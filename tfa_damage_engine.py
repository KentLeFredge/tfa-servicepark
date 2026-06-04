"""
tfa_damage_engine.py
====================
TFA — Moteur de calcul des dégâts

Architecture :
  - VehicleDB      : charge et expose les profils véhicules depuis JSON
  - DamageEngine   : calcule les scores depuis les impacts ACSM
  - RepairEngine   : calcule le coût et le résiduel d'une réparation partielle

Source de données :
  Fichier JSON multi-packs (tfa_vehicle_profiles.json)
  Impacts ACSM : ImpactSpeed (km/h), RelPosition {x, y, z}, Type (CAR|ENV)

Règles fondamentales :
  - Pas de trial-and-error : toutes les formules sont déterministes
  - La BDD véhicule est interchangeable (multi-packs, multi-fichiers)
  - Réparation partielle : le pilote choisit entre 0% et score_actuel%
"""

import json
import math
from dataclasses import dataclass, field
from typing import Dict, List, Optional


# ──────────────────────────────────────────────
# CONSTANTES (issues du fichier JSON, fallback ici)
# ──────────────────────────────────────────────

COMPONENTS     = ["refroidissement", "direction", "transmission", "suspension", "chassis"]
ZONES          = ["AVANT", "ARRIERE", "LATERAL_G", "LATERAL_D", "AERIEN"]

# Fallback si absent du JSON
DEFAULT_REPAIR_COST_MAX = {
    "refroidissement": 40,
    "direction":       35,
    "transmission":    30,
    "suspension":      15,   # reduced — less repair needed for robust 60s suspension
    "chassis":         50,
}

DEFAULT_PENALTY_MAX = {
    "refroidissement": {"type": "restrictor", "max": 12},
    "direction":       {"type": "ballast_kg",  "max": 20},
    "transmission":    {"type": "restrictor", "max":  8},
    "suspension":      {"type": "ballast_kg",  "max":  8},
    "chassis":         {"type": "ballast_kg",  "max": 20},
}

# ──────────────────────────────────────────────
# DATA CLASSES
# ──────────────────────────────────────────────

@dataclass
class Impact:
    """Un impact tel que reporté par ACSM."""
    speed_kmh: float
    rel_x: float          # positif = droite
    rel_y: float          # positif = haut (tonneau/décollage si > 0.5)
    rel_z: float          # positif = avant de la voiture heurtée
    impact_type: str      # "CAR" ou "ENV"


@dataclass
class ComponentState:
    """État d'un composant après calcul + réparation partielle."""
    id: str
    score: float          # 0.0 – 100.0 (dégât brut accumulé)
    repair_pct: float     # 0.0 – 100.0 (% réparé par le pilote, ≤ score)

    # Calculés depuis score et repair_pct
    score_residual: float = 0.0   # score - (score * repair_pct / 100)
    penalty_value: float  = 0.0   # ballast_kg ou restrictor résiduel
    penalty_type: str     = ""
    repair_cost_min: float = 0.0  # minutes consommées pour repair_pct

    def __post_init__(self):
        self._recalculate()

    def _recalculate(self):
        self.repair_pct = max(0.0, min(self.repair_pct, self.score))
        self.score_residual = self.score - self.repair_pct

    def apply_tables(self, repair_cost_max: Dict, penalty_max: Dict):
        """Calcule le coût de réparation et la pénalité résiduelle."""
        comp_id = self.id
        cost_max   = repair_cost_max.get(comp_id, DEFAULT_REPAIR_COST_MAX.get(comp_id, 30))
        pen_cfg    = penalty_max.get(comp_id, DEFAULT_PENALTY_MAX.get(comp_id, {}))
        pen_type   = pen_cfg.get("type", "ballast_kg")
        pen_max    = pen_cfg.get("max", 0)

        # Coût de réparation : proportionnel au % réparé × densité du dégât
        # La densité (score/100) reflète qu'une pièce très abîmée coûte plus à réparer
        if self.score > 0:
            self.repair_cost_min = (self.repair_pct / self.score) * cost_max * (self.score / 100)
        else:
            self.repair_cost_min = 0.0

        # Pénalité résiduelle : courbe progressive ^0.7
        # Concave légèrement : les premiers % de dégât pèsent peu,
        # les derniers % pèsent beaucoup
        if self.score_residual > 0:
            ratio = (self.score_residual / 100) ** 0.7
            self.penalty_value = round(ratio * pen_max, 1)
        else:
            self.penalty_value = 0.0

        self.penalty_type = pen_type
        self._recalculate()


# ──────────────────────────────────────────────
# VEHICLE DATABASE
# ──────────────────────────────────────────────

class VehicleDB:
    """
    Charge un ou plusieurs fichiers JSON de profils véhicules.
    Multi-packs : plusieurs packs peuvent coexister dans un même fichier,
    ou on peut charger plusieurs fichiers — le lookup est toujours par model_id.

    Format JSON attendu :
    {
      "_meta": {...},
      "repair_cost_max_min": {...},
      "penalty_max": {...},
      "damage_curve_exponent": 0.7,
      "impact_speed_min_kmh": 8,
      "env_multiplier": 1.0,
      "car_multiplier": 0.70,
      "speed_reference_kmh": 120,
      "speed_exponent": 1.8,
      "packs": {
        "pack_id": {
          "name": "...",
          "vehicles": {
            "model_id": {
              "zones": {
                "AVANT": {"refroidissement": 1.0, "direction": 1.0, ...},
                ...
              }
            }
          }
        }
      }
    }
    """

    def __init__(self):
        self._vehicles: Dict[str, Dict] = {}   # model_id → zones dict
        self._packs:    Dict[str, str]  = {}   # model_id → pack_id
        self.repair_cost_max: Dict = dict(DEFAULT_REPAIR_COST_MAX)
        self.penalty_max:     Dict = dict(DEFAULT_PENALTY_MAX)
        # Paramètres du moteur physique
        self.impact_speed_min     = 8.0
        self.env_multiplier       = 1.0
        self.car_multiplier       = 0.70
        self.speed_reference_car  = 100.0  # CAR ref — 100 km/h = impact catastrophique entre voitures
        self.speed_reference_env  = 150.0  # ENV ref — 150 km/h = sortie de route grave (Spa 66)
        self.speed_exponent       = 1.8

    def load(self, json_path: str) -> int:
        """
        Charge un fichier JSON.
        Peut être appelé plusieurs fois pour ajouter d'autres packs.
        Retourne le nombre de véhicules chargés depuis ce fichier.
        """
        with open(json_path, "r", encoding="utf-8") as f:
            data = json.load(f)

        # Paramètres globaux (premier fichier chargé gagne)
        if "repair_cost_max_min" in data:
            self.repair_cost_max = data["repair_cost_max_min"]
        if "penalty_max" in data:
            self.penalty_max = data["penalty_max"]
        if "impact_speed_min_kmh" in data:
            self.impact_speed_min = data["impact_speed_min_kmh"]
        if "env_multiplier" in data:
            self.env_multiplier = data["env_multiplier"]
        if "car_multiplier" in data:
            self.car_multiplier = data["car_multiplier"]
        if "speed_reference_car_kmh" in data:
            self.speed_reference_car = data["speed_reference_car_kmh"]
        if "speed_reference_env_kmh" in data:
            self.speed_reference_env = data["speed_reference_env_kmh"]
        # Legacy single reference
        if "speed_reference_kmh" in data and "speed_reference_car_kmh" not in data:
            self.speed_reference_car = data["speed_reference_kmh"]
            self.speed_reference_env = data["speed_reference_kmh"]
        if "speed_exponent" in data:
            self.speed_exponent = data["speed_exponent"]

        count = 0
        for pack_id, pack_data in data.get("packs", {}).items():
            for model_id, vehicle_data in pack_data.get("vehicles", {}).items():
                self._vehicles[model_id] = vehicle_data.get("zones", {})
                self._packs[model_id]    = pack_id
                count += 1

        return count

    def get_zones(self, model_id: str) -> Optional[Dict]:
        """Retourne le dict zones du véhicule, ou None si inconnu."""
        return self._vehicles.get(model_id)

    def get_pack(self, model_id: str) -> Optional[str]:
        return self._packs.get(model_id)

    def list_vehicles(self) -> List[str]:
        return list(self._vehicles.keys())

    def has_vehicle(self, model_id: str) -> bool:
        return model_id in self._vehicles


# ──────────────────────────────────────────────
# ZONE DETECTION
# ──────────────────────────────────────────────

def detect_zone(impact: Impact) -> str:
    """
    Détermine la zone d'impact depuis le vecteur RelPosition ACSM.

    Coordonnées ACSM (vérifiées sur logs réels) :
      rel_x : gauche/droite (positif = droite)
      rel_y : haut/bas      (rel_y > 0.5 ET speed > 30 = décollage/tonneau)
      rel_z : avant/arrière (positif = vers l'avant du véhicule impacté)

    Règle de priorité :
      1. AERIEN si rel_y > 0.5 et speed > 30 km/h
      2. Zone dominante entre Z (longitudinal) et X (latéral)
      3. En cas d'égalité, l'axe longitudinal prime
    """
    rx, ry, rz = impact.rel_x, impact.rel_y, impact.rel_z

    # Détection tonneau / décollage
    if abs(ry) > 0.5 and impact.speed_kmh > 30:
        return "AERIEN"

    az = abs(rz)
    ax = abs(rx)

    # Rapport 1.5 : une zone est dominante si elle est 50% plus forte
    if az >= ax:
        return "AVANT" if rz >= 0 else "ARRIERE"
    else:
        return "LATERAL_G" if rx < 0 else "LATERAL_D"


# ──────────────────────────────────────────────
# DAMAGE ENGINE
# ──────────────────────────────────────────────

class DamageEngine:
    """
    Calcule les scores de dégâts par composant à partir d'une liste d'impacts.

    Formule par impact :
      speed_eff   = speed × multiplier(type)
      energy      = min((speed_eff / speed_ref) ^ exponent × 100, 100)
      score[comp] += energy × zone_weight[comp]

    Score final plafonné à 100 par composant.
    """

    def __init__(self, db: VehicleDB):
        self.db = db

    def compute(self, model_id: str, impacts: List[Impact]) -> Dict[str, float]:
        """
        Retourne un dict {composant: score_float} pour la liste d'impacts.
        Si le modèle est inconnu, retourne des scores à 0.
        """
        scores = {c: 0.0 for c in COMPONENTS}
        zones  = self.db.get_zones(model_id)

        if zones is None:
            # Véhicule inconnu : profil générique prudent
            zones = _generic_fallback_zones()

        for impact in impacts:
            if impact.speed_kmh < self.db.impact_speed_min:
                continue

            # Filtre artefact ACSM : vitesses CAR absurdes (téléportation/spawn)
            # Observé sur données réelles : CAR 191 km/h avec |Y| < 0.3 = artefact
            if (impact.impact_type != "ENV"
                    and impact.speed_kmh > 150
                    and abs(impact.rel_y) < 0.3):
                continue

            # Vitesse effective et référence selon type de choc
            if impact.impact_type == "ENV":
                mult = self.db.env_multiplier
                ref  = self.db.speed_reference_env
            else:
                mult = self.db.car_multiplier
                ref  = self.db.speed_reference_car

            speed_eff = impact.speed_kmh * mult

            # Énergie normalisée 0–100 (courbe convexe)
            # Références calibrées sur données réelles (session Rusty Valley mai 2026) :
            #   CAR ref=80  : 80 km/h CAR → énergie 100 (impact catastrophique entre voitures)
            #   ENV ref=100 : 100 km/h ENV → énergie 100 (sortie de route grave)
            energy = min(
                (speed_eff / ref) ** self.db.speed_exponent * 100,
                100.0
            )

            zone = detect_zone(impact)
            zone_weights = zones.get(zone, {})

            for comp in COMPONENTS:
                w = zone_weights.get(comp, 0.0)
                scores[comp] = min(scores[comp] + energy * w, 100.0)

        return scores


def _generic_fallback_zones() -> Dict:
    """
    Profil générique utilisé quand un véhicule n'est pas dans la BDD.
    Exposition modérée uniforme sur toutes les zones.
    """
    return {
        zone: {comp: 0.35 for comp in COMPONENTS}
        for zone in ZONES
    }


# ──────────────────────────────────────────────
# REPAIR ENGINE
# ──────────────────────────────────────────────

class RepairEngine:
    """
    Gère la réparation partielle et calcule les pénalités résiduelles.

    Règle de réparation partielle :
      - repair_pct est compris entre 0 et score_actuel (pas au-delà)
      - Coût = proportionnel au % réparé et à la densité du dégât
      - Pénalité résiduelle = f(score_résiduel) en courbe ^0.7

    Budget total : 60 min
    Dépassement autorisé : 15 min max → pénalité TFA 1 min = 2 secondes
    """

    BUDGET_MIN       = 60
    OVERRUN_MAX_MIN  = 15
    OVERRUN_TO_SEC   = 2   # 1 min dépassement = 2 secondes sur chrono TFA

    def __init__(self, db: VehicleDB):
        self.db = db

    def build_components(
        self,
        scores: Dict[str, float],
        repair_choices: Dict[str, float]  # {comp: repair_pct}
    ) -> List[ComponentState]:
        """
        Construit la liste des ComponentState avec pénalités calculées.

        repair_choices : dict optionnel {composant: pct_à_réparer}
          Si absent ou 0 pour un composant → pas de réparation
        """
        states = []
        for comp in COMPONENTS:
            score      = round(scores.get(comp, 0.0), 2)
            repair_pct = repair_choices.get(comp, 0.0)
            # Clamp : on ne peut pas réparer plus que le dégât
            repair_pct = max(0.0, min(repair_pct, score))

            state = ComponentState(id=comp, score=score, repair_pct=repair_pct)
            state.apply_tables(self.db.repair_cost_max, self.db.penalty_max)
            states.append(state)
        return states

    def compute_totals(self, states: List[ComponentState]) -> Dict:
        """
        Calcule les totaux de ballast, restrictor, coût et pénalité TFA.
        """
        total_repair_min = sum(s.repair_cost_min for s in states)
        ballast_total    = sum(
            s.penalty_value for s in states
            if s.penalty_type == "ballast_kg"
        )
        # Restrictor : règle pondérée — max + 30% du second
        # Évite que le pilote ignore un composant pour ne payer que le max
        restr_values = sorted(
            [s.penalty_value for s in states if s.penalty_type == "restrictor"],
            reverse=True
        )
        if len(restr_values) == 0:
            restrictor_total = 0.0
        elif len(restr_values) == 1:
            restrictor_total = restr_values[0]
        else:
            restrictor_total = restr_values[0] + restr_values[1] * 0.30
        restrictor_max = min(restrictor_total, 12.0)

        overrun = max(0.0, total_repair_min - self.BUDGET_MIN)
        overrun = min(overrun, self.OVERRUN_MAX_MIN)
        penalty_seconds = overrun * self.OVERRUN_TO_SEC

        return {
            "repair_used_min":  round(total_repair_min, 1),
            "budget_min":       self.BUDGET_MIN,
            "overrun_min":      round(overrun, 1),
            "penalty_seconds":  penalty_seconds,
            "ballast_kg":       round(min(ballast_total, 40), 1),
            "restrictor":       round(restrictor_max, 1),
        }

    def repair_cost_preview(
        self,
        comp_id: str,
        score: float,
        repair_pct: float
    ) -> Dict:
        """
        Calcul préliminaire pour un seul composant — utilisé par l'UI slider
        pour afficher en temps réel le coût et la pénalité résiduelle.
        """
        state = ComponentState(id=comp_id, score=score, repair_pct=repair_pct)
        state.apply_tables(self.db.repair_cost_max, self.db.penalty_max)
        return {
            "repair_pct":     round(state.repair_pct, 1),
            "score_residual": round(state.score_residual, 1),
            "repair_cost_min":round(state.repair_cost_min, 1),
            "penalty_value":  round(state.penalty_value, 1),
            "penalty_type":   state.penalty_type,
        }


# ──────────────────────────────────────────────
# HELPERS : conversion depuis format ACSM JSON
# ──────────────────────────────────────────────

def impacts_from_acsm_result(acsm_result: Dict) -> Dict[str, List[Impact]]:
    """
    Convertit un résultat ACSM (dict JSON) en dict {driver_guid: [Impact, ...]}.

    Format ACSM attendu :
    {
      "Events": [
        {
          "Type": "CAR" | "ENV",
          "CarId": 0,
          "OtherCarId": 1,   (si CAR)
          "ImpactSpeed": 42.5,
          "RelPosition": {"X": 0.8, "Y": 0.0, "Z": 1.2}
        }
      ],
      "Cars": [
        {"CarId": 0, "Driver": {"Guid": "...", "Name": "..."}, "Model": "..."}
      ]
    }
    """
    # Map CarId → (guid, model)
    car_map = {}
    for car in acsm_result.get("Cars", []):
        cid   = car.get("CarId")
        guid  = car.get("Driver", {}).get("Guid", "")
        model = car.get("Model", "")
        car_map[cid] = (guid, model)

    impacts_by_guid: Dict[str, List[Impact]] = {}

    for event in acsm_result.get("Events", []):
        car_id = event.get("CarId")
        if car_id not in car_map:
            continue
        guid, _ = car_map[car_id]
        if not guid:
            continue

        speed   = float(event.get("ImpactSpeed", 0))
        rel     = event.get("RelPosition", {})
        rx      = float(rel.get("X", 0))
        ry      = float(rel.get("Y", 0))
        rz      = float(rel.get("Z", 0))
        itype   = event.get("Type", "ENV")

        imp = Impact(speed_kmh=speed, rel_x=rx, rel_y=ry, rel_z=rz,
                     impact_type=itype)

        if guid not in impacts_by_guid:
            impacts_by_guid[guid] = []
        impacts_by_guid[guid].append(imp)

    return impacts_by_guid


def model_from_acsm(acsm_result: Dict, driver_guid: str) -> Optional[str]:
    """Retourne le model_id d'un pilote depuis un résultat ACSM."""
    for car in acsm_result.get("Cars", []):
        if car.get("Driver", {}).get("Guid") == driver_guid:
            return car.get("Model")
    return None


# ──────────────────────────────────────────────
# SELF-TEST
# ──────────────────────────────────────────────

if __name__ == "__main__":
    import os

    DB_PATH = "/home/claude/tfa_vehicle_profiles.json"
    assert os.path.exists(DB_PATH), f"JSON not found: {DB_PATH}"

    db = VehicleDB()
    n = db.load(DB_PATH)
    print(f"Loaded {n} vehicles from {DB_PATH}")
    print(f"Packs: {set(db._packs.values())}")
    print(f"CAR ref={db.speed_reference_car} km/h  ENV ref={db.speed_reference_env} km/h")
    print()

    engine  = DamageEngine(db)
    repair  = RepairEngine(db)

    # ── Test 1 : Choc frontal 911 SWB à 27 km/h (tiré des logs réels)
    print("=== TEST 1 : 911 SWB — choc frontal CAR 27 km/h ===")
    impacts_911 = [
        Impact(speed_kmh=26.95, rel_x=0.80, rel_y=0.03, rel_z=1.22,
               impact_type="CAR")
    ]
    scores_911 = engine.compute("ac_legends_gtc_porsche_911_swb_8hlmc", impacts_911)
    print("Scores bruts:")
    for c, s in scores_911.items():
        print(f"  {c}: {s:.2f}")

    # Réparation partielle : pilote répare 50% de la direction
    choices = {c: 0.0 for c in COMPONENTS}
    choices["direction"] = scores_911["direction"] * 0.5  # 50% de réparation
    states = repair.build_components(scores_911, choices)
    totals = repair.compute_totals(states)
    print("\nAvec 50% réparation direction:")
    for s in states:
        if s.score > 0:
            print(f"  {s.id}: score={s.score:.1f} repair={s.repair_pct:.1f}% "
                  f"résiduel={s.score_residual:.1f} "
                  f"coût={s.repair_cost_min:.1f}min "
                  f"pénalité={s.penalty_value:.1f} {s.penalty_type}")
    print(f"  → Totaux: {totals}")

    print()

    # ── Test 2 : 250 LM — choc frontal à 80 km/h (radiateurs avant malgré moteur AR)
    print("=== TEST 2 : Ferrari 250 LM — choc frontal ENV 80 km/h ===")
    impacts_lm = [
        Impact(speed_kmh=80.0, rel_x=0.1, rel_y=0.0, rel_z=1.5,
               impact_type="ENV")
    ]
    scores_lm = engine.compute("wsc_legends_ferrari_250lm", impacts_lm)
    print("Scores bruts (refroidissement doit être élevé = radiateurs avant):")
    for c, s in scores_lm.items():
        if s > 0:
            print(f"  {c}: {s:.2f}")

    print()

    # ── Test 3 : Preview slider pour l'UI
    print("=== TEST 3 : Preview slider — direction score=45, repair_pct=30 ===")
    preview = repair.repair_cost_preview("direction", score=45.0, repair_pct=30.0)
    print(f"  {preview}")

    print()

    # ── Test 4 : Véhicule inconnu → fallback
    print("=== TEST 4 : Véhicule inconnu → fallback générique ===")
    impacts_unk = [Impact(speed_kmh=50.0, rel_x=0.0, rel_y=0.0, rel_z=1.0,
                          impact_type="CAR")]
    scores_unk = engine.compute("unknown_car_model", impacts_unk)
    print("Scores (fallback 0.35 uniforme):")
    for c, s in scores_unk.items():
        print(f"  {c}: {s:.2f}")

    print()
    print("All tests passed.")
