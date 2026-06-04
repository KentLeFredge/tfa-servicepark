# TFA SERVICE PARK — Référence Moteur de Dégâts

> Document de calibration — valeurs basées sur les défauts actuels du code

---

## 1. Calcul du Score (0–100)

### Formule
```
speed_eff = vitesse_impact × multiplicateur_type
energy    = min( (speed_eff / speed_ref)^1.8 × 100 , 100 )
score[composant] += energy × poids_zone[composant]
score plafonné à 100
```

### Paramètres physiques (feuille config)

| Paramètre | Valeur | Rôle |
|---|---|---|
| `SPEED_MIN_KMH` | 8 km/h | Choc ignoré en dessous |
| `SPEED_REF_CAR_KMH` | 100 km/h | 100 km/h voiture-voiture = score 100 |
| `SPEED_REF_ENV_KMH` | 150 km/h | 150 km/h mur = score 100 |
| `SPEED_EXPONENT` | 1.8 | Courbe convexe (gros chocs très pénalisés) |
| `CAR_MULTIPLIER` | 0.70 | Choc voiture-voiture atténué × 0.70 |
| `ENV_MULTIPLIER` | 1.00 | Choc mur = référence |

### Exemples de scores générés (composant primaire, poids = 1.0)

| Impact | Type | Speed eff. | Score brut |
|---|---|---|---|
| 15 km/h | CAR | 10.5 | 1.4 → **ignoré** |
| 25 km/h | CAR | 17.5 | 4.7 → **ignoré** |
| 30 km/h | CAR | 21.0 | 6.6 → **léger** |
| 50 km/h | CAR | 35.0 | 16.2 → **léger** |
| 80 km/h | CAR | 56.0 | 37.4 → **modéré** |
| 100 km/h | CAR | 70.0 | 55.3 → **sévère** |
| 30 km/h | ENV | 30.0 | 9.3 → **léger** |
| 60 km/h | ENV | 60.0 | 19.3 → **léger** |
| 100 km/h | ENV | 100.0 | 46.4 → **sévère** |
| 130 km/h | ENV | 130.0 | 72.3 → **critique** |

---

## 2. Score → Sévérité

| Sévérité | Score |
|---|---|
| Aucun dégât | 0 – 5 |
| **Léger** | 5 – 20 |
| **Modéré** | 20 – 45 |
| **Sévère** | 45 – 70 |
| **Critique** | 70 – 100 |

---

## 3. Pénalités par Composant

> Formule : `pénalité = (score_résiduel / 100)^0.7 × max`
> La courbe est **progressive** : 50% de dégât ≠ 50% de la pénalité max.

### Refroidissement → RESTRICTOR (max 12)

| Score | Restrictor |
|---|---|
| 5 | 0.9 |
| 20 | 2.9 |
| 45 | 5.5 |
| 70 | 7.8 |
| 100 | 12.0 |

### Direction → BALLAST kg (max 20 kg)

| Score | Ballast |
|---|---|
| 5 | 1.5 kg |
| 20 | 4.8 kg |
| 45 | 9.2 kg |
| 70 | 13.0 kg |
| 100 | 20.0 kg |

### Transmission → RESTRICTOR (max 8)

| Score | Restrictor |
|---|---|
| 5 | 0.6 |
| 20 | 1.9 |
| 45 | 3.7 |
| 70 | 5.2 |
| 100 | 8.0 |

### Suspension → BALLAST kg (max 8 kg)

| Score | Ballast |
|---|---|
| 5 | 0.6 kg |
| 20 | 1.9 kg |
| 45 | 3.7 kg |
| 70 | 5.2 kg |
| 100 | 8.0 kg |

### Châssis → BALLAST kg (max 20 kg)

| Score | Ballast |
|---|---|
| 5 | 1.5 kg |
| 20 | 4.8 kg |
| 45 | 9.2 kg |
| 70 | 13.0 kg |
| 100 | 20.0 kg |

---

## 4. Règle Restrictor (max 2 composants)

```
restrictor_total = max(R1, R2, ...) + 2e_plus_grand × 0.30
plafonné à 12
```

Exemple : refroidissement (R=8) + transmission (R=5)
→ 8 + 5 × 0.30 = **9.5**

---

## 5. Coûts de Réparation (budget : 60 min)

> Formule : `coût = (repair_pct / score) × cost_max × (score / 100)`

| Composant | Coût max (réparation complète à score 100) |
|---|---|
| Refroidissement | 40 min |
| Direction | 35 min |
| Transmission | 30 min |
| Suspension | 20 min |
| Châssis | 50 min |

**Budget total pilote : 60 min**
**Dépassement autorisé : 15 min max → 2 sec de pénalité chrono par min**

---

## 6. Accumulation sur un Pilote (exemple réel)

Pilote avec 3 incidents :
- Choc mur à 80 km/h → zone AVANT → refroidissement score **34** (restrictor résiduel = **5.8**)
- Choc voiture à 60 km/h → zone LATERAL → suspension score **18** (ballast résiduel = **1.7 kg**)
- Choc mur à 45 km/h → zone ARRIERE → transmission score **12** (restrictor résiduel = **1.7**)

Restrictor total = 5.8 + 1.7 × 0.30 = **6.3**
Ballast total = **1.7 kg**

Sans réparation → **6.3 restrictor + 1.7 kg ballast** pour l'étape suivante.

---

## 7. Paramètres à Ajuster (feuille config Google Sheet)

Pour rendre le moteur **plus sensible** (plus de dégâts) :
- Baisser `SPEED_REF_CAR_KMH` (ex: 70 au lieu de 100)
- Baisser `SPEED_REF_ENV_KMH` (ex: 100 au lieu de 150)
- Baisser `SCORE_RIEN_MAX` (ex: 3 au lieu de 5)

Pour rendre le moteur **moins sensible** :
- Augmenter les références de vitesse
- Augmenter `SCORE_RIEN_MAX`

Pour augmenter les **pénalités** :
- Augmenter `PENALTY_DIRECTION_MAX`, `PENALTY_CHASSIS_MAX`, etc.

---

*Généré automatiquement depuis TFA Service Park — code.gs CONFIG_DEFAULTS*
