# KFC Food Cost Hub — Pipeline

Manual monthly pipeline για το Food Cost Hub dashboard.

## Αρχεία σε αυτό το φάκελο

| File | Σκοπός |
|------|--------|
| `build_food_data.py` | Core parser — διαβάζει τα xlsx, παράγει `food_data.json` |
| `build_pipeline.py` | Wrapper — κάνει build + git push |
| `build_food_data.bat` | One-click wrapper για manual execution |

## Πώς το χρησιμοποιείς (μηνιαία)

### Setup (μία φορά)

Αντιγραφή των 3 αρχείων στο:

```
C:\Users\IT\Documents\GitHub\kfc-sales-dashboard\automation\foodcost\
  ├── build_food_data.py
  ├── build_pipeline.py
  └── build_food_data.bat
```

### Μηνιαίο workflow

**Στις ~15 του μήνα**, όταν παίρνεις τα food cost αρχεία από το λογιστήριο:

#### Βήμα 1: Αντιγραφή αρχείων
Αντιγράφεις στο `_work\` φάκελο:

```
C:\Users\IT\Documents\GitHub\kfc-sales-dashboard\_work\
  ├── FoodCost.xlsx          ← overwrite κάθε φορά (το κεντρικό αρχείο)
  └── CategoriesFC.xlsx      ← overwrite μόνο όταν αλλάξει
```

#### Βήμα 2: Run

Διπλό click στο `build_food_data.bat`

ή από PowerShell:
```powershell
cd C:\Users\IT\Documents\GitHub\kfc-sales-dashboard\automation\foodcost
.\build_food_data.bat
```

#### Βήμα 3: Έλεγχος

Αν όλα πάνε καλά, θα δεις:
```
============================================
✓ DONE — HH:MM:SS
============================================
```

Το dashboard ενημερώνεται μέσα σε 1-2 λεπτά (GitHub Pages deployment).

## Τι κάνει το script

1. **Verify input files** — ελέγχει ότι υπάρχουν τα 2 required files στο `_work\`
2. **Parse FoodCost.xlsx** — διαβάζει 268K rows ιστορικού
3. **Apply categories** — από το CategoriesFC.xlsx
4. **Filter to 22 KFC stores** — exclude warehouses
5. **Aggregate monthly** — per store × per month × per category × per reason
6. **Compute variance** — Theoretical vs Actual για κάθε ingredient
7. **Top variances** — top 20 items per store για τον latest month
8. **Cross-validate** — με το 2026_ΚΟΣΤΟΛΟΓΗΣΗ.xlsx (αν υπάρχει)
9. **Write food_data.json** στο `kfc-dashboard\food\`
10. **Git push** με 3 retries

## Output

Δημιουργείται το αρχείο:
```
kfc-dashboard\food\food_data.json   (~1.5 MB)
```

Schema:
```
meta:
  - 51+ periods (2022-01 onwards)
  - 22 stores
  - 16 categories

monthly: array of {year, month, store} records
  pure_food: { ideal_cost, actual_cost, variance, qty }
  total: { ίδια metrics συμπεριλαμβανομένων packaging/services }
  waste: { cost, qty }
  by_reason: { NORMAL, ΦΥΡΑ, ΦΥΣ.ΑΠΟΓΡ, ΠΑΡΑΓ_KFC, ... }
  by_category: { CHICKEN, SAUCES, ... }

top_variances_latest:
  Per store, top 20 ingredients by variance €
```

## Troubleshooting

### "MISSING REQUIRED FILES"
Δεν έχεις βάλει τα αρχεία στο `_work\`. Έλεγξε τα paths.

### "git push failed"
- Πιθανώς λήξε ο PAT (το `.github_pat`)
- Ή internet drop
- Δοκίμασε manual: `cd C:\...\kfc-sales-dashboard && git push`

### Λάθος νούμερα στο dashboard
- Έλεγξε αν το `FoodCost.xlsx` είναι το σωστό (όχι παλιά version)
- Έλεγξε ότι το CategoriesFC.xlsx έχει όλα τα ingredients mapped
- Run ξανά με `--no-push` και δες τα logs στο terminal

### Build only (χωρίς push)
```powershell
python build_pipeline.py --no-push
```

## Σημαντικά

- **Manual execution** — δεν τρέχει αυτόματα μέσω Task Scheduler
- **Monthly cadence** — τρέχεις όταν παίρνεις τα νέα αρχεία (~15 του μήνα)
- **Idempotent** — ασφαλές να τρέχει πολλές φορές, παράγει το ίδιο αποτέλεσμα
- **Backwards compatible** — διαβάζει και ιστορικά δεδομένα, οπότε προσθήκη νέου μήνα αυτόματα ενημερώνει όλο το ιστορικό
