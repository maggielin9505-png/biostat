import pandas as pd
import numpy as np
from scipy import stats

df = pd.read_excel('調查情報20260331.xlsx')

print('=== 1. CROSS TAB: drinking water vs disease status ===')
ct_water_dis = pd.crosstab(df['drinking water'], df['disease status'], margins=True)
print(ct_water_dis)
print('\nPercentages within drinking water:')
print(pd.crosstab(df['drinking water'], df['disease status'], normalize='index') * 100)

print('\n=== 2. CROSS TAB: water_group vs disease status ===')
ct_wg_dis = pd.crosstab(df['water_group'], df['disease status'], margins=True)
print(ct_wg_dis)
print('\nPercentages within water_group:')
print(pd.crosstab(df['water_group'], df['disease status'], normalize='index') * 100)

# Binary disease: diseased (minor + severe) vs healthy
df['is_diseased'] = df['disease status'].apply(lambda x: 1 if x in ['minor', 'severe'] else 0)
df['is_dead'] = df['survival outcome'].apply(lambda x: 1 if x == 'Dead' else 0)

print('\n=== Check if is_diseased == is_dead ===')
print((df['is_diseased'] == df['is_dead']).all())

# Since is_diseased == is_dead exactly, any test on disease status (Healthy vs Diseased)
# is mathematically IDENTICAL to survival outcome (alive vs Dead)!

print('\n=== 3. 2x2 TABLES & EPIDEMIOLOGICAL METRICS ===')

def calc_2x2(table_2x2, name=''):
    # table format: [[exposed_case, exposed_noncase], [unexposed_case, unexposed_noncase]]
    a, b = table_2x2[0][0], table_2x2[0][1]
    c, d = table_2x2[1][0], table_2x2[1][1]
    n1 = a + b
    n0 = c + d
    m1 = a + c
    m0 = b + d
    N = n1 + n0
    
    # Chi-square
    chi2, p, dof, ex = stats.chi2_contingency([[a, b], [c, d]], correction=False)
    chi2_corr, p_corr, _, _ = stats.chi2_contingency([[a, b], [c, d]], correction=True)
    odds_ratio, p_fisher = stats.fisher_exact([[a, b], [c, d]])
    
    # RR = (a / n1) / (c / n0)
    # If c == 0 or b == 0, handle
    r1 = a / n1 if n1 > 0 else 0
    r0 = c / n0 if n0 > 0 else 0
    rr = r1 / r0 if r0 > 0 else np.nan
    
    # RR 95% CI (Woolf / Katz method)
    # SE(ln RR) = sqrt( (1/a - 1/n1) + (1/c - 1/n0) )
    # If a, c > 0
    if a > 0 and c > 0 and n1 > a and n0 > c:
        se_ln_rr = np.sqrt((1/a - 1/n1) + (1/c - 1/n0))
        rr_ci_low = np.exp(np.log(rr) - 1.96 * se_ln_rr)
        rr_ci_up = np.exp(np.log(rr) + 1.96 * se_ln_rr)
    else:
        # Haldane correction if 0
        se_ln_rr = np.sqrt((1/(a+0.5) - 1/(n1+1)) + (1/(c+0.5) - 1/(n0+1)))
        rr_ci_low = np.exp(np.log(rr) - 1.96 * se_ln_rr)
        rr_ci_up = np.exp(np.log(rr) + 1.96 * se_ln_rr)
        
    # OR = (a * d) / (b * c)
    or_val = (a * d) / (b * c) if (b * c) > 0 else np.nan
    # SE(ln OR) = sqrt(1/a + 1/b + 1/c + 1/d)
    if min(a, b, c, d) > 0:
        se_ln_or = np.sqrt(1/a + 1/b + 1/c + 1/d)
        or_ci_low = np.exp(np.log(or_val) - 1.96 * se_ln_or)
        or_ci_up = np.exp(np.log(or_val) + 1.96 * se_ln_or)
    else:
        se_ln_or = np.sqrt(1/(a+0.5) + 1/(b+0.5) + 1/(c+0.5) + 1/(d+0.5))
        or_ci_low = np.exp(np.log((a*d)/(b*c) if b*c>0 else ((a+0.5)*(d+0.5))/((b+0.5)*(c+0.5))) - 1.96 * se_ln_or)
        or_ci_up = np.exp(np.log((a*d)/(b*c) if b*c>0 else ((a+0.5)*(d+0.5))/((b+0.5)*(c+0.5))) + 1.96 * se_ln_or)
        
    print(f'=== {name} ===')
    print(f'2x2 table: [[{a}, {b}], [{c}, {d}]]')
    print(f'Exposed Risk: {r1:.4f} ({a}/{n1}), Unexposed Risk: {r0:.4f} ({c}/{n0})')
    print(f'Pearson Chi-square: {chi2:.3f}, df={dof}, p={p:.4e}')
    print(f'Continuity Corrected Chi-square: {chi2_corr:.3f}, p={p_corr:.4e}')
    print(f'Fisher Exact Test p: {p_fisher:.4e}')
    print(f'Relative Risk (RR): {rr:.4f} (95% CI: {rr_ci_low:.4f} - {rr_ci_up:.4f})')
    print(f'Odds Ratio (OR): {or_val:.4f} (95% CI: {or_ci_low:.4f} - {or_ci_up:.4f})')
    print()

# 1. water_group: Group 2 (Deep well) vs Group 1 (Top water + shallow well)
# Outcome: Dead/Diseased (Case) vs Alive/Healthy (Noncase)
# Note SPSS coding in table:
# water_group 2: Dead=501, Alive=1
# water_group 1: Dead=14, Alive=153
# If Exposed = Group 2, Unexposed = Group 1:
# Exposed: Case=501, Noncase=1
# Unexposed: Case=14, Noncase=153
calc_2x2([[501, 1], [14, 153]], 'Exposure: Group 2 (Deep well) vs Group 1 (Non-deep well) -> Outcome: Dead/Diseased')

# Notice what SPSS printed for water_group Odds ratio (1 / 2):
# In SPSS: Row 1 = Group 1, Row 2 = Group 2. Outcome col 1 = alive, col 2 = Dead.
# SPSS: (153 * 501) / (14 * 1) = 76653 / 14 = 5475.214!
# If Group 2 vs Group 1 for Dead vs Alive:
# (501 * 153) / (1 * 14) = 5475.214!
# Risk of death in Group 2 = 501 / 502 = 99.8008%
# Risk of death in Group 1 = 14 / 167 = 8.3832%
# RR = 99.8008% / 8.3832% = 11.9048! (Group 2 has 11.90 times the risk of death/illness compared to Group 1)
# And if looking at SPSS cohort alive:
# Group 1 alive = 153/167 = 91.6168%
# Group 2 alive = 1/502 = 0.1992%
# RR(alive) = 91.6168% / 0.1992% = 459.916 (which matches SPSS line 1822: 459.916!)
# RR(Dead) = 8.3832% / 99.8008% = 0.084 (matches SPSS line 1826: .084!)

# 2. Deep well vs shallow well
# Deep well: Dead=501, Alive=1
# shallow well: Dead=14, Alive=101
calc_2x2([[501, 1], [14, 101]], 'Exposure: Deep well vs Shallow well -> Outcome: Dead/Diseased')

# 3. Deep well vs Top water
# Deep well: Dead=501, Alive=1
# Top water: Dead=0, Alive=52
calc_2x2([[501, 1], [0, 52]], 'Exposure: Deep well vs Top water -> Outcome: Dead/Diseased')

# 4. Shallow well vs Top water
# Shallow well: Dead=14, Alive=101
# Top water: Dead=0, Alive=52
calc_2x2([[14, 101], [0, 52]], 'Exposure: Shallow well vs Top water -> Outcome: Dead/Diseased')

# 5. Gender: Male vs Female
# M: Dead=200, Alive=60
# F: Dead=315, Alive=94
calc_2x2([[200, 60], [315, 94]], 'Exposure: Male vs Female -> Outcome: Dead/Diseased')

# 6. Age: >40 vs <=40
# >40: Dead=247, Alive=73
# <=40: Dead=268, Alive=81
calc_2x2([[247, 73], [268, 81]], 'Exposure: Age >40 vs Age <=40 -> Outcome: Dead/Diseased')
