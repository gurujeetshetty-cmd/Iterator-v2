# Iterator Metrics Reference

Comprehensive guide to all 50+ metrics calculated by Iterator for evaluating segmentation quality.

---

## Metric Categories

1. [Probability Distribution Metrics](#probability-distribution-metrics)
2. [Segment Size Metrics](#segment-size-metrics)
3. [Bimodality Metrics](#bimodality-metrics)
4. [Differentiation Metrics](#differentiation-metrics)
5. [ROV (Rule of Variance) Metrics](#rov-metrics)
6. [Per-Segment Metrics](#per-segment-metrics)
7. [Custom OF_ Metrics](#custom-of_-metrics)
8. [Factor Analysis Metrics](#factor-analysis-metrics)
9. [Correlation Metrics](#correlation-metrics)
10. [Input Performance Metrics](#input-performance-metrics)

---

## Probability Distribution Metrics

These metrics measure how confidently respondents are assigned to segments.

### PROB_95
- **Formula:** `% of respondents with max segment probability > 0.95`
- **Interpretation:** High values (>80%) indicate strong, clear segment assignments
- **Threshold:** >0.80 is excellent, 0.60-0.80 is good, <0.60 needs improvement
- **Example:** `PROB_95 = 0.85` means 85% of respondents are very confidently assigned

### PROB_90
- **Formula:** `% of resp with max probability > 0.90`
- **Threshold:** >0.85 is good
- **Use:** Slightly more lenient than PROB_95

### PROB_80
- **Formula:** `% of resp with max probability > 0.80`
- **Threshold:** >0.90 is acceptable

### PROB_75
- **Formula:** `% of resp with max probability > 0.75`
- **Threshold:** >0.95 is minimum acceptable

### PROB_LESS_THAN_50
- **Formula:** `% of resp with max probability < 0.50`
- **Interpretation:** **Lower is better**. High values indicate respondents are split between segments
- **Threshold:** <0.05 is excellent, 0.05-0.10 is acceptable, >0.10 is poor
- **Example:** `PROB_LESS_THAN_50 = 0.03` means only 3% of respondents are ambiguous

---

## Segment Size Metrics

Measure the distribution of respondents across segments.

### MAX_N_SIZE
- **Formula:** `Largest segment size (count)`
- **Interpretation:** Helps identify dominant segments
- **Threshold:** Should not exceed 40% of total sample

### MAX_N_SIZE_PERC
- **Formula:** `Largest segment / total sample`
- **Interpretation:** Percentage of sample in largest segment
- **Threshold:** <0.40 is ideal, >0.50 indicates one segment dominates

### MIN_N_SIZE
- **Formula:** `Smallest segment size (count)`
- **Interpretation:** Ensures no segment is too small for analysis
- **Threshold:** >100 respondents minimum (depends on total sample)

### MIN_N_SIZE_PERC
- **Formula:** `Smallest segment / total sample`
- **Interpretation:** Percentage of sample in smallest segment
- **Threshold:** >0.10 is preferred, >0.05 is minimum

### SOLUTION_N_SIZE
- **Formula:** `Total number of segments`
- **Interpretation:** How many segments the solution has
- **Typical Values:** 3-5 segments for most applications

---

## Bimodality Metrics

Measure how well variables differentiate between top/bottom (polarization).

### BIMODAL_VARS
- **Formula:** Count of variables showing bimodal distribution (strong top AND bottom)
- **Definition:** Variable is bimodal if:
  - Top box index > 110 AND Bottom box index > 110
  - AND Top box % > 0.2 AND Bottom box % > 0.2
- **Interpretation:** Higher values mean variables have clear "love it or hate it" patterns
- **Threshold:** >50% of variables is ideal

### BIMODAL_VARS_PERC
- **Formula:** `BIMODAL_VARS / PROPER_BUCKETED_VARS`
- **Interpretation:** Percentage of variables showing bimodality
- **Threshold:** >0.50 is excellent, 0.30-0.50 is good, <0.30 is weak

### PROPER_BUCKETED_VARS
- **Formula:** Count of variables with proper top/middle/bottom buckets
- **Interpretation:** Total "scorable" variables in the analysis
- **Usage:** Denominator for calculating BIMODAL_VARS_PERC

---

## Differentiation Metrics

Evaluate how well variables distinguish between segments.

### Performance Variables
- **Definition:** Variable performs if one segment scores >120 index and another <80 index
- **Purpose:** Identifies variables that separate "high vs low" performers

### Independent Top Variables (indT)
- **Definition:** Segment scores >120 index with >20% base size
- **Purpose:** Segment is independently high on this variable

### Independent Bottom Variables (indB)
- **Definition:** Segment scores <80 index with >20% base size
- **Purpose:** Segment is independently low on this variable

---

## ROV Metrics

Rule of Variance metrics assess overall variable discrimination power.

### ROV_SD
- **Formula:** `Standard deviation of ROV values across variables`
- **Interpretation:** Variability in how much variables discriminate
- **Threshold:** Higher values indicate diverse discrimination levels

### ROV_RANGE
- **Formula:** `Max ROV - Min ROV`
- **Interpretation:** Spread between best and worst discriminating variables
- **Usage:** Helps identify if some variables are much stronger than others

**ROV Formula:**  
`ROV = (χ² - df) / √(2 * df)` where χ² is chi-square statistic

---

## Per-Segment Metrics

Calculated for each segment (1-5).

### bi_[1-5]
- **Formula:** Count of bimodal variables for segment N
- **Example:** `bi_1 = 12` means Segment 1 has 12 bimodal variables
- **Interpretation:** How many variables show love/hate for this segment
- **Threshold:** Higher is better (more defined segment)

### perf_[1-5]
- **Formula:** Count of performance-differentiating variables for segment N
-  **Interpretation:** How many variables distinguish this segment from others
- **Threshold:** Higher values = more differentiated segment

### indT_[1-5]
- **Formula:** Count of top-box independent variables for segment N
- **Interpretation:** Variables where this segment scores uniquely high
- **Usage:** Identifies segment's strengths

### indB_[1-5]
- **Formula:** Count of bottom-box independent variables for segment N
- **Interpretation:** Variables where this segment scores uniquely low
- **Usage:** Identifies segment's weaknesses

### seg[1-5]_diff
- **Formula:** Count of differentiating (non-bimodal but perf/ind) variables
- **Interpretation:** Variables that differentiate without being bimodal
- **Usage:** Supplementary differentiation beyond bimodality

---

## Custom OF_ Metrics

User-defined objective function metrics extracted from Excel output.

### OF_[NAME]_VALUES
- **Format:** `"value1 | value2 | value3 | value4"`
- **Example:** `OF_BRAND_PREFERENCE_VALUES = "85 | 72 | 91 | 78"`
- **Interpretation:** Mean values for each segment
- **Usage:** Display segment profiles on custom KPIs

### OF_[NAME]_MAXDIFF
- **Formula:** `max(all pairwise differences)`
- **Example:** `OF_BRAND_PREFERENCE_MAXDIFF = 19` (91 - 72)
- **Interpretation:** Maximum spread between any two segments
- **Threshold:** Higher values indicate better differentiation

### OF_[NAME]_MINDIFF
- **Formula:** `min(all pairwise differences)`
- **Example:** `OF_BRAND_PREFERENCE_MINDIFF = 3` (85 - 82 or similar)
- **Interpretation:** Minimum spread between any two segments
- **Threshold:** Higher values mean all segments are well-separated

**How to Create OF_ Metrics:**

In your Excel summary output, add a table:
```
Row 1: OF_METRIC_NAME
Row 2+: Numeric values (one column per segment)
```

Example:
```
OF_PURCHASE_INTENT
   Seg1    Seg2    Seg3    Seg4
   4.2     3.1     4.8     3.5
   4.5     3.3     4.6     3.7
```

---

## Factor Analysis Metrics

Requires `FA_OP` sheet in INPUT.xlsx.

### FA_DISTINCT_FACTORS
- **Formula:** Number of unique factors represented in combination
- **Interpretation:** Factor diversity in variable set
- **Threshold:** Higher values indicate broader construct coverage

### FA_FACTOR_COUNTS
- **Format:** `"F1:3, F2:5, F3:2"`
- **Interpretation:** How many variables load on each factor
- **Usage:** Ensure balanced factor representation

### FA_ENTROPY
- **Formula:** Shannon entropy of factor distribution: `-Σ(p * log(p))`
- **Interpretation:** Higher = more even distribution across factors
- **Threshold:** >1.0 for well-balanced sets

### FA_MEAN_PRIMARY_LOADING
- **Formula:** Average of highest factor loading per variable
- **Interpretation:** Average strength of primary factor relationship
- **Threshold:** >0.60 is good, >0.70 is excellent

### FA_MIN_PRIMARY_LOADING
- **Formula:** Minimum primary factor loading
- **Interpretation:** Weakest variable-factor relationship
- **Threshold:** >0.40 ensures all variables load sufficiently

### FA_MEAN_COMMUNALITY
- **Formula:** Average communality across variables
- **Interpretation:** How much variance is explained on average
- **Threshold:** >0.50 is acceptable, >0.60 is good

### FA_MIN_COMMUNALITY
- **Formula:** Minimum communality
- **Interpretation:** Worst-explained variable
- **Threshold:** >0.30 ensures all variables are adequately explained

### FA_CROSSLOAD_RATIO
- **Formula:** Average `(secondary loading / primary loading)`
- **Interpretation:** Amount of cross-loading (lower is better)
- **Threshold:** <0.50 is clean, >0.75 indicates problematic cross-loading

### FA_WEAK_ITEM_COUNT
- **Formula:** Count of variables with:
  - Primary loading < 0.4 OR
  - Cross-load ratio > 0.75 OR
  - Communality < 0.3
- **Interpretation:** Number of "weak" variables
- **Threshold:** 0 is ideal, <20% of variables is acceptable

### FA_COVERAGE_MIN_K
- **Formula:** Minimum number of variables per factor
- **Interpretation:** Ensures no factor is under-represented
- **Threshold:** ≥2 variables per factor minimum

---

## Correlation Metrics

Requires `RAW_DATA` sheet in INPUT.xlsx.

### P_R_MEAN_ABS
- **Formula:** Mean of absolute pairwise correlations
- **Interpretation:** Average relationship strength between variables
- **Threshold:** 0.3-0.5 is ideal (not too correlated, not independent)

### P_R_MEDIAN_ABS
- **Formula:** Median of absolute pairwise correlations
- **Interpretation:** Typical correlation between variables
- **Usage:** More robust to outlier pairs than mean

### P_R_MAX_ABS
- **Formula:** Maximum absolute pairwise correlation
- **Interpretation:** Strongest variable relationship
- **Threshold:** <0.80 (avoid multicollinearity), <0.90 is critical

### P_EIGVAR_PC1
- **Formula:** `λ₁ / Σλᵢ` (first eigenvalue / total variance)
- **Interpretation:** % variance explained by first principal component
- **Threshold:** <0.50 is diverse, >0.70 indicates one dominant dimension

### P_EFF_DIM
- **Formula:** `(Σλᵢ)² / Σ(λᵢ²)` (effective dimensionality)
- **Interpretation:** Equivalent number of independent dimensions
- **Threshold:** Higher is better (more independent information)

### P_TOPPAIR1
- **Format:** `"VAR1;VAR2"`
- **Interpretation:** Most correlated variable pair
- **Usage:** Identify potential redundancy

### P_TOPPAIR1_R
- **Formula:** Correlation coefficient of most correlated pair
- **Interpretation:** Strength of top correlation
- **Threshold:** <0.80 to avoid redundancy

---

## Input Performance Metrics

Evaluate how well input-flagged variables differentiate segments.

### BIMODAL_VARS_INPUT_VARS
- **Formula:** Count of bimodal variables among INPUT-flagged variables
- **Interpretation:** How many input variables show polarization
- **Usage:** Quality check on variable selection

### INPUT_PERF_SEG_COVERED
- **Formula:** Number of segments covered by performance variables
- **Interpretation:** How many segments have performance-differentiating input vars
- **Threshold:** Should equal number of segments (all covered)

### INPUT_PERF_FLAG
- **Formula:** `1` if all segments covered, `0` otherwise
- **Interpretation:** Binary flag for full segment coverage
- **Threshold:** Should be 1

### INPUT_PERF_MINSEG
- **Formula:** Minimum count of performance variables across segments
- **Interpretation:** Segment with fewest differentiators
- **Threshold:** >0 ensures all segments have some differentiation

### INPUT_PERF_MAXSEG
- **Formula:** Maximum count of performance variables across segments
- **Interpretation:** Segment with most differentiators
- **Usage:** Identify best-defined segment

---

## Metric Interpretation Guidelines

### High-Quality Solution Indicators

✅ **Excellent:**
- PROB_95 > 0.80
- BIMODAL_VARS_PERC > 0.50
- MIN_N_SIZE_PERC > 0.10
- PROB_LESS_THAN_50 < 0.05
- All segments have perf/ind variables

✅ **Good:**
- PROB_95 > 0.70
- BIMODAL_VARS_PERC > 0.35
- MIN_N_SIZE_PERC > 0.08
- PROB_LESS_THAN_50 < 0.10

⚠️ **Acceptable:**
- PROB_95 > 0.60
- BIMODAL_VARS_PERC > 0.25
- MIN_N_SIZE_PERC > 0.05

❌ **Poor:**
- PROB_95 < 0.60
- BIMODAL_VARS_PERC < 0.25
- PROB_LESS_THAN_50 > 0.10

### Typical Filtering Strategy

When exploring results in the Results Explorer tab:

1. **Basic Quality:** `PROB_95 > 0.75 AND MIN_N_SIZE_PERC > 0.08`
2. **Strong Differentiation:** `BIMODAL_VARS_PERC > 0.40 AND PROPER_BUCKETED_VARS > 10`
3. **Balanced Segments:** `MAX_N_SIZE_PERC < 0.40 AND MIN_N_SIZE_PERC > 0.12`
4. **Custom KPI:** `OF_BRAND_PREFERENCE_MAXDIFF > 15`

---

## Metric Correlation Insights

**Highly Correlated Metrics** (r > 0.80):
- PROB_95 ↔ PROB_90
- BIMODAL_VARS ↔ BIMODAL_VARS_PERC
- MAX_N_SIZE ↔ MAX_N_SIZE_PERC

**Independent Metrics** (r < 0.30):
- PROB_95 vs BIMODAL_VARS_PERC
- ROV metrics vs probability metrics

**Key Tradeoffs:**
- More variables → Higher BIMODAL_VARS but potentially lower PROB_95
- More segments → More differentiation but lower MIN_N_SIZE

---

## Advanced Usage

### Creating Custom Composite Scores

You can create weighted composite scores in Excel after export:

```excel
=0.4*PROB_95 + 0.3*BIMODAL_VARS_PERC + 0.2*MIN_N_SIZE_PERC + 0.1*INPUT_PERF_FLAG
```

### Filtering for Business Use Cases

**Brand Positioning Study:**
```
PROB_95 > 0.75
BIMODAL_VARS_PERC > 0.40
OF_BRAND_PREFERENCE_MAXDIFF > 20
```

**Customer Segmentation:**
```
MIN_N_SIZE > 100
MAX_N_SIZE_PERC < 0.35
PROB_LESS_THAN_50 < 0.05
```

**Needs-Based Segmentation:**
```
BIMODAL_VARS_PERC > 0.50
FA_DISTINCT_FACTORS > 3
P_EFF_DIM > 4
```

---

## Glossary

- **Bimodal:** Distribution with strong representation at both top and bottom
- **Index:** Score relative to average (100 = average, 120 = 20% above avg)
- **Top Box:** Highest 1-2 response categories in a rating scale
- **Bottom Box:** Lowest 1-2 response categories
- **ROV:** Rule of Variance statistic measuring discrimination strength
- **Communality:** Proportion of variance explained by factors
- **Cross-loading:** Variable loading significantly on multiple factors

---

**Last Updated:** 2025-11-28
