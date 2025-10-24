# 📈 Automatic Linear Model Fitting and Evaluation in R

## Overview
`linear_model_function()` is an R utility designed to **automatically fit, diagnose, and evaluate linear models (`lm`)**.  
It generates diagnostic plots, forest plots, effect plots, and contrast analyses, saving all outputs in a structured folder system.

---

## ✳️ Features
- Automatically installs and loads required packages.
- Sequentially creates folders (`fit_01`, `fit_02`, …) to organize model runs.
- Fits a **linear model** using `lm()`.
- Generates and saves:
  - Diagnostic plots (residuals, QQ, leverage, etc.)
  - Forest plots of estimates.
  - Pairwise contrast plots (for factors).
  - Effect plots for each independent variable.
- Exports model summaries in:
  - `.html` (model summary)
  - `.docx` and `.png` (contrast tables)

---

## 🧩 Dependencies
The function uses the following packages:

```
sjPlot, parameters, performance, dplyr, tidyr, lme4, lmerTest,
gtsummary, emmeans, flextable, ggeffects, ggplot2
```

and relies on the following custom helper scripts (which must be available in your environment):

- `print_centered_note_v1.R`
- `Automate_Saving_ggplots.R`
- `automate_saving_dataframes_xlsx_format.R`
- `create_sequential_dir.R`

---

## ⚙️ Usage

```r
source("path/to/linear_model_function.R")

# Example using built-in dataset
fit <- linear_model_function(Sepal.Length ~ Species, iris)
```

All output files (plots, summaries, reports) will be stored in a new directory:
```
fit_01/
├── Diagnostic.png
├── ForestPlot.png
├── Contrast_Estimates_Plot_log.png
├── EffectPlot_Species.png
├── Contrast_LMM_Summary.docx
├── Contrast_LMM_Summary.png
└── fit_01_LMM_Summary.html
```

---

## 📤 Output
- **Model object**: returned invisibly (can be reused for further analysis).
- **Saved results**: Plots and reports in a dedicated subfolder.

---

## 🪄 License
MIT License © 2025 Raúl
