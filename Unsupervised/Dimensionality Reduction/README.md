# PCA_auto - Automated Principal Component Analysis in R  

![R](https://img.shields.io/badge/R-%E2%89%A54.0.0-blue.svg)  ![License](https://img.shields.io/badge/License-MIT-green.svg)  ![Version](https://img.shields.io/badge/Version-1.0.0-orange.svg)  

---

## 📊 Overview  
**PCA_auto** is a comprehensive R function that performs automated Principal Component Analysis (PCA) with built-in statistical assumption testing, professional reporting, and publication-ready visualizations.  
This function streamlines the entire PCA workflow into a single command, making complex multivariate analysis accessible and reproducible.  

---

## ✨ Features  
- 🔄 **Automated Workflow**: Complete PCA analysis in one function call  
- 📈 **Statistical Validation**: Comprehensive assumption testing (normality, multicolinearity)  
- 📊 **Professional Reporting**: Excel-based reports with detailed results  
- 🎨 **Publication-Ready Plots**: High-quality `ggplot2` visualizations  
- 📁 **Organized Output**: Automatic directory structure management  
- 🔍 **Quality Control**: Built-in data validation and error checking  

---

## 📥 Installation  

### Prerequisites  
Ensure you have **R (≥ 4.0.0)** installed along with these required packages:  

```r
# Install required packages
install.packages(c(
  'openxlsx', 'ggplot2', 'FactoMineR', 'factoextra', 'psych', 'MVN', 'dplyr',
  'corrr', 'corrplot', 'tidyr', 'ggpubr', 'nortest', 'patchwork',
  'ggcorrplot', 'ggplotify'
))
```
## 📊 Output Structure
```bash
PCA_My_Study_2024/
├── 📊 Combined_PCA_Figures_Plot.png
├── 📁 Figures/
│   ├── 🎯 ScreePlot.png
│   ├── 📈 PCA_Variable_Contribution_PC1_PC2.png
│   ├── 👥 Individual_Contribution.png
│   ├── 🎨 PCA_Ind_Colored_by_group.png
│   └── 🔗 Corrplot_Colinearity.png
└── 📁 Reports/
    ├── 📋 Assumptions_Report.xlsx
    │   ├── Univariate_Normality
    │   ├── Multivariate_Normality
    │   ├── Correlation_Matrix
    │   └── Colineality_Test
    └── 📈 PCA Results My_Study_2024.xlsx
        ├── Eigenvalues
        └── Dimension descriptions
```
## 🔍 Statistical Tests Performed

When `perform_assumtions = TRUE`:

### ✅ Data Validation  
- Numeric variable check  
- Missing value detection  

### 📊 Univariate Normality  
- Shapiro-Wilk test (n ≥ 50)  
- Lilliefors test (n < 50)  
- QQ-plots and distribution plots  

### 📈 Multivariate Normality  
- Mardia's test  

### 🔗 Multicolinearity Analysis  
- Correlation matrix  
- Bartlett's sphericity test  
- Kaiser-Meyer-Olkin (KMO) index  
- Determinant analysis  

---

## 📈 PCA Output  

The function returns a comprehensive PCA object containing:  

- **Eigenvalues**: Variance explained by each component  
- **Variable coordinates**: Loadings and contributions  
- **Individual coordinates**: Sample positions in PCA space  
- **Quality measures**: cos2, contributions, and representation quality  

```r
# Access results
summary(pca_results)
pca_results$eig          # Eigenvalues and variance
pca_results$var$coord    # Variable coordinates
pca_results$ind$coord    # Individual coordinates
```
## 📝 Examples
Example 1: Basic Analysis

```r
# Using built-in dataset
data(iris)
pca_iris <- PCA_auto(
  dataframe = iris,
  color_variables = "Species",
  title = "PCA_Iris_Analysis"
)
```
Example 2: Custom Analysis
```r
# Custom dataset with multiple grouping variables
pca_custom <- PCA_auto(
  dataframe = experimental_data,
  where_to_save = "analyses/2024/pca/",
  title = "PCA_Experiment_1",
  color_variables = c("condition", "batch", "replicate"),
  perform_assumtions = TRUE
)
```
