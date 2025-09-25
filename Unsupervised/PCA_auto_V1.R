########## 24/09/2025 ##########

# ================================================ #
# ==== AUTOMATIC PRINCIPAL COMPONENT ANALYSIS ==== #
# ================================================ #

#' Automated Principal Component Analysis (PCA) with Comprehensive Reporting
#' 
#' Performs a complete and automated Principal Component Analysis including 
#' assumption checking, statistical testing, and generation of comprehensive 
#' reports and visualizations.
#'
#' @param dataframe A data.frame containing the dataset for PCA analysis. 
#'   Must contain only numeric variables except those specified in color_variables.
#' @param where_to_save Character string specifying the directory path where 
#'   results will be saved. If NULL, uses current working directory.
#' @param title Character string for the analysis title and folder name. 
#'   Defaults to "PCA_DD_MM_YYYY".
#' @param color_variables Character vector specifying categorical variable names 
#'   to be used for coloring individuals in PCA plots.
#' @param scaled_data Logical indicating whether the data is already scaled. 
#'   If FALSE, PCA will scale the data internally. Defaults to FALSE.
#' @param perform_assumptions Logical indicating whether to perform statistical 
#'   assumption tests (normality, multicolinearity, etc.). Defaults to TRUE.
#'
#' @return Returns a PCA object from FactoMineR containing:
#'   \itemize{
#'     \item eigenvalues: Eigenvalues and percentage of variance explained
#'     \item ind: Individual coordinates, contributions, and cos2
#'     \item var: Variable coordinates, contributions, and cos2
#'     \item etc.: Other PCA results from FactoMineR
#'   }
#'   Additionally, generates comprehensive Excel reports and ggplot visualizations 
#'   in an organized directory structure.
#'
#' @details
#' This function performs a complete PCA workflow including:
#' \enumerate{
#'   \item Data validation and assumption testing
#'   \item Principal Component Analysis execution
#'   \item Comprehensive reporting in Excel format
#'   \item Professional visualizations for publication
#'   \item Organized output directory structure
#' }
#' 
#' The function automatically checks for numeric variables, missing values, 
#' and performs statistical assumption tests when requested.
#'
#' @section Directory Structure:
#' The function creates the following directory structure:
#' \preformatted{
#' output_dir/
#' ├── Figures/
#' │   ├── ScreePlot.png
#' │   ├── PCA_Variable_Contribution*.png
#' │   └── Individual_Contribution.png
#' ├── Reports/
#' │   ├── Assumptions_Report.xlsx
#' │   └── PCA Results [title].xlsx
#' └── Combined_PCA_Figures_Plot.png
#' }
#'
#' @section Statistical Assumptions:
#' When \code{perform_assumtions = TRUE}, the function tests:
#' \itemize{
#'   \item Univariate normality (Shapiro-Wilk/Lilliefors test)
#'   \item Multivariate normality (Mardia's test)
#'   \item Multicolinearity (Correlation matrix, Bartlett's test, KMO index)
#'   \item Variable suitability for PCA
#' }
#'
#' @examples
#' \dontrun{
#' # Basic usage
#' pca_result <- PCA_auto(
#'   dataframe = my_data,
#'   where_to_save = "results/pca_analysis/",
#'   title = "PCA_My_Study",
#'   color_variables = c("group", "treatment")
#' )
#'
#' # Quick analysis without assumption testing
#' pca_quick <- PCA_auto(
#'   dataframe = numeric_data,
#'   color_variables = "condition",
#'   perform_assumtions = FALSE
#' )
#'
#' # Access PCA results
#' summary(pca_result)
#' pca_result$eig # Eigenvalues
#' pca_result$ind$coord # Individual coordinates
#' }
#'
#' @seealso
#' \code{\link[FactoMineR]{PCA}}, \code{\link[factoextra]{fviz_pca_ind}},
#' \code{\link[psych]{KMO}}, \code{\link[MVN]{mvn}}
#'
#' @importFrom FactoMineR PCA dimdesc
#' @importFrom factoextra fviz_eig fviz_pca_var fviz_pca_ind get_pca_var
#' @importFrom ggplot2 ggplot aes geom_histogram geom_density stat_qq stat_qq_line labs theme
#' @importFrom patchwork wrap_plots plot_annotation
#' @importFrom psych KMO cortest.bartlett
#' @importFrom MVN mvn
#' @importFrom nortest lillie.test
#' @importFrom dplyr select select_if mutate tibble rownames_to_column
#' @importFrom openxlsx createWorkbook addWorksheet writeData saveWorkbook
#' @importFrom ggcorrplot ggcorrplot cor_pmat
#' @importFrom stats cor shapiro.test na.omit
#' 
#' @export
#'
#' @author [Raúl Fernández Contreras]
#'
#' @keywords multivariate PCA dimension-reduction

PCA_auto <- function(dataframe, 
                     where_to_save = NULL, 
                     title = paste0("PCA_", format(Sys.Date(), "%d_%m_%Y")),
                     color_variables,
                     scaled_data = F,
                     perform_assumptions = T){
  
  # ======================================================= #
  # ==== STEP 0: LOADING PACKAGES AND CUSTOM FUNCTIONS ==== 
  # ======================================================= #
  # Custom functions
  source("~/Documentos/09_scripts_R/create_sequential_dir.R")
  source("~/Documentos/09_scripts_R/automate_saving_dataframes_xlsx_format.R")
  source("~/Documentos/09_scripts_R/Automate_Saving_ggplots.R")
  source("~/Documentos/09_scripts_R/print_centered_note_v1.R")
  
  # Libraries
  list.of.packages = c('openxlsx', 'ggplot2', 'FactoMineR', 'factoextra', 'psych', 'MVN', 'dplyr', 
                       'corrr', 'corrplot', 'tidyr', "dplyr", 'ggpubr', "nortest", "patchwork",
                       "ggcorrplot", "ggplotify")
  
  new.packages = list.of.packages[!(list.of.packages %in% installed.packages())]
  if(length(new.packages) > 0) install.packages(new.packages)
  
  invisible(lapply(list.of.packages, FUN = library, character.only = T))
  rm(list.of.packages, new.packages)
  
  # Directory hierarchy
  where_to_save <- ifelse(is.null(where_to_save), getwd(), where_to_save)
  output_dir <- create_sequential_dir(path = where_to_save, name = title)
  figures <- create_sequential_dir(path = output_dir, name = "Figures")
  dataprocessed <- create_sequential_dir(path = output_dir, name = "Reports")
  
  wb <- createWorkbook()
  
  # ====================================================== #
  # ==== STEP 1: CHECKING ASSUPMTIONS AND CHECKPOINTS ====
  # ====================================================== #
  
  # Checkpoint 1: Check if all variables are numeric
  numeric_df <- dataframe %>%
    dplyr::select(-all_of(color_variables)) %>% 
    select_if(is.numeric)
  
  if(!(ncol(numeric_df) == (ncol(dataframe)-length(color_variables)))){
    stop("Please check that all variables in the dataset are numeric.")
  }else {
    print('Data only contains numeric variables. Continue')
  } # ifelse key 
  
  # Checkpoint 2: Assumptions
  if(perform_assumptions){
    normality_folder <- create_sequential_dir(path = dataprocessed, name = "Normality_Assumptions")
    
    # Univariate normality
    print_centered_note("Evaluating univariate normality ")
    
    # Univariate normality plots
    plot_normality <- function(variable_name, data) {
      # Extraer los datos de la variable
      x <- data[[variable_name]]
      
      # QQ-plot
      qq_plot <- ggplot(data.frame(x = x), aes(sample = x)) +
        stat_qq() +
        stat_qq_line() +
        labs(title = paste("QQ-plot for", variable_name),
             x = "Theoretical Quantiles", y = "Sample Quantiles")
      
      # Histograma con densidad
      hist_density_plot <- ggplot(data.frame(x = x), aes(x = x)) +
        geom_histogram(aes(y = ..density..), bins = 30, fill = "lightblue", color = "black") +
        geom_density(color = "red", linewidth = 1) +
        labs(title = paste("Distribution of", variable_name),
             x = variable_name, y = "Density")
      
      # Combinar los gráficos
      combined_plot <- qq_plot + hist_density_plot
      
      # Guardar el gráfico combinado
      save_ggplot(plot = combined_plot, 
                  title = paste0("Normality_", variable_name), 
                  folder = normality_folder, width = 12, height = 6, units = "in")
    }
    
    for (var in names(numeric_df)) {
      plot_normality(var, numeric_df)
    }
    
    # Univariate normality test
    if(nrow(dataframe) < 50){
      cat("Performing Lilliefors test for univariate normality\n")
      uninorm <- apply(numeric_df, 2, nortest::lillie.test)
      
      normality_assumptions <- data.frame(
        variable = names(uninorm),
        test = "Lilliefors",
        p_value = sapply(uninorm, function(x) x$p.value),
        distribution = ifelse(sapply(uninorm, function(x) x$p.value) > 0.05, 
                              "Normal Distribution", "Not Normal Distribution"),
        stringsAsFactors = FALSE
      )
      
      addWorksheet(wb, sheet = 'Univariate_Normality')
      writeData(wb, sheet = 'Univariate_Normality', normality_assumptions)
    
    } else{
      cat("Performing Shapiro-Wilk test for univariate normality\n")
      uninorm <- apply(numeric_df, 2, shapiro.test)
      
      normality_assumptions <- data.frame(
        variable = names(uninorm),
        test = "Shapiro Wilk",
        p_value = sapply(uninorm, function(x) x$p.value),
        distribution = ifelse(sapply(uninorm, function(x) x$p.value) > 0.05, 
                              "Normal Distribution", "Not Normal Distribution"),
        stringsAsFactors = FALSE
      )
      
      addWorksheet(wb, sheet = 'Univariate_Normality')
      writeData(wb, sheet = 'Univariate_Normality', normality_assumptions)
    } # ifelse key for univariate normality
    
    # Multivariate normality test
    print_centered_note("Evaluating multivariate normality")
    
    mvn_resultado <- mvn(
          data = numeric_df,
          mvnTest = 'mardia')$multivariateNormality
  
    mvn_resultado <- mvn_resultado %>%
      mutate(Statistic = as.numeric(as.character(Statistic)),
             "p value" = as.numeric(as.character(`p value`)),
             Result = ifelse(`p value` < 0.05, "Not normally", "Normally"))

    addWorksheet(wb, sheet = 'Multivariate_Normality')
    writeData(wb, sheet = 'Multivariate_Normality', mvn_resultado)
    
    # Multicolinearity
    print_centered_note("Evaluation of the multicolinearity ")
    
    # Calcular la matriz de correlación
    cor_matrix <- cor(numeric_df, use = "pairwise.complete.obs")
    
    addWorksheet(wb, sheet = 'Correlation_Matrix')
    writeData(wb, sheet = 'Correlation_Matrix', cor_matrix, rowNames = T)
    
    # Calcular la matriz de p-values (opcional, pero si quieres incluir significancia)
    # Usamos la función cor_pmat del paquete ggcorrplot
    p_matrix <- cor_pmat(numeric_df)
    
    addWorksheet(wb, sheet = 'Signif_Correlation_Matrix')
    writeData(wb, sheet = 'Signif_Correlation_Matrix', p_matrix, rowNames = T)
    
    # Crear el gráfico de correlación
    plot_corr <- ggcorrplot(
      cor_matrix, 
      #method = "circle", # o "square", "circle"
      type = "lower", # solo la triangular inferior
      hc.order = TRUE, # ordenar por hclust
      outline.color = "white",
      ggtheme = ggplot2::theme_minimal,
      colors = c("blue", "white", "red"),
      lab = TRUE, # mostrar coeficientes
      lab_size = 3, # tamaño de los coeficientes
      #p.mat = p_matrix, # matriz de p-values para insignificantes
      #sig.level = 0.05, # nivel de significancia
      #insig = "blank" # omitir las correlaciones no significativas
    ) +
      labs(title = "Correlation Plot") +
      theme(
        axis.text.x = element_text(angle = 45, hjust = 1, size = 12),
        axis.text.y = element_text(size = 12)
      )
    
    save_ggplot(plot_corr, title = "Corrplot_Colinearity", folder = figures, width = 2000, height = 2000)
    
    # Crear el gráfico de correlación
    signif_plot_corr <- ggcorrplot(
      cor_matrix, 
      #method = "circle", # o "square", "circle"
      type = "lower", # solo la triangular inferior
      hc.order = TRUE, # ordenar por hclust
      outline.color = "white",
      ggtheme = ggplot2::theme_minimal,
      colors = c("blue", "white", "red"),
      lab = TRUE, # mostrar coeficientes
      lab_size = 3, # tamaño de los coeficientes
      p.mat = p_matrix, # matriz de p-values para insignificantes
      sig.level = 0.05, # nivel de significancia
      insig = "blank" # omitir las correlaciones no significativas
    ) +
      labs(title = "Correlation Plot") +
      theme(
        axis.text.x = element_text(angle = 45, hjust = 1, size = 12),
        axis.text.y = element_text(size = 12)
      )
    
    save_ggplot(signif_plot_corr, title = "Corrplot_Colinearity_signification", folder = figures, width = 2000, height = 2000)
    
    # KMO
    colineality_test <- matrix(nrow =3, ncol = 3)
    colnames(colineality_test) = c('Test', 'p-value/KMO value', 'Interpretation')
    bartlett.res <- cortest.bartlett(numeric_df) ## We get colieality with p-value < 0.05
    colineality_test[1,] = c('Bartlett cor test', as.numeric(bartlett.res$p.value), ifelse(bartlett.res$p.value < 0.05, 'Significant colineality', 'Non-significant colineality'))
    
    tryCatch({
      kmo.res <- KMO(numeric_df) ## Nos fijamos en el Overall MSA
      colineality_test[2,] = c('KMO index', kmo.res$MSA, ifelse(kmo.res$MSA > 0.9, 'Strongly good',
                                                                ifelse(kmo.res$MSA < 0.9 & kmo.res$MSA > 0.8, 'Good',
                                                                       ifelse(kmo.res$MSA < 0.8 & kmo.res$MSA > 0.7, 'Aceptable',
                                                                              ifelse(kmo.res$MSA < 0.7 & kmo.res$MSA > 0.6, 'Normal',
                                                                                     ifelse(kmo.res$MSA < 0.6 & kmo.res$MSA > 0.5, 'Bad', 'Really bad'))))))
      
      
      
    }, error = function(e){
      print('KMO index can not be calculed')
      colineality_test[2,] = c('KMO index', NA, ifelse(kmo.res$MSA > 0.9, 'Strongly good',
                                                       ifelse(kmo.res$MSA < 0.9 & kmo.res$MSA > 0.8, 'Good',
                                                              ifelse(kmo.res$MSA < 0.8 & kmo.res$MSA > 0.7, 'Aceptable',
                                                                     ifelse(kmo.res$MSA < 0.7 & kmo.res$MSA > 0.6, 'Normal',
                                                                            ifelse(kmo.res$MSA < 0.6 & kmo.res$MSA > 0.5, 'Bad', 'Really bad'))))))
      
      
    })
    
    determinant <- det(cor_matrix)
    colineality_test[3,] = c('Determinant', determinant, 'Close to 0, the better. Absolute 0, worse.')
    colineality_test = as.data.frame(colineality_test)
    colineality_test$`p-value/KMO value` = as.numeric(colineality_test$`p-value/KMO value`)
    
    addWorksheet(wb, sheet = 'Colineality_Test')
    writeData(wb, sheet = 'Colineality_Test', colineality_test)
    
    saveWorkbook(wb = wb, file = file.path(dataprocessed, "Assumptions_Report.xlsx"))
  }#if key to evaluate statistical assumptions
  
  # ========================================================= #
  # ==== STEP 2: PERFORMING PRINCIPAL COMPONENT ANALYSIS ====
  # ========================================================= #
  
  print_centered_note("Performing PCA ")
  cat("\nRemember that PCA scales all variables inside. \n")
  
  # Checkpoint
  ## There can be no missing values to perform PCA.
  na_variables <- names(which(colSums(is.na(numeric_df)) > 0))
  
  if (length(na_variables) > 0) {
    print(na_variables)
    stop('There are NA values in these variables. Please check.')
  }
  
  # PCA
  wb_pca = createWorkbook()
  res.pca = PCA(na.omit(numeric_df), scale.unit=TRUE, ncp=ncol(numeric_df), graph=F)

  addWorksheet(wb_pca, sheetName = 'Eigenvalues')
  eigenvalues <- as.data.frame(res.pca$eig)
  eigenvalues <- eigenvalues %>%
    tibble::rownames_to_column('PC')
  
  eigenvalues$selected_component <- NA
  porcentaje_deseado <- 80
  acumulado <- 0
  
  for (i in 1:nrow(eigenvalues)) {
    acumulado <- acumulado + eigenvalues$`percentage of variance`[i]
    if (acumulado <= porcentaje_deseado) {
      
      eigenvalues$selected_component[1:i] <- "Use this component"
    }
  }
  
  eigenvalues$selected_component[is.na(eigenvalues$selected_component)] <- "Enough Components"
  writeData(wb_pca, sheet = 'Eigenvalues', eigenvalues)
  
  dimension_description <- dimdesc(res.pca, axes = 1:ncol(res.pca$ind$coord))
  for(dimension in seq(from = 1, to = ncol(numeric_df), length.out = length(numeric_df))){
    addWorksheet(wb_pca, sheetName = paste('Dim', dimension))
    my_description <- as.data.frame(dimension_description[dimension])
    my_description <- my_description %>%
      tibble::rownames_to_column('Variables')
    writeData(wb_pca, sheet = paste('Dim', dimension), my_description)
  }
  
  saveWorkbook(wb_pca, file = file.path(dataprocessed, paste0(paste('PCA Results ', title, Sys.Date(), sep = ' '), '.xlsx')))
  
  # =========================== #
  # ==== STEP 3: PCA PLOTS ====
  # =========================== #
  print_centered_note("Creating PCA visualizations")
  
  # Screeplot
  cat('\nScreeplot\n')
  scree_plot = fviz_eig(res.pca, addlabels = TRUE, ylim = c(0, 50)) ##Gráfico sedimentación
  save_ggplot(plot = scree_plot, title = "ScreePlot", folder = figures, height = 3000, width = 3000)
  
  # Variable Contribution I
  cat('\nVariable Contribution Plot\n')
  pca_var_contribution_plot = fviz_pca_var(res.pca,
                                           col.var = "contrib", # Color by contributions to the PC
                                           gradient.cols = c("#00AFBB", "#E7B800", "#FC4E07"),
                                           repel = TRUE     # Avoid text overlapping
  )
  save_ggplot(plot = pca_var_contribution_plot, 
              title = "PCA_Variable_Contribution_PC1_PC2", 
              folder = figures, width = 2000, height = 2000)
  
  # Variable Contribution II
  var <- get_pca_var(res.pca)
  cat('\nContribution variable for each component plot\n')
  png(filename = file.path(figures, "05_Contribution_Each_Component.png"), 
      res = 300, width = 2000, height = 2000, units = "px")
  contribution_each_component_plot_gg <- corrplot(var$cos2, is.corr=FALSE)
  dev.off()
  
  # Individual Contribution
  individual_pca_contribution = fviz_pca_ind(res.pca, pointsize = "contrib",
                                             pointshape = 21, fill = "#E7B800",
                                             repel = TRUE)
  
  save_ggplot(individual_pca_contribution, title = "Individual_Contribution", folder = figures,
              width = 3000, height = 3000)
  
  # PCA indiviuals by groups
  cat('\nIndividual PCA plot colored by selected variables\n')
  if(length(color_variables) != 0){
    for(variable in color_variables){
      groups <- unique(dataframe[[variable]])
      num_gruops <- length(groups)
      colors <- rainbow(num_gruops)
      #png(filename = paste0(paste('PCA Colored Individuals by', variable), '.png'), res = 300, height = 3000, width = 3000)
      fviz_ind <- fviz_pca_ind(res.pca,
                               col.ind = dataframe[[variable]], # color by groups
                               palette = colors,
                               axes = c(1, 2),
                               addEllipses = TRUE, # Concentration ellipses
                               ellipse.type = "confidence",
                               legend.title = variable,
                               label = "none",
                               repel = TRUE, 
                               title = paste0("Individuals - PCA Colored By ", variable)
      )
      save_ggplot(plot = fviz_ind, title = paste0("PCA_Ind_Colored_by_", variable), folder = figures,
                  width = 2000, height = 2000)
    } # for loop key for multiple color variables
  }# If key for color individuals
  
  # Patchwork
  plot_list <- list(pca_var_contribution_plot, fviz_ind, individual_pca_contribution, plot_corr)
  plots <- wrap_plots(plot_list, ncol = 2) + 
    plot_annotation(tag_levels = 'A')
  save_ggplot(plot= plots, 
              title = "Combined_PCA_Figures_Plot", 
              folder = output_dir,
              width = 5000, height = 4000) 
  print_centered_note("End of the script")
  return(res.pca)
}# Main function key
