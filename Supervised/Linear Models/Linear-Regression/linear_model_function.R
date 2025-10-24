########## 23/10/2025 ##########
#' @title Automatic Linear Model Fitting, Diagnosis, and Evaluation
#' @description 
#' This function automatically fits a linear model (`lm`) to a given dataset,
#' performs diagnostics, generates visualizations of effects and contrasts,
#' and saves outputs (plots, summaries, and reports) in a structured folder.
#'
#' @param formula_lm A formula specifying the linear model (e.g. `y ~ x1 + x2`).
#' @param dataframe A data frame containing the variables used in the model.
#' @param where_to_save Optional path to a directory where results will be saved. 
#' If not provided, the current working directory is used.
#'
#' @details 
#' The function performs the following main steps:
#' 1. Loads and installs required packages.
#' 2. Creates a sequential folder (`fit_XX`) for model outputs.
#' 3. Fits the linear model using `lm()`.
#' 4. Generates diagnostic plots, forest plots, contrast plots, and effect plots.
#' 5. Saves model summaries and contrast tables in HTML, DOCX, and PNG formats.
#'
#' @return The fitted model object (of class `lm`), returned invisibly.
#' @export
#'
#' @examples
#' \dontrun{
#' linear_model_function(Sepal.Length ~ Species, iris)
#' }
linear_model_function <- function(formula_lm, 
                         dataframe, 
                         where_to_save = NULL 
                         ){
  # ===================================================== #
  # STEP 0 --> Setting up functions and working directory #
  # ===================================================== #
  source("~/Documentos/09_scripts_R/print_centered_note_v1.R")
  source("~/Documentos/09_scripts_R/Automate_Saving_ggplots.R")
  source("~/Documentos/09_scripts_R/automate_saving_dataframes_xlsx_format.R")
  source("~/Documentos/09_scripts_R/create_sequential_dir.R")
  
  print_centered_note('Loading Packages ')
  list_of_packages = c("sjPlot", "parameters", "performance", "dplyr", "tidyr", "lme4", 
                       "lmerTest", "gtsummary", "emmeans", "flextable", "ggeffects", "ggplot2")
  
  new_packages = list_of_packages[!(list_of_packages %in% installed.packages())]
  if(length(new_packages) > 0){install.packages(new_packages)}
  
  invisible(lapply(list_of_packages, FUN = library, character.only = T))
  rm(list_of_packages, new_packages)
  print('Packages loaded successfully.')
  print_centered_note("Setting Up Working Directory ")
  
  if(is.null(where_to_save)){
    where_to_save <- getwd()
  }
  
  carpetas <- list.dirs(where_to_save, full.names = FALSE, recursive = FALSE)
  carpetas_fit <- grep("^fit_\\d{2}$", carpetas, value = TRUE) # Filtrar las que coinciden con el patrón "fit_XX"
  
  numeros <- gsub("^fit_(\\d{2})$", "\\1", carpetas_fit) # Extraer los números y convertirlos a enteros
  numeros <- as.integer(numeros)
  
  siguiente <- if (length(numeros) == 0) 1 else max(numeros) + 1 # Determinar el siguiente número
  
  # Crear nombre con formato "fit_XX"
  nombre_carpeta <- sprintf("fit_%02d", siguiente)
  
  ruta_completa <- create_sequential_dir(path = where_to_save, name = nombre_carpeta)
  message("Carpeta creada: ", ruta_completa)
  
  # ================================================= #
  # STEP1 --> Set the types of variables in the model #
  # ================================================= #
  
  # Extraer variable respuesta
  response_var <- all.vars(as.formula(formula_lm))[1]
  type_response_var <- class(dataframe[[response_var]]) # Tipo de la variable respuesta
  
  # Extraer variables independientes (fijas)
  vars_indep <- attr(terms(as.formula(formula_lm)), "term.labels")
  vars_indep <- vars_indep[!grepl("\\|", vars_indep)] # Filtrar solo efectos fijos (quitar términos con '|')
  
  # Tipos de cada variable fija
  tipos_indep <- sapply(vars_indep, function(v) class(dataframe[[v]]))
  print(unlist(tipos_indep))
  
  # =============================== #
  # STEP1 --> Ajustar el modelo LMM #
  # =============================== #
  print_centered_note("Fitting Linear Model ")
  fit <- lm(as.formula(formula_lm), data = dataframe)
  
  # =========================== #
  # STEP2 --> Create some Plots #
  # =========================== #
  # 1) Model assumptions
  print_centered_note("Evaluating the model ")
  cat("\n1.Diagnostics Plots \n")
  # Plot the result of check_model
  diagn_fig = plot(check_model(fit))
  save_ggplot(plot = diagn_fig, title = "Diagnostic", folder = ruta_completa, width = 3750, height = 3000)
  
  
  # 2) Forest Plot of Estimates
  cat("\n2.ForestPlot Plot Of Estimates \n")
  forest_plot <- plot_model(fit, type = "est", show.values = TRUE, value.offset = .3, width = 0.2
                            #title = as.character(formula_lmm)
  )
  
  save_ggplot(forest_plot, title = "ForestPlot", folder = ruta_completa,
              width = 2000, height = 2000)
  
  # 3) Estimates Plot for Each Contrast
  if(any(tipos_indep %in% c("character", "factor"))){
    fixed_term <- sub(".*~", "", formula_lm) # extraer RHS (todo lo que está después de "~")
    fixed_term <- gsub("\\([^)]*\\|[^)]*\\)", "", fixed_term)  # " Treatment + "
    fixed_term <- gsub("\\+\\s*\\+", "+", fixed_term)   # quita ++ si aparecen
    fixed_term <- gsub("^\\s*\\+|\\+\\s*$", "", fixed_term) # quitar + inicial / final
    fixed_term <- trimws(fixed_term) 
    
    res <- emmeans(fit, as.formula(paste0("pairwise ~ ", fixed_term)), infer = T) 
    # Esto nos vale para calcular la media de los errores y sus  p-values, pero raravez se utilizan. 
    # Lo voy a utilizar para generar un gráfico con las estimaciones de  todos los contrastes
    
    cat("\n3.By Contrast Effect Plot \n")
    pw_log_estimates <- res$contrasts %>% plot()
    save_ggplot(pw_log_estimates, title = "Contrast_Estimates_Plot_log", 
                folder = ruta_completa,
                width = 2000, height = 2000)
  }#if key for contrats estimates plot
  
  # 4) Effect Plot
  cat("\n4.Effect Plot \n")
  for(independ_var in seq(1, length(vars_indep))){
    print(paste0("Saving Effect Plot for variable ", vars_indep[independ_var]))
    eff <- plot(ggpredict(fit, terms = vars_indep[independ_var]))
    
    #effect_plot <- plot_model(modelo, type = "eff")[[independ_var]]
    save_ggplot(plot = eff, title = paste0("EffectPlot_", vars_indep[independ_var]), 
                folder = ruta_completa, width = 7, height = 4.5, units = "in")
  }
  
  if(length(vars_indep) == 2){
    print(paste0("Saving Effect Plot for both variables"))
    eff <- plot(ggpredict(fit, terms = vars_indep))
    save_ggplot(plot = eff, title = paste0("EffectPlot_", paste0(vars_indep, collapse = "_")), 
                folder = ruta_completa, width = 7, height = 4.5, units = "in")
  }
  
  # ===================================== #
  # STEP3 --> Resumen del modelo en .html #
  # ===================================== #
  print_centered_note("Saving Summary of the Model in .html ")
  # 1) Overall model summary
  cat("\nSaving overall model summary\n")
  print(tab_model(fit, file = paste0(ruta_completa, paste0("/", nombre_carpeta,"_LMM_Summary.html")),
                  show.reflvl = T, 
                  show.intercept = F,
                  p.style = "numeric_stars",
                  show.aic = T,
                  show.aicc = T))
  
  #2) Contrast model summary (if > 2 categories X variable)
  # Save tables in pdf and word format
  cat("\nSaving contrast summary\n")
  t <- tbl_regression(fit, add_pairwise_contrast = T) %>%
    add_significance_stars(
      hide_p = F,
      hide_se = T,
      hide_ci = F) %>%
    bold_p()
  
  t %>%
    as_flex_table() %>%
    save_as_docx(path = file.path(ruta_completa, "Contrast_LMM_Summary.docx"))
  
  t %>%
    as_flex_table() %>%
    save_as_image(path = file.path(ruta_completa, "Contrast_LMM_Summary.png"))
  
  # 6. Devolver el modelo por si lo quieres guardar o seguir usando
  print_centered_note("End of the script")
  return(invisible(fit))
  
} # main function key