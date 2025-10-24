########## 11/04/2025 ##########
#' Automated Linear and Generalized Mixed-Effects Model Fitting with Reporting
#'
#' This function automates the fitting of linear mixed-effects models (\code{lmer}) 
#' or generalized mixed-effects models (\code{glmer}) based on a user-defined formula and dataset.
#' It creates a sequential results folder and saves:
#' - Diagnostic plots of the fitted model  
#' - Forest plots of estimated coefficients  
#' - Effect plots and pairwise contrast plots  
#' - Summary tables in HTML, DOCX, and PNG formats  
#'
#' It also installs and loads any required R packages that are not already installed.
#'
#' @param formula_lmm \code{formula}. Model formula, e.g. \code{y ~ x1 + x2 + (1|group)}.
#' @param dataframe \code{data.frame}. Dataset used to fit the model.
#' @param where_to_save \code{character}, optional. Path where the results will be saved. 
#' If \code{NULL}, the current working directory is used. 
#' A sequential folder with the format \code{"fit_XX"} will be created.
#' @param transform_coeff \code{character}. Transformation applied to model coefficients in the HTML report. 
#' Defaults to \code{"exp"}. Accepts any option supported by \code{sjPlot::tab_model}.
#' @param family \code{character} or \code{family}. Distribution family for the model. 
#' Defaults to \code{"gaussian"} for linear mixed models. 
#' For generalized mixed models, specify e.g. \code{"binomial"}.
#'
#' @details 
#' The function automates a typical workflow for mixed-effects model analysis in R:
#' \enumerate{
#'   \item Install and load required packages if not available.
#'   \item Create a sequential results folder (e.g., \code{fit_01}, \code{fit_02}, ...).
#'   \item Fit a mixed-effects model using \code{lmer} or \code{glmer}.
#'   \item Generate and save diagnostic plots, forest plots, effect plots, and pairwise contrasts.
#'   \item Export model summaries in multiple formats (HTML, DOCX, PNG).
#' }
#'
#' Graphs are saved using a helper function called \code{save_ggplot}, which must be available in the environment.
#'
#' @return 
#' (Invisibly) returns the fitted model object (\code{lmerMod} or \code{glmerMod}) for further use.
#'
#' @seealso 
#' \code{\link[lme4]{lmer}}, \code{\link[lme4]{glmer}}, \code{\link[sjPlot]{tab_model}}, 
#' \code{\link[ggplot2]{ggplot}}, \code{\link[ggeffects]{ggpredict}}, \code{\link[emmeans]{emmeans}}
#'
#' @examples
#' \dontrun{
#' # Example: Fit a linear mixed-effects model with random intercepts by Subject
#' library(lme4)
#' data("sleepstudy")
#' mixed_effect_lm(
#'   formula_lmm = Reaction ~ Days + (1 | Subject),
#'   dataframe = sleepstudy,
#'   where_to_save = "model_results",
#'   transform_coeff = "exp",
#'   family = "gaussian"
#' )
#' }
#'
#' @export
mixed_effect_lm <- function(formula_lmm, 
                             dataframe, 
                             where_to_save = NULL, 
                             transform_coeff = "exp",
                             family = "gaussian") {
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
  response_var <- all.vars(as.formula(formula_lmm))[1]
  type_response_var <- class(dataframe[[response_var]]) # Tipo de la variable respuesta
  
  # Extraer variables independientes (fijas)
  vars_indep <- attr(terms(as.formula(formula_lmm)), "term.labels")
  vars_indep <- vars_indep[!grepl("\\|", vars_indep)] # Filtrar solo efectos fijos (quitar términos con '|')
  
  # Tipos de cada variable fija
  tipos_indep <- sapply(vars_indep, function(v) class(dataframe[[v]]))
  print(unlist(tipos_indep))
  
  # =============================== #
  # STEP1 --> Ajustar el modelo LMM #
  # =============================== #
  #dataframe <- dataframe
  
  if(family == "gaussian"){
    modelo <- lmer(formula_lmm, data = dataframe)
  }else{
    modelo <- glmer(formula_lmm, data = dataframe, family = family)
    
  }
  
  # =========================== #
  # STEP2 --> Create some Plots #
  # =========================== #
  # 1) Model assumptions
  print_centered_note("Creating Some Visualizations ")
  cat("\n1.Diagnostics Plots \n")
  # Plot the result of check_model
  diagn_fig = plot(check_model(modelo))
  save_ggplot(plot = diagn_fig, title = "Diagnostic", folder = ruta_completa, width = 3750, height = 3000)
  
  
  # 2) Forest Plot of Estimates
  cat("\n2.ForestPlot Plot Of Estimates \n")
  forest_plot <- plot_model(modelo, type = "est", show.values = TRUE, value.offset = .3, width = 0.2
                            #title = as.character(formula_lmm)
  )
  
  save_ggplot(forest_plot, title = "ForestPlot", folder = ruta_completa,
              width = 2000, height = 2000)
  
  # 3) Estimates Plot for Each Contrast
  # 3) Estimates Plot for Each Contrast (one plot per factor)
if(any(tipos_indep %in% c("character", "factor"))){
  
  cat("\n3. By Contrast Effect Plot (one per factor variable)\n")
  
  # Iterar sobre las variables categóricas
  for(var_factor in names(tipos_indep)[tipos_indep %in% c("character", "factor")]){
    print(paste0("Generating contrast plot for factor: ", var_factor))
    
    # Calcular los contrastes de esa variable individualmente
    res_factor <- emmeans(fit, as.formula(paste0("pairwise ~ ", var_factor)), infer = TRUE)
    
    # Crear el gráfico de los contrastes
    pw_plot <- plot(res_factor$contrasts) +
      ggtitle(paste("Pairwise Contrasts for", var_factor))
    
    # Guardar el gráfico en la carpeta correspondiente
    save_ggplot(
      plot = pw_plot,
      title = paste0("Contrast_Estimates_", var_factor),
      folder = ruta_completa,
      width = 2000,
      height = 2000
    )
  }
}
  # 4) Effect Plot
  cat("\n4.Effect Plot \n")
  for(independ_var in seq(1, length(vars_indep))){
    print(paste0("Saving Effect Plot for variable ", vars_indep[independ_var]))
    eff <- plot(ggpredict(modelo, terms = vars_indep[independ_var]))
    
    #effect_plot <- plot_model(modelo, type = "eff")[[independ_var]]
    save_ggplot(plot = eff, title = paste0("EffectPlot_", vars_indep[independ_var]), 
                folder = ruta_completa, width = 7, height = 4.5, units = "in")
  }
  
  if(length(vars_indep) == 2){
    print(paste0("Saving Effect Plot for both variables"))
    eff <- plot(ggpredict(modelo, terms = vars_indep))
    save_ggplot(plot = eff, title = paste0("EffectPlot_", paste0(vars_indep, collapse = "_")), 
                folder = ruta_completa, width = 7, height = 4.5, units = "in")
  }
  
  # ===================================== #
  # STEP3 --> Resumen del modelo en .html #
  # ===================================== #
  print_centered_note("Saving Summary of the Model in .html ")
  # 1) Overall model summary
  cat("\nSaving overall model summary\n")
  print(tab_model(modelo, file = paste0(ruta_completa, paste0("/", nombre_carpeta,"_LMM_Summary.html")),
                  transform = transform_coeff,
                  show.reflvl = T, 
                  show.intercept = F,
                  p.style = "numeric_stars",
                  show.aic = T,
                  show.aicc = T))
  
  #2) Contrast model summary (if > 2 categories X variable)
  # Save tables in pdf and word format
  cat("\nSaving contrast summary\n")
  t <- tbl_regression(modelo, add_pairwise_contrast = T) %>%
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
  return(invisible(modelo))
}
