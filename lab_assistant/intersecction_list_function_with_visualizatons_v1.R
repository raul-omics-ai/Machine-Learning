######## 08/06/2026 ##########
#' Plot intersections using Venn Diagram or UpSet Plot
#'
#' Computes all possible intersections from a list of sets and:
#' \itemize{
#'   \item Saves a dataframe with all intersection combinations in wide format
#'   \item Generates a visualization:
#'     \itemize{
#'       \item Venn Diagram (if number of sets \eqn{\le} 3)
#'       \item UpSet Plot (if number of sets \eqn{>} 3)
#'     }
#' }
#'
#' @param intersection_list A named list of vectors representing sets to intersect.
#' Each element must be named and contain comparable values.
#' @param where_to_save Character string specifying the output directory.
#' Default is \code{getwd()}.
#' @param title Character string used as the base name for output files.
#'
#' @details
#' This function:
#' \enumerate{
#'   \item Validates that the input list is named
#'   \item Computes all possible combinations of intersections
#'   \item Stores results in a wide-format dataframe
#'   \item Saves the dataframe using a custom helper function
#'   \item Generates and saves either a Venn diagram or an UpSet plot
#' }
#'
#' It relies on external utility scripts:
#' \itemize{
#'   \item \code{print_centered_note_v1.R}
#'   \item \code{Automate_Saving_ggplots.R}
#'   \item \code{automate_saving_dataframes_xlsx_format.R}
#' }
#'
#' Required packages are automatically installed if missing.
#'
#' @return Invisibly returns \code{NULL}. Outputs are saved to disk.
#'
#' @examples
#' \dontrun{
#' example_list <- list(
#'   A = c("gene1", "gene2", "gene3"),
#'   B = c("gene2", "gene3", "gene4"),
#'   C = c("gene3", "gene5")
#' )
#'
#' plot_intersection(
#'   intersection_list = example_list,
#'   where_to_save = "./results",
#'   title = "example_intersection"
#' )
#' }
#'
#' @import ggplot2
#' @import dplyr
#' @import tidyr
#' @import openxlsx
#'
#' @export
plot_intersection <- function(intersection_list, 
                              where_to_save = NULL, 
                              title = "Intersection_Result"){
  
  # Load custom scripts
  source("~/Documentos/09_scripts_R/print_centered_note_v1.R")
  source("~/Documentos/09_scripts_R/Automate_Saving_ggplots.R")
  source("~/Documentos/09_scripts_R/automate_saving_dataframes_xlsx_format.R")
  
  # Load packages
  print_centered_note('LOADING PACKAGES ')
  
  list_of_packages = c("ggplot2", "openxlsx", "dplyr", "tidyr")
  
  new_packages = list_of_packages[!(list_of_packages %in% installed.packages())]
  if(length(new_packages) > 0){
    install.packages(new_packages)
  }
  
  invisible(lapply(list_of_packages, FUN = library, character.only = TRUE))
  
  rm(list_of_packages, new_packages)
  print('Packages loaded successfully.')
  
  # Checkpoint 1: Validate names
  n_sets <- length(intersection_list)
  set_names <- names(intersection_list)
  
  if (is.null(set_names)) {
    stop("The input list must contain named elements (set identifiers).")
  }
  
  # Checkpoint 2: Output path
  if(is.null(where_to_save)){
    where_to_save <- getwd()
  }
  
  # Generate combinations
  combinations <- unlist(
    lapply(1:length(intersection_list), function(x) 
      combn(set_names, x, simplify = FALSE)),
    recursive = FALSE
  )
  
  # Compute intersections
  results <- lapply(combinations, function(cmb) {
    Reduce(intersect, intersection_list[cmb])
  })
  
  names(results) <- sapply(combinations, paste, collapse = "_")
  
  # Convert to dataframe
  df <- stack(results)
  colnames(df) <- c("Element", "Intersection")
  
  df_wide <- df %>%
    dplyr::group_by(Intersection) %>%
    dplyr::mutate(id = dplyr::row_number()) %>%
    tidyr::pivot_wider(
      names_from = Intersection,
      values_from = Element
    ) %>%
    dplyr::select(-id)
  
  print("Saving combination dataframe")
  save_dataframe(df = df_wide, title = title, folder = where_to_save)
  
  # Visualization
  print_centered_note(toupper("Making Plots "))
  
  if(n_sets <= 3){
    print("Saving Venn Diagram")
    library(ggvenn)
    
    ven <- ggvenn(
      intersection_list, 
      fill_color = c("#0073C2FF", "#EFC000FF", "#868686FF"),
      stroke_size = 0.5, 
      set_name_size = 4
    )
    
    save_ggplot(
      ven, 
      title = paste0("Venn_Diagram_", title), 
      folder = where_to_save, 
      width = 2000,
      height = 2000
    )
    
  } else {
    print("Saving Upset Plot")
    library(UpSetR)
    
    saving_path <- file.path(where_to_save, paste0("UpsetPlot_", title, ".png"))
    png(saving_path, width = 3000, height = 2000, res = 300)
    
    print(upset(fromList(intersection_list), order.by = "freq"))
    
    dev.off()
  }
  
  print_centered_note(toupper("End of the function"))
}
