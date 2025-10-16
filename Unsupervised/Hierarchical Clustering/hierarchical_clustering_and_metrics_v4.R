########## 26/02/2025 ##########

# =================================================== #
# ==== Hierarchical clustering automate function ==== #
# =================================================== #

#' Comprehensive Hierarchical Clustering Analysis with Validation Metrics
#'
#' This function performs a complete hierarchical clustering analysis that includes:
#' correlation calculations, data normalization, optimal cluster number determination,
#' internal and external validation, and comprehensive report generation.
#'
#' @param metadata Dataframe with data for analysis. Can contain numeric and
#'   categorical variables.
#' @param vect.external.val Optional vector with external labels for clustering
#'   validation. Default: NULL
#' @param where_to_save Directory to save results. If NULL, uses current working
#'   directory. Default: NULL
#' @param title Title for the analysis and output directory names.
#'   Default: "Hierarchical_clustering"
#' @param max_k_number Maximum number of clusters to evaluate. Default: 50
#' @param nboot Number of bootstrap iterations for gap statistic calculation. Default: 500
#'
#' @return Dataframe with cluster assignments for each observation.
#'   Additionally generates Excel reports and plots in the specified directory.
#'
#' @details
#' The function automates the entire hierarchical clustering process:
#' \itemize{
#'   \item{Calculates appropriate correlations based on data type (Pearson, Spearman, or mixed)}
#'   \item{Normalizes numeric variables}
#'   \item{Determines best distance metric and agglomeration method}
#'   \item{Evaluates optimal cluster count using multiple methods}
#'   \item{Generates professional visualizations (dendrograms, silhouettes, PCA)}
#'   \item{Provides internal and external validation}
#' }
#'
#' @examples
#' \dontrun{
#' # Example with numeric data
#' data(iris)
#' iris_data <- iris[, 1:4]
#' result <- hierarchical_clustering_and_metrics(
#'   metadata = iris_data,
#'   title = "Iris_Clustering",
#'   max_k_number = 10
#' )
#'
#' # Example with external validation
#' result_ext <- hierarchical_clustering_and_metrics(
#'   metadata = iris_data,
#'   vect.external.val = iris$Species,
#'   where_to_save = "./results",
#'   title = "Iris_With_Validation"
#' )
#' }
#'
#' @export
#' @importFrom vegan vegdist
#' @importFrom psych mixedCor polychoric tetrachoric
#' @importFrom polycor hetcor
#' @importFrom openxlsx createWorkbook addWorksheet writeData writeDataTable saveWorkbook
#' @importFrom readxl read_excel
#' @importFrom dplyr mutate across all_of mutate_if group_by summarise arrange desc
#' @importFrom cluster daisy silhouette
#' @importFrom StatMatch dist
#' @importFrom ggcorrplot ggcorrplot
#' @importFrom ggdendro ggdendrogram
#' @importFrom tibble rownames_to_column
#' @importFrom GGally ggpairs
#' @importFrom corrplot corrplot
#' @importFrom MASS isoMDS
#' @importFrom viridis viridis
#' @importFrom tidyr pivot_longer pivot_wider
#' @importFrom factoextra fviz_nbclust fviz_dend fviz_silhouette fviz_cluster fviz_dist
#' @importFrom igraph graph.adjacency
#' @importFrom Hmisc rcorr
#' @importFrom purrr map map_df
#' @importFrom psychTools bfi
#' @importFrom Rtsne Rtsne
#' @importFrom patchwork plot_annotation
#' @importFrom FactoMineR PCA MCA
#' @importFrom fpc cluster.stats
#' @importFrom NbClust NbClust
#' @importFrom ggplot2 ggplot geom_bar theme_minimal labs scale_y_continuous
#'   theme element_text
#' @importFrom stats dist hclust cutree cophenetic shapiro.test
#' @importFrom RColorBrewer brewer.pal
#' @importFrom cluster eclust

hierarchical_clustering_and_metrics <- function(metadata,
                                                vect.external.val = NULL,
                                                where_to_save = NULL,
                                                title = "Hierarchical_clustering",
                                                max_k_number = 30,
                                                nboot = 500,
                                                ties = "max"){
  # ---- BLOCK 1: Loading packages and custom functions ----
  source("~/Documentos/09_scripts_R/print_centered_note_v1.R")
  source("~/Documentos/09_scripts_R/create_sequential_dir.R")
  # Loading packages
  print_centered_note("LOADING REQUIRED PACKAGES AND FUNCTIONS")
  
  list_of_packages = c("vegan", "psych","polycor","openxlsx","readxl","dplyr","cluster","StatMatch",
                       "ggcorrplot", "ggdendro", "tibble", "GGally", "corrplot", "MASS", "viridis",        
                       "tidyr", "factoextra", "igraph", "Hmisc", "purrr", "psychTools", "Rtsne", 
                       "patchwork", "FactoMineR",  "fpc", "NbClust", "ggplot2","stats", "RColorBrewer")
  new_packages = list_of_packages[!(list_of_packages %in% installed.packages())]
  if(length(new_packages) > 0){install.packages(new_packages)}
  
  invisible(lapply(list_of_packages, FUN = library, character.only = T))
  rm(list_of_packages, new_packages)
  print('Packages loaded successfully.')
  
  # Saving directory
  where_to_save <- ifelse(is.null(where_to_save), getwd(), where_to_save)
  output_dir <- create_sequential_dir(path = where_to_save, name = paste0(title, "_", Sys.Date()))
  
  # ---- BLOCK 2: Computing correlation coeficients ---- 
  # In this step, I'll calculate correlation coefficients based on the type of
  # variable in the given dataset.
  
  # Check the best correlation coefficient:
  #   1) If all variables are numeric: Pearson/Spearman correlation coefficient
  #   2) If there are a mix of numeric and categorical variables:
  #       - When categorical variables had more than 2 categories: Polychoric Correlation
  #       - When categorical variable had only two categories: Tetrachoric Correlation
  
  # ---- CUSTOM FUNCTIONS
  cors <- function(df, type_corr) {
    # type_corr = c("pearson","spearman")
    M <- Hmisc::rcorr(as.matrix(df), type = type_corr)
    # turn all three matrices (r, n, and P into a data frame)
    Mdf <- purrr::map(M, ~data.frame(.x))
    # return the three data frames in a list
    return(Mdf)
  }
  
  formatted_cors <- function(df, type_corr){
    cors(df, type_corr) %>%
      map(~rownames_to_column(.x, var="measure1")) %>%
      # format each data set (r,P,n) long
      map(~pivot_longer(.x, -measure1, names_to ="measure2")) %>%
      # merge our three list elements by binding the rows
      bind_rows(.id = "id") %>%
      pivot_wider(names_from = id, values_from = value) %>%
      # change so everything is lower case
      rename(p = P) %>%
      mutate(sig_p = ifelse(p < .001, '***', 
                            ifelse(p < 0.001, '**',
                                   ifelse(p < 0.05, '*', 'NS'))),
             p_if_sig = ifelse(sig_p %in% c('***', '**', '*'), p, NA),
             r_if_sig = ifelse(sig_p %in% c('***', '**', '*'), r, NA))
  }
  
  scale2 <- function(x, na.rm = F){(x-min(x, na.rm = T))/(max(x, na.rm = T)-min(x, na.rm = T))}
  
  dir.create(path = file.path(output_dir, "01_Correlations"), showWarnings = F, recursive = T)
  hc_wb <- createWorkbook()
  
  # If all columns is numeric -> Pearson/Spearman correlation
  if(sum(sapply(metadata, is.numeric)) == ncol(metadata)){
    # ---- Block 2.1: All variables are numeric ----
    print_centered_note("ALL VARIABLES ARE NUMERIC. CHECKING NORMALITY: ")
    
    normality <- apply(metadata, 2, shapiro.test)
    normality.df <- as.data.frame(matrix(ncol=2, nrow = 0))
    colnames(normality.df) = c('Variable', 'P.value')
    
    for(variable in 1:length(normality)){
      normality.df[variable , 1] = names(normality)[variable]
      normality.df[variable, 2] = normality[[variable]]$p.value
    }
    
    normality.df$interpretation = ifelse(normality.df$P.value < 0.05, 'Not Normally distributed', 'Normally Distributed')
    if(sum(normality.df$P.value > 0.05) == ncol(metadata)){
      # ---- Block 2.1.1: All numeric variables are normally distributed ----
      # If all variables are normally distributed -> Pearson
      cat("\nAll variables are normally distributed. \nApplying Pearson correlation")
      corr.res <- cors(metadata, type_corr = 'pearson')
      # Pearson - Corrplot
      col<- colorRampPalette(c("blue", "white", "red"))
             
      png(file.path(output_dir, "01_Correlations", "01_Pearson_Correlation_Corrplot.png"), 
          width = 3000, height = 3000, res = 300)
      plot_corr = corrplot(as.matrix(corr.res$r), method="color", col=col(200),  
                           type="lower", 
                           order="hclust", ## "original", "AOE", "FPC", "hclust", "alphabet"
                           addCoef.col = "black", # Add coefficient of correlation
                           tl.col="black", 
                           tl.srt=45, 
                           tl.cex = 1, 
                           number.cex = 0.6,#Text label color and rotation
                           # Combine with significance
                           p.mat = as.matrix(corr.res$P), 
                           sig.level = 0.05, 
                           insig = "blank", # hide correlation coefficient on the principal diagonal
                           diag=F,
                           title = "Rho Pearson Correlation Coefficients. Only significant correlations are shown")         
      dev.off()
      
      # ---- Statistical reprot
      report = formatted_cors(metadata, type_corr = 'pearson')
      report$corr_coef <- 'Pearson'
      
      #wb = createWorkbook(title = paste('Correlation and P-val'))
      addWorksheet(hc_wb, sheetName = "Correlation_Rho")
      writeDataTable(hc_wb, sheet = "Correlation_Rho", x = corr.res$r, rowNames = T)
      
      addWorksheet(hc_wb, sheetName = "Correlation_pval")
      writeDataTable(hc_wb, sheet = "Correlation_pval", corr.res$P, rowNames = T)
      
      addWorksheet(hc_wb, 'Correlation_Summary')
      writeData(hc_wb, sheet = 'Correlation_Summary', report)
      #saveWorkbook(hc_wb, file = file.path(output_dir, "01_Correlations", "Correlation and P-val report.xlsx"))
      
    }else{
      # ---- Block 2.1.2: Not all numerical variables are normally distributed ----
      # If at least one variable is not normally distributed -> Spearman
      cat("There are at least one variable not normally distributed. \nApplying Spearman correlation")
      corr.res <- cors(metadata, type_corr = 'spearman')
      
      # Spearman - Corrplot
      col<- colorRampPalette(c("blue", "white", "red"))
      
      png(file.path(output_dir, "01_Correlations", "01_Corrplot_Spearman.png"),
          width = 3000, height = 3000, res = 300)
      plot_corr = corrplot(as.matrix(corr.res$r), method="color", col=col(200),  
                           type="lower", 
                           order="hclust", ## "original", "AOE", "FPC", "hclust", "alphabet"
                           addCoef.col = "black", # Add coefficient of correlation
                           tl.col="black", 
                           tl.srt=45, 
                           tl.cex = 1, 
                           number.cex = 0.6,#Text label color and rotation
                           # Combine with significance
                           p.mat = as.matrix(corr.res$P), 
                           sig.level = 0.05, 
                           insig = "blank", # hide correlation coefficient on the principal diagonal
                           diag=F,
                           title = "Rho Spearman Correlation Coefficients. Only significant correlations are shown")         
      
      dev.off()
      
      # ---- Statistical Report
      report = formatted_cors(metadata, type_corr = 'spearman')
      report$corr_coef <- 'Spearman'
      
      addWorksheet(hc_wb, sheetName = "Correlation_Rho")
      writeDataTable(hc_wb, sheet = "Correlation_Rho", x = corr.res$r, rowNames = T)
      
      addWorksheet(hc_wb, sheetName = "Correlation_pval")
      writeDataTable(hc_wb, sheet = "Correlation_pval", corr.res$P, rowNames = T)
      
      addWorksheet(hc_wb, 'Correlation_Summary')
      writeData(hc_wb, sheet = 'Correlation_Summary', report)
    }
  } else{
    # ---- BLOCK 2.2: Not all variables are numerical ----
    print_centered_note("NOT ALL VARIABLES ARE NUMERIC. APLYING MIXED CORRELATIONS: ")
    
    numeric_variables <- colnames(metadata)[sapply(metadata, is.numeric)]
    categorical_variables <- colnames(metadata)[sapply(metadata, is.factor) | sapply(metadata, is.character)]
    metadata <- metadata %>%
      mutate(across(all_of(categorical_variables), as.factor))
    
    categorical.df <- data.frame(
      `Cat Variable` = categorical_variables,
      `Levels` = sapply(metadata[categorical_variables], function(x) length(unique(x))),
      stringsAsFactors = FALSE
    )
    i <- 0
    for(cat.var in categorical_variables){
      
      cat.levels <- sapply(metadata[cat.var], nlevels)
      categorical.df[i+1, 1] = cat.var
      categorical.df[i+1, 2] = cat.levels
      i <- i+1
    }
    polytomous_variables <- colnames(metadata)[sapply(metadata, nlevels) > 2]
    dichotomous_variables <- colnames(metadata[sapply(metadata, nlevels) == 2])
    corr_method = 'spearman'
    
    # mixeCor function computes each type of correlation based on each type of variable automatically
    cor_df.tmp <- metadata
    
    cor_df.tmp[sapply(cor_df.tmp, is.factor)] <- as.data.frame(sapply(cor_df.tmp[sapply(cor_df.tmp, is.factor)], as.integer))
    corr.res <- mixedCor(data = cor_df.tmp, global = T,
                         c = which(colnames(cor_df.tmp) %in% numeric_variables), # Set the continuous variables
                         p = which(colnames(cor_df.tmp) %in% polytomous_variables), # Set the polytomous variables
                         d = which(colnames(cor_df.tmp) %in% dichotomous_variables), # Set the dichotomous variables
                         method = corr_method, # Set the Correlation Coefficient based on the coef previously selected
                         use = 'pairwise.complete.obs'
    )
    # Mixed - Corrplot 
    col<- colorRampPalette(c("blue", "white", "red"))
    
    png(file.path(output_dir, "01_Correlations","01_Mixed_Corrplot.png"), width = 3000, height = 3000, res = 300)
    plot_corr = corrplot(as.matrix(corr.res$rho), 
                         method="color", 
                         col=col(200),  
                         type="lower", 
                         order="hclust", ## "original", "AOE", "FPC", "hclust", "alphabet"
                         addCoef.col = "black", # Add coefficient of correlation
                         tl.col="black", tl.srt=45, tl.cex = 1, number.cex = 0.5,#Text label color and rotation
                         # Combine with significance
                         #p.mat = as.matrix(corr.res$P), sig.level = 0.05, insig = "blank", 
                         # hide correlation coefficient on the principal diagonal
                         diag=F ,
                         title = "Rho Mixed Correlation Coefficients are shown."
    )
    print(plot_corr)
    dev.off()
    
    # ---- Statistical Report
    corr.res = as.data.frame(corr.res$rho)
    ma_ln <- max(c(length(numeric_variables), length(dichotomous_variables), length(polytomous_variables))) # This object it's only to create unbalanced dataframe
    argument_report <- data.frame('Continuous Variables' = c(numeric_variables, rep(NA, ma_ln -length(numeric_variables))),
                                  'Dichotomous Variables' = c(dichotomous_variables, rep(NA, ma_ln-length(dichotomous_variables))),
                                  'Polytomous Variables'= c(polytomous_variables, rep(NA, ma_ln - length(polytomous_variables))))
    
    #wb = createWorkbook(title = paste('Mixed Correlation', Sys.Date(), sep=' '))
    addWorksheet(hc_wb, 'Mixed_Data_Report')
    writeData(hc_wb, sheet = 'Mixed_Data_Report', argument_report)
    
    addWorksheet(hc_wb, 'Correlation_Summary')
    writeData(hc_wb, sheet = 'Correlation_Summary', corr.res, rowNames = T)
    
    #addWorksheet(wb, 'Correlation Coefficients')
    #writeData(wb, sheet = 'Correlation Coefficients', corr.res, rowNames = T)
    saveWorkbook(wb, file = file.path(output_dir, "01_Correlations", "Mixed_Correlations_Report.xlsx"))
  }
  
  #---- BLOCK 3: Normalize data ---- 
  # Here, all numeric variables must be normalized previous clustering because this
  # method it's really sensitive to changes in units.
  print_centered_note("APLYING NORMALIZATION TO NUMERIC VARIABLES ")
  
  normalized_df <- metadata %>%
    mutate_if(is.numeric, scale2, na.rm = T)
  
  # NOTE: In this step, I scale the main dataset with all their variables, but for
  # the clustering, I use the distance matrix based on this normalized dataset.
  
  # Corr.res object is the result of the correlation to compute later the CCC
  # but it's not necessary for hierarchical clustering functions.
  
  print_centered_note("PERFORMING HIERARCHICAL CLUSTERING ")

  # First, the calculus of similarity matrix will be different if all data are numeric or
  # otherwise,the dataframe contains mixed type of data.
  
  # OBJECTIVE: Test all distance methods and select the best one based on CCC and 
  # elbow method.
  dir.create(path = file.path(output_dir, "02_HClustering"), showWarnings = F)

  if(sum(sapply(normalized_df, is.numeric)) == ncol(normalized_df)){
    # ---- BLOCK 4: Hierarchical Clustering for Numeric Data ----
    # ---- Block 4.1: HClust for only numerical dataset ----
    print("Optimising parameters for Hclust")
    # With qualitative variables, to compute the distance matrix:
    distance_measures <- c("euclidean", "maximum", "manhattan", "canberra")
    agglomeration_methods <- c("ward.D", "ward.D2", "single", "complete", 
                               "average", "mcquitty", "median","centroid")
    
    best_disimilarity_matrix <- NULL
    best_distance_measure <- NULL
    best_ccc <- 0
    best_agglomeration_method <- NULL
    
    for(distance_meth in distance_measures){
      for(agglo_meth in agglomeration_methods){
        disimilarity_matrix <- stats::dist(x = normalized_df, method = distance_meth)
        hcluster <- hclust(disimilarity_matrix, method = agglo_meth)
        mat.coef <- cophenetic(hcluster)
        ccc <-cor(mat.coef,
                  disimilarity_matrix,
                  method = 'pearson')
        
        if(ccc > best_ccc){
          best_disimilarity_matrix <- disimilarity_matrix
          best_ccc <- ccc
          best_distance_measure <- distance_meth
          best_agglomeration_method <- agglo_meth
        }
      }
    }
    hc_description <- rbind(best_distance_measure, best_agglomeration_method, best_ccc)
    addWorksheet(hc_wb, sheetName = 'HC_Parametrization')
    writeData(hc_wb, sheet = 'HC_Parametrization', hc_description, rowNames = T)
    
    # ---- Block 4.2: Determine the optimal number of clusters ----
    dir.create(file.path(output_dir, "02_HClustering", "Individual_Plots"), showWarnings = F)
    # Elbowplot
    print_centered_note(toupper("Determining the optimal number of clusters "))
    print('1.Elbowplot')
    elbow <- fviz_nbclust(normalized_df, FUN = hcut, 
                       diss = best_disimilarity_matrix, 
                       method = "wss", 
                       k.max = min(max_k_number, ncol(normalized_df))) + 
            xlab("Group Number") + 
            ylab("Intra-group variance") +
      labs(subtitle = "Elbow method", title = NULL)
    
    png(file.path(output_dir, "02_HClustering","Individual_Plots", "01_ElbowPlot.png"), 
        res = 300, height = 1500, width = 1500)
    print(elbow)
    dev.off()
    
    # Silhouet Width 
    print("2.Silhoutte Width")
    # Silhouette width is an internal validation metric which is an aggregated 
    # measure of how similar an observation is to its own cluster compared its 
    # closest neighboring cluster. It's range goes from -1 (poorly clustered) to
    # +1 (great clusterd)
    
    silhouette_plot <- fviz_nbclust(normalized_df, 
                                    FUN = hcut, 
                                    diss = best_disimilarity_matrix, 
                                    method = "silhouette", 
                                    k.max = min(max_k_number, ncol(normalized_df))) + 
      xlab("Group Number") + 
      ylab("Silhouette Width") +
      labs(subtitle = "Silhouette method", title = NULL)
    
    png(file.path(output_dir, "02_HClustering","Individual_Plots", "02_SilhouettePlot.png"),
        res = 300, width = 1500, height = 1500)
    print(silhouette_plot)
    dev.off()
    
    sil_width.df <- data.frame('K Clusters' = silhouette_plot$data$clusters,
                               'Silhouette With' = silhouette_plot$data$y)
    sil_width.df <- sil_width.df %>%
      arrange(desc(Silhouette.With))
    
    addWorksheet(hc_wb, 'Silhouette Width')
    writeDataTable(hc_wb, sheet = 'Silhouette Width', sil_width.df)
    
    # Gap Statistic
    print("3.Gap Statistic")
    set.seed(123)
    gap_stat <- fviz_nbclust(normalized_df, 
                 hcut, 
                 diss = best_disimilarity_matrix, 
                 k.max = min(max_k_number, nrow(normalized_df)-1),
                 method = "gap_stat", 
                 nboot = nboot)+
      labs(subtitle = "Gap statistic method", title = NULL)
    
    png(file.path(output_dir, "02_HClustering","Individual_Plots", "03_Gap_StatisticPlot.png"),
        res = 300, width = 1500, height = 1500)
    print(gap_stat)
    dev.off()
    
    # 30 indices
    majority <- NbClust(data = as.matrix(normalized_df), 
            diss = best_disimilarity_matrix, 
            method = best_agglomeration_method, 
            index = "all",
            distance = NULL,
            max.nc = min(max_k_number, ncol(normalized_df)-1),
            alphaBeale = 0.5)
    
    Nbindeices <- data.frame("Index" = colnames(majority$Best.nc),
                             "Value_Index" = majority$Best.nc[2,],
                             "Number_clusters" = majority$Best.nc[1,])
    
    addWorksheet(hc_wb, 'NbClust_Indices')
    writeDataTable(hc_wb, sheet = 'NbClust_Indices', x = Nbindeices)
    
    frequency_table <- as.data.frame(table(Nbindeices$Number_clusters))
    max_cluster <- as.numeric(as.character(frequency_table[which(frequency_table$Freq == max(frequency_table$Freq)), 1]))

    if(length(max_cluster) > 1){
      max_cluster <- ifelse(ties == "max", max(max_cluster), min(max_cluster))
    }
    addWorksheet(hc_wb, 'Frequency_NbClust_Indices')
    writeDataTable(hc_wb, sheet = 'Frequency_NbClust_Indices', frequency_table)
    
    # Frequency plot
    print("4.Frequency Optimal Number of Clusters")
    frequency_plot <- ggplot(Nbindeices, aes(x = factor(Number_clusters))) +
      geom_bar(fill = "#0073C2FF", color = "black") +  # Color azul de factoextra
      theme_minimal() +
      labs(
        subtitle  = "Majority of 30 Indices from Charrad (2014)",
        x = "Number of Clusters",
        y = "Frequency"
      ) +
      scale_y_continuous(breaks = seq(0, max(table(Nbindeices$Number_clusters)), by = 1)) +  # Escala de enteros
      theme(
        axis.text = element_text(size = 12),
        axis.title = element_text(size = 14, face = "bold")
        )
    
    png(file.path(output_dir, "02_HClustering","Individual_Plots","04_Frequency_Optimal_Clustering.png"),
        res = 300, width = 1500, height = 1500)
    print(frequency_plot)
    dev.off()
    
    # Summary plot Optimal Number of Clusters
    number_clusters_plot <- (elbow + silhouette_plot + gap_stat + frequency_plot) + 
      plot_annotation(title = 'Optimal Number of Clusters', tag_levels = "A", tag_prefix = "Fig. ")
    
    png(filename = file.path(output_dir, "02_HClustering", "01_Optimal_Clusters_Determination.png"),
        res = 300, width = 5000, height = 3000)
    print(number_clusters_plot)
    dev.off()
    
    # ---- Block 4.3: Selected Clusters Validation ----
    print_centered_note(toupper(paste0("Validating the selected number of clusters (", max_cluster, ") ")))
    # Si hay menos de 12 clusters, usar Set2 (colores más suaves)
    if (max_cluster <= 8) {
      colores <- brewer.pal(n = max_cluster, name = "Dark2")
      names(colores) <- c(1:max_cluster)
      
    } else {
      # Si hay más de 12 clusters, generar una paleta extendida con viridis
      colores <- viridis(max_cluster, option = "D")  # Puedes probar "D", "C", "B", etc.
      names(colores) <- c(1:max_cluster)
    }
    
    clust <- eclust(x = normalized_df, 
                    FUNcluster = "hclust", 
                    k = max_cluster, 
                    hc_metric = best_distance_measure, 
                    hc_method = best_agglomeration_method)
    # Disimilarity Matrix Heatmap
    print("1.Best Disimilarity Matrix")
    dismat_plot <- fviz_dist(best_disimilarity_matrix, 
             gradient = list(low = "#00AFBB",
                             mid = "white", 
                             high = "#FC4E07")) + 
      labs(title = "Best Disimilarity Matrix",
           subtitle = paste0("Distance Measure: ", best_distance_measure,
                             "\nAgglomerative Method: ", best_agglomeration_method,
                             "\nCCC: ", round(best_ccc, 2)))
    
    png(file.path(output_dir, "02_HClustering", "Individual_Plots","05_Best_Disimilarity_Matrix.png"), 
        res = 300, height = 3000, width = 3000)
    print(dismat_plot)
    dev.off()
    
    # Dendrogram
    print("2.Dendrogram")
    dendrogram <- fviz_dend(clust, 
                            show_labels = T,
                            as.ggplot = TRUE, 
                            k = max_cluster,
                            k_colors = colores) + 
      labs(title = paste0("Distance Measure: ", best_distance_measure, 
                          "\tAgglomerative Method: ", best_agglomeration_method))
    
    png(file.path(output_dir, "02_HClustering","Individual_Plots", "06_HClust_Dendrogram.png"), 
        res = 300, width = 3000, height = 3000)
    print(dendrogram)
    dev.off()
    
    # Silhouete Width Plot
    print("3.Silhouette")
    sil_plot <- fviz_silhouette(clust,
                                print.summary = F,
                                ggtheme = theme_minimal(),
                                palette = colores)
    png(filename = file.path(output_dir, "02_HClustering", "Individual_Plots", "07_Silhouette_Width_Plot.png"),
        res = 300, width = 1500, height = 1500)
    print(sil_plot)
    dev.off()
    
    silinfo <- clust$silinfo
    
    sil_wid_df <- silinfo$widths
    sil_wid_df <- sil_wid_df %>%
      group_by(cluster) %>%
      summarise(number = n(),
                ave.si.width = mean(sil_width)) %>%
      as.data.frame()
    
    addWorksheet(hc_wb, 'Sil_Width_Summary')
    writeDataTable(hc_wb, sheet = "Sil_Width_Summary", sil_wid_df, rowNames = T)
    
    addWorksheet(hc_wb, 'Sil_Width_Per_Sample')
    writeDataTable(hc_wb, sheet = "Sil_Width_Per_Sample", silinfo$widths, rowNames = T)
    
    # PCA
    print("4.PCA")
    pca <- fviz_cluster(clust, 
                 geom = "point", 
                 ellipse.type = "t",
                 palette = colores, 
                 ggtheme = theme_minimal())
    
    ggsave(filename = file.path(output_dir, "02_HClustering", "Individual_Plots", "08_PCA.png"),
           dpi = 300, width = 1500, height = 1500, units = "px", bg = "white")
    
    # Combined plot - Evaluation of Clustering
    print("5.CombinedPlot")
    combined <- (dismat_plot + dendrogram + sil_plot + pca) + 
      plot_annotation(title = 'Optimal Number of Clusters', tag_levels = "A", tag_prefix = "Fig. ")
    
    png(filename = file.path(output_dir, "02_HClustering", "02_Validation_Clustering_Plots.png"),
        res = 300, width = 6000, height = 4000)
    print(combined)
    dev.off()
    
    # ---- Block 4.4: External Validation ----
    if(!is.null(vect.external.val)){
      print_centered_note("ANALYZING EXTERNAL GROUP FOR VALIDATION")
      tabla <- table(vect.external.val, clust$cluster)
      tabla_df <- as.data.frame.matrix(tabla)
      tabla_df <- cbind(Categoria = rownames(tabla_df), tabla_df)
      
      addWorksheet(hc_wb, sheetName = "External Validation")
      writeData(hc_wb, "External Validation", tabla_df)
      
      clust_stats <- cluster.stats(d = stats::dist(x = normalized_df, method = best_distance_measure), 
                    as.numeric(as.factor(vect.external.val)), clust$cluster)
      
      external_validation_stats <- as.data.frame(t(data.frame("Corrected Rand index" = clust_stats$corrected.rand,
                 "Meila’s VI" = clust_stats$vi)))
      external_validation_stats$Interpretation <- c("From -1 (no agreement) to 1 (perfect agreement)", "Must be minimized")
      
      addWorksheet(hc_wb, sheetName = "External Val Metrics")
      writeDataTable(hc_wb, sheet = "External Val Metrics", x = external_validation_stats, rowNames = T)
    } # if for external validation group key
    
    print_centered_note("EXTRACTING CLUSTER ASSIGNMENTS ")
    clusters <- cutree(tree = clust, k = max_cluster)
    cluster_ass <- data.frame("ID" = names(clusters),
               "Cluster_Assignment" = clusters)
    
    addWorksheet(hc_wb, sheetName = "Cluster Assignment")
    writeDataTable(hc_wb, sheet = "Cluster Assignment", cluster_ass)
    
    
    print_centered_note(toupper("Saving Workbook"))
    saveWorkbook(hc_wb, file = file.path(output_dir, "HC_Report.xlsx"), overwrite = T)
  } else{
    # ---- Block 5: HClustering for mixed data ----
    print_centered_note("COMPUTING HCLUST FOR MIXED DATA")
    # ---- Calculating the best aglomerative method
    print("Computing the optimal agglomerative method")
    # With categorical variables: Gower distance - cluster::daisy
    # ref: https://www.r-bloggers.com/2016/06/clustering-mixed-data-types-in-r/
    distance_matrix <- daisy(x = normalized_df,
                             metric = 'gower')
    agglomeration_methods <- c("ward.D", "ward.D2", "single", "complete", 
                               "average", "mcquitty", "median","centroid")
    
    best_ccc <- 0
    best_agglomeration_method <- NULL
    
    for(agg_method in agglomeration_methods){
      hcluster <- hclust(distance_matrix, method = agg_method)
      mat.coef <- cophenetic(hcluster)
      ccc <-cor(mat.coef,
                distance_matrix,
                method = 'pearson')
      
      if(ccc > best_ccc){
        best_ccc <- ccc
        best_agglomeration_method <- agglo_meth
      }
    }# for loop for select the best agglomerative method

    hc_description <- rbind(best_agglomeration_method, best_ccc)
    addWorksheet(hc_wb, sheetName = 'HC_Parametrization')
    writeData(hc_wb, sheet = 'HC_Parametrization', hc_description, rowNames = T)

    # ---- Block 5.1: Determining the optimal number of clusters for mixed data ----
    dir.create(file.path(output_dir, "02_HClustering", "Individual_Plots"), showWarnings = F)
    # Elbowplot
    print_centered_note(toupper("Determining the optimal number of clusters "))
    print('1.Elbowplot')
    elbow <- fviz_nbclust(normalized_df, FUN = hcut, 
                          diss = distance_matrix, 
                          method = "wss", 
                          k.max = max_k_number) + 
      xlab("Group Number") + 
      ylab("Intra-group variance") +
      labs(subtitle = "Elbow method", title = NULL)
    
    png(file.path(output_dir, "02_HClustering","Individual_Plots", "01_ElbowPlot.png"), 
        res = 300, height = 1500, width = 1500)
    print(elbow)
    dev.off()
    
    # Silhouet Width 
    print("2.Silhoutte Width")
    # Silhouette width is an internal validation metric which is an aggregated 
    # measure of how similar an observation is to its own cluster compared its 
    # closest neighboring cluster. It's range goes from -1 (poorly clustered) to
    # +1 (great clusterd)
    
    silhouette_plot <- fviz_nbclust(normalized_df, FUN = hcut, diss = distance_matrix, 
                                    method = "silhouette", k.max = max_k_number) + 
      xlab("Group Number") + 
      ylab("Silhouette Width") +
      labs(subtitle = "Silhouette method", title = NULL)
    
    png(file.path(output_dir, "02_HClustering","Individual_Plots", "02_SilhouettePlot.png"),
        res = 300, width = 1500, height = 1500)
    print(silhouette_plot)
    dev.off()
    
    sil_width.df <- data.frame('K Clusters' = silhouette_plot$data$clusters,
                               'Silhouette With' = silhouette_plot$data$y)
    sil_width.df <- sil_width.df %>%
      arrange(desc(Silhouette.With))
    
    addWorksheet(hc_wb, 'Avg Sil Width Per Cluster')
    writeDataTable(hc_wb, sheet = 'Avg Sil Width Per Cluster', sil_width.df)
    
    # Gap Statistic
    print("3.Gap Statistic")
    set.seed(123)
    gap_stat <- fviz_nbclust(normalized_df, 
                             hcut, 
                             diss = distance_matrix, 
                             k.max = max_k_number,
                             method = "gap_stat", 
                             nboot = 500)+
      labs(subtitle = "Gap statistic method", title = NULL)
    
    png(file.path(output_dir, "02_HClustering","Individual_Plots", "03_Gap_StatisticPlot.png"),
        res = 300, width = 1500, height = 1500)
    print(gap_stat)
    dev.off()
    
    # 30 indices
    majority <- NbClust(data = as.matrix(normalized_df), 
                        diss = distance_matrix, 
                        method = best_agglomeration_method, 
                        index = "all",
                        distance = NULL,
                        max.nc <- nrow(normality.df)-1)
    
    Nbindeices <- data.frame("Index" = colnames(majority$Best.nc),
                             "Value_Index" = majority$Best.nc[2,],
                             "Number_clusters" = majority$Best.nc[1,])
    
    addWorksheet(hc_wb, 'NbClust_Indices')
    writeDataTable(hc_wb, sheet = 'NbClust_Indices Width', Nbindeices)
    
    frequency_table <- as.data.frame(table(Nbindeices$Number_clusters))
    max_cluster <- as.numeric(frequency_table$Var1[which.max(frequency_table$Freq)])
    
    addWorksheet(hc_wb, 'Frequency_NbClust_Indices')
    writeDataTable(hc_wb, sheet = 'Frequency_NbClust_Indices Width', frequency_table)
    
    # Frequency plot
    print("4.Frequency Optimal Number of Clusters")
    frequency_plot <- ggplot(Nbindeices, aes(x = factor(Number_clusters))) +
      geom_bar(fill = "#0073C2FF", color = "black") +  # Color azul de factoextra
      theme_minimal() +
      labs(
        subtitle  = "Majority of 30 Indices from Charrad (2014)",
        x = "Number of Clusters",
        y = "Frequency"
      ) +
      scale_y_continuous(breaks = seq(0, max(table(Nbindeices$Number_clusters)), by = 1)) +  # Escala de enteros
      theme(
        axis.text = element_text(size = 12),
        axis.title = element_text(size = 14, face = "bold")
      )
    
    png(file.path(output_dir, "02_HClustering","Individual_Plots","04_Frequency_Optimal_Clustering.png"),
        res = 300, width = 1500, height = 1500)
    print(frequency_plot)
    dev.off()
    
    # Summary plot Optimal Number of Clusters
    number_clusters_plot <- (elbow + silhouette_plot + gap_stat + frequency_plot) + 
      plot_annotation(title = 'Optimal Number of Clusters', tag_levels = "A", tag_prefix = "Fig. ")
    
    png(filename = file.path(output_dir, "02_HClustering", "01_Optimal_Clusters_Determination.png"),
        res = 300, width = 5000, height = 3000)
    print(number_clusters_plot)
    dev.off()
    
    # ---- Block 5.2: Selected Clusters Validation ----
    print_centered_note(toupper(paste0("Validating the selected number of clusters (", max_cluster, ") ")))
    # Si hay menos de 12 clusters, usar Set2 (colores más suaves)
    if (max_cluster <= 8) {
      colores <- brewer.pal(n = max_cluster, name = "Dark2")
      names(colores) <- c(1:max_cluster)
      
    } else {
      # Si hay más de 12 clusters, generar una paleta extendida con viridis
      colores <- viridis(max_cluster, option = "D")  # Puedes probar "D", "C", "B", etc.
      names(colores) <- c(1:max_cluster)
    }
    
    clust <- eclust(x = normalized_df, 
                    FUNcluster = "hclust", 
                    k = max_cluster, 
                    hc_metric = best_distance_measure, 
                    hc_method = best_agglomeration_method)
    # Disimilarity Matrix Heatmap
    print("1.Best Disimilarity Matrix")
    dismat_plot <- fviz_dist(distance_matrix, 
                             gradient = list(low = "#00AFBB",
                                             mid = "white", 
                                             high = "#FC4E07")) + 
      labs(title = "Best Disimilarity Matrix",
           subtitle = paste0("Distance Measure: Gower",
                             "\nAgglomerative Method: ", best_agglomeration_method,
                             "\nCCC: ", round(best_ccc, 2)))
    
    png(file.path(output_dir, "02_HClustering", "Individual_Plots","05_Best_Disimilarity_Matrix.png"), 
        res = 300, height = 3000, width = 3000)
    print(dismat_plot)
    dev.off()
    
    # Dendrogram
    print("2.Dendrogram")
    dendrogram <- fviz_dend(clust, 
                            show_labels = T,
                            as.ggplot = TRUE, 
                            k = max_cluster,
                            k_colors = colores) + 
      labs(title = paste0("Distance Measure: Gower", 
                          "\tAgglomerative Method: ", best_agglomeration_method))
    
    png(file.path(output_dir, "02_HClustering","Individual_Plots", "06_HClust_Dendrogram.png"), 
        res = 300, width = 3000, height = 3000)
    print(dendrogram)
    dev.off()
    
    # Silhouete Width Plot
    print("3.Silhouette")
    sil_plot <- fviz_silhouette(clust,
                                print.summary = F,
                                ggtheme = theme_minimal(),
                                palette = colores)
    png(filename = file.path(output_dir, "02_HClustering", "Individual_Plots", "07_Silhouette_Width_Plot.png"),
        res = 300, width = 1500, height = 1500)
    print(sil_plot)
    dev.off()
    
    silinfo <- clust$silinfo
    
    sil_wid_df <- silinfo$widths
    sil_wid_df <- sil_wid_df %>%
      group_by(cluster) %>%
      summarise(number = n(),
                ave.si.width = mean(sil_width)) %>%
      as.data.frame()
    
    addWorksheet(hc_wb, 'Sil_width_summary')
    writeDataTable(hc_wb, sheet = "Sil_width_summary", sil_wid_df, rowNames = T)
    
    addWorksheet(hc_wb, 'Sil_width_per_sample')
    writeDataTable(hc_wb, sheet = "Sil_width_per_sample", silinfo$widths, rowNames = T)
    
    # PCA
    print("4.PCA")
    pca <- fviz_cluster(clust, 
                        geom = "point", 
                        ellipse.type = "t",
                        palette = colores, 
                        ggtheme = theme_minimal())
    
    ggsave(filename = file.path(output_dir, "02_HClustering", "Individual_Plots", "08_PCA.png"),
           dpi = 300, width = 1500, height = 1500, units = "px", bg = "white")
    
    # Combined plot - Evaluation of Clustering
    print("5.CombinedPlot")
    combined <- (dismat_plot + dendrogram + sil_plot + pca) + 
      plot_annotation(title = 'Optimal Number of Clusters', tag_levels = "A", tag_prefix = "Fig. ")
    
    png(filename = file.path(output_dir, "02_HClustering", "02_Validation_Clustering_Plots.png"),
        res = 300, width = 6000, height = 4000)
    print(combined)
    dev.off()
    
    print("Saving Workbook")
    saveWorkbook(hc_wb, file = file.path(output_dir, "HC_Report.xlsx"), overwrite = T)
    
    #
  }# else statemente key
  
  print_centered_note("END OF THE FUNCTION")
  return(cluster_ass)
}# Function End Key