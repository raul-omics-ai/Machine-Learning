########## 01/04/2025 ##########

# Función para crear los directorios y añadirlos de formac onsecutiva para tener una jerarquía
# de directorios ordenada.

create_sequential_dir <- function(path, name, width = 2, start = 1, zero_pad = TRUE) {
  # 1) Validaciones iniciales
  if (!dir.exists(path)) {
    stop("La ruta especificada no existe: ", path)
  }
  if (!nzchar(name)) {
    stop("El parámetro 'name' no puede estar vacío.")
  }
  if (!is.numeric(width) || width < 1) {
    stop("'width' debe ser un entero >= 1.")
  }
  if (!is.numeric(start) || start < 0) {
    stop("'start' debe ser un entero >= 0.")
  }
  
  # 2) Listado de directorios (solo los de primer nivel)
  existing_dirs <- list.dirs(path, full.names = FALSE, recursive = FALSE)
  
  # 3) Filtrar por patrón: ^number(_rest)
  #    Ejemplos válidos: "01_Proyecto", "2_Datos", "123_ML"
  pattern <- "^([0-9]+)_.+"
  idx <- grepl(pattern, existing_dirs, perl = TRUE)
  
  # 4) Extraer números solo de los válidos
  numbers <- integer(0)
  if (any(idx)) {
    numbers <- as.integer(sub(pattern, "\\1", existing_dirs[idx], perl = TRUE))
  }
  
  # 5) Calcular el siguiente número disponible con manejo del caso vacío
  if (length(numbers) == 0) {
    next_num <- as.integer(start)
  } else {
    next_num <- max(numbers, na.rm = TRUE) + 1L
  }
  
  # 6) Formatear el nombre de la carpeta
  prefix <- if (isTRUE(zero_pad)) sprintf(paste0("%0", width, "d"), next_num) else as.character(next_num)
  dir_name <- paste0(prefix, "_", name)
  
  # 7) Ruta final y creación (con protección ante colisión)
  new_dir <- file.path(path, dir_name)
  if (dir.exists(new_dir)) {
    stop("Ya existe el directorio: ", new_dir)
  }
  
  dir.create(new_dir, recursive = FALSE, showWarnings = FALSE)
  if (!dir.exists(new_dir)) {
    stop("No se pudo crear el directorio: ", new_dir)
  }
  
  message("Directorio creado: ", new_dir)
  return(new_dir)
}
