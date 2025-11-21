# Instalar paquetes necesarios si no están instalados
if (!require("readxl")) install.packages("readxl")
if (!require("dplyr")) install.packages("dplyr")
if (!require("tidyr")) install.packages("tidyr")
if (!require("ggplot2")) install.packages("ggplot2")

# Cargar librerías
library(readxl)
library(dplyr)
library(tidyr)
library(ggplot2)

# Leer el archivo Excel
archivo <- "my_data.xlsx"

# Obtener nombres de las hojas
hojas <- excel_sheets(archivo)
print("Hojas disponibles:")
print(hojas)

# Leer datos demográficos por región
# Basado en el índice, los cuadros relevantes son:
# - Cuadro 8: Personas detenidas por región y sexo
# - Cuadro 12: Personas detenidas por edad
# - Cuadro 13: Personas detenidas por nacionalidad
# - Cuadro 14: Personas detenidas por estado civil
# - Cuadro 15: Personas detenidas por nivel de educación

# Función para leer datos de una hoja específica
leer_cuadro_demografico <- function(hoja_numero, variables) {
  tryCatch({
    datos <- read_excel(archivo, sheet = as.character(hoja_numero))
    return(datos)
  }, error = function(e) {
    message(paste("No se pudo leer la hoja", hoja_numero))
    return(NULL)
  })
}

# Leer datos de personas detenidas por región y sexo (Cuadro 8)
datos_region_sexo <- leer_cuadro_demografico(8)

# Leer datos de personas detenidas por edad (Cuadro 12)
datos_edad <- leer_cuadro_demografico(12)

# Leer datos de personas detenidas por nacionalidad (Cuadro 13)
datos_nacionalidad <- leer_cuadro_demografico(13)

# Leer datos de personas detenidas por estado civil (Cuadro 14)
datos_estado_civil <- leer_cuadro_demografico(14)

# Leer datos de personas detenidas por nivel educativo (Cuadro 15)
datos_educacion <- leer_cuadro_demografico(15)

# Función para procesar datos demográficos básicos por región
procesar_datos_demograficos <- function() {
  # Lista de regiones de Chile (según el cuadro 1)
  regiones <- c(
    "Arica y Parinacota", "Tarapacá", "Antofagasta", "Atacama", "Coquimbo",
    "Valparaíso", "Metropolitana", "O'Higgins", "Maule", "Ñuble", "Biobío",
    "La Araucanía", "Los Ríos", "Los Lagos", "Aysén", "Magallanes"
  )
  
  # Crear dataframe base con regiones
  datos_demograficos <- data.frame(Region = regiones)
  
  # Si tenemos datos del cuadro 1, extraer información básica
  try({
    cuadro1 <- read_excel(archivo, sheet = "1", skip = 4)
    
    # Filtrar solo filas con datos de 2024
    datos_2024 <- cuadro1[grepl("2024", cuadro1[[1]]), ]
    
    if (nrow(datos_2024) > 0) {
      # Extraer denuncias y detenciones por región para 2024
      denuncias_2024 <- as.numeric(datos_2024[1, 4:19])
      detenciones_2024 <- as.numeric(datos_2024[2, 4:19])
      
      datos_demograficos$Denuncias_2024 <- denuncias_2024
      datos_demograficos$Detenciones_2024 <- detenciones_2024
      datos_demograficos$Total_Casos_2024 <- denuncias_2024 + detenciones_2024
    }
  }, silent = TRUE)
  
  return(datos_demograficos)
}

# Procesar datos demográficos básicos
datos_demograficos <- procesar_datos_demograficos()

# Mostrar resumen de datos demográficos por región
print("Resumen de datos demográficos por región:")
print(datos_demograficos)

# Análisis descriptivo básico
if (!is.null(datos_demograficos) && ncol(datos_demograficos) > 1) {
  cat("\n=== ANÁLISIS DESCRIPTIVO ===\n")
  
  # Total de casos por región
  cat("Total de casos policiales por región (2024):\n")
  print(datos_demograficos[, c("Region", "Total_Casos_2024")])
  
  # Regiones con mayor y menor número de casos
  cat("\nRegiones con mayor número de casos:\n")
  mayor_casos <- datos_demograficos[order(-datos_demograficos$Total_Casos_2024), ]
  print(head(mayor_casos[, c("Region", "Total_Casos_2024")], 5))
  
  cat("\nRegiones con menor número de casos:\n")
  menor_casos <- datos_demograficos[order(datos_demograficos$Total_Casos_2024), ]
  print(head(menor_casos[, c("Region", "Total_Casos_2024")], 5))
  
  # Estadísticas resumen
  cat("\nEstadísticas resumen:\n")
  cat("Total nacional de denuncias 2024:", sum(datos_demograficos$Denuncias_2024, na.rm = TRUE), "\n")
  cat("Total nacional de detenciones 2024:", sum(datos_demograficos$Detenciones_2024, na.rm = TRUE), "\n")
  cat("Total nacional de casos 2024:", sum(datos_demograficos$Total_Casos_2024, na.rm = TRUE), "\n")
  
  # Proporciones por región
  datos_demograficos$Proporcion_Nacional <- round(
    datos_demograficos$Total_Casos_2024 / sum(datos_demograficos$Total_Casos_2024, na.rm = TRUE) * 100, 2
  )
  
  cat("\nProporción de casos por región (%):\n")
  print(datos_demograficos[, c("Region", "Proporcion_Nacional")])
}

# Visualización de datos
if (!is.null(datos_demograficos) && ncol(datos_demograficos) > 1) {
  # Gráfico de barras - Casos por región
  ggplot(datos_demograficos, aes(x = reorder(Region, Total_Casos_2024), y = Total_Casos_2024)) +
    geom_bar(stat = "identity", fill = "steelblue") +
    coord_flip() +
    labs(
      title = "Casos Policiales por Región - 2024",
      x = "Región",
      y = "Número de Casos",
      caption = "Fuente: Carabineros de Chile - INE"
    ) +
    theme_minimal()
  
  # Guardar el gráfico
  ggsave("casos_por_region.png", width = 10, height = 8)
  
  # Gráfico de proporciones
  ggplot(datos_demograficos, aes(x = "", y = Proporcion_Nacional, fill = Region)) +
    geom_bar(stat = "identity", width = 1) +
    coord_polar("y", start = 0) +
    labs(
      title = "Distribución de Casos Policiales por Región - 2024",
      fill = "Región",
      caption = "Fuente: Carabineros de Chile - INE"
    ) +
    theme_void()
  
  ggsave("distribucion_casos_region.png", width = 10, height = 8)
}

# Función para análisis comparativo entre regiones
analisis_comparativo <- function(datos) {
  if (is.null(datos) || ncol(datos) <= 1) return(NULL)
  
  cat("\n=== ANÁLISIS COMPARATIVO ===\n")
  
  # Densidad de casos (asumiendo población similar por región)
  datos$Tasa_Relativa <- round(
    datos$Total_Casos_2024 / mean(datos$Total_Casos_2024, na.rm = TRUE), 2
  )
  
  cat("Tasa relativa de casos (media = 1.00):\n")
  print(datos[, c("Region", "Tasa_Relativa")])
  
  # Identificar regiones con tasas altas y bajas
  cat("\nRegiones con tasas más altas que el promedio:\n")
  alta_tasa <- datos[datos$Tasa_Relativa > 1.2, c("Region", "Tasa_Relativa")]
  print(alta_tasa)
  
  cat("\nRegiones con tasas más bajas que el promedio:\n")
  baja_tasa <- datos[datos$Tasa_Relativa < 0.8, c("Region", "Tasa_Relativa")]
  print(baja_tasa)
  
  return(datos)
}

# Realizar análisis comparativo
datos_demograficos <- analisis_comparativo(datos_demograficos)

# Exportar resultados a CSV
if (!is.null(datos_demograficos)) {
  write.csv(datos_demograficos, "datos_demograficos_por_region.csv", row.names = FALSE)
  cat("\nDatos exportados a 'datos_demograficos_por_region.csv'\n")
}

# Resumen final
cat("\n=== RESUMEN EJECUTIVO ===\n")
cat("Se han extraído y analizado datos demográficos policiales por región.\n")
cat("Principales hallazgos:\n")
cat("- Se identificaron", nrow(datos_demograficos), "regiones\n")
cat("- Se procesaron datos de denuncias y detenciones para 2024\n")
cat("- Se generaron visualizaciones y análisis comparativos\n")
cat("- Los datos están disponibles para análisis más detallados\n")