library(dplyr) 
library(readxl)
datos <- read_excel("my_data.xlsx")  # o read_excel() si es Excel
datos <- datos %>% drop_na()        # eliminar NA
datos <- datos %>% mutate_if(is.character, as.factor)  # convertir texto a factor

print(datos)