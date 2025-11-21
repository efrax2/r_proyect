library(dplyr)
library(tidyverse)
library(readxl)

datos <- read_excel("my_data.xlsx")
datos <- datos %>% drop_na()
datos <- datos %>% mutate(across(where(is.character), as.factor))

# Ver estructura
str(datos)

# Gráfico de barras para variables categóricas
# Reemplaza 'variable_interes' con el nombre real de tu columna
print(ggplot(datos, aes(x = nombre_de_tu_variable)) + 
  geom_bar() +
  theme_minimal())

