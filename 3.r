library(tidyr)
library(dplyr)
library(tidyverse)
library(readxl)
library(GGally)
library(psych)

datos <- read_excel("my_data.xlsx")
convert_to_smart_factor <- function(df, threshold = 15) {
  df %>%
    mutate(across(where(is.character), 
                  ~ {
                    unique_vals <- unique(.)
                    if(length(unique_vals) <= threshold && 
                       length(unique_vals) > 1) {
                      as.factor(.)
                    } else {
                      .
                    }
                  }))
}

datos <- convert_to_smart_factor(datos)


print(ggplot(datos, aes(x = "Región"
)) + 
  geom_bar(data = data.frame(x="Región", y=50),
         aes(x = x, y = y),
         stat = "identity"))

