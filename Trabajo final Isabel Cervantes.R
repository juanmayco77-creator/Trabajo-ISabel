#================ PARTE I ANALIZANDO OUTLIERS ================

# Cargar librerías necesarias

install.packages("readxl")
install.packages("dplyr")
install.packages("ggplot2")

library(readxl)
library(dplyr)
library(ggplot2)

# Cargar los datos desde el archivo Excel
file_path <- "D:/Isabel/Rstudios/Outlier.xlsx"  # Reemplazar con la ruta correcta
data <- read_excel(file_path)

# Vista general de los datos
head(data)
summary(data)

# Método de rango intercuartílico (IQR) para detectar valores atípicos
detect_outliers <- function(column) {
  Q1 <- quantile(column, 0.25)
  Q3 <- quantile(column, 0.75)
  IQR <- Q3 - Q1
  lower_limit <- Q1 - 1.5 * IQR
  upper_limit <- Q3 + 1.5 * IQR
  outliers <- column[column < lower_limit | column > upper_limit]
  return(outliers)
}

# Aplicar a las columnas
outliers_var1 <- detect_outliers(data$Costo)


print("Valores atípicos en Variable_1:")
print(outliers_var1)



# Visualización con boxplots
ggplot(data, aes(x = "", y = Costo)) +
  geom_boxplot() +
  labs(title = "Boxplot de Variable_1", y = "Variable_1") +
  theme_minimal()


# Manejo de valores atípicos: eliminación
data_clean <- data %>%
  filter(Costo >= quantile(Costo, 0.25) - 1.5 * IQR(Costo),
         Costo <= quantile(Costo, 0.75) + 1.5 * IQR(Costo))


# Reemplazar valores atípicos por la media o mediana
replace_outliers <- function(column, method = "mean") {
  Q1 <- quantile(column, 0.25)
  Q3 <- quantile(column, 0.75)
  IQR <- Q3 - Q1
  lower_limit <- Q1 - 1.5 * IQR
  upper_limit <- Q3 + 1.5 * IQR
  
  # Calcular reemplazo según el método
  if (method == "mean") {
    replacement <- mean(column, na.rm = TRUE)
  } else if (method == "median") {
    replacement <- median(column, na.rm = TRUE)
  } else {
    stop("Método no reconocido. Usa 'mean' o 'median'.")
  }
  
  # Reemplazar valores atípicos
  column[column < lower_limit | column > upper_limit] <- replacement
  return(column)
}

# Aplicar a las columnas con método deseado
data$Costo_media <- replace_outliers(data$Costo, method = "mean")

data
sumary(data)

#================ PARTE II ANALIZANDO OUTLIERS ================

#Instalar paquetes
install.packages(c("ggplot2","dplyr","officer","flextable","nortest"))

# =========================
# 1. CARGAR LIBRERÍAS
# =========================
library(readxl)
library(dplyr)
library(officer)
library(flextable)

# =========================
# 2. IMPORTAR DATOS
# =========================
datos <- read_excel("ArchivoCambiado.xlsx")

# =========================
# 3. CONVERTIR DISTRITO A FACTOR
# =========================
datos$Distrito <- factor(datos$Distrito,
                         levels = c("El Tambo", "Pilcomayo", "Huancayo"))

# =========================
# 4. TABLAS DE FRECUENCIA
# =========================

tabla_distrito <- as.data.frame(table(datos$Distrito))
tabla_edad <- as.data.frame(table(datos$Edad))

# Guardar tablas
write.csv(tabla_distrito, "Tabla_Frecuencia_Distrito.csv", row.names = FALSE)
write.csv(tabla_edad, "Tabla_Frecuencia_Edad.csv", row.names = FALSE)

# =========================
# 5. PRUEBA SHAPIRO-WILK (Costo)
# =========================
shapiro_test <- shapiro.test(datos$Costo)

interpretacion_shapiro <- ifelse(shapiro_test$p.value < 0.05,
                                 "Se rechaza H0: El Costo NO sigue distribución normal.",
                                 "No se rechaza H0: El Costo sigue distribución normal.")

# =========================
# 6. PRUEBA DE BARTLETT
# =========================
bartlett_test <- bartlett.test(Costo ~ Distrito, data = datos)

interpretacion_bartlett <- ifelse(bartlett_test$p.value < 0.05,
                                  "Se rechaza H0: No hay homogeneidad de varianzas.",
                                  "No se rechaza H0: Las varianzas son homogéneas.")

# =========================
# 7. PRUEBA KRUSKAL-WALLIS
# =========================
kruskal_test <- kruskal.test(Costo ~ Distrito, data = datos)

interpretacion_kruskal <- ifelse(kruskal_test$p.value < 0.05,
                                 "Se rechaza H0: Existen diferencias significativas entre distritos.",
                                 "No se rechaza H0: No existen diferencias significativas entre distritos.")

# =========================
# 8. ANOVA
# =========================
anova_model <- aov(Costo ~ Distrito, data = datos)
anova_result <- summary(anova_model)

p_anova <- anova_result[[1]]$`Pr(>F)`[1]

interpretacion_anova <- ifelse(p_anova < 0.05,
                               "Se rechaza H0: Existen diferencias significativas entre medias.",
                               "No se rechaza H0: No existen diferencias significativas entre medias.")

# =========================
# 9. EXPORTAR RESULTADOS A WORD
# =========================

doc <- read_docx()

doc <- doc %>%
  body_add_par("ANÁLISIS ESTADÍSTICO", style = "heading 1") %>%
  
  body_add_par("Tabla de Frecuencia - Distrito", style = "heading 2") %>%
  body_add_flextable(flextable(tabla_distrito)) %>%
  
  body_add_par("Tabla de Frecuencia - Edad", style = "heading 2") %>%
  body_add_flextable(flextable(tabla_edad)) %>%
  
  body_add_par("Prueba Shapiro-Wilk", style = "heading 2") %>%
  body_add_par(paste("W =", round(shapiro_test$statistic,4),
                     " | p-value =", round(shapiro_test$p.value,5))) %>%
  body_add_par(interpretacion_shapiro) %>%
  
  body_add_par("Prueba Bartlett", style = "heading 2") %>%
  body_add_par(paste("K-squared =", round(bartlett_test$statistic,4),
                     " | p-value =", round(bartlett_test$p.value,5))) %>%
  body_add_par(interpretacion_bartlett) %>%
  
  body_add_par("Prueba Kruskal-Wallis", style = "heading 2") %>%
  body_add_par(paste("Chi-squared =", round(kruskal_test$statistic,4),
                     " | p-value =", round(kruskal_test$p.value,5))) %>%
  body_add_par(interpretacion_kruskal) %>%
  
  body_add_par("ANOVA", style = "heading 2") %>%
  body_add_par(paste("p-value =", round(p_anova,5))) %>%
  body_add_par(interpretacion_anova)

print(doc, target = "Resultados_Analisis.docx")