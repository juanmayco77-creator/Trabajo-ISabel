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

# =========================
# LIBRERÍAS
# =========================
library(ggplot2)
library(dplyr)
library(officer)

# =========================
# ASEGURAR FACTORES
# =========================
datos$Distrito <- factor(datos$Distrito,
                         levels = c("El Tambo", "Pilcomayo", "Huancayo"))

datos$Edad <- factor(datos$Edad)

# =========================
# CREAR GRÁFICOS
# =========================

g1 <- ggplot(datos, aes(x=Distrito, y=Costo)) +
  geom_boxplot(fill="lightblue") +
  theme_minimal() +
  labs(title="Boxplot del Costo por Distrito")

g2 <- ggplot(datos, aes(x=Edad, y=Costo)) +
  geom_boxplot(fill="lightgreen") +
  theme_minimal() +
  labs(title="Boxplot del Costo por Edad")

g3 <- ggplot(datos, aes(x=Distrito)) +
  geom_bar(fill="orange") +
  theme_minimal() +
  labs(title="Frecuencia por Distrito")

g4 <- ggplot(datos, aes(x=Edad)) +
  geom_bar(fill="purple") +
  theme_minimal() +
  labs(title="Frecuencia por Edad")

g5 <- ggplot(datos, aes(x=Costo)) +
  geom_histogram(bins=10, fill="steelblue", color="black") +
  theme_minimal() +
  labs(title="Histograma del Costo")

media_distrito <- datos %>%
  group_by(Distrito) %>%
  summarise(Media = mean(Costo, na.rm=TRUE))

g6 <- ggplot(media_distrito, aes(x=Distrito, y=Media)) +
  geom_col(fill="darkred") +
  theme_minimal() +
  labs(title="Media del Costo por Distrito")

media_edad <- datos %>%
  group_by(Edad) %>%
  summarise(Media = mean(Costo, na.rm=TRUE))

g7 <- ggplot(media_edad, aes(x=Edad, y=Media)) +
  geom_col(fill="darkgreen") +
  theme_minimal() +
  labs(title="Media del Costo por Edad")

# =========================
# GUARDAR GRÁFICOS COMO PNG
# =========================
ggsave("g1.png", g1, width=6, height=4)
ggsave("g2.png", g2, width=6, height=4)
ggsave("g3.png", g3, width=6, height=4)
ggsave("g4.png", g4, width=6, height=4)
ggsave("g5.png", g5, width=6, height=4)
ggsave("g6.png", g6, width=6, height=4)
ggsave("g7.png", g7, width=6, height=4)

# =========================
# EXPORTAR A WORD
# =========================
doc2 <- read_docx()

doc2 <- doc2 %>%
  body_add_par("GRÁFICOS ESTADÍSTICOS", style="heading 1") %>%
  
  body_add_par("Boxplot Costo por Distrito", style="heading 2") %>%
  body_add_img(src="g1.png", width=6, height=4) %>%
  
  body_add_par("Boxplot Costo por Edad", style="heading 2") %>%
  body_add_img(src="g2.png", width=6, height=4) %>%
  
  body_add_par("Frecuencia por Distrito", style="heading 2") %>%
  body_add_img(src="g3.png", width=6, height=4) %>%
  
  body_add_par("Frecuencia por Edad", style="heading 2") %>%
  body_add_img(src="g4.png", width=6, height=4) %>%
  
  body_add_par("Histograma del Costo", style="heading 2") %>%
  body_add_img(src="g5.png", width=6, height=4) %>%
  
  body_add_par("Media del Costo por Distrito", style="heading 2") %>%
  body_add_img(src="g6.png", width=6, height=4) %>%
  
  body_add_par("Media del Costo por Edad", style="heading 2") %>%
  body_add_img(src="g7.png", width=6, height=4)

print(doc2, target="Graficos_Analisis.docx")

#==tablas

# ===============================
# LIBRERÍAS
# ===============================
library(readxl)
library(dplyr)
library(ggplot2)
library(officer)
library(flextable)
library(rstatix)

# ===============================
# IMPORTAR DATOS
# ===============================
datos <- read_excel("ArchivoCambiado.xlsx")

datos$Distrito <- factor(datos$Distrito,
                         levels=c("El Tambo","Pilcomayo","Huancayo"))

datos$Edad <- factor(datos$Edad)

# ===============================
# PRUEBAS ESTADÍSTICAS
# ===============================

# Shapiro
shapiro <- shapiro.test(datos$Costo)
tabla_shapiro <- data.frame(
  Estadistico_W = shapiro$statistic,
  p_value = shapiro$p.value
)

# Bartlett
bartlett <- bartlett.test(Costo ~ Distrito, data=datos)
tabla_bartlett <- data.frame(
  Estadistico_K2 = bartlett$statistic,
  p_value = bartlett$p.value
)

# ANOVA
anova_model <- aov(Costo ~ Distrito, data=datos)
anova_tabla <- as.data.frame(summary(anova_model)[[1]])

# Kruskal
kruskal <- kruskal.test(Costo ~ Distrito, data=datos)
tabla_kruskal <- data.frame(
  Chi_cuadrado = kruskal$statistic,
  p_value = kruskal$p.value
)

# Mann-Whitney (comparaciones por pares)
mann_whitney <- datos %>%
  pairwise_wilcox_test(Costo ~ Distrito, p.adjust.method="bonferroni")

# ===============================
# INTERPRETACIONES
# ===============================

int_shapiro <- ifelse(shapiro$p.value<0.05,
                      "La variable Costo no presenta normalidad.",
                      "La variable Costo presenta normalidad.")

int_bartlett <- ifelse(bartlett$p.value<0.05,
                       "No existe homogeneidad de varianzas.",
                       "Existe homogeneidad de varianzas.")

p_anova <- anova_tabla$`Pr(>F)`[1]

int_anova <- ifelse(p_anova<0.05,
                    "Existen diferencias significativas entre medias.",
                    "No existen diferencias significativas entre medias.")

int_kruskal <- ifelse(kruskal$p.value<0.05,
                      "La prueba no paramétrica confirma diferencias.",
                      "No se detectan diferencias significativas.")

# ===============================
# CREAR DOCUMENTO
# ===============================
doc <- read_docx()

doc <- doc %>%
  body_add_par("INFORME ESTADÍSTICO FINAL", style="heading 1") %>%
  
  body_add_par("1. Prueba de Normalidad (Shapiro-Wilk)", style="heading 2") %>%
  body_add_flextable(flextable(tabla_shapiro)) %>%
  body_add_par(int_shapiro) %>%
  
  body_add_par("2. Prueba de Homogeneidad (Bartlett)", style="heading 2") %>%
  body_add_flextable(flextable(tabla_bartlett)) %>%
  body_add_par(int_bartlett) %>%
  
  body_add_par("3. ANOVA", style="heading 2") %>%
  body_add_flextable(flextable(anova_tabla)) %>%
  body_add_par(int_anova) %>%
  
  body_add_par("4. Kruskal-Wallis", style="heading 2") %>%
  body_add_flextable(flextable(tabla_kruskal)) %>%
  body_add_par(int_kruskal) %>%
  
  body_add_par("5. Comparaciones por pares (Mann–Whitney con Bonferroni)", style="heading 2") %>%
  body_add_flextable(flextable(mann_whitney))

print(doc, target="Informe_Final_Tesis.docx")

# ==========================================================
# CREAR DOCUMENTO INDEPENDIENTE:
# CONCLUSIONES Y DISCUSIÓN
# ==========================================================

doc_final <- read_docx()

# ==========================================================
# RESUMEN EJECUTIVO
# ==========================================================

resumen_ejecutivo <- paste(
  "El presente estudio evaluó diferencias en la variable Costo según Distrito.",
  "Se aplicaron análisis descriptivos e inferenciales.",
  "La prueba de normalidad indicó que", 
  ifelse(shapiro$p.value < 0.05,
         "la distribución no cumple el supuesto de normalidad.",
         "la distribución cumple el supuesto de normalidad."),
  "La prueba de homogeneidad de varianzas determinó que",
  ifelse(bartlett$p.value < 0.05,
         "no existe igualdad de varianzas.",
         "existe igualdad de varianzas."),
  "El análisis ANOVA mostró que",
  ifelse(p_anova < 0.05,
         "existen diferencias estadísticamente significativas entre distritos.",
         "no existen diferencias estadísticamente significativas entre distritos."),
  "La prueba no paramétrica de Kruskal-Wallis",
  ifelse(kruskal$p.value < 0.05,
         "confirmó la presencia de diferencias entre grupos.",
         "no evidenció diferencias significativas."),
  "En conjunto, los resultados permiten una interpretación robusta."
)

doc_final <- doc_final %>%
  body_add_par("RESUMEN EJECUTIVO", style="heading 1") %>%
  body_add_par(resumen_ejecutivo)



# ==========================================================
# CONCLUSIÓN GENERAL
# ==========================================================

conclusion_general <- paste(
  "A partir del análisis estadístico realizado, se concluye que:",
  ifelse(p_anova < 0.05 | kruskal$p.value < 0.05,
         "existen diferencias significativas en el Costo entre los distritos evaluados.",
         "no se encontraron diferencias estadísticamente significativas en el Costo entre distritos."),
  "El tamaño del efecto observado (Eta² =", round(eta2,4),
  "y Epsilon² =", round(epsilon2,4),
  ") permite interpretar la magnitud de dichas diferencias."
)

doc_final <- doc_final %>%
  body_add_par("CONCLUSIÓN GENERAL", style="heading 1") %>%
  body_add_par(conclusion_general)



# ==========================================================
# DISCUSIÓN ESTILO TESIS
# ==========================================================

discusion_texto <- paste(
  "Los resultados evidencian que el comportamiento del Costo varía según el distrito.",
  ifelse(shapiro$p.value < 0.05,
         "La ausencia de normalidad sugiere interpretar con cautela los resultados paramétricos.",
         "El cumplimiento del supuesto de normalidad fortalece la validez del ANOVA."),
  ifelse(p_anova < 0.05,
         "Las diferencias observadas podrían estar asociadas a factores estructurales propios de cada distrito.",
         "La homogeneidad observada podría indicar similitud en los patrones de comportamiento del Costo."),
  "Las comparaciones por pares mediante Mann–Whitney complementan el análisis al identificar diferencias específicas.",
  "El tamaño del efecto aporta evidencia sobre la magnitud práctica de los hallazgos.",
  "Se recomienda ampliar futuras investigaciones incorporando variables adicionales y mayor tamaño muestral."
)

doc_final <- doc_final %>%
  body_add_par("DISCUSIÓN", style="heading 1") %>%
  body_add_par(discusion_texto)



# ==========================================================
# GUARDAR DOCUMENTO INDEPENDIENTE
# ==========================================================

print(doc_final, target="Conclusiones_Discusion_Final.docx")

#================ PARTE II MEDIDAS DE TENDENCIA CENTRAL ================

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

# =========================
# LIBRERÍAS
# =========================
library(ggplot2)
library(dplyr)
library(officer)

# =========================
# ASEGURAR FACTORES
# =========================
datos$Distrito <- factor(datos$Distrito,
                         levels = c("El Tambo", "Pilcomayo", "Huancayo"))

datos$Edad <- factor(datos$Edad)

# =========================
# CREAR GRÁFICOS
# =========================

g1 <- ggplot(datos, aes(x=Distrito, y=Costo)) +
  geom_boxplot(fill="lightblue") +
  theme_minimal() +
  labs(title="Boxplot del Costo por Distrito")

g2 <- ggplot(datos, aes(x=Edad, y=Costo)) +
  geom_boxplot(fill="lightgreen") +
  theme_minimal() +
  labs(title="Boxplot del Costo por Edad")

g3 <- ggplot(datos, aes(x=Distrito)) +
  geom_bar(fill="orange") +
  theme_minimal() +
  labs(title="Frecuencia por Distrito")

g4 <- ggplot(datos, aes(x=Edad)) +
  geom_bar(fill="purple") +
  theme_minimal() +
  labs(title="Frecuencia por Edad")

g5 <- ggplot(datos, aes(x=Costo)) +
  geom_histogram(bins=10, fill="steelblue", color="black") +
  theme_minimal() +
  labs(title="Histograma del Costo")

media_distrito <- datos %>%
  group_by(Distrito) %>%
  summarise(Media = mean(Costo, na.rm=TRUE))

g6 <- ggplot(media_distrito, aes(x=Distrito, y=Media)) +
  geom_col(fill="darkred") +
  theme_minimal() +
  labs(title="Media del Costo por Distrito")

media_edad <- datos %>%
  group_by(Edad) %>%
  summarise(Media = mean(Costo, na.rm=TRUE))

g7 <- ggplot(media_edad, aes(x=Edad, y=Media)) +
  geom_col(fill="darkgreen") +
  theme_minimal() +
  labs(title="Media del Costo por Edad")

# =========================
# GUARDAR GRÁFICOS COMO PNG
# =========================
ggsave("g1.png", g1, width=6, height=4)
ggsave("g2.png", g2, width=6, height=4)
ggsave("g3.png", g3, width=6, height=4)
ggsave("g4.png", g4, width=6, height=4)
ggsave("g5.png", g5, width=6, height=4)
ggsave("g6.png", g6, width=6, height=4)
ggsave("g7.png", g7, width=6, height=4)

# =========================
# EXPORTAR A WORD
# =========================
doc2 <- read_docx()

doc2 <- doc2 %>%
  body_add_par("GRÁFICOS ESTADÍSTICOS", style="heading 1") %>%
  
  body_add_par("Boxplot Costo por Distrito", style="heading 2") %>%
  body_add_img(src="g1.png", width=6, height=4) %>%
  
  body_add_par("Boxplot Costo por Edad", style="heading 2") %>%
  body_add_img(src="g2.png", width=6, height=4) %>%
  
  body_add_par("Frecuencia por Distrito", style="heading 2") %>%
  body_add_img(src="g3.png", width=6, height=4) %>%
  
  body_add_par("Frecuencia por Edad", style="heading 2") %>%
  body_add_img(src="g4.png", width=6, height=4) %>%
  
  body_add_par("Histograma del Costo", style="heading 2") %>%
  body_add_img(src="g5.png", width=6, height=4) %>%
  
  body_add_par("Media del Costo por Distrito", style="heading 2") %>%
  body_add_img(src="g6.png", width=6, height=4) %>%
  
  body_add_par("Media del Costo por Edad", style="heading 2") %>%
  body_add_img(src="g7.png", width=6, height=4)

print(doc2, target="Graficos_Analisis.docx")

#==tablas

# ===============================
# LIBRERÍAS
# ===============================
library(readxl)
library(dplyr)
library(ggplot2)
library(officer)
library(flextable)
library(rstatix)

# ===============================
# IMPORTAR DATOS
# ===============================
datos <- read_excel("ArchivoCambiado.xlsx")

datos$Distrito <- factor(datos$Distrito,
                         levels=c("El Tambo","Pilcomayo","Huancayo"))

datos$Edad <- factor(datos$Edad)

# ===============================
# PRUEBAS ESTADÍSTICAS
# ===============================

# Shapiro
shapiro <- shapiro.test(datos$Costo)
tabla_shapiro <- data.frame(
  Estadistico_W = shapiro$statistic,
  p_value = shapiro$p.value
)

# Bartlett
bartlett <- bartlett.test(Costo ~ Distrito, data=datos)
tabla_bartlett <- data.frame(
  Estadistico_K2 = bartlett$statistic,
  p_value = bartlett$p.value
)

# ANOVA
anova_model <- aov(Costo ~ Distrito, data=datos)
anova_tabla <- as.data.frame(summary(anova_model)[[1]])

# Kruskal
kruskal <- kruskal.test(Costo ~ Distrito, data=datos)
tabla_kruskal <- data.frame(
  Chi_cuadrado = kruskal$statistic,
  p_value = kruskal$p.value
)

# Mann-Whitney (comparaciones por pares)
mann_whitney <- datos %>%
  pairwise_wilcox_test(Costo ~ Distrito, p.adjust.method="bonferroni")

# ===============================
# INTERPRETACIONES
# ===============================

int_shapiro <- ifelse(shapiro$p.value<0.05,
                      "La variable Costo no presenta normalidad.",
                      "La variable Costo presenta normalidad.")

int_bartlett <- ifelse(bartlett$p.value<0.05,
                       "No existe homogeneidad de varianzas.",
                       "Existe homogeneidad de varianzas.")

p_anova <- anova_tabla$`Pr(>F)`[1]

int_anova <- ifelse(p_anova<0.05,
                    "Existen diferencias significativas entre medias.",
                    "No existen diferencias significativas entre medias.")

int_kruskal <- ifelse(kruskal$p.value<0.05,
                      "La prueba no paramétrica confirma diferencias.",
                      "No se detectan diferencias significativas.")

# ===============================
# CREAR DOCUMENTO
# ===============================
doc <- read_docx()

doc <- doc %>%
  body_add_par("INFORME ESTADÍSTICO FINAL", style="heading 1") %>%
  
  body_add_par("1. Prueba de Normalidad (Shapiro-Wilk)", style="heading 2") %>%
  body_add_flextable(flextable(tabla_shapiro)) %>%
  body_add_par(int_shapiro) %>%
  
  body_add_par("2. Prueba de Homogeneidad (Bartlett)", style="heading 2") %>%
  body_add_flextable(flextable(tabla_bartlett)) %>%
  body_add_par(int_bartlett) %>%
  
  body_add_par("3. ANOVA", style="heading 2") %>%
  body_add_flextable(flextable(anova_tabla)) %>%
  body_add_par(int_anova) %>%
  
  body_add_par("4. Kruskal-Wallis", style="heading 2") %>%
  body_add_flextable(flextable(tabla_kruskal)) %>%
  body_add_par(int_kruskal) %>%
  
  body_add_par("5. Comparaciones por pares (Mann–Whitney con Bonferroni)", style="heading 2") %>%
  body_add_flextable(flextable(mann_whitney))

print(doc, target="Informe_Final_Tesis.docx")

# ==========================================================
# CREAR DOCUMENTO INDEPENDIENTE:
# CONCLUSIONES Y DISCUSIÓN
# ==========================================================

doc_final <- read_docx()

# ==========================================================
# RESUMEN EJECUTIVO
# ==========================================================

resumen_ejecutivo <- paste(
  "El presente estudio evaluó diferencias en la variable Costo según Distrito.",
  "Se aplicaron análisis descriptivos e inferenciales.",
  "La prueba de normalidad indicó que", 
  ifelse(shapiro$p.value < 0.05,
         "la distribución no cumple el supuesto de normalidad.",
         "la distribución cumple el supuesto de normalidad."),
  "La prueba de homogeneidad de varianzas determinó que",
  ifelse(bartlett$p.value < 0.05,
         "no existe igualdad de varianzas.",
         "existe igualdad de varianzas."),
  "El análisis ANOVA mostró que",
  ifelse(p_anova < 0.05,
         "existen diferencias estadísticamente significativas entre distritos.",
         "no existen diferencias estadísticamente significativas entre distritos."),
  "La prueba no paramétrica de Kruskal-Wallis",
  ifelse(kruskal$p.value < 0.05,
         "confirmó la presencia de diferencias entre grupos.",
         "no evidenció diferencias significativas."),
  "En conjunto, los resultados permiten una interpretación robusta."
)

doc_final <- doc_final %>%
  body_add_par("RESUMEN EJECUTIVO", style="heading 1") %>%
  body_add_par(resumen_ejecutivo)



# ==========================================================
# CONCLUSIÓN GENERAL
# ==========================================================

conclusion_general <- paste(
  "A partir del análisis estadístico realizado, se concluye que:",
  ifelse(p_anova < 0.05 | kruskal$p.value < 0.05,
         "existen diferencias significativas en el Costo entre los distritos evaluados.",
         "no se encontraron diferencias estadísticamente significativas en el Costo entre distritos."),
  "El tamaño del efecto observado (Eta² =", round(eta2,4),
  "y Epsilon² =", round(epsilon2,4),
  ") permite interpretar la magnitud de dichas diferencias."
)

doc_final <- doc_final %>%
  body_add_par("CONCLUSIÓN GENERAL", style="heading 1") %>%
  body_add_par(conclusion_general)



# ==========================================================
# DISCUSIÓN ESTILO TESIS
# ==========================================================

discusion_texto <- paste(
  "Los resultados evidencian que el comportamiento del Costo varía según el distrito.",
  ifelse(shapiro$p.value < 0.05,
         "La ausencia de normalidad sugiere interpretar con cautela los resultados paramétricos.",
         "El cumplimiento del supuesto de normalidad fortalece la validez del ANOVA."),
  ifelse(p_anova < 0.05,
         "Las diferencias observadas podrían estar asociadas a factores estructurales propios de cada distrito.",
         "La homogeneidad observada podría indicar similitud en los patrones de comportamiento del Costo."),
  "Las comparaciones por pares mediante Mann–Whitney complementan el análisis al identificar diferencias específicas.",
  "El tamaño del efecto aporta evidencia sobre la magnitud práctica de los hallazgos.",
  "Se recomienda ampliar futuras investigaciones incorporando variables adicionales y mayor tamaño muestral."
)

doc_final <- doc_final %>%
  body_add_par("DISCUSIÓN", style="heading 1") %>%
  body_add_par(discusion_texto)



# ==========================================================
# GUARDAR DOCUMENTO INDEPENDIENTE
# ==========================================================

print(doc_final, target="Conclusiones_Discusion_Final.docx")


#====================PARTE III CONTRASTE DE HIPOTESIS ==================
# =========================================
# 1. LIBRERÍAS
# =========================================
library(readxl)
library(dplyr)
library(broom)
library(officer)
library(flextable)

cat("✔ Librerías cargadas\n")

# =========================================
# 2. IMPORTAR ARCHIVO
# =========================================
data <- read_excel(file.choose())

cat("✔ Archivo cargado\n")
cat("Variables detectadas:\n")
print(names(data))

# =========================================
# 3. CONVERTIR TEXTO A FACTOR
# =========================================
data <- data %>% mutate(across(where(is.character), as.factor))

# =========================================
# 4. VERIFICAR VARIABLES NECESARIAS
# =========================================

edad <- grep("edad", names(data), ignore.case=TRUE, value=TRUE)[1]
distrito <- grep("distr", names(data), ignore.case=TRUE, value=TRUE)[1]
servicio <- grep("serv", names(data), ignore.case=TRUE, value=TRUE)[1]
costo <- grep("cost", names(data), ignore.case=TRUE, value=TRUE)[1]

binaria <- names(data)[sapply(data, function(x)
  length(unique(na.omit(x)))==2)][1]

cat("Edad:", edad,"\n")
cat("Distrito:", distrito,"\n")
cat("Servicio:", servicio,"\n")
cat("Costo:", costo,"\n")
cat("Variable binaria:", binaria,"\n")

# detener si falta algo
if(any(is.na(c(edad,distrito,servicio,costo,binaria)))){
  stop("❌ Faltan variables. Revisa los nombres mostrados arriba.")
}

cat("✔ Variables válidas\n")

# =========================================
# 5. MODELOS
# =========================================
modelo_costo <- lm(as.formula(paste(costo,"~",edad,"+",distrito,"+",servicio)), data=data)
modelo_uso <- glm(as.formula(paste(binaria,"~",edad,"+",distrito,"+",servicio)),
                  data=data, family=binomial)

res_costo <- tidy(modelo_costo)
res_uso <- tidy(modelo_uso)

cat("✔ Modelos estimados\n")

# =========================================
# 6. PRUEBAS NO PARAMÉTRICAS
# =========================================
prueba_B <- kruskal.test(as.formula(paste(costo,"~",edad)), data=data)
prueba_C <- kruskal.test(as.formula(paste(costo,"~",distrito)), data=data)

cat("✔ Pruebas realizadas\n")

# =========================================
# 7. INTERPRETACIONES AUTOMÁTICAS
# =========================================
interp_A <- if(any(res_costo$p.value < 0.05) | any(res_uso$p.value < 0.05)){
  "Existe evidencia de influencia de los factores sobre costo o uso."
} else {
  "No se encontró evidencia estadística de influencia."
}

interp_B <- if(prueba_B$p.value < 0.05){
  "Existe relación significativa entre edad y costo."
} else {
  "No existe relación significativa entre edad y costo."
}

interp_C <- if(prueba_C$p.value < 0.05){
  "Existe relación significativa entre distrito y costo."
} else {
  "No existe relación significativa entre distrito y costo."
}

# =========================================
# 8. EXPORTAR A WORD
# =========================================
doc <- read_docx()

doc <- body_add_par(doc, "CONTRASTE DE HIPÓTESIS", style="heading 1")

doc <- body_add_par(doc, "Hipótesis A", style="heading 2")
doc <- body_add_par(doc, interp_A)
doc <- body_add_flextable(doc, flextable(res_costo))
doc <- body_add_flextable(doc, flextable(res_uso))

doc <- body_add_par(doc, "Hipótesis B", style="heading 2")
doc <- body_add_par(doc, interp_B)
doc <- body_add_par(doc, paste("p =", round(prueba_B$p.value,4)))

doc <- body_add_par(doc, "Hipótesis C", style="heading 2")
doc <- body_add_par(doc, interp_C)
doc <- body_add_par(doc, paste("p =", round(prueba_C$p.value,4)))

ruta <- paste0(getwd(),"/Contraste_Hipotesis.docx")

print(doc, target=ruta)
