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
