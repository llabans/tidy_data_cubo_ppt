Poblacion \| departamento
================
LL

# FORMATO: ESTRICTURA DE CUBO

## TABLA SIN EDAD/CON EDAD

``` r
# CONDICIONES DE GENERACION DE TABLAS CUBO:  
# cubo filtrado por departamento: 
# 1. TIPO ATENCION: ALL
# 2. AÑO:2023
# 3. TIPO DX: D
# 4. AMBITO: ALL
# 5. DEPARTAMENTO: LIMA

# Vars-Primera tabla cubo sin edad y con provincia, distrito condiciÓn: objeto data

# Vars-segunda tabla cubo con edad y con provincia, distrito condiciÓn: objeto data_conedad4_clean
```

# DESCRIPTIVE LIMA BY DISTRICT data_clean \| data_conedad4_clean

# TIDY: procesamiento de tabla cubo \| data \| data_conedad4_clean

``` r
#calculo %: denominador x cada grupo de distrito-provincia
#estimacion de % por distrito-provincia, usando denominador total de cada distrito-provincia

#data_clean
data <- read.xlsx("raw/data_distlima_cond.xlsx")

data_ <- data %>%
    group_by(provincia_rh, distrito_rh) %>%
    mutate(total_distrito = sum(total)) %>%
    ungroup() %>%
    group_by(provincia_rh, distrito_rh, categoria) %>%
    mutate(porcentaje = round((total / total_distrito) * 100, 2)) %>%
    ungroup() %>%
    mutate(total = format(total, big.mark = ","))
write.xlsx(data_, "raw/porcentajes.xlsx")


#data_conedad4_clean
data_conedad4_clean <- read.xlsx("raw/data_distlimacond_edad.xlsx")
glimpse(data_conedad4_clean)
```

    ## Rows: 974,922
    ## Columns: 6
    ## $ provincia_rh <chr> "BARRANCA", "BARRANCA", "BARRANCA", "BARRANCA", "BARRANCA…
    ## $ distrito_rh  <chr> "BARRANCA", "BARRANCA", "BARRANCA", "BARRANCA", "BARRANCA…
    ## $ categoria    <chr> "A01 - FIEBRES TIFOIDEA Y PARATIFOIDEA", "A01 - FIEBRES T…
    ## $ edad         <dbl> 3, 4, 5, 8, 11, 16, 17, 18, 20, 23, 24, 31, 35, 39, 44, 4…
    ## $ total        <dbl> 1, 1, 1, 1, 4, 1, 1, 3, 1, 2, 1, 1, 1, 1, 1, 1, 1, 1, 2, …
    ## $ etapa_vida   <chr> "Niños", "Niños", "Niños", "Niños", "Niños", "Adolescente…

# PORCENTAJES

``` r
#crear variable ranking, a nivel de cada distrito-provincia
data_rank <- data_ %>%
    group_by(provincia_rh, distrito_rh) %>%
    mutate(ranking = dense_rank(desc(porcentaje))) %>%
    ungroup()
view(data_rank)
# ordenar de mayor a menor de acuerdo al ranking
data_order <- data_rank %>%
    arrange(provincia_rh, distrito_rh, ranking)

view(data_order)

write.xlsx(data_order, "raw/porcentajes_distritos_general_order.xlsx")

# 10 primeros de cada distrito
data_top10 <- data_order %>%
    filter(ranking <= 10)
view(data_top10)
glimpse(data_top10)
```

    ## Rows: 1,803
    ## Columns: 7
    ## $ provincia_rh   <chr> "BARRANCA", "BARRANCA", "BARRANCA", "BARRANCA", "BARRAN…
    ## $ distrito_rh    <chr> "BARRANCA", "BARRANCA", "BARRANCA", "BARRANCA", "BARRAN…
    ## $ categoria      <chr> "J02 - FARINGITIS AGUDA", "J00 - RINOFARINGITIS AGUDA […
    ## $ total          <chr> " 7,170", " 6,345", " 5,862", " 5,088", " 4,794", " 3,2…
    ## $ total_distrito <dbl> 104979, 104979, 104979, 104979, 104979, 104979, 104979,…
    ## $ porcentaje     <dbl> 6.83, 6.04, 5.58, 4.85, 4.57, 3.07, 2.59, 2.12, 1.69, 1…
    ## $ ranking        <int> 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 1, 2, 3, 4, 5, 6, 7, 8, …

``` r
# Data_conedad4_clean
# Agrupar y resumir datos por distrito, provincia y etapa de vida
data_resumen <- data_conedad4_clean %>%
    group_by(provincia_rh, distrito_rh, etapa_vida, categoria) %>%
    summarise(n = sum(total, na.rm = TRUE)) %>%
    ungroup()
```

    ## `summarise()` has grouped output by 'provincia_rh', 'distrito_rh',
    ## 'etapa_vida'. You can override using the `.groups` argument.

``` r
# Calcular los totales por distrito, provincia y etapa de vida
totales <- data_conedad4_clean %>%
    group_by(provincia_rh, distrito_rh, etapa_vida) %>%
    summarise(Total_individuos = sum(total, na.rm = TRUE)) %>%
    ungroup()
```

    ## `summarise()` has grouped output by 'provincia_rh', 'distrito_rh'. You can
    ## override using the `.groups` argument.

``` r
# Unir los datos resumidos con los totales de individuos
datos_porcentaje <- data_resumen %>%
    inner_join(totales, by = c("provincia_rh", "distrito_rh", "etapa_vida")) %>%
    mutate(Porcentaje = round((n / Total_individuos) * 100,1)) %>%
    arrange(provincia_rh, distrito_rh, etapa_vida, desc(Porcentaje))

# Ver la estructura de los datos resumidos con porcentajes
glimpse(datos_porcentaje)
```

    ## Rows: 188,514
    ## Columns: 7
    ## $ provincia_rh     <chr> "BARRANCA", "BARRANCA", "BARRANCA", "BARRANCA", "BARR…
    ## $ distrito_rh      <chr> "BARRANCA", "BARRANCA", "BARRANCA", "BARRANCA", "BARR…
    ## $ etapa_vida       <chr> "Adolescente", "Adolescente", "Adolescente", "Adolesc…
    ## $ categoria        <chr> "E66 - OBESIDAD", "K02 - CARIES DENTAL", "K05 - GINGI…
    ## $ n                <dbl> 1253, 1077, 690, 466, 358, 187, 185, 158, 137, 134, 1…
    ## $ Total_individuos <dbl> 8778, 8778, 8778, 8778, 8778, 8778, 8778, 8778, 8778,…
    ## $ Porcentaje       <dbl> 14.3, 12.3, 7.9, 5.3, 4.1, 2.1, 2.1, 1.8, 1.6, 1.5, 1…

``` r
view(datos_porcentaje) # con edad

# Verificar que sumatoria % x cada grupo de edad en cada distrito
datos_verificados <- datos_porcentaje %>%
    group_by(provincia_rh, distrito_rh, etapa_vida) %>%
    summarise(Total_porcentaje = sum(Porcentaje))
```

    ## `summarise()` has grouped output by 'provincia_rh', 'distrito_rh'. You can
    ## override using the `.groups` argument.

``` r
print(datos_verificados) #casi 100%
```

    ## # A tibble: 570 × 4
    ## # Groups:   provincia_rh, distrito_rh [114]
    ##    provincia_rh distrito_rh etapa_vida   Total_porcentaje
    ##    <chr>        <chr>       <chr>                   <dbl>
    ##  1 BARRANCA     BARRANCA    Adolescente              95.4
    ##  2 BARRANCA     BARRANCA    Adulto                   94.3
    ##  3 BARRANCA     BARRANCA    Adulto mayor             94.9
    ##  4 BARRANCA     BARRANCA    Joven                    95.3
    ##  5 BARRANCA     BARRANCA    Niños                    96  
    ##  6 BARRANCA     PARAMONGA   Adolescente              97.1
    ##  7 BARRANCA     PARAMONGA   Adulto                   95.5
    ##  8 BARRANCA     PARAMONGA   Adulto mayor             96.6
    ##  9 BARRANCA     PARAMONGA   Joven                    97  
    ## 10 BARRANCA     PARAMONGA   Niños                    95.8
    ## # ℹ 560 more rows

``` r
######################################
# Crear la variable de ranking usando dense_rank
datos_rankeados <- datos_porcentaje %>%
    group_by(provincia_rh, distrito_rh, etapa_vida) %>%
    mutate(Ranking = dense_rank(-Porcentaje)) %>%
    ungroup()

# Ver la estructura de los datos con el ranking
glimpse(datos_rankeados)
```

    ## Rows: 188,514
    ## Columns: 8
    ## $ provincia_rh     <chr> "BARRANCA", "BARRANCA", "BARRANCA", "BARRANCA", "BARR…
    ## $ distrito_rh      <chr> "BARRANCA", "BARRANCA", "BARRANCA", "BARRANCA", "BARR…
    ## $ etapa_vida       <chr> "Adolescente", "Adolescente", "Adolescente", "Adolesc…
    ## $ categoria        <chr> "E66 - OBESIDAD", "K02 - CARIES DENTAL", "K05 - GINGI…
    ## $ n                <dbl> 1253, 1077, 690, 466, 358, 187, 185, 158, 137, 134, 1…
    ## $ Total_individuos <dbl> 8778, 8778, 8778, 8778, 8778, 8778, 8778, 8778, 8778,…
    ## $ Porcentaje       <dbl> 14.3, 12.3, 7.9, 5.3, 4.1, 2.1, 2.1, 1.8, 1.6, 1.5, 1…
    ## $ Ranking          <int> 1, 2, 3, 4, 5, 6, 6, 7, 8, 9, 10, 10, 11, 11, 12, 12,…

``` r
view(datos_rankeados)

# Verificar que sumatoria % x cada grupo de edad en cada distrito
datos_verificados <- datos_rankeados %>%
    group_by(provincia_rh, distrito_rh, etapa_vida) %>%
    summarise(Total_porcentaje = sum(Porcentaje))
```

    ## `summarise()` has grouped output by 'provincia_rh', 'distrito_rh'. You can
    ## override using the `.groups` argument.

``` r
#filtrar los 10 primeros de cada distrito en cada edad
datosedad_top10 <- datos_rankeados %>%
    filter(Ranking <= 10)
    View(datosedad_top10)

glimpse(datosedad_top10)
```

    ## Rows: 8,952
    ## Columns: 8
    ## $ provincia_rh     <chr> "BARRANCA", "BARRANCA", "BARRANCA", "BARRANCA", "BARR…
    ## $ distrito_rh      <chr> "BARRANCA", "BARRANCA", "BARRANCA", "BARRANCA", "BARR…
    ## $ etapa_vida       <chr> "Adolescente", "Adolescente", "Adolescente", "Adolesc…
    ## $ categoria        <chr> "E66 - OBESIDAD", "K02 - CARIES DENTAL", "K05 - GINGI…
    ## $ n                <dbl> 1253, 1077, 690, 466, 358, 187, 185, 158, 137, 134, 1…
    ## $ Total_individuos <dbl> 8778, 8778, 8778, 8778, 8778, 8778, 8778, 8778, 8778,…
    ## $ Porcentaje       <dbl> 14.3, 12.3, 7.9, 5.3, 4.1, 2.1, 2.1, 1.8, 1.6, 1.5, 1…
    ## $ Ranking          <int> 1, 2, 3, 4, 5, 6, 6, 7, 8, 9, 10, 10, 1, 2, 3, 4, 5, …

### Resumen Nas data_clean

``` r
missing_data_top10 <- data_top10 %>%
  summarise_all(~ sum(is.na(.))) %>%
  pivot_longer(everything(), names_to = "variable", values_to = "n")

missing_datos_porcentaje <- datos_porcentaje %>%
    summarise_all(~ sum(is.na(.))) %>%
    pivot_longer(everything(), names_to = "variable", values_to = "n")
missing_datos_porcentaje #0 missing
```

    ## # A tibble: 7 × 2
    ##   variable             n
    ##   <chr>            <int>
    ## 1 provincia_rh         0
    ## 2 distrito_rh          0
    ## 3 etapa_vida           0
    ## 4 categoria            0
    ## 5 n                    0
    ## 6 Total_individuos     0
    ## 7 Porcentaje           0

``` r
missing_data_top10 #0 missing
```

    ## # A tibble: 7 × 2
    ##   variable           n
    ##   <chr>          <int>
    ## 1 provincia_rh       0
    ## 2 distrito_rh        0
    ## 3 categoria          0
    ## 4 total              0
    ## 5 total_distrito     0
    ## 6 porcentaje         0
    ## 7 ranking            0

# CREATE PPT

``` r
# new
ppt <- read_pptx()

# districts
distritos <- unique(data_top10$distrito_rh)
max_rows <- 11 # Max rows per table

for (i in seq_along(distritos)) {
    distrito <- distritos[i]
    distrito_capitalizado <- str_to_title(distrito)
    tabla_completa <- data_top10 %>%
        filter(distrito_rh == distrito) %>%
        select(ranking, categoria, total, porcentaje, provincia_rh, total_distrito) %>%
        mutate(porcentaje = round(porcentaje, 3)) # Round percentage to 3 decimals

    # Choose the maximum value for total_distrito if there are inconsistencies
    total_poblacion <- max(tabla_completa$total_distrito, na.rm = TRUE)

    num_partes <- ceiling(nrow(tabla_completa) / max_rows)

    for (j in seq_len(num_partes)) {
        tabla <- tabla_completa %>%
            slice(((j - 1) * max_rows + 1):min(j * max_rows, nrow(tabla_completa)))

        titulo <- paste("Diez primeras enfermedades diagnosticadas en atenciones por consulta externa, distrito ",
            distrito_capitalizado, " (n=", total_poblacion, "), Perú, 2023 - (", j, "/", num_partes, ").",
            sep = ""
        )
        nota <- "Fuente: HIS, OGTI, MINSA."

        flextable_tabla <- regulartable(tabla %>% select(-total_distrito)) %>%
            set_header_labels(ranking = "Orden", categoria = "Enfermedad", total = "Número de atenciones", porcentaje = "Porcentaje (%)", provincia_rh = "Provincia") %>%
            theme_vanilla() %>%
            fontsize(size = 10) %>%
            width(j = c("ranking", "categoria", "total", "porcentaje", "provincia_rh"), width = c(1, 3, 1.5, 1, 1)) %>%
            set_table_properties(width = 1, layout = "autofit")

        flextable_tabla <- align(flextable_tabla, j = c("ranking", "total", "porcentaje", "provincia_rh"), align = "center", part = "all")
        flextable_tabla <- align(flextable_tabla, j = "categoria", align = "left", part = "all")
        flextable_tabla <- align(flextable_tabla, j = "categoria", align = "center", part = "header")

        ppt <- add_slide(ppt, layout = "Title and Content", master = "Office Theme")

        title_style <- fp_text(font.size = 20, bold = TRUE)
        formatted_title <- fpar(ftext(titulo, title_style), fp_p = fp_par(text.align = "center"))
        ppt <- ph_with(ppt, value = formatted_title, location = ph_location_type(type = "title"))

        ppt <- ph_with(ppt, value = flextable_tabla, location = ph_location(left = 1, top = 1.5, width = 8, height = 5))

        note_style <- fp_text(font.size = 10)
        formatted_note <- fpar(ftext(nota, note_style))

        if (j == num_partes) {
            ppt <- ph_with(ppt, value = formatted_note, location = ph_location(left = 1, top = 7, width = 8, height = 0.5))
        }
    }
}

#disminuir tamaño de titulo 
ppt <- ph_with(ppt, value = formatted_title, location = ph_location(left = 1, top = 0.5, width = 8, height = 1))
```

# Function to generate districts tables inside each department

``` r
# Save the PowerPoint
print(ppt, target = "general_distritos.pptx")
```

# PPT por etapa de etapa_vida

``` r
library(officer)
library(flextable)
library(dplyr)
library(stringr)

# Function to capitalize words
capitalize_words <- function(s) {
    sapply(strsplit(s, " "), function(x) {
        paste(toupper(substring(x, 1, 1)), tolower(substring(x, 2)), sep = "", collapse = " ")
    })
}

# Crear una nueva presentación de PowerPoint
ppt2 <- read_pptx()

# Agrupar los datos por distrito y etapa de vida
distritos <- unique(datosedad_top10$distrito_rh)
max_rows <- 11 # Max rows per table

for (distrito in distritos) {
    distrito_capitalizado <- capitalize_words(distrito)
    data_distrito <- datosedad_top10 %>% filter(distrito_rh == distrito)

    etapas_vida <- unique(data_distrito$etapa_vida)

    for (etapa in etapas_vida) {
        data_etapa <- data_distrito %>% filter(etapa_vida == etapa)

        tabla_completa <- data_etapa %>%
            select(Ranking, categoria, n, Porcentaje, provincia_rh, Total_individuos) %>%
            mutate(Porcentaje = round(Porcentaje, 3)) # Redondear porcentaje a 3 decimales

        # Escoger el valor máximo de total_distrito si hay inconsistencias
        total_poblacion <- max(tabla_completa$Total_individuos, na.rm = TRUE)

        num_partes <- ceiling(nrow(tabla_completa) / max_rows)

        for (j in seq_len(num_partes)) {
            tabla <- tabla_completa %>%
                slice(((j - 1) * max_rows + 1):min(j * max_rows, nrow(tabla_completa)))

            titulo <- paste("Diez primeras enfermedades diagnosticadas en atenciones por consulta externa, etapa de vida: ",
                etapa, ", distrito ", distrito_capitalizado, " (n=", total_poblacion, "), Perú, 2023 - (", j, "/", num_partes, ").",
                sep = ""
            )
            nota <- "Fuente: HIS, OGTI, MINSA."

            flextable_tabla <- regulartable(tabla %>% select(-Total_individuos)) %>%
                set_header_labels(Ranking = "Orden", categoria = "Enfermedad", n = "Número de atenciones", Porcentaje = "Porcentaje (%)", provincia_rh = "Provincia") %>%
                theme_vanilla() %>%
                fontsize(size = 10) %>%
                width(j = c("Ranking", "categoria", "n", "Porcentaje", "provincia_rh"), width = c(1, 3, 1.5, 1, 1)) %>%
                set_table_properties(width = 1, layout = "autofit")

            flextable_tabla <- align(flextable_tabla, j = c("Ranking", "n", "Porcentaje", "provincia_rh"), align = "center", part = "all")
            flextable_tabla <- align(flextable_tabla, j = "categoria", align = "left", part = "all")
            flextable_tabla <- align(flextable_tabla, j = "categoria", align = "center", part = "header")

            ppt2 <- add_slide(ppt2, layout = "Title and Content", master = "Office Theme")

            title_style <- fp_text(font.size = 20, bold = TRUE)
            formatted_title <- fpar(ftext(titulo, title_style), fp_p = fp_par(text.align = "center"))
            ppt2 <- ph_with(ppt2, value = formatted_title, location = ph_location_type(type = "title"))

            ppt2 <- ph_with(ppt2, value = flextable_tabla, location = ph_location(left = 1, top = 1.5, width = 8, height = 5))

            note_style <- fp_text(font.size = 10)
            formatted_note <- fpar(ftext(nota, note_style))

            if (j == num_partes) {
                ppt2 <- ph_with(ppt2, value = formatted_note, location = ph_location(left = 1, top = 7, width = 8, height = 0.5))
            }
        }
    }
}

# Guardar la presentación
print(ppt2, target = "morbilidad_lima_etapavida.pptx")
```
