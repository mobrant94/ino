pacman::p_load(tidyverse, gtsummary, janitor)

setwd("C:/Users/MARCOS.ANTUNES/Downloads")

df = readxl::read_xlsx("no_covid.xlsx")

df$Prontuario = NULL
df$data_coleta = NULL
df$doses = NULL
df$vacinado = NULL
df$motivo_interna = NULL
df$outras = NULL
df$tempo = NULL
df$pa = NULL
df$ifn_alpha = NULL
str(df)

df <- df %>%
  mutate(status = ifelse(grepl("R$", id), "Recoleta", "Coleta"))
df = clean_names(df)
df = df %>% filter(!(id=="COV026"))

df <- df %>%
  mutate(categoria = factor(paste(status, grupo, sep = ": ")))

nova_ordem <- c("Coleta: Grupo 1: Tratamento usual", 
                "Recoleta: Grupo 1: Tratamento usual",
                "Coleta: Grupo 2: Óxido Nítrico",
                "Recoleta: Grupo 2: Óxido Nítrico")

# Reordenando a variável categoria
df$categoria <- factor(df$categoria, levels = nova_ordem)

# Verificando o resumo atualizado da variável categoria
summary(df$categoria)
df <- df[, colSums(is.na(df)) == 0]
# Convertendo a variável idade para uma variável categórica
df <- df %>%
  mutate(faixa_etaria = case_when(
    idade >= 30 & idade <= 39 ~ "30 - 39",
    idade >= 40 & idade <= 64 ~ "40 - 64",
    idade >= 65 ~ "65 >"
  ))



# Supondo que 'dados' seja o nome do seu dataframe
subset <- subset(df, select = c(28:272))
subset = clean_names(subset)

subset1 <- subset(df, select = c(1:41,272))
subset1 = clean_names(subset1)

subset2 <- subset(df, select = c(2,42:272))
subset2 = clean_names(subset2)
subset2 <- subset2 %>%
  mutate_at(vars(2:272), as.numeric)
subset1$id = NULL
subset1$tempo=NULL
subset1$peso = as.numeric(subset1$peso)
subset1$altura = as.numeric(subset1$altura)
subset1$imc = as.numeric(subset1$imc)
subset1$spo2_gaso = as.numeric(subset1$spo2_gaso)
subset1$peso = as.numeric(subset1$peso)
subset1$los = as.numeric(subset1$los)



subset$t8_8 = as.numeric(subset$t8_8)







library(gtsummary)
library(titanic)
library(tidyverse)
library(plotrix) 


# first, write a little function to get the 2-way ANOVA p-values in a table
# function to get 2-way ANOVA p-values in tibble
twoway_p <- function(variable) {
  paste(variable, "~ grupo * status") %>% #######checar anova de duas vias como medidas repetitivas
    as.formula() %>%
    aov(data = subset2) %>% 
    broom::tidy() %>%
    select(term, p.value) %>%
    filter(complete.cases(.)) %>%
    pivot_wider(names_from = term, values_from = p.value) %>%
    mutate(
      variable = .env$variable,
      row_type = "label"
    )
}

# add all results to a single table (will be merged with gtsummary table in next step)
twoway_results <- bind_rows(
  twoway_p("eritrocitos"),
  twoway_p("hemato"),
  twoway_p("hemog"),
  twoway_p("vcm"),
  twoway_p("hcm"),
  twoway_p("rdw"),
  twoway_p("eritoblastos"),
  twoway_p("leuco"),
  twoway_p("neutro"),
  twoway_p("bastonatos"),
  twoway_p("segmentados"),
  twoway_p("basofilos"),
  twoway_p("eosinofilos"),
  twoway_p("monocitos"),
  twoway_p("linfocitos"),
  twoway_p("plaquetas"),
  twoway_p("pcr"),
  twoway_p("dimeros"),
  twoway_p("ldh"),
  twoway_p("s_trem_1"),
  twoway_p("il_12p70"),
  twoway_p("tnf_alpha"),
  twoway_p("il_10"),
  twoway_p("il_6"),
  twoway_p("il_1_beta"),
  twoway_p("il_8"),
  twoway_p("ifn_alpha"),
  twoway_p("t1_1"),
  twoway_p("t1_2"),
  twoway_p("t1_3"),
  twoway_p("t1_4"),
  twoway_p("t1_5"),
  twoway_p("t1_6"),
  twoway_p("t1_7"),
  twoway_p("t1_8"),
  twoway_p("t1_9"),
  twoway_p("t1_10"),
  twoway_p("t1_12"),
  twoway_p("t1_13"),
  twoway_p("t1_14"),
  twoway_p("t1_15"),
  twoway_p("t1_16"),
  twoway_p("t1_17"),
  twoway_p("t1_20"),
  twoway_p("t1_21"),
  twoway_p("t1_22"),
  twoway_p("t1_23"),
  twoway_p("t1_24"),
  twoway_p("t1_25"),
  twoway_p("t1_26"),
  twoway_p("t1_27"),
  twoway_p("t1_28"),
  twoway_p("t1_29"),
  twoway_p("t1_30"),
  twoway_p("t2_1"),
  twoway_p("t2_2"),
  twoway_p("t2_3"),
  twoway_p("t2_4"),
  twoway_p("t2_5"),
  twoway_p("t2_6"),
  twoway_p("t2_7"),
  twoway_p("t2_8"),
  twoway_p("t2_9"),
  twoway_p("t2_10"),
  twoway_p("t2_11"),
  twoway_p("t2_12"),
  twoway_p("t2_13"),
  twoway_p("t2_14"),
  twoway_p("t2_15"),
  twoway_p("t2_16"),
  twoway_p("t2_17"),
  twoway_p("t2_18"),
  twoway_p("t2_19"),
  twoway_p("t2_20"),
  twoway_p("t2_21"),
  twoway_p("t2_22"),
  twoway_p("t2_23"),
  twoway_p("t2_24"),
  twoway_p("t2_25"),
  twoway_p("t2_26"),
  twoway_p("t2_27"),
  twoway_p("t2_28"),
  twoway_p("t3_1"),
  twoway_p("t3_2"),
  twoway_p("t3_3"),
  twoway_p("t3_4"),
  twoway_p("t3_5"),
  twoway_p("t3_6"),
  twoway_p("t3_8"),
  twoway_p("t3_9"),
  twoway_p("t3_10"),
  twoway_p("t3_11"),
  twoway_p("t3_12"),
  twoway_p("t3_13"),
  twoway_p("t3_15"),
  twoway_p("t3_16"),
  twoway_p("t3_17"),
  twoway_p("t3_18"),
  twoway_p("t3_19"),
  twoway_p("t3_20"),
  twoway_p("t3_22"),
  twoway_p("t3_23"),
  twoway_p("t3_24"),
  twoway_p("t3_25"),
  twoway_p("t3_26"),
  twoway_p("t3_27"),
  twoway_p("t3_28"),
  twoway_p("t3_29"),
  twoway_p("t4_1"),
  twoway_p("t4_2"),
  twoway_p("t4_3"),
  twoway_p("t4_4"),
  twoway_p("t4_5"),
  twoway_p("t4_6"),
  twoway_p("t4_7"),
  twoway_p("t4_8"),
  twoway_p("t4_9"),
  twoway_p("t4_10"),
  twoway_p("t4_11"),
  twoway_p("t4_13"),
  twoway_p("t4_14"),
  twoway_p("t4_15"),
  twoway_p("t4_16"),
  twoway_p("t4_18"),
  twoway_p("t4_19"),
  twoway_p("t4_20"),
  twoway_p("t4_22"),
  twoway_p("t4_23"),
  twoway_p("t4_24"),
  twoway_p("t4_25"),
  twoway_p("t4_26"),
  twoway_p("t4_27"),
  twoway_p("t4_29"),
  twoway_p("t4_30"),
  twoway_p("t4_31"),
  twoway_p("t4_32"),
  twoway_p("t4_33"),
  twoway_p("t4_34"),
  twoway_p("t4_35"),
  twoway_p("t4_36"),
  twoway_p("t4_37"),
  twoway_p("t4_38"),
  twoway_p("t4_39"),
  twoway_p("t4_41"),
  twoway_p("t4_42"),
  twoway_p("t4_43"),
  twoway_p("t4_44"),
  twoway_p("t4_45"),
  twoway_p("t4_46"),
  twoway_p("t4_47"),
  twoway_p("t4_49"),
  twoway_p("t4_50"),
  twoway_p("t4_51"),
  twoway_p("t4_51"),
  twoway_p("t4_52"),
  twoway_p("t4_53"),
  twoway_p("t4_54"),
  twoway_p("t5_1"),
  twoway_p("t5_2"),
  twoway_p("t5_3"),
  twoway_p("t5_4"),
  twoway_p("t5_5"),
  twoway_p("t5_6"),
  twoway_p("t5_7"),
  twoway_p("t5_8"),
  twoway_p("t5_9"),
  twoway_p("t5_10"),
  twoway_p("t5_11"),
  twoway_p("t5_12"),
  twoway_p("t5_13"),
  twoway_p("t5_14"),
  twoway_p("t5_15"),
  twoway_p("t5_16"),
  twoway_p("t5_17"),
  twoway_p("t5_18"),
  twoway_p("t6_1"),
  twoway_p("t6_2"),
  twoway_p("t6_3"),
  twoway_p("t6_4"),
  twoway_p("t6_5"),
  twoway_p("t6_6"),
  twoway_p("t6_7"),
  twoway_p("t6_8"),
  twoway_p("t6_9"),
  twoway_p("t6_10"),
  twoway_p("t7_1"),
  twoway_p("t7_2"),
  twoway_p("t7_3"),
  twoway_p("t7_4"),
  twoway_p("t7_5"),
  twoway_p("t7_6"),
  twoway_p("t7_7"),
  twoway_p("t7_8"),
  twoway_p("t7_10"),
  twoway_p("t7_11"),
  twoway_p("t7_12"),
  twoway_p("t7_13"),
  twoway_p("t7_14"),
  twoway_p("t7_15"),
  twoway_p("t7_16"),
  twoway_p("t7_17"),
  twoway_p("t7_18"),
  twoway_p("t7_19"),
  twoway_p("t7_20"),
  twoway_p("t7_22"),
  twoway_p("t7_23"),
  twoway_p("t7_24"),
  twoway_p("t8_1"),
  twoway_p("t8_2"),
  twoway_p("t8_4"),
  twoway_p("t8_5"),
  twoway_p("t8_6"),
  twoway_p("t8_8"),
  twoway_p("t8_9"),
  twoway_p("t8_10"),
  twoway_p("t8_11"),
  twoway_p("t8_12"),
  twoway_p("t8_13"),
  twoway_p("t8_14"),
  twoway_p("t8_16"),
  twoway_p("t8_17"),
  twoway_p("t8_20"),
  twoway_p("t8_21"),
  twoway_p("t8_23"),
  twoway_p("t8_24"),
  twoway_p("t8_26"),
  twoway_p("t8_27"),
  twoway_p("t8_30"),
  twoway_p("t8_31"),
  twoway_p("t8_33"),
  twoway_p("t8_34"),
  twoway_p("t8_35"),
  twoway_p("t8_36"),
  twoway_p("t8_37"),
  twoway_p("t8_38"),
  twoway_p("t8_39"),
  twoway_p("t8_40"),
  twoway_p("t8_42"),
  twoway_p("t8_43"),
  twoway_p("t8_45"),
  twoway_p("t8_46"),
  twoway_p("t8_47"),
  twoway_p("t8_49"),
  twoway_p("t8_50"),
  twoway_p("t8_51"),
  twoway_p("t8_52"),
  twoway_p("t8_53"),
  twoway_p("t8_54"),
  twoway_p("t8_55"),
  twoway_p("t8_57"),
  twoway_p("t8_62"),
  twoway_p("t8_64"),
  twoway_p("t8_65"),
  twoway_p("t8_67"),
  twoway_p("t8_68"),
  twoway_p("t8_74"),
  twoway_p("t8_75"),
  twoway_p("t8_76"),
  twoway_p("t8_77"),
  twoway_p("t8_78"),
  twoway_p("t8_79"),
  twoway_p("t8_80"),
  twoway_p("t8_81"))

twoway_results



tbl <-
  # first build a stratified `tbl_summary()` table to get summary stats by two variables
  subset2 %>%
  tbl_strata(
    strata =  grupo,
    .tbl_fun =
      ~.x %>%
      tbl_summary(
        by = status,
        missing = "no",
        type = list(il_12p70~ "continuous",
                    tnf_alpha~ "continuous",
                    il_1_beta~"continuous",
                    ifn_alpha~"continuous",
                    t8_8 ~ "continuous",
                    tnf_alpha ~ "continuous"),
        statistic = all_continuous() ~ "{mean} ({std.error})",
        digits = everything() ~ 1
      ) %>%
      modify_header(all_stat_cols() ~ "**{level}**")
  ) %>%
  # merge the 2way ANOVA results into tbl_summary table
  modify_table_body(
    ~.x %>%
      left_join(
        twoway_results,
        by = c("variable", "row_type")
      )
  ) %>%
  # by default the new columns are hidden, add a header to unhide them
  modify_header(list(
    grupo ~ "**Grupo**", 
    status ~ "**status**", 
    `grupo:status` ~ "**grupo * status**"
  )) %>%
  # adding spanning header to analysis results
  modify_spanning_header(c(grupo, status, `grupo:status`) ~ "**Two-way ANOVA p-values**") %>%
  # format the p-values with a pvalue formatting function
  modify_fmt_fun(c(grupo, status, `grupo:status`) ~ style_pvalue) %>%
  # update the footnote to be nicer looking
  modify_footnote(all_stat_cols() ~ "Mean (SE)")

tbl = tbl %>% as_flex_table() %>% save_as_docx(path = "two_waysaov.docx")


tbl1 = tbl_summary(subset2, 
                   by = categoria,
                   type = list(t8_8 ~ "continuous",
                               tnf_alpha ~ "continuous"),
                   missing = "no") %>% add_p(test = everything() ~ "aov")

tbl1 = tbl1 %>% as_flex_table() %>% save_as_docx(path = "indepent_aov.docx")


subset1$peso = as.numeric(subset1$peso)

tbl2 = tbl_summary(subset1, 
                   by = grupo,
                   type = list(peso~ "continuous",
                               altura~ "continuous",
                               imc~"continuous",
                               los~"continuous"),
                   missing = "no") %>% add_p()

tbl2 = tbl2 %>% as_flex_table() %>% save_as_docx(path = "demo_indepent_aov.docx")



# set theme to get MEAN (SD) by default in `tbl_summary()`
theme_gtsummary_mean_sd()

# function to add pairwise copmarisons to `tbl_summary()`
add_stat_pairwise <- function(data, variable, by, ...) {
  # calculate pairwise p-values
  pw <- pairwise.t.test(data[[variable]], data[[by]], p.adj = "none")
  
  # convert p-values to list
  index <- 0L
  p.value.list <- list()
  for (i in seq_len(nrow(pw$p.value))) {
    for (j in seq_len(nrow(pw$p.value))) {
      index <- index + 1L
      
      p.value.list[[index]] <- 
        c(pw$p.value[i, j]) %>%
        setNames(glue::glue("**{colnames(pw$p.value)[j]} vs. {rownames(pw$p.value)[i]}**"))
    }
  }
  
  # convert list to data frame
  p.value.list %>% 
    unlist() %>%
    purrr::discard(is.na) %>%
    t() %>%
    as.data.frame() %>%
    # formatting/roundign p-values
    dplyr::mutate(dplyr::across(everything(), style_pvalue))
}

d = subset2 %>% 
  tbl_summary(by = categoria, missing = "no",
              type = list(tnf_alpha~ "continuous")) %>%
  # add pariwaise p-values
  add_stat(everything() ~ add_stat_pairwise) 

d = d %>% as_flex_table() %>% save_as_docx(path = "comparacoes_aov.docx")



anova_results <- lapply(vars_to_analyze, function(var) {
  formula <- as.formula(paste(var, " ~ grupo_numerico + status_numerico + grupo_numerico:status_numerico"))
  model <- aov(formula, data = subset)
  tidy_results <- tidy(model)
  p_value <- tidy_results %>%
    filter(term == "grupo_numerico:status_numerico") %>%
    pull(p.value)
  return(list(variable = var, p_value = p_value))
})

# Converter os resultados em um data frame
anova_df <- bind_rows(anova_results)

# Exibir o data frame
print(anova_df)





subset

vars_to_analyze <- names(subset)[sapply(subset, is.numeric)]

anova_results <- lapply(vars_to_analyze, function(var) {
  formula <- as.formula(paste(var, " ~ categoria/status"))
  model <- aov(formula, data = subset)
  tidy_results <- tidy(model)
  p_value <- tidy_results %>%
    pull(p.value)
  return(list(variable = var, p_value = p_value))
}) %>% na.omit()


# Converter os resultados em um data frame
anova_df <- bind_rows(anova_results) %>% na.omit()

writexl::write_xlsx(anova_results2,"tabelag.xlsx")


t1 = subset %>%
  tbl_strata(
    strata =  categoria,
    .tbl_fun =
      ~.x %>%
      tbl_summary(
        by = status,
        missing = "no",
        type = list(tnf_alpha~"continuous",
                    il_12p70~"continuous",
                    il_1_beta	~"continuous"),
        statistic = all_continuous() ~ "{mean} ({sd})",
        digits = everything() ~ 1
      ) %>%
      modify_header(all_stat_cols() ~ "**{level}**")
  )


t1 = t1 %>% as_flex_table() %>% save_as_docx(path = "nested_aov.docx")

library(flextable)
######age class############################################################


subset3039 = subset %>% filter(faixa_etaria == "30 - 39")

vars_to_analyze1 <- names(subset3039)[sapply(subset3039, is.numeric)]

anova_results1 <- lapply(vars_to_analyze1, function(var) {
  formula <- as.formula(paste(var, " ~ categoria/status"))
  model <- aov(formula, data = subset3039)
  tidy_results <- tidy(model)
  p_value <- tidy_results %>%
    pull(p.value)
  return(list(variable = var, p_value = p_value))
}) %>% na.omit()

# Converter os resultados em um data frame
anova_df1 <- bind_rows(anova_results1) %>% na.omit()
subset3039$faixa_etaria = NULL

writexl::write_xlsx(anova_results1,"tabela3039.xlsx")

t2 = subset3039 %>%
  tbl_strata(
    strata =  categoria,
    .tbl_fun =
      ~.x %>%
      tbl_summary(
        by = status,
        missing = "no",
        type = everything()~"continuous",
        statistic = all_continuous() ~ "{mean} ({sd})",
        digits = everything() ~ 1
      ) %>%
      modify_header(all_stat_cols() ~ "**{level}**")
  )

t2 = t2 %>% as_flex_table() %>% save_as_docx(path = "3039_aov.docx")

###################


subset4064 = subset %>% filter(faixa_etaria == "40 - 64")
subset4064$faixa_etaria = NULL
vars_to_analyze2 <- names(subset4064)[sapply(subset4064, is.numeric)]

anova_results2 <- lapply(vars_to_analyze2, function(var) {
  formula <- as.formula(paste(var, " ~ categoria/status"))
  model <- aov(formula, data = subset4064)
  tidy_results <- tidy(model)
  p_value <- tidy_results %>%
    pull(p.value)
  return(list(variable = var, p_value = p_value))
}) %>% na.omit()

writexl::write_xlsx(anova_results2,"tabela4060.xlsx")

# Converter os resultados em um data frame
anova_df2 <- bind_rows(anova_results2) %>% na.omit()

t3 = subset4064 %>%
  tbl_strata(
    strata =  categoria,
    .tbl_fun =
      ~.x %>%
      tbl_summary(
        by = status,
        missing = "no",
        type = everything()~"continuous",
        statistic = all_continuous() ~ "{mean} ({sd})",
        digits = everything() ~ 1
      ) %>%
      modify_header(all_stat_cols() ~ "**{level}**")
  )

t3 = t3 %>% as_flex_table() %>% save_as_docx(path = "4064_aov.docx")

#####################

subset65 = subset %>% filter(faixa_etaria == "65 >")
subset65$faixa_etaria = NULL
vars_to_analyze3 <- names(subset65)[sapply(subset65, is.numeric)]
subset65$subset65 = NULL
anova_results3 <- lapply(vars_to_analyze3, function(var) {
  formula <- as.formula(paste(var, " ~ categoria/status"))
  model <- aov(formula, data = subset65)
  tidy_results <- tidy(model)
  p_value <- tidy_results %>%
    pull(p.value)
  return(list(variable = var, p_value = p_value))
}) %>% na.omit()

writexl::write_xlsx(anova_results3,"tabela65.xlsx")

# Converter os resultados em um data frame
anova_df3 <- bind_rows(anova_results3) %>% na.omit()

t4 = subset65 %>%
  tbl_strata(
    strata =  categoria,
    .tbl_fun =
      ~.x %>%
      tbl_summary(
        by = status,
        missing = "no",
        type = everything()~"continuous",
        statistic = all_continuous() ~ "{mean} ({sd})",
        digits = everything() ~ 1
      ) %>%
      modify_header(all_stat_cols() ~ "**{level}**")
  )

t4 = t4 %>% as_flex_table() %>% save_as_docx(path = "65_aov.docx")

