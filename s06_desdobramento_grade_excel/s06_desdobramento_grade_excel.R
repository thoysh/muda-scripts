## Dado
## - um EXCEL contendo duas abas;
##     - uma aba com dados e pelo menos uma coluna "Grade" (a ser desdobrado) e
##                                      uma coluna "Tamanho" (onde será sobreescrito com o desdobramento)
##     - uma aba com desdobramento com a coluna de referência da "Grade" e cada desdobramento
## o script desdobra cada linha cuja valor na coluna "Grade" da aba de dados tem referência
## na aba de desdobramento, caso contrário copia sem desdobrar
## No final, salva um relatório em EXCEL contendo todo o EXCEL de input desdobrado
## preenchendo a coluna quando possivel

library(tidyverse)
library(stringr)
library(readxl)
library(writexl)
library(rstudioapi)
rm(list = ls())
cat("\014")


## Dados de input ## >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> IMPORTANTE
# excel <- "s06_template.xlsx"
excel <- "FARM - REF PARA DESDOBRAR TAMANHO (1).xlsx"


## Dados de configuração
txt_data_hora <- format(Sys.time(), "%Y-%m-%d-%H%M%S")
dir <- rstudioapi::getSourceEditorContext()$path %>% dirname()
setwd(dir) # define o diretório de trabalho
if (!file.exists(excel))
  print("Excel não encontrado")


## Código
df_dados <- read_excel(excel, sheet = "dados") %>%
  rowid_to_column("Meu_Indice") %>%
  mutate(Grade = Grade %>% as.character(),
         Tamanho = Tamanho %>% as.character())
df_desdobramento <- read_excel(excel, sheet = "desdobramento")
df_desdobramento <- read_excel(excel, sheet = "desdobramento", col_types = rep("text", ncol(df_desdobramento)))
vec_grade <- df_desdobramento$Grade

if (any(duplicated(vec_grade))) {
  stop("Existe desdobramento duplicado")
} else {
  df_desdobramento <- df_desdobramento %>%
    pivot_longer(cols = c(-Grade), names_to = "Coluna", values_to = "Tamanho_Desdobrado") %>%
    select(-Coluna) %>%
    unique() %>%
    na.omit()

  df_planilha <- df_dados %>%
    merge(df_desdobramento, by = "Grade", all.x = T) %>%
    mutate(Tamanho = Tamanho_Desdobrado) %>%
    select(-Tamanho_Desdobrado) %>%
    arrange(Meu_Indice, Tamanho) %>%
    select(all_of(names(df_dados))) %>%
    select(-Meu_Indice)

}

View(df_planilha) # exibe o relatório
relatorio <- paste0(txt_data_hora, "-relatorio.xlsx")
write_xlsx(df_planilha, relatorio) # cria output: relatório
