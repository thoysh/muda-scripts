## Dado 
## - uma FOTO e
## - uma PLANILHA contendo NOMES,
## o script cria uma pasta contendo a mesma foto repetida com todos os nomes da planilha

library(tidyverse)
library(stringr)
library(readxl)
library(writexl)
library(rstudioapi)
rm(list = ls())
cat("\014")


## Dados de input ## >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> IMPORTANTE
foto <- "MUDA-01.jpg"
excel <- "s03_template.xlsx"
# excel <- "C:/Users/coord/Downloads/nome do excel.xlsx"


## Dados de configuração
txt_data_hora <- format(Sys.time(), "%Y-%m-%d-%H%M%S")
dir <- rstudioapi::getSourceEditorContext()$path %>% dirname() 
setwd(dir) # define o diretório de trabalho
if (!file.exists(excel))
  print("Excel não encontrado")
if (!file.exists(foto))
  print("Imagem não encontrado")


## Código
vec_planilha <- read_excel(excel) %>% 
  .[[1]] %>% 
  unique()
extensao <- strsplit(basename(foto), split="\\.")[[1]][-1]

out_pasta <- txt_data_hora
dir.create(file.path(out_pasta), showWarnings = F) # cria output: pasta de fotos

for(nome in vec_planilha) {
  file.copy(from = foto, 
            to = paste0(out_pasta, "/", nome, ".", extensao), 
            overwrite = F, recursive = F,
            copy.mode = T, copy.date = T) # copia/sobreescreve a mesma foto com todos os nomes da lista
}

View(length(vec_planilha)) # exibe o total de fotos
