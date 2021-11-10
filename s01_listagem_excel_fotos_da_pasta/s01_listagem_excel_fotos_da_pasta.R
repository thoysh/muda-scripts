## Dado 
## - uma PASTA contendo FOTOS,
## o script gera um relatório em EXCEL listando todas as fotos da pasta incluindo:
## Caminho:  Caminho da foto
## Arquivo:  Nome da foto com extensão
## Nome:     Nome da foto sem extensão
## Extensao: Extensão da foto
## Pasta:    Pasta da foto

library(tidyverse)
library(stringr)
library(readxl)
library(writexl)
library(rstudioapi)
rm(list = ls())
cat("\014")


## Dados de input ## >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> IMPORTANTE
pasta <- "fotos"
# pasta <- "C:/Users/coord/Downloads/nome da pasta"


## Dados de configuração
txt_data_hora <- format(Sys.time(), "%Y-%m-%d-%H%M%S")
dir <- rstudioapi::getSourceEditorContext()$path %>% dirname() 
setwd(dir) # define o diretório de trabalho

## Código
df_fotos <- list.files(pasta, pattern = ".jpg|.JPG", # extensão
                       all.files = F,                # arquivos visíveis
                       full.names = T,               # com Caminho
                       recursive = T) %>%            # inclui subdiretórios
  data.frame(Caminho = .) %>% 
  mutate(Arquivo = basename(Caminho),
         Nome = sub(pattern = "(.*)\\..*$", replacement = "\\1", basename(Caminho)),
         Extensao = strsplit(basename(Caminho), split="\\.")[[1]][-1],
         Pasta   = dirname(Caminho))

View(df_fotos) # exibe o relatório
relatorio <- paste0(txt_data_hora, "-relatorio.xlsx")
write_xlsx(df_fotos, relatorio) # cria output: relatório