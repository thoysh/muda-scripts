## Dado 
## - um EXCEL contendo uma relação CODIGO (123456) e COD_BARRAS (1234560001PP) e
## - uma PASTA contendo FOTOS (123456M_1.jpg),
## o script copia a FOTO _1 cujo nome bata com o CODIGO (123456XX_1) para uma pasta e renomeia para COD_BARRAS.jpg
## No final, salva um relatório em EXCEL contendo todo o EXCEL de input incluindo:
## RefBusca:   Referência da busca (CODIGO)
## CodBarras:  Código de barras (COD_BARRAS)
## Caminho:    Caminho da foto (se encontrada)
## Arquivo:    Nome da foto com extensão (se encontrada)
## Nome:       Nome da foto sem extensão (se encontrada)
## Extensao:   Extensão da foto (se encontrada)
## Pasta:      Pasta da foto (se encontrada)
## Encontrado: Flag TRUE caso encontrada, FALSO cc
## Numero:     Número da foto encontrada (se encontrada)

library(tidyverse)
library(stringr)
library(readxl)
library(writexl)
library(rstudioapi)
rm(list = ls())
cat("\014")


## Dados de input ## >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> IMPORTANTE
excel <- "s04_template.xlsx"
pasta <- "fotos"
# pasta <- "C:/Users/coord/Downloads/nome da pasta"


## Dados de configuração
txt_data_hora <- format(Sys.time(), "%Y-%m-%d-%H%M%S")
dir <- rstudioapi::getSourceEditorContext()$path %>% dirname() 
setwd(dir) # define o diretório de trabalho
define_refbusca <- function(ref) { ## >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> IMPORTANTE
  # substring(ref, 4, 9) ## Farm
  substring(ref, 1, 6) ## Cantão
}
if (!file.exists(excel))
  print("Excel não encontrado")


## Código
df_planilha <- read_excel(excel) %>% 
  .[1:2]
names(df_planilha) <- c("RefBusca", 'CodBarras')

df_fotos <- list.files(pasta, pattern = ".jpg|.JPG", # extensão
                       all.files = F,                # arquivos visíveis
                       full.names = T,               # com caminho
                       recursive = T) %>%            # inclui subdiretórios
  data.frame(Caminho = .) %>% 
  mutate(Arquivo = basename(Caminho),
         Nome = sub(pattern = "(.*)\\..*$", replacement = "\\1", basename(Caminho)),
         Extensao = strsplit(basename(Caminho), split="\\.")[[1]][-1],
         Pasta   = dirname(Caminho),
         Encontrado = "Não") %>% 
  mutate(RefBusca = Nome %>% 
           ## IMPORTANTE: define padrão de consulta do nome das fotos da pasta de input 
           define_refbusca()) %>% 
  mutate(Numero = Nome %>% 
           substr(nchar(Nome) - 1, nchar(Nome))) %>% 
  filter(Numero == "_1")

df_planilha <- df_planilha %>% 
  merge(df_fotos, by = "RefBusca", all.x = T) # faz o merge dos data.frames para o relatório

out_pasta <- txt_data_hora
dir.create(file.path(out_pasta), showWarnings = F) # cria output: pasta de fotos

for (lin in rownames(df_planilha)) {
  df_linha <- df_planilha[lin,]
  
  if (!is.na(df_linha$Caminho)) {
    file.copy(from = df_linha$Caminho, 
              to = paste0(out_pasta, "/", df_linha$CodBarras, ".", df_linha$Extensao), 
              overwrite = F, recursive = F,
              copy.mode = T, copy.date = T) # copia/sobreescreve todas as fotos que forem encontradas
    
    df_planilha[lin,]$Encontrado <- "Sim"
  }
}

View(df_planilha) # exibe o relatório
relatorio <- paste0(txt_data_hora, "-relatorio.xlsx")
write_xlsx(df_planilha, relatorio) # cria output: relatório
