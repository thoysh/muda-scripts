## Dado 
## - um XML contendo as produtos com tamanhos do estoque do PDVNET,
## o script converte a formato XML para um EXCEL legível incluindo:
## Ref:   Referência ddo produto
## Nome:  Nome do produto
## Total: Quantidade total do produto
## [Tamanhos]: Quantidade do produto por tamanho (ex.: U, PP, P, M, G, GG, 34, 35, 36 etc)

library(tidyverse)
library(stringr)
library(readxl)
library(writexl)
library(xml2)
library(rstudioapi)
rm(list = ls())
cat("\014")


## Dados de input ## >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> IMPORTANTE
arquivo <- "s05_template.xml"
# arquivo <- "C:/Users/coord/Downloads/nome do arquivo.xml"


## Dados de configuração
txt_data_hora <- format(Sys.time(), "%Y-%m-%d-%H%M%S")
dir <- rstudioapi::getSourceEditorContext()$path %>% dirname() 
setwd(dir) # define o diretório de trabalho
if (!file.exists(arquivo))
  print("Xml não encontrado")


## Código
data <- read_xml(arquivo)
xml_fap_l2 <- xml_find_all(data, ".//d1:FormattedAreaPair[@Level='2']") # Leitura do mais alto nível

lst_produto <- list()
for (pos in 1:length(xml_fap_l2)) {
  this_xml <- xml_fap_l2[[pos]] # define a parte do xml com a identificação
  
  ref <- xml_find_first(this_xml, ".//d1:FormattedReportObject[@FieldName='{REFERENCIAS.REF_REFERENCIA}']/d1:FormattedValue") %>%
    xml_text() # encontra a referência do produto
  nome <- xml_find_first(this_xml, ".//d1:FormattedReportObject[@FieldName='{REFERENCIAS.REF_DESCRICAO}']/d1:FormattedValue") %>%
    xml_text() # encontra o nome do produto
  
  this_xml <- xml_find_first(this_xml, # define a parte do xml com a tabela
                             ".//d1:FormattedSection[@SectionNumber='1']//d1:FormattedReportObject") 
  
  # vec_cor <- xml_find_all(this_xml, ".//d1:FormattedRowGroup/d1:FormattedRowGroup/d1:FormattedRowTotal") %>% 
  vec_cor <- xml_find_all(this_xml, ".//d1:FormattedRowTotal") %>% 
    xml_text() %>% 
    str_trim() # encontra as cores do produto
  # vec_tamanho <- xml_find_all(this_xml, ".//d1:FormattedColumnGroup/d1:FormattedColumnGroup/d1:FormattedColumnGroup/d1:FormattedColumnTotal") %>% 
  vec_tamanho <- xml_find_all(this_xml, ".//d1:FormattedColumnTotal") %>% 
    xml_text() %>% 
    str_trim() # encontra os tamanhos do produto
  
  df_codigo <- data.frame(matrix(ncol = length(vec_tamanho), nrow = length(vec_cor))) ## cria o data.frame
  colnames(df_codigo) <- vec_tamanho
  rownames(df_codigo) <- vec_cor
  
  for (cor in 1:length(vec_cor)) {
    vec_quantidade <- xml_find_all(this_xml, 
                                   paste0(".//d1:FormattedCell[@RowNumber='", cor + 1, "']", 
                                          "/d1:FormattedCellValues/d1:CellValue")) %>% 
      xml_text() %>% 
      str_trim() %>% 
      as.numeric() # encontra as quantidades da cor do produto
    df_codigo[cor,] <- vec_quantidade
  }
  
  vec_rm <- intersect(colnames(df_codigo), as.character(c(1:20, 50:70)))
  df_codigo <- df_codigo %>% 
    select(-all_of(vec_rm)) %>% # limpa as duplicações de tamanhos
    mutate(Ref = ref,
           Nome = nome)
  df_codigo <- df_codigo[row.names(df_codigo) != "Total", , drop = F] # remove a linha total
  
  lst_produto <- append(lst_produto, list(df_codigo), 0)
}

df_estoque <- bind_rows(lst_produto) # mescla todos os produtos em um data.frame

vec_ordem <- c("Ref", "Nome", "Total", "U", "PP", "P", "M", "G", "GG")
vec_ordem <- c(intersect(vec_ordem, colnames(df_estoque)), 
               sort(setdiff(colnames(df_estoque), vec_ordem)))

df_estoque <- df_estoque %>% 
  select(all_of(vec_ordem)) # ordena as colunas

View(df_estoque) # exibe o relatório
relatorio <- paste0(txt_data_hora, "-relatorio.xlsx")
write_xlsx(df_estoque, relatorio) # cria output: relatório