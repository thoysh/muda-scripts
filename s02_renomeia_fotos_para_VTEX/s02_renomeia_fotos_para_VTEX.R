## Dado
## - um EXCEL com a coluna _CodigoReferenciaProduto e
## - uma PASTA contendo FOTOS (528263031P_1.jpg) das sessões,
## o script cria uma pasta e copia dentro dela todas as fotos que encontra
## na qual a referência de busca (parte do _CodigoReferenciaProduto) no EXCEL
## bate com o nome da foto (parcialmente, via função)
## No final, salva um relatório em EXCEL contendo duas abas:
## aba: Referencias que contém todo o EXCEL de input incluindo
##   CodRefProd:       Dado original
##   RefBusca:         Valor usado para busca das fotos
##   FotosEncontradas: Quantidade de fotos encontradas (pode haver mais quando um produto tem mais de uma cor, pois o código da busca não limita cor)
##   FotosNomeCorreto: Quantidade de fotos com nome correto (o padrão é procurar - ou _ e depois um número antes da extensão, o que não acontece sempre)
##   FotosCopiadas:    Quantidade de fotos copiadas (só deve dar erro em caso de duplicação de nome de foto na pasta ou _CodigoReferenciaProduto repetido)
##   Status:           Breve descrição do que foi encontrado de erro
##   Repetido:         Flag CodRefProd repetido. Novamente, quando um produto tem mais de uma cor, todas as fotos são copiadas, então todos que tem essa flag devem ser verificados
## aba: Fotos que contém todas as fotos da pasta de input
##   Caminho:    Caminho da foto
##   Arquivo:    Nome da foto com extensão
##   Pasta:      Pasta da foto
##   Encontrada: Flag de foto encontrada usando as RefBusca. Uma foto não encontrada não foi aproveitada após a sessão de fotos.

library(tidyverse)
library(stringr)
library(readxl)
library(writexl)
library(tools)
library(rstudioapi)
rm(list = ls())
cat("\014")


## Dados de input ## >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> IMPORTANTE
pasta <- "fotos"
pasta <- "C:/Users/u4fl/Downloads/img"
excel <- "s02_template-2022.xls"
# excel <- "C:/Users/coord/Downloads/nome do excel.xls"


## Dados de configuração
txt_data_hora <- format(Sys.time(), "%Y-%m-%d-%H%M%S")
dir <- rstudioapi::getSourceEditorContext()$path %>% dirname()
setwd(dir) # define o diretório de trabalho
define_refbusca <- function(ref) { ## >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> IMPORTANTE
  ## FORMATO ANTIGO até 2022-03
  ## FARM   | FAR2835404438 | 283540
  # substring(ref, 4, 9) ## Farm
  ## CANTÃO | 521590031     | 521590
  # substring(ref, 1, 6) ## Cantão

  ## FORMATO NOVO após 2022-04
  prefix <- 3
  marca <- function(ref) {
    substr(ref, 1, prefix)
  }
  case_when(##AGILITÀ
            # AGI100075BL004 | 100075 | referência pode ter 5 ou 6 caracteres, mas
            # AGI97205BM008  | 97205  | o caracter seguinte é uma letra
            marca(ref) == "AGI" ~
              ifelse(substr(ref, prefix + 6, prefix + 6) %>%
                       as.numeric() %>%
                       is.na() %>%
                       suppressWarnings(),
                     substr(ref, prefix + 1, prefix + 5),
                     substr(ref, prefix + 1, prefix + 6)),
            ##CANTÃO
            # CAN521590031 | 521590 | referência tem 6 caracteres
            marca(ref) == "CAN" ~
              substr(ref, prefix + 1, prefix + 6),
            ##FARM
            # FAR2773011719 | 277301 | referência tem 6 caracteres
            marca(ref) == "FAR" ~
              substr(ref, prefix + 1, prefix + 6),
            ##DRESS TO
            # DRE0208132811 | 02081328 | referência tem 8 caracteres
            marca(ref) == "DRE" ~
              substr(ref, prefix + 1, prefix + 8),
            ##LOVE DRESS
            # LOV042904671638 | 04290467 | referência tem 8 caracteres
            marca(ref) == "LOV" ~
              substr(ref, prefix + 1, prefix + 8),
            ##Default
            TRUE ~ "Referência errada"
  )
}
define_regexBusca <- function(ref) { ## >>>>>>>>>>>>>>>>>>>>>>>>>>>>>> IMPORTANTE
  ## FORMATO ANTIGO até 2022-03
  ## FARM   | FAR2835404438 pode ter FAR ou não na frente do 283540 no nome da imagem
  # paste0("^(", ref, "|", "FAR", ref, ")") ## inicia com xxx ou FARxxx
  ## CANTÃO | 521590031 busca comum no início da imagem
  # paste0("^(", ref, ")") ## inicia com xxx


  paste0("^(", ref, ")") ## inicia com xxx
}
if (!file.exists(excel))
  print("Excel não encontrado")


## Código
# lista todas as fotos da pasta de input respeitando a extensão
df_fotos <- list.files(pasta, pattern = ".jpg|.JPG|.png|.PNG", # extensão
                       all.files = F,                          # arquivos visíveis
                       full.names = T,                         # com Caminho
                       recursive = T) %>%                      # inclui subdiretórios
  data.frame(Caminho = .) %>%
  mutate(Arquivo    = basename(Caminho),
         Nome       = file_path_sans_ext(Arquivo) %>% str_squish() %>% str_replace_all(" ", "-"),
         Extensao   = file_ext(Arquivo),
         Pasta      = dirname(Caminho),
         Encontrada = F
         ) %>%
  mutate(IniNum = Nome %>% str_match("(.*?)(_|-)") %>% .[, 2],
         IniWtv = Nome %>% substring(nchar(IniNum) + 1),
         FimNum = Nome %>% str_match("(?!.*(_|-))(.*)") %>% .[, 3],
         FimWtv = Nome %>% substr(1, nchar(Nome) - nchar(FimNum))
         ) %>%
  mutate(NumFinal  = ifelse(IniNum %in% c(1:9), IniNum,
                            ifelse(FimNum %in% c(1:9), FimNum,
                                   NA)),
         NomeFinal = ifelse(IniNum %in% c(1:9), IniWtv,
                            ifelse(FimNum %in% c(1:9), FimWtv,
                                   Nome)) %>%
           str_replace_all("-|_|\\.", "")
         ) %>%
  select(-IniNum, -IniWtv, -FimNum, -FimWtv)

# lista todas as referências de busca da marca listados no Excel
df_planilha <- read_excel(excel) %>%
  transmute(CodRefProd = `_CodigoReferenciaProduto` %>% as.character()) %>%
  mutate(RefBusca = CodRefProd %>%
           ## IMPORTANTE: define padrão de consulta do nome das fotos da pasta de input
           define_refbusca()) %>%
  # remove as referências duplicadas de tamanho (P, M, G têm a mesma foto de produto por exemplo)
  distinct() %>%
  mutate(
    FotosEncontradas = 0,
    FotosNomeCorreto = 0,
    FotosCopiadas = 0,
    Status = NA,
    # indica as potenciais referências duplicadas de cores (não há como diferenciar por falta de informação)
    Repetido = ifelse(ave(RefBusca, RefBusca, FUN = length) > 1,
                      "i. Verificar! Tem mais de uma cor",
                      NA)
  )

# passa por cada referência de produto (não apenas a de busca, devido ao problema das cores)
for (lin in seq_len(nrow(df_planilha))) {

  out_pasta <- txt_data_hora
  dir.create(file.path(out_pasta), showWarnings = F) # cria output: pasta de fotos

  # constrói o regex de busca para foto
  ref <- df_planilha$RefBusca[lin]
  padrao_regex_foto <- define_regexBusca(ref)

  # procura por todas as fotos que tenham no nome da foto CORRIGIDO a referência de busca conforme o regex
  # marca como encontrada no data.frame principal
  df_fotos <- df_fotos %>%
    mutate(Encontrada = ifelse(Encontrada == T, T,
                               str_detect(toupper(NomeFinal), pattern = padrao_regex_foto)))

  # procura por todas as fotos que tenham no nome da foto CORRIGIDO a referência de busca conforme o regex
  # filtra as encontradas no data.frame auxiliar
  df_encontradas <- df_fotos %>%
    filter(str_detect(toupper(Arquivo), pattern = padrao_regex_foto))

  num_fotos_encontradas <- nrow(df_encontradas)
  df_planilha$FotosEncontradas[lin] = num_fotos_encontradas # atualiza fotos encontradas

  if (num_fotos_encontradas == 0) {
    df_planilha$Status[lin] = "0. Alerta! Nenhuma foto encontrada"
  } else {
    tryCatch({
      # define o novo nome da foto encontrada (CodRefProd-1.jpg)
      df_encontradas <- df_encontradas %>%
        mutate(novo_arquivo_nome = df_planilha$CodRefProd[lin],
               novo_arquivo_sep  = "_",
               novo_arquivo_num  = NumFinal,
               novo_arquivo_ext  = paste0(".", tolower(Extensao))) %>%
        mutate(
          novo_arquivo = paste0(novo_arquivo_nome,
                                novo_arquivo_sep,
                                novo_arquivo_num,
                                novo_arquivo_ext))

      # verifica se tem mais que o limite de fotos
      num_fotos_nome_correto <- df_encontradas$novo_arquivo_num %>%
        as.numeric(.) %>%
        replace_na(., Inf)
      num_fotos_nome_correto <- sum(unique(num_fotos_nome_correto) < 7) ## limite de fotos
      df_planilha$FotosNomeCorreto[lin] = num_fotos_nome_correto # atualiza fotos nome correto

      # verifica se tem mais que o limite de fotos
      num_fotos_copiadas <- file.copy(from = df_encontradas$Caminho,
                            to = paste0(out_pasta, "/", df_encontradas$novo_arquivo),
                            overwrite = F, recursive = F,
                            copy.mode = T, copy.date = T) %>%
        sum()
      df_planilha$FotosCopiadas[lin] = num_fotos_copiadas # atualiza fotos copiadas

      df_planilha$Status[lin] = ifelse(num_fotos_encontradas == num_fotos_nome_correto,
                                       ifelse(num_fotos_encontradas == num_fotos_copiadas,
                                              "1. Ok! Fotos corretas e copiadas",
                                              "3. Verificar! Problema na cópia (estranho)"),
                                       "2. Verificar! Problema nos nomes das fotos (checar numeração)")
    }, error = function(err) {
      df_planilha$Status[lin] = "4. Erro! Chamar o desenvolvedor"
    })
  }
}

View(df_planilha) # exibe o relatório
relatorio <- paste0(txt_data_hora, "-relatorio.xlsx")
lst_relatorio <-
  list("Referencias" = df_planilha, "Fotos" = df_fotos)
write_xlsx(lst_relatorio, relatorio) # cria output: relatório
