rm(list=ls()); gc(); gc(); graphics.off(); #おまじない
setwd(dirname(rstudioapi::getActiveDocumentContext()$path)) #作業フォルダ


# 設定 ----------------------------------------------------------------------
get_keihyo <- 0
# 0：QとFYの両方
# 1：Qのみ
# 2：FYのみ

master_path <- "data/master.xlsx"        # マスタファイルのファイルパス
data_path_konki <- "data/ff_value"       # 今回データのフォルダパス
data_path_zenki <- "data/ff_value_zenki" # 前期データのフォルダパス
soku <- 202404                           # 速報期をYYYYMMで指定（例：202402）
kaku <- 202403                           # 速報期をYYYYMMで指定（例：202401）
soku_FY <- 2023                          # 年度をYYYYで指定（例：2023）

# プログラムの実行 ----------------------------------------------------------------
library(tictoc)
tic()

if (get_keihyo == 0 | get_keihyo == 1){
  Q_or_FY <- "Q"
  source("fof.R")
}

if (get_keihyo == 0 | get_keihyo == 2){
  Q_or_FY <- "FY"
  source("fof.R")
}

toc()