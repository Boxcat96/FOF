rm(list=ls()) # 変数のクリア
graphics.off() # 図表のクリア

get_keihyo <- 0
# 0：QとFYの両方
# 1：Qのみ
# 2：FYのみ

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