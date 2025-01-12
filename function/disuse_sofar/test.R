rm(list = ls())
library(tidyverse)
library(tictoc)

tic()

# サンプルデータフレーム
df <- data.frame(
  FSR = c("F", "F", "S", "S", "S", "R", "R", "R"),
  Value = c(1, 2, 3, 4, 5, 6, 7, 8),
  stringsAsFactors = FALSE
)

df_column <- as.data.frame(colnames(df)) %>% 
  rename(original = "colnames(df)") %>% 
  mutate(copy = str_c("c_",original))

df_trp <- as.data.frame(t(df_column))


# 3行のデータをデータフレームとして作成（例として）
header <- data.frame(
  FSR = c("Header1", "Header2", "Header3"),
  Value = c("Subheader1", "Subheader2", "Subheader3"),
  stringsAsFactors = FALSE
)


# ヘッダーとNAを追加する関数（三段表作成用）
format_sandan <- function(df, header, start_F, start_S, start_R) {
  
  source("function/add_na_rows.R")
  
  # データフレームとヘッダーの列数が異なる場合にエラーを出す
  if (!(ncol(df) == ncol(header))){
    stop("データフレームとヘッダーの列数が異なります。修正してください。")
  }
  
  # 空のデータフレームを用意したあと、F, S, Rの部分を順番に処理
  df_empty <- df %>% slice(0)
  df_F <- df %>% filter(FSR == "F")
  df_S <- df %>% filter(FSR == "S")
  df_R <- df %>% filter(FSR == "R")
  
  # NA行追加 -> F -> NA行追加 -> S -> NA行追加 -> R の順に結合
  # start_F, start_S, start_Rの1行前までNAを付け加える
  new_df <- add_na_rows(df_empty, start_F-1)
  new_df <- bind_rows(new_df, df_F)
  new_df <- add_na_rows(new_df, start_S-1)
  new_df <- bind_rows(new_df, df_S)
  new_df <- add_na_rows(new_df, start_R-1)
  new_df <- bind_rows(new_df, df_R)  
  
  # F、S、Rの前にそれぞれヘッダーを代入する
  new_df[(start_F-nrow(header)):(start_F-1), ] <- header
  new_df[(start_S-nrow(header)):(start_S-1), ] <- header
  new_df[(start_R-nrow(header)):(start_R-1), ] <- header
  
  return(new_df)
}

df <- format_sandan(df, header, 4, 15, 25)
# write_xlsx(df, "output.xlsx", col_names = FALSE)

toc()
