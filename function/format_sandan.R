# ヘッダーとNAを追加する関数（三段表作成用）
format_sandan <- function(df, header, start_F, start_S, start_R) {
  
  # source("function/add_na_rows.R")
  # library(tidyverse)
  
  # 空のデータフレームを用意したあと、F, S, Rの部分を順番に処理
  df_empty <- df %>% slice(0)
  df_F <- df %>% filter(FSR == "F")
  df_S <- df %>% filter(FSR == "S")
  df_R <- df %>% filter(FSR == "R")
  
  # NA行追加 -> F -> NA行追加 -> S -> NA行追加 -> R の順に結合
  new_df <- df_empty %>% 
    add_na_rows(header, start_F-1) %>% 
    bind_rows(df_F) %>% 
    add_na_rows(header, start_S-1) %>% 
    bind_rows(df_S) %>% 
    add_na_rows(header, start_R-1) %>% 
    bind_rows(df_R) 
  
  # F、S、Rの前にそれぞれヘッダーを代入する
  new_df[(start_F-nrow(header)):(start_F-1), ] <- header
  new_df[(start_S-nrow(header)):(start_S-1), ] <- header
  new_df[(start_R-nrow(header)):(start_R-1), ] <- header
  
  return(new_df)
}