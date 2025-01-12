# ヘッダーを追加する関数（マトリクス作成用）
format_matrix <- function(df, header) {
  
  # source("function/add_na_rows.R")
  # library(tidyverse)
  
  # 空のデータフレームを用意
  df_empty <- df %>% slice(0)
  
  # NA行追加
  new_df <- add_na_rows(df_empty, header, nrow(header))
  new_df <- bind_rows(new_df, df)
  
  # ヘッダーの代入
  new_df[1:nrow(header), ] <- header
  
  return(new_df)
}