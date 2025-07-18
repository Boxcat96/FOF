# フォルダ内のcsvを読み込む関数
read_folder <- function(fol_path, column_names) {
  # library(tidyverse)
  
  # フォルダ内のCSVファイルを取得
  file_list <- list.files(path = fol_path, pattern = "\\.csv$"
                          , full.names = TRUE)
  
  # エラーチェック
  if (length(file_list) == 0) {
    stop("指定したフォルダにCSVファイルが見つかりません。")
  }
  
  # 読み込み処理
  combined_df <- file_list %>%
    map(~ read_csv(.x, col_names = FALSE) %>% 
          setNames(column_names)) %>% 
    bind_rows()
  
  # データフレームを返す
  return(combined_df)
}