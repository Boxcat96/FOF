# フォルダ内のcsvを読み込む関数
read_folder <- function(fol_path, column_names) {
  # フォルダ内のCSVファイルを取得
  file_list <- list.files(path = fol_path, pattern = "\\.csv$"
                          , full.names = TRUE)
  
  # エラーチェック
  if (length(file_list) == 0) {
    stop("指定したフォルダにCSVファイルが見つかりません。")
  }
  
  # 読み込み処理
  combined_df <- do.call(rbind, lapply(file_list, function(file) {
    # ファイルを読み込み、列名を適用
    read_csv(file, col_names = FALSE) %>% 
      setNames(column_names)
  }))
  
  # データフレームを返す
  return(combined_df)
}