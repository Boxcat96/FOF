# あるデータフレームの後に、指定された行までNAを挿入する関数
add_na_rows <- function(df, header, end_row) {
  
  # end_rowが小さすぎる場合はエラーを出す
  if (end_row - nrow(df) < nrow(header)) {
    stop("開始行の値が小さすぎます。修正してください。")
  }
  
  # 指定した数のNA行を追加
  na_rows <- matrix(NA, nrow = (end_row - nrow(df)), ncol = ncol(df))
  colnames(na_rows) <- colnames(df)  # 列名を元のデータフレームと一致させる
  df <- rbind(df, na_rows)
  
  return(df)
}