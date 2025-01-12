# 四半期データを生成する関数
gen_Q <- function(start, end) {
  start_year <- as.integer(substr(start, 1, 4))
  start_qtr <- as.integer(substr(start, 5, 6))
  end_year <- as.integer(substr(end, 1, 4))
  end_qtr <- as.integer(substr(end, 5, 6))
  
  # 全ての四半期を生成
  quarters <- c()
  for (year in start_year:end_year) {
    for (qtr in 1:4) {
      if ((year == start_year && qtr < start_qtr) || (year == end_year && qtr > end_qtr)) {
        next
      }
      quarters <- c(quarters, year * 100 + qtr)
    }
  }
  
  return(quarters)
}
