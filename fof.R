# rm(list=ls()); gc(); gc(); graphics.off(); #おまじない
# setwd(dirname(rstudioapi::getActiveDocumentContext()$path)) #作業フォルダ

library(tidyverse)
library(readxl)
library(writexl)
library(tictoc)
source("function/read_folder.R") # フォルダ内のcsvを全て読み込む関数
source("function/format_sandan.R") # ヘッダーとNAを追加する関数（三段表作成用）
source("function/format_matrix.R") # ヘッダーを追加する関数（マトリクス作成用）
source("function/add_na_rows.R") # あるデータフレームの後に、指定された行までNAを挿入する関数

tic() # 時間計測開始

# 設定 ----------------------------------------------------------------------
# マスターファイルの読み込み
# setwd("C:/Users/tkero/OneDrive/経済ファイル/02_分析/93_FOF計表")

# マスタファイルのパスを指定
master_path <- "data/master.xlsx"

# ファイル内の全シート名を取得
sheet_names <- excel_sheets(master_path)

# すべてのシートをリストに読み込む
master_data <- lapply(sheet_names, function(sheet) {
  read_excel(master_path, sheet = sheet)
})

# リストの名前をシート名に設定
names(master_data) <- sheet_names

# フォルダパスと時期を指定
master_info <- master_data[["master"]]

folder_path <- as.character(master_info[1, 1]) # このファイルのパス
data_path_konki <- as.character(master_info[2, 1]) # 今回データのパス
data_path_zenki <- as.character(master_info[3, 1]) # 前期データのパス
soku <- as.numeric(master_info[4, 1])          # 速報期
kaku <- as.numeric(master_info[5, 1])          # 確報期

# 部門情報を格納
sector <- master_data[["部門"]] %>% 
  mutate(sec = as.character(sec))

# 項目情報を格納
item <- master_data[["項目"]] %>% 
  mutate(item = as.character(item))

# 初期データフレームの列名の指定
column_names = c("code", "QFY", "period", "value")

# 計表を読み込み、QFYをフィルターする関数
read_folder_keihyo <- function(fol_path, column_names, freq){
  read_folder(fol_path, column_names) %>%
    filter(QFY == freq) %>% 
    select(-QFY)
}

# 期種・期初指定（基準改定まで更新不要）
# Q_or_FY <- "Q"
if (Q_or_FY == "Q"){
  start_period <- 200501
}else if (Q_or_FY == "FY"){
  start_period <- 2004
}else{
  stop("期種を正しく指定してください")
}

# フォルダパス
result_folder <- str_c("result/", Q_or_FY)

# 今期フォルダの読み込み
df_raw <- read_folder_keihyo(data_path_konki, column_names, Q_or_FY)

# 今期フォルダの読み込み
df_raw_zenki <- read_folder_keihyo(data_path_zenki, column_names, Q_or_FY)
# value列をvalue_zenkiに変更

# 三段表の開始行を指定
start_row = 5
  
# 並び替え順を指定
FSR_order <- c("F","S","R")
AL_order <- c("A","L")
sec_order <- unlist(sector$sec)
item_order <- unlist(item$item)

# データセットの成形 -----------------------------------------------------------------

df <- df_raw %>% 
  # 今期データに、前期データをbind_rowsし、今期データがなければ前期データを採用
  mutate(FSR = substr(code, 8, 8),          # 左から8文字目、FSR
         sec = substr(code, 9, 11),         # 9文字目から11文字目、部門
         AL  = substr(code, 12, 12),        # 12文字目、資産／負債
         item = substr(code, 13, 15)        # 13文字目から15文字目、項目
  ) %>%
  filter(period >= start_period 
         & (FSR == "F" | FSR == "S" | FSR == "R")
         & (AL == "A" | AL == "L")
         & !(item == "750" | item == "751"))  %>% # 季調済系列を削除
  complete(period, FSR, AL, sec, item) %>% 
  filter(!((FSR == "F" | FSR == "R" ) & period <= start_period)) %>%  # 不要データの削除
  left_join(sector, by = "sec")  %>% 
  left_join(item, by = "item")  %>% 
  left_join(df_raw_zenki, by = c("code", "period")) %>% 
  rename(value = value.x,
         value_zenki = value.y) %>% 
  select(-code) %>% 
  mutate(
    # 今期コードが追加されたときも、（前期を0にして）前期差が表示されるようにする
    value_zenkisa =
      value - if_else(is.na(value_zenki), 0, value_zenki),
    FSR = factor(FSR, levels = FSR_order),    # FSRの順番設定
    AL = factor(AL, levels = AL_order),       # ALの順番設定
    sec = factor(sec, levels = sec_order),    # 部門の順番設定
    item = factor(item, levels = item_order)  # 項目の順番設定
  ) 

# 以後使用しないデータフレームを変数リストから削除
# rm(df_raw, df_raw_zenki)

# Excel出力データ作成関数 ----------------------------------------------------
prepare_excel <- function(df, column_name, style) {

  # フィルタリング基準となる列を取得
  categories <- unique(df[[column_name]])

  # フィルタリング結果をリストに保存
  output <- lapply(categories, function(cat) {
    df %>%
      filter(!!sym(column_name) == cat) %>%  # column_nameを動的に参照
      { 
        if (style == "sandan") {
          format_sandan(., header_sandan,
                        start_row, start_row + 200, start_row + 400)  # 三段表の場合
        } 
        else if (style == "matrix") {
          format_matrix(., header_matrix)  # マトリクスの場合
        }
        else{identity(.)}
      } %>% 
      select(-all_of(column_name))     # column_name列を削除
  })

  # リストを名前付きリストに変換（シート名に使用）
  names(output) <- categories

  return(output)
}


# 三段表ヘッダー関数の作成 ---------------------------------------------------------------
create_header_sandan <- function(df, df_header, header, header_name) {
  # 処理の実行
  enframe(colnames(df), name = NULL, value = "temp") %>%  # colnamesを簡潔に変換
    separate(temp, into = c("prefix", header_name), sep = "_", fill = "right") %>%
    left_join(df_header, by = header_name) %>%
    mutate(
      temp_header = as.numeric(as.character(.data[[header]])),  # 動的参照を簡略化
      # l_header（大部門）, m_header（中部門）, s_header（小部門）
      l_header = if_else(temp_header %% 100 == 0, str_trim(.data[[header_name]], side = "left"), NA),
      m_header = if_else(temp_header %% 100 != 0 & temp_header %% 10 == 0, str_trim(.data[[header_name]], side = "left"), NA),
      s_header = if_else(temp_header %% 10 != 0, str_trim(.data[[header_name]], side = "left"), NA)
    ) %>%
    select(-all_of(c(header_name, "temp_header", header))) %>%  # 不要な列を一括削除
    t() %>%  # 転置
    as.data.frame()
}

# 計表の作成 -------------------------------------------------------------------

# 部門別三段表
df_bumon <- df %>%
  select(-value_zenki, -value_zenkisa) %>% 
  arrange(AL, item, sec, FSR, period) %>% 
  select(-sec, -item) %>%
  mutate(
    sector_name = str_trim(sector_name, side = "left"),
    FSR = as.character(FSR) # FSRがfactor型であることによるエラーへの対応
  ) %>% # 全角スペースの削除
  pivot_wider(names_from = c(AL, item_name), values_from = value) 

# ヘッダーの指定 -> フィルター&シート分け -> Excel出力
header_sandan <- create_header_sandan(df_bumon, item, "item", "item_name")
output <- prepare_excel(df_bumon, "sector_name", "sandan")
write_xlsx(output, path = str_c(result_folder, "/1_部門別_三段表.xlsx"),
           col_names = FALSE)

# 項目別三段表
df_koumoku <- df %>%
  select(-value_zenki, -value_zenkisa) %>%
  arrange(AL, sec, item, FSR, period) %>% 
  select(-sec, -item) %>% 
  mutate(
    item_name = str_trim(item_name, side = "left"),
    FSR = as.character(FSR) # FSRがfactor型であることによるエラーへの対応
  ) %>% # 全角スペースの削除
  pivot_wider(names_from = c(AL, sector_name), values_from = value) 

# ヘッダーの指定 -> フィルター&シート分け -> Excel出力
header_sandan <- create_header_sandan(df_koumoku, sector, "sec", "sector_name")
output <- prepare_excel(df_koumoku, "item_name", "sandan")
write_xlsx(output, path = str_c(result_folder, "/2_項目別_三段表.xlsx"),
           col_names = FALSE)


# Q計表の場合は続行（mark1）
if (Q_or_FY == "Q"){
  
  # マトリクス用データフレーム（項目を大項目・中項目・小項目にソート）
  df_matrix_raw <- df %>% 
    filter(period == soku | period == kaku) %>%
    arrange(period, sec, item, FSR, AL) %>%
    mutate(
      item = as.numeric(as.character(item)),  
      # l_item（大項目）, m_item（中項目）, s_item（小項目）
      l_item = if_else(item %% 100 == 0, str_trim(item_name, side = "left"), NA),
      m_item = if_else(item %% 100 != 0 & item %% 10 == 0, str_trim(item_name, side = "left"), NA),
      s_item = if_else(item %% 10 != 0, str_trim(item_name, side = "left"), NA)
    ) %>% 
    select(-item_name)

  # マトリクスヘッダー関数の作成
  create_header_matrix <- function(df, df_header, header, header_name) {
    # 処理の実行
    enframe(colnames(df), name = NULL, value = "temp") %>%  # colnamesを簡潔に変換
      separate(temp, into = c(header_name, "prefix"), sep = "_", fill = "right") %>%
      left_join(df_header, by = header_name) %>%
      mutate(
        temp_header = as.numeric(as.character(.data[[header]])),  # 動的参照を簡略化
        # l_header（大部門）, m_header（中部門）, s_header（小部門）
        l_header = if_else(temp_header %% 100 == 0, str_trim(.data[[header_name]], side = "left"), NA),
        m_header = if_else(temp_header %% 100 != 0 & temp_header %% 10 == 0, str_trim(.data[[header_name]], side = "left"), NA),
        s_header = if_else(temp_header %% 10 != 0, str_trim(.data[[header_name]], side = "left"), NA)
      ) %>% 
      select(-all_of(c(header_name, "temp_header", header))) %>%  # 不要な列を一括削除
      select(setdiff(names(.), "prefix"), "prefix") %>%  # "prefix"列を最後尾に移動
      t() %>%  # 転置
      as.data.frame()
  }

  # 速確マトリクス
  df_matrix <- df_matrix_raw %>%
    select(-value_zenki, -value_zenkisa) %>%
    mutate(period = if_else(period == soku, "速報", "確報")) %>%
    select(-sec, -item) %>% 
    mutate(FSR = as.character(FSR)) %>% # FSRがfactor型であることによるエラーへの対応)
    pivot_wider(names_from = c(sector_name, AL), 
                values_from = value,
                names_sep = "_") %>%
    # シート名用に速確＋FSRの列を作成（「確報F」など）
    mutate(period = str_c(period, FSR)) %>% 
    select(-FSR)
  
  # ヘッダーの指定 -> フィルター&シート分け -> Excel出力
  header_matrix <- create_header_matrix(df_matrix, sector, "sec", "sector_name")
  output <- prepare_excel(df_matrix, "period", "matrix")
  write_xlsx(output, path = str_c(result_folder, "/3_速確マトリクス.xlsx"),
             col_names = FALSE)
  
  # 速確乖離
  df_kairi <- df_matrix_raw %>% 
    select(-value, -value_zenki) %>% 
    filter(period == kaku) %>% 
    select(-period, -sec, -item) %>% 
    mutate(FSR = as.character(FSR)) %>% # FSRがfactor型であることによるエラーへの対応)
    pivot_wider(names_from = c(sector_name, AL),
                values_from = value_zenkisa,
                names_sep = "_")
  
  # ヘッダーの指定 -> フィルター&シート分け -> Excel出力
  header_matrix <- create_header_matrix(df_kairi, sector, "sec", "sector_name")
  output <- prepare_excel(df_kairi, "FSR", "matrix")
  write_xlsx(output, path = str_c(result_folder, "/4_速確乖離.xlsx"),
             col_names = FALSE)
  
  # FSRグラフ
  df_graph <- df %>% 
    select(-value_zenki, -value_zenkisa) %>%
    arrange(period, sec, AL, item) %>% 
    select(-sec, -item) %>% 
    filter(!is.na(value)) %>% 
    pivot_wider(names_from = period, values_from = value) %>% 
    arrange(FSR) # シートの順番を整えるため
  
  output <- prepare_excel(df_graph, "FSR", "else")
  write_xlsx(output, path = str_c(result_folder, "/5_FSRグラフ.xlsx"))

} #（mark1）

toc() # 時間計測終了
