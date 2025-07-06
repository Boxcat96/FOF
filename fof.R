# rm(list=ls()); gc(); gc(); graphics.off(); #おまじない
# setwd(dirname(rstudioapi::getActiveDocumentContext()$path)) #作業フォルダ

library(tidyverse)
library(readxl)
library(writexl)
# library(tictoc)
source("function/read_folder.R") # フォルダ内のcsvを全て読み込む関数
source("function/format_sandan.R") # ヘッダーとNAを追加する関数（三段表作成用）
source("function/format_matrix.R") # ヘッダーを追加する関数（マトリクス作成用）
source("function/add_na_rows.R") # あるデータフレームの後に、指定された行までNAを挿入する関数

# tic() # 時間計測開始

# 設定 ----------------------------------------------------------------------

# 期種・期初指定（基準改定まで更新不要）
# Q_or_FY <- "Q"
if (Q_or_FY == "Q"){
  start_period <- 200501
}else if (Q_or_FY == "FY"){
  start_period <- 2004
}else{
  stop("期種を正しく指定してください")
}

# 結果格納フォルダパス
result_folder <- str_c(result_path, Q_or_FY)

# 三段表の開始行、FSR間の行数を指定
start_row <- 5
sandan_interval <- 200

# マスタ・データ読み込み ---------------------------------------------------------------
# マスタファイル内の全シート名を取得
sheet_names <- excel_sheets(master_path)

# すべてのシートをリストに読み込む
master_data <- lapply(sheet_names, function(sheet) {
  read_excel(master_path, sheet = sheet)
})

# リストの名前をシート名に設定
names(master_data) <- sheet_names

# 部門情報を格納
sector <- master_data[["部門"]] %>% 
  mutate(sec = as.character(sec))

# 項目情報を格納
item <- master_data[["項目"]] %>% 
  mutate(item = as.character(item))

# 空白トリム系列（部門・項目）
sector_trmd <- sector %>% 
  mutate(sector_name = str_trim(sector_name, side = "left"))
item_trmd <- item %>% 
  mutate(item_name = str_trim(item_name, side = "left"))

# 初期データフレームの列名の指定
column_names <- c("code", "QFY", "period", "value")

# 計表を読み込み、QFYをフィルターする関数
read_folder_keihyo <- function(fol_path, column_names, freq){
  read_folder(fol_path, column_names) %>%
    filter(QFY == freq) %>% 
    select(-QFY)
}

# 今期フォルダの読み込み
df_raw_konki <- read_folder_keihyo(data_path_konki, column_names, Q_or_FY)

# 前期フォルダの読み込み
df_raw_zenki <- read_folder_keihyo(data_path_zenki, column_names, Q_or_FY) 

# 前回フォルダの読み込み
df_raw_maekai <- read_folder_keihyo(data_path_maekai, column_names, Q_or_FY) 

# 前期フォルダと今期フォルダを結合（FOF作業用）
# 例えば、今期フォルダに速確期しか入れないようなケースを想定

# 今回計算
df_raw <- df_raw_zenki %>%
  filter(!period %in% df_raw_konki$period) %>% 
  bind_rows(df_raw_konki)

# 前回計算
df_maekai <- df_raw_zenki %>%
  filter(!period %in% df_raw_maekai$period) %>% 
  bind_rows(df_raw_maekai) %>% 
  # value列をvalue_maekaiに変更（後でjoinする際に重複しないように）
  rename(value_maekai = value)

# value列をvalue_zenkiに変更（後でjoinする際に重複しないように）
df_raw_zenki <- df_raw_zenki %>% 
  rename(value_zenki = value)
  
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
  complete(period, FSR, AL, sec, item) %>%        # 暗黙の欠損値を補完
  filter(!((FSR == "F" | FSR == "R" ) & period <= start_period)) %>%  # FRの冒頭期を削除
  left_join(sector, by = "sec")  %>% # 部門マスタの結合
  left_join(item, by = "item")  %>%  # 項目マスタの結合
  left_join(df_raw_zenki, by = c("code", "period")) %>% # 前期フォルダの結合 
  left_join(df_maekai, by = c("code", "period")) %>%    # 前回計算フォルダの結合
  select(-code) %>% # コードは以後不要なので削除
  mutate(
    # 今期コードが追加されたときも、（前期を0にして）前期差が表示されるようにする
    value_zenkisa =
      value - if_else(is.na(value_zenki), 0, value_zenki),
    value_maekaisa =
      value - if_else(is.na(value_maekai), 0, value_maekai),
    FSR = factor(FSR, levels = FSR_order),    # FSRの順番設定
    AL = factor(AL, levels = AL_order),       # ALの順番設定
    sec = factor(sec, levels = sec_order),    # 部門の順番設定
    item = factor(item, levels = item_order)  # 項目の順番設定
  ) 

# マイナス残高リストの作成
write_xlsx(
  list(
  "マイナス残高" = df %>% 
    filter(FSR == "S" & !(item == 700) & value < 0) %>% 
    select(period, AL, sec, item, sector_name, item_name, value)
  ), path = str_c(result_folder, "/0_マイナス残高.xlsx"))

# 以後使用しないデータフレームを変数リストから削除
# rm(df_raw, df_raw_zenki, df_raw_konki, df_raw_maekai, df_maekai)

# 三段表データフレーム作成関数 ----------------------------------------------------------
make_sandan_table <- function(df,
                              arrange_vars,     # データ並び替えの順序
                              names_from_vars,  # 列として展開したい系列。例：c("AL", "item_name")
                              values_target) {  # 参照する値。例：今期計数ならvalue
  df %>%
    select(c("period", "FSR", "AL", "sec", "item", "sector_name", "item_name",
             all_of(values_target))) %>% 
    arrange(across(all_of(arrange_vars))) %>%
    select(-sec, -item) %>%
    mutate(
      sector_name = str_trim(sector_name, side = "left"),
      item_name   = str_trim(item_name, side = "left"),
      FSR = as.character(FSR)
    ) %>%
    pivot_wider(
      names_from = all_of(names_from_vars),
      values_from = all_of(values_target)
    )
}

# 三段表ヘッダー関数の作成 ---------------------------------------------------------------
create_header_sandan <- function(df,
                                 df_header, # トリム後の部門／項目マスタ（sector_trmd / item_trmd）
                                 header,    # dfにおける部門／項目番号の列名（"sector" / "item"）
                                 header_name) { # dfにおける部門名／項目名の列名（"sector_name" / "item_name"）
  
  enframe(colnames(df), name = NULL, value = "temp") %>%  # colnamesを簡潔に変換
    separate(temp, into = c("prefix", header_name), sep = "_", fill = "right") %>%
    left_join(df_header, by = header_name) %>%
    mutate(
      temp_header = as.numeric(as.character(.data[[header]])),  # 動的参照を簡略化
      # l_header（大部門／項目）, m_header（中部門／項目）, s_header（小部門／項目）
      l_header = if_else(temp_header %% 100 == 0, .data[[header_name]], NA),
      m_header = if_else(temp_header %% 100 != 0 & temp_header %% 10 == 0, .data[[header_name]], NA),
      s_header = if_else(temp_header %% 10 != 0, .data[[header_name]], NA)
    ) %>%
    select(-all_of(c(header_name, "temp_header", header))) %>%  # 不要な列を一括削除
    t() %>%  # 転置
    as.data.frame()
}


# Excel出力データ作成関数 ----------------------------------------------------
prepare_excel <- function(df,
                          column_name, # シート名（sector_name, item_name, periodなど）
                          style,       # sandanかmatrixか""で指定（処理が異なる）
                          header_keihyo) { # 各フォーマットのヘッダー

  # フィルタリング基準となる列を取得
  categories <- unique(df[[column_name]])

  # フィルタリング結果をリストに保存
  output <- lapply(categories, function(cat) { # categories内の全要素に対して以下の作業を実施
    df %>%
      filter(!!sym(column_name) == cat) %>%  # column_nameを動的に参照し、各要素でフィルター
      { 
        if (style == "sandan") {
          format_sandan(., header_keihyo,
                        start_row, start_row + sandan_interval, start_row + sandan_interval*2)  # 三段表の場合
        } 
        else if (style == "matrix") {
          format_matrix(., header_keihyo)  # マトリクスの場合
        }
        else{identity(.)}
      } %>% 
      select(-all_of(column_name))     # column_name列を削除
  })

  # リストを名前付きリストに変換（シート名に使用）
  names(output) <- categories

  return(output)
}


# 計表の作成 -------------------------------------------------------------------

# 部門別三段表用データ
df_bumon <- df %>%
  make_sandan_table(c("AL", "item", "sec", "FSR", "period"),
                    c("AL", "item_name"), "value")

# ヘッダーの指定 -> フィルター&シート分け -> Excel出力
header_sandan_bumon <- create_header_sandan(df_bumon, item_trmd, "item", "item_name") 
prepare_excel(df_bumon, "sector_name", "sandan", header_sandan_bumon) %>% 
  write_xlsx(path = str_c(result_folder, "/1_部門別_三段表.xlsx"), col_names = FALSE)

# 項目別三段表
df_koumoku <- df %>%
  make_sandan_table(c("AL", "sec", "item", "FSR", "period"),
                    c("AL", "sector_name"), "value")

# ヘッダーの指定 -> フィルター&シート分け -> Excel出力
header_sandan_koumoku <- create_header_sandan(df_koumoku, sector_trmd, "sec", "sector_name")
prepare_excel(df_koumoku, "item_name", "sandan", header_sandan_koumoku) %>% 
  write_xlsx(path = str_c(result_folder, "/2_項目別_三段表.xlsx"), col_names = FALSE)


# マトリクス用データフレーム（項目を大項目・中項目・小項目にソート）
df_matrix_raw <- df %>% 
  filter(period == soku | period == kaku | period == soku_FY) %>%
  arrange(period, sec, item, FSR, AL) %>%
  mutate(
    item = as.numeric(as.character(item)),  
    # l_item（大項目）, m_item（中項目）, s_item（小項目）
    l_item = if_else(item %% 100 == 0, str_trim(item_name, side = "left"), NA),
    m_item = if_else(item %% 100 != 0 & item %% 10 == 0, str_trim(item_name, side = "left"), NA),
    s_item = if_else(item %% 10 != 0, str_trim(item_name, side = "left"), NA)
  ) %>% 
  select(-item_name)


# マトリクスヘッダー関数の作成 ----------------------------------------------------------
create_header_matrix <- function(df,
                                 df_header, # 部門マスタ(sector)
                                 header,    # dfにおける部門番号の列名("sec")
                                 header_name) { # dfにおける部門名の列名("sector_name")
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
# -----------------------------------------------------------------------------

# 速確マトリクス
df_matrix <- df_matrix_raw %>%
  select(-value_zenki, -value_zenkisa, -value_maekai, -value_maekaisa) %>%
  mutate(period = case_when(
    period == soku ~ "速報",
    period == kaku ~ "確報",
    period == soku_FY ~ "年度"
  )) %>% 
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
prepare_excel(df_matrix, "period", "matrix", header_matrix) %>% 
  write_xlsx(path = str_c(result_folder, "/3_速確マトリクス.xlsx"),
           col_names = FALSE)

# Q計表の場合は続行
if (Q_or_FY == "Q"){  
  
  # 調整額リスト（あとで直す！！！！！！！！！！！！！！！！！！！！！）
  df_soku_R <- df %>% 
    filter(period == soku & FSR == "R") %>% 
    select(period, FSR, AL, sec, item, sector_name, item_name, value)
  
  write_xlsx(
    list(
      "貸出" = df_soku_R %>% filter(item %in% c(200, 210)),
      "債券" = df_soku_R %>% filter(item %in% c(300, 310))
    ),
    path = str_c(result_folder, "/0_調整額リスト.xlsx")
  )
  
  # 部門別三段表_前回公表からの乖離
  df %>%
    make_sandan_table(c("AL", "item", "sec", "FSR", "period"),
                      c("AL", "item_name"), "value_zenkisa") %>% 
    prepare_excel("sector_name", "sandan", header_sandan_bumon) %>% 
    write_xlsx(path = str_c(result_folder, "/1_部門別_乖離表_前回公表.xlsx"), col_names = FALSE)
  
  # 部門別三段表_前回計算からの乖離
  df %>%
    make_sandan_table(c("AL", "item", "sec", "FSR", "period"),
                      c("AL", "item_name"), "value_maekaisa") %>% 
    prepare_excel("sector_name", "sandan", header_sandan_bumon) %>% 
    write_xlsx(path = str_c(result_folder, "/1_部門別_乖離表_前回計算.xlsx"), col_names = FALSE)
  
  # 項目別三段表_前回公表からの乖離
  df %>%
    make_sandan_table(c("AL", "sec", "item", "FSR", "period"),
                      c("AL", "sector_name"), "value_zenkisa") %>% 
    prepare_excel("item_name", "sandan", header_sandan_koumoku) %>% 
    write_xlsx(path = str_c(result_folder, "/2_項目別_乖離表_前回公表.xlsx"), col_names = FALSE)
  
  # 項目別三段表_前回計算からの乖離
  df %>%
    make_sandan_table(c("AL", "sec", "item", "FSR", "period"),
                      c("AL", "sector_name"), "value_maekaisa") %>% 
    prepare_excel("item_name", "sandan", header_sandan_koumoku) %>% 
    write_xlsx(path = str_c(result_folder, "/2_項目別_乖離表_前回計算.xlsx"), col_names = FALSE)
    
  # 速確乖離
  df_matrix_raw %>% 
    select(-value, -value_zenki, -value_maekai, -value_maekaisa) %>% 
    filter(period == kaku) %>% 
    select(-period, -sec, -item) %>% 
    mutate(FSR = as.character(FSR)) %>% # FSRがfactor型であることによるエラーへの対応)
    pivot_wider(names_from = c(sector_name, AL),
                values_from = value_zenkisa,
                names_sep = "_") %>% 
    prepare_excel("FSR", "matrix", header_matrix) %>% # フィルター&シート分け -> Excel出力 
    write_xlsx(path = str_c(result_folder, "/4_速確乖離.xlsx"),
             col_names = FALSE)
  
  # 速確乖離_前回計算
  df_matrix_raw %>% 
    select(-value, -value_zenki, -value_zenkisa, -value_maekai) %>% 
    filter(period == soku) %>% 
    select(-period, -sec, -item) %>% 
    mutate(FSR = as.character(FSR)) %>% # FSRがfactor型であることによるエラーへの対応)
    pivot_wider(names_from = c(sector_name, AL),
                values_from = value_maekaisa,
                names_sep = "_") %>% 
    prepare_excel("FSR", "matrix", header_matrix) %>% # フィルター&シート分け -> Excel出力 
    write_xlsx(path = str_c(result_folder, "/4_速報マトリクス_前回計算乖離.xlsx"),
               col_names = FALSE)


# FSRグラフ ------------------------------------------------------------------
  # データフレーム作成
  df_graph_raw <- df %>%
    select(period, FSR, AL, sec, item, sector_name, item_name, value, value_maekai) %>% 
    arrange(period) %>%
    filter(!is.na(value)) 
  
  # 今回計算データに前回計算データを縦結合
  bind_rows(
    df_graph_raw %>% 
      select(-value_maekai) %>% 
      pivot_wider(names_from = period, values_from = value) %>%
      arrange(FSR, sec, AL, item) %>% 
      mutate(FSR = str_c("今回", FSR)), # シート名用に"今回"とFSRを結合（prepare_excelで使用）
    df_graph_raw %>%
      select(-value) %>% 
      pivot_wider(names_from = period, values_from = value_maekai) %>%
      arrange(FSR, sec, AL, item) %>% 
      mutate(FSR = str_c("前回", FSR)) # シート名用に"前回"とFSRを結合（prepare_excelで使用）
  ) %>% 
    select(-sec, -item) %>% 
    prepare_excel("FSR", "", "") %>% 
    write_xlsx(path = str_c(result_folder, "/5_FSRグラフ.xlsx"))
}

# toc() # 時間計測終了
