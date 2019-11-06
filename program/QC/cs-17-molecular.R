# CS-17-Molecular_interim
# Mariko Ohtsuka
# 2019/11/5 created
# ------ Remove objects ------
rm(list=ls())
# ------ library ------
library("stringr")
library("dplyr")
library("here")
library("openxlsx")
# ------ constant ------
kCS_17 <- "CS_17"
kCS_17_Molecular <- "CS_17_Molecular"
# ------ function
#' @title ReadCsvFiles
#' @description Read the CSV file and save it in an object with the name of the CSV name with "raw_" appended.
#' @param file_path_list : Vector of strings
#' @param str_encode : File encoding string
#' @param obj_names : Specifying variable names
#' @return none
#' @example str_c(parent_path, "/input/ext/facilities.csv") %>% ReadCsvFiles("cp932", c(kCS_17, kCS_17_Molecular))
ReadCsvFiles <- function(file_path_list, str_encode, obj_names=c(kCS_17, kCS_17_Molecular)){
  for (i in 1:length(file_path_list)){
    temp_filename <- str_extract(file_path_list[i], pattern="([^/]+)(?=.csv)")
    if (str_detect(temp_filename, "CS-17_registration_[0-9]{6}_[0-9]{4}")){
      filename <- obj_names[1]
    } else if (str_detect(temp_filename, "CS-17-Molecular_registration_[0-9]{6}_[0-9]{4}")){
      filename <- obj_names[2]
    } else {
      filename <- str_c("raw_", temp_filename)
    }
    assign(filename, read.csv(file=file_path_list[i], stringsAsFactors=F, na.strings=NA, fileEncoding=str_encode),
           envir=globalenv())
    lockBinding(filename, globalenv())
  }
}
#' @title AddStyleTable2
#' @description Apply styles to worksheet cells.
#' @param output_wb : Workbook object
#' @param target_sheet_name : Sheet name string
#' @param target_row : Cell row number
#' @param border_style : Cell border string
#' @return none
#' @example AddStyleTable2(output_wb, "sheet1", 1, c("bottom", "left", "right"))
AddStyleTable2 <- function(output_wb, target_sheet_name, target_row, border_style){
  addStyle(output_wb, sheet=target_sheet_name,
           createStyle(border=border_style, borderColour="#000000", halign="center"),
           row=target_row, cols=c(1, 2, 4, 5, 9:12), gridExpand=F)
  addStyle(output_wb, sheet=target_sheet_name,
           createStyle(border=border_style, borderColour="#000000", halign="left"),
           row=target_row, cols=c(3, 6:8), gridExpand=F)
}
# ------ init ------
parent_path <- "/Volumes/Stat/Trials/JAGSE/CS-17-Molecular_interim"
output_path <- format(Sys.Date(), "%Y%m%d") %>% str_c("CS-17-Molecular定期報告_", ., ".xlsx") %>%
                here("output", "QC", .)
# read csv
file_list <- str_c(parent_path, "/input/rawdata") %>% list.files(full.names=T)
registration_list <- str_subset(file_list, pattern="([^/]+)(?=_[0-9]{6}_[0-9]{4}.csv)")
exclude_registration_list <- file_list[-which(file_list %in% registration_list)]
registration_list %>% ReadCsvFiles("cp932")
exclude_registration_list %>% ReadCsvFiles("utf-8")
str_c(parent_path, "/input/ext/diseases.csv") %>% ReadCsvFiles("utf-8")
str_c(parent_path, "/input/ext/facilities.csv") %>% ReadCsvFiles("cp932")
# ------ Table 1 ------
# 医療機関名 DM$facility_name
DM <- raw_DM
DM$str_SITEID <- formatC(DM$SITEID, width=9, flag="0")
# DM.SITEID, facilities.csv.A列をキーに facilities.csv.B列を付与
DM <- DM %>% left_join(raw_facilities, by=c("str_SITEID"="施設コード"))
# 症例数	DM.SITEID, DM.SUBJID	DM.SITEIDごとのDM.SUBJIDの合計
table1 <- summarise(group_by(DM, 施設名.ja.), number_of_cases=n())
# 合計の症例数	DM.SUBJID	DM.SUBJIDの合計を算出
table1 <- rbind(table1, c("合計", sum(table1$number_of_cases)))
# ------ Table 2 ------
# 症例番号
# CS-17-Molecular_registration_yyyymmdd_mmss.csv.症例登録番号
# CS-17_registration_yyyymmdd_mmss.csv.登録コードが存在するもののみ出力
# JALSG番号
# CS-17-Molecular_registration_yyyymmdd_mmss.csv.登録コード
# CS-17_registration_yyyymmdd_mmss.csv.登録コードが存在するもののみ出力
local({
  df_registration <- get(kCS_17) %>% select(c(症例登録番号, 登録コード))
  df_molecular <- get(kCS_17_Molecular) %>% select("登録コード") %>% right_join(df_registration, by=c("登録コード"))
  table2 <<- DM[, c("SUBJID", "USUBJID", "BRTHDTC", "SEX", "施設名.ja.")] %>%
               left_join(df_registration, by=c("SUBJID"="症例登録番号")) %>%
                   left_join(df_molecular, by="登録コード")
})
# 年齢（診断時）
# MH.MHDTC, DM.BRTHDTC
# MH.MHDTC where MHCAT=PRIMARY DIAGNOSISとDM.BRTHDTCから算出
table2 <- raw_MH %>% select(c(USUBJID, MHDTC)) %>% filter(MHDTC != "") %>% inner_join(table2, by="USUBJID")
table2$age <- apply(table2, 1, function(x){return(length(seq(as.Date(x["BRTHDTC"]), as.Date(x["MHDTC"]), "year")) - 1)})
# 性別
# DM.SEX
# "If DM.SEX=0 then M
# Else F"
table2$str_sex <- ifelse(table2$SEX == 0, "M", "F")
# WHO2016 診断病名
# MH.MHTERM, MH.MHCAT
# MH.MHTERM where MHCAT=PRIMARY DIAGNOSIS
MH <- raw_MH %>% select(c(USUBJID, MHTERM, MHCAT, MHDTC)) %>% filter(MHCAT=="PRIMARY DIAGNOSIS")
local({
  df_diseases <- raw_diseases
  df_diseases$str_code <- as.character(df_diseases$code)
  table2 <<- MH %>% select(c(USUBJID, MHTERM)) %>% right_join(table2, by="USUBJID") %>%
            left_join(df_diseases, by=c("MHTERM"="str_code"))
})
# 投与時期
# EC.EPOCH EC.ECTPT
# 治療内容
# EC.ECTRT
epoch_order <- data.frame(c(1, 2, 3),
                          c("INDUCTION THERAPY", "CONSOLIDATION THERAPY/MAINTENANCE THERAPY", "SALVAGE THERAPY"),
                          c("寛解導入療法", "地固め療法・維持療法", "サルベージ療法"),
                          stringsAsFactors=F)
colnames(epoch_order) <- c("epoch_order", "EPOCH", "epoch_ja")
table2 <- raw_EC %>% left_join(epoch_order, by="EPOCH") %>%
            select(c(USUBJID, EPOCH, ECTPT, ECTRT, epoch_order, epoch_ja, ECSPID)) %>% right_join(table2, by="USUBJID")
# 日本語ラベル
ectpt_label <- data.frame(c("COURSE 1", "COURSE 2", "COURSE 3", "COURSE 4"),
                          c("1コース", "2コース", "3コース", "4コース"), stringsAsFactors=F)
colnames(ectpt_label) <- c("ECTPT", "ectpt_ja")
table2 <- table2 %>% left_join(ectpt_label, by="ECTPT")
table2$epoch_ja <- apply(table2, 1, function(x){
  if (x["EPOCH"] == "CONSOLIDATION THERAPY/MAINTENANCE THERAPY" & !is.na(x["ectpt_ja"])){
    res = str_c(x["epoch_ja"], " ", x["ectpt_ja"])
  } else {
    res = x["epoch_ja"]
  }
  return(res)
})
# 寛解もしくは安定の有無
# RS.RSORRES, RS.EPOCH
# "If RS.RSORRES where EPOCH=INDUCTION THERAPY in (CR, CRi) then 有
# Else 無"
table2 <- raw_RS %>% select(c(USUBJID, RS_EPOCH="EPOCH", RSORRES, RSDTC, RSSPID)) %>%
            right_join(table2, by=c("USUBJID", "RS_EPOCH"="EPOCH"))
table2$remission <- ifelse((table2$RS_EPOCH == "INDUCTION THERAPY") & (table2$RSORRES == "CR" | table2$RSORRES == "CRi"),
                           "有", NA)
# 再発もしくは増悪の有無*1
# RS.RSORRES, RS.EPOCH, CE.CEDTC
# "寛解もしくは安定の有無が有について
# If CE.CEDTC^=Null 有
# Else 無"
table2 <- raw_CE %>% select(c(USUBJID, CEDTC)) %>% group_by(USUBJID) %>% filter(CEDTC == min(CEDTC)) %>%
            right_join(table2, by="USUBJID")
table2$remission <- ifelse(is.na(table2$remission), "無", "有")
table2$relapse <- ifelse(table2$remission == "無", NA, ifelse(!is.na(table2$CEDTC), "有", "無"))
# 寛解日数もしくは無増悪期間*1
# RS.RSORRES, RS.EPOCH, RS.RSDTC, CE.CEDTC
# "寛解もしくは安定の有無が有について
# 最初のCE.CEDTC-RS.RSDTC where (EPOCH=INDUCTION THERAPY and RS.RSORRES in (CR, CRi)) +1"
table2$days_of_remission <- ifelse(table2$remission == "無", NA, as.Date(table2$CEDTC) - as.Date(table2$RSDTC) + 1)
# 生存日数
# DS.DSSTDTC, DS.DSTERM, MH.MHDTC, MH.MHCAT
# "死亡症例について
# DS.DSSTDTC where DSTERM=DEATH - MH.MHDTC where MHCAT=PRIMARY DIAGNOSIS+1"
DS <- raw_DS %>% inner_join(MH, by="USUBJID") %>%
        select(c(USUBJID, DSSTDTC, DSTERM, MHDTC, MHCAT)) %>% filter(DSTERM == "DEATH")
DS$survival_days <- as.Date(DS$DSSTDTC) - as.Date(DS$MHDTC) + 1
table2 <- table2 %>% left_join(DS, by="USUBJID")
# 症例番号、JALSG番号昇順でソート
# EC.ECTPTで昇順
table2 <- table2 %>% arrange(USUBJID, 登録コード, epoch_order, ECTPT) %>% distinct()
output_table2 <- table2 %>% select(c(USUBJID, 登録コード, 施設名.ja., age, str_sex, name_en, epoch_ja, ECTRT, remission,
                                     relapse, days_of_remission, survival_days)) %>% distinct()
# ------ output ------
local({
  output_wb <- here("output", "CS-17-Molecular定期報告_YYYYMMDD.xlsx") %>% loadWorkbook(file=.)
  # table1
  table1_start_row <- 4
  table1_last_row <- nrow(table1) + table1_start_row - 1
  target_row <- table1_start_row:table1_last_row
  target_sheet_name <- "Table1"
  addStyle(output_wb, sheet=target_sheet_name,
           createStyle(border="TopBottomLeftRight", borderColour="#000000"),
           rows=target_row, cols=1, gridExpand=T)
  addStyle(output_wb, sheet=target_sheet_name,
           createStyle(border="TopBottomLeftRight", borderColour="#000000", halign="center"),
           rows=target_row, cols=2, gridExpand=T)
  writeData(output_wb, sheet=target_sheet_name, x=table1, withFilter=F, sep=",", colNames=F, startCol=1,
            startRow=table1_start_row)
  # table2
  target_sheet_name <- "Table2"
  target_col <- 1:12
  table2_start_row <- 4
  output_row <- table2_start_row - 1
  usubjid_list <- output_table2 %>% distinct(USUBJID, .keep_all=F)
  table2_x <- NULL
  for (i in 1:nrow(usubjid_list)){
    table2_each_case <- output_table2 %>% filter(USUBJID == usubjid_list[i, 1])
    # Replace repeated items such as case numbers with spaces
    if (nrow(table2_each_case) > 1){
      table2_each_case[2:nrow(table2_each_case), c(1:6, 9:12)] <- NA
    }
    table2_x <- rbind(table2_x, table2_each_case)
    # Draw ruled lines
    output_row <- output_row + 1
    if (nrow(table2_each_case) > 1){
      AddStyleTable2(output_wb, target_sheet_name, output_row, "TopLeftRight")
      if (nrow(table2_each_case) > 2){
        last_row <- nrow(table2_each_case) - 1
        for (j in 2:last_row){
          output_row <- output_row + 1
          AddStyleTable2(output_wb, target_sheet_name, output_row, "LeftRight")
        }
      }
      output_row <- output_row + 1
      AddStyleTable2(output_wb, target_sheet_name, output_row, c("bottom", "left", "right"))
    } else {
      AddStyleTable2(output_wb, target_sheet_name, output_row, "TopBottomLeftRight")
    }
  }
  writeData(output_wb, sheet=target_sheet_name, x=table2_x, withFilter=F, sep=",", colNames=F,
            startCol=1, startRow=table2_start_row)
  output_row <- output_row + 1
  comment <- data.frame(c("*1:寛解もしくは安定例の場合", "*2:死亡症例の場合"))
  writeData(output_wb, sheet=target_sheet_name, x=comment, withFilter=F, sep=",", colNames=F,
            startCol=1, startRow=output_row)
  saveWorkbook(output_wb, output_path, overwrite=T)
})
