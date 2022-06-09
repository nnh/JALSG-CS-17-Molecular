# CS-17-Molecular_interim
# Mariko Ohtsuka
# 2019/11/5 created
# 2022/6/2 fixed
# ------ Remove objects ------
rm(list=ls())
# ------ library ------
library("stringr")
library("dplyr")
library("here")
output_path <- format(Sys.Date(), "%Y%m%d") %>% str_c("CS-17-Molecular定期報告_", ., ".xlsx") %>%
  here("output", "QC", .)
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
parent_path <- "~/Library/CloudStorage/Box-Box/Stat/Trials/JAGSE/CS-17-Molecular_interim"
# read csv
file_list <- str_c(parent_path, "/input/rawdata") %>% list.files(full.names=T, pattern=".csv") # Exclude directories
registration_list <- str_subset(file_list, pattern="([^/]+)(?=_[0-9]{6}_[0-9]{4}.csv)")
exclude_registration_list <- file_list[-which(file_list %in% registration_list)]
registration_list %>% ReadCsvFiles("utf-8")
exclude_registration_list %>% ReadCsvFiles("utf-8")
str_c(parent_path, "/input/ext/diseases.csv") %>% ReadCsvFiles("utf-8")
str_c(parent_path, "/input/ext/facilities.csv") %>% ReadCsvFiles("utf-8")
# ------ common processing ------
local({
  target_cs_17 <- get(kCS_17) %>% select(c(cs_17_subjid=症例登録番号, 登録コード))
  target_molecular <- get(kCS_17_Molecular) %>% select(c(molecular_subjid=症例登録番号, 登録コード, 施設名.ja.=シート作成時施設名))
  # Combine CS-17 and molecular with registration code
  molecular_registration <- target_molecular %>% inner_join(target_cs_17, by="登録コード")
  DM <- raw_DM
  DM <<- DM[, c("SUBJID", "USUBJID", "BRTHDTC", "SEX")] %>%
    inner_join(molecular_registration, by=c("SUBJID"="cs_17_subjid"))
})
# ------ Table 1 ------
# Sort by descending number of cases, ascending order of facility name
table1 <- summarise(group_by(DM, 施設名.ja.), number_of_cases=n()) %>%
            arrange(desc(number_of_cases), 施設名.ja.)
table1 <- rbind(table1, c("合計", sum(as.numeric(table1$number_of_cases))))
# ------ Table 2 ------
table2 <- raw_MH %>% select(c(USUBJID, MHDTC)) %>% filter(MHDTC != "") %>% inner_join(DM, by="USUBJID")
table2$age <- apply(table2, 1, function(x){return(length(seq(as.Date(x["BRTHDTC"]), as.Date(x["MHDTC"]), "year")) - 1)})
table2$str_sex <- table2$SEX
MH <- raw_MH %>% select(c(USUBJID, MHTERM, MHCAT, MHDTC)) %>% filter(MHCAT=="PRIMARY DIAGNOSIS")
local({
  df_diseases <- raw_diseases
  df_diseases$str_code <- as.character(df_diseases$code)
  table2 <<- MH %>% select(c(USUBJID, MHTERM)) %>% right_join(table2, by="USUBJID") %>%
              left_join(df_diseases, by=c("MHTERM"="str_code"))
})
epoch_order <- data.frame(c(1, 2, 3),
                          c("INDUCTION THERAPY", "CONSOLIDATION THERAPY/MAINTENANCE THERAPY", "SALVAGE THERAPY"),
                          c("寛解導入療法", "地固め療法・維持療法", "サルベージ療法"),
                          stringsAsFactors=F)
colnames(epoch_order) <- c("epoch_order", "EPOCH", "epoch_ja")
table2 <- raw_EC %>% filter(ECTRT != "Not Done") %>% left_join(epoch_order, by="EPOCH") %>%
            select(c(USUBJID, EPOCH, ECTPT, ECTRT, epoch_order, epoch_ja, ECSPID)) %>% right_join(table2, by="USUBJID")
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
local({
  temp_RS <- raw_RS %>% select(c(USUBJID, RS_EPOCH="EPOCH", RSORRES, RSDTC, RSSPID))
  temp_CE <- raw_CE %>% select(c(USUBJID, CEDTC)) %>% group_by(USUBJID) %>% filter(CEDTC == min(CEDTC))
  table2 <<- table2 %>% left_join(temp_RS, by=c("USUBJID", "EPOCH"="RS_EPOCH")) %>% left_join(temp_CE, by="USUBJID")
})
table2$remission <- ifelse(is.na(table2["EPOCH"]), NA,
                           ifelse((table2["EPOCH"] == "INDUCTION THERAPY" &
                                   (table2["RSORRES"] == "CR" | table2["RSORRES"] == "CRi")), "有", "無"))
table2$relapse <- ifelse(table2$remission == "無", NA, ifelse(!is.na(table2$CEDTC), "有", "無"))
table2$days_of_remission <- ifelse(table2$remission == "無", NA, as.Date(table2$CEDTC) - as.Date(table2$RSDTC) + 1)
DS <- raw_DS %>% inner_join(MH, by="USUBJID") %>%
        select(c(USUBJID, DSSTDTC, DSTERM, MHDTC, MHCAT)) %>% filter(DSTERM == "DEATH")
DS$survival_days <- as.Date(DS$DSSTDTC) - as.Date(DS$MHDTC) + 1
table2 <- table2 %>% left_join(DS, by="USUBJID")
# sort
table2 <- table2 %>% arrange(molecular_subjid, 登録コード, epoch_order, ECTPT)
output_table2 <- table2 %>% ungroup %>%  # Ungroup data frames grouped by USUBJID
                  select(c(USUBJID=molecular_subjid, 登録コード, 施設名.ja., age, str_sex, name_en, epoch_ja, ECTRT,
                           remission, relapse, days_of_remission, survival_days)) %>% distinct()
# ------ output for validation ------
write.csv(table1, "/Users/mariko/Documents/GitHub/JALSG-CS-17-Molecular/output/validation_table1.csv",
          fileEncoding="cp932")
write.csv(output_table2, "/Users/mariko/Documents/GitHub/JALSG-CS-17-Molecular/output/validation_table2.csv",
          fileEncoding="cp932")
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
