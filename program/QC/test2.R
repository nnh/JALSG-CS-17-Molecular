# 確認用PG
library(dplyr)
library(openxlsx)
test_dm <- raw_DM
test_dm <- filter(test_dm, USUBJID=="CS-17-1859" | USUBJID=="CS-17-2530")
df_shisetsu <- raw_facilities
df_shisetsu$code <- as.numeric(df_shisetsu$施設コード)
# table1
test_dm$test <- 1
table1test <- summarise(group_by(test_dm, SITEID), total=sum(test))
table1test <- left_join(table1test, df_shisetsu, c("SITEID"="code"))
table1test <- select(table1test, c(施設名.ja., number_of_cases=total))
table1test <- arrange(table1test, desc(number_of_cases), 施設名.ja.)
table1test$number_of_cases <- as.character(table1test$number_of_cases)
print("table1合計を確認")
setdiff(table1, table1test)
nrow(test_dm)
# table2
CS_17_touroku <- CS_17$登録コード
molecular_touroku <- CS_17_Molecular$登録コード
cs_17_touroku_list <- NULL
for (i in 1:length(CS_17_touroku)){
  for (j in 1:length(molecular_touroku)){
    if(CS_17_touroku[i] == molecular_touroku[j]){
      cs_17_touroku_list <- c(cs_17_touroku_list, molecular_touroku[j])
    }
  }
}
CS_17_SUBJID <- NULL
for (i in 1:length(cs_17_touroku_list)){
  temp <- filter(CS_17, 登録コード==cs_17_touroku_list[i])
  CS_17_SUBJID <- c(CS_17_SUBJID, temp[1,"症例登録番号"])
}
DM_SUBJID <- raw_DM$SUBJID
SUBJID_list <- NULL
for (i in 1:length(DM_SUBJID)){
  for (j in 1:length(CS_17_SUBJID)){
    if (DM_SUBJID[i] == CS_17_SUBJID[j]){
      SUBJID_list <- c(SUBJID_list, DM_SUBJID[i])
    }
  }
}
USUBJID_list <- paste0("CS-17-", SUBJID_list)
test_table2_dm <- filter(raw_DM, SUBJID %in% SUBJID_list)
test_table2_mh <- filter(raw_MH, USUBJID %in% USUBJID_list)
test_table2_ec <- filter(raw_EC, USUBJID %in% USUBJID_list)
test_table2_rs <- filter(raw_RS, USUBJID %in% USUBJID_list)
# 施設名
df_shisetsumei <- left_join(test_table2_dm, df_shisetsu, by=c("SITEID"="code"))
df_shisetsumei <- select(df_shisetsumei, c(USUBJID, 施設名.ja.))
# 年齢性別
df_age <- inner_join(test_table2_dm, raw_MH, by="USUBJID")
df_age <- filter(df_age, MHCAT=="PRIMARY DIAGNOSIS")
df_age <- select(df_age, c(USUBJID, MHDTC, BRTHDTC, SEX))
df_age$seibetsu <- ifelse(df_age$SEX==0,"M", "F")
# 診断病名
df_byoumei <- filter(test_table2_mh, MHCAT=="PRIMARY DIAGNOSIS")
# DMになければ消す
df_byoumei <- inner_join(df_byoumei, test_table2_dm, by="USUBJID")
df_byoumei <- select(df_byoumei, c(USUBJID, MHTERM))
df_byoumei$intbyoumei <- as.numeric(df_byoumei$MHTERM)
df_byoumei <- left_join(df_byoumei, raw_diseases, by=c("intbyoumei"="code"))
df_byoumei <- select(df_byoumei, c(USUBJID, name_en))
# 投与時期,治療内容
# DMになければ消す
# 治療内容	EC.ECTRT	"日本語ラベル
df_touyo_chiryo <- inner_join(test_table2_ec, test_table2_dm, by="USUBJID")
df_touyo_chiryo <- select(df_touyo_chiryo, c(USUBJID, EPOCH, ECTPT, ECTRT))
# 寛解もしくは安定の有無
df_kankai <- select(test_table2_rs, c(USUBJID, RSORRES, EPOCH, RSDTC))
df_kankai$umu <- "無"
for (i in 1:nrow(df_kankai)){
  if (df_kankai[i, "EPOCH"] == "INDUCTION THERAPY"){
    if (df_kankai[i, "RSORRES"] == "CR"){
      df_kankai[i, "umu"] <- "有"
    }
    if (df_kankai[i, "RSORRES"] == "CRi"){
      df_kankai[i, "umu"] <- "有"
    }
  }
}
# 再発もしくは増悪の有無*1
df_kankai_ari <- filter(df_kankai, umu == "有")
df_saihatsu <- left_join(df_kankai_ari, raw_CE, by="USUBJID")
df_saihatsu$umu <- NA
for (i in 1:nrow(df_saihatsu)){
  if (is.na(df_saihatsu[i, "CEDTC"])){
    df_saihatsu[i, "umu"] <- "無"
  } else {
    df_saihatsu[i, "umu"] <- "有"
  }
}
# 寛解日数もしくは無増悪期間*1
df_muzouakukikan <- left_join(df_kankai_ari, raw_CE, by="USUBJID")
df_muzouakukikan <- select(df_muzouakukikan, c(USUBJID, RSDTC, CEDTC))
# 生存日数
df_shibou <- filter(raw_DS, DSTERM=="DEATH")
df_seizonnissuu <- left_join(df_shibou, raw_MH, by="USUBJID")
df_seizonnissuu <- filter(df_seizonnissuu, MHCAT=="PRIMARY DIAGNOSIS")
df_seizonnissuu <- select(df_seizonnissuu, c(USUBJID, DSSTDTC, MHDTC))
# excel出力
output_wb <- createWorkbook()
addWorksheet(output_wb, "table1")
addWorksheet(output_wb, "id_jplsg")
addWorksheet(output_wb, "年齢性別")
addWorksheet(output_wb, "医療機関名")
addWorksheet(output_wb, "診断病名")
addWorksheet(output_wb, "投与時期,治療内容")
addWorksheet(output_wb, "寛解")
addWorksheet(output_wb, "再発")
addWorksheet(output_wb, "寛解日数")
addWorksheet(output_wb, "生存日数")
writeData(output_wb, sheet="table1", x=table1test, withFilter=F, sep=",")
writeData(output_wb, sheet="id_jplsg", x=df_id, withFilter=F, sep=",")
writeData(output_wb, sheet="年齢性別", x=df_age, withFilter=F, sep=",")
writeData(output_wb, sheet="医療機関名", x=df_shisetsumei, withFilter=F, sep=",")
writeData(output_wb, sheet="診断病名", x=df_byoumei, withFilter=F, sep=",")
writeData(output_wb, sheet="投与時期,治療内容", x=df_touyo_chiryo, withFilter=F, sep=",")
writeData(output_wb, sheet="寛解", x=df_kankai, withFilter=F, sep=",")
writeData(output_wb, sheet="再発", x=df_saihatsu, withFilter=F, sep=",")
writeData(output_wb, sheet="寛解日数", x=df_muzouakukikan, withFilter=F, sep=",")
writeData(output_wb, sheet="生存日数", x=df_seizonnissuu, withFilter=F, sep=",")
saveWorkbook(output_wb, "/Users/admin/Documents/GitHub/CS-17-Molecular_interim/output/test.xlsx", overwrite=T)
