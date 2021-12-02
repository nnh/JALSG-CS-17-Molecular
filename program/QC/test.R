# 確認用PG
library(dplyr)
library(openxlsx)
# table1
test_dm <- raw_DM
test_dm$test <- 1
table1test <- summarise(group_by(test_dm, SITEID), total=sum(test))
df_shisetsu <- raw_facilities
df_shisetsu$code <- as.numeric(df_shisetsu$施設コード)
table1test <- left_join(table1test, df_shisetsu, c("SITEID"="code"))
nrow(test_dm)
# table2
# 症例番号、登録コード
# 症例番号	CS-17-Molecular_registration_yyyymmdd_mmss.csv.症例登録番号	"CS-17_registration_yyyymmdd_mmss.csv.登録コードが存在するもののみ出力
#DM.SUBJIDとCS-17_registration_yyyymmdd_mmss.csv.症例登録番号をキーにマージする"
#df_reg <- select(CS_17, c(症例登録番号, 登録コード))
#df_mol <- select(CS_17_Molecular, c(症例登録番号, 登録コード))
#df_mol_reg <- right_join(df_mol, df_reg, by="登録コード")
#df_id <- right_join(df_mol_reg, raw_DM["SUBJID"], by=c("症例登録番号.y"="SUBJID"))
#df_id <- select(df_id, c(症例登録番号="症例登録番号.y", 登録コード))
#df_id$id <- formatC(df_id$症例登録番号, digits=3, flag="0")
#df_id$id <- paste0("CS-17-", df_id$id)
# 施設名
df_shisetsumei <- left_join(raw_DM, df_shisetsu, by=c("SITEID"="code"))
df_shisetsumei <- select(df_shisetsumei, c(USUBJID, 施設名.ja.))
# 年齢性別
df_age <- inner_join(raw_DM, raw_MH, by="USUBJID")
df_age <- filter(df_age, MHCAT=="PRIMARY DIAGNOSIS")
df_age <- select(df_age, c(USUBJID, MHDTC, BRTHDTC, SEX))
df_age$seibetsu <- ifelse(df_age$SEX==0,"M", "F")
# 診断病名
df_byoumei <- filter(raw_MH, MHCAT=="PRIMARY DIAGNOSIS")
# DMになければ消す
df_byoumei <- inner_join(df_byoumei, raw_DM, by="USUBJID")
df_byoumei <- select(df_byoumei, c(USUBJID, MHTERM))
df_byoumei$intbyoumei <- as.numeric(df_byoumei$MHTERM)
df_byoumei <- left_join(df_byoumei, raw_diseases, by=c("intbyoumei"="code"))
df_byoumei <- select(df_byoumei, c(USUBJID, name_en))
# 投与時期,治療内容
# DMになければ消す
df_touyo_chiryo <- inner_join(raw_EC, raw_DM, by="USUBJID")
df_touyo_chiryo <- select(df_touyo_chiryo, c(USUBJID, EPOCH, ECTPT, ECTRT))
# 寛解もしくは安定の有無
df_id <- select(raw_DM, c(USUBJID))
df_kankai <- select(raw_RS, c(USUBJID, RSORRES, EPOCH, RSDTC))
df_kankai <- left_join(df_id, df_kankai, c("USUBJID"="USUBJID"))
df_kankai$umu <- "無"
sv_id <- "0"
for (i in 1:nrow(df_kankai)){
  if (is.na(df_kankai[i, "EPOCH"])){
    df_kankai[i, "umu"] <- NA
    df_kankai[i, "EPOCH"] <- "."
  }
  if (df_kankai[i, "EPOCH"] == ""){
    df_kankai[i, "umu"] <- NA
  }
  if (df_kankai[i, "EPOCH"] == "INDUCTION THERAPY"){
    if (df_kankai[i, "RSORRES"] == "CR"){
      df_kankai[i, "umu"] <- "有"
    }
    if (df_kankai[i, "RSORRES"] == "CRi"){
      df_kankai[i, "umu"] <- "有"
    }
  }
  if (sv_id == df_kankai[i, "USUBJID"]){
    sv_id <- df_kankai[i, "USUBJID"]
    df_kankai[i, "USUBJID"] <- NA
  } else {
    sv_id <- df_kankai[i, "USUBJID"]
  }
}
df_kankai <- subset(df_kankai, !is.na(USUBJID))
write.csv(df_kankai, "~/Documents/GitHub/JALSG-CS-17-Molecular/output/df_kankai.csv",
          fileEncoding="cp932")
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
