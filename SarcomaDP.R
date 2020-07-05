library(readxl)
library(dplyr)

#read data

setwd("C:/Users/USER/Desktop/2020-1학기/의학통계/sarcoma/Distal Pancreatectomy")
a <- excel_sheets("sarcoma data sheet SMC 20200629_DP추가.xlsx") %>% 
  lapply(function(x){read_excel("sarcoma data sheet SMC 20200629_DP추가.xlsx",sheet=x,skip=2, na = "UK")})
b <- a[[1]] %>% 
  left_join(a[[2]], by = "환자번호") %>% left_join(a[[3]], by = "환자번호") %>% left_join(a[[4]], by = "환자번호") %>%
  left_join(a[[5]], by = "환자번호") %>% left_join(a[[6]], by = "환자번호") %>% left_join(a[[7]], by = "환자번호")

#Age 계산
b$Age <- as.numeric(b[["수술날짜\r\n\r\ndd-mm-yyyy"]] - b[["생년월일\r\n\r\ndd-mm-yyyy"]])/365.25

#연구대상 추리기
c <- b %>% 
  filter(`1:DP    2:Non-DP 9:제외`== 1 | `1:DP    2:Non-DP 9:제외`== 2)

#out에 데이터 만들기 시작
out <- c %>% select(환자번호,Age,`성별\r\n\r\nM/F`)
names(out)[3] <- "Sex"; names(out)[1] <- "ID"
out$Sex <- as.factor(out$Sex)

#Distal Pancreatectomy : 1=DP, 2=non-DP
out$DP<-as.integer(c[["1:DP    2:Non-DP 9:제외"]])

#Death
out$Death<-ifelse(c[["사망여부\r\n\r\n0.Alive\r\n1.Dead\r\n2.Unknown.y"]] == "1", T,
                  ifelse(c[["사망여부\r\n\r\n0.Alive\r\n1.Dead\r\n2.Unknown.y"]]== "0", F, NA)) %>% as.integer
#관찰기간
out$day_FU <- as.numeric(c[["마지막 f/u\r\n\r\ndd-mm-yyyy"]] - c[["수술날짜\r\n\r\ndd-mm-yyyy"]])

out$recur_local <- c[["재발#1\r\n\r\n0: 무\r\n1: 유.x"]]
out$recur_site <- c$`Site of local recurrence`
out$recur_site <- ifelse(out$recur_site == "6", NA, out$recur_site)
out$recur_day <- ifelse(c[["재발#1\r\n\r\n0: 무\r\n1: 유.x"]] == 1, 
                        as.numeric(as.Date(as.integer(c[["Date of local recurrence"]]), origin = "1899-12-30") - as.Date(c[["수술날짜\r\n\r\ndd-mm-yyyy"]])),
                        as.numeric(c[["마지막 f/u\r\n\r\ndd-mm-yyyy"]] - c[["수술날짜\r\n\r\ndd-mm-yyyy"]]))

#BMI
height <- as.numeric(c[["키\r\n(cm)"]])/100
weight <- as.numeric(c[["몸무게\r\n(kg)"]])
out$BMI <- weight/height/height
out$BMI_cat <- factor(ifelse(out$BMI < 18.5, "< 18.5", ifelse(out$BMI < 25, "< 25", ifelse(out$BMI < 30, "< 30", "≥ 30"))))


#DM : 1=TRUE, 0=FALSE
out$DM <- as.integer(c[["DM\r\n\r\n0. No\r\n1.yes"]])

#HTN : 1=TRUE, 0=FALSE
out$HTN <- as.integer(c[["HTN\r\n\r\n0. No\r\n1.yes"]])

#COPD : 1=TRUE, 0=FALSE
out$COPD <- as.integer(c[["COPD\r\n\r\n0. No\r\n1.yes"]])

#Coronary artery disease : 1=TRUE, 0=FALSE
out$CoronaryArteryDisease <- as.integer(c[["Coronary artery disease\r\n\r\n0. No\r\n1.yes"]])

#Chronic renal disease : 1=TRUE, 0=FALSE
out$ChronicRenalDisease <- as.integer(c[["Chronic renal disease\r\n\r\n0. No\r\n1.yes"]])


#Prev abdominal OP Hx : 0="무", 1="유", 2="기타(laparo)"
out$PrevAbdominalOp <- c[["이전\r\nabdominal op Hx \r\n여부\r\n\r\n0. 무\r\n1. 유\r\n2. 기타(laparo)"]]
#out$PrevAbdominalOp <- as.factor(out$PrevAbdominalOp)

#PreOP chemo : 1=TRUE, 0=FALSE
out$preOpChemo <- as.integer(c[["Neoadjuvant chemo 여부\r\n\r\n0.No\r\n1.Yes"]])


#Hb
out$Hb <- as.numeric(c[["수술전 \r\n피검사\r\n\r\nHb\r\n(g/dL)"]])
out$Hb_below9 <- as.integer(out$Hb < 9)  
out$Hb_below10 <- as.integer(out$Hb < 10)


#Albumin
out$Albumin <- as.numeric(c[["수술전 피검사\r\n\r\nAlbumin\r\n(g/dL)"]])
out$Albumin_below3 <- as.integer(out$Albumin < 3)

#PLT
out$PLT <- as.numeric(c[["수술전 피검사\r\n\r\nPlatelet\r\n(1000/uL)"]])
out$PLT_below50 <- as.integer(out$PLT < 50)
out$PLT_below100 <- as.integer(out$PLT < 100)

#PT INR
out$PT_INR <- as.numeric(c[["수술전 피검사\r\n\r\nPT(INR)"]])
out$PT_INR_over1.5 <- as.integer(out$PT_INR > 1.5)

#Tumor size
out$TumorSize <- as.numeric(c[["종양크기\r\nFirst dimension\r\n(mm)"]])


#Tumor histologic subtype
#LPS : 0. WD Liposarcoma / 1. DD Liposarcoma / 2. Pleomorphic Liposarcoma / 7. 중 comment 에 liposarcoma
#nonLPS : 3. Leiomyosarcoma / 4. MPNST / 5. Solitary fibrous tumor / 6. PEComa / 7. 중 comment 에 liposarcoma 없음.
out$Liposarcoma_postop <- as.integer((c[["병리결과\r\n\r\n0. WD Liposarcoma\r\n1. DD Liposarcoma\r\n2. Pleomorphic Liposarcoma\r\n3. Leiomyosarcoma\r\n4. MPNST\r\n5. Solitary fibrous tumor\r\n6. PEComa\r\n7. Other"]] %in% c(0, 1, 2)) |
                                       (c[["병리결과\r\n\r\n0. WD Liposarcoma\r\n1. DD Liposarcoma\r\n2. Pleomorphic Liposarcoma\r\n3. Leiomyosarcoma\r\n4. MPNST\r\n5. Solitary fibrous tumor\r\n6. PEComa\r\n7. Other"]] == 7) &
                                       grepl("liposarcoma|Liposarcoma", c[["Other \r\n\r\ncomment"]]))  
#FNCLCC grade
out$FNCLCC <- as.factor(c[["FNCLCC grade\r\n\r\n1. total score 2-3\r\n2. total score 4-5\r\n3. total score 6,7,8"]])

#Tumor Resection
#R0/R1="0", R2="1", other=NA ("2"도 NA에 포함)
out$Resection <- c[["Surgical margins\r\n\r\n0. R0/R1\r\n1. R2\r\n2. Not available"]]
out$Resection <- as.factor(ifelse(out$Resection=="2", NA, out$Resection))

#Combined Organ Resection
# "colon resection" : Rt. + Lt. + rectum 
# "small bowel resection" : small bowel + duodenum
# "pancreas resection" : distal pan + PD 
# "liver resection" 
# "major vessel resection" : iliac a & v, IVC, aorta
out$Resection_Colon <- as.integer(
  (c[["동반절제 장기\r\nRight colon\r\n\r\n0. No\r\n1. Yes"]] == "1") | (c[["동반절제 장기\r\nLeft colon\r\n\r\n0. No\r\n1. Yes"]] == "1") | (c[["동반절제 장기\r\nRectum\r\n\r\n0. No\r\n1. Yes"]] == "1")
)
out$Resection_SmallBowel <- as.integer(
  (c[["동반절제 장기\r\nSmall bowel\r\n\r\n0. No\r\n1. Yes"]] == "1") | (c[["동반절제 장기\r\nDuodenum\r\n\r\n0. No\r\n1. Yes"]] == "1")
)
out$Resection_Pancreas <- as.integer(
  (c[["동반절제 \r\n장기\r\nDistal pancreas\r\n\r\n0. No\r\n1. Yes"]] == "1") | (c[["동반절제 \r\n장기\r\nPanreatico-duodenectomy\r\n\r\n0. No\r\n1. Yes"]] == "1")
)
out$Resection_Liver <- as.integer((c[["동반절제 장기\r\nLiver\r\n\r\n0. No\r\n1. Yes"]]=="1"))
out$Resection_MajorVesselResection <- as.integer(
  (c[["동반절제 장기\r\nIliac vein\r\n\r\n0. No\r\n1. Yes"]] == "1") | (c[["동반절제 장기\r\nIVC\r\n\r\n0. No\r\n1. Yes"]] == "1") | (c[["동반절제 장기\r\nIliac artery\r\n\r\n0. No\r\n1. Yes"]] == "1") | (c[["동반절제 장기\r\nAorta\r\n\r\n0. No\r\n1. Yes"]] == "1")
)

#OP time
out$opTime <- as.numeric(c[["수술시간\r\n(min)"]])

#intra OP transfusion
out$intraOpTransfusion <- as.integer(c[["PRBC 수혈 수"]])

#Estimated blood loss
out$EBL <- as.numeric(c[["EBL\r\n(ml)"]])

#C-D complication : complication==1 중에서 grade==2 는 complication=0으로 바꿈
out$ClavienDindoComplication <- as.integer(c[["Clavien-Dindo complication \r\n\r\n0. No\r\n1. Yes"]])
out$ClavienDindoComplication<-ifelse(out$ClavienDindoComplication==1 & c[["Clavien-Dindo grade \r\n\r\n2/3a/3b/4a/4b/5"]]=="2",0,out$ClavienDindoComplication)

#post OP transfusion
out$postOpTransfusion <- as.integer(c[["수술 후 PRBC 수혈 여부\r\n\r\n0. No\r\n1. Yes"]])

#ICU care
out$ICUcare <- as.integer(c[["ICU 입실여부\r\n\r\n0. No\r\n1. Yes"]])

#Return to OR
out$ReOP <- as.integer(c[["합병증으로 인한 Re-op 여부\r\n\r\n0. No\r\n1. Yes"]])

#Hospital Stay after OP
out$HospitalDay <- as.numeric(c[["재원일수(days)"]])

#RT gray
out$RTgray <- c[["RT dose\r\n(Gy)"]]

for (vname in names(c)[c(128:135, 137:139)]){
  vn.new <- gsub(" ", ".", strsplit(vname, "\r")[[1]][1])
  out[[vn.new]] <- as.integer(c[[vname]])
}

write.csv(out,"out.csv",row.names=F)

#-직접 Table1 만들어보기---------------------------------------------------------------------------------------------------

library(data.table)

tab1<-data.frame("variable"=character(),
                         "subgroup"=character(),
                         "DP"=character(),"nonDP"=character(),
                         "pvalue"=numeric())

for(i in 1:ncol(out)){
  va<-names(out)[i]
  
  if((is.numeric(out[[va]]) | is.integer(out[[va]])) & length(table(out[va]))>5){
    #continous
    ttest<-t.test(out[[va]]~out$DP,data=out,var.equal=F)
    tab1<-rbind(tab1,
                data.frame("variable"=va,
                           "subgroup"=" ",
                           "DP"=paste(round(ttest$estimate[1],2),"(",round(as.numeric(by(out[[va]],out$DP,sd)[1]),2),")"),
                           "nonDP"=paste(round(ttest$estimate[2],2),"(",round(as.numeric(by(out[[va]],out$DP,sd)[2]),2),")"),
                           "pvalue"=round(ttest$p.value,3)))
  }else if(length(table(out[va]))<=5 & length(table(out[va]))>=2){
    #factor
    tb.va<-table(data.frame(out[[va]],out$DP))
    cchisq<-chisq.test(tb.va)
    check.fisher<-0
    for(ii in 1:nrow(tb.va)){
      for(jj in 1:ncol(tb.va)){
        if(cchisq$expected[ii,jj]<5){
          check.fisher<-1
          break
        }
      }
    }

    if(check.fisher==0){
      #chi-square test
      for(ii in 1:nrow(tb.va)){
        if(ii==1){
          tab1<-rbind(tab1,
                      data.frame("variable"=paste(va,"(%)"),
                                 "subgroup"=rownames(tb.va)[ii],
                                 "DP"=paste(tb.va[ii,1],"(",round(tb.va[ii,1]/colSums(tb.va)[1]*100,2),")"),
                                 "nonDP"=paste(tb.va[ii,2],"(",round(tb.va[ii,2]/colSums(tb.va)[2]*100,2),")"),
                                 "pvalue"=round(cchisq$p.value,3)))
        }else{
          tab1<-rbind(tab1,
                      data.frame("variable"=" ",
                                 "subgroup"=rownames(tb.va)[ii],
                                 "DP"=paste(tb.va[ii,1],"(",round(tb.va[ii,1]/colSums(tb.va)[1]*100,2),")"),
                                 "nonDP"=paste(tb.va[ii,2],"(",round(tb.va[ii,2]/colSums(tb.va)[2]*100,2),")"),
                                 "pvalue"=" "))
        }
      }
      
    }else
    {
      #fisher's test
      ffisher<-fisher.test(tb.va)
      for(ii in 1:nrow(tb.va)){
        if(ii==1){
          tab1<-rbind(tab1,
                      data.frame("variable"=va,
                                 "subgroup"=rownames(tb.va)[ii],
                                 "DP"=paste(tb.va[ii,1],"(",round(tb.va[ii,1]/colSums(tb.va)[1]*100,2),")"),
                                 "nonDP"=paste(tb.va[ii,2],"(",round(tb.va[ii,2]/colSums(tb.va)[2]*100,2),")"),
                                 "pvalue"=round(ffisher$p.value,3)))
        }else{
          tab1<-rbind(tab1,
                      data.frame("variable"=" ",
                                 "subgroup"=rownames(tb.va)[ii],
                                 "DP"=paste(tb.va[ii,1],"(",round(tb.va[ii,1]/colSums(tb.va)[1]*100,2),")"),
                                 "nonDP"=paste(tb.va[ii,2],"(",round(tb.va[ii,2]/colSums(tb.va)[2]*100,2),")"),
                                 "pvalue"=" "))
        }
      }
    }
  }
}

write.csv(tab1,"tableone-handmade.csv",row.names=F)

#-Tableone R package 이용해서 만들어보기---------------------------------------------------------------------------------------------------

library(tableone)

## Variable list: For select UI in ShinyApps
varlist <- list(
  Base = c("DP", "Age", "Sex", "BMI", "BMI_cat",  "DM", "HTN", "COPD", "CoronaryArteryDisease", "ChronicRenalDisease", "PrevAbdominalOp", "preOpChemo", 
           "Hb", "Hb_below9", "Hb_below10", "Albumin", "Albumin_below3", "PLT", "PLT_below50", "PLT_below100", "PT_INR", "PT_INR_over1.5", "TumorSize",
           "Liposarcoma_postop", "RTgray","FNCLCC",
           "Resection", grep("Resection_", names(out), value = T), "opTime", "intraOpTransfusion", "EBL"),
  Complication = c("ClavienDindoComplication", "postOpTransfusion", "ICUcare", "ReOP", "HospitalDay", names(out)[47:ncol(out)]),
  Event = c("Death", "recur_local"),
  Day = c("day_FU", "recur_day")
)

out <- data.table(out[, unlist(varlist)])
my_vars<-names(out)
factor_vars <- c(names(out)[sapply(out, function(x){length(table(x))}) <= 5])
out[, (factor_vars) := lapply(.SD, factor), .SDcols = factor_vars]
conti_vars <- setdiff(my_vars, factor_vars)

tab2 <- CreateTableOne(vars=my_vars,strata="DP",data=out,factorVars=factor_vars)
write.csv(print(tab2,showAllLevels = TRUE),"tableone-package.csv",row.names=TRUE)

#fisher-test 적용 안되고 다 chisq test 로 되는 것 같고..
#t-test p value 값도 왜 인지 다르다.. 'Two-group ANOVA is equivalent of t-test.'라고 되어있는데 왜 다른지 모르겠다.