## Variable list: For select UI in ShinyApps
varlist <- list(
  Base = c("DP", "primaryTumor", "DDLPS_postop", "Age", "Sex", "BMI", "BMI_cat",  "DM", "HTN", "COPD", "CoronaryArteryDisease", "ChronicRenalDisease", "PrevAbdominalOp", "preOpChemo", 
           "Hb", "Hb_below9", "Hb_below10", "Albumin", "Albumin_below3", "PLT", "PLT_below50", "PLT_below100", "PT_INR", "PT_INR_over1.5", "TumorSize", "Liposarcoma_postop",
           "FNCLCC", "Resection", grep("Resection_", names(out), value = T), "opTime", "intraOpTransfusion", "EBL"),
  Complication = c("ClavienDindoComplication01", "ClavienDindoComplication", "postOpTransfusion", "ICUcare", "ReOP", "HospitalDay", "RTgray", "Abdominal.abscess", "Bowel.anastomosis.leak",
                   "Biliary.leak", "Bleeding", "Evisceration", "DVT", "Lymphatic.leak", "Pancreatic.leak", "Sepsis", "Urinary.leak", "Ileus"),
  Event = c("Death", "recur_local"),
  Day = c("day_FU", "recur_day")
)

## Don't need `data.table` 
out <- out[, unlist(varlist)]
my_vars <- unlist(varlist)

factor_vars <- c(names(out)[sapply(out, function(x){length(table(x))}) <= 5])

## Base R style
for (v in factor_vars){out[[v]] <- factor(out[[v]])}
## data.table style
#out[, (factor_vars) := lapply(.SD, factor), .SDcols = factor_vars]

## Factor with fisher : with trycatch 
vars.fisher <- sapply(factor_vars, function(x){is(tryCatch(chisq.test(table(out[["DP"]], out[[x]])), error = function(e) e, warning=function(w) w), "warning")})
vars.fisher <- factor_vars[vars.fisher]


conti_vars <- setdiff(my_vars, factor_vars)


tab1 <- lapply(setdiff(names(out), "DP"), function(va){
  if (va %in% conti_vars){
    forms <- as.formula(paste0(va, "~ DP"))
    mean_sd <- aggregate(forms, data = out, FUN = function(x){c(mean = mean(x), sd = sd(x))})
    p <- t.test(forms, data = out, var.equal = F)$p.value
    return(c(va, "", paste0(round(mean_sd[[va]][, "mean"], 2), " (", round(mean_sd[[va]][, "sd"], 2), ")"), ifelse(p < 0.001, "< 0.001", round(p, 3))))
  } else {
    tb <- table(out[[va]], out[["DP"]])
    tb.prop <- round(100 * prop.table(tb, 2), 2)      ## prop.table : 1- byrow 2 - bycol 
    tb.out <- matrix(paste0(tb, " (", tb.prop, ")"), ncol =2)
    p <- ifelse(va %in% vars.fisher, fisher.test(tb)$p.value, chisq.test(tb)$p.value)
    out.final <- cbind(c(paste0(va, " (%)"), rep("", nrow(tb.out) - 1)), rownames(tb), tb.out, c(ifelse(p < 0.001, "< 0.001", round(p, 3)), rep("", nrow(tb.out) - 1)))
    return(out.final)
  }
}) %>% Reduce(rbind, .)

colnames(tab1) <- c("Variable", "Subgroup", "DP", "nonDP", "Pvalue")

write.csv(tab1,"tableone-handmade.csv",row.names=F)


#-Tableone R package-------------------------------------------------------------------------------------------------

library(tableone)
tab2 <- CreateTableOne(vars=my_vars, strata = "DP", data=out,factorVars = factor_vars, argsNormal = list(var.equal = F))
write.csv(print(tab2, showAllLevels = TRUE, exact = vars.fisher),"tableone-package.csv", row.names=TRUE)