rm(list = ls())
#load packages
library(pacman)
pacman::p_load(tidyverse,
               readxl,
               xlsx,
               writexl,
               openxlsx,
               foreign,
               haven,
               janitor,
               lubridate,
               ggrcs,
               rms,
               pROC,
               xgboost,
               data.table,
               Matrix,
               skimr,
               DataExplorer,
               GGally,
               caret,
               pROC,
               dplyr,
               ggplot2,
               ggpubr,
               ggprism,
               rms,
               vip,
               lattice,
               nnet,
               foreign,
               ranger,
               gtsummary)
library(gt)
library(officer)
library(flextable)
library(pROC)

##### Step 1. Upload data ####
merged_coh<-readRDS('merged_coh.rds')  #total coh dara
merged_et<-readRDS('merged_et.rds')  # total cumulative cycles
df_fet<-readRDS('fet.rds')  #total fet cycles
df_fresh<-readRDS('fresh.rds')  #total fresh et cycles


##### Step 2. table1  baseline characteristics ####
tbl <- tbl_summary(
  data = merged_coh,  
  by = TyG_4cat, 
  include = c('PCOS','DOR','adenomyosis','endometriosis','tube factor','male factor',
              'COH_protocol','Gn_starting','Gn_duration','Gn_total','HCG_E2','HCG_P','HCG_LH'
              ),
  digits = all_continuous() ~ 3
) %>% 
  add_overall() %>%  
  add_p(pvalue_fun = ~style_pvalue(.x, digits = 3))  %>%  
  add_n(statistic = "{N_nonmiss}", col_label = "**N**", footnote = TRUE) %>% 
  modify_caption("**Table 1. Ovarian stimulation treatments,and oocyte/embryo parameters**") 

tbl_flex <- as_flex_table(tbl)
# save as docx 
read_docx() %>%
  body_add_flextable(tbl_flex) %>% 
  print(target = "table1.docx")

####fet logistic####
model <- glm(live_birth~  F_age + em_thickness+embryos_transferred+good_quality_embryos_transferred+
               FET_protocol+TyG_BMI_4cat+PCOS+DOR+adenomyosis+endometriosis, ,
             data = df_fet,
             family = binomial(),
             na.action =na.omit )
summary(model)
model <- glm(miscarriage~  F_age + em_thickness+embryos_transferred+good_quality_embryos_transferred+
               FET_protocol+TyG_BMI_4cat+PCOS+DOR+adenomyosis+endometriosis, ,
             data = df_fet,
             family = binomial(),
             na.action =na.omit )
summary(model)
model <- glm(cli_pregnancy~  F_age + em_thickness+embryos_transferred+good_quality_embryos_transferred+
               FET_protocol+TyG_BMI_4cat+PCOS+DOR+adenomyosis+endometriosis, ,
             data = df_fet,
             family = binomial(),
             na.action =na.omit )
summary(model)
model <- glm(bio_pregnancy~  F_age + em_thickness+embryos_transferred+good_quality_embryos_transferred+
               FET_protocol+TyG_BMI_4cat+PCOS+DOR+adenomyosis+endometriosis,,
             data = df_fet,
             family = binomial(),
             na.action =na.omit )
summary(model)


####fresh logistic
model <- glm(live_birth~  F_age +em_thickness+embryos_transferred+good_quality_embryos_transferred+
               TyG_BMI_4cat+PCOS+DOR+adenomyosis+endometriosis ,
             data = df_fresh,
             family = binomial(),
             na.action =na.omit )

summary(model) 
model <- glm(miscarriage~  F_age +em_thickness+embryos_transferred+good_quality_embryos_transferred+
               TyG_BMI_4cat+PCOS+DOR+adenomyosis+endometriosis ,
             data = df_fresh,
             family = binomial(),
             na.action =na.omit )

summary(model) 

model <- glm(cli_pregnancy~  F_age +em_thickness+embryos_transferred+good_quality_embryos_transferred+
               TyG_BMI_4cat+PCOS+DOR+adenomyosis+endometriosis ,
             data = df_fresh,
             family = binomial(),
             na.action =na.omit )

summary(model) 

model <- glm(bio_pregnancy~  F_age +em_thickness+embryos_transferred+good_quality_embryos_transferred+
               TyG_BMI_4cat+PCOS+DOR+adenomyosis+endometriosis ,
             data = df_fresh,
             family = binomial(),
             na.action =na.omit )

summary(model) 


####clbr logistic####
model <- glm(cLB_2cat ~ F_age+AFC+good_quality_embryos+PCOS+DOR+TyG_BMI_4cat+adenomyosis+endometriosis,
             data =merged_et,
             family = binomial(),
             na.action =na.omit )
summary(model)

results <- tidy(model, conf.int = TRUE)
formatted_results <- results %>%
  mutate(
    OR = exp(estimate),
    CI_lower = exp(conf.low),
    CI_upper = exp(conf.high),
    OR_CI = sprintf("%.2f (%.2f–%.2f)", OR, CI_lower, CI_upper),  # 格式化 OR 和 95% CI
    P_value = sprintf("%.3f", p.value)  
  ) %>%
  select(term, P_value, OR_CI,OR,CI_lower,CI_upper)

ft <- flextable(formatted_results) %>%
  set_header_labels(term = "characteristics", P_value = "P-value", OR_CI = "OR (95% CI)") %>%
  colformat_double(digits = 3) %>%  
  theme_zebra() %>%
  autofit()
read_docx() %>% 
  body_add_flextable(ft) %>% 
  print(target = "Logistic_clbr.docx")


model <- glm(cCP_2cat ~ F_age+AFC+good_quality_embryos+PCOS+DOR+TyG_BMI_4cat+adenomyosis+endometriosis,
             data =merged_et,
             family = binomial(),
             na.action =na.omit )
summary(model)

results <- tidy(model, conf.int = TRUE)
formatted_results <- results %>%
  mutate(
    OR = exp(estimate),
    CI_lower = exp(conf.low),
    CI_upper = exp(conf.high),
    OR_CI = sprintf("%.2f (%.2f–%.2f)", OR, CI_lower, CI_upper),  # 格式化 OR 和 95% CI
    P_value = sprintf("%.3f", p.value)  
  ) %>%
  select(term, P_value, OR_CI,OR,CI_lower,CI_upper)

ft <- flextable(formatted_results) %>%
  set_header_labels(term = "characteristics", P_value = "P-value", OR_CI = "OR (95% CI)") %>%
  colformat_double(digits = 3) %>%  
  theme_zebra() %>%
  autofit()
read_docx() %>% 
  body_add_flextable(ft) %>% 
  print(target = "Logistic_cCP.docx")


#subgroup analysis####

model <- glm(cLB_2cat ~ F_age+AFC+good_quality_embryos+PCOS+DOR+TyG_BMI_4cat+adenomyosis+endometriosis,
             data =merged_et,
             family = binomial(),
             subset = age_2cat==0,
             na.action =na.omit )
summary(model)
model <- glm(cLB_2cat ~ F_age+AFC+good_quality_embryos+PCOS+DOR+TyG_BMI_4cat+adenomyosis+endometriosis,
             data =merged_et,
             family = binomial(),
             subset = age_2cat==1,
             na.action =na.omit )
summary(model)

model <- glm(cLB_2cat ~ F_age+AFC+good_quality_embryos+PCOS+DOR+TyG_4cat+adenomyosis+endometriosis,
             data =merged_et,
             family = binomial(),
             subset = ovarian_res==0,
             na.action =na.omit )
summary(model)
model <- glm(cLB_2cat ~ F_age+AFC+good_quality_embryos+PCOS+DOR+TyG_BMI_4cat+adenomyosis+endometriosis,
             data =merged_et,
             family = binomial(),
             subset = ovarian_res==1,
             na.action =na.omit )
summary(model)
model <- glm(cLB_2cat ~ F_age+AFC+good_quality_embryos+PCOS+DOR+TyG_BMI_4cat+adenomyosis+endometriosis,
             data =merged_et,
             family = binomial(),
             subset = ovarian_res==2,
             na.action =na.omit )
summary(model)


####森林图####
dat <- read_excel("logistic+forest_plot/bio_pregnancy_fet_forest.xlsx", 
                  col_types = c("guess", "guess", "text", "guess", "guess", "guess"),
                  col_names = F)
dat <- read_excel("logistic+forest_plot/cli_pregnancy_fet_forest.xlsx", 
                  col_types = c("guess", "guess", "text", "guess", "guess", "guess"),
                  col_names = F)
dat <- read_excel("logistic+forest_plot/miscarriage_fet_forest.xlsx", 
                  col_types = c("guess", "guess", "text", "guess", "guess", "guess"),
                  col_names = F)
dat <- read_excel("logistic+forest_plot/live_birth_fet_forest.xlsx", 
                  col_types = c("guess", "guess", "text", "guess", "guess", "guess"),
                  col_names = F)

dat$...3 <- case_when(
  is.na(dat$...3) ~ NA_character_,          
  dat$...3 == "Ref" ~ "Ref",
  dat$...3 == "<0.001" ~ "<0.001",
  dat$...3 == "P-value" ~ "P-value",
  TRUE ~ format(as.numeric(dat$...3), nsmall = 3, trim = TRUE)
)


p1<-forestplot(as.matrix(dat[,1:3]), 
               dat$...4,
               dat$...5, 
               dat$...6, 
               zero=1, 
               graph.pos=4, 
               graphwidth=unit(50,"mm"), 
               lineheight=unit(5,"mm"),
               colgap = unit(10,"mm"),
               boxsize=0.3, 
               is.summary = c(TRUE,rep(FALSE,50)), 
               ci.vertices = TRUE,
               ci.vertices.height=0.15,
               xticks=(c(0,0.5,1.0,1.5,2.0)), 
               clip=c(0.5,2.0), 
               lwd.ci =2,
               col= fpColors(lines = "darkgreen",
                             box="darkgreen",
                             zero = "black"),
               title='Factors associated with biochemical pregnancy'
)


library(gridExtra)

# 假设 p1, p2, p3, p4 是 forestplot() 生成的图
p1_grob <- grid.grabExpr(print(p1))
p2_grob <- grid.grabExpr(print(p2))
p3_grob <- grid.grabExpr(print(p3))
p4_grob <- grid.grabExpr(print(p4))

tiff("myplot.tiff", res=300, width=2400, height=4800)
grid.arrange(p1_grob,p2_grob,p3_grob,p4_grob, ncol = 1)
dev.off()


####ROC for three models####
model1 <- glm(cLB_2cat ~ F_age+AFC+good_quality_embryos+PCOS+DOR+adenomyosis+endometriosis,
              data =merged_et,
              family = binomial(),
              na.action =na.omit )
model2 <- glm(cLB_2cat ~ F_age+AFC+good_quality_embryos+F_BMI+TG+FBG+PCOS+DOR+adenomyosis+endometriosis,
              data =merged_et,
              family = binomial(),
              na.action =na.omit )
model3 <- glm(cLB_2cat ~ F_age+AFC+good_quality_embryos+TyG_BMI+PCOS+DOR +adenomyosis+endometriosis,
              data =merged_et,
              family = binomial(),
              na.action =na.omit )
#predict
pred1 <- predict(model1, type = "response",newdata = merged_et)
pred2 <- predict(model2, type = "response",newdata = merged_et)
pred3 <- predict(model3, type = "response",newdata = merged_et)

# calculate ROC
roc1 <- roc(merged_et$cLB_2cat, pred1)
roc2 <- roc(merged_et$cLB_2cat, pred2)
roc3 <- roc(merged_et$cLB_2cat, pred3)

# plot
tiff("myplot.tiff", res = 300, width = 2000, height = 2200)
plot(roc1, col = "#558934FF", main = "ROC Curves of Three Models")
plot(roc2, col = "#F2A241FF", add = TRUE)
plot(roc3, col = "#EE2617FF", add = TRUE)
legend("topright", 
       legend = c("Model 1", "Model 2", "Model 3"),
       col = c("#F2A241FF", "#558934FF", "#EE2617FF"),
       lty = c(1, 2, 3),cex=0.8,
       bty = "n",
       inset = c(0.01, 0.6),  # 微调位置
       xjust = 1, yjust = 1)
dev.off()

# print AUC
auc(roc1)
ci(roc1)
auc(roc2)
ci(roc2)
auc(roc3)
ci(roc3)
# model1 vs model2
roc.test(roc1, roc2, method = "delong")

# model2 vs model3
roc.test(roc2, roc3, method = "delong")

# model1 vs model3
roc.test(roc1, roc3, method = "delong")

####clbr DCA curve####
library(rmda)
#拆分数据：训练集和测试集
set.seed(111)
index <- sort(sample(nrow(merged_et), nrow(merged_et) * 0.7))
train <- merged_et[index,] 
test <- merged_et[-index,] 

fml1 <- as.formula(cLB_2cat==1 ~ F_age+AFC+good_quality_embryos+PCOS+DOR+adenomyosis+endometriosis)
fml2<-as.formula(cLB_2cat==1 ~ F_age+AFC+good_quality_embryos+F_BMI+TG+FBG+PCOS+DOR+adenomyosis+endometriosis)
fml3<-as.formula(cLB_2cat==1 ~ F_age+AFC+good_quality_embryos+TyG_BMI+TG+FBG+PCOS+DOR+adenomyosis+endometriosis)
model1<-decision_curve(fml1,  
                       data=train,  
                       family = binomial(link = "logit"),
                       policy = c("opt-in"),
                       thresholds = seq(0,1,by=0.01),
                       bootstraps = 500,
                       confidence.intervals = 0.95,
                       study.design = "cohort")
model2<-decision_curve(fml2,  
                       data=train,  
                       family = binomial(link = "logit"),
                       policy = c("opt-in"),
                       thresholds = seq(0,1,by=0.01),
                       bootstraps = 500,
                       confidence.intervals = 0.95,
                       study.design = "cohort")
model3<-decision_curve(fml3,  
                       data=train,  
                       family = binomial(link = "logit"),
                       policy = c("opt-in"),
                       thresholds = seq(0,1,by=0.01),
                       bootstraps = 500,
                       confidence.intervals = 0.95,
                       study.design = "cohort")
head(model3$derived.data)

tiff("myplot.tiff", res=300, width=2000, height=1200)
plot_decision_curve(list(model1,model2,model3),
                    xlim=c(0,1),
                    cost.benefit.axis = T,
                    n.cost.benefit = 6,
                    col=c("#F2A241FF", "#558934FF", "#EE2617FF"),
                    lty=c(1,2,3),
                    confidence.intervals = F,
                    standardize = T,
                    legend.position="none")

legend("topright", 
       legend = c("Model 1", "Model 2", "Model 3"),
       col = c("#F2A241FF", "#558934FF", "#EE2617FF"),
       lty = c(1, 2, 3),cex=0.8,
       bty = "n",
       inset = c(0.01, 0.5), 
       xjust = 1, yjust = 1)
title(main = "Decision Curve Analysis of Three Models")
dev.off()



####RCS####
library(rcssci)
library(readxl)
library(rms)
library(ggplot2)
k_values <- 3:5
fits <- lapply(k_values, 
               function(k) {  
                 lrm(cLB_2cat ~ rcs(TyG_BMI, k)+F_age+AFC+good_quality_embryo+PCOS+DOR+adenomyosis+endometriosis, 
                     data = merged_et)})

aics <- sapply(fits, AIC)

best_k <- k_values[which.min(aics)]

print(paste("Best number of knots:", best_k))

print(paste("Minimum AIC:", min(aics)))


fit<-lrm(cLB_2cat ~ rcs(TyG_BMI, 3)+F_age+AFC+good_quality_embryos+PCOS+DOR+adenomyosis+endometriosis,data=merged_et)

anova(fit)


dd <- datadist(merged_et)

options(datadist = 'dd')

OR<-Predict(fit,TyG_BMI,ref.zero = TRUE)


rcssci_logistic(data=merged_et, y = "cLB_2cat",x = "TyG_BMI",
                covs=c('F_age','AFC','good_quality_embryos','PCOS','DOR','adenomyosis','endometriosis'),knot=3,               
                prob=0.1,filepath = 'new')



####desision tree####
library(rpart)
library(rpart.plot)
library(skimr)
library(GGally)
library(caret) 
data<-merged_et[,c('cLB_2cat','F_age','AFC','AMH','PCOS','DOR','TyG_BMI','adenomyosis','endometriosis')]
set.seed(123)
tree_model <- rpart(cLB_2cat~.,
                     data=data,
                     method = "class",    
                     parms = list(split = "gini"),       
                     control = rpart.control(
                       minsplit = 150,     
                       minbucket = 100,    
                       cp = 0.001,         
                       maxdepth = 8       
                     )
)

#view the importance of variable
varImp(tree_model)

#rpart.plot
tiff("treeplot.tiff", res=300, width=3200, height=2000)
rpart.plot(tree_model,
           extra = 106, 
           type = 3, 
           fallen.leaves = TRUE,
           cex = 0.5,
           box.palette = "RdBu",
           split.cex = 1.2,
           main = "CLBR"
)
dev.off()

