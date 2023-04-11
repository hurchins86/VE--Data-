
################################################################################
##################### Training DQR #############################################
################################################################################
rm(list=ls())
Sys.setenv(JAVA_HOME='C:\\Program Files (x86)\\Java\\jdk-11.0.16')
library(rJava)
library("mailR")
library("writexl")
library(dplyr, warn.conflicts = FALSE)
library(salesforcer)
Sys.setenv(JAVA_HOME='C:\\Program Files (x86)\\Java\\jdk-11.0.16')
library(rJava)
library("mailR")
library("writexl")


# Importing Training data from SF


train_sf <- sprintf("SELECT Cycle__c, BSG_Membership__c, BSG__c, Business_Name__c, Country__c, Mod_1_Business_Group_Formation__c, Mod_2_Basic_Savings__c, Mod_3_Business_Savings_Group_BSG_Form__c, Mod_4_Leadership_and_BSG_Executive_Comm__c, Mod_5_Introduction_to_Conservation__c, Mod_6_BSG_Constitution__c, Mod_7_BSG_Record_Keeping__c, Mod_8_BSG_Lending_and_Loans__c, Mod_9_Business_Basics__c, Mod_10_Business_Selection__c, Mod_11_Business_Planning__c, Mod_12_Managing_a_Business__c, Mod_13_Business_Record_Keeping__c, Mod_14_Value_Addition_and_Marketing__c, Mod_15_Livestock_as_a_business_or_inves__c, Group_Session__c, Training_Name__c, BSG__r.BM_Cycle__c, BSG__r.Name, BSG__r.Business_Mentor__c, Household__r.BM_Cycle__c, Household__r.BSG_Name__c, Household__r.Id,Household__r.Business_Group_name__c, Household__r.Business_Owner_Age__c, Household__r.Country__c, Household__r.Cycle__c, Household__r.Office__c, Household__r.Training_Attendance__c, Household__r.Name, Household__r.Village__c FROM Group_Training_Detail__c WHERE Household__r.Cycle__c IN ('FY23C2','FY23C1','FY22C1', 'FY22C2','FY22C3')  AND Household__r.Household_Participation_Status__c = 'Active'")
train_sf <- sf_query(train_sf)
train_sf$training<-train_sf$Training_Name__c

my_soql <- sprintf("SELECT Id, BM_Cycle_ID__c, Name, Country__c, Cycle__c, Office2__c, Cycle__r.Cycle_ID__c, Cycle__r.Cycle_Name__c, Cycle__r.Fiscal_Year_Cycle__c FROM BM_Cycle__c WHERE Cycle__r.Fiscal_Year_Cycle__c IN  ('a302I000005imjMQAQ','a302I000002pgr0QAA')")
bmcycle <- sf_query(my_soql)

my_soql <-sprintf("SELECT BO_Name__c,Business_Group__r.Business_Group_Name__c , Business_Group__c, Household_ID__c, Household_Name__c, TW_Household__c, TW_Household__r.Id FROM Business_Group_Membership__c WHERE TW_Household__r.Cycle__c IN ('FY23C2','FY23C1','FY22C1', 'FY22C2','FY22C3')")
hhs<- sf_query(my_soql)


my_soql <-sprintf("SELECT Id, Fiscal_Year__c, Name FROM Fiscal_Year__c")
fiscal<- sf_query(my_soql)


train_md<-merge(train_sf,hhs,by.x="Household__r.Id",by.y="TW_Household__r.Id")

## Number of Training Sessions per BM
agg_df <- aggregate(Training_Name__c ~  Household__r.BM_Cycle__c + Country__c , data = train_sf, function(x) length(unique(x)))
agg_df<-subset(agg_df,Training_Name__c<8)
agg_df<-merge(agg_df, bmcycle, by.x = "Household__r.BM_Cycle__c", by.y = "Id", all.x = TRUE)
agg_df<-merge(agg_df, fiscal, by.x = "Cycle__r.Fiscal_Year_Cycle__c", by.y = "Id", all.x = TRUE)
write_xlsx(agg_df,"C:\\Users\\Hurchins Manua\\Documents\\DQR_template\\training_BM_.xlsx" )





###Number of unique Training sessions for Every Business Group
agg_tall <- aggregate(Training_Name__c ~  TW_Household__c + Country__c + Household__r.Office__c + Business_Group__r.Business_Group_Name__c+Household_Name__c+Cycle__c+Household__r.Cycle__c, data = train_md, function(x) length(unique(x)))
agg_tall2 <- aggregate(Training_Name__c ~  TW_Household__c + Country__c + Household__r.Office__c + Business_Group__r.Business_Group_Name__c, data = train_md, function(x) length(x))
agg_tall<-subset(agg_tall,Training_Name__c<6)
write_xlsx(agg_tall,"C:\\Users\\Hurchins Manua\\Documents\\DQR_template\\training_hh_.xlsx" )


#Getting Unique # of households in a country

unique_hhs <- n_distinct(train_sf$Household__r.Id)

#Sending Data to  Individual countries  less than 8Training Sessions per BM

country<-c("Kenya","Rwanda","Uganda")
for (i in country) {
  k<-i
  
  agg_df1<-subset(agg_df, ( Country__c.x==k)) 
  training_hhs<-nrow(agg_df1)
  write_xlsx(agg_df1,"C:\\Users\\Hurchins Manua\\Documents\\DQR_template\\training.xlsx" )
  
  
  if ((training_hhs>0) & k=="Kenya"){
    
    send.mail(from = "hnyakwara@villageenterprise.org",
              to = c("mhurchins@gmail.com","hnyakwara@villageenterprise.org","bodhiambo@villageenterprise.org"),
              subject = "BMs with Lesss than 8 training Sessions FY22-23",
              body = "Hi Billy, 
            The following BMS have less than 8 Training Sessions",
              smtp = list(host.name = "smtp.gmail.com", port = 465, user.name = "hnyakwara@villageenterprise.org", passwd = "Neville@11", ssl = TRUE),
              attach.files=c("./training.xlsx"),
              authenticate = TRUE,
              send = TRUE)
    
  }
  
  if ((training_hhs>0) & k=="Uganda"){
    
    send.mail(from = "hnyakwara@villageenterprise.org",
              to = c("mhurchins@gmail.com","hnyakwara@villageenterprise.org","edwinm@villageenterprise.org"),
              subject = "BMs with Less than 8 Training Sessions FY22-23",
              body = "Hi Edwin, 
             The following BMS have less than 8 Training Sessions",
              smtp = list(host.name = "smtp.gmail.com", port = 465, user.name = "hnyakwara@villageenterprise.org", passwd = "Neville@11", ssl = TRUE),
              attach.files=c("./training.xlsx"),
              authenticate = TRUE,
              send = TRUE)
    
    
  }
  
  if ((training_hhs>0) & k=="Rwanda"){
    
    send.mail(from = "hnyakwara@villageenterprise.org",
              to = c("mhurchins@gmail.com","hnyakwara@villageenterprise.org","mauricen@villageenterprise.org "),
              subject = "BMs with Less than 8 Training Sessions FY22-23",
              body = "Hi Maurice, 
             The following BMS have less than 8 Training Sessions",
              smtp = list(host.name = "smtp.gmail.com", port = 465, user.name = "hnyakwara@villageenterprise.org", passwd = "Neville@11", ssl = TRUE),
              attach.files=c("./training.xlsx"),
              authenticate = TRUE,
              send = TRUE)
    
  } 
  
}




####Sending Data to  Individual countries Less than 6 Training Sessions per HHs 


country<-c("Kenya","Rwanda","Uganda")
for (i in country) {
  k<-i
  
  agg_df2<-subset(agg_tall, ( Country__c==k)) 
  training_hhs<-nrow(agg_df2)
  write_xlsx(agg_df2,"C:\\Users\\Hurchins Manua\\Documents\\DQR_template\\training_h.xlsx" )
  
  
  if ((training_hhs>0) & k=="Kenya"){
    
    send.mail(from = "hnyakwara@villageenterprise.org",
              to = c("mhurchins@gmail.com","hnyakwara@villageenterprise.org","bodhiambo@villageenterprise.org"),
              subject = "HHS  with Lesss than 6 training Sessions FY22-23",
              body = "Hi Billy, 
            The following HHs have less than 6 Training Sessions",
              smtp = list(host.name = "smtp.gmail.com", port = 465, user.name = "hnyakwara@villageenterprise.org", passwd = "Neville@11", ssl = TRUE),
              attach.files=c("./training_h.xlsx"),
              authenticate = TRUE,
              send = TRUE)
    
  }
  
  if ((training_hhs>0) & k=="Uganda"){
    
    send.mail(from = "hnyakwara@villageenterprise.org",
              to = c("mhurchins@gmail.com","hnyakwara@villageenterprise.org","edwinm@villageenterprise.org"),
              subject = "HHS  with Lesss than 6 training Sessions FY22-23",
              body = "Hi Edwin, 
              The following HHs have less than 6 Training Sessions FY22-23",
              smtp = list(host.name = "smtp.gmail.com", port = 465, user.name = "hnyakwara@villageenterprise.org", passwd = "Neville@11", ssl = TRUE),
              attach.files=c("./training_h.xlsx"),
              authenticate = TRUE,
              send = TRUE)
    
    
  }
  
  if ((training_hhs>0) & k=="Rwanda"){
    
    send.mail(from = "hnyakwara@villageenterprise.org",
              to = c("mhurchins@gmail.com","hnyakwara@villageenterprise.org","mauricen@villageenterprise.org "),
              subject = "HHS  with Lesss than 6 training Sessions FY22-23",
              body = "Hi Maurice, 
             The following HHs have less than 6 Training Sessions",
              smtp = list(host.name = "smtp.gmail.com", port = 465, user.name = "hnyakwara@villageenterprise.org", passwd = "Neville@11", ssl = TRUE),
              attach.files=c("./training_h.xlsx"),
              authenticate = TRUE,
              send = TRUE)
    
  } 
  
}

































