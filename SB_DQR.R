rm(list=ls())
Sys.setenv(JAVA_HOME='C:\\Program Files (x86)\\Java\\jdk-11.0.16')
library(rJava)
library("mailR")
library("writexl")
library(dplyr, warn.conflicts = FALSE)
library(salesforcer)
sf_auth()
FY<-c("FY23C1")



# Pulling data from a report 

my_soql <- sprintf("SELECT Id, Name, BM_Cycle__r.Country__c, BM_Cycle__r.Cycle__c, Business_Group_Name__c, Business_Health_1__c, Country__c, Grant_Size__c, Office__c,(SELECT  Business_Group_Name__c FROM SB_Grants__r) FROM Business_Group__c WHERE BM_Cycle__c = 'a0h2I00000CB9G0QAL'") 
queried_records <- sf_query(my_soql)
countrys<-data.frame(table(queried_records$BM_Cycle__r.Country__c))
countrys = as.character(countrys$Var1);
print(countrys)