#Linking G-sheet from EOC to Look at the Validity of Information

##install.packages('googlesheets4')
# Kenya
library(stringr)
library(tibble)
library('googlesheets4')

FY<-c("FY23C1")
eoc<-read_sheet("https://docs.google.com/spreadsheets/d/1xWt4eRO7OCcLCiUs6TDys9QWRBfgtceuLrTawcYhsF8/edit#gid=2022529918","FY23 EOC" )


# Reformatting the Google Sheet Template

N<-5
eoc<-tail(eoc,-N)

library(dplyr)
eoc<-eoc %>%
slice(-c(14))


Nt<-4
eoc<-head(eoc,-Nt)
eoc[14,1] <-c("Crops")
eoc[15,1] <-c("Retail")
eoc[16,1] <-c("Skilled")
eoc[17,1] <-c("Service")
eoc[18,1] <-c("Livestock")

eoc<-eoc[,1:14]
eoc<-eoc[,-2]   
eoc$...3<-as.character(eoc$...3)
eoc$...4<-as.character(eoc$...4)
eoc$...5<-as.character(eoc$...5)
eoc$...6<-as.character(eoc$...6)

 

eoc<-eoc[,-2]
eoc<-eoc[,-5]
eoc<-eoc[,-8]

colnames(eoc) <- eoc[1, ] 
eoc<-eoc[-1,]
eoc<-eoc[-1,]
eoc<-eoc[-3,]
eoc<-eoc[-9,]
eoc<-eoc[-19,]
eoc<-eoc[-29,]

## Rename to identify each cycle
names(eoc)[2] <- "Kenya_C1"
names(eoc)[3] <- "Uganda_C1"
names(eoc)[4] <- "Rwanda_C1"
names(eoc)[5] <- "Kenya_C2"
names(eoc)[6] <- "Uganda_C2"
names(eoc)[7] <- "Rwanda_C2"
names(eoc)[8] <- "Kenya_C3"
names(eoc)[9] <- "Uganda_C3"
names(eoc)[10] <- "Rwanda_C3"


######EOC Analysis for cycle Statistics #############

###Calculating and getting data from ###################

######## Data from Sales force ###################
#### Start with Cycle Name#######################
## Systems Prep
Sys.setenv(JAVA_HOME='C:\\Program Files (x86)\\Java\\jdk-11.0.16')
library(rJava)
library("mailR")
library("writexl")

setwd("C:\\Users\\Hurchins Manua\\Documents\\DQR_template")

library(dplyr, warn.conflicts = FALSE)
library(salesforcer)
library(RecordLinkage)
library(dplyr)
library(tidyr)
library("writexl")

my_soql <- sprintf("SELECT Id, BM_Cycle_ID__c, Name, Country__c, Cycle__c, Office2__c, Cycle__r.Cycle_ID__c, Cycle__r.Cycle_Name__c, Cycle__r.Fiscal_Year_Cycle__c FROM BM_Cycle__c WHERE Cycle__r.Fiscal_Year_Cycle__c = 'a302I000005imjMQAQ'")
bmcycle <- sf_query(my_soql)


my_soql1 <- sprintf("SELECT Business_Group__c, Business_Group_Name__c,Business_Group__r.BM_Cycle__c , SB_Funding_Status2__c, Name, Business_Group__r.Name, Business_Group__r.Business_Group_Name__c, Business_Group__r.Country__c, Business_Group__r.Office__c FROM SB_Grant__c WHERE CreatedDate = LAST_N_DAYS:180")
sbgrants<- sf_query(my_soql1)

my_soql2 <- sprintf("SELECT Business_Group__c, Business_Group_Name__c,Business_Group__r.BM_Cycle__c , SB_Funding_Status2__c, Name, Business_Group__r.Name, Business_Group__r.Business_Group_Name__c, Business_Group__r.Country__c, Business_Group__r.Office__c FROM SB_Grant__c WHERE CreatedDate = LAST_N_DAYS:180")
prgrants<- sf_query(my_soql2)



my_soql3 <- sprintf("SELECT Biz_Cash__c, Biz_Expenses__c, Biz_Inputs__c, Biz_Inventory__c, Biz_Profits__c, Biz_Revenue__c, Business_type_group_currently_operating__c, Biz_Value__c, BOs_Dropped__c, Business_Group__c, Business_Group_Name__c, Proportion_PR_Used__c, Proportion_SB_Used__c, Records_Kept__c, SB_Actual_Biz_Type__c, Why_Survey_Not_Completed__c, Why_Survey_Not_Conducted__c, Business_Group__r.Name, Business_Group__r.Business_Group_Name__c, Business_Group__r.Grant_Size__c, Business_Group__r.Group_Size_at_PR__c, Business_Group__r.Group_Size_at_SB__c, Business_Group__r.PR_Business_Type__c, Business_Group__r.PR_Funding_Status2__c, Business_Group__r.SB_Business_Type__c, Business_Group__r.SB_Funding_Status2__c FROM Business_Progress_Survey__c WHERE CreatedDate = LAST_N_DAYS:180")
Biz_exits<-sf_query(my_soql3)

my_soql4<-sprintf("SELECT Business_Group__c, Business_Group_Name__c, PR_Business_Type__c, PR_Funding_Status2__c, Name FROM PR_Grant__c WHERE CreatedDate = LAST_N_DAYS:180")
prgrants<-sf_query(my_soql4)






