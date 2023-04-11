### 17/11/2022
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
###Importing Data from System Prep Google sheet




#Linking G-sheet to FY

##install.packages('googlesheets4')
# Kenya
library(stringr)
library(tibble)
library('googlesheets4')
vill_kenya<-read_sheet("https://docs.google.com/spreadsheets/d/1gBkWwgrsqpgokjVy9Oe4i1aS2ltVafrrUm5itYVgP70/edit#gid=37501867","FY23C2 Nawiri" )
n<-ncol(vill_kenya)
long_vill<-vill_kenya[,4:10]
long_vill<-subset(long_vill,!is.na(long_vill$`BM Cycle name`))
long_vill<-data.frame(column_to_rownames(long_vill, var = "BM Cycle name"))
long_vill<- reshape(long_vill, idvar = "BM Cycle name", ids = row.names(long_vill),
                times = names(long_vill), timevar = "Characteristic",
                varying = list(names(long_vill)), direction = "long")

long_vill<-long_vill[!is.na(long_vill$Carryover.Village.1..Old.Villages.), ] 
long_vill<-rename(long_vill, villages =  Carryover.Village.1..Old.Villages. )
long_vill$`Village Assignment: Name`<-paste(long_vill$`BM Cycle name` ,long_vill$villages)
long_vill<-long_vill[order(long_vill$`BM Cycle name`),]
long_vill$`Village Assignment: Name`<-stringr::str_to_title(long_vill$`Village Assignment: Name`)
#write_xlsx(long_vill,"C:\\Users\\Hurchins Manua\\Documents\\DQR_template\\long_vill.xlsx" )



##Data from Salesforce

my_report_id <- "00O2I000007SyXzUAK"
vill_sf <- sf_run_report(my_report_id)
vill_saf_kenya<- subset(vill_sf, vill_sf$Country=="Uganda")
vill_saf_kenya<-vill_saf_kenya[order(vill_saf_kenya$`Village Assignment: Name`),]
vill_saf_kenya$`Village Assignment: Name`<-stringr::str_to_title(vill_saf_kenya$`Village Assignment: Name`)


#Checking villages 

tr<-as.data.frame(unique(long_vill$villages))
names(tr)[names(tr) == "unique(long_vill$villages)"] <- "var1"

sv<-as.data.frame(unique(vill_saf_kenya$`Village: Village Name`))
names(sv)[names(sv) ==  "unique(vill_saf_kenya$`Village: Village Name`)"] <- "var1"


dfv = sv %>% anti_join(tr, by = "var1" )
cou<-nrow(dfv)


#Checking Discrepancy 
df = long_vill %>% anti_join(vill_saf_kenya, by = "Village Assignment: Name" )
cou<-nrow(df)




if (cou > 0) {
  write_xlsx(df,"C:\\Users\\Hurchins Manua\\Documents\\DQR_template\\Mi_Sf_rwa.xlsx" )
}

#vill_saf_kenya$consistency<-RecordLinkage::levenshteinSim(long_vill$`Village Assignment: Name`, vill_saf_kenya$`Village Assignment: Name`)
#n_discp<-subset(vill_saf_kenya,subset = vill_saf_kenya$consistency<1)
#n_discp<-nrow(n_discp)
#write_xlsx(vill_saf_kenya,"C:\\Users\\Hurchins Manua\\Documents\\DQR_template\\saf.xlsx" )

# Checking Offered status Updates for last month

qual_status <- sprintf("SELECT BM_Cycle__c, Business_Owner_Full_Name__c, Gender__c, Country__c, Cycle__c, Household_Offered_status__c, Household_Participation_Status__c, Qualification_Status__c, Office__c, Name, BM_Cycle__r.BM_Cycle_ID__c, BM_Cycle__r.Name, BM_Cycle__r.Cycle__c, BM_Cycle__r.Project__c FROM Household__c WHERE BM_Cycle__r.Cycle__c = 'a2z2I000001dwd7QAA'")
qual_status <- sf_query(qual_status)
update_list<-subset(qual_status,qual_status$Household_Offered_status__c=="Waiting Program Offer")
up_no<-nrow(update_list)

if (up_no > 0) {
  write_xlsx(update_list,"C:\\Users\\Hurchins Manua\\Documents\\DQR_template\\update_offered_status.xlsx" )
}


if (up_no > 0) {
  
  send.mail(from = "hnyakwara@villageenterprise.org",
            to = c("mhurchins@gmail.com","hnyakwara@villageenterprise.org","bodhiambo@villageenterprise.org","edwinm@villageenterprise.org","mauricen@villageenterprise.org "),
            subject = "Last Cycle BM Offered Status Updates",
            body = "Hi Team, 
          Please update on the Offered Status for the BOs attached",
            smtp = list(host.name = "smtp.gmail.com", port = 465, user.name = "hnyakwara@villageenterprise.org", passwd = "Neville@11", ssl = TRUE),
            attach.files=c("./update_offered_status.xlsx"),
            authenticate = TRUE,
            send = TRUE)  
  
}

# Checking for record Assignment amongst FAs and BMs

Assignment <- sprintf("SELECT gfsurveys__Account__c, gfsurveys__Alias__c, gfsurveys__Assigned_Public_Group_IDs__c, gfsurveys__Assigned_Public_Group_Names__c, gfsurveys__Assigned_Record_Ids__c, gfsurveys__Assigned_Record_Objects__c, gfsurveys__Contact__c, gfsurveys__First_Name__c, gfsurveys__Manager__c, gfsurveys__Mobile__c, Name, gfsurveys__Role__c, gfsurveys__Username__c, gfsurveys__Contact__r.FirstName, gfsurveys__Contact__r.Name, gfsurveys__Contact__r.LastName, gfsurveys__Contact__r.stayclassy__Middle_Name__c, gfsurveys__Contact__r.gfsurveys__mobilesurveys_Username__c FROM gfsurveys__TaroWorks_Mobile_User__c WHERE gfsurveys__Active__c = true")
Assignment <- sf_query(Assignment)


## This is to convert Assigned records to ensure all BM cycles are assigned for FAs and BAs

my_soql <- sprintf("SELECT Id, gfsurveys__AssociatedIds__c, gfsurveys__Contact__c, CreatedById, CreatedDate, IsDeleted, gfsurveys__Instance__c, LastModifiedById, LastModifiedDate, gfsurveys__NumberOfRecords__c, gfsurveys__SObjectApiName__c, gfsurveys__SObjectFieldApiName__c, Name, SystemModstamp, gfsurveys__UniqueKey__c FROM gfsurveys__SObjectContactAssociation__c")
queried_records <- sf_query(my_soql)

Assignment = merge(x = Assignment, y = queried_records, by= "gfsurveys__Contact__c",
           all.x = TRUE)
Assignment<-subset(Assignment,!is.na(Assignment$gfsurveys__Assigned_Public_Group_IDs__c))


Assignment<-Assignment %>% filter(Assignment$LastModifiedDate > '2022-10-01')

## Importing BM cycle data 
my_bmc <- sprintf("SELECT Id, Name, BM__c, Country__c, BM__r.Id, Cycle__r.Id, Cycle__r.Cycle_Name__c FROM BM_Cycle__c WHERE Cycle__c = 'a2z2I000001jVmsQAE'")
queried_bmc <- sf_query(my_bmc)
Assignment<-merge(x = Assignment, y = queried_bmc, by.x = "gfsurveys__Contact__c",by.y = "BM__r.Id",
                  all = FALSE)

###Checking Assigned Contacts & BM cycles 
Assignment1<-subset(Assignment,select = c(Name,gfsurveys__AssociatedIds__c ))

Assignment1<-tidyr::separate(
  data = Assignment1,
  col = gfsurveys__AssociatedIds__c,
  sep = ",",
  into = c("Object1","Object2","Object3","Object4","Object5","Object6","Object7","Object8","Object9","Object10","Object11","Object12"),
  remove = FALSE
)





####









