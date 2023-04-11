
##########     DQR  ###########################


#This is DQR for Targeting  / Registration 




## Start by having all Libraries 
# Clearing the Environment 

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

my_soql <- sprintf("SELECT Id, Business_Owner_Age__c,BM_Cycle__r.Project__c, Business_Owner_Date_of_Birth__c, Gender__c, Household_ID__c, Household_Offered_status__c, Household_Participation_Status__c, Qualification_Status__c, BM_Cycle__r.Name, BM_Cycle__r.Country__c, BM_Cycle__r.Cycle__c FROM Household__c WHERE (BM_Cycle__r.Cycle__c = 'a2z2I000001dwd7QAA' OR BM_Cycle__r.Cycle__c = 'a2z2I000001jVmsQAE')")
queried_records <- sf_query(my_soql)
queried_records<-subset(queried_records,queried_records$BM_Cycle__r.Cycle__c=="a2z2I000001jVmsQAE")
countrys<-data.frame(table(queried_records$BM_Cycle__r.Country__c))
countrys = as.character(countrys$Var1);
print(countrys)

# Checking for duplicates 
# check  Missing BM cycle

BM_missing<-subset(queried_records,is.na(queried_records$BM_Cycle__r.Cycle__c) )
bmm<-nrow(BM_missing)


#Check Missing Age

Age_missing<-subset(queried_records,is.na(queried_records$Business_Owner_Age__c) &  queried_records$Household_Participation_Status__c!="Inactive - Declined Never Enrolled" & queried_records$Household_Participation_Status__c!="Inactive - Dropped" )
Am<-nrow(Age_missing)



#DQR  BO Registration 
#Missing Qualification Status

table(queried_records$Household_Participation_Status__c)
Qualification_status_Missing<- subset(queried_records,is.na(queried_records$Qualification_Status__c))
qsm<-nrow(Qualification_status_Missing)
write_xlsx(Qualification_status_Missing,"C:\\Users\\Hurchins Manua\\Documents\\DQR_template\\Qualification_mi_participation.xlsx" )



# Qualified Missing offered Status

Qualified_mi_offered<- subset(queried_records,queried_records$Qualification_Status__c=="Qualified" & is.na(queried_records$Household_Offered_status__c))
qmo<-(nrow(Qualified_mi_offered))
write_xlsx(Qualification_status_Missing,"C:\\Users\\Hurchins Manua\\Documents\\DQR_template\\Qualified_mi_offered.xlsx" )

# Not Qualified but Has Offered Program 

not_Qualified_offered<- subset(queried_records,queried_records$Qualification_Status__c!="Qualified" & !is.na(queried_records$Household_Offered_status__c))
nqo<-(nrow(not_Qualified_offered))
write_xlsx(Qualification_status_Missing,"C:\\Users\\Hurchins Manua\\Documents\\DQR_template\\not_Qualified_offered.xlsx" )

# Not Offered But Has participation Status

not_active_offered<-subset(queried_records,queried_records$Household_Offered_status__c!="Offered Program" & !is.na(queried_records$Household_Participation_Status__c))
na<-(nrow(not_active_offered))
write_xlsx(not_active_offered,"C:\\Users\\Hurchins Manua\\Documents\\DQR_template\\not_active_offered.xlsx" )

# Offered Missing Participation Status

offered_mi_participation<-subset(queried_records,queried_records$Household_Offered_status__c=="Offered Program" & is.na(queried_records$Household_Participation_Status__c))
omp=(nrow(offered_mi_participation))
write_xlsx(offered_mi_participation,"C:\\Users\\Hurchins Manua\\Documents\\DQR_template\\offered_mi_participation.xlsx" )


#########    Looping through the Country's Emailing List of      #####################
########     Managers with the information sending information on  ###################


######  Missing Qualification Status  #############

Sys.setenv(JAVA_HOME='C:\\Program Files (x86)\\Java\\jdk-11.0.16')
library(rJava)
library("mailR")
library("writexl")

setwd("C:\\Users\\Hurchins Manua\\Documents\\DQR_template")
#for (x in countrys){

country<-c("Kenya","Rwanda","Uganda")
for (i in country) {
k<-i
Qualification_status_Missing<- subset(queried_records,is.na(queried_records$Qualification_Status__c)& queried_records$BM_Cycle__r.Country__c==k)
qsm<-nrow(Qualification_status_Missing)  
write_xlsx(Qualification_status_Missing,"C:\\Users\\Hurchins Manua\\Documents\\DQR_template\\Qualification_mi_participation.xlsx" )

if ((qsm>0) & k=="Kenya"){

send.mail(from = "hnyakwara@villageenterprise.org",
          to = c("mhurchins@gmail.com","hnyakwara@villageenterprise.org","bodhiambo@villageenterprise.org"),
          subject = "Qualification Status Missing ",
          body = "Hi Team, 
          Please update on the Qualification Status for the BOs attached",
          smtp = list(host.name = "smtp.gmail.com", port = 465, user.name = "hnyakwara@villageenterprise.org", passwd = "Neville#11", ssl = TRUE),
          attach.files=c("./Qualification_mi_participation.xlsx"),
          authenticate = TRUE,
          send = TRUE)
  
  }
 
if ((qsm>0) & k=="Uganda"){
  
  send.mail(from = "hnyakwara@villageenterprise.org",
            to = c("mhurchins@gmail.com","hnyakwara@villageenterprise.org","edwinm@villageenterprise.org"),
            subject = "Qualification Status Missing ",
            body = "Hi Team, 
          Please update on the Qualification Status for the BOs attached",
            smtp = list(host.name = "smtp.gmail.com", port = 465, user.name = "hnyakwara@villageenterprise.org", passwd = "Neville#11", ssl = TRUE),
            attach.files=c("./Qualification_mi_participation.xlsx"),
            authenticate = TRUE,
            send = TRUE)
  
}
if ((qsm>0) & k=="Rwanda"){
  
  send.mail(from = "hnyakwara@villageenterprise.org",
            to = c("mhurchins@gmail.com","hnyakwara@villageenterprise.org","mauricen@villageenterprise.org"),
            subject = "Qualification Status Missing ",
            body = "Hi Team, 
          Please update on the Qualification Status for the BOs attached",
            smtp = list(host.name = "smtp.gmail.com", port = 465, user.name = "hnyakwara@villageenterprise.org", passwd = "Neville#11", ssl = TRUE),
            attach.files=c("./Qualification_mi_participation.xlsx"),
            authenticate = TRUE,
            send = TRUE)
 } 

}


setwd("C:\\Users\\Hurchins Manua\\Documents\\DQR_template")

#######################################################################################
#######################################################################################

# Qualified Missing offered Status

country<-c("Kenya","Rwanda","Uganda")
for (i in country) {
  k<-i
  Qualified_mi_offered<- subset(queried_records,queried_records$Qualification_Status__c=="Qualified" & is.na(queried_records$Household_Offered_status__c) & queried_records$BM_Cycle__r.Country__c==k)
  qmo<-nrow(Qualified_mi_offered)  
  write_xlsx(Qualified_mi_offered,"C:\\Users\\Hurchins Manua\\Documents\\DQR_template\\Qualified_mi_offered.xlsx" )
  
  if ((qmo>0) & k=="Kenya"){
    
    send.mail(from = "hnyakwara@villageenterprise.org",
              to = c("mhurchins@gmail.com","hnyakwara@villageenterprise.org","bodhiambo@villageenterprise.org"),
              subject = "Qualification Status Missing ",
              body = "Hi Team, 
          Please update on the Qualification Status for the BOs attached",
              smtp = list(host.name = "smtp.gmail.com", port = 465, user.name = "hnyakwara@villageenterprise.org", passwd = "Neville@11", ssl = TRUE),
              attach.files=c("./Qualified_mi_offered.xlsx"),
              authenticate = TRUE,
              send = TRUE)
    
  }
  
  if ((qmo>0) & k=="Uganda"){
    
    send.mail(from = "hnyakwara@villageenterprise.org",
              to = c("mhurchins@gmail.com","hnyakwara@villageenterprise.org","edwinm@villageenterprise.org"),
              subject = "Qualification Status Missing ",
              body = "Hi Team, 
          Please update on the Qualification Status for the BOs attached",
              smtp = list(host.name = "smtp.gmail.com", port = 465, user.name = "hnyakwara@villageenterprise.org", passwd = "Neville@11", ssl = TRUE),
              attach.files=c("./Qualified_mi_offered.xlsx"),
              authenticate = TRUE,
              send = TRUE)
    
  }
  if ((qmo>0) & k=="Rwanda"){
    
    send.mail(from = "hnyakwara@villageenterprise.org",
              to = c("mhurchins@gmail.com","hnyakwara@villageenterprise.org","mauricen@villageenterprise.org"),
              subject = "Qualification Status Missing ",
              body = "Hi Team, 
          Please update on the Qualification Status for the BOs attached",
              smtp = list(host.name = "smtp.gmail.com", port = 465, user.name = "hnyakwara@villageenterprise.org", passwd = "Neville@11", ssl = TRUE),
              attach.files=c("./Qualified_mi_offered.xlsx"),
              authenticate = TRUE,
              send = TRUE)
  } 
  
}

# Not Qualified but Has Offered Program 

country<-c("Kenya","Rwanda","Uganda")
for (i in country) {
  k<-i

  not_Qualified_offered<- subset(queried_records,queried_records$Qualification_Status__c!="Qualified" & !is.na(queried_records$Household_Offered_status__c)  & queried_records$BM_Cycle__r.Country__c==k)
  nqo<-(nrow(not_Qualified_offered))
  write_xlsx(Qualification_status_Missing,"C:\\Users\\Hurchins Manua\\Documents\\DQR_template\\not_Qualified_offered.xlsx" )
  
  
  if ((nqo>0) & k=="Kenya"){
    
    send.mail(from = "hnyakwara@villageenterprise.org",
              to = c("mhurchins@gmail.com","hnyakwara@villageenterprise.org","bodhiambo@villageenterprise.org"),
              subject = "Qualification Status Missing ",
              body = "Hi Team, 
          Please update on the Qualification Status for the BOs attached",
              smtp = list(host.name = "smtp.gmail.com", port = 465, user.name = "hnyakwara@villageenterprise.org", passwd = "Neville@11", ssl = TRUE),
              attach.files=c("./not_Qualified_offered.xlsx"),
              authenticate = TRUE,
              send = TRUE)
    
  }
  
  if ((nqo>0) & k=="Uganda"){
    
    send.mail(from = "hnyakwara@villageenterprise.org",
              to = c("mhurchins@gmail.com","hnyakwara@villageenterprise.org","edwinm@villageenterprise.org"),
              subject = "Qualification Status Missing ",
              body = "Hi Team, 
          Please update on the Qualification Status for the BOs attached",
              smtp = list(host.name = "smtp.gmail.com", port = 465, user.name = "hnyakwara@villageenterprise.org", passwd = "Neville@11", ssl = TRUE),
              attach.files=c("./not_Qualified_offered.xlsx"),
              authenticate = TRUE,
              send = TRUE)
    
  }
  if ((nqo>0) & k=="Rwanda"){
    
    send.mail(from = "hnyakwara@villageenterprise.org",
              to = c("mhurchins@gmail.com","hnyakwara@villageenterprise.org","mauricen@villageenterprise.org"),
              subject = "Qualification Status Missing ",
              body = "Hi Team, 
          Please update on the Qualification Status for the BOs attached",
              smtp = list(host.name = "smtp.gmail.com", port = 465, user.name = "hnyakwara@villageenterprise.org", passwd = "Neville@11", ssl = TRUE),
              attach.files=c("./not_Qualified_offered.xlsx"),
              authenticate = TRUE,
              send = TRUE)
  } 
  
}




# Not Offered But Has participation Status


country<-c("Kenya","Rwanda","Uganda")
for (i in country) {
  k<-i
  
  
  not_active_offered<-subset(queried_records,queried_records$Household_Offered_status__c!="Offered Program" & !is.na(queried_records$Household_Participation_Status__c) & queried_records$BM_Cycle__r.Country__c==k)
  na<-(nrow(not_active_offered))
  write_xlsx(not_active_offered,"C:\\Users\\Hurchins Manua\\Documents\\DQR_template\\not_active_offered.xlsx" )
  
  
  if ((na>0) & k=="Kenya"){
    
    send.mail(from = "hnyakwara@villageenterprise.org",
              to = c("mhurchins@gmail.com","hnyakwara@villageenterprise.org","bodhiambo@villageenterprise.org"),
              subject = "Not Offered But Has participation Status",
              body = "Hi Team, 
          Please update on this HHs Not offered program but has Participation Status ",
              smtp = list(host.name = "smtp.gmail.com", port = 465, user.name = "hnyakwara@villageenterprise.org", passwd = "Neville@11", ssl = TRUE),
              attach.files=c("./not_active_offered.xlsx"),
              authenticate = TRUE,
              send = TRUE)
    
  }
  
  if ((na>0) & k=="Uganda"){
    
    send.mail(from = "hnyakwara@villageenterprise.org",
              to = c("mhurchins@gmail.com","hnyakwara@villageenterprise.org","edwinm@villageenterprise.org"),
              subject = "Not Offered But Has participation Status ",
              body =  "Hi Team, 
          Please update on this HHs Not offered program but has Participation Status ",
              smtp = list(host.name = "smtp.gmail.com", port = 465, user.name = "hnyakwara@villageenterprise.org", passwd = "Neville@11", ssl = TRUE),
              attach.files=c("./not_active_offered.xlsx"),
              authenticate = TRUE,
              send = TRUE)
    
  }
  
  if ((na>0) & k=="Rwanda"){
    
    send.mail(from = "hnyakwara@villageenterprise.org",
              to = c("mhurchins@gmail.com","hnyakwara@villageenterprise.org","mauricen@villageenterprise.org"),
              subject = "Not Offered But Has participation Status",
              body =  "Hi Team, 
          Please update on this HHs Not offered program but has Participation Status ",
              smtp = list(host.name = "smtp.gmail.com", port = 465, user.name = "hnyakwara@villageenterprise.org", passwd = "Neville@11", ssl = TRUE),
              attach.files=c("./not_active_offered.xlsx"),
              authenticate = TRUE,
              send = TRUE)
  } 
  
}

# Offered Missing Participation Status


country<-c("Kenya","Rwanda","Uganda")
for (i in country) {
  k<-i
  
  offered_mi_participation<-subset(queried_records,queried_records$Household_Offered_status__c=="Offered Program" & is.na(queried_records$Household_Participation_Status__c)  & queried_records$BM_Cycle__r.Country__c==k)
  omp=(nrow(offered_mi_participation))
  write_xlsx(offered_mi_participation,"C:\\Users\\Hurchins Manua\\Documents\\DQR_template\\offered_mi_participation.xlsx" )
  
  
  
  
  if ((omp>0) & k=="Kenya"){
    
    send.mail(from = "hnyakwara@villageenterprise.org",
              to = c("mhurchins@gmail.com","hnyakwara@villageenterprise.org","bodhiambo@villageenterprise.org"),
              subject = "Qualification Status Missing ",
              body = "Hi Team, 
          Please update on the Qualification Status for the BOs attached",
              smtp = list(host.name = "smtp.gmail.com", port = 465, user.name = "hnyakwara@villageenterprise.org", passwd = "Neville@11", ssl = TRUE),
              attach.files=c("./offered_mi_participation.xlsx"),
              authenticate = TRUE,
              send = TRUE)
    
  }
  
  if ((omp>0) & k=="Uganda"){
    
    send.mail(from = "hnyakwara@villageenterprise.org",
              to = c("mhurchins@gmail.com","hnyakwara@villageenterprise.org","edwinm@villageenterprise.org"),
              subject = "Qualification Status Missing ",
              body = "Hi Team, 
          Please update on the Qualification Status for the BOs attached",
              smtp = list(host.name = "smtp.gmail.com", port = 465, user.name = "hnyakwara@villageenterprise.org", passwd = "Neville@11", ssl = TRUE),
              attach.files=c("./offered_mi_participation.xlsx"),
              authenticate = TRUE,
              send = TRUE)
    
  }
  
  if ((omp>0) & k=="Rwanda"){
    
    send.mail(from = "hnyakwara@villageenterprise.org",
              to = c("mhurchins@gmail.com","hnyakwara@villageenterprise.org","mauricen@villageenterprise.org"),
              subject = "Qualification Status Missing ",
              body = "Hi Team, 
          Please update on the Qualification Status for the BOs attached",
              smtp = list(host.name = "smtp.gmail.com", port = 465, user.name = "hnyakwara@villageenterprise.org", passwd = "Neville@11", ssl = TRUE),
              attach.files=c("./offered_mi_participation.xlsx"),
              authenticate = TRUE,
              send = TRUE)
  } 
  
}



# Missing Age


country<-c("Kenya","Rwanda","Uganda")
for (i in country) {
  k<-i
  

  Age_missing<-subset(queried_records,is.na(queried_records$Business_Owner_Age__c) &  queried_records$Household_Participation_Status__c!="Inactive - Declined Never Enrolled" & queried_records$Household_Participation_Status__c!="Inactive - Dropped"  & queried_records$BM_Cycle__r.Country__c==k)
  Am<-nrow(Age_missing)
  write_xlsx(offered_mi_participation,"C:\\Users\\Hurchins Manua\\Documents\\DQR_template\\Age_missing.xlsx" )
  
  
  
  if ((Am>0) & k=="Kenya"){
    
    send.mail(from = "hnyakwara@villageenterprise.org",
              to = c("mhurchins@gmail.com","hnyakwara@villageenterprise.org","bodhiambo@villageenterprise.org"),
              subject = "Qualification Status Missing ",
              body = "Hi Team, 
          Please update on the Qualification Status for the BOs attached",
              smtp = list(host.name = "smtp.gmail.com", port = 465, user.name = "hnyakwara@villageenterprise.org", passwd = "Neville@11", ssl = TRUE),
              attach.files=c("./Age_missing.xlsx"),
              authenticate = TRUE,
              send = TRUE)
    
  }
  
  if ((Am>0) & k=="Uganda"){
    
    send.mail(from = "hnyakwara@villageenterprise.org",
              to = c("mhurchins@gmail.com","hnyakwara@villageenterprise.org","edwinm@villageenterprise.org"),
              subject = "Qualification Status Missing ",
              body = "Hi Team, 
          Please update on the Qualification Status for the BOs attached",
              smtp = list(host.name = "smtp.gmail.com", port = 465, user.name = "hnyakwara@villageenterprise.org", passwd = "Neville@11", ssl = TRUE),
              attach.files=c("./Age_missing.xlsx"),
              authenticate = TRUE,
              send = TRUE)
    
  }
  
  if ((Am>0) & k=="Rwanda"){
    
    send.mail(from = "hnyakwara@villageenterprise.org",
              to = c("mhurchins@gmail.com","hnyakwara@villageenterprise.org","mauricen@villageenterprise.org"),
              subject = "Qualification Status Missing ",
              body = "Hi Team, 
          Please update on the Qualification Status for the BOs attached",
              smtp = list(host.name = "smtp.gmail.com", port = 465, user.name = "hnyakwara@villageenterprise.org", passwd = "Neville@11", ssl = TRUE),
              attach.files=c("./Age_missing.xlsx"),
              authenticate = TRUE,
              send = TRUE)
  } 
  
}



###############################################################################
##################### BG _ BSG Registration DQR################################
###############################################################################




## Checking duplicate HHs in BSG

BSG_sf <- sprintf("SELECT Role_in_BSG_Group__c, TW_Household__c, CreatedBy.Country, BSG__r.BM_Cycle__c, BSG__r.BSG_ID__c, BSG__r.Name, BSG__r.Country__c, TW_Household__r.Business_Group_name__c, TW_Household__r.BMC_Field_Associate__c, TW_Household__r.Household_ID__c, TW_Household__r.Household_Participation_Status__c, TW_Household__r.Name FROM Group_Membership__c WHERE TW_Household__r.Cycle__c = 'FY23C2'")
BSG_sf <- sf_query(BSG_sf)


#BSG_dup<-BSG_sf[,duplicated(BSG_sf$TW_Household__r.Household_ID__c)]
BSG_dup<-BSG_sf[which(duplicated(BSG_sf[, c("TW_Household__r.Household_ID__c" )])), ]
write_xlsx(BSG_sf,"C:\\Users\\Hurchins Manua\\Documents\\DQR_template\\BSG_dup.xlsx" )

mi_BSG_role<-subset(BSG_sf, is.na(BSG_sf$Role_in_BSG_Group__c)&  BSG_sf$TW_Household__r.Household_Participation_Status__c=="Active")
nm_bsg<-nrow(mi_BSG_role)



## Ensuring each BM has almost 2 BSGs

BSG_bm <- aggregate(BSG__r.BSG_ID__c ~ BSG__r.BM_Cycle__c + BSG__r.Country__c , data = BSG_sf, FUN = function(x) length(unique(x)))
BSG_bm<-subset(BSG_bm,BSG__r.BSG_ID__c>2)
no_BSG_bm<-nrow(BSG_bm)








if (nm_bsg > 0) {
    write_xlsx(mi_BSG_role,"C:\\Users\\Hurchins Manua\\Documents\\DQR_template\\Update_BSG_Roles.xlsx" )
  }  


#####Updating and sending out Emails ##########################
###############################################################


# Update Missing Roles in BSG 


country<-c("Kenya","Rwanda","Uganda")
for (i in country) {
  k<-i
  
 # BSG_dup_hhid<-BSG_sf[duplicated(BSG_sf$TW_Household__c),]
  BSG_sf<-as.data.frame(BSG_sf)
  BSG_dup_names<-BSG_sf[which(duplicated(BSG_sf[, c("BSG__r.BM_Cycle__c" , "BSG__r.Name")])), ]
  
  mi_BSG_role<-subset(BSG_sf, is.na(BSG_sf$Role_in_BSG_Group__c) &  BSG_sf$TW_Household__r.Household_Participation_Status__c=="Active" & BSG_sf$BSG__r.Country__c==k)
  nm_bsg<-nrow(mi_BSG_role)
  write_xlsx(mi_BSG_role,"C:\\Users\\Hurchins Manua\\Documents\\DQR_template\\Update_BSG_Roles.xlsx" )
  
  
  if ((nm_bsg>0) & k=="Kenya"){
    
    send.mail(from = "hnyakwara@villageenterprise.org",
              to = c("mhurchins@gmail.com","hnyakwara@villageenterprise.org","bodhiambo@villageenterprise.org"),
              subject = "Missing Role in BSG",
              body = "Hi Team, 
          Please update on the BSG Roles for the BOs attached",
              smtp = list(host.name = "smtp.gmail.com", port = 465, user.name = "hnyakwara@villageenterprise.org", passwd = "Neville@11", ssl = TRUE),
              attach.files=c("./Update_BSG_Roles.xlsx"),
              authenticate = TRUE,
              send = TRUE)
    
  }
  
  if ((nm_bsg>0) & k=="Uganda"){
    
    send.mail(from = "hnyakwara@villageenterprise.org",
              to = c("mhurchins@gmail.com","hnyakwara@villageenterprise.org","edwinm@villageenterprise.org"),
              subject = "Missing Role in BSG",
              body = "Hi Team, 
          Please update on the BSG Roles for the BOs attached",
              smtp = list(host.name = "smtp.gmail.com", port = 465, user.name = "hnyakwara@villageenterprise.org", passwd = "Neville@11", ssl = TRUE),
              attach.files=c("./Update_BSG_Roles.xlsx"),
              authenticate = TRUE,
              send = TRUE)
    
    
  }
  
  if ((nm_bsg>0) & k=="Rwanda"){
    
    send.mail(from = "hnyakwara@villageenterprise.org",
              to = c("mhurchins@gmail.com","hnyakwara@villageenterprise.org","mauricen@villageenterprise.org "),
              subject = "Missing Role in BSG",
              body = "Hi Team, 
          Please update on the BSG Roles for the BOs attached",
              smtp = list(host.name = "smtp.gmail.com", port = 465, user.name = "hnyakwara@villageenterprise.org", passwd = "Neville@11", ssl = TRUE),
              attach.files=c("./Update_BSG_Roles.xlsx"),
              authenticate = TRUE,
              send = TRUE)
    
  } 
  
}


### BM with  more than 3 BSGs


country<-c("Kenya","Rwanda","Uganda")
for (i in country) {
  k<-i

  BSG_bm <- aggregate(BSG__r.BSG_ID__c ~ BSG__r.BM_Cycle__c + BSG__r.Country__c , data = BSG_sf, FUN = function(x) length(unique(x)))
  BSG_bm<-subset(BSG_bm, (BSG__r.BSG_ID__c>2  & BSG__r.Country__c==k)) 
  no_BSG_bm<-nrow(BSG_bm)
  write_xlsx(BSG_bm,"C:\\Users\\Hurchins Manua\\Documents\\DQR_template\\BSG_bm.xlsx" )
  
  
  if ((no_BSG_bm>0) & k=="Kenya"){
    
    send.mail(from = "hnyakwara@villageenterprise.org",
              to = c("mhurchins@gmail.com","hnyakwara@villageenterprise.org","bodhiambo@villageenterprise.org"),
              subject = "BMs with more than 2 BSG",
              body = "Hi Team, 
          The following BMs have more than 2 BSGs",
              smtp = list(host.name = "smtp.gmail.com", port = 465, user.name = "hnyakwara@villageenterprise.org", passwd = "Neville@11", ssl = TRUE),
              attach.files=c("./BSG_bm.xlsx"),
              authenticate = TRUE,
              send = TRUE)
    
  }
  
  if ((no_BSG_bm>0) & k=="Uganda"){
    
    send.mail(from = "hnyakwara@villageenterprise.org",
              to = c("mhurchins@gmail.com","hnyakwara@villageenterprise.org","edwinm@villageenterprise.org"),
              subject = "BMs with more than 2 BSG",
              body = "Hi Team, 
           The following BMs have more than 2 BSGs",
              smtp = list(host.name = "smtp.gmail.com", port = 465, user.name = "hnyakwara@villageenterprise.org", passwd = "Neville@11", ssl = TRUE),
              attach.files=c("./BSG_bm.xlsx"),
              authenticate = TRUE,
              send = TRUE)
    
    
  }
  
  if ((no_BSG_bm>0) & k=="Rwanda"){
    
    send.mail(from = "hnyakwara@villageenterprise.org",
              to = c("mhurchins@gmail.com","hnyakwara@villageenterprise.org","mauricen@villageenterprise.org "),
              subject = "BMs with more than 2 BSG",
              body = "Hi Team, 
            The following BMs have more than 2 BSGs",
              smtp = list(host.name = "smtp.gmail.com", port = 465, user.name = "hnyakwara@villageenterprise.org", passwd = "Neville@11", ssl = TRUE),
              attach.files=c("./BSG_bm.xlsx"),
              authenticate = TRUE,
              send = TRUE)
    
  } 
  
}

## Members with more Belonging to 2 Business Groups 


country<-c("Kenya","Rwanda","Uganda")
for (i in country) {
  k<-i
  
  BSG_hhs <- aggregate(BSG__r.BSG_ID__c ~ TW_Household__r.Household_ID__c + BSG__r.Country__c , data = BSG_sf, FUN = function(x) length(unique(x)))
  BSG_hhs<-subset(BSG_hhs, (BSG__r.BSG_ID__c>1  & BSG__r.Country__c==k)) 
  no_BSG_hhs<-nrow(BSG_hhs)
  write_xlsx(BSG_hhs,"C:\\Users\\Hurchins Manua\\Documents\\DQR_template\\BSG_hhs.xlsx" )
  
  
  if ((no_BSG_hhs>0) & k=="Kenya"){
    
    send.mail(from = "hnyakwara@villageenterprise.org",
              to = c("mhurchins@gmail.com","hnyakwara@villageenterprise.org","bodhiambo@villageenterprise.org"),
              subject = "HHs  with more than 1 BSG",
              body = "Hi Team, 
          The following HHs  belong to more than 1  BSGs",
              smtp = list(host.name = "smtp.gmail.com", port = 465, user.name = "hnyakwara@villageenterprise.org", passwd = "Neville@11", ssl = TRUE),
              attach.files=c("./BSG_hhs.xlsx"),
              authenticate = TRUE,
              send = TRUE)
    
  }
  
  if ((no_BSG_hhs>0) & k=="Uganda"){
    
    send.mail(from = "hnyakwara@villageenterprise.org",
              to = c("mhurchins@gmail.com","hnyakwara@villageenterprise.org","edwinm@villageenterprise.org"),
              subject = "HHs  with more than 1 BSG",
              body = "Hi Team, 
            The following HHs  belong to more than 1  BSGs",
              smtp = list(host.name = "smtp.gmail.com", port = 465, user.name = "hnyakwara@villageenterprise.org", passwd = "Neville@11", ssl = TRUE),
              attach.files=c("./BSG_hhs.xlsx"),
              authenticate = TRUE,
              send = TRUE)
    
    
  }
  
  if ((no_BSG_hhs>0) & k=="Rwanda"){
    
    send.mail(from = "hnyakwara@villageenterprise.org",
              to = c("mhurchins@gmail.com","hnyakwara@villageenterprise.org","mauricen@villageenterprise.org "),
              subject = "HHs  with more than 1 BSG",
              body = "Hi Team, 
            The following HHs  belong to more than 1  BSGs",
              smtp = list(host.name = "smtp.gmail.com", port = 465, user.name = "hnyakwara@villageenterprise.org", passwd = "Neville@11", ssl = TRUE),
              attach.files=c("./BSG_hhs.xlsx"),
              authenticate = TRUE,
              send = TRUE)
    
  } 
  
}



################################################################################
##################### Training DQR #############################################
################################################################################

Sys.setenv(JAVA_HOME='C:\\Program Files (x86)\\Java\\jdk-11.0.16')
library(rJava)
library("mailR")
library("writexl")


# Importing Training data from SF


train_sf <- sprintf("SELECT Cycle__c, BSG_Membership__c, BSG__c, Business_Name__c, Country__c, Mod_1_Business_Group_Formation__c, Mod_2_Basic_Savings__c, Mod_3_Business_Savings_Group_BSG_Form__c, Mod_4_Leadership_and_BSG_Executive_Comm__c, Mod_5_Introduction_to_Conservation__c, Mod_6_BSG_Constitution__c, Mod_7_BSG_Record_Keeping__c, Mod_8_BSG_Lending_and_Loans__c, Mod_9_Business_Basics__c, Mod_10_Business_Selection__c, Mod_11_Business_Planning__c, Mod_12_Managing_a_Business__c, Mod_13_Business_Record_Keeping__c, Mod_14_Value_Addition_and_Marketing__c, Mod_15_Livestock_as_a_business_or_inves__c, Group_Session__c, Training_Name__c, BSG__r.BM_Cycle__c, BSG__r.Name, BSG__r.Business_Mentor__c, Household__r.BM_Cycle__c, Household__r.BSG_Name__c, Household__r.Id,Household__r.Business_Group_name__c, Household__r.Business_Owner_Age__c, Household__r.Country__c, Household__r.Cycle__c, Household__r.Office__c, Household__r.Training_Attendance__c, Household__r.Name, Household__r.Village__c FROM Group_Training_Detail__c WHERE Household__r.Cycle__c = 'FY23C2' AND Household__r.Household_Participation_Status__c = 'Active'")
train_sf <- sf_query(train_sf)
train_sf$training<-train_sf$Training_Name__c

my_soql <- sprintf("SELECT Id, BM_Cycle_ID__c, Name, Country__c, Cycle__c, Office2__c, Cycle__r.Cycle_ID__c, Cycle__r.Cycle_Name__c, Cycle__r.Fiscal_Year_Cycle__c FROM BM_Cycle__c WHERE Cycle__r.Fiscal_Year_Cycle__c = 'a302I000005imjMQAQ'")
bmcycle <- sf_query(my_soql)

my_soql <-sprintf("SELECT BO_Name__c,Business_Group__r.Business_Group_Name__c , Business_Group__c, Household_ID__c, Household_Name__c, TW_Household__c, TW_Household__r.Id FROM Business_Group_Membership__c WHERE TW_Household__r.Cycle__c = 'FY23C2'")
hhs<- sf_query(my_soql)


train_md<-merge(train_sf,hhs,by.x="Household__r.Id",by.y="TW_Household__r.Id")

## Number of Training Sessions per BM
agg_df <- aggregate(Training_Name__c ~  Household__r.BM_Cycle__c + Country__c , data = train_sf, function(x) length(unique(x)))
agg_df<-subset(agg_df,Training_Name__c<8)
agg_df<-merge(agg_df, bmcycle, by.x = "Household__r.BM_Cycle__c", by.y = "Id", all.x = TRUE)


###Number of unique Training sessions for Every Business Group
agg_tall <- aggregate(Training_Name__c ~  TW_Household__c + Country__c + Household__r.Office__c + Business_Group__r.Business_Group_Name__c+Household_Name__c, data = train_md, function(x) length(unique(x)))
agg_tall2 <- aggregate(Training_Name__c ~  TW_Household__c + Country__c + Household__r.Office__c + Business_Group__r.Business_Group_Name__c, data = train_md, function(x) length(x))
agg_tall<-subset(agg_tall,Training_Name__c<6)



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
              subject = "BMs with Lesss than 8 training Sessions",
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
              subject = "BMs with Less than 8 TS",
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
              subject = "BMs with Less than 8 TS",
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
              subject = "HHS  with Lesss than 6 training Sessions",
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
              subject = "HHS  with Lesss than 6 training Sessions",
              body = "Hi Edwin, 
              The following HHs have less than 6 Training Sessions",
              smtp = list(host.name = "smtp.gmail.com", port = 465, user.name = "hnyakwara@villageenterprise.org", passwd = "Neville@11", ssl = TRUE),
              attach.files=c("./training_h.xlsx"),
              authenticate = TRUE,
              send = TRUE)
    
    
  }
  
  if ((training_hhs>0) & k=="Rwanda"){
    
    send.mail(from = "hnyakwara@villageenterprise.org",
              to = c("mhurchins@gmail.com","hnyakwara@villageenterprise.org","mauricen@villageenterprise.org "),
              subject = "HHS  with Lesss than 6 training Sessions",
              body = "Hi Maurice, 
             The following HHs have less than 6 Training Sessions",
              smtp = list(host.name = "smtp.gmail.com", port = 465, user.name = "hnyakwara@villageenterprise.org", passwd = "Neville@11", ssl = TRUE),
              attach.files=c("./training_h.xlsx"),
              authenticate = TRUE,
              send = TRUE)
    
  } 
  
}
















































