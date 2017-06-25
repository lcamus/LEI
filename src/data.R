# dgs/dimos/sbdd/lc
# 25/06/2017

library(xml2)
library(openxlsx)

Sys.setenv("R_ZIPCMD" = "C:/Rtools/bin/zip.exe")

fRef <- "data/referentiel-adm._centrale-odac-apul-asso-resident-20170328.xlsx"
fData <- "data/leifrancefullfile20170624t2230-cf2.xml"
#src: https://lei-france.insee.fr/telechargement

fill <- function(p) {
  
  rm <- regmatches(
    p,
    regexec("^(?<parent>.+)\\/(?<node>[a-z,A-Z]+:[a-z,A-Z]+)$",p,perl=T)
  )[[1]]
  parent <- rm[2]
  node <- rm[3]
 
  s <- xml_text(xml_find_all(data,p))
  
  if (length(s) != RecordCount) {
    g <- grepl(node,as.character(xml_find_all(data,parent)))
    if (length(g)!=RecordCount) {
      path <- sub("/[a-z,A-Z]+:[a-z,A-Z]+$","",path)
      g <- grepl(node,as.character(xml_find_all(data,path)))
    }
    r <- rep("",length(g))
    r[which(g)]<- s
    res <- r
  }
  else
    res <- s
  
  return(res)
  
}

data <- read_xml(fData)

path <- "/lei:LEIData/lei:LEIHeader/lei:RecordCount"
RecordCount <- xml_integer(xml_find_first(data,path))

path <- "/lei:LEIData/lei:LEIRecords/lei:LEIRecord"
LEI <- fill(paste0(path,"/lei:LEI"))

path <- "/lei:LEIData/lei:LEIRecords/lei:LEIRecord/lei:Entity"
LegalName <- fill(paste0(path,"/lei:LegalName"))

path <- "/lei:LEIData/lei:LEIRecords/lei:LEIRecord/lei:Entity/lei:TransliteratedOtherEntityNames"
TransliteratedOtherEntityName <- fill(paste0(path,"/lei:TransliteratedOtherEntityName"))

path <- "/lei:LEIData/lei:LEIRecords/lei:LEIRecord/lei:Entity"
LegalAddress <- fill(paste0(path,"/lei:LegalAddress"))

path <- "/lei:LEIData/lei:LEIRecords/lei:LEIRecord/lei:Entity"
HeadquartersAddress <- fill(paste0(path,"/lei:HeadquartersAddress"))
BusinessRegisterEntityID <- fill(paste0(path,"/lei:BusinessRegisterEntityID"))
LegalJurisdiction <- fill(paste0(path,"/lei:LegalJurisdiction"))

path <- "/lei:LEIData/lei:LEIRecords/lei:LEIRecord/lei:Entity/lei:LegalForm"
EntityLegalFormCode <- fill(paste0(path,"/lei:EntityLegalFormCode"))
OtherLegalForm <- fill(paste0(path,"/lei:OtherLegalForm"))

path <- "/lei:LEIData/lei:LEIRecords/lei:LEIRecord/lei:Entity"
EntityStatus <- fill(paste0(path,"/lei:EntityStatus"))

path <- "/lei:LEIData/lei:LEIRecords/lei:LEIRecord/lei:Registration"
InitialRegistrationDate <- fill(paste0(path,"/lei:InitialRegistrationDate"))
LastUpdateDate <- fill(paste0(path,"/lei:LastUpdateDate"))
RegistrationStatus <- fill(paste0(path,"/lei:RegistrationStatus"))
NextRenewalDate <- fill(paste0(path,"/lei:NextRenewalDate"))
ManagingLOU <- fill(paste0(path,"/lei:ManagingLOU"))
ValidationSources <- fill(paste0(path,"/lei:ValidationSources"))

path <- "/lei:LEIData/lei:LEIRecords/lei:LEIRecord/lei:Extension"
SIREN <- fill(paste0(path,"/leifr:SIREN"))
LegalFormCodification <- fill(paste0(path,"/leifr:LegalFormCodification"))

path <- paste0(path,"/leifr:EconomicActivity")
NACEClassCode <- fill(paste0(path,"/leifr:NACEClassCode"))
SousClasseNAF <- fill(paste0(path,"/leifr:SousClasseNAF"))

path <- "/lei:LEIData/lei:LEIRecords/lei:LEIRecord/lei:Extension"
FundNumber <- fill(paste0(path,"/leifr:FundNumber"))
FundManagerBusinessRegisterID <- fill(paste0(path,"/leifr:FundManagerBusinessRegisterID"))

path <- "/lei:LEIData/lei:LEIRecords/lei:LEIRecord/lei:Entity/lei:AssociatedEntity"
AssociatedLEI <- fill(paste0(path,"/lei:AssociatedLEI"))
AssociatedEntityName <- fill(paste0(path,"/lei:AssociatedEntityName"))

path <- "/lei:LEIData/lei:LEIRecords/lei:LEIRecord/lei:Entity"
EntityExpirationDate <- fill(paste0(path,"/lei:EntityExpirationDate"))
EntityExpirationReason <- fill(paste0(path,"/lei:EntityExpirationReason"))

path <- "/lei:LEIData/lei:LEIRecords/lei:LEIRecord/lei:Entity/lei:SuccessorEntity"
SuccessorLEI <- fill(paste0(path,"/lei:SuccessorLEI"))

# write XML origin content to Excel file:

dfXml <- data.frame(
  cbind(LEI,LegalName,
        TransliteratedOtherEntityName,
        LegalAddress,HeadquartersAddress,BusinessRegisterEntityID,
        LegalJurisdiction,
        EntityLegalFormCode,OtherLegalForm,
        EntityStatus,InitialRegistrationDate,LastUpdateDate,
        RegistrationStatus,NextRenewalDate,ManagingLOU,ValidationSources,SIREN,
        LegalFormCodification,NACEClassCode,SousClasseNAF,
        FundNumber,FundManagerBusinessRegisterID,
        AssociatedLEI,AssociatedEntityName,
        EntityExpirationDate,EntityExpirationReason,
        SuccessorLEI),
  stringsAsFactors=F
)

openxlsx::write.xlsx(dfXml, paste0(fData,".xlsx"))

# Get APU:

fRef.lei <- paste0(regmatches(fRef,regexpr("^.+(?=\\.xlsx$)",fRef,perl=T)),"+lei.xlsx")

wb <- openxlsx::loadWorkbook(fRef)

l <- list()

for (s in seq_along(getSheetNames(fRef))) {
  df.apu <- read.xlsx(fRef, sheet = s)
  # hasLEI <- df.apu$SIREN %in% dfXml$SIREN
  df <- merge(x=df.apu,y=dfXml,by="SIREN",all.x=T,all.y=F)
  df$hasLEI <- ifelse(is.na(df$LEI),F,T)
  df <- df[,c(1:2,ncol(df),4:ncol(df)-1)]
  df <- setNames(df,sub("hasLEI","?LEI",names(df)))
  l[[length(l)+1]] <- df[,c(-1,-2)]
  writeData(wb, s, l[[s]], startCol = 3, startRow = 1)
}

saveWorkbook(wb, fRef.lei, overwrite = T)


