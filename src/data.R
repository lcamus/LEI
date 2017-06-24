# dgs/dimos/sbdd/lc
# 23/06/2017

#options(java.parameters = "-Xmx3072m")
# library(XLConnect)

library(xml2)
library(openxlsx)

fill <- function(s, tag, path) {
 
  if (length(s) != RecordCount) {
    g <- grepl(tag,as.character(xml_find_all(data,path)))
    if (length(g)!=RecordCount) {
      path <- sub("/[a-z,A-Z]+:[a-z,A-Z]+$","",path)
      g <- grepl(tag,as.character(xml_find_all(data,path)))
    }
    r <- rep("",length(g))
    r[which(g)]<- s
    res <- r
  }
  else
    res <- s
  
  return(res)
  
}

fData <- "data/leifrancefullfile20170614t2230-cf2.xml"
data <- read_xml(fData)

path <- "/lei:LEIData/lei:LEIHeader/lei:RecordCount"
RecordCount <- xml_integer(xml_find_first(data,path))

path <- "/lei:LEIData/lei:LEIRecords/lei:LEIRecord"
LEI <- fill(
  xml_text(xml_find_all(data,paste0(path,"/lei:LEI"))),
  "lei:LEI",
  path
)

path <- "/lei:LEIData/lei:LEIRecords/lei:LEIRecord/lei:Entity"
LegalName <- fill(
  xml_text(xml_find_all(data,paste0(path,"/lei:LegalName"))),
  "lei:LegalName",
  path
)

path <- "/lei:LEIData/lei:LEIRecords/lei:LEIRecord/lei:Entity/lei:TransliteratedOtherEntityNames"
TransliteratedOtherEntityName <- fill(
  xml_text(xml_find_all(data,paste0(path,"/lei:TransliteratedOtherEntityName"))),
  "lei:TransliteratedOtherEntityName",
  path
)

path <- "/lei:LEIData/lei:LEIRecords/lei:LEIRecord/lei:Entity"
LegalAddress <- fill(
  xml_text(xml_find_all(data,paste0(path,"/lei:LegalAddress"))),
  "lei:LegalAddress",
  path
)

path <- "/lei:LEIData/lei:LEIRecords/lei:LEIRecord/lei:Entity"

HeadquartersAddress <- fill(
  xml_text(xml_find_all(data,paste0(path,"/lei:HeadquartersAddress"))),
  "lei:HeadquartersAddress",
  path
)

BusinessRegisterEntityID <- fill(
  xml_text(xml_find_all(data,paste0(path,"/lei:BusinessRegisterEntityID"))),
  "lei:BusinessRegisterEntityID",
  path
)
  
LegalJurisdiction <- fill(
  xml_text(xml_find_all(data,paste0(path,"/lei:LegalJurisdiction"))),
  "lei:LegalJurisdiction",
  path
)

path <- "/lei:LEIData/lei:LEIRecords/lei:LEIRecord/lei:Entity/lei:LegalForm"
  
EntityLegalFormCode <- fill(
  xml_text(xml_find_all(data,paste0(path,"/lei:EntityLegalFormCode"))),
  "lei:EntityLegalFormCode",
  path
)

OtherLegalForm <- fill(
  xml_text(xml_find_all(data,paste0(path,"/lei:OtherLegalForm"))),
  "lei:OtherLegalForm",
  path
)

path <- "/lei:LEIData/lei:LEIRecords/lei:LEIRecord/lei:Entity"

EntityStatus <- fill(
  xml_text(xml_find_all(data,paste0(path,"/lei:EntityStatus"))),
  "lei:EntityStatus",
  path
)

path <- "/lei:LEIData/lei:LEIRecords/lei:LEIRecord/lei:Registration"

InitialRegistrationDate <- fill(
  xml_text(xml_find_all(data,paste0(path,"/lei:InitialRegistrationDate"))),
  "lei:InitialRegistrationDate",
  path
)

LastUpdateDate <- fill(
  xml_text(xml_find_all(data,paste0(path,"/lei:LastUpdateDate"))),
  "lei:LastUpdateDate",
  path
)

RegistrationStatus <- fill(
  xml_text(xml_find_all(data,paste0(path,"/lei:RegistrationStatus"))),
  "lei:RegistrationStatus",
  path
)

NextRenewalDate <- fill(
  xml_text(xml_find_all(data,paste0(path,"/lei:NextRenewalDate"))),
  "lei:NextRenewalDate",
  path
)

ManagingLOU <- fill(
  xml_text(xml_find_all(data,paste0(path,"/lei:ManagingLOU"))),
  "lei:ManagingLOU",
  path
)

ValidationSources <- fill(
  xml_text(xml_find_all(data,paste0(path,"/lei:ValidationSources"))),
  "lei:ValidationSources",
  path
)

path <- "/lei:LEIData/lei:LEIRecords/lei:LEIRecord/lei:Extension"

SIREN <- fill(
  xml_text(xml_find_all(data,paste0(path,"/leifr:SIREN"))),
  "leifr:SIREN",
  path
)

LegalFormCodification <- fill(
  xml_text(xml_find_all(data,paste0(path,"/leifr:LegalFormCodification"))),
  "leifr:LegalFormCodification",
  path
)

path <- paste0(path,"/leifr:EconomicActivity")

NACEClassCode <- fill(
  xml_text(xml_find_all(data,paste0(path,"/leifr:NACEClassCode"))),
  "leifr:NACEClassCode",
  path
)

SousClasseNAF <- fill(
  xml_text(xml_find_all(data,paste0(path,"/leifr:SousClasseNAF"))),
  "leifr:SousClasseNAF",
  path
)

path <- "/lei:LEIData/lei:LEIRecords/lei:LEIRecord/lei:Extension"

FundNumber <- fill(
  xml_text(xml_find_all(data,paste0(path,"/leifr:FundNumber"))),
  "leifr:FundNumber",
  path
)

FundManagerBusinessRegisterID <- fill(
  xml_text(xml_find_all(data,paste0(path,"/leifr:FundManagerBusinessRegisterID"))),
  "leifr:FundManagerBusinessRegisterID",
  path
)

path <- "/lei:LEIData/lei:LEIRecords/lei:LEIRecord/lei:Entity/lei:AssociatedEntity"

AssociatedLEI <- fill(
  xml_text(xml_find_all(data,paste0(path,"/lei:AssociatedLEI"))),
  "lei:AssociatedLEI",
  path
)

AssociatedEntityName <- fill(
  xml_text(xml_find_all(data,paste0(path,"/lei:AssociatedEntityName"))),
  "lei:AssociatedEntityName",
  path
)

path <- "/lei:LEIData/lei:LEIRecords/lei:LEIRecord/lei:Entity"

EntityExpirationDate <- fill(
  xml_text(xml_find_all(data,paste0(path,"/lei:EntityExpirationDate"))),
  "lei:EntityExpirationDate",
  path
)

EntityExpirationReason <- fill(
  xml_text(xml_find_all(data,paste0(path,"/lei:EntityExpirationReason"))),
  "lei:EntityExpirationReason",
  path
)

path <- "/lei:LEIData/lei:LEIRecords/lei:LEIRecord/lei:Entity/lei:SuccessorEntity"
SuccessorLEI <- fill(
  xml_text(xml_find_all(data,paste0(path,"/lei:SuccessorLEI"))),
  "lei:SuccessorLEI",
  path
)

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

fRef <- "data/referentiel-adm._centrale-odac-apul-asso-resident-20170328.xlsx"
fRef.lei <- paste0(regmatches(fRef,regexpr("^.+(?=\\.xlsx$)",fRef,perl=T)),"+lei.xlsx")

wb <- openxlsx::loadWorkbook(fRef)

l <- list()

for (s in seq_along(getSheetNames(fRef))) {
  df.apu <- read.xlsx(fRef, sheet = s)
  hasLEI <- df.apu$SIREN %in% dfXml$SIREN
  # df <- setNames(
  #   data.frame(hasLEI,rep("",length(hasLEI)),stringsAsFactors=F),
  #   c("hasLEI","LEI")
  # )
  df <- setNames(
    data.frame(hasLEI,
               matrix(data="",nrow=nrow(df.apu),ncol=ncol(dfXml)),
               stringsAsFactors=F),
    c("hasLEI",names(dfXml))
  )
  df[hasLEI,-1] <- dfXml[dfXml$SIREN %in% df.apu$SIREN,]
  # df[df$hasLEI,]$LEI <- dfXml[dfXml$SIREN %in% df.apu$SIREN,]$LEI
  l[[length(l)+1]] <- df
  # l[[length(l)+1]] <- dfXml[dfXml$SIREN %in% df.apu$SIREN,]
  writeData(wb, s, l[[s]], startCol = 3, startRow = 1)
}

saveWorkbook(wb, fRef.lei, overwrite = T)


