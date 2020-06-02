######################### Install packages ###########################

library(dplyr)
library(RODBC)
library(tibble)
library(xlsx)
library(latticeExtra)
library(lattice)
library(lubridate)

stdInd = 2
## sense use rate of boundaries
ChangeRate=0.1

## MoM limit change
#month_limit = 9999
MoM_limit = .075

## Top and bottom lines for usage
MinHours = .5*52/1000 # half an hour per week
MaxHours = 16*7*52/1000 # 52 weeks/yr, 7 days a week, 16 hours a day

MinMiles = 40 * 52/1000 # 40 miles  per week
MaxMiles = 300000/1000 # 300k per yr, from Chris estimate


######################### Read the input file ##################################
filepath = 'C:/Users/vanessa.li/Documents/GitHub/Usage-Adjustments'
setwd(filepath)
#plotFolder = paste("Plots",Sys.Date())
#dir.create(plotFolder)
excelfile_Usage = '20200311UsageManagement.xlsx'
loadFile = paste(Sys.Date(),'UsageImport_VL.csv')

### publish date is the last day of prior month
publishDate <- Sys.Date() - days(day(Sys.Date()))
channel<-odbcConnect("production")
#####################################################################################################################
#####################################################################################################################
############################################## Import SQL Table #####################################################
############# No Need to manually update DATE - Run with Last day of prior month as publish date ####################
#####################################################################################################################

Usage_Data<-sqlQuery(channel,"
                     SET NOCOUNT ON

                      Declare @StartDate date = CAST(DATEADD(MONTH, DATEDIFF(MONTH, -1, DATEADD(year,-2,GETDATE()))-2, -1) as date)                   
                      Declare @EffectiveDate date = CAST(DATEADD(MONTH, DATEDIFF(MONTH, -1, GETDATE())-1, -1) AS date)
             
                     Declare @KMtoM decimal = 0.621371
                     
                     
                     DROP TABLE IF exists #Data
                     
                     SELECT 
                     [InternetComparableId] 
                     ,CategoryID   
                     ,[CategoryName] 
                     ,SubcategoryID 
                     ,SubcategoryName
                     , Saleprice
                     ,[ModelYear]
                     ,CASE WHEN saletype='Auction' THEN SalePrice/[M1PrecedingFlv] ELSE SalePrice/[M1PrecedingFmv] END as Y
                     ,cast(YEAR(SaleDate)-ModelYear + (MONTH(SaleDate)-6)/12.00 as decimal(10,4))  as Age
                     ,CASE WHEN MilesHoursCode='K' THEN MilesHours*@KMtoM ELSE MilesHours END AS MilesHours
                     ,CASE WHEN MilesHoursCode='K' THEN 'M' ELSE MilesHoursCode END AS [MilesHoursCode]
					 ,EOMONTH([SaleDate]) as EffectiveDate
                     ,SaleMonth
                     ,[SaleDate]
                     ,[M1PrecedingFlv]
                     ,saletype         
                     INTO #Data     
                     FROM [ras_sas].[BI].[Comparables]
                     WHERE ([source]='internet' OR (saletype in ('retail','Dealer','Trade-In') 
                     AND M1Active='Active' 
                     AND customerid in (SELECT  [CustomerID] FROM [ras_sas].[BI].[Customers] where [IsUsedForComparablesUSNA]='y'))) 
                     AND Modelyear between 2008 and 2020
                     AND CategoryID in (3,	6,	14,	15,	20,	23,	25, 27,	28,	29,	30,	32,	35,	36,	164,	313,	314,	315,	316,	317,	360,	
                     362,	451,	452,	453,	2300,	2506, 2507, 2509,	2511,	2512,	2514, 2515,	2525,	2599,	2603,	2604,	2605,	2608,
                     2612,	2609,	5,	2614,	2613,	2610,	2616,	2611)
                     
                    AND MakeId NOT in (58137,78,14086,15766) --Miscellaneous,Not Attributed,Various,Mantis 
                  
                     AND MilesHours>0
                     AND SaleDate>@StartDate  AND [SaleDate]<=@EffectiveDate
                     --AND SaleDate>='2017-08-01'  AND [SaleDate]<='2019-12-31'
                     AND [M1PrecedingABCostUSNA] IS NOT NULL
                     AND (Option15 is NULL or Option15 ='0')
                     
                                SELECT CategoryId,CategoryName, SubcategoryId,SubcategoryName,([MilesHours]/Age)/1000 as Usage, MilesHoursCode,Y,SaleDate,SaleMonth,EffectiveDate,saletype 
                     FROM #Data WHERE Age>1.5 AND Y>0.1 AND Y<2
                
                     ")


########## Last month published values ################
LastMonthUsage<-sqlQuery(channel,"SELECT 
      [ClassificationId]
      ,[Slope]
      ,[Intercept]
   
  FROM [ras_sas].[BI].[AppraisalBookUsageAdjCoefficientsMKT]
  Where [AppraisalBookIssueId] = (Select max([AppraisalBookIssueId]) FROM [ras_sas].[BI].[AppraisalBookUsageAdjCoefficientsMKT] 
                         where MarketCode ='USNA')")


### MoM check function
MoMlimit_intc <- function(last_month,current_month,limit){
  upline = last_month * (1+ limit)
  btline = last_month * (1- limit)
  result = ifelse(is.na(last_month),current_month,pmin(upline,pmax(btline,current_month)))
  return(result)
}

MoMlimit_slp <- function(last_month,current_month,limit){
  upline = last_month * (1+ limit)
  btline = last_month * (1- limit)
  result = ifelse(is.na(last_month),current_month,pmax(upline,pmin(btline,current_month)))
  return(result)
}

############### Read input file by tabs ###################
inputFeed<-data.frame(read.xlsx(excelfile_Usage,sheetName='In')) %>%
  select(-CategoryName, -SubcategoryName, -MakeName, -CSMM, -ValidSchedule)
apply<-data.frame(read.xlsx(excelfile_Usage,sheetName='Out')) %>%
  select(-CSMM,DupeSchids,ValidSchedule)

inputBorw<-data.frame(read.xlsx(excelfile_Usage,sheetName='InA')) %>%
  select(-CategoryName, -SubcategoryName, -MakeName, -CSMM, -ValidSchedule,-CheckJoin)
applyBorw<-data.frame(read.xlsx(excelfile_Usage,sheetName='OutA')) %>%
  select(-CSMM,DupeSchids,ValidSchedule)

######### combine two output tabs ##########
applyall<-rbind(apply,applyBorw)


## boundaries 
BoundBorw<-merge(inputFeed %>% select(Schedule,MinFactor,MaxFactor,Overide) %>% rename(SlopeBorrow=Schedule) %>% distinct(), 
                 inputBorw %>% select(Schedule,SlopeBorrow,MeterCode) %>% distinct()
                 ,by='SlopeBorrow') %>% select (Schedule,MinFactor,MaxFactor,Overide,MeterCode) 

bounds <- rbind(data.frame(inputFeed) %>%
                  select(Schedule,MinFactor,MaxFactor,Overide,MeterCode) %>%
                  distinct(),BoundBorw)


################################### Data Cleaning #######################################
month_lost <- Usage_Data %>% select(EffectiveDate) %>% distinct() %>% arrange(EffectiveDate) %>% slice(1)
## Regression Data

JoinData <- merge(Usage_Data,inputFeed,by=c("CategoryId","SubcategoryId"))

## manage the meter code
unitMcode<-rbind(JoinData %>%
                   filter(is.na(MilesHoursCode)==T) %>%
                   mutate(MilesHoursCode = MeterCode),
                 JoinData %>% filter(MilesHoursCode==MeterCode))

unitMcode$MilesHoursCode<-as.factor(unitMcode$MilesHoursCode)

## Slope Borrow
JoinSlpBr <- merge(Usage_Data,inputBorw,by=c("CategoryId","SubcategoryId"))

unitMcode_SlpBr<-rbind(JoinSlpBr %>%
                         filter(is.na(MilesHoursCode)==T) %>%
                         mutate(MilesHoursCode = MeterCode),
                       JoinSlpBr %>% filter(as.character(MilesHoursCode)== as.character(MeterCode))) %>%
  filter(ifelse(MilesHoursCode=='M',Usage >= MinMiles & Usage <= MaxMiles, Usage >= MinHours & Usage <= MaxHours))

unitMcode_SlpBr$MilesHoursCode<-as.factor(unitMcode_SlpBr$MilesHoursCode) 


# Exclude bad data
BadptExc<-unitMcode %>%
  filter(ifelse(MilesHoursCode=='M',Usage >= MinMiles & Usage <= MaxMiles, Usage >= MinHours & Usage <= MaxHours)) %>%
  ## exclude the one-month data lost from last month to this month 
  filter(as.Date(EffectiveDate) != as.Date(month_lost$EffectiveDate))


Auction_Data<-BadptExc %>%
  filter(saletype=='Auction') %>%
  mutate(meterUse=Usage*1000) %>%
  group_by(Schedule) %>%
  filter(meterUse <= mean(meterUse) + stdInd*sd(meterUse) & meterUse>= mean(meterUse) - stdInd*sd(meterUse))


############################### Schedule List ############################### 
summary<-inputFeed %>%
  group_by(Schedule) %>%
  summarise(n=n()) 
ScheduleList<-data.frame(summary[,1])
# Number of groups
nSched<-dim(ScheduleList)[1]


###################################### Regression ######################################################

# define variables
m1<-rep(NA,nSched)
m2<-rep(NA,nSched)
nAuc_wOL<-rep(NA,nSched)
nAuc<-rep(NA,nSched)
AvgUsage<-rep(NA,nSched)
MedUsage<-rep(NA,nSched)

Usage_list<-matrix(NA,nSched,6)


### loop through the schedules
for (i in 1:nSched){

  ModelData<-subset(Auction_Data, Auction_Data$Schedule==ScheduleList[i,1])
  
  if(nrow(ModelData)>0){
    fitcat<-lm(log(Y)~Usage,data=ModelData)
    
    cooksd <- cooks.distance(fitcat)
    influential <- as.numeric(names(cooksd)[(cooksd > 20*mean(cooksd, na.rm=T) | cooksd >1)])
    
    remainData =
      if(length(influential) ==0){
        ModelData
      }
      else{ModelData[-influential,]}
    
    fit_good<-lm(log(Y)~Usage,data=remainData)
    m1[i]<-exp(coef(fit_good)[1])
    m2[i]<-coef(fit_good)[2]
    nAuc_wOL[i]<-dim(ModelData)[1]
    nAuc[i]<-dim(remainData)[1]
    AvgUsage[i]<-mean(remainData$Usage)*1000
    MedUsage[i]<-median(remainData$Usage)*1000
    
  }
  Usage_list<-data.frame(m1,m2,MedUsage,AvgUsage,nAuc_wOL,nAuc)
}


rownames(Usage_list)<-ScheduleList[,1]
Usage_model<-rownames_to_column(Usage_list) %>%
  rename(Schedule=rowname)


############################ Stat summary in Total Sales ###############################
# regular schedules

TS_median<- BadptExc %>%
  mutate(meterUse=Usage*1000) %>%
  group_by(Schedule) %>%
  #filter(meterUse <= median(meterUse) + stdInd*sd(meterUse) & meterUse>= median(meterUse) - stdInd*sd(meterUse)) %>%
  summarise(TSmed = median(meterUse),
            TSmean = mean(meterUse),
            nTotal=n())

RegOutput<-merge(Usage_model,TS_median,by=c("Schedule")) 

############################ Borrow Schedules ###############################
TS_medSlpBr<- unitMcode_SlpBr %>%
  mutate(meterUse=Usage*1000) %>%
  group_by(Schedule,SlopeBorrow) %>%
  filter(meterUse <= median(meterUse) + stdInd*sd(meterUse) & meterUse>= median(meterUse) - stdInd*sd(meterUse)) %>%
  summarise(TSmed = median(meterUse),
            TSmean = mean(meterUse),
            nTotal=n()) %>%
  rename(Schedule=SlopeBorrow, NewSched=Schedule)

## Join to regular
SlpBrOutput<-merge(Usage_model,TS_medSlpBr,by=c("Schedule")) %>%
  select(-Schedule) %>%
  select(NewSched,everything()) %>%
  rename(Schedule=NewSched)

AdjustOutput <- rbind(RegOutput,SlpBrOutput)



##########################  Apply schedules to all output cat/subcat ###############################
Application<-merge(AdjustOutput,applyall,by=c("Schedule"),all.y=T) %>%
  select(-Level1,-Level2,-MonthsOfData,-DupeSchids,-ValidSchedule)



######################### Get prior month values - prepare table to compare and set limit ####################

Cappedoutput<-merge(Application,LastMonthUsage,by=c('ClassificationId'),all.x=T) %>%
  select(ClassificationId,Schedule,CategoryId,SubcategoryId,CategoryName,SubcategoryName,TSmed,m1,m2,Slope,Intercept,nTotal,nAuc) %>%
  rename(SlopeLM=Slope,InterceptLM=Intercept,MedianFinal=TSmed)


###############################################################################################################################################
###################################################### Explore the share page and upload file ##################################################
###############################################################################################################################################


#### prepare the page for share
CapAdj <- merge(Cappedoutput,bounds,by=('Schedule'),all.x=T) %>%
  mutate(SteepBound = log(MaxFactor)/(ChangeRate*MedianFinal/1000),
         FlatBound = log(MinFactor)/(ChangeRate*MedianFinal/1000)) %>%
  mutate(capm2=ifelse(Overide == "Max",SteepBound,
                      ifelse(Overide =="Min",FlatBound,
                             ifelse(m2>0 | m2>FlatBound,FlatBound,
                                    ifelse(m2<SteepBound,SteepBound,m2))))) %>%
  mutate(m2Final = MoMlimit_slp(SlopeLM,capm2,MoM_limit)) %>%
  mutate(m1calc=1/(exp(m2Final*MedianFinal/1000)),
         m1Final = MoMlimit_intc(InterceptLM,m1calc,MoM_limit)) %>%
  mutate(Hrspct1 = (log(0.99)*1000)/m2Final,
         Hrspct1_LM = (log(0.99)*1000)/SlopeLM) %>%
  mutate(diff_Slope = m2Final/SlopeLM-1,
         diff_Intercept =m1Final/InterceptLM-1,
         diff_Hrs1pct =Hrspct1/Hrspct1_LM-1)%>%
  #filter(is.na(diff_Slope))
  select(ClassificationId,Schedule,CategoryId,SubcategoryId,CategoryName,SubcategoryName,MeterCode,
         SlopeLM,InterceptLM,Hrspct1_LM,
         nTotal,nAuc,
         MedianFinal,m2Final,m1Final,Hrspct1,
         diff_Slope,diff_Intercept,diff_Hrs1pct
         ) %>%
  arrange(MeterCode,Schedule)


SharePage<-CapAdj %>%
  select(-ClassificationId,-CategoryId,-SubcategoryId,-CategoryName,-SubcategoryName) %>%
  distinct()




################### Upload file ##########################

UsageOutput<-CapAdj %>%
  select(CategoryId,CategoryName,SubcategoryId,SubcategoryName,m2Final,m1Final,MeterCode) %>%
  rename(Category=CategoryName,Subcategory=SubcategoryName,Slope=m2Final,Intercept=m1Final)



write.xlsx2(as.data.frame(SharePage),file = paste(Sys.Date(),'MoMSharePage_Usage.xlsx'), sheetName = 'Sheet1',row.names = F)
write.csv(UsageOutput,loadFile,row.names = F)                      



