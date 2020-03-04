######################### Install packages ###########################

library(dplyr)
library(RODBC)
library(tibble)
library(readxl)
library(latticeExtra)
library(lattice)

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
setwd("C:/Users/vanessa.li/Documents/GitHub/Usage-Adjustments")
excelfile_Usage = '20190904UsageManagement.xlsx'
loadFile = paste(Sys.Date(),'UsageImport_VL.csv')

#####################################################################################################################
#####################################################################################################################
############################################## Import SQL Table #####################################################
############# No Need to manually update DATE - Run with Last day of prior month as publish date ####################
#####################################################################################################################


channel<-odbcConnect("Production")
Usage_Data<-sqlQuery(channel,"
                      SET NOCOUNT ON

                      Declare @StartDate date = CAST(DATEADD(MONTH, DATEDIFF(MONTH, -1, DATEADD(year,-2,GETDATE()))-1, -1) as date)                   
                      Declare @EffectiveDate date = CAST(DATEADD(MONTH, DATEDIFF(MONTH, -1, GETDATE())-1, -1) AS date)
             
                     Declare @KMtoM decimal = 0.621371
                     
                     IF OBJECT_ID('tempdb.dbo.#Data') is not NULL
                     DROP TABLE #Data
                     
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
                     ,SaleMonth
                     ,[SaleDate]
                     ,[M1PrecedingFlv]
                     ,saletype         
                     INTO #Data     
                     FROM [ras_sas].[BI].[Comparables]
                     WHERE ([source]='internet' OR (saletype in ('retail','Dealer','Trade-In') 
                     AND M1Active='Active' 
                     AND customerid in (SELECT  [CustomerID] FROM [ras_sas].[BI].[Customers] where [IsUsedForComparables]='y'))) 
                     AND Modelyear between 2008 and 2020
                     AND CategoryID in (3,	6,	14,	15,	20,	23,	25, 27,	28,	29,	30,	32,	35,	36,	164,	313,	314,	315,	316,	317,	360,	
                     362,	451,	452,	453,	2300,	2506, 2507, 2509,	2511,	2512,	2514, 2515,	2525,	2599,	2603,	2604,	2605,	2608,
                     2612,	2609,	5,	2614,	2613,	2610,	2616,	2611)
                     
                    AND MakeId NOT in (58137,78,14086,15766) --Miscellaneous,Not Attributed,Various,Mantis 
                  
                     AND MilesHours>0
                     AND SaleDate>@StartDate  AND [SaleDate]<=@EffectiveDate
                     --AND SaleDate>='2017-08-01'  AND [SaleDate]<='2019-12-31'
                     AND [M1PrecedingABCost] IS NOT NULL
                     AND (Option15 is NULL or Option15 ='0')
                     
                     SELECT CategoryId,CategoryName, SubcategoryId,SubcategoryName,([MilesHours]/Age)/1000 as Usage, MilesHoursCode,Y,SaleDate,SaleMonth,saletype 
                     FROM #Data WHERE Age>1.5 AND Y>0.1 AND Y<2
                
                     ")


########## Last month published values ################
channel<-odbcConnect("Production")
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

inputFeed<-read_excel(excelfile_Usage,sheet='In') %>%
  select(-CategoryName, -SubcategoryName, -MakeName, -CSMM, -ValidSchedule)
apply<-read_excel(excelfile_Usage,sheet='Out') %>%
  select(-CSMM,DupeSchids,ValidSchedule)

inputBorw<-read_excel(excelfile_Usage,sheet='InA') %>%
  select(-CategoryName, -SubcategoryName, -MakeName, -CSMM, -ValidSchedule,-CheckJoin)
applyBorw<-read_excel(excelfile_Usage,sheet='OutA') %>%
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

## Regression Data
JoinData <- merge(Usage_Data,inputFeed,by=c("CategoryId","SubcategoryId"))


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
                       JoinSlpBr %>% filter(MilesHoursCode==MeterCode)) %>%
  filter(ifelse(MilesHoursCode=='M',Usage >= MinMiles & Usage <= MaxMiles, Usage >= MinHours & Usage <= MaxHours))

unitMcode_SlpBr$MilesHoursCode<-as.factor(unitMcode_SlpBr$MilesHoursCode) 


# Exclude bad data
BadptExc<-unitMcode %>%
  filter(ifelse(MilesHoursCode=='M',Usage >= MinMiles & Usage <= MaxMiles, Usage >= MinHours & Usage <= MaxHours)) 


"""
### work on last month data
LastM_bySched<-merge(LastMonthUsage,inputFeed,by='ClassificationId') %>%
  select(Schedule,Slope,Intercept) %>%
  group_by(Schedule) %>%
  filter(row_number()==1)

LastM_est <-merge(BadptExc,LastM_bySched,by='Schedule') %>%
  mutate(est_y = Intercept * exp(Slope * Usage)) %>%
  mutate(index = Y/est_y)

################################### Data Regrouping #######################################

### subset to auction data only
Auction_Data<-LastM_est %>%
  filter(saletype=='Auction') %>%
  mutate(meterUse=Usage*1000) %>%
  group_by(Schedule) %>%
  filter(meterUse <= mean(meterUse) + stdInd*sd(meterUse) & meterUse>= mean(meterUse) - stdInd*sd(meterUse)) %>%
  filter(index <= mean(index) + stdInd*sd(index) & index>= mean(index) - stdInd*sd(index))
"""

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


head(Cappedoutput)
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


write.csv(SharePage,paste(Sys.Date(),'MoMSharePage_Usage.csv'))  
write.csv(UsageOutput,loadFile,row.names = F)                      









#################################### PLOTS ############################################
plotUse_adjuster <- CapAdj %>% select(Schedule,SlopeLM,InterceptLM,MedianFinal,m2Final,m1Final) %>% distinct()
plot_Use_dtpts<- rbind(data.frame(unitMcode_SlpBr %>% filter(saletype == 'Auction') %>% mutate(meterUse=Usage*1000) %>% select(Schedule,Y, meterUse)),
                       data.frame(Auction_Data %>% select(Schedule,Y, meterUse)))




seqMeter.1 = data.frame('meterRag' = seq(0,6000,100))
matrixMeter1<-merge(plotUse_adjuster %>% filter(MedianFinal<6000),seqMeter.1) %>%
  arrange(Schedule) %>% 
  mutate(sfLM = InterceptLM  * exp((SlopeLM * meterRag)/1000),
         sfcur = m1Final  * exp((m2Final  * meterRag)/1000)) 

seqMeter.2 = data.frame('meterRag' = seq(0,300000,5000))
matrixMeter2<-merge(plotUse_adjuster %>% filter(MedianFinal>6000),seqMeter.2) %>%
  arrange(Schedule) %>% 
  mutate(sfLM = InterceptLM  * exp((SlopeLM * meterRag)/1000),
         sfcur = m1Final  * exp((m2Final  * meterRag)/1000)) 

 
Regression_Curve = rbind(matrixMeter1,matrixMeter2)




plot_list<-plotUse_adjuster %>% select(Schedule)
for (i in 1:dim(plot_list)[1]){

  subsetAuc<-subset(plot_Use_dtpts, plot_Use_dtpts$Schedule==plot_list[i,1])
  
  Maxpt = max(subsetAuc$meterUse)
  ModelCurve<-subset(Regression_Curve,Regression_Curve$Schedule==plot_list[i,1])

  
  xaxis = c(0,Maxpt * 1.2)
  yaxis = c(0,2) 
  
  draw_data<-xyplot(Y ~ meterUse, subsetAuc,pch=20,cex=1.1,col='dodgerblue3',main=list(label=paste(plot_list[i,1],""),font=2,cex=2),ylim=yaxis,xlim=xaxis)

  draw_line.LM <- xyplot(sfLM ~ meterRag, ModelCurve,type=c('l'),col='dodgerblue3',lwd=3,pch=4,cex=1.5, lty=1,ylim=yaxis,xlim=xaxis)
  draw_line.Cur <- xyplot(sfcur ~ meterRag, ModelCurve,type=c('l'),col='light blue',lwd=3,pch=4,cex=1.5, lty=3,ylim=yaxis,xlim=xaxis)
  
  draw<- draw_data +as.layer(draw_line.LM) +as.layer(draw_line.Cur)
  
  mypath<-file.path('H:/Projects/52_UsageModel/202001/Plots',paste(plot_list[i,1],'.png'))
  png(file=mypath,width=1600,height=1200)
  print(draw)
  dev.off()
}
