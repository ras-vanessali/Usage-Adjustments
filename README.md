# Usage-Adjustments
This program is to calculate the usage adjusters for C or CS level schedules. For each schedule, this program will return the slope and intercept. 

## Prerequisites
a)  You need to have R studio installed in your computer. 
b)  Download the most recent `UsageManagement.xlsx` from git or valuation share folder.
c)  Install the following R libraries:
```
RODBC
readxl
tibble
dplyr
```
## Data scope
### what bi.views in ras_sas been used?
```
BI.Comparables
BI.Customers
BI.AppraisalBookUsageAdjCoefficients
```
### what sales data been used?
- US data only
- use auction data in regression model and use all sale type data calculating average usage
- only customers used for comparables
- categories in the input file
- sales in rolling 24 months
- model year in between 2008 to 2020
- make is not Miscellaneous, Not Attributed and Mantis
- M1precedingABCost IS NOT NULL
- M1Active = 'Active'
- option15 IS NULL or Option15=0
- age is greater than 1.5
- saleprice/M1value between .1 and 2

## Data cleaning
1) Meter code has to be making sense to the category. For example, truck tractors' meter code can only be miles, any kilometer will be transferred to miles; any missing meter code will be considered as miles; any other meter code will be removed from the data set.  

2) Bad data:
- meter code = 'M'
`MIN = 40 * 52/1000`  # 40 miles  per week
`MAX = 300000/1000`  # 300k per yr, from Chris C estimate

- meter code = 'H'
`MIN = .5*52/1000`  # half an hour per week
`MAX = 16*7*52/1000`  # 52 weeks/yr, 7 days a week, 16 hours a day
3) Outliers:
```
 > mean(meter usage) + 2*stddev(meter usage)
 < mean(meter usage) - 2*stddev(meter usage)
 ```
 4) Cook distance for regression:
 find and remove the influential points by each schedule using the following rule. 
 `coodsd > 20* mean(cooksd) or cooksd >1`
## Regression model
1) Calculate usage per year: `(MilesHours/Age)/1000`
2) Fit the following regression for each schedule through a loop:
```
log(SP/M1value) = Usage
```
output the slope and intercept, and the mean and median of the usage for  each schedule

3) Get the median of usage for all schedules

4) Applied slopes to the borrow schedules

5) Cap the slopes to a reasonable depreciation rate, 
```
steep bound = log(MaxFactor)/(10% * MedianFinal)
flat bound = log(MinFactor)/(10% * MedianFinal)
```
If slope is out of range, pull it back to the steep or flat boundaries. Or follow `Overide` to cap the slope.

6) Recalculate intercepts by the capped slope and median usage: 
`1/(exp(Slope * MedianUsage/1000)`
7) Limit the slope and intercept from last month by 7.5% 
8) Calculate `Hrspct1` for current month and last month:
`(log(0.99) * 1000)/ slope`
