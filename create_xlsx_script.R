################################
# Environment:
################################
# R version 3.6.1 (2019-07-05) -- "Action of the Toes"
# Platform: x86_64-apple-darwin15.6.0 (64-bit)
# Running under: macOS Catalina 10.15.7
library(dplyr) # dplyr_1.0.2 
library(tidyr) # tidyr_1.1.2
library(openxlsx) # openxlsx_4.2.3
library(ggplot2) # ggplot2_3.3.2
library(psych) # psych_2.0.9
sessionInfo()

################################
# Load Data:
################################
setwd("/Users/shlong/Desktop/IXIS Data Analyst") # Change directory
getwd()

addsToCart_csv = read.csv('DataAnalyst_Ecom_data_addsToCart.csv')
head(addsToCart_csv)
sessionCounts_csv = read.csv('DataAnalyst_Ecom_data_sessionCounts.csv')
head(sessionCounts_csv)

################################
# Data Error Checks:
################################
## Missing values ?
addsToCart_csv[!complete.cases(addsToCart_csv),] # No missing values
sessionCounts_csv[!complete.cases(sessionCounts_csv),] # No missing values

## Dates in 'sessionCounts_csv' consecutive ?
sessionDays = as.Date(unique(sessionCounts_csv$dim_date),"%m/%d/%y")
max(sessionDays) - min(sessionDays) # Timeframe from July 1st 2012 to June 30th 2013 (Time difference of 364 days)

date_range = seq(min(sessionDays), max(sessionDays), by = 1) 
date_range[!date_range %in% sessionDays] # No missing days in Timeframe

## Text columns in 'sessionCounts_csv'
str(sessionCounts_csv)
levels(sessionCounts_csv$dim_browser) # 57 different browsers (Chrome, Safari, Firefox, ...)
levels(sessionCounts_csv$dim_deviceCategory) # 3 different devices (desktop, mobile, tablet)

################################
# Data Cleaning:
################################
### Numerical variables in 'sessionCounts_csv'
## For a given browser * device * date:
  # sessions = number of site visits
  # transactions = number of transactions
  # QTY = "QTY is just "quantity" as in quantity of items in the transaction." -Jon

sessionCounts_csv1 = sessionCounts_csv %>% mutate(index = 1:nrow(sessionCounts_csv))

## A) Issue: sessions < transactions
# PROBLEM SEVERITY: low
moreTthanS_df = sessionCounts_csv1 %>% 
                filter(sessions - transactions < 0) %>% 
                arrange(desc(transactions)) 
head(moreTthanS_df)

## B) Issue: zero transactions but QTY > 0 
# PROBLEM SEVERITY: low
zeroTmoreQ_df = sessionCounts_csv1 %>% 
                filter((transactions == 0 & QTY > 0)) %>% 
                arrange(desc(sessions))  
head(zeroTmoreQ_df)

## C) Issue: transactions < QTY 
# PROBLEM SEVERITY: medium
moreTthanQ_df = sessionCounts_csv1 %>% 
  filter(QTY - transactions < 0) %>% 
  arrange(desc(transactions))  
head(moreTthanQ_df)

### Investigate Data Issues:
remove_indexs = c(moreTthanS_df$index,zeroTmoreQ_df$index, moreTthanQ_df$index)
length(remove_indexs)
length(unique(remove_indexs))
sessionCounts_csv1[remove_indexs[duplicated(remove_indexs)],] # instances w/ multiple issues

## Issues are constant over Timeframe
sessionCounts_csv1 %>% 
  mutate(date = as.Date(dim_date,"%m/%d/%y")) %>% 
  mutate(Issue = index %in% unique(remove_indexs)) %>% 
  ggplot()+
  geom_histogram(aes(date, color = Issue, fill = Issue)) +
  scale_colour_manual(values = c('#57606C','red')) +
  scale_fill_manual(values = c('#7B848F','pink'))+
  labs(x="Date",y="Browser * Device count")

## Issues vary across Browser * Device (Internet Explorer has the majority of issues)
BrowserDevices_df = sessionCounts_csv1 %>% 
  group_by(dim_browser,dim_deviceCategory) %>% 
  summarise(sessions = sum(sessions),transactions = sum(transactions),QTY = sum(QTY)) %>% 
  mutate(key = paste(dim_browser,dim_deviceCategory))

Issues_df = sessionCounts_csv1 %>% 
  mutate(Issue = index %in% unique(remove_indexs)) %>% 
  filter(Issue == TRUE) %>% 
  mutate(key = paste(dim_browser,dim_deviceCategory)) %>% 
  select(-c("dim_browser","dim_deviceCategory")) %>% 
  group_by(key) %>% 
  summarise(sessions_I = sum(sessions),transactions_I = sum(transactions),QTY_I = sum(QTY)) 

BrowserDevices_df %>% 
  left_join(Issues_df,by="key") %>% 
  mutate( S_per = (replace_na(sessions_I, 0) / sessions)*100, 
          T_per = (replace_na(transactions_I, 0) / transactions)*100,
          Q_per = (replace_na(QTY_I, 0) / QTY)*100) %>% 
  arrange(desc(sessions)) %>%
  head(15) %>% 
  ggplot(aes(x=reorder(key,sessions/1000))) +
  geom_bar(aes(y=sessions/1000), stat="identity", alpha=.3, fill='#7B848F', color='#57606C') +
  geom_bar(aes(y=sessions_I/1000), stat="identity", alpha=.8, fill='pink', color='red') +
  geom_text(aes(label=paste(round(S_per, 1), "%", sep=""), y =sessions/1000, hjust = -1,size=13*S_per),color='red') +
  labs(y="Sessions (K)",x="",title="Yearly Sessions across Browser * Device") +
  theme(axis.text.x = element_text(size = 20),
        axis.text.y = element_text(size = 20),  
        axis.title.x = element_text(size = 20),
        axis.title.y = element_text(size = 20),
        title = element_text(size = 20),
        legend.position = "none")+
  coord_flip()

### SUMMARY OF DATA AFTER CLEANING:
# sessions -> 95.66% remain
1 - sum(Issues_df$sessions_I) / sum(sessionCounts_csv1$sessions)
# transactions -> 95.66% remain
1 - sum(Issues_df$transactions_I) / sum(sessionCounts_csv1$transactions)
# QTY -> 99.25% remain
1 - sum(Issues_df$QTY_I) / sum(sessionCounts_csv1$QTY)

################################
# Create Excel Tables (.xlsx):
################################
sessionCounts_csv2 = sessionCounts_csv1 %>%
                      filter(!index %in%  unique(remove_indexs)) %>% ## DATA CLEAN 
                      select(-c("index")) %>% 
                      mutate(date = as.Date(dim_date,"%m/%d/%y"),
                             year = as.numeric(format(date, "%Y")),
                             month = as.numeric(format(date, "%m")))
## Quick Check:
nrow(sessionCounts_csv1) - nrow(sessionCounts_csv2) == length(unique(remove_indexs))
                      
## Sheet 1 - Option 1
Sheet1_Option1 = sessionCounts_csv2 %>% 
          select(year, month, sessions, transactions, QTY) %>% 
          group_by(year,month) %>% 
          summarise(sessions = sum(sessions),transactions = sum(transactions),QTY = sum(QTY)) %>% 
          mutate(ECR = transactions / sessions) %>% 
          mutate(QpT = QTY/transactions) 

## Sheet 1 - Option 2
Sheet1_Option2 = sessionCounts_csv2 %>% 
  select(year, month, dim_deviceCategory, sessions, transactions, QTY) %>% 
  group_by(year, month, dim_deviceCategory) %>% 
  summarise(sessions = sum(sessions),transactions = sum(transactions),QTY = sum(QTY)) %>% 
  mutate(ECR = transactions / sessions) %>% 
  mutate(QpT = QTY/transactions) 
  
## Sheet 2
Sheet2 = Sheet1_Option1 %>% 
  left_join(addsToCart_csv,by = c("month" = "dim_month")) %>% # Since 'month' is a primary key
  select(-c("dim_year")) %>% 
  mutate(ATCpS = addsToCart / sessions) %>% 
  ungroup() %>%
  mutate(absSessions = sessions - lag(sessions, k=1), # https://stats.mom.gov.sg/SL/Pages/Absolute-vs-Relative-Change-Concepts-and-Definitions.aspx
         relSessions = absSessions / lag(sessions, k=1),
         absAddsToCart = addsToCart - lag(addsToCart, k=1),
         relAddsToCart = absAddsToCart / lag(addsToCart, k=1),
         absTransactions = transactions - lag(transactions, k=1),
         relTransactions = absTransactions / lag(transactions, k=1),
         absQTY = QTY - lag(QTY, k=1),
         relQTY = absQTY / lag(QTY, k=1),
         absECR = ECR - lag(ECR, k=1),
         relECR = absECR / lag(ECR, k=1)) 
      
################################
# Write Excel Tables (.xlsx):
################################
workbook <- createWorkbook() # https://rdrr.io/cran/openxlsx/man/writeData.html

addWorksheet(workbook, "Sheet1")
addWorksheet(workbook, "Sheet2")

writeData(workbook, "Sheet1", Sheet1_Option2, startCol = 1, startRow = 1)
writeData(workbook, "Sheet2", Sheet2, startCol = 1, startRow = 1)

saveWorkbook(workbook, "output.xlsx", overwrite = TRUE)