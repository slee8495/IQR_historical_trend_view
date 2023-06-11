library(tidyverse)
library(magrittr)
library(openxlsx)
library(readxl)
library(writexl)
library(reshape2)
library(skimr)
library(janitor)
library(lubridate)

# I'm still working on data collection in the folder
# requested Linda for missing files
# data 1 data 2 numbers need to be reviewed. some of them are duplicated



############ Data collection ###############

###################################################################### Data 1 ############################################################################
data_1 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2021/Finished Goods Inventory Health (IQR) - 01.06.21.xlsx",
                     sheet = "FG Jan 2021 without BKO & BKM")


data_1[-2, ] -> data_1
data_1 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_1_date 


data_1[-1, ] -> data_1
colnames(data_1) <- data_1[1, ]
data_1[-1, ] -> data_1

data_1 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_1_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_1

rm(data_1_date)

###################################################################### Data 2 ############################################################################
data_2 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2021/Finished Goods Inventory Health (IQR) - 01.13.21.xlsx",
                     sheet = "FG Jan 2021 without BKO & BKM")


data_2[-2, ] -> data_2
data_2 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_2_date 


data_2[-1, ] -> data_2
colnames(data_2) <- data_2[1, ]
data_2[-1, ] -> data_2

data_2 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_2_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_2

rm(data_2_date)

###################################################################### Data 3 ############################################################################
data_3 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2021/Finished Goods Inventory Health (IQR) - 01.20.21.xlsx",
                     sheet = "FG Jan 2021 without BKO & BKM")


data_3[-2, ] -> data_3
data_3 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_3_date 


data_3[-1, ] -> data_3
colnames(data_3) <- data_3[1, ]
data_3[-1, ] -> data_3

data_3 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_3_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_3

rm(data_3_date)

###################################################################### Data 4 ############################################################################
data_4 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2021/Finished Goods Inventory Health (IQR) - 01.27.21.xlsx",
                     sheet = "FG Jan 2021 without BKO & BKM")


data_4[-2, ] -> data_4
data_4 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_4_date 


data_4[-1, ] -> data_4
colnames(data_4) <- data_4[1, ]
data_4[-1, ] -> data_4

data_4 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_4_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_4

rm(data_4_date)




###################################################################### Data 5 ############################################################################
data_5 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2021/Finished Goods Inventory Health (IQR) - 02.03.21.xlsx",
                     sheet = "FG Jan 2021 without BKO & BKM")


data_5[-2, ] -> data_5
data_5 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_5_date 


data_5[-1, ] -> data_5
colnames(data_5) <- data_5[1, ]
data_5[-1, ] -> data_5

data_5 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_5_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_5

rm(data_5_date)





###################################################################### Data 6 ############################################################################
data_6 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2021/Finished Goods Inventory Health (IQR) - 02.10.21.xlsx",
                     sheet = "FG Jan 2021 without BKO & BKM")


data_6[-2, ] -> data_6
data_6 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_6_date 


data_6[-1, ] -> data_6
colnames(data_6) <- data_6[1, ]
data_6[-1, ] -> data_6

data_6 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_6_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_6

rm(data_6_date)



###################################################################### Data 7 ############################################################################
data_7 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2021/Finished Goods Inventory Health (IQR) - 02.17.21.xlsx",
                     sheet = "FG Jan 2021 without BKO & BKM")


data_7[-2, ] -> data_7
data_7 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_7_date 


data_7[-1, ] -> data_7
colnames(data_7) <- data_7[1, ]
data_7[-1, ] -> data_7

data_7 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_7_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_7

rm(data_7_date)


###################################################################### Data 8 ############################################################################
data_8 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2021/Finished Goods Inventory Health (IQR) - 02.24.21.xlsx",
                     sheet = "FG Jan 2021 without BKO & BKM")


data_8[-2, ] -> data_8
data_8 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_8_date 


data_8[-1, ] -> data_8
colnames(data_8) <- data_8[1, ]
data_8[-1, ] -> data_8

data_8 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_8_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_8

rm(data_8_date)



###################################################################### Data 9 ############################################################################
data_9 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2021/Finished Goods Inventory Health (IQR) - 03.03.21.xlsx",
                     sheet = "FG Jan 2021 without BKO & BKM")


data_9[-2, ] -> data_9
data_9 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_9_date 


data_9[-1, ] -> data_9
colnames(data_9) <- data_9[1, ]
data_9[-1, ] -> data_9

data_9 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_9_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_9

rm(data_9_date)




###################################################################### Data 10 ############################################################################
data_10 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2021/Finished Goods Inventory Health (IQR) - 03.10.21.xlsx",
                     sheet = "FG Jan 2021 without BKO & BKM")


data_10[-2, ] -> data_10
data_10 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_10_date 


data_10[-1, ] -> data_10
colnames(data_10) <- data_10[1, ]
data_10[-1, ] -> data_10

data_10 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_10_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_10

rm(data_10_date)


###################################################################### Data 11 ############################################################################
data_11 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2021/Finished Goods Inventory Health (IQR) - 03.18.21.xlsx",
                     sheet = "FG Jan 2021 without BKO & BKM")


data_11[-2, ] -> data_11
data_11 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_11_date 


data_11[-1, ] -> data_11
colnames(data_11) <- data_11[1, ]
data_11[-1, ] -> data_11

data_11 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_11_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_11

rm(data_11_date)


###################################################################### Data 12 ############################################################################
data_12 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2021/Finished Goods Inventory Health (IQR) - 03.24.21.xlsx",
                     sheet = "FG Jan 2021 without BKO & BKM")


data_12[-2, ] -> data_12
data_12 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_12_date 


data_12[-1, ] -> data_12
colnames(data_12) <- data_12[1, ]
data_12[-1, ] -> data_12

data_12 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_12_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_12

rm(data_12_date)



###################################################################### Data 13 ############################################################################
data_13 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2021/Finished Goods Inventory Health (IQR) - 04.01.21.xlsx",
                     sheet = "FG Jan 2021 without BKO & BKM")


data_13[-2, ] -> data_13
data_13 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_13_date 


data_13[-1, ] -> data_13
colnames(data_13) <- data_13[1, ]
data_13[-1, ] -> data_13

data_13 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_13_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_13

rm(data_13_date)



###################################################################### Data 14 ############################################################################
data_14 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2021/Finished Goods Inventory Health (IQR) - 04.07.21.xlsx",
                     sheet = "FG Jan 2021 without BKO & BKM")


data_14[-2, ] -> data_14
data_14 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_14_date 


data_14[-1, ] -> data_14
colnames(data_14) <- data_14[1, ]
data_14[-1, ] -> data_14

data_14 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_14_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_14

rm(data_14_date)


###################################################################### Data 15 ############################################################################
data_15 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2021/Finished Goods Inventory Health (IQR) - 04.14.21.xlsx",
                      sheet = "FG Jan 2021 without BKO & BKM")


data_15[-2, ] -> data_15
data_15 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_15_date 


data_15[-1, ] -> data_15
colnames(data_15) <- data_15[1, ]
data_15[-1, ] -> data_15

data_15 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_15_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_15

rm(data_15_date)



###################################################################### Data 16 ############################################################################
data_16 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2021/Finished Goods Inventory Health (IQR) - 04.21.21.xlsx",
                      sheet = "FG Jan 2021 without BKO & BKM")


data_16[-2, ] -> data_16
data_16 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_16_date 


data_16[-1, ] -> data_16
colnames(data_16) <- data_16[1, ]
data_16[-1, ] -> data_16

data_16 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_16_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_16

rm(data_16_date)


###################################################################### Data 17 ############################################################################
data_17 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2021/Finished Goods Inventory Health (IQR) - 05.19.21.xlsx",
                      sheet = "FG Jan 2021 without BKO & BKM")


data_17[-2, ] -> data_17
data_17 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_17_date 


data_17[-1, ] -> data_17
colnames(data_17) <- data_17[1, ]
data_17[-1, ] -> data_17

data_17 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_17_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_17

rm(data_17_date)




###################################################################### Data 18 ############################################################################
data_18 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2021/Finished Goods Inventory Health (IQR) - 05.26.21.xlsx",
                      sheet = "FG Jan 2021 without BKO & BKM")


data_18[-2, ] -> data_18
data_18 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_18_date 


data_18[-1, ] -> data_18
colnames(data_18) <- data_18[1, ]
data_18[-1, ] -> data_18

data_18 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_18_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_18

rm(data_18_date)


###################################################################### Data 19 ############################################################################
data_19 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2021/Finished Goods Inventory Health (IQR) - 06.09.21.xlsx",
                      sheet = "FG Jan 2021 without BKO & BKM")


data_19[-2, ] -> data_19
data_19 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_19_date 


data_19[-1, ] -> data_19
colnames(data_19) <- data_19[1, ]
data_19[-1, ] -> data_19

data_19 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_19_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_19

rm(data_19_date)



###################################################################### Data 20 ############################################################################
data_20 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2021/Finished Goods Inventory Health (IQR) - 06.16.21.xlsx",
                      sheet = "FG Jan 2021 without BKO & BKM")


data_20[-2, ] -> data_20
data_20 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_20_date 


data_20[-1, ] -> data_20
colnames(data_20) <- data_20[1, ]
data_20[-1, ] -> data_20

data_20 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_20_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_20

rm(data_20_date)



###################################################################### Data 21 ############################################################################
data_21 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2021/Finished Goods Inventory Health (IQR) - 06.23.21.xlsx",
                      sheet = "FG Jan 2021 without BKO & BKM")


data_21[-2, ] -> data_21
data_21 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_21_date 


data_21[-1, ] -> data_21
colnames(data_21) <- data_21[1, ]
data_21[-1, ] -> data_21

data_21 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_21_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_21

rm(data_21_date)



###################################################################### Data 22 ############################################################################
data_22 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2021/Finished Goods Inventory Health (IQR) - 06.30.21.xlsx",
                      sheet = "FG Jan 2021 without BKO & BKM")


data_22[-2, ] -> data_22
data_22 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_22_date 


data_22[-1, ] -> data_22
colnames(data_22) <- data_22[1, ]
data_22[-1, ] -> data_22

data_22 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_22_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_22

rm(data_22_date)



###################################################################### Data 23 ############################################################################
data_23 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2021/Finished Goods Inventory Health (IQR) - 07.07.21.xlsx",
                      sheet = "FG Jan 2021 without BKO & BKM")


data_23[-2, ] -> data_23
data_23 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_23_date 


data_23[-1, ] -> data_23
colnames(data_23) <- data_23[1, ]
data_23[-1, ] -> data_23

data_23 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_23_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_23

rm(data_23_date)



###################################################################### Data 24 ############################################################################
data_24 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2021/Finished Goods Inventory Health Adjusted Forward (IQR) - 07.14.21.xlsx",
                      sheet = "FG Jan 2021 without BKO & BKM")


data_24[-2, ] -> data_24
data_24 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_24_date 


data_24[-1, ] -> data_24
colnames(data_24) <- data_24[1, ]
data_24[-1, ] -> data_24

data_24 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_24_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_24

rm(data_24_date)



###################################################################### Data 25 ############################################################################
data_25 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2021/Finished Goods Inventory Health Adjusted Forward (IQR) - 07.23.21.xlsx",
                      sheet = "FG Jan 2021 without BKO & BKM")


data_25[-2, ] -> data_25
data_25 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_25_date 


data_25[-1, ] -> data_25
colnames(data_25) <- data_25[1, ]
data_25[-1, ] -> data_25

data_25 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_25_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_25

rm(data_25_date)



###################################################################### Data 26 ############################################################################
data_26 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2021/Finished Goods Inventory Health Adjusted Forward (IQR) - 07.28.21.xlsx",
                      sheet = "FG Jan 2021 without BKO & BKM")


data_26[-2, ] -> data_26
data_26 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_26_date 


data_26[-1, ] -> data_26
colnames(data_26) <- data_26[1, ]
data_26[-1, ] -> data_26

data_26 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_26_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_26

rm(data_26_date)



###################################################################### Data 27 ############################################################################
data_27 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2021/Finished Goods Inventory Health Adjusted Forward (IQR) - 08.11.21.xlsx",
                      sheet = "FG Jan 2021 without BKO & BKM")


data_27[-2, ] -> data_27
data_27 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_27_date 


data_27[-1, ] -> data_27
colnames(data_27) <- data_27[1, ]
data_27[-1, ] -> data_27

data_27 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_27_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_27

rm(data_27_date)



###################################################################### Data 28 ############################################################################
data_28 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2021/Finished Goods Inventory Health Adjusted Forward (IQR) - 08.18.21.xlsx",
                      sheet = "FG Jan 2021 without BKO & BKM")


data_28[-2, ] -> data_28
data_28 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_28_date 


data_28[-1, ] -> data_28
colnames(data_28) <- data_28[1, ]
data_28[-1, ] -> data_28

data_28 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_28_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_28

rm(data_28_date)



###################################################################### Data 29 ############################################################################
data_29 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2021/Finished Goods Inventory Health Adjusted Forward (IQR) - 08.18.21.xlsx",
                      sheet = "FG Jan 2021 without BKO & BKM")


data_29[-2, ] -> data_29
data_29 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_29_date 


data_29[-1, ] -> data_29
colnames(data_29) <- data_29[1, ]
data_29[-1, ] -> data_29

data_29 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_29_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_29

rm(data_29_date)


###################################################################### Data 30 ############################################################################
data_30 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2021/Finished Goods Inventory Health Adjusted Forward (IQR) - 08.25.21.xlsx",
                      sheet = "FG Jan 2021 without BKO & BKM")


data_30[-2, ] -> data_30
data_30 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_30_date 


data_30[-1, ] -> data_30
colnames(data_30) <- data_30[1, ]
data_30[-1, ] -> data_30

data_30 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_30_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_30

rm(data_30_date)



###################################################################### Data 31 ############################################################################
data_31 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2021/Finished Goods Inventory Health Adjusted Forward (IQR) - 09.01.21.xlsx",
                      sheet = "FG Jan 2021 without BKO & BKM")


data_31[-2, ] -> data_31
data_31 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_31_date 


data_31[-1, ] -> data_31
colnames(data_31) <- data_31[1, ]
data_31[-1, ] -> data_31

data_31 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_31_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_31

rm(data_31_date)



###################################################################### Data 32 ############################################################################
data_32 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2021/Finished Goods Inventory Health Adjusted Forward (IQR) - 09.08.21.xlsx",
                      sheet = "FG Jan 2021 without BKO & BKM")


data_32[-2, ] -> data_32
data_32 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_32_date 


data_32[-1, ] -> data_32
colnames(data_32) <- data_32[1, ]
data_32[-1, ] -> data_32

data_32 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_32_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_32

rm(data_32_date)




###################################################################### Data 33 ############################################################################
data_33 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2021/Finished Goods Inventory Health Adjusted Forward (IQR) - 09.15.21.xlsx",
                      sheet = "FG Jan 2021 without BKO & BKM")


data_33[-2, ] -> data_33
data_33 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_33_date 


data_33[-1, ] -> data_33
colnames(data_33) <- data_33[1, ]
data_33[-1, ] -> data_33

data_33 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_33_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_33

rm(data_33_date)




###################################################################### Data 34 ############################################################################
data_34 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2021/Finished Goods Inventory Health Adjusted Forward (IQR) - 09.22.21.xlsx",
                      sheet = "FG Jan 2021 without BKO & BKM")


data_34[-2, ] -> data_34
data_34 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_34_date 


data_34[-1, ] -> data_34
colnames(data_34) <- data_34[1, ]
data_34[-1, ] -> data_34

data_34 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_34_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_34

rm(data_34_date)



###################################################################### Data 35 ############################################################################
data_35 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2021/Finished Goods Inventory Health Adjusted Forward (IQR) - 09.29.21.xlsx",
                      sheet = "FG without BKO BKM TST")


data_35[-2, ] -> data_35
data_35 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_35_date 


data_35[-1, ] -> data_35
colnames(data_35) <- data_35[1, ]
data_35[-1, ] -> data_35

data_35 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_35_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_35

rm(data_35_date)



###################################################################### Data 36 ############################################################################
data_36 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2021/Finished Goods Inventory Health Adjusted Forward (IQR) - 10.06.21.xlsx",
                      sheet = "FG without BKO BKM TST")


data_36[-2, ] -> data_36
data_36 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_36_date 


data_36[-1, ] -> data_36
colnames(data_36) <- data_36[1, ]
data_36[-1, ] -> data_36

data_36 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_36_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_36

rm(data_36_date)




###################################################################### Data 37 ############################################################################
data_37 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2021/Finished Goods Inventory Health Adjusted Forward (IQR) - 10.13.21.xlsx",
                      sheet = "FG without BKO BKM TST")


data_37[-2, ] -> data_37
data_37 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_37_date 


data_37[-1, ] -> data_37
colnames(data_37) <- data_37[1, ]
data_37[-1, ] -> data_37

data_37 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_37_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_37

rm(data_37_date)



###################################################################### Data 38 ############################################################################
data_38 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2021/Finished Goods Inventory Health Adjusted Forward (IQR) - 10.20.21.xlsx",
                      sheet = "FG without BKO BKM TST")


data_38[-2, ] -> data_38
data_38 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_38_date 


data_38[-1, ] -> data_38
colnames(data_38) <- data_38[1, ]
data_38[-1, ] -> data_38

data_38 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_38_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_38

rm(data_38_date)


###################################################################### Data 39 ############################################################################
data_39 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2021/Finished Goods Inventory Health Adjusted Forward (IQR) - 10.27.21.xlsx",
                      sheet = "FG without BKO BKM TST")


data_39[-2, ] -> data_39
data_39 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_39_date 


data_39[-1, ] -> data_39
colnames(data_39) <- data_39[1, ]
data_39[-1, ] -> data_39

data_39 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_39_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_39

rm(data_39_date)


###################################################################### Data 40 ############################################################################
data_40 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2021/Finished Goods Inventory Health Adjusted Forward (IQR) - 11.03.21.xlsx",
                      sheet = "FG without BKO BKM TST")


data_40[-2, ] -> data_40
data_40 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_40_date 


data_40[-1, ] -> data_40
colnames(data_40) <- data_40[1, ]
data_40[-1, ] -> data_40

data_40 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_40_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_40

rm(data_40_date)


###################################################################### Data 41 ############################################################################
data_41 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2021/Finished Goods Inventory Health Adjusted Forward (IQR) - 11.10.21.xlsx",
                      sheet = "FG without BKO BKM TST")


data_41[-2, ] -> data_41
data_41 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_41_date 


data_41[-1, ] -> data_41
colnames(data_41) <- data_41[1, ]
data_41[-1, ] -> data_41

data_41 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_41_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_41

rm(data_41_date)



###################################################################### Data 42 ############################################################################
data_42 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2021/Finished Goods Inventory Health Adjusted Forward (IQR) - 11.17.21.xlsx",
                      sheet = "FG without BKO BKM TST")


data_42[-2, ] -> data_42
data_42 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_42_date 


data_42[-1, ] -> data_42
colnames(data_42) <- data_42[1, ]
data_42[-1, ] -> data_42

data_42 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_42_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_42

rm(data_42_date)


###################################################################### Data 43 ############################################################################
data_43 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2021/Finished Goods Inventory Health Adjusted Forward (IQR) - 11.24.21.xlsx",
                      sheet = "FG without BKO BKM TST")


data_43[-2, ] -> data_43
data_43 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_43_date 


data_43[-1, ] -> data_43
colnames(data_43) <- data_43[1, ]
data_43[-1, ] -> data_43

data_43 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_43_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_43

rm(data_43_date)


###################################################################### Data 44 ############################################################################
data_44 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2021/Finished Goods Inventory Health Adjusted Forward (IQR) - 12.01.21.xlsx",
                      sheet = "FG without BKO BKM TST")


data_44[-2, ] -> data_44
data_44 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_44_date 


data_44[-1, ] -> data_44
colnames(data_44) <- data_44[1, ]
data_44[-1, ] -> data_44

data_44 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_44_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_44

rm(data_44_date)


###################################################################### Data 45 ############################################################################
data_45 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2021/Finished Goods Inventory Health Adjusted Forward (IQR) - 12.08.21.xlsx",
                      sheet = "FG without BKO BKM TST")


data_45[-2, ] -> data_45
data_45 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_45_date 


data_45[-1, ] -> data_45
colnames(data_45) <- data_45[1, ]
data_45[-1, ] -> data_45

data_45 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_45_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_45

rm(data_45_date)


###################################################################### Data 46 ############################################################################
data_46 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2021/Finished Goods Inventory Health Adjusted Forward (IQR) - 12.15.21.xlsx",
                      sheet = "FG without BKO BKM TST")


data_46[-2, ] -> data_46
data_46 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_46_date 


data_46[-1, ] -> data_46
colnames(data_46) <- data_46[1, ]
data_46[-1, ] -> data_46

data_46 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_46_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_46

rm(data_46_date)


###################################################################### Data 47 ############################################################################
data_47 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2021/Finished Goods Inventory Health Adjusted Forward (IQR) - 12.22.21.xlsx",
                      sheet = "FG without BKO BKM TST")


data_47[-2, ] -> data_47
data_47 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_47_date 


data_47[-1, ] -> data_47
colnames(data_47) <- data_47[1, ]
data_47[-1, ] -> data_47

data_47 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_47_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_47

rm(data_47_date)


###################################################################### Data 48 ############################################################################
data_48 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2021/Finished Goods Inventory Health Adjusted Forward (IQR) - 12.29.21.xlsx",
                      sheet = "FG without BKO BKM TST")


data_48[-2, ] -> data_48
data_48 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_48_date 


data_48[-1, ] -> data_48
colnames(data_48) <- data_48[1, ]
data_48[-1, ] -> data_48

data_48 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_48_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_48

rm(data_48_date)


###################################################################### Data 49 ############################################################################
data_49 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2022/Finished Goods Inventory Health Adjusted Forward (IQR) - 01.05.22.xlsx",
                      sheet = "FG without BKO BKM TST")


data_49[-2, ] -> data_49
data_49 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_49_date 


data_49[-1, ] -> data_49
colnames(data_49) <- data_49[1, ]
data_49[-1, ] -> data_49

data_49 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_49_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_49

rm(data_49_date)


###################################################################### Data 50 ############################################################################
data_50 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2022/Finished Goods Inventory Health Adjusted Forward (IQR) - 01.12.22.xlsx",
                      sheet = "FG without BKO BKM TST")


data_50[-2, ] -> data_50
data_50 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_50_date 


data_50[-1, ] -> data_50
colnames(data_50) <- data_50[1, ]
data_50[-1, ] -> data_50

data_50 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_50_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_50

rm(data_50_date)


###################################################################### Data 51 ############################################################################
data_51 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2022/Finished Goods Inventory Health Adjusted Forward (IQR) - 01.19.22.xlsx",
                      sheet = "FG without BKO BKM TST")


data_51[-2, ] -> data_51
data_51 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_51_date 


data_51[-1, ] -> data_51
colnames(data_51) <- data_51[1, ]
data_51[-1, ] -> data_51

data_51 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_51_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_51

rm(data_51_date)


###################################################################### Data 52 ############################################################################
data_52 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2022/Finished Goods Inventory Health Adjusted Forward (IQR) - 01.26.22.xlsx",
                      sheet = "FG without BKO BKM TST")


data_52[-2, ] -> data_52
data_52 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_52_date 


data_52[-1, ] -> data_52
colnames(data_52) <- data_52[1, ]
data_52[-1, ] -> data_52

data_52 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_52_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_52

rm(data_52_date)


###################################################################### Data 53 ############################################################################
data_53 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2022/Finished Goods Inventory Health Adjusted Forward (IQR) - 02.02.22.xlsx",
                      sheet = "FG without BKO BKM TST")


data_53[-2, ] -> data_53
data_53 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_53_date 


data_53[-1, ] -> data_53
colnames(data_53) <- data_53[1, ]
data_53[-1, ] -> data_53

data_53 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_53_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_53

rm(data_53_date)


###################################################################### Data 54 ############################################################################
data_54 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2022/Finished Goods Inventory Health Adjusted Forward (IQR) - 02.09.22.xlsx",
                      sheet = "FG without BKO BKM TST")


data_54[-2, ] -> data_54
data_54 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_54_date 


data_54[-1, ] -> data_54
colnames(data_54) <- data_54[1, ]
data_54[-1, ] -> data_54

data_54 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_54_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_54

rm(data_54_date)


###################################################################### Data 55 ############################################################################
data_55 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2022/Finished Goods Inventory Health Adjusted Forward (IQR) - 02.16.22.xlsx",
                      sheet = "FG without BKO BKM TST")


data_55[-2, ] -> data_55
data_55 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_55_date 


data_55[-1, ] -> data_55
colnames(data_55) <- data_55[1, ]
data_55[-1, ] -> data_55

data_55 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_55_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_55

rm(data_55_date)


###################################################################### Data 56 ############################################################################
data_56 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2022/Finished Goods Inventory Health Adjusted Forward (IQR) - 02.23.22.xlsx",
                      sheet = "FG without BKO BKM TST")


data_56[-2, ] -> data_56
data_56 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_56_date 


data_56[-1, ] -> data_56
colnames(data_56) <- data_56[1, ]
data_56[-1, ] -> data_56

data_56 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_56_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_56

rm(data_56_date)


###################################################################### Data 57 ############################################################################
data_57 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2022/Finished Goods Inventory Health Adjusted Forward (IQR) - 03.02.22.xlsx",
                      sheet = "FG without BKO BKM TST")


data_57[-2, ] -> data_57
data_57 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_57_date 


data_57[-1, ] -> data_57
colnames(data_57) <- data_57[1, ]
data_57[-1, ] -> data_57

data_57 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_57_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_57

rm(data_57_date)


###################################################################### Data 58 ############################################################################
data_58 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2022/Finished Goods Inventory Health Adjusted Forward (IQR) - 03.09.22.xlsx",
                      sheet = "FG without BKO BKM TST")


data_58[-2, ] -> data_58
data_58 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_58_date 


data_58[-1, ] -> data_58
colnames(data_58) <- data_58[1, ]
data_58[-1, ] -> data_58

data_58 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_58_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_58

rm(data_58_date)


###################################################################### Data 59 ############################################################################
data_59 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2022/Finished Goods Inventory Health Adjusted Forward (IQR) - 03.16.22.xlsx",
                      sheet = "FG without BKO BKM TST")


data_59[-2, ] -> data_59
data_59 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_59_date 


data_59[-1, ] -> data_59
colnames(data_59) <- data_59[1, ]
data_59[-1, ] -> data_59

data_59 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_59_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_59

rm(data_59_date)


###################################################################### Data 60 ############################################################################
data_60 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2022/Finished Goods Inventory Health Adjusted Forward (IQR) - 03.23.22.xlsx",
                      sheet = "FG without BKO BKM TST")


data_60[-2, ] -> data_60
data_60 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_60_date 


data_60[-1, ] -> data_60
colnames(data_60) <- data_60[1, ]
data_60[-1, ] -> data_60

data_60 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_60_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_60

rm(data_60_date)


###################################################################### Data 61 ############################################################################
data_61 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2022/Finished Goods Inventory Health Adjusted Forward (IQR) - 03.30.22.xlsx",
                      sheet = "FG without BKO BKM TST")


data_61[-2, ] -> data_61
data_61 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_61_date 


data_61[-1, ] -> data_61
colnames(data_61) <- data_61[1, ]
data_61[-1, ] -> data_61

data_61 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_61_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_61

rm(data_61_date)


###################################################################### Data 62 ############################################################################
data_62 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2022/Finished Goods Inventory Health Adjusted Forward (IQR) - 04.13.22.xlsx",
                      sheet = "FG without BKO BKM TST")


data_62[-2, ] -> data_62
data_62 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_62_date 


data_62[-1, ] -> data_62
colnames(data_62) <- data_62[1, ]
data_62[-1, ] -> data_62

data_62 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_62_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_62

rm(data_62_date)


###################################################################### Data 63 ############################################################################
data_63 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2022/Finished Goods Inventory Health Adjusted Forward (IQR) - 04.20.22.xlsx",
                      sheet = "FG without BKO BKM TST")


data_63[-2, ] -> data_63
data_63 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_63_date 


data_63[-1, ] -> data_63
colnames(data_63) <- data_63[1, ]
data_63[-1, ] -> data_63

data_63 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_63_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_63

rm(data_63_date)


###################################################################### Data 64 ############################################################################
data_64 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2022/Finished Goods Inventory Health Adjusted Forward (IQR) - 04.27.22.xlsx",
                      sheet = "FG without BKO BKM TST")


data_64[-2, ] -> data_64
data_64 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_64_date 


data_64[-1, ] -> data_64
colnames(data_64) <- data_64[1, ]
data_64[-1, ] -> data_64

data_64 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_64_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_64

rm(data_64_date)


###################################################################### Data 65 ############################################################################
data_65 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2022/Finished Goods Inventory Health Adjusted Forward (IQR) - 05.04.22.xlsx",
                      sheet = "FG without BKO BKM TST")


data_65[-2, ] -> data_65
data_65 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_65_date 


data_65[-1, ] -> data_65
colnames(data_65) <- data_65[1, ]
data_65[-1, ] -> data_65

data_65 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_65_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_65

rm(data_65_date)


###################################################################### Data 66 ############################################################################
data_66 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2022/Finished Goods Inventory Health Adjusted Forward (IQR) - 05.11.22.xlsx",
                      sheet = "FG without BKO BKM TST")


data_66[-2, ] -> data_66
data_66 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_66_date 


data_66[-1, ] -> data_66
colnames(data_66) <- data_66[1, ]
data_66[-1, ] -> data_66

data_66 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_66_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_66

rm(data_66_date)


###################################################################### Data 67 ############################################################################
data_67 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2022/Finished Goods Inventory Health Adjusted Forward (IQR) - 05.18.22.xlsx",
                      sheet = "FG without BKO BKM TST")


data_67[-2, ] -> data_67
data_67 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_67_date 


data_67[-1, ] -> data_67
colnames(data_67) <- data_67[1, ]
data_67[-1, ] -> data_67

data_67 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_67_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_67

rm(data_67_date)


###################################################################### Data 68 ############################################################################
data_68 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2022/Finished Goods Inventory Health Adjusted Forward (IQR) - 05.25.22.xlsx",
                      sheet = "FG without BKO BKM TST")


data_68[-2, ] -> data_68
data_68 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_68_date 


data_68[-1, ] -> data_68
colnames(data_68) <- data_68[1, ]
data_68[-1, ] -> data_68

data_68 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_68_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_68

rm(data_68_date)


###################################################################### Data 69 ############################################################################
data_69 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2022/Finished Goods Inventory Health Adjusted Forward (IQR) - 06.01.22.xlsx",
                      sheet = "FG without BKO BKM TST")


data_69[-2, ] -> data_69
data_69 %>% 
  janitor::clean_names() %>% 
  dplyr::select(1:2) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x2 = as.integer(x2),
                x2 = as.Date(x2, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_69_date 


data_69[-1, ] -> data_69
colnames(data_69) <- data_69[1, ]
data_69[-1, ] -> data_69

data_69 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_69_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_69

rm(data_69_date)


###################################################################### Data 70 ############################################################################
data_70 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2022/Finished Goods Inventory Health Adjusted Forward (IQR) - 06.08.22.xlsx",
                      sheet = "FG without BKO BKM TST")


data_70[-2, ] -> data_70
data_70 %>% 
  janitor::clean_names() %>% 
  dplyr::select(c(1, 3)) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x3 = as.integer(x3),
                x3 = as.Date(x3, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_70_date 


data_70[-1, ] -> data_70
colnames(data_70) <- data_70[1, ]
data_70[-1, ] -> data_70

data_70 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_70_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_70

rm(data_70_date)


###################################################################### Data 71 ############################################################################
data_71 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2022/Finished Goods Inventory Health Adjusted Forward (IQR) - 06.15.22.xlsx",
                      sheet = "FG without BKO BKM TST")


data_71[-2, ] -> data_71
data_71 %>% 
  janitor::clean_names() %>% 
  dplyr::select(c(1, 3)) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x3 = as.integer(x3),
                x3 = as.Date(x3, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_71_date 


data_71[-1, ] -> data_71
colnames(data_71) <- data_71[1, ]
data_71[-1, ] -> data_71

data_71 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_71_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_71

rm(data_71_date)


###################################################################### Data 72 ############################################################################
data_72 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2022/Finished Goods Inventory Health Adjusted Forward (IQR) - 06.22.22.xlsx",
                      sheet = "FG without BKO BKM TST")


data_72[-2, ] -> data_72
data_72 %>% 
  janitor::clean_names() %>% 
  dplyr::select(c(1, 3)) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x3 = as.integer(x3),
                x3 = as.Date(x3, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_72_date 


data_72[-1, ] -> data_72
colnames(data_72) <- data_72[1, ]
data_72[-1, ] -> data_72

data_72 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_72_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_72

rm(data_72_date)


###################################################################### Data 73 ############################################################################
data_73 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2022/Finished Goods Inventory Health Adjusted Forward (IQR) - 06.29.22.xlsx",
                      sheet = "FG without BKO BKM TST")


data_73[-2, ] -> data_73
data_73 %>% 
  janitor::clean_names() %>% 
  dplyr::select(c(1, 3)) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x3 = as.integer(x3),
                x3 = as.Date(x3, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_73_date 


data_73[-1, ] -> data_73
colnames(data_73) <- data_73[1, ]
data_73[-1, ] -> data_73

data_73 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_73_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_73

rm(data_73_date)


###################################################################### Data 74 ############################################################################
data_74 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2022/Finished Goods Inventory Health Adjusted Forward (IQR) - 07.06.22.xlsx",
                      sheet = "FG without BKO BKM TST")


data_74[-2, ] -> data_74
data_74 %>% 
  janitor::clean_names() %>% 
  dplyr::select(c(1, 3)) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x3 = as.integer(x3),
                x3 = as.Date(x3, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_74_date 


data_74[-1, ] -> data_74
colnames(data_74) <- data_74[1, ]
data_74[-1, ] -> data_74

data_74 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_74_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_74

rm(data_74_date)


###################################################################### Data 75 ############################################################################
data_75 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2022/Finished Goods Inventory Health Adjusted Forward (IQR) - 07.13.22.xlsx",
                      sheet = "FG without BKO BKM TST")


data_75[-2, ] -> data_75
data_75 %>% 
  janitor::clean_names() %>% 
  dplyr::select(c(1, 3)) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x3 = as.integer(x3),
                x3 = as.Date(x3, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_75_date 


data_75[-1, ] -> data_75
colnames(data_75) <- data_75[1, ]
data_75[-1, ] -> data_75

data_75 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_75_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_75

rm(data_75_date)


###################################################################### Data 76 ############################################################################
data_76 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2022/Finished Goods Inventory Health Adjusted Forward (IQR) - 07.20.22.xlsx",
                      sheet = "FG without BKO BKM TST")


data_76[-2, ] -> data_76
data_76 %>% 
  janitor::clean_names() %>% 
  dplyr::select(c(1, 3)) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x3 = as.integer(x3),
                x3 = as.Date(x3, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_76_date 


data_76[-1, ] -> data_76
colnames(data_76) <- data_76[1, ]
data_76[-1, ] -> data_76

data_76 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_76_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_76

rm(data_76_date)


###################################################################### Data 77 ############################################################################
data_77 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2022/Finished Goods Inventory Health Adjusted Forward (IQR) - 07.27.22.xlsx",
                      sheet = "FG without BKO BKM TST")


data_77[-2, ] -> data_77
data_77 %>% 
  janitor::clean_names() %>% 
  dplyr::select(c(1, 3)) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x3 = as.integer(x3),
                x3 = as.Date(x3, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_77_date 


data_77[-1, ] -> data_77
colnames(data_77) <- data_77[1, ]
data_77[-1, ] -> data_77

data_77 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_77_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_77

rm(data_77_date)



###################################################################### Data 78 ############################################################################
data_78 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2022/Finished Goods Inventory Health Adjusted Forward (IQR) - 08.03.22.xlsx",
                      sheet = "FG without BKO BKM TST")


data_78[-2, ] -> data_78
data_78 %>% 
  janitor::clean_names() %>% 
  dplyr::select(c(1, 3)) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x3 = as.integer(x3),
                x3 = as.Date(x3, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_78_date 


data_78[-1, ] -> data_78
colnames(data_78) <- data_78[1, ]
data_78[-1, ] -> data_78

data_78 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_78_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_78

rm(data_78_date)


###################################################################### Data 79 ############################################################################
data_79 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2022/Finished Goods Inventory Health Adjusted Forward (IQR) - 08.10.22.xlsx",
                      sheet = "FG without BKO BKM TST")


data_79[-2, ] -> data_79
data_79 %>% 
  janitor::clean_names() %>% 
  dplyr::select(c(1, 3)) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x3 = as.integer(x3),
                x3 = as.Date(x3, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_79_date 


data_79[-1, ] -> data_79
colnames(data_79) <- data_79[1, ]
data_79[-1, ] -> data_79

data_79 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_79_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_79

rm(data_79_date)


###################################################################### Data 80 ############################################################################
data_80 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2022/Finished Goods Inventory Health Adjusted Forward (IQR) - 08.17.22.xlsx",
                      sheet = "FG without BKO BKM TST")


data_80[-2, ] -> data_80
data_80 %>% 
  janitor::clean_names() %>% 
  dplyr::select(c(1, 3)) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x3 = as.integer(x3),
                x3 = as.Date(x3, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_80_date 


data_80[-1, ] -> data_80
colnames(data_80) <- data_80[1, ]
data_80[-1, ] -> data_80

data_80 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_80_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_80

rm(data_80_date)


###################################################################### Data 81 ############################################################################
data_81 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2022/Finished Goods Inventory Health Adjusted Forward (IQR) - 08.24.22.xlsx",
                      sheet = "FG without BKO BKM TST")


data_81[-2, ] -> data_81
data_81 %>% 
  janitor::clean_names() %>% 
  dplyr::select(c(1, 3)) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x3 = as.integer(x3),
                x3 = as.Date(x3, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_81_date 


data_81[-1, ] -> data_81
colnames(data_81) <- data_81[1, ]
data_81[-1, ] -> data_81

data_81 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_81_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_81

rm(data_81_date)


###################################################################### Data 82 ############################################################################
data_82 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2022/Finished Goods Inventory Health Adjusted Forward (IQR) - 08.29.22.xlsx",
                      sheet = "FG without BKO BKM TST")


data_82[-2, ] -> data_82
data_82 %>% 
  janitor::clean_names() %>% 
  dplyr::select(c(1, 3)) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x3 = as.integer(x3),
                x3 = as.Date(x3, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_82_date 


data_82[-1, ] -> data_82
colnames(data_82) <- data_82[1, ]
data_82[-1, ] -> data_82

data_82 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_82_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_82

rm(data_82_date)


###################################################################### Data 83 ############################################################################
data_83 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2022/Finished Goods Inventory Health Adjusted Forward (IQR) - 09.02.22.xlsx",
                      sheet = "FG without BKO BKM TST")


data_83[-2, ] -> data_83
data_83 %>% 
  janitor::clean_names() %>% 
  dplyr::select(c(1, 3)) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x3 = as.integer(x3),
                x3 = as.Date(x3, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_83_date 


data_83[-1, ] -> data_83
colnames(data_83) <- data_83[1, ]
data_83[-1, ] -> data_83

data_83 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_83_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_83

rm(data_83_date)


###################################################################### Data 84 ############################################################################
data_84 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2022/Finished Goods Inventory Health Adjusted Forward (IQR) - 09.14.22.xlsx",
                      sheet = "FG without BKO BKM TST")


data_84[-2, ] -> data_84
data_84 %>% 
  janitor::clean_names() %>% 
  dplyr::select(c(1, 3)) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x3 = as.integer(x3),
                x3 = as.Date(x3, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_84_date 


data_84[-1, ] -> data_84
colnames(data_84) <- data_84[1, ]
data_84[-1, ] -> data_84

data_84 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_84_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_84

rm(data_84_date)


###################################################################### Data 85 ############################################################################
data_85 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2022/Finished Goods Inventory Health Adjusted Forward (IQR) - 09.21.22.xlsx",
                      sheet = "FG without BKO BKM TST")


data_85[-2, ] -> data_85
data_85 %>% 
  janitor::clean_names() %>% 
  dplyr::select(c(1, 3)) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x3 = as.integer(x3),
                x3 = as.Date(x3, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_85_date 


data_85[-1, ] -> data_85
colnames(data_85) <- data_85[1, ]
data_85[-1, ] -> data_85

data_85 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_85_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_85

rm(data_85_date)


###################################################################### Data 86 ############################################################################
data_86 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2022/Finished Goods Inventory Health Adjusted Forward (IQR) - 09.28.22.xlsx",
                      sheet = "FG without BKO BKM TST")


data_86[-2, ] -> data_86
data_86 %>% 
  janitor::clean_names() %>% 
  dplyr::select(c(1, 3)) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x3 = as.integer(x3),
                x3 = as.Date(x3, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_86_date 


data_86[-1, ] -> data_86
colnames(data_86) <- data_86[1, ]
data_86[-1, ] -> data_86

data_86 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_86_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_86

rm(data_86_date)


###################################################################### Data 87 ############################################################################
data_87 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2022/Finished Goods Inventory Health Adjusted Forward (IQR) - 10.05.22.xlsx",
                      sheet = "FG without BKO BKM TST")


data_87[-2, ] -> data_87
data_87 %>% 
  janitor::clean_names() %>% 
  dplyr::select(c(1, 3)) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x3 = as.integer(x3),
                x3 = as.Date(x3, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_87_date 


data_87[-1, ] -> data_87
colnames(data_87) <- data_87[1, ]
data_87[-1, ] -> data_87

data_87 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_87_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_87

rm(data_87_date)


###################################################################### Data 88 ############################################################################
data_88 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2022/Finished Goods Inventory Health Adjusted Forward (IQR) - 10.12.22.xlsx",
                      sheet = "FG without BKO BKM TST")


data_88[-2, ] -> data_88
data_88 %>% 
  janitor::clean_names() %>% 
  dplyr::select(c(1, 3)) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x3 = as.integer(x3),
                x3 = as.Date(x3, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_88_date 


data_88[-1, ] -> data_88
colnames(data_88) <- data_88[1, ]
data_88[-1, ] -> data_88

data_88 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_88_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_88

rm(data_88_date)


###################################################################### Data 89 ############################################################################
data_89 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2022/Finished Goods Inventory Health Adjusted Forward (IQR) - 10.19.22.xlsx",
                      sheet = "FG without BKO BKM TST")


data_89[-2, ] -> data_89
data_89 %>% 
  janitor::clean_names() %>% 
  dplyr::select(c(1, 3)) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x3 = as.integer(x3),
                x3 = as.Date(x3, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_89_date 


data_89[-1, ] -> data_89
colnames(data_89) <- data_89[1, ]
data_89[-1, ] -> data_89

data_89 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_89_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_89

rm(data_89_date)


###################################################################### Data 90 ############################################################################
data_90 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2022/Finished Goods Inventory Health Adjusted Forward (IQR) - 10.26.22.xlsx",
                      sheet = "FG without BKO BKM TST")


data_90[-2, ] -> data_90
data_90 %>% 
  janitor::clean_names() %>% 
  dplyr::select(c(1, 3)) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x3 = as.integer(x3),
                x3 = as.Date(x3, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_90_date 


data_90[-1, ] -> data_90
colnames(data_90) <- data_90[1, ]
data_90[-1, ] -> data_90

data_90 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_90_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_90

rm(data_90_date)


###################################################################### Data 91 ############################################################################
data_91 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2022/Finished Goods Inventory Health Adjusted Forward (IQR) - 11.02.22.xlsx",
                      sheet = "FG without BKO BKM TST")


data_91[-2, ] -> data_91
data_91 %>% 
  janitor::clean_names() %>% 
  dplyr::select(c(1, 3)) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x3 = as.integer(x3),
                x3 = as.Date(x3, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_91_date 


data_91[-1, ] -> data_91
colnames(data_91) <- data_91[1, ]
data_91[-1, ] -> data_91

data_91 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_91_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_91

rm(data_91_date)


###################################################################### Data 92 ############################################################################
data_92 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2022/Finished Goods Inventory Health Adjusted Forward (IQR) - 11.09.22.xlsx",
                      sheet = "FG without BKO BKM TST")


data_92[-2, ] -> data_92
data_92 %>% 
  janitor::clean_names() %>% 
  dplyr::select(c(1, 3)) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x3 = as.integer(x3),
                x3 = as.Date(x3, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_92_date 


data_92[-1, ] -> data_92
colnames(data_92) <- data_92[1, ]
data_92[-1, ] -> data_92

data_92 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_92_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_92

rm(data_92_date)


###################################################################### Data 93 ############################################################################
data_93 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2022/Finished Goods Inventory Health Adjusted Forward (IQR) - 11.16.22.xlsx",
                      sheet = "FG without BKO BKM TST")


data_93[-2, ] -> data_93
data_93 %>% 
  janitor::clean_names() %>% 
  dplyr::select(c(1, 3)) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x3 = as.integer(x3),
                x3 = as.Date(x3, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_93_date 


data_93[-1, ] -> data_93
colnames(data_93) <- data_93[1, ]
data_93[-1, ] -> data_93

data_93 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_93_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_93

rm(data_93_date)


###################################################################### Data 94 ############################################################################
data_94 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2022/Finished Goods Inventory Health Adjusted Forward (IQR) - 11.21.22.xlsx",
                      sheet = "FG without BKO BKM TST")


data_94[-2, ] -> data_94
data_94 %>% 
  janitor::clean_names() %>% 
  dplyr::select(c(1, 3)) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x3 = as.integer(x3),
                x3 = as.Date(x3, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_94_date 


data_94[-1, ] -> data_94
colnames(data_94) <- data_94[1, ]
data_94[-1, ] -> data_94

data_94 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_94_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_94

rm(data_94_date)


###################################################################### Data 95 ############################################################################
data_95 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2022/Finished Goods Inventory Health Adjusted Forward (IQR) - 11.30.22.xlsx",
                      sheet = "FG without BKO BKM TST")


data_95[-2, ] -> data_95
data_95 %>% 
  janitor::clean_names() %>% 
  dplyr::select(c(1, 3)) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x3 = as.integer(x3),
                x3 = as.Date(x3, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_95_date 


data_95[-1, ] -> data_95
colnames(data_95) <- data_95[1, ]
data_95[-1, ] -> data_95

data_95 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_95_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_95

rm(data_95_date)


###################################################################### Data 96 ############################################################################
data_96 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2022/Finished Goods Inventory Health Adjusted Forward (IQR) - 12.07.22.xlsx",
                      sheet = "FG without BKO BKM TST")


data_96[-2, ] -> data_96
data_96 %>% 
  janitor::clean_names() %>% 
  dplyr::select(c(1, 3)) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x3 = as.integer(x3),
                x3 = as.Date(x3, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_96_date 


data_96[-1, ] -> data_96
colnames(data_96) <- data_96[1, ]
data_96[-1, ] -> data_96

data_96 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_96_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_96

rm(data_96_date)


###################################################################### Data 97 ############################################################################
data_97 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2022/Finished Goods Inventory Health Adjusted Forward (IQR) - 12.14.22.xlsx",
                      sheet = "FG without BKO BKM TST")


data_97[-2, ] -> data_97
data_97 %>% 
  janitor::clean_names() %>% 
  dplyr::select(c(1, 3)) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x3 = as.integer(x3),
                x3 = as.Date(x3, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_97_date 


data_97[-1, ] -> data_97
colnames(data_97) <- data_97[1, ]
data_97[-1, ] -> data_97

data_97 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_97_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_97

rm(data_97_date)


###################################################################### Data 98 ############################################################################
data_98 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2022/Finished Goods Inventory Health Adjusted Forward (IQR) - 12.21.22.xlsx",
                      sheet = "FG without BKO BKM TST")


data_98[-2, ] -> data_98
data_98 %>% 
  janitor::clean_names() %>% 
  dplyr::select(c(1, 3)) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x3 = as.integer(x3),
                x3 = as.Date(x3, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_98_date 


data_98[-1, ] -> data_98
colnames(data_98) <- data_98[1, ]
data_98[-1, ] -> data_98

data_98 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_98_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_98

rm(data_98_date)


###################################################################### Data 99 ############################################################################
data_99 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2022/Finished Goods Inventory Health Adjusted Forward (IQR) - 12.28.22.xlsx",
                      sheet = "FG without BKO BKM TST")


data_99[-2, ] -> data_99
data_99 %>% 
  janitor::clean_names() %>% 
  dplyr::select(c(1, 3)) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x3 = as.integer(x3),
                x3 = as.Date(x3, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_99_date 


data_99[-1, ] -> data_99
colnames(data_99) <- data_99[1, ]
data_99[-1, ] -> data_99

data_99 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_99_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_99

rm(data_99_date)


###################################################################### Data 100 ############################################################################
data_100 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2023/Finished Goods Inventory Health Adjusted Forward (IQR) - 01.04.23.xlsx",
                      sheet = "FG without BKO BKM TST")


data_100[-2, ] -> data_100
data_100 %>% 
  janitor::clean_names() %>% 
  dplyr::select(c(1, 3)) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x3 = as.integer(x3),
                x3 = as.Date(x3, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_100_date 


data_100[-1, ] -> data_100
colnames(data_100) <- data_100[1, ]
data_100[-1, ] -> data_100

data_100 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_100_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_100

rm(data_100_date)


###################################################################### Data 101 ############################################################################
data_101 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2023/Finished Goods Inventory Health Adjusted Forward (IQR) - 01.11.23.xlsx",
                       sheet = "FG without BKO BKM TST")


data_101[-2, ] -> data_101
data_101 %>% 
  janitor::clean_names() %>% 
  dplyr::select(c(1, 3)) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x3 = as.integer(x3),
                x3 = as.Date(x3, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_101_date 


data_101[-1, ] -> data_101
colnames(data_101) <- data_101[1, ]
data_101[-1, ] -> data_101

data_101 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_101_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_101

rm(data_101_date)


###################################################################### Data 102 ############################################################################
data_102 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2023/Finished Goods Inventory Health Adjusted Forward (IQR) - 01.18.23.xlsx",
                       sheet = "FG without BKO BKM TST")


data_102[-2, ] -> data_102
data_102 %>% 
  janitor::clean_names() %>% 
  dplyr::select(c(1, 3)) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x3 = as.integer(x3),
                x3 = as.Date(x3, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_102_date 


data_102[-1, ] -> data_102
colnames(data_102) <- data_102[1, ]
data_102[-1, ] -> data_102

data_102 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_102_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_102

rm(data_102_date)


###################################################################### Data 103 ############################################################################
data_103 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2023/Finished Goods Inventory Health Adjusted Forward (IQR) - 01.25.23.xlsx",
                       sheet = "FG without BKO BKM TST")


data_103[-2, ] -> data_103
data_103 %>% 
  janitor::clean_names() %>% 
  dplyr::select(c(1, 3)) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x3 = as.integer(x3),
                x3 = as.Date(x3, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_103_date 


data_103[-1, ] -> data_103
colnames(data_103) <- data_103[1, ]
data_103[-1, ] -> data_103

data_103 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_103_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_103

rm(data_103_date)


###################################################################### Data 104 ############################################################################
data_104 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2023/Finished Goods Inventory Health Adjusted Forward (IQR) - 02.01.23.xlsx",
                       sheet = "FG without BKO BKM TST")


data_104[-2, ] -> data_104
data_104 %>% 
  janitor::clean_names() %>% 
  dplyr::select(c(1, 3)) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x3 = as.integer(x3),
                x3 = as.Date(x3, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_104_date 


data_104[-1, ] -> data_104
colnames(data_104) <- data_104[1, ]
data_104[-1, ] -> data_104

data_104 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_104_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_104

rm(data_104_date)


###################################################################### Data 105 ############################################################################
data_105 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2023/Finished Goods Inventory Health Adjusted Forward (IQR) - 02.08.23.xlsx",
                       sheet = "FG without BKO BKM TST")


data_105[-2, ] -> data_105
data_105 %>% 
  janitor::clean_names() %>% 
  dplyr::select(c(1, 3)) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x3 = as.integer(x3),
                x3 = as.Date(x3, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_105_date 


data_105[-1, ] -> data_105
colnames(data_105) <- data_105[1, ]
data_105[-1, ] -> data_105

data_105 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_105_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_105

rm(data_105_date)


###################################################################### Data 106 ############################################################################
data_106 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2023/Finished Goods Inventory Health Adjusted Forward (IQR) - 02.15.23.xlsx",
                       sheet = "FG without BKO BKM TST")


data_106[-2, ] -> data_106
data_106 %>% 
  janitor::clean_names() %>% 
  dplyr::select(c(1, 3)) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x3 = as.integer(x3),
                x3 = as.Date(x3, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_106_date 


data_106[-1, ] -> data_106
colnames(data_106) <- data_106[1, ]
data_106[-1, ] -> data_106

data_106 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_106_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_106

rm(data_106_date)


###################################################################### Data 107 ############################################################################
data_107 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2023/Finished Goods Inventory Health Adjusted Forward (IQR) - 02.22.23.xlsx",
                       sheet = "FG without BKO BKM TST")


data_107[-2, ] -> data_107
data_107 %>% 
  janitor::clean_names() %>% 
  dplyr::select(c(1, 3)) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x3 = as.integer(x3),
                x3 = as.Date(x3, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_107_date 


data_107[-1, ] -> data_107
colnames(data_107) <- data_107[1, ]
data_107[-1, ] -> data_107

data_107 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_107_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_107

rm(data_107_date)


###################################################################### Data 108 ############################################################################
data_108 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2023/Finished Goods Inventory Health Adjusted Forward (IQR) - 03.01.23.xlsx",
                       sheet = "FG without BKO BKM TST")


data_108[-2, ] -> data_108
data_108 %>% 
  janitor::clean_names() %>% 
  dplyr::select(c(1, 3)) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x3 = as.integer(x3),
                x3 = as.Date(x3, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_108_date 


data_108[-1, ] -> data_108
colnames(data_108) <- data_108[1, ]
data_108[-1, ] -> data_108

data_108 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_108_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_108

rm(data_108_date)


###################################################################### Data 109 ############################################################################
data_109 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2023/Finished Goods Inventory Health Adjusted Forward (IQR) - 03.22.23.xlsx",
                       sheet = "FG without BKO BKM TST")


data_109[-2, ] -> data_109
data_109 %>% 
  janitor::clean_names() %>% 
  dplyr::select(c(1, 3)) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x3 = as.integer(x3),
                x3 = as.Date(x3, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_109_date 


data_109[-1, ] -> data_109
colnames(data_109) <- data_109[1, ]
data_109[-1, ] -> data_109

data_109 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_109_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_109

rm(data_109_date)


###################################################################### Data 110 ############################################################################
data_110 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2023/Finished Goods Inventory Health Adjusted Forward (IQR) - 03.29.23.xlsx",
                       sheet = "FG without BKO BKM TST")


data_110[-2, ] -> data_110
data_110 %>% 
  janitor::clean_names() %>% 
  dplyr::select(c(1, 3)) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x3 = as.integer(x3),
                x3 = as.Date(x3, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_110_date 


data_110[-1, ] -> data_110
colnames(data_110) <- data_110[1, ]
data_110[-1, ] -> data_110

data_110 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_110_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_110

rm(data_110_date)


###################################################################### Data 111 ############################################################################
data_111 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2023/Finished Goods Inventory Health Adjusted Forward (IQR) - 04.05.23.xlsx",
                       sheet = "FG without BKO BKM TST")


data_111[-2, ] -> data_111
data_111 %>% 
  janitor::clean_names() %>% 
  dplyr::select(c(1, 3)) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x3 = as.integer(x3),
                x3 = as.Date(x3, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_111_date 


data_111[-1, ] -> data_111
colnames(data_111) <- data_111[1, ]
data_111[-1, ] -> data_111

data_111 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_111_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_111

rm(data_111_date)


###################################################################### Data 112 ############################################################################
data_112 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2023/Finished Goods Inventory Health Adjusted Forward (IQR) - 04.12.23.xlsx",
                       sheet = "FG without BKO BKM TST")


data_112[-2, ] -> data_112
data_112 %>% 
  janitor::clean_names() %>% 
  dplyr::select(c(1, 3)) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x3 = as.integer(x3),
                x3 = as.Date(x3, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_112_date 


data_112[-1, ] -> data_112
colnames(data_112) <- data_112[1, ]
data_112[-1, ] -> data_112

data_112 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_112_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_112

rm(data_112_date)


###################################################################### Data 113 ############################################################################
data_113 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2023/Finished Goods Inventory Health Adjusted Forward (IQR) - 04.19.23.xlsx",
                       sheet = "FG without BKO BKM TST")


data_113[-2, ] -> data_113
data_113 %>% 
  janitor::clean_names() %>% 
  dplyr::select(c(1, 3)) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x3 = as.integer(x3),
                x3 = as.Date(x3, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_113_date 


data_113[-1, ] -> data_113
colnames(data_113) <- data_113[1, ]
data_113[-1, ] -> data_113

data_113 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_113_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_113

rm(data_113_date)


###################################################################### Data 114 ############################################################################
data_114 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2023/Finished Goods Inventory Health Adjusted Forward (IQR) - 04.26.23.xlsx",
                       sheet = "FG without BKO BKM TST")


data_114[-2, ] -> data_114
data_114 %>% 
  janitor::clean_names() %>% 
  dplyr::select(c(1, 3)) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x3 = as.integer(x3),
                x3 = as.Date(x3, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_114_date 


data_114[-1, ] -> data_114
colnames(data_114) <- data_114[1, ]
data_114[-1, ] -> data_114

data_114 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_114_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_114

rm(data_114_date)


###################################################################### Data 115 ############################################################################
data_115 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2023/Finished Goods Inventory Health Adjusted Forward (IQR) - 05.04.23.xlsx",
                       sheet = "FG without BKO BKM TST")


data_115[-2, ] -> data_115
data_115 %>% 
  janitor::clean_names() %>% 
  dplyr::select(c(1, 3)) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x3 = as.integer(x3),
                x3 = as.Date(x3, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_115_date 


data_115[-1, ] -> data_115
colnames(data_115) <- data_115[1, ]
data_115[-1, ] -> data_115

data_115 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_115_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_115

rm(data_115_date)


###################################################################### Data 116 ############################################################################
data_116 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2023/Finished Goods Inventory Health Adjusted Forward (IQR) - 05.10.23.xlsx",
                       sheet = "FG without BKO BKM TST")


data_116[-2, ] -> data_116
data_116 %>% 
  janitor::clean_names() %>% 
  dplyr::select(c(1, 3)) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x3 = as.integer(x3),
                x3 = as.Date(x3, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_116_date 


data_116[-1, ] -> data_116
colnames(data_116) <- data_116[1, ]
data_116[-1, ] -> data_116

data_116 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_116_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_116

rm(data_116_date)


###################################################################### Data 117 ############################################################################
data_117 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2023/Finished Goods Inventory Health Adjusted Forward (IQR) - 05.17.23.xlsx",
                       sheet = "FG without BKO BKM TST")


data_117[-2, ] -> data_117
data_117 %>% 
  janitor::clean_names() %>% 
  dplyr::select(c(1, 3)) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x3 = as.integer(x3),
                x3 = as.Date(x3, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_117_date 


data_117[-1, ] -> data_117
colnames(data_117) <- data_117[1, ]
data_117[-1, ] -> data_117

data_117 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_117_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_117

rm(data_117_date)


###################################################################### Data 118 ############################################################################
data_118 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2023/Finished Goods Inventory Health Adjusted Forward (IQR) - 05.24.23.xlsx",
                       sheet = "FG without BKO BKM TST")


data_118[-2, ] -> data_118
data_118 %>% 
  janitor::clean_names() %>% 
  dplyr::select(c(1, 3)) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x3 = as.integer(x3),
                x3 = as.Date(x3, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_118_date 


data_118[-1, ] -> data_118
colnames(data_118) <- data_118[1, ]
data_118[-1, ] -> data_118

data_118 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_118_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_118

rm(data_118_date)


###################################################################### Data 119 ############################################################################
data_119 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/2023/Finished Goods Inventory Health Adjusted Forward (IQR) - 05.31.23.xlsx",
                       sheet = "FG without BKO BKM TST")


data_119[-2, ] -> data_119
data_119 %>% 
  janitor::clean_names() %>% 
  dplyr::select(c(1, 3)) %>% 
  dplyr::slice_head(n = 1) %>% 
  dplyr::mutate(x3 = as.integer(x3),
                x3 = as.Date(x3, origin = "1899-12-30")) %>% 
  dplyr::pull() -> data_119_date 


data_119[-1, ] -> data_119
colnames(data_119) <- data_119[1, ]
data_119[-1, ] -> data_119

data_119 %>% 
  janitor::clean_names() %>% 
  data.frame() %>% 
  dplyr::mutate(date = data_119_date) %>% 
  dplyr::mutate(year = lubridate::year(date),
                month = lubridate::month(date),
                day = lubridate::day(date)) -> data_119

rm(data_119_date)



