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


