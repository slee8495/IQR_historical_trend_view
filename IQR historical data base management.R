library(tidyverse)
library(magrittr)
library(openxlsx)
library(readxl)
library(writexl)
library(reshape2)
library(skimr)
library(janitor)
library(lubridate)


############ Data collection ###############

###################################################################### Data 1 ############################################################################
data_1 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/Finished Goods Inventory Health (IQR) - 01.13.21.xlsx",
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
data_2 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/Finished Goods Inventory Health (IQR) - 01.20.21.xlsx",
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
data_3 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/Finished Goods Inventory Health (IQR) - 01.27.21.xlsx",
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
data_4 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/Finished Goods Inventory Health (IQR) - 02.10.21.xlsx",
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
data_5 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/Finished Goods Inventory Health (IQR) - 02.17.21.xlsx",
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
data_6 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/Finished Goods Inventory Health (IQR) - 02.24.21.xlsx",
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
data_7 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/Finished Goods Inventory Health (IQR) - 03.10.21.xlsx",
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


###################################################################### Data 7 ############################################################################
data_8 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/Finished Goods Inventory Health (IQR) - 03.18.21.xlsx",
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
data_9 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/Finished Goods Inventory Health (IQR) - 03.24.21.xlsx",
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
data_10 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/Finished Goods Inventory Health (IQR) - 04.07.21.xlsx",
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
data_11 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/Finished Goods Inventory Health (IQR) - 04.14.21.xlsx",
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
data_12 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/Finished Goods Inventory Health (IQR) - 04.21.21.xlsx",
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
data_13 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/Finished Goods Inventory Health (IQR) - 05.19.21.xlsx",
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
data_14 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/Finished Goods Inventory Health (IQR) - 05.26.21.xlsx",
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
data_15 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/Finished Goods Inventory Health (IQR) - 06.09.21.xlsx",
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
data_16 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/Finished Goods Inventory Health (IQR) - 06.16.21.xlsx",
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
data_17 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/Finished Goods Inventory Health (IQR) - 06.23.21.xlsx",
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
data_18 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/Finished Goods Inventory Health (IQR) - 06.30.21.xlsx",
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
data_19 <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/Inventory Days On Hand/IQR Historical Data Collection/FG/Finished Goods Inventory Health (IQR) - 07.07.21.xlsx",
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



