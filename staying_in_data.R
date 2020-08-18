# windows
setwd("C:/Users/User/Desktop/patients")

# linux
setwd("~/patients")

library(docxtractr)
library(readxl)
library(qdapTools)
library(stringr)
library(readr)
library(lubridate)
library(dplyr)

# here I scan all .docx files in the directory and read them
temp <- list.files(pattern="*.docx")
patients <- data.frame("шифр" = 1:length(temp))
patients_in <- sapply(temp, read_docx)

################## staying in #####################

ALL_DATA <- data.frame(шифр = NA, дата = NA, номер_суток = NA, время = NA, тип = NA, общее_состояние = NA, сознание = NA, ЧДД = NA, Spo2 = NA, Сист_АД = NA, Диаст_АД = NA, ЧСС = NA, NEWS = NA)

# for j iterates through patients
# for i iterates through the different check results of each patient

# this loop is really long ¯\_(ツ)_/¯ 
# scroll 40 lines to reach its end

for(j in 1:length(temp)) {
  current_patient <- patients_in[[j]]

key_words <- paste0(c("ПЕРВИЧНЫЙ ОСМОТР", "ОБХОД В РЕАНИМАЦИИ", "ОСМОТР В РЕАНИМАЦИИ", 
                        "ДНЕВНИК", "ОСМОТР ДЕЖУРНОГО ВРАЧА", "СОВМЕСТНЫЙ", "ОБХОД В РЕАНИМАЦИИ",
                        "ОСМОТР ЗАВЕДУЮЩЕГО ОТДЕЛЕНИЕМ"),collapse = '|')
  
# Some doctors write more than 2 keywords for each check
# Here I delete the earlier keyword, if there is another one after it
indexes <- which(str_detect(current_patient, key_words))
to_delete <- c(10000)
for(n in 1:(length(indexes)-1)){
    if(indexes[n] + 1 == indexes[n+1]){
      to_delete <- c(to_delete, n+1)}
  }
indexes <- indexes[-to_delete]

#

# temporary - data frame for each check, that is attached to the main dataset via bind_rows after each iteration
temporary <- data.frame(шифр = 1:(length(indexes)-1), дата = NA, номер_суток = NA, время = NA, тип = NA, общее_состояние = NA, сознание = NA, ЧДД = NA, Spo2 = NA, Сист_АД = NA, Диаст_АД = NA, ЧСС = NA, NEWS = NA)

# iterate through each patient's checks
for(i in 1:(length(indexes)-1)){
    current_check <- current_patient[indexes[i]:indexes[i+1]]
    
    temporary$шифр <- str_sub(temp[j], start = 1, end = -6)
    temporary$дата[i] <- current_check[1] %>% str_sub(start = 1, end = 10)
    temporary$время[i] <- current_check[1] %>% str_sub(start = 12, end = 16)
    temporary$тип[i] <- current_check[1] %>% str_sub(start = 18, end = -1)
    temporary$номер_суток[i] <- as.integer(dmy(temporary$дата[i]) - dmy(patients_days$first[j]))
    temporary$ЧСС[i] <- str_sub(str_extract(current_check[which(str_detect(current_check, "ЧСС"))], "ЧСС:...."), start = 6, end = 8)[1]
    temporary$температура[i] <- str_extract(current_check[which(str_detect(current_check, "3[:digit:],[:digit:]"))][1], "3[:digit:],[:digit:]")
    temporary$общее_состояние[i] <- str_sub(current_check[which(str_detect(current_check, "Общее состояние"))[1]], start = 18, end = 32)
    temporary$сознание[i] <- str_sub(str_extract(current_check[which(str_detect(current_check, "Общее состояние"))[1]], "Сознание.................."), start = 10, end = 15)
    temporary$ЧДД[i] <- str_sub(str_extract(current_check[which(str_detect(current_check, "ЧДД"))[1]], "ЧДД: ......."), start = 5, end = 8)
    temporary$Spo2[i] <- str_sub(str_extract(current_check[which(str_detect(current_check, "SPO2"))[1]], "SPO2:....."), start = 6, end = -3)
    temporary$Сист_АД[i] <- str_sub(str_extract(current_check[which(str_detect(current_check, "Систолическое"))[1]], "Систолическое\\sдавление..............."),  start = 24, end = 27)
    temporary$Диаст_АД[i] <- str_sub(str_extract((current_check[which(str_detect(current_check, "Диастолическое"))[1]]), "Диастолическое\\sдавление:.............."), start = 26, end = 28)
    temporary$NEWS[i] <- str_sub(str_extract(current_check[which(str_detect(current_check, "NEWS"))[1]], "NEWS....."), start = 7, end = 8)
    
  }
ALL_DATA <- bind_rows(ALL_DATA, temporary) 
}

ALL_DATA <- ALL_DATA[-1,]

# saving to excel file to share
write.xlsx(patients, "поступление_6.xlsx")
write.xlsx(patients_days_for_print, "количество_дней_6.xlsx")
write.xlsx(ALL_DATA, "стационар6.xlsx")

