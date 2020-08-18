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

################## checking in #####################


# creating dataset with detection of strings

for(i in 1:length(temp)){
  
  patients$шифр[i] <- str_sub(temp[i], start = 1, end = -6)
  
  patients$рост[i] <- str_extract(patients_in[[i]][which(str_detect(patients_in[[i]], "Рост"))[1]], "[:digit:][:digit:][:digit:]")
  
  patients$вес[i] <- str_extract(patients_in[[i]][which(str_detect(patients_in[[i]], "Вес"))[1]], "[:digit:][:digit:]\\s")
  
  patients$сахарный_диабет[i] <- str_sub(patients_in[[i]][which(str_detect(patients_in[[i]], "Сахарный"))[1]], start = 17, end = -1)
  
  patients$Систолическое_давление[i] <- str_sub(str_extract(patients_in[[i]][which(str_detect(patients_in[[i]], "Систолическое"))[1]], "Систолическое\\sдавление..............."),  start = 24, end = 27)
  
  patients$Диастолическое_давление[i] <- str_sub(str_extract((patients_in[[i]][which(str_detect(patients_in[[i]], "Диастолическое"))[1]]), "Диастолическое\\sдавление:.............."), start = 26, end = 28)
  
  patients$ЧСС[i] <- str_sub(str_extract((patients_in[[i]][which(str_detect(patients_in[[i]], "ЧСС"))[1]]), "ЧСС:\\s........"), start = 6, end = 8)
  
  patients$общее_состояние[i] <- str_sub(patients_in[[i]][which(str_detect(patients_in[[i]], "Общее состояние"))[1]], start = 18, end = 35)
  
  patients$сознание[i] <- str_sub(str_extract(patients_in[[i]][which(str_detect(patients_in[[i]], "Общее состояние"))[1]], "Сознание.................."), start = 10, end = 16)
  
  patients$ЧДД[i] <- str_sub(str_extract(patients_in[[i]][which(str_detect(patients_in[[i]], "ЧДД"))[1]], "ЧДД: ......."), start = 5, end = 8)
  
  patients$spo2[i] <- str_sub(str_extract(patients_in[[i]][which(str_detect(patients_in[[i]], "SPO2"))[1]], "SPO2:....."), start = 6, end = -3)
  
  patients$NEWS[i] <- str_extract(patients_in[[i]][which(str_detect(patients_in[[i]], "NEWS"))[1]], "NEWS.....")
  
  patients$дата_рождения[i] <- str_sub(patients_in[[i]][which(str_detect(patients_in[[i]], "Дата рождения"))[1]], start = 16, end = 25) 
  
  patients$возраст[i] <- str_sub(patients_in[[i]][which(str_detect(patients_in[[i]], "Дата рождения"))[1]], start = 27, end = 29) 
  
  patients$дата_поступления[i] <- str_sub(patients_in[[i]][which(str_detect(patients_in[[i]], "Дата поступления"))[1]], start = 31, end = -8) 
  
  patients$время_поступления[i] <- str_sub(patients_in[[i]][which(str_detect(patients_in[[i]], "Дата поступления"))[1]], start = 41, end = -3) 
  
}

################## number of days #####################

# here i create a dataset with some date related info - for temporary use

regex_date <- paste0("^[:digit:][:digit:].[:digit:][:digit:].[:digit:][:digit:][:digit:][:digit:]")

patients_days <- data.frame(шифр = NA, first = 1:length(patients_in), last = NA, n_days = NA)

for(n in 1:length(patients_in)){
  
  p <- patients_in[[n]]
  
  dates_index <- which(str_detect(patients_in[[n]], regex_date))
  
  patients_days$шифр[n] <- str_sub(temp[n], start = 1, end = 13)
  patients_days$first[n] <- str_sub(p[dates_index[1]], start = 1, end = 10)
  patients_days$last[n] <- str_sub(p[dates_index][length(dates_index)], start = 1, end = 10)
  patients_days$n_days[n] <- as.integer(dmy(str_sub(p[dates_index][length(dates_index)], start = 1, end = 10)) - dmy(str_sub(p[dates_index[1]], start = 1, end = 10)))
}


patients_days_for_print <- patients_days
colnames(patients_days_for_print) <- c("шифр","первый_день", "посл_день","всего_дней")

