library(tidyverse)
library(readxl)

Ben<- function(name,saved) {
raw <- read_excel(name,sheet = "Sheet1")

dim(raw) -> dimensions
dimensions[2] -> ncols
dimensions[1] -> nrows
names(raw) <- paste0("v",1:ncols)
which(str_detect(raw$v3,"Year")) -> startrow

d1 <- raw[startrow:nrows,]
names(d1) <- d1[1,]

d2 <- d1[-1,]
d2 %>% mutate_at(vars(1,13:27),as.numeric) -> d3

max(d3[1], na.rm=TRUE) -> maximum
which(d3[1]==maximum) -> maxrow

d3[1:maxrow,]-> d4
d4 %>% summarise(across(13:27, ~ sum(.x, na.rm = TRUE))) -> sums
write.csv(sums,paste0(saved,".csv"))
}
