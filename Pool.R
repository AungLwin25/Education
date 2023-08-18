library(readr)
library(tidyverse)

# read two data sets
Pool <- function(Table1,Table2,Saved) {

  i1 <- read_csv(Table1)
  i2 <- read_csv(Table2)

  i1 <- read_csv("Indicators_Results.csv")
  i2 <- read_csv("TrainSupport_Results.csv")

 full_join(i1,i2, by = c("ftInc","SRgp")) -> db

 # Identify Training and support rows
 ## Indicators set
 which(i1$ftInc=="T-1") -> TRows
 TStart <- TRows[1]
 TLength <- length(TRows)
 TLast <- TStart + TLength -1

 ## Training and support set
 which(i2$ftInc=="T-1") -> SRows
 SStart <- SRows[1]
 SLength <- length(SRows)
 SLast <- SStart + SLength -1

 ## Remove training support rows from both
 i1[-c(TStart:TLast),] -> f1
 i2[ c(SStart:SLast),] -> f2

 # bind rows
 rbind(f1,f2) -> f3

 #  adjust format
 f3$ftInc <- factor(f3$ftInc,levels = c("C-1","C-2","T-1", "M-1","S-1"))
 f3[order(c(f3$ftInc,f3$SRgp)), ] -> f4
 f4 %>% filter(!is.na(ftInc)|!is.na(SRgp)) -> f5

 write.csv(f5,"Edu.csv", row.names=FALSE)
}


