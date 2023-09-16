library(tidyverse)
library(readxl)

# Load file
Trs_sub <- function(Data,SheetName,Filter,Index,Saved,Month) {
  raw <- read_excel(Data,sheet = SheetName)
  # raw <- read_excel("AWPJun23.xlsx",sheet = "AWP1")

# Trs_sub("AWPJul23.xlsx", "AWPJul23_2", "Supports","[1],[4]", "test2", "Jul")

# Mark categories for filter
  Filter -> vars

  unlist(gregexpr('[[]', vars)) -> ind1
  unlist(gregexpr('[]]', vars)) -> ind2

  length(ind2) -> k

  for (i in 1:k) {
    substr(vars,ind1[i]+1,ind2[i]-1) -> ind3[i]
  }

  ind <- as.numeric(ind3)

# Assign cutpoint for outputs
  CutNum <-5
  # Month <-"Jun"
  Mth <- Month

# Get dimensions
  dim(raw) -> dimensions
  dimensions[2] -> ncols
  dimensions[1] -> nrows
  names(raw) <- paste0("v",1:ncols)
  which(str_detect(raw$v3,"Year")) -> startrow

# Get last row
  d1 <- raw[startrow:nrows,]
  names(d1) <- d1[1,]

# Get last column
  d2 <- d1[-1,]
  ncol(d2) -> lastcol
  lastcol-1 -> lcol

# Transform calculation variables as numeric
d2 %>% mutate_at(vars(1,14:lcol),as.numeric) -> d3

# Get maximum row
  max(d3[1], na.rm=TRUE) -> maximum
  which(d3[1]==maximum) -> maxrow

# rewrite data
  d3[1:maxrow,]-> d4

# Create index catalog
  d4 %>% select(ind) -> Var

  table(Var) -> tb
  data.frame(tb)-> tb2
  tb2$Index <- rownames(tb2)

  d4 %>% mutate(tv1=Supports) -> d4a

  merge(d4a, tb2, by ="tv1", all.x = TRUE) -> d4b

  write.csv(d4b, paste0(Saved,'.csv'))
}


  # https://stackoverflow.com/questions/27721008/how-do-i-deal-with-special-characters-like-in-my-regex

  d4 %>% summarise(across(14:lcol, ~ sum(.x, na.rm = TRUE))) -> sum

  #k table(d4$Indicator) -> tb
  #k write.csv(tb,"Indicator_tb.csv")

  #  write.csv(sum,paste0(OutputFile,".csv"))

  # get short names of Indicator
  #  table(d4$Indicator)-> tb
  #  data.frame(tb) -> df
  # df %>% mutate(xInc = substr(Var1,1,3))-> df2
  #  df2[3] -> AbInc

  d4 %>% mutate(AbInc= substr(Indicator,1,3)) -> d5

  # recode state region groups
  d5 %>%
    mutate(SRgp = case_when(
      str_detect(`State/ Region`,"Kachin") ~ "Kachin",
      str_detect(`State/ Region`,"Rakhine") ~ "Rakhine",
      str_detect(`State/ Region`,"Chin|Sagaing") ~ "Chin and Sagaing",
      str_detect(`State/ Region`,"Yangon|Ayeyarwady") ~ "Yangon and Ayeyarwady",
      str_detect(`State/ Region`,"Shan|Kayah") ~ "Shan (E+N+S) and Kayah",
      str_detect(`State/ Region`,"Kayin|Mon|Bago|Tanin") ~ "Kayin, Mon, Bago, Taninitharyi",
      str_detect(`State/ Region`,"Nay|Magw|Mandalay") ~ "Nay Pyi Taw, Magwe, Mandalay",
      TRUE ~ `State/ Region`)) -> d6

  # Split data sets into two groups: reference and indicators
  d6ref <- d6[c(1:13,29:31)]
  d6inc <- d6[14:28]

  # Organize indicators
  names(d6inc) <-c(paste0("O1",1:CutNum), paste0("O2",1:CutNum), paste0("O3",1:CutNum))
  cbind(d6ref[1:13], d6inc[1:15], d6ref[14:16])-> d7

  # Compute totals of indicators
  d7 %>% group_by(AbInc, SRgp) %>%
    summarise(
      O11 = sum(O11, na.rm=TRUE),
      O12 = sum(O12, na.rm=TRUE),
      O13 = sum(O13, na.rm=TRUE),
      O14 = sum(O14, na.rm=TRUE),
      O15 = sum(O15, na.rm=TRUE),

      O21 = sum(O21, na.rm=TRUE),
      O22 = sum(O22, na.rm=TRUE),
      O23 = sum(O23, na.rm=TRUE),
      O24 = sum(O24, na.rm=TRUE),
      O25 = sum(O25, na.rm=TRUE),

      O31 = sum(O31, na.rm=TRUE),
      O32 = sum(O32, na.rm=TRUE),
      O33 = sum(O33, na.rm=TRUE),
      O34 = sum(O34, na.rm=TRUE),
      O35 = sum(O35, na.rm=TRUE)) -> sumInc

  # Adjust with Excel templace: replace the missing row data
  data.frame(sumInc) -> r1
  table(r1$SRgp) -> tbSR
  data.frame(tbSR) -> SR2

  SR3 <- data.frame(SRgp = rep(SR2$Var1,5),
                    AbInc= c(rep("C-1",7), rep("C-2",7), rep("T-1",7), rep("M-1", 7), rep("S-1",7)),
                    Index = c(rep(1,7), rep(2,7), rep(3,7), rep(4,7), rep(5,7)))

  merge(r1,SR3, by = c("AbInc","SRgp"), all.y=TRUE) -> r2

  # Get Group totals of Each Indicator
  r2 %>% group_by(AbInc) %>%
    summarise(across(O11:O35, ~ sum(.x, na.rm = TRUE))) %>%
    bind_rows(r2, .) -> r3

  nrow(r3) -> Totrows
  Totrows-4 -> Start
  r3[Start:Totrows, 2] <-"ZZZ Total"
  r3$ftInc <- factor(r3$AbInc,levels = c("C-1","C-2","T-1", "M-1","S-1"))

  # ncol(r3) -> IncIndex
  # r3[Start:Totrows,IncIndex] <-9

  r3[order(c(as.numeric(r3$ftInc),r3$SRgp)), ] -> r4
  r4 %>% filter(!is.na(AbInc) | !is.na(SRgp)) -> r5

  r5 %>% mutate(SRgp = case_when(str_detect(SRgp,"ZZZ") ~ "Total",
                                 TRUE ~ SRgp)) %>% select(-c(Index)) -> r6
  write.csv(r6,"r6.csv")

  # Compute output totals
  r6 %>% mutate(All1 = O11 + O21 + O31,
                All2 = O12 + O22 + O32,
                All3 = O13 + O23 + O33,
                All4 = O14 + O24 + O34,
                All5 = O15 + O25 + O35) -> r7

  write.csv(r7,"r7.csv")

  # Compute target month stats
  colnames(d7)[4] = "Month"
  d7 %>% filter(Month== Mth ) %>% group_by(AbInc, SRgp) %>%
    summarise(
      O11 = sum(O11, na.rm=TRUE),
      O12 = sum(O12, na.rm=TRUE),
      O13 = sum(O13, na.rm=TRUE),
      O14 = sum(O14, na.rm=TRUE),
      O15 = sum(O15, na.rm=TRUE),

      O21 = sum(O21, na.rm=TRUE),
      O22 = sum(O22, na.rm=TRUE),
      O23 = sum(O23, na.rm=TRUE),
      O24 = sum(O24, na.rm=TRUE),
      O25 = sum(O25, na.rm=TRUE),

      O31 = sum(O31, na.rm=TRUE),
      O32 = sum(O32, na.rm=TRUE),
      O33 = sum(O33, na.rm=TRUE),
      O34 = sum(O34, na.rm=TRUE),
      O35 = sum(O35, na.rm=TRUE)) -> sumJun

  data.frame(sumJun) -> j1
  merge(j1,SR3, by = c("AbInc","SRgp"), all.y=TRUE) -> j2

  # Get Group totals of Each Indicator
  j2 %>% group_by(AbInc) %>%
    summarise(across(O11:O35, ~ sum(.x, na.rm = TRUE))) %>%
    bind_rows(j2, .) -> j3

  # nrow(r3) -> Totrows
  # Totrows-4 -> Start

  j3[Start:Totrows, 2] <-"ZZZ Total"
  j3$ftInc <- factor(j3$AbInc,levels = c("C-1","C-2","T-1", "M-1","S-1"))

  # ncol(r3) -> IncIndex
  # r3[Start:Totrows,IncIndex] <-9

  j3 %>% filter(!is.na(AbInc) | !is.na(SRgp)) -> j4


  # Compute target month outputs total

  j4 %>% mutate(Mth1 = O11 + O21 + O31,
                Mth2 = O12 + O22 + O32,
                Mth3 = O13 + O23 + O33,
                Mth4 = O14 + O24 + O34,
                Mth5 = O15 + O25 + O35) -> j5

  j5[order(c(as.numeric(j3$ftInc),j3$SRgp)), ] -> j6

  j6 %>% mutate(SRgp = case_when(str_detect(SRgp,"ZZZ") ~ "Total",
                                 TRUE ~ SRgp)) %>% select(-Index) -> j7

  # write.csv(j7,"j7.csv")

  ncol(j7) -> j7last
  j7last-4 -> j7start
  ft<-j7start-1

  j7[c(1,2,ft:j7last)] -> j8
  j8 %>% filter(!is.na(ftInc)| !is.na(SRgp)) %>%
    mutate(SRgp = case_when(str_detect(SRgp,"Total") ~ "ZZZ Total",
                            TRUE ~ SRgp)) -> j9
  r7  %>%
    mutate(SRgp = case_when(str_detect(SRgp,"Total") ~ "ZZZ Total",
                            TRUE ~ SRgp)) -> r8


  merge(r8,j9, by=c("ftInc","SRgp")) -> r9

  r9 %>% select(-contains("AbInc")) %>%
    mutate(SRgp = case_when(str_detect(SRgp,"ZZZ") ~ "Total",
                            TRUE ~ SRgp)) -> tb1

  tb1[order(c(tb1$ftInc,tb1$SRgp)),] -> tb2

  tb2 %>% filter(!is.na(ftInc)| !is.na(SRgp)) -> tb3

  tb3

  write.csv(tb3,paste0(Saved,".csv"), row.names=FALSE)
}

  # r2 %>%  select(1,2,3:7) %>% mutate(Output = "Output1") -> O1
  # r2 %>%  select(1,2,8:12) %>% mutate(Output = "Output2") -> O2
  # r2 %>%  select(1,2,13:17) %>% mutate(Output = "Output3") -> O3
