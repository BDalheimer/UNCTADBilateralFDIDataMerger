## This code transforms individual country data on Bilateral Foreign Direct investment, 
## available at the UNCTAD website, into 4 long data frames. For each measured variable, inflow,
## outflow, instocks and outstocks, a seperate csv file is generated containing all cross-invesments
## for all available countries. The Country Code system is adoped from UNCTAD (ISO-3166).

# Set working Directory. TO DO: download data directly from server

setwd("~/FDI_Data")
# load packages
library(xlsx)
library(reshape2)
library(data.table)

## function "combine" to read .xls files and reformat them individually in nx4 data frames

combine <- function(x, y){
  if (file.exists(paste(paste("webdiaeia2014d3_", x, sep=""), "xls", sep="."))){
    
  dat <- read.xlsx(paste(paste("webdiaeia2014d3_", x, sep=""), "xls", sep="."), y, startRow=5)
  
  
  
  numb <- which(grepl("Table", dat[,1]))
  
  if(length(numb) > 0){
    todel <- NULL
  for (i in 1:length(numb)) {
    lower <- numb[i]-1
    upper <- numb[i] + 3
    todel <- c(todel,c(lower:upper))
    dat <- dat[-todel,]
  }}
  
 
  l <- which(grepl("Source", dat[,1]))
  if(length(l) > 0){
  l <- l -3
  dat <- dat[1:l,]
  }
  dat <- dat[rowSums(is.na(dat))!=ncol(dat), ]
  if(ncol(dat) <= 13) {
  
  dat <- dat[!is.na(dat[,1]),]
  dat <- setnames(dat, 1, "Reporting Economy")
  dat$area <- rep(paste(x), length(dat[1]))
  dat <- melt(dat, id.vars=c("area", "Reporting Economy"))
  dat <- setnames(dat, 3, "Year")
  dat$Year <- gsub("X", "", dat$Year)
  dat
  
  }else{
    
    ids <- dat[, !grepl( "X" , names( dat ))]
    years <- dat[, grepl( "X" , names( dat ))]
    nam <- colnames(ids)
    ids$ReportingEconomy <- do.call(paste, c(ids[nam], sep=""))
    ids$ReportingEconomy <- gsub("NA", "", ids$ReportingEconomy)
    ids$ReportingEconomy <- gsub("[[:space:]]", "", ids$ReportingEconomy)
    ids$area <- rep(paste(x), length(dat[1]))
    res <- cbind(ids$area, ids$ReportingEconomy, years)
    res <- setnames(res, c(1,2), c("area", "Reporting Economy"))
    res <- melt(res, id.vars=c("area", "Reporting Economy"))
    res <- setnames(res, 3, "Year")
    res$Year <- gsub("X", "", res$Year)
    res
  }
}
}

## Run function for each .xls workbook and combine all country data. Result is nx4 data frame

inflow <- NULL
for (i in cclist){
  inflow <- rbind(inflow, combine(i, 1))
}

outflow <- NULL
for (i in cclist){
  outflow <- rbind(outflow, combine(i, 2))
}

instock <- NULL
for (i in cclist){
  instock <- rbind(instock, combine(i, 3))
}

outstock <- NULL
for (i in cclist){
  outstock <- rbind(outstock, combine(i, 4))
}

# Write .csv files
write.csv(inflow, "inflow.csv", row.names=F)
write.csv(outflow, "outflow.csv", row.names=F)
write.csv(instock, "instock.csv", row.names=F)
write.csv(outstock, "outstock.csv", row.names=F)