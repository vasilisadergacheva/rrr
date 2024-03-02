library(tidyverse)
library(readxl)
library(openxlsx)
library(XLConnect)
library(xlsx)


path = setwd('/home/vasilisa/work/r/')
path <- paste(path, '/data/1_Descriptives_26_02_2024.xlsx', sep = '')

df <- read_excel(path)
wb <- loadWorkbook(path)

sheetname <- "Sheet1"
sheets <- getSheets(wb)   
sheet <- sheets[[sheetname]] 
nrows = nrow(df)
ncols = ncol(df)
rows <- getRows(sheet, rowIndex=2:(nrows+1))
cells <- getCells(rows, colIndex = 2:(ncols+1)) 
values <- lapply(cells, getCellValue)

bold <- "_"
yellow <- "_"
for (row in 2: nrows + 1) {
  ind = as.character(paste(row, 5, sep="."))
  sfc <- values[[ind]]
  if (!is.null(sfc)) {
    bold <- append(bold, paste(row, 2, sep="."))
    
    sfc <- gsub("[^0-9.-]", "", sfc)
    if (sfc > 0.3) {
      yellow <- append(yellow, ind)
    }
  }
}
bold <- bold[-1]
yellow <- yellow[-1]

lapply(names(cells[bold]),
       function(ii)setCellStyle(cells[[ii]], 
                                CellStyle(wb) + Font(wb, isBold=TRUE)))

lapply(names(cells[yellow]),
       function(ii)setCellStyle(cells[[ii]], 
                                CellStyle(wb, fill=Fill(foregroundColor="yellow"))))

saveWorkbook(wb, path)
