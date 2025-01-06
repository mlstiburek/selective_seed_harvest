#This code returns correlated vector of male gametic contributions M
#for the input vector of breeding values
#change the desired correlation (r) in the function
library(readxl) 
library(openxlsx)
library(ggplot2)
library(dplyr)
library(tidyr)
library(faux)
library(clipr)

#BV <- scan()
data <- read_excel("G:/My Drive/SCA_optimization/Input_data/Scen_9/resampled_GCA_scen9.xlsx",range="Sheet1!A1:AD32",col_names=FALSE)
BV <- as.matrix(data)

MO <- matrix(1:960, nrow = 32, ncol = 30)
MO_shifted <- MO
M <- MO

for (i in 1:30) {
  #generate correlated values to GCA
  MO[,i] <- rnorm_pre(BV[,i], r = 0.5, empirical = TRUE)
  #convert to fractions
  MO_shifted[,i] <- MO[,i] - min(MO[,i])
  M[,i] <- MO_shifted[,i] / sum(MO_shifted[,i]) 
}

# Create a workbook
wb <- createWorkbook()

# Add a worksheet
addWorksheet(wb, sheetName = "Matrix Data")

# Write the matrix to the sheet
writeData(wb, sheet = "Matrix Data", x = M)

# Save the workbook
write.xlsx(as.data.frame(M), "G:/My Drive/SCA_optimization/Input_data/Scen_9/matrix_data.xlsx")
