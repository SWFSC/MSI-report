# Create Word doc Appendix of HCMSI cases

rm(list=ls())

library(flextable)
library(officer)
library(dplyr)

load("SeriousInjuryReport.RData")

# data frame 'x' should be a UTF-8 comma-delimited file to work with flextable

include.cols <- names(x)[-1]
shortNames <- c("Species", "Stock", "YYYY", "MM", "DD", "Source", "County", "State", "Initial", "Narrative", "Final", "SI.Criteria", "MSI.Value", "LOF", "PBR")
x <- x[-1]
names(x) <- shortNames

set_flextable_defaults(na_str = "NA", nan_str = "NaN")

A1 <- flextable(x)
A1 <- fontsize(A1, size=7, part="all")
A1 <- align(A1, align="center", part="all")
A1 <- font(A1, part="all", fontname="Arial Narrow")
A1 <- set_table_properties(A1, layout = "autofit")
A1 <- colformat_num(A1, big.mark="")
std_border = fp_border(color="gray")
A1 <- hline(A1, part="all", border = std_border)

doc <- read_docx()
doc <- body_add_flextable(doc, value = A1, align = "left")
print(doc, target = "Appendix_1.docx")

# Once in Word, Layout/autofit/autofit to contents

# Add Table header from placeholder page in HCM_SI_Report.docx
# Note carriage returns in Narrative can mess up Word table formatting

# # Additional table that includes only most-recent complete year of data for SRG review
# 
# latest.year <- latest.yr[-1]
# names(latest.year) <- shortNames
# A2 <- flextable(latest.year)
# A2 <- fontsize(A2, size=7, part="all")
# A2 <- align(A2, align="center", part="all")
# A2 <- font(A2, part="all", fontname="Arial Narrow")
# A2 <- set_table_properties(A2, layout = "autofit")
# A2 <- colformat_num(A2, big.mark="")
# std_border = fp_border(color="gray")
# A2 <- hline(A2, part="all", border = std_border)
# 
# doc <- officer::read_docx()
# doc <- body_add_flextable(doc, value = A2, align = "left")
# print(doc, target = "Appendix_2.docx")
# 
