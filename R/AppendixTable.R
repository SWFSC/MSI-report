# Create Word doc Appendix of HCMSI cases

library(flextable)
library(officer)

load("SeriousInjuryReport.RData")

# data frame 'x' should be a UTF-8 comma-delimited file to work with flextable

#include.cols <- names(x)[-1]
include.cols <- names(x)[c(-1,-2,-3,-c(19:24))] # JEM: remove columns that Alex or I added to HCMSI_Records_SWFSC_Main.xlsx but that don't align with shortNames in next row
shortNames <- c("Species", "Stock", "YYYY", "MM", "DD", "Source", "County", "State", "Initial", "Narrative", "Final", "SI.Criteria", "MSI.Value", "LOF", "PBR")
#x <- x[-1]
x <- x[c(-1,-2,-3,-c(19:24))] # JEM: remove columns in x that don't align with those specified by shortNames
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

# Once in Word:
 # - Layout/autofit/autofit to contents (see "MSI report steps.txt" in GitHub repo "MSI-report")
 # - Add Table header from placeholder page in HCM_SI_Report.docx
# Note carriage returns in Narrative can mess up Word table formatting