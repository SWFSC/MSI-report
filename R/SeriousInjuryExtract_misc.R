# Contains code to produce additional tables, CSVs, figures, etc. from MSI data
# !! Run SeriousInjuryExtract.R first

# Write Large Whale file for Serious Injury Working Group
## Include following species
lw.data = data %>% filter(Species %in% lg.whale.spp)
write.csv(lw.data, "Large Whales US West Coast Serious Injury Working Group.csv", row.names=F, fileEncoding="UTF-8")

# write large whale output file for proration analysis
lw.prorate.sp = c("SPERM WHALE", "GRAY WHALE", "BLUE WHALE", "HUMPBACK WHALE", "UNIDENTIFIED WHALE", "MINKE WHALE", "FIN WHALE", "SEI WHALE")
SI.whales <- data %>% filter(Species %in% lw.prorate.sp)
write.csv(SI.whales, "Large Whale SI data.csv", fileEncoding="UTF-8")

# write CSV for Northern fur seal
write.csv(CU, "Northern.Fur.Seal.csv")

# Harbor seal stock assignments by county in WA state (is this chunk a reference for editing XLSX by hand?)
# c("WHATCOM", "SKAGIT", "SNOHOMISH", "ISLAND", "SAN JUAN", "KING") = WA N INLAND WATERS
# c("THURSTON") = S PUGET SOUND
# c("PIERCE") = S PUGET SOUND *or* WA N INLAND WATERS
# c("KITSAP")  = WA N INLAND WATERS *or* HOOD CANAL *or* S PUGET SOUND
# c("JEFFERSON") = WA N INLAND WATERS *or* HOOD CANAL *or* OR/WA COAST
# c("CLALLAM") = WA N INLAND WATERS *or* OR/WA COAST
# c("MASON") = HOOD CANAL *or* S PUGET SOUND

# harbor seal stocks
PV.Hood.Canal <- PV[PV$Stock.or.Area%in%c("HOOD CANAL"),]
PV.SPuget.Sound <- PV[PV$Stock.or.Area%in%c("SOUTHERN PUGET SOUND"),]
PV.WA.N.Inland <- PV[PV$Stock.or.Area%in%c("WASHINGTON NORTHERN INLAND WATERS"),]
sort(tapply(PV.Hood.Canal$MSI.Value, PV.Hood.Canal$Interaction.Type, sum),decreasing=TRUE)
sort(tapply(PV.SPuget.Sound$MSI.Value, PV.SPuget.Sound$Interaction.Type, sum), decreasing=TRUE)
sort(tapply(PV.WA.N.Inland$MSI.Value, PV.WA.N.Inland$Interaction.Type, sum), decreasing=TRUE)

# write CSV for humpback whale
write.csv(MN, "Humpback Summary.csv", fileEncoding="UTF-8")
# MN.pot.trap = MN[grep("POT|TRAP", MN$Interaction.Type),]
# MN.vessel = MN[grep("VESSEL STRIKE", MN$Interaction.Type),]
# MN.other = MN[grep("DEBRIS|GILLNET|UNIDENTIFIED FISHERY INTERACTION|HOOK", MN$Interaction.Type),]

# X-year summary of humpback mortality / injury cases by source
MN.Xyr <- cbind.data.frame(table(MN$Year, MN$Interaction.Type))
names(MN.Xyr) <- c("Year", "Source", "Number")

# Humpback plot by source for all years
MN.all.data <- data[data$Species=="HUMPBACK WHALE",]
MN.all.data <- cbind.data.frame(table(MN.all.data$Year, MN.all.data$Interaction.Type))
names(MN.all.data) <- c("Year", "Source", "Number")
MN.all.data$Source <- as.character(MN.all.data$Source)
## group MSI sources
MN.all.data$Source[grep("POT|TRAP|CRAB|LOBSTER|PRAWN", MN.all.data$Source, ignore.case=TRUE)] <- "POT-TRAP FISHERY"
MN.all.data$Source[grep("VESSEL STRIKE", MN.all.data$Source, ignore.case=TRUE)] <- "VESSEL STRIKE"
MN.all.data$Source[grep("NET", MN.all.data$Source, ignore.case=TRUE)] <- "NET FISHERY"
MN.all.data$Source[grep("HOOK", MN.all.data$Source, ignore.case=TRUE)] <- "HOOK AND LINE FISHERY"
MN.all.data$Source[grep("DEBRIS", MN.all.data$Source, ignore.case=TRUE)] <- "MARINE DEBRIS"
MN.all.data$Source[grep("UNIDENTIFIED FISHERY", MN.all.data$Source, ignore.case=TRUE)] <- "UNIDENTIFIED FISHERY"
MN.all.data$Source[grep("RESEARCH", MN.all.data$Source, ignore.case=TRUE)] <- "RESEARCH-RELATED"
## generate plot
MN.barplot.sources <- ggplot(MN.all.data, aes(x=Year, y=Number, fill=Source)) +
  geom_bar(stat="identity") +
  scale_fill_manual("Legend", values = c("POT-TRAP FISHERY" = "#EEBAB4", "VESSEL STRIKE" = "#A856CC", 
                                         "NET FISHERY" = "#E57A77", "HOOK AND LINE FISHERY" = 
                                           "orange", "MARINE DEBRIS" = "yellow4", "UNIDENTIFIED FISHERY" = "#3D65A5",
                                         "RESEARCH-RELATED" = "black")) +
  ggtitle("Humpback Whale: Human-Caused Mortality and Injury by Source: CA-OR-WA") +
  ylab("Number of Cases") +
  theme_classic()
MN.barplot.sources
## save plot
ggsave(plot=MN.barplot.sources, filename="Humpback.Sources.All.Years.tiff",
       width=300, height = 170, units="mm", dpi = 1200, compression = "lzw")

# call script to generate parallel blue whale plot for Take Reduction Team
# source("c:/carretta/github/HCM_SI/R/Blue_Whale_Plot.R")

# write CSV for blue whales
write.csv(BM, "Blue X-yr Summary.csv", fileEncoding="UTF-8")
# table(BM$Interaction.Type, BM$MSI.Value, BM$Final.Injury.Assessment)

# write CSV for gray whales
write.csv(ER, "Gray X-yr Summary.csv", fileEncoding="UTF-8")
# table(ER$Interaction.Type, ER$MSI.Value, ER$Final.Injury.Assessment)

#	table(UW$Interaction.Type, UW$MSI.Value)


### Summarize and sum anthropogenic cases by category
ER.summary = sort(tapply(ER$MSI.Value, ER$Interaction.Type, sum), decreasing=T)

non.comm.fish = c("VESSEL", "DEBRIS", "HUNT", "TRIBAL")
non.comm = grep("VESS|DEBRIS|HUNT|TRIBAL", ER$Interaction.Type)
ER.comm.fish = ER[-non.comm,]
ER.pot.trap = ER[grep("POT|TRAP", ER$Interaction.Type),]
ER.vessel = ER[grep("VESSEL", ER$Interaction.Type),]
ER.other = ER[grep("HUNT|NET|UNIDENTIFIED FISHERY|DEBRIS", ER$Interaction.Type),]

tapply(ER.comm.fish$MSI.Value, ER.comm.fish$Final.Injury.Assessment, sum)

humpback.summary = sort(tapply(MN$MSI.Value, MN$Interaction.Type, sum), decreasing=T)

blue.summary = sort(tapply(BM$MSI.Value, BM$Interaction.Type, sum), decreasing=T)

fin = x[x$Species=="FIN WHALE",]
fin.summary = sort(tapply(fin$MSI.Value, fin$Interaction.Type, sum), decreasing=T)
#	table(fin$Interaction.Type, fin$MSI.Value, fin$Final.Injury.Assessment)

unid = x[x$Species=="UNIDENTIFIED WHALE",]
unid.summary = sort(tapply(unid$MSI.Value, unid$Interaction.Type, sum), decreasing=T)
#	table(unid$Interaction.Type, unid$MSI.Value, unid$Interaction.Type)


# write most-recent year data to CSV by taxonomic group for cross-center review
# write most-recent year large whale data to file
whales.review.yr <- whale.records[whale.records$Year==max.year,]
write.csv(whales.review.yr, "Whales_Cross_Center_Review.csv", fileEncoding="UTF-8")
# write most-recent year pinniped data to file
pinnipeds.review.yr <- pinn.records[pinn.records$Year==max.year,]
write.csv(pinnipeds.review.yr, "Pinnipeds_Cross_Center_Review.csv", fileEncoding="UTF-8")
# write most-recent year small cetaceans to file
small.cet.review.yr <- small.cet.records[small.cet.records$Year==max.year,]
write.csv(small.cet.review.yr, "Small_Cetaceans_Cross_Center_Review.csv", fileEncoding="UTF-8")

# write most-recent year data for all species to CSV file (for SRG review)
latest.yr <- x[x$Year==max.year,]
write.csv(latest.yr, "Latest_Year_HCMSI.csv", fileEncoding="UTF-8")


# Table that includes only most-recent complete year of data for SRG review
latest.year <- latest.yr[-1]
names(latest.year) <- shortNames
A2 <- flextable(latest.year)
A2 <- fontsize(A2, size=7, part="all")
A2 <- align(A2, align="center", part="all")
A2 <- font(A2, part="all", fontname="Arial Narrow")
A2 <- set_table_properties(A2, layout = "autofit")
A2 <- colformat_num(A2, big.mark="")
std_border = fp_border(color="gray")
A2 <- hline(A2, part="all", border = std_border)

doc <- officer::read_docx()
doc <- body_add_flextable(doc, value = A2, align = "left")
print(doc, target = "Appendix_2.docx")

# source randomForest model used to prorate unid. whale entanglement cases to species (there is no equivalent model for vessel strikes)
# these fractional cases are added to stock assessment reports (SARs) as species-specific fractional estimates, but the unid whale ID is retained for annual serious injury reports

#	source("c:/Carretta/Github/Prorate_Unid_Whale_Entanglements/Prorate_Unid_Whale_Entanglements.R")


