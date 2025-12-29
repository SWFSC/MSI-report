# Summarize serious injury and mortality data
# !! Note this script produces several in-console tables and CSVs that should be checked by eye.
# Adapted from Jim Carretta's 02-25-2025 script
# Input file renamed "HCMSI_Records_SWFSC_Main.xlsx" in Sep 2025

# Load required packages
  library(ggplot2)
	library(flextable)
	library(officer)
	library(readxl)
	library(dplyr)
	library(magrittr)
	
# Set local paths
	## set local working directory (make this different from GitHub repo)
	setwd("C:/Users/alex.curtis/Data/MSI")
	## set local path to main data file
	path.dat <- "C://Users/alex.curtis/Data/Github/MSI-data/data/"
	
# Import data
	data = read_excel(paste0(path.dat, "HCMSI_Records_SWFSC_Main.xlsx"))
	
	## Set years to include in annual MSI report, subset data
	min.year = 2019
	max.year = 2023
	inc.yrs = min.year:max.year
	x <- data %>% filter(Year %in% inc.yrs)
	
	## get species group lists
	spgroups <- read.csv(paste0(path.dat, "lt_SpeciesGroups.csv"))
	pinn.spp <- spgroups %>% filter(SpeciesGroup=="pinniped") %>% select(Species) %>% unlist()
	sm.cet.spp <- spgroups %>% filter(SpeciesGroup=="small cetacean") %>% select(Species) %>% unlist()
	lg.whale.spp <- spgroups %>% filter(SpeciesGroup=="large whale") %>% select(Species) %>% unlist()
	rm(spgroups)
	
	## get interaction type classifications (commercial or other)
	hcmsi.sources <- read.csv(paste0(path.dat, "lt_IntrxnTypes.csv"))
	commercial <- hcmsi.sources %>% filter(SARTable=="commercial") %>% select(Interaction.Type) %>% unlist()
	other <- hcmsi.sources %>% filter(SARTable=="other") %>% select(Interaction.Type) %>% unlist()
	rm(hcmsi.sources)
	
# Function to tabulate MSI by source and species, format for Markdown
	Table_MSI <- function(df) {    
	  Interaction.Type <- df$Interaction.Type
	  MSI.Value <- df$MSI.Value
	  Species <- df$Species
	  
	  df.new <- cbind.data.frame(Interaction.Type, MSI.Value)
	  names(df.new) <- c("Source", "MSI.Value")
	  Cases <- as.numeric(table(df.new$Source))
	  MSI.Total <- tapply(df.new$MSI.Value, df.new$Source,  sum)
	  
	  df.new <- cbind.data.frame(Cases, MSI.Total)
	  df.new <- cbind.data.frame(row.names(df.new), Cases, MSI.Total)
	  df.new <- df.new[order(df.new$Cases, decreasing=TRUE),]
	  names(df.new) <- c("Source", "Cases", "MSI.Total")
	  
	  Table_MSI <- flextable(df.new)
	  Table_MSI <- line_spacing (Table_MSI, space = 0.3, part = "body")
	  Table_MSI <- fontsize(Table_MSI, size=8, part="all")
	  std_border = fp_border(color="gray")
	  hline(Table_MSI, part="all", border = std_border)
	  line_spacing(Table_MSI, space = 0.4, part = "all")
	  Table_MSI 
	  }
	
# How many records were associated with observer programs or monitored activities?
	systematic.1 = c(grep("DETERRENT|OBSERVER PROGRAM|CODEND", x$Comments))
	# the following may include electronic monitoring, observation of catch at processing facility 
	systematic.2 = c(grep("HAWAII LONGLINE|CA SWORDFISH|CA HALIBUT|PROCESSOR|CATCH SHARES|CATCH MONITOR|OPEN ACCESS|LIMITED ENTRY|ROCKFISH EM|CATCHER PROCESSOR|MACKEREL TRAWL|FLATFISH TRAWL|AT-SEA HAKE", x$Interaction.Type))
  systematic.3 = c(grep("RESEARCH", x$Interaction.Type))
  systematic.4 = c(grep("AUTHORIZED", x$Interaction.Type))
	systematic.obs = unique(c(systematic.1, systematic.2, systematic.3, systematic.4))
	observer.program = unique(c(systematic.1, systematic.2))
	opportunistic.length <- nrow(x)-length(systematic.obs)

# summarized initial condition NSI/SI records by interaction.type
	mortality = which(x$Final.Injury.Assessment=="DEAD")
  alive = x[-mortality,]
  dead = x[mortality,]
  rm(mortality)
	
  # SI determinations
  PINNIPEDS.DETERMINED = alive %>% filter(Species %in% pinn.spp)
	SMALL.CET.DETERMINED = alive %>% filter(Species %in% sm.cet.spp)
	WHALES.DETERMINED = alive %>% filter(Species %in% lg.whale.spp)
	## check initial vs final assessments (*** MOVE TO A SCRIPT FOR CHECKING MOST RECENT YEAR OR FIVE YEARS?***)
	table(PINNIPEDS.DETERMINED$Initial.Injury.Assessment, PINNIPEDS.DETERMINED$Final.Injury.Assessment)
	table(SMALL.CET.DETERMINED$Initial.Injury.Assessment, SMALL.CET.DETERMINED$Final.Injury.Assessment)
	table(WHALES.DETERMINED$Initial.Injury.Assessment, WHALES.DETERMINED$Final.Injury.Assessment)
	
# How many large whale records involved an attempt to free the whale from gear?
  rescue.mentioned = grep("FAST|WET|REMOVE|CUT|TEAM|RELEASED|DISTENTANGLEMENT TEAM|GEAR REMOVED|REMOVED GEAR|DISENTANGLED|SUCCESSFULLY", WHALES.DETERMINED$Comments)
  no.rescue = grep("NO RESIGHT|NO RESCUE|COULD NOT LOCATE|NO ATTEMPT|NO RESPONSE", WHALES.DETERMINED$Comments)
	
# Summarize NSI/SI/Mortality by Interaction.Type

# GENERAL ENTANGLEMENTS
	  all.entanglements = grep("ENTANGLE|TRAP|TRAWL|NET|GILLNET|SEA BOTTOM|WAVE RIDER|DEBRIS", x$Interaction.Type)
	  all.gillnet = grep("GILL", x$Interaction.Type, ignore.case=T)

# Pinniped output
	   pinn.records = x %>% filter(Species %in% pinn.spp)
     pinn.net.trawl = pinn.records[grep("net|trawl", pinn.records$Interaction.Type, ignore.case=TRUE),]
     pinn.harassment = pinn.records[grep("harass", pinn.records$Interaction.Type, ignore.case=TRUE),]

# How many pinnipeds were injured / killed in commercial gillnet, longline, trawl fisheries?
  pinn.commercial.records = length(which(pinn.records$Interaction.Type %in% commercial))
	
# How many pinnipeds were euthanized?
	euthanized = grep("EUTH", pinn.records$Comments)	

# examine pinniped injury categories
  table(pinn.records$Interaction.Type, pinn.records$Initial.Injury.Assessment, pinn.records$Final.Injury.Assessment)

# large whale output
	whale.records <- x %>% filter(Species %in% lg.whale.spp)
	whales.dead <- whale.records[whale.records$Final.Injury.Assessment=="DEAD",]
	# # write CSV to check by eye
	# write.csv(whale.records, "X-yr Large Whales.csv", fileEncoding="UTF-8")
	
# small cetaceans 
	small.cet.records <- x %>% filter(Species %in% sm.cet.spp)
	small.cet.causes = as.data.frame(table(small.cet.records$Interaction.Type))
	small.cet.causes = small.cet.causes[order(small.cet.causes[,2], decreasing=T, na.last=F), ]
	# different subset from list of 'systematic' observation created above
	systematic.small.cet.records <- length(grep("SCIENTIFIC RESEARCH", small.cet.records$Interaction.Type, ignore.case=TRUE)) +
	                                 length(grep("OBSERVER PROGRAM", small.cet.records$Comments, ignore.case=TRUE))

#  Rehab mentioned
     rehab = grep("REHAB", x$Comments)
     rehab = x[rehab,]
     table(rehab$Interaction.Type, rehab$SI.Criteria.Supporting.Assessment)

# extract 5-year records for each species
	## pinnipeds
	AT <- x[x$Species=="GUADALUPE FUR SEAL",]
	CSL <- x[x$Species=="CALIFORNIA SEA LION",]
	CU <- x[x$Species=="NORTHERN FUR SEAL",]	
	MA <- x[x$Species=="NORTHERN ELEPHANT SEAL",]
	OT <- x[x$Species=="UNIDENTIFIED OTARIID",]
  PU <- x[x$Species=="UNIDENTIFIED PINNIPED",]
	PV <- x[x$Species=="HARBOR SEAL",]
	USeal <- x[x$Species=="UNIDENTIFIED SEAL",]
	USL <- x[x$Species=="UNIDENTIFIED SEA LION",]
	UFS <- x[x$Species=="UNIDENTIFIED FUR SEAL",]
  # small cetaceans (excluding beaked and pilot whales)
	CD <- x[x$Species=="COMMON DOLPHIN UNIDENTIFIED",]
	PP <- x[x$Species=="HARBOR PORPOISE",]
	LO <- x[x$Species=="PACIFIC WHITE-SIDED DOLPHIN",]
	DS <- x[x$Species=="COMMON DOLPHIN SHORT-BEAKED",]
	DL <- x[x$Species=="COMMON DOLPHIN LONG-BEAKED",]
	SC <- x[x$Species=="STRIPED DOLPHIN",]
	UD <- x[x$Species=="UNIDENTIFIED DOLPHIN",]
	LB <- x[x$Species=="NORTHERN RIGHT WHALE DOLPHIN",]
	PD <- x[x$Species=="DALL'S PORPOISE",]
	# large whales
	BA <- x[x$Species=="MINKE WHALE",]
	BM <- x[x$Species=="BLUE WHALE",]
	BP <- x[x$Species=="FIN WHALE",]
	ER <- x[x$Species=="GRAY WHALE",]
	MN <- x[x$Species=="HUMPBACK WHALE",]
	PM <- x[x$Species=="SPERM WHALE",]
	UW <- x[x$Species=="UNIDENTIFIED WHALE",]
	
	
# Write data outputs
  ##	write temporary data file
  save.image("SeriousInjuryReport.RData")
  ## write 5-year csv output file with added codes for Github upload
  ## (move manually to MSI-report/data directory after checking)
  write.csv(x, paste(min.year, "_", max.year, "_5yr_MSI.csv", sep=""), row.names=F, fileEncoding="UTF-8")
  ## write all data through most-recently published report for Github upload 
  ## (move manually to MSI-report/data directory after checking)
  MSI.all.data <- data[data$Year<=max.year,]
  write.csv(MSI.all.data, paste("HCMSI_Records_SWFSC_", min(data$Year), "_",
                                "LatestPublishedYear",
                                ".csv", sep=""), row.names=F, fileEncoding="UTF-8")


#* Then run AppendixTable.R *# (if including in MSI report)
	