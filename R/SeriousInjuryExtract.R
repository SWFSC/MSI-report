# Summarize serious injury and mortality data
# Note this script produces several in-console tables and CSVs that should be checked by eye.
# Jim Carretta 02-25-2025
# Updated to fully categorize all systematic (observer program, research-related, authorized removal) records as such
# Input file renamed "HCMSI_Records_SWFSC_Main.xlsx" in Sep 2025
  
	library(ggplot2)
	library(flextable)
	library(officer)
	library(readxl)
	
	setwd("C:/Users/alex.curtis/Data/MSI")
	data = read_excel("../Github/MSI-data/data/HCMSI_Records_SWFSC_Main.xlsx")
	data = data[,-1]   # Drop SWFSC Record ID column
	
#  Years to include in annual MSI report
	min.year = 2019
	max.year = 2023
	inc.yrs = seq(min.year, max.year, 1)
	
# Function to tabulate MSI by source and species
	
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
	  Table_MSI }
  
# check for missing data from critical fields
	
	missing.data <- function(df) {
	  
	  a <- which(is.na(df$Initial.Injury.Assessment==TRUE))
	  b <- which(is.na(df$Final.Injury.Assessment==TRUE))
	  c <- which(is.na(df$MSI.Value==TRUE))
	  d <- which(is.na(df$COUNT.AGAINST.LOF==TRUE))
	  e <- which(is.na(df$COUNT.AGAINST.PBR==TRUE))
	  f <- which(is.na(df$Year==TRUE))
	  
	  missing.list <- rbind.data.frame(a,b,c,d,e,f)
	  
	  if(nrow(missing.list)>=1) { warning("missing data in lines: ", missing.list)}
	  
	  if(nrow(missing.list)<1) { message("no missing data") }
	}
	
	missing.data(data)
	delay <- Sys.sleep(10)
	invisible(delay)
	
# subset targeted years of data   
  x = subset(data, data$Year>=min.year & data$Year<=max.year)
  
# How many records were associated with observer programs or monitored activities?
	systematic.1 = c(grep("DETERRENT|OBSERVER PROGRAM|CODEND", x$Comments))
	systematic.2 = c(grep("HAWAII LONGLINE|CA SWORDFISH|CA HALIBUT|PROCESSOR|CATCH SHARES|CATCH MONITOR", x$Interaction.Type))
  systematic.3 = c(grep("RESEARCH", x$Interaction.Type))
  systematic.4 = c(grep("AUTHORIZED", x$Interaction.Type))
	systematic.obs = unique(c(systematic.1, systematic.2, systematic.3, systematic.4))
	observer.program = unique(c(systematic.1, systematic.2))
	opportunistic.length <- nrow(x)-length(systematic.obs)

#  summarized initial condition NSI/SI records by interaction.type
	mortality = which(x$Final.Injury.Assessment=="DEAD")
  alive = x[-mortality,]
  dead = x[mortality,]
	
# correct MSI.Value field to correct for mis-typed entries
	x$MSI.Value[mortality]=1
	final.dead.captive=which(x$Final.Injury.Assessment%in%c("DEAD","CAPTIVITY"))
	x$MSI.Value[final.dead.captive]=1
	
	NSI.final = which(x$Final.Injury.Assessment=="NSI")
	x$MSI.Value[NSI.final]=0
  
	pinn.live = grep("SEA|PINN|OTAR", alive$Species)
	small.cet.live = grep("DOL|PORP|KILLER", alive$Species)
	
	PINNIPEDS.DETERMINED = alive[pinn.live,]
	WHALES.DETERMINED = alive[-c(pinn.live,small.cet.live),]
	SMALL.CET.DETERMINED = alive[small.cet.live,]
	
	pinn.live.table = table(PINNIPEDS.DETERMINED$Initial.Injury.Assessment, PINNIPEDS.DETERMINED$Final.Injury.Assessment)
	whale.live.table = table(WHALES.DETERMINED$Initial.Injury.Assessment, WHALES.DETERMINED$Final.Injury.Assessment)
	sm.cet.live.table = table(SMALL.CET.DETERMINED$Initial.Injury.Assessment, SMALL.CET.DETERMINED$Final.Injury.Assessment)
	
# How many large whale records involved an attempt to free the whale from gear?
  rescue.mentioned = grep("FAST|WET|REMOVE|CUT|TEAM|RELEASED|DISTENTANGLEMENT TEAM|GEAR REMOVED|REMOVED GEAR|DISENTANGLED|SUCCESSFULLY", WHALES.DETERMINED$Comments)
  no.rescue = grep("NO RESIGHT|NO RESCUE|COULD NOT LOCATE|NO ATTEMPT|NO RESPONSE", WHALES.DETERMINED$Comments)
	
# Summarize NSI/SI/Mortality by Interaction.Type
# add LOF and PBR fields to SI database                    ##############******** DO WE FEED THIS BACK TO XLS?
## identify commercial vs non.commercial sources of mortality / injury
	commercial = c("GILLNET FISHERY", "UNIDENTIFIED FISHERY INTERACTION", "CA SWORDFISH DRIFT GILLNET FISHERY",
	                       "CA HALIBUT AND WHITE SEABASS SET GILLNET FISHERY", "UNIDENTIFIED NET FISHERY", "CATCH SHARES BOTTOM TRAWL",
	                       "UNIDENTIFIED POT - TRAP FISHERY ENTANGLEMENT", "OPEN ACCESS CA HALIBUT BOTTOM TRAWL", "LIMITED ENTRY SABLEFISH HOOK AND LINE",
	                       "DUNGENESS CRAB POT FISHERY", "CRAB POT FISHERY", "CATCH SHARES HOOK AND LINE", "LIMITED ENTRY TRAWL BOTTOM TRAWL",
	                       "AT-SEA HAKE PELAGIC TRAWL", "HAWAII LONGLINE SHALLOW SET FISHERY", "CATCHER PROCESSOR MIDWATER TRAWL", "LOBSTER TRAP ENTANGLEMENT",
	                       "CA SPOT PRAWN TRAP FISHERY", "LIMITED ENTRY CA HALIBUT BOTTOM TRAWL", "TRAWL FISHERY (UNKNOWN)", "BAIT BARGE NET ENTANGLEMENT",
	                       "COD POT FISHERY ENTANGLEMENT", "LIMITED ENTRY TRAWL HOOK AND LINE", "OPEN ACCESS FIXED GEAR HOOK AND LINE", "SEAL BOMB",
	                       "SHORESIDE HAKE MIDWATER TRAWL", "WEST COAST GROUNDFISH NEARSHORE FIXED GEAR FISHERY", "WEST COAST GROUNDFISH TRAWL (PACIFIC HAKE AT-SEA PROCESSING COMPONENT) FISHERY",
	                       "AK BSAI ATKA MACKEREL TRAWL", "AK BSAI FLATFISH TRAWL", "AQUACULTURE FACILITY ENTANGLEMENT", "CA - OR - WA SABLEFISH POT FISHERY",
	                       "CRAB POT FISHERY - HOOK AND LINE FISHERY", "FIXED GEAR SABLEFISH FISHERY OBSERVER REPORT", "LIMITED ENTRY SABLEFISH POT FISHERY",
	                       "NEARSHORE HOOK AND LINE", "PACIFIC WHITING TRAWL FISHERY", "RED KING CRAB POT FISHERY ENTANGLEMENT", "SALMON GILLNET FISHERY",
	                       "SET GILLNET FISHERY INTERACTION", "SHOOTING - GILLNET FISHERY", "UNIDENTIFIED NET FISHERY - SHOOTING", "UNIDENTIFIED NET FISHERY - VESSEL STRIKE",
	                       "LIMITED ENTRY FIXED GEAR DTL HOOK AND LINE"
	                       )
	non.commercial = c("HOOK AND LINE FISHERY", "SHOOTING", "POWER PLANT ENTRAINMENT", "MARINE DEBRIS",
	                               "VESSEL STRIKE", "OIL  -  TAR", "SCIENTIFIC RESEARCH", "AUTHORIZED REMOVAL", "HARASSMENT",
	                               "HUMAN-INDUCED ABANDONMENT", "DOG ATTACK", "NORTHERN WASHINGTON MARINE SET GILLNET FISHERY, TRIBAL",
	                               "VEHICLE COLLISION", "GILLNET FISHERY, TRIBAL", "UNIDENTIFIED HUMAN INTERACTION", "STAB WOUNDS", "BLUNT FORCE TRAUMA",
	                               "ENTRAPMENT", "GAFF INJURY", "UNDERWATER DETONATION EXERCISES", "UNAUTHORIZED REMOVAL", "ELECTRIC SHOCK", 
	                               "NON-FISHERY ENTANGLEMENT", "NORTHERN WASHINGTON MARINE DRIFT GILLNET FISHERY, TRIBAL", "ROPE", "SHOOTING - HOOK AND LINE FISHERY",
	                               "SPEARING", "HARASSMENT - DOG ATTACK", "HARPOON", "NORTHERN WASHINGTON OCEAN TROLL FISHERY, TRIBAL", "SPRAY PAINTING",
	                               "TRIBAL CRAB POT GEAR"
	                               )
## add LOF codes to non.comm.fishery records 
	LOF.col.num = which(names(x)=="COUNT.AGAINST.LOF")
	commercial.records = which(x$Interaction.Type%in%commercial)
	non.commercial.records = which(x$Interaction.Type%in%non.commercial)
	x[commercial.records, LOF.col.num] = "Y"
  x[non.commercial.records, LOF.col.num]="N"
  
# the next 2 lines assign initial NSI designations as not counting against List of Fisheries
# [even if an interaction occurred with a commercial fishery, the interaction needs to have
# an initial designation of SI for it to 'count' against the List of Fisheries]
  non.serious.initial = which(x$Initial.Injury.Assessment=="NSI")
  x[non.serious.initial, LOF.col.num] = "N"

# add PBR codes to all records
  PBR.yes = grep("CAPTIVITY|DEAD|SI|SI (PRORATE)", x$Final.Injury.Assessment)
  PBR.no = grep("NSI", x$Final.Injury.Assessment)
  col.num.pbr = which(names(x)=="COUNT.AGAINST.PBR")
	x[PBR.yes, col.num.pbr] = "Y"
	x[PBR.no, col.num.pbr] = "N"

# GENERAL ENTANGLEMENTS
	  all.entanglements = grep("ENTANGLE|TRAP|TRAWL|NET|GILLNET|SEA BOTTOM|WAVE RIDER|DEBRIS", x$Interaction.Type)
	  all.gillnet = grep("GILL", x$Interaction.Type, ignore.case=T)

# Pinniped output and assign records with final SI determinations as proration value = 1
	   pinn = grep("SEA LION|SEAL|PINN|OTAR", x$Species)
     pinn.spp = unique(c(x$Species[pinn]))
     pinn.records = x[pinn,]
     pinn.net.trawl = pinn.records[grep("net|trawl", pinn.records$Interaction.Type, ignore.case=TRUE),]
     pinn.harassment = pinn.records[grep("harass", pinn.records$Interaction.Type, ignore.case=TRUE),]
	   pinn.SI.final = which(x$Final.Injury.Assessment=="SI" & x$Species%in%pinn.spp)
	   x$MSI.Value[pinn.SI.final]=1
  
# How many pinnipeds were injured / killed in commercial gillnet, longline, trawl fisheries?
  pinn.commercial.records = length(which(pinn.records$Interaction.Type%in%commercial))
	
# How many pinnipeds were euthanized?
	euthanized = grep("EUTH", pinn.records$Comments)	

# examine pinniped injury categories
  table(pinn.records$Interaction.Type, pinn.records$Initial.Injury.Assessment, pinn.records$Final.Injury.Assessment)

# large whale output
	whales = grep("BLUE|FIN WHALE|SEI|BRYDE'S|HUMP|GRAY|UNIDENTIFIED WHALE|MINKE|SPERM", x$Species)
	whale.records = x[whales,]
	whales.dead <- whale.records[whale.records$Final.Injury.Assessment=="DEAD",]
	# # write CSV to check by eye
	# write.csv(whale.records, "X-yr Large Whales.csv", fileEncoding="UTF-8")
	
# small cetaceans 
	small.cet = grep("DOL|PORP|SHORT|KILLER|BAIRD'S|BLAINVILLE", x$Species)
	small.cet.spp = unique(c(x$Species[small.cet]))
	# assign small cetacean records with final SI designation as MSI.Value=1
	sm.cet.SI.final = which(x$Final.Injury.Assessment=="SI" & x$Species%in%small.cet.spp)
	x$MSI.Value[sm.cet.SI.final]=1
	
	small.cet.records = x[small.cet,]
	small.cet.causes = as.data.frame((table(small.cet.records$Interaction.Type)))
	small.cet.causes = small.cet.causes[order(small.cet.causes[,2], decreasing=T, na.last=F), ]
	
	systematic.small.cet.records <- length(grep("SCIENTIFIC RESEARCH", small.cet.records$Interaction.Type, ignore.case=TRUE)) +
	                                 length(grep("OBSERVER PROGRAM", small.cet.records$Comments, ignore.case=TRUE))
	small.cet.gillnet.trawl.unid.fishery <- length(grep("GILLNET|TRAWL|RESEARCH|UNIDENTIFIED FISHERY INTERACTION", small.cet.records$Interaction.Type, ignore.case=TRUE))
	
#  Rehab mentioned
     rehab = grep("REHAB", x$Comments)
     rehab = x[rehab,]
     table(rehab$Interaction.Type, rehab$SI.Criteria.Supporting.Assessment)

# Harbor seal stock assignments by county in WA state   ##############******** DO WE FEED THIS BACK TO XLS?
# c("WHATCOM", "SKAGIT", "SNOHOMISH", "ISLAND", "SAN JUAN", "KING") = WA N INLAND WATERS
# c("THURSTON") = S PUGET SOUND
# c("PIERCE") = S PUGET SOUND *or* WA N INLAND WATERS
# c("KITSAP")  = WA N INLAND WATERS *or* HOOD CANAL *or* S PUGET SOUND
# c("JEFFERSON") = WA N INLAND WATERS *or* HOOD CANAL *or* OR/WA COAST
# c("CLALLAM") = WA N INLAND WATERS *or* OR/WA COAST
# c("MASON") = HOOD CANAL *or* S PUGET SOUND
	
# write 5-year csv output file with added codes (move manually to MSI-report/data directory after checking)
	write.csv(x, paste(min.year, "_", max.year, "_5yr_MSI.csv", sep=""), row.names=F, fileEncoding="UTF-8")
	
# write all data through most-recently published report for Github upload (move manually to MSI-report/data directory after checking)
	MSI.all.data <- data[data$Year<=max.year,]
	
	### BUT ###
	# > msi.all.5 <- MSI.all.data %>% filter(Year %in% 2019:2023)
	# > all.equal(x, msi.all.5)
	# [1] "Component “MSI.Value”: Mean absolute difference: 1"  
	# [2] "Component “COUNT.AGAINST.LOF”: 168 string mismatches"
	# [3] "Component “COUNT.AGAINST.PBR”: 19 string mismatches" 
	
	# write.csv(MSI.all.data, paste("HCMSI_Records_SWFSC_", min(data$Year), "_",
	                              "LatestPublishedYear",
	                              ".csv", sep=""), row.names=F, fileEncoding="UTF-8")
	
# extract each species
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
  # small cetaceans
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
	
#	write temporary data file
	save.image("SeriousInjuryReport.RData")

#* Then run AppendixTable.R *#
	