############################################################
# AHFTC Open Payments Data Generation
# Open Payments, 2018-2024
#
# PURPOSE
# Generates the physician-, manufacturer-, and device-level
# datasets used in downstream analyses.
#
# REQUIRED INPUTS
#
# 1. ahftc_npis
#    Final manually reviewed AHFTC physician cohort generated
#    from NPPES.
#
# 2. CMS Open Payments General Payment files, 2018-2024
#    Downloaded from:
#    https://openpaymentsdata.cms.gov/datasets/download
#
# MANUALLY REVIEWED INPUTS
#
# 1. all_unique_companies_master_tagged.xlsx
#    Manufacturer classification file. cv_device_flag = 1
#    identifies cardiovascular device manufacturers included
#    in the analysis.
#
#    If this file does not exist, the script creates an
#    automated draft and stops for manual review.
#
# 2. master_company_device_labeled.xlsx
#    Manually reviewed device-name file containing a "combine"
#    column identifying synonymous device names.
#
#    If this file does not exist, the script creates the raw
#    device file and stops for manual review.
#
# MAIN OUTPUTS
#
# 1. ahftc_summary_nature
#    Physician-level dataset
#    Expected structure: 1,271 x 242
#
# 2. master_company
#    Manufacturer-level dataset
#    Expected structure: 171 x 98
#
# 3. master_company_device_consolidated
#    Consolidated device-level dataset
#    Expected structure: 382 x 121
#
# 4. top50_devices
#
# ADDITIONAL OUTPUT
#
# ever_over
#    One row per physician-manufacturer relationship, indicating
#    whether >=$5,000 was received from that manufacturer during
#    at least one calendar year.
#
# EXPECTED STUDY TOTALS
#
# 1,271 AHFTC specialists
# 265 specialists with >=1 significant device manufacturer COI
# 96 specialists with COIs involving >=2 manufacturers
# 439 unique physician-manufacturer COIs
############################################################

############################################################
# 1. Setup
############################################################

library(data.table)
library(readxl)
library(openxlsx)
library(writexl)
library(lubridate)

base_dir <- path.expand(
  "~/Library/Mobile Documents/com~apple~CloudDocs/Documents/Cardiology Industry Payments"
)

file_map <- list(
  "2018"=file.path(base_dir,"OP_DTL_GNRL_PGYR2018_P06302025_06162025.csv"),
  "2019"=file.path(base_dir,"OP_DTL_GNRL_PGYR2019_P06302025_06162025.csv"),
  "2020"=file.path(base_dir,"OP_DTL_GNRL_PGYR2020_P06302025_06162025.csv"),
  "2021"=file.path(base_dir,"OP_DTL_GNRL_PGYR2021_P06302025_06162025.csv"),
  "2022"=file.path(base_dir,"OP_DTL_GNRL_PGYR2022_P06302025_06162025.csv"),
  "2023"=file.path(base_dir,"OP_DTL_GNRL_PGYR2023_P06302025_06162025.csv"),
  "2024"=file.path(base_dir,"OP_DTL_GNRL_PGYR2024_P06302025_06162025.csv")
)

years <- 2018:2024
threshold <- 5000

company_col <- "Applicable_Manufacturer_or_Applicable_GPO_Making_Payment_Name"
product_cols <- paste0(
  "Indicate_Drug_or_Biological_or_Device_or_Medical_Supply_",1:5
)

categories <- c(
  "Consulting.fees",
  "Food.and.beverage",
  "Honoraria",
  "Other",
  "Speaking.fees",
  "Travel.and.lodging"
)

setDT(ahftc_npis)
ahftc_npis[,physician_npi := as.character(physician_npi)]

stopifnot(
  nrow(ahftc_npis)==1271,
  uniqueN(ahftc_npis$physician_npi)==1271
)

############################################################
# 2. Shared classifications
############################################################

# Collapse Open Payments nature-of-payment fields
nature_map <- list(
  "Speaking.fees"=c(
    "Compensation for services other than consulting, including serving as faculty or as a speaker at a venue other than a continuing education program",
    "Compensation for serving as faculty or as a speaker for a non-accredited and noncertified continuing education program",
    "Compensation for serving as faculty or as a speaker for an accredited or certified continuing education program",
    "Compensation for serving as faculty or as a speaker for a medical education program"
  ),
  "Consulting.fees"="Consulting Fee",
  "Food.and.beverage"="Food and Beverage",
  "Travel.and.lodging"="Travel and Lodging",
  "Honoraria"="Honoraria"
)

classify_nature <- function(x) {
  for (cat in names(nature_map))
    if (x %in% nature_map[[cat]]) return(cat)
  "Other"
}

# Consolidate manufacturer subsidiaries, historical names, and mergers (requires prior review of outputs)
merge_map <- list(
  "CARDINAL HEALTH 200, LLC"=c(
    "CARDINAL HEALTH 200 LLC","CARDINAL HEALTH 200, LLC"
  ),
  "CVRX INC."=c("CVRX, INC.","CVRX INC."),
  "HOLOGIC"=c("HOLOGIC INC","HOLOGIC, LLC"),
  "KCI USA, INC."=c("KCI USA, INC","KCI USA, INC."),
  "MAQUET CARDIOVASCULAR, LLC."=c(
    "MAQUET CARDIOVASCULAR L.L.C.",
    "MAQUET CARDIOVASCULAR U.S. SALES, L.L.C."
  ),
  "MEDTRONIC, INC."=c(
    "MEDTRONIC VASCULAR, INC.",
    "MEDTRONIC, INC.",
    "MEDTRONIC USA, INC."
  ),
  "ROCHE"=c(
    "ROCHE DIAGNOSTICS CORPORATION",
    "ROCHE DIAGNOSTICS INTERNATIONAL LTD",
    "ROCHE MOLECULAR SYSTEMS, INC."
  ),
  "SENSIBLE MEDICAL INNOVATIONS, INC."=c(
    "SENSIBLE MEDICAL INNOVATIONS INC",
    "SENSIBLE MEDICAL INNOVATIONS LTD"
  ),
  "SMITH AND NEPHEW, INC."=c(
    "SMITH & NEPHEW, INC.",
    "SMITH+NEPHEW, INC."
  ),
  "NUWELLIS INC"=c(
    "NUWELLIS, INC.",
    "CHF SOLUTIONS, INC"
  ),
  "ZOLL MEDICAL CORP"=c(
    "RESPICARDIA, INC.",
    "ZOLL CIRCULATION INC",
    "ZOLL MEDICAL CORPORATION",
    "ZOLL RESPICARDIA, INC.",
    "ZOLL SERVICES LLC (A/K/A ZOLL LIFECOR CORP)"
  )
)

apply_company_map <- function(dt,col="company_clean") {
  for (new_name in names(merge_map))
    dt[get(col) %in% merge_map[[new_name]],(col) := new_name]
  invisible(dt)
}

standardize_company <- function(x) {
  out <- toupper(trimws(x))
  for (new_name in names(merge_map))
    out[out %in% toupper(merge_map[[new_name]])] <- new_name
  out
}

############################################################
# 3. Identify potential cardiovascular device manufacturers
############################################################

all_companies <- rbindlist(lapply(names(file_map),function(yr) {
  cat("Reading manufacturer names for",yr,"...\n")
  x <- fread(file_map[[yr]],select=company_col,showProgress=FALSE)
  company <- unique(trimws(x[[company_col]]))
  company <- company[!is.na(company) & company!=""]
  data.table(year=yr,company=company)
}))

master_companies <- all_companies[
  ,.(earliest_year=min(as.integer(year)),n_years=uniqueN(year)),
  by=company
]
master_companies[,company_type := NA_character_]

# This file is automated and can be regenerated
write.xlsx(
  master_companies,
  file.path(base_dir,"all_unique_companies_master.xlsx"),
  overwrite=TRUE
)

normalize_for_matching <- function(x) {
  x <- tolower(as.character(x))
  x <- gsub(
    "\\b(inc|llc|corp(?:oration)?|ltd|co\\.?|gmbh|s\\.?a\\.?|plc|ag|nv|sa|pte|pvt|holdings?|limited|lp|llp|kk|pty|bv|ab|sas)\\b",
    " ",x
  )
  x <- gsub("[^a-z0-9]+"," ",x)
  trimws(x)
}

dt <- copy(master_companies)
dt[,`:=`(
  company_norm=normalize_for_matching(company),
  company_lc=tolower(company),
  cv_device_flag=0L,
  cv_area=NA_character_,
  tag_source=NA_character_
)]

# Seed list of known cardiovascular device manufacturers
cv_seed <- data.table(
  pattern=c(
    "abiomed","edwards(\\s+lifesciences)?","boston\\s+scientific",
    "\\bmedtronic\\b","biotronik","\\bcvrx\\b","shockwave\\s+medical",
    "cardiac\\s+dimensions","ancora\\s+heart","\\bcordis\\b",
    "\\bpenumbra\\b","\\bterumo\\b","\\bteleflex\\b","heartflow",
    "\\birhythm\\b","alivecor","acist\\s+medical","acutus\\s+medical",
    "angiodynamics","volta\\s+medical","veryan\\s+medical",
    "silk\\s*road\\s+medical","\\binari\\b","artivion|cryolife",
    "w\\.?\\s*l\\.?\\s*gore","merit\\s+medical","asahi\\s+intecc",
    "atricure","liva\\s*nova","philips(\\s|$)|volcano",
    "siemens\\s+healthineers","ge\\s*healthcare",
    "canon\\s+medical|toshiba\\s+medical"
  ),
  cv_area=c(
    "Mechanical support (Impella)",
    "Structural heart (TAVR/TEER)",
    "EP/CRT/ICD; Structural/LAA; Coronary/peripheral",
    "EP/CRT/ICD; Coronary/peripheral; Structural heart",
    "EP/CRT/ICD",
    "Neuromodulation (Barostim)",
    "IVL (coronary/peripheral)",
    "Structural heart (Carillon)",
    "Structural heart (ventricular remodeling)",
    "Coronary/peripheral",
    "Peripheral/vascular",
    "Coronary/peripheral",
    "Vascular access",
    "FFR-CT (diagnostics)",
    "Ambulatory ECG",
    "Ambulatory ECG",
    "Contrast/hemodynamics",
    "EP mapping",
    "Peripheral vascular",
    "EP AI software",
    "Peripheral vascular",
    "TCAR (stroke prevention)",
    "Venous thrombectomy",
    "Cardiac surgery / tissue valves",
    "Vascular grafts / stent grafts",
    "Coronary/peripheral",
    "Guidewires (coronary/peripheral)",
    "Ablation (EP / surgical)",
    "Cardiac surgery / neuromodulation",
    "Cath-lab imaging / IVUS (adjacent)",
    "Imaging / angio (adjacent)",
    "Imaging / angio (adjacent)",
    "Imaging / angio (adjacent)"
  )
)

has_pat <- function(pattern,norm,raw) {
  grepl(pattern,norm,perl=TRUE) | grepl(pattern,raw,perl=TRUE)
}

for (i in seq_len(nrow(cv_seed))) {
  hit <- has_pat(cv_seed$pattern[i],dt$company_norm,dt$company_lc)
  dt[hit,`:=`(company_type="Device",cv_device_flag=1L)]
  dt[hit & is.na(cv_area),cv_area := cv_seed$cv_area[i]]
  dt[hit & is.na(tag_source),
     tag_source := paste0("cv_seed: ",cv_seed$pattern[i])]
}

# Flag additional likely device manufacturers for manual review
kw_device <- paste(c(
  "device","devices","surgical","surgery","spine","orthop",
  "implant","prosthet","vascular","endosc","robot","catheter",
  "stent","valve","defibrillat","pacemaker","electrophysiolog",
  "balloon","guidewire","sheath","graft","monitor(ing)?\\b",
  "diagnostic(s)?\\b","imaging","ophthalm","otolog|otolaryng|ent",
  "dental|orthodont","optics","wound","arthro","cardio(?!logy)"
),collapse="|")

kw_drug <- paste(c(
  "pharma","pharmaceutical","therapeutic","therapeutics",
  "biotech","biopharma","bioscience","biosciences",
  "laborator","medicin","oncology","genetic","genomics",
  "cell","gene","immune","immuno"
),collapse="|")

left <- which(is.na(dt$company_type) | dt$company_type=="")
if (length(left)) {
  device_match <- has_pat(kw_device,dt$company_norm[left],dt$company_lc[left])
  drug_match <- has_pat(kw_drug,dt$company_norm[left],dt$company_lc[left])
  idx <- left[device_match & !drug_match]
  
  if (length(idx)) {
    dt[idx,company_type := "Device"]
    dt[idx & is.na(tag_source),tag_source := "keyword:device"]
  }
}


############################################################
# 4. Manufacturer manual-review checkpoint
############################################################

tagged_file <- file.path(
  base_dir,
  "all_unique_companies_master_tagged.xlsx"
)

# Never overwrite an existing manually reviewed file
if (!file.exists(tagged_file)) {
  
  setcolorder(dt,unique(c(
    "company","earliest_year","n_years","company_type",
    "cv_device_flag","cv_area","tag_source","company_norm","company_lc"
  )))
  
  debug_hits <- rbindlist(list(
    head(dt[cv_device_flag==1L,
            .(company,company_type,cv_area,tag_source)],50),
    head(dt[company_type=="Device" & cv_device_flag==0L,
            .(company,company_type,cv_area,tag_source)],50)
  ),fill=TRUE)
  
  still_blank <- dt[
    is.na(company_type) | company_type=="",
    .(company)
  ]
  
  wb <- createWorkbook()
  addWorksheet(wb,"Tagged")
  writeData(wb,"Tagged",dt)
  addWorksheet(wb,"Debug_Matches")
  writeData(wb,"Debug_Matches",debug_hits)
  addWorksheet(wb,"Needs_Review")
  writeData(wb,"Needs_Review",still_blank)
  saveWorkbook(wb,tagged_file,overwrite=FALSE)
  
  stop(
    paste0(
      "\nManufacturer review file created:\n",tagged_file,
      "\n\nReview the file and set cv_device_flag = 1 for all ",
      "manufacturers included as cardiovascular device manufacturers. ",
      "Save the reviewed file with the same name, then rerun the script."
    ),
    call.=FALSE
  )
}

# Read the existing manually reviewed file
device_companies <- as.data.table(read_excel(tagged_file))
device_companies <- device_companies[cv_device_flag==1,.(company)]
device_companies[,company := toupper(trimws(company))]

cat("\nReviewed device manufacturers:",nrow(device_companies),"\n")

############################################################
# 5. Extract annual physician- and manufacturer-level data
############################################################

npis <- unique(ahftc_npis[,.(physician_npi)])
N_AHFTC_total <- uniqueN(npis$physician_npi)

extract_payment_year <- function(op_file,year) {
  
  op_vars <- c(
    "Covered_Recipient_NPI",
    company_col,
    "Total_Amount_of_Payment_USDollars",
    "Nature_of_Payment_or_Transfer_of_Value",
    product_cols
  )
  
  gp <- fread(op_file,select=op_vars,showProgress=TRUE)
  setnames(gp,"Covered_Recipient_NPI","physician_npi")
  gp[,physician_npi := as.character(physician_npi)]
  gp <- gp[physician_npi %in% npis$physician_npi]
  
  # Device classification uses the original Open Payments company name
  gp[,company_clean := toupper(trimws(get(company_col)))]
  gp[,is_device_company := company_clean %in% device_companies$company]
  
  gp[,is_device_product := rowSums(as.data.frame(
    lapply(.SD,function(x) x %in% c("Device","Medical Supply"))
  ))>0,.SDcols=product_cols]
  
  gp[,is_device := is_device_company | is_device_product]
  
  gp[,payment_category := vapply(
    Nature_of_Payment_or_Transfer_of_Value,
    classify_nature,
    character(1)
  )]
  
  
  ####################################
  # Physician-level annual summary
  ####################################
  
  total_summary <- gp[,.(total_payments=sum(
    Total_Amount_of_Payment_USDollars,na.rm=TRUE
  ),n_total_transactions=.N),by=physician_npi]
  
  setnames(
    total_summary,
    c("total_payments","n_total_transactions"),
    paste0("yr",year,c("_total_payments","_n_total_transactions"))
  )
  
  device_summary <- gp[is_device==TRUE,.(device_payments_combined=sum(
    Total_Amount_of_Payment_USDollars,na.rm=TRUE
  ),n_device_tx_combined=.N),by=physician_npi]
  
  setnames(
    device_summary,
    c("device_payments_combined","n_device_tx_combined"),
    paste0("yr",year,c(
      "_device_payments_combined",
      "_n_device_tx_combined"
    ))
  )
  
  # All payments by category
  total_cat <- gp[,.(total_payments_by_cat=sum(
    Total_Amount_of_Payment_USDollars,na.rm=TRUE
  ),n_total_tx_by_cat=.N),by=.(physician_npi,payment_category)]
  
  total_wide <- dcast(
    total_cat,
    physician_npi~payment_category,
    value.var=c("total_payments_by_cat","n_total_tx_by_cat"),
    fun.aggregate=sum,
    fill=0
  )
  
  # Device payments by category
  device_cat <- gp[is_device==TRUE,.(device_payments_by_cat=sum(
    Total_Amount_of_Payment_USDollars,na.rm=TRUE
  ),n_device_tx_by_cat=.N),by=.(physician_npi,payment_category)]
  
  device_wide <- dcast(
    device_cat,
    physician_npi~payment_category,
    value.var=c("device_payments_by_cat","n_device_tx_by_cat"),
    fun.aggregate=sum,
    fill=0
  )
  
  # Ensure every category exists every year
  for (cat in categories) {
    
    total_value_col <- paste0("total_payments_by_cat_",cat)
    total_n_col <- paste0("n_total_tx_by_cat_",cat)
    device_value_col <- paste0("device_payments_by_cat_",cat)
    device_n_col <- paste0("n_device_tx_by_cat_",cat)
    
    if (!total_value_col %in% names(total_wide))
      total_wide[,(total_value_col) := 0]
    if (!total_n_col %in% names(total_wide))
      total_wide[,(total_n_col) := 0L]
    
    if (!device_value_col %in% names(device_wide))
      device_wide[,(device_value_col) := 0]
    if (!device_n_col %in% names(device_wide))
      device_wide[,(device_n_col) := 0L]
  }
  
  setnames(
    total_wide,
    setdiff(names(total_wide),"physician_npi"),
    paste0("yr",year,"_",setdiff(names(total_wide),"physician_npi"))
  )
  
  setnames(
    device_wide,
    setdiff(names(device_wide),"physician_npi"),
    paste0("yr",year,"_",setdiff(names(device_wide),"physician_npi"))
  )
  
  physician_out <- Reduce(
    function(x,y) merge(x,y,by="physician_npi",all=TRUE),
    list(total_summary,device_summary,total_wide,device_wide)
  )
  
  
  ##########################################################
  # Manufacturer-level annual summary
  ##########################################################
  
  device_gp <- gp[is_device==TRUE]
  
  # Consolidation occurs before defining manufacturer-specific COI
  apply_company_map(device_gp)
  
  phys_pairs <- unique(
    device_gp[,.(company_clean,physician_npi)]
  )
  
  per_phys_year <- device_gp[,.(total_paid_year=sum(
    Total_Amount_of_Payment_USDollars,na.rm=TRUE
  )),by=.(company_clean,physician_npi)]
  
  per_phys_year[,`:=`(
    year=as.integer(year),
    over5000_year=as.integer(total_paid_year>=threshold)
  )]
  
  company_tx <- device_gp[,.(n_tx=.N,value=sum(
    Total_Amount_of_Payment_USDollars,na.rm=TRUE
  )),by=company_clean]
  
  setnames(
    company_tx,
    c("n_tx","value"),
    paste0(c("n_tx_","value_"),year)
  )
  
  company_phys <- device_gp[
    ,.(n_physicians=uniqueN(physician_npi)),
    by=company_clean
  ]
  
  setnames(
    company_phys,
    "n_physicians",
    paste0("n_physicians_",year)
  )
  
  company_over <- per_phys_year[
    ,.(n_phys_over5000=sum(over5000_year,na.rm=TRUE)),
    by=company_clean
  ]
  
  setnames(
    company_over,
    "n_phys_over5000",
    paste0("n_phys_over5000_",year)
  )
  
  company_out <- Reduce(
    function(x,y) merge(x,y,by="company_clean",all=TRUE),
    list(company_tx,company_phys,company_over)
  )
  
  company_out[,(paste0("pct_phys_over5000_among_paid_",year)) :=
                fifelse(
                  get(paste0("n_physicians_",year))>0,
                  100*get(paste0("n_phys_over5000_",year))/
                    get(paste0("n_physicians_",year)),
                  0
                )
  ]
  
  company_out[,(paste0("pct_phys_over5000_of_cohort_",year)) :=
                100*get(paste0("n_phys_over5000_",year))/N_AHFTC_total
  ]
  
  company_cat <- device_gp[,.(value=sum(
    Total_Amount_of_Payment_USDollars,na.rm=TRUE
  )),by=.(company_clean,payment_category)]
  
  company_cat <- dcast(
    company_cat,
    company_clean~payment_category,
    value.var="value",
    fun.aggregate=sum,
    fill=0
  )
  
  for (cat in categories)
    if (!cat %in% names(company_cat))
      company_cat[,(cat) := 0]
  
  setcolorder(
    company_cat,
    c("company_clean",categories)
  )
  
  setnames(
    company_cat,
    categories,
    paste0("cat_",year,"_",categories)
  )
  
  company_out <- merge(
    company_out,
    company_cat,
    by="company_clean",
    all=TRUE
  )
  
  for (col in names(company_out))
    if (is.numeric(company_out[[col]]))
      set(company_out,which(is.na(company_out[[col]])),col,0)
  
  list(
    physician=physician_out,
    company=company_out,
    phys_pairs=phys_pairs,
    per_phys_year=per_phys_year
  )
}

############################################################
# 6. Run all years
############################################################

year_results <- lapply(names(file_map),function(yr) {
  cat("Processing Open Payments",yr,"...\n")
  extract_payment_year(file_map[[yr]],yr)
})
names(year_results) <- names(file_map)

############################################################
# 7. Build manufacturer-level dataset
############################################################

master_company <- Reduce(
  function(x,y) merge(x,y,by="company_clean",all=TRUE),
  lapply(year_results,`[[`,"company")
)

for (col in names(master_company))
  if (is.numeric(master_company[[col]]))
    set(master_company,which(is.na(master_company[[col]])),col,0)

# Unique physicians ever paid by each manufacturer
phys_pairs_all <- unique(
  rbindlist(lapply(year_results,`[[`,"phys_pairs"),use.names=TRUE)
)

phys_total <- phys_pairs_all[
  ,.(n_physicians_total=as.integer(uniqueN(physician_npi))),
  by=company_clean
]

master_company <- merge(
  master_company,
  phys_total,
  by="company_clean",
  all.x=TRUE
)

master_company[is.na(n_physicians_total),n_physicians_total := 0L]

# Physician-manufacturer COI across the full study period
per_phys_all <- rbindlist(
  lapply(year_results,`[[`,"per_phys_year"),
  use.names=TRUE,
  fill=TRUE
)

ever_over <- per_phys_all[
  ,.(ever_over5000=as.integer(any(over5000_year==1L))),
  by=.(company_clean,physician_npi)
]

over_totals <- ever_over[
  ,.(n_phys_over5000_total=as.integer(sum(ever_over5000))),
  by=company_clean
]

master_company <- merge(
  master_company,
  over_totals,
  by="company_clean",
  all.x=TRUE
)

master_company[
  is.na(n_phys_over5000_total),
  n_phys_over5000_total := 0L
]

master_company[,pct_phys_over5000_among_paid_total :=
                 fifelse(
                   n_physicians_total>0,
                   100*n_phys_over5000_total/n_physicians_total,
                   0
                 )
]

master_company[,pct_phys_over5000_of_cohort_total :=
                 100*n_phys_over5000_total/N_AHFTC_total
]

# Cumulative payment value and transactions
master_company[,total_value := rowSums(
  .SD,na.rm=TRUE
),.SDcols=paste0("value_",years)]

master_company[,total_tx := rowSums(
  .SD,na.rm=TRUE
),.SDcols=paste0("n_tx_",years)]

# Cumulative category totals
for (cat in categories) {
  cols <- paste0("cat_",years,"_",cat)
  master_company[,(paste0("cat_total_",cat)) :=
                   rowSums(.SD,na.rm=TRUE),
                 .SDcols=cols
  ]
}

setorder(master_company,-total_value)
master_company[,Rank := as.integer(.I)]

# Reproduce historical 98-column order
company_total_cols <- c(
  "Rank",
  "company_clean",
  "n_physicians_total",
  "n_phys_over5000_total",
  "pct_phys_over5000_among_paid_total",
  "pct_phys_over5000_of_cohort_total",
  "total_tx",
  "total_value",
  paste0("cat_total_",categories)
)

company_year_cols <- unlist(lapply(years,function(yr) {
  c(
    paste0("n_tx_",yr),
    paste0("value_",yr),
    paste0("n_physicians_",yr),
    paste0("n_phys_over5000_",yr),
    paste0("pct_phys_over5000_among_paid_",yr),
    paste0("pct_phys_over5000_of_cohort_",yr),
    paste0("cat_",yr,"_",categories)
  )
}))

setcolorder(
  master_company,
  c(company_total_cols,company_year_cols)
)

manufacturer_file <- file.path(
  base_dir,
  "Device_Companies_Master_2018_2024.xlsx"
)

write_xlsx(master_company,manufacturer_file)


############################################################
# 8. Build physician-level dataset
############################################################

# Ensure both historical guideline variables are available
if (!"guideline_author" %in% names(ahftc_npis) &&
    "guideline_author_new" %in% names(ahftc_npis))
  ahftc_npis[,guideline_author := guideline_author_new]

if (!"guideline_author_old" %in% names(ahftc_npis))
  ahftc_npis[,guideline_author_old := guideline_author]

# Reproduce historical data types
for (v in c("academic_status","guideline_author_old","guideline_author")) {
  if (v %in% names(ahftc_npis)) {
    ahftc_npis[is.na(get(v)),(v) := "No"]
    ahftc_npis[,(v) := factor(get(v),levels=c("No","Yes"))]
  }
}

for (v in c("has_interventional","has_ep","has_critcare"))
  if (v %in% names(ahftc_npis))
    ahftc_npis[,(v) := as.numeric(get(v))]

if ("dual_specialty" %in% names(ahftc_npis))
  ahftc_npis[,dual_specialty := as.integer(dual_specialty)]

base_vars <- c(
  "physician_npi","first_name","last_name","gender",
  "academic_status","guideline_author_old","guideline_author",
  "institution_name","credential","city","state","zip",
  "enumeration_date","has_interventional","has_ep","has_critcare",
  "dual_specialty","dual_type"
)

ahftc_summary_nature <- copy(
  ahftc_npis[,..base_vars]
)

# Merge annual payment variables
for (yr in names(year_results)) {
  ahftc_summary_nature <- merge(
    ahftc_summary_nature,
    year_results[[yr]]$physician,
    by="physician_npi",
    all.x=TRUE,
    sort=FALSE
  )
}

# Missing payment values indicate no payments
annual_payment_cols <- grep(
  "^yr[0-9]{4}_",
  names(ahftc_summary_nature),
  value=TRUE
)

for (col in annual_payment_cols) {
  if (is.numeric(ahftc_summary_nature[[col]]))
    set(
      ahftc_summary_nature,
      which(is.na(ahftc_summary_nature[[col]])),
      col,
      0
    )
}


############################################################
# 9. Physician cumulative payments and COI variables
############################################################

ahftc_summary_nature[,device_payments_all_years :=
                       rowSums(.SD,na.rm=TRUE),
                     .SDcols=paste0("yr",years,"_device_payments_combined")
]

ahftc_summary_nature[,total_payments_all_years :=
                       rowSums(.SD,na.rm=TRUE),
                     .SDcols=paste0("yr",years,"_total_payments")
]

# Historical aggregate-payment COI variables
for (yr in years) {
  ahftc_summary_nature[
    ,(paste0("yr",yr,"_sig_device_coi")) :=
      as.integer(get(paste0("yr",yr,"_device_payments_combined"))>threshold)
  ]
  
  ahftc_summary_nature[
    ,(paste0("yr",yr,"_sig_total_coi")) :=
      as.integer(get(paste0("yr",yr,"_total_payments"))>threshold)
  ]
}

device_sig_cols <- paste0("yr",years,"_sig_device_coi")
total_sig_cols <- paste0("yr",years,"_sig_total_coi")

ahftc_summary_nature[,sig_device_coi_ever :=
                       as.integer(rowSums(.SD)>0),
                     .SDcols=device_sig_cols
]

ahftc_summary_nature[,sig_total_coi_ever :=
                       as.integer(rowSums(.SD)>0),
                     .SDcols=total_sig_cols
]

ahftc_summary_nature[,sig_device_coi_all_years :=
                       as.integer(rowSums(.SD)==length(years)),
                     .SDcols=device_sig_cols
]

ahftc_summary_nature[,sig_total_coi_all_years :=
                       as.integer(rowSums(.SD)==length(years)),
                     .SDcols=total_sig_cols
]

# Manufacturer-specific significant device COI
physician_device_coi <- ever_over[
  ever_over5000==1L,
  .(n_device_companies_over5000=as.integer(uniqueN(company_clean))),
  by=physician_npi
]

ahftc_summary_nature <- merge(
  ahftc_summary_nature,
  physician_device_coi,
  by="physician_npi",
  all.x=TRUE,
  sort=FALSE
)

ahftc_summary_nature[
  is.na(n_device_companies_over5000),
  n_device_companies_over5000 := 0L
]

ahftc_summary_nature[,sig_device_coi_ever_new :=
                       as.integer(n_device_companies_over5000>=1L)
]

ahftc_summary_nature[,multiple_device_company_coi :=
                       as.integer(n_device_companies_over5000>=2L)
]

# Compatibility variables retained from the historical dataset
ahftc_summary_nature[,any_company_over5000 :=
                       as.integer(n_device_companies_over5000>=1L)
]

ahftc_summary_nature[,two_or_more_companies_over5000 :=
                       as.integer(n_device_companies_over5000>=2L)
]


############################################################
# 10. Physician geographic and NPI variables
############################################################

ahftc_summary_nature[,enumeration_date := parse_date_time(
  enumeration_date,
  orders=c("mdy","ymd","dmy"),
  quiet=TRUE
)]

# Fixed reference date preserves the study-era NPI duration
npi_reference_date <- as.Date("2025-12-01")

ahftc_summary_nature[,years_since_npi :=
                       as.integer(floor(
                         interval(enumeration_date,npi_reference_date)/years(1)
                       ))
]

census_divisions <- list(
  "New England"=c("CT","ME","MA","NH","RI","VT"),
  "Middle Atlantic"=c("NJ","NY","PA"),
  "East North Central"=c("IL","IN","MI","OH","WI"),
  "West North Central"=c("IA","KS","MN","MO","NE","ND","SD"),
  "South Atlantic"=c("DE","DC","FL","GA","MD","NC","SC","VA","WV"),
  "East South Central"=c("AL","KY","MS","TN"),
  "West South Central"=c("AR","LA","OK","TX"),
  "Mountain"=c("AZ","CO","ID","MT","NV","NM","UT","WY"),
  "Pacific"=c("AK","CA","HI","OR","WA")
)

census_regions <- list(
  "Northeast"=unlist(
    census_divisions[c("New England","Middle Atlantic")]
  ),
  "Midwest"=unlist(
    census_divisions[c("East North Central","West North Central")]
  ),
  "South"=unlist(
    census_divisions[c(
      "South Atlantic","East South Central","West South Central"
    )]
  ),
  "West"=unlist(
    census_divisions[c("Mountain","Pacific")]
  )
)

get_division <- function(st) {
  for (x in names(census_divisions))
    if (st %in% census_divisions[[x]]) return(x)
  "Other"
}

get_region <- function(st) {
  for (x in names(census_regions))
    if (st %in% census_regions[[x]]) return(x)
  "Other"
}

ahftc_summary_nature[,division :=
                       vapply(state,get_division,character(1))
]

ahftc_summary_nature[,region :=
                       vapply(state,get_region,character(1))
]


############################################################
# 11. Reproduce historical physician-level column structure
############################################################

physician_front_cols <- c(
  "physician_npi",
  "first_name",
  "last_name",
  "gender",
  "academic_status",
  "guideline_author_old",
  "guideline_author",
  "institution_name",
  "credential",
  "city",
  "state",
  "division",
  "region",
  "zip",
  "enumeration_date",
  "years_since_npi",
  "has_interventional",
  "has_ep",
  "has_critcare",
  "dual_specialty",
  "dual_type",
  "sig_device_coi_ever",
  "n_device_companies_over5000",
  "sig_device_coi_ever_new",
  "multiple_device_company_coi",
  "sig_total_coi_ever",
  "device_payments_all_years",
  "total_payments_all_years"
)

physician_year_cols <- unlist(lapply(years,function(yr) {
  c(
    paste0("yr",yr,"_total_payments"),
    paste0("yr",yr,"_n_total_transactions"),
    paste0("yr",yr,"_device_payments_combined"),
    paste0("yr",yr,"_n_device_tx_combined"),
    
    paste0("yr",yr,"_total_payments_by_cat_",categories),
    paste0("yr",yr,"_n_total_tx_by_cat_",categories),
    
    paste0("yr",yr,"_device_payments_by_cat_",categories),
    paste0("yr",yr,"_n_device_tx_by_cat_",categories)
  )
}))

# Historical auxiliary COI fields appeared after the main payment data
physician_tail_cols <- c(
  as.vector(rbind(device_sig_cols,total_sig_cols)),
  "sig_device_coi_all_years",
  "sig_total_coi_all_years",
  "any_company_over5000",
  "two_or_more_companies_over5000"
)

setcolorder(
  ahftc_summary_nature,
  c(
    physician_front_cols,
    physician_year_cols,
    physician_tail_cols
  )
)

# Main study-level validation
stopifnot(
  nrow(ahftc_summary_nature)==1271,
  uniqueN(ahftc_summary_nature$physician_npi)==1271,
  sum(ahftc_summary_nature$sig_device_coi_ever_new)==265,
  sum(ahftc_summary_nature$multiple_device_company_coi)==96,
  sum(ahftc_summary_nature$n_device_companies_over5000)==439
)

physician_file <- file.path(
  base_dir,
  "AHFTC_Physician_Level_2018_2024.xlsx"
)

write_xlsx(
  ahftc_summary_nature,
  physician_file
)

############################################################
# 12. Device-level extraction
############################################################

# Manufacturer set used in the original device-level analysis
top20_input <- c(
  "ABBOTT LABORATORIES",
  "MEDTRONIC, INC.",
  "ABIOMED",
  "BOSTON SCIENTIFIC CORPORATION",
  "IMPULSE DYNAMICS (USA) INC.",
  "CVRX INC.",
  "ZOLL MEDICAL CORPORATION",
  "SENSIBLE MEDICAL INNOVATIONS, INC.",
  "TRANSMEDICS, INC.",
  "EDWARDS LIFESCIENCES CORPORATION",
  "NUWELLIS, INC.",
  "LIFENET HEALTH",
  "RESPICARDIA, INC.",
  "LIVANOVA USA, INC.",
  "BIOTRONIK INC.",
  "CHF SOLUTIONS, INC",
  "IRHYTHM TECHNOLOGIES, INC.",
  "SIEMENS MEDICAL SOLUTIONS USA, INC.",
  "GE HEALTHCARE",
  "ANCORA HEART, INC."
)

# Preserve historical device-dataset company labels
standardize_device_company <- function(x) {
  out <- standardize_company(x)
  out[out=="NUWELLIS INC"] <- "NUWELLIS, INC."
  out[out=="ZOLL MEDICAL CORP"] <- "ZOLL Medical"
  out
}

top20_canonical <- unique(
  standardize_device_company(top20_input)
)

build_device_year <- function(op_file,year) {
  
  device_name_cols <- paste0(
    "Name_of_Drug_or_Biological_or_Device_or_Medical_Supply_",1:5
  )
  
  indicate_cols <- paste0(
    "Indicate_Drug_or_Biological_or_Device_or_Medical_Supply_",1:5
  )
  
  op_vars <- c(
    "Covered_Recipient_NPI",
    company_col,
    "Total_Amount_of_Payment_USDollars",
    "Nature_of_Payment_or_Transfer_of_Value",
    device_name_cols,
    indicate_cols
  )
  
  gp <- fread(op_file,select=op_vars,showProgress=TRUE)
  setnames(gp,"Covered_Recipient_NPI","physician_npi")
  gp[,physician_npi := as.character(physician_npi)]
  gp <- gp[physician_npi %in% npis$physician_npi]
  
  gp[,company_clean :=
       standardize_device_company(get(company_col))
  ]
  
  gp <- gp[company_clean %in% top20_canonical]
  if (!nrow(gp)) return(data.table())
  
  # Specific device extraction requires an explicit OP device/supply field
  gp[,is_device_row := rowSums(as.data.frame(
    lapply(.SD,function(x) x %in% c("Device","Medical Supply"))
  ))>0,.SDcols=indicate_cols]
  
  gp <- gp[is_device_row==TRUE]
  if (!nrow(gp)) return(data.table())
  
  gp[,payment_category := vapply(
    Nature_of_Payment_or_Transfer_of_Value,
    classify_nature,
    character(1)
  )]
  
  # Remove product names from slots that are not devices
  for (k in seq_along(device_name_cols)) {
    gp[
      !(get(indicate_cols[k]) %in% c("Device","Medical Supply")),
      (device_name_cols[k]) := NA_character_
    ]
  }
  
  # Number of device names attached to each payment
  gp[,n_devices := rowSums(
    !is.na(.SD) & .SD!="",
    na.rm=TRUE
  ),.SDcols=device_name_cols]
  
  gp <- gp[n_devices>0]
  if (!nrow(gp)) return(data.table())
  
  gp_long <- melt(
    gp,
    id.vars=c(
      "physician_npi",
      "company_clean",
      "Total_Amount_of_Payment_USDollars",
      "n_devices",
      "payment_category"
    ),
    measure.vars=device_name_cols,
    variable.name="device_slot",
    value.name="device_raw"
  )
  
  gp_long <- gp_long[
    !is.na(device_raw) & trimws(device_raw)!=""
  ]
  
  gp_long[,device := toupper(trimws(device_raw))]
  
  # Divide payment equally when more than one device is named
  gp_long[,device_amount :=
            Total_Amount_of_Payment_USDollars/n_devices
  ]
  
  # Overall device totals
  base_summary <- gp_long[
    ,.(n_physicians=uniqueN(physician_npi),
       n_transactions=.N,
       total_value=sum(device_amount,na.rm=TRUE)),
    by=.(company_clean,device)
  ]
  
  setnames(
    base_summary,
    c("n_physicians","n_transactions","total_value"),
    paste0(
      c("n_physicians_","n_transactions_","total_value_"),
      year
    )
  )
  
  # Payment value by category
  cat_value <- gp_long[
    ,.(value=sum(device_amount,na.rm=TRUE)),
    by=.(company_clean,device,payment_category)
  ]
  
  cat_value <- dcast(
    cat_value,
    company_clean+device~payment_category,
    value.var="value",
    fun.aggregate=sum,
    fill=0
  )
  
  # Transaction count by category
  cat_n <- gp_long[
    ,.(n_tx=.N),
    by=.(company_clean,device,payment_category)
  ]
  
  cat_n <- dcast(
    cat_n,
    company_clean+device~payment_category,
    value.var="n_tx",
    fun.aggregate=sum,
    fill=0
  )
  
  # Ensure all six categories exist
  for (cat in categories) {
    if (!cat %in% names(cat_value))
      cat_value[,(cat) := 0]
    if (!cat %in% names(cat_n))
      cat_n[,(cat) := 0L]
  }
  
  setcolorder(
    cat_value,
    c("company_clean","device",categories)
  )
  
  setcolorder(
    cat_n,
    c("company_clean","device",categories)
  )
  
  setnames(
    cat_value,
    categories,
    paste0(categories,"_value_",year)
  )
  
  setnames(
    cat_n,
    categories,
    paste0(categories,"_n_tx_",year)
  )
  
  out <- Reduce(
    function(x,y) merge(
      x,y,
      by=c("company_clean","device"),
      all=TRUE
    ),
    list(base_summary,cat_value,cat_n)
  )
  
  for (col in names(out))
    if (is.numeric(out[[col]]))
      set(out,which(is.na(out[[col]])),col,0)
  
  out[]
}


############################################################
# 13. Build raw company-device dataset
############################################################

device_years <- lapply(names(file_map),function(yr) {
  cat("Building device summary for",yr,"...\n")
  build_device_year(file_map[[yr]],yr)
})
names(device_years) <- names(file_map)

device_years <- Filter(
  function(x) is.data.table(x) && nrow(x)>0,
  device_years
)

master_company_device <- Reduce(
  function(x,y) merge(
    x,y,
    by=c("company_clean","device"),
    all=TRUE
  ),
  device_years
)

for (col in names(master_company_device))
  if (is.numeric(master_company_device[[col]]))
    set(
      master_company_device,
      which(is.na(master_company_device[[col]])),
      col,
      0
    )

master_company_device[,total_value_all_years :=
                        rowSums(.SD,na.rm=TRUE),
                      .SDcols=paste0("total_value_",years)
]

master_company_device[,n_transactions_all_years :=
                        rowSums(.SD,na.rm=TRUE),
                      .SDcols=paste0("n_transactions_",years)
]

# Historical cumulative category order
device_total_category_order <- c(
  "Speaking.fees",
  "Consulting.fees",
  "Food.and.beverage",
  "Travel.and.lodging",
  "Honoraria",
  "Other"
)

for (cat in device_total_category_order) {
  
  master_company_device[,(paste0(cat,"_value_total")) :=
                          rowSums(.SD,na.rm=TRUE),
                        .SDcols=paste0(cat,"_value_",years)
  ]
  
  master_company_device[,(paste0(cat,"_n_tx_total")) :=
                          rowSums(.SD,na.rm=TRUE),
                        .SDcols=paste0(cat,"_n_tx_",years)
  ]
}

# Historical 121-column device order
device_front_cols <- c(
  "company_clean",
  "device",
  "total_value_all_years",
  "n_transactions_all_years",
  as.vector(rbind(
    paste0(device_total_category_order,"_value_total"),
    paste0(device_total_category_order,"_n_tx_total")
  ))
)

device_year_cols <- unlist(lapply(years,function(yr) {
  c(
    paste0("n_physicians_",yr),
    paste0("n_transactions_",yr),
    paste0("total_value_",yr),
    paste0(categories,"_value_",yr),
    paste0(categories,"_n_tx_",yr)
  )
}))

setcolorder(
  master_company_device,
  c(device_front_cols,device_year_cols)
)

setorder(
  master_company_device,
  -total_value_all_years
)

raw_device_file <- file.path(
  base_dir,
  "master_company_device_2018_2024_raw_devices.xlsx"
)

write_xlsx(
  master_company_device,
  raw_device_file
)


############################################################
# 14. Device manual-review checkpoint
############################################################

manual_device_file <- file.path(
  base_dir,
  "master_company_device_labeled.xlsx"
)

# This manually reviewed file is never overwritten
if (!file.exists(manual_device_file)) {
  stop(
    paste0(
      "\nRaw device file created:\n",raw_device_file,
      "\n\nReview the device names and save a copy as:\n",
      manual_device_file,
      "\n\nAdd a column named 'combine'. Assign the same value ",
      "to device names referring to the same underlying product. ",
      "Then rerun the script."
    ),
    call.=FALSE
  )
}


############################################################
# 15. Consolidate manually grouped device names
############################################################

manual <- as.data.table(
  read_excel(manual_device_file)
)

manual[,device := toupper(trimws(device))]

manual_groups <- manual[
  !is.na(combine) &
    trimws(as.character(combine))!=""
]

device_merge_map <- split(
  manual_groups$device,
  manual_groups$combine
)

# First listed name within each manually defined group is canonical
device_merge_map <- setNames(
  device_merge_map,
  vapply(device_merge_map,function(x) x[1],character(1))
)

standardize_device <- function(x) {
  out <- toupper(trimws(x))
  for (canonical in names(device_merge_map))
    out[out %in% toupper(device_merge_map[[canonical]])] <- canonical
  out
}

master_company_device[,device :=
                        standardize_device(device)
]

num_cols <- names(master_company_device)[
  vapply(master_company_device,is.numeric,logical(1))
]

master_company_device_consolidated <- master_company_device[
  ,lapply(.SD,sum,na.rm=TRUE),
  by=.(company_clean,device),
  .SDcols=num_cols
]

# Restore historical column order after consolidation
setcolorder(
  master_company_device_consolidated,
  c(device_front_cols,device_year_cols)
)

setorder(
  master_company_device_consolidated,
  -total_value_all_years
)

# TRUEBEAM remains in this full dataset, as in the historical object
final_device_file <- file.path(
  base_dir,
  "master_company_device_2018_2024_consolidated.xlsx"
)

write_xlsx(
  master_company_device_consolidated,
  final_device_file
)


############################################################
# 16. Reproduce historical top50_devices object
############################################################

# Select top 50 first, then remove the noncardiovascular
# radiation therapy device, leaving 49 rows
top50_devices <- copy(
  master_company_device_consolidated[
    order(-total_value_all_years)
  ][1:50]
)

top50_devices <- top50_devices[
  device!="TRUEBEAM"
]

# Shortened device labels used in the figure dataset
device_label_map <- c(
  "BAROSTIM NEO SYSTEM"="BAROSTIM NEO",
  "CLINICAL TRIAL PRODUCT"="CLINICAL TRIAL PROD.",
  "MITRA CLIP SYSTEM"="MITRACLIP",
  "ORGAN CARE SYSTEM"="ORGAN CARE SYS.",
  "VNS - VITARIA"="VNS VITARIA",
  "EDWARDS SAPIEN 3 TRANSCATHETER HEART VALVE"="SAPIEN 3 TAVR"
)

for (old in names(device_label_map))
  top50_devices[device==old,device := device_label_map[[old]]]

# Shortened company labels used in the figure dataset
top50_devices[device=="IMPELLA",company_clean := "ABIOMED"]
top50_devices[device=="CARDIOMEMS",company_clean := "ABBOTT"]
top50_devices[device=="HEARTMATE",company_clean := "ABBOTT"]
top50_devices[device=="OPTIMIZER",company_clean := "IMPULSE DYN."]
top50_devices[device=="BAROSTIM NEO",company_clean := "CVRx"]
top50_devices[device=="LIFEVEST",company_clean := "ZOLL"]
top50_devices[device=="CLINICAL TRIAL PROD.",company_clean := "BOSTON SCI."]
top50_devices[device=="HEARTWARE HVAD",company_clean := "MEDTRONIC"]
top50_devices[device=="MITRACLIP",company_clean := "ABBOTT"]
top50_devices[device=="REDS SYSTEM",company_clean := "SENSIBLE MED."]
top50_devices[device=="REMEDE SYSTEM",company_clean := "ZOLL"]
top50_devices[device=="ORGAN CARE SYS.",company_clean := "TRANSMEDICS"]
top50_devices[device=="AQUADEX",company_clean := "NUWELLIS"]
top50_devices[device=="VNS VITARIA",company_clean := "LIVANOVA"]
top50_devices[device=="WATCHMAN",company_clean := "BOSTON SCI."]
top50_devices[device=="COREVALVE EVOLUT",company_clean := "MEDTRONIC"]
top50_devices[device=="ACCUCINCH",company_clean := "ANCORA"]
top50_devices[device=="ENSITE",company_clean := "ABBOTT"]
top50_devices[device=="ZIO MONITOR",company_clean := "iRHYTHM"]
top50_devices[device=="SAPIEN 3 TAVR",company_clean := "EDWARDS"]

top50_file <- file.path(
  base_dir,
  "top50_devices.xlsx"
)

write_xlsx(
  top50_devices,
  top50_file
)