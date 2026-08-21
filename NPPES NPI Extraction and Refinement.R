############################################################
# AHFTC NPI Cohort Construction
#
# PURPOSE
# Identifies AHFTC specialists from NPPES taxonomy data,
# exports the initial cohort for manual review, then imports
# the reviewed cohort and creates the final NPI-level dataset.
#
# REQUIRED INPUTS
#   1. NPPES Data Dissemination file
#      https://download.cms.gov/nppes/NPI_Files.html
#
#   2. ahftc_npis_final.xlsx
#      Manually reviewed version of the initial extracted cohort
#
# OUTPUT
#   ahftc_npis
#      Final reviewed AHFTC physician cohort used in downstream
#      physician-, manufacturer-, and device-level analyses.
############################################################

library(data.table)
library(readxl)
library(openxlsx)

# Project directory
base_dir <- "/Users/arman/Documents/Cardiology Industry Payments"

# NPPES file
npi_file <- "/Users/arman/Documents/Cardiology Industry Payments/NPPES_Data_Dissemination_August_2026/npidata_pfile_20050523-20260809.csv"

if (!file.exists(npi_file))
  stop("NPPES file not found. Check npi_file path.")


############################################################
# 1. Extract AHFTC providers from NPPES
############################################################

base_cols <- c(
  "NPI","Entity Type Code",
  "Provider Organization Name (Legal Business Name)",
  "Provider Other Organization Name",
  "Provider Last Name (Legal Name)",
  "Provider First Name","Provider Middle Name",
  "Provider Credential Text",
  "Provider Business Practice Location Address City Name",
  "Provider Business Practice Location Address State Name",
  "Provider Business Practice Location Address Postal Code",
  "Provider Sex Code","Provider Enumeration Date"
)

tax_cols <- paste0("Healthcare Provider Taxonomy Code_",1:15)

# Confirm expected columns are present
npi_names <- names(fread(file=npi_file,nrows=0))
missing_cols <- setdiff(c(base_cols,tax_cols),npi_names)

if (length(missing_cols)) {
  cat("\nMissing NPPES columns:\n")
  print(missing_cols)
  stop("Required NPPES columns are missing.")
}

# Read only columns needed for cohort construction
npi_data <- fread(
  file=npi_file,
  select=c(base_cols,tax_cols)
)

# Standardize taxonomy values
for (col in tax_cols)
  npi_data[[col]] <- toupper(trimws(npi_data[[col]]))


############################################################
# 2. Identify individual providers with AHFTC taxonomy
############################################################

ahftc_code <- "207RA0001X"

npi_data[,ahftc_flag:=
           rowSums(.SD==ahftc_code,na.rm=TRUE)>0,
         .SDcols=tax_cols]

ahftc_npis <- npi_data[
  `Entity Type Code`=="1" & ahftc_flag==TRUE
]


############################################################
# 3. Rename key physician variables
############################################################

setnames(
  ahftc_npis,
  old=c(
    "Provider Organization Name (Legal Business Name)",
    "Provider Other Organization Name",
    "Provider Last Name (Legal Name)",
    "Provider First Name",
    "Provider Middle Name",
    "Provider Credential Text",
    "Provider Business Practice Location Address City Name",
    "Provider Business Practice Location Address State Name",
    "Provider Business Practice Location Address Postal Code",
    "Provider Sex Code",
    "Provider Enumeration Date"
  ),
  new=c(
    "org_name","org_other_name","last_name","first_name","middle_name",
    "credential","city","state","zip","gender","enumeration_date"
  )
)

# NPI variable used in downstream scripts
ahftc_npis[,physician_npi:=as.character(NPI)]


############################################################
# 4. Restrict to physicians
############################################################

# Blank credentials are retained for manual review
ahftc_npis <- ahftc_npis[
  grepl(
    "\\b(MD|M\\.D\\.|DO|D\\.O\\.|MBBS|M\\.B\\.B\\.S)\\b",
    credential,
    ignore.case=TRUE
  ) |
    credential=="" |
    is.na(credential)
]


############################################################
# 5. Identify additional cardiovascular subspecialties
############################################################

tax_cols_in_ahftc <- grep(
  "^Healthcare Provider Taxonomy Code_",
  names(ahftc_npis),
  value=TRUE
)

flag_presence <- function(dt,code,tax_cols) {
  as.integer(
    rowSums(dt[,..tax_cols]==code,na.rm=TRUE)>0
  )
}

ahftc_npis[,has_interventional:=
             flag_presence(.SD,"207RI0011X",tax_cols_in_ahftc)]

ahftc_npis[,has_ep:=
             flag_presence(.SD,"207RC0001X",tax_cols_in_ahftc)]

ahftc_npis[,has_critcare:=
             flag_presence(.SD,"207RC0200X",tax_cols_in_ahftc)]

ahftc_npis[,overlap_count:=
             has_interventional+has_ep+has_critcare]

ahftc_npis[,dual_type:=fcase(
  overlap_count==0,"AHFTC_only",
  overlap_count>1,"Multiple",
  has_interventional==1,"Interventional",
  has_ep==1,"EP",
  has_critcare==1,"CriticalCare"
)]


############################################################
# 6. Initial checks and export for manual review
############################################################

print(
  ahftc_npis[,.N,by=dual_type][order(-N)]
)

dup_npi_count <- ahftc_npis[duplicated(NPI),.N]

cat("\nInitial cohort size:",nrow(ahftc_npis),"\n")
cat("Duplicate NPIs found:",dup_npi_count,"\n")

setorder(ahftc_npis,state,city,last_name,first_name)

initial_file <- file.path(base_dir,"ahftc_npis.xlsx")

write.xlsx(
  ahftc_npis,
  file=initial_file,
  sheetName="AHFTC_NPIs",
  overwrite=TRUE
)

cat("Initial cohort exported to:\n",initial_file,"\n")


############################################################
# MANUAL REVIEW STEP
#
# Review ahftc_npis.xlsx outside R and save the finalized
# version as:
#
#   ahftc_npis_final.xlsx
#
# in the project directory above.
############################################################


############################################################
# 7. Import manually reviewed cohort
############################################################

reviewed_file <- file.path(base_dir,"ahftc_npis_final.xlsx")

if (!file.exists(reviewed_file))
  stop("Reviewed file not found. Complete manual review and save as ahftc_npis_final.xlsx.")

ahftc_npis <- as.data.table(read_excel(reviewed_file))

# Ensure downstream NPI variable exists
if (!"physician_npi" %in% names(ahftc_npis)) {
  ahftc_npis[,physician_npi:=as.character(NPI)]
} else {
  ahftc_npis[,physician_npi:=as.character(physician_npi)]
}


############################################################
# 8. Apply manual subspecialty classifications
############################################################

# Manual comments identify additional EP/interventional specialists
ahftc_npis[,comment_clean:=toupper(trimws(as.character(comment)))]

ahftc_npis[
  !is.na(comment_clean) & grepl("\\bEP\\b",comment_clean),
  has_ep:=1L
]

ahftc_npis[
  !is.na(comment_clean) & grepl("INTERVENTIONAL",comment_clean),
  has_interventional:=1L
]

# Recalculate dual-specialty variables
ahftc_npis[,dual_specialty:=as.integer(
  has_interventional==1 |
    has_ep==1 |
    has_critcare==1
)]

ahftc_npis[,dual_type:=fifelse(
  has_interventional==1 & has_ep==0 & has_critcare==0,
  "AHFTC+Interventional",
  fifelse(
    has_interventional==0 & has_ep==1 & has_critcare==0,
    "AHFTC+EP",
    fifelse(
      has_interventional==0 & has_ep==0 & has_critcare==1,
      "AHFTC+CritCare",
      fifelse(
        has_interventional+has_ep+has_critcare>1,
        "AHFTC+Multi",
        "AHFTC_only"
      )
    )
  )
)]


############################################################
# 9. Apply manual exclusions and clean variables
############################################################

if ("remove" %in% names(ahftc_npis))
  ahftc_npis <- ahftc_npis[is.na(remove) | remove!=1]

ahftc_npis[,comment_clean:=NULL]

# Missing authorship classifications treated as No
if ("guideline_author" %in% names(ahftc_npis))
  ahftc_npis[is.na(guideline_author),guideline_author:="No"]


############################################################
# 10. Final checks
############################################################

cat("\nFinal cohort size:",nrow(ahftc_npis),"\n")
cat("Unique NPIs:",uniqueN(ahftc_npis$physician_npi),"\n\n")

if ("academic_status" %in% names(ahftc_npis)) {
  academic_counts <- ahftc_npis[,.N,by=academic_status][order(-N)]
  print(academic_counts)
}

if ("guideline_author" %in% names(ahftc_npis)) {
  guideline_counts <- ahftc_npis[,.N,by=guideline_author][order(-N)]
  print(guideline_counts)
}

dual_specialty_counts <- ahftc_npis[,.N,by=dual_type][order(-N)]
print(dual_specialty_counts)

View(ahftc_npis)