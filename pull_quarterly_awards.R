install.packages(c("readxl", "openxlsx", "dplyr", "magrittr", "lubridate", "stringr", "ggplot2"))

library("readxl")
library("openxlsx")
library("dplyr")
library("magrittr")
library("lubridate")
library("stringr")
library("ggplot2")

# Load raw award data
dwq_awards <- read_excel(
  "C:\\Users\\sarkisj\\Downloads\\awards_df.xlsx"
)

# Load Master Faculty spreadsheet
faculty_master <- read_excel(
  "C:\\Users\\sarkisj\\OneDrive - UC Irvine\\BioSci Research Development\\Faculty_Master.xlsx"
  , sheet = "Master"
)

# Convert 'Award PI Campus ID' to numeric
faculty_master$`Award PI Campus ID` <- as.numeric(faculty_master$`Award PI Campus ID`)

# Select department and campus ID columns
deptsAndIDs <- faculty_master %>%
  select("Award PI Campus ID", "Department")

# Extract campus IDs
campusIDs <- deptsAndIDs[[1]]

# Filter by BioSci faculty only
BioSciFacultyOnly <- dwq_awards %>%
  filter(`Award PI Campus ID` %in% campusIDs)

# Add department info
BioSciFacultyOnly <- left_join(BioSciFacultyOnly, deptsAndIDs, by = "Award PI Campus ID")

# Filter awards by start date
awards2019ToPresent <- BioSciFacultyOnly %>%
  filter(between(`Award Finalize Date`, as.Date('2025-01-01'), as.Date('2025-12-31')))

# Filter by award transaction type
newOrRenewal <- awards2019ToPresent %>%
  filter(`Award Transaction Type Code` %in% c("1", "9", "13"))

# Add fiscal year column
newOrRenewal <- newOrRenewal %>%
  mutate(`Fiscal Year` = year(`Award Finalize Date`) + ifelse(month(`Award Finalize Date`) >= 7, 1, 0))

# Add academic quarter column
newOrRenewal <- newOrRenewal %>%
  mutate(`Quarter` = quarter(`Award Finalize Date`, with_year = FALSE, fiscal_start = 7))

# Select columns for internal analysis
finalColumns <- newOrRenewal %>%
  select(
    'Sponsor Award ID', 'Award PI Campus ID', 'Award PI Last Name', 'Award PI First Name',
    'Department', 'Award Lead Unit Name', 'Award Sponsor Name', 'Award Prime Sponsor Name',
    'Fiscal Year', 'Quarter', 'Award Project Title', 'Award Finalize Date', 'Project Start Date', 'Project End Date',
    'Award Obligated Direct Cost', 'Award Obligated F&A Cost', 'Award Obligated Total Cost',
    'Award Type Description', 'Award Transaction Type Description', 'NIH Activity Code'
  )

# Select columns for RAD Unit website
finalColumnsforWebsite <- newOrRenewal %>%
  select(
    'Award PI Last Name', 'Award PI First Name', 'Department', 'Award Sponsor Name',
    'Award Prime Sponsor Name', 'Award Project Title', 'Award Type Description',
    'Award Transaction Type Description', 'NIH Activity Code'
  ) %>%
  mutate(
    PI = paste(`Award PI Last Name`, `Award PI First Name`, sep = ", "),
    `Fellowship Recipient` = ''
  ) %>%
  rename(
    Sponsor = 'Award Sponsor Name',
    `Prime Sponsor` = 'Award Prime Sponsor Name',
    `Project Title` = 'Award Project Title',
    `Award Type` = 'Award Transaction Type Description',
    `NIH Mechanism` = 'NIH Activity Code'
  ) %>%
  select(PI, `Fellowship Recipient`, Department, `Project Title`, `Prime Sponsor`, Sponsor, `NIH Mechanism`, `Award Type`)

# Write Excel files
write.xlsx(finalColumns, "C:\\Users\\sarkisj\\OneDrive - UC Irvine\\BioSci Research Development\\Data Analysis, DWQuery\\Analysis, Quarterly\\Awards 2025-01-01 to 2025-12-31.xlsx")
write.xlsx(finalColumnsforWebsite, "C:\\Users\\sarkisj\\OneDrive - UC Irvine\\BioSci Research Development\\Data Analysis, DWQuery\\Analysis, Quarterly\\BioSci Faculty Awards for Website 2025-01-01 to 2025-12-31.xlsx")