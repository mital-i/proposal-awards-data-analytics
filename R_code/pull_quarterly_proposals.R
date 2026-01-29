library("readxl")
library("openxlsx")
library("dplyr")
library("magrittr")
library("lubridate")
library("stringr")
library("ggplot2")

# Load proposal data and rename column
dwq_proposals <- rename(
  read_excel(
    "C:\\Users\\sarkisj\\Downloads\\proposals_df.xlsx"
    ),
  `Award PI Campus ID` = `Proposal PI Campus ID`
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
BioSciFacultyOnly <- dwq_proposals %>%
  filter(`Award PI Campus ID` %in% campusIDs)

# Add department info
BioSciFacultyOnly <- left_join(BioSciFacultyOnly, deptsAndIDs, by = "Award PI Campus ID")

# Filter proposals submitted within fiscal year
proposals <- BioSciFacultyOnly %>%
  filter(between(`Proposal Process Date`, as.Date('2025-07-01'), as.Date('2025-09-30')))

# Filter by proposal type: new, renewal, supplement
proposal_type <- proposals %>%
  filter(`Proposal Type Code` %in% c("1", "3", "7"))

# Add fiscal year column
proposal_type <- proposal_type %>%
  mutate(`Fiscal Year` = year(`Proposal Process Date`) + ifelse(month(`Proposal Process Date`) >= 7, 1, 0))

# Add academic quarter column
proposal_type <- proposal_type %>%
  mutate(`Quarter` = quarter(`Proposal Process Date`, with_year = FALSE, fiscal_start = 7))

# Select columns for internal analysis
finalColumns <- proposal_type %>%
  select(
    'Proposal Development #', 'Award PI Campus ID', 'Proposal PI Last Name', 'Proposal PI First Name',
    'Department', 'Proposal Lead Unit Name', 'Proposal School Name', 'Proposal Sponsor Name',
    'Proposal Prime Sponsor Name', 'Fiscal Year', 'Quarter', 'Proposal Project Title',
    'Proposed Start Date', 'Proposed End Date', 'Proposal Process Date',
    'Proposal Total Direct Cost', 'Proposal F&A Cost', 'Proposal Total Cost',
    'Award Type Description', 'Proposal Type Description', 'Proposal Activity Type Description',
    'Status Description', 'Subaward Flag', 'NIH Activity Code'
  )

# Select columns for RAD Unit website
finalColumnsforWebsite <- proposal_type %>%
  mutate(PI = paste(`Proposal PI Last Name`, `Proposal PI First Name`, sep = ", ")) %>%
  rename(
    Sponsor = 'Proposal Sponsor Name',
    `Prime Sponsor` = 'Proposal Prime Sponsor Name',
    `Project Title` = 'Proposal Project Title',
    `Proposal Type` = 'Proposal Type Description',
    `NIH Mechanism` = 'NIH Activity Code'
  ) %>%
  select(PI, Department, `Project Title`, `Prime Sponsor`, Sponsor, `NIH Mechanism`, `Proposal Type`)

# Write Excel files
write.xlsx(finalColumns, "C:\\Users\\sarkisj\\OneDrive - UC Irvine\\BioSci Research Development\\Data Analysis, DWQuery\\Analysis, Quarterly\\Proposals 2025-07-01 to 2025-09-30.xlsx")
write.xlsx(finalColumnsforWebsite, "C:\\Users\\sarkisj\\OneDrive - UC Irvine\\BioSci Research Development\\Data Analysis, DWQuery\\Analysis, Quarterly\\BioSci Faculty Proposals for Website 2025-07-01 to 2025-09-30.xlsx")