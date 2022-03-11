library(tidyverse)
library(lubridate)
library(readxl)
library(openxlsx)
library(readr)
library(moderndive)

Ireland_Weather_Station_Codes <- read_csv("Ireland Weather Station Codes.csv")

Weather_Station_Headings <- read_csv("Weather Station CSV Files/hly518.csv", 
                                     col_names = FALSE, skip = 3, n_max=17) %>%
  separate(X1, c("Accronym", "Name"), sep = ":\\s+-\\s+") %>%
  mutate(Name = str_squish(str_to_title(
    gsub("\\b(and|for)\\b |erature|sure|(it)?ation|Predominant|Weather|\\(.*", "", Name))),
         Name = gsub(" ", "_", Name)) %>%
  filter(!grepl("^(ind|date)$", Accronym))

Weather_Station_Files <- list.files(path = "Weather Station CSV Files/", 
                                    pattern = "(csv)", recursive = T, full.names = T, include.dirs = T)


Weather_Station_Details <- map_df(Weather_Station_Files, ~read_csv(.x ,n_max=3, col_names = F) %>%
                                    mutate(filename = .x, .before = 1)) %>%
  splitstackshape::cSplit(., "X1", sep = ",", "long") %>%
  separate(X1, into = c("key", "value"), sep = ": ?") %>%
  pivot_wider(names_from = key, values_from = value) %>%
  set_names(~ str_replace_all(.," ", "_"))

Weather_Station_Merge <- map_df(Weather_Station_Files, ~
                                  # Skip Until 'Date' Column Appears
                                  data.table::fread(.x, skip = "Date") %>%
                                  select(-ind) %>%
                                  mutate(filename = .x,
                                         Station_No = gsub(".*hly|\\.csv", "", filename), .before = 1,
                                         date = dmy_hm(date),
                                         Decade = year(floor_date(date, years(10))),
                                         Year = year(date),
                                         Quarter = quarter(date, with_year = FALSE, fiscal_start = 1),
                                         Month = month(date),
                                         Week_No = week(date),
                                         Hour_No = hour(date)) %>%
                                  separate(date, into = c("Date", "Time"), sep = " ")) %>%
  # Separate Converts Time Object to Char
  mutate(Date = as_date(Date),
         Time = hms(Time)) %>%
  left_join(Weather_Station_Details, .) %>%
  select(-filename)

 
Weather_Station_Merge <- Weather_Station_Merge %>%
  # Replace/Substitute Certain Column Names
  rename_with(.fn = ~ gsubfn::gsubfn(".*", setNames(as.list(
    # Need to Define: names() like 'Weather_Station_Merge' - Will not work: names(.)
  Weather_Station_Headings$Name), names(Weather_Station_Merge)[14:28]), .), 
  .cols = c(rain:clamt))


write_rds(Weather_Station_Merge, "Weather_Station_Merge.rds")
write.xlsx(Weather_Station_Merge, "Weather_Station_Merge.xlsx", colWidths = "auto")

Weather_Station_Merge_National <- Weather_Station_Merge %>%
  #group_by(across(c(Station_Name:Time))) %>%
  group_by(across(c(Date:Hour_No))) %>%
  summarise(across(c(Precip_Amount:Cloud_Amount), ~ mean(.x, na.rm = TRUE))) %>%
  ungroup()

########################
Eirgrid_System_Data <- list.files(path = "Weather/System_Data/", 
                                    pattern = "(xlsx)", recursive = T, full.names = T, include.dirs = T)

# Arrange Files By Size
Eirgrid_File_Info<- file.info(Eirgrid_System_Data)
Eirgrid_System_Data[match(1:length(Eirgrid_System_Data),rank(-Eirgrid_File_Info$size))]

Eirgrid_System <- map_df(Eirgrid_System_Data, ~read_excel(.x, )) %>%
  set_names(~ str_replace_all(.," ", "_")) %>%
  set_names(~ str_replace_all(., "(Dem|Avail|Gen)(.*)", "\\1")) %>%
  select(DateTime, GMT_Offset, SNSP, matches("IE_"), matches("^NI\\S{0,20}$")) %>%
  arrange(DateTime) %>%
  separate(DateTime, into = c("Date", "Time"), sep = " ") %>%
  mutate(Date = ymd(Date),
         Time = hms(Time),
         Hour_No = hour(Time), .after = 2)

write.xlsx(Eirgrid_System, "Eirgrid_System.xlsx", colWidths = "auto")
  
Eirgrid_Hourly <- Eirgrid_System %>%
  group_by(across(c(Date,Hour_No, GMT_Offset))) %>%
  summarise(across(c(SNSP:NI_Solar_Gen), ~ mean(.x, na.rm = T))) %>%
  ungroup()

Eirgrid_Dates <- Eirgrid_System %>%
  select(Date, Time, GMT_Offset) %>%
  mutate(Decade = year(floor_date(Date, years(10))),
         Year = year(Date),
         Quarter = quarter(Date, with_year = FALSE, fiscal_start = 1),
         Month = month(Date, label = T),
         Week_No = week(Date),
         Weekday_No = wday(Date, label = T),
         Hour_No = hour(Time))
  

Weather_Eirgrid <- left_join(Eirgrid_Dates, Eirgrid_Hourly) %>%
  left_join(. , Weather_Station_Merge_National) %>%
  mutate(across(GMT_Offset:Hour_No, ~ as.factor(as.character(.))))
  
write.xlsx(Weather_Eirgrid, "Weather_Eirgrid.xlsx", colWidths = "auto")

Weather_Eirgrid_Log <- Weather_Eirgrid %>%
  # Removes Nevative Values by Converting Negative Celsius Values to Zero
  mutate(across(contains("Temp"), ~ weathermetrics::celsius.to.kelvin(., round = 5))) %>%
  # Log1p deals with Zero Values - Inverse - expm1
  mutate(across(SNSP:Cloud_Amount, ~ log1p(.))) #%>%
  #mutate(across(NI_Gen:Cloud_Amount, ~ 10^(.x + 1 - max(.x))))

write.xlsx(Weather_Eirgrid_Log, "Weather_Eirgrid_Log.xlsx", colWidths = "auto")

write.xlsx(mget(ls(pattern = "^(Weather_Eirgrid|Eirgrid_System(_Log)?)$")), file = "Weather_Eirgrid_List.xlsx", 
           asTable = T, tableStyle = "TableStyleMedium2", colWidths = "auto")

################################################################################################

Weather_Eirgrid_Demand_Train <- lm(IE_Dem ~  Year + Month + Weekday_No + Hour_No
                                 + Precip_Amount + Air_Temp + Wet_Bulb_Temp + Dew_Point_Temp + Relative_Humidity
                                 + Vapour_Pres + Mean_Sea_Level_Pres + Mean_Wind_Speed + Wind_Direction
                                 + Synop_Code_Past + Sunshine_Dur + Visibility + Cloud_Height + Cloud_Amount,
                                 data = Weather_Eirgrid_Log)

Weather_Eirgrid_Dem_Reg <- get_regression_points(Weather_Eirgrid_Demand_Train, digits = 6)

Weather_Eirgrid_Dem_Reg_Sum <- get_regression_table(Weather_Eirgrid_Demand_Train)

get_regression_points(Weather_Eirgrid_Demand_Train) %>%
  mutate(sq_residuals = residual^2) %>%
  summarize(rmse = sqrt(mean(sq_residuals))) %>%
  mutate(Accuracy = 1- rmse)

Weather_Eirgrid_Generation_Train <- lm(IE_Gen ~  Year + Month + Weekday_No + Hour_No
                                   + Precip_Amount + Air_Temp + Wet_Bulb_Temp + Dew_Point_Temp + Relative_Humidity
                                   + Vapour_Pres + Mean_Sea_Level_Pres + Mean_Wind_Speed + Wind_Direction
                                   + Synop_Code_Past + Sunshine_Dur + Visibility + Cloud_Height + Cloud_Amount,
                                   data = Weather_Eirgrid_Log)

Weather_Eirgrid_Gen_Reg <- get_regression_points(Weather_Eirgrid_Generation_Train, digits = 6)

Weather_Eirgrid_Gen_Reg_Sum <- get_regression_table(Weather_Eirgrid_Generation_Train)

get_regression_points(Weather_Eirgrid_Generation_Train) %>%
  mutate(sq_residuals = residual^2) %>%
  summarize(rmse = sqrt(mean(sq_residuals))) %>%
  mutate(Accuracy = 1- rmse)

Weather_Eirgrid_Wind_Gen_Train <- lm(IE_Wind_Gen ~  Year + Month + Weekday_No + Hour_No + Precip_Amount 
                                     + Air_Temp + Wet_Bulb_Temp + Dew_Point_Temp + Relative_Humidity
                                       + Vapour_Pres + Mean_Sea_Level_Pres + Mean_Wind_Speed + Wind_Direction
                                       + Synop_Code_Past + Sunshine_Dur + Visibility + Cloud_Height + Cloud_Amount,
                                       data = Weather_Eirgrid_Log)

Weather_Eirgrid_Wind_Gen_Reg <- get_regression_points(Weather_Eirgrid_Wind_Gen_Train, digits = 6)

Weather_Eirgrid_Wind_Gen_Sum <- get_regression_table(Weather_Eirgrid_Wind_Gen_Train)

get_regression_points(Weather_Eirgrid_Wind_Gen_Train) %>%
  mutate(sq_residuals = residual^2) %>%
  summarize(rmse = sqrt(mean(sq_residuals))) %>%
  mutate(Accuracy = 1- rmse)

Weather_Eirgrid_Solar_Data <- Weather_Eirgrid_Log %>%
  filter(!is.nan(NI_Solar_Gen)) %>%
  # Deals with Negative and Zero Values
  mutate(across(SNSP:Cloud_Amount, ~ log1p(.)))
  
Weather_Eirgrid_Solar_Gen_Train <- lm(NI_Solar_Gen ~  Year + Month + Weekday_No + Hour_No
                                      + Precip_Amount + Air_Temp + Wet_Bulb_Temp + Dew_Point_Temp + Relative_Humidity
                                      + Vapour_Pres + Mean_Sea_Level_Pres + Mean_Wind_Speed + Wind_Direction
                                      + Synop_Code_Past + Sunshine_Dur + Visibility + Cloud_Height + Cloud_Amount,
                                      data = Weather_Eirgrid_Solar_Data)

Weather_Eirgrid_Solar_Gen_Reg <- get_regression_points(Weather_Eirgrid_Solar_Gen_Train, digits = 6)

Weather_Eirgrid_Solar_Gen_Sum <- get_regression_table(Weather_Eirgrid_Solar_Gen_Train)

get_regression_points(Weather_Eirgrid_Solar_Gen_Train) %>%
  mutate(sq_residuals = residual^2) %>%
  summarize(rmse = sqrt(mean(sq_residuals))) %>%
  mutate(Accuracy = 1- rmse)

# Write Multiple Data Frames to Worksheets - Styles them to Tables
write.xlsx(mget(ls(pattern = "Sum$")), file = "Eirgrid_Weather_Regressions.xlsx", 
           asTable = T, tableStyle = "TableStyleMedium2", colWidths = "auto")
