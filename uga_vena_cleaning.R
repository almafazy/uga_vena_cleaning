library(tidyverse)
library(kobold)

# COMMODITY CLEANING

test <- read_xls_form("input/commodity_cleaning.xlsx", cleaning = "cleaning", choices = "choices", data = "data")

test$cleaning <- mutate(test$cleaning, sheet = test$survey$sheet[match(name, test$survey$name)])

for (i in test$data_sheets$sheets[-1]) {
  test[[i]] <- group_by(test[[i]], uuid) %>% mutate(position = 1:n()) %>% ungroup
}

indices <- pmap_dbl(
  list(
    test$cleaning$sheet,
    test$cleaning$uuid,
    test$cleaning$position
  ),

  function(sheet, uuid, position) {
    if (is.na(position) | is.na(uuid)) {
      NA
    } else {
      location_check <- test[[sheet]][["uuid"]] == uuid & test[[sheet]][["position"]] == position
      if (sum(location_check) == 0) {
        -99
      } else {
        test[[sheet]][["index"]][location_check]
      }
    }
  }

)

test$cleaning$index <- indices

openxlsx::write.xlsx(test, "output/commodity_cleaning_log.xlsx")

test$cleaning <- filter(test$cleaning, index != -99 | is.na(index))

test2 <- kobold_cleaner(test)

openxlsx::write.xlsx(test2, "output/commodity_cleaned.xlsx")

# VENA HOUSEHOLD CLEANING

test <- read_xls_form("input/vena_cleaning.xlsx", choices = "choices", data = "data", cleaning = "cleaning")

test$cleaning <- mutate(test$cleaning, sheet = ifelse(type != "remove_loop_entry", test$survey$sheet[match(name, test$survey$name)], sheet))

for (i in test$data_sheets$sheets[-1]) {
  test[[i]] <- mutate(test[[i]], index = 1:n()) %>% group_by(uuid) %>% mutate(position = 1:n()) %>% ungroup
}

indices <- pmap_dbl(
  list(
    test$cleaning$sheet,
    test$cleaning$uuid,
    test$cleaning$position
  ),

  function(sheet, uuid, position) {
    if (is.na(position) | is.na(uuid)) {
      NA
    } else {
      location_check <- test[[sheet]][["uuid"]] == uuid & test[[sheet]][["position"]] == position
      if (sum(location_check) == 0) {
        -99
      } else {
        test[[sheet]][["index"]][location_check]
      }
    }
  }

)

test$cleaning$index <- indices

test$cleaning <- mutate(test$cleaning,
                        index = ifelse(sheet == "data", NA, index),
                        position = ifelse(sheet == "data", NA, position))

openxlsx::write.xlsx(test, "input/vena_cleaning_final.xlsx")

test$cleaning <- filter(test$cleaning, type != "add_loop_entry")
test$cleaning <- mutate(test$cleaning, sheet = ifelse(!is.na(relevant), "data", sheet))

test2 <- kobold_cleaner(test)

openxlsx::write.xlsx(test2, "output/vena_cleaned.xlsx")
