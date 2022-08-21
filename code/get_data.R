library(openxlsx)
library(stringr)
library(dplyr)
library(tidyr)

parse_data <- function(url_name, sheet_n) {
  ca_ports <- c("Crescent City", "Eureka", "Fort Bragg", "San Francisco", "Monterey")
  or_ports <- c("Astoria", "Tillamook", "Newport", "Coos Bay", "South of Cape Falcon")
  wa_ports <- c("Neah Bay", "La Push", "Westport", "Ilwaco", "Area 4B")

  dat <- read.xlsx(url, sheet = sheet_n, startRow = 1)
  dat[, 1] <- stringr::str_replace(dat[, 1], "a/", "")
  dat[, 1] <- stringr::str_replace(dat[, 1], "b/", "")
  dat[, 1] <- stringr::str_replace(dat[, 1], "c/", "")
  dat[, 1] <- stringr::str_replace(dat[, 1], "d/", "")
  dat[, 1] <- stringr::str_replace(dat[, 1], "e/", "")

  # determine region / tribal / rec or commercial
  indian <- ""
  if (length(grep("Indian", colnames(dat)[1])) > 0) {
    indian <- "Indian"
    port_names <- wa_ports
    state <- "WA"
  }
  if (length(grep("Washington", colnames(dat)[1])) > 0) {
    port_names <- wa_ports
    state <- "WA"
  }
  if (length(grep("Oregon", colnames(dat)[1])) > 0) {
    port_names <- or_ports
    state <- "OR"
  }
  if (length(grep("California", colnames(dat)[1])) > 0) {
    port_names <- ca_ports
    state <- "CA"
  }
  fishery <- "commercial"
  if (length(grep("recreational", colnames(dat)[1])) > 0) {
    fishery <- "recreational"
  }
  if (length(grep("commercial", colnames(dat)[1])) > 0) {
    fishery <- "commercial"
  }
  metric <- "catch"
  if (length(grep("effort", colnames(dat)[1])) > 0) {
    metric <- "effort"
  }

  row_start <- which(dat[1:10, 2] %in% c("Jan.", "Feb.", "Mar.", "Apr.", "May", "Jan.-Apr."))
  col_names <- as.character(dat[row_start, ])
  col_names[1] <- "year"
  indx <- which(col_names %in% c("Season", "Year")) #
  if (length(indx) > 0) { # not all sheets have season summaries
    col_names <- col_names[1:indx[1]]
  }
  n_cols <- length(col_names[-1])
  col_indices <- which(!is.na(match(dat[row_start, ], col_names)))
  n_spp <- length(col_indices) / n_cols # some have chinook + coho

  raw_dat <- dat
  d <- list()
  for (sp in 1:n_spp) {
    first_col <- n_cols * (sp - 1) + 1 + 1 # 2 plus ones because year
    end_col <- first_col + n_cols - 1
    dat <- raw_dat[, c(1, first_col:end_col)]

    port_start <- match(port_names, dat[, 1])
    for (i in 1:length(which(!is.na(port_start)))) {
      sub <- dat[-c(1:port_start[1]), ] # trim first rows
      sub <- sub[1:(which(is.na(as.numeric(sub[, 1])))[1] - 1), ]
      colnames(sub) <- col_names
      sub$port <- port_names[i]
      sub_long <- tidyr::pivot_longer(sub, cols = which(colnames(sub) %in% c("year", "port") == FALSE))
      if (i == 1) {
        d[[sp]] <- sub_long
      } else {
        d[[sp]] <- rbind(d[[sp]], sub_long)
      }
    }

    # Cleaning up of names etc
    d[[sp]]$name <- stringr::str_remove_all(d[[sp]]$name, "[.]")
    d[[sp]]$name <- stringr::str_replace(d[[sp]]$name, "a/", "")
    d[[sp]]$name <- stringr::str_replace(d[[sp]]$name, "b/", "")
    names(d[[sp]])[1] <- "year"
    d[[sp]]$value <- as.numeric(d[[sp]]$value)
    d[[sp]]$state <- state
    d[[sp]]$indian <- indian
    d[[sp]]$fishery <- fishery
    d[[sp]]$metric <- metric
  }

  if (length(d) > 1) {
    d[[1]]$species <- "Chinook"
    d[[2]]$species <- "coho"
    d <- rbind(d[[1]], d[[2]])
  } else {
    d <- d[[1]]
  }

  # join in labels
  df <- data.frame(
    "indian" = c("", "", "Indian", "", "", "Indian"),
    "fishery" = rep(c("recreational", "commercial", "commercial"), 2),
    "metric" = c(rep("effort", 3), rep("catch", 3)),
    "units" = c("anger trips", "days fished", "deliveries", "numbers", "numbers", "numbers")
  )
  d <- dplyr::left_join(d, df)

  # add ports
  ports <- read.csv("data/ports.csv")
  d <- dplyr::left_join(d, ports)
  return(as.data.frame(d))
}

url <- "https://www.pcouncil.org/documents/2020/06/ocean-salmon-fishery-effort-and-landings-salmon-review-appendix-a-excel-file-format.xlsm"

# CA rec effort
sheet_n <- 4
ca_rec <- parse_data(url_name = url, sheet = sheet_n)
write.csv(ca_rec, "data/pfmc_ca_rec.csv", row.names=FALSE)

# OR rec effort
sheet_n <- 9
or_rec <- parse_data(url_name = url, sheet = sheet_n)
write.csv(or_rec, "data/pfmc_or_rec.csv", row.names=FALSE)

# WA rec effort
sheet_n <- 17
wa_rec <- parse_data(url_name = url, sheet = sheet_n)
wa_rec$units <- "angler trips"
write.csv(wa_rec, "data/pfmc_wa_rec.csv", row.names=FALSE)

# CA commercial troll
sheet_n <- 2
ca_troll <- parse_data(url_name = url, sheet = sheet_n)
ca_troll$units <- "days fished"
write.csv(ca_troll, "data/pfmc_ca_troll_effort.csv", row.names=FALSE)

# OR commercial rroll
sheet_n <- 7
or_troll <- parse_data(url_name = url, sheet = sheet_n)
or_troll$units <- "days fished"
write.csv(or_troll, "data/pfmc_or_troll_effort.csv", row.names=FALSE)

# WA commercial rroll
sheet_n <- 12
wa_troll <- parse_data(url_name = url, sheet = sheet_n)
wa_troll$units <- "days fished"
write.csv(wa_troll, "data/pfmc_wa_troll_effort.csv", row.names=FALSE)

sheet_n <- 14
wa_troll_indian <- parse_data(url_name = url, sheet = sheet_n)
wa_troll_indian$units <- "deliveries"
write.csv(wa_troll_indian, "data/pfmc_wa_troll_indian_effort.csv", row.names=FALSE)

# California recreational landings
sheet_n <- 3
ca_troll <- parse_data(url_name = url, sheet = sheet_n)
write.csv(ca_troll, "data/pfmc_ca_troll_catch.csv", row.names=FALSE)

sheet_n <- 5
ca_rec <- parse_data(url_name = url, sheet = sheet_n)
write.csv(ca_rec, "data/pfmc_ca_rec_catch.csv", row.names=FALSE)

sheet_n <- 8
or_troll <- parse_data(url_name = url, sheet = sheet_n)
write.csv(or_troll, "data/pfmc_or_troll_catch.csv", row.names=FALSE)

sheet_n <- 10
or_rec <- parse_data(url_name = url, sheet = sheet_n)
write.csv(or_rec, "data/pfmc_or_rec_catch.csv", row.names=FALSE)

sheet_n <- 13
wa_troll <- parse_data(url_name = url, sheet = sheet_n)
write.csv(wa_troll, "data/pfmc_wa_troll_catch.csv", row.names=FALSE)

sheet_n <- 18
wa_rec <- parse_data(url_name = url, sheet = sheet_n)
write.csv(wa_rec, "data/pfmc_wa_rec_catch.csv", row.names=FALSE)

sheet_n <- 15 
wa_tribal <- parse_data(url_name = url, sheet = sheet_n)
write.csv(wa_tribal, "data/pfmc_wa_troll_indian_catch.csv", row.names=FALSE)
