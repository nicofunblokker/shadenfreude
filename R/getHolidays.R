# ferien
getHolidays_schule <- function(pause = 5, progress){
  incProgress(amount = progress)
  Sys.sleep(pause)

  #url <- paste0("https://ferien-api.de/api/v1/holidays/NI/", jahr) # bei Jahresübergängen fehlen ggfs. Daten
  url <- paste0("https://ferien-api.de/api/v1/holidays/NI/")        # --> daher alles
  headers <- c("accept" = "application/json")

  response <- httr::GET(url, headers = headers)
  json_response <- httr::content(response, "text", encoding = "UTF-8")

  # Parse the JSON response
  ferien <- jsonlite::fromJSON(json_response)

  # Extract start and end dates as vectors
  start_dates <- as.Date(ferien$start)
  end_dates <- as.Date(ferien$end)

  ferienintervall <- purrr::map2_vec(start_dates, end_dates, lubridate::interval)
  return(ferienintervall)
}


# feiertage
getHolidays_feiertag <- function(jahr = lubridate::year(lubridate::today()), pause = 5, progress){
  Sys.sleep(pause)
  incProgress(amount = progress)
  url <- paste0("https://feiertage-api.de/api/?jahr=", jahr, "&nur_land=NI")
  headers <- c("accept" = "application/json")

  response <- httr::GET(url, headers = headers)
  json_response <- httr::content(response, "text", encoding = "UTF-8")
  feiertage <- jsonlite::fromJSON(json_response, flatten = T)

  feiertage <- purrr::map_vec(feiertage, ~lubridate::as_date(.x$datum[1])) |> unname()
  feiertage <- purrr::map2_vec(feiertage, feiertage, lubridate::interval)

  return(feiertage)
}





