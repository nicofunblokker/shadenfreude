
getHolidays <- function(start, end, pause = 5, progress){
  incProgress(amount = progress)
  Sys.sleep(pause)

  #ferien
  urlFerien <- glue::glue("https://openholidaysapi.org/SchoolHolidays?countryIsoCode=DE&languageIsoCode=DE&validFrom={start}&validTo={end}&subdivisionCode=DE-NI")
  headers <- c("accept" = "application/json")
  response <- httr::GET(urlFerien, headers = headers)
  json_response <- httr::content(response, "text", encoding = "UTF-8")
  ferien <- jsonlite::fromJSON(json_response)
  start_dates <- as.Date(ferien$startDate)
  end_dates <- as.Date(ferien$endDate)
  ferienintervall <- purrr::map2_vec(start_dates, end_dates, lubridate::interval)


  Sys.sleep(pause-4)
  incProgress(amount = progress)
  # feiertage
  urlFeiertag <- glue::glue("https://openholidaysapi.org/PublicHolidays?countryIsoCode=DE&languageIsoCode=DE&validFrom={start}&validTo={end}&subdivisionCode=DE-NI")
  headers <- c("accept" = "application/json")
  response <- httr::GET(urlFeiertag, headers = headers)
  json_response <- httr::content(response, "text", encoding = "UTF-8")
  feiertag <- jsonlite::fromJSON(json_response)
  start_dates <- as.Date(feiertag$startDate)
  end_dates <- as.Date(feiertag$endDate)
  feiertagintervall <- purrr::map2_vec(start_dates, end_dates, lubridate::interval)
  return(c(ferienintervall, feiertagintervall))
}




