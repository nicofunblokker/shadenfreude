generateSchuljahr <- function(ferien){
  sommer <- which(grepl("Sommer", ferien$name, ignore.case = T))
  jahr <- ferien[sommer[1]:sommer[2],]
  halbjahr <- which(grepl("Halb", jahr$name, ignore.case = T))
  hj1_start <- ymd(jahr$endDate[1]) + lubridate::days(1)
  hj1_ende <- ymd(jahr$startDate[halbjahr]) - lubridate::days(1)

  hj2_start <- ymd(jahr$endDate[halbjahr]) + lubridate::days(1)
  hj2_ende <- ymd(tail(jahr$startDate, 1)) - lubridate::days(1)

  # c(interval(hj1_start, hj1_ende),
  #      interval(hj2_start, hj2_ende))
  list('schuljahrenden' = c(hj1_ende, hj2_ende),
       'schuljahranfaenge' = c(hj1_start, hj2_start))
}

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

  # sets boundaries
  boundaries <- generateSchuljahr(ferien)


  Sys.sleep(1)
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
  return(list(c(ferienintervall, feiertagintervall), boundaries))
}




