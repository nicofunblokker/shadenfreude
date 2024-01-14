
getHolidays <- function(jahr = lubridate::year(lubridate::today()), pause = 0){
  Sys.sleep(pause)
  url <- paste0("https://ferien-api.de/api/v1/holidays/NI/", jahr)
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





