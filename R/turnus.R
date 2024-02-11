
turnus2dates <- function(turnusTage, schulanfang, schulende){
  #browser()
  wochentage <- c("Monday", "Tuesday", "Wednesday", "Thursday", "Friday")
  turnus_rhythm <- sapply(turnusTage, \(x) grep(x, wochentage))
  if(length(turnus_rhythm) > 1){
    turnus_rhythm2 <- c(0, sapply(turnus_rhythm[-1], \(x) diff(c(turnus_rhythm[1], x))))
  } else {
    turnus_rhythm2 <- 0
  }
  unterrichtsdaten <- purrr::map(turnus_rhythm2, \(x) seq(ymd(schulanfang)+x,ymd(schulende), by = '1 week'))
  unterrichtsdaten_uniq <- unterrichtsdaten |> purrr::list_c(ptype = vctrs::new_date()) |> unique()  #https://stackoverflow.com/questions/15659783/why-does-unlist-turn-date-types-into-numeric

  if(any(unterrichtsdaten_uniq < unterrichtsdaten_uniq[1])){
    unterrichtsdaten_uniq <- unterrichtsdaten_uniq[-which(unterrichtsdaten_uniq < unterrichtsdaten_uniq[1])]
  }
  return(sort(unterrichtsdaten_uniq))
}
