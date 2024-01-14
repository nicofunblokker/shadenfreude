# Install shiny package if not installed
# install.packages("shiny")

# Load shiny library
library(shiny)
library(openxlsx)
library(lubridate)
library(shinyjs)
library(dplyr)
source("./R/getHolidays.R")

wochentage <- c("Monday", "Tuesday", "Wednesday", "Thursday", "Friday")

# Define UI
ui <- fluidPage(
  titlePanel("Notentabelle"),

  # Numeric slider from 1 to 30
  sliderInput("sus", "Wie viele SuS?", min = 1, max = 30, value = 25),

  # Selector with numbers 1 to 5
  selectizeInput("number_selector", "Turnus, z.B. jeden Montag (max. 2 Tage)", choices = wochentage, selected = "", multiple = T, options = list(maxItems = 2, placeholder ="Monday")),

  # Date picker with start and end date. Default für Start nächster Montag
  shinyWidgets::airDatepickerInput("halbjahr", "Halbjahr (Start - Ende)",
                                   minDate = floor_date(Sys.Date(), "week") + days(1),
                                   maxDate = Sys.Date()+180,
                                   firstDay =  1,
                                   range = T,
                                   placeholder =  c(floor_date(Sys.Date(), "week") + days(1), Sys.Date()+180),
                                   clearButton = T),

  # tooltip hinzufügen
  tags$script(HTML('
    $(document).ready(function(){
      $("#halbjahr").tooltip({title: "Der gewählte Starttag des Halbjahres sollte dem ersten Unterrichtstag in der Klasse entsprechen. Wenn dies z.B. Donnerstag ist, dann sollte bei Turnus auch zuerst Donnerstag eingetragen werden (gefolgt von z.B. Montag).", placement: "top", trigger: "hover", container: "body"});
    });
  ')),

  # date pickers with only one date
  shinyWidgets::airDatepickerInput("klassenarbeiten", "Klassenarbeitstermin(e)", multiple = T,
                                   disabledDates = c(0,6), minDate = Sys.Date() - 7, maxDate = Sys.Date()+180),

  # filename
  shiny::textInput("filename", "Name der Notentabelle", placeholder = "10a_HJ1"),

  # Download button
  downloadButton("download_btn", "Generiere Tabelle")
)

# Define server
server <- function(input, output, session) {

  observe({
    req(input$number_selector)
    disabled_days <- which(!wochentage %in% input$number_selector)
    shinyWidgets::updateAirDateInput(session = session, "halbjahr", options = list(
      disabledDaysOfWeek = c(0,6,disabled_days),
      minDate = floor_date(Sys.Date(), "week") + days(which(wochentage == input$number_selector[1]))))

    shinyWidgets::updateAirDateInput(session = session, "klassenarbeiten", options = list(
      disabledDaysOfWeek = c(0,6,disabled_days)))
  })

  # Download button logic
  output$download_btn <- downloadHandler(
    filename = function() {
      if(input$filename != ""){
        glue::glue("Notentabelle_{input$filename}.xlsx")
      } else {
        glue::glue("Notentabelle_{Sys.Date()}.xlsx")
      }
    },
    content = function(file) {
      SuS <- as.numeric(input$sus)   # Specify the number SuS
      schulanfang <- as.character(input$halbjahr[1])   # Halbjahresintervall festlegen
      schulende <- as.character(input$halbjahr[2])
      klassenarbeiten <- as.character(input$klassenarbeiten)  # Klassenarbeiten festlegen

      if(!is.null(input$number_selector)){
        tageidx <- which(wochentage %in% input$number_selector) # wie viele tage zwischen Unterrichtstunden (0 wenn nur 1x pro Woche)
        tage <- max(tageidx) - min(tageidx)
      } else {
        tage <- 0                                            # ebenfalls 0, wenn keine Angabe
      }

      # consider edgecase: Schulbeginn an Donnerstag, Turnus aber Montag - Donnerstag. Dann nicht 3 Tage vor, sondern zurück
      if(length(input$number_selector) > 1){
        if(grep(input$number_selector[1], wochentage) > grep(input$number_selector[2], wochentage)) {
          tage <- tage*-1
        } else {
          tage
        }
      }

      # Feiertage und Ferien abrufen
      #ferienintervall <- getHolidays(year(input$halbjahr[1]))

      # Sequenz festlegen
      a <- seq(ymd(schulanfang),ymd(schulende), by = '1 week')
      b <- seq(ymd(schulanfang)+as.numeric(tage) ,ymd(schulende), by = '1 week')

      # consider edgecase: Schulbeginn an Donnerstag, Turnus aber Montag - Donnerstag. Dann nicht 3 Tage vor, sondern zurück.
      # Allerdings OHNE den Montag in der ausgewählten Woche -> zufrüh
      zuf <- which(b < a[1])
      if(length(zuf) > 0){
        termine <- unique(c(a,b[-zuf])) |> sort()
      } else {
        termine <- unique(c(a,b)) |> sort()
      }

      # Feiertage ermitteln
      ausfall <- sapply(termine, function(x) any(x %within% ferienintervall))

      # datensatz kreieren
      empty <- data.frame(matrix(0, ncol = length(termine), nrow = SuS))
      colnames(empty) <- termine

      # wenn keine Klassenarbeiten gewählt, mache nichts
      if(length(klassenarbeiten) != "character(0)"){
        klassenarbeitsdaten <- which(colnames(empty) %in% klassenarbeiten)
        colnames(empty)[klassenarbeitsdaten] <- paste("Klausur\n", colnames(empty)[klassenarbeitsdaten])
      }

      # Letzte Spalte
      to <- tail(c(LETTERS, paste0("A", LETTERS), paste0("B", LETTERS), paste0("C", LETTERS))[6:(length(termine)+7)],1)

      # Variablen einführen
      empty$ID = 1:SuS
      empty$`Name` = rep("", SuS)
      empty$Gesamtnote = sprintf('=AVERAGEIF(D%d:E%d, "<>0", D%d:E%d)', 2:(SuS+1), 2:(SuS+1), 2:(SuS+1), 2:(SuS+1))
      empty$muendlich = sprintf(glue::glue('= AVERAGEIFS(F%d:{to}%d, F1:{to}1, "<>*KLAUSUR*", F%d:{to}%d, "<>0")'), 2:(SuS+1), 2:(SuS+1), 2:(SuS+1), 2:(SuS+1))
      empty$schriftlich = sprintf(glue::glue('= AVERAGEIFS(F%d:{to}%d, F1:{to}1, "*KLAUSUR*", F%d:{to}%d, "<>0")'), 2:(SuS+1), 2:(SuS+1), 2:(SuS+1), 2:(SuS+1))

      # Datensatz neu-anordnen
      full <- empty %>% select(ID, Name, Gesamtnote, muendlich, schriftlich, everything())

      # Formelspalten als solche klassifizieren
      for(i in c(3:5)){
        class(full[,i]) <- c(class(full[,i]), "formula")
      }

      ## Notebook erstellen
      wb <- createWorkbook()
      addWorksheet(wb, "Noten")
      writeData(wb, "Noten", x = full)

      # Style festlegen
      neutralStyle <- createStyle(bgFill = "grey")
      posStyle <- createStyle(bgFill = "#C6EFCE")
      negStyle <- createStyle(bgFill = "#FFC7CE")

      idx0 <- which(colnames(full) %in% termine)
      conditionalFormatting(wb, sheet =  "Noten", cols = idx0, rows = 1:(SuS+1), style = posStyle, rule = "",
                            type = "contains")

      idx <- which(colnames(full) %in% termine[ausfall])
      for(i in idx){
        conditionalFormatting(wb, sheet =  "Noten", cols = i, rows = 1:(SuS+1), style = neutralStyle, rule = "",
                              type = "contains")
      }

      idx2 <- which(colnames(full) %in% colnames(empty)[klassenarbeitsdaten])
      for(i in idx2){
        conditionalFormatting(wb, sheet =  "Noten", cols = i, rows = 1:(SuS+1), style = negStyle, rule = "",
                              type = "contains")
      }

      # notenspiegel
      notenspiegel <- data.frame(Note = 1:6, Anzahl = sprintf("=COUNTIF(Noten!C%d:Noten!C%d, %d)", 2, SuS+1, 1:6))
      class(notenspiegel$Anzahl) <- c(class(notenspiegel$Anzahl), "formula")
      addWorksheet(wb, "Notenspiegel")
      writeData(wb, "Notenspiegel", x = notenspiegel)

      # speichern
      saveWorkbook(wb, file, overwrite = TRUE)
    }
  )
}

# Run the app
shinyApp(ui, server)
