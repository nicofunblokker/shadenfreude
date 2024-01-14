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
  sliderInput("sus", "SuS", min = 1, max = 30, value = 25),

  # Selector with numbers 1 to 5
  selectInput("number_selector", "Turnus (wähle 1-2)", choices = wochentage, selected = wochentage[c(1, 4)], multiple = T),

  # Date picker with start and end date
  dateRangeInput("halbjahr", "Halbjahr (Start - Ende)", start = floor_date(Sys.Date(), "week") + days(1), end = Sys.Date()+180, weekstart = 1),

  tags$script(HTML('
    $(document).ready(function(){
      $("#halbjahr").tooltip({title: "Wähle den ersten Unterrichtstag. Wochentag sollte mit Turnus übereinstimmen! Z.B., wenn erster Unterrichtstag an einem Donnerstag stattfindet und dann erst wieder am Montag, wähle Turnus: Thursday, Monday und Start: Thursday 11.01.2024", placement: "top", trigger: "hover", container: "body"});
    });
  ')),

  # date pickers with only one date
  shinyWidgets::airDatepickerInput("klassenarbeiten", "Klassenarbeit(en)", multiple = T,
                                   disabledDates = c(0,6), minDate = Sys.Date() - 7, maxDate = Sys.Date()+180),

  # Download button
  downloadButton("download_btn", "Generiere Tabelle")
)

# Define server
server <- function(input, output) {

  # Download button logic (you can replace this with your own data)
  output$download_btn <- downloadHandler(
    filename = function() {
      paste("Notentabelle_", Sys.Date(), ".xlsx", sep = "")
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

      # consider edgecase where erster später Wochentagtag, zb Donnerstag und nicht Montag
      if(length(input$number_selector) > 1){
        if(grep(input$number_selector[1], wochentage) > grep(input$number_selector[2], wochentage)) {
          tage <- tage*-1
        } else {
          tage
        }
      }

      # Feiertage und Ferien abrufen
      ferienintervall <- getHolidays(year(input$halbjahr[1]))

      # Sequenz festlegen
      a <- seq(ymd(schulanfang),ymd(schulende), by = '1 week')
      b <- seq(ymd(schulanfang)+as.numeric(tage) ,ymd(schulende), by = '1 week')

      # consider edgecase where erster Tag zb Donnerstag: entferne Montage vor dem gewählten Donnerstag
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

      # wenn keine Klassenarbeiten gewählt
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

      # Datensatz reshufflen
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
