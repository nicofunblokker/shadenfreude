# Install shiny package if not installed
# install.packages("shiny")

# Load shiny library
library(shiny)
library(openxlsx)
library(lubridate)
library(dplyr)
source("./R/getHolidays.R")

# Define UI
ui <- fluidPage(
  titlePanel("Notentabelle"),

  # Numeric slider from 1 to 30
  sliderInput("sus", "SuS", min = 1, max = 30, value = 25),

  # Date picker with start and end date
  dateRangeInput("halbjahr", "Halbjahr (Start - Ende)", start = Sys.Date() - 7, end = Sys.Date()+180),

  # Two date pickers with only one date
  dateInput("klassenarbeit1", "Klassenarbeit 1", value = Sys.Date()+30),
  dateInput("klassenarbeit2", "Klassenarbeit 2", value = Sys.Date()+60),

  # Selector with numbers 1 to 5
  selectInput("number_selector", "Turnus", choices = 0:4, selected = 2),

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
      klassenarbeiten <- c(as.character(input$klassenarbeit1), as.character(input$klassenarbeit2))  # Klassenarbeiten festlegen
      tage <- input$number_selector # wie viele tage zwischen Unterrichtstunden (0 wenn nur 1x pro Woche)

      # Feiertage und Ferien abrufen
      ferienintervall <- getHolidays(year(input$halbjahr[1]))

      # Sequenz festlegen
      a <- seq(ymd(schulanfang),ymd(schulende), by = '1 week')
      b <- seq(ymd(schulanfang)+as.numeric(tage) ,ymd(schulende), by = '1 week')
      termine <- unique(c(a,b)) |> sort()

      # Feiertage ermitteln
      ausfall <- sapply(termine, function(x) any(x %within% ferienintervall))

      # datensatz kreieren
      empty <- data.frame(matrix(0, ncol = length(termine), nrow = SuS))
      colnames(empty) <- termine
      klassenarbeitsdaten <- which(colnames(empty) %in% klassenarbeiten)
      colnames(empty)[klassenarbeitsdaten] <- paste("Klausur\n", colnames(empty)[klassenarbeitsdaten])

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
