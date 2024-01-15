# Install shiny package if not installed
# install.packages("shiny")

# Load shiny library
library(shiny)
library(openxlsx)
library(lubridate)
library(dplyr)
source("./R/getHolidays.R")
source("./R/turnus.R")

wochentage <- c("Monday", "Tuesday", "Wednesday", "Thursday", "Friday")

# Define UI
ui <- fluidPage(
  titlePanel("Notentabelle"),

  # Numeric slider from 1 to 30
  sliderInput("sus", "1. Wie viele SuS?", min = 1, max = 30, value = 25),

  # Selector with numbers 1 to 5
  selectizeInput("number_selector", "2. Turnus, z.B. jeden Montag", choices = wochentage, selected = "", multiple = T, options = list(maxItems = 5, placeholder ="Monday")),

  # Date picker with start and end date. Default für Start nächster Montag
  shinyWidgets::airDatepickerInput("halbjahr", "3. Halbjahr (Start - Ende)",
                                   minDate = floor_date(Sys.Date(), "week") + days(1),
                                   maxDate = Sys.Date()+180,
                                   firstDay =  1,
                                   range = T,
                                   update_on = "close",
                                   placeholder =  c(floor_date(Sys.Date(), "week") + days(1), Sys.Date()+180),
                                   clearButton = T),

  # date pickers with only one date
  shinyWidgets::airDatepickerInput("klassenarbeiten", "4. Klassenarbeitstermin(e)", multiple = T,
                                   firstDay =  1,
                                   disabledDates = c(0,6), minDate = Sys.Date() - 7, maxDate = Sys.Date()+180),

  # filename
  shiny::textInput("filename", "5. Name der Notentabelle", placeholder = "10a_2024_HJ1"),

  shinyWidgets::switchInput(
    inputId = "holiday",
    label = "6. freie Tage",
    value = FALSE,
    labelWidth = "90px"
  ),

  # Download button
  downloadButton("download_btn", "7. Generiere Tabelle")
)

# Define server
server <- function(input, output, session) {


  # update datepicker remove freie tage from klassenarbeiten und only allow ausgewählte wochentage
  observe({
    req(input$number_selector)
    disabled_days <- which(!wochentage %in% input$number_selector)
    shinyWidgets::updateAirDateInput(session = session, "halbjahr", options = list(
      disabledDaysOfWeek = c(0,6,disabled_days),
      update_on = "close",
      minDate = floor_date(Sys.Date(), "week") + days(which(wochentage == input$number_selector[1]))))

    shinyWidgets::updateAirDateInput(session = session, "klassenarbeiten", options = list(
      disabledDaysOfWeek = c(0,6,disabled_days),
      disabledDates = termine()[ausfall()],
      minDate = floor_date(Sys.Date(), "week") + days(which(wochentage == input$number_selector[1]))))
  })

  termine <- reactiveVal(0)
  ausfall <- reactiveVal(0)

  observe({
    req(input$halbjahr[2], input$number_selector)
    schulanfang <- as.character(input$halbjahr[1])   # Halbjahresintervall festlegen
    schulende <- as.character(input$halbjahr[2])

    # Feiertage und Ferien abrufen
    if(year(input$halbjahr[1]) != year(input$halbjahr[2])){
      ferienintervall <- c(getHolidays(year(input$halbjahr[1])),
                           getHolidays(year(input$halbjahr[2]), pause = 5))
    } else {
      ferienintervall <- getHolidays(year(input$halbjahr[1]))
    }

    # turnusgemäße Termine ermitteln
    termine <- turnus2dates(input$number_selector, schulanfang = as.character(input$halbjahr[1]), schulende = as.character(input$halbjahr[2]))


    # Feiertage ermitteln
    ausfall <- sapply(termine, function(x) any(x %within% ferienintervall))

    # reactive values auffüllen
    ausfall(ausfall)
    termine(termine)
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
      klassenarbeiten <- as.character(input$klassenarbeiten)  # Klassenarbeiten festlegen

      # datensatz kreieren
      empty <- data.frame(matrix(0, ncol = length(termine()), nrow = SuS))
      colnames(empty) <- termine()

      # wenn keine Klassenarbeiten gewählt, mache nichts
      if(length(klassenarbeiten) != "character(0)"){
        klassenarbeitsdaten <- which(colnames(empty) %in% klassenarbeiten)
        colnames(empty)[klassenarbeitsdaten] <- paste("Klausur\n", colnames(empty)[klassenarbeitsdaten])
      }


      if(any(colnames(empty) %in% termine()[ausfall()])){
        frei <- which(colnames(empty) %in% termine()[ausfall()])
        colnames(empty)[frei] <- paste("FREI", colnames(empty)[frei])
      }

      # Letzte Spalte
      to <- tail(c(LETTERS, paste0("A", LETTERS), paste0("B", LETTERS), paste0("C", LETTERS))[6:(length(termine())+7)],1)

      # Variablen einführen
      empty$ID = 1:SuS
      empty$`Name` = rep("", SuS)
      empty$Gesamtnote = sprintf('=AVERAGEIF(D%d:E%d, "<>0", D%d:E%d)', 2:(SuS+1), 2:(SuS+1), 2:(SuS+1), 2:(SuS+1))
      empty$muendlich = sprintf(glue::glue('= AVERAGEIFS(F%d:{to}%d, F1:{to}1, "<>*KLAUSUR*", F1:{to}1, "<>*FREI*", F%d:{to}%d, "<>0")'), 2:(SuS+1), 2:(SuS+1), 2:(SuS+1), 2:(SuS+1))
      empty$schriftlich = sprintf(glue::glue('= AVERAGEIFS(F%d:{to}%d, F1:{to}1, "*KLAUSUR*", F%d:{to}%d, "<>0")'), 2:(SuS+1), 2:(SuS+1), 2:(SuS+1), 2:(SuS+1))

      # Datensatz neu-anordnen
      full <- empty %>% select(ID, Name, Gesamtnote, muendlich, schriftlich, everything()) %>%
        mutate(across(contains("FREI"), ~ ""))

      if(!input$holiday){
        full <- full %>% select(!contains("FREI"))
      }

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

      idx0 <- which(colnames(full) %in% termine())
      conditionalFormatting(wb, sheet =  "Noten", cols = idx0, rows = 1:(SuS+1), style = posStyle, rule = "",
                            type = "contains")

      idx <- which(grepl("FREI", colnames(full)))
      for(i in idx){
        conditionalFormatting(wb, sheet =  "Noten", cols = i, rows = 1:(SuS+1), style = neutralStyle, rule = "=TRUE",
                              type = "expression")
      }

      idx2 <- which(colnames(full) %in% colnames(empty)[klassenarbeitsdaten])
      for(i in idx2){
        conditionalFormatting(wb, sheet =  "Noten", cols = i, rows = 1:(SuS+1), style = negStyle, rule = "",
                              type = "contains")
      }

      # notenspiegel
      notenspiegel <- data.frame(Note = 1:6, Anzahl = sprintf('=COUNTIFS(Noten!C%d:Noten!C%d, ">%d,5", Noten!C%d:Noten!C%d, "<=%d,5")', 2, SuS+1, 0:5, 2, SuS+1, 1:6))
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
