# Load shiny libraries
library(shiny)
library(openxlsx)
library(lubridate)
library(dplyr)
library(shinyjs)
#library(bslib)
source("getHolidays.R")
source("turnus.R")
wochentage <- c("Monday", "Tuesday", "Wednesday", "Thursday", "Friday")

# Define UI
ui <- fluidPage(
  style = 'margin: 10px 15px',
  #theme = bs_theme(preset = "shiny"),
  shinyjs::useShinyjs(),
  titlePanel("Notentabelle"),

  HTML('Schritte bitte nacheinander ausfüllen.<br> Im Zweifel neuladen.'),
  br(),
  br(),

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
                                   placeholder =  c(floor_date(Sys.Date(), "week") + days(1), Sys.Date()+182),
                                   clearButton = T),

  # date pickers with only one date
  shinyWidgets::airDatepickerInput("klassenarbeiten", "4. Klassenarbeitstermin(e)", multiple = T,
                                   firstDay =  1,
                                   clearButton = T,
                                   disabledDates = c(0,6), minDate = Sys.Date() - 7, maxDate = Sys.Date()+182),

  # filename
  shiny::textInput("filename", "5. Name der Notentabelle", placeholder = "10a_2024_HJ1"),

  shinyWidgets::switchInput(
    inputId = "holiday",
    label = "6. freie Tage",
    value = FALSE,
    labelWidth = "90px"
  ),

  # Download button
  downloadButton("download_btn", "7. Generiere Tabelle"),

  hr(),
  HTML("Referenz: <a href='https://ferien-api.de/'>Feriendaten</a> und <a href='https://feiertage-api.de/'>Feiertagsdaten</a>")
)

# Define server
server <- function(input, output, session) {
  #bslib::bs_themer()
  # disable turnus-selector, wenn halbjahr nicht gesetzt
  observe({
    req(input$number_selector)
    if(is.null(input$halbjahr[2])){
      shinyjs::enable("number_selector")
    }
  })

  # update datepicker only allow ausgewählte wochentage
  observeEvent(input$number_selector ,{
    req(input$number_selector)
    disabled_days <- which(!wochentage %in% input$number_selector)
    shinyWidgets::updateAirDateInput(session = session, "halbjahr", options = list(
      disabledDaysOfWeek = c(0,6,disabled_days),
      update_on = "close",
      minDate = floor_date(Sys.Date(), "week") - 7*8 + days(which(wochentage == input$number_selector[1]))))
  })

  # reactive Values setzen
  termine <- reactiveVal(0)
  ausfall <- reactiveVal(0)
  api <- reactiveVal(0)           # api results zwischenspeichern
  halbjahrend <- reactiveVal(0)   # alte Einträge zum HJ wieder verwenden, wenn HJ-Zeitraum nachträglich verkleinert werden
  halbjahranf <- reactiveVal(0)

  # remove freie tage zusätzlich from klassenarbeiten grenze Zeitbereich auf ausgewähltes Halbjahr ein
  observeEvent(input$halbjahr[2], {
    req(input$number_selector, input$halbjahr[2])
    shinyjs::disable("number_selector")

  # wenn ausgewählter Wochentag nicht dem ersten Turnustag entspricht, wird er automatisch angepasst
    if(input$number_selector[1] != weekdays(ymd(input$halbjahr[1]))){
      ersterTurnustag <- which(input$number_selector[1] == wochentage)
      schulanfangneu <- lubridate::floor_date(ymd(input$halbjahr[1]), "week") + days(ersterTurnustag)
      shinyWidgets::updateAirDateInput(session = session, "halbjahr",value = c(schulanfangneu, input$halbjahr[2]))
      shinyjs::alert(glue::glue("Erster Turnuswochentag und erster Unterrichtstag stimmen nicht überein, setze {input$number_selector[1]} aus der Woche als Starttag ({schulanfangneu})."))
    } else {
      schulanfangneu <- input$halbjahr[1]
    }

    # Feiertage und Ferien abrufen (nur wenn dies nicht bereits geschehen)
    if(api()[1] == 0 || year(halbjahrend()) < year(input$halbjahr[2]) || year(halbjahranf()) > year(input$halbjahr[1])){
      if(year(input$halbjahr[1]) != year(input$halbjahr[2])){
        withProgress(message = 'Mache API-Abfragen', value = 0.0, {
          disable("klassenarbeiten")
          ferienintervall <- c(getHolidays_schule(pause = 5, progress = .333),
                               getHolidays_feiertag(year(input$halbjahr[1]), pause = 1, progress = .333),
                               getHolidays_feiertag(year(input$halbjahr[2]), pause = 5, progress = .333))
          api(ferienintervall)
          halbjahrend(input$halbjahr[2])
          halbjahranf(input$halbjahr[1])
          enable("klassenarbeiten")
        })
      } else {
        withProgress(message = 'Mache API-Abfrage', value = 0.0, {
          disable("klassenarbeiten")
          ferienintervall <- c(getHolidays_schule(pause = 5, progress = .5),
                               getHolidays_feiertag(year(input$halbjahr[1]), pause = 1, progress = .5))
          api(ferienintervall)
          halbjahrend(input$halbjahr[2])
          halbjahranf(input$halbjahr[1])
          enable("klassenarbeiten")
        })
      }
    } else {
      ferienintervall <- api()    # sonst verwende bereits geladene Termine
    }

    # turnusgemäße Termine ermitteln
    termine <- turnus2dates(input$number_selector, schulanfang = as.character(schulanfangneu), schulende = as.character(input$halbjahr[2]))

    # Feiertage ermitteln
    ausfall <- sapply(termine, function(x) any(x %within% ferienintervall))

    # reactive values auffüllen
    ausfall(ausfall)
    termine(format(termine, "%a %d-%m-%y"))

    # update slider
    disabled_days <- which(!wochentage %in% input$number_selector)
    shinyWidgets::updateAirDateInput(session = session, "klassenarbeiten", options = list(
      disabledDaysOfWeek = c(0,6,disabled_days),
      disabledDates = format(dmy(termine()[ausfall()]), "%Y-%m-%d"),
      minDate = schulanfangneu, maxDate = input$halbjahr[2]))
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
      withProgress(message = 'Erstelle Datei', value = 0.0, {
      SuS <- as.numeric(input$sus)   # Specify the number SuS
      klassenarbeiten <- as.character(input$klassenarbeiten)  # Klassenarbeiten festlegen
      # datensatz kreieren
      empty <- data.frame(matrix(ncol = length(termine()), nrow = SuS))
      colnames(empty) <- termine()

      # wenn keine Klassenarbeiten gewählt, mache nichts
      if(length(klassenarbeiten) != "character(0)"){
        klassenarbeiten <- format(ymd(klassenarbeiten), "%a %d-%m-%y")
        klassenarbeitsdaten <- which(colnames(empty) %in% klassenarbeiten)
        colnames(empty)[klassenarbeitsdaten] <- paste("Klausur\n", colnames(empty)[klassenarbeitsdaten])
      }

      if(any(colnames(empty) %in% termine()[ausfall()])){
        frei <- which(colnames(empty) %in% termine()[ausfall()])
        colnames(empty)[frei] <- paste("FREI", colnames(empty)[frei])
      }

      if(!input$holiday){
        empty <- empty %>% select(!contains("FREI"))
      }

      # Letzte Spalte
      to_all <- as.vector(sapply(c("", LETTERS[1:10]), \(x) paste0(x, LETTERS)))[6:(ncol(empty)+5)]
      to <- tail(to_all,1)

      # Variablen einführen
      empty$ID = 1:SuS
      empty$`Nachname, Vorname` = rep("", SuS)
      empty$Gesamtnote = sprintf('=IFERROR(AVERAGEIF(D%d:E%d, "<>0", D%d:E%d), "")', 2:(SuS+1), 2:(SuS+1), 2:(SuS+1), 2:(SuS+1))
      empty$muendlich = sprintf(glue::glue('= IFERROR(AVERAGEIFS(F%d:{to}%d, F1:{to}1, "<>*KLAUSUR*", F1:{to}1, "<>*FREI*", F%d:{to}%d, "<>0"), "")'), 2:(SuS+1), 2:(SuS+1), 2:(SuS+1), 2:(SuS+1))
      empty$schriftlich = sprintf(glue::glue('= IFERROR(AVERAGEIFS(F%d:{to}%d, F1:{to}1, "*KLAUSUR*", F%d:{to}%d, "<>0"), "")'), 2:(SuS+1), 2:(SuS+1), 2:(SuS+1), 2:(SuS+1))

      # Datensatz neu-anordnen
      full <- empty %>% select(ID, `Nachname, Vorname`, Gesamtnote, muendlich, schriftlich, everything())

      # Formelspalten als solche klassifizieren
      for(i in c(3:5)){
        class(full[,i]) <- c(class(full[,i]), "formula")
      }

      ## Notebook erstellen
      wb <- createWorkbook()
      addWorksheet(wb, "Noten")
      freezePane(wb, "Noten", firstActiveCol = 6)
      writeData(wb, "Noten", x = full)
      setColWidths(wb, "Noten", cols = c(1:2, 6:ncol(full)), widths = "auto")
      setRowHeights(wb, "Noten", rows = c(1), heights = c(40))

      # center headers
      headerstyle <- createStyle(halign = "center", valign = "center")
      addStyle(wb, sheet = "Noten", headerstyle, rows = 1, cols = 1:ncol(full))

      # füge borderColor hinzu + jeweils für body and headers
      bodyStyle <- createStyle(fgFill = 'grey95', border = "TopBottomLeftRight", borderStyle = 'thin', borderColour = 'grey65', halign = "center", valign = "center")
      addStyle(wb, sheet = "Noten", bodyStyle, rows = 1:(SuS+1), cols = 1:5, gridExpand = TRUE)
      bodyStyle2 <- createStyle(fgFill = 'grey90', border = c("top","bottom","left","right"), borderStyle = c('thin','thick','thin','thin'), borderColour = 'grey45', halign = "center", valign = "center")
      addStyle(wb, sheet = "Noten", bodyStyle2, rows = 1, cols = 1:5, gridExpand = TRUE)

      # Style festlegen
      neutralStyle <- createStyle(bgFill = "grey", borderColour = 'grey65')
      posStyle <- createStyle(bgFill = "#c6d7ef", border = "TopBottomLeftRight", borderStyle = 'thin', borderColour = 'grey65')
      negStyle <- createStyle(bgFill = "#EFDEC6", border = "TopBottomLeftRight", borderStyle = 'thin', borderColour = 'grey65')

      # add border to header
      neutralStyleH1 <- createStyle(bgFill = "grey50", borderColour = 'grey45', border = "Bottom", borderStyle = 'thick')
      posStyleH1 <- createStyle(bgFill = "#8FB8FF", border = "Bottom", borderStyle = 'thick', borderColour = 'grey45')
      negStyleH1 <- createStyle(bgFill = "#FFD68F", border = "Bottom", borderStyle = 'thick', borderColour = 'grey45')

      # header regeln für gesamte reihe
      idx0 <- 6:ncol(full)
      conditionalFormatting(wb, sheet =  "Noten", cols = idx0, rows = 1, style = posStyleH1, rule = "Klausur",
                            type = "notContains")
      conditionalFormatting(wb, sheet =  "Noten", cols = idx0, rows = 1, style = posStyleH1, rule = "Frei",
                            type = "notContains")
      conditionalFormatting(wb, sheet =  "Noten", cols = idx0, rows = 1, style = negStyleH1, rule = "Klausur",
                            type = "contains")
      conditionalFormatting(wb, sheet =  "Noten", cols = idx0, rows = 1, style = neutralStyleH1, rule = "Frei",
                            type = "contains")

      # body regeln für einzelne Spalten
      all <- as.vector(sapply(c("", LETTERS[1:10]), \(x) paste0(x, LETTERS)))
      for (i in idx0) {
        conditionalFormatting(wb, sheet =  "Noten", cols = i, rows = 2:(SuS+1), style = posStyle, rule = glue::glue('NOT(OR(ISNUMBER(SEARCH("Klausur", ${all[i]}$1)), ISNUMBER(SEARCH("Frei", ${all[i]}$1))))'),
                              type = "expression")
        conditionalFormatting(wb, sheet =  "Noten", cols = i, rows = 2:(SuS+1), style = negStyle, rule = glue::glue('ISNUMBER(SEARCH("Klausur", ${all[i]}$1))'),
                              type = "expression")
        conditionalFormatting(wb, sheet =  "Noten", cols = i, rows = 2:(SuS+1), style = neutralStyle, rule = glue::glue('ISNUMBER(SEARCH("FREI", ${all[i]}$1))'),
                              type = "expression")
        incProgress(amount = 1/length(idx0))
      }
      # body regeln für alle spalten ab spalte 6 incl.
      conditionalFormatting(wb, sheet =  "Noten", cols = idx0, rows = 2:(SuS+1), style = createStyle(bgFill = "white"), rule = ">6",
                            type = "expression")
      conditionalFormatting(wb, sheet =  "Noten", cols = idx0, rows = 1, style = createStyle(bgFill = "white", border = "Bottom", borderStyle = 'thin'), rule = "-",
                            type = "notContains")

      # notenspiegel
      notenspiegel <- data.frame(Note = 1:6, Anzahl = sprintf('=COUNTIFS(Noten!C%d:Noten!C%d, ">%d,5", Noten!C%d:Noten!C%d, "<=%d,5")', 2, SuS+1, 0:5, 2, SuS+1, 1:6))
      class(notenspiegel$Anzahl) <- c(class(notenspiegel$Anzahl), "formula")
      addWorksheet(wb, "Notenspiegel")
      writeData(wb, "Notenspiegel", x = notenspiegel)
      # speichern
      saveWorkbook(wb, file, overwrite = TRUE)
      }
  )
    })
}

# Run the app
shinyApp(ui, server)
