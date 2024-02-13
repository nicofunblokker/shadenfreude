# Load shiny libraries
library(shiny)
library(openxlsx)
library(lubridate)
library(dplyr)
library(shinyjs)
library(bslib)
source("documentationSheet.R")
source("getHolidays.R")
wochentage <- c("Monday", "Tuesday", "Wednesday", "Thursday", "Friday")
names(wochentage) <- c("Mon", "Tue", "Wed", "Thu", "Fri")

# prepare choice for halbjahr (cut-off for display is generally july&january)
jahr <- today()
halbjahr1 <- floor_date(jahr, 'halfyear') |> year()
halbjahr1 <- as.numeric(gsub("^\\d{2}", "", halbjahr1))
halbjahr2 <- ceiling_date(jahr, 'halfyear') |> year()
halbjahr2 <- as.numeric(gsub("^\\d{2}", "", halbjahr2))
choices <- 1:2
names(choices) <- c(glue::glue('1. (Sommer {halbjahr1}/{halbjahr1+1})'), glue::glue('2. (Winter {halbjahr2-1}/{halbjahr2})'))

# Define UI
ui <- page_fluid(
  theme = bs_theme(bootswatch = "flatly", primary = "#3498db", secondary = "#2c3e50"),
  div(style = "float:right; padding-right: 30px;",
      input_dark_mode(id = "dark_mode", mode = "light")
  ),
  style = 'margin: 10px 15px',
  shinyjs::useShinyjs(),
  titlePanel("Notentabelle v2"),
  HTML('Schritte bitte nacheinander ausfüllen.<br> Im Zweifel neuladen.'),
  br(),
  br(),

  # Numeric slider from 1 to 30
  sliderInput("sus", "1. Anzahl SuS", min = 1, max = 30, value = 25),

  # Selector with weekdays
   shinyWidgets::checkboxGroupButtons(
    inputId = "turnus",
    label = "2. Turnus, z.B. jeden Montag",
    choices = wochentage,
    status = "primary",
    checkIcon = list(
      yes = icon("ok",
                 lib = "glyphicon"),
      no = icon("remove",
                lib = "glyphicon")),
    size = "sm",
    individual = FALSE),


  # select halbjahr
  shinyWidgets::radioGroupButtons(
    inputId = "halbjahr",
    label = "3. Halbjahr",
    choices = sort(choices, decreasing = halbjahr1 > (halbjahr2-1)),
    selected = character(0),
    #inline = T
  ),

  # date pickers with only one date
  shinyWidgets::airDatepickerInput("klassenarbeiten", "4. Klassenarbeitstermin(e)", multiple = T,
                                   firstDay =  1,
                                   clearButton = T,
                                   addon = "none",
                                   readonly = T,
                                   position = 'bottom right',
                                   disabledDates = c(0,6), minDate = Sys.Date() - 7, maxDate = Sys.Date()+182),

  # filename
  shiny::textInput("filename", "5. Name der Notentabelle", placeholder = "10a_2024_HJ1"),

  shinyWidgets::switchInput(
    inputId = "holiday",
    label = "6. freie Tage anzeigen",
    value = FALSE,
    labelWidth = "220px"
  ),

  shinyWidgets::switchInput(
    inputId = "rotate",
    label = "7. Spaltennamen rotieren",
    value = FALSE,
    labelWidth = "220px"
  ),
  # Download button
  downloadButton("download_btn", "8. Generiere Tabelle"),

  hr(),
  HTML("Referenz: <a href='https://www.openholidaysapi.org/en/'>Ferien- und Feiertagsdaten</a> unter <a href='https://raw.githubusercontent.com/openpotato/openholidaysapi.data/main/LICENSE'>dieser Lizenz</a>.")
)

server <- function(input, output, session) {
  disable("halbjahr")
  disable("klassenarbeiten")
  disable("download_btn")
  # reactive Values setzen
  termine <- reactiveVal(0)
  ausfall <- reactiveVal(0)
  api <- reactiveVal(0)           # api results zwischenspeichern
  halbjahrend <- reactiveVal(0)   # alte Einträge zum HJ wieder verwenden, wenn HJ-Zeitraum nachträglich verkleinert werden
  halbjahranf <- reactiveVal(0)

  observe({
    if(!is.null(input$klassenarbeiten)){
      disable("halbjahr")
      disable("turnus")
    } else {
      enable("halbjahr")
      enable("turnus")
    }

    if(length(input$turnus) == 0){
      disable("download_btn")
      disable("klassenarbeiten")
    }
  })

  observeEvent(input$turnus, {
    enable("halbjahr")
    # shinyWidgets::updateAirDateInput(session = session, "klassenarbeiten", clear = T) # this is how it can be "reset"
  })

  observeEvent(c(input$halbjahr, input$turnus), {
    req(input$halbjahr, input$turnus)
    disable("halbjahr")
    disable("turnus")
    #jahr <- today()   # oben festgelegt
    if(!is.list(api()[1])){
      mindate <- floor_date(jahr, "halfyear")
      maxdate <- ceiling_date(jahr, "halfyear") + years(1)
      withProgress(message = 'Mache API-Abfragen', value = 0.0, {
        ferienintervall <- getHolidays(start = mindate,
                                       end = maxdate,
                                       pause = 2.5,
                                       progress = .5)
      })
      api(ferienintervall)
    } else {
      ferienintervall <- api()
    }
    halbjahrend(ferienintervall[[2]]$schuljahrenden)
    halbjahranf(ferienintervall[[2]]$schuljahranfaenge)

    hjr <- as.numeric(input$halbjahr)
    end <- halbjahrend()[hjr]
    anf <- halbjahranf()[hjr]
    tage <- as.Date(anf:end)
    turnus <- which(wochentage  %in% input$turnus)
    turnustage <- tage[which(lubridate::wday(tage, week_start = 1) %in% turnus)]
    ausfall(sapply(turnustage, function(x) any(x %within% api()[[1]])))
    #print(tage[ausfall()])
    termine(format(turnustage, "%a %d-%m-%y"))
    enable("halbjahr")
    enable("turnus")
    enable("klassenarbeiten")

    disabled_days <- which(!wochentage %in% input$turnus)
    shinyWidgets::updateAirDateInput(session = session, "klassenarbeiten", options = list(
      disabledDaysOfWeek = c(0,6,disabled_days),
      disabledDates = ymd(turnustage[ausfall()]),
      minDate = anf, maxDate = end))
    enable("download_btn")
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
        #browser()
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
        empty$schriftlich = sprintf(glue::glue('= IFERROR(AVERAGEIFS(F%d:{to}%d, F1:{to}1, "*KLAUSUR*", F1:{to}1, "<>*FREI*", F%d:{to}%d, "<>0"), "")'), 2:(SuS+1), 2:(SuS+1), 2:(SuS+1), 2:(SuS+1))

        # Datensatz neu-anordnen
        full <- empty %>% select(ID, `Nachname, Vorname`, Gesamtnote, muendlich, schriftlich, everything())

        # Formelspalten als solche klassifizieren
        for(i in c(3:5)){
          class(full[,i]) <- c(class(full[,i]), "formula")
        }

        ## Notebook erstellen
        wb <- createWorkbook()
        namehjr <- glue::glue("Noten_HBJ{input$halbjahr}")
        addWorksheet(wb, namehjr)
        freezePane(wb, namehjr, firstActiveCol = 6)
        writeData(wb, namehjr, x = full)
        if(input$rotate == TRUE){
          setColWidths(wb, namehjr, cols = c(1:2), widths = "auto")
          setColWidths(wb, namehjr, cols = 6:ncol(full), widths = 5)
          setRowHeights(wb, namehjr, rows = c(1), heights = c(80))
          headerstyle <- createStyle(halign = "center", valign = "center", textRotation = -90, wrapText  = TRUE)
        } else {
          setColWidths(wb, namehjr, cols = c(1:2, 6:ncol(full)), widths = "auto")
          setRowHeights(wb, namehjr, rows = c(1), heights = c(40))
          headerstyle <- createStyle(halign = "center", valign = "center")
        }

        # center headers
        addStyle(wb, sheet = namehjr, headerstyle, rows = 1, cols = 1:ncol(full))

        # füge borderColor hinzu + jeweils für body and headers
        bodyStyle <- createStyle(fgFill = 'grey95', border = "TopBottomLeftRight", borderStyle = 'thin', borderColour = 'grey65', halign = "center", valign = "center")
        addStyle(wb, sheet = namehjr, bodyStyle, rows = 1:(SuS+1), cols = 1:5, gridExpand = TRUE)
        bodyStyle2 <- createStyle(fgFill = 'grey90', border = c("top","bottom","left","right"), borderStyle = c('thin','thick','thin','thin'), borderColour = 'grey45', halign = "center", valign = "center")
        addStyle(wb, sheet = namehjr, bodyStyle2, rows = 1, cols = 1:5, gridExpand = TRUE)

        # Style festlegen
        neutralStyle <- createStyle(bgFill = "grey", borderColour = 'grey', border = "TopBottomLeftRight", borderStyle = 'thin')
        posStyle <- createStyle(bgFill = "#c6d7ef", border = "TopBottomLeftRight", borderStyle = 'thin', borderColour = 'grey65')
        negStyle <- createStyle(bgFill = "#EFDEC6", border = "TopBottomLeftRight", borderStyle = 'thin', borderColour = 'grey65')

        # add border to header
        neutralStyleH1 <- createStyle(bgFill = "grey50", borderColour = 'grey45', border = "Bottom", borderStyle = 'thick')
        posStyleH1 <- createStyle(bgFill = "#8FB8FF", border = "Bottom", borderStyle = 'thick', borderColour = 'grey45')
        negStyleH1 <- createStyle(bgFill = "#FFD68F", border = "Bottom", borderStyle = 'thick', borderColour = 'grey45')

        # header regeln für gesamte reihe
        idx0 <- 6:ncol(full)
        conditionalFormatting(wb, sheet =  namehjr, cols = idx0, rows = 1, style = posStyleH1, rule = "Klausur",
                              type = "notContains")
        conditionalFormatting(wb, sheet =  namehjr, cols = idx0, rows = 1, style = posStyleH1, rule = "Frei",
                              type = "notContains")
        conditionalFormatting(wb, sheet =  namehjr, cols = idx0, rows = 1, style = negStyleH1, rule = "Klausur",
                              type = "contains")
        conditionalFormatting(wb, sheet =  namehjr, cols = idx0, rows = 1, style = neutralStyleH1, rule = "Frei",
                              type = "contains")

        # add data validation (only 1-6 for date cells)
        dataValidation(wb,
                       sheet =  namehjr, cols = idx0, rows = 2:(SuS+1),
                       type = "decimal", operator = "between", value = c(0, 6))   # use decimal for exam grades
        # hacky way to prevent overwriting formula by disallowing values with textlength < 31
        dataValidation(wb,
                       sheet =  namehjr, cols = 3:5, rows = 2:(SuS+1),
                       type = "textLength", operator = "greaterThan", value = 30)


        incProgress(amount = 1/3)

        # body regeln für gesamte Spalten
        conditionalFormatting(wb, sheet =  namehjr, cols = idx0, rows = 2:(SuS+1), style = posStyle, rule = glue::glue('NOT(OR(ISNUMBER(SEARCH("Klausur", INDEX($1:$1,COLUMN()))), ISNUMBER(SEARCH("FREI", INDEX(1:1,COLUMN())))))'),
                              type = "expression")
        conditionalFormatting(wb, sheet =  namehjr, cols = idx0, rows = 2:(SuS+1), style = negStyle, rule = glue::glue('ISNUMBER(SEARCH("Klausur", INDEX($1:$1,COLUMN())))'),
                              type = "expression")
        conditionalFormatting(wb, sheet =  namehjr, cols = idx0, rows = 2:(SuS+1), style = neutralStyle, rule = glue::glue('ISNUMBER(SEARCH("FREI", INDEX($1:$1,COLUMN())))'),
                              type = "expression")

        incProgress(amount = 2/3)

        # body regeln für alle spalten ab spalte 6 incl.
        conditionalFormatting(wb, sheet =  namehjr, cols = idx0, rows = 2:(SuS+1), style = createStyle(bgFill = "white"), rule = ">6",
                              type = "expression")
        conditionalFormatting(wb, sheet =  namehjr, cols = idx0, rows = 2:(SuS+1), style = createStyle(bgFill = "white"), rule = '=AND(NOT(ISBLANK(F2)), F2 < 1)',   # not blank but below 1
                              type = "expression")
        conditionalFormatting(wb, sheet =  namehjr, cols = idx0, rows = 1, style = createStyle(bgFill = "white", border = "Bottom", borderStyle = 'thin'), rule = "-",
                              type = "notContains")

        incProgress(amount = 3/3)

        #remove gridlines
        showGridLines(wb, namehjr, showGridLines = FALSE)

        # notenspiegel
        notenspiegel <- data.frame(Note = 1:6, Anzahl = sprintf('=COUNTIFS(Noten!C%d:Noten!C%d, ">%d,5", Noten!C%d:Noten!C%d, "<=%d,5")', 2, SuS+1, 0:5, 2, SuS+1, 1:6))
        notenspiegel$Anzahl <- gsub("Noten", namehjr, notenspiegel$Anzahl)
        class(notenspiegel$Anzahl) <- c(class(notenspiegel$Anzahl), "formula")
        addWorksheet(wb, glue::glue("Notenspiegel_HBJ{input$halbjahr}"))
        writeData(wb, glue::glue("Notenspiegel_HBJ{input$halbjahr}"), x = notenspiegel)
        # disallow editing
        protectWorksheet(wb, glue::glue("Notenspiegel_HBJ{input$halbjahr}"), protect = TRUE)


        # documentation
        wb <- doku(wb = wb,
                   #nsus=  input$sus,
                   #ntermine=length(6:ncol(full)),
                   turnust=input$turnus,
                   halbjr = names(choices[as.numeric(input$halbjahr)]),
                   sheetname = glue::glue("Dokumentation_HBJ{input$halbjahr}"),
                   notensheet = namehjr)

        # speichern
        saveWorkbook(wb, file, overwrite = TRUE)
      }
      )
    })
}

shinyApp(ui, server)

