# Load shiny libraries
library(shiny)
library(openxlsx)
library(lubridate)
library(dplyr)
library(shinyjs)
library(bslib)
library(ggplot2)
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
  tags$head(tags$link(rel = "icon", type = "image/png", href = "icon.png")),
  style = 'margin: 10px 15px',
  shinyjs::useShinyjs(),
  titlePanel("Notentabelle v3"),
  #HTML('Schritte bitte nacheinander ausfüllen.<br> Im Zweifel neuladen.'),
  HTML('Erstelle Exceltabelle mit einzelnen Sitzungen.'),
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


  shinyWidgets::switchInput(
    "show", label = '<i class="fa-solid fa-gear"></i>  Weitere Einstellungen',
    value = FALSE,
    labelWidth = "220px"),

  shinyjs::hidden(shinyWidgets::switchInput(
    inputId = "abwesend",
    label = '<i class="fa-solid fa-user-xmark"></i>  Abwesendheitssheet',
    value = TRUE,
    labelWidth = "220px"
  )),

  shinyjs::hidden(shinyWidgets::switchInput(
    inputId = "holiday",
    label = '<i class="fa-solid fa-champagne-glasses"></i>  freie Tage anzeigen',
    value = FALSE,
    labelWidth = "220px"
  )),

  shinyjs::hidden(shinyWidgets::switchInput(
    inputId = "rotate",
    label = '<i class="fa-solid fa-rotate"></i>  Spaltennamen rotieren',
    value = FALSE,
    labelWidth = "220px"
  )),

  tooltip(
  shinyjs::hidden(shinyWidgets::switchInput(
    inputId = "showplot",
    label = '<i class="fa-solid fa-layer-group"></i>  Übersicht anzeigen',
    value = FALSE,
    labelWidth = "220px"
  )), "Zeige nachfolgend eine kalendarische Übersicht des Halbjahrs"),

  shinyjs::hidden(plotOutput("plot", width = 300)),

  # filename
  shinyjs::hidden(shiny::textInput("filename", "Name der Notentabelle", placeholder = "10a_2024_HJ1")),

  # Download button
  downloadButton("download_btn", "Generiere Tabelle"),



  hr(),
  HTML("Referenz: <a href='https://www.openholidaysapi.org/en/'>Ferien- und Feiertagsdaten</a> unter <a href='https://raw.githubusercontent.com/openpotato/openholidaysapi.data/main/LICENSE'>dieser Lizenz</a>.")
)

server <- function(input, output, session) {
  disable("halbjahr")
  disable("klassenarbeiten")
  disable("download_btn")

  # hide optional settings
  observeEvent(c(input$show, input$showplot), {
    if(input$show == FALSE){
      shinyjs::hideElement(id= "plot")
      shinyjs::hideElement(id= "showplot")
      shinyjs::hideElement(id= "filename")
      shinyjs::hideElement(id= "holiday")
      shinyjs::hideElement(id= "rotate")
      shinyjs::hideElement(id= "abwesend")
    } else {
      if(input$showplot & is.list(api())){
        shinyjs::showElement(id= "plot")
      } else {
        shinyjs::hideElement(id= "plot")
      }
      shinyjs::showElement(id= "showplot")
      shinyjs::showElement(id= "filename")
      shinyjs::showElement(id= "holiday")
      shinyjs::showElement(id= "rotate")
      shinyjs::showElement(id= "abwesend")
    }

  })

  # reactive Values setzen
  termine <- reactiveVal(0)
  ausfall <- reactiveVal(0)
  api <- reactiveVal(0)           # api results zwischenspeichern
  halbjahrend <- reactiveVal(0)
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

    # showplot option nur möglich, wenn api abgefragt
    if(!is.list(api())){
      shinyWidgets::updateSwitchInput(session, "showplot", disabled = TRUE)
    } else {
      shinyWidgets::updateSwitchInput(session, "showplot", disabled = FALSE)
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
    #jahr <- today()   # oben global festgelegt
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
  # create plot
  output$plot <- renderPlot({
    req(input$halbjahr, is.list(api()))

    df <- data.frame(days = seq(halbjahranf()[as.numeric(input$halbjahr)], halbjahrend()[as.numeric(input$halbjahr)], "day")) %>%
      mutate(woche = as.numeric(lubridate::isoweek(days)),
             monat = month(floor_date(days, "month")),
             jahr = lubridate::isoyear(days),
             wd = weekdays(days)) %>%
      mutate(colr = if_else(wd %in% input$turnus, "wd", "nwd")) %>%
      group_by(jahr, woche) %>% mutate(idx = cur_group_id()) %>%
      ungroup() %>%
      mutate(woche_label = case_when(
        idx == min(idx) ~ paste("KW", woche),
        TRUE ~ as.character(woche)))

    df$wd <-  factor(df$wd, levels = c("Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"))

    idx <- which(sapply(df$days, function(x) any(x %within% api()[[1]])))
    df$colr[idx] <- "nwd"

    if(!is.null(input$klassenarbeiten)){
      klausuren <- ymd(input$klassenarbeiten)
      df$colr[df$days %in% klausuren] <- "Klausur"
    }

    ggplot(df, aes(x = wd, y = reorder(woche_label, -idx), color = "blue", fill = colr)) +
      geom_tile(show.legend = F) +
      scale_color_manual(values = "white") +
      scale_fill_manual(values = c("nwd" = "grey90", "Klausur" = "salmon", "wd" = "cornflowerblue")) +
      coord_equal() +
      theme_void() +
      theme(axis.text.x = element_text(angle = 90, hjust=0.95),
            axis.text.y = element_text(hjust = 1),
            plot.margin = margin(b=0.5, unit = "cm"))+
      scale_x_discrete(labels = substr(levels(df$wd), 1,3)) +
      ggtitle("Übersicht Hbj.")
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
        empty$Hausarbeit = sprintf(glue::glue('=IF(COUNTIF(F%d:{to}%d, "*") = 0, "", COUNTIF(F%d:{to}%d, "*"))'), 2:(SuS+1), 2:(SuS+1),2:(SuS+1), 2:(SuS+1))
        empty$`Nachname, Vorname` = rep("", SuS)
        empty$Gesamtnote = sprintf('=IFERROR(AVERAGEIF(C%d:D%d, "<>0", C%d:D%d), "")', 2:(SuS+1), 2:(SuS+1), 2:(SuS+1), 2:(SuS+1))
        #empty$muendlich = sprintf(glue::glue('= IFERROR(AVERAGEIFS(F%d:{to}%d, F1:{to}1, "<>*KLAUSUR*", F1:{to}1, "<>*FREI*", F%d:{to}%d, ">0"), "")'), 2:(SuS+1), 2:(SuS+1), 2:(SuS+1), 2:(SuS+1))
        muendlich = sprintf(glue::glue('= IFERROR(SUMPRODUCT(IF(IFERROR(ISNUMBER(SEARCH("-", F1:{to}1))*NOT(ISNUMBER(SEARCH("KLAUSUR", F1:{to}1)))*NOT(ISNUMBER(SEARCH("FREI", F1:{to}1)))*(--ISNUMBER(VALUE(SUBSTITUTE(F%d:{to}%d,"*",""))))*(VALUE(SUBSTITUTE(F%d:{to}%d,"*",""))>=1),0), VALUE(SUBSTITUTE(F%d:{to}%d,"*","")), 0)) / SUMPRODUCT(--(ISNUMBER(SEARCH("-", F1:{to}1)))*NOT(ISNUMBER(SEARCH("KLAUSUR", F1:{to}1)))*NOT(ISNUMBER(SEARCH("FREI", F1:{to}1)))*(--ISNUMBER(IFERROR(VALUE(SUBSTITUTE(F%d:{to}%d,"*","")),0)))*(IFERROR(VALUE(SUBSTITUTE(F%d:{to}%d,"*",""))>=1, 0))),"")'), 2:(SuS+1), 2:(SuS+1), 2:(SuS+1), 2:(SuS+1), 2:(SuS+1), 2:(SuS+1), 2:(SuS+1), 2:(SuS+1), 2:(SuS+1), 2:(SuS+1))
        empty$muendlich = NA
        schriftlich = sprintf(glue::glue('= IFERROR(SUMPRODUCT(IF(IFERROR(ISNUMBER(SEARCH("-", F1:{to}1))*(ISNUMBER(SEARCH("KLAUSUR", F1:{to}1)))*NOT(ISNUMBER(SEARCH("FREI", F1:{to}1)))*(--ISNUMBER(VALUE(SUBSTITUTE(F%d:{to}%d,"*",""))))*(VALUE(SUBSTITUTE(F%d:{to}%d,"*",""))>=1),0), VALUE(SUBSTITUTE(F%d:{to}%d,"*","")), 0)) / SUMPRODUCT(--(ISNUMBER(SEARCH("-", F1:{to}1)))*(ISNUMBER(SEARCH("KLAUSUR", F1:{to}1)))*NOT(ISNUMBER(SEARCH("FREI", F1:{to}1)))*(--ISNUMBER(IFERROR(VALUE(SUBSTITUTE(F%d:{to}%d,"*","")),0)))*(IFERROR(VALUE(SUBSTITUTE(F%d:{to}%d,"*",""))>=1, 0))),"")'), 2:(SuS+1), 2:(SuS+1), 2:(SuS+1), 2:(SuS+1), 2:(SuS+1), 2:(SuS+1), 2:(SuS+1), 2:(SuS+1), 2:(SuS+1), 2:(SuS+1))
        empty$schriftlich = NA
        #empty$schriftlich = sprintf(glue::glue('= IFERROR(AVERAGEIFS(F%d:{to}%d, F1:{to}1, "*KLAUSUR*", F1:{to}1, "<>*FREI*", F%d:{to}%d, ">0"), "")'), 2:(SuS+1), 2:(SuS+1), 2:(SuS+1), 2:(SuS+1))

        # Datensatz neu-anordnen
        full <- empty %>% select(`Nachname, Vorname`, Gesamtnote, muendlich, schriftlich, Hausarbeit,  everything())

        # Formelspalten als solche klassifizieren
        for(i in c(2,5)){
          class(full[,i]) <- c(class(full[,i]), "formula")
        }

        ## Notebook erstellen
        wb <- createWorkbook()
        namehjr <- glue::glue("Noten_HBJ{input$halbjahr}")
        addWorksheet(wb, namehjr)
        freezePane(wb, namehjr, firstActiveCol = 6)
        writeData(wb, namehjr, x = full)
        for(row in seq_along(muendlich)){
          writeFormula(wb, namehjr, x = muendlich[row], startCol = 3, startRow = row+1, array = T)
        }
        for(row in seq_along(schriftlich)){
          writeFormula(wb, namehjr, x = schriftlich[row], startCol = 4, startRow = row+1, array = T)
        }

        if(input$rotate == TRUE){
          setColWidths(wb, namehjr, cols = c(1), widths = 25)
          setColWidths(wb, namehjr, cols = 6:ncol(full), widths = 5)
          setRowHeights(wb, namehjr, rows = c(1), heights = c(80))
          headerstyle <- createStyle(halign = "center", valign = "center", textRotation = -90, wrapText  = TRUE)
        } else {
          setColWidths(wb, namehjr, cols = 1, widths = 25)
          setColWidths(wb, namehjr, cols = c(6:ncol(full)), widths = "auto")
          setRowHeights(wb, namehjr, rows = c(1), heights = c(40))
          headerstyle <- createStyle(halign = "center", valign = "center")
        }

        # center headers
        addStyle(wb, sheet = namehjr, headerstyle, rows = 1, cols = 1:ncol(full))

        # füge borderColor hinzu + jeweils für body and headers
        bodyStyle <- createStyle(fgFill = 'grey95', border = "TopBottomLeftRight", borderStyle = 'thin', borderColour = 'grey65', halign = "center", valign = "center")
        addStyle(wb, sheet = namehjr, bodyStyle, rows = 1:(SuS+1), cols = 1:5, gridExpand = TRUE)
        # Namenspalte left aligned
        bodyStyle <- createStyle(fgFill = 'grey95', border = "TopBottomLeftRight", borderStyle = 'thin', borderColour = 'grey65', halign = "left", valign = "center")
        addStyle(wb, sheet = namehjr, bodyStyle, rows = 1:(SuS+1), cols = 1, gridExpand = TRUE)
        bodyStyle2 <- createStyle(fgFill = 'grey90', border = c("top","bottom","left","right"), borderStyle = c('thin','thick','thin','thin'), borderColour = 'grey45', halign = "center", valign = "center")
        addStyle(wb, sheet = namehjr, bodyStyle2, rows = 1, cols = 1:5, gridExpand = TRUE)

        # center cells
        bodyStyleCELLS <- createStyle(halign = "right", valign = "center")
        addStyle(wb, sheet = namehjr, bodyStyleCELLS, rows = 2:(SuS+1), cols = 6:length(full), gridExpand = TRUE)


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
                       type = "list", value = glue::glue("Info_HBJ{input$halbjahr}!$C$17:$C$42"))   # use decimal for exam grades
        # hacky way to prevent overwriting formula by disallowing values with textlength < 31
        dataValidation(wb,
                       sheet =  namehjr, cols = 2:5, rows = 2:(SuS+1),
                       type = "textLength", operator = "greaterThan", value = 30)

        # progressbar update
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
        conditionalFormatting(wb, sheet =  namehjr, cols = idx0, rows = 2:(SuS+1), style = createStyle(bgFill = "white"), rule = '=OR(SUBSTITUTE(F2, "*", "") = "-1", SUBSTITUTE(F2, "*", "") = "0")',
                              type = "expression")
        conditionalFormatting(wb, sheet =  namehjr, cols = idx0, rows = 2:(SuS+1), style = createStyle(bgFill = "white"), rule = '=AND(NOT(ISBLANK(F2)), F2 < 1)',   # not blank but below 1
                              type = "expression")
        conditionalFormatting(wb, sheet =  namehjr, cols = idx0, rows = 1, style = createStyle(bgFill = "white", border = "Bottom", borderStyle = 'thin'), rule = "-",
                              type = "notContains")

        incProgress(amount = 3/3)

        #remove gridlines
        showGridLines(wb, namehjr, showGridLines = FALSE)

        # notenspiegel
        notenspiegel <- data.frame(Note = 1:6, Anzahl = sprintf('=COUNTIFS(Noten!B%d:Noten!B%d, ">%d,5", Noten!B%d:Noten!B%d, "<=%d,5")', 2, SuS+1, 0:5, 2, SuS+1, 1:6))
        notenspiegel$Anzahl <- gsub("Noten", namehjr, notenspiegel$Anzahl)
        class(notenspiegel$Anzahl) <- c(class(notenspiegel$Anzahl), "formula")
        addWorksheet(wb, glue::glue("Notenspiegel_HBJ{input$halbjahr}"))
        addStyle(wb, sheet = glue::glue("Notenspiegel_HBJ{input$halbjahr}"), style = createStyle(fgFill = 'grey95', textDecoration ="bold", border = "BottomTopRightLeft", borderColour = c("grey65", "grey95","grey95","grey95" ), borderStyle = "thick"), rows = 1, cols = 1:2, gridExpand = TRUE)
        addStyle(wb, sheet = glue::glue("Notenspiegel_HBJ{input$halbjahr}"), style = createStyle(fgFill = 'grey95'), rows = 2:7, cols = 1:2, gridExpand = TRUE)
        showGridLines(wb, glue::glue("Notenspiegel_HBJ{input$halbjahr}"), showGridLines = FALSE)
        writeData(wb, glue::glue("Notenspiegel_HBJ{input$halbjahr}"), x = notenspiegel)
        # disallow editing
        protectWorksheet(wb, glue::glue("Notenspiegel_HBJ{input$halbjahr}"), protect = TRUE)

        # abwesenheitszeiten
        if(input$abwesend){
          # variablen erstellen
          addWorksheet(wb, glue::glue("Abwesendheit_HBJ{input$halbjahr}"))
          #ID = sprintf(glue::glue('=IFERROR({namehjr}!A%d, "")'), 2:(SuS+1))
          Name = sprintf(glue::glue('=IFERROR({namehjr}!A%d, "")'), 2:(SuS+1))
          Abwesend_einzeln = sprintf(glue::glue('=IFERROR(COUNTIFS({namehjr}!F1:{to}1, "<>*Frei*", {namehjr}!F%d:{to}%d, "<1", {namehjr}!F%d:{to}%d, "<>"), "")'), 2:(SuS+1),2:(SuS+1),2:(SuS+1),2:(SuS+1))
          Abwesend = sprintf(glue::glue('=IFERROR(COUNTIFS({namehjr}!F1:{to}1, "<>*Frei*", {namehjr}!F%d:{to}%d, "<1", {namehjr}!F%d:{to}%d, "<>"), "") & " (" & IFERROR(COUNTIFS({namehjr}!F1:{to}1, "<>*Frei*", {namehjr}!F%d:{to}%d, -1, {namehjr}!F%d:{to}%d, "<>"), "") & ")"'), 2:(SuS+1),2:(SuS+1),2:(SuS+1),2:(SuS+1),2:(SuS+1),2:(SuS+1),2:(SuS+1),2:(SuS+1))
          Sitzungen = sprintf(glue::glue('=IFERROR(COUNTIFS({namehjr}!F1:{to}1, "<>*Frei*", {namehjr}!F%d:{to}%d, "<>"), "")'), 2:(SuS+1),2:(SuS+1))
          Prozent = glue::glue('=IFERROR(ROUND(({gsub("=","", Abwesend_einzeln)} / {gsub("=","", Sitzungen)})*100, 2), 0)')
          #Unentschuld = sprintf(glue::glue('=IFERROR(COUNTIFS({namehjr}!F1:{to}1, "<>*Frei*", {namehjr}!F%d:{to}%d, -1, {namehjr}!F%d:{to}%d, "<>"), "")'), 2:(SuS+1),2:(SuS+1),2:(SuS+1),2:(SuS+1))
          abwesendheit <- data.frame(Name, Sitzungen, Abwesend, Prozent)
          for(i in c(1:4)){
            class(abwesendheit[,i]) <- c(class(abwesendheit[,i]), "formula")
          }
          #style hinzufügen
          addStyle(wb, sheet = glue::glue("Abwesendheit_HBJ{input$halbjahr}"), style = createStyle(fgFill = 'grey95', textDecoration ="bold", border = "BottomTopRightLeft", borderColour = c("grey65", "grey95","grey95","grey95" ), borderStyle = "thick"), rows = 1, cols = 1:4, gridExpand = TRUE)
          addStyle(wb, sheet = glue::glue("Abwesendheit_HBJ{input$halbjahr}"), style = createStyle(fgFill = 'grey95'), rows = 2:(SuS+1), cols = 1:4, gridExpand = TRUE)
          addStyle(wb, sheet = glue::glue("Abwesendheit_HBJ{input$halbjahr}"), style = createStyle(fgFill = 'grey95', halign = "right"), rows = 2:(SuS+1), cols = 3, gridExpand = TRUE)
          writeData(wb, glue::glue("Abwesendheit_HBJ{input$halbjahr}"), x = abwesendheit, startCol = 1, startRow = 1)

          # add average
          avg_abwesend = sprintf(glue::glue('="Ø " & ROUND(SUMPRODUCT(VALUE(LEFT(C2:C%d, FIND("(", C2:C%d)-1))), 2) / COUNTA(C2:C%d)'), (SuS+1), (SuS+1), (SuS+1))   # instead of sumproduct use sum? cannot use average here because excel adds @
          avg_percent = sprintf(glue::glue('="Ø " & IFERROR(ROUND(AVERAGE(D2:D%d), 2), "")'), (SuS+1))
          averages <- data.frame(avg_abwesend, avg_percent)
          for(i in c(1:2)){
            class(averages[,i]) <- c(class(averages[,i]), "formula")
          }
          addStyle(wb, sheet = glue::glue("Abwesendheit_HBJ{input$halbjahr}"), style = createStyle(halign = "right", fg = "grey80"), rows = SuS+2, cols = 3:4, gridExpand = TRUE)
          writeData(wb, glue::glue("Abwesendheit_HBJ{input$halbjahr}"), x = averages, startCol = 3, startRow = SuS+2, colNames = FALSE)

          # add limitation note
          writeData(wb, glue::glue("Abwesendheit_HBJ{input$halbjahr}"), x = data.frame("Hinweis" = c("*Einträge in der Notentabelle mit '0' gelten als abwesend, '-1' als unentschuldigt (in Klammern angegeben).", "**Neuzugänge und Abgänge müssen händisch hinzugefügt und Formel ggfs. angepasst werden.")), startCol = 1, startRow = SuS+3, colNames = FALSE)

          # disallow editing
          protectWorksheet(wb, glue::glue("Abwesendheit_HBJ{input$halbjahr}"), protect = TRUE)
          #remove gridlines
          showGridLines(wb, glue::glue("Abwesendheit_HBJ{input$halbjahr}"), showGridLines = FALSE)
        }

        # documentation sheet
        wb <- doku(wb = wb,
                   #nsus=  input$sus,
                   #ntermine=length(6:ncol(full)),
                   turnust=input$turnus,
                   halbjr = names(choices[as.numeric(input$halbjahr)]),
                   sheetname = glue::glue("Info_HBJ{input$halbjahr}"),
                   notensheet = namehjr)

        # save entire workbook
        saveWorkbook(wb, file, overwrite = TRUE)
      }
      )
    })
}

shinyApp(ui, server)

