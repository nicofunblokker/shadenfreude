doku <- function(wb, turnust, halbjr, sheetname, notensheet){
# Add a worksheet for Dokumentation
addWorksheet(wb, sheetname)

# Increase the height of the rows
setRowHeights(wb, sheetname, rows = 2:8, height = 50) # Adjust the height as needed
setRowHeights(wb, sheetname, rows = 7, height = 80) # Adjust the height as needed

# Merge the first 5 columns without merging the rows
for(i in 2:8){
  mergeCells(wb, sheetname, cols = 2:6, rows = i)
}

# Apply formatting to the Dokumentation area
#addStyle(wb, sheet = sheetname, style = createStyle(textDecoration = "bold", fontSize = 12, wrapText = TRUE), rows = 1, cols = 1:5, gridExpand = TRUE)
addStyle(wb, sheet = sheetname, style = createStyle(fgFill = 'grey95', halign = "center", valign = "center", wrapText = TRUE, textDecoration ="bold", border = "BottomTopRightLeft", borderColour = c("white", "black","black","black" )), rows = 2, cols = 2:6, gridExpand = TRUE)
addStyle(wb, sheet = sheetname, style = createStyle(fgFill = 'grey95', halign = "left", valign = "center", wrapText = TRUE, border = "BottomTopRightLeft", borderColour = c("white", "white","black","black" )), rows = 3:7, cols = 2:6, gridExpand = TRUE)
addStyle(wb, sheet = sheetname, style = createStyle(fgFill = 'grey95', halign = "left", valign = "center", wrapText = TRUE, border = "LeftBottomRight", borderColour = "black"), rows = 8, cols = 2:6, gridExpand = TRUE)

# Write Dokumentation below the table
text <- "Zum Gebrauch der Notentabelle:\n– Spalten 1-2 frei editierbar. Spalten 3-5 beinhalten Formeln und sollten i.d.R nicht bearbeitet werden.\n– Die übrigen Spalten können mit Noten von -1-6 gefüllt werden (Nachkommastellen möglich). Werte <1 fallen nicht in die Gewichtung, ändern die Hintergrundfarbe jedoch auf weiß.\n– Dies kann genutzt werden, um Abwesenheiten einzutragen: 0 entschuldigte Abwesenheit, -1 unentschuldigt.\n– Blaue Spalten zählen in die mündliche Note, orangefarbene in die schriftliche. Grau ist für nicht-stattfindende Termine vorbehalten.\n–  Um Spalten nachträglich als Klausurtermine oder Ausfall zu deklarieren, muss 'Klausur' oder 'Frei' in den jeweiligen Spaltennamen eingetragen werden. Zusätzlich muss in jedem Fall '-' im Namen auftauchen.\n– So können ebenfalls Zusatztermine hinzugefügt und kodiert werden, z.B. '2025-02-12 Klausur' (Klausurtermin) oder '2025-02-12 Frei' (Ausfall) oder '2025-02-12' (mündlich)."

showGridLines(wb, sheetname, showGridLines = FALSE)
# Split the text by newline character
sentences <- unlist(strsplit(text, "\n"))

# Write each sentence in its own cell
for (i in 1:length(sentences)) {
  writeData(wb, sheetname, x = sentences[i], startCol = 2, startRow = i+1)
}

writeData(wb, sheetname, x = data.frame("Übersicht" = ""), startCol = 4, startRow = 10)
addStyle(wb, sheet = sheetname, style = createStyle(textDecoration = "bold", halign = "center", valign = "center", wrapText = F, border = "BottomTopRightLeft", borderColour = "white"), rows = 10, cols = 2:6, gridExpand = TRUE)

# erstell-infos + current infos
text2 <- glue::glue("{today()}")
d <- data.frame("Erstellt" = text2,
                "Halbjahr" = halbjr,
                "SuS" = glue::glue('=COUNTIF({notensheet}!A:A, ">0")'),
                "Termine" = glue::glue('=COUNTIFS({notensheet}!1:1, "*-*", {notensheet}!1:1, "<>*Klausur*", {notensheet}!1:1, "<>*Frei*") & " + " & COUNTIFS({notensheet}!1:1, "*Klausur*", {notensheet}!1:1, "<>*Frei*")'),
                "Turnus" = paste(turnust, collapse = "\n"))
for(i in c(3:4)){
  class(d[,i]) <- c(class(d[,i]), "formula")
}
writeData(wb, sheetname, x = d, startCol = 2, startRow = 11)
addStyle(wb, sheet = sheetname, style = createStyle(fgFill = 'grey95', halign = "center", valign = "center", wrapText = TRUE, border = "BottomTopRightLeft", borderColour = "black"), rows = 11:12, cols = 2:6, gridExpand = TRUE)

# addDropdown
value <- c(-1, 0, seq(1,6, by = 0.5))
value2 <- paste0(value[-c(1:2)], "H")

a = data.frame("Bedeutung" = c("untenschuldigt", "entschuldigt", rep("Notenwert", 11)), "Auswahl" = value)
b =data.frame("Bedeutung" = rep("Notenwert ohne Hausarbeit", 11), "Auswahl" = gsub("\\.", ",",  value2))

addStyle(wb, sheet = sheetname, style = createStyle(fgFill = 'grey95', wrapText = F), rows = 17:40, cols = 2:3, gridExpand = TRUE)
addStyle(wb, sheet = sheetname, style = createStyle(fgFill = 'grey65', wrapText = F), rows = 16, cols = 2:3, gridExpand = TRUE)

writeData(wb, sheetname, x = a, startCol = 2, startRow = 16)
writeData(wb, sheetname, x = b, startCol = 2, startRow = 17+length(value),colNames  =F)

# disallow editing
protectWorksheet(wb, sheetname, protect = TRUE)
return(wb)
}
