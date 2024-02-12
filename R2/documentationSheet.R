doku <- function(wb, nsus, ntermine, turnust, halbjr, sheetname){


# Create a new workbook
#wb <- createWorkbook()

# Add a worksheet for Dokumentation
addWorksheet(wb, sheetname)

# Increase the height of the rows
setRowHeights(wb, sheetname, rows = 1:6, height = 50) # Adjust the height as needed
setRowHeights(wb, sheetname, rows = 5, height = 80) # Adjust the height as needed

# Merge the first 5 columns without merging the rows
for(i in 1:6){
  mergeCells(wb, sheetname, cols = 1:5, rows = i)
}


# Apply formatting to the Dokumentation area
#addStyle(wb, sheet = sheetname, style = createStyle(textDecoration = "bold", fontSize = 12, wrapText = TRUE), rows = 1, cols = 1:5, gridExpand = TRUE)
addStyle(wb, sheet = sheetname, style = createStyle(halign = "center", valign = "center", wrapText = TRUE, textDecoration ="bold", border = "BottomTopRightLeft", borderColour = c("white", "black","black","black" )), rows = 1, cols = 1:5, gridExpand = TRUE)
addStyle(wb, sheet = sheetname, style = createStyle(halign = "left", valign = "center", wrapText = TRUE, border = "BottomTopRightLeft", borderColour = c("white", "white","black","black" )), rows = 2:5, cols = 1:5, gridExpand = TRUE)
addStyle(wb, sheet = sheetname, style = createStyle(halign = "left", valign = "center", wrapText = TRUE, border = "BottomRight", borderColour = "black"), rows = 6, cols = 1:5, gridExpand = TRUE)

# Write Dokumentation below the table
text <- "Dokumentation zum Gebrauch der Notentabelle:\n– Spalten 1-2 frei editierbar. Spalten 3-5 beinhalten Formeln und sollten i.d.R nicht bearbeitet werden.\n– Die übrigen Spalten können mit Noten von 1-6 gefüllt werden (Nachkommastellen möglich).\n– Blaue Spalten zählen in die mündliche Note, orange in die schriftliche. Grau ist für nicht-stattfindende Termine vorbehalten.\n–  Um Spalten nachträglich als Klausurtermine oder Ausfall zu deklarieren, muss 'Klausur' oder 'Frei' in den jeweiligen Spaltennamen eingetragen werden. Zusätzlich muss in jedem Fall '-' im Namen auftauchen.\n– So können ebenfalls Zusatztermine hinzugefügt und kodiert werden, z.B. '2025-02-12 Klausur' (Klausurtermin) oder '2025-02-12 Frei' (Ausfall) oder '2025-02-12' (mündlich)."

showGridLines(wb, sheetname, showGridLines = FALSE)
# Split the text by newline character
sentences <- unlist(strsplit(text, "\n"))

# Write each sentence in its own cell
for (i in 1:length(sentences)) {
  writeData(wb, sheetname, x = sentences[i], startCol = 1, startRow = i)
}

writeData(wb, sheetname, x = data.frame("Eingabe" = ""), startCol = 3, startRow = 8)
addStyle(wb, sheet = sheetname, style = createStyle(textDecoration = "bold", halign = "center", valign = "center", wrapText = F, border = "BottomTopRightLeft", borderColour = "white"), rows = 8, cols = 1:5, gridExpand = TRUE)

text2 <- glue::glue("{today()}")
d <- data.frame("Erstellt" = text2, "Halbjahr" = halbjr, "SuS" = nsus, "Termine" = ntermine, "Turnus" = paste(turnust, collapse = "\n"))
writeData(wb, sheetname, x = d, startCol = 1, startRow = 9)
addStyle(wb, sheet = sheetname, style = createStyle(halign = "center", valign = "center", wrapText = TRUE, border = "BottomTopRightLeft", borderColour = "black"), rows = 9:10, cols = 1:5, gridExpand = TRUE)

# disallow editing
protectWorksheet(wb, sheetname, protect = TRUE)
return(wb)

}
