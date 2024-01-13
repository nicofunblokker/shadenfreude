# Install and load the openxlsx package if not done already

library(openxlsx)
library(lubridate)
library(dplyr)
source("./R/getHolidays.R")

SuS <- 25   # Specify the number SuS

schulanfang <- "2024-01-08"   # Halbjahresintervall festlegen
schulende <- "2024-06-24"

klassenarbeiten <- c("2024-04-22", "2024-05-20")  # Klassenarbeiten festlegen

tage <- 2 # wie viele tage zwischen Unterrichtstunden (0 wenn nur 1x pro Woche)

# Feiertage und Ferien abrufen

#ferienintervall <- getHolidays()

# Sequenz festlegen
a <- seq(ymd(schulanfang),ymd(schulende), by = '1 week')
b <- seq(ymd(schulanfang)+tage ,ymd(schulende), by = '1 week')
termine <- c(a,b) |> sort()

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
addWorksheet(wb, "Sheet 1")
writeData(wb, "Sheet 1", x = full)

# Style festlegen
neutralStyle <- createStyle(bgFill = "grey")
posStyle <- createStyle(bgFill = "#C6EFCE")
negStyle <- createStyle(bgFill = "#FFC7CE")

idx0 <- which(colnames(full) %in% termine)
conditionalFormatting(wb, sheet =  "Sheet 1", cols = idx0, rows = 1:(SuS+1), style = posStyle, rule = "",
                      type = "contains")

idx <- which(colnames(full) %in% termine[ausfall])
for(i in idx){
  conditionalFormatting(wb, sheet =  "Sheet 1", cols = i, rows = 1:(SuS+1), style = neutralStyle, rule = "",
                        type = "contains")
}

idx2 <- which(colnames(full) %in% colnames(empty)[klassenarbeitsdaten])
for(i in idx2){
  conditionalFormatting(wb, sheet =  "Sheet 1", cols = i, rows = 1:(SuS+1), style = negStyle, rule = "",
                        type = "contains")
}

# speichern
saveWorkbook(wb, "writeFormulaHyperlinkExample.xlsx", overwrite = TRUE)


