# Erstellen der Tabelle
kontaktdaten <- tibble::tibble(
  MSB        = "msb_name",
  EMAIL      = "empfaenger@example.com",
  ANREDE     = "Sehr geehrte Damen und Herren," ,
  EINLEITUNG = "Bitte finden Sie die angeforderten Daten unten:"
)

# Abspeichern der Variablen in einer .RData-Datei
save(kontaktdaten, file = "EMails.RData")


