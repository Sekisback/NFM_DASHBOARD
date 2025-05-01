# Erstellen der Tabelle Kontaktdaten
kontaktdaten <- tibble::tibble(
  MSB        = "msb_name",
  EMAIL      = "empfaenger@example.com",
  ANREDE     = "Sehr geehrte Damen und Herren," ,
  EINLEITUNG = "Bitte finden Sie die angeforderten Daten unten:"
)

# Abspeichern der Variablen in einer .RData-Datei
save(kontaktdaten, file = "EMails.RData")

# Erstellen der Tabelle  Blacklist
blacklist_df <- tibble::tibble(
  MELDEPUNKT  = "test_MP1234",
  MSB         = "test_MSB",
  INFORMATION = "test_Daten"
)

# Abspeichern beider tebellen in einer .RData-Datei
save(kontaktdaten_df, blacklist_df, file = "EMails.RData")
