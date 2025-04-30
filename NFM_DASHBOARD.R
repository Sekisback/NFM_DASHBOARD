# --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --
#
# NFM DASHBOARD                                                             ----
#
# Author : Sascha Kornberger
# Datum  : 26.04.2025
# Version: 0.2.1
#
# History:
# 0.2.1  Bugfix : Diagramm baklenhöhe aktualisieren
# 0.2.0  Bugfix : Dropdown Filter aktualisieren 
#
# --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --
rm(list = ls())

## BENOETIGTE PAKETE ------------------------------------------------------------
# Optionen setzen – unterdrückt manche Dialoge zusätzlich
options(install.packages.check.source = "no")

# Liste der Pakete
pakete <- c(
  "devtools", "shiny", "bs4Dash", "DT", "tidyverse", "janitor",
  "lubridate", "base64enc", "shinyWidgets", "readr", "stringr", "purrr"
)

# Installiere fehlende Pakete ohne Rückfragen
installiere_fehlende <- pakete[!pakete %in% installed.packages()[, "Package"]]
if (length(installiere_fehlende) > 0) {
  install.packages(
    installiere_fehlende,
    repos = "https://cran.r-project.org",
    quiet = TRUE
  )
}

# Lade alle Pakete
invisible(lapply(pakete, function(pkg) {
  suppressPackageStartupMessages(library(pkg, character.only = TRUE))
}))

# RDCOMClient ist nicht in CRAN, daher muss es von GitHub installiert werden
# Prüfen, ob das System Windows ist
if (.Platform$OS.type == "windows") {
  if (!requireNamespace("RDCOMClient", quietly = TRUE)) {
    devtools::install_github("omegahat/RDCOMClient")
  }
  suppressPackageStartupMessages(library(RDCOMClient))
}


## OPTIONS ----------------------------------------------------------------------
# Vermeide Exponentialfunktion
options(scipen = 999)

# Keine Ausgabe in der CLI beim laden von Tidyverse
options(tidiverse.quiet = TRUE)


# --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- -#
# ----                           FUNKTIONEN                                 ----
# --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- -#


# Laden der RData-Datei mit den E-Mail Kontakten
load_email_data <- function() {
  load("EMails.RData") 
  return(kontaktdaten) 
}

# Speichern der RData-Datei mit den E-Mail Kontakten nach Änderung
save_email_data <- function(kontaktdaten) {
  save(kontaktdaten, file = "EMails.RData")
}


# Abrufen der CSV-Dateien aus dem Ordner
get_csv_files <- function(directory = "data") {
  # Hole alle CSV-Dateien im Ordner
  files <- list.files(directory, pattern = "\\.csv$", full.names = TRUE)
  return(files)
}

# Ereldigte CSV in done verschieben
move_csv_to_done <- function(file_path, done_folder = "done") {
  # Erstelle den Zielordner, falls er noch nicht existiert
  if (!dir.exists(done_folder)) {
    dir.create(done_folder, recursive = TRUE)
  }
  
  # Zielpfad
  target_path <- file.path(done_folder, basename(file_path))
  
  # Datei verschieben
  file.rename(file_path, target_path)
}


# Aktualisiert die CSV-Dateien und den File-Filter nach dem Verschieben einer CSV-Datei.
update_after_csv_move <- function(session, csv_files, current_msb_type, current_file_filter) {
  # Rufe die aktuelle Liste der CSV-Dateien ab
  new_files <- get_csv_files()
  # Aktualisiere die reaktive Variable csv_files mit der neuen Liste
  csv_files(new_files)
}


# Extrahieren des Namens zwischen "MSB" und "an" aus den Dateinamen
extract_msb_name <- function(file_names) {
  # Extrahiere Teile der Dateinamen nach "MSB " und vor " an"
  msb_names <- str_extract(file_names, "(?<=MSB\\s)(.*?)(?=\\san)")
  # Gib die extrahierten Namen zurück, ohne fehlende Werte
  return(msb_names[!is.na(msb_names)])  
}


# Extrahieren des MSB-Typs (wMSB oder gMSB) aus den Dateinamen
extract_msb_type <- function(file_names) {
  # Suche nach 'wMSB' oder 'gMSB' im Dateinamen
  msb_types <- str_extract(file_names, "(wMSB|gMSB)")
  # Gib die gefundenen Typen ohne leere Werte zurück
  return(msb_types[!is.na(msb_types)]) 
}


# Extrahieren des VNB aus den Dateinamen
extract_vnb <- function(filename) {
  # Suche Text zwischen " an " und "[" oder ".csv"
  vnb <- str_extract(filename, "(?<= an )(.*?)(?=\\[|\\.csv)")
  # Entferne Leerzeichen am Anfang und Ende
  vnb <- str_trim(vnb)
  # Gib den extrahierten Text zurück
  return(vnb)
}


# Extrahiere den oder die Monate aus dem Datenframe
extract_monate <- function(df) {
  # Wenn der Datenframe leer ist, gib einen leeren String zurück
  if (nrow(df) == 0) return("")
  
  # Bereinige die Spaltennamen (alles in Großbuchstaben)
  df <- df |> clean_names("all_caps")
  
  # Wähle die relevanten Spalten für den Zeitraum aus
  df_relevant <- df |>
    select(LUCKE_VON, LUCKE_BIS) |>
    # Konvertiere die Datums- und Zeitangaben
    mutate(
      LUCKE_VON = dmy_hm(LUCKE_VON),
      LUCKE_BIS = dmy_hm(LUCKE_BIS)
    )
  
  # Erstelle eine Liste der Monate zwischen Start- und Enddatum
  monate <- map2(
    df_relevant$LUCKE_VON,
    df_relevant$LUCKE_BIS,
    # Erzeugt eine Sequenz von Monatsanfängen
    ~ seq.Date(from = floor_date(.x, "month"), to = floor_date(.y, "month"), by = "1 month")
  ) |>
    # Reduziere die Liste zu einem Vektor
    unlist() |>
    # Stelle sicher, dass es Datumsangaben sind
    as.Date(origin = "1970-01-01") |>
    # Entferne doppelte Monatsangaben
    unique()
  
  # Formatiere die Monatszahlen zu Monatsnamen
  monatsnamen <- format(monate, "%B")
  # Extrahiere die Jahreszahlen
  jahre <- format(monate, "%Y")
  
  # Wenn nur ein Monat vorhanden ist
  if (length(monate) == 1) {
    # Erstelle eine formatierte Ausgabe: "Monatsname Jahr"
    monate_final <- paste0(monatsnamen[1], " ", jahre[1])
  } else {
    # Erstelle eine formatierte Ausgabe für mehrere Monate
    monate_final <- paste(
      # Alle Monatsnamen außer dem letzten, durch Komma getrennt
      paste(monatsnamen[-length(monatsnamen)], collapse = ", "),
      "und",
      # Der letzte Monatsname mit der zugehörigen Jahreszahl
      paste0(monatsnamen[length(monatsnamen)], " ", jahre[length(jahre)])
    )
  }
  # Gib die formatierte Zeichenkette der Monate zurück
  return(monate_final)
}


# Erstelle Plot als Balkendiagramm mit der Menge der Dateien pro MSB
msb_plot <- function(file_names) {
  # Extrahiere die MSB-Namen aus den Dateinamen
  msb_names <- extract_msb_name(basename(file_names))
  
  # Zähle, wie oft jeder MSB-Name vorkommt
  msb_count <- tibble(msb = msb_names) |> 
    count(msb, sort = TRUE)
  
  # Finde den höchsten Zählwert für die X-Achsenbegrenzung
  max_value <- max(msb_count$n, na.rm = TRUE)
  
  # Erstelle den ggplot
  ggplot(msb_count, aes(x = n, y = reorder(msb, n))) +
    # Erstelle Balken mit blauer Füllung und schmaler Breite
    geom_col(fill = "steelblue", width = 0.4, just = 1) +
    # Füge eine vertikale Linie bei x = 0 hinzu
    geom_vline(xintercept = 0) +
    # Füge Textbeschriftungen für die MSB-Namen hinzu
    geom_text(
      data = msb_count,
      mapping = aes(
        x = 0,
        y = msb,
        # Beschrifte mit großgeschriebenem MSB-Namen
        label = str_to_title(msb)
      ),
      # Horizontal linksbündig ausrichten
      hjust = 0,
      # Leichte horizontale Verschiebung
      nudge_x = 0.1,
      # Vertical mittig ausrichten
      vjust = 0,
      # Leichte vertikale Verschiebung
      nudge_y = 0.15,
      # Schwarze Textfarbe
      color = "black",
      # Fettdruck
      fontface = "bold",
      # Schriftgröße 5
      size = 5
    )+
    # Verwende minimalistisches Theme mit Arial-Schrift
    theme_minimal(base_family = "Arial") +
    # Entferne Achsenbeschriftungen und Titel
    labs(
      x = element_blank(),
      y = element_blank(),
      title = element_blank(),
    ) +
    # Entferne die Hauptgitterlinien
    theme(
      panel.grid.major = element_blank(),
    )+
    # Entferne die Beschriftungen der y-Achse
    scale_y_discrete(labels = NULL)+
    # Definiere die Skala der x-Achse
    scale_x_continuous(
      # Setze Limits oder verwende NA für automatische Anpassung
      limits = c(0, ifelse(max_value < 10, 10, NA)),
      # Definiere die Achsenbruche
      breaks = 0:max(10, max_value),
      # Füge etwas Platz am Rand hinzu
      expand = expansion(mult = c(0, 0.05))
    )
}


# Modal für Neu- und Edit-Einträge
entry_modal <- function(mode = c("new", "edit"), data = NULL) {
  # Stelle sicher, dass der Modus "new" oder "edit" ist
  mode <- match.arg(mode)
  
  # Setze den Titel des Dialogfensters basierend auf dem Modus
  dlg_title <- if (mode == "new") {
    "Neuen Eintrag hinzufügen"
  } else {
    "Eintrag bearbeiten"
  }
  
  # Standard-Einleitungstext für neue Einträge
  default_einleitung <- "für die unten aufgeführten Marktlokationen wurden, trotz zuvor versendeter ORDERS, noch keine vollständigen Lastgangdaten empfangen.<br>Wir möchten Sie daher bitten, die Lastgangdaten für den Monat {MONAT} noch einmal zu versenden.<br>Nutzen Sie hierzu bitte die aktuellen MSCONS-Version und die Ihnen bekannte EDIFACT-Adresse."
  # Standard-Anrede
  default_anrede <- "Sehr geehrte Damen und Herren,"
  
  # Setze Standardwerte oder verwende Daten für die Bearbeitung
  vals <- list(
    msb        = if (mode == "edit") data$MSB        else "",
    email      = if (mode == "edit") data$EMAIL      else "",
    anrede     = if (mode == "edit") data$ANREDE     else default_anrede,
    einleitung = if (mode == "edit") data$EINLEITUNG else default_einleitung
  )
  
  # Definiere die ID des Speicher-Buttons basierend auf dem Modus
  save_id <- if (mode == "new") "save_new_entry" else "save_edit_entry"
  
  # Erstelle das modale Dialogfenster
  modalDialog(
    size = "xl",
    title = dlg_title,
    fluidRow(
      column(4, textInput("modal_msb", "MSB Name", value = vals$msb)),
      column(4, textInput("modal_email", "E-Mail Adresse", value = vals$email)),
      column(4, textInput("modal_anrede", "Anrede", value = vals$anrede))
    ),
    # Eingabefeld für den Einleitungssatz (mehrzeilig)
    textAreaInput(
      "modal_einleitung", 
      "Einleitungssatz", 
      value = vals$einleitung, 
      # Breite setzen
      width = "100%",
      # Mindestens 3 Zeilen hoch
      rows = 3             
    ),
    # Fußzeile des Modals mit Buttons
    footer = tagList(
      modalButton("Abbrechen"),
      actionButton(save_id, "Speichern", class = "btn-primary")
    ),
    # Ermögliche das Schließen des Modals durch Klicken außerhalb
    easyClose = TRUE
  )
}

# --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --
#                             E-MAIL BODY VORLAGEN                          ----
# --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --

# CSV IN HTML-TABELLE 
generate_email_body <- function(df, contact, monate_final_plain) {
  # Wenn der Datenframe leer ist, gib eine Info-Nachricht zurück
  if (nrow(df) == 0) return("<i>Keine Daten vorhanden.</i>")
  
  # Bereinige die Spaltennamen (alles in Großbuchstaben)
  df <- df |> clean_names("all_caps")
  
  # Erstelle den Anredetext
  anrede_text <- contact$ANREDE

  # Erstelle den Einleitungstext
  einleitung_text <- contact$EINLEITUNG
  
  # Formatiere die Monate fett für HTML
  monate_final_html <- paste0("<b>", monate_final_plain, "</b>")
  
  # Ersetze Platzhalter {MONAT} im Einleitungstext
  einleitung_text <- str_replace_all(einleitung_text, "\\{MONAT\\}", monate_final_html)
  
  # Extrahiere die eindeutigen Meldepunkte
  meldepunkte <- df$MELDEPUNKT |> unique()
  # Formatiere die Meldepunkte als Liste für HTML
  meldepunkte_text <- paste0("&nbsp;&nbsp;• ", meldepunkte, collapse = "<br>")
  
  # Füge alle Teile zum E-Mail-Body zusammen
  email_body <- paste0(
    anrede_text, "<br><br>",
    einleitung_text, "<br><br>",
    "Betroffene Messlokationen:<br>",
    meldepunkte_text
  )
  # Gib den erstellten E-Mail-Body zurück
  return(email_body)
}


# E-MAIL SIGNATUR HTML
signatur_html <- function() {
 
  # Liest die Bilddatei und konvertiert sie in ein Base64-kodiertes URI-Schema 
  logo_eeg <- dataURI(file = "www/images/logo_eeg.png", mime = "image/png")
  icon_mail <- dataURI(file = "www/images/icon_mail.png", mime = "image/png")
  icon_www <- dataURI(file = "www/images/icon_www.png", mime = "image/png")
  icon_address <- dataURI(file = "www/images/icon_address.png", mime = "image/png")
  icon_leuchtturm <- dataURI(file = "www/images/icon_leuchtturm.png", mime = "image/png")
  icon_linkedin <- dataURI(file = "www/images/icon_linkedin.png", mime = "image/png")
  icon_xing <- dataURI(file = "www/images/icon_xing.png", mime = "image/png")
  icon_instagram <- dataURI(file = "www/images/icon_instagram.png", mime = "image/png")
  
  
  paste0(
    # Text über der Tabelle
    '<div style="padding-bottom:6px;">',
    "Für  Rückfragen stehen wir gerne zur Verfügung.",
    "<br><br>",
    "Vielen Dank und viele Grüße",
    "</div>",
    
    # Tabelle mit fixer Zeilenhöhe
    '<table style="border-collapse:collapse; font-family:Arial, sans-serif; font-size:10pt;">',
    
    # Zeile 1: leer, Logo rechts mit rowspan über 3 Zeilen
    '<tr style="height:24px;">',
    "<td></td>",
    '<td rowspan="3" style="text-align:right; vertical-align:bottom; padding-left:20px;">',
    '<img src="', logo_eeg, '" style="height:70px;" alt="EEG Logo">',
    "</td>",
    "</tr>",
    
    # Zeile 2: Team
    '<tr style="height:24px;">',
    '<td style="vertical-align:bottom;"><b>Team Energieservice</b></td>',
    "</tr>",
    
    # Zeile 3: Firma
    '<tr style="height:24px;">',
    '<td style="vertical-align:bottom;">Energieservice | EEG Energie- Einkaufs- und Service GmbH</td>',
    "</tr>",
    "</table>",
    
    # --- Tabelle 2: Icons ---
    '<table style="border-collapse:collapse; font-family:Arial, sans-serif; font-size:10pt; margin-top:10px;">',
    "<tr>",
    '<td style="width:30px;"><img src="', icon_mail, '" style="height:14px; vertical-align:middle;"></td>',
    '<td><a href="mailto:dlznn@eeg-energie.de">dlznn@eeg-energie.de</a></td>',
    "</tr>",
    "<tr>",
    '<td style="width:30px;"><img src="', icon_www, '" style="height:14px; vertical-align:middle;"></td>',
    '<td><a href="http://www.eeg-energie.de">www.eeg-energie.de</a></td>',
    "</tr>",
    "<tr>",
    '<td style="width:30px;"><img src="', icon_address, '" style="height:14px; vertical-align:middle;"></td>',
    '<td><a href="https://www.google.com/maps/...">Margarete-Steiff-Str. 1-3, 24558 Henstedt-Ulzburg</a></td>',
    "</tr>",
    "</table>",
    
    # --- Tabelle 3: Slogan + rechtlicher Hinweis ---
    '<table style="border-collapse:collapse; font-family:Arial, sans-serif; font-size:10pt; margin-top:10px;">',
    
    # Zeile 1: Slogan + Icon mit Abstand
    "<tr>",
    "<td><b>Wir sind DER Ansprechpartner – Rund um das Thema Energie!</b></td>",
    '<td style="width:40px; text-align:right; padding-left:18px;">',
    '<img src="', icon_leuchtturm, '" style="height:24px;">',
    "</td>",
    "</tr>",
    
    # Zeile 2: Rechtlicher Hinweis
    '<tr><td colspan="2" style="font-size:8pt; color:#AEAAAA;">',
    "Sitz: Henstedt-Ulzburg, Amtsgericht Kiel HRB 20927KI<br>",
    "Geschäftsführung: Marc Wiederhold (Sprecher) und Matthias Ewert",
    "</td></tr>",
    
    #  Zeile 3: Social Icons mit Abstand
    '<tr><td colspan="2" style="padding-top:10px;">',
    '<table style="border-collapse:collapse;">',
    '<tr style="vertical-align:middle;">',
    '<td style="vertical-align:middle;">',
    '<a href="https://www.linkedin.com/company/eeg-energie">',
    '<img src="', icon_linkedin, '" style="height:25px; vertical-align:middle;"></a>',
    '</td>',
    '<td style="width:32px;"></td>',
    '<td style="vertical-align:middle;">',
    '<a href="https://www.xing.com/pages/eegenergie-einkaufs-undservicegmbh">',
    '<img src="', icon_xing, '" style="height:25px; vertical-align:middle;"></a>',
    '</td>',
    '<td style="width:32px;"></td>',
    '<td style="vertical-align:middle;">',
    '<a href="https://www.instagram.com/eeg.energie/">',
    '<img src="', icon_instagram, '" style="height:25px; vertical-align:middle;"></a>',
    '</td>',
    '</tr>',
    '</table>',
    '</td></tr>'
  )
}


# Senden der Email über Outlook
sende_mail_outlook <- function(empfaenger, betreff = "", body_html = "", direktversand = FALSE) {
  # Starte die Outlook-Anwendung
  outlook <- COMCreate("Outlook.Application")
  
  # Erstelle neue Mail
  mail <- outlook$CreateItem(0)
  
  # Setze die Absenderadresse (im Auftrag von)
  mail[["SentOnBehalfOfName"]] <- "dlznn@eeg-energie.de"
  # Setze den/die Empfänger
  mail[["To"]] <- empfaenger
  # Setze die CC-Adresse
  mail[["CC"]] <- "dlznn@eeg-energie.de"
  # Setze den Betreff der E-Mail
  mail[["Subject"]] <- betreff
  # Setze den HTML-Inhalt des E-Mail-Bodys
  mail[["HTMLBody"]] <- body_html
  
  # Prüfe, ob die E-Mail direkt versendet werden soll
  if (direktversand) {
    # Aktiviere den direkten Versand
    mail$Display()
  } else {
    # Zeige die E-Mail zur Überprüfung an
    mail$Display()
  }
}


# Zeige eine Erfolgsmeldung oben rechts
show_success_toast <- function(message) {
  show_toast(
    # Der Haupttext der Meldung ist der Input
    title = message,
    # Kein weiterer Text unterhalb des Titels
    text = NULL,
    # Typ der Meldung: Erfolg (grünes Symbol)
    type = "success",
    # Dauer der Anzeige in Millisekunden (1,5 Sekunden)
    timer = 1500,
    # Zeige einen Fortschrittsbalken unterhalb der Zeit
    timerProgressBar = TRUE,
    # Positioniere die Meldung oben rechts
    position = "top-end",
    # Automatische Breiteanpassung
    width = NULL,
    # Aktuelle Shiny-Session
    session = getDefaultReactiveDomain()
  )
}

# --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- -#
# ----                         USER INTERFACE                               ----
# --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- -#
# UI-Definition mit bs4Dash
ui <- bs4DashPage(
  ## Browser Titel ----
  title = "LG-NACHFORDERUNG",
  help = NULL,
  dark = NULL,
  
  ## Navbar deaktiviert
  header = bs4DashNavbar(disable = TRUE),
  # Sidebar deaktiviert
  sidebar = bs4DashSidebar(disable = TRUE),  
 
  ## Body ----
  body = bs4DashBody(
    # Füge benutzerdefiniertes CSS hinzu
    includeCSS(file.path(getwd(), "www/styles.css")),
    # Füge benutzerdefiniertes JavaScript hinzu
    includeScript(file.path(getwd(), "www/custom.js")),
    
    # Erstelle eine Tabellen-Struktur für Haupt- und E-Mail-Bereich
    tabsetPanel(
      ### Tab Hauptbereich ----
      tabPanel(
        "Hauptbereich",
        # Füge einen oberen Rand hinzu
        fluidRow(style = "margin-top: 10px;",
          
          #### Filter ----
          column(
            width = 3,
            class = "messstellenbetreiber-box-parent",
            # Umfasst die Filter für Messstellenbetreiber
            box(
              title = "MESSSTELLENBETREIBER",
              width = 12,
              # Verwende eine Monospace-Schriftart
              style = "font-family: monospace;",
              # Box ist nicht einklappbar
              collapsible = FALSE,

              # Erste Zeile: Radiobuttons und Dropdown nebeneinander
              fluidRow(
                column(
                  width = 3,
                  # Auswahl zwischen allen, wMSB und gMSB
                  radioButtons(
                    inputId = "msb_type",
                    label = NULL,
                    choices = c("ALLE" = "alle", "wMSB" = "wMSB", "gMSB" = "gMSB"),
                    selected = "alle",
                    # Buttons untereinander
                    inline = FALSE  
                  )
                ),
                column(
                  width = 9,
                  # Dropdown zur Auswahl von Dateien/Filtern
                  selectInput(
                    inputId = "file_filter",
                    label = NULL,
                    choices = NULL,
                    multiple = FALSE,
                    width = "100%"
                  )
                )
              ),
              
              # Füge einen vertikalen Abstand ein
              br(),
              #### Plot-Ausgabe ----
              uiOutput("plot_ui")
            )
          ),
          
          # Bereich für die E-Mail-Ansicht
          column(
            width = 9,
            box(
              title = "E-Mail Vorschau",
              width = 12,
              # Box ist nicht einklappbar
              collapsible = FALSE,

              
              #### Buttons für die E-Mail-Aktionen ----
              fluidRow(
                column(
                  width = 4,
                  # Button zum Zurückgehen
                  actionButton("back_button", "Zurück", class =  "btn-primary", style = "width: 120px;"),
                  # Button zum Weitergehen
                  actionButton("next_button", "Nächster", class = "btn-primary", style = "width: 120px; margin-left: 10px;")
                ),
                column(
                  width = 8,
                  # Buttons rechtsbündig
                  div(
                    style = "float: right;",
                    # Button zum Löschen
                    actionButton("del_button", "löschen", class = "btn-primary", style = "width: 120px;"),
                    # Button zum Bearbeiten
                    actionButton("edi_button", "bearbeiten", class = "btn-primary", style = "width: 120px; margin-left: 10px;"),
                    # Button zum direkten Senden
                    actionButton("send_button", "direkt Senden", class = "btn-success", style = "width: 120px; margin-left: 10px;")
                  )
                )
              ),
              # Füge einen vertikalen Abstand ein
              br(),
              
              #### Bereich zur Anzeige der E-Mail ----
              # Eingabefeld für den E-Mail-Empfänger
              textInput("email_empfaenger", "Empfänger (E-Mail-Adresse)", width = "100%"),
              # Eingabefeld für den E-Mail-Betreff
              textInput("email_betreff", "Betreff", width = "100%"),
              # Bereich zur Anzeige des HTML-E-Mail-Bodys
              htmlOutput(
                "email_body_html", 
                container = span,
                style = "
                  display: block; 
                  width: 100%; 
                  height: 480px; 
                  border: 1px solid #ced4da; 
                  padding: 10px; 
                  background-color: white; 
                  overflow-y: auto;
                "
              )
            )
          )
        )
      ),
      
      ### Tab E-Mails ----
      tabPanel(
        "E-Mails",
        # Füge einen oberen Rand hinzu
        fluidRow(style = "margin-top: 10px;",
          box(
            title = "E-Mail-Details",
            width = 12,
            # Box ist nicht einklappbar
            collapsible = FALSE,
            class = "white",
            
            #### Buttons ----
            fluidRow(
              column(
                width = 12,
                # Füge einen unteren Abstand hinzu
                style = "margin-bottom: 10px;", 
                # Button zum Hinzufügen eines neuen Eintrags
                actionButton("add_entry", "Neuer Eintrag", class = "btn-primary"),
                # Button zum Bearbeiten eines Eintrags
                actionButton("edit_entry", "Eintrag bearbeiten", class = "btn-primary"),
                # Button zum Löschen eines Eintrags
                actionButton("delete_entry", "Eintrag löschen", class = "btn-primary")
              )
            ),
            
            #### Tabelle ----
            # Tabelle zur Anzeige der E-Mails
            DTOutput("email_table")
          )
        )
      )
    )
  )
)

# --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- -#
# ----                              SERVER                                  ----
# --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- -#
server <- function(input, output, session) {
  
  # E-Mail-Daten laden ----
  # E-Mail-Daten aus Datei laden und reaktiv speichern
  kontaktdaten <- reactiveVal(load_email_data())
  
  # CSV-Liste einlesen ----
  # Liste der CSV-Dateien beim Start einlesen und reaktiv speichern
  csv_files <- reactiveVal(get_csv_files())
  
  # Buttons index initialisieren ----
  # Reaktiver Wert für den aktuellen Index der angezeigten Daten (startet bei 1)
  current_index <- reactiveVal(1)
  
  # Email Body content initialisieren ----
  # Reaktiver Wert für den Inhalt des E-Mail-Bodys (initial leer)
  email_body_content <- reactiveVal("")
  
  # Reaktiver Wert für den globalen Monatszeitraum (als Klartext, initial leer)
  monate_final_plain <- reactiveVal("")
  
  # Reaktiver Wert für die Höhe des Plots (initial 75 Pixel)
  plot_height <- reactiveVal(75)
  
  # --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---#
  # ----                            FILTER                                  ----
  # --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---#
  
  ## MSB Typ ---- 
  filtered_files <- reactive({
    # Stelle sicher, dass die Liste der CSV-Dateien geladen ist
    req(csv_files())
    # Hole die aktuelle Liste der CSV-Dateien
    files <- csv_files()
    
    # Filter nach dem ausgewählten MSB-Typ (wMSB oder gMSB)
    if (input$msb_type != "alle") {
      # Behalte nur Dateien, deren Name den ausgewählten Typ enthält
      files <- files[str_detect(basename(files), fixed(input$msb_type))]
    }
    
    # Filter zusätzlich nach dem ausgewählten MSB-Namen (aus Dropdown)
    if (!is.null(input$file_filter) && input$file_filter != "Alle" && nzchar(input$file_filter)) {
      # Behalte nur Dateien, deren Name den ausgewählten Filter enthält
      files <- files[str_detect(basename(files), fixed(input$file_filter))]
    }
    # Gib die gefilterte Liste der Dateien zurück
    return(files)
  })
  
  
  # Beobachte Änderungen am MSB-Typ-Auswahlfeld
  observeEvent(input$msb_type, {
    # Wenn der MSB-Typ geändert wird, setze die Dateifilter-Auswahl zurück auf "Alle"
    updateSelectInput(session, "file_filter", selected = "Alle")
  })
  
  
  ## Beobachte Änderungen an der Liste der CSV-Dateien und dem MSB-Typ ----
  observe({
    # Stelle sicher, dass die Liste der CSV-Dateien geladen ist
    req(csv_files())
    # Hole die aktuelle Liste der CSV-Dateien
    files <- csv_files()
    
    # Filtere die Dateien basierend auf dem ausgewählten MSB-Typ
    if (input$msb_type != "alle") {
      files <- files[str_detect(basename(files), fixed(input$msb_type))]
    }
    
    # Extrahiere die MSB-Namen aus den gefilterten Dateinamen
    msb_names <- extract_msb_name(basename(files))
    
    # Führe die Aktualisierung des Dropdowns isoliert aus
    isolate({
      # Überprüfe, ob der aktuelle Filterwert noch in den neuen MSB-Namen vorhanden ist
      if (!(input$file_filter %in% msb_names)) {
        # Wenn nicht, setze die Dropdown-Auswahl auf "Alle" und aktualisiere die Optionen
        updateSelectInput(
          session,
          "file_filter",
          choices = c("Alle", sort(unique(msb_names))),
          selected = "Alle"
        )
      } else {
        # Wenn ja, behalte den aktuellen Filterwert bei und aktualisiere nur die Optionen
        updateSelectInput(
          session,
          "file_filter",
          choices = c("Alle", sort(unique(msb_names))),
          selected = input$file_filter
        )
      }
    })
  })
  
  
  ## MSB Name reaktive Liste ----
  filtered_msb_names <- reactive({
    # Stelle sicher, dass die Liste der CSV-Dateien geladen ist
    req(csv_files())
    # Nimm die Dateinamen (ohne Pfad) aller CSV-Dateien
    file_names <- basename(csv_files())
    # Extrahiere alle eindeutigen MSB-Namen aus diesen Dateinamen
    extract_msb_name(file_names) 
  })
  
  # --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---#
  # ----                             PLOT                                   ----
  # --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---#
  
  # Eindeutige MSB-Namen einmalig ermitteln ----
  max_unique_msb <- reactive({
    # Rufe die Liste aller CSV-Dateien ab
    all_csv_files <- get_csv_files()
    # Extrahiere alle MSB-Namen aus allen Dateinamen
    all_msb_names <- unlist(lapply(all_csv_files, function(file) {
      extract_msb_name(basename(file))
    }))
    # Gib die Anzahl der eindeutigen MSB-Namen zurück
    length(unique(all_msb_names))
  })
  
  # Plot Höhe aktualisieren ---
  update_plot_height <- function() {
    # Stelle sicher, dass die gefilterten Dateien vorhanden sind
    req(filtered_files())
    # Hole die Liste der gefilterten Dateien
    files <- filtered_files()
    # Extrahiere die MSB-Namen aus den gefilterten Dateinamen
    msb_names <- extract_msb_name(basename(files))
    # Zähle die eindeutigen MSB-Namen in den gefilterten Dateien
    unique_msb_count <- length(unique(msb_names))
    # Definiere eine Standard-Balkenhöhe für den Plot
    balken_hoehe <- 35
    # Definiere den Platz für die Achsenbeschriftungen
    achsen_platz <- 50
    # Definiere eine maximale Höhe für den Plot, bevor Scrollen nötig wird
    max_plot_hoehe <- 583 
    
    # Berechne die benötigte Höhe basierend auf der Anzahl der Balken
    benoetigte_hoehe <- (unique_msb_count * balken_hoehe) + achsen_platz
    # Wähle die kleinere Höhe zwischen der benötigten und der maximalen Höhe
    tatsaechliche_hoehe <- min(benoetigte_hoehe, max_plot_hoehe)
    
    # Speichere die tatsächliche Höhe in der reaktiven Variable plot_height
    plot_height(tatsaechliche_hoehe)
  }
  
  # Beobachte Änderungen an der gefilterten Dateiliste
  observeEvent(filtered_files(), {
    # Wenn die gefilterte Dateiliste sich ändert, aktualisiere die Plot-Höhe
    update_plot_height()
  })

  # Erstelle einen Container für den Plot mit dynamischer Höhe
  output$plot_ui <- renderUI({
    div(
      class = "plot-container",
      # Setze die Höhe des plotOutput dynamisch basierend auf plot_height()
      plotOutput("plot_msb", height = paste0(plot_height(), "px"))
    )
  })
  
  # Rendert den MSB-Plot
  output$plot_msb <- renderPlot({
    # Stelle sicher, dass die gefilterten Dateien vorhanden sind
    req(filtered_files())
    # Rufe die Funktion zum Erstellen des MSB-Plots mit den gefilterten Dateien auf
    msb_plot(filtered_files())
  })
  
  # --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---#
  # ----                          HAUPTABELLE                               ----
  # --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---#
  
  ## Daten aktualisieren ----
  ### Aktuell ausgewählte Datei ----
  # Reaktiver Wert zur Speicherung der aktuell ausgewählten Datei (initial NULL)
  current_file <- reactiveVal(NULL)
  
  # Beobachte Änderungen in der gefilterten Dateiliste und wähle die erste Datei aus
  observe({
    # Stelle sicher, dass gefilterte Dateien vorhanden sind
    req(filtered_files())
    # Setze die aktuell ausgewählte Datei auf die erste in der gefilterten Liste
    current_file(filtered_files()[1])
  })
  
  ### MSB-Name aus aktueller Datei extrahieren ----
  # Reaktiver Wert, der den MSB-Namen der aktuell ausgewählten Datei enthält
  current_msb <- reactive({
    # Stelle sicher, dass eine aktuelle Datei ausgewählt ist
    req(current_file())
    # Extrahiere den MSB-Namen aus dem Dateinamen
    basename(current_file()) |> extract_msb_name()
  })
  
  ### MSB-Kontakt aus Email ----
  # Reaktiver Wert, der die Kontaktdaten des aktuellen MSB enthält
  current_contact <- reactive({
    # Stelle sicher, dass die Kontaktdaten geladen und ein MSB vorhanden ist
    req(kontaktdaten())
    req(current_msb())
    
    # Filter die Kontaktdaten nach dem aktuellen MSB-Namen
    kontaktdaten() |> 
      filter(MSB == current_msb()) |> 
      # Nimm den ersten Eintrag, falls mehrere gefunden werden
      slice(1)  
  })
  
  ### Eingabefelder aktualisieren ----
  # Beobachte Änderungen in den Kontaktdaten des aktuellen MSB
  observe({
    # Hole die Kontaktdaten des aktuellen MSB
    contact <- current_contact()
    
    # Überprüfe, ob Kontaktdaten gefunden wurden
    if (nrow(contact) > 0) {
      #### Extrahiere den Monatszeitraum aus den aktuellen CSV-Datenn ----
      monate_plain <- extract_monate(current_csv_data())
      # Speichere den Monatszeitraum im reaktiven Wert
      monate_final_plain(monate_plain)  
      
      # Setze die E-Mail-Adresse des Empfängers im Eingabefeld
      updateTextInput(session, "email_empfaenger", value = contact$EMAIL)
      
      # Hole den Dateinamen der aktuellen Datei
      filename <- basename(current_file())
      # Extrahiere die VNB-Kennung aus dem Dateinamen
      vnb <- extract_vnb(filename)
      # Hole den Namen des aktuellen MSB
      msb_name <- current_msb()
      
      # Erstelle den Betreff der E-Mail
      betreff_text <- paste0(
        "Fehlende Lastgangdaten im ", monate_final_plain(), ".    MSB ", msb_name, " für VNB ", vnb
      )
      
      # Aktualisiere das Eingabefeld für den Betreff
      updateTextInput(session, "email_betreff", value = betreff_text)
      
      # Generiere den HTML-Inhalt des E-Mail-Bodys
      body_html <- generate_email_body(current_csv_data(), contact, monate_final_plain()) 
      
      # Rendere den HTML-Inhalt im Ausgabebereich
      output$email_body_html <- renderUI({
        HTML(body_html)
      })
      # Speichere den HTML-Inhalt im reaktiven Wert
      email_body_content(body_html)
    } else {
      # Kein Kontakt
      # Erstelle eine Hinweismeldung für fehlende Kontaktdaten
      missing_msb_html <- paste0(
        "<b>Hinweis:</b> Kein Eintrag in den E-Mail Kontakten für MSB <b>", current_msb(), "</b> vorhanden."
      )
      
      # Leere die Eingabefelder für Empfänger und Betreff
      updateTextInput(session, "email_empfaenger", value = "")
      updateTextInput(session, "email_betreff", value = "")
      
      # Zeige die Hinweismeldung im E-Mail-Body-Bereich an
      output$email_body_html <- renderUI({
        HTML(missing_msb_html)
      })
    }
  })
  
  # TODO hier gehts weiter
  ## Buttons ----
  observe({
    req(filtered_files())
    idx <- current_index()
    
    # Sicherstellen, dass der Index gültig ist
    if (idx >= 1 && idx <= length(filtered_files())) {
      current_file(filtered_files()[idx])
    }
  })
  
  ### Button: Zurück ----
  observeEvent(input$back_button, {
    idx <- current_index()
    
    if (idx > 1) {
      current_index(idx - 1)
    }
  })
  
  ### Button: nächster ----
  observeEvent(input$next_button, {
    idx <- current_index()
    max_idx <- length(filtered_files())
    
    if (idx < max_idx) {
      current_index(idx + 1)
    }
  })
  
  ### Button: löschen ----
  observeEvent(input$del_button, {
    idx <- current_index()
    max_idx <- length(filtered_files())
    
    showNotification(
      "CSV-Datei gelöscht.",
      type = "message",
      duration = 3,
      closeButton = TRUE,
    )
    
    # CSV verschieben
    move_csv_to_done(current_file())
    # --- Update nach Verschieben
    
    update_after_csv_move(session, csv_files, input$msb_type, input$file_filter)
    
    # Index erhöhen
    if (idx < max_idx) {
      current_index(idx + 1)
    }
  })
  
  ### Button: bearbeiten ----
  observeEvent(input$edi_button, {
    idx <- current_index()
    max_idx <- length(filtered_files())
    
    # Neuen Body zusammenbauen
    final_body <- paste0(
      email_body_content(),
      "<br><br>",
      signatur_html()
    )
    
    # E-Mail zur Bearbeitung öffnen
    sende_mail_outlook(
      empfaenger = input$email_empfaenger,
      betreff    = input$email_betreff,
      body_html  = final_body,
      direktversand = FALSE  # <<--- WICHTIG!
    )
    
    showNotification(
      "E-Mail zur Bearbeitung geöffnet.",
      type = "message",
      duration = 3,
      closeButton = TRUE,
    )
    
    # CSV verschieben
    move_csv_to_done(current_file())
    
    # --- Update nach Verschieben
    update_after_csv_move(session, csv_files, input$msb_type, input$file_filter)

    # Index erhöhen
    if (idx < max_idx) {
      current_index(idx + 1)
    }
  })
  
  
  ### Button: Senden ----
  observeEvent(input$send_button, {
    idx <- current_index()
    max_idx <- length(filtered_files())
    
    # Neuen Body zusammenbauen:
    final_body <- paste0(
      email_body_content(),  # der Hauptteil (Anrede, Einleitung, Tabelle)
      "<br><br>",            # kleiner Abstand
      signatur_html()        # DEINE Signatur unten dran
    )
    
    # E-Mail senden
    sende_mail_outlook(
      empfaenger = input$email_empfaenger,
      betreff    = input$email_betreff,
      body_html  = final_body,
      direktversand = TRUE  # <<--- WICHTIG!
    )
    
    showNotification(
      "E-Mail an Outlook übergeben.",
      type = "message",    # oder "default"
      duration = 3,        # Sekunden sichtbar
      closeButton = TRUE,  # Nutzer kann es auch selbst wegklicken
    )
    
    # CSV verschieben
    move_csv_to_done(current_file())
    
    # --- Update nach Verschieben
    update_after_csv_move(session, csv_files, input$msb_type, input$file_filter)

    # Index erhöhen
    if (idx < max_idx) {
      current_index(idx + 1)
    }
  })
  
  ## Tabelle laden ----
  # Inhalt der aktuellen CSV-Datei einlesen
  current_csv_data <- reactive({
    req(current_file())
    
    tryCatch({
      read_delim(
        current_file(),
        delim = ";",
        locale = locale(encoding = "Windows-1252"),
        show_col_types = FALSE
      )
    }, error = function(e) {
      # Falls beim Einlesen ein Fehler auftritt, gib eine leere Tabelle zurück
      tibble(Fehler = "Konnte CSV nicht laden.")
    })
  })
  
  # --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---#
  # ----                            E-MAIL                                  ----
  # --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---#
  
  ## E-Mail Tabelle ----
  output$email_table <- renderDT({
    req(kontaktdaten())  # Sicherstellen, dass Daten geladen sind
    
    datatable(
      kontaktdaten(),
      #escape = FALSE,
      selection = "single",
      options = list(
        # Deaktiviert horizontales Scrollen
        scrollX = FALSE,
        # Setzt die Höhe der Tabelle dynamisch basierend auf der Fensterhöhe
        scrollY = "calc(100vh - 300px)",
        # Zeigt alle Einträge auf einer Seite an
        pageLength = -1,
        # Zeigt nur die Tabelle selbst an (keine Suchfelder etc.)
        dom = 't',
        # Spaltenbreite festlegen
        autoWidth = FALSE, 
        columnDefs = list(
          list(width = '15%', targets = 0), # MSB
          list(width = '15%', targets = 1), # EMAIL
          list(width = '20%', targets = 2), # ANREDE
          list(width = '50%', targets = 3)  # EINLEITUNG
          #list(className = 'dt-body-wrap', targets = 3)
        )
      ),
      rownames = FALSE 
    )
  })
  
  ## E-Mail neu ----
  observeEvent(input$add_entry, {
    showModal(entry_modal("new"))
  })
  
  ## E-Mail neu speichern ----
  observeEvent(input$save_new_entry, {
    data <- kontaktdaten()
    new_row <- tibble(
      MSB        = input$modal_msb,
      EMAIL      = input$modal_email,
      ANREDE     = input$modal_anrede,
      EINLEITUNG = input$modal_einleitung
    )
    updated <- bind_rows(data, new_row)
    kontaktdaten(updated)
    save_email_data(kontaktdaten())
    removeModal()
    show_success_toast("Neuer Eintrag erfolgreich gespeichert!")
  })
  
  
  ## E-Mail bearbeiten ----
  observeEvent(input$edit_entry, {
    sel <- input$email_table_rows_selected
    if (length(sel) == 0) {
      showModal(modalDialog("Bitte erst einen Eintrag auswählen.", easyClose = TRUE))
    } else {
      dat <- kontaktdaten()[sel, ]
      showModal(entry_modal("edit", dat))
    }
  })
  
  ## E-Mail bearbeiten speichern ----
  observeEvent(input$save_edit_entry, {
    sel <- input$email_table_rows_selected
    data <- kontaktdaten()
    data[sel, ] <- tibble(
      MSB        = input$modal_msb,
      EMAIL      = input$modal_email,
      ANREDE     = input$modal_anrede,
      EINLEITUNG = input$modal_einleitung
    )
    kontaktdaten(data)
    save_email_data(kontaktdaten())
    removeModal()
    show_success_toast("Eintrag erfolgreich aktualisiert!")
  })
  
  ## E-Mail löschen ----
  observeEvent(input$delete_entry, {
    # Holen der ausgewählten Zeile
    selected_row <- input$email_table_rows_selected
    
    if (length(selected_row) == 0) {
      showModal(modalDialog(
        title = "Hinweis",
        "Bitte wähle zuerst einen Eintrag aus, den du löschen möchtest.",
        easyClose = TRUE
      ))
    } else {
      # Aktuelle Daten
      data <- kontaktdaten()
      
      # Löschen der ausgewählten Zeile
      updated_data <- data[-selected_row, ]
      
      # Update reactive Value
      kontaktdaten(updated_data)
      
      # Speicherung
      save_email_data(kontaktdaten())
      
      show_success_toast("Eintrag erfolgreich gelöscht!")
    }
  })
  
  
  

  

  

  

  # Wenn die Sitzung endet, beende die App  
  session$onSessionEnded(function() { 
    # Beendet die Shiny-Anwendung wenn der Browser geschlossen wird
    stopApp()
  })
}

# --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --
#                            *** APP STARTEN ***                            ----
# --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --

# Starte die App im Browser; Shiny wählt automatisch einen freien Port
runApp(
  list(ui = ui, server = server),
  launch.browser = TRUE
)
