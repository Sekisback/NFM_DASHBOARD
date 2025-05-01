# --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --
#
# NFM DASHBOARD                                                             ----
#
# Author : Sascha Kornberger
# Datum  : 01.05.2025
# Version: 0.4.0
#
# History:
# 0.4.0  Funktion: Blacklist
# 0.3.0  Bugfix  : Code Optimierung mit Claude
# 0.2.1  Bugfix  : Diagramm balkenhöhe aktualisieren
# 0.2.0  Bugfix  : Dropdown Filter aktualisieren 
#
# --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --
rm(list = ls())

## BENOETIGTE PAKETE ------------------------------------------------------------
# Optionen setzen – unterdrückt manche Dialoge zusätzlich
options(install.packages.check.source = "no")

# Funktion zum effizienten Laden von Paketen
lade_pakete <- function(pakete) {
  # Prüfe welche Pakete fehlen
  fehlende_pakete <- pakete[!pakete %in% installed.packages()[, "Package"]]
  
  # Installiere fehlende Pakete, falls vorhanden
  if (length(fehlende_pakete) > 0) {
    message("Installiere fehlende Pakete: ", paste(fehlende_pakete, collapse = ", "))
    install.packages(
      fehlende_pakete,
      repos = "https://cran.r-project.org",
      quiet = TRUE
    )
  }
  
  # Lade alle Pakete mit Fortschrittsindikator
  for (paket in pakete) {
    suppressPackageStartupMessages(library(paket, character.only = TRUE))
  }
  
  message("Alle Pakete erfolgreich geladen.")
}

# Liste der benötigten Pakete
standard_pakete <- c(
  "shiny", "bs4Dash", "DT", "tidyverse", "janitor",
  "lubridate", "base64enc", "shinyWidgets", "readr", "stringr", "purrr"
)

# Lade Standard-Pakete
lade_pakete(standard_pakete)

# Spezieller Fall für RDCOMClient (nur Windows)
if (.Platform$OS.type == "windows") {
  # Prüfe ob devtools installiert ist
  if (!requireNamespace("devtools", quietly = TRUE)) {
    install.packages("devtools", repos = "https://cran.r-project.org", quiet = TRUE)
    suppressPackageStartupMessages(library(devtools))
  }
  
  # Installiere RDCOMClient falls nötig
  if (!requireNamespace("RDCOMClient", quietly = TRUE)) {
    message("Installiere RDCOMClient von GitHub...")
    tryCatch(
      devtools::install_github("omegahat/RDCOMClient"),
      error = function(e) message("Fehler bei der Installation von RDCOMClient: ", e$message)
    )
  }
  
  # Lade RDCOMClient
  if (requireNamespace("RDCOMClient", quietly = TRUE)) {
    suppressPackageStartupMessages(library(RDCOMClient))
    message("RDCOMClient erfolgreich geladen.")
  } else {
    warning("RDCOMClient konnte nicht geladen werden. Outlook-Funktionen werden nicht verfügbar sein.")
  }
}

## OPTIONS ---------------------------------------------------------------------
# Vermeide Exponentialfunktion
options(scipen = 999)

# Keine Ausgabe in der CLI beim laden von Tidyverse
options(tidiverse.quiet = TRUE)


## Variablen und Konstanten ----------------------------------------------------
# Dateipfad Kontaktdaten
EMAIL_DATA_PATH <- "EMails.RData"

# Ordner
DATA_DIR <- "data"
DONE_DIR <- "done"

# Standardwerte für Modal-Dialoge
DEFAULT_EINLEITUNG <- paste(
  "für die unten aufgeführten Marktlokationen wurden, trotz zuvor versendeter ORDERS,", 
  "noch keine vollständigen Lastgangdaten empfangen.<br>",
  "Wir möchten Sie daher bitten, die Lastgangdaten für den Monat {MONAT} noch einmal zu versenden.<br>",
  "Nutzen Sie hierzu bitte die aktuellen MSCONS-Version und die Ihnen bekannte EDIFACT-Adresse."
)
DEFAULT_ANREDE <- "Sehr geehrte Damen und Herren,"

# Konstanten für Email-Konfiguration
DEFAULT_SENDER <- "dlznn@eeg-energie.de"
DEFAULT_CC <- "dlznn@eeg-energie.de"

# --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- -#
# ----                           FUNKTIONEN                                 ----
# --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- -#


# Laden der E-Mail Kontakten aus RData-Datei 
load_data <- function() {
  data_env <- new.env()
  load(EMAIL_DATA_PATH, envir = data_env)
  list(
    kontaktdaten = data_env$kontaktdaten_df,
    blacklist    = data_env$blacklist_df
  )
}

# Abrufen der CSV-Dateien aus dem Ordner
get_csv_files <- function(directory = DATA_DIR) {
  list.files(directory, pattern = "\\.csv$", full.names = TRUE)
}

# Ereldigte CSV in done verschieben
move_csv_to_done <- function(file_path, done_folder = DONE_DIR) {
  if (!dir.exists(done_folder)) {
    dir.create(done_folder, recursive = TRUE)
  }
  file.rename(file_path, file.path(done_folder, basename(file_path)))
}


# Aktualisiert die CSV-Dateien und den File-Filter nach dem Verschieben einer CSV-Datei.
update_after_csv_move <- function(session, csv_files, current_msb_type, current_file_filter) {
  csv_files(get_csv_files())
}


# Extrahieren des Namens zwischen "MSB" und "an" aus den Dateinamen
extract_msb_name <- function(file_names) {
  str_extract(file_names, "(?<=MSB\\s)(.*?)(?=\\san)") |>
    {\(x) x[!is.na(x)]}()
}


# Extrahieren des MSB-Typs (wMSB oder gMSB) aus den Dateinamen
extract_msb_type <- function(file_names) {
  str_extract(file_names, "(wMSB|gMSB)") |>
    {\(x) x[!is.na(x)]}()
}


# Extrahieren des VNB aus den Dateinamen
extract_vnb <- function(filename) {
  str_extract(filename, "(?<= an )(.*?)(?=\\[|\\.csv)") |>
    str_trim()
}


# Extrahiere den oder die Monate aus dem Datenframe
extract_monate <- function(df) {
  if (nrow(df) == 0) return("")
  
  # Bereinige die Spaltennamen und extrahiere Zeitraum
  monate <- df |> 
    clean_names("all_caps") |>
    select(LUCKE_VON, LUCKE_BIS) |>
    mutate(
      LUCKE_VON = dmy_hm(LUCKE_VON),
      LUCKE_BIS = dmy_hm(LUCKE_BIS)
    ) |>
    # Verwende eine einzige pmap-Operation statt map2
    pmap(function(LUCKE_VON, LUCKE_BIS) {
      seq.Date(
        from = floor_date(LUCKE_VON, "month"), 
        to = floor_date(LUCKE_BIS, "month"), 
        by = "1 month"
      )
    }) |>
    unlist() |>
    as.Date(origin = "1970-01-01") |>
    unique() |>
    sort()  # Stellt sicher, dass die Monate geordnet sind
  
  # Bereite formatierte Monatsnamen und Jahre vor
  monatsnamen <- format(monate, "%B")
  jahre <- format(monate, "%Y")
  
  # Formatiere Ausgabe basierend auf Anzahl der Monate
  if (length(monate) == 1) {
    return(paste0(monatsnamen[1], " ", jahre[1]))
  } else {
    return(paste(
      paste(monatsnamen[-length(monatsnamen)], collapse = ", "),
      "und",
      paste0(monatsnamen[length(monatsnamen)], " ", jahre[length(jahre)])
    ))
  }
}


# Erstelle Plot als Balkendiagramm mit der Menge der Dateien pro MSB
msb_plot <- function(file_names) {
  # Extrahiere und zähle MSB-Namen
  msb_count <- tibble(msb = extract_msb_name(basename(file_names))) |> 
    count(msb, sort = TRUE)
  
  if (nrow(msb_count) == 0) return(ggplot() + theme_void() + labs(title = "Keine MSB Daten verfügbar"))
  
  max_value <- max(msb_count$n, na.rm = TRUE)
  
  # Definiere gemeinsame Themen für den Plot
  theme_settings <- theme_minimal(base_family = "Arial") +
    theme(
      panel.grid.major = element_blank(),
      axis.title = element_blank(),
      plot.title = element_blank()
    )
  
  # Erstelle den ggplot mit einer Pipe
  ggplot(msb_count, aes(x = n, y = reorder(msb, n))) +
    geom_col(fill = "steelblue", width = 0.4, just = 1) +
    geom_vline(xintercept = 0) +
    geom_text(
      aes(x = 0, y = msb, label = str_to_title(msb)),
      hjust = 0, nudge_x = 0.1, vjust = 0, nudge_y = 0.15,
      color = "black", fontface = "bold", size = 5
    ) +
    scale_y_discrete(labels = NULL) +
    scale_x_continuous(
      limits = c(0, ifelse(max_value < 10, 10, NA)),
      breaks = 0:max(10, max_value),
      expand = expansion(mult = c(0, 0.05))
    ) +
    theme_settings
}


# Modal für Neu- und Edit-Einträge
entry_modal <- function(mode = c("new", "edit"), data = NULL) {
  mode <- match.arg(mode)
  
  # Verwende Standardwerte, wenn nicht im Edit-Modus oder bestimmte Werte fehlen
  vals <- list(
    msb        = if (mode == "edit" && !is.null(data$MSB)) data$MSB else "",
    email      = if (mode == "edit" && !is.null(data$EMAIL)) data$EMAIL else "",
    anrede     = if (mode == "edit" && !is.null(data$ANREDE)) data$ANREDE else DEFAULT_ANREDE,
    einleitung = if (mode == "edit" && !is.null(data$EINLEITUNG)) data$EINLEITUNG else DEFAULT_EINLEITUNG
  )
  
  dlg_title <- if (mode == "new") "Neuen Eintrag hinzufügen" else "Eintrag bearbeiten"
  save_id <- if (mode == "new") "save_new_entry" else "save_edit_entry"
  
  modalDialog(
    size = "xl",
    title = dlg_title,
    fluidRow(
      column(4, textInput("modal_msb", "MSB Name", value = vals$msb)),
      column(4, textInput("modal_email", "E-Mail Adresse", value = vals$email)),
      column(4, textInput("modal_anrede", "Anrede", value = vals$anrede))
    ),
    textAreaInput(
      "modal_einleitung", "Einleitungssatz", value = vals$einleitung, 
      width = "100%", rows = 3
    ),
    footer = tagList(
      modalButton("Abbrechen"),
      actionButton(save_id, "Speichern", class = "btn-primary")
    ),
    easyClose = TRUE
  )
}

# --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --
#                             E-MAIL BODY VORLAGEN                          ----
# --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --

# CSV IN HTML-TABELLE 
generate_email_body <- function(df, contact, monate_final_plain, blacklist_meldepunkte = NULL) {
  if (nrow(df) == 0) return("<i>Keine Daten vorhanden.</i>")
  if (is.null(contact) || is.null(contact$ANREDE) || is.null(contact$EINLEITUNG)) {
    return("<i>Kontaktdaten unvollständig.</i>")
  }
  
  # Bereite Daten vor
  df <- df |> clean_names("all_caps")
  
  # --- Blacklist-Filter ---
  blacklist_hinweis <- ""
  if (!is.null(blacklist_meldepunkte) && "MELDEPUNKT" %in% names(df)) {
    anzahl_vorher <- nrow(df)
    df <- df[!(df$MELDEPUNKT %in% blacklist_meldepunkte), ]
    anzahl_nachher <- nrow(df)
    if (anzahl_nachher < anzahl_vorher) {
      blacklist_hinweis <- "<i>Hinweis: Einige Meldepunkte wurden aufgrund der Blacklist entfernt.</i><br><br>"
    }
  }
  
  if (nrow(df) == 0) return(paste0(blacklist_hinweis, "<i>Alle Meldepunkte wurden ausgefiltert.</i>"))
  
  
  # Formatiere Monatstext für HTML
  monate_final_html <- paste0("<b>", monate_final_plain, "</b>")
  # Ersetze Platzhalter im Einleitungstext
  einleitung_text <- str_replace_all(contact$EINLEITUNG, "\\{MONAT\\}", monate_final_html)
  
  # Erstelle HTML für Meldepunkte - mit Fehlerbehandlung falls die Spalte fehlt
  meldepunkte_text <- if ("MELDEPUNKT" %in% names(df)) {
    meldepunkte <- unique(df$MELDEPUNKT)
    paste0("&nbsp;&nbsp;• ", meldepunkte, collapse = "<br>")
  } else {
    "<i>Keine Meldepunkte gefunden</i>"
  }
  
  # Baue E-Mail-Body zusammen
  paste0(
    contact$ANREDE, "<br><br>",
    einleitung_text, "<br><br>",
    "Betroffene Messlokationen:<br>",
    meldepunkte_text
  )
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
sende_mail_outlook <- function(
    empfaenger, 
    betreff = "", 
    body_html = "", 
    direktversand = FALSE,
    cc = DEFAULT_CC,
    sender = DEFAULT_SENDER
) {
  # Prüfe die Eingabeparameter
  if (missing(empfaenger) || is.null(empfaenger) || empfaenger == "") {
    toast("Kein Empfänger angegeben. E-Mail wurde nicht gesendet.", "error")
    return(invisible(FALSE))
  }
  
  tryCatch({
    # Starte die Outlook-Anwendung
    outlook <- COMCreate("Outlook.Application")
    
    # Erstelle neue Mail (0 = olMailItem)
    mail <- outlook$CreateItem(0)
    
    # Setze die E-Mail-Eigenschaften
    mail[["SentOnBehalfOfName"]] <- sender
    mail[["To"]] <- empfaenger
    mail[["CC"]] <- cc
    mail[["Subject"]] <- betreff
    mail[["HTMLBody"]] <- body_html
    
    # Führe die gewünschte Aktion aus
    if (direktversand) {
      mail$Send()
      return(invisible(TRUE))
    } else {
      mail$Display()
      return(invisible(TRUE))
    }
  }, error = function(e) {
    warning(paste("Fehler beim Erstellen der E-Mail:", e$message))
    return(invisible(FALSE))
  })
}


# Zeige eine Erfolgsmeldung oben rechts
toast <- function(message, type = "success") {
  show_toast(
    # Der Haupttext der Meldung ist der Input
    title = message,
    # Kein weiterer Text unterhalb des Titels
    text = NULL,
    # Typ der Meldung: Erfolg (grünes Symbol)
    type = type,
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
  # Browser Titel - kompakte Definition
  title = "LG-NACHFORDERUNG",
  help = NULL,
  dark = NULL,
  
  # Navbar und Sidebar direkt als deaktiviert definieren
  header = bs4DashNavbar(disable = TRUE),
  sidebar = bs4DashSidebar(disable = TRUE),
  
  # Body des Dashboards
  body = bs4DashBody(
    # CSS und JS mit relativen Pfaden einbinden
    tags$head(
      includeCSS("www/styles.css"),
      includeScript("www/custom.js")
    ),
    
    # Haupt-Tabset mit vereinfachter Struktur
    tabsetPanel(
      ### Tab Hauptbereich
      tabPanel(
        "Hauptbereich",
        fluidRow(
          style = "margin-top: 10px;",
          
          # Messstellenbetreiber-Box links
          column(
            width = 3,
            box(
              title = "MESSSTELLENBETREIBER",
              width = NULL, # NULL nutzt volle Breite der Spalte
              style = "font-family: monospace;",
              collapsible = FALSE,
              
              # Filter-Bereich mit MSB-Auswahl
              fluidRow(
                column(
                  width = 3,
                  radioButtons(
                    inputId = "msb_type",
                    label = NULL,
                    choices = c("ALLE" = "alle", "wMSB" = "wMSB", "gMSB" = "gMSB"),
                    selected = "alle"
                  )
                ),
                column(
                  width = 9,
                  selectInput(
                    inputId = "file_filter",
                    label = NULL,
                    choices = NULL,
                    width = "100%"
                  )
                )
              ),
              
              # Plot-Bereich mit einfacherem Abstand
              div(style = "margin-top: 15px;", 
                  uiOutput("plot_ui"))
            )
          ),
          
          # E-Mail-Vorschau rechts
          column(
            width = 9,
            box(
              title = "E-MAIL VORSCHAU",
              width = NULL, # NULL nutzt volle Breite der Spalte
              collapsible = FALSE,
              
              # Navigation und Aktions-Buttons
              fluidRow(
                # Navigations-Buttons links
                column(
                  width = 3,
                  # Stil in eine Variable auslagern für Konsistenz
                  actionButton("back_button", "Zurück", class = "btn-primary nav-btn"),
                  actionButton("next_button", "Nächster", class = "btn-primary nav-btn")
                ),
                
                # MITTE: Hinweistext
                column(
                  width = 5,
                  div(
                    style = "margin-top: 8px; text-align: center;",
                    htmlOutput("blacklist_info")
                  )
                ),
                
                # Aktions-Buttons rechts
                column(
                  width = 4,
                  div(
                    class = "text-right", # Rechtsausrichtung mit BS4-Klasse
                    actionButton("del_button", "löschen", class = "btn-primary action-btn"),
                    actionButton("edi_button", "bearbeiten", class = "btn-primary action-btn"),
                    actionButton("send_button", "direkt Senden", class = "btn-success action-btn")
                  )
                )
              ),
              
              # E-Mail-Formular mit konsistentem Styling
              div(
                style = "margin-top: 15px;",
                textInput("email_empfaenger", "Empfänger (E-Mail-Adresse)", width = "100%"),
                textInput("email_betreff", "Betreff", width = "100%"),
                div(
                  htmlOutput("email_body_html"),
                  class = "email-preview"
                )
              )
            )
          )
        )
      ),
      
      ### Tab E-Mails
      tabPanel(
        "E-Mails",
        fluidRow(
          style = "margin-top: 10px;",
          box(
            width = 12,
            collapsible = FALSE,
            class = "white",
            
            # Aktionsbuttons mit einheitlichem Styling
            div(
              class = "mb-3", # Margin-Bottom mit BS4-Klasse
              actionButton("add_entry", "Neuer Eintrag", class = "btn-primary table-btn"),
              actionButton("edit_entry", "Eintrag bearbeiten", class = "btn-primary table-btn"),
              actionButton("delete_entry", "Eintrag löschen", class = "btn-primary table-btn")
            ),
            
            # Datentabelle
            DTOutput("email_table")
          )
        )
      ),
      
      tabPanel(
        "Blacklist",
        fluidRow(
          style = "margin-top: 10px;",
          box(
            
            width = 12,
            collapsible = FALSE,
            class = "white",
            
            # Aktionsbuttons mit einheitlichem Styling
            div(
              class = "mb-3", # Margin-Bottom mit BS4-Klasse
              actionButton("add_blacklist_entry", "Neu"),
              actionButton("edit_blacklist_entry", "Bearbeiten"),
              actionButton("delete_blacklist_entry", "Löschen")
            ),
            
            # Datentabelle
            DTOutput("blacklist_table")
          )
        )
      )
    )
  )
)

server <- function(input, output, session) {
  
  # Initialdaten laden
  daten <- load_data()
  
  # Reaktive Werte initialisieren
  kontaktdaten <- reactiveVal(daten$kontaktdaten)
  blacklist    <- reactiveVal(daten$blacklist)
  
  # Speichern der RData-Datei
  save_data <- function() {
    kontaktdaten_df <- kontaktdaten()
    blacklist_df <- blacklist()
    save(kontaktdaten = kontaktdaten_df, blacklist = blacklist_df, file = EMAIL_DATA_PATH)
  }
  
  # Lädt CSV-Dateien einmalig und speichert sie reaktiv
  csv_files <- reactiveVal(get_csv_files())
  
  # Speichert den aktuellen Index der angezeigten Daten (startet bei 1)
  current_index <- reactiveVal(1)
  
  # Speichert den Inhalt des E-Mail-Bodys (initial leer)
  email_body_content <- reactiveVal("")
  
  # Speichert den Monatszeitraum als Text (initial leer)
  monate_final_plain <- reactiveVal("")
  
  # Speichert die Höhe des Plots in Pixel (initial 75px)
  plot_height <- reactiveVal(75)
  
  # Speichert die aktuell ausgewählte Datei (initial NULL)
  current_file <- reactiveVal(NULL)
  
  # Erstellt einen Cache für extrahierte MSB-Namen um wiederholte Berechnungen zu vermeiden
  msb_name_cache <- reactiveValues(names = list())
  
  blacklist_info_text <- reactiveVal("")
  
  # --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---#
  # ----                        HILFSFUNKTIONEN                            ----
  # --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---#
  
  # Extrahiert den MSB-Namen aus einem Dateinamen mit Caching
  get_cached_msb_name <- function(filename) {
    # Prüfe ob der MSB-Name bereits im Cache existiert
    if (is.null(msb_name_cache$names[[filename]])) {
      # Falls nicht, berechne ihn und speichere ihn im Cache
      msb_name_cache$names[[filename]] <- extract_msb_name(filename)
    }
    # Gib den MSB-Namen aus dem Cache zurück
    return(msb_name_cache$names[[filename]])
  }
  
  # Funktion zum Aktualisieren nach Verschieben einer CSV-Datei
  update_after_csv_move <- function() {
    # Aktualisiere die Liste der CSV-Dateien
    csv_files(get_csv_files())
    
    # Setze den Cache für MSB-Namen zurück
    msb_name_cache$names <- list()
    
    # Überprüfe, ob noch gefilterte Dateien vorhanden sind
    if (length(filtered_files()) == 0) {
      # Wenn keine Dateien mehr vorhanden sind, zeige eine Benachrichtigung
      showNotification(
        "Keine weiteren Dateien entsprechen den Filterkriterien.",
        type = "warning",
        duration = 5
      )
    }
  }
  
  # --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---#
  # ----                            FILTER                                  ----
  # --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---#
  
  # Filtert die CSV-Dateien nach den ausgewählten Kriterien
  filtered_files <- reactive({
    # Stelle sicher, dass die Liste der CSV-Dateien geladen ist
    req(csv_files())
    files <- csv_files()
    
    # Filter nach MSB-Typ (wMSB oder gMSB)
    if (input$msb_type != "alle") {
      files <- files[str_detect(basename(files), fixed(input$msb_type))]
    }
    
    # Filter nach MSB-Namen aus Dropdown
    if (!is.null(input$file_filter) && input$file_filter != "Alle" && nzchar(input$file_filter)) {
      files <- files[str_detect(basename(files), fixed(input$file_filter))]
    }
    
    return(files)
  })
  
  # Extrahiert eindeutige MSB-Namen aus den gefilterten Dateien
  filtered_msb_names <- reactive({
    req(filtered_files())
    # Extrahiere MSB-Namen mit Cache-Funktion
    sapply(basename(filtered_files()), get_cached_msb_name) |> unique()
  })
  
  # Setzt die Dateifilter-Auswahl zurück, wenn sich der MSB-Typ ändert
  observeEvent(input$msb_type, {
    updateSelectInput(session, "file_filter", selected = "Alle")
  })
  
  # Aktualisiert die Dropdown-Optionen für MSB-Namen
  observe({
    req(csv_files())
    files <- csv_files()
    
    if (input$msb_type != "alle") {
      files <- files[str_detect(basename(files), fixed(input$msb_type))]
    }
    
    # Extrahiere MSB-Namen mit Cache-Funktion für bessere Performance
    msb_names <- sapply(basename(files), get_cached_msb_name) |> unique() |> sort()
    
    # Aktualisiere das Dropdown-Menü, behalte die Auswahl wenn möglich
    isolate({
      current_filter <- input$file_filter
      updateSelectInput(
        session,
        "file_filter",
        choices = c("Alle", msb_names),
        selected = if (current_filter %in% c("Alle", msb_names)) current_filter else "Alle"
      )
    })
  })
  
  # --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---#
  # ----                             PLOT                                   ----
  # --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---#
  
  # Aktualisiert die Plot-Höhe basierend auf der Anzahl der MSBs
  update_plot_height <- function() {
    req(filtered_files())
    
    # Zähle eindeutige MSB-Namen in gefilterten Dateien
    unique_msb_count <- length(filtered_msb_names())
    
    # Definiere Parameter für die Plotgröße
    balken_hoehe <- 25
    achsen_platz <- 50
    max_plot_hoehe <- 583
    
    # Berechne optimale Plot-Höhe
    benoetigte_hoehe <- (unique_msb_count * balken_hoehe) + achsen_platz
    tatsaechliche_hoehe <- min(benoetigte_hoehe, max_plot_hoehe)
    
    plot_height(tatsaechliche_hoehe)
  }
  
  # Aktualisiere Plot-Höhe wenn sich die gefilterte Dateiliste ändert
  observeEvent(filtered_files(), {
    update_plot_height()
  })
  
  # Erstellt einen Container für den Plot mit dynamischer Höhe
  output$plot_ui <- renderUI({
    div(
      class = "plot-container",
      plotOutput("plot_msb", height = paste0(plot_height(), "px"))
    )
  })
  
  # Rendert den MSB-Plot
  output$plot_msb <- renderPlot({
    req(filtered_files())
    msb_plot(filtered_files())
  })
  
  # --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---#
  # ----                          HAUPTABELLE                               ----
  # --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---#
  
  # Beobachte Änderungen in der gefilterten Dateiliste und wähle die erste Datei
  observe({
    req(filtered_files())
    if (length(filtered_files()) > 0) {
      current_file(filtered_files()[1])
      current_index(1)  # Setze den Index zurück
    } else {
      current_file(NULL)
    }
  })
  
  # Extrahiert den MSB-Namen der aktuell ausgewählten Datei
  current_msb <- reactive({
    req(current_file())
    get_cached_msb_name(basename(current_file()))
  })
  
  # Findet die Kontaktdaten des aktuellen MSB
  current_contact <- reactive({
    req(kontaktdaten(), current_msb())
    
    kontaktdaten() |> 
      filter(MSB == current_msb()) |> 
      slice(1)  # Nimm den ersten Eintrag, falls mehrere gefunden werden
  })
  
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
      # Bei Fehler gib eine informative Tabelle zurück
      tibble(Fehler = paste("Fehler beim Laden der CSV:", e$message))
    })
  })
  
  # Aktualisiere UI-Elemente basierend auf Kontaktdaten
  observe({
    contact <- current_contact()
    
    # Prüfe ob aktuelle CSV-Daten verfügbar sind
    csv_data_available <- !is.null(current_csv_data()) && 
      !"Fehler" %in% names(current_csv_data())
    
    if (nrow(contact) > 0 && csv_data_available) {
      # Extrahiere den Monatszeitraum aus den CSV-Daten
      monate_plain <- extract_monate(current_csv_data())
      monate_final_plain(monate_plain)
      
      # Setze E-Mail-Empfänger
      updateTextInput(session, "email_empfaenger", value = contact$EMAIL)
      
      # Extrahiere Informationen für den Betreff
      filename <- basename(current_file())
      vnb <- extract_vnb(filename)
      msb_name <- current_msb()
      
      betreff_text <- paste0(
        "Fehlende Lastgangdaten im ", monate_final_plain(), ".    MSB ", msb_name, " für VNB ", vnb
      )
      updateTextInput(session, "email_betreff", value = betreff_text)
      
      # --- NEU: Blacklist prüfen ---
      daten_raw <- current_csv_data() |> janitor::clean_names("all_caps")
      blacklist_mps <- trimws(toupper(blacklist()$MELDEPUNKT))  # sicherheitshalber
      
      # Prüfen, ob ein Meldepunkt aus der CSV auf der Blacklist steht
      if ("MELDEPUNKT" %in% names(daten_raw)) {
        mp_csv <- trimws(toupper(daten_raw$MELDEPUNKT))
        hat_blacklist_treffer <- any(mp_csv %in% blacklist_mps)
      } else {
        hat_blacklist_treffer <- FALSE
      }
      
      # Hinweistext setzen (nur wenn es Treffer gibt)
      if (hat_blacklist_treffer) {
        blacklist_info_text("<b>Hinweis:</b> Meldepunkte auf der Blacklist wurden aus der E-Mail entfernt.")
      } else {
        blacklist_info_text("")
      }
      
      # Generiere E-Mail-Body mit Blacklist-Filter
      body_html <- generate_email_body(
        df = daten_raw,
        contact = contact,
        monate_final_plain = monate_final_plain(),
        blacklist_meldepunkte = blacklist_mps
      )
      
      output$email_body_html <- renderUI({
        HTML(body_html)
      })
      email_body_content(body_html)
      
    } else {
      # Kein Kontakt oder CSV-Fehler
      missing_reason <- if (nrow(contact) == 0) {
        paste0("Kein Eintrag in den E-Mail Kontakten für MSB <b>", current_msb(), "</b> vorhanden.")
      } else {
        "CSV-Datei konnte nicht eingelesen werden."
      }
      
      missing_html <- paste0("<b>Hinweis:</b> ", missing_reason)
      
      updateTextInput(session, "email_empfaenger", value = "")
      updateTextInput(session, "email_betreff", value = "")
      
      output$email_body_html <- renderUI({
        HTML(missing_html)
      })
      
      blacklist_info_text("")  # Hinweis leeren, wenn Daten fehlen
    }
  })
  
  output$blacklist_info <- renderUI({
    if (nzchar(blacklist_info_text())) {
        HTML(blacklist_info_text())
    }
  })
  
  # --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---#
  # ----                          BUTTONS                                   ----
  # --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---#
  
  # Aktualisiert die ausgewählte Datei basierend auf dem aktuellen Index
  observe({
    req(filtered_files())
    idx <- current_index()
    
    # Stelle sicher, dass der Index gültig ist
    if (idx >= 1 && idx <= length(filtered_files())) {
      current_file(filtered_files()[idx])
    }
  })
  
  # Button: Zurück
  observeEvent(input$back_button, {
    idx <- current_index()
    
    if (idx > 1) {
      current_index(idx - 1)
    }
  })
  
  # Button: Nächster
  observeEvent(input$next_button, {
    idx <- current_index()
    max_idx <- length(filtered_files())
    
    if (idx < max_idx) {
      current_index(idx + 1)
    }
  })
  
  # Gemeinsame Funktion für das Behandeln von Datei-Aktionen (Löschen/E-Mail)
  handle_file_action <- function(action_type, show_message) {
    idx <- current_index()
    max_idx <- length(filtered_files())
    
    # Zeige Benachrichtigung
    showNotification(
      show_message,
      type = "message",
      duration = 3,
      closeButton = TRUE
    )
    
    # CSV verschieben
    move_csv_to_done(current_file())
    
    # Update nach Verschieben
    update_after_csv_move()
    
    # Index anpassen für nächste Datei
    if (length(filtered_files()) > 0 && idx <= length(filtered_files())) {
      # Behalte aktuellen Index wenn möglich
      current_index(idx)
    } else if (length(filtered_files()) > 0) {
      # Wenn der aktuelle Index zu hoch ist, gehe zum letzten Element
      current_index(length(filtered_files()))
    }
  }
  
  # Button: Löschen
  observeEvent(input$del_button, {
    handle_file_action("delete", "CSV-Datei gelöscht.")
  })
  
  # Button: E-Mail bearbeiten
  observeEvent(input$edi_button, {
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
      direktversand = FALSE
    )
    
    handle_file_action("edit", "E-Mail zur Bearbeitung geöffnet.")
  })
  
  # Button: E-Mail senden
  observeEvent(input$send_button, {
    # Neuen Body zusammenbauen
    final_body <- paste0(
      email_body_content(),
      "<br><br>",
      signatur_html()
    )
    
    # E-Mail senden
    sende_mail_outlook(
      empfaenger = input$email_empfaenger,
      betreff    = input$email_betreff,
      body_html  = final_body,
      direktversand = TRUE
    )
    
    handle_file_action("send", "E-Mail an Outlook übergeben.")
  })
  
  # --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---#
  # ----                            E-MAIL                                  ----
  # --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---#
  
  # E-Mail-Tabelle rendern
  output$email_table <- renderDT({
    req(kontaktdaten())
    
    datatable(
      kontaktdaten(),
      selection = "single",
      options = list(
        scrollX = FALSE,
        scrollY = "calc(100vh - 300px)",
        pageLength = -1,
        dom = 't',
        autoWidth = FALSE, 
        columnDefs = list(
          list(width = '15%', targets = 0), # MSB
          list(width = '15%', targets = 1), # EMAIL
          list(width = '20%', targets = 2), # ANREDE
          list(width = '50%', targets = 3)  # EINLEITUNG
        )
      ),
      rownames = FALSE 
    )
  })
  
  # Datahandling - Management
  handle_data_action <- function(data_type, action_type) {
    if (data_type == "email") {
      data <- kontaktdaten()
      sel <- input$email_table_rows_selected
    } else {
      data <- blacklist()
      sel <- input$blacklist_table_rows_selected
    }
    
    if (action_type == "new") {
      new_row <- if (data_type == "email") {
        tibble(
          MSB        = input$modal_msb,
          EMAIL      = input$modal_email,
          ANREDE     = input$modal_anrede,
          EINLEITUNG = input$modal_einleitung
        )
      } else {
        tibble(
          MELDEPUNKT  = input$modal_meldepunkt,
          MSB         = input$modal_msb,
          INFORMATION = input$modal_info
        )
      }
      updated <- bind_rows(data, new_row)
      
    } else if (action_type == "edit") {
      if (length(sel) == 0) return()
      
      if (data_type == "email") {
        data[sel, ] <- tibble(
          MSB        = input$modal_msb,
          EMAIL      = input$modal_email,
          ANREDE     = input$modal_anrede,
          EINLEITUNG = input$modal_einleitung
        )
      } else {
        data[sel, ] <- tibble(
          MELDEPUNKT  = input$modal_meldepunkt,
          MSB         = input$modal_msb,
          INFORMATION = input$modal_info
        )
      }
      updated <- data
      
    } else if (action_type == "delete") {
      if (length(sel) == 0) return()
      updated <- data[-sel, ]
    }
    
    # Rückspeichern
    if (data_type == "email") {
      kontaktdaten(updated)
    } else {
      blacklist(updated)
    }
    
    save_data()
    toast(paste("Eintrag", switch(action_type, new = "angelegt", edit = "bearbeitet", delete = "gelöscht"), "!"))
    removeModal()
  }
  
  # E-Mail neu
  observeEvent(input$add_entry, {
    showModal(entry_modal("new"))
  })
  
  # E-Mail neu speichern
  observeEvent(input$save_new_entry, {
    handle_data_action("email", "new")
  })
  
  # E-Mail bearbeiten
  observeEvent(input$edit_entry, {
    sel <- input$email_table_rows_selected
    if (length(sel) == 0) {
      showModal(modalDialog("Bitte erst einen Eintrag auswählen.", easyClose = TRUE))
    } else {
      dat <- kontaktdaten()[sel, ]
      showModal(entry_modal("edit", dat))
    }
  })
  
  # E-Mail bearbeitet speichern
  observeEvent(input$save_edit_entry, {
    handle_data_action("email", "edit")
  })
  
  # E-Mail löschen
  observeEvent(input$delete_entry, {
    if (length(input$email_table_rows_selected) == 0) {
      showModal(modalDialog(
        title = "Hinweis",
        "Bitte wähle zuerst einen Eintrag aus, den du löschen möchtest.",
        easyClose = TRUE
      ))
    } else {
      handle_data_action("email", "delete")
    }
  })
  
  # --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---#
  # ----                           BLACKLIST                                ----
  # --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---#
  
  output$blacklist_table <- renderDT({
    req(blacklist())
    
    datatable(
      blacklist(),
      selection = "single",
      options = list(
        scrollX = FALSE,
        scrollY = "calc(100vh - 300px)",
        pageLength = -1,
        dom = 't',
        autoWidth = FALSE, 
        columnDefs = list(
          list(width = '15%', targets = 0), 
          list(width = '15%', targets = 1), 
          list(width = '70%', targets = 2)
        )
      ),
      rownames = FALSE 
    )
  })
  
  entry_modal_blacklist <- function(mode = "new", data = NULL) {
    modalDialog(
      textInput("modal_meldepunkt", "Meldepunkt", value = if (!is.null(data)) data$MELDEPUNKT else ""),
      textInput("modal_msb", "MSB", value = if (!is.null(data)) data$MSB else ""),
      textAreaInput("modal_info", "Info", value = if (!is.null(data)) data$INFORMATION else "", width = "100%"),
      footer = tagList(
        modalButton("Abbrechen"),
        actionButton(
          if (mode == "edit") "save_edit_blacklist_entry" else "save_new_blacklist_entry",
          "Speichern"
        )
      ),
      easyClose = TRUE
    )
  }
  
  # Neuen Blacklist-Eintrag anlegen
  observeEvent(input$add_blacklist_entry, {
    showModal(entry_modal_blacklist("new"))
  })
  
  observeEvent(input$save_new_blacklist_entry, {
    handle_data_action("blacklist", "new")
  })
  
  # Blacklist-Eintrag bearbeiten
  observeEvent(input$edit_blacklist_entry, {
    sel <- input$blacklist_table_rows_selected
    if (length(sel) == 0) {
      showModal(modalDialog("Bitte Eintrag auswählen.", easyClose = TRUE))
    } else {
      dat <- blacklist()[sel, ]
      showModal(entry_modal_blacklist("edit", dat))
    }
  })
  
  observeEvent(input$save_edit_blacklist_entry, {
    handle_data_action("blacklist", "edit")
  })
  
  # Blacklist-Eintrag löschen
  observeEvent(input$delete_blacklist_entry, {
    sel <- input$blacklist_table_rows_selected
    if (length(sel) == 0) {
      showModal(modalDialog("Bitte erst einen Eintrag auswählen.", easyClose = TRUE))
    } else {
      handle_data_action("blacklist", "delete")
    }
  })
  

  
  
  
  
  # Beende die App, wenn die Sitzung endet
  session$onSessionEnded(function() {
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
