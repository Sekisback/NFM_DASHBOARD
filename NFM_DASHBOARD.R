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
# TODO Ordner auf Laufwerk mappen
get_csv_files <- function(directory = "data") {
  # Hole alle CSV-Dateien im Ordner
  files <- list.files(directory, pattern = "\\.csv$", full.names = TRUE)
  # TODO Debugging Ausgabe entfernen
  # print(files)
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
  new_files <- get_csv_files()
  csv_files(new_files)
}


# Extrahieren des Namens zwischen "MSB" und "an" aus den Dateinamen
extract_msb_name <- function(file_names) {
  # Verwende regex, um den Text zwischen "MSB" und "an" zu extrahieren
  msb_names <- str_extract(file_names, "(?<=MSB\\s)(.*?)(?=\\san)")
  # Entferne NA-Werte
  return(msb_names[!is.na(msb_names)])  
}

# Extrahieren des MSB-Typs (wMSB oder gMSB) aus den Dateinamen
extract_msb_type <- function(file_names) {
  # Verwende regex, um nach 'wMSB' oder 'gMSB' zu suchen
  msb_types <- str_extract(file_names, "(wMSB|gMSB)")
  # Entferne NA-Werte
  return(msb_types[!is.na(msb_types)]) 
}

# Extrahieren des VNB aus den Dateinamen
extract_vnb <- function(filename) {
  # Extrahiere Text nach " an " und vor [ oder .csv
  vnb <- str_extract(filename, "(?<= an )(.*?)(?=\\[|\\.csv)")
  vnb <- str_trim(vnb)  # Whitespace entfernen
  return(vnb)
}


extract_monate_final <- function(df) {
  if (nrow(df) == 0) return("")
  
  df <- df |> janitor::clean_names("all_caps")
  
  df_relevant <- df |>
    select(LUCKE_VON, LUCKE_BIS) |>
    mutate(
      LUCKE_VON = dmy_hm(LUCKE_VON),
      LUCKE_BIS = dmy_hm(LUCKE_BIS)
    )
  
  monate <- purrr::map2(
    df_relevant$LUCKE_VON,
    df_relevant$LUCKE_BIS,
    ~ seq.Date(from = floor_date(.x, "month"), to = floor_date(.y, "month"), by = "1 month")
  ) |>
    unlist() |>
    as.Date(origin = "1970-01-01") |>
    unique()
  
  monatsnamen <- format(monate, "%B")
  jahre <- format(monate, "%Y")
  
  if (length(monate) == 1) {
    monate_final <- paste0(monatsnamen[1], " ", jahre[1])
  } else {
    monate_final <- paste(
      paste(monatsnamen[-length(monatsnamen)], collapse = ", "),
      "und",
      paste0(monatsnamen[length(monatsnamen)], " ", jahre[length(jahre)])
    )
  }
  
  return(monate_final)
}



# Erstelle Plot als Balkendiagramm mit der Menge der Dateien pro MSB
msb_plot <- function(file_names) {
  # MSB-Namen extrahieren
  msb_names <- extract_msb_name(basename(file_names))
  
  # Zähle Vorkommen
  msb_count <- tibble(msb = msb_names) |> 
    count(msb, sort = TRUE)
  
  # Maximalwert für die X-Achse
  max_value <- max(msb_count$n, na.rm = TRUE)
  
  # Plot erzeugen
  ggplot(msb_count, aes(x = n, y = reorder(msb, n))) +
    geom_col(fill = "steelblue", width = 0.4, just = 1) +
    geom_vline(xintercept = 0) +
    geom_text(
      data = msb_count,
      mapping = aes(
        x = 0,
        y = msb,
        label = str_to_title(msb)
      ),
      hjust = 0,
      vjust = 0,
      nudge_y = 0.15,
      nudge_x = 0.1,
      color = "black",
      fontface = "bold",
      size = 5
    )+
    theme_minimal(base_family = "Arial") +
    labs(
      x = element_blank(),
      y = element_blank(),
      title = element_blank(),
    ) +
    theme(
      panel.grid.major = element_blank(),
    )+
    scale_y_discrete(labels = NULL)+
    scale_x_continuous(
      limits = c(0, ifelse(max_value < 10, 10, NA)),
      breaks = 0:max(10, max_value),
      expand = expansion(mult = c(0, 0.05))
    )
}


# Modal für Neu- und Edit-Einträge
entry_modal <- function(mode = c("new", "edit"), data = NULL) {
  mode <- match.arg(mode)
  # Titel je nach Mode
  dlg_title <- if (mode == "new") {
    "Neuen Eintrag hinzufügen"
  } else {
    "Eintrag bearbeiten"
  }
  # Default Einleitung
  default_einleitung <- "für die unten aufgeführten Marktlokationen wurden, trotz zuvor versendeter ORDERS, noch keine vollständigen Lastgangdaten empfangen.<br>Wir möchten Sie daher bitten, die Lastgangdaten für den Monat {MONAT} noch einmal zu versenden.<br>Nutzen Sie hierzu bitte die aktuellen MSCONS-Version und die Ihnen bekannte EDIFACT-Adresse."
  # Defautl Anrede
  default_anrede <- "Sehr geehrte Damen und Herren,"
  
  # Default-Werte oder aus data übernehmen
  vals <- list(
    msb        = if (mode == "edit") data$MSB        else "",
    email      = if (mode == "edit") data$EMAIL      else "",
    anrede     = if (mode == "edit") data$ANREDE     else default_anrede,
    einleitung = if (mode == "edit") data$EINLEITUNG else default_einleitung
  )
  # Welcher Save-Button soll erzeugt werden?
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
      "modal_einleitung", 
      "Einleitungssatz", 
      value = vals$einleitung, 
      width = "100%",     # Breite setzen
      rows = 3             # Mindestens 4 Zeilen hoch
    ),
    footer = tagList(
      modalButton("Abbrechen"),
      actionButton(save_id, "Speichern", class = "btn-primary")
    ),
    easyClose = TRUE,
  )
}

# --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --
#                             E-MAIL BODY VORLAGEN                          ----
# --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --

# CSV IN HTML-TABELLE 
generate_email_body <- function(df, contact, monate_final_plain) {
  if (nrow(df) == 0) return("<i>Keine Daten vorhanden.</i>")
  
  df <- df |> clean_names("all_caps")
  
  # Anrede
  anrede_text <- contact$ANREDE
  if (is.null(anrede_text) || is.na(anrede_text)) {
    anrede_text <- "Sehr geehrte Damen und Herren,"
  }
  
  # Einleitung
  einleitung_text <- contact$EINLEITUNG
  if (is.null(einleitung_text) || is.na(einleitung_text)) {
    einleitung_text <- "Einleitungstext fehlt"
  }
  
  # Monate fett
  monate_final_html <- paste0("<b>", monate_final_plain, "</b>")
  
  # {MONAT} ersetzen
  einleitung_text <- str_replace_all(einleitung_text, "\\{MONAT\\}", monate_final_html)
  
  # Meldepunkte
  meldepunkte <- df$MELDEPUNKT |> unique()
  meldepunkte_text <- paste0("&nbsp;&nbsp;• ", meldepunkte, collapse = "<br>")
  
  # Zusammenbauen
  email_body <- paste0(
    anrede_text, "<br><br>",
    einleitung_text, "<br><br>",
    "Betroffene Messlokationen:<br>",
    meldepunkte_text
  )
  
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
  # Starte Outlook
  outlook <- COMCreate("Outlook.Application")
  
  # Erstelle neue Mail
  mail <- outlook$CreateItem(0)
  
  # Fülle Felder
  mail[["SentOnBehalfOfName"]] <- "dlznn@eeg-energie.de"
  mail[["To"]] <- empfaenger
  mail[["CC"]] <- "dlznn@eeg-energie.de"
  mail[["Subject"]] <- betreff
  mail[["HTMLBody"]] <- body_html
  
  # TODO direktverseasand freigeben
  if (direktversand) {
    #mail$Send()
    mail$Display()
  } else {
    mail$Display()
  }

}

# Erfolgsmeldung oben rechts
show_success_toast <- function(message) {
  show_toast(
    title = message,
    text = NULL,
    type = "success",
    timer = 1500,
    timerProgressBar = TRUE,
    position = "top-end",
    width = NULL,
    session = shiny::getDefaultReactiveDomain()
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
    includeCSS(file.path(getwd(), "www/styles.css")),
    includeScript(file.path(getwd(), "www/custom.js")),
    
    tabsetPanel(
      ### Tab Hauptbereich ----
      tabPanel(
        "Hauptbereich",
        fluidRow(style = "margin-top: 10px;",
          
          #### Filter ----
          column(
            width = 3,
            class = "messstellenbetreiber-box-parent",
            box(
              title = "MESSSTELLENBETREIBER",
              width = 12,
              style = "font-family: monospace;",
              collapsible = FALSE,

              # Erste Zeile: ALLE-Button und Dropdown nebeneinander
              fluidRow(
                column(
                  width = 3,
                  radioButtons(
                    inputId = "msb_type",
                    label = NULL,
                    choices = c("ALLE" = "alle", "wMSB" = "wMSB", "gMSB" = "gMSB"),
                    selected = "alle",
                    inline = FALSE  # untereinander, oder TRUE wenn gewünscht
                  )
                ),
                column(
                  width = 9,
                  selectInput(
                    inputId = "file_filter",
                    label = NULL,
                    choices = NULL,
                    multiple = FALSE,
                    width = "100%"
                  )
                )
              ),
              
              # Abstand
              br(),
              #### Plot ----
              uiOutput("plot_ui")
            )
          ),
          
          # E-Mail Ansicht 
          column(
            width = 9,
            box(
              title = "E-Mail Vorschau",
              width = 12,
              collapsible = FALSE,
              
              #### Buttons ----
              fluidRow(
                column(
                  width = 4,
                  actionButton("back_button", "Zurück", class =  "btn-primary", style = "width: 120px;"),
                  actionButton("next_button", "Nächster", class = "btn-primary", style = "width: 120px; margin-left: 10px;")
                ),
                column(
                  width = 8,
                  div(
                    style = "float: right;",
                    actionButton("del_button", "löschen", class = "btn-primary", style = "width: 120px;"),
                    actionButton("edi_button", "bearbeiten", class = "btn-primary", style = "width: 120px; margin-left: 10px;"),
                    actionButton("send_button", "direkt Senden", class = "btn-success", style = "width: 120px; margin-left: 10px;")
                  )
                )
              ),
              br(),
              
              #### E-Mail Ansicht ----
              # Empfänger
              textInput("email_empfaenger", "Empfänger (E-Mail-Adresse)", width = "100%"),
              # Betreff
              textInput("email_betreff", "Betreff", width = "100%"),
              # Body
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
        fluidRow(style = "margin-top: 10px;",
          box(
            title = "E-Mail-Details",
            width = 12,
            collapsible = FALSE,
            
            #### Buttons ----
            fluidRow(
              column(
                width = 12,
                style = "margin-bottom: 10px;", # Kleiner Abstand unter Buttons
                actionButton("add_entry", "Neuer Eintrag", class = "btn-primary"),
                actionButton("edit_entry", "Eintrag bearbeiten", class = "btn-primary"),
                actionButton("delete_entry", "Eintrag löschen", class = "btn-primary")
              )
            ),
            
            #### Tabelle ----
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
  kontaktdaten <- reactiveVal(load_email_data())
  
  # CSV-Liste einlesen ----
  csv_files <- reactiveVal(get_csv_files())
  
  # Buttons index initialisieren ----
  current_index <- reactiveVal(1)
  
  # Email Body content initialisieren ----
  email_body_content <- reactiveVal("")
  
  # Globale Variable für den aktuellen Monatszeitraum
  monate_final_plain <- reactiveVal("")
  
  # Plot höhe
  plot_height <- reactiveVal(75)
  
  # --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---#
  # ----                            FILTER                                  ----
  # --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---#
  
  ## MSB Typ ---- 
  filtered_files <- reactive({
    req(csv_files())  # Sicherstellen, dass CSVs da sind
    files <- csv_files()
    
    # Filter nach MSB Typ (wMSB / gMSB)
    if (input$msb_type != "alle") {
      files <- files[str_detect(basename(files), fixed(input$msb_type))]
    }
    
    # Filter nach MSB Name (Dropdown)
    if (!is.null(input$file_filter) && input$file_filter != "Alle" && nzchar(input$file_filter)) {
      files <- files[str_detect(basename(files), fixed(input$file_filter))]
    }
    
    files
  })
  

  observeEvent(input$msb_type, {
    updateSelectInput(session, "file_filter", selected = "Alle")
  })
  
  ## MSB Name extrahieren ----
  observe({
    req(csv_files())
    
    files <- csv_files()
    
    if (input$msb_type != "alle") {
      files <- files[str_detect(basename(files), fixed(input$msb_type))]
    }
    
    msb_names <- extract_msb_name(basename(files))
    
    isolate({
      # Nur neu setzen, wenn der aktuelle Filterwert nicht mehr gültig ist
      if (!(input$file_filter %in% msb_names)) {
        updateSelectInput(
          session,
          "file_filter",
          choices = c("Alle", sort(unique(msb_names))),
          selected = "Alle"
        )
      } else {
        updateSelectInput(
          session,
          "file_filter",
          choices = c("Alle", sort(unique(msb_names))),
          selected = input$file_filter
        )
      }
    })
  })
  

  ## MSB Name dynamisch Liste ----
  filtered_msb_names <- reactive({
    req(csv_files())
    
    file_names <- basename(csv_files())   # ALLE Dateien nehmen!
    extract_msb_name(file_names)           # Alle MSB Namen extrahieren
  })
  

  # --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---#
  # ----                             PLOT                                   ----
  # --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---#
  
  # --- Eindeutige MSB-Namen einmalig ermitteln ----
  max_unique_msb <- reactive({
    all_csv_files <- get_csv_files()
    all_msb_names <- unlist(lapply(all_csv_files, function(file) {
      extract_msb_name(basename(file))
    }))
    length(unique(all_msb_names))
  })
  
  # --- FUNKTION update_plot_height ---
  update_plot_height <- function() {
    req(filtered_files())
    files <- filtered_files()
    msb_names <- extract_msb_name(basename(files))
    unique_msb_count <- length(unique(msb_names))
    balken_hoehe <- 35
    achsen_platz <- 50
    max_plot_hoehe <- 583 # Maximale Plot-Höhe vor dem Scrollen
    
    benoetigte_hoehe <- (unique_msb_count * balken_hoehe) + achsen_platz
    tatsaechliche_hoehe <- min(benoetigte_hoehe, max_plot_hoehe)
    
    # Speichere die tatsächliche Höhe für das UI
    plot_height(tatsaechliche_hoehe)
  }
  
  observeEvent(filtered_files(), {
    update_plot_height()
  })

  output$plot_ui <- renderUI({
    div(
      class = "plot-container",
      plotOutput("plot_msb", height = paste0(plot_height(), "px"))
    )
  })
  
  output$plot_msb <- renderPlot({
    req(filtered_files())
    # Einfach die Funktion aufrufen
    msb_plot(filtered_files())
  })
  
  # --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---#
  # ----                          HAUPTABELLE                               ----
  # --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---#
  
  ## Daten aktualisieren ----
  ### Aktuell ausgewählte Datei ----
  # (Start: erste Datei oder NULL)
  current_file <- reactiveVal(NULL)
  
  # Beim Start automatisch die erste Datei
  observe({
    req(filtered_files())
    current_file(filtered_files()[1])
  })
  
  ### MSB-Name aus aktueller Datei extrahieren ----
  current_msb <- reactive({
    req(current_file())
    basename(current_file()) %>% extract_msb_name()
  })
  
  ### MSB-Kontakt aus Email ----
  current_contact <- reactive({
    req(kontaktdaten())
    req(current_msb())
    
    kontaktdaten() %>%
      filter(MSB == current_msb()) %>%
      slice(1)  # Falls mehrere Treffer, nimm den ersten
  })
  
  ### Eingabefelder aktualisieren ----
  observe({
    contact <- current_contact()
    
    if (nrow(contact) > 0) {
      # ---- Monatszeitraum einmal bestimmen ----
      monate_plain <- extract_monate_final(current_csv_data())
      monate_final_plain(monate_plain)  # In reactiveVal speichern
      
      # Empfänger setzen
      updateTextInput(session, "email_empfaenger", value = contact$EMAIL)
      
      # --- Betreff dynamisch aufbauen ---
      filename <- basename(current_file())
      vnb <- extract_vnb(filename)
      msb_name <- current_msb()
      
      betreff_text <- paste0(
        "Fehlende Lastgangdaten im ", monate_final_plain(), ".    MSB ", msb_name, " für VNB ", vnb
      )
      
      updateTextInput(session, "email_betreff", value = betreff_text)
      
      # --- Email Body generieren ---
      body_html <- generate_email_body(current_csv_data(), contact, monate_final_plain())  # <<<<< HIER angepasst!
      
      output$email_body_html <- renderUI({
        HTML(body_html)
      })
      email_body_content(body_html)
      
    } else {
      # Kein Kontakt
      missing_msb_html <- paste0(
        "<b>Hinweis:</b> Kein Eintrag in den E-Mail Kontakten für MSB <b>", current_msb(), "</b> vorhanden."
      )
      
      updateTextInput(session, "email_empfaenger", value = "")
      updateTextInput(session, "email_betreff", value = "")
      
      output$email_body_html <- renderUI({
        HTML(missing_msb_html)
      })
    }
  })
  
  
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
