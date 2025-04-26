# --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --
#
# NFM DASHBOARD                                                             ----
#
# Author : Sascha Kornberger
# Datum  : 26.04.2025
# Version: 0.1.0
#
# History:
# 0.1.0  Funktion: 
#
# --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --
rm(list = ls())

## BENOETIGTE PAKETE ------------------------------------------------------------
# Optionen setzen – unterdrückt manche Dialoge zusätzlich
options(install.packages.check.source = "no")

# Liste der Pakete
pakete <- c(
  "devtools", "shiny", "bs4Dash", "DT", "tidyverse", "janitor",
  "lubridate", "base64enc", "shinyWidgets", "readr", "stringr"
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

# Funktion zum Laden der RData-Datei mit den Einstellungen
load_email_data <- function() {
  load("EMails.RData") 
  return(kontaktdaten) 
}


# Funktion zum Abrufen der CSV-Dateien aus dem festen Ordner
get_csv_files <- function(directory = "data") {
  # Hole alle CSV-Dateien im Ordner
  files <- list.files(directory, pattern = "\\.csv$", full.names = TRUE)
  # TODO Debugging Ausgabe entfernen
  print(files)
  return(files)
}


# Funktion zum Extrahieren des Namens zwischen "MSB" und "an" aus den Dateinamen
extract_msb_name <- function(file_names) {
  # Verwende regex, um den Text zwischen "MSB" und "an" zu extrahieren
  msb_names <- str_extract(file_names, "(?<=MSB\\s)(.*?)(?=\\san)")
  # Entferne NA-Werte
  return(msb_names[!is.na(msb_names)])  
}


# Funktion zum Extrahieren des MSB-Typs (wMSB oder gMSB) aus den Dateinamen
extract_msb_type <- function(file_names) {
  # Verwende regex, um nach 'wMSB' oder 'gMSB' zu suchen
  msb_types <- str_extract(file_names, "(wMSB|gMSB)")
  # Entferne NA-Werte
  return(msb_types[!is.na(msb_types)]) 
}

# --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- -#
# ----                         USER INTERFACE                               ----
# --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- -#
# UI-Definition mit bs4Dash
ui <- bs4DashPage(
  title = "LG-NACHFORDERUNG",
  header = bs4DashNavbar(),
  
  sidebar = bs4DashSidebar(disable = TRUE),  # Sidebar deaktiviert
  controlbar = bs4DashControlbar(),
  
  body = bs4DashBody(
    includeCSS(file.path(getwd(), "www/styles.css")),
    includeScript(file.path(getwd(), "www/custom.js")),
    
    tabsetPanel(
      tabPanel(
        "Hauptbereich",
        fluidRow(
          
          # Linke Seite: Filter + Plot in einer einzigen Box
          column(
            width = 3,
            class = "messstellenbetreiber-box-parent",
            box(
              title = "MESSSTELLENBETREIBER",
              width = 12,
              style = "font-family: monospace;",

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
                    selected = "",
                    multiple = FALSE,
                    width = "100%"
                  )
                )
              ),
              
              # Abstand
              br(),
              
              # ggplot
              div(
                style = "border: 1px solid #ccc; padding: 10px;",
                plotOutput("plot_msb", height = "450px")
              )
            )
          ),
          
          # Rechte Seite: Tabelle mit CSV-Dateien
          column(
            width = 9,
            box(
              title = "CSV-Dateien",
              DTOutput("file_list"),
              width = 12
            )
          )
          
        )
      ),
      
      tabPanel(
        "E-Mails",
        fluidRow(
          box(
            title = "E-Mail-Details",
            textOutput("email_preview"),
            width = 12
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
  
  
  # MSB extrahieren ----
  observe({
    # Hole die CSV-Dateinamen aus dem Ordner (ohne den Pfad)
    file_names <- basename(csv_files())
    
    # Extrahiere die MSB-Kundennamen aus den Dateinamen
    msb_names <- extract_msb_name(file_names)
    
    # Extrahiere die MSB-Typen (wMSB oder gMSB) aus den Dateinamen
    msb_types <- extract_msb_type(file_names)
    
    # Erstelle die Dropdown-Optionen für den MSB-Kunden
    updateSelectInput(session, "file_filter", choices = c("Alle" = "", unique(msb_names)))
  })
  
  # Wenn MSB-Type gewechselt wird, Dropdown zurücksetzen
  observeEvent(input$msb_type, {
    updateSelectInput(
      session,
      "file_filter",
      selected = ""  # Leere Auswahl entspricht "Alle"
    )
  })
  

  # Reaktive gefilterte Dateien je nach MSB-Type
  filtered_files <- reactive({
    req(csv_files())  # Stelle sicher, dass CSV-Dateien vorhanden sind
    
    # Alle aktuellen Dateien
    files <- csv_files()
    
    # Filter anwenden je nach RadioButton-Auswahl
    if (input$msb_type == "alle") {
      files
    } else {
      files[str_detect(basename(files), fixed(input$msb_type))]
    }
  })
  
  
  
  filtered_msb_names <- reactive({
    file_names <- basename(filtered_files())
    extract_msb_name(file_names)
  })
  
  observe({
    updateSelectInput(
      session,
      "file_filter",
      choices = c("Alle" = "", unique(filtered_msb_names()))
    )
  })
  
  
  
  output$file_list <- renderDT({
    req(filtered_files())
    
    selected_msb <- input$file_filter
    
    # Wenn ein bestimmter MSB ausgewählt ist, weiter filtern
    if (selected_msb != "") {
      filtered <- filtered_files()[str_detect(basename(filtered_files()), fixed(selected_msb))]
    } else {
      filtered <- filtered_files()
    }
    
    datatable(
      data.frame(Dateien = filtered),
      options = list(pageLength = 15)
    )
  })
  
  
  output$plot_msb <- renderPlot({
    req(filtered_files())
    
    # Hole Dateinamen
    file_names <- basename(filtered_files())
    
    # Extrahiere MSB-Namen
    msb_names <- extract_msb_name(file_names)
    
    # Erzeuge Tabelle: Anzahl Dateien pro MSB
    msb_count <- tibble(msb = msb_names) %>%
      count(msb, sort = TRUE)
    
    # Plot erstellen
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
        nudge_y = 0.2,
        nudge_x = 0.1,
        color = "black",
        fontface = "bold",
        size = 4
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
      scale_y_discrete(labels = NULL)

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
