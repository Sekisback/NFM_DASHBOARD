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
  print(files)
  return(files)
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


# Erstelle Plot als Balkendiagramm mit der Menge der Dateien pro MSB
msb_plot <- function(file_names) {
  # MSB-Namen extrahieren
  msb_names <- extract_msb_name(basename(file_names))
  
  # Zähle Vorkommen
  msb_count <- tibble(msb = msb_names) %>%
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
  default_einleitung <- "für die in der Tabelle aufgeführten Marktlokationen wurden, trotz zuvor versendeter ORDERS, noch keine vollständigen Lastgangdaten empfangen.<br>
Wir möchten Sie daher bitten, diese schnellstmöglich in der aktuellen MSCONS Version an die Ihnen bekannte edifact Adresse zu senden."
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
  
  ## Navbar ----
  header = bs4DashNavbar(),
  
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
        fluidRow(
          
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
                    selected = "",
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
          
          #### Anzeige ----
          column(
            width = 9,
            box(
              title = "CSV-Dateien",
              DTOutput("file_list"),
              width = 12,
              collapsible = FALSE
            )
          )
          
        )
      ),
      
      ### Tab E-Mails ----
      tabPanel(
        "E-Mails",
        fluidRow(
          box(
            title = "E-Mail-Details",
            width = 12,
            collapsible = FALSE,
            
            # Neue Zeile: Buttons
            fluidRow(
              column(
                width = 12,
                style = "margin-bottom: 10px;", # Kleiner Abstand unter Buttons
                actionButton("add_entry", "Neuer Eintrag", class = "btn-primary"),
                actionButton("edit_entry", "Eintrag bearbeiten", class = "btn-primary"),
                actionButton("delete_entry", "Eintrag löschen", class = "btn-primary")
              )
            ),
            
            # Danach die Tabelle
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
  
  
  # --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---#
  # ----                            FILTER                                  ----
  # --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---#
  
  ## MSB Typ ---- 
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
  
  ## MSB Typ Überwachung ----
  observeEvent(input$msb_type, {
    # Bei Änderung des MSB Typ Filter MSB Name auf Alle setzen
    updateSelectInput(
      session,
      "file_filter",
      # Leere Auswahl entspricht "Alle"
      selected = ""  
    )
  })
  
  
  ## MSB Name extrahieren ----
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
  

  ## MSB Name dynamisch Liste ----
  filtered_msb_names <- reactive({
    # Extrahiert die Dateinamen aus dem reaktiven Ausdruck filtered_files()
    file_names <- basename(filtered_files())
    # Wendet die Funktion extract_msb_name auf die Dateinamen an
    extract_msb_name(file_names)
  })
  
  ## MSB Name Filter Überwachung ----
  observe({
    # Aktualisiert das selectInput-Feld mit der ID "file_filter"
    updateSelectInput(
      session,
      "file_filter",
      # Definiert die Auswahlmöglichkeiten für das Dropdown-Menü
      choices = c("Alle" = "", unique(filtered_msb_names()))
    )
  })

  # --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---#
  # ----                             PLOT                                   ----
  # --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---#
  
  ## Plot dynamische Höhe berechnen ----
  plot_height <- reactive({
    req(filtered_files())
    file_names <- basename(filtered_files())
    msb_names <- extract_msb_name(file_names)
    # Mindestens 75px, sonst 12px pro MSB
    max(75, length(msb_names) * 12)   
  })
  
  ## Plot Dynamisches UI ----
  output$plot_ui <- renderUI({
    div(
      style = "border: 1px solid #ccc; padding: 10px; background-color: white;",
      plotOutput("plot_msb", height = paste0(plot_height(), "px"))
    )
  })
  
  output$plot_msb <- renderPlot({
    req(filtered_files())
    # Einfach die Funktion aufrufen
    msb_plot(filtered_files())
  })
  

  
  ## TABELLE *******************************************************************
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
