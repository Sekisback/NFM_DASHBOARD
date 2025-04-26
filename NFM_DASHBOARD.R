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

# E-MAIL KONTAKTE SPEICHERN 
save_email_data <- function(kontaktdaten) {
  # Speichert die übergebenen Daten in der Datei "Emails.RData"
  save(kontaktdaten, file = "EMails.RData")
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


new_entry_modal <- function(title) {
  modalDialog(
    title = "Neuen Eintrag hinzufügen",
    textInput("new_msb", "MSB Name"),
    textInput("new_email", "E-Mail Adresse"),
    textInput("new_anrede", "Anrede", value = "Sehr geehrte Damen und Herren,"),
    textAreaInput("new_einleitung", "Einleitungssatz", value = "Bitte finden Sie die angeforderten Daten unten:"),
    footer = tagList(
      modalButton("Abbrechen"),
      actionButton("save_new_entry", "Speichern", class = "btn-primary")
    ),
    easyClose = TRUE
  )
}


edit_entry_modal <- function(selected_data) {
  modalDialog(
    title = "Eintrag bearbeiten",
    textInput("edit_msb", "MSB Name", value = selected_data$MSB),
    textInput("edit_email", "E-Mail Adresse", value = selected_data$EMAIL),
    textInput("edit_anrede", "Anrede", value = selected_data$ANREDE),
    textAreaInput("edit_einleitung", "Einleitungssatz", value = selected_data$EINLEITUNG),
    footer = tagList(
      modalButton("Abbrechen"),
      actionButton("save_edit_entry", "Speichern", class = "btn-primary")
    ),
    easyClose = TRUE
  )
}

show_success_toast <- function(message) {
  show_toast(
    title = message,
    text = NULL,
    type = "success",
    timer = 3000,
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
  
  
  
  # Plot dynamische Höhe berechnen
  plot_height <- reactive({
    req(filtered_files())
    file_names <- basename(filtered_files())
    msb_names <- extract_msb_name(file_names)
    max(75, length(msb_names) * 12)   # Mindestens 300px, sonst 50px pro MSB
  })
  
  # Dynamisches UI für Plot
  output$plot_ui <- renderUI({
    div(
      style = "border: 1px solid #ccc; padding: 10px; background-color: white;",
      plotOutput("plot_msb", height = paste0(plot_height(), "px"))
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
    
    # HIER: Maximalwert berechnen VOR ggplot!
    max_value <- max(msb_count$n, na.rm = TRUE)
    
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
      scale_y_discrete(labels = NULL)+
      scale_x_continuous(
        limits = c(0, ifelse(max_value < 10, 10, NA)),
        breaks = 0:max(10, max_value),
        expand = expansion(mult = c(0, 0.05))
      )

  })
  
  # --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---#
  # ----                            E-MAIL                                  ----
  # --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---#
  
  output$email_table <- renderDT({
    req(kontaktdaten())  # Sicherstellen, dass Daten geladen sind
    
    datatable(
      kontaktdaten(),
      selection = "single",
      options = list(
        paging = FALSE,         
        searching = FALSE,       
        ordering = TRUE,        
        autoWidth = TRUE,       
        dom = 't' 
      ),
      rownames = FALSE 
    )
  })
  
  
  observeEvent(input$add_entry, {
    showModal(new_entry_modal())
  })
  
  
  observeEvent(input$save_new_entry, {
    # 1. Aktuelle Daten holen
    data <- kontaktdaten()
    
    # 2. Neue Zeile bauen
    new_row <- tibble(
      MSB = input$new_msb,
      EMAIL = input$new_email,
      ANREDE = input$new_anrede,
      EINLEITUNG = input$new_einleitung
    )
    
    # 3. Neue Zeile unten anfügen
    updated_data <- bind_rows(data, new_row)
    
    # 4. Reactive Value aktualisieren
    kontaktdaten(updated_data)
    
    # 5. In die Datei speichern
    save_email_data(kontaktdaten())
    
    # 6. Modal schließen
    removeModal()
    show_success_toast("Neuer Eintrag erfolgreich gespeichert!")
  })
  
  
  
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
  
  observeEvent(input$edit_entry, {
    selected_row <- input$email_table_rows_selected
    
    if (length(selected_row) == 0) {
      showModal(modalDialog(
        title = "Hinweis",
        "Bitte wähle zuerst einen Eintrag aus, den du bearbeiten möchtest.",
        easyClose = TRUE
      ))
    } else {
      # Aktuellen Datensatz holen
      data <- kontaktdaten()
      selected_data <- data[selected_row, ]
      
      # Modal öffnen und vorausfüllen
      showModal(edit_entry_modal(selected_data))
    }
  })
  
  observeEvent(input$save_edit_entry, {
    selected_row <- input$email_table_rows_selected
    data <- kontaktdaten()
    
    # Überschreibe die ausgewählte Zeile mit neuen Werten
    data[selected_row, ] <- tibble(
      MSB = input$edit_msb,
      EMAIL = input$edit_email,
      ANREDE = input$edit_anrede,
      EINLEITUNG = input$edit_einleitung
    )
    
    # Aktualisiere reactiveVal
    kontaktdaten(data)
    
    # Speichern in RData
    save_email_data(kontaktdaten())
    
    # Modal schließen
    removeModal()
    show_success_toast("Eintrag erfolgreich aktualisiert!")
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
