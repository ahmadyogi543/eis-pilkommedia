# Load necessary libraries
library(dplyr)
library(forcats)
library(ggplot2)
library(openxlsx)
library(tidyr)
library(readxl)
library(shiny)

# Helpers
read_file_upload_xls <- function(filepath) {
  scores <- read_xls(filepath) |>
    rename(`JENIS MATA KULIAH` = `JENIS MATA KULIAH (1= WAJIB , 0= PILIHAN)`) |>
    rename_with(~gsub("\\ ", ".", .), everything()) |>
    select(-NILAI.ANGKA) |>
    mutate(
      NILAI.INDEKS = as.numeric(NILAI.INDEKS),
      SKS.MATA.KULIAH = as.numeric(SKS.MATA.KULIAH),
      JENIS.MATA.KULIAH = as.numeric(JENIS.MATA.KULIAH),
      NILAI.HURUF = if_else(
        is.na(NILAI.HURUF),
	    if_else(
	      is.na(NILAI.INDEKS),
	      NA,
	      case_when(
            NILAI.INDEKS >= 4.00 ~ "A",
            NILAI.INDEKS >= 3.75 ~ "A-",
            NILAI.INDEKS >= 3.50 ~ "B+",
            NILAI.INDEKS >= 3.00 ~ "B",
            NILAI.INDEKS >= 2.75 ~ "B-",
            NILAI.INDEKS >= 2.50 ~ "C+",
            NILAI.INDEKS >= 2.00 ~ "C",
            NILAI.INDEKS >= 1.75 ~ "D+",
            NILAI.INDEKS >= 1.50 ~ "D",
            NILAI.INDEKS >= 0.00 ~ "E",
	        TRUE ~ "E"
	      )
	    ),
	   NILAI.HURUF
    )
  )
  
  # change some column to use factors
  scores$NILAI.HURUF <- factor(
    scores$NILAI.HURUF,
    ordered = TRUE,
    levels = c(
      "E", "D",
      "D+", "C", "C+",
      "B-", "B", "B+",
      "A-", "A"
    )
  )
  
  return(scores)
}

get_courses_table <- function(curriculum, courses) {
  new_courses <- courses |>
    na.omit() |>
    filter(NAMA.KURIKULUM == curriculum) |>
    select(
      KODE.MATA.KULIAH,
      NAMA.MATA.KULIAH,
      NILAI.HURUF,
      NILAI.INDEKS
    ) |>
    group_by(KODE.MATA.KULIAH, NAMA.MATA.KULIAH) |>
    mutate(STATUS = if_else(
      NILAI.HURUF %in% c("A", "A-", "B+", "B", "B-", "C+", "C"),
      "Lulus",
      "Tidak Lulus"
      )
    ) |>
    summarise(
      `FREKUENSI LULUS` = sum(STATUS == "Lulus"),
      `PERSENTASE LULUS (%)` = round(sum(STATUS == "Lulus") / n() * 100, 2),
      `FREKUENSI TIDAK LULUS` = sum(STATUS == "Tidak Lulus"),
      `PERSENTASE TIDAK LULUS (%)` = round(sum(STATUS == "Tidak Lulus") / n() * 100, 2)
    ) |>
    ungroup() |>
    mutate(NO = row_number()) |>
    select(NO, everything()) |>
    rename_with(~gsub("\\.", " ", .), everything())

  return(new_courses)
}

get_courses_period <- function(scores) {
  get_course_period <- function(period) {
    course_semester <- ifelse(substr(
      period,
      start = 5,
      stop = 5
    ) == "1", "Ganjil", "Genap")
    course_year <- as.numeric(substr(
      period,
      start = 1,
      stop = 4
    ))

    return(sprintf("%d/%d %s", course_year, course_year+1, course_semester))
  }

  course_info <- scores |>
    select(PERIODE) |>
    mutate(
      PERIODE = get_course_period(PERIODE)
    ) |>
    slice(1)

  return(course_info$PERIODE)
}

write_xlsx_title <- function(
  wb,
  sheet_name,
  col_size,
  curriculum,
  period
) {
  titles <- c(
    "Rekap Kelulusan Mata Kuliah",
    "Progam Studi Pendidikan Komputer",
    "FKIP Universitas Lambung Mangkurat",
    "",
    paste0("Nama Kurikulum: ", curriculum),
    paste0("Periode: ", period),
    ""
  )

  for (i in seq_along(titles)) {
    writeData(wb, sheet = sheet_name, x = titles[i], startRow = i, startCol = 1)
    mergeCells(wb, sheet = sheet_name, cols = 1:col_size, rows = i)
  }
}

export_to_xlsx <- function(courses_table, curriculum, period) {
  # drop the no column
  courses_table <- courses_table |> select(-NO)
  
  wb <- createWorkbook()

  addWorksheet(wb, "data")

  #title
  write_xlsx_title(
    wb,
    "data",
    ncol(courses_table),
    curriculum,
    period
  )

  # data
  writeData(wb, sheet = "data", x = courses_table, startRow = 9, startCol = 1)

  # plot
  courses_table_plot <- courses_table |>
    ggplot(aes(x = `KODE MATA KULIAH`)) +
	  geom_line(aes(y = `PERSENTASE LULUS (%)`, color = "LULUS (%)"), group = 1) +
      geom_text(aes(y = `PERSENTASE LULUS (%)`, label = `PERSENTASE LULUS (%)`), vjust = -0.5, angle = 45) +
	  geom_line(aes(y = `PERSENTASE TIDAK LULUS (%)`, color = "TIDAK LULUS (%)"), group = 2) +
      geom_text(aes(y = `PERSENTASE TIDAK LULUS (%)`, label = `PERSENTASE TIDAK LULUS (%)`), vjust = -0.5, angle = 45) +
	  labs(x  = "", y = "", color = "STATUS") +
	  scale_color_manual(
        values = c(
          "LULUS (%)" = "cornflowerblue",
	      "TIDAK LULUS (%)" = "tomato"
	    )
	  ) +
	  theme_bw() +
	  theme(
	    axis.text = element_text(size = 12),
	    legend.text = element_text(size = 12),
	    legend.title = element_text(size = 12, face = "bold"),
	    axis.text.x = element_text(angle = 90, vjust = 0.5, hjust = 1)
	  )
  ggsave("plot.png", plot = courses_table_plot, width = 12, height = 5, units = "in")


  insertImage(
    wb,
    sheet = "data",
    "plot.png",
    startRow = 11 + nrow(courses_table),
    startCol = 1,
    width = 12,
    height = 5
  )

  # set column widths
  setColWidths(wb, sheet = "data", cols = 1:6, widths = "auto")

  tmp <- tempfile(fileext = ".xlsx")
  saveWorkbook(wb, file = tmp)
  file.remove("plot.png")

  return(tmp)
}

# Define UI for application
ui <- fluidPage(
  h4("Rekap Kelulusan Mata Kuliah"),
  h4("Program Studi Pendidikan Komputer"),
  h4("FKIP Universitas Lambung Mangkurat"),
  br(),
  fluidRow(
    column(
      width = 3,
      fluidRow(
        column(
          width = 12,
          fileInput("file", "Pilih File XLS",
            accept = c(
              "application/vnd.ms-excel",
              ".xls"
            )
          ),
          uiOutput("curriculum_ui")
        ),
        column(
          width = 12,
          htmlOutput("courses_period_ui")
        ),
        column(
          width = 12,
          br(),
          uiOutput("xlsx_export_ui")
        )
      )
    ),
    column(
      width = 9,
      uiOutput("courses_table_ui")
    )
  ),
  fluidRow(
    column(
      width = 12,
      uiOutput("courses_plot_ui")
    )
  ),
  br()
)

server <- function(input, output, session) {
  # states
  scores <- reactive({
    req(input$file)
    return(read_file_upload_xls(input$file$datapath))
  })
  curriculum <- reactive(input$curriculum)

  # render select curriculum ui
  output$curriculum_ui <- renderUI({
    req(scores())
    
    # get choices
    choices <- scores() |>
      select(NAMA.KURIKULUM) |>
      add_row(NAMA.KURIKULUM = "-") |>
      arrange(NAMA.KURIKULUM) |>
      distinct(NAMA.KURIKULUM)
     
    selectInput(
      "curriculum",
      "Pilih Kurikulum",
      choices = choices
    )
  })

  # render courses table
  output$courses_table_ui <- renderUI({
    req(scores())
    req(curriculum())
   
    if (curriculum() != "-") {
      renderDataTable(
        get_courses_table(curriculum(), scores()),
        options = list(
          pageLength = 10,
          columnDefs = list(
            list(targets = "_all", searchable = FALSE)
          )
        )
      ) 
    }
  })

  # render courses period ui
  output$courses_period_ui <- renderUI({
    req(scores())
    req(curriculum())

    courses_period <- get_courses_period(scores())

    if (curriculum() != "-") {
      HTML(
        paste(
          paste0("<br />"),
          paste0("<p>Periode: ", courses_period, "</p>")
        )
      )
    }
  })

  # render xlsx export ui
  output$xlsx_export_ui <- renderUI({
    req(scores())
    req(curriculum())

    if (curriculum() != "-") {
      downloadButton("xlsx", "Export to XLSX")
    }
  })

  output$xlsx <- downloadHandler(
    filename = function() {
      paste0(curriculum(), " - ", Sys.Date(), ".xlsx")
    },
    content = function(file) {
      courses_table <- get_courses_table(curriculum(), scores())
      course_period <- get_courses_period(scores())
      tmp <- export_to_xlsx(courses_table, curriculum(), course_period)
      file.copy(tmp, file)
    }
  )

  # render courses plot
  output$courses_plot_ui <- renderUI({
    req(scores())
    req(curriculum())

    if (curriculum() != "-") {
      courses_table <- get_courses_table(curriculum(), scores())

      renderPlot({
        courses_table |>
	      ggplot(aes(x = `KODE MATA KULIAH`)) +
	        geom_line(
	        aes(y = `PERSENTASE LULUS (%)`, color = "LULUS (%)"), group = 1) +
            geom_text(aes(y = `PERSENTASE LULUS (%)`, label = `PERSENTASE LULUS (%)`), vjust = -0.5, angle = 45) +
	        geom_line(aes(y = `PERSENTASE TIDAK LULUS (%)`, color = "TIDAK LULUS (%)"), group = 2) +
            geom_text(aes(y = `PERSENTASE TIDAK LULUS (%)`, label = `PERSENTASE TIDAK LULUS (%)`), vjust = -0.5, angle = 45) +
	        labs(x  = "", y = "", color = "STATUS") +
	        scale_color_manual(
              values = c(
                "LULUS (%)" = "cornflowerblue",
	            "TIDAK LULUS (%)" = "tomato"
	          )
	        ) +
	        theme_bw() +
	        theme(
	          axis.text = element_text(size = 14),
	          legend.text = element_text(size = 14),
	          legend.title = element_text(size = 14, face = "bold"),
	          axis.text.x = element_text(angle = 90, vjust = 0.5, hjust = 1)
	        )
      })
    }
  })
}

# Run the application
shinyApp(ui = ui, server = server)

