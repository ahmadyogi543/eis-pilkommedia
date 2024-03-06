# Load necessary libraries
library(dplyr)
library(forcats)
library(ggplot2)
library(openxlsx)
library(readxl)
library(tidyr)
library(shiny)

# Helpers
read_file_upload_xls <- function(filepath) {
  scores <- read_xls(filepath) |>
    rename(`JENIS MATA KULIAH` = `JENIS MATA KULIAH (1= WAJIB , 0= PILIHAN)`) |>
    rename_with(~gsub("\\ ", ".", .), everything()) |>
    mutate(
      NILAI.ANGKA = as.numeric(NILAI.ANGKA),
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

get_scores_table <- function(course_id, course_name, scores) {
  new_scores <- scores |>
    select(
      -JENIS.MATA.KULIAH,
      -NILAI.ANGKA,
      -NAMA.KELAS,
      -NAMA.KURIKULUM,
      -PERIODE,
      -SKS.MATA.KULIAH
    ) |>
    filter(
      KODE.MATA.KULIAH == course_id &
      NAMA.MATA.KULIAH == course_name
    ) |>
    select(
      -KODE.MATA.KULIAH,
      -NAMA.MATA.KULIAH
    ) |>
    arrange(NIM) |>
    mutate(NO = row_number()) |>
    select(NO, everything()) |>
    rename_with(~gsub("\\.", " ", .), everything())

  return(new_scores)
}

get_passed_percentage <- function(passed_scores) {
  passed_percentage <- passed_scores |>
    count(STATUS, name = "FREKUENSI") |>
    mutate(PERSENTASE = round(FREKUENSI / sum(FREKUENSI) * 100, 2))
  
  return(passed_percentage)
}

get_summarized_scores <- function(course_id, course_name, scores) {
 summarized_scores <- scores |>
    filter(
      KODE.MATA.KULIAH == course_id &
      NAMA.MATA.KULIAH == course_name
    ) |>
    group_by(NAMA.MATA.KULIAH, KODE.MATA.KULIAH) |>
    count(NILAI.HURUF, name = "FREKUENSI") |>
    complete(NILAI.HURUF, fill = list(n = 0)) |>
    replace_na(list(FREKUENSI = 0)) |>
    mutate(
      PERSENTASE = round(FREKUENSI / sum(FREKUENSI) * 100, 2)
    ) |>
    arrange(
      NAMA.MATA.KULIAH,
      KODE.MATA.KULIAH,
      desc(NILAI.HURUF),
      desc(FREKUENSI)
    ) |>
    ungroup(NAMA.MATA.KULIAH, KODE.MATA.KULIAH) |>
    select(
      -NAMA.MATA.KULIAH,
      -KODE.MATA.KULIAH
    ) |>
    rename_with(~gsub("\\.", " ", .), everything())
 
 return(summarized_scores)
}

get_passed_scores <- function(course_id, course_name, scores) {
  is_pass <- function(grade) {
    pass <- ifelse(
      grade %in% c("A", "A-", "B+", "B", "B-", "C+", "C"),
      "Lulus",
      "Tidak Lulus"
    )
    
    return(pass)
  }
  
  passed_scores <- scores |>
    select(-NILAI.ANGKA) |>
    na.omit() |>
    filter(
      KODE.MATA.KULIAH == course_id &
      NAMA.MATA.KULIAH == course_name
    ) |>
    mutate(STATUS = is_pass(NILAI.HURUF)) |>
    select(STATUS)
 
  return(passed_scores)
}

get_course_info <- function(course_id, course_name, scores) {
  get_course_type <- function(type) {
    type <- ifelse(type == 0, "Pilihan", "Wajib")
    return(type)
  }
  
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
    filter(
      KODE.MATA.KULIAH == course_id &
      NAMA.MATA.KULIAH == course_name
    ) |>
    select(SKS.MATA.KULIAH, JENIS.MATA.KULIAH, PERIODE) |>
    mutate(
      JENIS.MATA.KULIAH = get_course_type(JENIS.MATA.KULIAH),
      PERIODE = get_course_period(PERIODE)
    ) |>
    slice(1)
  
  return(course_info)
}

write_xlsx_title <- function(
  wb,
  sheet_name,
  col_size,
  curriculum,
  course_id,
  course_name,
  course_info
) {
  titles <- c(
    "Statistik Deskriptif Nilai Mata Kuliah",
    "Progam Studi Pendidikan Komputer",
    "FKIP Universitas Lambung Mangkurat",
    "",
    paste0("Nama Kurikulum: ", curriculum),
    paste0("Periode: ", course_info$PERIODE),
    paste0("Kode Mata Kuliah: ", course_id),
    paste0("Nama Mata Kuliah: ", course_name),
    paste0("SKS Mata Kuliah: ", course_info$SKS.MATA.KULIAH),
    paste0("Jenis Mata Kuliah: ", course_info$JENIS.MATA.KULIAH),
    ""
  )

  for (i in seq_along(titles)) {
    writeData(wb, sheet = sheet_name, x = titles[i], startRow = i, startCol = 1)
    mergeCells(wb, sheet = sheet_name, cols = 1:col_size, rows = i)
  }
}

export_to_xlsx <- function(
  scores_table,
  scores_summary,
  curriculum,
  course_id,
  course_name,
  course_info
) {
  wb <- createWorkbook()

  addWorksheet(wb, "data")
  #title
  write_xlsx_title(
    wb,
    "data",
    3,
    curriculum,
    course_id,
    course_name,
    course_info
  )
  # data
  writeData(wb, sheet = "data", x = scores_table |> select(-NO), startRow = 12, startCol = 1)
  # summary
  writeData(wb, sheet = "data", x = scores_summary, startRow = 14 + nrow(scores_table), startCol = 1)

  # plot
  summary_plot <- scores_summary |>
    ggplot(aes(x = forcats::fct_rev(`NILAI HURUF`), y = `FREKUENSI`)) +
      geom_line(group = 1) +
      geom_point() +
      geom_text(aes(label = FREKUENSI), vjust = -0.5) +
      labs(x = "", y = "") +
      theme_bw() +
      theme(axis.text = element_text(size = 14))
  ggsave("plot.png", plot = summary_plot)

  insertImage(
    wb,
    sheet = "data",
    "plot.png",
    startRow = 14 + nrow(scores_table) + nrow(scores_summary) + 2,
    startCol = 1,
    width = 5,
    height = 4
  )

  # set column widths
  setColWidths(wb, sheet = "data", cols = 1:3, widths = c(18, 18, 18))

  tmp <- tempfile(fileext = ".xlsx")
  saveWorkbook(wb, file = tmp)
  file.remove("plot.png")

  return(tmp)
}


# Define UI for application
ui <- fluidPage(
  h4("Statistik Deskriptif Nilai Mata Kuliah"),
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
          uiOutput("curriculum_ui"),
          uiOutput("course_ui"),
        ),
        column(
          width = 12,
          htmlOutput("course_info_ui")
        ),
        column(
          width = 12,
          htmlOutput("passed_scores_ui")
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
      uiOutput("scores_table_ui")
    )
  ),
  br(),
  fluidRow(
    column(
      width = 4,
      offset = 3,
      uiOutput("scores_summary_table_ui")
    ),
    column(
      width = 5,
      uiOutput("scores_summary_plot_ui")
    )
  )
)

server <- function(input, output, session) {
  # states
  scores <- reactive({
    req(input$file)
    return(read_file_upload_xls(input$file$datapath))
  })
  course <- reactive(input$course)
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

  # render select course ui
  output$course_ui <- renderUI({
    req(curriculum())
    
    if (curriculum() != "-") {
      # get choices
      choices <- scores() |>
        filter(NAMA.KURIKULUM == curriculum()) |>
        mutate(NAMA.MATA.KULIAH = paste0(
          KODE.MATA.KULIAH,
          " - ",
          NAMA.MATA.KULIAH
        )) |>
        select(NAMA.MATA.KULIAH) |>
        add_row(NAMA.MATA.KULIAH = "-") |>
        arrange(NAMA.MATA.KULIAH) |>
        distinct(NAMA.MATA.KULIAH)
       
      selectInput(
        "course",
        "Pilih Mata Kuliah",
        choices = choices
      )
    }
  })
  
  # render course info
  output$course_info_ui <- renderUI({
    req(scores())
    req(course())
    req(curriculum())
   
    # split the id and name of the selected course 
    split_text <- strsplit(course(), " - ")
    course_id <- trimws(split_text[[1]][1])
    course_name <- trimws(split_text[[1]][2])
    
    course_info <- get_course_info(
      course_id,
      course_name,
      scores()
    )
    
    if (curriculum() != "-" & course() != "-") {
      HTML(
        paste(
          paste0("<br />"),
          paste0("<p><b>Informasi Mata Kuliah</b></p>"),
          paste0("<p>SKS Mata Kuliah: ", course_info$SKS.MATA.KULIAH, "</p>"),
          paste0("<p>Jenis Mata Kuliah: ", course_info$JENIS.MATA.KULIAH, "</p>"),
          paste0("<p>Periode: ", course_info$PERIODE, "</p>")
        )
      )
    }
  })
 
  # render scores table
  output$scores_table_ui <- renderUI({
    req(scores())
    req(course())
    req(curriculum())

    shinycssloaders::showSpinner()
   
    # split the id and name of the selected course 
    split_text <- strsplit(course(), " - ")
    course_id <- trimws(split_text[[1]][1])
    course_name <- trimws(split_text[[1]][2])
   
    if (curriculum() != "-" & course() != "-") {
      shinycssloaders::hideSpinner()

      renderDataTable(
        get_scores_table(
          course_id,
          course_name,
          scores()
        ),
        options = list(
          pageLength = 10,
          columnDefs = list(
            list(targets = "_all", searchable = FALSE)
          )
        )
      ) 
    }
  })
  
  # render passed scores ui
  output$passed_scores_ui <- renderUI({
    req(scores())
    req(course())
    req(curriculum())
   
    # split the id and name of the selected course 
    split_text <- strsplit(course(), " - ")
    course_id <- trimws(split_text[[1]][1])
    course_name <- trimws(split_text[[1]][2])
    
    passed_scores <- get_passed_scores(
      course_id,
      course_name,
      scores()
    )
    
    passed_percentage <- get_passed_percentage(passed_scores)
    
    get_percentage <- function(x) {
      percentage <- passed_percentage |>
        filter(STATUS == unique(x)) |>
        pull(PERSENTASE)
      
      paste(
        unique(x),
        sprintf(
          " (%.1f%%)",
          percentage
        )
      )
    }
    
    if (curriculum() != "-" & course() != "-") {
      HTML(
        paste(
          paste0("<br />"),
          paste0("<p><b>Status Kelulusan</b></p>"),
          paste0("<p>", get_percentage("Lulus"), "</p>"),
          paste0("<p>", get_percentage("Tidak Lulus"), "</p>")
        )
      )
    }
  })

  # render export to xlsx button
  output$xlsx_export_ui <- renderUI({
    req(scores())
    req(course())
    req(curriculum())
    

    if (curriculum() != "-" & course() != "-") {
        downloadButton("xlsx", "Export to XLSX")
    }
  })

  output$xlsx <- downloadHandler(
    filename = function() {
      # Generate a dynamic filename, for example, based on the current date
      paste0(course(), " - ", Sys.Date(), ".xlsx")
    },
    content = function(file) {
        split_text <- strsplit(course(), " - ")
        course_id <- trimws(split_text[[1]][1])
        course_name <- trimws(split_text[[1]][2])
        
        scores_table <- get_scores_table(
          course_id,
          course_name,
          scores()
        )

        summarized_scores <- get_summarized_scores(
          course_id,
          course_name,
          scores()
        )

        course_info <- get_course_info(
          course_id,
          course_name,
          scores()
        )

        tmp <- export_to_xlsx(
          scores_table,
          summarized_scores,
          curriculum(),
          course_id,
          course_name,
          course_info
        )
        file.copy(tmp, file)
    }
  )
  
  # render scores summary table
  output$scores_summary_table_ui <- renderUI({
    req(scores())
    req(course())
    req(curriculum())
    
    if (curriculum() != "-" & course() != "-") {
      split_text <- strsplit(course(), " - ")
      course_id <- trimws(split_text[[1]][1])
      course_name <- trimws(split_text[[1]][2])
      
      summarized_scores <- get_summarized_scores(
        course_id,
        course_name,
        scores()
      )
      
      renderTable(summarized_scores)
    }
  })
  
  # render scores summary plot
  output$scores_summary_plot_ui <- renderUI({
    req(scores())
    req(course())
    req(curriculum())
    
    if (curriculum() != "-" & course() != "-") {
      split_text <- strsplit(course(), " - ")
      course_id <- trimws(split_text[[1]][1])
      course_name <- trimws(split_text[[1]][2])
      
      summarized_scores <- get_summarized_scores(
        course_id,
        course_name,
        scores()
      )
       
      renderPlot({
        ggplot(
          summarized_scores,
          aes(x = forcats::fct_rev(`NILAI HURUF`), y = `FREKUENSI`)
        ) +
          geom_line(group = 1) +
          geom_point() +
          geom_text(aes(label = FREKUENSI), vjust = -0.5) +
          labs(x = "", y = "") +
          theme_bw() +
          theme(
            axis.text = element_text(size = 14)
          )
      })
    }
  })
}

# Run the application
shinyApp(ui = ui, server = server)

