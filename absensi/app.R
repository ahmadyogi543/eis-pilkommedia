# load necessary libraries
library(dplyr)
library(ggplot2)
library(readxl)
library(openxlsx)
library(shiny)

# helpers
rename_cols <- function(col_name) {
    if (grepl("^X[0-9]+$", col_name)) {
        number <- as.integer(gsub("X", "", col_name))
        new_number <- number - 3
        return(paste0("P", new_number))
    } else {
        return(col_name)
    }
}

format_period <- function(period) {
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
    return(sprintf("%d/%d %s", course_year, course_year + 1, course_semester))
}

get_attributes <- function(wb, sheet_name) {
    attributes <- read.xlsx(
        wb,
        sheet = sheet_name,
        rows = 5:9,
        cols = 3,
        colNames = FALSE
    )
    attributes <- list(
        period = format_period(attributes[1, 1]),
        course_name = attributes[2, 1],
        class_name = attributes[3, 1],
        course_id = attributes[4, 1],
        program = attributes[5, 1]
    )
    return(attributes)
}

get_dataframe <- function(wb, sheet_name) {
    dataframe <- read.xlsx(wb, sheet = sheet_name, startRow = 11) |>
        as_tibble() |>
        na.omit() |>
        select(-Persentase.Kehadiran) |>
        rename(
            NO = No,
            NIM = Nim,
            NAMA = Nama,
            P1 = Perkuliahan.Ke
        ) |>
        rename_with(
            function(x) sapply(x, rename_cols),
            everything()
        )

    dataframe <- dataframe %>%
        mutate(FREKUENSI.HADIR = rowSums(. == "H")) %>%
        mutate(FREKUENSI.TIDAK.HADIR = rowSums(. == "S") + rowSums(. == "I") + rowSums(. == "T")) %>%
        mutate(PERSENTASE.HADIR = round(FREKUENSI.HADIR / (FREKUENSI.HADIR + FREKUENSI.TIDAK.HADIR) * 100, digits = 2)) %>%
        mutate(PERSENTASE.TIDAK.HADIR = round(FREKUENSI.TIDAK.HADIR / (FREKUENSI.HADIR + FREKUENSI.TIDAK.HADIR) * 100, digits = 2)) %>%
        select(NO, NIM, NAMA, FREKUENSI.HADIR, FREKUENSI.TIDAK.HADIR, PERSENTASE.HADIR, PERSENTASE.TIDAK.HADIR, everything()) %>%
        rename_with(~gsub("\\.", " ", .), everything())

    return(dataframe)
}

get_attendance_data <- function(filepath) {
    wb <- loadWorkbook(filepath)
    sheet_names <- getSheetNames(filepath)
    attendance_data <- list()

    course_ids <- c()
    for (sheet_name in sheet_names) {
        if (sheet_name != "Worksheet") {
            # attributes
            attributes <- get_attributes(wb, sheet_name)

            # skip data if it's RekogMBKM
            if (grepl("RekogMBKM", attributes$class_name)) {
                next
            }

            # skip data if already exists (need to check empty or not)
            course_id <- paste0(attributes$course_id, " - ", attributes$course_name)
            if (course_id %in% names(attendance_data)) {
                next
            }


            # add course id to course_ids
            course_ids <- c(course_ids, course_id)

            # dataframes
            dataframe <- get_dataframe(wb, sheet_name)

            attendance_data[[course_id]] <- list(
                attributes = attributes,
                dataframe = dataframe
            )
        }
    }
    return(list(course_ids = course_ids, data = attendance_data))
}

read_file_upload_xls <- function(filepath) {
    return(get_attendance_data(filepath))
}

# define UI for application
ui <- fluidPage(
  h4("Rekap Absensi Mahasiswa"),
  h4("Program Studi Pendidikan Komputer"),
  h4("FKIP Universitas Lambung Mangkurat"),
  br(),
  fluidRow(
    column(
      width = 3,
      fluidRow(
        column(
          width = 12,
          fileInput("file", "Pilih File XLSX",
            accept = c(
              "application/vnd.ms-excel",
              ".xlsx"
            )
          ),
          uiOutput("course_ids_ui"),
          htmlOutput("course_attributes_ui")
        )
      )
    ),
    column(
      width = 9,
      uiOutput("attendance_table_ui")
    )
  ),
  br(),
  fluidRow(
    column(
      width = 9,
      offset = 3,
      uiOutput("attendance_plot_ui")
    )
  )
)

server <- function(input, output, session) {
  # states
  attendance_data <- reactive({
    req(input$file)
    return(get_attendance_data(input$file$datapath))
  })
  course_ids <- reactive(input$course_ids)

  # render select course id ui
  output$course_ids_ui <- renderUI({
    req(attendance_data())

    selectInput(
      "course_ids",
      "Pilih Mata Kuliah",
      choices = c("-", attendance_data()$course_ids)
    )
  })

  # render course attributes ui
  output$course_attributes_ui <- renderUI({
      req(attendance_data())
      req(course_ids())

      if (course_ids() != "-") {
          attributes <- attendance_data()$data[[course_ids()]]$attributes
          HTML(
              paste(
                  paste0("<p><b>Informasi</b></p>"),
                  # paste0("<p>Kode Mata Kuliah: ", attributes$course_id, "</p>"),
                  # paste0("<p>Mata Kuliah: ", attributes$course_name, "</p>"),
                  paste0("<p>Kelas: ", attributes$class_name, "</p>"),
                  paste0("<p>Periode: ", attributes$period, "</p>"),
                  paste0("</br>"),
                  paste0("<p><b>Keterangan:</b></p>"),
                  paste0("<p>- H: Hadir</p>"),
                  paste0("<p>- S: Sakit</p>"),
                  paste0("<p>- I: Izin</p>"),
                  paste0("<p>- T: Tanpa Keterangan</p>")
              )
          )
      }
  })

  # render attendance table ui
  output$attendance_table_ui <- renderUI({
      req(attendance_data())
      req(course_ids())

      if (course_ids() != "-") {
          dataframe <- attendance_data()$data[[course_ids()]]$dataframe

          renderDataTable(
            dataframe,
            options = list(
              scrollX = TRUE,
              pageLength = -1,
              columnDefs = list(
                list(targets = "_all", searchable = FALSE, orderable = FALSE),
                list(targets = c(2), width = "300px"),
                list(targets = c(4), width = "120px")
              ),
              lengthMenu = list(c(10, 25, 50, 100, -1), c('10', '25', '50', '100', 'All'))
            )
          )
      }
  })

  # render attendance plot ui
  output$attendance_plot_ui <- renderUI({
      req(attendance_data())
      req(course_ids())

      if (course_ids() != "-") {
          dataframe <- attendance_data()$data[[course_ids()]]$dataframe

          renderPlot({
              dataframe |>
                  ggplot(aes(x = NIM, y = `PERSENTASE HADIR`)) +
                  geom_line(group = 1, color = "cornflowerblue") +
                  geom_point(color = "cornflowerblue", size = 2) +
                  geom_abline(intercept = 80, slope = 0, color = "tomato") +
                  scale_y_continuous(limits = c(0, 100), breaks = c(0, 25, 50, 75, 80, 100)) +
                  labs(x = "", y = "") +
                  theme_bw() +
                  theme(
                    axis.text = element_text(size = 14),
                    axis.text.x = element_text(angle = 90, vjust = 0.5, hjust = 1)
                  )
          })
      }
  })
}

# Run the application
shinyApp(ui = ui, server = server)

