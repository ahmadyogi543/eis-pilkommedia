# load necessary libraries
library(dplyr)
library(ggplot2)
library(readxl)
library(openxlsx)
library(shiny)

# helpers
read_file_upload_xls <- function(filepath) {
  gpas <- read_xls(filepath) |>
    rename_with(~gsub("\\ ", ".", .), everything()) |>
    select(
      -BOBOT,
      -JUMLAH.SKS,
      -SKS.MATAKULIAH.WAJIB.LULUS,
      -SKS.PILIHAN.LULUS,
      -STATUS.MAHASISWA,
      -Terakhir.Update
    ) |>
    mutate(
      IPK = round(as.numeric(IPK), digits = 2),
      IP.SEMESTER = round(as.numeric(IP.SEMESTER), digits = 2),
      SKS.SEMESTER = as.numeric(SKS.SEMESTER),
      SKS.LULUS = as.numeric(SKS.LULUS),
      SKS.TIDAK.LULUS = as.numeric(SKS.TIDAK.LULUS),
      SKS.TOTAL = as.numeric(SKS.TOTAL)
    )
  
  return(gpas)
}

get_generation_gpas_table <- function(gpas, generation) {
  get_gpas_pass_status <- function(IPK, SKS_TIDAK_LULUS) {
    status <- ifelse(
      IPK >= 3.51 & IPK <= 4.0,
      ifelse(SKS_TIDAK_LULUS == 0, "Pujian (3,51 - 4,00)", "Sangat Memuaskan (3,01 - 3,50)"),
      ifelse(
        IPK >= 3.01 & IPK <= 3.50, "Sangat Memuaskan (3,01 - 3,50)",
        ifelse(
          IPK >= 2.76 & IPK <= 3.0, "Memuaskan (2,76 - 3,00)", "Tidak Lulus (< 2,76)"
        )
      )
    )
    return(status)
  }

 get_remainder_credits <- function(generation, credits) {
   remainder_credits <- numeric(length(generation))

   for (i in seq_along(generation)) {
     if (generation[i] < 2017) {
       remainder_credits[i] <- 149 - credits[i]
     } else if (generation[i] < 2020) {
       remainder_credits[i] <- 145 - credits[i]
     } else {
       remainder_credits[i] <- 146 - credits[i]
     }

     if (remainder_credits[i] < 0) {
       remainder_credits[i] = 0
     }
   }

    return(remainder_credits)
  }
 

  gpas_table <- gpas |>
    filter(ANGKATAN == generation) |>
    mutate(PREDIKAT.KELULUSAN = get_gpas_pass_status(IPK, SKS.TIDAK.LULUS)) |>
    mutate(PREDIKAT.KELULUSAN = factor(
      PREDIKAT.KELULUSAN,
      ordered = TRUE,
      levels = c(
        "Pujian (3,51 - 4,00)",
        "Sangat Memuaskan (3,01 - 3,50)",
        "Memuaskan (2,76 - 3,00)",
        "Tidak Lulus (< 2,76)"
      )
    )) |>
    mutate(SISA.SKS = get_remainder_credits(ANGKATAN, SKS.TOTAL)) |>
    select(-ANGKATAN, -PERIODE) |>
    arrange(desc(IPK)) |>
    mutate(NO = row_number()) |>
    select(NO, everything()) |>
    rename_with(~gsub("\\.", " ", .), everything())
  return (gpas_table)
}

get_summarized_gpas_table <- function(gpas_table) {
  summarized_gpas_table <- gpas_table |>
    group_by(`PREDIKAT KELULUSAN`) |>
    summarize(
      FREKUENSI = n(),
      PERSENTASE = round((n() / nrow(gpas_table)) *100, digits = 2)
    ) |>
    ungroup()
    
  total_row <- summarized_gpas_table |>
    summarise(
      FREKUENSI = sum(FREKUENSI),
      PERSENTASE = round(sum(PERSENTASE)),
      `PREDIKAT KELULUSAN` = "TOTAL"
    )

  summarized_gpas_table <- bind_rows(summarized_gpas_table, total_row)

  return(summarized_gpas_table)
}

get_gpas_period <- function(gpas) {
  format_gpas_period <- function(period) {
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

  gpas_info <- gpas |>
    select(PERIODE) |>
    mutate(
      PERIODE = format_gpas_period(PERIODE)
    ) |>
    slice(1)

  return(gpas_info$PERIODE)
}

write_xlsx_title <- function(wb, sheet_name, col_size, period, generation) {
  titles <- c(
    "Rekap IPK Mahasiswa Per Semester",
    "Progam Studi Pendidikan Komputer",
    "FKIP Universitas Lambung Mangkurat",
    "",
    paste0("Periode: ", period),
    paste0("Angkatan: ", generation),
    ""
  )

  for (i in seq_along(titles)) {
    writeData(wb, sheet = sheet_name, x = titles[i], startRow = i, startCol = 1)
    mergeCells(wb, sheet = sheet_name, cols = 1:col_size, rows = i)
  }
}

export_to_xlsx <- function(gpas_table,summarized_gpas_table, period, generation) {
  wb <- createWorkbook()
  addWorksheet(wb, "data")

  #title
  write_xlsx_title(
    wb,
    "data",
    ncol(gpas_table),
    period,
    generation 
  )

  # data
  writeData(wb, sheet = "data", x = gpas_table |> select(-`PREDIKAT KELULUSAN`), startRow = 8, startCol = 1)
  writeData(wb, sheet = "data", x = summarized_gpas_table, startRow = 8 + nrow(gpas_table) + 2, startCol = 3)

  # plot
  plot <- gpas_table |>
      ggplot(aes(x = NIM, y = IPK)) +
        geom_point(alpha = 1.0, size = 3, aes(, color = `PREDIKAT KELULUSAN`)) +
        scale_color_manual(
          values = c(
            "Pujian (3,51 - 4,00)" = "limegreen",
            "Sangat Memuaskan (3,01 - 3,50)" = "cornflowerblue",
            "Memuaskan (2,76 - 3,00)" = "gold",
            "Tidak Lulus (< 2,76)" = "tomato"
          )
        ) +
        geom_line(group = 1, alpha = 0.25) +
        geom_abline(intercept = 2.76, slope = 0, color = "tomato", alpha = 0.5) +
        scale_y_continuous(limits = c(0, 4), breaks = c(0, 1, 2, 2.76, 3, 4)) +
        labs(x = "", y = "") +
        theme_bw() +
        theme(
	      axis.text = element_text(size = 14),
          legend.text = element_text(size = 14),
	      legend.title = element_text(size = 14, face = "bold"),
	      axis.text.x = element_text(angle = 90, hjust = 1, vjust = 0.5)
	    )
  ggsave("plot.png", plot = plot, width = 12, height = 5, units = "in")

  insertImage(
    wb,
    sheet = "data",
    "plot.png",
    startRow = 8 + nrow(gpas_table) + nrow(summarized_gpas_table) + 4,
    startCol = 1,
    width = 12,
    height = 5
  )

  # set column widths
  setColWidths(
    wb,
    sheet = "data",
    cols = 1:10,
    widths = c(5, 18, 35, 15, 15, 15, 15, 15, 20, 15)
  )

  tmp <- tempfile(fileext = ".xlsx")
  saveWorkbook(wb, file = tmp)
  file.remove("plot.png")

  return(tmp)
}

# define UI for application
ui <- fluidPage(
  h4("Rekap IPK Mahasiswa"),
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
          uiOutput("generation_ui"),
          htmlOutput("gpas_period_ui"),
          br(),
          uiOutput("export_to_xlsx_button_ui")
        )
      )
    ),
    column(
      width = 9,
      uiOutput("generation_gpas_table_ui")
    )
  ),
  fluidRow(
    column(
      width = 9,
      offset = 3,
      br(),
      uiOutput("generation_summarized_gpas_table_ui")
    )
  ),
  fluidRow(
    column(
      width = 12,
      offset = 0,
      br(),
      uiOutput("generation_gpas_plot_ui")
    )
  ),
  br()
)

server <- function(input, output, session) {
  # states
  gpas <- reactive({
    req(input$file)
    return(read_file_upload_xls(input$file$datapath))
  })
  generation <- reactive(input$generation)
  generation_gpas_table <- reactive({
    req(gpas())
    req(generation())

    if (generation() != "-") {
      generation_gpas_table <- get_generation_gpas_table(gpas(), generation())
      return(generation_gpas_table)
    } else {
        return(NULL)
    }
  })
  generation_gpas_table_no <- reactive(seq_len(nrow(generation_gpas_table())))

  # render select generation ui
  output$generation_ui <- renderUI({
    req(gpas())
    
    # get choices
    choices <- gpas() |>
      select(ANGKATAN) |>
      add_row(ANGKATAN = "-") |>
      arrange(ANGKATAN) |>
      distinct(ANGKATAN)
     
    selectInput(
      "generation",
      "Pilih Angkatan",
      choices = choices
    )
  })

  # render gpas period ui
  output$gpas_period_ui <- renderUI({
    req(gpas())

    gpas_period <- get_gpas_period(gpas())

    HTML(
      paste(
        paste0("<p><b>Informasi</b></p>"),
        paste0("<p>Periode: ", gpas_period, "</p>")
      )
    )
  })

  # render export_to_xlsx_button_ui
  output$export_to_xlsx_button_ui <- renderUI({
    req(gpas())
    req(generation())

    if (generation() != "-") {
      downloadButton("export_to_xlsx", "Export to XLSX")
    }
  })

  # export xlsx
  output$export_to_xlsx <- downloadHandler(
    filename = "export.xlsx",
    content = function(file) {
      gpas_period <- get_gpas_period(gpas())
      tmp <- export_to_xlsx(
        generation_gpas_table(),
        get_summarized_gpas_table(generation_gpas_table()),
        gpas_period,
        generation()
      )

      file.copy(tmp, file)
    }
  )

  # render gpas table
  output$generation_gpas_table_ui <- renderUI({
    req(gpas())
    req(generation())

    if (generation() != "-") {
      renderDataTable(
        generation_gpas_table() |> select(-`PREDIKAT KELULUSAN`),
        options = list(
          pageLength = 10,
          columnDefs = list(
            list(targets = "_all", searchable = FALSE, orderable = FALSE),
            list(targets = 5, width = "80px"),
            list(targets = 2, width = "100px")
          ),
          lengthMenu = list(c(10, 25, 50, 100, -1), c('10', '25', '50', '100', 'All'))
        )
      ) 
    }
  })

  # render summarized gpas table
  output$generation_summarized_gpas_table_ui <- renderUI({
    req(gpas())
    req(generation())

    if (generation() != "-") {
      summarized_gpas_table <- get_summarized_gpas_table(generation_gpas_table())
      renderTable(summarized_gpas_table)
    }
  })

  # render gpas plot
  output$generation_gpas_plot_ui <- renderUI({
    req(gpas())
    req(generation())

    if (generation() != "-") {
      renderPlot({
        generation_gpas_table() |>
          ggplot(aes(x = NIM, y = IPK)) +
            geom_point(alpha = 1.0, size = 3, aes(, color = `PREDIKAT KELULUSAN`)) +
            scale_color_manual(
              values = c(
                "Pujian (3,51 - 4,00)" = "limegreen",
                "Sangat Memuaskan (3,01 - 3,50)" = "cornflowerblue",
                "Memuaskan (2,76 - 3,00)" = "gold",
                "Tidak Lulus (< 2,76)" = "tomato"
              )
            ) +
            # geom_text(aes(label = sprintf("%.2f", IPK)), vjust = -0.5) +
            geom_line(group = 1, alpha = 0.25) +
            geom_abline(intercept = 2.76, slope = 0, color = "tomato", alpha = 0.5) +
            scale_y_continuous(limits = c(0, 4), breaks = c(0, 1, 2, 2.76, 3, 4)) +
            labs(x = "", y = "") +
            theme_bw() +
            theme(
	          axis.text = element_text(size = 14),
              legend.text = element_text(size = 14),
	          legend.title = element_text(size = 14, face = "bold"),
	          axis.text.x = element_text(angle = 90, hjust = 1, vjust = 0.5)
	        )
      }) 
    }
  })
}

# Run the application
shinyApp(ui = ui, server = server)

