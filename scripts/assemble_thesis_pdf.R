# scripts/assemble_thesis_pdf.R
#
# Splices the published article PDFs into the rendered thesis PDF at the
# marker positions defined in thesis_manuscript.Rmd.
#
# Workflow:
#   1. Render the thesis to .docx (rmarkdown::render or RStudio "Knit"):
#        rmarkdown::render("thesis_manuscript.Rmd",
#                          output_format = "bookdown::word_document2")
#
#   2. Convert the .docx to .pdf with your tool of choice. On macOS with
#      LibreOffice installed, e.g.:
#        soffice --headless --convert-to pdf thesis_manuscript.docx
#      (Or open in Word and "Save As PDF".)
#
#   3. Run this script:
#        Rscript scripts/assemble_thesis_pdf.R
#
# Requires R packages: pdftools, qpdf
#   install.packages(c("pdftools", "qpdf"))

suppressPackageStartupMessages({
  library(pdftools)
  library(qpdf)
})

# ---- Configuration ----
# Resolve paths relative to this script's location so the script works whether
# run via `Rscript scripts/assemble_thesis_pdf.R` or sourced from the project root.
script_path <- tryCatch(
  normalizePath(sys.frame(1)$ofile, mustWork = TRUE),
  error = function(e) {
    cmd_args <- commandArgs(trailingOnly = FALSE)
    file_arg <- grep("^--file=", cmd_args, value = TRUE)
    if (length(file_arg)) sub("^--file=", "", file_arg[1]) else NA_character_
  }
)
project_root <- if (!is.na(script_path)) {
  normalizePath(file.path(dirname(script_path), ".."), mustWork = FALSE)
} else {
  getwd()
}

input_pdf  <- file.path(project_root, "thesis_manuscript.pdf")
output_pdf <- file.path(project_root, "thesis_complete.pdf")

# Each entry: marker = unique string placed in thesis_manuscript.Rmd;
# file = path to the article PDF to splice in at the marker page.
articles <- list(
  list(marker = "ARTICLE-INSERT-1", file = file.path(project_root, "viral_suppression_paper.pdf")),
  list(marker = "ARTICLE-INSERT-2", file = file.path(project_root, "prep_paper.pdf")),
  list(marker = "ARTICLE-INSERT-3", file = file.path(project_root, "tgc_paper_2026_04_26.pdf"))
)

# ---- Sanity checks ----
if (!file.exists(input_pdf)) {
  stop("Rendered thesis PDF not found at: ", input_pdf,
       "\nRender the .Rmd to .docx first, then convert .docx to .pdf.")
}
for (a in articles) {
  if (!file.exists(a$file)) stop("Article PDF not found: ", a$file)
}

# ---- Locate marker pages in the rendered thesis ----
pages_text <- pdf_text(input_pdf)
marker_pages <- vapply(articles, function(a) {
  hits <- which(grepl(a$marker, pages_text, fixed = TRUE))
  if (length(hits) == 0L) {
    stop("Marker not found in thesis PDF: ", a$marker,
         "\nCheck that the marker text in thesis_manuscript.Rmd survived rendering.")
  }
  if (length(hits) > 1L) {
    warning("Multiple pages match marker '", a$marker, "'; using first (page ", hits[1], ").")
  }
  hits[1]
}, integer(1))

if (is.unsorted(marker_pages)) {
  stop("Markers must appear in order in the thesis. Got pages: ",
       paste(marker_pages, collapse = ", "))
}

# ---- Build merged PDF ----
n_pages <- pdf_length(input_pdf)
segments <- list()
prev_end <- 0L

for (i in seq_along(articles)) {
  page <- marker_pages[i]
  # Thesis content before this marker page
  if (page > prev_end + 1L) {
    segments[[length(segments) + 1L]] <- list(
      file = input_pdf,
      pages = (prev_end + 1L):(page - 1L)
    )
  }
  # The article itself (the marker page is dropped)
  segments[[length(segments) + 1L]] <- list(
    file = articles[[i]]$file,
    pages = NULL  # all pages
  )
  prev_end <- page
}

# Remaining thesis content after the last marker
if (prev_end < n_pages) {
  segments[[length(segments) + 1L]] <- list(
    file = input_pdf,
    pages = (prev_end + 1L):n_pages
  )
}

# ---- Write segments to temp files and combine ----
tmp_files <- character(length(segments))
on.exit(unlink(tmp_files), add = TRUE)

for (i in seq_along(segments)) {
  s <- segments[[i]]
  tmp_files[i] <- tempfile(fileext = ".pdf")
  if (is.null(s$pages)) {
    file.copy(s$file, tmp_files[i], overwrite = TRUE)
  } else {
    pdf_subset(s$file, pages = s$pages, output = tmp_files[i])
  }
}

pdf_combine(tmp_files, output = output_pdf)
cat("Assembled thesis written to: ", output_pdf, "\n", sep = "")
