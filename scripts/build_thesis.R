# scripts/build_thesis.R
#
# End-to-end thesis build: Rmd -> .docx -> .pdf -> splice articles -> insert TOC.
#
# Steps:
#   1. Render thesis_manuscript.Rmd to .docx via bookdown::word_document2.
#   2. Convert the .docx to .pdf using LibreOffice (soffice --headless).
#   3. Source assemble_thesis_pdf.R to splice in the published-article PDFs at
#      the ARTICLE-INSERT-* markers, producing thesis_complete.pdf.
#   4. Build a Table of Contents PDF (one page) with page numbers harvested
#      from the rendered manuscript PDF, and splice it in after the cover.
#
# We render with `toc: false` and build the TOC ourselves because LibreOffice's
# headless PDF export does not update Word TOC fields, leaving them empty.
#
# Run from the project root:
#     Rscript scripts/build_thesis.R
#
# Requires R packages: rmarkdown, bookdown, pdftools, qpdf, officer, flextable.

suppressPackageStartupMessages({
  library(rmarkdown)
  library(officer)
  library(flextable)
  library(pdftools)
  library(qpdf)
})

# ---- Resolve paths ----
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

rmd_path    <- file.path(project_root, "thesis_manuscript.Rmd")
docx_path   <- file.path(project_root, "thesis_manuscript.docx")
pdf_path    <- file.path(project_root, "thesis_manuscript.pdf")
final_path  <- file.path(project_root, "thesis_complete.pdf")
ref_docx    <- file.path(project_root, "thesis_reference.docx")

# Sections to include in the TOC, in order. Names must match the heading text
# (case-sensitive) as it appears in the rendered PDF.
toc_sections <- c(
  "Acknowledgments",
  "List of Publications",
  "Abstract",
  "Résumé en français",
  "Outline",
  "General Introduction",
  "Methodological Contributions",
  "Papers",
  "Conclusions and Perspectives",
  "Bibliography"
)

# ---- Locate LibreOffice ----
soffice <- Sys.which("soffice")
if (!nzchar(soffice)) {
  mac_app <- "/Applications/LibreOffice.app/Contents/MacOS/soffice"
  if (file.exists(mac_app)) soffice <- mac_app
}
if (!nzchar(soffice)) {
  stop("LibreOffice 'soffice' not found on PATH. Install LibreOffice or add it to PATH.")
}

# Shared LibreOffice user-installation profile (avoids conflicts with a running
# LibreOffice GUI and is reused across both PDF conversions in this script).
profile_dir <- tempfile("lo_profile_")
dir.create(profile_dir)
on.exit(unlink(profile_dir, recursive = TRUE), add = TRUE)

soffice_convert <- function(input, outdir) {
  status <- system2(
    soffice,
    args = c(
      sprintf("-env:UserInstallation=file://%s", profile_dir),
      "--headless",
      "--convert-to", "pdf",
      "--outdir", outdir,
      input
    ),
    stdout = TRUE, stderr = TRUE
  )
  if (!is.null(attr(status, "status")) && attr(status, "status") != 0) {
    cat(status, sep = "\n")
    stop("LibreOffice conversion failed for: ", input)
  }
}

# ---- 1. Render Rmd -> .docx ----
cat("[1/4] Rendering ", basename(rmd_path), " -> .docx\n", sep = "")
rmarkdown::render(
  rmd_path,
  output_format = "bookdown::word_document2",
  output_file   = basename(docx_path),
  output_dir    = project_root,
  quiet         = TRUE
)

# ---- 2. Convert .docx -> .pdf via LibreOffice ----
cat("[2/4] Converting .docx -> .pdf via LibreOffice\n")
soffice_convert(docx_path, project_root)
if (!file.exists(pdf_path)) {
  stop("Expected PDF not produced at: ", pdf_path)
}

# ---- 3. Splice in article PDFs (creates thesis_complete.pdf) ----
cat("[3/4] Splicing article PDFs at ARTICLE-INSERT-* markers\n")
source(file.path(project_root, "scripts", "assemble_thesis_pdf.R"))

# ---- 4. Build TOC and splice it in after the cover ----
cat("[4/4] Building TOC page and inserting after cover\n")

# Find each section's page number in the manuscript PDF. These are the page
# numbers as printed in the body footer — readers locate sections by these.
# Sections appear in the same order as toc_sections, so we search forward only
# (after the previous match) to avoid picking up false positives from the
# Outline section, which lists upcoming section names as bullet points.
pages_text <- pdftools::pdf_text(pdf_path)
heading_page_after <- function(name, start_page) {
  for (p in start_page:length(pages_text)) {
    # Heading rendered by Word appears within the first ~3 lines of its page.
    head <- substr(pages_text[p], 1, 80)
    if (grepl(name, head, fixed = TRUE)) return(p)
  }
  NA_integer_
}
section_pages <- integer(length(toc_sections))
last_page <- 1L
for (i in seq_along(toc_sections)) {
  p <- heading_page_after(toc_sections[i], last_page)
  section_pages[i] <- if (is.null(p)) NA_integer_ else p
  if (!is.na(section_pages[i])) last_page <- section_pages[i]
}
missing <- toc_sections[is.na(section_pages)]
if (length(missing)) {
  warning("TOC: could not locate these sections in the rendered PDF: ",
          paste(missing, collapse = ", "))
}

# Build the TOC docx — a single heading + a 2-column flextable (section, page)
toc_df <- data.frame(
  Section = toc_sections,
  Page    = ifelse(is.na(section_pages), "", as.character(section_pages)),
  stringsAsFactors = FALSE
)

ft <- flextable(toc_df) |>
  delete_part(part = "header") |>
  width(j = 1, width = 5.5) |>
  width(j = 2, width = 0.7) |>
  align(j = 1, align = "left",  part = "body") |>
  align(j = 2, align = "right", part = "body") |>
  border_remove() |>
  fontsize(size = 12, part = "body") |>
  padding(padding.top = 4, padding.bottom = 4, part = "body")

# Use a fresh docx (officer default). We override the section to have an empty
# footer so the TOC page itself doesn't carry a body-style page number.
toc_doc <- read_docx() |>
  body_set_default_section(prop_section(
    page_size      = page_size(width = 8.27, height = 11.69, orient = "portrait"),
    page_margins   = page_mar(top = 1, bottom = 1, left = 1, right = 1,
                              header = 0.5, footer = 0.5, gutter = 0),
    type           = "continuous",
    footer_default = block_list(fpar(ftext("")))
  )) |>
  body_add_fpar(fpar(
    ftext("Table of Contents", prop = fp_text(font.size = 20, bold = TRUE)),
    fp_p = fp_par(text.align = "left", padding.bottom = 12)
  )) |>
  body_add_flextable(ft)

toc_docx <- file.path(project_root, "_toc.docx")
toc_pdf  <- file.path(project_root, "_toc.pdf")
print(toc_doc, target = toc_docx)
on.exit(unlink(c(toc_docx, toc_pdf)), add = TRUE)

soffice_convert(toc_docx, project_root)
if (!file.exists(toc_pdf)) {
  stop("Expected TOC PDF not produced at: ", toc_pdf)
}
if (pdftools::pdf_length(toc_pdf) > 1) {
  warning("TOC PDF spans ", pdftools::pdf_length(toc_pdf),
          " pages; consider trimming entries or font size")
}

# Splice TOC PDF in after the cover (page 1) of thesis_complete.pdf
n <- pdftools::pdf_length(final_path)
cover_pdf <- tempfile(fileext = ".pdf")
rest_pdf  <- tempfile(fileext = ".pdf")
on.exit(unlink(c(cover_pdf, rest_pdf)), add = TRUE)

pdftools::pdf_subset(final_path, pages = 1, output = cover_pdf)
pdftools::pdf_subset(final_path, pages = 2:n, output = rest_pdf)
qpdf::pdf_combine(c(cover_pdf, toc_pdf, rest_pdf), output = final_path)

cat("Final thesis written to: ", final_path,
    " (", pdftools::pdf_length(final_path), " pages)\n", sep = "")
