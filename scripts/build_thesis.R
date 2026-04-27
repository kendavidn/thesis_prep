# scripts/build_thesis.R
#
# End-to-end thesis build: Rmd -> .docx -> .pdf -> spliced .pdf with articles.
#
# Steps:
#   1. Render thesis_manuscript.Rmd to .docx via bookdown::word_document2.
#   2. Convert the .docx to .pdf using LibreOffice (soffice --headless).
#   3. Source assemble_thesis_pdf.R to splice in the published-article PDFs at
#      the ARTICLE-INSERT-* markers, producing thesis_complete.pdf.
#
# Run from the project root:
#     Rscript scripts/build_thesis.R
#
# Requires: rmarkdown, bookdown, pdftools, qpdf, plus LibreOffice (`soffice`).

suppressPackageStartupMessages({
  library(rmarkdown)
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

rmd_path  <- file.path(project_root, "thesis_manuscript.Rmd")
docx_path <- file.path(project_root, "thesis_manuscript.docx")
pdf_path  <- file.path(project_root, "thesis_manuscript.pdf")

# ---- Locate LibreOffice ----
soffice <- Sys.which("soffice")
if (!nzchar(soffice)) {
  mac_app <- "/Applications/LibreOffice.app/Contents/MacOS/soffice"
  if (file.exists(mac_app)) soffice <- mac_app
}
if (!nzchar(soffice)) {
  stop("LibreOffice 'soffice' not found on PATH. Install LibreOffice or add it to PATH.")
}

# ---- 1. Render Rmd -> .docx ----
cat("[1/3] Rendering ", basename(rmd_path), " -> .docx\n", sep = "")
rmarkdown::render(
  rmd_path,
  output_format = "bookdown::word_document2",
  output_file   = basename(docx_path),
  output_dir    = project_root,
  quiet         = TRUE
)

# ---- 2. Convert .docx -> .pdf via LibreOffice ----
cat("[2/3] Converting .docx -> .pdf via LibreOffice\n")
# Use a unique user profile to avoid conflicts with a running LibreOffice GUI.
profile_dir <- tempfile("lo_profile_")
dir.create(profile_dir)
on.exit(unlink(profile_dir, recursive = TRUE), add = TRUE)

status <- system2(
  soffice,
  args = c(
    sprintf("-env:UserInstallation=file://%s", profile_dir),
    "--headless",
    "--convert-to", "pdf",
    "--outdir", project_root,
    docx_path
  ),
  stdout = TRUE, stderr = TRUE
)
if (!is.null(attr(status, "status")) && attr(status, "status") != 0) {
  cat(status, sep = "\n")
  stop("LibreOffice conversion failed.")
}
if (!file.exists(pdf_path)) {
  cat(status, sep = "\n")
  stop("Expected PDF not produced at: ", pdf_path)
}

# ---- 3. Splice in article PDFs ----
cat("[3/3] Splicing article PDFs at ARTICLE-INSERT-* markers\n")
source(file.path(project_root, "scripts", "assemble_thesis_pdf.R"))
