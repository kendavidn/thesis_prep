# scripts/make_reference_docx.R
#
# Build thesis_reference.docx — a pandoc reference document used by
# rmarkdown::render() to control Word styling.
#
# What this adds on top of pandoc's default reference.docx:
#   * A footer with a centered page number (page X) on every page.
#   * A "Centered" paragraph style usable in the Rmd via
#         ::: {custom-style="Centered"}
#         ...content...
#         :::
#     (used by the French cover page).
#
# Run from the project root:
#     Rscript scripts/make_reference_docx.R

suppressPackageStartupMessages({
  library(officer)
  library(xml2)
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
out_path <- file.path(project_root, "thesis_reference.docx")

# ---- 1. Get pandoc's default reference.docx ----
default_ref <- tempfile(fileext = ".docx")
res <- system2("pandoc",
               args = c("-o", shQuote(default_ref),
                        "--print-default-data-file=reference.docx"))
if (res != 0 || !file.exists(default_ref)) {
  stop("Failed to extract pandoc's default reference.docx")
}

# ---- 2. Add a centered page-number footer via officer ----
doc <- read_docx(default_ref)

footer_block <- block_list(
  fpar(run_word_field("PAGE"), fp_p = fp_par(text.align = "center"))
)

# body_set_default_section sets the section properties for the document's
# default section, which is what pandoc copies into the rendered output.
doc <- body_set_default_section(
  doc,
  prop_section(
    page_size      = page_size(width = 8.27, height = 11.69, orient = "portrait"),
    page_margins   = page_mar(top = 1, bottom = 1, left = 1, right = 1,
                              header = 0.5, footer = 0.5, gutter = 0),
    type           = "continuous",
    footer_default = footer_block
  )
)
print(doc, target = out_path)

# ---- 3. Inject a "Centered" paragraph style into styles.xml ----
# pandoc lets us reference custom styles from markdown via
# ::: {custom-style="Centered"} ... :::
# but the style must exist in the reference document.
work_dir <- tempfile("ref_docx_")
dir.create(work_dir)
on.exit(unlink(work_dir, recursive = TRUE), add = TRUE)

unzipped <- utils::unzip(out_path, exdir = work_dir)
styles_xml_path <- file.path(work_dir, "word", "styles.xml")
if (!file.exists(styles_xml_path)) stop("styles.xml not found in reference docx")

styles_doc <- read_xml(styles_xml_path)
ns <- c(w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")

if (length(xml_find_all(styles_doc, "//w:style[@w:styleId='Centered']", ns)) == 0L) {
  centered_xml <- read_xml(paste0(
    '<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ',
    'w:type="paragraph" w:customStyle="1" w:styleId="Centered">',
    '<w:name w:val="Centered"/>',
    '<w:basedOn w:val="Normal"/>',
    '<w:qFormat/>',
    '<w:pPr><w:jc w:val="center"/></w:pPr>',
    '</w:style>'
  ))
  xml_add_child(xml_root(styles_doc), centered_xml)
  write_xml(styles_doc, styles_xml_path)

  # Re-zip the docx (must preserve internal structure; zip from inside work_dir)
  old_wd <- getwd()
  on.exit(setwd(old_wd), add = TRUE)
  setwd(work_dir)
  files <- list.files(recursive = TRUE, all.files = TRUE, no.. = TRUE)
  unlink(out_path)
  utils::zip(out_path, files = files, flags = "-q -X")
  setwd(old_wd)
}

cat("Reference docx written to: ", out_path, "\n", sep = "")
