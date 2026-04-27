# scripts/make_reference_docx.R
#
# Build thesis_reference.docx — a pandoc reference document used by
# rmarkdown::render() to control Word styling.
#
# What this builds on top of pandoc's default reference.docx:
#   * A footer with a centered page number (page X) on every page.
#   * Custom paragraph styles usable in the Rmd via
#         ::: {custom-style="<name>"}
#         ...content...
#         :::
#     - "Centered"       — left-align cleared, paragraph centered.
#     - "CoverTitle"     — large bold UNIGE pink, centered (cover title).
#     - "CoverSubtitle"  — italic UNIGE pink, centered (cover subtitle).
#   * Heading 1/2/3 recoloured to UNIGE pink (#CF0063) and forced bold.
#   * Normal paragraph style set to 1.5x line spacing (everything in the body
#     gets 1.5x by default; the spliced-in article PDFs are unaffected).
#
# Run from the project root:
#     Rscript scripts/make_reference_docx.R

suppressPackageStartupMessages({
  library(officer)
  library(xml2)
})

UNIGE_PINK <- "CF0063"   # Pantone 214C — UNIGE brand colour
W_NS <- "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

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

# ---- 3. Edit styles.xml ----
work_dir <- tempfile("ref_docx_")
dir.create(work_dir)
on.exit(unlink(work_dir, recursive = TRUE), add = TRUE)
utils::unzip(out_path, exdir = work_dir)

styles_xml_path <- file.path(work_dir, "word", "styles.xml")
if (!file.exists(styles_xml_path)) stop("styles.xml not found in reference docx")
styles_doc <- read_xml(styles_xml_path)
ns <- c(w = W_NS)

# Helper: build a w:* element with the right namespace declared.
w_node <- function(xml_string) {
  read_xml(sprintf('<%s xmlns:w="%s"/>',
                   sub("/?$", "", sub("^<", "", strsplit(xml_string, "[ />]")[[1]][1])),
                   W_NS))
}

# Helper: ensure the <w:rPr> child of `style` has <w:color w:val="..."/> and <w:b/>.
# We replace any existing <w:color> rather than mutating attributes, because
# xml2's xml_set_attr doesn't reliably reach namespaced attributes (it ends up
# adding a sibling `val=` without the `w:` prefix, which Word ignores), and the
# default <w:color> carries themeColor/themeShade overrides we'd need to clear.
set_heading_color_and_bold <- function(style_node, color_hex) {
  rPr <- xml_find_first(style_node, "./w:rPr", ns)
  if (inherits(rPr, "xml_missing")) {
    xml_add_child(style_node, read_xml(sprintf('<w:rPr xmlns:w="%s"/>', W_NS)))
    rPr <- xml_find_first(style_node, "./w:rPr", ns)
  }
  existing_color <- xml_find_first(rPr, "./w:color", ns)
  if (!inherits(existing_color, "xml_missing")) xml_remove(existing_color)
  xml_add_child(rPr, read_xml(sprintf(
    '<w:color xmlns:w="%s" w:val="%s"/>', W_NS, color_hex)))
  if (inherits(xml_find_first(rPr, "./w:b", ns), "xml_missing")) {
    xml_add_child(rPr, read_xml(sprintf('<w:b xmlns:w="%s"/>', W_NS)))
  }
}

# 3a. Recolour Heading 1/2/3 to UNIGE pink + force bold.
for (sid in c("Heading1", "Heading2", "Heading3")) {
  s <- xml_find_first(styles_doc, sprintf("//w:style[@w:styleId='%s']", sid), ns)
  if (!inherits(s, "xml_missing")) set_heading_color_and_bold(s, UNIGE_PINK)
}

# 3b. Set 1.5x line spacing on Normal style.
normal <- xml_find_first(styles_doc, "//w:style[@w:styleId='Normal']", ns)
if (!inherits(normal, "xml_missing")) {
  pPr <- xml_find_first(normal, "./w:pPr", ns)
  if (inherits(pPr, "xml_missing")) {
    pPr <- read_xml(sprintf('<w:pPr xmlns:w="%s"/>', W_NS))
    xml_add_child(normal, pPr)
    pPr <- xml_find_first(normal, "./w:pPr", ns)
  }
  spacing <- xml_find_first(pPr, "./w:spacing", ns)
  if (inherits(spacing, "xml_missing")) {
    xml_add_child(pPr, read_xml(sprintf(
      '<w:spacing xmlns:w="%s" w:line="360" w:lineRule="auto"/>', W_NS)))
  } else {
    xml_set_attr(spacing, "line",     "360")    # 360 = 1.5x in twentieths
    xml_set_attr(spacing, "lineRule", "auto")
  }
}

# 3c. Add custom paragraph styles if not already present.
add_style <- function(xml_string) {
  # Insert as a child of the styles root so pandoc can resolve it.
  xml_add_child(xml_root(styles_doc), read_xml(xml_string))
}

if (length(xml_find_all(styles_doc, "//w:style[@w:styleId='Centered']", ns)) == 0L) {
  add_style(sprintf(paste0(
    '<w:style xmlns:w="%s" w:type="paragraph" w:customStyle="1" w:styleId="Centered">',
    '<w:name w:val="Centered"/>',
    '<w:basedOn w:val="Normal"/>',
    '<w:qFormat/>',
    '<w:pPr><w:jc w:val="center"/></w:pPr>',
    '</w:style>'), W_NS))
}

if (length(xml_find_all(styles_doc, "//w:style[@w:styleId='CoverTitle']", ns)) == 0L) {
  add_style(sprintf(paste0(
    '<w:style xmlns:w="%s" w:type="paragraph" w:customStyle="1" w:styleId="CoverTitle">',
    '<w:name w:val="Cover Title"/>',
    '<w:basedOn w:val="Normal"/>',
    '<w:qFormat/>',
    # 280 line ~= 1.15x to keep multi-line titles tight on the cover.
    '<w:pPr>',
    '<w:jc w:val="center"/>',
    '<w:spacing w:before="160" w:after="160" w:line="280" w:lineRule="auto"/>',
    '</w:pPr>',
    '<w:rPr>',
    '<w:b/>',
    '<w:color w:val="%s"/>',
    '<w:sz w:val="36"/>',     # 18pt (sz is half-points)
    '</w:rPr>',
    '</w:style>'), W_NS, UNIGE_PINK))
}

if (length(xml_find_all(styles_doc, "//w:style[@w:styleId='CoverSubtitle']", ns)) == 0L) {
  add_style(sprintf(paste0(
    '<w:style xmlns:w="%s" w:type="paragraph" w:customStyle="1" w:styleId="CoverSubtitle">',
    '<w:name w:val="Cover Subtitle"/>',
    '<w:basedOn w:val="Normal"/>',
    '<w:qFormat/>',
    '<w:pPr>',
    '<w:jc w:val="center"/>',
    '<w:spacing w:before="80" w:after="80" w:line="280" w:lineRule="auto"/>',
    '</w:pPr>',
    '<w:rPr>',
    '<w:i/>',
    '<w:color w:val="%s"/>',
    '<w:sz w:val="26"/>',     # 13pt
    '</w:rPr>',
    '</w:style>'), W_NS, UNIGE_PINK))
}

write_xml(styles_doc, styles_xml_path)

# ---- 4. Re-zip the modified docx ----
old_wd <- getwd()
on.exit(setwd(old_wd), add = TRUE)
setwd(work_dir)
files <- list.files(recursive = TRUE, all.files = TRUE, no.. = TRUE)
unlink(out_path)
utils::zip(out_path, files = files, flags = "-q -X")
setwd(old_wd)

cat("Reference docx written to: ", out_path, "\n", sep = "")
