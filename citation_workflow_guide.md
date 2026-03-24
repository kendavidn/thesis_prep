# Citation Resolution Workflow for R Markdown Theses

Guide for LLM assistants cleaning up citations in `.Rmd` manuscripts that use `bibliography: references.bib`.

## Pipeline

1. **For DOIs: use Crossref content negotiation, not Crossref search.** `curl -LH "Accept: application/x-bibtex" https://doi.org/{DOI}` returns clean BibTeX in one step. This is the single most reliable tool in the pipeline. However, beware of GBD-style mega-author papers — the returned BibTeX can be 60KB+ of author names. Truncate to first author + `and others` before writing to the `.bib` file; the CSL style handles "et al." rendering.

2. **For PMC links: convert to DOI first, then use content negotiation.** The NCBI ID converter (`https://pmc.ncbi.nlm.nih.gov/tools/idconv/api/v1/articles/?ids=PMC{ID}&format=json`) returns the DOI for a given PMCID. Then proceed with step 1. Do not try to construct BibTeX manually from the PMC page.

3. **Crossref search is unreliable for finding specific papers.** Queries like "GBD 2023 Lancet 2025" returned irrelevant results across multiple attempts with different keyword combinations. When Crossref search fails, fall back to PubMed E-utilities (`eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi`) which handles biomedical literature much better, or use a web search as a last resort to find the DOI, then resolve it via content negotiation. PubMed search worked well for finding Schwartz et al. and Shannon et al. by author + keywords.

4. **Grey literature (UNAIDS reports, government surveys, World Bank documents) won't have DOIs — write manual BibTeX entries.** Use `@techreport` for institutional reports (UNAIDS, NAIIS, IBBSS), `@incollection` for book chapters (Mboup et al.), and `@misc` for web resources (IHME GBD Results Tool, UNAIDS country pages). Include the URL in the entry. These are common in HIV/global health writing. You may need to do some web searching to find the relevant information for your manual BibTeX entry.
