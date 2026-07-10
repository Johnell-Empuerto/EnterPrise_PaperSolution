import psycopg2, zipfile, re, zlib, os, io

conn = psycopg2.connect(host="127.0.0.1", port=5432, dbname="irepodb", user="postgres", password="cimtops")
cur = conn.cursor()

f = open("q5_q6_analysis.txt", "w", encoding="utf-8")

def p(s=""):
    print(s)
    f.write(s + "\n")

p("=" * 80)
p("Q5: Attempt to Disprove the PDF Post-Processing Hypothesis")
p("=" * 80)

p("\nThe PDF hypothesis states: Legacy coordinates were computed by rendering the Excel")
p("worksheet to PDF, then extracting vector positions (grid lines, cell boundaries)")
p("from the PDF content stream, and normalizing them to 0-1 page ratios.")

p("\n--- Test 1: Evidence of PDF coordinate extraction code ---")
p("Result from Q3: NO PDF parsing/vector extraction code exists in the codebase.")
p("The only PDF library (PDFtoImage v5.1.0) is used solely for PDF-to-PNG rasterization.")
p("CONCLUSION: No current evidence of PDF coordinate extraction code.")
p("    -- Does NOT disprove the hypothesis; the legacy system may have had separate code --")

p("\n--- Test 2: Database cache of PDF-extracted coordinates ---")
p("Result from Q1: NO database table stores PDF geometry (bounding boxes, clip regions,")
p("grid line positions, cell boundaries extracted from PDF).")
p("The only position data is in def_cluster (text ratios already normalized).")
p("No 'cache', 'temp', 'geometry', 'vector', 'clip', or 'bbox' columns exist.")
p("CONCLUSION: No evidence of cached PDF measurements.")
p("    -- Does NOT disprove; legacy code may have extracted without caching --")

p("\n--- Test 3: Background PDF exists BEFORE ratio computation ---")
p("Check: Is there a timestamp showing the PDF was generated before ratios were stored?")
cur.execute("SELECT COUNT(*) FROM def_top WHERE background_image_file IS NOT NULL")
has_bg = cur.fetchone()[0]
cur.execute("SELECT COUNT(*) FROM def_top WHERE background_image_file IS NULL AND def_file IS NOT NULL")
has_excel_no_bg = cur.fetchone()[0]
p(f"Forms with background_image_file: {has_bg}")
p(f"Forms with def_file (Excel) but NO background_image_file: {has_excel_no_bg}")

# Check if background_image_file generation predates or postdates first def_cluster creation
p("\n--- Test 4: Timing analysis of oldest forms ---")
cur.execute("SELECT d.def_top_id, d.sys_regist_time, d.designer_version, "
            "CASE WHEN d.background_image_file IS NOT NULL THEN 1 ELSE 0 END as has_bg, "
            "CASE WHEN dc.def_top_id IS NOT NULL THEN 1 ELSE 0 END as has_clusters "
            "FROM def_top d LEFT JOIN def_cluster dc ON d.def_top_id = dc.def_top_id "
            "WHERE d.designer_version IS NOT NULL "
            "GROUP BY d.def_top_id, d.sys_regist_time, d.designer_version, d.background_image_file "
            "ORDER BY d.def_top_id LIMIT 20")
for r in cur.fetchall():
    p(f"  id={r[0]} ver={r[1]} regist={str(r[2])[:19]} bg={'Y' if r[3] else 'N'} clusters={'Y' if r[4] else 'N'}")

p("\n--- Test 5: Check if ALL forms with clusters have background PDFs ---")
cur.execute("""
    SELECT COUNT(DISTINCT dc.def_top_id) as forms_with_clusters,
           SUM(CASE WHEN d.background_image_file IS NOT NULL THEN 1 ELSE 0 END) as with_bg
    FROM def_cluster dc
    JOIN def_top d ON d.def_top_id = dc.def_top_id
    WHERE d.background_image_file IS NOT NULL
""")
r = cur.fetchone()
p(f"Forms with clusters AND background PDF: {r[1]} out of {r[0]}")

# Forms with clusters but no background PDF
cur.execute("""
    SELECT DISTINCT dc.def_top_id, d.designer_version 
    FROM def_cluster dc 
    JOIN def_top d ON d.def_top_id = dc.def_top_id 
    WHERE d.background_image_file IS NULL
    ORDER BY dc.def_top_id
""")
no_bg = cur.fetchall()
p(f"\nForms with clusters but NO background PDF: {len(no_bg)}")
for r in no_bg[:10]:
    p(f"  id={r[0]} ver={r[1]}")

p("\n" + "=" * 80)
p("Q6: Evidence Classification")
p("=" * 80)

p("""
EVIDENCE CATEGORIES:
A) Evidence supporting Font Metrics / COM-based theory
B) Evidence supporting PDF Post-processing theory
C) Evidence contradicting each theory
D) Ambiguous evidence
""")

p("--- A) Supporting Font Metrics / COM-based theory ---")
p("""
A1. Current engine uses Excel COM exclusively for coordinate extraction.
    - File: ExcelCaptureService.cs ExtractFields()
    - Coordinates derived from cell.Left, cell.Top, cell.Width, cell.Height
    - PDF is used only as intermediate for rasterization, not coordinate source

A2. def_cluster ratios stored as character varying (text) - consistent with
    COM-based division result, not PDF coordinate extraction.

A3. Legacy and current PDFs show grid lines at similar positions (within 0.15pt),
    suggesting the same Excel layout engine drives both. If a separate coordinate
    extraction were used, it would be redundant.

A4. pic_original_resolution field stores image DPI info, consistent with a
    rasterization pipeline (Excel->PDF->PNG), not a vector extraction pipeline.

A5. No 'pdf_fontsize_deduction' field exists in def_top_option but not in 
    def_cluster - the font deduction is a rendering parameter, not coordinate.
""")

p("--- B) Supporting PDF Post-processing theory ---")
p("""
B1. FORM 546 CROSS-SHEET PARADOX (STRONGEST EVIDENCE):
    - Clusters stored under def_sheet_no=1 which maps to '_Fields' sheet XML
    - _Fields sheet has center_h='0' (FALSE) in xl/charts.xml
    - BUT stored min origin = 205.92pt requires centering (155pt from 50.4pt margin)
    - With center_h=FALSE, origin should equal margin (50.4pt)
    - This contradiction is EXPLAINED by PDF rendering: ExportAsFixedFormat renders
      from the default/active Sheet1 (which HAS centering), ignoring the DB cluster mapping.
    - Conclusion: COM-based formulas using the cluster's sheet parameters WOULD fail.
      Only a PDF-based approach (rendering Sheet1 regardless) would get the right origin.

B2. FORM 228 OUTLIER:
    - 5 default columns, Calibri 11pt, centering ON
    - Stored left_position ratio = 0.283743 (origin = 173.65pt)
    - Expected (Calibri formula) ≈ 200.15pt, residual = 26.5pt ERROR
    - Expected (COM with current Aptos) ≈ 210pt, residual = 36.35pt ERROR
    - This massive discrepancy cannot be explained by font metrics.
    - BUT could be explained if the PDF was rendered at a specific scale or DPI
      and the clip region inside the PDF captured a different area.

B3. LEGACY PDF PRODUCER DIFFERENCE:
    - Old forms: 'Microsoft Print to PDF' driver
    - Current forms: ExportAsFixedFormat (xlTypePDF)
    - These render through different engines with potentially different internal
      measurement approaches.

B4. background_image_file EXISTS as bytea for most forms - the rendered PDF
    was available at the time of initial form creation. A PDF-based approach
    could have read coordinates from this exact same file.

B5. ratios are 0-1 normalizations of page dimensions - this is consistent with
    extracting pixel positions from a rendered image and dividing by total page size.
""")

p("--- C) Evidence Contradicting Each Theory ---")
p("""
C1. CONTRADICTING FONT METRICS THEORY:
    a) Form 228: 26.5pt error with Calibri 11pt, 5 default columns, centering
       - If font metrics theory were correct, Calibri forms that never changed
         font should have <1pt error. Instead, 26.5pt error.
    b) Form 546 cross-sheet paradox: COM formulas using sheet-specific settings
       cannot produce correct origins for clusters mapped to non-centered sheets.
    c) 12 candidate formulas attempted, none explain all 29 test forms (from 
       comprehensive report).
    d) Font on paper is Calibri (legacy) vs Aptos (current), but the stored
       ratios don't consistently match either font's column width predictions.

C2. CONTRADICTING PDF POST-PROCESSING THEORY:
    a) NO PDF parsing code exists in the codebase (anywhere).
    b) NO database cache of PDF-extracted vector data.
    c) NO historical or legacy code for PDF coordinate extraction found.
    d) If PDF was used, why store ratios as text (not float/numeric)?
       Text ratios suggest string formatting during COM computation.
    e) The cross-sheet paradox has only been verified for ONE form (546).
       Other forms need verification.
    f) Form 546's background PDF in the DB was generated by modern Excel 
       (CreatorTool='Microsoft Excel for Microsoft 365'), not the legacy driver.
       So the same pipeline exists today but produces different ratios.
""")

p("--- D) Ambiguous Evidence ---")
p("""
D1. Grid lines in form 546 PDF: 204.83pt vertical, 405.49pt vertical
    - These are 200.66pt apart (= 612pt * 0.3364706 legacy ratio - 50.4pt margin - 4.7pt?)
    - 200.66pt ≈ 4 columns × 50.165pt
    - BUT current COM Range.Width = 192.00pt (4 columns × 48.00pt)
    - AND PDF-drawn grid = 200.66pt
    - Both theories predict similar values (~50pt per column), so cannot discriminate

D2. Legacy column width = 200.09pt (backward-solved from 4 forms)
    - Calibri 11pt expected ≈ 200.15pt
    - PDF grid measured = 200.66pt
    - All three cluster within 0.57pt - both theories fit equally well here

D3. designer_version timeline:
    - v7.2.13950: oldest (2020), likely Calibri era
    - v8.2.25110: Nov 2025-Jan 2026, still Calibri in stored xlsx 
    - v8.2.26020: Feb 2026+, may have transitioned to Aptos
    - The ratio formulas seem consistent across versions, suggesting a 
      version-independent algorithm (PDF rendering is version-independent)
""")

p("\n" + "=" * 80)
p("FINAL PROBABILITY ASSESSMENT")
p("=" * 80)
p("""
Based on all collected evidence:

PDF Post-processing theory: 55%
  - Explains form 546 cross-sheet paradox
  - Explains form 228 outlier
  - Timing-independent
  - BUT: no evidence of PDF parsing code anywhere

Font Metrics / COM-based theory: 35%
  - Current engine uses COM
  - Ratios stored as text
  - BUT: cannot explain cross-sheet paradox or form 228

Third unknown mechanism: 10%
  - Neither theory fits all evidence perfectly
  - Some other approach entirely (e.g., hardcoded position tables,
    image pixel analysis, or manual calibration)
""")

p("\n--- Recommended decisive test ---")
p("""
1. For form 546: Generate the EXCEL file from the stored def_file, open in Excel 2016
   (Calibri 11pt), measure Range.Width. If width = 200.66pt (matches PDF grid), the
   font metrics theory gains support. If width = 192.00pt (matches current Aptos),
   then the stored ratios were NOT derived from current COM.

2. Examine more forms with the cross-sheet paradox pattern: find other forms where
   def_cluster.def_sheet_no maps to a sheet without centering but origin differs
   from margin. This would confirm or refute the PDF hypothesis.

3. Search the codebase (or git history) for any version that used PDF coordinate
   extraction -- check git log for removed features or deprecated code paths.
""")

conn.close()
f.close()
print("Q5/Q6 saved to q5_q6_analysis.txt")
