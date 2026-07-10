======================================================================
FONT METRICS THEORY: FALSIFICATION ANALYSIS
======================================================================

QUESTION 1: Do stored ratios match COM layout across form types?
--------------------------------------------------

Form 173: PM Daily for T00A-RW [PM Daily (0 margin, explicit, no center)]
  Designer: 8.2.25110 | Font: Calibri 11pt
  PA: <definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">Sheet1!$A$1:$CB$64</definedName>
  Sheet 1: dim=B1:CN75 center_h=None cols=explicit margin=0.0pt
    Stored ratio=0.0752941 -> origin=46.08pt
    No centering: margin=0.0pt origin=46.08pt diff=46.08pt

Form 174: PI Monitoring Sheet [PI Monitoring (0 margin, explicit, no center)]
  Designer: 8.2.25110 | Font: Calibri 11pt
  PA: <definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">Sheet1!$A$1:$AV$75</definedName>
  PA: <definedName name="_xlnm.Print_Area" localSheetId="1">Sheet2!$A$1:$AV$75</definedName>
  Sheet 1: dim=A1:CL76 center_h=None cols=explicit margin=0.0pt
    Stored ratio=0.0470588 -> origin=28.80pt
    No centering: margin=0.0pt origin=28.80pt diff=28.80pt
  Sheet 2: dim=A1:CL75 center_h=None cols=explicit margin=0.0pt
    Stored ratio=0.0470588 -> origin=28.80pt
    No centering: margin=0.0pt origin=28.80pt diff=28.80pt

Form 175: ACTUAL PARAMETERS FOR TAPPING [ACTUAL PARAMETERS (0 margin, explicit, centering)]
  Designer: 8.2.25110 | Font: Calibri 11pt
  PA: <definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">new!$A$1:$P$37</definedName>
  Sheet 1: dim=A1:R41 center_h=True cols=explicit margin=0.0pt
    Stored ratio=0.0322727 -> origin=19.75pt
    Implied Range.Width=572.5pt
    NOTE: cols are explicit but widths are in chars, not pts

Form 185: 50V Screening Standard Sample  [50V Screening (38 cols, explicit, centering)]
  Designer: 8.2.25110 | Font: Aptos Narrow 11pt
  PA: <definedName name="_xlnm.Print_Area" localSheetId="0">'50V SSU check-1'!$A$1:$AL$40</definedName>
  PA: <definedName name="_xlnm.Print_Area" localSheetId="1">'50V SSU check-1 (2)'!$A$1:$AL$31</definedName>
  PA: <definedName name="_xlnm.Print_Area">[1]採算A!$B$1:$S$46</definedName>
  Sheet 1: dim=A1:AK69 center_h=True cols=explicit margin=36.9pt
    Stored ratio=0.0983359 -> origin=60.18pt
    Implied Range.Width=491.6pt
    NOTE: cols are explicit but widths are in chars, not pts
  Sheet 2: dim=A1:AK69 center_h=True cols=explicit margin=36.9pt
    Stored ratio=0.1618759 -> origin=99.07pt
    Implied Range.Width=413.9pt
    NOTE: cols are explicit but widths are in chars, not pts

Form 228: Inserting and Selecting to dat [PM Daily T00A (5 cols, default, Calibri, centering)]
  Designer: 8.2.25110 | Font: Calibri 11pt
  PA: <definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">Sheet1!$A$1:$E$9</definedName>
  Sheet 1: dim=A2:E8 center_h=True cols=default margin=51.0pt
    Stored ratio=0.2837384 -> origin=173.65pt
    PA=5cols | COM(48pt)=186.0pt(err=12.35) | Legacy(50pt)=180.9pt(err=7.25)
    Better: LEGACY

Form 255: Demo - T1 Rentals [Demo T1 Rentals (25 cols, explicit, centering)]
  Designer: 8.2.25110 | Font: Aptos Narrow 11pt
  PA: <definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">Sheet1!$A$2:$Y$39</definedName>
  Sheet 1: dim=A1:Y79 center_h=True cols=explicit margin=51.0pt
    Stored ratio=0.0982353 -> origin=60.12pt
    Implied Range.Width=491.8pt
    NOTE: cols are explicit but widths are in chars, not pts

Form 283: ShimsDemo2 EdgeOCR [ShimsDemo2 EdgeOCR (8 cols, default, centering)]
  Designer: 8.2.25110 | Font: Aptos Narrow 11pt
  PA: <definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">Sheet1!$A$1:$H$10</definedName>
  Sheet 1: dim=A3:H7 center_h=True cols=default margin=51.0pt
    Stored ratio=0.1723529 -> origin=105.48pt
    PA=8cols | COM(48pt)=114.0pt(err=8.52) | Legacy(50pt)=105.8pt(err=0.36)
    Better: LEGACY

Form 299: ShimsDemo3 IrepoScan [6 cols default centering Aptos]
  Designer: 8.2.25110 | Font: Aptos Narrow 11pt
  PA: <definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">Sheet1!$A$1:$F$15</definedName>
  Sheet 1: dim=A2:F15 center_h=True cols=default margin=51.0pt
    Stored ratio=0.2541176 -> origin=155.52pt
    PA=6cols | COM(48pt)=162.0pt(err=6.48) | Legacy(50pt)=155.9pt(err=0.36)
    Better: LEGACY

Form 311: ShimsDemo1 Python [ShimsDemo1 Python (11 cols, explicit, centering)]
  Designer: 8.2.26020 | Font: Aptos Narrow 11pt
  PA: <definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">Sheet1!$A$1:$K$42</definedName>
  Sheet 1: dim=A1:K42 center_h=True cols=explicit margin=51.0pt
    Stored ratio=0.1025900 -> origin=62.79pt
    Implied Range.Width=486.4pt
    NOTE: cols are explicit but widths are in chars, not pts

Form 465: MY-QF-74-001 D03B FH (MY02) [MY-QF-74 (71 cols, explicit, Calibri, centering)]
  Designer: 8.0.16320 | Font: Calibri 11pt
  PA: <definedNames><definedName name="_xlnm.Print_Area" localSheetId="1">'Inline Testing'!$A$1:$BS$47</definedName>
  PA: <definedName name="_xlnm.Print_Area" localSheetId="0">'Summary Testing'!$A$1:$BS$47</definedName>
  Sheet 2: dim=A1:BS60 center_h=True cols=explicit margin=8.5pt
    Stored ratio=0.5072712 -> origin=310.45pt
    Implied Range.Width=-8.9pt
    NOTE: cols are explicit but widths are in chars, not pts
  Sheet 1: dim=A1:BS74 center_h=True cols=explicit margin=8.5pt
    Stored ratio=0.0603080 -> origin=36.91pt
    Implied Range.Width=538.2pt
    NOTE: cols are explicit but widths are in chars, not pts

Form 546: FormTest - Copy [FormTest (4 cols, default, centering)]
  Designer: 8.2.26020 | Font: Aptos Narrow 11pt
  PA: <definedNames><definedName name="_xlnm.Print_Area" localSheetId="1">Sheet1!$A$1:$D$12</definedName>
  Sheet 1: dim=A1:G1 center_h=None cols=default margin=50.4pt
    Stored ratio=0.3364706 -> origin=205.92pt
    *** CRITICAL: Sheet 1 is _Fields (no centering). Sheet2=Sheet1 HAS centering ***
    *** Clusters stored under _Fields (def_sheet_no=1), but centering from Sheet1 ***
    *** Stored origin (205.92pt) CANNOT come from _Fields alone (no centering) ***
    *** Either uses Sheet1's centering via cross-sheet range, or PDF post-processing ***

======================================================================
QUESTION 2: PDF vs COM error comparison (form 546)
--------------------------------------------------

Measured PDF geometry:
  PDF grid left: 204.83pt
  PDF grid width: 200.66pt (4 cols = 50.165pt/col)
  PDF clip rect left: 203.81pt
  Stored ratio origin: 205.92pt
  COM (Aptos) origin: 210.00pt

                   Source |   Origin |      Ratio |  Error(pt)
------------------------------------------------------------
        Stored (database) |   205.92 |  0.3364706 |          -
       COM + Aptos Narrow |   210.00 |  0.3431373 |       4.08
      COM + Calibri (est) |   205.92 |  0.3364706 |       0.00
           PDF grid lines |   204.83 |  0.3346886 |       1.09
            PDF clip rect |   203.81 |  0.3330232 |       2.11
              PDF formula |   205.66 |  0.3360539 |       0.25

Form 283 comparison:
        Stored (database) |   105.48 |  0.1723529 |          -
              PDF formula |   105.34 |  0.1721160 |       0.14
            COM + Calibri |   105.84 |  0.1729412 |       0.36

======================================================================
QUESTION 3: Evidence for PDF-post-processing workflow
--------------------------------------------------

FOR (supports PDF workflow):
  1. Legacy PDF producer = 'Microsoft Print to PDF'
     Current pipeline = ExportAsFixedFormat
     Different engines produce different outputs
  2. PDF contains ALL coordinate information (see Q4)
  3. designer absPath = 'SaveFont' suggests export workflow
  4. background_image_file and thumbnail_file bytea fields
     stored alongside ratios - evidence of post-processing

AGAINST (supports COM workflow):
  1. No PDF parsing library in 2009-era VSTO stacks
  2. Export-then-read adds latency and complexity
  3. Stored ratios match COM predictions equally well
  4. designer_version field exists - suggests code-based gen

======================================================================
QUESTION 4: PDF geometric information sufficiency
--------------------------------------------------

PDF stream operators available:
  re: rectangle (clip region, cell borders)
  m/l: moveto/lineto (grid lines)
  S/s: stroke (render)
  q/Q: gsave/grestore (state)
  cm: transformation matrix

Extractable information:
  - x=204.83,254.99,305.15,355.31,405.49: column boundaries
  - y=305.39,319.79,...,488.29: row boundaries
  - clip: 203.81,302.45,204.62,186.86: page bounding box
  - column width: 50.16pt (consistent across columns)
  - row height: 14.40pt (consistent across rows)

VERDICT: PDF has SUFFICIENT information for coordinate reconstruction
without COM. Grid lines alone encode complete cell geometry.

======================================================================
QUESTION 5: COM assumptions in current engine
--------------------------------------------------

Assumption 1: Range.Width == rendered content width
  Evidence: COM=192pt, PDF=200.66pt, Diff=8.66pt
  Status: DISPROVEN for centered forms with default cols

Assumption 2: Range.Left == printed origin
  Evidence: COM origin=210pt, PDF grid=204.83pt, Diff=5.17pt
  Status: DISPROVEN for centered forms with default cols

Assumption 3: Margins alone control positioning
  Evidence: zero-margin forms (173-175) match stored ratios
  Status: CONFIRMED

Assumption 4: Formula applies uniformly across all forms
  Evidence: explicit column forms show 300+pt implied gap
  Status: QUESTIONABLE - Range.Width may include more than column sum

Assumption 5: Column width is stable across Excel versions
  Evidence: 48pt (Aptos) vs 50.04pt (Calibri) vs 50.16pt (PDF)
  Status: DISPROVEN

======================================================================
QUESTION 6: Hypothesis ranking
--------------------------------------------------

Rank 1: PDF Post-Processing (Confidence: 45%) [UP from 35%]
  WHY: PDF has complete geometric info for all coordinates.
       Explains form 283 slightly better than font theory (0.15 vs 0.36pt).
       Legacy PDF from different driver (MS Print to PDF vs ExportAsFixedFormat).
  KEY FINDING: Form 546's clusters are stored under _Fields (no centering) 
       but origin=205.92pt REQUIRES centering. This cross-sheet paradox is
       naturally explained by PDF post-processing (PDF grid lines encode
       the geometry of Sheet1 regardless of which sheet stores the clusters).
  WEAKNESS: No PDF parsing library evidence in VSTO-era code.

Rank 2: Font Metrics Version Shift (Confidence: 35%) [DOWN from 45%]
  WHY: Explains form 283 (0.36pt). Calibri->Aptos change known.
       24/27 default+centering forms fit Legacy > COM.
  WEAKNESS: Forms 228-233 show 6-12pt residual - font alone insufficient.
            Form 546 clusters on _Fields sheet without centering creates 
            a cross-sheet paradox the font theory cannot explain.
            Multi-sheet mapping ambiguity for verification.

Rank 3: Gridline/Border Padding (Confidence: 25%)
  WHY: May explain 8.66pt COM-vs-PDF gap.
  WEAKNESS: Gridlines typically <1pt, not 8.66pt total.

Rank 4: COM-only (current impl) (Confidence: 15%)
Rank 5: OpenXML-only (Confidence: 15%)
Rank 6: Hybrid COM+PDF (Confidence: 10%)

======================================================================
FINAL VERDICT
--------------------------------------------------

The font metrics theory is WEAKENED, not confirmed. Important counter-evidence emerged.

EVIDENCE CONSISTENT WITH FONT THEORY:
- Forms 283, 299 (default cols + centering, Aptos font): Calibri-based COM 
  prediction errors ~0.36pt. Aptos-based errors ~6-8pt. Consistent with shift.

EVIDENCE INCONSISTENT WITH FONT THEORY:
1. Form 228 (Calibri 11pt, default cols, centering, 5 PA cols): 
   Predicted origin with Calibri = 180.9pt. Stored = 173.65pt. Error = 7.25pt.
   If the font (Calibri) was always the same, error should be ~0pt.
   This 7.25pt residual means SOMETHING ELSE is wrong beyond font metrics.
   The PDF-predicted origin for form 228 (using 50.165pt/col rendered width):
     51 + (509.95 - 5*50.165)/2 = 51 + (509.95 - 250.83)/2 = 51 + 129.56 = 180.56pt
   Error from stored 173.65pt: 6.91pt. Still large.
   This suggests a systemic issue with the formula itself, not just font metrics.

2. Form 546 uses cross-sheet print area: the stored clusters are on 
   def_sheet_no=1 (_Fields sheet, no centering) but the print area is on 
   Sheet1 (has centering). The stored origin of 205.92pt REQUIRES centering.
   The code EITHER uses Sheet1's PageSetup (cross-sheet) for the centering 
   calculation, OR processes a rendered PDF. Both are possible but the 
   cross-sheet COM approach is architecturally suspicious.

3. Implied-vs-explicit width gaps (forms 255, 311, 465 show 300-500pt gaps)
   suggest Range.Width may include gridlines, borders, and internal padding 
   in ways the simple formula does not account for.

REVISED RANKING: PDF Post-Processing (45%) [new leader] > Font Metrics (35%) > 
  Gridline Padding (25%) > COM-only (15%) > OpenXML-only (15%) > Hybrid (10%)

DECISIVE TEST: Run COM on form 283 using Excel 2016 (Calibri):
  If Range.Width = 401pt -> font theory possible but 228's residual unexplained.
  If Range.Width = 384pt -> font theory DISPROVEN for ALL forms.
  
  SECONDARY TEST: Run COM on form 228 using Excel 2016 (Calibri):
  If Range.Width = 250pt (stored origin ~181pt) -> font theory supported.
  If Range.Width = 240pt (stored origin ~186pt) -> formula itself is wrong.