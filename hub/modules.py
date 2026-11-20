"""modules.py - the hub's modules. Each one runs one skill.

Stage A: the Claude step is done by Ahmed in Claude Code, and its output (the spec) is uploaded to the job.
Stage B: `claude_step` becomes an API call, logged in the cost log.

Step modules (runner='step') are the PTE Phase 1 steps that are scripts only (toolkit/pte/run_step.py).
They are chained through the project's files: a job takes project files of the kinds in `inputs`
and puts its results back as project files of kind `output_kind`, which the next step picks up.
`inputs` is a list of groups; from each group the first kind the project has is ticked by default.
"""
MODULES = {
    'areas': dict(
        name='00 Area Calculations',
        desc='Drawings to AAA BUA vs IPMS 1: collected workbook, 3D model and offline client copy per asset, '
             'and the Annex D PowerPoint.',
        inputs='Plans for every level, a section, the roof plan; the report .docx for the BUA check.',
        claude_step='Read the drawings in Claude Code with the 00 Areas Calculation skill and save one spec per asset (.py or .json).',
        spec_types=('.py', '.json'),
        ready=True),
    'boq-unlock': dict(
        name='01 Unlock BOQs', runner='step', ready=True,
        desc='Phase 1 step 1: editable copies of the designer BOQs with sheet and workbook protection removed. '
             'Everything else stays byte-for-byte the same.',
        inputs=[['boq']], output_kind='boq-unlocked',
        what='Removes the protection tags only (no password needed), then checks every copy against its original: '
             'values, styles, merges and images. Files with no protection are copied as they are. '
             'A file with a password to open cannot be unlocked; legacy .xls is skipped.'),
    'asset-names': dict(
        name='02 Asset / Package Register', runner='step', review=True, ready=True,
        desc='Phase 1 step 2: File_Names.xlsx listing the assets and the packages (BOQs) of each, with a matrix, '
             'plus File_Names.txt.',
        inputs=[['boq-unlocked', 'boq'], ['register']], output_kind='register',
        what='Suggests an asset and a package for each BOQ from its file name (and sheet name), then you correct them '
             'on the review screen before the register is built. Tick a previous register to add to it: '
             'its rows keep their order and new ones are added.'),
    'divisions': dict(
        name='03 Stack Divisions into One Sheet', runner='step', ready=True,
        desc='Phase 1 step 3: for BOQs split into one tab per division, one ALL DIVISIONS sheet with formulas, '
             'page totals and print setup kept.',
        inputs=[['boq-unlocked', 'boq']], output_kind='boq-divisions',
        what='Exact copy, division detail tabs only (the DIVISION n SUMMARY tabs are left out). Every cell is checked '
             'against the source after Excel recalculates. BOQs that are already one sheet are reported and skipped.'),
    'name-assets': dict(
        name='04 Naming Assets in Master Sheet', runner='step', review=True, ready=True,
        desc='Phase 1 step 4: the PTE & Tender Evaluation master with one tab per asset (one block of divisions; '
             'BOQ codes go in column B at loading), named everywhere (Schedule of Assets, Asset / Project / '
             'Elemental Summary).',
        inputs=[['master'], ['register']], output_kind='master-named',
        review_step='Confirm the asset names, the spare tabs and the elemental amount.',
        what='Reads the master and the register, then you confirm the asset names, what happens to spare Asset tabs '
             'and which AAA amount the Elemental Summary reads. Excel does the build (tabs renamed, summary links); '
             'every link is then checked, and test amounts are planted in a '
             'throwaway copy to prove each total reaches the summaries. Supports the master that ships ready-made '
             'Asset #01..#25 tabs.'),
    'load-boq': dict(
        name='05 Load BOQs into Master Sheet', runner='step', review=True, ready=True,
        desc='Phase 1 step 5: every BOQ line, as written, into its asset tab of the named master, with the BOQ code '
             'in column B and amount and weight formulas on the priced lines.',
        inputs=[['master-named'], ['register'], ['boq-unlocked', 'boq']], output_kind='master-loaded',
        review_step='Check which tab and code each BOQ goes to, and the division each BOQ division maps to.',
        what='Reads each BOQ line by line, wording as written (page totals and repeated headings dropped), finds its '
             "asset and code in the register, and maps its divisions to the master's (old 16-division numbers "
             'included). Each loaded division gets 1 empty row under its heading and 3 empty rows before the next '
             'division. Excel inserts the rows, so every subtotal and the Asset Summary links follow. Every priced '
             'line is then read back against its BOQ, and in a throwaway copy all quantities are priced at 1 to prove '
             'each asset total. Column A (NRM), the QTY / Rate inputs and columns BC-BP are not touched.'),
    'link-rates': dict(
        name='06 Link Repeated Items to One Rate', runner='step', ready=True,
        desc='Phase 1 step 5b: the AAA Rate of an item that repeats an earlier item (higher up in its own asset, or '
             'in an earlier asset) is linked to it (BLUE certain, ORANGE check), so each rate is entered once.',
        inputs=[['master-loaded']], output_kind='master-linked',
        what='Assets in their master order, items top to bottom. Each item is compared first with the items above '
             'it in its own asset, then with earlier assets one at a time; the first match wins. BOQ-code headings '
             '("Arch - A2-1 - ...") are not counted as headings. BLUE: same division, heading, description and unit. ORANGE: same '
             'description and unit under another heading, or the same wording apart from spacing, punctuation and '
             'filler words with the same numbers. Openings (doors, windows) are never BLUE: the same code can be a '
             'different size. Only the AAA Rate cell is linked; quantities stay each asset\'s own. A Linking report '
             'lists every link, oranges first.'),
    'pricing-review': dict(name='Pricing review', desc='Phase 1 step 7: missing rates, outliers, mismatches, qty × rate checks.', ready=False),
    'rate-push': dict(name='06-08 Push rates into the BOQs', desc='Phase 1 step 8.', ready=False),
}

# project file kinds, in the order the project page lists them (label shown on upload and in lists)
FILE_KINDS = {
    'drawing': 'Drawing', 'report': 'Report (.docx)', 'boq': 'BOQ (as received)', 'boq-unlocked': 'BOQ (unlocked)',
    'register': 'Asset / package register', 'boq-divisions': 'BOQ (all divisions)',
    'master': 'PTE master (template)', 'master-named': 'PTE master (assets named)',
    'master-loaded': 'PTE master (BOQs loaded)', 'master-linked': 'PTE master (rates linked)',
    'workbook': 'Workbook', 'other': 'Other',
}
UPLOAD_KINDS = ('drawing', 'report', 'boq', 'register', 'master', 'workbook', 'other')

# A .py spec is executed by the pipeline, so only admins (Ahmed) may upload one (app.py py_ok); colleagues
# on the office network upload JSON specs. Set False to refuse .py from everyone (Stage C).
ALLOW_PY_SPECS = True
