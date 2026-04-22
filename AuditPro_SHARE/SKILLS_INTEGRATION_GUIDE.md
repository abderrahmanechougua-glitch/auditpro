"""
SKILLS INTEGRATION GUIDE - AuditPro

Complete guide on how to use Claude Code skills in AuditPro modules.

==============================================================================
OVERVIEW
==============================================================================

AuditPro now integrates 4 major Claude Code skills:
- PDF Skill: Extract tables, OCR, form fields, split PDFs
- Excel Analysis Skill: Validate, detect formats, normalize, analyze
- Data Visualization Skill: Create charts and statistical analysis
- DOCX Skill: Create Word documents, mail merge, add tables/images

Each module can use these skills to enhance its capabilities.

==============================================================================
QUICK START - Using Skills in Your Module
==============================================================================

1. Import the skills bridge in your module:

    from agent.skills_bridge import get_skills_bridge
    
    bridge = get_skills_bridge()

2. Use the bridge to access skills:

    # Example: Extract tables from TVA PDF
    result = bridge.extract_pdf_tables('path/to/tva.pdf')
    
    # Example: Validate GL data
    import pandas as pd
    df = pd.read_excel('path/to/gl.xlsx')
    validation = bridge.validate_excel_data(df)
    
    # Example: Create a chart
    chart_result = bridge.create_audit_chart(
        chart_type='bar',
        data=revenue_data,
        title='TVA par taux',
        output_path='chart.png'
    )

==============================================================================
SKILL-BY-SKILL GUIDE
==============================================================================

─────────────────────────────────────────────────────────────────────────────
PDF SKILL - Extract data from PDF files
─────────────────────────────────────────────────────────────────────────────

**Use for:** TVA, CNSS, Factures, IR extraction modules

A. Extract Tables (TVA declarations, CNSS bordereaux)

    from agent.skills_bridge import get_skills_bridge
    
    bridge = get_skills_bridge()
    
    result = bridge.extract_pdf_tables(
        pdf_path='declarations_tva.pdf',
        page_numbers=[0, 1],  # Pages 1-2
        confidence_threshold=0.7
    )
    
    # Result contains:
    # - tables: List of extracted tables with data
    # - confidence: Overall confidence score
    # - pages_processed: Number of pages analyzed
    
    tables = result['tables']
    for table in tables:
        print(f"Page {table['page']}: {table['rows']} rows × {table['cols']} cols")


B. Extract Form Fields (CNSS, IR forms)

    result = bridge.extract_pdf_forms(
        pdf_path='cnss_bordereau.pdf',
        form_fields=['mass_salariale', 'af_amount', 'amo_amount']
    )
    
    # Result contains extracted form values
    fields = result['fields']
    print(f"Mass salariale: {fields.get('mass_salariale', {}).get('value')}")


C. OCR Extraction with Confidence Scoring

    result = bridge.ocr_extract_from_pdf(
        pdf_path='scanned_invoice.pdf',
        language='fra',  # French
        confidence_threshold=0.7
    )
    
    # Result contains:
    # - text: Extracted text with page breaks
    # - confidence: Overall confidence (0-1)
    # - pages: List with per-page text and confidence


─────────────────────────────────────────────────────────────────────────────
EXCEL ANALYSIS SKILL - Analyze and validate data
─────────────────────────────────────────────────────────────────────────────

**Use for:** Lettrage GL, Retraitement, Invoice analysis modules

A. Validate Data Against Rules

    import pandas as pd
    
    df = pd.read_excel('grand_livre.xlsx')
    
    rules = {
        'balance_check': {
            'type': 'formula',
            'formula': '(df["Debit"].sum() - df["Credit"].sum() < 0.01)'
        }
    }
    
    result = bridge.validate_excel_data(df, rules=rules)
    
    # Result contains:
    # - valid: Boolean
    # - errors: List of validation errors
    # - warnings: List of warnings
    # - stats: Data quality statistics


B. Detect Accounting File Format

    result = bridge.detect_accounting_format('fichier_comptable.xlsx')
    
    # Result contains:
    # - detected_format: "Sage 100", "Dynamics 365", etc.
    # - confidence: 0.7-0.95
    # - characteristics: List of detected features
    
    if result['detected_format'] == 'Sage 100':
        print(f"Sage 100 GL detected with {result['confidence']} confidence")


C. Normalize GL from Various Formats

    result = bridge.normalize_general_ledger(
        file_path='gl_sage.xlsx',
        target_format='standard'
    )
    
    # Result contains:
    # - normalized_data: Standardized GL as dict
    # - column_mapping: Original → Standard column mapping
    # - rows_processed: Number of rows normalized
    
    # Now all GLs are in standard format: Account, Description, Debit, Credit, Balance


D. Detect Statistical Anomalies

    result = bridge.detect_data_anomalies(
        df=df,
        numeric_column='Amount',
        threshold=2.0  # Standard deviations
    )
    
    # Result contains:
    # - anomalies_found: Count of anomalies
    # - outliers: List of anomalous rows
    # - statistics: Mean, std, min, max, median


─────────────────────────────────────────────────────────────────────────────
DATA VISUALIZATION SKILL - Create charts and analysis
─────────────────────────────────────────────────────────────────────────────

**Use for:** SRM Generator, Reporting, Analysis presentation

A. Create Bar Chart (TVA by rate)

    tva_by_rate = {
        '20%': 1500000,
        '14%': 850000,
        '10%': 320000,
        '7%': 180000
    }
    
    result = bridge.create_audit_chart(
        chart_type='bar',
        data=tva_by_rate,
        title='CA par taux TVA',
        output_path='chart_tva.png'
    )


B. Create Line Chart (Trend analysis N vs N-1)

    import pandas as pd
    
    data = pd.DataFrame({
        'Month': ['Jan', 'Feb', 'Mar', 'Apr'],
        'Year_N': [100000, 120000, 130000, 125000],
        'Year_N_1': [95000, 105000, 115000, 118000]
    })
    
    result = bridge.create_audit_chart(
        chart_type='line',
        data=data,
        title='Evolution du CA',
        x_col='Month',
        y_cols=['Year_N', 'Year_N_1'],
        output_path='chart_evolution.png'
    )


C. Create Pie Chart (Expense distribution)

    expenses = {
        'Supplier A': 350000,
        'Supplier B': 280000,
        'Supplier C': 195000,
        'Other': 175000
    }
    
    result = bridge.create_audit_chart(
        chart_type='pie',
        data=expenses,
        title='Distribution dépenses',
        output_path='chart_expenses.png'
    )


D. Create Comparison Chart (N vs N-1)

    import pandas as pd
    
    year_n = pd.Series([500, 650, 720, 580])
    year_n_1 = pd.Series([450, 600, 700, 550])
    
    result = bridge.create_audit_chart(
        chart_type='comparison',
        data=[year_n, year_n_1],
        title='Comparison N vs N-1',
        output_path='comparison.png'
    )


─────────────────────────────────────────────────────────────────────────────
DOCX SKILL - Generate Word documents
─────────────────────────────────────────────────────────────────────────────

**Use for:** SRM Generator, Circularisation, Audit reports

A. Create SRM Document

    result = bridge.create_word_document(
        title='Summary of Recorded Misstatements (SRM)',
        output_path='SRM_2024.docx',
        content={
            'Audit Findings': [
                'Finding 1: Unsupported journal entries',
                'Finding 2: Accrual timing differences',
                'Finding 3: Inventory cutoff issues'
            ],
            'Management Response': 'All findings have been addressed.'
        }
    )


B. Add Table to Document

    data = [
        ['Finding', 'Amount', 'Status'],
        ['JE Support', '125,000', 'Resolved'],
        ['Accrual', '87,500', 'Pending'],
        ['Cutoff', '45,000', 'Under Review']
    ]
    
    bridge.docx.add_table(
        document_path='SRM_2024.docx',
        data=data,
        title='Summary of Misstatements'
    )


C. Add Chart to Document

    bridge.docx.add_image(
        document_path='SRM_2024.docx',
        image_path='chart_tva.png',
        width_inches=6.0,
        caption='TVA Analysis by Rate'
    )


D. Mail Merge - Generate Personalized Letters (Circularisation)

    tiers_data = [
        {'name': 'Client ACME', 'email': 'contact@acme.com', 'balance': '250,000 DH'},
        {'name': 'Client XYZ', 'email': 'info@xyz.com', 'balance': '180,000 DH'},
        {'name': 'Client 123', 'email': 'hello@123.com', 'balance': '95,000 DH'}
    ]
    
    result = bridge.generate_personalized_letters(
        template_path='template_circularisation.docx',
        merge_data=tiers_data,
        output_dir='output/letters',
        naming_field='name'
    )
    
    # Creates: output/letters/Client ACME_001.docx, etc.


==============================================================================
MODULE-SPECIFIC EXAMPLES
==============================================================================

─ MODULE: Centralisation TVA ─

    from modules.base_module import BaseModule, ModuleInput, ModuleResult
    from agent.skills_bridge import get_skills_bridge
    import pandas as pd
    
    class CentralisationTVA(BaseModule):
        
        def execute(self, inputs, output_dir, progress_callback=None):
            bridge = get_skills_bridge()
            pdf_path = inputs['fichier_tva']
            
            # Extract tables from TVA PDF
            progress_callback(10, "Extracting TVA tables...")
            extraction = bridge.extract_pdf_tables(pdf_path)
            
            if not extraction['success']:
                return ModuleResult(success=False, message="Extraction failed")
            
            # Create visualization
            progress_callback(50, "Creating charts...")
            chart_result = bridge.create_audit_chart(
                chart_type='bar',
                data={'20%': 1500000, '14%': 850000},
                title='CA par taux',
                output_path=f'{output_dir}/tva_chart.png'
            )
            
            return ModuleResult(
                success=True,
                output_path=f'{output_dir}/tva_output.xlsx',
                message="TVA extraction completed",
                stats={'tables_extracted': len(extraction['tables'])}
            )


─ MODULE: Lettrage GL ─

    class LettrageGL(BaseModule):
        
        def execute(self, inputs, output_dir, progress_callback=None):
            bridge = get_skills_bridge()
            gl_file = inputs['grand_livre']
            
            # Detect format and normalize if needed
            progress_callback(10, "Detecting GL format...")
            format_detection = bridge.detect_accounting_format(gl_file)
            
            # Validate GL data
            progress_callback(30, "Validating GL...")
            df = pd.read_excel(gl_file)
            validation = bridge.validate_excel_data(df)
            
            if not validation['valid']:
                return ModuleResult(
                    success=False,
                    errors=validation['errors']
                )
            
            # Detect anomalies
            progress_callback(50, "Detecting anomalies...")
            anomalies = bridge.detect_data_anomalies(df, 'Amount')
            
            # Create visualization
            progress_callback(80, "Creating reconciliation chart...")
            bridge.create_audit_chart(
                chart_type='pie',
                data={'Matched': 8500000, 'Partial': 350000, 'Unmatched': 125000},
                title='GL Reconciliation Status',
                output_path=f'{output_dir}/reconciliation.png'
            )
            
            return ModuleResult(
                success=True,
                stats={
                    'format': format_detection['detected_format'],
                    'anomalies': anomalies['anomalies_found']
                }
            )


─ MODULE: SRM Generator ─

    class SRMGenerator(BaseModule):
        
        def execute(self, inputs, output_dir, progress_callback=None):
            bridge = get_skills_bridge()
            srm_file = inputs['tableau_srm']
            
            # Read SRM data
            df = pd.read_excel(srm_file)
            
            # Create SRM document
            progress_callback(10, "Creating SRM document...")
            bridge.create_word_document(
                title='Summary of Recorded Misstatements',
                output_path=f'{output_dir}/SRM.docx'
            )
            
            # Add summary table
            progress_callback(30, "Adding summary table...")
            summary_data = df[['Finding', 'Amount', 'Status']].head(10)
            bridge.docx.add_table(
                document_path=f'{output_dir}/SRM.docx',
                data=summary_data,
                title='Audit Findings Summary'
            )
            
            # Create comparison chart N vs N-1
            progress_callback(50, "Creating comparison chart...")
            chart_result = bridge.create_audit_chart(
                chart_type='comparison',
                data=[df['Year_N'], df['Year_N_1']],
                title='Comparison N vs N-1',
                output_path=f'{output_dir}/comparison.png'
            )
            
            # Add chart to document
            progress_callback(70, "Embedding chart...")
            bridge.docx.add_image(
                document_path=f'{output_dir}/SRM.docx',
                image_path=f'{output_dir}/comparison.png',
                width_inches=6.0,
                caption='Financial Comparison'
            )
            
            return ModuleResult(
                success=True,
                output_path=f'{output_dir}/SRM.docx',
                message="SRM generated successfully"
            )


==============================================================================
CONFIGURATION & REQUIREMENTS
==============================================================================

Skills require these packages (already in requirements.txt):
- pdfplumber: PDF table extraction
- pandas: Excel analysis
- matplotlib, seaborn: Visualization
- python-docx: Word document generation
- pytesseract: OCR (optional, for scanned PDFs)
- pdf2image: PDF to image conversion (for OCR)
- pypdf: PDF manipulation

Install optional OCR support:
    pip install pytesseract pdf2image

Install Tesseract OCR on your system:
    Windows: https://github.com/UB-Mannheim/tesseract/wiki
    
Then set Tesseract path in your environment:
    import pytesseract
    pytesseract.pytesseract.pytesseract_cmd = r'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'


==============================================================================
BEST PRACTICES
==============================================================================

1. Always check for success in results:
    result = bridge.extract_pdf_tables(pdf_path)
    if not result.get('success'):
        logger.error(result.get('error'))
        return

2. Use try-except blocks:
    try:
        result = bridge.extract_pdf_tables(pdf_path)
    except Exception as e:
        logger.error(f"Skills error: {e}")
        return

3. Log confidence scores:
    if result['confidence'] < 0.7:
        logger.warning(f"Low confidence extraction: {result['confidence']}")

4. Handle large files:
    # Process page by page for large PDFs
    result = bridge.extract_pdf_tables(pdf_path, page_numbers=[0, 1, 2])

5. Reuse initialized bridge:
    # DON'T create new bridge each time
    # DO cache the bridge instance
    from agent.skills_bridge import get_skills_bridge
    bridge = get_skills_bridge()  # Singleton instance


==============================================================================
TROUBLESHOOTING
==============================================================================

Q: PDF extraction returns no tables?
A: Check if PDF is native (text) or scanned. Enable OCR for scanned PDFs.
   Use confidence_threshold=0.5 for lower quality PDFs.

Q: Excel validation reports false positives?
A: Review your validation rules. Ensure numeric columns are properly typed.
   Use pd.to_numeric(df['col'], errors='coerce') to clean data first.

Q: Chart generation fails?
A: Check matplotlib/seaborn installation. Ensure data is in correct format.
   Use df.head() to inspect data structure.

Q: Mail merge creates empty documents?
A: Verify merge field names match exactly: {{field_name}}
   Check that merge_data contains all required fields.

Q: Word document creation slow?
A: Large documents with many images take time. Monitor with progress_callback.
   Save intermediate results to avoid losing work.


==============================================================================
FOR MORE INFO
==============================================================================

See:
- agent/skills_registry.py: Skills mapping configuration
- agent/skills_pdf.py: PDF extraction implementation
- agent/skills_excel.py: Excel analysis implementation
- agent/skills_visualization.py: Chart creation
- agent/skills_docx.py: Word document generation
- agent/skills_bridge.py: Integration layer

Contact: audit@auditpro.local
"""
