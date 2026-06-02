# Monitoring Performa — Drive DOCX placeholders

Add these **plain text** tags in your blank Word template (data row only). Do not change layout, tables, fonts, or signatures.

| Column | Placeholder |
|--------|-------------|
| Name of Work | `{{WORK_NAME}}` |
| Name of Agency | `{{AGENCY_NAME}}` |
| Estt. Cost (in Lacs) | `{{EST_COST_LACS}}` |
| Gross Amt. (in Lacs) | `{{GROSS_AMT_LACS}}` |
| Net Amt. (in Lacs) | `{{NET_AMT_LACS}}` |
| Sample report if any | `{{SAMPLE_REPORT}}` |
| Comments if any | `{{COMMENTS}}` |

After upload to Google Drive, set `MONITORING_TEMPLATE_ID` in `code.gs.txt` to the file ID from the URL: `https://drive.google.com/file/d/FILE_ID/view`

Generated PDFs are saved under `MONITORING_OUTPUT_FOLDER_ID` (defaults to BG letter folder).
