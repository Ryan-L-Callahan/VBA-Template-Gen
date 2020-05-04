# VBA-Template-Gen
Assumes Column D, Row 2-1000 holds consecutive Standard Field Names.
Typically, Column A will hold consecutive field names, Column B and C will hold consecutive start position and length. 
Run the Template Macro, then the Hyperlink Macro.
Paste the results of the following query to the Batch Metadata tab, where <table> is changed to your table:
  SELECT btch_id, cyc_dt, file_nm, COUNT(1) as record_count
  FROM <table>
  GROUP BY btch_id, cyc_dt, file_nm
Select your working batch and paste it's value as in the 'Layout' sheet, cell F1.
Paste frequency distributions in the associated tab, formulas will autofill based on the frequency distribution
Any values in the fill count column will highlight in red if the record count in the frequency distribution does not match your chosen batch.
