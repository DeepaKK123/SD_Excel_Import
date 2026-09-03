SUBROUTINE SR.RECORD.BUILD(ROW, START.ATTR, REC.ID, NEW.REC, ERR.MSG)
*=============================================================================
* SR.RECORD.BUILD
*
* Purpose:
*   Converts a single parsed data row (VM-delimited fields, as produced by
*   SR.DATA.PARSE) into an MV dynamic array ready to be written to a data
*   file. Applies the client's column-offset rule:
*     - Column A of the row becomes the Record ID
*     - Column B onward is written starting at START.ATTR (normally 11,
*       since attributes 1-10 are reserved)
*
* Parameters:
*   ROW        (IN)  - One VM-delimited row, e.g. from DATA.ROWS<n>
*   START.ATTR (IN)  - First attribute number to write column data into
*                       (pass DATA.START.ATTR from IMPORT.EQUATES)
*   REC.ID     (OUT) - Record key, taken from Column A of the row
*   NEW.REC    (OUT) - Dynamic array built from the row, ready to WRITE
*   ERR.MSG    (OUT) - Empty on success, description on failure
*                       (e.g. blank Record ID)
*
* Notes:
* This subroutine only builds the in-memory record.
*=============================================================================
   $CATALOGUE

   REC.ID  = ''
   NEW.REC = ''
   ERR.MSG = ''

   REC.ID = TRIM(ROW<1, 1>)
   * Record IDs may not contain spaces. Preserve meaningful hyphens while
   * removing spaces around them and replacing other spaces with hyphens.
   REC.ID = CHANGE(REC.ID, ' - ', '-')
   REC.ID = CHANGE(REC.ID, ' ', '-')
   IF REC.ID = '' THEN
      ERR.MSG = 'Row has a blank Record ID (Column A) - skipped'
      RETURN
   END

   NUM.COLS = DCOUNT(ROW, @VM)

   FOR COL.NO = 2 TO NUM.COLS
      FIELD.VALUE = TRIM(ROW<1, COL.NO>)
      IF FIELD.VALUE = '-' THEN FIELD.VALUE = ''
      NEW.REC<COL.NO + (START.ATTR - 2)> = FIELD.VALUE
   NEXT COL.NO

   RETURN
*=============================================================================
END
