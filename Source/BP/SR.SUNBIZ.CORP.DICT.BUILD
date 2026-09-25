SUBROUTINE SR.SUNBIZ.CORP.DICT.BUILD(DATA.FILE, NEXT.ATTR, ERR.MSG)
*=============================================================================
* SR.SUNBIZ.CORP.DICT.BUILD
*
* Purpose:
*   Same job as the generic SR.DICT.BUILD, but sized to the actual
*   SunBiz Corporate Data File layout: each DICT item's display FORMAT
*   matches that field's real fixed-width length (SR.SUNBIZ.CORP.LENGTHS)
*   instead of a generic guess. Column names come from
*   SR.SUNBIZ.CORP.HEADER; column 1 is the Record ID (@ID), columns 2+
*   are data attributes starting at NEXT.ATTR.
*
* Parameters:
*   DATA.FILE  (IN)     - Name of the MV data file whose DICT is updated
*   NEXT.ATTR  (IN/OUT) - Next available attribute number to assign to a
*                          brand-new column. Seed with DATA.START.ATTR.
*   ERR.MSG    (OUT)    - Empty on success, description on failure
*
* Rules applied (same as SR.DICT.BUILD):
*   - Column 1 is the record ID and is described by the @ID item.
*   - DICT item-id  = column name, uppercased, spaces/'+' changed to '-'
*   - DICT attr 4   = display name, original casing/spaces/'+' preserved
*=============================================================================
   $CATALOGUE

   $INCLUDE IMPORT.EQUATES

   ERR.MSG = ''

   CALL SR.SUNBIZ.CORP.HEADER(HEADER.ROW)
   CALL SR.SUNBIZ.CORP.LENGTHS(LENGTHS.ROW)

   OPEN 'DICT', DATA.FILE TO F.DICT ELSE
      ERR.MSG = 'Cannot open dictionary for target MV file: ' : DATA.FILE
      RETURN
   END

   ID.LEN = LENGTHS.ROW<1, 1>
   ID.REC = ''
   ID.REC<DICT.TYPE.ATTR>   = 'D'
   ID.REC<DICT.LOC.ATTR>    = 0
   ID.REC<DICT.CONV.ATTR>   = ''
   ID.REC<DICT.NAME.ATTR>   = DATA.FILE
   ID.REC<DICT.FORMAT.ATTR> = ID.LEN : 'L'
   ID.REC<DICT.SM.ATTR>     = 'S'
   ID.REC<DICT.ASSOC.ATTR>  = ''
   WRITE ID.REC TO F.DICT, '@ID'

   NUM.COLS = DCOUNT(HEADER.ROW, @VM)

   FOR COL.IDX = 2 TO NUM.COLS
      COL.NAME = TRIM(HEADER.ROW<1, COL.IDX>)
      IF COL.NAME = '' THEN CONTINUE
      COL.LEN = LENGTHS.ROW<1, COL.IDX>

      FOUND.ATTR = 0
      SELECT F.DICT
      LOOP
         READNEXT SCAN.ID ELSE EXIT
         READ SCAN.REC FROM F.DICT, SCAN.ID THEN
            IF SCAN.REC<DICT.NAME.ATTR> = COL.NAME THEN
               IF SCAN.REC<DICT.CONV.ATTR> NE '' THEN
                  SCAN.REC<DICT.CONV.ATTR> = ''
                  WRITE SCAN.REC TO F.DICT, SCAN.ID
               END
               FOUND.ATTR = 1
               EXIT
            END
         END
      REPEAT

      IF NOT(FOUND.ATTR) THEN
         DICT.NAME = UPCASE(CHANGE(CHANGE(COL.NAME, ' ', '-'), '+', '-'))

         D.REC = ''
         D.REC<DICT.TYPE.ATTR>    = 'D'
         D.REC<DICT.LOC.ATTR>     = NEXT.ATTR
         D.REC<DICT.CONV.ATTR>    = ''
         D.REC<DICT.NAME.ATTR>    = COL.NAME
         D.REC<DICT.FORMAT.ATTR>  = COL.LEN : 'L'
         D.REC<DICT.SM.ATTR>      = 'S'
         D.REC<DICT.ASSOC.ATTR>   = ''

         WRITE D.REC TO F.DICT, DICT.NAME
         NEXT.ATTR += 1
      END
   NEXT COL.IDX

   CLOSE F.DICT
   RETURN
*=============================================================================
END
