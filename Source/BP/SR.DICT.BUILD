SUBROUTINE SR.DICT.BUILD(DATA.FILE, HEADER.ROW, NEXT.ATTR, ERR.MSG)
*=============================================================================
* SR.DICT.BUILD
*
* Purpose:
*   For each column name in HEADER.ROW, ensures a matching DICT item
*   exists in the dictionary of DATA.FILE. Creates any that are missing.
*   Existing columns (matched on original column name, DICT attr 4) retain
*   their attribute number, but their conversion code is cleared so imported
*   values remain text.
*
* Parameters:
*   DATA.FILE  (IN)     - Name of the MV data file whose DICT is updated
*   HEADER.ROW (IN)     - VM-delimited column names, as read from the file
*   NEXT.ATTR  (IN/OUT) - Next available attribute number to assign to a
*                          brand-new column. Caller should seed this with
*                          DATA.START.ATTR (11) on a first-ever load, or
*                          the highest attribute number in use otherwise.
*   ERR.MSG    (OUT)    - Empty on success, description on failure
*
* Rules applied (per client spec):
*   - Column A is the record ID and is described by the @ID item.
*   - DICT item-id  = column name with spaces changed to '-'
*   - DICT attr 4   = display name, spaces preserved
*                     (kept because future files reference columns by
*                      their original name)
*=============================================================================

   $CATALOGUE

   $INCLUDE IMPORT.EQUATES

   ERR.MSG = ''

   OPEN 'DICT', DATA.FILE TO F.DICT ELSE
      ERR.MSG = 'Cannot open dictionary for target MV file: ' : DATA.FILE
      RETURN
   END

   GOSUB BUILD.ID.DICT

   NUM.COLS = DCOUNT(HEADER.ROW, @VM)

   * Column A is the MV record ID, not a data attribute. Column B starts at 11.
   FOR COL.IDX = 2 TO NUM.COLS
      COL.NAME = TRIM(HEADER.ROW<1, COL.IDX>)
      IF COL.NAME = '' THEN CONTINUE

      GOSUB FIND.EXISTING.COLUMN

      IF NOT(FOUND.ATTR) THEN
         GOSUB CREATE.DICT.ITEM
      END
   NEXT COL.IDX

   CLOSE F.DICT

   RETURN

*-----------------------------------------------------------------------
BUILD.ID.DICT:
* Maintain the standard @ID dictionary item. Location zero refers to the
* item ID, and the manual's D-type layout places the display format in field 5.
*-----------------------------------------------------------------------
   ID.REC = ''
   ID.REC<DICT.TYPE.ATTR>   = 'D'
   ID.REC<DICT.LOC.ATTR>    = 0
   ID.REC<DICT.CONV.ATTR>   = ''
   ID.REC<DICT.NAME.ATTR>   = DATA.FILE
   ID.REC<DICT.FORMAT.ATTR> = '50L'
   ID.REC<DICT.SM.ATTR>     = 'S'
   ID.REC<DICT.ASSOC.ATTR>  = ''
   WRITE ID.REC TO F.DICT, '@ID'
   RETURN

*-----------------------------------------------------------------------
FIND.EXISTING.COLUMN:
* Scan existing DICT items for one whose attr 4 matches COL.NAME exactly.
* Sets FOUND.ATTR = 1 if found (existing attribute number is reused as-is;
* no further action needed for that column).
*-----------------------------------------------------------------------
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
   RETURN

*-----------------------------------------------------------------------
CREATE.DICT.ITEM:
* Creates a new D-type dictionary item for COL.NAME at NEXT.ATTR,
* then advances NEXT.ATTR for the following new column.
*-----------------------------------------------------------------------
   DICT.NAME = CHANGE(COL.NAME, ' ', '-')

   D.REC = ''
   D.REC<DICT.TYPE.ATTR>    = 'D'
   D.REC<DICT.LOC.ATTR>     = NEXT.ATTR
   D.REC<DICT.CONV.ATTR>    = ''
   D.REC<DICT.NAME.ATTR>    = COL.NAME
   D.REC<DICT.FORMAT.ATTR>  = '20L'
   D.REC<DICT.SM.ATTR>      = 'S'
   D.REC<DICT.ASSOC.ATTR>   = ''

   WRITE D.REC TO F.DICT, DICT.NAME

   NEXT.ATTR += 1
   RETURN
*=============================================================================
END
