SUBROUTINE SR.CONTROL.DICT.BUILD(CONTROL.FILE, ERR.MSG)
*=============================================================================
* SR.CONTROL.DICT.BUILD
*
* Purpose:
*   Ensure IMPORT.CONTROL exists and its DICT items are present/updated.
*   Centralizes control-dictionary maintenance for install/upgrade programs.
*
* Parameters:
*   CONTROL.FILE (IN)  - Control filename (normally IMPORT.CONTROL)
*   ERR.MSG      (OUT) - Empty on success, description on failure
*=============================================================================
   $CATALOGUE

   $INCLUDE IMPORT.EQUATES

   ERR.MSG = ''

   OPEN CONTROL.FILE TO F.CTRL ELSE
      EXECUTE 'CREATE.FILE ' : CONTROL.FILE : ' DYNAMIC' CAPTURING CREATE.OUT
      OPEN CONTROL.FILE TO F.CTRL ELSE
         ERR.MSG = 'Cannot create/open control file: ' : CONTROL.FILE : ' - ' : CREATE.OUT
         RETURN
      END
   END

   OPEN 'DICT', CONTROL.FILE TO F.CTRL.DICT ELSE
      EXECUTE 'CREATE.FILE ' : CONTROL.FILE : ' DYNAMIC' CAPTURING CREATE.OUT
      OPEN 'DICT', CONTROL.FILE TO F.CTRL.DICT ELSE
         ERR.MSG = 'Cannot create/open dictionary for file: ' : CONTROL.FILE : ' - ' : CREATE.OUT
         CLOSE F.CTRL
         RETURN
      END
   END

   GOSUB WRITE.ID.DICT

   DICT.CONV = ''
   DICT.ID = 'CURR.TYPE'      ; DICT.LOC = CTRL.CURR.TYPE      ; DICT.NAME = 'Current load type'                ; DICT.FMT = '20L' ; GOSUB WRITE.CONTROL.DICT.ITEM
   DICT.ID = 'CURR.DATE'      ; DICT.LOC = CTRL.CURR.DATE      ; DICT.NAME = 'Current load date'                ; DICT.CONV = 'D4/' ; DICT.FMT = '12L' ; GOSUB WRITE.CONTROL.DICT.ITEM
   DICT.ID = 'CURR.TIME'      ; DICT.LOC = CTRL.CURR.TIME      ; DICT.NAME = 'Current load time'                ; DICT.CONV = 'MTH' ; DICT.FMT = '10L' ; GOSUB WRITE.CONTROL.DICT.ITEM
   DICT.ID = 'CURR.COUNT'     ; DICT.LOC = CTRL.CURR.COUNT     ; DICT.NAME = 'Current record count'             ; DICT.FMT = '10R' ; GOSUB WRITE.CONTROL.DICT.ITEM
   DICT.ID = 'CURR.ADDED'     ; DICT.LOC = CTRL.CURR.ADDED     ; DICT.NAME = 'Current added count'              ; DICT.FMT = '10R' ; GOSUB WRITE.CONTROL.DICT.ITEM
   DICT.ID = 'CURR.CHANGED'   ; DICT.LOC = CTRL.CURR.CHANGED   ; DICT.NAME = 'Current changed count'            ; DICT.FMT = '10R' ; GOSUB WRITE.CONTROL.DICT.ITEM
   DICT.ID = 'CURR.DELETED'   ; DICT.LOC = CTRL.CURR.DELETED   ; DICT.NAME = 'Current deleted count'            ; DICT.FMT = '10R' ; GOSUB WRITE.CONTROL.DICT.ITEM
   DICT.ID = 'CURR.DIFFPATH'  ; DICT.LOC = CTRL.CURR.DIFFPATH  ; DICT.NAME = 'Current diff report file'         ; DICT.FMT = '80L' ; GOSUB WRITE.CONTROL.DICT.ITEM
   DICT.ID = 'CURR.ACCOUNT'   ; DICT.LOC = CTRL.CURR.ACCOUNT   ; DICT.NAME = 'Current run account'              ; DICT.FMT = '25L' ; GOSUB WRITE.CONTROL.DICT.ITEM
   DICT.ID = 'CURR.SRC.DIR'   ; DICT.LOC = CTRL.CURR.SRC.DIR   ; DICT.NAME = 'Current archive root directory'   ; DICT.FMT = '80L' ; GOSUB WRITE.CONTROL.DICT.ITEM
   DICT.ID = 'CURR.SKIPPED'   ; DICT.LOC = CTRL.CURR.SKIPPED   ; DICT.NAME = 'Current skipped count'            ; DICT.FMT = '10R' ; GOSUB WRITE.CONTROL.DICT.ITEM
   DICT.ID = 'CURR.SRC.NAME'  ; DICT.LOC = CTRL.CURR.SRC.NAME  ; DICT.NAME = 'Current source filename'          ; DICT.FMT = '50L' ; GOSUB WRITE.CONTROL.DICT.ITEM
   DICT.ID = 'CURR.CHANGE.LOG'; DICT.LOC = CTRL.CURR.CHANGE.LOG; DICT.NAME = 'Current detailed change log file' ; DICT.FMT = '80L' ; GOSUB WRITE.CONTROL.DICT.ITEM

   DICT.ID = 'PRIOR.TYPE'      ; DICT.LOC = CTRL.PRIOR.TYPE      ; DICT.NAME = 'Prior load type'                 ; DICT.FMT = '20L' ; GOSUB WRITE.CONTROL.DICT.ITEM
   DICT.ID = 'PRIOR.DATE'      ; DICT.LOC = CTRL.PRIOR.DATE      ; DICT.NAME = 'Prior load date'                 ; DICT.CONV = 'D4/' ; DICT.FMT = '12L' ; GOSUB WRITE.CONTROL.DICT.ITEM
   DICT.ID = 'PRIOR.TIME'      ; DICT.LOC = CTRL.PRIOR.TIME      ; DICT.NAME = 'Prior load time'                 ; DICT.CONV = 'MTH' ; DICT.FMT = '10L' ; GOSUB WRITE.CONTROL.DICT.ITEM
   DICT.ID = 'PRIOR.COUNT'     ; DICT.LOC = CTRL.PRIOR.COUNT     ; DICT.NAME = 'Prior record count'              ; DICT.FMT = '10R' ; GOSUB WRITE.CONTROL.DICT.ITEM
   DICT.ID = 'PRIOR.ADDED'     ; DICT.LOC = CTRL.PRIOR.ADDED     ; DICT.NAME = 'Prior added count'               ; DICT.FMT = '10R' ; GOSUB WRITE.CONTROL.DICT.ITEM
   DICT.ID = 'PRIOR.CHANGED'   ; DICT.LOC = CTRL.PRIOR.CHANGED   ; DICT.NAME = 'Prior changed count'             ; DICT.FMT = '10R' ; GOSUB WRITE.CONTROL.DICT.ITEM
   DICT.ID = 'PRIOR.DELETED'   ; DICT.LOC = CTRL.PRIOR.DELETED   ; DICT.NAME = 'Prior deleted count'             ; DICT.FMT = '10R' ; GOSUB WRITE.CONTROL.DICT.ITEM
   DICT.ID = 'PRIOR.DIFFPATH'  ; DICT.LOC = CTRL.PRIOR.DIFFPATH  ; DICT.NAME = 'Prior diff report file'          ; DICT.FMT = '80L' ; GOSUB WRITE.CONTROL.DICT.ITEM
   DICT.ID = 'PRIOR.ACCOUNT'   ; DICT.LOC = CTRL.PRIOR.ACCOUNT   ; DICT.NAME = 'Prior run account'               ; DICT.FMT = '25L' ; GOSUB WRITE.CONTROL.DICT.ITEM
   DICT.ID = 'PRIOR.SRC.DIR'   ; DICT.LOC = CTRL.PRIOR.SRC.DIR   ; DICT.NAME = 'Prior archive root directory'    ; DICT.FMT = '80L' ; GOSUB WRITE.CONTROL.DICT.ITEM
   DICT.ID = 'PRIOR.SKIPPED'   ; DICT.LOC = CTRL.PRIOR.SKIPPED   ; DICT.NAME = 'Prior skipped count'             ; DICT.FMT = '10R' ; GOSUB WRITE.CONTROL.DICT.ITEM
   DICT.ID = 'PRIOR.SRC.NAME'  ; DICT.LOC = CTRL.PRIOR.SRC.NAME  ; DICT.NAME = 'Prior source filename'           ; DICT.FMT = '50L' ; GOSUB WRITE.CONTROL.DICT.ITEM
   DICT.ID = 'PRIOR.CHANGE.LOG'; DICT.LOC = CTRL.PRIOR.CHANGE.LOG; DICT.NAME = 'Prior detailed change log file'  ; DICT.FMT = '80L' ; GOSUB WRITE.CONTROL.DICT.ITEM

   CLOSE F.CTRL.DICT
   CLOSE F.CTRL
   RETURN

*-----------------------------------------------------------------------
WRITE.ID.DICT:
* Ensure the standard @ID dictionary item exists.
*-----------------------------------------------------------------------
   ID.REC = ''
   ID.REC<DICT.TYPE.ATTR>   = 'D'
   ID.REC<DICT.LOC.ATTR>    = 0
   ID.REC<DICT.CONV.ATTR>   = ''
   ID.REC<DICT.NAME.ATTR>   = CONTROL.FILE
   ID.REC<DICT.FORMAT.ATTR> = '50L'
   ID.REC<DICT.SM.ATTR>     = 'S'
   ID.REC<DICT.ASSOC.ATTR>  = ''
   WRITE ID.REC TO F.CTRL.DICT, '@ID'
   RETURN

*-----------------------------------------------------------------------
WRITE.CONTROL.DICT.ITEM:
* Create/refresh one D-type dictionary item for IMPORT.CONTROL.
*-----------------------------------------------------------------------
   D.REC = ''
   D.REC<DICT.TYPE.ATTR>   = 'D'
   D.REC<DICT.LOC.ATTR>    = DICT.LOC
   D.REC<DICT.CONV.ATTR>   = DICT.CONV
   D.REC<DICT.NAME.ATTR>   = DICT.NAME
   D.REC<DICT.FORMAT.ATTR> = DICT.FMT
   D.REC<DICT.SM.ATTR>     = 'S'
   D.REC<DICT.ASSOC.ATTR>  = ''
   WRITE D.REC TO F.CTRL.DICT, DICT.ID
   DICT.CONV = ''
   RETURN
*=============================================================================
END
