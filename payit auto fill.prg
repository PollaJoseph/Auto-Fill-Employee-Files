

_SCREEN.Visible = .F.

PUBLIC oForm
oForm = CREATEOBJECT("PayITForm")
oForm.Show()
READ EVENTS

_SCREEN.Visible = .T.

DEFINE CLASS PayITForm AS FORM
    Caption = "PayIT Auto Fill System"
    BackColor = RGB(245,247,250) 
    FontName = "Segoe UI"
    FontSize = 10
    
    *-- Window Settings
    BorderStyle = 3  && Sizable
    MaxButton = .T.
    MinButton = .T.
    ControlBox = .T.
    ShowWindow = 2   && Top-Level Form
    WindowState = 2  && Start Maximized
    MinWidth = 800
    MinHeight = 600

    *-- UI HEADER
    ADD OBJECT cntHeader AS Container WITH Top = 0, Left = 0, Height = 100, BackColor = RGB(44, 62, 80), BorderWidth = 0, Anchor = 10, Width = 800
    ADD OBJECT imgLogo AS Image WITH Picture = "I:\PayIT\PayIT Update Master File\CCC.png", Top = 15, Height = 90, Width = 90, Stretch = 1, BackStyle = 0
    ADD OBJECT lblTitle AS Label WITH Caption = "PayIT Data Importer", FontName = "Segoe UI", FontSize = 20, FontBold = .T., ForeColor = RGB(255, 255, 255), BackStyle = 0, Top = 45, AutoSize = .T.
    
    *-- INPUT CONTROLS
    ADD OBJECT lblInstr AS Label WITH Caption = "Select Excel File (Old data will be deleted automatically)", Top = 140, AutoSize = .T., ForeColor = RGB(100, 100, 100)
    ADD OBJECT txtFile AS TextBox WITH Top = 165, Width = 400, Height = 30, FontName = "Segoe UI", FontSize = 10, ReadOnly = .T.
    ADD OBJECT cmdBrowse AS CommandButton WITH Caption = "...", Top = 165, Width = 40, Height = 30
    ADD OBJECT lblCalendar AS Label WITH Caption = "Calendar:", Top = 215, AutoSize = .T., FontName = "Segoe UI", FontSize = 11, FontBold = .T.
    ADD OBJECT cboCalendar AS ComboBox WITH Top = 210, Width = 200, Height = 30, FontName = "Segoe UI", FontSize = 10, RowSourceType = 1, RowSource = "Daily,Monthly,Hired", Style = 2, Value = "Daily"
    
    *-- ACTION BUTTONS
    * Import Button (Green)
    ADD OBJECT cmdImport AS CommandButton WITH Caption = "START IMPORT", Top = 260, Width = 200, Height = 40, FontName = "Segoe UI", FontSize = 11, FontBold = .T., BackColor = RGB(0, 255, 0), ForeColor = 0
    * Exit Button (Red)
    ADD OBJECT cmdExit AS CommandButton WITH Caption = "EXIT SYSTEM", Top = 310, Width = 200, Height = 40, FontName = "Segoe UI", FontSize = 11, FontBold = .T., BackColor = 255, ForeColor = 0 , BackStyle=0
    
    ADD OBJECT lblStatus AS Label WITH Caption = "", Top = 370, AutoSize = .T., ForeColor = RGB(100, 100, 100)

    PROCEDURE Init
        This.Resize()
    ENDPROC

    *-- Resize Logic to keep everything centered
    PROCEDURE Resize
        LOCAL lnCenter
        lnCenter = This.Width / 2
        
        This.imgLogo.Left    = 30
        This.lblTitle.Left   = 140
        
        * Center Inputs
        This.txtFile.Left    = lnCenter - 225
        This.lblInstr.Left   = This.txtFile.Left
        This.cmdBrowse.Left  = This.txtFile.Left + 410
        
        This.lblCalendar.Left = This.txtFile.Left
        This.cboCalendar.Left = This.txtFile.Left + 80
        
        * Center Buttons
        This.cmdImport.Left   = lnCenter - 100  && Center the button perfectly
        This.cmdExit.Left     = lnCenter - 100  && Center the Exit button
        
        This.lblStatus.Left   = This.txtFile.Left
    ENDPROC

    PROCEDURE Destroy
        CLEAR EVENTS
    ENDPROC

    PROCEDURE imgLogo.Init
        LOCAL lcPath
        lcPath = ADDBS(JUSTPATH(SYS(16,0))) 
        IF FILE(lcPath + "CCC.png")
            This.Picture = lcPath + "CCC.png"
        ENDIF
    ENDPROC

    *-- EXIT BUTTON LOGIC
    PROCEDURE cmdExit.Click
        LOCAL lnChoice
        lnChoice = MESSAGEBOX("Are you sure you want to exit?", 4+32, "Exit System")
        IF lnChoice = 6  && User clicked Yes
            ThisForm.Release
        ENDIF
    ENDPROC

    PROCEDURE cmdBrowse.Click
        LOCAL lcFile
        lcFile = GETFILE("XLS;XLSX", "Select Excel File", "Open", 0)
        IF !EMPTY(lcFile)
            ThisForm.txtFile.Value = lcFile
        ENDIF
    ENDPROC

    PROCEDURE cmdImport.Click
        IF EMPTY(ThisForm.txtFile.Value)
            MESSAGEBOX("Please select a file first.", 48, "Error")
            RETURN
        ENDIF
        ThisForm.ProcessImport(ThisForm.txtFile.Value)
    ENDPROC

    *-- IMPORT PROCESS
    PROCEDURE ProcessImport(tcFile)
        LOCAL loExcel, loBook, loSheetH, loSheetD
        LOCAL lnLastRow, i
        LOCAL llHeaderFound, llDetailsFound
        LOCAL lcBasePath, lcSubFolder, lcPayResultH, lcPayResultD, lcCalendarType
        LOCAL lnErrorCount
        lnErrorCount = 0
        
        SET SAFETY OFF 

        lcCalendarType = ThisForm.cboCalendar.Value
        lcBasePath = "FILE_PATH"
        
        DO CASE
            CASE lcCalendarType = "Daily"
                lcSubFolder = "Daily"
            CASE lcCalendarType = "Monthly"
                lcSubFolder = "Monthly"
            CASE lcCalendarType = "Hired"
                lcSubFolder = "Hired"
            OTHERWISE
                MESSAGEBOX("Please select a Calendar type.", 16, "Error")
                RETURN
        ENDCASE
        
        lcPayResultH = ADDBS(lcBasePath) + ADDBS(lcSubFolder) + "PayResultH.dbf"
        lcPayResultD = ADDBS(lcBasePath) + ADDBS(lcSubFolder) + "PayResultD.dbf"
        
        IF !FILE(lcPayResultH) OR !FILE(lcPayResultD)
            MESSAGEBOX("DBF files not found in: " + lcSubFolder, 16, "Error")
            RETURN
        ENDIF

        ThisForm.lblStatus.Caption = "Initializing Excel..."
        
        TRY
            loExcel = CREATEOBJECT("Excel.Application")
            loBook = loExcel.Workbooks.Open(tcFile)
        CATCH
            MESSAGEBOX("Could not open Excel. Make sure it is not already open.", 16, "Error")
            RETURN
        ENDTRY

        TRY
            loSheetH = loBook.Sheets("Header")
            llHeaderFound = .T.
        CATCH
            llHeaderFound = .F.
        ENDTRY

        TRY
            loSheetD = loBook.Sheets("Details")
            llDetailsFound = .T.
        CATCH
            llDetailsFound = .F.
        ENDTRY

        SET DATE YMD
        SET CENTURY ON

        IF llHeaderFound
            ThisForm.lblStatus.Caption = "Deleting Old Header Data..."
            lnLastRow = loSheetH.Cells(loSheetH.Rows.Count, 1).End(-4162).Row
            
            IF USED("PayResultH")
                USE IN PayResultH
            ENDIF
            USE (lcPayResultH) IN 0 ALIAS PayResultH EXCLUSIVE
            SELECT PayResultH
            ZAP 

            ThisForm.lblStatus.Caption = "Importing New Header Data..."
            FOR i = 2 TO lnLastRow
                TRY
                    APPEND BLANK
                    ThisForm.SmartReplace("ACTIVE",     loSheetH.Cells(i, 1).Value)
                    ThisForm.SmartReplace("BADGE_CD",   loSheetH.Cells(i, 2).Value)
                    ThisForm.SmartReplace("BADGE_CHR",  loSheetH.Cells(i, 3).Value)
                    ThisForm.SmartReplace("COMPANY",    loSheetH.Cells(i, 4).Value)
                    ThisForm.SmartReplace("GROUPNO",    loSheetH.Cells(i, 5).Value)
                    ThisForm.SmartReplace("NAME",       loSheetH.Cells(i, 6).Value)
                    ThisForm.SmartReplace("NATION",     loSheetH.Cells(i, 7).Value)
                    ThisForm.SmartReplace("TITLECODE",  loSheetH.Cells(i, 8).Value)
                    ThisForm.SmartReplace("TITLEDESC",  loSheetH.Cells(i, 9).Value)
                    ThisForm.SmartReplace("LOCCODE",    loSheetH.Cells(i, 10).Value)
                    ThisForm.SmartReplace("LOCDESC",    loSheetH.Cells(i, 11).Value)
                    ThisForm.SmartReplace("LOC_DT",     loSheetH.Cells(i, 12).Value)
                    ThisForm.SmartReplace("PERSONALNO", loSheetH.Cells(i, 13).Value)
                    ThisForm.SmartReplace("WP_NO",      loSheetH.Cells(i, 14).Value)
                    ThisForm.SmartReplace("WP_ISS",     loSheetH.Cells(i, 15).Value)
                    ThisForm.SmartReplace("WP_EXP",     loSheetH.Cells(i, 16).Value)
                    ThisForm.SmartReplace("WP_SPONSOR", loSheetH.Cells(i, 17).Value)
                    ThisForm.SmartReplace("WP_CITY",    loSheetH.Cells(i, 18).Value)
                    ThisForm.SmartReplace("WP_AUTHISS", loSheetH.Cells(i, 19).Value)
                    ThisForm.SmartReplace("PAY_STATUS", loSheetH.Cells(i, 20).Value)
                    ThisForm.SmartReplace("COMPSTAT",   loSheetH.Cells(i, 21).Value)
                    ThisForm.SmartReplace("CONTR_TYPE", loSheetH.Cells(i, 22).Value)
                    ThisForm.SmartReplace("DTCONTRACT", loSheetH.Cells(i, 23).Value)
                    ThisForm.SmartReplace("CONTR_END",  loSheetH.Cells(i, 24).Value)
                    ThisForm.SmartReplace("DTERMINATE", loSheetH.Cells(i, 25).Value)
                    ThisForm.SmartReplace("WORKING_HR", loSheetH.Cells(i, 26).Value)
                    ThisForm.SmartReplace("OT_CAT",     loSheetH.Cells(i, 27).Value)
                    ThisForm.SmartReplace("OT_CATDESC", loSheetH.Cells(i, 28).Value)
                    ThisForm.SmartReplace("OT_PAIDHRS", loSheetH.Cells(i, 29).Value)
                    ThisForm.SmartReplace("OT_NORMAL",  loSheetH.Cells(i, 30).Value)
                    ThisForm.SmartReplace("OT_WEEKEND", loSheetH.Cells(i, 31).Value)
                    ThisForm.SmartReplace("OT_HOLIDAY", loSheetH.Cells(i, 32).Value)
                    ThisForm.SmartReplace("OT_NIGHT",   loSheetH.Cells(i, 33).Value)
                    ThisForm.SmartReplace("OT_EXCPTN",  loSheetH.Cells(i, 34).Value)
                    ThisForm.SmartReplace("OTA",        loSheetH.Cells(i, 44).Value)
                    ThisForm.SmartReplace("OTA_PAYFRI", loSheetH.Cells(i, 46).Value)
                    
                    IF MOD(i, 50) = 0
                        WAIT WINDOW "Importing Header Row: " + ALLTRIM(STR(i)) + " / " + ALLTRIM(STR(lnLastRow)) NOWAIT
                    ENDIF
                CATCH TO oErr
                    lnErrorCount = lnErrorCount + 1
                ENDTRY
            ENDFOR
            USE IN PayResultH
        ENDIF

        IF llDetailsFound
            ThisForm.lblStatus.Caption = "Deleting Old Details Data..."
            lnLastRow = loSheetD.Cells(loSheetD.Rows.Count, 1).End(-4162).Row 
            
            IF USED("PayResultD")
                USE IN PayResultD
            ENDIF
            USE (lcPayResultD) IN 0 ALIAS PayResultD EXCLUSIVE
            SELECT PayResultD
            ZAP 

            ThisForm.lblStatus.Caption = "Importing New Details Data..."
            FOR i = 2 TO lnLastRow
                TRY
                    APPEND BLANK
                    ThisForm.SmartReplace("BADGE_CD",   loSheetD.Cells(i, 1).Value)
                    ThisForm.SmartReplace("CODE",       loSheetD.Cells(i, 2).Value)
                    ThisForm.SmartReplace("DESCR",      loSheetD.Cells(i, 3).Value)
                    ThisForm.SmartReplace("EFF_FROM",   loSheetD.Cells(i, 4).Value)
                    ThisForm.SmartReplace("EFF_TO",     loSheetD.Cells(i, 5).Value)
                    ThisForm.SmartReplace("TYPE",       loSheetD.Cells(i, 6).Value)
                    ThisForm.SmartReplace("VALUE",      loSheetD.Cells(i, 7).Value)
                    ThisForm.SmartReplace("CURRENCY",   loSheetD.Cells(i, 8).Value)
                    ThisForm.SmartReplace("AMOUNT",     loSheetD.Cells(i, 9).Value)
                    ThisForm.SmartReplace("FREQUENCY",  loSheetD.Cells(i, 10).Value)
                    ThisForm.SmartReplace("PAY_EOS",    loSheetD.Cells(i, 11).Value)
                    ThisForm.SmartReplace("PAY_LEAVE",  loSheetD.Cells(i, 12).Value)
                    ThisForm.SmartReplace("PAY_OVTIME", loSheetD.Cells(i, 13).Value)
                    ThisForm.SmartReplace("PAY_UNPAID", loSheetD.Cells(i, 14).Value)
                    ThisForm.SmartReplace("MAX_UNPAID", loSheetD.Cells(i, 15).Value)
                    ThisForm.SmartReplace("PAY_BUSNES", loSheetD.Cells(i, 16).Value)
                    ThisForm.SmartReplace("MAX_BUSNES", loSheetD.Cells(i, 17).Value)
                    ThisForm.SmartReplace("ACTSHT_NO",  loSheetD.Cells(i, 18).Value)
                    ThisForm.SmartReplace("CLOSE_ACTN", loSheetD.Cells(i, 19).Value)
                    ThisForm.SmartReplace("CREATED",    loSheetD.Cells(i, 20).Value)
                    ThisForm.SmartReplace("LIMPORT",    loSheetD.Cells(i, 24).Value)
                    ThisForm.SmartReplace("LNOCLOSING", loSheetD.Cells(i, 25).Value)

                    IF MOD(i, 50) = 0
                        WAIT WINDOW "Importing Details Row: " + ALLTRIM(STR(i)) + " / " + ALLTRIM(STR(lnLastRow)) NOWAIT
                    ENDIF
                CATCH TO oErr
                    lnErrorCount = lnErrorCount + 1
                ENDTRY
            ENDFOR
            USE IN PayResultD
        ENDIF

        WAIT CLEAR
        IF VARTYPE(loBook) = "O"
            loBook.Close(.F.)
        ENDIF
        IF VARTYPE(loExcel) = "O"
            loExcel.Quit
        ENDIF
        RELEASE loExcel
        SET SAFETY ON

        ThisForm.lblStatus.Caption = "Import Completed Successfully!"
        
        IF lnErrorCount > 0
            MESSAGEBOX("Import Complete with " + TRANSFORM(lnErrorCount) + " skipped rows/errors.", 48, "Done")
        ELSE
            MESSAGEBOX("Old Data Deleted & New Data Imported Successfully!", 64, "Success")
        ENDIF
    ENDPROC

    FUNCTION SmartReplace(tcField, tvValue)
        LOCAL lcType, lvCastValue
        IF TYPE(tcField) = "U"
             RETURN 
        ENDIF
        lcType = TYPE(tcField)
        IF ISNULL(tvValue) OR UPPER(TRANSFORM(tvValue)) == "NULL"
            tvValue = "" 
        ENDIF
        DO CASE
            CASE lcType = "D"
                IF VARTYPE(tvValue) = "D" OR VARTYPE(tvValue) = "T"
                    lvCastValue = tvValue
                ELSE
                    LOCAL lcDateStr
                    lcDateStr = ALLTRIM(TRANSFORM(tvValue))
                    IF EMPTY(lcDateStr)
                        lvCastValue = {//}
                    ELSE
                        TRY
                            lvCastValue = CTOD(lcDateStr)
                            IF EMPTY(lvCastValue) AND "-" $ lcDateStr
                                lvCastValue = EVALUATE("{^" + lcDateStr + "}")
                            ENDIF
                        CATCH
                            lvCastValue = {//}
                        ENDTRY
                    ENDIF
                ENDIF
            CASE lcType = "N"
                lvCastValue = VAL(TRANSFORM(tvValue))
            CASE lcType = "L"
                LOCAL cVal
                cVal = UPPER(TRANSFORM(tvValue))
                lvCastValue = (cVal == ".T." OR cVal == "TRUE" OR cVal == "YES" OR cVal == "1")
            CASE lcType = "C"
                lvCastValue = TRANSFORM(tvValue)
            OTHERWISE
                lvCastValue = tvValue
        ENDCASE
        REPLACE (tcField) WITH lvCastValue
    ENDFUNC

ENDDEFINE