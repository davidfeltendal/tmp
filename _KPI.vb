Imports Microsoft.VisualBasic

Public Class Class1
    'Attribute VB_Name "Danfoss_KPI"
    Private Declare Function ShellExecute _
  Lib "shell32.dll" Alias "ShellExecuteA" (
  ByVal hWnd As Long,
  ByVal Operation As String,
  ByVal fileName As String,
  Optional ByVal Parameters As String = "",
  Optional ByVal Directory As String = "",
  Optional ByVal WindowStyle As Long = vbMinimizedFocus
  ) As Long
    Public strMakroNavn As String
    'Variables used when dealing with Error
    Dim blDimensioned As Boolean
    'Variables used when dealing with Cargolink
    Dim period As String ': period = WorksheetFunction.Text(Date - Day(Date), "[$-409]mmm")
    Dim LastDay As Date ': LastDay = DateSerial(Year(Date), Month(Date), 0)
    Dim FirstDay As Date ': FirstDay = LastDay - Day(LastDay) + 1
    Dim layout As String ': layout = "DANF-KPI"
    Dim sLastDay As String ': sLastDay = Format(LastDay, "yymmdd")
    Dim sFirstDay As String ': sFirstDay = Format(FirstDay, "yymmdd")
    Dim iDays As Integer ': iDays = 10
    'Variables used when dealing with Outlook
    Dim iMail As Integer, sMailTo As String, sMailFrom As String
    Dim sBody As String, sSubject As String, Dest As Object

    'Variables used when dealing with Excel
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook, objTmpWorkbook As Excel.Workbook
    Dim xlFilename As String, xlFilePath As String, oldXl As String
    Dim objWorksheet As Excel.Worksheet, objTmpWorksheet As Excel.Worksheet
    Dim iXlrow As Integer, iRow As Integer, iLastXlRow As Integer
    Dim sDestArr As Object, rngStart As range
    Dim userinfo() As String, mailInfo() As String
    Dim vbResult As VbMsgBoxResult
    Sub testThis()
        Dim TEST() As String
        'Call CreateEmail("", "david.feltendal@dk.dsv.com")
        'Call WaitForMail("#J978886", 5, TEST)
        'test = DanfossKPI_getDeptMail("", "it")
        TEST = getUser()
        'Debug.Print TEST

    End Sub
    Sub DanfossKPI_Part1()
        'By DAFE 2017-06
        'sender mails rundt til div departments

        '*****************************************************************************************************
        strMakroNavn = "DanfossKPI_MailTilPolen"
        '*****************************************************************************************************

        'Variables used when dealing with Cargolink
        period = WorksheetFunction.text(Date - Day(Of Date), "[$-409]mmm")
        LastDay = DateSerial(Year(Of Date), Month(Of Date), 0)
        FirstDay = LastDay - Day(LastDay) + 1
        layout = "DANF-KPI"
        sLastDay = Format(LastDay, "yymmdd")
        sFirstDay = Format(FirstDay, "yymmdd")
        iDays = 10


        'With Session

        userinfo = getUser()
        iDays = 14 - Format(Now(), "d")

        iMail = InputBox("Send mail til:" & vbNewLine & "NL Alfred: tast 1" & vbNewLine & "PL Matyna: tast 2" & vbNewLine & "Send til begge: tast 3" _
        & vbNewLine & vbNewLine & "Send ikke mail: tast 0 eller efterlad blank", "Send Mail?", 0)
        If iMail = 1 Then
            sBody = "<p style=""margin: 0px 0px 16px;""><span style=""margin: 0px; color: black; font-family: 'Arial',sans-serif; font-size: 10pt;""><br/>" &
            "</span><span style=""margin: 0px; color: black; font-family: 'Arial',sans-serif; font-size: 10pt;"">Hi Alfred</p><br/><p>Do you have the IOD data regarding Danfoss (" & period & ") ready for me?<br />" &
            "If yes &ndash; please send it to me. No later than d." & Format(DateAdd("d", iDays, Date), "dd-mm-yyyy") & "&ndash; before 14:30 </p>" &
            "<br/><p>Med venlig hilsen, / Best regards, </p><br />" &
            "<p>" & userinfo(1) & "<br />Key Account </p>"
            sSubject = "Danfoss IOD data for " & period
            Call ExcelMail.CreateEmail("A.Bolink@ZwierVeldhoen.nl", userinfo(0), sBody, sSubject)

        ElseIf iMail = 2 Then
            'Debug.Print Format(DateAdd("d", 10, Date), "dd-mm-yyyy")
            sBody = "<p style=""margin: 0px 0px 16px;""><span style=""margin: 0px; color: black; font-family: 'Arial',sans-serif; font-size: 10pt;""><br/>" &
            "</span><span style=""margin: 0px; color: black; font-family: 'Arial',sans-serif; font-size: 10pt;"">Hi Martyna</p><br/><p>Please send me the Danfoss KPI report for " & period & ".<br />" &
            "I need the 2 reports. No later than d." & Format(DateAdd("d", iDays, Date), "dd-mm-yyyy") & "&ndash; before 14:30 </p>" &
            "<br/><p>Med venlig hilsen, / Best regards, </p><br />" &
            "<p>" & userinfo(1) & "<br />Key Account </p>"
            sSubject = "Danfoss KPI report for " & period
            Call ExcelMail.CreateEmail("Martyna.Kawinska@pl.dsv.com", userinfo(0), sBody, sSubject)

        ElseIf iMail = 3 Then
            sBody = "<p style=""margin: 0px 0px 16px;""><span style=""margin: 0px; color: black; font-family: 'Arial',sans-serif; font-size: 10pt;""><br/>" &
            "</span><span style=""margin: 0px; color: black; font-family: 'Arial',sans-serif; font-size: 10pt;"">Hi Alfred</p><br/><p>Do you have the IOD data regarding Danfoss (" & period & ") ready for me?<br />" &
            "If yes &ndash; please send it to me. No later than d." & Format(DateAdd("d", iDays, Date), "dd-mm-yyyy") & "&ndash; before 14:30 </p>" &
            "<br/><p>Med venlig hilsen, / Best regards, </p><br />" &
            "<p>" & userinfo(1) & "<br />Key Account </p>"
            sSubject = "Danfoss IOD data for " & period
            Call ExcelMail.CreateEmail("A.Bolink@ZwierVeldhoen.nl", userinfo(0), sBody, sSubject)

            sBody = "<p style=""margin: 0px 0px 16px;""><span style=""margin: 0px; color: black; font-family: 'Arial',sans-serif; font-size: 10pt;""><br/>" &
            "</span><span style=""margin: 0px; color: black; font-family: 'Arial',sans-serif; font-size: 10pt;"">Hi Martyna</p><br/><p>Please send me the Danfoss KPI report for " & period & ".<br />" &
            "I need the 2 reports. No later than d." & Format(DateAdd("d", iDays, Date), "dd-mm-yyyy") & "&ndash; before 14:30 </p>" &
            "<br/><p>Med venlig hilsen, / Best regards, </p><br />" &
            "<p>" & userinfo(1) & "<br />Key Account </p>"
            sSubject = "Danfoss KPI report for " & period
            Call ExcelMail.CreateEmail("Martyna.Kawinska@pl.dsv.com", userinfo(0), sBody, sSubject)

        Else
        End If

    End Sub
    Sub DanfossKPI_Part2()
        'By DAFE 2017-06
        'sender mails rundt til div departments

        '*****************************************************************************************************
        strMakroNavn = "DanfossKPI_MailTilDepartments"
        '*****************************************************************************************************

        'Variables used when dealing with Excel
        period = WorksheetFunction.text(Date - Day(Of Date), "[$-409]mmm")
        LastDay = DateSerial(Year(Of Date), Month(Of Date), 0)
        FirstDay = LastDay - Day(LastDay) + 1
        sLastDay = Format(LastDay, "yymmdd")
        sFirstDay = Format(FirstDay, "yymmdd")


        'With Session
        '*****************************************************************************************************
        'Definere indhold i mail
        '*****************************************************************************************************
        userinfo() = getUser()

        sBody = "<p style=""margin: 0px 0px 16px;""><span style=""margin: 0px; color: black; font-family: 'Arial',sans-serif; font-size: 10pt;""><br/>" &
            "</span><span style=""margin: 0px; color: black; font-family: 'Arial',sans-serif; font-size: 10pt;"">Hi</p><br/>" &
            "<p>Please update Cargolink IOD for these shipments, no later than d. " & Format(DateAdd("d", 7, Date), "dd-mm-yyyy") & ".<br />" &
            "<br/><p>Med venlig hilsen, / Best regards, </p><br />" &
            "<p>" & userinfo(1) & "<br />Key Account </p>"
        sSubject = "Danfoss KPI report for " & period

        '*****************************************************************************************************
        'sMailTo = "david.feltendal@dk.dsv.com"
        sMailFrom = userinfo(0)
        '*****************************************************************************************************

        objExcel = Application
        objWorkbook = objExcel.ActiveWorkbook
        xlFilename = objWorkbook.FullName

        If InStr(xlFilename, "\\") = 0 And InStr(xlFilename, ":") = 0 Then
            MsgBox("åben og gem fil fra cargolink, og kør derefter script!", vbCritical, "Error!")
            Exit Sub
        End If

        objWorksheet = objWorkbook.Sheets(1)
        'objExcel.Visible = True
        objExcel.DisplayAlerts = False

        oldXl = xlFilename
        xlFilePath = Replace(xlFilename, objWorkbook.Name, "")
        xlFilePath = xlFilePath & period & "\"

        'Opretter ny mappe hvis den ikke findes
        If Len(Dir(xlFilePath, vbDirectory)) = 0 Then
            MkDir(xlFilePath)
        End If

        'Finder sidste linje i Excel ark
        rngStart = objWorksheet.range("A1")
        iLastXlRow = rngStart.CurrentRegion.Rows.Count

        'Fjerner alle data med IOD
        objWorksheet.range("A1:AY1").AutoFilter(Field:=21, Criteria1:="<>")
        objWorksheet.range("A2:AY" & iLastXlRow).Delete(shift:=xlUp)
        objWorksheet.ShowAllData

        'Finder sidste linje i Excel ark
        iLastXlRow = rngStart.CurrentRegion.Rows.Count

        'Fjerner alle fra GB sendinger
        objWorksheet.range("A1:AY1").AutoFilter(Field:=9, Criteria1:="GB", Operator:=xlFilterValues)
        objWorksheet.range("A2:AY" & iLastXlRow).Delete(shift:=xlUp)
        objWorksheet.ShowAllData

        'Fjerner alle til NL sendinger
        objWorksheet.range("A1:AY1").AutoFilter(Field:=14, Criteria1:="NL", Operator:=xlFilterValues)
        objWorksheet.range("A2:AY" & iLastXlRow).Delete(shift:=xlUp)
        objWorksheet.ShowAllData

        'Fjerner ikke nødvendige colums
        objWorksheet.range("B:E").EntireColumn.Delete(shift:=xlRight)
        objWorksheet.range("M:N").EntireColumn.Delete(shift:=xlRight)
        objWorksheet.range("Q:AJ").EntireColumn.Delete(shift:=xlRight)
        objWorksheet.range("R:S").EntireColumn.Delete(shift:=xlRight)
        objWorksheet.range("S:U").EntireColumn.Delete(shift:=xlRight)

        'Gemmer Excel
        xlFilename = "DANF-KPI" & "-" & period & "-" & Format("#" & Date & "#", "yyyy")
        objWorkbook.SaveAs(fileName:=xlFilePath & xlFilename, FileFormat:=51)
        objWorkbook.ChangeFileAccess(XlFileAccess.xlReadOnly)

        'finder alle sendinger til IE og GB
        objWorksheet.range("A1:AY1").AutoFilter(Field:=10, Criteria1:=Array("IE", "GB"), Operator:=xlFilterValues)
        objWorksheet.range("A1:AY" & iLastXlRow).SpecialCells(xlCellTypeVisible).Copy

        objTmpWorkbook = objExcel.Workbooks.Add
        objTmpWorksheet = objTmpWorkbook.Sheets(1)
        objTmpWorksheet.range("A1").PasteSpecial

        'Gemmer Excel
        xlFilename = "IEGB" & "-" & period & "-" & Format("#" & Date & "#", "yyyy")
        objTmpWorkbook.SaveAs(fileName:=xlFilePath & xlFilename, FileFormat:=51)
        objTmpWorkbook.ChangeFileAccess(XlFileAccess.xlReadOnly)
        objTmpWorksheet.range("A1:AY" & iLastXlRow).Delete(shift:=xlUp)
        xlFilename = objTmpWorkbook.FullName

        'Finder modtager
        mailInfo = DanfossKPI_getDeptMail("IE", "GB")

        'Fjerner IE og GB sendinger
        If Trim(mailInfo(1)) <> "" Then objWorksheet.range("A2:AY" & iLastXlRow).Delete(shift:=xlUp)
        objWorksheet.ShowAllData

        'Sender Excel som mail
        Call ExcelMail.CreateEmail(mailInfo(1), userinfo(0), sBody, sSubject, mailAttachmentPath:=xlFilename) 'mailInfo(1)

        'Sletter Excel fil
        Kill(xlFilename)

        'finder alle sendinger til PL fra DK og DE (PL-IMP)
        objWorksheet.range("A1:AY1").AutoFilter(Field:=10, Criteria1:="PL")
        objWorksheet.range("A1:AY1").AutoFilter(Field:=5, Criteria1:=Array("DK", "DE"), Operator:=xlFilterValues)
        objWorksheet.range("A1:AY" & iLastXlRow).SpecialCells(xlCellTypeVisible).Copy

        objTmpWorksheet.range("A1").PasteSpecial

        'Gemmer Excel
        xlFilename = "PL-IMP" & "-" & period & "-" & Format("#" & Date & "#", "yyyy")
        objTmpWorkbook.SaveAs(fileName:=xlFilePath & xlFilename, FileFormat:=51)
        objTmpWorkbook.ChangeFileAccess(XlFileAccess.xlReadOnly)
        objTmpWorksheet.range("A1:AY" & iLastXlRow).Delete(shift:=xlUp)
        xlFilename = objTmpWorkbook.FullName

        'Finder modtager
        mailInfo = DanfossKPI_getDeptMail("", "PL")

        'Fjerner PL sendinger
        If Trim(mailInfo(1)) <> "" Then objWorksheet.range("A2:AY" & iLastXlRow).Delete(shift:=xlUp)
        objWorksheet.ShowAllData

        'Sender Excel som mail
        Call ExcelMail.CreateEmail(mailInfo(1), userinfo(0), sBody, sSubject, mailAttachmentPath:=xlFilename)

        'Sletter Excel fil
        Kill(xlFilename)

        'finder alle sendinger fra PL til DK og DE (PL-EXP)
        objWorksheet.range("A1:AY1").AutoFilter(Field:=5, Criteria1:="PL")
        objWorksheet.range("A1:AY1").AutoFilter(Field:=10, Criteria1:=Array("DK", "DE"), Operator:=xlFilterValues)
        objWorksheet.range("A1:AY" & iLastXlRow).SpecialCells(xlCellTypeVisible).Copy

        objTmpWorksheet.range("A1").PasteSpecial

        'Gemmer Excel
        xlFilename = "PL-EXP" & "-" & period & "-" & Format("#" & Date & "#", "yyyy")
        objTmpWorkbook.SaveAs(fileName:=xlFilePath & xlFilename, FileFormat:=51)
        objTmpWorkbook.ChangeFileAccess(XlFileAccess.xlReadOnly)
        objTmpWorksheet.range("A1:AY" & iLastXlRow).Delete(shift:=xlUp)
        xlFilename = objTmpWorkbook.FullName

        'Finder modtager
        mailInfo = DanfossKPI_getDeptMail("PL", "")

        'Fjerner PL sendinger
        If Trim(mailInfo(1)) <> "" Then objWorksheet.range("A2:AY" & iLastXlRow).Delete(shift:=xlUp)
        objWorksheet.ShowAllData

        'Sender Excel som mail
        Call ExcelMail.CreateEmail(mailInfo(1), userinfo(0), sBody, sSubject, mailAttachmentPath:=xlFilename)

        'Sletter Excel fil
        Kill(xlFilename)

        'finder alle sendinger fra DK til DK
        objWorksheet.range("A1:AY1").AutoFilter(Field:=5, Criteria1:="DK")
        objWorksheet.range("A1:AY1").AutoFilter(Field:=10, Criteria1:="DK", Operator:=xlFilterValues)
        objWorksheet.range("A1:AY" & iLastXlRow).SpecialCells(xlCellTypeVisible).Copy

        objTmpWorksheet.range("A1").PasteSpecial

        'Gemmer Excel
        xlFilename = "DK-DK" & "-" & period & "-" & Format("#" & Date & "#", "yyyy")
        objTmpWorkbook.SaveAs(fileName:=xlFilePath & xlFilename, FileFormat:=51)
        objTmpWorkbook.ChangeFileAccess(XlFileAccess.xlReadOnly)
        objTmpWorksheet.range("A1:AY" & iLastXlRow).Delete(shift:=xlUp)
        xlFilename = objTmpWorkbook.FullName

        'Finder modtager
        mailInfo = DanfossKPI_getDeptMail("DK", "DK")

        'Fjerner DK-DK sendinger
        If Trim(mailInfo(1)) <> "" Then objWorksheet.range("A2:AY" & iLastXlRow).Delete(shift:=xlUp)
        objWorksheet.ShowAllData

        'Sender Excel som mail
        Call ExcelMail.CreateEmail(mailInfo(1), userinfo(0), sBody, sSubject, mailAttachmentPath:=xlFilename)

        'Sletter Excel fil
        Kill(xlFilename)

        'finder alle Arvika sendinger
        objWorksheet.range("A1:AY1").AutoFilter(Field:=2, Criteria1:="ARVIKA", Operator:=xlFilterValues)
        objWorksheet.range("A1:AY" & iLastXlRow).SpecialCells(xlCellTypeVisible).Copy

        objTmpWorksheet.range("A1").PasteSpecial

        'Gemmer Excel
        xlFilename = "ARVIKA" & "-" & period & "-" & Format("#" & Date & "#", "yyyy")
        objTmpWorkbook.SaveAs(fileName:=xlFilePath & xlFilename, FileFormat:=51)
        objTmpWorkbook.ChangeFileAccess(XlFileAccess.xlReadOnly)
        objTmpWorksheet.range("A1:AY" & iLastXlRow).Delete(shift:=xlUp)
        xlFilename = objTmpWorkbook.FullName

        'Finder modtager
        mailInfo = DanfossKPI_getDeptMail("", "ARVIKA")

        'Fjerner alle ARVIKA sendinger
        If Trim(mailInfo(1)) <> "" Then objWorksheet.range("A2:AY" & iLastXlRow).Delete(shift:=xlUp)
        objWorksheet.ShowAllData

        'Sender Excel som mail
        Call ExcelMail.CreateEmail(mailInfo(1), userinfo(0), sBody, sSubject, mailAttachmentPath:=xlFilename)

        'Sletter Excel fil
        Kill(xlFilename)

        '*****************************************************************************************************
        'Sortere excel ark ud fra kolonne 'J'
        objWorksheet.range("A2:AY" & iLastXlRow).Sort(key1:=objWorksheet.range("J1:J" & iLastXlRow),
        order1:=xlAscending, Header:=xlNo)

        iXlrow = 2
        'Indlæser Destinationer
        Do Until iXlrow > iLastXlRow
            'Array Loop
            Call getArray(sDestArr, objWorksheet.range("J" & iXlrow).value, blDimensioned)

            iXlrow = iXlrow + 1

        Loop

        sDestArr = removeDuplicates(sDestArr)

        For Each Dest In sDestArr

            If Dest = "" Or Dest = Empty Then GoTo Break
            'finder alle Dest sendinger
            objWorksheet.range("A1:AY1").AutoFilter(Field:=10, Criteria1:=Dest, Operator:=xlFilterValues)
            objWorksheet.range("A1:AY" & iLastXlRow).SpecialCells(xlCellTypeVisible).Copy

            objTmpWorksheet.range("A1").PasteSpecial

            'Gemmer Excel
            xlFilename = Dest & "-" & period & "-" & Format("#" & Date & "#", "yyyy")
            objTmpWorkbook.SaveAs(fileName:=xlFilePath & xlFilename, FileFormat:=51)
            objTmpWorkbook.ChangeFileAccess(xlFilename.xlReadOnly)
            objTmpWorksheet.range("A1:AY" & iLastXlRow).Delete(shift:=xlUp)
            xlFilename = objTmpWorkbook.FullName

            'Finder modtager
            mailInfo = DanfossKPI_getDeptMail("", Dest)

            'Fjerner alle Dest sendinger
            If Trim(mailInfo(1)) <> "" Then objWorksheet.range("A2:AY" & iLastXlRow).Delete(shift:=xlUp)
            If objWorksheet.range("A2").value <> "" Or objWorksheet.range("A2").value <> Empty Then objWorksheet.ShowAllData

            'Sender Excel som mail
            Call ExcelMail.CreateEmail(mailInfo(1), userinfo(0), sBody, sSubject, mailAttachmentPath:=xlFilename)

            'Sletter Excel fil
            Kill(xlFilename)

        Next Dest
Break:

        'Vælger mappe hvor ExcelArk skal gemmes.
        objFSO = CreateObject("Scripting.FileSystemObject")
        objDialog = Excel.Application.FileDialog(msoFileDialogFolderPicker)
        objDialog.AllowMultiSelect = False
        objDialog.Title = "Select Folder To Save Workbook."
        objDialog.Show
        objfolder = objFSO.GetFolder(objDialog.SelectedItems(1))

        'Gemmer resten af Excel arket
        xlFilename = "DANF-KPI" & "-OBS-" & period & "-" & Format("#" & Date & "#", "yyyy")
        objWorkbook.SaveAs(fileName:=objfolder & "\" & xlFilename, FileFormat:=51)
        objWorkbook.ChangeFileAccess(XlFileAccess.xlReadOnly)

        'Sletter original excel fil
        Kill(oldXl)

        'Laver dogcen klar til næste gang
        iDays = 14 - Format(Now(), "d")
        vbResult = MsgBox("Skal der laves et DOCGEN?", vbYesNo, "Schedule a Night Docgen")
        'If vbResult = vbYes Then Call Danfoss_ScheduleDocgen(iDays)

        'End With
        '*****************************************************************************************************
        'Lukker Excel:
        'On Error GoTo ErrHandler
        objWorkbook.Close
        objTmpWorkbook.Close
        'objExcel.Visible = True
        'objExcel.DisplayAlerts = True
        'objExcel.Quit
        objWorksheet = Nothing
        objTmpWorksheet = Nothing
        objWorkbook = Nothing
        objTmpWorkbook = Nothing
        objExcel = Nothing
        MsgBox("NEXT RUN")
        Application.Quit
    End Sub
    Sub DanfossKPI_Part3()
        'By DAFE 2017-06
        'Retter Special Customers fra GB

        '*****************************************************************************************************
        strMakroNavn = "DanfossKPI_SpecialCustomers"
        '*****************************************************************************************************

        'Variables used when dealing with Cargolink
        Dim period As String : period = Format(DateAdd("m", -1, Now()), "yyyy/mm")
        period = Replace(period, "-", "/")

        'Variables used when dealing with Outlook
        Dim iMail As Integer, sMailTo As String, sMailFrom As String
        Dim sBody As String, sSubject As String

        userinfo = getUser()
        'Åbner CL fil til kopiering
        objExcel = Application
        objCLWorkbook = objExcel.ActiveWorkbook
        xlFilename = objCLWorkbook.FullName

        If InStr(xlFilename, "\\") = 0 And InStr(xlFilename, ":") = 0 Then
            MsgBox("åben og gem fil fra cargolink, og kør derefter script!", vbCritical, "Error!")
            Exit Sub
        End If


        objCLWorksheet = objCLWorkbook.Sheets(1)
        'Åbner Arbejds ark
        xlFilename = Excel.Application.GetOpenFilename(Title:="Åben Danfoss Arbejds Ark.")
        objWorkbook = objExcel.Workbooks.Open(xlFilename)
        objWorksheet = objWorkbook.Sheets("Data")
        objWorksheet.Activate
        objTmpWorksheet = objWorkbook.Sheets("Special Customers")
        objExcel.Visible = True
        objExcel.DisplayAlerts = False
        objCLWorksheet.Activate
        xlFilePath = Replace(xlFilename, objWorkbook.Name, "")
        '*****************************************************************************************************
        'Indsætter data fra cl
        '*****************************************************************************************************

        'Finder sidste linje i Excel ark
        rngStart = objCLWorksheet.range("A1")
        iLastXlRow = rngStart.CurrentRegion.Rows.Count

        '***INDSÆT KODE***
        'Tjek Excel-ark for blanke celler (Speciel fokus på AK+AL)

        objCLWorksheet.range("A2:AY" & iLastXlRow).Copy
        objWorksheet.Activate
        rngStart = objWorksheet.range("A1")
        iLastRow = rngStart.CurrentRegion.Rows.Count

        objWorksheet.range("A" & iLastRow).PasteSpecial
        iXlrow = iLastRow
        rngStart = objWorksheet.range("A1")
        iLastXlRow = rngStart.CurrentRegion.Rows.Count

        'Kopiere formler part 1
        objWorksheet.range("V2:AE2").Copy
        'formler fra V - AE (Transittid)
        objWorksheet.range("V" & iXlrow & ":AE" & iXlrow).PasteSpecial(xlPasteFormulas)
        objWorksheet.range("V" & iXlrow & ":AE" & iLastXlRow).FillDown
        objWorksheet.range("AZ2:BH2").Copy
        'formler fra AZ - BH (Kontrol)
        objWorksheet.range("AZ" & iXlrow & ":BH" & iXlrow).PasteSpecial(xlPasteFormulas)
        objWorksheet.range("AZ" & iXlrow & ":BH" & iLastXlRow).FillDown

        'Lukker CL Excel ark
        objCLWorkbook.Close
        objCLWorkbook = Nothing
        objCLWorksheet = Nothing

        '*****************************************************************************************************
        'Begynder at rette special customers
        '*****************************************************************************************************

        iXlrow = 2
        'Indlæser Destinationer
        iLastXlRow = objTmpWorksheet.Cells(objTmpWorksheet.Rows.Count, 1).End(xlUp).row

        iXlrow = 2
        'Indlæser Destinationer
        Do Until iXlrow > iLastXlRow
            If objTmpWorksheet.range("A" & iXlrow).value <> objTmpWorksheet.range("A" & iXlrow + 1).value Then
                'Array Loop
                Call getArray(vSpecial, objTmpWorksheet.range("A" & iXlrow).value, blDimensioned)
            End If
            iXlrow = iXlrow + 1

        Loop

        period = Format(DateAdd("m", -1, Now()), "yyyy/mm")
        period = Replace(period, "-", "/")
        'finder alle sendinger indenfor periode
        objWorksheet.range("A1:BH1").AutoFilter(Field:=36, Criteria1:=period, Operator:=xlFilterValues)

        'finder alle GB sendinger
        objWorksheet.range("A1:BH1").AutoFilter(Field:=14, Criteria1:="GB", Operator:=xlFilterValues)

        'finder alle Delay sendinger
        objWorksheet.range("A1:BH1").AutoFilter(Field:=27, Criteria1:="Delay", Operator:=xlFilterValues)

        iXlrow = 2
        'finder alle Special Customer sendinger
        objWorksheet.range("A1:BH1").AutoFilter(Field:=12, Criteria1:=Array(vSpecial), Operator:=xlFilterValues)

        rng = objWorksheet.UsedRange.SpecialCells(xlCellTypeVisible)
        For Each rngRow In rng.Rows
            If rngRow.row = 1 Then GoTo NextRngRow
            If Trim(objWorksheet.range("A" & rngRow.row).value) = "" Then Exit For
            objWorksheet.range("Y" & rngRow.row).value = objWorksheet.range("X" & rngRow.row).value
            iStartWeek = Format(objWorksheet.range("T" & rngRow.row).value, "ww")
            iEndWeek = Format(objWorksheet.range("U" & rngRow.row).value, "ww")
            LastDay = objWorksheet.range("U" & rngRow.row).value + 1
            'tilføjer ekstra tid hvis leveringsdato er lørdag eller søndag
            If Weekday(LastDay, vbMonday) = 6 Then iEndWeek = iEndWeek + 1
            If iStartWeek <> iEndWeek Then
                diff = (iEndWeek - iStartWeek) * 2
                objWorksheet.range("Z" & rngRow.row).value = objWorksheet.range("X" & rngRow.row).value - diff
                If Trim(objWorksheet.range("AF" & rngRow.row).value) = "" Then objWorksheet.range("AF" & rngRow.row).value = 33
            Else
                objWorksheet.range("Z" & rngRow.row).value = objWorksheet.range("X" & rngRow.row).value
                If Trim(objWorksheet.range("AF" & rngRow.row).value) = "" Then objWorksheet.range("AF" & rngRow.row).value = 33
            End If
NextRngRow:

        Next rngRow
        objWorksheet.ShowAllData
PL_Data:
        reply = MsgBox("Er begge raporter fra PL klar?", vbYesNo, "Import PL DATA.")
        If reply = vbNo Then GoTo SaveSheet

        'Indsætter PL(FI) data
        xlFilename = Excel.Application.GetOpenFilename(Title:="Åben FI fil fra Polen.")

        'Åbner CL fil til kopiering
        objCLWorkbook = objExcel.Workbooks.Open(xlFilename)
        objCLWorksheet = objCLWorkbook.Sheets(1)

        'Finder sidste linje i Excel ark
        rngStart = objCLWorksheet.range("A1")
        iLastXlRow = rngStart.CurrentRegion.Rows.Count

        'Ret Period
        objCLWorksheet.range("AJ2").value = period
        objCLWorksheet.range("AJ2:AJ" & iLastXlRow).FillDown

        'Ret ontime / delay text til On time eller Delay
        OntimeRng = objCLWorksheet.range("AA2:AA" & iLastXlRow)
        For Each cell In OntimeRng
            If LCase(cell.value) = "ontime" Then cell.value = "On time"
            If LCase(cell.value) = "delay" Then cell.value = "Delay"
        Next cell

        'Ret Cancelled shipments
        OntimeRng = objCLWorksheet.range("U2:U" & iLastXlRow)
        For Each cell In OntimeRng
            If LCase(cell.value) = "cancelled" Then
                cell.Copy
                cell.Offset(0, 13).PasteSpecial
                cell.value = cell.Offset(0, -1).value
            End If
        Next cell

        objCLWorksheet.range("A2:AY" & iLastXlRow).Copy

        rngStart = objWorksheet.range("A1")
        iLastXlRow = rngStart.CurrentRegion.Rows.Count

        objWorksheet.range("A" & iLastXlRow).PasteSpecial(xlPasteValues)

        'Lukker PL(FI) Excel ark
        objCLWorkbook.Close
        objCLWorkbook = Nothing
        objCLWorksheet = Nothing

        'Indsætter PL(DE) data
        xlFilename = Excel.Application.GetOpenFilename(Title:="Åben DE fil fra Polen.")

        'Åbner CL fil til kopiering
        objCLWorkbook = objExcel.Workbooks.Open(xlFilename)
        objCLWorksheet = objCLWorkbook.Sheets(1)

        'Finder sidste linje i Excel ark
        rngStart = objCLWorksheet.range("A1")
        iLastXlRow = rngStart.CurrentRegion.Rows.Count

        'Ret Period
        objCLWorksheet.range("AJ2").value = period
        objCLWorksheet.range("AJ2:AJ" & iLastXlRow).FillDown

        'Ret ontime / delay text til On time eller Delay
        OntimeRng = objCLWorksheet.range("AA2:AA" & iLastXlRow)
        For Each cell In OntimeRng
            If LCase(cell.value) = "ontime" Then cell.value = "On time"
            If LCase(cell.value) = "delay" Then cell.value = "Delay"
        Next cell

        'Ret Cancelled shipments
        OntimeRng = objCLWorksheet.range("U2:U" & iLastXlRow)
        For Each cell In OntimeRng
            If LCase(cell.value) = "cancelled" Then
                cell.Copy
                cell.Offset(0, 13).PasteSpecial
                cell.value = cell.Offset(0, -1).value
            End If
        Next cell

        objCLWorksheet.range("A2:AY" & iLastXlRow).Copy

        rngStart = objWorksheet.range("A1")
        iLastXlRow = rngStart.CurrentRegion.Rows.Count

        objWorksheet.range("A" & iLastXlRow).PasteSpecial(xlPasteValues)

        'Lukker PL(DE) Excel ark
        objCLWorkbook.Close
        objCLWorkbook = Nothing
        objCLWorksheet = Nothing

        'Kopiere formler part 3
        rngStart = objWorksheet.range("A1")
        iLastXlRow = rngStart.CurrentRegion.Rows.Count

        objWorksheet.range("AZ2:BH" & iLastXlRow).FillDown

        'Finder sidste linje i Excel ark
        rngStart = objWorksheet.range("A1")
        iLastXlRow = rngStart.CurrentRegion.Rows.Count

        'Fjerner ikke nødvendige reference numre i colums C
        objWorksheet.Columns("C").Replace(What:="/************************",
                        Replacement:="", LookAt:=xlPart,
                        SearchOrder:=xlByRows, MatchCase:=False,
                        SearchFormat:=False, ReplaceFormat:=False)

        'finder alle sendinger med formelfejl
        objWorksheet.range("A1:BH1").AutoFilter(Field:=27, Criteria1:="Transittime?", Operator:=xlFilterValues)



SaveSheet:

        'Vælger mappe hvor ExcelArk skal gemmes.
        objFSO = CreateObject("Scripting.FileSystemObject")
        objDialog = Excel.Application.FileDialog(msoFileDialogFolderPicker)
        objDialog.AllowMultiSelect = False
        objDialog.Title = "Select Folder To Save Workbook: 'Arbejdsfil Med Formler'."
        objDialog.Show
        objfolder = objFSO.GetFolder(objDialog.SelectedItems(1))

        'Gemmer Excel-ark
        objWorkbook.SaveAs(fileName:=objfolder & "\" & WorksheetFunction.text(Date - Day(Of Date), "[$-409]mmm") & " " & Format(Now(), "yyyy"))
        MsgBox("Ret alle formelfejl i kolonne 'AA + BD', og opdater tabellerne" & vbNewLine &
            "Formelfejl: transittime?, N/A#, REF#" & vbNewLine &
            "Send besked til Ole R. Thomsen, om at KPI'en er klar.", vbInformation, "Delay on SpecialCustomers is Done.")
        objWorkbook.Save


        objWorkbook.Close
        objWorkbook = Nothing
        objWorksheet = Nothing
        objExcel = Nothing
        Application.Quit

    End Sub
    Sub DanfossKPI_Part4()


        'Variables used when dealing with Edge
        Dim lSuccess As Long
        Dim PartOneURL As String : PartOneURL = "https://ecc.dsv.com/intro.php?ln="
        Dim PartTwoURL As String : PartTwoURL = "&sn=$2y$13$JgepqNK3csKW84Ji5rE33.cj8TX.qBxqHuV/nssAUlotLV3TDz"
        Dim url As String

        userinfo = getUser()
        url = PartOneURL & Mid(userinfo(0), 1, InStr(1, userinfo(0), "@") - 1) & PartTwoURL
        'Åbner Danfoss fil til kopiering
        objExcel = Excel.Application
        objWorkbook = objExcel.ActiveWorkbook
        objWorksheet = objWorkbook.Sheets("Data")
        objWorksheetComments = objWorkbook.Sheets("Comment")
        xlFilename = objWorkbook.FullName
        xlFilePath = Replace(xlFilename, objWorkbook.Name, "")

        If InStr(xlFilename, "\\") = 0 And InStr(xlFilename, ":") = 0 Then
            MsgBox("åben og gem fil fra cargolink, og kør derefter script!", vbCritical, "Error!")
            Exit Sub
        End If

        objWorksheet.ShowAllData

        'Åbner report der skal Uploades
        'On Error GoTo ErrHandler
        xlFilename = Excel.Application.GetOpenFilename(Title:="Åben seneste KPI Report der er blevet Uploadet.")

        objFinalWorkbook = objExcel.Workbooks.Open(xlFilename)
        objFinalWorksheet = objFinalWorkbook.Sheets("Data")
        objFinalWorksheetComments = objWorkbook.Sheets("Comment")

        xlFinalFilePath = Replace(xlFilename, objFinalWorkbook.Name, "")

        'Finder sidste linje i Excel ark
        rngStart = objWorksheet.range("A1")
        iLastXlRow = rngStart.CurrentRegion.Rows.Count

        period = Format(DateAdd("m", -1, Now()), "yyyy/mm")
        period = Replace(period, "-", "/")

        'finder alle sendinger indenfor periode
        objWorksheet.range("A1:BH1").AutoFilter(Field:=36, Criteria1:=period, Operator:=xlFilterValues)

        rngStart = objWorksheet.range("A2")
        iLastXlRow = rngStart.CurrentRegion.Rows.Count

        'Kopiere Data
        objWorksheet.range("A2:BH" & iLastXlRow).Copy

        'Finder sidste linje i Excel ark
        rngStart = objFinalWorksheet.range("A1")
        iLastXlRow = rngStart.CurrentRegion.Rows.Count

        'Indsætter/overskriver data
        objFinalWorksheet.range("A" & iLastXlRow).PasteSpecial(xlPasteValues)

        'Finder sidste linje i Excel ark
        rngStart = objFinalWorksheet.range("A1")
        iLastXlRow = rngStart.CurrentRegion.Rows.Count

        'Kopiere Performance Target ned
        objFinalWorksheet.range("BI2").AutoFill(objFinalWorksheet.range("BI2:BI" & iLastXlRow))

        'Indsætter/overskriver Kommentare fra Ole R.
        rngStart = objWorksheetComments.range("A1")
        iLastXlRow = rngStart.CurrentRegion.Rows.Count
        objWorksheetComments.range("A1:E" & iLastXlRow).Copy

        rngStart = objFinalWorksheetComments.range("A1")
        iLastXlRow = rngStart.CurrentRegion.Rows.Count
        objFinalWorksheetComments.range("A1").PasteSpecial(xlFilename)

        'Opdatere Pivot
        ' On Error GoTo ErrHandler
        'Set Variables Equal to Data Sheet and Pivot Sheet
        objFinalWorksheetPivot1 = objWorkbook.Sheets("Performance (Pivot)")
        objFinalWorksheetPivot2 = objWorkbook.Sheets("KPI (Pivot)")

        'Enter in Pivot Table Name
        PivotName = "PivotTable1"

        'Dynamically Retrieve Range Address of Data
        StartPoint = objFinalWorksheet.range("A1")
        iLastXlRow = StartPoint.CurrentRegion.Rows.Count
        EndPoint = objFinalWorksheet.range("BI" & iLastXlRow)
        DataRange = objFinalWorksheet.range(StartPoint, EndPoint)

        'Change Pivot Table Data Source Range Address
        NewRange = objFinalWorksheet.Name & "!" & DataRange.Address(ReferenceStyle:=xlR1C1)

        'Make sure every column in data set has a heading and is not blank (error prevention)
        If WorksheetFunction.CountBlank(DataRange.Rows(1)) > 0 Then
            MsgBox("One of your data columns has a blank heading." & vbNewLine & "Please fix and re-run!.", vbCritical, "Column Heading Missing!")
            Exit Sub
        End If

        'Update Pivot Table Data Source cache
        Call ChangeCaches(CStr(NewRange))

        'Vælger mappe hvor ExcelArk skal gemmes.
        objFSO = CreateObject("Scripting.FileSystemObject")
        objDialog = Excel.Application.FileDialog(msoFileDialogFolderPicker)
        objDialog.AllowMultiSelect = False
        objDialog.Title = ("Select Folder To Save Workbook: 'Arbejdsfil Med Formler'.")
        objDialog.Show
        objfolder = objFSO.GetFolder(objDialog.SelectedItems(1))

        'Gemmer Excel
        xlFilename = objWorkbook.Name
        objWorkbook.SaveAs(fileName:=objfolder & "\" & xlFilename, FileFormat:=51)
        objWorkbook.ChangeFileAccess(XlFileAccess.xlReadOnly)

        objWorkbook.Close
        objWorkbook = Nothing
        objWorksheet = Nothing

        'Vælger mappe hvor ExcelArk skal gemmes.
        objFSO = CreateObject("Scripting.FileSystemObject")
        objDialog = Excel.Application.FileDialog(msoFileDialogFolderPicker)
        objDialog.AllowMultiSelect = False
        objDialog.Title = "Select Folder To Save Workbook: 'Performance Report'."
        objDialog.Show
        objfolder = objFSO.GetFolder(objDialog.SelectedItems(1))

        'Gemmer Excel
        xlFilename = "Performance Report - " & UCase(WorksheetFunction.text(Date - Day(Of Date), "[$-409]mmmm"))
        objFinalWorkbook.SaveAs(fileName:=objfolder & "\" & xlFilename, FileFormat:=51)
        objFinalWorkbook.ChangeFileAccess(XlFileAccess.xlReadOnly)

        objFinalWorkbook.Close
        objFinalWorkbook = Nothing
        objFinalWorksheet = Nothing
        objExcel = Nothing

        lSuccess = ShellExecute(0, "Open", "https://ecc.dsv.com/intro.php") ' url)
        MsgBox("Upload " & xlFilename & " to Danfoss webportal.", vbOKOnly, "DANFOSS KPI READY.")

        'Mail to Danfoss Performance Group (+ TKA & OLT)
        sMailTo = "Anbo@danfoss.com;Bartosz.Lerch@danfoss.com;c.andersen@danfoss.com;drives.shipping@danfoss.com;" &
            "JIdzkowski@Sauer-Danfoss.com;Jaroslaw_Cynkier@danfoss.com;mlindemuth@sauer-danfoss.com;MHO@danfoss.com;" &
            "Michal_Miklaszewski@danfoss.com;mifl@danfoss.com;pjorgensen@danfoss.com;pf@danfoss.com;roman.miklovic@sk.dsv.com;" &
            "Thomas_Lyck@danfoss.com;TBruhnJensen@sauer-danfoss.com;thomas.kamper@dk.dsv.com;ole.r.thomsen@dk.dsv.com"

        sMailFrom = userinfo(0)
        sBody = "<p style=""margin: 0px 0px 16px;""><span style=""margin: 0px; color: black; font-family: 'Arial',sans-serif; font-size: 10pt;"">" &
            "Hi<br/><br/>Danfoss Performance report, have been uploaded to DSV online portal.<br/><br/>Med venlig hilsen, / Best regards,<br/><br/>" &
            userinfo(1) & "<br/>IT Consultant<br/>Key Account<p/>"

        sSubject = "Danfoss KPI " & period & Format("#" & Date & "#", "yyyy")

        Call CreateEmail(sMailTo, sMailFrom, sBody, sSubject)

        Application.Quit

    End Sub
    Sub DanfossKPI_Part5()

        'Åbner Danfoss fil til kopiering
        objExcel = Excel.Application
        objWorkbook = objExcel.ActiveWorkbook
        objWorksheet = objWorkbook.Sheets("Data")
        xlFilename = objWorkbook.FullName
        xlFilePath = Replace(xlFilename, objWorkbook.Name, "")

        'test for korrekt fil-sti
        If InStr(xlFilename, "\\") = 0 And InStr(xlFilename, ":") = 0 Then
            MsgBox("åben og gem fil fra cargolink, og kør derefter script!", vbCritical, "Error!")
            Exit Sub
        End If

        'Finder sidste linje i Excel ark
        rngStart = objWorksheet.range("A1")
        iLastXlRow = rngStart.CurrentRegion.Rows.Count

        period = Format(DateAdd("m", -1, Now()), "yyyy/mm")
        period = Replace(period, "-", "/")

        'finder alle sendinger indenfor periode
        objWorksheet.range("A1:BH1").AutoFilter(Field:=36, Criteria1:=period, Operator:=xlFilterValues)

        rngStart = objWorksheet.range("A1")
        iLastXlRow = rngStart.CurrentRegion.Rows.Count

        'Kopiere Data
        objWorksheet.range("A1:BI" & iLastXlRow).SpecialCells(xlCellTypeVisible).Copy

        'Opretter ny Workbook
        objNewWorkbook = objExcel.Workbooks.Add
        objNewWorksheet = objNewWorkbook.Sheets(1)

        objNewWorksheet.range("A1").PasteSpecial(xlPasteValuesAndNumberFormats)

        objWorkbook.Close
        objWorkbook = Nothing
        objWorksheet = Nothing

        rngStart = objNewWorksheet.range("A1")
        iLastXlRow = rngStart.CurrentRegion.Rows.Count

        'finder og sletter alle NL Linehaul sendinger.
        objNewWorksheet.range("A1:BH1").AutoFilter(Field:=4, Criteria1:="0", Operator:=xlFilterValues)
        objNewWorksheet.range("A2:BI" & iLastXlRow).SpecialCells(xlCellTypeVisible).Delete(shift:=xlUp)
        objNewWorksheet.ShowAllData

        rngStart = objNewWorksheet.range("A1")
        iLastXlRow = rngStart.CurrentRegion.Rows.Count

        On Error Resume Next
        'finder og sletter alle DSV Sea/Air sendinger.
        cCol = objWorksheet.Columns("C").Find(
        What:="2786996", LookAt:=xlPart,
            SearchOrder:=xlByRows, MatchCase:=False,
                SearchFormat:=False)

        If Not cCol Is Nothing Or IsEmpty(cCol) = False Then
            objNewWorksheet.range("A1:BH1").AutoFilter(Field:=5, Criteria1:="2786996", Operator:=xlFilterValues)
            objNewWorksheet.range("A2:BI" & iLastXlRow).SpecialCells(xlCellTypeVisible).Delete(shift:=xlUp)
            objNewWorksheet.ShowAllData
        End If
        On Error GoTo - 1

        'Vælger mappe hvor ExcelArk skal gemmes.
        objFSO = CreateObject("Scripting.FileSystemObject")
        objDialog = Excel.Application.FileDialog(msoFileDialogFolderPicker)
        objDialog.AllowMultiSelect = False
        objDialog.Title = "Select Folder To Save Workbook: 'Datafile sendt til Danfoss'."
        objDialog.Show
        objfolder = objFSO.GetFolder(objDialog.SelectedItems(1))

        'Gemmer Excel
        xlFilename = "DSV " & Format(Now(), "yyyy") & " " & UCase(WorksheetFunction.text(Date - Day(Of Date), "[$-409]mmm"))
        objNewWorkbook.SaveAs(fileName:=objfolder & "\" & xlFilename, FileFormat:=51)
        objNewWorkbook.ChangeFileAccess(XlFileAccess.xlReadOnly)

        objNewWorkbook.Close
        objNewWorkbook = Nothing
        objNewWorksheet = Nothing
        objExcel = Nothing

        'Mail to Danfoss
        userinfo = getUser()
        sMailFrom = userinfo(0)
        sBody = "<p style=""margin: 0px 0px 16px;""><span style=""margin: 0px; color: black; font-family: 'Arial',sans-serif; font-size: 10pt;"">" &
            "Hi<br/><br/>Please see enclosed file for Danfoss Performance report for " & UCase(WorksheetFunction.text(Date - Day(Of Date), "[$-409]mmm")) &
            ".<br/><br/>Med venlig hilsen, / Best regards,<br/><br/>" &
            userinfo(1) & "<br/>IT Consultant<br/>Key Account<p/>"

        sSubject = "Danfoss KPI " & UCase(WorksheetFunction.text(Date - Day(Of Date), "[$-409]mmm"))

        Call CreateEmail("GlobalPMS@danfoss.com", sMailFrom, sBody, sSubject, mailAttachmentPath:=objNewWorkbook.FullName)

        Application.Quit

    End Sub
    Private Function DanfossKPI_getDeptMail(countryFrom As Object, countryTo As Object) As Object

        '*****************************************************************************************************
        strMakroNavn = "DanfossKPI_getDeptMail"
        '*****************************************************************************************************
        Dim lane As String
        Dim vReturn(2) As String, bMail As Boolean

        If Trim(countryFrom) = "" Then countryFrom = "**"
        If Trim(countryTo) = "" Then countryTo = "**"
        lane = countryFrom & countryTo
        If Len(countryTo) > 3 Then lane = countryTo
        Dim Count As Long : Count = 0

        Const adOpenStatic = 3
        Const adLockOptimistic = 3
        Const adCmdText = &H1

        objconnection = CreateObject("ADODB.Connection")
        objRecordset = CreateObject("ADODB.Recordset")

        objconnection.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=H:\DAFE_Kunder\Danfoss\KPIMacro\Mail.xlsx;Extended Properties=""Excel 12.0 Xml;HDR=YES"";")

        objRecordset.Open("Select * FROM [Sheet1$]", objconnection, adOpenStatic, adLockOptimistic, adCmdText)

        Do Until objRecordset.EOF
            If lane = objRecordset.Fields.item("Kunde/by/Lane") Then
                vReturn(0) = objRecordset.Fields.item("Kontakt")
                vReturn(1) = objRecordset.Fields.item("e-mail")
                Exit Do
            ElseIf objRecordset.Fields.item("e-mail") = "Null" Then
                Exit Do
            End If
            objRecordset.MoveNext
        Loop

        objRecordset.Close
        objconnection.Close


        DanfossKPI_getDeptMail = vReturn

    End Function
    Private Function removeDuplicates(ByVal myArray As Object) As Object

        '*****************************************************************************************************
        strMakroNavn = "removeDuplicates"
        '*****************************************************************************************************

        Dim d As Object
        Dim v As Object 'Value for function
        Dim outputArray() As Object
        Dim i As Integer

        d = CreateObject("Scripting.Dictionary")

        For i = LBound(myArray) To UBound(myArray)

            d(myArray(i)) = 1

        Next i

        i = 0
        For Each v In d.Keys()

            ReDim Preserve outputArray(0 To i)
            outputArray(i) = v
            i = i + 1

        Next v

        removeDuplicates = outputArray

    End Function
    Private Function getUser() As String()
        Dim sInitial As String, sName As String, sPhone As String, sMail As String, sMobile As String
        Dim vReturn(2) As String
        Dim OutApp As Object
        OutApp = CreateObject("Outlook.Application")
        sMail = OutApp.Session.accounts.item(1).smtpaddress
        sName = OutApp.Session.accounts.item(1).userName
        vReturn(0) = sMail
        vReturn(1) = uCaseFirstLetter(sName)
        getUser = vReturn
    End Function
    Private Function uCaseFirstLetter(str)
        Dim arr, i
        arr = Split(str, ".")
        For i = LBound(arr) To UBound(arr)
            arr(i) = UCase(Left(arr(i), 1)) & Mid(arr(i), 2)
        Next
        uCaseFirstLetter = Join(arr, " ")
    End Function
    Private Function getArray(ByRef vArray As Object, value As Object, blDimensioned As Boolean)
        'by DAFE 2017/06

        If blDimensioned = True Then
            ReDim Preserve vArray(0 To UBound(vArray) + 1)
        Else
            ReDim vArray(0 To 0)
            blDimensioned = True
        End If
        vArray(UBound(vArray)) = value

        getArray = vArray
    End Function
    Private Function ChangeCaches(sNewSource As String)
        ' sample to change multiple pivots based off the same Excel Range-based pivotcache
        Dim pc As PivotCache
        Dim ws As Worksheet
        Dim pt As PivotTable
        Dim bCreated As Boolean

        For Each ws In ActiveWorkbook.Worksheets
            For Each pt In ws.PivotTables
                If Not bCreated Then
                    ' this only adds a new cache on the first run through
                    ' on subsequent passes, the pivot tables are simply assigned to the new cache
                    ' if multiple caches are desired, simply repeat this part for each pivot table.
                    pt.ChangePivotCache(ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase,
                    SourceData:=sNewSource, Version:=xlPivotTableVersion14))
                    pc = pt.PivotCache
                    bCreated = True
                Else
                    If pt.CacheIndex <> pc.Index Then pt.CacheIndex = pc.Index
                End If
            Next pt
        Next ws
    End Function
    Private Function CreateEmail(sMailTo As String, sMailFrom As String, Optional mailBody As String = "", Optional sMailSubject As String = "",
        Optional BCC As String = "", Optional CC As String = "", Optional mailAttachmentPath As String = "", Optional sender As String = "")
        On Error GoTo ErrHandler
        Dim mailSubject As String : mailSubject = sMailSubject
        Dim mailTo As String : mailTo = sMailTo
        Dim mailBCC As String : mailBCC = BCC
        Dim mailCC As String : mailCC = CC
        Dim mailFrom As String : mailFrom = sMailFrom
        Dim replyto As String
        Dim iMsg As Object, iConf As Object, OlApp As Object
        Dim Flds As Object, iBp As Object, OlMail As Object

        'olMailItem is the Outlook Application's constant,
        Const olMailItem As Long = 0

        'Try using of already open Outlook Application
        On Error Resume Next
        OlApp = GetObject(, "Outlook.Application")

        ' If Outlook Application was not active then create it
        If err Then OlApp = CreateObject("Outlook.Application")

        OlMail = OlApp.CreateItem(olMailItem)

        iMsg = CreateObject("CDO.Message")
        iConf = CreateObject("CDO.Configuration")

        iConf.Load(-1)    ' CDO Source Defaults
        Flds = iConf.Fields

        With Flds
            .item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 2
            .item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "DK.Sha.Migatronic@dk.dsv.com"
            .item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "KamH1605"
            .item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
            .item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.dsv.com"
            .item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 587
            .Update
        End With

        sBody = mailBody & "<br/><br/><img src=""\\dsv.com\corp\department\keyaccounthor\IT\Outlook\DSV.gif"" alt=""DSV logo"">" &
            "<p style=""margin: 0px 0px 16px;""><span style=""margin: 0px; color: black; font-family: 'Arial',sans-serif; font-size: 10pt;""><br/>" &
            "</span><span style=""margin: 0px; color: black; font-family: 'Arial',sans-serif; font-size: 10pt;""> DSV Road A/S <br/>Nokiavej 30 <br/>DK-8700 Horsens </span><br/>&nbsp;<br/><br/><br/><p style=""margin: 0px;""><strong><u><span lang=""EN-US"" style=""margin: 0px; color: #002664; font-family: 'Arial',sans-serif; font-size: 7.5pt;"">" &
            "DSV Standard Terms and Conditions</span></u></strong><span lang=""EN-US"" style=""margin: 0px; color: #002664; font-family: 'Arial',sans-serif; font-size: 7.5pt;"">" &
            "<br /> All services are rendered according to the DSV Standard Terms and Conditions and the General Conditions of the Nordic Association of Freight Forwarders - NSAB2015." &
            "In case of contradictions between the DSV Standard Terms and Conditions and the NSAB2015, the NSAB2015 shall prevail. Your legal position is materially altered due to" &
            "DSV&rsquo;s limited liability in case of loss of, damage to or delay of your cargo. DSV will furthermore obtain the right of lien over your cargo and all claims against" &
            "DSV are time-barred after 1 year. <strong>We recommend that you review the full text of the DSV Standard Terms and Conditions and the NSAB2015 prior to DSV&rsquo;s pick-up" &
            "of your cargo - <a href=""http://dasp.dk/sites/dasp.dk/files/pictures/nsab_2015_uk.pdf""><span style=""margin: 0px; color: blue;"">NSAB2015</span></a> &ndash;" &
            "<a href=""http://www.e-pages.dk/dsv/600/""><span style=""margin: 0px; color: blue;"">DSV Standard Terms &amp;Conditions</span></a>.</strong><br /> Orders undertaken" &
            "as carrier of overseas carriage are subject to conditions stipulated in the DSV Ocean Transport Bill of Lading/Sea Waybill. Your legal position is materially altered due to" &
            "DSV 's limited liability in case of loss of, damage to or delay of your cargo. DSV will furthermore obtain the right of lien over your cargo and all claims against" &
            "DSV are time-barred after 9 months. <strong>We recommend that you review the full version of the DSV Ocean Bill of Lading before DSV's pick-up of your cargo -" &
            "<a href=""http://www.e-pages.dk/dsv/634/""><span style=""margin: 0px; color: blue;"">DSV Ocean Transport B/L</span></a></strong><br /> Orders undertaken as carrier of carriage" &
            "by air are subject to conditions stipulated in DSV's House Air waybill. Your legal position is materially altered due to DSV's limited liability in case of loss of, damage to or" &
            "delay of your cargo. All claims against DSV are time-barred after 2 years. <strong>We recommend that you review the full version of the DSV House Air waybill prior to DSV's pick-up" &
            "of your cargo - <a href=""http://www.e-pages.dk/dsv/633/""><span style=""margin: 0px; color: blue;"">DSV House Air Waybill</span></a></strong><br /> In case of discrepancy between the" &
            "DSV Standard Terms and Conditions and the terms stipulated in the DSV Ocean Transport B/L or the DSV House Air Waybill, the terms of the DSV Ocean Transport B/L or the DSV House" &
            "Air Waybill shall prevail.</span></p>"

        If mailAttachmentPath <> "" Then iBp = iMsg.AddAttachment(mailAttachmentPath)
        mailBCC = mailBCC & ";" & mailFrom
        If sender <> "" Then
            replyto = mailFrom
            mailFrom = """" & sender & """,<kam.support@dk.dsv.com>"
        End If
        '    On Error GoTo Omail
        '    If mailTo <> "" Then
        '        With iMsg
        '            Set .Configuration = iConf
        '            .To = mailTo
        '            .CC = mailCC
        '            .BCC = mailBCC
        '            .From = mailFrom
        '            .replyto = replyto
        '            .Subject = mailSubject
        '            .HTMLBody = mailBody
        '            '.Send
        '            .Display
        '        End With
        '    Else
        'Omail:
        '    On Error Resume Next
        If mailAttachmentPath <> "" Then OlMail.Attachments.Add(mailAttachmentPath)
        OlMail.To = mailTo
        OlMail.From = mailFrom
        OlMail.Subject = mailSubject
        OlMail.HTMLBody = sBody
        OlMail.Display
        On Error GoTo - 1
        '    End If

        iBp = Nothing
        iMsg = Nothing
        iConf = Nothing
        mailTo = ""
        mailBCC = ""
        mailFrom = ""
        mailSubject = ""
        sBody = ""

        Exit Function
ErrHandler:
        Call ErrorModule.ErrHandler

    End Function

End Class
