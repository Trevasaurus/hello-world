# hello-world
tutorial
_this is a test.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''    email     ''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub Email()
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

    Dim OutApp As Object
    Dim OutMail As Object
    Dim rng As Range
    Dim SigString As String
    Dim Signature As String
    Dim StrBody As String
    Dim data_file As String
    Dim file_date As String
    Dim folder_date As String
    Dim week_num As String
    Dim subject_date As String
    Dim secondary_contact As String
    
    secondary_contact = "xxxxxxxxxxxxx"
    
    
    email_date = Range("email_date")
    file_date = Range("CD_DateSave")
    
    total_pnl = Format(Range("total_pnl"), "Currency")
    clean_pnl = Format(Range("clean_pnl"), "Currency")
    activity_pnl = Format(Range("activity_pnl"), "Currency")
    flash_activity = Format(Range("flash_activity"), "Currency")
    flash_legacy = Format(Range("flash_legacy"), "Currency")
    flash_diff = Format(Range("flash_diff"), "Currency")
    flash_diff_adj = Format(Range("flash_diff_adj"), "Currency")
    residual_pnl = Format(Range("residual_pnl"), "Currency")
    
    Set rng = Range("email_table")

''''''Attachment File name & location
    data_file = "Secondary.Daily_" & file_date & "_FINAL.xlsx"
    Dim sPath As String: sPath = "\\nnjxcarlfppr02.svr.us.jpmchase.net\na_rfs_mb_share_prod$\MBNJ1Capital Markets\Secondary\PNL\PnL Daily Reports\"
    Dim sFilename1 As String: sFilename1 = sPath & data_file

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
       
    'create email body
    StrBody = "All,<br><br>" & _
            "Attached is the FINAL PnL for COB " & "<b>" & email_date & "</b>. " & _
            "Total Secondary PnL is " & "<b>" & total_pnl & "mm " & "</b>" & _
            "with " & clean_pnl & "mm of Clean PnL and " & _
            residual_pnl & "mm in residual. Today's PnL is primarily driven by <br><br>" & _
            "Flash to Actual PnL difference of " & flash_diff & "mm is driven by " & _
            "Activity PnL of " & flash_activity & "mm and Legacy Book PnL of " & flash_legacy & "mm. " & _
            "Adjusting for these items that are not included in the Flash PnL " & _
            "the remaining difference is " & flash_diff_adj & "mm. <br><br>" & _
            RangetoHTML(rng) & "<br><br>" & _
            "<br><br>" & _
            "Note that voting functionality is for head of Secondary Marketing or designee sign-off of daily P&L for control purposes. Other recipients of this report can ignore the voting functionality. <br><br>" & _
            "Please contact xxxxxxxxxxxxxxxxxxxxxx " & _
            "or " & xxxxxxxxxxxx & " email group with any questions or comments. <br><br>" & _
            "PLEASE DO NOT 'Reply All' TO THIS EMAIL <br><br>"
            
    If Dir(SigString) <> "" Then
        Signature = ""
    Else
        Signature = ""
    End If

        On Error Resume Next
        With OutMail
            .SentOnBehalfOfName = "mb.secondary.marketing.product.control@jpmchase.com"
            .To = "xxxxxxxxxxxxxxxx.com"
            .CC = ""
            .BCC = "xxxxxxxxxxxxxxxxxx.com"
            .Subject = "xxxxxxxxxxxxxxxx " & email_date
            .HTMLBody = StrBody
            .Attachments.Add sFilename1
            .Display
            .VotingOptions = "Approve P&L; Reject P&L"
        End With
        
    On Error GoTo 0

    Set OutMail = Nothing
    Set OutApp = Nothing
    
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
End Sub

Function RangetoHTML(rng As Range)

    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook

    TempFile = Environ$("temp") & "\" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"

    'Copy the range and create a new workbook to past the data in
    Range("email_table").Copy
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        .Cells(1).Select
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        On Error GoTo 0
    End With

    'Publish the sheet to a htm file
    With TempWB.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         Filename:=TempFile, _
         Sheet:=TempWB.Sheets(1).Name, _
         Source:=TempWB.Sheets(1).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With

    'Read all data from the htm file into RangetoHTML
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.ReadAll
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                          "align=left x:publishsource=")

    'Close TempWB
    TempWB.Close savechanges:=False

    'Delete the htm file we used in this function
    Kill TempFile

    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub Delete_EmptyRows()
'
'This macro will delete all rows that are missing data in a cell
'underneath and including the selected cell.
'Importnant: To avoid run time error, get an accurate row count for your sheet!
'
Dim Counter
Dim i As Integer

Counter = InputBox("Enter the total number of rows to process")

ActiveCell.Select

For i = 1 To Counter

    If ActiveCell = "" Then

        Selection.EntireRow.Delete
        Counter = Counter - 1

    Else

ActiveCell.Offset(1, 0).Select
    End If

Next i

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub WrapFormula()

Dim rng As Range
Dim cell As Range
Dim x As String

'Determine if a single cell or range is selected
  If Selection.Cells.Count = 1 Then
    Set rng = Selection
    If Not rng.HasFormula Then GoTo NoFormulas
  Else
    'Get Range of Cells that Only Contain Formulas
      On Error GoTo NoFormulas
        Set rng = Selection.SpecialCells(xlCellTypeFormulas)
      On Error GoTo 0
  End If

'Loop Through Each Cell in Range and add *run_rate
  For Each cell In rng.Cells
    x = cell.Formula
    'cell = "=IFERROR(" & Right(x, Len(x) - 1) & "," & Chr(34) & Chr(34) & ")"
    cell = x & "* run_rate"
  Next cell

Exit Sub

'Error Handler
NoFormulas:
  MsgBox "There were no formulas found in your selection!"

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''








