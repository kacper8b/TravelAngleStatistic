Option Explicit

Private Const LIST_START_ROW As Long = 10
Private Const SOURCE_LIST_COLUMN As String = "A"

Private Const PYTHON_COMMAND_CELL As String = "B2"
Private Const SCRIPT_PATH_CELL As String = "B3"
Private Const STATE_PATH_CELL As String = "B4"
Private Const GENERATED_WORKBOOK_CELL As String = "B5"
Private Const STATUS_CELL As String = "B6"

Public Sub StatisticsWorkbook_Open()
    On Error GoTo Handler
    MarkResultsUnverified
    RefreshSourcesButton
    Exit Sub

Handler:
    SetStatus "Open refresh failed: " & Err.Description, 1
End Sub

Public Sub StatisticsWorkbook_BeforeSave()
    On Error Resume Next
    MarkResultsUnverified
End Sub

Public Sub AddSourcesButton()
    Dim files As Collection
    Set files = PickCsvFiles(True)
    If files Is Nothing Then
        Exit Sub
    End If
    If files.Count = 0 Then
        Exit Sub
    End If
    RunCommandAndSync "add", files
End Sub

Public Sub ReplaceSourcesButton()
    Dim files As Collection
    Set files = PickCsvFiles(True)
    If files Is Nothing Then
        Exit Sub
    End If
    If files.Count = 0 Then
        Exit Sub
    End If
    RunCommandAndSync "replace", files
End Sub

Public Sub RemoveSelectedSourceButton()
    Dim sheet As Worksheet
    Dim selectedPath As String
    Dim files As Collection

    Set sheet = ThisWorkbook.Worksheets("Data sources")
    If ActiveSheet.Name <> sheet.Name Then
        MsgBox "Select a source path cell in the Data sources sheet first.", vbInformation
        Exit Sub
    End If

    If ActiveCell.Column <> sheet.Range(SOURCE_LIST_COLUMN & "1").Column Or ActiveCell.Row < LIST_START_ROW Then
        MsgBox "Select a source path cell from the active source list.", vbInformation
        Exit Sub
    End If

    selectedPath = Trim$(CStr(ActiveCell.Value))
    If Len(selectedPath) = 0 Then
        MsgBox "Selected cell is empty.", vbInformation
        Exit Sub
    End If

    Set files = New Collection
    files.Add selectedPath
    RunCommandAndSync "remove", files
End Sub

Public Sub RefreshSourcesButton()
    RunCommandAndSync "refresh"
End Sub

Private Sub RunCommandAndSync(ByVal actionName As String, Optional ByVal files As Collection = Nothing)
    Dim command As String
    Dim output As String
    Dim exitCode As Long
    Dim generatedWorkbook As String

    On Error GoTo Handler

    command = BuildCommand(actionName, files)
    exitCode = RunProcess(command, output)
    SetStatus output, exitCode

    generatedWorkbook = ResolvePath(CStr(ThisWorkbook.Worksheets("Data sources").Range(GENERATED_WORKBOOK_CELL).Value))
    If Len(generatedWorkbook) > 0 And Len(Dir$(generatedWorkbook)) > 0 Then
        SyncFromGeneratedWorkbook generatedWorkbook
    End If

    If exitCode <> 0 Then
        MsgBox "Command failed: " & actionName & vbCrLf & output, vbExclamation
    End If
    Exit Sub

Handler:
    SetStatus "VBA runtime error: " & Err.Description, 1
    MsgBox "VBA runtime error: " & Err.Description, vbExclamation
End Sub

Private Function BuildCommand(ByVal actionName As String, Optional ByVal files As Collection = Nothing) As String
    Dim sheet As Worksheet
    Dim pythonCommand As String
    Dim scriptPath As String
    Dim statePath As String
    Dim generatedWorkbook As String
    Dim command As String
    Dim item As Variant

    Set sheet = ThisWorkbook.Worksheets("Data sources")
    pythonCommand = Trim$(CStr(sheet.Range(PYTHON_COMMAND_CELL).Value))
    scriptPath = ResolvePath(CStr(sheet.Range(SCRIPT_PATH_CELL).Value))
    statePath = ResolvePath(CStr(sheet.Range(STATE_PATH_CELL).Value))
    generatedWorkbook = ResolvePath(CStr(sheet.Range(GENERATED_WORKBOOK_CELL).Value))

    If Len(pythonCommand) = 0 Then
        pythonCommand = "python"
    End If
    If Len(scriptPath) = 0 Then
        Err.Raise vbObjectError + 1, "BuildCommand", "Script path is empty."
    End If
    If Len(statePath) = 0 Then
        Err.Raise vbObjectError + 2, "BuildCommand", "State path is empty."
    End If
    If Len(generatedWorkbook) = 0 Then
        Err.Raise vbObjectError + 3, "BuildCommand", "Generated workbook path is empty."
    End If

    command = QuoteArg(pythonCommand) _
        & " " & QuoteArg(scriptPath) _
        & " --workbook " & QuoteArg(generatedWorkbook) _
        & " --state " & QuoteArg(statePath) _
        & " " & actionName

    If Not files Is Nothing Then
        For Each item In files
            command = command & " " & QuoteArg(CStr(item))
        Next item
    End If

    BuildCommand = command
End Function

Private Function RunProcess(ByVal command As String, ByRef output As String) As Long
    Dim shell As Object
    Dim exec As Object
    Dim wrapped As String

    Set shell = CreateObject("WScript.Shell")
    wrapped = "cmd.exe /c " & QuoteArg(command)
    Set exec = shell.Exec(wrapped)

    Do While exec.Status = 0
        DoEvents
    Loop

    output = exec.StdOut.ReadAll
    output = output & exec.StdErr.ReadAll
    RunProcess = CLng(exec.ExitCode)
End Function

Private Function PickCsvFiles(ByVal allowMulti As Boolean) As Collection
    Dim dialog As FileDialog
    Dim selected As Collection
    Dim index As Long

    Set dialog = Application.FileDialog(msoFileDialogFilePicker)
    dialog.AllowMultiSelect = allowMulti
    dialog.Title = "Select CSV source files"
    dialog.Filters.Clear
    dialog.Filters.Add "CSV files", "*.csv"

    If dialog.Show <> -1 Then
        Set PickCsvFiles = New Collection
        Exit Function
    End If

    Set selected = New Collection
    For index = 1 To dialog.SelectedItems.Count
        selected.Add CStr(dialog.SelectedItems(index))
    Next index
    Set PickCsvFiles = selected
End Function

Private Function QuoteArg(ByVal value As String) As String
    QuoteArg = """" & Replace(value, """", """""") & """"
End Function

Private Function ResolvePath(ByVal value As String) As String
    Dim textValue As String
    textValue = Trim$(value)
    If Len(textValue) = 0 Then
        ResolvePath = ""
        Exit Function
    End If

    If InStr(textValue, ":") > 0 Or Left$(textValue, 2) = "\\" Then
        ResolvePath = textValue
        Exit Function
    End If

    ResolvePath = ThisWorkbook.Path & "\" & textValue
End Function

Private Sub SetStatus(ByVal message As String, ByVal exitCode As Long)
    Dim sheet As Worksheet
    Dim statusPrefix As String
    Set sheet = ThisWorkbook.Worksheets("Data sources")
    If exitCode = 0 Then
        statusPrefix = "OK"
    Else
        statusPrefix = "ERROR"
    End If
    sheet.Range(STATUS_CELL).Value = statusPrefix & " [" & Format$(Now, "yyyy-mm-dd hh:nn:ss") & "] " & message
End Sub

Private Sub SyncFromGeneratedWorkbook(ByVal generatedWorkbook As String)
    Dim sourceWorkbook As Workbook
    Dim dataSheet As Worksheet
    Dim outputList As String
    Dim lineItems() As String
    Dim index As Long
    Dim writeRow As Long

    Application.ScreenUpdating = False
    Set sourceWorkbook = Application.Workbooks.Open(generatedWorkbook, ReadOnly:=True)

    CopyWorksheetData sourceWorkbook.Worksheets("Movement angle"), ThisWorkbook.Worksheets("Movement angle")
    CopyWorksheetData sourceWorkbook.Worksheets("Tilt counts"), ThisWorkbook.Worksheets("Tilt counts")
    CopyWorksheetData sourceWorkbook.Worksheets("Final position"), ThisWorkbook.Worksheets("Final position")
    ClearResultsUnverified

    Set dataSheet = ThisWorkbook.Worksheets("Data sources")
    dataSheet.Range("A7").Value = sourceWorkbook.Worksheets("Data sources").Range("A2").Value
    dataSheet.Range("A8").Value = "Last successful refresh"
    dataSheet.Range("B8").Value = sourceWorkbook.Worksheets("Data sources").Range("B3").Value
    dataSheet.Range("A9").Value = sourceWorkbook.Worksheets("Data sources").Range("A5").Value
    dataSheet.Range("A10:A100000").ClearContents

    outputList = GetSourcesFromCommand()
    lineItems = Split(outputList, vbCrLf)
    writeRow = LIST_START_ROW
    For index = LBound(lineItems) To UBound(lineItems)
        If Left$(Trim$(lineItems(index)), 1) = "-" Then
            dataSheet.Cells(writeRow, SOURCE_LIST_COLUMN).Value = Trim$(Mid$(Trim$(lineItems(index)), 2))
            writeRow = writeRow + 1
        End If
    Next index
    dataSheet.Columns(SOURCE_LIST_COLUMN).EntireColumn.AutoFit

    sourceWorkbook.Close SaveChanges:=False
    Application.ScreenUpdating = True
End Sub

Private Function GetSourcesFromCommand() As String
    Dim command As String
    Dim output As String
    Dim exitCode As Long

    command = BuildCommand("list-sources")
    exitCode = RunProcess(command, output)
    If exitCode <> 0 Then
        GetSourcesFromCommand = ""
        Exit Function
    End If
    GetSourcesFromCommand = output
End Function

Private Sub CopyWorksheetData(ByVal sourceSheet As Worksheet, ByVal targetSheet As Worksheet)
    targetSheet.Cells.Clear
    sourceSheet.UsedRange.Copy Destination:=targetSheet.Range("A1")
End Sub

Private Sub MarkResultsUnverified()
    MarkSheetUnverified ThisWorkbook.Worksheets("Movement angle")
    MarkSheetUnverified ThisWorkbook.Worksheets("Tilt counts")
    MarkSheetUnverified ThisWorkbook.Worksheets("Final position")
End Sub

Private Sub MarkSheetUnverified(ByVal sheet As Worksheet)
    sheet.Range("A1").Value = "UNVERIFIED — enable macros and click Refresh to validate current results."
    sheet.Range("A1").Font.Color = RGB(156, 0, 6)
    sheet.Range("A1").Font.Bold = True
    sheet.Range("A1").Interior.Color = RGB(255, 199, 206)
End Sub

Private Sub ClearResultsUnverified()
    ClearSheetUnverified ThisWorkbook.Worksheets("Movement angle")
    ClearSheetUnverified ThisWorkbook.Worksheets("Tilt counts")
    ClearSheetUnverified ThisWorkbook.Worksheets("Final position")
End Sub

Private Sub ClearSheetUnverified(ByVal sheet As Worksheet)
    sheet.Range("A1").ClearContents
    sheet.Range("A1").Font.ColorIndex = xlColorIndexAutomatic
    sheet.Range("A1").Font.Bold = False
    sheet.Range("A1").Interior.Pattern = xlPatternNone
End Sub
