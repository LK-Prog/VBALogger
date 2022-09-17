VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vbLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'#===========================================#
'# Programm:    Logging functionality in VBA #
'# Author:      Lasse Koß                    #
'#===========================================#

Option Explicit

Private Const LOG_HEADER As String = _
    "LEVEL" & vbTab & _
    "|DATETIME" & vbTab & vbTab & vbTab & vbTab & _
    "|USER" & vbTab & vbTab & vbTab & _
    "|DETAILS"

' TODO: Place these values in a log config file
Private EMPTY_LINE_IN_LOG   As Boolean      ' Flag to get set if empty lines were found in a log file
Private LOG_LEVELS          As Variant      ' Array containing the levels for logging
Private PATH_SEP            As String       ' System path separator
Private LOG_FOLDER_PATH     As String       ' Path to the log folder
Private LOG_ARCHIVE_PATH    As String       ' Path to the log archive folder
Private LOG_FILE_PATH       As String       ' Absolute path to the log file
Private ROLL_CONFIG_FOLDER  As String       ' Path to the roll over configuration folder
Private ROLL_CONFIG_FILE    As String       ' Absolute path to the roll over configuration file
Private LOGGING_LEVEL       As LogLevel     ' Log everything from this level and above (default: WARN)

Public Enum LogLevel                        ' Enumeration of log levels corresponding to the array index
    TRACE = 0
    INFO = 1
    WARN = 2
    CRIT = 3
End Enum

Private Enum LogErrorNum                    ' Enumeration of custom error codes (added to vbObjectError)
    NOINIT = 100
    BAD_LOG_PATH = 101
    ROLL_INTERVAL = 102
    BAD_CONFIG_PATH = 103
    CONIFG_IDX = 104
    MSC_EXCEPTION = 999
End Enum

'------CONSTRUCTOR---------------------------------------------------------------------------------
Public Sub initLogger(Optional ByVal logFileFullName As String = "", _
Optional logFromLevel As LogLevel = WARN, Optional noRollOver As Boolean = False)

    LOG_LEVELS = Array("[TRACE]", "[INFO]", "[WARN]", "[CRIT]")
    PATH_SEP = Application.PathSeparator
    EMPTY_LINE_IN_LOG = False
    ROLL_CONFIG_FOLDER = ThisWorkbook.Path & PATH_SEP & "config"
    ROLL_CONFIG_FILE = ROLL_CONFIG_FOLDER & PATH_SEP & "roll_over_config.conf"
    LOGGING_LEVEL = logFromLevel

    Dim fileName As String, filePath As String
    Dim fso As Object, oFile As Object
           
    If Not logFileFullName = "" Then                                    ' Check if a path was supplied
        LOG_FOLDER_PATH = Left(logFileFullName, InStrRev(logFileFullName, PATH_SEP) - 1)
        LOG_FILE_PATH = logFileFullName
        
        If Not (InStrRev(LOG_FILE_PATH, ".txt") > 0) And Not (InStrRev(LOG_FILE_PATH, ".log") > 0) Then
            LOG_FILE_PATH = LOG_FILE_PATH & ".log"
        End If
    Else
        LOG_FOLDER_PATH = CurDir & PATH_SEP & "Logging"                 ' Use current path
        LOG_FILE_PATH = LOG_FOLDER_PATH & PATH_SEP & "AppLog.log"
    End If
    
    LOG_ARCHIVE_PATH = LOG_FOLDER_PATH & PATH_SEP & "Log_Archive"
    
    If Len(Dir(LOG_FILE_PATH)) = 0 Then                                 ' Check if the supplied path is valid
        Set fso = CreateObject("Scripting.FileSystemObject")
        
        If Not fso.FolderExists(LOG_FOLDER_PATH) Then
            fso.createFolder LOG_FOLDER_PATH                            ' Create folder if not
        End If
        
        Set oFile = fso.createTextFile(LOG_FILE_PATH)
        oFile.WriteLine LOG_HEADER
        oFile.Close
        
        Set fso = Nothing
        Set oFile = Nothing
        
        Me.logText INFO, _
                   Mid(LOG_FILE_PATH, InStrRev(LOG_FILE_PATH, PATH_SEP) + 1) & " created", _
                   "vbLogger", True
    ElseIf Not noRollOver Then
        Call rollOver
    End If

End Sub

'------DEFAULT LOG LEVEL---------------------------------------------------------------------------
' Changes the log level from default (WARN) to the supplied level
' objLogger.logText logs everything from the level and higher
Public Sub setDefaultLogLevel(level As LogLevel)

    LOGGING_LEVEL = level

End Sub

'------ROLL OVER STRATEGY--------------------------------------------------------------------------
Public Sub initRollOver(ByVal intervalInDays As Integer, _
ByVal archive As Boolean, Optional ByVal startDate As Date = #1/1/1970#)

    Dim fso As Object, oFile As Object, ff As Double, configStr As String, logFileFullName As String
    
    logFileFullName = LOG_FILE_PATH
    
    If intervalInDays <= 0 Then
        raiseLoggerError ROLL_INTERVAL, "VBLogger.vbLogger"
    End If
    
    If Len(Dir(logFileFullName)) = 0 Then
        raiseLoggerError BAD_LOG_PATH, "VBALogger.vbLogger"
        Exit Sub
    End If
    
    If Not Mid(logFileFullName, InStrRev(logFileFullName, ".") + 1) = "txt" _
    And Not Mid(logFileFullName, InStrRev(logFileFullName, ".") + 1) = "log" Then
        raiseLoggerError BAD_LOG_PATH, "VBALogger.vbLogger"
        Exit Sub
    End If
    
    configStr = logFileFullName & ";" & intervalInDays & ";" & archive & ";" & startDate
    
    If Len(Dir(ROLL_CONFIG_FILE)) = 0 Then
        Set fso = CreateObject("Scripting.FileSystemObject")
        If Not fso.FolderExists(ROLL_CONFIG_FOLDER) Then
            fso.createFolder ROLL_CONFIG_FOLDER
            Me.logText INFO, "Created roll over configuration folder", "vbLogger", True
        End If
        
        Set oFile = fso.createTextFile(ROLL_CONFIG_FILE)
        oFile.WriteLine "#---Roll Over Configuration---------------------------------------"
        oFile.WriteLine configStr
        oFile.Close
        
        Me.logText INFO, "Created roll over configuration file", "vbLogger", True
                         
        Me.logText INFO, "Configured " & Mid(LOG_FILE_PATH, InStrRev(LOG_FILE_PATH, PATH_SEP) + 1) & _
                         " for roll over (Every " & intervalInDays & " day(s) starting " & startDate & ")", _
                         "vbLogger", True
    Else
        Dim arrLines As Variant, i As Integer, oStr As String, fileFound As Boolean
        arrLines = readConfigFile
        
        On Error GoTo JmpNextIdx
        For i = LBound(arrLines) To UBound(arrLines)
            If Split(arrLines(i), ";")(0) = logFileFullName Then
                If arrLines(i) = configStr Then
                    GoTo JmpNothingToUpdate
                End If
                
                arrLines(i) = configStr
                fileFound = True
                Exit For
            End If
JmpNextIdx:
        Next i
        On Error GoTo 0
        
        ff = FreeFile()
        If fileFound Then                                               ' Update if found else append
            oStr = Join(arrLines, vbNewLine)
            Open ROLL_CONFIG_FILE For Output As #ff                     ' Update the config file
                Print #ff, oStr
            Close #ff
            
            Me.logText INFO, "Updated roll over configuration for " & _
                             Mid(LOG_FILE_PATH, InStrRev(LOG_FILE_PATH, PATH_SEP) + 1), _
                             "vbLogger", True
        Else
            Open ROLL_CONFIG_FILE For Append As #ff                     ' Append config string to config file
                Print #ff, configStr
            Close #ff
            
            Me.logText INFO, "Configured " & Mid(LOG_FILE_PATH, InStrRev(LOG_FILE_PATH, PATH_SEP) + 1) & _
                             " for roll over (Every " & intervalInDays & " day(s) starting " & startDate & ")", _
                             "vbLogger", True
        End If
    End If

JmpNothingToUpdate:
    Set oFile = Nothing
    Set fso = Nothing

End Sub

Private Sub rollOver()

    Dim fso As Object, oFile As Object, ff As Double, confStr As String, arrConf() As Variant, i As Integer

    If Len(Dir(ROLL_CONFIG_FILE)) = 0 Then                      ' If the config file does not exist exit
        Exit Sub
    End If

    arrConf = readConfigFile()

    On Error GoTo JmpNextIdx
    For i = LBound(arrConf) To UBound(arrConf)
        If Split(arrConf(i), ";")(0) = LOG_FILE_PATH Then
            confStr = arrConf(i)
            Erase arrConf
            Exit For
        End If
JmpNextIdx:
    Next i
    On Error GoTo 0
    
    If (Not Not arrConf) = 0 Then   ' If the array is not initialized the call was successfull
        Erase arrConf
        GoTo JmpRollOverStrFound
    Else
        Me.logText INFO, "Roll over config file exists but no strategy was defined", "vbLogger", True
        Exit Sub
    End If

    Dim logFileFullName As String, intervalInDays As Integer, archive As Boolean, startDate As Date
    Dim diff As Integer, arrConfStr As Variant
    
JmpRollOverStrFound:
    arrConfStr = Split(confStr, ";")
    
    logFileFullName = arrConfStr(0)
    intervalInDays = CInt(arrConfStr(1))
    archive = CBool(arrConfStr(2))
    startDate = CDate(arrConfStr(3))

    Erase arrConfStr

    If Date < startDate Then
        Exit Sub
    End If
    
    On Error GoTo JmpMscException
    Set fso = CreateObject("Scripting.FileSystemObject")

    diff = DateDiff("d", fso.GetFile(logFileFullName).DateCreated(), #10/1/2022#, 0, 0)
    If diff < intervalInDays Then
        Exit Sub
    End If

    If archive Then
        If Not fso.FolderExists(LOG_ARCHIVE_PATH) Then
            fso.createFolder LOG_ARCHIVE_PATH
            Me.logText INFO, "Log file archive folder not found -> created folder", "vbLogger", True
        End If
        
        Me.logText INFO, _
                   "Archived on " & Format(Date, "dd.mm.yyyy") & " at " & Format(Now, "hh:mm:ss"), _
                   "vbLogger", True
        
        fso.MoveFile source:=logFileFullName, _
                     Destination:=LOG_ARCHIVE_PATH & PATH_SEP & _
                                  Format(Now, "yyyy-mm-dd hh-nn-ss") & "_" & _
                                  Mid(logFileFullName, InStrRev(logFileFullName, PATH_SEP) + 1)
    Else
        fso.DeleteFile logFileFullName
    End If

    Set oFile = fso.createTextFile(logFileFullName)
    oFile.WriteLine LOG_HEADER
    oFile.Close
    
    Me.logText INFO, _
               "Created '" & Mid(logFileFullName, InStrRev(logFileFullName, PATH_SEP) + 1) & _
               "' after roll over", _
               "vbLogger", True
    
    Set oFile = Nothing
    Set fso = Nothing
    
Exit Sub
JmpMscException:
    raiseLoggerError MSC_EXCEPTION, "VBLogger.vbLogger"
    
End Sub

Public Sub removeRollOverConfigForThisFile()

    If Len(Dir(ROLL_CONFIG_FILE)) = 0 Then                      ' If the config file does not exist exit
        Exit Sub
    End If
    
    Dim ff As Double, arrLines() As Variant, i As Integer, j As Integer, arrO() As Variant, oStr As String
    
    arrLines = readConfigFile()
    
    If (Not Not arrLines) = 0 Then
        Exit Sub
    End If

    On Error GoTo JmpConfIdxOutOfBounds
    For i = LBound(arrLines) To UBound(arrLines)
        If Not Split(arrLines(i), ";")(0) = LOG_FILE_PATH Then
            ReDim Preserve arrO(j)
            arrO(j) = arrLines(i)
            j = j + 1
        End If
    Next i
    On Error GoTo 0

    oStr = Join(arrO, vbNewLine)
    
    Erase arrLines, arrO
    
    ff = FreeFile()
    Open ROLL_CONFIG_FILE For Output As #ff
        Print #ff, oStr
    Close #ff

    Me.logText INFO, "Removed roll over configuration for " & _
                     Mid(LOG_FILE_PATH, InStrRev(LOG_FILE_PATH, PATH_SEP) + 1), _
                     "vbLogger", True
Exit Sub
JmpConfIdxOutOfBounds:
    raiseLoggerError CONIFG_IDX, "VBLogger.vbLogger"

End Sub

' Reads the contents of the roll over configuration file
' @returns Array containing each line of the configuration file
Private Function readConfigFile() As Variant()

    Dim ff As Double, arrLines() As Variant, i As Integer
    
    ff = FreeFile()
    Open ROLL_CONFIG_FILE For Input As #ff
        While Not EOF(ff)
            ReDim Preserve arrLines(i)
            Line Input #ff, arrLines(i)
            i = i + 1
        Wend
    Close #ff

    readConfigFile = arrLines

End Function

'------WRTING TO LOG FILE--------------------------------------------------------------------------
' Use force to force the logger to log the text, even if the level does not match the LOGGING_LEVEL
Public Function logText(ByVal level As LogLevel, ByVal s As String, _
Optional userName As String = "", Optional force As Boolean = False) _
As String

    If PATH_SEP = "" Then
        raiseLoggerError NOINIT, "VBALogger.vbLogger"
    End If

    If level < LOGGING_LEVEL And Not force Then
        Exit Function
    End If

    Dim oStr As String, ff As Double
    
    If userName = "" Then
        userName = Application.userName
    End If
 
    s = Replace(s, vbNewLine, "<br/>")                          ' Change line breaks to line brake markers
    s = Replace(s, vbTab, "<t/>")                               ' Change tabs to tab markers

    oStr = LOG_LEVELS(level) & vbTab & _
           "|" & Format(Now, "yyyy-mm-dd hh-nn-ss") & vbTab & _
           "|" & userName
           
    If Len(userName) < 12 Then
        oStr = oStr & vbTab & vbTab
    Else
        oStr = oStr & vbTab
    End If
    
    oStr = Trim(oStr & "|" & s)                                 ' Remove any leading or trailing spaces
    
    
    Dim prepareResult As Boolean
    prepareResult = prepareTextFileForAppend(LOG_FILE_PATH)     ' Check if there is a new line at eof
    
    If Not prepareResult Then
        raiseLoggerError BAD_LOG_PATH, "VBLogger.vbLogger"
        Exit Function
    End If
    
    ff = FreeFile()
    Open LOG_FILE_PATH For Append As #ff                        ' Append output string to log file
        Print #ff, oStr
    Close #ff

    logText = oStr                                              ' Return ouput String

End Function

'------PARSING LOG FILE AND CREATING REPORT---------------------------------------------------------
Public Function createLogReport(ByRef wb As Workbook, Optional ByRef ws As Worksheet = Nothing, _
Optional ByVal logFileFullName As String = "") _
As Worksheet

    If PATH_SEP = "" Then
        raiseLoggerError NOINIT, "VBALogger.vbLogger"
    End If

    Dim arr, i As Long, x As Long, y As Integer, redI As Integer
    
    Application.ScreenUpdating = False

    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    Else
        Dim result As VbMsgBoxResult
        result = MsgBox( _
                    Title:="LogReportHelper", _
                    Prompt:="Alle auf dem Arbeitsblatt '" & ws.name & "' vorhandenen Daten werden gelöscht" & vbNewLine & _
                            "Soll fortgefahren werden?", _
                    Buttons:=vbYesNo + vbDefaultButton1)
                    
        If result = vbNo Then
            Exit Function
        End If
        
        ws.UsedRange.Clear
    End If

    arr = parseLogFile()

    With ws
        .name = "LogReport" & " " & Format(Now, "yyyy-mm-dd hh-nn-ss")
        
        For i = LBound(arr) To UBound(arr)                                  ' Add logging data to worksheet
            arr(i) = Replace(arr(i), "<t/>", vbTab)                         ' Change tab marker to real tab
            .Range("A1").Offset(x, y).Value2 = arr(i)
            y = y + 1
            If y = 4 Then
                y = 0
                x = x + 1
            End If
        Next i
        
        Erase arr
        
        Dim tbl As ListObject, r As Range, pvt As PivotTable                ' Create data table and pivot table
        Set tbl = turnRangeIntoTable(ws, .Range("A1").CurrentRegion, "tblLog-" & Format(Now, "yyyymmddhhnnss"))
        
        For Each r In tbl.ListColumns(2).DataBodyRange
            arr = Split(r.Value2, " ")
            r.Value2 = CDate(Replace(arr(0), "-", "/") & " " & Replace(arr(1), "-", ":"))
        Next r
        
        Erase arr
        tbl.ListColumns(2).DataBodyRange.NumberFormat = "dd.mm.yyyy"
        
        ' Create pivot table
        Set r = .Range("F1")
        Set pvt = createPivotTableFromTable(wb, tbl, r)
        
        ' Pivot table settings
        pvt.TableStyle2 = "PivotStyleMedium15"
        pvt.RowAxisLayout xlTabularRow
        
        ' Add fields
        With pvt.PivotFields("DATETIME")
            .Orientation = xlRowField
            .Position = 1
        End With
        .Range("F2").Group _
                        Start:=True, _
                        End:=True, _
                        Periods:=Array(False, False, False, False, True, False, True)   ' Group by month and year
'                                      Sec    Min    Hour   Day    Mnth  Qrtr   year
                        
        With pvt.PivotFields("LEVEL")
            .Orientation = xlRowField
            .Position = 3
        End With
        With pvt.PivotFields("LEVEL")
            .Orientation = xlDataField
            .Position = 1
            .Function = xlCount
            '.name = "Count per Level"
        End With
        
        Set tbl = Nothing
        Set pvt = Nothing
        
        .Columns.AutoFit
        .Activate
    End With

    Set createLogReport = ws

    Dim resultEmptyLines As VbMsgBoxResult
    If EMPTY_LINE_IN_LOG Then
        resultEmptyLines = _
            MsgBox( _
              Title:="TextFileHandler Warning", _
              Prompt:="Es wurden leere Zeilen in '" & _
                  Mid(LOG_FILE_PATH, InStrRev(LOG_FILE_PATH, PATH_SEP) + 1) & "' gefunden" & vbNewLine & _
                  "Leere Zielen werden grundsätzlich übersprungen, können aber entfernt werden" & vbNewLine & _
                      vbNewLine & _
                      "Sollen die leeren Zeilen entfernt werden?", _
              Buttons:=vbQuestion + vbYesNo + vbDefaultButton1)
        
        If resultEmptyLines = vbYes Then
            removeEmptyLinesFromLog
        End If
    End If

    Application.ScreenUpdating = True

End Function

Private Function parseLogFile(Optional fileFullName As String = "") As Variant()

    Dim ff As Double, curLine As String, arr, outputArr() As Variant, idx As Long, i As Long
    
    If fileFullName = "" Then
        fileFullName = LOG_FILE_PATH
    End If
    
    ReDim outputArr(0)
    
    ff = FreeFile()
    Open fileFullName For Input As #ff
        While Not EOF(ff)
            Line Input #ff, curLine                             ' Read Line
            
            If curLine = "" Then
                EMPTY_LINE_IN_LOG = True                        ' Set empty line flag, prompt user in caller function
                GoTo JmpSkipLine
            End If
            
            curLine = Replace(curLine, "|", "")                 ' Remove '|'
            curLine = Replace(curLine, "<br/>", vbNewLine)      ' Change line brake marker to real line break
            arr = removeEmptyIdx(Split(curLine, vbTab), -1)     ' Split on tab and remove empty indices
            
            If UBound(outputArr) = 0 Then
                idx = 0
            Else
                idx = UBound(outputArr) + 1
            End If
            
            If idx = 0 Then
                ReDim Preserve outputArr(0 To UBound(arr) - 1)
            Else
                ReDim Preserve outputArr(0 To (UBound(outputArr) + UBound(arr)))
            End If
            
            For i = LBound(arr) To UBound(arr)                  ' Add line data to ouput array
                outputArr((idx + i) - 1) = arr(i)
            Next i
            
            Erase arr
JmpSkipLine:
        Wend
    Close #ff

    parseLogFile = outputArr

End Function

Private Function removeEmptyIdx(ByRef arr, Optional ByVal startIdx As Integer = -1) As Variant()

    Dim i As Integer, nextIdx As Integer
    
    If startIdx = -1 Then
        startIdx = LBound(arr)
    End If
    
    While True 'Not arr(i) = ""
JmpIncrement:
        If Not arr(startIdx) = "" Then
            startIdx = startIdx + 1
            
            If startIdx > UBound(arr) Then
                removeEmptyIdx = WorksheetFunction.Transpose(WorksheetFunction.Transpose(Array(arr)))
                Exit Function
            End If
            
            GoTo JmpIncrement
        End If
        
        nextIdx = nextUsedIdx(arr, startIdx)
        
        If nextIdx = -1 Then
            ReDim Preserve arr(LBound(arr) To startIdx)
            removeEmptyIdx = WorksheetFunction.Transpose(WorksheetFunction.Transpose(Array(arr)))
            Exit Function
        Else
            arr(startIdx) = arr(nextIdx)
            arr(nextIdx) = ""
        End If
        
        If nextIdx = UBound(arr) Then
            ReDim Preserve arr(LBound(arr) To startIdx)
            removeEmptyIdx = WorksheetFunction.Transpose(WorksheetFunction.Transpose(Array(arr)))
            Exit Function
        End If
    Wend
    
End Function

' @returns Last not empty array index or -1 if array is empty after the given index
Private Function nextUsedIdx(ByRef arr As Variant, ByVal idx As Integer) As Integer

    Do
        idx = idx + 1
        
        If idx > UBound(arr) Then
            nextUsedIdx = -1
            Exit Function
        End If
    Loop While arr(idx) = ""

    nextUsedIdx = idx

End Function

'------CUSTOM EXCEPTIONS----------------------------------------------------------------------------
Private Function raiseLoggerError(ByVal errNum As LogErrorNum, ByVal qualifiedPath As String) As String

    Dim errTxt As String

    Select Case errNum
        Case 100 ' NOINIT
            errTxt = "Logger was not initialized after instantiation" & vbNewLine & _
                     "Use objLogger.initLogger to initalize the instance"
        Case 101 ' WRONG_CONFIG_NAME
            errTxt = "The given path does not lead to a valid log file" & vbNewLine & _
                     "Use a valid file path and retry"
        Case 102 ' ROLL_INTERVAL
            errTxt = "The supplied interval in days is not valid" & vbNewLine & _
                     "Use a number > 0 and retry"
        Case 103 ' BAD_CONFIG_PATH
            errTxt = "The internal path does not lead to a valid config file" & vbNewLine & _
                     "Check if the config file exists" & vbNewLine & _
                     "If it does not exist, use objLogger.initRollOver to recreate it"
        Case 104 ' CONIFG_IDX
            errTxt = "While parsing the config file an IndexOutOfBoundsException occured" & vbNewLine & _
                     "Check the config file and retry" & vbNewLine & vbNewLine & _
                     "Should the error occur again, remove the config file" & vbNewLine & _
                     "Then use objLogger.initRollOver to recreate it"
        Case 999 ' MSC_EXCEPTION
            Me.logText CRIT, "Unknown-exception vbLogger occurred", "vbLogger", True
            errTxt = "An unknown internal error occurred" & vbNewLine & _
                     "Please consult the AddIn-Supplier for help"
    End Select
    
    Err.Raise Number:=vbObjectError + errNum, _
              source:=qualifiedPath, _
              Description:=errTxt

End Function

'------MISC FUNCTIONS------------------------------------------------------------------------------
Private Function turnRangeIntoTable(ws As Worksheet, rng As Range, name As String) As ListObject
    
    Dim tbl As ListObject
    Set tbl = ws.ListObjects.Add( _
                    SourceType:=xlSrcRange, _
                    source:=rng, _
                    XlListObjectHasHeaders:=xlYes, _
                    TableStyleName:="TableStyleLight8")
    tbl.name = name
    
    Set turnRangeIntoTable = tbl
    
End Function

Private Function createPivotTableFromTable(ByRef wb As Workbook, ByRef tbl As ListObject, ByRef trgR As Range) _
As PivotTable

    Dim pvc As PivotCache
    
    Set pvc = wb.PivotCaches.Create(SourceType:=xlDatabase, _
                                    SourceData:=tbl.Range)
                                    
    Set createPivotTableFromTable = pvc.createPivotTable( _
                                        TableDestination:=trgR, _
                                        TableName:="pvtLogReport-" & Format(Now, "yyyymmddhhnnss"))

End Function

' Writes all data from the original log file to a temporary log file
' Removes the original log file and places the temporary one in it's stead
Private Sub removeEmptyLinesFromLog()

    Dim tmpFullFileName As String, fso As Object, oFile As Object, ff As Double, curLine As String

    tmpFullFileName = LOG_FILE_PATH & ".tmp"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set oFile = fso.createTextFile(tmpFullFileName)
    
    ff = FreeFile()
    
    Open LOG_FILE_PATH For Input As #ff
        While Not EOF(ff)
            Line Input #ff, curLine
        
            If Not curLine = "" Then
                oFile.WriteLine curLine
            End If
        Wend
    Close #ff

    oFile.Close
    
    Set oFile = Nothing
    
    Kill LOG_FILE_PATH
    fso.MoveFile tmpFullFileName, LOG_FILE_PATH
        
    Set fso = Nothing
    
    EMPTY_LINE_IN_LOG = False                               ' Reset empty line flag

End Sub

' Check if the last line of the file is empty and ready for appending
' If not enter a line break and close the file
' @returns True if the operation succeded, false if the supplied file path was not valid
Private Function prepareTextFileForAppend(ByVal fileFullName As String) As Boolean

    If Len(Dir(fileFullName)) = 0 Then
        prepareTextFileForAppend = False
        Exit Function
    End If

    Dim ff As Double, logTxt As String, arrLine As Variant
    
    ff = FreeFile()
    Open fileFullName For Input As #ff
        logTxt = Input(LOF(ff), #ff)                        ' Get contents of file
    Close #ff
    
    arrLine = Split(logTxt, vbLf)                           ' Split on line feed
    logTxt = ""
    
    If Not arrLine(UBound(arrLine)) = "" Then               ' Append empty line if last line is not empty
        ff = FreeFile()
        Open fileFullName For Append As #ff
            Print #ff, ""
        Close #ff
    End If

    Erase arrLine

    prepareTextFileForAppend = True

End Function






















































