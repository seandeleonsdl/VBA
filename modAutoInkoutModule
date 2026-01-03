



Private checkStep() As Class1



Private Sub Label11_Click()

End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub txtInkoutMapDirectory_Change()

End Sub

Private Sub UserForm_Initialize()
        
    '------------------------
    ' 1.) Program properties
    '------------------------
    PROPERTIES_PanelMapDirectory = ThisWorkbook.Sheets(1).Range("C2").Value2
    PROPERTIES_InkoutMapDirectory = ThisWorkbook.Sheets(1).Range("C3").Value2
    PROPERTIES_MergeMapDirectory = ThisWorkbook.Sheets(1).Range("C4").Value2
    PROPERTIES_MaxScheduleCount = ThisWorkbook.Sheets(1).Range("C5").Value2
    PROPERTIES_MaxStepCount = ThisWorkbook.Sheets(1).Range("C6").Value2
    
    ReDim checkStep(1 To PROPERTIES_MaxScheduleCount, 1 To PROPERTIES_MaxStepCount)
    ReDim labelSchedule(1 To PROPERTIES_MaxScheduleCount) As Object
    ReDim LABEL_ScheduleBackground(1 To PROPERTIES_MaxScheduleCount) As Object
    
    Me.txtPanelMapDirectory = PROPERTIES_PanelMapDirectory
    Me.txtInkoutMapDirectory = PROPERTIES_InkoutMapDirectory
    Me.txtMergeMapDirectory = PROPERTIES_MergeMapDirectory
    
    '--------------------------------
    ' 2.) Set-up userform aesthetics
    '--------------------------------
    Me.BackColor = RGB(252, 252, 252)
    Me.Label7.BackColor = RGB(220, 234, 247)
    Me.lblBackground.BackColor = RGB(240, 240, 240)
    
    
    For i = 1 To PROPERTIES_MaxScheduleCount
        
        '----------------------
        ' 3.) Set-up schedules
        '--------------------------
        LabelName = "labelSchedule_" & Format(i, "00")
        Set LABEL_ScheduleBackground(i) = Me.Controls.Add("Forms.Label.1", "LabelBackground_" & Format(i, "00"), True)
        Set labelSchedule(i) = Me.Controls.Add("Forms.Label.1", LabelName, True)
        
        '----------------
        ' a.) Dimensions
        '----------------
        LABEL_ScheduleBackground(i).Width = 65
        LABEL_ScheduleBackground(i).Height = 18
        
        labelSchedule(i).Width = 65
        
        '--------------
        ' b.) Position
        '--------------
        LABEL_ScheduleBackground(i).Left = Me.lblScheduleBackground.Left
        LABEL_ScheduleBackground(i).Top = Me.lblScheduleBackground.Top + Me.lblScheduleBackground.Height + 6 + 24 * (i - 1)
        
        labelSchedule(i).Left = Me.lblScheduleBackground.Left + 6
        labelSchedule(i).Top = Me.lblScheduleBackground.Top + Me.lblScheduleBackground.Height + 6 + 4 + 24 * (i - 1)
        
        '-------------
        ' c.) Styling
        '-------------
        LABEL_ScheduleBackground(i).BackColor = RGB(240, 240, 240)
        LABEL_ScheduleBackground(i).Visible = False
        
        labelSchedule(i).BackStyle = fmBackStyleTransparent
        labelSchedule(i).Caption = ""
        labelSchedule(i).Font.Name = "Courier New"
        labelSchedule(i).Font.Size = 8
        
        
        For j = 1 To PROPERTIES_MaxStepCount
        
            checkboxName = "checkStep_" & Format(i, "00") & "_" & Format(j, "00")
            Set checkboxInstance = Me.Controls.Add("Forms.Label.1", checkboxName, True)
            
            '----------------
            ' a.) Dimensions
            '----------------
            checkboxInstance.Width = 30
            checkboxInstance.Height = 18
            
            '--------------
            ' b.) Position
            '--------------
            checkboxInstance.Left = LABEL_ScheduleBackground(i).Left + LABEL_ScheduleBackground(i).Width + 6 + (checkboxInstance.Width + 6) * (j - 1)
            checkboxInstance.Top = LABEL_ScheduleBackground(i).Top
            
            '-------------
            ' c.) Styling
            '-------------
            checkboxInstance.SpecialEffect = fmButtonEffectFlat
            checkboxInstance.BackColor = RGB(240, 240, 240)
            checkboxInstance.ForeColor = RGB(255, 255, 255)
            checkboxInstance.Font.Name = "Courier New"
            checkboxInstance.Font.Size = 8
            checkboxInstance.Font.Bold = True
            checkboxInstance.Caption = ""
            If i > 1 Then checkboxInstance.Visible = False
            
            Set checkStep(i, j) = New Class1
            Set checkStep(i, j).checkbox = checkboxInstance
            Set checkStep(i, j).Range = ThisWorkbook.Sheets(2).Range("B2").Offset(i - 1, j - 1)
            checkStep(i, j).Row = i
            
            '----------------
            ' d.)
            '---------------
            'toggleYes = Me.Controls.Add("Forms.Label.1", checkboxName, True)
            
        
        Next
        
    Next
    
    LABEL_ScheduleBackground(1).Visible = True
    For j = 1 To PROPERTIES_MaxStepCount
        checkStep(1, j).checkbox.Visible = True
        
    Next
    
    
    
End Sub



Public Function CountF30(ByVal Schedule As String, ByVal Step As String)
    
    
    
    '-----------------
    ' Get Directories
    '-----------------
    '(1) Get Directories
    InkoutMapWorkWeek = Right(Year(Now), 2)
    InkoutMapScheduleDigits = Left(Schedule, 2)
    InkoutMapDirectory = Me.txtInkoutMapDirectory & "\W" & InkoutMapWorkWeek & InkoutMapScheduleDigits
    
    '(2) Get Map List from Bumping2 and PanelMap
    MapList = GetMapList(InkoutMapDirectory & "\" & Schedule & "\" & CStr(Step) & "\")
    
    For Each InkoutMap In MapList
        
        InkoutMapID = GetWaferID(InkoutMap)
        
        fileNum = FreeFile
        filePath = InkoutMapDirectory & "\" & Schedule & "\" & CStr(Step) & "\" & InkoutMap
        ThisWorkbook.Sheets(1).Range("A15").Value = filePath
        Open filePath For Input As #fileNum
        fileContent = ""
        
        Do Until EOF(fileNum)
            Line Input #fileNum, LineInstance
            fileContent = fileContent & vbNewLine & LineInstance
        Loop
        Close #fileNum
        
        DefectCount = CountString(fileContent, "019") + CountString(fileContent, "020") + CountString(fileContent, "021") + CountString(fileContent, "030")
        CountF30 = CountF30 & "S" & InkoutMapID & "-" & DefectCount & "ea,"
        '(6) Sort through each PanelMap
         
    Next
    
    

End Function

Public Function CountString(ByVal mainString As String, ByVal subString As String) As Long
    Dim position As Long
    Dim count As Long
    
    count = 0
    position = InStr(1, mainString, subString, vbTextCompare) ' vbTextCompare = case-insensitive
    'MsgBox position
    Do While position > 0
        count = count + 1
        position = InStr(position + 1, mainString, subString, vbTextCompare)
    Loop
    
    CountString = count
    
    
End Function



Private Sub Inkout(ByVal Schedule As String, ByVal Step As Integer)

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    '-----------------
    ' Get Directories
    '-----------------
    '(1) Get Directories
    InkoutMapWorkWeek = Right(Year(Now), 2)
    InkoutMapScheduleDigits = Left(Schedule, 2)
    InkoutMapDirectory = Me.txtInkoutMapDirectory & "\W" & InkoutMapWorkWeek & InkoutMapScheduleDigits
    'InkoutMapDirectory =WorkWeek = Right(Year(Now), 2)
    
    
    
    PanelMapDirectory = Me.txtPanelMapDirectory
    MergeMapDirectory = Me.txtMergeMapDirectory
    
    '(2) Create Inkout Folders in MergeMapDirectory
    If Dir(MergeMapDirectory, vbDirectory) = "" Then MkDir MergeMapDirectory
    If Dir(MergeMapDirectory & "\Inkout Map", vbDirectory) = "" Then MkDir (MergeMapDirectory & "\Inkout Map")
    If Dir(MergeMapDirectory & "\Panel Map", vbDirectory) = "" Then MkDir (MergeMapDirectory & "\Panel Map")
    If Dir(MergeMapDirectory & "\Merge Map", vbDirectory) = "" Then MkDir (MergeMapDirectory & "\Merge Map")
    
    
    '(3) Get Map List from Bumping2 and PanelMap
    MapList = GetMapList(InkoutMapDirectory & "\" & Schedule & "\" & CStr(Step) & "\")
    PanelMapList = GetMapList(PanelMapDirectory & "\" & Schedule & "*")
    
    For Each InkoutMap In MapList
           
        '(4) Get InkoutMap from Bumping2
        FileToCopy = InkoutMapDirectory & "\" & Schedule & "\" & CStr(Step) & "\" & InkoutMap
        SaveFileAs = MergeMapDirectory & "\Inkout Map\" & InkoutMap
        fso.CopyFile FileToCopy, SaveFileAs
        
        '(5) Convert Defect Code
        ConvertDefectCode SaveFileAs
        
        '(6) Sort through each PanelMap
        For Each PanelMap In PanelMapList
        
            InkoutMapID = GetWaferID(InkoutMap)
            PanelMapID = GetWaferID(PanelMap)
            
            If InkoutMapID = PanelMapID Then
                
                FileToCopy = PanelMapDirectory & "\" & PanelMap
                SaveFileAs = MergeMapDirectory & "\Panel Map\" & PanelMap
                fso.CopyFile FileToCopy, SaveFileAs
                
                MergeMapContent = MergeMap("C:\Users\Sean\OneDrive\Desktop\Onyx\PanelMap\" & PanelMap, MergeMapDirectory & "\Inkout Map\" & InkoutMap)
                filePath = MergeMapDirectory & "\Merge Map\" & PanelMap
                
                If Dir(MergeMapDirectory & "\Merge Map", vbDirectory) = "" Then
                    MkDir MergeMapDirectory & "\MergeMap"
                End If
                
                fileNum = FreeFile
                Open filePath For Output As #fileNum
                Print #fileNum, MergeMapContent
                Close #fileNum
                
                Exit For
            End If
        Next
            
    Next



End Sub




Private Sub btnInkout_Click()
    
    Dim ScheduleCount As Integer, StepCount As Integer
    ScheduleCount = ThisWorkbook.Sheets(2).Range("A" & Rows.count).End(xlUp).Row - 1
    StepCount = ThisWorkbook.Sheets(2).Cells(1, Columns.count).End(xlToLeft).Column - 1
    
    Dim InkoutSchedules As String
    
    For i = 1 To ScheduleCount
        Schedule = ThisWorkbook.Sheets(2).Range("A2").Offset(i - 1, 0).Value2
        For j = 1 To StepCount
            Step = ThisWorkbook.Sheets(2).Range("B1").Offset(0, j - 1).Value2
            NeedInkout = ThisWorkbook.Sheets(2).Range("B2").Offset(i - 1, j - 1).Value
            
            If NeedInkout = True Then
                InkoutSchedules = InkoutSchedules & Schedule & " " & Step & vbNewLine
            End If
            
        Next
    Next
    InkoutSchedules = Left(InkoutSchedules, Len(InkoutSchedules) - 1)
    
    
    For Each InkoutSchedule In Split(InkoutSchedules, vbNewLine)
        MsgBox InkoutSchedule
        Inkout Split(InkoutSchedule, " ")(0), Split(InkoutSchedule, " ")(1)
    Next
       
'    MsgBox InkoutSchedules

End Sub








Private Function GetWaferID(ByVal Filename As String) As Integer

    Filename = Replace(Filename, ".txt", "")
    Filename = Replace(Filename, ".", "-")
    Filename = Split(Filename, "-")(UBound(Split(Filename, "-")))
    Filename = Replace(Filename, "S", "")
    Filename = Left(Filename, 2)
    
    GetWaferID = Int(Filename)

End Function

Private Function MergeMap(ByVal PanelMapFilename As String, InkoutMapFilename As String) As String
    
    filePath = Filename
    fileNum = FreeFile
    
    
    '-----------------------
    ' Get Panel Map Content
    '-----------------------
    Open PanelMapFilename For Input As #fileNum
    
    Do Until EOF(fileNum)
        Line Input #fileNum, LineInstance
        PanelMapFileContent = PanelMapFileContent & vbNewLine & LineInstance
    Loop
    Close #fileNum
    PanelMapFileContent = Right(PanelMapFileContent, Len(PanelMapFileContent) - 2)
    
    '------------------------
    ' Get Inkout Map Content
    '------------------------
    Open InkoutMapFilename For Input As #fileNum
    
    Do Until EOF(fileNum)
        Line Input #fileNum, LineInstance
        InkoutMapFileContent = InkoutMapFileContent & vbNewLine & LineInstance
    Loop
    Close #fileNum
    InkoutMapFileContent = Right(InkoutMapFileContent, Len(InkoutMapFileContent) - 2)
    
    
    'MsgBox "PanelMapContent:" & vbNewLine & PanelMapFileContent
    'MsgBox "InkoutMapFileContent:" & vbNewLine & InkoutMapFileContent
    
    
    '-------------------------
    ' Merge Map
    '-------------------------
    
    
    
    Index = InStr(1, InkoutMapFileContent, "030")
    
    Do While Index > 0
        
        DefectCode = Mid(PanelMapFileContent, Index, 3)
        If DefectCode = "000" Then
            'MsgBox "BEFORE:" & vbNewLine & PanelMapFileContent
            'MsgBox "REPLACE: " & vbNewLine & Replace(PanelMapFileContent, "000", "030", Index, 1)
            PanelMapFileContent = Left(PanelMapFileContent, Index - 1) & Replace(PanelMapFileContent, "000", "030", Index, 1)
            'MsgBox "AFTER:" & vbNewLine & PanelMapFileContent
        End If
        Index = InStr(Index + 1, InkoutMapFileContent, "030")
    Loop
    
    
    MergeMap = PanelMapFileContent
    
    'MsgBox "[" & DefectCode & "]"
    
    
    'PanelMapFileContent
    
    
    
    
    
    'MsgBox PanelMapFileContent
    'MsgBox InkoutMapFileContent
    
    
    
    
End Function




Private Sub ConvertDefectCode(ByVal Filename As String)
    
    Dim filePath As String
    Dim fileNum As Integer
    Dim fileContent As String
    
    filePath = Filename
    fileNum = FreeFile
    
    Open filePath For Input As #fileNum
    fileContent = ""
    Do Until EOF(fileNum)
        Line Input #fileNum, LineInstance
        fileContent = fileContent & vbNewLine & LineInstance
    Loop
    
    If fileContent = "" Then
        MsgBox "No Map Found"
        Exit Sub
    End If
    fileContent = Right(fileContent, Len(fileContent) - 2)
    fileContent = Replace(fileContent, "019", "030")
    fileContent = Replace(fileContent, "020", "030")
    fileContent = Replace(fileContent, "021", "030")
    
    'MsgBox "[" & fileContent & "]"
    
    Close #fileNum
    
    Open filePath For Output As #fileNum
    Print #fileNum, fileContent
    Close #fileNum

End Sub



Private Function GetMapList(ByVal Directory As String)
    
    Files = ""
    Filename = Dir(Directory)
    ThisWorkbook.Sheets(1).Range("B7").Value = Directory
    Do While Filename <> ""
    
        '------------------------
        ' File Filter Conditions
        '------------------------
        FileIsMap = Filename <> vbNewLine And Len(Filename) >= 6
        FileIsMap = FileIsMap And InStr(Filename, Schedule) > 0
        FileIsMap = FileIsMap And InStr(Filename, "August") = 0
       
        If FileIsMap Then
            Files = Files & "\" & Filename
        End If
        Filename = Dir
    Loop
    
    If Len(Files) = 0 Then
        MsgBox "No maps found"
        Exit Function
        
    End If
    Files = Right(Files, Len(Files) - 1)
    
    GetMapList = Split(Files, "\")
    
End Function



Private Sub txtInkoutSchedules_Change()

    '----------------------
    ' TextBox Input Filter
    '----------------------
    txtInkoutSchedules.Value = UCase(txtInkoutSchedules.Value)
    
    If Right(txtInkoutSchedules.Value, 1) = vbLf Then
        txtInkoutSchedules.Value = Left(txtInkoutSchedules.Value, Len(txtInkoutSchedules.Value) - 2)
        Exit Sub
    End If
    
    If Len(txtInkoutSchedules.Value) < 10 Then Exit Sub
    Schedule = txtInkoutSchedules.Value
    
    '---------------------------------
    ' 1.) Retrieve program properties
    '---------------------------------
    PROPERTIES_MaxScheduleCount = ThisWorkbook.Sheets(1).Range("C5").Value2
    PROPERTIES_MaxStepCount = ThisWorkbook.Sheets(1).Range("C6").Value2
    'IN/6440/AEI/6730/6630/6733/6688/6905
    
    
    '-----------------------------
    ' 2. Reset/clear all controls
    '-----------------------------
    For i = 1 To PROPERTIES_MaxScheduleCount
        For j = 1 To PROPERTIES_MaxStepCount
            Me.Controls("checkStep_" & Format(i, "00") & "_" & Format(j, "00")).BackColor = RGB(240, 240, 240)
            Me.Controls("checkStep_" & Format(i, "00") & "_" & Format(j, "00")).Caption = ""
        Next
    Next
    
    LastCell = Replace(ThisWorkbook.Sheets(2).Cells(1 + PROPERTIES_MaxScheduleCount, 1 + PROPERTIES_MaxStepCount).Address, "$", "")
    
    ThisWorkbook.Sheets(2).Range("B2:" & LastCell).Value = ""
    
    '---------------
    '
    '---------------
    Dim ScheduleList As Variant, ScheduleCount As Integer, WorkWeek As String
    ScheduleList = Split(txtInkoutSchedules.Value, vbNewLine)
    ScheduleCount = UBound(ScheduleList) + 1
    WorkWeek = Right(Year(Now), 2)
    
    InkoutMapDirectory = Me.txtInkoutMapDirectory.Value
    
    Dim StepList() As Integer, StepCount As Integer
    ReDim StepList(1 To PROPERTIES_MaxStepCount)
    StepCount = 1
    
    
    Dim Active() As Variant
    ReDim Active(1 To ScheduleCount, 1 To PROPERTIES_MaxStepCount)

    Dim StepListString As String
    StepList(1) = 6730
    StepListString = StepListString & "[6730]"
    
    
    
    '-----------------------------------------------
    ' Get all the available steps for each schedule
    '-----------------------------------------------
    
    '(1) Get Directories
    InkoutMapWorkWeek = Right(Year(Now), 2)
    InkoutMapScheduleDigits = Left(Schedule, 2)
    InkoutMapDirectory = Me.txtInkoutMapDirectory & "\W" & InkoutMapWorkWeek & InkoutMapScheduleDigits
    
    For i = 1 To ScheduleCount
        For j = 1 To PROPERTIES_MaxStepCount
            Active(i, j) = "N/A"
        Next
    Next
    
    
    For i = 0 To UBound(ScheduleList)
        
        StepFolder = Dir(InkoutMapDirectory & "\" & ScheduleList(i) & "\", vbDirectory)
        
        Me.Controls("LabelBackground_" & Format(i + 1, "00")).Visible = True
        Do While StepFolder <> ""
            StepFolder = Dir
            
            '----------------------
            ' Step Folder Criteria
            '----------------------
            IsStepFolder = UCase(StepFolder) <> "ORG"
            IsStepFolder = StepFolder = "AEI"
            IsStepFolder = IsStep Or Len(StepFolder) = 4
            
            If IsStepFolder Then
            
                NotIncluded = InStr(1, StepListString, StepFolder) = 0
                If NotIncluded Then
                    StepListString = StepListString & "[" & StepFolder & "]"
                    StepCount = StepCount + 1
                    StepList(StepCount) = StepFolder
                End If
                
                For j = 1 To StepCount
                    
                    Me.Controls("checkStep_" & Format(i + 1, "00") & "_" & Format(j, "00")).Caption = StepList(j)
                          
                    If StepList(j) = StepFolder Then
                 
                        Active(i + 1, j) = False
                
                        Me.Controls("checkStep_" & Format(i + 1, "00") & "_" & Format(j, "00")).BackColor = RGB(220, 234, 247) ' Light Blue
                        Me.Controls("checkStep_" & Format(i + 1, "00") & "_" & Format(j, "00")).ForeColor = RGB(0, 0, 0)
                    
                    End If
                    
                Next
                
                Files = Files & "\" & StepFolder
            End If
        Loop
    Next
    LastCell = Replace(ThisWorkbook.Sheets(2).Cells(1 + ScheduleCount, 1 + StepCount).Address, "$", "")
    ThisWorkbook.Sheets(2).Range("B2:" & LastCell).Value = Active
    
    For i = 1 To ScheduleCount
        For j = 1 To StepCount
            
            Me.Controls("checkStep_" & Format(i, "00") & "_" & Format(j, "00")).BackColor = RGB(240, 240, 240)
            Me.Controls("checkStep_" & Format(i, "00") & "_" & Format(j, "00")).ForeColor = RGB(126, 126, 126)
            Me.Controls("checkStep_" & Format(i, "00") & "_" & Format(j, "00")).Visible = True
            
                        
            If Active(i, j) <> "N/A" Then
                Me.Controls("checkStep_" & Format(i, "00") & "_" & Format(j, "00")).BackColor = RGB(220, 234, 247) ' Light Blue
                Me.Controls("checkStep_" & Format(i, "00") & "_" & Format(j, "00")).ForeColor = RGB(0, 0, 0)
            
            End If
            
        Next
    Next
    
    
    
    
    
    '----------------------------------------------
    ' Update each control to reflect each schedule
    '----------------------------------------------
    Dim Schedules() As String, Steps() As Variant
    ReDim Schedules(0 To UBound(ScheduleList) + 1, 1)
    ReDim Steps(0, StepCount)
    For i = 1 To UBound(ScheduleList) + 1
        Me.Controls("labelSchedule_" & Format(i, "00")).Caption = ScheduleList(i - 1)
        Schedules(i - 1, 0) = ScheduleList(i - 1)
        
        For j = 1 To StepCount
    
            Me.Controls("checkStep_" & Format(i, "00") & "_" & Format(j, "00")).Caption = StepList(j)
            Steps(0, j - 1) = StepList(j)
        Next
    
    Next
    
    ThisWorkbook.Sheets(2).Range("A2:A" & 1 + ScheduleCount).Value = Schedules
    ThisWorkbook.Sheets(2).Range("B1:" & Replace(Cells(1, 1 + StepCount).Address, "$", "")).Value = Steps
    
    
    
    
    
    
    StepListString = Right(StepListString, Len(StepListString) - 1)
    
    
    Exit Sub
    
End Sub

Private Function CheckSteps(Schedule As String) As Variant
    
    '------------------
    ' Current Workweek
    '------------------
    WorkWeek = Right(Year(Now), 2)
    Directory = "\\bumping2\Map\W" & WorkWeek & Left(Schedule, 2) & "\" & Schedule
    
    '--------------
    ' Step Folders
    '--------------
    StepFolders = ""
    'StepFolder = Dir(Directory & "\", vbDirectory)
    StepFolder = Dir("C:\Users\Sean\OneDrive\Desktop\" & Schedule & "\", vbDirectory)
    Do While StepFolder <> ""
        StepFolder = Dir
            
        If StepFolder <> ".." And StepFolder <> "" Then
            StepFolders = StepFolders & "\" & StepFolder
        End If
    Loop
    
    If Len(StepFolders) > 0 Then StepFolders = Right(StepFolders, Len(StepFolders) - 1)
    
    CheckSteps = Split(StepFolders, "\")
    
    
    '----------------
    ' Update Folders
    '----------------
    ControlNo = 0
    
    For Each Step In Split(Files, "\")
    
        ControlNo = ControlNo + 1
        Me.Controls("checkStep" & Format(ControlNo, "00")).Caption = Step
    
    Next
        
End Function

Private Sub txtPanelMapDirectory_Change()
    ThisWorkbook.Sheets(1).Range("A2").Value2 = Me.txtPanelMapDirectory.Value
End Sub

Private Sub txtInkoutDirectory_Change()
    ThisWorkbook.Sheets(1).Range("A2").Value2 = Me.txtInkoutDirectory.Value
End Sub


