Attribute VB_Name = "modTransferNewFiles"
Public Sub MoveNewFiles(NewFileDirectory As String, Directory As String)
    'MsgBox "MoveNewFiles"
    '------------------------------
    ' Filter-out new files to move
    '------------------------------
    Dim File As Variant, FileStream As String
    Dim FilesToMove() As String
    FilesToMove = InterestList(modGeneral.LoopThroughContents(NewFileDirectory), ".txt")
    For Each File In FilesToMove
        MsgBox File
        FileStream = FileStream & NewFileDirectory & "\" & File & "|"
    Next
    FilesToMove = Split(Left(FileStream, Len(FileStream) - 1), "|")
    DisplayArray FilesToMove
    
    '------------------------------
    ' Move new files to new folder
    '------------------------------
    Dim FSO As Object
    Set FSO = VBA.CreateObject("Scripting.FileSystemObject")
    
    Dim Filename As String
    For Each File In FilesToMove
        Filename = Split(File, "\")(UBound(Split(File, "\")))
        If Filename = "" Then Exit For
        'MsgBox "File: " & File
        'MsgBox "Destination: " & Directory & "\" & Filename
        FSO.MoveFile File, Directory & "\" & Filename
    Next
    'MsgBox "MoveFiles succeeded"
    
    
    
End Sub

Public Function InterestList(FilesToCount() As String, SearchString As String) As String()
    
    
    Dim File As Variant
    Dim stream As String
    For Each File In FilesToCount
        
        If InStr(1, File, SearchString) Then
            stream = stream & File & "|"
            'MsgBox SearchString & " in " & File
        End If
    Next
    
    If stream = "" Then
        Dim EmptyString(1) As String
        EmptyString(0) = ""
        InterestList = EmptyString
        Exit Function
    End If
    
    stream = Left(stream, Len(stream) - 1)
    MsgBox "[" & stream & "]"
    InterestList = Split(stream, "|")
    
End Function

