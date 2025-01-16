Attribute VB_Name = "modBackUpFiles"





Public Sub BackUpFiles(Directory As String, FilesToBackup() As String)
    
     ExclusionList = ""
    
    
    '-----------------------
    ' Create Back-up Folder
    '-----------------------
    Dim BackUpFolderName As String
    BackUpFolderName = "AOI OLD"
    If Dir(Directory & "\" & BackUpFolderName, vbDirectory) = "" Then
        MkDir Directory & "\" & BackUpFolderName
    Else:
        MsgBox BackUpFolderName & " Exists"
    End If
    
    '-----------------------------
    ' Filter-out files to back-up
    '-----------------------------
    Dim ExclusionArray() As String
    ExclusionArray = Split(ExclusionList, "|")
    FilesToBackup = ExcludeList(modGeneral.LoopThroughContents(Directory), ExclusionArray)
    
    Dim File As Variant, FileStream As String
    For Each File In FilesToBackup
        FileStream = FileStream & Directory & "\" & File & "|"
    Next
    FilesToBackup = Split(Left(FileStream, Len(FileStream) - 1), "|")
    
    '----------------------------
    ' Move files to BackUpFolder
    '----------------------------
    'MsgBox Directory & "\" & BackUpFolderName

    MoveBackupFiles Directory & "\" & BackUpFolderName, FilesToBackup
    
End Sub





Public Function ExcludeList(OriginalList() As String, ExclusionList() As String) As String()
    
    Dim TempList() As String
    TempList = OriginalList
    
    For i = 0 To UBound(ExclusionList)
        TempList = RemoveElement(TempList, ExclusionList(i))
        'DisplayArray TempList
        
    Next
    
    ExcludeList = TempList
    
End Function

Public Function DisplayArray(StringArray() As String)
    
    Dim s As Variant, stream As String
    
    'MsgBox UBound(StringArray)
    
    For Each s In StringArray
        stream = stream & s & vbNewLine
    Next
    stream = Left(stream, Len(stream) - 1)
    MsgBox stream
    
End Function

Public Function RemoveElement(StringArray() As String, Element As String) As String()
    
    'MsgBox "RemoveElement" & vbNewLine & "Element: " & Element
    
    'DisplayArray StringArray
    'MsgBox "UBound1: " & UBound(StringArray)
 
    Dim s As Variant, stream As String
    stream = ""
    For Each s In StringArray
        If s <> Element Then
            stream = stream & s & "|"
        End If
    Next
    
       
    If stream = "" Then
        Dim EmptyString(1) As String
        EmptyString(0) = ""
        RemoveElement = EmptyString
        Exit Function
    End If
    
    stream = Left(stream, Len(stream) - 1)
    RemoveElement = Split(stream, "|")
    
   
    'MsgBox "[" & Stream & "]"
    'MsgBox "RemoveElement succeeded"
    

End Function

Public Sub MoveBackupFiles(BackUpFolderName As String, FilesToMove() As String)
    
    'MsgBox "MoveFiles start"
    Dim FSO As Object
    Set FSO = VBA.CreateObject("Scripting.FileSystemObject")

    Dim File As Variant
    Dim Filename As String
    
    For Each File In FilesToMove
        Filename = Split(File, "\")(UBound(Split(File, "\")))
        If Filename = "" Then Exit For
        'MsgBox "File: " & File
        'MsgBox "BackUpFolderName: " & BackUpFolderName & "\" & Filename
        FSO.MoveFile File, BackUpFolderName & "\" & Filename
    Next
    'MsgBox "MoveFiles succeeded"
    
End Sub


