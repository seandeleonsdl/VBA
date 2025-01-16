Attribute VB_Name = "modGeneral"
Public Function LoopThroughContents(FolderDirectory As String) As String()
    
    Dim Files As String
    Dim FolderDir As String
    
    Content = Dir(FolderDirectory & "\*", vbDirectory)
    
    Do While Len(Content) > 0
        
        If Content <> "" And Content <> "." And Content <> ".." Then
            Files = Files & Content & "|"
        
        End If
        
        
        Content = Dir
        
    Loop
    Files = Left(Files, Len(Files) - 1)
    'MsgBox "[" & Files & "]"
    LoopThroughContents = Split(Files, "|")
    
    
End Function
