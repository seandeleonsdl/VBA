Public Sub Main()


    LoopThroughContents ("C:\Users\Sean\Desktop\New folder")
    
End Sub


Public Sub LoopThroughContents(FolderDirectory As String)
    
    Dim Files As String
    Dim FolderDir As String
    
    Content = Dir(FolderDirectory & "\", vbDirectory)
    
    Do While Len(Content) > 0
        
        
        Content = Dir
        
        If Content <> "" And Content <> "." And Content <> ".." Then
            Files = Files & "[" & Content & "]" & vbNewLine
        End If
        
    Loop
    
    MsgBox Left(Files, Len(Files) - 1)
    
    
End Sub
