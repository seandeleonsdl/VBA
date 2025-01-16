Attribute VB_Name = "modMain"
Public Sub Main()

    Dim Directory As String
    Directory = "C:\Users\Sean\Desktop\New folder"
    mod_01_ReadReport.ReadReport Directory, "August"
    'mod_02_BackUpFiles.BackUpFiles Directory, LoopThroughContents("C:\Users\Sean\Desktop\New folder")
    'mod_03_TransferNewFiles.MoveNewFiles "C:\Users\Sean\Desktop\New Maps", Directory
    
    
End Sub
