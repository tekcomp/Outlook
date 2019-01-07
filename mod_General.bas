Attribute VB_Name = "mod_General"
Option Explicit

'************************************************************************************************
'
'
'           Common Functions below
'
'************************************************************************************************

Function FindInFolders(TheFolders As Outlook.Folders, ByVal Name As String) As Outlook.MAPIFolder
  Dim Subfolder As Outlook.MAPIFolder
   
  On Error Resume Next
   
  Set FindInFolders = Nothing
   
  For Each Subfolder In TheFolders
    If LCase(Subfolder.Name) Like LCase(Name) Then
      Set FindInFolders = Subfolder
      Exit For
    Else
      Set FindInFolders = FindInFolders(Subfolder.Folders, Name)
      If Not FindInFolders Is Nothing Then Exit For
    End If
  Next
End Function
Function FolderExist(objMyfolder As Outlook.Folders, Name As String) As Boolean

Dim bVal As Boolean
bVal = False
Dim i As Integer
'Loop through outlook folder to search for specific folder
For i = 1 To objMyfolder.Count
    'Check if folder exist in loop
    If objMyfolder.Item(i).Name = Name Then
        Debug.Print "Folder " & objMyfolder.Item(i).Name & " exists"
        bVal = True
        Exit For
    End If
Next i

'Return the resutls boolean
FolderExist = bVal

End Function
Function GetFolderPath(ByVal FolderPath As String) As Outlook.Folder
    Dim oFolder As Outlook.Folder
    Dim FoldersArray As Variant
    Dim i As Integer
        Debug.Print "GetFolderPaht for " & FolderPath
    On Error GoTo GetFolderPath_Error
    If Left(FolderPath, 2) = "\\" Then
        FolderPath = Right(FolderPath, Len(FolderPath) - 2)
    End If
    'Convert folderpath to array
    FoldersArray = Split(FolderPath, "\")
    Set oFolder = Application.Session.Folders.Item(FoldersArray(0))
    If Not oFolder Is Nothing Then
        For i = 1 To UBound(FoldersArray, 1)
            Dim SubFolders As Outlook.Folders
            Set SubFolders = oFolder.Folders
            Set oFolder = SubFolders.Item(FoldersArray(i))
            If oFolder Is Nothing Then
                Set GetFolderPath = Nothing
            End If
        Next
    End If
    'Return the oFolder
    Set GetFolderPath = oFolder
    Exit Function
        
GetFolderPath_Error:
    Set GetFolderPath = Nothing
    Exit Function
End Function


Public Sub FindFolder(ByVal Name As String)

  'Dim Name$
  Dim Folders As Outlook.Folders

  Set m_Folder = Nothing
  m_Find = ""
  m_Wildcard = False

  'Name = InputBox("Find name:", "Search folder")
  If Len(Trim$(Name)) = 0 Then Exit Sub
  m_Find = Name

  m_Find = LCase$(m_Find)
  m_Find = Replace(m_Find, "%", "*")
  m_Wildcard = (InStr(m_Find, "*"))

  Set Folders = Application.Session.Folders 'Only look in the Main MAPI Folder
  Set Folders = Folders.Item(2).Folders 'Only look in the inbox folder

  LoopFolders Folders

  If Not m_Folder Is Nothing Then
    'If MsgBox("Activate folder: " & vbCrLf & m_Folder.FolderPath, vbQuestion Or vbYesNo) = vbYes Then
      Set Application.ActiveExplorer.CurrentFolder = m_Folder
    'End If
  Else
    MsgBox "Not found", vbInformation
  End If
End Sub
Private Sub LoopFolders(Folders As Outlook.Folders)
  Dim F As Outlook.MAPIFolder
  Dim Found As Boolean
  
  If SpeedUp = False Then DoEvents

  For Each F In Folders
    If m_Wildcard Then
      Found = (LCase$(F.Name) Like m_Find)
    Else
      Found = (LCase$(F.Name) = m_Find)
    End If

    If Found Then
      If StopAtFirstMatch = False Then
        If MsgBox("Found: " & vbCrLf & F.FolderPath & vbCrLf & vbCrLf & "Continue?", vbQuestion Or vbYesNo) = vbYes Then
          Found = False
        End If
      End If
    End If
    If Found Then
      Set m_Folder = F
      Exit For
    Else
      LoopFolders F.Folders
      If Not m_Folder Is Nothing Then Exit For
    End If
  Next
End Sub
