VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private m_Folder As Outlook.MAPIFolder
Private m_Find As String
Private m_Wildcard As Boolean

Private Const SpeedUp As Boolean = True
Private Const StopAtFirstMatch As Boolean = True

'********************************************************************************
'   Author: Al Ramos
'   Name: CreateFolderFromXLS
'   Description: Script written to create parent child folder structure in outlook via csv file.
'   Date: 01/22/2018
'   Updated: 01/05/2019
'
'   Description: Creaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa
'   Reference Sample URL #1 = https://www.slipstick.com/developer/code-samples/create-outlook-folders-list-folder-names/
'   Reference Sample URL #2 = http://www.outlookcode.com/threads.aspx?forumid=2&messageid=33589
'********************************************************************************
Public Sub CreateFolderFromXLS()
On Error GoTo PROC_ERR
    Dim CurrentFolder As Outlook.MAPIFolder
    Dim Subfolder As Outlook.MAPIFolder
    Dim List As New VBA.Collection
    Dim Folders As Outlook.Folders
    Dim Item As Variant

    Dim objParentFolder, objPriorParentFolder As Outlook.Folder ' parent
    Dim newFolderName, parentname 'As String
    Dim strFilepath
    
    Dim xlApp As Object 'Excel.Application
    Dim xlWkb As Object ' As Workbook
    Dim xlSht As Object ' As Worksheet
    Dim rng As Object 'Range
    
    Dim i, iRow, iCol As Integer
    iRow = 2 'Start from line 2, first line is column header
    iCol = 1

    Dim FoundFolder As Folder

    Set xlApp = CreateObject("Excel.Application")
    
    'strFilepath = xlApp.GetOpenFilename("Comma Seperated Values (*.csv), *.csv")
    strFilepath = "C:\Temp\PS\Outlook\PSFT_FolderStructure.csv"
    If strFilepath = False Then
        xlApp.Quit
        Set xlApp = Nothing
        Exit Sub
    End If
      
    Set xlWkb = xlApp.Workbooks.Open(strFilepath)
    Set xlSht = xlWkb.Worksheets(1)
    
    Set objParentFolder = Session.GetDefaultFolder(olFolderInbox)
    Set Folders = objParentFolder.Folders
    Dim olFolder As Outlook.MAPIFolder

    Dim MaxLevel, iCell As Integer
    MaxLevel = xlSht.Range(xlSht.Cells(2, 1), xlSht.Cells(2, 2).End(-4121)).Count
    Set rng = xlSht.Range(xlSht.Cells(2, 1), xlSht.Cells(2, 2).End(-4121))

    Dim cellPar, PriorcellPar As String
    
    Set Application.ActiveExplorer.CurrentFolder = objParentFolder 'Set to inbox folder
    
    'If there are no childeren, there is no need to execute further
    If Not MaxLevel = 0 Then
    iCell = 0
    
    Set objParentFolder = Application.ActiveExplorer.CurrentFolder
    Dim objNewFolder As Outlook.Folder
                 
    While xlSht.Cells(iRow, 1) <> ""
    
        For iCol = 1 To 2
        
            parentname = xlSht.Cells(iRow, 1)
            newFolderName = xlSht.Cells(iRow, 2)
                                       
            On Error Resume Next
            
            If parentname = "Inbox" And iCol = 1 Then
                Set objParentFolder = Session.GetDefaultFolder(olFolderInbox)
            Else
            
                If parentname <> newFolderName Then
                    If Application.ActiveExplorer.CurrentFolder.Folders.Count = 0 Then 'No SubFolders exist
                        Set objNewFolder = Application.ActiveExplorer.CurrentFolder.Folders.Add(newFolderName)
                        Debug.Print "Add  New Folder " & newFolderName & " Created"
                    ElseIf Application.ActiveExplorer.CurrentFolder.Folders.Count = 1 And Application.ActiveExplorer.CurrentFolder.Folders.Item(1) = newFolderName Then
                        
                        Set Application.ActiveExplorer.CurrentFolder = objParentFolder.Item(CStr(newFolderName)) 'Worked Created SubFolder
                        Debug.Print "Set Application CurrentFolder = " & objParentFolder.Item(CStr(newFolderName)).FolderPath
                    
                    ElseIf Not FolderExist(Application.ActiveExplorer.CurrentFolder.Folders, CStr(newFolderName)) Then
                    
                        Set Application.ActiveExplorer.CurrentFolder = objParentFolder.Item(CStr(parentname)) 'Worked Created SubFolder
                        Debug.Print "Set Application CurrentFolder = \\Alberto.Ramos@miamidade.gov\Inbox\" & parentname
                        
                        Set objParentFolder = objParentFolder.Folders(CStr(parentname)) 'Worked Created SubFolder
                        Set objParentFolder = Application.ActiveExplorer.CurrentFolder.Folders 'Worked Created SubFolder
                                              
                        Set objNewFolder = Application.ActiveExplorer.CurrentFolder.Folders.Add(newFolderName)
                        
                        Debug.Print " New Folder: " & newFolderName & ", Parent Folder: " & parentname
                    Else
                        Debug.Print "Prep Folder for next Iteration: " & "\\Alberto.Ramos@miamidade.gov\Inbox\" & newFolderName
                        Debug.Print "Source ColNum" & iCol
                        Debug.Print "Source RowNum" & iRow
                        Debug.Print "newFolderName:" & newFolderName
                        Debug.Print "parentname:" & parentname
                        Set Application.ActiveExplorer.CurrentFolder = GetFolderPath("\\Alberto.Ramos@miamidade.gov\Inbox\" & newFolderName)
    
                    End If
                                    
                    iCell = 1
                    
                ElseIf xlSht.Cells(iRow, 1) = objParentFolder.Name Then
                    'Parent column matches Parent Object Name movenext keep parent name?
                    Debug.Print xlSht.Cells(iRow, 1)
                    Debug.Print objParentFolder.Name
                End If
            
            End If 'End Inbox Object Assignment
            
        Next iCol 'Goto Next Field(Column)
        
            iRow = iRow + 1
    
    Wend
         
PROC_EXIT:
  Exit Sub

PROC_ERR:

  MsgBox "Error " & Err.Number & " " & Err.Description
  If Err.Description <> "" Then
    If IsObject(xlWkb) Then
        'xlWkb.Close
        xlApp.Quit
    End If
    Set xlWkb = Nothing
    Set xlApp = Nothing
    Set objParentFolder = Nothing
    Err.Clear
    Resume PROC_EXIT
  End If

    xlWkb.Close
    xlApp.Quit
    Set xlWkb = Nothing
    Set xlApp = Nothing
    Set objParentFolder = Nothing
    End If

End Sub
