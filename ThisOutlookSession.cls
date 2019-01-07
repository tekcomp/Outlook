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
    Dim Item As Variant

    Dim objParentFolder, objPriorParentFolder As Outlook.Folder ' parent
    Dim sNewFolderName, sParentName 'As String
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
    

    Dim ParentMemberNum(1) As Double
    'ReDim Preserve ParentMemberNum(MaxLevel)
    
    While xlSht.Cells(iRow, 1) <> ""
    
        For iCol = 1 To 2
        
            sParentName = xlSht.Cells(iRow, 1)
            sNewFolderName = xlSht.Cells(iRow, 2)
                                       
            On Error Resume Next
            
            If sParentName = "Inbox" And iCol = 1 Then
                Set objParentFolder = Session.GetDefaultFolder(olFolderInbox)
                'ParentMemberNum(iRow) = sParentName
            Else
                         
                'If sParentName <> ParentMemberNum(iRow) Then
                    'ParentMemberNum(iRow) = sParentName
                'End If
                         
                If sParentName <> sNewFolderName Then
                
                    If Application.ActiveExplorer.CurrentFolder.Folders.Count = 0 Then 'No SubFolders exist
                        Set objNewFolder = Application.ActiveExplorer.CurrentFolder.Folders.Add(sNewFolderName)
                        Debug.Print "Add  New Folder " & sNewFolderName & " Created"
                    ElseIf Application.ActiveExplorer.CurrentFolder.Folders.Item(1) = sNewFolderName And iCol = 2 Then
                    
                        If sParentName <> Application.ActiveExplorer.CurrentFolder.Folders.Item(1) Then
                            'go back up to parent folder for continued iteration of childern folders
                            Set objParentFolder = objParentFolder.Folders(CStr(sParentName)) 'Worked Created SubFolder
                            Set objParentFolder = Application.ActiveExplorer.CurrentFolder.Folders 'Worked Created SubFolder
                            Set Application.ActiveExplorer.CurrentFolder = objParentFolder.Item(CStr(sParentName)) 'Worked Created SubFolder - Changed the side Folder Preview
                        Else
                        
                            'Current Folder Iteration equals current Outlook Folder
                            Set objParentFolder = objParentFolder.Folders(CStr(sNewFolderName)) 'Worked Created SubFolder
                            Set objParentFolder = Application.ActiveExplorer.CurrentFolder.Folders 'Worked Created SubFolder
                            Set Application.ActiveExplorer.CurrentFolder = objParentFolder.Item(CStr(sNewFolderName)) 'Worked Created SubFolder - Changed the side Folder Preview
                        
                        End If
                        
                   
                    ElseIf Application.ActiveExplorer.CurrentFolder.Folders.Count = 1 And Application.ActiveExplorer.CurrentFolder.Folders.Item(1) = sNewFolderName Then

                        Set Application.ActiveExplorer.CurrentFolder = objParentFolder.Item(CStr(sNewFolderName)) 'Worked Created SubFolder
                        Debug.Print "Set Application CurrentFolder = " & objParentFolder.Item(CStr(sNewFolderName)).FolderPath
                        Debug.Print "Set Application CurrentFolder = " & Application.ActiveExplorer.CurrentFolder.Folders.Item(1)
                    
                    ElseIf Not FolderExist(Application.ActiveExplorer.CurrentFolder.Folders, CStr(sNewFolderName)) Then
                    
                        Set objParentFolder = objParentFolder.Folders(CStr(sParentName)) 'Worked Created SubFolder
                        Set objParentFolder = Application.ActiveExplorer.CurrentFolder.Folders 'Worked Created SubFolder
                        Set Application.ActiveExplorer.CurrentFolder = objParentFolder.Item(CStr(sParentName)) 'Worked Created SubFolder - Changed the side Folder Preview
                        Set Application.ActiveExplorer.CurrentFolder = GetFolderPath(objParentFolder.Item(CStr(sParentName)).FolderPath)
                        
                        Set Application.ActiveExplorer.CurrentFolder = FindInFolders(objParentFolder.Item(CStr(sParentName)), sParentName)
                        
                        Set objNewFolder = Application.ActiveExplorer.CurrentFolder.Folders.Add(sNewFolderName)
                        
                        Debug.Print "Parent Folder: " & sParentName & ", Child Folder: " & sNewFolderName
                                               
                    Else 'Folder exist
                        Debug.Print "Prep Folder for next Iteration: " & objParentFolder.Item(CStr(sNewFolderName)).FolderPath
                        Debug.Print "Source ColNum: " & iCol
                        Debug.Print "Source RowNum: " & iRow
                        Debug.Print "newFolderName: " & sNewFolderName
                        Debug.Print "parentname: " & sParentName
                        'Set Application.ActiveExplorer.CurrentFolder = GetFolderPath("\\Alberto.Ramos@miamidade.gov\Inbox\" & sNewFolderName)
                        'Set Application.ActiveExplorer.CurrentFolder = GetFolderPath(objParentFolder.Item(CStr(sNewFolderName)).FolderPath)

                        'Set objParentFolder = objParentFolder.Folders(CStr(sNewFolderName)) 'Worked Created SubFolder
                        'Set objParentFolder = Application.ActiveExplorer.CurrentFolder.Folders 'Worked Created SubFolder
                        'Set Application.ActiveExplorer.CurrentFolder = objParentFolder.Item(CStr(sNewFolderName)) 'Worked Created SubFolder - Changed the side Folder Preview

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
