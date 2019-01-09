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
'   Get-Process | where ProcessName -like EXCEL | stop-process
'********************************************************************************
Public Sub CreateFolderFromXLS()
On Error GoTo PROC_ERR
    Dim CurrentFolder As Outlook.MAPIFolder
    Dim Subfolder As Outlook.MAPIFolder
    Dim List As New VBA.Collection
    Dim Item As Variant
    
    Dim objOutlook As Object, objnSpace As Object, objFolder As Object
    
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
    Dim iVal As Integer
    iVal = 1
    
    Set objParentFolder = Application.ActiveExplorer.CurrentFolder
    Dim objNewFolder As Outlook.Folder
    

    'Dim ParentMemberNum(1)
    
    
    While xlSht.Cells(iRow, 1) <> ""
    
        For iCol = 1 To 2
        
            sParentName = xlSht.Cells(iRow, 1)
            sNewFolderName = xlSht.Cells(iRow, 2)
                                       
            On Error Resume Next
            
            If sParentName = "Inbox" And iCol = 1 Then
                Set objParentFolder = Session.GetDefaultFolder(olFolderInbox)
                ReDim ParentMemberNum(1)
                ParentMemberNum(1) = sParentName
            Else
                         
                If sParentName <> sNewFolderName Then
                
                    If sParentName <> ParentMemberNum(iVal) And iCol = 1 And iRow > 2 And Not IsInArray(CStr(sParentName), ParentMemberNum) Then
                        'If Not IsInArray(CStr(sParentName), ParentMemberNum) Then
                            iVal = iVal + 1
                            ReDim Preserve ParentMemberNum(iVal)
                            ParentMemberNum(iVal) = sParentName
                            
                            ReDim ChildMemberNum(iVal)
                            ChildMemberNum(iVal) = sNewFolderName
                        'End If
                    End If
                
                    If Application.ActiveExplorer.CurrentFolder.Folders.Count = 0 Then 'No SubFolders exist


                        Set objNewFolder = Application.ActiveExplorer.CurrentFolder.Folders.Add(sNewFolderName)
                        Debug.Print "create New Folder " & sNewFolderName & " Objected Created"
'MsgBox ("New Folder Created: " & sNewFolderName)
                    
                    ElseIf Application.ActiveExplorer.CurrentFolder.Folders.Item(1) = sNewFolderName And iCol = 1 Then
                                            
                            Set objParentFolder = objParentFolder.Folders(CStr(sNewFolderName)) 'Worked Created SubFolder
                            Set objParentFolder = Application.ActiveExplorer.CurrentFolder.Folders 'Worked Created SubFolder
                            Set Application.ActiveExplorer.CurrentFolder = objParentFolder.Item(CStr(sNewFolderName)) 'Worked Created SubFolder - Changed the side Folder Preview
                                           
                    ElseIf Application.ActiveExplorer.CurrentFolder.Folders.Count = 1 And Application.ActiveExplorer.CurrentFolder.Folders.Item(1) = sNewFolderName Then
                        If CStr(xlSht.Cells(iRow, 1).Value) = CStr(ChildMemberNum(iVal)) Then 'Current cell value comparison - This will never happen??????
                            
                            Set Application.ActiveExplorer.CurrentFolder = objParentFolder.Item(CStr(sNewFolderName)) 'Worked Created SubFolder
                            Debug.Print "Set Application CurrentFolder = " & objParentFolder.Item(CStr(sNewFolderName)).FolderPath
                            Debug.Print "Set Application CurrentFolder = " & Application.ActiveExplorer.CurrentFolder.Folders.Item(1)
                        
                            Set objParentFolder = objParentFolder.Folders(CStr(sNewFolderName))  'Worked Created SubFolder
                            Set Application.ActiveExplorer.CurrentFolder = objParentFolder.Item(CStr(sNewFolderName)) 'Worked Created SubFolder - Changed the side Folder Preview
                            Set Application.ActiveExplorer.CurrentFolder = Application.ActiveExplorer.CurrentFolder.Folders.Item(1) 'Worked Created SubFolder - Changed the side Folder Preview - 2nd Time
                            Debug.Print "Folder Changed view! to " & sNewFolderName
                            
                        ElseIf CStr(xlSht.Cells(iRow + 1, 1).Value) = CStr(ChildMemberNum(iVal)) Then 'check current leaf with next Parent member
                            If Application.ActiveExplorer.CurrentFolder.Folders.Count = 1 Then
                                Debug.Print "Set Application CurrentFolder = " & Application.ActiveExplorer.CurrentFolder.Folders.Item(1) 'This works and keep 1/08/2019 11:21AM
                                Set Application.ActiveExplorer.CurrentFolder = Application.ActiveExplorer.CurrentFolder.Folders.Item(1) 'Worked Created SubFolder
                            End If
                        ElseIf CStr(xlSht.Cells(iRow + 1, 1).Value) <> CStr(ParentMemberNum(iVal)) Then 'check Future parent member does not match Current Parent member
                            Debug.Print "if Future Parent does not equals with Current Parent. do nothing keep Folder Selection"
                        ElseIf CStr(xlSht.Cells(iRow + 1, 1).Value) = CStr(ParentMemberNum(iVal)) Then 'check Future parent member matches Current Parent member
                            'Do Nothing we are cool
                            Debug.Print CStr(xlSht.Cells(iRow + 1, 1).Value) & " = " & CStr(ParentMemberNum(iVal))
                        ElseIf CStr(xlSht.Cells(iRow + 1, 2).Value) <> CStr(ChildMemberNum(iVal)) And CStr(xlSht.Cells(iRow + 1, 1).Value) = CStr(ParentMemberNum(iVal)) Then  'check parent leaf to current leaf source
                            Debug.Print "if Future child equals with Current Parent. do nothing keep Folder Selection"

                        ElseIf CStr(xlSht.Cells(iRow + 1, 2).Value) <> CStr(ChildMemberNum(iVal)) And CStr(xlSht.Cells(iRow + 1, 1).Value) <> CStr(ParentMemberNum(iVal)) Then  'check parent leaf to current leaf source
                            Debug.Print "if Future child do not equal with Current child and Parent dont equal.Then make next parent focused"
                        End If
                    
                    ElseIf Not FolderExist(Application.ActiveExplorer.CurrentFolder.Folders, CStr(sNewFolderName)) Then
                    
                        
                        If iRow > 2 Then
                         
                            If iCol = 1 Then 'Make sure we have a leaf member to for creating folders

                                If CStr(ParentMemberNum(iVal)) <> sNewFolderName And ParentMemberNum(iVal - 1) = "Inbox" Then
                                    
                                    Set objParentFolder = objParentFolder.Folders(CStr(ParentMemberNum(iVal)))  'Worked Created SubFolder
                                    Set Application.ActiveExplorer.CurrentFolder = objParentFolder.Item(CStr(sParentName)) 'Worked Created SubFolder - Changed the side Folder Preview
                                    Debug.Print "Folder Changed view! to " & ParentMemberNum(iVal)
                                    
                                ElseIf CStr(ParentMemberNum(iVal)) <> sNewFolderName And ParentMemberNum(iVal - 1) = sParentName And Application.ActiveExplorer.CurrentFolder.Name <> sParentName Then
                                    'Then go up to the parent of current member
                                    If objParentFolder.Item(5).Folders.Count = 1 And sParentName = objParentFolder.Item(5).Folders.Item(1) Then
                                        Set Application.ActiveExplorer.CurrentFolder = objParentFolder.Item(5).Folders.Item(1) 'go up a couple folders to reach parent folder, to continue creating second level
                                    Else
                                        Debug.Print sParentName
                                    End If
                                        
                                    DoEvents
                                    

                                ElseIf CStr(ParentMemberNum(iVal - 1)) <> Application.ActiveExplorer.CurrentFolder.Name Then 'Prior Member does not match current selected folder
                                    

                                    Debug.Print "Current Selected folder " & Application.ActiveExplorer.CurrentFolder.Name
                                    Debug.Print "lets focus on the last parent member " & ParentMemberNum(iVal)
                                    If objParentFolder.Item(5).Folders.Count > 1 Then
                                    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ Commend out first line; to better improve iteration. 1/9/2019 8:15AM @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
                                        'If ParentMemberNum(iVal) = objParentFolder.Item(objParentFolder.Count).Folders.Item(objParentFolder.Item(5).Folders.Count).Folders.Item(objParentFolder.Item(objParentFolder.Count).Folders.Item(objParentFolder.Item(5).Folders.Count).Folders.Count).Name Then
                                        'If ParentMemberNum(iVal) = objParentFolder.Item(objParentFolder.Count).Folders.Item(objParentFolder.Item(5).Folders.Count).Folders.Item(objParentFolder.Item(objParentFolder.Count).Folders.Item(objParentFolder.Item(5).Folders.Count).Folders.Count).Name And CStr(xlSht.Cells(iRow - 1, 2).Value) = CStr(xlSht.Cells(iRow, 1).Value) Then
                                        If CStr(xlSht.Cells(iRow - 1, 2).Value) = CStr(xlSht.Cells(iRow, 1).Value) Then
                                            
                                            Set Application.ActiveExplorer.CurrentFolder = objParentFolder.Item(5).Folders.Item(1).Folders.Item(2) 'go up one to the root folder childer folder to reset structure
                                            
                                        End If
                                    ElseIf objParentFolder.Item(5).Folders.Count = 1 And ParentMemberNum(iVal) <> Application.ActiveExplorer.CurrentFolder And CStr(xlSht.Cells(iRow, 2).Value) <> CStr(ChildMemberNum(iVal)) Then
                                        
                                        Set Application.ActiveExplorer.CurrentFolder = objParentFolder.Item(5).Folders.Item(1) 'go up two to the root folder childer folder to reset structure
                                        
                                    ElseIf objParentFolder.Item(5).Folders.Count = 1 And ParentMemberNum(iVal) = Application.ActiveExplorer.CurrentFolder Then
                                    
                                        Debug.Print "Do Nothing we kewl for next child member to be added to parent"
                                        Debug.Print Application.ActiveExplorer.CurrentFolder
                                        
                                    ElseIf objParentFolder.Item(5).Folders.Count = 1 And CStr(xlSht.Cells(iRow - 1, 2).Value) = CStr(xlSht.Cells(iRow, 1).Value) Then
                                        'If Child value equals current Parent value then, We Need to focus/set to the right folder - next Division"
                                        Debug.Print "We Need to set to the right folder - next Division"
                                        Debug.Print Application.ActiveExplorer.CurrentFolder
                                        
                                        Set objParentFolder = objParentFolder.Folders(CStr(ParentMemberNum(iVal))) 'Worked Created SubFolder
                                        Set Application.ActiveExplorer.CurrentFolder = objParentFolder
                                        Set objParentFolder = Application.ActiveExplorer.CurrentFolder 'Worked Created SubFolder
                                            
                                        Set objParentFolder = objParentFolder.Folders(CStr(ParentMemberNum(iVal))) 'Worked Created SubFolder
                                        Set Application.ActiveExplorer.CurrentFolder = objParentFolder
                                        Set objParentFolder = Application.ActiveExplorer.CurrentFolder 'Worked Created SubFolder
                                        
                                    End If
                                Else
                                    Set objParentFolder = objParentFolder.Folders(CStr(ParentMemberNum(iVal))) 'Worked Created SubFolder
                                    Set Application.ActiveExplorer.CurrentFolder = objParentFolder
                                    Set objParentFolder = Application.ActiveExplorer.CurrentFolder 'Worked Created SubFolder
                           
                                    DoEvents
                                End If
                            Else
                                Set objParentFolder = objParentFolder.Item(CStr(ParentMemberNum(iVal))) 'Worked Created SubFolder
                            End If
                        Else

                            Set objParentFolder = Application.ActiveExplorer.CurrentFolder.Folders 'Worked Created SubFolder
    
                            Set Application.ActiveExplorer.CurrentFolder = objParentFolder.Item(CStr(sParentName)) 'Worked Created SubFolder - Changed the side Folder Preview
                            Set Application.ActiveExplorer.CurrentFolder = GetFolderPath(objParentFolder.Item(CStr(sParentName)).FolderPath)
    
                            Set Application.ActiveExplorer.CurrentFolder = FindInFolders(objParentFolder.Item(CStr(sParentName)), sParentName)
                            Set Application.ActiveExplorer.CurrentFolder = objParentFolder
                        
                        End If
                        
                            Debug.Print "Source ColNum: " & iCol
                            Debug.Print "Source RowNum: " & iRow
                            Debug.Print "newFolderName: " & sNewFolderName
                            Debug.Print "parentname: " & sParentName

                            If iRow > 2 And Application.ActiveExplorer.CurrentFolder = "Inbox" Then
                                Set objParentFolder = Application.ActiveExplorer.CurrentFolder.Folders 'Worked Created SubFolder
                                Set Application.ActiveExplorer.CurrentFolder = objParentFolder.Item(CStr(sParentName)) 'Worked Created SubFolder - Changed the side Folder Preview
                                Set Application.ActiveExplorer.CurrentFolder = GetFolderPath(objParentFolder.Item(CStr(sParentName)).FolderPath)
        
                                Set Application.ActiveExplorer.CurrentFolder = FindInFolders(objParentFolder.Item(CStr(sParentName)), sParentName)
                                Set Application.ActiveExplorer.CurrentFolder = objParentFolder
                            End If
                            Set objNewFolder = Application.ActiveExplorer.CurrentFolder.Folders.Add(sNewFolderName)
                            'Application.ActiveExplorer.CurrentFolder = objParentFolder.Item(5).Folders.Item(1)

'MsgBox ("New Folder Created: " & sNewFolderName)

                    Else 'Folder exist
                        Debug.Print "Prep Folder for next Iteration: " & objParentFolder.Item(CStr(sNewFolderName)).FolderPath
                        Debug.Print "Source ColNum: " & iCol
                        Debug.Print "Source RowNum: " & iRow
                        Debug.Print "newFolderName: " & sNewFolderName
                        Debug.Print "parentname: " & sParentName

                    End If

                    iCell = 1

                ElseIf xlSht.Cells(iRow, 1) = objParentFolder.Item(CStr(sParentName)) Then
                    'Parent column matches Parent Object Name movenext keep parent name?
                    Debug.Print xlSht.Cells(iRow, 1)
                    Debug.Print objParentFolder.Item(CStr(sParentName))
                
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
        xlWkb.Close
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
