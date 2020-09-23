VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LineCounter"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   5790
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCount 
      Caption         =   "Count!"
      Default         =   -1  'True
      Height          =   495
      Left            =   3720
      TabIndex        =   19
      Top             =   7440
      Width           =   1935
   End
   Begin VB.CommandButton cmdClear 
      Cancel          =   -1  'True
      Caption         =   "Clear Settings"
      Height          =   495
      Left            =   1800
      TabIndex        =   18
      Top             =   7440
      Width           =   1935
   End
   Begin VB.Frame frameOutput 
      Caption         =   "Output"
      Height          =   2415
      Left            =   120
      TabIndex        =   11
      Top             =   4920
      Width           =   5535
      Begin VB.CheckBox chkShowFileDetails 
         Caption         =   "Show File Details"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   2040
         Value           =   1  'Checked
         Width           =   5175
      End
      Begin VB.CheckBox chkShowFolderDetails 
         Caption         =   "Show Folder Details"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1800
         Value           =   1  'Checked
         Width           =   5175
      End
      Begin VB.CheckBox chkRTF 
         Caption         =   "Output in Rich &Text Format"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1440
         Value           =   1  'Checked
         Width           =   5175
      End
      Begin VB.CommandButton cmdBrowseFile 
         Caption         =   "..."
         Height          =   285
         Left            =   4920
         TabIndex        =   16
         Top             =   1080
         Width           =   495
      End
      Begin VB.OptionButton optFile 
         Caption         =   "Fil&e: (Unselected)"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1080
         Width           =   4695
      End
      Begin VB.OptionButton optClipboard 
         Caption         =   "&Clipboard"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   840
         Width           =   5175
      End
      Begin VB.OptionButton optPrinter 
         Caption         =   "&Printer"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   600
         Width           =   5175
      End
      Begin VB.OptionButton optOutputWindow 
         Caption         =   "&Results Window"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Value           =   -1  'True
         Width           =   5175
      End
   End
   Begin MSComDlg.CommonDialog C 
      Left            =   5040
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frameInclude 
      Caption         =   "Include"
      Height          =   3375
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   5535
      Begin VB.ListBox lstFiles 
         Height          =   1035
         Left            =   1320
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   2160
         Width           =   4095
      End
      Begin VB.ListBox lstDirectories 
         Height          =   1035
         Left            =   1320
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   960
         Width           =   4095
      End
      Begin VB.TextBox txtMasks 
         Height          =   285
         Left            =   1320
         TabIndex        =   6
         Text            =   "*.bas; *.cls; *.txt; *.ctl; *.frm"
         Top             =   360
         Width           =   4095
      End
      Begin VB.Label lblFilesLbl 
         Alignment       =   1  'Right Justify
         Caption         =   "&Files:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label lblDirectoriesLbl 
         Alignment       =   1  'Right Justify
         Caption         =   "&Directories:"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   975
      End
      Begin VB.Line lnSep 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   15
         X2              =   5520
         Y1              =   855
         Y2              =   855
      End
      Begin VB.Line lnSep 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   0
         X2              =   5520
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label lblMasksLbl 
         Alignment       =   1  'Right Justify
         Caption         =   "&Masks:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame frameLocation 
      Caption         =   "Files &Location"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      Begin VB.CheckBox chkRecurseSubdirs 
         Caption         =   "&Recurse Subdirectories"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   5175
      End
      Begin VB.CommandButton cmdBrowseRoot 
         Caption         =   "..."
         Height          =   285
         Left            =   4920
         TabIndex        =   2
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtRoot 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   360
         Width           =   4680
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Private Data Members                                                        '
' =============================================================================

' The file to save the output results to.
Private m_resultsFile As String

' Private and Helper Methods                                                  '
' =============================================================================

' Method addFile:                                                             '
'                                                                             '
'   Adds the file only if another file with the same name and size hasn't     '
'   already been added.                                                       '
'                                                                             '
Private Sub addFile(filename As String)
    Dim i As Integer, searchFile As String, thisFileLen As Long
    
    ' Ignore missing files for now...
    On Error Resume Next
    
    searchFile = UCase(extractFilename(filename))
    thisFileLen = FileLen(filename)
    
    For i = 0 To lstFiles.ListCount - 1
        If UCase(extractFilename(lstFiles.List(i))) = searchFile Then
            ' If they're the same size, don't add it.
            If FileLen(lstFiles.List(i)) = thisFileLen Then
                Exit Sub
            End If
        End If
    Next
    
    ' Go ahead and add it.
    lstFiles.AddItem filename
End Sub

' Method getRootFiles:                                                        '
'                                                                             '
'   Scans the path directory and adds all files that match one of the         '
'   supplied masks.  If the chkRecurseSubdirs is checked, then it will also   '
'   perform the operation for all subdirectories.                             '
'                                                                             '
Private Sub getRootFiles(path As String, masks() As String)
    Dim fso As Object, fldr As Object, file As Object, subDir As Object
    Dim i As Integer
    
    ' Setup the FileSystemObject and folder.
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fldr = fso.GetFolder(path)
    
    ' Add this directory to the list.
    lstDirectories.AddItem addPathBackslash(path)
    
    ' Add each matching file to the list.
    For Each file In fldr.files
        ' Check to see if it matches any of the masks.
        For i = LBound(masks) To UBound(masks)
            If UCase(Trim(extractFilename(file.name))) Like masks(i) Then
                ' Add this file.
                addFile buildPath(path, file.name)
            End If
        Next
    Next
    
    ' Now recurse the subdirectories, if desired.
    If chkRecurseSubdirs.value = 1 Then
        For Each subDir In fldr.subFolders
            getRootFiles subDir.path, masks
        Next
    End If
End Sub

' Method getMasksList:                                                        '
'                                                                             '
'   Returns a string array containing the individual masks.                   '
'                                                                             '
Private Function getMasksList() As String()
    Dim rval() As String, i As Integer
    
    rval = Split(UCase(txtMasks.Text), ";")
    
    For i = LBound(rval) To UBound(rval)
        rval(i) = Trim(rval(i))
    Next
    
    getMasksList = rval
End Function

' Event Click on chkShowFolderDetails:                                        '
'                                                                             '
'   If this checkbox is unchecked, we will uncheck and disable the individual '
'   file details checkbox, since the folders must be reported individually if '
'   the file details are to be shown.                                         '
'                                                                             '
Private Sub chkShowFolderDetails_Click()
    If chkShowFolderDetails.value = 1 Then
        chkShowFileDetails.Enabled = True
    Else
        chkShowFileDetails.Enabled = False
        chkShowFileDetails.value = 0
    End If
End Sub

' Event Click on cmdBrowseFile:                                               '
'                                                                             '
'   Select the file to send the results information to.                       '
'                                                                             '
Private Sub cmdBrowseFile_Click()
    On Error GoTo cmdBrowseFile_Click_Cancel
    
    With C
        .CancelError = True
        .DialogTitle = "Choose Results File"
        .Filter = "Rich Text Format (*.rtf)|*.rtf|Text-Only (*.txt)|*.txt"
        .Flags = cdlOFNNoReadOnlyReturn + cdlOFNOverwritePrompt + _
                 cdlOFNPathMustExist + cdlOFNHideReadOnly
                 
        .ShowSave
        
        m_resultsFile = .filename
        
        If UCase(Right(m_resultsFile, 4)) = ".RTF" Then
            ' Set Rich-Text output.
            chkRTF.value = 1
        Else
            ' Set text output.
            chkRTF.value = 0
        End If
        
        optFile.Caption = "Fil&e: " & .FileTitle
        optFile.value = True
    End With
    
cmdBrowseFile_Click_Cancel:
    ' Canceled.
    Exit Sub
End Sub

' Event Click on cmdBrowseRoot:                                               '
'                                                                             '
'   Have the user browse for a root directory and then load all the files in  '
'   it.                                                                       '
'                                                                             '
Private Sub cmdBrowseRoot_Click()
    Dim selFolder As String

    If Trim(txtMasks.Text) = "" Then _
     txtMasks.Text = "*.cls; *.bas; *.frm; *.ctl; *.txt"

    selFolder = modBrowseForFolder.browseForFolder( _
      "Choose Root Folder", _
      MBFF_O_EDITBOX + MBFF_O_FILESYSTEMDIRSONLY, _
      Me.hWnd)
      
    If selFolder = "" Then Exit Sub Else txtRoot.Text = selFolder
        
    ' Scan the directories...
    Me.MousePointer = vbHourglass
    
    lstDirectories.Clear
    lstFiles.Clear
    getRootFiles txtRoot.Text, getMasksList
    
    Me.MousePointer = vbNormal
End Sub

' Event Click on cmdCount:                                                    '
'                                                                             '
'   Start counting.                                                           '
'                                                                             '
Private Sub cmdCount_Click()
    Dim outputForm As frmOutput

    If lstFiles.ListCount <> 0 Then
        Set outputForm = frmCounting.countLines(Me)
        
        ' Perform extra actions as necessary.
        If chkRTF.value = 0 Then outputForm.plainText = True

        If optClipboard.value = True Then
            outputForm.copy
            
        ElseIf optFile.value = True Then
            outputForm.saveFilename = m_resultsFile
            outputForm.save
        
        ElseIf optPrinter.value = True Then
            outputForm.printResults
        
        End If
    Else
        Beep
    End If
End Sub

' Event Load on Form:                                                         '
'                                                                             '
'   Remember the last form position.                                          '
'                                                                             '
Private Sub Form_Load()
    mbiFunctions.RestoreFormState Me
End Sub

' Event Unload on Form:                                                       '
'                                                                             '
'   Remember the last form position.                                          '
'                                                                             '
Private Sub Form_Unload(cancel As Integer)
    mbiFunctions.SaveFormState Me
    End
End Sub

' Event DblClick on lstFiles:                                                 '
'                                                                             '
'   Show the details about the file they selected.                            '
'                                                                             '
Private Sub lstFiles_DblClick()
    Dim msg As String, selFile As String

    On Error GoTo lstFiles_DblClick_error
    
    selFile = lstFiles.List(lstFiles.ListIndex)
    
    msg = "File details for " & extractFilename(selFile) & ":" & vbNewLine & _
          vbNewLine & _
          "Path: " & extractPathname(selFile) & vbNewLine & _
          "Size: " & FileLen(selFile) & vbNewLine & _
          "Date: " & FileDateTime(selFile)
    
    MsgBox msg, vbInformation
    
    Exit Sub
    
lstFiles_DblClick_error:
    MsgBox "Could not display the information for the selected file.", _
           vbExclamation
           
    Exit Sub
End Sub
