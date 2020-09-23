Attribute VB_Name = "modMain"
Option Explicit

' Public Data Members                                                         '
' =============================================================================

Public ready As Boolean

' Methods                                                                     '
' =============================================================================

' Method main:
'
'   Invoked by windows when the application begins.  This method will determine
'   the action taken by the program, based on the command line.  Only the first
'   of these values will be examined and used--thus, if you specify more than
'   one, the subsequent switches will be ignored.
'
'   Command line options:
'
'       None:           If no command line is specified, the default behavior
'                       of this program is to display a list of installed FSL
'                       packets to select for uninstallation.
'
'       /install x      Performs an installation based on the data in the FSL
'                       file x.
'
'       /uninstall x    Performs an uninstallation based on the data in the
'                       FSU file x.
'
'       /list x         Shows a list of previous installations related to the
'                       UID stored in the FSL file x and allows the user to
'                       select one to uninstall.
'
'       /edit [x]       Edits or creates an FSL file.  If the FSL file x is
'                       omitted, then FSLMan creates a new FSL file.
'
Public Sub Main()
    Dim cmdline As String, sppos As Integer, file As String
    Dim fsl As New FileScriptFSL, answer As VbMsgBoxResult
    
    cmdline = Trim(LCase(Command))
    sppos = InStr(1, cmdline, " ")
    
    If sppos <> 0 Then file = Mid(cmdline, sppos + 1)
    
    If cmdline = "" Then
        ' No command--show the global uninstall list.
        showUninstallList ""
    ElseIf Left(cmdline, Len("/install")) = "/install" Then
        ' Begin installing.
        installFSL file
    ElseIf Left(cmdline, Len("/uninstall")) = "/uninstall" Then
        ' Begin uninstalling.
        uninstallFSL file
    ElseIf Left(cmdline, Len("/list")) = "/list" Then
        ' Show the uninstall list for the given file.  We'll need to open the
        ' file and grab its UID.
        On Error Resume Next
        fsl.readFSL file
        On Error GoTo 0
        
        If Err.Number <> 0 Or fsl.uid = "" Then
            ' Can't get an uninstall list without a UID.
            answer = MsgBox("A list of previous installations cannot be " & _
                            "found for the install file """ & file & ".""" & _
                            vbNewLine & vbNewLine & _
                            "Would you like to see a list of all " & _
                            "previous installations?", _
                            vbExclamation + vbYesNo, "FSL Installations")
                            
            If answer = vbYes Then
                showUninstallList ""
            Else
                End
            End If
        Else
            showUninstallList fsl.uid
        End If
    ElseIf Left(cmdline, Len("/edit")) = "/edit" Then
        ' Edit the given FSL file or create a new one.
        editFSL file
    Else
        ' Unrecognized.
        MsgBox "Invalid command line: " & Command, vbExclamation, "FSLMan"
        End
    End If
    
    End
End Sub

' Method editFSL:
'
'   Edits the given FSL file or creates a new FSL file.
'
' Parameters:
'
'   fsl:                The FSL file to edit.  If empty (""), then a new FSL
'                       file will be created.
'
Public Sub editFSL(fsl As String)
    If fsl <> "" Then
        Load frmEdit
        frmEdit.openFSL fsl
    End If

    frmEdit.Show
    
    ' Keep looping until we should quit.
    Do
        DoEvents
    Loop Until ready
End Sub

' Method showUninstallList:
'
'   Displays the uninstall list for the given UID, or the global uninstall list
'   if the UID is not provided.
'
' Parameters:
'
'   [uid]:              The UID to display the uninstall list for.
'
Public Sub showUninstallList(Optional uid As String)
    Dim files() As String, listForm As frmUninstallListing, i As Integer
    Dim newFSLU As FileScriptFSLUninstall
    
    files = FileScript.getFSLUninstallLogs(uid)
    
    If UBound(files) = 0 And files(0) = "" Then
        ' No uninstall logs.
        MsgBox "There are no previous installations to show.", vbExclamation
        End
    Else
        ' Set up the list on the main form.
        Set listForm = New frmUninstallListing
        Load listForm
        
        If uid = "" Then
            listForm.Caption = "Previous Installations"
        Else
            listForm.Caption = "Previous Installations of " & uid
        End If
        
        listForm.numLogFiles = UBound(files) + 1
        
        For i = 0 To UBound(files)
            Set newFSLU = New FileScriptFSLUninstall
            
            On Error GoTo showUninstallList_nextFile
            newFSLU.Load files(i)
            On Error GoTo 0
            
            Set listForm.logFile(i) = newFSLU
showUninstallList_nextFile:
        Next
        
        On Error GoTo 0
        
        listForm.refreshLogs
        listForm.Show vbModal
    End If
End Sub

' Method installFSL:
'
'   Begins the installation process for the given FSL.
'
' Parameters:
'
'   fsl:                The FSL file to install.
'
Public Sub installFSL(fsl As String)
    Dim fslObj As New FileScriptFSL, installer As New frmInstall
    Dim overrideIngoreFSLErrors As Boolean
    
    ' Attempt to open the FSL file.
    On Error GoTo installFSL_error
    
    If fslObj.readFSL(fsl) Then
        On Error GoTo 0
        
        overrideIngoreFSLErrors = _
         mbiFunctions.getINISetting(FileScript.getFileScriptININame, _
                                    "FSL Install", _
                                    "override ignore fsl errors", _
                                    False)
    
        If Not overrideIngoreFSLErrors And fslObj.numFSLErrors > 0 And _
         Not fslObj.ignoreFSLErrors Then
            MsgBox "There is a problem with the selected installation " & _
                   "that is preventing it from being displayed.  The " & _
                   "installation cannot continue." & vbNewLine & _
                   vbNewLine & _
                   "Error: " & fslObj.fslError(0), _
                   vbCritical, "Install"
        ElseIf fslObj.numCopyFiles = 0 And fslObj.numScripts = 0 And _
         fslObj.numSources = 0 Then
            MsgBox "The selected FSL does not specify any files to " & _
                   "install.  The installation cannot continue.", _
                   vbCritical, "Install"
        Else
            ' Set the current path, based on the filename of the fsl.
            If InStr(1, fsl, ":") <> 0 Then
                ChDrive Left(fsl, InStr(1, fsl, ":"))
            End If
            
            If InStr(1, fsl, "\") <> 0 Then
                ChDir extractPathname(fsl)
            End If
        
            ' Load and show the installer.
            Load installer
            Set installer.fsl = fslObj
            installer.initInstall
            installer.Hide
            installer.Show vbModal
        End If
    Else
installFSL_error:
        MsgBox "The selected FSL file, """ & fsl & ","" could not be " & _
               "opened.  The FSL was not installed.", vbExclamation, "Install"
        End
    End If
End Sub

' Method uninstallFSL:
'
'   Begins the uninstallation process based on the settings in the given
'   uninstall log file.
'
' Parameters:
'
'   fsu:                The FileScript Uninstall Log file to use.
'
Public Sub uninstallFSL(fsu As String)
    Dim fslu As FileScriptFSLUninstall, uninstaller As frmUninstall
    
    ' Attempt to open the FSL file.
    Set fslu = New FileScriptFSLUninstall
    On Error GoTo uninstallFSL_cantOpen
    fslu.Load fsu
    On Error GoTo 0
    
    ' All good.  Load the uninstall form.
    Set uninstaller = New frmUninstall
    Load uninstaller
    
    Set uninstaller.uninstallLog = fslu
    uninstaller.setup
    uninstaller.Show vbModal
    
    Exit Sub
    
uninstallFSL_cantOpen:
    ' Couldn't open the FSU file.
    MsgBox "The uninstall log """ & fsu & """ could not be opened.", _
           vbExclamation
           
    Exit Sub
End Sub
