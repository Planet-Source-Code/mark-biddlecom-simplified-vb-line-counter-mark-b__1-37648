VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCounting 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Counting..."
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   1500
      TabIndex        =   2
      Top             =   1440
      Width           =   2295
   End
   Begin MSComctlLib.ProgressBar pbProgress 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lblCurFile 
      Caption         =   "Current File:"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "frmCounting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Private Types                                                               '
' =============================================================================

' This type defines the results for a given file.
Private Type FILECOUNT_RESULTS

    ' The filename.
    filename As String
    
    ' The size of the file, in bytes.
    size As Long
    
    ' This flag determines if the file was actually counted or not (due to
    ' error).
    excluded As Boolean
    
    ' The number of characters in the file, excluding line-break characters.
    numChars As Long

    ' The number of spaces and tabs in the file.
    numWhitespace As Long
    
    ' The total number of lines in this file.
    numLines As Long
    
    ' The number of pure comment lines.
    numCommentLines As Long
    
    ' The total number of lines in this file that were excluded.
    numExcluded As Long
    
End Type

' This type defines the counting results for an individual directory.
Private Type DIRCOUNT_RESULTS
    
    ' The path of this directory, including the final backslash.
    path As String
    
    ' The files included in this directory.
    files() As FILECOUNT_RESULTS
    numFiles As Integer
    
End Type

' Type defines details about the font selection.
Private Type OUTPUTFONT
    name As String
    size As Integer
    bold As Boolean
    italic As Boolean
    underline As Boolean
    color As Long
End Type

' Private Data Members                                                        '
' =============================================================================

' The last set output font.
Private m_lastFont As OUTPUTFONT

' The cancel flags.
Private m_cancel As Boolean
Private m_cancelComplete As Boolean

' The output window.
Private m_outputWindow As frmOutput

' The settings form.
Private m_settingsForm As frmMain

' The list of files to count, and the results of those counted.
Private m_countList() As DIRCOUNT_RESULTS
Private m_numCountList As Integer

' Methods                                                                     '
' =============================================================================

' Method countLines:                                                          '
'                                                                             '
'   Performs the line count and prepares and shows an output window, a        '
'   reference to which is returned.                                           '
'                                                                             '
' Parameters:                                                                 '
'                                                                             '
'   settingsForm:               A reference to the form defining the settings '
'                               for the line count operation.                 '
'                                                                             '
' Returns frmOutput:                                                          '
'                                                                             '
'   A reference to a displayed output form, or Nothing if an error ocurred or '
'   the process was canceled.                                                 '
'                                                                             '
Public Function countLines(settingsForm As frmMain) As frmOutput
    Dim i As Integer, j As Integer

    ' Setup...
    Set m_settingsForm = settingsForm
    initialize
    
    ' Count the data.
    For i = 0 To m_numCountList - 1
        For j = 0 To m_countList(i).numFiles - 1
            countFile m_countList(i).path, m_countList(i).files(j)
            
            ' Cancel if we have to.
            If m_cancel Then cancel: Exit Function
        Next
    Next
    
    ' Display the results.
    cmdCancel.Enabled = False
    Me.MousePointer = vbHourglass
    lblCurFile.Caption = "Generating report..."
    
    generateReport
    
    ' All done.
    Me.MousePointer = vbNormal
    Unload Me
    
    m_outputWindow.Show
    m_settingsForm.Enabled = True
    m_outputWindow.rtfResults.TextRTF = m_outputWindow.rtfResults.TextRTF
    DoEvents
    
    Set countLines = m_outputWindow
    Set m_outputWindow = Nothing
End Function

' Private and Helper Methods                                                  '
' =============================================================================

' Method bytes:                                                               '
'                                                                             '
'   Formats a number, showing it in the largest possible denomination (MB,    '
'   kB, etc.).                                                                '
'                                                                             '
Private Function bytes(ByVal val As Long) As String
    Dim bVal As Double, denom As String
    
    If val < 1000 Then
        ' Just the way it is.
        denom = " Bytes"
        bVal = val
    ElseIf val < 1000000 Then
        ' Kilobytes.
        denom = "kB"
        bVal = val / 1024
    Else
        ' Megabytes.
        denom = "MB"
        bVal = val / 1048576
    End If
    
    bytes = number(bVal) & denom
End Function

' Method number:                                                              '
'                                                                             '
'   Formats a number to a standard form.                                      '
'                                                                             '
Private Function number(ByVal val As Double) As String
    If CLng(val) <> val Then
        number = Format(val, "###,##0.##")
    Else
        number = Format(val, "###,##0")
    End If
End Function

' Method percent:                                                             '
'                                                                             '
'   Calculates and formats a percentage based on the supplied settings.       '
'                                                                             '
Private Function percent(ByVal value As Long, ByVal range As Long) As String
    percent = Format(CDbl(value) / CDbl(range) * 100#, "##0.##") & "%"
End Function

' Method directoryTotalCountedFiles:                                          '
'                                                                             '
'   Counts and returns the total number of files sucessfully processed in the '
'   supplied directory.                                                       '
'                                                                             '
Private Function directoryTotalCountedFiles(dir As DIRCOUNT_RESULTS) As Long
    Dim count As Long, i As Integer
    
    For i = 0 To dir.numFiles - 1
        If Not dir.files(i).excluded Then count = count + 1
    Next
    
    directoryTotalCountedFiles = count
End Function

' Method directoryTotalExcludedFiles:                                         '
'                                                                             '
'   Counts and returns the total number of files that were skipped due to     '
'   error in the supplied directory.                                          '
'                                                                             '
Private Function directoryTotalExcludedFiles(dir As DIRCOUNT_RESULTS) As Long
    Dim count As Long, i As Integer
    
    For i = 0 To dir.numFiles - 1
        If dir.files(i).excluded Then count = count + 1
    Next
    
    directoryTotalExcludedFiles = count
End Function

' Method getDirectoryTotalsCount:                                             '
'                                                                             '
'   Totals the statistics for every file counted in the supplied directory,   '
'   and returns the results in a FILECOUNT_RESULTS type.  The filename and    '
'   excluded properties of the return UDT will not be set.                    '
'                                                                             '
Private Function getDirectoryTotalsCount(dir As DIRCOUNT_RESULTS) _
As FILECOUNT_RESULTS
    Dim results As FILECOUNT_RESULTS, i As Integer
    
    With results
        For i = 0 To dir.numFiles - 1
            .numChars = .numChars + dir.files(i).numChars
            .numCommentLines = .numCommentLines + dir.files(i).numCommentLines
            .numExcluded = .numExcluded + dir.files(i).numExcluded
            .numLines = .numLines + dir.files(i).numLines
            .numWhitespace = .numWhitespace + dir.files(i).numWhitespace
            .size = .size + dir.files(i).size
        Next
    End With
    
    getDirectoryTotalsCount = results
End Function

' Method totalCountedFiles:                                                   '
'                                                                             '
'   Counts and returns the total number of files sucessfully processed.       '
'                                                                             '
Private Function totalCountedFiles() As Long
    Dim i As Integer, count As Long
    
    For i = 0 To m_numCountList - 1
        count = count + directoryTotalCountedFiles(m_countList(i))
    Next
    
    totalCountedFiles = count
End Function

' Method totalExcludedFiles:                                                  '
'                                                                             '
'   Counts and returns the total number of files that were skipped due to     '
'   error.                                                                    '
'                                                                             '
Private Function totalExcludedFiles() As Long
    Dim i As Integer, count As Long
    
    For i = 0 To m_numCountList - 1
        count = count + directoryTotalExcludedFiles(m_countList(i))
    Next
    
    totalExcludedFiles = count
End Function

' Method getTotalsCount:                                                      '
'                                                                             '
'   Totals the statistics for every file counted in this session, and returns '
'   the results in a FILECOUNT_RESULTS type.  The filename and excluded       '
'   properties of the return UDT will not be set.                             '
'                                                                             '
Private Function getTotalsCount() As FILECOUNT_RESULTS
    Dim totalResults As FILECOUNT_RESULTS, dirResults As FILECOUNT_RESULTS
    Dim i As Integer
    
    With totalResults
        For i = 0 To m_numCountList - 1
            dirResults = getDirectoryTotalsCount(m_countList(i))
            
            .numChars = .numChars + dirResults.numChars
            .numCommentLines = .numCommentLines + dirResults.numCommentLines
            .numExcluded = .numExcluded + dirResults.numExcluded
            .numLines = .numLines + dirResults.numLines
            .numWhitespace = .numWhitespace + dirResults.numWhitespace
            .size = .size + dirResults.size
        Next
    End With
    
    getTotalsCount = totalResults
End Function

' Method showFileCountStats:                                                  '
'                                                                             '
'   Displays the file count statistics for the supplied UDT.                  '
'                                                                             '
Private Sub showFileCountStats _
 (filecount As FILECOUNT_RESULTS, _
  totals As FILECOUNT_RESULTS, _
  Optional showAsPercentOfTotals As Boolean = False _
)
    If filecount.excluded Then
        ' Excluded.
        setOutputFont , , True, , , vbRed
        addOutputText "File could not be counted."
        setOutputFont
        
    ElseIf filecount.numLines = 0 Then
        ' No lines.
        setOutputFont , , True, , , vbRed
        addOutputText "No lines counted."
        setOutputFont
        
    Else
        ' Size...
        addOutputText "Size: " & bytes(filecount.size), False
        If showAsPercentOfTotals Then
            setOutputFont , , , True, , vbBlue
            addOutputText " [" & percent(filecount.size, totals.size) & "]"
            setOutputFont
        Else
            addOutputText
        End If
        
        addOutputText
    
        ' Lines...
        addOutputText "Number of lines: " & number(filecount.numLines), False
        If showAsPercentOfTotals Then
            setOutputFont , , , True, , vbBlue
            addOutputText " [" & percent(filecount.numLines, _
                          totals.numLines) & "]"
            setOutputFont
        Else
            addOutputText
        End If
        
        If filecount.numCommentLines <> 0 Then
            addOutputText "Number of pure comment lines (VB): " & _
                          number(filecount.numCommentLines) & " (" & _
                          percent(filecount.numCommentLines, _
                          filecount.numLines) & ")", False
                          
            If showAsPercentOfTotals Then
                setOutputFont , , , True, , vbBlue
                addOutputText " [" & percent(filecount.numCommentLines, _
                              totals.numCommentLines) & "]"
                setOutputFont
            Else
                addOutputText
            End If
        End If
        
        addOutputText "Number of blank lines: ", False
        If filecount.numExcluded = 0 Then
            addOutputText "None"
        Else
            addOutputText number(filecount.numExcluded) & " (" & _
                          percent(filecount.numExcluded, filecount.numLines) _
                          & ")", False
                          
            If showAsPercentOfTotals Then
                setOutputFont , , , True, , vbBlue
                addOutputText " [" & percent(filecount.numExcluded, _
                              totals.numExcluded) & "]"
                setOutputFont
            Else
                addOutputText
            End If
                          
            addOutputText
            setOutputFont , , , True
            addOutputText "Number of significant lines: " & _
                          number(filecount.numLines - filecount.numExcluded), _
                          False
                          
            If showAsPercentOfTotals Then
                setOutputFont , , , True, , vbBlue
                addOutputText " [" & percent(filecount.numLines - _
                              filecount.numExcluded, totals.numLines - _
                              totals.numExcluded) & "]"
                setOutputFont
            Else
                addOutputText
            End If
                          
            setOutputFont
        End If
        
        addOutputText
        
        ' Characters...
        addOutputText "Number of characters: " & number(filecount.numChars)
        addOutputText "Number of whitespace characters: ", False
        If filecount.numWhitespace = 0 Then
            addOutputText "None"
        Else
            addOutputText number(filecount.numWhitespace) & " (" & _
                          percent(filecount.numWhitespace, _
                          filecount.numChars) & ")"
                          
            addOutputText
            setOutputFont , , , True
            addOutputText "Number of significant characters: " & _
                          number(filecount.numChars - _
                          filecount.numWhitespace)
            setOutputFont
        End If
        
        addOutputText
        
        ' Average line length...
        addOutputText "Average line length: " & _
                      number(filecount.numChars / filecount.numLines)
        If filecount.numWhitespace > 0 Then
            addOutputText "Average line length (whitespace excluded): " & _
             number((filecount.numChars - filecount.numWhitespace) / _
             filecount.numLines)
        End If
    End If
    
    addOutputText
End Sub

' Method showTop10Files:                                                      '
'                                                                             '
'   Reports the top 10 files, by number of significant lines.                 '
'                                                                             '
Private Sub showTop10Files()
    setOutputFont "Arial", 10, True
    addOutputText "Top 10 Files (By Number of Significant Lines)"
    setOutputFont
End Sub

' Method addDirectorySubfoldersDetails:                                       '
'                                                                             '
'   Determines if the supplied directory has any subfolders, and, if so,      '
'   shows details about this folder and all subfolders.                       '
'                                                                             '
Private Sub addDirectorySubfoldersDetails _
 (fldr As DIRCOUNT_RESULTS, _
  totalCountedFiles As Long, _
  totals As FILECOUNT_RESULTS _
)
End Sub

' Method showDirectoryDetails:                                                '
'                                                                             '
'   Generates a report section showing the details about the supplied         '
'   directory.                                                                '
'                                                                             '
Private Sub showDirectoryDetails _
 (fldr As DIRCOUNT_RESULTS, _
  totalCountedFiles As Long, _
  totals As FILECOUNT_RESULTS _
)
    Dim i As Integer

    setOutputFont , , True, False, True
    addOutputText "Folder " & fldr.path
    setOutputFont
    
    addDirectorySubfoldersDetails fldr, totalCountedFiles, totals
    
    ' Statistics about only this folder.
    setOutputFont , , True
    addOutputText "Cumulative details for this folder:"
    setOutputFont
    
    showFileCountStats getDirectoryTotalsCount(fldr), totals, True
    
    If m_settingsForm.chkShowFileDetails.value = 1 Then
        addOutputText
    
        ' Statistics for each individual file.
        For i = 0 To fldr.numFiles - 1
            setOutputFont , , , , True
            addOutputText "File " & extractFilename(fldr.files(i).filename)
            setOutputFont
            
            showFileCountStats fldr.files(i), totals, True
        Next
    End If
    
    ' Done.
    addOutputText
End Sub

' Method showResultsOverview:                                                 '
'                                                                             '
'   Generates the report section showing the general and totaled results.     '
'                                                                             '
Private Sub showResultsOverview _
 (totalCounted As Long, _
  totalExcluded As Long, _
  totals As FILECOUNT_RESULTS _
)
    ' Count...
    totalCounted = totalCountedFiles
    totalExcluded = totalExcludedFiles
    totals = getTotalsCount
    
    ' Add information.
    setOutputFont "Arial", 12, True
    addOutputText "Results Overview"
    setOutputFont
    
    ' Total number of files...
    addOutputText "Total number of folders: " & number(m_numCountList)
    addOutputText "Total number of counted files: " & number(totalCounted)
    addOutputText "Total number of excluded files: ", False
    
    If totalExcluded = 0 Then
        addOutputText "None"
    Else
        setOutputFont , , True, , , vbRed
        
        addOutputText totalExcluded & " (" & _
                      percent(totalExcluded, totalCounted + totalExcluded) _
                      & ")"
                      
        setOutputFont
    End If
    
    addOutputText
    
    ' Totals.
    setOutputFont , , True, , True
    addOutputText "Line statistics for entire listing:"
    
    setOutputFont
    showFileCountStats totals, totals
End Sub

' Method showDirectories:                                                     '
'                                                                             '
'   Generates the reports for each directory.                                 '
'                                                                             '
Private Sub showDirectories(totalCounted As Long, totals As FILECOUNT_RESULTS)
    Dim i As Integer
    
    setOutputFont "Arial", 12, True
    addOutputText "Folder Specifics"
    setOutputFont
    
    For i = 0 To m_numCountList - 1
        lblCurFile.Caption = "Generating report: Folder " & _
                             m_countList(i).path
        pbProgress.value = i + 1
        DoEvents
    
        showDirectoryDetails m_countList(i), totalCounted, totals
    Next
End Sub

' Method generateReport:                                                      '
'                                                                             '
'   Generates the report.                                                     '
'                                                                             '
Private Sub generateReport()
    Dim totalCounted As Long, totalExcluded As Long
    Dim totals As FILECOUNT_RESULTS

    lblCurFile.Caption = "Generating report: Overview"
    pbProgress.value = 0
    pbProgress.Max = m_numCountList
    DoEvents

    showResultsOverview totalCounted, totalExcluded, totals
    
    If m_settingsForm.chkShowFolderDetails.value = 1 Then _
     showDirectories totalCounted, totals
End Sub

' Method countFile:                                                           '
'                                                                             '
'   Sets the current file label to reflect that the supplied file is being    '
'   counted, and accumulates the data for that file.                          '
'                                                                             '
Private Sub countFile(path As String, file As FILECOUNT_RESULTS)
    Dim filenum As Integer, curLine As String

    ' Update the UI...
    pbProgress.value = pbProgress.value + 1
    lblCurFile.Caption = "Current File:" & vbNewLine & path & file.filename
    DoEvents
    
    ' Begin the data collection.
    file.size = FileLen(path & file.filename)
    
    filenum = FreeFile
    Open path & file.filename For Input As filenum
    Do Until EOF(filenum)
    
        Line Input #filenum, curLine
        
        ' Update totals...
        file.numLines = file.numLines + 1
        file.numChars = file.numChars + Len(curLine)
        file.numWhitespace = file.numWhitespace + _
                             mbiFunctions.CharacterCount(curLine, " ") + _
                             mbiFunctions.CharacterCount(curLine, vbTab)
        If Trim(curLine) = "" Then file.numExcluded = file.numExcluded + 1
        If Left(Trim(curLine), 1) = "'" Then file.numCommentLines = _
         file.numCommentLines + 1
        
        ' Check to see if we need to cancel.
        DoEvents
        If m_cancel Then cancel: Exit Sub
    Loop
    Close filenum
    
    ' Show that this file was included.
    file.excluded = False
    Exit Sub
    
countFile_Error:
    ' Show that this file was not counted.
    Close filenum
    file.excluded = True
    
    Exit Sub
End Sub

' Method addFile:                                                             '
'                                                                             '
'   Adds the supplied filename to the list of files to count.  If the file's  '
'   path already exists, a new file is added to that directory.  Otherwise, a '
'   new directory is created and the new file is added to it.                 '
'                                                                             '
Private Sub addFile(filename As String)
    Dim filePath As String, i As Integer
    
    filePath = UCase(extractPathname(filename))
    
    ' See if the path exists.
    For i = 0 To m_numCountList - 1
        If UCase(m_countList(i).path) = filePath Then
            ' This is it.  Add a new file to this directory.
            ReDim Preserve m_countList(i).files(m_countList(i).numFiles)
            
            m_countList(i).files(m_countList(i).numFiles).filename = _
             extractFilename(filename)
             
            m_countList(i).numFiles = m_countList(i).numFiles + 1
            
            Exit Sub
        End If
    Next
    
    ' The directory wasn't found.  We'll have to add it.
    ReDim Preserve m_countList(m_numCountList)
    
    m_countList(m_numCountList).path = extractPathname(filename)
    ReDim m_countList(m_numCountList).files(0)
    m_countList(m_numCountList).numFiles = 1
    
    m_countList(m_numCountList).files(0).filename = extractFilename(filename)
    
    m_numCountList = m_numCountList + 1
End Sub

' Method buildCountList:                                                      '
'                                                                             '
'   Sets up the m_countList array with the list of files to be counted, from  '
'   the settings form.                                                        '
'                                                                             '
Private Sub buildCountList()
    Dim i As Integer
    
    ' Get the list of files to be counted from the settings form.
    For i = 0 To m_settingsForm.lstFiles.ListCount - 1
        addFile m_settingsForm.lstFiles.List(i)
    Next
End Sub

' Method setOutputFont:                                                       '
'                                                                             '
'   Changes the output font characteristics on the results window.            '
'                                                                             '
Private Sub setOutputFont _
 (Optional name As String = "Times New Roman", _
  Optional size As Integer = 10, _
  Optional bold As Boolean = False, _
  Optional italic As Boolean = False, _
  Optional underline As Boolean = False, _
  Optional color As Long = vbBlack _
)
    With m_outputWindow.rtfResults
        .SelFontName = name:        m_lastFont.name = name
        .SelFontSize = size:        m_lastFont.size = size
        .SelBold = bold:            m_lastFont.bold = bold
        .SelItalic = italic:        m_lastFont.italic = italic
        .SelUnderline = underline:  m_lastFont.underline = underline
        .SelColor = color:          m_lastFont.color = color
    End With
End Sub

' Method addOutputText:                                                       '
'                                                                             '
'   Appends the supplied output text to the output window.                    '
'                                                                             '
Private Sub addOutputText _
 (Optional Text As String = "", _
  Optional newLine As Boolean = True _
)
    If m_lastFont.name <> "" Then
        setOutputFont m_lastFont.name, _
                      m_lastFont.size, _
                      m_lastFont.bold, _
                      m_lastFont.italic, _
                      m_lastFont.underline, _
                      m_lastFont.color
    End If
    
    If newLine Then
        m_outputWindow.rtfResults.SelText = Text & vbNewLine
        m_outputWindow.rtfResults.SelStart = _
         m_outputWindow.rtfResults.SelStart + Len(Text) + 2
    Else
        m_outputWindow.rtfResults.SelText = Text
        m_outputWindow.rtfResults.SelStart = _
         m_outputWindow.rtfResults.SelStart + Len(Text)
    End If
End Sub

' Method initialize:                                                          '
'                                                                             '
'   Creates the output window and adds some opening comments regarding this   '
'   particular line-counting session.                                         '
'                                                                             '
Private Sub initialize()
    ' Clear.
    m_cancel = False
    m_cancelComplete = False
    
    ReDim m_countList(0)
    m_numCountList = 0
    
    m_lastFont.name = ""

    ' Show me.
    Me.Show , m_settingsForm
    m_settingsForm.Enabled = False

    ' Create and load the output window.
    Set m_outputWindow = New frmOutput
    Load m_outputWindow
    
    ' Determine the number of files to process.
    pbProgress.Min = 0
    pbProgress.value = 0
    pbProgress.Max = m_settingsForm.lstFiles.ListCount
    
    ' We will assume that the results are to be displayed in the rich text
    ' format.  If not, they will be converted to plain text after the counting
    ' is complete.
    setOutputFont "Arial", 12, True
    addOutputText "mbiSoft LineCounter 1.0"
    
    setOutputFont "Arial", 12
    addOutputText "Line count performed on " & Now
    addOutputText "Root: " & m_settingsForm.txtRoot.Text
    addOutputText
    
    setOutputFont
    
    ' Build the list of files to count.
    buildCountList
End Sub

' Method cancel:                                                              '
'                                                                             '
'   To be called by the countLines method if the counting process is          '
'   canceled.  This will unload the output form as appropriate and upon       '
'   completion set the hourglass back to normal.                              '
'                                                                             '
Private Sub cancel()
    If Not m_cancelComplete Then
        On Error Resume Next
        
        Close
        
        Unload m_outputWindow
        Me.MousePointer = vbNormal
        m_settingsForm.Enabled = True
        
        Unload Me
        
        m_cancelComplete = True
    End If
End Sub

' Event Click on cmdCancel:                                                   '
'                                                                             '
'   Sets the cancel flag to True, then disables the cmdCancel button and      '
'   changes the form mouse pointer to vbHourglass to indicate to the user     '
'   that the cancel will take place immediately.                              '
'                                                                             '
Private Sub cmdCancel_Click()
    m_cancel = True
    
    cmdCancel.Enabled = False
    Me.MousePointer = vbHourglass
End Sub

