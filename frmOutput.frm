VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmOutput 
   Caption         =   "LineCounter Results"
   ClientHeight    =   3645
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4725
   LinkTopic       =   "Form1"
   ScaleHeight     =   3645
   ScaleWidth      =   4725
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog C 
      Left            =   1320
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdConvertToText 
      Caption         =   "Convert to Plain Text"
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton cmdPrint 
      Height          =   375
      Left            =   840
      Picture         =   "frmOutput.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdSave 
      Height          =   375
      Left            =   480
      Picture         =   "frmOutput.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdCopy 
      Height          =   375
      Left            =   120
      Picture         =   "frmOutput.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin RichTextLib.RichTextBox rtfResults 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   5106
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmOutput.frx":0306
   End
End
Attribute VB_Name = "frmOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Private Data Members                                                        '
' =============================================================================

' Determines if the text should be copied as plain text.
Private m_plainText As Boolean

' The filename to save to.
Private m_saveFilename As String

' Properties                                                                  '
' =============================================================================

' Property plainText:                                                         '
'                                                                             '
'   Determines if the text to copy will be copied as Rich Text or plain text. '
'                                                                             '
' Access: Read/Write                                                          '
'                                                                             '
Public Property Get plainText() As Boolean
    plainText = m_plainText
End Property
Public Property Let plainText(newVal As Boolean)
    Dim txt As String

    m_plainText = newVal
    
    If m_plainText Then
        cmdConvertToText.Enabled = False
    
        txt = rtfResults.Text
        
        rtfResults.Text = ""
        rtfResults.Font.name = "Courier New"
        rtfResults.Font.size = 10
        rtfResults.Font.bold = False
        rtfResults.Font.italic = False
        rtfResults.Font.underline = False
        rtfResults.SelColor = vbBlack
        
        rtfResults.Text = txt
    End If
End Property

' Property saveFilename:                                                      '
'                                                                             '
'   The filename that the results will be saved to.                           '
'                                                                             '
' Access: Read/Write                                                          '
'                                                                             '
Public Property Get saveFilename() As String
    saveFilename = m_saveFilename
End Property
Public Property Let saveFilename(newFile As String)
    m_saveFilename = newFile
End Property

' Methods                                                                     '
' =============================================================================

' Method copy:                                                                '
'                                                                             '
'   Copies the results data to the clipboard, in the format specified by the  '
'   plainText property.                                                       '
'                                                                             '
Public Sub copy()
    If plainText Then
        Clipboard.SetText rtfResults.Text, vbCFText
    Else
        Clipboard.SetText rtfResults.TextRTF, vbCFRTF
    End If
End Sub

' Method save:                                                                '
'                                                                             '
'   Saves the data in the clipboard to the file specified by the saveFilename '
'   property, in the format specified by the plainText property.  If          '
'   saveFilename is "", then the browse dialog will be displayed.             '
'                                                                             '
Public Sub save()
    If m_saveFilename = "" Then
        With C
            .DialogTitle = "Save LineCounter Results"
            .CancelError = True
            
            If m_plainText Then
                .Filter = "Text Files (*.txt)|*.txt"
            Else
                .Filter = "Rich Text Files (*.rtf)|*.rtf"
            End If
            
            .ShowSave
            .Flags = cdlOFNHideReadOnly + cdlOFNNoReadOnlyReturn + _
                     cdlOFNPathMustExist + cdlOFNOverwritePrompt
            
            m_saveFilename = .filename
        End With
    End If
    
    If m_plainText Then
        rtfResults.SaveFile m_saveFilename, rtfText
    Else
        rtfResults.SaveFile m_saveFilename, rtfRTF
    End If
    
save_canceled:
    Exit Sub
End Sub

' Method printResults:                                                        '
'                                                                             '
'   Prints the selected text, or if no text is selected, all text.            '
'                                                                             '
Public Sub printResults()
    rtfResults.SelPrint Printer.hDC, True
End Sub

' Private and Helper Methods                                                  '
' =============================================================================

' Event Click on cmdConvertToText:                                            '
'                                                                             '
'   Display the results in text form.                                         '
'                                                                             '
Private Sub cmdConvertToText_Click()
    plainText = True
End Sub

' Event Click on cmdCopy:                                                     '
'                                                                             '
'   Copies.                                                                   '
'                                                                             '
Private Sub cmdCopy_Click()
    copy
End Sub

' Event Click on cmdPrint:                                                    '
'                                                                             '
'   Prints.                                                                   '
'                                                                             '
Private Sub cmdPrint_Click()
    printResults
End Sub

' Event Click on cmdSave:                                                     '
'                                                                             '
'   Saves.                                                                    '
'                                                                             '
Private Sub cmdSave_Click()
    save
End Sub

' Event Load on Form:                                                         '
'                                                                             '
'   Remember the last position of this form.                                  '
'                                                                             '
Private Sub Form_Load()
    mbiFunctions.RestoreFormState Me
End Sub

' Event Resize on Form:                                                       '
'                                                                             '
'   Adjusts the positions and sizes of the various form elements to fit the   '
'   window area of the document form.                                         '
'                                                                             '
Private Sub Form_Resize()
    On Error Resume Next
    
    cmdCopy.Move 100, 100
    cmdSave.Move cmdCopy.Left + cmdCopy.Width, cmdCopy.Top
    cmdPrint.Move cmdSave.Left + cmdSave.Width, cmdCopy.Top
    
    cmdConvertToText.Move Me.ScaleWidth - (cmdConvertToText.Width + 100), _
                          cmdCopy.Top
                          
    rtfResults.Move cmdCopy.Left, _
                    cmdCopy.Top + cmdCopy.Height + 100, _
                    Me.ScaleWidth - 200, _
                    Me.ScaleHeight - (cmdCopy.Top + cmdCopy.Height + 300)
End Sub

' Event Unload on Form:                                                       '
'                                                                             '
'   Remember the last position of this form.                                  '
'                                                                             '
Private Sub Form_Unload(cancel As Integer)
    mbiFunctions.SaveFormState Me
End Sub
