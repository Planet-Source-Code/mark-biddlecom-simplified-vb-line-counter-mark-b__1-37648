Attribute VB_Name = "modBrowseForFolder"
Option Explicit

' This module defines the API functionality that displays the "browse for     '
' folder" dialog box.                                                         '
'                                                                             '
' ** This module contains modified code from "The Access Web" article         '
'    entitled "API: BrowseFolder dialog" by Dev Ashish.                       '
'                                                                             '
'    (http://www.mvps.org/access/api/api0002.htm)                             '
'                                                                             '

' API                                                                         '
' =============================================================================

Private Declare Function SHGetPathFromIDList Lib "shell32.dll" _
Alias "SHGetPathFromIDListA" _
 (ByVal pidl As Long, _
  ByVal pszPath As String _
) As Long

Private Declare Function SHBrowseForFolder Lib "shell32.dll" _
Alias "SHBrowseForFolderA" _
 (lpBrowseInfo As BROWSEINFO _
) As Long

Private Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Private Const BIF_RETURNONLYFSDIRS = &H1
Private Const BIF_DONTGOBELOWDOMAIN = &H2
Private Const BIF_EDITBOX = &H10

Private Const BIF_BROWSEFORCOMPUTER = &H1000
Private Const BIF_BROWSEFORPRINTER = &H2000
Private Const BIF_BROWSEINCLUDEFILES = &H4000

' Enumerations                                                                '
' =============================================================================

' Enumeration MBFF_OPTIONS:                                                   '
'                                                                             '
'   Defines the available options for the browseForFolder method.  You can    '
'   combine more than one option by using the Or or + operators.              '
'                                                                             '
' Members:                                                                    '
'                                                                             '
'   MBFF_O_FILESYSTEMDIRSONLY:      Instructs the dialog to disable the       '
'                                   selection of folders that are not part of '
'                                   the physical file system.                 '
'                                                                             '
'   MBFF_O_DOMAINNETWORKONLY:       Instructs the dialog to disable the       '
'                                   selection of computers residing below the '
'                                   domain level of the network.              '
'                                                                             '
'   MBFF_O_EDITBOX:                 Displays an edit control in the dialog    '
'                                   box that allows the user to type in the   '
'                                   name of a folder explicitly.              '
'                                                                             '
'   MBFF_O_BROWSEFORCOMPUTER:       Disables the selection of any item other  '
'                                   than a network computer.                  '
'                                                                             '
'   MBFF_O_BROWSEFORPRINTER:        Disables the selection of any item other  '
'                                   than a local or network printer.          '
'                                                                             '
'   MBFF_O_BROWSEINCLUDEFILES:      Shows and allows the selection of files   '
'                                   as well as folders.                       '
'                                                                             '
Public Enum MBFF_OPTIONS
    MBFF_O_FILESYSTEMDIRSONLY = BIF_RETURNONLYFSDIRS
    MBFF_O_DOMAINNETWORKONLY = BIF_DONTGOBELOWDOMAIN
    MBFF_O_EDITBOX = BIF_EDITBOX
    MBFF_O_BROWSEFORCOMPUTER = BIF_BROWSEFORCOMPUTER
    MBFF_O_BROWSEFORPRINTER = BIF_BROWSEFORPRINTER
    MBFF_O_BROWSEINCLUDEFILES = BIF_BROWSEINCLUDEFILES
End Enum

' Methods                                                                     '
' =============================================================================

' Method browseForFolder:                                                     '
'                                                                             '
'   Displays the "Browse for Folder" dialog box and returns the user's folder '
'   selection.                                                                '
'                                                                             '
' Parameters:                                                                 '
'                                                                             '
'   title:                      A string to display above the tree view in    '
'                               dialog.  This can be used to convey           '
'                               instructions to the user.                     '
'                                                                             '
'   options:                    The options affecting the dialog display.     '
'                                                                             '
'   [owner]:                    The window handle of the dialog's owner form. '
'                                                                             '
' Returns String:                                                             '
'                                                                             '
'   The path selected by the user, or "" if the dialog was canceled.          '
'                                                                             '
Public Function browseForFolder _
 (title As String, _
  options As MBFF_OPTIONS, _
  Optional owner As Long = 0 _
) As String
    Dim x As Long, bi As BROWSEINFO, dwIList As Long
    Dim szPath As String, wPos As Integer
    
    With bi
        .hOwner = owner
        .lpszTitle = title
        .ulFlags = options
    End With
    
    dwIList = SHBrowseForFolder(bi)
    szPath = Space(512)
    x = SHGetPathFromIDList(ByVal dwIList, ByVal szPath)
    
    If x Then
        wPos = InStr(szPath, Chr(0))
        browseForFolder = Left(szPath, wPos - 1)
    Else
        browseForFolder = ""
    End If
End Function

