Attribute VB_Name = "mbiFunctions"
Option Explicit
'20020618 another test comment in newly created FileScriptDemo project on Mike's machine to try VSS
'20020410 test comment to straighten out VSS confusion

' Windows API
' =============================================================================

' This module contains generic helper functions used my mbiSoft programs

' NOTE:  THERE IS NO ERROR CHECKING IN THESE FUNCTIONS, YOU MUST DO THAT FOR YOURSELF.

' RoundDecimals         Rournds a number to the nth decimal place (1 = 10's, 2 = 100's, etc...).
'                       NOTE:  If you are using VB 6.0 or later, do not use this function.  Instead,
'                       use the Round() function supplied with the language itself.  Also be aware
'                       that despite the function's name, RoundDecimals does not actually ROUND the
'                       number up or down -- it mearly chops off any digits after DigitsAfterDecimal.
'   Number                  The number to round.
'   DigitsAfterDecimal      The number of digits after the decimal place to round to.

' InStrBack:            Works just like InStr, only backwards.
'   String1                 The string to search.
'   String2                 The string to find.
'   [StartAt]               The starting position of the search (default is at end).

'   NOTE                    Returns 0 if the character is not found.

' NthCharacterPos:      Returns the starting position of the Nth occurence of a string within another
'                       string.
'   SearchString            The string to search.
'   Character               The character (or string) to search for.
'   n                       The occurence to look for.
'   [StartAt]               The positon to start searching (default is at beginning).

'   NOTE                    Returns 0 if occurence n does not exist.

' CharacterCount        Returns the number of occurences of a string within another string.
'   Text                    The string to search.
'   Character               The character (or string) to search for.
'   [StartAt]               The position to start searching (default is at beginning).

' SaveFormState         Saves the visual state (Left, Top, Width, Height, WindowState) to the Windows
'                       registry so that it can be restored later.
'   Window                  The form whose state to save.
'   [SaveMaximized]         Specifies whether or not to save the Left, Top, Width, and Height variables
'                           if the form is maximized.  The default is False.

'   NOTE                    The registry path that the function saves to is:

'                           HKEY_CURRENT_USER\Software\VB and VBA Program Settings\App.ExeName\Window.Name

' RestoreFormState      Restores a form state saved using the SaveFormState function.
'   Window                  The form whose state to restore.
'   [AllowMinimized]        Specifies whether or not to restore the WindowState property if it is
'                           minimized.  The default is False.
'   [NameOverload]          If specified, the program will read from the registry the settings for a form
'                           of this name, instead of Window.Name.

' AddArrayElement       Adds an element to an array at a specified location.
'   Arry                    The array to add to.
'   Element                 The element to add to the array.
'   [Index]                 The index at which to add the array element.  The default is at the end.

'   NOTE                    If the index specified is out of bounds, the index is automatically defaulted
'                           to be the last element in the resized array.

' RemoveArrayElement    Removes an element from an a specified position in an array.
'   Arry                    The array from which to remove the element.
'   Index                   The index of the element to remove.

'   NOTE                    This sub performs no task is LBound(Arry) = UBound(Arry)

' Function api_getINISetting:
'
'   Reads and returns a string from the given INI file.
'
Private Declare Function api_getINISetting Lib "kernel32" Alias _
"GetPrivateProfileStringA" _
 (ByVal sectionName As String, _
  ByVal keyName As String, _
  ByVal default As String, _
  ByVal result As String, _
  ByVal size As Long, _
  ByVal filename As String _
 ) As Long

' Function api_setINISetting:
'
'   Writes the given setting to the given INI file.
'
Private Declare Function api_setINISetting Lib "kernel32" Alias _
"WritePrivateProfileStringA" _
 (ByVal sectionName As String, _
  ByVal keyName As String, _
  ByVal setting As String, _
  ByVal filename As String _
 ) As Long

Public Function isByte(value As String) As Boolean
    Dim byteTemp As Byte

    On Error GoTo isByte_error
    byteTemp = CByte(value)
    
    isByte = True
    Exit Function
    
isByte_error:
    isByte = False
End Function

Public Function isInteger(value As String) As Boolean
    Dim intTemp As Integer
    
    On Error GoTo isInteger_error
    intTemp = CInt(value)
    
    isInteger = True
    Exit Function
    
isInteger_error:
    isInteger = False
End Function

Public Function isLong(value As String) As Boolean
    Dim longTemp As Long
    
    On Error GoTo isLong_error
    longTemp = CLng(value)
    
    isLong = True
    Exit Function
    
isLong_error:
    isLong = False
End Function

Public Function isSingle(value As String) As Boolean
    Dim singleTemp As Single
    
    On Error GoTo isSingle_error
    singleTemp = CSng(value)
    
    isSingle = True
    Exit Function
    
isSingle_error:
    isSingle = False
End Function

Public Function isDouble(value As String) As Boolean
    Dim doubleTemp As Double
    
    On Error GoTo isDouble_error
    doubleTemp = CDbl(value)
    
    isDouble = True
    Exit Function
    
isDouble_error:
    isDouble = False
End Function

Public Function getINISetting(iniFile As String, section As String, key As String, ByVal default As String) As String
    Dim f As String, s As String, k As String, d As String, rval As String, buffLen As Long, res As Long
    
    f = iniFile & Chr(0)
    s = section & Chr(0)
    k = key & Chr(0)
    d = default & Chr(0)
    
    Do
        buffLen = buffLen + 250
        rval = String(buffLen, 0)
        
        res = api_getINISetting(s, k, d, rval, buffLen, f)
    Loop Until res <> buffLen - 1
    
    getINISetting = Left(rval, res)
End Function

Public Sub setINISetting(iniFile As String, section As String, key As String, ByVal value As String)
    api_setINISetting section & Chr(0), key & Chr(0), value & Chr(0), iniFile & Chr(0)
End Sub

Public Sub AddArrayElement(arry() As Variant, Element As Variant, Optional index As Integer = -1)
    ReDim Preserve arry(UBound(arry) + 1)
    
    If index < LBound(arry) Or index > UBound(arry) Then index = UBound(arry)
    
    Dim i As Integer
    For i = UBound(arry) To index + 1
        arry(i) = arry(i - 1)
    Next
    
    arry(index) = Element
End Sub

Public Sub RemoveArrayElement(arry() As Variant, index As Integer)
    Dim i  As Integer
    
    If LBound(arry) = UBound(arry) Then Exit Sub
    
    For i = index To UBound(arry) - 1
        arry(i) = arry(i + 1)
    Next
    
    ReDim Preserve arry(UBound(arry) - 1)
End Sub

Public Sub SaveFormState(Window As Form, Optional SaveMaximized As Boolean = False)
    If (Window.WindowState = vbMaximized And SaveMaximized) Or Window.WindowState <> vbMaximized Then
        SaveSetting App.EXEName, Window.name, "Left", Window.Left
        SaveSetting App.EXEName, Window.name, "Top", Window.top
        SaveSetting App.EXEName, Window.name, "Width", Window.Width
        SaveSetting App.EXEName, Window.name, "Height", Window.Height
    End If
    
    SaveSetting App.EXEName, Window.name, "WindowState", Window.WindowState
End Sub

Public Sub RestoreFormState(Window As Form, Optional AllowMinimized As Boolean = False, Optional NameOverload As String = "")
    Dim NameToUse As String
    
    If NameOverload <> "" Then NameToUse = NameOverload Else NameToUse = Window.name
    
    Dim temp As Integer
    temp = GetSetting(App.EXEName, NameToUse, "WindowState", Window.WindowState)
    If (temp = vbMinimized And AllowMinimized) Or temp <> vbMinimized Then Window.WindowState = temp
    
    Window.Left = GetSetting(App.EXEName, NameToUse, "Left", Window.Left)
    Window.top = GetSetting(App.EXEName, NameToUse, "Top", Window.top)
    Window.Width = GetSetting(App.EXEName, NameToUse, "Width", Window.Width)
    Window.Height = GetSetting(App.EXEName, NameToUse, "Height", Window.Height)
End Sub

Public Function InStrBack(String1 As String, String2 As String, Optional StartAt As Integer = -1) As Integer
    Dim i As Integer, CurPart As String
    
    If StartAt < 1 Or StartAt > Len(String1) Then StartAt = Len(String1)
    
    For i = StartAt To 1 Step -1
        CurPart = Mid(String1, i, Len(String2))
        If CurPart = String2 Then
            InStrBack = i
            Exit Function
        End If
    Next
End Function

Public Function NthCharacterPos(SearchString As String, Character As String, N As Integer, Optional StartAt As Integer = 1) As Integer
    Dim i As Integer, n2 As Integer, CurPart As String
    
    For i = StartAt To Len(SearchString) - Len(Character)
        CurPart = Mid(SearchString, i, Len(Character))
        If CurPart = Character Then
            n2 = n2 + 1
            If n2 = N Then NthCharacterPos = i: Exit Function
        End If
    Next
End Function

Public Function CharacterCount(Text As String, Character As String, Optional StartAt As Integer = 1) As Integer
    Dim i As Integer, N As Integer
    For i = StartAt To Len(Text) - Len(Character) + 1
        If Mid(Text, i, Len(Character)) = Character Then N = N + 1
    Next
    CharacterCount = N
End Function

' Takes the command line arguments and parses them into switches based on the following
' rules:
'
'   * Data enclosed between double quotes (") is left as a single parameter with the
'     quotation marks removed.
'
'   * Any character preceded by a forward slash (/) is treated as its own parameter
'     (i.e., /d/f /H would translate to 3 array elements: /d, /f, and /H.
'
'   * Any group of characters folling a forward slash and not separated by whitespace are
'     separated into individual settings (i.e., /dfH would translate to 3 array elements:
'     /d, /h, and /H).
'
'   * Any group of characters not separated by whitespace and not preceded by a forward
'     slash are left as a single array element.  If a forward slash appears elsewhere in
'     the group, it is left as part of the array element.
'
'   * Escape characters are supported: #" inserts a quotation mark, #t inserts a tab
'     character, #n inserts a newline character, #f inserts a line-feed character, and ##
'     inserts a the pound-sign.
'
'   EXAMPLE:
'
'       MyApp.exe "C:\My Documents\Somefile.txt" /ad/H  /i read-only="#"if requested#""
'
'     Becomes:
'
'       C:\My Documents\Somefile.txt
'       /a
'       /d
'       /H
'       /i
'       read-only="if requested"
'
'   NOTE: Only the space character is currently considered whitespace (tabs ARE NOT).
'
'   Arrays will always be zero-based.
'
'   (c) 2001 Mark Biddlecom
'   ====================================================================
'   Student, Rochester Institute of Technology
'   Software Engineer, SoftWright LLC (Aurora, CO)
'
'   markbiddlecom@mbisoft.com
'   www.mbisoft.com
'
Public Function splitCommandLine(ByVal cmdLine As String) As String()
    Dim i As Integer, inQuote As Boolean, inSwitchSet As Boolean
    Dim curParam As String, returnArray() As String, curChar As String
    Dim arrayInitialized As Boolean
    
    ' Don't let them pass an empty string into this function.
    If Trim(cmdLine) = "" Then
        Err.Raise 35353535, , "Cannot pass a blank or all-space string to " & _
         "splitCommandLine."
    End If
    
    ' Replace all escape characters except for the double-quote escape character.  That
    ' will be replaced with some temporary value for now.
    cmdLine = Replace(cmdLine, "##", Chr(1))
    cmdLine = Replace(cmdLine, "#""", Chr(2)) ' Note: to embed a double-quote in a string
                                              ' you place two double-quotes next to each-
                                              ' other.
    
    cmdLine = Replace(cmdLine, "#t", vbTab)
    cmdLine = Replace(cmdLine, "#n", vbCr)
    cmdLine = Replace(cmdLine, "#f", vbLf)
    
    cmdLine = Replace(cmdLine, Chr(1), "#")
    
    ' We'll put the quotes back later.  Get our parameters by parsing one character at a
    ' time.
    For i = 1 To Len(cmdLine)
        ' Get the current character (the ith character in cmdLine).
        curChar = Mid(cmdLine, i, 1)
    
        ' See what type of parameter we're currently dealing with.
        If inQuote Then
            ' Add the current character to the current parameter if it's not a double-
            ' quote.  If it's a double-quote, end the string.
            If curChar <> """" Then
                curParam = curParam & curChar
            Else
                ' End the string.  If we're in a switch set, also add the parameter to
                ' the array.
                inQuote = False
                
                If inSwitchSet Then
                    ' Resize the array.
                    If arrayInitialized Then
                        ReDim Preserve returnArray(UBound(returnArray) + 1)
                    Else
                        ' Start at zero.
                        ReDim returnArray(0)
                        arrayInitialized = True
                    End If
                    
                    ' Add the character.
                    returnArray(UBound(returnArray)) = curParam
                    curParam = ""
                End If
            End If
        ElseIf inSwitchSet Then
            ' OK, we're talking about rules 3 and 4 here.  Add the character as its own
            ' element if it's not the forward-slash.  We'll also deal with quotes here,
            ' and spaces should terminate the switch set.
            If curChar = """" Then
                ' Start a quote string.
                curParam = "/"
                inQuote = True
            ElseIf curChar = " " Then
                ' No more switch set!
                curParam = ""
                inSwitchSet = False
            ElseIf curChar <> "/" Then
                ' Resize the array.
                If arrayInitialized Then
                    ReDim Preserve returnArray(UBound(returnArray) + 1)
                Else
                    ' Start at zero.
                    ReDim returnArray(0)
                    arrayInitialized = True
                End If
                
                ' Add the character.
                returnArray(UBound(returnArray)) = "/" & curChar
                curParam = ""
            End If
        ElseIf curChar = " " Then
            ' We will flat out ignore spaces.  But, we'll add the current parameter if
            ' it's not empty.
            If curParam <> "" Then
                ' Give an extra space at the end of the array.
                If arrayInitialized Then
                    ReDim Preserve returnArray(UBound(returnArray) + 1)
                Else
                    ' Start at zero.
                    ReDim returnArray(0)
                    arrayInitialized = True
                End If
                
                ' Add the value and reset it.
                returnArray(UBound(returnArray)) = curParam
                curParam = ""
            End If
        ElseIf curChar = "/" And curParam = "" Then
            ' This is the start of a switch set.
            inSwitchSet = True
        Else
            ' OK, this is a parameter of the type described by rule 4.  Check to see if
            ' we're starting a string.
            If curChar = """" Then
                ' Yup.  Say so.
                inQuote = True
            Else
                ' Nope.  Just add the current character to the list.
                curParam = curParam & curChar
            End If
        End If
    Next
    
    ' Add the final parameter to the list, if it exists.
    If curParam <> "" Then
        ' Give an extra space at the end of the array.
        If arrayInitialized Then
            ReDim Preserve returnArray(UBound(returnArray) + 1)
        Else
            ' Start at zero.
            ReDim returnArray(0)
            arrayInitialized = True
        End If
        
        ' Add the value and reset it.
        returnArray(UBound(returnArray)) = curParam
        curParam = ""
    End If
    
    ' Put all the quotes back.
    For i = 0 To UBound(returnArray)
        returnArray(i) = Replace(returnArray(i), Chr(2), """")
    Next
    
    ' All done!  :)
    splitCommandLine = returnArray
End Function

Public Function addPathBackslash(path As String) As String
    If Right(path, 1) <> "\" Then
        addPathBackslash = path & "\"
    Else
        addPathBackslash = path
    End If
End Function

Public Function buildPath(path As String, filename As String) As String
    If Right(path, 1) <> "\" Then
        buildPath = path & "\" & filename
    Else
        buildPath = path & filename
    End If
End Function

Public Function extractPathname(filename As String) As String
    If InStr(1, filename, "\") Then
        extractPathname = Left(filename, InStrRev(filename, "\"))
    Else
        extractPathname = addPathBackslash(CurDir)
    End If
End Function

Public Function extractExtension(filename As String) As String
    If InStr(1, filename, ".") Then
        If InStrRev(filename, ".") <> Len(filename) Then
            extractExtension = Mid(filename, InStrRev(filename, ".") + 1)
        Else
            extractExtension = ""
        End If
    Else
        extractExtension = ""
    End If
End Function

Public Function extractFilename(filename As String) As String
    If InStr(1, filename, "\") Then
        If InStrRev(filename, "\") <> Len(filename) Then
            extractFilename = Mid(filename, InStrRev(filename, "\") + 1)
        Else
            extractFilename = ""
        End If
    Else
        extractFilename = filename
    End If
End Function

Public Function extractFileTitle(filename As String) As String
    Dim working As String
    
    working = filename
    
    ' Strip off the path.
    If InStr(1, working, "\") Then
        If InStrRev(working, "\") <> Len(filename) Then
            working = Mid(working, InStrRev(working, "\") + 1)
        Else
            Exit Function
        End If
    End If
    
    ' Strip off the extension.
    If InStr(1, working, ".") Then
        If InStrRev(working, ".") <> Len(working) And _
         InStrRev(working, ".") <> 1 Then
            working = Left(working, InStrRev(working, ".") - 1)
        Else
            Exit Function
        End If
    End If
    
    ' We've got it.
    extractFileTitle = working
End Function

' Method getSWDate:
'
'   Returns the current day, month, and year in the form YYYYMMDD.
'
Public Function getSWDate() As String
    Dim yearString As String, monthString As String, dayString As String
    
    yearString = Year(Now)
    monthString = Month(Now)
    dayString = Day(Now)
    
    monthString = String(2 - Len(monthString), "0") & monthString
    dayString = String(2 - Len(dayString), "0") & dayString
    
    getSWDate = yearString & monthString & dayString
End Function
