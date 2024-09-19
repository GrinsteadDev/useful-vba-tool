Attribute VB_Name = "IniReader"
''
' IniReader
'   Deserializes an Ini File/String into a VBA Dictionary Object.
'   Serializes a VBA Dictionary Object into an Ini File/String.
' Methods
'  ConvertFromFile
'       Takes a path to a valid Ini file and converts its contents into a
'       Dictionary object
'     Params
'       FilePath As String - This is any path the operating system is able to
'                            resolve into an absolute path. File names without
'                            a directory will assume the current working directory.
'     Returns
'       Dictionary         - An object represenation of the Ini Document.
'
'  ConvertFromString
'       Takes a String value, that matches Ini syntax, and converts it into a
'       Dictionary object
'     Params
'       StringData As String - This is a String value in Ini syntax.
'     Returns
'       Dictionary           - An object represenation of the Ini String.
'
'  ConvertToFile
'       Converts a dictionary object to an Ini File. Valid Dictionary Object Values are Dictionary, Collection,
'       String, Number, Date, Null, and Empty. File names that do not end in .ini will have it appended to them.
'     Params
'       IniDictionary     As Object  - A Dictionary Object represenation of an Ini File. Dictionary returned by this module's
'                                      ConvertFromString or ConvertFromFile functions will always be valid, however this function
'                                      attemp to serialize any provided Dictionary Object.
'       FilePath          As String  - This is any path the operating system is able to resolve into an absolute path. File
'                                      names without a directory will assume the current working directory.
'       OverWriteExisting As Boolean - Default = False. Determines if this function can overwrite an exisiting file.
'     Returns
'       String                       - The absolute path of the new ini file created.
'
'  ConvertToString
'       Converts a dictionary object to an Ini String. Valid Dictionary Object Values are Dictionary, Collection,
'       String, Number, Date, Null, and Empty. File names that do not end in .ini will have it appended to them.
'     Params
'       IniDictionary As Object - A Dictionary Object represenation of an Ini File. Dictionary returned by this module's
'                                 ConvertFromString or ConvertFromFile functions will always be valid, however this function
'                                 attemp to serialize any provided Dictionary Object.
'     Returns
'       String                  - A string that is identical to the contents of a valid ini file.
' Properties
'   LastError -
'       ErrNumber - The error number of the last error Zero means not Error.
'       ErrDesc   - The description of the last error.
'       ErrFunc   - The last module function ran.
' Examples
'   Sub Test()
'       Dim IniDoc As Object, tmp
'
'       Set IniDoc = IniReader.ConvertFromFile("IniReader_TestFile.ini")
'
'       Debug.Print IniDoc("Settings")("Strings")("Default")
'       Debug.Print IniDoc("Settings")("Strings")("StringLiteral1")
'       Debug.Print IniDoc("Settings")("Strings")("StringLiteral2")
'       Debug.Print IniDoc("Settings")("Booleans")("TrueValue")
'       Debug.Print IniDoc("Settings")("Strings")("NullOrEmpty").Count
'       Debug.Print IniDoc("Escapes")("UnicodeString")
'       Debug.Print IniDoc("Comment1")
'
'       Debug.Print vbCrLf & vbCrLf
'
'       Debug.Print IniReader.ConvertToString(IniDoc)
'
'       '' If Errors are encounted they will be saved to LastError
'       Debug.Print IniReader.LastError.ErrFunc
'
'   End Sub
' Notes
'   Comments Can be manually added to a Ini Dictionary by following this pattern.
'       iniDoc.Add "Comment" & iniDoc.Count, "; I am a comment"
'       iniDoc("SectionName").Add "Comment" & iniDoc("SectionName").Count, "; I am a section comment"
' IniReader Supports the follow Ini syntax elements
'   Comments
'     Comments are specificed with the simicolon(;) or number sign(#) character.
'     A comment starts at the comment character and continues to the end of the line.
'     This implementation does NOT support block comments.
'     Example(s):
'       # Comment line starts with number sign
'       ; Comment line starts with simicolon
'       Name=Value ; Inline Comment
'       Name2=Value2 # Inline Comment
'   Document Section
'     Keys that come before the first section def. This creates a logical
'     section it is unamed and is the top level dictionary object.
'     Example(s):
'       1:
'         GName=GValue
'         [FirstSection]
'         Name=Value
'   Sections
'     Example(s):
'       [SectionName]
'   Section Nesting
'     Two nesting styles are supported Literal and Relative. Nesting is done by
'     dot notation. All top-level Sections are nested under the logical Global.
'     Example(s):
'       1: Literal
'         [Section1.Section2]
'       2: Relative
'         [Section1]
'         [.Section2]
'   Multi-Line
'     Line continuation can accomplished by the use of the backslash(\) character.
'     before the newline(\r\n) or (\n).
'     Example(s):
'       Name1=This line is continued onto \
'       the next line.
'   Multi-Value
'     When a key is defined defined more than once in a section it will group the values into
'     an array.
'     Example(s):
'       Name1=Value1
'       Name1=Value2
'       Name1=Value3
'   Multi-Section
'     When a section is defined more than once it will be group together into a single logical section.
'     Example(s):
'       [Section2]
'       Option=1
'       Name=Value
'       [Section3]
'       [Section2]
'       MoreOptions=3
'   Quoted Values
'     Supports both single quotes (') and double quotes (")
'   Escape Characters
'     The escape character is backslash (\). Common escape sequences are below
'       \\ - A literal blackslash, escapes the backslash character.
'       \' - A literal apostrophe, escapes the apostrophe or single quote character
'       \" - A literal quote, escapes the double quote chatacter.
'       \0 - Null character
'       \t - Tab character
'       \r - Carriage return
'       \n - Line feed
'       \; - A literal semicolon, escapes the semicolon character.
'       \# - A literal number sign, escapes the number sign character.
'       \= - A literal equals sign, escapes the equals sign character.
'       \xhhhh - Unicode character with code point 0xhhhh, encoded in UTF-8
'   Read
'     The object can open an INI Formated text file.
'   Write
'     This object can write data to a new file or an exisiting opened file.
'   Type Literals
'     Boolean (case-insensitive)
'       True
'       False
'     Number (all numbers are coalesced into a double regardless of precision)
'       #.## - Double Literal
'       ##   - Long Literal
'       0x## - Hex Literal
'     String - Any data between single or double quotes or data that does not match
'              any other type are strings.
'       "002"  - Double Quote string literal
'       'dsad' - Single Quote string literal
'       also a string - Any value not matching a reconized literal format
'     NullOrEmpty (case-insensitive)
'       Null          - Null Literal
'       Empty         - Empty Literal
'       [white space] - No value give after the equals (=) sign.
Option Explicit

'' Character Encoding Converstion Functions

' VBA does not have native support for UTF-8 encoded; the current standard.
Private Const CP_UTF8 = 65001
''
' WideCharToMultiByte
'   Maps a UTF-16 (wide character) string to a new character string. The new character
'   string is not necessarily from a multibyte character set.
' Params
'    [in]            UINT                               CodePage,
'    [in]            DWORD                              dwFlags,
'    [in]            _In_NLS_string_(cchWideChar)LPCWCH lpWideCharStr,
'    [in]            int                                cchWideChar,
'    [out, optional] LPSTR                              lpMultiByteStr,
'    [in]            int                                cbMultiByte,
'    [in, optional]  LPCCH                              lpDefaultChar,
'    [out, optional] LPBOOL                             lpUsedDefaultChar
'
''
' MultiByteToWideChar
'   Maps a character string to a UTF-16 (wide character) string. The character string
'   is not necessarily from a multibyte character set.
' Params
'    [in]            UINT                              CodePage,
'    [in]            DWORD                             dwFlags,
'    [in]            _In_NLS_string_(cbMultiByte)LPCCH lpMultiByteStr,
'    [in]            int                               cbMultiByte,
'    [out, optional] LPWSTR                            lpWideCharStr,
'    [in]            int                               cchWideChar
#If Win64 Then
    Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" ( _
        ByVal CodePage As Long, _
        ByVal dwFlags As Long, _
        ByVal lpWideCharStr As LongPtr, _
        ByVal ccWideChar As Long, _
        ByVal lpMultiByteStr As LongPtr, _
        ByVal cbMultiByte As Long, _
        ByVal lpDefaultChar As LongPtr, _
        ByVal lpUsedDefaultChar As LongPtr _
    ) As Long
    Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" ( _
        ByVal CodePage As Long, _
        ByVal dwFlags As Long, _
        ByVal lpMultiByteStr As LongPtr, _
        ByVal cbMultiByte As Long, _
        ByVal lpWideCharStr As LongPtr, _
        ByVal cchWideChar As Long _
    ) As Long
#Else
    Private Declare Function WideCharToMultiByte Lib "kernel32" ( _
        ByVal CodePage As Long, _
        ByVal dwFlags As Long, _
        ByVal lpWideCharStr As Long, _
        ByVal ccWideChar As Long, _
        ByVal lpMultiByteStr As Long, _
        ByVal cbMultiByte As Long, _
        ByVal lpDefaultChar As Long, _
        ByVal lpUsedDefaultChar As Long _
    ) As Long
    Private Declare Function MultiByteToWideChar Lib "kernel32" ( _
        ByVal CodePage As Long, _
        ByVal dwFlags As Long, _
        ByVal lpMultiByteStr As Long, _
        ByVal cbMultiByte As Long, _
        ByVal lpWideCharStr As Long, _
        ByVal cchWideChar As Long _
    ) As Long
#End If

'' Regex Expressions
' Matches any comment when the comment character is not between quotes.
' Due to VBA's lack of look behind support check group 0 for slash (\)
' is none-empty to validate match. Group 1 is match.
'   Group 0 - Empty if Valid Match
'   Group 1 - Match
Private Const COMMENT_PATTERN = "(\\?)([;#](?=([^'""\\]*(\\.|""([^'""\\]*\\.)*[^'""\\]*""))*[^'""]*$).*?$)"
' Matches Ini Sections.
'   Group 0 - Match
Private Const SECTION_PATTERN = "^\s*?\[([^'""\r\n]*)\](?=.*[\r\n]*)"
' Matches values as Key Value Pairs
'   Group 0 - Key
'   Group 1 - Equals (can discard)
'   Group 2 - Value
Private Const KEY_VAL_PATTERN = "^([^=\[\];#\r\n]*?)(\s*?=[ \t\f]*)(.*)$"
'' Literal Patterns
' Matches Number formats
'   #.##
'   ##
'   0x##
Private Const NUMBER_PATTERN = "^(\d*\.?\d{2,}|0x[0-9a-zA-Z]*)(?=\s*?$)"
' Matches Boolean Format
Private Const BOOL_PATTERN = "^([Tt][Rr][Uu][Ee]|[Ff][Aa][Ll][Ss][Ee])(?=\s*?$)"
' Matches Null Format
Private Const NULL_PATTERN = "^([Nn][Uu][Ll][Ll]|[Ee][Mm][Pp][Tt][Yy]|\\0)(?=\s*?$)"
' No String literal is needed because all other values are string.
' Matches the Escape character pattern
Private Const ESCAPE_PATTERN = "\\[\\'""0trn;=#]|\\x[\da-zA-Z]{4}"
' Matches non Escaped characters
Private Const NON_ESCAPED_PATTERN = "([\\'""\t;=#\r\n]|[^\x00-\x7F])"

Public Enum IniReaderErrors
    NoError = 0
    GeneralError = vbObjectError + 1
    FileNotFound
    FileAlreadyExists
    FileReadOnly
End Enum

Private Type IniReaderError_Type
    ErrNumber As Long
    ErrDesc As String
    ErrFunc As String
End Type

Public LastError As IniReaderError_Type

Private reg As Object

'' Public Methods
''
' ConvertFromFile
'   Takes a path to a valid Ini file and converts its contents into a
'   Dictionary object
' Params
'   FilePath As String - This is any path the operating system is able to
'                        resolve into an absolute path. File names without
'                        a directory will assume the current working directory.
' Returns
'   Dictionary         - An object represenation of the Ini Document.
Public Function ConvertFromFile(FilePath As String) As Object
    Dim v_out As Object
    
    Set v_out = CreateObject("Scripting.Dictionary")
    Set reg = CreateObject("VBScript.RegExp")
    
    ReadIniFile FilePath, v_out
    
    Set reg = Nothing
    Set ConvertFromFile = v_out
End Function
''
' ConvertFromString
'   Takes a String value, that matches Ini syntax, and converts it into a
'   Dictionary object
' Params
'   StringData As String - This is a String value in Ini syntax.
' Returns
'   Dictionary           - An object represenation of the Ini String.
Public Function ConvertFromString(StringData As String) As Object
    Dim v_out As Object, s_arr() As String, tmp As String
    
    Set v_out = CreateObject("Scripting.Dictionary")
    Set reg = CreateObject("VBScript.RegExp")
    
    s_arr = VBA.Split(VBA.Replace(StringData, vbCr, ""), vbLf)
    
    ProcessLine "", Nothing, True
    For Each tmp In s_arr
        ProcessLine tmp, v_out
    Next
    
    Set reg = Nothing
    Set ConvertFromString = v_out
End Function
''
' ConvertToFile
'   Converts a dictionary object to an Ini File. Valid Dictionary Object Values are Dictionary, Collection,
'   String, Number, Date, Null, and Empty. File names that do not end in .ini will have it appended to them.
' Params
'   IniDictionary     As Object  - A Dictionary Object represenation of an Ini File. Dictionary returned by this module's
'                                  ConvertFromString or ConvertFromFile functions will always be valid, however this function
'                                  attemp to serialize any provided Dictionary Object.
'   FilePath          As String  - This is any path the operating system is able to resolve into an absolute path. File
'                                  names without a directory will assume the current working directory.
'   OverWriteExisting As Boolean - Default = False. Determines if this function can overwrite an exisiting file.
' Returns
'   String                       - The absolute path of the new ini file created.
Public Function ConvertToFile(IniDictionary As Object, FilePath As String, Optional OverWriteExisting As Boolean = False) As String
    '' Appends ".ini" to the file name if not already present.
    If Not VBA.Right(FilePath, 4) Like ".[Ii][Nn][In]" Then
        FilePath = FilePath & ".ini"
    End If
    
    WriteIniFile FilePath, IniDictionary, OverWriteExisting
    
    ConvertToFile = FilePath
End Function
''
' ConvertToString
'   Converts a dictionary object to an Ini String. Valid Dictionary Object Values are Dictionary, Collection,
'   String, Number, Date, Null, and Empty. File names that do not end in .ini will have it appended to them.
' Params
'   IniDictionary As Object - A Dictionary Object represenation of an Ini File. Dictionary returned by this module's
'                             ConvertFromString or ConvertFromFile functions will always be valid, however this function
'                             attemp to serialize any provided Dictionary Object.
' Returns
'   String                  - A string that is identical to the contents of a valid ini file.
Public Function ConvertToString(IniDictionary As Object) As String
    Dim v_out As String, tmp As Variant
    
    For Each tmp In IniDictionary.Keys
         v_out = v_out & ProcessItem(CStr(tmp), IniDictionary(tmp))
    Next
    
    ConvertToString = v_out
End Function

'' Helper Functions

''
' ReadIniFile
'   Opens the specified Ini File, copies its data, then processes it line by line,
'   creating appending to a dictionary object.
' Params
'   FilePath As String - A valid System File Path leading to an Ini File to process
'   IniDoc   As Object - A Dictionary Object to fill with the Ini File's data.
' Returns
'   void
Private Sub ReadIniFile(ByVal FilePath As String, IniDoc As Object)
    LastError.ErrNumber = 0
    LastError.ErrDesc = ""
    LastError.ErrFunc = "IniReader.ReadIniFile"
    On Error GoTo ErrHnd
    
    Dim fso As Object, fp As String, f_data() As Byte, _
        fn As Long, fl As Long, lines() As String, ln
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    fp = VBA.Trim$(FilePath)
    fp = fso.GetAbsolutePathName(fp)
    
    If Not fso.FileExists(fp) Then
        LastError.ErrNumber = IniReaderErrors.FileNotFound
        LastError.ErrDesc = "Ini File Not Found"
        GoTo CleanUp
    End If
    
    fl = FileLen(fp)
    fn = FreeFile
    
    ReDim f_data(0 To fl - 1)
    
    Open fp For Binary As #fn
    
    Get #fn, , f_data
    
    Close #fn
    
    lines = VBA.Split(StringFromUTF8(f_data), vbCrLf)
    
    ProcessLine "", Nothing, True
    For Each ln In lines
        ProcessLine CStr(ln), IniDoc
    Next
    
CleanUp:
    If fn <> 0 Then Close #fn
    Set fso = Nothing
    On Error GoTo 0
    Exit Sub
ErrHnd:
    LastError.ErrNumber = Err.Number
    LastError.ErrDesc = Err.Description
    Err.Clear
    
    Resume CleanUp
End Sub

''
' WriteIniFile
'   Writes to the specified Ini File, by serilizing the provided IniDoc Dictionary Object.
' Params
'   FilePath          As String  - A valid System File Path leading to an Ini File to process
'   IniDoc            As Object  - A Dictionary Object to fill with the Ini File's data.
'   OverWriteExisting As Boolean - Determines if this function can overwrite an exisiting file.
' Returns
'   void
Private Sub WriteIniFile(FilePath As String, IniDoc As Object, OverWriteExisting As Boolean)
    LastError.ErrNumber = 0
    LastError.ErrDesc = ""
    LastError.ErrFunc = "IniReader.WriteIniFile"
    On Error GoTo ErrHnd
    
    Dim fso As Object, fp As String, fn As Long, fl As Long, _
        data() As Byte, ini_str As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
        fp = fso.GetAbsolutePathName(VBA.Trim(FilePath))
    
    If fso.FileExists(fp) And OverWriteExisting Then
        If (GetAttr(fp) And vbReadOnly) = vbReadOnly Then
            LastError.ErrNumber = IniReaderErrors.FileReadOnly
            LastError.ErrDesc = "File is Read-Only"
            GoTo CleanUp
        End If
    
        fso.DeleteFile fp
    ElseIf fso.FileExists(fp) And Not OverWriteExisitng Then
        LastError.ErrNumber = IniReaderErrors.FileAlreadyExists
        LastError.ErrDesc = "File Already Exists"
        GoTo CleanUp
    End If
    
    ini_str = ConvertToString(IniDoc)
    data = UTF8FromString(ini_str)
    fn = FreeFile
    fl = FileLen(fp)
    
    Open fp For Binary As #fn
    
    Put #fn, , data
    
    Close #fn
    
    FilePath = fp
    
CleanUp:
    If fn <> 0 Then Close #fn
    Set fso = Nothing
    On Error GoTo 0
    Exit Sub
ErrHnd:
    LastError.ErrNumber = Err.Number
    LastError.ErrDesc = Err.Description
    Err.Clear
    
    Resume CleanUp
End Sub

''
' ProcessLine
'   Takes a string and turns into a Ini Dictionary Entry.
' Params
'   data          As String  - A single string representing ini data.
'   IniDictionary As Object  - A Dictionary Object to fill with the Ini File's data.
'   ClearCache    As Boolean - Default = False. If True this function clears it last process item.
' Returns
'   void
Private Sub ProcessLine(data As String, IniDictionary As Object, Optional ClearCache As Boolean = False)
    Static curr_item As Object
    
    If ClearCache Then
        Set curr_item = Nothing
        Exit Sub
    End If
    
    LastError.ErrNumber = 0
    LastError.ErrDesc = ""
    LastError.ErrFunc = "IniReader.ProcessLine"
    On Error GoTo ErrHnd
    
    Dim ln As String, m_col As Object, match As Object, tmp As Collection
    
    If reg Is Nothing Then Set reg = CreateObject("VBScript.RegExp")
    If curr_item Is Nothing Then Set curr_item = IniDictionary
    
    With reg
        .MultiLine = True
        .IgnoreCase = False
        .Global = True
    End With
    
    If VBA.Trim(data) = "" Then Exit Sub
    
    '' Sections
    reg.Pattern = SECTION_PATTERN
    If reg.Test(data) Then
        Set m_col = reg.Execute(data)
        
        For Each match In m_col
            Set curr_item = GetCurrSection(IniDictionary, curr_item, match.SubMatches(0))
        Next
    End If
    
    '' Comments
    reg.Pattern = COMMENT_PATTERN
    If reg.Test(data) Then
        
        Set m_col = reg.Execute(data)
        
        For Each match In m_col
            If Not match.SubMatches(0) = "\" Then
                curr_item.Add "Comment" & CStr(curr_item.Count), match.SubMatches(1)
                data = VBA.Replace(data, match.SubMatches(1), "", 1, 1)
            End If
        Next
    End If
    
    '' Key Values
    reg.Pattern = KEY_VAL_PATTERN
    If reg.Test(data) Then
        Set m_col = reg.Execute(data)
        
        For Each match In m_col
            If Not curr_item.Exists(match.SubMatches(0)) Then
                curr_item.Add match.SubMatches(0), ConvertToValueItem(match.SubMatches(2))
            Else
                If TypeOf curr_item(match.SubMatches(0)) Is Collection Then
                    Set tmp = curr_item(match.SubMatches(0))
                Else
                    Set tmp = New Collection
                    
                    tmp.Add curr_item(match.SubMatches(0))
                    
                    curr_item.Remove match.SubMatches(0)
                    curr_item.Add match.SubMatches(0), tmp
                End If
                
                tmp.Add ConvertToValueItem(match.SubMatches(2))
            End If
        Next
    End If
    
CleanUp:
    Set m_col = Nothing
    Set match = Nothing
    Set tmp = Nothing
    On Error GoTo 0
    Exit Sub
ErrHnd:
    LastError.ErrNumber = Err.Number
    LastError.ErrDesc = Err.Description
    Err.Clear
    
    Set reg = Nothing
    Set curr_item = Nothing
    Resume CleanUp
End Sub

''
' ProcessItem
'   Takes a Dictionary Item and turns it into an Ini String
' Params
'   Key        As String  - The Item's Name
'   Val        As Variant - A varint representing data to serialize. Complete objects or
'                           types cannot be serialized.
'   ParentName As String  - Default = "". If provided this function assumes the item
'                           is a nested section
' Returns
'   String                - A string represention of an Ini Data Line.
Private Function ProcessItem(Key As String, Val As Variant, Optional ParentName As String = "") As String
    LastError.ErrNumber = 0
    LastError.ErrDesc = ""
    LastError.ErrFunc = "IniReader.ProcessItem"
    On Error GoTo ErrHnd
    
    Dim tmp As Variant, v_out As String
    
    If Key Like "[C|c][O|o][M|m][M|m][E|e][N|n][T|t]*" Then
        v_out = CStr(Val) & vbCrLf
    ElseIf TypeOf Val Is Collection Then
        For Each tmp In Val
            v_out = v_out & ProcessItem(Key, tmp)
        Next
    ElseIf TypeOf Val Is Object  And TypeName(Val) = "Dictionary" Then
        If ParentName <> "" Then
            v_out = "[" & ParentName & "." & VBA.Trim(Key) & "]" & vbCrLf
        Else
            v_out = "[" & VBA.Trim(Key) & "]" & vbCrLf
        End If
        
        For Each tmp In Val.Keys
            If ParentName <> "" Then
                v_out = v_out & ProcessItem(CStr(tmp), Val(tmp), ParentName & "." & VBA.Trim(Key))
            Else
                v_out = v_out & ProcessItem(CStr(tmp), Val(tmp), VBA.Trim(Key))
            End If
        Next
    ElseIf IsEmpty(Val) Or IsNull(Val) Then
        v_out = VBA.Trim(Key) & " = " & vbCrLf
    Else
        v_out = VBA.Trim(Key) & " = " & ConvertToValueStr(CStr(Val)) & vbCrLf
    End If
    
    ProcessItem = v_out
    
CleanUp:
    On Error GoTo 0
    Exit Function
ErrHnd:
    LastError.ErrNumber = Err.Number
    LastError.ErrDesc = Err.Description
    Err.Clear
    
    Resume CleanUp
End Function

''
' GetCurrSection
'   Fetches the current section object from an Ini Dictionary.
' Params
'   IniDoc         As String - An Ini Dictionary Object to search.
'   currItem       As Object - The Last Retrived Ini Section Object.
'   IniSecFullName As String - The Section name.
' Returns
'   Object                   - An Ini Section Object
Private Function GetCurrSection(IniDoc As Object, currItem As Object, IniSecFullName As String) As Object
    LastError.ErrNumber = 0
    LastError.ErrDesc = ""
    LastError.ErrFunc = "IniReader.GetCurrSection"
    On Error GoTo ErrHnd
    
    Dim v_out As Object, s_names() As String, i As Long
    
    Set v_out = IniDoc
    
    s_names = VBA.Split(IniSecFullName, ".")
    
    If VBA.Trim(s_names(0)) = "" Then
        Set v_out = currItem
    End If
    
    For i = LBound(s_names) To UBound(s_names)
        If VBA.Trim(s_names(i)) <> "" Then
            If Not v_out.Exists(s_names(i)) Then
                v_out.Add s_names(i), CreateObject("Scripting.Dictionary")
            End If
            
            Set v_out = v_out(s_names(i))
        End If
    Next
    
    Set GetCurrSection = v_out
    
CleanUp:
    On Error GoTo 0
    Exit Function
ErrHnd:
    LastError.ErrNumber = Err.Number
    LastError.ErrDesc = Err.Description
    Err.Clear
    
    Resume CleanUp
End Function

''
' ConvertToValueItem
'   Converts a String to an Ini Value
' Params
'   Val As String - A String representation of an Ini Value.
' Returns
'   Variant       - An Ini Value
Private Function ConvertToValueItem(ByVal Val As String) As Variant
    LastError.ErrNumber = 0
    LastError.ErrDesc = ""
    LastError.ErrFunc = "IniReader.ConvertToValueItem"
    On Error GoTo ErrHnd
    
    Dim v_out As Variant, m_col As Object, match As Object, _
        matched As Boolean
    
    If reg Is Nothing Then Set reg = CreateObject("VBScript.RegExp")
    
    With reg
        .MultiLine = True
        .IgnoreCase = False
        .Global = True
    End With
    matched = False
    Val = VBA.Trim(Val)
    
    reg.Pattern = NUMBER_PATTERN
    If reg.Test(Val) Then
        matched = True
        Set m_col = reg.Execute(Val)
        
        On Error Resume Next
        For Each match In m_col
            v_out = VBA.Replace(match.SubMatches(0), "0x", "&h")
            v_out = CDbl(v_out)
        Next
        On Error GoTo 0
        
        If Err.Number <> 0 Then
            Err.Clear
            matched = False
        End If
    End If
    
    reg.Pattern = BOOL_PATTERN
    If reg.Test(Val) Then
        matched = True
        Set m_col = reg.Execute(Val)
        
        On Error Resume Next
        For Each match In m_col
            v_out = CBool(match.SubMatches(0))
        Next
        On Error GoTo 0
        
        If Err.Number <> 0 Then
            Err.Clear
            matched = False
        End If
    End If
    
    reg.Pattern = NULL_PATTERN
    If reg.Test(Val) Or Val = "" Then
        matched = True
        v_out = Empty
    End If
    
    If Not matched Then
        v_out = Val
        If VBA.Left(v_out, 1) = """" Or VBA.Left(v_out, 1) = "'" Then
            v_out = VBA.Mid(v_out, 2, Len(v_out) - 1)
        End If
        If VBA.Right(v_out, 1) = """" Or VBA.Right(v_out, 1) = "'" Then
            If VBA.Right(v_out, 2) <> "\""" And VBA.Right(v_out, 2) <> "\'" Then
                v_out = VBA.Mid(v_out, 1, Len(v_out) - 1)
            End If
        End If
        
        v_out = DecodeEscapes(CStr(v_out))
    End If
    
    ConvertToValueItem = v_out
    
CleanUp:
    Set m_col = Nothing
    Set match = Nothing
    
    On Error GoTo 0
    Exit Function
ErrHnd:
    LastError.ErrNumber = Err.Number
    LastError.ErrDesc = Err.Description
    Err.Clear
    
    Set reg = Nothing
    Resume CleanUp
End Function

''
' ConvertToValueStr
'   Converts a Variant to an Ini String Value
' Params
'   Val As Variant - A vba value.
' Returns
'   String         - An Ini String Value
Private Function ConvertToValueStr(ByVal Val As Variant) As String
    LastError.ErrNumber = 0
    LastError.ErrDesc = ""
    LastError.ErrFunc = "IniReader.ConvertToValueStr"
    On Error GoTo ErrHnd
    
    Dim v_out As String, m_col As Object, match As Object, _
        matched As Boolean
    
    If reg Is Nothing Then Set reg = CreateObject("VBScript.RegExp")
    
    With reg
        .MultiLine = True
        .IgnoreCase = False
        .Global = True
    End With
    
    matched = False
    
    If VBA.IsEmpty(Val) Or VBA.IsNull(Val) Then
        v_out = ""
    ElseIf VarType(Val) = vbBoolean Then
        v_out = CStr(CBool(Val))
    ElseIf IsNumeric(Val) Then
        v_out = CStr(VBA.Trim(Val))
    Else
        Val = CStr(Val)
        
        reg.Pattern = NULL_PATTERN
        If reg.Test(Val) Or Val = "" Then
            matched = True
            v_out = ""
        End If
        
        reg.Pattern = NUMBER_PATTERN
        If reg.Test(Val) Then
            matched = True
            v_out = CStr(VBA.Trim(Val))
        End If
        
        If Not matched Then
            v_out = Val
        End If
        
        v_out = EncodeEscapes(CStr(v_out))
    
        If Len(VBA.Trim(v_out)) <> Len(v_out) Then
            v_out = """" & v_out & """"
        End If
    End If
    
    ConvertToValueStr = v_out
    
CleanUp:
    Set m_col = Nothing
    Set match = Nothing
    
    On Error GoTo 0
    Exit Function
ErrHnd:
    LastError.ErrNumber = Err.Number
    LastError.ErrDesc = Err.Description
    Err.Clear
    
    Set reg = Nothing
    Resume CleanUp
End Function

''
' DecodeEscapes
'   Turns escaped unicode character in valid string data.
' Params
'   Val As String - A string with escaped unicode characters or special ini characters.
' Returns
'   String        - A string with the escaped decoded.
Private Function DecodeEscapes(ByVal Val As String) As String
    LastError.ErrNumber = 0
    LastError.ErrDesc = ""
    LastError.ErrFunc = "IniReader.DecodeEscapes"
    On Error GoTo ErrHnd
    
    Dim v_out As String, m_col As Object, match As Object
    
    If reg Is Nothing Then Set reg = CreateObject("VBScript.RegExp")
    v_out = VBA.Trim(Val)
    
    With reg
        .MultiLine = True
        .IgnoreCase = False
        .Global = True
        .Pattern = ESCAPE_PATTERN
    End With
    
    If reg.Test(v_out) Then
        Set m_col = reg.Execute(v_out)
        
        For Each match In m_col
            Select Case True
                Case "\\" = match.Value
                    v_out = VBA.Replace(v_out, match.Value, "\", 1, 1)
                Case "\'" = match.Value
                    v_out = VBA.Replace(v_out, match.Value, "'", 1, 1)
                Case "\""" = match.Value
                    v_out = VBA.Replace(v_out, match.Value, """", 1, 1)
                Case "\0" = match.Value
                    v_out = VBA.Replace(v_out, match.Value, VBA.Chr(0), 1, 1)
                Case "\t" = match.Value
                    v_out = VBA.Replace(v_out, match.Value, vbTab, 1, 1)
                Case "\r" = match.Value
                    v_out = VBA.Replace(v_out, match.Value, vbCr, 1, 1)
                Case "\n" = match.Value
                    v_out = VBA.Replace(v_out, match.Value, vbLf, 1, 1)
                Case "\;" = match.Value
                    v_out = VBA.Replace(v_out, match.Value, ";", 1, 1)
                Case "\=" = match.Value
                    v_out = VBA.Replace(v_out, match.Value, "=", 1, 1)
                Case "\#" = match.Value
                    v_out = VBA.Replace(v_out, match.Value, "#", 1, 1)
                Case match.Value Like "\x[0-9a-zA-Z][0-9a-zA-Z][0-9a-zA-Z][0-9a-zA-Z]"
                    v_out = VBA.Replace(v_out, match.Value, VBA.Chr(VBA.Replace(match.Value, "\x", "&h")), 1, 1)
            End Select
        Next
    End If
    
    DecodeEscapes = v_out
    
CleanUp:
    Set m_col = Nothing
    Set match = Nothing
    
    On Error GoTo 0
    Exit Function
ErrHnd:
    LastError.ErrNumber = Err.Number
    LastError.ErrDesc = Err.Description
    Err.Clear
    
    Set reg = Nothing
    Resume CleanUp
End Function

''
' EncodeEscapes
'   Encodes un-escaped unicode characters and special characters.
' Params
'   Val As String - A string with un-escaped unicode characters or special ini characters.
' Returns
'   String        - A string with escaped unicode characters or special ini characters.
Private Function EncodeEscapes(ByVal Val As String) As String
    LastError.ErrNumber = 0
    LastError.ErrDesc = ""
    LastError.ErrFunc = "IniReader.DecodeEscapes"
    On Error GoTo ErrHnd
    
    Dim v_out As String, m_col As Object, match As Object
    
    If reg Is Nothing Then Set reg = CreateObject("VBScript.RegExp")
    v_out = Val
    
    With reg
        .MultiLine = True
        .IgnoreCase = False
        .Global = True
        .Pattern = NON_ESCAPED_PATTERN
    End With
    
    If reg.Test(v_out) Then
        Set m_col = reg.Execute(v_out)
        
        For Each match In m_col
            Select Case True
                Case "\" = match.SubMatches(0)
                    v_out = VBA.Replace(v_out, match.SubMatches(0), "\\", 1, 1)
                Case "'" = match.SubMatches(0)
                    v_out = VBA.Replace(v_out, match.SubMatches(0), "\'", 1, 1)
                Case """" = match.SubMatches(0)
                    v_out = VBA.Replace(v_out, match.SubMatches(0), "\""", 1, 1)
                Case VBA.Chr(0) = match.SubMatches(0)
                    v_out = VBA.Replace(v_out, match.SubMatches(0), "\0", 1, 1)
                Case vbTab = match.SubMatches(0)
                    v_out = VBA.Replace(v_out, match.SubMatches(0), "\t", 1, 1)
                Case vbCr = match.SubMatches(0)
                    v_out = VBA.Replace(v_out, match.SubMatches(0), "\r", 1, 1)
                Case vbLf = match.SubMatches(0)
                    v_out = VBA.Replace(v_out, match.SubMatches(0), "\n", 1, 1)
                Case ";" = match.SubMatches(0)
                    v_out = VBA.Replace(v_out, match.SubMatches(0), "\;", 1, 1)
                Case "=" = match.SubMatches(0)
                    v_out = VBA.Replace(v_out, match.SubMatches(0), "\=", 1, 1)
                Case "#" = match.SubMatches(0)
                    v_out = VBA.Replace(v_out, match.SubMatches(0), "\#", 1, 1)
                Case match.SubMatches(1) <> ""
                    v_out = VBA.Replace( _
                        v_out, _
                        match.SubMatches(1), _
                        "\x" & VBA.Right("0000" & VBA.Hex(VBA.AscW(match.SubMatches(1))), 4), _
                        1, _
                        1 _
                    )
            End Select
        Next
    End If
    
    EncodeEscapes = v_out
    
CleanUp:
    Set m_col = Nothing
    Set match = Nothing
    
    On Error GoTo 0
    Exit Function
ErrHnd:
    LastError.ErrNumber = Err.Number
    LastError.ErrDesc = Err.Description
    Err.Clear
    
    Set reg = Nothing
    Resume CleanUp
End Function

''
' bArrLen
'   A helper function that returns the lenght of an array or zero if it
'   is uninitilzed.
' Params
'   bytes() As Byte - A byte array to get the lenght of.
' Returns
'   Long            - The lenght of the array or Zero.
Private Function bArrLen(bytes() As Byte) As Long
    On Error Resume Next
    bArrLen = UBound(bytes) - LBound(bytes) + 1
End Function

''
' StringFromUTF8
'   A helper function that converts a UTF-8 byte array in a vba string (UTF-16).
' Params
'   bytes() As Byte - A byte array containing UTF-8 encoded data.
' Returns
'   String          - A UTF-16 VBA string.
Private Function StringFromUTF8(bytes() As Byte) As String
    Dim n_bytes As Long
    Dim n_chars As Long
    Dim v_out() As Byte
    
    n_bytes = bArrLen(bytes)
    
    If n_bytes > 0 Then
        n_chars = MultiByteToWideChar(CP_UTF8, 0&, VarPtr(bytes(0)), n_bytes, 0^, 0&)
        
        ReDim v_out(0 To n_chars * 2 - 1) '' VBA Strings are Unicode, double-wide, or UTF-16, therefor char = 2 bytes
        
        n_chars = MultiByteToWideChar(CP_UTF8, 0&, VarPtr(bytes(0)), n_bytes, VarPtr(v_out(0)), n_chars)
    End If
    
    StringFromUTF8 = v_out
End Function

''
' UTF8FromString
'   A helper function the encodes string data into UTF-8 data.
' Params
'   str As String - A UTF-16 VBA string.
' Returns
'   Byte()        - A byte array containing UTF-8 encoded data.
Function UTF8FromString(str As String) As Byte()
    Dim n_bytes As Long
    Dim n_chars As Long
    Dim v_out() As Byte
    
    n_chars = Len(str)
    
    If n_chars > 0 Then
        n_bytes = WideCharToMultiByte(CP_UTF8, 0&, StrPtr(str), n_chars, 0^, 0&, 0^, 0^)
        
        ReDim v_out(0 To n_bytes - 1)
        
        n_bytes = WideCharToMultiByte(CP_UTF8, 0&, StrPtr(str), n_chars, VarPtr(v_out(0)), n_bytes, 0^, 0^)
    End If
    
    UTF8FromString = v_out
End Function
