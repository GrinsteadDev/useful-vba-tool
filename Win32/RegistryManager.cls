VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RegistryManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
''
'This Module is designed to be used as a standard interface to Update, Change, and Query the Windows Registry.
'This Module has complete error handling to allow easy debugging while using this Module as a Class Object.
'This Module accomplishes the following tasks.
'   - Verify the presence of registry keys.
'   - Verify the access permissions for a user.
'   - Create, modify, and delete registry keys.
'   - Create, modify, and delete named attributes.
'   - Read, write, and delete data values.
'
'Public Functions
'   - QueryRegistry              - Fetches a registry key as an object map.
'   - UpdateRegistryKey          - Creates or Updates an attribute on a registry key.
'   - CreateRegistryKey          - Creates a registry key.
'   - RemoveRegistryKey          - Deletes a registry key.
'   - RemoveRegistryKeyAttribute - Removes an attribute from a registry key.
'   - HasAccess                  - Verifies that the user has the specified access permissions.
'   - KeyExists                  - Verifies that the requested key exists.
'
'Public Settings
'   - SlientErrors - Determines if Registry errors throw.
'   - LogError     - Determines if Registry errors are logged.
'   - LogFile      - The log file that this module is to save/log errors to.
'Public Constants
'   - Enum AccessFlags   - The following access permission values used by the Windows Registry
'   - Enum AttributeType - The Registry data types.

''Registry Win32 Constants
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS = &H80000003
Private Const HKEY_CURRENT_CONFIG = &H80000005

Private Const REG_NONE = 0
Private Const REG_SZ = 1
Private Const REG_EXPAND_SZ = 2
Private Const REG_BINARY = 3
Private Const REG_DWORD = 4
Private Const REG_DWORD_LITTLE_ENDIAN = 4
Private Const REG_DWORD_BIG_ENDIAN = 5
Private Const REG_LINK = 6
Private Const REG_MULTI_SZ = 7
Private Const REG_RESOURCE_LIST = 8
Private Const REG_FULL_RESOURCE_DESCRIPTOR = 9
Private Const REG_RESOURCE_REQUIREMENTS_LIST = 10
Private Const REG_QWORD = 11
Private Const REG_QWORD_LITTLE_ENDIAN = 11

Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_READ_WRITE = KEY_QUERY_VALUE Or KEY_SET_VALUE
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE = &H20
Private Const KEY_DELETE = &HA0000
Private Const KEY_READ = &H20000
Private Const KEY_WRITE_DAC = &H40000
Private Const KEY_WRITE_OWNER = &H80000

Private Const REG_OPTION_RESERVED = &H0
Private Const REG_OPTION_NON_VOLATILE = &H0
Private Const REG_OPTION_VOLATILE = &H1
Private Const REG_OPTION_CREATE_LINK = &H2
Private Const REG_OPTION_BACKUP_RESTORE = &H4
Private Const REG_OPTION_OPEN_LINK = &H8
Private Const REG_LEGAL_OPTION = REG_OPTION_RESERVED Or REG_OPTION_NON_VOLATILE Or REG_OPTION_CREATE_LINK Or REG_OPTION_BACKUP_RESTORE Or REG_OPTION_OPEN_LINK
Private Const REG_CREATED_NEW_KEY = 1
Private Const REG_OPENED_EXISTING_KEY = 2
Private Const REG_WHOLE_HIVE_VOLATILE = &H1
Private Const REG_REFRESH_HIVE = &H2
Private Const REG_NO_LAZY_FLUSH = &H4
Private Const REG_FORCE_RESTORE = &H8

Private Const ERROR_FILE_NOT_FOUND = &H2 'Error Value associated with an incorrect path.
Private Const ERROR_ACCESS_DENIED = &H5
Private Const ERROR_ALREADY_EXISTS = &HB7

Private Const MAX_KEY_LENGTH = 514
Private Const MAX_VALUE_LENGTH = 32768

''Module Constants
Private Const ErrHelp = "https://learn.microsoft.com/en-us/windows/win32/debug/system-error-codes--0-499-"
Private Const MoudleName = "RegistryManager"

''Public Enums
Public Enum AccessFlags
    Default = KEY_READ_WRITE
    QueryValue = KEY_QUERY_VALUE
    SetValue = KEY_SET_VALUE
    CreateSubKeys = KEY_CREATE_SUB_KEY
    EnumSubKeys = KEY_ENUMERATE_SUB_KEYS
    KeyNotify = KEY_NOTIFY
    CreateKey = KEY_CREATE
    DeleteKey = KEY_DELETE
    ReadControl = KEY_READ
    WriteDac = KEY_WRITE_DAC
    WriteOwner = KEY_WRITE_OWNER
End Enum

Public Enum AttributeType
    RegString = REG_SZ
    RegExpandedString = REG_EXPAND_SZ
    RegBinary = REG_BINARY
    RegInteger = REG_DWORD
    RegLong = REG_QWORD
    RegMultiString = REG_MULTI_SZ
    RegInt32 = REG_DWORD
    RegInt64 = REG_QWORD
    RegByteArray = REG_BINARY
    RegDword = REG_DWORD
    RegQword = REG_QWORD
End Enum

''External API Kernal32
Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100&
Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000&
Private Const FORMAT_MESSAGE_FROM_HMODULE = &H800&
Private Const FORMAT_MESSAGE_FROM_STRING = &H400&
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000&
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200&

Private Const LANG_DEFAULT = 0

Private Declare PtrSafe Function FormatMessageA Lib "Kernel32.dll" ( _
    ByVal dwFlags As Long, _
    ByVal lpSource As Any, _
    ByVal dwMessageId As Long, _
    ByVal dwLanguageId As Long, _
    ByVal lpBuffer As String, _
    ByVal nSize As Long, _
    ByVal va_list As Any _
) As Long

Private WMIReg_ As Object

''Module Settings
''
'Determines if Registry errors throw.
'Boolean
'   Default False - Registry errors are thrown using the err.Raise Method.
'           True  - Registry errors are ignored and printed to the immediate window.
Public SlientErrors As Boolean
''
'Determines if Registry errors are logged.
'Boolean
'   Default False - Registry errors are not logged
'           True  - Registry errors are logged in the specified log file.
Public LogError As Boolean
''
'The log file that this module is to save/log errors to.
'String
'   Default ""             - Registry errors are not logged
'           "<file-path>"  - Registry errors are logged in the specified log file at <file-path>.
Public LogFile As String

''
'Gets the Windows Error Message for the StdRegProv Object's Error Codes.
'Param
'   ErrorNum Long - The System Error Number.
'Return
'   String - The System Defined Error Message/Description.
Private Function GetErrorMsg( _
    ErrorNum As Long _
) As String
    Const MAX_MESSAGE = 1024
    Const Flags = FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS
   
    Dim oStr As String * MAX_MESSAGE
    Dim hResult
   
    hResult = FormatMessageA( _
        Flags, _
        0^, _
        ErrorNum, _
        LANG_DEFAULT, _
        oStr, _
        MAX_MESSAGE, _
        0^ _
    )
   
    If hResult = 0 Then
        Debug.Print "FormatMesageA Error: "; Err.LastDllError
        oStr = vbNullString
    End If
   
    GetErrorMsg = VBA.Trim(VBA.Replace(oStr, VBA.Chr$(0), ""))
End Function

''
'This Module's error log. This function logs errors the specified file in the format of:
'   <username>--<datetime>  -  <ErrorNum> <ErrorMsg> @ <FunctionName>
'Param
'   ErrorNum     - The Error Number to log
'   ErrorMsg     - The Error Message to log
'   FunctionName - The Function Name where the error occured
'   UserName     - The Current User
'Return
'   void
Private Sub Log( _
    ErrorNum As Long, _
    ErrorMsg As String, _
    FunctionName As String, _
    UserName As String _
)
    If Not LogError Then Exit Sub
    If VBA.Trim(LogFile) = "" Then Exit Sub
   
    Dim fn As Integer
    Dim currDate As String
   
    fn = FreeFile
    currDate = VBA.Format(VBA.Now, "--mmm-dd-yyyy hh:mm:ss") & VBA.Right(VBA.Format(VBA.Timer, "0.000"), 4)
   
    Open LogFile For Append Access Write Lock Write As #fn
   
    Print #fn, VBA.Replace(UserName & currDate & "  -  0x" & VBA.Right("00000000" & VBA.Hex(ErrorNum), 8) & "  " & ErrorMsg & " @ " & FunctionName, vbCrLf, "")
   
    Close #fn
End Sub

''
'Slices an Array
'Param
'            Value  Variant - The array that is to be sliced
'            Start  Long    - The index to start the slice at
'   Optional Length Long    - The length of the new array
'Return
'   Variant - The new array or Empty
Private Function SliceArray( _
    Value As Variant, _
    Start As Long, _
    Optional Length As Long _
) As Variant
    Dim out As Variant
    Dim i As Long
    Dim maxLen As Long
       
    If IsArray(Value) Then
        If Start < LBound(Value) Then GoTo errHnd
       
        maxLen = UBound(Value) - LBound(Value) - Start + 1
        If Length < 1 Then Length = maxLen
        If Length > maxLen Then Length = maxLen
       
        ReDim out(0 To Length - 1)
       
        For i = Start To Start + Length - 1
            out(i - Start) = Value(i)
        Next i
    End If
   
    GoTo endf
errHnd:
   
endf:
    SliceArray = out
End Function

''
'Creates and Retrieves the WMIReg_ object from the Module's private scope.
Private Function WMIReg() As Object
    If WMIReg_ Is Nothing Then Set WMIReg_ = CreateObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
    Set WMIReg = WMIReg_
End Function

''
'Turns a Registry Path String into a Parts Array
'Param
'   Path String - The Registry Path's String Representation.
'Return
'   Variant Array(0 To 1)
'       [0] - The HKey Value, defaults to HKEY_LOCAL_MACHINE
'       [1] - The KeyName String
Private Function parsePath( _
    Path As String _
) As Variant
    Dim parts
   
    If VBA.Left(Path, 1) = "\" Then Path = VBA.Mid(Path, 2)
    If VBA.InStr(Path, "\") < 1 Then Path = Path & "\"
   
   
    parts = VBA.Split(Path, "\", 2)
    If VBA.Right(parts(1), 1) = "\" Then parts(1) = VBA.Mid(parts(1), 1, VBA.Len(parts(1)) - 1)
   
    If IsArray(parts) Then
        Select Case VBA.UCase(parts(0))
            Case "HKEY_CLASSES_ROOT"
                parts(0) = HKEY_CLASSES_ROOT
            Case "HKEY_CURRENT_USER"
                parts(0) = HKEY_CURRENT_USER
            Case "HKEY_LOCAL_MACHINE"
                parts(0) = HKEY_LOCAL_MACHINE
            Case "HKEY_USERS"
                parts(0) = HKEY_USERS
            Case "HKEY_CURRENT_CONFIG"
                parts(0) = HKEY_CURRENT_CONFIG
            Case Else
                parts(1) = parts(0) & "\" & parts(1)
                parts(0) = HKEY_LOCAL_MACHINE
        End Select
    End If
   
    parsePath = parts
End Function

''
'Retrieves a Value from the Registry
'Param
'   PathParts Variant - An Array[0 To 1].
'                       [0] The HKey Value as a Number/Number String
'                       [1] The Key Name as a String
'   Key       String  - The Name of the Value to Retrieve From the Registry Key
'   KeyType   Integer - The Key Type As a REG_ Constant. This value should be retrieved from the WMI Registry
'                       Object's EnumValues Function.
'                       Values:
'                           REG_SZ
'                           REG_EXPAND_SZ
'                           REG_BINARY
'                           REG_DWORD
'                           REG_MULTI_SZ
'                           REG_QWORD
'Return
'   Variant - The Value from the Registry
Private Function getRegValue( _
    pathParts As Variant, _
    KeyAttribute As String, _
    KeyAttributeType As Integer _
) As Variant
    Dim out, hResult As Long
   
    Select Case KeyAttributeType
        Case REG_SZ
            hResult = WMIReg.GetStringValue(CLng(pathParts(0)), CStr(pathParts(1)), KeyAttribute, out)
        Case REG_EXPAND_SZ
            hResult = WMIReg.GetExpandedStringValue(CLng(pathParts(0)), CStr(pathParts(1)), KeyAttribute, out)
        Case REG_BINARY
            hResult = WMIReg.GetBinaryValue(CLng(pathParts(0)), CStr(pathParts(1)), KeyAttribute, out)
        Case REG_DWORD
            hResult = WMIReg.GetDWORDValue(CLng(pathParts(0)), CStr(pathParts(1)), KeyAttribute, out)
        Case REG_MULTI_SZ
            hResult = WMIReg.GetMultiStringValue(CLng(pathParts(0)), CStr(pathParts(1)), KeyAttribute, out)
        Case REG_QWORD
            hResult = WMIReg.GetQWORDValue(CLng(pathParts(0)), CStr(pathParts(1)), KeyAttribute, out)
        Case Else
            hResult = -1
    End Select
   
endf:
    getRegValue = out
End Function

''
'The QueryRegistrary function is designed to take a Registry string path and return
'the desired key entry as a Scripting.Dictionary Object. With the expection of the
'value at the "SubKeys" table entry all other key=vaule pairs are maped to the Registry Key's
'key=value pairs.
'
'Param
'            Path          String          - A Valid Registry
'   Optional ReferAsObject Boolean [False] - Determins wheather the map object returned creates
'                                            Map objects for the subkeys
'                                            False - SubKey Collection as Path String
'                                            True  - SubKey Collection as Map Object
'                                            Warning, setting this value to true will cause ALL subkeys
'                                            to be maped as objects. Depending on where the request key
'                                            is in the Registry Tree this could be an expensive process.
'Return
'   On Success
'       Object Scripting.Dictionary
'           Keys defined as
'               SubKeys String[] - An Array of the Registry entry's sub keys or Empty
'               <Registry Value Name> - An undefined number of entries that reflects the
'                                    Registry values for the supplied key
'   On Failure - Nothing
'   Notes: While this object reflects the registry it is NOT a reference, changes made to the returned
'          object will NOT change the registry.
Public Function QueryRegistry( _
    Path As String, _
    Optional ReferAsObject As Boolean = False _
) As Object
    Dim parts As Variant: parts = parsePath(Path)
    Dim Keys As Variant
    Dim hResult As Long
    Dim map As Object
   
    If Not HasAccess(Path, ReadControl) Then
        hResult = ERROR_ACCESS_DENIED
        GoTo regErr
    End If
   
    hResult = WMIReg.EnumKey(CLng(parts(0)), CStr(parts(1)), Keys)
   
    If hResult = 0 Then
        Dim vKeys As Variant
        Dim vTypes As Variant
        Dim sKey As Variant
        Dim sOKeys As Variant
       
        Set map = CreateObject("Scripting.Dictionary")
       
        If ReferAsObject And IsArray(Keys) Then
            ReDim sOKeys(0)
            For Each sKey In Keys
                Set sOKeys(UBound(sOKeys)) = QueryRegistry(Path & "\" & sKey, ReferAsObject)
                ReDim Preserve sOKeys(UBound(sOKeys) + 1)
            Next
            ReDim Preserve sOKeys(UBound(sOKeys) - 1)
           
            map.Add "SubKeys", sOKeys
        ElseIf IsArray(Keys) Then
            ReDim sOKeys(0)
            For Each sKey In Keys
                sOKeys(UBound(sOKeys)) = Path & "\" & sKey
                ReDim Preserve sOKeys(UBound(sOKeys) + 1)
            Next
            ReDim Preserve sOKeys(UBound(sOKeys) - 1)
           
            map.Add "SubKeys", sOKeys
        Else
            map.Add "SubKeys", Empty
        End If
       
        hResult = WMIReg.EnumValues(CLng(parts(0)), parts(1), vKeys, vTypes)
       
        If hResult = 0 Then
            If IsArray(vKeys) Then
                Dim i As Long
               
                For i = 0 To UBound(vKeys)
                    map.Add CStr(vKeys(i)), getRegValue(parts, CStr(vKeys(i)), CInt(vTypes(i)))
                Next i
            End If
        End If
    End If
   
    If hResult <> 0 Then GoTo regErr
   
    GoTo endf
regErr:
    Set map = Nothing
    Dim errMsg As String
   
    errMsg = GetErrorMsg(hResult)
   
    If LogError Then
        Log hResult, errMsg, MoudleName & ".QueryRegistry", VBA.Environ("UserName")
    End If
    If SlientErrors Then
        Debug.Print "System Error 0x"; VBA.Hex(hResult)
        Debug.Print errMsg
    Else
        Err.Raise hResult, "StdRegProv", errMsg, ErrHelp
    End If
endf:
    Set QueryRegistry = map
End Function

''
'This function updates/creates an attribute value on a Registry key.
'Param
'   Path              - The Registry Key Path's String Representation.
'   KeyAttributeName  - The Attribute Name to Create or Change.
'   KeyAttributeValue - The Value to be added to the Registry.
'   KeyAttributeType  - The Value type.
'                       Values As AttributeType:
'                           RegString = REG_SZ                - System String, same as VBA String
'                           RegExpandedString = REG_EXPAND_SZ - System Expanded String. A String datatype with an environment Variable.
'                                                               Example: "%USERPROFILE%/MyApp/MyData"
'                           RegBinary = REG_BINARY            - A byte array, same as VBA Byte()
'                           RegByteArray = REG_BINARY
'                           RegInteger = REG_DWORD            - An Int32 Value, same as VBA Integer
'                           RegInt32 = REG_DWORD
'                           RegDword = REG_DWORD
'                           RegLong = REG_QWORD               - An Int64 Value, same as VBA Long
'                           RegInt64 = REG_QWORD
'                           RegQword = REG_QWORD
'                           RegMultiString = REG_MULTI_SZ     - An Array of Strings
'Return
'   Boolean
'       True on Success
'       False on Failure
Public Function UpdateRegistryKey( _
    Path As String, _
    KeyAttributeName As String, _
    KeyAttributeValue As Variant, _
    KeyAttributeType As AttributeType _
) As Boolean
    Dim parts As Variant: parts = parsePath(Path)
    Dim hResult As Long
    Dim i
   
    If Not HasAccess(Path, SetValue) Then
        hResult = ERROR_ACCESS_DENIED
        GoTo regErr
    End If
   
    Select Case KeyAttributeType
        Case REG_SZ
            Dim Str As String
           
            Str = CStr(KeyAttributeValue)
           
            hResult = WMIReg.SetStringValue(CLng(parts(0)), CStr(parts(1)), KeyAttributeName, Str)
        Case REG_EXPAND_SZ
            Dim exStr As String
           
            exStr = CStr(KeyAttributeValue)
           
            hResult = WMIReg.SetExpandedStringValue(CLng(parts(0)), CStr(parts(1)), KeyAttributeName, exStr)
        Case REG_BINARY
            Dim uint8() As Byte
           
            ReDim uint8(0 To UBound(KeyAttributeValue) - LBound(KeyAttributeValue))
           
            For i = 0 To UBound(uint8)
                uint8(i) = CByte(KeyAttributeValue(i + LBound(KeyAttributeValue)))
            Next i
           
            hResult = WMIReg.SetBinaryValue(CLng(parts(0)), CStr(parts(1)), KeyAttributeName, uint8)
        Case REG_DWORD
            Dim dword As Integer
           
            dword = CInt(KeyAttributeValue)
           
            hResult = WMIReg.SetDWORDValue(CLng(parts(0)), CStr(parts(1)), KeyAttributeName, dword)
        Case REG_MULTI_SZ
            hResult = WMIReg.SetMultiStringValue(CLng(parts(0)), CStr(parts(1)), CStr(KeyAttributeName), KeyAttributeValue)
        Case REG_QWORD
            Dim qword As Long
           
            qword = CLng(KeyAttributeValue)
           
            hResult = WMIReg.SetQWORDValue(CLng(parts(0)), CStr(parts(1)), KeyAttributeName, qword)
        Case Else
            hResult = -1
    End Select
   
    GoTo endf
regErr:
    Dim errMsg As String
   
    errMsg = GetErrorMsg(hResult)
   
    If LogError Then
        Log hResult, errMsg, MoudleName & ".UpdateRegistryKey", VBA.Environ("UserName")
    End If
    If SlientErrors Then
        Debug.Print "System Error 0x"; VBA.Hex(hResult)
        Debug.Print errMsg
    Else
        Err.Raise hResult, "StdRegProv", errMsg, ErrHelp
    End If
endf:
    UpdateRegistryKey = (hResult = 0)
End Function

''
'Creates a new Registry Key at the specified Key Path. The CreateKey method creates all subkeys specified
'in the path that do not exist.
'Params
'   Path - The Registry Key Path to be created.
'Return
'   Boolean
'       True on Success
'       False on Failure
Public Function CreateRegistryKey( _
    Path As String _
) As Boolean
    Dim pathParts As Variant: pathParts = parsePath(Path)
    Dim pathArr As Variant
    Dim pathTmp As String
    Dim i
    Dim hResult As Long
   
    If VBA.Trim(pathParts(1)) = "" Or KeyExists(Path) Then
        hResult = ERROR_ALREADY_EXISTS
        GoTo regErr
    End If
   
    ReDim Preserve pathParts(0 To 2)
    pathArr = VBA.Split(Path, "\")
   
    For i = 0 To UBound(pathArr)
       
        If Not KeyExists(pathTmp & "\" & CStr(pathArr(i))) Then
            If VBA.Left(pathTmp, 1) = "\" Then pathTmp = VBA.Mid(pathTmp, 2)
            If VBA.InStr(pathTmp, "\") < 1 Then pathTmp = pathTmp & "\"
           
            pathParts(1) = VBA.Mid(pathTmp, VBA.InStr(1, pathTmp, "\") + 1)
            pathParts(2) = VBA.Join(SliceArray(pathArr, CLng(i)), "\")
           
            Exit For
        End If
        pathTmp = pathTmp & "\" & CStr(pathArr(i))
    Next i
   
    If Not HasAccess(VBA.Replace(Path, pathParts(2), ""), CreateKey Or CreateSubKeys) Then
        hResult = ERROR_ACCESS_DENIED
        GoTo regErr
    End If
   
    pathTmp = pathParts(1) & "\" & pathParts(2)
    If VBA.Left(pathTmp, 1) = "\" Then pathTmp = VBA.Mid(pathTmp, 2)
   
    hResult = WMIReg.CreateKey(CLng(pathParts(0)), pathTmp)
   
    If hResult <> 0 Then
        GoTo regErr
    End If
   
    GoTo endf
regErr:
    Dim errMsg As String
   
    errMsg = GetErrorMsg(hResult)
   
    If LogError Then
        Log hResult, errMsg, MoudleName & ".CreateRegistryKey", VBA.Environ("UserName")
    End If
    If SlientErrors Then
        Debug.Print "System Error 0x"; VBA.Hex(hResult)
        Debug.Print errMsg
    Else
        Err.Raise hResult, "StdRegProv", errMsg, ErrHelp
    End If
endf:
    CreateRegistryKey = (hResult = 0)
End Function

''
'Deletes the specified Registry Key
'Params
'   Path - The Registry Key Path to be deleted.
'Return
'   Boolean
'       True on success
'       False on Failure
Public Function RemoveRegistryKey( _
    Path As String _
) As Boolean
    Dim parts As Variant: parts = parsePath(Path)
    Dim hResult As Long
   
   
    If Not KeyExists(Path) Then
        hResult = ERROR_FILE_NOT_FOUND
        GoTo regErr
    End If
    If Not HasAccess(Path, DeleteKey) Or parts(1) = "" Then
        hResult = ERROR_ACCESS_DENIED
        GoTo regErr
    End If
   
    hResult = WMIReg.DeleteKey(CLng(parts(0)), CStr(parts(1)))
   
    If hResult <> 0 Then
        GoTo regErr
    End If
   
    GoTo endf
regErr:
    Dim errMsg As String
   
    errMsg = GetErrorMsg(hResult)
   
    If LogError Then
        Log hResult, errMsg, MoudleName & ".RemoveRegistryKey", VBA.Environ("UserName")
    End If
    If SlientErrors Then
        Debug.Print "System Error 0x"; VBA.Hex(hResult)
        Debug.Print errMsg
    Else
        Err.Raise hResult, "StdRegProv", errMsg, ErrHelp
    End If
endf:
    RemoveRegistryKey = (hResult = 0)
End Function

''
'The function Removes and Attribute from the specified Registry Key.
'Params
'   Path         - The String Path to the Registry Key.
'   KeyAttribute - The Name of the Attribute to Remove.
'Return
'   Boolean
'       True on Success
'       False on Failure
Public Function RemoveRegistryKeyAttribute( _
    Path As String, _
    KeyAttribute As String _
) As Boolean
    Dim parts As Variant: parts = parsePath(Path)
    Dim hResult As Long
   
    If Not KeyExists(Path) Then
        hResult = ERROR_FILE_NOT_FOUND
        GoTo regErr
    End If
    If Not HasAccess(Path, DeleteKey Or SetValue) Or parts(1) = "" Then
        hResult = ERROR_ACCESS_DENIED
        GoTo regErr
    End If
   
    hResult = WMIReg.DeleteValue(CLng(parts(0)), CStr(parts(1)), KeyAttribute)
   
    If hResult <> 0 Then
        GoTo regErr
    End If
   
    GoTo endf
regErr:
    Dim errMsg As String
   
    errMsg = GetErrorMsg(hResult)
   
    If LogError Then
        Log hResult, errMsg, MoudleName & ".RemoveRegistryKeyAttribute", VBA.Environ("UserName")
    End If
    If SlientErrors Then
        Debug.Print "System Error 0x"; VBA.Hex(hResult)
        Debug.Print errMsg
    Else
        Err.Raise hResult, "StdRegProv", errMsg, ErrHelp
    End If
endf:
    RemoveRegistryKeyAttribute = (hResult = 0)
End Function

''
'This function queries if the user has access to perform the action specified by the Flags Parameter
'Param
'   KeyPath String      - The String Path to the Registry Key.
'   Flags   AccessFlags - This parameter specifies the action for which access is requested for.
'                         Values:
'                            Default = KEY_READ_WRITE
'                            QueryValue = KEY_QUERY_VALUE
'                            SetValue = KEY_SET_VALUE
'                            CreateSubKeys = KEY_CREATE_SUB_KEY
'                            EnumSubKeys = KEY_ENUMERATE_SUB_KEYS
'                            KeyNotify = KEY_NOTIFY
'                            CreateKey = KEY_CREATE
'                            DeleteKey = KEY_DELETE
'                            ReadControl = KEY_READ
'                            WriteDac = KEY_WRITE_DAC
'                            WriteOwner = KEY_WRITE_OWNER
'Return
'   Boolean - True if user has access to perform the request action.
'             False if the user does not have access.
Public Function HasAccess( _
    KeyPath As String, _
    Optional Flags As AccessFlags = Default _
) As Boolean
    Dim parts As Variant: parts = parsePath(KeyPath)
    Dim out As Boolean
    Dim hResult As Long
   
    hResult = WMIReg.CheckAccess(CLng(parts(0)), CStr(parts(1)), Flags, out)
   
    If hResult <> 0 And hResult <> ERROR_ACCESS_DENIED Then
        out = False
        GoTo regErr
    End If
   
    GoTo endf
regErr:
    Dim errMsg As String
   
    errMsg = GetErrorMsg(hResult)
   
    If LogError Then
        Log hResult, errMsg, MoudleName & ".HasAccess", VBA.Environ("UserName")
    End If
    If SlientErrors Then
        Debug.Print "System Error 0x"; VBA.Hex(hResult)
        Debug.Print errMsg
    Else
        Err.Raise hResult, "StdRegProv", errMsg, ErrHelp
    End If
endf:
    HasAccess = out
End Function

''
'This function queries if the request registry key exists.
'Param
'   KeyPath String      - The String Path to the Registry Key.
'Return
'   Boolean - True if the registry path is valid.
'             False otherwise.
Public Function KeyExists( _
    KeyPath As String _
) As Boolean
    Dim parts As Variant: parts = parsePath(KeyPath)
    Dim hResult As Long
    Dim out As Boolean
   
    hResult = WMIReg.CheckAccess(CLng(parts(0)), CStr(parts(1)), 1, 0)
   
    If hResult = ERROR_FILE_NOT_FOUND Then
        out = False
    End If
    If hResult = 0 Or hResult = ERROR_ACCESS_DENIED Then
        out = True
    End If
   
    If hResult <> 0 And hResult <> ERROR_ACCESS_DENIED And hResult <> ERROR_FILE_NOT_FOUND Then
        out = False
        GoTo regErr
    End If
   
    GoTo endf
regErr:
    Dim errMsg As String
   
    errMsg = GetErrorMsg(hResult)
   
    If LogError Then
        Log hResult, errMsg, MoudleName & ".KeyExists", VBA.Environ("UserName")
    End If
    If SlientErrors Then
        Debug.Print "System Error 0x"; VBA.Hex(hResult)
        Debug.Print errMsg
    Else
        Err.Raise hResult, "StdRegProv", errMsg, ErrHelp
    End If
endf:
    KeyExists = out
End Function

