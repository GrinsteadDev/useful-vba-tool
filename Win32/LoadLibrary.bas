Attribute VB_Name = "LoadLibrary"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Load Library Script                                                                        '
'Purpose:                                                                                   '
'   Allows library files to be loaded into the VBE by file path.                            '
'Description:                                                                               '
'   This module is designed to allow VBA scripters to load library files that would         '
'   otherwise be unloadable due to the limitations of the VBE's declare function API.       '
'Usage:                                                                                     '
'   It is recommended that all calls to LoadLibrary are places inside of the Workbook_Open  '
'   function.                                                                               '
'   Call LoadLibrary("<file-path>")                                                         '
'   Call LibraryFunction1                                                                   '
'   Call LibraryFunction2                                                                   '
'Notes:                                                                                     '
'   It is important to note that before calling functions loaded by the VBE's               '
'   declare function APIs to access library functions you must call this module's           '
'   LoadLibrary Method with the supplied file path.                                         '
'   The VBE will perform the freelibrary functions on its own. Calling them manually will   '
'   cause the VBE to crash due to the library                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Const DONT_RESOLVE_DLL_REFERENCES = &H1&
Private Const LOAD_IGNORE_CODE_AUTHZ_LEVEL = &H10&
Private Const LOAD_LIBRARY_AS_DATAFILE = &H2&
Private Const LOAD_LIBRARY_AS_DATAFILE_EXCLUSIVE = &H40&
Private Const LOAD_LIBRARY_AS_IMAGE_RESOURCE = &H20&
Private Const LOAD_LIBRARY_SEARCH_APPLICATION_DIR = &H200&
Private Const LOAD_LIBRARY_SEARCH_DEFAULT_DIRS = &H1000&
Private Const LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR = &H100&
Private Const LOAD_LIBRARY_SEARCH_SYSTEM32 = &H800&
Private Const LOAD_LIBRARY_SEARCH_USER_DIRS = &H400&
Private Const LOAD_WITH_ALTERED_SEARCH_PATH = &H8&
Private Const LOAD_LIBRARY_REQUIRE_SIGNED_TARGET = &H80&
Private Const LOAD_LIBRARY_SAFE_CURRENT_DIRS = &H2000&

Public Enum LoadType
    RelativePath = 0
    FullPath = LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR Or LOAD_LIBRARY_SEARCH_DEFAULT_DIRS
    LibName = LOAD_LIBRARY_SEARCH_DEFAULT_DIRS
End Enum

Private Const CTRUE As Byte = 0
Private Const CFALSE As Byte = 1
Private Const MAX_PATH = 256
''Support for x86, x64, VBA7, and VBA6
#If Win64 Or VBA7 Then
Private Const NULL_PTR As LongPtr = 0^

Private Declare PtrSafe Function LoadLibraryExA Lib "Kernel32.dll" ( _
    ByVal lpLibFileName As String, _
    ByVal hFile As LongPtr, _
    ByVal dwFlags As Long _
) As LongPtr
Private Declare PtrSafe Function AddDllDirectory Lib "Kernel32.dll" ( _
    ByVal NewDirectory As LongPtr _
) As LongPtr
Private Declare PtrSafe Function SetDefaultDllDirectories Lib "Kernel32.dll" ( _
    ByVal DirectoryFlags As Long _
) As Byte
Private Declare PtrSafe Function GetFullPathNameA Lib "Kernel32.dll" ( _
    ByVal lpFileName As String, _
    ByVal nBufferLength As Long, _
    ByVal lpBuffer As String, _
    ByVal lpFilePart As LongPtr _
) As Long
Private Declare PtrSafe Sub CopyMemory Lib "Kernel32.dll" Alias "RtlCopyMemory" ( _
    ByVal Destination As LongPtr, _
    ByVal Source As LongPtr, _
    ByVal Length As Long _
)
Private Declare PtrSafe Function FormatMessageW Lib "Kernel32.dll" ( _
    ByVal dwFlags As Long, _
    ByVal lpSource As LongPtr, _
    ByVal dwMessageId As Long, _
    ByVal dwLanguageId As Long, _
    ByVal lpBuffer As LongPtr, _
    ByVal nSize As Long, _
    ByVal Arguments As LongPtr _
) As Long
#Else
Private Const NULL_PTR As Long = 0&

Private Declare Function LoadLibraryExA Lib "Kernel32.dll" ( _
    ByVal lpLibFileName As String, _
    ByVal hFile As Long, _
    ByVal dwFlags As Long _
) As Long
Private Declare Function AddDllDirectory Lib "Kernel32.dll" ( _
    ByVal NewDirectory As Long _
) As Long
Private Declare Function SetDefaultDllDirectories Lib "Kernel32.dll" ( _
    ByVal DirectoryFlags As Long _
) As Byte
Private Declare Function GetFullPathNameA Lib "Kernel32.dll" ( _
    ByVal lpFileName As String, _
    ByVal nBufferLength As Long, _
    ByVal lpBuffer As String, _
    ByVal lpFilePart As Long _
) As Long
Private Declare Sub CopyMemory Lib "Kernel32.dll" Alias "RtlMoveMemory" ( _
    ByVal Destination As Long, _
    ByVal Source As Long, _
    ByVal Length As Long _
)
Private Declare Function FormatMessageW Lib "Kernel32.dll" ( _
    ByVal dwFlags As Long, _
    ByVal lpSource As Long, _
    ByVal dwMessageId As Long, _
    ByVal dwLanguageId As Long, _
    ByVal lpBuffer As Long, _
    ByVal nSize As Long, _
    ByVal Arguments As Long _
) As Long
#End If

''
'A setting for this module that specifies whether the module will throw error messages or
'print error messages to the immediate window.
'   True  -  Print error messages to the immediate window
'   False -  Raise and error event.
Dim QuiteErrors As Boolean

''
'Description:
'   Gets the Error string associated with a System Error Message
'Parameters:
'   e As Long  -  The error code
'Returns:
'   String  -  The system generated message associated with the supplied error code.
'Remarks:
'   Intended to provide the user as much information as possible should any of the
'   functions defined below fail.
Private Function GetDllErrorMessage(e As Long) As String
    Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100&
    Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000&
    Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200&
    
    Const LANG_USER_DEFAULT = &H400
#If Win64 Or VBA7 Then
    Dim ptr As LongPtr
#Else
    Dim ptr As Long
#End If
    Dim buff As String
    Dim rs As Long
    
    rs = FormatMessageW( _
        FORMAT_MESSAGE_ALLOCATE_BUFFER Or FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, _
        NULL_PTR, _
        e, _
        LANG_USER_DEFAULT, _
        VarPtr(ptr), _
        0, _
        NULL_PTR _
    )
    
    If rs = 0 Then
        GetDllErrorMessage = "Application-defined or object-defined error"
    Else
        buff = VBA.String(rs, VBA.Chr(0))
        
        CopyMemory StrPtr(buff), ptr, rs * 2
        
        GetDllErrorMessage = buff
    End If
End Function

''
'Description:
'   Turns a relative path into a full system path. It is not able to check the existence of
'   a directory or file and does not process wild cards.
'Parameters:
'   Path As String  -  A relative path to evaluate into a full system path.
'Returns:
'   String  -  The full system path as a null terminated string.
'Remarks:
'   Please note that due to the nature of SetCurrentDirectory relative paths may return
'   unexpected results. For example regardless of nest an Excel file is in a onedrive it
'   will return the base onedrive directory 'C:\Users\<user-name>\One-Drive' or
'   'C:\Users\<user-name>\One-Drive\Documents'
Function GetFullPath(Path As String)
    Dim fp As String * MAX_PATH
    Dim r As Long
    Dim e As Long
    
    r = GetFullPathNameA(Path, MAX_PATH, fp, NULL_PTR)
    
    If r = 0 Then
        e = Err.LastDllError
        
        If QuiteErrors Then
            Debug.Print "Win32 Error: 0x" & VBA.Right("00000000" & VBA.Hex$(e), 8)
            Debug.Print GetDllErrorMessage(e)
        Else
            Err.Raise _
                e, _
                "LoadLibrary.GetFullPath", _
                GetDllErrorMessage(e), _
                "https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-erref/18d8fbe8-a967-4f1c-ae50-99ca8e491d2d"
        End If
    End If
    
    GetFullPath = VBA.Left$(fp, r)
End Function

''
'Description:
'   Load a specified library into the VBE's active memory. This allows the user to hook into the functions
'   using standard VBE declare statements.
'Parameters:
'   Lib as String          -  The library to load.
'   Load_Type as LoadType  -  How the LoadLibrary function should process the Lib parameter.
'                               RelativePath  -  This option tells the LoadLibrary function that it should
'                                                evaluate the Lib parameter into a System Path.
'                               FullPath      -  This option tells the LoadLibrary function that it should
'                                                treat the Lib parameter as a complete System Path.
'                               LibName       -  This options tells the LoadLibrary function that a library
'                                                alias was supplied by the Lib parameter. This will cause the
'                                                function to act the same as the VBE's standard library loader.
'Returns:
'   Boolean  - True on Success
'              False on Failure (If QuiteErrors is False then an error is raised instead)
'Remarks:
'   The libraries loaded by the function are handed over to the VBE and will be unloaded by the VBE.
'   It is important to note that all of the dependencies specified in the library loaded will also be
'   loaded.
Public Function LoadLibrary(Lib As String, Optional Load_Type As LoadType = FullPath) As Boolean
    LoadLibrary = False
    
    Dim hModule As LongPtr
    
    If LoadType = RelativePath Then
        Lib = GetFullPath(Lib)
        
        LoadType = FullPath
    End If
    
    hModule = LoadLibraryExA(Lib, NULL_PTR, LoadType)
    
    If hModule = NULL_PTR Then
        e = Err.LastDllError
        
        If QuiteErrors Then
            Debug.Print "Win32 Error: 0x" & VBA.Right("00000000" & VBA.Hex$(e), 8)
            Debug.Print GetDllErrorMessage(e)
        Else
            Err.Raise _
                e, _
                "LoadLibrary.GetFullPath", _
                GetDllErrorMessage(e), _
                "https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-erref/18d8fbe8-a967-4f1c-ae50-99ca8e491d2d"
        End If
    Else
        LoadLibrary = True
    End If
End Function

''
'Description:
'   This function allows the user to add a directory to the default search directories searched by both
'   this module's LoadLibrary function and the VBE's default library loader.
'   For example supplying the directory 'U:\Dlls' will allow both this module's LoadLibrary function and
'   the VBE's default library loader to load and library in that directory by it's library alias.
'Paramaters:
'   Dir As String
'   Load_Type As LoadType  -  How the LoadLibrary function should process the Lib parameter.
'                               RelativePath  -  This option tells the LoadLibrary function that it should
'                                                evaluate the Lib parameter into a System Path.
'                               FullPath      -  This option tells the LoadLibrary function that it should
'                                                treat the Lib parameter as a complete System Path.
'                               LibName       -  Is NOT valid for this function and will be treated as if
'                                                RelativePath was supplied.
'Returns:
'   Boolean  -  True on Success
'               False on Failure (If QuiteErrors is False then an error is raised instead)
'Remarks:
'   This function is be used when large code libraries (however unlikely) are being referenced.
Public Function AddLibraryDirectory(Dir As String, Optional Load_Type As LoadType = FullPath) As Boolean
    AddLibraryDirectory = False
    
    If Load_Type <> FullPath Then
        Dir = GetFullPath(Dir)
    End If
    
    Dim cookie As LongPtr
    Dim r As Byte
    
    cookie = AddDllDirectory(StrPtr(Dir))
    
    If cookie = NULL_PTR Then
        e = Err.LastDllError
        
        If QuiteErrors Then
            Debug.Print "Win32 Error: 0x" & VBA.Right("00000000" & VBA.Hex$(e), 8)
            Debug.Print GetDllErrorMessage(e)
        Else
            Err.Raise _
                e, _
                "LoadLibrary.GetFullPath", _
                GetDllErrorMessage(e), _
                "https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-erref/18d8fbe8-a967-4f1c-ae50-99ca8e491d2d"
        End If
        
        Exit Function
    End If
    
    r = SetDefaultDllDirectories(LOAD_LIBRARY_SEARCH_DEFAULT_DIRS)
    
    If r = CFALSE Then
        e = Err.LastDllError
        
        If QuiteErrors Then
            Debug.Print "Win32 Error: 0x" & VBA.Right("00000000" & VBA.Hex$(e), 8)
            Debug.Print GetDllErrorMessage(e)
        Else
            Err.Raise _
                e, _
                "LoadLibrary.GetFullPath", _
                GetDllErrorMessage(e), _
                "https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-erref/18d8fbe8-a967-4f1c-ae50-99ca8e491d2d"
        End If
    Else
        AddLibraryDirectory = True
    End If
End Function
