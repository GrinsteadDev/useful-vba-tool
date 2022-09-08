VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JSON"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'JSON Object Class                                                          '
'Description:                                                               '
'   A JSON Class Module designed to turn valid JSON files and string data   '
'   and turn it into a nested JSON object                                   '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Type JSONKey
    Name As String
    Index As Long
End Type
Private Type JSONValue
    vType As JSONType
    strValue As String
    Value As JSON
End Type
Private RegEx As Object
Private objType As JSONType
Private keyCollection() As JSONKey
Private valCollection() As JSONValue
Private defValue As String
''
'Description:
'   The enum representing the JSON Object Type.
Public Enum JSONType
    UNDEFINED = -1
    JSONArray = 0
    JSONObj = 1
    JSONString = 2
    JSONNumber = 3
End Enum
''
'Description:
'   Takes a valid JSON string and turns it into a JSON Object
'Params
'   [in] string Str  -  The valid JSON String
Public Sub parseString(str As String)
    Dim mCol As Object
    
    Set mCol = createTokens(str)
    parseTokens mCol
End Sub
''
'Description:
'   Takes a valid file path and reads the contents into a JSON Object
'Params
'   [in] string File  -  The file path to a JSON file
Public Sub parseFile(File As String)
    Dim data As String
    Dim txtLn As String
    Dim fNum As Integer
    
    data = String(FileLen(File), " ")
    fNum = FreeFile
    
    Open File For Binary Access Read As #fNum
    Get #fNum, , data
    Close #fNum
    
    parseString data
End Sub
''
'Description:
'   The default property of THIS class module.
'Returns:
'   The JSON Object's default value when the JSON Object type is JSONString or JSONNumber
'   else it raises an "Object doesn't support this property or method" error
Public Property Get Value() As String
Attribute Value.VB_UserMemId = 0
    If objType = JSONNumber Or objType = JSONString Then
        Value = defValue
    ElseIf objType = UNDEFINED Then
        Value = ""
    Else
        Err.Raise 438
    End If
End Property
''
'Description:
'   The main way to access value key pairs in a JSON Object
'Params:
'   [in]  string Key       -  The key to search in the JSON Object
'Returns:
'   [out] JSON   getValue  -  The value assoicated with a provided key. This value
'                             will always be a JSON Object with one of the following types:
'                                 JSONArray
'                                 JSONObj
'                                 JSONString
'                                 JSONNumber
Public Function getValue(key As String) As JSON
    Dim k As JSONKey
    Dim v As JSONValue
    Dim j As JSON
    
    k = findKey(key)
    
    If k.Index <> -1 Then
        v = valCollection(k.Index)
        If v.Value Is Nothing Then
            If v.vType = JSONArray Or v.vType = JSONObj Then
                Set j = New JSON
                
                j.parseString v.strValue
                Set v.Value = j
            Else
                Set j = New JSON
                
                j.setType v.vType
                j.setValue v.strValue
                Set v.Value = j
            End If
            valCollection(k.Index) = v
        End If
        Set getValue = v.Value
    Else
        Set getValue = Nothing
    End If
End Function
''
'Description:
'   Sets a Value with a given Key. If a Key already exists the existing value will be overwritten with the
'   supplied Value. If JSON Object is of types JSONString or JSONNumber it will set the default value
'   regardless of the Key supplied.
'Params:
'            [in] Variant Value  -  The Value to set. The Values supplied must one of the following types:
'                                       vbLong
'                                       vbInteger
'                                       vbString
'                                       vbArray
'                                       vbObject (JSON)
'   optional [in] string Key     -  The Key to associate with the supplied Value. This paramater only applies
'                                   if the JSON Object is of types JSONObject or JSONArray. If the JSON Object
'                                   is of type JSONObject a Key must be supplied. An empty string "" is treated
'                                   as a valid JSON Key. If the JSON Object is of type JSONArray then a key is
'                                   not required. Non-numeric keys will be ignored. Any key provided that is not
'                                   already in the Index will be ignored. Any ignored key will have it's value
'                                   appened to the end of the array with a Key equal to +1 the current highest
'                                   index
Public Function setValue(Value As Variant, Optional key As String)
    Dim vT As VbVarType: vT = VarType(Value)
    
    If objType = JSONNumber Or objType = JSONString Then
        If vT = vbLong Or vT = vbInteger Then
            objType = JSONNumber
            defValue = Value
        ElseIf vT = vbString Then
            objType = JSONString
            If VBA.Left$(Value, 1) = """" Then Value = VBA.Right$(Value, VBA.Len(Value) - 1)
            If VBA.Right$(Value, 1) = """" Then Value = VBA.Left$(Value, VBA.Len(Value) - 1)
            defValue = Value
        Else
            Err.Raise 13
        End If
        
        Exit Function
    End If
    
    Dim k As JSONKey
    Dim v As JSONValue
    Dim j As JSON
    
    k = findKey(VBA.Trim$(key))
    
    If k.Index = -1 Then
        If isKeysAllocated Then
            ReDim Preserve keyCollection(0 To UBound(keyCollection) + 1)
        Else
            ReDim keyCollection(0 To 0)
        End If
        If isValsAllocated Then
            ReDim Preserve valCollection(0 To UBound(valCollection) + 1)
        Else
            ReDim valCollection(0 To 0)
        End If
        
        keyCollection(UBound(keyCollection)).Index = UBound(valCollection)
        If objType = JSONArray Then
            keyCollection(UBound(keyCollection)).Name = UBound(valCollection)
        Else
            keyCollection(UBound(keyCollection)).Name = VBA.Trim$(key)
        End If
        
        k.Index = keyCollection(UBound(keyCollection)).Index
        k.Name = keyCollection(UBound(keyCollection)).Name
    End If
    
    v = valCollection(k.Index)
    Set j = New JSON
    
    If vT = vbString Then
        v.vType = JSONString
        v.strValue = """" & Value & """"
        
        j.setType JSONString
        j.setValue v.strValue
        
        Set v.Value = j
    ElseIf vT = vbLong Or vT = vbInteger Then
        v.vType = JSONNumber
        v.strValue = CStr(Value)
        
        j.setType JSONNumber
        j.setValue v.strValue
        
        Set v.Value = j
    ElseIf vT = vbArray Then
        j.parseString "[" & VBA.Join(Value, ",") & "]"
        
        v.vType = JSONArray
        v.strValue = "[" & VBA.Join(Value, ",") & "]"
        Set v.Value = j
    ElseIf vT = vbObject Then
        If TypeOf Value Is JSON Then
            v.vType = Value.getType
            v.strValue = Value.toString
            Set v.Value = Value
        Else
            Err.Raise 13
            Exit Function
        End If
    Else
        Err.Raise 13
        Exit Function
    End If
    
    valCollection(k.Index) = v
End Function
''
'Description:
'   Removes a Value Key pair from the JSON Object
'Params:
'   [in] String Key  -  The Key of the Value Key pair to remove
Public Sub Remove(key As String)
    If objType <> JSONArray And objType <> JSONObj Then Exit Sub
    Dim k As JSONKey
    
    k = findKey(VBA.Trim$(key))
    
    If k.Index = -1 Then Exit Sub
    
    removeKey VBA.Trim$(key)
    removeValue k.Index
End Sub
''
'Description:
'   The function checks if a Key exists
'Params:
'   [in]  string  Key     -  The Key to check
'Returns:
'   [out] Boolean hasKey  -  True if the Key is found, otherwise False
Public Function hasKey(key As String) As Boolean
    Dim k As JSONKey
    
    k = findKey(key)
    hasKey = k.Index <> -1
End Function
''
'Description:
'   This function returns the current type of the current JSON Object
'Returns:
'   [out] JSONType getType  -  The JSONType of the current JSON Object
'                                  UNDEFINED
'                                  JSONArray
'                                  JSONObj
'                                  JSONString
'                                  JSONNumber
Public Function getType() As JSONType
    getType = objType
End Function
''
'Description:
'   This function sets the current type of the current JSON Object
'   but only if the current Object has not keys assigned to it.
'Params:
'   [in] JSONType jType  -  The JSONType to set
'                               UNDEFINED
'                               JSONArray
'                               JSONObj
'                               JSONString
'                               JSONNumber
Public Sub setType(jType As JSONType)
    If Not isKeysAllocated And Not isValsAllocated Then
        objType = jType
    Else
        
    End If
End Sub
''
'Description:
'   The Key collection of the JSON Object
'Params:
'   optional [in]  int     Index  -  If supplied the Index will be used to retrive
'                                    a Key from the collection at the specified
'                                    index.
'Returns:
'            [out] Variant Keys   -  If an Index is supplied the return value is a String
'                                    or number representing the requested key. If an Index
'                                    is not supplied an array mirroring the keyCollection
'                                    is returned
Public Function Keys(Optional Index) As Variant
    If IsMissing(Index) Then
        Dim out() As String
        Dim i
        
        ReDim out(0 To UBound(keyCollection))
        For i = 0 To UBound(out)
            out(i) = keyCollection(i).Name
        Next i
        Keys = out
        Exit Function
    End If
    If VarType(Index) <> vbLong And VarType(Index) <> vbInteger Then
        Err.Raise 13
        Exit Function
    End If
    
    Keys = keyCollection(Index).Name
End Function
''
'Description:
'   Creates a Valid JSON String the represents the current JSON Object
'Params:
'   optional [in]  Boolean Pretty    -  This Value determines if the returned JSON String
'                                       is formated a single line or human readable
'                                       Examples:
'                                           (True)  -  {"key_1":"value 1","key_2":{"key_2.1":"value2.1"}}
'                                           (False) -  {
'                                                          "key_1":"value 1",
'                                                          "key_2":{
'                                                              "key_2.1":"value2.1"
'                                                          }
'                                                       }
'   optional [in]  String  Indent    -  This value is the Indent or white space used in Pretty Mode.
'                                       It defaults to vbTab
'   optional [in]  Long    lvl       -  This is the number of times of the indent is repeated
'Returns:
'            [out] String  toString  -  A Valid JSON String
Public Function toString(Optional Pretty As Boolean = False, Optional Indent As String = vbTab, Optional lvl As Long = 0) As String
    If objType = JSONNumber Then
        toString = CStr(defValue)
        Exit Function
    ElseIf objType = JSONString Then
        toString = """" & defValue & """"
        Exit Function
    End If
    
    Dim tIndent As String
    Dim t2Indent As String
    Dim tLine As String
    
    If Pretty Then
        tIndent = repeat(Indent, lvl)
        t2Indent = repeat(Indent, lvl - 1)
        tLine = vbNewLine
    End If
    
    If Not isKeysAllocated Then
        If objType = JSONArray Then
            toString = "[" & tLine & t2Indent & "]"
        End If
        If objType = JSONObj Then
            toString = "{" & tLine & t2Indent & "}"
        End If
        Exit Function
    End If
    If UBound(keyCollection) = 0 Then
        If objType = JSONArray Then
            toString = "[" & tLine & tIndent & getValue(keyCollection(0).Name).toString(Pretty, Indent, lvl + 1) & tLine & t2Indent & "]"
        Else
            toString = "{" & tLine & tIndent & """" & keyCollection(0).Name & """:" & getValue(keyCollection(0).Name).toString(Pretty, Indent, lvl + 1) & tLine & t2Indent & "}"
        End If
        Exit Function
    End If
    
    Dim out() As String
    Dim i As Long
    
    ReDim out(0 To UBound(keyCollection))
    
    If objType = JSONArray Then
        out(0) = "[" & tLine & tIndent & getValue(keyCollection(0).Name).toString(Pretty, Indent, lvl + 1)
    Else
        out(0) = "{" & tLine & tIndent & """" & keyCollection(0).Name & """:" & getValue(keyCollection(0).Name).toString(Pretty, Indent, lvl + 1)
    End If
    
    For i = 1 To UBound(out) - 1
        If objType = JSONArray Then
            out(i) = "," & tLine & tIndent & getValue(keyCollection(i).Name).toString(Pretty, Indent, lvl + 1)
        Else
            out(i) = "," & tLine & tIndent & """" & keyCollection(i).Name & """:" & getValue(keyCollection(i).Name).toString(Pretty, Indent, lvl + 1)
        End If
    Next i
    
    If objType = JSONArray Then
        out(UBound(out)) = "," & tLine & tIndent & getValue(keyCollection(UBound(out)).Name).toString(Pretty, Indent, lvl + 1) & tLine & t2Indent & "]"
    Else
        out(UBound(out)) = "," & tLine & tIndent & """" & keyCollection(UBound(out)).Name & """:" & getValue(keyCollection(UBound(out)).Name).toString(Pretty, Indent, lvl + 1) & tLine & t2Indent & "}"
    End If
    
    toString = VBA.Join(out, "")
End Function
''
'Description:
'   The number of Keys containted in the current JSON Object
'Returns:
'   [out] Long Count  -  The number of Keys containted in the current JSON Object
Public Function Count() As Long
    If isValsAllocated Then
        Count = UBound(valCollection) + 1
    Else
        Count = 0
    End If
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'INTERNAL/HELPER FUNCTIONS                              '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function createTokens(str As String) As Object
    If RegEx Is Nothing Then
        Set RegEx = CreateObject("VBScript.RegExp")
    End If
    
    RegEx.Pattern = """(([^""\\]|\\.)*)""|[+\-]?(?:""|[1-9]\d*)(?:\.\d*)?(?:[eE][+\-]?\d+)?|\w+|[^\s""']+?"
    RegEx.IgnoreCase = True
    RegEx.Global = True
    RegEx.MultiLine = False
    
    Set createTokens = RegEx.Execute(str)
End Function

Private Sub parseTokens(mCol As Object)
    Dim i As Long
    Dim cLen As Long
    Dim m As Object
    Dim mV As String
    Dim mS As String
    Dim lvl As Long
    Dim cType As JSONType
    Dim keyDelimiter As Boolean
    Dim valDelimiter As Boolean
    Dim vArr() As String
    
    cLen = mCol.Count - 1
    keyDelimiter = False
    valDelimiter = True
    For i = 0 To cLen
        Set m = mCol(i)
        mV = m.Value
        mS = m.SubMatches(0)
        
        If mV = "{" Or mV = "[" Then
            lvl = lvl + 1
        End If
        If mV = "}" Or mV = "]" Then
            lvl = lvl - 1
        End If
        
        If lvl = 1 Then
            
            If mV = "{" Then
                objType = JSONObj
                GoTo NextToken
            End If
            If mV = "[" Then
                objType = JSONArray
                GoTo NextToken
            End If
            If mV = "}" Or mV = "]" Then GoTo JoinString
            
            If objType = JSONObj Then
                If mV = ":" Then
                    keyDelimiter = True
                    GoTo NextToken
                End If
                If mV = "," Then
                    valDelimiter = True
                    keyDelimiter = False
                    GoTo NextToken
                End If
                
                If keyDelimiter = False And valDelimiter = True Then
                    If isKeysAllocated Then
                        ReDim Preserve keyCollection(0 To UBound(keyCollection) + 1)
                    Else
                        ReDim keyCollection(0 To 0)
                    End If
                    
                    If mS <> "" Then
                        keyCollection(UBound(keyCollection)).Name = mS
                    Else
                        keyCollection(UBound(keyCollection)).Name = mV
                    End If
                    
                    valDelimiter = False
                    GoTo NextToken
                End If
                If keyDelimiter Then
                    If isValsAllocated Then
                        ReDim Preserve valCollection(0 To UBound(valCollection) + 1)
                    Else
                        ReDim valCollection(0 To 0)
                    End If
                    
                    If mS <> "" Then
                        valCollection(UBound(valCollection)).vType = JSONString
                        valCollection(UBound(valCollection)).strValue = mV
                    Else
                        valCollection(UBound(valCollection)).vType = JSONNumber
                        valCollection(UBound(valCollection)).strValue = mV
                    End If
                    keyCollection(UBound(keyCollection)).Index = UBound(valCollection)
                    
                    keyDelimiter = False
                    GoTo NextToken
                End If
            End If
            
            If objType = JSONArray Then
                If mV = "," Then
                    valDelimiter = True
                    GoTo NextToken
                End If
                
                If valDelimiter Then
                    If isKeysAllocated Then
                        ReDim Preserve keyCollection(0 To UBound(keyCollection) + 1)
                    Else
                        ReDim keyCollection(0 To 0)
                    End If
                    If isValsAllocated Then
                        ReDim Preserve valCollection(0 To UBound(valCollection) + 1)
                    Else
                        ReDim valCollection(0 To 0)
                    End If
                    
                    If mS <> "" Then
                        valCollection(UBound(valCollection)).vType = JSONString
                        valCollection(UBound(valCollection)).strValue = mV
                    Else
                        valCollection(UBound(valCollection)).vType = JSONNumber
                        valCollection(UBound(valCollection)).strValue = mV
                    End If
                    keyCollection(UBound(keyCollection)).Index = UBound(valCollection)
                    keyCollection(UBound(keyCollection)).Name = CStr(UBound(valCollection))
                    
                    valDelimiter = False
                    GoTo NextToken
                End If
            End If
            
            GoTo NextToken
        End If
        
        If lvl > 1 Then
            
            If lvl = 2 Then
                If (mV = "{" Or mV = "[") And objType = JSONObj Then
                    If isValsAllocated Then
                        ReDim Preserve valCollection(0 To UBound(valCollection) + 1)
                    Else
                        ReDim valCollection(0 To 0)
                    End If
                    
                    keyCollection(UBound(keyCollection)).Index = UBound(valCollection)
                    valCollection(UBound(valCollection)).vType = JSONObj
                    ReDim vArr(0 To 0)
                    vArr(0) = mV
                    
                    GoTo NextToken
                End If
                If (mV = "{" Or mV = "[") And objType = JSONArray Then
                    If isKeysAllocated Then
                        ReDim Preserve keyCollection(0 To UBound(keyCollection) + 1)
                    Else
                        ReDim keyCollection(0 To 0)
                    End If
                    If isValsAllocated Then
                        ReDim Preserve valCollection(0 To UBound(valCollection) + 1)
                    Else
                        ReDim valCollection(0 To 0)
                    End If
                    
                    keyCollection(UBound(keyCollection)).Index = UBound(valCollection)
                    keyCollection(UBound(keyCollection)).Name = CStr(UBound(valCollection))
                    valCollection(UBound(valCollection)).vType = JSONArray
                    ReDim vArr(0 To 0)
                    vArr(0) = mV
                    
                    GoTo NextToken
                End If
            End If
            
            ReDim Preserve vArr(0 To UBound(vArr) + 1)
            vArr(UBound(vArr)) = mV
            
            GoTo NextToken
        End If
        
        GoTo NextToken
JoinString:
        If mV = "}" Or mV = "]" Then
            ReDim Preserve vArr(0 To UBound(vArr) + 1)
            vArr(UBound(vArr)) = mV
            valCollection(UBound(valCollection)).strValue = VBA.Join(vArr, "")
            
            GoTo NextToken
        End If
        
NextToken:
    Next i
End Sub

Private Function isKeysAllocated() As Boolean
    isKeysAllocated = (Not Not keyCollection) <> 0
End Function
Private Function isValsAllocated() As Boolean
    isValsAllocated = (Not Not valCollection) <> 0
End Function

Private Function findKey(key As String) As JSONKey
    Dim i As Long
    
    If isKeysAllocated Then
        For i = 0 To UBound(keyCollection)
            If VBA.Trim$(key) = VBA.Trim$(keyCollection(i).Name) Then
                findKey.Index = keyCollection(i).Index
                findKey.Name = keyCollection(i).Name
                Exit Function
            End If
        Next i
    End If
    
    findKey.Index = -1
End Function

Private Function repeat(Value As String, Number As Long) As String
    If Number < 0 Then Exit Function
    
    Dim i
    Dim out() As String
    
    ReDim out(0 To Number)
    For i = 0 To Number
        out(i) = Value
    Next i
    
    repeat = VBA.Join(out, "")
End Function

Private Sub removeKey(key As String)
    If Not isKeysAllocated Then Exit Sub
    Dim i As Long
    Dim idx As Long
    
    For i = 0 To UBound(keyCollection)
        If VBA.Trim$(key) = VBA.Trim$(keyCollection(i).Name) Then
            idx = i
            Exit For
        End If
    Next i
    
    If idx > UBound(keyCollection) And UBound(keyCollection) > 0 Then Exit Sub
    If UBound(keyCollection) = 0 Then
        Erase keyCollection
        Exit Sub
    End If
    
    For i = idx To UBound(keyCollection) - 1
        keyCollection(i + 1).Index = i
        If objType = JSONArray Then keyCollection(i + 1).Name = CStr(i)
        keyCollection(i) = keyCollection(i + 1)
    Next i
    
    
    ReDim Preserve keyCollection(0 To UBound(keyCollection) - 1)
End Sub

Private Sub removeValue(idx As Long)
    If Not isValsAllocated Then Exit Sub
    Dim i As Long
    
    If idx > UBound(valCollection) And UBound(valCollection) > 0 Then Exit Sub
    If UBound(valCollection) = 0 Then
        Erase valCollection
        Exit Sub
    End If
    
    For i = idx To UBound(valCollection) - 1
        valCollection(i) = valCollection(i + 1)
    Next i
    
    ReDim Preserve valCollection(0 To UBound(valCollection) - 1)
End Sub

Private Sub Class_Initialize()
    objType = UNDEFINED
End Sub
