VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ComboBoxEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''
' Allows VBA Developers to Register Progammatically Created Form Objects to Events. Events will
' have a signature identical to VBA's Event Signature with a Preappened Sender Paramater.
' Such As:
'   "Sub Control_Click()" is now "Sub Control_Click(Sender As MSForms.Control)"
'
'   "Sub Control_DblClick(ByVal Cancel As MSForms.ReturnBoolean)" is now
'   "Sub Control_DblClick(Sender As MSForms.Control, ByVal Cancel As MSForms.ReturnBoolean)"
' Examples
'   1:
'     The follow example assumes that this event object is being created inside of a Form
'     code block, thus the Self Object Reference Me is passed.
'
'       Sub AddEventListener()
'           '' Allows the event object to persist.
'           Static cbboxEvent As New ComboBoxEvent
'
'           '' Sets Label Object Target for events
'           Set cbboxEvent.FormComboBox = Me.ComboBox1
'           '' Adds a Function Event Subscriber
'           cbboxEvent.AddEvent "My_Click", ComboBoxEventTypes.Click, Me
'           cbboxEvent.AddEvent "My_DblClick", ComboBoxEventTypes.DblClick, Me
'       End Sub
'
'       Sub My_Click(Sender As MSForms.ComboBox)
'           MsgBox "I was clicked"
'       End Sub
'       Sub My_DblClick(Sender As MSForms.ComboBox, ByVal Cancel As MSForms.ReturnBoolean)
'           MsgBox "I was Double Clicked"
'       End Sub
'
'   2:
'     The following example assumes that this event object is being create inside of a Module
'     code block, thus the third parameter is being ommited.
'
'       Sub AddEventListener()
'           '' Allows the event object to persist.
'           Static cbboxEvent As New ComboBoxEvent
'
'           '' Sets Label Object Target for events
'           Set cbboxEvent.FormComboBox = UserForm1.ComboBox1
'           '' Adds a Function Event Subscriber
'           cbboxEvent.AddEvent "My_Click", ComboBoxEventTypes.Click
'           cbboxEvent.AddEvent "My_DblClick", ComboBoxEventTypes.DblClick
'       End Sub
'
'       Sub My_Click(Sender As MSForms.ComboBox)
'           MsgBox "I was clicked"
'       End Sub
'       Sub My_DblClick(Sender As MSForms.ComboBox, ByVal Cancel As MSForms.ReturnBoolean)
'           MsgBox "I was Double Clicked"
'       End Sub
'
' Dependencies
'   Microsoft Froms 2.0 Object Library
'   Microsoft Excel 16.0 Object Library
'   Visual Basic For Applications VB7

Option Explicit

Private Enum EventObj
    NameIndex
    ObjectIndex
End Enum

Public Enum ComboBoxEventTypes
    BeforeDragOver
    BeforeDropOrPaste
    Change
    Click
    DblClick
    DropButtonClick
    ErrorEvent
    KeyDown
    KeyPress
    KeyUp
    MouseDown
    MouseMove
    MouseUp
End Enum

'' Public Event Object Property
Public WithEvents FormComboBox As MSForms.ComboBox
Attribute FormComboBox.VB_VarHelpID = -1

'' Public Function Calls
Private EventCollection As Collection

''
' AddEvent
'   This method allows events to added. Function Names MUST be unique even if an event is stored
'   in two different objects, having the same name will cause a collision.
' Parameters
'   FuncName  As String          - The name of the function to call. This name MUST be globally unique.
'   EventType As LabelEventTypes - An enum representing the event you wish to subscript to.
'   Obj       As Object          - An optional object instance.
Public Sub AddEvent(ByVal FuncName As String, ByVal EventType As ComboBoxEventTypes, Optional Obj As Object = Nothing)
    EventCollection(EventType).Add Array(FuncName, Obj), FuncName
End Sub
''
' RemoveEvent
'   This method allows events to be removed.
' Parameters
'   FuncName  As String          - The name of the function to remove.
'   EventType As LabelEventTypes - An enum representing the event the function was subscripted to.
Public Sub RemoveEvent(ByVal FuncName As String, ByVal EventType As ComboBoxEventTypes)
    EventCollection(EventType).Remove FuncName
End Sub

'' Initilize Class
Private Sub Class_Initialize()
    Set EventCollection = New Collection
    
    EventCollection.Add New Collection, CStr(ComboBoxEventTypes.BeforeDragOver)
    EventCollection.Add New Collection, CStr(ComboBoxEventTypes.BeforeDropOrPaste)
    EventCollection.Add New Collection, CStr(ComboBoxEventTypes.Change)
    EventCollection.Add New Collection, CStr(ComboBoxEventTypes.Click)
    EventCollection.Add New Collection, CStr(ComboBoxEventTypes.DblClick)
    EventCollection.Add New Collection, CStr(ComboBoxEventTypes.DropButtonClick)
    EventCollection.Add New Collection, CStr(ComboBoxEventTypes.ErrorEvent)
    EventCollection.Add New Collection, CStr(ComboBoxEventTypes.KeyDown)
    EventCollection.Add New Collection, CStr(ComboBoxEventTypes.KeyPress)
    EventCollection.Add New Collection, CStr(ComboBoxEventTypes.KeyUp)
    EventCollection.Add New Collection, CStr(ComboBoxEventTypes.MouseDown)
    EventCollection.Add New Collection, CStr(ComboBoxEventTypes.MouseMove)
    EventCollection.Add New Collection, CStr(ComboBoxEventTypes.MouseUp)
End Sub

''******************************ComboBox Events************************************''


Private Sub FormComboBox_BeforeDragOver( _
    ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, _
    ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, _
    ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer _
)
    Dim tmp As Variant
    
    For Each tmp In EventCollection(ComboBoxEventTypes.BeforeDragOver)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), FormComboBox, Cancel, Data, X, Y, DragState, Effect, Shift
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, FormComboBox, Cancel, Data, X, Y, DragState, Effect, Shift
        End If
    Next
End Sub

Private Sub FormComboBox_BeforeDropOrPaste( _
    ByVal Cancel As MSForms.ReturnBoolean, ByVal Action As MSForms.fmAction, _
    ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, _
    ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer _
)
    Dim tmp As Variant
    
    For Each tmp In EventCollection(ComboBoxEventTypes.BeforeDropOrPaste)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), FormComboBox, Cancel, Action, Data, X, Y, Effect, Shift
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, FormComboBox, Cancel, Action, Data, X, Y, Effect, Shift
        End If
    Next
End Sub

Private Sub FormComboBox_Change()
    Dim tmp As Variant
    
    For Each tmp In EventCollection(ComboBoxEventTypes.Change)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), FormComboBox
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, FormComboBox
        End If
    Next
End Sub

Private Sub FormComboBox_Click()
    Dim tmp As Variant
    
    For Each tmp In EventCollection(ComboBoxEventTypes.Click)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), FormComboBox
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, FormComboBox
        End If
    Next
End Sub

Private Sub FormComboBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim tmp As Variant
    
    For Each tmp In EventCollection(ComboBoxEventTypes.DblClick)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), FormComboBox, Cancel
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, FormComboBox, Cancel
        End If
    Next
End Sub

Private Sub FormComboBox_DropButtonClick()
    Dim tmp As Variant
    
    For Each tmp In EventCollection(ComboBoxEventTypes.DropButtonClick)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), FormComboBox
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, FormComboBox
        End If
    Next
End Sub

Private Sub FormComboBox_Error( _
    ByVal Number As Integer, ByVal Description As MSForms.ReturnString, _
    ByVal SCode As Long, ByVal Source As String, ByVal HelpFile As String, _
    ByVal HelpContext As Long, ByVal CancelDisplay As MSForms.ReturnBoolean _
)
    Dim tmp As Variant
    
    For Each tmp In EventCollection(ComboBoxEventTypes.ErrorEvent)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), FormComboBox, Number, Description, SCode, Source, HelpFile, HelpContext, CancelDisplay
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, FormComboBox, Number, Description, SCode, Source, HelpFile, HelpContext, CancelDisplay
        End If
    Next
End Sub

Private Sub FormComboBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim tmp As Variant
    
    For Each tmp In EventCollection(ComboBoxEventTypes.KeyDown)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), FormComboBox, KeyCode, Shift
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, FormComboBox, KeyCode, Shift
        End If
    Next
End Sub

Private Sub FormComboBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim tmp As Variant
    
    For Each tmp In EventCollection(ComboBoxEventTypes.KeyPress)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), FormComboBox, KeyAscii
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, FormComboBox, KeyAscii
        End If
    Next
End Sub

Private Sub FormComboBox_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim tmp As Variant
    
    For Each tmp In EventCollection(ComboBoxEventTypes.KeyUp)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), FormComboBox, KeyCode, Shift
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, FormComboBox, KeyCode, Shift
        End If
    Next
End Sub

Private Sub FormComboBox_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim tmp As Variant
    
    For Each tmp In EventCollection(ComboBoxEventTypes.MouseDown)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), FormComboBox, Button, Shift, X, Y
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, FormComboBox, Button, Shift, X, Y
        End If
    Next
End Sub

Private Sub FormComboBox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim tmp As Variant
    
    For Each tmp In EventCollection(ComboBoxEventTypes.MouseMove)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), FormComboBox, Button, Shift, X, Y
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, FormComboBox, Button, Shift, X, Y
        End If
    Next
End Sub

Private Sub FormComboBox_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim tmp As Variant
    
    For Each tmp In EventCollection(ComboBoxEventTypes.MouseUp)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), FormComboBox, Button, Shift, X, Y
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, FormComboBox, Button, Shift, X, Y
        End If
    Next
End Sub

''*****************************End ComboBox Events*********************************''



