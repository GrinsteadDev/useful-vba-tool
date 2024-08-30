VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CommandButtonEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''
' Allows VBA Developers to Register Progammatically Created Form Objects to Events
' Examples
'   1:
'     The follow example assumes that this event object is being created inside of a Form
'     code block, thus the Self Object Reference Me is passed.
'
'       Sub AddEventListener()
'           '' Allows the event object to persist.
'           Static cbboxEvent As New CommandButtonEvent
'
'           '' Sets Label Object Target for events
'           Set cbboxEvent.FormCommandButton = Me.ComboBox1
'           '' Adds a Function Event Subscriber
'           cbboxEvent.AddEvent "My_Click", CommandButtonEventTypes.Click, Me
'           cbboxEvent.AddEvent "My_DblClick", CommandButtonEventTypes.DblClick, Me
'       End Sub
'
'       Sub My_Click()
'           MsgBox "I was clicked"
'       End Sub
'       Sub My_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
'           MsgBox "I was Double Clicked"
'       End Sub
'
'   2:
'     The following example assumes that this event object is being create inside of a Module
'     code block, thus the third parameter is being ommited.
'
'       Sub AddEventListener()
'           '' Allows the event object to persist.
'           Static cbboxEvent As New CommandButtonEvent
'
'           '' Sets Label Object Target for events
'           Set cbboxEvent.FormCommandButton = UserForm1.ComboBox1
'           '' Adds a Function Event Subscriber
'           cbboxEvent.AddEvent "My_Click", CommandButtonEventTypes.Click
'           cbboxEvent.AddEvent "My_DblClick", CommandButtonEventTypes.DblClick
'       End Sub
'
'       Sub My_Click()
'           MsgBox "I was clicked"
'       End Sub
'       Sub My_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
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

Public Enum CommandButtonEventTypes
    BeforeDragOver
    BeforeDropOrPaste
    Click
    DblClick
    ErrorEvent
    KeyDown
    KeyPress
    KeyUp
    MouseDown
    MouseMove
    MouseUp
End Enum

'' Public Event Object Property
Public WithEvents FormCommandButton As MSForms.CommandButton
Attribute FormCommandButton.VB_VarHelpID = -1

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
Public Sub AddEvent(ByVal FuncName As String, ByVal EventType As CommandButtonEventTypes, Optional Obj As Object = Nothing)
    EventCollection(EventType).Add Array(FuncName, Obj), FuncName
End Sub
''
' RemoveEvent
'   This method allows events to be removed.
' Parameters
'   FuncName  As String          - The name of the function to remove.
'   EventType As LabelEventTypes - An enum representing the event the function was subscripted to.
Public Sub RemoveEvent(ByVal FuncName As String, ByVal EventType As CommandButtonEventTypes)
    EventCollection(EventType).Remove FuncName
End Sub

'' Initilize Class
Private Sub Class_Initialize()
    Set EventCollection = New Collection
    
    EventCollection.Add New Collection, CStr(CommandButtonEventTypes.BeforeDragOver)
    EventCollection.Add New Collection, CStr(CommandButtonEventTypes.BeforeDropOrPaste)
    EventCollection.Add New Collection, CStr(CommandButtonEventTypes.Click)
    EventCollection.Add New Collection, CStr(CommandButtonEventTypes.DblClick)
    EventCollection.Add New Collection, CStr(CommandButtonEventTypes.ErrorEvent)
    EventCollection.Add New Collection, CStr(CommandButtonEventTypes.KeyDown)
    EventCollection.Add New Collection, CStr(CommandButtonEventTypes.KeyPress)
    EventCollection.Add New Collection, CStr(CommandButtonEventTypes.KeyUp)
    EventCollection.Add New Collection, CStr(CommandButtonEventTypes.MouseDown)
    EventCollection.Add New Collection, CStr(CommandButtonEventTypes.MouseMove)
    EventCollection.Add New Collection, CStr(CommandButtonEventTypes.MouseUp)
End Sub

''******************************CommandButton Events************************************''


Private Sub FormCommandButton_BeforeDragOver( _
    ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, _
    ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, _
    ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer _
)
    Dim tmp As Variant
    
    For Each tmp In EventCollection(CommandButtonEventTypes.BeforeDragOver)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), Cancel, Data, X, Y, DragState, Effect, Shift
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, Cancel, Data, X, Y, DragState, Effect, Shift
        End If
    Next
End Sub

Private Sub FormCommandButton_BeforeDropOrPaste( _
    ByVal Cancel As MSForms.ReturnBoolean, ByVal Action As MSForms.fmAction, _
    ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, _
    ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer _
)
    Dim tmp As Variant
    
    For Each tmp In EventCollection(CommandButtonEventTypes.BeforeDropOrPaste)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), Cancel, Action, Data, X, Y, Effect, Shift
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, Cancel, Action, Data, X, Y, Effect, Shift
        End If
    Next
End Sub

Private Sub FormCommandButton_Click()
    Dim tmp As Variant
    
    For Each tmp In EventCollection(CommandButtonEventTypes.Click)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex)
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod
        End If
    Next
End Sub

Private Sub FormCommandButton_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim tmp As Variant
    
    For Each tmp In EventCollection(CommandButtonEventTypes.DblClick)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), Cancel
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, Cancel
        End If
    Next
End Sub

Private Sub FormCommandButton_Error( _
    ByVal Number As Integer, ByVal Description As MSForms.ReturnString, _
    ByVal SCode As Long, ByVal Source As String, ByVal HelpFile As String, _
    ByVal HelpContext As Long, ByVal CancelDisplay As MSForms.ReturnBoolean _
)
    Dim tmp As Variant
    
    For Each tmp In EventCollection(CommandButtonEventTypes.ErrorEvent)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), Number, Description, SCode, Source, HelpFile, HelpContext, CancelDisplay
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, Number, Description, SCode, Source, HelpFile, HelpContext, CancelDisplay
        End If
    Next
End Sub

Private Sub FormCommandButton_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim tmp As Variant
    
    For Each tmp In EventCollection(CommandButtonEventTypes.KeyDown)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), KeyCode, Shift
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, KeyCode, Shift
        End If
    Next
End Sub

Private Sub FormCommandButton_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim tmp As Variant
    
    For Each tmp In EventCollection(CommandButtonEventTypes.KeyPress)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), KeyAscii
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, KeyAscii
        End If
    Next
End Sub

Private Sub FormCommandButton_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim tmp As Variant
    
    For Each tmp In EventCollection(CommandButtonEventTypes.KeyUp)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), KeyCode, Shift
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, KeyCode, Shift
        End If
    Next
End Sub

Private Sub FormCommandButton_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim tmp As Variant
    
    For Each tmp In EventCollection(CommandButtonEventTypes.MouseDown)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), Button, Shift, X, Y
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, Button, Shift, X, Y
        End If
    Next
End Sub

Private Sub FormCommandButton_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim tmp As Variant
    
    For Each tmp In EventCollection(CommandButtonEventTypes.MouseMove)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), Button, Shift, X, Y
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, Button, Shift, X, Y
        End If
    Next
End Sub

Private Sub FormCommandButton_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim tmp As Variant
    
    For Each tmp In EventCollection(CommandButtonEventTypes.MouseUp)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), Button, Shift, X, Y
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, Button, Shift, X, Y
        End If
    Next
End Sub

''*****************************End CommandButton Events*********************************''



