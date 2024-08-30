VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OptionButtonEvent"
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
'           Static opbtnEvent As New OptionButtonEvent
'
'           '' Sets Label Object Target for events
'           Set opbtnEvent.FormOptionButton = Me.OptionButton1
'           '' Adds a Function Event Subscriber
'           opbtnEvent.AddEvent "My_Click", OptionButtonEventTypes.Click, Me
'           opbtnEvent.AddEvent "My_DblClick", OptionButtonEventTypes.DblClick, Me
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
'           Static opbtnEvent As New OptionButtonEvent
'
'           '' Sets Label Object Target for events
'           Set opbtnEvent.FormOptionButton = UserForm1.OptionButton1
'           '' Adds a Function Event Subscriber
'           opbtnEvent.AddEvent "My_Click", OptionButtonEventTypes.Click
'           opbtnEvent.AddEvent "My_DblClick", OptionButtonEventTypes.DblClick
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

Public Enum OptionButtonEventTypes
    BeforeDragOver
    BeforeDropOrPaste
    Change
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
Public WithEvents FormOptionButton As MSForms.OptionButton
Attribute FormOptionButton.VB_VarHelpID = -1

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
Public Sub AddEvent(ByVal FuncName As String, ByVal EventType As OptionButtonEventTypes, Optional Obj As Object = Nothing)
    EventCollection(EventType).Add Array(FuncName, Obj), FuncName
End Sub
''
' RemoveEvent
'   This method allows events to be removed.
' Parameters
'   FuncName  As String          - The name of the function to remove.
'   EventType As LabelEventTypes - An enum representing the event the function was subscripted to.
Public Sub RemoveEvent(ByVal FuncName As String, ByVal EventType As OptionButtonEventTypes)
    EventCollection(EventType).Remove FuncName
End Sub

'' Initilize Class
Private Sub Class_Initialize()
    Set EventCollection = New Collection
    
    EventCollection.Add New Collection, CStr(OptionButtonEventTypes.BeforeDragOver)
    EventCollection.Add New Collection, CStr(OptionButtonEventTypes.BeforeDropOrPaste)
    EventCollection.Add New Collection, CStr(OptionButtonEventTypes.Change)
    EventCollection.Add New Collection, CStr(OptionButtonEventTypes.Click)
    EventCollection.Add New Collection, CStr(OptionButtonEventTypes.DblClick)
    EventCollection.Add New Collection, CStr(OptionButtonEventTypes.ErrorEvent)
    EventCollection.Add New Collection, CStr(OptionButtonEventTypes.KeyDown)
    EventCollection.Add New Collection, CStr(OptionButtonEventTypes.KeyPress)
    EventCollection.Add New Collection, CStr(OptionButtonEventTypes.KeyUp)
    EventCollection.Add New Collection, CStr(OptionButtonEventTypes.MouseDown)
    EventCollection.Add New Collection, CStr(OptionButtonEventTypes.MouseMove)
    EventCollection.Add New Collection, CStr(OptionButtonEventTypes.MouseUp)
End Sub

''******************************ListBox Events************************************''

Private Sub FormOptionButton_BeforeDragOver( _
    ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, _
    ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, _
    ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer _
)
    Dim tmp As Variant
    
    For Each tmp In EventCollection(OptionButtonEventTypes.BeforeDragOver)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), Cancel, Data, X, Y, DragState, Effect, Shift
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, Cancel, Data, X, Y, DragState, Effect, Shift
        End If
    Next
End Sub

Private Sub FormOptionButton_BeforeDropOrPaste( _
    ByVal Cancel As MSForms.ReturnBoolean, ByVal Action As MSForms.fmAction, _
    ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, _
    ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer _
)
    Dim tmp As Variant
    
    For Each tmp In EventCollection(OptionButtonEventTypes.BeforeDropOrPaste)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), Cancel, Action, Data, X, Y, Effect, Shift
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, Cancel, Action, Data, X, Y, Effect, Shift
        End If
    Next
End Sub

Private Sub FormOptionButton_Change()
    Dim tmp As Variant
    
    For Each tmp In EventCollection(OptionButtonEventTypes.Change)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex)
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod
        End If
    Next
End Sub

Private Sub FormOptionButton_Click()
    Dim tmp As Variant
    
    For Each tmp In EventCollection(OptionButtonEventTypes.Click)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex)
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod
        End If
    Next
End Sub

Private Sub FormOptionButton_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim tmp As Variant
    
    For Each tmp In EventCollection(OptionButtonEventTypes.DblClick)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), Cancel
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, Cancel
        End If
    Next
End Sub

Private Sub FormOptionButton_Error( _
    ByVal Number As Integer, ByVal Description As MSForms.ReturnString, _
    ByVal SCode As Long, ByVal Source As String, ByVal HelpFile As String, _
    ByVal HelpContext As Long, ByVal CancelDisplay As MSForms.ReturnBoolean _
)
    Dim tmp As Variant
    
    For Each tmp In EventCollection(OptionButtonEventTypes.ErrorEvent)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), Number, Description, SCode, Source, HelpFile, HelpContext, CancelDisplay
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, Number, Description, SCode, Source, HelpFile, HelpContext, CancelDisplay
        End If
    Next
End Sub

Private Sub FormOptionButton_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim tmp As Variant
    
    For Each tmp In EventCollection(OptionButtonEventTypes.KeyDown)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), KeyCode, Shift
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, KeyCode, Shift
        End If
    Next
End Sub

Private Sub FormOptionButton_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim tmp As Variant
    
    For Each tmp In EventCollection(OptionButtonEventTypes.KeyPress)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), KeyAscii
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, KeyAscii
        End If
    Next
End Sub

Private Sub FormOptionButton_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim tmp As Variant
    
    For Each tmp In EventCollection(OptionButtonEventTypes.KeyUp)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), KeyCode, Shift
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, KeyCode, Shift
        End If
    Next
End Sub

Private Sub FormOptionButton_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim tmp As Variant
    
    For Each tmp In EventCollection(OptionButtonEventTypes.MouseDown)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), Button, Shift, X, Y
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, Button, Shift, X, Y
        End If
    Next
End Sub

Private Sub FormOptionButton_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim tmp As Variant
    
    For Each tmp In EventCollection(OptionButtonEventTypes.MouseMove)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), Button, Shift, X, Y
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, Button, Shift, X, Y
        End If
    Next
End Sub

Private Sub FormOptionButton_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim tmp As Variant
    
    For Each tmp In EventCollection(OptionButtonEventTypes.MouseUp)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), Button, Shift, X, Y
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, Button, Shift, X, Y
        End If
    Next
End Sub

''*****************************End ListBox Events*********************************''


