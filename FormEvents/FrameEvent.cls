VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FrameEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''
' Allows VBA Developers to Register Progammatically Created FormFrame to Events
' Examples
'   1:
'     The follow example assumes that this event object is being created inside of a Form
'     code block, thus the Self Object Reference Me is passed.
'
'       Sub AddEventListener()
'           '' Allows the event object to persist.
'           Static frmEvent As New FrameEvent
'
'           '' Sets Label Object Target for events
'           Set frmEvent.FormFrame = Me.Frame1
'           '' Adds a Function Event Subscriber
'           frmEvent.AddEvent "My_Click", FrameEventTypes.Click, Me
'           frmEvent.AddEvent "My_DblClick", FrameEventTypes.DblClick, Me
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
'           Static frmEvent As New FrameEvent
'
'           '' Sets Label Object Target for events
'           Set frmEvent.FormFrame = UserForm1.Frame1
'           '' Adds a Function Event Subscriber
'           frmEvent.AddEvent "My_Click", FrameEventTypes.Click
'           frmEvent.AddEvent "My_DblClick", FrameEventTypes.DblClick
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

Public Enum FrameEventTypes
    AddControl
    BeforeDragOver
    BeforeDropOrPaste
    Click
    DblClick
    ErrorEvent
    KeyDown
    KeyPress
    KeyUp
    Layout
    MouseDown
    MouseMove
    MouseUp
    RemoveControl
    Scroll
    Zoom
End Enum

'' Public Event Object Property
Public WithEvents FormFrame As MSForms.Frame
Attribute FormFrame.VB_VarHelpID = -1

'' Public Function Calls
Private EventCollection As Collection

''
' AddEvent
'   This method allows events to added. Function Names MUST be unique even if an event is stored
'   in two different objects, having the same name will cause a collision.
' Parameters
'   FuncName  As String         - The name of the function to call. This name MUST be globally unique.
'   EventType As FrameEventTypes - An enum representing the event you wish to subscript to.
'   Obj       As Object         - An optional object instance.
Public Sub AddEvent(ByVal FuncName As String, ByVal EventType As FrameEventTypes, Optional Obj As Object = Nothing)
    EventCollection(EventType).Add Array(FuncName, Obj), FuncName
End Sub
''
' RemoveEvent
'   This method allows events to be removed.
' Parameters
'   FuncName  As String         - The name of the function to remove.
'   EventType As FrameEventTypes - An enum representing the event the function was subscripted to.
Public Sub RemoveEvent(ByVal FuncName As String, ByVal EventType As FrameEventTypes)
    EventCollection(EventType).Remove FuncName
End Sub

'' Initilize Class
Private Sub Class_Initialize()
    Set EventCollection = New Collection
    
    EventCollection.Add New Collection, CStr(FrameEventTypes.AddControl)
    EventCollection.Add New Collection, CStr(FrameEventTypes.BeforeDragOver)
    EventCollection.Add New Collection, CStr(FrameEventTypes.BeforeDropOrPaste)
    EventCollection.Add New Collection, CStr(FrameEventTypes.Click)
    EventCollection.Add New Collection, CStr(FrameEventTypes.DblClick)
    EventCollection.Add New Collection, CStr(FrameEventTypes.ErrorEvent)
    EventCollection.Add New Collection, CStr(FrameEventTypes.KeyDown)
    EventCollection.Add New Collection, CStr(FrameEventTypes.KeyPress)
    EventCollection.Add New Collection, CStr(FrameEventTypes.KeyUp)
    EventCollection.Add New Collection, CStr(FrameEventTypes.Layout)
    EventCollection.Add New Collection, CStr(FrameEventTypes.MouseDown)
    EventCollection.Add New Collection, CStr(FrameEventTypes.MouseMove)
    EventCollection.Add New Collection, CStr(FrameEventTypes.MouseUp)
    EventCollection.Add New Collection, CStr(FrameEventTypes.RemoveControl)
    EventCollection.Add New Collection, CStr(FrameEventTypes.Scroll)
    EventCollection.Add New Collection, CStr(FrameEventTypes.Zoom)
End Sub

''******************************Form Events************************************''

Private Sub FormFrame_AddControl(ByVal Control As MSForms.Control)
    Dim tmp As Variant

    For Each tmp In EventCollection(FrameEventTypes.AddControl)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), Control
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, Control
        End If
    Next
End Sub

Private Sub FormFrame_BeforeDragOver( _
    ByVal Cancel As MSForms.ReturnBoolean, ByVal Control As MSForms.Control, _
    ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, _
    ByVal State As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, _
    ByVal Shift As Integer _
)
    Dim tmp As Variant

    For Each tmp In EventCollection(FrameEventTypes.BeforeDragOver)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), Cancel, Control, Data, X, Y, State, Effect, Shift
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, Cancel, Control, Data, X, Y, State, Effect, Shift
        End If
    Next
End Sub

Private Sub FormFrame_BeforeDropOrPaste( _
    ByVal Cancel As MSForms.ReturnBoolean, ByVal Control As MSForms.Control, _
    ByVal Action As MSForms.fmAction, ByVal Data As MSForms.DataObject, ByVal X As Single, _
    ByVal Y As Single, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer _
)
    Dim tmp As Variant

    For Each tmp In EventCollection(FrameEventTypes.BeforeDropOrPaste)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), Cancel, Control, Action, Data, X, Y, Effect, Shift
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, Cancel, Control, Action, Data, X, Y, Effect, Shift
        End If
    Next
End Sub

Private Sub FormFrame_Click()
    Dim tmp As Variant

    For Each tmp In EventCollection(FrameEventTypes.Click)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex)
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod
        End If
    Next
End Sub

Private Sub FormFrame_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim tmp As Variant

    For Each tmp In EventCollection(FrameEventTypes.DblClick)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), Cancel
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, Cancel
        End If
    Next
End Sub

Private Sub FormFrame_Error( _
    ByVal Number As Integer, ByVal Description As MSForms.ReturnString, _
    ByVal SCode As Long, ByVal Source As String, ByVal HelpFile As String, _
    ByVal HelpContext As Long, ByVal CancelDisplay As MSForms.ReturnBoolean _
)
    Dim tmp As Variant

    For Each tmp In EventCollection(FrameEventTypes.ErrorEvent)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), Number, Description, SCode, Source, HelpFile, HelpContext, CancelDisplay
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, Number, Description, SCode, Source, HelpFile, HelpContext, CancelDisplay
        End If
    Next
End Sub

Private Sub FormFrame_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim tmp As Variant

    For Each tmp In EventCollection(FrameEventTypes.KeyDown)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), KeyCode, Shift
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, KeyCode, Shift
        End If
    Next
End Sub

Private Sub FormFrame_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim tmp As Variant

    For Each tmp In EventCollection(FrameEventTypes.KeyPress)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), KeyAscii
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, KeyAscii
        End If
    Next
End Sub

Private Sub FormFrame_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim tmp As Variant

    For Each tmp In EventCollection(FrameEventTypes.KeyUp)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), KeyCode, Shift
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, KeyCode, Shift
        End If
    Next
End Sub

Private Sub FormFrame_Layout()
    Dim tmp As Variant

    For Each tmp In EventCollection(FrameEventTypes.Layout)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex)
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod
        End If
    Next
End Sub

Private Sub FormFrame_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim tmp As Variant

    For Each tmp In EventCollection(FrameEventTypes.MouseDown)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), Button, Shift, X, Y
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, Button, Shift, X, Y
        End If
    Next
End Sub

Private Sub FormFrame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim tmp As Variant

    For Each tmp In EventCollection(FrameEventTypes.MouseMove)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), Button, Shift, X, Y
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, Button, Shift, X, Y
        End If
    Next
End Sub

Private Sub FormFrame_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim tmp As Variant

    For Each tmp In EventCollection(FrameEventTypes.MouseUp)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), Button, Shift, X, Y
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, Button, Shift, X, Y
        End If
    Next
End Sub

Private Sub FormFrame_RemoveControl(ByVal Control As MSForms.Control)
    Dim tmp As Variant

    For Each tmp In EventCollection(FrameEventTypes.RemoveControl)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), Control
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, Control
        End If
    Next
End Sub

Private Sub FormFrame_Scroll( _
    ByVal ActionX As MSForms.fmScrollAction, ByVal ActionY As MSForms.fmScrollAction, _
    ByVal RequestDx As Single, ByVal RequestDy As Single, _
    ByVal ActualDx As MSForms.ReturnSingle, ByVal ActualDy As MSForms.ReturnSingle _
)
    Dim tmp As Variant

    For Each tmp In EventCollection(FrameEventTypes.Scroll)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), ActionX, ActionY, RequestDx, RequestDy, ActualDx, ActualDy
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, ActionX, ActionY, RequestDx, RequestDy, ActualDx, ActualDy
        End If
    Next
End Sub

Private Sub FormFrame_Zoom(Percent As Integer)
    Dim tmp As Variant

    For Each tmp In EventCollection(FrameEventTypes.Zoom)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), Percent
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, Percent
        End If
    Next
End Sub

''*****************************End FormFrame Events*********************************''


