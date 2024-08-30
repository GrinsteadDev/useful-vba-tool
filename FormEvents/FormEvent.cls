VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''
' Allows VBA Developers to Register Progammatically Created Form to Events. Events will
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
'           Static frmEvent As New FormEvent
'
'           '' Sets Label Object Target for events
'           Set frmEvent.Form = Me
'           '' Adds a Function Event Subscriber
'           frmEvent.AddEvent "My_Click", FormEventTypes.Click, Me
'           frmEvent.AddEvent "My_DblClick", FormEventTypes.DblClick, Me
'       End Sub
'
'       Sub My_Click(Sender As MSForms.UserForm)
'           MsgBox "I was clicked"
'       End Sub
'       Sub My_DblClick(Sender As MSForms.UserForm, ByVal Cancel As MSForms.ReturnBoolean)
'           MsgBox "I was Double Clicked"
'       End Sub
'
'   2:
'     The following example assumes that this event object is being create inside of a Module
'     code block, thus the third parameter is being ommited.
'
'       Sub AddEventListener()
'           '' Allows the event object to persist.
'           Static frmEvent As New FormEvent
'
'           '' Sets Label Object Target for events
'           Set frmEvent.Form = MyForm
'           '' Adds a Function Event Subscriber
'           frmEvent.AddEvent "My_Click", FormEventTypes.Click
'           frmEvent.AddEvent "My_DblClick", FormEventTypes.DblClick
'       End Sub
'
'       Sub My_Click(Sender As MSForms.UserForm)
'           MsgBox "I was clicked"
'       End Sub
'       Sub My_DblClick(Sender As MSForms.UserForm, ByVal Cancel As MSForms.ReturnBoolean)
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

Public Enum FormEventTypes
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
Public WithEvents Form As MSForms.UserForm
Attribute Form.VB_VarHelpID = -1

'' Public Function Calls
Private EventCollection As Collection

''
' AddEvent
'   This method allows events to added. Function Names MUST be unique even if an event is stored
'   in two different objects, having the same name will cause a collision.
' Parameters
'   FuncName  As String         - The name of the function to call. This name MUST be globally unique.
'   EventType As FormEventTypes - An enum representing the event you wish to subscript to.
'   Obj       As Object         - An optional object instance.
Public Sub AddEvent(ByVal FuncName As String, ByVal EventType As FormEventTypes, Optional Obj As Object = Nothing)
    EventCollection(EventType).Add Array(FuncName, Obj), FuncName
End Sub
''
' RemoveEvent
'   This method allows events to be removed.
' Parameters
'   FuncName  As String         - The name of the function to remove.
'   EventType As FormEventTypes - An enum representing the event the function was subscripted to.
Public Sub RemoveEvent(ByVal FuncName As String, ByVal EventType As FormEventTypes)
    EventCollection(EventType).Remove FuncName
End Sub

'' Initilize Class
Private Sub Class_Initialize()
    Set EventCollection = New Collection
    
    EventCollection.Add New Collection, CStr(FormEventTypes.AddControl)
    EventCollection.Add New Collection, CStr(FormEventTypes.BeforeDragOver)
    EventCollection.Add New Collection, CStr(FormEventTypes.BeforeDropOrPaste)
    EventCollection.Add New Collection, CStr(FormEventTypes.Click)
    EventCollection.Add New Collection, CStr(FormEventTypes.DblClick)
    EventCollection.Add New Collection, CStr(FormEventTypes.ErrorEvent)
    EventCollection.Add New Collection, CStr(FormEventTypes.KeyDown)
    EventCollection.Add New Collection, CStr(FormEventTypes.KeyPress)
    EventCollection.Add New Collection, CStr(FormEventTypes.KeyUp)
    EventCollection.Add New Collection, CStr(FormEventTypes.Layout)
    EventCollection.Add New Collection, CStr(FormEventTypes.MouseDown)
    EventCollection.Add New Collection, CStr(FormEventTypes.MouseMove)
    EventCollection.Add New Collection, CStr(FormEventTypes.MouseUp)
    EventCollection.Add New Collection, CStr(FormEventTypes.RemoveControl)
    EventCollection.Add New Collection, CStr(FormEventTypes.Scroll)
    EventCollection.Add New Collection, CStr(FormEventTypes.Zoom)
End Sub

''******************************Form Events************************************''

Private Sub Form_AddControl(ByVal Control As MSForms.Control)
    Dim tmp As Variant

    For Each tmp In EventCollection(FormEventTypes.AddControl)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), Form, Control
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, Form, Control
        End If
    Next
End Sub

Private Sub Form_BeforeDragOver( _
    ByVal Cancel As MSForms.ReturnBoolean, ByVal Control As MSForms.Control, _
    ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, _
    ByVal State As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, _
    ByVal Shift As Integer _
)
    Dim tmp As Variant

    For Each tmp In EventCollection(FormEventTypes.BeforeDragOver)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), Form, Cancel, Control, Data, X, Y, State, Effect, Shift
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, Form, Cancel, Control, Data, X, Y, State, Effect, Shift
        End If
    Next
End Sub

Private Sub Form_BeforeDropOrPaste( _
    ByVal Cancel As MSForms.ReturnBoolean, ByVal Control As MSForms.Control, _
    ByVal Action As MSForms.fmAction, ByVal Data As MSForms.DataObject, ByVal X As Single, _
    ByVal Y As Single, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer _
)
    Dim tmp As Variant

    For Each tmp In EventCollection(FormEventTypes.BeforeDropOrPaste)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), Form, Cancel, Control, Action, Data, X, Y, Effect, Shift
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, Form, Cancel, Control, Action, Data, X, Y, Effect, Shift
        End If
    Next
End Sub

Private Sub Form_Click()
    Dim tmp As Variant

    For Each tmp In EventCollection(FormEventTypes.Click)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), Form
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, Form
        End If
    Next
End Sub

Private Sub Form_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim tmp As Variant

    For Each tmp In EventCollection(FormEventTypes.DblClick)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), Form, Cancel
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, Form, Cancel
        End If
    Next
End Sub

Private Sub Form_Error( _
    ByVal Number As Integer, ByVal Description As MSForms.ReturnString, _
    ByVal SCode As Long, ByVal Source As String, ByVal HelpFile As String, _
    ByVal HelpContext As Long, ByVal CancelDisplay As MSForms.ReturnBoolean _
)
    Dim tmp As Variant

    For Each tmp In EventCollection(FormEventTypes.ErrorEvent)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), Form, Number, Description, SCode, Source, HelpFile, HelpContext, CancelDisplay
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, Form, Number, Description, SCode, Source, HelpFile, HelpContext, CancelDisplay
        End If
    Next
End Sub

Private Sub Form_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim tmp As Variant

    For Each tmp In EventCollection(FormEventTypes.KeyDown)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), Form, KeyCode, Shift
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, Form, KeyCode, Shift
        End If
    Next
End Sub

Private Sub Form_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim tmp As Variant

    For Each tmp In EventCollection(FormEventTypes.KeyPress)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), Form, KeyAscii
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, Form, KeyAscii
        End If
    Next
End Sub

Private Sub Form_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim tmp As Variant

    For Each tmp In EventCollection(FormEventTypes.KeyUp)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), Form, KeyCode, Shift
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, Form, KeyCode, Shift
        End If
    Next
End Sub

Private Sub Form_Layout()
    Dim tmp As Variant

    For Each tmp In EventCollection(FormEventTypes.Layout)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), Form
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, Form
        End If
    Next
End Sub

Private Sub Form_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim tmp As Variant

    For Each tmp In EventCollection(FormEventTypes.MouseDown)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), Form, Button, Shift, X, Y
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, Form, Button, Shift, X, Y
        End If
    Next
End Sub

Private Sub Form_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim tmp As Variant

    For Each tmp In EventCollection(FormEventTypes.MouseMove)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), Form, Button, Shift, X, Y
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, Form, Button, Shift, X, Y
        End If
    Next
End Sub

Private Sub Form_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim tmp As Variant

    For Each tmp In EventCollection(FormEventTypes.MouseUp)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), Form, Button, Shift, X, Y
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, Form, Button, Shift, X, Y
        End If
    Next
End Sub

Private Sub Form_RemoveControl(ByVal Control As MSForms.Control)
    Dim tmp As Variant

    For Each tmp In EventCollection(FormEventTypes.RemoveControl)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), Form, Control
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, Form, Control
        End If
    Next
End Sub

Private Sub Form_Scroll( _
    ByVal ActionX As MSForms.fmScrollAction, ByVal ActionY As MSForms.fmScrollAction, _
    ByVal RequestDx As Single, ByVal RequestDy As Single, _
    ByVal ActualDx As MSForms.ReturnSingle, ByVal ActualDy As MSForms.ReturnSingle _
)
    Dim tmp As Variant

    For Each tmp In EventCollection(FormEventTypes.Scroll)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), Form, ActionX, ActionY, RequestDx, RequestDy, ActualDx, ActualDy
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, Form, ActionX, ActionY, RequestDx, RequestDy, ActualDx, ActualDy
        End If
    Next
End Sub

Private Sub Form_Zoom(Percent As Integer)
    Dim tmp As Variant

    For Each tmp In EventCollection(FormEventTypes.Zoom)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), Form, Percent
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, Form, Percent
        End If
    Next
End Sub

''*****************************End Form Events**********************************''
