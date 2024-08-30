VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MultiPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''
' Allows VBA Developers to Register Progammatically Created FormMultiPage to Events. Events will
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
'           Static mltPageEvent As New MultiPageEvent
'
'           '' Sets Label Object Target for events
'           Set mltPageEvent.FormMultiPage = Me.MultiPage1
'           '' Adds a Function Event Subscriber
'           mltPageEvent.AddEvent "My_Click", MultiPageEventTypes.Click, Me
'           mltPageEvent.AddEvent "My_DblClick", MultiPageEventTypes.DblClick, Me
'       End Sub
'
'       Sub My_Click(Sender As MSForms.MultiPage)
'           MsgBox "I was clicked"
'       End Sub
'       Sub My_DblClick(Sender As MSForms.MultiPage, ByVal Cancel As MSForms.ReturnBoolean)
'           MsgBox "I was Double Clicked"
'       End Sub
'
'   2:
'     The following example assumes that this event object is being create inside of a Module
'     code block, thus the third parameter is being ommited.
'
'       Sub AddEventListener()
'           '' Allows the event object to persist.
'           Static mltPageEvent As New MultiPageEvent
'
'           '' Sets Label Object Target for events
'           Set mltPageEvent.FormMultiPage = UserForm1.MultiPage1
'           '' Adds a Function Event Subscriber
'           mltPageEvent.AddEvent "My_Click", MultiPageEventTypes.Click
'           mltPageEvent.AddEvent "My_DblClick", MultiPageEventTypes.DblClick
'       End Sub
'
'       Sub My_Click(Sender As MSForms.MultiPage)
'           MsgBox "I was clicked"
'       End Sub
'       Sub My_DblClick(Sender As MSForms.MultiPage, ByVal Cancel As MSForms.ReturnBoolean)
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

Public Enum MultiPageEventTypes
    AddControl
    BeforeDragOver
    BeforeDropOrPaste
    Change
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
Public WithEvents FormMultiPage As MSForms.MultiPage
Attribute FormMultiPage.VB_VarHelpID = -1

'' Public Function Calls
Private EventCollection As Collection

''
' AddEvent
'   This method allows events to added. Function Names MUST be unique even if an event is stored
'   in two different objects, having the same name will cause a collision.
' Parameters
'   FuncName  As String         - The name of the function to call. This name MUST be globally unique.
'   EventType As MultiPageEventTypes - An enum representing the event you wish to subscript to.
'   Obj       As Object         - An optional object instance.
Public Sub AddEvent(ByVal FuncName As String, ByVal EventType As MultiPageEventTypes, Optional Obj As Object = Nothing)
    EventCollection(EventType).Add Array(FuncName, Obj), FuncName
End Sub
''
' RemoveEvent
'   This method allows events to be removed.
' Parameters
'   FuncName  As String         - The name of the function to remove.
'   EventType As MultiPageEventTypes - An enum representing the event the function was subscripted to.
Public Sub RemoveEvent(ByVal FuncName As String, ByVal EventType As MultiPageEventTypes)
    EventCollection(EventType).Remove FuncName
End Sub

'' Initilize Class
Private Sub Class_Initialize()
    Set EventCollection = New Collection
    
    EventCollection.Add New Collection, CStr(MultiPageEventTypes.AddControl)
    EventCollection.Add New Collection, CStr(MultiPageEventTypes.BeforeDragOver)
    EventCollection.Add New Collection, CStr(MultiPageEventTypes.BeforeDropOrPaste)
    EventCollection.Add New Collection, CStr(MultiPageEventTypes.Change)
    EventCollection.Add New Collection, CStr(MultiPageEventTypes.Click)
    EventCollection.Add New Collection, CStr(MultiPageEventTypes.DblClick)
    EventCollection.Add New Collection, CStr(MultiPageEventTypes.ErrorEvent)
    EventCollection.Add New Collection, CStr(MultiPageEventTypes.KeyDown)
    EventCollection.Add New Collection, CStr(MultiPageEventTypes.KeyPress)
    EventCollection.Add New Collection, CStr(MultiPageEventTypes.KeyUp)
    EventCollection.Add New Collection, CStr(MultiPageEventTypes.Layout)
    EventCollection.Add New Collection, CStr(MultiPageEventTypes.MouseDown)
    EventCollection.Add New Collection, CStr(MultiPageEventTypes.MouseMove)
    EventCollection.Add New Collection, CStr(MultiPageEventTypes.MouseUp)
    EventCollection.Add New Collection, CStr(MultiPageEventTypes.RemoveControl)
    EventCollection.Add New Collection, CStr(MultiPageEventTypes.Scroll)
    EventCollection.Add New Collection, CStr(MultiPageEventTypes.Zoom)
End Sub

''******************************MultiPage Events************************************''

Private Sub FormMultiPage_AddControl(ByVal Control As MSForms.Control)
    Dim tmp As Variant

    For Each tmp In EventCollection(MultiPageEventTypes.AddControl)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), FormMultiPage, Control
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, FormMultiPage, Control
        End If
    Next
End Sub

Private Sub FormMultiPage_BeforeDragOver( _
    ByVal Cancel As MSForms.ReturnBoolean, ByVal Control As MSForms.Control, _
    ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, _
    ByVal State As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, _
    ByVal Shift As Integer _
)
    Dim tmp As Variant

    For Each tmp In EventCollection(MultiPageEventTypes.BeforeDragOver)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), FormMultiPage, Cancel, Control, Data, X, Y, State, Effect, Shift
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, FormMultiPage, Cancel, Control, Data, X, Y, State, Effect, Shift
        End If
    Next
End Sub

Private Sub FormMultiPage_BeforeDropOrPaste( _
    ByVal Cancel As MSForms.ReturnBoolean, ByVal Control As MSForms.Control, _
    ByVal Action As MSForms.fmAction, ByVal Data As MSForms.DataObject, ByVal X As Single, _
    ByVal Y As Single, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer _
)
    Dim tmp As Variant

    For Each tmp In EventCollection(MultiPageEventTypes.BeforeDropOrPaste)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), FormMultiPage, Cancel, Control, Action, Data, X, Y, Effect, Shift
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, FormMultiPage, Cancel, Control, Action, Data, X, Y, Effect, Shift
        End If
    Next
End Sub

Private Sub FormMultiPage_Change()
    Dim tmp As Variant

    For Each tmp In EventCollection(MultiPageEventTypes.Change)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), FormMultiPage
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, FormMultiPage
        End If
    Next
End Sub

Private Sub FormMultiPage_Click()
    Dim tmp As Variant

    For Each tmp In EventCollection(MultiPageEventTypes.Click)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), FormMultiPage
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, FormMultiPage
        End If
    Next
End Sub

Private Sub FormMultiPage_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim tmp As Variant

    For Each tmp In EventCollection(MultiPageEventTypes.DblClick)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), FormMultiPage, Cancel
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, FormMultiPage, Cancel
        End If
    Next
End Sub

Private Sub FormMultiPage_Error( _
    ByVal Number As Integer, ByVal Description As MSForms.ReturnString, _
    ByVal SCode As Long, ByVal Source As String, ByVal HelpFile As String, _
    ByVal HelpContext As Long, ByVal CancelDisplay As MSForms.ReturnBoolean _
)
    Dim tmp As Variant

    For Each tmp In EventCollection(MultiPageEventTypes.ErrorEvent)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), FormMultiPage, Number, Description, SCode, Source, HelpFile, HelpContext, CancelDisplay
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, FormMultiPage, Number, Description, SCode, Source, HelpFile, HelpContext, CancelDisplay
        End If
    Next
End Sub

Private Sub FormMultiPage_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim tmp As Variant

    For Each tmp In EventCollection(MultiPageEventTypes.KeyDown)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), FormMultiPage, KeyCode, Shift
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, FormMultiPage, KeyCode, Shift
        End If
    Next
End Sub

Private Sub FormMultiPage_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim tmp As Variant

    For Each tmp In EventCollection(MultiPageEventTypes.KeyPress)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), FormMultiPage, KeyAscii
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, FormMultiPage, KeyAscii
        End If
    Next
End Sub

Private Sub FormMultiPage_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim tmp As Variant

    For Each tmp In EventCollection(MultiPageEventTypes.KeyUp)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), FormMultiPage, KeyCode, Shift
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, FormMultiPage, KeyCode, Shift
        End If
    Next
End Sub

Private Sub FormMultiPage_Layout()
    Dim tmp As Variant

    For Each tmp In EventCollection(MultiPageEventTypes.Layout)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), FormMultiPage
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, FormMultiPage
        End If
    Next
End Sub

Private Sub FormMultiPage_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim tmp As Variant

    For Each tmp In EventCollection(MultiPageEventTypes.MouseDown)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), FormMultiPage, Button, Shift, X, Y
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, FormMultiPage, Button, Shift, X, Y
        End If
    Next
End Sub

Private Sub FormMultiPage_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim tmp As Variant

    For Each tmp In EventCollection(MultiPageEventTypes.MouseMove)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), FormMultiPage, Button, Shift, X, Y
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, FormMultiPage, Button, Shift, X, Y
        End If
    Next
End Sub

Private Sub FormMultiPage_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim tmp As Variant

    For Each tmp In EventCollection(MultiPageEventTypes.MouseUp)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), FormMultiPage, Button, Shift, X, Y
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, FormMultiPage, Button, Shift, X, Y
        End If
    Next
End Sub

Private Sub FormMultiPage_RemoveControl(ByVal Control As MSForms.Control)
    Dim tmp As Variant

    For Each tmp In EventCollection(MultiPageEventTypes.RemoveControl)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), FormMultiPage, Control
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, FormMultiPage, Control
        End If
    Next
End Sub

Private Sub FormMultiPage_Scroll( _
    ByVal ActionX As MSForms.fmScrollAction, ByVal ActionY As MSForms.fmScrollAction, _
    ByVal RequestDx As Single, ByVal RequestDy As Single, _
    ByVal ActualDx As MSForms.ReturnSingle, ByVal ActualDy As MSForms.ReturnSingle _
)
    Dim tmp As Variant

    For Each tmp In EventCollection(MultiPageEventTypes.Scroll)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), FormMultiPage, ActionX, ActionY, RequestDx, RequestDy, ActualDx, ActualDy
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, FormMultiPage, ActionX, ActionY, RequestDx, RequestDy, ActualDx, ActualDy
        End If
    Next
End Sub

Private Sub FormMultiPage_Zoom(Percent As Integer)
    Dim tmp As Variant

    For Each tmp In EventCollection(MultiPageEventTypes.Zoom)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), FormMultiPage, Percent
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, FormMultiPage, Percent
        End If
    Next
End Sub

''*****************************End MultiPage Events*********************************''


