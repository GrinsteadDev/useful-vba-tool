VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListBoxEvent"
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
'           Static lsboxEvent As New ListBoxEvent
'
'           '' Sets Label Object Target for events
'           Set lsboxEvent.FormListBox = Me.ListBox1
'           '' Adds a Function Event Subscriber
'           lsboxEvent.AddEvent "My_Click", ListBoxEventTypes.Click, Me
'           lsboxEvent.AddEvent "My_DblClick", ListBoxEventTypes.DblClick, Me
'       End Sub
'
'       Sub My_Click(Sender As MSForms.ListBox)
'           MsgBox "I was clicked"
'       End Sub
'       Sub My_DblClick(Sender As MSForms.ListBox, ByVal Cancel As MSForms.ReturnBoolean)
'           MsgBox "I was Double Clicked"
'       End Sub
'
'   2:
'     The following example assumes that this event object is being create inside of a Module
'     code block, thus the third parameter is being ommited.
'
'       Sub AddEventListener()
'           '' Allows the event object to persist.
'           Static lsboxEvent As New ListBoxEvent
'
'           '' Sets Label Object Target for events
'           Set lsboxEvent.FormListBox = UserForm1.ListBox1
'           '' Adds a Function Event Subscriber
'           lsboxEvent.AddEvent "My_Click", ListBoxEventTypes.Click
'           lsboxEvent.AddEvent "My_DblClick", ListBoxEventTypes.DblClick
'       End Sub
'
'       Sub My_Click(Sender As MSForms.ListBox)
'           MsgBox "I was clicked"
'       End Sub
'       Sub My_DblClick(Sender As MSForms.ListBox, ByVal Cancel As MSForms.ReturnBoolean)
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

Public Enum ListBoxEventTypes
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
Public WithEvents FormListBox As MSForms.ListBox
Attribute FormListBox.VB_VarHelpID = -1

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
Public Sub AddEvent(ByVal FuncName As String, ByVal EventType As ListBoxEventTypes, Optional Obj As Object = Nothing)
    EventCollection(EventType).Add Array(FuncName, Obj), FuncName
End Sub
''
' RemoveEvent
'   This method allows events to be removed.
' Parameters
'   FuncName  As String          - The name of the function to remove.
'   EventType As LabelEventTypes - An enum representing the event the function was subscripted to.
Public Sub RemoveEvent(ByVal FuncName As String, ByVal EventType As ListBoxEventTypes)
    EventCollection(EventType).Remove FuncName
End Sub

'' Initilize Class
Private Sub Class_Initialize()
    Set EventCollection = New Collection
    
    EventCollection.Add New Collection, CStr(ListBoxEventTypes.BeforeDragOver)
    EventCollection.Add New Collection, CStr(ListBoxEventTypes.BeforeDropOrPaste)
    EventCollection.Add New Collection, CStr(ListBoxEventTypes.Change)
    EventCollection.Add New Collection, CStr(ListBoxEventTypes.Click)
    EventCollection.Add New Collection, CStr(ListBoxEventTypes.DblClick)
    EventCollection.Add New Collection, CStr(ListBoxEventTypes.ErrorEvent)
    EventCollection.Add New Collection, CStr(ListBoxEventTypes.KeyDown)
    EventCollection.Add New Collection, CStr(ListBoxEventTypes.KeyPress)
    EventCollection.Add New Collection, CStr(ListBoxEventTypes.KeyUp)
    EventCollection.Add New Collection, CStr(ListBoxEventTypes.MouseDown)
    EventCollection.Add New Collection, CStr(ListBoxEventTypes.MouseMove)
    EventCollection.Add New Collection, CStr(ListBoxEventTypes.MouseUp)
End Sub

''******************************ListBox Events************************************''

Private Sub FormListBox_BeforeDragOver( _
    ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, _
    ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, _
    ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer _
)
    Dim tmp As Variant
    
    For Each tmp In EventCollection(ListBoxEventTypes.BeforeDragOver)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), FormListBox, Cancel, Data, X, Y, DragState, Effect, Shift
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, FormListBox, Cancel, Data, X, Y, DragState, Effect, Shift
        End If
    Next
End Sub

Private Sub FormListBox_BeforeDropOrPaste( _
    ByVal Cancel As MSForms.ReturnBoolean, ByVal Action As MSForms.fmAction, _
    ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, _
    ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer _
)
    Dim tmp As Variant
    
    For Each tmp In EventCollection(ListBoxEventTypes.BeforeDropOrPaste)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), FormListBox, Cancel, Action, Data, X, Y, Effect, Shift
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, FormListBox, Cancel, Action, Data, X, Y, Effect, Shift
        End If
    Next
End Sub

Private Sub FormListBox_Change()
    Dim tmp As Variant
    
    For Each tmp In EventCollection(ListBoxEventTypes.Change)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), FormListBox
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, FormListBox
        End If
    Next
End Sub

Private Sub FormListBox_Click()
    Dim tmp As Variant
    
    For Each tmp In EventCollection(ListBoxEventTypes.Click)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), FormListBox
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, FormListBox
        End If
    Next
End Sub

Private Sub FormListBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim tmp As Variant
    
    For Each tmp In EventCollection(ListBoxEventTypes.DblClick)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), FormListBox, Cancel
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, FormListBox, Cancel
        End If
    Next
End Sub

Private Sub FormListBox_Error( _
    ByVal Number As Integer, ByVal Description As MSForms.ReturnString, _
    ByVal SCode As Long, ByVal Source As String, ByVal HelpFile As String, _
    ByVal HelpContext As Long, ByVal CancelDisplay As MSForms.ReturnBoolean _
)
    Dim tmp As Variant
    
    For Each tmp In EventCollection(ListBoxEventTypes.ErrorEvent)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), FormListBox, Number, Description, SCode, Source, HelpFile, HelpContext, CancelDisplay
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, FormListBox, Number, Description, SCode, Source, HelpFile, HelpContext, CancelDisplay
        End If
    Next
End Sub

Private Sub FormListBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim tmp As Variant
    
    For Each tmp In EventCollection(ListBoxEventTypes.KeyDown)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), FormListBox, KeyCode, Shift
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, FormListBox, KeyCode, Shift
        End If
    Next
End Sub

Private Sub FormListBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim tmp As Variant
    
    For Each tmp In EventCollection(ListBoxEventTypes.KeyPress)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), FormListBox, KeyAscii
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, FormListBox, KeyAscii
        End If
    Next
End Sub

Private Sub FormListBox_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim tmp As Variant
    
    For Each tmp In EventCollection(ListBoxEventTypes.KeyUp)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), FormListBox, KeyCode, Shift
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, FormListBox, KeyCode, Shift
        End If
    Next
End Sub

Private Sub FormListBox_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim tmp As Variant
    
    For Each tmp In EventCollection(ListBoxEventTypes.MouseDown)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), FormListBox, Button, Shift, X, Y
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, FormListBox, Button, Shift, X, Y
        End If
    Next
End Sub

Private Sub FormListBox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim tmp As Variant
    
    For Each tmp In EventCollection(ListBoxEventTypes.MouseMove)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), FormListBox, Button, Shift, X, Y
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, FormListBox, Button, Shift, X, Y
        End If
    Next
End Sub

Private Sub FormListBox_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim tmp As Variant
    
    For Each tmp In EventCollection(ListBoxEventTypes.MouseUp)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), FormListBox, Button, Shift, X, Y
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, FormListBox, Button, Shift, X, Y
        End If
    Next
End Sub

''*****************************End ListBox Events*********************************''





