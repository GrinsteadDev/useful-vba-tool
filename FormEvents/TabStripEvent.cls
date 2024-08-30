VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TabStripEvent"
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
'           Static tbStrEvent As New TabStripEvent
'
'           '' Sets Label Object Target for events
'           Set tbStrEvent.FormTabStrip = Me.TabStrip1
'           '' Adds a Function Event Subscriber
'           tbStrEvent.AddEvent "My_Click", TabStripEventTypes.Click, Me
'           tbStrEvent.AddEvent "My_DblClick", TabStripEventTypes.DblClick, Me
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
'           Static tbStrEvent As New TabStripEvent
'
'           '' Sets Label Object Target for events
'           Set tbStrEvent.FormTabStrip = UserForm1.TabStrip1
'           '' Adds a Function Event Subscriber
'           tbStrEvent.AddEvent "My_Click", TabStripEventTypes.Click
'           tbStrEvent.AddEvent "My_DblClick", TabStripEventTypes.DblClick
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

Public Enum TabStripEventTypes
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
Public WithEvents FormTabStrip As MSForms.TabStrip
Attribute FormTabStrip.VB_VarHelpID = -1

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
Public Sub AddEvent(ByVal FuncName As String, ByVal EventType As TabStripEventTypes, Optional Obj As Object = Nothing)
    EventCollection(EventType).Add Array(FuncName, Obj), FuncName
End Sub
''
' RemoveEvent
'   This method allows events to be removed.
' Parameters
'   FuncName  As String          - The name of the function to remove.
'   EventType As LabelEventTypes - An enum representing the event the function was subscripted to.
Public Sub RemoveEvent(ByVal FuncName As String, ByVal EventType As TabStripEventTypes)
    EventCollection(EventType).Remove FuncName
End Sub

'' Initilize Class
Private Sub Class_Initialize()
    Set EventCollection = New Collection
    
    EventCollection.Add New Collection, CStr(TabStripEventTypes.BeforeDragOver)
    EventCollection.Add New Collection, CStr(TabStripEventTypes.BeforeDropOrPaste)
    EventCollection.Add New Collection, CStr(TabStripEventTypes.Change)
    EventCollection.Add New Collection, CStr(TabStripEventTypes.Click)
    EventCollection.Add New Collection, CStr(TabStripEventTypes.DblClick)
    EventCollection.Add New Collection, CStr(TabStripEventTypes.ErrorEvent)
    EventCollection.Add New Collection, CStr(TabStripEventTypes.KeyDown)
    EventCollection.Add New Collection, CStr(TabStripEventTypes.KeyPress)
    EventCollection.Add New Collection, CStr(TabStripEventTypes.KeyUp)
    EventCollection.Add New Collection, CStr(TabStripEventTypes.MouseDown)
    EventCollection.Add New Collection, CStr(TabStripEventTypes.MouseMove)
    EventCollection.Add New Collection, CStr(TabStripEventTypes.MouseUp)
End Sub

''******************************TabStrip Events************************************''


Private Sub FormTabStrip_BeforeDragOver( _
    ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, _
    ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, _
    ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer _
)
    Dim tmp As Variant
    
    For Each tmp In EventCollection(TabStripEventTypes.BeforeDragOver)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), Cancel, Data, X, Y, DragState, Effect, Shift
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, Cancel, Data, X, Y, DragState, Effect, Shift
        End If
    Next
End Sub

Private Sub FormTabStrip_BeforeDropOrPaste( _
    ByVal Cancel As MSForms.ReturnBoolean, ByVal Action As MSForms.fmAction, _
    ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, _
    ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer _
)
    Dim tmp As Variant
    
    For Each tmp In EventCollection(TabStripEventTypes.BeforeDropOrPaste)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), Cancel, Action, Data, X, Y, Effect, Shift
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, Cancel, Action, Data, X, Y, Effect, Shift
        End If
    Next
End Sub

Private Sub FormTabStrip_Change()
    Dim tmp As Variant
    
    For Each tmp In EventCollection(TabStripEventTypes.Change)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex)
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod
        End If
    Next
End Sub

Private Sub FormTabStrip_Click()
    Dim tmp As Variant
    
    For Each tmp In EventCollection(TabStripEventTypes.Click)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex)
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod
        End If
    Next
End Sub

Private Sub FormTabStrip_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim tmp As Variant
    
    For Each tmp In EventCollection(TabStripEventTypes.DblClick)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), Cancel
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, Cancel
        End If
    Next
End Sub

Private Sub FormTabStrip_Error( _
    ByVal Number As Integer, ByVal Description As MSForms.ReturnString, _
    ByVal SCode As Long, ByVal Source As String, ByVal HelpFile As String, _
    ByVal HelpContext As Long, ByVal CancelDisplay As MSForms.ReturnBoolean _
)
    Dim tmp As Variant
    
    For Each tmp In EventCollection(TabStripEventTypes.ErrorEvent)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), Number, Description, SCode, Source, HelpFile, HelpContext, CancelDisplay
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, Number, Description, SCode, Source, HelpFile, HelpContext, CancelDisplay
        End If
    Next
End Sub

Private Sub FormTabStrip_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim tmp As Variant
    
    For Each tmp In EventCollection(TabStripEventTypes.KeyDown)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), KeyCode, Shift
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, KeyCode, Shift
        End If
    Next
End Sub

Private Sub FormTabStrip_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim tmp As Variant
    
    For Each tmp In EventCollection(TabStripEventTypes.KeyPress)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), KeyAscii
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, KeyAscii
        End If
    Next
End Sub

Private Sub FormTabStrip_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim tmp As Variant
    
    For Each tmp In EventCollection(TabStripEventTypes.KeyUp)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), KeyCode, Shift
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, KeyCode, Shift
        End If
    Next
End Sub

Private Sub FormTabStrip_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim tmp As Variant
    
    For Each tmp In EventCollection(TabStripEventTypes.MouseDown)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), Button, Shift, X, Y
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, Button, Shift, X, Y
        End If
    Next
End Sub

Private Sub FormTabStrip_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim tmp As Variant
    
    For Each tmp In EventCollection(TabStripEventTypes.MouseMove)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), Button, Shift, X, Y
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, Button, Shift, X, Y
        End If
    Next
End Sub

Private Sub FormTabStrip_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim tmp As Variant
    
    For Each tmp In EventCollection(TabStripEventTypes.MouseUp)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), Button, Shift, X, Y
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, Button, Shift, X, Y
        End If
    Next
End Sub

''*****************************End TabStrip Events*********************************''



