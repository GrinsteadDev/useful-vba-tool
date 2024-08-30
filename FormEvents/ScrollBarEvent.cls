VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ScrollBarEvent"
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
'           Static scrollEvent As New ScrollBarEvent
'
'           '' Sets Label Object Target for events
'           Set scrollEvent.FormScrollBar = Me.ScrollBar1
'           '' Adds a Function Event Subscriber
'           scrollEvent.AddEvent "My_Change", ScrollBarEventTypes.Change, Me
'       End Sub
'
'       Sub My_Change()
'           MsgBox "I was Changed"
'       End Sub
'
'   2:
'     The following example assumes that this event object is being create inside of a Module
'     code block, thus the third parameter is being ommited.
'
'       Sub AddEventListener()
'           '' Allows the event object to persist.
'           Static scrollEvent As New ScrollBarEvent
'
'           '' Sets Label Object Target for events
'           Set scrollEvent.FormScrollBar = UserForm1.ScrollBar1
'           '' Adds a Function Event Subscriber
'           scrollEvent.AddEvent "My_Change", ScrollBarEventTypes.Change
'       End Sub
'
'       Sub My_Change()
'           MsgBox "I was Changed"
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

Public Enum ScrollBarEventTypes
    BeforeDragOver
    BeforeDropOrPast
    Change
    ErrorEvent
    KeyDown
    KeyPress
    KeyUp
    Scroll
End Enum

'' Public Event Object Property
Public WithEvents FormScrollBar As MSForms.ScrollBar
Attribute FormScrollBar.VB_VarHelpID = -1

'' Public Function Calls
Private EventCollection As Collection

''
' AddEvent
'   This method allows events to added. Function Names MUST be unique even if an event is stored
'   in two different objects, having the same name will cause a collision.
' Parameters
'   FuncName  As String          - The name of the function to call. This name MUST be globally unique.
'   EventType As ScrollBarEventTypes - An enum representing the event you wish to subscript to.
'   Obj       As Object          - An optional object instance.
Public Sub AddEvent(ByVal FuncName As String, ByVal EventType As ScrollBarEventTypes, Optional Obj As Object = Nothing)
    EventCollection(EventType).Add Array(FuncName, Obj), FuncName
End Sub
''
' RemoveEvent
'   This method allows events to be removed.
' Parameters
'   FuncName  As String          - The name of the function to remove.
'   EventType As ScrollBarEventTypes - An enum representing the event the function was subscripted to.
Public Sub RemoveEvent(ByVal FuncName As String, ByVal EventType As ScrollBarEventTypes)
    EventCollection(EventType).Remove FuncName
End Sub

'' Initilize Class
Private Sub Class_Initialize()
    Set EventCollection = New Collection
    
    EventCollection.Add New Collection, CStr(ScrollBarEventTypes.BeforeDragOver)
    EventCollection.Add New Collection, CStr(ScrollBarEventTypes.BeforeDropOrPast)
    EventCollection.Add New Collection, CStr(ScrollBarEventTypes.Change)
    EventCollection.Add New Collection, CStr(ScrollBarEventTypes.ErrorEvent)
    EventCollection.Add New Collection, CStr(ScrollBarEventTypes.KeyDown)
    EventCollection.Add New Collection, CStr(ScrollBarEventTypes.KeyPress)
    EventCollection.Add New Collection, CStr(ScrollBarEventTypes.KeyUp)
    EventCollection.Add New Collection, CStr(ScrollBarEventTypes.Scroll)
End Sub

''******************************ScrollBar Events************************************''

Private Sub FormScrollBar_BeforeDragOver( _
    ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, _
    ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, _
    ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer _
)
    Dim tmp As Variant
    
    For Each tmp In EventCollection(ScrollBarEventTypes.BeforeDragOver)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), Cancel, Date, X, Y, DragState, Effect, Shift
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, Cancel, Date, X, Y, DragState, Effect, Shift
        End If
    Next
End Sub

Private Sub FormScrollBar_BeforeDropOrPaste( _
    ByVal Cancel As MSForms.ReturnBoolean, ByVal Action As MSForms.fmAction, _
    ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, _
    ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer _
)
    Dim tmp As Variant
    
    For Each tmp In EventCollection(ScrollBarEventTypes.BeforeDropOrPast)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), Cancel, Action, Data, X, Y, Effect, Shift
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, Cancel, Action, Data, X, Y, Effect, Shift
        End If
    Next
End Sub

Private Sub FormScrollBar_Change()
    Dim tmp As Variant
    
    For Each tmp In EventCollection(ScrollBarEventTypes.Change)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex)
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod
        End If
    Next
End Sub

Private Sub FormScrollBar_Error( _
    ByVal Number As Integer, ByVal Description As MSForms.ReturnString, _
    ByVal SCode As Long, ByVal Source As String, _
    ByVal HelpFile As String, ByVal HelpContext As Long, _
    ByVal CancelDisplay As MSForms.ReturnBoolean _
)
    Dim tmp As Variant
    
    For Each tmp In EventCollection(ScrollBarEventTypes.ErrorEvent)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), Number, Description, SCode, Source, HelpFile, HelpContext, CancelDisplay
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, Number, Description, SCode, Source, HelpFile, HelpContext, CancelDisplay
        End If
    Next
End Sub

Private Sub FormScrollBar_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim tmp As Variant
    
    For Each tmp In EventCollection(ScrollBarEventTypes.KeyDown)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), KeyCode, Shift
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, KeyCode, Shift
        End If
    Next
End Sub

Private Sub FormScrollBar_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim tmp As Variant
    
    For Each tmp In EventCollection(ScrollBarEventTypes.KeyPress)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), KeyAscii
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, KeyAscii
        End If
    Next
End Sub

Private Sub FormScrollBar_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim tmp As Variant
    
    For Each tmp In EventCollection(ScrollBarEventTypes.KeyUp)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), KeyCode, Shift
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, KeyCode, Shift
        End If
    Next
End Sub

Private Sub FormScrollBar_Scroll( _
    ByVal ActionX As MSForms.fmScrollAction, ByVal ActionY As MSForms.fmScrollAction, _
    ByVal RequestDx As Single, ByVal RequestDy As Single, _
    ByVal ActualDx As MSForms.ReturnSingle, ByVal ActualDy As MSForms.ReturnSingle _
)
    Dim tmp As Variant

    For Each tmp In EventCollection(ScrollBarEventTypes.Scroll)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), ActionX, ActionY, RequestDx, RequestDy, ActualDx, ActualDy
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, ActionX, ActionY, RequestDx, RequestDy, ActualDx, ActualDy
        End If
    Next
End Sub

''*****************************End ScrollBar Events*********************************''


