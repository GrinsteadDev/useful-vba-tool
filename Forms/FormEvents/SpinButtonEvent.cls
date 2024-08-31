VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SpinButtonEvent"
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
'           Static spinBtnEvent As New SpinButtonEvent
'
'           '' Sets Label Object Target for events
'           Set spinBtnEvent.FormSpinButton = Me.SpinButton1
'           '' Adds a Function Event Subscriber
'           spinBtnEvent.AddEvent "My_Change", SpinButtonEventTypes.Change, Me
'       End Sub
'
'       Sub My_Change(Sender As MSForms.SpinButton)
'           MsgBox "I was Changed"
'       End Sub
'
'   2:
'     The following example assumes that this event object is being create inside of a Module
'     code block, thus the third parameter is being ommited.
'
'       Sub AddEventListener()
'           '' Allows the event object to persist.
'           Static spinBtnEvent As New SpinButtonEvent
'
'           '' Sets Label Object Target for events
'           Set spinBtnEvent.FormSpinButton = UserForm1.SpinButton1
'           '' Adds a Function Event Subscriber
'           spinBtnEvent.AddEvent "My_Change", SpinButtonEventTypes.Change
'       End Sub
'
'       Sub My_Change(Sender As MSForms.SpinButton)
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

Public Enum SpinButtonEventTypes
    BeforeDragOver
    BeforeDropOrPast
    Change
    ErrorEvent
    KeyDown
    KeyPress
    KeyUp
    SpinDown
    SpinUp
End Enum

'' Public Event Object Property
Public WithEvents FormSpinButton As MSForms.SpinButton
Attribute FormSpinButton.VB_VarHelpID = -1

'' Public Function Calls
Private EventCollection As Collection

''
' AddEvent
'   This method allows events to added. Function Names MUST be unique even if an event is stored
'   in two different objects, having the same name will cause a collision.
' Parameters
'   FuncName  As String          - The name of the function to call. This name MUST be globally unique.
'   EventType As SpinButtonEventTypes - An enum representing the event you wish to subscript to.
'   Obj       As Object          - An optional object instance.
Public Sub AddEvent(ByVal FuncName As String, ByVal EventType As SpinButtonEventTypes, Optional Obj As Object = Nothing)
    EventCollection(EventType).Add Array(FuncName, Obj), FuncName
End Sub
''
' RemoveEvent
'   This method allows events to be removed.
' Parameters
'   FuncName  As String          - The name of the function to remove.
'   EventType As SpinButtonEventTypes - An enum representing the event the function was subscripted to.
Public Sub RemoveEvent(ByVal FuncName As String, ByVal EventType As SpinButtonEventTypes)
    EventCollection(EventType).Remove FuncName
End Sub

'' Initilize Class
Private Sub Class_Initialize()
    Set EventCollection = New Collection
    
    EventCollection.Add New Collection, CStr(SpinButtonEventTypes.BeforeDragOver)
    EventCollection.Add New Collection, CStr(SpinButtonEventTypes.BeforeDropOrPast)
    EventCollection.Add New Collection, CStr(SpinButtonEventTypes.Change)
    EventCollection.Add New Collection, CStr(SpinButtonEventTypes.ErrorEvent)
    EventCollection.Add New Collection, CStr(SpinButtonEventTypes.KeyDown)
    EventCollection.Add New Collection, CStr(SpinButtonEventTypes.KeyPress)
    EventCollection.Add New Collection, CStr(SpinButtonEventTypes.KeyUp)
    EventCollection.Add New Collection, CStr(SpinButtonEventTypes.SpinDown)
    EventCollection.Add New Collection, CStr(SpinButtonEventTypes.SpinUp)
End Sub

''******************************SpinButton Events************************************''

Private Sub FormSpinButton_BeforeDragOver( _
    ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, _
    ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, _
    ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer _
)
    Dim tmp As Variant
    
    For Each tmp In EventCollection(SpinButtonEventTypes.BeforeDragOver)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), FormSpinButton, Cancel, Data, X, Y, DragState, Effect, Shift
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, FormSpinButton, Cancel, Data, X, Y, DragState, Effect, Shift
        End If
    Next
End Sub

Private Sub FormSpinButton_BeforeDropOrPaste( _
    ByVal Cancel As MSForms.ReturnBoolean, ByVal Action As MSForms.fmAction, _
    ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, _
    ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer _
)
    Dim tmp As Variant
    
    For Each tmp In EventCollection(SpinButtonEventTypes.BeforeDropOrPast)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), FormSpinButton, Cancel, Action, Data, X, Y, Effect, Shift
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, FormSpinButton, Cancel, Action, Data, X, Y, Effect, Shift
        End If
    Next
End Sub

Private Sub FormSpinButton_Change()
    Dim tmp As Variant
    
    For Each tmp In EventCollection(SpinButtonEventTypes.Change)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), FormSpinButton
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, FormSpinButton
        End If
    Next
End Sub

Private Sub FormSpinButton_Error( _
    ByVal Number As Integer, ByVal Description As MSForms.ReturnString, _
    ByVal SCode As Long, ByVal Source As String, _
    ByVal HelpFile As String, ByVal HelpContext As Long, _
    ByVal CancelDisplay As MSForms.ReturnBoolean _
)
    Dim tmp As Variant
    
    For Each tmp In EventCollection(SpinButtonEventTypes.ErrorEvent)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), FormSpinButton, Number, Description, SCode, Source, HelpFile, HelpContext, CancelDisplay
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, FormSpinButton, Number, Description, SCode, Source, HelpFile, HelpContext, CancelDisplay
        End If
    Next
End Sub

Private Sub FormSpinButton_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim tmp As Variant
    
    For Each tmp In EventCollection(SpinButtonEventTypes.KeyDown)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), FormSpinButton, KeyCode, Shift
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, FormSpinButton, KeyCode, Shift
        End If
    Next
End Sub

Private Sub FormSpinButton_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim tmp As Variant
    
    For Each tmp In EventCollection(SpinButtonEventTypes.KeyPress)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), FormSpinButton, KeyAscii
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, FormSpinButton, KeyAscii
        End If
    Next
End Sub

Private Sub FormSpinButton_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim tmp As Variant
    
    For Each tmp In EventCollection(SpinButtonEventTypes.KeyUp)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), FormSpinButton, KeyCode, Shift
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, FormSpinButton, KeyCode, Shift
        End If
    Next
End Sub

Private Sub FormSpinButton_SpinDown()
    Dim tmp As Variant
    
    For Each tmp In EventCollection(SpinButtonEventTypes.SpinDown)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), FormSpinButton
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, FormSpinButton
        End If
    Next
End Sub

Private Sub FormSpinButton_SpinUp()
    Dim tmp As Variant
    
    For Each tmp In EventCollection(SpinButtonEventTypes.SpinUp)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), FormSpinButton
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, FormSpinButton
        End If
    Next
End Sub

''*****************************End SpinButton Events*********************************''


