VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ControlEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''
' Allows VBA Developers to Register Progammatically Created FormFrame to Events. Events will
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
'           Static crtEvent As New ControlEvent
'
'           '' Sets Label Object Target for events
'           Set crtEvent.FormControl = Me.Image1
'           '' Adds a Function Event Subscriber
'           crtEvent.AddEvent "My_Enter", ControlEventTypes.Enter, Me
'       End Sub
'
'       Sub My_Enter(Sender As MSForms.Control)
'           Debug.Print "Control Entered"
'       End Sub
'   2:
'     The following example assumes that this event object is being create inside of a Module
'     code block, thus the third parameter is being ommited.
'
'       Sub AddEventListener()
'           '' Allows the event object to persist.
'           Static imgEvent As New ControlEvent
'
'           '' Sets Label Object Target for events
'           Set imgEvent.FormControl = UserForm1.Image1
'           '' Adds a Function Event Subscriber
'           crtEvent.AddEvent "My_Enter", ControlEventTypes.Enter
'       End Sub
'
'       Sub My_Enter(Sender As MSForms.Control)
'           Debug.Print "Control Entered"
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

Public Enum ControlEventTypes
    AfterUpdate
    BeforeUpdate
    Enter
    ExitEvent
End Enum

'' Public Event Object Property
Public WithEvents FormControl As MSForms.Control
Attribute FormControl.VB_VarHelpID = -1

'' Public Function Calls
Private EventCollection As Collection

''
' AddEvent
'   This method allows events to added. Function Names MUST be unique even if an event is stored
'   in two different objects, having the same name will cause a collision.
' Parameters
'   FuncName  As String         - The name of the function to call. This name MUST be globally unique.
'   EventType As ControlEventTypes - An enum representing the event you wish to subscript to.
'   Obj       As Object         - An optional object instance.
Public Sub AddEvent(ByVal FuncName As String, ByVal EventType As ControlEventTypes, Optional Obj As Object = Nothing)
    EventCollection(EventType).Add Array(FuncName, Obj), FuncName
End Sub
''
' RemoveEvent
'   This method allows events to be removed.
' Parameters
'   FuncName  As String         - The name of the function to remove.
'   EventType As ControlEventTypes - An enum representing the event the function was subscripted to.
Public Sub RemoveEvent(ByVal FuncName As String, ByVal EventType As ControlEventTypes)
    EventCollection(EventType).Remove FuncName
End Sub

'' Initilize Class
Private Sub Class_Initialize()
    Set EventCollection = New Collection
    
    EventCollection.Add New Collection, CStr(ControlEventTypes.AfterUpdate)
    EventCollection.Add New Collection, CStr(ControlEventTypes.BeforeUpdate)
    EventCollection.Add New Collection, CStr(ControlEventTypes.Enter)
    EventCollection.Add New Collection, CStr(ControlEventTypes.ExitEvent)
End Sub

''******************************Control Events************************************''

Private Sub FormControl_AfterUpdate()
    Dim tmp As Variant

    For Each tmp In EventCollection(ControlEventTypes.AfterUpdate)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), FormControl
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, FormControl
        End If
    Next
End Sub

Private Sub FormControl_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    Dim tmp As Variant

    For Each tmp In EventCollection(ControlEventTypes.BeforeUpdate)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), FormControl, Cancel
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, FormControl, Cancel
        End If
    Next
End Sub

Private Sub FormControl_Enter()
    Dim tmp As Variant

    For Each tmp In EventCollection(ControlEventTypes.Enter)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), FormControl
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, FormControl
        End If
    Next
End Sub

Private Sub FormControl_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Dim tmp As Variant

    For Each tmp In EventCollection(ControlEventTypes.ExitEvent)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), FormControl, Cancel
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, FormControl, Cancel
        End If
    Next
End Sub

''*****************************End Control Events*********************************''

