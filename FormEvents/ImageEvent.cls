VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ImageEvent"
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
'           Static imgEvent As New ImageEvent
'
'           '' Sets Label Object Target for events
'           Set imgEvent.FormImage = Me.Image1
'           '' Adds a Function Event Subscriber
'           imgEvent.AddEvent "My_MouseDown", ImageEventTypes.MouseDown, Me
'       End Sub
'
'       Sub My_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'           '' Prints the Mouse Button Pressed
'           Debug.Print "Mouse Button Pressed " & CStr(Button)
'       End Sub
'   2:
'     The following example assumes that this event object is being create inside of a Module
'     code block, thus the third parameter is being ommited.
'
'       Sub AddEventListener()
'           '' Allows the event object to persist.
'           Static imgEvent As New ImageEvent
'
'           '' Sets Label Object Target for events
'           Set imgEvent.FormImage = UserForm1.Image1
'           '' Adds a Function Event Subscriber
'           imgEvent.AddEvent "My_MouseDown", ImageEventTypes.MouseDown
'       End Sub
'
'       Sub My_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'           '' Prints the Mouse Button Pressed
'           Debug.Print "Mouse Button Pressed " & CStr(Button)
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

Public Enum ImageEventTypes
    BeforeDragOver
    BeforeDropOrPaste
    ErrorEvent
    MouseDown
    MouseMove
    MouseUp
End Enum

'' Public Event Object Property
Public WithEvents FormImage As MSForms.Image
Attribute FormImage.VB_VarHelpID = -1

'' Public Function Calls
Private EventCollection As Collection

''
' AddEvent
'   This method allows events to added. Function Names MUST be unique even if an event is stored
'   in two different objects, having the same name will cause a collision.
' Parameters
'   FuncName  As String         - The name of the function to call. This name MUST be globally unique.
'   EventType As ImageEventTypes - An enum representing the event you wish to subscript to.
'   Obj       As Object         - An optional object instance.
Public Sub AddEvent(ByVal FuncName As String, ByVal EventType As ImageEventTypes, Optional Obj As Object = Nothing)
    EventCollection(EventType).Add Array(FuncName, Obj), FuncName
End Sub
''
' RemoveEvent
'   This method allows events to be removed.
' Parameters
'   FuncName  As String         - The name of the function to remove.
'   EventType As ImageEventTypes - An enum representing the event the function was subscripted to.
Public Sub RemoveEvent(ByVal FuncName As String, ByVal EventType As ImageEventTypes)
    EventCollection(EventType).Remove FuncName
End Sub

'' Initilize Class
Private Sub Class_Initialize()
    Set EventCollection = New Collection
    
    EventCollection.Add New Collection, CStr(ImageEventTypes.BeforeDragOver)
    EventCollection.Add New Collection, CStr(ImageEventTypes.BeforeDropOrPaste)
    EventCollection.Add New Collection, CStr(ImageEventTypes.ErrorEvent)
    EventCollection.Add New Collection, CStr(ImageEventTypes.MouseDown)
    EventCollection.Add New Collection, CStr(ImageEventTypes.MouseMove)
    EventCollection.Add New Collection, CStr(ImageEventTypes.MouseUp)
End Sub

''******************************Form Events************************************''

Private Sub FormImage_BeforeDragOver( _
    ByVal Cancel As MSForms.ReturnBoolean, ByVal Control As MSForms.Control, _
    ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, _
    ByVal State As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, _
    ByVal Shift As Integer _
)
    Dim tmp As Variant

    For Each tmp In EventCollection(ImageEventTypes.BeforeDragOver)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), Cancel, Control, Data, X, Y, State, Effect, Shift
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, Cancel, Control, Data, X, Y, State, Effect, Shift
        End If
    Next
End Sub

Private Sub FormImage_BeforeDropOrPaste( _
    ByVal Cancel As MSForms.ReturnBoolean, ByVal Control As MSForms.Control, _
    ByVal Action As MSForms.fmAction, ByVal Data As MSForms.DataObject, ByVal X As Single, _
    ByVal Y As Single, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer _
)
    Dim tmp As Variant

    For Each tmp In EventCollection(ImageEventTypes.BeforeDropOrPaste)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), Cancel, Control, Action, Data, X, Y, Effect, Shift
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, Cancel, Control, Action, Data, X, Y, Effect, Shift
        End If
    Next
End Sub

Private Sub FormImage_Error( _
    ByVal Number As Integer, ByVal Description As MSForms.ReturnString, _
    ByVal SCode As Long, ByVal Source As String, ByVal HelpFile As String, _
    ByVal HelpContext As Long, ByVal CancelDisplay As MSForms.ReturnBoolean _
)
    Dim tmp As Variant

    For Each tmp In EventCollection(ImageEventTypes.ErrorEvent)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), Number, Description, SCode, Source, HelpFile, HelpContext, CancelDisplay
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, Number, Description, SCode, Source, HelpFile, HelpContext, CancelDisplay
        End If
    Next
End Sub

Private Sub FormImage_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim tmp As Variant

    For Each tmp In EventCollection(ImageEventTypes.MouseDown)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), Button, Shift, X, Y
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, Button, Shift, X, Y
        End If
    Next
End Sub

Private Sub FormImage_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim tmp As Variant

    For Each tmp In EventCollection(ImageEventTypes.MouseMove)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), Button, Shift, X, Y
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, Button, Shift, X, Y
        End If
    Next
End Sub

Private Sub FormImage_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim tmp As Variant

    For Each tmp In EventCollection(ImageEventTypes.MouseUp)
        If tmp(EventObj.ObjectIndex) Is Nothing Then
            Application.Run tmp(EventObj.NameIndex), Button, Shift, X, Y
        Else
            VBA.CallByName tmp(EventObj.ObjectIndex), tmp(EventObj.NameIndex), VbMethod, Button, Shift, X, Y
        End If
    Next
End Sub

''*****************************End FormFrame Events*********************************''

