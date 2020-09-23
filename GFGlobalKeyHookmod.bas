Attribute VB_Name = "GFGlobalKeyHookmod"
Option Explicit
'(c)2001, 2004 by Louis. Use to set up a global key hook to allow a target project to use hot keys.
'Code partially copied from NN99 (06.01.04).
'
'Interface sub (copy to target form):
'Public Sub GFGlobalKeyHookProc(ByVal SourceDescription As String, ByVal KeyCode As Integer, ByVal Shift As Integer, ByRef ReturnValueUsedFlag As Boolean, ByRef ReturnValue As Long)
'    'on error resume next
'End Sub
'
'NOTE: the target project must define and process hot keys.
'Also informing the user about hot keys is the task of the target project.
'It is recommended to use a sub called 'DefineHotKeys' for defining
'the hot keys at program start up.
'
'NN99 CODE >>>
'(c)1999, 2000 by daynight.
'NOTE: parts of code have been copied to the KeyHook Sonde File project (04-16-2000).
'[Set/Remove]KeyHook
Declare Sub SetKH Lib "GFGKH.dll" Alias "noname_sub001" (ByVal MsgTargetAddress As Long, ByVal HookDLLName As String)
Declare Sub RemoveKH Lib "GFGKH.dll" Alias "noname_sub002" ()
'KeyHookProcSub
Private Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal nVirtKey As Long) As Integer
'[Set/Remove]MessageHook
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
'<<< END OF NN99 CODE
'GFGlobalKeyHookStruct
Private Type GFGlobalKeyHookStruct
    KeyHookEnabledFlag As Boolean
    KeyHookTargetFormName As String
    KeyHookTargetForm As Object
End Type
Dim GFGlobalKeyHookStructNumber As Integer
Dim GFGlobalKeyHookStructArray() As GFGlobalKeyHookStruct
'other
Dim KeyHookEnabledFlag As Boolean 'if key hook has been set up once
Dim KeyHookHandle As Long
'old NN99 code
Dim MessageHookEnabledFlag As Boolean
Dim MessageHookKHhWndUnchanged As Long
Dim MessageHookhWndUnchanged As Long
Dim HookDLLHandle As Long

Public Sub GFGlobalKeyHook_SetKeyHook(ByVal KeyHookTargetFormName As String, ByRef KeyHookTargetForm As Object)
    'on error Resume Next 'add another form to the KeyHook target form buffer
    Dim StructIndex As Integer
    Dim StructLoop As Integer
    '
    'NOTE: call this sub to set a form into the 'key hook event notification queue'.
    'Call [...]_RemoveKeyHook to remove the form again.
    'The key hook itself will not be removed until GFGlobalKeyHook_Terminate is called.
    '
    'preset
    StructIndex = 0 'reset (error)
    For StructLoop = 1 To GFGlobalKeyHookStructNumber
        If GFGlobalKeyHookStructArray(StructLoop).KeyHookTargetFormName = KeyHookTargetFormName Then
            StructIndex = StructLoop
            Exit For
        End If
    Next StructLoop
    'begin
    If StructIndex = 0 Then
        'create new array element to add target form
        If Not (GFGlobalKeyHookStructNumber = 32766) Then 'verify
            GFGlobalKeyHookStructNumber = GFGlobalKeyHookStructNumber + 1
        Else
            MsgBox "internal error in GFGlobalKeyHook_SetKeyHook(): overflow !", vbOKOnly + vbExclamation 'damn it!
            Exit Sub 'error
        End If
        ReDim Preserve GFGlobalKeyHookStructArray(1 To GFGlobalKeyHookStructNumber) As GFGlobalKeyHookStruct
        GFGlobalKeyHookStructArray(GFGlobalKeyHookStructNumber).KeyHookEnabledFlag = True
        GFGlobalKeyHookStructArray(GFGlobalKeyHookStructNumber).KeyHookTargetFormName = KeyHookTargetFormName
        Set GFGlobalKeyHookStructArray(GFGlobalKeyHookStructNumber).KeyHookTargetForm = KeyHookTargetForm
        'enable key hook if not done yet
        If KeyHookEnabledFlag = False Then
            Call SetMessageHook(KeyHookTargetForm.KH)
            Call SetKeyHook(KeyHookTargetForm.KH)
            KeyHookEnabledFlag = True 'do here as RemoveKH also alters KeyHookEnabledFlag
        End If
    Else
        'add target form
        GFGlobalKeyHookStructArray(GFGlobalKeyHookStructNumber).KeyHookEnabledFlag = True
        GFGlobalKeyHookStructArray(GFGlobalKeyHookStructNumber).KeyHookTargetFormName = KeyHookTargetFormName 'senseless
        Set GFGlobalKeyHookStructArray(GFGlobalKeyHookStructNumber).KeyHookTargetForm = KeyHookTargetForm
    End If
    Exit Sub
End Sub
    
Public Sub GFGlobalKeyHook_RemoveKeyHook(ByVal KeyHookTargetFormName As String, ByRef KeyHookTargetForm As Object)
    'on error Resume Next 'call to prevent a target form from receiving messages (call [...]_SetKeyHook() to enable message receiving again)
    Dim StructIndex As Integer
    Dim StructLoop As Integer
    'preset
    StructIndex = 0 'reset (error)
    For StructLoop = 1 To GFGlobalKeyHookStructNumber
        If GFGlobalKeyHookStructArray(StructLoop).KeyHookTargetFormName = KeyHookTargetFormName Then
            StructIndex = StructLoop
            Exit For
        End If
    Next StructLoop
    'begin
    If Not (StructIndex = 0) Then 'verify
        GFGlobalKeyHookStructArray(StructIndex).KeyHookEnabledFlag = False
    End If
End Sub

Public Sub GFGlobalKeyHook_Terminate()
    'on error Resume Next 'call when unloading target project
    If KeyHookEnabledFlag = True Then
        Call RemoveKH
        Call RemoveMessageHook
        KeyHookEnabledFlag = False 'reset (do here as RemoveKH also alters KeyHookEnabledFlag)
    End If
End Sub

Public Function GFGlobalKeyHook_KeyHookProc(ByVal KeyCode As Long, ByVal KeyModifierCode As Long) As Long
    'on error Resume Next 'code mainly copied from NN99
    Dim ReturnValueUsedFlag As Boolean
    Dim ReturnValue As Long
    Dim Shift As Integer
    Dim StructLoop As Integer
    'begin
    For StructLoop = 1 To GFGlobalKeyHookStructNumber
        Call GFGlobalKeyHookStructArray(StructLoop).KeyHookTargetForm.GFGlobalKeyHookProc( _
            GFGlobalKeyHookStructArray(StructLoop).KeyHookTargetFormName, KeyCode, CInt(KeyModifierCode), ReturnValueUsedFlag, ReturnValue)
    Next StructLoop
End Function

'***NN99 CODE***
'NOTE: the following code has been copied from NN99 (06.01.04) and altered.

Public Sub SetKeyHook(ByRef KH As PictureBox)
    'On Error Resume Next
    If KeyHookEnabledFlag = False Then
        KeyHookEnabledFlag = True
        HookDLLHandle = LoadLibrary("GFGKH.dll")
        Call SetKH(KH.hWnd, "GFGKH.dll")
    End If
End Sub

Public Sub RemoveKeyHook()
    'On Error Resume Next
    If KeyHookEnabledFlag = True Then
        KeyHookEnabledFlag = False 'reset
        Call RemoveKH
        Call FreeLibrary(HookDLLHandle)
    End If
End Sub

Public Sub SetMessageHook(ByRef KH As PictureBox)
    'On Error Resume Next
    If MessageHookEnabledFlag = False Then
        MessageHookEnabledFlag = True
        MessageHookKHhWndUnchanged = KH.hWnd 'store current handle (used also in RemoveMessageHook)
        MessageHookhWndUnchanged = SetWindowLong(MessageHookKHhWndUnchanged, (-4), AddressOf MessageHookProcSub)
    End If
End Sub

Public Sub RemoveMessageHook()
    'On Error Resume Next
    If MessageHookEnabledFlag = True Then
        MessageHookEnabledFlag = False 'reset
        Call SetWindowLong(MessageHookKHhWndUnchanged, (-4), MessageHookhWndUnchanged)
    End If
End Sub

Public Function MessageHookProcSub(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    'On Error Resume Next
    Select Case Msg
    Case 0 'NULL message
        If Not ((GetAsyncKeyState(wParam) And &H8001) = 0) Then 'check for keydown event
            Call GFGlobalKeyHook_KeyHookProc(wParam, lParam)
        End If
    End Select
    MessageHookProcSub = CallWindowProc(MessageHookhWndUnchanged, hWnd, Msg, wParam, lParam)
End Function

'***END OF NN99 CODE***
'***END OF MODULE***

