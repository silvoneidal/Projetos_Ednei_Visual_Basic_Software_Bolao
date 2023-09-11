Attribute VB_Name = "mdIInputbox"
'Option Compare Database
'////////////////////////////////////////////////////////////////////
'Password masked inputbox
'Allows you to hide characters entered in a VBA Inputbox.
'
'Code written by Daniel Klann
'http://www.danielklann.com/
'March 2003

'// Kindly permitted to be amended
'// Amended by Ivan F Moala
'// [Você precisa estar registrado e conectado para ver este link.]
'// April 2003
'// Works for Xl2000+ due the AddressOf Operator
'////////////////////////////////////////////////////////////////////

'******************** CALL FROM FORM *********************************
' Dim pwd As String
'
' pwd = InputBoxDK("Please Enter Password Below!", "Database Administration Security Form.")
'
' 'If no password was entered.
' If pwd = "" Then
' MsgBox "You didn't enter a password! You must enter password to 'enter the Administration Screen!" _
' , vbInformation, "Security Warning"
' End If
'**************************************



'API functions to be used
Private Declare Function CallNextHookEx _
Lib "User32" ( _
ByVal hHook As Long, _
ByVal ncode As Long, _
ByVal wParam As Long, _
LParam As Any) _
As Long

Private Declare Function GetModuleHandle _
Lib "Kernel32" _
Alias "GetModuleHandleA" ( _
ByVal lpModuleName As String) _
As Long

Private Declare Function SetWindowsHookEx _
Lib "User32" _
Alias "SetWindowsHookExA" ( _
ByVal idHook As Long, _
ByVal lpfn As Long, _
ByVal hmod As Long, _
ByVal dwThreadId As Long) _
As Long

Private Declare Function UnhookWindowsHookEx _
Lib "User32" ( _
ByVal hHook As Long) _
As Long

Private Declare Function SendDlgItemMessage _
Lib "User32" Alias "SendDlgItemMessageA" ( _
ByVal hDlg As Long, _
ByVal nIDDlgItem As Long, _
ByVal wMsg As Long, _
ByVal wParam As Long, _
ByVal LParam As Long) _
As Long

Private Declare Function GetClassName _
Lib "User32" _
Alias "GetClassNameA" ( _
ByVal hWnd As Long, _
ByVal lpClassName As String, _
ByVal nMaxCount As Long) _
As Long

Private Declare Function GetCurrentThreadId _
Lib "Kernel32" () _
As Long

'Constants to be used in our API functions
Private Const EM_SETPASSWORDCHAR = &HCC
Private Const WH_CBT = 5
Private Const HCBT_ACTIVATE = 5
Private Const HC_ACTION = 0

Private hHook As Long

Public Function NewProc(ByVal lngCode As Long, _
ByVal wParam As Long, _
ByVal LParam As Long) As Long

Dim RetVal
Dim strClassName As String, lngBuffer As Long

If lngCode < HC_ACTION Then
NewProc = CallNextHookEx(hHook, lngCode, wParam, LParam)
Exit Function
End If

strClassName = String$(256, " ")
lngBuffer = 255

If lngCode = HCBT_ACTIVATE Then 'A window has been activated
RetVal = GetClassName(wParam, strClassName, lngBuffer)
If Left$(strClassName, RetVal) = "#32770" Then 'Class name of the Inputbox
'This changes the edit control so that it display the password character *.
'You can change the Asc("*") as you please.
SendDlgItemMessage wParam, &H1324, EM_SETPASSWORDCHAR, Asc("*"), &H0
End If
End If

'This line will ensure that any other hooks that may be in place are
'called correctly.
CallNextHookEx hHook, lngCode, wParam, LParam

End Function

'// Make it public = avail to ALL Modules
'// Lets simulate the VBA Input Function
Public Function InputBoxDK(Prompt As String, Optional Title As String, _
Optional Default As String, _
Optional Xpos As Long, _
Optional Ypos As Long, _
Optional Helpfile As String, _
Optional Context As Long) As String

Dim lngModHwnd As Long, lngThreadID As Long

'// Lets handle any Errors JIC! due to HookProc> App hang!
On Error GoTo ExitProperly
lngThreadID = GetCurrentThreadId
lngModHwnd = GetModuleHandle(vbNullString)

hHook = SetWindowsHookEx(WH_CBT, AddressOf NewProc, lngModHwnd, lngThreadID)
If Xpos Then
InputBoxDK = InputBox(Prompt, Title, Default, Xpos, Ypos, Helpfile, Context)
Else
InputBoxDK = InputBox(Prompt, Title, Default, , , Helpfile, Context)
End If

ExitProperly:
UnhookWindowsHookEx hHook

End Function

