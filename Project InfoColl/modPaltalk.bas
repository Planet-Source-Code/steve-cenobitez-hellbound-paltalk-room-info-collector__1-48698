Attribute VB_Name = "modPaltalk"
Option Explicit
Public Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Boolean

Global lngRhWnd As Long ' global to store handle
Global strWindowTitle As String ' global to store title
Global strClassname As String ' global to store class

Private Function EnumWindowsProc(ByVal hWnd As Long, ByVal lParam As Long) As Boolean
    Dim sSave As String, Ret As Long
    Ret = GetWindowTextLength(hWnd)
    sSave = Space(Ret)
    GetWindowText hWnd, sSave, Ret + 1
    If InStr(sSave, " Group") Then
    lngRhWnd = Str$(hWnd)
    strClassname = GetClass(lngRhWnd)
    strWindowTitle = GetWinTitle(lngRhWnd)
    End If
    EnumWindowsProc = True
End Function

Public Function GetPalWindow(Optional strClassname As String) As Long
    Dim retval As Boolean
    retval = EnumWindows(AddressOf EnumWindowsProc, 0)
    strClassname = GetClass(lngRhWnd)
    GetPalWindow = lngRhWnd
End Function

Private Function GetClass(hWnd As Long)
Dim sSave As String, intLenRet As Integer
sSave = Space(256)
intLenRet = GetClassName(hWnd, sSave, 256)
GetClass = Left$(sSave, intLenRet)
End Function

Private Function GetWinTitle(hWnd As Long) As String
GetWinTitle = GetText(hWnd)
End Function

Private Function GetText(lnghWnd As Long) As String
    Dim TheText As String, TL As Long, XT As Long
        TL = GetWindowTextLength(lnghWnd)
            TheText = String(TL + 1, " ")
        XT = GetWindowText(lnghWnd, TheText, TL + 1)
            TheText = Left(TheText, TL)
        GetText = TheText
End Function
