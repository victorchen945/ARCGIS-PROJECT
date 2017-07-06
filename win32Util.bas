Attribute VB_Name = "win32Util"
Option Explicit

Const SW_SHOW = 5, SW_RESTORE = 9
Const SWP_NOSIZE = 1
Const SWP_NOMOVE = 2
Const SWP_NOACTIVATE = 16
Const FLOAT_FLAGS = SWP_NOMOVE + SWP_NOSIZE + SWP_NOACTIVATE
Const HWND_TOP = 0
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, _
ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal _
hWndNewParent As Long) As Long

'*** Cause window to float/not float on top of all
'***   other windows; window will float if "Floating"
'***   is True, otherwise window will not float
'------------------------------------------------------------------------------------
Public Sub FloatWindow(frm As Form, FloatIt As Boolean, Optional bOnlyBringToFront _
As Boolean)
    
    Dim v As Variant
    
     If FloatIt Then
         If Not bOnlyBringToFront Then
             v = SetWindowPos(frm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLOAT_FLAGS)
         Else
             v = SetWindowPos(frm.hwnd, HWND_TOP, 0, 0, 0, 0, FLOAT_FLAGS)
         End If
     Else
         v = SetWindowPos(frm.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLOAT_FLAGS)
     End If
    
End Sub

' Sets the specified form to be above the window specified by handle
'------------------------------------------------------------------------------------
Public Sub FloatWindowAbove(frm As Form, lHandle As Long)
    Dim r As Long
     r = SetParent(frm.hwnd, lHandle)
     r = SetWindowPos(lHandle, frm.hwnd, 0, 0, 0, 0, FLOAT_FLAGS)
End Sub




