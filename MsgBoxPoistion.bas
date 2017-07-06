Attribute VB_Name = "MsgBoxPoistion"
 'Microsoft Knowledge Base Article - 180936
 'http://support.microsoft.com/default.aspx?scid=kb;en-us;180936
 
 Type RECT
         Left As Long
         Top As Long
         Right As Long
         Bottom As Long
      End Type

      Public Declare Function UnhookWindowsHookEx Lib "user32" ( _
         ByVal hHook As Long) As Long
      Public Declare Function GetWindowLong Lib "user32" Alias _
         "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) _
         As Long
      Public Declare Function GetCurrentThreadId Lib "kernel32" () As Long
      Public Declare Function SetWindowsHookEx Lib "user32" Alias _
         "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, _
         ByVal hmod As Long, ByVal dwThreadId As Long) As Long
      Public Declare Function SetWindowPos Lib "user32" ( _
         ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
         ByVal x As Long, ByVal y As Long, ByVal cx As Long, _
         ByVal cy As Long, ByVal wFlags As Long) As Long
      Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd _
         As Long, lpRect As RECT) As Long

      Public Const GWL_HINSTANCE = (-6)
      Public Const SWP_NOSIZE = &H1
      Public Const SWP_NOZORDER = &H4
      Public Const SWP_NOACTIVATE = &H10
      Public Const HCBT_ACTIVATE = 5
      Public Const WH_CBT = 5

      Public hHook As Long

      Function WinProc1(ByVal lMsg As Long, ByVal wParam As Long, _
         ByVal lParam As Long) As Long

         If lMsg = HCBT_ACTIVATE Then
            'Show the MsgBox at a fixed location (0,0)
            SetWindowPos wParam, 0, 0, 0, 0, 0, _
                         SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE
            'Release the CBT hook
            UnhookWindowsHookEx hHook
         End If
         WinProc1 = False

      End Function

      Function WinProc2(ByVal lMsg As Long, ByVal wParam As Long, _
         ByVal lParam As Long) As Long

      Dim rectForm As RECT, rectMsg As RECT
      Dim x As Long, y As Long

         'On HCBT_ACTIVATE, show the MsgBox centered over Form1
         If lMsg = HCBT_ACTIVATE Then
            'Get the coordinates of the form and the message box so that
            'you can determine where the center of the form is located
            GetWindowRect Form1.hwnd, rectForm
            GetWindowRect wParam, rectMsg
            x = (rectForm.Left + (rectForm.Right - rectForm.Left) / 2) - _
                ((rectMsg.Right - rectMsg.Left) / 2)
            y = (rectForm.Top + (rectForm.Bottom - rectForm.Top) / 2) - _
                ((rectMsg.Bottom - rectMsg.Top) / 2)
            'Position the msgbox
            SetWindowPos wParam, 0, x, y, 0, 0, _
                         SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE
            'Release the CBT hook
            UnhookWindowsHookEx hHook
         End If
         WinProc2 = False

      End Function
      
      Function WinProc3(ByVal lMsg As Long, ByVal wParam As Long, _
         ByVal lParam As Long) As Long

      Dim rectForm As RECT, rectMsg As RECT
      Dim x As Long, y As Long, x_left As Long, x_right As Long

        x_left = frmLineOfSight.Left - (rectMsg.Right - rectMsg.Left)
        x_right = frmLineOfSight.Left + frmLineOfSight.Width
       
        
         'On HCBT_ACTIVATE, show the MsgBox centered over Form1
         If lMsg = HCBT_ACTIVATE Then
            'Get the coordinates of the form and the message box so that
            'you can determine where the center of the form is located
            GetWindowRect Form1.hwnd, rectForm
            GetWindowRect wParam, rectMsg
            x = (rectForm.Left + (rectForm.Right - rectForm.Left) / 2) - _
                ((rectMsg.Right - rectMsg.Left) / 2)
            y = (rectForm.Top + (rectForm.Bottom - rectForm.Top) / 2) - _
                ((rectMsg.Bottom - rectMsg.Top) / 2)
        
            If (x < x_right) And (x > x_left) Then
                If (frmLineOfSight.Left + frmLineOfSight.Width / 2) < _
                    (rectMsg.Left + (rectMsg.Right - rectMsg.Left) / 2) Then
                    x = x_right
                  Else: x = x_left
                End If
            End If
            
            'Position the msgbox
            SetWindowPos wParam, 0, x, y, 0, 0, _
                         SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE
            'Release the CBT hook
            UnhookWindowsHookEx hHook
         End If
         WinProc3 = False

      End Function



