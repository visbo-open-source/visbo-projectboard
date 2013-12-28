Imports System
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices
Imports System.Windows.Forms

Public Class clsKeyboardEvents



    'Public Class KeyboardHooking
    '    ' Methods
    '    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    '    Private Shared Function CallNextHookEx(ByVal hhk As IntPtr, ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
    '    End Function

    '    <DllImport("kernel32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    '    Private Shared Function GetModuleHandle(ByVal lpModuleName As String) As IntPtr
    '    End Function

    '    Private Shared Function HookCallback(ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Integer
    '        If ((nCode >= 0) AndAlso (nCode = 0)) Then
    '            Dim keyData As Keys = DirectCast(CInt(wParam), Keys)
    '            If (((BindingFunctions.IsKeyDown(Keys.ControlKey) AndAlso BindingFunctions.IsKeyDown(Keys.ShiftKey)) AndAlso BindingFunctions.IsKeyDown(keyData)) AndAlso (keyData = Keys.D7)) Then
    '                'DO SOMETHING HERE
    '            End If
    '            If ((BindingFunctions.IsKeyDown(Keys.ControlKey) AndAlso BindingFunctions.IsKeyDown(keyData)) AndAlso (keyData = Keys.D7)) Then
    '                'DO SOMETHING HERE
    '            End If
    '        End If
    '        Return CInt(KeyboardHooking.CallNextHookEx(KeyboardHooking._hookID, nCode, wParam, lParam))
    '    End Function

    '    Public Shared Sub ReleaseHook()
    '        KeyboardHooking.UnhookWindowsHookEx(KeyboardHooking._hookID)
    '    End Sub

    '    Public Shared Sub SetHook()
    '        KeyboardHooking._hookID = KeyboardHooking.SetWindowsHookEx(2, KeyboardHooking._proc, IntPtr.Zero, Convert.ToUInt32(AppDomain.GetCurrentThreadId))
    '        KeyboardHooking._hookID = KeyboardHooking.SetWindowsHookEx(2, KeyboardHooking._proc, IntPtr.Zero, Convert.ToUInt32(AppDomain.managedThreadID))
    '    End Sub

    '    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    '    Private Shared Function SetWindowsHookEx(ByVal idHook As Integer, ByVal lpfn As LowLevelKeyboardProc, ByVal hMod As IntPtr, ByVal dwThreadId As UInt32) As IntPtr
    '    End Function

    '    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    '    Private Shared Function UnhookWindowsHookEx(ByVal hhk As IntPtr) As <MarshalAs(UnmanagedType.Bool)> Boolean
    '    End Function


    '    ' Fields
    '    Private Shared _hookID As IntPtr = IntPtr.Zero
    '    Private Shared _proc As LowLevelKeyboardProc = New LowLevelKeyboardProc(AddressOf KeyboardHooking.HookCallback)
    '    Private Const WH_KEYBOARD As Integer = 2
    '    Private Const WH_KEYBOARD_LL As Integer = 13
    '    Private Const WM_KEYDOWN As Integer = &H100

    '    ' Nested Types
    '    Public Delegate Function LowLevelKeyboardProc(ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Integer
    'End Class

    'Public Class BindingFunctions
    '    ' Methods
    '    <DllImport("user32.dll")> _
    '    Private Shared Function GetKeyState(ByVal nVirtKey As Integer) As Short
    '    End Function

    '    Public Shared Function IsKeyDown(ByVal keys As Keys) As Boolean
    '        Return ((BindingFunctions.GetKeyState(CInt(keys)) And &H8000) = &H8000)
    '    End Function

    'End Class



End Class
