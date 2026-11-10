Option Strict On
Imports System.Runtime.InteropServices

Public Module ModWindowsApi
    Public Const MB_ICONINFORMATION As UInteger = &H40UI
    Public Const WM_HOTKEY As Integer = &H312
    Public Const ID_MULAI As Integer = 1
    Public Const ID_ULANGI As Integer = 2
    Public Const MOD_CONTROL As UInteger = &H2UI
    Public Const MOD_ALT As UInteger = &H1UI
    Public Const MOD_NOREPEAT As UInteger = &H4000UI
    Public Const VK_F9 As UInteger = &H78UI
    Public Const VK_F10 As UInteger = &H79UI
    Public Const SWP_NOSIZE As UInteger = &H1UI
    Public Const SWP_NOMOVE As UInteger = &H2UI
    Public Const SWP_NOACTIVATE As UInteger = &H10UI
    Public ReadOnly HWND_TOPMOST As New IntPtr(-1)
    Public ReadOnly HWND_NOTOPMOST As New IntPtr(-2)

    Public Declare Function GetTickCount64 Lib "kernel32.dll" () As ULong

    Public Declare Function MessageBeep Lib "user32.dll" (
        uType As UInteger) As Boolean

    <DllImport("user32.dll", ExactSpelling:=True, SetLastError:=True)>
    Public Function SetWindowPos(hWnd As IntPtr, hWndInsertAfter As IntPtr,
        x As Integer, y As Integer, cx As Integer, cy As Integer,
        uFlags As UInteger) As Boolean
    End Function

    <DllImport("user32.dll", ExactSpelling:=True, SetLastError:=True)>
    Public Function RegisterHotKey(hWnd As IntPtr, id As Integer,
        fsModifiers As UInteger, vk As UInteger) As Boolean
    End Function

    <DllImport("user32.dll", ExactSpelling:=True, SetLastError:=True)>
    Public Function UnregisterHotKey(hWnd As IntPtr, id As Integer) As Boolean
    End Function

    Public Function HitungSisa(durasi As ULong, terpakai As ULong) As ULong
        If terpakai >= durasi Then Return 0UL
        Return durasi - terpakai
    End Function

    Public Function FormatSisa(milidetik As ULong) As String
        Dim detik As ULong = milidetik \ 1000UL
        If milidetik Mod 1000UL > 0UL Then detik += 1UL
        Return $"{detik \ 60UL:00}:{detik Mod 60UL:00}"
    End Function
End Module
