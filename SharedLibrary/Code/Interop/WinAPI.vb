' Part of: Red Ink Shared Library
' Copyright by David Rosenthal, david.rosenthal@vischer.com
' May only be used under with an appropriate license (see vischer.com/redink)


Option Strict On
Option Explicit On

Imports System.Runtime.InteropServices

Namespace SharedLibrary
    Module WinAPI
        <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
        Public Function FindWindow(lpClassName As String, lpWindowName As String) As IntPtr
        End Function
    End Module

End Namespace