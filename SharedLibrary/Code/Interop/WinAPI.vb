' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

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