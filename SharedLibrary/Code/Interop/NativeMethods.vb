' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

Option Strict On
Option Explicit On

Namespace SharedLibrary
    Public Class NativeMethods
        <Runtime.InteropServices.DllImport("user32.dll")>
        Public Shared Function SetForegroundWindow(hWnd As IntPtr) As Boolean
        End Function
    End Class

End Namespace