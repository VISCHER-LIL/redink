' Part of: Red Ink Shared Library
' Copyright by David Rosenthal, david.rosenthal@vischer.com
' May only be used under with an appropriate license (see vischer.com/redink)


Option Strict On
Option Explicit On

Namespace SharedLibrary
    Public Class NativeMethods
        <Runtime.InteropServices.DllImport("user32.dll")>
        Public Shared Function SetForegroundWindow(hWnd As IntPtr) As Boolean
        End Function
    End Class

End Namespace