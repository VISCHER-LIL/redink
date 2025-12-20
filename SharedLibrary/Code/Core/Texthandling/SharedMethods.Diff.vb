' Part of: Red Ink Shared Library
' Copyright by David Rosenthal, david.rosenthal@vischer.com
' May only be used under with an appropriate license (see vischer.com/redink)

Option Strict On
Option Explicit On

Namespace SharedLibrary

    Partial Public Class SharedMethods

        Public Class Diff
            Public Enum Operation
                Equal
                Insert
                Delete
            End Enum

            Public Property Op As Operation
            Public Property Text As String

            Public Sub New(op As Operation, text As String)
                Me.Op = op
                Me.Text = text
            End Sub
        End Class

    End Class

End Namespace