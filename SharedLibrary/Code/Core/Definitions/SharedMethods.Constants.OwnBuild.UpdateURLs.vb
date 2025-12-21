' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

Option Strict On
Option Explicit On

Namespace SharedLibrary
    Partial Public Class SharedMethods

        Private Shared _appsUrl As String = "https://redink.ai/apps"
        Private Shared _appsUrlDir As String = "/develop/"
        Private Shared _versionQualifier As String = " Develop"

        Public Shared ReadOnly Property AppsUrl As String
            Get
                Return _appsUrl
            End Get
        End Property

        Public Shared ReadOnly Property AppsUrlDir As String
            Get
                Return _appsUrlDir
            End Get
        End Property

        Public Shared ReadOnly Property VersionQualifier As String
            Get
                Return _versionQualifier
            End Get
        End Property

        ' The following code runs when the class is first accessed and calls a private .vb module
        ' that is not in the public repository to allow private overrides of the above values.
        ' You can delete it, if not needed.
        '
        ' The private module could look like this:
        '         Private Shared Sub ApplyPrivateOverrides()
        '            _appsUrl = "https://www.privatesite.com"
        '            _appsUrlDir = "/privatedevelop/"
        '            _versionQualifier = " Develop XYZ"
        '         End Sub

        Shared Sub New()
            ApplyPrivateOverrides()
        End Sub

        Partial Private Shared Sub ApplyPrivateOverrides()
        End Sub

    End Class
End Namespace
