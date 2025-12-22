' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

Option Strict On
Option Explicit On

Namespace SharedLibrary
    Partial Public Class SharedMethods

        ' Use this file to define URLs and version qualifiers specific to your own build, provided the Constant DEVELOP has been defined in the SharedLibrary.vbproj file

#If DEVELOP Then
        Public Shared ReadOnly Property AppsUrl As String = "https://redink.ai/apps"
        Public Shared ReadOnly Property AppsUrlDir As String = "/develop/"
        Public Shared ReadOnly Property VersionQualifier As String = " Develop"
#End If

    End Class
End Namespace