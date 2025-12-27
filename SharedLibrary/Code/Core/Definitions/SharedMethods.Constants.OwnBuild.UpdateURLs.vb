' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: SharedMethods.Constants.OwnBuild.UpdateURLs.vb
' Purpose: Defines update-related URLs and version qualifier values for a local/own
'          build, enabled only when the `DEVELOP` compilation constant is defined.
'
' Architecture / How it works:
'  - This file is part of the `SharedMethods` partial class and conditionally
'    compiles only for `#If DEVELOP Then` builds.
'  - It provides the `AppsUrl`, `AppsUrlDir`, `VersionQualifier`, and
'    `DefaultUpdateIntervalDays` members used by the update/configuration logic.
'  - For non-`DEVELOP` builds, these members are expected to be provided by other
'    compilation branches in `SharedMethods.Constants.vb`.
' =============================================================================


Option Strict On
Option Explicit On

Namespace SharedLibrary
    Partial Public Class SharedMethods

        ' Use this file to define URLs and version qualifiers specific to your own build, provided the Constant DEVELOP has been defined in the SharedLibrary.vbproj file

#If DEVELOP Then
        Public Shared ReadOnly Property AppsUrl As String = "https://redink.ai/apps"
        Public Shared ReadOnly Property AppsUrlDir As String = "/develop/"
        Public Shared ReadOnly Property VersionQualifier As String = " Develop"

        Public Shared ReadOnly DefaultUpdateIntervalDays As Integer = 1

#End If

    End Class
End Namespace