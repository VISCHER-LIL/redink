' Part of "Red Ink for Excel"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: BridgeSubs.vb
' Part of: Red Ink for Excel
' Purpose: COM-visible bridge exposing selected ThisAddIn functionality (primarily
'          asynchronous language and text processing operations plus utility
'          helpers) for consumption from Excel VBA or other COM automation. Needed
'          for context menu and shortcut key functionality (requires the VBA
'          helper)
'
' Architecture:
' - Class BridgeSubs (COM-visible): Thin wrapper delegating each public member to
'   corresponding members on Globals.ThisAddIn. Asynchronous methods return Task
'   (some with Boolean result) mirroring underlying add-in operations.
' - GetFileTextContent: Synchronous text extraction utility selecting a reader
'   based on file extension (.txt/.ini/.csv/.log/.json/.xml/.html/.htm/.rtf/.doc/.docx/.pdf).
'   Delegates specialized formats to helper functions (ReadTextFile, ReadRtfAsText,
'   ReadWordDocument, ReadPdfAsText). Returns error strings or empty string based
'   on ReturnErrorInsteadOfEmpty flag.
' - Partial ThisAddIn.GetAPIConfiguration: Builds a serialized configuration
'   string by collecting key-value pairs (delimited internally by "§§") into a
'   list and joining them with "@@@". Supports two sets of values selected by
'   UseSecondAPI.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Threading.Tasks
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports SLib = SharedLibrary.SharedLibrary.SharedMethods

''' <summary>
''' COM-visible bridge that exposes selected ThisAddIn automation helpers to VBA and other COM callers.
''' </summary>
<ComVisible(True)>
Public Class BridgeSubs
    ''' <summary>
    ''' Asynchronously invokes Globals.ThisAddIn.InLanguage1.
    ''' </summary>
    Public Async Function DoInLanguage1() As Task
        Dim Result As Boolean = Await Globals.ThisAddIn.InLanguage1()
    End Function

    ''' <summary>
    ''' Asynchronously invokes Globals.ThisAddIn.InLanguage2.
    ''' </summary>
    Public Async Function DoInLanguage2() As Task
        Dim Result As Boolean = Await Globals.ThisAddIn.InLanguage2()
    End Function

    ''' <summary>
    ''' Asynchronously invokes Globals.ThisAddIn.InOtherFormulas.
    ''' </summary>
    Public Async Function DoInOtherFormulas() As Task
        Dim Result As Boolean = Await Globals.ThisAddIn.InOtherFormulas()
    End Function

    ''' <summary>
    ''' Asynchronously invokes Globals.ThisAddIn.Correct.
    ''' </summary>
    Public Async Function DoCorrect() As Task
        Dim Result As Boolean = Await Globals.ThisAddIn.Correct()
    End Function

    ''' <summary>
    ''' Asynchronously invokes Globals.ThisAddIn.Improve.
    ''' </summary>
    Public Async Function DoImprove() As Task
        Dim Result As Boolean = Await Globals.ThisAddIn.Improve()
    End Function

    ''' <summary>
    ''' Asynchronously invokes Globals.ThisAddIn.Shorten.
    ''' </summary>
    Public Async Function DoShorten() As Task
        Dim Result As Boolean = Await Globals.ThisAddIn.Shorten()
    End Function

    ''' <summary>
    ''' Asynchronously invokes Globals.ThisAddIn.Anonymize.
    ''' </summary>
    Public Async Function DoAnonymize() As Task
        Dim Result As Boolean = Await Globals.ThisAddIn.Anonymize()
    End Function

    ''' <summary>
    ''' Asynchronously invokes Globals.ThisAddIn.SwitchParty.
    ''' </summary>
    Public Async Function DoSwitchParty() As Task
        Dim Result As Boolean = Await Globals.ThisAddIn.SwitchParty()
    End Function

    ''' <summary>
    ''' Asynchronously invokes Globals.ThisAddIn.FreestyleNM.
    ''' </summary>
    Public Async Function DoFreestyleNM() As Task
        Dim Result As Boolean = Await Globals.ThisAddIn.FreestyleNM()
    End Function

    ''' <summary>
    ''' Asynchronously invokes Globals.ThisAddIn.FreestyleAM.
    ''' </summary>
    Public Async Function DoFreestyleAM() As Task
        Dim Result As Boolean = Await Globals.ThisAddIn.FreestyleAM()
    End Function

    ''' <summary>
    ''' Invokes Globals.ThisAddIn.AdjustHeight with optional silent mode.
    ''' </summary>
    Public Sub DoAdjustHeight(Optional Silent As Boolean = False)
        Globals.ThisAddIn.AdjustHeight(Silent)
    End Sub

    ''' <summary>
    ''' Invokes Globals.ThisAddIn.RegexSearchReplace.
    ''' </summary>
    Public Sub DoRegexSearchReplace()
        Globals.ThisAddIn.RegexSearchReplace()
    End Sub

    ''' <summary>
    ''' Invokes Globals.ThisAddIn.AdjustLegacyNotes.
    ''' </summary>
    Public Sub DoAdjustLegacyNotes()
        Globals.ThisAddIn.AdjustLegacyNotes()
    End Sub

    ''' <summary>
    ''' Invokes Globals.ThisAddIn.AddContextMenu.
    ''' </summary>
    Public Sub DoAddContextMenu()
        Globals.ThisAddIn.AddContextMenu()
    End Sub

    ''' <summary>
    ''' Retrieves serialized API configuration from Globals.ThisAddIn.GetAPIConfiguration.
    ''' </summary>
    ''' <param name="UseSecondAPI">Selects alternative configuration set when True.</param>
    ''' <returns>Configuration string containing delimited key-value pairs.</returns>
    Public Function GetLLMConfig(UseSecondAPI As Boolean) As String
        Dim Result As String = Globals.ThisAddIn.GetAPIConfiguration(UseSecondAPI)
        Return Result
    End Function

    ''' <summary>
    ''' Signs a JWT using a PEM private key via shared library helper.
    ''' </summary>
    Public Function SignJWT(jwtUnsigned As String, privateKeyPem As String) As String
        Return SLib.SignJWT(jwtUnsigned, privateKeyPem)
    End Function

    ''' <summary>
    ''' Returns textual content of supported file types. Uses specialized readers
    ''' based on extension. Returns error strings or empty string depending on flag.
    ''' </summary>
    ''' <param name="filePath">Path to the file to read.</param>
    ''' <param name="ReturnErrorInsteadOfEmpty">When True returns error messages instead of empty string.</param>
    ''' <returns>Extracted text or error/empty string.</returns>
    Public Function GetFileTextContent(ByVal filePath As String, Optional ReturnErrorInsteadOfEmpty As Boolean = True) As String
        Try
            ' Normalize and check the path
            filePath = Path.GetFullPath(filePath)
            If Not File.Exists(filePath) Then
                Return If(ReturnErrorInsteadOfEmpty, "Error: File not found", "")
            End If

            ' Determine file type by extension
            Dim extension As String = Path.GetExtension(filePath).ToLower()

            Select Case extension
                Case ".txt", ".ini", ".csv", ".log", ".json", ".xml", ".html", ".htm"
                    Return ReadTextFile(filePath, ReturnErrorInsteadOfEmpty)

                Case ".rtf"
                    Return ReadRtfAsText(filePath, ReturnErrorInsteadOfEmpty)

                Case ".doc", ".docx"
                    Return ReadWordDocument(filePath, ReturnErrorInsteadOfEmpty)

                Case ".pdf"
                    Return ReadPdfAsText(filePath, ReturnErrorInsteadOfEmpty, False, False).Result

                Case Else
                    Return If(ReturnErrorInsteadOfEmpty, "Error: File type not supported (not txt, rtf, doc, docx, pdf, ini, csv, log, json, xml, html or htm)", "")
            End Select
        Catch ex As UnauthorizedAccessException
            Return If(ReturnErrorInsteadOfEmpty, "Error: Unauthorized access", "")
        Catch ex As IOException
            Return If(ReturnErrorInsteadOfEmpty, "Error: IO Error: " & ex.Message, "")
        Catch ex As System.Exception
            Return If(ReturnErrorInsteadOfEmpty, "Error: Unexpected error: " & ex.Message, "")
        End Try
    End Function

End Class

Partial Public Class ThisAddIn
    ''' <summary>
    ''' Builds a serialized configuration string from primary or secondary API values.
    ''' Each entry formatted as Key§§Value and entries joined by @@@.
    ''' </summary>
    ''' <param name="UseSecondAPI">Selects secondary set when True; otherwise primary.</param>
    ''' <returns>Serialized configuration string.</returns>
    Public Function GetAPIConfiguration(UseSecondAPI As Boolean) As String
        Dim config As New List(Of String)()

        If UseSecondAPI Then
            config.Add("INI_OAuth2§§" & INI_OAuth2_2.ToString)
            config.Add("INI_OAuth2ClientMail§§" & INI_OAuth2ClientMail_2)
            config.Add("INI_OAuth2Scopes§§" & INI_OAuth2Scopes_2)
            config.Add("INI_OAuth2Endpoint§§" & INI_OAuth2Endpoint_2)
            config.Add("INI_OAuth2ATExpiry§§" & INI_OAuth2ATExpiry_2.ToString)
            config.Add("INI_APIKey§§" & INI_APIKey_2)
            config.Add("INI_Temperature§§" & INI_Temperature_2.ToString)
            config.Add("INI_Timeout§§" & INI_Timeout_2)
            config.Add("INI_MaxOutputToken§§" & INI_MaxOutputToken_2.ToString)
            config.Add("INI_Model§§" & INI_Model_2)
            config.Add("INI_Endpoint§§" & INI_Endpoint_2)
            config.Add("INI_HeaderA§§" & INI_HeaderA_2)
            config.Add("INI_HeaderB§§" & INI_HeaderB_2)
            config.Add("INI_APICall§§" & INI_APICall_2.Replace("{objectcall}", ""))
            config.Add("INI_Response§§" & INI_Response_2)
            config.Add("DecodedAPI§§" & DecodedAPI_2)
            'config.Add("INI_APICall_Object§§" & INI_APICALL_Object_2)
        Else
            config.Add("INI_OAuth2§§" & INI_OAuth2.ToString)
            config.Add("INI_OAuth2ClientMail§§" & INI_OAuth2ClientMail)
            config.Add("INI_OAuth2Scopes§§" & INI_OAuth2Scopes)
            config.Add("INI_OAuth2Endpoint§§" & INI_OAuth2Endpoint)
            config.Add("INI_OAuth2ATExpiry§§" & INI_OAuth2ATExpiry.ToString)
            config.Add("INI_APIKey§§" & INI_APIKey)
            config.Add("INI_Temperature§§" & INI_Temperature.ToString)
            config.Add("INI_Timeout§§" & INI_Timeout)
            config.Add("INI_MaxOutputToken§§" & INI_MaxOutputToken.ToString)
            config.Add("INI_Model§§" & INI_Model)
            config.Add("INI_Endpoint§§" & INI_Endpoint)
            config.Add("INI_HeaderA§§" & INI_HeaderA)
            config.Add("INI_HeaderB§§" & INI_HeaderB)
            config.Add("INI_APICall§§" & INI_APICall.Replace("{objectcall}", ""))
            config.Add("INI_Response§§" & INI_Response)
            config.Add("DecodedAPI§§" & DecodedAPI)
            'config.Add("INI_APICall_Object§§" & INI_APICALL_Object)
        End If

        ' Join the list into a single string with a delimiter
        Return String.Join("@@@", config)
    End Function

End Class