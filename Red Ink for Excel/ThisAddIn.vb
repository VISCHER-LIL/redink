' Red Ink for Excel
' Copyright by David Rosenthal, david.rosenthal@vischer.com
' May only be used under the Red Ink License. See License.txt or https://vischer.com/redink for more information.
'
' 16.11.2025
'
' The compiled version of Red Ink also ...
'
' Includes DiffPlex in unchanged form; Copyright (c) 2023 Matthew Manela; licensed under the Appache-2.0 license (http://www.apache.org/licenses/LICENSE-2.0) at GitHub (https://github.com/mmanela/diffplex).
' Includes Newtonsoft.Json in unchanged form; Copyright (c) 2023 James Newton-King; licensed under the MIT license (https://licenses.nuget.org/MIT) at https://www.newtonsoft.com/json
' Includes HtmlAgilityPack in unchanged form; Copyright (c) 2024 ZZZ Projects, Simon Mourrier,Jeff Klawiter,Stephan Grell; licensed under the MIT license (https://licenses.nuget.org/MIT) at https://html-agility-pack.net/
' Includes Bouncycastle.Cryptography in unchanged form; Copyright (c) 2024 Legion of the Bouncy Castle Inc.; licensed under the MIT license (https://licenses.nuget.org/MIT) at https://www.bouncycastle.org/download/bouncy-castle-c/
' Includes PdfPig in unchanged form; Copyright (c) 2024 UglyToad, EliotJones PdfPig, BobLd; licensed under the Apache 2.0 license (https://licenses.nuget.org/Apache-2.0) at https://github.com/UglyToad/PdfPig
' Includes MarkDig in unchanged form; Copyright (c) 2024 Alexandre Mutel; licensed under the BSD 2 Clause (Simplified) license (https://licenses.nuget.org/BSD-2-Clause) at https://github.com/xoofx/markdig
' Includes NAudio in unchanged form; Copyright (c) 2020 Mark Heath; licensed under a proprietary open source license (https://www.nuget.org/packages/NAudio/2.2.1/license) at https://github.com/naudio/NAudio
' Includes Vosk in unchanged form; Copyright (c) 2022 Alpha Cephei Inc.; licensed under the Apache 2.0 license (https://licenses.nuget.org/Apache-2.0) at https://alphacephei.com/vosk/
' Includes Whisper.net in unchanged form; Copyright (c) 2024 Sandro Hanea; licensed under the MIT license (https://licenses.nuget.org/MIT) at https://github.com/sandrohanea/whisper.net
' Includes Grpc.core in unchanged form; Copyright (c) 2023 The gRPC Authors; licensed under the Apache 2.0 license (https://licenses.nuget.org/Apache-2.0) at https://github.com/grpc/grpc
' Includes Google Speech V1 library and related API libraries in unchanged form; Copyright (c) 2024 Google LLC; licensed under the Apache 2.0 license (https://licenses.nuget.org/Apache-2.0) at https://github.com/googleapis/google-cloud-dotnet
' Includes Google Protobuf in unchanged form; Copyright (c) 2025 Google Inc.; licensed under the BSD-3-Clause license (https://licenses.nuget.org/BSD-3-Clause) at https://github.com/protocolbuffers/protobuf
' Includes MarkdownToRTF in modified form; Copyright (c) 2025 Gustavo Hennig; original licensed under the MIT license (https://licenses.nuget.org/MIT) at https://github.com/GustavoHennig/MarkdownToRtf
' Includes Nito.AsyncEx in unchanged form; Copyright (c) 2021 Stephen Cleary; licensed under the MIT license (https://licenses.nuget.org/MIT) at https://github.com/StephenCleary/AsyncEx
' Includes NetOffice libraries in unchanged form; Copyright (c) 2020 Sebastian Lange, Erika LeBlanc; licensed under the MIT license (https://licenses.nuget.org/MIT) at https://github.com/netoffice/NetOffice-NuGet
' Includes NAudio.Lame in unchanged form; Copyright (c) 2019 Corey Murtagh; licensed under the MIT license (https://licenses.nuget.org/MIT) at https://github.com/Corey-M/NAudio.Lame
' Includes PdfiumViewer in unchanged form; Copyright (c) 2017 Pieter van Ginkel; licensed under the Apache 2.0 license (https://licenses.nuget.org/Apache-2.0) at https://github.com/pvginkel/PdfiumViewer
' Includes also various Microsoft libraries copyrighted by Microsoft Corporation and available, among others, under the Microsoft EULA and the MIT License; Copyright (c) 2016- Microsoft Corp.

Option Explicit On

Imports System.Diagnostics
Imports System.Drawing
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Web
Imports System.Windows.Forms
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Tools
Imports Microsoft.Vbe.Interop
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedContext
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports SLib = SharedLibrary.SharedLibrary.SharedMethods


Module Module1
    ' Correct attribute declaration for DllImport
    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Public Function GetAsyncKeyState(ByVal vKey As Integer) As Short
    End Function

End Module


#Region "BridgeSubs"

<ComVisible(True)>
Public Class BridgeSubs
    Public Async Function DoInLanguage1() As Task(Of Boolean)
        Dim Result As Boolean = Await Globals.ThisAddIn.InLanguage1()
    End Function

    Public Async Function DoInLanguage2() As Task
        Dim Result As Boolean = Await Globals.ThisAddIn.InLanguage2()
    End Function

    Public Async Function DoInOtherFormulas() As Task
        Dim Result As Boolean = Await Globals.ThisAddIn.InOtherFormulas()
    End Function

    Public Async Function DoCorrect() As Task
        Dim Result As Boolean = Await Globals.ThisAddIn.Correct()
    End Function

    Public Async Function DoImprove() As Task
        Dim Result As Boolean = Await Globals.ThisAddIn.Improve()
    End Function

    Public Async Function DoShorten() As Task
        Dim Result As Boolean = Await Globals.ThisAddIn.Shorten()
    End Function

    Public Async Function DoAnonymize() As Task
        Dim Result As Boolean = Await Globals.ThisAddIn.Anonymize()
    End Function

    Public Async Function DoSwitchParty() As Task
        Dim Result As Boolean = Await Globals.ThisAddIn.SwitchParty()
    End Function

    Public Async Function DoFreestyleNM() As Task
        Dim Result As Boolean = Await Globals.ThisAddIn.FreestyleNM()
    End Function

    Public Async Function DoFreestyleAM() As Task
        Dim Result As Boolean = Await Globals.ThisAddIn.FreestyleAM()
    End Function

    Public Sub DoAdjustHeight(Optional Silent As Boolean = False)
        Globals.ThisAddIn.AdjustHeight(Silent)
    End Sub
    Public Sub DoRegexSearchReplace()
        Globals.ThisAddIn.RegexSearchReplace()
    End Sub

    Public Sub DoAdjustLegacyNotes()
        Globals.ThisAddIn.AdjustLegacyNotes()
    End Sub

    Public Sub DoAddContextMenu()
        Globals.ThisAddIn.AddContextMenu()
    End Sub

    Public Function GetLLMConfig(UseSecondAPI As Boolean) As String
        Dim Result As String = Globals.ThisAddIn.GetAPIConfiguration(UseSecondAPI)
        Return Result
    End Function

    Public Function SignJWT(jwtUnsigned As String, privateKeyPem As String) As String
        Return SLib.SignJWT(jwtUnsigned, privateKeyPem)
    End Function

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

#End Region

Public Class ThisAddIn

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function FindWindow(
                                ByVal lpClassName As String,
                                ByVal lpWindowName As String
                            ) As IntPtr
    End Function

    Private mainThreadControl As New System.Windows.Forms.Control()

    Private Sub ThisAddIn_Startup() Handles Me.Startup

        ' Necessary for Update Handler to work correctly

        ' 1) Force the creation of the Control's handle on the Office UI thread
        Dim dummy = mainThreadControl.Handle

        ' 2) Give that Control to the UpdateHandler so it can Invoke on it
        UpdateHandler.MainControl = mainThreadControl

        ' 3) Capture the host window’s HWND (Word / Excel / Outlook)
        Dim hwnd As IntPtr
        Dim progId = Me.Application.GetType().Name.ToLowerInvariant()
        If progId.Contains("word") OrElse progId.Contains("excel") Then
            hwnd = New IntPtr(CInt(Me.Application.Hwnd))
        Else
            hwnd = FindWindow("rctrl_renwnd32", Nothing)
        End If

        ' Other startup tasks

        SharedMethods.Initialize(Me.CustomTaskPanes)
        InitializeAddInFeatures()
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
        RemoveOldContextMenu()
    End Sub

    Private Shared ReadOnly _uiContext As SynchronizationContext = WindowsFormsSynchronizationContext.Current

    Private Shared Async Function EnsureUIThread() As System.Threading.Tasks.Task
        If _uiContext IsNot Nothing AndAlso SynchronizationContext.Current IsNot _uiContext Then
            Await System.Threading.Tasks.Task.Factory.StartNew(
                Sub()
                    ' Continue processing here with Word objects
                End Sub,
                CancellationToken.None,
                TaskCreationOptions.None,
                TaskScheduler.FromCurrentSynchronizationContext())
        End If
    End Function

    ' Hardcoded config values

    Public Const Version As String = "V.161125 Gen2 Beta Test"

    Public Const AN As String = "Red Ink"
    Public Const AN2 As String = "redink"
    Public Const AN5 As String = "RI"

    Private Const ShortenPercent As Integer = 20
    Private Const TextPrefix As String = "TextOnly:"
    Private Const TextPrefix2 As String = "Text:"
    Private Const BatchPrefix As String = "Batch:"
    Private Const CellByCellPrefix As String = "CellByCell:"
    Private Const CellByCellPrefix2 As String = "CBC:"
    Private Const PurePrefix As String = "Pure:"
    Private Const PanePrefix As String = "Pane:"
    Private Const BubblesPrefix As String = "Bubbles:"
    Private Const ExtTrigger As String = "{doc}"
    Private Const ExtWSTrigger As String = "(addws)"
    Private Const ObjectTrigger As String = "(file)"
    Private Const ObjectTrigger2 As String = "(clip)"
    Private Const ColorTrigger As String = "(color)"
    Private Const RIMenu = AN
    Private Const MinHelperVersion = 1           ' Minimum version of the helper file that is required
    Public Const LargeWorksheetSize As Integer = 2500

    Public Shared DragDropFormLabel As String = ""
    Public Shared DragDropFormFilter As String = ""

    Public Shared allowedExtensions As System.Collections.Generic.HashSet(Of System.String) =
                        New System.Collections.Generic.HashSet(Of System.String)(
                            New System.String() {".txt", ".ini", ".csv", ".log", ".json", ".xml", ".html", ".htm", ".rtf", ".doc", ".docx", ".pdf"},
                            System.StringComparer.OrdinalIgnoreCase
                        )

    ' Definition of the SharedProperties for context for exchanging values with the SharedLibrary

    Private chatForm As frmAIChat

#Region "SharedProperties"

    Private Shared _context As ISharedContext = New SharedContext()

    Public Shared Property INI_APIKey As String
        Get
            Return _context.INI_APIKey
        End Get
        Set(value As String)
            _context.INI_APIKey = value
        End Set
    End Property

    Public Shared Property INI_APIKeyBack As String
        Get
            Return _context.INI_APIKeyBack
        End Get
        Set(value As String)
            _context.INI_APIKeyBack = value
        End Set
    End Property

    Public Shared Property INI_Temperature As String
        Get
            Return _context.INI_Temperature
        End Get
        Set(value As String)
            _context.INI_Temperature = value
        End Set
    End Property

    Public Shared Property INI_Timeout As Long
        Get
            Return _context.INI_Timeout
        End Get
        Set(value As Long)
            _context.INI_Timeout = value
        End Set
    End Property

    Public Shared Property INI_MaxOutputToken As Integer
        Get
            Return _context.INI_MaxOutputToken
        End Get
        Set(value As Integer)
            _context.INI_MaxOutputToken = value
        End Set
    End Property

    Public Shared Property INI_Model As String
        Get
            Return _context.INI_Model
        End Get
        Set(value As String)
            _context.INI_Model = value
        End Set
    End Property

    Public Shared Property INI_Endpoint As String
        Get
            Return _context.INI_Endpoint
        End Get
        Set(value As String)
            _context.INI_Endpoint = value
        End Set
    End Property

    Public Shared Property INI_HeaderA As String
        Get
            Return _context.INI_HeaderA
        End Get
        Set(value As String)
            _context.INI_HeaderA = value
        End Set
    End Property

    Public Shared Property INI_HeaderB As String
        Get
            Return _context.INI_HeaderB
        End Get
        Set(value As String)
            _context.INI_HeaderB = value
        End Set
    End Property

    Public Shared Property INI_APICall As String
        Get
            Return _context.INI_APICall
        End Get
        Set(value As String)
            _context.INI_APICall = value
        End Set
    End Property

    Public Shared Property INI_APICall_Object As String
        Get
            Return _context.INI_APICall_Object
        End Get
        Set(value As String)
            _context.INI_APICall_Object = value
        End Set
    End Property


    Public Shared Property INI_Response As String
        Get
            Return _context.INI_Response
        End Get
        Set(value As String)
            _context.INI_Response = value
        End Set
    End Property

    Public Shared Property INI_Anon As String
        Get
            Return _context.INI_Anon
        End Get
        Set(value As String)
            _context.INI_Anon = value
        End Set
    End Property

    Public Shared Property INI_TokenCount As String
        Get
            Return _context.INI_TokenCount
        End Get
        Set(value As String)
            _context.INI_TokenCount = value
        End Set
    End Property

    Public Shared Property INI_TokenCount_2 As String
        Get
            Return _context.INI_TokenCount_2
        End Get
        Set(value As String)
            _context.INI_TokenCount_2 = value
        End Set
    End Property


    Public Shared Property INI_DoubleS As Boolean
        Get
            Return _context.INI_DoubleS
        End Get
        Set(value As Boolean)
            _context.INI_DoubleS = value
        End Set
    End Property

    Public Shared Property INI_Clean As Boolean
        Get
            Return _context.INI_Clean
        End Get
        Set(value As Boolean)
            _context.INI_Clean = value
        End Set
    End Property

    Public Shared Property INI_NoDash As Boolean
        Get
            Return _context.INI_NoDash
        End Get
        Set(value As Boolean)
            _context.INI_NoDash = value
        End Set
    End Property

    Public Shared Property INI_DefaultPrefix As String
        Get
            Return _context.INI_DefaultPrefix
        End Get
        Set(value As String)
            _context.INI_DefaultPrefix = value
        End Set
    End Property

    Public Shared Property INI_MarkdownBubbles As Boolean
        Get
            Return _context.INI_MarkdownBubbles
        End Get
        Set(value As Boolean)
            _context.INI_MarkdownBubbles = value
        End Set
    End Property

    Public Shared Property INI_PreCorrection As String
        Get
            Return _context.INI_PreCorrection
        End Get
        Set(value As String)
            _context.INI_PreCorrection = value
        End Set
    End Property

    Public Shared Property INI_PostCorrection As String
        Get
            Return _context.INI_PostCorrection
        End Get
        Set(value As String)
            _context.INI_PostCorrection = value
        End Set
    End Property

    Public Shared Property INI_APIEncrypted As Boolean
        Get
            Return _context.INI_APIEncrypted
        End Get
        Set(value As Boolean)
            _context.INI_APIEncrypted = value
        End Set
    End Property

    Public Shared Property INI_APIKeyPrefix As String
        Get
            Return _context.INI_APIKeyPrefix
        End Get
        Set(value As String)
            _context.INI_APIKeyPrefix = value
        End Set
    End Property

    Public Shared Property INI_MarkupMethodOutlook As Integer
        Get
            Return _context.INI_MarkupMethodOutlook
        End Get
        Set(value As Integer)
            _context.INI_MarkupMethodOutlook = value
        End Set
    End Property

    Public Shared Property INI_MarkupDiffCap As Integer
        Get
            Return _context.INI_MarkupDiffCap
        End Get
        Set(value As Integer)
            _context.INI_MarkupDiffCap = value
        End Set
    End Property

    Public Shared Property INI_MarkupRegexCap As Integer
        Get
            Return _context.INI_MarkupRegexCap
        End Get
        Set(value As Integer)
            _context.INI_MarkupRegexCap = value
        End Set
    End Property

    Public Shared Property INI_OpenSSLPath As String
        Get
            Return _context.INI_OpenSSLPath
        End Get
        Set(value As String)
            _context.INI_OpenSSLPath = value
        End Set
    End Property


    Public Shared Property INI_OAuth2 As Boolean
        Get
            Return _context.INI_OAuth2
        End Get
        Set(value As Boolean)
            _context.INI_OAuth2 = value
        End Set
    End Property

    Public Shared Property INI_OAuth2ClientMail As String
        Get
            Return _context.INI_OAuth2ClientMail
        End Get
        Set(value As String)
            _context.INI_OAuth2ClientMail = value
        End Set
    End Property

    Public Shared Property INI_OAuth2Scopes As String
        Get
            Return _context.INI_OAuth2Scopes
        End Get
        Set(value As String)
            _context.INI_OAuth2Scopes = value
        End Set
    End Property

    Public Shared Property INI_OAuth2Endpoint As String
        Get
            Return _context.INI_OAuth2Endpoint
        End Get
        Set(value As String)
            _context.INI_OAuth2Endpoint = value
        End Set
    End Property

    Public Shared Property INI_OAuth2ATExpiry As Long
        Get
            Return _context.INI_OAuth2ATExpiry
        End Get
        Set(value As Long)
            _context.INI_OAuth2ATExpiry = value
        End Set
    End Property

    Public Shared Property INI_SecondAPI As Boolean
        Get
            Return _context.INI_SecondAPI
        End Get
        Set(value As Boolean)
            _context.INI_SecondAPI = value
        End Set
    End Property

    Public Shared Property INI_APIKey_2 As String
        Get
            Return _context.INI_APIKey_2
        End Get
        Set(value As String)
            _context.INI_APIKey_2 = value
        End Set
    End Property

    Public Shared Property INI_APIKeyBack_2 As String
        Get
            Return _context.INI_APIKeyBack_2
        End Get
        Set(value As String)
            _context.INI_APIKeyBack_2 = value
        End Set
    End Property

    Public Shared Property INI_Temperature_2 As String
        Get
            Return _context.INI_Temperature_2
        End Get
        Set(value As String)
            _context.INI_Temperature_2 = value
        End Set
    End Property

    Public Shared Property INI_Timeout_2 As Long
        Get
            Return _context.INI_Timeout_2
        End Get
        Set(value As Long)
            _context.INI_Timeout_2 = value
        End Set
    End Property
    Public Shared Property INI_MaxOutputToken_2 As Integer
        Get
            Return _context.INI_MaxOutputToken_2
        End Get
        Set(value As Integer)
            _context.INI_MaxOutputToken_2 = value
        End Set
    End Property

    Public Shared Property INI_Model_2 As String
        Get
            Return _context.INI_Model_2
        End Get
        Set(value As String)
            _context.INI_Model_2 = value
        End Set
    End Property

    Public Shared Property INI_Endpoint_2 As String
        Get
            Return _context.INI_Endpoint_2
        End Get
        Set(value As String)
            _context.INI_Endpoint_2 = value
        End Set
    End Property

    Public Shared Property INI_HeaderA_2 As String
        Get
            Return _context.INI_HeaderA_2
        End Get
        Set(value As String)
            _context.INI_HeaderA_2 = value
        End Set
    End Property

    Public Shared Property INI_HeaderB_2 As String
        Get
            Return _context.INI_HeaderB_2
        End Get
        Set(value As String)
            _context.INI_HeaderB_2 = value
        End Set
    End Property

    Public Shared Property INI_APICall_2 As String
        Get
            Return _context.INI_APICall_2
        End Get
        Set(value As String)
            _context.INI_APICall_2 = value
        End Set
    End Property

    Public Shared Property INI_APICall_Object_2 As String
        Get
            Return _context.INI_APICall_Object_2
        End Get
        Set(value As String)
            _context.INI_APICall_Object_2 = value
        End Set
    End Property


    Public Shared Property INI_Response_2 As String
        Get
            Return _context.INI_Response_2
        End Get
        Set(value As String)
            _context.INI_Response_2 = value
        End Set
    End Property

    Public Shared Property INI_Anon_2 As String
        Get
            Return _context.INI_Anon_2
        End Get
        Set(value As String)
            _context.INI_Anon_2 = value
        End Set
    End Property

    Public Shared Property INI_APIEncrypted_2 As Boolean
        Get
            Return _context.INI_APIEncrypted_2
        End Get
        Set(value As Boolean)
            _context.INI_APIEncrypted_2 = value
        End Set
    End Property

    Public Shared Property INI_APIKeyPrefix_2 As String
        Get
            Return _context.INI_APIKeyPrefix_2
        End Get
        Set(value As String)
            _context.INI_APIKeyPrefix_2 = value
        End Set
    End Property

    Public Shared Property INI_OAuth2_2 As Boolean
        Get
            Return _context.INI_OAuth2_2
        End Get
        Set(value As Boolean)
            _context.INI_OAuth2_2 = value
        End Set
    End Property

    Public Shared Property INI_OAuth2ClientMail_2 As String
        Get
            Return _context.INI_OAuth2ClientMail_2
        End Get
        Set(value As String)
            _context.INI_OAuth2ClientMail_2 = value
        End Set
    End Property

    Public Shared Property INI_OAuth2Scopes_2 As String
        Get
            Return _context.INI_OAuth2Scopes_2
        End Get
        Set(value As String)
            _context.INI_OAuth2Scopes_2 = value
        End Set
    End Property

    Public Shared Property INI_OAuth2Endpoint_2 As String
        Get
            Return _context.INI_OAuth2Endpoint_2
        End Get
        Set(value As String)
            _context.INI_OAuth2Endpoint_2 = value
        End Set
    End Property

    Public Shared Property INI_OAuth2ATExpiry_2 As Long
        Get
            Return _context.INI_OAuth2ATExpiry_2
        End Get
        Set(value As Long)
            _context.INI_OAuth2ATExpiry_2 = value
        End Set
    End Property

    Public Shared Property INI_APIDebug As Boolean
        Get
            Return _context.INI_APIDebug
        End Get
        Set(value As Boolean)
            _context.INI_APIDebug = value
        End Set
    End Property

    Public Shared Property INI_UsageRestrictions As String
        Get
            Return _context.INI_UsageRestrictions
        End Get
        Set(value As String)
            _context.INI_UsageRestrictions = value
        End Set
    End Property

    Public Shared Property INI_Language1 As String
        Get
            Return _context.INI_Language1
        End Get
        Set(value As String)
            _context.INI_Language1 = value
        End Set
    End Property

    Public Shared Property INI_Language2 As String
        Get
            Return _context.INI_Language2
        End Get
        Set(value As String)
            _context.INI_Language2 = value
        End Set
    End Property

    Public Shared Property INI_KeepFormat1 As Boolean
        Get
            Return _context.INI_KeepFormat1
        End Get
        Set(value As Boolean)
            _context.INI_KeepFormat1 = value
        End Set
    End Property

    Public Shared Property INI_KeepFormat2 As Boolean
        Get
            Return _context.INI_KeepFormat2
        End Get
        Set(value As Boolean)
            _context.INI_KeepFormat2 = value
        End Set
    End Property

    Public Shared Property INI_KeepParaFormatInline As Boolean
        Get
            Return _context.INI_KeepParaFormatInline
        End Get
        Set(value As Boolean)
            _context.INI_KeepParaFormatInline = value
        End Set
    End Property

    Public Shared Property INI_KeepFormatCap As Integer
        Get
            Return _context.INI_KeepFormatCap
        End Get
        Set(value As Integer)
            _context.INI_KeepFormatCap = value
        End Set
    End Property


    Public Shared Property INI_ReplaceText1 As Boolean
        Get
            Return _context.INI_ReplaceText1
        End Get
        Set(value As Boolean)
            _context.INI_ReplaceText1 = value
        End Set
    End Property

    Public Shared Property INI_ReplaceText2 As Boolean
        Get
            Return _context.INI_ReplaceText2
        End Get
        Set(value As Boolean)
            _context.INI_ReplaceText2 = value
        End Set
    End Property

    Public Shared Property INI_ReplaceText2Override As String
        Get
            Return _context.INI_ReplaceText2Override
        End Get
        Set(value As String)
            _context.INI_ReplaceText2Override = value
        End Set
    End Property


    Public Shared Property INI_DoMarkupOutlook As Boolean
        Get
            Return _context.INI_DoMarkupOutlook
        End Get
        Set(value As Boolean)
            _context.INI_DoMarkupOutlook = value
        End Set
    End Property

    Public Shared Property INI_DoMarkupWord As Boolean
        Get
            Return _context.INI_DoMarkupWord
        End Get
        Set(value As Boolean)
            _context.INI_DoMarkupWord = value
        End Set
    End Property

    Public Shared Property SP_Translate As String
        Get
            Return _context.SP_Translate
        End Get
        Set(value As String)
            _context.SP_Translate = value
        End Set
    End Property

    Public Shared Property SP_Correct As String
        Get
            Return _context.SP_Correct
        End Get
        Set(value As String)
            _context.SP_Correct = value
        End Set
    End Property

    Public Shared Property SP_Improve As String
        Get
            Return _context.SP_Improve
        End Get
        Set(value As String)
            _context.SP_Improve = value
        End Set
    End Property

    Public Shared Property SP_Explain As String
        Get
            Return _context.SP_Explain
        End Get
        Set(value As String)
            _context.SP_Explain = value
        End Set
    End Property

    Public Shared Property SP_FindClause As String
        Get
            Return _context.SP_FindClause
        End Get
        Set(value As String)
            _context.SP_FindClause = value
        End Set
    End Property

    Public Shared Property SP_FindClause_Clean As String
        Get
            Return _context.SP_FindClause_Clean
        End Get
        Set(value As String)
            _context.SP_FindClause_Clean = value
        End Set
    End Property
    Public Shared Property SP_DocCheck_Clause As String
        Get
            Return _context.SP_DocCheck_Clause
        End Get
        Set(value As String)
            _context.SP_DocCheck_Clause = value
        End Set
    End Property

    Public Shared Property SP_DocCheck_MultiClause As String
        Get
            Return _context.SP_DocCheck_MultiClause
        End Get
        Set(value As String)
            _context.SP_DocCheck_MultiClause = value
        End Set
    End Property

    Public Shared Property SP_DocCheck_MultiClauseSum As String
        Get
            Return _context.SP_DocCheck_MultiClauseSum
        End Get
        Set(value As String)
            _context.SP_DocCheck_MultiClauseSum = value
        End Set
    End Property

    Public Shared Property SP_DocCheck_MultiClauseSum_Bubbles As String
        Get
            Return _context.SP_DocCheck_MultiClauseSum_Bubbles
        End Get
        Set(value As String)
            _context.SP_DocCheck_MultiClauseSum_Bubbles = value
        End Set
    End Property

    Public Shared Property SP_SuggestTitles As String
        Get
            Return _context.SP_SuggestTitles
        End Get
        Set(value As String)
            _context.SP_SuggestTitles = value
        End Set
    End Property

    Public Shared Property SP_Friendly As String
        Get
            Return _context.SP_Friendly
        End Get
        Set(value As String)
            _context.SP_Friendly = value
        End Set
    End Property

    Public Shared Property SP_Convincing As String
        Get
            Return _context.SP_Convincing
        End Get
        Set(value As String)
            _context.SP_Convincing = value
        End Set
    End Property

    Public Shared Property SP_NoFillers As String
        Get
            Return _context.SP_NoFillers
        End Get
        Set(value As String)
            _context.SP_NoFillers = value
        End Set
    End Property

    Public Shared Property SP_Podcast As String
        Get
            Return _context.SP_Podcast
        End Get
        Set(value As String)
            _context.SP_Podcast = value
        End Set
    End Property


    Public Shared Property SP_MyStyle_Word As String
        Get
            Return _context.SP_MyStyle_Word
        End Get
        Set(value As String)
            _context.SP_MyStyle_Word = value
        End Set
    End Property

    Public Shared Property SP_MyStyle_Outlook As String
        Get
            Return _context.SP_MyStyle_Outlook
        End Get
        Set(value As String)
            _context.SP_MyStyle_Outlook = value
        End Set
    End Property

    Public Shared Property SP_MyStyle_Apply As String
        Get
            Return _context.SP_MyStyle_Apply
        End Get
        Set(value As String)
            _context.SP_MyStyle_Apply = value
        End Set
    End Property

    Public Shared Property SP_Shorten As String
        Get
            Return _context.SP_Shorten
        End Get
        Set(value As String)
            _context.SP_Shorten = value
        End Set
    End Property

    Public Shared Property SP_Summarize As String
        Get
            Return _context.SP_Summarize
        End Get
        Set(value As String)
            _context.SP_Summarize = value
        End Set
    End Property

    Public Shared Property SP_MailReply As String
        Get
            Return _context.SP_MailReply
        End Get
        Set(value As String)
            _context.SP_MailReply = value
        End Set
    End Property

    Public Shared Property SP_MailSumup As String
        Get
            Return _context.SP_MailSumup
        End Get
        Set(value As String)
            _context.SP_MailSumup = value
        End Set
    End Property

    Public Shared Property SP_MailSumup2 As String
        Get
            Return _context.SP_MailSumup2
        End Get
        Set(value As String)
            _context.SP_MailSumup2 = value
        End Set
    End Property

    Public Shared Property SP_FreestyleText As String
        Get
            Return _context.SP_FreestyleText
        End Get
        Set(value As String)
            _context.SP_FreestyleText = value
        End Set
    End Property

    Public Shared Property SP_FreestyleNoText As String
        Get
            Return _context.SP_FreestyleNoText
        End Get
        Set(value As String)
            _context.SP_FreestyleNoText = value
        End Set
    End Property

    Public Shared Property SP_SwitchParty As String
        Get
            Return _context.SP_SwitchParty
        End Get
        Set(value As String)
            _context.SP_SwitchParty = value
        End Set
    End Property
    Public Shared Property SP_Anonymize As String
        Get
            Return _context.SP_Anonymize
        End Get
        Set(value As String)
            _context.SP_Anonymize = value
        End Set
    End Property

    Public Shared Property SP_Redact As String
        Get
            Return _context.SP_Redact
        End Get
        Set(value As String)
            _context.SP_Redact = value
        End Set
    End Property

    Public Shared Property SP_CheckforII As String
        Get
            Return _context.SP_CheckforII
        End Get
        Set(value As String)
            _context.SP_CheckforII = value
        End Set
    End Property


    Public Shared Property SP_ContextSearch As String
        Get
            Return _context.SP_ContextSearch
        End Get
        Set(value As String)
            _context.SP_ContextSearch = value
        End Set
    End Property

    Public Shared Property SP_ContextSearchMulti As String
        Get
            Return _context.SP_ContextSearchMulti
        End Get
        Set(value As String)
            _context.SP_ContextSearchMulti = value
        End Set
    End Property


    Public Shared Property SP_RangeOfCells As String
        Get
            Return _context.SP_RangeOfCells
        End Get
        Set(value As String)
            _context.SP_RangeOfCells = value
        End Set
    End Property

    Public Shared Property SP_ParseFile As String
        Get
            Return _context.SP_ParseFile
        End Get
        Set(value As String)
            _context.SP_ParseFile = value
        End Set
    End Property


    Public Shared Property SP_WriteNeatly As String
        Get
            Return _context.SP_WriteNeatly
        End Get
        Set(value As String)
            _context.SP_WriteNeatly = value
        End Set
    End Property

    Public Shared Property SP_Add_KeepFormulasIntact As String
        Get
            Return _context.SP_Add_KeepFormulasIntact
        End Get
        Set(value As String)
            _context.SP_Add_KeepFormulasIntact = value
        End Set
    End Property
    Public Shared Property SP_Add_KeepHTMLIntact As String
        Get
            Return _context.SP_Add_KeepHTMLIntact
        End Get
        Set(value As String)
            _context.SP_Add_KeepHTMLIntact = value
        End Set
    End Property

    Public Shared Property SP_Add_KeepInlineIntact As String
        Get
            Return _context.SP_Add_KeepInlineIntact
        End Get
        Set(value As String)
            _context.SP_Add_KeepInlineIntact = value
        End Set
    End Property

    Public Shared Property SP_BubblesExcel As String
        Get
            Return _context.SP_BubblesExcel
        End Get
        Set(value As String)
            _context.SP_BubblesExcel = value
        End Set
    End Property

    Public Shared Property SP_Add_Bubbles As String
        Get
            Return _context.SP_Add_Bubbles
        End Get
        Set(value As String)
            _context.SP_Add_Bubbles = value
        End Set
    End Property

    Public Shared Property SP_Add_BubblesReply As String
        Get
            Return _context.SP_Add_BubblesReply
        End Get
        Set(value As String)
            _context.SP_Add_BubblesReply = value
        End Set
    End Property

    Public Shared Property SP_Add_BubblesExtract As String
        Get
            Return _context.SP_Add_BubblesExtract
        End Get
        Set(value As String)
            _context.SP_Add_BubblesExtract = value
        End Set
    End Property

    Public Shared Property SP_Add_Bubbles_Format As String
        Get
            Return _context.SP_Add_Bubbles_Format
        End Get
        Set(value As String)
            _context.SP_Add_Bubbles_Format = value
        End Set
    End Property

    Public Shared Property SP_Add_Batch As String
        Get
            Return _context.SP_Add_Batch
        End Get
        Set(value As String)
            _context.SP_Add_Batch = value
        End Set
    End Property


    Public Shared Property SP_Add_Revisions As String
        Get
            Return _context.SP_Add_Revisions
        End Get
        Set(value As String)
            _context.SP_Add_Revisions = value
        End Set
    End Property

    Public Shared Property SP_MarkupRegex As String
        Get
            Return _context.SP_MarkupRegex
        End Get
        Set(value As String)
            _context.SP_MarkupRegex = value
        End Set
    End Property


    Public Shared Property SP_ChatWord As String
        Get
            Return _context.SP_ChatWord
        End Get
        Set(value As String)
            _context.SP_ChatWord = value
        End Set
    End Property

    Public Shared Property SP_Chat As String
        Get
            Return _context.SP_Chat
        End Get
        Set(value As String)
            _context.SP_Chat = value
        End Set
    End Property

    Public Shared Property SP_HelpMe As String
        Get
            Return _context.SP_HelpMe
        End Get
        Set(value As String)
            _context.SP_HelpMe = value
        End Set
    End Property

    Public Shared Property SP_Add_ChatWord_Commands As String
        Get
            Return _context.SP_Add_ChatWord_Commands
        End Get
        Set(value As String)
            _context.SP_Add_ChatWord_Commands = value
        End Set
    End Property

    Public Shared Property SP_Add_Chat_NoCommands As String
        Get
            Return _context.SP_Add_Chat_NoCommands
        End Get
        Set(value As String)
            _context.SP_Add_Chat_NoCommands = value
        End Set
    End Property


    Public Shared Property SP_ChatExcel As String
        Get
            Return _context.SP_ChatExcel
        End Get
        Set(value As String)
            _context.SP_ChatExcel = value
        End Set
    End Property

    Public Shared Property SP_Add_ChatExcel_Commands As String
        Get
            Return _context.SP_Add_ChatExcel_Commands
        End Get
        Set(value As String)
            _context.SP_Add_ChatExcel_Commands = value
        End Set
    End Property

    Public Shared Property INI_ChatCap As Integer
        Get
            Return _context.INI_ChatCap
        End Get
        Set(value As Integer)
            _context.INI_ChatCap = value
        End Set
    End Property

    Public Shared ReadOnly Property RDV As String = "Excel (" & Version & ")"
    Public Shared Property DecodedAPI As String
        Get
            Return _context.DecodedAPI
        End Get
        Set(value As String)
            _context.DecodedAPI = value
        End Set
    End Property

    Public Shared Property DecodedAPI_2 As String
        Get
            Return _context.DecodedAPI_2
        End Get
        Set(value As String)
            _context.DecodedAPI_2 = value
        End Set
    End Property

    Public Shared Property TokenExpiry As DateTime
        Get
            Return _context.TokenExpiry
        End Get
        Set(value As DateTime)
            _context.TokenExpiry = value
        End Set
    End Property

    Public Shared Property TokenExpiry_2 As DateTime
        Get
            Return _context.TokenExpiry_2
        End Get
        Set(value As DateTime)
            _context.TokenExpiry_2 = value
        End Set
    End Property

    Public Shared Property Codebasis As String
        Get
            Return _context.Codebasis
        End Get
        Set(value As String)
            _context.Codebasis = value
        End Set
    End Property

    Public Shared Property GPTSetupError As Boolean
        Get
            Return _context.GPTSetupError
        End Get
        Set(value As Boolean)
            _context.GPTSetupError = value
        End Set
    End Property

    Public Shared Property INIloaded As Boolean
        Get
            Return _context.INIloaded
        End Get
        Set(value As Boolean)
            _context.INIloaded = value
        End Set
    End Property



    Public Shared Property INI_ISearch As Boolean
        Get
            Return _context.INI_ISearch
        End Get
        Set(value As Boolean)
            _context.INI_ISearch = value
        End Set
    End Property

    Public Shared Property INI_ISearch_Approve As Boolean
        Get
            Return _context.INI_ISearch_Approve
        End Get
        Set(value As Boolean)
            _context.INI_ISearch_Approve = value
        End Set
    End Property

    Public Shared Property INI_ISearch_URL As String
        Get
            Return _context.INI_ISearch_URL
        End Get
        Set(value As String)
            _context.INI_ISearch_URL = value
        End Set
    End Property

    Public Shared Property INI_ISearch_ResponseURLStart As String
        Get
            Return _context.INI_ISearch_ResponseURLStart
        End Get
        Set(value As String)
            _context.INI_ISearch_ResponseURLStart = value
        End Set
    End Property

    Public Shared Property INI_ISearch_ResponseMask1 As String
        Get
            Return _context.INI_ISearch_ResponseMask1
        End Get
        Set(value As String)
            _context.INI_ISearch_ResponseMask1 = value
        End Set
    End Property

    Public Shared Property INI_ISearch_ResponseMask2 As String
        Get
            Return _context.INI_ISearch_ResponseMask2
        End Get
        Set(value As String)
            _context.INI_ISearch_ResponseMask2 = value
        End Set
    End Property

    Public Shared Property INI_ISearch_Name As String
        Get
            Return _context.INI_ISearch_Name
        End Get
        Set(value As String)
            _context.INI_ISearch_Name = value
        End Set
    End Property

    Public Shared Property INI_ISearch_Tries As Integer
        Get
            Return _context.INI_ISearch_Tries
        End Get
        Set(value As Integer)
            _context.INI_ISearch_Tries = value
        End Set
    End Property

    Public Shared Property INI_ISearch_Results As Integer
        Get
            Return _context.INI_ISearch_Results
        End Get
        Set(value As Integer)
            _context.INI_ISearch_Results = value
        End Set
    End Property

    Public Shared Property INI_ISearch_MaxDepth As Integer
        Get
            Return _context.INI_ISearch_MaxDepth
        End Get
        Set(value As Integer)
            _context.INI_ISearch_MaxDepth = value
        End Set
    End Property

    Public Shared Property INI_ISearch_Timeout As Long
        Get
            Return _context.INI_ISearch_Timeout
        End Get
        Set(value As Long)
            _context.INI_ISearch_Timeout = value
        End Set
    End Property

    Public Shared Property INI_ISearch_SearchTerm_SP As String
        Get
            Return _context.INI_ISearch_SearchTerm_SP
        End Get
        Set(value As String)
            _context.INI_ISearch_SearchTerm_SP = value
        End Set
    End Property

    Public Shared Property INI_ISearch_Apply_SP_Markup As String
        Get
            Return _context.INI_ISearch_Apply_SP_Markup
        End Get
        Set(value As String)
            _context.INI_ISearch_Apply_SP_Markup = value
        End Set
    End Property
    Public Shared Property INI_ISearch_Apply_SP As String
        Get
            Return _context.INI_ISearch_Apply_SP
        End Get
        Set(value As String)
            _context.INI_ISearch_Apply_SP = value
        End Set
    End Property

    Public Shared Property INI_Lib As Boolean
        Get
            Return _context.INI_Lib
        End Get
        Set(value As Boolean)
            _context.INI_Lib = value
        End Set
    End Property

    Public Shared Property INI_Lib_File As String
        Get
            Return _context.INI_Lib_File
        End Get
        Set(value As String)
            _context.INI_Lib_File = value
        End Set
    End Property

    Public Shared Property INI_Lib_Timeout As Long
        Get
            Return _context.INI_Lib_Timeout
        End Get
        Set(value As Long)
            _context.INI_Lib_Timeout = value
        End Set
    End Property

    Public Shared Property INI_Lib_Find_SP As String
        Get
            Return _context.INI_Lib_Find_SP
        End Get
        Set(value As String)
            _context.INI_Lib_Find_SP = value
        End Set
    End Property

    Public Shared Property INI_Lib_Apply_SP_Markup As String
        Get
            Return _context.INI_Lib_Apply_SP_Markup
        End Get
        Set(value As String)
            _context.INI_Lib_Apply_SP_Markup = value
        End Set
    End Property

    Public Shared Property INI_Lib_Apply_SP As String
        Get
            Return _context.INI_Lib_Apply_SP
        End Get
        Set(value As String)
            _context.INI_Lib_Apply_SP = value
        End Set
    End Property

    Public Shared Property INI_MarkupMethodHelper As Integer
        Get
            Return _context.INI_MarkupMethodHelper
        End Get
        Set(value As Integer)
            _context.INI_MarkupMethodHelper = value
        End Set
    End Property

    Public Shared Property INI_MarkupMethodWord As Integer
        Get
            Return _context.INI_MarkupMethodWord
        End Get
        Set(value As Integer)
            _context.INI_MarkupMethodWord = value
        End Set
    End Property

    Public Shared Property INI_MarkupMethodWordOverride As String
        Get
            Return _context.INI_MarkupMethodWordOverride
        End Get
        Set(value As String)
            _context.INI_MarkupMethodWordOverride = value
        End Set
    End Property

    Public Shared Property INI_MarkupMethodOutlookOverride As String
        Get
            Return _context.INI_MarkupMethodOutlookOverride
        End Get
        Set(value As String)
            _context.INI_MarkupMethodOutlookOverride = value
        End Set
    End Property

    Public Shared Property INI_ContextMenu As Boolean
        Get
            Return _context.INI_ContextMenu
        End Get
        Set(value As Boolean)
            _context.INI_ContextMenu = value
        End Set
    End Property

    Public Shared Property INI_UpdateCheckInterval As Integer
        Get
            Return _context.INI_UpdateCheckInterval
        End Get
        Set(value As Integer)
            _context.INI_UpdateCheckInterval = value
        End Set
    End Property
    Public Shared Property INI_UpdatePath As String
        Get
            Return _context.INI_UpdatePath
        End Get
        Set(value As String)
            _context.INI_UpdatePath = value
        End Set
    End Property

    Public Shared Property INI_HelpMeInkyPath As String
        Get
            Return _context.INI_HelpMeInkyPath
        End Get
        Set(value As String)
            _context.INI_HelpMeInkyPath = value
        End Set
    End Property

    Public Shared Property INI_RedactionInstructionsPath As String
        Get
            Return _context.INI_RedactionInstructionsPath
        End Get
        Set(value As String)
            _context.INI_RedactionInstructionsPath = value
        End Set
    End Property

    Public Shared Property INI_RedactionInstructionsPathLocal As String
        Get
            Return _context.INI_RedactionInstructionsPathLocal
        End Get
        Set(value As String)
            _context.INI_RedactionInstructionsPathLocal = value
        End Set
    End Property


    Public Shared Property INI_SpeechModelPath As String
        Get
            Return _context.INI_SpeechModelPath
        End Get
        Set(value As String)
            _context.INI_SpeechModelPath = value
        End Set
    End Property

    Public Shared Property INI_TTSEndpoint As String
        Get
            Return _context.INI_TTSEndpoint
        End Get
        Set(value As String)
            _context.INI_TTSEndpoint = value
        End Set
    End Property
    Public Shared Property INI_ShortcutsWordExcel As String
        Get
            Return _context.INI_ShortcutsWordExcel
        End Get
        Set(value As String)
            _context.INI_ShortcutsWordExcel = value
        End Set
    End Property

    Public Shared Property INI_PromptLib As Boolean
        Get
            Return _context.INI_PromptLib
        End Get
        Set(value As Boolean)
            _context.INI_PromptLib = value
        End Set
    End Property

    Public Shared Property INI_MyStylePath As String
        Get
            Return _context.INI_MyStylePath
        End Get
        Set(value As String)
            _context.INI_MyStylePath = value
        End Set
    End Property

    Public Shared Property INI_PromptLibPath As String
        Get
            Return _context.INI_PromptLibPath
        End Get
        Set(value As String)
            _context.INI_PromptLibPath = value
        End Set
    End Property

    Public Shared Property INI_PromptLibPathLocal As String
        Get
            Return _context.INI_PromptLibPathLocal
        End Get
        Set(value As String)
            _context.INI_PromptLibPathLocal = value
        End Set
    End Property


    Public Shared Property INI_PromptLibPath_Transcript As String
        Get
            Return _context.INI_PromptLibPath_Transcript
        End Get
        Set(value As String)
            _context.INI_PromptLibPath_Transcript = value
        End Set
    End Property

    Public Shared Property INI_AlternateModelPath As String
        Get
            Return _context.INI_AlternateModelPath
        End Get
        Set(value As String)
            _context.INI_AlternateModelPath = value
        End Set
    End Property

    Public Shared Property INI_SpecialServicePath As String
        Get
            Return _context.INI_SpecialServicePath
        End Get
        Set(value As String)
            _context.INI_SpecialServicePath = value
        End Set
    End Property

    Public Shared Property INI_FindClausePath As String
        Get
            Return _context.INI_FindClausePath
        End Get
        Set(value As String)
            _context.INI_FindClausePath = value
        End Set
    End Property

    Public Shared Property INI_FindClausePathLocal As String
        Get
            Return _context.INI_FindClausePathLocal
        End Get
        Set(value As String)
            _context.INI_FindClausePathLocal = value
        End Set
    End Property

    Public Shared Property INI_WebAgentPath As String
        Get
            Return _context.INI_WebAgentPath
        End Get
        Set(value As String)
            _context.INI_WebAgentPath = value
        End Set
    End Property

    Public Shared Property INI_WebAgentPathLocal As String
        Get
            Return _context.INI_WebAgentPathLocal
        End Get
        Set(value As String)
            _context.INI_WebAgentPathLocal = value
        End Set
    End Property

    Public Shared Property INI_DocCheckPath As String
        Get
            Return _context.INI_DocCheckPath
        End Get
        Set(value As String)
            _context.INI_DocCheckPath = value
        End Set
    End Property

    Public Shared Property INI_DocCheckPathLocal As String
        Get
            Return _context.INI_DocCheckPathLocal
        End Get
        Set(value As String)
            _context.INI_DocCheckPathLocal = value
        End Set
    End Property

    Public Shared Property PromptLibrary() As List(Of String)
        Get
            Return _context.PromptLibrary
        End Get
        Set(value As List(Of String))
            _context.PromptLibrary = value
        End Set
    End Property

    Public Shared Property PromptTitles() As List(Of String)
        Get
            Return _context.PromptTitles
        End Get
        Set(value As List(Of String))
            _context.PromptTitles = value
        End Set
    End Property

    Public Shared Property MenusAdded As Boolean
        Get
            Return _context.MenusAdded
        End Get
        Set(value As Boolean)
            _context.MenusAdded = value
        End Set
    End Property

    Public Shared Property InitialConfigFailed As Boolean
        Get
            Return _context.InitialConfigFailed
        End Get
        Set(value As Boolean)
            _context.InitialConfigFailed = value
        End Set
    End Property

    Public Shared Property INI_Model_Parameter1 As String
        Get
            Return _context.INI_Model_Parameter1
        End Get
        Set(value As String)
            _context.INI_Model_Parameter1 = value
        End Set
    End Property

    Public Shared Property INI_Model_Parameter2 As String
        Get
            Return _context.INI_Model_Parameter2
        End Get
        Set(value As String)
            _context.INI_Model_Parameter2 = value
        End Set
    End Property

    Public Shared Property INI_Model_Parameter3 As String
        Get
            Return _context.INI_Model_Parameter3
        End Get
        Set(value As String)
            _context.INI_Model_Parameter3 = value
        End Set
    End Property

    Public Shared Property INI_Model_Parameter4 As String
        Get
            Return _context.INI_Model_Parameter4
        End Get
        Set(value As String)
            _context.INI_Model_Parameter4 = value
        End Set
    End Property

    Public Shared Property SP_FindPrompts As String
        Get
            Return _context.SP_FindPrompts
        End Get
        Set(value As String)
            _context.SP_FindPrompts = value
        End Set
    End Property

    Public Shared Property SP_MergePrompt As String
        Get
            Return _context.SP_MergePrompt
        End Get
        Set(value As String)
            _context.SP_MergePrompt = value
        End Set
    End Property
    Public Shared Property SP_Add_MergePrompt As String
        Get
            Return _context.SP_Add_MergePrompt
        End Get
        Set(value As String)
            _context.SP_Add_MergePrompt = value
        End Set
    End Property


#End Region

    ' Functions of SharedLibrary

#Region "SharedLibrary"

    Public Sub InitializeConfig(FirstTime As Boolean, Reload As Boolean)
        _context.InitialConfigFailed = False
        _context.RDV = "Excel (" & Version & ")"
        SharedMethods.InitializeConfig(_context, FirstTime, Reload)
    End Sub
    Private Function INIValuesMissing()
        Return SharedMethods.INIValuesMissing(_context)
    End Function
    Public Shared Async Function PostCorrection(inputText As String, Optional ByVal UseSecondAPI As Boolean = False) As Task(Of String)
        Return Await SharedMethods.PostCorrection(_context, inputText, UseSecondAPI)
    End Function
    Public Shared Async Function LLM(ByVal promptSystem As String, ByVal promptUser As String, Optional ByVal Model As String = "", Optional ByVal Temperature As String = "", Optional ByVal Timeout As Long = 0, Optional ByVal UseSecondAPI As Boolean = False, Optional ByVal Hidesplash As Boolean = True, Optional ByVal AddUserPrompt As String = "", Optional ByVal FileObject As String = "") As Task(Of String)
        Dim Response = Await SharedMethods.LLM(_context, promptSystem, promptUser, Model, Temperature, Timeout, UseSecondAPI, Hidesplash, AddUserPrompt, FileObject)
        Await EnsureUIThread()
        Return Response
    End Function
    Private Function ShowSettingsWindow(Settings As Dictionary(Of String, String), SettingsTips As Dictionary(Of String, String))
        SharedMethods.ShowSettingsWindow(Settings, SettingsTips, _context)
    End Function
    Private Function ShowPromptSelector(filePath As String, filepathlocal As String, enableMarkup As Boolean, enableBubbles As Boolean) As (String, Boolean, Boolean, Boolean)
        Return SharedMethods.ShowPromptSelector(filePath, filepathlocal, enableMarkup, enableBubbles, _context)
    End Function

#End Region

    Private automationObject As BridgeSubs

    Protected Overrides Function RequestComAddInAutomationService() As Object
        If automationObject Is Nothing Then
            automationObject = New BridgeSubs()
        End If
        Return automationObject
    End Function

    Public Sub InitializeAddInFeatures()
        InitializeConfig(True, True)
        AddContextMenu()
        UpdateHandler.PeriodicCheckForUpdates(INI_UpdateCheckInterval, RDV, INI_UpdatePath)
    End Sub

#Region "Build Menu"
    Public Sub AddContextMenu()

        Dim result = Globals.Ribbons.Ribbon1.InitializeAppAsync()

        If MenusAdded Then Exit Sub

        If RemoveMenu Then
            RemoveOldContextMenu()
            RemoveMenu = False
        End If

        If Not INI_ContextMenu Then Exit Sub

        If Not VBAModuleWorking() Then Exit Sub

        If INIloaded = False Then Exit Sub

        MenusAdded = True

        ' List of relevant context menus
        Dim contextMenus As String() = {"Cell", "Row", "Column", "List Range Popup", "PivotTable Context Menu", "Text Box", "Drawing Object", "Chart"}
        Dim application As Excel.Application = Globals.ThisAddIn.Application

        For Each cb As CommandBar In application.CommandBars
            If cb.Type = MsoBarType.msoBarTypePopup Then
                ' Check if the context menu is relevant
                If contextMenus.Contains(cb.Name) Then
                    ' Check if the menu already exists
                    If Not ContextMenuExists(cb, RIMenu) Then
                        Dim myControl As CommandBarPopup = Nothing
                        Try
                            myControl = CType(cb.Controls.Add(Type:=MsoControlType.msoControlPopup, Temporary:=True), CommandBarPopup)
                        Catch ex As System.Exception
                            ' Handle potential errors
                        End Try

                        If myControl IsNot Nothing Then
                            myControl.Caption = RIMenu
                            myControl.Visible = True
                            myControl.Enabled = True

                            ' Add submenu items
                            AddSubMenuItems(myControl)
                        End If
                    End If
                End If
            End If
        Next

    End Sub

    Private Function ContextMenuExists(cb As CommandBar, menuName As String) As Boolean
        For Each ctrl As CommandBarControl In cb.Controls
            If ctrl.Type = MsoControlType.msoControlPopup AndAlso ctrl.Caption = menuName Then
                Return True
            End If
        Next
        Return False
    End Function

    Public Sub AddSubMenuItems(myControl As CommandBarPopup)

        Try
            Dim subControl As CommandBarButton
            Dim excelHelpersMenu As CommandBarPopup
            Dim subSubControl As CommandBarButton
            Dim shortcutsArray() As String
            Dim shortcutPair() As String
            Dim shortcutDict As New Dictionary(Of String, String) ' Use native .NET Dictionary
            Dim i As Integer

            ' Parse the shortcuts from INI_ShortcutsWordExcel
            shortcutsArray = INI_ShortcutsWordExcel.Split(";")

            ' Populate the dictionary
            For i = 0 To shortcutsArray.Length - 1
                If shortcutsArray(i).Contains("=") Then
                    shortcutPair = shortcutsArray(i).Split("=")
                    shortcutDict(Trim(shortcutPair(0))) = Trim(shortcutPair(1))
                End If
            Next
            myControl.Visible = True

            ' Add menu items and assign shortcuts
            ' The OnAction refers to a Word Macro that has to be loaded as a helper for this to work; it will call up the BridgeSubs methods

            If Not String.IsNullOrWhiteSpace(INI_Language1) Then
                subControl = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
                subControl.Caption = "To " & INI_Language1
                subControl.FaceId = 6112
                subControl.Visible = True
                subControl.OnAction = "CallInLanguage1"
                If shortcutDict.ContainsKey(subControl.Caption) Then ' Check for key existence
                    subControl.TooltipText = "Shortcut: " & shortcutDict(subControl.Caption) ' Access the value
                End If
            End If
            If Not String.IsNullOrWhiteSpace(INI_Language2) Then
                subControl = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
                subControl.Caption = "To " & INI_Language2
                subControl.OnAction = "CallInLanguage2"
                subControl.FaceId = 6112

                If shortcutDict.ContainsKey(subControl.Caption) Then
                    subControl.TooltipText = "Shortcut: " & shortcutDict(subControl.Caption)
                End If
            End If

            subControl = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subControl.Caption = "To Other (text)"
            subControl.OnAction = "CallInOther"
            subControl.FaceId = 6112
            If shortcutDict.ContainsKey(subControl.Caption) Then
                subControl.TooltipText = "Shortcut: " & shortcutDict(subControl.Caption)
            End If

            subControl = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subControl.Caption = "To Other (cells)"
            subControl.OnAction = "CallInOtherFormulas"
            subControl.FaceId = 6112
            If shortcutDict.ContainsKey(subControl.Caption) Then
                subControl.TooltipText = "Shortcut: " & shortcutDict(subControl.Caption)
            End If

            subControl = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subControl.Caption = "Correct"
            subControl.OnAction = "CallCorrect"
            subControl.FaceId = 329

            If shortcutDict.ContainsKey(subControl.Caption) Then
                subControl.TooltipText = "Shortcut: " & shortcutDict(subControl.Caption)
            End If

            subControl = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subControl.Caption = "Write Neatly"
            subControl.OnAction = "CallNeatly"
            subControl.FaceId = 162

            If shortcutDict.ContainsKey(subControl.Caption) Then
                subControl.TooltipText = "Shortcut: " & shortcutDict(subControl.Caption)
            End If

            subControl = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subControl.Caption = "Shorten"
            subControl.OnAction = "CallShorten"
            subControl.FaceId = 292
            If shortcutDict.ContainsKey(subControl.Caption) Then
                subControl.TooltipText = "Shortcut: " & shortcutDict(subControl.Caption)
            End If

            subControl = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subControl.Caption = "Anonymize"
            subControl.OnAction = "CallAnonymize"
            subControl.FaceId = 7502
            If shortcutDict.ContainsKey(subControl.Caption) Then
                subControl.TooltipText = "Shortcut: " & shortcutDict(subControl.Caption)
            End If

            subControl = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subControl.Caption = "Switch Party"
            subControl.OnAction = "CallSwitchParty"
            subControl.FaceId = 327
            If shortcutDict.ContainsKey(subControl.Caption) Then
                subControl.TooltipText = "Shortcut: " & shortcutDict(subControl.Caption)
            End If

            subControl = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subControl.Caption = "Freestyle"
            subControl.OnAction = "CallFreestyleNM"
            subControl.FaceId = 346
            If shortcutDict.ContainsKey(subControl.Caption) Then
                subControl.TooltipText = "Shortcut: " & shortcutDict(subControl.Caption)
            End If

            If INI_SecondAPI Then
                subControl = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
                subControl.Caption = "Freestyle (" & INI_Model_2 & ")"
                subControl.OnAction = "CallFreestyleAM"
                subControl.FaceId = 346
                If shortcutDict.ContainsKey(subControl.Caption) Then
                    subControl.TooltipText = "Shortcut: " & shortcutDict(subControl.Caption)
                End If

            End If

            ' Create new submenu "Excel helpers"
            excelHelpersMenu = CType(myControl.Controls.Add(Type:=MsoControlType.msoControlPopup, Temporary:=True), CommandBarPopup)
            excelHelpersMenu.Caption = "Excel Helpers"

            subSubControl = CType(excelHelpersMenu.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subSubControl.Caption = "Adjust Cell Height"
            subSubControl.OnAction = "CallAdjustHeight"
            subSubControl.FaceId = 1647

            If shortcutDict.ContainsKey(subSubControl.Caption) Then
                subSubControl.TooltipText = "Shortcut: " & shortcutDict(subSubControl.Caption)
            End If


            subSubControl = CType(excelHelpersMenu.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subSubControl.Caption = "Adjust Size of Notes"
            subSubControl.OnAction = "CallAdjustLegacyNotes"
            subSubControl.FaceId = 1996

            If shortcutDict.ContainsKey(subSubControl.Caption) Then
                subSubControl.TooltipText = "Shortcut: " & shortcutDict(subSubControl.Caption)
            End If

            subSubControl = CType(excelHelpersMenu.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
            subSubControl.Caption = "Regex Search && Replace"
            subSubControl.OnAction = "CallRegexSearchReplace"
            subSubControl.FaceId = 288
            If shortcutDict.ContainsKey(subSubControl.Caption) Then
                subSubControl.TooltipText = "Shortcut: " & shortcutDict(subSubControl.Caption)
            End If

            If Not String.IsNullOrWhiteSpace(INI_ShortcutsWordExcel) Then
                ' Assign shortcuts using the dictionary
                If Not String.IsNullOrWhiteSpace(INI_Language1) Then AssignShortcut("To " & INI_Language1, "CallInLanguage1", shortcutDict)
                If Not String.IsNullOrWhiteSpace(INI_Language2) Then AssignShortcut("To " & INI_Language2, "CallInLanguage2", shortcutDict)
                AssignShortcut("To Other (text)", "CallInOther", shortcutDict)
                AssignShortcut("To Other (cells)", "CallInOther", shortcutDict)
                AssignShortcut("Correct", "CallCorrect", shortcutDict)
                AssignShortcut("Write Neatly", "CallImprove", shortcutDict)
                AssignShortcut("Shorten", "CallShorten", shortcutDict)
                AssignShortcut("Anonymize", "CallAnonymize", shortcutDict)
                AssignShortcut("Switch Party", "CallSwitchParty", shortcutDict)
                AssignShortcut("Freestyle", "CallFreestyleNM", shortcutDict)

                ' Assign shortcuts for second API if applicable
                If INI_SecondAPI Then
                    AssignShortcut("Freestyle (" & INI_Model_2 & ")", "CallFreestyleAM", shortcutDict)
                End If

                ' Assign shortcuts for submenu "Excel helpers"
                AssignShortcut("Adjust Cell Height", "CallAdjustheight", shortcutDict)
                AssignShortcut("Adjust Size of Notes", "CallAdjustLegacyNotes", shortcutDict)
                AssignShortcut("Regex Search & Replace", "CallRegexSearchReplace", shortcutDict)
                AssignShortcut("Regex Search && Replace", "CallRegexSearchReplace", shortcutDict)
            End If
        Catch ex As System.Exception

        End Try
    End Sub
    Public Sub AssignShortcut(ByVal controlName As String, ByVal macro As String, ByRef shortcutDict As Dictionary(Of String, String))
        Dim shortcutKey As String
        Dim keyCombination As String
        Try
            ' Check if there is a shortcut assigned for this control
            If shortcutDict.ContainsKey(controlName) Then
                shortcutKey = shortcutDict(controlName)
            Else
                Exit Sub ' No shortcut assigned
            End If

            ' Build the key combination string from the shortcutKey text
            keyCombination = BuildKeyCodeFromText(shortcutKey)

            If Not String.IsNullOrEmpty(keyCombination) Then
                ' Assign the shortcut key to the macro in Excel using Application.OnKey
                Globals.ThisAddIn.Application.OnKey(keyCombination, macro)
            End If
        Catch ex As System.Exception
            ' Handle exceptions gracefully
            ' Debug.WriteLine("Error in AssignShortcut: " & ex.Message)
        End Try
    End Sub

    Public Function BuildKeyCodeFromText(ByVal shortcutKey As String) As String
        Dim parts() As String
        Dim keysCollection As New List(Of String)()
        Dim keyCombination As String = ""

        Try
            parts = shortcutKey.Split("-"c)

            For Each part As String In parts
                Select Case part.Trim().ToUpper()
                    Case "CTRL"
                        keysCollection.Add("^") ' Control key representation in Excel
                    Case "SHIFT"
                        keysCollection.Add("+") ' Shift key representation in Excel
                    Case "ALT"
                        keysCollection.Add("%") ' Alt key representation in Excel

                ' Map digits directly
                    Case "0" To "9"
                        keysCollection.Add(part.Trim())

                ' Map function keys directly
                    Case "F1" To "F12"
                        keysCollection.Add(part.Trim())

                ' Letters mapped directly
                    Case "A" To "Z"
                        keysCollection.Add(part.Trim().ToUpper())

                    Case Else
                        ' Unknown key
                        Return ""
                End Select
            Next

            ' Combine the keys into a single shortcut string for VBA
            keyCombination = String.Join("", keysCollection)

            Return keyCombination

        Catch ex As System.Exception
            ' Handle errors gracefully
            ' Debug.WriteLine("Error in BuildKeyCodeFromText: " & ex.Message)
            Return ""
        End Try
    End Function


    Public Sub RemoveOldContextMenu()
        Dim application As Excel.Application = Globals.ThisAddIn.Application

        ' Array of relevant context menus
        Dim contextMenus As String() = {"Cell", "Row", "Column", "List Range Popup", "PivotTable Context Menu", "Text Box", "Drawing Object", "Chart"}

        ' Iterate through all CommandBars
        For Each cb As CommandBar In application.CommandBars
            If cb.Type = MsoBarType.msoBarTypePopup Then
                ' Check if the context menu is relevant
                If contextMenus.Contains(cb.Name) Then
                    ' Remove the context menu if it exists
                    For Each ctrl As CommandBarControl In cb.Controls
                        If ctrl.Type = MsoControlType.msoControlPopup AndAlso ctrl.Caption = RIMenu Then
                            Try
                                ctrl.Delete()
                            Catch ex As System.Exception
                                ' Handle errors if needed, e.g., logging
                                'Debug.WriteLine($"Error removing control: {ex.Message}")
                            End Try
                        End If
                    Next
                End If
            End If
        Next

    End Sub



#End Region

    ' Declare them publicly so that InterpolateAtRuntime can access them; case-sensitive

    Public TranslateLanguage As String
    Public ShortenLength As Double
    Public SummaryLength As Integer
    Public OtherPrompt As String = ""
    Public Separator As String = ""
    Public LineNumber As Integer = 0
    Public Context As String
    Public SysPrompt As String
    Public OldParty, NewParty As String
    Public SelectedText As String

    Public Structure CellState
        Public WorksheetName As String
        Public CellAddress As String
        Public OldValue As Object
        Public HadFormula As Boolean
        Public OldFormula As String
    End Structure

    Public Shared undoStates As New List(Of CellState)

    Public Async Function InLanguage1() As Task(Of Boolean)
        System.Windows.Forms.Application.DoEvents()
        TranslateLanguage = INI_Language1
        Dim result As Boolean = Await ProcessSelectedRange(SP_Translate, True, False, False, False, True, False)
    End Function
    Public Async Function InLanguage2() As Task(Of Boolean)
        System.Windows.Forms.Application.DoEvents()
        TranslateLanguage = INI_Language2
        Dim result As Boolean = Await ProcessSelectedRange(SP_Translate, True, False, False, False, True, False)
    End Function
    Public Async Function InOther() As Task(Of Boolean)
        System.Windows.Forms.Application.DoEvents()
        Dim selectedRange As Excel.Range = TryCast(Globals.ThisAddIn.Application.Selection, Excel.Range)
        TranslateLanguage = SLib.ShowCustomInputBox("Enter your target language:", $"{AN} Translate", True)
        If Not String.IsNullOrEmpty(TranslateLanguage) Then
            If selectedRange IsNot Nothing Then
                selectedRange.Select()
            End If

            Dim result As Boolean = Await ProcessSelectedRange(SP_Translate, True, False, False, False, True, False)
        End If

    End Function
    Public Async Function InOtherFormulas() As Task(Of Boolean)
        System.Windows.Forms.Application.DoEvents()
        TranslateLanguage = SLib.ShowCustomInputBox("Enter your target language:", $"{AN} Translate", True)
        If Not String.IsNullOrEmpty(TranslateLanguage) Then
            Dim result As Boolean = Await ProcessSelectedRange(SP_Translate, True, False, True, False, True, False)
        End If
    End Function
    Public Async Function Correct() As Task(Of Boolean)
        System.Windows.Forms.Application.DoEvents()
        Dim result As Boolean = Await ProcessSelectedRange(SP_Correct, True, False, False, False, True, False)
    End Function
    Public Async Function Improve() As Task(Of Boolean)
        System.Windows.Forms.Application.DoEvents()
        Dim selectedRange As Excel.Range = TryCast(Globals.ThisAddIn.Application.Selection, Excel.Range)
        If selectedRange Is Nothing Then
            ShowCustomMessageBox("Please select the cells to be processed.")
            Return False
        End If

        Context = Trim(SLib.ShowCustomInputBox("Please provide the context that should be taken into account, if any:", $"{AN} Write Neatly", True))

        If String.IsNullOrWhiteSpace(Context) Then
            Context = "n/a"
        End If

        If selectedRange IsNot Nothing Then
            selectedRange.Select()
        End If

        Dim result As Boolean = Await ProcessSelectedRange(SP_WriteNeatly, True, False, False, False, True, False)

    End Function
    Public Async Function Anonymize() As Task(Of Boolean)
        System.Windows.Forms.Application.DoEvents()
        Dim result As Boolean = Await ProcessSelectedRange(SP_Anonymize, True, False, False, False, True, False)
    End Function
    Public Async Function Shorten() As Task(Of Boolean)
        System.Windows.Forms.Application.DoEvents()
        Dim selectedRange As Excel.Range = TryCast(Globals.ThisAddIn.Application.Selection, Excel.Range)

        If selectedRange Is Nothing Then
            ShowCustomMessageBox("Please select the cells to be processed.")
            Return False
        End If

        Dim totalLength As Integer = 0
        Dim maxLength As Integer = 0
        Dim cellCount As Integer = 0

        For Each cell As Excel.Range In selectedRange.Cells
            If Not CellProtected(cell) AndAlso Not cell.HasFormula Then
                Dim cellText As String = CStr(cell.Value)
                If Not String.IsNullOrEmpty(cellText) Then
                    Dim textLength As Integer = cellText.Length
                    totalLength += textLength
                    If textLength > maxLength Then
                        maxLength = textLength
                    End If
                    cellCount += 1
                End If
            End If
        Next

        Dim averageLength As Double = If(cellCount > 0, totalLength / cellCount, 0)

        Dim UserInput As String
        Dim ShortenPercentValue As Integer = 0
        Do
            UserInput = Trim(SLib.ShowCustomInputBox($"Enter the percentage by which the text of each selected cell should be shortened (the cells have have of average {averageLength:n1} words and {maxLength} at max; {ShortenPercent}% will cut approx. " & (averageLength * ShortenPercent / 100) & " words in average):", $"{AN} Shortener", True, CStr(ShortenPercent) & "%"))
            If String.IsNullOrEmpty(UserInput) Then
                Return False
            End If
            UserInput = UserInput.Replace("%", "").Trim()
            If Integer.TryParse(UserInput, ShortenPercentValue) AndAlso ShortenPercentValue >= 1 AndAlso ShortenPercentValue <= 99 Then
                Exit Do
            Else
                ShowCustomMessageBox("Please enter a valid percentage between 1 and 99.")
            End If
        Loop
        If ShortenPercentValue = 0 Then Return False
        If selectedRange IsNot Nothing Then
            selectedRange.Select()
        End If

        Dim result As Boolean = Await ProcessSelectedRange(SP_Shorten, True, False, False, True, False, ShortenPercentValue)

    End Function
    Public Async Function SwitchParty() As Task(Of Boolean)
        System.Windows.Forms.Application.DoEvents()
        Dim selectedRange As Excel.Range = TryCast(Globals.ThisAddIn.Application.Selection, Excel.Range)

        If selectedRange Is Nothing Then
            ShowCustomMessageBox("Please select the cells to be processed.")
            Return False
        End If

        Dim UserInput As String
        Do
            UserInput = Trim(SLib.ShowCustomInputBox("Please provide the original party name and the new party name, separated by a comma (example: Elvis Presley, Taylor Swift):", $"{AN} Switch Party", True))

            If String.IsNullOrEmpty(UserInput) Then
                Return False
            End If

            Dim parts() As String = UserInput.Split(","c)
            If parts.Length = 2 Then
                OldParty = parts(0).Trim()
                NewParty = parts(1).Trim()
                Exit Do
            Else
                ShowCustomMessageBox("Please enter two names separated by a comma.")
            End If
        Loop
        If selectedRange IsNot Nothing Then
            selectedRange.Select()
        End If

        Dim result As Boolean = Await ProcessSelectedRange(SP_SwitchParty, True, False, False, False, True, False)


    End Function

    Public Sub ShowChatForm()
        If chatForm Is Nothing OrElse chatForm.IsDisposed Then
            chatForm = New frmAIChat(_context)

            ' Set the location and size before showing the form
            If My.Settings.FormLocation <> System.Drawing.Point.Empty AndAlso My.Settings.FormSize <> Size.Empty Then
                chatForm.StartPosition = FormStartPosition.Manual
                chatForm.Location = My.Settings.FormLocation
                chatForm.Size = My.Settings.FormSize
            Else
                ' Default to center screen if no settings are available
                chatForm.StartPosition = FormStartPosition.Manual
                Dim screenBounds As System.Drawing.Rectangle = Screen.PrimaryScreen.WorkingArea
                chatForm.Location = New System.Drawing.Point((screenBounds.Width - chatForm.Width) \ 2, (screenBounds.Height - chatForm.Height) \ 2)
                chatForm.Size = New Size(650, 500) ' Set default size if needed
            End If
        End If

        ' Show and bring the form to the front
        chatForm.Show()
        chatForm.BringToFront()
    End Sub

    Public Async Function FreestyleNM() As Task(Of Boolean)
        System.Windows.Forms.Application.DoEvents()
        Dim result As Boolean = Await Freestyle(False)
    End Function

    Public Async Function FreestyleAM() As Task(Of Boolean)
        System.Windows.Forms.Application.DoEvents()

        If Not String.IsNullOrWhiteSpace(INI_AlternateModelPath) Then

            If Not ShowModelSelection(_context, INI_AlternateModelPath) Then
                originalConfigLoaded = False
                Exit Function
            End If

        End If

        Dim result As Boolean = Await Freestyle(True)

        If originalConfigLoaded Then
            RestoreDefaults(_context, originalConfig)
            originalConfigLoaded = False
        End If

    End Function

    Public Async Function Freestyle(ByVal UseSecondAPI As Boolean) As Task(Of Boolean)

        Dim selectedRange As Excel.Range = TryCast(Globals.ThisAddIn.Application.Selection, Excel.Range)
        Dim NoSelectedCells As Boolean = False
        Dim DoClipboard As Boolean = False
        Dim DoFormulas As Boolean = True
        Dim DoBubbles As Boolean = False
        Dim DoColor As Boolean = False
        Dim DoFileObject As Boolean = False
        Dim DoFileObjectClip As Boolean = False
        Dim DoPane As Boolean = False
        Dim DoBatch As Boolean = False
        Dim BatchPath As String = ""

        Dim LastPromptInstruct As String = If(String.IsNullOrWhiteSpace(My.Settings.LastPrompt), "", "; Ctrl-P for your last prompt")
        Dim PureInstruct As String = $"; use '{PurePrefix}' for direct prompting"

        Dim DefaultPrefix As String = INI_DefaultPrefix
        Dim DefaultPrefixText As String = ""

        If selectedRange Is Nothing Then
            NoSelectedCells = True
        End If

        Dim DoRange As Boolean = True
        Dim CBCInstruct As String = $"with '{CellByCellPrefix}' or '{CellByCellPrefix2} if the instruction should be executed cell-by-cell"
        Dim TextInstruct As String = $"use '{TextPrefix}' or '{TextPrefix2}' if the instruction should apply cell-by-cell, but only to text cells"
        Dim BatchInstruct As String = $"use '{BatchPrefix}' if to process a directory of files"
        Dim BubblesInstruct As String = $"use '{BubblesPrefix}' for inserting comments only"
        Dim PaneInstruct As String = $"use '{PanePrefix}' for using the pane"
        Dim ExtInstruct As String = $"; insert '{ExtTrigger}' (multiple times) for text of (a) file(s) (txt, docx, pdf) or '{ExtWSTrigger}' to add more worksheet(s)"
        Dim AddonInstruct As String = $"; add '{ColorTrigger}' to check for colorcodes"
        Dim ObjectInstruct As String = $"; add '{ObjectTrigger}'/'{ObjectTrigger2}' for adding a file object"
        Dim FileObject As String = ""
        Dim InsertWS As String = ""

        If UseSecondAPI Then
            If Not String.IsNullOrWhiteSpace(INI_APICall_Object_2) Then
                AddonInstruct += ObjectInstruct.Replace("; add", ",")
                DoFileObject = True
            End If
        Else
            If Not String.IsNullOrWhiteSpace(INI_APICall_Object) Then
                AddonInstruct += ObjectInstruct.Replace("; add", ",")
                DoFileObject = True
            End If
        End If

        Dim PromptLibInstruct As String = ""
        If INI_PromptLib Then
            PromptLibInstruct = " or press 'OK' for the prompt library"
        End If

        If DefaultPrefix.Trim() <> "" Then
            DefaultPrefixText = $" (default prefix: '{DefaultPrefix}')"
        End If


        Dim OptionalButtons As System.Tuple(Of String, String, String)() = {
                            System.Tuple.Create("OK, use pane", $"Use this to automatically insert '{PanePrefix}' as a prefix.", PanePrefix)
                        }

        If Not NoSelectedCells Then
            OtherPrompt = Trim(SLib.ShowCustomInputBox($"Please provide the prompt you wish to execute on the selected cells (start {CBCInstruct}; {TextInstruct}; {PaneInstruct}; {BatchInstruct}; {BubblesInstruct})" & PromptLibInstruct & PureInstruct & ExtInstruct & AddonInstruct & LastPromptInstruct & DefaultPrefixText & ":", $"{AN} Freestyle (using " & If(UseSecondAPI, INI_Model_2, INI_Model) & ")", False, "", My.Settings.LastPrompt, OptionalButtons))
        Else
            OtherPrompt = Trim(SLib.ShowCustomInputBox($"Please provide the prompt you wish to execute {PromptLibInstruct} (the result will be shown to you before inserting anything into your worksheet); {PaneInstruct}{BatchInstruct}{PureInstruct}{ExtInstruct}{AddonInstruct}{LastPromptInstruct}{DefaultPrefixText}:", $"{AN} Freestyle (using " & If(UseSecondAPI, INI_Model_2, INI_Model) & ")", False, "", My.Settings.LastPrompt, OptionalButtons))
            DoRange = True
        End If

        If String.IsNullOrEmpty(OtherPrompt) And OtherPrompt <> "ESC" And INI_PromptLib Then

            Dim promptlibresult As (String, Boolean, Boolean, Boolean)

            promptlibresult = ShowPromptSelector(INI_PromptLibPath, INI_PromptLibPathLocal, Nothing, Nothing)

            OtherPrompt = promptlibresult.Item1
            DoClipboard = promptlibresult.Item4

            If OtherPrompt = "" Then
                Return False
            End If
        Else
            If String.IsNullOrEmpty(OtherPrompt) Or OtherPrompt = "ESC" Then Return False
        End If

        ' Check if OtherPrompt starts with a word ending with a colon
        If Not String.IsNullOrWhiteSpace(OtherPrompt) Then
            Dim firstWord As String = OtherPrompt.Split({" "c}, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault()
            If firstWord IsNot Nothing AndAlso Not firstWord.EndsWith(":"c) Then

                Dim prefix As String = DefaultPrefix.Trim()

                ' Ensure prefix ends with colon and space
                If prefix <> "" AndAlso Not prefix.EndsWith(":"c) Then
                    prefix &= ":"
                End If

                OtherPrompt = prefix & " " & OtherPrompt.Trim()
                OtherPrompt = OtherPrompt.Trim()
            End If
        End If


        My.Settings.LastPrompt = OtherPrompt
        My.Settings.Save()

        If Not SharedMethods.ProcessParameterPlaceholders(OtherPrompt) Then
            ShowCustomMessageBox("Freestyle canceled.", $"{AN} Freestyle")
            Return False
        End If


        If OtherPrompt.StartsWith(CellByCellPrefix, StringComparison.OrdinalIgnoreCase) And DoFormulas Then
            OtherPrompt = OtherPrompt.Substring(CellByCellPrefix.Length).Trim()
            DoRange = False
        End If
        If OtherPrompt.StartsWith(CellByCellPrefix2, StringComparison.OrdinalIgnoreCase) And DoFormulas Then
            OtherPrompt = OtherPrompt.Substring(CellByCellPrefix2.Length).Trim()
            DoRange = False
        End If
        If OtherPrompt.StartsWith(TextPrefix, StringComparison.OrdinalIgnoreCase) And DoFormulas Then
            OtherPrompt = OtherPrompt.Substring(TextPrefix.Length).Trim()
            DoRange = False
            DoFormulas = False
        End If
        If OtherPrompt.StartsWith(TextPrefix2, StringComparison.OrdinalIgnoreCase) And DoFormulas Then
            OtherPrompt = OtherPrompt.Substring(TextPrefix2.Length).Trim()
            DoRange = False
            DoFormulas = False
        End If
        If OtherPrompt.StartsWith(BubblesPrefix, StringComparison.OrdinalIgnoreCase) And selectedRange IsNot Nothing Then
            OtherPrompt = OtherPrompt.Substring(BubblesPrefix.Length).Trim()
            DoBubbles = True
            DoRange = True
        End If
        If OtherPrompt.StartsWith(PanePrefix, StringComparison.OrdinalIgnoreCase) And DoRange Then
            OtherPrompt = OtherPrompt.Substring(PanePrefix.Length).Trim()
            DoPane = True
            DoRange = True
        End If
        If OtherPrompt.StartsWith(BatchPrefix, StringComparison.OrdinalIgnoreCase) Then
            OtherPrompt = OtherPrompt.Substring(BatchPrefix.Length).Trim()
            DoPane = False
            DoRange = True
            DoBatch = True


            Try
                ' --- Excel context (as requested) ---
                Dim activeCell As Microsoft.Office.Interop.Excel.Range = Application.ActiveCell
                Dim ws As Microsoft.Office.Interop.Excel.Worksheet = CType(Application.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)

                Dim currentRow As System.Int32 = 1
                If activeCell IsNot Nothing Then
                    currentRow = System.Convert.ToInt32(activeCell.Row, System.Globalization.CultureInfo.InvariantCulture)
                End If

                Dim maxRow As System.Int32 = System.Convert.ToInt32(ws.Rows.Count, System.Globalization.CultureInfo.InvariantCulture)

                ' --- Loop until user provides a valid line number or cancels/ESC ---
                Dim lineNumberAnswer As System.String = System.String.Empty
                Do
                    lineNumberAnswer = ShowCustomInputBox(
                    "Please provide the starting line number at which the results should be inserted (1 for first line, 2 for second line etc.):", $"{AN} Freestyle Batch",
                    True,
                    currentRow.ToString(System.Globalization.CultureInfo.InvariantCulture)
                )

                    If lineNumberAnswer Is Nothing Then
                        Return Nothing
                    End If

                    lineNumberAnswer = lineNumberAnswer.Trim()
                    If lineNumberAnswer.Length = 0 Then
                        Return Nothing
                    End If

                    If System.String.Compare(lineNumberAnswer, "ESC", System.StringComparison.OrdinalIgnoreCase) = 0 Then
                        Return Nothing
                    End If

                    Dim parsedRow As System.Int32
                    If Not System.Int32.TryParse(lineNumberAnswer, parsedRow) OrElse parsedRow < 1 OrElse parsedRow > maxRow Then
                        ShowCustomMessageBox("The provided value is not a valid line number between 1 and " &
                                         maxRow.ToString(System.Globalization.CultureInfo.InvariantCulture) & ".")
                        ' Loop back to input box
                    Else
                        LineNumber = parsedRow
                        Exit Do
                    End If
                Loop

                ' --- Directory picker (browse or type) ---
                Dim initialPath As System.String
                If (Application.ActiveWorkbook IsNot Nothing) AndAlso
               (Application.ActiveWorkbook.Path IsNot Nothing) AndAlso
               (Application.ActiveWorkbook.Path.Length > 0) Then
                    initialPath = Application.ActiveWorkbook.Path
                Else
                    initialPath = System.Environment.GetFolderPath(System.Environment.SpecialFolder.MyDocuments)
                End If

                Using dlg As New System.Windows.Forms.FolderBrowserDialog()
                    dlg.Description = "Select the directory that contains the batch text files."
                    dlg.ShowNewFolderButton = False
                    dlg.SelectedPath = initialPath

                    Dim result As System.Windows.Forms.DialogResult = dlg.ShowDialog()
                    If result <> System.Windows.Forms.DialogResult.OK Then
                        Return Nothing
                    End If

                    Dim selectedPath As System.String = dlg.SelectedPath
                    If System.String.IsNullOrWhiteSpace(selectedPath) OrElse Not System.IO.Directory.Exists(selectedPath) Then
                        ShowCustomMessageBox("No directory was selected or it does not exist.")
                        Return Nothing
                    End If

                    ' --- Check for expected text file types 

                    Dim hasAny As System.Boolean = False

                    ' Enumerate files and check extensions
                    For Each filePath As System.String In System.IO.Directory.EnumerateFiles(selectedPath, "*.*", System.IO.SearchOption.TopDirectoryOnly)
                        Dim ext As System.String = System.IO.Path.GetExtension(filePath)
                        If allowedExtensions.Contains(ext) Then
                            hasAny = True
                            Exit For
                        End If
                    Next

                    ' Handle case when no allowed files exist
                    If Not hasAny Then
                        ShowCustomMessageBox(
                                        "The selected directory does not contain any files of the expected types: " &
                                        System.String.Join(", ", allowedExtensions) & "."
                                    )
                        Return Nothing
                    End If

                    BatchPath = selectedPath
                End Using

            Catch ex As System.Exception
                ShowCustomMessageBox("GetLineNumber and Path resulted in an Error: " & ex.Message)
                Return Nothing
            End Try
        End If
        If DoFileObject AndAlso OtherPrompt.IndexOf(ObjectTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
            OtherPrompt = OtherPrompt.Replace(ObjectTrigger, "(a file object follows)").Trim()
        ElseIf DoFileObject AndAlso OtherPrompt.IndexOf(ObjectTrigger2, StringComparison.OrdinalIgnoreCase) >= 0 Then
            OtherPrompt = OtherPrompt.Replace(ObjectTrigger2, "(a file object follows)").Trim()
            DoFileObjectClip = True
        Else
            DoFileObject = False
        End If

        If selectedRange IsNot Nothing Then
            selectedRange.Select()
        End If

        If Not String.IsNullOrEmpty(OtherPrompt) And OtherPrompt.IndexOf(ColorTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
            DoColor = True
            OtherPrompt = Regex.Replace(OtherPrompt, Regex.Escape(ColorTrigger), "", RegexOptions.IgnoreCase)
        End If

        If Not String.IsNullOrEmpty(OtherPrompt) AndAlso OtherPrompt.IndexOf(ExtTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
            ' Count total occurrences first (case-insensitive) so inserted file text containing {doc} does not trigger extra loops.
            Dim totalOccurrences As Integer = Regex.Matches(OtherPrompt, Regex.Escape(ExtTrigger), RegexOptions.IgnoreCase).Count

            ' Detect if a placeholder occurrence is already enclosed by any tag, e.g. <tag>...{doc}...</tag>
            Dim wrappedPattern As String =
                "<(?<name>[A-Za-z][\w\-]*)\b[^>]*>[^<]*" & Regex.Escape(ExtTrigger) & "[^<]*</\k<name>>"

            If totalOccurrences = 1 Then
                ' Single-occurrence behavior with optional auto-wrapping
                DragDropFormLabel = ""
                DragDropFormFilter = ""
                Dim doc As String = Await GetFileContent(Nothing, False, Not String.IsNullOrWhiteSpace(INI_APICall_Object), True)
                If String.IsNullOrWhiteSpace(doc) Then
                    ShowCustomMessageBox("The file you have selected is empty or not supported - exiting.")
                    Return False
                End If

                Dim isWrapped As Boolean = Regex.IsMatch(OtherPrompt, wrappedPattern, RegexOptions.IgnoreCase)
                Dim replacementText As String = If(isWrapped, doc, $"<document>{doc}</document>")

                OtherPrompt = Regex.Replace(OtherPrompt, Regex.Escape(ExtTrigger), replacementText, RegexOptions.IgnoreCase)
                ShowCustomMessageBox($"This file will be included in your prompt where you have referred to {ExtTrigger}: " & vbCrLf & vbCrLf & doc)
            Else
                ' Multi-occurrence behavior: prompt separately for each placeholder
                For occurrence As Integer = 1 To totalOccurrences
                    Dim idx As Integer = OtherPrompt.IndexOf(ExtTrigger, StringComparison.OrdinalIgnoreCase)
                    If idx < 0 Then Exit For

                    DragDropFormLabel = ""
                    DragDropFormFilter = ""
                    Dim docPart As String = Await GetFileContent(Nothing, False, Not String.IsNullOrWhiteSpace(INI_APICall_Object), True)
                    If String.IsNullOrWhiteSpace(docPart) Then
                        ShowCustomMessageBox($"The file you selected for occurrence #{occurrence} is empty or not supported - exiting.")
                        Return False
                    End If

                    ' Determine if this specific occurrence is already wrapped by any tag pair
                    Dim isWrappedThis As Boolean = False
                    Dim mcol As MatchCollection = Regex.Matches(OtherPrompt, wrappedPattern, RegexOptions.IgnoreCase)
                    For Each m As Match In mcol
                        If idx >= m.Index AndAlso idx < m.Index + m.Length Then
                            isWrappedThis = True
                            Exit For
                        End If
                    Next

                    Dim replacementText As String =
                        If(isWrappedThis, docPart, $"<document{occurrence}>{docPart}</document{occurrence}>")

                    ' Replace only the first remaining occurrence (manual replacement keeps later placeholders intact)
                    OtherPrompt = OtherPrompt.Substring(0, idx) & replacementText & OtherPrompt.Substring(idx + ExtTrigger.Length)

                    ShowCustomMessageBox($"This file will be included at occurrence #{occurrence} (of {totalOccurrences}) where you used {ExtTrigger}:" &
                                         vbCrLf & vbCrLf & docPart)
                Next
            End If
        End If



        If Not String.IsNullOrEmpty(OtherPrompt) And OtherPrompt.IndexOf(ExtWSTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
            If Not DoRange Then
                ShowCustomMessageBox($"{ExtWSTrigger} cannot be combined with cell by cell processing - exiting.")
                Return False
            End If
            InsertWS = GatherSelectedWorksheets(True)
            Debug.WriteLine($"GatherSelectedWorksheets returned: {Left(InsertWS, 3000)}")
            If String.IsNullOrWhiteSpace(InsertWS) Then
                ShowCustomMessageBox("No content was found or an error occurred in gathering the additional worksheet(s) - exiting.")
                Return False
            ElseIf InsertWS.StartsWith("ERROR", StringComparison.OrdinalIgnoreCase) Then
                ShowCustomMessageBox($"An error occured gathering the additional worksheet(s) ({InsertWS.Substring(6).Trim()}) - exiting.")
                Return False
            ElseIf InsertWS.StartsWith("NONE", StringComparison.OrdinalIgnoreCase) Then
                ShowCustomMessageBox($"There are no other worksheets to add - exiting.")
                Return False
            End If
            OtherPrompt = Regex.Replace(OtherPrompt, Regex.Escape(ExtWSTrigger), "", RegexOptions.IgnoreCase)
        End If

        If DoFileObject Then
            If DoFileObjectClip Then
                FileObject = "clipboard"
            Else
                DragDropFormLabel = "All file types that are supported by your LLM."
                DragDropFormFilter = "Supported Files|*.*"
                FileObject = GetFileName()
                DragDropFormLabel = ""
                DragDropFormFilter = ""
                If String.IsNullOrWhiteSpace(FileObject) Then
                    ShowCustomMessageBox("No file object has been selected - will abort. You can try again (use Ctrl-P to re-insert your prompt).")
                    Return False
                End If
            End If
        End If

        If OtherPrompt.StartsWith(PurePrefix, StringComparison.OrdinalIgnoreCase) Then
            OtherPrompt = OtherPrompt.Substring(PurePrefix.Length).Trim()
            Dim result As Boolean = Await ProcessSelectedRange(OtherPrompt, True, DoRange, DoFormulas, DoBubbles, False, UseSecondAPI, 0, True, DoColor, DoPane, FileObject, InsertWS)
        Else
            If Not NoSelectedCells Then
                If DoRange Then
                    Dim result As Boolean = Await ProcessSelectedRange(SP_RangeOfCells, True, DoRange, DoFormulas, DoBubbles, False, UseSecondAPI, 0, True, DoColor, DoPane, FileObject, InsertWS, BatchPath)
                Else
                    Dim result As Boolean = Await ProcessSelectedRange(SP_FreestyleText, True, DoRange, DoFormulas, DoBubbles, False, UseSecondAPI, DoColor, DoPane, FileObject, InsertWS)
                End If
            Else
                Dim result As Boolean = Await ProcessSelectedRange(SP_RangeOfCells, True, DoRange, DoFormulas, DoBubbles, False, UseSecondAPI, 0, True, DoColor, DoPane, FileObject, InsertWS, BatchPath)
            End If
        End If


    End Function


    ' ProcessSelectedRang
    '
    ' This function processes the selected range of cells in Excel. It takes the following parameters:
    ' - SysCommand: The system command to be executed
    ' - CheckMaxToken: A boolean value indicating whether the maximum token count should be checked
    ' - DoRange: A boolean value indicating whether the selected range should be processed
    ' - DoFormulas: A boolean value indicating whether formulas should be processed
    ' - DoBubbles: A boolean value indicating whether to insert comments
    ' - SelectionMandatory: A boolean value indicating whether a selection is mandatory
    ' - UseSecondAPI: A boolean value indicating whether the second API should be used
    ' - Optional: ShortenPercentValue: A percentage value by which the text should be shortened (for calculating the word count for each cell individually)
    ' - Optional: Freestyle: A boolean value indicating whether to use freestyle mode
    ' - Optional: DoColor: A boolean value indicating whether to check for color codes
    ' - Optional: DoPane: A boolean value indicating whether the output should go into the pane
    ' - Optional: FileObject: The name of the file (or clipboard) where to get an object to include in the LLM call
    ' - Optional: InsertWS: The text representation of additional worksheets to be included in the prompt
    ' - Optional: BatchFilePath: The path to a directory containing files to be processed in batch mode

    Private Async Function ProcessSelectedRange(ByVal SysCommand As String, CheckMaxToken As Boolean, DoRange As Boolean, DoFormulas As Boolean, DoBubbles As Boolean, SelectionMandatory As Boolean, ByVal UseSecondAPI As Boolean, Optional ShortenPercentValue As Integer = 0, Optional Freestyle As Boolean = False, Optional DoColor As Boolean = False, Optional DoPane As Boolean = False, Optional FileObject As String = "", Optional InsertWS As String = "", Optional BatchFilePath As String = "") As Task(Of Boolean)

        Dim excelApp As Excel.Application = CType(Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application"), Excel.Application)

        Dim selectedRange As Excel.Range = TryCast(Globals.ThisAddIn.Application.Selection, Excel.Range)
        Dim NoSelectedCells As Boolean = False
        Dim DoShorten As Boolean = False

        If DoBubbles Then SelectionMandatory = True

        ' Get the used range of the active sheet
        Dim activeSheet As Microsoft.Office.Interop.Excel.Worksheet = Globals.ThisAddIn.Application.ActiveSheet
        Dim usedRange As Excel.Range = activeSheet.UsedRange

        ' Check if a selection has been made
        If selectedRange Is Nothing Then
            NoSelectedCells = True
        Else
            ' If the entire row, column, or sheet is selected, limit to used range
            selectedRange = Globals.ThisAddIn.Application.Intersect(selectedRange, usedRange)

            ' If the intersection results in no cells, set NoSelectedCells to True
            If selectedRange Is Nothing Then
                NoSelectedCells = True
                If Freestyle Or Not SelectionMandatory Then
                    DoRange = True
                    Freestyle = True
                    SysCommand = SP_RangeOfCells
                End If
            End If
        End If

        ' Check if cells are selected and show message if mandatory selection is required
        If NoSelectedCells AndAlso SelectionMandatory Then
            ShowCustomMessageBox("Please select cells (with content) to be processed.")
            Return False
        End If

        ' Check if all selected cells are blocked
        If AreAllCellsBlocked(selectedRange) And Not DoRange Then
            ShowCustomMessageBox($"{AN} cannot do anything because the cells are blocked.")
            Return False
        End If

        If ShortenPercentValue > 0 Then
            DoShorten = True
        End If

        Dim MaxToken As Integer = If(UseSecondAPI, INI_MaxOutputToken_2, INI_MaxOutputToken)
        If Not NoSelectedCells And CheckMaxToken And MaxToken > 0 Then

            SelectedText = GetSelectedText(selectedRange, DoFormulas)

            Dim EstimatedTokens As Integer = EstimateTokenCount(SelectedText)

            If EstimatedTokens > MaxToken Then
                ShowCustomMessageBox("The content of the selected cells is larger than the maximum output your LLM can supposedly generate. Therefore, the output may be shorter than expected based on maximum tokens supported, which is " & MaxToken & " tokens. Your input (with formatting information, as the case may be) has an estimated to be " & EstimatedTokens & " tokens). Therefore, check whether the output is complete.", AN, 15)
            End If

        End If

        If Not DoShorten And BatchFilePath = "" Then SysCommand = InterpolateAtRuntime(SysCommand)

        If DoBubbles Then SysCommand = InterpolateAtRuntime(SP_BubblesExcel)

        undoStates.Clear()

        If Not DoRange Then

            Dim splash As New SplashScreen("Processing cells... press 'Esc' to abort")
            splash.Show()
            splash.Refresh()

            'Application.ScreenUpdating = False ' Prevent UI updates during processing
            Try
                For Each cell As Excel.Range In selectedRange.Cells

                    System.Windows.Forms.Application.DoEvents()

                    If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And &H8000) <> 0 Then Exit For

                    If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And 1) <> 0 Then
                        ' Exit the loop
                        Exit For
                    End If
                    Try
                        If Not IsNothing(cell.Value) AndAlso Not CellProtected(cell) Then
                            If cell.HasFormula AndAlso DoFormulas Then
                                ' Handle formulas
                                SelectedText = cell.Formula

                                If DoShorten Then
                                    Dim Textlength As Integer = SelectedText.Length
                                    ShortenLength = (Textlength - (Textlength * (100 - ShortenPercentValue) / 100))
                                    SysCommand = InterpolateAtRuntime(SysCommand)
                                End If

                                Await System.Threading.Tasks.Task.Delay(500)

                                Dim LLMResult As String = Await LLM(SysCommand & " " & SP_Add_KeepFormulasIntact, If(NoSelectedCells, "", "<TEXTTOPROCESS>" & SelectedText & "</TEXTTOPROCESS>"), "", "", 0, UseSecondAPI, True, "", FileObject)

                                LLMResult = Trim(LLMResult)

                                If Not String.IsNullOrEmpty(LLMResult) Then
                                    LLMResult = Await PostCorrection(LLMResult, UseSecondAPI)
                                End If
                                If Not String.IsNullOrWhiteSpace(LLMResult) Then
                                    Dim state As New CellState With {
                                                                    .WorksheetName = cell.Worksheet.Name,
                                                                    .CellAddress = cell.Address,
                                                                    .OldValue = cell.Value,
                                                                    .HadFormula = cell.HasFormula,
                                                                    .OldFormula = If(cell.HasFormula, cell.Formula, "")
                                                                }
                                    Try
                                        cell.Formula = LLMResult ' Replace cell formula
                                        undoStates.Add(state)
                                    Catch ex As Exception
                                        If ex.Message.Contains("HRESULT: 0x800A03EC") Then
                                            Try
                                                cell.FormulaLocal = LLMResult
                                                undoStates.Add(state)
                                            Catch ex2 As Exception
                                                If ex2.Message.Contains("HRESULT: 0x800A03EC") Then
                                                    Try
                                                        cell.FormulaLocal = Trim(ConvertFormulaToLocale(LLMResult, excelApp))
                                                        undoStates.Add(state)
                                                    Catch ex3 As Exception
                                                        If ex.Message.Contains("HRESULT: 0x800A03EC") Then
                                                            ShowCustomMessageBox($"Error: Excel rejected the formula '{LLMResult}' that {AN} tried to assign to the cell {cell.Address(False, False)}.")
                                                        Else
                                                            ShowCustomMessageBox($"An error occurred when trying to insert the formula '{LLMResult}' in cell {cell.Address(False, False)}: {ex.Message}")
                                                        End If
                                                    End Try
                                                Else
                                                    ShowCustomMessageBox($"An error occurred when trying to insert the formula '{LLMResult}' in cell {cell.Address(False, False)}: {ex.Message}")
                                                End If
                                            End Try
                                        Else
                                            ShowCustomMessageBox($"An error occurred when trying to insert the formula '{LLMResult}' in cell {cell.Address(False, False)}: {ex.Message}")
                                        End If
                                    End Try
                                End If
                            ElseIf Not cell.HasFormula Then
                                ' Handle plain text cells
                                SelectedText = CStr(cell.Value)

                                Dim regex As New Regex("((\r\n)|\n|\r){2,}$")
                                'Dim trailingCR As Boolean = regex.IsMatch(SelectedText)
                                'Dim trailingCR As Boolean = (SelectedText.EndsWith(vbCrLf) Or SelectedText.EndsWith(vbLf) Or SelectedText.EndsWith(vbCr))

                                If DoShorten Then
                                    Dim Textlength As Integer = SelectedText.Length
                                    ShortenLength = (Textlength - (Textlength * (100 - ShortenPercentValue) / 100))
                                    SysCommand = InterpolateAtRuntime(SysCommand)
                                End If

                                Await System.Threading.Tasks.Task.Delay(500)

                                Dim LLMResult As String = Await LLM(SysCommand, If(NoSelectedCells, "", "<TEXTTOPROCESS>" & SelectedText & "</TEXTTOPROCESS>"), "", "", 0, UseSecondAPI, True, "", FileObject)

                                If Not String.IsNullOrEmpty(LLMResult) Then
                                    LLMResult = Await PostCorrection(LLMResult, UseSecondAPI)
                                End If

                                'If Not trailingCR And LLMResult.EndsWith(ControlChars.CrLf) Then LLMResult = LLMResult.TrimEnd(ControlChars.CrLf)
                                'If Not trailingCR And LLMResult.EndsWith(ControlChars.Lf) Then LLMResult = LLMResult.TrimEnd(ControlChars.Lf)
                                'If Not trailingCR And LLMResult.EndsWith(ControlChars.Cr) Then LLMResult = LLMResult.TrimEnd(ControlChars.Cr)

                                LLMResult = Trim(LLMResult).TrimEnd(ControlChars.CrLf).TrimEnd(ControlChars.Lf).TrimEnd(ControlChars.Cr)
                                LLMResult = Trim(LLMResult).TrimEnd(ControlChars.CrLf).TrimEnd(ControlChars.Lf).TrimEnd(ControlChars.Cr)
                                LLMResult = Trim(LLMResult).TrimEnd(ControlChars.CrLf).TrimEnd(ControlChars.Lf).TrimEnd(ControlChars.Cr)

                                If Not String.IsNullOrWhiteSpace(LLMResult) Then
                                    Dim state As New CellState With {
                                                                    .WorksheetName = cell.Worksheet.Name,
                                                                    .CellAddress = cell.Address,
                                                                    .OldValue = cell.Value,
                                                                    .HadFormula = cell.HasFormula,
                                                                    .OldFormula = If(cell.HasFormula, cell.Formula, "")
                                                                }
                                    cell.Value = LLMResult ' Set the result back to the cell
                                    undoStates.Add(state)
                                End If
                            End If
                        End If

                    Catch ex As Exception
                        ' Log the error and continue with the next cell
                        Debug.WriteLine($"ProcessSelectedRange Error processing cell {cell.Address}: {ex.Message}")
                    End Try
                Next
            Finally
                'Application.ScreenUpdating = True ' Re-enable UI updates
            End Try

            splash.Close()

        Else
            Try

                If NoSelectedCells Then
                    activeSheet.Application.ActiveCell.Select()
                    selectedRange = TryCast(Globals.ThisAddIn.Application.Selection, Excel.Range)
                    If selectedRange Is Nothing Then
                        SelectedText = ""
                        Try
                            SelectedText = $"Current cell = {activeSheet.Application.ActiveCell.Address(False, False)} Text = '{activeSheet.Application.ActiveCell.Text}' Formula = '{activeSheet.Application.ActiveCell.Formula}' (use this for your output unless instructed otherwise)"
                            Debug.WriteLine("NoSelectedCell - SelectedText = " & SelectedText)
                        Catch
                        End Try
                    Else
                        NoSelectedCells = False
                    End If
                End If

                If Not NoSelectedCells Then
                    SelectedText = ConvertRangeToString(selectedRange, DoFormulas, DoColor)
                End If

                Dim RangeToInsert As String = ""

                If InsertWS = "" Then
                    RangeToInsert = "<RANGEOFCELLS>" & SelectedText & "</RANGEOFCELLS>"
                Else
                    RangeToInsert = "Currently active Worksheet: <RANGEOFCELLS>" & SelectedText & "</RANGEOFCELLS>  " & InsertWS
                End If

                If BatchFilePath = "" Then

                    Dim LLMResult As String = Await LLM(SysCommand, If(NoSelectedCells, SelectedText, RangeToInsert), "", "", 0, UseSecondAPI, False, OtherPrompt, FileObject)

                    If InsertWS = "" Then
                        LLMResult = LLMResult.Replace("<RANGEOFCELLS>", "").Replace("</RANGEOFCELLS>", "")
                    Else
                        LLMResult = Regex.Replace(LLMResult, "</?RANGEOFCELLS\d*>", "", RegexOptions.IgnoreCase)
                    End If

                    OtherPrompt = ""

                    If Not String.IsNullOrEmpty(LLMResult) Then
                        LLMResult = Await PostCorrection(LLMResult, UseSecondAPI)
                    End If

                    Dim instructions As New List(Of String)
                    instructions = ParseLLMResponse(LLMResult)

                    'Debug.WriteLine("LLMResult 1: " & LLMResult)

                    If instructions.Count > 0 Then

                        If DoPane Then
                            SP_MergePrompt_Cached = ""
                            ShowPaneAsync("The LLM has provided the following result (you can edit it):", LLMResult, $"You can let {AN} insert the square brackets into your worksheet, where possible", AN, False, True, True)
                        Else
                            Dim FinalText = ShowCustomWindow("The LLM has provided the following result (you can edit it):", LLMResult, $"Shall {AN} insert the square backets into your worksheet, where possible?", AN, False, False, False, True, Nothing, True)

                            ' Handle the user's response
                            If FinalText = "Pane" Then
                                SP_MergePrompt_Cached = ""
                                ShowPaneAsync("The LLM has provided the following result (you can edit it):", LLMResult, $"You can let {AN} insert the square brackets into your worksheet, where possible", AN, False, True, True)
                            ElseIf Not String.IsNullOrWhiteSpace(FinalText) Then
                                instructions = ParseLLMResponse(FinalText)
                                'Debug.WriteLine("FinalText: " & FinalText)
                                ApplyLLMInstructions(instructions, DoBubbles)
                                PutInClipboard(FinalText)
                                ShowCustomMessageBox("Implementation of the instructions completed (to the extent possible). They are also in the clipboard.")
                            End If
                        End If
                    Else
                        If DoPane Then
                            SP_MergePrompt_Cached = ""
                            ShowPaneAsync("The LLM has provided the following result (you can edit it):", LLMResult, "Choose to copy your edited or original text to clipboard. You can also copy & paste from the pane.", AN, False, True)
                        Else
                            Dim FinalText = ShowCustomWindow("The LLM has provided the following result (you can edit it):", LLMResult, "If you chose OK, it will be put in the clipboard.", AN)
                            If FinalText = "Pane" Then
                                SP_MergePrompt_Cached = ""
                                ShowPaneAsync("The LLM has provided the following result (you can edit it):", LLMResult, "Choose to copy your edited or original text to clipboard. You can also copy & paste from the pane.", AN, False, True)
                            ElseIf Not String.IsNullOrWhiteSpace(FinalText) Then
                                PutInClipboard(FinalText)
                            End If
                        End If
                    End If
                Else

                    Dim FileContentString As String = ""
                    Dim TempSysCommand As String = ""

                    ' Collect eligible files first (to know the maximum)
                    Dim eligibleFiles As New System.Collections.Generic.List(Of System.String)()
                    Dim DoOCR As Boolean = False
                    Dim HasPDF As Boolean = False

                    Try
                        If Not System.String.IsNullOrWhiteSpace(BatchFilePath) AndAlso System.IO.Directory.Exists(BatchFilePath) Then
                            For Each filePath As System.String In System.IO.Directory.EnumerateFiles(BatchFilePath, "*.*", System.IO.SearchOption.TopDirectoryOnly)
                                Dim ext As System.String = System.IO.Path.GetExtension(filePath)
                                If allowedExtensions.Contains(ext) Then
                                    eligibleFiles.Add(filePath)
                                    If String.Equals(ext, ".pdf", StringComparison.OrdinalIgnoreCase) Then
                                        HasPDF = True
                                    End If
                                End If
                            Next
                        Else
                            ' Directory not available; nothing to process
                            eligibleFiles.Clear()
                        End If
                    Catch ex As System.Exception
                        ' If directory enumeration itself fails, proceed with empty list
                        eligibleFiles.Clear()
                    End Try

                    If HasPDF Then
                        If Not String.IsNullOrWhiteSpace(INI_APICall_Object) Then
                            Dim answer As Integer = ShowCustomYesNoBox("The selected directory contains PDF files. Do you want to use your main model's OCR capabilities (if any) to extract text from PDFs that do not appear to contain searchable text?", "Yes, use OCR as needed", "No, do it without OCR")
                            If answer = 1 Then
                                DoOCR = True
                            ElseIf answer <> 2 Then
                                Return Nothing
                            End If
                        Else
                            Dim Answer = ShowCustomYesNoBox("The selected directory contains PDF files. However, you have not configured a model that can do OCR. Therefore, only searchable text can be extracted from the PDFs. Continue anyway?", "Yes, continue", "No, abort")
                            If Answer <> 1 Then
                                Return Nothing
                            End If
                        End If
                    End If

                        ' Progress values
                        Dim MaxEligibleFiles As System.Int32 = eligibleFiles.Count
                    Dim FileCounter As System.Int32 = 0

                    ShowProgressBarInSeparateThread($"{AN} Freestyle Batch", "Starting file processing...")
                    ProgressBarModule.CancelOperation = False
                    GlobalProgressValue = 0
                    GlobalProgressMax = MaxEligibleFiles

                    ' Main processing loop
                    For Each filePath As System.String In eligibleFiles
                        FileCounter += 1

                        If ProgressBarModule.CancelOperation Then
                            ShowCustomMessageBox("Batch processing cancelled by user.")
                            Exit For
                        End If

                        ' Update the current progress value and status label.
                        GlobalProgressValue = FileCounter
                        GlobalProgressLabel = $"Processing {FileCounter} of {MaxEligibleFiles} files..."

                        Try
                            If Not System.IO.File.Exists(filePath) Then
                                FileContentString = "Error: File not found: " & filePath
                                Continue For
                            End If

                            ' Your provided content loader
                            FileContentString = Await GetFileContent(filePath, True, DoOCR, False)

                        Catch ex As System.Exception
                            ' Do not throw; continue with next file
                            FileContentString = "File Error: " & ex.Message
                            ' Optionally log: ex.ToString()
                        End Try

                        TempSysCommand = InterpolateAtRuntime(SysCommand & " " & SP_Add_Batch)

                        Debug.WriteLine(TempSysCommand)

                        Dim LLMResult As String = Await LLM(TempSysCommand, $"File = '{System.IO.Path.GetFileName(filePath)}' <FILECONTENT>" & FileContentString & "</FILECONTENT>" & vbCrLf & "Other input to consider (if any): " & If(NoSelectedCells, SelectedText, RangeToInsert), "", "", 0, UseSecondAPI, False, OtherPrompt, FileObject)

                        If InsertWS = "" Then
                            LLMResult = LLMResult.Replace("<RANGEOFCELLS>", "").Replace("</RANGEOFCELLS>", "")
                        Else
                            LLMResult = Regex.Replace(LLMResult, "</?RANGEOFCELLS\d*>", "", RegexOptions.IgnoreCase)
                        End If

                        LLMResult = LLMResult.Replace("<FILECONTENT>", "").Replace("</FILECONTENT>", "")

                        If Not String.IsNullOrEmpty(LLMResult) Then
                            LLMResult = Await PostCorrection(LLMResult, UseSecondAPI)
                        End If

                        Dim instructions As New List(Of String)
                        instructions = ParseLLMResponse(LLMResult)

                        If instructions.Count > 0 Then
                            ApplyLLMInstructions(instructions, DoBubbles)
                        End If

                        LineNumber += 1

                    Next

                    If Not ProgressBarModule.CancelOperation Then
                        ProgressBarModule.CancelOperation = True
                        ShowCustomMessageBox($"Batch processing completed (processed {FileCounter} files).")
                    End If

                End If

            Catch ex As Exception
                MessageBox.Show("Error in Range: " & ex.Message)
            End Try

        End If

        Dim result = Globals.Ribbons.Ribbon1.UpdateUndoButton()


    End Function


    Private Async Sub ShowPaneAsync(
                          introLine As String,
                          bodyText As String,
                          finalRemark As String,
                          header As String,
                          Optional NoRtf As Boolean = False,
                          Optional insertMarkdown As Boolean = False,
                          Optional PreserveLiterals As Boolean = False
                        )
        Try
            Dim OriginalText As String = bodyText

            ' Ensure we're on the UI thread for the pane operation
            Dim result As String = ""
            If mainThreadControl.InvokeRequired Then
                result = Await mainThreadControl.Invoke(
                Function() As Task(Of String)
                    Return PaneManager.ShowMyPane(introLine, bodyText, finalRemark, header, NoRtf, insertMarkdown, New IntelligentMergeCallback(AddressOf HandleIntelligentMerge), PreserveLiterals)
                End Function
            )
            Else
                result = Await PaneManager.ShowMyPane(introLine, bodyText, finalRemark, header, NoRtf, insertMarkdown, New IntelligentMergeCallback(AddressOf HandleIntelligentMerge), PreserveLiterals)
            End If

        Catch ex As Exception
            MessageBox.Show("Error in ShowPaneAsync: " & ex.Message)
        End Try
    End Sub

    Private Sub HandleIntelligentMerge(selectedText As String)
        ' Ensure UI operations happen on the main thread
        If mainThreadControl.InvokeRequired Then
            mainThreadControl.Invoke(Sub() IntelligentMerge(selectedText))
        Else
            IntelligentMerge(selectedText)
        End If
    End Sub

    Public Async Sub IntelligentMerge(newtext As String)
        ' Ensure we're on UI thread for Excel COM operations
        If mainThreadControl.InvokeRequired Then
            mainThreadControl.Invoke(
            Sub()
                Dim instructions As New List(Of String)
                instructions = ParseLLMResponse(newtext)
                ApplyLLMInstructions(instructions, True)  ' Always DoBubbles
                ShowCustomMessageBox("Implementation of the instructions completed (to the extent possible). They are also in the clipboard.")
                Dim result = Globals.Ribbons.Ribbon1.UpdateUndoButton()
            End Sub
        )
        Else
            Dim instructions As New List(Of String)
            instructions = ParseLLMResponse(newtext)
            ApplyLLMInstructions(instructions, True)  ' Always DoBubbles
            ShowCustomMessageBox("Implementation of the instructions completed (to the extent possible). They are also in the clipboard.")
            Dim result = Globals.Ribbons.Ribbon1.UpdateUndoButton()
        End If
    End Sub


    Private Async Sub oldShowPaneAsync(
                              introLine As String,
                              bodyText As String,
                              finalRemark As String,
                              header As String,
                              Optional NoRtf As Boolean = False,
                              Optional insertMarkdown As Boolean = False,
                              Optional PreserveLiterals As Boolean = False
                            )
        Try

            Dim OriginalText As String = bodyText

            Dim result As String = Await PaneManager.ShowMyPane(introLine, bodyText, finalRemark, header, NoRtf, insertMarkdown, New IntelligentMergeCallback(AddressOf HandleIntelligentMerge), PreserveLiterals)

        Catch ex As Exception
            MessageBox.Show("Error in ShowPaneAsync: " & ex.Message)
        End Try
    End Sub


    Private Sub oldHandleIntelligentMerge(selectedText As String)
        IntelligentMerge(selectedText)
    End Sub

    Public Async Sub oldIntelligentMerge(newtext As String)
        Dim instructions As New List(Of String)
        instructions = ParseLLMResponse(newtext)
        ApplyLLMInstructions(instructions, True)  ' Always DoBubbles
        ShowCustomMessageBox("Implementation of the instructions completed (to the extent possible). They are also in the clipboard.")
        Dim result = Globals.Ribbons.Ribbon1.UpdateUndoButton()
    End Sub


    Public Function GatherSelectedWorksheets(
    Optional ByVal includeActiveWorksheet As System.Boolean = False
) As System.String
        Try
            Dim app As Microsoft.Office.Interop.Excel.Application =
            Globals.ThisAddIn.Application
            Dim activeWs As Microsoft.Office.Interop.Excel.Worksheet =
            TryCast(app.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)

            ' build list of worksheets (optionally including the active one)
            Dim sheetList As New System.Collections.Generic.List(
            Of Microsoft.Office.Interop.Excel.Worksheet)()
            Dim selItems As New System.Collections.Generic.List(Of SelectionItem)()

            For Each wb As Microsoft.Office.Interop.Excel.Workbook In app.Workbooks
                For Each ws As Microsoft.Office.Interop.Excel.Worksheet In wb.Worksheets
                    If includeActiveWorksheet OrElse ws IsNot activeWs Then
                        sheetList.Add(ws)
                        selItems.Add(New SelectionItem(
                        $"{ws.Name} ({wb.FullName})",
                        sheetList.Count))
                    End If
                Next
            Next

            ' if no sheets matched the filter, bail
            If sheetList.Count = 0 Then Return "NONE"

            ' add “All worksheets …” option
            Dim allOptionIndex As System.Int32 = selItems.Count + 1
            Dim allOptionText As System.String = If(
            includeActiveWorksheet,
            "Add all worksheets",
            "Add all other worksheets")
            selItems.Add(New SelectionItem(allOptionText, allOptionIndex))

            ' prompt user
            Dim itemsArray As SelectionItem() = selItems.ToArray()
            Dim picked As System.Int32 = SelectValue(itemsArray, allOptionIndex, "Choose worksheet to add …")
            If picked < 1 Then Return System.String.Empty

            ' determine targets
            Dim targets As New System.Collections.Generic.List(
            Of Microsoft.Office.Interop.Excel.Worksheet)()
            If picked = allOptionIndex Then
                targets.AddRange(sheetList)
            Else
                targets.Add(sheetList(picked - 1))
            End If

            ' build the result string
            Dim InsertedWorksheet As System.String = System.String.Empty
            Dim tagIndex As System.Int32 = 2
            For Each ws As Microsoft.Office.Interop.Excel.Worksheet In targets
                InsertedWorksheet &= $"<RANGEOFCELLS{tagIndex}>" & vbCrLf

                ' now just call your converter (which itself inserts the sheet/file header)
                InsertedWorksheet &= ConvertRangeToString(
                CellRange:=CType(ws.UsedRange, Microsoft.Office.Interop.Excel.Range),
                IncludeFormulas:=True,
                DoColor:=False,
                TargetWorksheet:=ws) & vbCrLf

                InsertedWorksheet &= $"</RANGEOFCELLS{tagIndex}>" & vbCrLf
                tagIndex += 1
            Next

            If System.String.IsNullOrEmpty(InsertedWorksheet) Then
                ShowCustomMessageBox("No content could be retrieved from the selected worksheet(s).")
                Return System.String.Empty
            End If

            Return InsertedWorksheet

        Catch ex As System.Exception
            Return "ERROR " & ex.Message
        End Try
    End Function


    Public Function oldGatherSelectedWorksheets() As String
        Try
            Dim app As Microsoft.Office.Interop.Excel.Application =
            Globals.ThisAddIn.Application
            Dim activeWs As Microsoft.Office.Interop.Excel.Worksheet =
            TryCast(app.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)

            ' build list of all worksheets except the active one
            Dim sheetList As New List(Of Microsoft.Office.Interop.Excel.Worksheet)()
            Dim selItems As New List(Of SelectionItem)()
            For Each wb As Microsoft.Office.Interop.Excel.Workbook In app.Workbooks
                For Each ws As Microsoft.Office.Interop.Excel.Worksheet In wb.Worksheets
                    If Not ws Is activeWs Then
                        sheetList.Add(ws)
                        selItems.Add(New SelectionItem(
                        $"{ws.Name} ({wb.FullName})",
                        sheetList.Count))
                    End If
                Next
            Next

            ' if no other sheets, bail
            If sheetList.Count = 0 Then Return "NONE"

            ' add “All worksheets, except current”
            Dim allExceptIndex As Integer = selItems.Count + 1
            selItems.Add(New SelectionItem("Add all other worksheets", allExceptIndex))

            ' prompt user
            Dim itemsArray = selItems.ToArray()
            Dim picked As Integer = SelectValue(itemsArray, allExceptIndex, "Choose worksheet to add …")
            If picked < 1 Then Return String.Empty

            ' determine targets
            Dim targets As New List(Of Microsoft.Office.Interop.Excel.Worksheet)()
            If picked = allExceptIndex Then
                targets.AddRange(sheetList)
            Else
                targets.Add(sheetList(picked - 1))
            End If

            ' build the result string
            Dim InsertedWorksheet As String = String.Empty
            Dim tagIndex As Integer = 2
            For Each ws In targets
                InsertedWorksheet &= $"<RANGEOFCELLS{tagIndex}>" & vbCrLf

                ' now just call your converter (which itself inserts the sheet/file header)
                InsertedWorksheet &= ConvertRangeToString(
                CellRange:=CType(ws.UsedRange, Microsoft.Office.Interop.Excel.Range),
                IncludeFormulas:=True,
                DoColor:=False,
                TargetWorksheet:=ws
            ) & vbCrLf

                InsertedWorksheet &= $"</RANGEOFCELLS{tagIndex}>" & vbCrLf
                tagIndex += 1
            Next

            If String.IsNullOrEmpty(InsertedWorksheet) Then
                ShowCustomMessageBox("No content could be retrieved from the selected worksheet(s).")
                Return String.Empty
            End If

            Return InsertedWorksheet

        Catch ex As System.Exception
            Return "ERROR " & ex.Message
        End Try
    End Function



    ' Helpers for the Range Functionality

    Public Function ConvertRangeToString(
    ByVal CellRange As Excel.Range,
    ByVal IncludeFormulas As Boolean,
    Optional ByVal DoColor As Boolean = False,
     Optional ByVal TargetWorksheet As Microsoft.Office.Interop.Excel.Worksheet = Nothing
) As String


        Dim splash As New SplashScreen("Gathering the content from your worksheet...")
        splash.Show()
        splash.Refresh()

        If CellRange Is Nothing AndAlso TargetWorksheet Is Nothing Then
            splash.Close()
            Return String.Empty
        End If

        ' Excel-UI abschalten
        Dim app As Excel.Application = Globals.ThisAddIn.Application
        Dim origWb As Microsoft.Office.Interop.Excel.Workbook = app.ActiveWorkbook
        Dim origWs As Microsoft.Office.Interop.Excel.Worksheet = TryCast(app.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)

        ' if caller specified a sheet, make sure it’s active
        If TargetWorksheet IsNot Nothing Then
            TargetWorksheet.Parent.Activate()
            TargetWorksheet.Activate()
        End If

        If CellRange Is Nothing AndAlso TargetWorksheet IsNot Nothing Then
            CellRange = TargetWorksheet.UsedRange
        End If

        Dim sb As New System.Text.StringBuilder()
        If TargetWorksheet IsNot Nothing Then
            sb.AppendLine($"From Worksheet: {TargetWorksheet.Name}, File: {TargetWorksheet.Parent.FullName}")
        Else
            sb.AppendLine($"From Worksheet {CellRange.Worksheet.Name}, File {CellRange.Worksheet.Parent.FullName}")
        End If

        With app
            .ScreenUpdating = False
            .EnableEvents = False
            .Calculation = Excel.XlCalculation.xlCalculationManual
        End With

        ' Standardfarben des "Normal"-Styles ermitteln
        Dim normalStyle As Excel.Style = app.ActiveWorkbook.Styles("Normal")
        Dim defaultFontColor As Long = normalStyle.Font.Color
        Dim defaultInteriorColor As Long = normalStyle.Interior.Color

        Try
            ' Werte auslesen
            Dim rawVals As Object = CellRange.Value2
            Dim vals(,) As Object

            If TypeOf rawVals Is Object(,) Then
                vals = CType(rawVals, Object(,))
            Else
                ReDim vals(0, 0)
                vals(0, 0) = rawVals
            End If

            Dim rowLB As Integer = vals.GetLowerBound(0)
            Dim rowUB As Integer = vals.GetUpperBound(0)
            Dim colLB As Integer = vals.GetLowerBound(1)
            Dim colUB As Integer = vals.GetUpperBound(1)

            For r As Integer = rowLB To rowUB
                For c As Integer = colLB To colUB
                    Dim raw = vals(r, c)

                    Dim relativeRow As Integer = r - rowLB + 1
                    Dim relativeCol As Integer = c - colLB + 1
                    Dim cell As Excel.Range = CellRange.Cells(relativeRow, relativeCol)
                    Dim addr As String = cell.Address(False, False)

                    ' Prüfen, ob wir diese Zelle *trotz* leerem Wert verarbeiten müssen
                    Dim shouldProcess As Boolean = False

                    ' 1) Ist ein Wert drin?
                    If raw IsNot Nothing Then
                        shouldProcess = True
                    End If

                    ' 2) Klassischer Kommentar?
                    If Not shouldProcess AndAlso cell.Comment IsNot Nothing Then
                        shouldProcess = True
                    End If

                    ' 3) Threaded Comment?
                    If Not shouldProcess Then
                        Try
                            Dim tc = CType(cell, Object).CommentThreaded
                            If tc IsNot Nothing Then shouldProcess = True
                        Catch ex As COMException When ex.ErrorCode = &H800A03EC
                            ' kein Support / kein ThreadedComment
                        End Try
                    End If

                    ' 4) DataValidation-List?
                    If Not shouldProcess Then
                        Try
                            If cell.Validation.Type = Excel.XlDVType.xlValidateList Then
                                shouldProcess = True
                            End If
                        Catch
                            ' keine Validierung
                        End Try
                    End If

                    ' 5) Farbe nur bei DoColor?
                    'If Not shouldProcess AndAlso DoColor Then
                    '                   If cell.Font.Color <> defaultFontColor OrElse cell.Interior.Color <> defaultInteriorColor Then
                    'shouldProcess = True
                    'End If
                    'End If

                    ' 5) Farbe nur bei DoColor?
                    If Not shouldProcess AndAlso DoColor Then
                        ' Include empty cells only when they actually have a background fill (not "No Fill").
                        ' Rationale: normal/transparent background is ColorIndex = xlColorIndexNone.
                        Dim hasFill As Boolean = False
                        Try
                            hasFill = (cell.Interior.ColorIndex <> Excel.XlColorIndex.xlColorIndexNone)
                        Catch
                            ' Some Excel versions/situations may throw on ColorIndex; ignore and treat as no fill.
                        End Try
                        If hasFill Then
                            shouldProcess = True
                        End If
                    End If

                    If shouldProcess Then

                        Try
                            sb.AppendLine($"Cell {addr} has")
                            sb.AppendLine($"- Value {raw}")

                            ' Formeln optional auslesen
                            If IncludeFormulas AndAlso cell.HasFormula Then
                                Dim f As String = String.Empty
                                Try
                                    f = cell.Formula2.ToString()
                                Catch comEx As System.Runtime.InteropServices.COMException _
                                When comEx.ErrorCode = &H800A03EC
                                    f = cell.Formula.ToString()
                                End Try
                                sb.AppendLine($"- Formula {If(String.IsNullOrEmpty(f), "none", f)}")
                            End If

                            ' Kommentare (klassisch)
                            If cell.Comment IsNot Nothing Then
                                sb.AppendLine($"- Comment {cell.Comment.Text()}")
                            End If

                            ' ThreadedComments per the Excel object model
                            Dim cellObj As Object = cell   ' late-bind the Range
                            Try
                                ' try to get the single threaded comment
                                Dim topObj As Object = cellObj.CommentThreaded
                                If topObj IsNot Nothing Then
                                    ' .Text
                                    Dim txt = topObj.Text
                                    ' .Author.Name
                                    Dim authName = topObj.Author.Name
                                    sb.AppendLine($"- Threaded comment {txt} (by {authName})")

                                    ' now each reply
                                    For Each rep In topObj.Replies  ' an IEnumerable
                                        sb.AppendLine($"- Reply comment {rep.Text} (by {rep.Author.Name})")
                                    Next
                                End If
                            Catch ex As System.Runtime.InteropServices.COMException When ex.ErrorCode = &H800A03EC
                                ' no threaded-comments support—ignore
                            End Try

                            ' 1) Drop-Down-Auswahl (DataValidation mit Named-Range-Support)
                            Try
                                ' 1a) Prüfen, ob die Zelle überhaupt eine Listen-Validation hat
                                Dim hasList As Boolean
                                Try
                                    hasList = (cell.Validation.Type = Excel.XlDVType.xlValidateList)
                                Catch comEx As COMException When comEx.ErrorCode = &H800A03EC
                                    hasList = False
                                End Try

                                If hasList Then
                                    Dim formula1 As String = cell.Validation.Formula1  ' z.B. "=MyList" oder "=Sheet2!$A$1$A$5"
                                    If Not String.IsNullOrWhiteSpace(formula1) Then
                                        Dim options As New List(Of String)()
                                        Dim wb As Microsoft.Office.Interop.Excel.Workbook = cell.Worksheet.Parent
                                        Dim refRange As Excel.Range = Nothing

                                        ' 1b) Versuch: Workbook-scoped Named Range
                                        If formula1.StartsWith("="c) Then
                                            Dim nameKey As String = formula1.Substring(1)  ' ohne "="
                                            Try
                                                Dim nm As Excel.Name = wb.Names(nameKey)
                                                refRange = nm.RefersToRange
                                            Catch ex As Exception
                                                ' kein Named Range gefunden → weiter unten
                                            End Try
                                        End If

                                        ' 1c) Fallback: Worksheet-lokaler Bereich oder sheet-qualifizierte Adresse
                                        If refRange Is Nothing AndAlso formula1.StartsWith("="c) Then
                                            Dim addrx As String = formula1.Substring(1)
                                            Try
                                                refRange = cell.Worksheet.Range(addrx)
                                            Catch ex1 As COMException
                                                ' vielleicht sheet-qualifiziert: "Sheet2!$B$1$B$10"
                                                Dim parts = addrx.Split("!"c)
                                                Dim sheetName = parts(0)
                                                Dim rangeAddr = parts(1)
                                                Dim otherSheet As Microsoft.Office.Interop.Excel.Worksheet = wb.Sheets(sheetName)
                                                refRange = otherSheet.Range(rangeAddr)
                                            End Try
                                        End If

                                        ' 1d) Werte aus dem ermittelten Bereich auslesen (nested loop)
                                        If refRange IsNot Nothing Then
                                            Dim tmp = refRange.Value2
                                            If TypeOf tmp Is Object(,) Then
                                                Dim arr = CType(tmp, Object(,))
                                                Dim rCount = arr.GetLength(0)
                                                Dim cCount = arr.GetLength(1)
                                                For rx As Integer = 1 To rCount
                                                    For cx As Integer = 1 To cCount
                                                        Dim v = arr(rx, cx)
                                                        options.Add(If(v Is Nothing, String.Empty, v.ToString()))
                                                    Next
                                                Next
                                            ElseIf tmp IsNot Nothing Then
                                                options.Add(tmp.ToString())
                                            End If
                                            Marshal.ReleaseComObject(refRange)
                                        Else
                                            ' 1e) Kein Range-Objekt gefunden? Dann nochmal die Fallback-Komma-Liste
                                            If formula1.StartsWith("="c) Then
                                                ' remove leading "=" before splitting
                                                options.AddRange(formula1.Substring(1) _
                                                .Split(","c) _
                                                .Select(Function(s) s.Trim()))
                                            Else
                                                ' pure comma-list, komplett splitten
                                                options.AddRange(formula1 _
                                                .Split(","c) _
                                                .Select(Function(s) s.Trim()))
                                            End If
                                        End If

                                        sb.AppendLine($"- Dropdown options (separated by §) {String.Join("§", options)}")
                                    End If
                                End If

                            Catch ex As Exception
                                sb.AppendLine($"- Error reading dropdown {ex.Message}")
                            End Try

                            ' 2) Farben (nur bei Abweichung)
                            If DoColor Then
                                ' Emit LLM-friendly CSS hex (#RRGGBB) plus rgb tuple.
                                If cell.Font.Color <> defaultFontColor Then
                                    Try
                                        Dim fc As System.Drawing.Color = System.Drawing.ColorTranslator.FromOle(CInt(cell.Font.Color))
                                        sb.AppendLine($"- FontColor #{fc.R:X2}{fc.G:X2}{fc.B:X2} (rgb {fc.R},{fc.G},{fc.B})")
                                    Catch
                                        ' Fallback: still emit the raw value if conversion fails (rare).
                                        sb.AppendLine($"- FontColor {cell.Font.Color}")
                                    End Try
                                End If

                                If cell.Interior.Color <> defaultInteriorColor Then
                                    Try
                                        Dim bc As System.Drawing.Color = System.Drawing.ColorTranslator.FromOle(CInt(cell.Interior.Color))
                                        sb.AppendLine($"- BackgroundColor #{bc.R:X2}{bc.G:X2}{bc.B:X2} (rgb {bc.R},{bc.G},{bc.B})")
                                    Catch
                                        sb.AppendLine($"- BackgroundColor {cell.Interior.Color}")
                                    End Try
                                End If
                            End If

                            sb.AppendLine(New String("-"c, 5))

                        Catch ex As System.Runtime.InteropServices.COMException _
                        When ex.ErrorCode = &H800A03EC
                            sb.AppendLine($"- COM-Error in Cell {addr} {ex.Message}")
                        Catch ex As System.Exception
                            sb.AppendLine($"- Error in Cell {addr} {ex.Message}")
                        Finally
                            Marshal.ReleaseComObject(cell)
                        End Try
                    End If
                    Marshal.ReleaseComObject(cell)
                Next
            Next

        Finally
            ' Excel-UI wieder aktivieren
            With app
                .ScreenUpdating = True
                .EnableEvents = True
                .Calculation = Excel.XlCalculation.xlCalculationAutomatic
            End With

            If origWb IsNot Nothing Then origWb.Activate()
            If origWs IsNot Nothing Then origWs.Activate()

            splash.Close()
        End Try

        Return sb.ToString()
    End Function


    Public Function GetFileName() As String
        Dim filePath As String = ""
        Try
            If String.IsNullOrWhiteSpace(filePath) Then
                Using form As New DragDropForm()
                    If form.ShowDialog() = DialogResult.OK Then
                        filePath = form.SelectedFilePath
                    Else
                        ' User cancelled or closed form
                        Return String.Empty
                    End If
                End Using
            End If

            filePath = RemoveCR(filePath.Trim())
            filePath = Path.GetFullPath(filePath)
            If Not File.Exists(filePath) Then
                ShowCustomMessageBox($"The file '{filePath}' was not found.")
                Return ""
        End If
        Return filePath

        Catch ex As System.Exception
            ShowCustomMessageBox($"An error occurred reading the file '{filePath}': {ex.Message}")
            Return ""
        End Try
    End Function



    Public Async Function GetFileContent(Optional ByVal optionalFilePath As String = Nothing, Optional Silent As Boolean = False, Optional DoOCR As Boolean = False, Optional AskUser As Boolean = True) As Task(Of String)
        Dim filePath As String = ""
        Try

            If optionalFilePath IsNot Nothing Then
                filePath = ExpandEnvironmentVariables(optionalFilePath)
            End If

            If String.IsNullOrWhiteSpace(filePath) Then
                Using form As New DragDropForm()
                    If form.ShowDialog() = DialogResult.OK Then
                        filePath = form.SelectedFilePath
                    Else
                        ' User cancelled or closed form
                        Return String.Empty
                    End If
                End Using
            End If

            filePath = RemoveCR(filePath.Trim())
            filePath = Path.GetFullPath(filePath)
            If Not File.Exists(filePath) Then
                If Not Silent Then ShowCustomMessageBox($"The file '{filePath}' was not found.")
                Return ""
            End If

            If Not String.IsNullOrWhiteSpace(filePath) AndAlso IO.File.Exists(filePath) Then
                Dim ext As String = IO.Path.GetExtension(filePath).ToLowerInvariant()
                Dim FromFile As String
                Select Case ext
                    Case ".txt", ".ini", ".csv", ".log", ".json", ".xml", ".html", ".htm"
                        FromFile = ReadTextFile(filePath)
                    Case ".rtf"
                        FromFile = ReadRtfAsText(filePath)
                    Case ".doc", ".docx"
                        FromFile = ReadWordDocument(filePath)
                    Case ".pdf"
                        FromFile = Await ReadPdfAsText(filePath, True, DoOCR, AskUser, _context)
                    Case Else
                        FromFile = "Error: File type not supported."
                End Select
                If FromFile.StartsWith("Error") And Len(FromFile) < 100 And Not Silent Then
                    ShowCustomMessageBox(FromFile)
                    Return ""
                Else
                    Return FromFile
                End If
            End If
        Catch ex As System.Exception
            If Not Silent Then ShowCustomMessageBox($"An error occurred reading the file '{filePath}': {ex.Message}")
            Return ""
        End Try
    End Function

    Public Function ParseLLMResponse(ByVal Response As String) As List(Of String)
        Dim instructions As New List(Of String)()
        Dim startPos As Integer, instructionEnd As Integer
        Dim tempInstruction As String
        Dim cellPattern As String

        ' Ensure we remove any newlines that might affect parsing
        Response = Response.Replace(vbCrLf, " ").Replace(vbLf, " ")

        ' Pattern for finding Cell
        cellPattern = "[Cell:"

        ' Start parsing the response
        startPos = Response.IndexOf(cellPattern, StringComparison.OrdinalIgnoreCase)

        Do While startPos >= 0
            ' Find next cell occurrence to extract the block between this and next [Cell:]
            instructionEnd = Response.IndexOf(cellPattern, startPos + cellPattern.Length, StringComparison.OrdinalIgnoreCase)

            ' If there's no further [Cell:], capture till the end of the string
            If instructionEnd = -1 Then instructionEnd = Response.Length

            ' Extract the instruction block between the current and next [Cell:]
            tempInstruction = Response.Substring(startPos, instructionEnd - startPos)

            ' Only keep blocks that contain a top-level action (ignore nested brackets inside other brackets)
            Dim extracted As String = GetFormulaOrValueFromInstruction(tempInstruction)
            Dim hasAction As Boolean = Not String.IsNullOrWhiteSpace(extracted) OrElse
                                       tempInstruction.IndexOf(AN5 & ":", StringComparison.OrdinalIgnoreCase) >= 0

            If hasAction Then
                instructions.Add(tempInstruction)
            End If

            ' Move to the next instruction start, exit if at the end
            startPos = Response.IndexOf(cellPattern, instructionEnd, StringComparison.OrdinalIgnoreCase)
        Loop

        Return instructions
    End Function

    Public Function oldParseLLMResponse(ByVal Response As String) As List(Of String)
        Dim instructions As New List(Of String)()
        Dim startPos As Integer, instructionEnd As Integer
        Dim tempInstruction As String
        Dim cellPattern As String

        ' Ensure we remove any newlines that might affect parsing
        Response = Response.Replace(vbCrLf, " ").Replace(vbLf, " ")

        ' Pattern for finding Cell
        cellPattern = "[Cell:"

        ' Start parsing the response
        startPos = Response.IndexOf(cellPattern)

        Do While startPos >= 0
            ' Find next cell occurrence to extract the block between this and next [Cell:]
            instructionEnd = Response.IndexOf(cellPattern, startPos + cellPattern.Length)

            ' If there's no further [Cell:], capture till the end of the string
            If instructionEnd = -1 Then instructionEnd = Response.Length

            ' Extract the instruction block between the current and next [Cell:]
            tempInstruction = Response.Substring(startPos, instructionEnd - startPos)
            instructions.Add(tempInstruction)

            ' Move to the next instruction start, exit if at the end
            startPos = Response.IndexOf(cellPattern, instructionEnd)
        Loop

        Return instructions
    End Function


    Private Function GetActiveWorksheetSafe(app As Excel.Application) As Worksheet
        ' Try the currently active sheet
        Dim ws = TryCast(app.ActiveSheet, Worksheet)
        If ws IsNot Nothing Then Return ws

        ' Fallback: first worksheet of first open workbook
        For Each wb As Workbook In app.Workbooks
            If wb IsNot Nothing AndAlso wb.Worksheets.Count > 0 Then
                Return CType(wb.Worksheets(1), Worksheet)
            End If
        Next

        Return Nothing
    End Function


    Sub ApplyLLMInstructions(ByVal instructions As List(Of String), DoAlsoBubbles As Boolean)

        Dim instruction As String
        Dim cellAddress As String
        Dim formulaOrValue As String
        Dim formulaOrValueLocale As String = ""
        Dim cleanedValue As String
        Dim ii As Integer

        ' Get the active Excel application and sheet
        Dim excelApp As Excel.Application = Globals.ThisAddIn.Application
        Dim activeSheet As Worksheet = GetActiveWorksheetSafe(excelApp)

        If activeSheet Is Nothing Then
            ShowCustomMessageBox("No worksheet available to apply instructions.")
            Exit Sub
        End If

        ii = 0

        undoStates.Clear()

        Dim splash As New SplashScreen("Implementing... press 'Esc' to abort")
        splash.Show()
        splash.Refresh()

        Debug.WriteLine("Instructions: " & String.Join(Environment.NewLine, instructions))

        Dim prevAutoFillFormulasInLists As Boolean = excelApp.AutoCorrect.AutoFillFormulasInLists
        Dim prevAutoExpandListRange As Boolean = excelApp.AutoCorrect.AutoExpandListRange
        Dim prevEnableAutoComplete As Boolean = excelApp.EnableAutoComplete
        Dim prevExtendList As Boolean = excelApp.ExtendList

        excelApp.AutoCorrect.AutoFillFormulasInLists = False
        excelApp.AutoCorrect.AutoExpandListRange = False
        excelApp.EnableAutoComplete = False
        excelApp.ExtendList = False

        Try
            ' Loop through the parsed instructions and ask for confirmation before applying
            For Each instruction In instructions

                If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And &H8000) <> 0 Then Exit For
                If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And 1) <> 0 Then Exit For

                cellAddress = GetCellFromInstruction(instruction)
                formulaOrValue = GetFormulaOrValueFromInstruction(instruction)

                'If Not String.IsNullOrWhiteSpace(cellAddress) AndAlso Not String.IsNullOrWhiteSpace(formulaOrValue) Then
                If Not String.IsNullOrWhiteSpace(cellAddress) Then
                    ii += 1
                    Debug.WriteLine($"Processing: Cell='{cellAddress}', Value='{formulaOrValue}'")

                    Try
                        If activeSheet IsNot Nothing AndAlso activeSheet.Range(cellAddress) IsNot Nothing Then
                            Dim targetRange As Range
                            Try
                                ' Ensure the address is valid before accessing it
                                If Regex.IsMatch(cellAddress, "^[A-Z]+\d+$") Then
                                    targetRange = activeSheet.Range(cellAddress)

                                    ' Store the state BEFORE any changes
                                    Dim state As New CellState With {
                                        .WorksheetName = targetRange.Worksheet.Name,
                                        .CellAddress = targetRange.Address,
                                        .OldValue = targetRange.Value,
                                        .HadFormula = targetRange.HasFormula,
                                        .OldFormula = If(targetRange.HasFormula, targetRange.Formula, "")
                                    }

                                    ' Handle merged cells properly
                                    If targetRange.MergeCells Then
                                        targetRange = targetRange.MergeArea.Cells(1, 1)
                                    End If

                                    ' Add the state to undoStates - do this BEFORE making changes
                                    undoStates.Add(state)

                                    ' Handle merged cells properly
                                    If targetRange.MergeCells Then
                                        targetRange = targetRange.MergeArea.Cells(1, 1)
                                    End If

                                    If formulaOrValue.StartsWith($"{AN5}: ") Then
                                        'If DoAlsoBubbles And formulaOrValue.StartsWith($"{AN5}: ") Then

                                        ' Add a comment to the cell
                                        Dim commentText As String = formulaOrValue.Trim()
                                        If commentText <> $"{AN5}: " Then
                                            If targetRange.CommentThreaded Is Nothing Then
                                                targetRange.AddCommentThreaded(Text:=$"{commentText}")
                                            Else
                                                targetRange.CommentThreaded.AddReply(Text:=$"{commentText}")
                                            End If
                                        End If

                                    ElseIf formulaOrValue.StartsWith("=") Then

                                        ' Fix cell format issues
                                        targetRange.Value = ""
                                        targetRange.NumberFormat = "General"

                                        SetFormulaSafe(targetRange, formulaOrValue, excelApp)

                                    Else

                                        ' Assign values properly
                                        If IsNumeric(formulaOrValue) Then
                                            targetRange.Value = formulaOrValue
                                        Else
                                            ' Remove unwanted apostrophes
                                            cleanedValue = CStr(formulaOrValue).Trim("'"c)
                                            cleanedValue = DecodeTextLiterals(cleanedValue)
                                            targetRange.NumberFormat = "@"
                                            Debug.WriteLine($"Set cleaned text value in {cellAddress}: '{cleanedValue}'")
                                            targetRange.Value = cleanedValue
                                        End If

                                    End If
                                Else
                                    Debug.WriteLine($"Invalid cell address: {cellAddress}")
                                End If
                            Catch ex As Exception
                                If ex.Message.Contains("HRESULT: 0x800A03EC") Then
                                    ShowCustomMessageBox($"Error: Excel rejected the formula or value '{formulaOrValue}' that {AN} tried to assign to the cell {cellAddress}.")
                                Else
                                    ShowCustomMessageBox($"An error occurred when trying to insert the formula or value '{formulaOrValue}' in cell {cellAddress}: {ex.Message}")
                                End If
                            End Try
                        Else
                            Debug.WriteLine($"Invalid or missing cell address: {cellAddress}")
                        End If
                    Catch ex As Exception
                        If ex.Message.Contains("HRESULT: 0x800A03EC") Then
                            ShowCustomMessageBox($"Error: Excel rejected the formula '{formulaOrValue}' that {AN} tried to assign to the cell {cellAddress}.")
                        Else
                            ShowCustomMessageBox($"An error occurred when trying to insert the formula '{formulaOrValue}' in cell {cellAddress}: {ex.Message}")
                        End If
                    End Try
                End If
            Next

        Finally
            ' --- Always restore Excel settings, even if the loop exits early or errors ---
            excelApp.AutoCorrect.AutoFillFormulasInLists = prevAutoFillFormulasInLists
            excelApp.AutoCorrect.AutoExpandListRange = prevAutoExpandListRange
            excelApp.EnableAutoComplete = prevEnableAutoComplete
            excelApp.ExtendList = prevExtendList
        End Try
        splash.Close()

    End Sub


    Private Function PreSanitizeNonJson(raw As String) As String
        If String.IsNullOrWhiteSpace(raw) Then Return raw

        ' 1. Convert RTF unicode escapes: \u####?  (decimal to hex → actual char)
        raw = Regex.Replace(raw, "\\u(-?\d+)\?", Function(m)
                                                     Dim dec = Integer.Parse(m.Groups(1).Value)
                                                     Return ChrW(dec)
                                                 End Function)

        ' 2. Replace RTF line markers
        raw = Regex.Replace(raw, "(?i)\\line", vbCrLf)

        ' 3. Collapse excessive backslashes that are not valid JSON escapes (optional)
        ' Leave valid JSON escapes (\n, \r, \t, \") alone
        raw = Regex.Replace(raw, "\\\\(?![nrt""\\/u])", "\")

        ' 4. Strip leading / trailing single quotes the LLM sometimes wraps around payload
        raw = raw.Trim()
        If raw.StartsWith("'") AndAlso raw.EndsWith("'") AndAlso raw.Length > 2 Then
            raw = raw.Substring(1, raw.Length - 2).Trim()
        End If

        ' 5. Remove zero-width or control chars except CR/LF/TAB
        raw = New String(raw.Where(Function(c) AscW(c) >= 32 OrElse c = vbCr OrElse c = vbLf OrElse c = vbTab).ToArray())

        Return raw
    End Function


    Public Sub SetFormulaSafe(cell As Excel.Range, formulaOrValue As String, excelApp As Excel.Application)
        ' 0. Hol Dir den Listentrenner (in DE ist das ";")
        Dim localSep As String = excelApp.International(XlApplicationInternational.xlListSeparator)

        ' 1. Unser "englischer" Ausgangs-String
        Dim englishFormula As String = formulaOrValue

        Try
            ' 2. Erstversuch: dynamic-array-Formel in Englisch
            Try
                cell.Formula2 = englishFormula
            Catch ex1 As System.Runtime.InteropServices.COMException When ex1.ErrorCode = &H800A03EC
                ' 0x800A03EC = Locale-Error
                ' → versuche gleich mit Formula2Local
                Try
                    cell.Formula2Local = englishFormula
                Catch ex2 As System.Runtime.InteropServices.COMException
                    ' ignorieren, kommt unten nochmal dran
                End Try
            End Try

            ' 3. Wenn #NAME? drinsteht, nochmal mit FormulaLocal und lokalem Trenner
            If cell.HasFormula AndAlso Trim(cell.Text.ToString()) = "#NAME?" Then
                Try
                    cell.FormulaLocal = englishFormula.Replace(",", localSep)
                Catch ex3 As System.Runtime.InteropServices.COMException
                    ' ignorieren
                End Try

                ' 4. Wenn immer noch #NAME? → Namen übersetzen lassen
                If Trim(cell.Text.ToString()) = "#NAME?" Then
                    Dim converted As String = Trim(ConvertFormulaToLocale(englishFormula, excelApp))
                    If Not String.IsNullOrEmpty(converted) Then
                        converted = converted.Replace(",", localSep)
                        Try
                            cell.FormulaLocal = converted
                        Catch ex4 As System.Runtime.InteropServices.COMException
                            ShowCustomMessageBox($"Failed to set converted formula: {ex4.Message}")
                        End Try
                    End If

                    ' 5. Letzter Check
                    If Trim(cell.Text.ToString()) = "#NAME?" Then
                        ShowCustomMessageBox(
                        $"Excel rejected the formula '{englishFormula}' for cell {cell.Address}. Resulted in #NAME?."
                    )
                    End If
                End If
            End If

            ' 6. Genereller COM-Fehler
        Catch comEx As System.Runtime.InteropServices.COMException
            ShowCustomMessageBox($"COM Error setting formula: {comEx.Message}")

            ' 7. Alle anderen Fehler
        Catch ex As System.Exception
            ShowCustomMessageBox($"General error setting formula: {ex.Message}")
        End Try
    End Sub



    Public Function ConvertFormulaToLocale(ByVal englishFormula As String, ByVal excelApp As Excel.Application) As String

        Dim wb As Workbook = Nothing
        Dim ws As Worksheet = Nothing
        Dim localizedFormula As String = ""

        ' Disable screen updating to prevent Excel from flashing
        Dim previousScreenUpdating As Boolean = excelApp.ScreenUpdating
        Dim previousDisplayAlerts As Boolean = excelApp.DisplayAlerts

        Try
            excelApp.ScreenUpdating = False ' Hide UI updates
            excelApp.DisplayAlerts = False ' Prevent pop-ups (e.g., when closing the temp workbook)

            ' Create a temporary worksheet
            wb = excelApp.Workbooks.Add()
            ws = CType(wb.Sheets(1), Worksheet)

            ' Set the formula using English syntax
            Dim tempRange As Excel.Range = ws.Range("A1")
            tempRange.Formula = englishFormula

            ' Retrieve the formula in the local Excel language
            localizedFormula = tempRange.FormulaLocal

            ' Close workbook without saving
            wb.Close(False)
        Catch ex As Exception
            localizedFormula = englishFormula ' Fallback to English if an error occurs
        Finally
            ' Restore Excel's UI settings
            excelApp.ScreenUpdating = previousScreenUpdating
            excelApp.DisplayAlerts = previousDisplayAlerts

            ' Release COM objects
            ReleaseObject(ws)
            ReleaseObject(wb)
        End Try

        Return localizedFormula
    End Function


    Public Function SizeOfWorksheet() As Integer
        Try
            Dim excelApp As Excel.Application = CType(Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application"), Excel.Application)
            Dim activeSheet As Worksheet = CType(excelApp.ActiveSheet, Worksheet)
            Dim usedRange As Excel.Range = activeSheet.UsedRange

            Dim rowCount As Integer = usedRange.Rows.Count
            Dim colCount As Integer = usedRange.Columns.Count
            Dim totalCells As Integer = rowCount * colCount

            Return totalCells

        Catch ex As System.Exception
            MsgBox("Error in SizeOfWorksheet: " & ex.Message, MsgBoxStyle.Critical)
        End Try

    End Function


    ' Helper function to release COM objects
    Private Sub ReleaseObject(ByVal obj As Object)
        Try
            If obj IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
                obj = Nothing
            End If
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub


    Function GetCellFromInstruction(ByVal instruction As String) As String
        Dim startPos As Integer = instruction.IndexOf("[Cell: ") + 7
        Dim endPos As Integer = instruction.IndexOf("]", startPos)
        If startPos > 6 AndAlso endPos > startPos Then
            Return instruction.Substring(startPos, endPos - startPos).Trim()
        End If
        Return String.Empty
    End Function


    Public Function GetFormulaOrValueFromInstruction(ByVal instruction As String) As String
        If String.IsNullOrEmpty(instruction) Then Return String.Empty

        Dim i As Integer = 0
        While i < instruction.Length
            If instruction(i) = "["c Then
                ' Only consider directives that start at top-level
                If StartsWithAt(instruction, "[Formula: ", i) Then
                    Dim content As String = ExtractBracketContent(instruction, i, "[Formula: ".Length)
                    Return content
                ElseIf StartsWithAt(instruction, "[Value: ", i) Then
                    Dim content As String = ExtractBracketContent(instruction, i, "[Value: ".Length)
                    Return content
                ElseIf StartsWithAt(instruction, "[Comment: ", i) Then
                    Dim content As String = ExtractBracketContent(instruction, i, "[Comment: ".Length)
                    If content IsNot Nothing Then
                        Return $"{AN5}: " & content
                    End If
                    Return String.Empty
                Else
                    ' Some other bracketed segment at top-level: skip it entirely (including nested), then continue
                    Dim closing As Integer = FindMatchingBracket(instruction, i)
                    If closing = -1 Then Exit While
                    i = closing + 1
                    Continue While
                End If
            End If
            i += 1
        End While

        Return String.Empty
    End Function

    ' Helpers used by GetFormulaOrValueFromInstruction
    Private Function StartsWithAt(s As String, prefix As String, index As Integer) As Boolean
        If index + prefix.Length > s.Length Then Return False
        Return s.IndexOf(prefix, index, StringComparison.OrdinalIgnoreCase) = index
    End Function

    Private Function ExtractBracketContent(s As String, openBracketIndex As Integer, tagLen As Integer) As String
        Dim closeIndex As Integer = FindMatchingBracket(s, openBracketIndex)
        If closeIndex = -1 Then Return String.Empty
        Dim contentStart As Integer = openBracketIndex + tagLen
        Dim len As Integer = closeIndex - contentStart
        If len <= 0 Then Return String.Empty
        Return s.Substring(contentStart, len).Trim()
    End Function

    Private Function FindMatchingBracket(s As String, openBracketIndex As Integer) As Integer
        Dim depth As Integer = 0
        For j As Integer = openBracketIndex To s.Length - 1
            Dim c As Char = s(j)
            If c = "["c Then
                depth += 1
            ElseIf c = "]"c Then
                depth -= 1
                If depth = 0 Then
                    Return j
                End If
            End If
        Next
        Return -1
    End Function


    Public Function oldGetFormulaOrValueFromInstruction(ByVal instruction As String) As String
        Dim pattern As String = ""
        ' Determine which marker is present
        If instruction.Contains("[Formula: ") Then
            pattern = "[Formula: "
        ElseIf instruction.Contains("[Value: ") Then
            pattern = "[Value: "
        ElseIf instruction.Contains("[Comment: ") Then
            pattern = "[Comment: "
        Else
            Return String.Empty
        End If

        Dim patternLength As Integer = pattern.Length
        ' Find the start of the entire bracketed expression (the "[" included in the marker)
        Dim openingBracketIndex As Integer = instruction.IndexOf(pattern)
        If openingBracketIndex = -1 Then
            Return String.Empty
        End If

        ' Use a bracket counter to handle nested brackets.
        ' Start scanning from the beginning of the outer bracketed expression.
        Dim counter As Integer = 0
        Dim matchingBracketIndex As Integer = -1
        For i As Integer = openingBracketIndex To instruction.Length - 1
            Dim c As Char = instruction(i)
            If c = "["c Then
                counter += 1
            ElseIf c = "]"c Then
                counter -= 1
                ' When counter reaches zero, we found the closing bracket that matches our opening bracket.
                If counter = 0 Then
                    matchingBracketIndex = i
                    Exit For
                End If
            End If
        Next

        ' If no matching closing bracket was found, return empty string.
        If matchingBracketIndex = -1 Then
            Return String.Empty
        End If

        ' The actual content starts right after the marker (e.g. after "[Value: " or "[Formula: ")
        Dim contentStart As Integer = openingBracketIndex + patternLength
        Dim contentLength As Integer = matchingBracketIndex - contentStart

        If contentLength > 0 Then
            Dim Response As String = instruction.Substring(contentStart, contentLength).Trim()
            If pattern = "[Comment: " Then
                Response = $"{AN5}: " & Response
            End If
            Return Response
        End If
    End Function

    ' Excel Helpers

    Public Sub AdjustHeight(Optional Silent As Boolean = False)

        Dim splash As New SplashScreen("Processing cells... press 'Esc' to abort")

        Try
            ' Get the active Excel worksheet
            Dim activeSheet As Microsoft.Office.Interop.Excel.Worksheet = Globals.ThisAddIn.Application.ActiveSheet
            Dim usedRange As Excel.Range = activeSheet.UsedRange

            ' Get the current selection
            Dim selectedRange As Excel.Range = Globals.ThisAddIn.Application.Selection
            selectedRange = Globals.ThisAddIn.Application.Intersect(selectedRange, usedRange)

            ' Check if the selection is empty or null
            If selectedRange Is Nothing OrElse selectedRange.Count = 0 Then
                Dim result As Integer = 0
                If Not Silent Then
                    result = ShowCustomYesNoBox("No cells are selected. Would you like to perform the operation on the entire worksheet?", "Yes", "No", "Adjust Height")
                End If
                If result = 1 Then
                    selectedRange = activeSheet.UsedRange
                Else
                    If Not Silent Then ShowCustomMessageBox("Operation cancelled.")
                    Exit Sub
                End If
            End If

            ' Perform AutoFit on the rows of the selected range to ensure initial proper height
            selectedRange.Rows.AutoFit()


            ' Prepare dictionaries for tracking row heights
            Dim rowOriginalHeights As New Dictionary(Of Integer, Double)()
            Dim rowMaxHeights As New Dictionary(Of Integer, Double)()

            ' Initialize these dictionaries for each row in the selection
            For Each oneRow As Excel.Range In selectedRange.Rows
                Dim rowIndex As Integer = oneRow.Row
                Dim currentHeight As Double = activeSheet.Rows(rowIndex).RowHeight
                rowOriginalHeights(rowIndex) = currentHeight
                ' Start the max at whatever the row is currently
                rowMaxHeights(rowIndex) = currentHeight
            Next

            splash.Show()
            splash.Refresh()


            ' Iterate through each cell in the selection
            For Each cell As Excel.Range In selectedRange

                System.Windows.Forms.Application.DoEvents()

                If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And &H8000) <> 0 Then Exit For

                If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And 1) <> 0 Then
                    ' Exit the loop
                    Exit For
                End If

                If cell Is Nothing Then Continue For

                ' We'll always enable wrapping so AutoFit will compute multi-line height
                cell.WrapText = True

                Dim wasMerged As Boolean = cell.MergeCells
                Dim mergeArea As Excel.Range = If(wasMerged, cell.MergeArea, cell)

                ' Temporarily store the row index for dictionary look-up
                Dim rowIndex As Integer = mergeArea.Row

                ' We'll measure how tall Excel wants to make this cell
                Dim newHeight As Double = 0

                If wasMerged Then

                    ' Store the original column widths for each column
                    Dim firstColIndex As Integer = mergeArea.Column
                    Dim totalCols As Integer = mergeArea.Columns.Count
                    Dim originalWidths As New List(Of Double)

                    For iCol As Integer = 0 To totalCols - 1
                        Dim colWidth As Double =
                        activeSheet.Columns(firstColIndex + iCol).ColumnWidth
                        originalWidths.Add(colWidth)
                    Next

                    ' Sum the widths so we can set it on the first column after unmerging
                    Dim combinedWidth As Double = originalWidths.Sum()

                    ' Unmerge
                    mergeArea.UnMerge()

                    ' Set only the first column to the combined width so AutoFit sees the "true" width
                    activeSheet.Columns(firstColIndex).ColumnWidth = combinedWidth

                    ' Autofit (note: must do autofit on entire row(s) that the cell spans)
                    mergeArea.Rows.AutoFit()

                    ' Capture the new row height
                    newHeight = mergeArea.RowHeight

                    ' Restore the original column widths
                    For iCol As Integer = 0 To totalCols - 1
                        activeSheet.Columns(firstColIndex + iCol).ColumnWidth = originalWidths(iCol)
                    Next

                    ' Re-merge
                    Dim remergeRange As Excel.Range = activeSheet.Range(
                    activeSheet.Cells(mergeArea.Row, firstColIndex),
                    activeSheet.Cells(mergeArea.Row, firstColIndex + totalCols - 1)
                )
                    remergeRange.Merge()

                Else
                    ' If not merged, simply use AutoFit
                    mergeArea.Rows.AutoFit()
                    newHeight = mergeArea.RowHeight
                End If

                ' Store the maximum needed height for this row so far
                If rowMaxHeights.ContainsKey(rowIndex) Then
                    ' Compare existing max with newly measured height
                    If newHeight > rowMaxHeights(rowIndex) Then
                        rowMaxHeights(rowIndex) = newHeight
                    End If
                End If

            Next


            ' Now set each row’s height to the maximum of:
            For Each rowIndex As Integer In rowMaxHeights.Keys.ToList()

                System.Windows.Forms.Application.DoEvents()

                If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And &H8000) <> 0 Then Exit For

                If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And 1) <> 0 Then
                    ' Exit the loop
                    Exit For
                End If
                Dim finalHeight As Double = Math.Max(rowMaxHeights(rowIndex), rowOriginalHeights(rowIndex))
                If finalHeight > 409 Then
                    finalHeight = 409
                End If

                activeSheet.Rows(rowIndex).RowHeight = finalHeight
            Next


        Catch ex As System.Exception
            MessageBox.Show($"Error in AdjustHeight: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            splash.Close()
        End Try

    End Sub

    Public Sub AdjustLegacyNotes()

        Dim splash As New SplashScreen("Processing cells... press 'Esc' to abort")

        Try
            ' Get the active Excel worksheet
            Dim activeSheet As Microsoft.Office.Interop.Excel.Worksheet = Globals.ThisAddIn.Application.ActiveSheet
            Dim usedRange As Excel.Range = activeSheet.UsedRange

            ' Get the current selection
            Dim selectedRange As Excel.Range = Globals.ThisAddIn.Application.Selection
            selectedRange = Globals.ThisAddIn.Application.Intersect(selectedRange, usedRange)

            ' Check if the selection is empty or null
            If selectedRange Is Nothing OrElse selectedRange.Count = 0 Then
                Dim result As Integer = ShowCustomYesNoBox(
                "No cells are selected. Would you like to perform the operation on the entire worksheet?",
                "Yes",
                "No",
                "Adjust Legacy Notes"
            )

                If result = 1 Then
                    selectedRange = activeSheet.UsedRange
                Else
                    ShowCustomMessageBox("Operation cancelled.")
                    Exit Sub
                End If
            End If

            ' Perform AutoFit on the rows of the selected range to ensure initial proper height
            selectedRange.Rows.AutoFit()


            splash.Show()
            splash.Refresh()

            For Each cell As Excel.Range In selectedRange

                System.Windows.Forms.Application.DoEvents()

                If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And &H8000) <> 0 Then Exit For

                If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And 1) <> 0 Then
                    ' Exit the loop
                    Exit For
                End If

                If cell Is Nothing Then Continue For

                If cell.Comment IsNot Nothing Then

                    ' Ensure the note box dimensions are at least 70 wide and 20 high, and no more than 200 wide
                    Dim comment As Excel.Comment = cell.Comment
                    With comment.Shape

                        .TextFrame.AutoSize = True
                        Dim MinimumHeight As Double = .Height

                        .TextFrame.AutoSize = False

                        ' Enforce width constraints
                        If .Width < 70 Then
                            .Width = 70
                        End If
                        If .Width > 250 Then
                            .Width = 250
                        End If

                        ' Dynamically calculate and set height
                        Dim textLength As Integer = Len(comment.Text)
                        Dim lines As Integer = CInt(Math.Ceiling(textLength / (250 / (.TextFrame.Characters.Font.Size - 2)))) ' Approximation based on average char width
                        Dim lineHeight As Double = .TextFrame.Characters.Font.Size + 2 ' Approximate height per line in points
                        Dim requiredHeight As Double = Math.Max(MinimumHeight, (lines * lineHeight)) + 10

                        If lines > 1 Then .Width = 250

                        .Height = requiredHeight

                    End With
                End If

            Next


        Catch ex As System.Exception
            MessageBox.Show($"Error in AdjustLegacyNotes: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            splash.Close()
        End Try

    End Sub

    Private Shared LastRegexPattern As String = String.Empty
    Private Shared LastRegexOptions As String = String.Empty
    Private Shared LastRegexReplace As String = String.Empty

    Public Sub RegexSearchReplace()

        Dim splash As New SplashScreen("Processing cells... press 'Esc' to abort")

        Try
            ' Get the active worksheet
            Dim activeSheet As Microsoft.Office.Interop.Excel.Worksheet = Globals.ThisAddIn.Application.ActiveSheet
            Dim usedRange As Excel.Range = activeSheet.UsedRange

            ' Get the selected range
            Dim selectedRange As Excel.Range = Globals.ThisAddIn.Application.Selection
            selectedRange = Globals.ThisAddIn.Application.Intersect(selectedRange, usedRange)

            Dim processEntireSheet As Boolean = False

            ' If no range is selected, ask to process the entire worksheet
            If selectedRange Is Nothing OrElse selectedRange.Count = 0 Then

                Dim result As Integer = ShowCustomYesNoBox("No cells are selected. Would you like to perform the operation on the entire worksheet?", "Yes", "No", "Regex Search & Replace")

                If result = 1 Then
                    selectedRange = activeSheet.UsedRange
                    processEntireSheet = True
                Else
                    ShowCustomMessageBox("Operation cancelled.")
                    Exit Sub
                End If
            End If

            ' Step 1: Get regex patterns
            Dim regexPattern As String = ShowCustomInputBox("Step 1: Enter your Regex pattern(s), one per line (more info about Regex: vischerlnk.com/regexinfo):", "Regex Search & Replace", False, LastRegexPattern)?.Trim()
            If String.IsNullOrEmpty(regexPattern) Then Exit Sub

            ' Step 2: Get regex options
            Dim optionsInput As String = ShowCustomInputBox("Enter regex option(s) (i for IgnoreCase, m for Multiline, s for Singleline, c for Compiled, r for RightToLeft, e for ExplicitCapture):", "Regex Search & Replace", True, LastRegexOptions)

            Dim regexOptions As RegexOptions = RegexOptions.None

            If Not String.IsNullOrEmpty(optionsInput) Then
                If optionsInput.Contains("i") Then regexOptions = regexOptions Or RegexOptions.IgnoreCase
                If optionsInput.Contains("m") Then regexOptions = regexOptions Or RegexOptions.Multiline
                If optionsInput.Contains("s") Then regexOptions = regexOptions Or RegexOptions.Singleline
                If optionsInput.Contains("c") Then regexOptions = regexOptions Or RegexOptions.Compiled
                If optionsInput.Contains("r") Then regexOptions = regexOptions Or RegexOptions.RightToLeft
                If optionsInput.Contains("e") Then regexOptions = regexOptions Or RegexOptions.ExplicitCapture
            End If

            ' Step 3: Get replacement text
            Dim replacementText As String = ShowCustomInputBox("Step 2: Enter your replacement text(s), one on each line, matching to your pattern(s):", "Regex Search & Replace", False, LastRegexReplace)

            ' Update the last-used regex pattern and options
            LastRegexPattern = regexPattern
            LastRegexOptions = optionsInput
            LastRegexReplace = replacementText

            ' Split patterns and replacements into lines
            Dim patterns() As String = regexPattern.Split(New String() {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)
            Dim replacements() As String = If(Not String.IsNullOrEmpty(replacementText), replacementText.Split(New String() {Environment.NewLine}, StringSplitOptions.None), Nothing)

            ' Check if patterns and replacements match
            If replacements IsNot Nothing AndAlso patterns.Length <> replacements.Length Then
                ShowCustomMessageBox("The number of regex patterns does not match the number of replacement lines. Aborting without any replacements done.")
                Exit Sub
            End If

            ' Validate all regex patterns first
            For Each pattern As String In patterns
                Try
                    Dim regexTest As New Regex(pattern, regexOptions)
                Catch ex As ArgumentException
                    ShowCustomMessageBox($"Your regex pattern '{pattern}' is invalid ({ex.Message}). Aborting without any replacements done.")
                    Exit Sub
                End Try
            Next

            splash.Show()
            splash.Refresh()

            ' Perform replacements
            Dim totalReplacements As Integer = 0

            For Each cell As Excel.Range In selectedRange

                System.Windows.Forms.Application.DoEvents()

                If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And &H8000) <> 0 Then Exit For

                If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And 1) <> 0 Then
                    ' Exit the loop
                    Exit For
                End If
                If cell.Value2 IsNot Nothing AndAlso TypeOf cell.Value2 Is String Then
                    Dim cellText As String = cell.Value2.ToString()

                    For i As Integer = 0 To patterns.Length - 1
                        Dim regex As New Regex(patterns(i), regexOptions)
                        Dim replacement As String = If(replacements IsNot Nothing, replacements(i), Nothing)

                        ' Perform replacement
                        Dim newText As String = regex.Replace(cellText, replacement)
                        If newText <> cellText Then
                            totalReplacements += 1
                            cell.Value2 = newText
                        End If
                    Next
                End If
            Next

            ShowCustomMessageBox($"{totalReplacements} replacement(s) made in the selected cells.")


        Catch ex As System.Exception
            MessageBox.Show($"Error in RegexSearchReplace: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            splash.Close()
        End Try
    End Sub

    ' Other Helpers 

    Function GetSelectedText(selectedRange As Excel.Range, DoFormulas As Boolean) As String
        Dim selectedTextBuilder As New StringBuilder()

        For Each cell As Excel.Range In selectedRange.Cells
            If Not IsNothing(cell.Value) AndAlso Not CellProtected(cell) Then
                If cell.HasFormula Then
                    If DoFormulas Then
                        selectedTextBuilder.AppendLine(cell.Formula)
                    End If
                Else
                    selectedTextBuilder.AppendLine(CStr(cell.Value))
                End If
            End If
        Next

        Return selectedTextBuilder.ToString()
    End Function
    Private Function AreAllCellsBlocked(ByVal rng As Excel.Range) As Boolean
        Dim allLocked As Boolean = True ' Assume all cells are locked by default

        If rng Is Nothing Then Return False

        ' Check if the worksheet is protected
        If rng.Worksheet.ProtectContents Then
            ' Iterate through each cell in the range
            For Each cell As Excel.Range In rng.Cells
                ' Check if the cell is locked and cannot be changed
                If Not CellProtected(cell) Then
                    allLocked = False
                    Exit For
                End If
            Next
        Else
            ' Worksheet is not protected, so cells can be modified
            allLocked = False
        End If

        ' Return True if all cells are locked and the worksheet is protected
        Return allLocked
    End Function


    Private Function CellProtected(ByVal cell As Excel.Range) As Boolean
        ' If the worksheet is not protected, no cell is actually protected
        If Not cell.Worksheet.ProtectContents Then
            Return False
        End If

        ' If the cell is not locked, it is not protected
        If Not cell.Locked Then
            Return False
        End If

        ' Check whether cell is in any AllowEditRange
        For Each aer As Excel.AllowEditRange In cell.Worksheet.Protection.AllowEditRanges
            ' If Intersect is not Nothing, the cell is within that allow-edit range
            If cell.Application.Intersect(aer.Range, cell) IsNot Nothing Then
                ' The cell can be edited => not protected
                Return False
            End If
        Next

        ' If it is locked, sheet is protected, and no allow-edit range covers the cell => it is effectively protected
        Return True
    End Function



    Public Sub UndoAction()
        Try
            Dim app As Excel.Application = Globals.ThisAddIn.Application
            Dim totalCount As Integer = undoStates.Count
            Dim restoredCount As Integer = 0

            ' Force complete shutdown of Excel's background calculation
            app.ScreenUpdating = False
            app.EnableEvents = False
            app.Calculation = Excel.XlCalculation.xlCalculationManual

            Debug.WriteLine($"Starting undo of {undoStates.Count} states")

            ' Process each saved state to restore the previous value or formula
            For i As Integer = 0 To undoStates.Count - 1
                Dim state = undoStates(i)
                Try
                    Dim ws As Worksheet = Nothing
                    ' Get worksheet - use error handling to be safe
                    Try
                        ws = app.Workbooks.Item(app.ActiveWorkbook.Name).Worksheets(state.WorksheetName)
                    Catch wsEx As Exception
                        Debug.WriteLine($"Could not find worksheet {state.WorksheetName}: {wsEx.Message}")
                        Continue For
                    End Try

                    ' Get the range on the worksheet
                    Dim rng As Excel.Range = Nothing
                    Try
                        rng = ws.Range(state.CellAddress)
                    Catch rngEx As Exception
                        Debug.WriteLine($"Failed to get range {state.CellAddress}: {rngEx.Message}")
                        Continue For
                    End Try

                    If rng IsNot Nothing Then
                        Debug.WriteLine($"Processing {i + 1}/{totalCount}: {state.WorksheetName}!{state.CellAddress}")

                        ' First, check if it's in a table
                        Dim isTableCell As Boolean = False
                        Try
                            For Each tbl As Microsoft.Office.Interop.Excel.ListObject In ws.ListObjects
                                If app.Intersect(tbl.Range, rng) IsNot Nothing Then
                                    isTableCell = True
                                    Debug.WriteLine($"  Cell is in table: {tbl.Name}")
                                    Exit For
                                End If
                            Next
                        Catch tableEx As Exception
                            Debug.WriteLine($"  Table check error: {tableEx.Message}")
                        End Try

                        ' Now restore the value using different strategies
                        If state.HadFormula Then
                            Debug.WriteLine($"  Restoring formula: {state.OldFormula}")

                            ' Try multiple approaches for formula restoration
                            Dim success As Boolean = False

                            ' Approach 1: Direct formula setting
                            If Not success Then
                                Try
                                    rng.Formula = state.OldFormula
                                    success = True
                                    Debug.WriteLine("  Set using Formula")
                                Catch ex As Exception
                                    Debug.WriteLine($"  Formula method failed: {ex.Message}")
                                End Try
                            End If

                            ' Approach 2: Formula2 (newer Excel versions)
                            If Not success Then
                                Try
                                    rng.Formula2 = state.OldFormula
                                    success = True
                                    Debug.WriteLine("  Set using Formula2")
                                Catch ex As Exception
                                    Debug.WriteLine($"  Formula2 method failed: {ex.Message}")
                                End Try
                            End If

                            ' Approach 3: FormulaR1C1 as fallback
                            If Not success Then
                                Try
                                    rng.FormulaR1C1 = state.OldFormula
                                    success = True
                                    Debug.WriteLine("  Set using FormulaR1C1")
                                Catch ex As Exception
                                    Debug.WriteLine($"  FormulaR1C1 method failed: {ex.Message}")
                                End Try
                            End If

                            ' Last resort: Set as value
                            If Not success Then
                                Try
                                    rng.Value = state.OldValue
                                    success = True
                                    Debug.WriteLine("  Set using Value (fallback)")
                                Catch ex As Exception
                                    Debug.WriteLine($"  Value fallback failed: {ex.Message}")
                                End Try
                            End If

                            If success Then restoredCount += 1
                        Else
                            Debug.WriteLine($"  Restoring value: {state.OldValue}")
                            Try
                                ' For non-formula cells, just set the value
                                rng.Value = state.OldValue
                                restoredCount += 1
                            Catch ex As Exception
                                Debug.WriteLine($"  Value restore error: {ex.Message}")
                            End Try
                        End If

                        ' Force immediate update of this cell
                        Try
                            rng.Calculate()
                        Catch ex As Exception
                            ' Ignore calculation errors
                        End Try
                    End If
                Catch ex As Exception
                    Debug.WriteLine($"Error processing state {i}: {ex.Message}")
                End Try

                ' Force periodic UI refresh during long undos
                If i Mod 5 = 0 Then
                    app.ScreenUpdating = True
                    System.Threading.Thread.Sleep(10)
                    app.ScreenUpdating = False
                End If
            Next

            Debug.WriteLine($"Undo complete: {restoredCount}/{totalCount} cells restored")
            undoStates.Clear()
            Dim result = Globals.Ribbons.Ribbon1.UpdateUndoButton()

        Catch ex As System.Exception
            MessageBox.Show("Error during undo: " & ex.Message)
        Finally
            ' Always restore Excel's calculation settings
            Dim app As Excel.Application = Globals.ThisAddIn.Application
            app.ScreenUpdating = True
            app.EnableEvents = True
            app.Calculation = Excel.XlCalculation.xlCalculationAutomatic

            ' Force a full recalculation to ensure all dependencies update
            Try
                app.CalculateFull()
            Catch ex As Exception
                Debug.WriteLine($"Final calculation error: {ex.Message}")
            End Try
        End Try
    End Sub


    Public Sub xxUndoAction()
        Try
            Dim app As Excel.Application = Globals.ThisAddIn.Application

            ' Process each saved state to restore the previous value or formula.
            For Each state In undoStates
                Dim rng As Excel.Range = app.Range(state.CellAddress)
                If state.HadFormula Then
                    rng.Formula = state.OldFormula
                Else
                    rng.Value = state.OldValue
                End If
            Next

            ' Clear the undo state after restoring.
            undoStates.Clear()

            Dim result = Globals.Ribbons.Ribbon1.UpdateUndoButton()


        Catch ex As System.Exception
            MessageBox.Show("Error during undo (" & ex.Message & ").")
        End Try
    End Sub

    Public Function InterpolateAtRuntime(ByVal template As String) As String
        If template Is Nothing Then
            MessageBox.Show("Error InterpolateAtRuntime: Template is Nothing.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return ""
        End If

        template = Regex.Replace(template, "{Codebasis}", "", RegexOptions.IgnoreCase)
        template = Regex.Replace(template, "{INI_DecodedAPI}", "", RegexOptions.IgnoreCase)
        template = Regex.Replace(template, "{INI_DecodedAPI_2}", "", RegexOptions.IgnoreCase)
        template = Regex.Replace(template, "{INI_APIKey}", "", RegexOptions.IgnoreCase)
        template = Regex.Replace(template, "{INI_APIKeyBack}", "", RegexOptions.IgnoreCase)
        template = Regex.Replace(template, "{INI_APIKey_2}", "", RegexOptions.IgnoreCase)
        template = Regex.Replace(template, "{INI_APIKeyBack_2}", "", RegexOptions.IgnoreCase)

        Dim result As String = template

        Dim placeholderPattern As String = "\{([^}]+)\}"
        Dim matches As MatchCollection = Regex.Matches(template, placeholderPattern)

        For Each m As Match In matches
            Dim placeholder As String = m.Value          ' e.g. "{Name}"
            Dim varName As String = m.Groups(1).Value    ' e.g. "Name"

            ' Debug.WriteLine($"placeholder = {placeholder}  Varname = {varName}")
            ' Search for Field
            Dim fieldInfo = Me.GetType().GetField(varName)
            If fieldInfo IsNot Nothing Then
                Dim fieldValue = fieldInfo.GetValue(Me)
                If fieldValue IsNot Nothing Then
                    result = result.Replace(placeholder, fieldValue.ToString())
                End If
                Continue For
            End If

            ' Search for Property
            Dim propInfo = Me.GetType().GetProperty(varName)
            If propInfo IsNot Nothing Then
                Dim propValue = propInfo.GetValue(Me)
                If propValue IsNot Nothing Then
                    result = result.Replace(placeholder, propValue.ToString())
                End If
            End If
        Next

        Return result
    End Function

    Public Function VBAModuleWorking() As Boolean

        Dim xlApp As Microsoft.Office.Interop.Excel.Application = Me.Application

        Try
            ' Call the VBA function
            Dim HelperVersion As Integer = CType(xlApp.Run("CheckAppHelper"), Integer)

            If HelperVersion >= MinHelperVersion Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
        End Try


    End Function


    Private _win As frmHelpMeInky = Nothing

    Public Sub HelpMeInky()
        If _win Is Nothing OrElse _win.IsDisposed Then
            _win = New frmHelpMeInky(_context, RDV)
        End If
        ' No owner needed
        _win.ShowRaised()
    End Sub


    Public Sub ShowSettings()

        Dim Settings As New Dictionary(Of String, String) From {
                {"Temperature", "Temperature of {model}"},
                {"Timeout", "Timeout of {model}"},
                {"Temperature_2", "Temperature of {model2}"},
                {"Timeout_2", "Timeout of {model2}"},
                {"DoubleS", "Convert '" & ChrW(223) & "' to 'ss'"},
                {"NoDash", "Convert em to en dash"},
                {"PreCorrection", "Additional instruction for prompts"},
                {"PostCorrection", "Prompt to apply after queries"},
                {"Language1", "Default translation language 1"},
                {"Language2", "Default translation language 2"},
                {"PromptLibPath", "Prompt library file"},
                {"PromptLibPathLocal", "Prompt library file (local)"},
                {"DefaultPrefix", "Default prefix to use in 'Freestyle'"}
            }

        Dim SettingsTips As New Dictionary(Of String, String) From {
                {"Temperature", "The higher, the more creative the LLM will be (0.0-2.0)"},
                {"Timeout", "In milliseconds"},
                {"Temperature_2", "The higher, the more creative the LLM will be (0.0-2.0)"},
                {"Timeout_2", "In milliseconds"},
                {"DoubleS", "For Switzerland"},
                {"NoDash", "This will convert long dashes typically generated by LLMs but that are not commonly used (thus suggesting that the text has been AI generated)"},
                {"PreCorrection", "Add prompting text that will be added to all basic requests (e.g., for special language tasks)"},
                {"PostCorrection", "Add a prompt that will be applied to each result before it is further processed (slow!)"},
                {"Language1", "The language (in English) that will be used for the first quick access button in the ribbon"},
                {"Language2", "The language (in English) that will be used for the second quick access button in the ribbon"},
                {"PromptLibPath", "The filename (including path, support environmental variables) for your prompt library (if any)"},
                {"PromptLibPathLocal", "The filename (including path, support environmental variables) for your local prompt library (if any)"},
                {"DefaultPrefix", "You can define here the default prefix to use within 'Freestyle' if no other prefix is used (will be added authomatically)."}
                }

        ShowSettingsWindow(Settings, SettingsTips)

        Dim splash As New SplashScreen("Updating menu following your changes ...")
        splash.Show()
        splash.Refresh()

        AddContextMenu()

        splash.Close()

    End Sub

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



    Public Async Sub AnalyzeCsvWithLLM()

        Dim UseSecondAPI As Boolean

        Try
            ' 1) Ask for file (prefer CSV/TXT)
            DragDropFormLabel = "CSV or Text files (*.csv; *.txt)"
            DragDropFormFilter =
            "Supported Files|*.csv;*.txt|" &
            "CSV Files (*.csv)|*.csv|" &
            "Text Files (*.txt)|*.txt"
            Dim filePath As String = GetFileName()
            DragDropFormLabel = ""
            DragDropFormFilter = ""
            If String.IsNullOrWhiteSpace(filePath) Then
                ShowCustomMessageBox("No file has been selected - will abort.")
                Return
            End If

            Dim ext As String = IO.Path.GetExtension(filePath).ToLowerInvariant()
            If ext <> ".csv" AndAlso ext <> ".txt" Then
                Dim answer = ShowCustomYesNoBox("The selected file is not .csv or .txt. Continue anyway?", "Yes", "No", $"{AN} CSV Analyzer")
                If answer <> 1 Then Return
            End If

            ' 2) Ask for separator first (needed to parse header correctly)
            Dim sepDefault As String = GetSetting("CSV_Separator", ";")
            Dim sepInput As String = SLib.ShowCustomInputBox("Enter the CSV separator (single character is recommended):", $"{AN} CSV Analyzer", True, sepDefault)
            If String.IsNullOrWhiteSpace(sepInput) Then sepInput = sepDefault
            Separator = sepInput
            SaveSetting("CSV_Separator", Separator)

            ' 3) Open minimally, read header + count lines efficiently
            Dim headerColumns As List(Of String) = Nothing
            Dim dataLinesCount As Long = 0
            Using fs As New FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite, 1 << 20, FileOptions.SequentialScan)
                Using sr As New StreamReader(fs, detectEncodingFromByteOrderMarks:=True)
                    Dim header As String = sr.ReadLine()
                    If header Is Nothing Then
                        ShowCustomMessageBox("The file appears to be empty.")
                        Return
                    End If
                    headerColumns = ParseCsvLine(header, Separator)

                    ' Count remaining lines efficiently
                    Dim line As String = Nothing
                    While True
                        line = sr.ReadLine()
                        If line Is Nothing Then Exit While
                        dataLinesCount += 1
                    End While
                End Using
            End Using

            ' Show header + line count and confirm
            Dim headerMsg As String = "Header columns (" & headerColumns.Count & "): " & String.Join(" | ", headerColumns)
            Dim linesMsg As String = "Number of lines (excluding header): " & dataLinesCount
            ShowCustomMessageBox(linesMsg & vbCrLf & headerMsg, $"{AN} CSV Analyzer")

            Dim proceed = ShowCustomYesNoBox("Proceed with parsing and analysis?", "Yes", "No", $"{AN} CSV Analyzer")
            If proceed <> 1 Then Return

            ' 4) Collect parameters (single form)
            Dim promptDefault As String = GetSetting("CSV_Prompt", "")
            Dim colsDefault As String = GetSetting("CSV_Columns", "")
            Dim chunkDefault As Integer = GetSetting(Of Integer)("CSV_ChunkSize", 50)
            Dim startSelDefault As Integer = GetSetting(Of Integer)("CSV_StartSelection", 0)
            Dim endSelDefault As Integer = GetSetting(Of Integer)("CSV_EndSelection", 0)
            Dim resultStartLineDefault As Integer = GetSetting(Of Integer)("CSV_ResultStartLine", 1)
            Dim resultStartColDefault As Integer = GetSetting(Of Integer)("CSV_ResultStartColumn", 1)
            Dim attemptsDefault As Integer = GetSetting(Of Integer)("CSV_Attempts", 2)

            Dim do2ndModel As Boolean? = GetSetting(Of Boolean)("CSV_UseSecondModel", False)
            If Not String.IsNullOrWhiteSpace(INI_AlternateModelPath) Then
                ' keep checkbox available
            ElseIf INI_SecondAPI Then
                ' keep checkbox available
            Else
                ' no secondary available -> disable by setting to Nothing
                do2ndModel = CType(Nothing, Boolean?)
            End If

            Dim p0 As New SLib.InputParameter("Prompt for analysis", promptDefault)
            Dim p1 As New SLib.InputParameter("CSV separator", Separator)
            Dim p2 As New SLib.InputParameter("Columns to process (empty = all; separate by same separator)", colsDefault)

            Dim p3 As New SLib.InputParameter("Chunk size in lines", chunkDefault)
            Dim p4 As New SLib.InputParameter("Starting line in file (0 = entire file)", startSelDefault)
            Dim p5 As New SLib.InputParameter("Ending line in file (0 = entire file)", endSelDefault)

            Dim p6 As New SLib.InputParameter("Result start line (row in Excel)", resultStartLineDefault)
            Dim p7 As New SLib.InputParameter("Result start column (1 = A)", resultStartColDefault)
            Dim p8 As New SLib.InputParameter("Number of attempts (in case of errors)", attemptsDefault)

            Dim p9 As SLib.InputParameter
            If do2ndModel.HasValue Then
                p9 = New SLib.InputParameter("Use the secondary model", do2ndModel.Value)
                If Not String.IsNullOrWhiteSpace(INI_AlternateModelPath) Then
                    p9 = New SLib.InputParameter("Use a secondary model", do2ndModel.Value)
                End If
            Else
                p9 = New SLib.InputParameter("Use a secondary model", CType(Nothing, Boolean?))
            End If

            Dim prms() As SLib.InputParameter = {p0, p1, p2, p3, p4, p5, p6, p7, p8, p9}
            If ShowCustomVariableInputForm("Please set the CSV analysis parameters:", $"{AN} CSV Analyzer", prms) = False Then
                Return
            End If

            ' Read back + persist
            OtherPrompt = CStr(prms(0).Value)
            Separator = CStr(prms(1).Value)
            Dim columnsToProcessRaw As String = CStr(prms(2).Value)

            Dim chunkSize As Integer = CInt(SafeToInt(prms(3).Value, 50))
            Dim startSelection As Integer = CInt(SafeToInt(prms(4).Value, 0))
            Dim endSelection As Integer = CInt(SafeToInt(prms(5).Value, 0))

            Dim resultStartLine As Integer = CInt(SafeToInt(prms(6).Value, 3))
            Dim resultStartCol As Integer = CInt(SafeToInt(prms(7).Value, 1))
            Dim llmAttempts As Integer = Math.Max(1, CInt(SafeToInt(prms(8).Value, 2)))

            UseSecondAPI = False
            If TypeOf prms(9).Value Is Boolean Then
                UseSecondAPI = CBool(prms(9).Value)
            End If

            SaveSetting("CSV_Separator", Separator)
            SaveSetting("CSV_Columns", columnsToProcessRaw)
            SaveSetting("CSV_ChunkSize", chunkSize)
            SaveSetting("CSV_StartSelection", startSelection)
            SaveSetting("CSV_EndSelection", endSelection)
            SaveSetting("CSV_ResultStartLine", resultStartLine)
            SaveSetting("CSV_ResultStartColumn", resultStartCol)
            SaveSetting("CSV_Attempts", llmAttempts)
            SaveSetting("CSV_UseSecondModel", UseSecondAPI)
            Try : My.Settings.Save() : Catch : End Try

            If OtherPrompt.Trim().Length < 5 Then
                ShowCustomMessageBox("Please provide a more detailed prompt (at least 5 characters).")
                Return
            End If

            SaveSetting("CSV_Prompt", OtherPrompt)
            Try : My.Settings.Save() : Catch : End Try

            ' 5) Resolve columns to extract (by header names)
            Dim selectedHeaders As List(Of String)
            Dim selectedIdx As List(Of Integer)
            Dim err As String = ResolveColumns(headerColumns, columnsToProcessRaw, Separator, selectedHeaders, selectedIdx)
            If Not String.IsNullOrEmpty(err) Then
                ShowCustomMessageBox(err)
                Return
            End If

            ' 6) If a second/alternate model is desired, prepare
            If UseSecondAPI Then
                If Not String.IsNullOrWhiteSpace(INI_AlternateModelPath) Then
                    If Not ShowModelSelection(_context, INI_AlternateModelPath) Then
                        originalConfigLoaded = False
                        ShowCustomMessageBox("The secondary model could not be loaded - aborting.")
                        Return
                    End If
                End If
            End If


            ' 7) Prepare Excel output header
            Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Globals.ThisAddIn.Application.ActiveSheet
            Dim outRow As Integer = Math.Max(1, resultStartLine)
            Dim outCol As Integer = Math.Max(1, resultStartCol)
            Dim headerInserted As Boolean = False
            Dim insertedRowsTotal As Integer = 0 ' track how many result rows were inserted

            ' In AnalyzeCsvWithLLM(), replace the InsertHeader definition with this (only alignment tweaks added):
            Dim InsertHeader As System.Action =
                    Sub()
                        If headerInserted Then Return

                        ' Remember where the header starts
                        Dim headerStartRow As Integer = outRow

                        ' Title
                        Dim defaultSize As Double
                        Try
                            defaultSize = CDbl(ws.Cells.Font.Size)
                        Catch
                            defaultSize = 11
                        End Try

                        ws.Cells(outRow, outCol).Value = "Analysis Report"
                        With ws.Cells(outRow, outCol).Font
                            .Bold = True
                            .Size = defaultSize + 2
                        End With
                        outRow += 1

                        ' Empty line
                        outRow += 1

                        ' Filename row
                        ws.Cells(outRow, outCol).Value = "Filename:"
                        ws.Cells(outRow, outCol + 1).Value = System.IO.Path.GetFileName(filePath)
                        outRow += 1

                        ' Date row
                        ws.Cells(outRow, outCol).Value = "Date:"
                        ws.Cells(outRow, outCol + 1).Value = DateTime.Now.ToString("G")
                        outRow += 1

                        ' Instruction row
                        ws.Cells(outRow, outCol).Value = "Prompt:"
                        ws.Cells(outRow, outCol + 1).Value = OtherPrompt

                        ' Ensure the column where the prompt is inserted has width 100 Characters and wraps
                        ws.Columns(outCol + 1).ColumnWidth = 100
                        Dim promptCell As Excel.Range = ws.Cells(outRow, outCol + 1)
                        promptCell.WrapText = True

                        outRow += 1

                        ' Model row
                        ws.Cells(outRow, outCol).Value = "Model:"
                        ws.Cells(outRow, outCol + 1).Value = If(UseSecondAPI, INI_Model_2, INI_Model)
                        outRow += 1

                        ' Chunksize row
                        ws.Cells(outRow, outCol).Value = "Chunksize:"
                        ws.Cells(outRow, outCol + 1).Value = chunkSize.ToString()
                        outRow += 1

                        ' Chunksize fields
                        ws.Cells(outRow, outCol).Value = "Columns:"
                        ws.Cells(outRow, outCol + 1).Value = columnsToProcessRaw.ToString()
                        outRow += 1


                        ' Empty line
                        outRow += 1

                        ' Data header
                        ws.Cells(outRow, outCol).Value = "Line(s)"
                        ws.Cells(outRow, outCol + 1).Value = "Result"
                        ws.Range(ws.Cells(outRow, outCol), ws.Cells(outRow, outCol + 1)).Font.Bold = True

                        ' Top- and left-align the entire header block (from title through the data header row)
                        Dim headerEndRow As Integer = outRow
                        Dim headerRange As Excel.Range = ws.Range(ws.Cells(headerStartRow, outCol), ws.Cells(headerEndRow, outCol + 1))
                        headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
                        headerRange.VerticalAlignment = Excel.XlVAlign.xlVAlignTop

                        ' Left-align "Line(s)" header cell (kept; redundant with headerRange but harmless)
                        ws.Cells(outRow, outCol).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft

                        outRow += 1
                        headerInserted = True
                    End Sub

            ' 8) Stream file again and process in chunks
            Dim firstDataAbsLine As Long = If(startSelection > 0, Math.Max(2, startSelection), 2) ' absolute line in file (1 = header)
            Dim lastDataAbsLine As Long = If(endSelection > 0, endSelection, dataLinesCount + 1) ' inclusive, note: header=1, so data runs to totalLines

            If lastDataAbsLine < firstDataAbsLine Then
                ShowCustomMessageBox("The ending line must be greater than or equal to the starting line.")
                Return
            End If

            Dim sysPrompt As String = InterpolateAtRuntime(SP_ParseFile)
            Dim chunkBuffer As New System.Text.StringBuilder(64 * 1024)
            Dim chunkFirstLine As Long = 0
            Dim chunkLastLine As Long = 0
            Dim chunkCounter As Integer = 0

            ' Compute the chunk header once, make it visible to FlushChunk
            Dim headerOut As String = BuildChunkHeader("LineInFile", selectedHeaders, Separator)

            ' Progress bar initialization
            Dim totalLinesToProcess As Integer = CInt(Math.Max(0, lastDataAbsLine - firstDataAbsLine + 1))
            Try
                ShowProgressBarInSeparateThread(AN & " CSV Analyzer", "Analyzing text...")
                ProgressBarModule.CancelOperation = False
                GlobalProgressValue = 0
                GlobalProgressMax = totalLinesToProcess
                GlobalProgressLabel = $"Processing 0 of {totalLinesToProcess} lines..."
            Catch
            End Try

            Dim processedLines As Integer = 0
            Dim cancelled As Boolean = False

            ' Local flush to handle both full and final partial chunks
            Dim FlushChunk As Func(Of Task) =
                            Async Function() As Task
                                If chunkCounter <= 0 OrElse chunkFirstLine <= 0 Then Return

                                ' Ensure header line is present even for the last, partial chunk
                                Dim body As String = chunkBuffer.ToString()
                                If body.Length > 0 AndAlso Not body.StartsWith(headerOut, StringComparison.Ordinal) Then
                                    body = headerOut & Environment.NewLine & body
                                End If

                                Await ProcessOneChunk(
                                    sysPrompt,
                                    body,
                                    chunkFirstLine,
                                    chunkLastLine,
                                    UseSecondAPI,
                                    llmAttempts,
                                    InsertHeader,
                                    Function() outRow,
                                    Sub(n As Integer)
                                        outRow += n
                                        insertedRowsTotal += n
                                    End Sub,
                                    ws,
                                    outCol,
                                    Separator
                                )
                                Dim linesProcessedThisChunk As Integer = CInt(chunkLastLine - chunkFirstLine + 1)
                                processedLines += linesProcessedThisChunk
                                Try
                                    GlobalProgressValue = processedLines
                                    GlobalProgressLabel = $"Processing {processedLines} of {totalLinesToProcess} lines..."
                                Catch
                                End Try

                                ' reset for next chunk
                                chunkBuffer.Clear()
                                chunkCounter = 0
                                chunkFirstLine = 0
                                chunkLastLine = 0
                            End Function

            Using fs2 As New FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite, 1 << 20, FileOptions.SequentialScan)
                Using sr2 As New StreamReader(fs2, detectEncodingFromByteOrderMarks:=True)
                    ' Skip header
                    Dim headerLine As String = sr2.ReadLine()
                    Dim absLine As Long = 1

                    ' Skip to firstDataAbsLine
                    While absLine < firstDataAbsLine - 1
                        If sr2.ReadLine() Is Nothing Then Exit While
                        absLine += 1
                    End While

                    chunkBuffer.Clear()
                    Dim line As String = Nothing

                    While True
                        If ProgressBarModule.CancelOperation Then
                            cancelled = True
                            Exit While
                        End If

                        line = sr2.ReadLine()
                        If line Is Nothing Then Exit While
                        absLine += 1
                        If absLine > lastDataAbsLine Then Exit While

                        Dim fields As List(Of String) = ParseCsvLine(line, Separator)
                        Dim one As String = BuildChunkLine(absLine, fields, selectedIdx, Separator)
                        If chunkCounter = 0 Then
                            chunkBuffer.AppendLine(headerOut)
                            chunkFirstLine = absLine
                        End If
                        chunkBuffer.AppendLine(one)
                        chunkCounter += 1
                        chunkLastLine = absLine

                        If chunkCounter >= Math.Max(1, chunkSize) Then
                            If ProgressBarModule.CancelOperation Then
                                cancelled = True
                                Exit While
                            End If
                            Await FlushChunk() ' flush full chunk
                        End If
                    End While

                    ' Always flush the remainder if any (final partial chunk)
                    If Not cancelled Then
                        Await FlushChunk()
                    End If

                    ' Ensure header + "No findings." if nothing was inserted
                    If Not cancelled Then
                        If insertedRowsTotal = 0 Then
                            If Not headerInserted Then InsertHeader()
                            ' First line of the table with "No findings." in the Result column
                            ws.Cells(outRow, outCol).Value = ""
                            ws.Cells(outRow, outCol + 1).Value = "No findings."
                            Dim nfRange As Excel.Range = ws.Range(ws.Cells(outRow, outCol), ws.Cells(outRow, outCol + 1))
                            nfRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
                            nfRange.VerticalAlignment = Excel.XlVAlign.xlVAlignTop
                            ws.Cells(outRow, outCol + 1).WrapText = True
                            outRow += 1
                        End If
                    End If

                    If Not cancelled AndAlso headerInserted Then
                        outRow += 1 ' empty line

                        ws.Cells(outRow, outCol).Value = "Created by " & AN & " (processed " & processedLines & " of " & totalLinesToProcess & " lines)"
                        ws.Cells(outRow, outCol).Font.Italic = True
                    End If

                End Using
            End Using

            ' Close progress bar and show cancellation/completion message
            Try
                If ProgressBarModule.CancelOperation Then
                    ' user cancelled
                    ShowCustomMessageBox("CSV analysis cancelled by user (processed " & processedLines & " of " & totalLinesToProcess & " lines).", $"{AN} CSV Analyzer")
                Else
                    ' mark finished so progress thread can close
                    ProgressBarModule.CancelOperation = True
                    ShowCustomMessageBox("CSV analysis completed (processed " & processedLines & " of " & totalLinesToProcess & " lines).", $"{AN} CSV Analyzer")
                End If
            Catch
            End Try

            If UseSecondAPI AndAlso originalConfigLoaded Then
                RestoreDefaults(_context, originalConfig)
                originalConfigLoaded = False
            End If

        Catch ex As Exception
            ShowCustomMessageBox("Error in AnalyzeCsvWithLLM: " & ex.Message)
        Finally

            If ProgressBarModule.CancelOperation = False Then
                ' ensure progress bar is closed
                Try
                    ProgressBarModule.CancelOperation = True
                Catch
                End Try
            End If

            If UseSecondAPI AndAlso originalConfigLoaded Then
                RestoreDefaults(_context, originalConfig)
                originalConfigLoaded = False
            End If
        End Try
    End Sub


    ' Process a single chunk with retries and insert results.
    Private Async Function ProcessOneChunk(
    ByVal sysPrompt As String,
    ByVal chunkBody As String,
    ByVal chunkFirstLine As Long,
    ByVal chunkLastLine As Long,
    ByVal useSecond As Boolean,
    ByVal attempts As Integer,
    ByVal ensureHeader As System.Action,
    ByVal getOutRow As System.Func(Of Integer),
    ByVal advanceOutRow As System.Action(Of Integer),
    ByVal ws As Microsoft.Office.Interop.Excel.Worksheet,
    ByVal outCol As Integer,
    ByVal separator As String
) As Task

        Dim userPrompt As String = "<LINESTOPROCESS>" & chunkBody & "</LINESTOPROCESS>"

        Dim lastError As String = ""
        Dim parsed As List(Of Tuple(Of String, String)) = Nothing
        Dim noResult As Boolean = False

        For retry = 1 To Math.Max(1, attempts)
            Dim resp As String = ""
            Try
                resp = Await LLM(sysPrompt, userPrompt, "", "", 0, useSecond, True, OtherPrompt)
            Catch ex As Exception
                resp = ""
                lastError = "LLM call failed: " & ex.Message
            End Try

            resp = If(resp, "").Trim()

            If String.Equals(resp, "[NORESULT]", StringComparison.OrdinalIgnoreCase) Then
                ' Do not retry and do not insert anything into Excel for NORESULT
                noResult = True
                Exit For
            End If

            If Not String.IsNullOrWhiteSpace(resp) Then
                parsed = TryParseLLMResponse(resp)
                If parsed IsNot Nothing AndAlso parsed.Count > 0 Then
                    ' success
                    ensureHeader?.Invoke()

                    Dim startRow As Integer = If(getOutRow Is Nothing, 1, getOutRow())
                    Dim row As Integer = startRow

                    For Each t In parsed
                        Dim lineKey As String = t.Item1.Trim()
                        Dim result As String = t.Item2
                        ws.Cells(row, outCol).Value = lineKey
                        ws.Cells(row, outCol + 1).Value = result
                        row += 1
                    Next

                    If row > startRow Then
                        ' Align both columns left/top for all inserted rows
                        Dim listRange As Excel.Range = ws.Range(ws.Cells(startRow, outCol), ws.Cells(row - 1, outCol + 1))
                        listRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
                        listRange.VerticalAlignment = Excel.XlVAlign.xlVAlignTop

                        ' Wrap text in the Result column
                        Dim resultRange As Excel.Range = ws.Range(ws.Cells(startRow, outCol + 1), ws.Cells(row - 1, outCol + 1))
                        resultRange.WrapText = True

                        ' Keep line keys as text (existing behavior)
                        Dim lineKeyRange As Excel.Range = ws.Range(ws.Cells(startRow, outCol), ws.Cells(row - 1, outCol))
                        lineKeyRange.NumberFormat = "@"
                    End If

                    If advanceOutRow IsNot Nothing Then
                        advanceOutRow(row - startRow)
                    End If

                    Return
                Else
                    lastError = "Could not parse the LLM response."
                End If
            Else
                lastError = "Empty response from LLM."
            End If

            ' Backoff before retry
            Await Task.Delay(750 * retry)
        Next

        ' If NORESULT: insert nothing and return silently
        If noResult Then
            Return
        End If

        ' For other failures, write one diagnostic line
        ensureHeader?.Invoke()
        Dim startErrRow As Integer = If(getOutRow Is Nothing, 1, getOutRow())
        Dim rangeKey As String = If(chunkFirstLine = chunkLastLine, chunkFirstLine.ToString(), $"{chunkFirstLine}-{chunkLastLine}")
        ws.Cells(startErrRow, outCol).Value = rangeKey
        ws.Cells(startErrRow, outCol + 1).Value = If(String.IsNullOrWhiteSpace(lastError), "Empty or unparsable response.", lastError)
        If advanceOutRow IsNot Nothing Then
            advanceOutRow(1)
        End If
    End Function

    ' Parses "line@@result§§§line@@result..." into list of (line, result).
    Private Function TryParseLLMResponse(ByVal response As String) As List(Of Tuple(Of String, String))
        Dim result As New List(Of Tuple(Of String, String))()

        Dim records = response.Split(New String() {"§§§"}, StringSplitOptions.RemoveEmptyEntries)
        For Each rec In records
            Dim trimmed = rec.Trim()
            If trimmed.Length = 0 Then Continue For
            Dim parts = trimmed.Split(New String() {"@@"}, 2, StringSplitOptions.None)
            If parts.Length = 2 Then
                Dim key = parts(0).Trim()
                Dim val = parts(1).Trim()
                If key.Length > 0 AndAlso val.Length >= 0 Then
                    result.Add(Tuple.Create(key, val))
                End If
            End If
        Next

        Return result
    End Function

    ' Build header line for chunk body: "LineInFile{sep}Col1{sep}Col2..."
    Private Function BuildChunkHeader(ByVal firstHeader As String, ByVal headers As List(Of String), ByVal separator As String) As String
        Return firstHeader & separator & String.Join(separator, headers)
    End Function

    ' Build a single data line: "<absLine>{sep}<v1>{sep}<v2>..."
    Private Function BuildChunkLine(ByVal absLine As Long, ByVal fields As List(Of String), ByVal selectedIdx As List(Of Integer), ByVal separator As String) As String
        Dim vals As New List(Of String)(selectedIdx.Count)
        For Each idx In selectedIdx
            Dim v As String = If(idx >= 0 AndAlso idx < fields.Count, fields(idx), "")
            vals.Add(v)
        Next
        Return absLine.ToString() & separator & String.Join(separator, vals)
    End Function

    ' Map selected headers to indices. If empty or whitespace, select all headers.
    Private Function ResolveColumns(
        ByVal headerColumns As List(Of String),
        ByVal columnsRaw As String,
        ByVal separator As String,
        ByRef selectedHeaders As List(Of String),
        ByRef selectedIdx As List(Of Integer)
    ) As String
        selectedHeaders = New List(Of String)()
        selectedIdx = New List(Of Integer)()

        If headerColumns Is Nothing OrElse headerColumns.Count = 0 Then
            Return "The header row could not be parsed."
        End If

        If String.IsNullOrWhiteSpace(columnsRaw) Then
            ' all columns
            For i = 0 To headerColumns.Count - 1
                selectedHeaders.Add(headerColumns(i))
                selectedIdx.Add(i)
            Next
            Return ""
        End If

        ' Split by the same separator (user-defined)
        Dim requested = columnsRaw.Split(New String() {separator}, StringSplitOptions.RemoveEmptyEntries).
                                 Select(Function(s) s.Trim()).
                                 Where(Function(s) s.Length > 0).
                                 ToList()

        If requested.Count = 0 Then
            ' fallback to all
            For i = 0 To headerColumns.Count - 1
                selectedHeaders.Add(headerColumns(i))
                selectedIdx.Add(i)
            Next
            Return ""
        End If

        ' Build a case-insensitive map from header -> index
        Dim map As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
        For i = 0 To headerColumns.Count - 1
            If Not map.ContainsKey(headerColumns(i)) Then
                map.Add(headerColumns(i), i)
            End If
        Next

        Dim missing As New List(Of String)()
        For Each name In requested
            Dim idx As Integer
            If map.TryGetValue(name, idx) Then
                selectedHeaders.Add(name)
                selectedIdx.Add(idx)
            Else
                missing.Add(name)
            End If
        Next

        If missing.Count > 0 Then
            Return "These columns were not found in the header: " & String.Join(", ", missing)
        End If

        Return ""
    End Function

    ' Robust CSV line parser for common cases: handles quoted fields and embedded separators/newlines within quotes.
    Private Function ParseCsvLine(ByVal line As String, ByVal separator As String) As List(Of String)
        Dim sep As Char = If(String.IsNullOrEmpty(separator), ";"c, separator(0))
        Dim res As New List(Of String)()
        If line Is Nothing Then Return res

        Dim sb As New System.Text.StringBuilder()
        Dim inQuotes As Boolean = False
        Dim i As Integer = 0

        While i < line.Length
            Dim c As Char = line(i)
            If c = """"c Then
                If inQuotes AndAlso i + 1 < line.Length AndAlso line(i + 1) = """"c Then
                    ' escaped quote
                    sb.Append(""""c)
                    i += 2
                    Continue While
                Else
                    inQuotes = Not inQuotes
                    i += 1
                    Continue While
                End If
            End If

            If Not inQuotes AndAlso c = sep Then
                res.Add(sb.ToString())
                sb.Clear()
                i += 1
                Continue While
            End If

            sb.Append(c)
            i += 1
        End While

        res.Add(sb.ToString())
        Return res
    End Function

    ' Helpers to use My.Settings defensively even if keys aren't present yet.
    Private Function GetSetting(Of T)(ByVal key As String, ByVal defaultValue As T) As T
        Try
            Dim p = My.Settings.Properties(key)
            If p IsNot Nothing Then
                Dim v = My.Settings.Item(key)
                If v IsNot Nothing Then
                    Return CType(v, T)
                End If
            End If
        Catch
        End Try
        Return defaultValue
    End Function

    Private Sub SaveSetting(Of T)(ByVal key As String, ByVal value As T)
        Try
            Dim p = My.Settings.Properties(key)
            If p IsNot Nothing Then
                My.Settings.Item(key) = value
            End If
        Catch
        End Try
    End Sub

    Private Function SafeToInt(value As Object, fallback As Integer) As Integer
        Try
            If value Is Nothing Then Return fallback
            Dim s = value.ToString().Trim()
            Dim n As Integer
            If Integer.TryParse(s, n) Then Return n
        Catch
        End Try
        Return fallback
    End Function


End Class
