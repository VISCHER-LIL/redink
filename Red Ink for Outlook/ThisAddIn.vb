﻿' Red Ink for Outlook
' Copyright by David Rosenthal, david.rosenthal@vischer.com
' May only be used under the Red Ink License. See License.txt or https://vischer.com/redink for more information.
'
' 15.4.2025
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
' Includes Whisper.net in unchanged form; Copyright (c) 2024 Sandro Hanea; licensed under the MIT License under the MIT license (https://licenses.nuget.org/MIT) at https://github.com/sandrohanea/whisper.net
' Includes also various Microsoft libraries copyrighted by Microsoft Corporation and available, among others, under the Microsoft EULA and the MIT License; Copyright (c) 2016- Microsoft Corp.

Option Explicit On

Imports Microsoft.Office.Interop.Outlook
Imports Microsoft.Office.Interop.Word
Imports System.Windows.Forms
Imports System.Threading.Tasks
Imports DiffPlex
Imports DiffPlex.DiffBuilder
Imports DiffPlex.DiffBuilder.Model
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedContext
Imports SLib = SharedLibrary.SharedLibrary.SharedMethods
Imports System.Text.RegularExpressions
Imports Markdig
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop
Imports System.Diagnostics
Imports System.IO
Imports System.Net
Imports System.Threading

Module Module1
    ' Correct attribute declaration for DllImport
    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Public Function GetAsyncKeyState(ByVal vKey As Integer) As Short
    End Function
End Module

Public Class ThisAddIn

    Public StartupInitialized As Boolean = False
    Private mainThreadControl As New System.Windows.Forms.Control()
    Private WithEvents outlookExplorer As Outlook.Explorer

    Private Sub ThisAddIn_Startup() Handles Me.Startup

        mainThreadControl.CreateControl()

        outlookExplorer = Application.ActiveExplorer()

        If outlookExplorer IsNot Nothing Then
            AddHandler outlookExplorer.Activate, AddressOf Explorer_Activate
        Else
            mainThreadControl.BeginInvoke(CType(AddressOf DelayedStartupTasks, MethodInvoker))
            StartupInitialized = True
        End If
    End Sub

    Private Sub Explorer_Activate()
        StartupInitialized = True
        RemoveHandler outlookExplorer.Activate, AddressOf Explorer_Activate
        DelayedStartupTasks()
    End Sub

    Private Sub DelayedStartupTasks()
        Try
            InitializeConfig(True, True)
            UpdateHandler.PeriodicCheckForUpdates(INI_UpdateCheckInterval, "Outlook", INI_UpdatePath)
            Dim result = Globals.Ribbons.Ribbon1.UpdateRibbon()
            result = Globals.Ribbons.Ribbon2.UpdateRibbon()
            mainThreadControl.CreateControl()
            StartupHttpListener()
        Catch ex As System.Exception
            ' Handling errors gracefully
        End Try
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
        ShutdownHttpListener()
    End Sub

    ' Hardcoded config values

    Public Const AN As String = "Red Ink"
    Public Const AN2 As String = "red_ink"

    Public Const Version As String = "V.150425 Gen2 Beta Test"

    ' Hardcoded configuration

    Public Const ShortenPercent As Integer = 20
    Public Const SummaryPercent As Integer = 20
    Private Const NetTrigger As String = "(net)"
    Private Const LibTrigger As String = "(Lib)"
    Private Const MarkupPrefix As String = "Markup:"
    Private Const MarkupPrefixDiff As String = "MarkupDiff:"
    Private Const MarkupPrefixDiffW As String = "MarkupDiffW:"
    Private Const MarkupPrefixWord As String = "MarkupWord:"
    Private Const MarkupPrefixAll As String = "Markup[Diff|DiffW|Word]:"
    Private Const ClipboardPrefix As String = "Clipboard:"
    Private Const InsertPrefix As String = "Insert:"
    Private Const NoFormatTrigger As String = "(noformat)"
    Private Const NoFormatTrigger2 As String = "(nf)"
    Private Const KFTrigger As String = "(keepformat)"
    Private Const KFTrigger2 As String = "(kf)"
    Private Const KPFTrigger As String = "(keepparaformat)"
    Private Const KPFTrigger2 As String = "(kpf)"
    Private Const InPlacePrefix As String = "Replace:"

    Private Const ESC_KEY As Integer = &H1B

    Private Const SecondAPICode As String = "(2nd)"

    ' Variables that are available to InterpolateAtRuntime

    Public TranslateLanguage As String = ""
    Public OtherPrompt As String = ""
    Public ShortenLength, SummaryLength As Long
    Public DateTimeNow As String

    Public InspectorOpened As Boolean = False

    ' Definition of the SharedProperties for context for exchanging values with the SharedLibrary

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

    Public Shared Property INI_Response As String
        Get
            Return _context.INI_Response
        End Get
        Set(value As String)
            _context.INI_Response = value
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

    Public Shared Property INI_Response_2 As String
        Get
            Return _context.INI_Response_2
        End Get
        Set(value As String)
            _context.INI_Response_2 = value
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

    Public Shared Property SP_Add_Bubbles As String
        Get
            Return _context.SP_Add_Bubbles
        End Get
        Set(value As String)
            _context.SP_Add_Bubbles = value
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

    Public Shared Property SP_Add_ChatWord_Commands As String
        Get
            Return _context.SP_Add_ChatWord_Commands
        End Get
        Set(value As String)
            _context.SP_Add_ChatWord_Commands = value
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

    Public Shared ReadOnly Property RDV As String = "Outlook (" & Version & ")"
    Public Shared ReadOnly Property InitialConfigFailed As Boolean = False
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

    Public Shared Property INI_Placeholder_03 As String
        Get
            Return _context.INI_Placeholder_03
        End Get
        Set(value As String)
            _context.INI_Placeholder_03 = value
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
    Public Shared Property INI_ISearch_Apply_SP_Markup As String
        Get
            Return _context.INI_ISearch_Apply_SP_Markup
        End Get
        Set(value As String)
            _context.INI_ISearch_Apply_SP_Markup = value
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

    Public Shared Property INI_Placeholder_01 As String
        Get
            Return _context.INI_Placeholder_01
        End Get
        Set(value As String)
            _context.INI_Placeholder_01 = value
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

    Public Shared Property INI_Lib_Apply_SP_Markup As String
        Get
            Return _context.INI_Lib_Apply_SP_Markup
        End Get
        Set(value As String)
            _context.INI_Lib_Apply_SP_Markup = value
        End Set
    End Property

    Public Shared Property INI_Placeholder_02 As String
        Get
            Return _context.INI_Placeholder_02
        End Get
        Set(value As String)
            _context.INI_Placeholder_02 = value
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

    Public Shared Property INI_PromptLibPath As String
        Get
            Return _context.INI_PromptLibPath
        End Get
        Set(value As String)
            _context.INI_PromptLibPath = value
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


    Public Shared Property INI_PromptLibPath_Transcript As String
        Get
            Return _context.INI_PromptLibPath_Transcript
        End Get
        Set(value As String)
            _context.INI_PromptLibPath_Transcript = value
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


#End Region

    ' Functions of SharedLibrary

#Region "SharedLibrary"

    Public Sub InitializeConfig(FirstTime As Boolean, Reload As Boolean)
        _context.InitialConfigFailed = False
        _context.RDV = "Outlook (" & Version & ")"
        SharedMethods.InitializeConfig(_context, FirstTime, Reload)
    End Sub
    Private Function INIValuesMissing()
        Return SharedMethods.INIValuesMissing(_context)
    End Function
    Public Shared Async Function PostCorrection(inputText As String, Optional ByVal UseSecondAPI As Boolean = False) As Task(Of String)
        Return Await SharedMethods.PostCorrection(_context, inputText, UseSecondAPI)
    End Function

    Public Shared Async Function LLM(ByVal promptSystem As String, ByVal promptUser As String, Optional ByVal Model As String = "", Optional ByVal Temperature As String = "", Optional ByVal Timeout As Long = 0, Optional ByVal UseSecondAPI As Boolean = False, Optional HideSplash As Boolean = False, Optional ByVal AddUserPrompt As String = "") As Task(Of String)
        Return Await SharedMethods.LLM(_context, promptSystem, promptUser, Model, Temperature, Timeout, UseSecondAPI, HideSplash, AddUserPrompt)
    End Function

    Private Function ShowSettingsWindow(Settings As Dictionary(Of String, String), SettingsTips As Dictionary(Of String, String))
        SharedMethods.ShowSettingsWindow(Settings, SettingsTips, _context)
    End Function
    Private Function ShowPromptSelector(filePath As String, enableMarkup As Boolean, enableBubbles As Boolean) As (String, Boolean, Boolean, Boolean)
        Return SharedMethods.ShowPromptSelector(filePath, enableMarkup, enableBubbles, _context)
    End Function

#End Region

    Enum Operation
        Insert = 1
        Delete = 2
        Equal = 3
    End Enum

    Public Sub MainMenu(RI_Command As String)

        If Not INIloaded Then
            If Not StartupInitialized Then
                Try
                    DelayedStartupTasks()
                    RemoveHandler outlookExplorer.Activate, AddressOf Explorer_Activate
                Catch ex As System.Exception
                End Try
                If Not INIloaded Then Exit Sub
            Else
                InitializeConfig(False, False)
                If Not INIloaded Then
                    Exit Sub
                End If
            End If
        End If

        Try
            ' Use fully qualified names to avoid ambiguity
            Dim outlookApp As New Microsoft.Office.Interop.Outlook.Application()
            Dim inspector As Microsoft.Office.Interop.Outlook.Inspector = outlookApp.ActiveInspector
            Dim Textlength As Long

            If inspector Is Nothing Then

                InspectorOpened = False

                OpenInspectorAndReapplySelection(RI_Command = "Sumup")

                If Not InspectorOpened Then Exit Sub

                inspector = outlookApp.ActiveInspector
                If inspector Is Nothing Then

                    System.Windows.Forms.MessageBox.Show("Error in MainMenu: No active email item found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If
            End If

            If inspector.CurrentItem.Class = Microsoft.Office.Interop.Outlook.OlObjectClass.olMail Then
                Dim mailItem As Microsoft.Office.Interop.Outlook.MailItem = DirectCast(inspector.CurrentItem, Microsoft.Office.Interop.Outlook.MailItem)
                Dim wordEditor As Microsoft.Office.Interop.Word.Document = DirectCast(inspector.WordEditor, Microsoft.Office.Interop.Word.Document)

                InitializeConfig(False, False)

                If GPTSetupError OrElse INIValuesMissing() Or Not INIloaded Then Return

                Select Case RI_Command

                    Case "Translate"
                        TranslateLanguage = ""
                        TranslateLanguage = SLib.ShowCustomInputBox("Enter your target language:", $"{AN} Translate", True)
                        If String.IsNullOrEmpty(TranslateLanguage) Then Return
                        Command_InsertAfter(InterpolateAtRuntime(SP_Translate), False, INI_KeepFormat1, INI_ReplaceText1)
                    Case "PrimLang"
                        TranslateLanguage = INI_Language1
                        Command_InsertAfter(InterpolateAtRuntime(SP_Translate), False, INI_KeepFormat1, INI_ReplaceText1)
                    Case "Correct"
                        Command_InsertAfter(InterpolateAtRuntime(SP_Correct), INI_DoMarkupOutlook, INI_KeepFormat2, INI_ReplaceText2, INI_MarkupMethodOutlook)
                    Case "Summarize"

                        Textlength = GetSelectedTextLength()

                        If Textlength = 0 Then
                            SLib.ShowCustomMessageBox("Please select the text to be processed.")
                            Exit Sub
                        End If

                        Dim UserInput As String
                        SummaryLength = 0

                        Do
                            UserInput = Trim(SLib.ShowCustomInputBox("Enter the number of words your summary shall have (the selected text has " & Textlength & " words; the proposal " & SummaryPercent & "%):", $"{AN} Summarizer", True, CStr(SummaryPercent * Textlength / 100)))

                            If String.IsNullOrEmpty(UserInput) Then
                                Exit Sub
                            End If

                            If Integer.TryParse(UserInput, SummaryLength) AndAlso SummaryLength >= 1 AndAlso SummaryLength <= Textlength Then
                                Exit Do
                            Else
                                SLib.ShowCustomMessageBox("Please enter a valid word count between 1 and " & Textlength & ".")
                            End If
                        Loop
                        If SummaryLength = 0 Then Exit Sub
                        'SummaryLength = (Textlength - (Textlength * SummaryPercent / 100))'

                        Command_InsertAfter(InterpolateAtRuntime(SP_Summarize), False)
                    Case "Improve"
                        Command_InsertAfter(InterpolateAtRuntime(SP_Improve), INI_DoMarkupOutlook, INI_KeepFormat2, INI_ReplaceText2, INI_MarkupMethodOutlook)
                    Case "NoFillers"
                        Command_InsertAfter(InterpolateAtRuntime(SP_NoFillers), INI_DoMarkupOutlook, INI_KeepFormat2, INI_ReplaceText2, INI_MarkupMethodOutlook)
                    Case "Friendly"
                        Command_InsertAfter(InterpolateAtRuntime(SP_Friendly), INI_DoMarkupOutlook, INI_KeepFormat2, INI_ReplaceText2, INI_MarkupMethodOutlook)
                    Case "Convincing"
                        Command_InsertAfter(InterpolateAtRuntime(SP_Convincing), INI_DoMarkupOutlook, INI_KeepFormat2, INI_ReplaceText2, INI_MarkupMethodOutlook)
                    Case "Shorten"
                        Textlength = GetSelectedTextLength()
                        If Textlength = 0 Then
                            SLib.ShowCustomMessageBox("Please select the text to be processed.")
                            Exit Sub
                        End If
                        Dim UserInput As String
                        Dim ShortenPercentValue As Integer = 0
                        Do
                            UserInput = Trim(SLib.ShowCustomInputBox("Enter the percentage by which your text should be shortened (it has " & Textlength & " words; " & ShortenPercent & "% will cut approx. " & (Textlength * ShortenPercent / 100) & " words)", $"{AN} Shortener", True, CStr(ShortenPercent) & "%"))
                            If String.IsNullOrEmpty(UserInput) Then
                                Exit Sub
                            End If
                            UserInput = UserInput.Replace("%", "").Trim()
                            If Integer.TryParse(UserInput, ShortenPercentValue) AndAlso ShortenPercentValue >= 1 AndAlso ShortenPercentValue <= 99 Then
                                Exit Do
                            Else
                                SLib.ShowCustomMessageBox("Please enter a valid percentage between 1 And 99.")
                            End If
                        Loop
                        ShortenLength = (Textlength - (Textlength * (100 - ShortenPercent) / 100))
                        Command_InsertAfter(InterpolateAtRuntime(SP_Shorten), INI_DoMarkupOutlook, INI_KeepFormat2, INI_ReplaceText2, INI_MarkupMethodOutlook)
                    Case "Sumup"
                        FreeStyle_InsertBefore(SP_MailSumup, False)
                    Case "Answers"
                        FreeStyle_InsertBefore(SP_MailReply, True)
                    Case "Freestyle"
                        FreeStyle_InsertAfter()
                    Case Else
                        System.Windows.Forms.MessageBox.Show("Error in MainMenu: Invalid internal command.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Select

            Else
                SLib.ShowCustomMessageBox($"Please open an email for editing for using {AN}.")
            End If
            If inspector IsNot Nothing Then Marshal.ReleaseComObject(inspector) : inspector = Nothing
            If outlookApp IsNot Nothing Then Marshal.ReleaseComObject(outlookApp) : outlookApp = Nothing
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show("Error in MainMenu: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Public Sub OpenInspectorAndReapplySelection(Sumup As Boolean)
        Try
            ' Grab Outlook instances
            Dim oApp As Outlook.Application = Globals.ThisAddIn.Application
            Dim oExplorer As Outlook.Explorer = oApp.ActiveExplorer()

            If oExplorer Is Nothing Then
                If Sumup Then
                    ShowCustomMessageBox("You can only use this function when you have selected an e-mail.")
                Else
                    ShowCustomMessageBox("You can only use this function when you are editing an e-mail.")
                End If
                Return
            End If

            ' Check for inline response
            Dim inlineResponse As Object = oExplorer.ActiveInlineResponse
            If inlineResponse Is Nothing Then

                ' Get the current selection in the explorer
                Dim selection As Outlook.Selection = oExplorer.Selection

                ' Check if any item is selected
                If selection.Count = 0 Then
                    ShowCustomMessageBox("No email is selected.")
                    Return
                End If

                If selection.Count > 1 Then
                    If Not Sumup Then
                        ShowCustomMessageBox("Multiple emails selected. Please select only one email when not using Sumup mode.")
                        Return
                    Else
                        ' Combine texts from all selected emails.
                        Dim mailItems As New List(Of Microsoft.Office.Interop.Outlook.MailItem)
                        For Each item As Object In selection
                            If TypeOf item Is Microsoft.Office.Interop.Outlook.MailItem Then
                                mailItems.Add(CType(item, Microsoft.Office.Interop.Outlook.MailItem))
                            End If
                        Next

                        If mailItems.Count = 0 Then
                            ShowCustomMessageBox("None of the selected items are emails.")
                            Return
                        End If

                        ' Order the emails: latest email first (descending order by ReceivedTime)
                        mailItems = mailItems.OrderByDescending(Function(m) m.ReceivedTime).ToList()

                        Dim selectedText As String = String.Empty
                        Dim count As Integer = 1
                        For Each mail As Microsoft.Office.Interop.Outlook.MailItem In mailItems
                            Dim tag As String = count.ToString("D4") ' Format count with four digits
                            Dim latestBody As String = GetLatestMailBody(mail.Body)
                            selectedText &= "<EMAIL" & tag & ">" & latestBody & "</EMAIL" & tag & ">"
                            count += 1
                        Next

                        ShowSumup2(selectedText)
                        Return
                    End If
                Else
                    ' Only one email is selected.
                    If Sumup Then
                        Dim selectedItem As Object = selection(1)
                        If TypeOf selectedItem Is Outlook.MailItem Then
                            Dim mail As Outlook.MailItem = CType(selectedItem, Outlook.MailItem)
                            Dim selectedText As String = mail.Body
                            ShowSumup(selectedText)
                            Return
                        Else
                            ShowCustomMessageBox("The selected item is not an email.")
                            Return
                        End If
                    Else
                        ShowCustomMessageBox("You can only use this function when you are editing one (single) e-mail.")
                        Return
                    End If
                End If

            End If

            ' Ensure it is a MailItem
            Dim mailItem As MailItem = TryCast(inlineResponse, MailItem)
            If mailItem Is Nothing Then
                ShowCustomMessageBox("You can only use this function when you are editing an e-mail (currently, there is no valid e-mail item).")
                Return
            End If

            ' Capture the user's current selection range (or caret) from the inline editor
            Dim oldSelStart As Integer = 0
            Dim oldSelEnd As Integer = 0
            If Not GetSelectionOrCaretRangeFromInlineEditor(oExplorer, oldSelStart, oldSelEnd) Then
                ' If this fails entirely (no Word editor, etc.), we can just open the window without reapplying.
                ' But no error is shown for "empty selection" anymore – only true failures (e.g., no WordEditor).
                ' We'll just continue and open the Inspector, albeit we can't set the cursor position.
            End If

            ' Open the Inspector modelessly
            Dim inspector As Inspector = mailItem.GetInspector
            If inspector Is Nothing Then
                MessageBox.Show("Error in OpenInspectorAndReapplySelection: Failed to open the ActiveInspector.",
                            "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If
            inspector.Display(False) ' modeless - do not block

            ' A short delay to let the new WordEditor initialize
            System.Threading.Thread.Sleep(500)

            ' Ensure it's still open
            If inspector.CurrentItem Is Nothing Then
                MessageBox.Show("Error in OpenInspectorAndReapplySelection: The Inspector window was closed before processing could complete.",
                            "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            ' Reapply the original selection (or caret position) to the new Inspector's WordEditor
            Try
                Dim wordDoc As Word.Document = TryCast(inspector.WordEditor, Word.Document)
                If wordDoc IsNot Nothing Then
                    Dim wordSel As Word.Selection = wordDoc.Application.Selection

                    ' Only reapply if we successfully retrieved the inline offsets
                    If oldSelStart <> 0 OrElse oldSelEnd <> 0 Then
                        wordSel.SetRange(Start:=oldSelStart, End:=oldSelEnd)
                        wordSel.Select()
                    End If
                End If

            Catch ex As System.Exception
                MessageBox.Show("Error in OpenInspectorAndReapplySelection: Failed to restore the original selection: " & ex.Message,
                            "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End Try

            ' Bring the new Inspector window to the foreground

            InspectorOpened = True

            inspector.Activate()

            ' Clean up COM references
            Marshal.ReleaseComObject(inspector)
            Marshal.ReleaseComObject(oExplorer)

            Return

        Catch ex As System.Exception
            MessageBox.Show("Error in OpenInspectorAndReapplySelection: " & ex.Message,
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Function GetLatestMailBody(ByVal fullBody As String) As String
        Try
            ' Define an array of candidate markers that are common indicators of quoted messages,
            ' including localized variants.
            Dim markers() As String = {
            "-----Original Message-----",
            "-----Ursprüngliche Nachricht-----",
            "-----Vorherige Nachricht-----",
            "-----Mensaje original-----",
            "-----Messaggio originale-----",
            "-----Courrier original-----",
            "On ",
            "wrote:"
        }

            ' Regular expression to detect header lines with a proper email address
            Dim emailPattern As String = "^(From:|Von:|De:|Da:)\s+[\w\.-]+@[\w\.-]+\.\w+"

            ' Split the email body into lines
            Dim lines() As String = fullBody.Split(New String() {Environment.NewLine}, StringSplitOptions.None)
            Dim sb As New StringBuilder()

            For i As Integer = 0 To lines.Length - 1
                Dim currentLine As String = lines(i)
                Dim trimmedLine As String = currentLine.TrimStart()
                Dim isChainMarker As Boolean = False

                ' First, check each line against our list of known chain markers.
                For Each marker As String In markers
                    If trimmedLine.StartsWith(marker, StringComparison.InvariantCultureIgnoreCase) Then
                        ' Only consider short lines (heuristically less than 100 characters) as markers.
                        If trimmedLine.Length < 100 Then
                            isChainMarker = True
                            Exit For
                        End If
                    End If
                Next

                ' If none of the above markers was found, try to detect headers indicating a quoted message.
                If Not isChainMarker Then
                    ' Check for email header markers using a regex pattern (with an @ symbol)
                    If Regex.IsMatch(trimmedLine, emailPattern, RegexOptions.IgnoreCase) Then
                        isChainMarker = True
                    Else
                        ' Additional check: headers with a name or parenthesized comment following the marker.
                        Dim headerMarkers() As String = {"From:", "Von:", "De:", "Da:"}
                        For Each header As String In headerMarkers
                            If trimmedLine.StartsWith(header, StringComparison.InvariantCultureIgnoreCase) Then
                                ' Extract the text after the header marker.
                                Dim remainingText As String = trimmedLine.Substring(header.Length).Trim()
                                ' Check if the remaining text contains a comma (e.g., "Doe, John") or a parenthesized group.
                                If remainingText.Contains(",") OrElse (remainingText.Contains("(") AndAlso remainingText.Contains(")")) Then
                                    isChainMarker = True
                                    Exit For
                                End If
                            End If
                        Next
                    End If
                End If

                ' If a marker is confidently detected, assume the latest mail ends here.
                If isChainMarker Then
                    Return sb.ToString().TrimEnd()
                End If

                ' Otherwise, add the current line to the accumulated result.
                sb.AppendLine(currentLine)
            Next

            ' No clear marker found; return the full email content.
            Return fullBody
        Catch ex As System.Exception
            ' In case of any error, return the full email body
            ' (Alternatively, you could log the exception as needed)
            Return fullBody
        End Try
    End Function


    Private Function GetSelectionOrCaretRangeFromInlineEditor(oExplorer As Outlook.Explorer, ByRef selStart As Integer, ByRef selEnd As Integer) As Boolean
        Try
            Dim inlineWordEditor As Object = oExplorer.ActiveInlineResponseWordEditor
            If inlineWordEditor Is Nothing Then
                ' No inline Word editor, so we can't read a selection/caret
                Return False
            End If

            Dim wordSel As Word.Selection =
            TryCast(inlineWordEditor.Application.Selection, Word.Selection)
            If wordSel Is Nothing Then
                Return False
            End If

            ' Even if there's no highlighted text, there's always a caret position
            ' So we record them (could be equal if there's no actual selection)
            selStart = wordSel.Start
            selEnd = wordSel.End

            Return True

        Catch ex As System.Exception
            MessageBox.Show("Failed to retrieve the selection: " & ex.Message,
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    Private Async Sub ShowSumup(selectedtext As String)

        Dim LLMResult As String = ""

        LLMResult = Await LLM(InterpolateAtRuntime(SP_MailSumup), "<MAILCHAIN>" & selectedtext & "</MAILCHAIN>", "", "", 0)

        If INI_PostCorrection <> "" Then
            LLMResult = Await PostCorrection(LLMResult)
        End If

        Dim markdownPipeline As MarkdownPipeline = New MarkdownPipelineBuilder().Build()
        Dim htmlText As String = Markdown.ToHtml(LLMResult, markdownPipeline)

        ShowHTMLCustomMessageBox(htmlText, $"{AN} Sum-up")

    End Sub

    Private Async Sub ShowSumup2(selectedtext As String)

        Dim LLMResult As String = ""

        DateTimeNow = DateTime.Now.ToString("yyyy-MMM-dd HH:mm")

        LLMResult = Await LLM(InterpolateAtRuntime(SP_MailSumup2), selectedtext, "", "", 0)

        If INI_PostCorrection <> "" Then
            LLMResult = Await PostCorrection(LLMResult)
        End If

        Dim markdownPipeline As MarkdownPipeline = New MarkdownPipelineBuilder().Build()
        Dim htmlText As String = Markdown.ToHtml(LLMResult, markdownPipeline)

        ShowHTMLCustomMessageBox(htmlText, $"{AN} Sum-up (multimail)")

    End Sub



    Private Async Sub FreeStyle_InsertBefore(Command As String, Optional AskForPrompt As Boolean = False)
        Try
            Dim outlookApp As New Microsoft.Office.Interop.Outlook.Application()
            Dim inspector As Inspector = outlookApp.ActiveInspector()

            ' Ensure the inspector is open and the item is a MailItem
            If inspector Is Nothing OrElse Not TypeOf inspector.CurrentItem Is MailItem Then
                SLib.ShowCustomMessageBox($"Please create or open an email for editing to use {AN}.")
                Return
            End If

            Dim mailItem As MailItem = DirectCast(inspector.CurrentItem, MailItem)

            ' Check if the email is in plain text format
            If mailItem.BodyFormat = OlBodyFormat.olFormatPlain Then
                SLib.ShowCustomMessageBox("This operation is not supported for plain text emails. Switch to HTML or RTF format.")
                Return
            End If

            ' Get the Word editor for the email
            Dim wordEditor As Microsoft.Office.Interop.Word.Document = TryCast(inspector.WordEditor, Microsoft.Office.Interop.Word.Document)

            If wordEditor Is Nothing Then
                SLib.ShowCustomMessageBox("Unable to access the necessary email editor. Ensure the email is in HTML or RTF format.")
                Return
            End If

            ' Get the selected text
            Dim selectedText As String = wordEditor.Application.Selection.Text
            If String.IsNullOrWhiteSpace(selectedText) Then
                selectedText = wordEditor.Content.Text
            End If

            OtherPrompt = ""
            Dim LLMResult As String = ""

            If AskForPrompt Then
                ' Prompt for additional instructions
                OtherPrompt = SLib.ShowCustomInputBox("Please provide additional information and instructions for drafting an answer:", $"{AN} Answers", False)
                If String.IsNullOrEmpty(OtherPrompt) Then Return

                ' Call your LLM function with the selected text
                LLMResult = Await LLM(InterpolateAtRuntime(SP_MailReply), "<MAILCHAIN>" & selectedText & "</MAILCHAIN>", "", "", 0)
            Else
                LLMResult = Await LLM(InterpolateAtRuntime(SP_MailSumup), "<MAILCHAIN>" & selectedText & "</MAILCHAIN>", "", "", 0)
            End If
            If INI_PostCorrection <> "" Then
                LLMResult = Await PostCorrection(LLMResult)
            End If

            'LLMResult = LLMResult.Replace("**", "")  ' Remove bold markers

            ' Convert Markdown to HTML using Markdig
            Dim markdownPipeline As MarkdownPipeline = New MarkdownPipelineBuilder().Build()
            Dim convertedHtml As String = Markdown.ToHtml(LLMResult, markdownPipeline)

            If mailItem.BodyFormat = OlBodyFormat.olFormatHTML Then
                ' Ensure consistent font and style for HTML emails
                Dim defaultStyle As String = "<div style='font-family:Arial, sans-serif; font-size:11pt;'>" ' Adjust as needed
                Dim formattedResult As String = defaultStyle & convertedHtml & "</div><br/><br/>"

                ' Append the formatted result to the HTML body
                mailItem.HTMLBody = formattedResult & mailItem.HTMLBody
            Else
                ' Convert HTML to plain text for non-HTML formats (optional)
                Dim doc As New HtmlAgilityPack.HtmlDocument()
                doc.LoadHtml(convertedHtml)
                Dim plainTextResult As String = doc.DocumentNode.InnerText

                ' Standard handling for Plain Text and Rich Text
                mailItem.Body = plainTextResult & vbCrLf & vbCrLf & mailItem.Body
            End If

            ' Refresh the inspector to show updated content
            inspector.Display()

            ' Release COM objects in reverse order of creation
            If wordEditor IsNot Nothing Then Marshal.ReleaseComObject(wordEditor) : wordEditor = Nothing
            If mailItem IsNot Nothing Then Marshal.ReleaseComObject(mailItem) : mailItem = Nothing
            If inspector IsNot Nothing Then Marshal.ReleaseComObject(inspector) : inspector = Nothing
            If outlookApp IsNot Nothing Then Marshal.ReleaseComObject(outlookApp) : outlookApp = Nothing

        Catch ex As System.Exception
            MessageBox.Show("Error in Freestyle_InsertBefore: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Async Sub Command_InsertAfter(ByVal SysCommand As String, Optional ByVal DoMarkup As Boolean = False, Optional KeepFormat As Boolean = False, Optional Inplace As Boolean = False, Optional MarkupMethod As Integer = 3)
        Try
            Dim outlookApp As New Microsoft.Office.Interop.Outlook.Application()
            Dim inspector As Microsoft.Office.Interop.Outlook.Inspector = outlookApp.ActiveInspector()

            ' Ensure the inspector is open and the item is a MailItem
            If inspector Is Nothing OrElse Not TypeOf inspector.CurrentItem Is Microsoft.Office.Interop.Outlook.MailItem Then
                ShowCustomMessageBox("Please open an email to use this function.")
                Return
            End If

            Dim mailItem As Microsoft.Office.Interop.Outlook.MailItem = DirectCast(inspector.CurrentItem, Microsoft.Office.Interop.Outlook.MailItem)

            ' Check if the email is in plain text format
            If mailItem.BodyFormat = Microsoft.Office.Interop.Outlook.OlBodyFormat.olFormatPlain Then
                ShowCustomMessageBox("This operation is not supported for plain text emails. Switch to HTML or RTF format.")
                Return
            End If

            ' Get the Word editor for the email
            Dim wordEditor As Microsoft.Office.Interop.Word.Document = TryCast(inspector.WordEditor, Microsoft.Office.Interop.Word.Document)

            If wordEditor Is Nothing Then
                ShowCustomMessageBox("Unable to access the email editor. Ensure the email is in HTML or RTF format.")
                Return
            End If

            ' Get the selected text and range
            Dim selection As Microsoft.Office.Interop.Word.Selection = wordEditor.Application.Selection
            Dim range As Microsoft.Office.Interop.Word.Range = selection.Range.Duplicate ' Duplicate to preserve original
            Dim SelectedText As String

            If INI_KeepFormatCap > 0 Then If Len(selection.Text) > INI_KeepFormatCap Then KeepFormat = False

            If KeepFormat Then
                SelectedText = SLib.GetRangeHtml(selection.Range)
            Else
                SelectedText = selection.Text
            End If

            If String.IsNullOrWhiteSpace(SelectedText) Then
                ShowCustomMessageBox($"Please select the text to be processed.")
                Return
            End If

            If DoMarkup And MarkupMethod = 2 And Len(SelectedText) > INI_MarkupDiffCap Then
                Dim MarkupChange As Integer = SLib.ShowCustomYesNoBox($"The selected text exceeds the defined cap for the Diff markup method at {INI_MarkupDiffCap} chars (your selection has {Len(SelectedText)} chars). {If(KeepFormat, "This may be because HTML codes have been inserted to keep the formatting (you can turn this off in the settings). ", "")}. How do you want to continue?", "Use Diff in Window compare instead", "Use Diff")
                Select Case MarkupChange
                    Case 1
                        MarkupMethod = 3
                    Case 2
                        MarkupMethod = 2
                    Case Else
                        Exit Sub
                End Select
            End If

            Dim trailingCR As Boolean = SelectedText.EndsWith(vbCrLf)

            ' Call your LLM function with the selected text
            Dim LLMResult As String = Await LLM(SysCommand & If(KeepFormat, " " & SP_Add_KeepHTMLIntact, ""), "<TEXTTOPROCESS>" & SelectedText & "</TEXTTOPROCESS>", "", "", 0)

            LLMResult = LLMResult.Replace("<TEXTTOPROCESS>", "").Replace("</TEXTTOPROCESS>", "")

            If INI_PostCorrection <> "" Then
                LLMResult = Await PostCorrection(LLMResult)
            End If

            ' Replace the selected text with the processed result
            If Not String.IsNullOrWhiteSpace(LLMResult) Then
                If KeepFormat Then

                    Dim Plaintext As String = ""

                    SelectedText = selection.Text
                    SLib.InsertTextWithFormat(LLMResult, range, Inplace)
                    If DoMarkup Then
                        LLMResult = SLib.RemoveHTML(LLMResult)
                        If MarkupMethod <> 3 Then
                            range.Text = vbCrLf & vbCrLf & "MARKUP:" & vbCrLf & vbCrLf
                        End If
                        range.Collapse(WdCollapseDirection.wdCollapseEnd)
                        selection.SetRange(range.Start, selection.End)

                        CompareAndInsertText(SelectedText, LLMResult, MarkupMethod = 3, "This is the markup of the text inserted:", True)
                    End If

                Else

                    If Inplace Then
                        If Not trailingCR And LLMResult.EndsWith(ControlChars.Lf) Then LLMResult = LLMResult.TrimEnd(ControlChars.Lf)
                        If Not trailingCR And LLMResult.EndsWith(ControlChars.Cr) Then LLMResult = LLMResult.TrimEnd(ControlChars.Cr)
                        If DoMarkup And MarkupMethod <> 3 Then
                            selection.TypeText(LLMResult & vbCrLf & vbCrLf & "MARKUP:" & vbCrLf & vbCrLf)
                        Else
                            selection.TypeText(LLMResult)
                        End If
                    Else
                        selection.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
                        If DoMarkup And MarkupMethod <> 3 Then
                            'selection.TypeText(vbCrLf & LLMResult & vbCrLf & vbCrLf & "MARKUP:" & vbCrLf & vbCrLf)
                            SLib.InsertTextWithBoldMarkers(selection, vbCrLf & LLMResult & vbCrLf & vbCrLf & "MARKUP:" & vbCrLf & vbCrLf)
                        Else
                            'selection.TypeText(vbCrLf & LLMResult & vbCrLf)
                            SLib.InsertTextWithBoldMarkers(selection, vbCrLf & LLMResult & vbCrLf)

                        End If
                    End If

                    ' Use Find to locate the nearest line break backward and adjust selection
                    range = selection.Range
                    With range.Find
                        .Text = vbCrLf
                        .Forward = False
                        .MatchWildcards = False
                        If .Execute() Then
                            selection.SetRange(range.Start, selection.End)
                        End If
                    End With

                    ' Perform markup comparison and insertion if necessary
                    If DoMarkup Then
                        If MarkupMethod = 2 Or MarkupMethod = 3 Then
                            CompareAndInsertText(SelectedText, LLMResult, MarkupMethod = 3, "This is the markup of the text inserted:", True)
                        Else
                            CompareAndInsertTextCompareDocs(SelectedText, LLMResult)
                        End If

                    End If

                End If

            Else
                ShowCustomMessageBox("The LLM did not return any content to insert.")

            End If

            ' Refresh the inspector to show updated content
            inspector.Display()

            ' Release COM objects in reverse order of creation
            If range IsNot Nothing Then Marshal.ReleaseComObject(range) : range = Nothing
            If selection IsNot Nothing Then Marshal.ReleaseComObject(selection) : selection = Nothing
            If wordEditor IsNot Nothing Then Marshal.ReleaseComObject(wordEditor) : wordEditor = Nothing
            If mailItem IsNot Nothing Then Marshal.ReleaseComObject(mailItem) : mailItem = Nothing
            If inspector IsNot Nothing Then Marshal.ReleaseComObject(inspector) : inspector = Nothing
            If outlookApp IsNot Nothing Then Marshal.ReleaseComObject(outlookApp) : outlookApp = Nothing

        Catch ex As System.Exception
            MessageBox.Show("Error in Command_InsertAfter: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Async Sub FreeStyle_InsertAfter()
        Try

            Dim DoMarkup As Boolean = False
            Dim DoInplace As Boolean = False
            Dim DoClipboard As Boolean = False
            Dim NoText As Boolean = False
            Dim MarkupMethod As Integer = INI_MarkupMethodOutlook
            Dim KeepFormatCap = INI_KeepFormatCap ' currently not used
            Dim DoKeepFormat As Boolean = INI_KeepFormat2 ' currently not used
            Dim DoKeepParaFormat As Boolean = INI_KeepParaFormatInline ' currently not used

            Dim UseSecondAPI As Boolean = False

            Dim MarkupInstruct As String = $"start with '{MarkupPrefixAll}' for markups"
            Dim InplaceInstruct As String = $"use '{InPlacePrefix}' for replacing the selection"
            Dim ClipboardInstruct As String = $"with '{ClipboardPrefix}' for separate output"
            Dim PromptLibInstruct As String = If(INI_PromptLib, " or press 'OK' for the prompt library", "")
            Dim NoFormatInstruct As String = $"; add '{NoFormatTrigger2}'/'{KFTrigger2}'/'{KPFTrigger2}' for overriding formatting defaults"
            Dim SecondAPIInstruct As String = If(INI_SecondAPI, $"'{SecondAPICode}' to use the secondary model ({INI_Model_2})", "")
            Dim LastPromptInstruct As String = If(String.IsNullOrWhiteSpace(My.Settings.LastPrompt), "", "; Ctrl-P for your last prompt")

            Dim AddOnInstruct As String = "; add " & SecondAPIInstruct

            Dim lastCommaIndex As Integer = AddOnInstruct.LastIndexOf(","c)
            If lastCommaIndex <> -1 Then
                AddOnInstruct = AddOnInstruct.Substring(0, lastCommaIndex) & ", and" & AddOnInstruct.Substring(lastCommaIndex + 1)
            End If

            Dim outlookApp As New Microsoft.Office.Interop.Outlook.Application()
            Dim inspector As Microsoft.Office.Interop.Outlook.Inspector = outlookApp.ActiveInspector()

            ' Ensure the inspector is open and the item is a MailItem
            If inspector Is Nothing OrElse Not TypeOf inspector.CurrentItem Is Microsoft.Office.Interop.Outlook.MailItem Then
                SLib.ShowCustomMessageBox($"Please create or open an email for editing to use {AN}.")
                Return
            End If

            Dim mailItem As Microsoft.Office.Interop.Outlook.MailItem = DirectCast(inspector.CurrentItem, Microsoft.Office.Interop.Outlook.MailItem)

            ' Check if the email is in plain text format
            If mailItem.BodyFormat = Microsoft.Office.Interop.Outlook.OlBodyFormat.olFormatPlain Then
                SLib.ShowCustomMessageBox("This operation is not supported for plain text emails. Switch to HTML or RTF format.")
                Return
            End If

            ' Get the Word editor for the email
            Dim wordEditor As Microsoft.Office.Interop.Word.Document = TryCast(inspector.WordEditor, Microsoft.Office.Interop.Word.Document)

            If wordEditor Is Nothing Then
                SLib.ShowCustomMessageBox("Unable to access the necessary email editor. Ensure the email is in HTML or RTF format.")
                Return
            End If

            ' Get the selected text
            Dim selection As Microsoft.Office.Interop.Word.Selection = wordEditor.Application.Selection
            Dim selectedText As String = selection.Text
            If String.IsNullOrWhiteSpace(selectedText) Then
                NoText = True
            End If

            ' Prompt for the text to process

            'SLib.StoreClipboard()

            'If Not String.IsNullOrWhiteSpace(My.Settings.LastPrompt) Then SLib.PutInClipboard(My.Settings.LastPrompt)

            If Not NoText Then
                OtherPrompt = SLib.ShowCustomInputBox($"Please provide the prompt you wish to execute on the selected text ({MarkupInstruct}, {InplaceInstruct}, {ClipboardInstruct}){PromptLibInstruct}{AddOnInstruct}{LastPromptInstruct}:", $"{AN} Freestyle", False, "", My.Settings.LastPrompt)
            Else
                OtherPrompt = SLib.ShowCustomInputBox($"Please provide the prompt you wish to execute ({ClipboardInstruct}){PromptLibInstruct}{AddOnInstruct}{LastPromptInstruct}:", $"{AN} Freestyle", False, "", My.Settings.LastPrompt)
            End If

            'SLib.RestoreClipboard()

            If String.IsNullOrEmpty(OtherPrompt) And OtherPrompt <> "ESC" And INI_PromptLib Then

                Dim promptlibresult As (String, Boolean, Boolean, Boolean)

                promptlibresult = ShowPromptSelector(INI_PromptLibPath, Not NoText, Nothing)

                OtherPrompt = promptlibresult.Item1
                DoMarkup = promptlibresult.Item2
                DoClipboard = promptlibresult.Item4

                If OtherPrompt = "" Then
                    Exit Sub
                End If
            Else
                If String.IsNullOrEmpty(OtherPrompt) Or OtherPrompt = "ESC" Then Exit Sub
            End If

            My.Settings.LastPrompt = OtherPrompt
            My.Settings.Save()

            ' Check if otherPrompt starts with "Markup:" (case-insensitive)

            If OtherPrompt.StartsWith(ClipboardPrefix, StringComparison.OrdinalIgnoreCase) Then
                OtherPrompt = OtherPrompt.Substring(ClipboardPrefix.Length).Trim()
                DoClipboard = True
            ElseIf OtherPrompt.StartsWith(MarkupPrefix, StringComparison.OrdinalIgnoreCase) And Not NoText Then
                OtherPrompt = OtherPrompt.Substring(MarkupPrefix.Length).Trim()
                DoMarkup = True
            ElseIf OtherPrompt.StartsWith(MarkupPrefixWord, StringComparison.OrdinalIgnoreCase) And Not NoText Then
                OtherPrompt = OtherPrompt.Substring(MarkupPrefixWord.Length).Trim()
                DoMarkup = True
                MarkupMethod = 1
            ElseIf OtherPrompt.StartsWith(MarkupPrefixDiffW, StringComparison.OrdinalIgnoreCase) And Not NoText Then
                OtherPrompt = OtherPrompt.Substring(MarkupPrefixDiffW.Length).Trim()
                DoMarkup = True
                MarkupMethod = 3
            ElseIf OtherPrompt.StartsWith(MarkupPrefixDiff, StringComparison.OrdinalIgnoreCase) And Not NoText Then
                OtherPrompt = OtherPrompt.Substring(MarkupPrefixDiff.Length).Trim()
                DoMarkup = True
                MarkupMethod = 2
            ElseIf OtherPrompt.StartsWith(InPlacePrefix, StringComparison.OrdinalIgnoreCase) And Not NoText Then
                OtherPrompt = OtherPrompt.Substring(InPlacePrefix.Length).Trim()
                DoMarkup = False
                MarkupMethod = 3
                DoInplace = True
            End If

            ' Formatting Trigger (currently not used)

            If OtherPrompt.IndexOf(NoFormatTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
                OtherPrompt = OtherPrompt.Replace(NoFormatTrigger, "").Trim()
                KeepFormatCap = 1
            End If
            If OtherPrompt.IndexOf(NoFormatTrigger2, StringComparison.OrdinalIgnoreCase) >= 0 Then
                OtherPrompt = OtherPrompt.Replace(NoFormatTrigger2, "").Trim()
                KeepFormatCap = 1
            End If
            If OtherPrompt.IndexOf(KFTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
                OtherPrompt = OtherPrompt.Replace(KFTrigger, "").Trim()
                DoKeepFormat = True
            End If
            If OtherPrompt.IndexOf(KFTrigger2, StringComparison.OrdinalIgnoreCase) >= 0 Then
                OtherPrompt = OtherPrompt.Replace(KFTrigger2, "").Trim()
                DoKeepFormat = True
            End If
            If OtherPrompt.IndexOf(KPFTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
                OtherPrompt = OtherPrompt.Replace(KPFTrigger, "").Trim()
                DoKeepParaFormat = True
            End If
            If OtherPrompt.IndexOf(KPFTrigger2, StringComparison.OrdinalIgnoreCase) >= 0 Then
                OtherPrompt = OtherPrompt.Replace(KPFTrigger2, "").Trim()
                DoKeepParaFormat = True
            End If

            If INI_SecondAPI Then
                If OtherPrompt.Contains(SecondAPICode) Then
                    UseSecondAPI = True
                    OtherPrompt = OtherPrompt.Replace(SecondAPICode, "").Trim()
                End If
            End If

            If DoMarkup And MarkupMethod = 2 And Len(selectedText) > INI_MarkupDiffCap Then
                Dim MarkupChange As Integer = SLib.ShowCustomYesNoBox($"The selected text exceeds the defined cap for the Diff markup method at {INI_MarkupDiffCap} chars (your selection has {Len(selectedText)} chars). How do you want to continue?", "Use Diff in Window compare instead", "Use Diff")
                Select Case MarkupChange
                    Case 1
                        MarkupMethod = 3
                    Case 2
                        MarkupMethod = 2
                    Case Else
                        Exit Sub
                End Select
            End If

            Dim trailingCR As Boolean = selectedText.EndsWith(vbCrLf)

            ' Call your LLM function with the selected text

            Dim LLMResult As String

            If Not NoText Then
                LLMResult = Await LLM(InterpolateAtRuntime(SP_FreestyleText), "<TEXTTOPROCESS>" & selectedText & "</TEXTTOPROCESS>", "", "", 0, UseSecondAPI, False, OtherPrompt)

                LLMResult = LLMResult.Replace("<TEXTTOPROCESS>", "").Replace("</TEXTTOPROCESS>", "")
            Else
                LLMResult = Await LLM(InterpolateAtRuntime(SP_FreestyleNoText), "", "", "", 0, UseSecondAPI, False, OtherPrompt)
            End If

            If INI_PostCorrection <> "" Then
                LLMResult = Await PostCorrection(LLMResult)
            End If

            OtherPrompt = ""

            If DoClipboard Then
                Dim FinalText As String = SLib.ShowCustomWindow("The LLM has provided the following result (you can edit it):", LLMResult, "You can choose whether you want to have the original text put into the clipboard or your text with any changes you have made. If you select Cancel, nothing will be put into the clipboard (without formatting).", AN, True)

                If FinalText <> "" Then
                    SLib.PutInClipboard(FinalText)
                End If
            Else
                ' Collapse the selection to the end

                If Not DoInplace Then
                    selection.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
                Else
                    If Not trailingCR And LLMResult.EndsWith(ControlChars.Lf) Then LLMResult = LLMResult.TrimEnd(ControlChars.Lf)
                    If Not trailingCR And LLMResult.EndsWith(ControlChars.Cr) Then LLMResult = LLMResult.TrimEnd(ControlChars.Cr)
                End If

                ' Insert the result as a new paragraph
                If DoMarkup And MarkupMethod <> 3 Then
                    SLib.InsertTextWithBoldMarkers(selection, vbCrLf & LLMResult & vbCrLf & vbCrLf & "MARKUP:" & vbCrLf & vbCrLf)
                Else
                    If DoInplace Then
                        SLib.InsertTextWithBoldMarkers(selection, LLMResult)
                    Else
                        SLib.InsertTextWithBoldMarkers(selection, vbCrLf & LLMResult & vbCrLf)
                    End If
                End If

                ' Use Find to locate the nearest line break backward and adjust selection
                Dim range As Microsoft.Office.Interop.Word.Range = selection.Range
                With range.Find
                    .Text = vbCrLf
                    .Forward = False
                    .MatchWildcards = False
                    If .Execute() Then
                        selection.SetRange(range.Start, selection.End)
                    End If
                End With

                ' Perform markup comparison and insertion if necessary
                If DoMarkup Then
                    If MarkupMethod = 2 Or MarkupMethod = 3 Then
                        CompareAndInsertText(selectedText, LLMResult, MarkupMethod = 3, "This is the markup of the text inserted:", True)
                    Else
                        CompareAndInsertTextCompareDocs(selectedText, LLMResult)
                    End If
                End If
            End If

            ' Refresh the inspector to show updated content
            inspector.Display()

            ' Release COM objects in reverse order of creation
            If selection IsNot Nothing Then Marshal.ReleaseComObject(selection) : selection = Nothing
            If wordEditor IsNot Nothing Then Marshal.ReleaseComObject(wordEditor) : wordEditor = Nothing
            If mailItem IsNot Nothing Then Marshal.ReleaseComObject(mailItem) : mailItem = Nothing
            If inspector IsNot Nothing Then Marshal.ReleaseComObject(inspector) : inspector = Nothing
            If outlookApp IsNot Nothing Then Marshal.ReleaseComObject(outlookApp) : outlookApp = Nothing

        Catch ex As System.Exception
            MessageBox.Show("Error in Freestyle_InsertAfter: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Private Sub CompareAndInsertTextCompareDocs(input1 As String, input2 As String)

        Dim splash As New SLib.SplashScreen("Creating markup using the Word compare functionality (ignore any flickering and press 'No' if prompted) ...")
        splash.Show()
        splash.Refresh()
        Try
            ' Get the active inspector (compose mail window)
            Dim outlookApp As Microsoft.Office.Interop.Outlook.Application = New Microsoft.Office.Interop.Outlook.Application()
            Dim inspector As Inspector = outlookApp.ActiveInspector

            ' Ensure the current item is a MailItem and in compose mode
            If TypeOf inspector.CurrentItem Is MailItem Then
                Dim mailItem As MailItem = CType(inspector.CurrentItem, MailItem)
                Dim editor As Object = inspector.WordEditor

                ' Cast the WordEditor to Word.Document
                Dim wordDoc As Document = CType(editor, Document)

                ' Create a new temporary Word application for comparison
                Dim wordApp As New Microsoft.Office.Interop.Word.Application()
                wordApp.Visible = False

                ' Create temporary documents for input1 and input2
                Dim tempDoc1 As Document = wordApp.Documents.Add()
                Dim tempDoc2 As Document = wordApp.Documents.Add()

                ' Insert the input texts into the temporary documents
                tempDoc1.Content.Text = input1
                tempDoc2.Content.Text = input2

                ' Perform the comparison
                Dim compareResult As Document = wordApp.CompareDocuments(tempDoc1, tempDoc2,
                                                            WdCompareDestination.wdCompareDestinationNew,
                                                            WdGranularity.wdGranularityWordLevel,
                                                            False, False, False, False, False, False)

                ' Convert tracked changes to static formatting
                For Each revision As Revision In compareResult.Revisions
                    Select Case revision.Type
                        Case WdRevisionType.wdRevisionInsert
                            ' Insertions: Apply blue color and underline
                            revision.Range.Font.Color = WdColor.wdColorBlue
                            revision.Range.Font.Underline = WdUnderline.wdUnderlineSingle
                        Case WdRevisionType.wdRevisionDelete
                            ' Deletions: Apply red color and strikethrough
                            revision.Range.Font.Color = WdColor.wdColorRed
                            revision.Range.Font.StrikeThrough = True
                    End Select
                    revision.Accept() ' Accept the revision to make the formatting static
                Next

                ' Copy the comparison result to clipboard
                compareResult.Content.Copy()

                ' Paste the comparison result into the Outlook compose window at the current selection
                wordDoc.Application.Selection.PasteAndFormat(WdRecoveryType.wdFormatOriginalFormatting)

                ' Clean up
                tempDoc1.Close(False)
                tempDoc2.Close(False)
                compareResult.Close(False)
                wordApp.Quit(False)

            Else
                MessageBox.Show("Error in CompareAndInsertTextCompareDocs: The mail compose window is not open (anymore).", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If

            ' Release COM objects in reverse order of creation
            If inspector IsNot Nothing Then Marshal.ReleaseComObject(inspector) : inspector = Nothing
            If outlookApp IsNot Nothing Then Marshal.ReleaseComObject(outlookApp) : outlookApp = Nothing

        Catch ex As System.Exception
            MessageBox.Show("Error in CompareAndInsertTextCompareDocs: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Finally
            splash.Close()

        End Try
    End Sub

    Private Sub CompareAndInsertText(text1 As String, text2 As String, Optional ShowInWindow As Boolean = False, Optional TextforWindow As String = "A text with these changes will be inserted ('Esc' to abort):", Optional DoNotWait As Boolean = False)
        Dim diffBuilder As New InlineDiffBuilder(New Differ())
        Dim sText As String = String.Empty

        ' Pre-process the texts to replace line breaks with a unique marker
        text1 = text1.Replace(vbCrLf, " {vbCrLf} ").Replace(vbCr, " {vbCr} ").Replace(vbLf, " {vbLf} ")
        text2 = text2.Replace(vbCrLf, " {vbCrLf} ").Replace(vbCr, " {vbCr} ").Replace(vbLf, " {vbLf} ")

        ' Normalize the texts by removing extra spaces
        text1 = text1.Replace("  ", " ").Trim()
        text2 = text2.Replace("  ", " ").Trim()

        ' Split the texts into words and convert them into a line-by-line format
        Dim words1 As String = String.Join(Environment.NewLine, text1.Split(" "c))
        Dim words2 As String = String.Join(Environment.NewLine, text2.Split(" "c))

        ' Generate word-based diff using DiffPlex
        Dim diffResult As DiffPaneModel = diffBuilder.BuildDiffModel(words1, words2)

        ' Build the formatted output based on the diff results
        For Each line In diffResult.Lines
            Select Case line.Type
                Case ChangeType.Inserted
                    sText &= "[INS_START]" & line.Text.Trim() & "[INS_END] "
                Case ChangeType.Deleted
                    sText &= "[DEL_START]" & line.Text.Trim() & "[DEL_END] "
                Case ChangeType.Unchanged
                    sText &= line.Text.Trim() & " "
            End Select
        Next

        ' Remove preceding and trailing spaces around placeholders

        sText = sText.Replace("{vbCr}", "{vbCrLf}")
        sText = sText.Replace("{vbLf}", "{vbCrLf}")
        sText = sText.Replace(" {vbCrLf} ", "{vbCrLf}")
        sText = sText.Replace(" {vbCrLf}", "{vbCrLf}")
        sText = sText.Replace("{vbCrLf} ", "{vbCrLf}")

        ' Remove instances of line breaks surrounded by [DEL_START] and [DEL_END]
        sText = sText.Replace("[DEL_START]{vbCrLf}[DEL_END] ", "")
        'sText = Regex.Replace(sText, "\[DEL_START\](.*?)\s*{vbCrLf}\s*(.*?)\[DEL_END\]", Function(m) $"[DEL_START]{m.Groups(1).Value}{m.Groups(2).Value}[DEL_END] ")

        ' Include instances of line breaks surrounded by [INS_START] and [INS_END] without the [INS...] text
        sText = sText.Replace("[INS_START]{vbCrLf}[INS_END] ", "{vbCrLf}")

        ' Replace placeholders with actual line breaks
        sText = sText.Replace("{vbCrLf}", vbCrLf)

        ' Adjust overlapping tags
        sText = sText.Replace("[DEL_END] [INS_START]", "[DEL_END][INS_START]")
        sText = sText.Replace("[INS_START][INS_END] ", "")

        ' Insert formatted text into the Word editor
        If Not ShowInWindow Then
            InsertFormattedText(sText & vbCrLf)
        Else
            Dim htmlContent As String = ConvertMarkupToRTF(TextforWindow & "\r\r" & sText)
            System.Threading.Tasks.Task.Run(Sub()
                                                ShowRTFCustomMessageBox(htmlContent)
                                            End Sub)
        End If

    End Sub


    Private Function ConvertRtfToPlainText(rtfContent As String) As String
        If String.IsNullOrWhiteSpace(rtfContent) Then Return String.Empty

        ' Remove RTF headers and control words
        Dim plainText As String = Regex.Replace(rtfContent, "{\\.*?}|\\[a-z]+[0-9]*|[{}]", String.Empty)

        ' Decode escaped characters (e.g., \'xx)
        plainText = Regex.Replace(plainText, "\\'([0-9a-fA-F]{2})", Function(m)
                                                                        Dim hex = Convert.ToByte(m.Groups(1).Value, 16)
                                                                        Return Chr(hex)
                                                                    End Function)

        ' Replace RTF line breaks (\par) with actual line breaks
        plainText = Regex.Replace(plainText, "\\par", Environment.NewLine, RegexOptions.IgnoreCase)

        ' Trim the result
        plainText = plainText.Trim()

        Return plainText
    End Function

    Private Sub InsertFormattedText(inputText As String)
        Dim objInspector As Microsoft.Office.Interop.Outlook.Inspector
        Dim objWordDoc As Microsoft.Office.Interop.Word.Document
        Dim objSelection As Object
        Dim objRange As Object
        Dim TextArray() As String = {}
        Dim FormatArray() As Integer = {}
        Dim i As Integer

        ' Store original font properties
        Dim originalFontColor As Integer = 0
        Dim originalUnderline As Integer = 0
        Dim originalStrikeThrough As Boolean = False
        Dim originalBold As Boolean = False
        Dim originalItalic As Boolean = False

        ' Check if there is an active inspector (open email)
        objInspector = TryCast(Globals.ThisAddIn.Application.ActiveInspector, Microsoft.Office.Interop.Outlook.Inspector)
        If objInspector Is Nothing Then
            MessageBox.Show("Error in InsertFormattedText: No open mail item found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        ' Get the Word editor and the current selection
        objWordDoc = TryCast(objInspector.WordEditor, Microsoft.Office.Interop.Word.Document)
        If objWordDoc Is Nothing Then
            MessageBox.Show("Error in InsertFormattedText: Unable to access the necessary mail editor for this mail.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        objSelection = objWordDoc.Application.Selection

        ' Store original font properties
        If objSelection.Font IsNot Nothing Then
            With objSelection.Font
                originalFontColor = .Color
                originalUnderline = .Underline
                originalStrikeThrough = .StrikeThrough
                originalBold = .Bold
                originalItalic = .Italic
            End With
        End If

        Dim splash As New SplashScreen("Creating your markup ... press 'Esc' to abort")
        splash.Show()
        splash.Refresh()

        ' Parse the input text into chunks with formatting information
        ParseText(inputText, TextArray, FormatArray)

        ' Reset formatting before starting
        If objSelection.Font IsNot Nothing Then objSelection.Font.Reset()

        ' Insert each text chunk with the appropriate formatting
        For i = 0 To TextArray.Length - 1

            System.Windows.Forms.Application.DoEvents()

            If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And &H8000) <> 0 Then
                Exit For
            End If

            If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And 1) <> 0 Then
                ' Exit the loop
                Exit For
            End If


            ' Reset formatting to original before each insertion
            If objSelection.Font IsNot Nothing Then
                With objSelection.Font
                    .Color = originalFontColor
                    .Underline = originalUnderline
                    .StrikeThrough = originalStrikeThrough
                    .Bold = originalBold
                    .Italic = originalItalic
                End With
            End If

            ' Insert the text at the current cursor position
            objSelection.Collapse(0) ' Collapse to insertion point
            objSelection.TypeText(TextArray(i))

            ' Define the range for the inserted text
            objRange = objSelection.Range
            objRange.Start = objSelection.Start - TextArray(i).Length
            objRange.End = objSelection.Start

            ' Apply formatting based on the tag
            Select Case FormatArray(i)
                Case 1 ' [INS_START]...[INS_END]: Blue underline
                    If objRange.Font IsNot Nothing Then
                        With objRange.Font
                            .Color = RGB(0, 0, 255)
                            .Underline = True
                            .StrikeThrough = False
                        End With
                    End If
                Case 2 ' [DEL_START]...[DEL_END]: Red strikethrough
                    If objRange.Font IsNot Nothing Then
                        With objRange.Font
                            .Color = RGB(255, 0, 0)
                            .StrikeThrough = True
                            .Underline = False
                        End With
                    End If
                Case Else ' Normal text
                    ' Already reset to original formatting
            End Select
        Next

        ' Ensure formatting is reset after all insertions
        If objSelection.Font IsNot Nothing Then
            With objSelection.Font
                .Color = originalFontColor
                .Underline = originalUnderline
                .StrikeThrough = originalStrikeThrough
                .Bold = originalBold
                .Italic = originalItalic
            End With
        End If

        splash.Close()

        ' Release COM objects in reverse order of creation
        If objInspector IsNot Nothing Then Marshal.ReleaseComObject(objInspector) : objInspector = Nothing
        If objWordDoc IsNot Nothing Then Marshal.ReleaseComObject(objWordDoc) : objWordDoc = Nothing

    End Sub

    Private Sub ParseText(inputText As String, ByRef TextArray() As String, ByRef FormatArray() As Integer)
        Dim pos As Integer = 1
        Dim lenText As Integer = inputText.Length
        Dim nextTagPos As Integer
        Dim tagEndPos As Integer
        Dim tagText As String
        Dim chunkIndex As Integer = 0
        Dim tagType As Integer
        Dim nextInsPos As Integer
        Dim nextDelPos As Integer

        While pos <= lenText
            If inputText.Substring(pos - 1, System.Math.Min(11, lenText - pos + 1)) = "[INS_START]" Then
                pos += 11
                tagType = 1 ' Insert formatting
                tagEndPos = inputText.IndexOf("[INS_END]", pos - 1) + 1
                If tagEndPos = -1 Then
                    MessageBox.Show("Error in ParseText: Missing [INS_END] tag.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If
                tagText = inputText.Substring(pos - 1, tagEndPos - pos)
                pos = tagEndPos + 9
            ElseIf inputText.Substring(pos - 1, System.Math.Min(11, lenText - pos + 1)) = "[DEL_START]" Then
                pos += 11
                tagType = 2 ' Delete formatting
                tagEndPos = inputText.IndexOf("[DEL_END]", pos - 1) + 1
                If tagEndPos = -1 Then
                    MessageBox.Show("Error in ParseText: Missing [DEL_END] tag.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If
                tagText = inputText.Substring(pos - 1, tagEndPos - pos)
                pos = tagEndPos + 9
            Else
                tagType = 0
                nextInsPos = inputText.IndexOf("[INS_START]", pos - 1) + 1
                If nextInsPos = 0 Then nextInsPos = lenText + 1
                nextDelPos = inputText.IndexOf("[DEL_START]", pos - 1) + 1
                If nextDelPos = 0 Then nextDelPos = lenText + 1
                nextTagPos = System.Math.Min(nextInsPos, nextDelPos)
                tagText = inputText.Substring(pos - 1, nextTagPos - pos)
                pos = nextTagPos
            End If

            chunkIndex += 1
            ReDim Preserve TextArray(chunkIndex - 1)
            ReDim Preserve FormatArray(chunkIndex - 1)
            TextArray(chunkIndex - 1) = tagText
            FormatArray(chunkIndex - 1) = tagType
        End While
    End Sub

    Private Function RGB(ByVal red As Integer, ByVal green As Integer, ByVal blue As Integer) As Integer
        Return red Or (green << 8) Or (blue << 16)
    End Function


    Private Function GetSelectedTextLength() As Integer
        Try
            Dim outlookApp As New Microsoft.Office.Interop.Outlook.Application()
            Dim inspector As Microsoft.Office.Interop.Outlook.Inspector = outlookApp.ActiveInspector()

            ' Ensure the inspector is open and the item is a MailItem
            If inspector Is Nothing OrElse Not TypeOf inspector.CurrentItem Is Microsoft.Office.Interop.Outlook.MailItem Then
                Return 0
            End If

            Dim mailItem As Microsoft.Office.Interop.Outlook.MailItem =
            DirectCast(inspector.CurrentItem, Microsoft.Office.Interop.Outlook.MailItem)

            ' Check if the email is in plain text format
            If mailItem.BodyFormat = Microsoft.Office.Interop.Outlook.OlBodyFormat.olFormatPlain Then
                Return 0
            End If

            ' Get the Word editor for the email
            Dim wordEditor As Microsoft.Office.Interop.Word.Document =
            TryCast(inspector.WordEditor, Microsoft.Office.Interop.Word.Document)

            If wordEditor Is Nothing Then
                Return 0
            End If

            ' Get the selected text
            Dim selection As Microsoft.Office.Interop.Word.Selection = wordEditor.Application.Selection
            Dim selectedText As String = selection.Text

            If String.IsNullOrWhiteSpace(selectedText) Then
                Return 0
            End If

            ' Split on whitespace to count words;
            ' filter out empty entries in case of multiple spaces/newlines
            Dim words = selectedText.Split(New Char() {" "c, ControlChars.Tab, ControlChars.Cr, ControlChars.Lf},
                                       StringSplitOptions.RemoveEmptyEntries)
            Return words.Length

            ' Release COM objects in reverse order of creation
            If selection IsNot Nothing Then Marshal.ReleaseComObject(selection) : selection = Nothing
            If wordEditor IsNot Nothing Then Marshal.ReleaseComObject(wordEditor) : wordEditor = Nothing
            If mailItem IsNot Nothing Then Marshal.ReleaseComObject(mailItem) : mailItem = Nothing
            If inspector IsNot Nothing Then Marshal.ReleaseComObject(inspector) : inspector = Nothing
            If outlookApp IsNot Nothing Then Marshal.ReleaseComObject(outlookApp) : outlookApp = Nothing

        Catch ex As System.Exception  ' Explicitly referencing System.Exception per your guideline
            Return 0
        End Try
    End Function

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

    Public Sub ShowSettings()

        Dim Settings As New Dictionary(Of String, String) From {
                        {"Temperature", "Temperature of {model}"},
                        {"Timeout", "Timeout of {model}"},
                        {"Temperature_2", "Temperature of {model2}"},
                        {"Timeout_2", "Timeout of {model2}"},
                        {"DoubleS", "Convert '" & ChrW(223) & "' to 'ss'"},
                        {"KeepFormat1", "Keep format (translations)"},
                        {"ReplaceText1", "Replace text (translations)"},
                        {"KeepFormat2", "Keep format (other commands)"},
                        {"ReplaceText2", "Replace text (other commands)"},
                        {"DoMarkupOutlook", "Also do a markup (other commands)"},
                        {"MarkupMethodOutlook", "Markup method (1 = Word, 2 = Diff, 3 = DiffW)"},
                        {"MarkupDiffCap", "Maximum characters for Diff Markup"},
                        {"PreCorrection", "Additional instruction for prompts"},
                        {"PostCorrection", "Prompt to apply after queries"},
                        {"Language1", "Default translation language"},
                        {"PromptLibPath", "Prompt library file"}
                    }

        Dim SettingsTips As New Dictionary(Of String, String) From {
                        {"Temperature", "The higher, the more creative the LLM will be (0.0-2.0)"},
                        {"Timeout", "In milliseconds"},
                        {"Temperature_2", "The higher, the more creative the LLM will be (0.0-2.0)"},
                        {"Timeout_2", "In milliseconds"},
                        {"DoubleS", "For Switzerland"},
                        {"KeepFormat1", "If selected, the original's text basic formatting of a translated text will be retained (by HTML encoding, takes time!)"},
                        {"ReplaceText1", "If selected, the response of the LLM for translations will replace the original text"},
                        {"KeepFormat2", "If selected, the original's text basic formatting of other text (other than translations) will be retained (by HTML encoding, takes time!)"},
                        {"ReplaceText2", "If selected, the response of the LLM for other commands (than translate) will replace the original text"},
                        {"DoMarkupOutlook", "Whether a markup should be done for functions that change only parts of a text"},
                        {"MarkupMethodOutlook", "Markup method to use: 1 = Compare using the Word compare function, 2 = Simple Differ, 3 = Simple Diff shown in a window"},
                        {"MarkupDiffCap", "The maximum size of the text that should be processed using the Diff method (to avoid you having to wait too long)"},
                        {"PreCorrection", "Add prompting text that will be added to all basic requests (e.g., for special language tasks)"},
                        {"PostCorrection", "Add a prompt that will be applied to each result before it is further processed (slow!)"},
                        {"Language1", "The language (in English) that will be used for the quick access button in the ribbon"},
                        {"PromptLibPath", "The filename (including path, support environmental variables) for your prompt library (if any)"}
                    }

        ShowSettingsWindow(Settings, SettingsTips)

        Globals.Ribbons.Ribbon1.UpdateRibbon()
        Globals.Ribbons.Ribbon2.UpdateRibbon()

    End Sub

    ' WebExtension integration

    Private httpListener As HttpListener
    Private listenerThread As Thread
    Private isShuttingDown As Boolean = False


    Private Sub StartupHttpListener()
        ' Start the HTTP listener on a background thread.

        listenerThread = New Thread(Async Sub()
                                        ' Await the function to ensure proper handling
                                        Await StartHttpListener()
                                    End Sub)

        listenerThread.IsBackground = True
        listenerThread.Start()
    End Sub


    Private Sub ShutdownHttpListener()
        ' Cleanly stop the listener if it's running.
        isShuttingDown = True
        If httpListener IsNot Nothing AndAlso httpListener.IsListening Then
            httpListener.Stop()
            httpListener.Close()
        End If
    End Sub

    Private Async Function NewStartHttpListener() As Task(Of String)
        Try
            httpListener = New HttpListener()
            httpListener.Prefixes.Add("http://127.0.0.1:12333/")
            httpListener.Start()

            isShuttingDown = False

            While Not isShuttingDown
                Dim context As HttpListenerContext = Await httpListener.GetContextAsync()
                ' Fire and forget: handle request in the background
                Dim result = System.Threading.Tasks.Task.Run(Async Function()
                                                                 Dim resultx = Await HandleHttpRequest(context)
                                                                 Return resultx
                                                             End Function)
            End While
        Catch ex As System.Exception
            ' Log or handle exceptions as needed
        End Try
    End Function



    Private Async Function StartHttpListener() As Task(Of String)
        Dim prefix As String = "http://127.0.0.1:12333/"
        Dim consecutiveFailures As Integer = 0

        Try
            ' Initialize the listener once.
            If httpListener Is Nothing Then
                httpListener = New HttpListener()
                httpListener.Prefixes.Add(prefix)
                httpListener.Start()
                Debug.WriteLine("HttpListener started.")
            End If

            While Not isShuttingDown
                Dim delayNeeded As Boolean = False

                ' If for some reason the listener is not active, restart it.
                If httpListener Is Nothing OrElse Not httpListener.IsListening Then
                    Try
                        If httpListener IsNot Nothing Then
                            httpListener.Close()
                        End If
                    Catch ex As System.Exception
                        Debug.WriteLine("Error closing HttpListener: " & ex.Message)
                    End Try

                    httpListener = New HttpListener()
                    httpListener.Prefixes.Add(prefix)
                    httpListener.Start()
                    Debug.WriteLine("HttpListener restarted.")
                End If

                Try
                    ' Asynchronously wait for an incoming request.
                    Dim context As HttpListenerContext = Await httpListener.GetContextAsync()
                    Dim result As String = Await HandleHttpRequest(context)
                    Debug.WriteLine("Request handled successfully.")
                    ' Reset the failure counter on success.
                    consecutiveFailures = 0
                Catch ex As System.ObjectDisposedException
                    Debug.WriteLine("HttpListener was disposed. Restarting listener...")
                    consecutiveFailures += 1
                    delayNeeded = True
                Catch ex As System.Exception
                    Debug.WriteLine("Error handling HTTP request: " & ex.Message)
                    consecutiveFailures += 1
                    delayNeeded = True
                End Try

                ' Check if we have reached the maximum number of consecutive failures.
                If consecutiveFailures >= 10 Then
                    Debug.WriteLine("Too many consecutive failures. Shutting down.")
                    isShuttingDown = True
                    Exit While
                End If

                ' If an error occurred, delay before restarting.
                If delayNeeded Then
                    Await System.Threading.Tasks.Task.Delay(5000)
                End If
            End While
        Catch ex As System.Exception
            Debug.WriteLine("Error in StartHttpListener: " & ex.Message)
        End Try

        Return ""
    End Function



    Private Async Function oldStartHttpListener() As Task(Of String)
        Dim prefix As String = "http://127.0.0.1:12333/"
        Try
            ' Initialize the listener once.
            If httpListener Is Nothing Then
                httpListener = New HttpListener()
                httpListener.Prefixes.Add(prefix)
                httpListener.Start()
                Debug.WriteLine("HttpListener started.")
            End If

            While Not isShuttingDown
                ' If for some reason the listener is not listening (disposed), restart it.
                If httpListener Is Nothing OrElse Not httpListener.IsListening Then
                    Try
                        If httpListener IsNot Nothing Then
                            httpListener.Close()
                        End If
                    Catch ex As System.Exception
                        Debug.WriteLine("Error closing HttpListener: " & ex.Message)
                    End Try

                    httpListener = New HttpListener()
                    httpListener.Prefixes.Add(prefix)
                    httpListener.Start()
                    Debug.WriteLine("HttpListener restarted.")
                End If

                Try
                    ' Use asynchronous call to wait for an incoming request.
                    Dim context As HttpListenerContext = Await httpListener.GetContextAsync()
                    Dim result As String = Await HandleHttpRequest(context)
                Catch ex As System.ObjectDisposedException
                    Debug.WriteLine("HttpListener was disposed. Restarting listener...")
                    ' Continue to the next iteration so that the above block restarts the listener.
                    Continue While
                Catch ex As System.Exception
                    Debug.WriteLine("Error httplistener handling request: " & ex.Message)
                End Try
            End While
        Catch ex As System.Exception
            Debug.WriteLine("Error in StartHttpListener: " & ex.Message)
        End Try
        Return ""
    End Function

    Private Async Function HandleHttpRequest(ByVal context As HttpListenerContext) As Task(Of String)
        Try
            ' 1) Retrieve the request
            Dim request As HttpListenerRequest = context.Request
            Dim response As HttpListenerResponse = context.Response

            ' Handle preflight (OPTIONS) request
            If request.HttpMethod = "OPTIONS" Then
                Debug.Print("Handling preflight (OPTIONS) request...")
                response.AddHeader("Access-Control-Allow-Origin", "*")
                response.AddHeader("Access-Control-Allow-Methods", "GET, POST, PUT, DELETE, OPTIONS")
                response.AddHeader("Access-Control-Allow-Headers", "Content-Type, Authorization")
                response.StatusCode = 204 ' No Content
                'response.OutputStream.Close()
                response.Close()
                Return ""
            End If

            ' Initialize request body variable
            Dim requestBody As String = String.Empty

            ' Handle entity body (for POST/PUT, etc.)
            If request.HasEntityBody Then
                Debug.Print("Processing request body...")
                Using reader As New StreamReader(request.InputStream, Encoding.UTF8)
                    requestBody = reader.ReadToEnd()
                End Using
                Debug.Print("Request Body: " & requestBody)
            End If

            ' 2) Process the request
            '    - Parse JSON or handle requestBody
            Dim responseText As String = Await ProcessRequestInAddIn(requestBody, request.RawUrl)

            ' 3) Write a response with CORS headers
            Dim buffer As Byte() = Encoding.UTF8.GetBytes(responseText)
            response.ContentLength64 = buffer.Length
            response.ContentType = "text/plain; charset=utf-8"
            response.AddHeader("Access-Control-Allow-Origin", "*") ' Allow cross-origin requests

            Using output As Stream = response.OutputStream
                output.Write(buffer, 0, buffer.Length)
            End Using
            response.Close()
            Debug.WriteLine("HTTP Request completed without errors.")
            Return ""

        Catch ex As System.Exception
            ' If there's an error, return an error response to the caller
            Try
                Dim errorStr = "Error: " & ex.Message
                Dim errorBytes = Encoding.UTF8.GetBytes(errorStr)
                context.Response.ContentLength64 = errorBytes.Length
                context.Response.StatusCode = 500  ' Internal server error
                context.Response.OutputStream.Write(errorBytes, 0, errorBytes.Length)
                'context.Response.OutputStream.Close()
                context.Response.Close()
            Catch
                ' If we can’t even write an error, just ignore
            End Try
            Debug.WriteLine("HTTP Request completed with errors.")
            Return ""
        End Try
    End Function

    Private Async Function RunLLMOnMainThread(sysprompt As String, userprompt As String) As Task(Of String)
        Dim tcs As New TaskCompletionSource(Of String)()

        ' Use Invoke to marshal work to the main thread
        mainThreadControl.Invoke(
        Sub()
            ' Run the LLM function on the main thread and set the TaskCompletionSource result
            Dim task = LLM(sysprompt, userprompt, "", "", 0)
            task.ContinueWith(
                Sub(t)
                    If t.IsFaulted Then
                        tcs.SetException(t.Exception)
                    ElseIf t.IsCanceled Then
                        tcs.SetCanceled()
                    Else
                        tcs.SetResult(t.Result)
                    End If
                End Sub)
        End Sub)

        ' Await the TaskCompletionSource's Task to ensure the result is returned
        Return Await tcs.Task
    End Function

    Private Sub RunCompareAndInsertTextOnMainThread(textbody As String, result As String, ShowInWindow As Boolean)
        mainThreadControl.Invoke(Sub()
                                     CompareAndInsertText(textbody, result, ShowInWindow) ' Blocking wait is safe here since we're on the main thread.
                                 End Sub)
    End Sub


    Private Async Function ProcessRequestInAddIn(requestBody As String, rawUrl As String) As Task(Of String)

        Dim result As String = ""

        Try
            ' Parse the JSON string
            Dim jsonObject As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.Linq.JObject.Parse(requestBody)

            Debug.WriteLine("Requestbody = " & requestBody)

            ' Check if the "command" segment contains "redink_sendtoword"

            Dim URL As String = jsonObject("URL")?.ToString()
            Dim Command As String = jsonObject("Command")?.ToString()
            Dim Instruction As String = jsonObject("Instruction")?.ToString()
            Dim Textbody As String = jsonObject("Text")?.ToString()

            If Command IsNot Nothing Then
                Select Case Command
                    Case "redink_sendtooutlook"

                        If Not String.IsNullOrWhiteSpace(Textbody) Then
                            Dim outlookApp As New Microsoft.Office.Interop.Outlook.Application()
                            Dim inspector As Microsoft.Office.Interop.Outlook.Inspector = outlookApp.ActiveInspector()

                            ' Ensure the inspector is open and the item is a MailItem
                            If inspector IsNot Nothing AndAlso TypeOf inspector.CurrentItem Is Microsoft.Office.Interop.Outlook.MailItem Then
                                Dim mailItem As Microsoft.Office.Interop.Outlook.MailItem = DirectCast(inspector.CurrentItem, Microsoft.Office.Interop.Outlook.MailItem)

                                ' Check if the email is in compose mode
                                If mailItem.Sent = False Then
                                    ' Get the Word editor for the email
                                    Dim wordEditor As Microsoft.Office.Interop.Word.Document = TryCast(inspector.WordEditor, Microsoft.Office.Interop.Word.Document)

                                    If wordEditor IsNot Nothing Then
                                        ' Insert the text at the current cursor position
                                        Dim selection As Microsoft.Office.Interop.Word.Selection = wordEditor.Application.Selection
                                        selection.TypeText(Textbody)
                                    End If
                                End If
                            End If
                            ' Release COM objects in reverse order of creation
                            If inspector IsNot Nothing Then Marshal.ReleaseComObject(inspector) : inspector = Nothing
                            If outlookApp IsNot Nothing Then Marshal.ReleaseComObject(outlookApp) : outlookApp = Nothing

                        End If
                        result = ""

                    Case "redink_translate"

                        If Not String.IsNullOrWhiteSpace(Textbody) Then
                            TranslateLanguage = SLib.ShowCustomInputBox("Enter your target language:", $"{AN} Translate (for Browser)", True, INI_Language1)

                            If Not String.IsNullOrWhiteSpace(TranslateLanguage) And TranslateLanguage <> "ESC" Then

                                Dim LLMResult As String = Await RunLLMOnMainThread(InterpolateAtRuntime(SP_Translate), "<TEXTTOPROCESS>" & Textbody & "</TEXTTOPROCESS>")
                                LLMResult = LLMResult.Replace("<TEXTTOPROCESS>", "").Replace("</TEXTTOPROCESS>", "")
                                If Not String.IsNullOrEmpty(LLMResult) Then
                                    result = LLMResult
                                    result = Trim(result).Replace("**", "")
                                Else
                                    result = ""
                                End If
                            Else
                                result = ""
                            End If
                        Else
                            result = ""
                        End If

                    Case "redink_correct"

                        If Not String.IsNullOrWhiteSpace(Textbody) Then

                            Dim LLMResult As String = Await RunLLMOnMainThread(InterpolateAtRuntime(SP_Correct), "<TEXTTOPROCESS>" & Textbody & "</TEXTTOPROCESS>")
                            LLMResult = LLMResult.Replace("<TEXTTOPROCESS>", "").Replace("</TEXTTOPROCESS>", "")
                            If Not String.IsNullOrEmpty(LLMResult) Then
                                result = LLMResult
                                RunCompareAndInsertTextOnMainThread(Textbody, result, True)
                                System.Windows.Forms.Application.DoEvents()
                                If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And &H8000) <> 0 Then result = ""
                                If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And 1) <> 0 Then result = ""
                            Else
                                result = ""
                            End If
                        Else
                            result = ""
                        End If

                    Case "redink_freestyle"

                        result = ""

                        Dim MarkupInstruct As String = $"start with '{MarkupPrefix}' for markups"
                        Dim InsertInstruct As String = $"with '{InsertPrefix}' for directly inserting the result"
                        Dim PromptLibInstruct As String = If(INI_PromptLib, " or press 'OK' for the prompt library", "")
                        Dim LastPromptInstruct As String = If(String.IsNullOrWhiteSpace(My.Settings.LastPrompt), "", "; ctrl-v for your last prompt")

                        Dim NoText As Boolean = False
                        Dim DoMarkup As Boolean = False
                        Dim DoInsert As Boolean = False

                        If String.IsNullOrWhiteSpace(Textbody) Then NoText = True

                        SLib.StoreClipboard()

                        If Not String.IsNullOrWhiteSpace(My.Settings.LastPrompt) Then SLib.PutInClipboard(My.Settings.LastPrompt)

                        If Not NoText Then
                            OtherPrompt = SLib.ShowCustomInputBox($"Please provide the prompt you wish to execute using the selected text ({MarkupInstruct}, {InsertInstruct}){PromptLibInstruct}{LastPromptInstruct}:", $"{AN} Freestyle (for Browser)", False)
                        Else
                            OtherPrompt = SLib.ShowCustomInputBox($"Please provide the prompt you wish to execute ({MarkupInstruct}, {InsertInstruct}){PromptLibInstruct}{LastPromptInstruct}:", $"{AN} Freestyle (for Browser)", False)
                        End If

                        SLib.RestoreClipboard()

                        If String.IsNullOrEmpty(OtherPrompt) And OtherPrompt <> "ESC" And INI_PromptLib Then

                            Dim promptlibresult As (String, Boolean, Boolean, Boolean)

                            promptlibresult = ShowPromptSelector(INI_PromptLibPath, Not NoText, Nothing)

                            OtherPrompt = promptlibresult.Item1
                            DoMarkup = promptlibresult.Item2
                            DoInsert = Not promptlibresult.Item4

                        Else
                            If String.IsNullOrEmpty(OtherPrompt) Or OtherPrompt = "ESC" Then OtherPrompt = ""
                        End If

                        If OtherPrompt <> "" Then
                            My.Settings.LastPrompt = OtherPrompt
                            My.Settings.Save()

                            If OtherPrompt.StartsWith(InsertPrefix, StringComparison.OrdinalIgnoreCase) Then
                                OtherPrompt = OtherPrompt.Substring(InsertPrefix.Length).Trim()
                                DoInsert = True
                            ElseIf OtherPrompt.StartsWith(MarkupPrefix, StringComparison.OrdinalIgnoreCase) And Not NoText Then
                                OtherPrompt = OtherPrompt.Substring(MarkupPrefix.Length).Trim()
                                DoMarkup = True
                                DoInsert = True
                            End If
                            Dim LLMResult As String = ""
                            If Not NoText Then
                                LLMResult = Await RunLLMOnMainThread(InterpolateAtRuntime(SP_FreestyleText), "<TEXTTOPROCESS>" & Textbody & "</TEXTTOPROCESS>")
                            Else
                                LLMResult = Await RunLLMOnMainThread(InterpolateAtRuntime(SP_FreestyleNoText), "")
                            End If
                            LLMResult = LLMResult.Replace("<TEXTTOPROCESS>", "").Replace("</TEXTTOPROCESS>", "")

                            If Not String.IsNullOrEmpty(LLMResult) Then
                                result = LLMResult
                                If DoMarkup Then
                                    RunCompareAndInsertTextOnMainThread(Textbody, result, True)
                                    System.Windows.Forms.Application.DoEvents()
                                    If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And &H8000) <> 0 Then result = ""
                                    If (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And 1) <> 0 Then result = ""
                                End If
                                If Not DoInsert Then
                                    Dim FinalText As String = SLib.ShowCustomWindow("The LLM has provided the following result (you can edit it):", result, "You can choose whether you want to have the original text put into the clipboard or your text with any changes you have made. If you select Cancel, nothing will be put into the clipboard (without formatting).", AN, True, True)

                                    If FinalText <> "" Then
                                        SLib.PutInClipboard(FinalText)
                                        Debug.WriteLine("Finaltext=" & FinalText)
                                    End If
                                    result = ""
                                End If
                            Else
                                result = ""
                            End If

                        End If


                End Select
            End If
        Catch ex As System.Exception
            MessageBox.Show($"Error in ProcessRequestInAddIn: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return result

    End Function

End Class
