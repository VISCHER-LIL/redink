' Red Ink for Outlook
' Copyright by David Rosenthal, david.rosenthal@vischer.com
' May only be used under the Red Ink License. See License.txt or https://vischer.com/redink for more information.
'
' 28.10.2025
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
' Includes also various Microsoft libraries copyrighted by Microsoft Corporation and available, among others, under the Microsoft EULA and the MIT License; Copyright (c) 2016- Microsoft Corp.

Option Explicit On

Imports System.Diagnostics
Imports System.Drawing
Imports System.IO
Imports System.Net
Imports System.Runtime.InteropServices
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel
Imports DiffPlex
Imports DiffPlex.DiffBuilder
Imports DiffPlex.DiffBuilder.Model
Imports Markdig
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Outlook
Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Interop.Powerpoint
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.VisualBasic.FileIO
Imports Nito.AsyncEx
Imports SharedLibrary.MarkdownToRtf
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

Public Class ThisAddIn

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function FindWindow(
                                ByVal lpClassName As String,
                                ByVal lpWindowName As String
                            ) As IntPtr
    End Function

    Public StartupInitialized As Boolean = False
    Private mainThreadControl As New System.Windows.Forms.Control()
    Private WithEvents outlookExplorer As Outlook.Explorer
    Private ReadOnly uiCtx As System.Threading.SynchronizationContext =
        System.Threading.SynchronizationContext.Current

    Private delayedStartupOnce As Integer = 0

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        Try
            RemoveHandler Microsoft.Win32.SystemEvents.PowerModeChanged, AddressOf OnPowerModeChanged
        Catch
        End Try

        StartPowerWatch()

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
        UpdateHandler.HostHandle = hwnd

        ' Other tasks that need to be done at startup
        mainThreadControl.CreateControl()

        outlookExplorer = ComRetry(Function() Application.ActiveExplorer())
        If outlookExplorer IsNot Nothing Then
            AddHandler outlookExplorer.Activate, AddressOf Explorer_Activate
        Else
            mainThreadControl.BeginInvoke(CType(AddressOf DelayedStartupTasks, MethodInvoker))
            StartupInitialized = True
        End If

        Try
            activeChatId = If(My.Settings.Inky_LastChat = 2, 2, 1)
        Catch
            activeChatId = 1
        End Try
    End Sub

    Private Sub Explorer_Activate()
        StartupInitialized = True
        RemoveHandler outlookExplorer.Activate, AddressOf Explorer_Activate
        DelayedStartupTasks()
    End Sub

    Private Sub DelayedStartupTasks()
        ' Run once even if scheduled twice (e.g., event + BeginInvoke)
        If System.Threading.Interlocked.CompareExchange(delayedStartupOnce, 1, 0) <> 0 Then Return

        Try
            InitializeConfig(True, True)
            UpdateHandler.PeriodicCheckForUpdates(INI_UpdateCheckInterval, "Outlook", INI_UpdatePath)
            Dim result = Globals.Ribbons.Ribbon1.UpdateRibbon()
            result = Globals.Ribbons.Ribbon2.UpdateRibbon()
            mainThreadControl.CreateControl()
            StartListenerWatchdog()
            StartupHttpListener()
        Catch ex As System.Exception
            ' Handling errors gracefully
        End Try
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
        ' 1) deterministically stop the HTTP listener (await synchronously)
        Try
            Dim t As System.Threading.Tasks.Task = ShutdownHttpListener()
            t.GetAwaiter().GetResult() ' safe: our shutdown continuations don’t capture the UI context
        Catch ex As System.Exception
            System.Diagnostics.Debug.WriteLine("ShutdownHttpListener failed: " & ex.Message)
        End Try

        ' 2) stop watchdog (if you added it)
        Try
            StopListenerWatchdog()
        Catch
        End Try

        ' 3) tear down power notifications window
        Try
            StopPowerWatch()
        Catch
        End Try

    End Sub




    ' Hardcoded config values

    Public Const AN As String = "Red Ink"
    Public Const AN2 As String = "red_ink"
    Public Const AN6 As String = "Inky"

    Public Const Version As String = "V.281025 Gen2 Beta Test"

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
    Private Const ClipboardPrefix2 As String = "Clip:"
    Private Const InsertPrefix As String = "Insert:"
    Private Const MyStyleTrigger As String = "(mystyle)"
    Private Const NoFormatTrigger As String = "(noformat)"
    Private Const NoFormatTrigger2 As String = "(nf)"
    Private Const KFTrigger As String = "(keepformat)"
    Private Const KFTrigger2 As String = "(kf)"
    Private Const KPFTrigger As String = "(keepparaformat)"
    Private Const KPFTrigger2 As String = "(kpf)"
    Private Const InPlacePrefix As String = "Replace:"
    Private Const NewDocPrefix As String = "Newdoc:"
    Private Const ObjectTrigger2 As String = "(clip)"

    Private Const ESC_KEY As Integer = &H1B

    Private Const SecondAPICode As String = "(2nd)"

    ' Variables that are available to InterpolateAtRuntime

    Public TranslateLanguage As String = ""
    Public OtherPrompt As String = ""
    Public Username As String = ""
    Public MyStyleInsert As String = ""
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

    Public Shared Property INI_MarkdownConvert As Boolean
        Get
            Return _context.INI_MarkdownConvert
        End Get
        Set(value As Boolean)
            _context.INI_MarkdownConvert = value
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

    Public Shared Property SP_InsertClipboard As String
        Get
            Return _context.SP_InsertClipboard
        End Get
        Set(value As String)
            _context.SP_InsertClipboard = value
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

    Public Shared Property SP_Add_Bubbles As String
        Get
            Return _context.SP_Add_Bubbles
        End Get
        Set(value As String)
            _context.SP_Add_Bubbles = value
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

    Public Shared Property SP_BubblesExcel As String
        Get
            Return _context.SP_BubblesExcel
        End Get
        Set(value As String)
            _context.SP_BubblesExcel = value
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

    Public Shared Property SP_Add_ChatWord_Commands As String
        Get
            Return _context.SP_Add_ChatWord_Commands
        End Get
        Set(value As String)
            _context.SP_Add_ChatWord_Commands = value
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

    Public Shared Property INI_MyStylePath As String
        Get
            Return _context.INI_MyStylePath
        End Get
        Set(value As String)
            _context.INI_MyStylePath = value
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


    '───────────────────────────────────────────────────────────────────────────
    ' Runs a Sub on the UI thread and waits asynchronously for it to complete.
    '───────────────────────────────────────────────────────────────────────────
    Private Function SwitchToUi(uiAction As System.Action) _
        As System.Threading.Tasks.Task

        Dim tcs As New System.Threading.Tasks.TaskCompletionSource(Of Object)()

        mainThreadControl.BeginInvoke(New MethodInvoker(
        Sub()
            Try
                uiAction.Invoke()
                tcs.SetResult(Nothing)
            Catch ex As System.Exception
                tcs.SetException(ex)
            End Try
        End Sub))

        Return tcs.Task
    End Function

    '───────────────────────────────────────────────────────────────────────────
    ' Runs a Function(Of T) on the UI thread and returns its result asynchronously.
    '───────────────────────────────────────────────────────────────────────────
    Private Function SwitchToUi(Of T)(uiFunc As System.Func(Of T)) _
        As System.Threading.Tasks.Task(Of T)

        Dim tcs As New System.Threading.Tasks.TaskCompletionSource(Of T)()

        mainThreadControl.BeginInvoke(New MethodInvoker(
        Sub()
            Try
                tcs.SetResult(uiFunc.Invoke())
            Catch ex As System.Exception
                tcs.SetException(ex)
            End Try
        End Sub))

        Return tcs.Task
    End Function

    '───────────────────────────────────────────────────────────────────────────
    ' Runs a Sub on the UI thread and *waits* for it to complete.
    '───────────────────────────────────────────────────────────────────────────
    Private Function OldSwitchToUi(uiAction As System.Action) _
        As System.Threading.Tasks.Task

        Dim tcs As New System.Threading.Tasks.TaskCompletionSource(Of Object)()

        mainThreadControl.Invoke(New MethodInvoker(
        Sub()
            Try
                uiAction.Invoke()
                tcs.SetResult(Nothing)
            Catch ex As System.Exception
                tcs.SetException(ex)
            End Try
        End Sub))

        Return tcs.Task        ' completes only after uiAction finished
    End Function

    ' OLE message filter to auto-retry transient COM call rejections
    ' OLE message filter to auto-retry transient COM call rejections (chained)
    Friend NotInheritable Class OleMessageFilter
        <System.Runtime.InteropServices.DllImport("ole32.dll")>
        Private Shared Function CoRegisterMessageFilter(newFilter As IOleMessageFilter, ByRef oldFilter As IOleMessageFilter) As Integer
        End Function

        <System.Runtime.InteropServices.ComImport(),
         System.Runtime.InteropServices.Guid("00000016-0000-0000-C000-000000000046"),
         System.Runtime.InteropServices.InterfaceType(System.Runtime.InteropServices.ComInterfaceType.InterfaceIsIUnknown)>
        Private Interface IOleMessageFilter
            <System.Runtime.InteropServices.PreserveSig()>
            Function HandleInComingCall(dwCallType As Integer,
                                        hTaskCaller As IntPtr,
                                        dwTickCount As Integer,
                                        lpInterfaceInfo As IntPtr) As Integer
            <System.Runtime.InteropServices.PreserveSig()>
            Function RetryRejectedCall(hTaskCallee As IntPtr,
                                       dwTickCount As Integer,
                                       dwRejectType As Integer) As Integer
            <System.Runtime.InteropServices.PreserveSig()>
            Function MessagePending(hTaskCallee As IntPtr,
                                    dwTickCount As Integer,
                                    dwPendingType As Integer) As Integer
        End Interface

        ' Keep a reference to the filter that Outlook installed before ours,
        ' so we can forward to it and restore it later.
        Private Shared prevFilter As IOleMessageFilter = Nothing
        Private Shared registered As Boolean

        Private Class Filter
            Implements IOleMessageFilter

            Public Function HandleInComingCall(dwCallType As Integer, hTaskCaller As IntPtr, dwTickCount As Integer, lpInterfaceInfo As IntPtr) As Integer Implements IOleMessageFilter.HandleInComingCall
                If prevFilter IsNot Nothing Then
                    Try : Return prevFilter.HandleInComingCall(dwCallType, hTaskCaller, dwTickCount, lpInterfaceInfo) : Catch : End Try
                End If
                Return 0 ' SERVERCALL_ISHANDLED
            End Function

            Public Function RetryRejectedCall(hTaskCallee As IntPtr, dwTickCount As Integer, dwRejectType As Integer) As Integer Implements IOleMessageFilter.RetryRejectedCall
                ' Ask Outlook’s filter first
                Dim prevRet As Integer = -1
                If prevFilter IsNot Nothing Then
                    Try : prevRet = prevFilter.RetryRejectedCall(hTaskCallee, dwTickCount, dwRejectType) : Catch : prevRet = -1 : End Try
                End If

                ' Only adjust RETRYLATER if Outlook would cancel (-1)
                If dwRejectType = 2 Then ' SERVERCALL_RETRYLATER
                    If prevRet >= 0 Then Return prevRet
                    Return 150 ' retry after 150ms
                End If

                ' For all other cases, preserve Outlook’s behavior
                Return prevRet
            End Function

            Public Function MessagePending(hTaskCallee As IntPtr, dwTickCount As Integer, dwPendingType As Integer) As Integer Implements IOleMessageFilter.MessagePending
                If prevFilter IsNot Nothing Then
                    Try : Return prevFilter.MessagePending(hTaskCallee, dwTickCount, dwPendingType) : Catch : End Try
                End If
                Return 2 ' PENDINGMSG_WAITDEFPROCESS
            End Function
        End Class

        Public Shared Sub Register()
            If registered Then Return
            Dim oldF As IOleMessageFilter = Nothing
            ' Register our filter and capture the previous (Outlook’s) filter
            CoRegisterMessageFilter(New Filter(), oldF)
            prevFilter = oldF
            registered = True
        End Sub

        Public Shared Sub Revoke()
            If Not registered Then Return
            Dim oldF As IOleMessageFilter = Nothing
            ' Restore Outlook’s original filter (do NOT set Nothing here)
            CoRegisterMessageFilter(prevFilter, oldF)
            prevFilter = Nothing
            registered = False
        End Sub
    End Class

    Private Sub EnableOleFilterFor(durationMs As Integer)
        ' must run on the Outlook UI thread
        Dim t As New System.Windows.Forms.Timer() With {.Interval = Math.Max(500, durationMs)}
        AddHandler t.Tick,
            Sub()
                Try : OleMessageFilter.Revoke() : Catch : End Try
                Try : t.Stop() : t.Dispose() : Catch : End Try
            End Sub
        Try : OleMessageFilter.Register() : Catch : End Try
        t.Start()
    End Sub

    Private Shared Function ComRetry(Of T)(work As System.Func(Of T)) As T
        For i As Integer = 0 To 2
            Try
                Return work()
            Catch ex As System.Runtime.InteropServices.COMException When _
            ex.HResult = &H80010001 OrElse   ' RPC_E_CALL_REJECTED
            ex.HResult = &H8001010A OrElse   ' RPC_E_SERVERCALL_RETRYLATER
            ex.HResult = &H80004005          ' E_FAIL (some busy states)
                ' Avoid Application.DoEvents here to prevent re-entrancy into COM/Ribbon
                System.Threading.Thread.Sleep(150)
            End Try
        Next
        Return work() ' last try to surface real error
    End Function

    '───────────────────────────────────────────────────────────────────────────
    ' Runs a Function(Of T) on the UI thread and waits for its return value.
    '───────────────────────────────────────────────────────────────────────────
    Private Function OldSwitchToUi(Of T)(uiFunc As System.Func(Of T)) _
        As System.Threading.Tasks.Task(Of T)

        Dim tcs As New System.Threading.Tasks.TaskCompletionSource(Of T)()

        mainThreadControl.Invoke(New MethodInvoker(
        Sub()
            Try
                tcs.SetResult(uiFunc.Invoke())
            Catch ex As System.Exception
                tcs.SetException(ex)
            End Try
        End Sub))

        Return tcs.Task        ' completes only after uiFunc returns
    End Function


    '───────────────────────────────────────────────────────────────────────────
    ' SwitchToUiTask  –  runs an *async* function (returns Task(Of T)) on the
    ' Outlook UI thread and gives you a Task(Of T) you can Await from any thread.
    '───────────────────────────────────────────────────────────────────────────
    Private Function SwitchToUiTask(Of T)(
        uiFunc As System.Func(Of System.Threading.Tasks.Task(Of T))) _
        As System.Threading.Tasks.Task(Of T)

        Dim tcs As New System.Threading.Tasks.TaskCompletionSource(Of T)()

        mainThreadControl.BeginInvoke(New MethodInvoker(
        Sub()
            Try
                Dim inner As System.Threading.Tasks.Task(Of T) = uiFunc.Invoke()
                inner.ContinueWith(
                    Sub(taskObj)
                        If taskObj.IsFaulted Then
                            tcs.SetException(taskObj.Exception.InnerExceptions)
                        ElseIf taskObj.IsCanceled Then
                            tcs.SetCanceled()
                        Else
                            tcs.SetResult(taskObj.Result)
                        End If
                    End Sub,
                    System.Threading.Tasks.TaskScheduler.Default)
            Catch ex As System.Exception
                tcs.SetException(ex)
            End Try
        End Sub))

        Return tcs.Task
    End Function


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

    Public Shared Async Function LLM(ByVal promptSystem As String, ByVal promptUser As String, Optional ByVal Model As String = "", Optional ByVal Temperature As String = "", Optional ByVal Timeout As Long = 0, Optional ByVal UseSecondAPI As Boolean = False, Optional HideSplash As Boolean = False, Optional ByVal AddUserPrompt As String = "", Optional ByVal FileObject As String = "", Optional cancellationToken As Threading.CancellationToken = Nothing) As Task(Of String)
        Return Await SharedMethods.LLM(_context, promptSystem, promptUser, Model, Temperature, Timeout, UseSecondAPI, HideSplash, AddUserPrompt, FileObject, cancellationToken)
    End Function

    Private Function ShowSettingsWindow(Settings As Dictionary(Of String, String), SettingsTips As Dictionary(Of String, String))
        If Not INIloaded Then
            If Not StartupInitialized Then
                Try
                    DelayedStartupTasks()
                    RemoveHandler outlookExplorer.Activate, AddressOf Explorer_Activate
                Catch ex As System.Exception
                End Try
                If Not INIloaded Then Return Nothing
            Else
                InitializeConfig(False, False)
                If Not INIloaded Then
                    Return Nothing
                End If
            End If
        End If
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

    Friend NotInheritable Class WordUndoScope
        Implements System.IDisposable

        Private ReadOnly _app As Microsoft.Office.Interop.Word.Application
        Private ReadOnly _undo As Microsoft.Office.Interop.Word.UndoRecord
        Private ReadOnly _iStarted As System.Boolean

        Public Sub New(app As Microsoft.Office.Interop.Word.Application, Optional name As System.String = Nothing)
            _app = app
            _undo = _app.UndoRecord

            ' Word < 2013 (Version < 15.0) hat kein UndoRecord.
            Dim ver As System.Version = New System.Version(_app.Version)
            If ver.Major < 15 Then
                Return
            End If

            ' Nur starten, wenn gerade kein anderer Custom-Record läuft
            If Not _undo.IsRecordingCustomRecord Then
                If name IsNot Nothing AndAlso name.Length > 0 Then
                    _undo.StartCustomRecord(name)
                Else
                    _undo.StartCustomRecord("VSTO-Aktion")
                End If
                _iStarted = True
            End If
        End Sub

        Public Sub Dispose() Implements System.IDisposable.Dispose
            Try
                If _iStarted AndAlso _undo.IsRecordingCustomRecord Then
                    _undo.EndCustomRecord()
                End If
            Catch ex As System.Exception
                ' Nichts werfen – wir sind in Dispose
            End Try
        End Sub
    End Class

    Public Sub MainMenu(RI_Command As String)

        ' Acquire single-entry guard; if already in MainMenu, bail out
        If System.Threading.Interlocked.CompareExchange(inMainMenu, 1, 0) <> 0 Then Return

        Try
            If IsInResumeCooldown() Then
                SLib.ShowCustomMessageBox("Outlook is resuming from sleep. Please try again in a few seconds.")
                Return
            End If

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

            InitializeConfig(False, False)

            If GPTSetupError OrElse INIValuesMissing() Or Not INIloaded Then Return

            ' Use fully qualified names to avoid ambiguity
            Dim outlookApp As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application
            Dim inspector As Microsoft.Office.Interop.Outlook.Inspector = ComRetry(Function() outlookApp.ActiveInspector())

            Dim Textlength As Long

            If inspector Is Nothing Then

                InspectorOpened = False

                OpenInspectorAndReapplySelection(RI_Command)

                If Not InspectorOpened Then Exit Sub

                inspector = ComRetry(Function() outlookApp.ActiveInspector())
                If inspector Is Nothing Then
                    System.Windows.Forms.MessageBox.Show("Error in MainMenu: No active email item found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If
            End If

            Dim curr As Object = ComRetry(Function() inspector.CurrentItem)
            If curr Is Nothing OrElse Not TypeOf curr Is Microsoft.Office.Interop.Outlook.MailItem Then
                SLib.ShowCustomMessageBox($"Please open an email for editing for using {AN}.")
                Return
            End If

            Dim mailItem As Microsoft.Office.Interop.Outlook.MailItem = CType(curr, Microsoft.Office.Interop.Outlook.MailItem)
            Dim wordEditor As Microsoft.Office.Interop.Word.Document = ComRetry(Function() CType(inspector.WordEditor, Microsoft.Office.Interop.Word.Document))

            Select Case RI_Command

                Case "Translate"
                    TranslateLanguage = ""
                    TranslateLanguage = SLib.ShowCustomInputBox("Enter your target language:", $"{AN} Translate", True, INI_Language2)
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
                        UserInput = Trim(SLib.ShowCustomInputBox("Enter the number of words your summary shall have (the selected text has " & Textlength & " words; the proposal " & SummaryPercent & "%):", $"{AN} Summarizer", True, CStr(Math.Round(SummaryPercent * Textlength / 100 / 5) * 5)))

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
                Case "ApplyMyStyle"
                    Dim StylePath As String = ExpandEnvironmentVariables(INI_MyStylePath)

                    If String.IsNullOrWhiteSpace(StylePath) Then
                        ShowCustomMessageBox("You have not defined a MyStyle prompt file. Please do so first in the configuration file or using 'Settings'.")
                        Return
                    End If
                    If Not IO.File.Exists(StylePath) Then
                        ShowCustomMessageBox("No MyStyle prompt file has been found. You may have to first create a MyStyle prompt. Go to 'Analyze' and use 'Define MyStyle' to do so - exiting.")
                        Return
                    End If

                    Textlength = GetSelectedTextLength()
                    If Textlength = 0 Then
                        SLib.ShowCustomMessageBox("Please select the text to be processed.")
                        Return
                    End If

                    MyStyleInsert = MyStyleHelpers.SelectPromptFromMyStyle(StylePath, "Outlook", 0, "Choose the style prompt to apply …", $"{AN} MyStyle", False)
                    If MyStyleInsert = "ERROR" Then Return
                    If MyStyleInsert = "NONE" OrElse String.IsNullOrWhiteSpace(MyStyleInsert) Then
                        Return
                    End If

                    Command_InsertAfter(InterpolateAtRuntime(SP_MyStyle_Apply) & " " & MyStyleInsert, INI_DoMarkupOutlook, INI_KeepFormat2, INI_ReplaceText2, INI_MarkupMethodOutlook)

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

                    Dim selectedText As String = mailItem.Body
                    ShowSumup(selectedText)

                'FreeStyle_InsertBefore(SP_MailSumup, False)
                Case "Answers"
                    FreeStyle_InsertBefore(SP_MailReply, True)
                Case "Freestyle"
                    FreeStyle_InsertAfter()
                Case "InsertClipboard"
                    InsertClipboard()
                Case Else
                    System.Windows.Forms.MessageBox.Show("Error in MainMenu: Invalid internal command.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Select

            If inspector IsNot Nothing Then Marshal.ReleaseComObject(inspector) : inspector = Nothing
            'If outlookApp IsNot Nothing Then Marshal.ReleaseComObject(outlookApp) : outlookApp = Nothing

        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show("Error in MainMenu: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ' Always release the reentrancy guard so subsequent calls work
            System.Threading.Interlocked.Exchange(inMainMenu, 0)
        End Try
    End Sub

    Public Sub OldMainMenu(RI_Command As String)

        If System.Threading.Interlocked.Exchange(inMainMenu, 1) = 1 Then Return

        If IsInResumeCooldown() Then
            SLib.ShowCustomMessageBox("Outlook is resuming from sleep. Please try again in a few seconds.")
            Return
        End If

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

            InitializeConfig(False, False)

            If GPTSetupError OrElse INIValuesMissing() Or Not INIloaded Then Return

            ' Use fully qualified names to avoid ambiguity
            Dim outlookApp As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application
            Dim inspector As Microsoft.Office.Interop.Outlook.Inspector = ComRetry(Function() outlookApp.ActiveInspector())

            Dim Textlength As Long

            If inspector Is Nothing Then

                InspectorOpened = False

                OpenInspectorAndReapplySelection(RI_Command)

                If Not InspectorOpened Then Exit Sub

                inspector = ComRetry(Function() outlookApp.ActiveInspector())
                If inspector Is Nothing Then

                    System.Windows.Forms.MessageBox.Show("Error in MainMenu: No active email item found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If
            End If


            Dim curr As Object = ComRetry(Function() inspector.CurrentItem)
            If curr Is Nothing OrElse Not TypeOf curr Is Microsoft.Office.Interop.Outlook.MailItem Then
                SLib.ShowCustomMessageBox($"Please open an email for editing for using {AN}.")
                Return
            End If

            Dim mailItem As Microsoft.Office.Interop.Outlook.MailItem = CType(curr, Microsoft.Office.Interop.Outlook.MailItem)
            Dim wordEditor As Microsoft.Office.Interop.Word.Document = ComRetry(Function() CType(inspector.WordEditor, Microsoft.Office.Interop.Word.Document))

            Select Case RI_Command

                Case "Translate"
                    TranslateLanguage = ""
                    TranslateLanguage = SLib.ShowCustomInputBox("Enter your target language:", $"{AN} Translate", True, INI_Language2)
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
                        UserInput = Trim(SLib.ShowCustomInputBox("Enter the number of words your summary shall have (the selected text has " & Textlength & " words; the proposal " & SummaryPercent & "%):", $"{AN} Summarizer", True, CStr(Math.Round(SummaryPercent * Textlength / 100 / 5) * 5)))

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
                Case "ApplyMyStyle"
                    Dim StylePath As String = ExpandEnvironmentVariables(INI_MyStylePath)

                    If String.IsNullOrWhiteSpace(StylePath) Then
                        ShowCustomMessageBox("You have not defined a MyStyle prompt file. Please do so first in the configuration file or using 'Settings'.")
                        Return
                    End If
                    If Not IO.File.Exists(StylePath) Then
                        ShowCustomMessageBox("No MyStyle prompt file has been found. You may have to first create a MyStyle prompt. Go to 'Analyze' and use 'Define MyStyle' to do so - exiting.")
                        Return
                    End If

                    Textlength = GetSelectedTextLength()
                    If Textlength = 0 Then
                        SLib.ShowCustomMessageBox("Please select the text to be processed.")
                        Return
                    End If

                    MyStyleInsert = MyStyleHelpers.SelectPromptFromMyStyle(StylePath, "Outlook", 0, "Choose the style prompt to apply …", $"{AN} MyStyle", False)
                    If MyStyleInsert = "ERROR" Then Return
                    If MyStyleInsert = "NONE" OrElse String.IsNullOrWhiteSpace(MyStyleInsert) Then
                        Return
                    End If

                    Command_InsertAfter(InterpolateAtRuntime(SP_MyStyle_Apply) & " " & MyStyleInsert, INI_DoMarkupOutlook, INI_KeepFormat2, INI_ReplaceText2, INI_MarkupMethodOutlook)

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

                    Dim selectedText As String = mailItem.Body
                    ShowSumup(selectedText)

                        'FreeStyle_InsertBefore(SP_MailSumup, False)
                Case "Answers"
                    FreeStyle_InsertBefore(SP_MailReply, True)
                Case "Freestyle"
                    FreeStyle_InsertAfter()
                Case "InsertClipboard"
                    InsertClipboard()
                Case Else
                    System.Windows.Forms.MessageBox.Show("Error in MainMenu: Invalid internal command.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Select

            If inspector IsNot Nothing Then Marshal.ReleaseComObject(inspector) : inspector = Nothing
            'If outlookApp IsNot Nothing Then Marshal.ReleaseComObject(outlookApp) : outlookApp = Nothing
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show("Error in MainMenu: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Function GetActiveInspector() As Outlook.Inspector
        Try
            Dim activeWindow = Globals.ThisAddIn.Application.ActiveWindow()
            If activeWindow IsNot Nothing AndAlso TypeOf activeWindow Is Outlook.Inspector Then
                ' The active window is an inspector, return it.
                Return CType(activeWindow, Outlook.Inspector)
            End If

            ' If the active window is not an inspector (e.g., it's the Explorer),
            ' or if there's no active window, return Nothing.
            If activeWindow IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(activeWindow)
            End If
            Return Nothing
        Catch
            Return Nothing
        End Try
    End Function


    Public Sub OpenInspectorAndReapplySelection(Command As String)
        Try

            If Command = "InsertClipboard" Then InsertClipboard() : Return

            Dim Sumup As Boolean = (Command = "Sumup")
            Dim Translate As Boolean = (Command = "Translate" OrElse Command = "PrimLang")

            ' Grab Outlook instances
            Dim oApp As Outlook.Application = Globals.ThisAddIn.Application
            Dim oExplorer As Outlook.Explorer = ComRetry(Function() oApp.ActiveExplorer())

            ' Check for inline response
            Dim inlineResponse As Object = If(oExplorer Is Nothing, Nothing,
                                              ComRetry(Function() oExplorer.ActiveInlineResponse))

            If inlineResponse Is Nothing OrElse Sumup OrElse Translate Then
                ' Get the current selection in the explorer
                Dim selection As Outlook.Selection = If(oExplorer Is Nothing, Nothing,
                                                       ComRetry(Function() oExplorer.Selection))
                Dim selectionCount As Integer = If(selection Is Nothing, 0, ComRetry(Function() selection.Count))

                If selectionCount = 0 Then
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

                        Const PR_LAST_VERB_EXECUTED As String = "http://schemas.microsoft.com/mapi/proptag/0x10810003"

                        Dim selectedText As String = String.Empty
                        Dim count As Integer = 1
                        For Each mail As Microsoft.Office.Interop.Outlook.MailItem In mailItems

                            Dim lastVerb As Integer = 0

                            Try
                                lastVerb = mail.PropertyAccessor.GetProperty(PR_LAST_VERB_EXECUTED)
                            Catch comEx As COMException
                                ' Property nicht gesetzt → noch nicht beantwortet
                                lastVerb = 0
                            Catch ex As System.Exception
                                ' Sicherstellen, dass System.Exception voll qualifiziert ist
                                lastVerb = 0
                            End Try


                            If lastVerb <> 102 AndAlso lastVerb <> 103 Then
                                Dim tag As String = count.ToString("D4") ' Format count with four digits
                                Dim latestBody As String = GetLatestMailBody(mail.Body)
                                selectedText &= "<EMAIL" & tag & ">" & latestBody & "</EMAIL" & tag & ">"
                                count += 1
                            End If
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
                    ElseIf Translate Then
                        Dim selectedItem As Object = selection(1)
                        If TypeOf selectedItem Is Outlook.MailItem Then

                            If Command = "Translate" Then
                                TranslateLanguage = ""
                                TranslateLanguage = SLib.ShowCustomInputBox("Enter your target language:", $"{AN} Translate", True, INI_Language2)
                                If String.IsNullOrEmpty(TranslateLanguage) Then Return
                            Else
                                TranslateLanguage = INI_Language1
                            End If

                            Dim mail As Outlook.MailItem = CType(selectedItem, Outlook.MailItem)
                            Dim selectedText As String = mail.Body

                            ShowTranslate(selectedText)
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

            ' Ensure it's still open and usable (guard COM with retries)
            If inspector Is Nothing Then
                inspector = ComRetry(Function() Globals.ThisAddIn.Application.ActiveInspector())
                If inspector Is Nothing Then
                    MessageBox.Show("Error in OpenInspectorAndReapplySelection: No active Inspector available.",
                                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If
            End If

            Dim curr As Object = ComRetry(Function() inspector.CurrentItem)
            If curr Is Nothing OrElse Not TypeOf curr Is Microsoft.Office.Interop.Outlook.MailItem Then
                MessageBox.Show("Error in OpenInspectorAndReapplySelection: The Inspector is not ready or no email item is active.",
                                "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            ' Reapply the original selection (or caret position) to the new Inspector's WordEditor
            Try
                Dim wordDoc As Word.Document = ComRetry(Function() TryCast(inspector.WordEditor, Word.Document))
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

    Public Sub oldOpenInspectorAndReapplySelection(Command As String)
        Try

            If Command = "InsertClipboard" Then InsertClipboard() : Return

            ' Grab Outlook instances
            Dim oApp As Outlook.Application = Globals.ThisAddIn.Application
            Dim oExplorer As Outlook.Explorer = oApp.ActiveExplorer()

            Dim Sumup As Boolean = (Command = "Sumup")
            Dim Translate As Boolean = (Command = "Translate" OrElse Command = "PrimLang")

            If oExplorer Is Nothing Then
                If Sumup Or Translate Then
                    ShowCustomMessageBox("You can only use this function when you have selected an e-mail.")
                Else
                    ShowCustomMessageBox("You can only use this function when you are editing an e-mail.")
                End If
                Return
            End If

            ' Check for inline response
            Dim inlineResponse As Object = oExplorer.ActiveInlineResponse
            'If inlineResponse Is Nothing Then
            If inlineResponse Is Nothing OrElse Sumup OrElse Translate Then


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

                        Const PR_LAST_VERB_EXECUTED As String = "http://schemas.microsoft.com/mapi/proptag/0x10810003"

                        Dim selectedText As String = String.Empty
                        Dim count As Integer = 1
                        For Each mail As Microsoft.Office.Interop.Outlook.MailItem In mailItems

                            Dim lastVerb As Integer = 0

                            Try
                                lastVerb = mail.PropertyAccessor.GetProperty(PR_LAST_VERB_EXECUTED)
                            Catch comEx As COMException
                                ' Property nicht gesetzt → noch nicht beantwortet
                                lastVerb = 0
                            Catch ex As System.Exception
                                ' Sicherstellen, dass System.Exception voll qualifiziert ist
                                lastVerb = 0
                            End Try


                            If lastVerb <> 102 AndAlso lastVerb <> 103 Then
                                Dim tag As String = count.ToString("D4") ' Format count with four digits
                                Dim latestBody As String = GetLatestMailBody(mail.Body)
                                selectedText &= "<EMAIL" & tag & ">" & latestBody & "</EMAIL" & tag & ">"
                                count += 1
                            End If
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
                    ElseIf Translate Then
                        Dim selectedItem As Object = selection(1)
                        If TypeOf selectedItem Is Outlook.MailItem Then

                            If Command = "Translate" Then
                                TranslateLanguage = ""
                                TranslateLanguage = SLib.ShowCustomInputBox("Enter your target language:", $"{AN} Translate", True, INI_Language2)
                                If String.IsNullOrEmpty(TranslateLanguage) Then Return
                            Else
                                TranslateLanguage = INI_Language1
                            End If

                            Dim mail As Outlook.MailItem = CType(selectedItem, Outlook.MailItem)
                            Dim selectedText As String = mail.Body

                            ShowTranslate(selectedText)
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

            ' Ensure it's still open and usable (guard COM with retries)
            If inspector Is Nothing Then
                inspector = ComRetry(Function() Globals.ThisAddIn.Application.ActiveInspector())
                If inspector Is Nothing Then
                    MessageBox.Show("Error in OpenInspectorAndReapplySelection: No active Inspector available.",
                                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If
            End If

            Dim curr As Object = ComRetry(Function() inspector.CurrentItem)
            If curr Is Nothing OrElse Not TypeOf curr Is Microsoft.Office.Interop.Outlook.MailItem Then
                MessageBox.Show("Error in OpenInspectorAndReapplySelection: The Inspector is not ready or no email item is active.",
                                "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            ' Reapply the original selection (or caret position) to the new Inspector's WordEditor
            Try
                Dim wordDoc As Word.Document = ComRetry(Function() TryCast(inspector.WordEditor, Word.Document))
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
            Dim inlineWordEditor As Object =
                If(oExplorer Is Nothing, Nothing,
                   ComRetry(Function() oExplorer.ActiveInlineResponseWordEditor))
            If inlineWordEditor Is Nothing Then
                Return False
            End If

            Dim wordSel As Word.Selection =
                ComRetry(Function() TryCast(inlineWordEditor.Application.Selection, Word.Selection))
            If wordSel Is Nothing Then
                Return False
            End If

            selStart = wordSel.Start
            selEnd = wordSel.End
            Return True
        Catch ex As System.Exception
            MessageBox.Show("Failed to retrieve the selection: " & ex.Message,
                            "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    Private Function oldGetSelectionOrCaretRangeFromInlineEditor(oExplorer As Outlook.Explorer, ByRef selStart As Integer, ByRef selEnd As Integer) As Boolean
        Try
            Dim inlineWordEditor As Object = oExplorer.ActiveInlineResponseWordEditor
            If inlineWordEditor Is Nothing Then
                ' No inline Word editor, so we can't read a selection/caret
                Return False
            End If

            Dim wordSel As Word.Selection =
            ComRetry(Function() TryCast(inlineWordEditor.Application.Selection, Word.Selection))
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

        'Dim markdownPipeline As MarkdownPipeline = New MarkdownPipelineBuilder().Build()

        Dim builder As New MarkdownPipelineBuilder()

        builder.UsePipeTables()
        builder.UseGridTables()
        builder.UseSoftlineBreakAsHardlineBreak()
        builder.UseListExtras()
        builder.UseFootnotes()
        builder.UseDefinitionLists()
        builder.UseAbbreviations()
        builder.UseAutoLinks()
        builder.UseTaskLists()
        builder.UseEmojiAndSmiley()
        builder.UseMathematics()
        builder.UseFigures()
        builder.UseAdvancedExtensions()
        builder.UseGenericAttributes()

        Dim markdownPipeline As MarkdownPipeline = builder.Build()

        Dim htmlText As String = Markdown.ToHtml(LLMResult, markdownPipeline)

        Dim fullHtml As String =
              "<!DOCTYPE html>" &
              "<html><head>" &
              "  <meta charset=""utf-8"" />" &
              "  <style>" &
              "    ul { margin-left: 0.5em; padding-left: 0; list-style-position: outside; }" &
              "    ul ul { margin-left: 1em; padding-left: 0; list-style-type: circle; }" &
              "    ul ul ul { margin-left: 1.5em; padding-left: 0; list-style-type: square; }" &
              "  </style>" &
              "</head><body>" &
                htmlText &
              "</body></html>"

        ShowHTMLCustomMessageBox(fullHtml, $"{AN} Sum-up")

    End Sub

    Private Async Sub ShowTranslate(selectedtext As String)

        Dim LLMResult As String = ""

        LLMResult = Await LLM(InterpolateAtRuntime(SP_Translate), "<TEXTTOPROCESS>" & selectedtext & "</TEXTTOPROCESS>", "", "", 0)

        If INI_PostCorrection <> "" Then
            LLMResult = Await PostCorrection(LLMResult)
        End If

        'Dim markdownPipeline As MarkdownPipeline = New MarkdownPipelineBuilder().Build()

        Dim builder As New MarkdownPipelineBuilder()

        builder.UsePipeTables()
        builder.UseGridTables()
        builder.UseSoftlineBreakAsHardlineBreak()
        builder.UseListExtras()
        builder.UseFootnotes()
        builder.UseDefinitionLists()
        builder.UseAbbreviations()
        builder.UseAutoLinks()
        builder.UseTaskLists()
        builder.UseEmojiAndSmiley()
        builder.UseMathematics()
        builder.UseFigures()
        builder.UseAdvancedExtensions()
        builder.UseGenericAttributes()

        Dim markdownPipeline As MarkdownPipeline = builder.Build()

        Dim htmlText As String = Markdown.ToHtml(LLMResult, markdownPipeline)

        ShowHTMLCustomMessageBox(htmlText, $"{AN} Translation")

    End Sub

    Private Async Sub ShowSumup2(selectedtext As String)

        Dim LLMResult As String = ""

        DateTimeNow = DateTime.Now.ToString("yyyy-MMM-dd HH:mm")

        LLMResult = Await LLM(InterpolateAtRuntime(SP_MailSumup2), selectedtext, "", "", 0)

        If INI_PostCorrection <> "" Then
            LLMResult = Await PostCorrection(LLMResult)
        End If

        ' Dim markdownPipeline As MarkdownPipeline = New MarkdownPipelineBuilder().Build()

        Dim builder As New MarkdownPipelineBuilder()

        builder.UsePipeTables()
        builder.UseGridTables()
        builder.UseSoftlineBreakAsHardlineBreak()
        builder.UseListExtras()
        builder.UseFootnotes()
        builder.UseDefinitionLists()
        builder.UseAbbreviations()
        builder.UseAutoLinks()
        builder.UseTaskLists()
        builder.UseEmojiAndSmiley()
        builder.UseMathematics()
        builder.UseFigures()
        builder.UseAdvancedExtensions()
        builder.UseGenericAttributes()

        Dim markdownPipeline As MarkdownPipeline = builder.Build()

        Dim htmlText As String = Markdown.ToHtml(LLMResult, markdownPipeline)

        Dim fullHtml As String =
              "<!DOCTYPE html>" &
              "<html><head>" &
              "  <meta charset=""utf-8"" />" &
              "  <style>" &
              "    ul { margin-left: 0.5em; padding-left: 0; list-style-position: outside; }" &
              "    ul ul { margin-left: 1em; padding-left: 0; list-style-type: circle; }" &
              "    ul ul ul { margin-left: 1.5em; padding-left: 0; list-style-type: square; }" &
              "  </style>" &
              "</head><body>" &
                htmlText &
              "</body></html>"

        ShowHTMLCustomMessageBox(fullHtml, $"{AN} Sum-up (of unanswered mails)")

    End Sub


    ' Helper: robustly set clipboard with retries to avoid "clipboard locked" errors.
    Private Sub SafeSetClipboard(dataObj As System.Windows.Forms.DataObject)
        Const maxAttempts As Integer = 8
        For attempt As Integer = 1 To maxAttempts
            Try
                System.Windows.Forms.Clipboard.SetDataObject(dataObj, True)
                Return
            Catch ex As System.Runtime.InteropServices.ExternalException
                ' Clipboard likely locked by another process, retry
                System.Threading.Thread.Sleep(40 * attempt)
            Catch ex As System.Exception
                ' Non‑transient; bail out
                SLib.ShowCustomMessageBox($"Clipboard copy failed: {ex.Message}")
                Return
            End Try
        Next
        SLib.ShowCustomMessageBox("Could not access the clipboard after several retries (another application may be holding it).")
    End Sub

    Private Async Function InsertClipboard() As System.Threading.Tasks.Task
        Try
            ' 1) Configuration check (original behavior)
            If System.String.IsNullOrWhiteSpace(INI_APICall_Object) Then
                SLib.ShowCustomMessageBox($"Your model ({INI_Model}) is not configured to process clipboard data (binary/object).")
                Return
            End If

            ' 2) Call LLM (may attempt to read the current clipboard object internally)
            Dim result As String = Await LLM(
                InterpolateAtRuntime(SP_InsertClipboard),
                "", "", "", 0, False, False, "", "clipboard"
            ).ConfigureAwait(False)

            If String.IsNullOrWhiteSpace(result) Then Return

            ' 3) Determine whether we have an open mail inspector + mail item
            Dim outlookApp As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application
            Dim inspector As Microsoft.Office.Interop.Outlook.Inspector = ComRetry(Function() outlookApp.ActiveInspector())
            Dim curr As Object = Nothing
            If inspector IsNot Nothing Then
                Try
                    curr = ComRetry(Function() inspector.CurrentItem)
                Catch
                    curr = Nothing
                End Try
            End If

            Dim haveMail As Boolean =
                inspector IsNot Nothing AndAlso
                curr IsNot Nothing AndAlso
                TypeOf curr Is Microsoft.Office.Interop.Outlook.MailItem

            If Not haveMail Then
                ' 4a) No mail open (Explorer context): put result onto clipboard
                Dim displayText As String = If(result.Length > 11000, result.Substring(0, 11000) & "…", result)

                Await SwitchToUi(
                    Sub()
                        Dim rtfText As String = Nothing
                        Dim dataObj As New System.Windows.Forms.DataObject()

                        ' RTF conversion guarded
                        Try
                            rtfText = MarkdownToRtfConverter.Convert(result)
                        Catch ex As System.Exception
                            ' Ignore; fallback to plain text only
                            rtfText = Nothing
                        End Try

                        If Not String.IsNullOrEmpty(rtfText) Then
                            dataObj.SetData(System.Windows.Forms.DataFormats.Rtf, rtfText)
                        End If
                        dataObj.SetData(System.Windows.Forms.DataFormats.Text, result)

                        SafeSetClipboard(dataObj)
                        SLib.ShowCustomMessageBox(
                            $"The content has been copied to the clipboard:{Environment.NewLine}{Environment.NewLine}{displayText}"
                        )
                    End Sub
                ).ConfigureAwait(False)

                Return
            End If

            ' 4b) Mail is open – insert at cursor
            Dim wordEditor As Microsoft.Office.Interop.Word.Document =
                ComRetry(Function() CType(inspector.WordEditor, Microsoft.Office.Interop.Word.Document))
            If wordEditor Is Nothing Then
                ' Fallback to clipboard if Word editor not available
                Await SwitchToUi(
                    Sub()
                        Dim dobj As New System.Windows.Forms.DataObject()
                        dobj.SetData(System.Windows.Forms.DataFormats.Text, result)
                        SafeSetClipboard(dobj)
                        SLib.ShowCustomMessageBox("Could not access the mail editor; result copied to clipboard instead.")
                    End Sub
                )
                Return
            End If

            Dim selection As Microsoft.Office.Interop.Word.Selection = wordEditor.Application.Selection
            If selection IsNot Nothing Then
                If selection.Start <> selection.End Then
                    selection.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
                End If
                selection.TypeParagraph()
                InsertTextWithMarkdown(selection, result, True)
            End If

            ' Release COM objects explicitly (only those we created here)
            If selection IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(selection) : selection = Nothing
            End If
            If wordEditor IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wordEditor) : wordEditor = Nothing
            End If
            If inspector IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(inspector) : inspector = Nothing
            End If

        Catch ex As System.Runtime.InteropServices.ExternalException
            ' Likely a clipboard contention scenario
            SLib.ShowCustomMessageBox($"InsertClipboard clipboard error (probably locked): {ex.Message}")
        Catch ex As System.Exception
            SLib.ShowCustomMessageBox($"InsertClipboard failed: {ex.GetType().FullName}: {ex.Message}")
        End Try
    End Function

    Private Async Function OldInsertClipboard() As System.Threading.Tasks.Task
        Try
            ' 1) Configure check
            If System.String.IsNullOrWhiteSpace(INI_APICall_Object) Then
                SLib.ShowCustomMessageBox($"Your model ({INI_Model}) is not configured to process clipboard data (i.e. binary objects).")
                Return
            End If

            ' 2) Call LLM
            Dim result As System.String = Await LLM(
            InterpolateAtRuntime(SP_InsertClipboard),
            "", "", "", 0, False, False, "", "clipboard"
        ).ConfigureAwait(False)

            If System.String.IsNullOrEmpty(result) Then Return

            ' 3) Check for open MailItem (prefer the running instance)
            Dim outlookApp As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application
            Dim inspector As Microsoft.Office.Interop.Outlook.Inspector = ComRetry(Function() outlookApp.ActiveInspector())

            ' Guard Inspector and CurrentItem via ComRetry
            Dim curr As Object = Nothing
            If inspector IsNot Nothing Then
                Try
                    curr = ComRetry(Function() inspector.CurrentItem)
                Catch
                    curr = Nothing
                End Try
            End If

            If inspector Is Nothing _
               OrElse curr Is Nothing _
               OrElse Not TypeOf curr Is Microsoft.Office.Interop.Outlook.MailItem Then

                ' No open email: copy to clipboard (cut to 6000 chars + ellipsis)
                Dim displayText As System.String = If(result.Length > 6000, result.Substring(0, 6000) & "…", result)

                ' Ensure this runs on the Outlook UI STA thread:
                Await SwitchToUi(
                    Sub()
                        Dim rtfText As System.String = MarkdownToRtfConverter.Convert(result)
                        Dim dataObj As New System.Windows.Forms.DataObject()
                        dataObj.SetData(System.Windows.Forms.DataFormats.Rtf, rtfText)
                        dataObj.SetData(System.Windows.Forms.DataFormats.Text, result)
                        System.Windows.Forms.Clipboard.SetDataObject(dataObj, True)
                        SLib.ShowCustomMessageBox($"The content has been copied to the clipboard:{System.Environment.NewLine}{System.Environment.NewLine}{displayText}")
                    End Sub
                ).ConfigureAwait(True)

                Return
            End If

            ' 4) Insert into the current email at the cursor
            Dim wordEditor As Microsoft.Office.Interop.Word.Document =
            ComRetry(Function() CType(inspector.WordEditor, Microsoft.Office.Interop.Word.Document))
            Dim selection As Microsoft.Office.Interop.Word.Selection = wordEditor.Application.Selection

            If selection.Start <> selection.End Then
                selection.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
            End If

            selection.TypeParagraph()
            InsertTextWithMarkdown(selection, result, True)

        Catch ex As System.Exception
            ' Log and show a friendly message instead of crashing Outlook
            SLib.ShowCustomMessageBox($"InsertClipboard failed: {ex.GetType().FullName}: {ex.Message}")
        End Try
    End Function



    Public Async Sub DefineMyStyle()

        Try
            ' --- Check MyStyle path (like Word) ---
            Dim stylePath As System.String = System.Environment.ExpandEnvironmentVariables(INI_MyStylePath)

            If System.String.IsNullOrWhiteSpace(stylePath) Then
                ShowCustomMessageBox("You have not configured a MyStyle prompt file path. Please do so in the configuration file or using 'Settings'.")
                Return
            End If

            ' --- Intro info box (adapted to Outlook workflow) ---
            Dim introLabel As System.String =
                $"You are about to have {AN} create a profile of your writing style from selected emails. There are six steps:" & vbCrLf & vbCrLf &
                "1. You will enter your name (used by the prompt to detect your mails)." & vbCrLf &
                "2. All currently open emails (including those opened from .MSG files) will be gathered as samples." & vbCrLf &
                "3. You can provide additional instructions (e.g., links or aspects to focus on)." & vbCrLf &
                "4. You select the model to perform the analysis (e.g., a reasoning model, Internet access if links are to be consulted)." & vbCrLf &
                "5. You can review and amend the analysis, including the final prompt for the AI to implement your style." & vbCrLf &
                $"6. The analysis will be saved to your personal MyStyle prompt file ({stylePath})."

            Dim answer As System.Int32 = ShowCustomYesNoBox(introLabel, "Continue", "Cancel", $"{AN} Define MyStyle (Outlook)",
                                                            extraButtonText:="Edit MyStyle prompt file",
                                                            extraButtonAction:=Sub()
                                                                                   SLib.ShowTextFileEditor(stylePath, "Edit your MyStyle prompt file (use 'Define MyStyle' to create new prompts automatically):")
                                                                               End Sub)
            If answer <> 1 Then
                Return
            End If

            ' --- Ask for Username (default = OS user) ---
            Dim defaultUser As System.String = System.Environment.UserName
            Username = SLib.ShowCustomInputBox("Please enter your name (will be used to identify your mails within mailchains):", $"{AN} Define MyStyle (Outlook)", True, defaultUser)
            If Username Is Nothing OrElse Username.Trim().Length = 0 Then
                ShowCustomMessageBox("No username provided - exiting.")
                Return
            End If
            Username = Username.Trim()

            ' --- Collect all open emails from Outlook inspectors ---
            Dim app As Outlook.Application = Globals.ThisAddIn.Application
            Dim inspectors As Outlook.Inspectors = ComRetry(Function() app.Inspectors)

            '            Dim mailItems As New System.Collections.Generic.List(Of Outlook.MailItem)()

            'For i As System.Int32 = 1 To inspectors.Count
            'Dim insp As Outlook.Inspector = inspectors.Item(i)
            'If insp IsNot Nothing AndAlso insp.CurrentItem IsNot Nothing Then
            '   If TypeOf insp.CurrentItem Is Outlook.MailItem Then
            '       Dim mi As Outlook.MailItem = CType(insp.CurrentItem, Outlook.MailItem)
            '       If mi IsNot Nothing Then
            '            mailItems.Add(mi)
            '         End If
            '      End If
            '   End If
            'Next

            Dim mailItems As New System.Collections.Generic.List(Of Outlook.MailItem)()

            ' Get count safely
            Dim inspCount As Integer = 0
            Try
                inspCount = ComRetry(Function() inspectors.Count)
            Catch
                inspCount = 0
            End Try

            For i As System.Int32 = 1 To inspCount
                Dim insp As Outlook.Inspector = Nothing
                Try
                    insp = ComRetry(Function() inspectors.Item(i))
                    If insp Is Nothing Then Continue For

                    Dim curr As Object = ComRetry(Function() insp.CurrentItem)
                    Dim mi As Outlook.MailItem = TryCast(curr, Outlook.MailItem)
                    If mi IsNot Nothing Then
                        ' Intentionally keep the MailItem reference; used later in this method.
                        mailItems.Add(mi)
                    End If
                Catch
                    ' Ignore and continue scanning remaining inspectors
                Finally
                    If insp IsNot Nothing Then
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(insp)
                        insp = Nothing
                    End If
                End Try
            Next

            If mailItems.Count = 0 Then
                ShowCustomMessageBox("No open emails were found. Please open all emails you want to include and try again.")
                Return
            End If

            ' --- Show list of all emails that will be included (via MessageBox), then explicit proceed confirm ---
            Dim sbList As New System.Text.StringBuilder()
            sbList.AppendLine("The following emails will be included:").AppendLine()
            For idx As System.Int32 = 0 To mailItems.Count - 1
                Dim mi As Outlook.MailItem = mailItems(idx)
                Dim subj As System.String = If(mi.Subject, "(no subject)")
                Dim sender As System.String = If(mi.SenderName, "(unknown sender)")
                Dim sentOn As System.String
                Try
                    sentOn = If(mi.SentOn = Date.MinValue, "(no sent date)", mi.SentOn.ToString())
                Catch ex As System.Exception
                    sentOn = "(no sent date)"
                End Try
                sbList.AppendLine($"{(idx + 1).ToString("000")}. {subj} — {sender} — {sentOn}")
            Next

            ShowCustomMessageBox(sbList.ToString())
            Dim confirm As System.Int32 = ShowCustomYesNoBox($"Proceed with these emails (the AI will get the full text and be instructed to learn only from those that refer to '{Username}')?", "Continue", "Cancel", $"{AN} Define MyStyle (Outlook)")
            If confirm <> 1 Then
                Return
            End If

            ' --- Additional instructions (like Word: ESC cancels) ---
            OtherPrompt = ""
            OtherPrompt = SLib.ShowCustomInputBox("You can provide additional instructions for the analysis (e.g., Internet links to check [if your model will understand so], aspects to focus on etc.). This is optional.",
                                                    $"{AN} Define MyStyle (Outlook)", False).Trim()
            If OtherPrompt = "ESC" Then
                Return
            End If

            ' --- Optional: use alternate model (like Word) ---
            Dim useSecondAPI As System.Boolean = False
            If Not System.String.IsNullOrWhiteSpace(INI_AlternateModelPath) Then
                answer = ShowCustomYesNoBox($"Do you want to use one of your alternate models?", "Yes, use alternate", "No, use primary", $"{AN} Define MyStyle (Outlook)")
                If answer = 1 Then
                    If Not ShowModelSelection(_context, INI_AlternateModelPath) Then
                        originalConfigLoaded = False
                        Return
                    End If
                    useSecondAPI = True
                ElseIf answer <> 2 Then
                    Return
                End If
            End If

            ' --- Build SelectedText from all open emails ---
            ' Format: <EMAILxxx> Mailtext </EMAILxxx>, xxx = 001, 002, ...
            Dim sbEmails As New System.Text.StringBuilder()

            For idx As System.Int32 = 0 To mailItems.Count - 1
                Dim mi As Outlook.MailItem = mailItems(idx)

                Dim bodyText As System.String = If(mi.Body, System.String.Empty)

                If System.String.IsNullOrWhiteSpace(bodyText) Then
                    Dim html As System.String = If(mi.HTMLBody, System.String.Empty)
                    If Not System.String.IsNullOrWhiteSpace(html) Then
                        ' simple HTML -> text (strip tags, decode)
                        Dim noTags As System.String = System.Text.RegularExpressions.Regex.Replace(html, "<[^>]+>", System.String.Empty)
                        bodyText = System.Net.WebUtility.HtmlDecode(noTags)
                    End If
                End If

                If bodyText Is Nothing Then
                    bodyText = System.String.Empty
                End If

                bodyText = bodyText.Trim()

                Dim tagId As System.String = (idx + 1).ToString("000")
                sbEmails.Append("<EMAIL").Append(tagId).Append(">").AppendLine()
                sbEmails.Append(bodyText).AppendLine()
                sbEmails.Append("</EMAIL").Append(tagId).Append(">").AppendLine().AppendLine()
            Next

            Dim SelectedText As String = sbEmails.ToString()

            ' --- Call LLM with SP_MyStyle_Outlook (like Word) ---
            ' Hinweis: Dein SP_MyStyle_Outlook sollte {OtherPrompt} etc. bereits über InterpolateAtRuntime aufnehmen.
            Dim llmResponse As System.String =
                Await LLM(InterpolateAtRuntime(SP_MyStyle_Outlook), SelectedText, "", "", 0, useSecondAPI)

            ' --- Show analysis and (on OK) save prompt + copy full report to clipboard (like Word) ---
            If Not System.String.IsNullOrWhiteSpace(llmResponse) Then
                Dim analysis As System.String = SLib.ShowCustomWindow($"The AI provided the following style analysis for {Username} and MyStyle prompt of your email samples:",
                                                                        llmResponse,
                                                                        "If you choose 'OK', the prompt and its title at the end of the analysis will be stored in your MyStyle prompt file for future usage (and the full report copied to the clipboard).",
                                                                        AN, False, False, False, False)

                If Not System.String.IsNullOrWhiteSpace(analysis) Then
                    SLib.PutInClipboard(analysis)
                    SLib.ExtractAndStorePromptFromAnalysis(analysis, stylePath, "Outlook")
                End If
            End If

            If useSecondAPI AndAlso originalConfigLoaded Then
                RestoreDefaults(_context, originalConfig)
                originalConfigLoaded = False
            End If


        Catch ex As System.Exception
            ShowCustomMessageBox($"An error occurred: {ex.Message}")
        End Try

    End Sub




    Private Async Sub FreeStyle_InsertBefore(Command As String, Optional AskForPrompt As Boolean = False)
        Try
            Dim outlookApp As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application
            Dim inspector As Inspector = ComRetry(Function() outlookApp.ActiveInspector())

            ' Ensure the inspector is open and the item is a MailItem
            'If inspector Is Nothing OrElse Not TypeOf inspector.CurrentItem Is MailItem Then
            'SLib.ShowCustomMessageBox($"Please create or open an email for editing to use {AN}.")
            'Return
            'End If

            'Dim mailItem As MailItem = DirectCast(inspector.CurrentItem, MailItem)


            ' Guard CurrentItem via ComRetry to avoid transient COM rejections
            Dim curr As Object = Nothing
            If inspector IsNot Nothing Then
                Try
                    curr = ComRetry(Function() inspector.CurrentItem)
                Catch
                    curr = Nothing
                End Try
            End If

            If inspector Is Nothing _
               OrElse curr Is Nothing _
               OrElse Not TypeOf curr Is Microsoft.Office.Interop.Outlook.MailItem Then
                SLib.ShowCustomMessageBox($"Please create or open an email for editing to use {AN}.")
                Return
            End If

            Dim mailItem As Microsoft.Office.Interop.Outlook.MailItem =
                CType(curr, Microsoft.Office.Interop.Outlook.MailItem)


            ' Check if the email is in plain text format
            If mailItem.BodyFormat = OlBodyFormat.olFormatPlain Then
                SLib.ShowCustomMessageBox("This operation is not supported for plain text emails. Switch to HTML or RTF format.")
                Return
            End If

            ' Get the Word editor for the email
            Dim wordEditor As Microsoft.Office.Interop.Word.Document = ComRetry(Function() TryCast(inspector.WordEditor, Microsoft.Office.Interop.Word.Document))

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

                MyStyleInsert = ""
                Dim DoMyStyle As Boolean = False
                Dim StylePath As String = ExpandEnvironmentVariables(INI_MyStylePath)
                If Not String.IsNullOrWhiteSpace(StylePath) And IO.File.Exists(StylePath) Then DoMyStyle = True

                ' Prompt for additional instructions
                OtherPrompt = SLib.ShowCustomInputBox("Please provide additional instructions for drafting an answer (or leave it empty for the most likely substantive response):", $"{AN} Answers", False)
                If OtherPrompt = "ESC" Then Return

                If DoMyStyle Then
                    MyStyleInsert = MyStyleHelpers.SelectPromptFromMyStyle(StylePath, "Outlook", 0, "Choose the style prompt to apply …", $"{AN} MyStyle", True)
                    If MyStyleInsert = "ERROR" Then Return
                    If MyStyleInsert = "NONE" OrElse String.IsNullOrWhiteSpace(MyStyleInsert) Then DoMyStyle = False
                End If

                ' Call your LLM function with the selected text
                LLMResult = Await LLM(InterpolateAtRuntime(SP_MailReply) & If(DoMyStyle, " " & MyStyleInsert, ""), "<MAILCHAIN>" & selectedText & "</MAILCHAIN>", "", "", 0)
            Else
                LLMResult = Await LLM(InterpolateAtRuntime(SP_MailSumup), "<MAILCHAIN>" & selectedText & "</MAILCHAIN>", "", "", 0)
            End If
            If INI_PostCorrection <> "" Then
                LLMResult = Await PostCorrection(LLMResult)
            End If

            'LLMResult = LLMResult.Replace("**", "")  ' Remove bold markers

            ' Convert Markdown to HTML using Markdig
            ' Dim markdownPipeline As MarkdownPipeline = New MarkdownPipelineBuilder().Build()

            Dim builder As New MarkdownPipelineBuilder()

            builder.UsePipeTables()
            builder.UseGridTables()
            builder.UseSoftlineBreakAsHardlineBreak()
            builder.UseListExtras()
            builder.UseFootnotes()
            builder.UseDefinitionLists()
            builder.UseAbbreviations()
            builder.UseAutoLinks()
            builder.UseTaskLists()
            builder.UseEmojiAndSmiley()
            builder.UseMathematics()
            builder.UseFigures()
            builder.UseAdvancedExtensions()
            builder.UseGenericAttributes()

            Dim markdownPipeline As MarkdownPipeline = builder.Build()

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
            'If outlookApp IsNot Nothing Then Marshal.ReleaseComObject(outlookApp) : outlookApp = Nothing

        Catch ex As System.Exception
            MessageBox.Show("Error in Freestyle_InsertBefore: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Async Sub Command_InsertAfter(ByVal SysCommand As String, Optional ByVal DoMarkup As Boolean = False, Optional KeepFormat As Boolean = False, Optional Inplace As Boolean = False, Optional MarkupMethod As Integer = 3)
        Try
            Dim outlookApp As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application
            Dim inspector As Microsoft.Office.Interop.Outlook.Inspector = ComRetry(Function() outlookApp.ActiveInspector())

            ' Ensure the inspector is open and the item is a MailItem
            'If inspector Is Nothing OrElse Not TypeOf inspector.CurrentItem Is Microsoft.Office.Interop.Outlook.MailItem Then
            '   ShowCustomMessageBox("Please open an email to use this function.")
            '   Return
            'End If

            'Dim mailItem As Microsoft.Office.Interop.Outlook.MailItem = DirectCast(inspector.CurrentItem, Microsoft.Office.Interop.Outlook.MailItem)

            ' Guard CurrentItem via ComRetry to avoid transient COM rejections
            Dim curr As Object = Nothing
            If inspector IsNot Nothing Then
                Try
                    curr = ComRetry(Function() inspector.CurrentItem)
                Catch
                    curr = Nothing
                End Try
            End If

            If inspector Is Nothing _
               OrElse curr Is Nothing _
               OrElse Not TypeOf curr Is Microsoft.Office.Interop.Outlook.MailItem Then
                ShowCustomMessageBox("Please open an email to use this function.")
                Return
            End If

            Dim mailItem As Microsoft.Office.Interop.Outlook.MailItem =
                CType(curr, Microsoft.Office.Interop.Outlook.MailItem)

            ' Check if the email is in plain text format
            If mailItem.BodyFormat = Microsoft.Office.Interop.Outlook.OlBodyFormat.olFormatPlain Then
                ShowCustomMessageBox("This operation is not supported for plain text emails. Switch to HTML or RTF format.")
                Return
            End If

            ' Get the Word editor for the email
            Dim wordEditor As Microsoft.Office.Interop.Word.Document = ComRetry(Function() TryCast(inspector.WordEditor, Microsoft.Office.Interop.Word.Document))

            If wordEditor Is Nothing Then
                ShowCustomMessageBox("Unable to access the email editor. Ensure the email is in HTML or RTF format.")
                Return
            End If



            ' Get the selected text and range
            Dim selection As Microsoft.Office.Interop.Word.Selection = wordEditor.Application.Selection
            Dim range As Microsoft.Office.Interop.Word.Range = selection.Range.Duplicate ' Duplicate to preserve original
            Dim SelectedText As String

            'Try
            'Using New WordUndoScope(wordEditor, $"{AN} Changes")


            If INI_KeepFormatCap > 0 Then If Len(selection.Text) > INI_KeepFormatCap Then KeepFormat = False

            If KeepFormat Then
                SelectedText = SLib.GetRangeHtml(selection.Range)
            Else
                If INI_MarkdownConvert Then ConvertRangeToMarkdown(selection.Range)
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

            Dim trailingCR As Boolean = SelectedText.EndsWith(vbCrLf) Or SelectedText.EndsWith(vbCr) Or SelectedText.EndsWith(vbLf)

            ' Call your LLM function with the selected text
            Dim LLMResult As String = Await LLM(SysCommand & If(KeepFormat, " " & SP_Add_KeepHTMLIntact, SP_Add_KeepInlineIntact), "<TEXTTOPROCESS>" & SelectedText & "</TEXTTOPROCESS>", "", "", 0)

            LLMResult = LLMResult.Replace("<TEXTTOPROCESS>", "").Replace("</TEXTTOPROCESS>", "")

            If INI_PostCorrection <> "" Then
                LLMResult = Await PostCorrection(LLMResult)
            End If

            Debug.WriteLine("TrailingCR=" & trailingCR)
            Debug.WriteLine($"Selection='{selection.Text}'")

            ' Replace the selected text with the processed result
            If Not String.IsNullOrWhiteSpace(LLMResult) Then
                If KeepFormat Then

                    Dim Plaintext As String = ""

                    SelectedText = selection.Text
                    SLib.InsertTextWithFormat(LLMResult, range, Inplace, Not trailingCR)
                    If DoMarkup Then
                        LLMResult = SLib.RemoveHTML(LLMResult)
                        If MarkupMethod <> 3 Then
                            range.Text = vbCrLf & vbCrLf & "MARKUP:" & vbCrLf
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
                            SLib.InsertTextWithMarkdown(selection, LLMResult & vbCrLf & "<p>MARKUP:<br></p>", trailingCR)
                            'selection.TypeText(LLMResult & vbCrLf & vbCrLf & "MARKUP:" & vbCrLf & vbCrLf)
                        Else
                            SLib.InsertTextWithMarkdown(selection, LLMResult, trailingCR)
                            'selection.TypeText(LLMResult)
                        End If
                    Else
                        ' Replace this line:
                        ' selection.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)

                        ' With the following code to insert two new lines and select the last one, preserving formatting:
                        Dim selRange As Microsoft.Office.Interop.Word.Range = selection.Range.Duplicate
                        Dim originalFont As Microsoft.Office.Interop.Word.Font = selRange.Font.Duplicate

                        ' Insert two new lines at the end of the selection
                        selRange.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd)
                        selRange.Text = vbCrLf & vbCrLf

                        ' Select the last new line
                        Dim newStart As Integer = selRange.End - 2 ' Position at the start of the last vbCrLf
                        Dim newEnd As Integer = selRange.End
                        selection.SetRange(newStart, newEnd)

                        ' Reapply the original formatting to the new selection
                        selection.Font = originalFont

                        If DoMarkup And MarkupMethod <> 3 Then
                            'selection.TypeText(vbCrLf & LLMResult & vbCrLf & vbCrLf & "MARKUP:" & vbCrLf & vbCrLf)
                            SLib.InsertTextWithMarkdown(selection, LLMResult & vbCrLf & "<p>MARKUP:<br></p>" & vbCrLf, trailingCR)
                        Else
                            'selection.TypeText(vbCrLf & LLMResult & vbCrLf)
                            SLib.InsertTextWithMarkdown(selection, LLMResult, trailingCR)

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

            ' End Using

            'Catch ex As System.Exception
            '   Debug.WriteLine("Error in Undo: " & ex.Message)
            'End Try

            ' Refresh the inspector to show updated content
            inspector.Display()

            ' Release COM objects in reverse order of creation
            If range IsNot Nothing Then Marshal.ReleaseComObject(range) : range = Nothing
            If selection IsNot Nothing Then Marshal.ReleaseComObject(selection) : selection = Nothing
            If wordEditor IsNot Nothing Then Marshal.ReleaseComObject(wordEditor) : wordEditor = Nothing
            If mailItem IsNot Nothing Then Marshal.ReleaseComObject(mailItem) : mailItem = Nothing
            If inspector IsNot Nothing Then Marshal.ReleaseComObject(inspector) : inspector = Nothing
            'If outlookApp IsNot Nothing Then Marshal.ReleaseComObject(outlookApp) : outlookApp = Nothing

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
            Dim DoFileObject As Boolean = False
            Dim FileObject As String = ""
            Dim DoNewDoc As Boolean = False
            Dim DoMyStyle As Boolean = False

            Dim UseSecondAPI As Boolean = False

            Dim MarkupInstruct As String = $"start with '{MarkupPrefixAll}' for markups"
            Dim InplaceInstruct As String = $"use '{InPlacePrefix}' for replacing your current selection"
            Dim ClipboardInstruct As String = $"with '{ClipboardPrefix}'/'{ClipboardPrefix2}' or '{NewDocPrefix}' to have the result in a window or new Word document"
            Dim PromptLibInstruct As String = If(INI_PromptLib, " or press 'OK' for the prompt library", "")
            Dim NoFormatInstruct As String = $"; add '{NoFormatTrigger2}'/'{KFTrigger2}'/'{KPFTrigger2}' for overriding formatting defaults"
            Dim MyStyleInstruct As String = $"; add '{MyStyleTrigger}' to apply your personal style"
            Dim SecondAPIInstruct As String = If(INI_SecondAPI, $"'{SecondAPICode}' to use {If(String.IsNullOrWhiteSpace(INI_AlternateModelPath), $"the secondary model ({INI_Model_2})", "one of the other models")}", "")
            Dim LastPromptInstruct As String = If(String.IsNullOrWhiteSpace(My.Settings.LastPrompt), "", "; Ctrl-P for your last prompt")
            Dim ObjectInstruct As String = $"; add '{ObjectTrigger2}' for including a clipboard object"

            Dim AddOnInstruct As String = "; add " & SecondAPIInstruct

            Dim lastCommaIndex As Integer = AddOnInstruct.LastIndexOf(","c)
            If lastCommaIndex <> -1 Then
                AddOnInstruct = AddOnInstruct.Substring(0, lastCommaIndex) & ", and" & AddOnInstruct.Substring(lastCommaIndex + 1)
            End If
            If Not String.IsNullOrWhiteSpace(INI_MyStylePath) Then
                AddOnInstruct += MyStyleInstruct.Replace("; add ", ", ")
            End If

            Dim outlookApp As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application
            Dim inspector As Microsoft.Office.Interop.Outlook.Inspector = ComRetry(Function() outlookApp.ActiveInspector())

            ' Ensure the inspector is open and the item is a MailItem
            'If inspector Is Nothing OrElse Not TypeOf inspector.CurrentItem Is Microsoft.Office.Interop.Outlook.MailItem Then
            '  SLib.ShowCustomMessageBox($"Please create or open an email for editing to use {AN}.")
            '   Return
            'End If

            'Dim mailItem As Microsoft.Office.Interop.Outlook.MailItem = DirectCast(inspector.CurrentItem, Microsoft.Office.Interop.Outlook.MailItem)

            ' Guard CurrentItem via ComRetry to avoid transient COM rejections
            Dim curr As Object = Nothing
            If inspector IsNot Nothing Then
                Try
                    curr = ComRetry(Function() inspector.CurrentItem)
                Catch
                    curr = Nothing
                End Try
            End If

            If inspector Is Nothing _
               OrElse curr Is Nothing _
               OrElse Not TypeOf curr Is Microsoft.Office.Interop.Outlook.MailItem Then
                SLib.ShowCustomMessageBox($"Please create or open an email for editing to use {AN}.")
                Return
            End If

            Dim mailItem As Microsoft.Office.Interop.Outlook.MailItem =
                CType(curr, Microsoft.Office.Interop.Outlook.MailItem)

            ' Check if the email is in plain text format
            If mailItem.BodyFormat = Microsoft.Office.Interop.Outlook.OlBodyFormat.olFormatPlain Then
                SLib.ShowCustomMessageBox("This operation is not supported for plain text emails. Switch to HTML or RTF format.")
                Return
            End If

            ' Get the Word editor for the email
            Dim wordEditor As Microsoft.Office.Interop.Word.Document = ComRetry(Function() TryCast(inspector.WordEditor, Microsoft.Office.Interop.Word.Document))

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

            If UseSecondAPI Then
                If Not String.IsNullOrWhiteSpace(INI_APICall_Object_2) Then
                    AddOnInstruct += ObjectInstruct.Replace("; add", ",")
                    DoFileObject = True
                End If
            Else
                If Not String.IsNullOrWhiteSpace(INI_APICall_Object) Then
                    AddOnInstruct += ObjectInstruct.Replace("; add", ",")
                    DoFileObject = True
                End If
            End If

            ' Prompt for the text to process

            If Not NoText Then
                Dim OptionalButtons As System.Tuple(Of String, String, String)() = {
                            System.Tuple.Create("OK, use window", $"Use this to automatically insert '{ClipboardPrefix}' as a prefix.", ClipboardPrefix),
                            System.Tuple.Create("OK, do a new doc", $"Use this to automatically insert '{NewDocPrefix}' as a prefix.", NewDocPrefix),
                            System.Tuple.Create("OK, do a markup", $"Use this to automatically insert '{MarkupPrefixDiff}' as a prefix.", MarkupPrefixDiff)
                        }
                OtherPrompt = SLib.ShowCustomInputBox($"Please provide the prompt you wish to execute on the selected text ({MarkupInstruct}, {InplaceInstruct}, {ClipboardInstruct}){PromptLibInstruct}{AddOnInstruct}{LastPromptInstruct}:", $"{AN} Freestyle", False, "", My.Settings.LastPrompt, OptionalButtons)
            Else
                Dim OptionalButtons As System.Tuple(Of String, String, String)() = {
                            System.Tuple.Create("OK, use window", $"Use this to automatically insert '{ClipboardPrefix}' as a prefix.", ClipboardPrefix),
                            System.Tuple.Create("OK, do a new doc", $"Use this to automatically insert '{NewDocPrefix}' as a prefix.", NewDocPrefix)
                        }

                OtherPrompt = SLib.ShowCustomInputBox($"Please provide the prompt you wish to execute ({ClipboardInstruct}){PromptLibInstruct}{AddOnInstruct}{LastPromptInstruct}:", $"{AN} Freestyle", False, "", My.Settings.LastPrompt, OptionalButtons)
            End If

            If String.IsNullOrEmpty(OtherPrompt) AndAlso OtherPrompt <> "ESC" AndAlso INI_PromptLib Then

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

            If Not SharedMethods.ProcessParameterPlaceholders(OtherPrompt) Then
                ShowCustomMessageBox("Freestyle canceled.", $"{AN} Freestyle")
                Exit Sub
            End If

            ' Check if otherPrompt starts with "Markup:" (case-insensitive)

            If OtherPrompt.StartsWith(ClipboardPrefix, StringComparison.OrdinalIgnoreCase) Then
                OtherPrompt = OtherPrompt.Substring(ClipboardPrefix.Length).Trim()
                DoClipboard = True
            ElseIf OtherPrompt.StartsWith(ClipboardPrefix2, StringComparison.OrdinalIgnoreCase) Then
                OtherPrompt = OtherPrompt.Substring(ClipboardPrefix2.Length).Trim()
                DoClipboard = True
            ElseIf OtherPrompt.StartsWith(NewDocPrefix, StringComparison.OrdinalIgnoreCase) Then
                OtherPrompt = OtherPrompt.Substring(NewDocPrefix.Length).Trim()
                DoClipboard = True
                DoNewDoc = True

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
            If DoFileObject AndAlso OtherPrompt.IndexOf(ObjectTrigger2, StringComparison.OrdinalIgnoreCase) >= 0 Then
                OtherPrompt = OtherPrompt.Replace(ObjectTrigger2, "(a file object follows)").Trim()
                FileObject = "clipboard"
            End If

            If Not String.IsNullOrWhiteSpace(INI_MyStylePath) And OtherPrompt.IndexOf(MyStyleTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
                Dim StylePath As String = ExpandEnvironmentVariables(INI_MyStylePath)
                If Not IO.File.Exists(StylePath) Then
                    ShowCustomMessageBox("No MyStyle prompt file has been found. You may have to first create a MyStyle prompt. Go to 'Analyze' and use 'Define MyStyle' to do so - exiting.")
                    Return
                End If
                OtherPrompt = OtherPrompt.Replace(MyStyleTrigger, "").Trim()
                MyStyleInsert = MyStyleHelpers.SelectPromptFromMyStyle(StylePath, "Word", 0, "Choose the style prompt to apply …", $"{AN} MyStyle", True)
                If MyStyleInsert = "ERROR" Then Return
                If MyStyleInsert = "NONE" OrElse String.IsNullOrWhiteSpace(MyStyleInsert) Then Return
                DoMyStyle = True
            End If

            If INI_SecondAPI Then
                If OtherPrompt.Contains(SecondAPICode) Then
                    UseSecondAPI = True
                    OtherPrompt = OtherPrompt.Replace(SecondAPICode, "").Trim()

                    If Not String.IsNullOrWhiteSpace(INI_AlternateModelPath) Then

                        If Not ShowModelSelection(_context, INI_AlternateModelPath) Then
                            originalConfigLoaded = False
                            Return
                        End If

                    End If

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
                LLMResult = Await LLM(InterpolateAtRuntime(SP_FreestyleText) & If(DoMyStyle, " " & MyStyleInsert, ""), "<TEXTTOPROCESS>" & selectedText & "</TEXTTOPROCESS>", "", "", 0, UseSecondAPI, False, OtherPrompt, FileObject)

                LLMResult = LLMResult.Replace("<TEXTTOPROCESS>", "").Replace("</TEXTTOPROCESS>", "")
            Else
                LLMResult = Await LLM(InterpolateAtRuntime(SP_FreestyleNoText) & If(DoMyStyle, " " & MyStyleInsert, ""), "", "", "", 0, UseSecondAPI, False, OtherPrompt, FileObject)
            End If

            If INI_PostCorrection <> "" Then
                LLMResult = Await PostCorrection(LLMResult)
            End If

            OtherPrompt = ""

            If DoNewDoc Then
                Try
                    ' Create a new instance of Word
                    Dim wordApp As New Microsoft.Office.Interop.Word.Application
                    wordApp.Visible = True

                    ' Add a new document
                    Dim newDoc As Microsoft.Office.Interop.Word.Document = wordApp.Documents.Add()

                    ' Insert your text (LLMResult) at the beginning
                    Dim docSelection As Microsoft.Office.Interop.Word.Selection = wordApp.Selection
                    InsertTextWithMarkdown(docSelection, LLMResult, True)

                Catch Ex As System.Exception
                    Dim FinalText As String = SLib.ShowCustomWindow("The Word document could not be created or the LLM output not inserted. Here is the result of the LLM (you can edit it):", LLMResult, "You can choose whether you want to have the original text put into the clipboard or your text with any changes you have made. If you select Cancel, nothing will be put into the clipboard (without formatting).", AN, False)

                    If FinalText <> "" Then
                        SLib.PutInClipboard(FinalText)
                    End If

                End Try

            ElseIf DoClipboard Then
                Dim FinalText As String = SLib.ShowCustomWindow("The LLM has provided the following result (you can edit it):", LLMResult, "You can choose whether you want to have the original text put into the clipboard or your text with any changes you have made. If you select Cancel, nothing will be put into the clipboard (without formatting).", AN, False)

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
                    SLib.InsertTextWithMarkdown(selection, vbCrLf & LLMResult & vbCrLf & "<p>MARKUP:<br></p>", trailingCR)
                Else
                    If DoInplace Then
                        SLib.InsertTextWithMarkdown(selection, LLMResult, trailingCR)
                    Else
                        SLib.InsertTextWithMarkdown(selection, vbCrLf & LLMResult & vbCrLf, trailingCR)
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
            'If outlookApp IsNot Nothing Then Marshal.ReleaseComObject(outlookApp) : outlookApp = Nothing

            If UseSecondAPI And originalConfigLoaded Then
                RestoreDefaults(_context, originalConfig)
                originalConfigLoaded = False
            End If

        Catch ex As System.Exception
            MessageBox.Show("Error in Freestyle_InsertAfter: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Private Sub CompareAndInsertTextCompareDocs(input1 As String, input2 As String)

        Dim splash As New SplashScreen("Creating markup using the Word compare functionality (ignore any flickering and press 'No' if prompted) ...")
        splash.Show()
        splash.Refresh()
        Try
            ' Get the active inspector (compose mail window)
            Dim outlookApp As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application
            Dim inspector As Inspector = ComRetry(Function() outlookApp.ActiveInspector)

            ' Ensure the current item is a MailItem and in compose mode

            ' Ensure the current item is a MailItem and in compose mode (COM-safe)
            If inspector Is Nothing Then
                System.Windows.Forms.MessageBox.Show("Error in CompareAndInsertTextCompareDocs: No active inspector.",
                                         "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            Dim curr As Object = Nothing
            Try
                curr = ComRetry(Function() inspector.CurrentItem)
            Catch
                curr = Nothing
            End Try

            If curr Is Nothing OrElse Not TypeOf curr Is Microsoft.Office.Interop.Outlook.MailItem Then
                System.Windows.Forms.MessageBox.Show("Error in CompareAndInsertTextCompareDocs: No active email item.",
                                         "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            Dim mailItem As Microsoft.Office.Interop.Outlook.MailItem =
    ComRetry(Function() CType(curr, Microsoft.Office.Interop.Outlook.MailItem))

            Dim editor As Object = ComRetry(Function() inspector.WordEditor)


            'If TypeOf inspector.CurrentItem Is MailItem Then
            'Dim mailItem As MailItem = CType(inspector.CurrentItem, MailItem)
            'Dim editor As Object = inspector.WordEditor

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

            'Else
            'MessageBox.Show("Error in CompareAndInsertTextCompareDocs: The mail compose window is not open (anymore).", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            'End If

            ' Release COM objects in reverse order of creation
            If inspector IsNot Nothing Then Marshal.ReleaseComObject(inspector) : inspector = Nothing
            'If outlookApp IsNot Nothing Then Marshal.ReleaseComObject(outlookApp) : outlookApp = Nothing

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
            InsertFormattedTextFast(sText & vbCrLf)
        Else
            Dim htmlContent As String = ConvertMarkupToRTF(TextforWindow & "\r\r" & sText)
            System.Threading.Tasks.Task.Run(Sub()
                                                ShowRTFCustomMessageBox(htmlContent)
                                            End Sub)
        End If

    End Sub


    Private Shared Sub ConvertRangeToMarkdown(WorkingRange As Word.Range)

        Dim listRegex As New Regex("^(\s*)([-*+]|\d+[\.\)])\s+", RegexOptions.Compiled)


        Dim rng As Word.Range = WorkingRange.Duplicate
        If rng.End < rng.Document.Content.End - 1 Then
            rng.End = rng.End + 1
        End If

        Dim doc As Microsoft.Office.Interop.Word.Document = rng.Document

        ' 0) Überschriften & Listen
        For Each para As Microsoft.Office.Interop.Word.Paragraph In rng.Paragraphs
            Dim styleName As String = CType(para.Style, Microsoft.Office.Interop.Word.Style).NameLocal

            Select Case styleName
                Case doc.Styles(Word.WdBuiltinStyle.wdStyleTitle).NameLocal
                    para.Range.InsertBefore("# ")
                Case doc.Styles(Word.WdBuiltinStyle.wdStyleHeading1).NameLocal
                    para.Range.InsertBefore("# ")
                Case doc.Styles(Word.WdBuiltinStyle.wdStyleHeading2).NameLocal
                    para.Range.InsertBefore("## ")
                Case doc.Styles(Word.WdBuiltinStyle.wdStyleHeading3).NameLocal
                    para.Range.InsertBefore("### ")
                    ' … und so weiter bis Heading6 …
            End Select

            ' — Listen erkennen
            With para.Range.ListFormat
                Try
                    ' Nur fortfahren, wenn eine Liste vorliegt
                    If .ListType <> Microsoft.Office.Interop.Word.WdListType.wdListNoNumbering Then

                        ' 1) Alle nötigen Infos VOR RemoveNumbers speichern
                        Dim lvl As Integer = .ListLevelNumber
                        Dim lt As Microsoft.Office.Interop.Word.WdListType = .ListType
                        Dim ls As String = .ListString.Trim()

                        ' 2) Prefix berechnen (4 Spaces pro Ebene)
                        Dim indent As String = New String(" "c, (lvl - 1) * 4)
                        Dim prefix As String
                        Select Case lt
                            Case Microsoft.Office.Interop.Word.WdListType.wdListBullet,
                                 Microsoft.Office.Interop.Word.WdListType.wdListPictureBullet
                                prefix = indent & "- "
                            Case Microsoft.Office.Interop.Word.WdListType.wdListSimpleNumbering,
                                 Microsoft.Office.Interop.Word.WdListType.wdListOutlineNumbering,
                                 Microsoft.Office.Interop.Word.WdListType.wdListMixedNumbering,
                                 Microsoft.Office.Interop.Word.WdListType.wdListListNumOnly
                                prefix = indent & ls & " "
                            Case Else
                                prefix = indent & "- "
                        End Select

                        ' 3) Liste entfernen
                        .RemoveNumbers()

                        ' 4) Markdown-Präfix am Zeilenanfang einfügen
                        Dim insertRange As Microsoft.Office.Interop.Word.Range = para.Range.Duplicate()
                        insertRange.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseStart)
                        insertRange.InsertBefore(prefix)

                        ' Range für das eingefügte Prefix erstellen
                        Dim prefixRange As Word.Range = insertRange.Duplicate
                        prefixRange.End = prefixRange.Start + prefix.Length

                        ' Font zurücksetzen (Standard-Formatierung)
                        prefixRange.Font.Reset()

                    End If

                Catch ex As System.Exception
                    System.Diagnostics.Debug.WriteLine("Fehler bei Listenkonvertierung: " & ex.ToString())
                End Try
            End With

        Next

        ' 1) Fett + Italic  (Absatz)
        ReplaceWithinRange(rng,
                        Sub(f)
                            f.Font.Bold = True
                            f.Font.Italic = True
                            f.Font.Underline = Word.WdUnderline.wdUnderlineNone
                            f.Text = "(*)^13"
                            f.MatchWildcards = True
                        End Sub,
                        "***\1***^13",
                        Sub(rep)                          ' nur Bold & Italic abstellen
                            rep.Bold = False
                            rep.Italic = False
                        End Sub)

        ' 2) Fett + Italic  (Inline)
        ReplaceWithinRange(rng,
                        Sub(f)
                            f.Font.Bold = True
                            f.Font.Italic = True
                            f.Font.Underline = Word.WdUnderline.wdUnderlineNone
                            f.Text = ""
                            f.MatchWildcards = False
                        End Sub,
                        "***^&***",
                        Sub(rep)
                            rep.Bold = False
                            rep.Italic = False
                        End Sub)


        ' 3) Nur Fett  (Absatz)
        ReplaceWithinRange(rng,
                        Sub(f)
                            f.Font.Bold = True
                            f.Text = "(*)^13"
                            f.MatchWildcards = True
                        End Sub,
                        "**\1**^13",
                        Sub(rep)
                            rep.Bold = False
                        End Sub)

        ' 4) Nur Fett  (Inline)
        ReplaceWithinRange(rng,
                        Sub(f)
                            f.Font.Bold = True
                            f.Text = ""
                            f.MatchWildcards = False
                        End Sub,
                        "**^&**",
                        Sub(rep)
                            rep.Bold = False
                        End Sub)


        ' 5) Nur Italic  (Absatz)
        ReplaceWithinRange(rng,
                        Sub(f)
                            f.Font.Italic = True
                            f.Text = "(*)^13"
                            f.MatchWildcards = True
                        End Sub,
                        "*\1*^13",
                        Sub(rep)
                            rep.Italic = False
                        End Sub)

        ' 6) Nur Italic  (Inline)
        ReplaceWithinRange(rng,
                        Sub(f)
                            f.Font.Italic = True
                            f.Text = ""
                            f.MatchWildcards = False
                        End Sub,
                        "*^&*",
                        Sub(rep)
                            rep.Italic = False
                        End Sub)


        ' 7) Underline  (Absatz)
        ReplaceWithinRange(rng,
                        Sub(f)
                            f.Font.Underline = Word.WdUnderline.wdUnderlineSingle
                            f.Text = "(*)^13"
                            f.MatchWildcards = True
                        End Sub,
                        "<u>\1</u>^13",
                        Sub(rep)
                            rep.Underline = Word.WdUnderline.wdUnderlineNone
                        End Sub)

        ' 8) Underline  (Inline)
        ReplaceWithinRange(rng,
                        Sub(f)
                            f.Font.Underline = Word.WdUnderline.wdUnderlineSingle
                            f.Text = ""
                            f.MatchWildcards = False
                        End Sub,
                        "<u>^&</u>",
                        Sub(rep)
                            rep.Underline = Word.WdUnderline.wdUnderlineNone
                        End Sub)

        ' 9) Strikethrough  (Absatz)
        ReplaceWithinRange(rng,
                        Sub(f)
                            f.Font.StrikeThrough = True
                            f.Text = "(*)^13"
                            f.MatchWildcards = True
                        End Sub,
                        "~~\1~~^13",
                        Sub(rep)
                            rep.StrikeThrough = False
                        End Sub)

        '10) Strikethrough  (Inline)
        ReplaceWithinRange(rng,
                        Sub(f)
                            f.Font.StrikeThrough = True
                            f.Text = ""
                            f.MatchWildcards = False
                        End Sub,
                        "~~^&~~",
                        Sub(rep)
                            rep.StrikeThrough = False
                        End Sub)


        ' Auswahl wiederherstellen
        'rng = workingrange.Duplicate

        rng.End = rng.End - 1
        rng.Select()

    End Sub

    Private Shared Sub ReplaceWithinRange(
    ByVal rng As Word.Range,
    ByVal configureFind As Action(Of Word.Find),
    ByVal replacementText As String,
    ByVal tweakReplacement As Action(Of Word.Font))

        Dim doc As Word.Document = rng.Document
        Dim originalStart As Long = rng.Start
        Dim originalEnd As Long = rng.End
        Dim currentPosition As Long = originalStart

        Do
            ' Create a range from current position to the end of the original range
            Dim searchRange As Word.Range = doc.Range(currentPosition, originalEnd)
            Dim f As Word.Find = searchRange.Find

            Debug.WriteLine($"Searchrange = '{searchRange.Text}'")

            f.ClearFormatting()
            f.Replacement.ClearFormatting()

            configureFind(f)
            f.Replacement.Text = replacementText
            tweakReplacement(f.Replacement.Font)

            f.Forward = True
            f.Wrap = Word.WdFindWrap.wdFindStop
            f.Format = True

            ' If no more matches, exit
            If Not f.Execute(Replace:=Word.WdReplace.wdReplaceOne) Then Exit Do

            Debug.WriteLine($"Searchrange = '{searchRange.Text}' (after change)")

            ' After replacement, searchRange now points to the match
            ' Check if this match/replacement went beyond our boundary
            If searchRange.End > originalEnd Then
                Debug.WriteLine("Went too far!")
                doc.Undo()
                Exit Do
            End If

            ' Set the current position to continue from the end of this match
            currentPosition = searchRange.End
            originalEnd = rng.End

        Loop While currentPosition < originalEnd

        ' Update the original range to reflect the final processed area
        rng.SetRange(originalStart, originalEnd)
    End Sub



    Private Function ConvertRtfToPlainText(rtfContent As String) As String
        If String.IsNullOrWhiteSpace(rtfContent) Then Return String.Empty

        ' Remove RTF headers and control words
        Dim plainText As String = Regex.Replace(rtfContent, "{\\.*?}|\\[a-z]+[0-9]*|[{}]", String.Empty)

        ' Decode escaped characters (e.g., \'xx)
        plainText = Regex.Replace(plainText, "\\'([0-9a-fA-F]{2})", Function(m)
                                                                        Dim hex = System.Convert.ToByte(m.Groups(1).Value, 16)
                                                                        Return Chr(hex)
                                                                    End Function)

        ' Replace RTF line breaks (\par) with actual line breaks
        plainText = Regex.Replace(plainText, "\\par", Environment.NewLine, RegexOptions.IgnoreCase)

        ' Trim the result
        plainText = plainText.Trim()

        Return plainText
    End Function


    Private Sub InsertFormattedTextFast(ByVal inputText As String)

        '------------------------------------------------------------
        ' 1. Convert the pseudo-tags to plain HTML
        '------------------------------------------------------------
        Dim markup As String = System.Net.WebUtility.HtmlEncode(inputText)

        'Preserve line breaks (optional – remove if you prefer real paragraphs)
        markup = markup.Replace(vbCrLf, "<br>")

        'Replace each tag with an inline <span>
        markup = markup.Replace("[INS_START]",
                "<span style=""color:#0000FF;text-decoration:underline;"">") _
                   .Replace("[INS_END]", "</span>") _
                   .Replace("[DEL_START]",
                "<span style=""color:#FF0000;text-decoration:line-through;"">") _
                   .Replace("[DEL_END]", "</span>")

        '------------------------------------------------------------
        ' 2. Get the current Outlook inspector / Word selection
        '------------------------------------------------------------
        Dim inspector As Microsoft.Office.Interop.Outlook.Inspector =
        ComRetry(Function() TryCast(Globals.ThisAddIn.Application.ActiveInspector,
                Microsoft.Office.Interop.Outlook.Inspector))

        If inspector Is Nothing Then
            System.Windows.Forms.MessageBox.Show(
            "No open mail item found.",
            "InsertFormattedTextFast",
            System.Windows.Forms.MessageBoxButtons.OK,
            System.Windows.Forms.MessageBoxIcon.Error)
            Exit Sub
        End If

        Dim wordDoc As Microsoft.Office.Interop.Word.Document =
        ComRetry(Function() TryCast(inspector.WordEditor,
                Microsoft.Office.Interop.Word.Document))

        If wordDoc Is Nothing Then
            System.Windows.Forms.MessageBox.Show(
            "Unable to access the Word editor.",
            "InsertFormattedTextFast",
            System.Windows.Forms.MessageBoxButtons.OK,
            System.Windows.Forms.MessageBoxIcon.Error)
            Exit Sub
        End If

        Dim app As Microsoft.Office.Interop.Word.Application = wordDoc.Application
        Dim selRange As Microsoft.Office.Interop.Word.Range = app.Selection.Range

        '------------------------------------------------------------
        ' 3. Insert the fragment in one shot
        '------------------------------------------------------------
        Dim oldScreenUpdating As Boolean = app.ScreenUpdating
        app.ScreenUpdating = False

        Try
            'Your existing helper that pastes an HTML fragment
            InsertTextWithFormat(markup, selRange, True, True)

        Catch ex As System.Exception
            'Handle or log as needed – shows a message box here for completeness
            System.Windows.Forms.MessageBox.Show(
            ex.Message,
            "InsertFormattedTextFast",
            System.Windows.Forms.MessageBoxButtons.OK,
            System.Windows.Forms.MessageBoxIcon.Error)

        Finally
            'Restore Word UI state
            app.ScreenUpdating = oldScreenUpdating

            '--------------------------------------------------------
            ' 4. Clean up COM objects in reverse order of creation
            '--------------------------------------------------------
            If selRange IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(selRange)
                selRange = Nothing
            End If
            If wordDoc IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wordDoc)
                wordDoc = Nothing
            End If
            If inspector IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(inspector)
                inspector = Nothing
            End If
        End Try
    End Sub



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
        objInspector = ComRetry(Function() TryCast(Globals.ThisAddIn.Application.ActiveInspector, Microsoft.Office.Interop.Outlook.Inspector))
        If objInspector Is Nothing Then
            MessageBox.Show("Error in InsertFormattedText: No open mail item found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        ' Get the Word editor and the current selection
        objWordDoc = ComRetry(Function() TryCast(objInspector.WordEditor, Microsoft.Office.Interop.Word.Document))
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

            'System.Windows.Forms.Application.DoEvents()

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
            Dim outlookApp As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application
            Dim inspector As Microsoft.Office.Interop.Outlook.Inspector = ComRetry(Function() outlookApp.ActiveInspector())

            ' Ensure the inspector is open and the item is a MailItem
            'If inspector Is Nothing OrElse Not TypeOf inspector.CurrentItem Is Microsoft.Office.Interop.Outlook.MailItem Then
            '   Return 0
            'End If

            'Dim mailItem As Microsoft.Office.Interop.Outlook.MailItem =
            'DirectCast(inspector.CurrentItem, Microsoft.Office.Interop.Outlook.MailItem)

            ' Guard CurrentItem via ComRetry to avoid transient COM rejections
            Dim curr As Object = Nothing
            If inspector IsNot Nothing Then
                Try
                    curr = ComRetry(Function() inspector.CurrentItem)
                Catch
                    curr = Nothing
                End Try
            End If

            If inspector Is Nothing _
               OrElse curr Is Nothing _
               OrElse Not TypeOf curr Is Microsoft.Office.Interop.Outlook.MailItem Then
                Return 0
            End If

            Dim mailItem As Microsoft.Office.Interop.Outlook.MailItem =
                CType(curr, Microsoft.Office.Interop.Outlook.MailItem)

            ' Check if the email is in plain text format
            If mailItem.BodyFormat = Microsoft.Office.Interop.Outlook.OlBodyFormat.olFormatPlain Then
                Return 0
            End If

            ' Get the Word editor for the email
            Dim wordEditor As Microsoft.Office.Interop.Word.Document =
            ComRetry(Function() TryCast(inspector.WordEditor, Microsoft.Office.Interop.Word.Document))

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
            'If outlookApp IsNot Nothing Then Marshal.ReleaseComObject(outlookApp) : outlookApp = Nothing

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
                        {"Clean", "Clean the LLM response"},
                        {"MarkdownConvert", "Keep character formatting"},
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
                        {"Clean", "To remove double-spaces and hidden markers that may have been inserted by the LLM"},
                        {"MarkdownConvert", "If selected, bold, italic, underline and some more formatting will be preserved converting it to Markdown coding before passing it to the LLM (most LLM support it)"},
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
    Private isShuttingDown As Boolean = False
    Private listenerTask As System.Threading.Tasks.Task

    Private llmOperationCts As System.Threading.CancellationTokenSource

    Private activeRequests As Integer = 0
    Private ModelTimeout As Integer = 300

    ' --- Threading & Listener State (Klassenebene) ---
    Private llmSyncContext As System.Threading.SynchronizationContext
    Private llmThread As System.Threading.Thread
    Private cts As System.Threading.CancellationTokenSource

    Private wasListenerRunningBeforeSleep As System.Boolean = False
    Private wasLlmThreadAliveBeforeSleep As System.Boolean = False
    Private restartingAfterResume As System.Int32 = 0  ' 0/1 via Interlocked

    ' guard to mute watchdog and concurrent restarts during power transitions
    Private powerChanging As System.Int32 = 0

    ' Generation protection (pre/post sleep)
    Private listenerGeneration As System.Int64 = 0

    ' Progress watchdog
    Private lastListenerProgressUtc As System.DateTime = System.DateTime.UtcNow
    Private watchdog As System.Threading.Timer

    ' --- Power notifications via hidden window ---
    Private powerWindow As PowerNotificationWindow

    ' Add near your other fields
    Private resumeCooldownUntilUtc As System.DateTime = System.DateTime.MinValue
    Private inMainMenu As System.Int32 = 0

    ' Add a helper
    Private Function IsInResumeCooldown() As System.Boolean
        Return System.DateTime.UtcNow < resumeCooldownUntilUtc
    End Function


    Private NotInheritable Class PowerNotificationWindow
        Inherits System.Windows.Forms.NativeWindow
        Implements System.IDisposable

        Private Const WM_POWERBROADCAST As System.Int32 = &H218
        Private Const PBT_APMSUSPEND As System.Int32 = &H4
        Private Const PBT_APMRESUMEAUTOMATIC As System.Int32 = &H12
        Private Const PBT_APMRESUMESUSPEND As System.Int32 = &H7

        Private ReadOnly owner As ThisAddIn

        Public Sub New(ByVal owner As ThisAddIn)
            Me.owner = owner
            Dim cp As New System.Windows.Forms.CreateParams()
            cp.Caption = "InkyPowerWnd"
            cp.X = 0 : cp.Y = 0 : cp.Height = 0 : cp.Width = 0
            cp.Style = 0 : cp.ExStyle = 0
            ' WICHTIG: Message-only window
            cp.Parent = New System.IntPtr(-3) ' HWND_MESSAGE
            Me.CreateHandle(cp)
        End Sub



        Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
            Const WM_POWERBROADCAST As System.Int32 = &H218
            Const PBT_APMQUERYSUSPEND As System.Int32 = &H0
            Const PBT_APMSUSPEND As System.Int32 = &H4
            Const PBT_APMRESUMEAUTOMATIC As System.Int32 = &H12
            Const PBT_APMRESUMESUSPEND As System.Int32 = &H7

            If m.Msg = WM_POWERBROADCAST Then
                Dim wp As System.Int32 = m.WParam.ToInt32()
                Select Case wp
                    Case PBT_APMQUERYSUSPEND
                        ' Sofort zustimmen und NICHTS tun.
                        m.Result = New System.IntPtr(1)
                        Return

                    Case PBT_APMSUSPEND
                        System.Threading.ThreadPool.QueueUserWorkItem(
                    Sub(stateObj As System.Object)
                        Try : owner.HandlePowerSuspendAsync() : Catch : End Try
                    End Sub)
                        m.Result = New System.IntPtr(1)
                        Return

                    Case PBT_APMRESUMEAUTOMATIC
                        System.Threading.ThreadPool.QueueUserWorkItem(
                        Sub(stateObj As System.Object)
                            Try : owner.HandlePowerResumeAsync(userPresent:=False) : Catch : End Try
                        End Sub)
                        m.Result = New System.IntPtr(1)
                        Return

                    Case PBT_APMRESUMESUSPEND
                        System.Threading.ThreadPool.QueueUserWorkItem(
                        Sub(stateObj As System.Object)
                            Try : owner.HandlePowerResumeAsync(userPresent:=True) : Catch : End Try
                        End Sub)
                        m.Result = New System.IntPtr(1)
                        Return
                End Select
            End If

            MyBase.WndProc(m)
        End Sub



        Public Sub Dispose() Implements System.IDisposable.Dispose
            If Me.Handle <> System.IntPtr.Zero Then
                Me.DestroyHandle()
            End If
        End Sub
    End Class


    ' Ensure only one suspend/resume sequence runs at a time
    Private suspendResumeGate As New System.Threading.SemaphoreSlim(1, 1)

    Friend Sub HandlePowerSuspendAsync()
        System.Threading.Tasks.Task.Run(
        Async Function() As System.Threading.Tasks.Task
            If Not Await TryEnterGateAsync().ConfigureAwait(False) Then Return
            System.Threading.Interlocked.Exchange(powerChanging, 1)
            Try

                ' cancel all jobs
                For Each kv In jobMap
                    Try
                        kv.Value.Cts.Cancel()
                    Catch
                    End Try
                Next

                ' Mute watchdog during suspend
                Try : StopListenerWatchdog() : Catch : End Try

                ' Remember state
                Try
                    wasListenerRunningBeforeSleep =
                        (httpListener IsNot Nothing AndAlso httpListener.IsListening)
                Catch
                    wasListenerRunningBeforeSleep = False
                End Try
                Try
                    wasLlmThreadAliveBeforeSleep =
                        (llmThread IsNot Nothing AndAlso llmThread.IsAlive)
                Catch
                    wasLlmThreadAliveBeforeSleep = False
                End Try

                ' Proactively cancel any in-flight LLM op (prevents wake-up on dead STA)
                Try
                    If llmOperationCts IsNot Nothing Then
                        If Not llmOperationCts.IsCancellationRequested Then llmOperationCts.Cancel()
                        llmOperationCts.Dispose()
                    End If
                Catch
                Finally
                    llmOperationCts = Nothing
                End Try

                ' Force any stale listener loop to exit quickly
                System.Threading.Interlocked.Increment(listenerGeneration)

                ' Listener stoppen – OHNE UI-STA zu stoppen
                Try
                    Dim t As System.Threading.Tasks.Task = ShutdownHttpListener(stopUiThread:=False)
                    Await System.Threading.Tasks.Task.WhenAny(
                        t,
                        System.Threading.Tasks.Task.Delay(1000)
                    ).ConfigureAwait(False)
                Catch
                End Try

                ' LLM-STA: request exit without waiting (no Join while suspending)
                Try
                    If wasLlmThreadAliveBeforeSleep Then
                        StopLlmUiThreadNonBlocking()
                    End If
                Catch
                End Try

            Finally
                suspendResumeGate.Release()
            End Try
        End Function)
    End Sub


    Friend Sub HandlePowerResumeAsync(userPresent As Boolean)
        System.Threading.Tasks.Task.Run(
    Async Function() As System.Threading.Tasks.Task
        If Not Await TryEnterGateAsync().ConfigureAwait(False) Then Return
        Try
            Await System.Threading.Tasks.Task.Delay(1500).ConfigureAwait(False)

            isShuttingDown = False

            ' Tear down any stale listener
            Try
                If httpListener IsNot Nothing Then
                    Try
                        If httpListener.IsListening Then httpListener.Stop()
                    Catch
                    End Try
                    Try : httpListener.Abort() : Catch : End Try
                    Try : httpListener.Close() : Catch : End Try
                End If
            Catch
            Finally
                httpListener = Nothing
            End Try

            ' Only start listener once the user is present
            If userPresent AndAlso wasListenerRunningBeforeSleep Then
                Try
                    Await System.Threading.Tasks.Task.Delay(3000).ConfigureAwait(False)
                    StartupHttpListener()
                Catch
                End Try
            End If

            ' Do NOT recreate the legacy STA LLM thread anymore
            ' If you leave this field in place it stays unused by design.
            'If wasLlmThreadAliveBeforeSleep Then
            '    Try : EnsureLlmUiThread() : Catch : End Try
            'End If

            Try : StartListenerWatchdog() : Catch : End Try

            ' Longer post-resume cooldown
            resumeCooldownUntilUtc = System.DateTime.UtcNow.AddSeconds(60)

            Await SwitchToUi(Sub() EnableOleFilterFor(60000)).ConfigureAwait(False)

        Finally
            System.Threading.Interlocked.Exchange(powerChanging, 0)
            suspendResumeGate.Release()
        End Try
    End Function)
    End Sub

    Friend Sub oldHandlePowerResumeAsync()
        System.Threading.Tasks.Task.Run(
    Async Function() As System.Threading.Tasks.Task
        If Not Await TryEnterGateAsync().ConfigureAwait(False) Then Return
        Try
            ' Small cushion; never Thread.Sleep on resume paths
            Await System.Threading.Tasks.Task.Delay(1500).ConfigureAwait(False)

            ' Allow listener loop to run again
            isShuttingDown = False

            ' Hard-release any stale listener quickly
            Try
                If httpListener IsNot Nothing Then
                    Try
                        If httpListener.IsListening Then httpListener.Stop()
                    Catch
                    End Try
                    Try : httpListener.Abort() : Catch : End Try
                    Try : httpListener.Close() : Catch : End Try
                End If
            Catch
            Finally
                httpListener = Nothing
            End Try

            If wasListenerRunningBeforeSleep Then
                Try
                    Await System.Threading.Tasks.Task.Delay(3000).ConfigureAwait(False)
                    StartupHttpListener()
                Catch
                End Try
            End If

            If wasLlmThreadAliveBeforeSleep Then
                Try : EnsureLlmUiThread() : Catch : End Try
            End If

            ' Re-enable watchdog after we’ve (re)started things
            Try : StartListenerWatchdog() : Catch : End Try

            ' Increased cool-down window after resume
            resumeCooldownUntilUtc = System.DateTime.UtcNow.AddSeconds(45)

            Await SwitchToUi(Sub() EnableOleFilterFor(60000)).ConfigureAwait(False)

        Finally
            System.Threading.Interlocked.Exchange(powerChanging, 0)
            suspendResumeGate.Release()
        End Try
    End Function)
    End Sub

    Private Sub StopLlmUiThreadNonBlocking()
        Try
            If llmSyncContext IsNot Nothing Then
                llmSyncContext.Post(Sub() System.Windows.Forms.Application.ExitThread(), Nothing)
            End If
        Catch
        End Try
        ' KEIN Join hier!
        llmScheduler = Nothing
        llmSyncContext = Nothing
        llmThread = Nothing
    End Sub

    Private Async Function TryEnterGateAsync() As System.Threading.Tasks.Task(Of System.Boolean)
        Try
            Return Await suspendResumeGate.WaitAsync(100).ConfigureAwait(False)
        Catch
            Return False
        End Try
    End Function


    Private Sub StartupHttpListener()
        ' Make sure the loop can run again
        isShuttingDown = False

        Dim gen As System.Int64 = System.Threading.Interlocked.Increment(listenerGeneration)
        cts = New System.Threading.CancellationTokenSource()

        ' ← Add this log (generation + UTC timestamp)
        System.Diagnostics.Debug.WriteLine(
        "HttpListener START gen=" &
        gen.ToString(System.Globalization.CultureInfo.InvariantCulture) &
        " at " &
        System.DateTime.UtcNow.ToString("o", System.Globalization.CultureInfo.InvariantCulture))

        lastListenerProgressUtc = System.DateTime.UtcNow
        listenerTask = StartHttpListener(cts.Token, gen)
    End Sub



    Private Async Function ShutdownHttpListener(
    Optional ByVal stopUiThread As System.Boolean = True
) As System.Threading.Tasks.Task
        isShuttingDown = True

        ' Cancel current loop
        Try
            If cts IsNot Nothing Then cts.Cancel()
        Catch
        End Try

        ' Force-break any pending GetContextAsync and clean up thoroughly
        Try
            If httpListener IsNot Nothing Then
                Try
                    If httpListener.IsListening Then httpListener.Stop()
                Catch
                End Try
                Try
                    httpListener.Abort() ' harsher than Close; reliably breaks GetContextAsync
                Catch
                End Try
                Try
                    If httpListener.Prefixes IsNot Nothing Then httpListener.Prefixes.Clear()
                Catch
                End Try
                Try
                    httpListener.Close()
                Catch
                End Try
            End If
        Catch
        Finally
            httpListener = Nothing
        End Try

        ' Await the running listener task to completion
        Try
            If listenerTask IsNot Nothing Then
                Await listenerTask.ConfigureAwait(False)
            End If
        Catch
        Finally
            listenerTask = Nothing
        End Try

        ' Dispose CTS after we've awaited its dependents
        Try
            If cts IsNot Nothing Then cts.Dispose()
        Catch
        Finally
            cts = Nothing
        End Try

        System.Diagnostics.Debug.WriteLine(
        "HttpListener STOP at " &
        System.DateTime.UtcNow.ToString("o", System.Globalization.CultureInfo.InvariantCulture))

        ' UI-STA nur stoppen, wenn gewünscht
        If stopUiThread Then
            StopLlmUiThread()
        End If
    End Function



    Private Async Function StartHttpListener(
    ByVal token As System.Threading.CancellationToken,
    ByVal gen As System.Int64) _
    As System.Threading.Tasks.Task

        Const prefix As System.String = "http://127.0.0.1:12333/"
        Dim consecutiveFailures As System.Int32 = 0
        Dim lastMetrics As System.DateTime = System.DateTime.UtcNow

        While (Not isShuttingDown) AndAlso (Not token.IsCancellationRequested)
            ' Bail out if a newer generation has started
            If gen <> listenerGeneration Then Return

            Dim needShortDelay As System.Boolean = False

            Try
                ' Inside StartHttpListener, where the HttpListener is created
                If httpListener Is Nothing Then
                    httpListener = New System.Net.HttpListener()
                    httpListener.IgnoreWriteExceptions = True
                    With httpListener.TimeoutManager
                        ' More permissive settings to accommodate slow models
                        .IdleConnection = System.TimeSpan.FromMinutes(10)
                        .HeaderWait = System.TimeSpan.FromSeconds(30)
                        .EntityBody = System.TimeSpan.FromMinutes(10)
                        .DrainEntityBody = System.TimeSpan.FromSeconds(30)
                        ' Disable minimum send rate; we send the body in one shot
                        .MinSendBytesPerSecond = CType(0UL, System.UInt64)
                    End With
                    If Not httpListener.Prefixes.Contains(prefix) Then
                        httpListener.Prefixes.Add(prefix)
                    End If
                    httpListener.Start()
                    System.Diagnostics.Debug.WriteLine("HttpListener started.")
                ElseIf Not httpListener.IsListening Then
                    Try : httpListener.Close() : Catch : End Try
                    httpListener = Nothing
                    Continue While
                End If

                Dim ctx As System.Net.HttpListenerContext =
                Await httpListener.GetContextAsync().ConfigureAwait(False)

                ' Progress heartbeat for watchdog
                lastListenerProgressUtc = System.DateTime.UtcNow

                Dim ctxLocal As System.Net.HttpListenerContext = ctx
                System.Threading.Tasks.Task.Run(
                Async Function()
                    Dim resLocal As System.Net.HttpListenerResponse = Nothing
                    Try
                        Await HandleHttpRequest(ctxLocal).ConfigureAwait(False)
                    Catch
                        Try
                            resLocal = ctxLocal.Response
                            resLocal.StatusCode = 500
                            resLocal.KeepAlive = False
                            resLocal.Headers("Connection") = "close"
                            resLocal.SendChunked = False
                            Dim bufErr() As System.Byte = System.Text.Encoding.UTF8.GetBytes("Internal server error.")
                            resLocal.ContentType = "text/plain; charset=utf-8"
                            resLocal.ContentLength64 = bufErr.LongLength
                            Using os As System.IO.Stream = resLocal.OutputStream
                                os.Write(bufErr, 0, bufErr.Length)
                            End Using
                        Catch
                        Finally
                            Try
                                If resLocal IsNot Nothing Then resLocal.Close()
                            Catch
                            End Try
                        End Try
                    Finally
                        ' Mark progress at the end of a handled request too
                        lastListenerProgressUtc = System.DateTime.UtcNow
                    End Try
                End Function)

                ' --- metrics (unchanged) ---
                Dim now As System.DateTime = System.DateTime.UtcNow
                If (now - lastMetrics).TotalSeconds >= 10.0 Then
                    Dim gdi As System.UInt32 = GetGdiCount()
                    Dim usr As System.UInt32 = GetUserCount()
                    System.Diagnostics.Debug.WriteLine(
                    System.String.Format(
                        System.Globalization.CultureInfo.InvariantCulture,
                        "RES {0:HH:mm:ss}: GDI={1}  USER={2}",
                        now, gdi, usr))
                    If gdi >= 8000UI OrElse usr >= 8000UI Then
                        System.Diagnostics.Debug.WriteLine("WARN: Hohe Handle-Zahl – prüfe GDI/USER-Leaks.")
                    End If
                    lastMetrics = now
                End If

                consecutiveFailures = 0

            Catch ex As System.ObjectDisposedException
                consecutiveFailures += 1
                needShortDelay = True

            Catch ex As System.Exception
                consecutiveFailures += 1
                System.Diagnostics.Debug.WriteLine(System.String.Concat("Listener error: ", ex.Message))
            End Try

            If needShortDelay AndAlso (Not token.IsCancellationRequested) Then
                Try
                    Await System.Threading.Tasks.Task.Delay(50, token).ConfigureAwait(False)
                Catch
                End Try
            End If

            If consecutiveFailures >= 10 AndAlso (Not isShuttingDown) AndAlso (Not token.IsCancellationRequested) Then
                System.Diagnostics.Debug.WriteLine("Restarting HttpListener after 10 failures.")
                Try
                    If httpListener IsNot Nothing Then
                        Try : httpListener.Abort() : Catch : End Try
                        Try : httpListener.Close() : Catch : End Try
                    End If
                Catch
                Finally
                    httpListener = Nothing
                End Try
                consecutiveFailures = 0
                Try
                    Await System.Threading.Tasks.Task.Delay(5000, token).ConfigureAwait(False)
                Catch
                End Try
            End If
        End While
    End Function


    Private Sub StartPowerWatch()
        If powerWindow Is Nothing Then
            powerWindow = New PowerNotificationWindow(Me)
        End If
    End Sub

    Private Sub StopPowerWatch()
        If powerWindow IsNot Nothing Then
            powerWindow.Dispose()
            powerWindow = Nothing
        End If
    End Sub



    Private Sub OnPowerModeChanged(ByVal sender As System.Object,
                               ByVal e As Microsoft.Win32.PowerModeChangedEventArgs)
        If e Is Nothing Then Return

        Select Case e.Mode
            Case Microsoft.Win32.PowerModes.Suspend
                ' Graceful listener stop in the background
                System.Threading.ThreadPool.QueueUserWorkItem(
                Sub(state As Object)
                    Try : ShutdownHttpListener().GetAwaiter().GetResult() : Catch : End Try
                End Sub)

            Case Microsoft.Win32.PowerModes.Resume
                ' Avoid re-entrancy; delegate to unified resume path (userPresent:=True)
                If System.Threading.Interlocked.Exchange(restartingAfterResume, 1) = 1 Then Return
                System.Threading.ThreadPool.QueueUserWorkItem(
                Sub(state As Object)
                    Try
                        HandlePowerResumeAsync(userPresent:=True)
                    Finally
                        System.Threading.Interlocked.Exchange(restartingAfterResume, 0)
                    End Try
                End Sub)
        End Select
    End Sub


    Private Sub StartListenerWatchdog()
        If watchdog IsNot Nothing Then Return

        watchdog = New System.Threading.Timer(
    Sub(stateObj As System.Object)
        Try
            ' Skip while suspend/resume is in progress
            If System.Threading.Interlocked.CompareExchange(powerChanging, 0, 0) <> 0 Then Return

            ' do not kill the listener if a request is currently being processed
            Dim inFlight As Integer = Threading.Interlocked.CompareExchange(activeRequests, 0, 0)
            Dim jobsInFlight As Integer = Threading.Interlocked.CompareExchange(activeJobs, 0, 0)
            If inFlight > 0 OrElse jobsInFlight > 0 Then Return

            Dim age As System.Double =
                (System.DateTime.UtcNow - lastListenerProgressUtc).TotalSeconds

            ' Be more lenient to avoid false positives with slow models
            If age > ModelTimeout + 15 AndAlso httpListener IsNot Nothing Then
                If Not isShuttingDown Then
                    Try : httpListener.Abort() : Catch : End Try
                    Try : httpListener.Close() : Catch : End Try
                    httpListener = Nothing

                    ' Only restart if our CTS is alive
                    If cts IsNot Nothing AndAlso Not cts.IsCancellationRequested Then
                        StartupHttpListener()
                    End If
                End If
            End If
        Catch
        End Try
    End Sub,
    state:=Nothing,
    dueTime:=System.TimeSpan.FromSeconds(20),
    period:=System.TimeSpan.FromSeconds(5))
    End Sub

    Private Sub StopListenerWatchdog()
        Try
            If watchdog IsNot Nothing Then
                watchdog.Dispose()
            End If
        Catch
        Finally
            watchdog = Nothing
        End Try
    End Sub


    ' ---------------------------------------------------------------------------

    Private Const InkyBasePath As String = "/inky"
    Private Const InkyUiRoute As String = "/inky"          ' GET → serves HTML
    Private Const InkyApiRoute As String = "/inky/api"      ' POST (JSON) → commands
    Private Const InkyName As String = "Inky"               ' Fallback; AN6 preferred

    Private activeChatId As Integer = 1   ' 1 or 2 – in‑memory only (not persisted)

    Private Class LlmJob
        Public Property Id As String
        Public Property CreatedUtc As DateTime
        Public Property Tcs As TaskCompletionSource(Of String)
        Public Property Cts As CancellationTokenSource
        Public Property UseSecond As Boolean
        Public Property FileObject As String
    End Class

    Private ReadOnly jobMap As New System.Collections.Concurrent.ConcurrentDictionary(Of String, LlmJob)()
    Private activeJobs As Integer = 0
    Private Const JobTtlMinutes As Integer = 45



    Private Async Function HandleHttpRequest(
    ByVal ctx As System.Net.HttpListenerContext
) As System.Threading.Tasks.Task

        Dim req As System.Net.HttpListenerRequest = ctx.Request
        Dim res As System.Net.HttpListenerResponse = ctx.Response

        ' Count in-flight requests
        System.Threading.Interlocked.Increment(activeRequests)

        ' Heartbeat to keep watchdog calm during long processing
        Dim hb As System.Threading.Timer = Nothing

        Try
            hb = New System.Threading.Timer(
                    Sub(stateObj As System.Object)
                        Try
                            lastListenerProgressUtc = System.DateTime.UtcNow
                        Catch
                        End Try
                    End Sub,
                    state:=Nothing,
                    dueTime:=System.TimeSpan.FromSeconds(5),
                    period:=System.TimeSpan.FromSeconds(5))

            If System.Threading.Interlocked.CompareExchange(powerChanging, 0, 0) <> 0 Then
                Try
                    res = ctx.Response
                    res.StatusCode = 503
                    res.StatusDescription = "Service Unavailable (suspend/resume)"
                    res.AddHeader("Retry-After", "2")
                    res.KeepAlive = False
                    res.Headers("Connection") = "close"
                    res.SendChunked = False
                    Using os = res.OutputStream
                        Dim msgBytes() As System.Byte = System.Text.Encoding.UTF8.GetBytes("Temporarily unavailable during power transition.")
                        res.ContentType = "text/plain; charset=utf-8"
                        res.ContentLength64 = msgBytes.LongLength
                        os.Write(msgBytes, 0, msgBytes.Length)
                    End Using
                    res.Close()
                Catch
                End Try
                Return
            End If

            If IsInResumeCooldown() Then
                Try
                    res = ctx.Response
                    res.StatusCode = 503
                    res.StatusDescription = "Service Unavailable (resume cooldown)"
                    res.AddHeader("Retry-After", "5")
                    res.AddHeader("Access-Control-Allow-Origin", "*")
                    res.KeepAlive = False
                    res.Headers("Connection") = "close"
                    res.SendChunked = False
                    Using os = res.OutputStream
                        Dim msgBytes() As Byte = System.Text.Encoding.UTF8.GetBytes("Resuming from sleep; please retry in a few seconds.")
                        res.ContentType = "text/plain; charset=utf-8"
                        res.ContentLength64 = msgBytes.LongLength
                        os.Write(msgBytes, 0, msgBytes.Length)
                    End Using
                    res.Close()
                Catch
                End Try
                Return
            End If

            ' ---- CORS Preflight ---------------------------------------------------
            If req.HttpMethod.Equals("OPTIONS", System.StringComparison.OrdinalIgnoreCase) Then
                res.AddHeader("Access-Control-Allow-Origin", "*")
                res.AddHeader("Access-Control-Allow-Methods", "GET, POST, PUT, DELETE, OPTIONS")
                res.AddHeader("Access-Control-Allow-Headers", "Content-Type, Authorization")
                res.StatusCode = 204
                res.KeepAlive = False
                res.Headers("Connection") = "close"
                res.SendChunked = False
                res.Close()
                Return
            End If

            ' ---- Favicon ----------------------------------------------------------
            If req.HttpMethod.Equals("GET", System.StringComparison.OrdinalIgnoreCase) AndAlso
           req.RawUrl.Equals("/favicon.ico", System.StringComparison.OrdinalIgnoreCase) Then

                Dim png() As System.Byte = GetLogoPngBytes()

                res.ContentType = "image/png"
                res.AddHeader("Cache-Control", "public, max-age=86400")
                res.KeepAlive = False
                res.Headers("Connection") = "close"
                res.SendChunked = False
                res.ContentLength64 = png.LongLength

                Using os As System.IO.Stream = res.OutputStream
                    Await os.WriteAsync(png, 0, png.Length).ConfigureAwait(False)
                End Using
                res.Close()
                Return
            End If

            ' ---- Inky UI (GET /inky[/]) ------------------------------------------
            If req.HttpMethod.Equals("GET", System.StringComparison.OrdinalIgnoreCase) AndAlso
           (req.RawUrl.Equals(InkyUiRoute, System.StringComparison.OrdinalIgnoreCase) OrElse
            req.RawUrl.Equals(InkyUiRoute & "/", System.StringComparison.OrdinalIgnoreCase)) Then

                Dim html As System.String = BuildInkyHtmlPage()
                Dim bufUi() As System.Byte = System.Text.Encoding.UTF8.GetBytes(html)

                res.ContentType = "text/html; charset=utf-8"
                res.AddHeader("Cache-Control", "no-store")
                res.KeepAlive = False
                res.Headers("Connection") = "close"
                res.SendChunked = False
                res.ContentLength64 = bufUi.LongLength

                Using os As System.IO.Stream = res.OutputStream
                    Await os.WriteAsync(bufUi, 0, bufUi.Length).ConfigureAwait(False)
                End Using
                res.Close()
                Return
            End If

            ' ---- Normal flow (POST JSON / API) -----------------------------------
            Dim body As System.String = System.String.Empty
            If req.HasEntityBody Then
                Using rdr As New System.IO.StreamReader(req.InputStream, System.Text.Encoding.UTF8, detectEncodingFromByteOrderMarks:=False, bufferSize:=8192, leaveOpen:=False)
                    body = Await rdr.ReadToEndAsync().ConfigureAwait(False)
                End Using
            End If

            Dim responseText As System.String = Await ProcessRequestInAddIn(body, req.RawUrl).ConfigureAwait(False)
            If responseText Is Nothing Then responseText = System.String.Empty

            ' Content-Type Hints
            Dim contentType As System.String = "text/plain; charset=utf-8"
            If responseText.StartsWith("CT:html" & vbLf, System.StringComparison.Ordinal) Then
                contentType = "text/html; charset=utf-8"
                responseText = responseText.Substring(("CT:html" & vbLf).Length)
            ElseIf responseText.StartsWith("CT:json" & vbLf, System.StringComparison.Ordinal) Then
                contentType = "application/json; charset=utf-8"
                responseText = responseText.Substring(("CT:json" & vbLf).Length)
            End If

            Dim buf() As System.Byte = System.Text.Encoding.UTF8.GetBytes(responseText)

            res.AddHeader("Access-Control-Allow-Origin", "*")
            res.ContentType = contentType
            res.KeepAlive = False
            res.Headers("Connection") = "close"
            res.SendChunked = False
            res.ContentLength64 = buf.LongLength

            Using os As System.IO.Stream = res.OutputStream
                Await os.WriteAsync(buf, 0, buf.Length).ConfigureAwait(False)
            End Using
            res.Close()

        Catch ex As System.Exception
            Try
                Dim err As System.String = "Internal server error: " & ex.Message
                Dim bufErr() As System.Byte = System.Text.Encoding.UTF8.GetBytes(err)

                res.StatusCode = 500
                res.AddHeader("Access-Control-Allow-Origin", "*")
                res.ContentType = "text/plain; charset=utf-8"
                res.KeepAlive = False
                res.Headers("Connection") = "close"
                res.SendChunked = False
                res.ContentLength64 = bufErr.LongLength
                Using os As System.IO.Stream = res.OutputStream
                    os.Write(bufErr, 0, bufErr.Length)
                End Using
                res.Close()
            Catch
            End Try
        Finally
            Try
                If hb IsNot Nothing Then hb.Dispose()
            Catch
            End Try
            System.Threading.Interlocked.Decrement(activeRequests)
            ' Mark progress at the end of a handled request too
            lastListenerProgressUtc = System.DateTime.UtcNow
        End Try
    End Function




    Private Function GetLogoPngBytes() As System.Byte()
        Try
            Using src As System.Drawing.Bitmap = CType(My.Resources.Red_Ink_Logo.Clone(), System.Drawing.Bitmap)
                Using ms As New System.IO.MemoryStream()
                    src.Save(ms, System.Drawing.Imaging.ImageFormat.Png)
                    Return ms.ToArray()
                End Using
            End Using
        Catch
            ' 1x1 transparent PNG fallback
            Return System.Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVQImWNgYGD4DwABdQF+8m3rXQAAAABJRU5ErkJggg==")
        End Try
    End Function


    ' --------------- LLM helper (runs off the UI thread) -----------------------


    ' ----------------------------------------
    ' 1) Feld für den Scheduler (Klassen-/Modul-Ebene)
    ' ----------------------------------------
    Private Shared llmScheduler As System.Threading.Tasks.TaskScheduler

    ' ----------------------------------------
    ' 2) STA-Thread mit eigener WinForms-Message-Loop initialisieren
    ' ----------------------------------------

    Private Sub EnsureLlmUiThread()
        If llmThread IsNot Nothing AndAlso llmThread.IsAlive Then Return
        If llmScheduler IsNot Nothing Then Return

        Dim tcs As New System.Threading.Tasks.TaskCompletionSource(Of System.Threading.Tasks.TaskScheduler)()

        llmThread = New System.Threading.Thread(
        Sub()
            System.Threading.SynchronizationContext.SetSynchronizationContext(
                New System.Windows.Forms.WindowsFormsSynchronizationContext())

            llmSyncContext = System.Threading.SynchronizationContext.Current
            tcs.SetResult(System.Threading.Tasks.TaskScheduler.FromCurrentSynchronizationContext())

            System.Windows.Forms.Application.Run()
        End Sub)

        llmThread.SetApartmentState(System.Threading.ApartmentState.STA)
        llmThread.IsBackground = True
        llmThread.Start()

        llmScheduler = tcs.Task.Result
    End Sub

    Private Sub StopLlmUiThread()
        Try
            If llmSyncContext IsNot Nothing Then
                llmSyncContext.Post(
                Sub() System.Windows.Forms.Application.ExitThread(),
                Nothing)
            End If
        Catch
        End Try
        Try
            If llmThread IsNot Nothing AndAlso llmThread.IsAlive Then
                If Not llmThread.Join(2000) Then
                    ' Optional: Log that the thread did not terminate in time.
                End If
            End If
        Catch
        End Try

        llmScheduler = Nothing
        llmSyncContext = Nothing
        llmThread = Nothing
    End Sub


    Public Function RunLlmAsync(
        ByVal sysPrompt As System.String,
        ByVal userPrompt As System.String,
        Optional ByVal UseSecondAPI As System.Boolean = False,
        Optional ByVal ShowTimer As System.Boolean = True,
        Optional ByVal FileObject As System.String = "",
        Optional ByVal cancellationToken As System.Threading.CancellationToken = Nothing
    ) As System.Threading.Tasks.Task(Of System.String)

        ' Use the thread pool – no STA message loop
        Dim effectiveTimeout As Integer = If(UseSecondAPI, INI_Timeout_2, INI_Timeout)
        ModelTimeout = effectiveTimeout

        Return System.Threading.Tasks.Task.Run(
            Async Function() As System.Threading.Tasks.Task(Of System.String)
                Using linkedCts As System.Threading.CancellationTokenSource =
                    System.Threading.CancellationTokenSource.CreateLinkedTokenSource(cancellationToken)

                    Try
                        Dim llmOut As String =
                            Await LLM(sysPrompt, userPrompt, "", "", 0, UseSecondAPI, Not ShowTimer, "", FileObject, linkedCts.Token).
                                ConfigureAwait(False)

                        If UseSecondAPI AndAlso originalConfigLoaded Then
                            RestoreDefaults(_context, originalConfig)
                            originalConfigLoaded = False
                        End If

                        Return If(llmOut, String.Empty)

                    Catch ex As OperationCanceledException
                        Return "Operation was canceled by the user."
                    End Try
                End Using
            End Function,
            cancellationToken)
    End Function

    Public Function oldRunLlmAsync(
    ByVal sysPrompt As System.String,
    ByVal userPrompt As System.String,
    Optional ByVal UseSecondAPI As System.Boolean = False,
    Optional ByVal ShowTimer As System.Boolean = True,
    Optional ByVal FileObject As System.String = "",
    Optional ByVal cancellationToken As System.Threading.CancellationToken = Nothing
) As System.Threading.Tasks.Task(Of System.String)

        EnsureLlmUiThread()

        Return System.Threading.Tasks.Task.Factory.StartNew(
        Function() As System.String
            Using linkedCts As System.Threading.CancellationTokenSource =
                System.Threading.CancellationTokenSource.CreateLinkedTokenSource(cancellationToken)

                If UseSecondAPI Then ModelTimeout = INI_Timeout_2 Else ModelTimeout = INI_Timeout

                Try
                    Return Nito.AsyncEx.AsyncContext.Run(
                        Async Function() As System.Threading.Tasks.Task(Of System.String)
                            Dim llmOut As System.String =
                                Await LLM(sysPrompt, userPrompt, "", "", 0, UseSecondAPI, Not ShowTimer, "", FileObject, linkedCts.Token).ConfigureAwait(True)

                            If UseSecondAPI AndAlso originalConfigLoaded Then
                                RestoreDefaults(_context, originalConfig)
                                originalConfigLoaded = False
                            End If

                            Return If(llmOut, System.String.Empty)
                        End Function)
                Catch ex As OperationCanceledException
                    Return "Operation was canceled by the user."
                End Try
            End Using
        End Function,
        cancellationToken,
        System.Threading.Tasks.TaskCreationOptions.None,
        llmScheduler)
    End Function



    ' ---------------------------------------------------------------------------

    ' --------------- Compare & insert helper (runs on UI) ----------------------
    '───────────────────────────────────────────────────────────────────────────
    ' Shows the compare window on the UI thread and BLOCKS the calling
    ' Task until the user dismisses it.  Returns True if the user accepted,
    ' False if Esc was pressed (like the original code).
    '───────────────────────────────────────────────────────────────────────────

    Private Async Function CompareAndInsertSyncConfirm(
    originalText As String,
    llmResult As String) _
    As System.Threading.Tasks.Task(Of Boolean)

        ' Show compare on UI thread
        Await SwitchToUi(Sub() CompareAndInsertText(originalText, llmResult, True)).ConfigureAwait(False)

        ' Await user decision (Esc = False, OK/close = True)
        Dim accepted As Boolean = Await WaitForPreviewDecisionAsync().ConfigureAwait(False)
        Return accepted
    End Function

    Private Function oldCompareAndInsertSyncConfirm(
        originalText As String,
        llmResult As String) _
        As System.Threading.Tasks.Task(Of Boolean)

        Dim tcs As New System.Threading.Tasks.TaskCompletionSource(Of Boolean)

        ' marshal to UI thread with BeginInvoke so listener thread never blocks
        mainThreadControl.BeginInvoke(New MethodInvoker(
        Sub()

            ' 1) show compare window (modal for this thread)
            CompareAndInsertText(originalText, llmResult, True)

            ' 2) pump one message cycle so the Esc keystroke is processed
            System.Windows.Forms.Application.DoEvents()

            ' 3) read Esc status exactly like the old code
            Dim escNow As Boolean =
                (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And &H8000) <> 0
            Dim escDown As Boolean =
                (GetAsyncKeyState(System.Windows.Forms.Keys.Escape) And 1) <> 0

            Dim accepted As Boolean = Not (escNow Or escDown)

            tcs.SetResult(accepted)      ' unblock the awaiting thread
        End Sub))

        Return tcs.Task       ' caller awaits without blocking the listener
    End Function
    ' ---------------------------------------------------------------------------


    '───────────────────────────────────────────────────────────────────────────
    ' Waits asynchronously until the preview window (ShowRTFCustomMessageBox)
    ' is either closed with OK  → returns True
    '                     or Esc is pressed       → returns False
    ' Works even though the preview window is created on its own worker thread.
    '───────────────────────────────────────────────────────────────────────────
    Private Async Function WaitForPreviewDecisionAsync() _
    As System.Threading.Tasks.Task(Of System.Boolean)

        Dim tcs As New System.Threading.Tasks.TaskCompletionSource(Of System.Boolean)()

        ' Handler-Anheftung einmalig auf dem UI-Thread
        Await SwitchToUi(Sub()

                             Dim previewForm As System.Windows.Forms.Form = Nothing
                             Dim searchTimer As New System.Windows.Forms.Timer() With {.Interval = 100}

                             AddHandler searchTimer.Tick,
                             Sub()
                                 If previewForm Is Nothing OrElse previewForm.IsDisposed Then
                                     previewForm = System.Windows.Forms.Application.OpenForms _
                                         .Cast(Of System.Windows.Forms.Form)() _
                                         .FirstOrDefault(Function(f As System.Windows.Forms.Form) f.Name = "ShowRTFCustomMessageBox" _
                                                             OrElse f.Text.StartsWith(AN))

                                     If previewForm Is Nothing Then Return

                                     previewForm.KeyPreview = True

                                     AddHandler previewForm.KeyDown,
                                         Sub(_s As System.Object, e As System.Windows.Forms.KeyEventArgs)
                                             If e.KeyCode = System.Windows.Forms.Keys.Escape Then
                                                 tcs.TrySetResult(False)
                                             End If
                                         End Sub

                                     AddHandler previewForm.FormClosed,
                                         Sub(_s As System.Object, _e As System.Windows.Forms.FormClosedEventArgs)
                                             tcs.TrySetResult(True)
                                             ' Failsafe: Timer aufräumen
                                             If searchTimer.Enabled Then
                                                 searchTimer.Stop()
                                                 searchTimer.Dispose()
                                             End If
                                         End Sub
                                 End If

                                 If tcs.Task.IsCompleted Then
                                     searchTimer.Stop()
                                     searchTimer.Dispose() ' Patch C
                                 End If
                             End Sub

                             searchTimer.Start()
                         End Sub).ConfigureAwait(False)

        ' WICHTIG: In Async-Funktion Task(Of Boolean) ==>> "Await tcs.Task"
        Return Await tcs.Task.ConfigureAwait(False)
    End Function




    ' ===== Chatbot 

    ' ---------- Safe accessors & utilities ----------
    Private Function TryGetAppSetting(Of T)(ByVal key As System.String, ByVal fallback As T) As T
        Try
            Dim p = GetType(My.MySettings).GetProperty(key, System.Reflection.BindingFlags.Public Or System.Reflection.BindingFlags.Instance)
            If p IsNot Nothing Then
                Dim v = DirectCast(p.GetValue(My.Settings, Nothing), Object)
                If v IsNot Nothing Then Return DirectCast(v, T)
            End If
        Catch
        End Try
        Return fallback
    End Function

    Private Function GetBotName() As System.String
        ' Try: My.Settings("AN6") → else "Inky"
        Dim v As System.String = TryGetAppSetting(Of System.String)("AN6", Nothing)
        If Not System.String.IsNullOrWhiteSpace(v) Then Return v
        Return "Inky"
    End Function

    Private Function GetLogoDataUrl() As System.String
        Try
            Using src As System.Drawing.Bitmap = CType(My.Resources.Red_Ink_Logo.Clone(), System.Drawing.Bitmap)
                Using ms As New System.IO.MemoryStream()
                    src.Save(ms, System.Drawing.Imaging.ImageFormat.Png)
                    Dim b64 As System.String = System.Convert.ToBase64String(ms.ToArray())
                    Return "data:image/png;base64," & b64
                End Using
            End Using
        Catch
            Return ""
        End Try
    End Function

    Private Function GetSystemPromptChat() As System.String
        ' Try: My.Settings("SP_Chat") → else fallback        
        Dim v As System.String = TryGetAppSetting(Of System.String)("SP_Chat", Nothing)
        If Not System.String.IsNullOrWhiteSpace(v) Then Return v
        Return "You are a helpful assistant."
    End Function

    ' Extract a human label/key from ModelConfig even if it lacks “Description”.
    ' --------- DROP-IN: ersetzt GetModelDisplayKey ---------
    Private Function GetModelDisplayKey(ByVal model As SharedLibrary.SharedLibrary.ModelConfig) As System.String
        If model Is Nothing Then Return ""
        ' Bevorzugt die sprechende Bezeichnung:
        If Not System.String.IsNullOrWhiteSpace(model.ModelDescription) Then
            Return model.ModelDescription
        End If
        ' Fallback: der interne Modellname
        If Not System.String.IsNullOrWhiteSpace(model.Model) Then
            Return model.Model
        End If
        Return "Model"
    End Function


    Private Function GetFriendlyGreeting() As System.String
        Dim name As System.String = GetBotName()
        Dim tl As System.String
        Try
            tl = System.Globalization.CultureInfo.CurrentUICulture.TwoLetterISOLanguageName
        Catch
            tl = "en"
        End Try

        Select Case tl
            Case "de" : Return $"Hallo! Ich bin {name}. Wie kann ich helfen?"
            Case "fr" : Return $"Salut ! Je suis {name}. Comment puis-je aider ?"
            Case "it" : Return $"Ciao! Sono {name}. In cosa posso aiutarti?"
            Case "es" : Return $"¡Hola! Soy {name}. ¿En qué puedo ayudarte?"
            Case Else : Return $"Hi! I’m {name}. How can I help?"
        End Select
    End Function

    Dim botName As System.String = GetBotName()
    Dim brandName As System.String = AN
    Dim logoUrl As System.String = GetLogoDataUrl()
    Dim greet As System.String = GetFriendlyGreeting()


    ' Simple persisted state container for chat.
    <Serializable>
    Private Class InkyState
        Public History As System.Collections.Generic.List(Of ChatTurn) = New System.Collections.Generic.List(Of ChatTurn)()
        Public SelectedModelKey As System.String = ""
        Public UseSecondApi As System.Boolean = False
        Public LastAssistantText As System.String = ""
        Public DarkMode As System.Boolean = False
        Public SupportsFileUploads As System.Boolean = False
    End Class

    <Serializable>
    Private Class ChatTurn
        Public Role As System.String   ' "user" or "assistant"
        Public Markdown As System.String
        Public Html As System.String
        Public Utc As System.DateTime
    End Class

    Private Function GetUserLanguageTwoLetter() As System.String
        Try
            Return System.Globalization.CultureInfo.CurrentUICulture.TwoLetterISOLanguageName
        Catch
            Return "en"
        End Try
    End Function


    Private Function ComputeSupportsFiles(ByVal useSecond As System.Boolean,
                                      ByVal selectedKey As System.String) As System.Boolean
        Try
            ' Primary API (kein Alternate gewählt)
            If Not useSecond Then
                Return Not System.String.IsNullOrWhiteSpace(INI_APICall_Object)
            End If

            ' Second API, Default-Modell
            If System.String.IsNullOrWhiteSpace(selectedKey) Then
                Return Not System.String.IsNullOrWhiteSpace(INI_APICall_Object_2)
            End If

            ' Second API, Alternate-Modell -> aus ModelConfig.APICall_Object lesen
            Dim alts As System.Collections.Generic.List(Of SharedLibrary.SharedLibrary.ModelConfig) = Nothing
            Try
                alts = LoadAlternativeModels(INI_AlternateModelPath, _context)
            Catch
                alts = Nothing
            End Try
            If alts Is Nothing Then Return False

            Dim sel As SharedLibrary.SharedLibrary.ModelConfig =
            alts.FirstOrDefault(Function(m As SharedLibrary.SharedLibrary.ModelConfig)
                                    If m Is Nothing Then Return False
                                    If Not System.String.IsNullOrWhiteSpace(m.ModelDescription) AndAlso
                                       System.String.Equals(m.ModelDescription, selectedKey, System.StringComparison.OrdinalIgnoreCase) Then
                                        Return True
                                    End If
                                    If Not System.String.IsNullOrWhiteSpace(m.Model) AndAlso
                                       System.String.Equals(m.Model, selectedKey, System.StringComparison.OrdinalIgnoreCase) Then
                                        Return True
                                    End If
                                    Return False
                                End Function)

            If sel Is Nothing Then Return False

            ' Direkter Zugriff – falls Property fehlt, fallback via Reflection
            Dim v As System.String = Nothing
            Try
                v = sel.APICall_Object
            Catch
                Try
                    Dim p As System.Reflection.PropertyInfo =
                    GetType(SharedLibrary.SharedLibrary.ModelConfig).GetProperty("APICall_Object",
                        System.Reflection.BindingFlags.Public Or System.Reflection.BindingFlags.Instance)
                    If p IsNot Nothing Then
                        Dim o As System.Object = p.GetValue(sel, Nothing)
                        If o IsNot Nothing Then v = System.Convert.ToString(o, System.Globalization.CultureInfo.InvariantCulture)
                    End If
                Catch
                End Try
            End Try

            Return Not System.String.IsNullOrWhiteSpace(v)
        Catch
            Return False
        End Try
    End Function



    Private Function LoadInkyState(Optional chatId As Integer = -1) As InkyState
        If chatId = -1 Then chatId = activeChatId
        Dim settingKey As String = If(chatId = 2, "ChatHistory_Inky2", "ChatHistory_Inky")
        Try
            Dim raw As String = ""
            Try : raw = DirectCast(My.Settings.[GetType]().GetProperty(settingKey).GetValue(My.Settings, Nothing), String) : Catch : raw = "" : End Try
            If String.IsNullOrWhiteSpace(raw) Then
                Dim st As New InkyState()
                ' Default dark mode on first empty chat
                st.DarkMode = True
                Return st
            End If
            Dim stLoaded = Newtonsoft.Json.JsonConvert.DeserializeObject(Of InkyState)(raw)
            If stLoaded Is Nothing Then stLoaded = New InkyState()
            Return stLoaded
        Catch
            Dim st As New InkyState() : st.DarkMode = True
            Return st
        End Try
    End Function

    Private Sub SaveInkyState(st As InkyState, Optional chatId As Integer = -1)
        If chatId = -1 Then chatId = activeChatId
        Dim settingKey As String = If(chatId = 2, "ChatHistory_Inky2", "ChatHistory_Inky")
        Try
            Dim json = Newtonsoft.Json.JsonConvert.SerializeObject(st)
            Try
                My.Settings.[GetType]().GetProperty(settingKey).SetValue(My.Settings, json, Nothing)
                My.Settings.Save()
            Catch
                ' ignore
            End Try
        Catch
        End Try
    End Sub

    Private Function MarkdownToHtml(ByVal md As System.String) As System.String
        Try
            ' Maximale Markdown-Funktionalität + SoftlineBreaks als <br/>
            Dim pipeline As Markdig.MarkdownPipeline =
            New Markdig.MarkdownPipelineBuilder().
                UseAdvancedExtensions().
                UseSoftlineBreakAsHardlineBreak().
                UsePipeTables().
                UseGridTables().
                UseListExtras().
                UseFootnotes().
                UseDefinitionLists().
                UseAbbreviations().
                UseAutoLinks().
                UseTaskLists().
                UseEmojiAndSmiley().
                UseMathematics().
                UseFigures().
                UseGenericAttributes().
                Build()

            Return Markdig.Markdown.ToHtml(md, pipeline)
        Catch ex As System.Exception
            ' Fallback: sicher encoden UND Zeilenumbrüche erhalten
            Return System.Net.WebUtility.HtmlEncode(md).Replace(vbLf, "<br>")
        End Try
    End Function

    Private Function CapHistoryToChars(ByVal st As InkyState, ByVal maxChars As System.Int32) As System.Collections.Generic.List(Of ChatTurn)
        If maxChars <= 0 Then Return New System.Collections.Generic.List(Of ChatTurn)(st.History)
        Dim acc As New System.Text.StringBuilder()
        Dim clipped As New System.Collections.Generic.List(Of ChatTurn)()
        ' iterate from the end (most recent) backwards until cap reached
        For i As System.Int32 = st.History.Count - 1 To 0 Step -1
            Dim turn As ChatTurn = st.History(i)
            Dim piece As System.String = $"[{turn.Role}]{turn.Markdown}" & vbLf
            If acc.Length + piece.Length > maxChars Then Exit For
            acc.Insert(0, piece)
            clipped.Insert(0, turn)
        Next
        Return clipped
    End Function

    Private Function GetSelectedModelLabel(ByVal useSecond As System.Boolean, ByVal selectedKey As System.String) As System.String
        If Not useSecond Then
            Return If(System.String.IsNullOrWhiteSpace(INI_Model), "Default model", INI_Model)
        End If
        If Not System.String.IsNullOrWhiteSpace(selectedKey) Then
            Return selectedKey
        End If
        Return If(System.String.IsNullOrWhiteSpace(INI_Model_2), "Second API model", INI_Model_2)
    End Function

    ' Builds the entire HTML UI (single file; no external assets)

    Private Function BuildInkyHtmlPage() As System.String
        Dim botName As String = GetBotName()
        Dim brandName As String = If(Not String.IsNullOrWhiteSpace(AN), AN, botName)
        Dim logoUrl As String = GetLogoDataUrl()
        Dim greet As String = GetFriendlyGreeting()

        Dim html As New System.Text.StringBuilder()

        html.AppendLine("<!doctype html>")
        html.AppendLine("<html lang=""en""><head><meta charset=""utf-8"">")
        html.AppendLine("<meta name=""viewport"" content=""width=device-width, initial-scale=1"">")
        html.AppendLine("<link rel=""shortcut icon"" type=""image/png"" href=""" & System.Net.WebUtility.HtmlEncode(logoUrl) & """>")
        html.AppendLine("<link rel=""icon"" type=""image/png"" href=""" & System.Net.WebUtility.HtmlEncode(logoUrl) & """>")
        html.AppendLine("<title>" & System.Net.WebUtility.HtmlEncode(brandName) & " — Local Chat</title>")

        ' CSS
        html.AppendLine("<style>")
        html.AppendLine(":root{--bg:#0b0f14;--card:#11161d;--fg:#e8eef6;--muted:#9aa8b7;--border:#1b2430;--border-strong:#2d3744;--elev:#1a222c;--press-shadow:inset 0 2px 6px rgba(0,0,0,.45);}")
        html.AppendLine(":root.light{--bg:#f6f7f9;--card:#ffffff;--fg:#0e1116;--muted:#5d6a77;--border:#e2e5e9;--border-strong:#c9cfd6;--elev:#eef1f4;--press-shadow:inset 0 2px 5px rgba(0,0,0,.08);}")
        html.AppendLine("html,body{height:100%;margin:0;font-family:system-ui,Segoe UI,Roboto,Arial,sans-serif;background:var(--bg);color:var(--fg);}")
        html.AppendLine(".wrap{display:flex;flex-direction:column;height:100%;}")
        html.AppendLine(".topbar{display:flex;gap:.5rem;align-items:center;padding:.75rem 1rem;border-bottom:1px solid var(--border);background:var(--card);position:sticky;top:0;z-index:5;flex-wrap:nowrap;overflow:hidden;}")
        html.AppendLine(".topline{display:flex;align-items:center;gap:.6rem;min-width:0;}")
        html.AppendLine(".topline img.logo{width:24px;height:24px;border-radius:6px;display:block}")
        html.AppendLine(".topline .brandbig{font-weight:700;white-space:nowrap;}")
        html.AppendLine(".topline .sub{color:var(--muted);font-size:.9rem;white-space:nowrap;}")
        html.AppendLine(".muted{color:var(--muted);font-size:.85rem}")
        html.AppendLine(".spacer{flex:0 0 0;}")
        html.AppendLine("select,button,input,textarea{background:var(--card);color:var(--fg);border:1px solid var(--border);border-radius:.6rem;font:inherit;}")
        html.AppendLine("select,button,input{padding:.5rem .7rem;}")
        html.AppendLine("button{cursor:pointer;transition:background .16s,filter .12s,transform .08s,box-shadow .18s;}")
        html.AppendLine("button:hover{filter:brightness(1.07)}")
        html.AppendLine("button:disabled{opacity:.5;cursor:not-allowed}")
        html.AppendLine("button.is-pressed, .chatTab.is-pressed{transform:translateY(1px);box-shadow:var(--press-shadow);filter:brightness(.92);}")
        html.AppendLine("button:active:not(:disabled){transform:translateY(1px);box-shadow:var(--press-shadow);filter:brightness(.9);}")
        html.AppendLine(".chat{flex:1;overflow:auto;padding:1rem;}")
        html.AppendLine(".row{display:flex;margin:0 auto 1rem auto;max-width:1000px;padding:0 .25rem;}")
        html.AppendLine(".row.bot{justify-content:flex-start}")
        html.AppendLine(".row.user{justify-content:flex-end}")
        html.AppendLine(".bubble{max-width:75%;padding:1rem;border:1px solid var(--border);background:var(--card);border-radius:1rem;box-shadow:0 1px 3px rgba(0,0,0,.25)}")
        html.AppendLine(".bot .bubble{border-top-right-radius:.35rem}")
        html.AppendLine(".user .bubble{border-top-left-radius:.35rem}")
        html.AppendLine(".role{font-size:.75rem;color:var(--muted);margin-bottom:.25rem}")
        html.AppendLine(".inputbar{display:flex;gap:.5rem;padding:1rem;border-top:1px solid var(--border);background:var(--card)}")
        html.AppendLine("textarea{flex:1;resize:vertical;min-height:52px;max-height:220px;border-radius:.8rem;padding:.75rem;line-height:1.25;}")
        html.AppendLine(".hint{font-size:.7rem;letter-spacing:.3px;color:var(--muted);padding:.25rem 1rem 1rem}")
        html.AppendLine("a{color:inherit;text-decoration:underline;text-decoration-color:rgba(255,255,255,.35)}")
        html.AppendLine(":root.light a{text-decoration-color:rgba(0,0,0,.4)}")
        html.AppendLine("a:hover{filter:brightness(1.15)}")
        html.AppendLine("code,pre{font-family:ui-monospace,Consolas,monospace;font-size:.85rem}")
        html.AppendLine("pre{overflow:auto;padding:.75rem;border:1px solid var(--border);border-radius:.6rem;position:relative;background:var(--elev);}")
        html.AppendLine(".code-copy-btn{position:absolute;top:6px;right:6px;padding:4px 8px;font-size:.65rem;line-height:1;border:1px solid var(--border);border-radius:4px;background:rgba(0,0,0,.45);backdrop-filter:blur(3px);cursor:pointer;display:flex;align-items:center;gap:6px;color:var(--fg);opacity:0;transition:opacity .18s,background .18s;}")
        html.AppendLine("pre:hover .code-copy-btn{opacity:1}")
        html.AppendLine(".code-copy-btn svg{width:16px;height:16px;display:block}")
        html.AppendLine(".code-copy-btn.copied{background:#2c3440;color:#fff}")
        html.AppendLine(":root.light .code-copy-btn.copied{background:#d5d9dd;color:#111}")
        html.AppendLine(".code-copy-btn:focus{outline:2px solid var(--border-strong);}")
        html.AppendLine(".chatTab{padding:.45rem .55rem;min-width:32px;font-size:.7rem;font-weight:600;line-height:1;border:1px solid var(--border);background:var(--card);color:var(--muted);transition:background .18s,border-color .18s,color .18s,transform .08s,box-shadow .18s;flex-shrink:0;}")
        html.AppendLine(".chatTab:hover:not(:disabled){background:var(--elev);color:var(--fg);}")
        html.AppendLine(".chatTab.active{background:#222b35;border-color:var(--border-strong);color:#fff;box-shadow:inset 0 0 0 1px #303c46;}")
        html.AppendLine(":root.light .chatTab.active{background:#e2e5e9;border-color:var(--border-strong);color:#0e1116;box-shadow:inset 0 0 0 1px #c9cfd6;}")
        html.AppendLine(".chatTab:focus{outline:2px solid var(--border-strong);outline-offset:1px;}")
        ' Typing + elapsed
        html.AppendLine(".typing-dots{display:inline-flex;gap:6px;align-items:center;}")
        html.AppendLine(".typing-dots span{width:7px;height:7px;border-radius:50%;background:currentColor;opacity:.35;animation:tdots 1.2s infinite ease-in-out;}")
        html.AppendLine(".typing-dots span:nth-child(2){animation-delay:.2s}")
        html.AppendLine(".typing-dots span:nth-child(3){animation-delay:.4s}")
        html.AppendLine("@keyframes tdots{0%,80%,100%{transform:translateY(0);opacity:.3}40%{transform:translateY(-5px);opacity:.85}}")
        html.AppendLine(".typing-elapsed{margin-left:8px;font-size:.65rem;color:var(--muted);font-family:ui-monospace,monospace;opacity:.8;}")
        html.AppendLine(".actions{display:flex;flex-direction:row;gap:.5rem;align-items:stretch;}")
        html.AppendLine(".actions .stack{display:flex;flex-direction:column;gap:.5rem;}")
        html.AppendLine("#cancelBtn{display:none;align-self:stretch;height:auto;}")
        ' Responsive model select behaviour
        html.AppendLine("#modelSel{flex:1 1 260px;max-width:420px;min-width:110px;overflow:hidden;white-space:nowrap;text-overflow:ellipsis;box-sizing:border-box;}")
        html.AppendLine("#modelSel.squeezed{max-width:55vw;}")
        html.AppendLine(".topbar button,#modelSel{flex-shrink:0;}")
        html.AppendLine("@media (max-width:1000px){#modelSel{max-width:360px;}}")
        html.AppendLine("@media (max-width:880px){.topline .sub{display:none;}}")
        html.AppendLine("@media (max-width:760px){.topline .brandbig{max-width:140px;overflow:hidden;text-overflow:ellipsis;}#modelSel{max-width:300px;}}")
        html.AppendLine("@media (max-width:640px){#modelSel{max-width:55vw;} .topline .brandbig{max-width:110px;}}")
        html.AppendLine("@media (max-width:640px){.actions{flex-direction:column;}.actions .stack{flex-direction:row;}.actions .stack button{flex:1;}#cancelBtn{align-self:auto;height:auto;}}")
        html.AppendLine("</style>")

        html.AppendLine("</head><body>")
        html.AppendLine("<div class=""wrap"">")

        ' Top bar
        html.AppendLine("  <div class=""topbar"">")
        html.AppendLine("    <div class=""topline"">")
        If Not String.IsNullOrWhiteSpace(logoUrl) Then
            html.AppendLine("      <img class=""logo"" src=""" & System.Net.WebUtility.HtmlEncode(logoUrl) & """ alt=""logo"">")
        End If
        html.AppendLine("      <div class=""brandbig"">" & System.Net.WebUtility.HtmlEncode(brandName) & "</div>")
        html.AppendLine("      <div class=""sub"">Local Chat</div>")
        html.AppendLine("    </div>")
        html.AppendLine("    <div class=""spacer""></div>")
        html.AppendLine("    <select id=""modelSel"" title=""Model""></select>")
        html.AppendLine("    <button id=""copyBtn"" title=""Copy last answer to clipboard"">Copy last</button>")
        html.AppendLine("    <button id=""clearBtn"" title=""Clear current conversation"">Clear</button>")
        html.AppendLine("    <button id=""chat1Btn"" class=""chatTab"" data-chat=""1"" title=""Chat 1"">1</button>")
        html.AppendLine("    <button id=""chat2Btn"" class=""chatTab"" data-chat=""2"" title=""Chat 2"">2</button>")
        html.AppendLine("    <button id=""themeBtn"" title=""Toggle theme"">Theme</button>")
        html.AppendLine("  </div>")

        html.AppendLine("  <div id=""chat"" class=""chat""></div>")

        html.AppendLine("  <div class=""inputbar"">")
        html.AppendLine("    <textarea id=""msg"" placeholder=""" & System.Net.WebUtility.HtmlEncode(greet) & """ autofocus></textarea>")
        html.AppendLine("    <div class=""actions"">" &
                            "<div class=""stack"">" &
                                "<button id=""sendBtn"">Send</button>" &
                                "<button id=""pureBtn"" title=""Send only this raw text (no system prompt, no history)"">Pure</button>" &
                            "</div>" &
                            "<button id=""cancelBtn"" style=""display:none;"">Cancel</button>" &
                        "</div>")
        html.AppendLine("  </div>")
        html.AppendLine("  <div class=""hint"">Drag & drop a file • Enter=send • Shift+Enter=newline • Ctrl+L=clear</div>")
        html.AppendLine("</div>")

        ' JS
        html.AppendLine("<script>")
        html.AppendLine("window.__botName=" & Newtonsoft.Json.JsonConvert.SerializeObject(botName) & ";")
        html.AppendLine("let __supportsFiles=false;")
        html.AppendLine("let __pendingFilePath='';")
        html.AppendLine("let dark=false;")
        html.AppendLine("let __currentJobId=null;")
        html.AppendLine("let __jobCanceled=false;")
        html.AppendLine("let __typingBubbleId=null;")
        html.AppendLine("let __jobStartTs=0;")
        html.AppendLine("let __elapsedTimer=null;")

        ' Press feedback
        html.AppendLine("(function(){const pressOn=e=>{const b=e.target.closest('button');if(!b||b.disabled)return;b.classList.add('is-pressed');};const pressOff=()=>{document.querySelectorAll('button.is-pressed').forEach(b=>b.classList.remove('is-pressed'));};['mousedown','touchstart'].forEach(ev=>document.addEventListener(ev,pressOn,{passive:true}));['mouseup','mouseleave','blur'].forEach(ev=>document.addEventListener(ev,pressOff));document.addEventListener('keydown',e=>{if((e.key===' '||e.key==='Enter')){const b=e.target.closest('button');if(b&&!b.disabled)b.classList.add('is-pressed');}});document.addEventListener('keyup',e=>{if(e.key===' '||e.key==='Enter')pressOff();});})();")

        ' Helpers
        html.AppendLine("function copyText(t){if(navigator.clipboard){return navigator.clipboard.writeText(t);}return new Promise((res,rej)=>{try{const ta=document.createElement('textarea');ta.value=t;ta.style.position='fixed';ta.style.left='-9999px';document.body.appendChild(ta);ta.select();document.execCommand('copy');ta.remove();res();}catch(e){rej(e);}});}")
        html.AppendLine("function enhanceCodeBlocks(scope){(scope||document).querySelectorAll('pre').forEach(pre=>{if(pre.dataset.enhanced==='1')return;const btn=document.createElement('button');btn.type='button';btn.className='code-copy-btn';btn.innerHTML='<svg viewBox=""0 0 24 24"" fill=""none"" stroke=""currentColor"" stroke-width=""2"" stroke-linecap=""round"" stroke-linejoin=""round""><rect x=""9"" y=""9"" width=""13"" height=""13"" rx=""2"" ry=""2""/><path d=""M5 15H4a2 2 0 0 1-2-2V4c0-1.1.9-2 2-2h9a2 2 0 0 1 2 2v1""/></svg>';btn.addEventListener('click',()=>{const code=pre.querySelector('code');const txt=code?code.innerText:pre.innerText;copyText(txt).then(()=>{btn.classList.add('copied');setTimeout(()=>btn.classList.remove('copied'),1500);});});pre.appendChild(btn);pre.dataset.enhanced='1';});}")
        html.AppendLine("const api=async(cmd,data={})=>{try{const r=await fetch('/inky/api',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(Object.assign({Command:cmd},data))});const txt=await r.text();try{return JSON.parse(txt);}catch{return{ok:false,error:txt}}}catch(e){return{ok:false,error:e.message||'Network error'}}};")

        html.AppendLine("const chatEl=document.getElementById('chat');")
        html.AppendLine("const msgEl=document.getElementById('msg');")
        html.AppendLine("const modelSel=document.getElementById('modelSel');")
        html.AppendLine("const copyBtn=document.getElementById('copyBtn');")
        html.AppendLine("const clearBtn=document.getElementById('clearBtn');")
        html.AppendLine("const themeBtn=document.getElementById('themeBtn');")
        html.AppendLine("const cancelBtn=document.getElementById('cancelBtn');")
        html.AppendLine("const chat1Btn=document.getElementById('chat1Btn');")
        html.AppendLine("const chat2Btn=document.getElementById('chat2Btn');")
        html.AppendLine("const sendBtn=document.getElementById('sendBtn');")
        html.AppendLine("const pureBtn=document.getElementById('pureBtn');")
        html.AppendLine("const topbar=document.querySelector('.topbar');")

        html.AppendLine("function setTheme(isDark){dark=!!isDark;document.documentElement.classList.toggle('light',!dark);} ")
        html.AppendLine("function forceExternalLinks(scope){try{(scope||document).querySelectorAll('a[href]').forEach(a=>{a.target='_blank';a.rel='noopener noreferrer';});}catch{}}")
        html.AppendLine("function setActiveChatBtn(id){document.querySelectorAll('.chatTab').forEach(b=>b.classList.toggle('active',b.dataset.chat==String(id)));}")
        html.AppendLine("function disableChatSwitch(dis){chat1Btn.disabled=dis;chat2Btn.disabled=dis;}")

        ' Tooltip updates & responsive adjustment
        html.AppendLine("function updateModelTooltip(){try{if(!modelSel) return;const opt=modelSel.options[modelSel.selectedIndex];if(opt){modelSel.title=opt.textContent||'Model';}}catch{}}")
        html.AppendLine("function adjustModelSel(){if(!topbar) return;requestAnimationFrame(()=>{if(topbar.scrollWidth>topbar.clientWidth){modelSel.classList.add('squeezed');}else{modelSel.classList.remove('squeezed');}});}")
        html.AppendLine("window.addEventListener('resize',adjustModelSel);")

        html.AppendLine("function render(turns){chatEl.innerHTML='';for(const t of (turns||[])){const row=document.createElement('div');row.className='row '+(t.role==='user'?'user':'bot');const bub=document.createElement('div');bub.className='bubble';const rl=document.createElement('div');rl.className='role';rl.textContent=(t.role==='user'?'You':(window.__botName||'Bot'));bub.appendChild(rl);const cont=document.createElement('div');if(t && t.html){cont.innerHTML=t.html;forceExternalLinks(cont);}else if(t && t.markdown){const safe=t.markdown.replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;').replaceAll('\n','<br>');cont.innerHTML=safe;}bub.appendChild(cont);row.appendChild(bub);chatEl.appendChild(row);}chatEl.scrollTop=chatEl.scrollHeight;enhanceCodeBlocks(chatEl);} ")
        html.AppendLine("function addTempAssistantBubble(html){const id='tmp-'+Math.random().toString(36).slice(2);chatEl.insertAdjacentHTML('beforeend',`<div class=""row bot"" id=""${id}""><div class=""bubble""><div class=""role"">${window.__botName||'Bot'}</div><div class=""tmpContent"">${html}</div></div></div>`);chatEl.scrollTop=chatEl.scrollHeight;return id;}")
        html.AppendLine("function removeTempBubble(id){const el=document.getElementById(id);if(el)el.remove();}")
        html.AppendLine("function replaceAssistantBubble(id,html){const row=document.getElementById(id);if(!row)return;const c=row.querySelector('.tmpContent');if(c){c.innerHTML=html;forceExternalLinks(row);enhanceCodeBlocks(row);}}")

        ' Typing + elapsed
        html.AppendLine("function ensureTypingBubble(){if(__typingBubbleId)return;const content='<div class=""typing-container""><span class=""typing-dots""><span></span><span></span><span></span></span><span id=""typingElapsed"" class=""typing-elapsed"" style=""display:none;"">(0s)</span></div>';__typingBubbleId=addTempAssistantBubble(content);}")
        html.AppendLine("function updateElapsed(){if(!__typingBubbleId)return;const el=document.getElementById('typingElapsed');if(!el)return;const sec=Math.floor((Date.now()-__jobStartTs)/1000);if(sec>=10){el.style.display='inline-block';el.textContent='(' + sec + 's)';}}")
        html.AppendLine("function startElapsedTimer(){stopElapsedTimer();__jobStartTs=Date.now();__elapsedTimer=setInterval(updateElapsed,1000);}")
        html.AppendLine("function stopElapsedTimer(){if(__elapsedTimer){clearInterval(__elapsedTimer);__elapsedTimer=null;}const el=document.getElementById('typingElapsed');if(el)el.style.display='none';}")
        html.AppendLine("function removeTypingBubble(){if(__typingBubbleId){removeTempBubble(__typingBubbleId);__typingBubbleId=null;}stopElapsedTimer();}")

        ' boot
        html.AppendLine("async function boot(){const st=await api('inky_getstate');if(!st.ok){alert(st.error||'Init failed');return;}__supportsFiles=(st.supportsFiles===true);setTheme(st.darkMode!==false);render(st.history||[]);modelSel.innerHTML='';for(const m of (st.models||[])){const o=document.createElement('option');o.value=m.key||'';o.textContent=m.label||'';o.disabled=!!m.disabled;o.title=o.textContent;if(m.selected&&!o.disabled)o.selected=true;modelSel.appendChild(o);}if(!modelSel.value){const fe=[...modelSel.options].find(o=>!o.disabled&&o.value);if(fe)fe.selected=true;}updateModelTooltip();if(st.greeting && (!Array.isArray(st.history)||st.history.length===0)){msgEl.placeholder=st.greeting;}setActiveChatBtn(st.activeChat||1);adjustModelSel();} ")

        ' pollJob
        html.AppendLine("async function pollJob(jobId){if(!jobId)return;__currentJobId=jobId;__jobCanceled=false;ensureTypingBubble();startElapsedTimer();cancelBtn.style.display='inline-block';disableChatSwitch(true);try{for(;;){await new Promise(r=>setTimeout(r,2000));if(__jobCanceled)break;const s=await api('inky_jobstatus',{Job:jobId});if(!s.ok){console.warn('job status error',s.error);break;}if(s.status==='running'){continue;}const st=await api('inky_getstate');if(st.ok){render(st.history||[]);}break;} }finally{cancelBtn.style.display='none';removeTypingBubble();sendBtn.disabled=false;pureBtn.disabled=false;disableChatSwitch(false);__currentJobId=null;adjustModelSel();}}")

        ' send (normal)
        html.AppendLine("async function send(){if(__currentJobId){return;}const t=msgEl.value.trim();if(!t)return;msgEl.value='';sendBtn.disabled=true;pureBtn.disabled=true;chatEl.insertAdjacentHTML('beforeend',`<div class=""row user""><div class=""bubble""><div class=""role"">You</div><div>${t.replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;').replaceAll('\n','<br>')}</div></div></div>`);let typingId=addTempAssistantBubble('<span class=""typing-dots""><span></span><span></span><span></span></span>');const payload={Text:t};if(__pendingFilePath)payload.FileObject=__pendingFilePath;let r;try{r=await api('inky_send',payload);}catch(e){r={ok:false,error:e.message||'Network error'};}if(!r||!r.ok){removeTempBubble(typingId);sendBtn.disabled=false;pureBtn.disabled=false;alert(r&&r.error||'Error');__pendingFilePath='';adjustModelSel();return;}__pendingFilePath='';if(r.job){if(r.history){render(r.history||[]);}removeTempBubble(typingId);__typingBubbleId=null;ensureTypingBubble();startElapsedTimer();cancelBtn.style.display='inline-block';disableChatSwitch(true);pollJob(r.job);}else{removeTempBubble(typingId);sendBtn.disabled=false;pureBtn.disabled=false;if(r.history){render(r.history||[]);}adjustModelSel();}}")

        ' pureSend
        html.AppendLine("async function pureSend(){if(__currentJobId){return;}const t=msgEl.value.trim();if(!t)return;msgEl.value='';sendBtn.disabled=true;pureBtn.disabled=true;chatEl.insertAdjacentHTML('beforeend',`<div class=""row user""><div class=""bubble""><div class=""role"">You</div><div>${('Pure: '+t).replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;').replaceAll('\n','<br>')}</div></div></div>`);let typingId=addTempAssistantBubble('<span class=""typing-dots""><span></span><span></span><span></span></span>');const payload={Text:t};if(__pendingFilePath)payload.FileObject=__pendingFilePath;let r;try{r=await api('inky_pure',payload);}catch(e){r={ok:false,error:e.message||'Network error'};}if(!r||!r.ok){removeTempBubble(typingId);sendBtn.disabled=false;pureBtn.disabled=false;alert(r&&r.error||'Error');__pendingFilePath='';adjustModelSel();return;}__pendingFilePath='';if(r.job){if(r.history){render(r.history||[]);}removeTempBubble(typingId);__typingBubbleId=null;ensureTypingBubble();startElapsedTimer();cancelBtn.style.display='inline-block';disableChatSwitch(true);pollJob(r.job);}else{removeTempBubble(typingId);sendBtn.disabled=false;pureBtn.disabled=false;if(r.history){render(r.history||[]);}adjustModelSel();}}")

        ' drag/drop
        html.AppendLine("(function(){const stop=e=>{e.preventDefault();e.stopPropagation();};['dragenter','dragover','dragleave','drop'].forEach(ev=>document.addEventListener(ev,stop,false));document.addEventListener('drop',async e=>{const files=[...(e.dataTransfer&&e.dataTransfer.files)||[]];if(!files.length)return;const f=files[0];if(!__supportsFiles){addTempAssistantBubble('File uploads are not supported for the current model.');return;}const tempId=addTempAssistantBubble(`Uploading <b>${f.name.replaceAll('&','&amp;')}</b> (${(f.size/1024).toFixed(1)} KB)…`);try{const fr=new FileReader();const dataUrl=await new Promise((res,rej)=>{fr.onerror=()=>rej(new Error('read error'));fr.onload=()=>res(fr.result);fr.readAsDataURL(f);});const r=await api('inky_upload',{Name:f.name,DataUrl:String(dataUrl||'')});if(!r.ok){replaceAssistantBubble(tempId,'Upload failed: '+(r.error||'unknown'));return;}if(r.supported===false){replaceAssistantBubble(tempId,'File uploads are not supported for this model.');return;}__pendingFilePath=r.path||'';replaceAssistantBubble(tempId,`Added file: <b>${(r.name||f.name).replaceAll('&','&amp;')}</b>`);}catch(err){replaceAssistantBubble(tempId,'Upload failed: '+(err&&err.message?err.message:'unknown'));}} ,false);})();")

        ' events
        html.AppendLine("modelSel.addEventListener('change',async()=>{if(__currentJobId)return;const opt=modelSel.options[modelSel.selectedIndex];if(!opt||opt.disabled||!opt.value){const fe=[...modelSel.options].find(o=>!o.disabled&&o.value);if(fe)fe.selected=true;}const r=await api('inky_setmodel',{Key:opt.value});updateModelTooltip();adjustModelSel();if(!r.ok){alert(r.error||'Failed to set model');return;}if(typeof r.supportsFiles==='boolean')__supportsFiles=r.supportsFiles;});")
        html.AppendLine("clearBtn.addEventListener('click',async()=>{if(__currentJobId)return;const r=await api('inky_clear');if(r.ok){render([]);if(r.greeting)msgEl.placeholder=r.greeting;}else{alert(r.error||'Failed to clear');}adjustModelSel();});")
        html.AppendLine("copyBtn.addEventListener('click',async()=>{const r=await api('inky_copylast');if(!r.ok){alert(r.error||'Nothing to copy')}});")
        html.AppendLine("themeBtn.addEventListener('click',async()=>{if(__currentJobId)return;const target=!dark;setTheme(target);const r=await api('inky_toggletheme');if(!r.ok){setTheme(!target);alert(r.error||'Theme switch failed');return;}if(typeof r.darkMode==='boolean')setTheme(r.darkMode===true);adjustModelSel();});")
        html.AppendLine("msgEl.addEventListener('keydown',e=>{if(e.key==='Enter'&&!e.shiftKey){e.preventDefault();send();}if(e.ctrlKey&&e.key.toLowerCase()==='l'){e.preventDefault();clearBtn.click();}});")
        html.AppendLine("sendBtn.addEventListener('click',send);")
        html.AppendLine("pureBtn.addEventListener('click',pureSend);")
        html.AppendLine("cancelBtn.addEventListener('click',async()=>{if(!__currentJobId)return;__jobCanceled=true;await api('inky_cancel',{Job:__currentJobId});});")
        html.AppendLine("chatEl.addEventListener('click',e=>{const a=e.target&&e.target.closest&&e.target.closest('a[href]');if(!a)return;if(a.target!=='_blank'){a.target='_blank';a.rel='noopener noreferrer';}});")
        html.AppendLine("async function switchChat(n){if(__currentJobId)return;const r=await api('inky_switch',{Chat:String(n)});if(!r.ok){alert(r.error||'Switch failed');return;}setActiveChatBtn(r.activeChat||n);render(r.history||[]);if(r.greeting){msgEl.placeholder=r.greeting;}updateModelTooltip();adjustModelSel();}")
        html.AppendLine("chat1Btn.addEventListener('click',()=>switchChat(1));")
        html.AppendLine("chat2Btn.addEventListener('click',()=>switchChat(2));")

        html.AppendLine("boot();")
        html.AppendLine("</script>")
        html.AppendLine("</body></html>")
        Return html.ToString()
    End Function

    Private Function oldBuildInkyHtmlPage() As System.String
        Dim botName As String = GetBotName()
        Dim brandName As String = If(Not String.IsNullOrWhiteSpace(AN), AN, botName)
        Dim logoUrl As String = GetLogoDataUrl()
        Dim greet As String = GetFriendlyGreeting()

        Dim html As New System.Text.StringBuilder()

        html.AppendLine("<!doctype html>")
        html.AppendLine("<html lang=""en""><head><meta charset=""utf-8"">")
        html.AppendLine("<meta name=""viewport"" content=""width=device-width, initial-scale=1"">")
        html.AppendLine("<link rel=""shortcut icon"" type=""image/png"" href=""" & System.Net.WebUtility.HtmlEncode(logoUrl) & """>")
        html.AppendLine("<link rel=""icon"" type=""image/png"" href=""" & System.Net.WebUtility.HtmlEncode(logoUrl) & """>")
        html.AppendLine("<title>" & System.Net.WebUtility.HtmlEncode(brandName) & " — Local Chat</title>")

        ' CSS
        html.AppendLine("<style>")
        html.AppendLine(":root{--bg:#0b0f14;--card:#11161d;--fg:#e8eef6;--muted:#9aa8b7;--border:#1b2430;--border-strong:#2d3744;--elev:#1a222c;--press-shadow:inset 0 2px 6px rgba(0,0,0,.45);}")
        html.AppendLine(":root.light{--bg:#f6f7f9;--card:#ffffff;--fg:#0e1116;--muted:#5d6a77;--border:#e2e5e9;--border-strong:#c9cfd6;--elev:#eef1f4;--press-shadow:inset 0 2px 5px rgba(0,0,0,.08);}")
        html.AppendLine("html,body{height:100%;margin:0;font-family:system-ui,Segoe UI,Roboto,Arial,sans-serif;background:var(--bg);color:var(--fg);}")
        html.AppendLine(".wrap{display:flex;flex-direction:column;height:100%;}")
        html.AppendLine(".topbar{display:flex;gap:.5rem;align-items:center;padding:.75rem 1rem;border-bottom:1px solid var(--border);background:var(--card);position:sticky;top:0;z-index:5}")
        html.AppendLine(".topline{display:flex;align-items:center;gap:.6rem}")
        html.AppendLine(".topline img.logo{width:24px;height:24px;border-radius:6px;display:block}")
        html.AppendLine(".topline .brandbig{font-weight:700}")
        html.AppendLine(".topline .sub{color:var(--muted);font-size:.9rem}")
        html.AppendLine(".muted{color:var(--muted);font-size:.85rem}")
        html.AppendLine(".spacer{flex:1}")
        html.AppendLine("select,button,input,textarea{background:var(--card);color:var(--fg);border:1px solid var(--border);border-radius:.6rem;font:inherit;}")
        html.AppendLine("select,button,input{padding:.5rem .7rem;}")
        html.AppendLine("button{cursor:pointer;transition:background .16s,filter .12s,transform .08s,box-shadow .18s;}")
        html.AppendLine("button:hover{filter:brightness(1.07)}")
        html.AppendLine("button:disabled{opacity:.5;cursor:not-allowed}")
        html.AppendLine("button.is-pressed, .chatTab.is-pressed{transform:translateY(1px);box-shadow:var(--press-shadow);filter:brightness(.92);}")
        html.AppendLine("button:active:not(:disabled){transform:translateY(1px);box-shadow:var(--press-shadow);filter:brightness(.9);}")
        html.AppendLine(".chat{flex:1;overflow:auto;padding:1rem;}")
        html.AppendLine(".row{display:flex;margin:0 auto 1rem auto;max-width:1000px;padding:0 .25rem;}")
        html.AppendLine(".row.bot{justify-content:flex-start}")
        html.AppendLine(".row.user{justify-content:flex-end}")
        html.AppendLine(".bubble{max-width:75%;padding:1rem;border:1px solid var(--border);background:var(--card);border-radius:1rem;box-shadow:0 1px 3px rgba(0,0,0,.25)}")
        html.AppendLine(".bot .bubble{border-top-right-radius:.35rem}")
        html.AppendLine(".user .bubble{border-top-left-radius:.35rem}")
        html.AppendLine(".role{font-size:.75rem;color:var(--muted);margin-bottom:.25rem}")
        html.AppendLine(".inputbar{display:flex;gap:.5rem;padding:1rem;border-top:1px solid var(--border);background:var(--card)}")
        html.AppendLine("textarea{flex:1;resize:vertical;min-height:52px;max-height:220px;border-radius:.8rem;padding:.75rem;line-height:1.25;}")
        html.AppendLine(".hint{font-size:.7rem;letter-spacing:.3px;color:var(--muted);padding:.25rem 1rem 1rem}")
        html.AppendLine("a{color:inherit;text-decoration:underline;text-decoration-color:rgba(255,255,255,.35)}")
        html.AppendLine(":root.light a{text-decoration-color:rgba(0,0,0,.4)}")
        html.AppendLine("a:hover{filter:brightness(1.15)}")
        html.AppendLine("code,pre{font-family:ui-monospace,Consolas,monospace;font-size:.85rem}")
        html.AppendLine("pre{overflow:auto;padding:.75rem;border:1px solid var(--border);border-radius:.6rem;position:relative;background:var(--elev);}")
        html.AppendLine(".code-copy-btn{position:absolute;top:6px;right:6px;padding:4px 8px;font-size:.65rem;line-height:1;border:1px solid var(--border);border-radius:4px;background:rgba(0,0,0,.45);backdrop-filter:blur(3px);cursor:pointer;display:flex;align-items:center;gap:6px;color:var(--fg);opacity:0;transition:opacity .18s,background .18s;}")
        html.AppendLine("pre:hover .code-copy-btn{opacity:1}")
        html.AppendLine(".code-copy-btn svg{width:16px;height:16px;display:block}")
        html.AppendLine(".code-copy-btn.copied{background:#2c3440;color:#fff}")
        html.AppendLine(":root.light .code-copy-btn.copied{background:#d5d9dd;color:#111}")
        html.AppendLine(".code-copy-btn:focus{outline:2px solid var(--border-strong);}")
        html.AppendLine(".chatTab{padding:.45rem .55rem;min-width:32px;font-size:.7rem;font-weight:600;line-height:1;border:1px solid var(--border);background:var(--card);color:var(--muted);transition:background .18s,border-color .18s,color .18s,transform .08s,box-shadow .18s;}")
        html.AppendLine(".chatTab:hover:not(:disabled){background:var(--elev);color:var(--fg);}")
        html.AppendLine(".chatTab.active{background:#222b35;border-color:var(--border-strong);color:#fff;box-shadow:inset 0 0 0 1px #303c46;}")
        html.AppendLine(":root.light .chatTab.active{background:#e2e5e9;border-color:var(--border-strong);color:#0e1116;box-shadow:inset 0 0 0 1px #c9cfd6;}")
        html.AppendLine(".chatTab:focus{outline:2px solid var(--border-strong);outline-offset:1px;}")
        ' Typing + elapsed
        html.AppendLine(".typing-dots{display:inline-flex;gap:6px;align-items:center;}")
        html.AppendLine(".typing-dots span{width:7px;height:7px;border-radius:50%;background:currentColor;opacity:.35;animation:tdots 1.2s infinite ease-in-out;}")
        html.AppendLine(".typing-dots span:nth-child(2){animation-delay:.2s}")
        html.AppendLine(".typing-dots span:nth-child(3){animation-delay:.4s}")
        html.AppendLine("@keyframes tdots{0%,80%,100%{transform:translateY(0);opacity:.3}40%{transform:translateY(-5px);opacity:.85}}")
        html.AppendLine(".typing-elapsed{margin-left:8px;font-size:.65rem;color:var(--muted);font-family:ui-monospace,monospace;opacity:.8;}")
        html.AppendLine(".actions{display:flex;flex-direction:row;gap:.5rem;align-items:stretch;}")
        html.AppendLine(".actions .stack{display:flex;flex-direction:column;gap:.5rem;}")
        html.AppendLine("#cancelBtn{display:none;align-self:stretch;height:auto;}")
        html.AppendLine("@media (max-width:640px){.actions{flex-direction:column;}.actions .stack{flex-direction:row;}.actions .stack button{flex:1;}#cancelBtn{align-self:auto;height:auto;}}")
        html.AppendLine("</style>")

        html.AppendLine("</head><body>")
        html.AppendLine("<div class=""wrap"">")

        ' Top bar
        html.AppendLine("  <div class=""topbar"">")
        html.AppendLine("    <div class=""topline"">")
        If Not String.IsNullOrWhiteSpace(logoUrl) Then
            html.AppendLine("      <img class=""logo"" src=""" & System.Net.WebUtility.HtmlEncode(logoUrl) & """ alt=""logo"">")
        End If
        html.AppendLine("      <div class=""brandbig"">" & System.Net.WebUtility.HtmlEncode(brandName) & "</div>")
        html.AppendLine("      <div class=""sub"">Local Chat</div>")
        html.AppendLine("    </div>")
        html.AppendLine("    <div class=""spacer""></div>")
        html.AppendLine("    <select id=""modelSel"" title=""Model""></select>")
        html.AppendLine("    <button id=""copyBtn"" title=""Copy last answer to clipboard"">Copy last</button>")
        html.AppendLine("    <button id=""clearBtn"" title=""Clear current conversation"">Clear</button>")
        html.AppendLine("    <button id=""chat1Btn"" class=""chatTab"" data-chat=""1"" title=""Chat 1"">1</button>")
        html.AppendLine("    <button id=""chat2Btn"" class=""chatTab"" data-chat=""2"" title=""Chat 2"">2</button>")
        html.AppendLine("    <button id=""themeBtn"" title=""Toggle theme"">Theme</button>")
        html.AppendLine("  </div>")

        html.AppendLine("  <div id=""chat"" class=""chat""></div>")

        html.AppendLine("  <div class=""inputbar"">")
        html.AppendLine("    <textarea id=""msg"" placeholder=""" & System.Net.WebUtility.HtmlEncode(greet) & """ autofocus></textarea>")
        html.AppendLine("    <div class=""actions"">" &
                            "<div class=""stack"">" &
                                "<button id=""sendBtn"">Send</button>" &
                                "<button id=""pureBtn"" title=""Send only this raw text (no system prompt, no history)"">Pure</button>" &
                            "</div>" &
                            "<button id=""cancelBtn"" style=""display:none;"">Cancel</button>" &
                        "</div>")
        html.AppendLine("  </div>")
        html.AppendLine("  <div class=""hint"">Drag & drop a file • Enter=send • Shift+Enter=newline • Ctrl+L=clear</div>")
        html.AppendLine("</div>")

        ' JS
        html.AppendLine("<script>")
        html.AppendLine("window.__botName=" & Newtonsoft.Json.JsonConvert.SerializeObject(botName) & ";")
        html.AppendLine("let __supportsFiles=false;")
        html.AppendLine("let __pendingFilePath='';")
        html.AppendLine("let dark=false;")
        html.AppendLine("let __currentJobId=null;")
        html.AppendLine("let __jobCanceled=false;")
        html.AppendLine("let __typingBubbleId=null;")
        html.AppendLine("let __jobStartTs=0;")
        html.AppendLine("let __elapsedTimer=null;")

        ' Press feedback
        html.AppendLine("(function(){const pressOn=e=>{const b=e.target.closest('button');if(!b||b.disabled)return;b.classList.add('is-pressed');};const pressOff=()=>{document.querySelectorAll('button.is-pressed').forEach(b=>b.classList.remove('is-pressed'));};['mousedown','touchstart'].forEach(ev=>document.addEventListener(ev,pressOn,{passive:true}));['mouseup','mouseleave','blur'].forEach(ev=>document.addEventListener(ev,pressOff));document.addEventListener('keydown',e=>{if((e.key===' '||e.key==='Enter')){const b=e.target.closest('button');if(b&&!b.disabled)b.classList.add('is-pressed');}});document.addEventListener('keyup',e=>{if(e.key===' '||e.key==='Enter')pressOff();});})();")

        ' Helpers
        html.AppendLine("function copyText(t){if(navigator.clipboard){return navigator.clipboard.writeText(t);}return new Promise((res,rej)=>{try{const ta=document.createElement('textarea');ta.value=t;ta.style.position='fixed';ta.style.left='-9999px';document.body.appendChild(ta);ta.select();document.execCommand('copy');ta.remove();res();}catch(e){rej(e);}});}")
        html.AppendLine("function enhanceCodeBlocks(scope){(scope||document).querySelectorAll('pre').forEach(pre=>{if(pre.dataset.enhanced==='1')return;const btn=document.createElement('button');btn.type='button';btn.className='code-copy-btn';btn.innerHTML='<svg viewBox=""0 0 24 24"" fill=""none"" stroke=""currentColor"" stroke-width=""2"" stroke-linecap=""round"" stroke-linejoin=""round""><rect x=""9"" y=""9"" width=""13"" height=""13"" rx=""2"" ry=""2""/><path d=""M5 15H4a2 2 0 0 1-2-2V4c0-1.1.9-2 2-2h9a2 2 0 0 1 2 2v1""/></svg>';btn.addEventListener('click',()=>{const code=pre.querySelector('code');const txt=code?code.innerText:pre.innerText;copyText(txt).then(()=>{btn.classList.add('copied');setTimeout(()=>btn.classList.remove('copied'),1500);});});pre.appendChild(btn);pre.dataset.enhanced='1';});}")
        html.AppendLine("const api=async(cmd,data={})=>{try{const r=await fetch('/inky/api',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(Object.assign({Command:cmd},data))});const txt=await r.text();try{return JSON.parse(txt);}catch{return{ok:false,error:txt}}}catch(e){return{ok:false,error:e.message||'Network error'}}};")

        html.AppendLine("const chatEl=document.getElementById('chat');")
        html.AppendLine("const msgEl=document.getElementById('msg');")
        html.AppendLine("const modelSel=document.getElementById('modelSel');")
        html.AppendLine("const copyBtn=document.getElementById('copyBtn');")
        html.AppendLine("const clearBtn=document.getElementById('clearBtn');")
        html.AppendLine("const themeBtn=document.getElementById('themeBtn');")
        html.AppendLine("const cancelBtn=document.getElementById('cancelBtn');")
        html.AppendLine("const chat1Btn=document.getElementById('chat1Btn');")
        html.AppendLine("const chat2Btn=document.getElementById('chat2Btn');")
        html.AppendLine("const sendBtn=document.getElementById('sendBtn');")
        html.AppendLine("const pureBtn=document.getElementById('pureBtn');") ' NEW

        html.AppendLine("function setTheme(isDark){dark=!!isDark;document.documentElement.classList.toggle('light',!dark);} ")
        html.AppendLine("function forceExternalLinks(scope){try{(scope||document).querySelectorAll('a[href]').forEach(a=>{a.target='_blank';a.rel='noopener noreferrer';});}catch{}}")
        html.AppendLine("function setActiveChatBtn(id){document.querySelectorAll('.chatTab').forEach(b=>b.classList.toggle('active',b.dataset.chat==String(id)));}")
        html.AppendLine("function disableChatSwitch(dis){chat1Btn.disabled=dis;chat2Btn.disabled=dis;}")

        html.AppendLine("function render(turns){chatEl.innerHTML='';for(const t of (turns||[])){const row=document.createElement('div');row.className='row '+(t.role==='user'?'user':'bot');const bub=document.createElement('div');bub.className='bubble';const rl=document.createElement('div');rl.className='role';rl.textContent=(t.role==='user'?'You':(window.__botName||'Bot'));bub.appendChild(rl);const cont=document.createElement('div');if(t && t.html){cont.innerHTML=t.html;forceExternalLinks(cont);}else if(t && t.markdown){const safe=t.markdown.replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;').replaceAll('\n','<br>');cont.innerHTML=safe;}bub.appendChild(cont);row.appendChild(bub);chatEl.appendChild(row);}chatEl.scrollTop=chatEl.scrollHeight;enhanceCodeBlocks(chatEl);} ")
        html.AppendLine("function addTempAssistantBubble(html){const id='tmp-'+Math.random().toString(36).slice(2);chatEl.insertAdjacentHTML('beforeend',`<div class=""row bot"" id=""${id}""><div class=""bubble""><div class=""role"">${window.__botName||'Bot'}</div><div class=""tmpContent"">${html}</div></div></div>`);chatEl.scrollTop=chatEl.scrollHeight;return id;}")
        html.AppendLine("function removeTempBubble(id){const el=document.getElementById(id);if(el)el.remove();}")
        html.AppendLine("function replaceAssistantBubble(id,html){const row=document.getElementById(id);if(!row)return;const c=row.querySelector('.tmpContent');if(c){c.innerHTML=html;forceExternalLinks(row);enhanceCodeBlocks(row);}}")

        ' Typing + elapsed
        html.AppendLine("function ensureTypingBubble(){if(__typingBubbleId)return;const content='<div class=""typing-container""><span class=""typing-dots""><span></span><span></span><span></span></span><span id=""typingElapsed"" class=""typing-elapsed"" style=""display:none;"">(0s)</span></div>';__typingBubbleId=addTempAssistantBubble(content);}")
        html.AppendLine("function updateElapsed(){if(!__typingBubbleId)return;const el=document.getElementById('typingElapsed');if(!el)return;const sec=Math.floor((Date.now()-__jobStartTs)/1000);if(sec>=10){el.style.display='inline-block';el.textContent='(' + sec + 's)';}}")
        html.AppendLine("function startElapsedTimer(){stopElapsedTimer();__jobStartTs=Date.now();__elapsedTimer=setInterval(updateElapsed,1000);}")
        html.AppendLine("function stopElapsedTimer(){if(__elapsedTimer){clearInterval(__elapsedTimer);__elapsedTimer=null;}const el=document.getElementById('typingElapsed');if(el)el.style.display='none';}")
        html.AppendLine("function removeTypingBubble(){if(__typingBubbleId){removeTempBubble(__typingBubbleId);__typingBubbleId=null;}stopElapsedTimer();}")

        ' boot
        html.AppendLine("async function boot(){const st=await api('inky_getstate');if(!st.ok){alert(st.error||'Init failed');return;}__supportsFiles=(st.supportsFiles===true);setTheme(st.darkMode!==false);render(st.history||[]);modelSel.innerHTML='';for(const m of (st.models||[])){const o=document.createElement('option');o.value=m.key||'';o.textContent=m.label||'';o.disabled=!!m.disabled;if(m.selected&&!o.disabled)o.selected=true;modelSel.appendChild(o);}if(!modelSel.value){const fe=[...modelSel.options].find(o=>!o.disabled&&o.value);if(fe)fe.selected=true;}if(st.greeting && (!Array.isArray(st.history)||st.history.length===0)){msgEl.placeholder=st.greeting;}setActiveChatBtn(st.activeChat||1);} ")

        ' pollJob
        html.AppendLine("async function pollJob(jobId){if(!jobId)return;__currentJobId=jobId;__jobCanceled=false;ensureTypingBubble();startElapsedTimer();cancelBtn.style.display='inline-block';disableChatSwitch(true);try{for(;;){await new Promise(r=>setTimeout(r,2000));if(__jobCanceled)break;const s=await api('inky_jobstatus',{Job:jobId});if(!s.ok){console.warn('job status error',s.error);break;}if(s.status==='running'){continue;}const st=await api('inky_getstate');if(st.ok){render(st.history||[]);}break;} }finally{cancelBtn.style.display='none';removeTypingBubble();sendBtn.disabled=false;pureBtn.disabled=false;disableChatSwitch(false);__currentJobId=null;}}")

        ' send (normal)
        html.AppendLine("async function send(){if(__currentJobId){return;}const t=msgEl.value.trim();if(!t)return;msgEl.value='';sendBtn.disabled=true;pureBtn.disabled=true;chatEl.insertAdjacentHTML('beforeend',`<div class=""row user""><div class=""bubble""><div class=""role"">You</div><div>${t.replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;').replaceAll('\n','<br>')}</div></div></div>`);let typingId=addTempAssistantBubble('<span class=""typing-dots""><span></span><span></span><span></span></span>');const payload={Text:t};if(__pendingFilePath)payload.FileObject=__pendingFilePath;let r;try{r=await api('inky_send',payload);}catch(e){r={ok:false,error:e.message||'Network error'};}if(!r||!r.ok){removeTempBubble(typingId);sendBtn.disabled=false;pureBtn.disabled=false;alert(r&&r.error||'Error');__pendingFilePath='';return;}__pendingFilePath='';if(r.job){if(r.history){render(r.history||[]);}removeTempBubble(typingId);__typingBubbleId=null;ensureTypingBubble();startElapsedTimer();cancelBtn.style.display='inline-block';disableChatSwitch(true);pollJob(r.job);}else{removeTempBubble(typingId);sendBtn.disabled=false;pureBtn.disabled=false;if(r.history){render(r.history||[]);}}}")

        ' pureSend (NEW)
        html.AppendLine("async function pureSend(){if(__currentJobId){return;}const t=msgEl.value.trim();if(!t)return;msgEl.value='';sendBtn.disabled=true;pureBtn.disabled=true;chatEl.insertAdjacentHTML('beforeend',`<div class=""row user""><div class=""bubble""><div class=""role"">You</div><div>${('Pure: '+t).replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;').replaceAll('\n','<br>')}</div></div></div>`);let typingId=addTempAssistantBubble('<span class=""typing-dots""><span></span><span></span><span></span></span>');const payload={Text:t};if(__pendingFilePath)payload.FileObject=__pendingFilePath;let r;try{r=await api('inky_pure',payload);}catch(e){r={ok:false,error:e.message||'Network error'};}if(!r||!r.ok){removeTempBubble(typingId);sendBtn.disabled=false;pureBtn.disabled=false;alert(r&&r.error||'Error');__pendingFilePath='';return;}__pendingFilePath='';if(r.job){if(r.history){render(r.history||[]);}removeTempBubble(typingId);__typingBubbleId=null;ensureTypingBubble();startElapsedTimer();cancelBtn.style.display='inline-block';disableChatSwitch(true);pollJob(r.job);}else{removeTempBubble(typingId);sendBtn.disabled=false;pureBtn.disabled=false;if(r.history){render(r.history||[]);}}}")

        ' drag/drop
        html.AppendLine("(function(){const stop=e=>{e.preventDefault();e.stopPropagation();};['dragenter','dragover','dragleave','drop'].forEach(ev=>document.addEventListener(ev,stop,false));document.addEventListener('drop',async e=>{const files=[...(e.dataTransfer&&e.dataTransfer.files)||[]];if(!files.length)return;const f=files[0];if(!__supportsFiles){addTempAssistantBubble('File uploads are not supported for the current model.');return;}const tempId=addTempAssistantBubble(`Uploading <b>${f.name.replaceAll('&','&amp;')}</b> (${(f.size/1024).toFixed(1)} KB)…`);try{const fr=new FileReader();const dataUrl=await new Promise((res,rej)=>{fr.onerror=()=>rej(new Error('read error'));fr.onload=()=>res(fr.result);fr.readAsDataURL(f);});const r=await api('inky_upload',{Name:f.name,DataUrl:String(dataUrl||'')});if(!r.ok){replaceAssistantBubble(tempId,'Upload failed: '+(r.error||'unknown'));return;}if(r.supported===false){replaceAssistantBubble(tempId,'File uploads are not supported for this model.');return;}__pendingFilePath=r.path||'';replaceAssistantBubble(tempId,`Added file: <b>${(r.name||f.name).replaceAll('&','&amp;')}</b>`);}catch(err){replaceAssistantBubble(tempId,'Upload failed: '+(err&&err.message?err.message:'unknown'));}} ,false);})();")

        ' events
        html.AppendLine("modelSel.addEventListener('change',async()=>{if(__currentJobId)return;const opt=modelSel.options[modelSel.selectedIndex];if(!opt||opt.disabled||!opt.value){const fe=[...modelSel.options].find(o=>!o.disabled&&o.value);if(fe)fe.selected=true;return;}const r=await api('inky_setmodel',{Key:opt.value});if(!r.ok){alert(r.error||'Failed to set model');return;}if(typeof r.supportsFiles==='boolean')__supportsFiles=r.supportsFiles;});")
        html.AppendLine("clearBtn.addEventListener('click',async()=>{if(__currentJobId)return;const r=await api('inky_clear');if(r.ok){render([]);if(r.greeting)msgEl.placeholder=r.greeting;}else{alert(r.error||'Failed to clear');}});")
        html.AppendLine("copyBtn.addEventListener('click',async()=>{const r=await api('inky_copylast');if(!r.ok){alert(r.error||'Nothing to copy')}});")
        html.AppendLine("themeBtn.addEventListener('click',async()=>{if(__currentJobId)return;const target=!dark;setTheme(target);const r=await api('inky_toggletheme');if(!r.ok){setTheme(!target);alert(r.error||'Theme switch failed');return;}if(typeof r.darkMode==='boolean')setTheme(r.darkMode===true);});")
        html.AppendLine("msgEl.addEventListener('keydown',e=>{if(e.key==='Enter'&&!e.shiftKey){e.preventDefault();send();}if(e.ctrlKey&&e.key.toLowerCase()==='l'){e.preventDefault();clearBtn.click();}});")
        html.AppendLine("sendBtn.addEventListener('click',send);")
        html.AppendLine("pureBtn.addEventListener('click',pureSend);") ' NEW
        html.AppendLine("cancelBtn.addEventListener('click',async()=>{if(!__currentJobId)return;__jobCanceled=true;await api('inky_cancel',{Job:__currentJobId});});")
        html.AppendLine("chatEl.addEventListener('click',e=>{const a=e.target&&e.target.closest&&e.target.closest('a[href]');if(!a)return;if(a.target!=='_blank'){a.target='_blank';a.rel='noopener noreferrer';}});")
        html.AppendLine("async function switchChat(n){if(__currentJobId)return;const r=await api('inky_switch',{Chat:String(n)});if(!r.ok){alert(r.error||'Switch failed');return;}setActiveChatBtn(r.activeChat||n);render(r.history||[]);if(r.greeting){msgEl.placeholder=r.greeting;}}")
        html.AppendLine("chat1Btn.addEventListener('click',()=>switchChat(1));")
        html.AppendLine("chat2Btn.addEventListener('click',()=>switchChat(2));")

        html.AppendLine("boot();")
        html.AppendLine("</script>")
        html.AppendLine("</body></html>")
        Return html.ToString()
    End Function



    ' Builds a simple JSON response
    Private Function JsonOk(o As Object) As System.String
        Return "CT:json" & vbLf & Newtonsoft.Json.JsonConvert.SerializeObject(o)
    End Function
    Private Function JsonErr(msg As System.String) As System.String
        Return JsonOk(New With {.ok = False, .error = msg})
    End Function

    ' ===== (D) EXTEND ProcessRequestInAddIn with the Inky commands =====

    Private Async Function ProcessRequestInAddIn(
        body As System.String,
        rawUrl As System.String) As System.Threading.Tasks.Task(Of System.String)

        ' If this is a browser POST to our Inky API, j may be JSON; otherwise keep your existing flow
        If rawUrl IsNot Nothing AndAlso rawUrl.StartsWith(InkyApiRoute, System.StringComparison.OrdinalIgnoreCase) Then
            Try
                Dim j As Newtonsoft.Json.Linq.JObject = If(
                Not System.String.IsNullOrWhiteSpace(body),
                Newtonsoft.Json.Linq.JObject.Parse(body),
                New Newtonsoft.Json.Linq.JObject())
                Dim cmd As System.String = j("Command")?.ToString()

                Select Case cmd
                    Case "inky_getstate"
                        Dim st As InkyState = LoadInkyState()

                        Try
                            st.DarkMode = My.Settings.Inky_DarkMode
                        Catch
                        End Try

                        ' Re-compute per current selection on every getstate
                        Try
                            st.SupportsFileUploads = ComputeSupportsFiles(st.UseSecondApi, st.SelectedModelKey)
                            SaveInkyState(st)
                        Catch
                            st.SupportsFileUploads = False
                        End Try

                        Dim greeting As System.String = Nothing
                        If st.History.Count = 0 Then greeting = GetFriendlyGreeting()

                        Dim models As System.Collections.Generic.List(Of System.Object) =
                                    Await GetModelListForBrowserAsync(st)

                        Return JsonOk(New With {
                            .ok = True,
                            .history = ToBrowserTurns(LoadInkyState().History),
                            .greeting = greeting,
                            .models = models,
                            .modelLabel = GetSelectedModelLabel(st.UseSecondApi, st.SelectedModelKey),
                            .darkMode = st.DarkMode,
                            .supportsFiles = st.SupportsFileUploads,
                            .activeChat = activeChatId
                        })

                    Case "inky_switch"
                        Dim which As String = j("Chat")?.ToString()
                        activeChatId = If(which = "2", 2, 1)
                        Dim stSw = LoadInkyState()
                        If stSw.History.Count = 0 AndAlso Not stSw.DarkMode Then stSw.DarkMode = True
                        Dim greetingSwitch As String = If(stSw.History.Count = 0, GetFriendlyGreeting(), Nothing)

                        Try
                            My.Settings.Inky_LastChat = activeChatId
                            My.Settings.Save()
                        Catch
                        End Try

                        Return JsonOk(New With {
                            .ok = True,
                            .history = ToBrowserTurns(stSw.History),
                            .activeChat = activeChatId,
                            .darkMode = stSw.DarkMode,
                            .supportsFiles = ComputeSupportsFiles(stSw.UseSecondApi, stSw.SelectedModelKey),
                            .greeting = greetingSwitch
                        })

                    Case "inky_upload"
                        Try
                            Dim stU As InkyState = LoadInkyState()

                            ' Hard-enforce on server: do not accept data when unsupported
                            Dim supports As System.Boolean = False
                            Try
                                supports = ComputeSupportsFiles(stU.UseSecondApi, stU.SelectedModelKey)
                            Catch
                                supports = False
                            End Try
                            If Not supports Then
                                ' Tell client it is not supported; do NOT create any temp file
                                Return JsonOk(New With {.ok = True, .supported = False})
                            End If

                            Dim name As System.String = j("Name")?.ToString()
                            Dim dataUrl As System.String = j("DataUrl")?.ToString()
                            If System.String.IsNullOrWhiteSpace(name) OrElse System.String.IsNullOrWhiteSpace(dataUrl) Then
                                Return JsonErr("Missing file data.")
                            End If

                            Dim commaIx As System.Int32 = dataUrl.IndexOf(","c)
                            If commaIx < 0 Then Return JsonErr("Bad DataURL.")
                            Dim b64 As System.String = dataUrl.Substring(commaIx + 1)

                            Dim bytes() As System.Byte
                            Try
                                bytes = System.Convert.FromBase64String(b64)
                            Catch exB64 As System.Exception
                                Return JsonErr("Invalid base64: " & exB64.Message)
                            End Try

                            Dim dir As System.String = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "InkyUploads")
                            If Not System.IO.Directory.Exists(dir) Then System.IO.Directory.CreateDirectory(dir)

                            Dim safeName As System.String = System.IO.Path.GetFileName(name)
                            For Each c As System.Char In System.IO.Path.GetInvalidFileNameChars()
                                safeName = safeName.Replace(c, "_"c)
                            Next

                            Dim unique As System.String = System.Guid.NewGuid().ToString("N")
                            Dim target As System.String = System.IO.Path.Combine(dir, unique & "_" & safeName)

                            System.IO.File.WriteAllBytes(target, bytes)

                            Return JsonOk(New With {.ok = True, .supported = True, .path = target, .name = safeName, .size = bytes.LongLength})
                        Catch exUp As System.Exception
                            Return JsonErr("Upload failed: " & exUp.Message)
                        End Try

                    Case "inky_cancel"

                        ' Optional job id (preferred modern path)
                        Dim jobId As String = j("Job")?.ToString()

                        If Not String.IsNullOrWhiteSpace(jobId) Then
                            Dim job As LlmJob = Nothing
                            If Not jobMap.TryGetValue(jobId, job) Then
                                Return JsonErr("Unknown job.")
                            End If
                            Try
                                If job Is Nothing OrElse job.Cts Is Nothing Then
                                    Return JsonErr("Job has no cancellation handle.")
                                End If
                                If Not job.Cts.IsCancellationRequested Then
                                    job.Cts.Cancel()
                                    Return JsonOk(New With {
                                        .ok = True,
                                        .job = jobId,
                                        .status = "cancelRequested"
                                    })
                                Else
                                    Return JsonOk(New With {
                                        .ok = True,
                                        .job = jobId,
                                        .status = "alreadyCanceled"
                                    })
                                End If
                            Catch ex As System.Exception
                                Return JsonErr("Cancel failed: " & ex.Message)
                            End Try
                        End If

                        ' Legacy path: global CTS (kept for backward compatibility)
                        If llmOperationCts IsNot Nothing AndAlso Not llmOperationCts.IsCancellationRequested Then
                            Try
                                llmOperationCts.Cancel()
                                Return JsonOk(New With {
                                    .ok = True,
                                    .status = "cancelRequestedLegacy"
                                })
                            Catch ex As System.Exception
                                Return JsonErr("Legacy cancel failed: " & ex.Message)
                            End Try
                        End If

                        ' Fallback: try to find the most recent running job (if user pressed cancel without specifying id)
                        Dim fallbackJob As LlmJob =
                            jobMap.Values _
                                  .Where(Function(x) x IsNot Nothing AndAlso
                                                     x.Tcs IsNot Nothing AndAlso
                                                     Not x.Tcs.Task.IsCompleted) _
                                  .OrderByDescending(Function(x) x.CreatedUtc) _
                                  .FirstOrDefault()

                        If fallbackJob IsNot Nothing Then
                            Try
                                If Not fallbackJob.Cts.IsCancellationRequested Then
                                    fallbackJob.Cts.Cancel()
                                    Return JsonOk(New With {
                                        .ok = True,
                                        .job = fallbackJob.Id,
                                        .status = "cancelRequestedFallback"
                                    })
                                Else
                                    Return JsonOk(New With {
                                        .ok = True,
                                        .job = fallbackJob.Id,
                                        .status = "alreadyCanceledFallback"
                                    })
                                End If
                            Catch ex As System.Exception
                                Return JsonErr("Fallback cancel failed: " & ex.Message)
                            End Try
                        End If

                        Return JsonErr("No active operation to cancel.")


                    Case "inky_send"

                        ' ------------------ (A) Read request & validate ------------------
                        Dim fileObject As System.String = j("FileObject")?.ToString()
                        Dim uploadedTempPath As System.String = fileObject

                        Dim textBody As System.String = j("Text")?.ToString()
                        If System.String.IsNullOrWhiteSpace(textBody) Then
                            Return JsonErr("Please enter a message.")
                        End If

                        Dim st As InkyState = LoadInkyState()

                        ' Recompute upload capability (client may be stale)
                        Dim supportsFilesNow As System.Boolean = False
                        Try
                            supportsFilesNow = ComputeSupportsFiles(st.UseSecondApi, st.SelectedModelKey)
                        Catch
                            supportsFilesNow = False
                        End Try

                        ' ------------------ (B) File / clipboard object extraction (unchanged logic) ------------
                        Dim extractedDoc As System.String = Nothing
                        Dim extractedLabel As System.String = Nothing
                        Dim attachedType As System.String = Nothing
                        Dim hadInlineExtraction As System.Boolean = False

                        If Not System.String.IsNullOrWhiteSpace(fileObject) Then
                            Dim okOffice As System.Boolean = False
                            Try
                                okOffice = TryExtractOfficeText(fileObject, extractedDoc, extractedLabel)
                            Catch
                                okOffice = False
                            End Try

                            If okOffice Then
                                hadInlineExtraction = True
                                attachedType = "office"
                                fileObject = Nothing    ' Do NOT pass a path to the model
                            Else
                                Dim okText As System.Boolean = False
                                Try
                                    okText = TryExtractTextLike(fileObject, extractedDoc, extractedLabel)
                                Catch
                                    okText = False
                                End Try
                                If okText Then
                                    hadInlineExtraction = True
                                    attachedType = "text"
                                    fileObject = Nothing
                                Else
                                    If (Not supportsFilesNow) AndAlso (Not System.String.IsNullOrWhiteSpace(fileObject)) Then
                                        ' Model does not support file objects
                                        st.History.Add(New ChatTurn With {
                                            .Role = "assistant",
                                            .Markdown = "This model does not support file attachments.",
                                            .Html = MarkdownToHtml("This model does not support file attachments."),
                                            .Utc = System.DateTime.UtcNow
                                        })
                                        SaveInkyState(st)
                                        ' Cleanup temp
                                        Try
                                            If Not System.String.IsNullOrWhiteSpace(uploadedTempPath) AndAlso IO.File.Exists(uploadedTempPath) Then
                                                IO.File.Delete(uploadedTempPath)
                                            End If
                                        Catch
                                        End Try
                                        Return JsonOk(New With {.ok = True, .history = ToBrowserTurns(st.History)})
                                    End If
                                    ' Else: keep fileObject (e.g. PDF) for API that supports raw path/object
                                End If
                            End If
                        End If

                        ' ------------------ (C) Append user turn immediately ------------------
                        Dim userTurn As New ChatTurn With {
                            .Role = "user",
                            .Markdown = textBody,
                            .Html = MarkdownToHtml(textBody),
                            .Utc = Date.UtcNow
                        }
                        st.History.Add(userTurn)

                        ' Cap history (unchanged)
                        Dim cap As Integer = 0
                        Try : cap = INI_ChatCap : Catch : cap = 4000 : End Try
                        Dim clipped As List(Of ChatTurn) = CapHistoryToChars(st, cap)

                        ' Build dialog prompt now (passed into job)
                        Dim sbDialog As New System.Text.StringBuilder()
                        sbDialog.AppendLine("<DIALOG>")
                        For Each t In clipped
                            If t.Role = "user" Then
                                sbDialog.AppendLine("[USER] " & t.Markdown)
                            Else
                                sbDialog.AppendLine("[ASSISTANT] " & t.Markdown)
                            End If
                        Next
                        sbDialog.AppendLine("</DIALOG>")

                        ' Add inline extracted document block if any (unchanged)
                        If hadInlineExtraction AndAlso Not String.IsNullOrWhiteSpace(extractedDoc) Then
                            sbDialog.AppendLine()
                            Dim lbl As String = EscapeForXml(If(extractedLabel, "Attached document"))
                            Dim typ As String = If(String.IsNullOrWhiteSpace(attachedType), "text", attachedType)
                            sbDialog.AppendLine("<ATTACHED_DOCUMENT type=""" & typ & """ label=""" & lbl & """>")
                            sbDialog.AppendLine(extractedDoc)
                            sbDialog.AppendLine("</ATTACHED_DOCUMENT>")
                        End If

                        ' Persist state with user turn BEFORE returning (important)
                        SaveInkyState(st)

                        ' ------------------ (D) Prepare system prompt (same logic) ------------------
                        Dim sysPromptBase As String = GetSystemPromptChat()
                        Dim nowLocal As String = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss zzz", Globalization.CultureInfo.InvariantCulture)
                        sysPromptBase &= Environment.NewLine & "Current local date/time: " & nowLocal
                        sysPromptBase &= Environment.NewLine & $"Your name is '{AN6}'. Only if you are expressly asked you can say that you have been developped by David Rosenthal of the law firm VISCHER in Switzerland."

                        Dim useSecondApiLocal As Boolean = st.UseSecondApi
                        Dim selectedModelKeyLocal As String = st.SelectedModelKey

                        ' Capture file object (may be Nothing after extraction)
                        Dim finalFileObject As String = fileObject
                        Dim tempUploadPathCopy As String = uploadedTempPath

                        ' ------------------ (E) Create background job ------------------
                        Dim jobId As String = Guid.NewGuid().ToString("N")
                        Dim jobCts As New CancellationTokenSource()
                        Dim tcs As New TaskCompletionSource(Of String)(TaskCreationOptions.RunContinuationsAsynchronously)

                        Dim job As New LlmJob With {
                            .Id = jobId,
                            .CreatedUtc = Date.UtcNow,
                            .Tcs = tcs,
                            .Cts = jobCts,
                            .UseSecond = useSecondApiLocal,
                            .FileObject = finalFileObject
                        }

                        If Not jobMap.TryAdd(jobId, job) Then
                            jobCts.Dispose()
                            Return JsonErr("Failed to register job.")
                        End If
                        Threading.Interlocked.Increment(activeJobs)

                        ' Keep old global CTS reference for legacy inky_cancel support
                        Try
                            If llmOperationCts IsNot Nothing Then
                                llmOperationCts.Cancel()
                                llmOperationCts.Dispose()
                            End If
                        Catch
                        End Try
                        llmOperationCts = jobCts

                        System.Threading.Tasks.Task.Run(
                            Sub()
                                Dim originalCfgLoaded As Boolean = False
                                Dim usedAlternate As Boolean = False
                                Dim localOutput As String = ""
                                Try
                                    ' (1) Alternate model application (moved from synchronous path)
                                    If useSecondApiLocal AndAlso Not String.IsNullOrWhiteSpace(selectedModelKeyLocal) Then
                                        Try
                                            originalConfig = GetCurrentConfig(_context) : originalConfigLoaded = True
                                            Dim alts = LoadAlternativeModels(INI_AlternateModelPath, _context)
                                            Dim sel = alts?.FirstOrDefault(
                                                Function(m)
                                                    If m Is Nothing Then Return False
                                                    If Not String.IsNullOrWhiteSpace(m.ModelDescription) AndAlso
                                                       String.Equals(m.ModelDescription, selectedModelKeyLocal, StringComparison.OrdinalIgnoreCase) Then Return True
                                                    If Not String.IsNullOrWhiteSpace(m.Model) AndAlso
                                                       String.Equals(m.Model, selectedModelKeyLocal, StringComparison.OrdinalIgnoreCase) Then Return True
                                                    Return False
                                                End Function)
                                            If sel IsNot Nothing Then
                                                ApplyModelConfig(_context, sel)
                                                usedAlternate = True
                                            End If
                                        Catch
                                        End Try
                                    End If

                                    ' (2) Run LLM
                                    localOutput = RunLlmAsync(sysPromptBase, sbDialog.ToString(),
                                                              useSecondApiLocal,
                                                              False,
                                                              finalFileObject,
                                                              jobCts.Token).GetAwaiter().GetResult()

                                    If localOutput Is Nothing Then localOutput = String.Empty
                                    localOutput = SanitizeModelOutputForBrowser(localOutput).Trim()

                                    ' Normalize cancellation marker
                                    If localOutput.Length > 0 AndAlso
                                       localOutput.Equals("Operation was canceled by the user.", StringComparison.OrdinalIgnoreCase) Then
                                        localOutput = "Aborted by user."
                                    End If

                                    ' (3) Build assistant turn or error turn
                                    Dim assistantText As String = localOutput
                                    Dim wasCanceled As Boolean = jobCts.IsCancellationRequested

                                    If assistantText.Length = 0 Then
                                        assistantText = If(wasCanceled,
                                                           "Aborted by user.",
                                                           "Error: The model did not provide a response.")
                                    End If

                                    Dim htmlOut = MarkdownToHtml(assistantText)

                                    ' Reload latest state (in case user switched chats meanwhile)
                                    Dim stJob = LoadInkyState()

                                    stJob.History.Add(New ChatTurn With {
                                        .Role = "assistant",
                                        .Markdown = assistantText,
                                        .Html = htmlOut,
                                        .Utc = Date.UtcNow
                                    })
                                    stJob.LastAssistantText = assistantText
                                    SaveInkyState(stJob)

                                    If wasCanceled AndAlso localOutput.Length = 0 Then
                                        tcs.TrySetCanceled()
                                    Else
                                        tcs.TrySetResult(assistantText)
                                    End If

                                Catch exOp As OperationCanceledException
                                    tcs.TrySetCanceled()
                                Catch ex As System.Exception
                                    tcs.TrySetException(ex)
                                Finally
                                    ' (4) Cleanup temp upload
                                    Try
                                        If Not String.IsNullOrWhiteSpace(tempUploadPathCopy) AndAlso IO.File.Exists(tempUploadPathCopy) Then
                                            IO.File.Delete(tempUploadPathCopy)
                                        End If
                                    Catch
                                    End Try
                                    ' (5) Restore original config if needed
                                    Try
                                        If useSecondApiLocal AndAlso (usedAlternate OrElse originalCfgLoaded) AndAlso originalConfigLoaded Then
                                            RestoreDefaults(_context, originalConfig)
                                            originalConfigLoaded = False
                                        End If
                                    Catch
                                    End Try
                                    Threading.Interlocked.Decrement(activeJobs)
                                End Try
                            End Sub)

                        ' ------------------ (F) Immediate response with job id ------------------
                        Return JsonOk(New With {
                            .ok = True,
                            .job = jobId,
                            .status = "running",
                            .history = ToBrowserTurns(st.History)
                        })

                    Case "inky_pure"
                        ' PURE MODE: send only raw user text (no history, no system prompt),
                        ' optionally with a binary file object (no text extraction).
                        Dim fileObject As String = j("FileObject")?.ToString()
                        Dim textBody As String = j("Text")?.ToString()

                        If String.IsNullOrWhiteSpace(textBody) Then
                            Return JsonErr("Please enter a message.")
                        End If

                        Dim stPure As InkyState = LoadInkyState()

                        ' Decide model context (still respect selected model choice)
                        Dim useSecondApiLocal As Boolean = stPure.UseSecondApi
                        Dim selectedModelKeyLocal As String = stPure.SelectedModelKey

                        ' Validate file object support (only pass through if supported; no extraction)
                        Dim supportsFilesNow As Boolean = False
                        Try
                            supportsFilesNow = ComputeSupportsFiles(useSecondApiLocal, selectedModelKeyLocal)
                        Catch
                            supportsFilesNow = False
                        End Try
                        If Not supportsFilesNow Then
                            fileObject = Nothing
                        End If

                        ' Record user turn (prefixed)
                        stPure.History.Add(New ChatTurn With {
                            .Role = "user",
                            .Markdown = "Pure: " & textBody,
                            .Html = MarkdownToHtml("Pure: " & textBody),
                            .Utc = Date.UtcNow
                        })
                        SaveInkyState(stPure)

                        ' Prepare background job (mirrors inky_send logic but no dialog/system prompt)
                        Dim jobIdP As String = Guid.NewGuid().ToString("N")
                        Dim jobCtsP As New CancellationTokenSource()
                        Dim tcsP As New TaskCompletionSource(Of String)(TaskCreationOptions.RunContinuationsAsynchronously)

                        Dim jobP As New LlmJob With {
                            .Id = jobIdP,
                            .CreatedUtc = Date.UtcNow,
                            .Tcs = tcsP,
                            .Cts = jobCtsP,
                            .UseSecond = useSecondApiLocal,
                            .FileObject = fileObject
                        }

                        If Not jobMap.TryAdd(jobIdP, jobP) Then
                            jobCtsP.Dispose()
                            Return JsonErr("Failed to register job.")
                        End If
                        Threading.Interlocked.Increment(activeJobs)

                        ' Maintain legacy cancellation
                        Try
                            If llmOperationCts IsNot Nothing Then
                                llmOperationCts.Cancel()
                                llmOperationCts.Dispose()
                            End If
                        Catch
                        End Try
                        llmOperationCts = jobCtsP

                        System.Threading.Tasks.Task.Run(
                            Sub()
                                Dim originalCfgLoadedP As Boolean = False
                                Dim usedAlternate As Boolean = False
                                Try
                                    ' Apply alternate model if a specific alternate is selected
                                    If useSecondApiLocal AndAlso Not String.IsNullOrWhiteSpace(selectedModelKeyLocal) Then
                                        Try
                                            originalConfig = GetCurrentConfig(_context) : originalCfgLoadedP = True
                                            Dim alts = LoadAlternativeModels(INI_AlternateModelPath, _context)
                                            Dim sel = alts?.FirstOrDefault(
                                                Function(m)
                                                    If m Is Nothing Then Return False
                                                    If Not String.IsNullOrWhiteSpace(m.ModelDescription) AndAlso
                                                       String.Equals(m.ModelDescription, selectedModelKeyLocal, StringComparison.OrdinalIgnoreCase) Then Return True
                                                    If Not String.IsNullOrWhiteSpace(m.Model) AndAlso
                                                       String.Equals(m.Model, selectedModelKeyLocal, StringComparison.OrdinalIgnoreCase) Then Return True
                                                    Return False
                                                End Function)
                                            If sel IsNot Nothing Then
                                                ApplyModelConfig(_context, sel)
                                                usedAlternate = True
                                            End If
                                        Catch
                                        End Try
                                    End If

                                    ' Raw call: NO system prompt, NO history packaging
                                    Dim output = RunLlmAsync("", textBody, useSecondApiLocal, False, fileObject, jobCtsP.Token).GetAwaiter().GetResult()
                                    If output Is Nothing Then output = String.Empty
                                    output = SanitizeModelOutputForBrowser(output).Trim()
                                    If jobCtsP.IsCancellationRequested AndAlso output.Length = 0 Then
                                        tcsP.TrySetCanceled()
                                    Else
                                        If output.Length = 0 Then
                                            output = If(jobCtsP.IsCancellationRequested, "Aborted by user.", "Error: The model returned no content.")
                                        End If
                                        Dim stFin = LoadInkyState()
                                        stFin.History.Add(New ChatTurn With {
                                            .Role = "assistant",
                                            .Markdown = output,
                                            .Html = MarkdownToHtml(output),
                                            .Utc = Date.UtcNow
                                        })
                                        stFin.LastAssistantText = output
                                        SaveInkyState(stFin)
                                        tcsP.TrySetResult(output)
                                    End If
                                Catch exOp As OperationCanceledException
                                    tcsP.TrySetCanceled()
                                Catch ex As System.Exception
                                    tcsP.TrySetException(ex)
                                Finally
                                    ' Restore config if alternate used
                                    Try
                                        If useSecondApiLocal AndAlso (usedAlternate OrElse originalCfgLoadedP) AndAlso originalConfigLoaded Then
                                            RestoreDefaults(_context, originalConfig)
                                            originalConfigLoaded = False
                                        End If
                                    Catch
                                    End Try
                                    Threading.Interlocked.Decrement(activeJobs)
                                End Try
                            End Sub)

                        Return JsonOk(New With {
                            .ok = True,
                            .job = jobIdP,
                            .status = "running",
                            .history = ToBrowserTurns(stPure.History)
                        })

                    Case "inky_jobstatus"
                        Dim jobId = j("Job")?.ToString()
                        If String.IsNullOrWhiteSpace(jobId) Then Return JsonErr("Missing Job id.")
                        Dim job As LlmJob = Nothing
                        If Not jobMap.TryGetValue(jobId, job) Then
                            Return JsonErr("Unknown job.")
                        End If
                        ' Cleanup old jobs (lazy TTL)
                        If (Date.UtcNow - job.CreatedUtc).TotalMinutes > JobTtlMinutes Then
                            Dim dump As LlmJob = Nothing
                            jobMap.TryRemove(jobId, dump)
                            Return JsonErr("Job expired.")
                        End If
                        Dim t = job.Tcs.Task
                        If Not t.IsCompleted Then
                            Return JsonOk(New With {.ok = True, .job = jobId, .status = "running"})
                        End If
                        If t.IsCanceled Then
                            Return JsonOk(New With {.ok = True, .job = jobId, .status = "canceled"})
                        ElseIf t.IsFaulted Then
                            Return JsonOk(New With {.ok = False, .job = jobId, .status = "error", .error = t.Exception.GetBaseException().Message})
                        Else
                            Return JsonOk(New With {.ok = True, .job = jobId, .status = "done", .result = t.Result})
                        End If

                    Case "inky_canceljob"
                        Dim jobId = j("Job")?.ToString()
                        If String.IsNullOrWhiteSpace(jobId) Then Return JsonErr("Missing Job id.")
                        Dim job As LlmJob = Nothing
                        If Not jobMap.TryGetValue(jobId, job) Then Return JsonErr("Unknown job.")
                        Try
                            job.Cts.Cancel()
                            Return JsonOk(New With {.ok = True, .job = jobId, .status = "cancelRequested"})
                        Catch ex As System.Exception
                            Return JsonErr("Cancel failed: " & ex.Message)
                        End Try


                    Case "inky_clear"
                        Dim stClear As New InkyState()
                        SaveInkyState(stClear)
                        Return JsonOk(New With {
                            .ok = True,
                            .activeChat = activeChatId,
                            .greeting = GetFriendlyGreeting()
                        })

                    Case "inky_copylast"
                        Dim stCopy As InkyState = LoadInkyState()
                        If System.String.IsNullOrWhiteSpace(stCopy.LastAssistantText) Then
                            Return JsonErr("No assistant response available to copy.")
                        End If

                        ' Synchronous wait on UI switch to avoid Await in environments that warn here                   
                        ' SwitchToUi(Sub() SLib.PutInClipboard(MarkdownToRtfConverter.Convert(stCopy.LastAssistantText)) End Sub).Wait()

                        Await SwitchToUi(Sub()
                                             SLib.PutInClipboard(MarkdownToRtfConverter.Convert(stCopy.LastAssistantText))
                                         End Sub).ConfigureAwait(False)

                        Return JsonOk(New With {.ok = True})


                    Case "inky_setmodel"
                        Dim key As System.String = j("Key")?.ToString()
                        Dim st As InkyState = LoadInkyState()

                        If System.String.IsNullOrWhiteSpace(key) OrElse System.String.Equals(key, "default", System.StringComparison.OrdinalIgnoreCase) Then
                            st.UseSecondApi = False
                            st.SelectedModelKey = ""
                        ElseIf System.String.Equals(key, "__second__", System.StringComparison.OrdinalIgnoreCase) Then
                            st.UseSecondApi = True
                            st.SelectedModelKey = ""
                        Else
                            st.UseSecondApi = True
                            st.SelectedModelKey = key
                        End If

                        Try
                            My.Settings.Inky_UseSecondApiSelected = st.UseSecondApi
                            My.Settings.Inky_SelectedModelKey = st.SelectedModelKey
                            My.Settings.Save()
                        Catch
                        End Try

                        ' Re-evaluate upload capability for selected model
                        Try
                            st.SupportsFileUploads = ComputeSupportsFiles(st.UseSecondApi, st.SelectedModelKey)
                        Catch
                            st.SupportsFileUploads = False
                        End Try
                        SaveInkyState(st)

                        Return JsonOk(New With {.ok = True, .supportsFiles = st.SupportsFileUploads, .activeChat = activeChatId})



                    Case "inky_toggletheme"
                        Dim st As InkyState = LoadInkyState()
                        st.DarkMode = Not st.DarkMode
                        SaveInkyState(st)
                        Try
                            My.Settings.Inky_DarkMode = st.DarkMode
                            My.Settings.Save()
                        Catch
                        End Try
                        Return JsonOk(New With {.ok = True, .darkMode = st.DarkMode, .activeChat = activeChatId})

                    Case Else
                        Return JsonErr("Unknown command.")
                End Select

            Catch ex As System.Exception
                Return JsonErr("Bad request: " & ex.Message)
            End Try
        End If

        ' ---- FALLBACK to your existing command dispatcher (unchanged) ----
        ' (Your original Select Case ... from earlier)
        ' NOTE: keep all your existing cases here. Below is your original body:

        Dim j0 = Newtonsoft.Json.Linq.JObject.Parse(If(body, "{}"))
        Dim cmd0 = j0("Command")?.ToString()
        Dim textBody0 = j0("Text")?.ToString()
        Dim sourceUrl = j0("URL")?.ToString()

        Select Case cmd0

            Case "redink_sendtooutlook"
                If String.IsNullOrWhiteSpace(textBody0) Then Return ""
                ' All Outlook automation on UI thread
                Await SwitchToUi(Sub()
                                     Dim olApp As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application
                                     Dim insp As Microsoft.Office.Interop.Outlook.Inspector = ComRetry(Function() olApp.ActiveInspector())
                                     If insp Is Nothing Then Exit Sub

                                     ' Guard CurrentItem (never access inline)
                                     Dim curr As Object = Nothing
                                     Try
                                         curr = ComRetry(Function() insp.CurrentItem)
                                     Catch
                                         curr = Nothing
                                     End Try
                                     If curr Is Nothing OrElse Not TypeOf curr Is Microsoft.Office.Interop.Outlook.MailItem Then
                                         Exit Sub
                                     End If

                                     Dim mail As Microsoft.Office.Interop.Outlook.MailItem =
                                         CType(curr, Microsoft.Office.Interop.Outlook.MailItem)

                                     ' Guard Sent property
                                     Try
                                         If ComRetry(Function() mail.Sent) Then
                                             If insp IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(insp) : insp = Nothing
                                             Exit Sub
                                         End If
                                     Catch
                                         If insp IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(insp) : insp = Nothing
                                         Exit Sub
                                     End Try

                                     ' Guard WordEditor and selection
                                     Dim doc As Microsoft.Office.Interop.Word.Document = Nothing
                                     Try
                                         doc = ComRetry(Function() CType(insp.WordEditor, Microsoft.Office.Interop.Word.Document))
                                         If doc Is Nothing Then Exit Sub

                                         Dim rng As Microsoft.Office.Interop.Word.Range = Nothing
                                         Try
                                             rng = doc.Application.Selection.Range
                                             doc.Application.ScreenUpdating = False
                                             rng.Text = textBody0 & " (" & sourceUrl & ")"
                                             doc.Application.ScreenUpdating = True
                                         Finally
                                             If rng IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(rng) : rng = Nothing
                                         End Try
                                     Finally
                                         If doc IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(doc) : doc = Nothing
                                         If insp IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(insp) : insp = Nothing
                                     End Try
                                 End Sub)
                Return ""
        ' -------------------------------------------------------------------
            Case "redink_translate"
                ' ─── 1  guard clauses ─────────────────────────────────────────
                If String.IsNullOrWhiteSpace(textBody0) Then Return ""

                ' Ask the user for a target language (UI thread)
                Dim targetLang As String = Await SwitchToUi(Function()
                                                                Return SLib.ShowCustomInputBox(
                       "Enter your target language:",
                       AN & " Translate (for Browser)",
                       True, INI_Language1)
                                                            End Function)

                If String.IsNullOrWhiteSpace(targetLang) OrElse targetLang = "ESC" Then
                    Return ""                                   ' user cancelled
                End If

                TranslateLanguage = targetLang.Trim()

                ' ─── 2  call the LLM on the UI thread, get Task(Of String) ─────
                Dim llmOut As String = Await RunLlmAsync(
                        InterpolateAtRuntime(SP_Translate),
                        $"<TEXTTOPROCESS>{textBody0}</TEXTTOPROCESS>")

                ' ─── 3  clean up the wrapper tags / markdown ──────────────────
                llmOut = llmOut.Replace("<TEXTTOPROCESS>", "") _
                       .Replace("</TEXTTOPROCESS>", "") _
                       .Replace("**", "").Trim()

                If llmOut = "" Then Return ""                  ' safety net

                ' Optional: copy to clipboard so the user can paste manually
                Await SwitchToUi(Sub() SLib.PutInClipboard(llmOut))

                ' ─── 4  SEND the translation back to the caller ───────────────
                Return llmOut

            ' -------------------------------------------------------------------
            Case "redink_correct"

                If String.IsNullOrWhiteSpace(textBody0) Then Return ""

                ' 1)  Run the LLM on the UI thread
                Dim llmOut As String = Await RunLlmAsync(
                                                        InterpolateAtRuntime(SP_Correct),
                                                        $"<TEXTTOPROCESS>{textBody0}</TEXTTOPROCESS>")
                llmOut = llmOut.Replace("<TEXTTOPROCESS>", "").Replace("</TEXTTOPROCESS>", "")

                If llmOut = "" Then Return ""

                ' 2)  Show the compare / preview window (synchronous)
                Await SwitchToUi(Sub()
                                     CompareAndInsertText(textBody0, llmOut, True)
                                 End Sub)

                ' 3)  
                Dim accepted As Boolean = Await WaitForPreviewDecisionAsync()

                If Not accepted Then Return ""          ' Esc pressed → abort

                Return llmOut

        ' -------------------------------------------------------------------
            Case "redink_freestyle"

                '─── A  gather prompt on UI thread ──────────────────────────────
                Dim noText As Boolean = String.IsNullOrWhiteSpace(textBody0)

                Dim promptCaption As String = AN & " Freestyle (for Browser)"
                Dim wordInstalled As Boolean = False
                Try
                    Dim wordApp As Object = CreateObject("Word.Application")
                    wordInstalled = True
                    wordApp.Quit()
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp)
                Catch ex As System.Exception
                    wordInstalled = False
                End Try

                Dim sb As New System.Text.StringBuilder()
                If noText Then
                    sb.Append("Please provide the prompt you wish to execute ")
                Else
                    sb.Append("Please provide the prompt you wish to execute using the selected text ")
                End If

                sb.Append("('" & MarkupPrefix & "' for markups, '" & InsertPrefix & "' for direct insert" & If(wordInstalled, " and '" & NewDocPrefix & "' to put the output in a new Word document)", ")"))
                If INI_PromptLib Then sb.Append(" or press 'OK' for the prompt library")
                If INI_SecondAPI Then sb.Append($"; add '{SecondAPICode}' to use {If(String.IsNullOrWhiteSpace(INI_AlternateModelPath), $"the secondary model ({INI_Model_2})", "one of the other models")}")

                If Not String.IsNullOrWhiteSpace(My.Settings.LastPrompt) Then sb.Append("; ctrl-p for your last prompt")
                sb.Append(":")
                Dim promptMsg As String = sb.ToString()

                Dim OptionalButtons As System.Tuple(Of String, String, String)()
                If wordInstalled Then
                    OptionalButtons = {
                                System.Tuple.Create("OK, do a new doc", $"Use this to automatically insert '{NewDocPrefix}' as a prefix.", NewDocPrefix)
                            }
                End If

                OtherPrompt = Await SwitchToUi(Function()
                                                   Return SLib.ShowCustomInputBox(promptMsg, promptCaption, False, "", My.Settings.LastPrompt, If(wordInstalled, OptionalButtons, Nothing))
                                               End Function)

                Dim doMarkupFlag As Boolean = False
                Dim doInsertFlag As Boolean = False
                Dim UseSecondAPI As Boolean = False
                Dim DoNewDoc As Boolean = False

                '─── prompt library branch ─────────────────────────────────────
                If String.IsNullOrEmpty(OtherPrompt) AndAlso OtherPrompt <> "ESC" AndAlso INI_PromptLib Then
                    Dim sel = Await SwitchToUi(Function()
                                                   Return ShowPromptSelector(INI_PromptLibPath, Not noText, Nothing)
                                               End Function)                         ' (prompt, doMarkup, doInsert, canceled)

                    OtherPrompt = sel.Item1
                    doMarkupFlag = sel.Item2
                    doInsertFlag = Not sel.Item4         ' library’s “canceled” → insert = False
                End If

                ' user cancelled
                If String.IsNullOrWhiteSpace(OtherPrompt) OrElse OtherPrompt = "ESC" Then
                    Return ""
                End If

                ' remember last prompt
                My.Settings.LastPrompt = OtherPrompt
                My.Settings.Save()

                '─── decode prefix flags ───────────────────────────────────────
                If OtherPrompt.StartsWith(InsertPrefix, StringComparison.OrdinalIgnoreCase) Then
                    OtherPrompt = OtherPrompt.Substring(InsertPrefix.Length).Trim()
                    doInsertFlag = True
                ElseIf OtherPrompt.StartsWith(MarkupPrefix, StringComparison.OrdinalIgnoreCase) AndAlso Not noText Then
                    OtherPrompt = OtherPrompt.Substring(MarkupPrefix.Length).Trim()
                    doMarkupFlag = True
                    doInsertFlag = True          ' old logic: markup implies insert
                ElseIf OtherPrompt.StartsWith(NewDocPrefix, StringComparison.OrdinalIgnoreCase) AndAlso Not noText Then
                    OtherPrompt = OtherPrompt.Substring(NewDocPrefix.Length).Trim()
                    DoNewDoc = True
                    doMarkupFlag = False
                End If

                If INI_SecondAPI Then
                    If OtherPrompt.Contains(SecondAPICode) Then
                        UseSecondAPI = True
                        OtherPrompt = OtherPrompt.Replace(SecondAPICode, "").Trim()

                        If Not String.IsNullOrWhiteSpace(INI_AlternateModelPath) Then

                            Dim sel = Await SwitchToUi(Function()
                                                           Return Not ShowModelSelection(_context, INI_AlternateModelPath)
                                                       End Function)                         ' (prompt, doMarkup, doInsert, canceled)
                            If sel Then
                                originalConfigLoaded = False
                                Return ""
                            End If

                        End If

                    End If
                End If

                '─── B  call the LLM on UI thread (async) ──────────────────────
                Dim llmResult As String
                If noText Then
                    llmResult = Await RunLlmAsync(InterpolateAtRuntime(SP_FreestyleNoText), "", UseSecondAPI)
                Else
                    llmResult = Await RunLlmAsync(InterpolateAtRuntime(SP_FreestyleText), $"<TEXTTOPROCESS>{textBody0}</TEXTTOPROCESS>", UseSecondAPI)
                End If

                llmResult = llmResult.Replace("<TEXTTOPROCESS>", "") _
                             .Replace("</TEXTTOPROCESS>", "") _
                             .Trim()

                If String.IsNullOrEmpty(llmResult) Then Return ""

                '─── C  present / insert / clipboard exactly like old code ─────

                ' A) markup path (implies insert)  -----------------------------
                If doMarkupFlag Then
                    Await SwitchToUi(Sub()
                                         CompareAndInsertText(textBody0, llmResult, True)
                                     End Sub)

                    Dim accepted As Boolean = Await WaitForPreviewDecisionAsync()

                    If Not accepted Then Return ""          ' Esc pressed → abort

                    Return llmResult                    ' user accepted
                End If

                ' B) plain insert path  ----------------------------------------
                If doInsertFlag Then
                    'Await InsertTextIntoCurrentMailAsync(llmResult)
                    Return llmResult                        ' send text back
                End If

                ' Replace the original block inside Case "redink_freestyle" with this updated version:

                If DoNewDoc And wordInstalled Then
                    If Await TryCreateWordDocFromMarkdown(llmResult) Then
                        Return ""
                    End If
                    Await SwitchToUi(Sub()
                                         ShowCustomMessageBox("Could not create new Word document and insert the LLM output; providing your output to a separate window.")
                                     End Sub)
                End If

                ' C) clipboard-only path  --------------------------------------
                Dim finalTxt As String = Await SwitchToUi(Function()
                                                              Return SLib.ShowCustomWindow(
                                                                  "The LLM has provided the following result (you can edit it):",
                                                                  llmResult,
                                                                  "You can choose whether you want to have the original text put into the clipboard or your text with any changes you have made (without formatting)." & If(wordInstalled, " If you choose to insert the original text with formatting, a new Word document will be created with it. ", " ") & "If you select Cancel, nothing will be put into the clipboard (you can yourself copy it to the clipboard).",
                                                                  AN, False, True, If(wordInstalled, True, False))
                                                          End Function)

                ' user chose to create a Word doc now (sentinel "Markdown")
                If String.Equals(finalTxt, "Markdown", StringComparison.OrdinalIgnoreCase) Then
                    If wordInstalled AndAlso Await TryCreateWordDocFromMarkdown(llmResult) Then
                        Return ""
                    Else
                        Await SwitchToUi(Sub()
                                             ShowCustomMessageBox("Could not create new Word document and insert the LLM output (however, it will be in the clipboard).")
                                             finalTxt = MarkdownToRtfConverter.Convert(llmResult)  ' fall back to clipboard
                                         End Sub)
                    End If
                End If

                If Not String.IsNullOrWhiteSpace(finalTxt) Then
                    Await SwitchToUi(Sub() SLib.PutInClipboard(finalTxt))
                End If

                Return ""

        End Select

        Return ""
    End Function

    ' Add this helper anywhere inside the ThisAddIn class (e.g. near other small helpers)
    Private Async Function TryCreateWordDocFromMarkdown(markdown As String) As Task(Of Boolean)
        Try
            Dim wordApp As New Microsoft.Office.Interop.Word.Application()
            wordApp.Visible = True
            Dim newDoc As Microsoft.Office.Interop.Word.Document = wordApp.Documents.Add()
            Dim docSelection As Microsoft.Office.Interop.Word.Selection = wordApp.Selection
            InsertTextWithMarkdown(docSelection, markdown, True)
            Await SwitchToUi(Sub()
                                 ShowCustomMessageBox("Your Word document has been created. It may be hidden behind the other windows.")
                             End Sub)
            Return True
        Catch
            Return False
        End Try
    End Function


    ' Provide model list for the browser dropdown (default, second, alternates)   
    Private Async Function GetModelListForBrowserAsync(ByVal st As InkyState) _
        As System.Threading.Tasks.Task(Of System.Collections.Generic.List(Of Object))

        Dim list As New System.Collections.Generic.List(Of Object)()

        ' --- 1) Load persisted preference FIRST and reconcile with current state ---
        Try
            Dim savedSecond As Boolean = My.Settings.Inky_UseSecondApiSelected
            Dim savedKey As String = My.Settings.Inky_SelectedModelKey

            ' Only adopt saved values if they are consistent / meaningful
            Dim shouldApply As Boolean = False

            If savedSecond <> st.UseSecondApi Then
                shouldApply = True
            ElseIf savedSecond AndAlso Not String.Equals(savedKey, st.SelectedModelKey, StringComparison.OrdinalIgnoreCase) Then
                shouldApply = True
            End If

            If shouldApply Then
                st.UseSecondApi = savedSecond
                st.SelectedModelKey = If(savedSecond, savedKey, "")
                SaveInkyState(st)
            End If
        Catch
            ' Ignore; keep in-memory state
        End Try

        ' --- 2) Gather availability info based on (potentially updated) state ---
        Dim hasPrimary As Boolean = Not String.IsNullOrWhiteSpace(INI_Model)
        Dim hasSecondApi As Boolean = INI_SecondAPI
        Dim hasSecondModelName As Boolean = Not String.IsNullOrWhiteSpace(INI_Model_2)
        Dim hasSecondary As Boolean = hasSecondApi AndAlso hasSecondModelName

        Dim alts As System.Collections.Generic.List(Of SharedLibrary.SharedLibrary.ModelConfig) = Nothing
        Dim altCount As Integer = 0
        Try
            If hasSecondApi AndAlso Not String.IsNullOrWhiteSpace(INI_AlternateModelPath) Then
                alts = LoadAlternativeModels(INI_AlternateModelPath, _context)
                If alts IsNot Nothing Then altCount = alts.Count
            End If
        Catch
            altCount = 0
            alts = Nothing
        End Try

        ' --- 3) If saved SelectedModelKey no longer exists, normalize state ---
        If st.UseSecondApi AndAlso Not String.IsNullOrWhiteSpace(st.SelectedModelKey) Then
            Dim exists As Boolean = False
            If alts IsNot Nothing Then
                exists = alts.Any(Function(m)
                                      If m Is Nothing Then Return False
                                      Dim label = If(Not String.IsNullOrWhiteSpace(m.ModelDescription), m.ModelDescription, m.Model)
                                      Return String.Equals(label, st.SelectedModelKey, StringComparison.OrdinalIgnoreCase)
                                  End Function)
            End If
            If Not exists Then
                ' Fallback to second default (still second API)
                st.SelectedModelKey = ""
                SaveInkyState(st)
            End If
        End If

        ' --- 4) Simple case: only primary model available ---
        If hasPrimary AndAlso Not hasSecondary AndAlso altCount = 0 Then
            list.Add(New With {
                .key = "default",
                .label = INI_Model,
                .selected = (Not st.UseSecondApi),
                .disabled = False,
                .isSeparator = False
            })
            Return list
        End If

        ' --- 5) Build list with correct selection flags ---
        If hasPrimary Then
            list.Add(New With {
                .key = "__hdr_primary__",
                .label = "Primary model:",
                .selected = False,
                .disabled = True,
                .isSeparator = False
            })
            list.Add(New With {
                .key = "default",
                .label = INI_Model,
                .selected = (Not st.UseSecondApi),
                .disabled = False,
                .isSeparator = False
            })
        End If

        If hasSecondary Then
            list.Add(New With {
                .key = "__hdr_secondary__",
                .label = "Secondary model:",
                .selected = False,
                .disabled = True,
                .isSeparator = False
            })
            list.Add(New With {
                .key = "__second__",
                .label = INI_Model_2,
                .selected = (st.UseSecondApi AndAlso String.IsNullOrWhiteSpace(st.SelectedModelKey)),
                .disabled = False,
                .isSeparator = False
            })
        End If

        If altCount > 0 AndAlso alts IsNot Nothing Then
            list.Add(New With {
                .key = "__sep__",
                .label = "Alternative models:",
                .selected = False,
                .disabled = True,
                .isSeparator = True
            })

            For Each m In alts
                If m Is Nothing Then Continue For
                Dim label = If(Not String.IsNullOrWhiteSpace(m.ModelDescription), m.ModelDescription, m.Model)
                If String.IsNullOrWhiteSpace(label) Then label = "Model"
                list.Add(New With {
                    .key = label,
                    .label = label,
                    .selected = (st.UseSecondApi AndAlso
                                 String.Equals(st.SelectedModelKey, label, StringComparison.OrdinalIgnoreCase)),
                    .disabled = False,
                    .isSeparator = False
                })
            Next
        End If

        Return list
    End Function




    ' Entfernt führende Rollenmarker wie [ASSISTANT], [USER] oder "ASSISTANT:" am Zeilenanfang.
    Private Function SanitizeModelOutputForBrowser(ByVal raw As System.String) As System.String
        If raw Is Nothing Then Return System.String.Empty

        Dim s As System.String = raw

        ' 1) Vollständige Rollen-Zeilen weg (z. B. nur "[ASSISTANT]" oder "[USER]:")
        s = System.Text.RegularExpressions.Regex.Replace(
            s,
            "(?im)^\s*\[(?:assistant|user)\]\s*:?\s*$",
            "",
            System.Text.RegularExpressions.RegexOptions.None)

        ' 2) Rollenmarker am Zeilenanfang entfernen (belässt den eigentlichen Text)
        s = System.Text.RegularExpressions.Regex.Replace(
            s,
            "(?im)^\s*\[(?:assistant|user)\]\s*",
            "",
            System.Text.RegularExpressions.RegexOptions.None)

        ' 3) Alternative Schreibweise "ASSISTANT:" / "USER:" am Zeilenanfang entfernen
        '    (nur am Zeilenanfang, um normalen Fließtext nicht zu beschädigen)
        s = System.Text.RegularExpressions.Regex.Replace(
            s,
            "(?im)^\s*(?:assistant|user)\s*:\s*",
            "",
            System.Text.RegularExpressions.RegexOptions.None)

        ' 4) Überzählige Leerzeilen normalisieren (optional)
        s = System.Text.RegularExpressions.Regex.Replace(
            s,
            "(\r?\n){3,}",
            System.Environment.NewLine & System.Environment.NewLine,
            System.Text.RegularExpressions.RegexOptions.None)

        Return s
    End Function


    ' Maps ChatTurn → browser DTO (camelCase)
    Private Function ToBrowserTurns(list As System.Collections.Generic.List(Of ChatTurn)) _
        As System.Collections.Generic.List(Of Object)

        Dim out As New System.Collections.Generic.List(Of Object)()
        For Each t In list
            out.Add(New With {
            .role = t.Role,
            .markdown = t.Markdown,
            .html = t.Html,
            .utc = t.Utc
        })
        Next
        Return out
    End Function

    <System.Runtime.InteropServices.DllImport("user32.dll")>
    Private Shared Function GetGuiResources(hProcess As System.IntPtr, uiFlags As System.Int32) As System.UInt32
    End Function
    Private Shared Function GetGdiCount() As System.UInt32
        Return GetGuiResources(System.Diagnostics.Process.GetCurrentProcess().Handle, 0UI)
    End Function
    Private Shared Function GetUserCount() As System.UInt32
        Return GetGuiResources(System.Diagnostics.Process.GetCurrentProcess().Handle, 1UI)
    End Function

    '──────────────────────────────────────────────────────────────────────────────
    ' Office → Plaintext
    '──────────────────────────────────────────────────────────────────────────────
    Private Function TryExtractOfficeText(
    ByVal filePath As System.String,
    ByRef extracted As System.String,
    ByRef label As System.String
) As System.Boolean

        extracted = Nothing
        label = Nothing

        If System.String.IsNullOrWhiteSpace(filePath) Then Return False
        If Not System.IO.File.Exists(filePath) Then Return False

        Dim ext As System.String = System.IO.Path.GetExtension(filePath).ToLowerInvariant()

        Try
            Select Case ext
                Case ".doc", ".docx", ".rtf"
                    extracted = ExtractWordText(filePath)
                    label = "Word document: " & System.IO.Path.GetFileName(filePath)
                Case ".xls", ".xlsx"
                    extracted = ExtractExcelText(filePath)
                    label = "Excel workbook: " & System.IO.Path.GetFileName(filePath)
                Case ".ppt", ".pptx"
                    extracted = ExtractPowerPointText(filePath)
                    label = "PowerPoint presentation: " & System.IO.Path.GetFileName(filePath)
                Case Else
                    Return False
            End Select
        Catch ex As System.Exception
            ' Optional: Loggen
            System.Diagnostics.Debug.WriteLine("Office extract failed: " & ex.Message)
            extracted = Nothing
            label = Nothing
            Return False
        End Try

        If System.String.IsNullOrWhiteSpace(extracted) Then Return False
        ' Soft-Limit (optional): bremst extrem große Arbeitsmappen
        If extracted.Length > 1_500_000 Then
            extracted = extracted.Substring(0, 1_500_000) & System.Environment.NewLine & "[…truncated…]"
        End If

        Return True
    End Function

    '──────────────────────────────────────────────────────────────────────────────
    ' WORD
    '──────────────────────────────────────────────────────────────────────────────
    Private Function ExtractWordText(ByVal path As System.String) As System.String
        Dim app As Microsoft.Office.Interop.Word.Application = Nothing
        Dim doc As Microsoft.Office.Interop.Word.Document = Nothing
        Try
            app = New Microsoft.Office.Interop.Word.Application()
            app.Visible = False
            doc = app.Documents.Open(FileName:=path, ReadOnly:=True, Visible:=False, AddToRecentFiles:=False)

            ' Volltext – simpel & robust
            Dim raw As System.String = doc.Content.Text

            ' Normalize line breaks
            raw = raw.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)
            raw = System.Text.RegularExpressions.Regex.Replace(raw, "[\f\v]+", vbLf)

            Return raw.Trim()
        Catch ex As System.Exception
            Throw
        Finally
            SafeCloseWord(doc, app)
        End Try
    End Function

    Private Sub SafeCloseWord(
    ByVal doc As Microsoft.Office.Interop.Word.Document,
    ByVal app As Microsoft.Office.Interop.Word.Application
)
        Try
            If doc IsNot Nothing Then
                Try : doc.Close(SaveChanges:=False) : Catch : End Try
                Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(doc) : Catch : End Try
            End If
        Finally
            If app IsNot Nothing Then
                Try : app.Quit(SaveChanges:=False) : Catch : End Try
                Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(app) : Catch : End Try
            End If
        End Try
    End Sub

    '──────────────────────────────────────────────────────────────────────────────
    ' EXCEL
    '──────────────────────────────────────────────────────────────────────────────
    Private Function ExtractExcelText(ByVal path As System.String) As System.String
        Dim app As Microsoft.Office.Interop.Excel.Application = Nothing
        Dim wb As Microsoft.Office.Interop.Excel.Workbook = Nothing
        Dim sb As New System.Text.StringBuilder(4096)

        Try
            app = New Microsoft.Office.Interop.Excel.Application()
            app.Visible = False
            wb = app.Workbooks.Open(Filename:=path, ReadOnly:=True, AddToMru:=False)

            For Each shObj As System.Object In wb.Worksheets
                Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                Try
                    ws = CType(shObj, Microsoft.Office.Interop.Excel.Worksheet)
                    Dim used As Microsoft.Office.Interop.Excel.Range = ws.UsedRange
                    If used Is Nothing Then Continue For

                    sb.AppendLine("=== Sheet: " & ws.Name & " ===")

                    Dim rows As System.Int32 = used.Rows.Count
                    Dim cols As System.Int32 = used.Columns.Count
                    Dim rowOffset As System.Int32 = used.Row      ' 1-basiert
                    Dim colOffset As System.Int32 = used.Column   ' 1-basiert

                    ' Schnellpfad: beide Arrays auf einmal holen
                    Dim dataValues As System.Object(,) = Nothing
                    Dim dataFormulas As System.Object(,) = Nothing
                    Try
                        dataValues = TryCast(used.Value2, System.Object(,))
                    Catch
                        dataValues = Nothing
                    End Try
                    Try
                        dataFormulas = TryCast(used.Formula, System.Object(,))
                    Catch
                        dataFormulas = Nothing
                    End Try

                    If dataValues IsNot Nothing AndAlso dataFormulas IsNot Nothing Then
                        Dim rL As System.Int32 = dataValues.GetLength(0)
                        Dim cL As System.Int32 = dataValues.GetLength(1)
                        For r As System.Int32 = 1 To rL
                            For c As System.Int32 = 1 To cL
                                Dim absRow As System.Int32 = rowOffset + r - 1
                                Dim absCol As System.Int32 = colOffset + c - 1
                                Dim addr As System.String = ColToLetters(absCol) & absRow.ToString(System.Globalization.CultureInfo.InvariantCulture)

                                Dim vObj As System.Object = dataValues(r, c)
                                Dim fObj As System.Object = dataFormulas(r, c)

                                Dim vStr As System.String = System.Convert.ToString(vObj, System.Globalization.CultureInfo.InvariantCulture)
                                Dim fStr As System.String = System.Convert.ToString(fObj, System.Globalization.CultureInfo.InvariantCulture)

                                ' Manche Zellen mit Konstante haben Formula="" oder Nothing
                                If fObj IsNot Nothing Then
                                    ' Excel liefert bei Konstanten oft den Wert statt einer Formel.
                                    ' Wenn die Formel identisch zum Wert aussieht (häufig leer), lassen wir sie leer.
                                End If

                                sb.Append(addr)
                                sb.Append(vbTab)
                                sb.Append("FORMULA:")
                                If Not System.String.IsNullOrEmpty(fStr) Then
                                    sb.Append("="c)
                                    sb.Append(fStr.TrimStart("="c))
                                End If
                                sb.Append(vbTab)
                                sb.Append("VALUE: ")
                                sb.AppendLine(If(vStr, ""))
                            Next
                        Next
                    Else
                        ' Fallback: Zell-für-Zell (langsamer, aber robust)
                        For r As System.Int32 = 1 To rows
                            For c As System.Int32 = 1 To cols
                                Dim cell As Microsoft.Office.Interop.Excel.Range = Nothing
                                Try
                                    cell = CType(used.Cells(r, c), Microsoft.Office.Interop.Excel.Range)

                                    Dim absRow As System.Int32 = rowOffset + r - 1
                                    Dim absCol As System.Int32 = colOffset + c - 1
                                    Dim addr As System.String = ColToLetters(absCol) & absRow.ToString(System.Globalization.CultureInfo.InvariantCulture)

                                    Dim vObj As System.Object = Nothing
                                    Dim fObj As System.Object = Nothing
                                    Try : vObj = cell.Value2 : Catch : vObj = Nothing : End Try
                                    Try : fObj = cell.Formula : Catch : fObj = Nothing : End Try

                                    Dim vStr As System.String = System.Convert.ToString(vObj, System.Globalization.CultureInfo.InvariantCulture)
                                    Dim fStr As System.String = System.Convert.ToString(fObj, System.Globalization.CultureInfo.InvariantCulture)

                                    sb.Append(addr)
                                    sb.Append(vbTab)
                                    sb.Append("FORMULA:")
                                    If Not System.String.IsNullOrEmpty(fStr) Then
                                        sb.Append("="c)
                                        sb.Append(fStr.TrimStart("="c))
                                    End If
                                    sb.Append(vbTab)
                                    sb.Append("VALUE: ")
                                    sb.AppendLine(If(vStr, ""))
                                Finally
                                    If cell IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(cell)
                                End Try
                            Next
                        Next
                    End If

                    If used IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(used)
                    sb.AppendLine()
                Finally
                    If ws IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(ws)
                End Try
            Next

            Return sb.ToString().Trim()
        Catch ex As System.Exception
            Throw
        Finally
            SafeCloseExcel(wb, app)
        End Try
    End Function

    ' A1-Spaltenbezeichner
    Private Function ColToLetters(ByVal col As System.Int32) As System.String
        ' col: 1-basiert (1=A, 27=AA, …)
        Dim n As System.Int32 = col
        Dim chars As New System.Text.StringBuilder()
        While n > 0
            n -= 1
            Dim ch As System.Char = System.Convert.ToChar((n Mod 26) + System.Convert.ToInt32("A"c))
            chars.Insert(0, ch)
            n \= 26
        End While
        Return chars.ToString()
    End Function


    Private Sub SafeCloseExcel(
    ByVal wb As Microsoft.Office.Interop.Excel.Workbook,
    ByVal app As Microsoft.Office.Interop.Excel.Application
)
        Try
            If wb IsNot Nothing Then
                Try : wb.Close(SaveChanges:=False) : Catch : End Try
                Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(wb) : Catch : End Try
            End If
        Finally
            If app IsNot Nothing Then
                Try : app.Quit() : Catch : End Try
                Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(app) : Catch : End Try
            End If
        End Try
    End Sub

    '──────────────────────────────────────────────────────────────────────────────
    ' POWERPOINT
    '──────────────────────────────────────────────────────────────────────────────
    Private Function ExtractPowerPointText(ByVal path As System.String) As System.String
        Dim app As System.Object = Nothing
        Dim pres As System.Object = Nothing
        Dim sb As New System.Text.StringBuilder(2048)

        Try
            ' Late binding: keine PIAs nötig
            app = Microsoft.VisualBasic.Interaction.CreateObject("PowerPoint.Application")

            ' Presentations.Open(FileName, ReadOnly, Untitled, WithWindow)
            ' Late bound: True/False als -1/0; hier 1=True, 0=False
            Dim presentations As System.Object = app.Presentations
            pres = presentations.Open(path, 1, 0, 0)

            Dim slideCount As System.Int32 = System.Convert.ToInt32(pres.Slides.Count, System.Globalization.CultureInfo.InvariantCulture)
            For i As System.Int32 = 1 To slideCount
                Dim sld As System.Object = pres.Slides(i)
                Try
                    sb.AppendLine("=== Slide " & i.ToString(System.Globalization.CultureInfo.InvariantCulture) & " ===")

                    Dim shapeCount As System.Int32 = System.Convert.ToInt32(sld.Shapes.Count, System.Globalization.CultureInfo.InvariantCulture)
                    For j As System.Int32 = 1 To shapeCount
                        Dim shp As System.Object = sld.Shapes(j)
                        Try
                            Dim hasTf As System.Boolean = False
                            Try
                                ' In Office-Interop: True = -1, False = 0
                                hasTf = (System.Convert.ToInt32(shp.HasTextFrame, System.Globalization.CultureInfo.InvariantCulture) <> 0) AndAlso
                                    (Not shp.TextFrame Is Nothing) AndAlso
                                    (System.Convert.ToInt32(shp.TextFrame.HasText, System.Globalization.CultureInfo.InvariantCulture) <> 0)
                            Catch
                                hasTf = False
                            End Try

                            If hasTf Then
                                Dim txt As System.String = System.Convert.ToString(shp.TextFrame.TextRange.Text, System.Globalization.CultureInfo.InvariantCulture)
                                If Not System.String.IsNullOrWhiteSpace(txt) Then
                                    sb.AppendLine(txt.Trim())
                                End If
                            End If
                        Finally
                            Try
                                If shp IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(shp)
                            Catch
                            End Try
                        End Try
                    Next

                    ' Notes (optional)
                    Try
                        Dim notesShapes As System.Object = sld.NotesPage.Shapes
                        Dim nCount As System.Int32 = System.Convert.ToInt32(notesShapes.Count, System.Globalization.CultureInfo.InvariantCulture)
                        For k As System.Int32 = 1 To nCount
                            Dim shp2 As System.Object = notesShapes(k)
                            Try
                                Dim hasTf2 As System.Boolean = False
                                Try
                                    hasTf2 = (System.Convert.ToInt32(shp2.HasTextFrame, System.Globalization.CultureInfo.InvariantCulture) <> 0) AndAlso
                                         (Not shp2.TextFrame Is Nothing) AndAlso
                                         (System.Convert.ToInt32(shp2.TextFrame.HasText, System.Globalization.CultureInfo.InvariantCulture) <> 0)
                                Catch
                                    hasTf2 = False
                                End Try
                                If hasTf2 Then
                                    Dim note As System.String = System.Convert.ToString(shp2.TextFrame.TextRange.Text, System.Globalization.CultureInfo.InvariantCulture)
                                    If Not System.String.IsNullOrWhiteSpace(note) Then
                                        sb.AppendLine("--- Notes ---")
                                        sb.AppendLine(note.Trim())
                                    End If
                                End If
                            Finally
                                Try
                                    If shp2 IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(shp2)
                                Catch
                                End Try
                            End Try
                        Next
                    Catch
                    End Try

                    sb.AppendLine()
                Finally
                    Try
                        If sld IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(sld)
                    Catch
                    End Try
                End Try
            Next

            Return sb.ToString().Trim()
        Catch ex As System.Exception
            Throw
        Finally
            Try
                If pres IsNot Nothing Then
                    Try : pres.Close() : Catch : End Try
                    Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(pres) : Catch : End Try
                End If
            Catch
            End Try
            Try
                If app IsNot Nothing Then
                    Try : app.Quit() : Catch : End Try
                    Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(app) : Catch : End Try
                End If
            Catch
            End Try
        End Try
    End Function



    Private Sub SafeClosePowerPoint(
    ByVal pres As Microsoft.Office.Interop.PowerPoint.Presentation,
    ByVal app As Microsoft.Office.Interop.PowerPoint.Application
)
        Try
            If pres IsNot Nothing Then
                Try : pres.Close() : Catch : End Try
                Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(pres) : Catch : End Try
            End If
        Finally
            If app IsNot Nothing Then
                Try : app.Quit() : Catch : End Try
                Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(app) : Catch : End Try
            End If
        End Try
    End Sub

    '──────────────────────────────────────────────────────────────────────────────
    ' XML-/Tag-sicher
    '──────────────────────────────────────────────────────────────────────────────
    Private Function EscapeForXml(ByVal s As System.String) As System.String
        If s Is Nothing Then Return ""
        Return System.Security.SecurityElement.Escape(s)
    End Function

    Private Function TryExtractTextLike(
    ByVal filePath As System.String,
    ByRef extracted As System.String,
    ByRef label As System.String
) As System.Boolean

        extracted = Nothing
        label = Nothing

        If System.String.IsNullOrWhiteSpace(filePath) Then Return False
        If Not System.IO.File.Exists(filePath) Then Return False

        Dim ext As System.String = System.IO.Path.GetExtension(filePath).ToLowerInvariant()

        ' Liste gängiger Text-/Code-Endungen (erweiterbar)
        Dim textLike As System.String() = {
        ".txt", ".log", ".csv", ".tsv", ".md",
        ".json", ".xml", ".yaml", ".yml", ".ini", ".cfg", ".conf", ".toml",
        ".sql",
        ".cs", ".vb", ".vbs", ".js", ".ts", ".jsx", ".tsx",
        ".py", ".rb", ".php", ".java", ".kt", ".kts",
        ".c", ".h", ".hpp", ".hh", ".cpp", ".cc",
        ".ps1", ".psm1", ".bat", ".cmd", ".sh", ".zsh",
        ".rtf" ' Hinweis: RTF könnte man auch via Word-Interop extrahieren – hier als Text belassen
    }

        If Not textLike.Contains(ext) Then
            Return False
        End If

        Try
            ' RTF optional als Office behandeln (falls du lieber echtes Plaintext-RTF willst, nimm Word-Interop)
            If ext = ".rtf" Then
                Try
                    Dim tmp As System.String = ExtractWordText(filePath) ' nutzt Word-Interop, wenn vorhanden
                    If Not System.String.IsNullOrWhiteSpace(tmp) Then
                        extracted = tmp
                        label = "Word-readable (RTF): " & System.IO.Path.GetFileName(filePath)
                        If extracted.Length > 1_500_000 Then
                            extracted = extracted.Substring(0, 1_500_000) & System.Environment.NewLine & "[…truncated…]"
                        End If
                        Return True
                    End If
                Catch
                    ' Fallback: als Text lesen
                End Try
            End If

            Dim content As System.String = ReadAllTextSmart(filePath)
            If System.String.IsNullOrWhiteSpace(content) Then Return False

            ' Für CSV/TSV eine kleine Kopfzeile hinzufügen
            If ext = ".csv" OrElse ext = ".tsv" Then
                Dim sep As System.String = If(ext = ".csv", ",", vbTab)
                Dim header As System.String = "=== CSV/TSV Detected (" & ext.Trim("."c).ToUpperInvariant() & ", sep=""" & If(ext = ".csv", ",", "\t") & """) ==="
                extracted = header & System.Environment.NewLine & content
                label = "Spreadsheet text: " & System.IO.Path.GetFileName(filePath)
            Else
                extracted = content
                label = "Text/code file: " & System.IO.Path.GetFileName(filePath)
            End If

            If extracted.Length > 1_500_000 Then
                extracted = extracted.Substring(0, 1_500_000) & System.Environment.NewLine & "[…truncated…]"
            End If

            Return True
        Catch ex As System.Exception
            System.Diagnostics.Debug.WriteLine("Text-like extract failed: " & ex.Message)
            extracted = Nothing
            label = Nothing
            Return False
        End Try
    End Function

    Private Function ReadAllTextSmart(ByVal path As System.String) As System.String
        ' UTF-8 (mit BOM-Erkennung), Fallback: Windows-1252 → UTF-8
        Try
            Using sr As New System.IO.StreamReader(path, System.Text.Encoding.UTF8, detectEncodingFromByteOrderMarks:=True)
                Dim s As System.String = sr.ReadToEnd()
                If Not System.String.IsNullOrEmpty(s) Then Return s
            End Using
        Catch
        End Try
        Try
            Dim enc As System.Text.Encoding = System.Text.Encoding.GetEncoding(1252) ' Westeuropa Win-1252
            Return System.IO.File.ReadAllText(path, enc)
        Catch
            ' letzter Fallback
            Try
                Return System.IO.File.ReadAllText(path)
            Catch
                Return Nothing
            End Try
        End Try
    End Function


End Class


