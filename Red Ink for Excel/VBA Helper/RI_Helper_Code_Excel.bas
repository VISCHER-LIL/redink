Attribute VB_Name = "RI_Helper_Code"

' Helper Code for "Red Ink for Excel"
'
' 11.1.2025
'
' These procedures are used if you configure "Red Ink for Excel" to assign a key to a particular
' functionality. If that happens, the key is assigned to these macros, which then call the relevant
' procedure within the VSTO Add-in of Red Ink for Excel. If you are not allowed to run them, the
' add-in will still work, but you can't use the key shortcuts defined.
'
' All Rights Reserved. david.rosenthal@vischer.com  https://www.vischer.com/redink

Option Explicit

Const CurrentVersion As Integer = 1
Const AddinName As String = "Red Ink for Excel"
Const AN As String = "Red Ink"

Private INI_APIKey As String
Public INI_Temperature As String
Public INI_Timeout As Long
Public INI_Model As String
Private INI_Endpoint As String
Private INI_HeaderA As String
Private INI_HeaderB As String
Private INI_APICall As String
Private INI_Response As String
Private INI_DoubleS As Boolean
Public INI_PreCorrection As String
Public INI_PostCorrection As String
Public INI_MaxOutputToken As Integer

Private INI_OAuth2 As Boolean
Private INI_OAuth2ClientMail As String
Private INI_OAuth2Scopes As String
Private INI_OAuth2Endpoint As String
Private INI_OAuth2ATExpiry As Long
Private INI_OpenSSLPath As String

Private INI_APIDebug As Boolean

Private DecodedAPI As String

Private gAccessToken_1 As String
Private gAccessTokenExpires_1 As Date
Private gAccessToken_2 As String
Private gAccessTokenExpires_2 As Date

Public ModuleRunning As Integer


#If VBA7 Then
    ' For 64-bit systems
    Private Declare PtrSafe Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
#Else
    ' For 32-bit systems
    Private Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
#End If

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Type TIME_ZONE_INFORMATION
    Bias As Long
    StandardName(0 To 31) As Byte
    StandardDate As SYSTEMTIME
    StandardBias As Long
    DaylightName(0 To 31) As Byte
    DaylightDate As SYSTEMTIME
    DaylightBias As Long
End Type


' When starting

Sub Auto_Open()
    ModuleRunning = CurrentVersion
    CallAddContextMenu
End Sub

' Loopback Function

Public Function CheckAppHelper() As Integer
    CheckAppHelper = ModuleRunning
End Function

' Testing the LLM

Public Sub TestLLM()
    MsgBox ("The LLM was asked to write a poem for the current month. This is what it produced:" & vbCrLf & vbCrLf & LLM("Write me a poem for the month of " & MonthName(Month(Date)), ""))
End Sub

' Helpers to Call Functions within the Add-in

Sub CallAddContextMenu()
    Dim AddIn As Object
    On Error Resume Next
    Set AddIn = Application.COMAddIns(AddinName).Object
    On Error GoTo 0
    If Not AddIn Is Nothing Then
        AddIn.DoAddContextMenu
    End If
End Sub

Sub CallInLanguage1()
    Dim AddIn As Object
    On Error Resume Next
    Set AddIn = Application.COMAddIns(AddinName).Object
    On Error GoTo 0
    If Not AddIn Is Nothing Then
        AddIn.DoInLanguage1
    End If
End Sub

Sub CallInLanguage2()
    Dim AddIn As Object
    On Error Resume Next
    Set AddIn = Application.COMAddIns(AddinName).Object
    On Error GoTo 0
    If Not AddIn Is Nothing Then
        AddIn.DoInLanguage2
    End If
End Sub

Sub CallInOther()
    Dim AddIn As Object
    On Error Resume Next
    Set AddIn = Application.COMAddIns(AddinName).Object
    On Error GoTo 0
    If Not AddIn Is Nothing Then
        AddIn.DoInOther
    End If
End Sub

Sub CallInOtherFormulas()
    Dim AddIn As Object
    On Error Resume Next
    Set AddIn = Application.COMAddIns(AddinName).Object
    On Error GoTo 0
    If Not AddIn Is Nothing Then
        AddIn.DoInOtherFormulas
    End If
End Sub

Sub CallCorrect()
    Dim AddIn As Object
    On Error Resume Next
    Set AddIn = Application.COMAddIns(AddinName).Object
    On Error GoTo 0
    If Not AddIn Is Nothing Then
        AddIn.DoCorrect
    End If
End Sub

Sub CallNeatly()
    Dim AddIn As Object
    On Error Resume Next
    Set AddIn = Application.COMAddIns(AddinName).Object
    On Error GoTo 0
    If Not AddIn Is Nothing Then
        AddIn.DoImprove
    End If
End Sub

Sub CallShorten()
    Dim AddIn As Object
    On Error Resume Next
    Set AddIn = Application.COMAddIns(AddinName).Object
    On Error GoTo 0
    If Not AddIn Is Nothing Then
        AddIn.DoShorten
    End If
End Sub

Sub CallAnonymize()
    Dim AddIn As Object
    On Error Resume Next
    Set AddIn = Application.COMAddIns(AddinName).Object
    On Error GoTo 0
    If Not AddIn Is Nothing Then
        AddIn.DoAnonymize
    End If
End Sub

Sub CallSwitchParty()
    Dim AddIn As Object
    On Error Resume Next
    Set AddIn = Application.COMAddIns(AddinName).Object
    On Error GoTo 0
    If Not AddIn Is Nothing Then
        AddIn.DoSwitchParty
    End If
End Sub

Sub CallFreestyleNM()
    Dim AddIn As Object
    On Error Resume Next
    Set AddIn = Application.COMAddIns(AddinName).Object
    On Error GoTo 0
    If Not AddIn Is Nothing Then
        AddIn.DoFreestyleNM
    End If
End Sub

Sub CallFreestyleNMF()
    Dim AddIn As Object
    On Error Resume Next
    Set AddIn = Application.COMAddIns(AddinName).Object
    On Error GoTo 0
    If Not AddIn Is Nothing Then
        AddIn.DoFreestyleNMF
    End If
End Sub

Sub CallFreestyleAM()
    Dim AddIn As Object
    On Error Resume Next
    Set AddIn = Application.COMAddIns(AddinName).Object
    On Error GoTo 0
    If Not AddIn Is Nothing Then
        AddIn.DoFreestyleAM
    End If
End Sub

Sub CallFreestyleAMF()
    Dim AddIn As Object
    On Error Resume Next
    Set AddIn = Application.COMAddIns(AddinName).Object
    On Error GoTo 0
    If Not AddIn Is Nothing Then
        AddIn.DoFreestyleAMF
    End If
End Sub

Sub CallAdjustHeight()
    Dim AddIn As Object
    On Error Resume Next
    Set AddIn = Application.COMAddIns(AddinName).Object
    On Error GoTo 0
    If Not AddIn Is Nothing Then
        AddIn.DoAdjustHeight
    End If
End Sub

Public Sub AdjustHeight()
    Dim AddIn As Object
    On Error Resume Next
    Set AddIn = Application.COMAddIns(AddinName).Object
    On Error GoTo 0
    If Not AddIn Is Nothing Then
        AddIn.DoAdjustHeight(True)
    End If
End Sub

Sub CallAdjustLegacyNotes()
    Dim AddIn As Object
    On Error Resume Next
    Set AddIn = Application.COMAddIns(AddinName).Object
    On Error GoTo 0
    If Not AddIn Is Nothing Then
        AddIn.DoAdjustLegacyNotes
    End If
End Sub

Sub CallRegexSearchReplace()
    Dim AddIn As Object
    On Error Resume Next
    Set AddIn = Application.COMAddIns(AddinName).Object
    On Error GoTo 0
    If Not AddIn Is Nothing Then
        AddIn.DoRegexSearchReplace
    End If
End Sub

' Helper to Read Files, including PDF

Public Function GetFileTextContent(filename As String, Optional ReturnErrorInsteadOfEmpty As Boolean = True) As String
    
    Dim AddIn As Object
    Dim Result As String

    On Error Resume Next

    ' Access the COM Add-In
    Set AddIn = Application.COMAddIns(AddinName).Object

    ' Call the COM Add-In Function
    GetFileTextContent = AddIn.GetFileTextContent(filename, ReturnErrorInsteadOfEmpty)

End Function

' Helper to access the LLM

Private Function LoadAPIConfigValues(UseSecondAPI As Boolean) As Boolean

    Dim DictType As String

    Dim AddIn As Object
    Dim APIConfig As String

    On Error Resume Next

    ' Access the COM Add-In
    Set AddIn = Application.COMAddIns(AddinName).Object

    If AddIn Is Nothing Then
        MsgBox "The " & AN & " Excel Add-In object is not accessible for the " & AN & " Helper. Make sure it is installed."
        Exit Function
    End If

    On Error Resume Next
    
    APIConfig = AddIn.GetLLMConfig(UseSecondAPI)

    'Debug.Print APIConfig

    ' Ensure APIConfig is retrieved correctly
    If APIConfig <> "" Then

        DecodeAPIConfigurationString (APIConfig)

        INI_APIDebug = True

        LoadAPIConfigValues = True

    Else
        LoadAPIConfigValues = False
    End If

    ' Clean up
    Set AddIn = Nothing

End Function

Sub DecodeAPIConfigurationString(encodedString As String)
    Dim pairs() As String
    Dim pair As Variant
    Dim keyValue() As String

    ' Split the encoded string into key-value pairs
    pairs = Split(encodedString, Chr(64) & Chr(64) & Chr(64))

    ' Loop through each key-value pair
    For Each pair In pairs
        keyValue = Split(pair, Chr(167) & Chr(167))
        If UBound(keyValue) = 1 Then
            Select Case keyValue(0)
                Case "INI_OAuth2": INI_OAuth2 = CBool(keyValue(1))
                Case "INI_OAuth2ClientMail": INI_OAuth2ClientMail = keyValue(1)
                Case "INI_OAuth2Scopes": INI_OAuth2Scopes = keyValue(1)
                Case "INI_OAuth2Endpoint": INI_OAuth2Endpoint = keyValue(1)
                Case "INI_OAuth2ATExpiry": INI_OAuth2ATExpiry = CLng(keyValue(1))
                Case "INI_APIKey": INI_APIKey = keyValue(1)
                Case "INI_Temperature": INI_Temperature = CDbl(keyValue(1))
                Case "INI_Timeout": INI_Timeout = CLng(keyValue(1))
                Case "INI_MaxOutputToken": INI_MaxOutputToken = CLng(keyValue(1))
                Case "INI_Model": INI_Model = keyValue(1)
                Case "INI_Endpoint": INI_Endpoint = keyValue(1)
                Case "INI_HeaderA": INI_HeaderA = keyValue(1)
                Case "INI_HeaderB": INI_HeaderB = keyValue(1)
                Case "INI_APICall": INI_APICall = keyValue(1)
                Case "INI_Response": INI_Response = keyValue(1)
                Case "DecodedAPI": DecodedAPI = keyValue(1)
            End Select
            'Debug.Print keyValue(0) & "=" & keyValue(1)
        End If
    Next pair

End Sub


Public Function LLM(ByVal promptSystem As String, ByVal promptUser As String, Optional ByVal Model As String = "", Optional ByVal Temperature As String = "", Optional ByVal TimeOut As Long = 0, Optional ByVal UseSecondAPI As Boolean = False, Optional ByVal HideSplash As Boolean = False) As String

    On Error GoTo ErrorHandler

    LLM = ""
    
    If Not LoadAPIConfigValues(UseSecondAPI) Then
        MsgBox AN & " could not load the API configuration from Excel COM Add-in. Make sure it is running and properly configured.", vbCritical
        Exit Function
    End If

    If INI_OAuth2 Then
        DecodedAPI = GetFreshAccessToken(INI_OAuth2ClientMail, INI_OAuth2Scopes, INI_APIKey, INI_OAuth2Endpoint, INI_OAuth2ATExpiry, UseSecondAPI)
        If DecodedAPI = "" Then Exit Function
    End If

    If (Len(promptUser) > 3000 Or Len(promptSystem) > 800) And Not Hidesplash Then
        RI_Wait.PictureSizeMode = 1 'fmPictureSizeModeStretch
        RI_Wait.Show vbModeless
        DoEvents
    End If

    promptSystem = CleanString(promptSystem)
    promptUser = CleanString(promptUser)

    ' Define local variables to hold the API configuration
    Dim Endpoint As String
    Dim HeaderA As String
    Dim HeaderB As String
    Dim APICall As String
    Dim ResponseKey As String
    Dim TemperatureValue As String
    Dim ModelValue As String
    Dim TimeoutValue As Long
    Dim DoubleS As Boolean

    ' Use first set of variables
    Endpoint = Replace(Replace(INI_Endpoint, "{model}", INI_Model), "{apikey}", DecodedAPI)
    HeaderA = Replace(Replace(INI_HeaderA, "{model}", INI_Model), "{apikey}", DecodedAPI)
    HeaderB = Replace(Replace(INI_HeaderB, "{model}", INI_Model), "{apikey}", DecodedAPI)
    APICall = INI_APICall
    ResponseKey = INI_Response
    DoubleS = INI_DoubleS
    ' Handle default values
    If Temperature = "" Or Temperature = "Default" Then
        TemperatureValue = INI_Temperature
    Else
        TemperatureValue = Temperature
    End If
    If Model = "" Or Model = "Default" Then
        ModelValue = INI_Model
    Else
        ModelValue = Model
    End If
    If TimeOut = 0 Then
        TimeoutValue = INI_Timeout
    Else
        TimeoutValue = TimeOut
    End If

    ' Create and configure the HTTP request
    Dim Headers As Object
    Set Headers = CreateObject("MSXML2.ServerXMLHTTP")
    Headers.Open "POST", Endpoint, False
    Headers.setRequestHeader "Content-Type", "application/json"

    If HeaderA & HeaderB <> "" Then
        Headers.setRequestHeader HeaderA, HeaderB
    End If

    ' Prepare the request body
    Dim Body As String
    Body = Replace(APICall, "{model}", ModelValue)
    Body = Replace(Body, "{promptsystem}", promptSystem)
    Body = Replace(Body, "{promptuser}", promptUser)
    Body = Replace(Body, "{temperature}", CStr(TemperatureValue))

    If INI_APIDebug Then
        Debug.Print "SENT TO API:" & vbNewLine & Body
    End If

    ' Set timeouts and send the request
    Headers.SetTimeouts 20000, 20000, TimeoutValue, TimeoutValue
    Headers.send (Body)

    ' Receive and process the response
    Dim Response As String
    Response = Headers.responseText

    Unload RI_Wait

    If INI_APIDebug Then
        Debug.Print "RECEIVED FROM API:" & vbNewLine & Response
    End If

    Dim text As String
    text = ExtractJSONValue(Response, "error")
    If Len(text) > 0 Then
        text = ExtractJSONValue(Response, "message")
        MsgBox "The LLM API generated the following error message: " & vbNewLine & vbNewLine & text & vbNewLine & vbNewLine & Response & vbNewLine, vbCritical
        LLM = ""
    Else
        text = ExtractJSONValue(Response, ResponseKey)
        text = ConvertEscapeCharacters(text)
        If DoubleS Then
            text = Replace(text, ChrW(223), "ss")
        End If
        LLM = text
    End If

    INI_APIKey = ""
    DecodedAPI = ""

    Exit Function

ErrorHandler:
    Unload RI_Wait
    MsgBox "Internal error message in the " & AN & " Helper LLM procedure: " & Err.Description, vbCritical

End Function

Public Function GetFreshAccessToken(ByVal clientEmail As String, ByVal ClientScopes As String, _
                                    ByVal PrivateKey As String, ByVal AuthServer As String, _
                                    ByVal TLife As Long, ByVal SecondAPI As Boolean) As String
    On Error GoTo ErrorHandler
    
    Dim currentToken As String
    Dim currentExpiry As Date
    
    ' Decide which token/account we are dealing with
    If SecondAPI = True Then
        currentToken = gAccessToken_2
        currentExpiry = gAccessTokenExpires_2
    Else
        currentToken = gAccessToken_1
        currentExpiry = gAccessTokenExpires_1
    End If

    ' Check if we already have a valid token that is not about to expire
    ' Consider a 60-second buffer
    If Len(currentToken) > 0 And (DateDiff("s", Now, currentExpiry) > 60) Then
        ' Token still valid
        GetFreshAccessToken = currentToken
        Exit Function
    End If

    ' If we reach here, we need to obtain a new token
    Dim newToken As String
    Dim newExpiry As Date

    newToken = ObtainAccessTokenFromJWT(clientEmail, ClientScopes, PrivateKey, AuthServer, TLife, newExpiry)
    If Len(newToken) = 0 Then
        ' Error obtaining token
        ' ObtainAccessTokenFromJWT should have shown a message box, just exit
        Exit Function
    End If

    ' Update global variables
    If SecondAPI = True Then
        gAccessToken_2 = newToken
        gAccessTokenExpires_2 = newExpiry
    Else
        gAccessToken_1 = newToken
        gAccessTokenExpires_1 = newExpiry
    End If

    GetFreshAccessToken = newToken
    Exit Function

ErrorHandler:
    MsgBox "Error in " & AN & " Helper / GetFreshAccessToken: " & Err.Number & " - " & Err.Description, vbCritical
End Function

' This function handles creating the JWT, signing it, and calling the Auth endpoint

Private Function ObtainAccessTokenFromJWT(ByVal clientEmail As String, ByVal ClientScopes As String, _
                                          ByVal PrivateKey As String, ByVal AuthServer As String, _
                                          ByVal TLife As Long, ByRef TokenExpiry As Date) As String
    On Error GoTo ErrHandler
    
    Dim iat As Long
    Dim exp As Long
    Dim JWT_Header As String
    Dim JWT_ClaimSet As String
    Dim JWT_Unsigned As String
    Dim JWT_Signed As String
    
    ' Current Unix time
    iat = GetUnixTimestamp()
    exp = iat + TLife - 30

    ' Construct the header and claim set
    JWT_Header = "{""alg"":""RS256"",""typ"":""JWT""}"
    JWT_ClaimSet = "{" & _
        """iss"":""" & clientEmail & """," & _
        """scope"":""" & ClientScopes & """," & _
        """aud"":""" & AuthServer & """," & _
        """exp"":" & exp & "," & _
        """iat"":" & iat & "}"
    
    ' Encode header and claim set
    Dim encodedHeader As String
    Dim encodedClaimSet As String

    encodedHeader = EncodeBase64(JWT_Header)
    encodedClaimSet = EncodeBase64(JWT_ClaimSet)

    JWT_Unsigned = encodedHeader & "." & encodedClaimSet

    ' Call VB.NET function to sign the JWT
    Dim vbnetHelper As Object
    'Set vbnetHelper = CreateObject("YourAddinNamespace.AddInFunctions")
    Set vbnetHelper = Application.COMAddIns(AddinName).Object

    Dim base64Signature As String
    base64Signature = vbnetHelper.SignJWT(JWT_Unsigned, PrivateKey)
    
    JWT_Signed = JWT_Unsigned & "." & base64Signature

    ' Now we have the signed JWT. Let's request the Access Token
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")

    Dim postData As String
    postData = "grant_type=urn%3Aietf%3Aparams%3Aoauth%3Agrant-type%3Ajwt-bearer&assertion=" & Replace(JWT_Signed, "+", "%2B")

    http.Open "POST", AuthServer, False
    http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    http.send postData

    If http.Status <> 200 Then
        MsgBox "Error obtaining token: HTTP " & http.Status & " - " & http.responseText, vbCritical
        GoTo Cleanup
    End If

    Dim responseJSON As String
    responseJSON = http.responseText

    ' Parse the JSON response for "access_token" and "expires_in"
    Dim accessToken As String
    Dim expiresIn As Long

    accessToken = ExtractJSONValue(responseJSON, "access_token")
    
    expiresIn = CLng(ExtractJSONValue(responseJSON, "expires_in"))

    If Len(accessToken) = 0 Then
        MsgBox "Red " & AN & " Helper: No access token found in the response received from the authentication server: " & vbCrLf & vbCrLf & responseJSON, vbCritical
        GoTo Cleanup
    End If

    ' Calculate the expiry time in local time
    TokenExpiry = DateAdd("s", expiresIn, Now())
    ObtainAccessTokenFromJWT = accessToken

Cleanup:
    On Error GoTo 0

    Exit Function

ErrHandler:
    MsgBox AN & " Helper: Error in obtaining access token: " & Err.Number & " - " & Err.Description, vbCritical
    Resume Cleanup
End Function

Private Function ExtractJSONValue(jsonString As String, objectName As String) As String

On Error GoTo ErrorHandler

    Dim searchKey As String
    Dim keyPos As Long
    Dim pos As Long
    Dim valueStartPos As Long
    Dim valueEndPos As Long
    Dim valueStartChar As String
    Dim c As String
    Dim braceCount As Long
    Dim bracketCount As Long
    Dim backslashCount As Long
    Dim tempPos As Long
    Dim valueString As String

    searchKey = """" & objectName & """"  ' Enclose objectName in double quotes

    ' Find the position of the key in the JSON string
    keyPos = InStr(1, jsonString, searchKey)
    If keyPos = 0 Then
        ' Key not found
        ExtractJSONValue = ""
        Exit Function
    End If

    ' Move past the key
    pos = keyPos + Len(searchKey)

    ' Skip any whitespace
    Do While pos <= Len(jsonString) And _
        (Mid(jsonString, pos, 1) = " " Or Mid(jsonString, pos, 1) = vbTab Or _
         Mid(jsonString, pos, 1) = vbCr Or Mid(jsonString, pos, 1) = vbLf)
        pos = pos + 1
    Loop

    ' Check for colon ':'
    If Mid(jsonString, pos, 1) <> ":" Then
        ' Invalid JSON format
        ExtractJSONValue = ""
        Exit Function
    End If

    pos = pos + 1 ' Move past the ':'

    ' Skip any whitespace after the colon
    Do While pos <= Len(jsonString) And _
        (Mid(jsonString, pos, 1) = " " Or Mid(jsonString, pos, 1) = vbTab Or _
         Mid(jsonString, pos, 1) = vbCr Or Mid(jsonString, pos, 1) = vbLf)
        pos = pos + 1
    Loop

    ' Get the first character of the value
    valueStartChar = Mid(jsonString, pos, 1)

    Select Case valueStartChar
        Case """"
            ' String value
            valueStartPos = pos + 1  ' Skip the opening quote
            valueEndPos = valueStartPos
            Do While valueEndPos <= Len(jsonString)
                If Mid(jsonString, valueEndPos, 1) = """" Then
                    ' Check if the quote is escaped
                    backslashCount = 0
                    tempPos = valueEndPos - 1
                    Do While tempPos >= valueStartPos And Mid(jsonString, tempPos, 1) = "\"
                        backslashCount = backslashCount + 1
                        tempPos = tempPos - 1
                    Loop
                    If backslashCount Mod 2 = 0 Then
                        ' Even number of backslashes, so the quote is not escaped
                        Exit Do
                    End If
                End If
                valueEndPos = valueEndPos + 1
            Loop
            ' Extract the string value
            valueString = Mid(jsonString, valueStartPos, valueEndPos - valueStartPos)
            ' Replace escaped characters
            valueString = Replace(valueString, "\""", """")
            valueString = Replace(valueString, "\\", "\")
            ExtractJSONValue = valueString
            Exit Function

        Case "{"
            ' Object value
            valueStartPos = pos
            braceCount = 1
            valueEndPos = pos + 1
            Do While valueEndPos <= Len(jsonString) And braceCount > 0
                c = Mid(jsonString, valueEndPos, 1)
                If c = "{" Then
                    braceCount = braceCount + 1
                ElseIf c = "}" Then
                    braceCount = braceCount - 1
                ElseIf c = """" Then
                    ' Skip strings inside the object
                    valueEndPos = valueEndPos + 1
                    Do While valueEndPos <= Len(jsonString)
                        If Mid(jsonString, valueEndPos, 1) = """" Then
                            ' Check if the quote is escaped
                            backslashCount = 0
                            tempPos = valueEndPos - 1
                            Do While tempPos >= valueStartPos And Mid(jsonString, tempPos, 1) = "\"
                                backslashCount = backslashCount + 1
                                tempPos = tempPos - 1
                            Loop
                            If backslashCount Mod 2 = 0 Then
                                Exit Do
                            End If
                        End If
                        valueEndPos = valueEndPos + 1
                    Loop
                End If
                valueEndPos = valueEndPos + 1
            Loop
            valueString = Mid(jsonString, valueStartPos, valueEndPos - valueStartPos)
            ExtractJSONValue = valueString
            Exit Function

        Case "["
            ' Array value
            valueStartPos = pos
            bracketCount = 1
            valueEndPos = pos + 1
            Do While valueEndPos <= Len(jsonString) And bracketCount > 0
                c = Mid(jsonString, valueEndPos, 1)
                If c = "[" Then
                    bracketCount = bracketCount + 1
                ElseIf c = "]" Then
                    bracketCount = bracketCount - 1
                ElseIf c = """" Then
                    ' Skip strings inside the array
                    valueEndPos = valueEndPos + 1
                    Do While valueEndPos <= Len(jsonString)
                        If Mid(jsonString, valueEndPos, 1) = """" Then
                            backslashCount = 0
                            tempPos = valueEndPos - 1
                            Do While tempPos >= valueStartPos And Mid(jsonString, tempPos, 1) = "\"
                                backslashCount = backslashCount + 1
                                tempPos = tempPos - 1
                            Loop
                            If backslashCount Mod 2 = 0 Then
                                Exit Do
                            End If
                        End If
                        valueEndPos = valueEndPos + 1
                    Loop
                End If
                valueEndPos = valueEndPos + 1
            Loop
            valueString = Mid(jsonString, valueStartPos, valueEndPos - valueStartPos)
            ExtractJSONValue = valueString
            Exit Function

        Case "t"
            ' Check for "true"
            If Mid(jsonString, pos, 4) = "true" Then
                ExtractJSONValue = "true"
                Exit Function
            End If

        Case "f"
            ' Check for "false"
            If Mid(jsonString, pos, 5) = "false" Then
                ExtractJSONValue = "false"
                Exit Function
            End If

        Case "n"
            ' Check for "null"
            If Mid(jsonString, pos, 4) = "null" Then
                ExtractJSONValue = "null"
                Exit Function
            End If

        Case Else
            ' Number value
            valueStartPos = pos
            valueEndPos = pos
            Do While valueEndPos <= Len(jsonString)
                c = Mid(jsonString, valueEndPos, 1)
                If c = "," Or c = "}" Or c = "]" Or c = vbCr Or c = vbLf Or _
                   c = " " Or c = vbTab Then
                    Exit Do
                End If
                valueEndPos = valueEndPos + 1
            Loop
            valueString = Mid(jsonString, valueStartPos, valueEndPos - valueStartPos)
            ExtractJSONValue = valueString
            Exit Function
    End Select

    ' If none of the cases match
    ExtractJSONValue = ""
    
    Exit Function
    
ErrorHandler:
    MsgBox "Error in " & AN & " Helper - ExtractJSONValue: " & Err.Description, vbCritical
End Function

Private Function ConvertEscapeCharacters(ByVal inputtext As String) As String
    'inputText = Replace(inputText, "\n\n", vbCrLf & vbCrLf) ' Doppelte neue Zeilen
    inputtext = Replace(inputtext, "\n\n", "\n")            ' Doppelte Zeilen eliminieren
    inputtext = Replace(inputtext, "\n", vbLf)              ' Neue Zeile
    inputtext = Replace(inputtext, "\r", vbCr)              ' Wagenr�cklauf
    inputtext = Replace(inputtext, "\t", vbTab)             ' Tabulator
    inputtext = Replace(inputtext, "\\", "\")
    inputtext = Replace(inputtext, "\""", """")

    inputtext = ConvertUnicodeEscapes(inputtext)

    ConvertEscapeCharacters = inputtext
End Function

Function ConvertUnicodeEscapes(ByVal inputtext As String) As String
    Dim regex As Object
    Dim matches As Object
    Dim m As Object
    Dim unicodePattern As String
    Dim unicodeValue As Long
    Dim i As Long
    
    Set regex = CreateObject("VBScript.RegExp")
    
    unicodePattern = "\\u([0-9A-Fa-f]{4})"
    
    With regex
        .pattern = unicodePattern
        .Global = True
        .IgnoreCase = True
    End With
    
    Set matches = regex.Execute(inputtext)
    
    ' Process matches in reverse order
    For i = matches.Count - 1 To 0 Step -1
        Set m = matches(i)
        unicodeValue = CLng("&H" & m.SubMatches(0))
        inputtext = Left(inputtext, m.FirstIndex) & ChrW(unicodeValue) & Mid(inputtext, m.FirstIndex + m.Length + 1)
    Next i

    ConvertUnicodeEscapes = inputtext
End Function

Private Function CleanString(ByVal XInputS As String) As String

    Dim inputs As String
    Dim cleanedString As String
    inputs = XInputS
    cleanedString = ""
    
    Dim i As Long
    Dim currentChar As String
    Dim charCode As Long
    
    For i = 1 To Len(inputs)
        currentChar = Mid(inputs, i, 1)
        charCode = AscW(currentChar)
        
        Select Case charCode
            Case 8
                cleanedString = cleanedString & "\b"
            Case 9
                cleanedString = cleanedString & "\t"
            Case 10
                cleanedString = cleanedString & "\n"
            Case 12
                cleanedString = cleanedString & "\f"
            Case 13
                cleanedString = cleanedString & "\r"
            Case 34
                cleanedString = cleanedString & "\"""
            Case 92
                cleanedString = cleanedString & "\\"
            Case 0 To 31
                ' Control characters are not allowed in JSON strings
                ' Represent them as Unicode escape sequences
                cleanedString = cleanedString & "\u" & Right("000" & Hex(charCode), 4)
            Case Else
                ' Include all other characters, including Unicode characters like Umlaute
                cleanedString = cleanedString & currentChar
        End Select
    Next i
    
    ' Condense multiple spaces to a single space
    Do While InStr(cleanedString, "  ") > 0
        cleanedString = Replace(cleanedString, "  ", " ")
    Loop
    
    CleanString = cleanedString

End Function

Private Function GetUnixTimestamp() As Long
    Dim timezoneOffset As Long
    timezoneOffset = GetUTCOffsetSeconds()
    
    ' Convert local time to UTC
    GetUnixTimestamp = DateDiff("s", "01/01/1970 00:00:00", Now() + timezoneOffset / 86400)
End Function

Private Function GetUTCOffsetSeconds() As Long
    Dim timeZoneInfo As TIME_ZONE_INFORMATION
    Dim Result As Long
    
    ' Get timezone information
    Result = GetTimeZoneInformation(timeZoneInfo)
    
    ' Calculate timezone offset in seconds
    ' Bias is the number of minutes west of UTC. Convert it to seconds.
    GetUTCOffsetSeconds = timeZoneInfo.Bias * 60
    
    ' If Daylight Saving Time is active, assume a 1-hour adjustment
    If Result = 2 Then ' Daylight Saving Time active
        GetUTCOffsetSeconds = GetUTCOffsetSeconds - 3600
    End If

End Function

Private Function EncodeBase64(ByVal inputtext As String) As String
    Dim arrData() As Byte
    Dim xmlObj As Object
    Dim node As Object
    Dim base64String As String

    ' Convert the text into a byte array
    arrData = StrConv(inputtext, vbFromUnicode)

    ' Use MSXML to perform Base64 encoding
    Set xmlObj = CreateObject("MSXML2.DOMDocument")
    Set node = xmlObj.createElement("b64")
    node.DataType = "bin.base64"
    node.nodeTypedValue = arrData
    base64String = node.text

    ' Remove any line breaks from the encoded string
    base64String = Replace(base64String, vbCrLf, "")
    base64String = Replace(base64String, vbLf, "")
    base64String = Replace(base64String, vbCr, "")
    
    ' Convert to URL-safe Base64
    base64String = Replace(base64String, "+", "-")
    base64String = Replace(base64String, "/", "_")
    base64String = Replace(base64String, "=", "") ' Remove padding

    EncodeBase64 = base64String

    ' Clean up
    Set node = Nothing
    Set xmlObj = Nothing
End Function

Private Function DecodeBase64(ByVal base64String As String) As String
    Dim arrData() As Byte
    Dim xmlObj As Object
    Dim node As Object

    On Error GoTo ErrorHandler

    ' Remove any whitespace or line breaks from the Base64 string
    base64String = Replace(base64String, vbCrLf, "")
    base64String = Replace(base64String, vbLf, "")
    base64String = Replace(base64String, vbCr, "")
    base64String = Replace(base64String, " ", "")

    ' Convert URL-safe Base64 to standard Base64
    base64String = Replace(base64String, "-", "+")
    base64String = Replace(base64String, "_", "/")

    ' Add padding "=" characters if necessary
    Do While (Len(base64String) Mod 4) <> 0
        base64String = base64String & "="
    Loop

    ' Use MSXML to perform Base64 decoding
    Set xmlObj = CreateObject("MSXML2.DOMDocument")
    Set node = xmlObj.createElement("b64")
    node.DataType = "bin.base64"

    ' Assign the cleaned Base64 string
    node.text = base64String

    ' Extract the binary data
    arrData = node.nodeTypedValue

    ' Convert the byte array back to a string
    DecodeBase64 = StrConv(arrData, vbUnicode)

    ' Clean up
    Set node = Nothing
    Set xmlObj = Nothing
    Exit Function

ErrorHandler:
    DecodeBase64 = "Error: Invalid Base64 input"
    Resume Next
End Function

' Helper: Insert line breaks in Base64 string every N characters
Private Function InsertLineBreaksInBase64(ByVal b64 As String, ByVal LineLen As Integer) As String
    Dim i As Long
    Dim Result As String
    
    i = 1
    Do While i <= Len(b64)
        Result = Result & Mid(b64, i, LineLen) & vbLf
        i = i + LineLen
    Loop
    InsertLineBreaksInBase64 = Trim(Result)
End Function

' Helper: Write string to file
Private Sub WriteStringToFile(ByVal filename As String, ByVal content As String)
    Dim fnum As Integer
    fnum = FreeFile
    Open filename For Output As #fnum
    Print #fnum, content;
    Close #fnum
End Sub

' Helper: Run command and wait for completion
Private Function RunCommandAndWait(ByVal cmd As String) As Long
    Dim wsh As Object
    Dim execObj As Object
    
    Set wsh = CreateObject("WScript.Shell")
    Set execObj = wsh.Exec(cmd)

    Do While execObj.Status = 0
        DoEvents
    Loop
    RunCommandAndWait = execObj.ExitCode
End Function

' Helper: read file into byte array
Private Function ReadFileToBytes(ByVal filename As String) As Byte()
    Dim fnum As Integer
    If Dir(filename) = "" Then
        Exit Function
    End If

    fnum = FreeFile
    Open filename For Binary As #fnum
    If LOF(fnum) > 0 Then
        ReDim ReadFileToBytes(LOF(fnum) - 1)
        Get #fnum, , ReadFileToBytes
    End If
    Close #fnum
End Function

' Helper: Encode signature bytes to URL-safe Base64
Private Function EncodeSignatureToBase64(sigBytes() As Byte) As String
    Dim xmlObj As Object
    Dim node As Object
    Dim base64String As String

    Set xmlObj = CreateObject("MSXML2.DOMDocument")
    Set node = xmlObj.createElement("b64")
    node.DataType = "bin.base64"
    node.nodeTypedValue = sigBytes
    base64String = node.text

    ' Remove line breaks
    base64String = Replace(base64String, vbCrLf, "")
    base64String = Replace(base64String, vbLf, "")
    base64String = Replace(base64String, vbCr, "")

    ' Convert to URL-safe Base64
    base64String = Replace(base64String, "+", "-")
    base64String = Replace(base64String, "/", "_")
    base64String = Replace(base64String, "=", "")

    EncodeSignatureToBase64 = base64String

    Set node = Nothing
    Set xmlObj = Nothing
End Function

Function ExpandEnvironmentVariables(ByVal filePath As String) As String
    ' Expand common environment variables like %APPDATA%, %USERPROFILE%, %WINDIR%, etc.
    Dim expandedPath As String
    expandedPath = filePath
    
    ' Expand known variables manually using Environ function
    expandedPath = Replace(expandedPath, "%APPDATA%", Environ("APPDATA"))
    expandedPath = Replace(expandedPath, "%USERPROFILE%", Environ("USERPROFILE"))
    expandedPath = Replace(expandedPath, "%WINDIR%", Environ("WINDIR"))
    expandedPath = Replace(expandedPath, "%TEMP%", Environ("TEMP"))
    expandedPath = Replace(expandedPath, "%HOMEPATH%", Environ("HOMEPATH"))
    expandedPath = Replace(expandedPath, "%APPSTARTUPPATH%", Application.StartupPath)

    ' Return the expanded path
    ExpandEnvironmentVariables = expandedPath
    
End Function

