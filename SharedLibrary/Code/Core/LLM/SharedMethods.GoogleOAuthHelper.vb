' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: SharedMethods.GoogleOAuthHelper.vb
' Purpose: Builds and signs a Google-style OAuth 2.0 JWT assertion (RS256) and exchanges it
'          for an access token by POSTing the assertion to the configured token endpoint.
'
' Architecture:
'  - Configuration Inputs (shared fields): `client_email`, `private_key`, `scopes`, `token_uri`,
'    `token_life` (currently not used by this implementation).
'  - JWT Construction: Creates a compact JWT with header `{alg=RS256, typ=JWT}` and payload containing
'    `iss`, `scope`, `aud`, `exp`, `iat`, then Base64Url-encodes header/payload.
'  - Signing: Uses BouncyCastle to parse a PEM RSA private key and signs `<header>.<payload>` via
'    SHA256withRSA (RS256), then Base64Url-encodes the signature.
'  - Token Exchange: Sends JSON `{ grant_type, assertion }` to `token_uri` and returns `access_token`
'    from the JSON response.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Text
Imports System.IO
Imports System.Net.Http
Imports Newtonsoft.Json
Imports Org.BouncyCastle.Crypto.Parameters
Imports Org.BouncyCastle.Security

Namespace SharedLibrary
    Partial Public Class SharedMethods

        ''' <summary>
        ''' Helper for generating an RS256-signed JWT assertion and exchanging it for an OAuth access token.
        ''' </summary>
        Public Class GoogleOAuthHelper
            ' Public variables

            ''' <summary>
            ''' Service account email used as the JWT issuer (`iss`).
            ''' </summary>
            Public Shared client_email As String = ""

            ''' <summary>
            ''' PEM-encoded RSA private key used to sign the JWT (expected to be readable by BouncyCastle's PEM reader).
            ''' </summary>
            Public Shared private_key As String = ""

            ''' <summary>
            ''' OAuth scope string placed into the JWT payload (`scope`).
            ''' </summary>
            Public Shared scopes As String = ""

            ''' <summary>
            ''' OAuth token endpoint URI used as audience (`aud`) and POST destination for token exchange.
            ''' </summary>
            Public Shared token_uri As String = ""

            ''' <summary>
            ''' Token lifetime in seconds.
            ''' </summary>
            ''' <remarks>
            ''' This value is currently not used by `GenerateJWT`, which uses a fixed 3600 second expiry.
            ''' </remarks>
            Public Shared token_life As Long = 0

            ' Base64Url encoding

            ''' <summary>
            ''' Base64Url-encodes a UTF-8 string (no padding) per JWT requirements.
            ''' </summary>
            ''' <param name="input">Text to encode using UTF-8.</param>
            ''' <returns>Base64Url-encoded string without padding characters.</returns>
            Private Shared Function Base64UrlEncode(input As String) As String
                Return System.Convert.ToBase64String(Encoding.UTF8.GetBytes(input)).
                Replace("+", "-").
                Replace("/", "_").
                Replace("=", "")
            End Function

            ''' <summary>
            ''' Base64Url-encodes a byte array (no padding) per JWT requirements.
            ''' </summary>
            ''' <param name="inputBytes">Bytes to encode.</param>
            ''' <returns>Base64Url-encoded string without padding characters.</returns>
            Private Shared Function Base64UrlEncode(inputBytes As Byte()) As String
                Return System.Convert.ToBase64String(inputBytes).
                Replace("+", "-").
                Replace("/", "_").
                Replace("=", "")
            End Function

            ' Sign data using BouncyCastle

            ''' <summary>
            ''' Signs the provided data with the configured RSA private key using SHA256withRSA (RS256).
            ''' </summary>
            ''' <param name="data">Data to sign.</param>
            ''' <returns>Signature bytes.</returns>
            Private Shared Function SignData(data As Byte()) As Byte()
                Dim rsaKey As RsaPrivateCrtKeyParameters
                Dim formattedPrivateKey As String = private_key.Replace("\n", Environment.NewLine)

                Using reader As New StringReader(formattedPrivateKey)
                    Dim pemReader = New Org.BouncyCastle.OpenSsl.PemReader(reader)
                    rsaKey = DirectCast(pemReader.ReadObject(), RsaPrivateCrtKeyParameters)
                End Using

                Dim signer = SignerUtilities.GetSigner("SHA256withRSA")
                signer.Init(True, rsaKey)
                signer.BlockUpdate(data, 0, data.Length)
                Return signer.GenerateSignature()
            End Function

            ' Generate JWT

            ''' <summary>
            ''' Generates a compact serialized JWT signed with RS256 containing `iss`, `scope`, `aud`, `exp`, and `iat`.
            ''' </summary>
            ''' <returns>Compact JWT string (`Base64Url(header).Base64Url(payload).Base64Url(signature)`).</returns>
            Public Shared Function GenerateJWT() As String
                Dim issuedAt As Long = DateTimeOffset.UtcNow.ToUnixTimeSeconds()
                Dim expiry As Long = issuedAt + 3600 ' 1 hour expiry

                Dim header = New With {.alg = "RS256", .typ = "JWT"}
                Dim payload = New With {
                                        .iss = client_email,
                                        .scope = scopes,
                                        .aud = token_uri,
                                        .exp = expiry,
                                        .iat = issuedAt
                                    }

                Dim headerBase64 = Base64UrlEncode(JsonConvert.SerializeObject(header))
                Dim payloadBase64 = Base64UrlEncode(JsonConvert.SerializeObject(payload))
                Dim unsignedToken = $"{headerBase64}.{payloadBase64}"
                Dim signature = SignData(Encoding.UTF8.GetBytes(unsignedToken))
                Dim signatureBase64 = Base64UrlEncode(signature)

                Return $"{unsignedToken}.{signatureBase64}"
            End Function

            ' Get Access Token

            ''' <summary>
            ''' Requests an OAuth access token by exchanging a signed JWT assertion at the configured token endpoint.
            ''' </summary>
            ''' <returns>Access token string on success; otherwise an empty string.</returns>
            Public Shared Async Function GetAccessToken() As Task(Of String)
                Dim jwt = GenerateJWT()
                Dim requestBody As String = JsonConvert.SerializeObject(New With {
                                .grant_type = "urn:ietf:params:oauth:grant-type:jwt-bearer",
                                .assertion = jwt
                                            })

                Using client As New HttpClient()
                    Dim content = New StringContent(requestBody, Encoding.UTF8, "application/json")
                    Dim response = Await client.PostAsync(token_uri, content)

                    If response.IsSuccessStatusCode Then
                        Dim responseBody = Await response.Content.ReadAsStringAsync()
                        Dim tokenData = JsonConvert.DeserializeObject(Of Dictionary(Of String, String))(responseBody)
                        Return tokenData("access_token")
                    Else
                        ShowCustomMessageBox($"Error getting access token: {response.ReasonPhrase}")
                        Return ""
                    End If
                End Using
            End Function
        End Class

    End Class

End Namespace