' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

Option Strict On
Option Explicit On

Imports System.IO
Imports System.Text
Imports Org.BouncyCastle.Crypto
Imports Org.BouncyCastle.Crypto.Parameters
Imports Org.BouncyCastle.Security

Namespace SharedLibrary
    Partial Public Class SharedMethods
        Public Shared Function SignJWT(jwtUnsigned As String, privateKeyPem As String) As String
            Try

                'Dim privateKey As AsymmetricCipherKeyPair
                If Left(privateKeyPem, 3) <> "---" Then
                    privateKeyPem = "-----BEGIN PRIVATE KEY-----" & vbLf & ConvertToPemFormat(privateKeyPem) & vbLf & "-----END PRIVATE KEY-----"
                End If

                ' Read the private key properly
                Dim privateKeyObject As Object
                Using reader As New StringReader(privateKeyPem)
                    Dim pemReader = New Org.BouncyCastle.OpenSsl.PemReader(reader)
                    privateKeyObject = pemReader.ReadObject()
                End Using

                ' Determine if we have a key pair or just a private key
                Dim privateKeyParams As RsaKeyParameters
                If TypeOf privateKeyObject Is AsymmetricCipherKeyPair Then
                    Dim keyPair = CType(privateKeyObject, AsymmetricCipherKeyPair)
                    privateKeyParams = CType(keyPair.Private, RsaKeyParameters)
                ElseIf TypeOf privateKeyObject Is RsaPrivateCrtKeyParameters Then
                    privateKeyParams = CType(privateKeyObject, RsaPrivateCrtKeyParameters)
                Else
                    Throw New ApplicationException("Invalid private key format.")
                End If

                ' Convert unsigned JWT to bytes
                Dim unsignedDataBytes = Encoding.UTF8.GetBytes(jwtUnsigned)

                ' Use SHA256 for the signature
                Dim signer = SignerUtilities.GetSigner("SHA256withRSA")
                signer.Init(True, privateKeyParams)
                signer.BlockUpdate(unsignedDataBytes, 0, unsignedDataBytes.Length)
                Dim signatureBytes = signer.GenerateSignature()

                ' Base64 encode the signature
                Dim base64Signature = System.Convert.ToBase64String(signatureBytes)

                Return base64Signature
            Catch ex As Exception
                Throw New ApplicationException("Error signing JWT: " & ex.Message, ex)
            End Try
        End Function

        Private Shared Function ConvertToPemFormat(rawKey As String) As String
            Dim sb As New StringBuilder()
            Dim index As Integer = 0
            While index < rawKey.Length
                Dim chunk As String = rawKey.Substring(index, Math.Min(64, rawKey.Length - index))
                sb.AppendLine(chunk)
                index += 64
            End While
            Return sb.ToString().Trim()
        End Function



    End Class
End Namespace
