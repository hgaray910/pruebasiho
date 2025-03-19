Attribute VB_Name = "modFirmaSAT"
' $Id: basFirmaSAT.bas $

' This module contains the full list of declaration statements
' for FirmaSAT v5.3.
' VB6/VBA version.
' Last updated:
'   $Date: 2014-01-26 08:23:00 $
'   $Revision: 5.3.0 $

'************************* COPYRIGHT NOTICE*************************
' Copyright (c) 2010-14 DI Management Services Pty Limited.
' All rights reserved.
' This code may only be used by licensed users of FirmaSAT.
' Refer to licence for conditions of use.
' See <http://www.cryptosys.net/fsa/>
' This copyright notice must always be left intact.
'****************** END OF COPYRIGHT NOTICE*************************

Option Explicit
Option Base 0

' OPTIONS FLAGS
Public Const SAT_GEN_PLATFORM      As Long = &H40
Public Const SAT_HASH_DEFAULT      As Long = 0        ' Default is SHA-1
Public Const SAT_HASH_MD5          As Long = &H10     ' For legacy apps pre-2011
Public Const SAT_HASH_SHA1         As Long = &H20     ' Deprecated - use default 0 instead
Public Const SAT_DATE_NOTBEFORE    As Long = &H1000
Public Const SAT_TFD               As Long = &H8000   ' New in [v4.0]
Public Const SAT_XML_LOOSE         As Long = &H4000   ' New option in [v5.0]
Public Const SAT_XML_STRICT        As Long = 0        ' New default in [v5.0]
Public Const SAT_ENCODE_UTF8       As Long = 0
Public Const SAT_ENCODE_LATIN1     As Long = 1
Public Const SAT_FILE_NO_BOM       As Long = &H2000   ' New in [v5.2]
Public Const SAT_KEY_ENCRYPTED     As Long = &H10000  ' New in [v5.3]
Public Const SAT_XML_EMPTYELEMTAG  As Long = &H20000  ' New in [v5.3]

' CONSTANTS
Public Const SAT_MAX_HASH_CHARS    As Long = 40
Public Const SAT_MAX_ERROR_CHARS   As Long = (4073 - 1)  ' Added [v5.2]

' ENUMERATION
Public Enum HashAlgorithm
    hashMD5 = SAT_HASH_MD5
    hashSHA1 = SAT_HASH_SHA1
End Enum

' DIAGNOSTIC FUNCTIONS
Public Declare Function SAT_Version Lib "diFirmaSAT2.dll" () As Long
Public Declare Function SAT_CompileTime Lib "diFirmaSAT2.dll" (ByVal strOutput As String, ByVal nOutChars As Long) As Long
Public Declare Function SAT_ModuleName Lib "diFirmaSAT2.dll" (ByVal strOutput As String, ByVal nOutChars As Long, ByVal reserved As Long) As Long
Public Declare Function SAT_LicenceType Lib "diFirmaSAT2.dll" () As Long

' ERROR-RELATED FUNCTIONS
Public Declare Function SAT_LastError Lib "diFirmaSAT2.dll" (ByVal strErrMsg As String, ByVal nMsgLen As Long) As Long
Public Declare Function SAT_ErrorLookup Lib "diFirmaSAT2.dll" (ByVal strErrMsg As String, ByVal nMsgLen As Long, ByVal nErrCode As Long) As Long

' OLD CRYPTOSYS PKI INTERROGATE FUNCTIONS -- REDUNDANT AS OF [v4.0] (because CryptoSys PKI is no longer required)
Public Declare Function SAT_PKIVersion Lib "diFirmaSAT2.dll" () As Long
Public Declare Function SAT_PKICompileTime Lib "diFirmaSAT2.dll" (ByVal strOutput As String, ByVal nOutChars As Long) As Long
Public Declare Function SAT_PKIModuleName Lib "diFirmaSAT2.dll" (ByVal strOutput As String, ByVal nOutChars As Long, ByVal reserved As Long) As Long

' SAT XML FUNCTIONS
Public Declare Function SAT_MakePipeStringFromXml Lib "diFirmaSAT2.dll" (ByVal strOut As String, ByVal nOutChars As Long, ByVal strXmlFile As String, ByVal nOptions As Long) As Long
Public Declare Function SAT_MakeSignatureFromXml Lib "diFirmaSAT2.dll" (ByVal strOut As String, ByVal nOutChars As Long, ByVal strXmlFile As String, ByVal strKeyFile As String, ByVal strPassword As String) As Long
Public Declare Function SAT_MakeSignatureFromXmlEx Lib "diFirmaSAT2.dll" (ByVal strOut As String, ByVal nOutChars As Long, ByVal strXmlFile As String, ByVal strKeyFile As String, ByVal strPassword As String, ByVal nOptions As Long) As Long
Public Declare Function SAT_ValidateXml Lib "diFirmaSAT2.dll" (ByVal strXmlFile As String, ByVal nOptions As Long) As Long
Public Declare Function SAT_VerifySignature Lib "diFirmaSAT2.dll" (ByVal strXmlFile As String, ByVal strCertFile As String, ByVal nOptions As Long) As Long
Public Declare Function SAT_SignXml Lib "diFirmaSAT2.dll" (ByVal strOutputFile As String, ByVal strInputXmlFile As String, ByVal strKeyFile As String, ByVal strPassword As String, ByVal strCertFile As String, ByVal nOptions As Long) As Long
Public Declare Function SAT_GetXmlAttribute Lib "diFirmaSAT2.dll" (ByVal strOut As String, ByVal nOutChars As Long, ByVal strXmlFile As String, ByVal strAttribute As String, ByVal strElement As String) As Long
Public Declare Function SAT_MakeDigestFromXml Lib "diFirmaSAT2.dll" (ByVal strOut As String, ByVal nOutChars As Long, ByVal strXmlFile As String, ByVal nOptions As Long) As Long
Public Declare Function SAT_ExtractDigestFromSignature Lib "diFirmaSAT2.dll" (ByVal strOut As String, ByVal nOutChars As Long, ByVal strXmlFile As String, ByVal strCertFile As String, ByVal nOptions As Long) As Long
Public Declare Function SAT_GetCertNumber Lib "diFirmaSAT2.dll" (ByVal strOut As String, ByVal nOutChars As Long, ByVal strFileName As String, ByVal nOptions As Long) As Long
Public Declare Function SAT_GetCertExpiry Lib "diFirmaSAT2.dll" (ByVal strOut As String, ByVal nOutChars As Long, ByVal strFileName As String, ByVal nOptions As Long) As Long
Public Declare Function SAT_GetCertAsString Lib "diFirmaSAT2.dll" (ByVal strOut As String, ByVal nOutChars As Long, ByVal strFileName As String, ByVal nOptions As Long) As Long
Public Declare Function SAT_CheckKeyAndCert Lib "diFirmaSAT2.dll" (ByVal strKeyFile As String, ByVal strPassword As String, ByVal strCertFile As String, ByVal nOptions As Long) As Long
Public Declare Function SAT_XmlReceiptVersion Lib "diFirmaSAT2.dll" (ByVal strXmlFile As String, ByVal nOptions As Long) As Long
Public Declare Function SAT_FixBOM Lib "diFirmaSAT2.dll" (ByVal strOutputFile As String, ByVal strInputFile As String, ByVal nOptions As Long) As Long
Public Declare Function SAT_GetKeyAsString Lib "diFirmaSAT2.dll" (ByVal strOut As String, ByVal nOutChars As Long, ByVal strKeyFile As String, ByVal strPassword As String, ByVal nOptions As Long) As Long
Public Declare Function SAT_WritePfxFile Lib "diFirmaSAT2.dll" (ByVal strOutputFile As String, ByVal strPfxPassword As String, ByVal strKeyFile As String, ByVal strKeyPassword As String, ByVal strCertFile As String, ByVal nOptions As Long) As Long
Public Declare Function SAT_QueryCert Lib "diFirmaSAT2.dll" (ByVal strOut As String, ByVal nOutChars As Long, ByVal strFileName As String, ByVal strQuery As String, ByVal nOptions As Long) As Long
' New in [v5.3]
Public Declare Function SAT_GetXmlAttributeEx Lib "diFirmaSAT2.dll" (ByVal strOut As String, ByVal nOutChars As Long, ByVal strXmlFile As String, ByVal strAttribute As String, ByVal strElement As String, ByVal nOptions As Long) As Long
Public Declare Function SAT_SignXmlToString Lib "diFirmaSAT2.dll" (ByVal strOut As String, ByVal nOutChars As Long, ByVal strXmlData As String, ByVal strKeyFile As String, ByVal strPassword As String, ByVal strCertFile As String, ByVal nOptions As Long) As Long
' Alias for VB6
Public Declare Function SAT_SignXmlToBytes Lib "diFirmaSAT2.dll" Alias "SAT_SignXmlToString" (ByRef lpOut As Byte, ByVal nOutBytes As Long, ByVal strXmlData As String, ByVal strKeyFile As String, ByVal strPassword As String, ByVal strCertFile As String, ByVal nOptions As Long) As Long

Public Declare Function SAT_InsertCert Lib "diFirmaSAT2.dll" ( _
    ByVal strOutputFile As String, _
    ByVal strXmlFile As String, _
    ByVal strCertFile As String, _
    ByVal nOptions As Long _
) As Long


' *** END OF FIRMASAT DECLARATIONS

' *****************
' WRAPPER FUNCTIONS
' *****************
' Direct calls to the DLL begin with "SAT_", wrapper functions begin with "sat"
' We choose to provide these wrappers as functions rather than class methods.
' It is a simple matter to convert these wrapper functions into a class should you so desire.

Public Function satModuleName() As String
    Dim nc As Long
    Dim strOut As String
    nc = SAT_ModuleName("", 0, 0)
    If nc <= 0 Then Exit Function
    strOut = String(nc, " ")
    nc = SAT_ModuleName(strOut, nc, 0)
    If nc > 0 Then
        satModuleName = strOut
    End If
End Function

Public Function satPlatform() As String
' NB This will *always* return "Win32" (because VB6 is only 32-bit)
    Dim nc As Long
    Dim strOut As String
    nc = SAT_ModuleName("", 0, SAT_GEN_PLATFORM)
    If nc <= 0 Then Exit Function
    strOut = String(nc, " ")
    nc = SAT_ModuleName(strOut, nc, SAT_GEN_PLATFORM)
    If nc > 0 Then
        satPlatform = strOut
    End If
End Function

Public Function satCompileTime() As String
    Dim nc As Long
    Dim strOut As String
    nc = SAT_CompileTime("", 0)
    If nc <= 0 Then Exit Function
    strOut = String(nc, " ")
    nc = SAT_CompileTime(strOut, nc)
    If nc > 0 Then
        satCompileTime = strOut
    End If
End Function

Public Function satPKIModuleName() As String
' REDUNDANT AS OF [v4.0]
    Dim nc As Long
    Dim strOut As String
    nc = SAT_PKIModuleName("", 0, 0)
    If nc <= 0 Then Exit Function
    strOut = String(nc, " ")
    nc = SAT_PKIModuleName(strOut, nc, 0)
    If nc > 0 Then
        satPKIModuleName = strOut
    End If
End Function

Public Function satPKICompileTime() As String
' REDUNDANT AS OF [v4.0]
    Dim nc As Long
    Dim strOut As String
    nc = SAT_PKICompileTime("", 0)
    If nc <= 0 Then Exit Function
    strOut = String(nc, " ")
    nc = SAT_PKICompileTime(strOut, nc)
    If nc > 0 Then
        satPKICompileTime = strOut
    End If
End Function

Public Function satLastError() As String
    Dim nc As Long
    Dim strOut As String
    nc = SAT_LastError("", 0)
    If nc <= 0 Then Exit Function
    strOut = String(nc, " ")
    nc = SAT_LastError(strOut, nc)
    If nc > 0 Then
        satLastError = strOut
    End If
End Function

Public Function satErrorLookup(nErrCode As Long) As String
    Dim nc As Long
    Dim strOut As String
    strOut = String(255, " ")
    nc = SAT_ErrorLookup(strOut, Len(strOut), nErrCode)
    If nc > 0 Then
        satErrorLookup = Trim(strOut)
    End If
End Function

Public Function satMakePipeStringFromXml(strXmlFile As String) As String
    Dim nc As Long
    Dim strOut As String
    nc = SAT_MakePipeStringFromXml("", 0, strXmlFile, 0)
    If nc <= 0 Then Exit Function
    strOut = String(nc, " ")
    nc = SAT_MakePipeStringFromXml(strOut, nc, strXmlFile, 0)
    If nc > 0 Then
        satMakePipeStringFromXml = Trim(strOut)
    End If
End Function

' [v2.1] Updated to include option for SHA-1
Public Function satMakeDigestFromXml(strXmlFile As String, Optional HashAlg As HashAlgorithm = 0) As String
    Dim nc As Long
    Dim strOut As String
    nc = SAT_MakeDigestFromXml("", 0, strXmlFile, HashAlg)
    If nc <= 0 Then Exit Function
    strOut = String(nc, " ")
    nc = SAT_MakeDigestFromXml(strOut, nc, strXmlFile, HashAlg)
    If nc > 0 Then
        satMakeDigestFromXml = strOut
    End If
End Function

Public Function satExtractDigestFromSignature(strXmlFile As String, Optional strCertFile As String = vbNullString) As String
    Dim nc As Long
    Dim strOut As String
    nc = SAT_ExtractDigestFromSignature("", 0, strXmlFile, strCertFile, 0)
    If nc <= 0 Then Exit Function
    strOut = String(nc, " ")
    nc = SAT_ExtractDigestFromSignature(strOut, nc, strXmlFile, strCertFile, 0)
    If nc > 0 Then
        satExtractDigestFromSignature = strOut
    End If
End Function

Public Function satVerifySignature(strXmlFile As String, Optional strCertFile As String = vbNullString) As Long
    satVerifySignature = SAT_VerifySignature(strXmlFile, strCertFile, 0)
End Function

Public Function satMakeSignatureFromXml(strXmlFile As String, strKeyFile As String, strPassword As String, Optional HashAlg As HashAlgorithm = 0) As String
    Dim nc As Long
    Dim strOut As String
    nc = SAT_MakeSignatureFromXmlEx("", 0, strXmlFile, strKeyFile, strPassword, HashAlg)
    If nc <= 0 Then Exit Function
    strOut = String(nc, " ")
    nc = SAT_MakeSignatureFromXmlEx(strOut, nc, strXmlFile, strKeyFile, strPassword, HashAlg)
    If nc > 0 Then
        satMakeSignatureFromXml = strOut
    End If
End Function

Public Function satGetXmlAttribute(strXmlFile As String, strAttributeName As String, strElementName As String) As String
    Dim nc As Long
    Dim strOut As String
    nc = SAT_GetXmlAttribute("", 0, strXmlFile, strAttributeName, strElementName)
    If nc <= 0 Then Exit Function
    strOut = String(nc, " ")
    nc = SAT_GetXmlAttribute(strOut, nc, strXmlFile, strAttributeName, strElementName)
    If nc > 0 Then
        satGetXmlAttribute = strOut
    End If
End Function

Public Function satGetCertNumber(strFileName As String) As String
    Dim nc As Long
    Dim strOut As String
    nc = SAT_GetCertNumber("", 0, strFileName, 0)
    If nc <= 0 Then Exit Function
    strOut = String(nc, " ")
    nc = SAT_GetCertNumber(strOut, nc, strFileName, 0)
    If nc > 0 Then
        satGetCertNumber = strOut
    End If
End Function

Public Function satGetCertExpiry(strFileName As String) As String
    Dim nc As Long
    Dim strOut As String
    nc = SAT_GetCertExpiry("", 0, strFileName, 0)
    If nc <= 0 Then Exit Function
    strOut = String(nc, " ")
    nc = SAT_GetCertExpiry(strOut, nc, strFileName, 0)
    If nc > 0 Then
        satGetCertExpiry = strOut
    End If
End Function

Public Function satGetCertStart(strFileName As String) As String
' [v3.0] Added option to get certificate start date
' Deprecated as of [v5.1] - use satQueryCert
    Dim nc As Long
    Dim strOut As String
    nc = SAT_GetCertExpiry("", 0, strFileName, SAT_DATE_NOTBEFORE)
    If nc <= 0 Then Exit Function
    strOut = String(nc, " ")
    nc = SAT_GetCertExpiry(strOut, nc, strFileName, SAT_DATE_NOTBEFORE)
    If nc > 0 Then
        satGetCertStart = strOut
    End If
End Function

Public Function satGetCertAsString(strFileName As String) As String
    Dim nc As Long
    Dim strOut As String
    nc = SAT_GetCertAsString("", 0, strFileName, 0)
    If nc <= 0 Then Exit Function
    strOut = String(nc, " ")
    nc = SAT_GetCertAsString(strOut, nc, strFileName, 0)
    If nc > 0 Then
        satGetCertAsString = strOut
    End If
End Function

Public Function satGetKeyAsString(strFileName As String, strPassword As String) As String
' Returns unencrypted key as a plain base64 string
    Dim nc As Long
    Dim strOut As String
    nc = SAT_GetKeyAsString("", 0, strFileName, strPassword, 0)
    If nc <= 0 Then Exit Function
    strOut = String(nc, " ")
    nc = SAT_GetKeyAsString(strOut, nc, strFileName, strPassword, 0)
    If nc > 0 Then
        satGetKeyAsString = strOut
    End If
End Function

Public Function satGetKeyAsPEMString(strFileName As String, strPassword As String) As String
' Returns encrypted private key as PEM string
    Dim nc As Long
    Dim strOut As String
    nc = SAT_GetKeyAsString("", 0, strFileName, strPassword, SAT_KEY_ENCRYPTED)
    If nc <= 0 Then Exit Function
    strOut = String(nc, " ")
    nc = SAT_GetKeyAsString(strOut, nc, strFileName, strPassword, SAT_KEY_ENCRYPTED)
    If nc > 0 Then
        satGetKeyAsPEMString = strOut
    End If
End Function

Public Function satQueryCert(strFileName As String, strQuery As String) As String
    Dim nc As Long
    Dim strOut As String
    nc = SAT_QueryCert("", 0, strFileName, strQuery, 0)
    If nc <= 0 Then Exit Function
    strOut = String(nc, " ")
    nc = SAT_QueryCert(strOut, nc, strFileName, strQuery, 0)
    If nc > 0 Then
        satQueryCert = strOut
    End If
End Function

Public Function satSignXmlToString(strXmlData As String, strKeyFile As String, strPassword As String, strCertFile As String, nOptions As Long) As String
    Dim nc As Long
    Dim strOut As String
    nc = SAT_SignXmlToString("", 0, strXmlData, strKeyFile, strPassword, strCertFile, nOptions)
    If nc <= 0 Then Exit Function
    strOut = String(nc, " ")
    nc = SAT_SignXmlToString(strOut, nc, strXmlData, strKeyFile, strPassword, strCertFile, nOptions)
    If nc > 0 Then
        satSignXmlToString = strOut
    End If
End Function


' **********************************************
' [v4.0] Variants for TimbreFiscalDigital (TFD)
' **********************************************

Public Function tfdMakePipeStringFromXml(strXmlFile As String) As String
    Dim nc As Long
    Dim strOut As String
    nc = SAT_MakePipeStringFromXml("", 0, strXmlFile, SAT_TFD)
    If nc <= 0 Then Exit Function
    strOut = String(nc, " ")
    nc = SAT_MakePipeStringFromXml(strOut, nc, strXmlFile, SAT_TFD)
    If nc > 0 Then
        tfdMakePipeStringFromXml = Trim(strOut)
    End If
End Function

Public Function tfdMakeDigestFromXml(strXmlFile As String, Optional HashAlg As HashAlgorithm = 0) As String
    Dim nc As Long
    Dim strOut As String
    nc = SAT_MakeDigestFromXml("", 0, strXmlFile, HashAlg + SAT_TFD)
    If nc <= 0 Then Exit Function
    strOut = String(nc, " ")
    nc = SAT_MakeDigestFromXml(strOut, nc, strXmlFile, HashAlg + SAT_TFD)
    If nc > 0 Then
        tfdMakeDigestFromXml = strOut
    End If
End Function

Public Function tfdExtractDigestFromSignature(strXmlFile As String, strCertFile As String) As String
' NB Certificate file is mandatory for TFD.
    Dim nc As Long
    Dim strOut As String
    nc = SAT_ExtractDigestFromSignature("", 0, strXmlFile, strCertFile, SAT_TFD)
    If nc <= 0 Then Exit Function
    strOut = String(nc, " ")
    nc = SAT_ExtractDigestFromSignature(strOut, nc, strXmlFile, strCertFile, SAT_TFD)
    If nc > 0 Then
        tfdExtractDigestFromSignature = strOut
    End If
End Function

Public Function tfdMakeSignatureFromXml(strXmlFile As String, strKeyFile As String, strPassword As String, Optional HashAlg As HashAlgorithm = 0) As String
    Dim nc As Long
    Dim strOut As String
    nc = SAT_MakeSignatureFromXmlEx("", 0, strXmlFile, strKeyFile, strPassword, HashAlg + SAT_TFD)
    If nc <= 0 Then Exit Function
    strOut = String(nc, " ")
    nc = SAT_MakeSignatureFromXmlEx(strOut, nc, strXmlFile, strKeyFile, strPassword, HashAlg + SAT_TFD)
    If nc > 0 Then
        tfdMakeSignatureFromXml = strOut
    End If
End Function

Public Function tfdVerifySignature(strXmlFile As String, strCertFile As String) As Long
' NB Certificate file is mandatory for TFD.
    tfdVerifySignature = SAT_VerifySignature(strXmlFile, strCertFile, SAT_TFD)
End Function

