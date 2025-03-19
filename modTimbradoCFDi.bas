Attribute VB_Name = "modTimbradoCFDi"
Option Explicit

Dim vlstrRefID As String 'ID para el archivo Request Timbrado (RFC Emisor (3 primeros y 3 últimos) + Tipo + Serie||Folio)
Dim vlstrRutaRequestTimbrado As String 'Ruta para el archivo Request Timbrado
Public vlstrMensajeErrorCancelacionCFDi As String 'se requiere para saber si tiene error la cancelacion del cfdi
Public vlintTipoMensajeErrorCancelacionCFDi As Variant 'se requiere para saber si tiene error la cancelacion del cfdi
'para guardar el log de un error de timbrado después de hacer rollback
Public Type typLogTimbrado
       vgXMLREQUEST As String
       vgRESPUESTAWS As String
       vgIDREFERENCIA As String
       vgMENSAJEERROR As String
End Type
Public vlArrLogTimbrado() As typLogTimbrado
Public intContadorArrlogTimbrado As Integer
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Type typCFDiPendienteCancelar
          INTCOMPROBANTE        As Long
          VCHTIPOCOMPROBANTE    As String
          CHRFOLIOCOMPROBANTE   As String
          dtmFechahora          As Date
          MNYSUBTOTAL           As Double
          MNYDESCUENTO          As Double
          MNYIVA                As Double
          MNYTOTAL              As Double
          BITPESOS              As Integer
          CHRNOMBRE             As String
          SMIDEPARTAMENTO       As Integer
          INTMOVPACIENTE        As Long
          CHRTIPOPACIENTE       As String
          intNumCliente         As Long
          VCHRFCEMISOR          As String
          VCHRFCRECEPTOR        As String
          INTIDCOMPROBANTE      As Long
          VCHUUID               As String
          CLBCOMPROBANTEFISCAL  As String
End Type
Public intcontadorCFDiPendienteCancelar As Integer
Public vlArrCFDiPendienteCancelar() As typCFDiPendienteCancelar
Dim vgStrSolicitudEnPlanoProdigia As String
Dim vgStrSolicitudEnPlanoCancProdigia As String
Public vgStrCertBase64 As String
Public vgStrKeyBase64 As String
Public vgStrKeyToKey As String

Public Sub TimbrarCFDI(DOMFacturaXMLsinTimbrar As MSXML2.DOMDocument, RefID As String, strRutaRequestTimbrado As String)
On Error GoTo NotificaError:
    
    Dim rsPAC As New ADODB.Recordset
    Dim intPAC As Integer

    'Se obtiene el PAC con el que se realizará el proceso de timbrado (Buzón Fiscal: INTIDPAC = 1) (PAX: INTIDPAC = 2)
1   Set rsPAC = frsEjecuta_SP(CStr(vgintClaveEmpresaContable), "SP_GNSELCONFIGPAC")
2   If rsPAC.RecordCount > 0 Then
3      intPAC = Val(rsPAC!PAC)
        
4      Select Case intPAC
            Case 1 '(Buzón Fiscal: INTIDPAC = 1)
6                Call pTimbradoBuzonFiscal(DOMFacturaXMLsinTimbrar, RefID, strRutaRequestTimbrado)
            Case 2 '(PAX: INTIDPAC = 2)
8                Call pTimbradoPax(DOMFacturaXMLsinTimbrar, RefID, strRutaRequestTimbrado)
            Case 3 '(Prodigia: INTIDPAC = 3)
                 Call pTimbradoProdigia(DOMFacturaXMLsinTimbrar, RefID, strRutaRequestTimbrado)
       End Select
    Else
    '|  ¡No se ha configurado un PAC activo para realizar el servicio de timbrado!
11        MsgBox SIHOMsg(1155), vbCritical, "Mensaje"
12        GoTo NotificaError
    End If
    
Exit Sub
NotificaError:
    ' si hay timbre o no aqui no afecta mucho el resultado de los errores
    If Not CFDiblnBanError Then 'si todavia no hay ningun error en el proceso
       CFDiblnBanError = True
       CFDistrProcesoError = "TimbrarCFDI"
       If Err.Number <> 0 Then
          CFDiintLineaError = Erl()
          CFDistrDescripError = Err.Description
          CFDilngNumError = Err.Number
          CFDiMostrarMensajeError = True
          Err.Clear 'limpiamos el error para que no se active en otros procesos
       Else 'error de que no hay pac configurado, viene del Goto
          CFDiintLineaError = 12
          CFDistrDescripError = SIHOMsg(1155)
          CFDilngNumError = -1
          CFDiMostrarMensajeError = False
       End If
    End If
End Sub
' Aqui es donde se llama al Web Service de Timbre Fiscal
Public Sub pTimbradoBuzonFiscal(DOMFacturaXMLsinTimbre As MSXML2.DOMDocument, strRefID As String, strRutaRequestTimbrado As String)
On Error GoTo NotificaErrorTimbre:
    
    Dim DOMRequestXML As MSXML2.DOMDocument
    Dim rsConexion As New ADODB.Recordset
    Dim strURLWSTimbrado As String
    Dim strURLWMTimbrado As String
    Dim SerializerWS As SoapSerializer30 'Para serializar el XML
    Dim ReaderRespuestaWS As SoapReader30      'Para leer la respuesta del WebService
    Dim ConectorWS As ISoapConnector 'Para conectarse al WebService
    Dim intlineaGoto As Integer
    Dim strUUId As String
    
    intlineaGoto = 0
    
    'Se especifican las rutas para la conexión con el servicio de timbrado
1    Set rsConexion = frsEjecuta_SP(CStr(vgintClaveEmpresaContable), "SP_GNSELCONFIGPAC")
       
    'Se especifica el valor de la URL de conexión al Web Service
    '(Pruebas = "https://demotf.buzonfiscal.com/timbrado", Producción = "https://tf.buzonfiscal.com/timbrado")
2    strURLWSTimbrado = rsConexion!URLWSTimbrado
3    strURLWMTimbrado = rsConexion!URLWMTimbrado
       
    'Se especifica la ruta para la creación del archivo Request del Timbrado
4    vlstrRutaRequestTimbrado = Trim(strRutaRequestTimbrado)
    
    'Se especifica el valor del strRefID para el archivo Request Timbrado (RFC Emisor (3 primeros y 3 últimos) + Tipo + Serie||Folio)
5    vlstrRefID = Trim(strRefID)
    
    'Se forma el archivo Request del Timbrado
6    Set DOMRequestXML = fnDOMRequestXMLBuzonFiscal(DOMFacturaXMLsinTimbre)
    
7    Set ConectorWS = New HttpConnector30
    ' La URL que atenderá nuestra solicitud
8    ConectorWS.Property("EndPointURL") = strURLWSTimbrado
                
    ' Ruta del WebMethod para el timbrado ("http://www.buzonfiscal.com/TimbradoCFDI/timbradoCFD")
9    ConectorWS.Property("SoapAction") = strURLWMTimbrado
    
    'TimeOUT
10    ConectorWS.Property("Timeout") = "300000"

    ' El certificado debe estar almacenado en la cuenta del usuario actual de Windows, el cual extraemos desde el XML.
11    ConectorWS.Property("SSLClientCertificateName") = DOMFacturaXMLsinTimbre.selectSingleNode("cfdi:Comprobante/cfdi:Emisor/@rfc").Text
      '########################################## INICIA CONEXIÓN CON EL SERVICIO DE TIMBRADO ##########################################
12    ConectorWS.Connect
13    ConectorWS.BeginMessage
14        Set SerializerWS = New SoapSerializer30
15        SerializerWS.Init ConectorWS.InputStream
16        SerializerWS.StartEnvelope
17            SerializerWS.StartBody
18                SerializerWS.WriteXml DOMRequestXML.xml
19            SerializerWS.EndBody
20        SerializerWS.EndEnvelope
21    ConectorWS.EndMessage
22    Set ReaderRespuestaWS = New SoapReader30
23    ReaderRespuestaWS.Load ConectorWS.OutputStream
      '########################################## FINAL DE CONEXIÓN CON EL SERVICIO DE TIMBRADO ##########################################
24    If Not ReaderRespuestaWS.Fault Is Nothing Then
         Dim strMensajeError As String
25       If Not ReaderRespuestaWS.FaultDetail Is Nothing Then
            Dim DOMNodoCodigoError As MSXML2.IXMLDOMNode
           'Se obtiene el codigo del error devuelto por Timbre Fiscal
26          Set DOMNodoCodigoError = ReaderRespuestaWS.FaultDetail.childNodes(0).Attributes.getNamedItem("codigo")
           'Se captura el error
27          If DOMNodoCodigoError Is Nothing Then
28             strMensajeError = "Ocurrió un error al solicitar el servicio de timbrado" & vbNewLine & vbNewLine & _
                                 "Número de error: 1000" & vbNewLine & _
                                 "Descripción: Error de timbrado de nivel de capa 1"
29             pLogTimbrado 0, DOMRequestXML.Text, strMensajeError, Mid(strRefID, 14, Len(strRefID) - 13), strMensajeError
30             CFDiintResultadoTimbrado = 1 'queda pendiente el timbrado
               intlineaGoto = 31
31             GoTo NotificaErrorTimbre
            Else
32             CFDiintResultadoTimbrado = 2
33             strMensajeError = "Ocurrió un error al solicitar el servicio de timbrado" & vbNewLine & vbNewLine & _
                                            "Número de error: " & DOMNodoCodigoError.Text & vbNewLine & _
                                            "Descripción: " & ReaderRespuestaWS.FaultString.Text
34             pLogTimbrado 0, DOMRequestXML.Text, ReaderRespuestaWS.FaultString.Text, Mid(strRefID, 14, Len(strRefID) - 13), strMensajeError
               intlineaGoto = 35
35             GoTo NotificaErrorTimbre
            End If
         End If
      Else
36        Dim DOMElementoRespuestaWS As MSXML2.IXMLDOMElement
37        Set DOMElementoRespuestaWS = ReaderRespuestaWS.Body
39        pAgregarTimbreBuzonFiscal DOMFacturaXMLsinTimbre, DOMElementoRespuestaWS
40        strUUId = DOMFacturaXMLsinTimbre.selectSingleNode("cfdi:Comprobante/cfdi:Complemento/tfd:TimbreFiscalDigital/@UUID").Text
41        If strUUId <> "" Then
42           CFDiblnHaytimbre = True '*tenemos timbre,
43           pLogTimbrado 0, DOMRequestXML.Text, ReaderRespuestaWS.Body.Text, Mid(strRefID, 14, Len(strRefID) - 13), "TIMBRADO EXITOSO"
          Else
44           CFDiintResultadoTimbrado = 1 'error no identificado, queda pendiente el timbre
45           pLogTimbrado 0, DOMRequestXML.Text, ReaderRespuestaWS.Body.Text, Mid(strRefID, 14, Len(strRefID) - 13), "TIMBRADO PENDIENTE"
46           strMensajeError = "Ocurrió un error al solicitar el servicio de timbrado" & vbNewLine & vbNewLine & _
                                "Descripción: No fue posible recuperar el folio fiscal."
47           intlineaGoto = 48
48           GoTo NotificaErrorTimbre
          End If
50        Set DOMElementoRespuestaWS = Nothing
      End If
    
Exit Sub
NotificaErrorTimbre:
       CFDiblnBanError = True
       CFDistrProcesoError = "pTimbradoBuzonFiscal"
    If Err.Number <> 0 Then 'llegó por error del código
       CFDiintResultadoTimbrado = 1 'el proceso queda pendiente de confirmación de timbre no sabemos si alcanzo a timbrar o no
       CFDiintLineaError = Erl()
       CFDistrDescripError = Err.Description
       CFDilngNumError = Err.Number
       strMensajeError = "Ocurrió un error al solicitar el servicio de timbrado" & vbNewLine & vbNewLine & _
                         "Número de error: " & Err.Number & " <" & Err.Description & "> " & " pTimbradoBuzonFiscal, Linea:" & Erl() & vbNewLine & _
                         "Origen: " & Err.Source
       pLogTimbrado 0, DOMRequestXML.Text, "Error: " & Err.Number & " " & Err.Description & " " & "Origen: " & Err.Source, Mid(strRefID, 14, Len(strRefID) - 13), strMensajeError
       Err.Clear 'limpiamos el error para que no se active en los demás procesos
       CFDiMostrarMensajeError = False
    Else ' llego por error del proceso de timbre(puede o no estar pendiente de timbre fiscal)
       CFDiintLineaError = intlineaGoto
       CFDistrDescripError = strMensajeError
       CFDilngNumError = -1
       CFDiMostrarMensajeError = False
       If CFDiintResultadoTimbrado = 2 Then MsgBox strMensajeError, vbCritical + vbOKOnly, "Mensaje"
    End If
End Sub
' Aqui es donde se llama al Web Service de PAX
Public Sub pTimbradoPax(DOMFacturaXMLsinTimbrar As MSXML2.DOMDocument, strRefID As String, strRutaRequestTimbrado As String)
On Error GoTo NotificaErrorTimbre:

    Dim DOMRequestXML As MSXML2.DOMDocument
    Dim rsConexion As New ADODB.Recordset
    Dim strURLWSTimbrado As String
    Dim strURLWMTimbrado As String
    Dim strURLXMLNSTimbrado As String       'Agregado para evitar la comparación que indica si el WS es de pruebas o no
    Dim SerializerWS As SoapSerializer30    'Para serializar el XML
    Dim ReaderRespuestaWS As SoapReader30   'Para leer la respuesta del WebService
    Dim ConectorWS As ISoapConnector        'Para conectarse al WebService
    
    Dim strTipoDocumento As String
    Dim strUsuario As String
    Dim strPassword As String
    Dim strVersion As String
    Dim intEstructura As Integer
    Dim intlineaGoto As Integer
    Dim strUUId As String
    Dim strPrimeraError As String
    intlineaGoto = 0
    
    'Se especifican las rutas para la conexión con el servicio de timbrado
1   Set rsConexion = frsEjecuta_SP(CStr(vgintClaveEmpresaContable), "SP_GNSELCONFIGPAC")
    
2   strURLWSTimbrado = rsConexion!URLWSTimbrado '(Pruebas = "https://test.paxfacturacion.com.mx:453/webservices/wcfRecepcionASMX.asmx", Producción = "https://www.paxfacturacion.mx/webservices/wcfRecepcionasmx.asmx")
3   strURLWMTimbrado = rsConexion!URLWMTimbrado
4   strURLXMLNSTimbrado = rsConexion!URLXMLNSTimbrado  '(Pruebas = "https://test.paxfacturacion.com.mx:453")
5   strUsuario = rsConexion!Usuario
6   strPassword = rsConexion!Password
    If vgstrVersionCFDI = "3.2" Then
        Select Case Mid(strRefID, 14, 2)
            Case "FA"
                strTipoDocumento = "Factura"
            Case "CR"
                strTipoDocumento = "Nota de Crédito"
            Case "CA"
                strTipoDocumento = "Nota de Cargo"
            Case "DO"
                strTipoDocumento = "Recibo de donativos"
            Case "NO"
                strTipoDocumento = "Recibo de Nomina"
            Case Else
                strTipoDocumento = "XX"
        End Select
    Else
        Select Case Mid(strRefID, 14, 2)
            Case "FA"
                strTipoDocumento = "01"
            Case "CR"
                strTipoDocumento = "02"
            Case "CA"
                strTipoDocumento = "03"
            Case "DO"
                strTipoDocumento = "08"
            Case "RE"
                strTipoDocumento = "09"
            Case "NO"
                strTipoDocumento = "10"
            Case "AN"
                strTipoDocumento = "01"
            Case "AA"
                strTipoDocumento = "02"
            Case Else
                strTipoDocumento = "XX"
        End Select
    End If
    intEstructura = 0 'Se especifica el tipo de estructura (0 por default)
    If vgstrVersionCFDI = "3.2" Then
7       strVersion = DOMFacturaXMLsinTimbrar.selectSingleNode("cfdi:Comprobante/@version").Text 'Se especifica la versión
    Else
        strVersion = DOMFacturaXMLsinTimbrar.selectSingleNode("cfdi:Comprobante/@Version").Text 'Se especifica la versión
    End If
    
    'Se especifica la ruta para la creación del archivo Request del Timbrado
8   vlstrRutaRequestTimbrado = Trim(strRutaRequestTimbrado)

    'Se forma el archivo Request del Timbrado
9   Set DOMRequestXML = fnDOMRequestXMLPAX(DOMFacturaXMLsinTimbrar, intEstructura, strTipoDocumento, strUsuario, strPassword, strVersion, strURLXMLNSTimbrado, False)
   
10  Set ConectorWS = New HttpConnector30
    
    'La URL que atenderá nuestra solicitud
11  ConectorWS.Property("EndPointURL") = strURLWSTimbrado
     
    'Ruta del WebMethod para el timbrado ("https://www.paxfacturacion.com.mx/fnEnviarXML")
12  ConectorWS.Property("SoapAction") = strURLWMTimbrado
    
    'Se configura timeOUT
13  ConectorWS.Property("Timeout") = "300000"
    '########################################## INICIA CONEXIÓN CON EL SERVICIO DE TIMBRADO ####################################
    
14  ConectorWS.Connect
    
15  ConectorWS.BeginMessage
16  Set SerializerWS = New SoapSerializer30
17      SerializerWS.Init ConectorWS.InputStream
        
18      SerializerWS.StartEnvelope
19                   SerializerWS.StartBody
20                   SerializerWS.WriteXml DOMRequestXML.xml
21                   SerializerWS.EndBody
22      SerializerWS.EndEnvelope
        
23      ConectorWS.EndMessage

    
24  Set ReaderRespuestaWS = New SoapReader30
25  ReaderRespuestaWS.Load ConectorWS.OutputStream

    '########################################## FINAL DE CONEXIÓN CON EL SERVICIO DE TIMBRADO ####################################
26  If Not ReaderRespuestaWS.Fault Is Nothing Then
27     Dim strMensajeError As String
28     If Not ReaderRespuestaWS.FaultDetail Is Nothing Then
29        Dim DOMNodoCodigo As MSXML2.IXMLDOMNode
          '------------------------------------------------------------------------------------------------- ERROR A NIVEL 1
          'Se obtiene el codigo del error devuelto por PAX
30        Set DOMNodoCodigo = ReaderRespuestaWS.FaultDetail.childNodes(0).Attributes.getNamedItem("codigo")
          'Se captura el error
31        If DOMNodoCodigo Is Nothing Then
32           CFDiintResultadoTimbrado = 1 'queda pendiente el timbrado
             intlineaGoto = 33
             strMensajeError = "Ocurrió un error al solicitar el servicio de timbrado" & vbNewLine & vbNewLine & _
                               "Número de error: 1000" & vbNewLine & _
                               "Descripción: Error de timbrado de nivel de capa 1"
             pLogTimbrado 0, DOMRequestXML.Text, strMensajeError, Mid(strRefID, 14, Len(strRefID) - 13), strMensajeError
33           GoTo NotificaErrorTimbre
          Else
35           If fDetieneProcesoErrorPAX(CLng(DOMNodoCodigo.Text)) Then
36              CFDiintResultadoTimbrado = 2 'error identificado, no queda pendiente el timbre
37           Else
38              CFDiintResultadoTimbrado = 1 'error no identificado, queda pendiente el timbre
             End If
39           strMensajeError = "Ocurrió un error al solicitar el servicio de timbrado" & vbNewLine & vbNewLine & _
                                            "Número de error: " & DOMNodoCodigo.Text & vbNewLine & _
                                            "Descripción: " & ReaderRespuestaWS.FaultString.Text
40           pLogTimbrado 0, DOMRequestXML.Text, ReaderRespuestaWS.FaultString.Text, Mid(strRefID, 14, Len(strRefID) - 13), strMensajeError
             intlineaGoto = 41
41           GoTo NotificaErrorTimbre
          End If
       End If
    Else
    
        Dim strErrorWS As String
        Dim strDescErrorWS As String
        Dim strRes As String
        strRes = ReaderRespuestaWS.Body.Text
        'Se verifica si el WS regresó algún código de error...
        If vgstrVersionCFDI = "3.2" Then
43          strErrorWS = Trim(Mid(ReaderRespuestaWS.Body.Text, 1, 3))
44          strDescErrorWS = Trim(Mid(ReaderRespuestaWS.Body.Text, 6))
        ElseIf vgstrVersionCFDI = "3.3" Then
            If ReaderRespuestaWS.Body.Text <> "" Then
                strPrimeraError = Split(ReaderRespuestaWS.Body.Text)(0)
                If Len(strPrimeraError) = 3 Then
                    strErrorWS = strPrimeraError
                Else
                    strErrorWS = Mid(strPrimeraError, InStr(strPrimeraError, "33") + 2)
                End If
                strDescErrorWS = Trim(Split(ReaderRespuestaWS.Body.Text, "-")(1))
            End If
        Else
            If strRes <> "" Then
                strPrimeraError = Split(strRes)(0)
                If Len(strPrimeraError) = 9 Or Len(strPrimeraError) = 8 Or Len(strPrimeraError) = 3 Then
                    strErrorWS = strPrimeraError
                End If
                strDescErrorWS = Trim(Split(strRes, "-")(1))
            End If
        End If
        
        strDescErrorWS = Replace(strDescErrorWS, "|", "/")
        
45      If Val(strErrorWS) > 0 Or InStr(strErrorWS, "CFDI") > 0 Or InStr(strErrorWS, "CRP") > 0 Then  'Para determinar si regresó un número de error
        '------------------------------------------------------------------------------------------------- ERROR A NIVEL 2
46         If fDetieneProcesoErrorPAX(strErrorWS) Then
47            CFDiintResultadoTimbrado = 2 'error identificado, no queda pendiente el timbre
           Else
48            CFDiintResultadoTimbrado = 1 'error no identificado, queda pendiente el timbre
           End If
           
49         strMensajeError = "Ocurrió un error al solicitar el servicio de timbrado" & vbNewLine & vbNewLine & _
                             "Número de error: " & strErrorWS & vbNewLine & _
                             "Descripción: " & strDescErrorWS
53         pLogTimbrado 0, DOMRequestXML.Text, ReaderRespuestaWS.Body.Text, Mid(strRefID, 14, Len(strRefID) - 13), strMensajeError
           intlineaGoto = 54
54         GoTo NotificaErrorTimbre
        Else '-------------------------------------------------------------------------------------------- TIMBRADO CORRECTO
55         DOMFacturaXMLsinTimbrar.loadXML (ReaderRespuestaWS.Body.Text)
56         strUUId = DOMFacturaXMLsinTimbrar.selectSingleNode("cfdi:Comprobante/cfdi:Complemento/tfd:TimbreFiscalDigital/@UUID").Text
57         If strUUId <> "" Then
58            CFDiblnHaytimbre = True '*tenemos timbre,
59            pLogTimbrado 0, DOMRequestXML.Text, ReaderRespuestaWS.Body.Text, Mid(strRefID, 14, Len(strRefID) - 13), "TIMBRADO EXITOSO"
           Else
60            CFDiintResultadoTimbrado = 1 'error no identificado, queda pendiente el timbre
61            pLogTimbrado 0, DOMRequestXML.Text, ReaderRespuestaWS.Body.Text, Mid(strRefID, 14, Len(strRefID) - 13), "TIMBRADO PENDIENTE"
              strMensajeError = "Ocurrió un error al solicitar el servicio de timbrado" & vbNewLine & vbNewLine & _
                                "Descripción: No fue posible recuperar el folio fiscal."
              intlineaGoto = 62
62            GoTo NotificaErrorTimbre
           End If
        End If
    End If
Exit Sub
NotificaErrorTimbre:
       CFDiblnBanError = True
       CFDistrProcesoError = "pTimbradoPax"
    If Err.Number <> 0 Then 'llegó por error del código
       CFDiintResultadoTimbrado = 1 'el proceso queda pendiente de confirmación de timbre no sabemos si alcanzo a timbrar o no
       CFDiintLineaError = Erl()
       CFDistrDescripError = Err.Description
       CFDilngNumError = Err.Number
       strMensajeError = "Ocurrió un error al solicitar el servicio de timbrado" & vbNewLine & vbNewLine & _
                         "Número de error: " & Err.Number & " <" & Err.Description & "> " & " pTimbradoPax, Linea:" & Erl() & vbNewLine & _
                         "Origen: " & Err.Source
       pLogTimbrado 0, DOMRequestXML.Text, "Error: " & Err.Number & " " & Err.Description & " " & "Origen: " & Err.Source, Mid(strRefID, 14, Len(strRefID) - 13), strMensajeError
       Err.Clear 'limpiamos el error para que no se active en los demás procesos
       CFDiMostrarMensajeError = False
    Else ' llego por error del proceso de timbre(puede o no estar pendiente de timbre fiscal)
       CFDiintLineaError = intlineaGoto
       CFDistrDescripError = strMensajeError
       CFDilngNumError = -1
       CFDiMostrarMensajeError = False
       If CFDiintResultadoTimbrado = 2 Then MsgBox strMensajeError, vbCritical + vbOKOnly, "Mensaje"
    End If
End Sub

Public Function fDetieneProcesoErrorPAX(vlintCodigoError As String) As Boolean
    fDetieneProcesoErrorPAX = False
              
    If vlintCodigoError <> "0" Then
        If vlintCodigoError <> "999" Then fDetieneProcesoErrorPAX = True 'el error 999 es intente de nuevo por eso este error se manejará como no controlado
    End If
End Function

Public Function fDetieneProcesoErrorProdigia(vlintCodigoError As String) As Boolean
       
    fDetieneProcesoErrorProdigia = False
              
    If vlintCodigoError <> "5" And vlintCodigoError <> "6" Then fDetieneProcesoErrorProdigia = True
       
End Function

' Aqui es donde se llama al Web Service de PAX (Método de ensobretado por SOAP)
Public Sub pTimbradoPax2(DOMFacturaXMLsinTimbrar As MSXML2.DOMDocument, strRefID As String, strRutaRequestTimbrado As String)
1     On Error GoTo NotificaErrorTimbre:

          Dim DOMRequestXML As MSXML2.DOMDocument
          Dim rsConexion As New ADODB.Recordset
          Dim strURLWSTimbrado As String
          Dim strURLWMTimbrado As String
          Dim strURLXMLNSTimbrado As String           'Agregado para evitar la comparación que indica si el WS es de pruebas o no
          Dim ConectorWebService As ISoapConnector    'Para conectarse al WebService
          
          Dim strTipoDocumento As String
          Dim strUsuario As String
          Dim strPassword As String
          Dim strVersion As String
          Dim intEstructura As Integer
          'Dim blnWSPruebas As Boolean
          
          Dim strFacturaXMLTimbrado As String
          Dim MensajeError As String

          'Se especifican las rutas para la conexión con el servicio de timbrado
2         Set rsConexion = frsEjecuta_SP(CStr(vgintClaveEmpresaContable), "SP_GNSELCONFIGPAC")
             
3         strURLWSTimbrado = rsConexion!URLWSTimbrado '(Pruebas = "https://test.paxfacturacion.com.mx:453/webservices/wcfRecepcionASMX.asmx", Producción = "https://www.paxfacturacion.mx/webservices/wcfRecepcionasmx.asmx")
4         strURLWMTimbrado = rsConexion!URLWMTimbrado
5         strURLXMLNSTimbrado = rsConexion!URLXMLNSTimbrado  '(Pruebas = "https://test.paxfacturacion.com.mx:453")
6         strUsuario = rsConexion!Usuario
7         strPassword = rsConexion!Password
          'strTipoDocumento = IIf(Mid(strRefID, 14, 2) = "FA", "factura", "XX")
8         Select Case Mid(strRefID, 14, 2)
              Case "FA"
9                 strTipoDocumento = "Factura"
10            Case "CR"
11                strTipoDocumento = "Nota de Crédito"
12            Case "CA"
13                strTipoDocumento = "Nota de Cargo"
14            Case "DO"
15                strTipoDocumento = "Recibo de donativos"
16            Case Else
17                strTipoDocumento = "XX"
18        End Select
19        intEstructura = 0 'Se especifica el tipo de estructura (0 por default)
20        strVersion = DOMFacturaXMLsinTimbrar.selectSingleNode("cfdi:Comprobante/@version").Text 'Se especifica la versión
          
          'Se especifica la ruta para la creación del archivo Request del Timbrado
21        vlstrRutaRequestTimbrado = Trim(strRutaRequestTimbrado)
          
          'Se define si el archivo request es para el WS de pruebas o para producción
          'blnWSPruebas = IIf(Trim(strURLWSTimbrado) = "https://www.paxfacturacion.mx/webservices/wcfRecepcionasmx.asmx", False, True)
          
          'Se forma el archivo Request del Timbrado
          'Set DOMRequestXML = fnDOMRequestXMLPAX(DOMFacturaXMLsinTimbrar, intEstructura, strTipoDocumento, strUsuario, strPassword, strVersion, blnWSPruebas, True)
22        Set DOMRequestXML = fnDOMRequestXMLPAX(DOMFacturaXMLsinTimbrar, intEstructura, strTipoDocumento, strUsuario, strPassword, strVersion, strURLXMLNSTimbrado, True)
          
              
          '###########################################################################################################################
          '########################################## INICIA CONEXIÓN CON EL SERVICIO DE TIMBRADO ##########################################
          '###########################################################################################################################
          
          '                                           ('Método (modificado) que nos proporcionó Javier Greco como alternativa a la conexión con su webservice de PAX)
          
          'Usar XMLHTTPRequest para enviar la información al servicio Web
          Dim oHttReq As MSXML2.XMLHTTP60
23        Set oHttReq = New MSXML2.XMLHTTP60
          
          Dim RespuestaWSXML As MSXML2.DOMDocument
          
          ' Enviar el comando de forma síncrona (se espera a que se reciba la respuesta)
24        oHttReq.Open "POST", strURLWSTimbrado, False
          
          ' Las cabeceras a enviar al servicio Web (no incluir los dos puntos en el nombre de la cabecera)
25        oHttReq.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
26        oHttReq.setRequestHeader "SOAPAction", strURLWMTimbrado
          
          'Enviar el comando
27        oHttReq.send CStr(DOMRequestXML.xml)
              
          'Se especifica el valor regresado por el WS en el archivo DOM de respuesta para poder manipular mejor su información (XMLDOM)
28        If Trim(oHttReq.statusText) = "OK" Then
29            Set RespuestaWSXML = oHttReq.responseXML
30            strFacturaXMLTimbrado = RespuestaWSXML.Text
31        Else
32            strFacturaXMLTimbrado = "Error"
33        End If
          '###########################################################################################################################
          '########################################## FINAL DE CONEXIÓN CON EL SERVICIO DE TIMBRADO ##########################################
          '#############################################################################################################################
                  
          
34        If strFacturaXMLTimbrado = "Error" Then
      '   -------------------------------------------------------------------------------------------------------- ERROR A NIVEL 1
35            Err.Raise 1000, "Error de timbrado de nivel de capa 1", "Error"
36        Else
              Dim ReaderRespuestaWS As MSXML2.IXMLDOMElement
              Dim strErrorWS As String
              Dim strDescErrorWS As String
              
              'Se verifica si el WS regresó algún código de error...
37            strErrorWS = Trim(Mid(strFacturaXMLTimbrado, 1, 3))
38            strDescErrorWS = Trim(Mid(strFacturaXMLTimbrado, 6))
              
39            If Val(strErrorWS) > 0 Then 'Para determinar si regresó un número de error...
              '------------------------------------------------------------------------------------------------- ERROR A NIVEL 2
40                MensajeError = "Ocurrió un error al solicitar el servicio de timbrado" & vbNewLine & vbNewLine & _
                                              "Número de error: " & strErrorWS & vbNewLine & _
                                              "Descripción: " & strDescErrorWS
                  'Se muestra el mensaje de error en pantalla
41                MsgBox MensajeError, vbCritical + vbOKOnly, "Mensaje"
                  'Se levanta el error
42                Err.Raise 1001, "Error de timbrado de nivel de capa 2", "Error"
43            Else '--------------------------------------------------------------------------------------------- TIMBRADO CORRECTO
44                DOMFacturaXMLsinTimbrar.loadXML (strFacturaXMLTimbrado)
45            End If
46        End If
          
47    Exit Sub
NotificaErrorTimbre:
        If Err.Number > 0 And Err.Number <> 1001 Then
            MensajeError = "Ocurrió un error al solicitar el servicio de timbrado" & vbNewLine & vbNewLine & _
                                          "Número de error: " & Err.Number & vbNewLine & " pTimbradoPax2, Linea:" & Erl() & _
                                          "Origen: " & Err.Source
            MsgBox MensajeError, vbCritical + vbOKOnly, "Mensaje"
        End If
End Sub

'----- Llamar al Web Service de PAX para cancelar el comprobante fiscal -----'
Public Function fblnCancelarCFDiPAX(strUUId As String, strRFC As String, strRFCReceptor As String, strTotalComprobante As String, strRutaAcuse As String)
1     On Error GoTo NotificaErrorTimbre:

          '-- Variables para el request de la cancelación -''
          Dim SerializerWS As SoapSerializer30     'Para serializar el XML
          Dim ConectorWS As ISoapConnector         'Para conectarse al WebService
          Dim DOMRequestXML As MSXML2.DOMDocument  'XML con los datos para la cancelación
          Dim rsConexion As New ADODB.Recordset    'Para leer datos de configuración del PAC
          Dim strURLWSCancelacion As String        'Dirección del Web Service para la cancelación
          Dim strURLWMCancelacion As String        'Dirección del Web Method para la cancelación
          Dim strURLXMLNSTimbrado As String        'Dirección del metodo XMLNS
          Dim strUsuario As String                 'Almacena el usuario del hospital para el uso del WebService
          Dim strPassword As String                'Almacena la contraseña del hospital para el uso del WebService
          Dim intEstructura As Integer             'Indica la estructura la cancelación
          
          '-- Variables para la respuesta del WebService --'
          Dim ReaderRespuestaWS As SoapReader30    'Para leer la respuesta del WebService
          Dim DOMResponseXML As MSXML2.DOMDocument 'XML para leer los datos de la respuesta del WebService
          Dim strCodigoWS As String                'Almacena el código de respuesta del WebService
          Dim strUUIDWS As String                  'Almacena el código del comprobante regresado por el WebService
          Dim strDescripcionWS As String           'Almacena el mensaje de respuesta del WebService
          Dim strFechaWS As String                 'Almacena la fecha de cancelación del WebService
          
          '-- Variables generales --'
          Dim strSentencia As String
          Dim strMensaje As String
          
          '--Variable para saber si se da la opción de cancelar el CFDi en forma manual, cuando el servicio del SAT no esta disponible
          Dim blnBitCancelaCDFiNOSAT As Boolean
          Dim RsbitCancelaCFDiNOSAT As New ADODB.Recordset
          
          Dim strSignatureWS As String
          Dim strPruebaCancelacionNE As String
              
2         vlstrMensajeErrorCancelacionCFDi = ""
3         vlintTipoMensajeErrorCancelacionCFDi = 0
          
          'Se especifican las rutas para la conexión con el servicio de cancelación
4         Set rsConexion = frsEjecuta_SP(CStr(vgintClaveEmpresaContable), "SP_GNSELCONFIGPAC")
          
5         strUsuario = rsConexion!Usuario
6         strPassword = rsConexion!Password
          
7         If vgblnNuevoEsquemaCancelacion Then
8             strURLWSCancelacion = rsConexion!URLWSCancelacionNE   '(Pruebas = "https://test.paxfacturacion.com.mx:476/webservices/wcfCancelaasmx.asmx", Producción = "https://www.paxfacturacion.com.mx:453/webservices/wcfCancelaasmx.asmx")
9             strURLWMCancelacion = rsConexion!URLWMCancelacionNE
10            strURLXMLNSTimbrado = rsConexion!URLXMLNSTimbradoNE   '(Pruebas = "https://test.paxfacturacion.com.mx:476")
          
              'Se forma el archivo Request de la Cancelación
11            Set DOMRequestXML = fnDOMReqCancelXMLPAX_NE(strUUId, strRFC, strRFCReceptor, strTotalComprobante, strUsuario, strPassword, strURLXMLNSTimbrado, False, vgMotivoCancelacion, vgstrFolioFiscalSustituye)
12        Else
13            strURLWSCancelacion = rsConexion!URLWSCancelacion   '(Pruebas = "https://test.paxfacturacion.com.mx:453/webservices/wcfCancelaasmx.asmx", Producción = "https://www.paxfacturacion.com.mx:453/webservices/wcfCancelaasmx.asmx")
14            strURLWMCancelacion = rsConexion!URLWMCancelacion
15            strURLXMLNSTimbrado = rsConexion!URLXMLNSTimbrado   '(Pruebas = "https://test.paxfacturacion.com.mx:453")
16            intEstructura = 0 'Se especifica el tipo de estructura (0 por default)
              
              'Se especifica la ruta para la creación del archivo Request para la cancelación
              'vlstrRutaRequestTimbrado = Trim(strRutaAcuse)
              
              'Se forma el archivo Request de la Cancelación
17            Set DOMRequestXML = fnDOMReqCancelXMLPAX(strUUId, strRFC, intEstructura, strUsuario, strPassword, strURLXMLNSTimbrado, False)
18        End If


          ' se carga parametro para saber si en caso de un error se puede permitir la cancelación del documento en el SIHO para después cancelar en SAT
19        strSentencia = "Select vchvalor from SiParametro where VCHNOMBRE = 'BITCANCELACFDINOSAT' and INTCVEEMPRESACONTABLE = " & vgintClaveEmpresaContable
20        Set RsbitCancelaCFDiNOSAT = frsRegresaRs(strSentencia, adLockOptimistic)
21        If RsbitCancelaCFDiNOSAT.RecordCount > 0 Then
22           blnBitCancelaCDFiNOSAT = IIf(IsNull(RsbitCancelaCFDiNOSAT!vchvalor), 0, Val(RsbitCancelaCFDiNOSAT!vchvalor))
23        Else
24           blnBitCancelaCDFiNOSAT = 0
25        End If
             
          '#####################################################################################################################################
          '########################################### INICIA CONEXIÓN CON EL SERVICIO DE CANCELACIÓN ##########################################
          '#####################################################################################################################################
26        Set ConectorWS = New HttpConnector30
          
27        ConectorWS.Property("EndPointURL") = strURLWSCancelacion    'Ruta del WebService
28        ConectorWS.Property("SoapAction") = strURLWMCancelacion     'Ruta del WebMethod para la cancelación ("https://www.paxfacturacion.com.mx/fnCancelarXML")
29        ConectorWS.Connect
          

30        ConectorWS.BeginMessage
31            Set SerializerWS = New SoapSerializer30
32            SerializerWS.Init ConectorWS.InputStream
              
              'Agrega información del request
33            SerializerWS.StartEnvelope
                  '|  Solo si se está trabajando con el nuevo esquema de cancelación se estableme el xmlns, el esquema viejo lo trae en el body
34                If vgblnNuevoEsquemaCancelacion Then SerializerWS.SoapDefaultNamespace strURLXMLNSTimbrado
35                SerializerWS.StartBody
36                    SerializerWS.WriteXml DOMRequestXML.xml
                      
      '                MsgBox strURLXMLNSTimbrado & Chr(13) & Chr(13) & strURLWSCancelacion & Chr(13) & Chr(13) & strURLWMCancelacion & Chr(13) & Chr(13) & Replace(DOMRequestXML.xml, "><", ">" & Chr(13) & "<"), vbInformation, "Envío de cancelación"
                      
37                SerializerWS.EndBody
38            SerializerWS.EndEnvelope
39        ConectorWS.EndMessage
          
40        Set ReaderRespuestaWS = New SoapReader30
41        ReaderRespuestaWS.Load ConectorWS.OutputStream
          
          '|<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
          '|  SOLO PARA EMULAR RESPUESTAS DEL PAX EN AMBIENTE DE PRUEBAS, DEBERÁ ESTAR COMENTADO EN TODOS LOS EJECUTABLES QUE SE MANDEN AL CLIENTE
          '|>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
42        strPruebaCancelacionNE = fstrPruebaCancelacionNE(strUUId)
43        If strPruebaCancelacionNE <> "" Then ReaderRespuestaWS.Body.Text = strPruebaCancelacionNE
          '|<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
          '|>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
          
          
          '#######################################################################################################################################
          '########################################### FINAL DE CONEXIÓN CON EL SERVICIO DE CANCELACIÓN ##########################################
          '#######################################################################################################################################
          
44       If Not ReaderRespuestaWS.Fault Is Nothing Then
45            If Not ReaderRespuestaWS.FaultDetail Is Nothing Then
                  Dim DOMNodoCodigo As MSXML2.IXMLDOMNode
                  '------------------------------------------------------------------------------------------------ ERROR A NIVEL 1
                  'Se obtiene el codigo del error devuelto por PAX
46                Set DOMNodoCodigo = ReaderRespuestaWS.FaultDetail.childNodes(0).Attributes.getNamedItem("codigo")
                  'Se captura el error
47                If DOMNodoCodigo Is Nothing Then
48                    Err.Raise 1000, "Error de cancelación de nivel de capa 1", "Error"
49                    fblnCancelarCFDiPAX = False
                      
50                    frsEjecuta_SP strUUId & "|1||1|'Error de cancelación de nivel de capa 1'", "Sp_PvPendientesCancelarSAT"
51                Else
52                    If blnBitCancelaCDFiNOSAT Then
53                       fblnCancelarCFDiPAX = True
54                       frsEjecuta_SP strUUId & "|1|PC|0|'Pendiente de cancelar ante el SAT'", "Sp_PvPendientesCancelarSAT" 'Se agrega
55                    Else
56                       strMensaje = "Ocurrió un error al solicitar el servicio de cancelación" & vbNewLine & vbNewLine & _
                                       "Número de error: " & DOMNodoCodigo.Text & vbNewLine & _
                                       "Descripción: " & ReaderRespuestaWS.FaultString.Text
57                       vlstrMensajeErrorCancelacionCFDi = strMensaje
58                       vlintTipoMensajeErrorCancelacionCFDi = vbCritical
59                       fblnCancelarCFDiPAX = False
                         
60                       frsEjecuta_SP strUUId & "|1||1|'Error: '" & Trim(DOMNodoCodigo.Text) & " Descripción: " & Trim(ReaderRespuestaWS.FaultString.Text) & "'", "Sp_PvPendientesCancelarSAT"
61                    End If
62                End If
63            End If
64       Else
65            Set DOMResponseXML = New MSXML2.DOMDocument
66            DOMResponseXML.loadXML (ReaderRespuestaWS.Body.Text)
              'Se verifica si el WS regresó información...
67            Set DOMNodoCodigo = DOMResponseXML.selectSingleNode("Cancelacion/Folios")
68            If Not DOMNodoCodigo Is Nothing Then
69                strCodigoWS = Trim(DOMNodoCodigo.selectSingleNode("UUIDEstatus").Text)
70                strDescripcionWS = DOMNodoCodigo.selectSingleNode("UUIDdescripcion").Text

                  '|  Si se está trabajando con el nuevo esquema de cancelación se realiza la validación con los nuevos códigos
71                If vgblnNuevoEsquemaCancelacion Then
72                    Select Case Val(strCodigoWS)
                          '-------------------------------------------------------------------------------------------------------------------------------
                          '| Este código está aparte porque tiene dos significados diferentes (facepalm)
                          '-------------------------------------------------------------------------------------------------------------------------------
                          Case 201
73                            If UCase(strDescripcionWS) = "201 - COMPROBANTE EN PROCESO DE SER CANCELADO." Then
                                  '| 201 - Comprobante En Proceso de ser Cancelado"
74                                fblnCancelarCFDiPAX = False
                                  '| Elimina el registro en la tabla de comprobantes pendientes de cancelar con el estado PA "Pendiente de autorización"
75                                frsEjecuta_SP strUUId & "|0|PA", "Sp_PvPendientesCancelarSAT" 'Se borra
                                  '| Inserta un registro en la tabla de comprobantes pendientes de cancelar con el estado PA "Pendiente de autorización"
76                                frsEjecuta_SP strUUId & "|1|PA|0|'Pendiente de autorización de cancelación. " & UCase(strDescripcionWS) & "'", "Sp_PvPendientesCancelarSAT" 'Se agrega
77                                strMensaje = "El comprobante se encuentra en proceso de ser cancelado." & vbNewLine & _
                                               "Folio en espera: " & strUUId & vbNewLine
78                                vlintTipoMensajeErrorCancelacionCFDi = vbCritical
79                            Else
                                  '| 201 - Comprobante Cancelado sin Aceptación.
80                                fblnCancelarCFDiPAX = True
                                  'Se graba el archivo Response de cancelación en la ruta especificada, sólo cuando se cancela la primera vez
81                                If strRutaAcuse <> "" Then
82                                    DOMResponseXML.Save strRutaAcuse
83                                End If
                                  
84                                frsEjecuta_SP strUUId & "|0|PA|0|'" & UCase(strDescripcionWS) & "'", "Sp_PvPendientesCancelarSAT" 'Se borra
                                                              
85                                strMensaje = ""
86                                vlintTipoMensajeErrorCancelacionCFDi = vbInformation
87                            End If
                          '-------------------------------------------------------------------------------------------------------------------------------
                          '| Cancelaciones exitosas
                          '-------------------------------------------------------------------------------------------------------------------------------
88                        Case 202, 107, 103
                              '| 202 - Comprobante previamente cancelado
                              '| 107 – El CFDI ha sido Cancelado por Plazo Vencido
                              '| 103 – El CFDI ha sido Cancelado Previamente por Aceptación del Receptor
89                            fblnCancelarCFDiPAX = True
                              '| Elimina el registro en la tabla de comprobantes pendientes de cancelar con el estado PA "Pendiente de autorización"
90                            frsEjecuta_SP strUUId & "|0|PA|0|'" & UCase(strDescripcionWS) & "'", "Sp_PvPendientesCancelarSAT" ' Se borra
                              'Se graba el archivo Response de cancelación en la ruta especificada, sólo cuando se cancela la primera vez
91                            If strRutaAcuse <> "" Then
92                                DOMResponseXML.Save strRutaAcuse
93                            End If
94                            strMensaje = ""
95                            vlintTipoMensajeErrorCancelacionCFDi = vbInformation
                          '-------------------------------------------------------------------------------------------------------------------------------
                          '| Cancelaciones que se quedaron en espera
                          '-------------------------------------------------------------------------------------------------------------------------------
96                        Case 105, 106
                              '| 105 – El CFDI no se puede Cancelar por que tiene Estatus De “En espera De Aceptación”
                              '| 106 – El CFDI no se puede Cancelar por que tiene Estatus de “En Proceso”.
97                            fblnCancelarCFDiPAX = False
                              '| Elimina el registro en la tabla de comprobantes pendientes de cancelar con el estado PA "Pendiente de autorización"
98                            frsEjecuta_SP strUUId & "|0|PA", "Sp_PvPendientesCancelarSAT" ' Se borra
                              '| Inserta un registro en la tabla de comprobantes pendientes de cancelar con el estado PA "Pendiente de autorización"
99                            frsEjecuta_SP strUUId & "|1|PA|0|'" & UCase(strDescripcionWS) & "'", "Sp_PvPendientesCancelarSAT" ' Se agrega
100                           strMensaje = "El comprobante se encuentra en espera de aceptación." & vbNewLine & _
                                           "Folio en espera: " & strUUId & vbNewLine
101                           vlintTipoMensajeErrorCancelacionCFDi = vbExclamation
                          '-------------------------------------------------------------------------------------------------------------------------------
                          '| Cancelaciones rechazadas por el receptor
                          '-------------------------------------------------------------------------------------------------------------------------------
102                       Case 104
                              '| 104 – El CFDI no se puede Cancelar por que fue Rechazado Previamente.
103                           fblnCancelarCFDiPAX = False
                              '| Elimina el registro en la tabla de comprobantes pendientes de cancelar con el estado PA "Pendiente de autorización"
104                           frsEjecuta_SP strUUId & "|0|PA", "Sp_PvPendientesCancelarSAT" ' Se borra
                              '| Inserta un registro en la tabla de comprobantes pendientes de cancelar con el estado CR "Cancelación rechazada"
105                           frsEjecuta_SP strUUId & "|1|CR|0|'" & UCase(strDescripcionWS) & "'", "Sp_PvPendientesCancelarSAT" ' Se agrega
106                           strMensaje = "El comprobante no se puede cancelar por que fue rechazado por el receptor." & vbNewLine & _
                                           "Folio rechazado: " & strUUId & vbNewLine
107                           vlintTipoMensajeErrorCancelacionCFDi = vbCritical
                          '-------------------------------------------------------------------------------------------------------------------------------
                          '|  Código de error genérico para las excepciones no controladas del SAT
                          '-------------------------------------------------------------------------------------------------------------------------------
108                       Case 999
109                           strMensaje = "El servicio del SAT no responde a la petición de cancelación. " & vbNewLine & _
                                           "Por favor intente más tarde."
110                           If blnBitCancelaCDFiNOSAT Then
111                               fblnCancelarCFDiPAX = True
112                               frsEjecuta_SP strUUId & "|1|PC", "Sp_PvPendientesCancelarSAT" ' Se agrega
113                           Else
114                               fblnCancelarCFDiPAX = False
115                               vlintTipoMensajeErrorCancelacionCFDi = vbCritical
116                           End If
                              
117                           frsEjecuta_SP strUUId & "|1||1|'" & UCase(strDescripcionWS) & " - " & Trim(strMensaje) & "'", "Sp_PvPendientesCancelarSAT"
                          '-------------------------------------------------------------------------------------------------------------------------------
                          '| Cancelaciones fallidas
                          '-------------------------------------------------------------------------------------------------------------------------------
118                       Case Else
119                           fblnCancelarCFDiPAX = False
120                           strMensaje = "Ocurrió un error al solicitar la cancelación del comprobante." & vbNewLine & vbNewLine & _
                                           "Número de error: " & strCodigoWS & vbNewLine & _
                                           "Descripción: " & strDescripcionWS
121                           vlintTipoMensajeErrorCancelacionCFDi = vbCritical
                              
122                           frsEjecuta_SP strUUId & "|1||1|'" & UCase(strDescripcionWS) & "'", "Sp_PvPendientesCancelarSAT"
123                   End Select
124                   vlstrMensajeErrorCancelacionCFDi = strMensaje
125               Else
                  
126                   If Val(strCodigoWS) = 201 Or Val(strCodigoWS) = 202 Then
                          '201 - El folio se ha cancelado con éxito.
                          '202 - El CFDI ya había sido cancelado previamente.
                          '-------------------------------------------------------------------------------------------- CANCELACIÓN CORRECTA
127                       fblnCancelarCFDiPAX = True
                          'Se graba el archivo Response de cancelación en la ruta especificada, sólo cuando se cancela la primera vez
128                       If Val(strCodigoWS) = 201 And strRutaAcuse <> "" Then
129                           DOMResponseXML.Save strRutaAcuse
130                       End If
                          
131                       frsEjecuta_SP strUUId & "|1||1|'" & Trim(strDescripcionWS) & "'", "Sp_PvPendientesCancelarSAT"
132                   Else
                          '-------------------------------------------------------------------------------------------- ERROR A NIVEL 2
133                       If Val(strCodigoWS) = 999 Then  'Código de error genérico para las excepciones no controladas del SAT
134                          strDescripcionWS = "El servicio del SAT no responde a la petición de cancelación. " & vbNewLine & _
                                                 "Por favor intente más tarde."
                                                 
135                       End If
136                       If blnBitCancelaCDFiNOSAT Then
137                          fblnCancelarCFDiPAX = True
138                          frsEjecuta_SP strUUId & "|1|PC", "Sp_PvPendientesCancelarSAT" ' Se agrega
139                       Else
140                          strMensaje = "Ocurrió un error al solicitar la cancelación del comprobante." & vbNewLine & vbNewLine & _
                                          "Número de error: " & strCodigoWS & vbNewLine & _
                                          "Descripción: " & strDescripcionWS
141                          vlstrMensajeErrorCancelacionCFDi = strMensaje
142                          fblnCancelarCFDiPAX = False
143                       End If
                          
144                       frsEjecuta_SP strUUId & "|1||1|'" & Val(strCodigoWS) & " - " & Trim(strDescripcionWS) & "'", "Sp_PvPendientesCancelarSAT"
145                   End If
146               End If
147           Else
148               If blnBitCancelaCDFiNOSAT Then
149                  fblnCancelarCFDiPAX = True
150                  frsEjecuta_SP strUUId & "|1|PC|0|'Pendiente de cancelar ante el SAT'", "Sp_PvPendientesCancelarSAT" ' Se agrega
151               Else
152                  vlstrMensajeErrorCancelacionCFDi = "Ocurrió un error al solicitar la cancelación del comprobante." & vbNewLine & vbNewLine & _
                                                        "No se recibió respuesta del web service."
153                  fblnCancelarCFDiPAX = False
                     
154                  frsEjecuta_SP strUUId & "|1||1|'" & vlstrMensajeErrorCancelacionCFDi & "'", "Sp_PvPendientesCancelarSAT"
155               End If
156           End If
157       End If
          
158   Exit Function
NotificaErrorTimbre:
       If Err.Number > 0 And Err.Number <> 1001 Then
          If blnBitCancelaCDFiNOSAT Then
             fblnCancelarCFDiPAX = True
             frsEjecuta_SP strUUId & "|1|PC|0|'NotificaErrorTimbre: " & Err.Number & "'", "Sp_PvPendientesCancelarSAT" ' Se agrega
          Else
             strMensaje = "Ocurrió un error al solicitar la cancelación del comprobante." & vbNewLine & _
                              "Número de error: " & Err.Number & vbNewLine & _
                              "Origen: " & Err.Source & vbNewLine & _
                              " fblnCancelarCFDiPAX, Linea:" & Erl() & vbNewLine & _
                              "Descripción: " & Err.Description
             vlstrMensajeErrorCancelacionCFDi = strMensaje
             fblnCancelarCFDiPAX = False
                
             frsEjecuta_SP strUUId & "|1||1|'" & strMensaje & "'", "Sp_PvPendientesCancelarSAT"
          End If
       End If
End Function

'----- Llamar al Web Service de Prodigia para cancelar el comprobante fiscal -----'
Public Function fblnCancelarCFDiProdigia(strUUId As String, strRFC As String, strRFCReceptor As String, strTotalComprobante As String, strRutaAcuse As String)
1     On Error GoTo NotificaErrorTimbre:

          '-- Variables para el request de la cancelación -''
          Dim SerializerWS As SoapSerializer30     'Para serializar el XML
          Dim SerializerWSQuery As SoapSerializer30 'Para serializar el XML de la consulta
          Dim ConectorWS As ISoapConnector         'Para conectarse al WebService
          Dim ConectorWSQuery As ISoapConnector    'Para conectarse al WebService de consulta
          Dim DOMRequestXML As MSXML2.DOMDocument  'XML con los datos para la cancelación
          Dim DOMQueryXML As MSXML2.DOMDocument    'XML con los datos de la consulta inmediata del estatus del CFDI
          Dim rsConexion As New ADODB.Recordset    'Para leer datos de configuración del PAC
          Dim rsPruebaOReal As New ADODB.Recordset
          Dim strURLWSCancelacion As String        'Dirección del Web Service para la cancelación
          Dim strURLWMCancelacion As String        'Dirección del Web Method para la cancelación
          Dim strURLXMLNSTimbrado As String        'Dirección del metodo XMLNS
          Dim strUsuario As String                 'Almacena el usuario del hospital para el uso del WebService
          Dim strPassword As String                'Almacena la contraseña del hospital para el uso del WebService
          Dim strContrato As String                'Almacena el contrato del hospital para el uso del timbrado
          Dim intEstructura As Integer             'Indica la estructura la cancelación
          
          '-- Variables para la respuesta del WebService --'
          Dim ReaderRespuestaWS As SoapReader30    'Para leer la respuesta del WebService
          Dim ReaderRespuestaWSQuery As SoapReader30 'Para leer la consulta del WebService
          Dim DOMResponseXML As MSXML2.DOMDocument 'XML para leer los datos de la respuesta del WebService
          Dim DOMAcuseXML As MSXML2.DOMDocument    'XML para leer el acuse en la respuesta del WebService
          Dim strCodigoWS As String                'Almacena el código de respuesta del WebService
          Dim strUUIDWS As String                  'Almacena el código del comprobante regresado por el WebService
          Dim strDescripcionWS As String           'Almacena el mensaje de respuesta del WebService
          Dim strFechaWS As String                 'Almacena la fecha de cancelación del WebService
          
          '-- Variables generales --'
          Dim strSentencia As String
          Dim strMensaje As String
          Dim strSQLPrueba As String
          
          '--Variable para saber si se da la opción de cancelar el CFDi en forma manual, cuando el servicio del SAT no esta disponible
          Dim blnBitCancelaCDFiNOSAT As Boolean
          Dim RsbitCancelaCFDiNOSAT As New ADODB.Recordset
          
          Dim strSignatureWS As String
          Dim strPruebaCancelacionNE As String
          
          '----tls1.2---
          Dim httpRequest As WinHttp.WinHttpRequest
          Set httpRequest = New WinHttp.WinHttpRequest
          Dim httpRequestC As WinHttp.WinHttpRequest
          Set httpRequestC = New WinHttp.WinHttpRequest
          '----tls1.2---
              
2         vlstrMensajeErrorCancelacionCFDi = ""
3         vlintTipoMensajeErrorCancelacionCFDi = 0
          
4         strSQLPrueba = "SELECT VCHVALOR FROM SIPARAMETRO WHERE VCHNOMBRE = 'BITFACTURACIONMODOPRUEBA'"
5         Set rsPruebaOReal = frsRegresaRs(strSQLPrueba, adLockReadOnly, adOpenForwardOnly)
          
          'Se especifican las rutas para la conexión con el servicio de cancelación
6         Set rsConexion = frsEjecuta_SP(CStr(vgintClaveEmpresaContable), "SP_GNSELCONFIGPAC")
          
7         strContrato = rsConexion!contrato
8         strUsuario = rsConexion!Usuario
9         strPassword = rsConexion!Password
      
      If vgstrVersionCFDI = "4.0" Then
        strURLWSCancelacion = rsConexion!URLWSTimbrado40 '(Pruebas = "https://test.paxfacturacion.com.mx:453/webservices/wcfRecepcionASMX.asmx", Producción = "https://www.paxfacturacion.mx/webservices/wcfRecepcionasmx.asmx")
      Else
        strURLWSCancelacion = rsConexion!URLWSTimbrado '(Pruebas = "https://test.paxfacturacion.com.mx:453/webservices/wcfRecepcionASMX.asmx", Producción = "https://www.paxfacturacion.mx/webservices/wcfRecepcionasmx.asmx")
      End If

11        If rsPruebaOReal!vchvalor Then
12            strURLWMCancelacion = "cancelarConOpciones"
13        Else
14            strURLWMCancelacion = rsConexion!URLWMCancelacion
15        End If
16        strURLXMLNSTimbrado = rsConexion!URLXMLNSTimbrado


          'Se forma el archivo Request de la Cancelación
17        Set DOMRequestXML = fnDOMReqCancelXMLProdigia_NE(strUUId, strRFC, strRFCReceptor, strTotalComprobante, strContrato, strUsuario, strPassword, strURLWMCancelacion, strURLXMLNSTimbrado, rsPruebaOReal!vchvalor, False, vgMotivoCancelacion, vgstrFolioFiscalSustituye, rsConexion!URLWMConsulta)

            'MsgBox strURLXMLNSTimbrado & Chr(13) & Chr(13) & strURLWSCancelacion & Chr(13) & Chr(13) & strURLWMCancelacion & Chr(13) & Chr(13) & Replace(DOMRequestXML.xml, "><", ">" & Chr(13) & "<"), vbInformation, "Envío de cancelación"
            'MsgBox strURLXMLNSTimbrado & Chr(13) & Chr(13) & strURLWSCancelacion & Chr(13) & Chr(13) & strURLWMCancelacion & Chr(13) & Chr(13) & Replace(DOMRequestXML, "><", ">" & Chr(13) & "<"), vbInformation, "Envío de cancelación"

          ' se carga parametro para saber si en caso de un error se puede permitir la cancelación del documento en el SIHO para después cancelar en SAT
18        strSentencia = "Select vchvalor from SiParametro where VCHNOMBRE = 'BITCANCELACFDINOSAT' and INTCVEEMPRESACONTABLE = " & vgintClaveEmpresaContable
19        Set RsbitCancelaCFDiNOSAT = frsRegresaRs(strSentencia, adLockOptimistic)
20        If RsbitCancelaCFDiNOSAT.RecordCount > 0 Then
21           blnBitCancelaCDFiNOSAT = IIf(IsNull(RsbitCancelaCFDiNOSAT!vchvalor), 0, Val(RsbitCancelaCFDiNOSAT!vchvalor))
22        Else
23           blnBitCancelaCDFiNOSAT = 0
24        End If
             
          '#####################################################################################################################################
          '########################################### INICIA CONEXIÓN CON EL SERVICIO DE CANCELACIÓN ##########################################
          '#####################################################################################################################################
25        Set ConectorWS = New HttpConnector30
          
26        ConectorWS.Property("EndPointURL") = strURLWSCancelacion    'Ruta del WebService
27        ConectorWS.Property("SoapAction") = strURLWMCancelacion     'Ruta del WebMethod para la cancelación ("https://www.paxfacturacion.com.mx/fnCancelarXML")
'28        ConectorWS.Connect
'
'
'29        ConectorWS.BeginMessage
'30            Set SerializerWS = New SoapSerializer30
'31            SerializerWS.Init ConectorWS.InputStream
'
'32            SerializerWS.StartEnvelope
'33                        SerializerWS.StartBody
'34                        SerializerWS.WriteXml DOMRequestXML.xml
'35                        SerializerWS.EndBody
'36            SerializerWS.EndEnvelope
'37            ConectorWS.EndMessage
'
'38        Set ReaderRespuestaWS = New SoapReader30
'39        ReaderRespuestaWS.Load ConectorWS.OutputStream
        '----tls1.2---
         httpRequest.Open "POST", strURLWSCancelacion, False
         httpRequest.setRequestHeader "Content-Type", "text/xml"
         httpRequest.setRequestHeader "SOAPAction", strURLWMCancelacion
         httpRequest.send DOMRequestXML.xml
            
         Set ReaderRespuestaWS = New SoapReader30
         ReaderRespuestaWS.loadXML httpRequest.responseBody
        '----tls1.2---
          
          
          '#####################################################################################################################################
          '############################################# INICIA CONEXIÓN CON EL SERVICIO DE CONSULTA ###########################################
          '#####################################################################################################################################
40        Set ConectorWSQuery = New HttpConnector30
          
41        ConectorWSQuery.Property("EndPointURL") = strURLWSCancelacion    'Ruta del WebService
42        ConectorWSQuery.Property("SoapAction") = rsConexion!URLWMConsulta     'Ruta del WebMethod para la cancelación ("https://www.paxfacturacion.com.mx/fnCancelarXML")
'43        ConectorWSQuery.Connect
'
'44        ConectorWSQuery.BeginMessage
'45            Set SerializerWSQuery = New SoapSerializer30
'46            SerializerWSQuery.Init ConectorWSQuery.InputStream
'
'47            SerializerWSQuery.StartEnvelope
'48                        SerializerWSQuery.StartBody
'49                        SerializerWSQuery.WriteXml vgStrSolicitudEnPlanoCancProdigia
'50                        SerializerWSQuery.EndBody
'51            SerializerWSQuery.EndEnvelope
'52        ConectorWSQuery.EndMessage
'
'53        Set ReaderRespuestaWSQuery = New SoapReader30
'54        ReaderRespuestaWSQuery.Load ConectorWSQuery.OutputStream

        '----tls1.2---
            httpRequestC.Open "POST", strURLWSCancelacion, False
            httpRequestC.setRequestHeader "Content-Type", "text/xml"
            httpRequestC.setRequestHeader "SOAPAction", rsConexion!URLWMConsulta
            httpRequestC.send vgStrSolicitudEnPlanoCancProdigia
            
            Set ReaderRespuestaWS = New SoapReader30
            ReaderRespuestaWS.loadXML httpRequestC.responseBody
        '----tls1.2---
          
          '|<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
          '|  SOLO PARA EMULAR RESPUESTAS DEL PAX EN AMBIENTE DE PRUEBAS, DEBERÁ ESTAR COMENTADO EN TODOS LOS EJECUTABLES QUE SE MANDEN AL CLIENTE
          '|>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
          'strPruebaCancelacionNE = fstrPruebaCancelacionNE(strUUId)
          'If strPruebaCancelacionNE <> "" Then ReaderRespuestaWS.Body.Text = strPruebaCancelacionNE
          '|<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
          '|>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
          
          
          '#######################################################################################################################################
          '########################################### FINAL DE CONEXIÓN CON EL SERVICIO DE CANCELACIÓN ##########################################
          '#######################################################################################################################################
          
55       If Not ReaderRespuestaWS.Fault Is Nothing Then
56            If Not ReaderRespuestaWS.FaultDetail Is Nothing Then
                  Dim DOMNodoCodigo As MSXML2.IXMLDOMNode
                  '------------------------------------------------------------------------------------------------ ERROR A NIVEL 1
                  'Se obtiene el codigo del error devuelto por PAX
57                Set DOMNodoCodigo = ReaderRespuestaWS.FaultDetail.childNodes(0).Attributes.getNamedItem("codigo")
                  'Se captura el error
58                If DOMNodoCodigo Is Nothing Then
59                    Err.Raise 1000, "Error de cancelación de nivel de capa 1", "Error"
60                    fblnCancelarCFDiProdigia = False
                      
61                    frsEjecuta_SP strUUId & "|1||1|'Error de cancelación de nivel de capa 1'", "Sp_PvPendientesCancelarSAT"
62                Else
63                    If blnBitCancelaCDFiNOSAT Then
64                       fblnCancelarCFDiProdigia = True
65                       frsEjecuta_SP strUUId & "|1|PC|0|'Pendiente de cancelar ante el SAT'", "Sp_PvPendientesCancelarSAT" 'Se agrega
66                    Else
67                       strMensaje = "Ocurrió un error al solicitar el servicio de cancelación" & vbNewLine & vbNewLine & _
                                       "Número de error: " & DOMNodoCodigo.Text & vbNewLine & _
                                       "Descripción: " & ReaderRespuestaWS.FaultString.Text
68                       vlstrMensajeErrorCancelacionCFDi = strMensaje
69                       vlintTipoMensajeErrorCancelacionCFDi = vbCritical
70                       fblnCancelarCFDiProdigia = False
                         
71                       frsEjecuta_SP strUUId & "|1||1|'Error: '" & Trim(DOMNodoCodigo.Text) & " Descripción: " & Trim(ReaderRespuestaWS.FaultString.Text) & "'", "Sp_PvPendientesCancelarSAT"
72                    End If
73                End If
74            End If
75       Else
              Dim DOMCodigoQuery As String
76            Set DOMResponseXML = New MSXML2.DOMDocument
77            Set DOMQueryXML = New MSXML2.DOMDocument
78            Set DOMAcuseXML = New MSXML2.DOMDocument

              Dim resp As String
              Dim resp2 As String
              If rsPruebaOReal!vchvalor Then
                 resp = Replace(Replace(Replace(Replace(Replace(httpRequest.responseText, "</return></ns2:cancelarConOpcionesResponse></S:Body></S:Envelope>", ""), "<S:Envelope xmlns:S=""http://schemas.xmlsoap.org/soap/envelope/""><S:Body><ns2:cancelarConOpcionesResponse xmlns:ns2=""timbrado.ws.pade.mx""><return>", ""), "&lt;", "<"), "&gt;", ">"), "<?xml version='1.0' encoding='UTF-8'?><?xml version=""1.0"" encoding=""UTF-8""?>", "")
              Else
                resp = Replace(Replace(Replace(Replace(Replace(httpRequest.responseText, "</return></ns2:cancelarResponse></S:Body></S:Envelope>", ""), "<S:Envelope xmlns:S=""http://schemas.xmlsoap.org/soap/envelope/""><S:Body><ns2:cancelarResponse xmlns:ns2=""timbrado.ws.pade.mx""><return>", ""), "&lt;", "<"), "&gt;", ">"), "<?xml version='1.0' encoding='UTF-8'?><?xml version=""1.0"" encoding=""UTF-8""?>", "")
              End If
                        
              resp2 = Replace(Replace(Replace(Replace(Replace(httpRequestC.responseText, "</return></ns2:consultarEstatusComprobanteResponse></S:Body></S:Envelope>", ""), "<S:Envelope xmlns:S=""http://schemas.xmlsoap.org/soap/envelope/""><S:Body><ns2:consultarEstatusComprobanteResponse xmlns:ns2=""timbrado.ws.pade.mx""><return>", ""), "&lt;", "<"), "&gt;", ">"), "<?xml version='1.0' encoding='UTF-8'?><?xml version=""1.0"" encoding=""UTF-8""?>", "")

79            DOMResponseXML.loadXML (resp)
              'MsgBox (ReaderRespuestaWS.Body.Text), vbOKOnly, "Cancelacion"
80            DOMQueryXML.loadXML (resp2)
              'MsgBox (ReaderRespuestaWSQuery.Body.Text), vbOKOnly, "Consulta"
              'Se verifica si el WS regresó información...
81            Set DOMNodoCodigo = DOMResponseXML.selectSingleNode("servicioCancel/cancelaciones/cancelacion/uuid")
82            DOMCodigoQuery = Trim(DOMQueryXML.selectSingleNode("servicioConsultaComprobante/estatusCfdi").Text)
83            If Not DOMNodoCodigo Is Nothing Then
84                strCodigoWS = Trim(DOMResponseXML.selectSingleNode("servicioCancel/cancelaciones/cancelacion/codigo").Text)
85                strDescripcionWS = DOMResponseXML.selectSingleNode("servicioCancel/cancelaciones/cancelacion/mensaje").Text
86                DOMAcuseXML.loadXML (resp)
87                If rsPruebaOReal!vchvalor = "1" Then
88                    DOMCodigoQuery = "9"
89                End If
                  '    DOMAcuseXML.loadXML Decode(DOMResponseXML.selectSingleNode("servicioCancel/acuseCancelBase64").Text)
                  'End If
90                    Select Case Val(DOMCodigoQuery)
                          '-------------------------------------------------------------------------------------------------------------------------------
                          '| Este código está aparte porque tiene dos significados diferentes (facepalm)
                          '-------------------------------------------------------------------------------------------------------------------------------
                          Case 7
                              '| 7  - Cancelación en proceso
                              '| 96 - Comprobante En Proceso de ser Cancelado"
91                            fblnCancelarCFDiProdigia = False
                              '| Elimina el registro en la tabla de comprobantes pendientes de cancelar con el estado PA "Pendiente de autorización"
92                            frsEjecuta_SP strUUId & "|0|PA", "Sp_PvPendientesCancelarSAT" 'Se borra
                              '| Inserta un registro en la tabla de comprobantes pendientes de cancelar con el estado PA "Pendiente de autorización"
93                            frsEjecuta_SP strUUId & "|1|PA|0|'Pendiente de autorización de cancelación. " & UCase(strCodigoWS & " - " & strDescripcionWS) & "'", "Sp_PvPendientesCancelarSAT" 'Se agrega
94                            strMensaje = "El comprobante se encuentra en proceso de ser cancelado." & vbNewLine & _
                                           "Folio en espera: " & strUUId & vbNewLine
95                            vlintTipoMensajeErrorCancelacionCFDi = vbCritical
96                        Case 9, 6
                              '| 9 - Comprobante Cancelado sin Aceptación.
97                            fblnCancelarCFDiProdigia = True
                              'Se graba el archivo Response de cancelación en la ruta especificada, sólo cuando se cancela la primera vez
98                            If strRutaAcuse <> "" Then
99                                DOMAcuseXML.Save strRutaAcuse
100                           End If
                              
101                           frsEjecuta_SP strUUId & "|0|PA|0|'" & UCase(strCodigoWS & " - " & strDescripcionWS) & "'", "Sp_PvPendientesCancelarSAT" 'Se borra
                                                          
102                           strMensaje = ""
103                           vlintTipoMensajeErrorCancelacionCFDi = vbInformation

                          '-------------------------------------------------------------------------------------------------------------------------------
                          '| Cancelaciones exitosas
                          '-------------------------------------------------------------------------------------------------------------------------------
104                       Case 8, 10
                              '| 202 - Comprobante previamente cancelado
                              '| 107|98 – El CFDI ha sido Cancelado por Plazo Vencido
                              '| 103|103 – El CFDI ha sido Cancelado Previamente por Aceptación del Receptor
105                           fblnCancelarCFDiProdigia = True
                              '| Elimina el registro en la tabla de comprobantes pendientes de cancelar con el estado PA "Pendiente de autorización"
106                           frsEjecuta_SP strUUId & "|0|PA|0|'" & UCase(strCodigoWS & " - " & strDescripcionWS) & "'", "Sp_PvPendientesCancelarSAT" ' Se borra
                              'Se graba el archivo Response de cancelación en la ruta especificada, sólo cuando se cancela la primera vez
107                           If strRutaAcuse <> "" Then
108                               DOMAcuseXML.Save strRutaAcuse
109                           End If
110                           strMensaje = ""
111                           vlintTipoMensajeErrorCancelacionCFDi = vbInformation
                          '-------------------------------------------------------------------------------------------------------------------------------
                          '| Cancelaciones que se quedaron en espera
                          '-------------------------------------------------------------------------------------------------------------------------------
      '                    Case 105, 106
      '                        '| 105 – El CFDI no se puede Cancelar por que tiene Estatus De “En espera De Aceptación”
      '                        '| 106 – El CFDI no se puede Cancelar por que tiene Estatus de “En Proceso”.
      '                        fblnCancelarCFDiProdigia = False
      '                        '| Elimina el registro en la tabla de comprobantes pendientes de cancelar con el estado PA "Pendiente de autorización"
      '                        frsEjecuta_SP strUUId & "|0|PA", "Sp_PvPendientesCancelarSAT" ' Se borra
      '                        '| Inserta un registro en la tabla de comprobantes pendientes de cancelar con el estado PA "Pendiente de autorización"
      '                        frsEjecuta_SP strUUId & "|1|PA|0|'" & UCase(strDescripcionWS) & "'", "Sp_PvPendientesCancelarSAT" ' Se agrega
      '                        strMensaje = "El comprobante se encuentra en espera de aceptación." & vbNewLine & _
      '                                     "Folio en espera: " & strUUId & vbNewLine
      '                        vlintTipoMensajeErrorCancelacionCFDi = vbExclamation
                          '-------------------------------------------------------------------------------------------------------------------------------
                          '| Cancelaciones rechazadas por el receptor
                          '-------------------------------------------------------------------------------------------------------------------------------
      '                    Case 97
      '                        '| 104|97 – El CFDI no se puede Cancelar por que fue Rechazado Previamente.
      '                        fblnCancelarCFDiProdigia = False
      '                        '| Elimina el registro en la tabla de comprobantes pendientes de cancelar con el estado PA "Pendiente de autorización"
      '                        frsEjecuta_SP strUUId & "|0|PA", "Sp_PvPendientesCancelarSAT" ' Se borra
      '                        '| Inserta un registro en la tabla de comprobantes pendientes de cancelar con el estado CR "Cancelación rechazada"
      '                        frsEjecuta_SP strUUId & "|1|CR|0|'" & UCase(strDescripcionWS) & "'", "Sp_PvPendientesCancelarSAT" ' Se agrega
      '                        strMensaje = "El comprobante no se puede cancelar por que fue rechazado por el receptor." & vbNewLine & _
      '                                     "Folio rechazado: " & strUUId & vbNewLine
      '                        vlintTipoMensajeErrorCancelacionCFDi = vbCritical
                          '-------------------------------------------------------------------------------------------------------------------------------
                          '|  Código de error genérico para las excepciones no controladas del SAT
                          '-------------------------------------------------------------------------------------------------------------------------------
112                       Case 99
113                           strMensaje = "El servicio del SAT no responde a la petición de cancelación. " & vbNewLine & _
                                           "Por favor intente más tarde."
114                           If blnBitCancelaCDFiNOSAT Then
115                               fblnCancelarCFDiProdigia = True
116                               frsEjecuta_SP strUUId & "|1|PC", "Sp_PvPendientesCancelarSAT" ' Se agrega
117                           Else
118                               fblnCancelarCFDiProdigia = False
119                               vlintTipoMensajeErrorCancelacionCFDi = vbCritical
120                           End If
                              
121                           frsEjecuta_SP strUUId & "|1||1|'" & UCase(strDescripcionWS) & " - " & Trim(strMensaje) & "'", "Sp_PvPendientesCancelarSAT"
                          '-------------------------------------------------------------------------------------------------------------------------------
                          '| Cancelaciones fallidas
                          '-------------------------------------------------------------------------------------------------------------------------------
122                       Case 0, 1
123                           fblnCancelarCFDiProdigia = False
                              
                              
124                           strMensaje = "Ocurrió un error al solicitar la cancelación del comprobante." & vbNewLine & vbNewLine & _
                                           "Número de error: " & DOMCodigoQuery & vbNewLine & _
                                           "Descripción: " & Trim(DOMQueryXML.selectSingleNode("servicioConsultaComprobante/estado").Text)
125                           vlintTipoMensajeErrorCancelacionCFDi = vbCritical
                              
126                           frsEjecuta_SP strUUId & "|1||1|'" & UCase(strDescripcionWS) & "'", "Sp_PvPendientesCancelarSAT"
127                       Case Else
128                           fblnCancelarCFDiProdigia = False
                              
                              
129                           strMensaje = "Ocurrió un error al solicitar la cancelación del comprobante." & vbNewLine & vbNewLine & _
                                           "Número de error: " & DOMCodigoQuery & vbNewLine & _
                                           "Descripción: " & Trim(DOMQueryXML.selectSingleNode("servicioConsultaComprobante/estado").Text) & " - " & Trim(DOMQueryXML.selectSingleNode("servicioConsultaComprobante/esCancelable").Text)
130                           vlintTipoMensajeErrorCancelacionCFDi = vbCritical
                              
131                           frsEjecuta_SP strUUId & "|1||1|'" & UCase(strDescripcionWS) & "'", "Sp_PvPendientesCancelarSAT"
132                   End Select
133                   vlstrMensajeErrorCancelacionCFDi = strMensaje
                  
134           Else
135               If blnBitCancelaCDFiNOSAT Then
136                  fblnCancelarCFDiProdigia = True
137                  frsEjecuta_SP strUUId & "|1|PC|0|'Pendiente de cancelar ante el SAT'", "Sp_PvPendientesCancelarSAT" ' Se agrega
138               Else
139                  vlstrMensajeErrorCancelacionCFDi = "Ocurrió un error al solicitar la cancelación del comprobante." & vbNewLine & vbNewLine & _
                                                        "No se recibió respuesta del web service."
140                  fblnCancelarCFDiProdigia = False
                     
141                  frsEjecuta_SP strUUId & "|1||1|'" & vlstrMensajeErrorCancelacionCFDi & "'", "Sp_PvPendientesCancelarSAT"
142               End If
143           End If
144       End If
          
145   Exit Function
NotificaErrorTimbre:
       If Err.Number > 0 And Err.Number <> 1001 Then
          If blnBitCancelaCDFiNOSAT Then
             fblnCancelarCFDiProdigia = True
             frsEjecuta_SP strUUId & "|1|PC|0|'NotificaErrorTimbre: " & Err.Number & "'", "Sp_PvPendientesCancelarSAT" ' Se agrega
          Else
             strMensaje = "Ocurrió un error al solicitar la cancelación del comprobante." & vbNewLine & _
                              "Número de error: " & Err.Number & vbNewLine & _
                              "Origen: " & Err.Source & " fblnCancelarCFDiProdigia, Linea:" & Erl() & vbNewLine & _
                              "Descripción: " & Err.Description
             vlstrMensajeErrorCancelacionCFDi = strMensaje
             fblnCancelarCFDiProdigia = False
                
             frsEjecuta_SP strUUId & "|1||1|'" & strMensaje & "'", "Sp_PvPendientesCancelarSAT"
          End If
       End If
End Function

'- Función para formar el Request de la cancelación de un Comprobante Fiscal para Prodigia -'
Private Function fnDOMReqCancelXMLProdigia_NE(strUUId As String, strRFCEmisor As String, strRFCReceptor As String, strTotal As String, strContrato As String, strUsuario As String, strPassword As String, strURLWMCancelacion As String, strURLXMLNSTimbrado As String, strPruebaOReal As String, blnEnsobretadoSOAP As Boolean, strMotivo As String, strFolioFiscalSustituye As String, strURLWMConsulta As String) As MSXML2.DOMDocument
 Dim DOMRequestCancelacionXML As MSXML2.DOMDocument
    Set DOMRequestCancelacionXML = New MSXML2.DOMDocument

    '-------------------------------------------------------------------------------------------
    '---   xmlns   -----------------------------------------------------------------------------
    '-------------------------------------------------------------------------------------------
    Dim DOMelementoSOAP As MSXML2.IXMLDOMElement
    Set DOMelementoSOAP = DOMRequestCancelacionXML.createElement("soap:Envelope")
    
    DOMelementoSOAP.setAttribute "xmlns:soap", "http://schemas.xmlsoap.org/soap/envelope/"
    DOMelementoSOAP.setAttribute "xmlns:tim", strURLXMLNSTimbrado
    
    DOMRequestCancelacionXML.appendChild DOMelementoSOAP
    
    Dim DOMelementoSOAPBody As MSXML2.IXMLDOMElement
    Set DOMelementoSOAPBody = DOMRequestCancelacionXML.createElement("soap:Body")
        
    DOMelementoSOAP.appendChild DOMelementoSOAPBody
    
    Dim DOMelementoRaiz As MSXML2.IXMLDOMElement
    If strPruebaOReal = "1" Then
        Set DOMelementoRaiz = DOMRequestCancelacionXML.createElement("tim:cancelarConOpciones")
    Else
        Set DOMelementoRaiz = DOMRequestCancelacionXML.createElement("tim:" & strURLWMCancelacion)
    End If
    
    DOMelementoSOAPBody.appendChild DOMelementoRaiz
    '-------------------------------------------
    '--- CONTRATO ------------------------------
    '-------------------------------------------
    Dim DOMelementoContrato As MSXML2.IXMLDOMElement
    Set DOMelementoContrato = DOMRequestCancelacionXML.createElement("contrato")

    DOMelementoContrato.Text = strContrato
    DOMelementoRaiz.appendChild DOMelementoContrato
    '-------------------------------------------
    '--- USUARIO -------------------------------
    '-------------------------------------------
    Dim DOMelementoUsuario As MSXML2.IXMLDOMElement
    Set DOMelementoUsuario = DOMRequestCancelacionXML.createElement("usuario")

    DOMelementoUsuario.Text = strUsuario
    DOMelementoRaiz.appendChild DOMelementoUsuario
    '-------------------------------------------
    '--- CONTRASEÑA ----------------------------
    '-------------------------------------------
    Dim DOMelementoContra As MSXML2.IXMLDOMElement
    Set DOMelementoContra = DOMRequestCancelacionXML.createElement("passwd")
    
    DOMelementoContra.Text = strPassword
    DOMelementoRaiz.appendChild DOMelementoContra
    '-------------------------------------------------------------------------------------------
    '---   RFC Emisor   ------------------------------------------------------------------------
    '-------------------------------------------------------------------------------------------
    Dim DOMElementoRFCEmisor As MSXML2.IXMLDOMElement
    Set DOMElementoRFCEmisor = DOMRequestCancelacionXML.createElement("rfcEmisor")
    
    DOMElementoRFCEmisor.Text = Trim(strRFCEmisor)
    DOMelementoRaiz.appendChild DOMElementoRFCEmisor
    '-------------------------------------------------------------------------------------------
    '---   ListUUID   --------------------------------------------------------------------------
    '-------------------------------------------------------------------------------------------
    '|  Crea la lista de UUID
    Dim DOMElementoListUUID As MSXML2.IXMLDOMElement
    Set DOMElementoListUUID = DOMRequestCancelacionXML.createElement("arregloUUID")

    If Trim(strMotivo) = "01" And Trim(strFolioFiscalSustituye) <> "" Then
        DOMElementoListUUID.Text = strUUId & "|" & Trim(strRFCReceptor) & "|" & Trim(strRFCEmisor) & "|" & Trim(strTotal) & "|" & Trim(strMotivo) & "|" & Trim(strFolioFiscalSustituye)
    Else
        DOMElementoListUUID.Text = strUUId & "|" & Trim(strRFCReceptor) & "|" & Trim(strRFCEmisor) & "|" & Trim(strTotal) & "|" & Trim(strMotivo)
    End If
    
    DOMelementoRaiz.appendChild DOMElementoListUUID
    '-------------------------------------------------------------------------------------------
    '---   CERT   --------------------------------------------------------------------------
    '-------------------------------------------------------------------------------------------
    '|  Certificado
    Dim DOMElementoCert As MSXML2.IXMLDOMElement
    Set DOMElementoCert = DOMRequestCancelacionXML.createElement("cert")

    DOMElementoCert.Text = vgStrCertBase64
    DOMelementoRaiz.appendChild DOMElementoCert
    '-------------------------------------------------------------------------------------------
    '---   Key64   --------------------------------------------------------------------------
    '-------------------------------------------------------------------------------------------
    '|  Certificado
    Dim DOMElementoKey64 As MSXML2.IXMLDOMElement
    Set DOMElementoKey64 = DOMRequestCancelacionXML.createElement("key")

    DOMElementoKey64.Text = vgStrKeyBase64
    DOMelementoRaiz.appendChild DOMElementoKey64
    '-------------------------------------------------------------------------------------------
    '---   Pass   --------------------------------------------------------------------------
    '-------------------------------------------------------------------------------------------
    '|  Certificado
    Dim DOMElementoPass As MSXML2.IXMLDOMElement
    Set DOMElementoPass = DOMRequestCancelacionXML.createElement("keyPass")

    DOMElementoPass.Text = vgStrKeyToKey
    DOMelementoRaiz.appendChild DOMElementoPass
    
    'Opciones
    
    If strPruebaOReal = "1" Then
        Dim DOMelementoOpcion As MSXML2.IXMLDOMElement
        Set DOMelementoOpcion = DOMRequestCancelacionXML.createElement("opciones")
        
        DOMelementoOpcion.Text = "MODO_PRUEBA:3"
        DOMelementoRaiz.appendChild DOMelementoOpcion
    End If

    
    vgStrSolicitudEnPlanoCancProdigia = "<contrato>" & strContrato & "</contrato>"
    vgStrSolicitudEnPlanoCancProdigia = vgStrSolicitudEnPlanoCancProdigia & "<usuario>" & strUsuario & "</usuario>"
    vgStrSolicitudEnPlanoCancProdigia = vgStrSolicitudEnPlanoCancProdigia & "<passwd>" & strPassword & "</passwd>"
    vgStrSolicitudEnPlanoCancProdigia = vgStrSolicitudEnPlanoCancProdigia & "<uuid>" & strUUId & "</uuid>"
    vgStrSolicitudEnPlanoCancProdigia = vgStrSolicitudEnPlanoCancProdigia & "<rfcEmisor>" & Replace(Trim(strRFCEmisor), "&", "&#38;") & "</rfcEmisor>"
    vgStrSolicitudEnPlanoCancProdigia = vgStrSolicitudEnPlanoCancProdigia & "<rfcReceptor>" & Replace(Trim(strRFCReceptor), "&", "&#38;") & "</rfcReceptor>"
    vgStrSolicitudEnPlanoCancProdigia = vgStrSolicitudEnPlanoCancProdigia & "<total>" & strTotal & "</total>"
    vgStrSolicitudEnPlanoCancProdigia = "<tim:" & strURLWMConsulta & " xmlns:tim=""" & strURLXMLNSTimbrado & """>" & vgStrSolicitudEnPlanoCancProdigia & "</tim:" & strURLWMConsulta & ">"
    vgStrSolicitudEnPlanoCancProdigia = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/""  xmlns:tim=""" & strURLXMLNSTimbrado & """><soapenv:Body>" & vgStrSolicitudEnPlanoCancProdigia & "</soapenv:Body></soapenv:Envelope>"
    
    '|  Se regresa el valor de la función
    Set fnDOMReqCancelXMLProdigia_NE = DOMRequestCancelacionXML
    
        'Se graba el archivo Request Timbrado en la ruta especificada
End Function

Private Function fnDOMRequestXMLBuzonFiscal(DOMFacturaXMLsinTimbrar As MSXML2.DOMDocument) As MSXML2.DOMDocument
    Dim DOMRequestTimbradoXML As MSXML2.DOMDocument
    Set DOMRequestTimbradoXML = New MSXML2.DOMDocument
    
    Dim strNombreArchivoRequest As String
    strNombreArchivoRequest = Trim(DOMFacturaXMLsinTimbrar.selectSingleNode("cfdi:Comprobante/@serie").Text) + Trim(DOMFacturaXMLsinTimbrar.selectSingleNode("cfdi:Comprobante/@folio").Text) + ".xml"
    
    Dim strComprobanteXML64 As String
    'Se codifican los carácteres a UTF-8 (ANSI 8-bytes) y se codifica el XML para agregarlo en el atributo "DOMElementoDocumento/Archivo"
    strComprobanteXML64 = CStr(DOMFacturaXMLsinTimbrar.xml)   'Se convierte a STRING el contenido del XML
    strComprobanteXML64 = fnstrANSI2UTF8(strComprobanteXML64) 'Codifica de ANSI 7-bytes a 8-bytes
    strComprobanteXML64 = Encode(strComprobanteXML64)   'Se codifica a Base64
    
    
    
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim DOMElementoReq As MSXML2.IXMLDOMElement
    Set DOMElementoReq = DOMRequestTimbradoXML.createElement("tim:RequestTimbradoCFD")
    
    ' Aquí van los Namespaces que contendrá la llamada al Web Service
    DOMElementoReq.setAttribute "xmlns:tim", "http://www.buzonfiscal.com/ns/xsd/bf/TimbradoCFD"
    DOMElementoReq.setAttribute "xmlns:req", "http://www.buzonfiscal.com/ns/xsd/bf/RequestTimbraCFDI"
'    DOMElementoReq.setAttribute "xmlns:cfdi", "http://www.sat.gob.mx/cfd/3"
    DOMElementoReq.setAttribute "req:RefID", vlstrRefID 'Se agregó el RefID para no "quemar" timbres
        
    DOMRequestTimbradoXML.appendChild DOMElementoReq

    pIdentarNodoXML DOMRequestTimbradoXML, DOMElementoReq, 1
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' Solo usamos el nodo DOMElementoInfoBasica, pero se pueden usar tambien los nodos "DOMElementoDocumento" e "InfoAdicional"
    Dim DOMElementoInfoBasica As MSXML2.IXMLDOMElement
    Set DOMElementoInfoBasica = DOMRequestTimbradoXML.createElement("req:InfoBasica")
        
    DOMElementoInfoBasica.setAttribute "RfcEmisor", DOMFacturaXMLsinTimbrar.selectSingleNode("cfdi:Comprobante/cfdi:Emisor/@rfc").Text
    DOMElementoInfoBasica.setAttribute "RfcReceptor", DOMFacturaXMLsinTimbrar.selectSingleNode("cfdi:Comprobante/cfdi:Receptor/@rfc").Text
    DOMElementoReq.appendChild DOMElementoInfoBasica
    
    pIdentarNodoXML DOMRequestTimbradoXML, DOMElementoReq, 1
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim DOMElementoDocumento As MSXML2.IXMLDOMElement
    Set DOMElementoDocumento = DOMRequestTimbradoXML.createElement("req:Documento")
                
    DOMElementoDocumento.setAttribute "Archivo", strComprobanteXML64
    DOMElementoDocumento.setAttribute "NombreArchivo", strNombreArchivoRequest
    DOMElementoDocumento.setAttribute "Tipo", "XML"
    DOMElementoDocumento.setAttribute "Version", "3.2"
    DOMElementoReq.appendChild DOMElementoDocumento
    
    pIdentarNodoXML DOMRequestTimbradoXML, DOMElementoReq, 1
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
      
    'Se regresa el valor de la función
    Set fnDOMRequestXMLBuzonFiscal = DOMRequestTimbradoXML

    'Se graba el archivo Request Timbrado en la ruta especificada
    fnDOMRequestXMLBuzonFiscal.Save vlstrRutaRequestTimbrado
End Function

'Private Function fnDOMRequestXMLPAX(DOMFacturaXMLsinTimbrar As MSXML2.DOMDocument, intEstructura As Integer, strTipo As String, strUsuario As String, strPassword As String, strVersion As String, blnWSPruebas As Boolean, blnEnsobretadoSOAP As Boolean) As MSXML2.DOMDocument
Private Function fnDOMRequestXMLPAX(DOMFacturaXMLsinTimbrar As MSXML2.DOMDocument, intEstructura As Integer, strTipo As String, strUsuario As String, strPassword As String, strVersion As String, strWSRuta As String, blnEnsobretadoSOAP As Boolean) As MSXML2.DOMDocument
    Dim DOMRequestTimbradoXML As MSXML2.DOMDocument
    Set DOMRequestTimbradoXML = New MSXML2.DOMDocument
        
    Dim strNombreArchivoRequest As String
    If vgstrVersionCFDI = "3.2" Then
        strNombreArchivoRequest = Trim(DOMFacturaXMLsinTimbrar.selectSingleNode("cfdi:Comprobante/@serie").Text) + Trim(DOMFacturaXMLsinTimbrar.selectSingleNode("cfdi:Comprobante/@folio").Text) + ".xml"
    Else
        strNombreArchivoRequest = Trim(DOMFacturaXMLsinTimbrar.selectSingleNode("cfdi:Comprobante/@Serie").Text) + Trim(DOMFacturaXMLsinTimbrar.selectSingleNode("cfdi:Comprobante/@Folio").Text) + ".xml"
    End If
    Dim strComprobanteXML As String
    'Se codifican los carácteres a UTF-8 (ANSI 8-bytes) y se codifica el XML para agregarlo en el atributo "Documento/Archivo"
    strComprobanteXML = CStr(DOMFacturaXMLsinTimbrar.xml)   'Se convierte a STRING el contenido del XML
'    strComprobanteXML = fnstrANSI2UTF8(strComprobanteXML) 'Codifica de ANSI 7-bytes a 8-bytes
    strComprobanteXML = Replace(strComprobanteXML, "<?xml version=""1.0""?>", "") 'Se elimina el encabezado
    strComprobanteXML = Replace(Replace(Trim(strComprobanteXML), Chr(10), ""), Chr(13), "") 'Se eliminan los saltos de linea
'    strComprobanteXML = Encode(strComprobanteXML)   'Se codifica a Base64
'    strComprobanteXML = Replace(strComprobanteXML, "&", "&amp;")
'    strComprobanteXML = Replace(strComprobanteXML, "<", "&lt;")
'    strComprobanteXML = Replace(strComprobanteXML, ">", "&gt;")
    
    If blnEnsobretadoSOAP = True Then
        '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Dim DOMelementoSOAP As MSXML2.IXMLDOMElement
        Set DOMelementoSOAP = DOMRequestTimbradoXML.createElement("soap:Envelope")
    
        ' Aquí van los Namespaces
        DOMelementoSOAP.setAttribute "xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance"
        DOMelementoSOAP.setAttribute "xmlns:xsd", "http://www.w3.org/2001/XMLSchema"
        DOMelementoSOAP.setAttribute "xmlns:soap", "http://schemas.xmlsoap.org/soap/envelope/"
        
        DOMRequestTimbradoXML.appendChild DOMelementoSOAP
        '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Dim DOMelementoSOAPBody As MSXML2.IXMLDOMElement
        Set DOMelementoSOAPBody = DOMRequestTimbradoXML.createElement("soap:Body")
        
        DOMelementoSOAP.appendChild DOMelementoSOAPBody
        '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    End If
    
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim DOMelementoRaiz As MSXML2.IXMLDOMElement
    Set DOMelementoRaiz = DOMRequestTimbradoXML.createElement("fnEnviarXML")

    'Esta ruta cambia según el WS de pruebas o el de producción.....después de mil pruebas y a falta de documentación lo descubrí a la mala :/
    'DOMelementoRaiz.setAttribute "xmlns", IIf(blnWSPruebas = True, "https://test.paxfacturacion.com.mx:453", "https://www.paxfacturacion.com.mx:453")
    DOMelementoRaiz.setAttribute "xmlns", strWSRuta
    
    If blnEnsobretadoSOAP = True Then
        DOMelementoSOAPBody.appendChild DOMelementoRaiz
    Else
        DOMRequestTimbradoXML.appendChild DOMelementoRaiz
    End If
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim DOMelementoComprobante As MSXML2.IXMLDOMElement
    Set DOMelementoComprobante = DOMRequestTimbradoXML.createElement("psComprobante")
    
    DOMelementoComprobante.Text = Trim(strComprobanteXML)
    DOMelementoRaiz.appendChild DOMelementoComprobante
    
    DOMelementoRaiz.appendChild DOMelementoComprobante
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim DOMelementoTipoDocumento As MSXML2.IXMLDOMElement
    Set DOMelementoTipoDocumento = DOMRequestTimbradoXML.createElement("psTipoDocumento")
        
    DOMelementoTipoDocumento.Text = Trim(strTipo)
    DOMelementoRaiz.appendChild DOMelementoTipoDocumento
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim DOMElementoIdEstructura As MSXML2.IXMLDOMElement
    Set DOMElementoIdEstructura = DOMRequestTimbradoXML.createElement("pnId_Estructura")
        
    DOMElementoIdEstructura.Text = Trim(intEstructura)
    DOMelementoRaiz.appendChild DOMElementoIdEstructura
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim DOMelementoUsuario As MSXML2.IXMLDOMElement
    Set DOMelementoUsuario = DOMRequestTimbradoXML.createElement("sNombre")
        
    DOMelementoUsuario.Text = Trim(strUsuario)
    DOMelementoRaiz.appendChild DOMelementoUsuario
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim DOMElementoPassword As MSXML2.IXMLDOMElement
    Set DOMElementoPassword = DOMRequestTimbradoXML.createElement("sContraseña")
        
    DOMElementoPassword.Text = Trim(strPassword)
    DOMelementoRaiz.appendChild DOMElementoPassword
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim DOMelementoVersion As MSXML2.IXMLDOMElement
    Set DOMelementoVersion = DOMRequestTimbradoXML.createElement("sVersion")
        
    DOMelementoVersion.Text = Trim(strVersion)
    DOMelementoRaiz.appendChild DOMelementoVersion
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    'Se regresa el valor de la función
    Set fnDOMRequestXMLPAX = DOMRequestTimbradoXML

    'Se graba el archivo Request Timbrado en la ruta especificada
    fnDOMRequestXMLPAX.Save vlstrRutaRequestTimbrado
End Function

'- Función para formar el Request de la cancelación de un Comprobante Fiscal para PAX -'
Public Function fnDOMReqCancelXMLPAX(strUUId As String, strRFC As String, intEstructura As Integer, strUsuario As String, strPassword As String, strWSRuta As String, blnEnsobretadoSOAP As Boolean) As MSXML2.DOMDocument
    Dim DOMRequestCancelacionXML As MSXML2.DOMDocument
    Set DOMRequestCancelacionXML = New MSXML2.DOMDocument
        
    If blnEnsobretadoSOAP = True Then
        '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Dim DOMelementoSOAP As MSXML2.IXMLDOMElement
        Set DOMelementoSOAP = DOMRequestCancelacionXML.createElement("soap:Envelope")
    
        ' Aquí van los Namespaces
        DOMelementoSOAP.setAttribute "xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance"
        DOMelementoSOAP.setAttribute "xmlns:xsd", "http://www.w3.org/2001/XMLSchema"
        DOMelementoSOAP.setAttribute "xmlns:soap", "http://schemas.xmlsoap.org/soap/envelope/"
        
        DOMRequestCancelacionXML.appendChild DOMelementoSOAP
        '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Dim DOMelementoSOAPBody As MSXML2.IXMLDOMElement
        Set DOMelementoSOAPBody = DOMRequestCancelacionXML.createElement("soap:Body")
        
        DOMelementoSOAP.appendChild DOMelementoSOAPBody
        '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    End If
    
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim DOMelementoRaiz As MSXML2.IXMLDOMElement
    Set DOMelementoRaiz = DOMRequestCancelacionXML.createElement("fnCancelarXML")
    DOMelementoRaiz.setAttribute "xmlns", strWSRuta
    
    If blnEnsobretadoSOAP = True Then
        DOMelementoSOAPBody.appendChild DOMelementoRaiz
    Else
        DOMRequestCancelacionXML.appendChild DOMelementoRaiz
    End If
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim DOMElementoListUUID As MSXML2.IXMLDOMElement
    Set DOMElementoListUUID = DOMRequestCancelacionXML.createElement("sListaUUID")
    
    Dim DOMNodoUUID As MSXML2.IXMLDOMElement
    Set DOMNodoUUID = DOMRequestCancelacionXML.createElement("string")
    DOMNodoUUID.Text = strUUId
    
    DOMElementoListUUID.appendChild DOMNodoUUID
    DOMelementoRaiz.appendChild DOMElementoListUUID
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim DOMElementoRFC As MSXML2.IXMLDOMElement
    Set DOMElementoRFC = DOMRequestCancelacionXML.createElement("psRFC")
        
    DOMElementoRFC.Text = Trim(strRFC)
    DOMelementoRaiz.appendChild DOMElementoRFC
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim DOMElementoIdEstructura As MSXML2.IXMLDOMElement
    Set DOMElementoIdEstructura = DOMRequestCancelacionXML.createElement("pn_IdEstructura")
        
    DOMElementoIdEstructura.Text = Trim(intEstructura)
    DOMelementoRaiz.appendChild DOMElementoIdEstructura
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim DOMelementoUsuario As MSXML2.IXMLDOMElement
    Set DOMelementoUsuario = DOMRequestCancelacionXML.createElement("sNombre")
        
    DOMelementoUsuario.Text = Trim(strUsuario)
    DOMelementoRaiz.appendChild DOMelementoUsuario
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim DOMElementoPassword As MSXML2.IXMLDOMElement
    Set DOMElementoPassword = DOMRequestCancelacionXML.createElement("sContraseña")
        
    DOMElementoPassword.Text = Trim(strPassword)
    DOMelementoRaiz.appendChild DOMElementoPassword
     '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    'Se regresa el valor de la función
    Set fnDOMReqCancelXMLPAX = DOMRequestCancelacionXML
End Function


'- Función para formar el Request de la cancelación de un Comprobante Fiscal para PAX -'
Private Function fnDOMReqCancelXMLPAX_NE(strUUId As String, strRFCEmisor As String, strRFCReceptor As String, strTotal As String, strUsuario As String, strPassword As String, strWSRuta As String, blnEnsobretadoSOAP As Boolean, strMotivo As String, strFolioFiscalSustituyePA As String) As MSXML2.DOMDocument
    Dim DOMRequestCancelacionXML As MSXML2.DOMDocument
    Set DOMRequestCancelacionXML = New MSXML2.DOMDocument
        
    If blnEnsobretadoSOAP = True Then
        '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Dim DOMelementoSOAP As MSXML2.IXMLDOMElement
        Set DOMelementoSOAP = DOMRequestCancelacionXML.createElement("soap:Envelope")
    
        ' Aquí van los Namespaces
        DOMelementoSOAP.setAttribute "xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance"
        DOMelementoSOAP.setAttribute "xmlns:xsd", "http://www.w3.org/2001/XMLSchema"
        DOMelementoSOAP.setAttribute "xmlns:soap", "http://www.w3.org/2003/05/soap-envelope/"
        
        DOMRequestCancelacionXML.appendChild DOMelementoSOAP
        '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Dim DOMelementoSOAPBody As MSXML2.IXMLDOMElement
        Set DOMelementoSOAPBody = DOMRequestCancelacionXML.createElement("soap:Body")
        
        DOMelementoSOAP.appendChild DOMelementoSOAPBody
        '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    End If
    
    '-------------------------------------------------------------------------------------------
    '---   xmlns   -----------------------------------------------------------------------------
    '-------------------------------------------------------------------------------------------
    Dim DOMelementoRaiz As MSXML2.IXMLDOMElement
    Set DOMelementoRaiz = DOMRequestCancelacionXML.createElement("fnCancelarXML20")
    'DOMelementoRaiz.setAttribute "xmlns", strWSRuta
    
    If blnEnsobretadoSOAP = True Then
        DOMelementoSOAPBody.appendChild DOMelementoRaiz
    Else
        DOMRequestCancelacionXML.appendChild DOMelementoRaiz
    End If
    
    '-------------------------------------------------------------------------------------------
    '---   ListUUID   --------------------------------------------------------------------------
    '-------------------------------------------------------------------------------------------
    '|  Crea la lista de UUID
    Dim DOMElementoListUUID As MSXML2.IXMLDOMElement
    Set DOMElementoListUUID = DOMRequestCancelacionXML.createElement("sListaUUID")
    '|  Crea el nodo UUID
    Dim DOMNodoUUID As MSXML2.IXMLDOMElement
    Set DOMNodoUUID = DOMRequestCancelacionXML.createElement("string")
    DOMNodoUUID.Text = strUUId
    
    '| Agrega el nodo UUID
    DOMElementoListUUID.appendChild DOMNodoUUID
    '| Agrega la lista elemento raíz
    DOMelementoRaiz.appendChild DOMElementoListUUID
    
    '-------------------------------------------------------------------------------------------
    '---   RFC Emisor   ------------------------------------------------------------------------
    '-------------------------------------------------------------------------------------------
    Dim DOMElementoRFCEmisor As MSXML2.IXMLDOMElement
    Set DOMElementoRFCEmisor = DOMRequestCancelacionXML.createElement("psRFCEmisor")
    DOMElementoRFCEmisor.Text = Trim(strRFCEmisor)
    DOMelementoRaiz.appendChild DOMElementoRFCEmisor
    
    '-------------------------------------------------------------------------------------------
    '---   Lista de RFC Receptor   -------------------------------------------------------------
    '-------------------------------------------------------------------------------------------
    '|  Crea la lista de RFCs del receptor
    Dim DOMElementoListRFCReceptor As MSXML2.IXMLDOMElement
    Set DOMElementoListRFCReceptor = DOMRequestCancelacionXML.createElement("psRFCReceptor")
    '|  Crea el nodo RFC del receptor
    Dim DOMNodoRFCReceptor As MSXML2.IXMLDOMElement
    Set DOMNodoRFCReceptor = DOMRequestCancelacionXML.createElement("string")
    DOMNodoRFCReceptor.Text = Trim(strRFCReceptor)
    
    '| Agrega el nodo a la lista
    DOMElementoListRFCReceptor.appendChild DOMNodoRFCReceptor
    '| Agrega la lista elemento raíz
    DOMelementoRaiz.appendChild DOMElementoListRFCReceptor
    
    '-------------------------------------------------------------------------------------------
    '---   Totales   ---------------------------------------------------------------------------
    '-------------------------------------------------------------------------------------------
    '|  Crea la lista de los totales
    Dim DOMElementoListaTotales As MSXML2.IXMLDOMElement
    Set DOMElementoListaTotales = DOMRequestCancelacionXML.createElement("sListaTotales")
    '|  Crea el nodo RFC del receptor
    Dim DOMNodoTotales As MSXML2.IXMLDOMElement
    Set DOMNodoTotales = DOMRequestCancelacionXML.createElement("string")
    DOMNodoTotales.Text = Trim(strTotal)
        
    '| Agrega el nodo a la lista
    DOMElementoListaTotales.appendChild DOMNodoTotales
    '| Agrega la lista elemento raíz
    DOMelementoRaiz.appendChild DOMElementoListaTotales
    
    
'    -------------------------------------------------------------------------------------------
'    ---   MOTIVOS DE CANCELACIÓN   ------------------------------------------------------------
'    -------------------------------------------------------------------------------------------
    Dim DOMElementoListaMotivos As MSXML2.IXMLDOMElement
    Set DOMElementoListaMotivos = DOMRequestCancelacionXML.createElement("sMotivosCancelacion")
    Dim DOMNodoMotivos As MSXML2.IXMLDOMElement
    Set DOMNodoMotivos = DOMRequestCancelacionXML.createElement("string")
    DOMNodoMotivos.Text = Trim(strMotivo)

    '| Agrega el nodo a la lista
    DOMElementoListaMotivos.appendChild DOMNodoMotivos
    '| Agrega la lista elemento raíz
    DOMelementoRaiz.appendChild DOMElementoListaMotivos
    
    '-------------------------------------------------------------------------------------------
    '---   FOLIOS DE SUSTITUCIÓN  --------------------------------------------------------------
    '-------------------------------------------------------------------------------------------
    Dim DOMElementoListaFolios As MSXML2.IXMLDOMElement
    Set DOMElementoListaFolios = DOMRequestCancelacionXML.createElement("sFoliosSustitucion")
    Dim DOMNodoFolios As MSXML2.IXMLDOMElement
    Set DOMNodoFolios = DOMRequestCancelacionXML.createElement("string")
    DOMNodoFolios.Text = Trim(strFolioFiscalSustituyePA)
    
    '| Agrega el nodo a la lista
    DOMElementoListaFolios.appendChild DOMNodoFolios
    '| Agrega la lista elemento raíz
    DOMelementoRaiz.appendChild DOMElementoListaFolios
    
    
    '-------------------------------------------------------------------------------------------
    '---   Nombre (Usuario)   ------------------------------------------------------------------
    '-------------------------------------------------------------------------------------------
    Dim DOMelementoUsuario As MSXML2.IXMLDOMElement
    Set DOMelementoUsuario = DOMRequestCancelacionXML.createElement("sNombre")
    DOMelementoUsuario.Text = Trim(strUsuario)
    DOMelementoRaiz.appendChild DOMelementoUsuario
    
    '-------------------------------------------------------------------------------------------
    '---   Contraseña   ------------------------------------------------------------------------
    '-------------------------------------------------------------------------------------------
    Dim DOMElementoPassword As MSXML2.IXMLDOMElement
    Set DOMElementoPassword = DOMRequestCancelacionXML.createElement("sContrasena")
    DOMElementoPassword.Text = Trim(strPassword)
    DOMelementoRaiz.appendChild DOMElementoPassword

    '|  Se regresa el valor de la función
    Set fnDOMReqCancelXMLPAX_NE = DOMRequestCancelacionXML
    
        'Se graba el archivo Request Timbrado en la ruta especificada
End Function

'- Función para formar el Request de la cancelación de un Comprobante Fiscal para Buzón Fiscal -'
Private Function fnDOMReqCancelXMLBuzon(strUUId As String, strRFC As String, strRFCReceptor As String, blnEnsobretadoSOAP As Boolean) As MSXML2.DOMDocument
    Dim DOMRequestCancelacionXML As MSXML2.DOMDocument
    Set DOMRequestCancelacionXML = New MSXML2.DOMDocument
    
    If blnEnsobretadoSOAP = True Then
        '-------------- Sección del ensobretado del mensaje ----------'
        Dim DOMelementoSOAP As MSXML2.IXMLDOMElement
        Set DOMelementoSOAP = DOMRequestCancelacionXML.createElement("soapenv:Envelope")
    
        '- Namespaces que contendrá la llamada al Web Service -'
        DOMelementoSOAP.setAttribute "xmlns:soapenv", "http://schemas.xmlsoap.org/soap/envelope/"
        DOMelementoSOAP.setAttribute "xmlns:ns", "http://www.buzonfiscal.com/ns/xsd/bf/bfcorp/3"
        
        DOMRequestCancelacionXML.appendChild DOMelementoSOAP
        
        '-------------- Sección del cuerpo del mensaje ----------'
        Dim DOMelementoSOAPBody As MSXML2.IXMLDOMElement
        Set DOMelementoSOAPBody = DOMRequestCancelacionXML.createElement("soapenv:Body")
        DOMelementoSOAP.appendChild DOMelementoSOAPBody
        '--------------------------------------------------------'
    End If
    
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim DOMElementoReq As MSXML2.IXMLDOMElement
    Set DOMElementoReq = DOMRequestCancelacionXML.createElement("ns:RequestCancelaCFDi")
    
    '- Namespaces que contendrá la llamada al Web Service -'
    If blnEnsobretadoSOAP = False Then
        DOMElementoReq.setAttribute "xmlns:ns", "http://www.buzonfiscal.com/ns/xsd/bf/bfcorp/3"
        DOMElementoReq.setAttribute "xmlns:soapenv", "http://schemas.xmlsoap.org/soap/envelope/"
    End If
    
    DOMElementoReq.setAttribute "uuid", strUUId
    DOMElementoReq.setAttribute "rfcReceptor", strRFCReceptor
    DOMElementoReq.setAttribute "rfcEmisor", strRFC
        
    If blnEnsobretadoSOAP = True Then
        DOMelementoSOAPBody.appendChild DOMElementoReq
    Else
        DOMRequestCancelacionXML.appendChild DOMElementoReq
    End If
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
      
    'Se regresa el valor de la función
    Set fnDOMReqCancelXMLBuzon = DOMRequestCancelacionXML
End Function

Public Function fnstrANSI2UTF8(strCadena As String) As String
    Dim strDecodificado As String
    
    strDecodificado = strCadena
    
    'Se deben de reemplazar en este orden para que no afecte la codificación
    strDecodificado = Replace$(strDecodificado, "´", Chr(194) & Chr(180))
    strDecodificado = Replace$(strDecodificado, "°", Chr(194) & Chr(176))
    strDecodificado = Replace$(strDecodificado, "¡", Chr(194) & Chr(161))
    strDecodificado = Replace$(strDecodificado, "Ñ", Chr(195) & Chr(145))
    strDecodificado = Replace$(strDecodificado, "ñ", Chr(195) & Chr(177))
    strDecodificado = Replace$(strDecodificado, "¿", Chr(194) & Chr(191))
    strDecodificado = Replace$(strDecodificado, "Á", Chr(195) & Chr(129))
    strDecodificado = Replace$(strDecodificado, "É", Chr(195) & Chr(137))
    strDecodificado = Replace$(strDecodificado, "Í", Chr(195) & Chr(141))
    strDecodificado = Replace$(strDecodificado, "Ó", Chr(195) & Chr(147))
    strDecodificado = Replace$(strDecodificado, "Ú", Chr(195) & Chr(154))
    strDecodificado = Replace$(strDecodificado, "á", Chr(195) & Chr(161))
    strDecodificado = Replace$(strDecodificado, "é", Chr(195) & Chr(169))
    strDecodificado = Replace$(strDecodificado, "í", Chr(195) & Chr(173))
    strDecodificado = Replace$(strDecodificado, "ó", Chr(195) & Chr(179))
    strDecodificado = Replace$(strDecodificado, "ú", Chr(195) & Chr(186))
    
    fnstrANSI2UTF8 = strDecodificado
End Function

Private Sub pAgregarTimbreBuzonFiscal(DOMFacturaXMLsinTimbre As MSXML2.DOMDocument, DOMNodoTimbre As MSXML2.IXMLDOMNode)
    Dim NodoTimbre As MSXML2.IXMLDOMElement
    Dim NodoComplemento As MSXML2.IXMLDOMNode
    Dim Atributos As MSXML2.IXMLDOMNamedNodeMap
    
    Set Atributos = DOMNodoTimbre.childNodes(0).Attributes
    Set NodoTimbre = DOMFacturaXMLsinTimbre.createElement("tfd:TimbreFiscalDigital")
    NodoTimbre.setAttribute "selloSAT", Atributos.getNamedItem("selloSAT").Text
    NodoTimbre.setAttribute "noCertificadoSAT", Atributos.getNamedItem("noCertificadoSAT").Text
    NodoTimbre.setAttribute "selloCFD", Atributos.getNamedItem("selloCFD").Text
    NodoTimbre.setAttribute "FechaTimbrado", Atributos.getNamedItem("FechaTimbrado").Text
    NodoTimbre.setAttribute "UUID", Atributos.getNamedItem("UUID").Text
    NodoTimbre.setAttribute "version", Atributos.getNamedItem("version").Text
    
    Set NodoComplemento = DOMFacturaXMLsinTimbre.createNode(1, "cfdi:Complemento", "http://www.sat.gob.mx/cfd/3")
    
    DOMFacturaXMLsinTimbre.documentElement.appendChild DOMFacturaXMLsinTimbre.createTextNode(vbTab)
    NodoComplemento.appendChild DOMFacturaXMLsinTimbre.createTextNode(vbNewLine & vbTab & vbTab)
    
    NodoComplemento.appendChild NodoTimbre
    DOMFacturaXMLsinTimbre.documentElement.appendChild NodoComplemento
    
    NodoComplemento.appendChild DOMFacturaXMLsinTimbre.createTextNode(vbNewLine & vbTab)
    DOMFacturaXMLsinTimbre.documentElement.appendChild DOMFacturaXMLsinTimbre.createTextNode(vbNewLine)
    
    Set Atributos = Nothing
    Set NodoComplemento = Nothing
    Set NodoTimbre = Nothing
End Sub

' Formatea el aspecto de los nodos para que sea mas comoda su lectura. (No aplica para todos los XML Request, ya que algúnos WS como PAX, no admiten saltos de linea o espacios!!!
Private Sub pIdentarNodoXML(DOM As MSXML2.DOMDocument, Nodo As MSXML2.IXMLDOMElement, ByVal Nivel As Long)
    Nodo.appendChild DOM.createTextNode(vbNewLine & String(Nivel, vbTab))
End Sub
'- Formatear la fecha del XML -'
Private Function fstrFormatoFechaHora(lstrFecha As String) As String
    Dim lintPosicion As Integer
On Error GoTo NOFECHA
    
    
    lstrFecha = Replace(lstrFecha, "Z", "")
    lintPosicion = InStr(lstrFecha, "T")
    If lintPosicion > 0 Then
        fstrFormatoFechaHora = Format(Left(lstrFecha, lintPosicion - 1), "dd/mmm/yyyy") & " " & Format(Mid(lstrFecha, lintPosicion + 1, Len(lstrFecha)), "hh:mm")
    Else
        'fstrFormatoFechaHora = Format(lstrFecha, "dd/mmm/yyyy hh:mm")
         fstrFormatoFechaHora = strFormatoDate(lstrFecha)
    End If
    
Exit Function
NOFECHA:
  fstrFormatoFechaHora = "NO DISPONIBLE"
End Function

Private Function strFormatoDate(objSTR As String) As String
'11/1/2012 ó '11/13/2013 ó 1/1/2013'
Dim datoI As String 'mes
Dim datoII As String 'dia
Dim datoIII As String 'año
Dim datoIV As String 'hora
Dim objcont As Integer
Dim objBan As Integer
On Error GoTo NOFECHA
    objBan = 1
    datoI = ""
    datoII = ""
    datoIII = ""
    datoIV = ""

    For objcont = 1 To Len(objSTR)
        If Not IsNumeric(Mid(objSTR, objcont, 1)) Then
            
            If datoI = "" Then
            datoI = Mid(objSTR, objBan, objcont - objBan)
            objBan = objcont + 1
            ElseIf datoII = "" Then
            datoII = Mid(objSTR, objBan, objcont - objBan)
            objBan = objcont + 1
            ElseIf datoIII = "" Then
            datoIII = Mid(objSTR, objBan, objcont - objBan)
            objBan = objcont + 1
            Else
            datoIV = Mid(objSTR, objBan, Len(objSTR) - (objBan - 1))
            Exit For
            End If
         End If
    Next objcont
    
    If datoIII = "" Then
    datoIII = Mid(objSTR, objBan, Len(objSTR) - (objBan - 1))
    End If
    
    strFormatoDate = Format(datoII & "/" & datoI & "/" & Trim(datoIII), "dd/mmm/yyyy") & " " & Format(datoIV, "HH:mm")
Exit Function
NOFECHA:
    strFormatoDate = "FECHA NO DISPONIBLE"
End Function


'----- Llamar al Web Service de Buzón Fiscal para cancelar el comprobante fiscal (Agregado CR) -----'
Public Function fblnCancelarCFDiBuzon(strUUId As String, strRFC As String, strRFCReceptor As String, strRutaAcuse As String) As Boolean
1     On Error GoTo NotificaErrorTimbre:

          '-- Variables para el request de la cancelación -''
          Dim SerializerWS As SoapSerializer30     'Para serializar el XML
          Dim ConectorWS As ISoapConnector         'Para conectarse al WebService
          Dim DOMRequestXML As MSXML2.DOMDocument  'XML con los datos para la cancelación
          Dim rsConexion As New ADODB.Recordset    'Para leer datos de configuración del PAC
          Dim strURLWSCancelacion As String        'Dirección del Web Service para la cancelación
          Dim strURLWMCancelacion As String        'Dirección del Web Method para la cancelación
          
          '-- Variables para la respuesta del WebService --'
          Dim ReaderRespuestaWS As SoapReader30    'Para leer la respuesta del WebService
          Dim DOMResponseXML As MSXML2.DOMDocument 'XML para leer los datos de la respuesta del WebService
          Dim strCodigoWS As String                'Almacena el código de respuesta del WebService
          Dim strUUIDWS As String                  'Almacena el código del comprobante regresado por el WebService
          Dim strDescripcionWS As String           'Almacena el mensaje de respuesta del WebService
          Dim strFechaWS As String                 'Almacena la fecha de cancelación del WebService
          
          '-- Variables generales --'
          Dim strSentencia As String
          Dim strMensaje As String
          
          '--Variable para saber si se da la opción de cancelar el CFDi en forma manual, cuando el servicio del SAT no esta disponible
          Dim blnBitCancelaCDFiNOSAT As Boolean
          Dim RsbitCancelaCFDiNOSAT As New ADODB.Recordset
          
          Dim XmlHTTP As MSXML2.XmlHTTP
          Dim strResponse As String
          
2         vlstrMensajeErrorCancelacionCFDi = "" 'Limpiar variable del mensaje de error
          
          ' Se carga parámetro para saber si en caso de un error se puede permitir la cancelación del documento en el SIHO para después cancelar en SAT
3         strSentencia = "SELECT vchvalor FROM SiParametro WHERE VCHNOMBRE = 'BITCANCELACFDINOSAT' AND INTCVEEMPRESACONTABLE = " & vgintClaveEmpresaContable
4         Set RsbitCancelaCFDiNOSAT = frsRegresaRs(strSentencia, adLockOptimistic)
5         If RsbitCancelaCFDiNOSAT.RecordCount > 0 Then
6            blnBitCancelaCDFiNOSAT = IIf(IsNull(RsbitCancelaCFDiNOSAT!vchvalor), 0, Val(RsbitCancelaCFDiNOSAT!vchvalor))
7         Else
8            blnBitCancelaCDFiNOSAT = 0
9         End If
          
          'Se especifican las rutas para la conexión con el servicio de cancelación
10        Set rsConexion = frsEjecuta_SP(CStr(vgintClaveEmpresaContable), "SP_GNSELCONFIGPAC")
11        strURLWSCancelacion = rsConexion!URLWSCancelacion   '(Pruebas = "https://demonegocios.buzonfiscal.com/bfcorpcfdiws", Producción = "https://serviciostf.buzonfiscal.com/bfcorpcfdiws")
12        strURLWMCancelacion = rsConexion!URLWMCancelacion
          
          'Se forma el archivo Request de la Cancelación
13        Set DOMRequestXML = fnDOMReqCancelXMLBuzon(strUUId, strRFC, strRFCReceptor, True)
          
          '#####################################################################################################################################
          '########################################### INICIA CONEXIÓN CON EL SERVICIO DE CANCELACIÓN ##########################################
          '#####################################################################################################################################
      '    Set ConectorWS = New HttpConnector30
      '
      '    ConectorWS.Property("EndPointURL") = strURLWSCancelacion    'Ruta del WebService
      '    ConectorWS.Property("SoapAction") = strURLWMCancelacion     'Ruta del WebMethod para la cancelación ("http://www.buzonfiscal.com/CorporativoWS3.0/cancelaCFDi")
      '
      '
      '    ConectorWS.Connect
      '
      '    ConectorWS.BeginMessage
      '        Set SerializerWS = New SoapSerializer30
      '        SerializerWS.Init ConectorWS.InputStream
      '        'Agrega información del request
      '        SerializerWS.StartEnvelope
      '            SerializerWS.StartBody
      '                SerializerWS.WriteXml DOMRequestXML.xml
      '            SerializerWS.EndBody
      '        SerializerWS.EndEnvelope
      '    ConectorWS.EndMessage
      '
      '    Set ReaderRespuestaWS = New SoapReader30
      '    ReaderRespuestaWS.Load ConectorWS.OutputStream
          '#######################################################################################################################################
          '########################################### FINAL DE CONEXIÓN CON EL SERVICIO DE CANCELACIÓN ##########################################
          '#######################################################################################################################################
            
      '   If Not ReaderRespuestaWS.Fault Is Nothing Then
      '        If Not ReaderRespuestaWS.FaultDetail Is Nothing Then
      '            Dim DOMNodoCodigo As MSXML2.IXMLDOMNode
      '            '------------------------------------------------------------------------------------------------ ERROR A NIVEL 1
      '            'Se obtiene el codigo del error devuelto
      '            Set DOMNodoCodigo = ReaderRespuestaWS.FaultDetail.childNodes(0).Attributes.getNamedItem("codigo")
      '            If DOMNodoCodigo Is Nothing Then
      '                Err.Raise 1000, "Error de cancelación de nivel de capa 1", "Error"
      '                fblnCancelarCFDiBuzon = False
      '            Else
      '                If blnBitCancelaCDFiNOSAT Then
      '                    fblnCancelarCFDiBuzon = True
      '                    frsEjecuta_SP strUUID & "|1", "Sp_PvPendientesCancelarSAT" 'Agregar el UUID como pendiente de cancelar
      '                Else
      '                    strMensaje = "Ocurrió un error al solicitar el servicio de cancelación" & vbNewLine & vbNewLine & _
      '                                 "Número de error: " & DOMNodoCodigo.Text & vbNewLine & _
      '                                 "Descripción: " & ReaderRespuestaWS.FaultString.Text
      '
      '                    vlstrMensajeErrorCancelacionCFDi = strMensaje
      '                    fblnCancelarCFDiBuzon = False
      '                End If
      '            End If
      '        End If

14        Set XmlHTTP = New XmlHTTP
          
          'Abrir la conexión con el Web service
15        XmlHTTP.Open "POST", strURLWSCancelacion, False
          
          'Crear las cabeceras del XML
16        XmlHTTP.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
17        XmlHTTP.setRequestHeader "SOAPAction", strURLWMCancelacion 'Web Method de la cancelación
          
          'Enviar el comando con el XML de request
18        XmlHTTP.send DOMRequestXML.xml
          
          'Almacenar el resultado en la variable strResponse
19        strResponse = XmlHTTP.responseText
          
20        Set XmlHTTP = Nothing 'Liberar el objeto
         
21        If Trim(strResponse) = "" Then
22            If blnBitCancelaCDFiNOSAT Then
23                fblnCancelarCFDiBuzon = True
24                frsEjecuta_SP strUUId & "|1", "Sp_PvPendientesCancelarSAT" ' Se agrega como pendiente de cancelar
25            Else
26                vlstrMensajeErrorCancelacionCFDi = "Ocurrió un error al solicitar la cancelación del comprobante." & vbNewLine & vbNewLine & _
                                                     "No se recibió respuesta del web service."
27                fblnCancelarCFDiBuzon = False
28            End If
29        Else
              Dim strEnv As String
              Dim strBody As String
              
30            strEnv = ""
31            strBody = ""
              
32            If InStr(strResponse, "<soap:Envelope") > 0 Then
33                strEnv = "soap:Envelope/"
34            ElseIf InStr(strResponse, "<S:Envelope") > 0 Then
35                strEnv = "S:Envelope/"
36            End If
              
37            If Trim(strEnv) <> "" Then
38                If InStr(strResponse, "<soap:Body>") > 0 Then
39                    strBody = "soap:Body/"
40                ElseIf InStr(strResponse, "<S:Body>") > 0 Then
41                    strBody = "S:Body/"
42                End If
43            End If
          
44            Set DOMResponseXML = New MSXML2.DOMDocument
45            DOMResponseXML.loadXML strResponse '(ReaderRespuestaWS.Body.Text)
              'Se verifica si el WS regresó información...
              Dim DOMNodoCodigo As MSXML2.IXMLDOMNode
              'Set DOMNodoCodigo = DOMResponseXML.selectSingleNode("ns3:ResponseCancelaCFDi/ns2:Result")
46            Set DOMNodoCodigo = DOMResponseXML.selectSingleNode(strEnv & strBody & "ns3:ResponseCancelaCFDi/ns2:Result")
47            If Not DOMNodoCodigo Is Nothing Then
48                strCodigoWS = Trim(DOMNodoCodigo.selectSingleNode("ns2:Message/@code").Text)
49                strDescripcionWS = Trim(DOMNodoCodigo.selectSingleNode("ns2:Message/@message").Text)
50                If Val(strCodigoWS) = 0 Or Val(strCodigoWS) = 19 Then
                      '0  - Proceso realizado con exito.
                      '19 - El CFD ya fue cancelado, previamente.
                      '-------------------------------------------------------------------------------------------- CANCELACIÓN CORRECTA
51                    fblnCancelarCFDiBuzon = True
                      'Se graba el archivo Response de cancelación en la ruta especificada, sólo cuando se cancela la primera vez
52                    If Val(strCodigoWS) = 0 And strRutaAcuse <> "" Then
53                        DOMResponseXML.Save strRutaAcuse
54                    End If

55                Else
                      '-------------------------------------------------------------------------------------------- ERROR A NIVEL 2
56                    If blnBitCancelaCDFiNOSAT Then
57                        fblnCancelarCFDiBuzon = True
58                        frsEjecuta_SP strUUId & "|1", "Sp_PvPendientesCancelarSAT"
59                    Else
60                        strMensaje = "Ocurrió un error al solicitar la cancelación del comprobante." & vbNewLine & vbNewLine & _
                                       "Número de error: " & strCodigoWS & vbNewLine & _
                                       "Descripción: " & strDescripcionWS
                                         
61                        vlstrMensajeErrorCancelacionCFDi = strMensaje
62                        fblnCancelarCFDiBuzon = False
63                    End If
64                End If
65            Else
66                If blnBitCancelaCDFiNOSAT Then
67                    fblnCancelarCFDiBuzon = True
68                    frsEjecuta_SP strUUId & "|1", "Sp_PvPendientesCancelarSAT"
69                Else
70                    vlstrMensajeErrorCancelacionCFDi = "Ocurrió un error al solicitar la cancelación del comprobante." & vbNewLine & vbNewLine & _
                                                         "No se recibió respuesta del web service."
71                    fblnCancelarCFDiBuzon = False
72                End If
73            End If
74        End If
          
75    Exit Function
NotificaErrorTimbre:
        If Err.Number > 0 And Err.Number <> 1001 Then
            If blnBitCancelaCDFiNOSAT Then
                fblnCancelarCFDiBuzon = True
                frsEjecuta_SP strUUId & "|1", "Sp_PvPendientesCancelarSAT"
            Else
                strMensaje = "Ocurrió un error al solicitar la cancelación del comprobante." & vbNewLine & _
                                "Número de error: " & Err.Number & vbNewLine & _
                                "Origen: " & Err.Source & " fblnCancelarCFDiBuzon, Linea:" & Erl() & vbNewLine & _
                                "Descripción: " & Err.Description
                vlstrMensajeErrorCancelacionCFDi = strMensaje
                fblnCancelarCFDiBuzon = False
            End If
        End If
End Function

'----- Llamar al Web Service de Buzón Fiscal para cancelar el comprobante fiscal en forma masiva (Agregadp CR) -----
Public Function fintCancelarCFDiBuzon(strUUId As String, strRFC As String, strRFCReceptor As String, strRutaAcuse As String, strFolio As String) As Integer
On Error GoTo NotificaErrorTimbre:

    '-- Variables para el request de la cancelación -''
    Dim SerializerWS As SoapSerializer30     'Para serializar el XML
    Dim ConectorWS As ISoapConnector         'Para conectarse al WebService
    Dim DOMRequestXML As MSXML2.DOMDocument  'XML con los datos para la cancelación
    Dim rsConexion As New ADODB.Recordset    'Para leer datos de configuración del PAC
    Dim strURLWSCancelacion As String        'Dirección del Web Service para la cancelación
    Dim strURLWMCancelacion As String        'Dirección del Web Method para la cancelación
    
    '-- Variables para la respuesta del WebService --'
    Dim ReaderRespuestaWS As SoapReader30    'Para leer la respuesta del WebService
    Dim DOMResponseXML As MSXML2.DOMDocument 'XML para leer los datos de la respuesta del WebService
    Dim strCodigoWS As String                'Almacena el código de respuesta del WebService
    Dim strUUIDWS As String                  'Almacena el código del comprobante regresado por el WebService
    Dim strDescripcionWS As String           'Almacena el mensaje de respuesta del WebService
    
    '-- Variables generales --'
    Dim strSentencia As String
    Dim strMensaje As String
    
    '-- Variables para conexión directa con Web Service --'
    Dim XmlHTTP As MSXML2.XmlHTTP
    Dim strResponse As String
    
    '''A ESTA VALIDACION NUNCA DEBE DE ENTRAR EL PROCESO DE LOS CLIENTES, SOLAMENTE ES PARA PRUEBAS YA QUE NO HAY FORMA DE HACER UNA PRUEBA
    ''' PARA LA CANCELACION DE UN CFDI ANTE LA SAT, EN EL CLIENTE NUNCA DEBE DE EXISTIR EL PARAMETRO 'BITPRUEBACANCELACIONCFDI'
    Dim ObjRs As New ADODB.Recordset
    Set ObjRs = frsRegresaRs("Select vchvalor from siparametro where vchnombre = 'BITPRUEBACANCELACIONCFDI'", adLockOptimistic)
    If ObjRs.RecordCount > 0 Then
       If ObjRs!vchvalor = "1" Then
          fintCancelarCFDiBuzon = 3
          Exit Function
       End If
    End If
    '''''''--------------------------------------------------------------------------------------------------------------------------------
    
    Set rsConexion = frsEjecuta_SP(CStr(vgintClaveEmpresaContable), "SP_GNSELCONFIGPAC")
    strURLWSCancelacion = rsConexion!URLWSCancelacion   '(Pruebas = "https://demonegocios.buzonfiscal.com/bfcorpcfdiws", Producción = "https://serviciostf.buzonfiscal.com/bfcorpcfdiws")
    strURLWMCancelacion = rsConexion!URLWMCancelacion
    
    'Se forma el archivo Request de la Cancelación
    Set DOMRequestXML = fnDOMReqCancelXMLBuzon(strUUId, strRFC, strRFCReceptor, True)
   
'    Set ConectorWS = New HttpConnector30
'
'    ConectorWS.Property("EndPointURL") = strURLWSCancelacion    'Ruta del WebService
'    ConectorWS.Property("SoapAction") = strURLWMCancelacion     'Ruta del WebMethod para la cancelación ("https://www.paxfacturacion.com.mx/fnCancelarXML")
    
   
    '#####################################################################################################################################
    '########################################### INICIA CONEXIÓN CON EL SERVICIO DE CANCELACIÓN ##########################################
    '#####################################################################################################################################
'    ConectorWS.Connect
'
'    ConectorWS.BeginMessage
'        Set SerializerWS = New SoapSerializer30
'        SerializerWS.Init ConectorWS.InputStream
'        'Agrega información del request
'        SerializerWS.StartEnvelope
'            SerializerWS.StartBody
'                SerializerWS.WriteXml DOMRequestXML.xml
'            SerializerWS.EndBody
'        SerializerWS.EndEnvelope
'    ConectorWS.EndMessage
'
'    Set ReaderRespuestaWS = New SoapReader30
'    ReaderRespuestaWS.Load ConectorWS.OutputStream
    '#######################################################################################################################################
    '########################################### FINAL DE CONEXIÓN CON EL SERVICIO DE CANCELACIÓN ##########################################
    '#######################################################################################################################################
    
    fintCancelarCFDiBuzon = 0
    
'    If Not ReaderRespuestaWS.Fault Is Nothing Then
'        If Not ReaderRespuestaWS.FaultDetail Is Nothing Then
'            Dim DOMNodoCodigo As MSXML2.IXMLDOMNode
'            '------------------------------------------------------------------------------------------------ ERROR A NIVEL 1
'            'Se obtiene el codigo del error devuelto por PAX
'            Set DOMNodoCodigo = ReaderRespuestaWS.FaultDetail.childNodes(0).Attributes.getNamedItem("codigo")
'            'Se captura el error
'            If DOMNodoCodigo Is Nothing Then
'                Err.Raise 1000, "Error de cancelación de nivel de capa 1", "Error"
'                fIntCancelarCFDiPAX = 0
'            Else
'                strMensaje = "Ocurrió un error al solicitar el servicio de cancelación" & vbNewLine & vbNewLine & _
'                             "Folio del documento: " & strFolio & vbNewLine & _
'                             "Número de error: " & DOMNodoCodigo.Text & vbNewLine & _
'                             "Descripción: " & ReaderRespuestaWS.FaultString.Text & vbNewLine & _
'                             "¿Desea continuar?"
'                'Se muestra el mensaje de error en pantalla
'                If MsgBox(strMensaje, vbCritical + vbYesNo, "Mensaje") = vbYes Then
'                   fintCancelarCFDiBuzon = 0
'                Else
'                   fintCancelarCFDiBuzon = 2
'                End If
'            End If
'        End If
'    Else

    Set XmlHTTP = New XmlHTTP
    
    'Abrir la conexión con el Web service
    XmlHTTP.Open "POST", strURLWSCancelacion, False
    
    'Crear las cabeceras del XML
    XmlHTTP.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    XmlHTTP.setRequestHeader "SOAPAction", strURLWMCancelacion 'Web Method de la cancelación
    
    'Enviar el comando con el XML de request
    XmlHTTP.send DOMRequestXML.xml
    
    'Almacenar el resultado en la variable strResponse
    strResponse = XmlHTTP.responseText
    
    Set XmlHTTP = Nothing 'Liberar el objeto
   
    If Trim(strResponse) = "" Then
        strMensaje = "Ocurrió un error al solicitar el servicio de cancelación" & vbNewLine & vbNewLine & _
                     "Folio del documento: " & strFolio & vbNewLine & _
                     "Descripción: No se recibió respuesta del web service." & vbNewLine & _
                     "¿Desea continuar?"
        'Se muestra el mensaje de error en pantalla
        If MsgBox(strMensaje, vbCritical + vbYesNo, "Mensaje") = vbYes Then
           fintCancelarCFDiBuzon = 0
        Else
           fintCancelarCFDiBuzon = 2
        End If
    Else
        Dim strEnv As String
        Dim strBody As String
        
        strEnv = ""
        strBody = ""
        
        If InStr(strResponse, "<soap:Envelope") > 0 Then
            strEnv = "soap:Envelope/"
        ElseIf InStr(strResponse, "<S:Envelope") > 0 Then
            strEnv = "S:Envelope/"
        End If
        
        If Trim(strEnv) <> "" Then
            If InStr(strResponse, "<soap:Body>") > 0 Then
                strBody = "soap:Body/"
            ElseIf InStr(strResponse, "<S:Body>") > 0 Then
                strBody = "S:Body/"
            End If
        End If
    
        Set DOMResponseXML = New MSXML2.DOMDocument
        DOMResponseXML.loadXML strResponse '(ReaderRespuestaWS.Body.Text)
        'Se verifica si el WS regresó información...
        Dim DOMNodoCodigo As MSXML2.IXMLDOMNode
        Set DOMNodoCodigo = DOMResponseXML.selectSingleNode(strEnv & strBody & "ns3:ResponseCancelaCFDi/ns2:Result")
        If Not DOMNodoCodigo Is Nothing Then
            strCodigoWS = Trim(DOMNodoCodigo.selectSingleNode("ns2:Message/@code").Text)
            strDescripcionWS = Trim(DOMNodoCodigo.selectSingleNode("ns2:Message/@message").Text)
            If Val(strCodigoWS) = 0 Or Val(strCodigoWS) = 19 Then
                '0  - Proceso realizado con exito.
                '19 - El CFD ya fue cancelado, previamente.
                '-------------------------------------------------------------------------------------------- CANCELACIÓN CORRECTA
                fintCancelarCFDiBuzon = 1
                'Se graba el archivo Response de cancelación en la ruta especificada, sólo cuando se cancela la primera vez
                If Val(strCodigoWS) = 0 And strRutaAcuse <> "" Then DOMResponseXML.Save strRutaAcuse
            Else
                '-------------------------------------------------------------------------------------------- ERROR A NIVEL 2
                strMensaje = "Ocurrió un error al solicitar la cancelación del comprobante." & vbNewLine & vbNewLine & _
                             "Folio del documento: " & strFolio & vbNewLine & _
                             "Número de error: " & strCodigoWS & vbNewLine & _
                             "Descripción: " & strDescripcionWS & vbNewLine & _
                             "¿Desea continuar?"
                                
                If MsgBox(strMensaje, vbCritical + vbYesNo, "Mensaje") = vbYes Then
                   fintCancelarCFDiBuzon = 0
                Else
                   fintCancelarCFDiBuzon = 2
                End If
            End If
        Else
            If MsgBox("Ocurrió un error al solicitar la cancelación del comprobante." & vbNewLine & vbNewLine & _
                      "Folio del documento: " & strFolio & vbNewLine & _
                      "No se recibió respuesta del web service." & vbNewLine & "¿Desea continuar?", vbCritical + vbYesNo, "Mensaje") = vbYes Then
                 fintCancelarCFDiBuzon = 0
            Else
                 fintCancelarCFDiBuzon = 2
            End If
        End If 'Not DOMNodoCodigo Is Nothing
    End If 'Not ReaderRespuestaWS.Fault Is Nothing
    
Exit Function
NotificaErrorTimbre:
    If Err.Number > 0 And Err.Number <> 1001 Then
           strMensaje = "Ocurrió un error al solicitar la cancelación del comprobante." & vbNewLine & _
                        "Folio del documento: " & strFolio & vbNewLine & _
                        "Número de error: " & Err.Number & vbNewLine & _
                        "Origen: " & Err.Source & vbNewLine & _
                        "Descripción: " & Err.Description & vbNewLine & _
                        "¿Desea continuar?"
         'Se muestra el mensaje de error en pantalla
          If MsgBox(strMensaje, vbCritical + vbYesNo, "Mensaje") = vbYes Then
             fintCancelarCFDiBuzon = 0
          Else
             fintCancelarCFDiBuzon = 2
          End If
    End If
End Function


Public Sub pMensajeCanelacionCFDi(vllngIDComprobante As Long, strTipo As String)

    Dim DOMResponseXML As MSXML2.DOMDocument 'XML para leer los datos de la respuesta del WebService
    Dim strCodigoWS As String                'Almacena el código de respuesta del WebService
    Dim strUUIDWS As String                  'Almacena el código del comprobante regresado por el WebService
    Dim strDescripcionWS As String           'Almacena el mensaje de respuesta del WebService
    Dim strFechaWS As String                 'Almacena la fecha de cancelación del WebService
    
    Dim ObjRs As New ADODB.Recordset
    Dim ObjSentencia As String
    
    Dim strMensaje As String
    Dim blnCFDI As Boolean
    Dim lngConsecutivo As Long
    Dim blnMensajeCFDI As Boolean
    
    Dim strSQLPrueba As String
    Dim rsConexion As New ADODB.Recordset
    Dim rsPruebaOReal As New ADODB.Recordset
    
    
    
    blnMensajeCFDI = True
    
    'Se busca información del comprobante fiscal digital para revisar si es CFDi
    ObjSentencia = "SELECT intIdComprobante,vchUUID FROM GnComprobanteFiscalDigital " & _
                   "WHERE intComprobante = " & vllngIDComprobante & " AND chrTipoComprobante = '" & strTipo & "'"
    
    Set ObjRs = frsRegresaRs(ObjSentencia, adLockOptimistic, adOpenDynamic)
    
    'Si no hay información, significa que se está utilizando formato FÍSICO, se permitirá cancelar el comprobante
    If ObjRs.RecordCount = 0 Then
       blnMensajeCFDI = False
    Else
        blnCFDI = IIf(IsNull(ObjRs!VCHUUID), False, True)
        If blnCFDI Then
           lngConsecutivo = ObjRs!INTIDCOMPROBANTE
           blnMensajeCFDI = True
        Else
           blnMensajeCFDI = False
        End If
    End If
  
    If blnMensajeCFDI = True Then
        ObjSentencia = "Select * from GNACUSECANCELACIONCFDI where INTIDCOMPROBANTE = " & lngConsecutivo
        Set ObjRs = frsRegresaRs(ObjSentencia)
        If ObjRs.RecordCount > 0 Then
            Dim rsPAC As New ADODB.Recordset
            Dim intPAC As Integer
            
            'Se obtiene el PAC para el proceso de cancelación (Buzón Fiscal: INTIDPAC = 1) (PAX: INTIDPAC = 2)
            Set rsPAC = frsEjecuta_SP(CStr(vgintClaveEmpresaContable), "SP_GNSELCONFIGPAC")
            If rsPAC.RecordCount > 0 Then
                intPAC = Val(rsPAC!PAC)
                
                Dim DOMNodoCodigo As MSXML2.IXMLDOMNode
             
                Select Case intPAC
                    Case 1 '>> Buzón Fiscal <<
                        Set DOMResponseXML = New MSXML2.DOMDocument
                        DOMResponseXML.loadXML (CStr(ObjRs!CLBXMLACUSE))
                        Set DOMNodoCodigo = DOMResponseXML.selectSingleNode("S:Envelope/S:Body/ns3:ResponseCancelaCFDi/ns2:Result")
                        If Not DOMNodoCodigo Is Nothing Then
                            strCodigoWS = Trim(DOMNodoCodigo.selectSingleNode("ns2:Message/@code").Text)
                            strDescripcionWS = Trim(DOMNodoCodigo.selectSingleNode("ns2:Message/@message").Text)
                            If Val(strCodigoWS) = 0 Then
                                 strMensaje = "El comprobante se ha cancelado con éxito."
                                 MsgBox strMensaje, vbInformation + vbOKOnly, "Mensaje"
                            Else
                                 strMensaje = strDescripcionWS
                                 MsgBox strMensaje, vbExclamation + vbOKOnly, "Mensaje"
                            End If
                        End If
                    
                    Case 2 '>> PAX <<
                        Set DOMResponseXML = New MSXML2.DOMDocument
                        DOMResponseXML.loadXML (CStr(ObjRs!CLBXMLACUSE))
                        'Se verifica si el WS regresó información...
                        Set DOMNodoCodigo = DOMResponseXML.selectSingleNode("Cancelacion/Folios")
                        'If Not DOMNodoCodigo Is Nothing Then
                        strCodigoWS = Trim(DOMNodoCodigo.selectSingleNode("UUIDEstatus").Text)
                        strDescripcionWS = Trim(Replace(DOMNodoCodigo.selectSingleNode("UUIDdescripcion").Text, "- ", ""))
                          '201 - El folio se ha cancelado con éxito.
                          '202 - El CFDI ya había sido cancelado previamente.
                        strUUIDWS = Trim(DOMNodoCodigo.selectSingleNode("UUID").Text)
                        strFechaWS = fstrFormatoFechaHora(Replace(DOMNodoCodigo.selectSingleNode("UUIDfecha").Text, "- ", ""))
                        If vgblnNuevoEsquemaCancelacion Then
                            Select Case Val(strCodigoWS)
                                '-------------------------------------------------------------------------------------------------------------------------------
                                '| Este código está aparte porque tiene dos significados diferentes (facepalm)
                                '-------------------------------------------------------------------------------------------------------------------------------
                                Case 201
                                    If UCase(strDescripcionWS) = "201 - COMPROBANTE EN PROCESO DE SER CANCELADO." Then
                                        '| 201 - Comprobante En Proceso de ser Cancelado"
                                        strMensaje = "El comprobante se encuentra en proceso de ser cancelado." & vbNewLine & _
                                                     "Folio en espera: " & strUUIDWS & vbNewLine & _
                                                     "Fecha en espera: " & strFechaWS
                                        MsgBox strMensaje, vbInformation + vbOKOnly, "Mensaje"
                                    Else
                                        '| 201 - El folio se ha cancelado con éxito.
                                        strMensaje = "El comprobante se ha cancelado con éxito." & vbNewLine & _
                                                     "Folio cancelado: " & strUUIDWS & vbNewLine & _
                                                     "Fecha cancelación: " & strFechaWS
                                        MsgBox strMensaje, vbInformation + vbOKOnly, "Mensaje"
                                    End If
                                '-------------------------------------------------------------------------------------------------------------------------------
                                '| Cancelaciones exitosas
                                '-------------------------------------------------------------------------------------------------------------------------------
                                Case 202, 107, 103
                                    '| 202 - Comprobante previamente cancelado
                                    '| 107 – El CFDI ha sido Cancelado por Plazo Vencido
                                    '| 103 – El CFDI ha sido Cancelado Previamente por Aceptación del Receptor
                                    strMensaje = "El comprobante se ha cancelado con éxito." & vbNewLine & _
                                                 "Folio cancelado: " & strUUIDWS & vbNewLine & _
                                                 "Fecha cancelación: " & strFechaWS
                                    MsgBox strMensaje, vbInformation + vbOKOnly, "Mensaje"
                                '-------------------------------------------------------------------------------------------------------------------------------
                                '| Cancelaciones que se quedaron en espera
                                '-------------------------------------------------------------------------------------------------------------------------------
                                Case 105
                                    '| 105 – El CFDI no se puede Cancelar por que tiene Estatus De “En espera De Aceptación”
                                    strMensaje = "El comprobante se encuentra en espera de aceptación." & vbNewLine & _
                                                 "Folio en espera: " & strUUIDWS & vbNewLine & _
                                                 "Fecha en espera: " & strFechaWS
                                    MsgBox strMensaje, vbInformation + vbOKOnly, "Mensaje"
                                '-------------------------------------------------------------------------------------------------------------------------------
                                '| Cancelaciones que se quedaron en espera
                                '-------------------------------------------------------------------------------------------------------------------------------
                                Case 106
                                    '| 106 – El CFDI no se puede Cancelar por que tiene Estatus de “En Proceso”.
                                    strMensaje = "El comprobante se encuentra en proceso de aceptación." & vbNewLine & _
                                                 "Folio en proceso: " & strUUIDWS & vbNewLine & _
                                                 "Fecha en proceso: " & strFechaWS
                                    MsgBox strMensaje, vbInformation + vbOKOnly, "Mensaje"
                                '-------------------------------------------------------------------------------------------------------------------------------
                                '| Cancelaciones rechazadas por el receptor
                                '-------------------------------------------------------------------------------------------------------------------------------
                                Case 104
                                    '| 104 – El CFDI no se puede Cancelar por que fue Rechazado Previamente.
                                    strMensaje = "El comprobante no se puede cancelar por que fue rechazado por el emisor." & vbNewLine & _
                                                 "Folio rechazado: " & strUUIDWS & vbNewLine & _
                                                 "Fecha rechazo: " & strFechaWS
                                    MsgBox strMensaje, vbError + vbOKOnly, "Mensaje"
                                '-------------------------------------------------------------------------------------------------------------------------------
                                '|  Código de error genérico para las excepciones no controladas del SAT
                                '-------------------------------------------------------------------------------------------------------------------------------
                                Case 999
                                    strMensaje = "El servicio del SAT no responde a la petición de cancelación. " & vbNewLine & _
                                                       "Por favor intente más tarde."
                                    MsgBox strMensaje, vbError + vbOKOnly, "Mensaje"
                                '-------------------------------------------------------------------------------------------------------------------------------
                                '| Cancelaciones fallidas
                                '-------------------------------------------------------------------------------------------------------------------------------
                                Case Else
                                    strMensaje = strDescripcionWS & vbNewLine & _
                                                 "Folio fiscal: " & strUUIDWS
                                    MsgBox strMensaje, vbExclamation + vbOKOnly, "Mensaje"
                            End Select
                        Else
                            If Val(strCodigoWS) = 201 Then
                                 strMensaje = "El comprobante se ha cancelado con éxito." & vbNewLine & _
                                              "Folio cancelado: " & strUUIDWS & vbNewLine & _
                                              "Fecha cancelación: " & strFechaWS
                                 MsgBox strMensaje, vbInformation + vbOKOnly, "Mensaje"
                            Else
                                 strMensaje = strDescripcionWS & vbNewLine & _
                                              "Folio fiscal: " & strUUIDWS
                                 MsgBox strMensaje, vbExclamation + vbOKOnly, "Mensaje"
                            End If
                        End If
                    Case 3 '>>Prodigia<<'
                        strSQLPrueba = "SELECT VCHVALOR FROM SIPARAMETRO WHERE VCHNOMBRE = 'BITFACTURACIONMODOPRUEBA'"
                        Set rsPruebaOReal = frsRegresaRs(strSQLPrueba, adLockReadOnly, adOpenForwardOnly)
                        Set DOMResponseXML = New MSXML2.DOMDocument
                        DOMResponseXML.loadXML (CStr(ObjRs!CLBXMLACUSE))
                            'If rsPruebaOReal!vchvalor = 0 Then
                            'Set DOMNodoCodigo = DOMResponseXML.selectSingleNode("Cancelacion/Folios")
                            'Else
                                'Se verifica si el WS regresó información...
                            Set DOMNodoCodigo = DOMResponseXML.selectSingleNode("servicioCancel")
                            'End If
                            If Not DOMNodoCodigo Is Nothing Then
                                'If rsPruebaOReal!vchvalor = 0 Then
                                    'If Not DOMNodoCodigo Is Nothing Then
                                    'strCodigoWS = Trim(DOMNodoCodigo.selectSingleNode("UUIDEstatus").Text)
                                    'strDescripcionWS = Trim(Replace(DOMNodoCodigo.selectSingleNode("UUIDdescripcion").Text, "- ", ""))
                                      '201 - El folio se ha cancelado con éxito.
                                      '202 - El CFDI ya había sido cancelado previamente.
                                    'strUUIDWS = Trim(DOMNodoCodigo.selectSingleNode("UUID").Text)
                                    'strFechaWS = fstrFormatoFechaHora(Replace(DOMNodoCodigo.selectSingleNode("UUIDfecha").Text, "- ", ""))
                                'Else
                                    strCodigoWS = Trim(DOMResponseXML.selectSingleNode("servicioCancel/cancelaciones/cancelacion/codigo").Text)
                                    strDescripcionWS = DOMResponseXML.selectSingleNode("servicioCancel/cancelaciones/cancelacion/mensaje").Text
                                    strUUIDWS = DOMResponseXML.selectSingleNode("servicioCancel/cancelaciones/cancelacion/uuid").Text
                                'End If
                                
                                Select Case Val(strCodigoWS)
                                '-------------------------------------------------------------------------------------------------------------------------------
                                '| Este código está aparte porque tiene dos significados diferentes (facepalm)
                                '-------------------------------------------------------------------------------------------------------------------------------
                                Case 96
                                    '| 96 - Comprobante En Proceso de ser Cancelado"
                                    strMensaje = "El comprobante se encuentra en proceso de ser cancelado." & vbNewLine & _
                                                "Folio en espera: " & strUUIDWS ' & vbNewLine & _
                                                '"Fecha en espera: " & strFechaWS
                                        MsgBox strMensaje, vbInformation + vbOKOnly, "Mensaje"
                                Case 201
                                        '| 201 - El folio se ha cancelado con éxito.
                                    strMensaje = "El comprobante se ha cancelado con éxito." & vbNewLine & _
                                            "Folio cancelado: " & strUUIDWS ' & vbNewLine & _
                                            '"Fecha cancelación: " & strFechaWS
                                        MsgBox strMensaje, vbInformation + vbOKOnly, "Mensaje"
                                '-------------------------------------------------------------------------------------------------------------------------------
                                '| Cancelaciones exitosas
                                '-------------------------------------------------------------------------------------------------------------------------------
                                Case 202, 98, 95
                                    '| 202 - Comprobante previamente cancelado
                                    '| 107 – El CFDI ha sido Cancelado por Plazo Vencido
                                    '| 103 – El CFDI ha sido Cancelado Previamente por Aceptación del Receptor
                                    strMensaje = "El comprobante se ha cancelado con éxito." & vbNewLine & _
                                                 "Folio cancelado: " & strUUIDWS ' & vbNewLine & _
                                                 '"Fecha cancelación: " & strFechaWS
                                    MsgBox strMensaje, vbInformation + vbOKOnly, "Mensaje"
                                '-------------------------------------------------------------------------------------------------------------------------------
                                '| Cancelaciones que se quedaron en espera
                                '-------------------------------------------------------------------------------------------------------------------------------
                                Case 105
                                    '| 105 – El CFDI no se puede Cancelar por que tiene Estatus De “En espera De Aceptación”
                                    strMensaje = "El comprobante se encuentra en espera de aceptación." & vbNewLine & _
                                                 "Folio en espera: " & strUUIDWS ' & vbNewLine & _
                                                 '"Fecha en espera: " & strFechaWS
                                    MsgBox strMensaje, vbInformation + vbOKOnly, "Mensaje"
                                Case 106
                                    '| 106 – El CFDI no se puede Cancelar por que tiene Estatus de “En Proceso”.
                                    strMensaje = "El comprobante se encuentra en proceso de aceptación." & vbNewLine & _
                                                 "Folio en proceso: " & strUUIDWS ' & vbNewLine & _
                                                 '"Fecha en proceso: " & strFechaWS
                                    MsgBox strMensaje, vbInformation + vbOKOnly, "Mensaje"
                                '-------------------------------------------------------------------------------------------------------------------------------
                                '| Cancelaciones rechazadas por el receptor
                                '-------------------------------------------------------------------------------------------------------------------------------
                                Case 97
                                    strMensaje = "El comprobante no se puede cancelar por que fue rechazado por el emisor." & vbNewLine & _
                                                 "Folio rechazado: " & strUUIDWS '& vbNewLine & _
                                                 '"Fecha rechazo: " & strFechaWS
                                    MsgBox strMensaje, vbError + vbOKOnly, "Mensaje"
                                '-------------------------------------------------------------------------------------------------------------------------------
                                '|  Código de error genérico para las excepciones no controladas del SAT
                                '-------------------------------------------------------------------------------------------------------------------------------
                                Case 7
                                '-------------------------------------------------------------------------------------------------------------------------------
                                '| Cancelaciones fallidas
                                '-------------------------------------------------------------------------------------------------------------------------------
                                    strMensaje = "El servicio del SAT no responde a la petición de cancelación. " & vbNewLine & _
                                                       "Por favor intente más tarde."
                                    MsgBox strMensaje, vbError + vbOKOnly, "Mensaje"
                                Case Else
                                '-------------------------------------------------------------------------------------------------------------------------------
                                '| Cancelaciones fallidas
                                '-------------------------------------------------------------------------------------------------------------------------------
                                    strMensaje = strDescripcionWS & vbNewLine & _
                                                 "Folio fiscal: " & strUUIDWS
                                    MsgBox strMensaje, vbInformation + vbOKOnly, "Mensaje"
                            End Select
                        
                        Else
                            If Val(strCodigoWS) = 201 Then
                                 strMensaje = "El comprobante se ha cancelado con éxito." & vbNewLine & _
                                              "Folio cancelado: " & strUUIDWS '& vbNewLine & _
                                              '"Fecha cancelación: " & strFechaWS
                                 MsgBox strMensaje, vbInformation + vbOKOnly, "Mensaje"
                            Else
                                 strMensaje = "El comprobante se ha cancelado con éxito."
                                 MsgBox strMensaje, vbInformation + vbOKOnly, "Mensaje"
                            End If
                        End If
 
                End Select
            End If
        Else
             Select Case strTipo
                Case "FA"
                     'La factura se canceló satisfactoriamente.
                      MsgBox SIHOMsg(365), vbInformation, "Mensaje"
                Case "CA"
                       
                      MsgBox Replace(SIHOMsg(365), "factura", "nota de cargo"), vbInformation, "Mensaje"
                Case "CR"
            
                      MsgBox Replace(SIHOMsg(365), "factura", "nota de crédito"), vbInformation, "Mensaje"
                Case "DO"
                      MsgBox Replace(SIHOMsg(365), "La factura", "El donativo"), vbInformation, "Mensaje"
             End Select
        End If ' ObjRS.RecordCount > 0
    Else
       Select Case strTipo
            Case "FA"
                 'La factura se canceló satisfactoriamente.
                  MsgBox SIHOMsg(365), vbInformation, "Mensaje"
            Case "CA"
                   
                  MsgBox Replace(SIHOMsg(365), "factura", "nota de cargo"), vbInformation, "Mensaje"
            Case "CR"
        
                  MsgBox Replace(SIHOMsg(365), "factura", "nota de crédito"), vbInformation, "Mensaje"
            Case "DO"
                  MsgBox Replace(SIHOMsg(365), "La factura", "El donativo"), vbInformation, "Mensaje"
        End Select
    End If 'blnMensajeCFDI = True
End Sub


Public Sub pLogTimbrado(vlOperacion As Integer, Optional VCBLXMLREQUEST As String, Optional VCBLRESPUESTAWS As String, Optional VVCHIDREFERENCIA As String, Optional VVCHMENSAJEERROR As String)
    'vlOperacion = 0 carga variables
    'vlOperacion = 1 guarda variables
    'vloperacion = 2 inicializa el arreglo
    
    Dim ObjRs As New ADODB.Recordset
    Dim ObjRsBan As Boolean
    Dim contador As Integer
    Dim vlstrsql As String
On Error GoTo NotificaError

1    Select Case vlOperacion
           Case 0
3                ReDim Preserve vlArrLogTimbrado(intContadorArrlogTimbrado)
4                vlArrLogTimbrado(intContadorArrlogTimbrado).vgXMLREQUEST = VCBLXMLREQUEST
5                vlArrLogTimbrado(intContadorArrlogTimbrado).vgRESPUESTAWS = VCBLRESPUESTAWS
6                vlArrLogTimbrado(intContadorArrlogTimbrado).vgIDREFERENCIA = VVCHIDREFERENCIA
7                vlArrLogTimbrado(intContadorArrlogTimbrado).vgMENSAJEERROR = VVCHMENSAJEERROR
8                intContadorArrlogTimbrado = intContadorArrlogTimbrado + 1
           Case 1
11               ObjRsBan = False
12               For contador = 0 To intContadorArrlogTimbrado - 1
13                    If Not ObjRsBan Then
14                       Set ObjRs = frsRegresaRs("Select * from GNLOGERRORTIMBRADOCFDI where intconsecutivo = -1", adLockOptimistic)
15                       ObjRsBan = True
16                    End If
17                    ObjRs.AddNew
18                    ObjRs!CBLXMLREQUEST = vlArrLogTimbrado(contador).vgXMLREQUEST
19                    ObjRs!CBLRESPUESTAWS = vlArrLogTimbrado(contador).vgRESPUESTAWS
20                    ObjRs!VCHIDREFERENCIA = vlArrLogTimbrado(contador).vgIDREFERENCIA
21                    ObjRs!VCHMENSAJEERROR = vlArrLogTimbrado(contador).vgMENSAJEERROR
22                    ObjRs!dtmFechahora = fdtmServerFechaHora
23                    ObjRs.Update
24               Next contador
25               If ObjRsBan Then
26                  ObjRs.Close
27               End If
28               intContadorArrlogTimbrado = 0
29               ReDim vlArrLogTimbrado(intContadorArrlogTimbrado)
          Case 2
31               intContadorArrlogTimbrado = 0
32               ReDim vlArrLogTimbrado(intContadorArrlogTimbrado)
33   End Select
Exit Sub
NotificaError:
   Call pRegistraError(Err.Number, Err.Description, cgstrModulo, ("pLogTimbrado" & " ,Linea:" & Erl()), , False)
   Err.Clear
End Sub
Public Sub pCFDiPendienteCancelar(lngIdComprobante As Long, strTipoComprobante As String, Optional intAccion As Integer = 0)

'intAccion = 0 agrega información al arreglo de los CFDi
'intAccion = 1 guarda la información del arreglo en la base de datos

Dim RsComprobante As New ADODB.Recordset
Dim StrParamtros As String
Dim intcontador As Integer

On Error GoTo NotificaError

1  If intAccion = 0 Then 'obtener datos de la base
2    StrParamtros = CStr(lngIdComprobante) & "|" & strTipoComprobante
3    Set RsComprobante = frsEjecuta_SP(StrParamtros, "sp_PvSelCFDIpCancelar")
4     If RsComprobante.RecordCount > 0 Then
5        intcontadorCFDiPendienteCancelar = intcontadorCFDiPendienteCancelar + 1
6        ReDim Preserve vlArrCFDiPendienteCancelar(intcontadorCFDiPendienteCancelar)
7        vlArrCFDiPendienteCancelar(intcontadorCFDiPendienteCancelar).INTCOMPROBANTE = lngIdComprobante
8        vlArrCFDiPendienteCancelar(intcontadorCFDiPendienteCancelar).VCHTIPOCOMPROBANTE = strTipoComprobante
9        vlArrCFDiPendienteCancelar(intcontadorCFDiPendienteCancelar).CHRFOLIOCOMPROBANTE = RsComprobante!CHRFOLIOCOMPROBANTE
10       vlArrCFDiPendienteCancelar(intcontadorCFDiPendienteCancelar).dtmFechahora = RsComprobante!dtmFechahora
11       vlArrCFDiPendienteCancelar(intcontadorCFDiPendienteCancelar).MNYSUBTOTAL = RsComprobante!MNYSUBTOTAL
12       vlArrCFDiPendienteCancelar(intcontadorCFDiPendienteCancelar).MNYDESCUENTO = RsComprobante!MNYDESCUENTO
13       vlArrCFDiPendienteCancelar(intcontadorCFDiPendienteCancelar).MNYIVA = RsComprobante!MNYIVA
14       vlArrCFDiPendienteCancelar(intcontadorCFDiPendienteCancelar).MNYTOTAL = RsComprobante!MNYTOTAL
15       vlArrCFDiPendienteCancelar(intcontadorCFDiPendienteCancelar).BITPESOS = RsComprobante!BITPESOS
16       vlArrCFDiPendienteCancelar(intcontadorCFDiPendienteCancelar).SMIDEPARTAMENTO = RsComprobante!SMIDEPARTAMENTO
17       vlArrCFDiPendienteCancelar(intcontadorCFDiPendienteCancelar).CHRNOMBRE = Trim(RsComprobante!CHRNOMBRE)
18       vlArrCFDiPendienteCancelar(intcontadorCFDiPendienteCancelar).INTMOVPACIENTE = RsComprobante!INTMOVPACIENTE
19       vlArrCFDiPendienteCancelar(intcontadorCFDiPendienteCancelar).CHRTIPOPACIENTE = RsComprobante!CHRTIPOPACIENTE
20       vlArrCFDiPendienteCancelar(intcontadorCFDiPendienteCancelar).intNumCliente = IIf(IsNull(RsComprobante!intNumCliente), 0, RsComprobante!intNumCliente)
21       vlArrCFDiPendienteCancelar(intcontadorCFDiPendienteCancelar).VCHRFCEMISOR = RsComprobante!VCHRFCEMISOR
22       vlArrCFDiPendienteCancelar(intcontadorCFDiPendienteCancelar).VCHRFCRECEPTOR = RsComprobante!VCHRFCRECEPTOR
23       vlArrCFDiPendienteCancelar(intcontadorCFDiPendienteCancelar).INTIDCOMPROBANTE = RsComprobante!INTIDCOMPROBANTE
24       vlArrCFDiPendienteCancelar(intcontadorCFDiPendienteCancelar).VCHUUID = IIf(IsNull(RsComprobante!VCHUUID), "", RsComprobante!VCHUUID)
25       vlArrCFDiPendienteCancelar(intcontadorCFDiPendienteCancelar).CLBCOMPROBANTEFISCAL = RsComprobante!CLBCOMPROBANTEFISCAL
26    End If
27    RsComprobante.Close
28 Else
29    Set RsComprobante = frsRegresaRs("Select * from PVCANCELARCOMPROBANTES where INTCOMPROBANTE  = -1", adLockOptimistic)
30    For intcontador = 1 To intcontadorCFDiPendienteCancelar
31        RsComprobante.AddNew
32        RsComprobante!INTCOMPROBANTE = vlArrCFDiPendienteCancelar(intcontador).INTCOMPROBANTE
33        RsComprobante!VCHTIPOCOMPROBANTE = vlArrCFDiPendienteCancelar(intcontador).VCHTIPOCOMPROBANTE
34        RsComprobante!CHRFOLIOCOMPROBANTE = vlArrCFDiPendienteCancelar(intcontador).CHRFOLIOCOMPROBANTE
35        RsComprobante!dtmFechahora = vlArrCFDiPendienteCancelar(intcontador).dtmFechahora
36        RsComprobante!MNYSUBTOTAL = vlArrCFDiPendienteCancelar(intcontador).MNYSUBTOTAL
37        RsComprobante!MNYDESCUENTO = vlArrCFDiPendienteCancelar(intcontador).MNYDESCUENTO
38        RsComprobante!MNYIVA = vlArrCFDiPendienteCancelar(intcontador).MNYIVA
39        RsComprobante!MNYTOTAL = vlArrCFDiPendienteCancelar(intcontador).MNYTOTAL
40        RsComprobante!BITPESOS = vlArrCFDiPendienteCancelar(intcontador).BITPESOS
41        RsComprobante!SMIDEPARTAMENTO = vlArrCFDiPendienteCancelar(intcontador).SMIDEPARTAMENTO
42        RsComprobante!CHRNOMBRE = vlArrCFDiPendienteCancelar(intcontador).CHRNOMBRE
43        RsComprobante!INTMOVPACIENTE = vlArrCFDiPendienteCancelar(intcontador).INTMOVPACIENTE
44        RsComprobante!CHRTIPOPACIENTE = vlArrCFDiPendienteCancelar(intcontador).CHRTIPOPACIENTE
45        RsComprobante!intNumCliente = vlArrCFDiPendienteCancelar(intcontador).intNumCliente
46        RsComprobante!VCHRFCEMISOR = vlArrCFDiPendienteCancelar(intcontador).VCHRFCEMISOR
47        RsComprobante!VCHRFCRECEPTOR = vlArrCFDiPendienteCancelar(intcontador).VCHRFCRECEPTOR
48        RsComprobante!INTIDCOMPROBANTE = vlArrCFDiPendienteCancelar(intcontador).INTIDCOMPROBANTE
49        RsComprobante!VCHUUID = vlArrCFDiPendienteCancelar(intcontador).VCHUUID
50        RsComprobante!CLBCOMPROBANTEFISCAL = vlArrCFDiPendienteCancelar(intcontador).CLBCOMPROBANTEFISCAL
51        RsComprobante!BITPENDIENTECANCELAR = 1
52        RsComprobante.Update
53    Next intcontador
54    RsComprobante.Close
55 End If
Exit Sub
NotificaError:
   Call pRegistraError(Err.Number, Err.Description, cgstrModulo, ("pCFDiPendienteCancelar" & " ,Linea:" & Erl()), , False)
   Err.Clear
End Sub


Public Sub pTimbradoProdigia(DOMFacturaXMLsinTimbrar As MSXML2.DOMDocument, strRefID As String, strRutaRequestTimbrado As String)
1     On Error GoTo NotificaErrorTimbre:

          Dim DOMRequestXML As MSXML2.DOMDocument
          Dim DOMTempXML As MSXML2.DOMDocument
2         Set DOMTempXML = New MSXML2.DOMDocument
          Dim strRequestXML As String
          Dim rsConexion As New ADODB.Recordset
          Dim rsPruebaOReal As New ADODB.Recordset
          Dim strURLWSTimbrado As String
          Dim strURLWMTimbrado As String
          Dim strURLXMLNSTimbrado As String       'Agregado para evitar la comparación que indica si el WS es de pruebas o no
          Dim SerializerWS As SoapSerializer30    'Para serializar el XML
          Dim ReaderRespuestaWS As SoapReader30   'Para leer la respuesta del WebService
          Dim ConectorWS As ISoapConnector        'Para conectarse al WebService
          Dim DOMResponseXML2 As MSXML2.DOMDocument 'XML para leer los datos de la respuesta del WebService
          
          Dim strTipoDocumento As String
          Dim strUsuario As String
          Dim strPassword As String
          Dim strContrato As String
          Dim strVersion As String
          Dim intEstructura As Integer
          Dim intlineaGoto As Integer
          Dim strUUId As String
          Dim strPrimeraError As String
          Dim strSQLPrueba As String
          Dim httpRequest As WinHttp.WinHttpRequest
          Set httpRequest = New WinHttp.WinHttpRequest
          
3         intlineaGoto = 0
          
4         strSQLPrueba = "SELECT VCHVALOR FROM SIPARAMETRO WHERE VCHNOMBRE = 'BITFACTURACIONMODOPRUEBA'"
5         Set rsPruebaOReal = frsRegresaRs(strSQLPrueba, adLockReadOnly, adOpenForwardOnly)
          
          'Se especifican las rutas para la conexión con el servicio de timbrado
6     Set rsConexion = frsEjecuta_SP(CStr(vgintClaveEmpresaContable), "SP_GNSELCONFIGPAC")
      
      If vgstrVersionCFDI = "4.0" Then
        strURLWSTimbrado = rsConexion!URLWSTimbrado40 '(Pruebas = "https://test.paxfacturacion.com.mx:453/webservices/wcfRecepcionASMX.asmx", Producción = "https://www.paxfacturacion.mx/webservices/wcfRecepcionasmx.asmx")
      Else
7       strURLWSTimbrado = rsConexion!URLWSTimbrado '(Pruebas = "https://test.paxfacturacion.com.mx:453/webservices/wcfRecepcionASMX.asmx", Producción = "https://www.paxfacturacion.mx/webservices/wcfRecepcionasmx.asmx")
      End If
8         If rsPruebaOReal!vchvalor = 0 Then
9       strURLWMTimbrado = rsConexion!URLWMTimbrado
10        Else
11      strURLWMTimbrado = "timbradoPrueba"
12        End If
13    strURLXMLNSTimbrado = rsConexion!URLXMLNSTimbrado  '(Pruebas = "https://test.paxfacturacion.com.mx:453")
14    strUsuario = rsConexion!Usuario
15    strPassword = rsConexion!Password
16        strContrato = rsConexion!contrato
17        If vgstrVersionCFDI = "3.2" Then
18      Select Case Mid(strRefID, 14, 2)
            Case "FA"
19              strTipoDocumento = "Factura"
20          Case "CR"
21              strTipoDocumento = "Nota de Crédito"
22          Case "CA"
23              strTipoDocumento = "Nota de Cargo"
24          Case "DO"
25              strTipoDocumento = "Recibo de donativos"
26          Case "NO"
27              strTipoDocumento = "Recibo de Nomina"
28          Case Else
29              strTipoDocumento = "XX"
30      End Select
31        Else
32      Select Case Mid(strRefID, 14, 2)
            Case "FA"
33              strTipoDocumento = "01"
34          Case "CR"
35              strTipoDocumento = "02"
36          Case "CA"
37              strTipoDocumento = "03"
38          Case "DO"
39              strTipoDocumento = "08"
40          Case "RE"
41              strTipoDocumento = "09"
42          Case "NO"
43              strTipoDocumento = "10"
44          Case "AN"
45              strTipoDocumento = "01"
46          Case "AA"
47              strTipoDocumento = "02"
48          Case Else
49              strTipoDocumento = "XX"
50      End Select
51        End If
52        intEstructura = 0 'Se especifica el tipo de estructura (0 por default)
53        If vgstrVersionCFDI = "3.2" Then
54      strVersion = DOMFacturaXMLsinTimbrar.selectSingleNode("cfdi:Comprobante/@version").Text 'Se especifica la versión
55        Else
56      strVersion = DOMFacturaXMLsinTimbrar.selectSingleNode("cfdi:Comprobante/@Version").Text 'Se especifica la versión
57        End If
          
          'Se especifica la ruta para la creación del archivo Request del Timbrado
58    vlstrRutaRequestTimbrado = Trim(strRutaRequestTimbrado)

          'Se forma el archivo Request del Timbrado
59    Set DOMRequestXML = fnDOMRequestXMLProdigia(DOMFacturaXMLsinTimbrar, intEstructura, strTipoDocumento, strContrato, strUsuario, strPassword, strVersion, strURLXMLNSTimbrado, strURLWMTimbrado)
         
60    Set ConectorWS = New HttpConnector30
          
          'La URL que atenderá nuestra solicitud
61    ConectorWS.Property("EndPointURL") = strURLWSTimbrado
           
          'Ruta del WebMethod para el timbrado ("https://www.paxfacturacion.com.mx/fnEnviarXML")
62    ConectorWS.Property("SoapAction") = strURLWMTimbrado
          
          'Se configura timeOUT
63    ConectorWS.Property("Timeout") = "300000"
          '########################################## INICIA CONEXIÓN CON EL SERVICIO DE TIMBRADO ####################################
'64    ConectorWS.Connect
'
'65    ConectorWS.BeginMessage
'66    Set SerializerWS = New SoapSerializer30
'67      SerializerWS.Init ConectorWS.InputStream
'
'68      SerializerWS.StartEnvelope
'69                   SerializerWS.StartBody
'70                   SerializerWS.WriteXml vgStrSolicitudEnPlanoProdigia
'71                   SerializerWS.EndBody
'72      SerializerWS.EndEnvelope
'
'73      ConectorWS.EndMessage
      httpRequest.Open "POST", strURLWSTimbrado, False
      httpRequest.setRequestHeader "Content-Type", "text/xml"
      httpRequest.setRequestHeader "SOAPAction", strURLWMTimbrado
      httpRequest.send vgStrSolicitudEnPlanoProdigia

          
74    Set ReaderRespuestaWS = New SoapReader30
      ReaderRespuestaWS.loadXML httpRequest.responseBody
          '########################################## FINAL DE CONEXIÓN CON EL SERVICIO DE TIMBRADO ####################################
76    If Not ReaderRespuestaWS.Fault Is Nothing Then
       Dim strMensajeError As String
77     If Not ReaderRespuestaWS.FaultDetail Is Nothing Then
          Dim DOMNodoCodigo As MSXML2.IXMLDOMNode
          '----------------------------------------------------------------------------------------------- ERROR A NIVEL 1
          'Se obtiene el codigo del error devuelto por PAX
78        Set DOMNodoCodigo = ReaderRespuestaWS.FaultDetail.childNodes(0).Attributes.getNamedItem("codigo")
          'Se captura el error
79        If DOMNodoCodigo Is Nothing Then
80           CFDiintResultadoTimbrado = 1 'queda pendiente el timbrado
81           intlineaGoto = 33
82           strMensajeError = "Ocurrió un error al solicitar el servicio de timbrado" & vbNewLine & vbNewLine & _
                               "Número de error: 1000" & vbNewLine & _
                               "Descripción: Error de timbrado de nivel de capa 1"
83           pLogTimbrado 0, DOMRequestXML.Text, strMensajeError, Mid(strRefID, 14, Len(strRefID) - 13), strMensajeError
84           GoTo NotificaErrorTimbre
85        Else
86           If fDetieneProcesoErrorProdigia(DOMNodoCodigo.Text) Then
87              CFDiintResultadoTimbrado = 2 'error identificado, no queda pendiente el timbre
88           Else
89              CFDiintResultadoTimbrado = 1 'error no identificado, queda pendiente el timbre
90           End If
91           strMensajeError = "Ocurrió un error al solicitar el servicio de timbrado" & vbNewLine & vbNewLine & _
                                            "Número de error: " & DOMNodoCodigo.Text & vbNewLine & _
                                            "Descripción: " & ReaderRespuestaWS.FaultString.Text
92           pLogTimbrado 0, DOMRequestXML.Text, ReaderRespuestaWS.FaultString.Text, Mid(strRefID, 14, Len(strRefID) - 13), strMensajeError
93           intlineaGoto = 41
94           GoTo NotificaErrorTimbre
95        End If
96     End If
97        Else
        '---tls1.2---
        Dim strErrorWS As String
        Dim strDescErrorWS As String
98      Set DOMResponseXML2 = New MSXML2.DOMDocument
        Dim resp As String
        'MsgBox (httpRequest.responseText), vbOKOnly, "Timbrado"
        If rsPruebaOReal!vchvalor = 0 Then
            resp = Replace(Replace(Replace(Replace(Replace(httpRequest.responseText, "</return></ns2:timbradoResponse></S:Body></S:Envelope>", ""), "<S:Envelope xmlns:S=""http://schemas.xmlsoap.org/soap/envelope/""><S:Body><ns2:timbradoResponse xmlns:ns2=""timbrado.ws.pade.mx""><return>", ""), "&lt;", "<"), "&gt;", ">"), "<?xml version='1.0' encoding='UTF-8'?><?xml version=""1.0"" encoding=""UTF-8""?>", "")
        Else
            resp = Replace(Replace(Replace(Replace(Replace(httpRequest.responseText, "</return></ns2:timbradoPruebaResponse></S:Body></S:Envelope>", ""), "<S:Envelope xmlns:S=""http://schemas.xmlsoap.org/soap/envelope/""><S:Body><ns2:timbradoPruebaResponse xmlns:ns2=""timbrado.ws.pade.mx""><return>", ""), "&lt;", "<"), "&gt;", ">"), "<?xml version='1.0' encoding='UTF-8'?><?xml version=""1.0"" encoding=""UTF-8""?>", "")
        End If
            
99      DOMResponseXML2.loadXML (resp)
        '---tls1.2---
'99      DOMResponseXML2.loadXML (ReaderRespuestaWS.Body.Text)
        'Se verifica si el WS regresó algún código de error...
        'MsgBox (resp), vbOKOnly, "Timbrado"
        
100     strErrorWS = Trim(DOMResponseXML2.selectSingleNode("servicioTimbrado/codigo").Text)
101     If strErrorWS <> "0" Then
102         strDescErrorWS = DOMResponseXML2.selectSingleNode("servicioTimbrado/mensaje").Text
103     Else
104         strDescErrorWS = ""
105     End If
            
106     strDescErrorWS = Replace(strDescErrorWS, "|", "/")
        
107     If (Val(strErrorWS) > 0 And Val(strErrorWS) <> 307) Or InStr(strErrorWS, "CFDI") > 0 Or InStr(strErrorWS, "CRP") > 0 Then 'Para determinar si regresó un número de error
        '------------------------------------------------------------------------------------------------- ERROR A NIVEL 2
108        If fDetieneProcesoErrorProdigia(strErrorWS) Then
109           CFDiintResultadoTimbrado = 2 'error identificado, no queda pendiente el timbre
110        Else
111           CFDiintResultadoTimbrado = 1 'error no identificado, queda pendiente el timbre
112        End If
113        strMensajeError = "Ocurrió un error al solicitar el servicio de timbrado" & vbNewLine & vbNewLine & _
                             "Número de error: " & strErrorWS & vbNewLine & _
                             "Descripción: " & strDescErrorWS
114        pLogTimbrado 0, DOMRequestXML.Text, resp, Mid(strRefID, 14, Len(strRefID) - 13), strMensajeError
115        intlineaGoto = 54
116        GoTo NotificaErrorTimbre
117     Else '-------------------------------------------------------------------------------------------- TIMBRADO CORRECTO
118        DOMTempXML.loadXML resp
           Dim strBase64 As String
119        strBase64 = DOMTempXML.selectSingleNode("servicioTimbrado/xmlBase64").Text
120        DOMFacturaXMLsinTimbrar.loadXML fstrUTF8ToUni(StrConv(Decode(strBase64), vbFromUnicode))
           'DOMFacturaXMLsinTimbrar.loadXML Decode(strBase64)
121        strUUId = DOMFacturaXMLsinTimbrar.selectSingleNode("cfdi:Comprobante/cfdi:Complemento/tfd:TimbreFiscalDigital/@UUID").Text
122        If strUUId <> "" Then
123           CFDiblnHaytimbre = True '*tenemos timbre,
124           pLogTimbrado 0, DOMRequestXML.Text, resp, Mid(strRefID, 14, Len(strRefID) - 13), "TIMBRADO EXITOSO"
125        Else
126           CFDiintResultadoTimbrado = 1 'error no identificado, queda pendiente el timbre
127           pLogTimbrado 0, DOMRequestXML.Text, resp, Mid(strRefID, 14, Len(strRefID) - 13), "TIMBRADO PENDIENTE"
128           strMensajeError = "Ocurrió un error al solicitar el servicio de timbrado" & vbNewLine & vbNewLine & _
                                "Descripción: No fue posible recuperar el folio fiscal."
129           intlineaGoto = 62
130           GoTo NotificaErrorTimbre
131        End If
132     End If
133       End If
134   Exit Sub
NotificaErrorTimbre:
    CFDiblnBanError = True
    CFDistrProcesoError = "pTimbradoProdigia"
       If Err.Number <> 0 Then 'llegó por error del código
    CFDiintResultadoTimbrado = 1 'el proceso queda pendiente de confirmación de timbre no sabemos si alcanzo a timbrar o no
    CFDiintLineaError = Erl()
    CFDistrDescripError = Err.Description
    CFDilngNumError = Err.Number
    strMensajeError = "Ocurrió un error al solicitar el servicio de timbrado" & vbNewLine & vbNewLine & _
                         "Número de error: " & Err.Number & " <" & Err.Description & "> " & " pTimbradoProdigia, Linea:" & Erl() & vbNewLine & _
                         "Origen: " & Err.Source
    pLogTimbrado 0, DOMRequestXML.Text, "Error: " & Err.Number & " " & Err.Description & " " & "Origen: " & Err.Source, Mid(strRefID, 14, Len(strRefID) - 13), strMensajeError
    Err.Clear 'limpiamos el error para que no se active en los demás procesos
    CFDiMostrarMensajeError = False
       Else ' llego por error del proceso de timbre(puede o no estar pendiente de timbre fiscal)
    CFDiintLineaError = intlineaGoto
    CFDistrDescripError = strMensajeError
    CFDilngNumError = -1
    CFDiMostrarMensajeError = False
    If CFDiintResultadoTimbrado = 2 Then MsgBox strMensajeError, vbCritical + vbOKOnly, "Mensaje"
       End If
End Sub

Private Function fnDOMRequestXMLProdigia(DOMFacturaXMLsinTimbrar As MSXML2.DOMDocument, intEstructura As Integer, strTipo As String, strContrato As String, strUsuario As String, strPassword As String, strVersion As String, strURLXMLNSTimbrado As String, strURLWMTimbrado As String) As MSXML2.DOMDocument
    Dim DOMRequestTimbradoXML As MSXML2.DOMDocument
    Set DOMRequestTimbradoXML = New MSXML2.DOMDocument
    
    Dim strNombreArchivoRequest As String
    If vgstrVersionCFDI = "3.2" Then
        strNombreArchivoRequest = Trim(DOMFacturaXMLsinTimbrar.selectSingleNode("cfdi:Comprobante/@serie").Text) + Trim(DOMFacturaXMLsinTimbrar.selectSingleNode("cfdi:Comprobante/@folio").Text) + ".xml"
    Else
        strNombreArchivoRequest = Trim(DOMFacturaXMLsinTimbrar.selectSingleNode("cfdi:Comprobante/@Serie").Text) + Trim(DOMFacturaXMLsinTimbrar.selectSingleNode("cfdi:Comprobante/@Folio").Text) + ".xml"
    End If
    Dim strComprobanteXML As String
    strComprobanteXML = CStr(DOMFacturaXMLsinTimbrar.xml)   'Se convierte a STRING el contenido del XML
    strComprobanteXML = Replace(strComprobanteXML, "<?xml version=""1.0""?>", "") 'Se elimina el encabezado
    strComprobanteXML = Replace(Replace(Trim(strComprobanteXML), Chr(10), ""), Chr(13), "") 'Se eliminan los saltos de linea
    
        Dim DOMelementoSOAP As MSXML2.IXMLDOMElement
    Set DOMelementoSOAP = DOMRequestTimbradoXML.createElement("soap:Envelope")
    
    ' Aquí van los Namespaces
    DOMelementoSOAP.setAttribute "xmlns:soap", "http://schemas.xmlsoap.org/soap/envelope/"
    DOMelementoSOAP.setAttribute "xmlns:tim", strURLXMLNSTimbrado
        
    DOMRequestTimbradoXML.appendChild DOMelementoSOAP
            '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim DOMelementoSOAPBody As MSXML2.IXMLDOMElement
    Set DOMelementoSOAPBody = DOMRequestTimbradoXML.createElement("soap:Body")
        
    DOMelementoSOAP.appendChild DOMelementoSOAPBody
    
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


    Dim DOMelementoRaiz As MSXML2.IXMLDOMElement
    Set DOMelementoRaiz = DOMRequestTimbradoXML.createElement("tim:" & strURLWMTimbrado)
    
    'DOMelementoRaiz.setAttribute "xmlns:tim", strURLXMLNSTimbrado
    
    DOMelementoSOAPBody.appendChild DOMelementoRaiz
    
    'DOMRequestTimbradoXML.appendChild DOMelementoRaiz
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim DOMelementoContrato As MSXML2.IXMLDOMElement
    Set DOMelementoContrato = DOMRequestTimbradoXML.createElement("contrato")

    DOMelementoContrato.Text = strContrato
    DOMelementoRaiz.appendChild DOMelementoContrato
    'DOMelementoRaiz.appendChild DOMelementoTipoDocumento
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim DOMelementoUsuario As MSXML2.IXMLDOMElement
    Set DOMelementoUsuario = DOMRequestTimbradoXML.createElement("usuario")

    DOMelementoUsuario.Text = strUsuario
    DOMelementoRaiz.appendChild DOMelementoUsuario
    'DOMelementoRaiz.appendChild DOMelementoTipoDocumento
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim DOMelementoContra As MSXML2.IXMLDOMElement
    Set DOMelementoContra = DOMRequestTimbradoXML.createElement("passwd")
    
    DOMelementoContra.Text = strPassword
    DOMelementoRaiz.appendChild DOMelementoContra
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim DOMelementoCDATA As MSXML2.IXMLDOMElement
    Set DOMelementoCDATA = DOMRequestTimbradoXML.createElement("cfdiXml")

    DOMelementoCDATA.Text = Replace$(Replace$("<![CDATA[" & strComprobanteXML & "]]>", ">", Chr(62)), "<", Chr(60))
    DOMelementoRaiz.appendChild DOMelementoCDATA
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim DOMelementoOpcion As MSXML2.IXMLDOMElement
    Set DOMelementoOpcion = DOMRequestTimbradoXML.createElement("opciones")

    DOMelementoOpcion.Text = "REGRESAR_CON_ERROR_307_XML"
    DOMelementoRaiz.appendChild DOMelementoOpcion

    vgStrSolicitudEnPlanoProdigia = "<contrato>" & strContrato & "</contrato>"
    vgStrSolicitudEnPlanoProdigia = vgStrSolicitudEnPlanoProdigia & "<usuario>" & strUsuario & "</usuario>"
    vgStrSolicitudEnPlanoProdigia = vgStrSolicitudEnPlanoProdigia & "<passwd>" & strPassword & "</passwd>"
    vgStrSolicitudEnPlanoProdigia = vgStrSolicitudEnPlanoProdigia & Replace$(Replace$("<cfdiXml><![CDATA[" & strComprobanteXML & "]]></cfdiXml>", ">", Chr(62)), "<", Chr(60))
    vgStrSolicitudEnPlanoProdigia = vgStrSolicitudEnPlanoProdigia & "<opciones>REGRESAR_CON_ERROR_307_XML</opciones>"
    vgStrSolicitudEnPlanoProdigia = "<tim:" & strURLWMTimbrado & " xmlns:tim=""" & strURLXMLNSTimbrado & """>" & vgStrSolicitudEnPlanoProdigia & "</tim:" & strURLWMTimbrado & ">"
    vgStrSolicitudEnPlanoProdigia = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/""  xmlns:tim=""" & strURLXMLNSTimbrado & """><soapenv:Body>" & vgStrSolicitudEnPlanoProdigia & "</soapenv:Body></soapenv:Envelope>"
    'Se regresa el valor de la función
    'fnDOMRequestXMLProdigia = solicitudEnPlano
    Set fnDOMRequestXMLProdigia = DOMRequestTimbradoXML

    'Se graba el archivo Request Timbrado en la ruta especificada
    fnDOMRequestXMLProdigia.Save vlstrRutaRequestTimbrado
End Function

''----- Llamar al Web Service de PAX para cancelar el comprobante fiscal -----'
'Public Function fblnCancelarCFDiProdigia(strUUId As String, strRFC As String, strRFCReceptor As String, strTotalComprobante As String, strRutaAcuse As String)
'On Error GoTo NotificaErrorTimbre:
'
'    '-- Variables para el request de la cancelación -''
'    Dim SerializerWS As SoapSerializer30     'Para serializar el XML
'    Dim ConectorWS As ISoapConnector         'Para conectarse al WebService
'    Dim DOMRequestXML As MSXML2.DOMDocument  'XML con los datos para la cancelación
'    Dim rsConexion As New ADODB.Recordset    'Para leer datos de configuración del PAC
'    Dim strURLWSCancelacion As String        'Dirección del Web Service para la cancelación
'    Dim strURLWMCancelacion As String        'Dirección del Web Method para la cancelación
'    Dim strURLXMLNSTimbrado As String        'Dirección del metodo XMLNS
'    Dim strUsuario As String                 'Almacena el usuario del hospital para el uso del WebService
'    Dim strPassword As String                'Almacena la contraseña del hospital para el uso del WebService
'    Dim intEstructura As Integer             'Indica la estructura la cancelación
'
'    '-- Variables para la respuesta del WebService --'
'    Dim ReaderRespuestaWS As SoapReader30    'Para leer la respuesta del WebService
'    Dim DOMResponseXML As MSXML2.DOMDocument 'XML para leer los datos de la respuesta del WebService
'    Dim strCodigoWS As String                'Almacena el código de respuesta del WebService
'    Dim strUUIDWS As String                  'Almacena el código del comprobante regresado por el WebService
'    Dim strDescripcionWS As String           'Almacena el mensaje de respuesta del WebService
'    Dim strFechaWS As String                 'Almacena la fecha de cancelación del WebService
'
'    '-- Variables generales --'
'    Dim strSentencia As String
'    Dim strMensaje As String
'
'    '--Variable para saber si se da la opción de cancelar el CFDi en forma manual, cuando el servicio del SAT no esta disponible
'    Dim blnBitCancelaCDFiNOSAT As Boolean
'    Dim RsbitCancelaCFDiNOSAT As New ADODB.Recordset
'
'    Dim strSignatureWS As String
'    Dim strPruebaCancelacionNE As String
'
'    vlstrMensajeErrorCancelacionCFDi = ""
'    vlintTipoMensajeErrorCancelacionCFDi = 0
'
'    'Se especifican las rutas para la conexión con el servicio de cancelación
'    Set rsConexion = frsEjecuta_SP(CStr(vgintClaveEmpresaContable), "SP_GNSELCONFIGPAC")
'
'    strUsuario = rsConexion!Usuario
'    strPassword = rsConexion!Password
'
'    If vgblnNuevoEsquemaCancelacion Then
'        strURLWSCancelacion = rsConexion!URLWSCancelacionNE   '(Pruebas = "https://test.paxfacturacion.com.mx:476/webservices/wcfCancelaasmx.asmx", Producción = "https://www.paxfacturacion.com.mx:453/webservices/wcfCancelaasmx.asmx")
'        strURLWMCancelacion = rsConexion!URLWMCancelacionNE
'        strURLXMLNSTimbrado = rsConexion!URLXMLNSTimbradoNE   '(Pruebas = "https://test.paxfacturacion.com.mx:476")
'
'        'Se forma el archivo Request de la Cancelación
'        Set DOMRequestXML = fnDOMReqCancelXMLPAX_NE(strUUId, strRFC, strRFCReceptor, strTotalComprobante, strUsuario, strPassword, strURLXMLNSTimbrado, False)
'    Else
'        strURLWSCancelacion = rsConexion!URLWSCancelacion   '(Pruebas = "https://test.paxfacturacion.com.mx:453/webservices/wcfCancelaasmx.asmx", Producción = "https://www.paxfacturacion.com.mx:453/webservices/wcfCancelaasmx.asmx")
'        strURLWMCancelacion = rsConexion!URLWMCancelacion
'        strURLXMLNSTimbrado = rsConexion!URLXMLNSTimbrado   '(Pruebas = "https://test.paxfacturacion.com.mx:453")
'        intEstructura = 0 'Se especifica el tipo de estructura (0 por default)
'
'        'Se especifica la ruta para la creación del archivo Request para la cancelación
'        'vlstrRutaRequestTimbrado = Trim(strRutaAcuse)
'
'        'Se forma el archivo Request de la Cancelación
'        Set DOMRequestXML = fnDOMReqCancelXMLPAX(strUUId, strRFC, intEstructura, strUsuario, strPassword, strURLXMLNSTimbrado, False)
'    End If
'
'
'    ' se carga parametro para saber si en caso de un error se puede permitir la cancelación del documento en el SIHO para después cancelar en SAT
'    strSentencia = "Select vchvalor from SiParametro where VCHNOMBRE = 'BITCANCELACFDINOSAT' and INTCVEEMPRESACONTABLE = " & vgintClaveEmpresaContable
'    Set RsbitCancelaCFDiNOSAT = frsRegresaRs(strSentencia, adLockOptimistic)
'    If RsbitCancelaCFDiNOSAT.RecordCount > 0 Then
'       blnBitCancelaCDFiNOSAT = IIf(IsNull(RsbitCancelaCFDiNOSAT!vchvalor), 0, Val(RsbitCancelaCFDiNOSAT!vchvalor))
'    Else
'       blnBitCancelaCDFiNOSAT = 0
'    End If
'
'    '#####################################################################################################################################
'    '########################################### INICIA CONEXIÓN CON EL SERVICIO DE CANCELACIÓN ##########################################
'    '#####################################################################################################################################
'    Set ConectorWS = New HttpConnector30
'
'    ConectorWS.Property("EndPointURL") = strURLWSCancelacion    'Ruta del WebService
'    ConectorWS.Property("SoapAction") = strURLWMCancelacion     'Ruta del WebMethod para la cancelación ("https://www.paxfacturacion.com.mx/fnCancelarXML")
'    ConectorWS.Connect
'
'
'    ConectorWS.BeginMessage
'        Set SerializerWS = New SoapSerializer30
'        SerializerWS.Init ConectorWS.InputStream
'
'        'Agrega información del request
'        SerializerWS.StartEnvelope
'            '|  Solo si se está trabajando con el nuevo esquema de cancelación se estableme el xmlns, el esquema viejo lo trae en el body
'            If vgblnNuevoEsquemaCancelacion Then SerializerWS.SoapDefaultNamespace strURLXMLNSTimbrado
'            SerializerWS.StartBody
'                SerializerWS.WriteXml DOMRequestXML.xml
'            SerializerWS.EndBody
'        SerializerWS.EndEnvelope
'    ConectorWS.EndMessage
'
'    Set ReaderRespuestaWS = New SoapReader30
'    ReaderRespuestaWS.Load ConectorWS.OutputStream
'
'    '|<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'    '|  SOLO PARA EMULAR RESPUESTAS DEL PAX EN AMBIENTE DE PRUEBAS, DEBERÁ ESTAR COMENTADO EN TODOS LOS EJECUTABLES QUE SE MANDEN AL CLIENTE
'    '|>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'    strPruebaCancelacionNE = fstrPruebaCancelacionNE(strUUId)
'    If strPruebaCancelacionNE <> "" Then ReaderRespuestaWS.Body.Text = strPruebaCancelacionNE
'    '|<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'    '|>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'
'
'    '#######################################################################################################################################
'    '########################################### FINAL DE CONEXIÓN CON EL SERVICIO DE CANCELACIÓN ##########################################
'    '#######################################################################################################################################
'
'   If Not ReaderRespuestaWS.Fault Is Nothing Then
'        If Not ReaderRespuestaWS.FaultDetail Is Nothing Then
'            Dim DOMNodoCodigo As MSXML2.IXMLDOMNode
'            '------------------------------------------------------------------------------------------------ ERROR A NIVEL 1
'            'Se obtiene el codigo del error devuelto por PAX
'            Set DOMNodoCodigo = ReaderRespuestaWS.FaultDetail.childNodes(0).Attributes.getNamedItem("codigo")
'            'Se captura el error
'            If DOMNodoCodigo Is Nothing Then
'                Err.Raise 1000, "Error de cancelación de nivel de capa 1", "Error"
'                fblnCancelarCFDiPAX = False
'
'                frsEjecuta_SP strUUId & "|1||1|'Error de cancelación de nivel de capa 1'", "Sp_PvPendientesCancelarSAT"
'            Else
'                If blnBitCancelaCDFiNOSAT Then
'                   fblnCancelarCFDiPAX = True
'                   frsEjecuta_SP strUUId & "|1|PC|0|'Pendiente de cancelar ante el SAT'", "Sp_PvPendientesCancelarSAT" 'Se agrega
'                Else
'                   strMensaje = "Ocurrió un error al solicitar el servicio de cancelación" & vbNewLine & vbNewLine & _
'                                 "Número de error: " & DOMNodoCodigo.Text & vbNewLine & _
'                                 "Descripción: " & ReaderRespuestaWS.FaultString.Text
'                   vlstrMensajeErrorCancelacionCFDi = strMensaje
'                   vlintTipoMensajeErrorCancelacionCFDi = vbCritical
'                   fblnCancelarCFDiPAX = False
'
'                   frsEjecuta_SP strUUId & "|1||1|'Error: '" & Trim(DOMNodoCodigo.Text) & " Descripción: " & Trim(ReaderRespuestaWS.FaultString.Text) & "'", "Sp_PvPendientesCancelarSAT"
'                End If
'            End If
'        End If
'   Else
'        Set DOMResponseXML = New MSXML2.DOMDocument
'        DOMResponseXML.loadXML (ReaderRespuestaWS.Body.Text)
'        'Se verifica si el WS regresó información...
'        Set DOMNodoCodigo = DOMResponseXML.selectSingleNode("Cancelacion/Folios")
'        If Not DOMNodoCodigo Is Nothing Then
'            strCodigoWS = Trim(DOMNodoCodigo.selectSingleNode("UUIDEstatus").Text)
'            strDescripcionWS = DOMNodoCodigo.selectSingleNode("UUIDdescripcion").Text
'
'            '|  Si se está trabajando con el nuevo esquema de cancelación se realiza la validación con los nuevos códigos
'            If vgblnNuevoEsquemaCancelacion Then
'                Select Case Val(strCodigoWS)
'                    '-------------------------------------------------------------------------------------------------------------------------------
'                    '| Este código está aparte porque tiene dos significados diferentes (facepalm)
'                    '-------------------------------------------------------------------------------------------------------------------------------
'                    Case 201
'                        If UCase(strDescripcionWS) = "201 - COMPROBANTE EN PROCESO DE SER CANCELADO." Then
'                            '| 201 - Comprobante En Proceso de ser Cancelado"
'                            fblnCancelarCFDiPAX = False
'                            '| Elimina el registro en la tabla de comprobantes pendientes de cancelar con el estado PA "Pendiente de autorización"
'                            frsEjecuta_SP strUUId & "|0|PA", "Sp_PvPendientesCancelarSAT" 'Se borra
'                            '| Inserta un registro en la tabla de comprobantes pendientes de cancelar con el estado PA "Pendiente de autorización"
'                            frsEjecuta_SP strUUId & "|1|PA|0|'Pendiente de autorización de cancelación. " & UCase(strDescripcionWS) & "'", "Sp_PvPendientesCancelarSAT" 'Se agrega
'                            strMensaje = "El comprobante se encuentra en proceso de ser cancelado." & vbNewLine & _
'                                         "Folio en espera: " & strUUId & vbNewLine
'                            vlintTipoMensajeErrorCancelacionCFDi = vbCritical
'                        Else
'                            '| 201 - Comprobante Cancelado sin Aceptación.
'                            fblnCancelarCFDiPAX = True
'                            'Se graba el archivo Response de cancelación en la ruta especificada, sólo cuando se cancela la primera vez
'                            If strRutaAcuse <> "" Then
'                                DOMResponseXML.Save strRutaAcuse
'                            End If
'
'                            frsEjecuta_SP strUUId & "|0|PA|0|'" & UCase(strDescripcionWS) & "'", "Sp_PvPendientesCancelarSAT" 'Se borra
'
'                            strMensaje = ""
'                            vlintTipoMensajeErrorCancelacionCFDi = vbInformation
'                        End If
'                    '-------------------------------------------------------------------------------------------------------------------------------
'                    '| Cancelaciones exitosas
'                    '-------------------------------------------------------------------------------------------------------------------------------
'                    Case 202, 107, 103
'                        '| 202 - Comprobante previamente cancelado
'                        '| 107 – El CFDI ha sido Cancelado por Plazo Vencido
'                        '| 103 – El CFDI ha sido Cancelado Previamente por Aceptación del Receptor
'                        fblnCancelarCFDiPAX = True
'                        '| Elimina el registro en la tabla de comprobantes pendientes de cancelar con el estado PA "Pendiente de autorización"
'                        frsEjecuta_SP strUUId & "|0|PA|0|'" & UCase(strDescripcionWS) & "'", "Sp_PvPendientesCancelarSAT" ' Se borra
'                        'Se graba el archivo Response de cancelación en la ruta especificada, sólo cuando se cancela la primera vez
'                        If strRutaAcuse <> "" Then
'                            DOMResponseXML.Save strRutaAcuse
'                        End If
'                        strMensaje = ""
'                        vlintTipoMensajeErrorCancelacionCFDi = vbInformation
'                    '-------------------------------------------------------------------------------------------------------------------------------
'                    '| Cancelaciones que se quedaron en espera
'                    '-------------------------------------------------------------------------------------------------------------------------------
'                    Case 105, 106
'                        '| 105 – El CFDI no se puede Cancelar por que tiene Estatus De “En espera De Aceptación”
'                        '| 106 – El CFDI no se puede Cancelar por que tiene Estatus de “En Proceso”.
'                        fblnCancelarCFDiPAX = False
'                        '| Elimina el registro en la tabla de comprobantes pendientes de cancelar con el estado PA "Pendiente de autorización"
'                        frsEjecuta_SP strUUId & "|0|PA", "Sp_PvPendientesCancelarSAT" ' Se borra
'                        '| Inserta un registro en la tabla de comprobantes pendientes de cancelar con el estado PA "Pendiente de autorización"
'                        frsEjecuta_SP strUUId & "|1|PA|0|'" & UCase(strDescripcionWS) & "'", "Sp_PvPendientesCancelarSAT" ' Se agrega
'                        strMensaje = "El comprobante se encuentra en espera de aceptación." & vbNewLine & _
'                                     "Folio en espera: " & strUUId & vbNewLine
'                        vlintTipoMensajeErrorCancelacionCFDi = vbExclamation
'                    '-------------------------------------------------------------------------------------------------------------------------------
'                    '| Cancelaciones rechazadas por el receptor
'                    '-------------------------------------------------------------------------------------------------------------------------------
'                    Case 104
'                        '| 104 – El CFDI no se puede Cancelar por que fue Rechazado Previamente.
'                        fblnCancelarCFDiPAX = False
'                        '| Elimina el registro en la tabla de comprobantes pendientes de cancelar con el estado PA "Pendiente de autorización"
'                        frsEjecuta_SP strUUId & "|0|PA", "Sp_PvPendientesCancelarSAT" ' Se borra
'                        '| Inserta un registro en la tabla de comprobantes pendientes de cancelar con el estado CR "Cancelación rechazada"
'                        frsEjecuta_SP strUUId & "|1|CR|0|'" & UCase(strDescripcionWS) & "'", "Sp_PvPendientesCancelarSAT" ' Se agrega
'                        strMensaje = "El comprobante no se puede cancelar por que fue rechazado por el receptor." & vbNewLine & _
'                                     "Folio rechazado: " & strUUId & vbNewLine
'                        vlintTipoMensajeErrorCancelacionCFDi = vbCritical
'                    '-------------------------------------------------------------------------------------------------------------------------------
'                    '|  Código de error genérico para las excepciones no controladas del SAT
'                    '-------------------------------------------------------------------------------------------------------------------------------
'                    Case 999
'                        strMensaje = "El servicio del SAT no responde a la petición de cancelación. " & vbNewLine & _
'                                     "Por favor intente más tarde."
'                        If blnBitCancelaCDFiNOSAT Then
'                            fblnCancelarCFDiPAX = True
'                            frsEjecuta_SP strUUId & "|1|PC", "Sp_PvPendientesCancelarSAT" ' Se agrega
'                        Else
'                            fblnCancelarCFDiPAX = False
'                            vlintTipoMensajeErrorCancelacionCFDi = vbCritical
'                        End If
'
'                        frsEjecuta_SP strUUId & "|1||1|'" & UCase(strDescripcionWS) & " - " & Trim(strMensaje) & "'", "Sp_PvPendientesCancelarSAT"
'                    '-------------------------------------------------------------------------------------------------------------------------------
'                    '| Cancelaciones fallidas
'                    '-------------------------------------------------------------------------------------------------------------------------------
'                    Case Else
'                        fblnCancelarCFDiPAX = False
'                        strMensaje = "Ocurrió un error al solicitar la cancelación del comprobante." & vbNewLine & vbNewLine & _
'                                     "Número de error: " & strCodigoWS & vbNewLine & _
'                                     "Descripción: " & strDescripcionWS
'                        vlintTipoMensajeErrorCancelacionCFDi = vbCritical
'
'                        frsEjecuta_SP strUUId & "|1||1|'" & UCase(strDescripcionWS) & "'", "Sp_PvPendientesCancelarSAT"
'                End Select
'                vlstrMensajeErrorCancelacionCFDi = strMensaje
'            Else
'
'                If Val(strCodigoWS) = 201 Or Val(strCodigoWS) = 202 Then
'                    '201 - El folio se ha cancelado con éxito.
'                    '202 - El CFDI ya había sido cancelado previamente.
'                    '-------------------------------------------------------------------------------------------- CANCELACIÓN CORRECTA
'                    fblnCancelarCFDiPAX = True
'                    'Se graba el archivo Response de cancelación en la ruta especificada, sólo cuando se cancela la primera vez
'                    If Val(strCodigoWS) = 201 And strRutaAcuse <> "" Then
'                        DOMResponseXML.Save strRutaAcuse
'                    End If
'
'                    frsEjecuta_SP strUUId & "|1||1|'" & Trim(strDescripcionWS) & "'", "Sp_PvPendientesCancelarSAT"
'                Else
'                    '-------------------------------------------------------------------------------------------- ERROR A NIVEL 2
'                    If Val(strCodigoWS) = 999 Then  'Código de error genérico para las excepciones no controladas del SAT
'                       strDescripcionWS = "El servicio del SAT no responde a la petición de cancelación. " & vbNewLine & _
'                                           "Por favor intente más tarde."
'
'                    End If
'                    If blnBitCancelaCDFiNOSAT Then
'                       fblnCancelarCFDiPAX = True
'                       frsEjecuta_SP strUUId & "|1|PC", "Sp_PvPendientesCancelarSAT" ' Se agrega
'                    Else
'                       strMensaje = "Ocurrió un error al solicitar la cancelación del comprobante." & vbNewLine & vbNewLine & _
'                                    "Número de error: " & strCodigoWS & vbNewLine & _
'                                    "Descripción: " & strDescripcionWS
'                       vlstrMensajeErrorCancelacionCFDi = strMensaje
'                       fblnCancelarCFDiPAX = False
'                    End If
'
'                    frsEjecuta_SP strUUId & "|1||1|'" & Val(strCodigoWS) & " - " & Trim(strDescripcionWS) & "'", "Sp_PvPendientesCancelarSAT"
'                End If
'            End If
'        Else
'            If blnBitCancelaCDFiNOSAT Then
'               fblnCancelarCFDiPAX = True
'               frsEjecuta_SP strUUId & "|1|PC|0|'Pendiente de cancelar ante el SAT'", "Sp_PvPendientesCancelarSAT" ' Se agrega
'            Else
'               vlstrMensajeErrorCancelacionCFDi = "Ocurrió un error al solicitar la cancelación del comprobante." & vbNewLine & vbNewLine & _
'                                                  "No se recibió respuesta del web service."
'               fblnCancelarCFDiPAX = False
'
'               frsEjecuta_SP strUUId & "|1||1|'" & vlstrMensajeErrorCancelacionCFDi & "'", "Sp_PvPendientesCancelarSAT"
'            End If
'        End If
'    End If
'
'Exit Function
'NotificaErrorTimbre:
'    If Err.Number > 0 And Err.Number <> 1001 Then
'       If blnBitCancelaCDFiNOSAT Then
'          fblnCancelarCFDiPAX = True
'          frsEjecuta_SP strUUId & "|1|PC|0|'NotificaErrorTimbre: " & Err.Number & "'", "Sp_PvPendientesCancelarSAT" ' Se agrega
'       Else
'          strMensaje = "Ocurrió un error al solicitar la cancelación del comprobante." & vbNewLine & _
'                        "Número de error: " & Err.Number & vbNewLine & _
'                        "Origen: " & Err.Source & vbNewLine & _
'                        "Descripción: " & Err.Description
'          vlstrMensajeErrorCancelacionCFDi = strMensaje
'          fblnCancelarCFDiPAX = False
'
'          frsEjecuta_SP strUUId & "|1||1|'" & strMensaje & "'", "Sp_PvPendientesCancelarSAT"
'       End If
'    End If
'End Function



