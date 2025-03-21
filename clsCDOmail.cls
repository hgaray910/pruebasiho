VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCDOmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' para la conexi�n a internet
Private Declare Function InternetGetConnectedState _
    Lib "wininet.dll" ( _
    ByRef lpdwFlags As Long, _
    ByVal dwReserved As Long) As Long

Private Const INTERNET_CONNECTION_MODEM_BUSY As Long = &H8
Private Const INTERNET_RAS_INSTALLED As Long = &H10
Private Const INTERNET_CONNECTION_OFFLINE As Long = &H20
Private Const INTERNET_CONNECTION_CONFIGURED As Long = &H40

' variables locales
Private mServidor As String
Private mPara As String
Private mDe As String
Private mCC As String
Private mAsunto As String
Private mMensaje As String
Private mAdjuntoXML As String
Private mAdjuntoPDF As String
Private mAdjuntoZIP As String
Private mPuerto As Variant
Private mUsuario As String
Private mPassword As String
Private mUseAuntentificacion As Boolean
Private mSSL As Boolean

Function fblnEnviarCorreo() As Boolean
On Error GoTo NotificaError:
    
    ' Variable de objeto CDO.Message
    Dim oCDO As Object
              
    ' Inicializa en error
    fblnEnviarCorreo = False
          
    ' Verifica si hay conexi�n a Internet
    If InternetGetConnectedState(0&, 0&) = False Then
        Err.Raise -10000
    End If
    
    ' Verifica nuevamente el puerto
    If Not IsNumeric(Puerto) Then
        Err.Raise -20000
    End If
    
    ' Crea un Nuevo objeto CDO.Message
    Set oCDO = CreateObject("CDO.Message")
    oCDO.BodyPart.Charset = "utf-8"
    
    ' Indica el servidor Smtp para poder enviar el Mail ( puede ser el nombre del servidor o su direcci�n IP )
    oCDO.Configuration.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = mServidor
    oCDO.Configuration.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    
    ' Puerto. Por defecto se usa el puerto 25, en el caso de Gmail se usa el puerto 465
    oCDO.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = mPuerto

    ' Indica el tipo de autentificaci�n con el servidor de correo.  (0 = no requiere autentificarse, 1= requiere autentificaci�n)
    oCDO.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = Abs(mUseAuntentificacion)
    
    ' Tiempo m�ximo de espera en segundos para la conexi�n
    oCDO.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 10

    ' Configura las opciones para el login en el SMTP
    If mUseAuntentificacion Then
        ' Id de usuario del servidor Smtp ( en el caso de gmail, debe ser la direcci�n de correro mas el @gmail.com )
        oCDO.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = mUsuario
    
        ' Password de la cuenta
        oCDO.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = mPassword
    
        ' Indica si se usa SSL para el env�o. En el caso de Gmail requiere que est� en True
        oCDO.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = mSSL
    End If
    
    ' ********************************** Informaci�n de los datos de env�o **********************************
    
    ' Direcci�n del Destinatario
    oCDO.To = mPara
    
    ' Direcci�n de copia (CC)
    oCDO.CC = mCC
    
    ' Direcci�n del remitente
    oCDO.From = mDe
    
    ' Asunto del mensaje
    oCDO.Subject = mAsunto
    
    ' Cuerpo del mensaje
    oCDO.TextBody = mMensaje
    
    'Ruta del archivo adjunto XML
    If mAdjuntoXML <> "" Then
        If Len(Dir(mAdjuntoXML)) = 0 Then
            ' La ruta del archivo XML no es v�lida
            Err.Raise -30000
        Else
            ' Adjunta el archivo
            oCDO.AddAttachment (mAdjuntoXML)
        End If
    End If
    
    'Ruta del archivo adjunto PDF
    If mAdjuntoPDF <> "" Then
        If Len(Dir(mAdjuntoPDF)) = 0 Then
            ' La ruta del archivo PDF no es v�lida
            Err.Raise -40000
        Else
            ' Adjunta el archivo
            oCDO.AddAttachment (mAdjuntoPDF)
        End If
    End If
    
    'Ruta del archivo adjunto ZIP (archivo comprimido para Paquete de Cobranza)
    If mAdjuntoZIP <> "" Then
        If Len(Dir(mAdjuntoZIP)) = 0 Then
            ' La ruta del archivo ZIP no es v�lida
            Err.Raise -50000
        Else
            ' Adjunta el archivo
            oCDO.AddAttachment (mAdjuntoZIP)
        End If
    End If
    
    ' Actualiza los datos antes de enviar
    oCDO.Configuration.Fields.Update
    
    ' ********************************** Env�o del correo **********************************
'    On Error Resume Next
    Screen.MousePointer = vbHourglass
    
    ' Enviar!!!!!
    oCDO.send
    
    ' Poner el cursor en modo de espera
    Screen.MousePointer = 0
    
    ' Si no hay error, establer el resultado de la funci�n = True
    If vgblnAutomatico = False Then
        If Err.Number = 0 Then
            'El env�o se realiz� satisfactoriamente.
            MsgBox SIHOMsg(1188), vbInformation, "Mensaje"
            fblnEnviarCorreo = True
        End If
    Else
        If Err.Number = 0 Then
            fblnEnviarCorreo = True
        End If
    End If

    ' Descarga la referencia
    If Not oCDO Is Nothing Then
        Set oCDO = Nothing
    End If
    
    'Limpia los c�digos de error
    Err.Clear
    
    'Reestablece el cursor
    Screen.MousePointer = vbNormal
    
Exit Function
NotificaError:
    'Reestablece el cursor
    Screen.MousePointer = vbNormal
    
    'Establece el resultado de la funci�n = False
    fblnEnviarCorreo = False
    
    'Validaci�n de mensajes de error
'    If Err.Number = -2147220973 Then
'        MsgBox "Error al enviar correo: " & vbNewLine & SIHOMsg(1189), vbCritical, "Mensaje"
'    ElseIf Err.Number = -2147220975 Then
'        MsgBox "Error al enviar correo: " & vbNewLine & SIHOMsg(1190), vbCritical, "Mensaje"

    If Err.Number = -10000 Then
        'Verifique que la conexi�n a Internet est� disponible.
        MsgBox "Error al enviar correo: " & vbNewLine & SIHOMsg(1189), vbCritical, "Mensaje"
'    ElseIf Err.Number = -20000 Then
'        'El puerto de conexi�n no es v�lido.
'        MsgBox "Error al enviar correo: " & vbNewLine & 'El puerto de conexi�n no es v�lido.', vbCritical, "Mensaje"
    ElseIf Err.Number = -30000 Then
        'Verifique la ruta del archivo XML.
        MsgBox "Error al enviar correo: " & vbNewLine & SIHOMsg(1190), vbCritical, "Mensaje"
    ElseIf Err.Number = -40000 Then
        'Verifique la ruta del archivo PDF
        MsgBox "Error al enviar correo: " & vbNewLine & SIHOMsg(1191), vbCritical, "Mensaje"
    ElseIf Err.Number = -50000 Then
        'Verifique la ruta del archivo ZIP.
        MsgBox "Error al enviar correo: " & vbNewLine & SIHOMsg(1192), vbCritical, "Mensaje"
    Else
        'No se pudo enviar el correo, revise la configuraci�n para env�o de correo electr�nico y/o el correo del destinatario.
        MsgBox SIHOMsg(1193), vbCritical, "Mensaje"
    End If
End Function

Property Get Servidor() As String
    Servidor = mServidor
End Property

Property Let Servidor(Value As String)
    mServidor = Value
End Property

Property Get Para() As String
    Para = mPara
End Property

Property Get CC() As String
    CC = mCC
End Property

Property Let CC(Value As String)
    mCC = Value
End Property

Property Let Para(Value As String)
    mPara = Value
End Property

Property Get De() As String
    De = mDe
End Property

Property Let De(Value As String)
    mDe = Value
End Property

Property Get Asunto() As String
    Asunto = mAsunto
End Property

Property Let Asunto(Value As String)
    mAsunto = Value
End Property

Property Get Mensaje() As String
    Mensaje = mMensaje
End Property

Property Let Mensaje(Value As String)
    mMensaje = Value
End Property

Property Get AdjuntoPDF() As String
    AdjuntoPDF = mAdjuntoPDF
End Property

Property Let AdjuntoPDF(Value As String)
    mAdjuntoPDF = Value
End Property

Property Get AdjuntoXML() As String
    AdjuntoXML = mAdjuntoXML
End Property

Property Let AdjuntoXML(Value As String)
    mAdjuntoXML = Value
End Property

Property Get AdjuntoZIP() As String
    AdjuntoZIP = mAdjuntoZIP
End Property

Property Let AdjuntoZIP(Value As String)
    mAdjuntoZIP = Value
End Property

Property Get Puerto() As Variant
    Puerto = mPuerto
End Property

Property Let Puerto(Value As Variant)
    mPuerto = Value
End Property

Property Get Usuario() As String
    Usuario = mUsuario
End Property

Property Let Usuario(Value As String)
    mUsuario = Value
End Property

Property Get Password() As String
    Password = mPassword
End Property

Property Let Password(Value As String)
    mPassword = Value
End Property

Property Get UseAuntentificacion() As Boolean
    UseAuntentificacion = mUseAuntentificacion
End Property

Property Let UseAuntentificacion(Value As Boolean)
    mUseAuntentificacion = Value
End Property

Property Get SSL() As Boolean
    SSL = mSSL
End Property

Property Let SSL(Value As Boolean)
    mSSL = Value
End Property
