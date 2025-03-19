VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmDatosCorreo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datos de envío de correo electrónico"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6480
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   6480
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab sstObj 
      Height          =   6375
      Left            =   -45
      TabIndex        =   11
      Top             =   -75
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   11245
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   2
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmDatosCorreo.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin VB.Frame Frame4 
         Height          =   750
         Left            =   2900
         TabIndex        =   10
         Top             =   3800
         Width           =   680
         Begin VB.CommandButton cmdEnviar 
            Height          =   550
            Left            =   60
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmDatosCorreo.frx":001C
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Enviar correo"
            Top             =   130
            UseMaskColor    =   -1  'True
            Width           =   555
         End
      End
      Begin VB.Frame Frame1 
         Height          =   3675
         Left            =   180
         TabIndex        =   7
         Top             =   120
         Width           =   6220
         Begin VB.TextBox txtCC 
            Height          =   315
            Left            =   1485
            MaxLength       =   300
            TabIndex        =   1
            ToolTipText     =   "Dirección de correo a la que se copiará"
            Top             =   765
            Width           =   4470
         End
         Begin VB.TextBox txtAsunto 
            Height          =   315
            Left            =   1485
            MaxLength       =   100
            TabIndex        =   2
            ToolTipText     =   "Asunto del correo electrónico"
            Top             =   1230
            Width           =   4470
         End
         Begin VB.Frame fraAdjuntos 
            Caption         =   "Adjuntar"
            Height          =   1215
            Left            =   285
            TabIndex        =   12
            Top             =   2200
            Width           =   975
            Begin VB.CheckBox chkXML 
               Caption         =   "XML"
               Height          =   255
               Left            =   180
               TabIndex        =   5
               ToolTipText     =   "Archivo XML"
               Top             =   720
               Width           =   720
            End
            Begin VB.CheckBox chkPDF 
               Caption         =   "PDF"
               Height          =   255
               Left            =   180
               TabIndex        =   4
               ToolTipText     =   "Archivo PDF"
               Top             =   300
               Width           =   735
            End
         End
         Begin VB.TextBox txtPara 
            Height          =   315
            Left            =   1485
            MaxLength       =   100
            TabIndex        =   0
            ToolTipText     =   "Dirección de correo electrónico del destinatario"
            Top             =   330
            Width           =   4470
         End
         Begin VB.TextBox txtMensaje 
            Height          =   1755
            Left            =   1485
            MaxLength       =   250
            MultiLine       =   -1  'True
            TabIndex        =   3
            ToolTipText     =   "Mensaje del correo electrónico"
            Top             =   1680
            Width           =   4470
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "CC"
            Height          =   195
            Left            =   285
            TabIndex        =   15
            Top             =   765
            Width           =   210
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Para"
            Height          =   195
            Left            =   285
            TabIndex        =   14
            Top             =   330
            Width           =   330
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Asunto"
            Height          =   195
            Left            =   285
            TabIndex        =   13
            Top             =   1230
            Width           =   495
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Height          =   195
            Left            =   285
            TabIndex        =   9
            Top             =   375
            Width           =   45
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Mensaje"
            Height          =   195
            Left            =   285
            TabIndex        =   8
            Top             =   1680
            Width           =   600
         End
      End
   End
End
Attribute VB_Name = "frmDatosCorreo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------------------
'| Nombre del Proyecto      : Caja
'| Nombre del Formulario    : frmDatosCorreo
'-----------------------------------------------------------------------------------------
'| Objetivo: Agregar la pantalla de los datos de envío de correo electrónico para CFD/CFDi
'-----------------------------------------------------------------------------------------
'| Análisis y Diseño            : Fernando Martínez
'| Autor                        : Fernando Martínez
'| Fecha de Creación            : 29/Dic/2011
'| Modificó                     : Claudia Ruvalcaba
'| Fecha última modificación    : 31/Mayo/2013
'-----------------------------------------------------------------------------------------

Option Explicit

Dim rsCorreo As New ADODB.Recordset
Dim temp As Integer                 'Variable auxiliar para el focus del txtMensaje
Public vlblnEnvio As Boolean           'Variable auxiliar para saber si se envió el correo
Dim vlblnNuevo As Boolean           'Variable auxiliar para determinar si la pantalla se carga por primera vez

Public strTipoDocumento As String       'Tipo de documento FA = Factura, CR = Nota de crédito, CA = Nota de cargo, DO = Donativo
Public strCorreoDestinatario As String  'Correo del destinatario cargado automáticamente del catálogo o tabla correspondiente
Public strFolioDocumento As String      'Folio del documento que se enviará por correo
Public strRutaPDF As String             'Ruta donde se almacenan los archivos PDF
Public strRutaXML As String             'Ruta donde se almacenan los archivos XML
Public strRutaZIP As String             'Ruta donde se almacena el archivo comprimido en formato ZIP para paquete de cobranza
Public lngEmpleado As Long              'Clave del empleado que envía el correo
Public blnArchivoZIP As Boolean         'Indica si se anexará un archivo comprimido ZIP en lugar de archivos PDF y XML individuales
Public strMensaje As String             'Mensaje adicional a agregarse al final del configurado
Public lngIdDocumento As Long           'Clave del documento, dependiendo del valor de strTipoDocumento
Public strEnvioMasivo As Boolean

Private clsEmail As clsCDOmail 'Variable de tipo clase con los parámetros para el envío de correo

Public Sub pGuardarLogCorreos(strEmisor As String, strReceptor As String, strCC As String, strAsunto As String, strPDF As String, strXML As String, strMensaje As String, lngEmpleado As Long)
On Error GoTo NotificaError
   
    Dim rsGenerarLog As New ADODB.Recordset
    Dim vlstrsql As String
    
    vlstrsql = "INSERT INTO SiLogCorreos (VCHEMISOR, VCHRECEPTOR, VCHCC, VCHASUNTO, VCHPDF, VCHXML, VCHMENSAJE, INTIDEMPLEADO, DTMFECHAHORA ) " & _
               " VALUES ('" & Trim(strEmisor) & "', '" & Trim(strReceptor) & "', '" & Trim(strCC) & "', '" & Trim(strAsunto) & "', '" & Trim(strPDF) & "', '" & Trim(strXML) & "', '" & strMensaje & "', '" & lngEmpleado & "', GetDate())"
    EntornoSIHO.ConeccionSIHO.Execute vlstrsql
   
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ": pGuardarLogCorreos"))
End Sub

Private Function fblnValidaCampos() As Boolean
On Error GoTo NotificaError

    fblnValidaCampos = True
    
    If Trim(txtPara) = "" Then
        'Debe indicar un destinatario.
        MsgBox SIHOMsg(1194), vbCritical, "Mensaje"
        fblnValidaCampos = False
        pEnfocaTextBox txtPara
        Exit Function
    End If
    
    If InStr(Trim(txtPara.Text), "@") <= 0 Then
        'La dirección del correo destinatario no es válida.
        MsgBox SIHOMsg(1195), vbCritical, "Mensaje"
        fblnValidaCampos = False
        pEnfocaTextBox txtPara
        Exit Function
    End If
    
    If Trim(txtCC.Text) <> "" And InStr(Trim(txtCC.Text), "@") <= 0 Then
        'La dirección del correo no es válida.
        MsgBox SIHOMsg(1196), vbCritical, "Mensaje"
        fblnValidaCampos = False
        pEnfocaTextBox txtCC
        Exit Function
    End If
    
    If Not blnArchivoZIP And chkPDF.Value = vbUnchecked And chkXML.Value = vbUnchecked Then
        'Debe de seleccionar al menos un documento adjunto.
        MsgBox SIHOMsg(1197), vbCritical, "Mensaje"
        fblnValidaCampos = False
        chkPDF.SetFocus
        Exit Function
    End If
    
    If blnArchivoZIP And Trim(strRutaZIP) = "" Then
        'No se encontró la ruta del archivo comprimido.
        MsgBox SIHOMsg(1198), vbCritical, "Mensaje"
        fblnValidaCampos = False
        chkPDF.SetFocus
        Exit Function
    End If
   
Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ": fblnValidaCampos"))
End Function

Public Sub cmdEnviar_Click()
On Error GoTo NotificaError

    Dim strArchivoPDF As String
    Dim strArchivoXML As String
    Dim vlstrTipoDocumento As String
    Dim vlPassDecoded As String
    
    'If lngEmpleado = 0 Then
        'lngEmpleado = flngPersonaGraba(vgintNumeroDepartamento)
        'If lngEmpleado = 0 Then Exit Sub
    'End If
    
    If fblnValidaCampos Then
        If strTipoDocumento = "FA" Then
            vlstrTipoDocumento = "FACTURA"
        ElseIf strTipoDocumento = "NA" Or strTipoDocumento = "CA" Then
            vlstrTipoDocumento = "NOTA DE CARGO"
        ElseIf strTipoDocumento = "NC" Or strTipoDocumento = "CR" Then
            vlstrTipoDocumento = "NOTA DE CRÉDITO"
        ElseIf strTipoDocumento = "DO" Then
            vlstrTipoDocumento = "DONATIVO"
        ElseIf strTipoDocumento = "PA" Then '<-- Agregado para paquete de Cobranza
            vlstrTipoDocumento = "PAQUETE"
        End If
            
        'En caso que sea un envío desde la pantalla de consulta de CFD/CFDi se toma el asunto de la pantalla
        If vlstrTipoDocumento = "" Then
            If frmComprobanteFiscalDigital.Visible = True Then
                vlstrTipoDocumento = Trim(frmComprobanteFiscalDigital.lblComprobante.Caption)
            ElseIf frmComprobanteFiscalDigitalInternet.Visible = True Then
                vlstrTipoDocumento = Trim(frmComprobanteFiscalDigitalInternet.lblComprobante.Caption)
            End If
        End If
        
        If strFolioDocumento = "" Then
            If frmComprobanteFiscalDigital.Visible = True Then
                strFolioDocumento = Trim(frmComprobanteFiscalDigital.TxtFolio.Text)
            ElseIf frmComprobanteFiscalDigitalInternet.Visible = True Then
                strFolioDocumento = Trim(frmComprobanteFiscalDigitalInternet.TxtFolio.Text)
            End If
        End If
        
        'actualizar correo del paciente
        If strTipoDocumento = "FA" Then pActualizaCorreo
        
        Set rsCorreo = frsEjecuta_SP(CStr(vgintClaveEmpresaContable) & "|0", "Sp_CnSelCnCorreo")
            
        strArchivoPDF = Trim(strRutaPDF) & "\" & Trim(strFolioDocumento) & ".pdf"
        strArchivoXML = Trim(strRutaXML) & "\" & Trim(strFolioDocumento) & ".xml"
        
        Set clsEmail = New clsCDOmail
        
        'Se decodifica la contraseña
        vlPassDecoded = Trim(rsCorreo!vchPassword)
        
        'Se reemplazan los caracteres especiales ("U"<-"?"   "l"<-"ñ"   "="<-"Ñ"   "=="<-"Ñ?")
        vlPassDecoded = Replace(vlPassDecoded, "Ñ?", "==")
        vlPassDecoded = Replace(vlPassDecoded, "Ñ", "=")
        vlPassDecoded = Replace(vlPassDecoded, "ñ", "l")
        vlPassDecoded = Replace(vlPassDecoded, "?", "U")
        
        vlPassDecoded = Decode(vlPassDecoded) 'Se decodifica la contraseña a partir de vlPassDecoded (1a vez)
        vlPassDecoded = Decode(vlPassDecoded) 'Se decodifica la contraseña a partir de vlPassDecoded (2a vez)
        vlPassDecoded = Decode(vlPassDecoded) 'Se decodifica la contraseña a partir de vlPassDecoded (3a vez)
                        
        With clsEmail
            'Datos para enviar
            .Servidor = Trim(rsCorreo!VCHSERVIDORSMTP)
            .Puerto = Val(rsCorreo!intPuerto)
            .UseAuntentificacion = True
            .SSL = IIf(rsCorreo!BITSSL = 1, True, False)
            .Usuario = Trim(rsCorreo!vchCorreo)
            .Password = vlPassDecoded
            If Trim(txtAsunto) = "" Then
'                .Asunto = "(" & Trim(vlstrTipoDocumento) & " " & Trim(strFolioDocumento) & ")"
                .Asunto = "Envío de Comprobante Fiscal Digital"
            Else
                .Asunto = Trim(txtAsunto.Text) '& " (" & Trim(vlstrTipoDocumento) & " " & Trim(strFolioDocumento) & ")"
            End If
            .AdjuntoPDF = IIf(chkPDF.Value = vbChecked, Trim(strArchivoPDF), "")
            .AdjuntoXML = IIf(chkXML.Value = vbChecked, Trim(strArchivoXML), "")
            .AdjuntoZIP = IIf(blnArchivoZIP, Trim(strRutaZIP), "")
            .De = Trim(rsCorreo!vchNombre) & " <" & Trim(rsCorreo!vchCorreo) & ">"
            .Para = Trim(txtPara.Text)
            .CC = Trim(txtCC.Text)
            .mensaje = Trim(txtMensaje.Text) & Chr(13) & Chr(13) & "DOCUMENTO ENVIADO: " & Trim(vlstrTipoDocumento) & " " & Trim(strFolioDocumento) & strMensaje
            'Enviar el correo
            If .fblnEnviarCorreo = True Then
                'Si se envía correctamente, grabar en el log de correos (SILOGCORREOS)
                Call pGuardarLogCorreos(.Usuario, .Para, .CC, .Asunto, .AdjuntoPDF, .AdjuntoXML, .mensaje, lngEmpleado)
                
                'Guardar el log de transacciones
                Call pGuardarLogTransaccion(Me.Name, EnmGrabar, lngEmpleado, Me.Caption, Trim(strFolioDocumento))
                vgblnEnvioExitosoCorreo = True
            End If
            
        End With
        Set clsEmail = Nothing 'Se libera el objeto para el envío del correo
        
        pLimpiaValores 'Se limpian los valores de las variables públicas
        
        vlblnEnvio = True 'Se establece que ya terminó el proceso de envío
        Unload Me
    End If
            
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ": cmdEnviar_Click"))
End Sub

Private Sub pActualizaCorreo()
On Error GoTo NotificaError
Dim rs As New ADODB.Recordset

    Set rs = frsRegresaRs("select intMovPaciente, chrTipoPaciente, chrTipoFactura from pvFactura where intConsecutivo = " & CStr(lngIdDocumento))
    If rs.RecordCount <> 0 Then
        If rs!chrTipoFactura = "P" Then
            pEjecutaSentencia "update pvDatosFiscales set vchCorreoElectronico = '" & Trim(txtPara.Text) & "' where intNumCuenta=" & rs!INTMOVPACIENTE & " and chrTipoPaciente='" & rs!CHRTIPOPACIENTE & "'"
        End If
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pActualizaCorreo"))
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError

    If KeyAscii = 27 Then
        Unload Me
    Else
        If KeyAscii = 13 And temp <> 1 Then
            SendKeys vbTab
        End If
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyPress"))
    Unload Me
End Sub

Public Sub Form_Load()
On Error GoTo NotificaError

    Dim rsCorreo As New ADODB.Recordset

    Me.Icon = frmMenuPrincipal.Icon

    vlblnEnvio = False
    vlblnNuevo = True
    fraAdjuntos.Enabled = True
    
    Set rsCorreo = frsEjecuta_SP(CStr(vgintClaveEmpresaContable) & "|0", "Sp_CnSelCnCorreo")
    With rsCorreo
        If .RecordCount <> 0 Then
            txtPara = Trim(strCorreoDestinatario)
            txtCC = ""
            txtAsunto = IIf(IsNull(Trim(!vchasunto)), "", Trim(!vchasunto))
            txtMensaje = IIf(IsNull(Trim(!vchmensaje)), "", Trim(!vchmensaje))
            chkPDF.Value = IIf(IsNull(!BITPDF), vbUnchecked, IIf(!BITPDF = 1, vbChecked, vbUnchecked))
            chkXML.Value = IIf(IsNull(!BITXML), vbUnchecked, IIf(!BITXML = 1, vbChecked, vbUnchecked))
        End If
    End With
    
    '- Si se envía un archivo comprimido desde Paquete de Cobranza -'
    If blnArchivoZIP Then
        chkPDF.Value = vbUnchecked
        chkXML.Value = vbUnchecked
        chkPDF.Enabled = False
        chkXML.Enabled = False
        fraAdjuntos.Enabled = False
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If vlblnEnvio <> True Then
        ' ¿Desea abandonar la operación?
        If MsgBox(SIHOMsg(17), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
            blnArchivoZIP = False
            Unload Me
        Else
            Cancel = True
        End If
    End If
End Sub



Private Sub txtAsunto_GotFocus()
    pEnfocaTextBox txtAsunto
End Sub

Private Sub txtCC_GotFocus()
    pEnfocaTextBox txtCC
End Sub

Private Sub txtMensaje_GotFocus()
    temp = 1
End Sub

Private Sub txtMensaje_LostFocus()
    temp = 0
End Sub

Private Sub txtPara_GotFocus()
    If Trim(txtPara.Text) <> "" And vlblnNuevo = True Then
        cmdEnviar.SetFocus
    Else
        pEnfocaTextBox txtPara
    End If
    
    vlblnNuevo = False
End Sub

'- Limpiar variables públicas para que no se queden valores anteriores -'
Private Sub pLimpiaValores()
    strTipoDocumento = ""
    lngIdDocumento = 0
    strCorreoDestinatario = ""
    strFolioDocumento = ""
    strRutaPDF = ""
    strRutaXML = ""
    strRutaZIP = ""
    lngEmpleado = 0
    strMensaje = ""
End Sub
