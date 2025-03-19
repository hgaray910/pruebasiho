VERSION 5.00
Begin VB.Form frmDatosCorreoVitamedica 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datos de envío de correo electrónico Vitamédica"
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
   Begin VB.PictureBox sstObj 
      AutoRedraw      =   -1  'True
      Height          =   6375
      Left            =   -45
      ScaleHeight     =   6315
      ScaleWidth      =   6555
      TabIndex        =   10
      Top             =   -75
      Width           =   6615
      Begin VB.Frame Frame4 
         Height          =   750
         Left            =   2900
         TabIndex        =   9
         Top             =   3800
         Width           =   680
         Begin VB.CommandButton cmdEnviar 
            Height          =   550
            Left            =   60
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmDatosCorreoVitamedica.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Enviar correo"
            Top             =   130
            UseMaskColor    =   -1  'True
            Width           =   555
         End
      End
      Begin VB.Frame Frame1 
         Height          =   3675
         Left            =   180
         TabIndex        =   6
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
            TabIndex        =   11
            Top             =   2200
            Width           =   975
            Begin VB.CheckBox chkTXT 
               Height          =   255
               Left            =   390
               TabIndex        =   4
               ToolTipText     =   "Archivo TXT"
               Top             =   330
               Width           =   465
            End
            Begin VB.Label Label5 
               Caption         =   "Archivo de texto"
               Height          =   495
               Left            =   180
               TabIndex        =   15
               Top             =   660
               Width           =   615
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
            TabIndex        =   14
            Top             =   765
            Width           =   210
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Para"
            Height          =   195
            Left            =   285
            TabIndex        =   13
            Top             =   330
            Width           =   330
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Asunto"
            Height          =   195
            Left            =   285
            TabIndex        =   12
            Top             =   1230
            Width           =   495
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Height          =   195
            Left            =   285
            TabIndex        =   8
            Top             =   375
            Width           =   45
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Mensaje"
            Height          =   195
            Left            =   285
            TabIndex        =   7
            Top             =   1680
            Width           =   600
         End
      End
   End
End
Attribute VB_Name = "frmDatosCorreoVitamedica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------------------
'| Nombre del Proyecto      : Caja
'| Nombre del Formulario    : frmDatosCorreoVitamedica
'-----------------------------------------------------------------------------------------
'| Objetivo: Agregar la pantalla de los datos de envío de correo electrónico para CFD/CFDi
'-----------------------------------------------------------------------------------------
'| Análisis y Diseño            : Teresita de J. Zubía Ramos
'| Autor                        : Teresita de J. Zubía Ramos
'| Fecha de Creación            : 15/Mayo/2020
'| Modificó                     :
'| Fecha última modificación    :
'-----------------------------------------------------------------------------------------

Option Explicit

Dim rsCorreo As New ADODB.Recordset
Dim temp As Integer                         'Variable auxiliar para el focus del txtMensaje
Dim vlblnEnvio As Boolean                   'Variable auxiliar para saber si se envió el correo
Dim vlblnNuevo As Boolean                   'Variable auxiliar para determinar si la pantalla se carga por primera vez



Public strCorreoDestinatario As String      'Correo del destinatario cargado automáticamente de la tabla correspondiente
Public strAsunto As String                  'Asunto del correo electrónico
Public strMensaje As String                 'Mensaje del correo electrónico
Public strRutaTXT As String                 'Ruta donde se almacenan los archivos TXT para Vitamédica
Public strNombreArchivoTXT As String        'Nombre del archivo TXT

Public lngEmpleado As Long                  'Clave del empleado que envía el correo
Public blnArchivoTXT As Boolean             'Indica si se anexará un archivo TXT

Private clsEmail As clsCDOmail              'Variable de tipo clase con los parámetros para el envío de correo



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
    
    If chkTXT.Value = vbUnchecked Then
        'Debe de seleccionar al menos un documento adjunto.
        MsgBox SIHOMsg(1197), vbCritical, "Mensaje"
        fblnValidaCampos = False
        chkTXT.SetFocus
        Exit Function
    End If
    
    If Trim(strRutaTXT) = "" Then
        'No se encontró la ruta del archivo de texto.
        MsgBox SIHOMsg(1198), vbCritical, "Mensaje"
        fblnValidaCampos = False
        chkTXT.SetFocus
        Exit Function
    End If
   
Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ": fblnValidaCampos"))
End Function



Public Sub cmdEnviar_Click()
On Error GoTo NotificaError

Dim strArchivoTXT As String
Dim vlPassDecoded As String

    If fblnValidaCampos Then
    
        Set rsCorreo = frsEjecuta_SP(CStr(vgintClaveEmpresaContable) & "|0", "Sp_CnSelCnCorreo")
        
        strArchivoTXT = Trim(strRutaTXT)
        
        Set clsEmail = New clsCDOmail
        
        'Se decodifica la contraseña
        vlPassDecoded = Trim(rsCorreo!vchPassword)
        
        'Se reemplazan los caracteres especiales ("U"<-"?" "l"<-"ñ" "="<-"Ñ" "=="<-"Ñ?")
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
                .Asunto = "Envío de la interfaz de Vitamedica"
            Else
                .Asunto = Trim(txtAsunto.Text) '& " (" & Trim(vlstrTipoDocumento) & " " & Trim(strFolioDocumento) & ")"
            End If
            .AdjuntoPDF = IIf(chkTXT.Value = vbChecked, Trim(strArchivoTXT), "")
            
            .De = Trim(rsCorreo!vchNombre) & " <" & Trim(rsCorreo!vchCorreo) & ">"
            .Para = Trim(txtPara.Text)
            .CC = Trim(txtCC.Text)
            .mensaje = Trim(txtMensaje.Text) & Chr(13) & Chr(13) & "DOCUMENTO ENVIADO: " & Trim(strNombreArchivoTXT)
            
            ' Enviar el correo
            If .fblnEnviarCorreo = True Then
                ' Si se envía correctamente, grabar en el log de correos (SILOGCORREOS)
                Call pGuardarLogCorreos(.Usuario, .Para, .CC, .Asunto, .AdjuntoPDF, .AdjuntoXML, .mensaje, lngEmpleado)
                
                ' Guardar el log de transacciones
                Call pGuardarLogTransaccion(Me.Name, EnmGrabar, lngEmpleado, Me.Caption, Trim(strNombreArchivoTXT))
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

Private Sub Form_Load()
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
            txtAsunto = IIf(IsNull(Trim(strAsunto)), "", Trim(strAsunto))
            txtMensaje = IIf(IsNull(Trim(strMensaje)), "", Trim(strMensaje))
            chkTXT.Value = 1
            'chkXML.Value = IIf(IsNull(!BITXML), vbUnchecked, IIf(!BITXML = 1, vbChecked, vbUnchecked))
        End If
    End With
    
    '- Si se envía un archivo comprimido desde Paquete de Cobranza -'
'    If blnArchivoZIP Then
        'chkPDF.Value = vbUnchecked
        'chkXML.Value = vbUnchecked
        'chkPDF.Enabled = False
        'chkXML.Enabled = False
'        fraAdjuntos.Enabled = False
'    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
    Unload Me
End Sub



Private Sub Form_Unload(Cancel As Integer)
    If vlblnEnvio <> True Then
        ' ¿Desea abandonar la operación?
        If MsgBox(SIHOMsg(17), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
            blnArchivoTXT = False
            Unload Me
        Else
            Cancel = True
        End If
    Else
        blnArchivoTXT = True 'Generación y envío del archivo TXT exitosa
    End If
End Sub



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

'Limpiar variables públicas para que no se queden valores anteriores
Private Sub pLimpiaValores()
    'strCorreoDestinatario = ""
    strRutaTXT = ""
    lngEmpleado = 0
    strMensaje = ""
End Sub
