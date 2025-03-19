VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmComprobanteFiscalDigitalInternet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Comprobante fiscal digital a través de Internet"
   ClientHeight    =   9495
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8775
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9495
   ScaleWidth      =   8775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cdlgDestino 
      Left            =   120
      Top             =   8600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame freDatosFiscales 
      Caption         =   "Datos fiscales"
      Enabled         =   0   'False
      Height          =   2350
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   8535
      Begin VB.TextBox txtRegimen 
         Height          =   315
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   1920
         Width           =   6255
      End
      Begin VB.TextBox txtCiudad 
         Height          =   315
         Left            =   6150
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   1605
         Width           =   2145
      End
      Begin VB.TextBox txtNumeroInterior 
         Height          =   315
         Left            =   4410
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1290
         Width           =   855
      End
      Begin VB.TextBox txtNumeroExterior 
         Height          =   315
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1290
         Width           =   1095
      End
      Begin VB.TextBox txtNombre 
         Height          =   315
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   660
         Width           =   6255
      End
      Begin VB.TextBox txtCP 
         Height          =   315
         Left            =   6150
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1290
         Width           =   2145
      End
      Begin VB.TextBox txtColonia 
         Height          =   315
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1605
         Width           =   3225
      End
      Begin VB.TextBox txtRFC 
         Height          =   315
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   350
         Width           =   1545
      End
      Begin VB.TextBox txtCalle 
         Height          =   315
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   975
         Width           =   6255
      End
      Begin VB.Label Label17 
         Caption         =   "Régimen fiscal"
         Height          =   195
         Left            =   240
         TabIndex        =   41
         Top             =   1950
         Width           =   1200
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Número exterior"
         Height          =   195
         Left            =   240
         TabIndex        =   30
         Top             =   1352
         Width           =   1110
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Razón social"
         Height          =   195
         Left            =   240
         TabIndex        =   29
         Top             =   724
         Width           =   915
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Ciudad"
         Height          =   195
         Left            =   5400
         TabIndex        =   28
         Top             =   1665
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Colonia"
         Height          =   195
         Left            =   240
         TabIndex        =   27
         Top             =   1666
         Width           =   525
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Calle"
         Height          =   195
         Left            =   240
         TabIndex        =   26
         Top             =   1038
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "RFC"
         Height          =   195
         Left            =   240
         TabIndex        =   25
         Top             =   410
         Width           =   315
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "CP"
         Height          =   195
         Left            =   5400
         TabIndex        =   15
         Top             =   1320
         Width           =   210
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Número interior"
         Height          =   195
         Left            =   3240
         TabIndex        =   14
         Top             =   1320
         Width           =   1065
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del comprobante"
      Height          =   5895
      Left            =   120
      TabIndex        =   5
      Top             =   2600
      Width           =   8535
      Begin VB.TextBox txtCadenaTFD 
         Height          =   915
         Left            =   2040
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   39
         Top             =   4670
         Width           =   6255
      End
      Begin VB.TextBox txtSelloDigitalSAT 
         Height          =   915
         Left            =   2040
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   37
         Top             =   3750
         Width           =   6255
      End
      Begin VB.TextBox txtCertificadoSAT 
         Height          =   315
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   3440
         Width           =   6255
      End
      Begin VB.TextBox txtFolio 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   390
         Width           =   2295
      End
      Begin VB.TextBox txtUUID 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   3120
         Width           =   6255
      End
      Begin VB.TextBox txtCadenaOriginal 
         Height          =   915
         Left            =   2040
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   24
         Top             =   1940
         Width           =   6255
      End
      Begin VB.TextBox txtSelloDigital 
         Height          =   915
         Left            =   2040
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   22
         Top             =   1015
         Width           =   6255
      End
      Begin VB.TextBox txtCertificado 
         Height          =   315
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   710
         Width           =   6255
      End
      Begin VB.TextBox txtFechaTimbrado 
         Height          =   315
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   390
         Width           =   2055
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Cadena del complemento TFD"
         Height          =   390
         Left            =   240
         TabIndex        =   40
         Top             =   4665
         Width           =   1680
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Sello digital SAT"
         Height          =   195
         Left            =   240
         TabIndex        =   38
         Top             =   3750
         Width           =   1155
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Certificado SAT"
         Height          =   195
         Left            =   240
         TabIndex        =   36
         Top             =   3440
         Width           =   1110
      End
      Begin VB.Label lblComprobante 
         AutoSize        =   -1  'True
         Caption         =   "COMPROBANTE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   32
         Top             =   390
         Width           =   1425
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Cadena original"
         Height          =   195
         Left            =   240
         TabIndex        =   23
         Top             =   1940
         Width           =   1095
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Sello digital"
         Height          =   195
         Left            =   240
         TabIndex        =   21
         Top             =   1015
         Width           =   795
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Certificado"
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   710
         Width           =   750
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de certificación"
         Height          =   195
         Left            =   4550
         TabIndex        =   18
         Top             =   390
         Width           =   1575
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "UUID (Folio fiscal)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   3120
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   2850
      TabIndex        =   0
      Top             =   8580
      Width           =   3080
      Begin VB.CommandButton cmdEnviar 
         Height          =   615
         Left            =   2260
         Picture         =   "frmComprobanteFiscalDigitalInternet.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Enviar por correo el comprobante fiscal digital"
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton cmdDescargaAcuse 
         Height          =   615
         Left            =   1530
         Picture         =   "frmComprobanteFiscalDigitalInternet.frx":101A
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Descargar acuse de cancelación"
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton cmdDescargaPDF 
         Height          =   615
         Left            =   60
         Picture         =   "frmComprobanteFiscalDigitalInternet.frx":2892
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Descargar archivo PDF"
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton cmdDescargaXML 
         Height          =   615
         Left            =   795
         Picture         =   "frmComprobanteFiscalDigitalInternet.frx":2D68
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Descargar archivo XML"
         Top             =   150
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmComprobanteFiscalDigitalInternet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public lngComprobante As Long       '|  Clave del comprobante fiscal digital
Public strTipoComprobante As String '|  Tipo del comprobante FA = Factura, CR = Nota de crédito, CA = Nota de cargo, DO = Donativo, NO = Nómina
Public blnCancelado As Boolean      '|  Indica si el comprobante está cancelado (habilita/deshabilita el botón de descargar acuse de cancelación)
Public blnFacturaSinComprobante As Boolean '| Indica si la factura se canceló y no generó un comprobante CFDi
Public llngMovPaciente As Long      '| Apoya al generar el reporte sin comprobante CFDi
Public strFolioFactura As String    '| Apoya al generar el reporte sin comprobante CFDi
Public dblTotal As Double           '| Apoya al generar el reporte sin comprobante CFDi para la cantidad en letras
Public strTipoReceptor As String    '| Tipo de cliente CO = Empresas, EM = Empleados, ME = Médicos, PI = Pacientes Internos, PE = Pacientes Externos
Public intReferenciaReceptor As Long 'Integer '|'Número de cuenta de paciente interno o externo, número de empresa, número de empleado o médico según corresponda (relacionado con EXPACIENTEINGRESO, CCEMPRESA, NOEMPLEADO, HOMEDICO)

Dim strRuta As String               '|  Ruta en la que se descargarán los archivos
Dim lngConsecutivo As Long          '|  Consecutivo del registro en la tabla GNCOMPROBANTEFISCALDIGITAL

Private Function fstrDestino(strNombreArchivo As String) As String
On Error GoTo NotificaError

    fstrDestino = ""
    cdlgDestino.FileName = strNombreArchivo
    cdlgDestino.DialogTitle = "Seleccione ubicación de descarga"
    cdlgDestino.CancelError = True
    cdlgDestino.ShowSave
    '|  Se se seleccionó un archivo, se regresa la ruta en la función, sino regresa vacío
    If cdlgDestino.FileName <> "" Then
        fstrDestino = cdlgDestino.FileName
    Else
        '|  Seleccione el dato.
        MsgBox SIHOMsg(431) & vbCrLf & "Destino del archivo.", vbCritical, "Mensaje"
        Exit Function
    End If
    
NotificaError:
    Err.Clear
End Function

''Descargar XML del acuse de cancelación
Private Sub cmdDescargaAcuse_Click()
On Error GoTo NotificaError:

    strRuta = fstrDestino(txtFolio.Text & "_AcuseCancelacion.xml")
    If strRuta = "" Then Exit Sub
    
    If fblnDescargaXMLCancelacion(lngConsecutivo, strRuta) Then
        MsgBox "Archivo XML descargado exitosamente.", vbInformation, "Mensaje"
    End If

Exit Sub
NotificaError:
    pRegistraError Err.Number, Err.Description, cgstrModulo, Me.Name
End Sub

Private Sub cmdDescargaPDF_Click()
    Dim lngCveFormato As Long
    Dim rsTipoPaciente As New ADODB.Recordset
    Dim lngCveTipoPaciente As Long
    Dim strTipoPaciente As String
    Dim intTipoAgrupacion As Integer
    Dim strSentencia As String
    Dim RsComprobante As New Recordset
    Dim rsEmpresaCFD As New ADODB.Recordset
    Dim vlintEmpresaCFD As Long
    Dim rsRutas As New ADODB.Recordset
    Dim strError As String
    Dim strRutaPDF As String
    
On Error GoTo NotificaError:
    If strTipoComprobante = "NO" Then
        strRutaPDF = fstrDestino(txtFolio.Text & ".pdf")
        If Not fblnImprimeComprobanteDigitalNom(lngComprobante, "PDF", strRutaPDF, True) Then Exit Sub
    Else
        If blnFacturaSinComprobante = False Then
            lngCveTipoPaciente = -2
            strTipoPaciente = ""
            intTipoAgrupacion = -1
            Set rsTipoPaciente = frsEjecuta_SP(lngComprobante & "|" & strTipoComprobante, "SP_GNSELTIPOPACIENTECFD")
            If rsTipoPaciente.RecordCount > 0 Then
                lngCveTipoPaciente = rsTipoPaciente!intCveTipoPaciente
                strTipoPaciente = rsTipoPaciente!VCHTIPOPACIENTE
                intTipoAgrupacion = rsTipoPaciente!intTipoDetalleFactura
            End If
            strRuta = fstrDestino(txtFolio.Text & ".pdf")
            
            If strRuta = "" Then Exit Sub
            
            If strTipoComprobante = "CR" Or strTipoComprobante = "CA" Then
                lngCveFormato = flngFormatoDepto(vgintNumeroDepartamento, 8, "*")
            ' Se agregó esta condición para determinar si es un donativo de tipo CFDi
            ElseIf Trim(strTipoComprobante) = "DO" Then
                lngCveFormato = -10  '----Se especifica el valor fijo para imprimir DONATIVOS de tipo CFDi (-10)
            ElseIf Trim(strTipoComprobante) = "RE" Then
                lngCveFormato = -20  '----Se especifica el valor fijo para imprimir comprobantes de pagos
            ElseIf Trim(strTipoComprobante) = "AN" Or Trim(strTipoComprobante) = "AA" Then
                'Anticipos
                lngCveFormato = 1
                frsEjecuta_SP vgintNumeroDepartamento & "|" & vlintEmpresaCFD & "|" & lngCveTipoPaciente & "|" & strTipoPaciente, "fn_PVSelFormatoFactura2", True, lngCveFormato
            Else
                If blnFacturaSinComprobante = True Then
                     If Not fblnImprimeComprobanteDigital(lngComprobante, strTipoComprobante, "PDF", lngCveFormato, 2, strRuta) Then Exit Sub
                Else
                    If strTipoComprobante = "FA" And strTipoPaciente = "C" Or strTipoPaciente = "S" Then
                        lngCveFormato = flngFormatoDepto(vgintNumeroDepartamento, 9, "*")
                    Else
                        lngCveFormato = 1
                        'Se verifica si el CFD es para una empresa, para obtener la clave de esta y seleccionar el formato correspondiente
                        Set rsEmpresaCFD = frsEjecuta_SP(CStr(lngComprobante), "SP_PVSELCVEEMPRESACFD")
                        
                        If rsEmpresaCFD.RecordCount > 0 And strTipoComprobante = "FA" Then
                            vlintEmpresaCFD = rsEmpresaCFD!ClaveEmpresa
                        Else
                            vlintEmpresaCFD = 0
                        End If
                            
                        frsEjecuta_SP vgintNumeroDepartamento & "|" & vlintEmpresaCFD & "|" & lngCveTipoPaciente & "|" & strTipoPaciente, "fn_PVSelFormatoFactura2", True, lngCveFormato
                    End If
               End If
            End If
                
            ' Se manda imprimir el comprobante
            If Not fblnImprimeComprobanteDigital(lngComprobante, strTipoComprobante, "PDF", lngCveFormato, intTipoAgrupacion, strRuta, True, strTipoReceptor, intReferenciaReceptor) Then Exit Sub
       
       ElseIf blnFacturaSinComprobante = True Then
            lngCveTipoPaciente = -2
            strTipoPaciente = ""
            intTipoAgrupacion = -1
            Set rsTipoPaciente = frsEjecuta_SP(CStr(lngComprobante), "SP_GNSELTIPOPACIENTEFACTURA")
            If rsTipoPaciente.RecordCount > 0 Then
                lngCveTipoPaciente = IIf(IsNull(rsTipoPaciente!intCveTipoPaciente), 0, rsTipoPaciente!intCveTipoPaciente)
                strTipoPaciente = rsTipoPaciente!VCHTIPOPACIENTE
                intTipoAgrupacion = rsTipoPaciente!intTipoDetalleFactura
            End If
            strRuta = fstrDestino(txtFolio.Text & ".pdf")
            
            If strRuta = "" Then Exit Sub
            ' Se manda imprimir la factura
            If Not fblnImprimeComprobanteFactura(dblTotal, lngComprobante, "PDF", lngCveFormato, 2, strRuta) Then Exit Sub
       End If
    End If
Exit Sub
NotificaError:
    pRegistraError Err.Number, Err.Description, cgstrModulo, Me.Name
End Sub

Private Sub cmdDescargaXML_Click()
On Error GoTo NotificaError:

    strRuta = fstrDestino(txtFolio.Text & ".xml")
    If strRuta = "" Then Exit Sub
    
    If strTipoComprobante = "NO" Then
        If Not fblnDescargaXMLCFDiNom(lngComprobante, strTipoComprobante, strRuta) Then Exit Sub
    Else
        If Not fblnDescargaXMLCFDi(lngComprobante, strTipoComprobante, strRuta) Then Exit Sub
    End If
    MsgBox "Archivo XML descargado exitosamente.", vbInformation, "Mensaje"
    
Exit Sub
NotificaError:
    pRegistraError Err.Number, Err.Description, cgstrModulo, Me.Name
End Sub

'- CASO 6217: Agregado para el envío de CFD por correo -'
Public Sub cmdEnviar_Click()
On Error GoTo NotificaError:

    If vgblnAutomatico = False Then
        pEnviarCFD strTipoComprobante, lngComprobante, CLng(vgintClaveEmpresaContable), Trim(txtRFC), 0, Me
    Else
        pEnviarCFD strTipoComprobante, lngComprobante, CLng(vgintClaveEmpresaContable), Trim(txtRFC), 0, Me, True
    End If
Exit Sub
NotificaError:
    pRegistraError Err.Number, Err.Description, cgstrModulo, Me.Name
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
    Me.Icon = frmMenuPrincipal.Icon
    pConsultaComprobante
    Me.cmdDescargaXML.Enabled = Me.txtUUID.Text <> " "
    cmdDescargaAcuse.Enabled = fblnAcuseXML And blnCancelado
    cmdEnviar.Enabled = fblnRevisaEnvioCorreo(vgintClaveEmpresaContable) And Not blnCancelado
End Sub

Public Sub pConsultaComprobante()
    Dim strSentencia As String
    Dim RsComprobante As New ADODB.Recordset
    Dim vlstRegimeFiscal As String
    Dim rsRegimen As New ADODB.Recordset
    
    Select Case strTipoComprobante
        Case "FA"
            lblComprobante.Caption = "FACTURA"
        Case "CR"
            lblComprobante.Caption = "NOTA DE CREDITO"
        Case "CA"
            lblComprobante.Caption = "NOTA DE CARGO"
        Case "DO"
            lblComprobante.Caption = "DONATIVO"
        Case Else
            lblComprobante.Caption = "COMPROBANTE"
    End Select
    
    If strTipoComprobante = "NO" Then
        strSentencia = "SELECT CFD.VCHRFCRECEPTOR " & _
                       "     , CFD.VCHNOMBRERECEPTOR " & _
                       "     , CFD.VCHCALLEDFRECEPTOR " & _
                       "     , CFD.VCHNOEXTERIORDFRECEPTOR " & _
                       "     , CFD.VCHNOINTERIORDFRECEPTOR " & _
                       "     , CFD.VCHCOLONIADFRECEPTOR " & _
                       "     , CFD.VCHMUNICIPIODFRECEPTOR " & _
                       "     , ' ' VCHCODIGOPOSTALDFRECEPTOR " & _
                       "     , CFD.VCHNUMEROCERTIFICADO, CFD.VCHCERTIFICADOSAT " & _
                       "     , CFD.VCHUUID, CFD.VCHSELLOSAT " & _
                       "     , CFD.VCHFECHATIMBRADO, CFD.VCHCADENATFD " & _
                       "     , CFD.CLBSELLODIGITAL " & _
                       "     , CFD.CLBCADENAORIGINAL " & _
                       "     , CFD.VCHSERIECOMPROBANTE || CFD.VCHFOLIOCOMPROBANTE Folio " & _
                       "     , CFD.INTIDCOMPROBANTE Consecutivo " & _
                       "  FROM GNCFDIGITALNOMINA CFD " & _
                       " WHERE CFD.INTIDCOMPROBANTE = " & lngComprobante & _
                       "   AND CFD.CHRTIPOCOMPROBANTE = '" & strTipoComprobante & "'"
    Else
        strSentencia = "SELECT CFD.VCHRFCRECEPTOR " & _
                       "     , CFD.VCHNOMBRERECEPTOR " & _
                       "     , CFD.VCHCALLEDFRECEPTOR " & _
                       "     , CFD.VCHNOEXTERIORDFRECEPTOR " & _
                       "     , CFD.VCHNOINTERIORDFRECEPTOR " & _
                       "     , CFD.VCHCOLONIADFRECEPTOR " & _
                       "     , CFD.VCHMUNICIPIODFRECEPTOR " & _
                       "     , CFD.VCHCODIGOPOSTALDFRECEPTOR " & _
                       "     , CFD.VCHNUMEROCERTIFICADO, CFD.VCHCERTIFICADOSAT " & _
                       "     , CFD.VCHUUID, CFD.VCHSELLOSAT " & _
                       "     , CFD.VCHFECHATIMBRADO, CFD.VCHCADENATFD " & _
                       "     , CFD.CLBSELLODIGITAL " & _
                       "     , CFD.CLBCADENAORIGINAL " & _
                       "     , CFD.VCHSERIECOMPROBANTE || CFD.VCHFOLIOCOMPROBANTE Folio " & _
                       "     , CFD.INTIDCOMPROBANTE Consecutivo " & _
                       "     , CFD.VCHREGIMENFISCALRECEPTOR regimen " & _
                       "  FROM GNCOMPROBANTEFISCALDIGITAL CFD " & _
                       " WHERE CFD.intComprobante = " & lngComprobante & _
                       "   AND CFD.CHRTIPOCOMPROBANTE = '" & strTipoComprobante & "'"
    End If
    Set RsComprobante = frsRegresaRs(strSentencia)
    If Not RsComprobante.EOF Then
        With RsComprobante
            txtRFC.Text = IIf(IsNull(!VCHRFCRECEPTOR), " ", !VCHRFCRECEPTOR)
            txtNombre.Text = IIf(IsNull(!VCHNOMBRERECEPTOR), " ", !VCHNOMBRERECEPTOR)
            txtCalle.Text = IIf(IsNull(!VCHCALLEDFRECEPTOR), " ", !VCHCALLEDFRECEPTOR)
            txtNumeroExterior.Text = IIf(IsNull(!VCHNOEXTERIORDFRECEPTOR), " ", !VCHNOEXTERIORDFRECEPTOR)
            txtNumeroInterior.Text = IIf(IsNull(!VCHNOINTERIORDFRECEPTOR), " ", !VCHNOINTERIORDFRECEPTOR)
            txtColonia.Text = IIf(IsNull(!VCHCOLONIADFRECEPTOR), " ", !VCHCOLONIADFRECEPTOR)
            txtCiudad.Text = IIf(IsNull(!VCHMUNICIPIODFRECEPTOR), " ", !VCHMUNICIPIODFRECEPTOR)
            txtCP.Text = IIf(IsNull(!VCHCODIGOPOSTALDFRECEPTOR), " ", !VCHCODIGOPOSTALDFRECEPTOR)
            txtCertificado.Text = IIf(IsNull(!VCHNUMEROCERTIFICADO), " ", !VCHNUMEROCERTIFICADO)
            txtSelloDigital.Text = IIf(IsNull(!CLBSELLODIGITAL), " ", !CLBSELLODIGITAL)
            txtCadenaOriginal.Text = IIf(IsNull(!CLBCADENAORIGINAL), " ", !CLBCADENAORIGINAL)
            txtFolio.Text = IIf(IsNull(!Folio), " ", !Folio)
            txtCadenaTFD.Text = IIf(IsNull(!VCHCADENATFD), " ", !VCHCADENATFD)
            txtUUID.Text = IIf(IsNull(!VCHUUID), " ", !VCHUUID)
            'Se formatea la fecha y hora de timbrado
            txtFechaTimbrado.Text = IIf(IsNull(!VCHFECHATIMBRADO), " ", !VCHFECHATIMBRADO)
            txtSelloDigitalSAT.Text = IIf(IsNull(!VCHSELLOSAT), " ", !VCHSELLOSAT)
            txtCertificadoSAT.Text = IIf(IsNull(!VCHCERTIFICADOSAT), " ", !VCHCERTIFICADOSAT)
            lngConsecutivo = IIf(IsNull(!Consecutivo), 0, !Consecutivo) 'Agregado para caso 7994
            vlstRegimeFiscal = IIf(IsNull(!regimen), 0, !regimen) 'Agregado para caso 7994
        End With
    Else
        If blnFacturaSinComprobante = True Then
            strSentencia = "SELECT * from PVFACTURA WHERE INTMOVPACIENTE = " & llngMovPaciente & " and CHRFOLIOFACTURA = '" & Trim(CStr(strFolioFactura)) & "'"
            Set RsComprobante = frsRegresaRs(strSentencia)
            With RsComprobante
                If RsComprobante.RecordCount > 0 Then
                    txtRFC.Text = IIf(IsNull(!chrRFC), " ", !chrRFC)
                    txtNombre.Text = IIf(IsNull(!CHRNOMBRE), " ", !CHRNOMBRE)
                    txtCalle.Text = IIf(IsNull(!CHRCALLE), " ", !CHRCALLE)
                    txtNumeroExterior.Text = IIf(IsNull(!VCHNUMEROEXTERIOR), " ", !VCHNUMEROEXTERIOR)
                    txtNumeroInterior.Text = IIf(IsNull(!VCHNUMEROINTERIOR), " ", !VCHNUMEROINTERIOR)
                    txtColonia.Text = IIf(IsNull(!VCHCOLONIA), " ", !VCHCOLONIA)
                    txtCiudad.Text = IIf(IsNull(!VCHCIUDAD), " ", !VCHCIUDAD)
                    txtCP.Text = IIf(IsNull(!VCHCODIGOPOSTAL), " ", !VCHCODIGOPOSTAL)
                    txtCertificado.Text = " "
                    txtSelloDigital.Text = " "
                    txtCadenaOriginal.Text = " "
                    txtFolio.Text = Trim(strFolioFactura)
                    txtCadenaTFD.Text = " "
                    txtUUID.Text = " "
                    txtFechaTimbrado.Text = " "
                    txtSelloDigitalSAT.Text = " "
                    txtCertificadoSAT.Text = " "
                    vlstRegimeFiscal = IIf(IsNull(!VCHREGIMENFISCALRECEPTOR), "", !VCHREGIMENFISCALRECEPTOR) 'VCHREGIMENFISCALRECEPTOR
                End If
            End With
        End If
    End If
    
    If vlstRegimeFiscal <> "" Then
        Set rsRegimen = frsRegresaRs("Select GNCatalogoSATDetalle.vchClave || ' - ' || GNCatalogoSATDetalle.vchDescripcion As RegimenFiscal FROM GNCATALOGOSATDETALLE WHERE VCHCLAVE = '" & vlstRegimeFiscal & "'")
        If Not rsRegimen.EOF Then
            txtRegimen.Text = IIf(IsNull(rsRegimen!RegimenFiscal), "", rsRegimen!RegimenFiscal)
        End If
    End If
End Sub

'- CASO 7994: Función para indicar si existe Acuse de Cancelación del comprobante -'
Private Function fblnAcuseXML() As Boolean
On Error GoTo NotificaError:

    Dim rsAcuse As New ADODB.Recordset
    Dim strSentencia As String
    
    strSentencia = "SELECT * FROM GNACUSECANCELACIONCFDI WHERE INTIDCOMPROBANTE = " & lngConsecutivo
    Set rsAcuse = frsRegresaRs(strSentencia, adLockOptimistic, adOpenDynamic)
    fblnAcuseXML = (rsAcuse.RecordCount > 0)
    rsAcuse.Close
    
Exit Function
NotificaError:
    pRegistraError Err.Number, Err.Description, cgstrModulo, Me.Name
End Function


