VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmComprobanteFiscalDigital 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Comprobante fiscal digital"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8775
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   8775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cdlgDestino 
      Left            =   1080
      Top             =   6240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame freDatosFiscales 
      Caption         =   "Datos fiscales"
      Enabled         =   0   'False
      Height          =   2085
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   8535
      Begin VB.TextBox txtCiudad 
         Height          =   315
         Left            =   6150
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   1605
         Width           =   2145
      End
      Begin VB.TextBox txtNumeroInterior 
         Height          =   315
         Left            =   4410
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1290
         Width           =   855
      End
      Begin VB.TextBox txtNumeroExterior 
         Height          =   315
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1290
         Width           =   1095
      End
      Begin VB.TextBox txtNombre 
         Height          =   315
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   660
         Width           =   6255
      End
      Begin VB.TextBox txtCP 
         Height          =   315
         Left            =   6150
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1290
         Width           =   2145
      End
      Begin VB.TextBox txtColonia 
         Height          =   315
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1605
         Width           =   3225
      End
      Begin VB.TextBox txtRFC 
         Height          =   315
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   350
         Width           =   1545
      End
      Begin VB.TextBox txtCalle 
         Height          =   315
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   975
         Width           =   6255
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Número exterior"
         Height          =   195
         Left            =   240
         TabIndex        =   29
         Top             =   1352
         Width           =   1110
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Razón social"
         Height          =   195
         Left            =   240
         TabIndex        =   28
         Top             =   724
         Width           =   915
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Ciudad"
         Height          =   195
         Left            =   5400
         TabIndex        =   27
         Top             =   1665
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Colonia"
         Height          =   195
         Left            =   240
         TabIndex        =   26
         Top             =   1666
         Width           =   525
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Calle"
         Height          =   195
         Left            =   240
         TabIndex        =   25
         Top             =   1038
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "RFC"
         Height          =   195
         Left            =   240
         TabIndex        =   24
         Top             =   410
         Width           =   315
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "C.P."
         Height          =   195
         Left            =   5400
         TabIndex        =   14
         Top             =   1320
         Width           =   300
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Número interior"
         Height          =   195
         Left            =   3240
         TabIndex        =   13
         Top             =   1320
         Width           =   1065
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del comprobante"
      Height          =   3495
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   8535
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
         TabIndex        =   33
         Top             =   390
         Width           =   3375
      End
      Begin VB.TextBox txtNumeroAprobacion 
         Height          =   315
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   720
         Width           =   3375
      End
      Begin VB.TextBox txtCadenaOriginal 
         Height          =   915
         Left            =   2040
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   23
         Top             =   2265
         Width           =   6255
      End
      Begin VB.TextBox txtSelloDigital 
         Height          =   915
         Left            =   2040
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   21
         Top             =   1350
         Width           =   6255
      End
      Begin VB.TextBox txtCertificado 
         Height          =   315
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   1035
         Width           =   6255
      End
      Begin VB.TextBox txtAnoAprobacion 
         Height          =   315
         Left            =   7080
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   720
         Width           =   1215
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
         TabIndex        =   31
         Top             =   450
         Width           =   1425
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Cadena original"
         Height          =   195
         Left            =   240
         TabIndex        =   22
         Top             =   2265
         Width           =   1095
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Sello digital"
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   1350
         Width           =   795
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Certificado"
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   1065
         Width           =   750
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Año de aprobación"
         Height          =   195
         Left            =   5640
         TabIndex        =   17
         Top             =   780
         Width           =   1350
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Número de aprobación"
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   780
         Width           =   1620
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   3270
      TabIndex        =   0
      Top             =   5880
      Width           =   2335
      Begin VB.CommandButton cmdEnviar 
         Height          =   615
         Left            =   1530
         Picture         =   "frmComprobanteFiscalDigital.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Enviar por correo el comprobante fiscal digital"
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton cmdDescargaPDF 
         Height          =   615
         Left            =   60
         Picture         =   "frmComprobanteFiscalDigital.frx":101A
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton cmdDescargaXML 
         Height          =   615
         Left            =   795
         Picture         =   "frmComprobanteFiscalDigital.frx":14F0
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   150
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmComprobanteFiscalDigital"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public lngComprobante As Long       '|  Clave del comprobante fiscal digital
Public strTipoComprobante As String '|  Tipo del comprobante FA = Factura, CR = Nota de crédito, CA = Nota de cargo, DO = Donativo
Public blnCancelado As Boolean      '|  Indica si el comprobante está cancelado (habilita/deshabilita el botón de enviar)

Private strRuta As String           '|

Private Sub cmdDescargaPDF_Click()
    Dim lngCveFormato As Long
    Dim rsTipoPaciente As New ADODB.Recordset
    Dim lngCveTipoPaciente As Long
    Dim strTipoPaciente As String
    Dim intTipoAgrupacion As Integer
    Dim strSentencia As String
    Dim rsComprobante As New Recordset
    Dim strRuta As String
    Dim rsEmpresaCFD As New ADODB.Recordset
    Dim vlintEmpresaCFD As Long
    
On Error GoTo NotificaError:

    lngCveTipoPaciente = -2
    strTipoPaciente = ""
    intTipoAgrupacion = -1
    Set rsTipoPaciente = frsEjecuta_SP(lngComprobante & "|" & strTipoComprobante, "SP_GNSELTIPOPACIENTECFD")
    If rsTipoPaciente.RecordCount > 0 Then
        lngCveTipoPaciente = rsTipoPaciente!INTCVETIPOPACIENTE
        strTipoPaciente = rsTipoPaciente!VCHTIPOPACIENTE
        intTipoAgrupacion = rsTipoPaciente!intTipoDetalleFactura
    End If
    strRuta = fstrDestino(txtFolio.Text & ".pdf")
    If strRuta = "" Then Exit Sub
    
    If strTipoComprobante = "CR" Or strTipoComprobante = "CA" Then
        lngCveFormato = flngFormatoDepto(vgintNumeroDepartamento, 8, "*")
    ' Se agregó esta condición para determinar si es un donativo de tipo CFD
    ElseIf Trim(strTipoComprobante) = "DO" Then
        lngCveFormato = -1  '----Se especifica el valor fijo para imprimir DONATIVOS de tipo CFD (-1)
    Else
        If strTipoComprobante = "FA" And strTipoPaciente = "C" Then
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
    
    ' Se manda imprimir el comprobante
    If Not fblnImprimeComprobanteDigital(lngComprobante, strTipoComprobante, "PDF", lngCveFormato, intTipoAgrupacion, strRuta) Then Exit Sub
    
    Exit Sub
NotificaError:
    pRegistraError Err.Number, Err.Description, cgstrModulo, Me.Name
End Sub

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

Private Sub cmdDescargaXML_Click()
    Dim strRuta As String
On Error GoTo NotificaError:

    strRuta = fstrDestino(txtFolio.Text & ".xml")
    If strRuta = "" Then Exit Sub
    
    If Not fblnDescargaXML(lngComprobante, strTipoComprobante, strRuta) Then Exit Sub
    MsgBox "Archivo XML descargado exitosamente.", vbInformation, "Mensaje"

Exit Sub
NotificaError:
    pRegistraError Err.Number, Err.Description, cgstrModulo, Me.Name
End Sub

'- CASO 6217: Agregado para el envío de CFD por correo -'
Private Sub cmdEnviar_Click()
On Error GoTo NotificaError:

'    Dim vllngPersonaGraba As Long
'
'    'Persona que envía el correo
'    vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
'    If vllngPersonaGraba = 0 Then Exit Sub
    
    pEnviarCFD strTipoComprobante, lngComprobante, CLng(vgintClaveEmpresaContable), Trim(txtRFC), 0, Me
    
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
    
    cmdEnviar.Enabled = fblnRevisaEnvioCorreo(vgintClaveEmpresaContable) And Not blnCancelado
End Sub

Private Sub pConsultaComprobante()
    Dim strSentencia As String
    Dim rsComprobante As New ADODB.Recordset
    
    Select Case strTipoComprobante
        Case "FA"
            lblComprobante.Caption = "FACTURA"
        Case "CR"
            lblComprobante.Caption = "NOTA DE CREDITO"
        Case "CA"
            lblComprobante.Caption = "NOTA DE CARGO"
        Case "DO"
            lblComprobante.Caption = "DONATIVO"
    End Select
        
    strSentencia = "SELECT CFD.VCHRFCRECEPTOR " & _
                   "     , CFD.VCHNOMBRERECEPTOR " & _
                   "     , CFD.VCHCALLEDFRECEPTOR " & _
                   "     , CFD.VCHNOEXTERIORDFRECEPTOR " & _
                   "     , CFD.VCHNOINTERIORDFRECEPTOR " & _
                   "     , CFD.VCHCOLONIADFRECEPTOR " & _
                   "     , CFD.VCHMUNICIPIODFRECEPTOR " & _
                   "     , CFD.VCHCODIGOPOSTALDFRECEPTOR " & _
                   "     , CFD.VCHNUMEROCERTIFICADO " & _
                   "     , CFD.INTNUMEROAPROBACION " & _
                   "     , CFD.INTANOAPROBACION " & _
                   "     , CFD.CLBSELLODIGITAL " & _
                   "     , CFD.CLBCADENAORIGINAL " & _
                   "     , CFD.VCHSERIECOMPROBANTE || CFD.VCHFOLIOCOMPROBANTE Folio " & _
                   "  FROM GNCOMPROBANTEFISCALDIGITAL CFD " & _
                   " WHERE CFD.intComprobante = " & lngComprobante & _
                   "   AND CFD.CHRTIPOCOMPROBANTE = '" & strTipoComprobante & "'"
    Set rsComprobante = frsRegresaRs(strSentencia)
    If Not rsComprobante.EOF Then
        With rsComprobante
            txtRFC.Text = IIf(IsNull(!VCHRFCRECEPTOR), " ", !VCHRFCRECEPTOR)
            txtNombre.Text = IIf(IsNull(!VCHNOMBRERECEPTOR), " ", !VCHNOMBRERECEPTOR)
            txtCalle.Text = IIf(IsNull(!VCHCALLEDFRECEPTOR), " ", !VCHCALLEDFRECEPTOR)
            txtNumeroExterior.Text = IIf(IsNull(!VCHNOEXTERIORDFRECEPTOR), " ", !VCHNOEXTERIORDFRECEPTOR)
            txtNumeroInterior.Text = IIf(IsNull(!VCHNOINTERIORDFRECEPTOR), " ", !VCHNOINTERIORDFRECEPTOR)
            txtColonia.Text = IIf(IsNull(!VCHCOLONIADFRECEPTOR), " ", !VCHCOLONIADFRECEPTOR)
            txtCiudad.Text = IIf(IsNull(!VCHMUNICIPIODFRECEPTOR), " ", !VCHMUNICIPIODFRECEPTOR)
            txtCP.Text = IIf(IsNull(!VCHCODIGOPOSTALDFRECEPTOR), " ", !VCHCODIGOPOSTALDFRECEPTOR)
            txtCertificado.Text = IIf(IsNull(!VCHNUMEROCERTIFICADO), " ", !VCHNUMEROCERTIFICADO)
            txtNumeroAprobacion.Text = IIf(IsNull(!INTNUMEROAPROBACION), " ", !INTNUMEROAPROBACION)
            txtAnoAprobacion.Text = IIf(IsNull(!intAnoAprobacion), " ", !intAnoAprobacion)
            txtSelloDigital.Text = IIf(IsNull(!CLBSELLODIGITAL), " ", !CLBSELLODIGITAL)
            txtCadenaOriginal.Text = IIf(IsNull(!CLBCADENAORIGINAL), " ", !CLBCADENAORIGINAL)
            txtFolio.Text = IIf(IsNull(!Folio), " ", !Folio)
        End With
    End If
End Sub

