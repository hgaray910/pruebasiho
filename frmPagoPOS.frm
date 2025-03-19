VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmPagoPos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Formas de pago"
   ClientHeight    =   8460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9135
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8460
   ScaleWidth      =   9135
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraPinPad 
      Height          =   1005
      Left            =   0
      TabIndex        =   34
      Top             =   -80
      Visible         =   0   'False
      Width           =   9131
      Begin VB.Timer Timer1 
         Left            =   4920
         Top             =   480
      End
      Begin VB.CommandButton cmdOkPP 
         Caption         =   "Aceptar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3240
         TabIndex        =   39
         Top             =   5760
         Width           =   2655
      End
      Begin VB.Timer tmrMsgPP 
         Enabled         =   0   'False
         Interval        =   3000
         Left            =   3960
         Top             =   480
      End
      Begin VB.Label lblErr 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   120
         TabIndex        =   38
         Top             =   2520
         Width           =   8895
      End
      Begin VB.Label lblDeclinada 
         Alignment       =   2  'Center
         Caption         =   "DECLINADA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   120
         TabIndex        =   37
         Top             =   3720
         Visible         =   0   'False
         Width           =   8895
      End
      Begin VB.Label lblAprobada 
         Alignment       =   2  'Center
         Caption         =   "APROBADA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   615
         Left            =   120
         TabIndex        =   36
         Top             =   3720
         Visible         =   0   'False
         Width           =   8895
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Siga las instrucciones en el Pinpad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   120
         TabIndex        =   35
         Top             =   1200
         Width           =   8895
      End
   End
   Begin VB.Frame fraPrincipales 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2680
      Left            =   4630
      TabIndex        =   27
      Top             =   1680
      Width           =   4345
      Begin VB.TextBox txtCantidad 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   0
         TabIndex        =   10
         ToolTipText     =   "Cantidad a pagar con la forma de pago"
         Top             =   2190
         Width           =   4320
      End
      Begin VB.TextBox txtFolio 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   0
         MaxLength       =   20
         TabIndex        =   9
         ToolTipText     =   "Folio de referencia de la forma de pago"
         Top             =   1440
         Width           =   4320
      End
      Begin VB.ComboBox cboTipoCargoBancario 
         Height          =   315
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   8
         ToolTipText     =   "Seleccione el tipo de cargo bancario"
         Top             =   840
         Width           =   4335
      End
      Begin VB.ComboBox cboBanco 
         Height          =   315
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   7
         ToolTipText     =   "Seleccione la cuenta bancaria receptora del pago"
         Top             =   230
         Width           =   4335
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad"
         Height          =   195
         Left            =   1852
         TabIndex        =   31
         Top             =   1970
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Referencia"
         Height          =   195
         Left            =   1777
         TabIndex        =   30
         Top             =   1200
         Width           =   780
      End
      Begin VB.Label lblTipoCargoBancario 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de cargo bancario"
         Height          =   195
         Left            =   0
         TabIndex        =   29
         Top             =   600
         Width           =   1710
      End
      Begin VB.Label lblBanco 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta bancaria receptora del pago"
         Height          =   195
         Left            =   0
         TabIndex        =   28
         Top             =   0
         Width           =   2550
      End
   End
   Begin SysInfoLib.SysInfo SysInfo 
      Left            =   240
      Top             =   5280
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Frame fraAbajo 
      BorderStyle     =   0  'None
      Height          =   3910
      Left            =   0
      TabIndex        =   20
      Top             =   4440
      Width           =   9160
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   1200
         Top             =   1080
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Frame Frame1 
         Height          =   520
         Left            =   2040
         TabIndex        =   40
         Top             =   3300
         Visible         =   0   'False
         Width           =   2000
         Begin VB.CommandButton cmdClear 
            Caption         =   "Reiniciar"
            Height          =   375
            Left            =   990
            TabIndex        =   42
            Top             =   120
            Width           =   975
         End
         Begin VB.CommandButton cmdReimprimir 
            Caption         =   "Reimprimir"
            Height          =   375
            Left            =   20
            TabIndex        =   41
            Top             =   120
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         Height          =   750
         Left            =   120
         TabIndex        =   21
         Top             =   3170
         Width           =   1785
         Begin VB.CommandButton cmdAceptar 
            Caption         =   "Aceptar"
            Enabled         =   0   'False
            Height          =   495
            Left            =   105
            TabIndex        =   12
            ToolTipText     =   "Aceptar los registros realizados "
            Top             =   180
            Width           =   1605
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdFormas 
         Height          =   1710
         Left            =   30
         TabIndex        =   11
         ToolTipText     =   "Formas registradas"
         Top             =   0
         Width           =   9060
         _ExtentX        =   15981
         _ExtentY        =   3016
         _Version        =   393216
         Rows            =   0
         FixedRows       =   0
         GridColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Frame fraTotales 
         Height          =   2220
         Left            =   4440
         TabIndex        =   22
         Top             =   1700
         Width           =   4590
         Begin VB.TextBox txtImporte 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   2325
            Locked          =   -1  'True
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   240
            Width           =   2205
         End
         Begin VB.TextBox txtCantidadPagada 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   2325
            Locked          =   -1  'True
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   690
            Width           =   2205
         End
         Begin VB.TextBox txtDiferencia 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   420
            Left            =   2325
            Locked          =   -1  'True
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   1140
            Width           =   2205
         End
         Begin VB.TextBox txtCambio 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   480
            Left            =   2325
            Locked          =   -1  'True
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   1590
            Width           =   2205
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Importe a pagar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   165
            TabIndex        =   26
            Top             =   330
            Width           =   1665
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Cantidad pagada"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   165
            TabIndex        =   25
            Top             =   780
            Width           =   1815
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Diferencia"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   165
            TabIndex        =   24
            Top             =   1245
            Width           =   1080
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Cambio"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   165
            TabIndex        =   23
            Top             =   1710
            Width           =   810
         End
      End
   End
   Begin VB.Frame fraInformacionExtra 
      BorderStyle     =   0  'None
      Height          =   2385
      Left            =   120
      TabIndex        =   17
      Top             =   1680
      Width           =   4345
      Begin VB.ComboBox cboCuentasPrevias 
         Height          =   315
         Left            =   0
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Seleccione la cuenta bancaria"
         Top             =   1440
         Visible         =   0   'False
         Width           =   4335
      End
      Begin VB.TextBox txtCuentaBancaria 
         Height          =   315
         Left            =   0
         MaxLength       =   50
         TabIndex        =   4
         ToolTipText     =   "Cuenta bancaria emisora del cheque, transferencia o tarjeta"
         Top             =   1440
         Width           =   4335
      End
      Begin VB.TextBox txtRFC 
         Height          =   315
         Left            =   0
         MaxLength       =   13
         TabIndex        =   1
         ToolTipText     =   "RFC relacionado con el pago"
         Top             =   230
         Width           =   1740
      End
      Begin VB.TextBox txtBancoExtranjero 
         Height          =   315
         Left            =   0
         MaxLength       =   150
         TabIndex        =   3
         Text            =   "BancoExtranjero"
         ToolTipText     =   "Banco extranjero emisor del del cheque, transferencia o tarjeta"
         Top             =   840
         Visible         =   0   'False
         Width           =   4335
      End
      Begin VB.ComboBox cboBancoSAT 
         Height          =   315
         Left            =   0
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Seleccione el banco del SAT emisor del cheque, transferencia o tarjeta"
         Top             =   840
         Width           =   4335
      End
      Begin MSMask.MaskEdBox MskFecha 
         Height          =   315
         Left            =   0
         TabIndex        =   6
         ToolTipText     =   "Fecha del cheque o transferencia"
         Top             =   2040
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mmm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label lblCuentaBancaria 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta bancaria emisora del cheque, transferencia o tarjeta"
         Height          =   195
         Left            =   0
         TabIndex        =   33
         Top             =   1200
         Width           =   4215
      End
      Begin VB.Label lblFecha 
         AutoSize        =   -1  'True
         Caption         =   "Fecha del cheque o transferencia"
         Height          =   195
         Left            =   0
         TabIndex        =   32
         Top             =   1800
         Width           =   2385
      End
      Begin VB.Label lblBancoSAT 
         AutoSize        =   -1  'True
         Caption         =   "Banco emisor del cheque, transferencia o tarjeta"
         Height          =   195
         Left            =   0
         TabIndex        =   19
         Top             =   600
         Width           =   3420
      End
      Begin VB.Label lblRFC 
         AutoSize        =   -1  'True
         Caption         =   "RFC relacionado con el pago"
         Height          =   195
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   2070
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdFormasPago 
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Seleccione la forma de pago"
      Top             =   120
      Width           =   4410
      _ExtentX        =   7779
      _ExtentY        =   2566
      _Version        =   393216
      GridColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frmPagoPos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Programa para registrar las formas de pago en que se paga una cantidad x, esta forma
' es llamada por la funcion fblnFormasPagoPos
' Fecha de programación: Miércoles 24 de Enero de 2001
' Fecha de última modificación: 21/Mar/2002
'---------------------------------------------------------------------------------------
Private strFormaPadre As String

Public vlstrTipoCliente As String
Public vldblTipoCambioDia As Double
Public vldblCantidadPago As Double
Public vlblnPesos As Boolean
Public vlblnIncluirFormasCredito As Boolean
Public vlblnCambioCantidad As Boolean
Public vlLngReferencia As Long
Public vgintSalidaOK As String
Public vgblnPermiteSalir As Boolean 'Si está prendida significa que se habilitará el botón de aceptar con cualquier cantidad tecleada
                                    'sino será necesario que la cantidad pagada sea igual al importe a pagar para habilitar el botón
Public vgstrForma As String 'Se utiliza para saber de que que forma mando llamar. Esto es porque de la forma de Honorarios cuando se
                            'pague a crédito, no se debe de poder combinar con otra forma de pago

Public lblnFormaTrans As Boolean 'Para saber cuando se cargarán formas de pago tipo "Transferencia"
Public vlstrRFCOriginal As String
Public vllngCveEmpleado As Long

Dim vl_blnTime As Boolean

'Columnas <grdFormasPago>
Const cintColIdForma = 0
Const cintColNombre = 1
Const cintColMoneda = 2
Const cintColReferencia = 3
Const cintColTipo = 4
Const cintColCtaContable = 5
Const cintColClaveFormaPagoSAT = 6
Const cintColUsarPinpad1 = 7
Const cintColUriPinpad1 = 8
Const cintColImprVoucher1 = 9
Const cintColCveMoneda1 = 10
Const cintColPpProv1 = 11
Const cintColPpUsr1 = 12
Const cintColPpPwd1 = 13
Const cintColPpHost1 = 14
Const cintColPpPort1 = 15
Const cintColPpCve1 = 16
Const cintColPpTId1 = 17
Const cintColumnas = 18

'Columnas <grdFormas>
Const cintColCol = 0
Const cintColDescripcion = 1
Const cIntColFolio = 2
Const cintColCantidad = 3
Const cintColCtaContableFormaPago = 4
Const cintColTipoCambio = 5
Const cintColCantidadReal = 6
Const cintColCredito = 7
Const cintColDolares = 8
Const cintColIdBanco = 9
Const cintColMonedaBanco = 10
Const cintColCtaComisionBanc = 11
Const cintColComisionBanc = 12
Const cintColIVAComisionBanc = 13
Const cIntColRFC = 14
Const cintColBancoSAT = 15
Const cintColBancoExt = 16
Const cIntColCuentaBancaria = 17
Const cintColFechaCqTrans = 18
Const cintColUsarPinpad2 = 19
Const cintColUriPinpad2 = 20
Const cintColImprVoucher2 = 21
Const cintColCveMoneda2 = 22
Const cintColPpProv2 = 23
Const cintColPpUsr2 = 24
Const cintColPpPwd2 = 25
Const cintColPpHost2 = 26
Const cintColPpPort2 = 27
Const cintColPpCve2 = 28
Const cintColPpTId2 = 29

Dim vlstrsSQL As String
Dim vlstrValorAnterior As String
Dim vldblTipoCambio As Double
Dim vldblValorSinFormato As Double
Dim vldblValorConDecimales As Double
Dim vldblLimiteCredito As Double
Dim vllngNumCliente As Long
Dim vllngDeptoCliente As Long
Dim vllngCuentaContableCredito As Long
Dim rsFormasPago As New ADODB.Recordset
Dim vlblnExisteError As Boolean

Dim vlblnRegistroCredito As Boolean
Dim vlblnExisteFormaCredito As Boolean
Dim vldblCantidadCredito As Double
Dim vlintFormaCredito As Integer
Dim vlintNumRegistro As Integer
Dim rsBancos As New ADODB.Recordset
Dim rsTipoCargoBancario As New ADODB.Recordset
Dim vlblnMuestraCargoBancario As Boolean

Dim vldtmFecha As Date
Dim vldtmfechaServer As Date
Dim vlblnLicenciaContaElectronica As Boolean
Dim vlClaveBancoSAT As String
Dim vlstrRFCTemporal As String

Dim lngCtaDevolucionesCuentasPagar As Long

Dim WithEvents ws As WebSocketWrap.Client
Attribute ws.VB_VarHelpID = -1
Dim blnRespuestaEsperar As Boolean
Dim strRespuestaEsperar As String
Dim intTimeout As Integer
Dim vlblnTimeout As Boolean
Dim vlstrRef As String

Dim Conectado As Boolean
Public intModoMasivo As Integer

Dim rsBanco As ADODB.Recordset

Public Function fblnValidaCtaDevolucionesxCP(intEmpresa As Integer) As Boolean
On Error GoTo NotificaError

    Dim rs As ADODB.Recordset
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    '---------------------------------------------------------------------------------'
    ' Valida la "Cuenta para devoluciones a paciente por cuentas por pagar"           '
    ' para el uso de la forma de pago "DEVOLUCIÓN A PACIENTE POR CUENTAS POR PAGAR" '
    '---------------------------------------------------------------------------------'
    
    fblnValidaCtaDevolucionesxCP = True
    
    ' Regresa valor de parámetro que indica la cuenta para devoluciones por cuentas por pagar
    lngCtaDevolucionesCuentasPagar = 0
    Set rsTemp = frsSelParametros("CN", vgintClaveEmpresaContable, "INTNUMCTADEVOLUCIONESPACIENTE")
    If rsTemp.EOF Then
        'No se ha registrado la cuenta para devoluciones a paciente por cuentas por pagar.
        MsgBox SIHOMsg(1403), vbOKOnly + vbInformation, "Mensaje"
        fblnValidaCtaDevolucionesxCP = False
    Else
        lngCtaDevolucionesCuentasPagar = IIf(IsNull(rsTemp!Valor), 0, CLng(rsTemp!Valor))
        If lngCtaDevolucionesCuentasPagar = 0 Then
            'No se ha registrado la cuenta para devoluciones a paciente por cuentas por pagar.
            MsgBox SIHOMsg(1403), vbOKOnly + vbInformation, "Mensaje"
            fblnValidaCtaDevolucionesxCP = False
        End If
    End If
    rsTemp.Close
               
    ' Regresa valores de la cuenta para devoluciones por cuentas por pagar
    If fblnValidaCtaDevolucionesxCP Then
        Set rsTemp = frsEjecuta_SP(CStr(lngCtaDevolucionesCuentasPagar), "Sp_CnSelCuentaContable")
        If rsTemp.RecordCount > 0 Then
            If IsNull(rsTemp!bitEstatusActiva) Or rsTemp!bitEstatusActiva = 0 Then
                'La cuenta para devoluciones a paciente por cuentas por pagar, no está activa.
                MsgBox SIHOMsg(1404), vbOKOnly + vbInformation, "Mensaje"
                fblnValidaCtaDevolucionesxCP = False
            Else
                If IsNull(rsTemp!Bitestatusmovimientos) Or rsTemp!Bitestatusmovimientos = 0 Then
                    'La cuenta para devoluciones a paciente por cuentas por pagar, no acepta movimientos.
                    MsgBox SIHOMsg(1405), vbOKOnly + vbInformation, "Mensaje"
                    fblnValidaCtaDevolucionesxCP = False
                End If
            End If
        End If
        rsTemp.Close
    End If

Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnValidaCtadevolucionesxCP"))
End Function


Private Function fintLocMshFGrd(ObjGrid As MSHFlexGrid, vlstrCriterio As String, vlintColBus As Integer) As Integer
    '-------------------------------------------------------------------------------------------
    ' Realiza la busqueda de un criterio dentro de una de las columnas del grdHBusqueda
    ' señalando los criterios encontrados
    '-------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    Dim vlintNumFilas As Integer 'Almacena el número de filas que contiene el grdHBusqueda
    Dim vlintseq As Integer 'Contador del número de filas del grdHBusqueda
    Dim vlintEFila As Integer 'Fila que se encuentra mediante el criterio de búsqueda
    Dim vlintLargo As Integer 'Almacena el largo del criterio de búsqueda
    Dim vlstrTexto As String 'Almacena los caracteres obtenidos de las celda del grid, segun el largo de busqueda del criterio, para su comparacion con el criterio de búsqueda
    
    vlintLargo = Len(vlstrCriterio)
    vlintEFila = 0 'Inicializa la búsqueda desde la primera fila
    vlintNumFilas = ObjGrid.Rows - 1
    
    If vlintLargo > 0 And vlintNumFilas > 0 Then
        For vlintseq = 1 To vlintNumFilas 'Realiza la búsqueda en todo el grid
            vlstrTexto = ObjGrid.TextMatrix(vlintseq, vlintColBus)
            If UCase(vlstrCriterio) = UCase(vlstrTexto) Then
                If vlintLargo > 0 Then
                    vlintEFila = vlintseq
                End If
            End If
        Next vlintseq
        fintLocMshFGrd = vlintEFila
    Else
        fintLocMshFGrd = 0
    End If
    
Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fintLocMshFGrd"))

End Function

Private Sub pCargosBancarios()
    cboTipoCargoBancario.Clear

    If vlblnMuestraCargoBancario And Trim(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColTipo)) = "T" Then
        'Cargar los tipos de cargos bancarios
        vgstrParametrosSP = Trim(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColIdForma)) & "|-1| " & CStr(vgintClaveEmpresaContable)
        Set rsTipoCargoBancario = frsEjecuta_SP(vgstrParametrosSP, "sp_pvSelComisionCargoBancario")
        If rsTipoCargoBancario.RecordCount <> 0 Then
            pLlenarCboRs cboTipoCargoBancario, rsTipoCargoBancario, 0, 1
            rsTipoCargoBancario.MoveFirst
            Do While Not rsTipoCargoBancario.EOF
                If rsTipoCargoBancario!bitpredeterminado = 1 Then cboTipoCargoBancario.Text = rsTipoCargoBancario!chrDescripcion
                rsTipoCargoBancario.MoveNext
            Loop
            cboTipoCargoBancario.Enabled = True
            lblTipoCargoBancario.Enabled = True
        Else
            cboTipoCargoBancario.Enabled = False
            lblTipoCargoBancario.Enabled = False
        End If
    Else
        cboTipoCargoBancario.Enabled = False
        lblTipoCargoBancario.Enabled = False
    End If
End Sub

Private Sub cboBanco_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If cboTipoCargoBancario.Enabled Then
            SendKeys vbTab
        Else
            If Val(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColReferencia)) = 1 Then
                txtFolio.SetFocus
            Else
                txtCantidad.SetFocus
            End If
        End If
    End If
End Sub

Private Sub cboBancoSAT_Change()
    pHabilitaFechaChequeTrans
End Sub

Private Sub cboBancoSAT_Click()
    pObtieneCtasBancariasPrevias
    pHabilitaFechaChequeTrans
End Sub

Private Sub cboBancoSAT_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError

    If KeyAscii = 13 Then
        If Trim(cboBancoSAT.Text) <> "<BANCO EXTRANJERO>" Then
            If cboCuentasPrevias.Enabled Or txtCuentaBancaria.Enabled Or MskFecha.Enabled Or cboBanco.Enabled Or cboTipoCargoBancario.Enabled Then
                SendKeys vbTab
            Else
                If Val(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColReferencia)) = 1 Then
                    txtFolio.SetFocus
                Else
                    txtCantidad.SetFocus
                End If
            End If
        Else
            txtBancoExtranjero.Text = ""
            txtBancoExtranjero.Visible = True
            txtBancoExtranjero.SetFocus
            
            cboBancoSAT.Visible = False
        End If
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboBancoSAT_KeyPress"))
End Sub

Private Sub cboBancoSAT_KeyUp(KeyCode As Integer, Shift As Integer)
    pHabilitaFechaChequeTrans
End Sub

Private Sub cboBancoSAT_Validate(Cancel As Boolean)
    pHabilitaFechaChequeTrans
End Sub

Private Sub cboCuentasPrevias_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError

    If KeyAscii = 13 Then
        If cboCuentasPrevias.ItemData(cboCuentasPrevias.ListIndex) = 0 Then
            txtCuentaBancaria.Visible = True
            txtCuentaBancaria.Enabled = True
            txtCuentaBancaria.SetFocus
            cboCuentasPrevias.Visible = False
            cboCuentasPrevias.Enabled = False
        Else
            If MskFecha.Enabled Or cboBanco.Enabled Or cboTipoCargoBancario.Enabled Then
                SendKeys vbTab
            Else
                If Val(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColReferencia)) = 1 Then
                    txtFolio.SetFocus
                Else
                    txtCantidad.SetFocus
                End If
            End If
        End If
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboCuentasPrevias_KeyPress"))
End Sub

Private Sub cboCuentasPrevias_LostFocus()
    If cboCuentasPrevias.ItemData(cboCuentasPrevias.ListIndex) = 0 Then
        txtCuentaBancaria.Visible = True
        cboCuentasPrevias.Visible = False
    End If
End Sub

Private Sub cboCuentasPrevias_Validate(Cancel As Boolean)
    pHabilitaFechaChequeTrans
End Sub

Private Sub cboTipoCargoBancario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColReferencia)) = 1 Then
            txtFolio.SetFocus
        Else
            txtCantidad.SetFocus
        End If
    End If
End Sub

Private Sub cmdClear_Click()
    Dim X As Integer
    For X = 1 To grdFormas.Rows - 1
        If grdFormas.TextMatrix(X, cintColUsarPinpad2) = "3" Then
            grdFormas.TextMatrix(X, cintColUsarPinpad2) = "1"
            grdFormas.TextMatrix(X, cIntColFolio) = ""
        End If
    Next
End Sub

Private Sub cmdOkPP_Click()
    vl_blnTime = True
End Sub

Private Sub cmdReimprimir_Click()
    pReimprimir "R"
End Sub

Private Sub grdFormasPago_Click()
    pLimpia
    
    If Trim(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColTipo)) = "B" Or Trim(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColTipo)) = "H" Or Trim(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColTipo)) = "T" Then
        If Trim(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColClaveFormaPagoSAT)) = "02" _
            Or Trim(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColClaveFormaPagoSAT)) = "03" _
            Or Trim(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColClaveFormaPagoSAT)) = "04" _
            Or Trim(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColClaveFormaPagoSAT)) = "05" _
            Or Trim(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColClaveFormaPagoSAT)) = "28" _
            Or Trim(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColClaveFormaPagoSAT)) = "29" _
        Then
            lblBanco.Enabled = True
            cboBanco.Enabled = True
        Else
            lblBanco.Enabled = False
            cboBanco.Enabled = False
        End If
    Else
        lblBanco.Enabled = False
        cboBanco.Enabled = False
    End If
    
    phabilitaInfoExtra
    pCargosBancarios
    
     pReceptoraBanco grdFormasPago.TextMatrix(grdFormasPago.Row, 5)
End Sub

Private Sub grdFormasPago_GotFocus()
    pLimpia

    If Trim(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColTipo)) = "B" Or Trim(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColTipo)) = "H" Or Trim(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColTipo)) = "T" Then
        If Trim(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColClaveFormaPagoSAT)) = "02" _
            Or Trim(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColClaveFormaPagoSAT)) = "03" _
            Or Trim(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColClaveFormaPagoSAT)) = "04" _
            Or Trim(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColClaveFormaPagoSAT)) = "05" _
            Or Trim(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColClaveFormaPagoSAT)) = "28" _
            Or Trim(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColClaveFormaPagoSAT)) = "29" _
        Then
            lblBanco.Enabled = True
            cboBanco.Enabled = True
        Else
            lblBanco.Enabled = False
            cboBanco.Enabled = False
        End If
    Else
        lblBanco.Enabled = False
        cboBanco.Enabled = False
    End If
    
    phabilitaInfoExtra
    pCargosBancarios
End Sub

Private Sub grdFormasPago_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyPageDown Or KeyCode = vbKeyPageUp Or (KeyCode = vbKeyEnd And Shift = 2) Then
        pLimpia
        
        If Trim(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColTipo)) = "B" Or Trim(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColTipo)) = "H" Or Trim(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColTipo)) = "T" Then
            If Trim(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColClaveFormaPagoSAT)) = "02" _
                Or Trim(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColClaveFormaPagoSAT)) = "03" _
                Or Trim(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColClaveFormaPagoSAT)) = "04" _
                Or Trim(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColClaveFormaPagoSAT)) = "05" _
                Or Trim(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColClaveFormaPagoSAT)) = "28" _
                Or Trim(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColClaveFormaPagoSAT)) = "29" _
            Then
                lblBanco.Enabled = True
                cboBanco.Enabled = True
            Else
                lblBanco.Enabled = False
                cboBanco.Enabled = False
            End If
        Else
            lblBanco.Enabled = False
            cboBanco.Enabled = False
        End If
    
        phabilitaInfoExtra
        pCargosBancarios
        
        pReceptoraBanco grdFormasPago.TextMatrix(grdFormasPago.Row, 5)
    End If
End Sub

Private Sub grdFormasPago_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With grdFormasPago
            If .Row >= 0 Then
                If Trim(.TextMatrix(.Row, cintColTipo)) = "B" Then
                    SendKeys vbTab
                Else
                    If vlblnLicenciaContaElectronica Or vgblnCapturaraDatosBanco Then
                        If txtRFC.Enabled Then
                            txtRFC.SetFocus
                        Else
                            If Val(.TextMatrix(.Row, cintColReferencia)) = 1 Then
                                txtFolio.SetFocus
                            Else
                                txtCantidad.SetFocus
                            End If
                        End If
                    Else
                        If cboBanco.Enabled Or cboTipoCargoBancario.Enabled Then
                            SendKeys vbTab
                        Else
                            If Val(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColReferencia)) = 1 Then
                                txtFolio.SetFocus
                            Else
                                txtCantidad.SetFocus
                            End If
                        End If
                    End If
                End If
            End If
        End With
    End If
End Sub

Private Sub pLimpia()
    txtFolio.Text = ""
    txtCantidad.Text = txtDiferencia.Text
    cboBanco.ListIndex = -1
    cboTipoCargoBancario.ListIndex = -1
    
    If grdFormasPago.Row >= 0 Then
        
        If Trim(txtDiferencia.Text) <> "" Then
            If Val(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColMoneda)) = 0 And vlblnPesos Then
                vldblValorSinFormato = Round(Val(Format(txtDiferencia.Text, "############.00")) / vldblTipoCambioDia, 2)
                vldblValorConDecimales = Val(Format(txtDiferencia.Text, "############.00")) / vldblTipoCambioDia
                txtCantidad.Text = Format(vldblValorSinFormato, "###,###,###,###.00")
            Else
                vldblValorSinFormato = Round(Val(Format(txtDiferencia.Text, "############.00")), 2)
                vldblValorConDecimales = Val(Format(txtDiferencia.Text, "############.00"))
            End If
        Else
            vldblValorSinFormato = 0
            vldblValorConDecimales = 0
        End If
    End If
End Sub


Public Sub cmdAceptar_Click()
    Dim X As Integer
    Dim rs As New ADODB.Recordset
    Dim blnAllOk As Boolean
    Dim blnPinPadOK As Boolean
    Dim strResultadoPinpad As String
    Dim arrDatosPinPad() As String
    Dim tmpData As String
    Dim intPpProv As Integer
    Dim blnImprimeVoucher As Boolean
    
    blnAllOk = False
    blnPinPadOK = True
    blnImprimeVoucher = False
    
    'El paciente tiene registro de crédito
    If vlblnRegistroCredito Then
        For X = 1 To grdFormas.Rows - 1
            If grdFormas.RowData(X) = vlintFormaCredito Then
                If grdFormas.TextMatrix(X, cintColCantidad) = Format(str(vldblCantidadCredito), "###,###,###,###.00") Then
                    vlstrsSQL = "Update TsRegistroCredito "
                    vlstrsSQL = vlstrsSQL & " set chrEstatus = 'A'"
                    vlstrsSQL = vlstrsSQL & " where intnumregistro = " & vlintNumRegistro
                    pEjecutaSentencia vlstrsSQL
                    
                    blnAllOk = True
                    'Hide
                    Exit For
                Else
                    'La cantidad a crédito debe ser igual a la que registró Trabajo social
                    MsgBox SIHOMsg(777), vbOKOnly + vbExclamation, "Mensaje"
                    Exit For
                End If
            End If
        Next
    Else
        blnAllOk = True
        
        'Hide
    End If
    
    If vlblnLicenciaContaElectronica Or vgblnCapturaraDatosBanco Then
        For X = 1 To grdFormas.Rows - 1
            If Trim(grdFormas.TextMatrix(X, cIntColRFC)) <> "ACO560518KW7" And Trim(grdFormas.TextMatrix(X, cIntColRFC)) <> "" And Trim(grdFormas.TextMatrix(X, cintColBancoSAT)) <> "000" Then
                Set rs = frsRegresaRs("SELECT CHRRFC FROM PVRFCCTABANCOSAT WHERE TRIM(CHRRFC) = '" & Trim(grdFormas.TextMatrix(X, cIntColRFC)) & "' AND TRIM(CHRCLAVEBANCOSAT) = '" & Trim(grdFormas.TextMatrix(X, cintColBancoSAT)) & "' AND TRIM(VCHCUENTABANCARIA) = '" & Trim(grdFormas.TextMatrix(X, cIntColCuentaBancaria)) & "'", adLockReadOnly, adOpenForwardOnly)
                If rs.RecordCount = 0 Then
                    pEjecutaSentencia "INSERT INTO PVRFCCTABANCOSAT (CHRRFC,CHRCLAVEBANCOSAT,VCHCUENTABANCARIA) VALUES ('" & Trim(grdFormas.TextMatrix(X, cIntColRFC)) & "','" & Trim(grdFormas.TextMatrix(X, cintColBancoSAT)) & "','" & Trim(grdFormas.TextMatrix(X, cIntColCuentaBancaria)) & "')"
                End If
            End If
        Next
    End If
    Dim booleo As Boolean
    
    If blnAllOk Then
        For X = 1 To grdFormas.Rows - 1
            If grdFormas.TextMatrix(X, cintColUsarPinpad2) = "3" Then
                MsgBox "Existen operaciones pendientes de resolver con el Pinpad." & vbCrLf & "Si fueron aprobadas presione el botón Reimprimir, de lo contrario presione Reiniciar para intentar de nuevo.", vbExclamation, "Mensaje"
                blnPinPadOK = False
                Exit For
            End If
            If grdFormas.TextMatrix(X, cintColUsarPinpad2) = "1" Then
                intPpProv = 1 'IIf(grdFormas.TextMatrix(X, cintColPpProv2) = "", 1, CInt(grdFormas.TextMatrix(X, cintColPpProv2)))
                fraPinPad.Height = 8550
                lblDeclinada.Visible = False
                lblAprobada.Visible = False
                lblErr.Visible = False
                cmdOkPP.Visible = False
                fraPinPad.Visible = True
            
               
                '"04TRN01|AMT12000|TIP0|TAM0|CUR484|PRV[Signature:3][EMVARQC:53D52317BAF79BF9][EMVAPPNAME:Debit Mastercard][EMVAID:A0000000041010]|TID53836-0000040|TRS40|AUC538197|IRC000|IMTN|QPS0|MCN510125******7681|ACNC|CBNMASTERCARD|CPNMASTERCARD|ISUMASTER CARD|AFF1463648|ERR00|RBMFFFFFDFFFFFFFF|MDLP400|VOS31343300|VSN806-674-014|VPN67116328|ADQFirstData-MX|TSIE800|RQC53D52317BAF79BF9|ANMDebit Mastercard|AIDA0000000041010|FSM84196280002|FST1947706|FSA2222286|TDT26/11/2024 21:24:50|BIN510125|"
                'strResultadoPinpad = "04TRN01|AMT12000|TIP0|TAM0|CUR484|NOTCANCELO USUARIO|ERRE2|RBMFFFFFDFFFFFFFF|MDLP400|VOS31343300|VSN806-674-014|VPN67116328|"
                strResultadoPinpad = fstrPinPad(grdFormas.TextMatrix(X, cintColUriPinpad2), CStr(CDbl(grdFormas.TextMatrix(X, cintColCantidad)) * 100), IIf(grdFormas.TextMatrix(X, cintColCveMoneda2) = "", "484", grdFormas.TextMatrix(X, cintColCveMoneda2)), intPpProv, grdFormas.TextMatrix(X, cintColPpHost2), grdFormas.TextMatrix(X, cintColPpPort2), grdFormas.TextMatrix(X, cintColPpUsr2), grdFormas.TextMatrix(X, cintColPpPwd2), CLng(grdFormas.TextMatrix(X, cintColPpCve2)), grdFormas.TextMatrix(X, cintColPpTId2), grdFormas.RowData(X), "")
                If strResultadoPinpad <> "" Then
                   ' arrDatosPinPad = Split(strResultadoPinpad, "|")
                    If intPpProv = 1 Then
                        tmpData = ObtenerValorVariable(strResultadoPinpad, "IRC")
                        If tmpData = "000" Or tmpData = "002" Or tmpData = "003" Then
                        
                            tmpData = ObtenerValorVariable(strResultadoPinpad, "ERR")
                            
                            If tmpData = "00" Then
                                lblAprobada.Visible = True
                                grdFormas.TextMatrix(X, cintColUsarPinpad2) = "2"
                                grdFormas.TextMatrix(X, cIntColFolio) = ObtenerValorVariable(strResultadoPinpad, "AUC")
                            Else
                                lblDeclinada.Visible = True
                                blnPinPadOK = False
                            End If
                            blnImprimeVoucher = True
                            pMsgWait False
                        Else
                            
                            
                            Select Case tmpData
                                Case "01": lblErr.Caption = "Error en lectura de tarjeta"
                                Case "02": lblErr.Caption = "Tarjeta retirada"
                                Case "942": lblErr.Caption = "Operación cancelada"
                                Case "05": lblErr.Caption = "Tiempo de espera de tarjeta excedido"
                                Case "06": lblErr.Caption = "Error en datos de tarjeta"
                                Case "07": lblErr.Caption = "Validación de últ 4 dígitos no exitosa"
                                Case "08": lblErr.Caption = "Error en datos o formato de mensaje"
                                Case "09": lblErr.Caption = "Comando no reconocido"
                                Case "10": lblErr.Caption = "Falta carga de llave para cifrado"
                                Case "11": lblErr.Caption = "Error al cifrar datos"
                                Case "12": lblErr.Caption = "Aplicación bloqueada en tarjeta"
                                Case "033": lblErr.Caption = "Tarjeta expirada"
                                Case "16": lblErr.Caption = "Código de moneda no soportado"
                                Case "51": lblErr.Caption = "Terminal no configurada"
                                Case "52": lblErr.Caption = "Terminal no pudo conectarse a host"
                                Case "53": lblErr.Caption = "Terminal no recibió respuesta de host"
                                Case Else: lblErr.Caption = ObtenerValorVariable(strResultadoPinpad, "NOT")
                            End Select
                            lblErr.Visible = True
                            cmdOkPP.Visible = True
                            cmdOkPP.SetFocus
                            blnPinPadOK = False
                            pMsgWait True
                        End If
                    ElseIf intPpProv = 2 Then
                         tmpData = fstrGetPPData(arrDatosPinPad, "trn_internal_respcode")
                         If tmpData = "-1" Then
                            lblAprobada.Visible = True
                            grdFormas.TextMatrix(X, cintColUsarPinpad2) = "2"
                            grdFormas.TextMatrix(X, cIntColFolio) = fstrGetPPData(arrDatosPinPad, "trn_auth_code")
                            blnImprimeVoucher = True
                            pMsgWait False
                         Else
                            If tmpData = "51" Then
                                lblDeclinada.Visible = True
                                blnImprimeVoucher = False
                                
                            End If
                            tmpData = fstrGetPPData(arrDatosPinPad, "trn_msg_host")
                            lblErr.Caption = tmpData
                            lblErr.Visible = True
                            cmdOkPP.Visible = True
                            cmdOkPP.SetFocus
                            blnPinPadOK = False
                            pMsgWait True
                         End If
                         
                       'santander
                       ElseIf intPpProv = 3 Then
                       tmpData = arrDatosPinPad(0)
                       If tmpData = "approved" Then
                         lblAprobada.Visible = True
                            grdFormas.TextMatrix(X, cintColUsarPinpad2) = "2"
                         grdFormas.TextMatrix(X, cIntColFolio) = fstrGetPPData(arrDatosPinPad, "NOperacion")
                           blnImprimeVoucher = True
                           blnPinPadOK = True
                            pMsgWait False
                            Else 'declinado
                                If tmpData = "denied" Then
                                lblDeclinada.Visible = True
                                blnImprimeVoucher = False
                                cmdOkPP.Visible = True
                                 cmdOkPP.SetFocus
                                 pMsgWait True
                                End If
                                
                                If tmpData = "error" Then
                                
                                tmpData = fstrGetPPData(arrDatosPinPad, "DescError")
                                lblErr.Caption = tmpData
                                lblErr.Visible = True
                                cmdOkPP.Visible = True
                                 cmdOkPP.SetFocus
                                blnPinPadOK = False
                                pMsgWait True
                                End If
                                
                                
                            
                            
                         End If
                         
                    End If
                  
                    If blnImprimeVoucher And Trim(grdFormas.TextMatrix(X, cintColImprVoucher2)) <> "" Then
                        
                        If intPpProv = 1 Then
                        pImprimeVoucher2 grdFormas.TextMatrix(X, cintColImprVoucher2), strResultadoPinpad, intPpProv
                        
                       Else
                        
                        
                        pImprimeVoucher grdFormas.TextMatrix(X, cintColImprVoucher2), arrDatosPinPad, intPpProv
                        
                        End If
                        
                    End If
                    pGuardarLogTransaccion "frmPagoPos", EnmPinPad, vllngCveEmpleado, "PINPADRESPONSE", strResultadoPinpad
                    
                Else
                    blnPinPadOK = False
                    If vlblnTimeout Then
                        pGuardarLogTransaccion "frmPagoPos", EnmPinPad, vllngCveEmpleado, "PINPADRESPONSE", "TIMEOUT"
                        grdFormas.TextMatrix(X, cintColUsarPinpad2) = "3"
                        grdFormas.TextMatrix(X, cIntColFolio) = vlstrRef
                    End If
                End If
                fraPinPad.Visible = False
                
                If Not blnPinPadOK Then Exit For
            End If
        Next
        If blnPinPadOK Then
            vgintSalidaOK = 1 ' Que si se pudo hacer
            If Me.Visible = True Then
                Hide
            Else
            
            End If
        End If
    End If
    
End Sub
Function ObtenerValorVariable(ByVal Cadena As String, ByVal clave As String) As String
    Dim partes() As String
    Dim i As Integer
    Dim claveConDelimitador As String
    Dim Resultado As String
    
    ' Definir la clave con el delimitador
    claveConDelimitador = clave & ""
    
    ' Dividir la cadena en partes usando "|" como delimitador
    partes = Split(Cadena, "|")
    
    ' Recorrer las partes buscando la clave
    For i = LBound(partes) To UBound(partes)
        If Left(partes(i), Len(claveConDelimitador)) = claveConDelimitador Then
            ' Extraer el valor después de la clave
            Resultado = Mid(partes(i), Len(claveConDelimitador) + 1)
            ObtenerValorVariable = Resultado
            Exit Function
        End If
    Next i
    
    ' Si no se encuentra la clave, devolver una cadena vacía
    ObtenerValorVariable = ""
End Function

Private Function fstrGetPPData(arrDatos() As String, strNombre As String) As String
    On Error GoTo Errs
    Dim intIndex As Integer
    For intIndex = 0 To UBound(arrDatos)
         If strNombre = Split(arrDatos(intIndex), "=")(0) Then
             fstrGetPPData = Replace(Split(arrDatos(intIndex), "=")(1), "_", " ")
             Exit Function
         End If
    Next
Errs:
    fstrGetPPData = ""
End Function
Private Function fstrGetFirma(strNombre As String) As String
   
  
If strNombre = "AutorizadosinFirma" Then
             fstrGetFirma = "AUTORIZAD SIN FIRMA"
Else
    If strNombre = "VALIDADOCONFIRMAELECTRONICA" Then
              fstrGetFirma = "VALIDADO CON FIRMA ELECTRONICA"
              Else
              fstrGetFirma = "*"
End If
End If


End Function

Private Function fstrPinPad(strUriPinpad As String, strCantidad As String, strMoneda As String, intPpProvider As Integer, strHost As String, strPort As String, strUsr As String, strPwd As String, lngCve As Long, strTermId As String, lngCveFormaPago As Long, strMSI As String) As String
    On Error GoTo Errs
    Dim intRespLen As Long
    Dim strReturn As String
    Dim lngIdLog As Long
    Dim strRef As String
    
    
    vlblnTimeout = False
    Set ws = New WebSocketWrap.Client
    ws.Timeout = 3000
    
    If intPpProvider = 2 Then
    
    ws.Uri = strUriPinpad & "?host=" & strHost & "&port=" & strPort & "&prov=" & IIf(intPpProvider = 1, "FISERV", "EVO") & "&usr=" & strUsr & "&pwd=" & strPwd
 Else ' snatander, firve
If strUriPinpad <> "" And Conectado = False Then

 Winsock1.Connect strUriPinpad, strPort  'Conectamos el winsock
 If Winsock1.State = sckConnected Then
        
        ' Puedes realizar acciones adicionales aquí después de una conexión exitosa
    Else
   
    
       strRespuestaEsperar = "Error de socket"
    End If
 End If
 
    End If
 

    'blnRespuestaEsperar = False
    'ws.SendMessage "INIT:" & strHost & ":" & strPort & ":" & IIf(intPpProvider = 1, "FISERV", "EVO") & ":" & strUsr & ":" & strPwd
    'Do While Not blnRespuestaEsperar
    '    DoEvents
    'Loop
    blnRespuestaEsperar = False
    lngIdLog = pRegLog(0, "", lngCve, "N", lngCveFormaPago)
    Select Case intPpProvider
        Case 1
            lngIdLog = pRegLog(lngIdLog, "TR0_01:" & strMoneda & ":" & strCantidad, lngCve, "T", lngCveFormaPago)
             If Conectado Then Winsock1.SendData Trim("TRN01|AMT" & Replace(strCantidad, ".", "") & "|CUR" & 484 & "|")
        Case 2
            If strTermId <> "" Then
                strRef = strTermId & Mid(CStr(1000000000 + lngIdLog), 2, 9)
            End If
            lngIdLog = pRegLog(lngIdLog, "T060S000:" & strMoneda & ":" & CDbl(strCantidad) / 100 & ":" & IIf(strRef = "", "-", strRef) & ":" & IIf(strMSI = "", "-", strMSI), lngCve, "T", lngCveFormaPago)
            ws.SendMessage "T060S000:" & strMoneda & ":" & CDbl(strCantidad) / 100 & ":" & IIf(strRef = "", "-", strRef) & ":" & IIf(strMSI = "", "-", strMSI)
            
        Case 3
          If strTermId <> "" Then
                strRef = strTermId & Mid(CStr(1000000000 + lngIdLog), 2, 9)
            End If
         lngIdLog = pRegLog(lngIdLog, "INIPAGO:" & strMoneda & ":" & CDbl(strCantidad) / 100 & ":" & IIf(strRef = "", Mid(CStr(1000000000 + lngIdLog), 2, 9) & strRef, strRef) & ":" & IIf(strMSI = "", "-", strMSI), lngCve, "T", lngCveFormaPago)
        If Conectado Then Winsock1.SendData Trim("INIPAGO|" & "MXN|" & CDbl(strCantidad) / 100 & "|" & "11|" & Mid(CStr(1000000000 + lngIdLog), 2, 9) & strRef & "|*")
        If Conectado = False Then blnRespuestaEsperar = True
  
    
    
    End Select
    


    Do While Not blnRespuestaEsperar
        DoEvents
    Loop
    strReturn = strRespuestaEsperar

   
    
    If intPpProvider = 1 Then
        strReturn = Trim(strReturn)
    ElseIf intPpProvider = 2 Then
        intRespLen = InStr(strReturn, "}") - 14
        strReturn = Mid(strReturn, 13, intRespLen)
        strReturn = Replace(strReturn, "|Respuesta=", "")
        strReturn = Replace(strReturn, "&", "|")
        
        ElseIf intPpProvider = 3 Then
        strReturn = Trim(strReturn)  'santander
        
    End If
  
    pRegLog lngIdLog, strReturn, lngCve, "R", lngCveFormaPago
    If InStr(strReturn, "Error de socket") > 0 Then
        MsgBox "Error de conexión con el socket:" & vbCrLf & strUriPinpad & ":" & strPort, vbExclamation, "Mensaje"
        fstrPinPad = ""
    Else
        fstrPinPad = strReturn
    End If
    
  


    
    Exit Function
Errs:
    fstrPinPad = ""
    If InStr(Err.Description, "Error de conexión") > 0 Then
        pRegLog lngIdLog, "Error de conexión con el Web Socket: " & strUriPinpad, lngCve, "R", lngCveFormaPago
        MsgBox "Error de conexión con el Web Socket: " & vbCrLf & strUriPinpad, vbExclamation, "Mensaje"
    Else
        pRegLog lngIdLog, Err.Description, lngCve, "R", lngCveFormaPago
        If Err.Number = 555 And Err.Description = "Timeout" Then
            vlblnTimeout = True
            vlstrRef = strRef
            MsgBox "El tiempo de espera terminó." & vbCrLf & "Si la transacción fue aprobada en el Pinpad presione el botón Reimprimir, de lo contrario presione Reiniciar para intentar de nuevo.", vbExclamation, "Mensaje"
        Else
            MsgBox Err.Description, vbExclamation, "Mensaje"
        End If
    End If
End Function

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

Dim Dato As String
Conectado = True
Winsock1.GetData strRespuestaEsperar
blnRespuestaEsperar = True
End Sub
Private Sub Timer1_Timer()
'Si el winsock está conectado, cambiamos la variable a true
DoEvents
If Winsock1.State <> sckConnected Then
Conectado = False
Else
Conectado = True
End If
End Sub
Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'En caso de error cerramos la conexión
Winsock1.Close
End Sub

Private Sub pImprimeVoucher(strImpresora As String, arrDatos() As String, intPpProv As Integer)
    
    On Error GoTo Errs
    Dim vgrptReporte As CRAXDRT.Report
    Dim rsDummy As ADODB.Recordset
    Dim alstrParametros(22) As String
    Dim tmpData As String
    Dim strModo As String
    Dim strFecha As String
    Dim strHora As String
    Dim strFechaHora As String
    Dim blnAprobado As Boolean
    Dim intFirma As Integer
    
    
    Set rsDummy = frsRegresaRs("select sysdate from dual", adLockReadOnly, adOpenForwardOnly)
    If intPpProv = 1 Then
        tmpData = ObtenerValorVariable(strResultadoPinpad, "IRC")
        blnAprobado = IIf(Mid(tmpData, 1, 2) = "00", True, False)
        
       
        strFechaHora = fstrGetPPData(arrDatos, "TDT")
        strFecha = Format(Date, "DDMMYY")
        strHora = Format(Time, "HHMMSS")
        alstrParametros(0) = "NombreComercio; " & fstrGetPPData(arrDatos, "COMERCIO")
        alstrParametros(1) = "Afiliacion;" & fstrGetPPData(arrDatos, "PROSA_BDU")
        alstrParametros(2) = "MerchID;" & fstrGetPPData(arrDatos, "MERCH_ID_SOUTH")
        alstrParametros(3) = "TerminalID;" & fstrGetPPData(arrDatos, "TERMINAL_ID")
        alstrParametros(4) = "Fecha;" & strFecha
        alstrParametros(5) = "Hora;" & strHora
        alstrParametros(6) = "Mensaje1;" & Trim(vgstrDireccionCH) & ", " & Trim(vgstrColoniaCH)
        alstrParametros(7) = "Mensaje2;" & vgstrCiudadCH & ", " & vgstrEstadoCH
        alstrParametros(8) = "PANTarjeta;" & fstrGetPPData(arrDatos, "PAN")
        alstrParametros(9) = "Autorizacion;" & fstrGetPPData(arrDatos, "NUM_AUTORIZACION")
        alstrParametros(10) = "Emisor;" & fstrGetPPData(arrDatos, "INFO_EMISOR")
        alstrParametros(11) = "EtiquetaAplicacion;" & fstrGetPPData(arrDatos, "EMV_APP_LABEL")
        alstrParametros(12) = "Total;$" & Format(CDbl(fstrGetPPData(arrDatos, "MONTO_TOTAL")) / 100, "0.00")
        alstrParametros(13) = "Proceso;" & IIf(Mid(tmpData, 2, 1) = "0", "APROBADA FUERA DE LINEA", "APROBADA EN LINEA")
        alstrParametros(14) = "ModoIngreso;" & strModo
        alstrParametros(15) = "ARQC;" & fstrGetPPData(arrDatos, "EMV_ARQC")
        alstrParametros(16) = "AID;" & fstrGetPPData(arrDatos, "EMV_AID")
        alstrParametros(17) = "TC;" & fstrGetPPData(arrDatos, "EMV_TC")
        alstrParametros(18) = "NombreCliente;" & Trim(fstrGetPPData(arrDatos, "TARJETAHABIENTE"))
        alstrParametros(19) = "Promocion;" & IIf(fstrGetPPData(arrDatos, "PROMOCION") = "-", "0", fstrGetPPData(arrDatos, "PROMOCION"))
        alstrParametros(20) = "Copia;0"
        alstrParametros(21) = "Firma;" & Mid(tmpData, 5, 1)
        alstrParametros(22) = "Operacion;" & IIf(fstrGetPPData(arrDatos, "OPERACION") = "01", "VENTA", "DEVOLUCION")
    ElseIf intPpProv = 2 Then
        
        blnAprobado = True
        tmpData = fstrGetPPData(arrDatos, "trn_qty_pay")
        
        'aqui esta lo nuevo en EVO
        
        
        Select Case CInt(fstrGetPPData(arrDatos, "trn_input_mode"))
         Case 0, 1
            strModo = "BANDA MAGNETICA"
            'LA FIRMA ES FIRMA:________________
            intFirma = 1
            alstrParametros(0) = "Firma;1"
             alstrParametros(20) = "EtiquetaAplicacion;" & fstrGetPPData(arrDatos, "trn_fii")
        Case 2
            strModo = "DIGITADO EN PINPAD"
            'LA FIRMA ES FIRMA______________
            
            intFirma = 1
            alstrParametros(0) = "Firma;1"
             alstrParametros(20) = "EtiquetaAplicacion;" & fstrGetPPData(arrDatos, "trn_fii")
        Case 5
            strModo = "CONTACTLES"
            'LA FIRMA ES FII: + TRN_FII
            alstrParametros(20) = "EtiquetaAplicacion;" & fstrGetPPData(arrDatos, "trn_fii")
            If CLng(fstrGetPPData(arrDatos, "trn_amount")) < 250 Then
            
            firma = 2
            'firma=2 es igual ha autorizado sin firma
             alstrParametros(0) = "Firma;2"
              alstrParametros(20) = "EtiquetaAplicacion;" & fstrGetPPData(arrDatos, "trn_fii")
            Else
            intFirma = 4
            'FIRMA_________________
             alstrParametros(0) = "Firma;4"
              alstrParametros(20) = "EtiquetaAplicacion;" & fstrGetPPData(arrDatos, "trn_fii")
            End If
            
        Case 7
           strModo = "CHIP"
           If CInt(fstrGetPPData(arrDatos, "trn_fe")) = 1 Then
           'SI EL BIT TRN_FE=1 ENTONCES AUTORIZACION CON FIRMA ELECTRONICA
           'SI ES DISTINTO A 1 ENTONCES FIRMA_________________
           intFirma = 3
            alstrParametros(0) = "Firma;3"
             alstrParametros(20) = "EtiquetaAplicacion;" & fstrGetPPData(arrDatos, "trn_fii")
           Else
           intFirma = 1
            alstrParametros(0) = "Firma;1"
             alstrParametros(20) = "EtiquetaAplicacion;" & fstrGetPPData(arrDatos, "trn_fii")
           
           End If
        
        End Select
       
        alstrParametros(1) = "NombreComercio;" & fstrGetPPData(arrDatos, "mer_legend1")
        alstrParametros(2) = "Mensaje1;" & fstrGetPPData(arrDatos, "mer_legend2")
        alstrParametros(3) = "Mensaje2;" & fstrGetPPData(arrDatos, "mer_legend3")
        alstrParametros(4) = "Afiliacion;" & fstrGetPPData(arrDatos, "trn_external_mer_id")
        alstrParametros(5) = "TerminalID;" & fstrGetPPData(arrDatos, "trn_external_ter_id")
        alstrParametros(6) = "Fecha;" & fstrGetPPData(arrDatos, "trn_fechaTrans")
        alstrParametros(7) = "Copia;0"
        alstrParametros(8) = "trn_label;" & fstrGetPPData(arrDatos, "trn_label")
        alstrParametros(9) = "trn_aprnam;" & fstrGetPPData(arrDatos, "trn_aprnam")
        alstrParametros(10) = "ARQC;" & fstrGetPPData(arrDatos, "trn_emv_cryptogram")
        alstrParametros(11) = "AID;" & fstrGetPPData(arrDatos, "trn_AID")
        alstrParametros(12) = "Emisor;" & fstrGetPPData(arrDatos, "trn_pro_name")
        alstrParametros(13) = "PANTarjeta;" & fstrGetPPData(arrDatos, "trn_aco_id")
        alstrParametros(14) = "Autorizacion;" & fstrGetPPData(arrDatos, "trn_auth_code")
        alstrParametros(15) = "TC;" & fstrGetPPData(arrDatos, "trn_id")
        alstrParametros(16) = "Total;" & fstrGetPPData(arrDatos, "trn_amount")
        alstrParametros(17) = "Operacion;" & IIf(tmpData = "1", "COMPRA NORMAL", tmpData & " MESES SIN INTERESES")
        alstrParametros(18) = "trn_fe;" & fstrGetPPData(arrDatos, "trn_fe")
        alstrParametros(19) = "NombreCliente;" & fstrGetPPData(arrDatos, "trn_internal_ter_id")
   ElseIf intPpProv = 3 Then
   
    tmpData = arrDatos(0) 'aprobado o declinado
    
    blnAprobado = True
    alstrParametros(0) = "Firma;1"
   alstrParametros(1) = "NombreComercio;" & fstrGetPPData(arrDatos, "cadena4")
        alstrParametros(2) = "Mensaje1;" & fstrGetPPData(arrDatos, "cadena5")
        alstrParametros(3) = "Mensaje2;" & fstrGetPPData(arrDatos, "cadena6")
        alstrParametros(4) = "Mensaje3;" & fstrGetPPData(arrDatos, "cadena7")
        alstrParametros(5) = "trn_label;" & fstrGetPPData(arrDatos, "cadena9")
        alstrParametros(6) = "PANTarjeta;" & fstrGetPPData(arrDatos, "cadena13")
        alstrParametros(7) = "Copia;0"
        alstrParametros(8) = "MerchID;" & fstrGetPPData(arrDatos, "cadena15")
        alstrParametros(9) = "tipoVocher;" & fstrGetPPData(arrDatos, "cadena17")
        alstrParametros(10) = "Total;" & fstrGetPPData(arrDatos, "cadena23")
        alstrParametros(11) = "Operacion;" & fstrGetPPData(arrDatos, "cadena24")
        alstrParametros(12) = "REF;" & fstrGetPPData(arrDatos, "cadena25")
        alstrParametros(13) = "ARQC;" & fstrGetPPData(arrDatos, "cadena26")
        alstrParametros(14) = "AID;" & fstrGetPPData(arrDatos, "cadena27")
        alstrParametros(15) = "trn_aprnam;" & fstrGetPPData(arrDatos, "cadena28")
        alstrParametros(16) = "Total;" & fstrGetPPData(arrDatos, "trn_amount")
        alstrParametros(18) = "Fecha;" & fstrGetPPData(arrDatos, "cadena33")
        alstrParametros(19) = "NombreCliente;" & fstrGetPPData(arrDatos, "cadena40")
         alstrParametros(20) = "BarrasCode;" & fstrGetPPData(arrDatos, "cadena48")
          alstrParametros(21) = "Autorizacion;" & fstrGetFirma(fstrGetPPData(arrDatos, "cadena39"))
   
   
   
   
    End If
    If blnAprobado Then
    If intPpProv = 2 Then
    pInstanciaReporte vgrptReporte, "voucher4.rpt"
    Else
    If intPpProv = 3 Then
     pInstanciaReporte vgrptReporte, "vouchersantander.rpt"
    
    End If
    If intPpProv = 1 Then
        pInstanciaReporte vgrptReporte, "voucher.rpt"
        End If
        
        End If
        
    Else
        pInstanciaReporte vgrptReporte, "voucherdec6.rpt"
    End If
    vgrptReporte.DiscardSavedData
    
    fblnAsignaImpresoraReportePorNombre strImpresora, vgrptReporte
    
    pCargaParameterFields alstrParametros, vgrptReporte
    pImprimeReporte vgrptReporte, rsDummy, "I", "Voucher venta"
    
    If blnAprobado Then
        If intPpProv = 1 Then
            alstrParametros(20) = "Copia;1"
        ElseIf intPpProv = 2 Then
            alstrParametros(7) = "Copia;1"
         ElseIf intPpProv = 3 Then
             alstrParametros(9) = "tipoVocher;" & "-C-L-I-E-N-T-E-"
        End If
        pCargaParameterFields alstrParametros, vgrptReporte
        pImprimeReporte vgrptReporte, rsDummy, "I", "Voucher venta"
    End If
    
    Exit Sub
Errs:
    MsgBox Err.Description, vbExclamation, "Mensaje"
End Sub

Private Sub pImprimeVoucher2(strImpresora As String, arrDatos As String, intPpProv As Integer)
    
    On Error GoTo Errs
    Dim vgrptReporte As CRAXDRT.Report
    Dim rsDummy As ADODB.Recordset
    Dim alstrParametros(22) As String
    Dim tmpData As String
    Dim strModo As String
    Dim strFecha As String
    Dim strHora As String
    Dim strFechaHora As String
    Dim blnAprobado As Boolean
    Dim intFirma As Integer
    
    
    Set rsDummy = frsRegresaRs("select sysdate from dual", adLockReadOnly, adOpenForwardOnly)
   
        tmpData = ObtenerValorVariable(arrDatos, "IRC")
        blnAprobado = IIf(tmpData = "000", True, False)
        
       
        strFechaHora = ObtenerValorVariable(arrDatos, "TDT")
        strFechaHora = Format(Date, "DD/MM/YYYY")
        strFecha = Format(Time, "HH:MM:SS")
        alstrParametros(0) = "NombreComercio; " & "COMERCIO PRUEBA"
        alstrParametros(1) = "Afiliacion;" & ObtenerValorVariable(arrDatos, "FSA")
        alstrParametros(2) = "MerchID;" & ObtenerValorVariable(arrDatos, "FSM")
        alstrParametros(3) = "TerminalID;" & ObtenerValorVariable(arrDatos, "FST")
        alstrParametros(4) = "Fecha;" & strFechaHora
        alstrParametros(5) = "Hora;" & strFecha
        alstrParametros(6) = "Mensaje1;" & Trim(vgstrDireccionCH) & ", " & Trim(vgstrColoniaCH)
        alstrParametros(7) = "Mensaje2;" & vgstrCiudadCH & ", " & vgstrEstadoCH
        alstrParametros(8) = "PANTarjeta;" & ObtenerValorVariable(arrDatos, "MCN")
        alstrParametros(9) = "Autorizacion;" & ObtenerValorVariable(arrDatos, "AUC")
        alstrParametros(10) = "Emisor;" & ObtenerValorVariable(arrDatos, "ANM")
        alstrParametros(11) = "EtiquetaAplicacion;" & ObtenerValorVariable(arrDatos, "CPN")
        alstrParametros(12) = "Total;$" & Format(CDbl(ObtenerValorVariable(arrDatos, "AMT")) / 100, "0.00")
        alstrParametros(13) = "Proceso;" & IIf(MtmpData = "00", "APROBADA FUERA DE LINEA", "APROBADA EN LINEA")
        alstrParametros(14) = "ModoIngreso;" & strModo
        alstrParametros(15) = "ARQC;" & ObtenerValorEntreDelimitadores(ObtenerValorVariable(arrDatos, "PRV"), "EMVARQC")
        alstrParametros(16) = "AID;" & ObtenerValorEntreDelimitadores(ObtenerValorVariable(arrDatos, "PRV"), "EMVAID")
        alstrParametros(17) = "TC;" & ObtenerValorVariable(arrDatos, "TSI")
        alstrParametros(18) = "NombreCliente;" & " "
        alstrParametros(19) = "Promocion;" & "  "
        alstrParametros(20) = "Copia;0"
        alstrParametros(21) = "Firma;" & Mid(tmpData, 5, 1)
        alstrParametros(22) = "Operacion;" & "VENTA"
    

    If intPpProv = 1 Then
        pInstanciaReporte vgrptReporte, "voucher.rpt"
    

        
    Else
        pInstanciaReporte vgrptReporte, "voucherdec6.rpt"
    End If
    vgrptReporte.DiscardSavedData
    
    fblnAsignaImpresoraReportePorNombre strImpresora, vgrptReporte
    
    pCargaParameterFields alstrParametros, vgrptReporte
    pImprimeReporte vgrptReporte, rsDummy, "I", "Voucher venta"
    
    If blnAprobado Then
        If intPpProv = 1 Then
            alstrParametros(20) = "Copia;1"
        ElseIf intPpProv = 2 Then
            alstrParametros(7) = "Copia;1"
         ElseIf intPpProv = 3 Then
             alstrParametros(9) = "tipoVocher;" & "-C-L-I-E-N-T-E-"
        End If
        pCargaParameterFields alstrParametros, vgrptReporte
        pImprimeReporte vgrptReporte, rsDummy, "I", "Voucher venta"
    End If
    
    Exit Sub
Errs:
    MsgBox Err.Description, vbExclamation, "Mensaje"
End Sub







Private Sub pMsgWait(blnUntilOk As Boolean)
    vl_blnTime = False
    If Not blnUntilOk Then
        tmrMsgPP.Enabled = True
    End If
    Do Until vl_blnTime
        DoEvents
    Loop
    tmrMsgPP.Enabled = False
End Sub

Public Sub Form_Activate()
    Dim SQL As String
    Dim rsTemp As New ADODB.Recordset
    
    vgintSalidaOK = 0 'Inicializada
    If vlblnRegistroCredito Then
        If vlblnExisteFormaCredito = False Then
            'No existe una forma de pago tipo crédito para este departamento.
            MsgBox SIHOMsg(622), vbOKOnly + vbExclamation, "Mensaje"
            frmPagoPos.Hide
        Else
            grdFormasPago.Row = fintLocRegMshFGrd(grdFormasPago, str(vlintFormaCredito), cintColIdForma)
            txtCantidad.Text = Format(str(vldblCantidadCredito), "###,###,###,###.00")
            pAgregaForma
        End If
    End If
    
    If rsFormasPago.RecordCount = 0 Then
        'No existen formas de pago para registrar el pago.
        MsgBox SIHOMsg(293), vbOKOnly + vbInformation, "Mensaje"
        frmPagoPos.Hide
    Else
        If vgstrForma = "frmFacturacionDirecta" Then
            SQL = " select intFormaPago,chrDescripcion from PvFormaPago where bitEstatusActivo = 1 and chrTipo = 'C' "
            SQL = SQL & " and smiDepartamento = " & Trim(str(vgintNumeroDepartamento))
            Set rsTemp = frsRegresaRs(SQL, adLockOptimistic, adOpenDynamic)
            If rsTemp.EOF Then
                grdFormasPago.Row = 0
            Else
                vlintFormaCredito = rsTemp.Fields(0)
                grdFormasPago.Row = fintLocMshFGrd(grdFormasPago, Trim(str(vlintFormaCredito)), cintColIdForma)
                grdFormasPago_Click
                If grdFormasPago.Enabled Then If fblnCanFocus(grdFormasPago) Then grdFormasPago.SetFocus
            End If
        Else
            grdFormasPago.Row = 0
        End If
      
        vlblnMuestraCargoBancario = False
        If vgstrForma = "frmFacturacion" Or vgstrForma = "frmEntradaSalidaDinero" Or vgstrForma = "frmPOS" Or vgstrForma = "frmFacturacionDirecta" Or vgstrForma = "frmPagosCredito" Or vgstrForma = "frmPaqueteCobranza" Then
            vlblnMuestraCargoBancario = True
            grdFormasPago_Click
        End If
    End If
End Sub
Function ObtenerValorEntreDelimitadores(ByVal Cadena As String, ByVal clave As String) As String
    Dim inicioClave As Long
    Dim inicioValor As Long
    Dim finValor As Long
    Dim Resultado As String
    
    ' Buscar la posición inicial de la clave
    inicioClave = InStr(Cadena, "[" & clave & ":")
    
    ' Si no encuentra la clave, retornar vacío
    If inicioClave = 0 Then
        ObtenerValorEntreDelimitadores = ""
        Exit Function
    End If
    
    ' Encontrar el inicio del valor después de la clave y el delimitador ":"
    inicioValor = inicioClave + Len(clave) + 2 ' Suma 2 para incluir "[" y ":"
    
    ' Encontrar el cierre del valor antes de "]"
    finValor = InStr(inicioValor, Cadena, "]")
    
    ' Extraer el valor
    If finValor > 0 Then
        Resultado = Mid(Cadena, inicioValor, finValor - inicioValor)
    Else
        Resultado = ""
    End If
    
    ' Retornar el valor extraído
    ObtenerValorEntreDelimitadores = Resultado
End Function

Private Sub pGrid()
    With grdFormas
        .Rows = 2
        .Cols = 30
        .FixedCols = 1
        .FixedRows = 1
        .FormatString = "|Descripción|Folio|Cantidad"
        .ColWidth(cintColCol) = 100
        .ColWidth(cintColDescripcion) = 3700        'Descripcion
        .ColWidth(cIntColFolio) = 3000              'Folio
        .ColWidth(cintColCantidad) = 1800           'Cantidad
        .ColWidth(cintColCtaContableFormaPago) = 0  'Cuenta contable de la forma de pago
        .ColWidth(cintColTipoCambio) = 0            'Tipo de cambio
        .ColWidth(cintColCantidadReal) = 0          'Cantidad real
        .ColWidth(cintColCredito) = 0               'Credito
        .ColWidth(cintColDolares) = 0               'Dolares
        .ColWidth(cintColIdBanco) = 0               'Id del banco
        .ColWidth(cintColMonedaBanco) = 0           'moneda del banco
        .ColWidth(cintColCtaComisionBanc) = 0       'Cuenta contable comisión bancaria
        .ColWidth(cintColComisionBanc) = 0          'Comisión bancaria
        .ColWidth(cintColIVAComisionBanc) = 0       'Iva comisión bancaria
        .ColWidth(cIntColRFC) = 0                   'RFC
        .ColWidth(cintColBancoSAT) = 0              'Clave del banco del SAT
        .ColWidth(cintColBancoExt) = 0              'Descripcion del banco extranjero
        .ColWidth(cIntColCuentaBancaria) = 0        'Cuenta bancaria
        .ColWidth(cintColFechaCqTrans) = 0          'Fecha del cheque o transferencia
        .ColWidth(cintColUsarPinpad2) = 0           'Usar Pinpad
        .ColWidth(cintColUriPinpad2) = 0            'URI del Pinpad
        .ColWidth(cintColImprVoucher2) = 0          'Impresora del voucher
        .ColWidth(cintColCveMoneda2) = 0            'Moneda del Pinpad
        .ColWidth(cintColPpProv2) = 0               'Proveedor del Pinpad
        .ColWidth(cintColPpUsr2) = 0                'Usuario del Pinpad
        .ColWidth(cintColPpPwd2) = 0                'Password del Pinpad
        .ColWidth(cintColPpHost2) = 0               'Host del Pinpad
        .ColWidth(cintColPpPort2) = 0               'Puerto del Pinpad
        .ColWidth(cintColPpCve2) = 0                'Clave del Pinpad
        .ColWidth(cintColPpTId2) = 0                'Id de la terminal
        
        .RowData(1) = -1
    End With
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        pGrid
        Hide
    End If
End Sub

Private Sub Form_Load()
    Dim SQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim vlstrsql As String
    Timer1.Interval = 100
    Set rsTemp = frsSelParametros("SI", vgintClaveEmpresaContable, "INTTERMINALSTIMEOUT")
    If Not rsTemp.EOF Then
        intTimeout = CInt(rsTemp!Valor)
    Else
        intTimeout = 300
    End If
    rsTemp.Close
    
    intModoMasivo = 0
    Me.Icon = frmMenuPrincipal.Icon
    
    strFormaPadre = vgstrForma
    
    vlblnLicenciaContaElectronica = fblnLicenciaContaElectronica
    vldtmFecha = fdtmServerFecha
    vldtmfechaServer = fdtmServerFecha
    
    MskFecha.Text = vldtmFecha
    
    txtRFC.Text = vlstrRFCOriginal
    txtBancoExtranjero.Text = ""
    txtCuentaBancaria.Text = ""
    
    If vlblnLicenciaContaElectronica Or vgblnCapturaraDatosBanco Then
        pCargaBancosSAT
        fraInformacionExtra.Visible = True
        grdFormasPago.Width = 8895
        grdFormasPago.Height = 1455
        fraPrincipales.Top = 1680
        fraAbajo.Top = 4440
        
        frmPagoPos.Refresh
        frmPagoPos.Height = 8940
        frmPagoPos.Top = Int((SysInfo.WorkAreaHeight - frmPagoPos.Height) / 2)
        frmPagoPos.Refresh
    Else
        fraInformacionExtra.Visible = False
        grdFormasPago.Width = 4410
        grdFormasPago.Height = 2670
        fraPrincipales.Top = 120
        fraAbajo.Top = 2880
        
        frmPagoPos.Refresh
        frmPagoPos.Height = 7410
        frmPagoPos.Top = Int((SysInfo.WorkAreaHeight - frmPagoPos.Height) / 2)
        frmPagoPos.Refresh
    End If
    
    'Cargar bancos
    
    
    Set rsBancos = frsEjecuta_SP("-1|" & CStr(vgintClaveEmpresaContable), "sp_CpSelBanco")
    If rsBancos.RecordCount <> 0 Then
        rsBancos.Filter = "BITESTATUS = 1"
        pLlenarCboRs cboBanco, rsBancos, 4, 5
    End If
    
    vlblnExisteFormaCredito = True
    'Revisa si existe registro por parte de Trabajo Social
    '-----------------------------------------------------
    vlblnRegistroCredito = False
    SQL = "  SELECT mnyCantidadCredito, intnumRegistro "
    SQL = SQL & "  FROM  TSREGISTROCREDITO"
    SQL = SQL & "         inner join CCCLIENTE ON CCCLIENTE.INTNUMCLIENTE = TSREGISTROCREDITO.INTNUMCLIENTE "
    SQL = SQL & "         inner join NODEPARTAMENTO ON CCCLIENTE.SMICVEDEPARTAMENTO = NODEPARTAMENTO.SMICVEDEPARTAMENTO "
    SQL = SQL & "   WHERE CCCLIENTE.CHRTIPOCLIENTE = '" & vlstrTipoCliente & "'"
    SQL = SQL & "          AND CCCLIENTE.INTNUMREFERENCIA =  " & vlLngReferencia
    SQL = SQL & "      AND chrEstatus = 'E'"
    SQL = SQL & "      AND CCCLIENTE.BITACTIVO = 1"
    SQL = SQL & "      AND NODEPARTAMENTO.TNYCLAVEEMPRESA = " & vgintClaveEmpresaContable
    Set rsTemp = frsRegresaRs(SQL, adLockOptimistic, adOpenDynamic)
    If rsTemp.RecordCount > 0 Then
        vldblCantidadCredito = rsTemp.Fields(0)
        vlintNumRegistro = rsTemp.Fields(1)
        vlblnRegistroCredito = True
        SQL = " select intFormaPago,chrDescripcion from PvFormaPago where bitEstatusActivo = 1 and chrTipo = 'C' "
        SQL = SQL & " and smiDepartamento = " & Trim(str(vgintNumeroDepartamento))
        Set rsTemp = frsRegresaRs(SQL, adLockOptimistic, adOpenDynamic)
        If rsTemp.EOF Then
            vlblnExisteFormaCredito = False
        Else
            vlintFormaCredito = rsTemp.Fields(0)
        End If
    End If
    
    vllngCuentaContableCredito = 0
   
    If Not vlblnIncluirFormasCredito Then
        If vlblnPesos Then
            If lblnFormaTrans Then
                vlstrsSQL = "select intFormaPago,chrDescripcion,BITPESOS,BITPREGUNTAFOLIO,CHRTIPO,INTCUENTACONTABLE, FN_CVECATALOGOSATPORNOMBRETIPO('c_FormaPago', INTFORMAPAGO, 'FP', 0) AS CHRFORMAPAGOSAT, BITUTILIZARPINPAD, VCHIMPRESORAVOUCHER, INTCVETERMINAL from PvFormaPago where bitEstatusActivo = 1 and chrTipo <> 'C' "
            Else
                If vgstrForma = "frmEntradaSalidaDinero-S" Then
                    vlstrsSQL = "select intFormaPago,chrDescripcion,BITPESOS,BITPREGUNTAFOLIO,CHRTIPO,INTCUENTACONTABLE, FN_CVECATALOGOSATPORNOMBRETIPO('c_FormaPago', INTFORMAPAGO, 'FP', 0) AS CHRFORMAPAGOSAT, BITUTILIZARPINPAD, VCHIMPRESORAVOUCHER, INTCVETERMINAL from PvFormaPago where bitEstatusActivo = 1 and chrTipo = 'E' "
                Else
                    vlstrsSQL = "select intFormaPago,chrDescripcion,BITPESOS,BITPREGUNTAFOLIO,CHRTIPO,INTCUENTACONTABLE, FN_CVECATALOGOSATPORNOMBRETIPO('c_FormaPago', INTFORMAPAGO, 'FP', 0) AS CHRFORMAPAGOSAT, BITUTILIZARPINPAD, VCHIMPRESORAVOUCHER, INTCVETERMINAL from PvFormaPago where bitEstatusActivo = 1 and chrTipo <> 'C' and chrTipo <> 'B' "
                End If
            End If
        Else
            If lblnFormaTrans Then
                vlstrsSQL = "select intFormaPago,chrDescripcion,BITPESOS,BITPREGUNTAFOLIO,CHRTIPO,INTCUENTACONTABLE, FN_CVECATALOGOSATPORNOMBRETIPO('c_FormaPago', INTFORMAPAGO, 'FP', 0) AS CHRFORMAPAGOSAT, BITUTILIZARPINPAD, VCHIMPRESORAVOUCHER, INTCVETERMINAL from PvFormaPago where bitEstatusActivo = 1 and chrTipo <> 'C' and bitPesos = 0 "
            Else
                If vgstrForma = "frmEntradaSalidaDinero-S" Then
                    vlstrsSQL = "select intFormaPago,chrDescripcion,BITPESOS,BITPREGUNTAFOLIO,CHRTIPO,INTCUENTACONTABLE, FN_CVECATALOGOSATPORNOMBRETIPO('c_FormaPago', INTFORMAPAGO, 'FP', 0) AS CHRFORMAPAGOSAT, BITUTILIZARPINPAD, VCHIMPRESORAVOUCHER, INTCVETERMINAL from PvFormaPago where bitEstatusActivo = 1 and chrTipo = 'E' and bitPesos = 0 "
                Else
                    vlstrsSQL = "select intFormaPago,chrDescripcion,BITPESOS,BITPREGUNTAFOLIO,CHRTIPO,INTCUENTACONTABLE, FN_CVECATALOGOSATPORNOMBRETIPO('c_FormaPago', INTFORMAPAGO, 'FP', 0) AS CHRFORMAPAGOSAT, BITUTILIZARPINPAD, VCHIMPRESORAVOUCHER, INTCVETERMINAL from PvFormaPago where bitEstatusActivo = 1 and chrTipo <> 'C' and chrTipo <> 'B' and bitPesos = 0 "
                End If
            End If
        End If
        'Mensaje si no esta la forma de crédito disponible y se necesita por el registro de trabajo social
        vlblnExisteFormaCredito = False
    Else
        If fblnCreditoVigente() Then
            If vlblnPesos Then
                If lblnFormaTrans Then
                    vlstrsSQL = "select intFormaPago,chrDescripcion,BITPESOS,BITPREGUNTAFOLIO,CHRTIPO,INTCUENTACONTABLE, FN_CVECATALOGOSATPORNOMBRETIPO('c_FormaPago', INTFORMAPAGO, 'FP', 0) AS CHRFORMAPAGOSAT, BITUTILIZARPINPAD, VCHIMPRESORAVOUCHER, INTCVETERMINAL from PvFormaPago where bitEstatusActivo = 1 "
                Else
                    vlstrsSQL = "select intFormaPago,chrDescripcion,BITPESOS,BITPREGUNTAFOLIO,CHRTIPO,INTCUENTACONTABLE, FN_CVECATALOGOSATPORNOMBRETIPO('c_FormaPago', INTFORMAPAGO, 'FP', 0) AS CHRFORMAPAGOSAT, BITUTILIZARPINPAD, VCHIMPRESORAVOUCHER, INTCVETERMINAL from PvFormaPago where bitEstatusActivo = 1 and chrTipo <> 'B' "
                End If
            Else
                If lblnFormaTrans Then
                    vlstrsSQL = "select intFormaPago,chrDescripcion,BITPESOS,BITPREGUNTAFOLIO,CHRTIPO,INTCUENTACONTABLE, FN_CVECATALOGOSATPORNOMBRETIPO('c_FormaPago', INTFORMAPAGO, 'FP', 0) AS CHRFORMAPAGOSAT, BITUTILIZARPINPAD, VCHIMPRESORAVOUCHER, INTCVETERMINAL from PvFormaPago where bitEstatusActivo = 1 and bitPesos = 0 "
                Else
                    vlstrsSQL = "select intFormaPago,chrDescripcion,BITPESOS,BITPREGUNTAFOLIO,CHRTIPO,INTCUENTACONTABLE, FN_CVECATALOGOSATPORNOMBRETIPO('c_FormaPago', INTFORMAPAGO, 'FP', 0) AS CHRFORMAPAGOSAT, BITUTILIZARPINPAD, VCHIMPRESORAVOUCHER, INTCVETERMINAL from PvFormaPago where bitEstatusActivo = 1 and bitPesos = 0 and chrTipo <> 'B' "
                End If
            End If
        Else
            If vlblnPesos Then
                If lblnFormaTrans Then
                    vlstrsSQL = "select intFormaPago,chrDescripcion,BITPESOS,BITPREGUNTAFOLIO,CHRTIPO,INTCUENTACONTABLE, FN_CVECATALOGOSATPORNOMBRETIPO('c_FormaPago', INTFORMAPAGO, 'FP', 0) AS CHRFORMAPAGOSAT, BITUTILIZARPINPAD, VCHIMPRESORAVOUCHER, INTCVETERMINAL from PvFormaPago where bitEstatusActivo = 1 and chrTipo <> 'C' "
                Else
                    vlstrsSQL = "select intFormaPago,chrDescripcion,BITPESOS,BITPREGUNTAFOLIO,CHRTIPO,INTCUENTACONTABLE, FN_CVECATALOGOSATPORNOMBRETIPO('c_FormaPago', INTFORMAPAGO, 'FP', 0) AS CHRFORMAPAGOSAT, BITUTILIZARPINPAD, VCHIMPRESORAVOUCHER, INTCVETERMINAL from PvFormaPago where bitEstatusActivo = 1 and chrTipo <> 'C' and chrTipo <> 'B' "
                End If
            Else
                If lblnFormaTrans Then
                    vlstrsSQL = "select intFormaPago,chrDescripcion,BITPESOS,BITPREGUNTAFOLIO,CHRTIPO,INTCUENTACONTABLE, FN_CVECATALOGOSATPORNOMBRETIPO('c_FormaPago', INTFORMAPAGO, 'FP', 0) AS CHRFORMAPAGOSAT, BITUTILIZARPINPAD, VCHIMPRESORAVOUCHER, INTCVETERMINAL from PvFormaPago where bitEstatusActivo = 1 and chrTipo <> 'C' and bitPesos = 0 "
                Else
                    vlstrsSQL = "select intFormaPago,chrDescripcion,BITPESOS,BITPREGUNTAFOLIO,CHRTIPO,INTCUENTACONTABLE, FN_CVECATALOGOSATPORNOMBRETIPO('c_FormaPago', INTFORMAPAGO, 'FP', 0) AS CHRFORMAPAGOSAT, BITUTILIZARPINPAD, VCHIMPRESORAVOUCHER, INTCVETERMINAL from PvFormaPago where bitEstatusActivo = 1 and chrTipo <> 'C' and chrTipo <> 'B' and bitPesos = 0 "
                End If
            End If
            'Mensaje si no esta la forma de crédito disponible y se necesita por el registro de trabajo social
            vlblnExisteFormaCredito = False
        End If
    End If
    
    If vgstrForma = "frmEntradaSalidaDinero-S" Then
        vlstrsSQL = vlstrsSQL & " and smiDepartamento = " & Trim(str(vgintNumeroDepartamento))
        vlstrsSQL = vlstrsSQL & " UNION select -9, 'DEVOLUCIÓN A PACIENTE POR CUENTAS POR PAGAR' , 1, 0, 'P', -1, null, 0, null, null from dual "
        vlstrsSQL = vlstrsSQL & " Order by chrDescripcion "
    Else
        vlstrsSQL = vlstrsSQL & " and smiDepartamento = " & Trim(str(vgintNumeroDepartamento)) & " Order by chrDescripcion"
    End If
    Set rsFormasPago = frsRegresaRs(vlstrsSQL)
    
    With grdFormasPago
        .Clear
        .Cols = cintColumnas
        .Rows = 1
        .ColWidth(cintColIdForma) = 0
        
        If vlblnLicenciaContaElectronica Or vgblnCapturaraDatosBanco Then
            .ColWidth(cintColNombre) = 8570
        Else
            .ColWidth(cintColNombre) = 4075
        End If
        
        .ColWidth(cintColMoneda) = 0
        .ColWidth(cintColReferencia) = 0
        .ColWidth(cintColTipo) = 0
        .ColWidth(cintColCtaContable) = 0
        .ColWidth(cintColClaveFormaPagoSAT) = 0
        .ColWidth(cintColUsarPinpad1) = 0
        .ColWidth(cintColUriPinpad1) = 0
        .ColWidth(cintColImprVoucher1) = 0
        .ColWidth(cintColCveMoneda1) = 0
        .ColWidth(cintColPpProv1) = 0
        .ColWidth(cintColPpUsr1) = 0
        .ColWidth(cintColPpPwd1) = 0
        .ColWidth(cintColPpHost1) = 0
        .ColWidth(cintColPpPort1) = 0
        .ColWidth(cintColPpCve1) = 0
        .ColWidth(cintColPpTId1) = 0
    End With
    
    If rsFormasPago.RecordCount <> 0 Then
        pCargaFormas
    End If
    vgblnPermiteSalir = False
    
    pGrid
    
    Dim inicioSocket As Boolean
      Set rsTemp = frsSelParametros("PV", vgintClaveEmpresaContable, "INTTERMINALSANTANDER")
    If Not rsTemp.EOF Then
    inicioSockect = Cconecta("127.0.0.1", rsTemp!Valor)
    End If
    
    
End Sub

Private Sub pCargaFormas()
    Dim rsPinPad As ADODB.Recordset
    With grdFormasPago
        rsFormasPago.MoveFirst
        Do While Not rsFormasPago.EOF
            If rsFormasPago!BITUTILIZARPINPAD <> 0 And Not IsNull(rsFormasPago!intCveTerminal) Then
                Set rsPinPad = frsRegresaRs("select * from PVTerminal where intCveTerminal = " & rsFormasPago!intCveTerminal, adLockReadOnly, adOpenForwardOnly)
                If Not rsPinPad.EOF Then
                    .TextMatrix(.Rows - 1, cintColUriPinpad1) = IIf(IsNull(rsPinPad!VCHURI), "", rsPinPad!VCHURI)
                    .TextMatrix(.Rows - 1, cintColPpProv1) = IIf(IsNull(rsPinPad!INTPROVIDER), "", rsPinPad!INTPROVIDER)
                    If rsFormasPago!BITPESOS = 0 Then
                         .TextMatrix(.Rows - 1, cintColCveMoneda1) = IIf(IsNull(rsPinPad!vchUSD), "", rsPinPad!vchUSD)
                    Else
                         .TextMatrix(.Rows - 1, cintColCveMoneda1) = IIf(IsNull(rsPinPad!vchMXN), "", rsPinPad!vchMXN)
                    End If
                    .TextMatrix(.Rows - 1, cintColPpUsr1) = IIf(IsNull(rsPinPad!VCHUSR), "", rsPinPad!VCHUSR)
                    .TextMatrix(.Rows - 1, cintColPpPwd1) = IIf(IsNull(rsPinPad!VCHPWD), "", rsPinPad!VCHPWD)
                    .TextMatrix(.Rows - 1, cintColPpHost1) = IIf(IsNull(rsPinPad!VCHIP), "", rsPinPad!VCHIP)
                    .TextMatrix(.Rows - 1, cintColPpPort1) = IIf(IsNull(rsPinPad!VCHPORT), "", rsPinPad!VCHPORT)
                    .TextMatrix(.Rows - 1, cintColPpCve1) = rsPinPad!intCveTerminal
                    .TextMatrix(.Rows - 1, cintColPpTId1) = IIf(rsFormasPago!BITPESOS = 0, IIf(IsNull(rsPinPad!VCHTUSD), "", rsPinPad!VCHTUSD), IIf(IsNull(rsPinPad!VCHTMXN), "", rsPinPad!VCHTMXN))
                End If
                rsPinPad.Close
            End If
            .TextMatrix(.Rows - 1, cintColIdForma) = rsFormasPago!intFormaPago
            .TextMatrix(.Rows - 1, cintColNombre) = rsFormasPago!chrDescripcion
            .TextMatrix(.Rows - 1, cintColMoneda) = rsFormasPago!BITPESOS
            .TextMatrix(.Rows - 1, cintColTipo) = rsFormasPago!chrTipo
            
            If Trim(rsFormasPago!chrTipo) = "E" Then 'Si es efectivo, no forza la referencia
                .TextMatrix(.Rows - 1, cintColReferencia) = rsFormasPago!bitpreguntafolio
            ElseIf Trim(rsFormasPago!chrTipo) = "C" Then 'Si es crédito, no forza la referencia
                .TextMatrix(.Rows - 1, cintColReferencia) = rsFormasPago!bitpreguntafolio
            ElseIf Trim(rsFormasPago!chrTipo) = "P" Then 'Si es devoluciones por cuentas por pagar, no forza la referencia
                .TextMatrix(.Rows - 1, cintColReferencia) = rsFormasPago!bitpreguntafolio
            Else
                .TextMatrix(.Rows - 1, cintColReferencia) = "1" 'Los demás si los forza, cambio para CFD y CFDi
            End If
                
            .TextMatrix(.Rows - 1, cintColCtaContable) = rsFormasPago!INTCUENTACONTABLE
            .TextMatrix(.Rows - 1, cintColClaveFormaPagoSAT) = IIf(IsNull(rsFormasPago!CHRFORMAPAGOSAT), "", rsFormasPago!CHRFORMAPAGOSAT)
            .TextMatrix(.Rows - 1, cintColUsarPinpad1) = rsFormasPago!BITUTILIZARPINPAD
            .TextMatrix(.Rows - 1, cintColImprVoucher1) = IIf(IsNull(rsFormasPago!VCHIMPRESORAVOUCHER), "", rsFormasPago!VCHIMPRESORAVOUCHER)
            .Rows = .Rows + 1
            rsFormasPago.MoveNext
        Loop
        .Rows = .Rows - 1
        .ColAlignment(cintColNombre) = flexAlignLeftCenter
        .Col = cintColNombre
        .Row = 0
    End With
End Sub

Private Function fblnCreditoVigente() As Boolean
    Dim rsCreditoVigenteAsignado As New ADODB.Recordset
    Dim rsCuentaContableCredito As New ADODB.Recordset

    vllngNumCliente = 0
    vllngDeptoCliente = 0
    vldblLimiteCredito = 0
    fblnCreditoVigente = False
    vlstrsSQL = "select count(*) from CcCliente Inner join nodepartamento on cccliente.smicvedepartamento = nodepartamento.smicvedepartamento " & _
                " Where cccliente.intNumReferencia = " + str(vlLngReferencia) + " And cccliente.chrTipoCliente = " + " '" + vlstrTipoCliente + "'" + " and cccliente.bitActivo=1 and nodepartamento.tnyclaveempresa = " & vgintClaveEmpresaContable
    Set rsCreditoVigenteAsignado = frsRegresaRs(vlstrsSQL)
    If rsCreditoVigenteAsignado.Fields(0) <> 0 Then
        fblnCreditoVigente = True
        vlstrsSQL = "select CcCliente.intnumcliente, CCCLIENTE.SMICVEDEPARTAMENTO, CcCliente.intNumCuentaContable, CcCliente.mnyLimiteCredito from CcCliente " & _
                    " Inner join nodepartamento on cccliente.smicvedepartamento = nodepartamento.smicvedepartamento " & _
                    " Where CcCliente.intNumReferencia = " + str(vlLngReferencia) + " And CcCliente.chrTipoCliente = " + " '" + vlstrTipoCliente + "'" + " and CcCliente.bitActivo=1" & _
                    " and nodepartamento.tnyclaveempresa = " & vgintClaveEmpresaContable
        Set rsCuentaContableCredito = frsRegresaRs(vlstrsSQL)
        If rsCuentaContableCredito.RecordCount <> 0 Then
            vllngNumCliente = rsCuentaContableCredito!intNumCliente
            vllngDeptoCliente = rsCuentaContableCredito!smicvedepartamento
            vllngCuentaContableCredito = rsCuentaContableCredito!INTNUMCUENTACONTABLE
            vldblLimiteCredito = rsCuentaContableCredito!mnyLimiteCredito
        End If
    End If
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = 1
        frmPagoPos.Hide
    End If
End Sub

Private Sub grdFormas_DblClick()
    Dim vldblCantidad As Double
    If grdFormas.RowData(1) <> 0 Then
        vldblCantidad = Val(grdFormas.TextMatrix(grdFormas.Row, cintColCantidadReal))
        pActualizaTotales vldblCantidad, False
        pBorrarRegMshFGrdData grdFormas.Row, grdFormas, True
        If grdFormas.Rows = 2 And grdFormas.RowData(1) = 0 Then
            pGrid
        End If
        grdFormasPago.SetFocus
        
        txtCambio.Text = " 0.00"
    End If
End Sub

Private Sub mskFecha_GotFocus()
On Error GoTo NotificaError

    pSelMkTexto MskFecha

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":MskFecha_GotFocus"))
End Sub

Private Sub MskFecha_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError

    If KeyAscii = 13 Then
        If fblnFechaValidaChequeTrans Then
            If cboBanco.Enabled Or cboTipoCargoBancario.Enabled Then
                SendKeys vbTab
            Else
                If Val(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColReferencia)) = 1 Then
                    txtFolio.SetFocus
                Else
                    txtCantidad.SetFocus
                End If
            End If
        Else
            pSelMkTexto MskFecha
        End If
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
        
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":MskFecha_KeyPress"))
End Sub

Private Sub mskFecha_LostFocus()
On Error GoTo NotificaError
    
    If Not fblnFechaValidaChequeTrans Then MskFecha.SetFocus
        
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFecha_LostFocus"))
End Sub

Private Sub tmrMsgPP_Timer()
    vl_blnTime = True
End Sub

Private Sub txtBancoExtranjero_GotFocus()
On Error GoTo NotificaError

    pSelTextBox txtBancoExtranjero

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtBancoExtranjero_GotFocus"))
End Sub

Public Sub txtBancoExtranjero_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError

    KeyAscii = Asc(StrConv(Chr$(KeyAscii), vbUpperCase))

    If KeyAscii = 13 Then
        If Trim(txtBancoExtranjero.Text) = "" Then
            cboBancoSAT.Visible = True
            If fblnCanFocus(cboBancoSAT) Then cboBancoSAT.SetFocus
            txtBancoExtranjero.Visible = False
        Else
            If txtCuentaBancaria.Enabled Or cboCuentasPrevias.Enabled Or MskFecha.Enabled Or cboBanco.Enabled Or cboTipoCargoBancario.Enabled Then
                SendKeys vbTab
            Else
                If Val(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColReferencia)) = 1 Then
                    txtFolio.SetFocus
                Else
                    txtCantidad.SetFocus
                End If
            End If
        End If
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtBancoExtranjero_KeyPress"))
End Sub

Private Sub txtBancoExtranjero_KeyUp(KeyCode As Integer, Shift As Integer)
    pHabilitaFechaChequeTrans
End Sub

Private Sub txtBancoExtranjero_LostFocus()
    If txtBancoExtranjero.Text = "" Then
        cboBancoSAT.Visible = True
        txtBancoExtranjero.Visible = False
    End If
End Sub

Private Sub txtBancoExtranjero_Validate(Cancel As Boolean)
    pHabilitaFechaChequeTrans
End Sub

Public Sub txtCantidad_GotFocus()
    vlstrValorAnterior = Trim(txtCantidad.Text)
    vlblnCambioCantidad = False
    pSelTextBox txtCantidad
End Sub

Public Sub txtCantidad_KeyPress(KeyAscii As Integer)
    Dim vldblTMPTipoCambio As Double
    
    If Val(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColMoneda)) = 0 And vlblnPesos Then
        vldblTMPTipoCambio = vldblTipoCambioDia
    Else
        vldblTMPTipoCambio = 1
    End If
    
    If Not fblnFormatoCantidad(txtCantidad, KeyAscii, 2) Then
       KeyAscii = 7
    Else
        If KeyAscii = 13 Then
        
            If Val(Format(txtDiferencia.Text, "###############.##")) = 0 Then Exit Sub
            
            vldblValorSinFormato = Val(Format(txtCantidad.Text, "############.00"))
            
            If grdFormasPago.TextMatrix(grdFormasPago.Row, cintColCtaContable) = 0 And Trim(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColTipo)) <> "B" And Trim(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColTipo)) <> "C" Then
                MsgBox SIHOMsg(1617), vbExclamation, "Mensaje"
                Exit Sub
            End If
            
            If Not fblnValidaFormaPagoSAT(grdFormasPago.TextMatrix(grdFormasPago.Row, 0)) Then
                MsgBox "La forma de pago no tiene definido el método de pago del SAT para CFDI ver " & vgstrVersionCFDI, vbExclamation, "Mensaje"
                Exit Sub
            End If
            
            If fblnInformacionValidaExtra Then
                If fblnDatosValidos Then
                    If (vlblnCambioCantidad And vldblValorSinFormato > Val(Format(txtDiferencia.Text, "############.00")) / vldblTMPTipoCambio) Or (vldblValorSinFormato = 0) Or Val(Format(txtCantidad.Text, "############.00")) < 0 Then
                        txtCantidad.Text = Format(Val(Format(txtDiferencia.Text, "############.##")) / vldblTMPTipoCambio, "############.00")
                        vldblValorConDecimales = Val(Format(txtDiferencia.Text, "############.##")) / vldblTMPTipoCambio
                        pEnfocaTextBox txtCantidad
                        txtCambio.Text = Format(vldblValorSinFormato * vldblTMPTipoCambio - Val(Format(txtDiferencia.Text, "############.00")), "$###,###,###,###.00")
                        vldblValorSinFormato = Val(Format(txtCantidad.Text, "############.00"))
                    Else
                        txtCambio.Text = " 0.00"
                    End If
                    
                    'Se verifica el tipo de pago
                    '----------------------------------------
                    
                    'Si es crédito o efectivo, revisará que no esté seleccionada esa forma de pago
                    If (Trim(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColTipo)) = "E") Or (Trim(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColTipo)) = "C") Then
                    
                        If Not fblnExisteForma(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColIdForma)) Then
                            'Validación para que no se puedan combinar la forma de pago crédito con otras, para Honoraios
                            If grdFormasPago.TextMatrix(grdFormasPago.Row, cintColTipo) = "C" And vgstrForma = "Honorarios" Then
                              If fblnExisteFormas Then
                                    'Seleccione otra forma de pago
                                    MsgBox SIHOMsg(934), vbOKOnly + vbExclamation, "Mensaje"
                              Else
                                txtCantidad.Text = Format(str(vlstrValorAnterior), "###,###,###,###.00")
                                pAgregaForma
                                If vlblnExisteError Then Exit Sub
                              End If
                            Else
                                pAgregaForma
                                If vlblnExisteError Then Exit Sub
                            End If
                        Else
                            'Esta forma de pago ya está seleccionada.
                            MsgBox SIHOMsg(294), vbOKOnly + vbInformation, "Mensaje"
                        End If
                    Else
                        'Verifica que la referencia no esté repetida
                        If Not fblnExisteReferencia(txtFolio.Text) Then
                            If grdFormasPago.TextMatrix(grdFormasPago.Row, cintColTipo) = "C" And vgstrForma = "Honorarios" Then
                              If fblnExisteFormas Then
                                'Seleccione otra forma de pago
                                MsgBox SIHOMsg(934), vbOKOnly + vbExclamation, "Mensaje"
                                Exit Sub
                              Else
                                txtCantidad.Text = Format(str(vlstrValorAnterior), "###,###,###,###.00")
                                pAgregaForma
                                If vlblnExisteError Then Exit Sub
                              End If
                            Else
                                pAgregaForma
                                If vlblnExisteError Then Exit Sub
                            End If
                        Else
                            '¡Ya existe esta forma de pago con esta referencia!
                            MsgBox SIHOMsg(1116), vbOKOnly + vbExclamation, "Mensaje"
                        End If
                        
                    End If
                    
                    If vlblnLicenciaContaElectronica Or vgblnCapturaraDatosBanco Then
                        vlstrRFCOriginal = txtRFC.Text
                        vldtmFecha = CDate(MskFecha.Text)
                    End If
                    
                    If txtImporte.Text <> txtCantidadPagada.Text Then
                        grdFormasPago.SetFocus
                    Else
                        pLimpia
                        phabilitaInfoExtra
                        If fblnCanFocus(cmdAceptar) Then cmdAceptar.SetFocus
                    End If
                Else
                    txtCantidad.Text = vlstrValorAnterior
                End If
            Else
                txtCantidad.Text = vlstrValorAnterior
            End If
            vlblnCambioCantidad = False
        Else
            vlblnCambioCantidad = True
        End If
    End If
End Sub
Private Function Cconecta(Ip As String, Puerto As Long)
If Ip <> "" And Conectado = False Then

 Winsock1.Connect Ip, Puerto  'Conectamos el winsock
Cconecta = True
Timer1.Interval = 100

End If

End Function
Private Function fblnDatosValidos()
    fblnDatosValidos = True
    
    If Val(Format(txtCantidad.Text, "###########.##")) = 0 Then
        fblnDatosValidos = False
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbOKOnly + vbExclamation, "Mensaje"
        txtCantidad.SetFocus
    End If
    If fblnDatosValidos And Val(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColReferencia)) = 1 And Trim(txtFolio.Text) = "" And Val(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColUsarPinpad1)) <> 1 Then
        fblnDatosValidos = False
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbOKOnly + vbExclamation, "Mensaje"
        txtFolio.SetFocus
    End If
    'Validación para que se incluyan al menos 4 caracteres de referencia
    If fblnDatosValidos And Val(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColReferencia)) = 1 And Len(Trim(txtFolio.Text)) < 4 And Val(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColUsarPinpad1)) <> 1 Then
        fblnDatosValidos = False
        '¡No ha ingresado datos!
        MsgBox "¡Se deben de indicar al menos 4 dígitos de la referencia de pago!", vbOKOnly + vbExclamation, "Mensaje"
        txtFolio.SetFocus
    End If
    If fblnDatosValidos And Trim(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColTipo)) = "B" And cboBanco.ListIndex = -1 Then
        fblnDatosValidos = False
        '¡Dato no válido, seleccione un valor de la lista!
        MsgBox SIHOMsg(3), vbOKOnly + vbExclamation, "Mensaje"
        cboBanco.SetFocus
    End If
    If fblnDatosValidos And Val(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColIdForma)) = -9 Then
        If Not fblnValidaCtaDevolucionesxCP(vgintClaveEmpresaContable) Then
            fblnDatosValidos = False
            txtCantidad.SetFocus
        End If
    End If
End Function

Private Function fblnExisteFormas() As Boolean
     Dim X As Integer
     
     fblnExisteFormas = False
     For X = 1 To grdFormas.Rows - 1
        If grdFormas.RowData(X) > 0 Then
            fblnExisteFormas = True
        End If
     Next X
End Function

Private Function fblnExisteForma(vllngxNumeroForma As Long) As Boolean
     Dim X As Integer
     
     fblnExisteForma = False
     For X = 1 To grdFormas.Rows - 1
        If grdFormas.RowData(X) = vllngxNumeroForma Then
            fblnExisteForma = True
        End If
     Next X
End Function

Private Function fblnExisteReferencia(vllngxNumeroReferencia As String) As Boolean
     Dim X As Integer
     Dim Ref As Boolean
     Dim Desc As Boolean
     
     Dim IntCveFormaGrabada As Integer
     Dim IntCveFormaNueva As Integer
     
     IntCveFormaGrabada = 0
     IntCveFormaNueva = 0
     fblnExisteReferencia = False
     
     For X = 1 To grdFormas.Rows - 1
        If Trim(grdFormas.TextMatrix(X, cIntColFolio)) = Trim(vllngxNumeroReferencia) Then
            IntCveFormaGrabada = grdFormas.RowData(X)
            If fblnExisteDescripcion(grdFormasPago.TextMatrix(grdFormasPago.Row, 1)) Then
                IntCveFormaNueva = CInt(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColIdForma))
                If IntCveFormaGrabada = IntCveFormaNueva Then
                    fblnExisteReferencia = True
                    Exit Function
                End If
            End If
        End If
     Next X
End Function

Private Function fblnExisteDescripcion(vlstrDescipcion As String) As Boolean
     Dim X As Integer
     
     fblnExisteRef = False
     For X = 1 To grdFormas.Rows - 1
        If grdFormas.TextMatrix(X, cintColDescripcion) = (vlstrDescipcion) Then
            fblnExisteDescripcion = True
            Exit Function
        End If
     Next X
End Function

Private Sub pAgregaForma()
    Dim X As Integer
    Dim vldblCantidad As Double
    Dim rsCuentaContableFormaPago As New ADODB.Recordset
    Dim vllngCuentaContableFormaPago As Long
    Dim vldblTipoCambioParaForma As Double
    Dim vldblCantidadReal As Double
    Dim vldblDolares As Double
    Dim vlbolCredito As Boolean
    Dim vldblCantidadComisionBancaria As Double
    Dim vldblIvaComisionBancaria As Double
    Dim vldblSaldoActual As Double
    Dim vgstrParametrosSP As String
    Dim blnTipoCredito As Boolean
    
    If grdFormasPago.TextMatrix(grdFormasPago.Row, cintColTipo) = "C" Then
        blnTipoCredito = True
    Else
        blnTipoCredito = False
    End If
    
    If vlblnPesos And Val(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColMoneda)) = 0 Then
        vldblTipoCambio = vldblTipoCambioDia
    Else
        vldblTipoCambio = 1
    End If
    If CDbl(vlstrValorAnterior) = CDbl(Trim(txtCantidad.Text)) Then
        vldblCantidad = vldblValorConDecimales * vldblTipoCambio
        vldblCantidadReal = vldblValorConDecimales * vldblTipoCambio
        vldblDolares = vldblValorConDecimales
    Else
        vldblCantidad = Val(Format(Trim(txtCantidad.Text), "############.##")) * vldblTipoCambio
        vldblCantidadReal = vldblCantidad
        vldblDolares = Val(Format(Trim(txtCantidad.Text), "############.##"))
    End If
    
    vllngCuentaContableFormaPago = 0
    If Trim(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColTipo)) = "C" Then
        vllngCuentaContableFormaPago = vllngCuentaContableCredito
        vlbolCredito = True
    Else
        If Trim(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColTipo)) = "B" Then
            pPosicionaBanco
            vllngCuentaContableFormaPago = rsBancos!intNumeroCuenta
        ElseIf Trim(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColTipo)) = "P" Then
            vllngCuentaContableFormaPago = lngCtaDevolucionesCuentasPagar
        Else
            vllngCuentaContableFormaPago = Val(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColCtaContable))
        End If
        vlbolCredito = False
    End If
    If Val(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColMoneda)) = 1 Then
        vldblTipoCambioParaForma = 0
    Else
        vldblTipoCambioParaForma = vldblTipoCambioDia
    End If
    
    vlblnExisteError = False
    
    ' Determina el saldo actual
    If vlbolCredito Then
        vgstrParametrosSP = vllngNumCliente & "|" & vllngDeptoCliente & "|" & fstrFechaSQL(Trim(fdtmServerFecha))
        Set rsReporte = frsEjecuta_SP(vgstrParametrosSP, "Sp_CCSelSaldoALaFecha")
        If rsReporte.RecordCount <> 0 Then
            vldblSaldoActual = Round(IIf(IsNull(rsReporte!Saldo), 0, rsReporte!Saldo), 2)
        End If
    End If
    
    
    If (Val(Format(Trim(txtCantidad.Text), "############.##")) + vldblSaldoActual) > vldblLimiteCredito And vlbolCredito And vldblLimiteCredito > 0 Then
        'La cantidad capturada más el saldo actual del cliente excede el límite de crédito otorgado.
        MsgBox SIHOMsg(734), vbOKOnly + vbExclamation, "Mensaje"
        vlblnExisteError = True
        pEnfocaTextBox txtCantidad
        Exit Sub
    End If
           
    If grdFormas.RowData(1) = -1 Then
        X = 1
    Else
        grdFormas.Rows = grdFormas.Rows + 1
        X = grdFormas.Rows - 1
    End If
    
    grdFormas.TextMatrix(X, cintColDescripcion) = grdFormasPago.TextMatrix(grdFormasPago.Row, cintColNombre)
    grdFormas.TextMatrix(X, cIntColFolio) = txtFolio.Text
    grdFormas.TextMatrix(X, cintColCantidad) = txtCantidad.Text
    grdFormas.TextMatrix(X, cintColCtaContableFormaPago) = vllngCuentaContableFormaPago
    grdFormas.TextMatrix(X, cintColTipoCambio) = vldblTipoCambioParaForma
    grdFormas.TextMatrix(X, cintColCantidadReal) = vldblCantidadReal
    grdFormas.TextMatrix(X, cintColCredito) = vlbolCredito
    grdFormas.TextMatrix(X, cintColDolares) = vldblDolares
    If Trim(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColTipo)) = "B" Then
        grdFormas.TextMatrix(X, cintColIdBanco) = cboBanco.ItemData(cboBanco.ListIndex)
        grdFormas.TextMatrix(X, cintColMonedaBanco) = rsBancos!bitestatusmoneda
    Else
        grdFormas.TextMatrix(X, cintColIdBanco) = 0
        grdFormas.TextMatrix(X, cintColMonedaBanco) = 0
    End If
    grdFormas.TextMatrix(X, cintColUsarPinpad2) = grdFormasPago.TextMatrix(grdFormasPago.Row, cintColUsarPinpad1)
    grdFormas.TextMatrix(X, cintColUriPinpad2) = grdFormasPago.TextMatrix(grdFormasPago.Row, cintColUriPinpad1)
    grdFormas.TextMatrix(X, cintColImprVoucher2) = grdFormasPago.TextMatrix(grdFormasPago.Row, cintColImprVoucher1)
    grdFormas.TextMatrix(X, cintColCveMoneda2) = grdFormasPago.TextMatrix(grdFormasPago.Row, cintColCveMoneda1)
    grdFormas.TextMatrix(X, cintColPpProv2) = grdFormasPago.TextMatrix(grdFormasPago.Row, cintColPpProv1)
    grdFormas.TextMatrix(X, cintColPpUsr2) = grdFormasPago.TextMatrix(grdFormasPago.Row, cintColPpUsr1)
    grdFormas.TextMatrix(X, cintColPpPwd2) = grdFormasPago.TextMatrix(grdFormasPago.Row, cintColPpPwd1)
    grdFormas.TextMatrix(X, cintColPpHost2) = grdFormasPago.TextMatrix(grdFormasPago.Row, cintColPpHost1)
    grdFormas.TextMatrix(X, cintColPpPort2) = grdFormasPago.TextMatrix(grdFormasPago.Row, cintColPpPort1)
    grdFormas.TextMatrix(X, cintColPpCve2) = grdFormasPago.TextMatrix(grdFormasPago.Row, cintColPpCve1)
    grdFormas.TextMatrix(X, cintColPpTId2) = grdFormasPago.TextMatrix(grdFormasPago.Row, cintColPpTId1)
    
    'Datos de la comisión bancaria seleccionada
    If cboTipoCargoBancario.ListIndex <> -1 Then
        If rsTipoCargoBancario.RecordCount <> 0 Then
            rsTipoCargoBancario.MoveFirst
            Do While Not rsTipoCargoBancario.EOF
                If cboTipoCargoBancario.ItemData(cboTipoCargoBancario.ListIndex) = rsTipoCargoBancario!intcvetipocargo Then
                    If vlstrValorAnterior = Trim(txtCantidad.Text) Then
                        vldblCantidadComisionBancaria = vldblValorConDecimales * vldblTipoCambio * rsTipoCargoBancario!mnycomision / 100
                        vldblIvaComisionBancaria = (vldblValorConDecimales * vldblTipoCambio * rsTipoCargoBancario!mnycomision / 100) * rsTipoCargoBancario!smyIVA / 100
                    Else
                        vldblCantidadComisionBancaria = Val(Format(txtCantidad.Text)) * vldblTipoCambio * rsTipoCargoBancario!mnycomision / 100
                        vldblIvaComisionBancaria = (Val(Format(txtCantidad.Text)) * vldblTipoCambio * rsTipoCargoBancario!mnycomision / 100) * rsTipoCargoBancario!smyIVA / 100
                    End If
                    grdFormas.TextMatrix(X, cintColCtaComisionBanc) = rsTipoCargoBancario!intNumeroCuenta    'Cuenta contable de la comisión bancaria
                    grdFormas.TextMatrix(X, cintColComisionBanc) = vldblCantidadComisionBancaria          'Cantidad de la comisión bancaria
                    grdFormas.TextMatrix(X, cintColIVAComisionBanc) = vldblIvaComisionBancaria               'Iva de la comisión bancaria
                    Exit Do
                End If
                rsTipoCargoBancario.MoveNext
            Loop
        End If
    End If
    
    If vlblnLicenciaContaElectronica Or vgblnCapturaraDatosBanco Then
        grdFormas.TextMatrix(X, cIntColRFC) = Trim(txtRFC.Text)
        
        If cboBancoSAT.ListIndex < 1 Then
            vlClaveBancoSAT = "000"
        Else
            If Len(Trim(str(cboBancoSAT.ItemData(cboBancoSAT.ListIndex)))) = 3 Then
                vlClaveBancoSAT = Trim(str(cboBancoSAT.ItemData(cboBancoSAT.ListIndex)))
            Else
                If Len(Trim(str(cboBancoSAT.ItemData(cboBancoSAT.ListIndex)))) = 2 Then
                    vlClaveBancoSAT = "0" & Trim(str(cboBancoSAT.ItemData(cboBancoSAT.ListIndex)))
                Else
                    vlClaveBancoSAT = "00" & Trim(str(cboBancoSAT.ItemData(cboBancoSAT.ListIndex)))
                End If
            End If
        End If
        
        grdFormas.TextMatrix(X, cintColBancoSAT) = vlClaveBancoSAT
        grdFormas.TextMatrix(X, cintColBancoExt) = IIf(vlClaveBancoSAT = "000", Trim(txtBancoExtranjero.Text), "")
        
        If Trim(txtCuentaBancaria.Text) <> "" And txtCuentaBancaria.Enabled Then
            grdFormas.TextMatrix(X, cIntColCuentaBancaria) = Trim(txtCuentaBancaria.Text)
        Else
            If Trim(cboCuentasPrevias.Text) <> "" And cboCuentasPrevias.Enabled Then
                grdFormas.TextMatrix(X, cIntColCuentaBancaria) = Trim(cboCuentasPrevias.Text)
            Else
                grdFormas.TextMatrix(X, cIntColCuentaBancaria) = ""
            End If
        End If
        
        If intModoMasivo = 1 Then
            grdFormas.TextMatrix(X, cIntColCuentaBancaria) = Trim(txtCuentaBancaria.Text)
        End If
                
        grdFormas.TextMatrix(X, cintColFechaCqTrans) = Trim(MskFecha.Text)
    End If
    
    grdFormas.RowData(X) = grdFormasPago.TextMatrix(grdFormasPago.Row, cintColIdForma)
    pActualizaTotales vldblCantidadReal, True
    
    'Oculta las formas de pago de crédito o las que no son de crédito, segun sea el caso, para CFDI 3.3
    If vgstrVersionCFDI <> "3.2" Then
        For X = 0 To grdFormasPago.Rows - 1
            If grdFormasPago.TextMatrix(X, cintColTipo) = "C" Then
                grdFormasPago.RowHeight(X) = IIf(blnTipoCredito, grdFormasPago.RowHeight(X), 0)
            Else
                grdFormasPago.RowHeight(X) = IIf(blnTipoCredito, 0, grdFormasPago.RowHeight(X))
            End If
        Next
    End If
End Sub

Private Sub pPosicionaBanco()
    Dim blnTermina As Boolean

    rsBancos.MoveFirst
    Do While Not rsBancos.EOF And Not blnTermina
        If cboBanco.ListIndex = -1 Then
            blnTermina = rsBancos!tnynumerobanco = 0
        Else
            blnTermina = rsBancos!tnynumerobanco = cboBanco.ItemData(cboBanco.ListIndex)
        End If
        If Not blnTermina Then
            rsBancos.MoveNext
        End If
    Loop
End Sub

Private Sub pActualizaTotales(vldblxCantidad As Double, vlblnSuma As Boolean)
    If vlblnSuma Then
        txtCantidadPagada.Text = str(Val(Format(txtCantidadPagada.Text, "############.00")) + vldblxCantidad)
    Else
        txtCantidadPagada.Text = str(Val(Format(txtCantidadPagada.Text, "############.00")) - vldblxCantidad)
    End If
    txtDiferencia.Text = str(Val(Format(txtImporte.Text, "############.00")) - Val(Format(txtCantidadPagada.Text, "############.00")))

    If Val(Format(txtCantidadPagada.Text, "############.00")) = 0 Then
        txtCantidadPagada.Text = ""
    Else
        txtCantidadPagada.Text = Format(txtCantidadPagada.Text, "###,###,###,###.00")
    End If
    If Val(txtDiferencia.Text) = 0 Then
        txtDiferencia.Text = ""
    Else
        txtDiferencia.Text = Format(txtDiferencia.Text, "###,###,###,###.00")
    End If
    If txtImporte.Text = txtCantidadPagada.Text Or vgblnPermiteSalir Then
        cmdAceptar.Enabled = True
    Else
        cmdAceptar.Enabled = False
    End If
End Sub

Private Sub txtCuentaBancaria_GotFocus()
On Error GoTo NotificaError

    pSelTextBox txtCuentaBancaria

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtCuentaBancaria_GotFocus"))
End Sub

Private Sub txtCuentaBancaria_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then KeyAscii = 7
    If KeyAscii = 13 Then
        pHabilitaFechaChequeTrans
        If MskFecha.Enabled Or cboBanco.Enabled Or cboTipoCargoBancario.Enabled Then
            SendKeys vbTab
        Else
            If Val(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColReferencia)) = 1 Then
                txtFolio.SetFocus
            Else
                txtCantidad.SetFocus
            End If
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtCuentaBancaria_KeyPress"))
End Sub

Private Sub txtCuentaBancaria_KeyUp(KeyCode As Integer, Shift As Integer)
    pHabilitaFechaChequeTrans
End Sub

Private Sub txtCuentaBancaria_Validate(Cancel As Boolean)
    pHabilitaFechaChequeTrans
End Sub

Private Sub TxtFolio_GotFocus()
    pSelTextBox txtFolio
End Sub

Private Sub TxtFolio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then KeyAscii = 7
    End If
End Sub

Private Sub phabilitaInfoExtra()
    If vlblnLicenciaContaElectronica Or vgblnCapturaraDatosBanco Then
        With grdFormasPago
            cboBancoSAT.Enabled = Trim(.TextMatrix(.Row, cintColTipo)) = "H" Or Trim(.TextMatrix(.Row, cintColTipo)) = "B" Or Trim(.TextMatrix(.Row, cintColTipo)) = "T"
            txtBancoExtranjero.Enabled = Trim(.TextMatrix(.Row, cintColTipo)) = "H" Or Trim(.TextMatrix(.Row, cintColTipo)) = "B" Or Trim(.TextMatrix(.Row, cintColTipo)) = "T"
            lblBancoSAT.Enabled = Trim(.TextMatrix(.Row, cintColTipo)) = "H" Or Trim(.TextMatrix(.Row, cintColTipo)) = "B" Or Trim(.TextMatrix(.Row, cintColTipo)) = "T"
'            txtCuentaBancaria.Enabled = Trim(.TextMatrix(.Row, cintColTipo)) = "H" Or Trim(.TextMatrix(.Row, cintColTipo)) = "B"
'            lblCuentaBancaria.Enabled = Trim(.TextMatrix(.Row, cintColTipo)) = "H" Or Trim(.TextMatrix(.Row, cintColTipo)) = "B"
'            cboCuentasPrevias.Enabled = Trim(.TextMatrix(.Row, cintColTipo)) = "H" Or Trim(.TextMatrix(.Row, cintColTipo)) = "B"
            txtRFC.Enabled = Trim(.TextMatrix(.Row, cintColTipo)) = "E" Or Trim(.TextMatrix(.Row, cintColTipo)) = "T" Or Trim(.TextMatrix(.Row, cintColTipo)) = "H" Or Trim(.TextMatrix(.Row, cintColTipo)) = "B" Or Trim(.TextMatrix(.Row, cintColTipo)) = "P"
            lblRFC.Enabled = Trim(.TextMatrix(.Row, cintColTipo)) = "E" Or Trim(.TextMatrix(.Row, cintColTipo)) = "T" Or Trim(.TextMatrix(.Row, cintColTipo)) = "H" Or Trim(.TextMatrix(.Row, cintColTipo)) = "B" Or Trim(.TextMatrix(.Row, cintColTipo)) = "P"
            
            cboBancoSAT.ListIndex = -1
            txtBancoExtranjero.Text = ""
            cboBancoSAT.Visible = True
            txtBancoExtranjero.Visible = False
            txtCuentaBancaria.Text = ""
            txtCuentaBancaria.Visible = True
            cboCuentasPrevias.Visible = False
            txtRFC.Text = vlstrRFCOriginal
            
            pHabilitaFechaChequeTrans
        End With
    End If
End Sub

Private Sub pHabilitaFechaChequeTrans()
    With grdFormasPago
        If vlblnLicenciaContaElectronica Or vgblnCapturaraDatosBanco Then
            If .Cols = cintColumnas Then
                If Trim(.TextMatrix(.Row, cintColTipo)) = "H" Or Trim(.TextMatrix(.Row, cintColTipo)) = "B" Or Trim(.TextMatrix(.Row, cintColTipo)) = "T" Then
                    
                    If ((cboBancoSAT.Visible And cboBancoSAT.ListIndex > 0) _
                        Or (txtBancoExtranjero.Visible And Trim(txtBancoExtranjero.Text) <> "" And Trim(txtBancoExtranjero.Text) <> "<BANCO EXTRANJERO>")) _
                            And cboCuentasPrevias.Visible Then
                                cboCuentasPrevias.Enabled = True
                                txtCuentaBancaria.Text = ""
                                txtCuentaBancaria.Enabled = False
                                lblCuentaBancaria.Enabled = True
                    Else
                        If ((cboBancoSAT.Visible And cboBancoSAT.ListIndex > 0) _
                            Or (txtBancoExtranjero.Visible And Trim(txtBancoExtranjero.Text) <> "" And Trim(txtBancoExtranjero.Text) <> "<BANCO EXTRANJERO>")) _
                                And txtCuentaBancaria.Visible Then
                                    cboCuentasPrevias.ListIndex = -1
                                    cboCuentasPrevias.Enabled = False
                                    txtCuentaBancaria.Enabled = True
                                    lblCuentaBancaria.Enabled = True
                        Else
                            cboCuentasPrevias.ListIndex = -1
                            cboCuentasPrevias.Enabled = False
                            txtCuentaBancaria.Text = ""
                            txtCuentaBancaria.Enabled = False
                            lblCuentaBancaria.Enabled = False
                        End If
                    End If
                    
                    If cboCuentasPrevias.Visible And cboCuentasPrevias.ListIndex > 0 Then
                        If cboBancoSAT.Visible And cboBancoSAT.ListIndex > 0 Then
                            MskFecha.Enabled = True
                            lblFecha.Enabled = True
                            Exit Sub
                        End If
                    End If
                                            
                    If txtCuentaBancaria.Visible And Trim(txtCuentaBancaria.Text) <> "" Then
                        If txtBancoExtranjero.Visible And Trim(txtBancoExtranjero.Text) <> "" And Trim(txtBancoExtranjero.Text) <> "<BANCO EXTRANJERO>" Then
                            MskFecha.Enabled = True
                            lblFecha.Enabled = True
                            Exit Sub
                        End If
                        If cboBancoSAT.Visible And cboBancoSAT.ListIndex > 0 Then
                            MskFecha.Enabled = True
                            lblFecha.Enabled = True
                            Exit Sub
                        End If
                    End If
    '
    '                If cboCuentasPrevias.Enabled = False And txtCuentaBancaria.Enabled = False Then
    '                    cboCuentasPrevias.ListIndex = -1
    '                    cboCuentasPrevias.Enabled = False
    '                    txtCuentaBancaria.Text = ""
    '                    txtCuentaBancaria.Enabled = False
    '                    lblCuentaBancaria.Enabled = False
    '                End If
    
                    MskFecha.Text = vldtmFecha
                    MskFecha.Enabled = False
                    lblFecha.Enabled = False
                Else
                    cboCuentasPrevias.ListIndex = -1
                    cboCuentasPrevias.Enabled = False
                    txtCuentaBancaria.Text = ""
                    txtCuentaBancaria.Enabled = False
                    lblCuentaBancaria.Enabled = False
                
                    MskFecha.Text = vldtmFecha
                    MskFecha.Enabled = False
                    lblFecha.Enabled = False
                End If
            End If
        Else
            cboCuentasPrevias.ListIndex = -1
            cboCuentasPrevias.Enabled = False
            txtCuentaBancaria.Text = ""
            txtCuentaBancaria.Enabled = False
            lblCuentaBancaria.Enabled = False
        
            MskFecha.Text = vldtmFecha
            MskFecha.Enabled = False
            lblFecha.Enabled = False
        End If
    End With
End Sub

Private Sub pCargaBancosSAT()
    'Llenado del listado de los bancos publicados por el SAT
    On Error GoTo NotificaError

    pLlenarCboSentencia cboBancoSAT, "SELECT chrclave, CASE WHEN TRIM(vchnombrecorto) = TRIM(vchnombrerazonsocial) THEN TRIM(vchnombrecorto) ELSE TRIM(vchnombrecorto) || ' - ' || TRIM(vchnombrerazonsocial) END descripcion FROM CPBANCOSAT WHERE bitactivo = 1 ORDER BY descripcion", 1, 0
    
    cboBancoSAT.AddItem "<BANCO EXTRANJERO>", 0
    cboBancoSAT.ItemData(cboBancoSAT.newIndex) = -1
    cboBancoSAT.ListIndex = 0
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCargaBancosSAT"))
End Sub

Private Sub txtRFC_GotFocus()
On Error GoTo NotificaError

    pSelTextBox txtRFC
    vlstrRFCTemporal = Trim(txtRFC.Text)

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtRFC_GotFocus"))
End Sub

Private Sub txtRFC_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
Dim vlstrcaracter As String

    With grdBusqueda
        If KeyAscii <> 8 Then
            If KeyAscii = 13 Then
                If Trim(txtRFC.Text) = "" Or (Len(Trim(txtRFC.Text)) = 12 Or Len(Trim(txtRFC.Text)) = 13) Then
                    If txtBancoExtranjero.Enabled Or cboBancoSAT.Enabled Or cboCuentasPrevias.Enabled Or txtCuentaBancaria.Enabled Or MskFecha.Enabled Or cboBanco.Enabled Or cboTipoCargoBancario.Enabled Then
                        SendKeys vbTab
                    Else
                        If Val(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColReferencia)) = 1 Then
                            txtFolio.SetFocus
                        Else
                            txtCantidad.SetFocus
                        End If
                    End If
                Else
                    If Trim(txtRFC.Text) <> "" And Len(Trim(txtRFC.Text)) <> 12 And Len(Trim(txtRFC.Text)) <> 13 Then
                        'El RFC ingresado no tiene un tamaño válido, favor de verificar:
                        MsgBox SIHOMsg(1345), vbOKOnly + vbInformation, "Mensaje"
                        txtRFC.SetFocus
                        pSelTextBox txtRFC
                    End If
                End If
            Else
                vlstrcaracter = fStrRFCValido(Chr(KeyAscii))
                If vlstrcaracter <> "" Then
                    KeyAscii = Asc(UCase(vlstrcaracter))
                Else
                    KeyAscii = 7
                End If
            End If
        End If
    End With

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtRFC_KeyPress"))
End Sub

Private Sub txtRFC_LostFocus()
    If Trim(txtRFC.Text) <> "" And Len(Trim(txtRFC.Text)) <> 12 And Len(Trim(txtRFC.Text)) <> 13 Then
        'El RFC ingresado no tiene un tamaño válido, favor de verificar:
        MsgBox SIHOMsg(1345), vbOKOnly + vbInformation, "Mensaje"
        txtRFC.SetFocus
        pSelTextBox txtRFC
    End If
End Sub

Private Function fblnFechaValidaChequeTrans() As Boolean

    fblnFechaValidaChequeTrans = True
    
    If Not IsDate(MskFecha) Then
        '¡Fecha no válida!, formato de fecha dd/mm/aaaa"
        MsgBox SIHOMsg(29), vbOKOnly + vbInformation, "Mensaje"
        MskFecha.Text = vldtmFecha
        MskFecha.SetFocus
        fblnFechaValidaChequeTrans = False
        Exit Function
    End If
    
    If Year(CDate(MskFecha.Text)) < 1900 Then
        '¡Fecha no válida!
        MsgBox SIHOMsg(254), vbOKOnly + vbExclamation, "Mensaje"
        MskFecha.Text = vldtmFecha
        MskFecha.SetFocus
        fblnFechaValidaChequeTrans = False
        Exit Function
    End If
    
    If CDate(MskFecha.Text) > vldtmfechaServer Then
        '¡La fecha debe ser menor o igual a la del sistema!
        MsgBox SIHOMsg(40), vbOKOnly + vbExclamation, "Mensaje"
        MskFecha.Text = vldtmFecha
        MskFecha.SetFocus
        fblnFechaValidaChequeTrans = False
        Exit Function
    End If
    
Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnFechaValidaChequeTrans"))
End Function

Private Sub pObtieneCtasBancariasPrevias()
    Dim rs As New ADODB.Recordset
    Dim vlintNuevaCuenta As Integer

    If vlblnLicenciaContaElectronica Or vgblnCapturaraDatosBanco Then
        If Trim(txtCuentaBancaria.Text) = "" Then
            If cboBancoSAT.ListIndex < 1 Then
                vlClaveBancoSAT = "000"
            Else
                If Len(Trim(str(cboBancoSAT.ItemData(cboBancoSAT.ListIndex)))) = 3 Then
                    vlClaveBancoSAT = Trim(str(cboBancoSAT.ItemData(cboBancoSAT.ListIndex)))
                Else
                    If Len(Trim(str(cboBancoSAT.ItemData(cboBancoSAT.ListIndex)))) = 2 Then
                        vlClaveBancoSAT = "0" & Trim(str(cboBancoSAT.ItemData(cboBancoSAT.ListIndex)))
                    Else
                        vlClaveBancoSAT = "00" & Trim(str(cboBancoSAT.ItemData(cboBancoSAT.ListIndex)))
                    End If
                End If
            End If
        
            Set rs = frsRegresaRs("SELECT intidRegistro, vchcuentabancaria FROM PVRFCCTABANCOSAT WHERE TRIM(CHRRFC) = '" & Trim(txtRFC.Text) & "' AND TRIM(CHRCLAVEBANCOSAT) = '" & vlClaveBancoSAT & "' ORDER BY vchcuentabancaria", adLockReadOnly, adOpenForwardOnly)
            If rs.RecordCount <> 0 Then
                cboCuentasPrevias.Clear

                cboCuentasPrevias.AddItem ""
                cboCuentasPrevias.ItemData(cboCuentasPrevias.newIndex) = 0
                
                Do While Not rs.EOF
                    cboCuentasPrevias.AddItem rs!VCHCUENTABANCARIA
                    cboCuentasPrevias.ItemData(cboCuentasPrevias.newIndex) = rs!intIdRegistro
                    rs.MoveNext
                Loop
                
                txtCuentaBancaria.Visible = False
                cboCuentasPrevias.ListIndex = 1
                cboCuentasPrevias.Visible = True
                
                pHabilitaFechaChequeTrans
            Else
                txtCuentaBancaria.Visible = True
                cboCuentasPrevias.Visible = False
            End If
        Else
            txtCuentaBancaria.Visible = True
            cboCuentasPrevias.Visible = False
        End If
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pObtieneCtasBancariasPrevias"))
End Sub

Private Sub txtRFC_Validate(Cancel As Boolean)
    If Trim(txtRFC.Text) <> vlstrRFCTemporal Then
        pObtieneCtasBancariasPrevias
    End If
End Sub

Private Function fblnInformacionValidaExtra() As Boolean

    fblnInformacionValidaExtra = True
    
    If vlblnLicenciaContaElectronica Or vgblnCapturaraDatosBanco Then
        If Trim(txtCuentaBancaria.Text) <> "" And (cboBancoSAT.ListIndex > 0 Or Trim(txtBancoExtranjero.Text) <> "") Then
            If Trim(txtRFC.Text) = "" Then
                'Favor de registrar el RFC.
                MsgBox SIHOMsg(1013), vbExclamation + vbOKOnly, "Mensaje"
                fblnInformacionValidaExtra = False
                txtRFC.SetFocus
                Exit Function
            End If
        End If
        
        If txtRFC.Enabled Then
            If Trim(txtRFC.Text) <> "" Then
                If Len(Trim(txtRFC.Text)) <> 12 And Len(Trim(txtRFC.Text)) <> 13 Then
                    'El RFC ingresado no tiene un tamaño válido, favor de verificar:
                    MsgBox SIHOMsg(1345), vbOKOnly + vbInformation, "Mensaje"
                    fblnInformacionValidaExtra = False
                    txtRFC.SetFocus
                    Exit Function
                End If
            End If
        End If
    
        If Trim(txtCuentaBancaria.Text) <> "" Then
            If cboBancoSAT.Visible Then
                If Not cboBancoSAT.ListIndex > 0 Then
                    'Seleccione el banco emisor del cheque o transferencia.
                    MsgBox SIHOMsg(1380), vbOKOnly + vbInformation, "Mensaje"
                    fblnInformacionValidaExtra = False
                    cboBancoSAT.SetFocus
                    Exit Function
                End If
            End If
            
            If txtBancoExtranjero.Visible Then
                If Trim(txtBancoExtranjero.Text) = "" Or Trim(txtBancoExtranjero.Text) = "<BANCO EXTRANJERO>" Then
                    '¡No ha ingresado datos!
                    MsgBox SIHOMsg(2), vbExclamation, "Mensaje"
                    fblnInformacionValidaExtra = False
                    txtBancoExtranjero.SetFocus
                    Exit Function
                End If
            End If
        End If
        
        If cboBancoSAT.Visible Then
            If cboBancoSAT.ListIndex > 0 And Trim(txtCuentaBancaria.Text) = "" And txtCuentaBancaria.Visible Then
                '¡No ha ingresado la cuenta bancaria!
                MsgBox SIHOMsg(1289), vbOKOnly + vbExclamation, "Mensaje"
                fblnInformacionValidaExtra = False
                txtCuentaBancaria.Enabled = True
                txtCuentaBancaria.SetFocus
                Exit Function
            Else
                If cboBancoSAT.ListIndex > 0 And Trim(txtCuentaBancaria.Text) <> "" And txtCuentaBancaria.Visible Then
                    If vgblnCapturaraDatosBanco Then
                        If fblnCuentaOrdenanteValida(True) = False Then
                            MsgBox SIHOMsg(1547), vbExclamation + vbOKOnly, "Mensaje"
                            'El tamaño de la cuenta bancaria emisora del pago no tiene la longitud esperada según la forma de pago seleccionada, favor de verificar.
                            fblnInformacionValidaExtra = False
                            txtCuentaBancaria.Enabled = True
                            txtCuentaBancaria.SetFocus
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
        
        If txtBancoExtranjero.Visible Then
            If txtBancoExtranjero.Text <> "" And Trim(txtCuentaBancaria.Text) = "" Then
                '¡No ha ingresado la cuenta bancaria!
                MsgBox SIHOMsg(1289), vbOKOnly + vbExclamation, "Mensaje"
                fblnInformacionValidaExtra = False
                txtCuentaBancaria.Enabled = True
                txtCuentaBancaria.SetFocus
                Exit Function
            Else
                If txtBancoExtranjero.Text <> "" And Trim(txtCuentaBancaria.Text) <> "" Then
                    If vgblnCapturaraDatosBanco Then
                        If fblnCuentaOrdenanteValida(True) = False Then
                            MsgBox SIHOMsg(1547), vbExclamation + vbOKOnly, "Mensaje"
                            'El tamaño de la cuenta bancaria emisora del pago no tiene la longitud esperada según la forma de pago seleccionada, favor de verificar.
                            fblnInformacionValidaExtra = False
                            txtCuentaBancaria.Enabled = True
                            txtCuentaBancaria.SetFocus
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
        
        If cboCuentasPrevias.Visible = True And txtCuentaBancaria.Visible = False Then
            If vgblnCapturaraDatosBanco Then
                If fblnCuentaOrdenanteValida(False) = False Then
                    MsgBox SIHOMsg(1547), vbExclamation + vbOKOnly, "Mensaje"
                    'El tamaño de la cuenta bancaria emisora del pago no tiene la longitud esperada según la forma de pago seleccionada, favor de verificar.
                    fblnInformacionValidaExtra = False
                    cboCuentasPrevias.Enabled = True
                    cboCuentasPrevias.SetFocus
                    Exit Function
                End If
            End If
        End If
        
        If cboBanco.ListIndex <> -1 Then
            If vgblnCapturaraDatosBanco Then
                If fblnCuentaBeneficiariaValida = False Then
                    MsgBox SIHOMsg(1550), vbExclamation + vbOKOnly, "Mensaje"
                    'El tamaño de la cuenta bancaria receptora del pago no tiene la longitud esperada según la forma de pago seleccionada, favor de verificar.
                    fblnInformacionValidaExtra = False
                    If cboBanco.Enabled Then
                        cboBanco.SetFocus
                    End If
                    Exit Function
                End If
            End If
        End If
    End If
    
Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnInformacionValidaExtra"))
End Function

Private Function fblnValidaFormaPagoSAT(lngCveFormaPago As Long) As Boolean
    Dim rs As ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Dim rs3 As ADODB.Recordset
    fblnValidaFormaPagoSAT = True
    If vgstrVersionCFDI = "3.2" Then
        Set rs = frsRegresaRs("select * from PVFORMAPAGO where intFormaPago = " & lngCveFormaPago)
        If Not rs.EOF Then
            If rs!chrTipo = "C" Then
                Set rs2 = frsRegresaRs("select * from CCCliente Inner join NODepartamento on CCCliente.smicvedepartamento = nodepartamento.smicvedepartamento where CCCliente.intNumReferencia = " & vlLngReferencia & " and CCCliente.chrTipoCliente = '" & vlstrTipoCliente & "' and CCCliente.bitActivo = 1 and nodepartamento.tnyclaveempresa = " & vgintClaveEmpresaContable)
                If IsNull(rs2!VCHTIPOPAGO) Then
                    fblnValidaFormaPagoSAT = False
                Else
                    Set rs3 = frsRegresaRs("select * from PVMETODOPAGOSATCFDI where CHRCLAVE = '" & rs2!VCHTIPOPAGO & "'")
                    If Not rs3.EOF Then
                        fblnValidaFormaPagoSAT = True
                    Else
                        fblnValidaFormaPagoSAT = False
                    End If
                    rs3.Close
                End If
                rs2.Close
            Else
                If IsNull(rs!VCHDESCRIPCIONCFD) Then
                    fblnValidaFormaPagoSAT = False
                Else
                    Set rs3 = frsRegresaRs("select * from PVMETODOPAGOSATCFDI where CHRCLAVE = '" & rs!VCHDESCRIPCIONCFD & "'")
                    If Not rs3.EOF Then
                        fblnValidaFormaPagoSAT = True
                    Else
                        fblnValidaFormaPagoSAT = False
                    End If
                    rs3.Close
                End If
            End If
        End If
        rs.Close
    Else
        Set rs = frsRegresaRs("select * from PVFormaPago where intFormaPago = " & lngCveFormaPago)
        If Not rs.EOF Then
            If rs!chrTipo <> "C" Then
                Set rs2 = frsRegresaRs("select R.intIdRegistro from GNCatalogoSATRelacion R inner join GNCatalogoSATDetalle D on R.intIdRegistro = D.intIdRegistro where R.intCveConcepto = " & lngCveFormaPago & " and R.chrTipoConcepto = 'FP' and D.bitActivo <> 0", adLockReadOnly, adOpenForwardOnly)
                If rs2.EOF Then
                    fblnValidaFormaPagoSAT = False
                End If
                rs2.Close
            End If
        End If
        rs.Close
    End If
End Function

Private Function fblnCuentaOrdenanteValida(vlblnTexto As Boolean) As Boolean
    fblnCuentaOrdenanteValida = False
    
    If vlblnTexto Then
        If Trim(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColClaveFormaPagoSAT)) = "" Then
            fblnCuentaOrdenanteValida = True
        Else
            If Trim(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColClaveFormaPagoSAT)) = "02" Then
                If Len(Trim(txtCuentaBancaria.Text)) = 11 Or Len(Trim(txtCuentaBancaria.Text)) = 18 Then
                    fblnCuentaOrdenanteValida = True
                Else
                    fblnCuentaOrdenanteValida = False
                End If
            Else
                If Trim(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColClaveFormaPagoSAT)) = "03" Then
                    If Len(Trim(txtCuentaBancaria.Text)) = 10 Or Len(Trim(txtCuentaBancaria.Text)) = 16 Or Len(Trim(txtCuentaBancaria.Text)) = 18 Then
                        fblnCuentaOrdenanteValida = True
                    Else
                        fblnCuentaOrdenanteValida = False
                    End If
                Else
                    If Trim(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColClaveFormaPagoSAT)) = "04" Then
                        If Len(Trim(txtCuentaBancaria.Text)) = 16 Then
                            fblnCuentaOrdenanteValida = True
                        Else
                            fblnCuentaOrdenanteValida = False
                        End If
                    Else
                        If Trim(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColClaveFormaPagoSAT)) = "05" Then
                            If Len(Trim(txtCuentaBancaria.Text)) = 10 Or Len(Trim(txtCuentaBancaria.Text)) = 11 Or Len(Trim(txtCuentaBancaria.Text)) = 15 Or Len(Trim(txtCuentaBancaria.Text)) = 16 Or Len(Trim(txtCuentaBancaria.Text)) = 18 Or (Len(Trim(txtCuentaBancaria.Text)) >= 10 And Len(Trim(txtCuentaBancaria.Text)) <= 50) Then
                                fblnCuentaOrdenanteValida = True
                            Else
                                fblnCuentaOrdenanteValida = False
                            End If
                        Else
                            If Trim(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColClaveFormaPagoSAT)) = "06" Then
                                If Len(Trim(txtCuentaBancaria.Text)) = 10 Then
                                    fblnCuentaOrdenanteValida = True
                                Else
                                    fblnCuentaOrdenanteValida = False
                                End If
                            Else
                                If Trim(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColClaveFormaPagoSAT)) = "28" Then
                                    If Len(Trim(txtCuentaBancaria.Text)) = 16 Then
                                        fblnCuentaOrdenanteValida = True
                                    Else
                                        fblnCuentaOrdenanteValida = False
                                    End If
                                Else
                                    If Trim(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColClaveFormaPagoSAT)) = "29" Then
                                        If Len(Trim(txtCuentaBancaria.Text)) = 15 Or Len(Trim(txtCuentaBancaria.Text)) = 16 Then
                                            fblnCuentaOrdenanteValida = True
                                        Else
                                            fblnCuentaOrdenanteValida = False
                                        End If
                                    Else
                                        fblnCuentaOrdenanteValida = True
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Else
        If Trim(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColClaveFormaPagoSAT)) = "" Then
            fblnCuentaOrdenanteValida = True
        Else
            If Trim(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColClaveFormaPagoSAT)) = "02" Then
                If Len(Trim(cboCuentasPrevias.Text)) = 11 Or Len(Trim(cboCuentasPrevias.Text)) = 18 Then
                    fblnCuentaOrdenanteValida = True
                Else
                    fblnCuentaOrdenanteValida = False
                End If
            Else
                If Trim(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColClaveFormaPagoSAT)) = "03" Then
                    If Len(Trim(cboCuentasPrevias.Text)) = 10 Or Len(Trim(cboCuentasPrevias.Text)) = 16 Or Len(Trim(cboCuentasPrevias.Text)) = 18 Then
                        fblnCuentaOrdenanteValida = True
                    Else
                        fblnCuentaOrdenanteValida = False
                    End If
                Else
                    If Trim(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColClaveFormaPagoSAT)) = "04" Then
                        If Len(Trim(cboCuentasPrevias.Text)) = 16 Then
                            fblnCuentaOrdenanteValida = True
                        Else
                            fblnCuentaOrdenanteValida = False
                        End If
                    Else
                        If Trim(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColClaveFormaPagoSAT)) = "05" Then
                            If Len(Trim(cboCuentasPrevias.Text)) = 10 Or Len(Trim(cboCuentasPrevias.Text)) = 11 Or Len(Trim(cboCuentasPrevias.Text)) = 15 Or Len(Trim(cboCuentasPrevias.Text)) = 16 Or Len(Trim(cboCuentasPrevias.Text)) = 18 Or (Len(Trim(cboCuentasPrevias.Text)) >= 10 And Len(Trim(cboCuentasPrevias.Text)) <= 50) Then
                                fblnCuentaOrdenanteValida = True
                            Else
                                fblnCuentaOrdenanteValida = False
                            End If
                        Else
                            If Trim(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColClaveFormaPagoSAT)) = "06" Then
                                If Len(Trim(cboCuentasPrevias.Text)) = 10 Then
                                    fblnCuentaOrdenanteValida = True
                                Else
                                    fblnCuentaOrdenanteValida = False
                                End If
                            Else
                                If Trim(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColClaveFormaPagoSAT)) = "28" Then
                                    If Len(Trim(cboCuentasPrevias.Text)) = 16 Then
                                        fblnCuentaOrdenanteValida = True
                                    Else
                                        fblnCuentaOrdenanteValida = False
                                    End If
                                Else
                                    If Trim(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColClaveFormaPagoSAT)) = "29" Then
                                        If Len(Trim(cboCuentasPrevias.Text)) = 15 Or Len(Trim(cboCuentasPrevias.Text)) = 16 Then
                                            fblnCuentaOrdenanteValida = True
                                        Else
                                            fblnCuentaOrdenanteValida = False
                                        End If
                                    Else
                                        fblnCuentaOrdenanteValida = True
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    
End Function

Private Function fblnCuentaBeneficiariaValida() As Boolean
    Dim rs As New ADODB.Recordset
    Dim vlstrCuenta As String
    
    fblnCuentaBeneficiariaValida = False
    
    vlstrCuenta = ""
    Set rs = frsRegresaRs("select vchcuentabancaria from cpbanco where tnynumerobanco = " & cboBanco.ItemData(cboBanco.ListIndex), adLockReadOnly, adOpenForwardOnly)
    If rs.RecordCount <> 0 Then
        vlstrCuenta = Trim(IIf(IsNull(rs!VCHCUENTABANCARIA), "", rs!VCHCUENTABANCARIA))
    Else
        vlstrCuenta = ""
    End If
    
    If Trim(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColClaveFormaPagoSAT)) = "" Then
        fblnCuentaBeneficiariaValida = True
    Else
        If Trim(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColClaveFormaPagoSAT)) = "02" Then
            If Len(vlstrCuenta) = 10 Or Len(vlstrCuenta) = 11 Or Len(vlstrCuenta) = 15 Or Len(vlstrCuenta) = 16 Or Len(vlstrCuenta) = 18 Then
                fblnCuentaBeneficiariaValida = True
            Else
                fblnCuentaBeneficiariaValida = False
            End If
        Else
            If Trim(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColClaveFormaPagoSAT)) = "03" Then
                If Len(vlstrCuenta) = 10 Or Len(vlstrCuenta) = 18 Then
                    fblnCuentaBeneficiariaValida = True
                Else
                    fblnCuentaBeneficiariaValida = False
                End If
            Else
                If Trim(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColClaveFormaPagoSAT)) = "04" Then
                    If Len(vlstrCuenta) = 10 Or Len(vlstrCuenta) = 11 Or Len(vlstrCuenta) = 15 Or Len(vlstrCuenta) = 16 Or Len(vlstrCuenta) = 18 Then
                        fblnCuentaBeneficiariaValida = True
                    Else
                        fblnCuentaBeneficiariaValida = False
                    End If
                Else
                    If Trim(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColClaveFormaPagoSAT)) = "05" Then
                        If Len(vlstrCuenta) = 10 Or Len(vlstrCuenta) = 11 Or Len(vlstrCuenta) = 15 Or Len(vlstrCuenta) = 16 Or Len(vlstrCuenta) = 18 Then
                            fblnCuentaBeneficiariaValida = True
                        Else
                            fblnCuentaBeneficiariaValida = False
                        End If
                    Else
                        If Trim(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColClaveFormaPagoSAT)) = "06" Then
                            fblnCuentaBeneficiariaValida = True
                        Else
                            If Trim(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColClaveFormaPagoSAT)) = "28" Then
                                If Len(vlstrCuenta) = 10 Or Len(vlstrCuenta) = 11 Or Len(vlstrCuenta) = 15 Or Len(vlstrCuenta) = 16 Or Len(vlstrCuenta) = 18 Then
                                    fblnCuentaBeneficiariaValida = True
                                Else
                                    fblnCuentaBeneficiariaValida = False
                                End If
                            Else
                                If Trim(grdFormasPago.TextMatrix(grdFormasPago.Row, cintColClaveFormaPagoSAT)) = "29" Then
                                    If Len(vlstrCuenta) = 10 Or Len(vlstrCuenta) = 11 Or Len(vlstrCuenta) = 15 Or Len(vlstrCuenta) = 16 Or Len(vlstrCuenta) = 18 Then
                                        fblnCuentaBeneficiariaValida = True
                                    Else
                                        fblnCuentaBeneficiariaValida = False
                                    End If
                                Else
                                    fblnCuentaBeneficiariaValida = True
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    
End Function

Private Sub ws_Answer(answ As String)
    strRespuestaEsperar = answ
   
    blnRespuestaEsperar = True

   
End Sub

'-------------------------------------------------------------------------------|
'-------------------------------------------------------------------------------|
'------ Funciones para configurar la pantalla sin mostrarla (Facturación masiva)|
'-------------------------------------------------------------------------------|
'-------------------------------------------------------------------------------|

Public Sub pConfiguraMasiva(vlstrRFC As String, vlstrFolio As String, vlstrFecha As String, vlstrCuenta As String, tipoPago As String)
    txtRFC.Text = vlstrRFC
    txtFolio.Text = vlstrFolio
    MskFecha.Text = vlstrFecha
    If tipoPago = "H" Or tipoPago = "B" Or tipoPago = "T" Then
        txtCuentaBancaria.Text = vlstrCuenta
    End If

End Sub

Public Sub pSeleccionaPago(vlstrCveFormaPago As String)
    
    intcontador = 1
    For intcontador = 1 To grdFormasPago.Rows - 1
        If grdFormasPago.TextMatrix(intcontador, cintColIdForma) = vlstrCveFormaPago Then
            With grdFormasPago
                .Row = intcontador
                .Col = 0
                .ColSel = .Cols - 1
            End With
            grdFormasPago.Redraw = True
            Call grdFormasPago_Click
            Exit Sub
        End If
    Next intcontador
    
End Sub

Public Sub pSeleccionaBancoSAT(vlstrCveBancoSAT As String, tipoPago As String)
    
    If tipoPago = "H" Or tipoPago = "B" Or tipoPago = "T" Then
    
        intcontador = 1
        For intcontador = 0 To cboBancoSAT.ListCount - 1
            If cboBancoSAT.ItemData(intcontador) = vlstrCveBancoSAT Then
                cboBancoSAT.ListIndex = intcontador
                Exit Sub
            End If
        Next intcontador
    
    End If
End Sub

Public Sub pSeleccionaCuentaBancaria(vlstrClaveCuentaBancaria As String, tipoPago As String)
    
    If tipoPago = "H" Or tipoPago = "B" Or tipoPago = "T" Then
    
        intcontador = 1
        For intcontador = 0 To cboBanco.ListCount - 1
            If cboBanco.ItemData(intcontador) = vlstrClaveCuentaBancaria Then
                cboBanco.ListIndex = intcontador
                Exit Sub
            End If
        Next intcontador
    
    End If
End Sub

Public Sub pSeleccionaTipoCargo(vlStrTipoCargoBancario As String, tipoPago As String)
    
    If (tipoPago = "H" Or tipoPago = "B" Or tipoPago = "T") And cboTipoCargoBancario.Enabled = True Then
    
        intcontador = 1
        For intcontador = 0 To cboTipoCargoBancario.ListCount - 1
            If cboTipoCargoBancario.ItemData(intcontador) = vlStrTipoCargoBancario Then
                cboTipoCargoBancario.ListIndex = intcontador
                Exit Sub
            End If
        Next intcontador
    
    End If
End Sub

Public Function pRegLog(lngId, strMessage, intCveTerminal, strTipo As String, lngCveFormaPago As Long) As Long
     Dim rs As ADODB.Recordset
     pRegLog = 0
     Set rs = frsRegresaRs("select * from PVTerminalLog where intID = " & lngId, adLockOptimistic, adOpenStatic)
     If rs.EOF Then
        rs.AddNew
        rs!intCveTerminal = intCveTerminal
        rs!intCveFormaPago = lngCveFormaPago
        rs.Update
        pRegLog = flngObtieneIdentity("SEC_PVTERMINALLOG", 1)
     Else
        If strTipo = "T" Then
            rs!VCHMESSAGETX = strMessage
            rs!DTMDATETX = Now
        Else
            rs!VCHMESSAGERX = strMessage
            rs!dtmDateRX = Now
        End If
        pRegLog = lngId
        rs.Update
     End If
     rs.Close
End Function


Private Sub pReimprimir(strTipo As String)
    On Error GoTo Errs
    Dim X As Integer
    Dim strResp As String
    Dim tmpData As String
    Dim arrDatos() As String
    Dim intPpProv As Integer
    Dim vgrptReporte As CRAXDRT.Report
    Dim rsDummy As ADODB.Recordset
    Dim alstrParametros(22) As String
    For X = 1 To grdFormas.Rows - 1
        If grdFormas.TextMatrix(X, cintColUsarPinpad2) = "3" Then
            intPpProv = IIf(grdFormas.TextMatrix(X, cintColPpProv2) = "", 1, CInt(grdFormas.TextMatrix(X, cintColPpProv2)))
            strResp = fstrPinPad2(grdFormas.TextMatrix(X, cintColUriPinpad2), IIf(grdFormas.TextMatrix(X, cintColCveMoneda2) = "", "484", grdFormas.TextMatrix(X, cintColCveMoneda2)), intPpProv, grdFormas.TextMatrix(X, cIntColFolio), grdFormas.TextMatrix(X, cintColPpHost2), grdFormas.TextMatrix(X, cintColPpPort2), grdFormas.TextMatrix(X, cintColPpUsr2), grdFormas.TextMatrix(X, cintColPpPwd2), CLng(grdFormas.TextMatrix(X, cintColPpCve2)), grdFormas.RowData(X))
            If strResp <> "" Then
                arrDatos = Split(strResp, "|")
                tmpData = fstrGetPPData(arrDatos, "trn_internal_respcode")
                If tmpData = "-1" Then
                
                    grdFormas.TextMatrix(X, cintColUsarPinpad2) = "2"
                    grdFormas.TextMatrix(X, cIntColFolio) = fstrGetPPData(arrDatos, "trn_auth_code")
                    
                    Set rsDummy = frsRegresaRs("select sysdate from dual", adLockReadOnly, adOpenForwardOnly)
                    
                    tmpData = fstrGetPPData(arrDatos, "trn_qty_pay")
                    alstrParametros(0) = "Operacion;REIMPRESION VENTA"
                    alstrParametros(1) = "mer_legend1;" & fstrGetPPData(arrDatos, "mer_legend1")
                    alstrParametros(2) = "mer_legend2;" & fstrGetPPData(arrDatos, "mer_legend2")
                    alstrParametros(3) = "mer_legend3;" & fstrGetPPData(arrDatos, "mer_legend3")
                    alstrParametros(4) = "trn_external_mer_id;" & fstrGetPPData(arrDatos, "trn_external_mer_id")
                    alstrParametros(5) = "trn_external_ter_id;" & fstrGetPPData(arrDatos, "trn_external_ter_id")
                    alstrParametros(6) = "trn_fechaTrans;" & fstrGetPPData(arrDatos, "trn_fechaTrans")
                    alstrParametros(7) = "Copia;0"
                    alstrParametros(8) = "trn_label;" & fstrGetPPData(arrDatos, "trn_label")
                    alstrParametros(9) = "trn_aprnam;" & fstrGetPPData(arrDatos, "trn_aprnam")
                    alstrParametros(10) = "trn_emv_cryptogram;" & fstrGetPPData(arrDatos, "trn_emv_cryptogram")
                    alstrParametros(11) = "trn_AID;" & fstrGetPPData(arrDatos, "trn_AID")
                    alstrParametros(12) = "trn_pro_name;" & fstrGetPPData(arrDatos, "trn_pro_name")
                    alstrParametros(13) = "trn_aco_id;" & fstrGetPPData(arrDatos, "trn_aco_id")
                    alstrParametros(14) = "trn_auth_code;" & fstrGetPPData(arrDatos, "trn_auth_code")
                    alstrParametros(15) = "trn_id;" & fstrGetPPData(arrDatos, "trn_id")
                    alstrParametros(16) = "trn_amount;" & fstrGetPPData(arrDatos, "trn_amount")
                    alstrParametros(17) = "Compra;" & IIf(tmpData = "1", "COMPRA NORMAL", tmpData & " MESES SIN INTERESES")
                    alstrParametros(18) = "trn_fe;" & fstrGetPPData(arrDatos, "trn_fe")
                    alstrParametros(19) = "trn_internal_ter_id;" & fstrGetPPData(arrDatos, "trn_internal_ter_id")
   
                    pInstanciaReporte vgrptReporte, "voucher.rpt"
   
                    vgrptReporte.DiscardSavedData
                    fblnAsignaImpresoraReportePorNombre grdFormas.TextMatrix(X, cintColImprVoucher2), vgrptReporte
                    pCargaParameterFields alstrParametros, vgrptReporte
                    pImprimeReporte vgrptReporte, rsDummy, "I", "Voucher venta"
                    
                    alstrParametros(7) = "Copia;1"
                    pCargaParameterFields alstrParametros, vgrptReporte
                    pImprimeReporte vgrptReporte, rsDummy, "I", "Voucher venta"
                Else
                    tmpData = fstrGetPPData(arrDatos, "trn_msg_host")
                    MsgBox tmpData, vbExclamation, "Mensaje"
                End If
            
            End If
           
        End If
    Next
    Exit Sub
Errs:
     MsgBox Err.Description, vbExclamation, "Mensaje"
End Sub

Private Function fstrPinPad2(strUriPinpad As String, strMoneda As String, intPpProvider As Integer, strReferencia As String, strHost As String, strPort As String, strUsr As String, strPwd As String, lngCve As Long, lngCveFormaPago As Long) As String
    On Error GoTo Errs
    Dim intRespLen As Long
    Dim strReturn As String
    Dim lngIdLog As Long
    Set ws = New WebSocketWrap.Client
    ws.Timeout = intTimeout
    ws.Uri = strUriPinpad & "?host=" & strHost & "&port=" & strPort & "&prov=" & IIf(intPpProvider = 1, "FISERV", "EVO") & "&usr=" & strUsr & "&pwd=" & strPwd
   
    blnRespuestaEsperar = False
    lngIdLog = pRegLog(0, "", lngCve, "N", lngCveFormaPago)
    Select Case intPpProvider
        Case 2
            lngIdLog = pRegLog(lngIdLog, "PRINT060:" & strMoneda & ":" & strReferencia, lngCve, "T", lngCveFormaPago)
            ws.SendMessage "PRINT060:" & strMoneda & ":" & strReferencia
    End Select
    
    Do While Not blnRespuestaEsperar
        DoEvents
    Loop
    strReturn = strRespuestaEsperar
    If intPpProvider = 1 Then
        intRespLen = CInt(Mid(strReturn, 13, 4)) - 1
        strReturn = Replace(strReturn, ":", "=")
        strReturn = Mid(strReturn, 18, intRespLen)
    ElseIf intPpProvider = 2 Then
        intRespLen = InStr(strReturn, "}") - 14
        strReturn = Mid(strReturn, 13, intRespLen)
        strReturn = Replace(strReturn, "|Respuesta=", "")
        strReturn = Replace(strReturn, "&", "|")
    End If
    pRegLog lngIdLog, strReturn, lngCve, "R", lngCveFormaPago
    If InStr(strReturn, "Error de socket") > 0 Then
        MsgBox "Error de conexión con el socket:" & vbCrLf & strHost & ":" & strPort, vbExclamation, "Mensaje"
        fstrPinPad2 = ""
    Else
        fstrPinPad2 = strReturn
    End If
    
    Exit Function
Errs:
    fstrPinPad2 = ""
    If InStr(Err.Description, "Error de conexión") > 0 Then
        pRegLog lngIdLog, "Error de conexión con el Web Socket: " & strUriPinpad, lngCve, "R", lngCveFormaPago
        MsgBox "Error de conexión con el Web Socket: " & vbCrLf & strUriPinpad, vbExclamation, "Mensaje"
    Else
        pRegLog lngIdLog, Err.Description, lngCve, "R", lngCveFormaPago
        MsgBox Err.Description, vbExclamation, "Mensaje"
    End If
End Function


Public Sub pReceptoraBanco(IntBanco As Long)

     Set rsBanco = frsRegresaRs("Select distinct VCHNOMBREBANCO from cpbanco where cpbanco.INTNUMEROCUENTA = " & IntBanco)
     cboBanco.Enabled = True
     If rsBanco.RecordCount > 0 Then
        For i = 0 To cboBanco.ListCount
            If rsBanco!VCHNOMBREBANCO = cboBanco.List(i) Then
                cboBanco.ListIndex = i
                cboBanco.Enabled = False
                Exit For
            End If
        Next i
    End If
End Sub
