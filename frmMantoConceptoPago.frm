VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMantoConceptoPago 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conceptos de entradas y salidas de dinero"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8415
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   8415
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab sstabConcepto 
      Height          =   4320
      Left            =   -60
      TabIndex        =   22
      Top             =   -555
      Width           =   8640
      _ExtentX        =   15240
      _ExtentY        =   7620
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmMantoConceptoPago.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmMantoConceptoPago.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "frmMantoConceptoPago.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame5"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      Begin VB.Frame Frame5 
         Height          =   690
         Left            =   -71583
         TabIndex        =   37
         Top             =   3120
         Width           =   1580
         Begin VB.CommandButton cmdRegresar 
            Caption         =   "Regresar"
            Height          =   495
            Left            =   560
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Regresar a la pantalla principal"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   975
         End
         Begin VB.CommandButton cmdBorrarCuenta 
            Enabled         =   0   'False
            Height          =   495
            Left            =   50
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmMantoConceptoPago.frx":0054
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Borrar cuenta contable"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
      End
      Begin VB.Frame Frame4 
         Height          =   2250
         Left            =   -74870
         TabIndex        =   31
         Top             =   570
         Width           =   8280
         Begin VB.CheckBox chkConceptoLiquidacion 
            Caption         =   "Concepto de liquidación"
            Height          =   195
            Left            =   1440
            TabIndex        =   19
            Top             =   1800
            Width           =   2130
         End
         Begin VB.ComboBox cboEmpresas 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmMantoConceptoPago.frx":0756
            Left            =   1440
            List            =   "frmMantoConceptoPago.frx":0769
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   795
            Width           =   6645
         End
         Begin MSMask.MaskEdBox mskCuenta 
            Height          =   315
            Left            =   1440
            TabIndex        =   18
            ToolTipText     =   "Cuenta contable"
            Top             =   1245
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            PromptChar      =   " "
         End
         Begin VB.Label lblConcepto 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   315
            Left            =   1440
            TabIndex        =   36
            ToolTipText     =   "Descripción del concepto de entrada/salida de dinero"
            Top             =   360
            Width           =   6645
         End
         Begin VB.Label lblDescripcionCuenta 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   315
            Left            =   3405
            TabIndex        =   35
            Top             =   1245
            Width           =   4680
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta contable"
            Height          =   195
            Left            =   120
            TabIndex        =   34
            Top             =   1300
            Width           =   1170
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Descripción"
            Height          =   195
            Left            =   120
            TabIndex        =   33
            Top             =   400
            Width           =   840
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Empresa"
            Height          =   195
            Left            =   120
            TabIndex        =   32
            Top             =   850
            Width           =   615
         End
      End
      Begin VB.Frame Frame3 
         Height          =   3690
         Left            =   -74910
         TabIndex        =   27
         Top             =   495
         Width           =   8355
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdConceptos 
            Height          =   3450
            Left            =   60
            TabIndex        =   38
            Top             =   165
            Width           =   8220
            _ExtentX        =   14499
            _ExtentY        =   6085
            _Version        =   393216
            GridColor       =   12632256
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.Frame Frame2 
         Height          =   705
         Left            =   1770
         TabIndex        =   26
         Top             =   3480
         Width           =   4875
         Begin VB.CommandButton cmdCuentascontables 
            Caption         =   "Cuentas contables"
            Height          =   495
            Left            =   3600
            TabIndex        =   9
            ToolTipText     =   "Configurar cuentas contables por empresa"
            Top             =   150
            Width           =   1215
         End
         Begin VB.CommandButton cmdTop 
            Height          =   495
            Left            =   60
            Picture         =   "frmMantoConceptoPago.frx":0802
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Primer registro"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdBack 
            Height          =   495
            Left            =   570
            Picture         =   "frmMantoConceptoPago.frx":0924
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Anterior registro"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdLocate 
            Height          =   495
            Left            =   1080
            Picture         =   "frmMantoConceptoPago.frx":0A96
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Búsqueda"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdNext 
            Height          =   495
            Left            =   1590
            Picture         =   "frmMantoConceptoPago.frx":0C08
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Siguiente registro"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdEnd 
            Height          =   495
            Left            =   2100
            Picture         =   "frmMantoConceptoPago.frx":0D7A
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Ultimo registro"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdSave 
            Height          =   495
            Left            =   2595
            Picture         =   "frmMantoConceptoPago.frx":0EEC
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Grabar"
            Top             =   150
            Width           =   495
         End
         Begin VB.CommandButton cmdDelete 
            Height          =   495
            Left            =   3105
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmMantoConceptoPago.frx":122E
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Borrar"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2850
         Left            =   140
         TabIndex        =   23
         Top             =   570
         Width           =   8280
         Begin VB.OptionButton OptDesglosarIVAoExento 
            Caption         =   "No desglosa IVA"
            Height          =   255
            Index           =   2
            Left            =   3120
            TabIndex        =   4
            ToolTipText     =   "No desglosa el IVA cuando se realice un pago"
            Top             =   1560
            Width           =   1575
         End
         Begin VB.OptionButton OptDesglosarIVAoExento 
            Caption         =   "Exento de IVA"
            Height          =   255
            Index           =   1
            Left            =   4920
            TabIndex        =   5
            ToolTipText     =   "El concepto será exento de IVA"
            Top             =   1560
            Width           =   1335
         End
         Begin VB.OptionButton OptDesglosarIVAoExento 
            Caption         =   "Desglosar IVA"
            Height          =   255
            Index           =   0
            Left            =   1560
            TabIndex        =   3
            ToolTipText     =   "Desglosar el IVA cuando se realice un pago"
            Top             =   1560
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.CheckBox chkcfdi 
            Caption         =   "Generar comprobante fiscal"
            Height          =   255
            Left            =   1560
            TabIndex        =   6
            ToolTipText     =   "Generar comprobante fiscal"
            Top             =   1920
            Width           =   2415
         End
         Begin VB.CheckBox chkPagoCancelarFactura 
            Caption         =   "Concepto para generar pagos automáticos al cancelar facturas"
            Enabled         =   0   'False
            Height          =   195
            Left            =   1560
            TabIndex        =   7
            ToolTipText     =   "Marcar para usar este concepto para generar un pago automático por el importe que afectó a bancos al cancelar una factura"
            Top             =   2235
            Width           =   4845
         End
         Begin VB.ComboBox cboTipoConcepto 
            Height          =   315
            ItemData        =   "frmMantoConceptoPago.frx":13D0
            Left            =   1560
            List            =   "frmMantoConceptoPago.frx":13D2
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   1130
            Width           =   4590
         End
         Begin VB.CheckBox chkActivo 
            Caption         =   "Concepto activo"
            Height          =   195
            Left            =   1560
            TabIndex        =   8
            ToolTipText     =   "Estado"
            Top             =   2520
            Width           =   1650
         End
         Begin VB.TextBox txtClave 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1560
            MaxLength       =   9
            TabIndex        =   0
            ToolTipText     =   "Clave del concepto "
            Top             =   240
            Width           =   945
         End
         Begin VB.TextBox txtDescripcion 
            Height          =   315
            Left            =   1560
            MaxLength       =   50
            TabIndex        =   1
            ToolTipText     =   "Descripción del concepto "
            Top             =   670
            Width           =   6555
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de concepto"
            Height          =   195
            Left            =   120
            TabIndex        =   29
            Top             =   1190
            Width           =   1260
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Estado"
            Height          =   195
            Left            =   240
            TabIndex        =   28
            Top             =   2520
            Width           =   495
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Clave"
            Height          =   195
            Left            =   120
            TabIndex        =   25
            Top             =   300
            Width           =   405
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Descripción"
            Height          =   195
            Left            =   120
            TabIndex        =   24
            Top             =   730
            Width           =   840
         End
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grdCuentas 
      CausesValidation=   0   'False
      Height          =   795
      Left            =   120
      TabIndex        =   30
      Top             =   3960
      Width           =   4095
      _cx             =   7223
      _cy             =   1402
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   0
      HighLight       =   0
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   0
      Cols            =   4
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   0
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   0
      AutoSizeMouse   =   0   'False
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
End
Attribute VB_Name = "frmMantoConceptoPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------
' Programa para dar mantenimiento a los conceptos de entrada/salida de dinero (PvConceptoPago)
' Fecha de programación: Lunes 26 de Febrero de 2001
'---------------------------------------------------------------------------
' Ultimas modificaciones, especificar:
' 28/Abril/2003 Se cambio de nombre a la pantalla para que fueran "entradas/salidas"
'               de dinero en vez de "Conceptos de Pago"
'               Se incluyó un bit para diferenciar si ese concepto es de "Salida de dinero"
''---------------------------------------------------------------------------
Option Explicit

Dim rsPvConceptoPago As New ADODB.Recordset
Dim rsCuentasConceptoES As New ADODB.Recordset
Dim rsConceptoLiq As New ADODB.Recordset
Dim vlstrx As String
Dim vlstrSentenciaConsulta As String
Dim vlblnConsulta As Boolean
Dim lngContador As Long
Dim lngRow As Long
Dim strSentencia As String
Dim vlintbitPagoCancelaFactura As Integer

Private Sub cboEmpresas_Click()
On Error GoTo NotificaError

    mskCuenta.Mask = ""
    mskCuenta.Text = ""
    lblDescripcionCuenta.Caption = ""

    For lngContador = 0 To grdCuentas.Rows - 1
        If CLng(grdCuentas.TextMatrix(lngContador, 0)) = cboEmpresas.ItemData(cboEmpresas.ListIndex) Then
            If grdCuentas.TextMatrix(lngContador, 1) > 0 Then
                mskCuenta.Text = fstrCuentaContable(grdCuentas.TextMatrix(lngContador, 1))
                lblDescripcionCuenta.Caption = fstrDescripcionCuenta(mskCuenta.Text, grdCuentas.TextMatrix(lngContador, 0))
            End If
            mskCuenta.Mask = Trim(grdCuentas.TextMatrix(lngContador, 2))
            lngRow = lngContador
            chkConceptoLiquidacion.Value = CInt(grdCuentas.TextMatrix(lngContador, 3))
            
        End If
    Next lngContador
    
    mskCuenta.Enabled = False
    If fblnRevisaPermiso(vglngNumeroLogin, IIf(cgstrModulo = "PV", 316, 1122), "C", True) Then
        mskCuenta.Enabled = fConceptoCuentaModificable
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboEmpresas_Click"))
End Sub

Private Sub cboTipoConcepto_Click()
On Error GoTo NotificaError
    
    'Si se seleccionó "CONCEPTO NORMAL"
    If cboTipoConcepto.ListIndex <> -1 Then chkConceptoLiquidacion.Visible = IIf(cboTipoConcepto.ItemData(cboTipoConcepto.ListIndex) = 0, True, False)
            
    If cboTipoConcepto.ListIndex = 0 And vlintbitPagoCancelaFactura = 1 Then
        'Si se seleccionó "CONCEPTO NORMAL" y PvConceptoPago.bitPagoCancelaFactura está activado
        chkPagoCancelarFactura.Enabled = True
    Else
        chkPagoCancelarFactura.Value = 0
        chkPagoCancelarFactura.Enabled = False
    End If
    
    chkcfdi.Enabled = cboTipoConcepto.ListIndex <> 5
        
    If Not chkcfdi.Enabled Then chkcfdi.Value = 0
    
Exit Sub
NotificaError:
    chkConceptoLiquidacion.Visible = False
End Sub

Private Sub cboTipoConcepto_GotFocus()
    pHabilita 0, 0, 0, 0, 0, 1, 0
End Sub

Private Sub chkActivo_GotFocus()
    On Error GoTo NotificaError
    pHabilita 0, 0, 0, 0, 0, 1, 0
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":chkActivo_GotFocus"))
End Sub

Private Sub chkcfdi_GotFocus()
    On Error GoTo NotificaError
    pHabilita 0, 0, 0, 0, 0, 1, 0
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":chkcfdi_GotFocus"))

End Sub

Private Sub chkConceptoLiquidacion_Click()
    On Error GoTo NotificaError
    
    If Trim(txtDescripcion.Text) <> "" Then
        grdCuentas.TextMatrix(lngRow, 3) = CStr(chkConceptoLiquidacion.Value)
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":chkConceptoLiquidacion_Click"))
End Sub

Private Sub chkConceptoLiquidacion_LostFocus()
    On Error GoTo NotificaError
    Dim rsLiquidacion As New ADODB.Recordset
    
    If chkConceptoLiquidacion.Value = 1 Then
        
        strSentencia = "Select Count(*) Co From PvConceptoPagoEmpresa " & _
                       "Where PvConceptoPagoEmpresa.bitConceptoLiquidacion = 1 " & _
                       "And PvConceptoPagoEmpresa.intCveEmpresa = " & CStr(cboEmpresas.ItemData(cboEmpresas.ListIndex)) & " " & _
                       "And PvConceptoPagoEmpresa.intNumConcepto <> " & txtClave.Text
                       
        Set rsConceptoLiq = frsRegresaRs(strSentencia, adLockReadOnly, adOpenForwardOnly)
        If rsConceptoLiq.RecordCount > 0 Then
            If rsConceptoLiq!CO > 0 Then
                'Ya se estableció un concepto de liquidación.
                MsgBox SIHOMsg(791), vbCritical, "Mensaje"
                chkConceptoLiquidacion.Value = 0
                Exit Sub
            End If
        End If
        
        grdCuentas.TextMatrix(lngRow, 3) = 1
    Else
        strSentencia = "select nvl(bitConceptoLiquidacion,0) bitLiquidacion from pvConceptoPagoEmpresa where intNumConcepto = " & txtClave.Text _
                     & " and intCveEmpresa = " & CStr(cboEmpresas.ItemData(cboEmpresas.ListIndex)) _
                     & " and intNumeroCuenta = " & grdCuentas.TextMatrix(lngRow, 1)
        Set rsLiquidacion = frsRegresaRs(strSentencia)
        If rsLiquidacion.RecordCount <> 0 Then
            If rsLiquidacion!bitLiquidacion = 1 Then
                strSentencia = " Select Count(*) Co From PvPago " & _
                               " Inner Join NoDepartamento On PvPago.smiDepartamento = NoDepartamento.smiCveDepartamento And NoDepartamento.tnyClaveEmpresa = " & CStr(cboEmpresas.ItemData(cboEmpresas.ListIndex)) & _
                               " Inner Join PvConceptoPagoEmpresa On PvPago.intNumConcepto = PvConceptoPagoEmpresa.intNumConcepto And PvConceptoPagoEmpresa.intCveEmpresa = " & CStr(cboEmpresas.ItemData(cboEmpresas.ListIndex)) & _
                               " Where PvPago.intNumConcepto = " & txtClave.Text
                Set rsConceptoLiq = frsRegresaRs(strSentencia, adLockReadOnly, adOpenForwardOnly)
                If rsConceptoLiq.RecordCount > 0 Then
                    If rsConceptoLiq!CO > 0 Then
                        ' El concepto de liquidación ya fue usado y no puede ser modificado.
                        MsgBox SIHOMsg(792), vbCritical, "Mensaje"
                        chkConceptoLiquidacion.Value = 1
                        chkConceptoLiquidacion.SetFocus
                        Exit Sub
                    End If
                End If
            End If
        End If
        grdCuentas.TextMatrix(lngRow, 3) = 0
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":chkConceptoLiquidacion_LostFocus"))
End Sub

Private Sub chkPagoCancelarFactura_Click()
    pHabilita 0, 0, 0, 0, 0, 1, 0
End Sub

Private Sub cmdBack_Click()
  On Error GoTo NotificaError
    
  If rsPvConceptoPago.RecordCount > 0 Then
    If rsPvConceptoPago.BOF Then
      rsPvConceptoPago.MoveFirst
    Else
      rsPvConceptoPago.MovePrevious
      If rsPvConceptoPago.BOF Then rsPvConceptoPago.MoveFirst
    End If
    pMuestraConcepto rsPvConceptoPago!intNumConcepto
    pHabilita 1, 1, 1, 1, 1, 0, 1
  End If
  
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdBack_Click"))
End Sub

Private Sub cmdBorrarCuenta_Click()
On Error GoTo NotificaError

    '¿Está seguro de eliminar los datos?
    If MsgBox(SIHOMsg("6"), (vbYesNo + vbQuestion), "Mensaje") = vbYes Then
        grdCuentas.TextMatrix(lngRow, 1) = 0
        mskCuenta.Mask = ""
        mskCuenta.Text = ""
        mskCuenta.Mask = Trim(grdCuentas.TextMatrix(lngRow, 2))
        lblDescripcionCuenta.Caption = ""
        chkConceptoLiquidacion.Value = False
    End If
   
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdBorrarCuenta_Click"))
End Sub

Private Sub cmdCuentasContables_Click()
On Error GoTo NotificaError

    If Trim(txtDescripcion.Text) <> "" Then
        sstabConcepto.Tab = 2
        lblConcepto.Caption = txtDescripcion.Text
        If cboEmpresas.Enabled Then
            cboEmpresas.SetFocus
        Else
            If mskCuenta.Enabled Then
                mskCuenta.SetFocus
            Else
                cmdRegresar.SetFocus
            End If
        End If
    End If
  
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdCuentascontables_Click"))
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo NotificaError
    Dim rs As New ADODB.Recordset
    Dim vllngConcepto As Long

    If MsgBox(SIHOMsg(6), vbQuestion + vbYesNo, "Mensaje") = vbYes Then
        
        Set rs = frsRegresaRs("SELECT intnumpago FROM PVPAGO WHERE intnumconcepto = " & CStr(rsPvConceptoPago!intNumConcepto))
        If rs.RecordCount > 0 Then
            MsgBox SIHOMsg(257), vbOKOnly + vbCritical, "Mensaje"
            Exit Sub
        End If
        rs.Close
        
        vllngConcepto = CStr(rsPvConceptoPago!intNumConcepto)
        
        EntornoSIHO.ConeccionSIHO.BeginTrans
        pEjecutaSentencia "DELETE FROM PVCONCEPTOPAGOEMPRESA WHERE intnumconcepto = " & vllngConcepto
        pEjecutaSentencia "DELETE FROM PVCONCEPTOPAGO WHERE intnumconcepto = " & vllngConcepto
        EntornoSIHO.ConeccionSIHO.CommitTrans
        
        rsPvConceptoPago.Requery
        
        txtClave.SetFocus
    End If
    
Exit Sub
NotificaError:
    If Err.Number = -2147217900 Then
        MsgBox SIHOMsg(257), vbOKOnly + vbCritical, "Mensaje"
        Unload Me
    End If
End Sub

Private Sub cmdEnd_Click()
  On Error GoTo NotificaError
    
  If rsPvConceptoPago.RecordCount > 0 Then
    rsPvConceptoPago.MoveLast
    pMuestraConcepto rsPvConceptoPago!intNumConcepto
    pHabilita 1, 1, 1, 1, 1, 0, 1
  End If
  
Exit Sub
NotificaError:
  Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdEnd_Click"))
End Sub

Private Sub cmdLocate_Click()
    On Error GoTo NotificaError
    
    sstabConcepto.Tab = 1

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdLocate_Click"))
End Sub

Private Sub cmdNext_Click()
  On Error GoTo NotificaError
    
  If rsPvConceptoPago.RecordCount > 0 Then
    If rsPvConceptoPago.EOF Then
      rsPvConceptoPago.MoveLast
    Else
      rsPvConceptoPago.MoveNext
      If rsPvConceptoPago.EOF Then rsPvConceptoPago.MoveLast
    End If
    pMuestraConcepto rsPvConceptoPago!intNumConcepto
    pHabilita 1, 1, 1, 1, 1, 0, 1
  End If
  
Exit Sub
NotificaError:
  Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdNext_Click"))
End Sub

Private Sub cmdRegresar_Click()
    sstabConcepto.Tab = 0
    pHabilita 0, 0, 0, 0, 0, 1, 0
    cmdSave.SetFocus
End Sub

Private Sub cmdSave_Click()
    On Error GoTo NotificaError
    Dim vllngPersonaGraba As Long
    Dim lngConcepto As Long
    
    If fblnDatosValidos() Then
        
        vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
        If vllngPersonaGraba <> 0 Then
            
            With rsPvConceptoPago
                If Not vlblnConsulta Then
                    .AddNew
                End If
                !chrDescripcion = Trim(txtDescripcion.Text)
                If cboTipoConcepto.ListIndex = 0 Then 'Normal
                    !chrTipo = "NO"
                ElseIf cboTipoConcepto.ListIndex = 1 Then 'Deducible
                    !chrTipo = "DE"
                ElseIf cboTipoConcepto.ListIndex = 2 Then 'Coaseguro
                    !chrTipo = "CO"
                ElseIf cboTipoConcepto.ListIndex = 3 Then 'Coaseguro adicional
                    !chrTipo = "CA"
                ElseIf cboTipoConcepto.ListIndex = 4 Then 'Copago
                    !chrTipo = "CP"
                ElseIf cboTipoConcepto.ListIndex = 5 Then 'Salidas de dinero
                    !chrTipo = "SD"
                End If
                !bitestatusactivo = chkActivo.Value
                !bitdesglosaiva = IIf(OptDesglosarIVAoExento(0).Value, 1, 0)
                !bitExentoIva = IIf(OptDesglosarIVAoExento(1).Value, 1, 0)
                !bitpagocancelafactura = chkPagoCancelarFactura.Value
                !bitGenerarCFDI = chkcfdi.Value
                                                
                .Update
                
                If Not vlblnConsulta Then
                    lngConcepto = flngObtieneIdentity("SEC_PVCONCEPTOPAGO", rsPvConceptoPago!intNumConcepto)
                Else
                    lngConcepto = txtClave.Text
                End If
                             
                pEjecutaSentencia ("Delete From PvConceptoPagoEmpresa Where intNumConcepto = " & CStr(lngConcepto))
                Set rsCuentasConceptoES = frsRegresaRs("Select * From PvConceptoPagoEmpresa Where intNumConcepto = -1 ", adLockOptimistic, adOpenDynamic)
                For lngContador = 0 To grdCuentas.Rows - 1
                    If grdCuentas.TextMatrix(lngContador, 1) > "0" Then
                        rsCuentasConceptoES.AddNew
                        rsCuentasConceptoES!intNumConcepto = lngConcepto
                        rsCuentasConceptoES!intcveempresa = CLng(grdCuentas.TextMatrix(lngContador, 0))
                        rsCuentasConceptoES!intNumeroCuenta = CLng(grdCuentas.TextMatrix(lngContador, 1))
                        rsCuentasConceptoES!bitConceptoLiquidacion = CLng(grdCuentas.TextMatrix(lngContador, 3))
                        rsCuentasConceptoES.Update
                    End If
                Next lngContador
                rsCuentasConceptoES.Close
                
                If Not vlblnConsulta Then
                    txtClave.Text = lngConcepto
                    Call pGuardarLogTransaccion(Me.Name, EnmGrabar, vglngNumeroLogin, "CONCEPTOPAGO", txtClave.Text)
                Else
                    Call pGuardarLogTransaccion(Me.Name, EnmCambiar, vglngNumeroLogin, "CONCEPTOPAGO", txtClave.Text)
                End If
                
            End With
            rsPvConceptoPago.Requery
            txtClave.SetFocus
        End If
    End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdSave_Click"))
End Sub

Private Function fblnDatosValidos() As Boolean
    Dim lngContador As Long

    On Error GoTo NotificaError
    
    fblnDatosValidos = True
    
    If Trim(txtDescripcion.Text) = "" Then
        fblnDatosValidos = False
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbOKOnly + vbInformation, "Mensaje"
        txtDescripcion.SetFocus
    End If
    If fblnDatosValidos And Not vlblnConsulta Then
        If fblnExisteConcepto(txtDescripcion.Text, "txtDescripcion") Then
            fblnDatosValidos = False
            'Este concepto ya está registrado.
            MsgBox SIHOMsg(319), vbOKOnly + vbInformation, "Mensaje"
            txtDescripcion.SetFocus
        End If
    End If
    If fblnDatosValidos Then
        If chkPagoCancelarFactura.Value = 1 Then
            If fblnExisteConcepto(chkPagoCancelarFactura.Value, "chkPagoCancelarFactura") Then
                fblnDatosValidos = False
                'Ya existe registrado un concepto para generar pagos al cancelar facturas.
                MsgBox SIHOMsg(1324), vbOKOnly + vbInformation, "Mensaje"
                chkPagoCancelarFactura.SetFocus
            End If
        End If
    End If
    
    If fblnDatosValidos Then
        For lngContador = 0 To grdCuentas.Rows - 1
            If grdCuentas.TextMatrix(lngContador, 1) = "0" Or grdCuentas.TextMatrix(lngContador, 1) = "" Then
                fblnDatosValidos = False
                'Seleccione la cuenta contable.
                MsgBox SIHOMsg(211), vbOKOnly + vbInformation, "Mensaje"
                If cmdCuentascontables.Enabled Then
                    cmdCuentascontables.SetFocus
                End If
                
                Exit For
            End If
        Next lngContador
    End If

Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fblnDatosValidos"))
End Function

Private Function fblnExisteConcepto(vlstrxValor As String, vlstrCampo As String) As Boolean
    On Error GoTo NotificaError
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    
    fblnExisteConcepto = False
    If vlstrCampo = "chkPagoCancelarFactura" Then
        strSQL = "select count(*) conceptobitpagocancelafactura from pvconceptopago where pvconceptopago.bitpagocancelafactura = 1 and intnumconcepto <> " & CLng(txtClave.Text)
        Set rs = frsRegresaRs(strSQL, adLockOptimistic, adOpenDynamic)
        If rs.RecordCount > 0 Then
            If rs!conceptobitpagocancelafactura > 0 Then
                fblnExisteConcepto = True
            End If
        End If
    Else
        If rsPvConceptoPago.RecordCount <> 0 Then
            rsPvConceptoPago.MoveFirst
            Do While Not rsPvConceptoPago.EOF
                If Trim(rsPvConceptoPago!chrDescripcion) = vlstrxValor Then
                    fblnExisteConcepto = True
                End If
                rsPvConceptoPago.MoveNext
            Loop
        End If
    End If
Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fblnExisteConcepto"))
End Function

Private Sub cmdTop_Click()
    On Error GoTo NotificaError
    
    If rsPvConceptoPago.RecordCount > 0 Then
        rsPvConceptoPago.MoveFirst
        pMuestraConcepto rsPvConceptoPago!intNumConcepto
        pHabilita 1, 1, 1, 1, 1, 0, 1
    End If
  
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdTop_Click"))
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 27 Then
        If sstabConcepto.Tab = 0 Then
            If Not vlblnConsulta Then
                Unload Me
            Else
                '¿Desea abandonar la operación?
                If MsgBox(SIHOMsg(17), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
                    pLimpia
                    txtClave.SetFocus
                End If
            End If
        Else
            sstabConcepto.Tab = 0
            txtClave.SetFocus
        End If
    Else
        If KeyAscii = 13 Then
            SendKeys vbTab
        End If
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_KeyPress"))
End Sub

Private Sub Form_Load()
    On Error GoTo NotificaError
    Dim rsEmpresas As New ADODB.Recordset
    
    Me.Icon = frmMenuPrincipal.Icon
    
    vlintbitPagoCancelaFactura = fintBitCuentaPuenteBanco(vgintClaveEmpresaContable)
    
    vlstrx = "select * from PvConceptoPago"
    Set rsPvConceptoPago = frsRegresaRs(vlstrx, adLockOptimistic, adOpenDynamic)
    
    vlstrSentenciaConsulta = "select intNumConcepto,chrDescripcion,case when bitEstatusActivo = 1 then 'Activo' when bitEstatusActivo = 0 then 'Inactivo' end as Estatus from PvConceptoPago"
    
    Set rsEmpresas = frsRegresaRs("Select tnyClaveEmpresa, Trim(vchNombreCorto) From CnEmpresaContable Where bitActiva = 1 ")
    If rsEmpresas.RecordCount > 0 Then
        pLlenarCboRs cboEmpresas, rsEmpresas, 0, 1
        rsEmpresas.Close
    End If
    cboEmpresas.ListIndex = fintLocalizaCbo(cboEmpresas, CStr(vgintClaveEmpresaContable))
    
    'Tipos de conceptos
    cboTipoConcepto.Clear
    cboTipoConcepto.AddItem "CONCEPTO NORMAL", 0
    cboTipoConcepto.AddItem "CONCEPTO PARA PAGOS DE DEDUCIBLE", 1
    cboTipoConcepto.AddItem "CONCEPTO PARA PAGOS DE COASEGURO", 2
    cboTipoConcepto.AddItem "CONCEPTO PARA PAGOS DE COASEGURO ADICIONAL", 3
    cboTipoConcepto.AddItem "CONCEPTO PARA PAGOS DE COPAGO", 4
    cboTipoConcepto.AddItem "CONCEPTO PARA SALIDAS DE DINERO", 5
    cboTipoConcepto.ListIndex = 0
    
    mskCuenta.Enabled = False
    If fblnRevisaPermiso(vglngNumeroLogin, IIf(cgstrModulo = "PV", 316, 1122), "C", True) Then
        cboEmpresas.Enabled = True
        mskCuenta.Enabled = fConceptoCuentaModificable
        cmdBorrarCuenta.Enabled = True
    End If
       
    sstabConcepto.Tab = 0
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_Load"))
End Sub

Private Function fintBitCuentaPuenteBanco(intEmpresa As Integer) As Integer
On Error GoTo NotificaError

    Dim rs As ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "select bitUtilizaCuentaPuenteBanco from pvParametro where tnyclaveempresa = " & intEmpresa
    Set rs = frsRegresaRs(strSQL, adLockOptimistic, adOpenDynamic)

    If Not rs.EOF Then
        If IsNull(rs!bitUtilizaCuentaPuenteBanco) Then
            fintBitCuentaPuenteBanco = -1
        Else
            fintBitCuentaPuenteBanco = Val(rs!bitUtilizaCuentaPuenteBanco)
        End If
    Else
        fintBitCuentaPuenteBanco = -1
    End If
    rs.Close
    
Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fintBitCuentaPuenteBanco"))
End Function

Private Sub grdConceptos_DblClick()
    On Error GoTo NotificaError
    
    pMuestraConcepto grdConceptos.RowData(grdConceptos.Row)
    pHabilita 1, 1, 1, 1, 1, 0, 1
    sstabConcepto.Tab = 0
    cmdLocate.SetFocus

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":grdConceptos_DblClick"))
End Sub

Private Sub grdConceptos_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    If KeyCode = vbKeyReturn Then
        grdConceptos_DblClick
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":grdConceptos_KeyDown"))
End Sub

Private Sub mskCuenta_GotFocus()
    pSelMkTexto mskCuenta
End Sub

Private Sub mskCuenta_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 13 Then pAsignaCuenta mskCuenta, lblDescripcionCuenta

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskCuenta_KeyPress"))
End Sub

Private Sub OptDesglosarIVAoExento_GotFocus(Index As Integer)
    On Error GoTo NotificaError
    pHabilita 0, 0, 0, 0, 0, 1, 0
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":OptDesglosarIVAoExento_GotFocus"))
End Sub

Private Sub sstabConcepto_Click(PreviousTab As Integer)
    On Error GoTo NotificaError
    
    If sstabConcepto.Tab = 1 Then
        grdConceptos.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":sstabConcepto_Click"))
End Sub

Private Sub txtClave_GotFocus()
    On Error GoTo NotificaError
    
    If rsPvConceptoPago.RecordCount > 0 Then
        pHabilita 1, 1, 1, 1, 1, 0, 0
    Else
        pHabilita 0, 0, 0, 0, 0, 0, 0
    End If
    pLimpia
    pSelTextBox txtClave

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtClave_GotFocus"))
End Sub

Private Sub pLimpia()
    On Error GoTo NotificaError
    
    If rsPvConceptoPago.RecordCount = 0 Then
        txtClave.Text = "1"
    Else
        txtClave.Text = frsRegresaRs("select max(intNumConcepto)+1 from PvConceptoPago").Fields(0)
    End If
    
    txtDescripcion.Text = ""
    cboTipoConcepto.ListIndex = 0
    chkActivo.Value = 1
    OptDesglosarIVAoExento(0).Value = 1
    OptDesglosarIVAoExento(1).Value = 0
    OptDesglosarIVAoExento(2).Value = 0
    chkPagoCancelarFactura.Value = 0
    chkcfdi.Value = 0
    
    If rsPvConceptoPago.RecordCount <> 0 Then
        pLlenarMshFGrdRs grdConceptos, frsRegresaRs(vlstrSentenciaConsulta), 0
        With grdConceptos
            .FormatString = "|Clave|Descripción|Estado"
            .ColWidth(0) = 100
            .ColWidth(1) = 1000
            .ColWidth(2) = 5000
            .ColWidth(3) = 1500
        End With
    End If
    
    cboEmpresas.ListIndex = fintLocalizaCbo(cboEmpresas, CStr(vgintClaveEmpresaContable))
    
    pLlenaGrid "-1"
    
    vlblnConsulta = False
    
    mskCuenta.Enabled = False
    If fblnRevisaPermiso(vglngNumeroLogin, IIf(cgstrModulo = "PV", 316, 1122), "C", True) Then
        mskCuenta.Enabled = fConceptoCuentaModificable
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pLimpia"))
End Sub

Private Sub pHabilita(vlb1 As Integer, vlb2 As Integer, vlb3 As Integer, vlb4 As Integer, vlb5 As Integer, vlb6 As Integer, vlb7 As Integer)
    On Error GoTo NotificaError
    
    If vlb1 = 1 Then
        cmdTop.Enabled = True
    Else
        cmdTop.Enabled = False
    End If
    If vlb2 = 1 Then
        cmdBack.Enabled = True
    Else
        cmdBack.Enabled = False
    End If
    If vlb3 = 1 Then
        cmdLocate.Enabled = True
    Else
        cmdLocate.Enabled = False
    End If
    If vlb4 = 1 Then
        cmdNext.Enabled = True
    Else
        cmdNext.Enabled = False
    End If
    If vlb5 = 1 Then
        cmdEnd.Enabled = True
    Else
        cmdEnd.Enabled = False
    End If
    If vlb6 = 1 Then
        cmdSave.Enabled = True
    Else
        cmdSave.Enabled = False
    End If
    If vlb7 = 1 Then
        cmdDelete.Enabled = True
    Else
        cmdDelete.Enabled = False
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pHabilita"))
End Sub

Private Sub txtClave_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
        KeyAscii = 7
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtClave_KeyPress"))
End Sub

Private Sub txtClave_LostFocus()
    On Error GoTo NotificaError
    
    If Trim(txtClave.Text) = "" Then
        pLimpia
    Else
        pMuestraConcepto Val(txtClave.Text)
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtClave_LostFocus"))
End Sub

Private Sub pMuestraConcepto(vllngxNumero As Long)
    On Error GoTo NotificaError
    
    If fintLocalizaPkRs(rsPvConceptoPago, 0, Str(vllngxNumero)) <> 0 Then
        txtClave.Text = rsPvConceptoPago!intNumConcepto
        txtDescripcion.Text = rsPvConceptoPago!chrDescripcion
        
        If rsPvConceptoPago!chrTipo = "NO" Then 'Normal
            cboTipoConcepto.ListIndex = 0
        ElseIf rsPvConceptoPago!chrTipo = "DE" Then 'Deducible
            cboTipoConcepto.ListIndex = 1
        ElseIf rsPvConceptoPago!chrTipo = "CO" Then 'Coaseguro
            cboTipoConcepto.ListIndex = 2
        ElseIf rsPvConceptoPago!chrTipo = "CA" Then 'Coaseguro adicional
            cboTipoConcepto.ListIndex = 3
        ElseIf rsPvConceptoPago!chrTipo = "CP" Then 'Copago
            cboTipoConcepto.ListIndex = 4
        ElseIf rsPvConceptoPago!chrTipo = "SD" Then 'Salidas de dinero
            cboTipoConcepto.ListIndex = 5
        End If
        
        chkActivo.Value = IIf(rsPvConceptoPago!bitestatusactivo Or rsPvConceptoPago!bitestatusactivo = 1, 1, 0)
        
        OptDesglosarIVAoExento(0).Value = IIf(rsPvConceptoPago!bitdesglosaiva Or rsPvConceptoPago!bitdesglosaiva = 1, 1, 0)
        OptDesglosarIVAoExento(1).Value = IIf(rsPvConceptoPago!bitExentoIva Or rsPvConceptoPago!bitExentoIva = 1, 1, 0)
        
        If rsPvConceptoPago!bitdesglosaiva = 0 And rsPvConceptoPago!bitExentoIva = 0 Then
            OptDesglosarIVAoExento(2).Value = 1
        End If
        
        chkPagoCancelarFactura.Value = rsPvConceptoPago!bitpagocancelafactura
        chkcfdi.Value = rsPvConceptoPago!bitGenerarCFDI
        
        pLlenaGrid CStr(rsPvConceptoPago!intNumConcepto)
        
        cmdCuentascontables.SetFocus
        
        mskCuenta.Enabled = False
        If fblnRevisaPermiso(vglngNumeroLogin, IIf(cgstrModulo = "PV", 316, 1122), "C", True) Then
            mskCuenta.Enabled = fConceptoCuentaModificable
        End If
        
        vlblnConsulta = True
    Else
        pLimpia
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pMuestraConcepto"))
End Sub

Private Sub txtDescripcion_GotFocus()
    On Error GoTo NotificaError
    
    pHabilita 0, 0, 0, 0, 0, 1, 0
    pSelTextBox txtDescripcion

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtDescripcion_GotFocus"))
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    KeyAscii = Asc(UCase(Chr(KeyAscii)))

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtDescripcion_KeyPress"))
End Sub

Private Sub pAsignaCuenta(mskObject As MaskEdBox, lblObject As Label)
    On Error GoTo NotificaError
    
    Dim vllngNumeroCuenta As Long
    Dim vlstrCuentaCompleta As String
    Dim rs As New ADODB.Recordset

    If Trim(mskObject.ClipText) = "" Then
        vllngNumeroCuenta = flngBusquedaCuentasContables(False, cboEmpresas.ItemData(cboEmpresas.ListIndex))
    Else
        vlstrCuentaCompleta = fstrCuentaCompleta(mskObject.Text)
        vllngNumeroCuenta = flngNumeroCuenta(vlstrCuentaCompleta, cboEmpresas.ItemData(cboEmpresas.ListIndex))
    End If
    
    If vllngNumeroCuenta <> 0 Then
    
        Set rs = frsRegresaRs("Select vchClasificacionTipo, vchSubclasificacionTipo From Cncuenta Where intNumeroCuenta = " & CStr(vllngNumeroCuenta) & " And tnyClaveEmpresa = " & CStr(cboEmpresas.ItemData(cboEmpresas.ListIndex)))
        If rs.RecordCount > 0 Then
'            If rs!vchClasificacionTipo = "Pasivo" And rs!vchSubclasificacionTipo = "Circulante" Then
                mskObject.Text = fstrCuentaContable(vllngNumeroCuenta)
                grdCuentas.TextMatrix(lngRow, 1) = vllngNumeroCuenta
'            Else
'                'Seleccione una cuenta de Pasivo circulante
'                MsgBox SIHOMsg(906), vbOKOnly + vbExclamation, "Mensaje"
'                mskObject.Mask = ""
'                mskObject.Text = ""
'                mskObject.Mask = Trim(grdCuentas.TextMatrix(lngRow, 2))
'                lblDescripcionCuenta.Caption = ""
'                Exit Sub
'            End If
        End If
        lblDescripcionCuenta.Caption = fstrDescripcionCuenta(mskObject.Text, cboEmpresas.ItemData(cboEmpresas.ListIndex))
    
    Else
        'No se encontró la cuenta contable.
        MsgBox SIHOMsg(222), vbOKOnly + vbExclamation, "Mensaje"
        mskObject.Mask = ""
        mskObject.Text = ""
        mskObject.Mask = Trim(grdCuentas.TextMatrix(lngRow, 2))
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pAsignaCuenta"))
    Unload Me
End Sub

Private Function fstrCuentaCompleta(vlstrCuenta As String, Optional vllngDigitos As Long) As String
    On Error GoTo NotificaError
    Dim vllngContador As Integer
    Dim vllngTotalDigitos As Long
    
    fstrCuentaCompleta = ""
    
    If vllngDigitos = 0 Then
        vllngTotalDigitos = Len(Trim(grdCuentas.TextMatrix(lngRow, 2)))
    Else
        vllngTotalDigitos = vllngDigitos
    End If
    
    vllngContador = 1
    Do While vllngContador <= vllngTotalDigitos
        If Mid(vgstrEstructuraCuentaContable, vllngContador, 1) = "#" Then
            If Trim(Mid(vlstrCuenta, vllngContador, 1)) <> "" Then
                fstrCuentaCompleta = fstrCuentaCompleta + Mid(vlstrCuenta, vllngContador, 1)
            Else
                fstrCuentaCompleta = fstrCuentaCompleta + "0"
            End If
        Else
            fstrCuentaCompleta = fstrCuentaCompleta + "."
        End If
        vllngContador = vllngContador + 1
    Loop
    
Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fstrCuentaCompleta"))
End Function

Private Sub pLlenaGrid(strCveConcepto As String)
On Error GoTo NotificaError

    grdCuentas.Clear
    grdCuentas.Rows = 0
    
    Set rsCuentasConceptoES = frsEjecuta_SP(strCveConcepto, "sp_PvSelCuentasConceptoPago")
    
    If rsCuentasConceptoES.RecordCount > 0 Then
        For lngContador = 0 To rsCuentasConceptoES.RecordCount - 1
            grdCuentas.Rows = lngContador + 1
            grdCuentas.TextMatrix(lngContador, 0) = CStr(rsCuentasConceptoES!empresa)
            grdCuentas.TextMatrix(lngContador, 1) = CStr(rsCuentasConceptoES!NumeroCuenta)
            grdCuentas.TextMatrix(lngContador, 2) = Trim(rsCuentasConceptoES!Estructura)
            grdCuentas.TextMatrix(lngContador, 3) = CStr(rsCuentasConceptoES!Liquidacion)
            
            If rsCuentasConceptoES!empresa = cboEmpresas.ItemData(cboEmpresas.ListIndex) Then
                mskCuenta.Mask = ""
                mskCuenta.Text = IIf(rsCuentasConceptoES!NumeroCuenta > 0, fstrCuentaContable(CStr(rsCuentasConceptoES!NumeroCuenta)), "")
                mskCuenta.Mask = Trim(rsCuentasConceptoES!Estructura)
            
                lblDescripcionCuenta.Caption = IIf(rsCuentasConceptoES!NumeroCuenta > 0, fstrDescripcionCuenta(mskCuenta.Text, rsCuentasConceptoES!empresa), "")
                             
                lngRow = lngContador
                
                chkConceptoLiquidacion.Value = rsCuentasConceptoES!Liquidacion
                
            End If
            rsCuentasConceptoES.MoveNext
        Next lngContador
        rsCuentasConceptoES.Close
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pLlenaGrid"))
End Sub

Private Function fConceptoCuentaModificable() As Boolean
On Error GoTo NotificaError
    Dim rs As New ADODB.Recordset

    fConceptoCuentaModificable = True
    
    If txtClave.Text = "" Then
        fConceptoCuentaModificable = True
        Exit Function
    End If
    
    Set rs = frsRegresaRs("SELECT intnumconcepto FROM PVPAGO WHERE intnumconcepto = " & txtClave.Text & " and smidepartamento in (select smicvedepartamento from nodepartamento where tnyclaveempresa = " & CStr(cboEmpresas.ItemData(cboEmpresas.ListIndex)) & ")")
    If rs.RecordCount > 0 Then
        fConceptoCuentaModificable = False
    End If
    rs.Close
    
    Set rs = frsRegresaRs("SELECT intnumconcepto FROM PVSALIDADINERO WHERE intnumconcepto = " & txtClave.Text & " and smidepartamento in (select smicvedepartamento from nodepartamento where tnyclaveempresa = " & CStr(cboEmpresas.ItemData(cboEmpresas.ListIndex)) & ")")
    If rs.RecordCount > 0 Then
        fConceptoCuentaModificable = False
    End If
    rs.Close
        
Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fConceptoCuentaModificable"))
End Function
