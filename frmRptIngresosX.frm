VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRptIngresosX 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingresos diarios"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   525
   ClientWidth     =   7920
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   7920
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      Height          =   825
      Left            =   3217
      TabIndex        =   13
      Top             =   2595
      Width           =   1605
      Begin VB.CommandButton cmdExportar 
         Caption         =   "Exportar"
         Height          =   500
         Left            =   100
         TabIndex        =   17
         ToolTipText     =   "Exportar a documento de Excel"
         Top             =   200
         Width           =   1400
      End
   End
   Begin VB.Frame freBarra 
      Height          =   860
      Left            =   113
      TabIndex        =   9
      Top             =   1080
      Visible         =   0   'False
      Width           =   7695
      Begin MSComctlLib.ProgressBar pgbBarra 
         Height          =   375
         Left            =   60
         TabIndex        =   11
         Top             =   420
         Width           =   7595
         _ExtentX        =   13388
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblTextoBarra 
         BackColor       =   &H80000002&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   50
         TabIndex        =   10
         Top             =   120
         Width           =   7605
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid gridFinal 
      Height          =   3375
      Left            =   120
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   8640
      Width           =   16575
      _cx             =   29236
      _cy             =   5953
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
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
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   0
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
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VSFlex7LCtl.VSFlexGrid grdIngresosDiarios 
      Height          =   2655
      Left            =   120
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3840
      Width           =   16575
      _cx             =   29236
      _cy             =   4683
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
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
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmRptIngresosX.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   3
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
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Frame Frame1 
      Height          =   2325
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   7695
      Begin VB.Frame frmFechas 
         Caption         =   "Rango de fechas de búsqueda"
         Height          =   760
         Left            =   1800
         TabIndex        =   14
         Top             =   1320
         Width           =   4215
         Begin MSMask.MaskEdBox txtFechaInicio 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "d/MMM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
            Height          =   315
            Left            =   720
            TabIndex        =   2
            ToolTipText     =   "Ingresar fecha inicial de reporte"
            Top             =   285
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            ClipMode        =   1
            MaxLength       =   10
            Format          =   "dd/mmm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtFechaFin 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "d/MMM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
            Height          =   315
            Left            =   2760
            TabIndex        =   3
            ToolTipText     =   "Ingresar fecha final de reporte"
            Top             =   285
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            ClipMode        =   1
            MaxLength       =   10
            Format          =   "dd/mmm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            Height          =   195
            Left            =   2160
            TabIndex        =   16
            Top             =   345
            Width           =   420
         End
         Begin VB.Label lblNumeroCuenta 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   345
            Width           =   465
         End
      End
      Begin VB.ComboBox cboDepartamento 
         Height          =   315
         Left            =   3000
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Selección del departamento"
         Top             =   720
         Width           =   4500
      End
      Begin VB.ComboBox cboHospital 
         Height          =   315
         Left            =   3000
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Selección de empresa contable"
         Top             =   300
         Width           =   4500
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Departamento que facturó"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   780
         Width           =   1860
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Empresa contable"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1515
         WordWrap        =   -1  'True
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grdFormasPago 
      Height          =   1935
      Left            =   120
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   6600
      Width           =   16575
      _cx             =   29236
      _cy             =   3413
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
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
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
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
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
End
Attribute VB_Name = "frmRptIngresosX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Cliente|Fecha|RFC|UUID|Tipo|Folio|Departamento|Importe|Descuento|IVA Cobrado|IVA no cobrado|Total|Ingreso cobrado gravado|Ingreso cobrado no gravado|Ingreso no cobrado gravado|Ingreso no cobrado no gravado
Dim cintColCliente As Integer
Dim cintColFecha As Integer
Dim cIntColRFC As Integer
Dim cintColUUID As Integer
Dim cintColTipo As Integer
Dim cIntColFolio As Integer
Dim cintColDepartamento As Integer
Dim cintColImporte As Integer
Dim cintColDescuento As Integer
Dim cintColIVACobrado As Integer
Dim cintColIVANoCobrado As Integer
Dim cintColTotal As Integer
Dim cintColICG As Integer
Dim cintColICNG As Integer
Dim cintColINCG As Integer
Dim cintColINCNG As Integer
Dim cintColTipoPaciente As Integer
Dim cintColAnticipo As Integer
Dim cintColObservacion As Integer
Dim cintColDif As Integer
Dim cintColPruebaIVA As Integer

Public vglngNumeroOpcion As Long
Dim vlpermiso As String

Dim cintColClienteI As Integer
Dim cintColFechaI As Integer
Dim cintColRFCI As Integer
Dim cintColUUIDI As Integer
Dim cintColTipoI As Integer
Dim cintColFolioI As Integer
Dim cintColDepartamentoI As Integer
Dim cintColImporteI As Integer
Dim cintColDescuentoI As Integer
Dim cintColIVACobradoI As Integer
Dim cintColIVANoCobradoI As Integer
Dim cintColTotalI As Integer
Dim cintColICGI As Integer
Dim cintColICNGI As Integer
Dim cintColINCGI As Integer
Dim cintColINCNGI As Integer
Dim cintColTipoPacienteI As Integer


Dim cintColFechaFP As Integer
Dim cintColFolioFP As Integer
Dim cintColTipoDocumentoFP As Integer
Dim cintColClaveFormaPagoFP As Integer
Dim cintColDescripcionFormaPagoFP As Integer
Dim cintColReferenciaFP As Integer
Dim cintColCantidadPagadaFP As Integer
Dim cintColMonedaFP As Integer
Dim cintColEmpleadoFP As Integer
Dim cintColEstatusFP As Integer
Dim cintColChrTipoDocFP As Integer
Dim cintColChrFolioReciboFP As Integer

Dim vlblnfechavalida As Boolean

Private Sub cboDepartamento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub cboHospital_Click()
    Dim rs As New ADODB.Recordset
    Dim vgstrParametrosSP As String
    If cboHospital.ListIndex <> -1 Then
       cboDepartamento.Clear
       vgstrParametrosSP = "select smicvedepartamento, vchdescripcion from nodepartamento inner join pvcorte on nodepartamento.smicvedepartamento = pvcorte.smidepartamento where nodepartamento.tnyclaveempresa = '" & cboHospital.ItemData(cboHospital.ListIndex) & "' group by smicvedepartamento, vchdescripcion order by smicvedepartamento"
       Set rs = frsRegresaRs(vgstrParametrosSP)
  
       If rs.RecordCount <> 0 Then
          pLlenarCboRs cboDepartamento, rs, 0, 1
          cboDepartamento.ListIndex = flngLocalizaCbo(cboDepartamento, str(vgintNumeroDepartamento))
          'Caso 20550 se agrega <TODOS>
          cboDepartamento.AddItem "<TODOS>", 0
          cboDepartamento.ItemData(cboDepartamento.newIndex) = -1
          cboDepartamento.ListIndex = 0
       End If
       
       
       
    End If
End Sub

Private Sub cboHospital_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub cmdExportar_Click()
If cboDepartamento.ListCount > 0 And cboDepartamento.Text <> "" Then

    'caso 20550 correccion en la validacion de Fecha GIRM-09/10/2024
     vlblnfechavalida = False

    If IsDate(txtFechaInicio.Text) Then
        Dim fecha As Date
        fecha = txtFechaInicio.Text
        If DateDiff("d", Date, fecha) <= 0 Then
            vlblnfechavalida = True
        
            If KeyCode = vbKeyReturn Then SendKeys vbTab
        Else
            Call MsgBox("¡Fecha invalida!", vbExclamation, "Mensaje")
            pEnfocaMkTexto txtFechaInicio
        End If
    Else
        Call MsgBox("¡Fecha invalida!", vbExclamation, "Mensaje")
        pEnfocaMkTexto txtFechaInicio
    End If
    If IsDate(txtFechaFin.Text) Then
        Dim Fechafinal As Date
        Fechafinal = txtFechaFin.Text
        If DateDiff("d", Date, Fechafinal) <= 0 Then
            vlblnfechavalida = True
        
            If KeyCode = vbKeyReturn Then SendKeys vbTab
        Else
            Call MsgBox("¡Fecha invalida!", vbExclamation, "Mensaje")
            pEnfocaMkTexto txtFechaFin
        End If
    Else
        Call MsgBox("¡Fecha invalida!", vbExclamation, "Mensaje")
        pEnfocaMkTexto txtFechaFin
    End If
    
    If vlblnfechavalida Then
        Screen.MousePointer = 11 'caso 20550
        cmdExportar.Enabled = False 'caso 20550
        realizaExportacion
        Screen.MousePointer = 0 'caso 20550
        cmdExportar.Enabled = True 'caso 20550
    End If
Else
   Call MsgBox("¡No se ha seleccionado ningun departamento!", vbExclamation, "Mensaje")
End If
End Sub

Private Sub realizaExportacion()
    On Error GoTo NotificaError

        Dim rsedo As New ADODB.Recordset
        Dim rsIngresosDiarios As New ADODB.Recordset
        Dim queryFormasPago As String
        Dim rsFormasPago As New ADODB.Recordset
        Dim rsiva As New ADODB.Recordset
        Dim paramsForma As String
        Dim fecha1 As String
        Dim fecha2 As String
        Dim rscantidadformaspago As New ADODB.Recordset
        Dim querycolsformaspago As String
        Dim colsFormaPago As String
        Dim intformaspago As Integer
        Dim queryiva As String
        Dim relporcentajeiva As Double
        queryiva = "select siparametro.vchnombre, siparametro.vchvalor, siparametro.intcveempresacontable, cnimpuesto.relporcentaje relporcentaje from siparametro inner join cnimpuesto on siparametro.vchvalor = to_char(cnimpuesto.SMICVEIMPUESTO) where vchnombre = 'INTTASAIMPUESTOHOSPITAL' and intcveempresacontable = " & cboHospital.ItemData(cboHospital.ListIndex)
        colsFormaPago = "|"
        querycolsformaspago = "select intformapago, chrdescripcion from pvformapago where smidepartamento = '" & cboDepartamento.ItemData(cboDepartamento.ListIndex) & "' and bitestatusactivo = 1 order by intformapago"
       
        '*********************************
        'columnas de grid final
        cintColCliente = 0
        cintColFecha = 1
        cIntColRFC = 2
        cintColUUID = 3
        cintColTipo = 4
        cIntColFolio = 5
        cintColDepartamento = 6
        cintColImporte = 7
        cintColDescuento = 8
        cintColIVACobrado = 9
        cintColIVANoCobrado = 10
        cintColTotal = 11
        cintColICG = 12
        cintColICNG = 13
        cintColINCG = 14
        cintColINCNG = 15
        cintColTipoPaciente = 17
        cintColAnticipo = 16
        cintColObservacion = 18
        cintColDif = 19
        cintColPruebaIVA = 20
        
        'Grid ingresos diarios
        cintColClienteI = 0
        cintColFechaI = 1
        cintColRFCI = 2
        cintColUUIDI = 3
        cintColTipoI = 4
        cintColFolioI = 5
        cintColDepartamentoI = 6
        cintColImporteI = 7
        cintColDescuentoI = 8
        cintColIVACobradoI = 9
        cintColIVANoCobradoI = 10
        cintColTotalI = 11
        cintColICGI = 12
        cintColICNGI = 13
        cintColINCGI = 14
        cintColINCNGI = 15
        cintColTipoPacienteI = 16
        
        'Grid formas de pago
        cintColFechaFP = 0
        cintColFolioFP = 1
        cintColTipoDocumentoFP = 2
        cintColClaveFormaPagoFP = 3
        cintColDescripcionFormaPagoFP = 4
        cintColReferenciaFP = 5
        cintColCantidadPagadaFP = 6
        cintColMonedaFP = 7
        cintColEmpleadoFP = 8
        cintColEstatusFP = 9
        cintColChrTipoDocFP = 10
        cintColChrFolioReciboFP = 11

        '*********************************
        
        lblTextoBarra.Caption = "Exportando información, por favor espere..."
        freBarra.Visible = True
        freBarra.Top = 720
        pgbBarra.Value = 10
        freBarra.Refresh
        
        
        fecha1 = fstrFechaSQL(txtFechaInicio.Text)
        fecha2 = fstrFechaSQL(txtFechaFin.Text)
        'formato de llamada SP_FORMASPAGOINS(0, 1, IN_chrFechaIni, IN_chrFechaIni, IN_intDepto, 1, 1, 1, 1, 1, 1, 1, 1, 0, 0, 0, -1, 0);
        paramsForma = 0 & _
                       "|" & 1 & _
                       "|" & fecha1 & _
                       "|" & fecha2 & _
                       "|" & cboDepartamento.ItemData(cboDepartamento.ListIndex) & _
                       "|" & 1 & _
                       "|" & 1 & _
                       "|" & 1 & _
                       "|" & 1 & _
                       "|" & 1 & _
                       "|" & 1 & _
                       "|" & 1 & _
                       "|" & 1 & _
                       "|" & 0 & _
                       "|" & 0 & _
                       "|" & 0 & _
                       "|" & -1 & _
                       "|" & 0
        'formato de llamada SP_PVINGRESOSDIARIOS('15/06/2022', '15/06/2022', 1, 1, 1, 1, 1, 0, 2, 29, -1, -1, 1, 0, :rc1);
        vgstrParametrosSP = "'" & txtFechaInicio.Text & "'" & _
                       "|" & "'" & txtFechaFin.Text & "'" & _
                       "|" & 1 & _
                       "|" & 1 & _
                       "|" & 1 & _
                       "|" & 1 & _
                       "|" & 1 & _
                       "|" & 0 & _
                       "|" & 2 & _
                       "|" & cboDepartamento.ItemData(cboDepartamento.ListIndex) & _
                       "|" & -1 & _
                       "|" & -1 & _
                       "|" & cboHospital.ItemData(cboHospital.ListIndex) & _
                       "|" & 0
                       
        frsEjecuta_SP paramsForma, "SP_PVRPTFORMASPAGOINS"
        
        
        Set rsIngresosDiarios = frsEjecuta_SP(vgstrParametrosSP, "SP_PVRPTINGRESOSDIARIOS")
        
        queryFormasPago = "SELECT VCHFECHAMOVIMIENTO, VCHFOLIODOCUMENTO, VCHTIPODOCUMENTO, VCHCLAVEFORMAPAGO, VCHDESCRIPCIONFORMAPAGO, VCHREFERENCIA, NUMCANTIDADPAGADA, VCHMONEDA, VCHEMPLEADO, VCHESTATUS, CNTMPCORTEFORMASPAGO.VCHCHRTIPODOCUMENTO, AN.CHRFOLIORECIBO FROM CNTMPCORTEFORMASPAGO " & _
                            "LEFT JOIN PVPAGO AN ON TRIM(CNTMPCORTEFORMASPAGO.VCHFOLIODOCUMENTO) = TRIM(AN.CHRFOLIOFACTURA) AND CNTMPCORTEFORMASPAGO.NUMCANTIDADPAGADA = AN.MNYCANTIDAD " & _
                            "Group By VCHFECHAMOVIMIENTO, VCHFOLIODOCUMENTO, VCHTIPODOCUMENTO, VCHCLAVEFORMAPAGO, VCHDESCRIPCIONFORMAPAGO, VCHREFERENCIA, NUMCANTIDADPAGADA, VCHMONEDA, VCHEMPLEADO, VCHESTATUS, CNTMPCORTEFORMASPAGO.VCHCHRTIPODOCUMENTO, AN.CHRFOLIORECIBO Order By VCHTIPODOCUMENTO, VCHFOLIODOCUMENTO, AN.CHRFOLIORECIBO"

        Set rsiva = frsRegresaRs(queryiva)
        Set rsFormasPago = frsRegresaRs(queryFormasPago)
        If rsiva.RecordCount > 0 Then
            relporcentajeiva = rsiva!relPorcentaje / 100
        End If
    
        If rsIngresosDiarios.RecordCount <> 0 Then
            pLlenaVsfGrid grdIngresosDiarios, rsIngresosDiarios, True, True, True
            If rsFormasPago.RecordCount <> 0 Then
                pLlenaVsfGrid grdFormasPago, rsFormasPago, True, True, True
            Else
                MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje"
                freBarra.Visible = False
                freBarra.Top = 720
                pgbBarra.Value = 0
                freBarra.Refresh
                pgbBarra.Value = 0
                Exit Sub
            End If
        Else
            freBarra.Visible = False
            freBarra.Top = 720
            pgbBarra.Value = 0
            freBarra.Refresh
            pgbBarra.Value = 0
            MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje"
            Exit Sub
        End If
        rsIngresosDiarios.Close
        rsFormasPago.Close
        
        Set rscantidadformaspago = frsRegresaRs(querycolsformaspago)
        If rscantidadformaspago.RecordCount <> 0 Then
            Do While Not rscantidadformaspago.EOF
                Dim valFormat As String
                valFormat = Format(rscantidadformaspago!intFormaPago, "0##")
                colsFormaPago = colsFormaPago & valFormat & " " & Trim(rscantidadformaspago!chrDescripcion) & "|"
                rscantidadformaspago.MoveNext
                
            Loop
            gridFinal.Cols = 21 + rscantidadformaspago.RecordCount
            intformaspago = rscantidadformaspago.RecordCount
        End If
        rscantidadformaspago.Close
        
        With gridFinal
            .FormatString = "Cliente|Fecha|RFC|UUID|Tipo|Folio|Departamento|Importe|Descuento|IVA cobrado|IVA no cobrado|Total|Ingreso cobrado gravado|Ingreso cobrado no gravado|Ingreso no cobrado gravado|Ingreso no cobrado no gravado" & colsFormaPago & "Anticipo facturado|Tipo paciente|Observación|Diferencia|Prueba IVA"
            .ColWidth(cIntColFolio) = 1300
            .ColWidth(cintColFecha) = 3000
            .ColWidth(cintColAnticipo + intformaspago) = 1600
            
            
        End With
        pgbBarra.Value = 25
        freBarra.Refresh
        
        Dim rowsIngresos As Long
        Dim i As Long
        rowsIngresos = grdIngresosDiarios.Rows - 1
        
        For i = 1 To rowsIngresos
            If i = 1 Then
                gridFinal.TextMatrix(i, cintColCliente) = grdIngresosDiarios.TextMatrix(i, cintColClienteI)
                gridFinal.TextMatrix(i, cintColFecha) = Format(grdIngresosDiarios.TextMatrix(i, cintColFechaI), "dd/MM/YYYY")
                gridFinal.TextMatrix(i, cIntColRFC) = grdIngresosDiarios.TextMatrix(i, cintColRFCI)
                gridFinal.TextMatrix(i, cintColUUID) = grdIngresosDiarios.TextMatrix(i, cintColUUIDI)
                gridFinal.TextMatrix(i, cintColTipo) = grdIngresosDiarios.TextMatrix(i, cintColTipoI)
                gridFinal.TextMatrix(i, cIntColFolio) = grdIngresosDiarios.TextMatrix(i, cintColFolioI)
                gridFinal.TextMatrix(i, cintColDepartamento) = grdIngresosDiarios.TextMatrix(i, cintColDepartamentoI)
                gridFinal.TextMatrix(i, cintColImporte) = grdIngresosDiarios.TextMatrix(i, cintColImporteI)
                gridFinal.TextMatrix(i, cintColDescuento) = grdIngresosDiarios.TextMatrix(i, cintColDescuentoI)
                gridFinal.TextMatrix(i, cintColIVACobrado) = grdIngresosDiarios.TextMatrix(i, cintColIVACobradoI)
                gridFinal.TextMatrix(i, cintColIVANoCobrado) = grdIngresosDiarios.TextMatrix(i, cintColIVANoCobradoI)
                gridFinal.TextMatrix(i, cintColTotal) = grdIngresosDiarios.TextMatrix(i, cintColTotalI)
                gridFinal.TextMatrix(i, cintColICG) = grdIngresosDiarios.TextMatrix(i, cintColICGI)
                gridFinal.TextMatrix(i, cintColICNG) = grdIngresosDiarios.TextMatrix(i, cintColICNGI)
                gridFinal.TextMatrix(i, cintColINCG) = grdIngresosDiarios.TextMatrix(i, cintColINCGI)
                gridFinal.TextMatrix(i, cintColINCNG) = grdIngresosDiarios.TextMatrix(i, cintColINCNGI)
                
                 If Trim(grdIngresosDiarios.TextMatrix(i, cintColTipoPacienteI)) = "I" Then
                    gridFinal.TextMatrix(i, cintColTipoPaciente + intformaspago) = "Interno"
                ElseIf Trim(grdIngresosDiarios.TextMatrix(i, cintColTipoPacienteI)) = "E" Then
                    gridFinal.TextMatrix(i, cintColTipoPaciente + intformaspago) = "Externo"
                ElseIf Trim(grdIngresosDiarios.TextMatrix(i, cintColTipoPacienteI)) = "C" Then
                    gridFinal.TextMatrix(i, cintColTipoPaciente + intformaspago) = "Cliente"
                ElseIf Trim(grdIngresosDiarios.TextMatrix(i, cintColTipoPacienteI)) = "G" Then
                    gridFinal.TextMatrix(i, cintColTipoPaciente + intformaspago) = "Grupo de cuentas"
                ElseIf Trim(grdIngresosDiarios.TextMatrix(i, cintColTipoPacienteI)) = "V" Then
                    gridFinal.TextMatrix(i, cintColTipoPaciente + intformaspago) = "Venta al público"
                End If
                
                gridFinal.AddItem 1
            Else
                gridFinal.TextMatrix(i, cintColCliente) = grdIngresosDiarios.TextMatrix(i, cintColClienteI)
                gridFinal.TextMatrix(i, cintColFecha) = Format(grdIngresosDiarios.TextMatrix(i, cintColFechaI), "dd/MM/YYYY")
                gridFinal.TextMatrix(i, cIntColRFC) = grdIngresosDiarios.TextMatrix(i, cintColRFCI)
                gridFinal.TextMatrix(i, cintColUUID) = grdIngresosDiarios.TextMatrix(i, cintColUUIDI)
                gridFinal.TextMatrix(i, cintColTipo) = grdIngresosDiarios.TextMatrix(i, cintColTipoI)
                gridFinal.TextMatrix(i, cIntColFolio) = grdIngresosDiarios.TextMatrix(i, cintColFolioI)
                gridFinal.TextMatrix(i, cintColDepartamento) = grdIngresosDiarios.TextMatrix(i, cintColDepartamentoI)
                gridFinal.TextMatrix(i, cintColImporte) = grdIngresosDiarios.TextMatrix(i, cintColImporteI)
                gridFinal.TextMatrix(i, cintColDescuento) = grdIngresosDiarios.TextMatrix(i, cintColDescuentoI)
                gridFinal.TextMatrix(i, cintColIVACobrado) = grdIngresosDiarios.TextMatrix(i, cintColIVACobradoI)
                gridFinal.TextMatrix(i, cintColIVANoCobrado) = grdIngresosDiarios.TextMatrix(i, cintColIVANoCobradoI)
                gridFinal.TextMatrix(i, cintColTotal) = grdIngresosDiarios.TextMatrix(i, cintColTotalI)
                gridFinal.TextMatrix(i, cintColICG) = grdIngresosDiarios.TextMatrix(i, cintColICGI)
                gridFinal.TextMatrix(i, cintColICNG) = grdIngresosDiarios.TextMatrix(i, cintColICNGI)
                gridFinal.TextMatrix(i, cintColINCG) = grdIngresosDiarios.TextMatrix(i, cintColINCGI)
                gridFinal.TextMatrix(i, cintColINCNG) = grdIngresosDiarios.TextMatrix(i, cintColINCNGI)
                
                If Trim(grdIngresosDiarios.TextMatrix(i, cintColTipoPacienteI)) = "I" Then
                    gridFinal.TextMatrix(i, cintColTipoPaciente + intformaspago) = "Interno"
                ElseIf Trim(grdIngresosDiarios.TextMatrix(i, cintColTipoPacienteI)) = "E" Then
                    gridFinal.TextMatrix(i, cintColTipoPaciente + intformaspago) = "Externo"
                ElseIf Trim(grdIngresosDiarios.TextMatrix(i, cintColTipoPacienteI)) = "C" Then
                    gridFinal.TextMatrix(i, cintColTipoPaciente + intformaspago) = "Cliente"
                ElseIf Trim(grdIngresosDiarios.TextMatrix(i, cintColTipoPacienteI)) = "G" Then
                    gridFinal.TextMatrix(i, cintColTipoPaciente + intformaspago) = "Grupo de cuentas"
                ElseIf Trim(grdIngresosDiarios.TextMatrix(i, cintColTipoPacienteI)) = "V" Then
                    gridFinal.TextMatrix(i, cintColTipoPaciente + intformaspago) = "Venta al público"
                End If
                gridFinal.AddItem 1
            End If

            Dim Folio As String
            Folio = gridFinal.TextMatrix(i, cIntColFolio)
            For j = 1 To grdFormasPago.Rows - 1
                If Trim(Folio) = Trim(grdFormasPago.TextMatrix(j, cintColFolioFP)) Then
                    Dim clave As String
                    Dim cantPagada As Double
                    Dim blnreciboAnt As Boolean
                    Dim foliorecibo As String
                    If Trim(grdFormasPago.TextMatrix(j, cintColChrFolioReciboFP)) = "" Then
                        blnreciboAnt = False
                    Else
                        foliorecibo = Trim(grdFormasPago.TextMatrix(j, cintColChrFolioReciboFP))
                        blnreciboAnt = True
                    End If
                    cantPagada = grdFormasPago.TextMatrix(j, cintColCantidadPagadaFP)
                    clave = Format(Trim(grdFormasPago.TextMatrix(j, cintColClaveFormaPagoFP)), "0##")
                    If blnreciboAnt = False Then
                        For k = 16 To 16 + intformaspago
                            If Mid(gridFinal.TextMatrix(0, k), 1, 3) = clave Then
                              gridFinal.TextMatrix(i, k) = cantPagada
                            End If
                        Next k
                    Else
                        Dim valrecibo As String
                        valrecibo = gridFinal.TextMatrix(i, cintColObservacion + intformaspago)
                        gridFinal.TextMatrix(i, cintColObservacion + intformaspago) = valrecibo & ", " & foliorecibo
                        If gridFinal.TextMatrix(i, cintColAnticipo + intformaspago) = "" Then
                            gridFinal.TextMatrix(i, cintColAnticipo + intformaspago) = 0
                        End If
                        Dim cantAcum As Double
                        cantAcum = gridFinal.TextMatrix(i, cintColAnticipo + intformaspago)
                        gridFinal.TextMatrix(i, cintColAnticipo + intformaspago) = cantAcum + cantPagada
                    End If
                    
                End If
                
            Next j
            
        Next i
        For quitaComas = 1 To gridFinal.Rows - 1
            gridFinal.TextMatrix(quitaComas, cintColObservacion + intformaspago) = Mid(gridFinal.TextMatrix(quitaComas, cintColObservacion + intformaspago), 2, Len(gridFinal.TextMatrix(quitaComas, cintColObservacion + intformaspago)))
        Next
        gridFinal.TextMatrix(gridFinal.Rows - 1, 0) = ""
        
        For p = 1 To grdFormasPago.Rows - 1
            Dim foliorec As String
            Dim foliofac As String
                If Not Trim(grdFormasPago.TextMatrix(p, cintColChrFolioReciboFP)) = "" Then
                    foliorec = Trim(grdFormasPago.TextMatrix(p, cintColChrFolioReciboFP))
                    foliofac = Trim(grdFormasPago.TextMatrix(p, cintColFolioFP))
                    For q = 0 To gridFinal.Rows - 1
                        If foliorec = Trim(gridFinal.TextMatrix(q, cIntColFolio)) Then
                            gridFinal.TextMatrix(q, cintColObservacion + intformaspago) = foliofac
                        End If
                    Next q
                End If
        Next p
        Dim rsCancelados As New ADODB.Recordset
        Set rsCancelados = frsRegresaRs("SELECT CLIENTE, FOLIO, CASE NVL(DC.CHRFOLIODOCUMENTO, '0') WHEN '0' THEN '' ELSE 'MD' END MISMODIA FROM CNTMPRPTIVATRASLADADO " & _
        "INNER JOIN PVFACTURA ON TRIM(CNTMPRPTIVATRASLADADO.FOLIO) = TRIM(PVFACTURA.CHRFOLIOFACTURA) " & _
        "INNER JOIN PVDOCUMENTOCANCELADO DC ON TRIM(PVFACTURA.CHRFOLIOFACTURA) = TRIM(DC.CHRFOLIODOCUMENTO) AND TO_CHAR(PVFACTURA.DTMFECHAHORA, 'DD/MM/YYYY') = TO_CHAR(DC.DTMFECHA, 'DD/MM/YYYY')  AND DC.CHRTIPODOCUMENTO = 'FA' " & _
        "GROUP BY CLIENTE, FOLIO, DC.CHRFOLIODOCUMENTO")
        '---------------------------------------------------------------------------------------
        '  Facturas generadas y canceladas el mismo dia poniendo montos en 0(Funcion CANCELADA)
        '---------------------------------------------------------------------------------------
        If rsCancelados.RecordCount > 0 Then
            For s = 0 To gridFinal.Rows - 1
                If Trim(rsCancelados!Folio) = Trim(gridFinal.TextMatrix(s, cIntColFolio)) And Trim(gridFinal.TextMatrix(s, cintColTipo)) = "Factura" Then
'                    For t = cintColImporte To intformaspago + cintColAnticipo
'                        gridFinal.TextMatrix(s, t) = "0"
'                    Next t
                End If
            Next s
        End If
        rsCancelados.Close
        
        
'       exec sp_PvRptFacturaCancelada('2022-06-15','2022-06-15',29,'*', 0, 1,0, :rc1);
'-- fecha inicio, fecha fin , departamento, tipopaciente, procedencia, clavehospital, tipo
        Dim rsfoliossustitutos As New ADODB.Recordset
        Dim queryfoliossustitutos As String
        queryfoliossustitutos = fecha1 & "|" & fecha1 & "|" & cboDepartamento.ItemData(cboDepartamento.ListIndex) & "|" & "'*'" & "|" & 0 & "|" & 1 & "|" & 0
        Set rsfoliossustitutos = frsEjecuta_SP(queryfoliossustitutos, "sp_pvrptfacturacancelada")
        If rsfoliossustitutos.RecordCount > 0 Then
            Do While Not rsfoliossustitutos.EOF
                For t = 1 To gridFinal.Rows - 1
                Dim folact As String
                folact = Trim(gridFinal.TextMatrix(t, cIntColFolio))
                    If Trim(rsfoliossustitutos!Folio) = folact Then
                        gridFinal.TextMatrix(t, cintColObservacion + intformaspago) = gridFinal.TextMatrix(t, cintColObservacion + intformaspago) & IIf(Trim(rsfoliossustitutos!FolioNuevo) = "", "", "Folio sustituto " & rsfoliossustitutos!FolioNuevo)
                    End If
                Next t
                rsfoliossustitutos.MoveNext
            Loop
        End If
        
        '--------------------------------------------------------
        '       Rellenando las columnas no utilizadas con 0
        '--------------------------------------------------------
        For intfillzero = 1 To gridFinal.Rows - 2
            For fillcols = 16 To cintColAnticipo + intformaspago
                If Trim(gridFinal.TextMatrix(intfillzero, fillcols)) = "" Then
                    gridFinal.TextMatrix(intfillzero, fillcols) = "0"
                End If
            Next
            For fillcols2 = cintColDif + intformaspago To cintColPruebaIVA + intformaspago
                If gridFinal.TextMatrix(intfillzero, cintColIVACobrado) = "0" And gridFinal.TextMatrix(intfillzero, cintColIVANoCobrado) = "0" Then
                    gridFinal.TextMatrix(intfillzero, cintColPruebaIVA + intformaspago) = "NO IVA"
                End If
            Next
        Next
        
        
         '----------------------------------------------------------------------------------------
        '      Cambiando los montos a negativo en facturas canceladas y entradas canceladas
        '----------------------------------------------------------------------------------------
        For intfillnegform = 1 To gridFinal.Rows - 2
            If Trim(gridFinal.TextMatrix(intfillnegform, cintColTipo)) = "Factura cancelada" Or Trim(gridFinal.TextMatrix(intfillnegform, cintColTipo)) = "Entrada cancelada" Or Trim(gridFinal.TextMatrix(intfillnegform, cintColTipo)) = "Entrada aplicada" Then
                For fillcols = 16 To cintColAnticipo + intformaspago
                    Dim valoractual As Double
                    valoractual = gridFinal.TextMatrix(intfillnegform, fillcols)
                    If valoractual > 0 Then
                        gridFinal.TextMatrix(intfillnegform, fillcols) = valoractual * -1
                    End If
                Next
                'En esta parte elimine la factura en la conversion a negativo
            ElseIf Trim(gridFinal.TextMatrix(intfillnegform, cintColTipo)) = "Entrada" Then
                 For fillcols = 16 To cintColAnticipo + intformaspago
                    Dim valoractual2 As Double
                    valoractual2 = gridFinal.TextMatrix(intfillnegform, fillcols)
                    If valoractual2 < 0 Then
                        gridFinal.TextMatrix(intfillnegform, fillcols) = valoractual2 * -1
                    End If
                Next
            
            End If
        Next
        
        
        'REALIZANDO CALCULOS DE DIFERENCIA Y DIFERENCIA DE IVA
        '--------------------------------------------------------
        '       Calculo de diferencia
        '--------------------------------------------------------
        For introwsdif = 1 To gridFinal.Rows - 1
            Dim sumaformaspago As Double
            Dim resultadodif As Double
            sumaformaspago = 0
            For colsgrid = 16 To cintColAnticipo + intformaspago
                sumaformaspago = sumaformaspago + Val(gridFinal.TextMatrix(introwsdif, colsgrid))
            Next colsgrid
            If sumaformaspago < 0 Then
                resultadodif = Val(gridFinal.TextMatrix(introwsdif, cintColTotal)) - sumaformaspago
            Else
                resultadodif = Val(gridFinal.TextMatrix(introwsdif, cintColTotal)) - sumaformaspago
            End If
            gridFinal.TextMatrix(introwsdif, cintColDif + intformaspago) = Format(resultadodif, "##.00")
        Next introwsdif
        '-----------------------------------------
        '       Calculo de diferencia IVA
        '-----------------------------------------
        For introwsdifiva = 1 To gridFinal.Rows - 2
            Dim ICG As Double
            Dim INCG As Double
            Dim ivacobrado As Double
            Dim ivanocobrado As Double
            Dim resultadodifiva As Double
            
            ICG = Val(gridFinal.TextMatrix(introwsdifiva, cintColICG))
            INCG = Val(gridFinal.TextMatrix(introwsdifiva, cintColINCG))
            ivacobrado = Val(gridFinal.TextMatrix(introwsdifiva, cintColIVACobrado))
            ivanocobrado = Val(gridFinal.TextMatrix(introwsdifiva, cintColIVANoCobrado))
            resultadodifiva = ((ICG + INCG) * relporcentajeiva) - (ivanocobrado + ivacobrado)
            gridFinal.TextMatrix(introwsdifiva, cintColPruebaIVA + intformaspago) = Format(resultadodifiva, "##.00")
        Next introwsdifiva
        
'        Dim intCuentaEliminadas As Integer
'        intCuentaEliminadas = -3
'        For introwsentradasdobles = 1 To gridFinal.Rows + intCuentaEliminadas
'            If Trim(gridFinal.TextMatrix(introwsentradasdobles, cintColTipo)) = "Entrada aplicada" Then
'                Dim folioFact As String
'                folioFact = Trim(gridFinal.TextMatrix(introwsentradasdobles, cintColFolio))
'                    For intentradasrow = 1 To gridFinal.Rows - 2
'                        If Trim(gridFinal.TextMatrix(intentradasrow, cintColTipo)) = "Entrada" Then
'                            If Trim(gridFinal.TextMatrix(intentradasrow, cintColFolio)) = folioFact Then
'                                gridFinal.RemoveItem (intentradasrow)
'                                introwsentradasdobles = introwsentradasdobles - 1
'                                intCuentaEliminadas = intCuentaEliminadas - 1
'                            End If
'
'                        End If
'                    Next
'            End If
'        Next
        
        For o = cintColImporte To intformaspago + cintColAnticipo
            gridFinal.Subtotal flexSTSum, cintColTipo, o, , , , , , cintColTipo, True
        Next o
        
        For intfillMayus = 1 To gridFinal.Rows - 2
            If Trim(gridFinal.TextMatrix(intfillMayus, 0)) = "Total Entrada" Then
                gridFinal.TextMatrix(intfillMayus, 0) = "Total entrada"
            End If
            If Trim(gridFinal.TextMatrix(intfillMayus, 0)) = "Total Entrada aplicada" Then
                gridFinal.TextMatrix(intfillMayus, 0) = "Total entrada aplicada"
            End If
            If Trim(gridFinal.TextMatrix(intfillMayus, 0)) = "Total Entrada cancelada" Then
                gridFinal.TextMatrix(intfillMayus, 0) = "Total entrada cancelada"
            End If
            If Trim(gridFinal.TextMatrix(intfillMayus, 0)) = "Total Factura" Then
                gridFinal.TextMatrix(intfillMayus, 0) = "Total factura"
            End If
            If Trim(gridFinal.TextMatrix(intfillMayus, 0)) = "Total Factura cancelada" Then
                gridFinal.TextMatrix(intfillMayus, 0) = "Total factura cancelada"
            End If
        Next
        
        For correccionEspacios = 1 To gridFinal.Rows - 1
            For columnasEspacios = 1 To gridFinal.Cols - 1
                Dim ValNow As String
                ValNow = gridFinal.TextMatrix(correccionEspacios, columnasEspacios)
                gridFinal.TextMatrix(correccionEspacios, columnasEspacios) = Trim(ValNow)
            Next
        Next
        
        '-------------------------------------------------------------------------------------------------
        
        '*********************************************************
        '           COMENZAMOS EXPORTACION
        '*********************************************************
    
On Error GoTo NotificaErrorExportacion
    
        'caso 20550 se realiza los ajustes para el cambio del reporte GIRM 09-10-2024
1        Dim o_Excel As Object
2        Dim o_Libro As Object
3        Dim o_Sheet As Object
4        Dim intRow As Long
5        Dim intCol As Integer
6        Dim dblAvance As Double
7        Dim intRowExcel As Long
8        Dim intDia As Integer
9        Dim dteFechaInicio As Date
10        Dim dteFechafin As Date
11        Dim intMeses  As Long
12        Dim lngMes As Long
13        Dim dblImporte As Double 'Columna 7
14        Dim dblDescuento As Double 'Columna 8
15        Dim dblIVAcobrado As Double 'Columna 9
16        Dim dblIVAnocobrado As Double 'Columna 10
17        Dim dblTotal As Double 'Columna 11
18        Dim dblIngresocobradogravado As Double 'Columna 12
19        Dim dblIngresocobradonogravado As Double 'Columna 13
20        Dim dblIngresonocobradogravado As Double 'Columna 14
21        Dim dblIngresonocobradonogravado As Double 'Columna 15
22        Dim lngFechasExpo As Long
23        Dim blnExisteHoja As Boolean
          Dim nameSheet As String
          
        
25       If gridFinal.Rows > 1 And gridFinal.TextMatrix(1, 1) <> "" Then
26           'Se crea Libro y hoja de excel
27            Set o_Excel = CreateObject("Excel.Application")

28            Set o_Libro = o_Excel.Workbooks.Add
29            Set o_Sheet = o_Libro.worksheets(1)

              nameSheet = o_Sheet.Name

            'Valida si no existe Excel ene el equipo
30            If Not IsObject(o_Excel) Then
                MsgBox "Necesitas Microsoft Excel para utilizar esta funcionalidad", _
                vbExclamation, "Mensaje"
31                Exit Sub
            End If
            
            'datos del repote
32            o_Excel.cells(2, 1).Value = Trim(vgstrNombreHospitalCH)
33            o_Excel.cells(3, 1).Value = "INGRESOS DIARIOS"
34            o_Excel.cells(4, 1).Value = "Rango de fechas del " & CStr(Format(CDate(txtFechaInicio.Text), "dd/MMM/yyyy")) & " al " & CStr(Format(CDate(txtFechaFin.Text), "dd/MMM/yyyy"))
            
            'Columnas titulos
35            For l = 0 To gridFinal.Cols - 1
36                Dim minus As String
37                Dim mayus As String
38                mayus = gridFinal.TextMatrix(0, l)
    
39                o_Excel.cells(6, l + 1).Value = mayus
40                o_Excel.cells(6, l + 1).HorizontalAlignment = -4108
41                o_Excel.cells(6, l + 1).Interior.ColorIndex = 15 '15 48
42                o_Excel.cells(6, l + 1).Font.Bold = True
43                o_Excel.cells(6, l + 1).Font.Name = "Times New Roman"
44                o_Excel.cells(6, l + 1).WrapText = True
            Next l
            
45            For n = 16 To cintColAnticipo + intformaspago - 1
46                Dim minus1 As String
47                Dim mayus1 As String
48                Dim pref As String
49                pref = Mid(gridFinal.TextMatrix(0, n), 1, 3)
50                mayus1 = UCase(Mid(gridFinal.TextMatrix(0, n), 5, 1))
51                minus1 = UCase(Mid(gridFinal.TextMatrix(0, n), 6, Len(gridFinal.TextMatrix(0, n))))
52                o_Excel.cells(6, n + 1).Value = pref & " " & mayus1 & minus1
            Next n
            
            'Formato de monedas
53            o_Sheet.range(o_Excel.cells(6, cintColImporte), o_Excel.cells(gridFinal.Rows - 1 + 6, cintColAnticipo + intformaspago + 2)).NumberFormat = "$ ###,###,###,##0.00"
54            o_Sheet.range(o_Excel.cells(6, cintColDif + intformaspago), o_Excel.cells(gridFinal.Rows - 1 + 6, cintColPruebaIVA + intformaspago + 2)).NumberFormat = "$ ###,###,###,##0.00"
            
55            o_Sheet.range(o_Excel.cells(6, cintColCliente + 1), o_Excel.cells(gridFinal.Rows - 1 + 6, cintColPruebaIVA + intformaspago + 2)).Columnwidth = 15

56            o_Sheet.range("A:A").Columnwidth = 35

            'Variable para la asignacion de el tamaño de la barra de progreso
57            dblAvance = 100 / gridFinal.Rows
    
58            intRowExcel = 2
            
            'Validamos el numero de meses entre las fechas
59            dteFechaInicio = txtFechaInicio.Text
60            dteFechafin = txtFechaFin.Text
            
            'Almacenas las fechas en un arreglo
61            Dim strFechas() As String
62            Dim lngNumber As Long
63            Dim lngFechas As Long
64            For lngNumber = 2 To gridFinal.Rows - 1
65                If gridFinal.TextMatrix(lngNumber - 1, 1) <> "" Then
66                    If CDate(gridFinal.TextMatrix(lngNumber - 1, 1)) >= CDate(txtFechaInicio.Text) And CDate(gridFinal.TextMatrix(lngNumber - 1, 1)) <= CDate(txtFechaFin.Text) Then
67                        If lngNumber = 2 Then
68                            ReDim strFechas(lngNumber - 2)
69                            strFechas(lngNumber - 2) = gridFinal.TextMatrix(lngNumber - 1, 1)
70                        Else
71                            If gridFinal.TextMatrix(lngNumber - 1, 1) <> "" Then
72                                ReDim Preserve strFechas(UBound(strFechas()) + 1)
73                                strFechas(UBound(strFechas())) = gridFinal.TextMatrix(lngNumber - 1, 1)
                              End If
                          End If
                     End If
                  End If
            Next lngNumber
            
            
            'Eliminar valores repetidos
74            Dim strUnicosCollection As Collection
75            Dim lngContador As Long
            
            'Creamos la colecciona para alamacenar los vamores únicos
76            Set strUnicosCollection = New Collection
            
            'Purgamos los valores repetidos
77            For lngContador = LBound(strFechas) To UBound(strFechas)
78                If Not ExisteEnColeccion(strUnicosCollection, strFechas(lngContador)) Then
79                    strUnicosCollection.Add strFechas(lngContador), CStr(strFechas(lngContador))
                End If
            Next lngContador
            
            'Transferir los valores únicos de la colección al arreglo
80            ReDim strFechas(strUnicosCollection.Count - 1)
81            For lngContador = 1 To strUnicosCollection.Count
82                strFechas(lngContador - 1) = strUnicosCollection(lngContador)
            Next
            
            'Ordenamos el arreglo por el método BubbleSort
83           Call ArraySortBS(strFechas())
            
            'Empezamos la exportación
84            For lngFechasExpo = 0 To UBound(strFechas)
85               intRowExcel = 2
                 blnExisteHoja = False
                
                'Agrega titulos a la nueva hoja
                If lngFechasExpo = 0 Then
                For Each Sheet In o_Libro.sheets
1101                intHoja = intHoja + 1
1102                If intHoja > 1 And intHoja <= 3 Then
1103                    If Sheet.Name = "Hoja" & intHoja Then
1105                        Set o_Sheet = o_Libro.worksheets(IIf(idiomaExcel = 3082, "Hoja", "Sheet") & intHoja)
                            o_Sheet.Delete
                        End If
                    End If
                Next
                Set o_Sheet = o_Libro.worksheets(nameSheet)
            ElseIf lngFechasExpo > 0 Then
                    
87                  If Month(strFechas(lngFechasExpo)) <> Month(strFechas(lngFechasExpo - 1)) Then
88                        o_Excel.Visible = True
89                        Set o_Excel = Nothing
                        'se crea un nuevo archivo
90                        Set o_Excel = CreateObject("Excel.Application")
91                        Set o_Libro = o_Excel.Workbooks.Add
                          intHoja = 0
                          For Each Sheet In o_Libro.sheets
                            intHoja = intHoja + 1
                            If intHoja > 1 And intHoja <= 3 Then
                                If Sheet.Name = "Hoja" & intHoja Then
                                    Set o_Sheet = o_Libro.worksheets(IIf(idiomaExcel = 3082, "Hoja", "Sheet") & intHoja)
                                     o_Sheet.Delete
                                 End If
                             Else
                                blnExisteHoja = True
                                nameSheet = Sheet.Name
                             End If
                        Next
                    End If
                    
                    If blnExisteHoja = False Then
                        Set o_Sheet = o_Libro.worksheets.Add
                        nameSheet = o_Sheet.Name
                    End If
                    
                    Set o_Sheet = o_Libro.worksheets(nameSheet)
                    o_Sheet.Move after:=o_Libro.sheets(o_Libro.sheets.Count)

                    
103                    If Not IsObject(o_Excel) Then
                        MsgBox "Necesitas Microsoft Excel para utilizar esta funcionalidad", _
                            vbExclamation, "Mensaje"
104                        Exit Sub
                    End If
                    
                    'datos del repote
105                    o_Sheet.cells(2, 1).Value = Trim(vgstrNombreHospitalCH)
106                    o_Sheet.cells(3, 1).Value = "INGRESOS DIARIOS"
107                    o_Sheet.cells(4, 1).Value = "Rango de fechas del " & CStr(Format(CDate(txtFechaInicio.Text), "dd/MMM/yyyy")) & " al " & CStr(Format(CDate(txtFechaFin.Text), "dd/MMM/yyyy"))

                    'Columnas titulos
108                    For l = 0 To gridFinal.Cols - 1
109                        mayus = gridFinal.TextMatrix(0, l)
            
110                        o_Sheet.cells(6, l + 1).Value = mayus
111                        o_Sheet.cells(6, l + 1).HorizontalAlignment = -4108
112                        o_Sheet.cells(6, l + 1).Interior.ColorIndex = 15 '15 48
113                        o_Sheet.cells(6, l + 1).Font.Bold = True
114                        o_Sheet.cells(6, l + 1).Font.Name = "Times New Roman"
115                        o_Sheet.cells(6, l + 1).WrapText = True
            
                    Next l
                       
116                    For n = 16 To cintColAnticipo + intformaspago - 1
117                        pref = Mid(gridFinal.TextMatrix(0, n), 1, 3)
118                        mayus1 = UCase(Mid(gridFinal.TextMatrix(0, n), 5, 1))
119                        minus1 = UCase(Mid(gridFinal.TextMatrix(0, n), 6, Len(gridFinal.TextMatrix(0, n))))
120                        o_Excel.cells(6, n + 1).Value = pref & " " & mayus1 & minus1
                    Next n
            
121                    o_Sheet.range(o_Sheet.cells(6, cintColImporte), o_Sheet.cells(gridFinal.Rows - 1 + 6, cintColAnticipo + intformaspago + 2)).NumberFormat = "$ ###,###,###,##0.00"
122                    o_Sheet.range(o_Sheet.cells(6, cintColDif + intformaspago), o_Sheet.cells(gridFinal.Rows - 1 + 6, cintColPruebaIVA + intformaspago + 2)).NumberFormat = "$ ###,###,###,##0.00"
            
            
123                    o_Sheet.range(o_Sheet.cells(6, cintColCliente + 1), o_Sheet.cells(gridFinal.Rows - 1 + 6, cintColPruebaIVA + intformaspago + 2)).Columnwidth = 15
            
124                    o_Sheet.range("A:A").Columnwidth = 35
            
                    'Variable para la asignacion de el tamaño de la barra de progreso
125                    dblAvance = 100 / gridFinal.Rows
            
126                    intRowExcel = 2
                    
                End If
                
                'Rorremos la cuadricula
127                For intRow = 2 To gridFinal.Rows - 2
128                    If gridFinal.TextMatrix(intRow - 1, 1) <> "" Then
129                        If strFechas(lngFechasExpo) = gridFinal.TextMatrix(intRow - 1, 1) Then
                            ' Actualización de la barra de estado
130                            If pgbBarra.Value + dblAvance < 100 Then
131                                pgbBarra.Value = pgbBarra.Value + dblAvance
                            Else
132                                pgbBarra.Value = 100
                            End If
                                        
133                            pgbBarra.Refresh
                            
                            'Variable para la asignacion de el tamaño de la barra de progreso
134                             dblAvance = 100 / gridFinal.Rows
                            
135                            If gridFinal.RowHeight(intRow - 1) > 0 Then
136                                With gridFinal
                                    
137                                    For m = 0 To gridFinal.Cols - 1
138                                        If m <> 1 Then
139                                            o_Sheet.cells(intRowExcel + 5, m + 1).Value = Trim(.TextMatrix(intRow - 1, m))
                                           End If
                                        
140                                        If m = 1 Then
141                                            o_Sheet.cells(intRowExcel + 5, m + 1).Value = "'" & Format(.TextMatrix(intRow - 1, m), "dd/mm/yyyy")
                                           End If
                                        
142                                        Select Case m
                                               Case 7:
143                                                    dblImporte = dblImporte + CDbl(Trim(.TextMatrix(intRow - 1, m)))
                                               Case 8:
144                                                    dblDescuento = dblDescuento + CDbl(Trim(.TextMatrix(intRow - 1, m)))
                                               Case 9:
145                                                    dblIVAcobrado = dblIVAcobrado + CDbl(Trim(.TextMatrix(intRow - 1, m)))
                                               Case 10:
146                                                    dblIVAnocobrado = dblIVAnocobrado + CDbl(Trim(.TextMatrix(intRow - 1, m)))
                                               Case 11:
147                                                    dblTotal = dblTotal + CDbl(Trim(.TextMatrix(intRow - 1, m)))
                                               Case 12:
148                                                    dblIngresocobradogravado = dblIngresocobradogravado + CDbl(Trim(.TextMatrix(intRow - 1, m)))
                                               Case 13:
149                                                    dblIngresocobradonogravado = dblIngresocobradonogravado + CDbl(Trim(.TextMatrix(intRow - 1, m)))
                                               Case 14:
150                                                    dblIngresonocobradogravado = dblIngresonocobradogravado + CDbl(Trim(.TextMatrix(intRow - 1, m)))
                                               Case 15:
151                                                    dblIngresonocobradonogravado = dblIngresonocobradonogravado + CDbl(Trim(.TextMatrix(intRow - 1, m)))
                                        End Select
                                    Next m
                                    
                                End With
152                                intRowExcel = intRowExcel + 1
                            End If
                        End If
                    Else
153                        With gridFinal
154                            For m = 0 To gridFinal.Cols - 1
155                                If m <> 1 Then
156                                    o_Sheet.cells(intRowExcel + 5, m + 1).Value = Trim(.TextMatrix(intRow - 1, m))
                                End If
                    
                                'sumatorias
157                                Select Case m
                                    Case 7:
158                                        o_Sheet.cells(intRowExcel + 5, m + 1).Value = dblImporte
159                                        dblImporte = 0
                                    Case 8:
160                                        o_Sheet.cells(intRowExcel + 5, m + 1).Value = dblDescuento
161                                        dblDescuento = 0
                                    Case 9:
162                                        o_Sheet.cells(intRowExcel + 5, m + 1).Value = dblIVAcobrado
163                                        dblIVAcobrado = 0
                                    Case 10:
164                                        o_Sheet.cells(intRowExcel + 5, m + 1).Value = dblIVAnocobrado
165                                        dblIVAnocobrado = 0
                                    Case 11:
166                                        o_Sheet.cells(intRowExcel + 5, m + 1).Value = dblTotal
167                                        dblTotal = 0
                                    Case 12:
168                                        o_Sheet.cells(intRowExcel + 5, m + 1).Value = dblIngresocobradogravado
169                                        dblIngresocobradogravado = 0
                                    Case 13:
170                                        o_Sheet.cells(intRowExcel + 5, m + 1).Value = dblIngresocobradonogravado
171                                        dblIngresocobradonogravado = 0
                                    Case 14:
172                                        o_Sheet.cells(intRowExcel + 5, m + 1).Value = dblIngresonocobradogravado
173                                        dblIngresonocobradogravado = 0
                                    Case 15:
174                                        o_Sheet.cells(intRowExcel + 5, m + 1).Value = dblIngresonocobradonogravado
175                                        dblIngresonocobradonogravado = 0
                                End Select
                                    
176                                o_Sheet.cells(intRowExcel + 5, m + 1).Font.Bold = True
                            Next m
                        End With
177                        intRowExcel = intRowExcel + 2
                    End If
                Next intRow
                
178                o_Sheet.Name = Format(strFechas(lngFechasExpo), "dd-mmm-yyyy")
179                o_Sheet.range("B:B").Columnwidth = 11
180                o_Sheet.range("B:B").HorizontalAlignment = -4131
181                o_Sheet.range("D:D").Columnwidth = 40
                'o_Sheet.range("D:D").HorizontalAlignment = -4108
                
            Next lngFechasExpo
            
182            o_Excel.Visible = True
183            Set o_Excel = Nothing
            
184            txtFechaInicio.Text = dteFechaInicio
185            txtFechaFin.Text = dteFechafin
            
            'La información ha sido exportada exitosamente
            MsgBox SIHOMsg(1185), vbOKOnly + vbInformation, "Mensaje"
186            freBarra.Visible = False
            
            'Fin caso 20550

            For intRows = gridFinal.Rows - 1 To 1 Step -1
                  gridFinal.RemoveItem intRows
            Next intRows

            gridFinal.AddItem 1
            gridFinal.Clear flexCleaeverywhere, flexClearText
            gridFinal.Cols = 7

            For intRows = grdFormasPago.Rows - 1 To 1 Step -1
                  grdFormasPago.RemoveItem intRows
            Next intRows

            grdFormasPago.AddItem 1
            grdFormasPago.Clear flexCleaeverywhere, flexClearText
            grdFormasPago.Cols = 7

            For intRows = grdIngresosDiarios.Rows - 1 To 1 Step -1
                  grdIngresosDiarios.RemoveItem intRows
            Next intRows

            grdIngresosDiarios.AddItem 1
            grdIngresosDiarios.Clear flexCleaeverywhere, flexClearText
            grdIngresosDiarios.Cols = 7

        Else
            'No existe información con esos parámetros
            MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje"
        End If
    
    '-----------------------------------------------------------------------------------------------
    
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdExportar_Click"))
    
NotificaErrorExportacion:
  MsgBox "Ocurrió un error en la línea " & Erl & ": " & Err.Description, vbCritical, "Error"
  
  
End Sub

Private Sub Form_Activate()
    Dim fechaactual As Date
    Dim fechaInicial As Date
    fechaactual = fdtmServerFecha
    'txtFecha.Text = Format(fechaactual, "dd/mmm/yyyy")
    'txtFecha.Text = fechaactual
    fechaInicial = DateSerial(Year(fechaactual), Month(fechaactual), 1)
    pMkTextAsignaValor txtFechaInicio, Format(fechaInicial, "dd/mm/yyyy")
    pMkTextAsignaValor txtFechaFin, Format(fechaactual, "dd/mm/yyyy")
End Sub

Private Sub Form_Load()
Dim rs As New ADODB.Recordset
Dim rspermisodelusuario As New ADODB.Recordset
Dim fechaactual As String
Dim vlstrsql As String
If vglngNumeroLogin <> 0 Then
vlstrsql = "select chrPermiso from Permiso where intNumeroLogin= '" & str(vglngNumeroLogin) & "' and intNumeroOpcion= '" & 7020 & "'"
Set rspermisodelusuario = frsRegresaRs(vlstrsql)
If rspermisodelusuario.RecordCount > 0 Then

    If Not rspermisodelusuario!chrpermiso = "C" Then
        cboHospital.Enabled = False
        cboDepartamento.Enabled = False
        txtFecha.Enabled = False
        cmdExportar.Enabled = False
    Else
        
    End If
Else

End If
End If
Me.Icon = frmMenuPrincipal.Icon
fechaactual = Format(Now, "dd/mmm/yyyy")
    Set rs = frsEjecuta_SP("-1", "Sp_Gnselempresascontable")
    If rs.RecordCount <> 0 Then
       pLlenarCboRs cboHospital, rs, 1, 0
       cboHospital.ListIndex = flngLocalizaCbo(cboHospital, str(vgintClaveEmpresaContable))
    End If
    
End Sub


Private Sub txtFechaInicio_GotFocus()
    txtFechaInicio.SelStart = 0
    txtFechaInicio.SelLength = Len(txtFechaInicio.Text)
End Sub

Private Sub txtFechaFin_GotFocus()
    txtFechaFin.SelStart = 0
    txtFechaFin.SelLength = Len(txtFechaFin.Text)
End Sub

Private Sub txtFechaInicio_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub txtFechaInicio_LostFocus()
        vlblnfechavalida = False
    
        If IsDate(txtFechaInicio.Text) Then
            Dim fecha As Date
            fecha = txtFechaInicio.Text
            If DateDiff("d", Date, fecha) <= 0 Then
                vlblnfechavalida = True
            
                If KeyCode = vbKeyReturn Then SendKeys vbTab
            Else
                Call MsgBox("¡Fecha invalida!", vbExclamation, "Mensaje")
                pEnfocaMkTexto txtFechaInicio
            End If
        Else
            Call MsgBox("¡Fecha invalida!", vbExclamation, "Mensaje")
            pEnfocaMkTexto txtFechaInicio
        End If
    
End Sub

Private Sub txtFechaFin_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Function ExisteEnColeccion(Col As Collection, clave As String) As Boolean
    Dim contador As Long
    ExisteEnColeccion = False
    
    For contador = 1 To Col.Count
        If Col(contador) = clave Then
            ExisteEnColeccion = True
            Exit For
        End If
    Next
       
End Function

Sub ArraySortBS(arreglo() As String)
    
    Dim i As Long, j As Long
    Dim tmp As Variant
    
    For i = LBound(arreglo) To UBound(arreglo) - 1
        For j = LBound(arreglo) To UBound(arreglo) - i - 1
            If CDate(arreglo(j)) > CDate(arreglo(j + 1)) Then
                ' Intercambiar elementos
                temp = arreglo(j)
                arreglo(j) = arreglo(j + 1)
                arreglo(j + 1) = temp
            End If
        Next j
    Next i
End Sub


