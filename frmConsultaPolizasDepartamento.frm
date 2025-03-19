VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmConsultaPolizasDepartamento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pólizas"
   ClientHeight    =   8010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11385
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   11385
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cdgExcel 
      Left            =   3360
      Top             =   7320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame freBarra 
      Height          =   815
      Left            =   2575
      TabIndex        =   26
      Top             =   8040
      Width           =   6265
      Begin MSComctlLib.ProgressBar pgbBarra 
         Height          =   300
         Left            =   45
         TabIndex        =   27
         Top             =   480
         Width           =   6180
         _ExtentX        =   10901
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblTextoBarra 
         BackColor       =   &H80000002&
         Caption         =   " Exportando información, por favor espere..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   45
         TabIndex        =   28
         Top             =   180
         Width           =   6165
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000002&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   345
         Left            =   45
         Top             =   135
         Width           =   6180
      End
   End
   Begin VB.ComboBox cboInterfazPolizas 
      Height          =   315
      ItemData        =   "frmConsultaPolizasDepartamento.frx":0000
      Left            =   8520
      List            =   "frmConsultaPolizasDepartamento.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   24
      ToolTipText     =   "Interfaz para importación y exportación de pólizas contables"
      Top             =   7470
      Width           =   2685
   End
   Begin VB.Frame Frame2 
      Height          =   700
      Left            =   4320
      TabIndex        =   22
      Top             =   7095
      Width           =   2550
      Begin VB.CommandButton cmdImprimirPoliza 
         Height          =   495
         Left            =   555
         Picture         =   "frmConsultaPolizasDepartamento.frx":0004
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Imprimir"
         Top             =   138
         UseMaskColor    =   -1  'True
         Width           =   540
      End
      Begin VB.CommandButton cmdExportar 
         Caption         =   "&Exportar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1080
         TabIndex        =   13
         ToolTipText     =   "Exportar póliza"
         Top             =   135
         Width           =   1380
      End
      Begin VB.CommandButton cmdPreview 
         Height          =   495
         Left            =   60
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmConsultaPolizasDepartamento.frx":01A6
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Vista previa"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdImportacion 
         Caption         =   "I&mportar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2475
         TabIndex        =   14
         ToolTipText     =   "Importar póliza"
         Top             =   135
         Width           =   1380
      End
   End
   Begin VB.CheckBox chkConcentrado 
      Caption         =   "Concentrado"
      Height          =   255
      Left            =   2880
      TabIndex        =   9
      ToolTipText     =   "Mostrar el reporte concentrado"
      Top             =   6860
      Width           =   1260
   End
   Begin VB.CheckBox chkCuentaMayor 
      Caption         =   "Incluir la cuenta inmediata mayor"
      Height          =   255
      Left            =   4320
      TabIndex        =   10
      ToolTipText     =   "Incluir la cuenta inmediata mayor en el reporte"
      Top             =   6860
      Width           =   2695
   End
   Begin VB.Frame Frame4 
      Height          =   920
      Left            =   120
      TabIndex        =   21
      Top             =   6900
      Width           =   2490
      Begin VB.CommandButton cmdInvertir 
         Caption         =   "&Invertir selección"
         Height          =   375
         Left            =   40
         TabIndex        =   8
         ToolTipText     =   "Invertir selección de pólizas"
         Top             =   500
         Width           =   2400
      End
      Begin VB.CommandButton cmdSeleccionar 
         Caption         =   "&Seleccionar / Quitar selección"
         Height          =   375
         Left            =   40
         TabIndex        =   7
         ToolTipText     =   "Seleccionar / quitar selección de póliza"
         Top             =   120
         Width           =   2400
      End
   End
   Begin MSComDlg.CommonDialog CDgArchivo 
      Left            =   2760
      Top             =   7320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   ".txt"
      DialogTitle     =   "Exportación de pólizas"
      FileName        =   "poliza.txt"
      Filter          =   "Texto (*.txt)|*.txt| Todos los archivos (*.*)|*.*"
   End
   Begin VB.Frame Frame1 
      Height          =   1125
      Left            =   120
      TabIndex        =   23
      Top             =   0
      Width           =   11175
      Begin VB.TextBox txtFolioFin 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6795
         TabIndex        =   5
         ToolTipText     =   "Folio final"
         Top             =   720
         Width           =   1140
      End
      Begin VB.TextBox txtFolioIni 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4785
         TabIndex        =   4
         ToolTipText     =   "Folio inicial"
         Top             =   720
         Width           =   1140
      End
      Begin VB.ComboBox cboDepartamento 
         Height          =   315
         Left            =   1200
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Departamento"
         Top             =   285
         Width           =   2775
      End
      Begin VB.ComboBox cboTipo 
         Height          =   315
         ItemData        =   "frmConsultaPolizasDepartamento.frx":0348
         Left            =   8640
         List            =   "frmConsultaPolizasDepartamento.frx":035B
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "Selección del tipo de póliza"
         Top             =   285
         Width           =   1530
      End
      Begin MSMask.MaskEdBox mskFechaFin 
         Height          =   315
         Left            =   6795
         TabIndex        =   2
         ToolTipText     =   "Fecha fin de consulta"
         Top             =   285
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mmm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskFechaInicio 
         Height          =   315
         Left            =   4785
         TabIndex        =   1
         ToolTipText     =   "Fecha de inicio de consulta"
         Top             =   285
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mmm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label lblAl 
         Caption         =   "al"
         Height          =   195
         Left            =   6150
         TabIndex        =   31
         Top             =   780
         Width           =   420
      End
      Begin VB.Label lblFolio 
         Caption         =   "Folio"
         Height          =   195
         Left            =   4155
         TabIndex        =   30
         Top             =   780
         Width           =   465
      End
      Begin VB.Label Label5 
         Caption         =   "Departamento"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   315
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
         Height          =   195
         Left            =   8115
         TabIndex        =   19
         Top             =   345
         Width           =   315
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         Height          =   195
         Left            =   6150
         TabIndex        =   17
         Top             =   345
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Left            =   4155
         TabIndex        =   16
         Top             =   345
         Width           =   465
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grdDetallePoliza1 
      Height          =   1935
      Left            =   105
      TabIndex        =   15
      TabStop         =   0   'False
      ToolTipText     =   "Detalle de la póliza"
      Top             =   4725
      Width           =   11160
      _cx             =   19685
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
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   0
      Cols            =   5
      FixedRows       =   0
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   7
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   -1  'True
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   0
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VSFlex7LCtl.VSFlexGrid grdPolizas1 
      Height          =   3030
      Left            =   120
      TabIndex        =   6
      ToolTipText     =   "Listado de pólizas"
      Top             =   1320
      Width           =   11160
      _cx             =   19685
      _cy             =   5345
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
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   0
      Cols            =   7
      FixedRows       =   0
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   7
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   -1  'True
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   0
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Desde"
      Height          =   195
      Left            =   120
      TabIndex        =   29
      Top             =   0
      Width           =   465
   End
   Begin VB.Label lblInterfaz 
      Caption         =   "Interfaz para pólizas"
      Height          =   255
      Left            =   8520
      TabIndex        =   25
      Top             =   7230
      Width           =   1575
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Detalle de la póliza"
      Height          =   195
      Left            =   120
      TabIndex        =   18
      Top             =   4425
      Width           =   1335
   End
End
Attribute VB_Name = "frmConsultaPolizasDepartamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************
' Consulta de las pólizas generadas en un rago de fechas, solo el departamento
' puede consultar e imprimir la póliza
' Fecha de programación: Jueves 16 de Noviembre del 2000
'*******************************************************************************

Private vgrptReporte As CRAXDRT.Report
Public vllngNumOpcion As Long

'- Caso 7442: Agregados para consulta de pólizas fuera del corte -'
Public vlblnPolizasFueraCorte As Boolean
Public vllngNumCorte As Long
Public vldtmFechaIni As Date
Public vldtmFechaFin As Date
'-----------------------------------------------------------------'

'- Agregados para la importación de pólizas en formato Microsip -'
Const CONTPAQ = 1
Const MICROSIP = 2
Const APSI = 3
Const intPosNivel = 0       '|  Indica la posición del nivel de la póliza (para indicar si es maestro o detalle)
'|  Maestro de la póliza
Const intPosTipo = 1        '|  Indica la posición del tipo de la póliza
Const intPosPoliza = 2      '|  Indica la posición de la clave de la póliza
Const intPosFecha = 2       '|  Indica la posición de la fecha
Const intPosMoneda = 3      '|  Indica la posición del tipo de la moneda
Const intPosTipoCambio = 4  '|  Indica la posición del tipo de cambio
Const intPosCancelada = 5   '|  Indica la posición del estatus de la póliza
Const intPosDescripcion = 5 '|  Indica la descripción de la póliza
Const intCamposMaestro = 6  '|  Indica el número de campos que debe tener el maestro de cada póliza
'|  Detalle de la póliza
Const intPosCuenta = 1      '|  Indica la posición del número de la cuenta
Const intPosNaturaleza = 3  '|  Indica la posición de la naturaleza de la cuenta
Const intPosCantidad = 4    '|  Indica la cantidad del detalle de la póliza
Const intPosReferencia = 5  '|  Indica la referencia del detalle de la póliza
Const intPosConcepto = 6    '|  Indica el concepto del detalle de la póliza
Const intCamposDetalle = 7  '|  Indica el número de campos que debe tener el detalle de cada póliza
'|  Configuración de la exportación
Const lstrDelimitador = ","             '|  Indica el delimitador de los datos de cada renglón de la póliza
Const lstrSeparador = "|"               '|  Indica el separador para el tipo de póliza
Const lstrFormatoFechas = "dd/mm/yyyy"  '|  Indica el formato de la fecha de la póliza
Const lstrCalificadorTexto = """"       '|  Indica el calificador de texto para las cadenas en cada renglón de la póliza
'-----------------------------------------------------------------'

'- Agregados para la importación de pólizas en formato APSI -'
'Columnas que el archivo excel debe de tener para guardar la información correctamente
Const colCuenta = 1
Const colDebito = 2
Const colCredito = 3
Const colTipo = 4
Const colFecha = 5
Const colConcepto = 6
Const colUuid = 7
Const colTotal = 8
Const colRfc = 9
Const colNombre = 10

Dim vlstrNombredelDepartamento As String
Dim vlstrsql As String
Dim rsCnSelPoliza As New ADODB.Recordset
Dim lblnSel As Boolean
Dim lblnExistePolizas As Boolean
Dim lblnEjecutarBusqueda As Boolean
Dim vldtmFechaInicial As Date
Dim vldtmFechaFinal As Date

Dim rsCuentaMayor As New ADODB.Recordset    'Agregada para indicar si se seleccionó la opción de Incluir cuenta inmediata mayor
Dim lintProceso As Integer                  'Agregado para indicar el proceso para la opción Incluir cuenta inmediata mayor

Dim lintInterfazPolizas As Integer          'Agregado para indicar que interfaz para la importación/exportación de pólizas se utilizará
Dim lstrArchivoUuid As String       'Ruta de archivo excel con uuids de polizas a importar
Dim lblnArchivoUuid As Boolean      'indica si el archivo a importar incluye los UUIDs, por lo tanto se utiliza otro formato

Private Sub cboDepartamento_Click()
    If cboDepartamento.ListIndex > -1 Then
        If lblnEjecutarBusqueda Then
            pCargaPolizas
            grdPolizas1.SetFocus
        End If
    End If
End Sub

Private Sub cboDepartamento_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub cboInterfazPolizas_Click()
    If cboInterfazPolizas.ListIndex <> -1 Then
        lintInterfazPolizas = cboInterfazPolizas.ItemData(cboInterfazPolizas.ListIndex)
        If fblnValidarLicenciaPolizas(vgintClaveEmpresaContable, cboInterfazPolizas.ItemData(cboInterfazPolizas.ListIndex), cboInterfazPolizas.Text) Then
            cmdExportar.Enabled = IIf(chkConcentrado.Value, False, IIf(grdPolizas1.Rows > 0, True, False))
            cmdImportacion.Enabled = True
        Else
            cmdExportar.Enabled = False
            cmdImportacion.Enabled = False
        End If
    End If
End Sub

Private Sub cboTipo_Click()
    If CDate(mskFechaInicio.Text) > CDate(mskFechaFin.Text) Then
        '¡Rango de fechas no válido!
        MsgBox SIHOMsg(64), vbOKOnly + vbInformation, "Mensaje"
        mskFechaInicio.SetFocus
    Else
        If lblnEjecutarBusqueda Then
            If cboTipo.List(cboTipo.ListIndex) = "Todas" Then
                lblFolio.Enabled = False
                lblAl.Enabled = False
                txtFolioIni.Enabled = False
                txtFolioFin.Enabled = False
                pLimpia
                pCargaPolizas
                grdPolizas1.SetFocus
            Else:
                lblFolio.Enabled = True
                lblAl.Enabled = True
                txtFolioIni.Enabled = True
                txtFolioFin.Enabled = True
                pCargaPolizas
                txtFolioIni.SetFocus
            End If
        End If
    End If
End Sub

Private Sub cboTipo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub chkConcentrado_Click()
    If cboInterfazPolizas.ListIndex <> -1 Then
        If fblnValidarLicenciaPolizas(vgintClaveEmpresaContable, cboInterfazPolizas.ItemData(cboInterfazPolizas.ListIndex), cboInterfazPolizas.Text) Then
            cmdExportar.Enabled = IIf(chkConcentrado.Value, False, IIf(grdPolizas1.Rows > 0, True, False))
        End If
    End If
End Sub

Private Sub chkConcentrado_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub chkCuentaMayor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub cmdExportar_Click()
    On Error GoTo NotificaError

    Dim vllngCont As Long
    Dim vllngCuadre As Long
    Dim vllngCancelada As Long
    Dim vllngCuentas As Long
    Dim vllngDescuadre As Long
    Dim llngPolizaDescuadrada As String
    
    For vllngCont = 1 To grdPolizas1.Rows - 1
        If grdPolizas1.TextMatrix(vllngCont, 0) = "*" Then
            vllngCuentas = vllngCuentas + 1
            '- Revisar si alguna póliza contiene alguna cuenta de cuadre -'
            If fblnCuentaCuadre(CLng(grdPolizas1.TextMatrix(vllngCont, 6))) Then
                vllngCuadre = vllngCuadre + 1
            End If
            '- Revisar si la póliza está cancelada -'
            If fblnPolizaCancelada(CLng(grdPolizas1.TextMatrix(vllngCont, 6))) Then
                vllngCancelada = vllngCancelada + 1
            End If
            '- Revisar si la póliza está descuadrada -'
            If fblnPolizaDescuadrada(CLng(grdPolizas1.TextMatrix(vllngCont, 6))) Then
                vllngDescuadre = vllngDescuadre + 1
                llngPolizaDescuadrada = CLng(grdPolizas1.TextMatrix(vllngCont, 2))
                If vllngDescuadre = 1 Then
                    llngPolizasDescuadradas = llngPolizaDescuadrada
                Else
                    llngPolizasDescuadradas = llngPolizasDescuadradas & Chr(13) & llngPolizaDescuadrada
                End If
            End If
        End If
    Next vllngCont
    
    If vllngCuentas = 0 Then
        'No se ha seleccionado ninguna póliza para exportar.
        MsgBox SIHOMsg(1212), vbInformation, "Mensaje"
        Exit Sub
    End If
    
    '- Validaciones antes de la exportación de Microsip -'
    If lintInterfazPolizas = 2 Then
        ' Si solo hay una póliza a exportar y ésta contiene cuenta de cuadre
        If vllngCuentas = 1 And vllngCuadre = 1 Then
            'No se puede realizar la exportación, la póliza contiene una cuenta de cuadre.
            MsgBox SIHOMsg(1208), vbInformation, "Mensaje"
            Exit Sub
        End If
        ' Si solo hay una póliza a exportar y ésta está cancelada
        If vllngCuentas = 1 And vllngCancelada = 1 Then
            'No se puede realizar la exportación, la póliza está cancelada.
            MsgBox SIHOMsg(1211), vbInformation, "Mensaje"
            Exit Sub
        End If
        ' Si solo hay una póliza a exportar y ésta está cancelada
        If vllngCuentas = 1 And vllngDescuadre = 1 Then
            'No se puede realizar la exportación, la póliza está descuadrada.
            MsgBox SIHOMsg(1394), vbInformation, "Mensaje"
            Exit Sub
        ElseIf vllngCuentas > 1 And vllngDescuadre = 1 Then
            'La póliza número n está descuadrada.
            Call MsgBox("La póliza número " & llngPolizaDescuadrada & " está descuadrada.", vbInformation, "Mensaje")
            Exit Sub
        ElseIf vllngCuentas > 1 And vllngDescuadre > 1 Then
            'No se puede realizar la exportación, existen n pólizas sin cuadrar.
            Call MsgBox("No se puede realizar la exportación, las siguientes pólizas están descuadradas: " & Chr(13) & llngPolizasDescuadradas & "", vbInformation, "Mensaje")
            Exit Sub
        End If
    End If

    CDgArchivo.CancelError = True
    CDgArchivo.InitDir = App.Path
    CDgArchivo.Flags = cdlOFNOverwritePrompt
    CDgArchivo.ShowSave
    
    If lintInterfazPolizas = CONTPAQ Then
        pExportaPoliza  ' Formato de la interfaz de Contpaq
    Else
        pExportaPolizaMicrosip  ' Formato de la interfaz de Microsip
    End If
  
NotificaError:
    err.Clear
End Sub
        
Private Sub pExportaPoliza()
    Dim X As Long
    Dim rsPoliza As New ADODB.Recordset
    Dim vlstrx As String, vlstrCadena As String
    Dim vllngPoliza As Long
    
    vlstrx = " SELECT dp.INTNUMEROPOLIZA,dp.BITNATURALEZAMOVIMIENTO, dp.MNYCANTIDADMOVIMIENTO,"
    vlstrx = vlstrx & "   p.DTMFECHAPOLIZA, p.CHRTIPOPOLIZA, p.INTCLAVEPOLIZA, p.VCHCONCEPTOPOLIZA, c.VCHCUENTACONTABLE"
    vlstrx = vlstrx & "  FROM CnDetallePoliza dp, CnPoliza p, CnCuenta c"
    vlstrx = vlstrx & "  WHERE dp.INTNUMEROPOLIZA=p.INTNUMEROPOLIZA AND dp.INTNUMEROCUENTA = c.INTNUMEROCUENTA"
    vlstrx = vlstrx & "   AND dp.INTNUMEROPOLIZA IN (-1"
    For X = 1 To grdPolizas1.Rows - 1
        If grdPolizas1.TextMatrix(X, 0) = "*" Then
            vlstrx = vlstrx & ", " & grdPolizas1.TextMatrix(X, 6)
        End If
    Next X
    vlstrx = vlstrx & ")"
    If Trim(cboTipo.List(cboTipo.ListIndex)) = "Diario" Then
        vlstrx = vlstrx & "   AND p.CHRTIPOPOLIZA='D'"
    End If
    If Trim(cboTipo.List(cboTipo.ListIndex)) = "Ingreso" Then
        vlstrx = vlstrx & "   AND p.CHRTIPOPOLIZA='I'"
    End If
    If Trim(cboTipo.List(cboTipo.ListIndex)) = "Egreso" Then
        vlstrx = vlstrx & "   AND p.CHRTIPOPOLIZA='E'"
    End If
    vlstrx = vlstrx & IIf(cboDepartamento.ListIndex > 0, " AND p.SMICVEDEPARTAMENTO = " & cboDepartamento.ItemData(cboDepartamento.ListIndex), "")
    vlstrx = vlstrx & "  ORDER BY dp.INTNUMEROPOLIZA, dp.BITNATURALEZAMOVIMIENTO DESC, c.VCHCUENTACONTABLE"
    Set rsPoliza = frsRegresaRs(vlstrx)
    With rsPoliza
        If .State <> adStateClosed Then
            If .RecordCount = 0 Then
                ' No existe información
                MsgBox SIHOMsg(12), vbOKOnly + vbExclamation, "Exportación de poliza(s)"
            Else
                Open CDgArchivo.FileName For Output As #1  ' Open file for output.
                .MoveFirst
                vllngPoliza = -1
                While Not .EOF
                    If vllngPoliza <> !intNumeroPoliza Then
                        vllngPoliza = !intNumeroPoliza
                        vlstrCadena = "P" ' Identificación
                        vlstrCadena = vlstrCadena & " " & Format(!dtmFechaPoliza, "YYYYMMDD") ' Fecha de alta de la póliza
                        vlstrCadena = vlstrCadena & " " & IIf(!chrTipoPoliza = "I", "1", IIf(!chrTipoPoliza = "E", "2", IIf(!chrTipoPoliza = "D", "3", "0"))) ' Tipo de Póliza
                        vlstrCadena = vlstrCadena & " " & Format(!intClavePoliza, "00000000") 'Número de Póliza
                        vlstrCadena = vlstrCadena & " " & "1" ' Clase de Póliza
                        vlstrCadena = vlstrCadena & " " & "000" ' Diario Especial de Agrupación
                        vlstrCadena = vlstrCadena & " " & Mid(!vchConceptoPoliza & String(100, " "), 1, 100) ' Concepto de la póliza
                        vlstrCadena = vlstrCadena & " " & "01" 'Sistema Origen 1 : ContPAQ 98
                        vlstrCadena = vlstrCadena & " " & "2" ' No Póliza Impresa
                        vlstrCadena = vlstrCadena & " " ' Espacio Final
                        Print #1, vlstrCadena ' Maestro o Encabezado de la Póliza
                    End If
                    vlstrCadena = "M" ' Identificación
                    vlstrCadena = vlstrCadena & " " & Mid(Replace(!VchCuentaContable, ".", "") & String(20, " "), 1, 20) ' Cuenta Contable a la que afecta
                    vlstrCadena = vlstrCadena & " " & String(10, " ") ' Referencia del Movimiento
                    vlstrCadena = vlstrCadena & " " & IIf(!bitNaturalezaMovimiento = 0, 2, 1) ' Tipo de Movimiento 1 Cargo, 2 Abono
                    vlstrCadena = vlstrCadena & " " & String(16 - Len(Format(!mnyCantidadMovimiento, "###0.00")), " ") & Format(!mnyCantidadMovimiento, "###0.00") 'Importe del movimiento
                    vlstrCadena = vlstrCadena & " " & "000" ' Diario Especial del Movimiento
                    vlstrCadena = vlstrCadena & " " & String(16 - Len(Format(0, "###0.00")), " ") & Format(0, "###0.00") 'Importe del movimiento en moneda extranjera
                    vlstrCadena = vlstrCadena & " " & String(30, " ") ' Concepto del Movimiento
                    vlstrCadena = vlstrCadena & " " ' Espacio Final
                    Print #1, vlstrCadena ' Detalle o Movimientos de la Póliza
                    .MoveNext
                Wend
                Close #1 ' Close file.
                
                '¡Los datos han sido guardados satisfactoriamente!
                MsgBox SIHOMsg(358), vbOKOnly + vbInformation, "Exportación de póliza(s)"
            End If
            .Close
        End If
    End With
End Sub

Private Sub cmdImportacion_Click()
On Error GoTo NotificaError

'    pImportaPoliza

    If lintInterfazPolizas = CONTPAQ Then
        pImportaPoliza  'Importar pólizas con formato de CONTPAQ
    ElseIf lintInterfazPolizas = MICROSIP Then
        pImportaPolizaMicrosip  'Importar pólizas con formato de Microsip
    Else
        pImportaPolizaApsi  'Importar con formato de APSI
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":cmdImportacion_Click"))
End Sub

Private Sub cmdImprimirPoliza_Click()
On Error GoTo NotificaError
    
    Dim vlintContador As Long
    Dim vllngFuncion As Long
    Dim vlblnHaySeleccionadas As Boolean
    
    If chkConcentrado.Value Then
        pEjecutaSentencia "DELETE FROM CNTMPRPTPOLIZASCONCENTRADO"
        
        vlblnHaySeleccionadas = False
        For vlintContador = 1 To grdPolizas1.Rows - 1
            If grdPolizas1.TextMatrix(vlintContador, 0) = "*" Then
                vlblnHaySeleccionadas = True
                vllngFuncion = 1
                frsEjecuta_SP Val(Trim(str(grdPolizas1.TextMatrix(vlintContador, 6)))), "FN_CNINSTMPPOLIZASCONCENTRADO", True, vllngFuncion
            End If
        Next vlintContador
        
        If vlblnHaySeleccionadas Then
            pImprimePoliza3 "I"
        End If
    Else
        For vlintContador = 1 To grdPolizas1.Rows - 1
            If grdPolizas1.TextMatrix(vlintContador, 0) = "*" Then
                'pImprimePoliza Str(grdPolizas1.TextMatrix(vlintContador, 6)), "I"
                pImprimePoliza2 str(grdPolizas1.TextMatrix(vlintContador, 6)), "I", 1 'Modificado para caso 4593
            End If
        Next vlintContador
    End If

Exit Sub
NotificaError:
    Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":cmdImprimirPoliza_Click"))
End Sub

Private Sub cmdInvertir_Click()
    Dim intIndex As Long
    
    If grdPolizas1.Rows > 0 Then
        lblnSel = False
        For intIndex = 1 To grdPolizas1.Rows - 1
            grdPolizas1.TextMatrix(intIndex, 0) = IIf(grdPolizas1.TextMatrix(intIndex, 0) = "*", "", "*")
            If grdPolizas1.TextMatrix(intIndex, 0) = "*" Then
                lblnSel = True
            End If
        Next
    End If
End Sub

Private Sub cmdPreview_Click()
On Error GoTo NotificaError

    Dim vlintContador As Long
    Dim vllngFuncion As Long
    Dim vlblnHaySeleccionadas As Boolean

    If chkConcentrado.Value Then
        pEjecutaSentencia "DELETE FROM CNTMPRPTPOLIZASCONCENTRADO"
        
        vlblnHaySeleccionadas = False
        For vlintContador = 1 To grdPolizas1.Rows - 1
            If grdPolizas1.TextMatrix(vlintContador, 0) = "*" Then
                vlblnHaySeleccionadas = True
                vllngFuncion = 1
                frsEjecuta_SP Val(Trim(str(grdPolizas1.TextMatrix(vlintContador, 6)))), "FN_CNINSTMPPOLIZASCONCENTRADO", True, vllngFuncion
            End If
        Next vlintContador
        
        If vlblnHaySeleccionadas Then
            pImprimePoliza3 "P"
        End If
    Else
        For vlintContador = 1 To grdPolizas1.Rows - 1
            If grdPolizas1.TextMatrix(vlintContador, 0) = "*" Then
                'pImprimePoliza Str(grdPolizas1.TextMatrix(vlintContador, 6)), "P"
                pImprimePoliza2 str(grdPolizas1.TextMatrix(vlintContador, 6)), "P", 1 'Modificado para caso 4593
            End If
        Next vlintContador
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":cmdPreview_Click"))
End Sub

Public Sub pImprimePoliza3(vlstrDestino As String)
'Procedimiento para imprimir las polizas seleccionadas en formato concentrado

    Dim vgrptReporte As CRAXDRT.Report
    Dim rsReporte As New ADODB.Recordset
    Dim alstrParametros(5) As String

    pInstanciaReporte vgrptReporte, "rptPolizaConcentrado.rpt"
    vgrptReporte.DiscardSavedData
    
    Set rsReporte = frsEjecuta_SP("", "SP_CNRPTPOLIZACONCENTRADO")
    If rsReporte.RecordCount > 0 Then
        alstrParametros(0) = "NombreHospital;" & Trim(vgstrNombreHospitalCH)
        alstrParametros(1) = "Departamento;" & cboDepartamento.Text
        alstrParametros(2) = "Del;" & Format(IIf(vlblnPolizasFueraCorte, vldtmFechaIni, vldtmFechaInicial), "DD/MMM/YYYY")
        alstrParametros(3) = "Al;" & Format(IIf(vlblnPolizasFueraCorte, vldtmFechaFin, vldtmFechaFinal), "DD/MMM/YYYY")
        alstrParametros(4) = "Tipo;" & StrConv(cboTipo.Text, vbUpperCase)
        pCargaParameterFields alstrParametros, vgrptReporte
        pImprimeReporte vgrptReporte, rsReporte, vlstrDestino, "Póliza"
    Else
        'No existe información con esos parámetro
        MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje"
    End If
    rsReporte.Close
    
    pEjecutaSentencia "DELETE FROM CNTMPRPTPOLIZASCONCENTRADO"
    
    frsEjecuta_SP 1 & "|" & Me.Name & "|" & chkConcentrado.Name & "|Value|" & vglngNumeroLogin & "|" & Trim(str(chkConcentrado.Value)), "SP_GNSELULTIMACONFIGURACION", True
End Sub

Private Sub cmdSeleccionar_Click()
    Dim intIndex As Long
    
    If grdPolizas1.Rows > 0 Then
        lblnSel = Not lblnSel
        For intIndex = 1 To grdPolizas1.Rows - 1
            grdPolizas1.TextMatrix(intIndex, 0) = IIf(lblnSel, "*", "")
        Next
    End If
End Sub

Private Sub pCargaPolizas()
On Error GoTo NotificaError

    Dim rsPolizas As New ADODB.Recordset
    Dim intCol As Integer
    
    lblnSel = False
    
    grdPolizas1.Rows = 0
    grdDetallePoliza1.Rows = 0
    
    cmdImprimirPoliza.Enabled = False
    cmdExportar.Enabled = False
    cmdPreview.Enabled = False
    chkConcentrado.Enabled = False
        
    If cboDepartamento.ListIndex < 0 Then Exit Sub
    
    vgstrParametrosSP = str(vgintClaveEmpresaContable) & _
                        "|" & cboDepartamento.ItemData(cboDepartamento.ListIndex) & _
                        "|" & IIf(cboTipo.ItemData(cboTipo.ListIndex) = 0, "*", IIf(cboTipo.ItemData(cboTipo.ListIndex) = 1, "D", IIf(cboTipo.ItemData(cboTipo.ListIndex) = 2, "I", IIf(cboTipo.ItemData(cboTipo.ListIndex) = 3, "E", "O")))) & _
                        "|" & mskFechaInicio.Text & _
                        "|" & mskFechaFin.Text & _
                        "|" & IIf(txtFolioIni.Text = "", -1, txtFolioIni.Text) & _
                        "|" & IIf(txtFolioFin.Text = "", -1, txtFolioFin.Text)
    Set rsPolizas = frsEjecuta_SP(vgstrParametrosSP, "SP_GNSELPOLIZAS")
    
    If Not rsPolizas.EOF Then
        With grdPolizas1
            .Rows = 1
            .FixedRows = 1
            .TextMatrix(0, 1) = "Fecha"
            .TextMatrix(0, 2) = "Número"
            .TextMatrix(0, 3) = "Concepto"
            .TextMatrix(0, 4) = "Tipo"
            .TextMatrix(0, 5) = "Departamento"
            .ColWidth(0) = 180
            .ColWidth(1) = 1100
            .ColWidth(2) = 800
            .ColWidth(3) = 6000
            .ColWidth(4) = 800
            .ColWidth(5) = 1000
            .ColWidth(6) = 0
        End With
    
        Do Until rsPolizas.EOF
            grdPolizas1.AddItem ""
            For intCol = 0 To rsPolizas.Fields.Count - 1
                grdPolizas1.TextMatrix(grdPolizas1.Rows - 1, intCol + 1) = rsPolizas.Fields(intCol).Value
            Next
            grdPolizas1.TextMatrix(grdPolizas1.Rows - 1, 1) = Format(rsPolizas.Fields(0).Value, "dd/mmm/yyyy")
            rsPolizas.MoveNext
        Loop
        
        chkConcentrado.Enabled = True
        cmdImprimirPoliza.Enabled = True
        If cboInterfazPolizas.ListIndex <> -1 Then
            If fblnValidarLicenciaPolizas(vgintClaveEmpresaContable, cboInterfazPolizas.ItemData(cboInterfazPolizas.ListIndex), cboInterfazPolizas.Text) Then
                cmdExportar.Enabled = IIf(chkConcentrado.Value, False, IIf(grdPolizas1.Rows > 0, True, False))
            Else
                cmdExportar.Enabled = False
            End If
        Else
            cmdExportar.Enabled = False
        End If
        lblnExistePolizas = True
        cmdPreview.Enabled = True
        'Posicionar en la primer poliza
        grdPolizas1.Row = 1
        'Desplegar detallado de la primer poliza
        pCargaDetallePoliza
        lblnExistePolizas = True
        
        vldtmFechaInicial = mskFechaInicio
        vldtmFechaFinal = mskFechaFin
    Else
        lblnExistePolizas = False
        vldtmFechaInicial = fdtmServerFecha
        vldtmFechaFinal = fdtmServerFecha
    End If
    rsPolizas.Close
    
Exit Sub
NotificaError:
    Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":pCargaPolizas"))
End Sub

Private Sub Command2_Click()
End Sub

Private Sub Folio_Click()

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    
    If KeyAscii = 27 Then Unload Me

Exit Sub
NotificaError:
    Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyPress"))
End Sub

Private Sub Form_Load()
On Error GoTo NotificaError

    Dim rsNombredelDepartamento As New ADODB.Recordset
    Dim rsValor As New ADODB.Recordset
    Dim lstrInterfaz As String
    
    vgstrNombreForm = Me.Name
    
    Dim strtmp As Boolean
    
    Me.Icon = frmMenuPrincipal.Icon
    mskFechaInicio.Text = fdtmServerFecha
    mskFechaFin.Text = fdtmServerFecha
    
    Set rsValor = frsEjecuta_SP(0 & "|" & Me.Name & "|" & chkConcentrado.Name & "|Value|" & vglngNumeroLogin & "|" & Trim(str(chkConcentrado.Value)), "SP_GNSELULTIMACONFIGURACION")
    If rsValor.RecordCount <> 0 Then
        chkConcentrado.Value = IIf(Trim(rsValor!vchvalor) = "0", 0, 1)
    Else
        chkConcentrado.Value = 0
    End If
    rsValor.Close
    
    lblnEjecutarBusqueda = False
    
    cboDepartamento.Enabled = False
    If fblnRevisaPermiso(vglngNumeroLogin, IIf(cgstrModulo = "IV", 3017, IIf(cgstrModulo = "PV", 3018, IIf(cgstrModulo = "CN", 3014, IIf(cgstrModulo = "CC", 3019, IIf(cgstrModulo = "CP", 3015, IIf(cgstrModulo = "NO", 3016, 0)))))), "L", True) Or fblnRevisaPermiso(vglngNumeroLogin, IIf(cgstrModulo = "IV", 3017, IIf(cgstrModulo = "PV", 3018, IIf(cgstrModulo = "CN", 3014, IIf(cgstrModulo = "CC", 3019, IIf(cgstrModulo = "CP", 3015, IIf(cgstrModulo = "NO", 3016, 0)))))), "E", True) Then
        cboDepartamento.Enabled = True
    End If
    cboDepartamento.Clear
    
'    vlstrsql = "SELECT smiCveDepartamento, vchDescripcion FROM NoDepartamento WHERE TNYCLAVEEMPRESA = " & vgintClaveEmpresaContable
    vlstrsql = "SELECT smicvedepartamento, vchdescripcion FROM NODEPARTAMENTO WHERE tnyclaveempresa = " & vgintClaveEmpresaContable & " AND smicvedepartamento IN (SELECT DISTINCT(smicvedepartamento) deptos FROM CNPOLIZA WHERE tnyclaveempresa = " & vgintClaveEmpresaContable & ")"
    Set rsNombredelDepartamento = frsRegresaRs(vlstrsql)
    
    If rsNombredelDepartamento.RecordCount > 0 Then
        Call pLlenarCboRs(cboDepartamento, rsNombredelDepartamento, 0, 1, 3)
        cboDepartamento.ListIndex = fintLocalizaCbo(cboDepartamento, CStr(vgintNumeroDepartamento)) 'se posiciona en el depto con el que se dio login
    End If
    
    cboTipo.ListIndex = 0
    pInstanciaReporte vgrptReporte, "rptPoliza.rpt"
    lblnSel = False
    
    '|  Solamente desde el módulo de contabilidad se podrán importar pólizas
    If cgstrModulo = "CN" Then
        Frame2.width = 3945
        cmdImportacion.Visible = True
    Else
        Frame2.width = 2550
        cmdImportacion.Visible = False
    End If
    
    lblnEjecutarBusqueda = True
    
    '- Caso 7323: Revisar si el usuario seleccionó la opción de mostrar la cuenta inmediata mayor -'
    lintProceso = 1 'Por defecto el proceso es de la Consulta de pólizas
    Set rsCuentaMayor = frsEjecuta_SP(CStr(vglngNumeroLogin) & "|" & lintProceso, "Sp_CpSelProcesoCuentaMayor")
    If rsCuentaMayor.RecordCount > 0 Then
        chkCuentaMayor.Value = rsCuentaMayor!Cuenta
    Else
        chkCuentaMayor.Value = 1
    End If
    rsCuentaMayor.Close
    
    '- Caso 7442: Mostrar las pólizas fuera del corte'
    If vlblnPolizasFueraCorte Then
        Me.Caption = "Pólizas fuera del corte"
        cmdImportacion.Visible = False  'En consulta de pólizas fuera de corte no se pueden importar pólizas
        pCargaPolizasFueraCorte
    Else
        Me.Caption = "Pólizas"
        pCargaPolizas
    End If
    
    pPosicionaObjetos 'Agregado caso 7442
    
    '- Caso 7762: Validar licenciamiento para la importación/exportación de pólizas -'
    lintInterfazPolizas = fintTipoInterfazPoliza(vgintClaveEmpresaContable)
    Select Case lintInterfazPolizas
        Case CONTPAQ: lstrInterfaz = "CONTPAQ"
        Case MICROSIP: lstrInterfaz = "Microsip"
        Case APSI: lstrInterfaz = "APSI"
    End Select
    If fblnValidarLicenciaPolizas(vgintClaveEmpresaContable, lintInterfazPolizas, lstrInterfaz) Then
        cmdImportacion.Enabled = True
        If cmdExportar.Visible Then cmdExportar.Enabled = True
    Else
        cmdImportacion.Enabled = False
        cmdExportar.Enabled = False
        lblInterfaz.Visible = False
        cboInterfazPolizas.Visible = False
    End If
    pAgregarInterfacesPolizas
    '---------------------------------------------------------------------'
    txtFolioIni.Enabled = False
    txtFolioFin.Enabled = False
    lblFolio.Enabled = False
    lblAl.Enabled = False
Exit Sub
NotificaError:
    Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '- Actualizar el proceso de la opción "Incluir cuenta inmediata mayor" -'
    pActualizaProcesoCuenta
    
    vlblnPolizasFueraCorte = False
    cboDepartamento.Enabled = True
End Sub

Private Sub grdPolizas1_AfterSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long)
    If grdPolizas1.Row > 0 Then
        pCargaDetallePoliza
    End If
End Sub

Private Sub grdPolizas1_DblClick()
    If grdPolizas1.Row > 0 Then
        grdPolizas1.TextMatrix(grdPolizas1.Row, 0) = IIf(grdPolizas1.TextMatrix(grdPolizas1.Row, 0) = "*", "", "*")
    End If
End Sub

Private Sub grdPolizas1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub mskFechaFin_GotFocus()
On Error GoTo NotificaError
    
    pSelMkTexto mskFechaFin

Exit Sub
NotificaError:
    Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":mskFechaFin_GotFocus"))
End Sub

Private Sub mskFechaFin_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        
        'Si existen polizas en ese rango de fechas dar foco al grid de lo contrario dar foco a la fecha de inicio
        If lblnExistePolizas = True Then
            grdPolizas1.SetFocus

        Else
            mskFechaInicio.SetFocus
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":mskFechaFin_KeyPress"))
End Sub

Private Sub mskFechaFin_LostFocus()
On Error GoTo NotificaError
    
    If mskFechaFin.ClipText = "" Then
        mskFechaFin.Text = fdtmServerFecha
    End If
    
    If Not fblnValidaFecha(mskFechaFin) Then
        mskFechaFin.Text = fdtmServerFecha
    End If
    
    If CDate(mskFechaInicio.Text) > CDate(mskFechaFin.Text) Then
        '¡Rango de fechas no válido!
        MsgBox SIHOMsg(64), vbOKOnly + vbInformation, "Mensaje"
        mskFechaInicio.SetFocus
    Else
        pCargaPolizas
    End If

Exit Sub
NotificaError:
    Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":mskFechaFin_LostFocus"))
End Sub

Private Sub pCargaDetallePoliza()
On Error GoTo NotificaError

    Dim rsDetallePoliza As New ADODB.Recordset
    Dim rsNumeroDepartamentoPoliza As New ADODB.Recordset
    Dim X As Long
    Dim intCol As Integer
    
    grdDetallePoliza1.Rows = 0
    
    If grdPolizas1.Row > 0 Then
        vgstrParametrosSP = CStr(grdPolizas1.TextMatrix(grdPolizas1.Row, 6))
        Set rsDetallePoliza = frsEjecuta_SP(vgstrParametrosSP, "Sp_CpSelDetallePoliza")
        If Not rsDetallePoliza.EOF Then
            With grdDetallePoliza1
                .Rows = 1
                .FixedRows = 1
                .TextMatrix(0, 1) = "Cuenta"
                .TextMatrix(0, 2) = "Descripción"
                .TextMatrix(0, 3) = "Cargo"
                .TextMatrix(0, 4) = "Abono"
                .ColWidth(0) = 100
                .ColWidth(1) = 2000
                .ColWidth(2) = 5700
                .ColWidth(3) = 1500
                .ColWidth(4) = 1500
            End With
            
            Do Until rsDetallePoliza.EOF
                grdDetallePoliza1.AddItem ""
                For intCol = 0 To rsDetallePoliza.Fields.Count - 3
                    If intCol = 2 Or intCol = 3 Then
                        grdDetallePoliza1.TextMatrix(grdDetallePoliza1.Rows - 1, intCol + 1) = FormatCurrency(rsDetallePoliza.Fields(intCol).Value, 2) 'Format(rsDetallePoliza.Fields(intCol).Value, "#,##0.00")
                    Else
                        grdDetallePoliza1.TextMatrix(grdDetallePoliza1.Rows - 1, intCol + 1) = rsDetallePoliza.Fields(intCol).Value
                    End If
                Next
                rsDetallePoliza.MoveNext
            Loop
        End If
        rsDetallePoliza.Close
    End If

Exit Sub
NotificaError:
    Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":pCargaDetallePoliza"))
End Sub

Private Sub mskFechaInicio_GotFocus()
On Error GoTo NotificaError
    
    pSelMkTexto mskFechaInicio

Exit Sub
NotificaError:
    Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":mskFechaInicio_GotFocus"))
End Sub

Private Sub mskFechaInicio_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        mskFechaFin.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":mskFechaInicio_KeyPress"))
End Sub

Private Sub mskFechaInicio_LostFocus()
On Error GoTo NotificaError
    
    If mskFechaInicio.ClipText = "" Then
        mskFechaInicio.Text = fdtmServerFecha
    End If
    
    If Not fblnValidaFecha(mskFechaInicio) Then
        mskFechaInicio.Text = fdtmServerFecha
    End If

Exit Sub
NotificaError:
    Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":mskFechaInicio_LostFocus"))
End Sub

'|  Importa una póliza del sistema ContPaq
Public Sub pImportaPoliza()
    Dim fs As FileSystemObject
    Dim f As Variant
    Dim strLinea As String
    Dim rsPoliza As New ADODB.Recordset
    Dim rsDetallePoliza As New ADODB.Recordset
    Dim rsPolizasImportadas As New ADODB.Recordset
    Dim strArchivoImportacion As String         '|  Archivo que se intenta importar
    Dim strSentencia As String                  '|  Variable genérica para armar instrucciones SQL
    Dim lngCvePolizaNueva As Long               '|  Clave de la póliza que se desea importar
    Dim lngCveCuentaNueva As Long               '|  Clave de la cuenta del detalle de la póliza que se intenta importar
    Dim lngPersonaGraba As Long
    Dim lngResultado As Long
    Dim lintInterfazImporta As Integer
    Dim rsUuid As New ADODB.Recordset
    Dim xlsApp As Object 'Excel.Application
    Dim hoja As Object 'Excel.Worksheet
    Dim lngContador As Long
    Dim rngfnd As Object 'Excel.Range
    
    Const xlValues = -4163
    Const xlWhole = 1
    Const xlByRows = 1
    Const xlNext = 1
    
On Error GoTo NotificaError
    
    '-----------------------------------------------------------------------------
    '|  Validaciones iniciales que pueden suspender la ejecución de la rutina
    '-----------------------------------------------------------------------------
    '|  Selecciona el archivo que se desea importar
    strArchivoImportacion = fstrAbreArchivo
    '|  Si se canceló el diálogo de abrir
    If strArchivoImportacion = "" Then Exit Sub
    '|  Valida que el formato del archivo de entrada sea del tipo esperado (CONTPAQ),
    '|  que las pólizas que se intentan importar pertenezcan a un periodo abierto y
    '|  que las pólizas no hayan sido importadas anteriormente
    '|  si se obtendra el uuid relacionado con la poliza por medio de otro archivo
        
    If Not fblnValidacionesGenerales(strArchivoImportacion) Then Exit Sub
    

    '|  Pide autorización para hacer la importación de la póliza
    lngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
    If lngPersonaGraba = 0 Then Exit Sub
    lstrArchivoUuid = strArchivoImportacion
    
    '------------------------------------------------
    '|  Inicialización de variables
    '------------------------------------------------
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    Set f = fs.OpenTextFile(strArchivoImportacion, ForReading, , TristateUseDefault)
    strSentencia = "SELECT bitasentada, chrtipopoliza, dtmfechapoliza, intclavepoliza, intcveempleado, intnumeropoliza, smicvedepartamento, smiejercicio, tnyclaveempresa, tnymes, vchconceptopoliza, vchnumero " & _
                   "FROM   CNPOLIZA " & _
                   "WHERE  intnumeropoliza = -1"
   Set rsPoliza = frsRegresaRs(strSentencia, adLockOptimistic, adOpenDynamic)
    
    strSentencia = "SELECT bitnaturalezamovimiento, intnumerocuenta, intnumeropoliza, intnumeroregistro, mnycantidadmovimiento, vchconcepto, vchreferencia " & _
                   "FROM   CNDETALLEPOLIZA " & _
                   "WHERE  intnumeropoliza = -1"
    Set rsDetallePoliza = frsRegresaRs(strSentencia, adLockOptimistic, adOpenDynamic)
    
    strSentencia = "SELECT * " & _
                   "FROM   CNPOLIZASIMPORTADAS " & _
                   "WHERE  CNPOLIZASIMPORTADAS.chrtipopoliza = '*' "
    Set rsPolizasImportadas = frsRegresaRs(strSentencia, adLockOptimistic, adOpenDynamic)
    
    strSentencia = "SELECT * " & _
                   "FROM   CNCFDIPOLIZA " & _
                   "WHERE  CNCFDIPOLIZA.intIdComprobante = -1 "
    Set rsUuid = frsRegresaRs(strSentencia, adLockOptimistic, adOpenDynamic)
       
    Me.MousePointer = 11 '|  Reloj de arena
    
    '|  Clave de la interfaz seleccionada para la importación
    If cboInterfazPolizas.ListIndex <> -1 Then
        lintInterfazImporta = cboInterfazPolizas.ItemData(cboInterfazPolizas.ListIndex)
    Else
        lintInterfazImporta = lintInterfazPolizas
    End If
    
    EntornoSIHO.ConeccionSIHO.BeginTrans
    
    '|  Bloquea el cierre
    lngResultado = 1
    vgstrParametrosSP = vgintClaveEmpresaContable & "|" & "Grabando poliza"
    frsEjecuta_SP vgstrParametrosSP, "sp_CnUpdEstatusCierre", True, lngResultado
    If lngResultado = 1 Then
        If lstrArchivoUuid <> "" Then
            Set xlsApp = CreateObject("Excel.Application")
            DoEvents
            xlsApp.Workbooks.Open lstrArchivoUuid
            DoEvents
            Set hoja = xlsApp.Worksheets(1)
            DoEvents
        End If
        Do Until f.AtEndOfStream
            DoEvents
            strLinea = f.ReadLine
            '|  Indica el tipo de registro P = Póliza (Maestro)
            If Mid(strLinea, 1, 1) = "P" Then
                '--------------------------------------------
                '|  Graba el maestro de la póliza
                '---------------------------------------------
                With rsPoliza
                    .AddNew
                    If lblnArchivoUuid Then
                        !dtmFechaPoliza = Mid(strLinea, 4, 4) & "/" & Mid(strLinea, 8, 2) & "/" & Mid(strLinea, 10, 2)
                        !chrTipoPoliza = IIf(Trim(Mid(strLinea, 13, 4)) = "1", "I", IIf(Trim(Mid(strLinea, 13, 4)) = "2", "E", IIf(Trim(Mid(strLinea, 13, 4)) = "3", "D", "")))
                        !vchConceptoPoliza = Trim(Mid(strLinea, 41, 100))
                    Else
                        '|  Fecha de alta de la póliza
                        !dtmFechaPoliza = Mid(strLinea, 3, 4) & "/" & Mid(strLinea, 7, 2) & "/" & Mid(strLinea, 9, 2)
                        '|  Tipo de póliza 1 Ingresos, 2 Egresos, 3 Diario, 4 De orden, 5 Estadísticas, 6 en adelante = creadas por el usuario
                        !chrTipoPoliza = IIf(Mid(strLinea, 12, 1) = "1", "I", IIf(Mid(strLinea, 12, 1) = "2", "E", IIf(Mid(strLinea, 12, 1) = "3", "D", "")))
                        '|  Concepto de la póliza
                        !vchConceptoPoliza = Mid(strLinea, 29, 100)
                    End If
                    '|  Indica si la póliza está incluida en un cierre contable ( guarda 1 = si está incluida, 0 = no está incluida)
                    !BITASENTADA = 0
                    '|  Empleado que esta realizando la importación
                    !intCveEmpleado = lngPersonaGraba
                    '|  Clave de la empresa contable (relacionado con CnEmpresaContable)
                    !TNYCLAVEEMPRESA = vgintClaveEmpresaContable
                    '|  Departamento del usuario que esta en el sistema
                    !smicvedepartamento = vgintNumeroDepartamento
                    '|  Año del ejercicio contable.
                    !SMIEJERCICIO = Format(!dtmFechaPoliza, "yyyy")
                    '|  Número de mes contable.
                    !tnyMes = Format(!dtmFechaPoliza, "mm")
                    '|  Número de póliza para manejo interno del área de contabilidad
                    !vchNumero = " "
                    '|  El número de póliza es generado por el sistema según este configurado
                    !intClavePoliza = flngFolioPoliza(vgintClaveEmpresaContable, !chrTipoPoliza, !SMIEJERCICIO, !tnyMes, False)
                    .Update
                    lngCvePolizaNueva = flngObtieneIdentity("SEC_CNPOLIZA", !intNumeroPoliza)
                End With
                '--------------------------------------------------------------------------------------
                '|  Registra las pólizas que se están importando si previamente no se han registrado
                '--------------------------------------------------------------------------------------
                    With rsPolizasImportadas
                        .AddNew
                        If lblnArchivoUuid Then
                            !chrTipoPoliza = IIf(Trim(Mid(strLinea, 13, 4)) = "1", "I", IIf(Trim(Mid(strLinea, 13, 4)) = "2", "E", IIf(Trim(Mid(strLinea, 13, 4)) = "3", "D", "")))
                            !dtmFechaPoliza = Mid(strLinea, 10, 2) & "/" & Mid(strLinea, 8, 2) & "/" & Mid(strLinea, 4, 4)
                        Else
                            !chrTipoPoliza = IIf(Mid(strLinea, 12, 1) = "1", "I", IIf(Mid(strLinea, 12, 1) = "2", "E", IIf(Mid(strLinea, 12, 1) = "3", "D", "")))
                            !dtmFechaPoliza = Mid(strLinea, 9, 2) & "/" & Mid(strLinea, 7, 2) & "/" & Mid(strLinea, 3, 4)
                        End If
                        '- Campos agregados para la relación de la póliza importada con la clave interna del SiHO -'
                        !intInterfazPoliza = lintInterfazImporta
                        !intNumeroPoliza = lngCvePolizaNueva
                        If lblnArchivoUuid Then
                            !intCvePolizaExterna = Val(Mid(strLinea, 18, 9))
                        Else
                            !intCvePolizaExterna = Val(Mid(strLinea, 14, 8))
76                        End If
77                        .Update
78                    End With
79            ElseIf Mid(strLinea, 1, 1) = "M" Then  '|  M = Movimiento (Detalle)
                '--------------------------------------------
                '|  Graba el detalle de la póliza
                '---------------------------------------------
80                With rsDetallePoliza
81                    .AddNew
82                    lngCveCuentaNueva = fintFormateaCuenta(Trim(IIf(lblnArchivoUuid, Mid(strLinea, 4, 30), Mid(strLinea, 3, 20))), vgintClaveEmpresaContable)
83                    If lngCveCuentaNueva = -1 Then
84                        EntornoSIHO.ConeccionSIHO.RollbackTrans
85                        Me.MousePointer = 0
86                        If lstrArchivoUuid <> "" Then
87                            xlsApp.Quit
88                            Set xlsApp = Nothing
89                        End If
90                        Exit Sub
91                    End If
92                    !intNumeroCuenta = lngCveCuentaNueva
93                    !intNumeroPoliza = lngCvePolizaNueva
94                    If lblnArchivoUuid Then
95                        !mnyCantidadMovimiento = CDbl(Mid(strLinea, 58, 20))
96                        !bitNaturalezaMovimiento = IIf(CInt(Mid(strLinea, 56, 1)) = 0, 1, 0)
97                        !vchReferencia = IIf(Trim(Mid(strLinea, 35, 20)) = "", " ", Trim(Mid(strLinea, 35, 20)))
98                        !vchConcepto = IIf(Trim(Mid(strLinea, 111, 100)) = "", " ", Trim(Mid(strLinea, 111, 100)))
99                    Else
100                        !mnyCantidadMovimiento = CDbl(Mid(strLinea, 37, 16))
101                        !bitNaturalezaMovimiento = IIf(CInt(Mid(strLinea, 35, 1)) = 2, 0, CInt(Mid(strLinea, 35, 1)))
102                        !vchReferencia = IIf(Trim(Mid(strLinea, 54, 20)) = "", " ", Trim(Mid(strLinea, 54, 20)))
103                        !vchConcepto = IIf(Trim(Mid(strLinea, 75, 100)) = "", " ", Trim(Mid(strLinea, 75, 100)))
104                     End If
105                    .Update
106                End With
107            ElseIf Mid(strLinea, 1, 1) = "A" Then '| uuid
108                With rsUuid
109                    .AddNew
110                    !INTNUMPOLIZA = lngCvePolizaNueva
111                    !intNumCuenta = lngCveCuentaNueva
112                    !VCHUUID = Trim(Mid(strLinea, 3))

113                    Set rngfnd = hoja.Cells.Find(What:=Trim(Mid(strLinea, 3)), After:=hoja.Cells(1, 1), LookIn:= _
                            xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:= _
                            xlNext, MatchCase:=False, SearchFormat:=False)

114                    If Not rngfnd Is Nothing Then
115                        !numTotalComprobante = hoja.Cells(rngfnd.Row, 17)
116                        !VCHNOMBRERECEPTOR = hoja.Cells(rngfnd.Row, 10)
117                        !VCHRFCRECEPTOR = hoja.Cells(rngfnd.Row, 11)
118                    End If
119                    .Update
120                End With
121            Else
122                MsgBox "La información contenida en el archivo no cumple con el formato requerido.", vbCritical, "Mensaje"
123                lngResultado = 0
124                Exit Do
125            End If
126       Loop
127        If lstrArchivoUuid <> "" Then
128            xlsApp.Quit
129            Set xlsApp = Nothing
130        End If

131        pEjecutaSentencia "update CnEstatusCierre set vchEstatus = 'Libre' where tnyClaveEmpresa = " & str(vgintClaveEmpresaContable)
                
132        Call pGuardarLogTransaccion(Me.Name, EnmBorrar, lngPersonaGraba, "IMPORTACION DE POLIZAS", strArchivoImportacion)
133    End If
    
134    EntornoSIHO.ConeccionSIHO.CommitTrans
    
135    f.Close
136    Me.MousePointer = 0
137    If lngResultado <> 0 Then
        '|  La póliza ha sido importada exitosamente
138        MsgBox SIHOMsg(744), vbOKOnly + vbInformation, "Mensaje"
139        pCargaPolizas
140    End If
    
Exit Sub
NotificaError:
    EntornoSIHO.ConeccionSIHO.RollbackTrans
    Me.MousePointer = 0
    Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":pImportaPoliza " & Erl()))
    err.Clear
    
    If lstrArchivoUuid <> "" Then
        xlsApp.Quit
        Set xlsApp = Nothing
    End If
End Sub

'|  Da formato a una cuenta según la mascara definida en la base de datos
Private Function fintFormateaCuenta(strCuenta As String, intEmpresa As Integer) As Long
    Dim strCuentaFormateada As String
    Dim strMascaraCuenta As String
    Dim intContMascara As Integer
    Dim strSentencia As String
    Dim rsCuenta As New ADODB.Recordset
    Dim intContCuenta As Integer
    Dim rs As ADODB.Recordset
    
On Error GoTo NotificaError

    strCuentaFormateada = ""
    Set rs = frsSelParametros("CN", intEmpresa, "VCHESTRUCTURACUENTACONTABLE")
    If Not rs.EOF Then
        If IsNull(rs!Valor) Then
            strMascaraCuenta = -1
        Else
            strMascaraCuenta = rs!Valor
        End If
    Else
        strMascaraCuenta = -1
    End If
    rs.Close
    
    '|  Si no se he configurado una mascara de cuenta
    If strMascaraCuenta = "-1" Then
        '|  No existe registrada la estructura de cuentas contables.
        MsgBox SIHOMsg(246), vbCritical, "Mensaje"
        fintFormateaCuenta = -1
        Exit Function
    End If
    
    '|  Si no coincide la cuenta de entrada con la configurada en parametros
    If Len(Replace(Trim(strMascaraCuenta), ".", "")) <> Len(Trim(strCuenta)) Then
        '|  ¡El formato de la cuenta es incorrecto!
        MsgBox SIHOMsg(66), vbCritical, "Mensaje"
        fintFormateaCuenta = -1
        Exit Function
    End If
    
    intContMascara = 1
    intContCuenta = 1
    While intContMascara <= Len(strMascaraCuenta)
        If Mid(strMascaraCuenta, intContMascara, 1) <> "." Then
            strCuentaFormateada = strCuentaFormateada & Mid(strCuenta, intContCuenta, 1)
            intContCuenta = intContCuenta + 1
        Else
            strCuentaFormateada = strCuentaFormateada & "."
        End If
        intContMascara = intContMascara + 1
    Wend
    
    strSentencia = "SELECT intNumeroCuenta " & _
                   "FROM   CnCuenta " & _
                   "WHERE  vchCuentaContable = '" & strCuentaFormateada & "' " & _
                   "  AND  tnyClaveEmpresa = " & vgintClaveEmpresaContable & _
                   "  AND  bitEstatusActiva = 1"
    Set rsCuenta = frsRegresaRs(strSentencia, adLockOptimistic, adOpenDynamic)
    If rsCuenta.RecordCount > 0 Then
        fintFormateaCuenta = rsCuenta!intNumeroCuenta
    Else
        '|  ¡La cuenta no existe!
        MsgBox SIHOMsg(67) & vbCrLf & vbCrLf & strCuentaFormateada, vbCritical, "Mensaje"
        fintFormateaCuenta = -1
    End If
    
Exit Function
NotificaError:
    Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":fintFormateaCuenta(" & strCuenta & ")"))
End Function

'|  Muestra el diálogo de abrir archivo y regresa la ruta del archivo seleccionado
Private Function fstrAbreArchivo() As String
On Error GoTo NotificaError
    
    fstrAbreArchivo = ""
    CDgArchivo.CancelError = True

    With CDgArchivo
        .DialogTitle = "Abrir archivo para importación"
        .InitDir = App.Path
        .Flags = cdlOFNHideReadOnly
        '|  Solo archivos de tipo texto
        .Filter = "Text Files(*.txt)|*.txt|"
        '|  Muestra el diálogo Abrir
        .ShowOpen
        fstrAbreArchivo = .FileName
    End With
    
Exit Function
NotificaError:
    '|  Se presionó el botón salir
End Function

'|  Muestra el diálogo de abrir archivo y regresa la ruta del archivo seleccionado
Private Function fstrAbreArchivoUuid() As String
On Error GoTo NotificaError
    
    fstrAbreArchivoUuid = ""
    CDgArchivo.CancelError = True

    With cdgExcel
        .DialogTitle = "Abrir archivo para importación de datos del comprobante"
        .InitDir = App.Path
        .Flags = cdlOFNHideReadOnly
        .Filter = "Excel (*.xls;*.xlsx)|*.xls;*.xlsx"
        .ShowOpen
        fstrAbreArchivoUuid = .FileName
    End With
    
Exit Function
NotificaError:
End Function
Private Function fblnEstructuraUUID(strArchivo As String) As Boolean
    Dim fs As FileSystemObject
    Dim f As Variant
    Dim strLinea As String
    
    fblnEstructuraUUID = False
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.OpenTextFile(strArchivo, ForReading, TristateFalse)
    'Se identifica si el archivo incluye los UUIDs, de ser asi, el formato es diferente
    Do Until f.AtEndOfStream
        DoEvents
        strLinea = f.ReadLine
        If Mid(strLinea, 1, 1) = "A" Then
            fblnEstructuraUUID = True
            Exit Do
        End If
    Loop
    f.Close
    
End Function

'|  Muestra el diálogo de abrir archivo y regresa la ruta del archivo seleccionado
Private Function fstrAbreArchivoImportaApsi() As String
On Error GoTo NotificaError
    
    fstrAbreArchivoImportaApsi = ""
    CDgArchivo.CancelError = False

    With CDgArchivo
        .DialogTitle = "Abrir archivo para importación"
        .InitDir = App.Path
        .Flags = cdlOFNHideReadOnly
        .FileName = ""
        '|  Solo archivos de tipo texto
        .Filter = "Excel (*.xls;*.xlsx)|*.xls;*.xlsx"
        '|  Muestra el diálogo Abrir
        .ShowOpen
        fstrAbreArchivoImportaApsi = .FileName
    End With
    
Exit Function
NotificaError:
    Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":fstrAbreArchivoImportaApsi()"))
End Function

'|  Verifica que el formato del archivo de texto sea el especificado por el layout de CONTPAQ,
'|  que las fechas de las polizas tengan cierre abierto y
'|  que las pólizas no hayan sido importadas anteriormente
Private Function fblnValidacionesGenerales(strArchivo As String) As Boolean
On Error GoTo NotificaError
    
    Dim fs As FileSystemObject
    Dim f As Variant
    Dim strLinea As String
    Dim strPeriodosCerrados As String                   '|  Lista de los periodos cerrados que se informará se deben abrir antes de continuar
    Dim strEjercicio As String
    Dim strMes As String
    Dim strFechaPoliza As String
    Dim strTipoPoliza As String
    Dim strClavePoliza As String
    Dim blnFormatoErroneo As String
    Dim strPolizasImportadasAnteriormente As String     '|  Lista de las pólizas que ya han sido importadas con anterioridad
    
    fblnValidacionesGenerales = False
    strPeriodosCerrados = ""
    strPolizasImportadasAnteriormente = ""
    blnFormatoErroneo = False
    lblnArchivoUuid = fblnEstructuraUUID(strArchivo) 'Formato de txt con uuid
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.OpenTextFile(strArchivo, ForReading, TristateFalse)
    
    '--------------------------------------------------------------------'
    '|  Recorre el archivo para validar los encabezados de las pólizas  |'
    '--------------------------------------------------------------------'
    Do Until f.AtEndOfStream
        DoEvents
        strLinea = f.ReadLine
        
        If Mid(strLinea, 1, 1) = "P" Then
            '---------------------------------------------------------------------------------------------'
            '|  Verifica que el encabezado del archivo coincida con el formato especificado por CONTPAQ  |'
            '---------------------------------------------------------------------------------------------'
            If lblnArchivoUuid Then
                If Mid(strLinea, 3, 1) <> " " _
                   Or Mid(strLinea, 12, 1) <> " " _
                   Or Mid(strLinea, 17, 1) <> " " _
                   Or Mid(strLinea, 27, 1) <> " " _
                   Or Mid(strLinea, 29, 1) <> " " _
                   Or Mid(strLinea, 40, 1) <> " " Then
                   blnFormatoErroneo = True
                   Exit Do
                End If
            Else
                If Mid(strLinea, 2, 1) <> " " _
                   Or Mid(strLinea, 11, 1) <> " " _
                   Or Mid(strLinea, 13, 1) <> " " _
                   Or Mid(strLinea, 22, 1) <> " " _
                   Or Mid(strLinea, 24, 1) <> " " _
                   Or Mid(strLinea, 28, 1) <> " " _
                   Or Mid(strLinea, 129, 1) <> " " _
                   Or Mid(strLinea, 132, 1) <> " " _
                   Or Mid(strLinea, 134, 1) <> " " Then
                   blnFormatoErroneo = True
                   Exit Do
                End If
            End If
            '-------------------------------------------------------------------------------------------------'
            '|  Verifica que el periodo contable de la(s) póliza(s) que se intenta(n) importar esté abierto  |'
            '|  de lo contrario lo guarda en una cadena para desplegarlo posteriormente                      |'
            '-------------------------------------------------------------------------------------------------'
            strEjercicio = Mid(strLinea, IIf(lblnArchivoUuid, 4, 3), 4)
            strMes = Mid(strLinea, IIf(lblnArchivoUuid, 8, 7), 2)
            If fblnPeriodoCerrado(vgintClaveEmpresaContable, CInt(strEjercicio), CInt(strMes)) Then '|  fblnPeriodoCerrado(CveEmpresa, Año, Mes)
                '|  Si no se ha agregado a la cadena strPeriodosCerrados
                If Len(Replace(strPeriodosCerrados, fstrMesLetras(CInt(Mid(strLinea, IIf(lblnArchivoUuid, 8, 7), 2))) & " " & strEjercicio, "")) = _
                    Len(strPeriodosCerrados) Then
                    strPeriodosCerrados = strPeriodosCerrados & vbCrLf & fstrMesLetras(CInt(Mid(strLinea, IIf(lblnArchivoUuid, 8, 7), 2))) & " " & strEjercicio
                End If
            End If
            
            '----------------------------------------------------------------------------------------------'
            '|  Verifica que las pólizas que se intentan importar no hayan sido importadas anteriormente  |'
            '|  de lo contrario las guarda en una cadena para desplegarlas posteriormente                 |'
            '----------------------------------------------------------------------------------------------'
            If lblnArchivoUuid Then
                strFechaPoliza = Mid(strLinea, 10, 2) & "/" & Mid(strLinea, 8, 2) & "/" & Mid(strLinea, 4, 4)
                strTipoPoliza = IIf(Trim(Mid(strLinea, 13, 4)) = "1", "I", IIf(Trim(Mid(strLinea, 13, 4)) = "2", "E", IIf(Trim(Mid(strLinea, 13, 4)) = "3", "D", "")))
                strClavePoliza = Val(Trim(Mid(strLinea, 18, 9)))
            Else
                strFechaPoliza = Mid(strLinea, 9, 2) & "/" & Mid(strLinea, 7, 2) & "/" & Mid(strLinea, 3, 4)
                strTipoPoliza = IIf(Mid(strLinea, 12, 1) = "1", "I", IIf(Mid(strLinea, 12, 1) = "2", "E", IIf(Mid(strLinea, 12, 1) = "3", "D", "")))
                strClavePoliza = Val(Mid(strLinea, 14, 8))
            End If
            If fblnPolizaImportada(strFechaPoliza, strTipoPoliza, strClavePoliza, CONTPAQ) Then
                '|  Si no se ha agregado a la cadena strPolizasImportadasAnteriormente
                If Len(Replace(strPolizasImportadasAnteriormente, strFechaPoliza & " " & IIf(strTipoPoliza = "I", "Ingreso", IIf(strTipoPoliza = "E", "Egreso", IIf(strTipoPoliza = "D", "Diario", "Desconocido"))), "")) = _
                    Len(strPolizasImportadasAnteriormente) Then
                    strPolizasImportadasAnteriormente = strPolizasImportadasAnteriormente & vbCrLf & strFechaPoliza & " " & IIf(strTipoPoliza = "I", "Ingreso", IIf(strTipoPoliza = "E", "Egreso", IIf(strTipoPoliza = "D", "Diario", "Desconocido")))
                End If
            End If
        End If
    Loop
    f.Close
    
    If strPeriodosCerrados <> "" Then
        '|  El periodo contable esta cerrado.
        MsgBox SIHOMsg(209) & vbCrLf & "No se puede continuar." & vbCrLf & strPeriodosCerrados, vbCritical, "Mensaje"
        Exit Function
    End If
    
    If strPolizasImportadasAnteriormente <> "" Then
        '|  Ya han sido importadas pólizas con las siguientes fechas y tipos. ¿Desea continuar?
        If MsgBox(SIHOMsg(745) & strPolizasImportadasAnteriormente, vbYesNo + vbCritical, "Mensaje") = vbNo Then
            Exit Function
        End If
    End If
    
    If blnFormatoErroneo Then
        '|  Formato interno de archivo erróneo.
        MsgBox SIHOMsg(746), vbCritical, "Mensaje"
        Exit Function
    End If
    
    lstrArchivoUuid = ""
    If lblnArchivoUuid Then
        lstrArchivoUuid = fstrAbreArchivoUuid()
        If lstrArchivoUuid = "" Then
            Exit Function
        End If
    End If

    fblnValidacionesGenerales = True
    
Exit Function
NotificaError:
    Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":fblnValidacionesGenerales(" & strArchivo & ")"))
End Function

'|  Convierte un número de mes a la palabra del mes correspondiente, si el mes no esta entre 1 y 12
'|  regresará una cadena vacía
'|     intMes.    Número del mes
'|     intEstilo. Indica el formato en que regresará la palabra la función, por defecto y número desconocido regresara con formato 1.
'|                  0 = Primera mayúscula las demás minúsculas
'|                  1 = Mayúsculas
'|                  2 = Minúsculas
Private Function fstrMesLetras(intMes As Integer, Optional intFormato As Integer) As String
    fstrMesLetras = ""
    Select Case intMes
        Case 1
            fstrMesLetras = "Enero"
        Case 2
            fstrMesLetras = "Febrero"
        Case 3
            fstrMesLetras = "Marzo"
        Case 4
            fstrMesLetras = "Abril"
        Case 5
            fstrMesLetras = "Mayo"
        Case 6
            fstrMesLetras = "Junio"
        Case 7
            fstrMesLetras = "Julio"
        Case 8
            fstrMesLetras = "Agosto"
        Case 9
            fstrMesLetras = "Septiembre"
        Case 10
            fstrMesLetras = "Octubre"
        Case 11
            fstrMesLetras = "Noviembre"
        Case 12
            fstrMesLetras = "Diciembre"
    End Select
    If intFormato = 1 Then fstrMesLetras = UCase(fstrMesLetras)
    If intFormato = 2 Then fstrMesLetras = LCase(fstrMesLetras)
End Function

'|  Verifica si ya se importó una póliza con la fecha y el tipo de póliza que se reciben como parámetros
Private Function fblnPolizaImportada(strFecha As String, strTipo As String, strClavePoliza As String, intInterfaz As Integer) As Boolean
    Dim strSentencia As String
On Error GoTo NotificaError
    
    fblnPolizaImportada = False
    strSentencia = "SELECT COUNT(*) Co " & _
                   "FROM   CNPOLIZASIMPORTADAS " & _
                   "WHERE  dtmFechaPoliza = " & fstrFechaSQL(strFecha) & _
                   "  AND  NVL(chrTipoPoliza, ' ') = '" & IIf(Trim(strTipo) = "", " ", strTipo) & "'" & _
                   "  AND  intInterfazPoliza = " & intInterfaz
    If Trim(strClavePoliza) <> "" Then strSentencia = strSentencia & "  AND  intCvePolizaExterna = " & strClavePoliza
    If frsRegresaRs(strSentencia, adLockOptimistic, adOpenDynamic)!co > 0 Then
        fblnPolizaImportada = True
    End If
    
Exit Function
NotificaError:
    Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":fblnPolizaImportada(" & strFecha & ", " & strTipo & ")"))
End Function

'(CR) - Caso 4593: Procedimiento modificado para mostrar o no las cuentas inmediatas mayores -'
Public Sub pImprimePoliza2(vlstrNumPoliza As String, vlstrDestino As String, Optional vlintDesglosadoCuenta As Integer)
'Procedimiento para imprimir una póliza
' vlstrNumPoliza => Consecutivo de la póliza (CnPoliza.intNumeroPoliza)
' vlstrDestino   => "P" = vista previa, "I" = impresora

    Dim vgrptReporte As CRAXDRT.Report
    Dim rsReporte As New ADODB.Recordset

    Dim alstrParametros(1) As String
    Dim vlstrNumUsuario As String
    Dim vlstrEmpresa As String
    Dim vlstrNumDepartamento As String
    Dim vlstrNumEmpleado As String
    Dim vlstrTipoPoliza As String
    Dim vlstrAsentada As String
    Dim vlstrDescuadradas As String
    Dim vlstrDetallado As String
    Dim vlstrCuadre As String
    Dim vlstrNumCuentaCuadre As String
    Dim vlstrOrden As String
    Dim vlstrIncluirCuentaInmediata As String
    Dim vlstrFiltroFolios As String
    Dim vlstrFolioInicio As String
    Dim vlstrFolioFin As String
    Dim vlstrFiltroFechas As String
    Dim vlstrFechaInicio As String
    Dim vlstrFechaFin As String

    vlstrNumUsuario = "0"
    vlstrEmpresa = "0"
    vlstrNumDepartamento = "0"
    vlstrNumEmpleado = "0"
    vlstrTipoPoliza = "*"
    vlstrAsentada = "0"
    vlstrDescuadradas = "0"
    vlstrDetallado = "0"
    vlstrCuadre = "0"
    vlstrNumCuentaCuadre = "0"
    vlstrOrden = "4"
    vlstrIncluirCuentaInmediata = IIf(chkCuentaMayor.Value, "0", "1") 'Indicar si se va a imprimir la cuenta inmediata mayor
    vlstrFiltroFolios = "0"
    vlstrFolioInicio = "0"
    vlstrFolioFin = "0"
    vlstrFiltroFechas = "0"
    vlstrFechaInicio = fstrFechaSQL(fdtmServerFecha, , True)
    vlstrFechaFin = fstrFechaSQL(fdtmServerFecha, , True)
    
    pInstanciaReporte vgrptReporte, "rptPoliza.rpt"
    vgrptReporte.DiscardSavedData
    
    vgstrParametrosSP = vlstrNumPoliza & "|" & _
                        vlstrNumUsuario & "|" & _
                        vlstrEmpresa & "|" & _
                        vlstrNumDepartamento & "|" & _
                        vlstrNumEmpleado & "|" & _
                        vlstrTipoPoliza & "|" & _
                        vlstrAsentada & "|" & _
                        vlstrDescuadradas & "|" & _
                        vlstrDetallado & "|" & _
                        vlstrCuadre & "|" & _
                        vlstrNumCuentaCuadre & "|" & _
                        vlstrOrden & "|" & _
                        vlstrIncluirCuentaInmediata & "|" & _
                        vlstrFiltroFolios & "|" & _
                        vlstrFolioInicio & "|" & _
                        vlstrFolioFin & "|" & _
                        vlstrFiltroFechas & "|" & _
                        vlstrFechaInicio & "|" & _
                        vlstrFechaFin
    Set rsReporte = frsEjecuta_SP(vgstrParametrosSP, "Sp_CnRptPoliza")
    Set rsReporte = frsUltimoRecordset(rsReporte)
    If rsReporte.RecordCount > 0 Then
        alstrParametros(0) = "DesglosadoCuenta;" & Trim(vlintDesglosadoCuenta)
        alstrParametros(1) = "IN_POLIZAPAGINA;" & "0"
        pCargaParameterFields alstrParametros, vgrptReporte
        pImprimeReporte vgrptReporte, rsReporte, vlstrDestino, "Póliza"
    Else
        'No existe información con esos parámetro
        MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje"
    End If
    rsReporte.Close
    
    frsEjecuta_SP 1 & "|" & Me.Name & "|" & chkConcentrado.Name & "|Value|" & vglngNumeroLogin & "|" & Trim(str(chkConcentrado.Value)), "SP_GNSELULTIMACONFIGURACION", True
End Sub

'(CR) Caso 7323: Guardar la selección de la opción "Incluir cuenta inmediata mayor"'
Private Sub pActualizaProcesoCuenta()
On Error GoTo NotificaError

    Dim strSentencia As String

    EntornoSIHO.ConeccionSIHO.BeginTrans
    
    strSentencia = "SELECT * FROM CpProcesoCuentaMayor" & _
                   " WHERE intNumeroLogin = " & CStr(vglngNumeroLogin) & _
                   " AND   intProceso = " & CStr(lintProceso)
    Set rsCuentaMayor = frsRegresaRs(strSentencia, adLockOptimistic, adOpenDynamic)
    With rsCuentaMayor
        If .RecordCount = 0 Then .AddNew
        !intProceso = lintProceso
        !intNumeroLogin = vglngNumeroLogin
        !intCuentaMayor = IIf(chkCuentaMayor.Value = 0, 0, 1)
        .Update
    End With
    rsCuentaMayor.Close
    
    EntornoSIHO.ConeccionSIHO.CommitTrans
    
Exit Sub
NotificaError:
    Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":pActualizaProcesoCuenta"))
End Sub

'(CR) Caso 7442: Mostrar pólizas fuera del corte '
Private Sub pCargaPolizasFueraCorte()
On Error GoTo NotificaError

    Dim rsPolizas As New ADODB.Recordset
    Dim intCol As Integer
    
    grdPolizas1.Rows = 0
    grdDetallePoliza1.Rows = 0
    
    cmdImprimirPoliza.Enabled = False
    cmdExportar.Enabled = False
    cmdPreview.Enabled = False
    chkConcentrado.Enabled = False
        
    If cboDepartamento.ListIndex < 0 Then Exit Sub

    vgstrParametrosSP = str(vgintClaveEmpresaContable) & _
                        "|" & vllngNumCorte & _
                        "|" & cboDepartamento.ItemData(cboDepartamento.ListIndex) & _
                        "|" & "*" & "|" & 0 & _
                        "|" & Format(vldtmFechaIni, "dd/mm/yyyy") & _
                        "|" & Format(vldtmFechaFin, "dd/mm/yyyy")
    Set rsPolizas = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELPOLIZASFUERACORTE")
    If Not rsPolizas.EOF Then
        With grdPolizas1
            .Rows = 1
            .FixedRows = 1
            .TextMatrix(0, 1) = "Fecha"
            .TextMatrix(0, 2) = "Número"
            .TextMatrix(0, 3) = "Concepto"
            .TextMatrix(0, 4) = "Tipo"
            .TextMatrix(0, 5) = "Departamento"
            .ColWidth(0) = 180
            .ColWidth(1) = 1100
            .ColWidth(2) = 800
            .ColWidth(3) = 6000
            .ColWidth(4) = 800
            .ColWidth(5) = 1000
            .ColWidth(6) = 0
        End With
    
        Do Until rsPolizas.EOF
            grdPolizas1.AddItem ""
            For intCol = 0 To rsPolizas.Fields.Count - 1
                grdPolizas1.TextMatrix(grdPolizas1.Rows - 1, intCol + 1) = rsPolizas.Fields(intCol).Value
            Next
            grdPolizas1.TextMatrix(grdPolizas1.Rows - 1, 1) = Format(rsPolizas.Fields(0).Value, "dd/mmm/yyyy")
            
            rsPolizas.MoveNext
        Loop
        
        chkConcentrado.Enabled = True
        cmdImprimirPoliza.Enabled = True
        If cboInterfazPolizas.ListIndex <> -1 Then
            If fblnValidarLicenciaPolizas(vgintClaveEmpresaContable, cboInterfazPolizas.ItemData(cboInterfazPolizas.ListIndex), cboInterfazPolizas.Text) Then
                cmdExportar.Enabled = IIf(chkConcentrado.Value, False, IIf(grdPolizas1.Rows > 0, True, False))
            Else
                cmdExportar.Enabled = False
            End If
        Else
            cmdExportar.Enabled = False
        End If
        cmdPreview.Enabled = True
        grdPolizas1.Row = 1 'Posicionar en la primer póliza
        pCargaDetallePoliza 'Desplegar detallado de la primer póliza
        lblnExistePolizas = True
    Else
        lblnExistePolizas = False
    End If
    rsPolizas.Close
    
Exit Sub
NotificaError:
    Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":pCargaPolizasFueraCorte"))
End Sub

'- Presentación de la forma -'
Private Sub pPosicionaObjetos()
    If vlblnPolizasFueraCorte Then 'Si se están consultando pólizas fuera del corte
        Frame1.Visible = False
        Frame2.width = 2565
        grdPolizas1.Top = Frame1.Top
        grdPolizas1.height = grdPolizas1.height + Frame1.height + 120
    Else
        Frame1.Visible = True
        grdPolizas1.Top = 1150
        grdPolizas1.height = 3190
    End If
End Sub

'***************************** PROCEDIMIENTOS AGREGADOS PARA INTERFAZ CON MICROSIP *****************************'
'|  Importa una póliza del sistema MICROSIP  |'
Public Sub pImportaPolizaMicrosip()
On Error GoTo NotificaError

    Dim fs As FileSystemObject
    Dim f As Variant
    Dim strLinea As String
    Dim rsPoliza As New ADODB.Recordset
    Dim rsDetallePoliza As New ADODB.Recordset
    Dim rsPolizasImportadas As New ADODB.Recordset
    Dim strArchivoImportacion As String         '|  Archivo que se intenta importar
    Dim strSentencia As String                  '|  Variable genérica para armar instrucciones SQL
    Dim lngCvePolizaNueva As Long               '|  Clave de la póliza que se desea importar
    Dim lngCveCuentaNueva As Long               '|  Clave de la cuenta del detalle de la póliza que se intenta importar
    Dim lngPersonaGraba As Long
    Dim lngResultado As Long
    Dim tArreglo() As String
    Dim strPolizasCanceladas As String
    Dim blnCancelada As Boolean
    Dim lintInterfazImporta As Integer
    
    '---------------------------------------------------------------------------'
    '|  Validaciones iniciales que pueden suspender la ejecución de la rutina  |'
    '---------------------------------------------------------------------------'
    '|  Selecciona el archivo que se desea importar
    strArchivoImportacion = fstrAbreArchivo
    '|  Si se canceló el diálogo de abrir
    If strArchivoImportacion = "" Then Exit Sub
    
    '|  Valida que el formato del archivo de entrada sea del tipo esperado (MICROSIP)
    If Not fblnValidacionesGeneralesMicrosip(strArchivoImportacion) Then Exit Sub
    
    '|  Pide autorización para hacer la importación de la póliza
    lngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
    If lngPersonaGraba = 0 Then Exit Sub
    
    '---------------------------------'
    '|  Inicialización de variables  |'
    '---------------------------------'
    '- Crear objeto para manejo de archivo -'
    Set fs = CreateObject("Scripting.FileSystemObject")
    '- Abrir archivo -'
    Set f = fs.OpenTextFile(strArchivoImportacion, ForReading, , TristateUseDefault)
    '- Abrir recordset para manejo de datos de la póliza -'
    vlstrsql = "SELECT bitAsentada, chrTipoPoliza, dtmFechaPoliza, intClavePoliza, intCveEmpleado, intNumeroPoliza, smiCveDepartamento, smiEjercicio, tnyClaveEmpresa, tnyMes, vchConceptoPoliza, vchNumero " & _
               "FROM   CNPOLIZA " & _
               "WHERE  intNumeroPoliza = -1"
    Set rsPoliza = frsRegresaRs(vlstrsql, adLockOptimistic, adOpenDynamic)
    '- Abrir recordset para manejo de datos del detalle de la póliza -'
    vlstrsql = "SELECT bitNaturalezaMovimiento, intNumeroCuenta, intNumeroPoliza, intNumeroRegistro, mnyCantidadMovimiento, vchConcepto, vchReferencia " & _
               "FROM   CNDETALLEPOLIZA " & _
               "WHERE  intNumeroPoliza = -1"
    Set rsDetallePoliza = frsRegresaRs(vlstrsql, adLockOptimistic, adOpenDynamic)
    '- Abrir recordset para manejo de datos pólizas importadas -'
    vlstrsql = "SELECT * " & _
               "FROM   CNPOLIZASIMPORTADAS " & _
               "WHERE  chrTipoPoliza = '*' "
    Set rsPolizasImportadas = frsRegresaRs(vlstrsql, adLockOptimistic, adOpenDynamic)
    
    Me.MousePointer = 11 '|  Reloj de arena
    
    '|  Clave de la interfaz seleccionada para la importación
    If cboInterfazPolizas.ListIndex <> -1 Then
        lintInterfazImporta = cboInterfazPolizas.ItemData(cboInterfazPolizas.ListIndex)
    Else
        lintInterfazImporta = lintInterfazPolizas
    End If
    
    EntornoSIHO.ConeccionSIHO.BeginTrans    '|  Comienza transacción
    
    '|  Bloquea el cierre
    lngResultado = 1
    vgstrParametrosSP = vgintClaveEmpresaContable & "|" & "Grabando poliza"
    frsEjecuta_SP vgstrParametrosSP, "sp_CnUpdEstatusCierre", True, lngResultado
    If lngResultado = 1 Then
        Do Until f.AtEndOfStream
            DoEvents
            ReDim tArreglo(0)
            strLinea = f.ReadLine '|  Leer la línea del archivo de texto
            strLinea = Mid(strLinea, 2, Len(strLinea)) '|  Quitar el primer elemento (pipe)
            strLinea = Replace(strLinea, """", "") '|  Sustituir las comillas dobles
            strLinea = Replace(strLinea, "|", ",") '|  Sustituir los pipes por comas para que se dividan los datos
            tArreglo = Split(strLinea, ",") '|  Convertir los elementos a arreglo para un manejo más fácil
            '|  Indica el tipo de registro 1 = Datos generales
            If tArreglo(intPosNivel) = "1" Then
'                If Trim(tArreglo(intPosCancelada)) = "N" Then 'Si la póliza no está cancelada
'                    blnCancelada = False
                    '-----------------------------------'
                    '|  Graba el maestro de la póliza  |'
                    '-----------------------------------'
                    With rsPoliza
                        .AddNew
                        '|  Tipo de póliza "I" Ingresos, "E" Egresos, "D" Diario, "O" De orden
                        !chrTipoPoliza = tArreglo(intPosTipo)
                        '|  Fecha de alta de la póliza
                        !dtmFechaPoliza = tArreglo(intPosFecha)
                        '|  Concepto de la póliza
                        !vchConceptoPoliza = Left(tArreglo(intPosDescripcion), 100)
                        '|  Indica si la póliza está incluida en un cierre contable (guarda 1 = si está incluida, 0 = no está incluida)
                        !BITASENTADA = 0
                        '|  Empleado que esta realizando la importación
                        !intCveEmpleado = lngPersonaGraba
                        '|  Clave de la empresa contable (relacionado con CnEmpresaContable)
                        !TNYCLAVEEMPRESA = vgintClaveEmpresaContable
                        '|  Departamento del usuario que esta en el sistema
                        !smicvedepartamento = vgintNumeroDepartamento
                        '|  Año del ejercicio contable.
                        !SMIEJERCICIO = Format(!dtmFechaPoliza, "yyyy")
                        '|  Número de mes contable.
                        !tnyMes = Format(!dtmFechaPoliza, "mm")
                        '|  Número de póliza para manejo interno del área de contabilidad
                        !vchNumero = " "
                        '|  El número de póliza es generado por el sistema según este configurado
                        !intClavePoliza = flngFolioPoliza(vgintClaveEmpresaContable, !chrTipoPoliza, !SMIEJERCICIO, !tnyMes, False)
                        .Update
                        
                        lngCvePolizaNueva = flngObtieneIdentity("SEC_CNPOLIZA", !intNumeroPoliza)
                    End With
                    
                    '--------------------------------------------------------------------------------------'
                    '|  Registra las pólizas que se están importando si previamente no se han registrado  |'
                    '--------------------------------------------------------------------------------------'
'                    vlstrsql = "SELECT Count(*) Co " & _
'                               " FROM  CNPOLIZASIMPORTADAS " & _
'                               " WHERE NVL(chrTipoPoliza, ' ') = '" & IIf(Trim(tArreglo(intPosTipo)) = "", " ", tArreglo(intPosTipo)) & "' " & _
'                               " AND   dtmFechaPoliza = " & fstrFechaSQL(tArreglo(intPosFecha)) & _
'                               " AND   intInterfazPoliza = " & lintInterfazImporta
'                               '" AND   intCvePolizaExterna = " & Val(Format(tArreglo(intPosPoliza), "000000"))
'                    If frsRegresaRs(vlstrsql, adLockOptimistic, adOpenDynamic)!CO = 0 Then
                        With rsPolizasImportadas
                            .AddNew
                            !chrTipoPoliza = tArreglo(intPosTipo)
                            !dtmFechaPoliza = tArreglo(intPosFecha)
                            !intInterfazPoliza = lintInterfazImporta
                            !intNumeroPoliza = lngCvePolizaNueva
                            '!intCvePolizaExterna = Val(Format(tArreglo(intPosPoliza), "000000"))
                            .Update
                        End With
'                    End If
'                Else
'                    blnCancelada = True
'                    strPolizasCanceladas = strPolizasCanceladas & tArreglo(intPosPoliza) & ", "
'                End If
            Else 'If Not blnCancelada Then 'Si la póliza NO está cancelada
            '|  1.1 = Asientos de la póliza
                '-----------------------------------'
                '|  Graba el detalle de la póliza  |'
                '-----------------------------------'
                With rsDetallePoliza
                    .AddNew
                    lngCveCuentaNueva = fintFormatoCuentaMicrosip(tArreglo(intPosCuenta), vgintClaveEmpresaContable)
                    If lngCveCuentaNueva = -1 Then
                        EntornoSIHO.ConeccionSIHO.RollbackTrans
                        Me.MousePointer = 0
                        Exit Sub
                    End If
                    !intNumeroCuenta = lngCveCuentaNueva
                    !mnyCantidadMovimiento = CDbl(tArreglo(intPosCantidad))
                    !bitNaturalezaMovimiento = IIf(Trim(tArreglo(intPosNaturaleza)) = "C", 1, 0)
                    !intNumeroPoliza = lngCvePolizaNueva
                    !vchReferencia = IIf(Trim(tArreglo(intPosReferencia)) = "", " ", Trim(tArreglo(intPosReferencia)))
                    !vchConcepto = IIf(Trim(tArreglo(intPosConcepto)) = "", " ", Trim(tArreglo(intPosReferencia)))
                    .Update
                End With
            End If
        Loop
        
        pEjecutaSentencia "UPDATE CnEstatusCierre SET vchEstatus = 'Libre' WHERE tnyClaveEmpresa = " & str(vgintClaveEmpresaContable)
        Call pGuardarLogTransaccion(Me.Name, EnmBorrar, lngPersonaGraba, "IMPORTACION DE POLIZAS", strArchivoImportacion)
    End If
    f.Close
    
    EntornoSIHO.ConeccionSIHO.CommitTrans
    Me.MousePointer = 0
    
    If strPolizasCanceladas <> "" Then
        strPolizasCanceladas = Left(Trim(strPolizasCanceladas), Len(Trim(strPolizasCanceladas)) - 1)
        MsgBox SIHOMsg(1210) & vbCrLf & strPolizasCanceladas, vbInformation, "Mensaje"
    End If

    If lngResultado <> 0 Then
        '|  La póliza ha sido importada exitosamente
        MsgBox SIHOMsg(744), vbOKOnly + vbInformation, "Mensaje"
    End If
    
Exit Sub
NotificaError:
    EntornoSIHO.ConeccionSIHO.RollbackTrans
    Me.MousePointer = 0
    Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":pImportaPolizaMicrosip"))
End Sub

'|  Exporta una póliza con el formato del sistema MICROSIP  |'
Private Sub pExportaPolizaMicrosip()
On Error GoTo NotificaError
    Dim vllngCont As Long
    Dim vllngNumeroMayor As Long
    Dim vllngNumeroMenor As Long
    Dim rsPoliza As New ADODB.Recordset
    Dim vlstrCadena As String
    Dim vllngPoliza As Long
    Dim vlstrTemp As String
    Dim vlblnPolizasExcuidas As String
    Dim lngVal As Long
    Dim intcontador As Long
    Dim rsTemp As New ADODB.Recordset
    Dim rsXMLs As New ADODB.Recordset
    Dim vllngGenerado As Long
    Dim aRenglon() As String
    Dim lintCont As Long
    Dim vlblnSiGeneraraInfo As Boolean
    Dim stmXMLs As New ADODB.Stream
    Dim fsoArchivoXML As New FileSystemObject
    Dim vlStrRuta As String
    Dim vlstrArchivo As String
    Dim vlstrUUID As String
    Dim vlfechahoraserver As Date
    Dim vlRutaXMLs As String
    Dim vldblAumBarra As Double

    lblTextoBarra.Caption = " Obteniendo el listado de las pólizas, por favor espere..."
    pgbBarra.Value = 0
    fblnMuestraBarraProgreso True
    pgbBarra.Refresh
    lblTextoBarra.Refresh

    ' -- Nuevo metodo de exportación incluyendo la información de los UUID relacionados a la póliza --
        EntornoSIHO.ConeccionSIHO.BeginTrans

        'Guarda la información de lo que se está exportando
        Set rsTemp = frsRegresaRs("SELECT * FROM CNGENERACIONPOLIZASMICROSIP WHERE INTIDREGISTRO = -1", adLockOptimistic, adOpenDynamic)
        With rsTemp
            .AddNew
            !smicvedepartamento = IIf(cboDepartamento.ItemData(cboDepartamento.ListIndex) <= 0, 0, cboDepartamento.ItemData(cboDepartamento.ListIndex))
            !chrTipoPoliza = IIf(Trim(cboTipo.List(cboTipo.ListIndex)) = "Diario", "D", IIf(Trim(cboTipo.List(cboTipo.ListIndex)) = "Ingreso", "I", IIf(Trim(cboTipo.List(cboTipo.ListIndex)) = "Egreso", "E", "T")))
            !DTMFECHAPOLIZAINI = CDate(mskFechaInicio)
            !DTMFECHAPOLIZAFIN = CDate(mskFechaFin)
            !CLBPOLIZAS = ""
            .Update
        End With
        'Se trae la secuencia del que acaba de generar
        vllngGenerado = flngObtieneIdentity("SEC_CNGENERAPOLIZASMICROSIP", rsTemp!intIdRegistro)

        vlblnPolizasExcuidas = ""
        vlblnSiGeneraraInfo = False

        'Llena la tabla temporal con las pólizas que se exportarán
        lngVal = 1
        pEjecutaSentencia "DELETE FROM CNTMPGENERAPOLIZASMICROSIP"
        vldblAumBarra = 10 / (grdPolizas1.Rows - 1)
        For vllngCont = 1 To grdPolizas1.Rows - 1
            If grdPolizas1.TextMatrix(vllngCont, 0) = "*" Then
                'Revisar si alguna póliza contiene alguna cuenta de cuadre
                If Not fblnCuentaCuadre(CLng(grdPolizas1.TextMatrix(vllngCont, 6))) Then
                    lngVal = 1
                    frsEjecuta_SP grdPolizas1.TextMatrix(vllngCont, 6), "FN_CNINSTMPGENPOLIZASMICROSIP", True, lngVal
                    vlblnSiGeneraraInfo = True
                Else
                    'Almacenar las pólizas con cuenta de cuadre para indicar que no fueron exportadas
                    vlblnPolizasExcuidas = vlblnPolizasExcuidas & grdPolizas1.TextMatrix(vllngCont, 6) & ", "
                End If
            End If
            pgbBarra.Value = IIf(pgbBarra.Value + vldblAumBarra > 100, 100, pgbBarra.Value + vldblAumBarra)
        Next vllngCont
    
    
        If vlblnSiGeneraraInfo Then
            pgbBarra.Value = 10
            lblTextoBarra.Caption = " Recabando información de pólizas y comprobantes, por favor espere..."
            pgbBarra.Refresh
            lblTextoBarra.Refresh
            
            'Genera la cadena a exportar y la guarda
            frsEjecuta_SP str(vllngGenerado) & "|" & vgintClaveEmpresaContable, "SP_CNGENERAPOLIZASMICROSIP", True

            Set rsPoliza = frsRegresaRs("SELECT CLBPOLIZAS FROM CNGENERACIONPOLIZASMICROSIP WHERE INTIDREGISTRO = " & vllngGenerado)
            With rsPoliza
                If .State <> adStateClosed Then
                    If .RecordCount = 0 Then
                        fblnMuestraBarraProgreso False
                        
                        ' No existe información
                        MsgBox SIHOMsg(12), vbOKOnly + vbExclamation, "Exportación de póliza(s)"
                        EntornoSIHO.ConeccionSIHO.RollbackTrans
                    Else
                        pgbBarra.Value = 20
                        lblTextoBarra.Caption = " Exportando las pólizas en archivo de texto, por favor espere..."
                        pgbBarra.Refresh
                        lblTextoBarra.Refresh
                        
                        Open CDgArchivo.FileName For Output As #1  ' Open file for output.
                        .MoveFirst

                        aRenglon = Split(!CLBPOLIZAS, "|Rng|")
                        vldblAumBarra = 40 / UBound(aRenglon)
                        For lintCont = 1 To UBound(aRenglon)
                            Print #1, Trim(aRenglon(lintCont))
                            
                            pgbBarra.Value = IIf(pgbBarra.Value + vldblAumBarra > 100, 100, pgbBarra.Value + vldblAumBarra)
                        Next lintCont

                        Close #1

                        'Quitar a la información exportada los identificadores de cada renglon
                        pEjecutaSentencia "UPDATE CNGENERACIONPOLIZASMICROSIP SET CLBPOLIZAS = REPLACE(CLBPOLIZAS,'|Rng|','') WHERE INTIDREGISTRO = " & vllngGenerado

                        vlStrRuta = Replace(CDgArchivo.FileName, CDgArchivo.FileTitle, "")
                        vlstrArchivo = Replace(CDgArchivo.FileTitle, ".txt", "")
                        
                        vlstrUUID = ""
                        
                        Set rsXMLs = frsEjecuta_SP("", "SP_CNSELTMPEXPORTAXMLSMICROSIP")
                        If rsXMLs.RecordCount <> 0 Then
                            pgbBarra.Value = 60
                            lblTextoBarra.Caption = " Exportando los comprobantes fiscales digitales, por favor espere..."
                            pgbBarra.Refresh
                            lblTextoBarra.Refresh
                            
                            rsXMLs.MoveFirst
                    
                            vlfechahoraserver = fdtmServerFechaHora
                            vlRutaXMLs = Trim(vlStrRuta) & "\" & vlstrArchivo & " " & Replace(Replace(vlfechahoraserver, "/", "."), ":", ".")
                            pCreaDirectorio vlRutaXMLs
                            
                            vldblAumBarra = 40 / rsXMLs.RecordCount
                            While Not rsXMLs.EOF
                                If Trim(vlstrUUID) = "" Or Trim(vlstrUUID) <> Trim(rsXMLs!VCHUUIDCFDI) Then
                                    'Crea físicamente el XML
                                    stmXMLs.Open
                                    stmXMLs.Charset = "utf-8"
                                    stmXMLs.WriteText rsXMLs!CLBXML, adWriteChar
                                    stmXMLs.SaveToFile vlRutaXMLs & "\" & rsXMLs!VCHUUIDCFDI & ".xml", adSaveCreateNotExist
                                    stmXMLs.Close
    
                                    vlstrUUID = Trim(rsXMLs!VCHUUIDCFDI)
                                End If
                                
                                rsXMLs.MoveNext
                                
                                pgbBarra.Value = IIf(pgbBarra.Value + vldblAumBarra > 100, 100, pgbBarra.Value + vldblAumBarra)
                            Wend
                        End If
                        
                        pgbBarra.Value = 100
                        
                        rsXMLs.Close
                        
                        fblnMuestraBarraProgreso False

                        If vlblnPolizasExcuidas <> "" Then
                            vlblnPolizasExcuidas = Left(Trim(vlblnPolizasExcuidas), Len(Trim(vlblnPolizasExcuidas)) - 1)
                            'Las siguientes pólizas contienen una cuenta de cuadre por lo que no fueron exportadas:
                            MsgBox SIHOMsg(1209) & vbCrLf & vlblnPolizasExcuidas, vbInformation, "Mensaje"
                        End If

                        '¡Los datos han sido guardados satisfactoriamente!
                        MsgBox SIHOMsg(358), vbOKOnly + vbInformation, "Exportación de póliza(s)"

                        EntornoSIHO.ConeccionSIHO.CommitTrans
                    End If
                    .Close
                Else
                    fblnMuestraBarraProgreso False
                    
                    EntornoSIHO.ConeccionSIHO.RollbackTrans
                End If
            End With
        Else
            fblnMuestraBarraProgreso False
        
            EntornoSIHO.ConeccionSIHO.RollbackTrans

            If vlblnPolizasExcuidas <> "" Then
                vlblnPolizasExcuidas = Left(Trim(vlblnPolizasExcuidas), Len(Trim(vlblnPolizasExcuidas)) - 1)
                'Las siguientes pólizas contienen una cuenta de cuadre por lo que no fueron exportadas:
                MsgBox SIHOMsg(1209) & vbCrLf & vlblnPolizasExcuidas, vbInformation, "Mensaje"
            End If
        End If
        
    fblnMuestraBarraProgreso False
        
    ' --

'' -- Forma anterior en la que se generaba el archivo
'    vlstrsql = "SELECT DP.intNumeroPoliza,DP.bitNaturalezaMovimiento, DP.mnyCantidadMovimiento, DP.vchReferencia, DP.vchConcepto, "
'    vlstrsql = vlstrsql & " P.dtmFechaPoliza, P.chrTipoPoliza, P.intClavePoliza, P.vchConceptoPoliza, C.vchCuentaContable, p.smicvedepartamento "
'    vlstrsql = vlstrsql & " FROM CnDetallePoliza DP, CnPoliza P, CnCuenta C"
'    vlstrsql = vlstrsql & " WHERE DP.intNumeroPoliza = P.intNumeroPoliza AND DP.intNumeroCuenta = C.intNumeroCuenta"
'    vlstrsql = vlstrsql & " AND DP.intNumeroPoliza IN (-1"
'    For vlLngCont = 1 To grdPolizas1.Rows - 1
'        If grdPolizas1.TextMatrix(vlLngCont, 0) = "*" Then
'            '- Revisar si alguna póliza contiene alguna cuenta de cuadre -'
'            If Not fblnCuentaCuadre(CLng(grdPolizas1.TextMatrix(vlLngCont, 6))) Then
'                vlstrsql = vlstrsql & ", " & grdPolizas1.TextMatrix(vlLngCont, 6)
'            Else
'                '- Almacenar las pólizas con cuenta de cuadre para indicar que no fueron exportadas -'
'                vlblnPolizasExcuidas = vlblnPolizasExcuidas & grdPolizas1.TextMatrix(vlLngCont, 6) & ", "
'            End If
'        End If
'    Next vlLngCont
'    vlstrsql = vlstrsql & ")"
'    If Trim(cboTipo.List(cboTipo.ListIndex)) = "Diario" Then
'        vlstrsql = vlstrsql & " AND P.chrTipoPoliza = 'D'"
'    End If
'    If Trim(cboTipo.List(cboTipo.ListIndex)) = "Ingreso" Then
'        vlstrsql = vlstrsql & " AND P.chrTipoPoliza = 'I'"
'    End If
'    If Trim(cboTipo.List(cboTipo.ListIndex)) = "Egreso" Then
'        vlstrsql = vlstrsql & " AND P.chrTipoPoliza = 'E'"
'    End If
'    vlstrsql = vlstrsql & IIf(cboDepartamento.ListIndex > 0, " AND P.smiCveDepartamento = " & cboDepartamento.ItemData(cboDepartamento.ListIndex), "")
'    vlstrsql = vlstrsql & " ORDER BY DP.intNumeroPoliza, DP.bitNaturalezaMovimiento DESC, C.vchCuentaContable"
'    Set rsPoliza = frsRegresaRs(vlstrsql)
'    With rsPoliza
'        If .State <> adStateClosed Then
'            If .RecordCount = 0 Then
'                ' No existe información
'                MsgBox SIHOMsg(12), vbOKOnly + vbExclamation, "Exportación de póliza(s)"
'            Else
'                Open CDgArchivo.FileName For Output As #1  ' Open file for output.
'                .MoveFirst
'                vllngPoliza = -1
'                vllngNumeroMayor = 1
'                vllngNumeroMenor = 1
'                While Not .EOF
'                    vlstrTemp = ""
'                    If vllngPoliza <> !intNumeroPoliza Then
'                        vllngPoliza = !intNumeroPoliza
'                        '|  Identificación (Maestro de la póliza)
'                        vlstrCadena = lstrSeparador & CStr(vllngNumeroMayor) & lstrSeparador
'                        '|  Tipo de Póliza
'                        vlstrCadena = vlstrCadena & lstrCalificadorTexto & Trim(!chrTipoPoliza) & lstrCalificadorTexto & lstrDelimitador
'                        '|  Clave de la Póliza
'                        vlstrCadena = vlstrCadena & lstrCalificadorTexto & CStr(!intClavePoliza) & lstrCalificadorTexto & lstrDelimitador
'                        '|  Fecha de alta de la póliza
'                        vlstrCadena = vlstrCadena & Format(!dtmFechaPoliza, lstrFormatoFechas) & lstrDelimitador
'                        '|  Clave de la moneda, tipo de cambio, estatus de la póliza (S = Cancelada, N = No cancelada)
'                        vlstrCadena = vlstrCadena & lstrCalificadorTexto & "1" & lstrCalificadorTexto & lstrDelimitador & 1 & lstrDelimitador & lstrCalificadorTexto & "N" & lstrCalificadorTexto & lstrDelimitador
'                        '|  Concepto de la póliza
'                        If Not IsNull(!vchConceptoPoliza) Then
'                            vlstrTemp = Trim(Mid(Replace(!vchConceptoPoliza, """", "'"), 1, 200)) 'Quitar comillas dobles antes de agregar el calificador de texto
'                        End If
'                        vlstrCadena = vlstrCadena & lstrCalificadorTexto & vlstrTemp & lstrCalificadorTexto
'                        Print #1, vlstrCadena ' Maestro o Encabezado de la Póliza
'                    End If
'                    '|  Identificación (Detalle de la póliza)
'                    vlstrCadena = lstrSeparador & CStr(vllngNumeroMayor) & "." & CStr(vllngNumeroMenor) & lstrSeparador
'                    '|  Cuenta Contable a la que afecta
'                    vlstrCadena = vlstrCadena & lstrCalificadorTexto & fstrFormateaCuentaMicrosip(!vchCuentaContable) & lstrCalificadorTexto & lstrDelimitador
'                    '|  Clave del departamento
'                    vlstrCadena = vlstrCadena & lstrCalificadorTexto & CStr(!smicvedepartamento) & lstrCalificadorTexto & lstrDelimitador
'                    '|  Tipo de Movimiento 1 Cargo, 0 Abono
'                    vlstrCadena = vlstrCadena & lstrCalificadorTexto & IIf(!bitNaturalezaMovimiento = 0, "A", "C") & lstrCalificadorTexto & lstrDelimitador
'                    '|  Importe del movimiento
'                    vlstrCadena = vlstrCadena & Format(!mnyCantidadMovimiento, "###0.00") & lstrDelimitador
'                    '|  Referencia del Movimiento
'                    vlstrTemp = IIf(IsNull(!vchReferencia), "", Trim(Mid(Replace(!vchReferencia, """", "'"), 1, 10))) 'Quitar comillas dobles antes de agregar el calificador de texto
'                    vlstrCadena = vlstrCadena & lstrCalificadorTexto & vlstrTemp & lstrCalificadorTexto & lstrDelimitador
'                    '|  Concepto del Movimiento
'                    vlstrTemp = IIf(IsNull(!vchConcepto), "", Trim(Mid(Replace(!vchConcepto, """", "'"), 1, 200))) 'Quitar comillas dobles antes de agregar el calificador de texto
'                    vlstrCadena = vlstrCadena & lstrCalificadorTexto & vlstrTemp & lstrCalificadorTexto
'                    Print #1, vlstrCadena ' Detalle o Movimientos de la Póliza
'                    .MoveNext
'                Wend
'                Close #1 ' Close file.
'
'                If vlblnPolizasExcuidas <> "" Then
'                    vlblnPolizasExcuidas = Left(Trim(vlblnPolizasExcuidas), Len(Trim(vlblnPolizasExcuidas)) - 1)
'                    'Las siguientes pólizas contienen una cuenta de cuadre por lo que no fueron exportadas:
'                    MsgBox SIHOMsg(1209) & vbCrLf & vlblnPolizasExcuidas, vbInformation, "Mensaje"
'                End If
'
'                '¡Los datos han sido guardados satisfactoriamente!
'                MsgBox SIHOMsg(358), vbOKOnly + vbInformation, "Exportación de póliza(s)"
'            End If
'            .Close
'        End If
'    End With
'' --
    
Exit Sub
NotificaError:
    Close #1
    EntornoSIHO.ConeccionSIHO.RollbackTrans
    Me.MousePointer = 0
    Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":pExportaPolizaMicrosip"))
End Sub

'|  Importa una póliza del sistema Apsi
Public Sub pImportaPolizaApsi()
    Dim fs As FileSystemObject
    'Dim f As Variant
    Dim strLinea As String
    Dim rsPoliza As New ADODB.Recordset
    Dim rsDetallePoliza As New ADODB.Recordset
    Dim rsPolizasImportadas As New ADODB.Recordset
    Dim strArchivoImportacion As String         '|  Archivo que se intenta importar
    Dim strSentencia As String                  '|  Variable genérica para armar instrucciones SQL
    Dim lngCvePolizaNueva As Long               '|  Clave de la póliza que se desea importar
    Dim lngCveCuentaNueva As Long               '|  Clave de la cuenta del detalle de la póliza que se intenta importar
    Dim lngPersonaGraba As Long
    Dim lngResultado As Long
    Dim lintInterfazImporta As Integer
    Dim rsUuid As New ADODB.Recordset
    
    Dim ExcelObj As Object
    Dim ExcelBook As Object
    Dim ExcelSheet As Object
    Dim i As Integer
    'variables para la lecura del archivo
    Dim strCuenta As String
    Dim strDebito As String
    Dim strCredito As String
    Dim strTipo As String
    Dim strFecha As String
    Dim strConcepto As String
    Dim strUUId As String
    Dim strTotal As String
    Dim strRFC As String
    Dim strNombre As String
    Dim strTmpUuid As String
    
    Const xlValues = -4163
    Const xlWhole = 1
    Const xlByRows = 1
    Const xlNext = 1
    
On Error GoTo NotificaError
    
    '-----------------------------------------------------------------------------
    '|  Validaciones iniciales que pueden suspender la ejecución de la rutina
    '-----------------------------------------------------------------------------
    '|  Selecciona el archivo que se desea importar
    strArchivoImportacion = fstrAbreArchivoImportaApsi
    '|  Si se canceló el diálogo de abrir
    If strArchivoImportacion = "" Then Exit Sub
    '|  Valida que el formato del archivo de entrada sea del tipo esperado (CONTPAQ),
    '|  que las pólizas que se intentan importar pertenezcan a un periodo abierto y
    '|  que las pólizas no hayan sido importadas anteriormente
    '|  si se obtendra el uuid relacionado con la poliza por medio de otro archivo
   If Not fblnValidacionesGeneralesApsi(strArchivoImportacion) Then Exit Sub
        
    '|  Pide autorización para hacer la importación de la póliza
    lngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
    If lngPersonaGraba = 0 Then Exit Sub
    
    '------------------------------------------------
    '|  Inicialización de variables
    '------------------------------------------------
    strSentencia = "SELECT bitasentada, chrtipopoliza, dtmfechapoliza, intclavepoliza, intcveempleado, intnumeropoliza, smicvedepartamento, smiejercicio, tnyclaveempresa, tnymes, vchconceptopoliza, vchnumero " & _
                   "FROM   CNPOLIZA " & _
                   "WHERE  intnumeropoliza = -1"
    Set rsPoliza = frsRegresaRs(strSentencia, adLockOptimistic, adOpenDynamic)
    
    strSentencia = "SELECT bitnaturalezamovimiento, intnumerocuenta, intnumeropoliza, intnumeroregistro, mnycantidadmovimiento, vchconcepto, vchreferencia " & _
                   "FROM   CNDETALLEPOLIZA " & _
                   "WHERE  intnumeropoliza = -1"
    Set rsDetallePoliza = frsRegresaRs(strSentencia, adLockOptimistic, adOpenDynamic)
    
    strSentencia = "SELECT * " & _
                   "FROM   CNPOLIZASIMPORTADAS " & _
                   "WHERE  CNPOLIZASIMPORTADAS.chrtipopoliza = '*' "
    Set rsPolizasImportadas = frsRegresaRs(strSentencia, adLockOptimistic, adOpenDynamic)
    
    strSentencia = "SELECT * " & _
                   "FROM   CNCFDIPOLIZA " & _
                   "WHERE  CNCFDIPOLIZA.intIdComprobante = -1 "
    Set rsUuid = frsRegresaRs(strSentencia, adLockOptimistic, adOpenDynamic)
       
    Me.MousePointer = 11 '|  Reloj de arena
    
    '|  Clave de la interfaz seleccionada para la importación
    If cboInterfazPolizas.ListIndex <> -1 Then
        lintInterfazImporta = cboInterfazPolizas.ItemData(cboInterfazPolizas.ListIndex)
    Else
        lintInterfazImporta = lintInterfazPolizas
    End If
    
    EntornoSIHO.ConeccionSIHO.BeginTrans
    
    '|  Bloquea el cierre
    lngResultado = 1
    vgstrParametrosSP = vgintClaveEmpresaContable & "|" & "Grabando poliza"
    frsEjecuta_SP vgstrParametrosSP, "sp_CnUpdEstatusCierre", True, lngResultado
    
    If lngResultado = 1 Then
        Set ExcelObj = CreateObject("Excel.Application")
        Set ExcelSheet = CreateObject("Excel.Sheet")
        
        ExcelObj.Workbooks.Open strArchivoImportacion 'Ruta del archivo
        Set ExcelBook = ExcelObj.Workbooks(1)
        Set ExcelSheet = ExcelBook.Worksheets(1)
        
        strTmpUuid = ""
        
        With ExcelSheet
            i = 2 'Como la primera línea es de encabezados no se toma en cuenta
                                    
            
            Do Until .Cells(i, 1) & "" = "" 'si la primer columna esta en blanco dejará de leer
            
                'Se obtienen los valores de las columnas con el formato adecuado
                strCuenta = .Cells(i, colCuenta)
                strDebito = .Cells(i, colDebito)
                strCredito = .Cells(i, colCredito)
                strTipo = .Cells(i, colTipo)
                strFecha = .Cells(i, colFecha)
                strConcepto = .Cells(i, colConcepto)
                strUUId = .Cells(i, colUuid)
                strTotal = .Cells(i, colTotal)
                strRFC = .Cells(i, colRfc)
                strNombre = .Cells(i, colNombre)
                
                'Se quitan espacios en blanco
                strCuenta = Trim(strCuenta)
                strDebito = Trim(strDebito)
                strCredito = Trim(strCredito)
                strTipo = Trim(strTipo)
                strFecha = Trim(strFecha)
                strConcepto = Trim(strConcepto)
                strUUId = Trim(strUUId)
                strTotal = Trim(strTotal)
                strRFC = Trim(strRFC)
                strNombre = Trim(strNombre)
                
            
                If i = 2 Then 'se guarda el primer registro en las tablas que no se repiten
                    
                    '--------------------------------------------
                    '|  Graba el maestro de la póliza
                    '--------------------------------------------
                    With rsPoliza
                        .AddNew
                        
                        !dtmFechaPoliza = strFecha
                        !chrTipoPoliza = strTipo
                        !vchConceptoPoliza = strConcepto
                        
                        '|  Indica si la póliza está incluida en un cierre contable ( guarda 1 = si está incluida, 0 = no está incluida)
                        !BITASENTADA = 0
                        '|  Empleado que esta realizando la importación
                        !intCveEmpleado = lngPersonaGraba
                        '|  Clave de la empresa contable (relacionado con CnEmpresaContable)
                        !TNYCLAVEEMPRESA = vgintClaveEmpresaContable
                        '|  Departamento del usuario que esta en el sistema
                        !smicvedepartamento = vgintNumeroDepartamento
                        '|  Año del ejercicio contable.
                        !SMIEJERCICIO = Format(!dtmFechaPoliza, "yyyy")
                        '|  Número de mes contable.
                        !tnyMes = Format(!dtmFechaPoliza, "mm")
                        '|  Número de póliza para manejo interno del área de contabilidad
                        !vchNumero = " "
                        '|  El número de póliza es generado por el sistema según este configurado
                        !intClavePoliza = flngFolioPoliza(vgintClaveEmpresaContable, !chrTipoPoliza, !SMIEJERCICIO, !tnyMes, False)
                        .Update
                        lngCvePolizaNueva = flngObtieneIdentity("SEC_CNPOLIZA", !intNumeroPoliza)
                    End With
                    '--------------------------------------------------------------------------------------
                    '|  Registra las pólizas que se están importando si previamente no se han registrado
                    '--------------------------------------------------------------------------------------
                    With rsPolizasImportadas
                        .AddNew
                        !chrTipoPoliza = strTipo
                        !dtmFechaPoliza = strFecha
                        '- Campos agregados para la relación de la póliza importada con la clave interna del SiHO -'
                        !intInterfazPoliza = lintInterfazImporta
                        !intNumeroPoliza = lngCvePolizaNueva
                        .Update
                    End With
                    
                End If 'Fin if primer registro
                
                '--------------------------------------------
                '|  Graba el detalle de la póliza
                '---------------------------------------------
                With rsDetallePoliza
                    .AddNew
                    !intNumeroCuenta = CLng(strCuenta)
                    !intNumeroPoliza = lngCvePolizaNueva
                    
                    !mnyCantidadMovimiento = IIf(strDebito = "0", CDbl(strCredito), CDbl(strDebito))
                    !bitNaturalezaMovimiento = IIf(strDebito <> "0", 1, 0)
                    !vchReferencia = ""
                    !vchConcepto = strConcepto
                    .Update
                End With
                
                'Se deberá insertar el uuid en caso de que sea diferente
                If strTmpUuid <> strUUId Then
                    With rsUuid
                        .AddNew
                        !INTNUMPOLIZA = lngCvePolizaNueva
                        !intNumCuenta = CLng(strCuenta)
                        !VCHUUID = strUUId
                        !numTotalComprobante = CDbl(strTotal)
                        !VCHNOMBRERECEPTOR = strNombre
                        !VCHRFCRECEPTOR = strRFC
                        .Update
                    End With
                    
                    strTmpUuid = strUUId
                    
                End If
                
                i = i + 1
                
            Loop
            
        End With
    
        ExcelObj.Workbooks.Close
    
        Set ExcelSheet = Nothing
        Set ExcelBook = Nothing
        Set ExcelObj = Nothing

        pEjecutaSentencia "update CnEstatusCierre set vchEstatus = 'Libre' where tnyClaveEmpresa = " & str(vgintClaveEmpresaContable)
                
        Call pGuardarLogTransaccion(Me.Name, EnmBorrar, lngPersonaGraba, "IMPORTACION DE POLIZAS", strArchivoImportacion)
    End If
    
    EntornoSIHO.ConeccionSIHO.CommitTrans
    
    'f.Close
    Me.MousePointer = 0
    If lngResultado <> 0 Then
        '|  La póliza ha sido importada exitosamente
        MsgBox SIHOMsg(744), vbOKOnly + vbInformation, "Mensaje"
        pCargaPolizas
    End If
    
Exit Sub
NotificaError:
    EntornoSIHO.ConeccionSIHO.RollbackTrans
    Me.MousePointer = 0
    Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":pImportaPolizaApsi " & Erl()))
    err.Clear
       
End Sub

'|  Verifica que el formato del archivo de texto sea el especificado por el layout de CONTPAQ,
'|  que las fechas de las polizas tengan cierre abierto y
'|  que las pólizas no hayan sido importadas anteriormente
Private Function fblnValidacionesGeneralesApsi(strArchivo As String) As Boolean
On Error GoTo NotificaError
    
    Dim rs As New ADODB.Recordset
    Dim fs As FileSystemObject
    Dim f As Variant
    Dim strErrorFalta As String
    Dim strPeriodosCerrados As String                   '|  Lista de los periodos cerrados que se informará se deben abrir antes de continuar
    Dim strEjercicio As String
    Dim strMes As String
    Dim strFechaPoliza As String
    Dim strTipoPoliza As String
    Dim strClavePoliza As String
    Dim blnFormatoErroneo As String
    Dim blnPolizaImportada As Boolean     'Indica si ya se importo la póliza
    
    fblnValidacionesGeneralesApsi = False
    strPeriodosCerrados = ""
    strErrorFalta = ""
    blnPolizaImportada = False
    blnFormatoErroneo = False
    'Variables para leer el archivo
    Dim strCuenta As String
    Dim strDebito As String
    Dim strCredito As String
    Dim strTipo As String
    Dim strFecha As String
    Dim strConcepto As String
    Dim strUUId As String
    Dim strTotal As String
    Dim strRFC As String
    Dim strNombre As String
    'Variables para abrir el excel
    Dim ExcelObj As Object
    Dim ExcelBook As Object
    Dim ExcelSheet As Object
    Dim i As Integer
    
    Dim strfinaliza As String
    'Si la extensión no es de un archivo de excel no debe continuar
    strfinaliza = Mid(strArchivo, (Len(strArchivo) - 4), 5)
    If InStr(strfinaliza, ".xls") = 0 Then
        fblnValidacionesGeneralesApsi = False
        blnFormatoErroneo = True
    Else
        fblnValidacionesGeneralesApsi = True
        blnFormatoErroneo = False
    End If
    
    'Si el formato no es válido no tiene sentido abrir el archivo
    If blnFormatoErroneo = False Then

        Set ExcelObj = CreateObject("Excel.Application")
        Set ExcelSheet = CreateObject("Excel.Sheet")
    
        ExcelObj.Workbooks.Open strArchivo 'path donde esta el archivo
    
        Set ExcelBook = ExcelObj.Workbooks(1)
        Set ExcelSheet = ExcelBook.Worksheets(1)
        
        With ExcelSheet
        i = 1
        Do Until .Cells(i, 1) & "" = "" ' si no hay nada se sale
                    
            strCuenta = .Cells(i, colCuenta)
            strDebito = .Cells(i, colDebito)
            strCredito = .Cells(i, colCredito)
            strTipo = .Cells(i, colTipo)
            strFecha = .Cells(i, colFecha)
            strConcepto = .Cells(i, colConcepto)
            strUUId = .Cells(i, colUuid)
            strTotal = .Cells(i, colTotal)
            strRFC = .Cells(i, colRfc)
            strNombre = .Cells(i, colNombre)
                
            '---------------------------------------------------------------------------------------------'
            '|  Verifica que el encabezado del archivo coincida con el formato especificado por APSI  |'
            '---------------------------------------------------------------------------------------------'
            If i = 1 Then
                'si no tiene el formato correcto se sale
                If strCuenta = "" _
                   Or strDebito = "" _
                   Or strCredito = "" _
                   Or strTipo = "" _
                   Or strFecha = "" _
                   Or strConcepto = "" _
                   Or strUUId = "" _
                   Or strTotal = "" _
                   Or strRFC = "" _
                   Or strNombre = "" Then
                   blnFormatoErroneo = True
                   Exit Do
                End If
                        
            Else
                'Se valida que esten todas las columnas con la información requerida ya que puede provocar errores
                If strCuenta = "" Or strDebito = "" Or strCredito = "" Or strTipo = "" Or strFecha = "" Or strConcepto = "" Or strUUId = "" Or strTotal = "" Or strRFC = "" Or strNombre = "" Then
                    
                    If strCuenta = "" Then
                        strErrorFalta = strErrorFalta & "Cuenta, "
                    End If
                    If strDebito = "" Then
                        strErrorFalta = strErrorFalta & "Débito, "
                    End If
                    If strCredito = "" Then
                        strErrorFalta = strErrorFalta & "Crédito, "
                    End If
                    If strTipo = "" Then
                        strErrorFalta = strErrorFalta & "Tipo, "
                    End If
                    If strFecha = "" Then
                        strErrorFalta = strErrorFalta & "Fecha, "
                    End If
                    If strConcepto = "" Then
                        strErrorFalta = strErrorFalta & "Concepto, "
                    End If
                    If strUUId = "" Then
                        strErrorFalta = strErrorFalta & "UUID, "
                    End If
                    If strTotal = "" Then
                        strErrorFalta = strErrorFalta & "Total comprobante, "
                    End If
                    If strRFC = "" Then
                        strErrorFalta = strErrorFalta & "RFC, "
                    End If
                    If strNombre = "" Then
                        strErrorFalta = strErrorFalta & "Nombre receptor, "
                    End If
                    strErrorFalta = Left(strErrorFalta, Len(strErrorFalta) - 2)
                    
                    Exit Do
                End If
            
                If i = 2 Then
                    '-------------------------------------------------------------------------------------------------'
                    '|  Verifica que el periodo contable de la(s) póliza(s) que se intenta(n) importar esté abierto  |'
                    '|  de lo contrario lo guarda en una cadena para desplegarlo posteriormente                      |'
                    '-------------------------------------------------------------------------------------------------'
                    strFechaPoliza = .Cells(i, colFecha) 'vendra en formato DD/MM/YYYY

                    strEjercicio = Mid(strFechaPoliza, 7, 4)
                    strMes = Mid(strFechaPoliza, 4, 2)
                    If fblnPeriodoCerrado(vgintClaveEmpresaContable, CInt(strEjercicio), CInt(strMes)) Then '|  fblnPeriodoCerrado(CveEmpresa, Año, Mes)
                        '|  Si no se ha agregado a la cadena strPeriodosCerrados
                        If Len(Replace(strPeriodosCerrados, fstrMesLetras(CInt(strMes)) & " " & strEjercicio, "")) = _
                            Len(strPeriodosCerrados) Then
                            strPeriodosCerrados = strPeriodosCerrados & vbCrLf & fstrMesLetras(CInt(strMes)) & " " & strEjercicio
                        End If
                    End If


                    '----------------------------------------------------------------------------------------------'
                    '|  Verifica que las pólizas que se intentan importar no hayan sido importadas anteriormente  |'
                    '|  de lo contrario las guarda en una cadena para desplegarlas posteriormente                 |'
                    '----------------------------------------------------------------------------------------------'
                    Set rs = frsRegresaRs("select count(*) val from CNPOLIZA pol where trim(pol.VCHCONCEPTOPOLIZA) = '" & Trim(strConcepto) & "'")
                    If rs!Val > 0 Then
                        blnPolizaImportada = True
                    End If

                    'Exit Do
                Else
                    'Exit Do
                End If
            End If
            
            i = i + 1
        Loop
    
        End With
    
        ExcelObj.Workbooks.Close
    
        Set ExcelSheet = Nothing
        Set ExcelBook = Nothing
        Set ExcelObj = Nothing
    
    End If
    
    If strPeriodosCerrados <> "" Then
        fblnValidacionesGeneralesApsi = False
        '|  El periodo contable esta cerrado.
        MsgBox SIHOMsg(209) & vbCrLf & "No se puede continuar." & vbCrLf & strPeriodosCerrados, vbCritical, "Mensaje"
        Exit Function
    End If
    
    If blnPolizaImportada Then
        fblnValidacionesGeneralesApsi = False
        
        '|  La póliza que desea importar, ya se encuentra registrada.
        MsgBox "La póliza que desea importar, ya se encuentra registrada.", vbCritical, "Mensaje"
        Exit Function
        
    End If
    
    If blnFormatoErroneo Then
        fblnValidacionesGeneralesApsi = False
        
        '|  Formato interno de archivo erróneo.
        MsgBox SIHOMsg(746), vbCritical, "Mensaje"
        Exit Function
    End If
    
    If strErrorFalta <> "" Then
        fblnValidacionesGeneralesApsi = False
        
        '|  Las siguientes columnas del archivo no contienen información:
        MsgBox "Las siguientes columnas del archivo no contienen información: " & strErrorFalta, vbCritical, "Mensaje"
        Exit Function
    End If
       
    fblnValidacionesGeneralesApsi = True
    
Exit Function
NotificaError:
    Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":fblnValidacionesGeneralesApsi(" & strArchivo & ")"))
End Function


Private Function fblnCuentaCuadre(llngNumPoliza As Long) As Boolean
On Error GoTo NotificaError

    Dim vllngCont As Long
    Dim rsPoliza As New ADODB.Recordset
    Dim vlstrSentencia As String

    vlstrSentencia = "SELECT DP.intNumeroPoliza, C.vchCuentaContable " & _
                     "FROM CnDetallePoliza DP, CnCuenta C " & _
                     "WHERE DP.intNumeroCuenta = C.intNumeroCuenta " & _
                     "AND DP.intNumeroPoliza = " & llngNumPoliza
    Set rsPoliza = frsRegresaRs(vlstrSentencia)
    If rsPoliza.RecordCount <> 0 Then
        rsPoliza.MoveFirst
        While Not rsPoliza.EOF
            If Trim(rsPoliza!VchCuentaContable) = "*" Then
                fblnCuentaCuadre = True
                Exit Function
            End If
            rsPoliza.MoveNext
        Wend
    End If
        
    fblnCuentaCuadre = False
    
Exit Function
NotificaError:
    Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":fblnCuentaCuadre"))
End Function

Private Function fblnPolizaCancelada(llngNumPoliza As Long) As Boolean
On Error GoTo NotificaError
    
    Dim vllngCont As Long
    Dim rsPoliza As New ADODB.Recordset
    Dim vlstrSentencia As String

    vlstrSentencia = "SELECT DP.intNumeroPoliza FROM CnDetallePoliza DP WHERE DP.intNumeroPoliza = " & llngNumPoliza
    Set rsPoliza = frsRegresaRs(vlstrSentencia)
    If rsPoliza.RecordCount = 0 Then
        fblnPolizaCancelada = True
    Else
        fblnPolizaCancelada = False
    End If
    
Exit Function
NotificaError:
    Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":fblnPolizaCancelada"))
End Function
Private Function fblnPolizaDescuadrada(llngNumPoliza As Long) As Boolean
On Error GoTo NotificaError

    Dim vllngCont As Long
    Dim rsPoliza As New ADODB.Recordset
    Dim vlstrSentencia As String
    fblnPolizaDescuadrada = False
 
    vlstrSentencia = "Select poliza.*, poliza.Cargo - poliza.Abono diferencia From (Select Cndetallepoliza.Intnumeropoliza, Cnpoliza.Vchconceptopoliza, Cnpoliza.Dtmfechapoliza, Sum(Case When Cndetallepoliza.Bitnaturalezamovimiento = 1 Then Round(Cndetallepoliza.Mnycantidadmovimiento, 2) Else 0 End) Cargo, Sum(Case When Cndetallepoliza.Bitnaturalezamovimiento = 0 Then Round(Cndetallepoliza.Mnycantidadmovimiento, 2)  Else 0 End) Abono From Cndetallepoliza Inner Join Cnpoliza On Cnpoliza.Intnumeropoliza = Cndetallepoliza.Intnumeropoliza where Cnpoliza.Intnumeropoliza = " & llngNumPoliza & " Group by Cndetallepoliza.Intnumeropoliza, Cnpoliza.Vchconceptopoliza,Cnpoliza.Dtmfechapoliza) poliza where Cargo - Abono > 0 "
    Set rsPoliza = frsRegresaRs(vlstrSentencia)
    If rsPoliza.RecordCount <> 0 Then
        fblnPolizaDescuadrada = True
    Else
        fblnPolizaDescuadrada = False
    End If


Exit Function
NotificaError:
    Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":fblnPolizaDescuadrada"))
End Function

'|  Da formato a una cuenta según la mascara definida en la base de datos
Private Function fintFormatoCuentaMicrosip(lstrCuenta As String, intEmpresa As Integer) As Long
    Dim strCuentaFormateada As String
    Dim strMascaraCuenta As String
    Dim intContMascara As Integer
    Dim rsCuenta As New ADODB.Recordset
    Dim intContCuenta As Integer
    Dim rs As ADODB.Recordset
    Dim aCuenta() As String
    Dim aEstructura() As String
    
On Error GoTo NotificaError

    strCuentaFormateada = ""
    Set rs = frsSelParametros("CN", intEmpresa, "VCHESTRUCTURACUENTACONTABLE")
    If Not rs.EOF Then
        If IsNull(rs!Valor) Then
            strMascaraCuenta = -1
        Else
            strMascaraCuenta = rs!Valor
        End If
    Else
        strMascaraCuenta = -1
    End If
    rs.Close
    
    '|  Si no se he configurado una mascara de cuenta
    If strMascaraCuenta = "-1" Then
        '|  No existe registrada la estructura de cuentas contables.
        MsgBox SIHOMsg(246), vbCritical, "Mensaje"
        fintFormatoCuentaMicrosip = -1
        Exit Function
    End If
    
    '|  Crea el arreglo con la cuenta de la póliza
    ReDim aCuenta(0)
    aCuenta = Split(lstrCuenta, ".")
    
    '|  Crea el arreglo de referencia con la estructura de cuentas de la empresa
    ReDim aEstructura(0)
    aEstructura = Split(strMascaraCuenta, ".")
    
    '|  Si la estructura de la cuenta es menor, agregar los elementos faltantes al final del arreglo
    If UBound(aCuenta) < UBound(aEstructura) Then
        For intContMascara = UBound(aCuenta) To UBound(aEstructura)
            ReDim Preserve aCuenta(UBound(aCuenta) + 1)
            aCuenta(UBound(aCuenta)) = "0"
        Next intContMascara
    ElseIf UBound(aCuenta) > UBound(aEstructura) Then
        '|  ¡El formato de la cuenta es incorrecto!
        MsgBox SIHOMsg(66), vbCritical, "Mensaje"
        fintFormatoCuentaMicrosip = -1
        Exit Function
    End If
    
    '|  Dar formato a la cuenta contable con la información del arreglo de la cuenta
    For intContMascara = 0 To UBound(aEstructura)
        If Len(aCuenta(intContMascara)) > Len(aEstructura(intContMascara)) Then
            '|  ¡El formato de la cuenta es incorrecto!
            MsgBox SIHOMsg(66) & Chr(13) & "La siguiente cuenta no tiene el formato correcto: " & lstrCuenta, vbCritical, "Mensaje"
            fintFormatoCuentaMicrosip = -1
            Exit Function
        Else
            strCuentaFormateada = strCuentaFormateada & String(Len(aEstructura(intContMascara)) - Len(aCuenta(intContMascara)), "0") & aCuenta(intContMascara) & "."
        End If
    Next intContMascara
    strCuentaFormateada = Left(strCuentaFormateada, Len(strCuentaFormateada) - 1) '|  Borrar el último "."
    
    vlstrsql = "SELECT intNumeroCuenta FROM CnCuenta " & _
               " WHERE vchCuentaContable = '" & strCuentaFormateada & "' " & _
               " AND tnyClaveEmpresa = " & vgintClaveEmpresaContable & " AND bitEstatusActiva = 1"
    Set rsCuenta = frsRegresaRs(vlstrsql, adLockOptimistic, adOpenDynamic)
    If rsCuenta.RecordCount > 0 Then
        fintFormatoCuentaMicrosip = rsCuenta!intNumeroCuenta
    Else
        '|  ¡La cuenta no existe!
        MsgBox SIHOMsg(67) & vbCrLf & vbCrLf & strCuentaFormateada, vbCritical, "Mensaje"
        fintFormatoCuentaMicrosip = -1
    End If
    
Exit Function
NotificaError:
    Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":fintFormateaCuenta(" & lstrCuenta & ")"))
End Function

'|  Convierte el número de cuenta al formato de Microsip
Private Function fstrFormateaCuentaMicrosip(lstrCuenta As String) As String
On Error GoTo NotificaError

    Dim aCuenta() As String
    Dim lstrCuentaFormateada As String
    Dim lintCont As Integer
    
    fstrFormateaCuentaMicrosip = ""
    
    If Trim(lstrCuenta) = "" Then Exit Function
    
    aCuenta = Split(lstrCuenta, ".")
    lstrCuentaFormateada = ""
    For lintCont = 0 To UBound(aCuenta)
        lstrCuentaFormateada = lstrCuentaFormateada & IIf(CLng(aCuenta(lintCont)) > 0, CLng(aCuenta(lintCont)) & ".", "")
    Next lintCont
    lstrCuentaFormateada = Left(lstrCuentaFormateada, Len(lstrCuentaFormateada) - 1) '|  Borrar el último "."
    
    fstrFormateaCuentaMicrosip = lstrCuentaFormateada

Exit Function
NotificaError:
    Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":fstrFormateaCuentaMicrosip(" & lstrCuenta & ")"))
End Function

'|  Verifica que el formato del archivo de texto sea el especificado por el layout de MICROSIP,
'|  que las fechas de las polizas tengan cierre abierto y que las pólizas no hayan sido importadas anteriormente
Private Function fblnValidacionesGeneralesMicrosip(strArchivo As String) As Boolean
On Error GoTo NotificaError

    Dim fs As FileSystemObject
    Dim f As Variant
    Dim strLinea As String
    Dim strPeriodosCerrados As String                   '|  Lista de los periodos cerrados que se informará se deben abrir antes de continuar
    Dim strEjercicio As String
    Dim strMes As String
    Dim strFechaPoliza As String
    Dim strTipoPoliza As String
    Dim strClavePoliza As String
    Dim blnFormatoErroneo As String
    Dim strPolizasImportadasAnteriormente As String     '|  Lista de las pólizas que ya han sido importadas con anterioridad
    Dim tArreglo() As String

    fblnValidacionesGeneralesMicrosip = False
    
    strPeriodosCerrados = ""
    strPolizasImportadasAnteriormente = ""
    blnFormatoErroneo = False
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.OpenTextFile(strArchivo, ForReading, TristateFalse)
    
    '--------------------------------------------------------------------'
    '|  Recorre el archivo para validar los encabezados de las pólizas  |'
    '--------------------------------------------------------------------'
    Do Until f.AtEndOfStream
        DoEvents
        ReDim tArreglo(0)
        strLinea = f.ReadLine
        strLinea = Mid(strLinea, 2, Len(strLinea)) '|  Quitar el primer elemento (pipe)
        strLinea = Replace(strLinea, """", "") '|  Sustituir las comillas dobles
        strLinea = Replace(strLinea, "|", ",") '|  Sustituir los pipes por comas para que se dividan los datos
        tArreglo = Split(strLinea, ",") '|  Convertir los elementos a arreglo para un manejo más fácil
        If tArreglo(intPosNivel) = "1" Then
            '----------------------------------------------------------------------------------------------'
            '|  Verifica que el encabezado del archivo coincida con el formato especificado por MICROSIP  |'
            '----------------------------------------------------------------------------------------------'
            If Len(tArreglo(intPosNivel)) <> 1 Or fstrVerificaFecha(tArreglo(intPosFecha)) = "" Or Len(tArreglo(intPosTipo)) <> 1 Then
                blnFormatoErroneo = True
                Exit Do
            End If
            
            '-------------------------------------------------------------------------------------------------'
            '|  Verifica que el periodo contable de la(s) póliza(s) que se intenta(n) importar esté abierto  |'
            '|  de lo contrario lo guarda en una cadena para desplegarlo posteriormente                      |'
            '-------------------------------------------------------------------------------------------------'
            strEjercicio = Format(CDate(tArreglo(intPosFecha)), "yyyy")
            strMes = Format(CDate(tArreglo(intPosFecha)), "mm")
            If fblnPeriodoCerrado(vgintClaveEmpresaContable, CInt(strEjercicio), CInt(strMes)) Then '|  fblnPeriodoCerrado(CveEmpresa, Año, Mes)
                '|  Si no se ha agregado a la cadena strPeriodosCerrados
                If Len(Replace(strPeriodosCerrados, fstrMesLetras(CInt(strMes)) & " " & strEjercicio, "")) = Len(strPeriodosCerrados) Then
                    strPeriodosCerrados = strPeriodosCerrados & vbCrLf & fstrMesLetras(CInt(strMes)) & " " & strEjercicio
                End If
            End If
            
            '----------------------------------------------------------------------------------------------'
            '|  Verifica que las pólizas que se intentan importar no hayan sido importadas anteriormente  |'
            '|  de lo contrario las guarda en una cadena para desplegarlas posteriormente                 |'
            '----------------------------------------------------------------------------------------------'
            strFechaPoliza = tArreglo(intPosFecha)
            strTipoPoliza = tArreglo(intPosTipo)
            strClavePoliza = "" 'tArreglo(intPosPoliza)
            If fblnPolizaImportada(strFechaPoliza, strTipoPoliza, strClavePoliza, MICROSIP) Then
                '|  Si no se ha agregado a la cadena strPolizasImportadasAnteriormente
                If Len(Replace(strPolizasImportadasAnteriormente, strFechaPoliza & " " & IIf(strTipoPoliza = "I", "Ingreso", IIf(strTipoPoliza = "E", "Egreso", IIf(strTipoPoliza = "D", "Diario", "Desconocido"))), "")) = Len(strPolizasImportadasAnteriormente) Then
                    strPolizasImportadasAnteriormente = strPolizasImportadasAnteriormente & vbCrLf & strFechaPoliza & " " & IIf(strTipoPoliza = "I", "Ingreso", IIf(strTipoPoliza = "E", "Egreso", IIf(strTipoPoliza = "D", "Diario", "Desconocido")))
                End If
            End If
        ElseIf UBound(tArreglo) = 0 Then '|  Si el arreglo esta vacío el formato es incorrecto
            blnFormatoErroneo = True
        End If
    Loop
    f.Close
    
    If strPeriodosCerrados <> "" Then
        '|  El periodo contable esta cerrado.
        MsgBox SIHOMsg(209) & vbCrLf & "No se puede continuar." & vbCrLf & strPeriodosCerrados, vbCritical, "Mensaje"
        Exit Function
    End If
    
    If strPolizasImportadasAnteriormente <> "" Then
        '|  Ya han sido importadas pólizas con las siguientes fechas y tipos. ¿Desea continuar?
        If MsgBox(SIHOMsg(745) & strPolizasImportadasAnteriormente, vbYesNo + vbCritical, "Mensaje") = vbNo Then
            Exit Function
        End If
    End If
    
    If blnFormatoErroneo Then
        '|  Formato interno de archivo erróneo.
        MsgBox SIHOMsg(746), vbCritical, "Mensaje"
        Exit Function
    End If

    fblnValidacionesGeneralesMicrosip = True
    
Exit Function
NotificaError:
    Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":fblnValidacionesGeneralesMicrosip(" & strArchivo & ")"))
End Function

'|  Regresa la interfaz que está utilizando la empresa contable actual  |'
Private Function fintTipoInterfazPoliza(intEmpresa As Integer) As Integer
On Error GoTo NotificaError

    Set rs = frsSelParametros("SI", intEmpresa, "INTINTERFAZPOLIZA")
    If Not rs.EOF Then
        If IsNull(rs!Valor) Then
            fintTipoInterfazPoliza = -1
        Else
            fintTipoInterfazPoliza = Val(rs!Valor)
        End If
    Else
        fintTipoInterfazPoliza = -1
    End If
    rs.Close
    
Exit Function
NotificaError:
    Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":fintTipoInterfazPoliza"))
End Function

Private Sub pAgregarInterfacesPolizas()
On Error GoTo NotificaError

    Dim rs As New ADODB.Recordset
    
    cboInterfazPolizas.Clear

    Set rs = frsEjecuta_SP(CStr(vgintClaveEmpresaContable), "Sp_GnSelInterfazEmpresa")
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        Do While Not rs.EOF
            If rs!clave = CONTPAQ Then
                If fblnValidarLicenciaPolizas(vgintClaveEmpresaContable, rs!clave, "CONTPAQ") Then
                    cboInterfazPolizas.AddItem "CONTPAQ"
                    cboInterfazPolizas.ItemData(cboInterfazPolizas.NewIndex) = rs!clave
                End If
            End If
            
            If rs!clave = MICROSIP Then
                If fblnValidarLicenciaPolizas(vgintClaveEmpresaContable, rs!clave, "Microsip") Then
                    cboInterfazPolizas.AddItem "Microsip"
                    cboInterfazPolizas.ItemData(cboInterfazPolizas.NewIndex) = rs!clave
                End If
            End If
            
            If rs!clave = APSI Then
                If fblnValidarLicenciaPolizas(vgintClaveEmpresaContable, rs!clave, "APSI") Then
                    cboInterfazPolizas.AddItem "APSI"
                    cboInterfazPolizas.ItemData(cboInterfazPolizas.NewIndex) = rs!clave
                End If
            End If
            
            rs.MoveNext
        Loop
        
        If cboInterfazPolizas.ListCount > 0 Then 'And lintInterfazPolizas <> 0 Then
            cboInterfazPolizas.ListIndex = fintLocalizaCbo(cboInterfazPolizas, CStr(lintInterfazPolizas))
            If cboInterfazPolizas.ListIndex < 0 Then cboInterfazPolizas.ListIndex = 0
        End If
        
'        If cboInterfazPolizas.ListIndex < 0 And lintInterfazPolizas <> 0 Then
'            cboInterfazPolizas.AddItem "<NINGUNO>", 0
'            cboInterfazPolizas.ItemData(0) = 0
'            cboInterfazPolizas.ListIndex = 0
'        End If
    Else
        cboInterfazPolizas.AddItem "<NINGUNO>"
        cboInterfazPolizas.ItemData(cboInterfazPolizas.NewIndex) = 0
        cboInterfazPolizas.ListIndex = 0
    End If
    rs.Close
    
    lblInterfaz.Visible = (cboInterfazPolizas.ListCount > 1)
    cboInterfazPolizas.Visible = (cboInterfazPolizas.ListCount > 1)
    
Exit Sub
NotificaError:
    Call pRegistraError(err.Number, err.Description, cgstrModulo, (Me.Name & ":pAgregarInterfacesPolizas"))
End Sub

'|  Verifica que se tenga la licencia para la importación/exportación de pólizas
Private Function fblnValidarLicenciaPolizas(lintClaveEmpresa As Integer, lintClaveInterfaz As Integer, lstrInterfaz As String) As Boolean
On Error GoTo NotificaError

    Dim rs As New ADODB.Recordset
    Dim lstrEncriptado As String
    Dim lstrRFC As String
    Dim lstrClave As String
    Dim lstrsql As String
    Dim lblnLicenciamiento As Boolean
    
    fblnValidarLicenciaPolizas = False

    ' Buscar el RFC de la empresa para formar la licencia
    lstrsql = "SELECT TRIM(REPLACE(REPLACE(REPLACE(CNEMPRESACONTABLE.VCHRFC,'-',''),'_',''),' ','')) AS RFC " & _
              " FROM CNEMPRESACONTABLE  " & _
              " WHERE CNEMPRESACONTABLE.TNYCLAVEEMPRESA = " & lintClaveEmpresa
    Set rs = frsRegresaRs(lstrsql)
    If Not rs.EOF Then
        lstrRFC = Trim(rs!RFC)
    End If
    rs.Close
    
    If Trim(lstrRFC) = "" Then
        Exit Function
    End If

    lblnLicenciamiento = False
        
    vgstrParametrosSP = CStr(vgintClaveEmpresaContable)
    Set rs = frsEjecuta_SP(vgstrParametrosSP, "Sp_GnSelInterfazEmpresa")
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        Do While Not rs.EOF
            If rs!clave = lintClaveInterfaz Then
                lstrClave = "Dx50Dg3aR612" & lstrRFC & lstrInterfaz & CStr(lintClaveInterfaz) & "0D1Gn6R12x"
                
                'Se decodifica la licencia a partir de lstrClave (1a vez)
                lstrEncriptado = Encode(Trim(lstrClave))
                
                'Se decodifica la licencia a partir de lstrClave (2a vez)
                lstrEncriptado = Encode(lstrEncriptado)
                
                'Se reemplazan los caracteres especiales ("U"<-"?"   "l"<-"ñ"   "="<-"Ñ"   "=="<-"Ñ?")
                lstrEncriptado = Replace(lstrEncriptado, "U", "?")
                lstrEncriptado = Replace(lstrEncriptado, "l", "ñ")
                lstrEncriptado = Replace(lstrEncriptado, "==", "Ñ?")
                lstrEncriptado = Replace(lstrEncriptado, "=", "Ñ")
                lstrEncriptado = Replace(Replace(Trim(lstrEncriptado), Chr(10), ""), Chr(13), "") 'Elimina los saltos de linea
                
                lblnLicenciamiento = IIf(rs!Licencia = lstrEncriptado, True, False)
            End If
            rs.MoveNext
        Loop
    End If

    fblnValidarLicenciaPolizas = lblnLicenciamiento

Exit Function
NotificaError:
    Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":fblnValidarLicenciaPolizas"))
End Function

Private Function fblnMuestraBarraProgreso(vlMostrar As Boolean) As Boolean
On Error GoTo NotificaError

    freBarra.Left = IIf(vlMostrar, 2575, 2575)
    freBarra.Top = IIf(vlMostrar, 3575, 8040)

Exit Function
NotificaError:
    Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":fblnMuestraBarraProgreso"))
End Function



Private Sub txtFolioFin_Click()
    If txtFolioFin.Text <> "" Then
        txtFolioFin.SelLength = txtFolioFin.Text
    End If
End Sub

Private Sub txtFolioFin_GotFocus()
   If txtFolioFin.Text <> "" Then
         pSelTextBox txtFolioFin
    End If
End Sub

Private Sub txtFolioFin_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Or KeyAscii = 46 Or KeyAscii = 32 Then KeyAscii = 7
   
    If KeyAscii = 13 Then
        If txtFolioIni.Text = "" Then
            txtFolioIni.SetFocus
        Else
        pCargaPolizas
        grdPolizas1.SetFocus
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":mskFechaFin_KeyPress"))
End Sub

Private Sub txtFolioFin_LostFocus()
On Error GoTo NotificaError
    
    If Val(txtFolioIni.Text) > Val(txtFolioFin.Text) Then
        '¡Rango no válido!
        MsgBox SIHOMsg(26), vbOKOnly + vbInformation, "Mensaje"
        txtFolioIni.SetFocus
    Else
        pCargaPolizas
        grdPolizas1.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":mskFechaFin_LostFocus"))
End Sub


Private Sub txtFolioIni_Click()
    If txtFolioIni.Text <> "" Then
        txtFolioIni.SelLength = txtFolioIni.Text
    End If
End Sub

Private Sub txtFolioIni_GotFocus()
   If txtFolioIni.Text <> "" Then
         pSelTextBox txtFolioIni
    End If
End Sub

Private Sub txtFolioIni_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Or KeyAscii = 46 Or KeyAscii = 32 Then KeyAscii = 7

    If KeyAscii = 13 Or KeyAscii = 11 Then
        If txtFolioIni.Text = "" Then
            txtFolioIni.SetFocus
        Else
         txtFolioFin.SetFocus
        End If
        
    End If
Exit Sub
NotificaError:
    Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":mskFechaInicio_KeyPress"))
End Sub

Private Sub pLimpia()
    txtFolioIni.Text = ""
    txtFolioFin.Text = ""
End Sub
