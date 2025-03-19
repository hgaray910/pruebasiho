VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRptResumenDiarioCargos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resumen diario de cargos y formas de pago"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8130
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   8130
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraSoloCargos 
      Enabled         =   0   'False
      Height          =   690
      Left            =   120
      TabIndex        =   30
      Top             =   3600
      Width           =   1845
      Begin VB.CheckBox ChkIncluirCargos 
         Caption         =   "Incluir sólo cargos"
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   10
         ToolTipText     =   "Incluye sólo los cargos al reporte"
         Top             =   240
         Width           =   1635
      End
   End
   Begin VB.Frame Frame1 
      Height          =   690
      Left            =   2780
      TabIndex        =   29
      Top             =   3570
      Width           =   2570
      Begin VB.CommandButton cmdImprimir 
         Height          =   495
         Left            =   560
         Picture         =   "frmRptResumenDiarioCargos.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Imprimir"
         Top             =   140
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdVistaPreliminar 
         Height          =   495
         Left            =   50
         Picture         =   "frmRptResumenDiarioCargos.frx":01A2
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Vista previa"
         Top             =   140
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdExportar 
         Caption         =   "Exportar a Excel"
         Height          =   495
         Left            =   1080
         TabIndex        =   13
         ToolTipText     =   "Exportar a Excel"
         Top             =   140
         Width           =   1455
      End
   End
   Begin VB.Frame freBarra 
      Height          =   815
      Left            =   1833
      TabIndex        =   25
      Top             =   1835
      Visible         =   0   'False
      Width           =   4465
      Begin MSComctlLib.ProgressBar pgbBarra 
         Height          =   300
         Left            =   45
         TabIndex        =   26
         Top             =   480
         Width           =   4380
         _ExtentX        =   7726
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label1 
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
         Width           =   4365
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000002&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   345
         Left            =   45
         Top             =   135
         Width           =   4380
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
         TabIndex        =   27
         Top             =   180
         Width           =   4370
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Fecha "
      Height          =   795
      Left            =   128
      TabIndex        =   22
      Top             =   1680
      Width           =   3795
      Begin MSMask.MaskEdBox mskFecha 
         Height          =   315
         Left            =   480
         TabIndex        =   5
         ToolTipText     =   "Fecha del"
         Top             =   280
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mmm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskFechaFin 
         Height          =   315
         Left            =   2280
         TabIndex        =   6
         ToolTipText     =   "Fecha al"
         Top             =   280
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mmm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Del"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   340
         Width           =   240
      End
      Begin VB.Label lblAl 
         AutoSize        =   -1  'True
         Caption         =   "Al"
         Height          =   195
         Left            =   1920
         TabIndex        =   23
         Top             =   340
         Width           =   135
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Periodo "
      Height          =   795
      Left            =   4088
      TabIndex        =   19
      Top             =   1680
      Width           =   3915
      Begin VB.ComboBox cboEjercicio 
         Height          =   315
         ItemData        =   "frmRptResumenDiarioCargos.frx":0344
         Left            =   840
         List            =   "frmRptResumenDiarioCargos.frx":0346
         Style           =   2  'Dropdown List
         TabIndex        =   7
         ToolTipText     =   "Selección del ejercicio"
         Top             =   280
         Width           =   1080
      End
      Begin VB.ComboBox cboMes 
         Height          =   315
         ItemData        =   "frmRptResumenDiarioCargos.frx":0348
         Left            =   2490
         List            =   "frmRptResumenDiarioCargos.frx":034A
         Style           =   2  'Dropdown List
         TabIndex        =   8
         ToolTipText     =   "Selección del mes"
         Top             =   280
         Width           =   1320
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ejercicio"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   340
         Width           =   600
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Mes"
         Height          =   195
         Left            =   2040
         TabIndex        =   20
         Top             =   345
         Width           =   300
      End
   End
   Begin VB.Frame Frame2 
      Height          =   795
      Left            =   128
      TabIndex        =   17
      Top             =   2640
      Width           =   7875
      Begin VB.ComboBox cboConcepto 
         Height          =   315
         Left            =   1680
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   9
         ToolTipText     =   "Concepto de factura"
         Top             =   280
         Width           =   6090
      End
      Begin VB.Label lblConcepto 
         AutoSize        =   -1  'True
         Caption         =   "Concepto de factura"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   340
         Width           =   1455
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Tipo de reporte "
      Height          =   825
      Left            =   128
      TabIndex        =   16
      Top             =   750
      Width           =   7875
      Begin VB.OptionButton optPresentacion 
         Caption         =   "Agrupado por día"
         Height          =   315
         Index           =   3
         Left            =   6120
         TabIndex        =   4
         ToolTipText     =   "Reporte agrupado por día"
         Top             =   330
         Width           =   1635
      End
      Begin VB.OptionButton optPresentacion 
         Caption         =   "Desglosado por cargo"
         Height          =   315
         Index           =   2
         Left            =   4100
         TabIndex        =   3
         ToolTipText     =   "Reporte desglosado por cargo"
         Top             =   330
         Width           =   1875
      End
      Begin VB.OptionButton optPresentacion 
         Caption         =   "Agrupado por cargo"
         Height          =   315
         Index           =   1
         Left            =   2240
         TabIndex        =   2
         ToolTipText     =   "Reporte agrupado por cargo"
         Top             =   330
         Width           =   1815
      End
      Begin VB.OptionButton optPresentacion 
         Caption         =   "Agrupado por concepto"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "Reporte agrupado por concepto"
         Top             =   330
         Value           =   -1  'True
         Width           =   2055
      End
   End
   Begin VB.Frame Frame5 
      Height          =   700
      Left            =   128
      TabIndex        =   14
      Top             =   0
      Width           =   7875
      Begin VB.ComboBox cboHospital 
         Height          =   315
         Left            =   915
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Seleccione la empresa"
         Top             =   240
         Width           =   6855
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Empresa"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   300
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmRptResumenDiarioCargos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private vgrptReporte As CRAXDRT.Report

Private Sub cmdExportar_Click()
    pExporta 'Exportación a excel
End Sub

Private Sub cmdImprimir_Click()
    pImprime "I"
End Sub

Private Sub cmdVistaPreliminar_Click()
    pImprime "P"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo NotificaError
    
    If KeyCode = vbKeyEscape Then
        Unload Me
        KeyCode = 0
    ElseIf KeyCode = vbKeyReturn Then
        SendKeys vbTab
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyDown"))
End Sub

Private Sub Form_Load()
On Error GoTo NotificaError

    Dim vlintEjercicios As Integer
    Dim vlfechaServer As Date
    Dim vlintIndexCombos As Integer

    pCargaParametrosContabilidad vgintClaveEmpresaContable

    Me.Icon = frmMenuPrincipal.Icon
    vgstrNombreForm = Me.Name
    
    vlfechaServer = fdtmServerFecha
    
    cboMes.Clear
    cboMes.AddItem "Enero", 0
    cboMes.AddItem "Febrero", 1
    cboMes.AddItem "Marzo", 2
    cboMes.AddItem "Abril", 3
    cboMes.AddItem "Mayo", 4
    cboMes.AddItem "Junio", 5
    cboMes.AddItem "Julio", 6
    cboMes.AddItem "Agosto", 7
    cboMes.AddItem "Septiembre", 8
    cboMes.AddItem "Octubre", 9
    cboMes.AddItem "Noviembre", 10
    cboMes.AddItem "Diciembre", 11
    
    cboMes.ListIndex = Month(vlfechaServer) - 1
    
    cboEjercicio.Clear
    For vlintEjercicios = vgintEjercicioInicioOperaciones To Year(vlfechaServer)
        cboEjercicio.AddItem Trim(Str(vlintEjercicios)), vlintIndexCombos
        vlintIndexCombos = vlintIndexCombos + 1
    Next
    cboEjercicio.ListIndex = vlintIndexCombos - 1
    
    pCargaHospital IIf(cgstrModulo = "PV", 4036, 4037)
    mskFecha.Text = fdtmServerFecha
    mskFechaFin.Text = fdtmServerFecha
    pCargaConceptos
    optPresentacion_Click 0
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
End Sub

Private Sub pCargaConceptos()
    On Error GoTo NotificaError
    Dim rs As New ADODB.Recordset

    Set rs = frsEjecuta_SP("0|1|-1", "sp_PvSelConceptoFactura")
    If rs.RecordCount <> 0 Then
        pLlenarCboRs cboConcepto, rs, 0, 1
    End If
    cboConcepto.AddItem "<TODOS>", 0
    cboConcepto.ItemData(cboConcepto.newIndex) = -1
    cboConcepto.ListIndex = 0

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargaConceptos"))
End Sub

Private Sub pCargaHospital(lngNumOpcion As Long)
On Error GoTo NotificaError
    Dim rs As New ADODB.Recordset
    
    Set rs = frsEjecuta_SP("-1", "Sp_Gnselempresascontable")
    If rs.RecordCount <> 0 Then
        pLlenarCboRs cboHospital, rs, 1, 0
        cboHospital.ListIndex = flngLocalizaCbo(cboHospital, Str(vgintClaveEmpresaContable))
    End If
    
    cboHospital.Enabled = fblnRevisaPermiso(vglngNumeroLogin, lngNumOpcion, "C")
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCargaHospital"))
    Unload Me
End Sub

Private Sub mskFecha_GotFocus()
    pSelMkTexto mskFecha
End Sub

Private Sub MskFecha_LostFocus()
    If Not IsDate(mskFecha.Text) Then
       MsgBox SIHOMsg(29), vbExclamation, "Mensaje"
       Me.mskFecha.SetFocus
    End If
End Sub

Private Sub pImprime(strDestino As String)
On Error GoTo NotificaError
    Dim rsReporte As ADODB.Recordset
    Dim alstrParametros(2) As String
    Dim intTipoReporte As Integer
    
    If fblnValidos() Then
    
        btnEnable False
        
        If strDestino = "P" Then
            Label1.Caption = " Obteniendo información, por favor espere..."
            pgbBarra.Value = 0
            freBarra.Visible = True
            Screen.MousePointer = vbHourglass
        End If
    
        intTipoReporte = IIf(optPresentacion(0).Value, 0, IIf(optPresentacion(1).Value, 1, 2))
        pInstanciaReporte vgrptReporte, IIf(intTipoReporte = 2, "rptResumenCargosDesglosado.rpt", "rptResumenCargos.rpt")
        vgstrParametrosSP = fstrFechaSQL(mskFecha.Text) & "|" & intTipoReporte & "|" & cboHospital.ItemData(cboHospital.ListIndex) & "|" & _
                            fstrFechaSQL(IIf(optPresentacion(0).Value, mskFecha.Text, mskFechaFin.Text)) & "|" & _
                            IIf(optPresentacion(0).Value, "-1", cboConcepto.ItemData(cboConcepto.ListIndex)) & "|" & IIf(ChkIncluirCargos.Value = 1, 1, 0)
        If strDestino = "P" Then
            pgbBarra.Value = 50
        End If
        Set rsReporte = frsEjecuta_SP(vgstrParametrosSP, "Sp_PvRptResumenDiarioCargos")
        
        If strDestino = "P" Then
            pgbBarra.Value = 75
        End If
        
        If strDestino = "P" Then
            Label1.Caption = "  Exportando información, por favor espere..."
            pgbBarra.Value = 100
            freBarra.Visible = False
            Screen.MousePointer = vbDefault
        End If
        
        If rsReporte.RecordCount = 0 Then
            MsgBox SIHOMsg(13), vbInformation, "Mensaje"
        Else
           
            vgrptReporte.DiscardSavedData
            alstrParametros(0) = "hospital; " & cboHospital.List(cboHospital.ListIndex)
            If optPresentacion(0).Value Then
                alstrParametros(1) = "titulo;" & "RESUMEN DE CARGOS Y FORMAS DE PAGO DEL " & UCase(Format(mskFecha.Text, "dd/MMM/yyyy"))
            Else
                alstrParametros(1) = "titulo;" & "RESUMEN DE CARGOS DEL " & UCase(Format(mskFecha.Text, "dd/MMM/yyyy")) & " AL " & UCase(Format(mskFechaFin.Text, "dd/MMM/yyyy"))
            End If
            alstrParametros(2) = "tipoReporte;" & intTipoReporte
            pCargaParameterFields alstrParametros, vgrptReporte
            pImprimeReporte vgrptReporte, rsReporte, strDestino, "Resumen diario de cargos y formas de pago"
        End If
        rsReporte.Close
    End If
    
    btnEnable

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCargaHospital"))
End Sub

Private Sub pExporta()
On Error GoTo NotificaError
    Dim rs As New ADODB.Recordset
    Dim rsReporte As ADODB.Recordset
    Dim rsReporteAnoAnterior As ADODB.Recordset
    Dim intTipoReporte As Integer
    Dim vlstrFechaInicial As String
    Dim vlstrFechaFinal As String
    Dim intRenglonesDias As Long
    Dim intRenglones As Long
    Dim intColumnas As Long
    Dim intRenglonesAnterior As Long
    Dim vlintAumentosBarra As Double
    Dim o_Excel As Object
    Dim o_ExcelAbrir As Object
    Dim o_Libro As Object
    Dim o_Sheet As Object
    Dim vldtmfechaServer As Date
    Dim vlstrMesAnterior As String
    
    Dim vldtmFechainicial As Date
    Dim vldtmFechaFinal As Date
    
    Dim intcontador As Integer
    Dim strTitulos As Variant
    Dim strValor As Variant
    Dim intAutoFit As Integer
    Dim intAnchoColumna As Integer
    Dim intContColumnas As Integer
    
    
    btnEnable False
    
    strTitulos = Array("Fecha del cargo (Fecha y hora completa)", "Año del cargo", "Mes del cargo", "Día del cargo", "Identificador del cargo", "Clave del cargo", "Clave externa", "Descripción", "Cantidad", "Costo unitario", "Precio unitario", "Importe", "Descuento", "Subtotal", "IVA", "Total", "Expediente", "Cuenta", "Nombre paciente", "Tipo de paciente", "Tipo de convenio", "Empresa", "Tipo de cargo", "Familia", "Subfamilia", "Concepto de facturación", "Número de solicitud (Laboratorio o imagenología)", "Número de requisición")
    
    vldtmfechaServer = fdtmServerFecha
    
    'intTipoReporte = 3
    intTipoReporte = IIf(optPresentacion(3).Value = True, 3, IIf(optPresentacion(2).Value = True, 2, 0))
    
    If optPresentacion(3).Value Then
        
        vlstrFechaInicial = Format("01/" & IIf((cboMes.ListIndex + 1) < 10, "0" & cboMes.ListIndex + 1, cboMes.ListIndex + 1) & "/" & cboEjercicio.Text, "dd/MMM/yyyy")
        vlstrFechaFinal = Format(DateAdd("d", -1, DateAdd("m", 1, CDate("01/" & IIf((cboMes.ListIndex + 1) < 10, "0" & cboMes.ListIndex + 1, cboMes.ListIndex + 1) & "/" & cboEjercicio.Text))), "dd/MMM/yyyy")
    
        vldtmFechainicial = "01/" & IIf((cboMes.ListIndex + 1) < 10, "0" & cboMes.ListIndex + 1, cboMes.ListIndex + 1) & "/" & cboEjercicio.Text
        vldtmFechaFinal = DateAdd("d", -1, DateAdd("m", 1, CDate("01/" & IIf((cboMes.ListIndex + 1) < 10, "0" & cboMes.ListIndex + 1, cboMes.ListIndex + 1) & "/" & cboEjercicio.Text)))
    
        If vldtmFechainicial > vldtmfechaServer Then
            'El periodo seleccionado debe ser menor o igual al periodo actual.
            MsgBox SIHOMsg(1291), vbOKOnly + vbExclamation, "Mensaje"
            
            cboMes.SetFocus
            Exit Sub
        End If
        
        Screen.MousePointer = vbHourglass
          
        vgstrParametrosSP = fstrFechaSQL(vlstrFechaInicial) _
                            & "|" _
                            & intTipoReporte _
                            & "|" _
                            & cboHospital.ItemData(cboHospital.ListIndex) _
                            & "|" _
                            & fstrFechaSQL(vlstrFechaFinal) _
                            & "|" _
                            & cboConcepto.ItemData(cboConcepto.ListIndex) _
                            & "|0"
                            
        Set rsReporte = frsEjecuta_SP(vgstrParametrosSP, "Sp_PvRptResumenDiarioCargos")
        
        If rsReporte.RecordCount = 0 Then
            Screen.MousePointer = vbDefault
            MsgBox SIHOMsg(13), vbInformation, "Mensaje"
        Else
            pgbBarra.Value = 0
            freBarra.Visible = True
        
            Set o_Excel = CreateObject("Excel.Application")
            Set o_Libro = o_Excel.Workbooks.Add
            Set o_Sheet = o_Libro.Worksheets(1)
                            
            o_Excel.ActiveWorkbook.ActiveSheet.Name = "Resumen diario"
                
            o_Sheet.range("A:A").HorizontalAlignment = -4131
            
            'Encabezados
            'datos del repote
            o_Excel.cells(1, 1).Value = Trim(vgstrNombreHospitalCH)
            o_Excel.range("A1:E1").Select
            o_Excel.Selection.Merge
            o_Sheet.range("A1:E1").HorizontalAlignment = -4108
            
            o_Excel.cells(2, 1).Value = "RESUMEN DE CARGOS DIARIOS DE " & UCase(cboMes.Text) & " DE " & UCase(cboEjercicio.Text)
            o_Excel.range("A2:E2").Select
            o_Excel.Selection.Merge
            o_Sheet.range("A2:E2").HorizontalAlignment = -4108
            
            o_Excel.cells(3, 1).Value = "CONCEPTO DE FACTURA: " & cboConcepto.Text
            o_Excel.range("A3:E3").Select
            o_Excel.Selection.Merge
            o_Sheet.range("A3:E3").HorizontalAlignment = -4131
        
            vlstrMesAnterior = ""
            If cboMes.Text = "Enero" Then
                vlstrMesAnterior = "Diciembre"
            Else
                If cboMes.Text = "Febrero" Then
                    vlstrMesAnterior = "Enero"
                Else
                    If cboMes.Text = "Marzo" Then
                        vlstrMesAnterior = "Febrero"
                    Else
                        If cboMes.Text = "Abril" Then
                            vlstrMesAnterior = "Marzo"
                        Else
                            If cboMes.Text = "Mayo" Then
                                vlstrMesAnterior = "Abril"
                            Else
                                If cboMes.Text = "Junio" Then
                                    vlstrMesAnterior = "Mayo"
                                Else
                                    If cboMes.Text = "Julio" Then
                                        vlstrMesAnterior = "Junio"
                                    Else
                                        If cboMes.Text = "Agosto" Then
                                            vlstrMesAnterior = "Julio"
                                        Else
                                            If cboMes.Text = "Septiembre" Then
                                                vlstrMesAnterior = "Agosto"
                                            Else
                                                If cboMes.Text = "Octubre" Then
                                                    vlstrMesAnterior = "Septiembre"
                                                Else
                                                    If cboMes.Text = "Noviembre" Then
                                                        vlstrMesAnterior = "Octubre"
                                                    Else
                                                        If cboMes.Text = "Diciembre" Then
                                                            vlstrMesAnterior = "Noviembre"
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            
            pgbBarra.Min = 0
            pgbBarra.Max = 100
        
            o_Excel.cells(5, 1).Value = "Fecha"
            o_Excel.cells(5, 2).Value = "Ingresos " & LCase(cboMes.Text) & " " & Val(cboEjercicio.Text) - 1
            o_Excel.cells(5, 3).Value = "Ingresos " & LCase(vlstrMesAnterior) & " " & IIf(LCase(vlstrMesAnterior) = "diciembre", Val(cboEjercicio.Text) - 1, Val(cboEjercicio.Text))
            o_Excel.cells(5, 4).Value = "Ingresos " & LCase(cboMes.Text) & " " & Val(cboEjercicio.Text)
            o_Excel.cells(5, 5).Value = "Descuentos " & LCase(cboMes.Text) & " " & Val(cboEjercicio.Text)
            
            o_Sheet.range("A5:E5").HorizontalAlignment = -4108
                        
            pgbBarra.Value = 5
            
            pgbBarra.Value = 15
                    
            'Agrega información de las cuentas y pone en 0s los rangos
            intRenglonesDias = 6
            vlintAumentosBarra = 30 / IIf(rsReporte.RecordCount = 0, 30, rsReporte.RecordCount)
                    
            o_Excel.range("A:A").NumberFormat = "dd/mmm/yyyy"
            o_Sheet.range("B:B").NumberFormat = "$ ###,###,###,##0.00"
            o_Sheet.range("C:C").NumberFormat = "$ ###,###,###,##0.00"
            o_Sheet.range("D:D").NumberFormat = "$ ###,###,###,##0.00"
            o_Sheet.range("E:E").NumberFormat = "$ ###,###,###,##0.00"
                    
            Do While vldtmFechainicial <= vldtmFechaFinal And vldtmFechainicial <= vldtmfechaServer
                o_Excel.cells(intRenglonesDias, 1).Value = vldtmFechainicial
                o_Excel.cells(intRenglonesDias, 2).Value = "0"
                o_Excel.cells(intRenglonesDias, 3).Value = "0"
                o_Excel.cells(intRenglonesDias, 4).Value = "0"
                o_Excel.cells(intRenglonesDias, 5).Value = "0"
                
                vldtmFechainicial = DateAdd("d", 1, vldtmFechainicial)
                intRenglonesDias = intRenglonesDias + 1
            Loop
            
            intRenglones = 6
            
            If rsReporte.RecordCount <> 0 Then
                rsReporte.MoveFirst
            End If
                    
            Do While Not rsReporte.EOF
                If IsDate(o_Excel.cells(intRenglones, 1).Value) Then
                    Do While CDate(o_Excel.cells(intRenglones, 1).Value) < rsReporte!fecha
                        intRenglones = intRenglones + 1
                    Loop
                
                    o_Excel.cells(intRenglones, 1).Value = rsReporte!fecha
                    o_Excel.cells(intRenglones, 4).Value = rsReporte!Importe
                    o_Excel.cells(intRenglones, 5).Value = rsReporte!Descuento
                        
                    intRenglones = intRenglones + 1
                    
                    rsReporte.MoveNext
                    pgbBarra.Value = pgbBarra.Value + vlintAumentosBarra
                Else
                    Exit Do
                End If
            Loop
            
            intRenglones = intRenglonesDias 'Se asigna el número de renglones de los dias que al final de cuentas es el que mas importa mas que los renglones que si traían importes
            
            pgbBarra.Value = 45
        
            '///// Mes solicitado pero del año anterior /////
                vlstrFechaInicial = Format(DateAdd("yyyy", -1, CDate("01/" & IIf((cboMes.ListIndex + 1) < 10, "0" & cboMes.ListIndex + 1, cboMes.ListIndex + 1) & "/" & cboEjercicio.Text)), "dd/MMM/yyyy")
                vlstrFechaFinal = Format(DateAdd("yyyy", -1, CDate(DateAdd("d", -1, DateAdd("m", 1, CDate("01/" & IIf((cboMes.ListIndex + 1) < 10, "0" & cboMes.ListIndex + 1, cboMes.ListIndex + 1) & "/" & cboEjercicio.Text))))), "dd/MMM/yyyy")
                
                vldtmFechainicial = DateAdd("yyyy", -1, CDate("01/" & IIf((cboMes.ListIndex + 1) < 10, "0" & cboMes.ListIndex + 1, cboMes.ListIndex + 1) & "/" & cboEjercicio.Text))
                vldtmFechaFinal = DateAdd("yyyy", -1, CDate(DateAdd("d", -1, DateAdd("m", 1, CDate("01/" & IIf((cboMes.ListIndex + 1) < 10, "0" & cboMes.ListIndex + 1, cboMes.ListIndex + 1) & "/" & cboEjercicio.Text)))))
                
                vgstrParametrosSP = fstrFechaSQL(vlstrFechaInicial) _
                                    & "|" _
                                    & intTipoReporte _
                                    & "|" _
                                    & cboHospital.ItemData(cboHospital.ListIndex) _
                                    & "|" _
                                    & fstrFechaSQL(vlstrFechaFinal) _
                                    & "|" _
                                    & cboConcepto.ItemData(cboConcepto.ListIndex) _
                                    & "|0"
                Set rsReporteAnoAnterior = frsEjecuta_SP(vgstrParametrosSP, "Sp_PvRptResumenDiarioCargos")
            
                vlintAumentosBarra = 20 / IIf(rsReporteAnoAnterior.RecordCount = 0, 20, rsReporteAnoAnterior.RecordCount)
                
                intRenglonesAnterior = 6
                
                Do While Not rsReporteAnoAnterior.EOF
                    If IsDate(o_Excel.cells(intRenglonesAnterior, 1).Value) Then
                        
                        Do While IsDate(o_Excel.cells(intRenglonesAnterior, 1).Value)
                            If CDate(o_Excel.cells(intRenglonesAnterior, 1).Value) = DateAdd("yyyy", 1, rsReporteAnoAnterior!fecha) Then
                                o_Excel.cells(intRenglonesAnterior, 2).Value = rsReporteAnoAnterior!Importe
                                intRenglonesAnterior = intRenglonesAnterior + 1
                                Exit Do
                            Else
                                intRenglonesAnterior = intRenglonesAnterior + 1
                            End If
                        Loop
                        
                        rsReporteAnoAnterior.MoveNext
                        pgbBarra.Value = pgbBarra.Value + vlintAumentosBarra
                    Else
                        Exit Do
                    End If
                Loop
            '///// Mes solicitado pero del año anterior /////
            
            pgbBarra.Value = 65
            
            '///// Mes anterior al solicitado /////
                vlstrFechaInicial = Format(DateAdd("m", -1, CDate("01/" & IIf((cboMes.ListIndex + 1) < 10, "0" & cboMes.ListIndex + 1, cboMes.ListIndex + 1) & "/" & cboEjercicio.Text)), "dd/MMM/yyyy")
                vlstrFechaFinal = Format(DateAdd("m", -1, CDate(DateAdd("d", -1, DateAdd("m", 1, CDate("01/" & IIf((cboMes.ListIndex + 1) < 10, "0" & cboMes.ListIndex + 1, cboMes.ListIndex + 1) & "/" & cboEjercicio.Text))))), "dd/MMM/yyyy")
                
                vldtmFechainicial = DateAdd("m", -1, CDate("01/" & IIf((cboMes.ListIndex + 1) < 10, "0" & cboMes.ListIndex + 1, cboMes.ListIndex + 1) & "/" & cboEjercicio.Text))
                vldtmFechaFinal = DateAdd("m", -1, CDate(DateAdd("d", -1, DateAdd("m", 1, CDate("01/" & IIf((cboMes.ListIndex + 1) < 10, "0" & cboMes.ListIndex + 1, cboMes.ListIndex + 1) & "/" & cboEjercicio.Text)))))
                
                vgstrParametrosSP = fstrFechaSQL(vlstrFechaInicial) _
                                    & "|" _
                                    & intTipoReporte _
                                    & "|" _
                                    & cboHospital.ItemData(cboHospital.ListIndex) _
                                    & "|" _
                                    & fstrFechaSQL(vlstrFechaFinal) _
                                    & "|" _
                                    & cboConcepto.ItemData(cboConcepto.ListIndex) _
                                    & "|0"
                Set rsReporteAnoAnterior = frsEjecuta_SP(vgstrParametrosSP, "Sp_PvRptResumenDiarioCargos")
            
                vlintAumentosBarra = 25 / IIf(rsReporteAnoAnterior.RecordCount = 0, 25, rsReporteAnoAnterior.RecordCount)
                
                intRenglonesAnterior = 6
                
                Do While Not rsReporteAnoAnterior.EOF
                    If IsDate(o_Excel.cells(intRenglonesAnterior, 1).Value) Then
                        
                        Do While IsDate(o_Excel.cells(intRenglonesAnterior, 1).Value)
                            If CDate(o_Excel.cells(intRenglonesAnterior, 1).Value) = DateAdd("m", 1, rsReporteAnoAnterior!fecha) Then
                                o_Excel.cells(intRenglonesAnterior, 3).Value = rsReporteAnoAnterior!Importe
                                intRenglonesAnterior = intRenglonesAnterior + 1
                                Exit Do
                            Else
                                intRenglonesAnterior = intRenglonesAnterior + 1
                            End If
                        Loop
                        
                        rsReporteAnoAnterior.MoveNext
                        pgbBarra.Value = pgbBarra.Value + vlintAumentosBarra
                    Else
                        Exit Do
                    End If
                Loop
            '///// Mes anterior al solicitado /////
            
            pgbBarra.Value = 90
            
            intRenglones = intRenglones + 1
            
            o_Excel.cells(intRenglones, 1).Value = "Acumulado al día " & Format(o_Excel.cells(intRenglones - 2, 1).Value, "dd/mmm/yyyy")
            o_Excel.cells(intRenglones + 1, 1).Value = "Promedio diario"
            o_Excel.cells(intRenglones + 2, 1).Value = "Proyectado"
            o_Excel.cells(intRenglones + 3, 1).Value = "Presupuesto"
            o_Excel.cells(intRenglones + 4, 1).Value = "Porcentaje presupuesto"
            
            If cboMes.Text = "Enero" Then
                Set rs = frsRegresaRs("SELECT NVL(MNYENERO,0) Valor FROM CNIMPORTEMENSUALPRESUPUESTADO WHERE TNYCLAVEEMPRESA = " & CStr(vgintClaveEmpresaContable) & " and SMIEJERCICIO = " & cboEjercicio.Text)
                If Not rs.EOF Then o_Excel.cells(intRenglones + 3, 4).Value = rs!valor
                
                Set rs = frsRegresaRs("SELECT NVL(MNYDICIEMBRE,0) Valor FROM CNIMPORTEMENSUALPRESUPUESTADO WHERE TNYCLAVEEMPRESA = " & CStr(vgintClaveEmpresaContable) & " and SMIEJERCICIO = " & Val(cboEjercicio.Text) - 1)
                If Not rs.EOF Then o_Excel.cells(intRenglones + 3, 3).Value = rs!valor
                
                Set rs = frsRegresaRs("SELECT NVL(MNYENERO,0) Valor FROM CNIMPORTEMENSUALPRESUPUESTADO WHERE TNYCLAVEEMPRESA = " & CStr(vgintClaveEmpresaContable) & " and SMIEJERCICIO = " & Val(cboEjercicio.Text) - 1)
                If Not rs.EOF Then o_Excel.cells(intRenglones + 3, 2).Value = rs!valor
            Else
                If cboMes.Text = "Febrero" Then
                    Set rs = frsRegresaRs("SELECT NVL(MNYFebrero,0) Valor FROM CNIMPORTEMENSUALPRESUPUESTADO WHERE TNYCLAVEEMPRESA = " & CStr(vgintClaveEmpresaContable) & " and SMIEJERCICIO = " & cboEjercicio.Text)
                    If Not rs.EOF Then o_Excel.cells(intRenglones + 3, 4).Value = rs!valor
                    
                    Set rs = frsRegresaRs("SELECT NVL(MNYENERO,0) Valor FROM CNIMPORTEMENSUALPRESUPUESTADO WHERE TNYCLAVEEMPRESA = " & CStr(vgintClaveEmpresaContable) & " and SMIEJERCICIO = " & cboEjercicio.Text)
                    If Not rs.EOF Then o_Excel.cells(intRenglones + 3, 3).Value = rs!valor
                    
                    Set rs = frsRegresaRs("SELECT NVL(MNYFebrero,0) Valor FROM CNIMPORTEMENSUALPRESUPUESTADO WHERE TNYCLAVEEMPRESA = " & CStr(vgintClaveEmpresaContable) & " and SMIEJERCICIO = " & Val(cboEjercicio.Text) - 1)
                    If Not rs.EOF Then o_Excel.cells(intRenglones + 3, 2).Value = rs!valor
                Else
                    If cboMes.Text = "Marzo" Then
                        Set rs = frsRegresaRs("SELECT NVL(MNYMarzo,0) Valor FROM CNIMPORTEMENSUALPRESUPUESTADO WHERE TNYCLAVEEMPRESA = " & CStr(vgintClaveEmpresaContable) & " and SMIEJERCICIO = " & cboEjercicio.Text)
                        If Not rs.EOF Then o_Excel.cells(intRenglones + 3, 4).Value = rs!valor
                        
                        Set rs = frsRegresaRs("SELECT NVL(MNYFebrero,0) Valor FROM CNIMPORTEMENSUALPRESUPUESTADO WHERE TNYCLAVEEMPRESA = " & CStr(vgintClaveEmpresaContable) & " and SMIEJERCICIO = " & cboEjercicio.Text)
                        If Not rs.EOF Then o_Excel.cells(intRenglones + 3, 3).Value = rs!valor
                        
                        Set rs = frsRegresaRs("SELECT NVL(MNYMarzo,0) Valor FROM CNIMPORTEMENSUALPRESUPUESTADO WHERE TNYCLAVEEMPRESA = " & CStr(vgintClaveEmpresaContable) & " and SMIEJERCICIO = " & Val(cboEjercicio.Text) - 1)
                        If Not rs.EOF Then o_Excel.cells(intRenglones + 3, 2).Value = rs!valor
                    Else
                        If cboMes.Text = "Abril" Then
                            Set rs = frsRegresaRs("SELECT NVL(MNYAbril,0) Valor FROM CNIMPORTEMENSUALPRESUPUESTADO WHERE TNYCLAVEEMPRESA = " & CStr(vgintClaveEmpresaContable) & " and SMIEJERCICIO = " & cboEjercicio.Text)
                            If Not rs.EOF Then o_Excel.cells(intRenglones + 3, 4).Value = rs!valor
                            
                            Set rs = frsRegresaRs("SELECT NVL(MNYMarzo,0) Valor FROM CNIMPORTEMENSUALPRESUPUESTADO WHERE TNYCLAVEEMPRESA = " & CStr(vgintClaveEmpresaContable) & " and SMIEJERCICIO = " & cboEjercicio.Text)
                            If Not rs.EOF Then o_Excel.cells(intRenglones + 3, 3).Value = rs!valor
                            
                            Set rs = frsRegresaRs("SELECT NVL(MNYAbril,0) Valor FROM CNIMPORTEMENSUALPRESUPUESTADO WHERE TNYCLAVEEMPRESA = " & CStr(vgintClaveEmpresaContable) & " and SMIEJERCICIO = " & Val(cboEjercicio.Text) - 1)
                            If Not rs.EOF Then o_Excel.cells(intRenglones + 3, 2).Value = rs!valor
                        Else
                            If cboMes.Text = "Mayo" Then
                                Set rs = frsRegresaRs("SELECT NVL(MNYMayo,0) Valor FROM CNIMPORTEMENSUALPRESUPUESTADO WHERE TNYCLAVEEMPRESA = " & CStr(vgintClaveEmpresaContable) & " and SMIEJERCICIO = " & cboEjercicio.Text)
                                If Not rs.EOF Then o_Excel.cells(intRenglones + 3, 4).Value = rs!valor
                                
                                Set rs = frsRegresaRs("SELECT NVL(MNYAbril,0) Valor FROM CNIMPORTEMENSUALPRESUPUESTADO WHERE TNYCLAVEEMPRESA = " & CStr(vgintClaveEmpresaContable) & " and SMIEJERCICIO = " & cboEjercicio.Text)
                                If Not rs.EOF Then o_Excel.cells(intRenglones + 3, 3).Value = rs!valor
                                
                                Set rs = frsRegresaRs("SELECT NVL(MNYMayo,0) Valor FROM CNIMPORTEMENSUALPRESUPUESTADO WHERE TNYCLAVEEMPRESA = " & CStr(vgintClaveEmpresaContable) & " and SMIEJERCICIO = " & Val(cboEjercicio.Text) - 1)
                                If Not rs.EOF Then o_Excel.cells(intRenglones + 3, 2).Value = rs!valor
                            Else
                                If cboMes.Text = "Junio" Then
                                    Set rs = frsRegresaRs("SELECT NVL(MNYJunio,0) Valor FROM CNIMPORTEMENSUALPRESUPUESTADO WHERE TNYCLAVEEMPRESA = " & CStr(vgintClaveEmpresaContable) & " and SMIEJERCICIO = " & cboEjercicio.Text)
                                    If Not rs.EOF Then o_Excel.cells(intRenglones + 3, 4).Value = rs!valor
                                    
                                    Set rs = frsRegresaRs("SELECT NVL(MNYMayo,0) Valor FROM CNIMPORTEMENSUALPRESUPUESTADO WHERE TNYCLAVEEMPRESA = " & CStr(vgintClaveEmpresaContable) & " and SMIEJERCICIO = " & cboEjercicio.Text)
                                    If Not rs.EOF Then o_Excel.cells(intRenglones + 3, 3).Value = rs!valor
                                    
                                    Set rs = frsRegresaRs("SELECT NVL(MNYJunio,0) Valor FROM CNIMPORTEMENSUALPRESUPUESTADO WHERE TNYCLAVEEMPRESA = " & CStr(vgintClaveEmpresaContable) & " and SMIEJERCICIO = " & Val(cboEjercicio.Text) - 1)
                                    If Not rs.EOF Then o_Excel.cells(intRenglones + 3, 2).Value = rs!valor
                                Else
                                    If cboMes.Text = "Julio" Then
                                        Set rs = frsRegresaRs("SELECT NVL(MNYJulio,0) Valor FROM CNIMPORTEMENSUALPRESUPUESTADO WHERE TNYCLAVEEMPRESA = " & CStr(vgintClaveEmpresaContable) & " and SMIEJERCICIO = " & cboEjercicio.Text)
                                        If Not rs.EOF Then o_Excel.cells(intRenglones + 3, 4).Value = rs!valor
                                        
                                        Set rs = frsRegresaRs("SELECT NVL(MNYJunio,0) Valor FROM CNIMPORTEMENSUALPRESUPUESTADO WHERE TNYCLAVEEMPRESA = " & CStr(vgintClaveEmpresaContable) & " and SMIEJERCICIO = " & cboEjercicio.Text)
                                        If Not rs.EOF Then o_Excel.cells(intRenglones + 3, 3).Value = rs!valor
                                            
                                        Set rs = frsRegresaRs("SELECT NVL(MNYJulio,0) Valor FROM CNIMPORTEMENSUALPRESUPUESTADO WHERE TNYCLAVEEMPRESA = " & CStr(vgintClaveEmpresaContable) & " and SMIEJERCICIO = " & Val(cboEjercicio.Text) - 1)
                                        If Not rs.EOF Then o_Excel.cells(intRenglones + 3, 2).Value = rs!valor
                                    Else
                                        If cboMes.Text = "Agosto" Then
                                            Set rs = frsRegresaRs("SELECT NVL(MNYAgosto,0) Valor FROM CNIMPORTEMENSUALPRESUPUESTADO WHERE TNYCLAVEEMPRESA = " & CStr(vgintClaveEmpresaContable) & " and SMIEJERCICIO = " & cboEjercicio.Text)
                                            If Not rs.EOF Then o_Excel.cells(intRenglones + 3, 4).Value = rs!valor
                                            
                                            Set rs = frsRegresaRs("SELECT NVL(MNYJulio,0) Valor FROM CNIMPORTEMENSUALPRESUPUESTADO WHERE TNYCLAVEEMPRESA = " & CStr(vgintClaveEmpresaContable) & " and SMIEJERCICIO = " & cboEjercicio.Text)
                                            If Not rs.EOF Then o_Excel.cells(intRenglones + 3, 3).Value = rs!valor
                                            
                                            Set rs = frsRegresaRs("SELECT NVL(MNYAgosto,0) Valor FROM CNIMPORTEMENSUALPRESUPUESTADO WHERE TNYCLAVEEMPRESA = " & CStr(vgintClaveEmpresaContable) & " and SMIEJERCICIO = " & Val(cboEjercicio.Text) - 1)
                                            If Not rs.EOF Then o_Excel.cells(intRenglones + 3, 2).Value = rs!valor
                                        Else
                                            If cboMes.Text = "Septiembre" Then
                                                Set rs = frsRegresaRs("SELECT NVL(MNYSeptiembre,0) Valor FROM CNIMPORTEMENSUALPRESUPUESTADO WHERE TNYCLAVEEMPRESA = " & CStr(vgintClaveEmpresaContable) & " and SMIEJERCICIO = " & cboEjercicio.Text)
                                                If Not rs.EOF Then o_Excel.cells(intRenglones + 3, 4).Value = rs!valor
                                                
                                                Set rs = frsRegresaRs("SELECT NVL(MNYAgosto,0) Valor FROM CNIMPORTEMENSUALPRESUPUESTADO WHERE TNYCLAVEEMPRESA = " & CStr(vgintClaveEmpresaContable) & " and SMIEJERCICIO = " & cboEjercicio.Text)
                                                If Not rs.EOF Then o_Excel.cells(intRenglones + 3, 3).Value = rs!valor
                                                
                                                Set rs = frsRegresaRs("SELECT NVL(MNYSeptiembre,0) Valor FROM CNIMPORTEMENSUALPRESUPUESTADO WHERE TNYCLAVEEMPRESA = " & CStr(vgintClaveEmpresaContable) & " and SMIEJERCICIO = " & Val(cboEjercicio.Text) - 1)
                                                If Not rs.EOF Then o_Excel.cells(intRenglones + 3, 2).Value = rs!valor
                                            Else
                                                If cboMes.Text = "Octubre" Then
                                                    Set rs = frsRegresaRs("SELECT NVL(MNYOctubre,0) Valor FROM CNIMPORTEMENSUALPRESUPUESTADO WHERE TNYCLAVEEMPRESA = " & CStr(vgintClaveEmpresaContable) & " and SMIEJERCICIO = " & cboEjercicio.Text)
                                                    If Not rs.EOF Then o_Excel.cells(intRenglones + 3, 4).Value = rs!valor
                                                    
                                                    Set rs = frsRegresaRs("SELECT NVL(MNYSeptiembre,0) Valor FROM CNIMPORTEMENSUALPRESUPUESTADO WHERE TNYCLAVEEMPRESA = " & CStr(vgintClaveEmpresaContable) & " and SMIEJERCICIO = " & cboEjercicio.Text)
                                                    If Not rs.EOF Then o_Excel.cells(intRenglones + 3, 3).Value = rs!valor
                                                    
                                                    Set rs = frsRegresaRs("SELECT NVL(MNYOctubre,0) Valor FROM CNIMPORTEMENSUALPRESUPUESTADO WHERE TNYCLAVEEMPRESA = " & CStr(vgintClaveEmpresaContable) & " and SMIEJERCICIO = " & Val(cboEjercicio.Text) - 1)
                                                    If Not rs.EOF Then o_Excel.cells(intRenglones + 3, 2).Value = rs!valor
                                                Else
                                                    If cboMes.Text = "Noviembre" Then
                                                        Set rs = frsRegresaRs("SELECT NVL(MNYNoviembre,0) Valor FROM CNIMPORTEMENSUALPRESUPUESTADO WHERE TNYCLAVEEMPRESA = " & CStr(vgintClaveEmpresaContable) & " and SMIEJERCICIO = " & cboEjercicio.Text)
                                                        If Not rs.EOF Then o_Excel.cells(intRenglones + 3, 4).Value = rs!valor
                                                        
                                                        Set rs = frsRegresaRs("SELECT NVL(MNYOctubre,0) Valor FROM CNIMPORTEMENSUALPRESUPUESTADO WHERE TNYCLAVEEMPRESA = " & CStr(vgintClaveEmpresaContable) & " and SMIEJERCICIO = " & cboEjercicio.Text)
                                                        If Not rs.EOF Then o_Excel.cells(intRenglones + 3, 3).Value = rs!valor
                                                        
                                                        Set rs = frsRegresaRs("SELECT NVL(MNYNoviembre,0) Valor FROM CNIMPORTEMENSUALPRESUPUESTADO WHERE TNYCLAVEEMPRESA = " & CStr(vgintClaveEmpresaContable) & " and SMIEJERCICIO = " & Val(cboEjercicio.Text) - 1)
                                                        If Not rs.EOF Then o_Excel.cells(intRenglones + 3, 2).Value = rs!valor
                                                    Else
                                                        If cboMes.Text = "Diciembre" Then
                                                            Set rs = frsRegresaRs("SELECT NVL(MNYDiciembre,0) Valor FROM CNIMPORTEMENSUALPRESUPUESTADO WHERE TNYCLAVEEMPRESA = " & CStr(vgintClaveEmpresaContable) & " and SMIEJERCICIO = " & cboEjercicio.Text)
                                                            If Not rs.EOF Then o_Excel.cells(intRenglones + 3, 4).Value = rs!valor
                                                            
                                                            Set rs = frsRegresaRs("SELECT NVL(MNYNoviembre,0) Valor FROM CNIMPORTEMENSUALPRESUPUESTADO WHERE TNYCLAVEEMPRESA = " & CStr(vgintClaveEmpresaContable) & " and SMIEJERCICIO = " & cboEjercicio.Text)
                                                            If Not rs.EOF Then o_Excel.cells(intRenglones + 3, 3).Value = rs!valor
                                                            
                                                            Set rs = frsRegresaRs("SELECT NVL(MNYDiciembre,0) Valor FROM CNIMPORTEMENSUALPRESUPUESTADO WHERE TNYCLAVEEMPRESA = " & CStr(vgintClaveEmpresaContable) & " and SMIEJERCICIO = " & Val(cboEjercicio.Text) - 1)
                                                            If Not rs.EOF Then o_Excel.cells(intRenglones + 3, 2).Value = rs!valor
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            
    '        o_Excel.Cells(intRenglones + 3, 4).Value
            
            o_Excel.range("B" & intRenglones).Select
            o_Excel.ActiveCell.FormulaR1C1 = "=SUM(R[-" & intRenglones - 6 & "]C:R[-1]C)"
            
            o_Excel.range("C" & intRenglones).Select
            o_Excel.ActiveCell.FormulaR1C1 = "=SUM(R[-" & intRenglones - 6 & "]C:R[-1]C)"
            
            o_Excel.range("D" & intRenglones).Select
            o_Excel.ActiveCell.FormulaR1C1 = "=SUM(R[-" & intRenglones - 6 & "]C:R[-1]C)"
            
            o_Excel.range("E" & intRenglones).Select
            o_Excel.ActiveCell.FormulaR1C1 = "=SUM(R[-" & intRenglones - 6 & "]C:R[-1]C)"
            
            o_Excel.cells(intRenglones + 1, 2).Value = o_Excel.cells(intRenglones, 2).Value / (intRenglones - 7)
            o_Excel.cells(intRenglones + 1, 3).Value = o_Excel.cells(intRenglones, 3).Value / (intRenglones - 7)
            o_Excel.cells(intRenglones + 1, 4).Value = o_Excel.cells(intRenglones, 4).Value / (intRenglones - 7)
            o_Excel.cells(intRenglones + 1, 5).Value = o_Excel.cells(intRenglones, 5).Value / (intRenglones - 7)
            
            o_Excel.cells(intRenglones + 2, 2).Value = o_Excel.cells(intRenglones + 1, 2).Value * Day(CDate(DateAdd("d", -1, DateAdd("m", 1, CDate("01/" & IIf((cboMes.ListIndex + 1) < 10, "0" & cboMes.ListIndex + 1, cboMes.ListIndex + 1) & "/" & cboEjercicio.Text)))))
            o_Excel.cells(intRenglones + 2, 3).Value = o_Excel.cells(intRenglones + 1, 3).Value * Day(CDate(DateAdd("d", -1, DateAdd("m", 1, CDate("01/" & IIf((cboMes.ListIndex + 1) < 10, "0" & cboMes.ListIndex + 1, cboMes.ListIndex + 1) & "/" & cboEjercicio.Text)))))
            o_Excel.cells(intRenglones + 2, 4).Value = o_Excel.cells(intRenglones + 1, 4).Value * Day(CDate(DateAdd("d", -1, DateAdd("m", 1, CDate("01/" & IIf((cboMes.ListIndex + 1) < 10, "0" & cboMes.ListIndex + 1, cboMes.ListIndex + 1) & "/" & cboEjercicio.Text)))))
            o_Excel.cells(intRenglones + 2, 5).Value = o_Excel.cells(intRenglones + 1, 5).Value * Day(CDate(DateAdd("d", -1, DateAdd("m", 1, CDate("01/" & IIf((cboMes.ListIndex + 1) < 10, "0" & cboMes.ListIndex + 1, cboMes.ListIndex + 1) & "/" & cboEjercicio.Text)))))
            
            If Not o_Excel.cells(intRenglones + 3, 4).Value = "" Then
                o_Excel.cells(intRenglones + 4, 4).Value = ((o_Excel.cells(intRenglones + 2, 4).Value - o_Excel.cells(intRenglones + 3, 4).Value) / o_Excel.cells(intRenglones + 3, 4).Value) + 1
                o_Excel.cells(intRenglones + 4, 4).Style = "Percent"
            End If
            
            If Not o_Excel.cells(intRenglones + 3, 3).Value = "" Then
                o_Excel.cells(intRenglones + 4, 3).Value = ((o_Excel.cells(intRenglones + 2, 3).Value - o_Excel.cells(intRenglones + 3, 3).Value) / o_Excel.cells(intRenglones + 3, 3).Value) + 1
                o_Excel.cells(intRenglones + 4, 3).Style = "Percent"
            End If
            
            If Not o_Excel.cells(intRenglones + 3, 2).Value = "" Then
                o_Excel.cells(intRenglones + 4, 2).Value = ((o_Excel.cells(intRenglones + 2, 2).Value - o_Excel.cells(intRenglones + 3, 2).Value) / o_Excel.cells(intRenglones + 3, 2).Value) + 1
                o_Excel.cells(intRenglones + 4, 2).Style = "Percent"
            End If
            
            pgbBarra.Value = 95
                    
            o_Excel.range("A1:E2").Select
            o_Excel.Selection.Font.Bold = True
                    
            o_Excel.Columns("A:E").Select
            o_Excel.range("A4").Activate
            o_Excel.Columns("A:E").entirecolumn.AutoFit
                    
            o_Excel.Selection.CurrentRegion.Columns.AutoFit
            o_Excel.Selection.CurrentRegion.Rows.AutoFit
            
            o_Excel.range("A1:E" & intRenglones + 4).Select
            o_Excel.range("E" & intRenglones + 4).Activate
            
            o_Excel.Selection.Borders(7).LineStyle = 1
            o_Excel.Selection.Borders(8).LineStyle = 1
            o_Excel.Selection.Borders(9).LineStyle = 1
            
            o_Excel.Selection.Borders(10).LineStyle = 1
            o_Excel.Selection.Borders(11).LineStyle = 1
            o_Excel.Selection.Borders(12).LineStyle = 1
            
            o_Excel.range("A" & intRenglones & ":E" & intRenglones + 4).Select
            o_Excel.range("E" & intRenglones + 4).Activate
            o_Excel.Selection.Font.Bold = True
                     
            o_Excel.cells(1).Select
            
            pgbBarra.Value = 100
            freBarra.Visible = False
            
            'La información ha sido exportada exitosamente
            MsgBox SIHOMsg(1185), vbOKOnly + vbInformation, "Mensaje"
            
            Screen.MousePointer = vbDefault
                    
            o_Excel.Visible = True
        
            Set o_Excel = Nothing
        
            rsReporte.Close
        End If
        
    ElseIf optPresentacion(2).Value Then
    
        intAnchoColumna = 120
    
        vlstrFechaInicial = Format(mskFecha.Text, "dd/MMM/yyyy")
        vlstrFechaFinal = Format(mskFechaFin.Text, "dd/MMM/yyyy")
        
        If vldtmFechainicial > vldtmfechaServer Then
            'El periodo seleccionado debe ser menor o igual al periodo actual.
            MsgBox SIHOMsg(1291), vbOKOnly + vbExclamation, "Mensaje"
        
            mskFecha.SetFocus
            Exit Sub
        End If
        
        vgstrParametrosSP = fstrFechaSQL(vlstrFechaInicial) _
                            & "|" _
                            & intTipoReporte _
                            & "|" _
                            & cboHospital.ItemData(cboHospital.ListIndex) _
                            & "|" _
                            & fstrFechaSQL(vlstrFechaFinal) _
                            & "|" _
                            & cboConcepto.ItemData(cboConcepto.ListIndex) _
                            & "|" _
                            & IIf(ChkIncluirCargos.Value = 1, 1, 0)
                            
        Set rsReporte = frsEjecuta_SP(vgstrParametrosSP, "Sp_PvRptResumenDiarioCargos")
        
        If rsReporte.RecordCount = 0 Then
            Screen.MousePointer = vbDefault
            MsgBox SIHOMsg(13), vbInformation, "Mensaje"
        Else
        
            pgbBarra.Max = rsReporte.RecordCount
            
            pgbBarra.Value = 0
            freBarra.Visible = True
            
            Set o_Excel = CreateObject("Excel.Application")
            Set o_Libro = o_Excel.Workbooks.Add
            Set o_Sheet = o_Libro.Worksheets(1)
            
            'Colocar los valores exportar
            o_Sheet.cells(4, 1).CopyFromRecordset rsReporte, rsReporte.RecordCount
            
            'Se insertar 26 columnas
            For intAutoFit = 1 To 28
                o_Sheet.range("A:A").Insert
            Next intAutoFit
            
            o_Sheet.Columns(47).Copy Destination:=o_Sheet.range("A:A")   ' Fecha del cargo (Fecha y hora completa)
            o_Sheet.Columns(47).Copy Destination:=o_Sheet.range("B:B")   ' Año del cargo
            o_Sheet.Columns(47).Copy Destination:=o_Sheet.range("C:C")   ' Mes del cargo
            o_Sheet.Columns(47).Copy Destination:=o_Sheet.range("D:D")   ' Día del cargo
            o_Sheet.Columns(53).Copy Destination:=o_Sheet.range("E:E")   ' Identificador del cargo
            o_Sheet.Columns(54).Copy Destination:=o_Sheet.range("F:F")   ' Clave del artículo
            o_Sheet.Columns(55).Copy Destination:=o_Sheet.range("G:G")   ' Clave externa
            o_Sheet.Columns(31).Copy Destination:=o_Sheet.range("H:H")   ' Descripción
            o_Sheet.Columns(32).Copy Destination:=o_Sheet.range("I:I")   ' Cantidad
            o_Sheet.Columns(41).Copy Destination:=o_Sheet.range("J:J")   ' Costo unitario
            o_Sheet.Columns(34).Copy Destination:=o_Sheet.range("K:K")   ' Precio unitario
            o_Sheet.Columns(35).Copy Destination:=o_Sheet.range("L:L")   ' Importe
            o_Sheet.Columns(36).Copy Destination:=o_Sheet.range("M:M")   ' Descuento
            o_Sheet.Columns(37).Copy Destination:=o_Sheet.range("N:N")   ' Subtotal
            o_Sheet.Columns(38).Copy Destination:=o_Sheet.range("O:O")   ' IVA
            o_Sheet.Columns(39).Copy Destination:=o_Sheet.range("P:P")   ' Total
            o_Sheet.Columns(56).Copy Destination:=o_Sheet.range("Q:Q")   ' Expediente
            o_Sheet.Columns(42).Copy Destination:=o_Sheet.range("R:R")   ' Cuenta
            o_Sheet.Columns(43).Copy Destination:=o_Sheet.range("S:S")   ' Nombre paciente
            o_Sheet.Columns(61).Copy Destination:=o_Sheet.range("T:T")   ' Tipo de paciente
            o_Sheet.Columns(62).Copy Destination:=o_Sheet.range("U:U")   ' Tipo de convenio
            o_Sheet.Columns(63).Copy Destination:=o_Sheet.range("V:V")   ' Empresa
            o_Sheet.Columns(52).Copy Destination:=o_Sheet.range("W:W")   ' Tipo de cargo
            o_Sheet.Columns(57).Copy Destination:=o_Sheet.range("X:X")   ' Familia
            o_Sheet.Columns(58).Copy Destination:=o_Sheet.range("Y:Y")   ' Subfamilia
            o_Sheet.Columns(30).Copy Destination:=o_Sheet.range("Z:Z")   ' Concepto de facturación
            o_Sheet.Columns(59).Copy Destination:=o_Sheet.range("AA:AA") ' Número de solicitud (Laboratorio o imagenología)
            o_Sheet.Columns(60).Copy Destination:=o_Sheet.range("AB:AB") ' Número de requisición
            
            'Eliminar Columnas no utilizadas
            For intColumnas = 63 To 29 Step -1
                o_Sheet.Columns(intColumnas).Delete
            Next intColumnas
            
            o_Excel.ActiveWorkbook.ActiveSheet.Name = "Resumen de cargos" ' Nombere de la hoja
            
            'Titulos de las columnas
            intcontador = LBound(strTitulos)
            Do While intcontador <= UBound(strTitulos)
                o_Sheet.cells(3, intcontador + 1).Value = strTitulos(intcontador)
                intcontador = intcontador + 1
            Loop
            
            'Colocar la información y ajustar
            intRenglones = 0
            o_Sheet.range("A4").Select

            Do While Not IsEmpty(o_Sheet.Application.ActiveCell.offset(intRenglones, 0).Value)
            
                For intColumnas = 0 To 27
                    Select Case intColumnas
                        Case 0: strValor = o_Sheet.Application.ActiveCell.offset(intRenglones, intColumnas).Value
                        Case 1: strValor = Format(o_Sheet.Application.ActiveCell.offset(intRenglones, intColumnas).Value, "yyyy")
                        Case 2: strValor = "'" & Format(o_Sheet.Application.ActiveCell.offset(intRenglones, intColumnas).Value, "mm")
                        Case 3: strValor = "'" & Format(o_Sheet.Application.ActiveCell.offset(intRenglones, intColumnas).Value, "dd")
                        Case 5:
                            If o_Sheet.Application.ActiveCell.offset(intRenglones, intColumnas).Value = "" Then
                                strValor = o_Sheet.Application.ActiveCell.offset(intRenglones, intColumnas - 1).Value
                            Else
                                If Len(o_Sheet.Application.ActiveCell.offset(intRenglones, intColumnas).Value) < 10 Then
                                    strValor = "'00" & o_Sheet.Application.ActiveCell.offset(intRenglones, intColumnas).Value
                                Else
                                    strValor = "'" & o_Sheet.Application.ActiveCell.offset(intRenglones, intColumnas).Value
                                End If
                            End If
                        Case 26, 27:
                            If CLng(o_Sheet.Application.ActiveCell.offset(intRenglones, intColumnas).Value) = 0 Then
                                strValor = ""
                            Else
                                strValor = o_Sheet.Application.ActiveCell.offset(intRenglones, intColumnas).Value
                            End If
                    End Select
                    
                    If intColumnas <= 3 Or intColumnas = 5 Or intColumnas = 26 Or intColumnas = 27 Then
                        o_Sheet.Application.ActiveCell.offset(intRenglones, intColumnas).Value = IIf(strValor = Null, "", strValor)
                    End If
                  
                Next intColumnas
                intRenglones = intRenglones + 1
                pgbBarra.Value = pgbBarra.Value + 1
            Loop
            
            'Formatear Celdas
            o_Sheet.range("A:A").NumberFormat = "dd/mmm/yyyy"
            o_Sheet.range("B:B").NumberFormat = "0#"
            o_Sheet.range("C:C").NumberFormat = "0#"
            o_Sheet.range("D:D").NumberFormat = "0#"
            o_Sheet.range("J:J").NumberFormat = "$ ###,###,###,##0.00"
            o_Sheet.range("K:K").NumberFormat = "$ ###,###,###,##0.00"
            o_Sheet.range("L:L").NumberFormat = "$ ###,###,###,##0.00"
            o_Sheet.range("M:M").NumberFormat = "$ ###,###,###,##0.00"
            o_Sheet.range("N:N").NumberFormat = "$ ###,###,###,##0.00"
            o_Sheet.range("O:O").NumberFormat = "$ ###,###,###,##0.00"
            o_Sheet.range("P:P").NumberFormat = "$ ###,###,###,##0.00"
            
            o_Sheet.range("C:C").HorizontalAlignment = 4
            o_Sheet.range("D:D").HorizontalAlignment = 4
            o_Sheet.range("E:E").HorizontalAlignment = 4
            o_Sheet.range("F:F").HorizontalAlignment = 4
            o_Sheet.range("Q:Q").HorizontalAlignment = 4
            
            o_Sheet.range("A3:AA3").HorizontalAlignment = -4131
            
            o_Excel.range("3:3").Select
            With o_Excel.Selection
                .WrapText = True
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .MergeCells = False
            End With

            'Filtros
            'o_Excel.range("3:3").Select
            'o_Excel.selection.AutoFilter
            
            'Borrar la columna 5
            o_Sheet.Columns(5).Delete

            'AutoAjustar
            o_Excel.Selection.CurrentRegion.Columns.AutoFit
                        
            For intAutoFit = 1 To 27
                Select Case intAutoFit
                    Case 1, 26, 27
                        o_Excel.Columns(intAutoFit).Columnwidth = (intAnchoColumna / 2) - 40 'Ancho de columna a 20
                    Case 2, 3, 4, 5, 6, 8
                        o_Excel.Columns(intAutoFit).Columnwidth = (intAnchoColumna / 2) - 45 'Ancho de columna a 15
                    Case 9, 10, 11, 12, 13, 14, 15, 16, 17
                        o_Excel.Columns(intAutoFit).Columnwidth = (intAnchoColumna / 2) - 35 'Ancho de columna a 35
                    Case 19, 20, 21, 22, 23, 24, 25
                        o_Excel.Columns(intAutoFit).Columnwidth = (intAnchoColumna / 2) - 10 'Ancho de columna a 50
                    Case Else
                        o_Excel.Columns(intAutoFit).Columns.AutoFit
                End Select
            Next intAutoFit
            
            o_Excel.Selection.CurrentRegion.Rows.AutoFit
            
            'Titulos del reporte
            o_Sheet.range("A:A").HorizontalAlignment = -4131
            
            o_Excel.cells(1, 1).Value = Trim(vgstrNombreHospitalCH)
            o_Excel.range("A1:AA1").Select
            o_Excel.Selection.Merge
            o_Sheet.range("A1:AA1").HorizontalAlignment = -4108
            
            o_Excel.cells(2, 1).Value = "RESUMEN DE CARGOS DESGLOSADO POR CARGO DEL " & UCase(vlstrFechaInicial) & " AL " & UCase(vlstrFechaFinal)
            o_Excel.range("A2:AA2").Select
            o_Excel.Selection.Merge
            o_Sheet.range("A2:AA2").HorizontalAlignment = -4108
            
            Screen.MousePointer = vbDefault
            freBarra.Visible = False
            
            'La información ha sido exportada exitosamente
            MsgBox SIHOMsg(1185), vbOKOnly + vbInformation, "Mensaje"
                    
            o_Excel.Visible = True
        
            Set o_Excel = Nothing
        
            rsReporte.Close
            
        End If
    End If
    
    btnEnable
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pExporta"))
End Sub

Private Function fblnValidos() As Boolean
    On Error GoTo NotificaError
    
    fblnValidos = True
    If Not optPresentacion(0).Value Then
        If CDate(mskFecha.Text) > CDate(mskFechaFin.Text) Then
            fblnValidos = False
            MsgBox SIHOMsg(64), vbOKOnly + vbInformation, "Mensaje" '¡Rango de fechas no válido!
            mskFecha.SetFocus
        End If
    End If
   
    Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fblnValidos"))
End Function

Private Sub mskFechaFin_GotFocus()
    pSelMkTexto mskFechaFin
End Sub

Private Sub mskFechaFin_LostFocus()
    If Not IsDate(mskFechaFin.Text) Then
       MsgBox SIHOMsg(29), vbExclamation, "Mensaje"
       Me.mskFechaFin.SetFocus
    End If
End Sub

Private Sub optPresentacion_Click(Index As Integer)
    'mskFechaFin.Mask = ""
    If Index = 0 Then
        mskFecha.Enabled = True
        mskFechaFin.Enabled = True
        
        cboEjercicio.Enabled = False
        cboMes.Enabled = False
    
        lblAl.Visible = False
        mskFechaFin.Visible = False
        lblConcepto.Enabled = False
        cboConcepto.Enabled = False
        
        cmdVistaPreliminar.Enabled = True
        cmdImprimir.Enabled = True
        cmdExportar.Enabled = False
        FraSoloCargos.Enabled = False
        ChkIncluirCargos.Enabled = False
        mskFecha.ToolTipText = "Ingresa fecha de reporte"
    Else
        If Index = 1 Or Index = 2 Then
            mskFecha.Enabled = True
            mskFechaFin.Enabled = True
            
            cboEjercicio.Enabled = False
            cboMes.Enabled = False
            
            lblAl.Visible = True
            mskFechaFin.Visible = True
            lblConcepto.Enabled = True
            cboConcepto.Enabled = True
            
            cmdVistaPreliminar.Enabled = True
            cmdImprimir.Enabled = True
            cmdExportar.Enabled = False
            FraSoloCargos.Enabled = False
            ChkIncluirCargos.Enabled = False
            
            '#Region Caco 19209
            If Index = 2 Then
                cmdExportar.Enabled = True
                FraSoloCargos.Enabled = True
                ChkIncluirCargos.Enabled = True
            End If
            'End Region
        Else
            mskFecha.Enabled = False
            mskFechaFin.Enabled = False
            
            cboEjercicio.Enabled = True
            cboMes.Enabled = True
        
            lblAl.Visible = True
            mskFechaFin.Visible = True
            lblConcepto.Enabled = True
            cboConcepto.Enabled = True
            
            cmdVistaPreliminar.Enabled = False
            cmdImprimir.Enabled = False
            cmdExportar.Enabled = True
            FraSoloCargos.Enabled = False
            ChkIncluirCargos.Enabled = False
        End If
        mskFecha.ToolTipText = "Fecha inicial"
        mskFechaFin.ToolTipText = "Fecha final"
    End If
    cboConcepto.ListIndex = 0
End Sub


'#region caso 19744
    Private Sub btnEnable(Optional blnValor As Boolean = True)
        cmdVistaPreliminar.Enabled = blnValor
        cmdImprimir.Enabled = blnValor
        cmdExportar.Enabled = blnValor
    End Sub
'#End Region

