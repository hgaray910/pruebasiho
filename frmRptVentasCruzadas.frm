VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRptVentasCruzadas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ventas cruzadas"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5895
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   5895
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame freBarra 
      Height          =   815
      Left            =   715
      TabIndex        =   14
      Top             =   3000
      Visible         =   0   'False
      Width           =   4465
      Begin MSComctlLib.ProgressBar pgbBarra 
         Height          =   300
         Left            =   45
         TabIndex        =   15
         Top             =   480
         Width           =   4380
         _ExtentX        =   7726
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
         TabIndex        =   16
         Top             =   180
         Width           =   4370
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
   End
   Begin VB.Frame Frame3 
      Height          =   5915
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   5655
      Begin VB.ListBox lstDepartamento 
         Height          =   1860
         ItemData        =   "frmRptVentasCruzadas.frx":0000
         Left            =   120
         List            =   "frmRptVentasCruzadas.frx":0007
         Style           =   1  'Checkbox
         TabIndex        =   4
         ToolTipText     =   "Departamentos"
         Top             =   3600
         Width           =   5415
      End
      Begin VB.ListBox lstConceptosFact 
         Height          =   1860
         ItemData        =   "frmRptVentasCruzadas.frx":001C
         Left            =   120
         List            =   "frmRptVentasCruzadas.frx":0023
         Style           =   1  'Checkbox
         TabIndex        =   3
         ToolTipText     =   "Conceptos de facturación"
         Top             =   1440
         Width           =   5415
      End
      Begin VB.CheckBox chkdetalle 
         Caption         =   "Mostrar sólo facturados"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "Mostrar sólo los cargos facturados"
         Top             =   5560
         Width           =   2535
      End
      Begin VB.ComboBox cboProcedencia 
         Height          =   315
         Left            =   2040
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Grupo de tipos de clientes"
         Top             =   240
         Width           =   3495
      End
      Begin MSComCtl2.DTPicker dtpRango 
         CausesValidation=   0   'False
         Height          =   315
         Index           =   1
         Left            =   3960
         TabIndex        =   2
         ToolTipText     =   "Fecha final"
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MMM/yyyy"
         Format          =   92274691
         CurrentDate     =   39204
      End
      Begin MSComCtl2.DTPicker dtpRango 
         CausesValidation=   0   'False
         Height          =   315
         Index           =   0
         Left            =   1200
         TabIndex        =   1
         ToolTipText     =   "Fecha inicial"
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MMM/yyyy"
         Format          =   92798979
         CurrentDate     =   39204
      End
      Begin VB.Label Label5 
         Caption         =   "Departamento"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   3360
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Concepto de facturación"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha final"
         Height          =   255
         Left            =   3000
         TabIndex        =   11
         Top             =   750
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha inicial"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   750
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Grupo de tipos de cliente"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   270
         Width           =   1935
      End
   End
   Begin VB.Frame Frame6 
      Height          =   725
      Left            =   2155
      TabIndex        =   8
      Top             =   5950
      Width           =   1585
      Begin VB.CommandButton cmdExportar 
         Caption         =   "Exportar a Excel"
         Height          =   495
         Left            =   60
         TabIndex        =   6
         ToolTipText     =   "Exportar a Excel"
         Top             =   150
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmRptVentasCruzadas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private lrptReporteVer As CRAXDRT.Report
Public vgstrmodulo As String
Dim vllbnSalir As Boolean

Dim o_Excel As Object
Dim o_ExcelAbrir As Object
Dim o_Libro As Object
Dim o_Sheet As Object

Private Sub cmdExportar_Click()
    Dim alstrParametros(16) As String
    Dim rsReporte As ADODB.Recordset
    Dim intIndex As Integer
    Dim strAgrupar As String
    Dim strTipoIngreso As String
    Dim strTipoIngreso2 As String
    Dim strParametros As String
    Dim rptReporte As CRAXDRT.Report
    Dim lngVal As Long
    Dim intcontadorREN As Integer
    Dim intcontadorCOL As Integer
    
    Dim rs As ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Dim rsImporte As ADODB.Recordset
    Dim vgstrParametrosSP As String
    Dim strInicial As String
    Dim strFinal As String
    Dim strInicialP As String
    Dim intNivel As Integer
    Dim vlstrCuentaInicial As String
    Dim vlstrCuentaFinal As String
    Dim intRenglones As Integer
    Dim vlintVueltas As Integer
    Dim vllngRengArregloFechas As Long
    Dim vlIntCiclo As Integer
    Dim vlIntCiclo2 As Integer
    Dim vlIntCiclo3 As Integer
    Dim vlintNumeroColumna As Integer   'Número de columna que se ha utilizado en la hoja de calculo
    Dim vlArchivo As String
    Dim vlblnIncluiraMesesForzoso As Boolean
    Dim vlblnYaPasoCuenta As Boolean
    Dim vlStrRenglon As String
    Dim vlStrColumna As String
    Dim vlintAumentosBarra As Double
    Dim vlStrColumnaFinal As String
    Dim vlintRenglonColocado As Integer
    Dim vlintColumnaColocada As Integer
    Dim intFunct As Long
    Dim vlContadorColumnas As Double
    Dim vlContadorRenglones As Double
    Dim vlContadorVueltas As Double
    Dim intcontador As Integer
    Dim vldblImporteDeptos As Double
    Dim vldblImporteDescuentos As Double
'    Dim vldblImporteNotasCredito As Double
'    Dim vldblImporteNotasCargo As Double

    If Not fValidaInformacion Then Exit Sub

    pgbBarra.Value = 0
    freBarra.Visible = True
    
    Set o_Excel = CreateObject("Excel.Application")
    Set o_Libro = o_Excel.Workbooks.Add
    Set o_Sheet = o_Libro.Worksheets(1)
    
    o_Excel.ActiveWorkbook.ActiveSheet.Name = "Ventas cruzadas"

    'Da formato a la hoja completa
    o_Sheet.range("A:IV").HorizontalAlignment = -4152 'Derecha
    o_Sheet.range("A:IV").NumberFormat = "$ ###,###,###,##0.00"
    o_Sheet.range("A:IV").Font.Size = 8
    o_Sheet.range("A:IV").Font.Name = "Times New Roman"
    o_Sheet.range("A:IV").Font.Bold = False

    o_Sheet.range("A:A").HorizontalAlignment = -4131 'Izquierda
    o_Sheet.range("A:A").NumberFormat = "@"
    
    o_Sheet.range("A1:IV1").HorizontalAlignment = -4131 'Izquierda
    o_Sheet.range("A1:IV1").NumberFormat = "@"

    'Datos del repote y su formato
'    o_Sheet.Range("A2:B4").HorizontalAlignment = -4108 'Centrado
'    o_Sheet.Range("A2:B6").NumberFormat = "@"
    
'    o_Sheet.Range("A2:B2").Font.Size = 12
'    o_Sheet.Range("A2:B2").Font.Name = "Times New Roman"
'    o_Sheet.Range("A2:B2").Font.Bold = True
'    o_Excel.Cells(2, 1).Value = StrConv(Trim(vgstrNombreHospitalCH), vbUpperCase)
    
'    o_Sheet.Range("A3:B4").Font.Size = 10
'    o_Sheet.Range("A3:B4").Font.Name = "Times New Roman"
'    o_Sheet.Range("A3:B4").Font.Bold = False
'    o_Excel.Cells(3, 1).Value = "VENTAS CRUZADAS"
'    o_Excel.Cells(4, 1).Value = "DEL " & StrConv(CStr(Format(dtpRango(0).Value, "dd/MMM/yyyy")), vbUpperCase) & " AL " & StrConv(CStr(Format(dtpRango(1).Value, "dd/MMM/yyyy")), vbUpperCase)
    
'    o_Sheet.Range("A6:B6").Font.Size = 10
'    o_Excel.Cells(6, 1).Value = "GRUPO DE TIPOS DE CLIENTE  " & cboProcedencia.Text
    
'    o_Sheet.Range("A7:B7").Font.Size = 10
'    o_Excel.Cells(7, 1).Value = IIf(chkdetalle.Value, "INCLUYE SÓLO CARGOS FACTURADOS", "INCLUYE TODOS LOS CARGOS")
    
    o_Sheet.range("A1:B1").HorizontalAlignment = -4108 'Centrado
    o_Sheet.range("A1:B1").Font.Size = 10
    o_Sheet.range("A1:B1").Font.Name = "Times New Roman"
    o_Sheet.range("A1:B1").Font.Bold = False
    o_Excel.cells(1, 1).Value = "Conceptos de facturación"
    o_Excel.cells(1, 2).Value = "Departamentos"
        
'    o_Sheet.Range("A2:B2").MergeCells = True
'    o_Sheet.Range("A3:B3").MergeCells = True
'    o_Sheet.Range("A4:B4").MergeCells = True
    
    o_Sheet.range("B3").Select
    
    o_Excel.ActiveWindow.FreezePanes = True
                        
    'Llena la tabla temporal de los departamentos
    lngVal = 1
    frsEjecuta_SP "", "FN_PVDELTMPDEPARTAMENTOS", True, lngVal
    For intcontador = 1 To lstDepartamento.ListCount - 1
        If lstDepartamento.Selected(intcontador) = True Then
            lngVal = 1
            frsEjecuta_SP Str(lstDepartamento.ItemData(intcontador)), "FN_PVINSTMPDEPARTAMENTOS", True, lngVal
        End If
    Next intcontador
                        
    'Llena la tabla temporal de los conceptos de facturación
    lngVal = 1
    frsEjecuta_SP "", "FN_PVDELTMPCONCEPTOSFACT", True, lngVal
    For intcontador = 1 To lstConceptosFact.ListCount - 1
        If lstConceptosFact.Selected(intcontador) = True Then
            lngVal = 1
            frsEjecuta_SP Str(lstConceptosFact.ItemData(intcontador)), "FN_IVINSTMPCONCEPTOSFACT", True, lngVal
        End If
    Next intcontador
                        
    vlContadorRenglones = 0
    For intcontador = 1 To lstConceptosFact.ListCount - 1
        If lstConceptosFact.Selected(intcontador) = True Then
            vlContadorRenglones = vlContadorRenglones + 1
        End If
    Next intcontador
    
    vlContadorColumnas = 0
    For intcontador = 1 To lstDepartamento.ListCount - 1
        If lstDepartamento.Selected(intcontador) = True Then
            vlContadorColumnas = vlContadorColumnas + 1
        End If
    Next intcontador

    vlContadorVueltas = 75 / (vlContadorRenglones * vlContadorColumnas)
    
    Set rsImporte = frsEjecuta_SP(CStr(vgintClaveEmpresaContable) & "|" & IIf(cboProcedencia.ItemData(cboProcedencia.ListIndex) = 0, -1, cboProcedencia.ItemData(cboProcedencia.ListIndex)) & "|" & fstrFechaSQL(CStr(dtpRango(0).Value), "00:00:00", True) & "|" & fstrFechaSQL(CStr(dtpRango(1).Value), "23:59:58", True) & "|" & chkdetalle.Value, "SP_PVSELIMPORTEVENTASCRUZADAS")
    
    'Ciclo para colocar los renglones CONCEPTOS DE FACTURACIÓN
    vlintRenglonColocado = 3
    For intcontadorREN = 1 To lstConceptosFact.ListCount - 1
        If lstConceptosFact.Selected(intcontadorREN) = True Then
            o_Excel.cells(vlintRenglonColocado, 1).Value = lstConceptosFact.List(intcontadorREN)
            
            vldblImporteDeptos = 0
            
            'Ciclo para colocar las columnas DEPARTAMENTOS
            vlintColumnaColocada = 2
            For intcontadorCOL = 1 To lstDepartamento.ListCount - 1
                If lstDepartamento.Selected(intcontadorCOL) = True Then
                    o_Excel.cells(2, vlintColumnaColocada).Value = lstDepartamento.List(intcontadorCOL)
                                        
                    o_Excel.cells(vlintRenglonColocado, vlintColumnaColocada).Value = 0
                                        
                    If rsImporte.RecordCount > 0 Then
                        rsImporte.MoveFirst
                        Do While Not rsImporte.EOF
                            If Trim(rsImporte!DescripcionConcepto) = Trim(lstConceptosFact.List(intcontadorREN)) Then
                                If Trim(rsImporte!DescripcionDepartamento) = Trim(lstDepartamento.List(intcontadorCOL)) Then
                                    o_Excel.cells(vlintRenglonColocado, vlintColumnaColocada).Value = rsImporte!Importe
                                    vldblImporteDeptos = vldblImporteDeptos + rsImporte!Importe
                                    Exit Do
                                End If
                            End If
                    
                            rsImporte.MoveNext
                        Loop
                    End If
                    
'                    'Coloca el importe para ese CONCEPTO y ese DEPARTAMENTO
'                    Set rsImporte = frsEjecuta_SP(CStr(vgintClaveEmpresaContable) & "|" & IIf(cboProcedencia.ItemData(cboProcedencia.ListIndex) = 0, -1, cboProcedencia.ItemData(cboProcedencia.ListIndex)) & "|" & fstrFechaSQL(CStr(dtpRango(0).Value), "00:00:00", True) & "|" & fstrFechaSQL(CStr(dtpRango(1).Value), "23:59:58", True) & "|" & Str(lstConceptosFact.ItemData(intcontadorREN)) & "|" & Str(lstDepartamento.ItemData(intcontadorCOL)) & "|" & chkdetalle.Value, "SP_PVSELIMPORTEVENTASCRUZADAS")
'                    If rsImporte.RecordCount <> 0 Then
'                        o_Excel.Cells(vlintRenglonColocado, vlintColumnaColocada).Value = rsImporte!Importe
'                        vldblImporteDeptos = vldblImporteDeptos + rsImporte!Importe
'                    Else
'
'                    End If
                    
                    vlintColumnaColocada = vlintColumnaColocada + 1
                    
                    pgbBarra.Value = IIf((pgbBarra.Value + vlContadorVueltas) > 100, 100, (pgbBarra.Value + vlContadorVueltas))
                End If
            Next intcontadorCOL
                        
            'Si son todos los Grupos de tipos de clientes y todos los departamentos
'            vldblImporteNotasCredito = 0
'            vldblImporteNotasCargo = 0
'            If cboProcedencia.Text = "<TODOS>" And lstDepartamento.Selected(0) = True Then
                
'                o_Excel.Cells(2, vlintColumnaColocada).Value = "NOTAS DE CRÉDITO"
                
                'Coloca el importe de las NOTAS DE CRÉDITO
'                Set rsImporte = frsEjecuta_SP(fstrFechaSQL(CStr(dtpRango(0).Value), "00:00:00", True) & "|" & fstrFechaSQL(CStr(dtpRango(1).Value), "23:59:58", True) & "|" & Str(lstConceptosFact.ItemData(intcontadorREN)) & "|" & "CR", "SP_PVSELIMPORTENOTASVENTASCRUZ")
'                If rsImporte.RecordCount <> 0 Then
'                    o_Excel.Cells(vlintRenglonColocado, vlintColumnaColocada).Value = rsImporte!Importe
'                    vldblImporteNotasCredito = rsImporte!Importe
'                Else
'                    o_Excel.Cells(vlintRenglonColocado, vlintColumnaColocada).Value = 0
'                End If
'
'                vlintColumnaColocada = vlintColumnaColocada + 1
                
'                o_Excel.Cells(2, vlintColumnaColocada).Value = "NOTAS DE CARGO"
                
                'Coloca el importe de las NOTAS DE CARGO
'                Set rsImporte = frsEjecuta_SP(fstrFechaSQL(CStr(dtpRango(0).Value), "00:00:00", True) & "|" & fstrFechaSQL(CStr(dtpRango(1).Value), "23:59:58", True) & "|" & Str(lstConceptosFact.ItemData(intcontadorREN)) & "|" & "CA", "SP_PVSELIMPORTENOTASVENTASCRUZ")
'                If rsImporte.RecordCount <> 0 Then
'                    o_Excel.Cells(vlintRenglonColocado, vlintColumnaColocada).Value = rsImporte!Importe
'                    vldblImporteNotasCargo = rsImporte!Importe
'                Else
'                    o_Excel.Cells(vlintRenglonColocado, vlintColumnaColocada).Value = 0
'                End If
'
'                vlintColumnaColocada = vlintColumnaColocada + 1
'            End If
            
            o_Excel.cells(2, vlintColumnaColocada).Value = "TOTAL"
'            o_Excel.Cells(vlintRenglonColocado, vlintColumnaColocada).Value = vldblImporteDeptos - vldblImporteNotasCredito + vldblImporteNotasCargo
            o_Excel.cells(vlintRenglonColocado, vlintColumnaColocada).Value = vldblImporteDeptos
            
            vlintRenglonColocado = vlintRenglonColocado + 1
        End If
    Next intcontadorREN
    'Caso: 19297 OverFlow  "Entra a SP_PVSELDESCVENTASCRUZADAS"
    vlContadorVueltas = 25 / (vlContadorColumnas)
    
    'Coloca los DESCUENTOS y los TOTALES
    o_Excel.cells(vlintRenglonColocado, 1).Value = "DESCUENTOS"
    o_Excel.cells(vlintRenglonColocado + 1, 1).Value = "TOTAL"
    
    vlintColumnaColocada = 2
    vldblImporteDescuentos = 0
    For intcontadorCOL = 1 To lstDepartamento.ListCount - 1
        If lstDepartamento.Selected(intcontadorCOL) = True Then
            'Coloca el DESCUENTO para todos los conceptos y ese departamento
            Set rsImporte = frsEjecuta_SP(CStr(vgintClaveEmpresaContable) & "|" & IIf(cboProcedencia.ItemData(cboProcedencia.ListIndex) = 0, -1, cboProcedencia.ItemData(cboProcedencia.ListIndex)) & "|" & fstrFechaSQL(CStr(dtpRango(0).Value), "00:00:00", True) & "|" & fstrFechaSQL(CStr(dtpRango(1).Value), "23:59:58", True) & "|" & Str(lstDepartamento.ItemData(intcontadorCOL)) & "|" & chkdetalle.Value, "SP_PVSELDESCVENTASCRUZADAS")
            If rsImporte.RecordCount <> 0 Then
                o_Excel.cells(vlintRenglonColocado, vlintColumnaColocada).Value = rsImporte!Importe
                vldblImporteDescuentos = vldblImporteDescuentos + rsImporte!Importe
            Else
                o_Excel.cells(vlintRenglonColocado, vlintColumnaColocada).Value = 0
            End If
            
            o_Excel.cells(vlintRenglonColocado + 1, vlintColumnaColocada).FormulaR1C1 = "=SUM(R[-" & vlintRenglonColocado - 2 & "]C:R[-2]C) - R[-1]C"
            
'            ActiveCell.FormulaR1C1 = "=SUM(R[-3]C:R[-2]C) - R[-1]C"
'            Range("B6").Select
            
            vlintColumnaColocada = vlintColumnaColocada + 1
            
            pgbBarra.Value = IIf((pgbBarra.Value + vlContadorVueltas) > 100, 100, (pgbBarra.Value + vlContadorVueltas))
        End If
    Next intcontadorCOL
    
'    If cboProcedencia.Text = "<TODOS>" And lstDepartamento.Selected(0) = True Then
'        'Sumatoria descuento
'        o_Excel.Cells(vlintRenglonColocado, vlintColumnaColocada + 2).Value = vldblImporteDescuentos
'
'        'Totales
'        o_Excel.Cells(vlintRenglonColocado + 1, vlintColumnaColocada).FormulaR1C1 = "=SUM(R[-" & vlintRenglonColocado - 2 & "]C:R[-2]C)"
'        o_Excel.Cells(vlintRenglonColocado + 1, vlintColumnaColocada + 1).FormulaR1C1 = "=SUM(R[-" & vlintRenglonColocado - 2 & "]C:R[-2]C)"
'        o_Excel.Cells(vlintRenglonColocado + 1, vlintColumnaColocada + 2).FormulaR1C1 = "=SUM(R[-" & vlintRenglonColocado - 2 & "]C:R[-2]C) - R[-1]C"
'    Else
        'Sumatoria descuento
        o_Excel.cells(vlintRenglonColocado, vlintColumnaColocada).Value = vldblImporteDescuentos
        
        'Totales
        o_Excel.cells(vlintRenglonColocado + 1, vlintColumnaColocada).FormulaR1C1 = "=SUM(R[-" & vlintRenglonColocado - 2 & "]C:R[-2]C) - R[-1]C"
'    End If
           
    o_Sheet.PageSetup.Orientation = 2
    o_Excel.ActiveSheet.PageSetup.PrintTitleRows = "$1:$" & 2
    o_Excel.ActiveSheet.PageSetup.PrintTitleColumns = "$A:$B"
    o_Excel.ActiveSheet.PageSetup.PrintArea = ""
    o_Excel.ActiveSheet.PageSetup.CenterHeader = "&""Times New Roman,Negrita""&12" & Trim(vgstrNombreHospitalCH) & "&""Times New Roman,Normal""&10" & Chr(10) & "VENTAS CRUZADAS DEL " & StrConv(Format(dtpRango(0).Value, "DD/MMM/YYYY"), vbUpperCase) & " AL " & StrConv(Format(dtpRango(1).Value, "DD/MMM/YYYY"), vbUpperCase)
    o_Excel.ActiveSheet.PageSetup.LeftHeader = Chr(10) & Chr(10) & "&""Times New Roman,Normal""&10GRUPO DE TIPOS DE CLIENTE  " & StrConv(cboProcedencia.Text, vbUpperCase) & Chr(10) & IIf(chkdetalle.Value, "INCLUYE SÓLO CARGOS FACTURADOS", "INCLUYE TODOS LOS CARGOS") & Chr(10)
    o_Excel.ActiveSheet.PageSetup.RightHeader = ""
    o_Excel.ActiveSheet.PageSetup.CenterFooter = ""
    o_Excel.ActiveSheet.PageSetup.LeftFooter = "&""Times New Roman,Normal""&8Impresión " & Format(fdtmServerFechaHora, "DD/MM/YYYY HH:MM")
    o_Excel.ActiveSheet.PageSetup.RightFooter = "&""Times New Roman,Normal""&8Página &p de &n"
    o_Excel.ActiveSheet.PageSetup.LeftMargin = o_Excel.Application.InchesToPoints(0.590551181102362)
    o_Excel.ActiveSheet.PageSetup.RightMargin = o_Excel.Application.InchesToPoints(0.590551181102362)
    o_Excel.ActiveSheet.PageSetup.TopMargin = o_Excel.Application.InchesToPoints(1.37795275590551)
    o_Excel.ActiveSheet.PageSetup.BottomMargin = o_Excel.Application.InchesToPoints(0.78740157480315)
    o_Excel.ActiveSheet.PageSetup.HeaderMargin = o_Excel.Application.InchesToPoints(0.393700787401575)
    o_Excel.ActiveSheet.PageSetup.FooterMargin = o_Excel.Application.InchesToPoints(0.393700787401575)
                        
    o_Excel.Selection.CurrentRegion.Columns.AutoFit
    o_Excel.Selection.CurrentRegion.Rows.AutoFit
            
    o_Excel.ActiveSheet.DisplayPageBreaks = False
    o_Excel.cells(1).Select
    
    pgbBarra.Value = 100
    freBarra.Visible = False
    
    'La información ha sido exportada exitosamente
    MsgBox SIHOMsg(1185), vbOKOnly + vbInformation, "Mensaje"
            
    o_Excel.Visible = True
    
    Me.MousePointer = 0
          
    Exit Sub
NotificaError:
    ' -- Cierra la hoja y la aplicación Excel
    If Not o_Libro Is Nothing Then: o_Libro.Close False
    If Not o_Excel Is Nothing Then: o_Excel.Quit
    If Not o_Excel Is Nothing Then Set o_Excel = Nothing
    If Not o_Libro Is Nothing Then Set o_Libro = Nothing
    If Not o_Sheet Is Nothing Then Set o_Sheet = Nothing
    If Err.Number = 32755 Then
        Me.MousePointer = 0
        Exit Sub
    End If
    If Err.Number <> 1004 Then MsgBox Err.Description, vbCritical
End Sub

Private Sub dtpRango_GotFocus(Index As Integer)
    dtpRango(Index).CustomFormat = "dd/MM/yyyy"
End Sub

Private Sub dtpRango_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub dtpRango_LostFocus(Index As Integer)
    dtpRango(Index).CustomFormat = "dd/MMM/yyyy"
End Sub

Private Sub Form_Activate()
    If cboProcedencia.ListCount = 1 Then
        '¡No existen grupos de tipos de clientes configurados!
        MsgBox SIHOMsg(1462), vbOKOnly + vbExclamation, "Mensaje"
        Unload Me
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
        On Error GoTo NotificaError
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If KeyAscii = 27 Then
            Unload Me
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyPress"))
End Sub

Private Sub Form_Load()
    Dim rs As ADODB.Recordset
    Dim lngContador As Integer
    Dim vlstrSentencia As String
    
    Me.Icon = frmMenuPrincipal.Icon
    pIniciaForma
    
    vlstrSentencia = "select smicveconcepto clave, trim(chrdescripcion) descripcion from pvconceptofacturacion where  bitactivo = 1 order by trim(chrdescripcion)"
    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    With lstConceptosFact
        .Clear
        
        .AddItem "<TODOS>"
        .ItemData(.newIndex) = 0
        .Selected(.newIndex) = True
        
        Do While Not rs.EOF
            .AddItem rs!Descripcion
            .ItemData(.newIndex) = rs!clave
            .Selected(.newIndex) = True
            rs.MoveNext
        Loop
    End With
    rs.Close
    
    lstConceptosFact.ListIndex = 0
    
    vlstrSentencia = "select smicvedepartamento clave, trim(vchdescripcion) descripcion from nodepartamento where tnyclaveempresa = " & vgintClaveEmpresaContable & " and bitestatus = 1 order by trim(vchdescripcion)"
    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    With lstDepartamento
        .Clear
        
        .AddItem "<TODOS>"
        .ItemData(.newIndex) = 0
        .Selected(.newIndex) = True
        
        Do While Not rs.EOF
            .AddItem rs!Descripcion
            .ItemData(.newIndex) = rs!clave
            .Selected(.newIndex) = True
            rs.MoveNext
        Loop
    End With
    rs.Close
    
    lstDepartamento.ListIndex = 0
    
'    pInstanciaReporte lrptReporteVer, "rpt.rpt"
        
End Sub

Private Sub pIniciaForma()
    Dim intIndex As Integer
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    
    dtpRango(0).Value = DateSerial(Year(Date), Month(Date), 1)
    dtpRango(1).Value = DateSerial(Year(Date), Month(Date), Day(Date))
    
    For intIndex = 0 To lstConceptosFact.ListCount - 1
        lstConceptosFact.Selected(intIndex) = True
    Next
    
    For intIndex = 0 To lstDepartamento.ListCount - 1
        lstDepartamento.Selected(intIndex) = True
    Next
        
    strSQL = "select intcvegrupo clave, trim(vchdescripciongrupo) descrip from pvgrupotipocliente order by trim(vchdescripciongrupo)"
    Set rs = frsRegresaRs(strSQL)
    Do Until rs.EOF
        cboProcedencia.AddItem rs!Descrip
        cboProcedencia.ItemData(cboProcedencia.newIndex) = rs!clave
        rs.MoveNext
    Loop
    rs.Close
        
    cboProcedencia.AddItem "<TODOS>", 0
    cboProcedencia.ItemData(cboProcedencia.newIndex) = 0
    cboProcedencia.ListIndex = 0
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = 0
End Sub

Private Sub lstDepartamento_ItemCheck(Item As Integer)
    Dim intIndex As Integer
    Dim vlblnHayApagado As Boolean
    Dim vllngSeleccion As Long

    With lstDepartamento
        If .ListIndex = 0 Then
            If .Selected(.ListIndex) = True Then
                For intIndex = 1 To .ListCount - 1
                    .Selected(intIndex) = True
                Next
            Else
                For intIndex = 1 To .ListCount - 1
                    .Selected(intIndex) = False
                Next
            End If
            
            .ListIndex = 0
        Else
            vllngSeleccion = .ListIndex
            vlblnHayApagado = False
        
            For intIndex = 1 To .ListCount - 1
                If .Selected(intIndex) = False Then
                    vlblnHayApagado = True
                    Exit For
                End If
            Next
            
            If vlblnHayApagado Then
                .Selected(0) = False
            Else
                .Selected(0) = True
            End If
            
            .ListIndex = vllngSeleccion
        End If
    End With
End Sub

Private Sub lstConceptosFact_ItemCheck(Item As Integer)
    Dim intIndex As Integer
    Dim vlblnHayApagado As Boolean
    Dim vllngSeleccion As Long

    With lstConceptosFact
        If .ListIndex = 0 Then
            If .Selected(.ListIndex) = True Then
                For intIndex = 1 To .ListCount - 1
                    .Selected(intIndex) = True
                Next
            Else
                For intIndex = 1 To .ListCount - 1
                    .Selected(intIndex) = False
                Next
            End If
            
            .ListIndex = 0
        Else
            vllngSeleccion = .ListIndex
            vlblnHayApagado = False
        
            For intIndex = 1 To .ListCount - 1
                If .Selected(intIndex) = False Then
                    vlblnHayApagado = True
                    Exit For
                End If
            Next
            
            If vlblnHayApagado Then
                .Selected(0) = False
            Else
                .Selected(0) = True
            End If
            
            .ListIndex = vllngSeleccion
        End If
    End With
End Sub

Private Function fValidaInformacion() As Boolean
    fValidaInformacion = True
    
    If dtpRango(0).Value > fdtmServerFecha Then
       '¡La fecha debe ser menor o igual a la del sistema!
       MsgBox SIHOMsg(40), vbOKOnly + vbExclamation, "Mensaje"
       dtpRango(0).SetFocus
       fValidaInformacion = False
       Exit Function
    End If
    
    If dtpRango(1).Value > fdtmServerFecha Then
       '¡La fecha debe ser menor o igual a la del sistema!
       MsgBox SIHOMsg(40), vbOKOnly + vbExclamation, "Mensaje"
       dtpRango(1).SetFocus
       fValidaInformacion = False
       Exit Function
    End If
    
    If dtpRango(0).Value > dtpRango(1).Value Then
       '¡La fecha final debe ser mayor a la fecha inicial!
       MsgBox SIHOMsg(379), vbOKOnly + vbExclamation, "Mensaje"
       dtpRango(1).SetFocus
       fValidaInformacion = False
       Exit Function
    End If
    
    If lstConceptosFact.SelCount = 0 Then
       '¡Dato no válido, seleccione un valor de la lista!
       MsgBox SIHOMsg(3), vbOKOnly + vbExclamation, "Mensaje"
       Me.lstConceptosFact.SetFocus
       fValidaInformacion = False
       Exit Function
    End If
    
    If lstDepartamento.SelCount = 0 Then
       '¡Dato no válido, seleccione un valor de la lista!
       MsgBox SIHOMsg(3), vbOKOnly + vbExclamation, "Mensaje"
       Me.lstDepartamento.SetFocus
       fValidaInformacion = False
       Exit Function
    End If
        
End Function

Private Sub pMargenes(vlIntRen As Integer, vlintCol As Integer, vlblnSuperior As Boolean, vlblnInferior As Boolean)
    On Error GoTo NotificaError
    'If o_Excel.Cells(vlIntRen, vlintCol).Borders(1).LineStyle <> 1 Then o_Excel.Cells(vlIntRen, vlintCol).Borders(1).LineStyle = 1
    'If o_Excel.Cells(vlIntRen, vlintCol).Borders(2).LineStyle <> 1 Then o_Excel.Cells(vlIntRen, vlintCol).Borders(2).LineStyle = 1
    If o_Excel.cells(vlIntRen, vlintCol).Borders(3).LineStyle <> 1 And vlblnSuperior Then o_Excel.cells(vlIntRen, vlintCol).Borders(3).LineStyle = 1
    If o_Excel.cells(vlIntRen, vlintCol).Borders(4).LineStyle <> 1 And vlblnInferior Then o_Excel.cells(vlIntRen, vlintCol).Borders(4).LineStyle = 1
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pMargenes"))
End Sub
