VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmrptIngresosporDepartamento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingresos por departamento"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7920
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   7920
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkCanceladas 
      Caption         =   "Mostrar documentos cancelados"
      Height          =   375
      Left            =   120
      TabIndex        =   18
      ToolTipText     =   "Mostrar documentos cancelados"
      Top             =   2640
      Width           =   2655
   End
   Begin VB.Frame freBarra 
      Height          =   815
      Left            =   1680
      TabIndex        =   23
      Top             =   960
      Visible         =   0   'False
      Width           =   4465
      Begin MSComctlLib.ProgressBar pgbBarra 
         Height          =   300
         Left            =   45
         TabIndex        =   24
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
         TabIndex        =   26
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
      Begin VB.Label Label9 
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
         TabIndex        =   25
         Top             =   225
         Width           =   4365
      End
   End
   Begin VB.Frame Frame6 
      Height          =   735
      Left            =   2790
      TabIndex        =   22
      Top             =   2520
      Width           =   2610
      Begin VB.CommandButton cmdExportar 
         Caption         =   "Exportar"
         Height          =   495
         Left            =   1080
         TabIndex        =   21
         ToolTipText     =   "Exportar a Excel"
         Top             =   165
         Width           =   1455
      End
      Begin VB.CommandButton cmdImprimir 
         Height          =   495
         Left            =   570
         Picture         =   "frmrptIngresosporDepartamento.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Imprimir"
         Top             =   165
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdVista 
         Height          =   495
         Left            =   75
         Picture         =   "frmrptIngresosporDepartamento.frx":01A2
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Vista previa"
         Top             =   165
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Incluir"
      Height          =   1065
      Left            =   3970
      TabIndex        =   17
      Top             =   1440
      Width           =   2295
      Begin VB.CheckBox chkNotas 
         Caption         =   "Notas de cargo y crédito"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "Incluir notas en el reporte"
         Top             =   480
         Width           =   2055
      End
      Begin VB.CheckBox chkTickets 
         Caption         =   "Tickets"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   "Incluir tickets en el reporte"
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox chkFacturas 
         Caption         =   "Facturas"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "Incluir facturas en el reporte"
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   705
      Left            =   120
      TabIndex        =   15
      Top             =   720
      Width           =   7695
      Begin VB.ComboBox cboDepartamento 
         Height          =   315
         Left            =   1470
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Selección del departamento"
         Top             =   240
         Width           =   6105
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Departamento "
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   300
         Width           =   1050
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Presentaciones"
      Height          =   1065
      Left            =   6285
      TabIndex        =   14
      Top             =   1440
      Width           =   1530
      Begin VB.OptionButton optPresentacion 
         Caption         =   "Detallada"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "Presentación detallada"
         Top             =   720
         Width           =   1050
      End
      Begin VB.OptionButton optPresentacion 
         Caption         =   "Concentrada"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   8
         ToolTipText     =   "Presentación concentrada"
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Rango de fechas del documento"
      Height          =   1065
      Left            =   120
      TabIndex        =   11
      Top             =   1440
      Width           =   3845
      Begin MSMask.MaskEdBox mskFechaInicio 
         Height          =   315
         Left            =   660
         TabIndex        =   3
         ToolTipText     =   "Fecha de inicio"
         Top             =   420
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mmm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskFechaFin 
         Height          =   315
         Left            =   2480
         TabIndex        =   4
         ToolTipText     =   "Fecha final"
         Top             =   420
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mmm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         Height          =   195
         Left            =   2000
         TabIndex        =   13
         Top             =   480
         Width           =   420
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   465
      End
   End
   Begin VB.Frame Frame7 
      Height          =   675
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7695
      Begin VB.ComboBox cboHospital 
         Height          =   315
         Left            =   1470
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Seleccione la empresa contable"
         Top             =   240
         Width           =   6120
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Empresa"
         Height          =   195
         Left            =   130
         TabIndex        =   10
         Top             =   300
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmrptIngresosporDepartamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub pImprime(strDestino As String)
    Dim rsReporte As New ADODB.Recordset
    Dim rptReporte As CRAXDRT.Report
    Dim alstrParametros(3) As String
    Dim intSeleccion As Integer
        
    'Validaciones.
    If Me.cboDepartamento.ListIndex = -1 Then
       '3 ¡Dato no válido, seleccione un valor de la lista!
       MsgBox SIHOMsg(3), vbExclamation + vbOKOnly, "Mensaje"
       Me.cboDepartamento.SetFocus
       Exit Sub
    End If
    
    If Not IsDate(Me.mskFechaInicio) Then
       '29 ¡Fecha no válida! Formato de fecha dd/mm/aaaa.
       MsgBox SIHOMsg(29), vbExclamation + vbOKOnly, "Mensaje"
       Me.mskFechaInicio.SetFocus
       Exit Sub
    End If
    
    If Not IsDate(Me.mskFechaFin) Then
       '29 ¡Fecha no válida! Formato de fecha dd/mm/aaaa.
       MsgBox SIHOMsg(29), vbExclamation + vbOKOnly, "Mensaje"
       Me.mskFechaFin.SetFocus
       Exit Sub
    End If
    
    If CDate(Me.mskFechaInicio) > CDate(Me.mskFechaFin) Then
       '64 ¡Rango de fechas no válido!
       MsgBox SIHOMsg(64), vbExclamation + vbOKOnly, "Mensaje"
       Me.mskFechaInicio.SetFocus
       Exit Sub
    End If
    
    If Me.chkFacturas.Value = vbUnchecked And Me.chkNotas.Value = vbUnchecked And Me.chkTickets.Value = vbUnchecked Then
       '819 No se han seleccionado datos.
       MsgBox SIHOMsg(819), vbExclamation + vbOKOnly, "Mensaje"
       Me.chkFacturas.SetFocus
       Exit Sub
    End If
       
    vgstrParametrosSP = Me.cboHospital.ItemData(Me.cboHospital.ListIndex) & _
                       "|" & IIf(Me.cboDepartamento.Text = "<TODOS>", -1, IIf(Me.cboDepartamento.Text = "<PAQUETES>", -2, Me.cboDepartamento.ItemData(Me.cboDepartamento.ListIndex))) & _
                       "|" & fstrFechaSQL(mskFechaInicio.Text, "00:00:00") & _
                       "|" & fstrFechaSQL(mskFechaFin.Text, "23:59:59") & _
                       "|" & IIf(Me.chkFacturas.Value = vbChecked, 1, 0) & _
                       "|" & IIf(Me.chkNotas.Value = vbChecked, 1, 0) & _
                       "|" & IIf(Me.chkTickets.Value = vbChecked, 1, 0) & _
                       "|" & IIf(chkCanceladas.Value = vbChecked, 1, 0)
    
    Set rsReporte = frsEjecuta_SP(vgstrParametrosSP, "Sp_PvRptIngresosDepartamento")

   If rsReporte.RecordCount <> 0 Then
      pInstanciaReporte rptReporte, IIf(Me.optPresentacion(1), "rptIngresosDepartamentoConcentrado.rpt", "rptIngresosDepartamentoDetallado.rpt")
        rptReporte.DiscardSavedData

        alstrParametros(0) = "NombreHospital;" & Trim(cboHospital.List(cboHospital.ListIndex)) & ";TRUE"
        alstrParametros(1) = "Fechas;" & "Desde " & Format(mskFechaInicio.Text, "dd/mmm/yyyy") & " hasta " & Format(mskFechaFin.Text, "dd/mmm/yyyy")
        alstrParametros(2) = "Departamento;" & cboDepartamento.Text
        
        pCargaParameterFields alstrParametros, rptReporte
        pImprimeReporte rptReporte, rsReporte, strDestino, "Ingresos por departamento"

    Else
        'No existe información con esos parámetros.
        MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje"
    End If
    rsReporte.Close

End Sub
Private Sub cboHospital_Click()
    Dim rs As New ADODB.Recordset
    Dim IntIndice As Integer
    
    If cboHospital.ListIndex <> -1 Then
       cboDepartamento.Clear
       vgstrParametrosSP = "-1|1|*|" & CStr(cboHospital.ItemData(cboHospital.ListIndex))
       Set rs = frsEjecuta_SP(vgstrParametrosSP, "Sp_Gnseldepartamento")
  
       If rs.RecordCount <> 0 Then
          pLlenarCboRs cboDepartamento, rs, 0, 1
       End If
       cboDepartamento.AddItem "<TODOS>", 0
       cboDepartamento.ItemData(cboDepartamento.newIndex) = -2
       cboDepartamento.AddItem "<PAQUETES>", 1
       cboDepartamento.ItemData(cboDepartamento.newIndex) = -1
       cboDepartamento.ListIndex = 0
    End If
End Sub

Private Sub cmdExportar_Click()
On Error GoTo NotificaError
    Dim rsAux As New ADODB.Recordset
    
    'Validaciones.
    If Me.cboDepartamento.ListIndex = -1 Then
       '3 ¡Dato no válido, seleccione un valor de la lista!
       MsgBox SIHOMsg(3), vbExclamation + vbOKOnly, "Mensaje"
       Me.cboDepartamento.SetFocus
       Exit Sub
    End If
    
    If Not IsDate(Me.mskFechaInicio) Then
       '29 ¡Fecha no válida! Formato de fecha dd/mm/aaaa.
       MsgBox SIHOMsg(29), vbExclamation + vbOKOnly, "Mensaje"
       Me.mskFechaInicio.SetFocus
       Exit Sub
    End If
    
    If Not IsDate(Me.mskFechaFin) Then
       '29 ¡Fecha no válida! Formato de fecha dd/mm/aaaa.
       MsgBox SIHOMsg(29), vbExclamation + vbOKOnly, "Mensaje"
       Me.mskFechaFin.SetFocus
       Exit Sub
    End If
    
    If CDate(Me.mskFechaInicio) > CDate(Me.mskFechaFin) Then
       '64 ¡Rango de fechas no válido!
       MsgBox SIHOMsg(64), vbExclamation + vbOKOnly, "Mensaje"
       Me.mskFechaInicio.SetFocus
       Exit Sub
    End If
    
    If Me.chkFacturas.Value = vbUnchecked And Me.chkNotas.Value = vbUnchecked And Me.chkTickets.Value = vbUnchecked Then
       '819 No se han seleccionado datos.
       MsgBox SIHOMsg(819), vbExclamation + vbOKOnly, "Mensaje"
       Me.chkFacturas.SetFocus
       Exit Sub
    End If
    
    If optPresentacion(1).Value = True Then 'Exportar reporte a excel en presentación concentrada
    
        vgstrParametrosSP = cboHospital.ItemData(Me.cboHospital.ListIndex) & _
                       "|" & IIf(cboDepartamento.Text = "<TODOS>", -1, IIf(cboDepartamento.Text = "<PAQUETES>", -2, cboDepartamento.ItemData(cboDepartamento.ListIndex))) & _
                       "|" & fstrFechaSQL(mskFechaInicio.Text, "00:00:00") & _
                       "|" & fstrFechaSQL(mskFechaFin.Text, "23:59:59") & _
                       "|" & IIf(chkFacturas.Value = vbChecked, 1, 0) & _
                       "|" & IIf(chkNotas.Value = vbChecked, 1, 0) & _
                       "|" & IIf(chkTickets.Value = vbChecked, 1, 0) & _
                       "|" & IIf(chkCanceladas.Value = vbChecked, 1, 0)
                       
        Set rsAux = frsEjecuta_SP(vgstrParametrosSP, "Sp_PvRptIngresoDeptoConcent")
        If rsAux.RecordCount <> 0 Then
            freBarra.Visible = True
            pgbBarra.Value = 0
            
            Set o_Excel = CreateObject("Excel.Application")
            Set o_Libro = o_Excel.Workbooks.Add
            Set o_Sheet = o_Libro.Worksheets(1)
            
            'datos del repote
            o_Excel.Cells(2, 1).Value = Trim(vgstrNombreHospitalCH)
            o_Excel.Cells(3, 1).Value = "INGRESOS POR DEPARTAMENTO"
            o_Excel.Cells(4, 1).Value = "Del " & CStr(Format(CDate(mskFechaInicio), "dd/MMM/yyyy")) & " Al " & CStr(Format(CDate(mskFechaFin), "dd/MMM/yyyy"))
            'columnas titulos
            o_Excel.Cells(6, 1).Value = "Departamento"
            o_Excel.Cells(6, 2).Value = "Razon social"
            o_Excel.Cells(6, 3).Value = "Cuenta"
            o_Excel.Cells(6, 4).Value = "Paciente"
            o_Excel.Cells(6, 5).Value = "Fecha"
            o_Excel.Cells(6, 6).Value = "Documento"
            o_Excel.Cells(6, 7).Value = "Folio"
            o_Excel.Cells(6, 8).Value = "Estado"
            o_Excel.Cells(6, 9).Value = "Importe"
            o_Excel.Cells(6, 10).Value = "Descuento"
            o_Excel.Cells(6, 11).Value = "Subtotal"
            o_Excel.Cells(6, 12).Value = "IVA"
            o_Excel.Cells(6, 13).Value = "Total"

            pgbBarra.Value = 30
        
            o_Sheet.Range("A6:M6").HorizontalAlignment = -4108
            o_Sheet.Range("A6:M6").VerticalAlignment = -4108
            o_Sheet.Range("A6:M6").WrapText = True
            o_Sheet.Range("A7").Select
            o_Excel.ActiveWindow.FreezePanes = True
            o_Sheet.Range("A6:M6").Interior.ColorIndex = 16
            o_Sheet.Range("A:A").ColumnWidth = 15
            o_Sheet.Range("B:B").ColumnWidth = 30
            o_Sheet.Range("C:C").ColumnWidth = 8
            o_Sheet.Range("D:D").ColumnWidth = 30
            o_Sheet.Range("E:E").ColumnWidth = 8
            o_Sheet.Range("F:F").ColumnWidth = 9
            o_Sheet.Range("G:G").ColumnWidth = 10
            o_Sheet.Range("H:H").ColumnWidth = 14
            o_Sheet.Range("I:I").ColumnWidth = 15
            o_Sheet.Range("J:J").ColumnWidth = 13
            o_Sheet.Range("K:K").ColumnWidth = 13
            o_Sheet.Range("L:L").ColumnWidth = 15
            o_Sheet.Range("M:M").ColumnWidth = 15
             
            o_Sheet.Range(o_Excel.Cells(rsAux.RecordCount + 6, 1), o_Excel.Cells(rsAux.RecordCount + 6, 13)).Borders(4).LineStyle = 1
            
            'info del rs
            o_Sheet.Range("A:M").Font.Size = 9
            o_Sheet.Range("A:M").Font.Name = "Times New Roman" '
            o_Sheet.Range("A:M").Font.Bold = False
            pgbBarra.Value = 70
            'titulos
            o_Sheet.Range("A6:M6").Font.Bold = True
            o_Sheet.Range(o_Excel.Cells(rsAux.RecordCount + 7, 1), o_Excel.Cells(rsAux.RecordCount + 7, 13)).Font.Bold = True
            o_Sheet.Range(o_Excel.Cells(2, 1), o_Excel.Cells(5, 1)).Font.Bold = True
            'centrado, auto ajustar texto, alinear medio
            o_Sheet.Range("H:M").NumberFormat = "$ ###,###,###,##0.00"
            
            'rs,maxRows,maxCols
            o_Sheet.Range("A7").CopyFromRecordset rsAux, , 13
            
            pgbBarra.Value = 100
            freBarra.Visible = False
            'La información ha sido exportada exitosamente
            MsgBox SIHOMsg(1185), vbOKOnly + vbInformation, "Mensaje"
            o_Excel.Visible = True
            
            Set o_Excel = Nothing
    
        Else
            'No existe información con esos parámetros.
            MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje"
        End If
        rsAux.Close
    
    ElseIf optPresentacion(0).Value = True Then 'Exportar reporte a excel en presentación detallada
    
        vgstrParametrosSP = cboHospital.ItemData(Me.cboHospital.ListIndex) & _
                       "|" & IIf(cboDepartamento.Text = "<TODOS>", -1, IIf(cboDepartamento.Text = "<PAQUETES>", -2, cboDepartamento.ItemData(cboDepartamento.ListIndex))) & _
                       "|" & fstrFechaSQL(mskFechaInicio.Text, "00:00:00") & _
                       "|" & fstrFechaSQL(mskFechaFin.Text, "23:59:59") & _
                       "|" & IIf(chkFacturas.Value = vbChecked, 1, 0) & _
                       "|" & IIf(chkNotas.Value = vbChecked, 1, 0) & _
                       "|" & IIf(chkTickets.Value = vbChecked, 1, 0) & _
                       "|" & IIf(chkCanceladas.Value = vbChecked, 1, 0)
                       
        Set rsAux = frsEjecuta_SP(vgstrParametrosSP, "Sp_PvRptIngresosDepartamento")
        If rsAux.RecordCount <> 0 Then
    
            freBarra.Visible = True
            pgbBarra.Value = 0
            
            Set o_Excel = CreateObject("Excel.Application")
            Set o_Libro = o_Excel.Workbooks.Add
            Set o_Sheet = o_Libro.Worksheets(1)
            
            'datos del repote
            o_Excel.Cells(2, 1).Value = Trim(vgstrNombreHospitalCH)
            o_Excel.Cells(3, 1).Value = "INGRESOS POR DEPARTAMENTO"
            o_Excel.Cells(4, 1).Value = "Del " & CStr(Format(CDate(mskFechaInicio), "dd/MMM/yyyy")) & " Al " & CStr(Format(CDate(mskFechaFin), "dd/MMM/yyyy"))
            'columnas titulos
            o_Excel.Cells(6, 1).Value = "Departamento"
            o_Excel.Cells(6, 2).Value = "Razon social"
            o_Excel.Cells(6, 3).Value = "Cuenta"
            o_Excel.Cells(6, 4).Value = "Paciente"
            o_Excel.Cells(6, 5).Value = "Fecha"
            o_Excel.Cells(6, 6).Value = "Documento"
            o_Excel.Cells(6, 7).Value = "Folio"
            o_Excel.Cells(6, 8).Value = "Estado"
            o_Excel.Cells(6, 9).Value = "Concepto facturación"
            o_Excel.Cells(6, 10).Value = "Cargo"
            o_Excel.Cells(6, 11).Value = "Importe"
            o_Excel.Cells(6, 12).Value = "Descuento"
            o_Excel.Cells(6, 13).Value = "Subtotal"
            o_Excel.Cells(6, 14).Value = "IVA"
            o_Excel.Cells(6, 15).Value = "Total"
            
            pgbBarra.Value = 30
            
            o_Sheet.Range("A6:O6").HorizontalAlignment = -4108
            o_Sheet.Range("A6:O6").VerticalAlignment = -4108
            o_Sheet.Range("A6:O6").WrapText = True
            o_Sheet.Range("A7").Select
            o_Excel.ActiveWindow.FreezePanes = True
            o_Sheet.Range("A6:O6").Interior.ColorIndex = 16
            o_Sheet.Range("A:A").ColumnWidth = 20
            o_Sheet.Range("B:B").ColumnWidth = 30
            o_Sheet.Range("C:C").ColumnWidth = 8
            o_Sheet.Range("D:D").ColumnWidth = 30
            o_Sheet.Range("E:E").ColumnWidth = 8
            o_Sheet.Range("F:F").ColumnWidth = 9
            o_Sheet.Range("G:G").ColumnWidth = 10
            o_Sheet.Range("H:H").ColumnWidth = 10
            o_Sheet.Range("I:I").ColumnWidth = 30
            o_Sheet.Range("J:J").ColumnWidth = 30
            o_Sheet.Range("K:K").ColumnWidth = 15
            o_Sheet.Range("L:L").ColumnWidth = 13
            o_Sheet.Range("M:M").ColumnWidth = 15
            o_Sheet.Range("N:N").ColumnWidth = 15
            o_Sheet.Range("O:O").ColumnWidth = 15
             
            o_Sheet.Range(o_Excel.Cells(rsAux.RecordCount + 6, 1), o_Excel.Cells(rsAux.RecordCount + 6, 15)).Borders(4).LineStyle = 1
            
            'info del rs
            o_Sheet.Range("A:O").Font.Size = 9
            o_Sheet.Range("A:O").Font.Name = "Times New Roman" '
            o_Sheet.Range("A:O").Font.Bold = False
            pgbBarra.Value = 70
            'titulos
            o_Sheet.Range("A6:O6").Font.Bold = True
            o_Sheet.Range(o_Excel.Cells(rsAux.RecordCount + 7, 1), o_Excel.Cells(rsAux.RecordCount + 7, 15)).Font.Bold = True
            o_Sheet.Range(o_Excel.Cells(2, 1), o_Excel.Cells(5, 1)).Font.Bold = True
            'centrado, auto ajustar texto, alinear medio
            o_Sheet.Range("J:O").NumberFormat = "$ ###,###,###,##0.00"
            
            'rs,maxRows,maxCols
            o_Sheet.Range("A7").CopyFromRecordset rsAux, , 15
            
            pgbBarra.Value = 100
            freBarra.Visible = False
            'La información ha sido exportada exitosamente
            MsgBox SIHOMsg(1185), vbOKOnly + vbInformation, "Mensaje"
            o_Excel.Visible = True
            
            Set o_Excel = Nothing
    
    Else
            'No existe información con esos parámetros.
            MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje"
        End If
        rsAux.Close
    End If
    
Exit Sub
NotificaError:
    ' -- Cierra la hoja y la aplicación Excel
    freBarra.Visible = False
    If Not o_Libro Is Nothing Then: o_Libro.Close False
    If Not o_Excel Is Nothing Then: o_Excel.Quit
    If Not o_Excel Is Nothing Then Set o_Excel = Nothing
    If Not o_Libro Is Nothing Then Set o_Libro = Nothing
    If Not o_Sheet Is Nothing Then Set o_Sheet = Nothing
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdExportar_Click"))
End Sub

Private Sub cmdImprimir_Click()
    pImprime "I"
End Sub
Private Sub cmdVista_Click()
    pImprime "P"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If KeyAscii = 27 Then
            Unload Me
        End If
    End If
End Sub
Private Sub Form_Load()
Dim rs As New ADODB.Recordset
Dim dtmfecha As Date
    Me.Icon = frmMenuPrincipal.Icon
    Set rs = frsEjecuta_SP("-1", "Sp_Gnselempresascontable")
    If rs.RecordCount <> 0 Then
       pLlenarCboRs cboHospital, rs, 1, 0
       cboHospital.ListIndex = flngLocalizaCbo(cboHospital, Str(vgintClaveEmpresaContable))
    End If
    Me.cboHospital.Enabled = cgstrModulo = "SE" 'sólo se activa en AE
    
    If cgstrModulo = "SE" Then 'AE
        If fblnRevisaPermiso(vglngNumeroLogin, 3070, "C", True) Then
           Me.cboHospital.Enabled = True
        Else
           Me.cboHospital.Enabled = False
        End If
    Else 'CAJA
        If fblnRevisaPermiso(vglngNumeroLogin, 3069, "C", True) Then
            Me.cboHospital.Enabled = True
        Else
            Me.cboHospital.Enabled = False
        End If
    End If
  
    
    
    dtmfecha = fdtmServerFecha
    mskFechaInicio.Mask = ""
    mskFechaInicio.Text = dtmfecha
    mskFechaInicio.Mask = "##/##/####"
    mskFechaFin.Mask = ""
    mskFechaFin.Text = dtmfecha
    mskFechaFin.Mask = "##/##/####"
    Me.chkFacturas.Value = vbChecked
    Me.chkNotas.Value = vbChecked
    Me.chkTickets.Value = vbChecked
End Sub
Private Sub mskFechaFin_GotFocus()
    pSelMkTexto mskFechaFin
End Sub
Private Sub mskFechaInicio_GotFocus()
    pSelMkTexto mskFechaInicio
End Sub
