VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReporteFacturacionIntegrada 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Facturación integrada"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6360
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   6360
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame freBarra 
      Height          =   815
      Left            =   948
      TabIndex        =   14
      Top             =   1985
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
         TabIndex        =   20
         Top             =   225
         Width           =   4365
      End
   End
   Begin VB.Frame Frame6 
      Height          =   710
      Left            =   2390
      TabIndex        =   13
      Top             =   3960
      Width           =   1580
      Begin VB.CommandButton cmdExportar 
         Caption         =   "Exportar a Excel"
         Height          =   495
         Left            =   60
         TabIndex        =   9
         ToolTipText     =   "Exportar a Excel el reporte"
         Top             =   150
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de facturación "
      Height          =   2325
      Left            =   90
      TabIndex        =   12
      Top             =   1605
      Width           =   6165
      Begin VB.CheckBox ChkIncluirPXFechaFinalRango 
         Caption         =   "Incluir pacientes que a la fecha final del rango tengan cargos por facturar"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         ToolTipText     =   "Incluir pacientes que a la fecha final del rango tengan cargos por facturar"
         Top             =   1320
         Value           =   1  'Checked
         Width           =   5775
      End
      Begin VB.CheckBox ChkIncluirPXFeacRango 
         Caption         =   "Incluir pacientes facturados en el rango"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         ToolTipText     =   "Incluir pacientes facturados en el rango"
         Top             =   975
         Value           =   1  'Checked
         Width           =   3135
      End
      Begin VB.CheckBox ChkIncluitPXIngresadosCargo 
         Caption         =   "Incluir pacientes ingresados en el rango"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         ToolTipText     =   "Incluir pacientes ingresados en el rango"
         Top             =   645
         Value           =   1  'Checked
         Width           =   3135
      End
      Begin VB.CheckBox chkDirectas 
         Caption         =   "Facturación directa"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   "Incluir facturas directas"
         Top             =   1650
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox chkVentasPublico 
         Caption         =   "Ventas al público"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         ToolTipText     =   "Incluir ventas al público"
         Top             =   2000
         Value           =   1  'Checked
         Width           =   1530
      End
      Begin VB.CheckBox chkPacientes 
         Caption         =   "Facturación de pacientes"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Incluir facturas de pacientes"
         Top             =   310
         Value           =   1  'Checked
         Width           =   2175
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Rango de fechas "
      Height          =   885
      Left            =   90
      TabIndex        =   11
      Top             =   720
      Width           =   6165
      Begin MSMask.MaskEdBox mskFecIni 
         Height          =   315
         Left            =   840
         TabIndex        =   1
         ToolTipText     =   "Fecha inicial"
         Top             =   360
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mmm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskFecFin 
         Height          =   315
         Left            =   4710
         TabIndex        =   2
         ToolTipText     =   "Fecha final"
         Top             =   360
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
         Caption         =   "Hasta"
         Height          =   195
         Left            =   4080
         TabIndex        =   19
         Top             =   420
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   420
         Width           =   465
      End
   End
   Begin VB.Frame Frame5 
      Height          =   675
      Left            =   90
      TabIndex        =   10
      Top             =   0
      Width           =   6165
      Begin VB.ComboBox cboHospital 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Seleccione la empresa"
         Top             =   240
         Width           =   5175
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Empresa"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   300
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmReporteFacturacionIntegrada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------------
' Reporte de las facturas canceladas en un rango de fechas del departamento
' Fecha de programación: Martes 16 de Abril de 2002
'--------------------------------------------------------------------------------------
'Ultimas modificaciones, especificar:
'Fecha:
'Descripción del cambio:
'--------------------------------------------------------------------------------------
Option Explicit

Dim o_Excel As Object
Dim o_ExcelAbrir As Object
Dim o_Libro As Object
Dim o_Sheet As Object

Dim vlstrx As String
Private vgrptReporte As CRAXDRT.Report
Public vglngNumeroOpcion As Long

Private Sub chkPacientes_Click()
    If chkPacientes.Value Then
        ChkIncluitPXIngresadosCargo.Enabled = True
        ChkIncluirPXFeacRango.Enabled = True
        ChkIncluirPXFechaFinalRango.Enabled = True
        ChkIncluitPXIngresadosCargo.Value = 1
        ChkIncluirPXFeacRango.Value = 1
        ChkIncluirPXFechaFinalRango.Value = 1
    Else
        ChkIncluitPXIngresadosCargo.Enabled = False
        ChkIncluirPXFeacRango.Enabled = False
        ChkIncluirPXFechaFinalRango.Enabled = False
        ChkIncluitPXIngresadosCargo.Value = 0
        ChkIncluirPXFeacRango.Value = 0
        ChkIncluirPXFechaFinalRango.Value = 0
    End If
End Sub

Private Sub cmdExportar_Click()
On Error GoTo NotificaError
    Dim rsAux As New ADODB.Recordset
    Dim rsColumnas As New ADODB.Recordset
    Dim VlColumnas As String
    Dim vlCont As Integer
    Dim vlContAux As Integer
    Dim vlContCero As Integer
    Dim arrSplitStrings() As String
        
    If chkPacientes.Value = 0 And chkDirectas.Value = 0 And chkVentasPublico.Value = 0 Then
        'Seleccione el dato.
        MsgBox SIHOMsg(431), vbExclamation + vbOKOnly, "Mensaje"
        chkPacientes.SetFocus
        Exit Sub
    End If
        
    If fblnFechasValidas() Then
    
        vgstrParametrosSP = str(cboHospital.ItemData(cboHospital.ListIndex)) _
            & "," & fstrFechaSQL(mskFecIni.Text, "00:00:00") _
            & "," & fstrFechaSQL(mskFecFin.Text, "23:59:59") _
            & "," & 0
            
        Set rsColumnas = frsRegresaRs("select FN_PVSelObtenCadenaFormaPago(" & vgstrParametrosSP & ") cadena from dual", adLockOptimistic, adOpenDynamic)
         
         If Trim(rsColumnas!Cadena) <> "" Then
           VlColumnas = Replace(rsColumnas!Cadena, "@", "")
           arrSplitStrings = Split(VlColumnas, ",")
        Else
           arrSplitStrings = Split("", ",")
        End If
     
        Screen.MousePointer = vbHourglass
        lblTextoBarra.Caption = " Obteniendo información, por favor espere..."
        freBarra.Visible = True
        
        DoEvents
        
        pgbBarra.Value = 0
        
        pgbBarra.Value = 25
        
        vgstrParametrosSP = str(cboHospital.ItemData(cboHospital.ListIndex)) _
            & "|" & fstrFechaSQL(mskFecIni.Text, "00:00:00") _
            & "|" & fstrFechaSQL(mskFecFin.Text, "23:59:59") _
            & "|" & chkPacientes.Value _
            & "|" & chkDirectas.Value _
            & "|" & chkVentasPublico.Value _
            & "|" & ChkIncluitPXIngresadosCargo.Value _
            & "|" & ChkIncluirPXFeacRango.Value _
            & "|" & ChkIncluirPXFechaFinalRango.Value
        
        pgbBarra.Value = 50
            
        Set rsAux = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELFACTURACIONINTEGRADA")

        
        pgbBarra.Value = 100
        freBarra.Visible = False
        
        If rsAux.RecordCount > 0 Then
            lblTextoBarra.Caption = " Exportando información, por favor espere..."
            pgbBarra.Value = 0
            freBarra.Visible = True
            
            Set o_Excel = CreateObject("Excel.Application")
            Set o_Libro = o_Excel.Workbooks.Add
            Set o_Sheet = o_Libro.Worksheets(1)
            
            'datos del repote
            o_Excel.cells(2, 1).Value = Trim(cboHospital.Text)
            o_Excel.cells(3, 1).Value = "FACTURACIÓN INTEGRADA"
            o_Excel.cells(4, 1).Value = "Del " & CStr(Format(CDate(mskFecIni), "dd/MMM/yyyy")) & " Al " & CStr(Format(CDate(mskFecFin), "dd/MMM/yyyy"))
            'columnas titulos
            o_Excel.cells(6, 1).Value = "Tipo de facturación"
            o_Excel.cells(6, 2).Value = "Fecha de ingreso / Factura / Venta / Cancelación"
            o_Excel.cells(6, 3).Value = "Cuenta"
            o_Excel.cells(6, 4).Value = "Número de expediente / Cliente"
            o_Excel.cells(6, 5).Value = "Tipo de ingreso"
            o_Excel.cells(6, 6).Value = "Cuarto"
            o_Excel.cells(6, 7).Value = "Departamento de ingreso"
            o_Excel.cells(6, 8).Value = "Fecha de egreso"
            o_Excel.cells(6, 9).Value = "Estado actual de la cuenta"
            
            o_Excel.cells(6, 10).Value = "Estado de la cuenta en el rango de fechas"
            
            o_Excel.cells(6, 11).Value = "Nombre del paciente / Empleado / Médico / Cliente"
            o_Excel.cells(6, 12).Value = "Tipo del paciente"
            o_Excel.cells(6, 13).Value = "Facturado antes del rango de fechas"
            o_Excel.cells(6, 14).Value = "Folios de facturas de lo facturado antes del rango de fechas"
            o_Excel.cells(6, 15).Value = "Facturado en el rango"
            o_Excel.cells(6, 16).Value = "Folios de facturas de lo facturado entre el rango de fechas"
            o_Excel.cells(6, 17).Value = "Importe pendiente de facturar"
            o_Excel.cells(6, 18).Value = "Subtotal"
            o_Excel.cells(6, 19).Value = "IVA"
            o_Excel.cells(6, 20).Value = "Total"
            o_Excel.cells(6, 21).Value = "Pagos"
            
            'encabezados nuevas columnas
            For vlCont = LBound(arrSplitStrings, 1) To UBound(arrSplitStrings, 1)
            o_Excel.cells(6, 23 + vlCont).Value = UCase(Mid(arrSplitStrings(vlCont), 1, 1)) & LCase(Mid(arrSplitStrings(vlCont), 2, Len(arrSplitStrings(vlCont))))
            Next vlCont
            
            'Sumatorias
            o_Excel.cells(rsAux.RecordCount + 7, 12).Formula = "TOTAL"
            o_Excel.cells(rsAux.RecordCount + 7, 13).Formula = "=SUM(M7:M" & (rsAux.RecordCount + 6) & ")"
            o_Excel.cells(rsAux.RecordCount + 7, 15).Formula = "=SUM(O7:O" & (rsAux.RecordCount + 6) & ")"
            o_Excel.cells(rsAux.RecordCount + 7, 17).Formula = "=SUM(Q7:Q" & (rsAux.RecordCount + 6) & ")"
            o_Excel.cells(rsAux.RecordCount + 7, 18).Formula = "=SUM(R7:R" & (rsAux.RecordCount + 6) & ")"
            o_Excel.cells(rsAux.RecordCount + 7, 19).Formula = "=SUM(S7:S" & (rsAux.RecordCount + 6) & ")"
            o_Excel.cells(rsAux.RecordCount + 7, 20).Formula = "=SUM(T7:T" & (rsAux.RecordCount + 6) & ")"
            o_Excel.cells(rsAux.RecordCount + 7, 21).Formula = "=SUM(U7:U" & (rsAux.RecordCount + 6) & ")"
                                    
                                    
            pgbBarra.Value = 30
            
            o_Sheet.range("A6:BX6").HorizontalAlignment = -4108
            o_Sheet.range("A6:BX6").VerticalAlignment = -4108
            o_Sheet.range("A6:BX6").WrapText = True
            o_Sheet.range("A7").Select
            o_Excel.ActiveWindow.FreezePanes = True
            'color de fondo encabezados
            o_Sheet.range("A6:" & LetraColumna(23 + UBound(arrSplitStrings, 1)) & "6").Interior.ColorIndex = 15 '15 48
            'color de fondo totales
            o_Sheet.range(o_Excel.cells(rsAux.RecordCount + 7, 13), o_Excel.cells(rsAux.RecordCount + 7, 23 + UBound(arrSplitStrings, 1))).Interior.ColorIndex = 15
            o_Sheet.range("A:A").Columnwidth = 15
            o_Sheet.range("B:B").Columnwidth = 10
            o_Sheet.range("C:C").Columnwidth = 6
            o_Sheet.range("D:D").Columnwidth = 9
            o_Sheet.range("E:E").Columnwidth = 25
            o_Sheet.range("F:F").Columnwidth = 7
            o_Sheet.range("G:G").Columnwidth = 20
            o_Sheet.range("H:H").Columnwidth = 10
            o_Sheet.range("I:I").Columnwidth = 24
            o_Sheet.range("J:J").Columnwidth = 20
            o_Sheet.range("K:K").Columnwidth = 38
            o_Sheet.range("L:L").Columnwidth = 15
            o_Sheet.range("M:M").Columnwidth = 13
            o_Sheet.range("N:N").Columnwidth = 13
            o_Sheet.range("O:O").Columnwidth = 13
            o_Sheet.range("P:P").Columnwidth = 13
            o_Sheet.range("Q:Q").Columnwidth = 13
            o_Sheet.range("R:R").Columnwidth = 13
            o_Sheet.range("S:S").Columnwidth = 13
            o_Sheet.range("T:T").Columnwidth = 13
            o_Sheet.range("U:U").Columnwidth = 13
            
            o_Sheet.range(o_Excel.cells(rsAux.RecordCount + 6, 1), o_Excel.cells(rsAux.RecordCount + 6, 23 + UBound(arrSplitStrings, 1))).Borders(4).LineStyle = 1
            
            'info del rs
            o_Sheet.range("A:BX").Font.Size = 9
            o_Sheet.range("A:BX").Font.Name = "Times New Roman" '
            o_Sheet.range("A:BX").Font.Bold = False
            pgbBarra.Value = 70
            
            'titulos
            o_Sheet.range("A6:BX6").Font.Bold = True
            o_Sheet.range(o_Excel.cells(rsAux.RecordCount + 7, 1), o_Excel.cells(rsAux.RecordCount + 7, 23 + UBound(arrSplitStrings, 1))).Font.Bold = True
            o_Sheet.range(o_Excel.cells(2, 1), o_Excel.cells(5, 1)).Font.Bold = True
            'centrado, auto ajustar texto, alinear medio
                                    
            'rs,maxRows,maxCols
            o_Sheet.range("A7").CopyFromRecordset rsAux, , 60
            
            o_Sheet.range("M:M").NumberFormat = "$ ###,###,###,##0.00"
            o_Sheet.range("O:O").NumberFormat = "$ ###,###,###,##0.00"
            o_Sheet.range("Q:Q").NumberFormat = "$ ###,###,###,##0.00"
            o_Sheet.range("R:R").NumberFormat = "$ ###,###,###,##0.00"
            o_Sheet.range("S:S").NumberFormat = "$ ###,###,###,##0.00"
            o_Sheet.range("T:T").NumberFormat = "$ ###,###,###,##0.00"
            o_Sheet.range("U:U").NumberFormat = "$ ###,###,###,##0.00"
            
            
            o_Sheet.range("B:B").NumberFormat = "dd/mmm/yyyy"
            o_Sheet.range("H:H").NumberFormat = "dd/mmm/yyyy"
            
                        
            'Sumatorias - Ancho de la columna - Formato numeros
            For vlCont = LBound(arrSplitStrings, 1) To UBound(arrSplitStrings, 1)
            
             o_Excel.cells(rsAux.RecordCount + 7, 23 + vlCont).Formula = "=SUM(" & LetraColumna(23 + vlCont) & "7:" & LetraColumna(23 + vlCont) & (rsAux.RecordCount + 6) & ")"
                          
             o_Sheet.range(LetraColumna(23 + vlCont) & ":" & LetraColumna(23 + vlCont)).Columnwidth = 15
                   
           
            'Rellenar de 0 las columnas vacias
            If rsAux.RecordCount > 1 Then
             o_Sheet.range(LetraColumna(23 + vlCont) & "7:" & LetraColumna(23 + vlCont) & (rsAux.RecordCount + 6)).Replace What:="", Replacement:="0"
             End If
           
                        
            o_Sheet.range(LetraColumna(23 + vlCont) & ":" & LetraColumna(23 + vlCont)).NumberFormat = "$ ###,###,###,##0.00"
            Next vlCont
                                                                                      
            'ELIMINA cOLUMNA Datos movpaciente
             o_Excel.Columns("V").Delete
             
             'eliminar sumatorias en 0
            vlContAux = LBound(arrSplitStrings, 1)
            For vlCont = LBound(arrSplitStrings, 1) To UBound(arrSplitStrings, 1)
                    If o_Excel.cells(rsAux.RecordCount + 7, 22 + vlContAux).Value = 0 Then
                         o_Excel.Columns(LetraColumna(22 + vlContAux)).Delete
                    Else
                        vlContAux = vlContAux + 1
                    End If
            Next vlCont
            
            
                        
            pgbBarra.Value = 100
            freBarra.Visible = False
            'La información ha sido exportada exitosamente
            MsgBox SIHOMsg(1185), vbOKOnly + vbInformation, "Mensaje"
            o_Excel.Visible = True
            
            Set o_Excel = Nothing
            
            Screen.MousePointer = vbDefault
        Else
            'No existe información con esos parámetros
            MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje"
            Screen.MousePointer = vbDefault
        End If
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


Function LetraColumna(numero As Integer) As String
    
    If (numero <= 26) Then
         LetraColumna = (Chr(numero + 64))
    ElseIf (numero <= 52) Then
        LetraColumna = "A" & (Chr(numero + 64 - 26))
    ElseIf (numero <= 78) Then
        LetraColumna = "B" & (Chr(numero + 64 - 52))
    ElseIf (numero <= 104) Then
        LetraColumna = "C" & (Chr(numero + 64 - 78))
    ElseIf (numero <= 130) Then
        LetraColumna = "D" & (Chr(numero + 64 - 104))
    End If
        
End Function



Private Sub Form_KeyPress(KeyAscii As Integer)
        On Error GoTo NotificaError
    
    If KeyAscii = 27 Then
        Unload Me
    Else
        If KeyAscii = 13 Then
            SendKeys vbTab
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyPress"))
End Sub

Private Sub Form_Load()
    On Error GoTo NotificaError
    
    Dim rs As New ADODB.Recordset
    Dim lngNumOpcion As Long
    Dim dtmfecha As Date
    
    Me.Icon = frmMenuPrincipal.Icon

    Select Case cgstrModulo
    Case "PV"
         lngNumOpcion = 343
    Case "SE"
         lngNumOpcion = 1535
    End Select
    
    pCargaHospital lngNumOpcion

    pInstanciaReporte vgrptReporte, "rptFacturasCanceladas.rpt"
    
    dtmfecha = fdtmServerFecha
   
    mskFecIni.Mask = ""
    mskFecIni.Text = dtmfecha
    mskFecIni.Mask = "##/##/####"
    
    mskFecFin.Mask = ""
    mskFecFin.Text = dtmfecha
    mskFecFin.Mask = "##/##/####"
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
End Sub

Private Sub mskFecFin_GotFocus()
    On Error GoTo NotificaError
    
    pSelMkTexto mskFecFin

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFecFin_GotFocus"))
End Sub

Private Sub mskFecFin_LostFocus()
    On Error GoTo NotificaError

    If Trim(mskFecFin.ClipText) = "" Then
        mskFecFin.Mask = ""
        mskFecFin.Text = fdtmServerFecha
        mskFecFin.Mask = "##/##/####"
    Else
        If Not IsDate(mskFecFin.Text) Then
            mskFecFin.Mask = ""
            mskFecFin.Text = fdtmServerFecha
            mskFecFin.Mask = "##/##/####"
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFecFin_LostFocus"))
End Sub

Private Sub mskFecIni_GotFocus()
    On Error GoTo NotificaError
    
    pSelMkTexto mskFecIni

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFecIni_GotFocus"))
End Sub

Private Sub mskFecIni_LostFocus()
    On Error GoTo NotificaError
    
    If Trim(mskFecIni.ClipText) = "" Then
        mskFecIni.Mask = ""
        mskFecIni.Text = fdtmServerFecha
        mskFecIni.Mask = "##/##/####"
    Else
        If Not IsDate(mskFecIni.Text) Then
            mskFecIni.Mask = ""
            mskFecIni.Text = fdtmServerFecha
            mskFecIni.Mask = "##/##/####"
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFecIni_LostFocus"))
End Sub

Private Sub pCargaHospital(lngNumOpcion As Long)
    On Error GoTo NotificaError
    
    Dim rs As New ADODB.Recordset
    
    Set rs = frsEjecuta_SP("-1", "Sp_Gnselempresascontable")
    If rs.RecordCount <> 0 Then
        pLlenarCboRs cboHospital, rs, 1, 0
        cboHospital.ListIndex = flngLocalizaCbo(cboHospital, str(vgintClaveEmpresaContable))
    End If
    
    cboHospital.Enabled = fblnRevisaPermiso(vglngNumeroLogin, lngNumOpcion, "C")
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCargaHospital"))
    Unload Me
End Sub

Private Function fblnFechasValidas() As Boolean
    On Error GoTo NotificaError
        
        fblnFechasValidas = True
        If Not IsDate(mskFecIni) Then
            MsgBox SIHOMsg(29), vbOKOnly + vbInformation, "Mensaje"
            fblnFechasValidas = False
            mskFecIni.SetFocus
        ElseIf Not IsDate(mskFecFin) Then
            MsgBox SIHOMsg(29), vbCritical, "Mensaje"
            fblnFechasValidas = False
            mskFecFin.SetFocus
        'fecha final valida
        ElseIf CDate(mskFecIni) > CDate(mskFecFin) Then
            MsgBox SIHOMsg(64), vbCritical, "Mensaje"
            fblnFechasValidas = False
            mskFecIni.SetFocus
        End If
        
    Exit Function
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnFechasValidas"))
End Function




