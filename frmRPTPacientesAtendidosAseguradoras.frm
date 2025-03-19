VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmPacientesAtendidosAseguradoras 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pacientes atendidos de aseguradoras"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7260
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Height          =   675
      Index           =   1
      Left            =   60
      TabIndex        =   19
      Top             =   50
      Width           =   7125
      Begin VB.ComboBox cboHospital 
         Height          =   315
         Left            =   915
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Empresa contable"
         Top             =   240
         Width           =   6015
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Empresa"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   300
         Width           =   615
      End
   End
   Begin VB.Frame freBarra 
      Height          =   815
      Left            =   1398
      TabIndex        =   15
      Top             =   1528
      Visible         =   0   'False
      Width           =   4465
      Begin MSComctlLib.ProgressBar pgbBarra 
         Height          =   300
         Left            =   45
         TabIndex        =   16
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
         TabIndex        =   18
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
         TabIndex        =   17
         Top             =   225
         Width           =   4365
      End
   End
   Begin VB.Frame Frame5 
      Height          =   705
      Index           =   0
      Left            =   2858
      TabIndex        =   14
      Top             =   3000
      Width           =   1545
      Begin VB.CommandButton cmdExportar 
         Caption         =   "Exportar a Excel"
         Height          =   495
         Left            =   50
         TabIndex        =   7
         ToolTipText     =   "Exportar a Excel"
         Top             =   150
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Rango de fechas de ingreso"
      Height          =   750
      Left            =   1455
      TabIndex        =   10
      Top             =   2160
      Width           =   4350
      Begin MSMask.MaskEdBox mskInicio 
         Height          =   315
         Left            =   720
         TabIndex        =   5
         ToolTipText     =   "Fecha de inicio"
         Top             =   285
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mmm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskFin 
         Height          =   315
         Left            =   2880
         TabIndex        =   6
         ToolTipText     =   "Fecha de fin"
         Top             =   285
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
         Left            =   2280
         TabIndex        =   12
         Top             =   360
         Width           =   420
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   345
         Width           =   465
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1155
      Left            =   60
      TabIndex        =   8
      Top             =   840
      Width           =   7125
      Begin VB.OptionButton optInternoExterno 
         Caption         =   "Interno"
         Height          =   195
         Index           =   0
         Left            =   2325
         TabIndex        =   3
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton optInternoExterno 
         Caption         =   "Externo"
         Height          =   195
         Index           =   1
         Left            =   3270
         TabIndex        =   4
         Top             =   720
         Width           =   960
      End
      Begin VB.OptionButton optInternoExterno 
         Caption         =   "Todos"
         Height          =   195
         Index           =   2
         Left            =   1425
         TabIndex        =   2
         Top             =   720
         Width           =   945
      End
      Begin VB.ComboBox cboEmpresa 
         Height          =   315
         Left            =   1425
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Empresa"
         Top             =   240
         Width           =   5505
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de paciente"
         Height          =   270
         Left            =   120
         TabIndex        =   13
         Top             =   705
         Width           =   1200
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Convenio"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   300
         Width           =   675
      End
   End
End
Attribute VB_Name = "frmPacientesAtendidosAseguradoras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------
'| Nombre del Proyecto      : Caja                                                   -
'| Nombre del Formulario    : frmPacientesAtendidosAseguradoras
'-------------------------------------------------------------------------------------
'| Objetivo: Sacar un reporte de pacientes externos atendidos por departamento
'-------------------------------------------------------------------------------------
'| Análisis y Diseño        : Jesús Valles
'| Autor                    : Jesus Valles
'| Fecha de Creación        : 09/Dic/2019
'| Modificó                 : Nombre(s)
'| Fecha Terminación        : hoy
'| Fecha última modificación: 09/Dic/2019
'-------------------------------------------------------------------------------------

Dim dtmfecha As Date

Dim o_Excel As Object
Dim o_ExcelAbrir As Object
Dim o_Libro As Object
Dim o_Sheet As Object

Option Explicit
Private vgrptReporte As CRAXDRT.Report

Private Sub cboDepartamento_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub cboEmpresa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub cboHospital_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub chkDetallado_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub ChkFacturados_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub ChkSinFacturar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub
 
Private Sub cmdExportar_Click()
    Dim rsAux As New ADODB.Recordset
    Dim rsAux2 As New ADODB.Recordset
    Dim vlstrInternoExterno As String
    Dim vlstrParametros As String
    Dim intRenglones As Integer
    Dim intColumnas As Integer
    Dim vlintAumentosBarra As Double
    Dim vlintseq As Integer
    
    If CDate(mskInicio.Text) > CDate(mskFin.Text) Then
        '¡Rango de fechas no válido!
        MsgBox SIHOMsg(64), vbExclamation, "Mensaje"
        mskInicio.SetFocus
        
        Exit Sub
    End If
    
    Me.MousePointer = 11
    
    If optInternoExterno(0) Then
        vlstrInternoExterno = "I"
    ElseIf optInternoExterno(1) Then
        vlstrInternoExterno = "E"
    ElseIf optInternoExterno(2) Then
        vlstrInternoExterno = "T"
    End If
        
    Set rsAux = frsEjecuta_SP(fstrFechaSQL(mskInicio.Text, "00:00:00") & "|" & fstrFechaSQL(mskFin.Text, "23:59:59") & "|" & cboEmpresa.ItemData(cboEmpresa.ListIndex) & "|" & CStr(cboHospital.ItemData(cboHospital.ListIndex)) & "|" & vlstrInternoExterno & "|" & 0, "SP_PVRPTPACATENDIDOSASEG", , , , True)

    If rsAux.EOF Then
        'No existe información con esos parámetros
        MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje"
        
        Me.MousePointer = 0
        
        Exit Sub
    Else
        If rsAux.RecordCount > 0 Then
            pgbBarra.Value = 0
            freBarra.Visible = True
            
            Set o_Excel = CreateObject("Excel.Application")
            Set o_Libro = o_Excel.Workbooks.Add
            Set o_Sheet = o_Libro.Worksheets(1)
            
            o_Excel.ActiveWorkbook.ActiveSheet.Name = "PACIENTES ATENDIDOS"
            
            'datos del repote
            o_Excel.Cells(2, 1).Value = Trim(vgstrNombreHospitalCH)
            o_Excel.Cells(3, 1).Value = "PACIENTES ATENDIDOS DE " & Trim(cboEmpresa.Text)
            o_Excel.Cells(4, 1).Value = "Del " & CStr(Format(CDate(mskInicio), "dd/MMM/yyyy")) & " Al " & CStr(Format(CDate(mskFin), "dd/MMM/yyyy"))
            'columnas titulos
            o_Excel.Cells(6, 1).Value = "Fecha de ingreso"
            o_Excel.Cells(6, 2).Value = "Nombre del paciente"
            o_Excel.Cells(6, 3).Value = "Número de expediente"
            o_Excel.Cells(6, 4).Value = "Número de cuenta"
            o_Excel.Cells(6, 5).Value = "Número de afiliación"
            o_Excel.Cells(6, 6).Value = "Diagnóstico"
            o_Excel.Cells(6, 7).Value = "Médico tratante"
            o_Excel.Cells(6, 8).Value = "Especialidad"
            
'            o_Excel.Cells(6, 9).Value = "Cuenta/ Clave"
'            o_Excel.Cells(6, 10).Value = "Tipo"
'            o_Excel.Cells(6, 11).Value = "Paciente/ Cliente"
'            o_Excel.Cells(6, 12).Value = "Empresa"
'            o_Excel.Cells(6, 13).Value = "Médico tratante"
'            o_Excel.Cells(6, 14).Value = "Importe gravado factura"
'            o_Excel.Cells(6, 15).Value = "Importe no gravado factura"

'            o_Excel.Cells(6, 16).Value = "Importe gravado nota"
'            o_Excel.Cells(6, 17).Value = "Importe no gravado nota"

'            o_Excel.Cells(6, 18).Value = "Descuento gravado"
'            o_Excel.Cells(6, 19).Value = "Descuento no gravado"
'            o_Excel.Cells(6, 20).Value = "Subtotal gravado"
'            o_Excel.Cells(6, 21).Value = "Subtotal no gravado"
'            o_Excel.Cells(6, 22).Value = "IVA"
'            o_Excel.Cells(6, 23).Value = "Total"
            
            'sumatorias
'            o_Excel.Cells(rsAux.RecordCount + 7, 13).Formula = "TOTAL"
'            o_Excel.Cells(rsAux.RecordCount + 7, 14).Formula = "=SUM(N7:N" & (rsAux.RecordCount + 6) & ")"
'            o_Excel.Cells(rsAux.RecordCount + 7, 15).Formula = "=SUM(O7:O" & (rsAux.RecordCount + 6) & ")"
'
'            o_Excel.Cells(rsAux.RecordCount + 7, 16).Formula = "=SUM(P7:P" & (rsAux.RecordCount + 6) & ")"
'            o_Excel.Cells(rsAux.RecordCount + 7, 17).Formula = "=SUM(Q7:Q" & (rsAux.RecordCount + 6) & ")"
'
'            o_Excel.Cells(rsAux.RecordCount + 7, 18).Formula = "=SUM(R7:R" & (rsAux.RecordCount + 6) & ")"
'            o_Excel.Cells(rsAux.RecordCount + 7, 19).Formula = "=SUM(S7:S" & (rsAux.RecordCount + 6) & ")"
'            o_Excel.Cells(rsAux.RecordCount + 7, 20).Formula = "=SUM(T7:T" & (rsAux.RecordCount + 6) & ")"
'            o_Excel.Cells(rsAux.RecordCount + 7, 21).Formula = "=SUM(U7:U" & (rsAux.RecordCount + 6) & ")"
'            o_Excel.Cells(rsAux.RecordCount + 7, 22).Formula = "=SUM(V7:V" & (rsAux.RecordCount + 6) & ")"
'            o_Excel.Cells(rsAux.RecordCount + 7, 23).Formula = "=SUM(W7:W" & (rsAux.RecordCount + 6) & ")"
            
            pgbBarra.Value = 15
            
            o_Sheet.Range("A6:X6").HorizontalAlignment = -4108
            o_Sheet.Range("A6:X6").VerticalAlignment = -4108
            o_Sheet.Range("A6:X6").WrapText = True
            o_Sheet.Range("A7").Select
            o_Excel.ActiveWindow.FreezePanes = True
            
            o_Sheet.Range("A:A").ColumnWidth = 10
            o_Sheet.Range("B:B").ColumnWidth = 30
            o_Sheet.Range("C:C").ColumnWidth = 9
            o_Sheet.Range("D:D").ColumnWidth = 7
            o_Sheet.Range("E:E").ColumnWidth = 9
            o_Sheet.Range("F:F").ColumnWidth = 30
            o_Sheet.Range("G:G").ColumnWidth = 30
            o_Sheet.Range("H:H").ColumnWidth = 16
            o_Sheet.Range("I:Z").ColumnWidth = 18
                                    
            'info del rs
            o_Sheet.Range("A:X").Font.Size = 9
            o_Sheet.Range("A:X").Font.Name = "Times New Roman" '
            o_Sheet.Range("A:X").Font.Bold = False
            
            pgbBarra.Value = 20
            
            'titulos
            o_Sheet.Range("A6:X6").Font.Bold = True
            o_Sheet.Range(o_Excel.Cells(rsAux.RecordCount + 7, 1), o_Excel.Cells(rsAux.RecordCount + 7, 23)).Font.Bold = True
            o_Sheet.Range(o_Excel.Cells(2, 1), o_Excel.Cells(5, 1)).Font.Bold = True
                        
            'Columnas de los conecptos de facturación
            Set rsAux2 = frsEjecuta_SP(fstrFechaSQL(mskInicio.Text, "00:00:00") & "|" & fstrFechaSQL(mskFin.Text, "23:59:59") & "|" & cboEmpresa.ItemData(cboEmpresa.ListIndex) & "|" & CStr(cboHospital.ItemData(cboHospital.ListIndex)) & "|" & vlstrInternoExterno & "|" & 1, "SP_PVRPTPACATENDIDOSASEG", , , , True)
            intColumnas = 9
            If Not rsAux2.EOF Then
                If rsAux2.RecordCount > 0 Then
                    Do While Not rsAux2.EOF
                        o_Excel.Cells(5, intColumnas).Value = Trim(rsAux2(9).Value) 'Clave del concepto
                        o_Excel.Cells(6, intColumnas).Value = Trim(rsAux2(10).Value) 'Descripción del concepto
                        intColumnas = intColumnas + 1
                        rsAux2.MoveNext
                    Loop
                End If
            End If
            
            o_Excel.Cells(6, intColumnas).Value = "SUBTOTAL"
            o_Excel.Cells(6, intColumnas + 1).Value = "IVA"
            o_Excel.Cells(6, intColumnas + 2).Value = "TOTAL"
            
            'Agrega información
            intRenglones = 7
            vlintAumentosBarra = 30 / rsAux.RecordCount
            Do While Not rsAux.EOF
                    If Trim(o_Excel.Cells(intRenglones - 1, 3).Value) = Trim(rsAux(2).Value) And Trim(o_Excel.Cells(intRenglones - 1, 4).Value) = Trim(rsAux(3).Value) Then
                        intRenglones = intRenglones - 1
                    End If
                    
                    o_Excel.Cells(intRenglones, 1).Value = rsAux(0).Value       'Fecha de ingreso
                    o_Excel.Cells(intRenglones, 2).Value = Trim(rsAux(1).Value) 'Nombre del paciente
                    o_Excel.Cells(intRenglones, 3).Value = Trim(rsAux(2).Value) 'Núm. expediente
                    o_Excel.Cells(intRenglones, 4).Value = Trim(rsAux(3).Value) 'Núm. cuenta
                    o_Excel.Cells(intRenglones, 5).Value = Trim(rsAux(4).Value) 'Núm. afiliación
                    o_Excel.Cells(intRenglones, 6).Value = IIf(Trim(rsAux(5).Value) <> "", Trim(rsAux(5).Value), Trim(rsAux(6).Value)) 'Diagnostico
                    o_Excel.Cells(intRenglones, 7).Value = Trim(rsAux(7).Value) 'Médico tratante
                    o_Excel.Cells(intRenglones, 8).Value = Trim(rsAux(8).Value) 'Especialidad
                    
                    For vlintseq = 9 To intColumnas + 2
                        If vlintseq <= intColumnas - 1 Then
                            If Trim(o_Excel.Cells(5, vlintseq).Value) = Trim(rsAux(9).Value) Then
                                o_Excel.Cells(intRenglones, vlintseq).Value = Trim(rsAux(11).Value - rsAux(12).Value) 'Importe
                            Else
                                If Trim(o_Excel.Cells(intRenglones, vlintseq).Value) = "" Then
                                    o_Excel.Cells(intRenglones, vlintseq).Value = "0"
                                End If
                            End If
                        Else
                            If vlintseq = intColumnas Then
                                'Columna del subtotal
                                o_Excel.Cells(intRenglones, vlintseq).Value = CDbl(o_Excel.Cells(intRenglones, vlintseq).Value) + (Trim(rsAux(11).Value - rsAux(12).Value))
                            End If
                            If vlintseq = intColumnas + 1 Then
                                'Columna del IVA
                                o_Excel.Cells(intRenglones, vlintseq).Value = CDbl(o_Excel.Cells(intRenglones, vlintseq).Value) + (Trim(rsAux(13).Value))
                            End If
                            If vlintseq = intColumnas + 2 Then
                                'Columna del TOTAL
                                o_Excel.Cells(intRenglones, vlintseq).Value = CDbl(o_Excel.Cells(intRenglones, vlintseq).Value) + (Trim(rsAux(11).Value - rsAux(12).Value + rsAux(13).Value))
                            End If
                        End If
                    Next vlintseq
                    
                    intRenglones = intRenglones + 1

                rsAux.MoveNext
                pgbBarra.Value = pgbBarra.Value + vlintAumentosBarra
            Loop
            
            o_Sheet.Range(o_Excel.Cells(6, 1), o_Excel.Cells(6, intColumnas + 2)).Interior.ColorIndex = 15
            
            o_Sheet.Range(o_Excel.Cells(intRenglones, 9), o_Excel.Cells(intRenglones, intColumnas + 2)).Interior.ColorIndex = 15
            o_Sheet.Range(o_Excel.Cells(intRenglones - 1, 9), o_Excel.Cells(intRenglones - 1, intColumnas + 2)).Borders(4).LineStyle = 1
            
            o_Sheet.Range("5:5").Select
            o_Excel.Selection.ClearContents
                                   
            'centrado, auto ajustar texto, alinear medio
            o_Sheet.Range("I:Z").NumberFormat = "$ ###,###,###,##0.00"
            
            For vlintseq = 9 To intColumnas + 2
                o_Excel.Cells(intRenglones, vlintseq).FormulaR1C1 = "=SUM(R[-" & intRenglones - 7 & "]C:R[-1]C)"
            Next vlintseq
            
            o_Sheet.Range("I" & intRenglones & ":X" & intRenglones).Font.Bold = True
            
            o_Sheet.Range("A7:A7").Select
            
            pgbBarra.Value = 100
            freBarra.Visible = False
            'La información ha sido exportada exitosamente
            MsgBox SIHOMsg(1185), vbOKOnly + vbInformation, "Mensaje"
            o_Excel.Visible = True
            
            Set o_Excel = Nothing
        Else
            'No existe información con esos parámetros
            MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje"
        End If
    End If
    rsAux.Close
    
    Me.MousePointer = 0
     
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

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    Dim vlstrSentencia As String
    Dim rs As New ADODB.Recordset
    Dim vlintTipoArticulo As Integer
    Dim lngNumOpcion As Long
    
    Me.Icon = frmMenuPrincipal.Icon
    
    Select Case cgstrModulo
    Case "PV"
         lngNumOpcion = 4107
    Case "SE"
         lngNumOpcion = 4108
    End Select
    
    pCargaHospital lngNumOpcion
    
    'Empresas
    vlstrSentencia = "select intCveEmpresa, ccEmpresa.vchDescripcion from ccEmpresa inner join cctipoconvenio on cctipoconvenio.tnycvetipoconvenio = ccEmpresa.tnycvetipoconvenio and cctipoconvenio.bitaseguradora = 1"
    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    pLlenarCboRs cboEmpresa, rs, 0, 1
    rs.Close
    
    cboEmpresa.ListIndex = 0
    
    'Fechas
    dtmfecha = fdtmServerFecha
    mskInicio.Text = dtmfecha
    mskFin.Text = dtmfecha
    
    optInternoExterno(2) = True
End Sub

Private Sub mskFin_GotFocus()
    pSelMkTexto mskFin
End Sub

Private Sub mskFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub mskFin_LostFocus()
    If Not IsDate(mskFin.Text) Then
           MsgBox SIHOMsg(29), vbExclamation, "Mensaje"
           mskFin.Text = dtmfecha
           Me.mskFin.SetFocus
    Else
        If Year(CDate(mskFin.Text)) < 1900 Then
            '¡Fecha no válida!
            MsgBox SIHOMsg(254), vbExclamation, "Mensaje"
            mskFin.Text = dtmfecha
            mskFin.SetFocus
        Else
            If CDate(mskFin.Text) > fdtmServerFecha Then
                '¡La fecha debe ser menor o igual a la del sistema!
                MsgBox SIHOMsg(40), vbExclamation, "Mensaje"
                mskFin.Text = dtmfecha
                mskFin.SetFocus
            End If
        End If
    End If
End Sub

Private Sub mskInicio_GotFocus()
    pSelMkTexto mskInicio
End Sub

Private Sub mskInicio_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then pEnfocaMkTexto mskFin
End Sub

Private Sub mskInicio_LostFocus()
    If Not IsDate(mskInicio.Text) Then
           MsgBox SIHOMsg(29), vbExclamation, "Mensaje"
           mskInicio.Text = dtmfecha
           Me.mskInicio.SetFocus
    Else
        If Year(CDate(mskInicio.Text)) < 1900 Then
            '¡Fecha no válida!
            MsgBox SIHOMsg(254), vbExclamation, "Mensaje"
            mskInicio.Text = dtmfecha
            mskInicio.SetFocus
        Else
            If CDate(mskInicio.Text) > fdtmServerFecha Then
                '¡La fecha debe ser menor o igual a la del sistema!
                MsgBox SIHOMsg(40), vbExclamation, "Mensaje"
                mskInicio.Text = dtmfecha
                mskInicio.SetFocus
            End If
        End If
    End If
End Sub

Private Sub optAgrupado_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub optInternoExterno_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub optOrdenFactura_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub optOrdenFecha_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub optOrdenNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub txtNumPaciente_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then KeyAscii = 7
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
