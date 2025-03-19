VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRptFacturacionCronologica 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cronológico de facturas y notas"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9270
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   9270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame freBarra 
      Height          =   815
      Left            =   2400
      TabIndex        =   30
      Top             =   1500
      Visible         =   0   'False
      Width           =   4465
      Begin MSComctlLib.ProgressBar pgbBarra 
         Height          =   300
         Left            =   45
         TabIndex        =   31
         Top             =   480
         Width           =   4380
         _ExtentX        =   7726
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   1
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
         TabIndex        =   33
         Top             =   225
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
         TabIndex        =   32
         Top             =   180
         Width           =   4370
      End
   End
   Begin VB.Frame Frame5 
      Height          =   705
      Left            =   3855
      TabIndex        =   26
      Top             =   3720
      Width           =   1545
      Begin VB.CommandButton cmdExportar 
         Caption         =   "Exportar a Excel"
         Height          =   495
         Left            =   50
         TabIndex        =   14
         ToolTipText     =   "Exportar a Excel"
         Top             =   150
         Width           =   1455
      End
   End
   Begin VB.Frame FrmBotonera 
      Height          =   735
      Left            =   2977
      TabIndex        =   17
      Top             =   7800
      Width           =   1140
      Begin VB.CommandButton cmdVista 
         Height          =   495
         Left            =   75
         Picture         =   "frmRptFacturacionCronologica.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Vista previa"
         Top             =   165
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdImprimir 
         Height          =   495
         Left            =   570
         Picture         =   "frmRptFacturacionCronologica.frx":01A2
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Imprimir"
         Top             =   165
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3495
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   9015
      Begin VB.ComboBox cboTipoPaciente 
         Height          =   315
         ItemData        =   "frmRptFacturacionCronologica.frx":0344
         Left            =   1680
         List            =   "frmRptFacturacionCronologica.frx":034B
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   11
         ToolTipText     =   "Selección del tipo de paciente"
         Top             =   2040
         Width           =   7080
      End
      Begin VB.ComboBox cboTipoConvenio 
         Height          =   315
         ItemData        =   "frmRptFacturacionCronologica.frx":0358
         Left            =   1680
         List            =   "frmRptFacturacionCronologica.frx":035F
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   12
         ToolTipText     =   "Selección del tipo de convenio"
         Top             =   2400
         Width           =   7080
      End
      Begin VB.ComboBox cboEmpresa 
         Height          =   315
         ItemData        =   "frmRptFacturacionCronologica.frx":036C
         Left            =   1680
         List            =   "frmRptFacturacionCronologica.frx":0373
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   13
         ToolTipText     =   "Selección de la empresa"
         Top             =   2760
         Width           =   7080
      End
      Begin VB.CheckBox chkFacVentapublico 
         Caption         =   "Facturas de ventas al público"
         Height          =   255
         Left            =   6240
         TabIndex        =   8
         ToolTipText     =   "Mostrar facturas de ventas al público"
         Top             =   1425
         Width           =   2415
      End
      Begin VB.CheckBox chkFacPacientes 
         Caption         =   "Facturas a pacientes"
         Height          =   375
         Left            =   1680
         TabIndex        =   6
         ToolTipText     =   "Mostrar facturas a pacientes"
         Top             =   1365
         Width           =   1815
      End
      Begin VB.CheckBox chkFacDirectas 
         Caption         =   "Facturas directas"
         Height          =   255
         Left            =   4080
         TabIndex        =   7
         ToolTipText     =   "Mostrar facturas directas a clientes"
         Top             =   1425
         Width           =   1575
      End
      Begin VB.CheckBox chkNotasCredito 
         Caption         =   "Notas de crédito"
         Height          =   255
         Left            =   1680
         TabIndex        =   9
         ToolTipText     =   "Mostrar notas de crédito"
         Top             =   1690
         Width           =   1575
      End
      Begin VB.CheckBox chkNotasCargo 
         Caption         =   "Notas de cargo"
         Height          =   255
         Left            =   4080
         TabIndex        =   10
         ToolTipText     =   "Mostrar notas de cargo"
         Top             =   1690
         Width           =   1455
      End
      Begin VB.ComboBox cboDepartamento 
         Height          =   315
         ItemData        =   "frmRptFacturacionCronologica.frx":0380
         Left            =   1680
         List            =   "frmRptFacturacionCronologica.frx":0387
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Selección del departamento"
         Top             =   960
         Width           =   7080
      End
      Begin VB.Frame Frame2 
         Height          =   465
         Left            =   1680
         TabIndex        =   15
         Top             =   120
         Width           =   3200
         Begin VB.OptionButton optFormaPago 
            Caption         =   "Todos"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   0
            ToolTipText     =   "Todas las formas de pago"
            Top             =   175
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton optFormaPago 
            Caption         =   "Crédito"
            Height          =   195
            Index           =   1
            Left            =   1080
            TabIndex        =   1
            ToolTipText     =   "Forma de pago crédito"
            Top             =   175
            Width           =   915
         End
         Begin VB.OptionButton optFormaPago 
            Caption         =   "Contado"
            Height          =   195
            Index           =   2
            Left            =   2160
            TabIndex        =   2
            ToolTipText     =   "Forma de pago contado"
            Top             =   175
            Width           =   930
         End
      End
      Begin MSMask.MaskEdBox mskFechaInicio 
         Height          =   315
         Left            =   2280
         TabIndex        =   3
         ToolTipText     =   "Fecha inicial"
         Top             =   600
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
         Left            =   4320
         TabIndex        =   4
         ToolTipText     =   "Fecha final"
         Top             =   600
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mmm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de paciente"
         Height          =   195
         Left            =   240
         TabIndex        =   29
         Top             =   2100
         Width           =   1200
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de convenio"
         Height          =   195
         Left            =   240
         TabIndex        =   28
         Top             =   2460
         Width           =   1245
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Empresa"
         Height          =   195
         Left            =   240
         TabIndex        =   27
         Top             =   2820
         Width           =   615
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo documento"
         Height          =   195
         Left            =   240
         TabIndex        =   25
         Top             =   1455
         Width           =   1155
      End
      Begin VB.Label lblDepartamento 
         AutoSize        =   -1  'True
         Caption         =   "Departamento"
         Height          =   195
         Left            =   240
         TabIndex        =   24
         Top             =   1020
         Width           =   1005
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Left            =   1680
         TabIndex        =   23
         Top             =   660
         Width           =   465
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         Height          =   255
         Left            =   3750
         TabIndex        =   22
         Top             =   630
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Rango de fechas"
         Height          =   195
         Left            =   240
         TabIndex        =   21
         Top             =   660
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Forma de pago"
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   255
         Width           =   1065
      End
   End
End
Attribute VB_Name = "frmRptFacturacionCronologica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim o_Excel As Object
Dim o_ExcelAbrir As Object
Dim o_Libro As Object
Dim o_Sheet As Object
Dim lstrTipoPac As String 'Indica el tipo de paciente seleccionado CO, EM, ME, PA

Private Sub cboTipoConvenio_Click()
    pLlenaComboEmpresa IIf(cboTipoConvenio.ListIndex = 0, -1, cboTipoConvenio.ItemData(cboTipoConvenio.ListIndex))
End Sub

Private Sub cboTipoPaciente_Click()
On Error GoTo NotificaError
    Dim rsAux As New ADODB.Recordset
    
    lstrTipoPac = ""
    If cboTipoPaciente.ListIndex = 0 Then
        pLlenaComboConvenio True
    Else
        Set rsAux = frsEjecuta_SP(cboTipoPaciente.ItemData(cboTipoPaciente.ListIndex), "Sp_Adseltipopaciente")
        If rsAux.RecordCount > 0 Then
            lstrTipoPac = rsAux!chrTipo
            pLlenaComboConvenio IIf(lstrTipoPac = "CO", True, False)
        Else
            pLlenaComboConvenio True
        End If
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboTipoPaciente_Click"))
End Sub

Private Sub cmdExportar_Click()
On Error GoTo NotificaError
    Dim rsAux As New ADODB.Recordset
        
    If fblnFechasValidas() Then
        'IN_INTCVEDEPTO      IN INTEGER, -- Clave del departamento, -1 = todos
        'IN_INTCVEHOSPITAL   IN INTEGER, -- Clave de la empresa contable
        'IN_CHRFACPACIENTES  IN CHAR,    -- P= Mostrar facturas de pacientes y grupos de cuentas
        'IN_CHRFACDIRECTAS   IN CHAR,    -- C= Mostrar facturas directas
        'IN_CHRFACVP         IN CHAR,    -- V= Mostrar facturas de venta al público
        'IN_CHRNOTASCREDITO  IN CHAR,    -- CR= Mostrar notas de crédito
        'IN_CHRNOTASCARGO    IN CHAR,    -- CA= Mostrar notas de cargo
        'IN_INTCVETIPOPAC    IN INTEGER, -- Clave del tipo de paciente, -1 =Todos
        'IN_CHRTIPOPAC       IN CHAR,    -- Tipo de paciente ME, EM, PA, CO, * = Todos
        'IN_INTCVECONVENIO   IN INTEGER, -- Clave del tipo de convenio, -1 =Todos
        'IN_INTCVEEMPRESA    IN INTEGER, -- Clave de la empresa, -1 =Todos
        'IN_CHRFORMAPAGO     IN CHAR,    -- Forma de pago *=todas, C=crédito E=Contado
        'IN_VCHFECHAINICIO   IN VARCHAR2,-- Fecha de inicio
        'IN_VCHFECHAFIN      IN VARCHAR2,-- Fecha fin
        
        vgstrParametrosSP = IIf(cboDepartamento.ListIndex = 0, "-1", Str(cboDepartamento.ItemData(cboDepartamento.ListIndex))) _
        & "|" & vgintClaveEmpresaContable _
        & "|" & IIf(chkFacPacientes.Value = vbChecked, "'P'", "'*'") _
        & "|" & IIf(chkFacDirectas.Value = vbChecked, "'C'", "'*'") _
        & "|" & IIf(chkFacVentapublico.Value = vbChecked, "'V'", "'*'") _
        & "|" & IIf(chkNotasCredito.Value = vbChecked, "'CR'", "'*'") _
        & "|" & IIf(chkNotasCargo.Value = vbChecked, "'CA'", "'*'") _
        & "|" & IIf(cboTipoPaciente.ListIndex = 0, "-1", Str(cboTipoPaciente.ItemData(cboTipoPaciente.ListIndex))) _
        & "|" & IIf(lstrTipoPac = "", "'*'", "'" & lstrTipoPac & "'") _
        & "|" & IIf(cboTipoConvenio.ListIndex = 0, "-1", Str(cboTipoConvenio.ItemData(cboTipoConvenio.ListIndex))) _
        & "|" & IIf(cboEmpresa.ListIndex = 0, "-1", Str(cboEmpresa.ItemData(cboEmpresa.ListIndex))) _
        & "|" & IIf(optFormaPago(0).Value, "'*'", IIf(optFormaPago(1).Value, "'C'", "'E'")) _
        & "|" & fstrFechaSQL(mskFechaInicio.Text) _
        & "|" & fstrFechaSQL(mskFechaFin.Text)
        Set rsAux = frsEjecuta_SP(vgstrParametrosSP, "SP_PVRPTCRONOLOGICOFACTURAS")
        If rsAux.RecordCount > 0 Then
            freBarra.Visible = True
            pgbBarra.Value = 0
            
            Set o_Excel = CreateObject("Excel.Application")
            Set o_Libro = o_Excel.Workbooks.Add
            Set o_Sheet = o_Libro.Worksheets(1)
            
            'datos del repote
            o_Excel.Cells(2, 1).Value = Trim(vgstrNombreHospitalCH)
            o_Excel.Cells(3, 1).Value = "CRONOLÓGICO DE FACTURAS Y NOTAS"
            o_Excel.Cells(4, 1).Value = "Del " & CStr(Format(CDate(mskFechaInicio), "dd/MMM/yyyy")) & " Al " & CStr(Format(CDate(mskFechaFin), "dd/MMM/yyyy"))
            'columnas titulos
            o_Excel.Cells(6, 1).Value = "Fecha movimiento"
            o_Excel.Cells(6, 2).Value = "Folio"
            o_Excel.Cells(6, 3).Value = "Fecha documento"
            o_Excel.Cells(6, 4).Value = "Estado"
            o_Excel.Cells(6, 5).Value = "Forma de pago"
            o_Excel.Cells(6, 6).Value = "Fecha primer documento"
            o_Excel.Cells(6, 7).Value = "R.F.C."
            o_Excel.Cells(6, 8).Value = "Razón social"
            o_Excel.Cells(6, 9).Value = "Cuenta/ Clave"
            o_Excel.Cells(6, 10).Value = "Tipo"
            o_Excel.Cells(6, 11).Value = "Paciente/ Cliente"
            o_Excel.Cells(6, 12).Value = "Empresa"
            o_Excel.Cells(6, 13).Value = "Médico tratante"
            o_Excel.Cells(6, 14).Value = "Importe gravado factura"
            o_Excel.Cells(6, 15).Value = "Importe no gravado factura"
            '***
            o_Excel.Cells(6, 16).Value = "Importe gravado nota"
            o_Excel.Cells(6, 17).Value = "Importe no gravado nota"
            '***
            o_Excel.Cells(6, 18).Value = "Descuento gravado"
            o_Excel.Cells(6, 19).Value = "Descuento no gravado"
            o_Excel.Cells(6, 20).Value = "Subtotal gravado"
            o_Excel.Cells(6, 21).Value = "Subtotal no gravado"
            o_Excel.Cells(6, 22).Value = "IVA"
            o_Excel.Cells(6, 23).Value = "Total"
            
            'sumatorias
            o_Excel.Cells(rsAux.RecordCount + 7, 13).Formula = "TOTAL"
            o_Excel.Cells(rsAux.RecordCount + 7, 14).Formula = "=SUM(N7:N" & (rsAux.RecordCount + 6) & ")"
            o_Excel.Cells(rsAux.RecordCount + 7, 15).Formula = "=SUM(O7:O" & (rsAux.RecordCount + 6) & ")"
            '***
            o_Excel.Cells(rsAux.RecordCount + 7, 16).Formula = "=SUM(P7:P" & (rsAux.RecordCount + 6) & ")"
            o_Excel.Cells(rsAux.RecordCount + 7, 17).Formula = "=SUM(Q7:Q" & (rsAux.RecordCount + 6) & ")"
            '***
            o_Excel.Cells(rsAux.RecordCount + 7, 18).Formula = "=SUM(R7:R" & (rsAux.RecordCount + 6) & ")"
            o_Excel.Cells(rsAux.RecordCount + 7, 19).Formula = "=SUM(S7:S" & (rsAux.RecordCount + 6) & ")"
            o_Excel.Cells(rsAux.RecordCount + 7, 20).Formula = "=SUM(T7:T" & (rsAux.RecordCount + 6) & ")"
            o_Excel.Cells(rsAux.RecordCount + 7, 21).Formula = "=SUM(U7:U" & (rsAux.RecordCount + 6) & ")"
            o_Excel.Cells(rsAux.RecordCount + 7, 22).Formula = "=SUM(V7:V" & (rsAux.RecordCount + 6) & ")"
            o_Excel.Cells(rsAux.RecordCount + 7, 23).Formula = "=SUM(W7:W" & (rsAux.RecordCount + 6) & ")"
            
            pgbBarra.Value = 30
            
            o_Sheet.Range("A6:X6").HorizontalAlignment = -4108
            o_Sheet.Range("A6:X6").VerticalAlignment = -4108
            o_Sheet.Range("A6:X6").WrapText = True
            o_Sheet.Range("A7").Select
            o_Excel.ActiveWindow.FreezePanes = True
            o_Sheet.Range("A6:W6").Interior.ColorIndex = 15 '15 48
            o_Sheet.Range(o_Excel.Cells(rsAux.RecordCount + 7, 13), o_Excel.Cells(rsAux.RecordCount + 7, 23)).Interior.ColorIndex = 15
            o_Sheet.Range("A:A").ColumnWidth = 10
            o_Sheet.Range("B:B").ColumnWidth = 8
            o_Sheet.Range("C:C").ColumnWidth = 9
            o_Sheet.Range("D:D").ColumnWidth = 8
            o_Sheet.Range("E:E").ColumnWidth = 9
            o_Sheet.Range("G:G").ColumnWidth = 14
            o_Sheet.Range("H:H").ColumnWidth = 20
            o_Sheet.Range("I:I").ColumnWidth = 7
            o_Sheet.Range("J:J").ColumnWidth = 6
            o_Sheet.Range("K:K").ColumnWidth = 20
            o_Sheet.Range("L:L").ColumnWidth = 20
            o_Sheet.Range("M:M").ColumnWidth = 20
            
            o_Sheet.Range(o_Excel.Cells(rsAux.RecordCount + 6, 1), o_Excel.Cells(rsAux.RecordCount + 6, 23)).Borders(4).LineStyle = 1
            
            'info del rs
            o_Sheet.Range("A:X").Font.Size = 9
            o_Sheet.Range("A:X").Font.Name = "Times New Roman" '
            o_Sheet.Range("A:X").Font.Bold = False
            pgbBarra.Value = 70
            'titulos
            o_Sheet.Range("A6:X6").Font.Bold = True
            o_Sheet.Range(o_Excel.Cells(rsAux.RecordCount + 7, 1), o_Excel.Cells(rsAux.RecordCount + 7, 23)).Font.Bold = True
            o_Sheet.Range(o_Excel.Cells(2, 1), o_Excel.Cells(5, 1)).Font.Bold = True
            'centrado, auto ajustar texto, alinear medio
            o_Sheet.Range("N:W").NumberFormat = "$ ###,###,###,##0.00"
            
            'rs,maxRows,maxCols
            o_Sheet.Range("A7").CopyFromRecordset rsAux, , 23
            
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
Private Function fblnFechasValidas() As Boolean
On Error GoTo NotificaError
    fblnFechasValidas = True
    If Not IsDate(mskFechaInicio) Then
        MsgBox SIHOMsg(29), vbOKOnly + vbInformation, "Mensaje"
        fblnFechasValidas = False
        mskFechaInicio.SetFocus
    ElseIf Not IsDate(mskFechaFin) Then
        MsgBox SIHOMsg(29), vbCritical, "Mensaje"
        fblnFechasValidas = False
        mskFechaFin.SetFocus
    'fecha final valida
    ElseIf CDate(mskFechaInicio) > CDate(mskFechaFin) Then
        MsgBox SIHOMsg(64), vbCritical, "Mensaje"
        fblnFechasValidas = False
        mskFechaInicio.SetFocus
    End If
Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnFechasValidas"))
End Function
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
    Me.Icon = frmMenuPrincipal.Icon
    vgstrNombreForm = Me.Name
    pIniciar
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
End Sub

Private Sub pIniciar()
On Error GoTo NotificaError
    Dim rsAux As New ADODB.Recordset
    
    optFormaPago(0).Value = True
    mskFechaInicio.Mask = ""
    mskFechaInicio.Text = ""
    mskFechaInicio.Mask = "##/##/####"
    mskFechaFin.Mask = ""
    mskFechaFin.Text = ""
    mskFechaFin.Mask = "##/##/####"
    mskFechaInicio = FormatDateTime(fdtmServerFecha, vbShortDate)
    mskFechaFin = FormatDateTime(fdtmServerFecha, vbShortDate)
    chkFacPacientes.Value = vbChecked
    chkFacDirectas.Value = vbChecked
    chkFacVentapublico.Value = vbChecked
    chkNotasCredito.Value = vbChecked
    chkNotasCargo.Value = vbChecked
        
    '--- departamento ---
     Set rsAux = frsEjecuta_SP("-1", "SP_GNSELDEPARTAMENTOS")
    If rsAux.RecordCount > 0 Then
        pLlenarCboRs cboDepartamento, rsAux, 0, 1, 0
    End If
    cboDepartamento.AddItem "<TODOS>", 0
    cboDepartamento.ListIndex = 0
    
    '--- tipo paciente ---
    Set rsAux = frsEjecuta_SP("-1", "Sp_Adseltipopaciente")
    If rsAux.RecordCount > 0 Then
        pLlenarCboRs cboTipoPaciente, rsAux, 0, 1, 0
    End If
    cboTipoPaciente.AddItem "<TODOS>", 0
    cboTipoPaciente.ListIndex = 0
    
    pLlenaComboConvenio True
    pLlenaComboEmpresa -1
    
    lstrTipoPac = ""
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pIniciar"))
End Sub
Private Sub pLlenaComboConvenio(blnCatalogo As Boolean)
On Error GoTo NotificaError
    Dim rsAux As New ADODB.Recordset
    cboTipoConvenio.Clear
    If blnCatalogo Then
        Set rsAux = frsEjecuta_SP("", "SP_ADTIPOCONVENIO")
        If rsAux.RecordCount > 0 Then
            pLlenarCboRs cboTipoConvenio, rsAux, 0, 1, 0
        End If
    End If
    cboTipoConvenio.AddItem "<TODOS>", 0
    cboTipoConvenio.ListIndex = 0
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLlenaComboConvenio"))
End Sub
Private Sub pLlenaComboEmpresa(lngId As Long)
On Error GoTo NotificaError
    Dim rsAux As New ADODB.Recordset
    cboEmpresa.Clear
    If lstrTipoPac = "CO" Or lstrTipoPac = "" Then
        Set rsAux = frsEjecuta_SP("-1|" & lngId & "|-1", "SP_CCSELEMPRESA")
        If rsAux.RecordCount > 0 Then
            pLlenarCboRs cboEmpresa, rsAux, 0, 1, 0
        End If
    End If
    cboEmpresa.AddItem "<TODOS>", 0
    cboEmpresa.ListIndex = 0
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLlenaComboEmpresa"))
End Sub
Private Sub mskFechaFin_GotFocus()
On Error GoTo NotificaError
    pSelMkTexto mskFechaFin
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFechaFin_GotFocus"))
End Sub

Private Sub mskFechaInicio_GotFocus()
On Error GoTo NotificaError
    pSelMkTexto mskFechaInicio
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFechaInicio_GotFocus"))
End Sub
