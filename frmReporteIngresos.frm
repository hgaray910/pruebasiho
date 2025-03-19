VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmReporteIngresos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ingresos por tipo de paciente"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7230
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkMostrarCancelados 
      Caption         =   "Mostrar cargos cancelados"
      Height          =   375
      Left            =   4920
      TabIndex        =   34
      Top             =   4500
      Value           =   1  'Checked
      Width           =   2300
   End
   Begin VB.Frame Frame5 
      Height          =   675
      Left            =   60
      TabIndex        =   32
      Top             =   -30
      Width           =   7125
      Begin VB.ComboBox cboHospital 
         Height          =   315
         Left            =   915
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Seleccione la empresa"
         Top             =   240
         Width           =   6015
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Empresa"
         Height          =   195
         Left            =   120
         TabIndex        =   33
         Top             =   300
         Width           =   615
      End
   End
   Begin VB.Frame Frame6 
      Height          =   735
      Left            =   3075
      TabIndex        =   31
      Top             =   4260
      Width           =   1140
      Begin VB.CommandButton cmdImprimir 
         Height          =   495
         Left            =   570
         Picture         =   "frmReporteIngresos.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Imprimir"
         Top             =   165
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdVistaPreliminar 
         Height          =   495
         Left            =   75
         Picture         =   "frmReporteIngresos.frx":01A2
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Vista previa"
         Top             =   165
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin TabDlg.SSTab sstReporte 
      Height          =   3495
      Left            =   45
      TabIndex        =   21
      Top             =   720
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   6165
      _Version        =   393216
      Style           =   1
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tipo de reporte"
      TabPicture(0)   =   "frmReporteIngresos.frx":0344
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Filtros avanzados"
      TabPicture(1)   =   "frmReporteIngresos.frx":0360
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Cargos"
      TabPicture(2)   =   "frmReporteIngresos.frx":037C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4"
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame4 
         Height          =   3045
         Left            =   -74910
         TabIndex        =   29
         Top             =   315
         Width           =   6945
         Begin VB.TextBox txtCargo 
            Height          =   285
            Left            =   120
            TabIndex        =   17
            Top             =   705
            Width           =   6705
         End
         Begin VB.OptionButton optTipoGrupoExamen 
            Caption         =   "Grupo de examen"
            Height          =   225
            Left            =   2730
            TabIndex        =   16
            Top             =   435
            Width           =   1620
         End
         Begin VB.OptionButton optTipoOtroConcepto 
            Caption         =   "Otros conceptos"
            Height          =   210
            Left            =   1140
            TabIndex        =   14
            Top             =   435
            Width           =   1470
         End
         Begin VB.OptionButton optTipoExamen 
            Caption         =   "Exámenes"
            Height          =   240
            Left            =   2730
            TabIndex        =   15
            Top             =   165
            Width           =   1065
         End
         Begin VB.OptionButton optTipoEstudio 
            Caption         =   "Estudios"
            Height          =   210
            Left            =   1140
            TabIndex        =   13
            Top             =   165
            Width           =   1050
         End
         Begin VB.OptionButton optTipoArticulos 
            Caption         =   "Artículos"
            Height          =   210
            Left            =   120
            TabIndex        =   12
            Top             =   435
            Width           =   960
         End
         Begin VB.OptionButton optTipoTodos 
            Caption         =   "Todos"
            Height          =   210
            Left            =   120
            TabIndex        =   11
            Top             =   165
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.ListBox lstCargo 
            Height          =   1815
            Left            =   120
            TabIndex        =   18
            Top             =   1035
            Width           =   6705
         End
      End
      Begin VB.Frame Frame2 
         Height          =   3045
         Left            =   90
         TabIndex        =   23
         Top             =   315
         Width           =   6960
         Begin VB.ComboBox cboTipoIngreso 
            Height          =   315
            Left            =   1965
            Style           =   2  'Dropdown List
            TabIndex        =   5
            ToolTipText     =   "Tipo de ingreso"
            Top             =   810
            Width           =   4800
         End
         Begin VB.ComboBox cboDepartamento 
            Height          =   315
            Left            =   1965
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   1575
            Width           =   4800
         End
         Begin VB.ComboBox cboEmpresa 
            Height          =   315
            Left            =   1965
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   2370
            Width           =   4800
         End
         Begin VB.ComboBox cboConcepto 
            Height          =   315
            Left            =   1965
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   1980
            Width           =   4800
         End
         Begin VB.OptionButton optTipoPaciente 
            Caption         =   "Internos"
            Height          =   255
            Index           =   0
            Left            =   2850
            TabIndex        =   3
            Top             =   450
            Width           =   960
         End
         Begin VB.OptionButton optTipoPaciente 
            Caption         =   "Externos"
            Height          =   285
            Index           =   1
            Left            =   3930
            TabIndex        =   4
            Top             =   435
            Width           =   945
         End
         Begin VB.OptionButton optTipoPaciente 
            Caption         =   "Todos"
            Height          =   285
            Index           =   2
            Left            =   1935
            TabIndex        =   2
            Top             =   450
            Value           =   -1  'True
            Width           =   870
         End
         Begin MSMask.MaskEdBox mskInicio 
            Height          =   315
            Left            =   1965
            TabIndex        =   6
            ToolTipText     =   "Fecha de inicio"
            Top             =   1200
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
            Left            =   4035
            TabIndex        =   7
            ToolTipText     =   "Fecha de fin"
            Top             =   1200
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mmm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin VB.Label Label8 
            Caption         =   "Tipo de ingreso"
            Height          =   195
            Left            =   180
            TabIndex        =   35
            Top             =   870
            Width           =   1200
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Departamento"
            Height          =   195
            Left            =   180
            TabIndex        =   30
            Top             =   1635
            Width           =   1005
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Concepto de facturación"
            Height          =   195
            Left            =   180
            TabIndex        =   28
            Top             =   2040
            Width           =   1755
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Procedencia"
            Height          =   195
            Left            =   180
            TabIndex        =   27
            Top             =   2430
            Width           =   900
         End
         Begin VB.Label Label6 
            Caption         =   "Tipo de paciente"
            Height          =   195
            Left            =   180
            TabIndex        =   26
            Top             =   480
            Width           =   1440
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fechas de factura"
            Height          =   195
            Left            =   180
            TabIndex        =   25
            Top             =   1260
            Width           =   1290
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            Height          =   195
            Left            =   3450
            TabIndex        =   24
            Top             =   1260
            Width           =   420
         End
      End
      Begin VB.Frame Frame1 
         Height          =   3045
         Left            =   -74910
         TabIndex        =   22
         Top             =   315
         Width           =   6975
         Begin VB.ListBox lstTipoReporte 
            Height          =   2595
            Left            =   120
            TabIndex        =   1
            Top             =   270
            Width           =   6765
         End
      End
   End
End
Attribute VB_Name = "frmReporteIngresos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'-------------------------------------------------------------------------------------
'| Nombre del Proyecto      : Caja                                                   -
'| Nombre del Formulario    : frmReporteIngresos                                     -
'-------------------------------------------------------------------------------------
'| Objetivo: Es el reporte de ingresos de caja
'-------------------------------------------------------------------------------------
'| Análisis y Diseño        : Rodolfo Ramos G.
'| Autor                    : Rodolfo Ramos G.
'| Fecha de Creación        : 15/Mar/2002
'| Modificó                 : Nombre(s)
'| Fecha Terminación        : hoy
'| Fecha última modificación: 16/Jun/2003
'-------------------------------------------------------------------------------------
Option Explicit

Private Sub pCargaCargos()
    Dim vlstrSentencia As String
    Dim rs As New ADODB.Recordset
    
    If optTipoArticulos.Value Then
        vlstrSentencia = "select intIdArticulo Clave, vchNombreComercial Descripcion from ivArticulo "
        PSuperBusqueda txtCargo, vlstrSentencia, lstCargo, "vchNombreComercial", 20, " and CHRCVEARTMEDICAMEN <> 2 and vchEstatus = 'ACTIVO' ", "vchNombreComercial"
    ElseIf optTipoEstudio.Value Then
        vlstrSentencia = "select intCveEstudio Clave, vchNombre Descripcion from imEstudio "
        PSuperBusqueda txtCargo, vlstrSentencia, lstCargo, "vchNombre", 20, " and bitStatusActivo = 1 ", "vchNombre"
    ElseIf optTipoExamen.Value Then
        vlstrSentencia = "select intCveExamen Clave, chrNombre Descripcion from laExamen "
        PSuperBusqueda txtCargo, vlstrSentencia, lstCargo, "chrNombre", 20, " and bitEstatusActivo = 1 ", "chrNombre"
    ElseIf optTipoGrupoExamen.Value Then
        vlstrSentencia = "select intCveGrupo Clave, chrNombre Descripcion from laGrupoExamen "
        PSuperBusqueda txtCargo, vlstrSentencia, lstCargo, "chrNombre", 20, " and bitEstatusActivo = 1 ", "chrNombre"
    ElseIf optTipoOtroConcepto.Value Then
        vlstrSentencia = "select intCveConcepto Clave, chrDescripcion Descripcion from pvOtroConcepto "
        PSuperBusqueda txtCargo, vlstrSentencia, lstCargo, "chrDescripcion", 20, " and bitEstatus = 1 ", "chrDescripcion"
    End If
    
'    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
'    With lstCargo
'        .Enabled = False
'        .Visible = False
'        .Clear
'        Do While Not rs.EOF
'            .AddItem rs!Descripcion
'            .ItemData(.NewIndex) = rs!Clave
'            rs.MoveNext
'        Loop
'        .Visible = True
'        .Enabled = rs.RecordCount > 0
'    End With
'    rs.Close
End Sub

Private Sub txtArticulo_KeyUp(KeyCode As Integer, Shift As Integer)
    pCargaCargos
End Sub


Private Sub cboDepartamento_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        cboConcepto.SetFocus
    End If

End Sub


Private Sub cboHospital_Click()
    On Error GoTo NotificaError
    
    Dim rs As New ADODB.Recordset

    If cboHospital.ListIndex <> -1 Then
        cboDepartamento.Clear
        vgstrParametrosSP = "-1|1|*|" & CStr(cboHospital.ItemData(cboHospital.ListIndex))
        Set rs = frsEjecuta_SP(vgstrParametrosSP, "Sp_Gnseldepartamento")
        If rs.RecordCount <> 0 Then
            pLlenarCboRs cboDepartamento, rs, 0, 1
        End If
        cboDepartamento.AddItem "<TODOS>", 0
        cboDepartamento.ItemData(cboDepartamento.newIndex) = -1
        cboDepartamento.ListIndex = 0
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboHospital_Click"))
    Unload Me
End Sub

Private Sub cboHospital_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        lstTipoReporte.SetFocus
    End If

End Sub

Private Sub cboTipoIngreso_KeyDown(KeyCode As Integer, Shift As Integer)
     
    If KeyCode = vbKeyReturn Then
        mskInicio.SetFocus
    End If
        
End Sub


Private Sub cmdImprimir_Click()
    pImprime "I"
End Sub

Private Sub cmdPreview_Click()

End Sub

Private Sub cmdPrint_Click()

End Sub

Private Sub cmdVistaPreliminar_Click()
    pImprime "P"
End Sub

Private Sub Combo1_Change()

End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        mskInicio.SetFocus
    End If

End Sub
Private Sub Form_Activate()
    If cboConcepto.ListCount = 1 Then
        MsgBox SIHOMsg(430), vbExclamation, "Mensaje"
        Unload Me
    Else
        cboTipoIngreso.SetFocus
    End If
End Sub

Private Sub Form_Load()
    
    Dim vlstrSentencia As String
    Dim rs As New ADODB.Recordset
    Dim lngNumOpcion As Long
    Dim dtmFecha As Date
    
    Me.Icon = frmMenuPrincipal.Icon
    
    Select Case cgstrModulo
    Case "PV"
         lngNumOpcion = 341
    Case "SE"
         lngNumOpcion = 1533
    End Select
   
    pCargaHospital lngNumOpcion
    
    ' Cargar el de Tipos de Reportes
    
    With lstTipoReporte
        .AddItem "Fecha - Concepto"
        .ItemData(.newIndex) = 1
        .AddItem "Fecha - Empresa"
        .ItemData(.newIndex) = 8
        .AddItem "Fecha - Médico"
        .ItemData(.newIndex) = 9
        .AddItem "Tipo paciente - Empresa"
        .ItemData(.newIndex) = 2
        .AddItem "Empresa - Concepto"
        .ItemData(.newIndex) = 3
        .AddItem "Empresa - Paciente"
        .ItemData(.newIndex) = 4
        .AddItem "Concepto Facturación - Médico"
        .ItemData(.newIndex) = 5
        .AddItem "Concepto Facturación - Empresa"
        .ItemData(.newIndex) = 6
        .AddItem "Concepto Facturación - Paciente"
        .ItemData(.newIndex) = 7
        .AddItem "Mes - Concepto"
        .ItemData(.newIndex) = 10
        .AddItem "Mes - Médico"
        .ItemData(.newIndex) = 11
        .AddItem "Médico - Mes"
        .ItemData(.newIndex) = 12
        .AddItem "Médico - Cargo"
        .ItemData(.newIndex) = 13
        .ListIndex = 0
    End With
    
    'Conceptos de Facturación
    'No está filtrado por departamento porque así lo pidio Sergio
    vlstrSentencia = "select smiCveConcepto, chrDescripcion from pvConceptoFacturacion where bitActivo = 1"
    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    pLlenarCboRs cboConcepto, rs, 0, 1
    rs.Close
    
    cboConcepto.AddItem "<TODOS>", 0
    cboConcepto.ItemData(cboConcepto.newIndex) = -1
    cboConcepto.ListIndex = 0
    
    'Tipos de Ingreso
    vlstrSentencia = "select intcvetipoingreso, vchnombre from siTipoIngreso"
    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    pLlenarCboRs cboTipoIngreso, rs, 0, 1
    rs.Close
    
    cboTipoIngreso.AddItem "<TODOS>", 0
    cboTipoIngreso.ItemData(cboTipoIngreso.newIndex) = -1
    cboTipoIngreso.ListIndex = 0
    
    'Empresas
    vlstrSentencia = "select intCveEmpresa, vchDescripcion from ccEmpresa "
    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    pLlenarCboRs cboEmpresa, rs, 0, 1
    rs.Close
    
    'Tipos de Paciente
    vlstrSentencia = "Select tnyCveTipoPaciente, vchDescripcion from adTipoPaciente order by tnyCveTipoPaciente desc"
    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    Do While Not rs.EOF
        With cboEmpresa
            .AddItem rs!VCHDESCRIPCION, 0
            .ItemData(.newIndex) = rs!tnyCveTipoPaciente * -1
        End With
        rs.MoveNext
    Loop
    rs.Close
    
    cboEmpresa.AddItem "<TODOS>", 0
    cboEmpresa.ItemData(0) = 0
    cboEmpresa.ListIndex = 0
    
    dtmFecha = fdtmServerFecha
    
    mskInicio.Mask = ""
    mskInicio.Text = dtmFecha
    mskInicio.Mask = "##/##/####"
    
    mskFin.Mask = ""
    mskFin.Text = dtmFecha
    mskFin.Mask = "##/##/####"


End Sub


Private Sub lstCargo_DblClick()
    lstCargo_KeyDown vbKeyReturn, 0
End Sub

Private Sub lstCargo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        sstReporte.Tab = 1
        If fblnCanFocus(cmdImprimir) Then cmdImprimir.SetFocus
    End If
End Sub

Private Sub lstTipoReporte_DblClick()
    lstTipoReporte_KeyDown vbKeyReturn, 0
End Sub

Private Sub lstTipoReporte_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        sstReporte.Tab = 1
        OptTipoPaciente(0).SetFocus
    End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub mskFin_GotFocus()
    pSelMkTexto mskInicio
End Sub

Private Sub mskFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cboDepartamento.SetFocus
    End If
End Sub

Private Sub mskInicio_GotFocus()
    
    pSelMkTexto mskInicio
    
End Sub

Private Sub mskInicio_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        pEnfocaMkTexto mskFin
    End If
End Sub


Private Sub optTipoPaciente_Click(Index As Integer)

    pCargaCboTipoIngreso

End Sub

Private Sub optTipoPaciente_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cboTipoIngreso.SetFocus
    End If
End Sub

Private Sub cboConcepto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cboEmpresa.SetFocus
    End If
End Sub

Private Sub cboEmpresa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       cmdImprimir.SetFocus
    End If
End Sub

Private Sub optTipoArticulos_Click()
    pCargaCargos
End Sub

Private Sub optTipoEstudio_Click()
    pCargaCargos
End Sub

Private Sub optTipoExamen_Click()
    pCargaCargos
End Sub

Private Sub optTipoGrupoExamen_Click()
    pCargaCargos
End Sub

Private Sub optTipoOtroConcepto_Click()
    pCargaCargos
End Sub

Private Sub optTipoPaciente_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    pCargaCargos
End Sub
Private Sub optTipoArticulos_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtCargo.Enabled = True
    pCargaCargos
    pEnfocaTextBox txtCargo
End Sub

Private Sub optTipoEstudio_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtCargo.Enabled = True
    pCargaCargos
    pEnfocaTextBox txtCargo
End Sub

Private Sub optTipoExamen_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtCargo.Enabled = True
    pCargaCargos
    pEnfocaTextBox txtCargo
End Sub

Private Sub optTipoGrupoExamen_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtCargo.Enabled = True
    pCargaCargos
    pEnfocaTextBox txtCargo
End Sub

Private Sub optTipoOtroConcepto_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtCargo.Enabled = True
    pCargaCargos
    pEnfocaTextBox txtCargo
End Sub

Private Sub optTipoTodos_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtCargo.Enabled = False
    lstCargo.Clear
    lstCargo.Enabled = False
End Sub

Private Sub txtCargo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtcargo_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyReturn Or KeyCode = vbKeyRight Or KeyCode = vbKeyDown) And lstCargo.Enabled Then
        lstCargo.SetFocus
    End If
End Sub

Private Sub txtCargo_KeyUp(KeyCode As Integer, Shift As Integer)
    pCargaCargos
End Sub


Sub pImprime(pstrDestino As String)
    Dim vgrptReporte As CRAXDRT.Report
    Dim vlstrTipoCargo As String
    Dim vlstrCveCargo As String
    Dim vlstrSentencia As String
    Dim rs As New ADODB.Recordset
    Dim rsReporte As ADODB.Recordset
    Dim alstrParametros(5) As String
    Dim strInternoExterno As String
    Dim strTitulo1 As String
    
    
    If Not IsDate(mskInicio.Text) Then
        MsgBox SIHOMsg(29), vbInformation, "Mensaje"
        pEnfocaMkTexto mskInicio
        Exit Sub
    ElseIf Not IsDate(mskFin.Text) Then
        MsgBox SIHOMsg(29), vbInformation, "Mensaje"
        pEnfocaMkTexto mskFin
        Exit Sub
    End If
    
    If OptTipoPaciente(0).Value Then
        strInternoExterno = "(Internos)"
    ElseIf OptTipoPaciente(1).Value Then
        strInternoExterno = "(Externos)"
    ElseIf OptTipoPaciente(2).Value Then
        strInternoExterno = "(Internos y Externos)"
    End If
        
    If lstTipoReporte.ListIndex = 1 Then
        strTitulo1 = "Fecha-Concepto " & strInternoExterno
    ElseIf lstTipoReporte.ListIndex = 2 Then
        strTitulo1 = "Tipo-Empresa " & strInternoExterno
    ElseIf lstTipoReporte.ListIndex = 3 Then
        strTitulo1 = "Empresa-Concepto " & strInternoExterno
    ElseIf lstTipoReporte.ListIndex = 4 Then
        strTitulo1 = "Empresa-Paciente " & strInternoExterno
    ElseIf lstTipoReporte.ListIndex = 5 Then
        strTitulo1 = "Concepto-Médico " & strInternoExterno
    ElseIf lstTipoReporte.ListIndex = 6 Then
        strTitulo1 = "Concepto-Empresa " & strInternoExterno
    ElseIf lstTipoReporte.ListIndex = 7 Then
        strTitulo1 = "Concepto-Paciente " & strInternoExterno
    ElseIf lstTipoReporte.ListIndex = 8 Then
        strTitulo1 = "Fecha-Empresa " & strInternoExterno
    ElseIf lstTipoReporte.ListIndex = 9 Then
        strTitulo1 = "Fecha-Médico " & strInternoExterno
    ElseIf lstTipoReporte.ListIndex = 10 Then
        strTitulo1 = "Mes-Concepto " & strInternoExterno
    ElseIf lstTipoReporte.ListIndex = 11 Then
        strTitulo1 = "Mes-Médico " & strInternoExterno
    ElseIf lstTipoReporte.ListIndex = 12 Then
        strTitulo1 = "Médico y Mes " & strInternoExterno
    ElseIf lstTipoReporte.ListIndex = 13 Then
        strTitulo1 = "Médico y Cargo " & strInternoExterno
    End If

    
    If optTipoArticulos.Value Then
        vlstrTipoCargo = "AR"
    ElseIf optTipoEstudio.Value Then
        vlstrTipoCargo = "ES"
    ElseIf optTipoExamen.Value Then
        vlstrTipoCargo = "EX"
    ElseIf optTipoGrupoExamen.Value Then
        vlstrTipoCargo = "GE"
    ElseIf optTipoOtroConcepto.Value Then
        vlstrTipoCargo = "OC"
    ElseIf optTipoTodos.Value Then
        vlstrTipoCargo = "TO"
    End If
    vlstrCveCargo = "0"
    If lstCargo.ListIndex <> -1 And Not optTipoTodos.Value Then
        If optTipoArticulos.Value Then
            vlstrCveCargo = Trim(Str(lstCargo.ItemData(lstCargo.ListIndex)))
        Else
            vlstrCveCargo = Trim(Str(lstCargo.ItemData(lstCargo.ListIndex)))
        End If
    End If
    
    vgstrParametrosSP = _
    CStr(lstTipoReporte.ItemData(lstTipoReporte.ListIndex)) & _
    "|" & IIf(OptTipoPaciente(0).Value, "I", IIf(OptTipoPaciente(1).Value, "E", "A")) & _
    "|" & CStr(IIf(cboConcepto.ListIndex = 0, -1, cboConcepto.ItemData(cboConcepto.ListIndex))) & _
    "|" & CStr(cboDepartamento.ItemData(cboDepartamento.ListIndex)) & _
    "|" & CStr(cboEmpresa.ItemData(cboEmpresa.ListIndex)) & _
    "|" & vlstrTipoCargo & _
    "|" & vlstrCveCargo & _
    "|" & fstrFechaSQL(mskInicio.Text) & _
    "|" & fstrFechaSQL(mskFin.Text) & _
    "|" & Str(cboHospital.ItemData(cboHospital.ListIndex)) & _
    "|" & IIf(chkMostrarCancelados.Value = 1, 1, 0) & _
    "|" & CStr(cboTipoIngreso.ItemData(cboTipoIngreso.ListIndex))
    
    
    
    Set rs = frsEjecuta_SP(vgstrParametrosSP, "Sp_Pvselreporteingresos", , , , True)
    If rs.EOF Then
        MsgBox SIHOMsg(13), vbInformation, "Mensaje"
    Else
        pInstanciaReporte vgrptReporte, "rptIngresos.rpt"
        
        vgrptReporte.DiscardSavedData
        
        alstrParametros(0) = "FechaInicial" & ";" & Format(mskInicio.Text, "dd/mmm/yyyy") & ";"
        alstrParametros(1) = "FechaFinal" & ";" & Format(mskFin.Text, "dd/mmm/yyyy") & ";"
        alstrParametros(2) = "NombreHospital" & ";" & Trim(cboHospital.List(cboHospital.ListIndex)) & ";"
        alstrParametros(3) = "Titulo1" & ";" & strTitulo1 & ";"
        alstrParametros(4) = "TipoDB" & ";" & vgstrBaseDatosUtilizada & ";"
        'alstrParametros(5) = "TipoIngreso;" & Trim(cboTipoIngreso.List(cboTipoIngreso.ListIndex))
        alstrParametros(5) = "TipoIngreso;" & IIf(Trim(cboTipoIngreso.List(cboTipoIngreso.ListIndex)) = "<TODOS>", "Todos", Trim(cboTipoIngreso.List(cboTipoIngreso.ListIndex)))
        

        pCargaParameterFields alstrParametros, vgrptReporte
        
        pImprimeReporte vgrptReporte, rs, pstrDestino, "Reporte de ingresos"
    End If
    rs.Close
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



Public Sub pCargaCboTipoIngreso()

    Dim vlstrSentencia As String
    Dim rs As New ADODB.Recordset

    'Tipos de Paciente
    If OptTipoPaciente(2) = True Then
        ' Todos los Pacientes
        vlstrSentencia = "select intcvetipoingreso, vchnombre from siTipoIngreso"
    Else
        If OptTipoPaciente(0) = True Then
            'Pacientes Internos
            vlstrSentencia = "select intcvetipoingreso, vchnombre from siTipoIngreso where CHRTIPOINGRESO = 'I'"
        Else
            'Pacientes Externos
            vlstrSentencia = "select intcvetipoingreso, vchnombre from siTipoIngreso where CHRTIPOINGRESO = 'E'"
        End If
    End If
    
    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    pLlenarCboRs cboTipoIngreso, rs, 0, 1
    rs.Close
    
    cboTipoIngreso.AddItem "<TODOS>", 0
    cboTipoIngreso.ItemData(cboTipoIngreso.newIndex) = -1
    cboTipoIngreso.ListIndex = 0

End Sub
