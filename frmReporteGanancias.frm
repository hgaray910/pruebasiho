VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmReporteGanancias 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ganancias"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7275
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   7275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      Height          =   735
      Left            =   3067
      TabIndex        =   37
      Top             =   3825
      Width           =   1140
      Begin VB.CommandButton cmdVistaPreliminar 
         Height          =   495
         Left            =   75
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmReporteGanancias.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Vista previa"
         Top             =   165
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdImprimir 
         Height          =   495
         Left            =   570
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmReporteGanancias.frx":01A2
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Imprimir"
         Top             =   165
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.Frame Frame8 
      Height          =   675
      Left            =   75
      TabIndex        =   35
      Top             =   -15
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
         TabIndex        =   36
         Top             =   300
         Width           =   615
      End
   End
   Begin TabDlg.SSTab sstReporte 
      Height          =   3045
      Left            =   90
      TabIndex        =   6
      Top             =   735
      Width           =   7110
      _ExtentX        =   12541
      _ExtentY        =   5371
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Filtros del paciente"
      TabPicture(0)   =   "frmReporteGanancias.frx":0344
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Filtros del inventario"
      TabPicture(1)   =   "frmReporteGanancias.frx":0360
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frmInventario"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Filtro por artículo"
      TabPicture(2)   =   "frmReporteGanancias.frx":037C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1"
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   2370
         Left            =   -74880
         TabIndex        =   30
         Top             =   330
         Width           =   6870
         Begin VB.ListBox lstArticulo 
            Height          =   1425
            Left            =   120
            TabIndex        =   17
            Top             =   765
            Width           =   6645
         End
         Begin VB.TextBox txtArticulo 
            Height          =   285
            Left            =   120
            TabIndex        =   16
            Top             =   465
            Width           =   6645
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Descripción del artículo"
            Height          =   195
            Left            =   120
            TabIndex        =   31
            Top             =   210
            Width           =   1680
         End
      End
      Begin VB.Frame frmInventario 
         Height          =   2370
         Left            =   -74880
         TabIndex        =   26
         Top             =   330
         Width           =   6870
         Begin VB.ComboBox cboDepartamento 
            Height          =   315
            Left            =   3870
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   945
            Width           =   2895
         End
         Begin VB.ComboBox cboTipoSalida 
            Height          =   315
            Left            =   3870
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   570
            Width           =   2895
         End
         Begin VB.ComboBox cboFamilia 
            Height          =   315
            Left            =   975
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   1605
            Width           =   5790
         End
         Begin VB.ComboBox cboSubfamilia 
            Height          =   315
            Left            =   975
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   1965
            Width           =   5790
         End
         Begin VB.OptionButton optMedicamento 
            Caption         =   "Medicamentos"
            Height          =   195
            Left            =   450
            TabIndex        =   9
            Top             =   750
            Width           =   1335
         End
         Begin VB.OptionButton optMaterial 
            Caption         =   "Material"
            Height          =   240
            Left            =   450
            TabIndex        =   10
            Top             =   990
            Width           =   1215
         End
         Begin VB.OptionButton optTodos 
            Caption         =   "Todos"
            Height          =   285
            Left            =   450
            TabIndex        =   11
            Top             =   495
            Value           =   -1  'True
            Width           =   1395
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Departamento"
            Height          =   195
            Left            =   1875
            TabIndex        =   33
            Top             =   1005
            Width           =   1005
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de salida de inventario"
            Height          =   195
            Left            =   1875
            TabIndex        =   32
            Top             =   615
            Width           =   1950
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Familia"
            Height          =   195
            Left            =   150
            TabIndex        =   29
            Top             =   1665
            Width           =   480
         End
         Begin VB.Label lblSubfamilia 
            AutoSize        =   -1  'True
            Caption         =   "Subfamilia"
            Height          =   195
            Left            =   150
            TabIndex        =   28
            Top             =   2025
            Width           =   720
         End
         Begin VB.Label Label8 
            Caption         =   "Tipo de artículo"
            Height          =   195
            Left            =   150
            TabIndex        =   27
            Top             =   285
            Width           =   1440
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2370
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   6870
         Begin VB.Frame Frame7 
            Height          =   495
            Left            =   1605
            TabIndex        =   34
            Top             =   1140
            Width           =   3120
            Begin VB.OptionButton optTipoReporte 
               Caption         =   "Detallado"
               Height          =   195
               Index           =   1
               Left            =   1350
               TabIndex        =   20
               Top             =   195
               Width           =   1215
            End
            Begin VB.OptionButton optTipoReporte 
               Caption         =   "Concentrado"
               Height          =   195
               Index           =   0
               Left            =   75
               TabIndex        =   5
               Top             =   180
               Value           =   -1  'True
               Width           =   1215
            End
         End
         Begin VB.ComboBox cboEmpresa 
            Height          =   315
            Left            =   1605
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   735
            Width           =   5160
         End
         Begin VB.OptionButton optTipoPaciente 
            Caption         =   "Internos"
            Height          =   255
            Index           =   0
            Left            =   2475
            TabIndex        =   1
            Top             =   360
            Width           =   960
         End
         Begin VB.OptionButton optTipoPaciente 
            Caption         =   "Externos"
            Height          =   285
            Index           =   1
            Left            =   3435
            TabIndex        =   2
            Top             =   360
            Width           =   945
         End
         Begin VB.OptionButton optTipoPaciente 
            Caption         =   "Todos"
            Height          =   285
            Index           =   2
            Left            =   1605
            TabIndex        =   3
            Top             =   345
            Value           =   -1  'True
            Width           =   795
         End
         Begin MSMask.MaskEdBox mskInicio 
            Height          =   315
            Left            =   1605
            TabIndex        =   7
            ToolTipText     =   "Fecha de inicio"
            Top             =   1740
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mmm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskFin 
            Height          =   315
            Left            =   3525
            TabIndex        =   8
            ToolTipText     =   "Fecha de fin"
            Top             =   1740
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mmm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Tipo reporte"
            Height          =   195
            Left            =   195
            TabIndex        =   38
            Top             =   1290
            Width           =   855
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Procedencia"
            Height          =   195
            Left            =   195
            TabIndex        =   25
            Top             =   795
            Width           =   900
         End
         Begin VB.Label Label6 
            Caption         =   "Tipo de paciente"
            Height          =   195
            Left            =   195
            TabIndex        =   24
            Top             =   360
            Width           =   1440
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Rango de fechas"
            Height          =   195
            Left            =   195
            TabIndex        =   23
            Top             =   1785
            Width           =   1230
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            Height          =   195
            Left            =   2955
            TabIndex        =   22
            Top             =   1800
            Width           =   465
         End
      End
   End
End
Attribute VB_Name = "frmReporteGanancias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------
'| Nombre del Proyecto      : Caja                                                   -
'| Nombre del Formulario    : frmReporteGanancias                                    -
'-------------------------------------------------------------------------------------
'| Objetivo: Es el reporte de Ganancias de caja
'-------------------------------------------------------------------------------------
'| Análisis y Diseño        : Rodolfo Ramos G.
'| Autor                    : Rodolfo Ramos G.
'| Fecha de Creación        : 18/Sep/2002
'| Modificó                 : Nombre(s)
'| Fecha Terminación        : hoy
'| Fecha última modificación: 19/Sep/2002
'-------------------------------------------------------------------------------------
Option Explicit
Private vgrptReporte As CRAXDRT.Report
Private Sub cboFamilia_Click()
    If cboFamilia.ListIndex > 0 Then
        cboSubfamilia.Enabled = True
        pCargaSubFamilias
    Else
        cboSubfamilia.Enabled = False
    End If
End Sub

Private Sub cboFamilia_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And cboSubfamilia.Enabled Then
       cboSubfamilia.SetFocus
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
        cboDepartamento.ItemData(cboDepartamento.NewIndex) = 0
        cboDepartamento.ListIndex = 0
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboHospital_Click"))
    Unload Me
End Sub

Private Sub cboSubfamilia_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       cmdImprimir.SetFocus
    End If
End Sub

Private Sub cboTipoSalida_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       cboFamilia.SetFocus
    End If
End Sub


Private Sub cmdImprimir_Click()
    pImprime "I"
End Sub

Private Sub cmdVistaPreliminar_Click()
    pImprime "P"
End Sub

Private Sub Form_Activate()
    Me.Refresh
    'Familias
    pCargaFamilias
    'SubFamilias
    pCargaSubFamilias
End Sub

Private Sub Form_Load()
    Dim vlstrsentencia As String
    Dim rs As New ADODB.Recordset
    Dim vlintTipoArticulo As Integer
    Dim lngNumOpcion As Long
    Dim dtmfecha As Date
    
    Me.Icon = frmMenuPrincipal.Icon
    
    Select Case cgstrModulo
    Case "PV"
         lngNumOpcion = 345
    Case "SE"
         lngNumOpcion = 1534
    End Select
    
    pCargaHospital lngNumOpcion
    
    pInstanciaReporte vgrptReporte, "rptGanancias.rpt"
    
    'Empresas
    vlstrsentencia = "select intCveEmpresa, vchDescripcion from ccEmpresa "
    Set rs = frsRegresaRs(vlstrsentencia, adLockReadOnly, adOpenForwardOnly)
    pLlenarCboRs cboEmpresa, rs, 0, 1
    rs.Close
    
    'Tipos de Paciente
    vlstrsentencia = "Select tnyCveTipoPaciente, vchDescripcion from adTipoPaciente order by tnyCveTipoPaciente desc"
    Set rs = frsRegresaRs(vlstrsentencia, adLockReadOnly, adOpenForwardOnly)
    Do While Not rs.EOF
        With cboEmpresa
            .AddItem rs!vchDescripcion, 0
            .ItemData(.NewIndex) = rs!tnyCveTipoPaciente * -1
        End With
        rs.MoveNext
    Loop
    rs.Close
    cboEmpresa.AddItem "<TODOS>", 0
    cboEmpresa.ItemData(0) = 0
    cboEmpresa.ListIndex = 0
    
    'Fechas
    
    dtmfecha = fdtmServerFecha
    mskInicio.Text = dtmfecha
    mskFin.Text = dtmfecha
    
    'Tipos de Salidas de inventario
    With cboTipoSalida
        .AddItem "VENTA AL PUBLICO"
        .AddItem "SALIDA POR VALES"
        .AddItem "CARGOS DIRECTOS"
        .ListIndex = 0
    End With
       
End Sub

Private Sub pCargaSubFamilias()
    Dim rs As New ADODB.Recordset
    Dim vlstrsentencia As String
    
    vlstrsentencia = "select cast((chrCveArtMedicamen || chrcveFamilia+chrCveSubFamilia) AS INTEGER) Clave, vchDescripcion from ivSubFamilia "
    If cboFamilia.ListIndex <> 0 Then
        vlstrsentencia = vlstrsentencia & " where CAST(chrCveArtMedicamen+chrcveFamilia AS INTEGER) = " & Trim(Str(cboFamilia.ItemData(cboFamilia.ListIndex)))
    End If
    
    Set rs = frsRegresaRs(vlstrsentencia, adLockOptimistic, adOpenForwardOnly)
    pLlenarCboRs cboSubfamilia, rs, 0, 1
    rs.Close

    cboSubfamilia.AddItem "<TODAS>", 0
    cboSubfamilia.ItemData(0) = -1 'Para decir que son todas
    cboSubfamilia.ListIndex = 0
End Sub

Private Sub pCargaFamilias()
    Dim rs As New ADODB.Recordset
    Dim vlstrsentencia As String
    
    vlstrsentencia = "select CAST((chrCveArtMedicamen || chrcveFamilia) AS INTEGER) Clave, vchDescripcion from ivFamilia "
    If optMedicamento.Value Then
        vlstrsentencia = vlstrsentencia & " where chrCveArtMedicamen = 1 "
    ElseIf optMaterial.Value Then
        vlstrsentencia = vlstrsentencia & " where chrCveArtMedicamen = 0 "
    End If
    Set rs = frsRegresaRs(vlstrsentencia, adLockOptimistic, adOpenDynamic)
    pLlenarCboRs cboFamilia, rs, 0, 1
    rs.Close
    cboFamilia.AddItem "<TODAS>", 0
    cboFamilia.ItemData(0) = -1 'Para decir que son todas
    cboFamilia.ListIndex = 0
End Sub

Private Sub cboEmpresa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        optTipoReporte(0).SetFocus
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub mskFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdVistaPreliminar.SetFocus
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

Private Sub optMaterial_Click()
    pCargaFamilias
End Sub

Private Sub optMaterial_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cboTipoSalida.SetFocus
    End If
End Sub

Private Sub optMedicamento_Click()
    pCargaFamilias
End Sub

Private Sub optMedicamento_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cboTipoSalida.SetFocus
    End If
End Sub

Private Sub OptTipoPaciente_Click(Index As Integer)
    optTipoPaciente_KeyDown Index, vbKeyReturn, 0
End Sub

Private Sub optTipoPaciente_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cboEmpresa.SetFocus
    End If
End Sub

Private Sub optTipoReporte_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        mskInicio.SetFocus
    End If

End Sub

Private Sub optTodos_Click()
    pCargaFamilias
End Sub

Private Sub optTodos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cboTipoSalida.SetFocus
    End If
End Sub

Private Sub sstReporte_Click(PreviousTab As Integer)
    If sstReporte.Tab = 2 Then
        pEnfocaTextBox txtArticulo
    End If
End Sub

Private Sub txtArticulo_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyReturn Or KeyCode = vbKeyRight Or KeyCode = vbKeyDown) And lstArticulo.Enabled Then
        lstArticulo.SetFocus
    End If
End Sub

Private Sub txtArticulo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtArticulo_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim vlstrsentencia As String
    vlstrsentencia = "Select intIdArticulo Clave, vchNombreComercial Descripcion from ivArticulo"
    PSuperBusqueda txtArticulo, vlstrsentencia, lstArticulo, "vchNombreComercial", 20, , "vchNombreComercial"
End Sub

Sub pImprime(pstrDestino As String)

    Dim vlstrSPParam As String
    Dim rs As New ADODB.Recordset
    Dim vlstrsentencia As String
    Dim vlstrFamilia As String
    Dim vlstrSubFamilia As String
    Dim vlstrTipoArticulo As String
    Dim vlstrTipoSalida As String
    Dim vlstrArticulo As String
    Dim alstrParametros(0) As String
    
    If Not IsDate(mskInicio.Text) Then
        MsgBox SIHOMsg(29), vbInformation, "Mensaje"
        pEnfocaMkTexto mskInicio
        Exit Sub
    ElseIf Not IsDate(mskFin.Text) Then
        MsgBox SIHOMsg(29), vbInformation, "Mensaje"
        pEnfocaMkTexto mskFin
        Exit Sub
    End If

    'Cual Tipo de Articulo?
    If optMedicamento.Value Then
        vlstrTipoArticulo = "1"
    ElseIf optMaterial.Value Then
        vlstrTipoArticulo = "0"
    Else
        vlstrTipoArticulo = "-1"
    End If

    'Cual familia es??
    If cboFamilia.ListIndex = 0 Then
        vlstrFamilia = "-1"
    Else
        vlstrsentencia = "SELECT chrCveArtMedicamen TipoArticulo, chrCveFamilia Familia from ivFamilia " & _
                " WHERE cast(chrCveArtMedicamen+chrcveFamilia as integer) = " & Trim(Str(cboFamilia.ItemData(cboFamilia.ListIndex)))
        Set rs = frsRegresaRs(vlstrsentencia, adLockReadOnly, adOpenForwardOnly)
        vlstrFamilia = rs!Familia
        vlstrTipoArticulo = rs!TipoArticulo
        rs.Close
    End If
    'Cual SubFamilia es?
    If cboSubfamilia.ListIndex = 0 Then
        vlstrSubFamilia = "-1"
    Else
        vlstrsentencia = "SELECT chrCveSubFamilia SubFamilia from ivSubFamilia " & _
                " WHERE cast(chrCveArtMedicamen+chrcveFamilia+chrCveSubFamilia as integer) = " & Trim(Str(cboSubfamilia.ItemData(cboSubfamilia.ListIndex)))
        Set rs = frsRegresaRs(vlstrsentencia, adLockReadOnly, adOpenForwardOnly)
        vlstrSubFamilia = rs!SubFamilia
        rs.Close
    End If
    'Tipo de Salida de inventario
    If cboTipoSalida.ListIndex = 0 Then
        vlstrTipoSalida = "SVP"
    ElseIf cboTipoSalida.ListIndex = 1 Then
        vlstrTipoSalida = "SCP"
    Else
        vlstrTipoSalida = "SCD" 'Aun no esta grabando esta clave los cargos directos
    End If
    'Cual Tipo de Articulo
    If lstArticulo.ListCount = 0 Then
        vlstrArticulo = "*"
    Else
        vlstrArticulo = Trim(Str(lstArticulo.ItemData(lstArticulo.ListIndex)))
    End If
    vlstrSPParam = vlstrTipoArticulo & "|" & _
    vlstrFamilia & "|" & _
    vlstrSubFamilia & "|" & _
    vlstrArticulo & "|" & _
    IIf(optTipoPaciente(0).Value, "I", IIf(optTipoPaciente(1).Value, "E", "A")) & "|" & _
    cboEmpresa.ItemData(cboEmpresa.ListIndex) & "|" & _
    cboDepartamento.ItemData(cboDepartamento.ListIndex) & "|" & _
    fstrFechaSQL(mskInicio.Text, " 00:00:00") & "|" & _
    fstrFechaSQL(mskFin.Text, " 23:59:59") & "|" & _
    IIf(optTipoReporte(0).Value, "G", "D") & "|" & _
    vlstrTipoSalida & "|" & _
    Str(cboHospital.ItemData(cboHospital.ListIndex))
    
    Set rs = frsEjecuta_SP(vlstrSPParam, "sp_pvSelReporteGanancia")
    If rs.EOF Then
        MsgBox SIHOMsg(13), vbInformation, "Mensaje"
    Else
        vgrptReporte.DiscardSavedData
        alstrParametros(0) = "NombreHospital; " & Trim(cboHospital.List(cboHospital.ListIndex))
        pCargaParameterFields alstrParametros, vgrptReporte
        pImprimeReporte vgrptReporte, rs, pstrDestino, "Reporte de ganancias"
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


