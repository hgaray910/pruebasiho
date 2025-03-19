VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmReporteProductividadMedicos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Productividad de médicos"
   ClientHeight    =   8790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6765
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8790
   ScaleWidth      =   6765
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame8 
      Height          =   675
      Left            =   75
      TabIndex        =   58
      Top             =   2250
      Width           =   6615
      Begin VB.ComboBox cboProcedencia 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   7
         ToolTipText     =   "Selección de la procedencia del paciente"
         Top             =   210
         Width           =   5265
      End
      Begin VB.Label Label9 
         Caption         =   "Procedencia"
         Height          =   195
         Left            =   120
         TabIndex        =   59
         Top             =   240
         Width           =   915
      End
   End
   Begin VB.ListBox lstBusqueda 
      Height          =   840
      Left            =   1320
      TabIndex        =   53
      Top             =   9720
      Visible         =   0   'False
      Width           =   5130
   End
   Begin VB.Frame Frame3 
      Caption         =   "Rango de fechas"
      Height          =   675
      Left            =   75
      TabIndex        =   42
      Top             =   1560
      Width           =   6615
      Begin MSMask.MaskEdBox mskFecIni 
         Height          =   315
         Left            =   1200
         TabIndex        =   5
         ToolTipText     =   "Fecha inicial"
         Top             =   240
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
         Left            =   3600
         TabIndex        =   6
         ToolTipText     =   "Fecha final"
         Top             =   240
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
         Left            =   3000
         TabIndex        =   44
         Top             =   300
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Left            =   120
         TabIndex        =   43
         Top             =   300
         Width           =   465
      End
   End
   Begin TabDlg.SSTab sstTipoReporte 
      Height          =   4995
      Left            =   75
      TabIndex        =   41
      Top             =   2940
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   8811
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Productividad general  "
      TabPicture(0)   =   "frmReporteProductividadMedicos.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame10"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Productividad en venta al público"
      TabPicture(1)   =   "frmReporteProductividadMedicos.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame4 
         Height          =   4635
         Left            =   -74880
         TabIndex        =   48
         Top             =   360
         Width           =   6375
         Begin VB.ComboBox cboDepartamento 
            Height          =   315
            Left            =   1125
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   23
            ToolTipText     =   "Departamento que realizó la venta"
            Top             =   510
            Width           =   5130
         End
         Begin VB.TextBox txtDescripcion 
            Height          =   315
            Left            =   1125
            TabIndex        =   31
            ToolTipText     =   "Descripción del artículo u otro concepto"
            Top             =   2985
            Width           =   5130
         End
         Begin VB.Frame fraSubClasificacion 
            Height          =   420
            Left            =   1110
            TabIndex        =   54
            Top             =   2250
            Width           =   5130
            Begin VB.OptionButton optSubArticulos 
               Caption         =   "Medicamentos"
               Height          =   255
               Index           =   2
               Left            =   3360
               TabIndex        =   30
               Top             =   120
               Width           =   1455
            End
            Begin VB.OptionButton optSubArticulos 
               Caption         =   "Artículos"
               Height          =   255
               Index           =   1
               Left            =   1600
               TabIndex        =   29
               Top             =   120
               Width           =   975
            End
            Begin VB.OptionButton optSubArticulos 
               Caption         =   "Todos"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   28
               Top             =   120
               Width           =   855
            End
         End
         Begin VB.CheckBox chkDetallado 
            Caption         =   "Detallado"
            Height          =   255
            Left            =   5280
            TabIndex        =   34
            Top             =   3840
            Width           =   975
         End
         Begin VB.Frame Frame9 
            Caption         =   "Agrupar por "
            Height          =   615
            Left            =   1110
            TabIndex        =   51
            Top             =   3630
            Width           =   2655
            Begin VB.OptionButton optAgrupar 
               Caption         =   "Médico"
               Height          =   255
               Index           =   1
               Left            =   1560
               TabIndex        =   33
               Top             =   240
               Width           =   975
            End
            Begin VB.OptionButton optAgrupar 
               Caption         =   "Artículo"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   32
               Top             =   240
               Width           =   975
            End
         End
         Begin VB.Frame Frame7 
            Height          =   420
            Left            =   1125
            TabIndex        =   50
            Top             =   1635
            Width           =   5130
            Begin VB.OptionButton optClasificacion 
               Caption         =   "Otros conceptos"
               Height          =   255
               Index           =   2
               Left            =   3360
               TabIndex        =   27
               Top             =   120
               Width           =   1575
            End
            Begin VB.OptionButton optClasificacion 
               Caption         =   "Artículos"
               Height          =   255
               Index           =   1
               Left            =   1600
               TabIndex        =   26
               Top             =   120
               Width           =   975
            End
            Begin VB.OptionButton optClasificacion 
               Caption         =   "Todos"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   25
               Top             =   120
               Width           =   855
            End
         End
         Begin VB.ComboBox cboLocalizacion 
            Height          =   315
            Left            =   1125
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   24
            ToolTipText     =   "Localización"
            Top             =   1110
            Width           =   5130
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Departamento"
            Height          =   195
            Left            =   90
            TabIndex        =   55
            Top             =   570
            Width           =   1005
         End
         Begin VB.Label Label8 
            Caption         =   "Descripción"
            Height          =   195
            Left            =   90
            TabIndex        =   52
            Top             =   3045
            Width           =   855
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Localización"
            Height          =   195
            Left            =   90
            TabIndex        =   49
            Top             =   1170
            Width           =   885
         End
      End
      Begin VB.Frame Frame10 
         Height          =   4515
         Left            =   120
         TabIndex        =   45
         Top             =   330
         Width           =   6375
         Begin VB.Frame Frame11 
            Caption         =   "Agrupación"
            Height          =   780
            Left            =   3240
            TabIndex        =   60
            Top             =   2420
            Width           =   3000
            Begin VB.OptionButton optAgrupacion 
               Caption         =   "Por procedencia"
               Height          =   195
               Index           =   1
               Left            =   1440
               TabIndex        =   17
               ToolTipText     =   "Agrupación por procedencia"
               Top             =   360
               Width           =   1480
            End
            Begin VB.OptionButton optAgrupacion 
               Caption         =   "Por médico"
               Height          =   195
               Index           =   0
               Left            =   240
               TabIndex        =   16
               ToolTipText     =   "Agrupado por médico"
               Top             =   360
               Width           =   1215
            End
         End
         Begin VB.ComboBox cboDeptoCredito 
            Height          =   315
            Left            =   2040
            TabIndex        =   20
            ToolTipText     =   "Selección de departamentos del crédito"
            Top             =   4050
            Width           =   4215
         End
         Begin VB.CheckBox chkMostrarCreditosMedico 
            Caption         =   "Mostrar los créditos otorgados al médico"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            ToolTipText     =   "Mostrar los créditos otorgados al médico"
            Top             =   3600
            Value           =   1  'Checked
            Width           =   3135
         End
         Begin VB.CheckBox chkProcedencia 
            Caption         =   "Agrupar los pacientes por tipo de paciente / Empresa"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   18
            ToolTipText     =   "Agrupar los pacientes por tipo de paciente / Empresa"
            Top             =   3240
            Width           =   4155
         End
         Begin VB.Frame fraAreaProductividad 
            Caption         =   "Área de productividad"
            Height          =   1455
            Left            =   120
            TabIndex        =   56
            Top             =   900
            Width           =   6135
            Begin VB.CheckBox chkPacientesExtAtendidos 
               Caption         =   "Pacientes externos atendidos"
               Height          =   255
               Left            =   120
               TabIndex        =   9
               ToolTipText     =   "Pacientes externos atendidos"
               Top             =   360
               Value           =   1  'Checked
               Width           =   2535
            End
            Begin VB.CheckBox chkPacientesrefLab 
               Caption         =   "Pacientes referidos a laboratorio"
               Height          =   255
               Left            =   2880
               TabIndex        =   13
               ToolTipText     =   "Pacientes referidos a laboratorio"
               Top             =   720
               Value           =   1  'Checked
               Width           =   2655
            End
            Begin VB.CheckBox chkpacientesRefImagen 
               Caption         =   "Pacientes referidos a imagenología"
               Height          =   255
               Left            =   2880
               TabIndex        =   12
               ToolTipText     =   "Pacientes referidos a imagenología"
               Top             =   360
               Value           =   1  'Checked
               Width           =   2895
            End
            Begin VB.CheckBox chkPacientesRefFarmacia 
               Caption         =   "Pacientes referidos a farmacia"
               Height          =   255
               Left            =   120
               TabIndex        =   11
               ToolTipText     =   "Pacientes referidos a farmacia"
               Top             =   1080
               Value           =   1  'Checked
               Width           =   2535
            End
            Begin VB.CheckBox chkPacientesIngAlhospital 
               Caption         =   "Pacientes ingresados al hospital"
               Height          =   255
               Left            =   120
               TabIndex        =   10
               ToolTipText     =   "Pacientes ingresados al hospital"
               Top             =   720
               Value           =   1  'Checked
               Width           =   2775
            End
         End
         Begin VB.ComboBox cboEmpresa 
            Height          =   315
            Left            =   120
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   8
            ToolTipText     =   "Selección del tipo de paciente / Empresa"
            Top             =   480
            Width           =   6135
         End
         Begin VB.Frame Frame2 
            Caption         =   "Presentación"
            Height          =   780
            Left            =   120
            TabIndex        =   46
            Top             =   2420
            Width           =   3000
            Begin VB.OptionButton optTipoReporte 
               Caption         =   "Detallado"
               Height          =   195
               Index           =   1
               Left            =   1680
               TabIndex        =   15
               ToolTipText     =   "Detallado"
               Top             =   360
               Width           =   1095
            End
            Begin VB.OptionButton optTipoReporte 
               Caption         =   "Concentrado"
               Height          =   195
               Index           =   0
               Left            =   240
               TabIndex        =   14
               ToolTipText     =   "Concentrado"
               Top             =   360
               Value           =   -1  'True
               Width           =   1335
            End
         End
         Begin VB.Label lbDeptoCredito 
            Caption         =   "Departamento del crédito"
            Height          =   255
            Left            =   120
            TabIndex        =   57
            Top             =   4080
            Width           =   1935
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de paciente / Empresa"
            Height          =   195
            Left            =   90
            TabIndex        =   47
            Top             =   210
            Width           =   1980
         End
      End
   End
   Begin VB.Frame Frame5 
      Height          =   675
      Left            =   75
      TabIndex        =   39
      Top             =   -45
      Width           =   6615
      Begin VB.ComboBox cboHospital 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Seleccione la empresa"
         Top             =   220
         Width           =   5265
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Empresa"
         Height          =   195
         Left            =   120
         TabIndex        =   40
         Top             =   280
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Height          =   255
      Left            =   2550
      TabIndex        =   38
      Top             =   9735
      Width           =   855
   End
   Begin VB.Frame Frame6 
      Height          =   735
      Left            =   2800
      TabIndex        =   36
      Top             =   7965
      Width           =   1165
      Begin VB.CommandButton cmdPreview 
         Height          =   495
         Left            =   70
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmReporteProductividadMedicos.frx":0038
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Vista previa"
         Top             =   165
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdPrint 
         Height          =   495
         Left            =   600
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmReporteProductividadMedicos.frx":043B
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Imprimir"
         Top             =   165
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Height          =   900
      Left            =   75
      TabIndex        =   35
      Top             =   600
      Width           =   6615
      Begin VB.ComboBox cboMedicos 
         Height          =   315
         Left            =   1200
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "Selección del médico"
         Top             =   480
         Width           =   5265
      End
      Begin VB.OptionButton optInternos 
         Caption         =   "Staff"
         Height          =   200
         Left            =   2520
         TabIndex        =   2
         ToolTipText     =   "Tipo de médico"
         Top             =   210
         Width           =   930
      End
      Begin VB.OptionButton optExternos 
         Caption         =   "Externos"
         Height          =   200
         Left            =   3840
         TabIndex        =   3
         ToolTipText     =   "Tipo de médico"
         Top             =   210
         Width           =   960
      End
      Begin VB.OptionButton optAmbos 
         Caption         =   "Todos"
         Height          =   200
         Left            =   1200
         TabIndex        =   1
         ToolTipText     =   "Tipo de médico"
         Top             =   210
         Value           =   -1  'True
         Width           =   795
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Médicos"
         Height          =   195
         Left            =   150
         TabIndex        =   37
         Top             =   215
         Width           =   600
      End
   End
End
Attribute VB_Name = "frmReporteProductividadMedicos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim vlstrGrupo As String
Dim dtmFecha1 As Date
Dim lintAlmacenVenta As Integer
Dim llngClaveCargo As Long 'clave del articulo o del otro concepto

Private vgrptReporte As CRAXDRT.Report

Private Sub pCargaProcedencia()
On Error GoTo NotificaError
    
    Dim vlstrSentencia As String
    Dim rsProcedencia As New ADODB.Recordset

    cboProcedencia.Clear
    cboProcedencia.AddItem "<TODOS>", 0
    cboProcedencia.ItemData(cboProcedencia.newIndex) = 0
    cboProcedencia.ListIndex = 0
    
    vlstrSentencia = "SELECT * FROM AdProcedencia ORDER BY VCHDESCRIPCION"
    Set rsProcedencia = frsRegresaRs(vlstrSentencia)
    
    Do While Not rsProcedencia.EOF
        cboProcedencia.AddItem rsProcedencia!VCHDESCRIPCION
        cboProcedencia.ItemData(cboProcedencia.newIndex) = rsProcedencia!intcveprocedencia
        rsProcedencia.MoveNext
    Loop
    rsProcedencia.Close
    cboProcedencia.ListIndex = 1
    
    cboProcedencia.AddItem "SIN PROCEDENCIA"
    cboProcedencia.ListIndex = 0

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCargaProcedencia"))
End Sub
Private Sub cboDepartamento_Click()
    If cboDepartamento.ListIndex = 0 Then
        If cboLocalizacion.ListCount > 0 Then cboLocalizacion.ListIndex = 0
        cboLocalizacion.Enabled = False
        lintAlmacenVenta = -1
    Else
        pCargaLocalizacion
        cboLocalizacion.Enabled = True
    End If
End Sub

Private Sub cboDepartamento_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
                
        SendKeys vbTab
        
    End If
    
End Sub

Private Sub cboDeptoCredito_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
    
        SendKeys vbTab
        
    End If
    
End Sub


Private Sub cboEmpresa_Click()
'    If cboEmpresa.ItemData(cboEmpresa.ListIndex) = 0 And cboEmpresa.ListIndex > 0 Then
'        cboEmpresa.ListIndex = cboEmpresa.ListIndex + 1
'    End If
End Sub

Private Sub cboEmpresa_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
                
        SendKeys vbTab
        
    End If
    
End Sub


Private Sub cboProcedencia_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
                
        SendKeys vbTab
        
    End If
    
End Sub


Private Sub cboHospital_Click()
     pCargaDepartamentos
     If cboDepartamento.ListCount > 0 Then cboDepartamento.ListIndex = 0
End Sub

Private Sub cboHospital_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
    
        If optAmbos.Value = True Then optAmbos.SetFocus
        If optInternos.Value = True Then optInternos.SetFocus
        If optExternos.Value = True Then optExternos.SetFocus
        
    End If

End Sub



Private Sub cboLocalizacion_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
                
        SendKeys vbTab
        
    End If
    
End Sub

Private Sub cboMedicos_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
    
        SendKeys vbTab
        
    End If
    
End Sub



Private Sub chkDetallado_Click()
    If chkDetallado.Value = 0 Then
        Frame2.Enabled = False
        optTipoReporte(0).Enabled = False
        optTipoReporte(1).Enabled = False
        
    Else
        optTipoReporte(0).Enabled = True
        optTipoReporte(1).Enabled = True
        
    End If
End Sub

Private Sub chkDetallado_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
                
        cmdPreview.SetFocus
        
    End If
    
End Sub

Private Sub chkMostrarCreditosMedico_Click()


    If chkMostrarCreditosMedico.Value = 1 Then
    
        cboDeptoCredito.Enabled = True
        lbDeptoCredito.Enabled = True
                                        
    Else
    
        cboDeptoCredito.Enabled = False
        lbDeptoCredito.Enabled = False
        
    End If
    
End Sub

Private Sub chkMostrarCreditosMedico_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
    
        SendKeys vbTab
        
    End If
    
End Sub


Private Sub chkPacientesExtAtendidos_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
    
        SendKeys vbTab
        
    End If
    
End Sub


Private Sub chkPacientesIngAlhospital_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
    
        SendKeys vbTab
        
    End If
    
End Sub


Private Sub chkPacientesRefFarmacia_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
    
        SendKeys vbTab
        
    End If
    
End Sub

Private Sub chkpacientesRefImagen_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
    
        SendKeys vbTab
        
    End If
    
    
End Sub


Private Sub chkPacientesrefLab_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
    
        SendKeys vbTab
        
    End If
    
End Sub


Private Sub chkProcedencia_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
    
        SendKeys vbTab
        
    End If
    
    
End Sub


Private Sub cmdPreview_Click()
    On Error GoTo NotificaError
    
    pReporte "P"
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdPreview_Click"))
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo NotificaError
    pReporte "I"
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdPrint_Click"))
End Sub

Private Sub cmdSalir_Click()
    If lstBusqueda.Visible Then
        txtDescripcion.Text = ""
        lstBusqueda.Visible = False
        txtDescripcion.SetFocus
    Else
        Unload Me
    End If
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
'    If KeyCode = vbKeyReturn Then
'            If frmReporteProductividadMedicos.ActiveControl.Name = "mskFecFin" Then
'                If sstTipoReporte.Tab = 0 Then
'                    cboEmpresa.SetFocus
'                Else
'                    If cboDepartamento.Enabled Then
'                        cboDepartamento.SetFocus
'                    Else
'                        cboLocalizacion.SetFocus
'                    End If
'                End If
'            ElseIf frmReporteProductividadMedicos.ActiveControl.Name = "chkDetallado" Then
'                cmdPreview.SetFocus
'            ElseIf frmReporteProductividadMedicos.ActiveControl.Name = "txtDescripcion" Then
'                If lstBusqueda.Visible Then
'                    lstBusqueda.SetFocus
'                Else
'                    optAgrupar(0).SetFocus
'                End If
'            ElseIf frmReporteProductividadMedicos.ActiveControl.Name = "lstBusqueda" Then
'                optAgrupar(0).SetFocus
'            Else
'                SendKeys vbTab
'            End If
'    End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_KeyDown"))
End Sub

Private Sub Form_Load()
    On Error GoTo NotificaError
    Dim rs As New ADODB.Recordset
    Dim vlstrx As String
    Dim lngNumOpcion As Long
        
    Me.Icon = frmMenuPrincipal.Icon
    
    Select Case cgstrModulo
    Case "PV"
         lngNumOpcion = 329
    Case "SE"
         lngNumOpcion = 2008
    End Select
    
    pCargaHospital lngNumOpcion
    
    dtmFecha1 = fdtmServerFecha
    
    vlstrx = "select intcvemedico, vchApellidoPaterno || ' ' || vchApellidoMaterno || ' ' || vchNombre as Nombre from HoMedico order by vchApellidoPaterno, vchApellidoMaterno, vchNombre"
    Set rs = frsRegresaRs(vlstrx)
    pLlenarCboRs cboMedicos, rs, 0, 1, 3
    cboMedicos.ListIndex = 0
    vlstrx = "select tnyCveTipoPaciente*-1 Clave, vchDescripcion Descripcion from AdTipoPaciente "
    vlstrx = vlstrx & "order by Descripcion"
    cboEmpresa.AddItem "<TODOS>"
    cboEmpresa.ItemData(cboEmpresa.newIndex) = 0
    Set rs = frsRegresaRs(vlstrx)
    If Not rs.EOF Then
        cboEmpresa.AddItem "------------------------------------------------------Tipos de Paciente"
        cboEmpresa.ItemData(cboEmpresa.newIndex) = 0
        Do Until rs.EOF
            cboEmpresa.AddItem rs!Descripcion
            cboEmpresa.ItemData(cboEmpresa.newIndex) = rs!Clave
            rs.MoveNext
        Loop
    End If
    
    pCargaProcedencia
    
    vlstrx = "select intCveEmpresa Clave, vchDescripcion Descripcion from CcEmpresa "
    vlstrx = vlstrx & "order by Descripcion"
    Set rs = frsRegresaRs(vlstrx)
    If Not rs.EOF Then
        cboEmpresa.AddItem "--------------------------------------------------------------------Empresas"
        cboEmpresa.ItemData(cboEmpresa.newIndex) = 0
        Do Until rs.EOF
            cboEmpresa.AddItem rs!Descripcion
            cboEmpresa.ItemData(cboEmpresa.newIndex) = rs!Clave
            rs.MoveNext
        Loop
    End If
    cboEmpresa.ListIndex = 0
    
    If cboEmpresa.List(cboEmpresa.ListIndex) <> "<TODOS>" Then
        
        For i = 1 To cboEmpresa.ListCount
        
            If cboEmpresa.List(i) = "<TODOS>" Then
            
                cboEmpresa.ListIndex = i
                i = cboEmpresa.ListCount
            
            End If
            
        Next i
    
    End If
    
    mskFecIni.Mask = ""
    mskFecIni.Text = dtmFecha1
    mskFecIni.Mask = "##/##/####"
    mskFecFin.Mask = ""
    mskFecFin.Text = dtmFecha1
    mskFecFin.Mask = "##/##/####"
    
    pCargaDepartamentos
    cboDepartamento.ListIndex = flngLocalizaCbo(cboDepartamento, Str(vgintNumeroDepartamento))
    cboDepartamento.Enabled = fblnRevisaPermiso(vglngNumeroLogin, lngNumOpcion, "C")
    pCargaLocalizacion
                   
    optClasificacion(0).Value = True
    optSubArticulos(0).Value = True
    optAgrupar(0).Value = True
        
    sstTipoReporte.Tab = 0
    
    pCargaDepartamentosCredito
    
    If fstrValorPorDefecto(vglngNumeroLogin, 1) = "P" Then
        optAgrupacion(1).Value = True
    Else
        optAgrupacion(0).Value = True
    End If
    
        
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_Load"))
End Sub
Private Sub pCargaDepartamentos()
On Error GoTo NotificaError
    Dim rs As New ADODB.Recordset
    cboDepartamento.Clear
    vgstrParametrosSP = "-1|1|*|" & CStr(cboHospital.ItemData(cboHospital.ListIndex))
    Set rs = frsEjecuta_SP(vgstrParametrosSP, "Sp_Gnseldepartamento")
    If rs.RecordCount <> 0 Then
        pLlenarCboRs cboDepartamento, rs, 0, 1
    End If
    cboDepartamento.AddItem "<TODOS>", 0
    cboDepartamento.ItemData(cboDepartamento.newIndex) = -1
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargaDepartamentos"))
End Sub
Private Sub pCargaLocalizacion()
On Error GoTo NotificaError
    Dim rs As New ADODB.Recordset
    Dim rsAlmacen As New ADODB.Recordset
    Dim lstrSentencia As String
    
    lintAlmacenVenta = -1
    lstrSentencia = "select smicvedepartamento from nodepartamento where nodepartamento.chrClasificacion = 'A' and smicvedepartamento=" & cboDepartamento.ItemData(cboDepartamento.ListIndex)
    Set rs = frsRegresaRs(lstrSentencia, adLockOptimistic, adOpenDynamic)
    If Not rs.RecordCount = 0 Then
        lintAlmacenVenta = cboDepartamento.ItemData(cboDepartamento.ListIndex)
    Else
        lstrSentencia = "select intnumalmacen from pvalmacenes where intnumdepartamento =" & cboDepartamento.ItemData(cboDepartamento.ListIndex)
        Set rsAlmacen = frsRegresaRs(lstrSentencia, adLockOptimistic, adOpenDynamic)
        If rsAlmacen.RecordCount > 0 Then
            lintAlmacenVenta = rsAlmacen!intnumalmacen
        End If
    End If
    
    cboLocalizacion.Clear
    Set rs = frsEjecuta_SP(CStr(lintAlmacenVenta), "SP_IVSELLOCALIZACION")
    If rs.RecordCount > 0 Then
        pLlenarCboRs cboLocalizacion, rs, 0, 1
    End If
    cboLocalizacion.AddItem "<TODAS>", 0
    cboLocalizacion.ListIndex = 0
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargaLocalizacion"))
End Sub

Private Sub Form_Unload(Cancel As Integer)
    pActualizaValorPorDefecto vglngNumeroLogin, 1, IIf(optAgrupacion(0).Value, "M", "P")
End Sub

Private Sub lstBusqueda_DblClick()
    lstBusqueda_KeyDown 13, 0
End Sub

Private Sub lstBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtDescripcion.Text = lstBusqueda.List(lstBusqueda.ListIndex)
        llngClaveCargo = lstBusqueda.ItemData(lstBusqueda.ListIndex)
        lstBusqueda.Visible = False
        cboLocalizacion.ListIndex = 0
        optSubArticulos(0).Value = True
        optAgrupar(0).SetFocus
    End If
End Sub

Private Sub mskFecFin_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
    
        SendKeys vbTab
        
    End If
    
End Sub

Private Sub mskFecIni_GotFocus()
    On Error GoTo NotificaError
    pSelMkTexto mskFecIni
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFecIni_GotFocus"))
End Sub
Private Sub mskFecFin_GotFocus()
    On Error GoTo NotificaError
    pSelMkTexto mskFecFin
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFecFin_GotFocus"))
End Sub

Private Sub mskFecIni_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
    
        SendKeys vbTab
        
    End If
    
End Sub

Private Sub mskFecIni_LostFocus()
    On Error GoTo NotificaError
    If Trim(mskFecIni.ClipText) = "" Then
        mskFecIni.Mask = ""
        mskFecIni.Text = dtmFecha1
        mskFecIni.Mask = "##/##/####"
    Else
        If Not IsDate(mskFecIni.Text) Then
            mskFecIni.Mask = ""
            mskFecIni.Text = dtmFecha1
            mskFecIni.Mask = "##/##/####"
        End If
    End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFecIni_LostFocus"))
End Sub
Private Sub mskFecFin_LostFocus()
    On Error GoTo NotificaError
    If Trim(mskFecFin.ClipText) = "" Then
        mskFecFin.Mask = ""
        mskFecFin.Text = dtmFecha1
        mskFecFin.Mask = "##/##/####"
    Else
        If Not IsDate(mskFecFin.Text) Then
            mskFecFin.Mask = ""
            mskFecFin.Text = dtmFecha1
            mskFecFin.Mask = "##/##/####"
        End If
    End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFecFin_LostFocus"))
End Sub

Private Sub pReporte(vlstrDestino As String)
    On Error GoTo NotificaError
    Dim fecha1 As Date
    Dim fecha2 As Date
    Dim alstrParametros(5) As String
    Dim rsReporte As ADODB.Recordset
   
    
    Dim vlstrSPPar As String
    Dim rs As New ADODB.Recordset
    
    If sstTipoReporte.Tab = 0 Then
        If optInternos.Value = True Then
            vlstrSPPar = "I"
        End If
        If optExternos.Value = True Then
            vlstrSPPar = "E"
        End If
        If optAmbos.Value = True Then
            vlstrSPPar = "*"
        End If
    
        
        vlstrSPPar = vlstrSPPar & "|" & cboMedicos.ItemData(cboMedicos.ListIndex)
        If Val(cboEmpresa.ItemData(cboEmpresa.ListIndex)) < 0 Then
            vlstrSPPar = vlstrSPPar & "|0|" & Abs(Val(cboEmpresa.ItemData(cboEmpresa.ListIndex)))
        Else
            If Val(cboEmpresa.ItemData(cboEmpresa.ListIndex)) > 0 Then
                vlstrSPPar = vlstrSPPar & "|" & cboEmpresa.ItemData(cboEmpresa.ListIndex) & "|0"
            Else
                vlstrSPPar = vlstrSPPar & "|0|0"
            End If
        End If
        vlstrSPPar = vlstrSPPar & "|" & fstrFechaSQL(mskFecIni.Text, "00:00:00") & "|" & fstrFechaSQL(mskFecFin.Text, "23:59:59")
        
        'Tipo de Reporte
        If optTipoReporte(0).Value = True Then
        
                vlstrSPPar = vlstrSPPar & "|0|" & Str(cboHospital.ItemData(cboHospital.ListIndex)) _
                & "|" & IIf(chkPacientesExtAtendidos.Value = 1, 1, 0) & "|" & IIf(chkPacientesIngAlhospital.Value = 1, 1, 0) _
                & "|" & IIf(chkPacientesRefFarmacia.Value = 1, 1, 0) & "|" & IIf(chkpacientesRefImagen.Value = 1, 1, 0) _
                & "|" & IIf(chkPacientesrefLab.Value = 1, 1, 0) & "|" & Str(cboDeptoCredito.ItemData(cboDeptoCredito.ListIndex)) _
                & "|" & IIf(cboProcedencia.ListIndex = 0, -1, cboProcedencia.ItemData(cboProcedencia.ListIndex)) _
                & "|" & IIf(optAgrupacion(0).Value, 0, 1)
                Set rs = frsEjecuta_SP(vlstrSPPar, "sp_PvRptProductividadMedicos")
                pInstanciaReporte vgrptReporte, "rptproductividadmedicossum.rpt"
                
        Else
                
                vlstrSPPar = vlstrSPPar & "|1|" & Str(cboHospital.ItemData(cboHospital.ListIndex)) _
                & "|" & IIf(chkPacientesExtAtendidos.Value = 1, 1, 0) & "|" & IIf(chkPacientesIngAlhospital.Value = 1, 1, 0) _
                & "|" & IIf(chkPacientesRefFarmacia.Value = 1, 1, 0) & "|" & IIf(chkpacientesRefImagen.Value = 1, 1, 0) _
                & "|" & IIf(chkPacientesrefLab.Value = 1, 1, 0) & "|" & Str(cboDeptoCredito.ItemData(cboDeptoCredito.ListIndex)) _
                & "|" & IIf(cboProcedencia.ListIndex = 0, -1, cboProcedencia.ItemData(cboProcedencia.ListIndex)) _
                & "|" & IIf(optAgrupacion(0).Value, 0, 1)
                Set rs = frsEjecuta_SP(vlstrSPPar, "sp_PvRptProductividadMedicos")
                pInstanciaReporte vgrptReporte, "rptproductividadmedicosdet.rpt"
                
       End If
            
        
        If rs.EOF Then
            MsgBox SIHOMsg(13), vbInformation, "Mensaje"
        Else
            vgrptReporte.DiscardSavedData
            
            alstrParametros(0) = "NombreHospital;" & cboHospital.List(cboHospital.ListIndex)
            alstrParametros(1) = "FechaIni;" & UCase(Format(mskFecIni.Text, "dd/mmm/yyyy"))
            alstrParametros(2) = "FechaFin;" & UCase(Format(mskFecFin.Text, "dd/mmm/yyyy"))
            alstrParametros(3) = "MostrarCreditos;" & IIf(chkMostrarCreditosMedico.Value = 1, "True", "False") & ";BOOLEAN"
            alstrParametros(4) = "Procedencia;" & IIf(chkProcedencia.Value = 1, "True", "False") & ";BOOLEAN"
            pCargaParameterFields alstrParametros, vgrptReporte
    
            pImprimeReporte vgrptReporte, rs, vlstrDestino, "Productividad de médicos"
                        
                        
        End If
        rs.Close
    Else
    'Reporte productividad de médicos por artículo
        If txtDescripcion <> "" Then   'validacion del filtro por articulo u otro concepto
            cboLocalizacion.ListIndex = 0
            optSubArticulos(0).Value = True
            vgstrParametrosSP = IIf(optClasificacion(1).Value, "AR", IIf(optClasificacion(2).Value, "OC", "*")) & "|-1|%|" & llngClaveCargo
            Set rs = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELARTICULOOTROCONCEPTO")
            If rs.RecordCount > 0 Then
                If Trim(txtDescripcion.Text) <> Trim(rs!Descripcion) Then
                    MsgBox SIHOMsg(539) & Chr(13) & "Descripción", vbCritical, "Mensaje"
                    txtDescripcion.SetFocus
                    Exit Sub
                End If
            Else
                'Existen datos incorrectos. Verifique
                MsgBox SIHOMsg(539) & Chr(13) & "Descripción", vbCritical, "Mensaje"
                txtDescripcion.SetFocus
                Exit Sub
            End If
            rs.Close
        End If
        
        vgstrParametrosSP = IIf(cboDepartamento.ListIndex = 0, -1, cboDepartamento.ItemData(cboDepartamento.ListIndex)) & "|" & _
                            cboHospital.ItemData(cboHospital.ListIndex) & "|" & _
                            IIf(optInternos.Value, "I", IIf(optExternos.Value, "E", "*")) & "|" & _
                            IIf(cboMedicos.ListIndex = 0, -1, cboMedicos.ItemData(cboMedicos.ListIndex)) & "|" & _
                            fstrFechaSQL(mskFecIni.Text, "00:00:00") & "|" & _
                            fstrFechaSQL(mskFecFin.Text, "23:59:59") & "|" & _
                            IIf(cboLocalizacion.ListIndex = 0, -1, cboLocalizacion.ItemData(cboLocalizacion.ListIndex)) & "|" & _
                            IIf(optClasificacion(1).Value, "AR", IIf(optClasificacion(2).Value, "OC", "*")) & "|" & _
                            IIf(optSubArticulos(1).Value, 0, IIf(optSubArticulos(2).Value, 1, -1)) & "|" & _
                            IIf(Trim(txtDescripcion.Text) = "", 0, llngClaveCargo)


                            
        Set rsReporte = frsEjecuta_SP(vgstrParametrosSP, "sp_PvRptProductividadMedicoArt")
        pInstanciaReporte vgrptReporte, "rptProductividadMedicoArt.rpt"
        If rsReporte.EOF Then
            MsgBox SIHOMsg(13), vbInformation, "Mensaje"
        Else
            vgrptReporte.DiscardSavedData
            
            alstrParametros(0) = "NombreHospital;" & cboHospital.List(cboHospital.ListIndex)
            alstrParametros(1) = "FechaInicio;" & UCase(Format(mskFecIni.Text, "dd/mmm/yyyy"))
            alstrParametros(2) = "FechaFin;" & UCase(Format(mskFecFin.Text, "dd/mmm/yyyy"))
            alstrParametros(3) = "Grupo;" & IIf(optAgrupar(0).Value, 0, 1)
            alstrParametros(4) = "Detallado;" & IIf(chkDetallado.Value, True, False)
            alstrParametros(5) = "Medico;" & IIf(optAmbos.Value, "TODOS LOS MÉDICOS", IIf(optInternos.Value, "MÉDICOS STAFF", "MÉDICOS EXTERNOS"))
            pCargaParameterFields alstrParametros, vgrptReporte
            
            pImprimeReporte vgrptReporte, rsReporte, vlstrDestino, "Productividad de médicos por artículo"
        End If
        rsReporte.Close
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pReporte"))
End Sub

Private Sub optAgrupacion_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
    
        SendKeys vbTab
        
    End If
End Sub

Private Sub optAgrupar_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
                
        SendKeys vbTab
        
    End If
    
End Sub

Private Sub optAmbos_Click()
       Dim vlstrx As String
    Dim rs As New ADODB.Recordset

       vlstrGrupo = ""
    vlstrx = "select intcvemedico, vchApellidoPaterno || ' ' || vchApellidoMaterno || ' ' || vchNombre as Nombre from HoMedico order by vchApellidoPaterno, vchApellidoMaterno, vchNombre"
    Set rs = frsRegresaRs(vlstrx)
    pLlenarCboRs cboMedicos, rs, 0, 1, 3
    cboMedicos.ListIndex = 0
End Sub

Private Sub optAmbos_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
    
        SendKeys vbTab
        
    End If
        
End Sub

Private Sub optClasificacion_Click(Index As Integer)
    If Index = 1 Then
        optSubArticulos(0).Enabled = True
        optSubArticulos(1).Enabled = True
        optSubArticulos(2).Enabled = True
    Else
        optSubArticulos(0).Value = True
        optSubArticulos(0).Enabled = False
        optSubArticulos(1).Enabled = False
        optSubArticulos(2).Enabled = False
    End If
End Sub

Private Sub optClasificacion_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
                
        SendKeys vbTab
        
    End If
    
End Sub

Private Sub optExternos_Click()
    Dim vlstrx As String
    Dim rs As New ADODB.Recordset
    
    vlstrGrupo = "E"
    
    vlstrx = "select intcvemedico, vchApellidoPaterno || ' ' || vchApellidoMaterno || ' ' || vchNombre as Nombre from HoMedico Where vchGrupo = '" & vlstrGrupo & "' order by vchApellidoPaterno, vchApellidoMaterno, vchNombre"
    Set rs = frsRegresaRs(vlstrx)
    pLlenarCboRs cboMedicos, rs, 0, 1, 3
    cboMedicos.ListIndex = 0
    
End Sub

Private Sub optInternos_Click()
    Dim vlstrx As String
    Dim rs As New ADODB.Recordset
    
    vlstrGrupo = "I"
    
    vlstrx = "select intcvemedico, vchApellidoPaterno || ' ' || vchApellidoMaterno || ' ' || vchNombre as Nombre from HoMedico Where vchGrupo = '" & vlstrGrupo & "' order by vchApellidoPaterno, vchApellidoMaterno, vchNombre"
    Set rs = frsRegresaRs(vlstrx)
    pLlenarCboRs cboMedicos, rs, 0, 1, 3
    cboMedicos.ListIndex = 0
    
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




Private Sub optSubArticulos_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
                
        SendKeys vbTab
        
    End If
    
End Sub

Private Sub optTipoReporte_Click(Index As Integer)

    If optTipoReporte(1).Value = True Then
    
        chkProcedencia.Enabled = True
        
    Else
        
        chkProcedencia.Enabled = False
        
    End If
    
    
End Sub

Private Sub optTipoReporte_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
    
        SendKeys vbTab
        
    End If
    
    
End Sub


Private Sub sstTipoReporte_Click(PreviousTab As Integer)
    If sstTipoReporte.Tab = 1 Then
        If cboDepartamento.Enabled Then
            cboDepartamento.SetFocus
        Else
            cboLocalizacion.SetFocus
        End If
        
        'sstTipoReporte.Height = 3500
        ' frmReporteProductividadMedicos.Height = 7100
        'Frame6.Top = 5800
        'Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
        
    'Else
    
         'sstTipoReporte.Height = 5205
         ' frmReporteProductividadMedicos.Height = 8865
         ' Frame6.Top = 7550
         ' Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
        
    End If
End Sub

Private Sub txtDescripcion_GotFocus()
    pSelTextBox txtDescripcion
End Sub

Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyReturn Or KeyCode = vbKeyRight Or KeyCode = vbKeyDown) And lstBusqueda.Visible Then
        
        lstBusqueda.SetFocus
               
    End If
    
    
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then
        KeyAscii = 0
        Exit Sub
    End If
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtDescripcion_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim rsConsulta As New ADODB.Recordset
    
    lstBusqueda.Visible = False
    lstBusqueda.Clear
    lstBusqueda.Top = 4740
    lstBusqueda.Left = 1330 '1200
    If Trim(txtDescripcion.Text) <> "" Then
        'vgstrParametrosSP = vgintNumeroDepartamento & "|" & lintAlmacenVenta & "|" & IIf(cboLocalizacion.ListIndex = 0, -1, cboLocalizacion.ItemData(cboLocalizacion.ListIndex)) & "|" & IIf(optClasificacion(1).Value, "AR", IIf(optClasificacion(2).Value, "OC", "*")) & "|" & IIf(optSubArticulos(1).Value, 0, IIf(optSubArticulos(2).Value, 1, -1)) & "|" & Trim(txtDescripcion.Text) & "|0"
        vgstrParametrosSP = IIf(optClasificacion(1).Value, "AR", IIf(optClasificacion(2).Value, "OC", "*")) & "|-1|" & Trim(txtDescripcion.Text) & "|0"
        Set rsConsulta = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELARTICULOOTROCONCEPTO")
        If rsConsulta.RecordCount <> 0 Then
            Do While Not rsConsulta.EOF
                lstBusqueda.AddItem Trim(rsConsulta!Descripcion)
                lstBusqueda.ItemData(lstBusqueda.newIndex) = rsConsulta!Clave
                rsConsulta.MoveNext
            Loop
            lstBusqueda.ListIndex = 0
            lstBusqueda.Visible = True
        End If
    End If
End Sub

Private Sub pCargaDepartamentosCredito()

Dim rs As ADODB.Recordset

        cboDeptoCredito.Clear
        
        vgstrParametrosSP = "-1|1|*|" & CStr(cboHospital.ItemData(cboHospital.ListIndex))
        Set rs = frsEjecuta_SP(vgstrParametrosSP, "Sp_Gnseldepartamento")
        
        If rs.RecordCount <> 0 Then
            
            pLlenarCboRs cboDeptoCredito, rs, 0, 1, 3
            cboDeptoCredito.ListIndex = 0
                                    
        End If

End Sub

