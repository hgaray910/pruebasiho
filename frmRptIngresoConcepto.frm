VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmRptIngresoConcepto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingresos por concepto de factura"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8565
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   8565
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab 
      Height          =   6465
      Left            =   -45
      TabIndex        =   27
      Top             =   -495
      Width           =   8730
      _ExtentX        =   15399
      _ExtentY        =   11404
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmRptIngresoConcepto.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame6"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame7"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fraAgrupado"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "chkDesglosar"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmRptIngresoConcepto.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label10"
      Tab(1).Control(1)=   "Frame5"
      Tab(1).ControlCount=   2
      Begin VB.CheckBox chkDesglosar 
         Caption         =   "Desglosar importes"
         Height          =   255
         Left            =   6870
         TabIndex        =   24
         ToolTipText     =   "Desglosar importes gravados"
         Top             =   6060
         Width           =   1628
      End
      Begin VB.Frame fraAgrupado 
         Caption         =   "Agrupación"
         Height          =   825
         Left            =   3645
         TabIndex        =   55
         Top             =   4740
         Width           =   4890
         Begin VB.OptionButton OptGroupConcentrada 
            Caption         =   "Tipo de ingreso"
            Height          =   345
            Index           =   1
            Left            =   1920
            TabIndex        =   23
            ToolTipText     =   "Agrupar por tipo de ingreso"
            Top             =   210
            Width           =   1400
         End
         Begin VB.OptionButton OptGroupConcentrada 
            Caption         =   "Ninguna"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   22
            ToolTipText     =   "No agrupar"
            Top             =   270
            Width           =   1125
         End
         Begin VB.OptionButton OptGroup 
            Caption         =   "Factura"
            Height          =   255
            Index           =   1
            Left            =   1920
            TabIndex        =   21
            ToolTipText     =   "Agrupar por factura"
            Top             =   510
            Width           =   975
         End
         Begin VB.OptionButton OptGroup 
            Caption         =   "Concepto"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   20
            ToolTipText     =   "Agrupar por concepto de factura"
            Top             =   510
            Width           =   1095
         End
      End
      Begin VB.Frame Frame7 
         Height          =   675
         Left            =   120
         TabIndex        =   53
         Top             =   480
         Width           =   8415
         Begin VB.ComboBox cboHospital 
            Height          =   315
            Left            =   915
            Style           =   2  'Dropdown List
            TabIndex        =   0
            ToolTipText     =   "Seleccione la empresa"
            Top             =   240
            Width           =   7380
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Empresa"
            Height          =   195
            Left            =   130
            TabIndex        =   54
            Top             =   300
            Width           =   615
         End
      End
      Begin VB.Frame Frame5 
         Height          =   5475
         Left            =   -74895
         TabIndex        =   49
         Top             =   525
         Width           =   7875
         Begin VB.TextBox txtAux 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   2610
            MaxLength       =   40
            TabIndex        =   56
            Top             =   4980
            Visible         =   0   'False
            Width           =   2685
         End
         Begin VB.CommandButton cmdSeleccionar 
            Caption         =   "Seleccionar"
            Height          =   495
            Left            =   5565
            TabIndex        =   52
            ToolTipText     =   "Seleccionar el cargo de la lista"
            Top             =   4830
            Width           =   2130
         End
         Begin VB.OptionButton opTipoCargo 
            Caption         =   "Otros conceptos"
            Height          =   195
            Index           =   5
            Left            =   6255
            TabIndex        =   33
            ToolTipText     =   "Cargos de otros conceptos"
            Top             =   285
            Width           =   1575
         End
         Begin VB.OptionButton opTipoCargo 
            Caption         =   "Grupos de exámenes"
            Height          =   195
            Index           =   4
            Left            =   4365
            TabIndex        =   32
            ToolTipText     =   "Cargos de grupos"
            Top             =   285
            Width           =   1800
         End
         Begin VB.OptionButton opTipoCargo 
            Caption         =   "Exámenes"
            Height          =   195
            Index           =   3
            Left            =   3210
            TabIndex        =   31
            ToolTipText     =   "Cargos de exámenes"
            Top             =   285
            Width           =   1050
         End
         Begin VB.OptionButton opTipoCargo 
            Caption         =   "Estudios"
            Height          =   195
            Index           =   2
            Left            =   2175
            TabIndex        =   30
            ToolTipText     =   "Cargos de estudios"
            Top             =   285
            Width           =   945
         End
         Begin VB.OptionButton opTipoCargo 
            Caption         =   "Artículos"
            Height          =   195
            Index           =   1
            Left            =   1050
            TabIndex        =   29
            ToolTipText     =   "Cargos artículos"
            Top             =   285
            Width           =   945
         End
         Begin VB.OptionButton opTipoCargo 
            Caption         =   "Todos"
            Height          =   195
            Index           =   0
            Left            =   105
            TabIndex        =   28
            ToolTipText     =   "Todos los tipos de cargos"
            Top             =   285
            Width           =   885
         End
         Begin VB.ListBox lstCargo 
            Height          =   3765
            Left            =   90
            TabIndex        =   35
            ToolTipText     =   "Lista de cargos encontrados"
            Top             =   960
            Width           =   7635
         End
         Begin VB.TextBox txtIniciales 
            Height          =   285
            Left            =   840
            MaxLength       =   20
            TabIndex        =   34
            ToolTipText     =   "Escriba las iniciales para realizar la búsqueda"
            Top             =   615
            Width           =   6885
         End
         Begin VB.Label lblIniciales 
            AutoSize        =   -1  'True
            Caption         =   "Iniciales"
            Height          =   195
            Left            =   120
            TabIndex        =   51
            Top             =   660
            Width           =   570
         End
         Begin VB.Label Label9 
            Caption         =   "Presione <ESC> para regresar"
            Height          =   195
            Left            =   90
            TabIndex        =   50
            Top             =   5100
            Width           =   3075
         End
      End
      Begin VB.Frame Frame6 
         Height          =   735
         Left            =   3752
         TabIndex        =   48
         Top             =   5580
         Width           =   1140
         Begin VB.CommandButton cmdImprimir 
            Height          =   495
            Left            =   570
            Picture         =   "frmRptIngresoConcepto.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "Imprimir"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdVista 
            Height          =   495
            Left            =   75
            Picture         =   "frmRptIngresoConcepto.frx":01DA
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Vista previa"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   495
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Presentaciones"
         Height          =   590
         Left            =   3645
         TabIndex        =   47
         Top             =   4140
         Width           =   4890
         Begin VB.OptionButton optPresentacion 
            Caption         =   "Detallada por cargo"
            Height          =   195
            Index           =   2
            Left            =   1560
            TabIndex        =   18
            ToolTipText     =   "Presentación por cargo"
            Top             =   285
            Width           =   1800
         End
         Begin VB.OptionButton optPresentacion 
            Caption         =   "Concentrada"
            Height          =   195
            Index           =   1
            Left            =   3500
            TabIndex        =   19
            ToolTipText     =   "Presentación concentrada"
            Top             =   285
            Width           =   1215
         End
         Begin VB.OptionButton optPresentacion 
            Caption         =   "Detallada"
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   17
            ToolTipText     =   "Presentación detallada"
            Top             =   285
            Width           =   1050
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Rango de fechas de factura"
         Height          =   1430
         Left            =   113
         TabIndex        =   44
         Top             =   4140
         Width           =   3465
         Begin MSMask.MaskEdBox mskFechaInicio 
            Height          =   315
            Left            =   765
            TabIndex        =   13
            ToolTipText     =   "Fecha de inicio"
            Top             =   315
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
            Left            =   765
            TabIndex        =   15
            ToolTipText     =   "Fecha final"
            Top             =   715
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mmm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskHoraInicial 
            Height          =   315
            Left            =   2160
            TabIndex        =   14
            ToolTipText     =   "Hora de inicio"
            Top             =   315
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "hh:mm:ss"
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskHoraFinal 
            Height          =   315
            Left            =   2160
            TabIndex        =   16
            ToolTipText     =   "Hora final"
            Top             =   720
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "hh:mm:ss"
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            Height          =   195
            Left            =   130
            TabIndex        =   46
            Top             =   375
            Width           =   465
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            Height          =   195
            Left            =   130
            TabIndex        =   45
            Top             =   775
            Width           =   420
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2925
         Left            =   120
         TabIndex        =   36
         Top             =   1140
         Width           =   8415
         Begin VB.ComboBox cboTipoIngreso 
            Height          =   315
            Left            =   2070
            Style           =   2  'Dropdown List
            TabIndex        =   12
            ToolTipText     =   "Tipo de ingreso"
            Top             =   2520
            Width           =   6225
         End
         Begin VB.ComboBox cboConcepto 
            Height          =   315
            Left            =   2070
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   1
            ToolTipText     =   "Selección del concepto de factura"
            Top             =   240
            Width           =   6225
         End
         Begin VB.ComboBox cboDepartamento 
            Height          =   315
            Left            =   2070
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   2
            ToolTipText     =   "Selección del departamento"
            Top             =   600
            Width           =   6225
         End
         Begin VB.ComboBox cboTipoPaciente 
            Height          =   315
            Left            =   2070
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   3
            ToolTipText     =   "Selección del tipo de paciente"
            Top             =   960
            Width           =   6225
         End
         Begin VB.CommandButton cmdLocate 
            Height          =   495
            Left            =   7800
            Picture         =   "frmRptIngresoConcepto.frx":037C
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Buscar el cargo"
            Top             =   1320
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.Frame Frame2 
            Height          =   400
            Left            =   2070
            TabIndex        =   37
            Top             =   1680
            Width           =   5625
            Begin VB.OptionButton optPaciente 
               Caption         =   "Ventas al público"
               Height          =   195
               Index           =   3
               Left            =   3315
               TabIndex        =   9
               Top             =   155
               Width           =   1510
            End
            Begin VB.OptionButton optPaciente 
               Caption         =   "Todos"
               Height          =   195
               Index           =   0
               Left            =   45
               TabIndex        =   6
               Top             =   155
               Width           =   745
            End
            Begin VB.OptionButton optPaciente 
               Caption         =   "Internos"
               Height          =   195
               Index           =   1
               Left            =   1065
               TabIndex        =   7
               Top             =   155
               Width           =   870
            End
            Begin VB.OptionButton optPaciente 
               Caption         =   "Externos"
               Height          =   195
               Index           =   2
               Left            =   2175
               TabIndex        =   8
               Top             =   155
               Width           =   910
            End
         End
         Begin VB.TextBox txtNumeroCuenta 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2070
            TabIndex        =   10
            ToolTipText     =   "Número de cuenta"
            Top             =   2160
            Width           =   1065
         End
         Begin VB.Label lblTipoIngreso 
            Caption         =   "Tipo de ingreso"
            Height          =   195
            Left            =   135
            TabIndex        =   57
            Top             =   2580
            Width           =   1935
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Concepto de factura"
            Height          =   195
            Left            =   130
            TabIndex        =   43
            Top             =   300
            Width           =   1455
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Departamento que facturó"
            Height          =   195
            Left            =   130
            TabIndex        =   42
            Top             =   660
            Width           =   1860
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de paciente"
            Height          =   195
            Left            =   130
            TabIndex        =   41
            Top             =   1020
            Width           =   1200
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Cargo"
            Height          =   195
            Left            =   130
            TabIndex        =   40
            Top             =   1380
            Width           =   420
         End
         Begin VB.Label lblCargo 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2070
            TabIndex        =   4
            ToolTipText     =   "Descripción del cargo seleccionado"
            Top             =   1320
            Width           =   5625
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Paciente"
            Height          =   195
            Left            =   130
            TabIndex        =   39
            Top             =   1783
            Width           =   630
         End
         Begin VB.Label lblNumeroCuenta 
            AutoSize        =   -1  'True
            Caption         =   "Número de cuenta"
            Height          =   195
            Left            =   135
            TabIndex        =   38
            Top             =   2220
            Width           =   1920
         End
         Begin VB.Label lblNombrePaciente 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   3150
            TabIndex        =   11
            ToolTipText     =   "Nombre del paciente"
            Top             =   2160
            Width           =   5145
         End
      End
      Begin VB.Label Label10 
         Caption         =   "Label10"
         Height          =   195
         Left            =   -74820
         TabIndex        =   58
         Top             =   5670
         Width           =   7275
      End
   End
End
Attribute VB_Name = "frmRptIngresoConcepto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs As New ADODB.Recordset   'Varios usos
Dim ldtmFecha As Date           'Fecha actual
Dim vlstrHoraInicio As String
Dim vlstrHoraFin As String
Dim vlblnLicenciaIEPS As Boolean


Public Sub pCargaCboTipoIngreso()

    Dim vlstrSentencia As String
    Dim rsTipoIngreso As New ADODB.Recordset

    'Tipos de Paciente
    If optPaciente(0) = True Then
        ' Todos los Pacientes
        vlstrSentencia = "select intcvetipoingreso, vchnombre from siTipoIngreso"
    Else
        If optPaciente(1) = True Then
            'Pacientes Internos
            vlstrSentencia = "select intcvetipoingreso, vchnombre from siTipoIngreso where CHRTIPOINGRESO = 'I'"
        Else
            'Pacientes Externos
            If optPaciente(2) = True Then
                vlstrSentencia = "select intcvetipoingreso, vchnombre from siTipoIngreso where CHRTIPOINGRESO = 'E'"
            End If
        End If
    End If
    
    Set rsTipoIngreso = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    pLlenarCboRs cboTipoIngreso, rsTipoIngreso, 0, 1
    rsTipoIngreso.Close
    
    cboTipoIngreso.AddItem "<TODOS>", 0
    cboTipoIngreso.ItemData(cboTipoIngreso.newIndex) = -1
    cboTipoIngreso.ListIndex = 0

End Sub

Private Sub cboHospital_Click()
    If cboHospital.ListIndex <> -1 Then pCargaDeptos
End Sub




Private Sub cmdImprimir_Click()
    On Error GoTo NotificaError
        If fblnValidos() Then pImprime "I"
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdImprimir_Click"))
End Sub

Private Sub cmdLocate_Click()
    On Error GoTo NotificaError

    SSTab.Tab = 1
    If opTipoCargo(0).Value Then
        opTipoCargo(0).SetFocus
    ElseIf opTipoCargo(1).Value Then
        opTipoCargo(1).SetFocus
    ElseIf opTipoCargo(2).Value Then
        opTipoCargo(2).SetFocus
    ElseIf opTipoCargo(3).Value Then
        opTipoCargo(3).SetFocus
    ElseIf opTipoCargo(4).Value Then
        opTipoCargo(4).SetFocus
    ElseIf opTipoCargo(5).Value Then
        opTipoCargo(5).SetFocus
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdLocate_Click"))
End Sub

Private Sub cmdSeleccionar_Click()
    On Error GoTo NotificaError

    lblCargo.Caption = lstCargo.List(lstCargo.ListIndex)
    SSTab.Tab = 0
    If optPaciente(0).Value Then
        optPaciente(0).SetFocus
    ElseIf optPaciente(1).Value Then
        optPaciente(1).SetFocus
    ElseIf optPaciente(2).Value Then
        optPaciente(2).SetFocus
    End If
        
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdSeleccionar_Click"))
End Sub

Private Sub pBuscaPaciente()
    On Error GoTo NotificaError

    If Trim(txtNumeroCuenta.Text) = "" Then
        With FrmBusquedaPacientes
            If optPaciente(2).Value Then    'Externos
                .vgstrTipoPaciente = "E"
                .Caption = .Caption & " externos"
                .vgblnPideClave = False
                .vgIntMaxRecords = 100
                .vgstrMovCve = "M"
                .optSoloActivos.Enabled = True
                .optSinFacturar.Enabled = True
                .optTodos.Enabled = True
                '.optSinFacturar.Value = True
                .optTodos.Value = True
                .vgStrOtrosCampos = ", CCempresa.vchDescripcion as Empresa, " & _
                " (Select Trim(GnDomicilio.vchCalle)||' '||Trim(GnDomicilio.vchNumeroExterior)||Case When GnDomicilio.vchNumeroInterior Is Null Then '' Else ' Int. '||Trim(GnDomicilio.vchNumeroInterior) End " & _
                " From ExPacienteDomicilio " & _
                " Inner Join GnDomicilio ON ExPacienteDomicilio.intCveDomicilio = GnDomicilio.intCveDomicilio " & _
                " And GnDomicilio.intCveTipoDomicilio = 1 " & _
                " Where ExPacienteDomicilio.intNumPaciente = ExPaciente.intNumPaciente) as Dirección, " & _
                " ExPaciente.dtmFechaNacimiento as ""Fecha Nac."", " & _
                " (Select GnTelefono.vchTelefono " & _
                " From ExPacienteTelefono " & _
                " Inner Join GnTelefono On ExPacienteTelefono.intCveTelefono = GnTelefono.intCveTelefono " & _
                " And GnTelefono.intCveTipoTelefono = 1 " & _
                " Where ExPacienteTelefono.intNumPaciente = ExPaciente.intNumpaciente) as Telefono "
                .vgstrTamanoCampo = "800,3400,2800,4100,990,980"
            Else
                .vgstrTipoPaciente = "I"  'Internos
                .vgblnPideClave = False
                .Caption = .Caption & " internos"
                .vgIntMaxRecords = 100
                .vgstrMovCve = "M"
                .optSinFacturar.Value = True
                .optSinFacturar.Enabled = True
                .optSoloActivos.Enabled = True
                .optTodos.Enabled = True
                .optTodos.Value = True
                .vgStrOtrosCampos = ", CCempresa.vchDescripcion as Empresa, " & _
                    " (Select Trim(GnDomicilio.vchCalle)||' '||Trim(GnDomicilio.vchNumeroExterior)||Case When GnDomicilio.vchNumeroInterior Is Null Then '' Else ' Int. '||Trim(GnDomicilio.vchNumeroInterior) End " & _
                " From ExPacienteDomicilio " & _
                " Inner Join GnDomicilio ON ExPacienteDomicilio.intCveDomicilio = GnDomicilio.intCveDomicilio " & _
                " And GnDomicilio.intCveTipoDomicilio = 1 " & _
                " Where ExPacienteDomicilio.intNumPaciente = ExPaciente.intNumPaciente) as Dirección, " & _
                " ExPaciente.dtmFechaNacimiento as ""Fecha Nac."", " & _
                " (Select GnTelefono.vchTelefono " & _
                " From ExPacienteTelefono " & _
                " Inner Join GnTelefono On ExPacienteTelefono.intCveTelefono = GnTelefono.intCveTelefono " & _
                " And GnTelefono.intCveTipoTelefono = 1 " & _
                " Where ExPacienteTelefono.intNumPaciente = ExPaciente.intNumpaciente) as Telefono "
                .vgstrTamanoCampo = "950,3400,2800,4100,990,980"
            End If
            
            txtNumeroCuenta.Text = .flngRegresaPaciente()
               
            If Val(txtNumeroCuenta.Text) <> -1 Then
                pCargaNombrePaciente
            Else
                txtNumeroCuenta.Text = ""
            End If
        End With
    Else
        pCargaNombrePaciente
    End If
    If Trim(lblNombrePaciente.Caption) <> "" Then
        mskFechaInicio.SetFocus
    Else
        pEnfocaTextBox txtNumeroCuenta
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pBuscaPaciente"))
End Sub

Private Sub pCargaNombrePaciente()
    On Error GoTo NotificaError

    vgstrParametrosSP = txtNumeroCuenta.Text & "|" & "0" & "|" & IIf(optPaciente(1).Value, "I", "E") & "|" & vgintClaveEmpresaContable
    Set rs = frsEjecuta_SP(vgstrParametrosSP, "sp_PvSelDatosPaciente")
    If rs.RecordCount <> 0 Then
        lblNombrePaciente.Caption = rs!Nombre
    Else
        MsgBox SIHOMsg(355), vbOKOnly + vbInformation, "Mensaje" 'No se encontró la información del paciente.
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargaNombrePaciente"))
End Sub

Private Sub cmdVista_Click()
    On Error GoTo NotificaError
        If fblnValidos() Then pImprime "P"
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdVista_Click"))
End Sub

Private Function fblnValidos() As Boolean
    On Error GoTo NotificaError
    
    fblnValidos = True
    If fstrFechaSQL(mskFechaInicio.Text, mskHoraInicial.Text) > fstrFechaSQL(mskFechaFin.Text, mskHoraFinal.Text) Then
        fblnValidos = False
        MsgBox SIHOMsg(64), vbOKOnly + vbInformation, "Mensaje" '¡Rango de fechas no válido!
        mskFechaInicio.SetFocus
    End If
    If fblnValidos And Val(txtNumeroCuenta.Text) <> 0 And Trim(lblNombrePaciente.Caption) = "" Then
        fblnValidos = False
        MsgBox SIHOMsg(27), vbOKOnly + vbInformation, "Mensaje" 'Debe seleccionar un paciente.
        txtNumeroCuenta.SetFocus
    End If

    Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fblnValidos"))
End Function
Private Sub pImprime(strDestino As String)
    On Error GoTo NotificaError
    Dim rptReporte As CRAXDRT.Report
    Dim alstrParametros(15) As String
    Dim lngCveCargo As Long
    Dim strTipoCargo As String
    Dim strTipoPaciente As String

    lngCveCargo = -1
    If lstCargo.ListCount <> 0 Then
        If lstCargo.ListIndex >= 0 Then lngCveCargo = lstCargo.ItemData(lstCargo.ListIndex)
    End If
    
    strTipoCargo = IIf(opTipoCargo(0).Value, "*", IIf(opTipoCargo(1).Value, "AR", IIf(opTipoCargo(2).Value, "ES", IIf(opTipoCargo(3).Value, "EX", IIf(opTipoCargo(4).Value, "GE", "OC")))))
    strTipoPaciente = IIf(optPaciente(0).Value, "*", IIf(optPaciente(1).Value, "I", IIf(optPaciente(2).Value, "E", "T")))

    vgstrParametrosSP = Str(cboConcepto.ItemData(cboConcepto.ListIndex)) _
        & "|" & CStr(cboDepartamento.ItemData(cboDepartamento.ListIndex)) _
        & "|" & CStr(cboTipoPaciente.ItemData(cboTipoPaciente.ListIndex)) _
        & "|" & CStr(lngCveCargo) _
        & "|" & strTipoCargo _
        & "|" & strTipoPaciente _
        & "|" & IIf(Trim(lblNombrePaciente.Caption) = "", "-1", txtNumeroCuenta.Text) _
        & "|" & "'" & Trim(mskFechaInicio.Text) & " " & Trim(mskHoraInicial.Text) & "'" _
        & "|" & "'" & Trim(mskFechaFin.Text) & " " & Trim(mskHoraFinal.Text) & "'" _
        & "|" & CStr(cboHospital.ItemData(cboHospital.ListIndex)) _
        & "|" & IIf(optPresentacion(0).Value, 0, 1) & "|" & vgintClaveEmpresaContable _
        & "|" & chkDesglosar.Value _
        & "|" & CStr(cboTipoIngreso.ItemData(cboTipoIngreso.ListIndex))
        
        
    
    Set rs = frsEjecuta_SP(vgstrParametrosSP, "SP_PVRPTINGRESOCONCEPTO")
    If rs.RecordCount <> 0 Then
        vlblnLicenciaIEPS = fblLicenciaIEPS
        
        If optPresentacion(2).Value Then
            pInstanciaReporte rptReporte, "rptIngresoDetalladoCargo.rpt"
        Else
            pInstanciaReporte rptReporte, IIf(optPresentacion(0).Value, IIf(OptGroup(0).Value, "rptIngresoConcepto.rpt", "rptIngresoConceptoAgrupadoFactura.rpt"), "rptIngresoConcentrado.rpt")
        End If
        rptReporte.DiscardSavedData
        alstrParametros(0) = "NombreHospital;" & Trim(cboHospital.List(cboHospital.ListIndex))
        alstrParametros(1) = "Concepto;" & Trim(cboConcepto.List(cboConcepto.ListIndex))
        alstrParametros(2) = "TipoPaciente;" & Trim(cboTipoPaciente.List(cboTipoPaciente.ListIndex))
        alstrParametros(3) = "Cargo;" & Trim(lblCargo.Caption)
        alstrParametros(4) = "Paciente;" & IIf(strTipoPaciente = "*", "<TODOS>", IIf(strTipoPaciente = "I", "INTERNOS", IIf(strTipoPaciente = "E", "EXTERNOS", "VENTAS AL PÚBLICO")))
        alstrParametros(5) = "FechaInicio;" & UCase(Format(mskFechaInicio.Text, "dd/mmm/yyyy"))
        alstrParametros(6) = "FechaFin;" & UCase(Format(mskFechaFin.Text, "dd/mmm/yyyy"))
        alstrParametros(7) = "Departamento;" & Trim(cboDepartamento.List(cboDepartamento.ListIndex))
        alstrParametros(8) = "Desglosar;" & IIf(vlblnLicenciaIEPS, IIf(chkDesglosar.Value = 0, 1, 2), IIf(chkDesglosar.Value = 0, 3, 4))
        alstrParametros(9) = "Cuenta" & ";" & IIf(Trim(txtNumeroCuenta.Text) = "", "<TODOS>", txtNumeroCuenta.Text & " - ") & ";STRING"
        alstrParametros(10) = "NombrePaciente" & ";" & IIf(Trim(lblNombrePaciente.Caption) = "", " ", lblNombrePaciente.Caption) & ";STRING"
        alstrParametros(11) = "HoraInicio;" & (Format(mskHoraInicial.Text, "hh:mm:ss"))
        alstrParametros(12) = "HoraFin;" & (Format(mskHoraFinal.Text, "hh:mm:ss"))
        alstrParametros(13) = "TipoIngreso;" & Trim(cboTipoIngreso.List(cboTipoIngreso.ListIndex))
        alstrParametros(14) = "Agrupar;" & IIf(OptGroupConcentrada(0) = 0, 0, 1)
        
        pCargaParameterFields alstrParametros, rptReporte
        pImprimeReporte rptReporte, rs, strDestino, "Ingresos por concepto de facturación"
    Else
        MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje" 'No existe información con esos parámetro
    End If
    rs.Close
    
    frsEjecuta_SP 1 & "|" & Me.Name & "|" & chkDesglosar.Name & "|Value|" & vglngNumeroLogin & "|" & Trim(Str(chkDesglosar.Value)), "SP_GNSELULTIMACONFIGURACION", True

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pImprime"))
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then
        If Me.ActiveControl.Name = "txtNumeroCuenta" Then
            pBuscaPaciente
        Else
            SendKeys vbTab
        End If
    Else
        If KeyCode = vbKeyEscape Then
            If SSTab.Tab = 1 Then
                SSTab.Tab = 0
                If optPaciente(0).Value Then
                    optPaciente(0).SetFocus
                ElseIf optPaciente(1).Value Then
                    optPaciente(1).SetFocus
                ElseIf optPaciente(2).Value Then
                    optPaciente(2).SetFocus
                ElseIf optPaciente(3).Value Then
                    optPaciente(3).SetFocus
                End If
            Else
                Unload Me
            End If
        End If
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_KeyDown"))
End Sub

Private Sub Form_Load()
    On Error GoTo NotificaError
    Dim lngNumOpcion As Long
    Dim rsValor As New ADODB.Recordset
    
    Me.Icon = frmMenuPrincipal.Icon
    Select Case cgstrModulo
    Case "PV"
         lngNumOpcion = 338
    Case "SE"
         lngNumOpcion = 1361
    End Select
    
    Set rsValor = frsEjecuta_SP(0 & "|" & Me.Name & "|" & chkDesglosar.Name & "|Value|" & vglngNumeroLogin & "|" & Trim(Str(chkDesglosar.Value)), "SP_GNSELULTIMACONFIGURACION")
    If rsValor.RecordCount <> 0 Then
        chkDesglosar.Value = IIf(Trim(rsValor!VCHVALOR) = "0", 0, 1)
    Else
        chkDesglosar.Value = 0
    End If
    rsValor.Close
    
    optPaciente(0) = True
    pCargaHospital lngNumOpcion
    pCargaConceptos      'Conceptos
    pCargaTipoPaciente   'Tipos de paciente
    pCargaCboTipoIngreso 'Tipo de ingreso
    

    lblCargo.Caption = "<TODOS>"
    
    optPaciente(0).Value = True
    OptPaciente_Click 0
    
    ldtmFecha = fdtmServerFecha
    
    mskFechaInicio.Mask = ""
    mskFechaInicio.Text = ldtmFecha
    mskFechaInicio.Mask = "##/##/####"
    mskFechaFin.Mask = ""
    mskFechaFin.Text = ldtmFecha
    mskFechaFin.Mask = "##/##/####"
    
    mskHoraInicial.Mask = ""
    mskHoraInicial.Text = "00:00:00"
    mskHoraInicial.Mask = "##:##:##"
    mskHoraFinal.Mask = ""
    mskHoraFinal.Text = "23:59:59"
    mskHoraFinal.Mask = "##:##:##"
    
    vlstrHoraInicio = "00:00:00"
    vlstrHoraFin = "23:59:59"
    
    optPresentacion(0).Value = True
    opTipoCargo(0).Value = True
    SSTab.Tab = 0
    
    
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_Load"))
End Sub

Private Sub pCargaDeptos()
    On Error GoTo NotificaError

    Set rs = frsEjecuta_SP("-1|1|*|" & CStr(cboHospital.ItemData(cboHospital.ListIndex)), "sp_GnSelDepartamento")
    If rs.RecordCount <> 0 Then
        pLlenarCboRs cboDepartamento, rs, 0, 1
    End If
    cboDepartamento.AddItem "<TODOS>", 0
    cboDepartamento.ItemData(cboDepartamento.newIndex) = -1
    cboDepartamento.ListIndex = 0

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargaDeptos"))
End Sub

Private Sub pCargaTipoPaciente()
    On Error GoTo NotificaError

    Set rs = frsEjecuta_SP("1|*", "SP_PVSELTIPOPACIENTE")
    If rs.RecordCount <> 0 Then
        pLlenarCboRs cboTipoPaciente, rs, 0, 1
    End If
    cboTipoPaciente.AddItem "<TODOS>", 0
    cboTipoPaciente.ItemData(cboTipoPaciente.newIndex) = 0
    cboTipoPaciente.ListIndex = 0

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargaTipoPaciente"))
End Sub

Private Sub pCargaConceptos()
    On Error GoTo NotificaError

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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If SSTab.Tab = 1 Then
        Cancel = 1
        SSTab.Tab = 0
        cmdLocate.SetFocus
    End If
End Sub

Private Sub lstCargo_DblClick()
    On Error GoTo NotificaError
        cmdSeleccionar_Click
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":lstCargo_DblClick"))
End Sub

Private Sub lstCargo_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
        If KeyCode = vbKeyReturn Then lstCargo_DblClick
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":lstCargo_KeyDown"))
End Sub

Private Sub mskFechaFin_GotFocus()
    On Error GoTo NotificaError
        pSelMkTexto mskFechaFin
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFechaFin_GotFocus"))
End Sub

Private Sub mskFechaFin_LostFocus()
    On Error GoTo NotificaError
        If Not IsDate(Format(mskFechaFin.Text, "dd/mm/yyyy")) Then
            mskFechaFin.Mask = ""
            mskFechaFin.Text = ldtmFecha
            mskFechaFin.Mask = "##/##/####"
        Else
            mskFechaFin.Text = Format(mskFechaFin.Text, "dd/mm/yyyy")
        End If
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFechaFin_LostFocus"))
End Sub

Private Sub mskFechaInicio_GotFocus()
    On Error GoTo NotificaError
        pSelMkTexto mskFechaInicio
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFechaInicio_GotFocus"))
End Sub

Private Sub mskFechaInicio_LostFocus()
    On Error GoTo NotificaError
        If Not IsDate(Format(mskFechaInicio.Text, "dd/mm/yyyy")) Then
            mskFechaInicio.Mask = ""
            mskFechaInicio.Text = ldtmFecha
            mskFechaInicio.Mask = "##/##/####"
        Else
            mskFechaInicio.Text = Format(mskFechaInicio.Text, "dd/mm/yyyy")
        End If
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFechaInicio_LostFocus"))
End Sub

Private Sub mskHoraFinal_GotFocus()
    pSelMkTexto mskHoraFinal
End Sub

Private Sub mskHoraFinal_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    If KeyCode = vbKeyReturn Then
        If Not IsDate(mskHoraFinal) Then
            '¡Hora no válida!, formato de hora hh:mm:ss
            MsgBox SIHOMsg(41), vbOKOnly + vbExclamation, "Mensaje"
            If fblnCanFocus(mskHoraFinal) Then mskHoraFinal.SetFocus
        End If
    End If

    If KeyCode = vbKeyLeft Then If fblnCanFocus(mskFechaFin) Then mskFechaFin.SetFocus

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskHoraFinal_KeyDown"))
End Sub

Private Sub mskHoraFinal_LostFocus()
    If Not IsDate(mskHoraFinal.Text) Then
        mskHoraFinal.Mask = ""
        mskHoraFinal.Text = vlstrHoraFin
        mskHoraFinal.Mask = "##:##:##"
    Else
        vlstrHoraFin = Trim(mskHoraFinal.Text)
    End If
End Sub

Private Sub mskHoraInicial_GotFocus()
    pSelMkTexto mskHoraInicial
End Sub

Private Sub mskHoraInicial_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    If KeyCode = vbKeyReturn Then
        If Not IsDate(mskHoraInicial) Then
            '¡Hora no válida!, formato de hora hh:mm:ss
            MsgBox SIHOMsg(41), vbOKOnly + vbExclamation, "Mensaje"
            If fblnCanFocus(mskHoraInicial) Then mskHoraInicial.SetFocus
        End If
    End If

    If KeyCode = vbKeyLeft Then If fblnCanFocus(mskFechaInicio) Then mskFechaInicio.SetFocus

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskHoraInicial_KeyDown"))
End Sub

Private Sub mskHoraInicial_LostFocus()
    If Not IsDate(mskHoraInicial.Text) Then
        mskHoraInicial.Mask = ""
        mskHoraInicial.Text = vlstrHoraInicio
        mskHoraInicial.Mask = "##:##:##"
    Else
        vlstrHoraInicio = Trim(mskHoraInicial.Text)
    End If
End Sub

Private Sub opTipoCargo_Click(Index As Integer)
    On Error GoTo NotificaError

    If opTipoCargo(0).Value Then
        lblCargo.Caption = "<TODOS>"
    ElseIf opTipoCargo(1).Value Then
        lblCargo.Caption = "<TODOS LOS ARTICULOS>"
    ElseIf opTipoCargo(2).Value Then
        lblCargo.Caption = "<TODOS LOS ESTUDIOS>"
    ElseIf opTipoCargo(3).Value Then
        lblCargo.Caption = "<TODOS LOS EXAMENES>"
    ElseIf opTipoCargo(4).Value Then
        lblCargo.Caption = "<TODOS LOS GRUPOS DE EXAMENES>"
    ElseIf opTipoCargo(5).Value Then
        lblCargo.Caption = "<TODOS LOS OTROS CONCEPTOS>"
    End If
    
    txtIniciales.Text = ""
    txtIniciales_KeyUp 13, 0
    lblIniciales.Enabled = Index <> 0
    txtIniciales.Enabled = Index <> 0
    lstCargo.Enabled = Index <> 0
    cmdSeleccionar.Enabled = Index <> 0 And lstCargo.ListCount <> 0

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":opTipoCargo_Click"))
End Sub

Private Sub OptPaciente_Click(Index As Integer)
    On Error GoTo NotificaError

    
    lblNumeroCuenta.Enabled = Index <> 0 And Index <> 3
    txtNumeroCuenta.Enabled = Index <> 0 And Index <> 3
    lblNombrePaciente.Enabled = Index <> 0 And Index <> 3
    txtNumeroCuenta.Text = ""
    If optPaciente(3) = True Then
        lblTipoIngreso.Enabled = False
        cboTipoIngreso.Enabled = False
    Else
        lblTipoIngreso.Enabled = True
        cboTipoIngreso.Enabled = True
        pCargaCboTipoIngreso
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":optPaciente_Click"))
End Sub

Private Sub optPresentacion_Click(Index As Integer)
    If optPresentacion(1).Value = True Then
        fraAgrupado.Enabled = True
        OptGroup(0).Value = False
        OptGroup(0).Enabled = False
        OptGroup(1).Value = False
        OptGroup(1).Enabled = False
        OptGroupConcentrada(1).Enabled = True
        OptGroupConcentrada(0).Enabled = True
        OptGroupConcentrada(0).Value = True
    ElseIf optPresentacion(0).Value = True Then
        fraAgrupado.Enabled = True
        OptGroup(0).Value = True
        OptGroup(0).Enabled = True
        OptGroup(1).Enabled = True
        OptGroupConcentrada(1).Enabled = False
        OptGroupConcentrada(0).Enabled = False
        OptGroupConcentrada(0).Value = False
    Else
        fraAgrupado.Enabled = True
        OptGroup(0).Value = False
        OptGroup(0).Enabled = False
        OptGroup(1).Value = False
        OptGroup(1).Enabled = False
        OptGroupConcentrada(1).Enabled = False
        OptGroupConcentrada(0).Enabled = True
        OptGroupConcentrada(0).Value = True
    End If
End Sub
Private Sub txtIniciales_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtIniciales_KeyPress"))
End Sub

Private Sub txtIniciales_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    pCargaCargos
    lstCargo.Enabled = lstCargo.ListCount <> 0
    cmdSeleccionar.Enabled = lstCargo.ListCount <> 0
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtIniciales_KeyUp"))
End Sub

Private Sub pCargaCargos()
    On Error GoTo NotificaError
    Dim strSentencia As String
    
    txtAux.Text = Replace(txtIniciales.Text, "'", "''")
    
    lstCargo.Clear
    If opTipoCargo(1).Value Then
        strSentencia = "select intIdArticulo Clave, vchNombreComercial Descripcion from ivArticulo "
        PSuperBusqueda txtAux, strSentencia, lstCargo, "vchNombreComercial", 100, " and CHRCOSTOGASTO <> 'G' and vchEstatus = 'ACTIVO' ", "vchNombreComercial"
    ElseIf opTipoCargo(2).Value Then
        strSentencia = "select intCveEstudio Clave, vchNombre Descripcion from imEstudio "
        PSuperBusqueda txtAux, strSentencia, lstCargo, "vchNombre", 100, " and bitStatusActivo = 1 ", "vchNombre"
    ElseIf opTipoCargo(3).Value Then
        strSentencia = "select intCveExamen Clave, chrNombre Descripcion from laExamen "
        PSuperBusqueda txtAux, strSentencia, lstCargo, "chrNombre", 100, " and bitEstatusActivo = 1 ", "chrNombre"
    ElseIf opTipoCargo(4).Value Then
        strSentencia = "select intCveGrupo Clave, chrNombre Descripcion from laGrupoExamen "
        PSuperBusqueda txtAux, strSentencia, lstCargo, "chrNombre", 100, " and bitEstatusActivo = 1 ", "chrNombre"
    ElseIf opTipoCargo(5).Value Then
        strSentencia = "select intCveConcepto Clave, chrDescripcion Descripcion from pvOtroConcepto "
        PSuperBusqueda txtAux, strSentencia, lstCargo, "chrDescripcion", 100, " and bitEstatus = 1 ", "chrDescripcion"
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargaCargos"))
End Sub

Private Sub txtNumeroCuenta_Change()
    On Error GoTo NotificaError
        lblNombrePaciente.Caption = ""
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtNumeroCuenta_Change"))
End Sub

Private Sub txtNumeroCuenta_GotFocus()
    On Error GoTo NotificaError
        pSelTextBox txtNumeroCuenta
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtNumeroCuenta_GotFocus"))
End Sub

Private Sub txtNumeroCuenta_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
        If UCase(Chr(KeyAscii)) = "E" Or UCase(Chr(KeyAscii)) = "I" Then
            optPaciente(1).Value = UCase(Chr(KeyAscii)) = "I"
            optPaciente(2).Value = UCase(Chr(KeyAscii)) = "E"
        End If
        KeyAscii = 7
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtNumeroCuenta_KeyPress"))
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

Private Sub txtNumeroCuenta_LostFocus()

    If txtNumeroCuenta <> " " Then
        lblTipoIngreso.Enabled = False
        cboTipoIngreso.Enabled = False
    Else
        lblTipoIngreso.Enabled = False
        cboTipoIngreso.Enabled = False
    End If

End Sub


