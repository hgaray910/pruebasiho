VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCorte 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Corte"
   ClientHeight    =   8325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11595
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   11595
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTabCorte 
      Height          =   9350
      Left            =   -75
      TabIndex        =   35
      Top             =   -420
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   16484
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmCorte.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label12"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label13"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblTotalPesos"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblTotalDolares"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lstFormas"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame4"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame5"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "freRangos"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "tmrHora"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "fraRptCorte"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "fraRptFueraCorte"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmCorte.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grdCortes"
      Tab(1).Control(1)=   "Frame3"
      Tab(1).ControlCount=   2
      Begin VB.Frame fraRptFueraCorte 
         Caption         =   "Movimientos fuera del corte"
         Height          =   800
         Left            =   5780
         TabIndex        =   78
         Top             =   7785
         Width           =   2870
         Begin VB.CommandButton cmdMovsFueraCorte 
            Caption         =   "Detallado"
            Height          =   495
            Left            =   60
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmCorte.frx":0038
            TabIndex        =   25
            ToolTipText     =   "Imprimir detalle de movimientos fuera del corte actual"
            Top             =   200
            UseMaskColor    =   -1  'True
            Width           =   1350
         End
         Begin VB.CommandButton cmdPolizaFueraCorte 
            Caption         =   "Pólizas"
            Height          =   495
            Left            =   1410
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmCorte.frx":0566
            TabIndex        =   26
            ToolTipText     =   "Imprimir pólizas de movimientos fuera del corte actual"
            Top             =   200
            UseMaskColor    =   -1  'True
            Width           =   1350
         End
      End
      Begin VB.Frame fraRptCorte 
         Caption         =   "Movimientos del corte"
         Height          =   800
         Left            =   135
         TabIndex        =   77
         Top             =   7785
         Width           =   5570
         Begin VB.CommandButton cmdPrintPoliza 
            Caption         =   "Póliza"
            Height          =   495
            Left            =   4110
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmCorte.frx":0924
            TabIndex        =   24
            ToolTipText     =   "Imprimir póliza"
            Top             =   200
            UseMaskColor    =   -1  'True
            Width           =   1350
         End
         Begin VB.CommandButton cmdCorteFormasPago 
            Caption         =   "Formas de pago"
            Height          =   495
            Left            =   2760
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmCorte.frx":0CE2
            TabIndex        =   23
            ToolTipText     =   "Imprimir el detalle del corte por formas de pago"
            Top             =   200
            UseMaskColor    =   -1  'True
            Width           =   1350
         End
         Begin VB.CommandButton cmdCorteCronologico 
            Caption         =   "Cronológico"
            Height          =   495
            Left            =   1410
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmCorte.frx":12C8
            TabIndex        =   22
            ToolTipText     =   "Imprimir el detalle del corte cronológicamente"
            Top             =   200
            UseMaskColor    =   -1  'True
            Width           =   1350
         End
         Begin VB.CommandButton cmdCorteporMovimiento 
            Caption         =   "Movimientos"
            Height          =   495
            Left            =   60
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmCorte.frx":190A
            TabIndex        =   21
            ToolTipText     =   "Imprimir detalle del corte por tipo de movimiento"
            Top             =   200
            UseMaskColor    =   -1  'True
            Width           =   1350
         End
      End
      Begin VB.Timer tmrHora 
         Interval        =   1000
         Left            =   11160
         Top             =   6000
      End
      Begin VB.Frame freRangos 
         Caption         =   "Rango de reporte"
         Height          =   1980
         Left            =   135
         TabIndex        =   55
         Top             =   5700
         Width           =   4500
         Begin VB.ComboBox cboEmpleadoRegistro 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   1185
            Width           =   4290
         End
         Begin VB.CheckBox chkAgrupadoPorEmpleado 
            Caption         =   "Agrupado por empleado"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   1605
            Width           =   2020
         End
         Begin MSMask.MaskEdBox mskHoraIni 
            Height          =   300
            Left            =   1440
            TabIndex        =   6
            ToolTipText     =   "Hora inicial"
            Top             =   405
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   5
            Format          =   "hh:mm"
            Mask            =   "##:##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.DTPicker dtpFechaInicio 
            Height          =   315
            Left            =   120
            TabIndex        =   4
            ToolTipText     =   "Fecha inicial"
            Top             =   390
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            Format          =   109707265
            CurrentDate     =   37351
         End
         Begin MSComCtl2.DTPicker dtpFechaFin 
            Height          =   315
            Left            =   2505
            TabIndex        =   5
            ToolTipText     =   "Fecha final"
            Top             =   390
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            Format          =   109707265
            CurrentDate     =   37351
         End
         Begin MSMask.MaskEdBox mskHoraFin 
            Height          =   300
            Left            =   3810
            TabIndex        =   7
            ToolTipText     =   "Hora final"
            Top             =   405
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   5
            Format          =   "hh:mm"
            Mask            =   "##:##"
            PromptChar      =   " "
         End
         Begin VB.Label lbla 
            Caption         =   "a"
            Height          =   180
            Left            =   2235
            TabIndex        =   57
            Top             =   450
            Width           =   105
         End
         Begin VB.Label lblEmpleadoRegistro 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Empleado que registró"
            Height          =   195
            Left            =   120
            TabIndex        =   56
            Top             =   915
            Width           =   1575
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Incluir en reportes"
         Height          =   2490
         Left            =   4695
         TabIndex        =   53
         Top             =   5190
         Width           =   3600
         Begin VB.CheckBox chkSalidaChica 
            Caption         =   "Salidas caja chica"
            Height          =   195
            Left            =   1725
            TabIndex        =   76
            Top             =   585
            Value           =   1  'Checked
            Width           =   1800
         End
         Begin VB.CheckBox chkEntradaChica 
            Caption         =   "Entradas caja chica"
            Height          =   195
            Left            =   1725
            TabIndex        =   75
            Top             =   315
            Value           =   1  'Checked
            Width           =   1800
         End
         Begin VB.CheckBox chkTransferencias 
            Caption         =   "Transferencias"
            Height          =   195
            Left            =   150
            TabIndex        =   16
            Top             =   1845
            Value           =   1  'Checked
            Width           =   1410
         End
         Begin VB.CheckBox chkFondoFijo 
            Caption         =   "Fondo fijo"
            Height          =   195
            Left            =   150
            TabIndex        =   17
            Top             =   2085
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CheckBox chkPagosCredito 
            Caption         =   "Pagos a créditos"
            Height          =   195
            Left            =   150
            TabIndex        =   14
            Top             =   1350
            Value           =   1  'Checked
            Width           =   1485
         End
         Begin VB.CheckBox chkHonorarios 
            Caption         =   "Honorarios"
            Height          =   195
            Left            =   150
            TabIndex        =   15
            Top             =   1590
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CheckBox chkFacturas 
            Caption         =   "Facturas"
            Height          =   195
            Left            =   150
            TabIndex        =   10
            Top             =   315
            Value           =   1  'Checked
            Width           =   945
         End
         Begin VB.CheckBox chkSoloCancelados 
            Caption         =   "Sólo documentos cancelados"
            Height          =   375
            Left            =   1725
            TabIndex        =   18
            Top             =   960
            Width           =   1650
         End
         Begin VB.CheckBox chkPagos 
            Caption         =   "Pagos"
            Height          =   195
            Left            =   150
            TabIndex        =   11
            Top             =   585
            Value           =   1  'Checked
            Width           =   990
         End
         Begin VB.CheckBox chkTickets 
            Caption         =   "Tickets"
            Height          =   195
            Left            =   150
            TabIndex        =   12
            Top             =   855
            Value           =   1  'Checked
            Width           =   990
         End
         Begin VB.CheckBox chkSalidas 
            Caption         =   "Salidas"
            Height          =   195
            Left            =   150
            TabIndex        =   13
            Top             =   1110
            Value           =   1  'Checked
            Width           =   1425
         End
         Begin VB.OptionButton optGrupoTipoDocumento 
            Caption         =   "Tipo documento"
            Height          =   195
            Left            =   1725
            TabIndex        =   20
            Top             =   2085
            Width           =   1470
         End
         Begin VB.OptionButton optGrupoFormaPago 
            Caption         =   "Forma de pago"
            Height          =   210
            Left            =   1725
            TabIndex        =   19
            Top             =   1830
            Value           =   -1  'True
            Width           =   1500
         End
         Begin VB.Label Label16 
            Caption         =   "Agrupado por"
            Height          =   225
            Left            =   1755
            TabIndex        =   54
            Top             =   1440
            Width           =   1035
         End
      End
      Begin VB.Frame Frame4 
         Height          =   480
         Left            =   135
         TabIndex        =   52
         Top             =   5190
         Width           =   4500
         Begin VB.CheckBox chkIntermedios 
            Caption         =   "Reportes intermedios"
            Height          =   195
            Left            =   1170
            TabIndex        =   3
            ToolTipText     =   "Seleccionar filtros para reportes intermedios"
            Top             =   180
            Width           =   1800
         End
         Begin VB.CheckBox chkAcumulado 
            Caption         =   "Acumulado"
            Height          =   195
            Left            =   60
            TabIndex        =   2
            ToolTipText     =   "Mostrar el acumulado o el detalle del corte"
            Top             =   180
            Value           =   1  'Checked
            Width           =   1125
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command1"
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   -480
         Width           =   1455
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdCortes 
         Height          =   7440
         Left            =   -74850
         TabIndex        =   46
         ToolTipText     =   "Cortes"
         Top             =   1095
         Width           =   11460
         _ExtentX        =   20214
         _ExtentY        =   13123
         _Version        =   393216
         Rows            =   0
         FixedRows       =   0
         GridColor       =   12632256
         SelectionMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Frame Frame3 
         Height          =   630
         Left            =   -74850
         TabIndex        =   43
         Top             =   435
         Width           =   5865
         Begin VB.CommandButton cmdCargar 
            Caption         =   "Cargar"
            Height          =   315
            Left            =   4110
            TabIndex        =   62
            Top             =   195
            Width           =   1590
         End
         Begin MSMask.MaskEdBox mskFinBusqueda 
            Height          =   315
            Left            =   2685
            TabIndex        =   45
            ToolTipText     =   "Fecha final de la búsqueda"
            Top             =   195
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mmm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskInicioBusqueda 
            Height          =   315
            Left            =   780
            TabIndex        =   44
            ToolTipText     =   "Fecha inicial de la búsqueda"
            Top             =   195
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mmm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            Height          =   195
            Left            =   2175
            TabIndex        =   61
            Top             =   255
            Width           =   420
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            Height          =   195
            Left            =   135
            TabIndex        =   60
            Top             =   255
            Width           =   465
         End
      End
      Begin VB.Frame Frame2 
         Height          =   800
         Left            =   8700
         TabIndex        =   33
         Top             =   7785
         Width           =   2945
         Begin VB.CommandButton cmdSave 
            Height          =   495
            Left            =   2385
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmCorte.frx":1E38
            Style           =   1  'Graphical
            TabIndex        =   32
            ToolTipText     =   "Cerrar el corte"
            Top             =   200
            Width           =   465
         End
         Begin VB.CommandButton cmdEnd 
            Height          =   495
            Left            =   1920
            Picture         =   "frmCorte.frx":217A
            Style           =   1  'Graphical
            TabIndex        =   31
            ToolTipText     =   "Ultimo corte"
            Top             =   200
            UseMaskColor    =   -1  'True
            Width           =   465
         End
         Begin VB.CommandButton cmdNext 
            Height          =   495
            Left            =   1455
            Picture         =   "frmCorte.frx":266C
            Style           =   1  'Graphical
            TabIndex        =   30
            ToolTipText     =   "Siguiente corte"
            Top             =   200
            UseMaskColor    =   -1  'True
            Width           =   465
         End
         Begin VB.CommandButton cmdLocate 
            Height          =   495
            Left            =   990
            Picture         =   "frmCorte.frx":2B5E
            Style           =   1  'Graphical
            TabIndex        =   29
            ToolTipText     =   "Búsqueda"
            Top             =   200
            UseMaskColor    =   -1  'True
            Width           =   465
         End
         Begin VB.CommandButton cmdBack 
            Height          =   495
            Left            =   525
            Picture         =   "frmCorte.frx":3050
            Style           =   1  'Graphical
            TabIndex        =   28
            ToolTipText     =   "Anterior corte"
            Top             =   200
            UseMaskColor    =   -1  'True
            Width           =   465
         End
         Begin VB.CommandButton cmdTop 
            Height          =   495
            Left            =   60
            Picture         =   "frmCorte.frx":3542
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "Primer corte"
            Top             =   200
            UseMaskColor    =   -1  'True
            Width           =   465
         End
      End
      Begin VB.ListBox lstFormas 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2700
         ItemData        =   "frmCorte.frx":3944
         Left            =   135
         List            =   "frmCorte.frx":3946
         TabIndex        =   1
         ToolTipText     =   "Detalle del corte"
         Top             =   2430
         Width           =   11490
      End
      Begin VB.Frame Frame1 
         Height          =   4695
         Left            =   135
         TabIndex        =   34
         Top             =   435
         Width           =   11490
         Begin VB.CheckBox chkDesglosaCorte 
            Caption         =   "Desglosar póliza"
            Enabled         =   0   'False
            Height          =   195
            Left            =   9675
            TabIndex        =   79
            ToolTipText     =   "Desglosa la póliza del corte"
            Top             =   1675
            Value           =   1  'Checked
            Width           =   1575
         End
         Begin VB.TextBox txtNumCorte 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1590
            TabIndex        =   0
            ToolTipText     =   "Número de corte"
            Top             =   195
            Width           =   1140
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "al"
            Height          =   195
            Left            =   3750
            TabIndex        =   74
            Top             =   945
            Width           =   120
         End
         Begin VB.Label lblFolio 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   9675
            TabIndex        =   71
            Top             =   1230
            Width           =   1725
         End
         Begin VB.Label lblMes 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   10875
            TabIndex        =   70
            Top             =   885
            Width           =   525
         End
         Begin VB.Label lblEjercicio 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   9675
            TabIndex        =   69
            Top             =   885
            Width           =   725
         End
         Begin VB.Label lblEstatus 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   9675
            TabIndex        =   68
            Top             =   540
            Width           =   1725
         End
         Begin VB.Label lblPersonaCierra 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1590
            TabIndex        =   67
            Top             =   1575
            Width           =   4515
         End
         Begin VB.Label lblPersonaAbre 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1590
            TabIndex        =   66
            Top             =   1230
            Width           =   4515
         End
         Begin VB.Label lblFin 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   4005
            TabIndex        =   65
            Top             =   885
            Width           =   2100
         End
         Begin VB.Label lblInicio 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1590
            TabIndex        =   64
            Top             =   885
            Width           =   2100
         End
         Begin VB.Label lblDepartamento 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1590
            TabIndex        =   63
            Top             =   540
            Width           =   4515
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Tipo corte"
            Height          =   195
            Left            =   8670
            TabIndex        =   59
            Top             =   255
            Width           =   720
         End
         Begin VB.Label lblTipoCorte 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   9675
            TabIndex        =   58
            Top             =   195
            Width           =   1725
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Mes"
            Height          =   195
            Left            =   10480
            TabIndex        =   48
            Top             =   945
            Width           =   300
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Ejercicio"
            Height          =   195
            Left            =   8670
            TabIndex        =   47
            Top             =   945
            Width           =   600
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Persona cierra"
            Height          =   195
            Left            =   135
            TabIndex        =   42
            Top             =   1635
            Width           =   1020
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Persona abre"
            Height          =   195
            Left            =   135
            TabIndex        =   41
            Top             =   1290
            Width           =   945
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Periodo"
            Height          =   195
            Left            =   135
            TabIndex        =   40
            Top             =   945
            Width           =   540
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Folio"
            Height          =   195
            Left            =   8670
            TabIndex        =   39
            Top             =   1290
            Width           =   330
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Departamento"
            Height          =   195
            Left            =   135
            TabIndex        =   38
            Top             =   600
            Width           =   1005
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Estado"
            Height          =   195
            Left            =   8670
            TabIndex        =   37
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Número"
            Height          =   195
            Left            =   135
            TabIndex        =   36
            Top             =   255
            Width           =   555
         End
      End
      Begin VB.Label lblTotalDolares 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9780
         TabIndex        =   73
         Top             =   7365
         Width           =   1845
      End
      Begin VB.Label lblTotalPesos 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9780
         TabIndex        =   72
         Top             =   7005
         Width           =   1845
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Total dólares"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   8340
         TabIndex        =   50
         Top             =   7402
         Width           =   1410
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Total pesos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   8340
         TabIndex        =   49
         Top             =   7042
         Width           =   1260
      End
   End
End
Attribute VB_Name = "frmCorte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
' Forma para registro y consulta del corte de caja, se muestran los cortes de todas
' las cajas, se puede cerrar el corte de una caja que no es del departamento que
' consulta, con advertencia.
' Fecha de programación: Jueves 1 de Febrero de 2001
'------------------------------------------------------------------------------------

'Columnas de la búsqueda:
Const cintColInicio = 1
Const cintColCierre = 2
Const cintColNumero = 3
Const cintColDepartamento = 4
Const cintColEmpleado = 5
Const cintColTipoCorte = 6
Const cIntColEstado = 7
Const cintColEmpleadoCierra = 8
Const cintColumnas = 9

Private vgrptReporte As CRAXDRT.Report

Private Type RegistroPoliza
    vllngNumeroCuenta As Long
    vldblCantidadMovimiento As Double
    vlintTipoMovimiento As Integer
    lstrFolioDocumento As String
    lstrDescSalidaCajaChica As String
End Type

Dim apoliza() As RegistroPoliza

Dim rs As New ADODB.Recordset                   'Varios usos
Dim ldtmFecha As Date                           'Fecha actual
Dim llngCveEmpleado As Long                     'Empleado que abrió el corte
Dim lintCveDepto As Integer                     'Clave del departamento
Dim lstrTipoCorte As String                     'Tipo de corte, P = ingresos, C = caja chica
Dim lintTipoCorte As Integer                    'Tipo de corte que se maneja por departamento o por empleado y departamento
Dim lblnCorte As Boolean                        'Para saber si se encontró la información del corte cuando el número fue introducido manualmente
Dim lstrConceptoDetallePoliza As String
Dim lstrEstado As String                        'Para saber el estado del corte, A = abierto, C = cerrado
Dim vlstrSentencia As String
Dim vllngNumPoliza As Long
Dim lblnMensaje As Boolean                      'Para saber cuando mostrar el mensaje de que no existe información en la consulta
Dim vlblnDesglosaPolizaConfig As Boolean        'Configuración de si los cortes del departamento desglosarán la póliza
Private lstrFechaInicio As String
Private lstrFechaFin As String
Dim lblnValidarFacturaSinXml As Boolean         'Indica si se debe validar que se hayan asignado los xmls a las facturas de caja chica

Private Sub pHabilita(intTop As Integer, intBack As Integer, intlocate As Integer, intNext As Integer, intEnd As Integer, intSave As Integer, intMovimientos As Integer, intCronologico As Integer, intFormaPago As Integer, intPoliza As Integer)
On Error GoTo NotificaError
    
    cmdTop.Enabled = intTop = 1
    cmdBack.Enabled = intBack = 1
    cmdLocate.Enabled = intlocate = 1
    cmdNext.Enabled = intNext = 1
    cmdEnd.Enabled = intEnd = 1
    cmdSave.Enabled = intSave = 1
    cmdCorteporMovimiento.Enabled = intMovimientos = 1
    cmdCorteCronologico.Enabled = intCronologico = 1
    cmdCorteFormasPago.Enabled = intFormaPago = 1
    cmdPrintPoliza.Enabled = intPoliza = 1

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pHabilita"))
End Sub

Private Sub chkAcumulado_Click()
On Error GoTo NotificaError
    
    pMuestraCorte Val(txtNumCorte.Text), 0, Format(ldtmFecha, "dd/mm/yyyy"), Format(ldtmFecha, "dd/mm/yyyy"), -1, -1

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":chkAcumulado_Click"))
End Sub

Private Sub chkIntermedios_Click()
    '|  Inicialización de los filtros
    DtpFechaInicio.Value = fdtmServerFecha
    dtpFechaFin.Value = fdtmServerFecha
    mskHoraIni.Text = "00:00"
    mskHoraFin.Text = Format(fdtmServerHora, "hh:mm")
    cboEmpleadoRegistro.ListIndex = 0
    chkAgrupadoPorEmpleado.Value = 0
    
    tmrHora.Enabled = True
    
    cmdCorteCronologico.Enabled = chkIntermedios.Value = 1 Or Len(Trim(lblPersonaCierra.Caption)) <> 0
    cmdCorteFormasPago.Enabled = chkIntermedios.Value = 1 Or Len(Trim(lblPersonaCierra.Caption)) <> 0
    cmdCorteporMovimiento.Enabled = chkIntermedios.Value = 1 Or Len(Trim(lblPersonaCierra.Caption)) <> 0

    freRangos.Enabled = chkIntermedios.Value = 1
    DtpFechaInicio.Enabled = chkIntermedios.Value = 1
    mskHoraIni.Enabled = chkIntermedios.Value = 1
    lblA.Enabled = chkIntermedios.Value = 1
    dtpFechaFin.Enabled = chkIntermedios.Value = 1
    mskHoraFin.Enabled = chkIntermedios.Value = 1
    lblEmpleadoRegistro.Enabled = chkIntermedios.Value = 1
    cboEmpleadoRegistro.Enabled = chkIntermedios.Value = 1
    chkAgrupadoPorEmpleado.Enabled = chkIntermedios.Value = 1
End Sub

Private Sub cmdBack_Click()
On Error GoTo NotificaError
    
    If grdCortes.Row > 1 Then
        grdCortes.Row = grdCortes.Row - 1
    End If
            
    pMuestraCorte Val(grdCortes.TextMatrix(grdCortes.Row, cintColNumero)), 0, Format(ldtmFecha, "dd/mm/yyyy"), Format(ldtmFecha, "dd/mm/yyyy"), vgintNumeroDepartamento, IIf(lintTipoCorte = 1, -1, vglngNumeroEmpleado)
    pHabilita 1, 1, 1, 1, 1, IIf(lstrEstado = "A", 1, 0), IIf(lstrEstado = "C" Or chkIntermedios.Value = 1, 1, 0), IIf(lstrEstado = "C" Or chkIntermedios.Value = 1, 1, 0), IIf(lstrEstado = "C" Or chkIntermedios.Value = 1, 1, 0), IIf(lstrEstado = "C", 1, 0)
    pHabilitaFueraCorte

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdBack_Click"))
End Sub

Private Sub cmdCargar_Click()
On Error GoTo NotificaError
    
    Dim intcontador As Integer
    
    With grdCortes
        .Cols = cintColumnas
        .Rows = 2
        .FixedRows = 1
        .FixedCols = 1
        .FormatString = "|Fecha apertura|Fecha cierre|Número|Departamento|Persona abre|Tipo corte|Estado|Persona cierra"
        
        For intcontador = 1 To .Cols - 1
            .TextMatrix(1, intcontador) = ""
        Next intcontador
        
        .ColWidth(0) = 100
        .ColWidth(cintColInicio) = 1500
        .ColWidth(cintColCierre) = 1500
        .ColWidth(cintColNumero) = 900
        .ColWidth(cintColDepartamento) = 1900
        .ColWidth(cintColEmpleado) = 3000
        .ColWidth(cintColTipoCorte) = 1500
        .ColWidth(cIntColEstado) = 1000
        .ColWidth(cintColEmpleadoCierra) = 3000
        
        .ColAlignment(cintColInicio) = flexAlignLeftCenter
        .ColAlignment(cintColCierre) = flexAlignLeftCenter
        .ColAlignment(cintColNumero) = flexAlignRightCenter
        .ColAlignment(cintColDepartamento) = flexAlignLeftCenter
        .ColAlignment(cintColEmpleado) = flexAlignLeftCenter
        .ColAlignment(cintColTipoCorte) = flexAlignLeftCenter
        .ColAlignment(cIntColEstado) = flexAlignLeftCenter
        .ColAlignment(cintColEmpleadoCierra) = flexAlignLeftCenter
        
        .ColAlignmentFixed(cintColInicio) = flexAlignCenterCenter
        .ColAlignmentFixed(cintColCierre) = flexAlignCenterCenter
        .ColAlignmentFixed(cintColNumero) = flexAlignCenterCenter
        .ColAlignmentFixed(cintColDepartamento) = flexAlignCenterCenter
        .ColAlignmentFixed(cintColEmpleado) = flexAlignCenterCenter
        .ColAlignmentFixed(cintColTipoCorte) = flexAlignCenterCenter
        .ColAlignmentFixed(cIntColEstado) = flexAlignCenterCenter
        .ColAlignmentFixed(cintColEmpleadoCierra) = flexAlignCenterCenter
        
        vgstrParametrosSP = "-1" _
                            & "|" & "1" _
                            & "|" & fstrFechaSQL(mskInicioBusqueda.Text) _
                            & "|" & fstrFechaSQL(mskFinBusqueda.Text) _
                            & "|" & CStr(vgintNumeroDepartamento) _
                            & "|" & IIf(lintTipoCorte = 1, -1, vglngNumeroEmpleado)
        Set rs = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELCORTE")
        If rs.RecordCount <> 0 Then
            Do While Not rs.EOF
                .TextMatrix(.Rows - 1, cintColInicio) = Format(rs!FechaAbre, "dd/mmm/yyyy hh:mm:ss")
                .TextMatrix(.Rows - 1, cintColCierre) = IIf(IsNull(rs!FechaCierra), " ", Format(rs!FechaCierra, "dd/mmm/yyyy hh:mm:ss"))
                .TextMatrix(.Rows - 1, cintColNumero) = rs!IdCorte
                .TextMatrix(.Rows - 1, cintColDepartamento) = rs!departamento
                .TextMatrix(.Rows - 1, cintColEmpleado) = rs!PersonaAbre
                .TextMatrix(.Rows - 1, cintColTipoCorte) = UCase(rs!TipoCorteDescripcion)
                .TextMatrix(.Rows - 1, cIntColEstado) = UCase(rs!Estado)
                .TextMatrix(.Rows - 1, cintColEmpleadoCierra) = rs!PersonaCierra
                .Rows = .Rows + 1
                rs.MoveNext
            Loop
            .Rows = .Rows - 1
        Else
            If lblnMensaje Then
                'No existe información con esos parámetros.
                MsgBox SIHOMsg(236), vbInformation + vbOKOnly, "Mensaje"
            Else
                lblnMensaje = True
            End If
        End If
    End With
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCargaCortes"))
End Sub

Private Sub cmdCorteCronologico_Click()
On Error GoTo NotificaError

    Dim alstrParametros(15) As String
    Dim rsPvSelCorteCronologico As New ADODB.Recordset
    Dim rsPvSelCorteResumenDocumentos As New ADODB.Recordset

    Dim vlstrFechaIni As String
    Dim vlstrFechaFin As String
    Dim vlstrResumenDocumentos As String
    Dim vlstrResumenDocumentosCxC As String
    
    Dim vllngCuantosFacturas As Long
    Dim vllngCuantosRecibos As Long
    Dim vllngCuantosTickets As Long
    Dim vllngCuantosSalidas As Long
    Dim vllngCuantosFacturasCanceladas As Long
    Dim vllngCuantosRecibosCancelados As Long
    Dim vllngCuantosTicketsCancelados As Long
    Dim vllngCuantosSalidasCancelados As Long
    Dim rsDevolucines As New ADODB.Recordset
    Dim ldblTotalDevoluciones As Double
    Dim lstrDevolucionesCxP As String
    Dim lstrImportesDevolucion As String
    
    If chkIntermedios.Value = 1 Then
        vlstrFechaIni = fstrFechaSQL(DtpFechaInicio.Value, mskHoraIni & ":00", True)
        vlstrFechaFin = fstrFechaSQL(dtpFechaFin.Value, mskHoraFin & ":00", True)
    Else
        vlstrFechaIni = "1900/01/01"
        vlstrFechaFin = "1900/01/01"
    End If
        
    vgstrParametrosSP = IIf(chkIntermedios.Value = 1, 0, Val(txtNumCorte.Text)) & "|" & _
                        IIf(chkIntermedios.Value = 1, 1, 0) & "|" & vlstrFechaIni & "|" & _
                        vlstrFechaFin & "|" & IIf(chkIntermedios.Value = 1, vgintNumeroDepartamento, 0) & "|" & _
                        chkFacturas.Value & "|" & chkPagos.Value & "|" & _
                        chkTickets.Value & "|" & chkSalidas.Value & "|" & _
                        chkPagosCredito.Value & "|" & chkHonorarios.Value & "|" & chkFondoFijo.Value & "|" & chkTransferencias.Value & "|" & _
                        chkSalidaChica.Value & "|" & chkEntradaChica.Value & "|" & chkSoloCancelados.Value & "|"
    vgstrParametrosSP = vgstrParametrosSP & _
                        IIf(cboEmpleadoRegistro.ListIndex = 0, -1, cboEmpleadoRegistro.ItemData(cboEmpleadoRegistro.ListIndex)) & "|" & _
                        chkAgrupadoPorEmpleado.Value
    Set rsPvSelCorteCronologico = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELCORTECRONOLOGICO")
    If rsPvSelCorteCronologico.EOF Then
        MsgBox SIHOMsg(13), vbInformation, "Mensaje"
    Else
        'devoluciones de pacientes por el proceso de cXp
        vgstrParametrosSP = IIf(chkIntermedios.Value = 1, 0, Val(txtNumCorte.Text)) & "|" & _
                        IIf(chkIntermedios.Value = 1, 1, 0) & "|" & vlstrFechaIni & "|" & _
                        vlstrFechaFin & "|" & IIf(chkIntermedios.Value = 1, vgintNumeroDepartamento, 0)
        Set rsDevolucines = frsEjecuta_SP(vgstrParametrosSP, "sp_pvSelDevolucionesPacCxP")
        Do While Not rsDevolucines.EOF
            lstrDevolucionesCxP = lstrDevolucionesCxP & rsDevolucines!FolioFactura & Chr(13)
            lstrImportesDevolucion = lstrImportesDevolucion & FormatCurrency(rsDevolucines!Importe, 2) & Chr(13)
            ldblTotalDevoluciones = ldblTotalDevoluciones + rsDevolucines!Importe
            rsDevolucines.MoveNext
        Loop
        'lstrDevolucionesCxP = LSET("FOLFACTURA",20)
        pInstanciaReporte vgrptReporte, "rptCorteCronologico.rpt"
        vgrptReporte.DiscardSavedData
        alstrParametros(0) = "Departamento;" & lblDepartamento.Caption
        alstrParametros(1) = "Ejercicio;" & lblEjercicio.Caption
        alstrParametros(2) = "Fecha;" & fdtmServerFecha
        alstrParametros(3) = "FechaFin;" & lstrFechaFin
        alstrParametros(4) = "FechaInicio;" & lstrFechaInicio
        alstrParametros(5) = "Hora;" & fdtmServerHora
        alstrParametros(6) = "Mes;" & lblMes.Caption
        alstrParametros(7) = "NombreHospital;" & Trim(vgstrNombreHospitalCH)
        alstrParametros(8) = "NumeroCorte;" & txtNumCorte.Text
        alstrParametros(9) = "NumeroPoliza;" & lblFolio.Caption
        alstrParametros(10) = "PersonaCierra;" & lblPersonaCierra.Caption
        alstrParametros(11) = "PersonaInicia;" & lblPersonaAbre.Caption
        alstrParametros(12) = "TipoCorte;" & lblTipoCorte.Caption
        alstrParametros(13) = "foliosFactura;" & lstrDevolucionesCxP
        alstrParametros(14) = "importesDevolucion;" & lstrImportesDevolucion
        alstrParametros(15) = "TotalDevoluciones;" & ldblTotalDevoluciones
        pCargaParameterFields alstrParametros, vgrptReporte
        pImprimeReporte vgrptReporte, rsPvSelCorteCronologico, "P", "Corte cronológico"
    End If
    rsPvSelCorteCronologico.Close

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdCorteCronologico_Click"))
End Sub

Private Sub cmdCorteFormasPago_Click()
On Error GoTo NotificaError

    Dim alstrParametros(12) As String
    Dim rsPvSelCorteFormasPago As New ADODB.Recordset
    
    Dim vlstrFechaIni As String
    Dim vlstrFechaFin As String
    
    If chkIntermedios.Value = 1 Then
        vlstrFechaIni = fstrFechaSQL(DtpFechaInicio.Value, mskHoraIni & ":00", True)
        vlstrFechaFin = fstrFechaSQL(dtpFechaFin.Value, mskHoraFin & ":00", True)
    Else
        vlstrFechaIni = "1900/01/01"
        vlstrFechaFin = "1900/01/01"
    End If
    
    vgstrParametrosSP = IIf(chkIntermedios.Value = 1, 0, Val(txtNumCorte.Text)) & "|" & _
                        IIf(chkIntermedios.Value = 1, 1, 0) & "|" & _
                        vlstrFechaIni & "|" & _
                        vlstrFechaFin & "|" & _
                        IIf(chkIntermedios.Value = 1, vgintNumeroDepartamento, 0) & "|" & _
                        chkFacturas.Value & "|" & _
                        chkPagos.Value & "|" & _
                        chkTickets.Value & "|" & _
                        chkSalidas.Value & "|" & _
                        chkPagosCredito.Value & "|" & _
                        chkHonorarios.Value & "|" & _
                        chkFondoFijo.Value & "|" & _
                        chkTransferencias.Value & "|" & _
                        chkSalidaChica.Value & "|" & _
                        chkEntradaChica.Value & "|" & _
                        chkSoloCancelados.Value & "|"
    vgstrParametrosSP = vgstrParametrosSP & _
                        IIf(cboEmpleadoRegistro.ListIndex = 0, -1, cboEmpleadoRegistro.ItemData(cboEmpleadoRegistro.ListIndex)) & "|" & _
                        chkAgrupadoPorEmpleado.Value
    Set rsPvSelCorteFormasPago = frsEjecuta_SP(vgstrParametrosSP, "Sp_PvSelCorteFormasPago")
    If rsPvSelCorteFormasPago.EOF Then
        MsgBox SIHOMsg(13), vbInformation, "Mensaje"
    Else
        pInstanciaReporte vgrptReporte, "rptCorteFormaspago.rpt"
        vgrptReporte.DiscardSavedData
        alstrParametros(0) = "Departamento;" & lblDepartamento.Caption
        alstrParametros(1) = "Ejercicio;" & lblEjercicio.Caption
        alstrParametros(2) = "Fecha;" & ldtmFecha
        alstrParametros(3) = "FechaFin;" & lblFin.Caption
        alstrParametros(4) = "FechaInicio;" & lblInicio.Caption
        alstrParametros(5) = "Hora;" & fdtmServerHora
        alstrParametros(6) = "Mes;" & lblMes.Caption
        alstrParametros(7) = "NombreHospital;" & Trim(vgstrNombreHospitalCH)
        alstrParametros(8) = "NumeroCorte;" & txtNumCorte.Text
        alstrParametros(9) = "NumeroPoliza;" & lblFolio.Caption
        alstrParametros(10) = "PersonaCierra;" & lblPersonaCierra.Caption
        alstrParametros(11) = "PersonaInicia;" & lblPersonaAbre.Caption
        alstrParametros(12) = "TipoCorte;" & lblTipoCorte.Caption
        pCargaParameterFields alstrParametros, vgrptReporte
        pImprimeReporte vgrptReporte, rsPvSelCorteFormasPago, "P", "Formas de pago"
    End If
    rsPvSelCorteFormasPago.Close

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdCorteFormasPago_Click"))
End Sub

Private Sub cmdCorteporMovimiento_Click()
On Error GoTo NotificaError
    
    Dim rsPvSelCorteMovimiento As New ADODB.Recordset
    Dim alstrParametros(13) As String
    Dim vlstrFechaIni As String
    Dim vlstrFechaFin As String
    Dim vldtDateIni As Date
    Dim vldtDateFin As Date
        
    If chkIntermedios.Value = 1 Then
        vlstrFechaIni = fstrFechaSQL(DtpFechaInicio.Value, mskHoraIni & ":00", True)
        vlstrFechaFin = fstrFechaSQL(dtpFechaFin.Value, mskHoraFin & ":00", True)
    Else
        vlstrFechaIni = fstrFechaSQL("01/01/1900 00:00:00")
        vlstrFechaFin = fstrFechaSQL("01/01/1900 00:00:00")
    End If
 
    vgstrParametrosSP = IIf(chkIntermedios.Value = 1, 0, Val(txtNumCorte.Text)) & "|" & _
                        IIf(chkIntermedios.Value = 1, 1, 0) & "|" & _
                        vlstrFechaIni & "|" & _
                        vlstrFechaFin & "|" & _
                        IIf(chkIntermedios.Value = 1, vgintNumeroDepartamento, 0) & "|" & _
                        chkFacturas.Value & "|" & _
                        chkPagos.Value & "|" & _
                        chkTickets.Value & "|" & _
                        chkSalidas.Value & "|" & _
                        chkSoloCancelados.Value & "|" & _
                        chkPagosCredito.Value & "|" & _
                        chkHonorarios.Value & "|" & _
                        chkFondoFijo.Value & "|" & _
                        chkTransferencias.Value & "|" & _
                        chkSalidaChica.Value & "|" & _
                        chkEntradaChica.Value & "|" & _
                        IIf(optGrupoFormaPago.Value, 1, 2) & "|"
    vgstrParametrosSP = vgstrParametrosSP & _
                        IIf(cboEmpleadoRegistro.ListIndex = 0, -1, cboEmpleadoRegistro.ItemData(cboEmpleadoRegistro.ListIndex)) & "|" & _
                        chkAgrupadoPorEmpleado.Value & "|" & _
                        vgintClaveEmpresaContable
    Set rsPvSelCorteMovimiento = frsEjecuta_SP(vgstrParametrosSP, "Sp_PvSelCorteMovimiento")
    Set rsPvSelCorteMovimiento = frsUltimoRecordset(rsPvSelCorteMovimiento)
    If rsPvSelCorteMovimiento.EOF Then
        MsgBox SIHOMsg(13), vbInformation, "Mensaje"
    Else
        pInstanciaReporte vgrptReporte, "rptDetalleMovimientos.rpt"
        vgrptReporte.DiscardSavedData
        alstrParametros(0) = "FechaInicio;" & lstrFechaInicio
        alstrParametros(1) = "FechaFin;" & lstrFechaFin
        alstrParametros(2) = "NombreHospital;" & Trim(vgstrNombreHospitalCH)
        alstrParametros(3) = "NumCorte;" & Val(txtNumCorte.Text)
        alstrParametros(4) = "Titulo;" & "CORTE POR MOVIMIENTOS"
        alstrParametros(5) = "Departamento;" & lblDepartamento.Caption
        alstrParametros(6) = "PersonaAbre;" & lblPersonaAbre.Caption
        alstrParametros(7) = "PersonaCierra;" & lblPersonaCierra.Caption
        alstrParametros(8) = "Ejercicio;" & lblEjercicio.Caption
        alstrParametros(9) = "Mes;" & lblMes.Caption
        alstrParametros(10) = "Folio;" & lblFolio.Caption
        alstrParametros(11) = "Acumulado;" & chkAcumulado.Value
        alstrParametros(12) = "TipoCorte;" & lblTipoCorte.Caption
        alstrParametros(13) = "TipoCorteCaracter;" & lstrTipoCorte
        
        pCargaParameterFields alstrParametros, vgrptReporte
        pImprimeReporte vgrptReporte, rsPvSelCorteMovimiento, "P", "Corte por movimientos"
    End If
    rsPvSelCorteMovimiento.Close

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdCorteporMovimiento_Click"))
End Sub

Private Sub cmdEnd_Click()
On Error GoTo NotificaError
    
    grdCortes.Row = grdCortes.Rows - 1
    pMuestraCorte Val(grdCortes.TextMatrix(grdCortes.Row, cintColNumero)), 0, Format(ldtmFecha, "dd/mm/yyyy"), Format(ldtmFecha, "dd/mm/yyyy"), vgintNumeroDepartamento, IIf(lintTipoCorte = 1, -1, vglngNumeroEmpleado)
    pHabilita 1, 1, 1, 1, 1, IIf(lstrEstado = "A", 1, 0), IIf(lstrEstado = "C" Or chkIntermedios.Value = 1, 1, 0), IIf(lstrEstado = "C" Or chkIntermedios.Value = 1, 1, 0), IIf(lstrEstado = "C" Or chkIntermedios.Value = 1, 1, 0), IIf(lstrEstado = "C", 1, 0)
    pHabilitaFueraCorte

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdEnd_Click"))
End Sub

Private Sub cmdLocate_Click()
On Error GoTo NotificaError
    
    SSTabCorte.Tab = 1
    
    mskInicioBusqueda.SetFocus
    lblnMensaje = False
    cmdCargar_Click

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdLocate_Click"))
End Sub

Private Sub cmdMovsFueraCorte_Click()
'***** Reporte de movimientos fuera del corte *****'
On Error GoTo NotificaError
    
    Dim rsPvSelMovFueraCorte As New ADODB.Recordset
    Dim alstrParametros(13) As String
    Dim vlstrFechaIni As String
    Dim vlstrFechaFin As String
    Dim vldtDateIni As Date
    Dim vldtDateFin As Date
        
    vlstrFechaIni = fstrFechaSQL(lblInicio.Caption, mskHoraIni & ":00", True)
    vlstrFechaFin = fstrFechaSQL(dtpFechaFin.Value, mskHoraFin & ":00", True)
 
    vgstrParametrosSP = Val(txtNumCorte.Text) & "|" & _
                        vlstrFechaIni & "|" & _
                        vlstrFechaFin & "|" & _
                        vgintNumeroDepartamento & "|" & _
                        chkFacturas.Value & "|" & _
                        chkPagos.Value & "|" & _
                        chkPagos.Value & "|" & _
                        1 & "|" & _
                        chkSoloCancelados.Value & "|" & _
                        0 & "|" & _
                        IIf(optGrupoFormaPago.Value, 1, 2) & "|" & _
                        -1 & "|" & _
                        chkAgrupadoPorEmpleado.Value & "|" & _
                        Str(vgintClaveEmpresaContable)
    Set rsPvSelMovFueraCorte = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELMOVIMIENTOFUERACORTE")
    Set rsPvSelMovFueraCorte = frsUltimoRecordset(rsPvSelMovFueraCorte)
    If rsPvSelMovFueraCorte.EOF Then
        MsgBox SIHOMsg(13), vbInformation, "Mensaje"
    Else
        pInstanciaReporte vgrptReporte, "rptDetalleMovFueraCorte.rpt"
        vgrptReporte.DiscardSavedData
        alstrParametros(0) = "FechaInicio;" & Format(lblInicio.Caption, "dd/mmm/yyyy hh:mm:ss am/pm")
        alstrParametros(1) = "FechaFin;" & Format(lblFin.Caption, "dd/mmm/yyyy hh:mm:ss am/pm")
        alstrParametros(2) = "NombreHospital;" & Trim(vgstrNombreHospitalCH)
        alstrParametros(3) = "NumCorte;" & Val(txtNumCorte.Text)
        alstrParametros(4) = "Titulo;" & "MOVIMIENTOS FUERA DEL CORTE"
        alstrParametros(5) = "Departamento;" & lblDepartamento.Caption
        alstrParametros(6) = "PersonaAbre;" & lblPersonaAbre.Caption
        alstrParametros(7) = "PersonaCierra;" & lblPersonaCierra.Caption
        alstrParametros(8) = "Ejercicio;" & lblEjercicio.Caption
        alstrParametros(9) = "Mes;" & lblMes.Caption
        alstrParametros(10) = "Folio;" & lblFolio.Caption
        alstrParametros(11) = "Acumulado;" & chkAcumulado.Value
        alstrParametros(12) = "TipoCorte;" & lblTipoCorte.Caption
        alstrParametros(13) = "TipoCorteCaracter;" & lstrTipoCorte
        
        pCargaParameterFields alstrParametros, vgrptReporte
        pImprimeReporte vgrptReporte, rsPvSelMovFueraCorte, "P", "Detalle de movimientos fuera del corte"
    End If
    rsPvSelMovFueraCorte.Close

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdMovsFueraCorte_Click"))
End Sub

Private Sub cmdNext_Click()
On Error GoTo NotificaError
    
    If grdCortes.Row < grdCortes.Rows - 1 Then
        grdCortes.Row = grdCortes.Row + 1
    End If
        
    pMuestraCorte Val(grdCortes.TextMatrix(grdCortes.Row, cintColNumero)), 0, Format(ldtmFecha, "dd/mm/yyyy"), Format(ldtmFecha, "dd/mm/yyyy"), vgintNumeroDepartamento, IIf(lintTipoCorte = 1, -1, vglngNumeroEmpleado)
    pHabilita 1, 1, 1, 1, 1, IIf(lstrEstado = "A", 1, 0), IIf(lstrEstado = "C" Or chkIntermedios.Value = 1, 1, 0), IIf(lstrEstado = "C" Or chkIntermedios.Value = 1, 1, 0), IIf(lstrEstado = "C" Or chkIntermedios.Value = 1, 1, 0), IIf(lstrEstado = "C", 1, 0)
    pHabilitaFueraCorte
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdNext_Click"))
End Sub

Private Sub cmdPolizaFueraCorte_Click()
'***** Pólizas de movimientos fuera del corte *****'
On Error GoTo NotificaError

    frmConsultaPolizasDepartamento.vllngNumCorte = CLng(txtNumCorte.Text)
    frmConsultaPolizasDepartamento.vllngNumOpcion = IIf(cgstrModulo = "PV", 314, 630)
    frmConsultaPolizasDepartamento.vlblnPolizasFueraCorte = True
    frmConsultaPolizasDepartamento.vldtmFechaIni = CDate(lblInicio.Caption)
    If Trim(lblFin.Caption) <> "" Then
        frmConsultaPolizasDepartamento.vldtmFechaFin = CDate(lblFin.Caption)
    Else
        frmConsultaPolizasDepartamento.vldtmFechaFin = CDate(dtpFechaFin.Value)
    End If
    frmConsultaPolizasDepartamento.Show vbModal
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdPolizaFueraCorte_Click"))
End Sub

Private Sub cmdPrintPoliza_Click()
On Error GoTo NotificaError
    
    pImprimePoliza Str(vllngNumPoliza), "P", 1, 3
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdPrintPoliza_Click"))
End Sub

Private Sub cmdSave_Click()
On Error GoTo NotificaError
    
    '------------------------------------'
    '|   Verifica que tenga permisos:   |'
    '------------------------------------'
    If fblnRevisaPermiso(vglngNumeroLogin, cintNumOpcionCorte, "E") Then
        '----------------------------------------------------------------------'
        '|   Verifica si hay cuentas liquidadas que no han sido facturadas:   |'
        '----------------------------------------------------------------------'
        vgstrParametrosSP = fstrFechaSQL(Format(ldtmFecha - 1, "dd/mm/yyyy")) & "|" & fstrFechaSQL(Format(ldtmFecha - 1, "dd/mm/yyyy")) & "|" & Str(vgintNumeroDepartamento)
        Set rs = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELCUENTALIQUIDADA")
        If rs!Total > 0 Then
            '|  Realice la facturación de grupo antes de cerrar el corte.
            MsgBox SIHOMsg(790), vbCritical, "Mensaje"
        Else
            '-----------------------------------------'
            '|   Verifica que existan movimientos:   |'
            '-----------------------------------------'
            vlstrSentencia = "SELECT NVL(COUNT(*),0) FROM PvCortePoliza WHERE intNumCorte = " + Trim(txtNumCorte.Text)
            If frsRegresaRs(vlstrSentencia).Fields(0) = 0 Then 'Si hay documentos registrados en este corte?
                '¿Desea actualizar la fecha del corte?
                If MsgBox(SIHOMsg(1651), vbQuestion + vbYesNo, "Mensaje") = vbYes Then
'                    frsEjecuta_SP txtNumCorte.Text, "SP_PVUPDFECHACORTE"
                    pCierraCorte
                Else
                    If MsgBox(SIHOMsg(1652), vbQuestion + vbYesNo, "Mensaje") = vbYes Then
                        frsEjecuta_SP txtNumCorte.Text, "SP_PVUPDFECHACORTE"
                        txtNumCorte.SetFocus
                    End If
                End If
            Else
                pCierraCorte
            End If
        End If
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdSave_Click"))
End Sub

Private Sub pIncluyeMovimiento(vllngxNumeroCuenta As Long, vldblxCantidad As Double, vlintxTipoMovto As Integer, lstrFolio As String, lstrDescSalidaCajaChica As String)
On Error GoTo NotificaError

'   10/11/2014 Se modifica para que no agrupe ni totalice las cuentas, esto para la contabilidad electrónica
    
    Dim vlblnEstaCuenta As Boolean
    Dim vllngContador As Long
        
    If chkDesglosaCorte.Value = 1 Then
        ' Desglosará el detalle de la póliza
        If apoliza(0).vllngNumeroCuenta = 0 Then
            apoliza(0).vllngNumeroCuenta = vllngxNumeroCuenta
            apoliza(0).vldblCantidadMovimiento = vldblxCantidad
            apoliza(0).vlintTipoMovimiento = vlintxTipoMovto
            apoliza(0).lstrFolioDocumento = lstrFolio
            apoliza(0).lstrDescSalidaCajaChica = lstrDescSalidaCajaChica
        Else
            ReDim Preserve apoliza(UBound(apoliza, 1) + 1)
            apoliza(UBound(apoliza, 1)).vllngNumeroCuenta = vllngxNumeroCuenta
            apoliza(UBound(apoliza, 1)).vldblCantidadMovimiento = vldblxCantidad
            apoliza(UBound(apoliza, 1)).vlintTipoMovimiento = vlintxTipoMovto
            apoliza(UBound(apoliza, 1)).lstrFolioDocumento = lstrFolio
            apoliza(UBound(apoliza, 1)).lstrDescSalidaCajaChica = lstrDescSalidaCajaChica
        End If
    Else
        ' Agrupará por cuenta el detalle de la póliza
        If apoliza(0).vllngNumeroCuenta = 0 Then
            apoliza(0).vllngNumeroCuenta = vllngxNumeroCuenta
            apoliza(0).vldblCantidadMovimiento = vldblxCantidad
            apoliza(0).vlintTipoMovimiento = vlintxTipoMovto
        Else
            vlblnEstaCuenta = False
            vllngContador = 0
            Do While vllngContador <= UBound(apoliza, 1) And Not vlblnEstaCuenta
                If apoliza(vllngContador).vllngNumeroCuenta = vllngxNumeroCuenta And apoliza(vllngContador).vlintTipoMovimiento = vlintxTipoMovto Then
                    vlblnEstaCuenta = True
                End If
                If Not vlblnEstaCuenta Then
                    vllngContador = vllngContador + 1
                End If
            Loop
    
            If vlblnEstaCuenta Then
                apoliza(vllngContador).vldblCantidadMovimiento = apoliza(vllngContador).vldblCantidadMovimiento + vldblxCantidad
            Else
                ReDim Preserve apoliza(UBound(apoliza, 1) + 1)
                apoliza(UBound(apoliza, 1)).vllngNumeroCuenta = vllngxNumeroCuenta
                apoliza(UBound(apoliza, 1)).vldblCantidadMovimiento = vldblxCantidad
                apoliza(UBound(apoliza, 1)).vlintTipoMovimiento = vlintxTipoMovto
            End If
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pIncluyeMovimiento"))
End Sub

Private Sub pLiberaCierre()
On Error GoTo NotificaError
    
    vlstrSentencia = "UPDATE CnEstatusCierre SET vchEstatus = 'Libre' WHERE tnyClaveEmpresa=" + Str(vgintClaveEmpresaContable)
    pEjecutaSentencia vlstrSentencia

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLiberaCierre"))
End Sub

Private Sub cmdTop_Click()
On Error GoTo NotificaError

    grdCortes.Row = 1
    
    pMuestraCorte Val(grdCortes.TextMatrix(grdCortes.Row, cintColNumero)), 0, Format(ldtmFecha, "dd/mm/yyyy"), Format(ldtmFecha, "dd/mm/yyyy"), vgintNumeroDepartamento, IIf(lintTipoCorte = 1, -1, vglngNumeroEmpleado)
    pHabilita 1, 1, 1, 1, 1, IIf(lstrEstado = "A", 1, 0), IIf(lstrEstado = "C" Or chkIntermedios.Value = 1, 1, 0), IIf(lstrEstado = "C" Or chkIntermedios.Value = 1, 1, 0), IIf(lstrEstado = "C" Or chkIntermedios.Value = 1, 1, 0), IIf(lstrEstado = "C", 1, 0)
    pHabilitaFueraCorte

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdTop_Click"))
End Sub

Private Sub Form_Activate()
On Error GoTo NotificaError
    
    vgstrNombreForm = Me.Name
    tmrHora.Enabled = True
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Activate"))
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    
    If KeyAscii = 27 Then
        If SSTabCorte.Tab = 1 Then
            SSTabCorte.Tab = 0
            txtNumCorte.SetFocus
        Else
            Unload Me
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyPress"))
End Sub

Private Sub Form_Load()
On Error GoTo NotificaError
    Dim rsTemp As New ADODB.Recordset
      
    Me.Icon = frmMenuPrincipal.Icon
     
    ldtmFecha = fdtmServerFecha
    
    Set rs = frsEjecuta_SP(CStr(vgintNumeroDepartamento), "SP_GNSELTIPOCORTE")
    If rs.RecordCount > 0 Then
        lintTipoCorte = rs!intTipoCorte
        vlblnDesglosaPolizaConfig = IIf(rs!INTDESGLOSAPOLIZACORTE = 1, True, False)
    Else
        frsEjecuta_SP CStr(vgintNumeroDepartamento) & "|" & CStr(vglngNumeroEmpleado), "SP_GNINSCORTE"
        vlblnDesglosaPolizaConfig = True
    End If
    
    Set rs = frsEjecuta_SP(CStr(vgintNumeroDepartamento), "Sp_GnSelEmpleadosPorDepto")
    pLlenarCboRs cboEmpleadoRegistro, rs, 0, 1, 3
    cboEmpleadoRegistro.ListIndex = 0
    
    SSTabCorte.Tab = 0
    
    mskInicioBusqueda.Mask = ""
    mskInicioBusqueda.Text = ldtmFecha
    mskInicioBusqueda.Mask = "##/##/####"
    
    mskFinBusqueda.Mask = ""
    mskFinBusqueda.Text = ldtmFecha
    mskFinBusqueda.Mask = "##/##/####"
    
    DtpFechaInicio.Value = Date
    dtpFechaFin.Value = Date
    
    chkIntermedios.Value = 0
    chkIntermedios_Click
    
    lblnValidarFacturaSinXml = False
    If fblnLicenciaContaElectronica Then
        Set rsTemp = frsSelParametros("PV", vgintClaveEmpresaContable, "BITCORTECAJACHICASINXML")
        If Not rsTemp.EOF Then
            lblnValidarFacturaSinXml = IIf(rsTemp!valor = "0", True, False)
        End If
        rsTemp.Close
    End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
End Sub

Private Sub grdCortes_DblClick()
On Error GoTo NotificaError
    
    If Val(grdCortes.TextMatrix(1, cintColNumero)) <> 0 Then
        pMuestraCorte Val(grdCortes.TextMatrix(grdCortes.Row, cintColNumero)), 0, Format(ldtmFecha, "dd/mm/yyyy"), Format(ldtmFecha, "dd/mm/yyyy"), vgintNumeroDepartamento, IIf(lintTipoCorte = 1, -1, vglngNumeroEmpleado)
        pHabilita 1, 1, 1, 1, 1, IIf(lstrEstado = "A", 1, 0), IIf(lstrEstado = "C" Or chkIntermedios.Value = 1, 1, 0), IIf(lstrEstado = "C" Or chkIntermedios.Value = 1, 1, 0), IIf(lstrEstado = "C" Or chkIntermedios.Value = 1, 1, 0), IIf(lstrEstado = "C", 1, 0)
        pHabilitaFueraCorte
        SSTabCorte.Tab = 0
        cmdLocate.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdCortes_DblClick"))
End Sub

Private Sub grdCortes_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        grdCortes_DblClick
    End If
End Sub



Private Sub mskFinBusqueda_GotFocus()
On Error GoTo NotificaError
    
    pSelMkTexto mskFinBusqueda

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFinBusqueda_GotFocus"))
End Sub

Private Sub mskFinBusqueda_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        cmdCargar.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFinBusqueda_KeyPress"))
End Sub

Private Sub mskFinBusqueda_LostFocus()
On Error GoTo NotificaError
    
    If Trim(mskFinBusqueda.ClipText) = "" Or Not IsDate(mskFinBusqueda.Text) Then
        mskFinBusqueda.Mask = ""
        mskFinBusqueda.Text = ldtmFecha
        mskFinBusqueda.Mask = "##/##/####"
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFinBusqueda_LostFocus"))
End Sub

Private Sub mskHoraFin_GotFocus()
On Error GoTo NotificaError

    tmrHora.Enabled = False
    pSelMkTexto mskHoraFin
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskHoraFin_GotFocus"))
End Sub

Private Sub mskHoraFin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdCorteporMovimiento.SetFocus
    End If
End Sub

Private Sub mskHoraFin_LostFocus()
On Error GoTo NotificaError

    If Trim(mskHoraFin.ClipText) = "" Then
        mskHoraFin.Mask = ""
        mskHoraFin.Text = fdtmServerHora
        mskHoraFin.Mask = "##:##"
    End If
    
   If Not IsDate(mskHoraFin.Text) Then MsgBox SIHOMsg(41), vbOKOnly + vbCritical, "Mensaje"
      
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskHoraFin_LostFocus"))
End Sub

Private Sub mskHoraIni_GotFocus()
On Error GoTo NotificaError
    
    pSelMkTexto mskHoraIni

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskHoraIni_GotFocus"))
End Sub

Private Sub mskHoraIni_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        mskHoraFin.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskHoraIni_KeyPress"))
End Sub

Private Sub mskHoraIni_LostFocus()
On Error GoTo NotificaError
    
    If Trim(mskHoraIni.ClipText) = "" Then
        mskHoraIni.Mask = ""
        mskHoraIni.Text = fdtmServerHora
        mskHoraIni.Mask = "##:##"
    End If
    
    If Not IsDate(mskHoraIni.Text) Then MsgBox SIHOMsg(41), vbOKOnly + vbCritical, "Mensaje"

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskHoraIni_LostFocus"))
End Sub

Private Sub mskInicioBusqueda_GotFocus()
On Error GoTo NotificaError
    
    pSelMkTexto mskInicioBusqueda

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskInicioBusqueda_GotFocus"))
End Sub

Private Sub mskInicioBusqueda_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        mskFinBusqueda.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskInicioBusqueda_KeyPress"))
End Sub

Private Sub mskInicioBusqueda_LostFocus()
On Error GoTo NotificaError
    
    If Trim(mskInicioBusqueda.ClipText) = "" Or Not IsDate(mskInicioBusqueda.Text) Then
        mskInicioBusqueda.Mask = ""
        mskInicioBusqueda.Text = ldtmFecha
        mskInicioBusqueda.Mask = "##/##/####"
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskInicioBusqueda_LostFocus"))
End Sub

Private Sub tmrHora_Timer()
On Error GoTo NotificaError

    mskHoraFin.Text = Format(fdtmServerHora, "hh:mm")
        
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":tmrHora_Timer"))
End Sub

Private Sub txtNumCorte_GotFocus()
On Error GoTo NotificaError
    
    Dim lngNumCorteIngresos As Long
    Dim lngNumCorteCajaChica As Long
    
    lngNumCorteIngresos = flngNumeroCorte(vgintNumeroDepartamento, vglngNumeroEmpleado, "P")
    lngNumCorteCajaChica = flngNumeroCorte(vgintNumeroDepartamento, vglngNumeroEmpleado, "C")
        
    If lngNumCorteIngresos = 0 And lngNumCorteCajaChica = 0 Then
        'No se encontró un corte abierto.
        'MsgBox SIHOMsg(659), vbOKOnly + vbExclamation, "Mensaje"
        'Unload Me
        frsEjecuta_SP CStr(vgintNumeroDepartamento) & "|" & CStr(vglngNumeroEmpleado), "SP_GNINSCORTE"
        chkAcumulado.SetFocus
        txtNumCorte.SetFocus
    Else
        pMuestraCorte IIf(lngNumCorteIngresos > lngNumCorteCajaChica, lngNumCorteIngresos, lngNumCorteCajaChica), 0, Format(ldtmFecha, "dd/mm/yyyy"), Format(ldtmFecha, "dd/mm/yyyy"), -1, -1
        pHabilita 0, 0, 1, 0, 0, 1, 0, 0, 0, 0
        pHabilitaFueraCorte
        pSelTextBox txtNumCorte
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtNumCorte_GotFocus"))
End Sub

Private Sub txtNumCorte_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        pMuestraCorte Val(txtNumCorte.Text), 0, Format(ldtmFecha, "dd/mm/yyyy"), Format(ldtmFecha, "dd/mm/yyyy"), vgintNumeroDepartamento, IIf(lintTipoCorte = 1, -1, vglngNumeroEmpleado)
        
        If lblnCorte Then
            pHabilita 0, 0, 1, 0, 0, IIf(lstrEstado = "A", 1, 0), IIf(lstrEstado = "C" Or chkIntermedios.Value = 1, 1, 0), IIf(lstrEstado = "C" Or chkIntermedios.Value = 1, 1, 0), IIf(lstrEstado = "C" Or chkIntermedios.Value = 1, 1, 0), IIf(lstrEstado = "C", 1, 0)
            pHabilitaFueraCorte
            pEnfocaTextBox txtNumCorte
        Else
            '¡La información no existe!
            MsgBox SIHOMsg(12), vbExclamation + vbOKOnly, "Mensaje"
            txtNumCorte_GotFocus
            pEnfocaTextBox txtNumCorte
        End If
    Else
        If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
            KeyAscii = 7
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtNumCorte_KeyPress"))
End Sub

Private Sub pCierraCorte()
'Cierra corte sin pasar por el proceso de entrega recepción
                    
    Dim rsCnPoliza As New ADODB.Recordset
    Dim rsCnDetallePoliza As New ADODB.Recordset
    Dim rsPvCortePoliza As New ADODB.Recordset      'Recordset de lectura (movimientos contables del corte)
    
    Dim vlstrSentencia As String
    Dim vllngNumeroPoliza As Long           'Consecutivo de la póliza del corte
    Dim vldtmFecha As Date
    Dim dblTotalCargos As Double            'Total de cargos de la póliza
    Dim dblTotalAbonos As Double            'Total de abonos de la póliza
    Dim dbldiferencia As Double             'Diferencia entre cargos y abonos
    Dim intCargoAbono As Integer            'Indica si el movimiento es cargo, 1, o abono 0
    Dim dblCantidadMovto As Double          'Cantidad del movimiento en la póliza
    Dim lngCuentaAfectada As Long           'Cuenta afectada con el cuadre por centavos
    Dim vllngPersonaGraba As Long
    Dim blnCorteAbierto As Boolean          'Para verificar el estado del corte
    Dim vlblnGraboDetallePoliza As Boolean  'Bandera para validar si se guardó algún registro en PvDetallePoliza
    Dim vllngClavePoliza As Long            'Consecutivo anual de la póliza
    Dim vllngContador As Long               'Contador
    Dim vllngCorteGrabando As Long          'Variable para almacenar si se puede guardar el corte
    Dim vllngEstatusPoliza As Long          'Estatus poliza
    Dim lngNuevoCorte As Long               'Num. de corte siguiente
    Dim dblTipoCambio As Double             'Tipo de cambio
    Dim vlPvCortePoliza As Long             'Variable temporal para guardar en el log de transacciones el detalle del grabado del corte las filas afectadas en PvCortePoliza
    Dim vlCnDetallePoliza As Long           'Variable temporal para guardar en el log de transacciones el detalle del grabado del corte las filas afectadas en CnDetallePoliza
    
    If lblnValidarFacturaSinXml And Trim(lstrTipoCorte) = "C" Then
        If fblnSinXml(Trim(txtNumCorte.Text)) Then
            MsgBox SIHOMsg(1482), vbOKOnly + vbInformation, "Mensaje"
            Exit Sub
        End If
    End If
    vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
    If vllngPersonaGraba <> 0 Then
        '---------------------------------------------------------------------------------------------------------------------------'
        ' En caso de que el corte se maneje por empleado valida que sea el mismo de persona graba si no pone 0 en vllngPersonaGraba '
        '---------------------------------------------------------------------------------------------------------------------------'
        If lintTipoCorte = 2 And vllngPersonaGraba <> llngCveEmpleado Then
            'La persona que inició el proceso no corresponde a la persona que quiere grabar.
            MsgBox SIHOMsg(593), vbInformation + vbOKOnly, "Mensaje"
        Else
            EntornoSIHO.ConeccionSIHO.BeginTrans
            
            blnCorteAbierto = fblnCorteAbierto()
            If blnCorteAbierto Then
                '------------------------------------------------------------'
                ' Recordset tipo tabla para registro de la póliza (CnPoliza) '
                '------------------------------------------------------------'
                vlstrSentencia = "SELECT * FROM CnPoliza WHERE intNumeroPoliza = -1"
                Set rsCnPoliza = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
                
                '-------------------------------------------------------------------------------'
                ' Recordset tipo tabla para registro del detalle de la póliza (CnDetallePoliza) '
                '-------------------------------------------------------------------------------'
                vlstrSentencia = "SELECT * FROM CnDetallePoliza WHERE intNumeroPoliza = -1"
                Set rsCnDetallePoliza = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
                
                vlblnGraboDetallePoliza = True 'Esta variable se actualiza en el procedimiento pIncluyeMovimiento. Nota: el cambio
                'solicitado permite seguir con el corte aun no existan movimientos, por eso se cambio el valor del parametro a TRUE
            
                Me.MousePointer = 11
            
                vllngCorteGrabando = 1
                vgstrParametrosSP = Trim(txtNumCorte.Text) & "|" & "Grabando"
                frsEjecuta_SP vgstrParametrosSP, "Sp_PvUpdEstatusCorte", True, vllngCorteGrabando
                            
                If vllngCorteGrabando = 2 Then
                    vllngEstatusPoliza = 1
                    vgstrParametrosSP = vgintClaveEmpresaContable & "|" & "Grabando poliza"
                    frsEjecuta_SP vgstrParametrosSP, "SP_CNUPDESTATUSCIERRE", True, vllngEstatusPoliza
                    
                    If vllngEstatusPoliza = 1 Then
                        vldtmFecha = CDate(Format(lblInicio.Caption, "dd/mmm/yyyy"))
                        
                        If Not fblnPeriodoCerrado(vgintClaveEmpresaContable, Year(vldtmFecha), Month(vldtmFecha)) Then
                            ReDim apoliza(0)
                            apoliza(0).vllngNumeroCuenta = 0
                    
                            '------------------------------------------------------------------'
                            ' Dejar en un arreglo los movimientos registrados en PvCortePoliza '
                            ' que se maten movimientos que no deben reflejarse en la póliza    '
                            ' al sumarse positivos con negativos                               '
                            '------------------------------------------------------------------'
                            vlPvCortePoliza = 0     'Inicialización de variable
                            vlCnDetallePoliza = 0   'Inicialización de variable
                            
                            vlstrSentencia = "SELECT * FROM PvCortePoliza WHERE intNumCorte = " & Trim(txtNumCorte.Text) & " ORDER BY intconsecutivo"
                            Set rsPvCortePoliza = frsEjecuta_SP(Trim(txtNumCorte.Text), "sp_PvSelCortePoliza")
                            vlPvCortePoliza = rsPvCortePoliza.RecordCount
                            If rsPvCortePoliza.RecordCount <> 0 Then
                                With rsPvCortePoliza
                                    .MoveFirst
                                    Do While Not .EOF
                                        pIncluyeMovimiento !intNumCuenta, !MNYCantidad, !bitcargo, IIf(IsNull(!chrFolioDocumento), "", !chrFolioDocumento), IIf(IsNull(!descSalidaCajaChica), "", !descSalidaCajaChica)
                                        .MoveNext
                                    Loop
                                End With
                            End If
                            rsPvCortePoliza.Close
                            lstrConceptoPoliza = ""
                            If fblnObtenerConfigPolizanva(0, IIf(Trim(lstrTipoCorte) = "P", "G", "K"), frmCorte.Name, lblTipoCorte) = True Then
                                For intIndex = 0 To UBound(ConceptosPoliza, 1)
                                    Select Case ConceptosPoliza(intIndex).tipo
                                        Case "1"
                                            lstrConceptoPoliza = lstrConceptoPoliza & ConceptosPoliza(intIndex).Texto
                                        Case "16"
                                            lstrConceptoPoliza = lstrConceptoPoliza & lblDepartamento.Caption
                                        Case "17"
                                            lstrConceptoPoliza = lstrConceptoPoliza & Trim(txtNumCorte.Text)
                                    End Select
                                Next intIndex
                                
                                '------------------------------------------------'
                                ' Vaciar el arreglo a las tablas de contabilidad '
                                '------------------------------------------------'
                                vllngClavePoliza = flngFolioPoliza(vgintClaveEmpresaContable, IIf(Trim(lstrTipoCorte) = "P", "I", "D"), Year(vldtmFecha), Month(vldtmFecha), False)
                            
                                With rsCnPoliza
                                    .AddNew
                                    !tnyclaveempresa = vgintClaveEmpresaContable
                                    !smiEjercicio = Year(vldtmFecha)
                                    !tnyMes = Month(vldtmFecha)
                                    !intClavePoliza = vllngClavePoliza
                                    !dtmFechaPoliza = vldtmFecha
                                    !chrTipoPoliza = IIf(Trim(lstrTipoCorte) = "P", "I", "D")
                                    !vchConceptoPoliza = lstrConceptoPoliza
                                    !smicvedepartamento = vgintNumeroDepartamento
                                    !INTCVEEMPLEADO = vllngPersonaGraba
                                    !bitAsentada = 0
                                    .Update
                                End With
                                vllngNumeroPoliza = flngObtieneIdentity("SEC_CNPOLIZA", rsCnPoliza!intNumeroPoliza)
                                rsCnPoliza.Close
                                
                                With rsCnDetallePoliza
                                    dblTotalCargos = 0
                                    dblTotalAbonos = 0
                                    dbldiferencia = 0
                                    vllngContador = 0
                                    Do While vllngContador <= UBound(apoliza(), 1)
                                        If apoliza(vllngContador).vldblCantidadMovimiento <> 0 Then
                                            vlblnGraboDetallePoliza = True
                                            '- Modificado para el caso 6723. Se cambió la función a Format porque se presentaron diferencias de centavos -'
                                            dblCantidadMovto = Format(apoliza(vllngContador).vldblCantidadMovimiento, "Fixed")
                                            
                                            If apoliza(vllngContador).vlintTipoMovimiento = 1 Then
                                                dblTotalCargos = dblTotalCargos + dblCantidadMovto
                                            Else
                                                dblTotalAbonos = dblTotalAbonos + dblCantidadMovto
                                            End If
                                            
                                            If vllngContador = UBound(apoliza(), 1) Then
                                                'Si es el último movimiento de la póliza
                                                dbldiferencia = dblTotalCargos - dblTotalAbonos
                                                lngCuentaAfectada = apoliza(vllngContador).vllngNumeroCuenta
                                                
                                                If Abs(dbldiferencia) <> 0 And Abs(dbldiferencia) < 1 Then
                                                    'Si existen diferencias entre cargos y abonos
                                                    intCargoAbono = apoliza(vllngContador).vlintTipoMovimiento
                                                    
                                                    If dbldiferencia > 0 And intCargoAbono = 1 Then
                                                        'Hay mas cargos y se hará un cargo, se resta
                                                        dblCantidadMovto = dblCantidadMovto - dbldiferencia
                                                    ElseIf dbldiferencia < 0 And intCargoAbono = 1 Then
                                                        'Hay mas abonos y se hará un cargo, se suma
                                                        dblCantidadMovto = dblCantidadMovto + Abs(dbldiferencia)
                                                    ElseIf dbldiferencia > 0 And intCargoAbono = 0 Then
                                                        'Hay mas cargos y se hará un abono, se suma
                                                        dblCantidadMovto = dblCantidadMovto + dbldiferencia
                                                    ElseIf dbldiferencia < 0 And intCargoAbono = 0 Then
                                                        'Hay mas abonos y se hará un abono, se resta
                                                        dblCantidadMovto = dblCantidadMovto - Abs(dbldiferencia)
                                                    End If
                                                End If
                                            End If
                                            dblCantidadMovto = Format(dblCantidadMovto, "Fixed")
                                            .AddNew
                                            !intNumeroPoliza = vllngNumeroPoliza
                                            !intNumeroCuenta = apoliza(vllngContador).vllngNumeroCuenta
                                            !bitNaturalezaMovimiento = apoliza(vllngContador).vlintTipoMovimiento
                                            !mnyCantidadMovimiento = dblCantidadMovto
                                                                                        
                                            Dim vl_indexrefConDet As Integer
                                            Dim vl_CD, vl_RD As String
                                            For vl_indexrefConDet = 1 To 2
                                                lstrConceptoDetallePoliza = " "
                                                If fblnObtenerConfigPolizanva(vl_indexrefConDet, IIf(Trim(lstrTipoCorte) = "P", "G", "K"), frmCorte.Name) = True Then      'Referencia de detalle
                                                    For intIndex = 0 To UBound(ConceptosPoliza, 1)
                                                        Select Case ConceptosPoliza(intIndex).tipo
                                                            Case "1"
                                                                lstrConceptoDetallePoliza = lstrConceptoDetallePoliza & ConceptosPoliza(intIndex).Texto
                                                            Case "16"
                                                                lstrConceptoDetallePoliza = lstrConceptoDetallePoliza & lblDepartamento.Caption
                                                            Case "17"
                                                                lstrConceptoDetallePoliza = lstrConceptoDetallePoliza & Trim(txtNumCorte.Text)
                                                            Case "3"
                                                                If chkDesglosaCorte.Value = 1 Then
                                                                    lstrConceptoDetallePoliza = lstrConceptoDetallePoliza & apoliza(vllngContador).lstrFolioDocumento
                                                                End If
                                                            Case "21"
                                                                If chkDesglosaCorte.Value = 1 Then
                                                                    lstrConceptoDetallePoliza = lstrConceptoDetallePoliza & apoliza(vllngContador).lstrDescSalidaCajaChica
                                                                End If
                                                        End Select
                                                    Next intIndex
                                                    If vl_indexrefConDet = 1 Then vl_CD = lstrConceptoDetallePoliza
                                                    If vl_indexrefConDet = 2 Then vl_RD = lstrConceptoDetallePoliza
                                                End If
                                            Next vl_indexrefConDet
                                            !vchConcepto = vl_CD
                                            !vchReferencia = Left(vl_RD, 50)
                                            .Update
                                            vlCnDetallePoliza = vlCnDetallePoliza + 1
                                        End If
                                        vllngContador = vllngContador + 1
                                    Loop
                                    
                                    .Close
                                End With
                            
                                '--------------------------------------------------------------'
                                ' Cerrar el corte, ponerle el número de póliza generada, fecha '
                                ' y hora de cierre y persona que cerró                         '
                                '--------------------------------------------------------------'
                                If vlblnGraboDetallePoliza Then
                                    '-----------------------------------'
                                    ' Diferencia entre cargos y abonos: '
                                    '-----------------------------------'
                                    If Abs(dbldiferencia) <> 0 Then
                                        vgstrParametrosSP = Trim(txtNumCorte.Text) & "|" & CStr(lngCuentaAfectada) & "|" & CStr(Abs(dbldiferencia))
                                        frsEjecuta_SP vgstrParametrosSP, "Sp_PvInsCorteAjuste", True
                                    End If
                                    
                                    '------------------------------------'
                                    ' Cerrar el corte y abrir uno nuevo: '
                                    '------------------------------------'
                                    lngNuevoCorte = 1
                                    vgstrParametrosSP = txtNumCorte.Text & "|" & CStr(lintCveDepto) & "|" & CStr(vllngPersonaGraba) & "|" & Trim(lstrTipoCorte) & "|" & CStr(vllngNumeroPoliza) & "|" & chkDesglosaCorte.Value
                                    frsEjecuta_SP vgstrParametrosSP, "SP_PVUPDCIERRACORTE", True, lngNuevoCorte
                                    
                                    '-------------------'
                                    ' Liberar el corte: '
                                    '-------------------'
                                    pLiberaCorte CLng(txtNumCorte.Text)
                                    
                                    If Trim(lstrTipoCorte) = "P" Then
                                        '--------------------------------------------------------'
                                        ' Transferir el dinero de fondo fijo al siguiente corte: '
                                        '--------------------------------------------------------'
                                        vgstrParametrosSP = txtNumCorte.Text & "|" & CStr(lngNuevoCorte)
                                        frsEjecuta_SP vgstrParametrosSP, "Sp_PvUpdTransfiereFondo"
                                    Else
                                        '----------------------------------------------------------'
                                        ' Transferir el saldo de la caja chica al siguiente corte: '
                                        '----------------------------------------------------------'
                                        dblTipoCambio = fdblTipoCambio(fdtmServerFecha, "O")
                                        vgstrParametrosSP = txtNumCorte.Text & "|" & CStr(lngNuevoCorte) & "|" & CStr(dblTipoCambio)
                                        frsEjecuta_SP vgstrParametrosSP, "SP_PVINSSALDOCAJACHICA"
                                    End If
                                    
                                    '*************************************************************************'
                                    '- Registrar los movimientos de las formas de pago relacionados al corte -'
                                    pRegistrarMovimientosBancos CLng(txtNumCorte.Text), 0, Format(lblInicio.Caption, "dd/mmm/yyyy"), Format(fdtmServerFecha, "dd/mmm/yyyy")
                                    '*************************************************************************'
                                End If
                            Else
                                Me.MousePointer = 0
                                EntornoSIHO.ConeccionSIHO.RollbackTrans
                                Exit Sub
                            End If
                            
                        Else
                            frmMensajePeriodoContableCerrado.Show vbModal
                            frmMensajePeriodoContableCerrado.Show vbModal
                            pLiberaCierre
                            Me.MousePointer = 0
                            EntornoSIHO.ConeccionSIHO.RollbackTrans
                            Exit Sub
                        End If
                        pLiberaCierre
                    End If
                    
                    If Not vlblnGraboDetallePoliza Then
                        pLiberaCorte (CLng(Trim(txtNumCorte.Text)))
                    End If
                End If
            End If
        
            If vlblnGraboDetallePoliza And blnCorteAbierto Then
                EntornoSIHO.ConeccionSIHO.CommitTrans
                
                'Candadito para detectar movimientos del corte
                Call pGuardarLogTransaccion(Me.Name, EnmGrabar, vllngPersonaGraba, "GRABADO DE CORTE DE CAJA", " Corte " & Trim(txtNumCorte.Text) & "; PvCortePoliza " & vlPvCortePoliza & "; CnDetallePoliza " & vlCnDetallePoliza & "; Contador " & vllngContador)
                vlCortePoliza = 0 'Reinicialización de la variable
                vlCnDetallePoliza = 0 'Reinicialización de la variable
                
                'La operación se realizó satisfactoriamente.
                MsgBox SIHOMsg(420), vbOKOnly + vbInformation, "Mensaje"
                
                pMuestraCorte Val(txtNumCorte.Text), 0, Format(ldtmFecha, "dd/mm/yyyy"), Format(ldtmFecha, "dd/mm/yyyy"), vgintNumeroDepartamento, IIf(lintTipoCorte = 1, -1, vglngNumeroEmpleado)
                pHabilita 0, 0, 1, 0, 0, IIf(lstrEstado = "A", 1, 0), IIf(lstrEstado = "C" Or chkIntermedios.Value = 1, 1, 0), IIf(lstrEstado = "C" Or chkIntermedios.Value = 1, 1, 0), IIf(lstrEstado = "C" Or chkIntermedios.Value = 1, 1, 0), IIf(lstrEstado = "C", 1, 0)
                pHabilitaFueraCorte
                
                If cmdPrintPoliza.Enabled Then
                    cmdPrintPoliza.SetFocus
                End If
            Else
                If Not blnCorteAbierto Then
                    EntornoSIHO.ConeccionSIHO.RollbackTrans
                    
                    'La información ha cambiado, consulte de nuevo.
                    MsgBox SIHOMsg(381), vbExclamation + vbOKOnly, "Mensaje"
                    txtNumCorte.SetFocus
                Else
                    EntornoSIHO.ConeccionSIHO.RollbackTrans
                    
                    'Aún no se registran movimientos.
                    '¿Desea actualizar la fecha del corte?
                    If MsgBox(SIHOMsg(311), vbQuestion + vbYesNo, "Mensaje") = vbYes Then
                        frsEjecuta_SP txtNumCorte.Text, "SP_PVUPDFECHACORTE"
                    End If
                    txtNumCorte.SetFocus
                End If
            End If
        
            Me.MousePointer = 0
        End If
    End If
End Sub
Private Function fblnSinXml(strCorte As String) As Boolean
On Error GoTo NotificaError
    Dim rs As New ADODB.Recordset

    fblnSinXml = False
    Set rs = frsRegresaRs("SELECT * FROM cpFacturaCajaChica WHERE intNumCorte = " & strCorte & _
                          " and chrTipoDocumento in ('F','H','L') and chrEstado not in ('C','R') and (Trim(chrTipoComprobante) = '' Or chrTipoComprobante = '0' Or chrTipoComprobante is null) ")
    If rs.RecordCount <> 0 Then
        fblnSinXml = True
    End If
    
Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnSinXml"))
End Function

Private Sub pMuestraCorte(lngNumCorte As Long, intFiltroFecha As Integer, strFechaIni As String, strFechaFin As String, intCveDepto As Integer, lngCveEmpleado As Long)
On Error GoTo NotificaError
    
    Dim rsTotales As New ADODB.Recordset
    
    lblnCorte = False
    vgstrParametrosSP = CStr(lngNumCorte) & "|" & CStr(intFiltroFecha) & "|" & fstrFechaSQL(strFechaIni) & "|" & fstrFechaSQL(strFechaFin) & "|" & CStr(intCveDepto) & "|" & CStr(lngCveEmpleado)
    Set rs = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELCORTE")
    If rs.RecordCount <> 0 Then
        lblnCorte = True
        vllngNumPoliza = IIf(IsNull(rs!IdPoliza), 0, rs!IdPoliza)
        
        txtNumCorte.Text = rs!IdCorte
        lintCveDepto = rs!IdDepartamento
        lblDepartamento.Caption = rs!departamento
        
        lblInicio.Caption = rs!FechaAbre
        lblFin.Caption = IIf(IsNull(rs!FechaCierra), "", rs!FechaCierra)
        
        lstrFechaInicio = Format(rs!FechaAbre, "dd/mmm/yyyy hh:mm")
        lstrFechaFin = ""
        If Not IsNull(rs!FechaCierra) Then
            lstrFechaFin = Format(rs!FechaCierra, "dd/mmm/yyyy hh:mm")
        End If
                
        llngCveEmpleado = rs!IdEmpleadoAbre
        lblPersonaAbre.Caption = rs!PersonaAbre
        lblPersonaCierra.Caption = rs!PersonaCierra
        
        lstrTipoCorte = rs!TipoCorte
        lblTipoCorte.Caption = rs!TipoCorteDescripcion
    
        lblEstatus.Caption = rs!Estado
        lstrEstado = rs!Estatus
        
        lblEjercicio.Caption = IIf(IsNull(rs!FechaCierra), "", rs!Ejercicio)
        lblMes.Caption = IIf(IsNull(rs!FechaCierra), "", rs!mes)
        lblFolio.Caption = IIf(IsNull(rs!FechaCierra), "", rs!folio)
        
        If Trim(lstrEstado) = "A" Then
            chkDesglosaCorte.Value = IIf(vlblnDesglosaPolizaConfig, 1, 0)
        Else
            chkDesglosaCorte.Value = IIf(IsNull(rs!DesglosaCorte), 0, rs!DesglosaCorte)
        End If
    
        vgstrParametrosSP = CStr(lngNumCorte) & "|" & CStr(chkAcumulado.Value)
        Set rsDetalleCorte = frsEjecuta_SP(vgstrParametrosSP, "Sp_PvSelMovtosCorte")
        lstFormas.Clear
        If rsDetalleCorte.RecordCount <> 0 Then
            rsDetalleCorte.MoveFirst
            Do While Not rsDetalleCorte.EOF
                lstFormas.AddItem rsDetalleCorte.Fields(0)
                rsDetalleCorte.MoveNext
            Loop
        End If
    
        lblTotalPesos.Caption = FormatCurrency(0, 2)
        lblTotalDolares.Caption = FormatCurrency(0, 2)
    
        Set rsTotales = frsEjecuta_SP(CStr(lngNumCorte), "SP_PVSELCORTETOTALES")
        If rsTotales.RecordCount <> 0 Then
            lblTotalPesos.Caption = FormatCurrency(rsTotales!TotalPesos, 2)
            lblTotalDolares.Caption = FormatCurrency(rsTotales!TotalDolares, 2)
        End If
        
        chkFacturas.Enabled = lstrTipoCorte = "P"
        chkPagos.Enabled = lstrTipoCorte = "P"
        chkTickets.Enabled = lstrTipoCorte = "P"
        chkSalidas.Enabled = lstrTipoCorte = "P"
        chkPagosCredito.Enabled = lstrTipoCorte = "P"
        chkHonorarios.Enabled = lstrTipoCorte = "P"
        chkTransferencias.Enabled = lstrTipoCorte = "P"
        chkFondoFijo.Enabled = lstrTipoCorte = "P"
        chkEntradaChica.Enabled = lstrTipoCorte = "C"
        chkSalidaChica.Enabled = lstrTipoCorte = "C"
        
        chkFacturas.Value = IIf(lstrTipoCorte = "P", 1, 0)
        chkPagos.Value = IIf(lstrTipoCorte = "P", 1, 0)
        chkTickets.Value = IIf(lstrTipoCorte = "P", 1, 0)
        chkSalidas.Value = IIf(lstrTipoCorte = "P", 1, 0)
        chkPagosCredito.Value = IIf(lstrTipoCorte = "P", 1, 0)
        chkHonorarios.Value = IIf(lstrTipoCorte = "P", 1, 0)
        chkTransferencias.Value = IIf(lstrTipoCorte = "P", 1, 0)
        chkFondoFijo.Value = IIf(lstrTipoCorte = "P", 1, 0)
        chkEntradaChica.Value = IIf(lstrTipoCorte = "C", 1, 0)
        chkSalidaChica.Value = IIf(lstrTipoCorte = "C", 1, 0)
    Else
        '¡No se encontró ese número de corte.
        MsgBox SIHOMsg(306), vbOKOnly + vbInformation, "Mensaje"
        txtNumCorte_GotFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pMuestraCorte"))
End Sub

Private Sub cboEmpleadoRegistro_Click()
    chkAgrupadoPorEmpleado.Enabled = IIf(cboEmpleadoRegistro.ListIndex = 0, True, False)
    If Not chkAgrupadoPorEmpleado.Enabled Then chkAgrupadoPorEmpleado.Value = False
End Sub

Private Function fblnCorteAbierto() As Boolean
    fblnCorteAbierto = False
    
    vgstrParametrosSP = Trim(txtNumCorte.Text) & "|" & "0" & "|" & fstrFechaSQL(Format(ldtmFecha, "dd/mm/yyyy")) & "|" & fstrFechaSQL(Format(ldtmFecha, "dd/mm/yyyy")) & "|" & "-1" & "|" & "-1"
    Set rs = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELCORTE")
    If rs.RecordCount <> 0 Then
        fblnCorteAbierto = IsNull(rs!FechaCierra)
    End If
End Function

'----- (CR) Agregado caso 7442: Registrar los movimientos realizados por las formas de pago al libro de bancos -----'
'Los movimientos provienen de facturación, pagos, entradas y salidas de dinero en donde se selecciona una forma de pago'
Private Sub pRegistrarMovimientosBancos(lngNumCorte As Long, intFiltroFecha As Integer, strFechaIni As String, strFechaFin As String)
On Error GoTo NotificaError
    
    Dim rsMovimientos As New ADODB.Recordset
    Dim llngIdKardex As Long
    Dim ldblCargo As Double, ldblAbono As Double
    
    vgstrParametrosSP = CStr(lngNumCorte) & "|" & CStr(intFiltroFecha) & "|" & fstrFechaSQL(strFechaIni) & "|" & fstrFechaSQL(strFechaFin) & "|" & lintCveDepto & "|" & "-1" & "|" & "1"
    Set rsMovimientos = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELMOVIMIENTOBANCOFORMA")
    If rsMovimientos.RecordCount <> 0 Then
        rsMovimientos.MoveFirst
        Do While Not rsMovimientos.EOF
            '- Revisar si la cantidad es positiva, si lo es, será cargo, si no, será abono -'
            If rsMovimientos!cantidad > 0 Then
                ldblCargo = rsMovimientos!cantidad
                ldblAbono = 0
            Else
                ldblAbono = rsMovimientos!cantidad * (-1) 'Convertir la cantidad a positiva para guardarla como abono
                ldblCargo = 0
            End If
            
            '- Registra en el libro de bancos el movimiento de la forma de pago ligada a una cuenta de banco -'
            vgstrParametrosSP = CStr(rsMovimientos!IdBanco) & "|" & fstrFechaSQL(Format(rsMovimientos!fecha, "dd/mm/yyyy"), Format(rsMovimientos!fecha, "hh:mm:ss")) & "|" & _
                                rsMovimientos!Movimiento & "|" & ldblCargo & "|" & ldblAbono & "|" & CStr(rsMovimientos!Referencia)
            frsEjecuta_SP vgstrParametrosSP, "Sp_CpInsKardexBanco"
            llngIdKardex = flngObtieneIdentity("Sec_CpKardexBanco", 0) 'Id del movimiento del kárdex
        
            '- Registra en la tabla kárdex del banco del módulo correspondiente -'
            If Trim(rsMovimientos!Origen) = "CC" And (rsMovimientos!Documento <> "FA" And rsMovimientos!Documento <> "HO") Then  'CRÉDITO
                vgstrParametrosSP = Trim(rsMovimientos!Documento) & "|" & CStr(rsMovimientos!Referencia) & "|" & CStr(llngIdKardex)
                frsEjecuta_SP vgstrParametrosSP, "Sp_CcInsPagoKardexBanco"
            Else 'CAJA
                If (Trim(rsMovimientos!Origen) = "CC" And (rsMovimientos!Documento = "FA" Or rsMovimientos!Documento = "HO")) Or Trim(rsMovimientos!Origen) = "PV" Then
                    vgstrParametrosSP = CStr(rsMovimientos!Referencia) & "|" & CStr(llngIdKardex) & "|" & Trim(rsMovimientos!Documento)
                    frsEjecuta_SP vgstrParametrosSP, "Sp_PvInsPagoKardexBanco"
                End If
            End If
            
            '- Actualizar la fecha y el estado a "Registrado" del movimiento en tabla intermedia -'
            vlstrSentencia = "UPDATE PvMovimientoBancoForma SET bitEstado = 2, dtmFechaRegistro = " & fstrFechaSQL(fdtmServerFecha, fdtmServerHora) & _
                             " WHERE intIdMovimiento = " & rsMovimientos!IdMovimiento
            pEjecutaSentencia vlstrSentencia
            
            rsMovimientos.MoveNext
        Loop
    End If
    rsMovimientos.Close

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pRegistrarMovimientosBancos"))
End Sub

'----- (CR) Agregado caso 7442: Revisar si el corte tiene movimientos fuera de corte para habilitar los botones -----'
Private Sub pHabilitaFueraCorte()
On Error GoTo NotificaError
    
    Dim rsMovimientos As New ADODB.Recordset
    Dim lblnHayMovimientos As Boolean
    
    vgstrParametrosSP = Val(txtNumCorte.Text) & "|" & _
                        fstrFechaSQL(lblInicio.Caption, mskHoraIni & ":00", True) & "|" & _
                        fstrFechaSQL(dtpFechaFin.Value, mskHoraFin & ":00", True) & "|" & _
                        vgintNumeroDepartamento & "|" & _
                        1 & "|" & 1 & "|" & 1 & "|" & 1 & "|" & 0 & "|" & 0 & "|" & _
                        IIf(optGrupoFormaPago.Value, 1, 2) & "|" & -1 & "|" & _
                        chkAgrupadoPorEmpleado.Value & "|" & Str(vgintClaveEmpresaContable)
    Set rsMovimientos = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELMOVIMIENTOFUERACORTE")
    lblnHayMovimientos = (rsMovimientos.RecordCount > 0)
    cmdMovsFueraCorte.Enabled = lblnHayMovimientos
    cmdPolizaFueraCorte.Enabled = lblnHayMovimientos
    rsMovimientos.Close

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pHabilitaFueraCorte"))
End Sub
