VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.1#0"; "MyCommandButton.ocx"
Object = "{FF14BD24-9F8A-41E3-B5B8-7F0D45EE9F16}#15.0#0"; "hsflatcontrols.ocx"
Begin VB.Form frmReportesIngreso 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingresos de pacientes"
   ClientHeight    =   9975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9975
   ScaleWidth      =   8655
   StartUpPosition =   2  'CenterScreen
   Begin HSFlatControls.MyCombo cboMedico 
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   5700
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   661
      Style           =   1
      Enabled         =   -1  'True
      Text            =   "MyCombo1"
      Sorted          =   -1  'True
      List            =   ""
      ItemData        =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   8970
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8415
      Begin VB.Frame Frame10 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Ocultar"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   4560
         TabIndex        =   54
         Top             =   7680
         Width           =   3735
         Begin VB.CheckBox chkOcultaEtiquetas 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Etiquetas de tipo de ingreso"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   250
            Left            =   240
            TabIndex        =   40
            Top             =   600
            Width           =   3320
         End
         Begin VB.CheckBox chkOcultarFiltros 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Filtros utilizados"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   250
            Left            =   240
            TabIndex        =   39
            ToolTipText     =   "Ocultar filtros utilizados"
            Top             =   360
            Width           =   2655
         End
      End
      Begin VB.Frame Frame9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Mostrar"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2670
         Left            =   4560
         TabIndex        =   53
         Top             =   4880
         Width           =   3735
         Begin VB.CheckBox chkFechaEgreso 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Fecha de egreso"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   250
            Left            =   240
            TabIndex        =   32
            ToolTipText     =   "Mostrar fecha de egreso"
            Top             =   1080
            Width           =   3320
         End
         Begin VB.CheckBox chkTotalCuenta 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Total de la cuenta"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   250
            Left            =   240
            TabIndex        =   36
            ToolTipText     =   "Mostrar total de la cuenta"
            Top             =   2040
            Width           =   3320
         End
         Begin VB.CheckBox chkNoControl 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Número de control"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   250
            Left            =   240
            TabIndex        =   33
            ToolTipText     =   "Mostrar número de control del paciente en la empresa de convenio"
            Top             =   1320
            Width           =   3320
         End
         Begin VB.CheckBox chkNoPoliza 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Número de póliza"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   250
            Left            =   240
            TabIndex        =   34
            ToolTipText     =   "Mostrar número de póliza"
            Top             =   1560
            Width           =   3320
         End
         Begin VB.CheckBox chkAutoriza 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Persona que autoriza"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   250
            Left            =   240
            TabIndex        =   35
            ToolTipText     =   "Mostrar persona que autoriza por parte de la aseguradora"
            Top             =   1800
            Width           =   3320
         End
         Begin VB.CheckBox chkAdmisionCancelada 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Admisiones canceladas"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   250
            Left            =   240
            TabIndex        =   31
            Top             =   840
            Width           =   3320
         End
         Begin VB.CheckBox ChkOrden 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Orden de internamiento"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   250
            Left            =   240
            TabIndex        =   30
            Top             =   600
            Width           =   3320
         End
         Begin VB.CheckBox chkDiagnostico 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Diagnóstico del paciente"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   250
            Left            =   240
            TabIndex        =   29
            Top             =   360
            Width           =   3320
         End
         Begin VB.CheckBox chkdetalle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Detalle de la cuenta"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   250
            Left            =   240
            TabIndex        =   28
            ToolTipText     =   "Mostrar detalle de la cuenta"
            Top             =   2400
            Width           =   3320
         End
      End
      Begin HSFlatControls.MyCombo cboProcedencia2 
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   7150
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   661
         Style           =   1
         Enabled         =   -1  'True
         Text            =   "MyCombo1"
         Sorted          =   -1  'True
         List            =   ""
         ItemData        =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin HSFlatControls.MyCombo cboArea 
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   6400
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   661
         Style           =   1
         Enabled         =   -1  'True
         Text            =   "MyCombo1"
         Sorted          =   -1  'True
         List            =   ""
         ItemData        =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin HSFlatControls.MyCombo cboProcedencia 
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   5000
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   661
         Style           =   1
         Enabled         =   -1  'True
         Text            =   "MyCombo1"
         Sorted          =   -1  'True
         List            =   ""
         ItemData        =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CheckBox chkHorizontal 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Presentación horizontal"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   250
         Left            =   120
         TabIndex        =   38
         Top             =   8280
         Width           =   2655
      End
      Begin VB.CheckBox chkDetallado 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Detallado"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   250
         Left            =   120
         TabIndex        =   37
         ToolTipText     =   "Detallado"
         Top             =   8040
         Width           =   2655
      End
      Begin VB.Frame Frame7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Ordenar por"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   800
         Left            =   5280
         TabIndex        =   25
         Top             =   3900
         Width           =   3015
         Begin VB.OptionButton optOrdenar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Cuarto"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   250
            Index           =   1
            Left            =   240
            TabIndex        =   27
            Top             =   480
            Width           =   1335
         End
         Begin VB.OptionButton optOrdenar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Fecha de ingreso"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   250
            Index           =   0
            Left            =   240
            TabIndex        =   26
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Agrupar por"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1780
         Left            =   5280
         TabIndex        =   18
         Top             =   2000
         Width           =   3015
         Begin VB.OptionButton optAgrupar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Procedencia"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   250
            Index           =   5
            Left            =   240
            TabIndex        =   24
            Top             =   1460
            Width           =   1575
         End
         Begin VB.OptionButton optAgrupar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Tipo de ingreso"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   250
            Index           =   4
            Left            =   240
            TabIndex        =   23
            Top             =   1200
            Width           =   1815
         End
         Begin VB.OptionButton optAgrupar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Médico tratante"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   250
            Index           =   3
            Left            =   240
            TabIndex        =   22
            Top             =   960
            Width           =   1935
         End
         Begin VB.OptionButton optAgrupar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Tipo de paciente"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   250
            Index           =   2
            Left            =   240
            TabIndex        =   21
            Top             =   720
            Width           =   1935
         End
         Begin VB.OptionButton optAgrupar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Área"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   250
            Index           =   1
            Left            =   240
            TabIndex        =   20
            Top             =   480
            Width           =   855
         End
         Begin VB.OptionButton optAgrupar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Sin agrupar"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   250
            Index           =   0
            Left            =   240
            TabIndex        =   19
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame Frame5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Sexo"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1020
         Left            =   5280
         TabIndex        =   14
         Top             =   860
         Width           =   3015
         Begin VB.OptionButton optSexo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Femenino"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   250
            Index           =   2
            Left            =   240
            TabIndex        =   17
            Top             =   720
            Width           =   1335
         End
         Begin VB.OptionButton optSexo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Masculino"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   250
            Index           =   1
            Left            =   240
            TabIndex        =   16
            Top             =   470
            Width           =   1335
         End
         Begin VB.OptionButton optSexo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Ambos"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   250
            Index           =   0
            Left            =   240
            TabIndex        =   15
            Top             =   230
            Width           =   1215
         End
      End
      Begin VB.OptionButton optFecha 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Actualmente internos"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   250
         Index           =   1
         Left            =   2640
         TabIndex        =   4
         Top             =   850
         Width           =   2415
      End
      Begin VB.OptionButton optFecha 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Por rango de fechas"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   250
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   850
         Width           =   2415
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   850
         Left            =   120
         TabIndex        =   5
         Top             =   860
         Width           =   5055
         Begin MSMask.MaskEdBox mskRango 
            Height          =   375
            Index           =   0
            Left            =   520
            TabIndex        =   6
            Top             =   280
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd/mmm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskRango 
            Height          =   375
            Index           =   1
            Left            =   2520
            TabIndex        =   7
            Top             =   280
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd/mmm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Del"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   250
            Left            =   120
            TabIndex        =   45
            Top             =   340
            Width           =   375
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "al"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   250
            Left            =   2280
            TabIndex        =   44
            Top             =   340
            Width           =   255
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Tipo de ingreso"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2985
         Left            =   120
         TabIndex        =   8
         Top             =   1700
         Width           =   5055
         Begin VB.ListBox lstTipoIngreso 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2595
            ItemData        =   "frmReportesIngreso.frx":0000
            Left            =   120
            List            =   "frmReportesIngreso.frx":0002
            Style           =   1  'Checkbox
            TabIndex        =   9
            Top             =   250
            Width           =   4815
         End
      End
      Begin HSFlatControls.MyCombo cboEmpresa 
         Height          =   375
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Empresa contable"
         Top             =   400
         Width           =   8170
         _ExtentX        =   14420
         _ExtentY        =   661
         Style           =   1
         Enabled         =   -1  'True
         Text            =   "MyCombo1"
         Sorted          =   -1  'True
         List            =   ""
         ItemData        =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Empresa contable"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   120
         TabIndex        =   1
         Top             =   160
         Width           =   3855
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Procedencia"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   51
         Top             =   6900
         Width           =   3855
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Área"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   120
         TabIndex        =   52
         Top             =   6150
         Width           =   3855
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Médico"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   120
         TabIndex        =   47
         Top             =   5450
         Width           =   3855
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tipo de paciente"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   120
         TabIndex        =   46
         Top             =   4750
         Width           =   3855
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   48
      Top             =   10335
      Width           =   975
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "Salir"
         Height          =   255
         Left            =   0
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdSiguiente 
         Caption         =   "Siguiente"
         Default         =   -1  'True
         Height          =   255
         Left            =   0
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.Frame Frame6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   810
      Left            =   3670
      TabIndex        =   43
      Top             =   9030
      Width           =   1320
      Begin MyCommandButton.MyButton cmdPreview 
         Height          =   600
         Left            =   60
         TabIndex        =   41
         ToolTipText     =   "Vista previa"
         Top             =   170
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   1058
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         MaskColor       =   16777215
         Picture         =   "frmReportesIngreso.frx":0004
         BackColorOver   =   -2147483633
         BackColorFocus  =   -2147483633
         BackColorDisabled=   -2147483633
         BorderColor     =   -2147483627
         TransparentColor=   16777215
         Caption         =   ""
         DepthEvent      =   1
         PictureDisabled =   "frmReportesIngreso.frx":0988
         PictureAlignment=   4
         PictureDisabledEffect=   0
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdPrint 
         Height          =   600
         Left            =   660
         TabIndex        =   42
         ToolTipText     =   "Imprimir"
         Top             =   170
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   1058
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         MaskColor       =   16777215
         Picture         =   "frmReportesIngreso.frx":130A
         BackColorOver   =   -2147483633
         BackColorFocus  =   -2147483633
         BackColorDisabled=   -2147483633
         BorderColor     =   -2147483627
         TransparentColor=   16777215
         Caption         =   ""
         DepthEvent      =   1
         PictureDisabled =   "frmReportesIngreso.frx":1C8E
         PictureAlignment=   4
         PictureDisabledEffect=   0
         ShowFocus       =   -1  'True
      End
   End
End
Attribute VB_Name = "frmReportesIngreso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lrptReporteVer As CRAXDRT.Report
Public vgstrmodulo As String

Private Sub cboEmpresa_Click()
    Dim strSql As String
    Dim rs As ADODB.Recordset
    
    cboArea.Clear
    strSql = "select ADArea.tnyCveArea Clave, ADArea.vchDescripcion Descrip " & _
    "from ADArea inner join nodepartamento on adarea.tnycvedepto = nodepartamento.smicvedepartamento " & _
    "where nodepartamento.tnyclaveempresa = " & cboEmpresa.ItemData(cboEmpresa.ListIndex) & "  order by Descrip"
    Set rs = frsRegresaRs(strSql)
    
    Do Until rs.EOF
        cboArea.AddItem rs!Descrip
        cboArea.ItemData(cboArea.NewIndex) = rs!Clave
        rs.MoveNext
    Loop
    cboArea.AddItem "<TODOS>", 0
    cboArea.ItemData(cboArea.NewIndex) = -1
    cboArea.ListIndex = 0
End Sub

Private Sub chkAutoriza_Click()
    If chkAutoriza.Value = vbChecked Then chkDetallado.Value = vbChecked
End Sub

Private Sub chkDetallado_Click()
    If chkDetallado.Value = vbUnchecked Then
        chkDiagnostico.Value = vbUnchecked
        chkDiagnostico.Enabled = False
        chkFechaEgreso.Value = vbUnchecked
        chkFechaEgreso.Enabled = False
        chkNoControl.Value = vbUnchecked
        chkNoControl.Enabled = False
        chkNoPoliza.Value = vbUnchecked
        chkNoPoliza.Enabled = False
        chkAutoriza.Value = vbUnchecked
        chkAutoriza.Enabled = False
        chkTotalCuenta.Value = vbUnchecked
        chkTotalCuenta.Enabled = False
        ChkOrden.Value = vbUnchecked
        ChkOrden.Enabled = False
    Else
        chkDiagnostico.Enabled = True
        chkFechaEgreso.Enabled = True
        chkNoControl.Enabled = True
        chkNoPoliza.Enabled = True
        chkAutoriza.Enabled = True
        chkTotalCuenta.Enabled = True
        ChkOrden.Enabled = True
    End If
End Sub

Private Sub chkDiagnostico_Click()
    If chkDiagnostico.Value = vbChecked Then chkDetallado.Value = vbChecked
End Sub

Private Sub chkFechaEgreso_Click()
    If chkFechaEgreso.Value = vbChecked Then chkDetallado.Value = vbChecked
End Sub

Private Sub chkNoControl_Click()
    If chkNoControl.Value = vbChecked Then chkDetallado.Value = vbChecked
End Sub

Private Sub chkNoPoliza_Click()
    If chkNoPoliza.Value = vbChecked Then chkDetallado.Value = vbChecked
End Sub

Private Sub chkTotalCuenta_Click()
    If chkTotalCuenta.Value = vbChecked Then chkDetallado.Value = vbChecked
End Sub

Private Sub cmdPreview_Click()
    pImprime "P"
End Sub

Private Sub cmdPrint_Click()
    pImprime "I"
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdSiguiente_Click()
    SendKeys vbTab
End Sub

Private Sub Form_Load()
    Dim rsTipo As ADODB.Recordset
    Dim lngContador As Integer
    Dim vlintFechaInicio As String
    mskRango(0) = CDate("01/" & CStr(Month(fdtmServerFecha)) & "/" & CStr(Year(fdtmServerFecha)))
    vlintFechaInicio = Mid(mskRango(0), 3, 8)
    mskRango(0) = "01" & vlintFechaInicio
    mskRango(1) = fdtmServerFecha
    Me.Icon = frmMenuPrincipal.Icon
    pIniciaForma
    
    Set rsTipo = frsEjecuta_SP("", "SP_SISELTIPOINGRESO")
    For lngContador = 1 To rsTipo.RecordCount
        'If rsTipo!intCveTipoIngreso >= 1 And rsTipo!intCveTipoIngreso <> 10 Then  Caso 14420. Se quitó para que muestre «CONSULTA EXTERNA»
        If rsTipo!intCveTipoIngreso >= 1 Then
            lstTipoIngreso.AddItem rsTipo!vchNombre
            lstTipoIngreso.ItemData(lstTipoIngreso.NewIndex) = rsTipo!intCveTipoIngreso
        End If
        rsTipo.MoveNext
    Next lngContador
    rsTipo.Close
    
    lstTipoIngreso.AddItem "PREVIO"
    lstTipoIngreso.ItemData(lstTipoIngreso.NewIndex) = 0
    
   If cgstrModulo = "PV" Or vgstrmodulo = "PV" Then
        chkdetalle.Enabled = True
        chkdetalle.Visible = True
        chkdetalle.Height = 450
        chkdetalle.Top = 220
        chkdetalle.Caption = "Mostrar solo pendientes de facturar"
        chkdetalle.ToolTipText = "Mostrar solo pendientes de facturar"
        chkdetalle.Value = vbChecked
        chkDetallado.Value = vbUnchecked
        chkDetallado.Top = 5750 '5000
        chkDetallado.Left = 4800
        chkDiagnostico.Visible = False
        ChkOrden.Visible = False      'caso 4260
        chkAdmisionCancelada.Visible = False
        '--- INICIO FM
        chkHorizontal.Visible = False
        Label6.Visible = False
        cboProcedencia2.Visible = False
        optAgrupar(5).Visible = False
        chkOcultaEtiquetas.Visible = False
        
        chkFechaEgreso.Visible = False
        chkNoControl.Visible = False
        chkNoPoliza.Visible = False
        chkAutoriza.Visible = False
        chkTotalCuenta.Visible = False
        
        chkOcultarFiltros.Top = 280
        Frame9.Height = 800
        'Frame9.Top = 4850
        Frame10.Height = 700
        Frame10.Top = 6080      '5690
        Frame3.Height = 6900 '6100
        Frame4.Height = 1530
        Frame4.Top = 2000 '1390
        Frame6.Top = 6900 '6000
        frmReportesIngreso.Height = 8220 '7320
        '--- FIN FM
    Else
        chkdetalle.Visible = False
        chkdetalle.Enabled = False
        pInstanciaReporte lrptReporteVer, "rptGeneralIngresosPacientes.rpt"
    End If
End Sub

Private Sub pIniciaForma()
    Dim intIndex As Integer
    Dim rs As ADODB.Recordset
    Dim strSql As String
    Dim cont As Integer
    
    optFecha(0).Value = True
    'mskRango(0) = DateSerial(Year(Date), Month(Date), 1)
    'mskRango(1) = DateSerial(Year(Date), Month(Date), Day(Date))
    
    
    For intIndex = 0 To lstTipoIngreso.ListCount - 1
        lstTipoIngreso.Selected(intIndex) = True
    Next
    
    
    optSexo(0).Value = True
    optAgrupar(0).Value = True
    optOrdenar(0).Value = True
    chkDiagnostico.Value = vbUnchecked
    chkDetallado.Value = vbChecked
    chkOcultarFiltros.Value = vbUnchecked
    chkHorizontal.Value = vbUnchecked
    chkOcultaEtiquetas.Value = vbUnchecked
    lstTipoIngreso_LostFocus
    
    strSql = "select  tnyclaveempresa Clave, vchnombre Descrip from cnempresacontable where bitactiva = 1 order by vchnombre" 'vgintClaveEmpresaContable
    Set rs = frsRegresaRs(strSql)
    Do Until rs.EOF
        cboEmpresa.AddItem rs!Descrip
        cboEmpresa.ItemData(cboEmpresa.NewIndex) = rs!Clave
        rs.MoveNext
    Loop
    rs.Close
    For cont = 0 To cboEmpresa.ListCount
        If cboEmpresa.ItemData(cont) = vgintClaveEmpresaContable Then
            cboEmpresa.ListIndex = cont
            Exit For
        End If
    Next
    If Not fblnRevisaPermiso(vglngNumeroLogin, 2175, "C", True) Then cboEmpresa.Enabled = False
    
    strSql = "select -1 * tnyCveTipoPaciente Clave" & _
    ", vchDescripcion Descrip" & _
    ", 'A'" & _
    " from ADTipoPaciente" & _
    " union" & _
    " select intCveEmpresa" & _
    ", vchDescripcion" & _
    ", 'B'" & _
    " from CCEmpresa" & _
    " order by 3,2"
    Set rs = frsRegresaRs(strSql)
    Do Until rs.EOF
        cboProcedencia.AddItem rs!Descrip
        cboProcedencia.ItemData(cboProcedencia.NewIndex) = rs!Clave
        rs.MoveNext
    Loop
    rs.Close
    cboProcedencia.AddItem "<TODOS>", 0
    cboProcedencia.ItemData(cboProcedencia.NewIndex) = 0
    cboProcedencia.ListIndex = 0
    
    strSql = "select intCveMedico Clave" & _
    ", vchApellidoPaterno || ' ' || vchApellidoMaterno || ' ' || vchNombre Nombre" & _
    " from HOMedico" & _
    " order by Nombre"
    Set rs = frsRegresaRs(strSql)
    Do Until rs.EOF
        cboMedico.AddItem rs!Nombre
        cboMedico.ItemData(cboMedico.NewIndex) = rs!Clave
        rs.MoveNext
    Loop
    rs.Close
    cboMedico.AddItem "<TODOS>", 0
    cboMedico.ItemData(cboMedico.NewIndex) = -1
    cboMedico.ListIndex = 0
    
    'strSQL = "select ADArea.tnyCveArea Clave, ADArea.vchDescripcion Descrip " & _
    '"from ADArea inner join nodepartamento on adarea.tnycvedepto = nodepartamento.smicvedepartamento " & _
    '"where nodepartamento.tnyclaveempresa = " & cboEmpresa.ItemData(cboEmpresa.ListIndex) & "  order by Descrip"
    'Set rs = frsRegresaRs(strSQL)
    'Do Until rs.EOF
    '    cboArea.AddItem rs!Descrip
    '    cboArea.ItemData(cboArea.newIndex) = rs!Clave
    '    rs.MoveNext
    'Loop
    'cboArea.AddItem "<TODOS>", 0
    'cboArea.ItemData(cboArea.newIndex) = -1
    'cboArea.ListIndex = 0
    
    strSql = "select intcveprocedencia Clave" & _
    ", vchDescripcion Descrip" & _
    " from ADProcedencia" & _
    " where bitactivo = 1" & _
    " order by Descrip"
    Set rs = frsRegresaRs(strSql)
    Do Until rs.EOF
        cboProcedencia2.AddItem rs!Descrip
        cboProcedencia2.ItemData(cboProcedencia2.NewIndex) = rs!Clave
        rs.MoveNext
    Loop
    rs.Close
    cboProcedencia2.AddItem "<TODOS>", 0
    cboProcedencia2.ItemData(cboProcedencia.NewIndex) = 0
    cboProcedencia2.ListIndex = 0
    
End Sub

Private Sub lstTipoIngreso_GotFocus()
    lstTipoIngreso.ListIndex = 0
End Sub

Private Sub lstTipoIngreso_LostFocus()
    lstTipoIngreso.ListIndex = -1
End Sub

Private Sub MyCombo1_Click()

End Sub

Private Sub mskRango_GotFocus(Index As Integer)
    mskRango(Index).Format = "dd/MM/yyyy"
    mskRango(Index).SelStart = 0
    mskRango(Index).SelLength = Len(mskRango(Index).Text)
End Sub

Private Sub mskRango_LostFocus(Index As Integer)
    mskRango(Index).Format = "dd/MMM/yyyy"
End Sub

Private Sub mskRango_Validate(Index As Integer, Cancel As Boolean)
    If Not IsDate(mskRango(0).Text) Then
        mskRango(0) = CDate("01/" & CStr(Month(fdtmServerFecha)) & "/" & CStr(Year(fdtmServerFecha)))
    Else
    
    End If
    If Not IsDate(mskRango(1).Text) Then
        mskRango(1) = fdtmServerFecha
    Else
    
    End If
End Sub

Private Sub optFecha_Click(Index As Integer)
    mskRango(0).Enabled = optFecha(0).Value
    mskRango(1).Enabled = optFecha(0).Value
End Sub

Private Sub pImprime(strDestino As String)
    Dim alstrParametros(21) As String
    Dim rsReporte As ADODB.Recordset
    Dim intIndex As Integer
    Dim strAgrupar As String
    Dim strTipoIngreso As String
    Dim strTipoIngreso2 As String
    Dim strParametros As String
    Dim rptReporte As CRAXDRT.Report

    If cgstrModulo = "PV" Or vgstrmodulo = "PV" Then
        If chkDetallado.Value = vbChecked Then
            pInstanciaReporte lrptReporteVer, "rptIngresosDetalleCuenta.rpt"
        Else
            pInstanciaReporte lrptReporteVer, "rptIngresosconmonto.rpt"
        End If
    Else
        If chkHorizontal.Value = vbChecked Then
            pInstanciaReporte lrptReporteVer, "rptGeneralIngresosPacientesHorizontal.rpt"
        Else
            pInstanciaReporte lrptReporteVer, "rptGeneralIngresosPacientes.rpt"
        End If
    End If
    
    Set rptReporte = lrptReporteVer

    For intIndex = 0 To 5
        If optAgrupar(intIndex).Value Then
            strAgrupar = CStr(intIndex)
            Exit For
        End If
    Next
    
    strAgrupar = IIf(strAgrupar = "0", "6", strAgrupar)
    strTipoIngreso = "_"
    strTipoIngreso2 = ""
    For intIndex = 0 To lstTipoIngreso.ListCount - 1
        If lstTipoIngreso.Selected(intIndex) Then
            strTipoIngreso = strTipoIngreso & lstTipoIngreso.ItemData(intIndex) & "_"
            strTipoIngreso2 = strTipoIngreso2 & lstTipoIngreso.List(intIndex) & ", "
        End If
    Next
    If strTipoIngreso2 <> "" Then
        strTipoIngreso2 = Mid(strTipoIngreso2, 1, Len(strTipoIngreso2) - 2)
    End If
    rptReporte.DiscardSavedData
    
    If cgstrModulo = "PV" Or vgstrmodulo = "PV" Then
        strParametros = Format(mskRango(0), "yyyy-MM-dd") & "|" & Format(mskRango(1), "yyyy-MM-dd") & "|" & IIf(optFecha(0).Value, "0", "1") & "|" & strTipoIngreso & "|" & cboProcedencia.ItemData(cboProcedencia.ListIndex) & "|" & cboMedico.ItemData(cboMedico.ListIndex) & "|" & cboArea.ItemData(cboArea.ListIndex) & "|" & IIf(optSexo(0).Value, "A", IIf(optSexo(1).Value, "M", "F")) & "|" & cboEmpresa.ItemData(cboEmpresa.ListIndex) & "|" & cboProcedencia2.ItemData(cboProcedencia2.ListIndex) & "|" & chkdetalle.Value
    Else
        strParametros = Format(mskRango(0), "yyyy-MM-dd") & "|" & Format(mskRango(1), "yyyy-MM-dd") & "|" & IIf(optFecha(0).Value, "0", "1") & "|" & strTipoIngreso & "|" & cboProcedencia.ItemData(cboProcedencia.ListIndex) & "|" & cboMedico.ItemData(cboMedico.ListIndex) & "|" & cboArea.ItemData(cboArea.ListIndex) & "|" & IIf(optSexo(0).Value, "A", IIf(optSexo(1).Value, "M", "F")) & "|" & chkAdmisionCancelada.Value & "|" & cboProcedencia2.ItemData(cboProcedencia2.ListIndex) & "|" & cboEmpresa.ItemData(cboEmpresa.ListIndex) & "|-1"
    End If
    
    alstrParametros(0) = "intGrupo1;" & strAgrupar
    alstrParametros(1) = "intOrden1;" & IIf(optOrdenar(0).Value, "1", "2")
    alstrParametros(2) = "blnInternos;" & IIf(optFecha(1).Value, "True", "False") & ";BOOLEAN"
    alstrParametros(3) = "dtmFechaIni;" & mskRango(0) & ";DATE"
    alstrParametros(4) = "dtmFechaFin;" & mskRango(1) & ";DATE"
    alstrParametros(5) = "strNombreHospital;" & cboEmpresa.Text 'Trim(vgstrNombreHospitalCH)
    alstrParametros(6) = "strProcedencia;" & Trim(cboProcedencia.Text)
    alstrParametros(7) = "strMedico;" & Trim(cboMedico.Text)
    alstrParametros(8) = "strArea;" & Trim(cboArea.Text)
    alstrParametros(9) = "strSexo;" & IIf(optSexo(0).Value, "AMBOS", IIf(optSexo(1).Value, "MASCULINO", "FEMENINO"))
    alstrParametros(10) = "strTipoIngreso;" & strTipoIngreso2
    alstrParametros(11) = "blnDetallado;" & IIf(chkDetallado.Value = vbChecked, "True", "False") & ";BOOLEAN"
    alstrParametros(12) = "blnDiagnostico;" & IIf(chkDiagnostico.Value = vbChecked, "True", "False") & ";BOOLEAN"
    alstrParametros(13) = "blnOcultarFiltros;" & IIf(chkOcultarFiltros.Value = vbChecked, "True", "False") & ";BOOLEAN"
    alstrParametros(14) = "MostrarOrdInt;" & IIf(ChkOrden.Value = vbChecked, 1, 0)
    alstrParametros(15) = "strProcedencia2;" & Trim(cboProcedencia2.Text)
    alstrParametros(16) = "blnOcultarEtiquetas;" & IIf(chkOcultaEtiquetas.Value = vbChecked, 1, 0)
    If cgstrModulo = "AD" Or vgstrmodulo = "AD" Then
        alstrParametros(17) = "blnFechaEgreso;" & IIf(chkFechaEgreso.Value = vbChecked, "True", "False") & ";BOOLEAN"
        alstrParametros(18) = "blnNoControl;" & IIf(chkNoControl.Value = vbChecked, "True", "False") & ";BOOLEAN"
        alstrParametros(19) = "blnNoPoliza;" & IIf(chkNoPoliza.Value = vbChecked, "True", "False") & ";BOOLEAN"
        alstrParametros(20) = "blnAutoriza;" & IIf(chkAutoriza.Value = vbChecked, "True", "False") & ";BOOLEAN"
        alstrParametros(21) = "blnTotalCuenta;" & IIf(chkTotalCuenta.Value = vbChecked, "True", "False") & ";BOOLEAN"
    End If
    pCargaParameterFields alstrParametros, rptReporte
    
    If cgstrModulo = "PV" Or vgstrmodulo = "PV" Then
        Set rsReporte = frsEjecuta_SP(strParametros, "sp_PvRptIngresosConMonto")
    Else
        Set rsReporte = frsEjecuta_SP(strParametros, "sp_ADWRptIngresosPacientes")
    End If
    
    If rsReporte.EOF Then
        MsgBox SIHOMsg(13), vbInformation, "Mensaje"
    Else
        pImprimeReporte rptReporte, rsReporte, strDestino, "Relación de cuentas de pacientes"
    End If
    rsReporte.Close

End Sub
