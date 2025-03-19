VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.1#0"; "MyCommandButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{FF14BD24-9F8A-41E3-B5B8-7F0D45EE9F16}#15.0#0"; "hsflatcontrols.ocx"
Begin VB.Form frmMaximoMinimo 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asignación de máximos y mínimos"
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14220
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   14220
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab 
      Height          =   10080
      Left            =   -15
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   -15
      Width           =   14300
      _ExtentX        =   25215
      _ExtentY        =   17780
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   2
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmMaximoMinimo.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraArticulos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdManejos"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraRequisitar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraRangoConsumo"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fraTipoAsignacion"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdCalcular"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fraRequisiciones"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "fraLocalizaciones"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "fraGrabar"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Frame6"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "fraFiltros"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "CommonDialog1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmMaximoMinimo.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frmBotonera"
      Tab(1).Control(1)=   "Frame1"
      Tab(1).ControlCount=   2
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   120
         Top             =   8280
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame fraFiltros 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2850
         Left            =   120
         TabIndex        =   27
         Top             =   0
         Width           =   14000
         Begin VB.CheckBox chkMostrarArti 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Mostrar solo artículos asignados al departamento"
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
            Height          =   700
            Left            =   9050
            TabIndex        =   12
            ToolTipText     =   "Mostrar solo artículos asignados al departamento"
            Top             =   2030
            Width           =   2800
         End
         Begin HSFlatControls.MyCombo cboLocalizacion 
            Height          =   375
            Left            =   1800
            TabIndex        =   5
            ToolTipText     =   "Selección de la localización"
            Top             =   1110
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   661
            Style           =   1
            Enabled         =   -1  'True
            Text            =   ""
            Sorted          =   0   'False
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
         Begin VB.OptionButton optOpcion 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Todos"
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
            Left            =   1780
            TabIndex        =   1
            ToolTipText     =   "Tipo de artículo"
            Top             =   760
            Width           =   850
         End
         Begin VB.OptionButton optOpcion 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Artículos"
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
            Left            =   2700
            TabIndex        =   2
            ToolTipText     =   "Tipo de artículo"
            Top             =   760
            Width           =   1130
         End
         Begin VB.OptionButton optOpcion 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Medicamentos"
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
            Left            =   3880
            TabIndex        =   3
            ToolTipText     =   "Tipo de artículo"
            Top             =   760
            Width           =   1750
         End
         Begin VB.OptionButton optOpcion 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Insumos"
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
            Left            =   5680
            TabIndex        =   4
            ToolTipText     =   "Tipo de artículo"
            Top             =   760
            Width           =   1065
         End
         Begin VB.TextBox txtClave 
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
            Height          =   375
            Left            =   9050
            TabIndex        =   11
            Top             =   700
            Width           =   4815
         End
         Begin VB.CheckBox chkVarias 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Varias localizaciones"
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
            Left            =   4800
            TabIndex        =   6
            Top             =   1170
            Width           =   2310
         End
         Begin VB.ListBox lstManejos 
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
            Height          =   885
            Left            =   9050
            Style           =   1  'Checkbox
            TabIndex        =   77
            Top             =   1110
            Width           =   4815
         End
         Begin VB.CheckBox chkCompradirecta 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Compra directa"
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
            Left            =   9050
            TabIndex        =   15
            Top             =   2500
            Width           =   2535
         End
         Begin MyCommandButton.MyButton cmdAgregar 
            Height          =   375
            Left            =   11870
            TabIndex        =   14
            ToolTipText     =   "Cargar los artículos según los filtros"
            Top             =   2025
            Width           =   1995
            _ExtentX        =   3519
            _ExtentY        =   661
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   1
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   16777215
            Caption         =   "Agregar artículos"
            DepthEvent      =   1
            ShowFocus       =   -1  'True
         End
         Begin HSFlatControls.MyCombo cboTipoRequisicion 
            Height          =   375
            Left            =   9050
            TabIndex        =   13
            Top             =   2025
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   661
            Style           =   1
            Enabled         =   -1  'True
            Text            =   ""
            Sorted          =   0   'False
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
         Begin HSFlatControls.MyCombo cboNombreComercial 
            Height          =   375
            Left            =   1800
            TabIndex        =   9
            ToolTipText     =   "Selección del artículo"
            Top             =   2325
            Width           =   5295
            _ExtentX        =   9340
            _ExtentY        =   661
            Enabled         =   -1  'True
            Text            =   ""
            Sorted          =   0   'False
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
         Begin HSFlatControls.MyCombo cboSubfamilia 
            Height          =   375
            Left            =   1800
            TabIndex        =   8
            ToolTipText     =   "Selección de la subfamilia"
            Top             =   1920
            Width           =   5295
            _ExtentX        =   9340
            _ExtentY        =   661
            Style           =   1
            Enabled         =   -1  'True
            Text            =   ""
            Sorted          =   0   'False
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
         Begin HSFlatControls.MyCombo cboDepartamento 
            Height          =   375
            Left            =   1800
            TabIndex        =   0
            ToolTipText     =   "Seleccione el departamento"
            Top             =   300
            Width           =   5295
            _ExtentX        =   9340
            _ExtentY        =   661
            Style           =   1
            Enabled         =   -1  'True
            Text            =   ""
            Sorted          =   0   'False
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
         Begin HSFlatControls.MyCombo cboFamilia 
            Height          =   375
            Left            =   1800
            TabIndex        =   7
            ToolTipText     =   "Selección de la familia"
            Top             =   1515
            Width           =   5295
            _ExtentX        =   9340
            _ExtentY        =   661
            Style           =   1
            Enabled         =   -1  'True
            Text            =   ""
            Sorted          =   0   'False
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
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Localización"
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
            TabIndex        =   28
            Top             =   1170
            Width           =   1215
         End
         Begin VB.Label lblFamilia 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Familia"
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
            TabIndex        =   29
            Top             =   1570
            Width           =   690
         End
         Begin VB.Label lblSubfamilia 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Subfamilia"
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
            TabIndex        =   30
            Top             =   1980
            Width           =   1005
         End
         Begin VB.Label lblNombreComercial 
            BackColor       =   &H80000005&
            Caption         =   "Nombre comercial"
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
            TabIndex        =   31
            Top             =   2380
            Width           =   1635
         End
         Begin VB.Label Label8 
            BackColor       =   &H80000005&
            Caption         =   "Nombre genérico"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7250
            TabIndex        =   32
            Top             =   360
            Width           =   1710
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Clave"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7250
            TabIndex        =   33
            Top             =   765
            Width           =   585
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Departamento"
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
            TabIndex        =   35
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label lblNombreGenerico 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   375
            Left            =   9050
            TabIndex        =   10
            Top             =   300
            Width           =   4815
         End
         Begin VB.Label lblTipoRequisicion 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Tipo requisición"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7250
            TabIndex        =   47
            Top             =   2085
            Width           =   1470
         End
         Begin VB.Label Label1 
            BackColor       =   &H80000005&
            Caption         =   "Manejos"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7250
            TabIndex        =   71
            Top             =   1170
            Width           =   1320
         End
      End
      Begin VB.Frame Frame6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   " Nombre comercial completo"
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
         Height          =   960
         Left            =   120
         TabIndex        =   72
         Top             =   7160
         Width           =   12495
         Begin VB.Label lblNombreComercialCompleto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   585
            Left            =   120
            TabIndex        =   73
            ToolTipText     =   "Nombre completo del artículo"
            Top             =   240
            Width           =   12255
         End
      End
      Begin VB.Frame fraGrabar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   840
         Left            =   2150
         TabIndex        =   67
         Top             =   8060
         Width           =   9900
         Begin MyCommandButton.MyButton cmdAsignarDesactivar 
            Height          =   600
            Index           =   0
            Left            =   660
            TabIndex        =   68
            Top             =   200
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   1058
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   1
            BackColorDown   =   -2147483643
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   16777215
            Caption         =   "Asignar al departamento"
            DepthEvent      =   1
         End
         Begin MyCommandButton.MyButton cmdGrabar 
            Height          =   600
            Left            =   60
            TabIndex        =   69
            ToolTipText     =   "Guardar"
            Top             =   200
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
            Picture         =   "frmMaximoMinimo.frx":0038
            BackColorDown   =   -2147483643
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   16777215
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmMaximoMinimo.frx":09BC
            PictureAlignment=   4
            PictureDisabledEffect=   0
         End
         Begin MyCommandButton.MyButton cmdAsignarDesactivar 
            Height          =   600
            Index           =   1
            Left            =   2950
            TabIndex        =   70
            Top             =   200
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   1058
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   1
            BackColorDown   =   -2147483643
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   16777215
            Caption         =   "Desasignar del departamento"
            DepthEvent      =   1
         End
         Begin MyCommandButton.MyButton cmdExportar 
            Height          =   600
            Index           =   2
            Left            =   5250
            TabIndex        =   75
            ToolTipText     =   "Exportar listado a excel"
            Top             =   200
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   1058
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   1
            BackColorDown   =   -2147483643
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   16777215
            Caption         =   "Exportar"
            DepthEvent      =   1
         End
         Begin MyCommandButton.MyButton cmdImportar 
            Height          =   600
            Index           =   0
            Left            =   7550
            TabIndex        =   76
            ToolTipText     =   "Importar listado de excel"
            Top             =   200
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   1058
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   1
            BackColorDown   =   -2147483643
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   16777215
            Caption         =   "Importar"
            DepthEvent      =   1
         End
      End
      Begin VB.Frame fraLocalizaciones 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3615
         Left            =   3840
         TabIndex        =   53
         Top             =   2760
         Visible         =   0   'False
         Width           =   6495
         Begin VB.Frame freBotones 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   2640
            Left            =   2880
            TabIndex        =   57
            Top             =   240
            Width           =   750
            Begin MyCommandButton.MyButton cmdAgregarTodas 
               Height          =   600
               Left            =   80
               TabIndex        =   60
               ToolTipText     =   "Agregar todas las localizaciones"
               Top             =   200
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
               Picture         =   "frmMaximoMinimo.frx":1340
               BackColorOver   =   -2147483633
               BackColorFocus  =   -2147483633
               BackColorDisabled=   -2147483633
               BorderColor     =   -2147483627
               TransparentColor=   16777215
               Caption         =   ""
               DepthEvent      =   1
               PictureDisabled =   "frmMaximoMinimo.frx":1CC2
               PictureAlignment=   4
               PictureDisabledEffect=   0
               ShowFocus       =   -1  'True
            End
            Begin MyCommandButton.MyButton cmdSelecciona 
               Height          =   600
               Index           =   0
               Left            =   80
               TabIndex        =   58
               ToolTipText     =   "Agregar una localización"
               Top             =   800
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
               Picture         =   "frmMaximoMinimo.frx":2644
               BackColorOver   =   -2147483633
               BackColorFocus  =   -2147483633
               BackColorDisabled=   -2147483633
               BorderColor     =   -2147483627
               TransparentColor=   16777215
               Caption         =   ""
               DepthEvent      =   1
               DropDownPicture =   "frmMaximoMinimo.frx":2FC6
               PictureDisabled =   "frmMaximoMinimo.frx":2FE2
               PictureAlignment=   4
               PictureDisabledEffect=   0
               ShowFocus       =   -1  'True
            End
            Begin MyCommandButton.MyButton cmdDeleteAll 
               Height          =   600
               Left            =   75
               TabIndex        =   61
               ToolTipText     =   "Quitar todas las localizaciones"
               Top             =   2000
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
               Picture         =   "frmMaximoMinimo.frx":3964
               BackColorOver   =   -2147483633
               BackColorFocus  =   -2147483633
               BackColorDisabled=   -2147483633
               BorderColor     =   -2147483627
               TransparentColor=   16777215
               Caption         =   ""
               DepthEvent      =   1
               PictureDisabled =   "frmMaximoMinimo.frx":42E6
               PictureAlignment=   4
               PictureDisabledEffect=   0
               ShowFocus       =   -1  'True
            End
            Begin MyCommandButton.MyButton cmdSelecciona 
               Height          =   600
               Index           =   1
               Left            =   75
               TabIndex        =   59
               ToolTipText     =   "Quitar localización"
               Top             =   1400
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
               Picture         =   "frmMaximoMinimo.frx":4C68
               BackColorOver   =   -2147483633
               BackColorFocus  =   -2147483633
               BackColorDisabled=   -2147483633
               BorderColor     =   -2147483627
               TransparentColor=   16777215
               Caption         =   ""
               DepthEvent      =   1
               PictureDisabled =   "frmMaximoMinimo.frx":55EA
               PictureAlignment=   4
               PictureDisabledEffect=   0
               ShowFocus       =   -1  'True
            End
         End
         Begin VB.ListBox lstDisponibles 
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
            Height          =   2835
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   54
            Top             =   240
            Width           =   2670
         End
         Begin VB.ListBox lstSeleccionadas 
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
            Height          =   2835
            Left            =   3720
            Sorted          =   -1  'True
            TabIndex        =   55
            Top             =   240
            Width           =   2665
         End
         Begin MyCommandButton.MyButton cmdAceptar 
            Height          =   375
            Left            =   2640
            TabIndex        =   56
            Top             =   3120
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   661
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   1
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   16777215
            Caption         =   "Aceptar"
            DepthEvent      =   1
            ShowFocus       =   -1  'True
         End
      End
      Begin VB.Frame fraRequisiciones 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Requisiciones del artículo"
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
         Height          =   2880
         Left            =   2520
         TabIndex        =   49
         Top             =   4920
         Width           =   9285
         Begin MyCommandButton.MyButton cmdCerrar 
            Height          =   375
            Left            =   4005
            TabIndex        =   52
            Top             =   2400
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   661
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   1
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   16777215
            Caption         =   "Cerrar"
            DepthEvent      =   1
            ShowFocus       =   -1  'True
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdConsultaReq 
            Height          =   1740
            Left            =   75
            TabIndex        =   51
            Top             =   600
            Width           =   9120
            _ExtentX        =   16087
            _ExtentY        =   3069
            _Version        =   393216
            ForeColor       =   0
            Cols            =   6
            ForeColorFixed  =   0
            ForeColorSel    =   0
            BackColorBkg    =   -2147483643
            BackColorUnpopulated=   -2147483643
            GridColor       =   -2147483638
            GridColorFixed  =   -2147483638
            GridColorUnpopulated=   -2147483638
            GridLinesFixed  =   1
            GridLinesUnpopulated=   1
            Appearance      =   0
            FormatString    =   "|Fecha|Requisición|Tipo|Departamento que surte|Persona requisitó"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   6
         End
         Begin VB.Label lblNombreArticulo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   300
            Left            =   80
            TabIndex        =   50
            Top             =   300
            Width           =   9120
         End
      End
      Begin MyCommandButton.MyButton cmdCalcular 
         Height          =   375
         Left            =   8520
         TabIndex        =   20
         ToolTipText     =   "Calcular el consumo"
         Top             =   3200
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         BackColorOver   =   -2147483633
         BackColorFocus  =   -2147483633
         BackColorDisabled=   -2147483633
         BorderColor     =   -2147483627
         TransparentColor=   16777215
         Caption         =   "Calcular el consumo"
         DepthEvent      =   1
         ShowFocus       =   -1  'True
      End
      Begin VB.Frame fraTipoAsignacion 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Asignar máximos y mínimos "
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
         Height          =   695
         Left            =   120
         TabIndex        =   25
         Top             =   2880
         Width           =   4695
         Begin VB.OptionButton optconsumo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Con base en el consumo"
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
            Left            =   1800
            TabIndex        =   17
            ToolTipText     =   "Asignar con base en el consumo"
            Top             =   300
            Value           =   -1  'True
            Width           =   2745
         End
         Begin VB.OptionButton optconsumo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Manualmente"
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
            Left            =   120
            TabIndex        =   16
            ToolTipText     =   "Asignación manual del máximo y mínimo"
            Top             =   300
            Width           =   1695
         End
      End
      Begin VB.Frame fraRangoConsumo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Rango de fechas de consumo"
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
         Height          =   695
         Left            =   4920
         TabIndex        =   26
         Top             =   2880
         Width           =   3525
         Begin MSMask.MaskEdBox mskInicio 
            Height          =   375
            Left            =   120
            TabIndex        =   18
            ToolTipText     =   "Fecha inicial"
            Top             =   250
            Width           =   1380
            _ExtentX        =   2434
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskFin 
            Height          =   375
            Left            =   2040
            TabIndex        =   19
            ToolTipText     =   "Fecha inicial"
            Top             =   255
            Width           =   1380
            _ExtentX        =   2434
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
            PromptChar      =   " "
         End
         Begin VB.Label lblAl 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            Height          =   255
            Left            =   1680
            TabIndex        =   34
            Top             =   315
            Width           =   180
         End
      End
      Begin VB.Frame fraRequisitar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   840
         Left            =   5000
         TabIndex        =   38
         Top             =   8060
         Width           =   4230
         Begin MyCommandButton.MyButton cmdConsultarReq 
            Height          =   600
            Left            =   60
            TabIndex        =   48
            ToolTipText     =   "Consultar las requisiciones donde está incluido el artículo"
            Top             =   200
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   1058
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   1
            BackColorDown   =   -2147483643
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   16777215
            Caption         =   "Consultar requisición"
            DepthEvent      =   1
         End
         Begin MyCommandButton.MyButton cmdRequisitar 
            Height          =   600
            Left            =   1880
            TabIndex        =   37
            ToolTipText     =   "Hacer la requisición por faltantes"
            Top             =   200
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   1058
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   1
            BackColorDown   =   -2147483643
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   16777215
            Caption         =   "Enviar requisición"
            DepthEvent      =   1
         End
         Begin MyCommandButton.MyButton cmdVistaPreliminar 
            Height          =   600
            Left            =   3570
            TabIndex        =   65
            ToolTipText     =   "Vista preliminar del reporte"
            Top             =   200
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
            Picture         =   "frmMaximoMinimo.frx":5F6C
            BackColorDown   =   -2147483643
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   16777215
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmMaximoMinimo.frx":68F0
            PictureAlignment=   4
            PictureDisabledEffect=   0
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   8060
         Left            =   -74880
         TabIndex        =   42
         Top             =   0
         Width           =   13995
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdRequisicion 
            Height          =   1515
            Left            =   840
            TabIndex        =   46
            Top             =   5520
            Visible         =   0   'False
            Width           =   10845
            _ExtentX        =   19129
            _ExtentY        =   2672
            _Version        =   393216
            ForeColor       =   0
            Cols            =   5
            ForeColorFixed  =   0
            ForeColorSel    =   0
            BackColorBkg    =   -2147483643
            BackColorUnpopulated=   -2147483643
            GridColor       =   -2147483638
            GridColorFixed  =   -2147483638
            GridColorUnpopulated=   -2147483638
            GridLinesFixed  =   1
            GridLinesUnpopulated=   1
            BorderStyle     =   0
            Appearance      =   0
            FormatString    =   "|CveAlmacen|CveArticulo|Cantidad|TipoUnidad"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   5
         End
         Begin VSFlex7LCtl.VSFlexGrid VsfBusArticulos 
            Height          =   7695
            Left            =   120
            TabIndex        =   43
            ToolTipText     =   "Artículos que están por debajo o iguales al punto de reorden"
            Top             =   240
            Width           =   13770
            _cx             =   24289
            _cy             =   13573
            _ConvInfo       =   1
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483638
            GridColorFixed  =   -2147483638
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   15
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmMaximoMinimo.frx":7272
            ScrollTrack     =   -1  'True
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   0
            AutoSearch      =   0
            AutoSearchDelay =   2
            MultiTotals     =   -1  'True
            SubtotalPosition=   1
            OutlineBar      =   0
            OutlineCol      =   0
            Ellipsis        =   0
            ExplorerBar     =   5
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   0
            Editable        =   0
            ShowComboButton =   -1  'True
            WordWrap        =   0   'False
            TextStyle       =   0
            TextStyleFixed  =   0
            OleDragMode     =   0
            OleDropMode     =   0
            ComboSearch     =   3
            AutoSizeMouse   =   -1  'True
            FrozenRows      =   0
            FrozenCols      =   0
            AllowUserFreezing=   0
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            Begin HSFlatControls.MyCombo cboUnidad 
               Height          =   375
               Left            =   7200
               TabIndex        =   66
               Top             =   1560
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   661
               Style           =   1
               Enabled         =   -1  'True
               Text            =   ""
               Sorted          =   0   'False
               List            =   $"frmMaximoMinimo.frx":7346
               ItemData        =   $"frmMaximoMinimo.frx":734D
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
         End
      End
      Begin MyCommandButton.MyButton cmdManejos 
         Height          =   375
         Left            =   12720
         TabIndex        =   74
         Top             =   7470
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         BackColorOver   =   -2147483633
         BackColorFocus  =   -2147483633
         BackColorDisabled=   -2147483633
         BorderColor     =   -2147483627
         TransparentColor=   16777215
         Caption         =   "Manejos"
         DepthEvent      =   1
         ShowFocus       =   -1  'True
      End
      Begin VB.Frame frmBotonera 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   780
         Left            =   -69120
         TabIndex        =   45
         Top             =   8110
         Width           =   2520
         Begin MyCommandButton.MyButton cmdImprimir 
            Height          =   600
            Left            =   1860
            TabIndex        =   64
            ToolTipText     =   "Imprimir localmente"
            Top             =   140
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
            Picture         =   "frmMaximoMinimo.frx":7354
            BackColorDown   =   -2147483643
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   16777215
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmMaximoMinimo.frx":7CD8
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdDeshacer 
            Height          =   600
            Left            =   1260
            TabIndex        =   63
            ToolTipText     =   "Deshacer cambios"
            Top             =   140
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
            Picture         =   "frmMaximoMinimo.frx":865A
            BackColorDown   =   -2147483643
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   16777215
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmMaximoMinimo.frx":8FDE
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdDelete 
            Height          =   600
            Left            =   660
            TabIndex        =   62
            ToolTipText     =   "Eliminar registro"
            Top             =   140
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
            Picture         =   "frmMaximoMinimo.frx":9962
            BackColorDown   =   -2147483643
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   16777215
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmMaximoMinimo.frx":A2E4
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdEnviarRequisicion 
            Height          =   600
            Left            =   60
            TabIndex        =   44
            ToolTipText     =   "Grabar y enviar la requisición"
            Top             =   140
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
            Picture         =   "frmMaximoMinimo.frx":AC68
            BackColorDown   =   -2147483643
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   16777215
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmMaximoMinimo.frx":B5EC
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
      End
      Begin VB.Frame fraArticulos 
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
         Height          =   4400
         Left            =   120
         TabIndex        =   36
         Top             =   2760
         Width           =   14000
         Begin MyCommandButton.MyButton cmdInvertirSeleccion 
            Height          =   375
            Left            =   11700
            TabIndex        =   23
            Top             =   3140
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   1
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   16777215
            Caption         =   "Invertir selección"
            CaptionPosition =   4
            DepthEvent      =   1
            ForeColorDisabled=   -2147483629
            ForeColorOver   =   13003064
            ForeColorFocus  =   13003064
            ForeColorDown   =   13003064
            PictureAlignment=   4
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdMarcar 
            Height          =   375
            Left            =   9500
            TabIndex        =   22
            Top             =   3140
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   1
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   16777215
            Caption         =   "Marcar / desmarcar"
            DepthEvent      =   1
            ShowFocus       =   -1  'True
         End
         Begin VSFlex7LCtl.VSFlexGrid vsfArticulos 
            Height          =   4035
            Left            =   120
            TabIndex        =   21
            ToolTipText     =   "Lista de artículos"
            Top             =   240
            Width           =   13755
            _cx             =   24262
            _cy             =   7117
            _ConvInfo       =   1
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483638
            GridColorFixed  =   -2147483638
            TreeColor       =   -2147483632
            FloodColor      =   49152
            SheetBorder     =   -2147483643
            FocusRect       =   2
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   15
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmMaximoMinimo.frx":BF70
            ScrollTrack     =   -1  'True
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   0
            AutoSearch      =   0
            AutoSearchDelay =   2
            MultiTotals     =   -1  'True
            SubtotalPosition=   1
            OutlineBar      =   0
            OutlineCol      =   0
            Ellipsis        =   0
            ExplorerBar     =   5
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   0
            Editable        =   1
            ShowComboButton =   -1  'True
            WordWrap        =   0   'False
            TextStyle       =   0
            TextStyleFixed  =   0
            OleDragMode     =   0
            OleDropMode     =   0
            ComboSearch     =   3
            AutoSizeMouse   =   -1  'True
            FrozenRows      =   0
            FrozenCols      =   0
            AllowUserFreezing=   0
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            Begin VB.Frame fraBarra 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BeginProperty Font 
                  Name            =   "Century Gothic"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   735
               Left            =   2400
               TabIndex        =   39
               Top             =   840
               Width           =   8940
               Begin MSComctlLib.ProgressBar pgbBarra 
                  Height          =   150
                  Left            =   75
                  TabIndex        =   40
                  Top             =   525
                  Width           =   8790
                  _ExtentX        =   15505
                  _ExtentY        =   265
                  _Version        =   393216
                  Appearance      =   0
                  Scrolling       =   1
               End
               Begin VB.Label lblTituloBarra 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
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
                  Height          =   300
                  Left            =   75
                  TabIndex        =   41
                  Top             =   195
                  Width           =   8775
               End
            End
         End
      End
   End
End
Attribute VB_Name = "frmMaximoMinimo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const clngTopModoAsignacion = 3500
Const clngHeightModoAsignacion = 3610
Const clngHeightVSFModoAsignacion = 2820

Const clngTopModoConsulta = 2760
Const clngHeightModoConsulta = 4400
Const clngHeightVSFModoConsulta = 4035

Const clngColorNegro = &H80000008
Const clngColorRojo = &HC0&
Const clngColorRojoOscuro = &H80&
Const clngColorVerde = &H8000&
Const clngColorVerdeBrillante = &HC000&
Const clngColorAzul = &HFF0000

Const cstrExcedida = "EXCEDIDA"
Const cstrOptima = "OPTIMA"
Const cstrSolicitado = "SOLICITADO"
Const cstrReorden = "REORDEN"
Const cstrFaltante = "FALTANTE"
Const cstrInsuficiente = "INSUFICIENTE"

Const cintItemDataReubicacion = 0           'Itemdata para el tipo de reubicación REUBICACION
Const cintItemDataCompra = 1                'Itemdata para el tipo de reubicación COMPRA - PEDIDO

Const xlRight_ = -4152                      'Alineación derecha

'Columnas vsfArticulos:                     (VSF de los artículos que se van a incluir como máximos, o que se consultan)
'Manejos
Const cintColVsfManejo1 = 1                 '
Const cintColVsfManejo2 = 2                 '
Const cintColVsfManejo3 = 3                 '
Const cintColVsfManejo4 = 4                 '

Const cintColVsfIdArticulo = 5              'Id. del articulo, IvArticulo.intIdArticulo
Const cintColVsfNombreComercial = 6         'Nombre comercial
'Máximo:
Const cintColVsfCapturaMaximo = 7           'Dato capturado como máximo
Const cintColVsfCapturaUnidadMax = 8        'Descripción de la unidad seleccionada
Const cintColVsfCapturaCveUnidadMax = 9     'Clave de la unidad seleccionada
'Punto de reorden:
Const cintColVsfCapturaPunto = 10            'Dato capturado como punto de reorden
Const cintColVsfCapturaUnidadPun = 11        'Descripción de la unidad seleccionada
Const cintColVsfCapturaCveUnidadPun = 12     'Clave de la unidad seleccionada
'Mínimo:
Const cintColVsfCapturaMinimo = 13           'Dato capturado como mínimo
Const cintColVsfCapturaUnidadMin = 14       'Descripción de la unidad seleccionada
Const cintColVsfCapturaCveUnidadMin = 15    'Clave de la unidad seleccionada
'Almacén:
Const cintColVsfCapturaAlmacen = 16         'Almacén que surtirá el artículo
Const cintColVsfCapturaCveAlmacen = 17      'Clave del almacén seleccionado
'Departamento compra:
Const cintColVsfCapturaDeptoCompra = 18     'Departamento que comprará el artículo
Const cintColVsfCapturaCveDeptoCompra = 19  'Clave del departamento seleccionado
'Otros datos:
Const cintColVsfClaveArt = 20               'Clave del artículo, IvArticulo.chrCveArticulo
Const cintColVsfContenidoArt = 21           'Contenido del artículo, IvArticulo.intContenido
Const cintColVsfCveUniAlternaArt = 22       'Clave de la unidad alterna en la que se maneja el artículo
Const cintColvsfUnidadAlternaArt = 23       'Descripción de la unidad alterna en la que se maneja el artículo
Const cintColVsfCveUniMinimaArt = 24        'Clave de la unidad mínima en la que se maneja el artículo
Const cintColvsfUnidadMinimaArt = 25        'Descripción de la unidad mínima en la que se maneja el artículo
Const cintColvsfCapturaTipoUniMax = 26      'Tipo de unidad seleccionada para el máximo, A = alterna, M = mínima
Const cintColvsfCapturaTipoUniPun = 27      'Tipo de unidad seleccionada para el punto de reorden, A = alterna, M = mínima
Const cintColvsfCapturaTipoUniMin = 28      'Tipo de unidad seleccionada para el mínimo, A = alterna, M = mínima
Const cintcolvsfGuardar = 29                'Indica que se guardará este registro
Const cintColVsfExistencia = 30             'Existencia (alterna o mínima según el tipo de requisición y artículo) en el departamento
Const cintColVsfExistenciaUnidad = 31       'Unidad en la que está expresada la existencia
Const cintColvsfEstadoArt = 32              'Estatus del artículo respecto al máximo y la existencia
Const cintColvsfRequisitar = 33             'Indica que de este artículo se generará requisición
Const cintColvsfCveArt = 34                 'Clave del artículo, dato duplicado pero sirve en algunos casos
Const cintColvsfControlado = 35             'Estatus de medicamento controlado
Const cintColvsfExistenciaAlterna = 36      'Existencia enss unidades alternas
Const cintColvsfExistenciaMinima = 37       'Existencia en unidades mínimas
Const cintColvsfLocalizacion = 38           'Localizacion del artículo
Const cintColsVsfArticulos = 39

Const cintColVsfArtiExisEsclavos = 39
Const cintColVsfArtiMaxEsclavos = 40
Const cintColVsfArtiReordenEsclavos = 41
Const cintColVsfArtiDeptoEsclavos = 42
Const cintColVsfArtiMinEsclavos = 43
Const cintColsVsfArtiBusArtiConsol = 44


Const cstrTitulosVsfArticulos = "Id|Nombre|Máximo|Unidad|Id|Reorden|Unidad|Id|Mínimo|Unidad|Id|Almacén reubica|Id|Departamento compra|Id|Clave|Contenido|CveAlterna|Nombre|CveMinima|Nombre|TipoUniMax|TipoUniPun|TipoUniMin|Guardar|Existencia|Unidad|Estado|Requisitar|CveArt|Controlado|ExistAlterna|ExistMinima|Localización"
Const cstrTitulosVsfArticulos2 = "Id|Nombre|Máximo|Unidad|Id|Reorden|Unidad|Id|Mínimo|Unidad|Id|Almacén recibe|Id|Departamento compra|Id|Clave|Contenido|CveAlterna|Nombre|CveMinima|Nombre|TipoUniMax|TipoUniPun|TipoUniMin|Guardar|Existencia|Unidad|Estado|Requisitar|CveArt|Controlado|ExistAlterna|ExistMinima|Localización"

'Columnas vsfBusArticulos:                  (VSF de los artículos que se van a requisitar)
Const cintColBusClave = 1                   'Clave del artículo, IvArticulo.chrCveArticulo
Const cintColBusNombre = 2                  'Nombre comercial
Const cintColBusLocalizacion = 3            'Localizacion del artículo
Const cintColBusAlmacen = 4                 'Almacén que surtirá
Const cintColBusIdAlmacen = 5               'Clave del almacén que surtirá
Const cintColBusExist = 6                   'Existencia del artículo
Const cintColBusUnidadExist = 7             'Unidad correspondiente a la existencia del artículo
Const cintColBusPendiente = 8               'Cantidad pendiente de surtir
Const cintColBusUnidadPendiente = 9         'Unidad del pendiente de surtir
Const cintColBusCantidad = 10               'Cantidad a requisitar
Const cintColBusUnidad = 11                 'Unidad de la requisición, descripción
Const cintColBusTipoUnidad = 12             'Tipo de unidad, A = alterna, M = mínima
Const cintColBusCveDeptoAutoriza = 13       'Clave del departamento que autoriza la requisición tipo compra
Const cintColBusCveDeptoRecibe = 14         'Clave del departamento que recibe la compra de los artículos
Const cintColBusCantidadOriginal = 15       'Cantidad original a requisitar
Const cintColsVsfBusArticulos = 16

Const cintColBusExisPrin = 16
Const cintColBusMaxPrin = 17
Const cintColBusExisEsclavos = 18
Const cintColBusMaxEsclavos = 19
Const cintColBusDeptoEsclavos = 20
Const cintColsVsfBusArtiConsol = 21

Const cstrTitulosVsfBusArticulos = "|Clave|Nombre|Localización|Departamento surte|Cve. almacen|Existencia|Unidad|Pendiente surtir|Unidad|Cantidad faltante|Unidad|Tipo|CveDeptoAutoriza|CveDeptoRecibe"

'Columnas grdRequisicion:
Const cintColGrdCveDeptoSurte = 1
Const cintColGrdCveArticulo = 2
Const cintColGrdCantidad = 3
Const cintColGrdTipoUnidad = 4
Const cintColGrdCveDeptoAutoriza = 5
Const cintColGrdCveDeptoRecibe = 6
Const cintColsRequisicion = 7

Const cintColReqArtiExisPrin = 7
Const cintColReqArtiMaxPrin = 8
Const cintColReqArtiExisEsclavos = 9
Const cintColReqArtiMaxEsclavos = 10
Const cintColReqArtiDeptoEsclavos = 11
Const cintColReqArtiConsol = 12

'Columnas grdConsultaReq:
Const cintColConsFecha = 1
Const cintColConsRequisicion = 2
Const cintColConsTipo = 3
Const cintColConsDepartamento = 4
Const cintColConsPersona = 5
Const cintColsConsulta = 6
Const cstrTitulosConsulta = "|Fecha|Número|Tipo|Departamento surte|Persona solicitó"

Private vgrptReporte As CRAXDRT.Report

Public lstrModo As String                   'Modo en que se manda llamar la pantalla, A = asignacion, C = consulta

Dim rs As New ADODB.Recordset               'Usos varios

Dim lblnEntrando As Boolean                 'Para saber si se está entrando a la pantalla

Dim lintTipoArticulo As Integer             'Indica el tipo de artículo, medicamento, insumo
Dim intMaxManejos As Integer
Dim lintCiclos As Integer

Dim lblnAlmacenConsigna As Boolean
Dim lblnPermisoAsignar As Boolean           'Para saber si el usuario tiene permisos para asignar artículos al almacén
Dim lblnBusquedaNombre As Boolean           'Indica si se ejecutó la búsqueda por nombre
Dim blnNoCargarFiltros As Boolean
Dim lbFlag As Boolean
Dim lbDeptoConSigna                         'Indica si el Departamento Seleccionado es de consignación
Dim lbConSigna As Boolean                   'Indica si el Almacén Reubica es de consignación




Dim lstrTotalLocalizaciones As String       'guarda las claves de la localizacion cuando se filtra por varias localizaciones
Dim lstrTipoRequisicion As String           'Tipo de requisicion a generar, RE = reubicación, CO = compra - pedido
Dim lstrDeptosReubican As String            'Lista de almacenes que reubican el artículo
Dim lstrDeptosConsigna As String            'Lista de departamentos de consignación
Dim lstrDeptosCompran As String             'Lista de departamentos que compran el artículo
Dim lstrCodigoBarras As String
Dim lstrDeptoCompras As String             'Departamento Compras configurado S o N
Dim strTodosManejos As String
Dim strUnidad As String
Dim strClavesConsol As String               'Almacenes esclavos consolidacion de almacenes
Dim lbCveAlmacenPrin As Integer


Dim llngNumOpcionRequisitar As Long         'Número de opción para requisitar, según el módulo de donde se llame
Dim llngNumOpcionGuardar As Long            'Número de opción para guardar datos, segun el módulo de donde se llame
Dim llngNumOpcionAsignar As Long            'Número de opción para asignar articulos al almacén, según el módulo de donde se llame
Dim llngTotalDesactivar As Long             'No. de marcas en la columna fija, para desactivar en depto
Dim llngTotalMarcados As Long               'No. de marcas en la columna fija, para asignar al depto
Dim llngTotalRequerir As Long               'Número de artículos a requisitar
Dim llngPersonaGraba As Long                'Clave del empleado que guarda la información
Dim liTemp2 As Long
Dim liTemp As Long

Dim ldblAumento  As Double                  'Para calcular el estado del progress bar
Dim blnPermisoCompraDirecta As Boolean  'Si el usuario tiene permiso para realizar requisiciones compra directa

Private Type EstRequisiciones
    lgnNumRequisicion As Long
End Type

Dim AsignacionAutomatica As Boolean
Dim aRequisiciones() As EstRequisiciones

Private Type DeptoArtiConsol
    intCveDepto  As Integer
    SrtCveArticulo As String
    intCantidad  As Integer
End Type
Dim aDeptoArtiConsol() As DeptoArtiConsol

Dim vlstrColSort As String
Dim lblnReubicarInsumos As Boolean          'Indica si se pueden reubicar insumos (cuando el departamento que solicita es almacén)
Dim llngDeptoSubrogado As Long              'Indica el depto subrogado configurado en el parametro de IV
Dim lintInterfazFarmaciaSJP As Integer      'Indica si se usa la interfaz de farmacia subrogada de SJP

Dim lblnProductoExiste As Boolean
Dim oExcel As Object
Dim oLibro As Object
Dim oHoja As Object
    

Private Sub pCargarOpciones()
On Error GoTo NotificaError

    llngNumOpcionGuardar = 0
    llngNumOpcionAsignar = 0
    llngNumOpcionRequisitar = 0
    
    Select Case cgstrModulo
    Case "IV"
        llngNumOpcionGuardar = 1032
        llngNumOpcionAsignar = 1982
        llngNumOpcionRequisitar = 1082
    Case "AD"
        llngNumOpcionRequisitar = 752
    Case "BS"
        llngNumOpcionRequisitar = 1299
    Case "PV"
        llngNumOpcionRequisitar = 327
    Case "CA"
        llngNumOpcionRequisitar = 916
    Case "CO"
        llngNumOpcionRequisitar = 855
    Case "CN"
        llngNumOpcionRequisitar = 1739
    Case "CC"
        llngNumOpcionRequisitar = 619
    Case "CP"
        llngNumOpcionRequisitar = 3
    Case "DI"
        llngNumOpcionRequisitar = 1619
    Case "EX"
        llngNumOpcionGuardar = 488
        llngNumOpcionAsignar = 1986
        llngNumOpcionRequisitar = 454
    Case "IM"
        llngNumOpcionRequisitar = 65
    Case "LA"
        llngNumOpcionRequisitar = 126
    Case "NO"
        llngNumOpcionRequisitar = 1873
    Case "SI"
        llngNumOpcionGuardar = 1117
        llngNumOpcionRequisitar = 1237
    Case "SE"
        llngNumOpcionRequisitar = 1569
    Case "TS"
        llngNumOpcionRequisitar = 1613
    End Select

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargarOpciones"))
End Sub

Private Sub pCargaDeptos()
1         On Error GoTo NotificaError
          
          Dim rs As New ADODB.Recordset
          Dim rsIvparametro As New ADODB.Recordset
          Dim strSentencia As String
          Dim strAlmaConsolidar As String
          Dim IntCvePrincipal As Integer
          Dim IntBitConsolAlmacen As Integer
          
2         If lstrModo = "A" Or lblnAlmacenConsigna Then
              'En modo de asignación, cargar los deptos donde el usuario puede manipular datos
3             vgstrParametrosSP = CStr(vglngNumeroLogin)
4             Set rs = frsEjecuta_SP(vgstrParametrosSP, "SP_IVSELDEPTOSMAXMIN")
5             If rs.RecordCount <> 0 Then
6                 pLlenarCboRs_new cboDepartamento, rs, 0, 1
7             Else
8                 If lstrModo = "C" Then
9                     cboDepartamento.AddItem vgstrNombreDepartamento
10                    cboDepartamento.ItemData(cboDepartamento.NewIndex) = vgintNumeroDepartamento
11                End If
12            End If
13        Else
              'En modo consulta antes de requisitar
              
14            cboDepartamento.Clear
              '| Carga el combo con los departamentos de la empresa, activos
              '| y clasificados como Almacén o como Enfermería que manejen Stock
              
              Set rsIvparametro = frsRegresaRs("SELECT BITCONSOLIDARALMACEN,TNYALMACENPRINCIPAL,VCHCVESALMACENCONSOLIDAR as claves FROM IvParametro WHERE tnyClaveEmpresa = " & CStr(vgintClaveEmpresaContable), adLockOptimistic, adOpenKeyset)
                If rsIvparametro.RecordCount > 0 Then
                   If IIf(IsNull(rsIvparametro!BITCONSOLIDARALMACEN), 0, rsIvparametro!BITCONSOLIDARALMACEN) = 1 Then
                   
                      strAlmaConsolidar = IIf(IsNull(rsIvparametro!claves), "", rsIvparametro!claves)
                      IntCvePrincipal = IIf(IsNull(rsIvparametro!TNYALMACENPRINCIPAL), 0, rsIvparametro!TNYALMACENPRINCIPAL)
                      IntBitConsolAlmacen = IIf(IsNull(rsIvparametro!BITCONSOLIDARALMACEN), 0, rsIvparametro!BITCONSOLIDARALMACEN)
                      
                         If Trim(strAlmaConsolidar) <> "" Then
                           strAlmaConsolidar = " And SMICVEDEPARTAMENTO NOT IN (" & rsIvparametro!claves & " )"
                         End If
                   End If
              End If
              
15            strSentencia = "Select smiCveDepartamento " & _
                             "  , Trim(vchDescripcion) " & _
                             "  From NoDepartamento " & _
                             " Where (chrClasificacion = 'A' Or (chrClasificacion = 'E' and chrEnfermeria = 'E')) " & _
                             "   And bitEstatus = 1 " & _
                             "   And NoDepartamento.BITCONSIGNACION = 0 " & _
                             "   And tnyClaveEmpresa = " & CStr(vgintClaveEmpresaContable) & _
                             "   And smiCveDepartamento <> " & CStr(vgintNumeroDepartamento)
              If strAlmaConsolidar <> "" Then
                    strSentencia = strSentencia & strAlmaConsolidar
              End If
16            Set rs = frsRegresaRs(strSentencia, adLockReadOnly, adOpenForwardOnly)
17            pLlenarCboRs_new cboDepartamento, rs, 0, 1

              If IntBitConsolAlmacen = 1 Then
                If IntCvePrincipal = vgintNumeroDepartamento Then
                  '| Agrega el departamento del usuario
18                cboDepartamento.AddItem vgstrNombreDepartamento
19                cboDepartamento.ItemData(cboDepartamento.NewIndex) = vgintNumeroDepartamento
20                cboDepartamento.ListIndex = fintLocalizaCbo_new(cboDepartamento, CStr(vgintNumeroDepartamento))
                End If
              Else
               '| Agrega el departamento del usuario
21              cboDepartamento.AddItem vgstrNombreDepartamento
22              cboDepartamento.ItemData(cboDepartamento.NewIndex) = vgintNumeroDepartamento
23              cboDepartamento.ListIndex = fintLocalizaCbo_new(cboDepartamento, CStr(vgintNumeroDepartamento))
              End If

              '| Si el usuario tiene permiso de control total se habilita el combo para que pueda realizar la
              '| requisición automática de cualquier departamento, sino solo la del departamento del usuario
24            cboDepartamento.Enabled = fblnRevisaPermiso(vglngNumeroLogin, llngNumOpcionRequisitar, "C")
              
        End If

    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargaDeptos" & " Linea:" & Erl()))
End Sub

Private Sub cboDepartamento_Click()
1         On Error GoTo NotificaError
          Dim rsConSigna As ADODB.Recordset
          Dim vgstrParametrosSP As String
          Dim rs As New ADODB.Recordset
          Dim rsIvparametro As ADODB.Recordset
          Dim strAlmaConsolidar As String
          Dim vlArrayAlmacen() As String
          Dim intContAlmaConsol As Integer

2         vgstrParametrosSP = 1 & "|" & vgintClaveEmpresaContable & "|" & CStr(cboDepartamento.ItemData(cboDepartamento.ListIndex))
3         Set rsConSigna = frsEjecuta_SP(vgstrParametrosSP, "Sp_Ivselalmacenconsigna")
          
4         lblnReubicarInsumos = False
5         If cboDepartamento.ListIndex <> -1 Then
              
6             pLimpiar
          
7             cboLocalizacion.Clear
8             vgstrParametrosSP = CStr(cboDepartamento.ItemData(cboDepartamento.ListIndex)) & "|1|0|" & vgintClaveEmpresaContable
9             Set rs = frsEjecuta_SP(vgstrParametrosSP, "SP_IVSELLOCALIZACIONBIT")
              
              
10            If rs.RecordCount > 0 Then
11                chkVarias.Enabled = True
12                pLlenarCboRs_new cboLocalizacion, rs, 0, 1
13                pLlenarListRs lstDisponibles, rs, 0, 1
14            Else
15                chkVarias.Enabled = False
16            End If
              
             
17            cboLocalizacion.AddItem "<TODAS>", 0
18            cboLocalizacion.ItemData(cboLocalizacion.NewIndex) = -1
19            cboLocalizacion.ListIndex = 0
20            lstrDeptosReubican = fstrDeptosReubican("RE")
21            lstrDeptosCompran = fstrDeptosReubican("CO")
              
              'Cargar los tipos de requisiciones que tiene configuradas el departamento:
22            cboTipoRequisicion.Clear
23            If Trim(lstrDeptosCompran) <> "" Then
24                cboTipoRequisicion.AddItem "COMPRA - PEDIDO", 0
25                cboTipoRequisicion.ItemData(cboTipoRequisicion.NewIndex) = cintItemDataCompra
26                cboTipoRequisicion.ListIndex = 0
27            End If
28            If Trim(lstrDeptosReubican) <> "" Then
29                cboTipoRequisicion.AddItem "REUBICACION", 0
30                cboTipoRequisicion.ItemData(cboTipoRequisicion.NewIndex) = cintItemDataReubicacion
31                cboTipoRequisicion.ListIndex = 0
32            End If
                
               lbDeptoConSigna = False
33            If rsConSigna.RecordCount > 0 Then
                  lbDeptoConSigna = True
34                cboTipoRequisicion.AddItem "CONSIGNACION", 0
35                cboTipoRequisicion.ItemData(cboTipoRequisicion.NewIndex) = cintItemDataReubicacion
36                cboTipoRequisicion.ListIndex = 0
37            End If
              
38            lblTipoRequisicion.Visible = lstrModo = "C"
39            cboTipoRequisicion.Visible = lstrModo = "C"
40            chkCompradirecta.Visible = lstrModo = "C"
              
41            If lstrModo = "C" Then
42                Set rs = frsEjecuta_SP(CStr(cboDepartamento.ItemData(cboDepartamento.ListIndex)), "SP_GNSELDEPARTAMENTOS")
43                If rs.RecordCount <> 0 Then lblnReubicarInsumos = rs!chrClasificacion = "A"
44                rs.Close
45            End If
46        End If
47        rsConSigna.Close
          If lstrModo = "C" Then
            cmdAgregar.Enabled = True
          End If

48    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboDepartamento_Click" & " Linea:" & Erl()))
End Sub

Private Function fstrDeptosReubican(strTipoRequisicion As String) As String
On Error GoTo NotificaError

    fstrDeptosReubican = ""
    vgstrParametrosSP = Str(cboDepartamento.ItemData(cboDepartamento.ListIndex)) & "|" & strTipoRequisicion
    
    Set rs = frsEjecuta_SP(vgstrParametrosSP, "SP_IVSELDEPTOSURTEREQUISICION")
    If rs.RecordCount <> 0 Then
        fstrDeptosReubican = "|#" & Trim(Str(0)) & ";" & " "
        Do While Not rs.EOF
            fstrDeptosReubican = fstrDeptosReubican & "|#" & Trim(Str(rs!smicvedepartamento)) & ";" & Trim(rs!VCHDESCRIPCION)
            rs.MoveNext
        Loop
    End If

Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fstrDeptosReubican"))
End Function

Private Sub cboFamilia_Click()
On Error GoTo NotificaError

    cboSubFamilia.Clear

    If cboFamilia.ListIndex <> -1 Then
        If cboFamilia.ItemData(cboFamilia.ListIndex) <> -1 Then pCargaSubfamilias
    End If

    cboSubFamilia.AddItem "<TODAS>", 0
    cboSubFamilia.ItemData(cboSubFamilia.NewIndex) = -1
    cboSubFamilia.ListIndex = 0

    lblSubfamilia.Enabled = cboSubFamilia.ListCount > 1
    cboSubFamilia.Enabled = cboSubFamilia.ListCount > 1

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboFamilia_Click"))
End Sub

Private Sub pCargaSubfamilias()
On Error GoTo NotificaError

    vgstrParametrosSP = CStr(cboFamilia.ItemData(cboFamilia.ListIndex)) & "|" & lintTipoArticulo
    Set rs = frsEjecuta_SP(vgstrParametrosSP, "Sp_Ivselsubfamiliaxfamilia")
    If rs.RecordCount <> 0 Then
        pLlenarCboRs_new cboSubFamilia, rs, 2, 3
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargaSubfamilias"))
End Sub

Private Sub pRefrescarinfoArticulo() ' cada que se elige un nombre comercial del combo nombre comercial se actualiza la informacion de manejos,clave y nombre generico CGR
      ' se agrego en caso 6472
1     On Error GoTo NotificaError
          Dim lngIdArticulo As Long
          Dim intcontador As Integer
          Dim rsNGenericoManejos As New ADODB.Recordset 'para buscar con un SP el nombre generico y los articulos
          Dim vlblnSinManejo As Boolean ' indica que el articulo no tiene manejos por lo tanto se selecciona de la lista "SIN MANEJO"
          
2         lngIdArticulo = -1
3         vlblnSinManejo = True
4         Me.txtClave.Text = "" ' se borra la clave de cualquier articulo
5         Me.lblNombreGenerico.Caption = "" ' se borra el nombre generico
6         lstManejos.Enabled = True ' se habilita el list de manejos
          
7         If cboNombreComercial.ListIndex = 0 Then '<TODOS>
8            lstManejos.Enabled = True ' se habilita el list de manejos
            ' Me.optOpcion(0).Value = True
9            For intcontador = 0 To lstManejos.ListCount - 1 'se colocan activos todos los manejos
10               lstManejos.Selected(intcontador) = True
11           Next intcontador
12        Else 'otro articulo
13            lngIdArticulo = cboNombreComercial.ItemData(cboNombreComercial.ListIndex) 'primero debemos obtener el ID del articulo seleccionado
                       
              'se limpia el list de los manejos
14                For intcontador = 0 To lstManejos.ListCount - 1
15                   lstManejos.Selected(intcontador) = False
16                Next intcontador
              
              'se ejecuta el SP con el ID del articulo
17                Set rsNGenericoManejos = frsEjecuta_SP(CStr(lngIdArticulo), "sp_IvSelNombreGenManejo")
                   
18            With rsNGenericoManejos
19                If Not .EOF Then ' si se encontraron datos se procede con la carga
                         
20                       txtClave.Text = !CHRCVEARTICULO ' se asigna la clave del articulo
21                       lblNombreGenerico.Caption = !VCHDESCRIPCION ' se asigna el nombre generico si es que lo hay
                                           
                        
22                       Do While Not .EOF
23                            For intcontador = 0 To lstManejos.ListCount - 1
24                                If lstManejos.ItemData(intcontador) = !intCveManejo Then
25                                   lstManejos.Selected(intcontador) = True
26                                   vlblnSinManejo = False 'tiene por lo menos un manejo
27                                End If
28                            Next intcontador
29                           .MoveNext
30                       Loop
31                      .Close
32                End If
33           End With
              'se verifica si hubo por lo menos un manejo
34             If vlblnSinManejo Then  ' si no hubo un manejo entonces se marca la casilla de "SIN MANEJO"
35                     lstManejos.Selected(0) = True
36            End If
37             lstManejos.Enabled = False ' se deshabilita el list de manejos
               
38       End If
39       Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ": pRefrescarinfoArticulo" & " Linea:" & Erl()))
End Sub

Private Sub cboNombreComercial_Click()
On Error GoTo NotificaError

        pRefrescarinfoArticulo
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboNombreComercial_Click"))
End Sub

Private Sub pDatosArticulo(strClave As String, lngIdArticulo As Long)
1     On Error GoTo NotificaError
          Dim rsManejos As New ADODB.Recordset
          Dim intcontador As Integer

2         txtClave.Text = "" ' limpia clave
3         lblNombreGenerico.Caption = "" ' limpia nombre generico
          
4         For intcontador = 0 To lstManejos.ListCount - 1 ' activa todos los manejos
5             If lstManejos.ItemData(intcontador) = 0 Then lstManejos.Selected(intcontador) = True
6         Next intcontador

7         vgstrParametrosSP = Str(lngIdArticulo) & "|" & strClave
8         Set rs = frsEjecuta_SP(vgstrParametrosSP, "Sp_Ivseldatosarticulo")
9         If rs.RecordCount <> 0 Then
10            blnNoCargarFiltros = True
              
11            cboNombreComercial.AddItem rs!VCHNOMBRECOMERCIAL, 1
12            cboNombreComercial.ItemData(cboNombreComercial.NewIndex) = rs!intIdArticulo
13            cboNombreComercial.ListIndex = cboNombreComercial.NewIndex
          
14            txtClave.Text = rs!CHRCVEARTICULO
15            lblNombreGenerico.Caption = rs!NombreGenerico
                     
              
16            optOpcion(rs!CHRCVEARTMEDICAMEN + 1).Value = True
                              
17            If rs!CHRCVEARTMEDICAMEN <> 2 Then
              'Manejos
18                For intcontador = 0 To lstManejos.ListCount - 1
19                   lstManejos.Selected(intcontador) = False
20                Next intcontador
                  
21                Set rsManejos = frsEjecuta_SP(rs!intIdArticulo, "Sp_IvSelManejosArticulo")
22                Do While Not rsManejos.EOF
23                    For intcontador = 0 To lstManejos.ListCount - 1
24                        If lstManejos.ItemData(intcontador) = rsManejos!intCveManejo Then
25                            lstManejos.Selected(intcontador) = True
26                        End If
27                    Next intcontador
28                    rsManejos.MoveNext
29                Loop
30                rsManejos.Close
31            End If
32            lstManejos.Enabled = False
33            blnNoCargarFiltros = False
34        End If

35    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pDatosArticulo" & " Linea:" & Erl()))
End Sub

Private Sub cboNombreComercial_KeyDown(KeyCode As Integer, Shift As Integer)
1     On Error GoTo NotificaError
          Dim intcontador As Integer

2         If KeyCode = vbKeyReturn Then
3             If Len(cboNombreComercial.Text) > 0 Then
4                 If Trim(cboNombreComercial.Text) <> "<TODOS>" Then
5                     lblnBusquedaNombre = True
                      
6                     vgstrVarIntercam = UCase(cboNombreComercial.Text)
7                     vgstrVarIntercam2 = "Lista por nombre comercial"
8                     vgstrNombreCbo = "Comercial"
                      
9                     frmLista.gintEstatus = 1
10                    frmLista.gintFamilia = IIf(cboFamilia.ItemData(cboFamilia.ListIndex) = -1, 0, cboFamilia.ItemData(cboFamilia.ListIndex))
11                    frmLista.gintSubfamilia = IIf(cboSubFamilia.ItemData(cboSubFamilia.ListIndex) = -1, 0, cboSubFamilia.ItemData(cboSubFamilia.ListIndex))
12                    frmLista.Tag = ""
13                    If lstrModo = "C" Then
14                        frmLista.Tag = IIf(cboTipoRequisicion.List(cboTipoRequisicion.ListIndex) = "REUBICACION" And lblnReubicarInsumos, "T", Mid(cboTipoRequisicion.List(cboTipoRequisicion.ListIndex), 1, 1))
15                    End If
16                    frmLista.Show vbModal, Me
                      
17                    If Trim(vgstrVarIntercam) <> "" Then
18                        pDatosArticulo Trim(vgstrVarIntercam), -1
                            
                       If lstrModo = "A" Then
                          chkMostrarArti.SetFocus
                       Else
                          cmdAgregar.SetFocus
                       End If

19
20                    End If
21                Else
22                    lblNombreGenerico.Caption = ""
23                    txtClave.Text = ""
24                    lstManejos.Enabled = Not optOpcion(3).Value
25                End If
26            Else
27                lblNombreGenerico.Caption = ""
28                txtClave.Text = ""
29                For intcontador = 0 To lstManejos.ListCount - 1
30                    lstManejos.Selected(intcontador) = True
31                Next intcontador
32                lstManejos.Enabled = Not optOpcion(3).Value
33                cboNombreComercial.ListIndex = 0
34            End If
35        End If

36    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboNombreComercial_KeyDown" & " Linea:" & Erl()))
End Sub

Private Sub cboNombreComercial_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboNombreComercial_KeyPress"))
End Sub

Private Sub cboNombreComercial_LostFocus()
    Dim intcontador As Integer

    If cboNombreComercial.ListIndex = -1 Then
        cboNombreComercial.ListIndex = 0
        For intcontador = 0 To lstManejos.ListCount - 1
            lstManejos.Selected(intcontador) = True
        Next intcontador
        lstManejos.Enabled = Not optOpcion(3).Value
    End If
    
End Sub

Private Sub cboSubFamilia_Click()
    On Error GoTo NotificaError

    If Not lblnBusquedaNombre Then
        cboNombreComercial.Clear
        lblNombreGenerico.Caption = ""
        txtClave.Text = ""
        If cboSubFamilia.ListIndex <> -1 Then
            If cboSubFamilia.ItemData(cboSubFamilia.ListIndex) <> -1 Then
                pCargaArticulos
            End If
        End If
        cboNombreComercial.AddItem "<TODOS>", 0
        cboNombreComercial.ItemData(cboNombreComercial.NewIndex) = -1
        cboNombreComercial.ListIndex = 0
    Else
        lblnBusquedaNombre = False
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboSubfamilia_Click"))
End Sub

Private Sub pCargaArticulos()
On Error GoTo NotificaError

    vgstrParametrosSP = CStr(lintTipoArticulo) & "|" & CStr(cboFamilia.ItemData(cboFamilia.ListIndex)) & "|" & CStr(cboSubFamilia.ItemData(cboSubFamilia.ListIndex))
    Set rs = frsEjecuta_SP(vgstrParametrosSP, "SP_IVSELARTICULOFAMILIASUB")
    If rs.RecordCount <> 0 Then
        pLlenarCboRs_new cboNombreComercial, rs, 0, 2
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargaArticulos"))
End Sub

Private Sub pRecargarArticulos() ' despues de agregar o cambiar de asignacion carga de nuevo todo los articulos al grid se agregó en caso 6472
1     On Error GoTo NotificaError
          Dim rsClavesArticulos As New ADODB.Recordset ' este recordser cargar las claves de los articulos del vsfARticulos, para asi cargarlos de nuevo despues de inicializado
          Dim vlintRenglones As Integer 'utilizada para recorrer el vsfarticulos
          Dim vlblnCboNombreComercialUnaVez As Boolean
2         vlblnCboNombreComercialUnaVez = True
3         vlintRenglones = 0
          'inicializamos el recordset con una sola columna que almacene la clave de los articulos, es lo unico que re requiere
4         With rsClavesArticulos
5         .CursorType = adOpenDynamic
         ' .Fields.Append "ID", adVarChar, 10
6         .Fields.Append "clave", adVarChar, 10
7         .Open
8         For vlintRenglones = 1 To vsfArticulos.Rows - 1
9             If vsfArticulos.TextMatrix(vlintRenglones, 20) <> "" Then
10                .AddNew
11                !clave = vsfArticulos.TextMatrix(vlintRenglones, cintColVsfClaveArt)
12               .Update
13            End If
14        Next
15      .MoveFirst
          
16          pLimpiarVsfArticulos '1
17          pConfVsfArticulos '2
           ' pAgregar 3 -estas tres lineas estaban en el boton de activar/desactiva, nos las traemos para aca y
           'descativamos el pagregar, ya que este se manda llamar en las siguientes lineas
          
          ' ahora se debe recorrer el recordset y mandar llamar a un par de procedimientos
          ' con la clave de los articulos que de deben de cargar de nuevo
18             Do While Not .EOF
19                If vlblnCboNombreComercialUnaVez Then
20                    cboNombreComercial.Clear
21                    cboNombreComercial.AddItem "<TODOS>", 0
22                    cboNombreComercial.ItemData(cboNombreComercial.NewIndex) = -1
23                    vlblnCboNombreComercialUnaVez = False
24                End If
25                pDatosArticulo !clave, -1
26                pAgregar
27                .MoveNext
28            Loop
29            .Close
30       End With
31     vsfArticulos_Click
32    Exit Sub
NotificaError:
       Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pRecargarArticulos" & " Linea:" & Erl()))
End Sub

Private Sub pAgregar()
1     On Error GoTo NotificaError

          Dim rsRequisicionConsigna As ADODB.Recordset
          Dim rsIvparametro As ADODB.Recordset
          Dim rsArticulos As ADODB.Recordset
          Dim rsTemp As New ADODB.Recordset
          Dim rsConSigna As ADODB.Recordset
          Dim vgstrParametrosSPConsigna As String
          Dim vgstrParametrosSP As String
          Dim strManejos As String
          Dim intRefrigerado As Integer
          Dim intControlado As Integer
          Dim intcontador As Integer
          Dim intTipoArt As Integer
          Dim intColumna As Integer
          Dim bitManejo As Integer
          Dim lngTotalFaltante As Long
          Dim lngIdArticulo As Long
          Dim lngTotal As Long
          Dim blnSinManejo As Boolean
          Dim rsArticulosVal As ADODB.Recordset
          Dim strSentencia As String
          
          strClavesConsol = ""

2         vgstrParametrosSP = 1 & "|" & vgintClaveEmpresaContable & "|" & CStr(cboDepartamento.ItemData(cboDepartamento.ListIndex))
3         Set rsConSigna = frsEjecuta_SP(vgstrParametrosSP, "Sp_Ivselalmacenconsigna")
          
4         If optOpcion(0).Value Then
              'Todos
5             intTipoArt = -1
6         ElseIf optOpcion(1).Value Then
              'Artículos
7             intTipoArt = 0
8         ElseIf optOpcion(2).Value Then
              'Medicamentos
9             intTipoArt = 1
10        ElseIf optOpcion(3).Value Then
              'Insumos
11            intTipoArt = 2
12        End If
          
          'Manejos
13        strManejos = "_"
14        blnSinManejo = False
          
15        For intcontador = 0 To lstManejos.ListCount - 1
16            If lstManejos.Selected(intcontador) = True Or cboNombreComercial.ListIndex > 0 Then
17                strManejos = strManejos & lstManejos.ItemData(intcontador) & "_"
18                If lstManejos.ItemData(intcontador) = 0 Then blnSinManejo = True
19            End If
20        Next intcontador
          
21        bitManejo = IIf(strManejos = "_", 3, IIf(blnSinManejo, 1, 2))
          
22        lngIdArticulo = -1
23        If cboNombreComercial.ListIndex <> -1 Then
24            lngIdArticulo = cboNombreComercial.ItemData(cboNombreComercial.ListIndex)
25        End If
          
26        If chkVarias.Value = 0 Then
27            lstrTotalLocalizaciones = IIf(cboLocalizacion.ItemData(cboLocalizacion.ListIndex) = -1, "-1", "_" & cboLocalizacion.ItemData(cboLocalizacion.ListIndex) & "_")
28        Else
29            If lstSeleccionadas.ListCount = 0 Then
30                MsgBox SIHOMsg(3), vbOKOnly + vbCritical, "Mensaje"
31                fraLocalizaciones.Visible = True
32                Exit Sub
33            End If
34        End If
          
35        vgstrParametrosSP = Str(cboDepartamento.ItemData(cboDepartamento.ListIndex)) _
                              & "|" & Str(IIf(lngIdArticulo = -1, intTipoArt, -1)) _
                              & "|" & IIf(lngIdArticulo = -1, lstrTotalLocalizaciones, "-1") _
                              & "|" & Str(IIf(lngIdArticulo = -1, cboFamilia.ItemData(cboFamilia.ListIndex), -1)) _
                              & "|" & Str(IIf(lngIdArticulo = -1, cboSubFamilia.ItemData(cboSubFamilia.ListIndex), -1)) _
                              & "|" & Trim(Str(lngIdArticulo)) _
                              & "|" & bitManejo _
                              & "|" & strManejos
36        If lstrModo = "C" Then
37            lbCveAlmacenPrin = 0
              Set rsIvparametro = frsRegresaRs("select tnyalmacenprincipal as valor,vchcvesalmacenconsolidar as claves,bitconsolidaralmacen from ivparametro where tnyclaveempresa = " & CStr(vgintClaveEmpresaContable), adLockOptimistic, adOpenKeyset)
           
              If IIf(IsNull(rsIvparametro!BITCONSOLIDARALMACEN), 0, rsIvparametro!BITCONSOLIDARALMACEN) = 1 Then
                strClavesConsol = Replace(rsIvparametro!claves, ",", "_")
                lbCveAlmacenPrin = IIf(IsNull(rsIvparametro!Valor), 0, rsIvparametro!Valor)
                
                If IIf(IsNull(rsIvparametro!Valor), 0, rsIvparametro!Valor) = cboDepartamento.ItemData(cboDepartamento.ListIndex) Then
                  vgstrParametrosSP = vgstrParametrosSP & "|" & strClavesConsol
                  Set rs = frsEjecuta_SP(vgstrParametrosSP, "sp_ivselmaxmindeptoconsol")
                Else
                  Set rs = frsEjecuta_SP(vgstrParametrosSP, "sp_ivselmaxmindeptoexclusivo")
                End If
             Else
               Set rs = frsEjecuta_SP(vgstrParametrosSP, "sp_ivselmaxmindeptoexclusivo")
             End If
38        Else
39              If chkMostrarArti.Value = 1 Then
                    vgstrParametrosSP = vgstrParametrosSP & "|0"
40                  Set rs = frsEjecuta_SP(vgstrParametrosSP, "SP_IVSELMAXMINASIGDEPTO")
                Else
                  vgstrParametrosSP = vgstrParametrosSP & "|0"
41                Set rs = frsEjecuta_SP(vgstrParametrosSP, "SP_IVSELMAXMINDEPTO")
               End If
          End If
          
42        If rs.RecordCount = 0 Then
43            MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje" 'No existe información con esos parámetros.
44        Else
45            If lstrModo = "C" Then 'Se limpia el grid, no funciona de modo incluyente:
46                pLimpiarVsfArticulos
47                pConfVsfArticulos
48            Else
                  'Si es modo asignación, se habilita la botonera según el permiso que tenga asignado el usuario
49                cmdMarcar.Enabled = lblnPermisoAsignar
50                cmdInvertirSeleccion.Enabled = lblnPermisoAsignar
51            End If
              
52            With vsfArticulos
53                lngTotalFaltante = 0
54                lngTotal = 0
55                .Redraw = False
56                Do While Not rs.EOF
57                    If lstrModo = "C" And (Mid(cboTipoRequisicion.List(cboTipoRequisicion.ListIndex), 1, 1) <> "R" Or rs!CostoGasto <> "G" Or lblnReubicarInsumos) _
                          Or lstrModo = "A" Then
                         'Si es consulta, se incluye, solo si es parte del stock (solo se toma en cuenta el máximo porque el punto de reorden o el mínimo si pueden ser cero), si es asignación se incluye:
58                        If (lstrModo = "C" And rs!Maximo <> 0) Or lstrModo = "A" Then
59                            If Not fblnExiste(rs!idArticulo) Then 'Esta función nunca va a encontrar cuando es modo consulta, pero que tiene, asi la dejé
60                                If Val(.TextMatrix(.Rows - 1, cintColVsfIdArticulo)) <> 0 Then .Rows = .Rows + 1
                                  
61                                .TextMatrix(.Rows - 1, cintColVsfIdArticulo) = rs!idArticulo
62                                .TextMatrix(.Rows - 1, cintColVsfNombreComercial) = rs!NombreComercial
63                                If lstrModo = "C" Then
64                                    .TextMatrix(.Rows - 1, cintColvsfLocalizacion) = IIf(IsNull(rs!Localizacion), " ", rs!Localizacion)
65                                End If
                                  
66                                If lstrModo = "C" Then
67                                    If cboTipoRequisicion.ItemData(cboTipoRequisicion.ListIndex) = cintItemDataReubicacion Then
68                                        If rsConSigna.RecordCount = 0 Then
                                              'Se va a hacer una requisición tipo reubicación, en unidades mínimas:
69                                            .TextMatrix(.Rows - 1, cintColVsfExistencia) = rs!ExistenciaAlterna * rs!Contenido + rs!ExistenciaMinima
70                                            .TextMatrix(.Rows - 1, cintColVsfExistenciaUnidad) = rs!NombreMinima
71                                        Else
                                              'Se va a hacer una requisición consignación, en unidades alternas completas:
72                                            .TextMatrix(.Rows - 1, cintColVsfExistencia) = CLng((rs!ExistenciaAlterna * rs!Contenido + rs!ExistenciaMinima) / rs!Contenido)
73                                            .TextMatrix(.Rows - 1, cintColVsfExistenciaUnidad) = rs!NombreAlterna
74                                        End If
75                                    Else
                                          'Se va a hacer una requisición tipo compra - pedido, en unidades alternas completas:
76                                        .TextMatrix(.Rows - 1, cintColVsfExistencia) = CLng((rs!ExistenciaAlterna * rs!Contenido + rs!ExistenciaMinima) / rs!Contenido)
77                                        .TextMatrix(.Rows - 1, cintColVsfExistenciaUnidad) = rs!NombreAlterna
78                                    End If
79                                Else
80                                    .TextMatrix(.Rows - 1, cintColVsfExistencia) = ""
81                                    .TextMatrix(.Rows - 1, cintColVsfExistenciaUnidad) = ""
82                                End If
                                  'Máximo:
                                   If strClavesConsol <> "" Then
                                        .TextMatrix(.Rows - 1, cintColVsfCapturaMaximo) = rs!MaximoAlmacenes
                                   Else
                                        .TextMatrix(.Rows - 1, cintColVsfCapturaMaximo) = rs!Maximo
                                   End If
83
84                                .TextMatrix(.Rows - 1, cintColVsfCapturaUnidadMax) = rs!UnidadMaximo
85                                .TextMatrix(.Rows - 1, cintColVsfCapturaCveUnidadMax) = rs!IdUnidadMaximo
86                                .TextMatrix(.Rows - 1, cintColvsfCapturaTipoUniMax) = IIf(IsNull(rs!intUnidadMaximo), " ", rs!intUnidadMaximo)
                                  'Reorden:
87                                .TextMatrix(.Rows - 1, cintColVsfCapturaPunto) = rs!Reorden
88                                .TextMatrix(.Rows - 1, cintColVsfCapturaUnidadPun) = rs!UnidadReorden
89                                .TextMatrix(.Rows - 1, cintColVsfCapturaCveUnidadPun) = rs!IdUnidadReorden
90                                .TextMatrix(.Rows - 1, cintColvsfCapturaTipoUniPun) = IIf(IsNull(rs!INTUNIDADPUNTOREORDEN), " ", rs!INTUNIDADPUNTOREORDEN)
                                  'Mínimo:
91                                .TextMatrix(.Rows - 1, cintColVsfCapturaMinimo) = rs!Minimo
92                                .TextMatrix(.Rows - 1, cintColVsfCapturaUnidadMin) = rs!UnidadMinimo
93                                .TextMatrix(.Rows - 1, cintColVsfCapturaCveUnidadMin) = rs!IdUnidadMinimo
94                                .TextMatrix(.Rows - 1, cintColvsfCapturaTipoUniMin) = IIf(IsNull(rs!intUnidadMinimo), " ", rs!intUnidadMinimo)
                                  
                                  'Almacén:
                                  ' Valida artículos de consignación
95                                If rsConSigna.RecordCount = 0 Then
96                                    .TextMatrix(.Rows - 1, cintColVsfCapturaAlmacen) = rs!ALMACEN
97                                Else
98                                    If rs!clave = rs!cveArticulo Then
99                                        If cboDepartamento.ListCount > 0 Then
100                                           .TextMatrix(.Rows - 1, cintColVsfCapturaAlmacen) = cboDepartamento.List(cboDepartamento.ListIndex)
101                                       End If
102                                   End If
103                               End If
                                  
104                               .TextMatrix(.Rows - 1, cintColVsfCapturaCveAlmacen) = rs!IdAlmacen
                                  
                                  'Departamento compra:
105                               .TextMatrix(.Rows - 1, cintColVsfCapturaDeptoCompra) = rs!DepartamentoCompra
106                               .TextMatrix(.Rows - 1, cintColVsfCapturaCveDeptoCompra) = rs!IdDepartamentoCompra
107                               .TextMatrix(.Rows - 1, cintColVsfClaveArt) = rs!clave
108                               .TextMatrix(.Rows - 1, cintColVsfContenidoArt) = rs!Contenido
109                               .TextMatrix(.Rows - 1, cintColVsfCveUniAlternaArt) = rs!CveAlterna
110                               .TextMatrix(.Rows - 1, cintColvsfUnidadAlternaArt) = rs!NombreAlterna
111                               .TextMatrix(.Rows - 1, cintColVsfCveUniMinimaArt) = rs!CveMinima
112                               .TextMatrix(.Rows - 1, cintColvsfUnidadMinimaArt) = rs!NombreMinima
113                               .TextMatrix(.Rows - 1, cintcolvsfGuardar) = 0
114                               .TextMatrix(.Rows - 1, cintColvsfCveArt) = IIf(IsNull(rs!cveArticulo), "*", rs!cveArticulo) 'Con esta columna se sabe si está ubicado o no el art.
115                               .TextMatrix(.Rows - 1, cintColvsfControlado) = rs!Controlado
116                               .TextMatrix(.Rows - 1, cintColvsfExistenciaAlterna) = rs!ExistenciaAlterna
117                               .TextMatrix(.Rows - 1, cintColvsfExistenciaMinima) = rs!ExistenciaMinima
                                  
                                  If strClavesConsol <> "" Then
                                    .TextMatrix(.Rows - 1, cintColVsfArtiExisEsclavos) = IIf(IsNull(rs!existenciaalternaesclavos), "", rs!existenciaalternaesclavos)
                                    .TextMatrix(.Rows - 1, cintColVsfArtiMaxEsclavos) = IIf(IsNull(rs!maximoesclavos), "", rs!maximoesclavos)
                                    .TextMatrix(.Rows - 1, cintColVsfArtiReordenEsclavos) = IIf(IsNull(rs!reordenesclavos), "", rs!reordenesclavos)
                                    .TextMatrix(.Rows - 1, cintColVsfArtiDeptoEsclavos) = IIf(IsNull(rs!deptoesclavos), "", rs!deptoesclavos)
                                    .TextMatrix(.Rows - 1, cintColVsfArtiMinEsclavos) = IIf(IsNull(rs!MinimaEsclavos), "", rs!MinimaEsclavos)
                                  End If
                                  
118                               If lstrModo = "C" Then 'Modo consulta:
119                                   If rsConSigna.RecordCount = 0 Then ' Valida requisiciones de consignación
120                                        If strClavesConsol <> "" Then
                                               .TextMatrix(.Rows - 1, cintColvsfEstadoArt) = fstrEstadoArtConsol(.Rows - 1)
                                           Else
                                            .TextMatrix(.Rows - 1, cintColvsfEstadoArt) = fstrEstadoArt(.Rows - 1)
                                           End If
121                                        lngTotalFaltante = lngTotalFaltante + IIf(Val(.TextMatrix(.Rows - 1, cintColvsfRequisitar)) = 1, 1, 0)
122                                   Else
123                                       .TextMatrix(.Rows - 1, cintColvsfEstadoArt) = fstrEstadoArt(.Rows - 1)
124                                       lngTotalFaltante = lngTotalFaltante + IIf(Val(.TextMatrix(.Rows - 1, cintColvsfRequisitar)) = 1, 1, 0)
                                          
125                                       vgstrParametrosSPConsigna = vsfArticulos.TextMatrix(vsfArticulos.Row, cintColVsfClaveArt) & "|" & Str(vgintDiasRequisicion) & "|" & Str(cboDepartamento.ItemData(cboDepartamento.ListIndex))
126                                       Set rsRequisicionConsigna = frsEjecuta_SP(vgstrParametrosSPConsigna, "SP_IVSELREQARTPEN_CONSIGNA")
                                                  
127                                       If rsRequisicionConsigna.RecordCount = 0 And .TextMatrix(.Rows - 1, cintColvsfEstadoArt) = "SOLICITADO" Then
128                                           .RemoveItem (.RowSel)
129                                       End If
130                                       rsRequisicionConsigna.Close
131                                   End If 'End if valida requisiciones de consignación
132                               End If
                                  
                                  'Manejos
                                  'Set rsTemp = frsEjecuta_SP(CStr(rs!idArticulo), "Sp_IvSelArticuloManejos")
133                               Set rsTemp = frsEjecuta_SP(CStr(rs!idArticulo), "Sp_IvSelArticuloManejos")
134                               With rsTemp
135                                   intcontador = cintColVsfManejo4
136                                   Do While Not .EOF
137                                       If intcontador >= cintColVsfManejo1 Then
138                                           If Not IsNull(!vchSimbolo) Then
139                                               vsfArticulos.Col = intcontador
140                                               vsfArticulos.Row = vsfArticulos.Rows - 1
141                                               vsfArticulos.CellFontName = "Wingdings"
142                                               vsfArticulos.CellFontSize = 12
143                                               vsfArticulos.CellForeColor = CLng(!vchColor)
144                                               vsfArticulos.TextMatrix(vsfArticulos.Row, intcontador) = !vchSimbolo
145                                               intcontador = intcontador - 1
                                                  
146                                           End If
147                                       End If
148                                       .MoveNext
149                                   Loop
150                                   If .RecordCount > intMaxManejos Then
151                                       intMaxManejos = IIf(.RecordCount > 4, 4, .RecordCount)
152                                   End If
153                               End With
154                               rsTemp.Close
155                               lngTotal = lngTotal + 1
156                           End If
157                       End If
                      
158                   Else
159                   End If
                                  
160                   rs.MoveNext
161               Loop
                  'rsArticulosVal.Close
162               vsfArticulos.ColWidth(cintColVsfManejo1) = IIf(intMaxManejos = 4, 300, 0)
163               vsfArticulos.ColWidth(cintColVsfManejo2) = IIf(intMaxManejos >= 3, 300, 0)
164               vsfArticulos.ColWidth(cintColVsfManejo3) = IIf(intMaxManejos >= 2, 300, 0)
165               vsfArticulos.ColWidth(cintColVsfManejo4) = IIf(intMaxManejos >= 1, 300, 0)
                  
166               If lngTotal = 0 And lstrModo = "C" Then
                      'No existe información con esos parámetros.
167                   MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje"
168               End If
                  
169               .Row = 1
170               .Col = cintColVsfNombreComercial
171               .SetFocus
172               .Redraw = True
               'txtClave.Text = ""
               'lblNombreGenerico.Caption = ""
173           End With
              
174           If lstrModo = "A" Then
175               fraTipoAsignacion.Enabled = True
176               optconsumo(0).Enabled = True
177               optconsumo(1).Enabled = True
178           Else
179               cmdRequisitar.Enabled = lngTotalFaltante > 0
180               If lngTotal > 0 Then cmdVistaPreliminar.Enabled = True
181           End If
182       End If
183       rsConSigna.Close
          
184   Exit Sub
NotificaError:
       Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pAgregar" & " Linea:" & Erl()))
End Sub

Private Sub pAgreImportar(StrIdArticuloAgrega As String)
1     On Error GoTo NotificaError

          Dim rsRequisicionConsigna As ADODB.Recordset
          Dim rsArticulos As ADODB.Recordset
          Dim rsTemp As New ADODB.Recordset
          Dim rsConSigna As ADODB.Recordset
          Dim vgstrParametrosSPConsigna As String
          Dim vgstrParametrosSP As String
          Dim strManejos As String
          Dim intRefrigerado As Integer
          Dim intControlado As Integer
          Dim intcontador As Integer
          Dim intTipoArt As Integer
          Dim intColumna As Integer
          Dim bitManejo As Integer
          Dim lngTotalFaltante As Long
          Dim lngIdArticulo As Long
          Dim lngTotal As Long
          Dim blnSinManejo As Boolean
          Dim rsArticulosVal As ADODB.Recordset
          Dim strSentencia As String
          
          lstrModo = "A"

2         vgstrParametrosSP = 1 & "|" & vgintClaveEmpresaContable & "|" & CStr(cboDepartamento.ItemData(cboDepartamento.ListIndex))
3         Set rsConSigna = frsEjecuta_SP(vgstrParametrosSP, "Sp_Ivselalmacenconsigna")
4
          'Todos
5         intTipoArt = -1

          'Manejos
13        strManejos = "_"
14        blnSinManejo = False
          
15        For intcontador = 0 To lstManejos.ListCount - 1
16            If lstManejos.Selected(intcontador) = True Or cboNombreComercial.ListIndex > 0 Then
17                strManejos = strManejos & lstManejos.ItemData(intcontador) & "_"
18                If lstManejos.ItemData(intcontador) = 0 Then blnSinManejo = True
19            End If
20        Next intcontador
          
21        bitManejo = IIf(strManejos = "_", 3, IIf(blnSinManejo, 1, 2))
          
22        lngIdArticulo = -1
23
24        lngIdArticulo = StrIdArticuloAgrega
25
27        lstrTotalLocalizaciones = "-1"
          
35        vgstrParametrosSP = Str(cboDepartamento.ItemData(cboDepartamento.ListIndex)) _
                              & "|" & Str(-1) _
                              & "|" & ("-1") _
                              & "|" & Str(-1) _
                              & "|" & Str(-1) _
                              & "|" & Str(lngIdArticulo) _
                              & "|" & bitManejo _
                              & "|" & strManejos
                              
39           vgstrParametrosSP = vgstrParametrosSP & "|0"
40           Set rs = frsEjecuta_SP(vgstrParametrosSP, "SP_IVSELMAXMINDEPTO")
41
          lblnProductoExiste = True
42        If rs.RecordCount = 0 Then
                lblnProductoExiste = False
43           ' MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje" 'No existe información con esos parámetros.
44        Else
45
46
                  'Si es modo asignación, se habilita la botonera según el permiso que tenga asignado el usuario
49                cmdMarcar.Enabled = lblnPermisoAsignar
50                cmdInvertirSeleccion.Enabled = lblnPermisoAsignar
51
              
52            With vsfArticulos
53                lngTotalFaltante = 0
54                lngTotal = 0
55                .Redraw = False
56                Do While Not rs.EOF
57
58                        If lstrModo = "A" Then
59
60                                If Val(.TextMatrix(.Rows - 1, cintColVsfIdArticulo)) <> 0 Then .Rows = .Rows + 1
                                  
61                                .TextMatrix(.Rows - 1, cintColVsfIdArticulo) = rs!idArticulo
62                                .TextMatrix(.Rows - 1, cintColVsfNombreComercial) = rs!NombreComercial
80                                .TextMatrix(.Rows - 1, cintColVsfExistencia) = ""
81                                .TextMatrix(.Rows - 1, cintColVsfExistenciaUnidad) = ""
82
                                  'Máximo:
83                                '.TextMatrix(.Rows - 1, cintColVsfCapturaMaximo) = ""
84                                .TextMatrix(.Rows - 1, cintColVsfCapturaUnidadMax) = rs!UnidadMaximo
85                                .TextMatrix(.Rows - 1, cintColVsfCapturaCveUnidadMax) = rs!IdUnidadMaximo
86                                .TextMatrix(.Rows - 1, cintColvsfCapturaTipoUniMax) = IIf(IsNull(rs!intUnidadMaximo), " ", rs!intUnidadMaximo)
                                  'Reorden:
87                                '.TextMatrix(.Rows - 1, cintColVsfCapturaPunto) = rs!Reorden
88                                .TextMatrix(.Rows - 1, cintColVsfCapturaUnidadPun) = rs!UnidadReorden
89                                .TextMatrix(.Rows - 1, cintColVsfCapturaCveUnidadPun) = rs!IdUnidadReorden
90                                .TextMatrix(.Rows - 1, cintColvsfCapturaTipoUniPun) = IIf(IsNull(rs!INTUNIDADPUNTOREORDEN), " ", rs!INTUNIDADPUNTOREORDEN)
                                  'Mínimo:
91                                '.TextMatrix(.Rows - 1, cintColVsfCapturaMinimo) = rs!Minimo
92                                .TextMatrix(.Rows - 1, cintColVsfCapturaUnidadMin) = rs!UnidadMinimo
93                                .TextMatrix(.Rows - 1, cintColVsfCapturaCveUnidadMin) = rs!IdUnidadMinimo
94                                .TextMatrix(.Rows - 1, cintColvsfCapturaTipoUniMin) = IIf(IsNull(rs!intUnidadMinimo), " ", rs!intUnidadMinimo)
                                  
                                  'Almacén:
                                  ' Valida artículos de consignación
95                                If rsConSigna.RecordCount = 0 Then
96                                    .TextMatrix(.Rows - 1, cintColVsfCapturaAlmacen) = rs!ALMACEN
97                                Else
98                                    If rs!clave = rs!cveArticulo Then
99                                        If cboDepartamento.ListCount > 0 Then
100                                           .TextMatrix(.Rows - 1, cintColVsfCapturaAlmacen) = cboDepartamento.List(cboDepartamento.ListIndex)
101                                       End If
102                                   End If
103                               End If
                                  
104                               .TextMatrix(.Rows - 1, cintColVsfCapturaCveAlmacen) = rs!IdAlmacen
                                  
                                  'Departamento compra:
105                               .TextMatrix(.Rows - 1, cintColVsfCapturaDeptoCompra) = rs!DepartamentoCompra
106                               .TextMatrix(.Rows - 1, cintColVsfCapturaCveDeptoCompra) = rs!IdDepartamentoCompra
107                               .TextMatrix(.Rows - 1, cintColVsfClaveArt) = rs!clave
108                               .TextMatrix(.Rows - 1, cintColVsfContenidoArt) = rs!Contenido
109                               .TextMatrix(.Rows - 1, cintColVsfCveUniAlternaArt) = rs!CveAlterna
110                               .TextMatrix(.Rows - 1, cintColvsfUnidadAlternaArt) = rs!NombreAlterna
111                               .TextMatrix(.Rows - 1, cintColVsfCveUniMinimaArt) = rs!CveMinima
112                               .TextMatrix(.Rows - 1, cintColvsfUnidadMinimaArt) = rs!NombreMinima
113                               .TextMatrix(.Rows - 1, cintcolvsfGuardar) = 0
114                               .TextMatrix(.Rows - 1, cintColvsfCveArt) = IIf(IsNull(rs!cveArticulo), "*", rs!cveArticulo) 'Con esta columna se sabe si está ubicado o no el art.
115                               .TextMatrix(.Rows - 1, cintColvsfControlado) = rs!Controlado
116                               .TextMatrix(.Rows - 1, cintColvsfExistenciaAlterna) = rs!ExistenciaAlterna
117                               .TextMatrix(.Rows - 1, cintColvsfExistenciaMinima) = rs!ExistenciaMinima
                                 
                                  'Manejos
133                               Set rsTemp = frsEjecuta_SP(CStr(rs!idArticulo), "Sp_IvSelArticuloManejos")
134                               With rsTemp
135                                   intcontador = cintColVsfManejo4
136                                   Do While Not .EOF
137                                       If intcontador >= cintColVsfManejo1 Then
138                                           If Not IsNull(!vchSimbolo) Then
139                                               vsfArticulos.Col = intcontador
140                                               vsfArticulos.Row = vsfArticulos.Rows - 1
141                                               vsfArticulos.CellFontName = "Wingdings"
142                                               vsfArticulos.CellFontSize = 12
143                                               vsfArticulos.CellForeColor = CLng(!vchColor)
144                                               vsfArticulos.TextMatrix(vsfArticulos.Row, intcontador) = !vchSimbolo
145                                               intcontador = intcontador - 1
                                                  
146                                           End If
147                                       End If
148                                       .MoveNext
149                                   Loop
150                                   If .RecordCount > intMaxManejos Then
151                                       intMaxManejos = IIf(.RecordCount > 4, 4, .RecordCount)
152                                   End If
153                               End With
154                               rsTemp.Close
155                               lngTotal = lngTotal + 1
156                           End If
157
160                   rs.MoveNext
161               Loop
                  'rsArticulosVal.Close
162               vsfArticulos.ColWidth(cintColVsfManejo1) = IIf(intMaxManejos = 4, 300, 0)
163               vsfArticulos.ColWidth(cintColVsfManejo2) = IIf(intMaxManejos >= 3, 300, 0)
164               vsfArticulos.ColWidth(cintColVsfManejo3) = IIf(intMaxManejos >= 2, 300, 0)
165               vsfArticulos.ColWidth(cintColVsfManejo4) = IIf(intMaxManejos >= 1, 300, 0)
                  
166               If lngTotal = 0 And lstrModo = "C" Then
                      'No existe información con esos parámetros.
167                   MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje"
168               End If
                  
169               .Row = 1
170               .Col = cintColVsfNombreComercial
171               .SetFocus
172               '.Redraw = True
               'txtClave.Text = ""
               'lblNombreGenerico.Caption = ""
173           End With
              
174           If lstrModo = "A" Then
175               fraTipoAsignacion.Enabled = True
176               optconsumo(0).Enabled = True
177               optconsumo(1).Enabled = True
181           End If
182       End If
183       rsConSigna.Close
          
184   Exit Sub
NotificaError:
       Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pAgregar" & " Linea:" & Erl()))
End Sub
 

Private Sub cboTipoRequisicion_Click()
On Error GoTo NotificaError

    If cboTipoRequisicion.ListIndex <> -1 Then
        pLimpiarVsfArticulos
        pConfVsfArticulos
        chkCompradirecta.Enabled = IIf(Trim(cboTipoRequisicion.List(cboTipoRequisicion.ListIndex)) = "COMPRA - PEDIDO" And blnPermisoCompraDirecta, True, False)
        chkCompradirecta.Value = IIf(Trim(cboTipoRequisicion.List(cboTipoRequisicion.ListIndex)) <> "COMPRA - PEDIDO", 0, chkCompradirecta.Value)
        lstrTipoRequisicion = IIf(cboTipoRequisicion.ItemData(cboTipoRequisicion.ListIndex) = cintItemDataReubicacion, "R", "C")
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdAgregar_Click"))
End Sub

Private Sub cboUnidad_Click()
    pConvierteUnidad
    VsfBusArticulos.Col = cintColBusCantidad
End Sub

Private Sub chkVarias_Click()
    If chkVarias.Value = 1 Then
        If lstrModo = "C" Then
            fraLocalizaciones.Visible = True
            If lstDisponibles.ListCount > 0 Then lstDisponibles.ListIndex = 0
            fraFiltros.Enabled = False
            cboLocalizacion.Enabled = False
            cboLocalizacion.ListIndex = 0
        End If
    Else
        cboLocalizacion.Enabled = True
    End If
End Sub

Private Sub cmdAceptar_Click()
    lstrTotalLocalizaciones = ""
    
    If lstSeleccionadas.ListCount > 0 Then
        lstrTotalLocalizaciones = "_"
        For lintCiclos = 0 To lstSeleccionadas.ListCount - 1
            lstrTotalLocalizaciones = lstrTotalLocalizaciones & lstSeleccionadas.ItemData(lintCiclos) & "_"
        Next lintCiclos
    End If
    fraFiltros.Enabled = True
    fraLocalizaciones.Visible = False
    
    If fblnCanFocus(lstManejos) Then
        lstManejos.SetFocus
    Else
        cmdAgregar.SetFocus
    End If
    
End Sub

Private Sub cmdAgregar_Click()
    On Error GoTo NotificaError
    Dim blnSinManejo As Boolean
    Dim intcontador As Integer

    If cboTipoRequisicion.ListCount = 0 And lstrModo = "C" Then
        'No se ha registrado la configuración del usuario para las requisiciones.
        MsgBox SIHOMsg(913), vbExclamation + vbOKOnly, "Mensaje"
        Exit Sub
    End If
    
    If Not optOpcion(3).Value And cboNombreComercial.ItemData(cboNombreComercial.ListIndex) = -1 Then
        blnSinManejo = False
        For intcontador = 0 To lstManejos.ListCount - 1
            If lstManejos.Selected(intcontador) Then
                blnSinManejo = True
                Exit For
            End If
        Next intcontador
       
        If Not blnSinManejo Then
            'Seleccione el manejo
            MsgBox SIHOMsg(1109), vbOKOnly + vbInformation, "Mensaje"
            lstManejos.SetFocus
            Exit Sub
        End If
    End If
    
    pAgregar
    cmdExportar(2).Enabled = True
    cmdImportar(0).Enabled = False
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdAgregar_Click"))
End Sub

Private Function fstrEstadoArt(lngRenglon As Long) As String
1         On Error GoTo NotificaError
          Dim lngMaximo As Long
          Dim lngReorden As Long
          Dim lngMinimo As Long
          Dim lngExistencia As Long
          Dim lngColor As Long
          Dim intColumna As Integer
          Dim lngPendientes As Long
          Dim lngContenido As Long
          Dim rsConSigna As ADODB.Recordset
          Dim vgstrParametrosSP As String

2         vgstrParametrosSP = 1 & "|" & vgintClaveEmpresaContable & "|" & CStr(cboDepartamento.ItemData(cboDepartamento.ListIndex))
3         Set rsConSigna = frsEjecuta_SP(vgstrParametrosSP, "Sp_Ivselalmacenconsigna")
4         fstrEstadoArt = ""

5         With vsfArticulos
              'Convertir todo a la unidad mínima, si el artículo solo se maneja en alterna, es la misma:
6             lngContenido = Val(.TextMatrix(lngRenglon, cintColVsfContenidoArt))
7             lngMaximo = Val(.TextMatrix(lngRenglon, cintColVsfCapturaMaximo)) * IIf(Val(.TextMatrix(lngRenglon, cintColvsfCapturaTipoUniMax)) = 1, lngContenido, 1)
8             lngReorden = Val(.TextMatrix(lngRenglon, cintColVsfCapturaPunto)) * IIf(Val(.TextMatrix(lngRenglon, cintColvsfCapturaTipoUniPun)) = 1, lngContenido, 1)
9             lngMinimo = Val(.TextMatrix(lngRenglon, cintColVsfCapturaMinimo)) * IIf(Val(.TextMatrix(lngRenglon, cintColvsfCapturaTipoUniMin)) = 1, lngContenido, 1)
10            lngExistencia = Val(.TextMatrix(lngRenglon, cintColvsfExistenciaAlterna)) * lngContenido + Val(.TextMatrix(lngRenglon, cintColvsfExistenciaMinima))
11            lngPendientes = flngCantidadPendiente(.TextMatrix(lngRenglon, cintColVsfClaveArt), cboDepartamento.ItemData(cboDepartamento.ListIndex))

12            If lngExistencia > lngMaximo Then
13                fstrEstadoArt = cstrExcedida
14                lngColor = clngColorVerdeBrillante
15            ElseIf lngExistencia > lngReorden Then
16                fstrEstadoArt = cstrOptima
17                lngColor = clngColorVerde
18            ElseIf _
                  (lngPendientes >= lngMaximo - lngExistencia) _
                  Or (lngPendientes <> 0 And lngPendientes < lngMaximo - lngExistencia And lstrTipoRequisicion = "C" And (lngMaximo - (lngPendientes + lngExistencia)) / lngContenido < 1) Then
19                fstrEstadoArt = cstrSolicitado
20                lngColor = clngColorAzul
21            ElseIf lngExistencia = lngReorden Then
22                fstrEstadoArt = cstrReorden
23                lngColor = clngColorRojo
24                .TextMatrix(lngRenglon, cintColvsfRequisitar) = "1"  ''Identificarlo para la requisición
25            ElseIf lngExistencia >= lngMinimo Then
26                fstrEstadoArt = cstrFaltante
27                lngColor = clngColorRojo
28                .TextMatrix(lngRenglon, cintColvsfRequisitar) = "1" ''Identificarlo para la requisición
29            ElseIf lngExistencia < lngMinimo Then
30                fstrEstadoArt = cstrInsuficiente
31                lngColor = clngColorRojoOscuro
32                .TextMatrix(lngRenglon, cintColvsfRequisitar) = "1" ''Identificarlo para la requisición
33            End If

34            .Row = lngRenglon
              
35            For intColumna = 1 To .Cols - 1
36                .Col = intColumna
37                .CellForeColor = lngColor
38            Next intColumna
39        End With

40    Exit Function
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fstrEstadoArt" & " Linea:" & Erl()))
End Function
Private Function fstrEstadoArtConsol(lngRenglon As Long) As String
1         On Error GoTo NotificaError
          Dim lngMaximo As Long
          Dim lngReorden As Long
          Dim lngMinimo As Long
          Dim lngExistencia As Long
          Dim lngColor As Long
          Dim intColumna As Integer
          Dim lngPendientes As Long
          Dim lngContenido As Long
          Dim rsConSigna As ADODB.Recordset
          Dim vgstrParametrosSP As String
          
          
          Dim intCount As Integer
          Dim intReordenDeptoPrin As Integer
          Dim intExisDeptoPrin As Integer
          Dim intMaxDeptoPrin As Integer
          Dim intMinDeptoPrin As Integer
          Dim intTotDeptosConsol As Integer
          Dim intExisDeptoEsclavo As Integer
          Dim intMaxDeptoEsclavo As Integer
          Dim intMiniDeptoEsclave As Integer
          Dim intReordenDeptoEsclave As Integer
          Dim DeptosConsol() As String
          Dim ExisEsclavoConsol() As String
          Dim MaximEsclavoConsol() As String
          Dim ReordenEsclavoConsol() As String
          Dim MinimaEsclavoConsol() As String
          
        
          Dim vlBandRequisicion As Boolean
          
2         fstrEstadoArtConsol = ""
          vlBandRequisicion = False
        
3       If strClavesConsol <> "" Then
          
4         With vsfArticulos
            intCount = 0
            lngContenido = Val(.TextMatrix(lngRenglon, cintColVsfContenidoArt))
            lngPendientes = flngCantidadPendiente(.TextMatrix(lngRenglon, cintColVsfClaveArt), cboDepartamento.ItemData(cboDepartamento.ListIndex))
            
            ' Almacen Principal
            intExisDeptoPrin = 0
            intMaxDeptoPrin = 0
            intMinDeptoPrin = 0
            
            'Total de Departamentos
            intTotDeptosConsol = 0
            DeptosConsol = Split(.TextMatrix(lngRenglon, cintColVsfArtiDeptoEsclavos), ",")
            intTotDeptosConsol = UBound(DeptosConsol) + 1
            
         
            'Almacen Esclavo Arreglo (n)
            intExisDeptoEsclavo = 0
            intMaxDeptoEsclavo = 0
            intMiniDeptoEsclave = 0
            intReordenDeptoEsclave = 0
            
            ExisEsclavoConsol = Split(.TextMatrix(lngRenglon, cintColVsfArtiExisEsclavos), ",")
            MaximEsclavoConsol = Split(.TextMatrix(lngRenglon, cintColVsfArtiMaxEsclavos), ",")
            ReordenEsclavoConsol = Split(.TextMatrix(lngRenglon, cintColVsfArtiReordenEsclavos), ",")
            MinimaEsclavoConsol = Split(.TextMatrix(lngRenglon, cintColVsfArtiMinEsclavos), ",")
            
            If UBound(ExisEsclavoConsol) < intTotDeptosConsol Then
             ReDim Preserve ExisEsclavoConsol(intTotDeptosConsol)
            End If
            
            If UBound(MaximEsclavoConsol) < intTotDeptosConsol Then
             ReDim Preserve MaximEsclavoConsol(intTotDeptosConsol)
            End If
            
            If UBound(ReordenEsclavoConsol) < intTotDeptosConsol Then
             ReDim Preserve ReordenEsclavoConsol(intTotDeptosConsol)
            End If
            
            If UBound(MinimaEsclavoConsol) < intTotDeptosConsol Then
             ReDim Preserve MinimaEsclavoConsol(intTotDeptosConsol)
            End If
           
            For intCount = 0 To UBound(ExisEsclavoConsol)
            intExisDeptoEsclavo = intExisDeptoEsclavo + IIf(ExisEsclavoConsol(intCount) = "", 0, ExisEsclavoConsol(intCount))
            Next intCount
            
            For intCount = 0 To UBound(MaximEsclavoConsol)
            intMaxDeptoEsclavo = intMaxDeptoEsclavo + IIf(MaximEsclavoConsol(intCount) = "", 0, MaximEsclavoConsol(intCount))
            Next intCount
            
            For intCount = 0 To UBound(ReordenEsclavoConsol)
            intReordenDeptoEsclave = intReordenDeptoEsclave + IIf(ReordenEsclavoConsol(intCount) = "", 0, ReordenEsclavoConsol(intCount))
            Next intCount
            
            For intCount = 0 To UBound(MinimaEsclavoConsol)
            intMiniDeptoEsclave = intMiniDeptoEsclave + IIf(MinimaEsclavoConsol(intCount) = "", 0, MinimaEsclavoConsol(intCount))
            Next intCount
                            
            'Almacen Principal
            intExisDeptoPrin = (.TextMatrix(lngRenglon, cintColVsfExistencia))
            intMaxDeptoPrin = (.TextMatrix(lngRenglon, cintColVsfCapturaMaximo))
            intReordenDeptoPrin = (.TextMatrix(lngRenglon, cintColVsfCapturaPunto))
            intMinDeptoPrin = (.TextMatrix(lngRenglon, cintColVsfCapturaMinimo))
            
            ExisEsclavoConsol(intTotDeptosConsol) = CStr(intExisDeptoPrin)
            MaximEsclavoConsol(intTotDeptosConsol) = CStr(intMaxDeptoPrin)
            ReordenEsclavoConsol(intTotDeptosConsol) = CStr(intReordenDeptoPrin)
            
            'ExisEsclavoConsol = Split(CStr(intExisDeptoPrin) & "," & .TextMatrix(lngRenglon, cintColVsfArtiExisEsclavos), ",")
            'MaximEsclavoConsol = Split(CStr(intMaxDeptoPrin) & "," & .TextMatrix(lngRenglon, cintColVsfArtiMaxEsclavos), ",")
            'ReordenEsclavoConsol = Split(CStr(intReordenDeptoPrin) & "," & .TextMatrix(lngRenglon, cintColVsfArtiReordenEsclavos), ",")
                                      
            lngExistencia = Val(intExisDeptoPrin * IIf(Val(.TextMatrix(lngRenglon, cintColvsfCapturaTipoUniMax)) = 1, lngContenido, 1))
            lngMaximo = Val(intMaxDeptoPrin * IIf(Val(.TextMatrix(lngRenglon, cintColvsfCapturaTipoUniMax)) = 1, lngContenido, 1))
            lngReorden = Val(intReordenDeptoPrin * IIf(Val(.TextMatrix(lngRenglon, cintColvsfCapturaTipoUniMax)) = 1, lngContenido, 1))
            lngMinimo = Val(intMinDeptoPrin) * IIf(Val(intMinDeptoPrin) = 1, lngContenido, 1)
            
            If lngExistencia > lngMaximo Then
                fstrEstadoArtConsol = cstrExcedida
                lngColor = clngColorVerdeBrillante
            ElseIf lngExistencia > lngReorden Then
                fstrEstadoArtConsol = cstrOptima
                lngColor = clngColorVerde
            ElseIf _
                (lngPendientes >= lngMaximo - lngExistencia) _
                Or (lngPendientes <> 0 And lngPendientes < lngMaximo - lngExistencia And lstrTipoRequisicion = "C" And (lngMaximo - (lngPendientes + lngExistencia)) / lngContenido < 1) Then
                fstrEstadoArtConsol = cstrSolicitado
                lngColor = clngColorAzul
            ElseIf lngExistencia = lngReorden Then
                fstrEstadoArtConsol = cstrReorden
                lngColor = clngColorRojo
                .TextMatrix(lngRenglon, cintColvsfRequisitar) = "1"  ''Identificarlo para la requisición
                vlBandRequisicion = True
            ElseIf lngExistencia >= lngMinimo Then
                fstrEstadoArtConsol = cstrFaltante
                lngColor = clngColorRojo
                .TextMatrix(lngRenglon, cintColvsfRequisitar) = "1" ''Identificarlo para la requisición
                vlBandRequisicion = True
            ElseIf lngExistencia < lngMinimo Then
                fstrEstadoArtConsol = cstrInsuficiente
                lngColor = clngColorRojoOscuro
                .TextMatrix(lngRenglon, cintColvsfRequisitar) = "1" ''Identificarlo para la requisición
                vlBandRequisicion = True
            End If
            
            .Row = lngRenglon
              
            For intColumna = 1 To .Cols - 1
                .Col = intColumna
                .CellForeColor = lngColor
            Next intColumna
            
            If vlBandRequisicion = True Then
               Exit Function
            End If
         
          End With
        End If
6    Exit Function
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fstrEstadoArt" & " Linea:" & Erl()))
End Function
Private Function fblnExiste(lngIdArticulo As Long) As Boolean
    On Error GoTo NotificaError
    Dim lngContador As Long

    fblnExiste = False
    lngContador = 1
    Do While Not fblnExiste And lngContador <= vsfArticulos.Rows - 1
        fblnExiste = Val(vsfArticulos.TextMatrix(lngContador, cintColVsfIdArticulo)) = lngIdArticulo
        lngContador = lngContador + 1
    Loop

Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fblnExiste"))
End Function

Private Sub cmdAgregarTodas_Click()
    For lintCiclos = 0 To lstDisponibles.ListCount - 1
        lstSeleccionadas.AddItem lstDisponibles.List(lintCiclos)
        lstSeleccionadas.ItemData(lstSeleccionadas.NewIndex) = lstDisponibles.ItemData(lintCiclos)
    Next lintCiclos
    
    lstSeleccionadas.ListIndex = 0
    lstDisponibles.Clear
End Sub

Private Sub cmdAsignarDesactivar_Click(Index As Integer)
1     On Error GoTo NotificaError
      Dim lngContador As Long
      Dim llngDesasignar As Long
      Dim lstrFaltantes As String

2         If AsignacionAutomatica = False Then ' para que no ocurra este evento cuando se llama este procedimiento desde el boton guardar
3             llngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
4         Else
5             llngPersonaGraba = 1
6         End If
          
7         If llngPersonaGraba <> 0 Then
8             With vsfArticulos
9                 ldblAumento = 100 / (.Rows - 1)
10                pgbBarra.Value = 0
11                fraBarra.Visible = True
12                lblTituloBarra.Caption = IIf(Index = 0, "Asignando artículos al departamento", "Desactivando artículos en el departamento") & ", espere un momento"
                  
13                lstrFaltantes = ""
14                For lngContador = 1 To .Rows - 1
                  
                      '*- Barra de avance
15                    fraBarra.Refresh
16                    If pgbBarra.Value + ldblAumento > 100 Then
17                        pgbBarra.Value = 100
18                    Else
19                        pgbBarra.Value = pgbBarra.Value + ldblAumento
20                    End If
                      '*- Barra de avance
                  
21                    If Trim(.TextMatrix(lngContador, 0)) = "*" Then
22                        .TextMatrix(lngContador, 0) = ""
23                        If Index = 0 Then
                              'Si no esta asignado
24                            If Trim(.TextMatrix(lngContador, cintColvsfCveArt)) = "*" Then
25                                .TextMatrix(lngContador, cintColvsfCveArt) = .TextMatrix(lngContador, cintColVsfClaveArt)
26                                vgstrParametrosSP = _
                                                      .TextMatrix(lngContador, cintColVsfClaveArt) _
                                              & "|" & cboDepartamento.ItemData(cboDepartamento.ListIndex) _
                                              & "|" & "0" & "|" & "0" _
                                              & "|" & "0" & "|" & "0"
27                                frsEjecuta_SP vgstrParametrosSP, "SP_IVINSUBICACION"
                                  
28                                pGuardarLogTransaccion Me.Name, EnmGrabar, llngPersonaGraba, "ASIGNAR A ALMACÉN", Trim(Str(cboDepartamento.ItemData(cboDepartamento.ListIndex)) & " - " & .TextMatrix(lngContador, cintColVsfClaveArt))
29                            Else
                                  'Articulos que ya estaban asignados al depto
30                                lstrFaltantes = lstrFaltantes & Chr(13) & Trim(.TextMatrix(lngContador, cintColVsfNombreComercial))
31                            End If
32                        Else
                              'si esta asignado al depto
33                            If Trim(.TextMatrix(lngContador, cintColvsfCveArt)) <> "*" Then
34                                vgstrParametrosSP = cboDepartamento.ItemData(cboDepartamento.ListIndex) & "|" & _
                                                      .TextMatrix(lngContador, cintColVsfClaveArt)
35                                llngDesasignar = 1
36                                frsEjecuta_SP vgstrParametrosSP, "FN_IvDelIvUbicacion", True, llngDesasignar
37                                If llngDesasignar = -1 Then
                                      'Articulos que tienen existencias
38                                    lstrFaltantes = lstrFaltantes & Chr(13) & Trim(.TextMatrix(lngContador, cintColVsfNombreComercial))
39                                Else
40                                    pGuardarLogTransaccion Me.Name, EnmGrabar, llngPersonaGraba, "DESASIGNAR DEL ALMACÉN", Trim(Str(cboDepartamento.ItemData(cboDepartamento.ListIndex)) & " - " & .TextMatrix(lngContador, cintColVsfClaveArt))
41                                End If
42                            Else
                                  'Articulos que no estaban asignados al depto
43                                lstrFaltantes = lstrFaltantes & Chr(13) & Trim(.TextMatrix(lngContador, cintColVsfNombreComercial))
44                            End If
45                        End If
46                    End If
47                Next lngContador
                  
48                fraBarra.Visible = False
49            End With
              
50            If AsignacionAutomatica = False Then '= true es cuando se mando llamar este procedimiento desde el boton guardar
51                If lstrFaltantes <> "" Then
52                    If Index = 0 Then
                      'Los siguientes artículos ya pertenecían al departamento
53                        MsgBox SIHOMsg(994) & Chr(13) & lstrFaltantes, vbOKOnly + vbInformation, "Mensaje"
54                    Else
                      'Los siguientes artículos no pudieron ser desasignados del departamento ya que tienen existencia o no pertenecen al departamento
55                       MsgBox SIHOMsg(995) & Chr(13) & lstrFaltantes, vbOKOnly + vbInformation, "Mensaje"
56                     End If
57                Else
                      'La operación se realizó satisfactoriamente.
58                    MsgBox SIHOMsg(420), vbOKOnly + vbInformation, "Mensaje"
59                End If

60          End If
61          cmdAsignarDesactivar(0).Enabled = False
62          cmdAsignarDesactivar(1).Enabled = False
63          If fblnCanFocus(cboDepartamento) Then cboDepartamento.SetFocus Else optOpcion(0).SetFocus
                   
64          pRecargarArticulos ' procedimiento que carga de nuevo los articulos al vsfarticulos

       

65        End If

66    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdAsignarDesactivar_Click" & " Linea:" & Erl()))
End Sub

Private Sub cmdCalcular_Click()
1     On Error GoTo NotificaError
      Dim lngContador As Long
      Dim strFechaIni As String
      Dim strFechaFin As String
      Dim lngTotalSinConsumo As Long

2         If fblnFechasValidas() Then

3             strFechaIni = fstrFechaSQL(mskInicio.Text)
4             strFechaFin = fstrFechaSQL(mskFin.Text)
          
           
              
          
          
5             With vsfArticulos
              
6                 ldblAumento = 100 / (vsfArticulos.Rows - 1)
              
7                 pgbBarra.Value = 0
8                 fraBarra.Visible = True
9                 lblTituloBarra.Caption = "Calculando el consumo de los artículos, espere un momento"
          
10                lngTotalSinConsumo = 0
11                For lngContador = 1 To vsfArticulos.Rows - 1
                  
12                    .TextMatrix(lngContador, cintcolvsfGuardar) = "1"
                  
                      '*- Barra de avance
13                    fraBarra.Refresh
14                    If pgbBarra.Value + ldblAumento > 100 Then
15                        pgbBarra.Value = 100
16                    Else
17                        pgbBarra.Value = pgbBarra.Value + ldblAumento
18                    End If
                      '*- Barra de avance
                  
19                    vgstrParametrosSP = _
                      vsfArticulos.TextMatrix(lngContador, cintColVsfIdArticulo) _
                      & "|" & Str(cboDepartamento.ItemData(cboDepartamento.ListIndex)) _
                      & "|" & strFechaIni _
                      & "|" & strFechaFin
                      
20                    Set rs = frsEjecuta_SP(vgstrParametrosSP, "SP_IVSELCALCULACONSUMO")
                      
21                    If rs!cantidad <> 0 Then
22                        If Val(.TextMatrix(lngContador, cintColVsfContenidoArt)) = 1 Then
                              'El articulo se maneja en una sola unidad, todo en unidad alterna
                              
                              'Máximo = consumo
23                            .TextMatrix(lngContador, cintColVsfCapturaMaximo) = rs!cantidad
24                            .TextMatrix(lngContador, cintColVsfCapturaUnidadMax) = .TextMatrix(lngContador, cintColvsfUnidadAlternaArt)
25                            .TextMatrix(lngContador, cintColVsfCapturaCveUnidadMax) = .TextMatrix(lngContador, cintColVsfCveUniAlternaArt)
                              
                              'Mínimo = consumo / 2
26                            .TextMatrix(lngContador, cintColVsfCapturaMinimo) = Round(rs!cantidad / 2, 0)
27                            .TextMatrix(lngContador, cintColVsfCapturaUnidadMin) = .TextMatrix(lngContador, cintColvsfUnidadAlternaArt)
28                            .TextMatrix(lngContador, cintColVsfCapturaCveUnidadMin) = .TextMatrix(lngContador, cintColVsfCveUniAlternaArt)
29                        Else
                              'El articulo se maneja en una sola unidad, se dejan las cantidades en mínima
                              
                              'Máximo = consumo
30                            .TextMatrix(lngContador, cintColVsfCapturaMaximo) = rs!cantidad
31                            .TextMatrix(lngContador, cintColVsfCapturaUnidadMax) = .TextMatrix(lngContador, cintColvsfUnidadMinimaArt)
32                            .TextMatrix(lngContador, cintColVsfCapturaCveUnidadMax) = .TextMatrix(lngContador, cintColVsfCveUniMinimaArt)
                              
                              'Mínimo = consumo / 2
33                            .TextMatrix(lngContador, cintColVsfCapturaMinimo) = Round(rs!cantidad / 2, 0)
34                            .TextMatrix(lngContador, cintColVsfCapturaUnidadMin) = .TextMatrix(lngContador, cintColvsfUnidadMinimaArt)
35                            .TextMatrix(lngContador, cintColVsfCapturaCveUnidadMin) = .TextMatrix(lngContador, cintColVsfCveUniMinimaArt)
36                        End If
                          
37                        pPinta lngContador, clngColorNegro
38                    Else
                          
39                        .TextMatrix(lngContador, cintColVsfCapturaMaximo) = "0" ' max
40                        .TextMatrix(lngContador, cintColVsfCapturaUnidadMax) = "" 'unidad max
41                        .TextMatrix(lngContador, cintColVsfCapturaCveUnidadMax) = "" '????
42                        .TextMatrix(lngContador, cintColVsfCapturaMinimo) = "0" 'min
43                        .TextMatrix(lngContador, cintColVsfCapturaUnidadMin) = "" 'unidad min
44                        .TextMatrix(lngContador, cintColVsfCapturaCveUnidadMin) = "" '???
45                        .TextMatrix(lngContador, cintColVsfCapturaPunto) = "0"  'reorden
46                        .TextMatrix(lngContador, cintColVsfCapturaUnidadPun) = "" 'unidad min
47                        .TextMatrix(lngContador, cintColVsfCapturaCveUnidadPun) = "" '???
48                        .TextMatrix(lngContador, cintColVsfCapturaAlmacen) = ""
49                        .TextMatrix(lngContador, cintColVsfCapturaCveAlmacen) = ""
                          
50                        lngTotalSinConsumo = lngTotalSinConsumo + 1
                          
51                        pPinta lngContador, clngColorRojo
52                    End If
53                Next lngContador
                  
54                fraBarra.Visible = False
                  
55                If lngTotalSinConsumo <> 0 Then
                      'Se encontraron artículos sin consumo en el periodo
56                    MsgBox SIHOMsg(871), vbOKOnly + vbInformation, "Mensaje"
57                End If
              
58            End With
59        End If

60    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdCalcular_Click" & " Linea:" & Erl()))
End Sub

Private Function fblnFechasValidas() As Boolean
On Error GoTo NotificaError

    fblnFechasValidas = True
    
   
    If Len(Trim(mskInicio.Text)) < 10 Then
        fblnFechasValidas = False
        MsgBox SIHOMsg(29), vbOKOnly + vbExclamation, "Mensaje"
        mskInicio.SetFocus
        Exit Function
        
    Else
         If fblnValidaFecha(mskInicio) = False Then
            fblnFechasValidas = False
            MsgBox SIHOMsg(29), vbOKOnly + vbExclamation, "Mensaje"
            mskInicio.SetFocus
            Exit Function
          Else
           If CDate(mskInicio.Text) > CDate(fdtmServerFecha) Then
                 fblnFechasValidas = False
                'FECHA MENOR O IGUAL AL SISTEMA
                MsgBox SIHOMsg(40), vbOKOnly + vbExclamation, "Mensaje"
                mskInicio.SetFocus
                  Exit Function
           End If
         End If
    End If
    
    If Len(Trim(mskFin.Text)) < 10 Then
        fblnFechasValidas = False
         MsgBox SIHOMsg(29), vbOKOnly + vbExclamation, "Mensaje"
         mskFin.SetFocus
        Exit Function
     Else
        If fblnValidaFecha(mskFin) = False Then
            fblnFechasValidas = False
            MsgBox SIHOMsg(29), vbOKOnly + vbExclamation, "Mensaje"
            mskFin.SetFocus
            Exit Function
          Else
           If CDate(mskFin.Text) > CDate(fdtmServerFecha) Then
                 fblnFechasValidas = False
                'FECHA MENOR O IGUAL AL SISTEMA
                MsgBox SIHOMsg(40), vbOKOnly + vbExclamation, "Mensaje"
                mskFin.SetFocus
                  Exit Function
           End If
         End If
         
    End If
    
    

    If CDate(mskInicio.Text) < CDate("01/01/1900") Then   ' fechas no menores a 1900
    'If Not IsDate(mskInicio.Text) Or CDate(mskInicio.Text) < CDate("01/01/1900") Then ' fechas no menores a 1900
        fblnFechasValidas = False
        '¡Fecha no válida!, formato de fecha dd/mm/aaaa
        MsgBox SIHOMsg(29), vbOKOnly + vbExclamation, "Mensaje"
        mskInicio.SetFocus
    Else
        If CDate(mskFin.Text) < CDate("01/01/1900") Then  ' fechas no menores a 1900
        'If Not IsDate(mskFin.Text) Or CDate(mskFin.Text) < CDate("01/01/1900") Then ' fechas no menores a 1900
            fblnFechasValidas = False
            '¡Fecha no válida!, formato de fecha dd/mm/aaaa
            MsgBox SIHOMsg(29), vbOKOnly + vbExclamation, "Mensaje"
            mskFin.SetFocus
        Else
            If CDate(mskInicio.Text) > CDate(mskFin.Text) Then
                fblnFechasValidas = False
                '¡Rango de fechas no válido!
                MsgBox SIHOMsg(64), vbOKOnly + vbExclamation, "Mensaje"
                mskFin.SetFocus
            End If
        End If
    End If

Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fblnFechasValidas"))
End Function

Private Sub pPinta(lngRenglon As Long, lngColor As Long)
On Error GoTo NotificaError
Dim intcontador As Integer
    
    vsfArticulos.Row = lngRenglon
    
    For intcontador = cintColVsfIdArticulo To vsfArticulos.Cols - 1
        vsfArticulos.Col = intcontador
        vsfArticulos.CellForeColor = lngColor
    Next intcontador

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pPinta"))
End Sub

Private Sub cmdCerrar_Click()
On Error GoTo NotificaError

    fraArticulos.Enabled = True
    fraRequisitar.Enabled = True
    fraFiltros.Enabled = True

    fraRequisiciones.Visible = False
    
    cmdConsultarReq.SetFocus
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdCerrar_Click"))
End Sub

Private Sub cmdConsultarReq_Click()
1     On Error GoTo NotificaError
      Dim rsConSigna As ADODB.Recordset
      Dim vgstrParametrosConsigna As String

2         vgstrParametrosConsigna = 1 & "|" & vgintClaveEmpresaContable & "|" & CStr(cboDepartamento.ItemData(cboDepartamento.ListIndex))
3         Set rsConSigna = frsEjecuta_SP(vgstrParametrosConsigna, "Sp_Ivselalmacenconsigna")
          
          ' Valida artículos de consignación
4           vgstrParametrosSP = vsfArticulos.TextMatrix(vsfArticulos.Row, cintColVsfClaveArt) & "|" & Str(vgintDiasRequisicion) & "|" & Str(cboDepartamento.ItemData(cboDepartamento.ListIndex))
5           If rsConSigna.RecordCount = 0 Then
6             Set rs = frsEjecuta_SP(vgstrParametrosSP, "SP_IVSELREQARTPENDIENTE")
7           Else
8             Set rs = frsEjecuta_SP(vgstrParametrosSP, "SP_IVSELREQARTPEN_CONSIGNA")
9           End If
          
10        If rs.RecordCount <> 0 Then
              
11            lblNombreArticulo.Caption = vsfArticulos.TextMatrix(vsfArticulos.Row, cintColVsfNombreComercial)
              
12            With grdConsultaReq
13                .Clear
14                .Rows = 2
15                .Cols = cintColsConsulta
16                .FormatString = cstrTitulosConsulta
              
17                .ColWidth(0) = 100
18                .ColWidth(cintColConsFecha) = 1100
19                .ColWidth(cintColConsRequisicion) = 1000
20                .ColWidth(cintColConsTipo) = 1700
21                .ColWidth(cintColConsDepartamento) = 2500
22                .ColWidth(cintColConsPersona) = 3300
                  
23                .ColAlignment(cintColConsFecha) = flexAlignLeftCenter
24                .ColAlignment(cintColConsRequisicion) = flexAlignRightCenter
25                .ColAlignment(cintColConsTipo) = flexAlignLeftCenter
26                .ColAlignment(cintColConsDepartamento) = flexAlignLeftCenter
27                .ColAlignment(cintColConsPersona) = flexAlignLeftCenter
                  
28                .ColAlignmentFixed(cintColConsFecha) = flexAlignCenterCenter
29                .ColAlignmentFixed(cintColConsRequisicion) = flexAlignCenterCenter
30                .ColAlignmentFixed(cintColConsTipo) = flexAlignCenterCenter
31                .ColAlignmentFixed(cintColConsDepartamento) = flexAlignCenterCenter
32                .ColAlignmentFixed(cintColConsPersona) = flexAlignCenterCenter
                  
33                Do While Not rs.EOF
34                    .TextMatrix(.Rows - 1, cintColConsFecha) = Format(rs!fecha, "dd/mmm/yyyy")
35                    .TextMatrix(.Rows - 1, cintColConsRequisicion) = rs!numero
36                    .TextMatrix(.Rows - 1, cintColConsTipo) = rs!tipo
                      
37                    If rsConSigna.RecordCount = 0 Then
38                        .TextMatrix(.Rows - 1, cintColConsDepartamento) = rs!DepartamentoSurte
39                        .TextMatrix(.Rows - 1, cintColConsPersona) = rs!Empleado
40                    Else
41                        .TextMatrix(.Rows - 1, cintColConsPersona) = rs!Empleado
42                        .ColWidth(cintColConsDepartamento) = 0
43                    End If
                      
44                    .Rows = .Rows + 1
45                    rs.MoveNext
46                Loop
47                .Rows = .Rows - 1
48            End With
49        End If
          
50        fraArticulos.Enabled = rs.RecordCount = 0
51        fraRequisitar.Enabled = rs.RecordCount = 0
52        fraFiltros.Enabled = rs.RecordCount = 0
53        fraRequisiciones.Visible = rs.RecordCount <> 0
          
54        If rs.RecordCount <> 0 Then
55            cmdCerrar.SetFocus
56        End If

57        rsConSigna.Close

58    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdConsultarReq_Click" & " Linea:" & Erl()))
End Sub

Private Sub cmdDelete_Click()
On Error GoTo NotificaError

    If MsgBox(SIHOMsg(6), vbQuestion + vbYesNo, "Mensaje") = vbYes Then
        VsfBusArticulos.RemoveItem (VsfBusArticulos.RowSel)
        VsfBusArticulos.Refresh
        VsfBusArticulos.SetFocus
        If VsfBusArticulos.Row = -1 Then VsfBusArticulos.Row = VsfBusArticulos.Rows - 1
        
        VsfBusArticulos.Redraw = False
        pConfVsfBusArticulos
        VsfBusArticulos.Redraw = True
        
        If VsfBusArticulos.Rows < 2 Then
            cmdEnviarRequisicion.Enabled = False
            cmdDelete.Enabled = False
            cmdImprimir.Enabled = False
            SSTab.Tab = 0
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdEliminarRegistro_Click"))
End Sub

Private Sub cmdDeleteAll_Click()
    
    For lintCiclos = 0 To lstSeleccionadas.ListCount - 1
        lstDisponibles.AddItem lstSeleccionadas.List(lintCiclos)
        lstDisponibles.ItemData(lstDisponibles.NewIndex) = lstSeleccionadas.ItemData(lintCiclos)
    Next lintCiclos
    
    lstDisponibles.ListIndex = 0
    lstSeleccionadas.Clear
    
End Sub

Private Sub cmdDeshacer_Click()
    lbFlag = True
    cmdRequisitar_Click
End Sub

Private Sub cmdEnviarRequisicion_Click()
1     On Error GoTo NotificaError
      Dim lngContador As Long
      Dim lngRenglon As Long
      Dim lngCveDeptoSurte As Long
      Dim lngNumRequisicion As Long
      Dim rsConSigna As ADODB.Recordset
      Dim vgstrParametrosSP As String
      Dim i As Integer
      Dim rsAutorizaciones As New ADODB.Recordset
      Dim rsTemp As New ADODB.Recordset
      Dim lintAutorizaRequi As Integer

      Dim lngContadorOrden As Long
      Dim rsBuscaOrden  As New ADODB.Recordset
      Dim strSentencia As String
      Dim lstrOrdenOrdenada As String
      Dim vstrArticuloOrden As String
      Dim lblnIRemotaSubrogados As Boolean
      Dim arrRequis() As Long     ' requisiciones reubicacion a incluir en impresion remota
      Dim lintTotalRequis As Integer
      Dim intCount As Integer
      
      Dim intCountDeptosEsclavo As Integer
      Dim intCantArtiTotDistConsol As Integer
      
      Dim intExisDeptoPrin As Integer
      Dim intMaxDeptoPrin As Integer
      Dim intExisDeptoEsclavo As Integer
      Dim intMaxDeptoEsclavo As Integer
          
      Dim intTotDeptosConsol As Integer
      Dim DeptosConsol() As String
      Dim ExisEsclavoConsol() As String
      Dim MaximEsclavoConsol() As String

2         i = 1
3         vgstrParametrosSP = 1 & "|" & vgintClaveEmpresaContable & "|" & CStr(cboDepartamento.ItemData(cboDepartamento.ListIndex))
4         Set rsConSigna = frsEjecuta_SP(vgstrParametrosSP, "SP_IVSELALMACENCONSIGNA")
          
5         If fblnValidaRequisicion() Then 'Validar permisos y contraseña
6             cmdEnviarRequisicion.Enabled = False
7             cmdDelete.Enabled = False
8             cmdDeshacer.Enabled = False

9             frsEjecuta_SP 1 & "|" & Me.Name & "|" & VsfBusArticulos.Name & "|Orden|" & vglngNumeroLogin & "|" & vlstrColSort, "SP_GNSELULTIMACONFIGURACION", True

10            If chkCompradirecta.Value = 0 And cboTipoRequisicion.ItemData(cboTipoRequisicion.ListIndex) = 1 And Not lblnAlmacenConsigna Then
11                Set rsTemp = frsSelParametros("CO", -1, "BITAUTORIZAREQUISICION")
12                If Not rsTemp.EOF Then
13                    lintAutorizaRequi = IIf(rsTemp!Valor = "0", 0, 1)
14                End If
15                rsTemp.Close
16            End If
                      
              
17            strSentencia = " Select coordendetalle.INTNUMORDEN, coordendetalle.NUMNUMREQUISICION, coordendetalle.CHRCVEARTICULO, coordendetalle.VCHESTATUSARTICULO, ivarticulo.VCHNOMBRECOMERCIAL  " & _
                         " from coordendetalle " & _
                         " inner join coordenmaestro on coordenmaestro.INTNUMORDEN = coordendetalle.INTNUMORDEN  " & _
                         " inner join ivarticulo on ivarticulo.CHRCVEARTICULO = coordendetalle.CHRCVEARTICULO " & _
                         " inner join ivrequisicionmaestro on ivrequisicionmaestro.NUMNUMREQUISICION = coordendetalle.NUMNUMREQUISICION " & _
                         " where coordenmaestro.VCHESTATUSORDEN = 'ORDENADA' AND coordendetalle.VCHESTATUSARTICULO = 'ORDENADA' " & _
                         " and ivrequisicionmaestro.CHRDESTINO = 'C' " & _
                         " and ivrequisicionmaestro.SMICVEDEPTOREQUIS = " & cboDepartamento.ItemData(cboDepartamento.ListIndex)
                         
                         
18            Set rsBuscaOrden = frsRegresaRs(strSentencia, adLockReadOnly, adOpenForwardOnly)
          
19            grdRequisicion.Clear
20            grdRequisicion.Rows = 2

              If strClavesConsol <> "" Then
                grdRequisicion.Cols = cintColReqArtiConsol
              Else
21              grdRequisicion.Cols = cintColsRequisicion
              End If
              
22            lstrOrdenOrdenada = ""
          
23            For lngContador = 1 To VsfBusArticulos.Rows - 1
24                If rsBuscaOrden.RecordCount > 0 Then
25                    rsBuscaOrden.MoveFirst
26                    lngContadorOrden = 0
27                    vstrArticuloOrden = ""
28                    For lngContadorOrden = 1 To rsBuscaOrden.RecordCount
29                         If VsfBusArticulos.TextMatrix(lngContador, cintColBusClave) = rsBuscaOrden!CHRCVEARTICULO Then
30                            lstrOrdenOrdenada = IIf(lstrOrdenOrdenada <> "0", lstrOrdenOrdenada, "") & Chr(13) & rsBuscaOrden!intnumOrden & " , " & VsfBusArticulos.TextMatrix(lngContador, cintColBusClave) & " , " & rsBuscaOrden!VCHNOMBRECOMERCIAL
31                            vstrArticuloOrden = rsBuscaOrden!CHRCVEARTICULO
32                            If rsBuscaOrden.EOF Then
33                                Exit For
34                            End If
35                         End If
36                         rsBuscaOrden.MoveNext
37                    Next lngContadorOrden
38                End If
                       
39                If vstrArticuloOrden <> VsfBusArticulos.TextMatrix(lngContador, cintColBusClave) Then
40                   grdRequisicion.TextMatrix(grdRequisicion.Rows - 1, cintColGrdCveDeptoSurte) = VsfBusArticulos.TextMatrix(lngContador, cintColBusIdAlmacen)
41                   grdRequisicion.TextMatrix(grdRequisicion.Rows - 1, cintColGrdCveArticulo) = VsfBusArticulos.TextMatrix(lngContador, cintColBusClave)
42                   grdRequisicion.TextMatrix(grdRequisicion.Rows - 1, cintColGrdCantidad) = VsfBusArticulos.TextMatrix(lngContador, cintColBusCantidad)
43                   grdRequisicion.TextMatrix(grdRequisicion.Rows - 1, cintColGrdTipoUnidad) = VsfBusArticulos.TextMatrix(lngContador, cintColBusTipoUnidad)
44                   grdRequisicion.TextMatrix(grdRequisicion.Rows - 1, cintColGrdCveDeptoAutoriza) = VsfBusArticulos.TextMatrix(lngContador, cintColBusCveDeptoAutoriza)
45                   grdRequisicion.TextMatrix(grdRequisicion.Rows - 1, cintColGrdCveDeptoRecibe) = VsfBusArticulos.TextMatrix(lngContador, cintColBusCveDeptoRecibe)
                     If strClavesConsol <> "" Then
                        grdRequisicion.TextMatrix(grdRequisicion.Rows - 1, cintColReqArtiExisPrin) = VsfBusArticulos.TextMatrix(lngContador, cintColBusExisPrin)
                        grdRequisicion.TextMatrix(grdRequisicion.Rows - 1, cintColReqArtiMaxPrin) = VsfBusArticulos.TextMatrix(lngContador, cintColBusMaxPrin)
                    
                        grdRequisicion.TextMatrix(grdRequisicion.Rows - 1, cintColReqArtiExisEsclavos) = VsfBusArticulos.TextMatrix(lngContador, cintColBusExisEsclavos)
                        grdRequisicion.TextMatrix(grdRequisicion.Rows - 1, cintColReqArtiMaxEsclavos) = VsfBusArticulos.TextMatrix(lngContador, cintColBusMaxEsclavos)
                        grdRequisicion.TextMatrix(grdRequisicion.Rows - 1, cintColReqArtiDeptoEsclavos) = VsfBusArticulos.TextMatrix(lngContador, cintColBusDeptoEsclavos)
                    End If
46                  grdRequisicion.Rows = grdRequisicion.Rows + 1
47                End If
48            Next lngContador
           
49            grdRequisicion.Rows = grdRequisicion.Rows - 1
              
50            EntornoSIHO.ConeccionSIHO.BeginTrans

              'Ordenar el grid por almacén y enviar las requisiciones:
51            With grdRequisicion
52                .Col = cintColGrdCveDeptoSurte
53                .Sort = 1
          
54                lintTotalRequis = 0
55                ReDim arrRequis(lintTotalRequis)
56                lngCveDeptoSurte = 0
57                For lngContador = 1 To .Rows - 1
58                    lblnIRemotaSubrogados = False
                      
59                    If lngCveDeptoSurte <> Val(.TextMatrix(lngContador, cintColGrdCveDeptoSurte)) Then
60                        vgstrParametrosSP = CStr(cboDepartamento.ItemData(cboDepartamento.ListIndex)) _
                                              & "|" & Str(llngPersonaGraba) _
                                              & "|" & .TextMatrix(lngContador, cintColGrdCveDeptoSurte) _
                                              & "|" & fstrFechaSQL(fdtmServerFecha) _
                                              & "|" & Format(fdtmServerHora, "hh:mm:ss") _
                                              & "|" & lstrTipoRequisicion _
                                              & "|" & "0" & "|" & "0" & "|" & "0" & "|" & " " & "|" & " " _
                                              & "|" & Str(vglngNumeroLogin) _
                                              & "|" & " " _
                                              & "|" & CStr(vgintNumeroDepartamento) _
                                              & "|" & chkCompradirecta.Value
61                        lngNumRequisicion = 1
62                        frsEjecuta_SP vgstrParametrosSP, "SP_IVINSREQUISICIONMAESTRO", True, lngNumRequisicion

                          ' Guarda referencia de requisición automática
63                        frsEjecuta_SP CStr(lngNumRequisicion), "SP_IVINSREQUISICIONAUTOMATICA"
                                             
64                        ReDim Preserve aRequisiciones(i)
                          
                          ' Guarda número de requisiciones en arreglo
65                        aRequisiciones(i).lgnNumRequisicion = lngNumRequisicion
                          
66                        i = i + 1
67                        If lstrTipoRequisicion = "C" Then ' Valida requisición tipo consignación
68                            If rsConSigna.RecordCount = 0 Then
                                  'Insertar datos en IVREQUISICIONCOMPRA:
69                                vgstrParametrosSP = Str(lngNumRequisicion) _
                                  & "|" & IIf(Val(.TextMatrix(lngContador, cintColGrdCveDeptoAutoriza)) = 0, Null, .TextMatrix(lngContador, cintColGrdCveDeptoAutoriza)) _
                                  & "|" & IIf(Val(.TextMatrix(lngContador, cintColGrdCveDeptoRecibe)) = 0, Null, .TextMatrix(lngContador, cintColGrdCveDeptoRecibe))
70                                frsEjecuta_SP vgstrParametrosSP, "SP_IVINSREQUISICIONCOMPRA"
71                            End If
72                        End If
                          
73                        If chkCompradirecta.Value = 0 Then
74                            If lblnAlmacenConsigna Then
75                                pImpresionRemota "RR", lngNumRequisicion, .TextMatrix(lngContador, cintColGrdCveDeptoSurte)
76                            Else
77                                If cboTipoRequisicion.ItemData(cboTipoRequisicion.ListIndex) = 0 Then ' Reubicacion
78                                    Set rsAutorizaciones = frsRegresaRs("SELECT NVL(COUNT(*),0) Autoriza FROM IVAUTORIZACIONREQUISICIONES WHERE intcvedepartamento = " & cboDepartamento.ItemData(cboDepartamento.ListIndex) & " AND chrtiporequisicion = 'R'", adLockOptimistic, adOpenDynamic)
79                                    If rsAutorizaciones!Autoriza = 0 Then
80                                        lblnIRemotaSubrogados = True
81                                        pImpresionRemota "RR", lngNumRequisicion, .TextMatrix(lngContador, cintColGrdCveDeptoSurte)
82                                    End If
83                                Else
84                                    If cboTipoRequisicion.ItemData(cboTipoRequisicion.ListIndex) = 1 And lintAutorizaRequi = 0 Then 'Compra - pedido
85                                        pImpresionRemota "RC", lngNumRequisicion, .TextMatrix(lngContador, cintColGrdCveDeptoSurte)
86                                    End If
87                                End If
88                            End If
89                        End If
                          
90                        pGuardarLogTransaccion Me.Name, EnmGrabar, llngPersonaGraba, "REQUISITAR FALTANTES AUTOMATICAMENTE", Str(lngNumRequisicion)

91                        lngCveDeptoSurte = Val(.TextMatrix(lngContador, cintColGrdCveDeptoSurte))
92                    End If

93                    If rsConSigna.RecordCount > 0 Then
94                        vgstrParametrosSP = CStr(cboDepartamento.ItemData(cboDepartamento.ListIndex)) _
                                              & "|" & Str(llngPersonaGraba) _
                                              & "|" & .TextMatrix(lngContador, cintColGrdCveDeptoSurte) _
                                              & "|" & fstrFechaSQL(fdtmServerFecha) _
                                              & "|" & Format(fdtmServerHora, "hh:mm:ss") _
                                              & "|" & "U" & "|" & "0" & "|" & "0" & "|" & "0" & "|" & " " & "|" & " " _
                                              & "|" & Str(vglngNumeroLogin) _
                                              & "|" & " " _
                                              & "|" & CStr(vgintNumeroDepartamento) _
                                              & "|" & chkCompradirecta.Value
95                        If lngContador = 1 Then
96                            lngNumRequisicion = 1
97                            frsEjecuta_SP vgstrParametrosSP, "SP_IVINSREQUISICIONMAESTRO", True, lngNumRequisicion
                              
                              ' Guarda referencia de requisición automática
98                            frsEjecuta_SP CStr(lngNumRequisicion), "SP_IVINSREQUISICIONAUTOMATICA"
99                        End If
100                   End If

101                   vgstrParametrosSP = Str(lngNumRequisicion) _
                                          & "|" & .TextMatrix(lngContador, cintColGrdCveArticulo) _
                                          & "|" & .TextMatrix(lngContador, cintColGrdCantidad) _
                                          & "|" & .TextMatrix(lngContador, cintColGrdTipoUnidad) _
                                          & "|" & "PENDIENTE||"
102                   frsEjecuta_SP vgstrParametrosSP, "SP_IVINSREQUISICIONDETALLE"

                If strClavesConsol <> "" Then
                 
                    intCount = 0
                    intCantArtiTotDistConsol = 0
                 
                    ' Almacen Principal
                    intExisDeptoPrin = 0
                    intMaxDeptoPrin = 0
                 
                    'Total de Departamentos
                    intTotDeptosConsol = 0
                    DeptosConsol = Split(grdRequisicion.TextMatrix(lngContador, cintColReqArtiDeptoEsclavos), ",")
                    intTotDeptosConsol = UBound(DeptosConsol) + 1
                 
                    'Registro de Articulos Departamento Requisición Consolida
                    ReDim Preserve aDeptoArtiConsol(intTotDeptosConsol)
                 
                    'Almacen Esclavo Arreglo (n)
                    intExisDeptoEsclavo = 0
                    intMaxDeptoEsclavo = 0
                 
                    ExisEsclavoConsol = Split(grdRequisicion.TextMatrix(lngContador, cintColReqArtiExisEsclavos), ",")
                    MaximEsclavoConsol = Split(grdRequisicion.TextMatrix(lngContador, cintColReqArtiMaxEsclavos), ",")
                 
                    If UBound(ExisEsclavoConsol) < intTotDeptosConsol - 1 Then
                       ReDim Preserve ExisEsclavoConsol(intTotDeptosConsol)
                    End If
            
                    If UBound(MaximEsclavoConsol) < intTotDeptosConsol - 1 Then
                        ReDim Preserve MaximEsclavoConsol(intTotDeptosConsol)
                    End If
                 
                    For intCount = 0 To UBound(ExisEsclavoConsol)
                        intExisDeptoEsclavo = intExisDeptoEsclavo + IIf(ExisEsclavoConsol(intCount) = "", 0, ExisEsclavoConsol(intCount))
                    Next intCount
                 
                    For intCount = 0 To UBound(MaximEsclavoConsol)
                        intMaxDeptoEsclavo = intMaxDeptoEsclavo + IIf(MaximEsclavoConsol(intCount) = "", 0, MaximEsclavoConsol(intCount))
                    Next intCount
                                 
                    'Almacen Principal
                    intExisDeptoPrin = grdRequisicion.TextMatrix(lngContador, cintColReqArtiExisPrin) - intExisDeptoEsclavo
                    intMaxDeptoPrin = grdRequisicion.TextMatrix(lngContador, cintColReqArtiMaxPrin) - intMaxDeptoEsclavo
                 
                    'Cantidad Total Solicitar
                    intCantArtiTotDistConsol = .TextMatrix(lngContador, cintColGrdCantidad)
                 
                    intCount = 0
                 
                    'Distribución de Articulos Alamcen Principal
                    If (intExisDeptoPrin < intMaxDeptoPrin) Then
                        aDeptoArtiConsol(intCount).intCveDepto = vgintNumeroDepartamento
                        aDeptoArtiConsol(intCount).SrtCveArticulo = .TextMatrix(lngContador, cintColGrdCveArticulo)
                        
                        If (intMaxDeptoPrin - intExisDeptoPrin) >= intCantArtiTotDistConsol Then
                            aDeptoArtiConsol(intCount).intCantidad = intCantArtiTotDistConsol
                            intCantArtiTotDistConsol = 0
                        Else
                            aDeptoArtiConsol(intCount).intCantidad = (intMaxDeptoPrin - intExisDeptoPrin)
                            intCantArtiTotDistConsol = intCantArtiTotDistConsol - (intMaxDeptoPrin - intExisDeptoPrin)
                        End If
                        intCount = intCount + 1
                    End If
                 
                    'Distribución de Articulos Almacen Esclavo
                    If intCantArtiTotDistConsol > 0 Then
                        For intCountDeptosEsclavo = 0 To intTotDeptosConsol - 1
                            If (ExisEsclavoConsol(intCountDeptosEsclavo) < MaximEsclavoConsol(intCountDeptosEsclavo)) Then
                                aDeptoArtiConsol(intCount).intCveDepto = DeptosConsol(intCountDeptosEsclavo)
                                aDeptoArtiConsol(intCount).SrtCveArticulo = .TextMatrix(lngContador, cintColGrdCveArticulo)
                             If (MaximEsclavoConsol(intCountDeptosEsclavo) - ExisEsclavoConsol(intCountDeptosEsclavo)) >= intCantArtiTotDistConsol Then
                                 aDeptoArtiConsol(intCount).intCantidad = intCantArtiTotDistConsol
                                 intCantArtiTotDistConsol = 0
                             Else
                                 aDeptoArtiConsol(intCount).intCantidad = (MaximEsclavoConsol(intCountDeptosEsclavo) - ExisEsclavoConsol(intCountDeptosEsclavo))
                                 intCantArtiTotDistConsol = intCantArtiTotDistConsol - (MaximEsclavoConsol(intCountDeptosEsclavo) - ExisEsclavoConsol(intCountDeptosEsclavo))
                             End If
                            
                             intCount = intCount + 1
                            End If
                        Next intCountDeptosEsclavo
                    End If
                 
                    For intCantArtiTotDistConsol = 0 To intCount - 1
                        If aDeptoArtiConsol(intCantArtiTotDistConsol).SrtCveArticulo <> "" Then
                        
                          vgstrParametrosSP = Str(lngNumRequisicion) _
                                    & "|" & aDeptoArtiConsol(intCantArtiTotDistConsol).SrtCveArticulo _
                                    & "|" & .TextMatrix(lngContador, cintColGrdTipoUnidad) _
                                    & "|" & aDeptoArtiConsol(intCantArtiTotDistConsol).intCantidad _
                                    & "|0" _
                                    & "|" & aDeptoArtiConsol(intCantArtiTotDistConsol).intCveDepto _
                                    & "|" & IIf(aDeptoArtiConsol(intCantArtiTotDistConsol).intCveDepto = lbCveAlmacenPrin, "P", "S")
                                    
                          frsEjecuta_SP vgstrParametrosSP, "sp_IvInsRequiConsoli"
                        End If
                    Next intCantArtiTotDistConsol
                End If
                    
                      
103                   If lblnIRemotaSubrogados = True Then
104                       ReDim Preserve arrRequis(lintTotalRequis)
105                       arrRequis(lintTotalRequis) = lngNumRequisicion
106                       lintTotalRequis = lintTotalRequis + 1
107                   End If
108               Next lngContador
109           End With
              
110           For intCount = 0 To UBound(arrRequis())
                  ' Se inserta registro de impresion remota de subrogados, se hace aqui ya que se guardó el detalle de la requi
111               pInsertaSubrrogados arrRequis(intCount), "RR"
112           Next intCount
                  
113           EntornoSIHO.ConeccionSIHO.CommitTrans
              
114           If lblnAlmacenConsigna Then
115               pImprimeLocal lngNumRequisicion
                  
                  'La operación se realizó satisfactoriamente.
116               MsgBox SIHOMsg(420), vbOKOnly + vbInformation, "Mensaje"
                  
117               SSTab.Tab = 0
118               VsfBusArticulos.Editable = 0
119               cmdAgregar_Click
120           Else
                  'La operación se realizó satisfactoriamente.
121               MsgBox SIHOMsg(420), vbOKOnly + vbInformation, "Mensaje"
                  
122               cmdImprimir.Enabled = True
123               cmdImprimir.SetFocus
124           End If
125       End If
          
126       If lstrOrdenOrdenada <> "" Then
127           MsgBox ("Artículos no incluidos en la requisición por tener orden de compra.") & Chr(13) & "Orden:" & " , " & "Artículo:" & " , " & "Nombre:" & Chr(13) & lstrOrdenOrdenada, vbOKOnly + vbInformation, "Mensaje"
128       End If
          
129       rsConSigna.Close

130   Exit Sub
NotificaError:
       Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdEnviarRequisicion_Click" & " Linea:" & Erl()))
End Sub

Private Function fblnValidaRequisicion() As Boolean
On Error GoTo NotificaError

    fblnValidaRequisicion = True
    
    fblnValidaRequisicion = fblnRevisaPermiso(vglngNumeroLogin, llngNumOpcionRequisitar, "E")
    
    If fblnValidaRequisicion Then
        llngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
        fblnValidaRequisicion = llngPersonaGraba <> 0
    End If

Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fblnValidaRequisicion"))
End Function

Private Sub pImprimeLocal(lngRequisicion As Long)
On Error GoTo NotificaError
Dim alstrParametros(1) As String
Dim rsReporte As ADODB.Recordset
    
    pInstanciaReporte vgrptReporte, "rptrequisicion.rpt"
    
    Set rsReporte = frsEjecuta_SP(CStr(lngRequisicion) & "|" & CStr(vglngCveAlmacenGeneral), "SP_IVRPTREQUISICIONES")
    If rsReporte.RecordCount > 0 Then
          vgrptReporte.DiscardSavedData
                   
          alstrParametros(0) = "tienehistorico;" & "0"
          alstrParametros(1) = "empresa;" & Trim(vgstrNombreHospitalCH)
          pCargaParameterFields alstrParametros, vgrptReporte
        
          pImprimeReporte vgrptReporte, rsReporte, "I", "Requisición de artículos"
    Else
        'No existe información con esos parámetros.
        MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje"
    End If
    If rsReporte.State <> adStateClosed Then rsReporte.Close

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pImprimeLocal"))
End Sub
Public Sub Exportar_Excel()

    Dim oExcel As Object
    Dim oBook As Object
    Dim oSheet As Object
    
    Dim vlCont As Integer
       
    
     On Error GoTo ErrHandler
    
        CommonDialog1.CancelError = True
        CommonDialog1.Flags = cdlOFNHideReadOnly + cdlOFNOverwritePrompt + cdlOFNPathMustExist
        CommonDialog1.Filter = "Libro de Excel(*.xlsx)|*.xlsx"
        CommonDialog1.FileName = ""
        CommonDialog1.ShowSave
    
    
    'Start a new workbook in Excel
    Set oExcel = CreateObject("Excel.Application")
    Set oBook = oExcel.Workbooks.Add
      
    'Add data to cells of the first worksheet in the new workbook
    Set oSheet = oBook.Worksheets(1)
    
    oExcel.Cells(2, 1).Value = Trim(vgstrNombreHospitalCH)
    oExcel.Cells(3, 1).Value = "MÁXIMOS, MÍNIMOS Y PUNTOS DE REORDEN"
    oExcel.Cells(4, 1).Value = "DEPARTAMENTO " + Trim(CStr(cboDepartamento.List(cboDepartamento.ListIndex)))
    
    
            
    oSheet.Range("A6").Value = "Clave del artículo"
    oSheet.Range("B6").Value = "Nombre comercial"
    oSheet.Range("C6").Value = "Máximo"
    oSheet.Range("D6").Value = "Unidad del máximo"
    oSheet.Range("E6").Value = "Reorden"
    oSheet.Range("F6").Value = "Unidad del reorden"
    oSheet.Range("G6").Value = "Mínimo"
    oSheet.Range("H6").Value = "Unidad del mínimo"
    oSheet.Range("I6").Value = "Almacén reubica"
    oSheet.Range("J6").Value = "Departamento compra"
    oSheet.Range("A1:J6").Font.Bold = True
    oSheet.Range("A6:J6").Interior.ColorIndex = 15 '15 48
    
    oSheet.Range("A7").Select
    oExcel.ActiveWindow.FreezePanes = True
    
    oSheet.Range("A:J").Font.Name = "Times New Roman" '
    oSheet.Range("A:J").Font.Size = 10
    
    oSheet.Range("A:A").ColumnWidth = 20
    oSheet.Range("B:B").ColumnWidth = 40
    oSheet.Range("C:C").ColumnWidth = 10
    oSheet.Range("D:D").ColumnWidth = 20
    oSheet.Range("E:E").ColumnWidth = 10
    oSheet.Range("F:F").ColumnWidth = 20
    oSheet.Range("G:G").ColumnWidth = 10
    oSheet.Range("H:H").ColumnWidth = 20
    oSheet.Range("I:I").ColumnWidth = 20
    oSheet.Range("J:J").ColumnWidth = 20
    

    For vlCont = 1 To vsfArticulos.Rows - 1
        oExcel.Cells(vlCont + 6, 1).NumberFormat = "@"
        oExcel.Cells(vlCont + 6, 1).HorizontalAlignment = xlRight_
        oExcel.Cells(vlCont + 6, 1).Value = Trim(vsfArticulos.TextMatrix(vlCont, 20))
        oExcel.Cells(vlCont + 6, 2).Value = Trim(vsfArticulos.TextMatrix(vlCont, 6))
        oExcel.Cells(vlCont + 6, 3).Value = Trim(vsfArticulos.TextMatrix(vlCont, 7))
        oExcel.Cells(vlCont + 6, 4).Value = Trim(vsfArticulos.TextMatrix(vlCont, 8))
        oExcel.Cells(vlCont + 6, 5).Value = Trim(vsfArticulos.TextMatrix(vlCont, 10))
        oExcel.Cells(vlCont + 6, 6).Value = Trim(vsfArticulos.TextMatrix(vlCont, 11))
        oExcel.Cells(vlCont + 6, 7).Value = Trim(vsfArticulos.TextMatrix(vlCont, 13))
        oExcel.Cells(vlCont + 6, 8).Value = Trim(vsfArticulos.TextMatrix(vlCont, 14))
        oExcel.Cells(vlCont + 6, 9).Value = Trim(vsfArticulos.TextMatrix(vlCont, 16))
        oExcel.Cells(vlCont + 6, 10).Value = Trim(vsfArticulos.TextMatrix(vlCont, 18))
    Next
        
      
            'Guardamos el libro y salimos de Excel
            oBook.SaveAs CommonDialog1.FileName
            MsgBox "La exportación se realizó con éxito.", vbInformation, "Mensaje"
            oExcel.Visible = True
            Set oExcel = Nothing
     
     
ErrHandler:
    Select Case Err
        Case 32755
            MsgBox "Usted canceló la exportación", vbInformation + vbOKOnly, "Mensaje"
        Case Else
            'MsgBox "Unexpected error. Err" & Err & ": " & Error, vbInformation + vbOKOnly, "Mensaje"
    End Select
    Exit Sub
    
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdExportar_Click"))
    Exit Sub
     
End Sub
Private Sub cmdExportar_Click(Index As Integer)
Call Exportar_Excel
End Sub

Private Sub cmdGrabar_Click()
1     On Error GoTo NotificaError
      Dim lngContador As Long

2         If fblnDatosValidos() Then
          
3             llngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
              
4             If llngPersonaGraba <> 0 Then
              
5                 With vsfArticulos
6                     ldblAumento = 100 / (.Rows - 1)
                  
7                     pgbBarra.Value = 0
8                     fraBarra.Visible = True
9                     lblTituloBarra.Caption = "Actualizando máximos y mínimos, espere un momento"
                      
10                    For lngContador = 1 To .Rows - 1
                      
                          '*- Barra de avance
11                        fraBarra.Refresh
12                        If pgbBarra.Value + ldblAumento > 100 Then
13                            pgbBarra.Value = 100
14                        Else
15                            pgbBarra.Value = pgbBarra.Value + ldblAumento
16                        End If
                          '*- Barra de avance
                      
17                        If Val(.TextMatrix(lngContador, cintcolvsfGuardar)) = 1 Then
                              'Guardar:
18                            vgstrParametrosSP = _
                              Str(cboDepartamento.ItemData(cboDepartamento.ListIndex)) _
                              & "|" & .TextMatrix(lngContador, cintColVsfIdArticulo) _
                              & "|" & Str(Val(.TextMatrix(lngContador, cintColVsfCapturaCveAlmacen))) _
                              & "|" & Str(Val(.TextMatrix(lngContador, cintColVsfCapturaCveDeptoCompra))) _
                              & "|" & .TextMatrix(lngContador, cintColVsfCapturaMaximo) _
                              & "|" & IIf(Val(.TextMatrix(lngContador, cintColvsfCapturaTipoUniMax)) = 1, "1", "0") _
                              & "|" & .TextMatrix(lngContador, cintColVsfCapturaPunto) _
                              & "|" & IIf(Val(.TextMatrix(lngContador, cintColVsfCapturaPunto)) = 0 And Val(.TextMatrix(lngContador, cintColvsfCapturaTipoUniPun)) = 0, "3", IIf(Val(.TextMatrix(lngContador, cintColvsfCapturaTipoUniPun)) = 1, "1", "0")) _
                              & "|" & .TextMatrix(lngContador, cintColVsfCapturaMinimo) _
                              & "|" & IIf(Val(.TextMatrix(lngContador, cintColVsfCapturaMinimo)) = 0 And Val(.TextMatrix(lngContador, cintColvsfCapturaTipoUniMin)) = 0, "3", IIf(Val(.TextMatrix(lngContador, cintColvsfCapturaTipoUniMin)) = 1, "1", "0")) _
      
19                            frsEjecuta_SP vgstrParametrosSP, "SP_IVINSMAXIMOMINIMO"
                              
20                            .TextMatrix(lngContador, cintcolvsfGuardar) = ""
                          
21                        ElseIf Val(.TextMatrix(lngContador, cintcolvsfGuardar)) = -1 Then
                              'Eliminar:
22                            vgstrParametrosSP = Str(cboDepartamento.ItemData(cboDepartamento.ListIndex)) & "|" & .TextMatrix(lngContador, cintColVsfIdArticulo)
23                            frsEjecuta_SP vgstrParametrosSP, "SP_IVDELMAXIMOMINIMO"
                              
24                            .TextMatrix(lngContador, cintcolvsfGuardar) = ""
25                        End If
26                    Next lngContador
                      
27                    fraBarra.Visible = False
28                End With
                  
29                pGuardarLogTransaccion Me.Name, EnmGrabar, llngPersonaGraba, "APLICAR MAXIMOS Y MINIMOS", Str(cboDepartamento.ItemData(cboDepartamento.ListIndex))
                  
                        
                  
30                AsignacionAutomatica = True
31                cmdAsignarDesactivar_Click (0)
32                AsignacionAutomatica = False
                  
                  
                  
                  
                  'La información se actualizó satisfactoriamente.
33                MsgBox SIHOMsg(284), vbOKOnly + vbInformation, "Mensaje"
34                cmdGrabar.Enabled = False
35                If fblnCanFocus(cboDepartamento) Then cboDepartamento.SetFocus Else optOpcion(0).SetFocus
                  
36            End If
37        End If

38    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdGrabar_Click" & " Linea:" & Erl()))
End Sub

Private Function fblnDatosValidos() As Boolean
1     On Error GoTo NotificaError
      Dim lngCuenta As Long
      Dim intMensaje As Integer
      Dim vlblnAsignarDepto As Integer ' indica si un articulo cuenta con max, min y punto de reorden para asi ser asignado al departamento
      Dim vlCveAlmaDeptoReCo As String

          'Revisar si tiene permisos:
2         fblnDatosValidos = fblnRevisaPermiso(vglngNumeroLogin, llngNumOpcionGuardar, "E")
          
3         If fblnDatosValidos Then
          
4             With vsfArticulos
              
5                 lngCuenta = 1
6                 Do While fblnDatosValidos And lngCuenta <= .Rows - 1
7                     vlblnAsignarDepto = True ' cada nuevo renglon la varible se activa
8                     If Val(.TextMatrix(lngCuenta, cintcolvsfGuardar)) = 1 Then
                      
                          'Cuando todo está vacío:
9                         If Val(.TextMatrix(lngCuenta, cintColVsfCapturaMaximo)) = 0 And _
                          (Val(.TextMatrix(lngCuenta, cintColVsfCapturaCveUnidadMax)) = -1 Or Val(.TextMatrix(lngCuenta, cintColVsfCapturaCveUnidadMax)) = 0) And _
                          Val(.TextMatrix(lngCuenta, cintColVsfCapturaMinimo)) = 0 And _
                          (Val(.TextMatrix(lngCuenta, cintColVsfCapturaCveUnidadMin)) = -1 Or Val(.TextMatrix(lngCuenta, cintColVsfCapturaCveUnidadMin)) = 0) And _
                          Val(.TextMatrix(lngCuenta, cintColVsfCapturaCveAlmacen)) = 0 And _
                          Val(.TextMatrix(lngCuenta, cintColVsfCapturaPunto)) = 0 And _
                          (Val(.TextMatrix(lngCuenta, cintColVsfCapturaCveUnidadPun)) = -1 Or Val(.TextMatrix(lngCuenta, cintColVsfCapturaCveUnidadPun)) = 0) And _
                          Val(.TextMatrix(lngCuenta, cintColVsfCapturaCveDeptoCompra)) = 0 Then
                              'Se marca al artículo para ser borrado:
10                            .TextMatrix(lngCuenta, cintcolvsfGuardar) = "-1"
                              
11                        Else
12                            If Val(.TextMatrix(lngCuenta, cintColVsfCapturaMaximo)) = 0 Then
                                  'Si no se capturó máximo:
13                                fblnDatosValidos = False
14                                intMensaje = 2
15                                vsfArticulos.Row = lngCuenta
16                                vsfArticulos.Col = cintColVsfCapturaMaximo
                                  
17                            ElseIf Val(.TextMatrix(lngCuenta, cintColVsfCapturaCveUnidadMax)) = -1 Or Val(.TextMatrix(lngCuenta, cintColVsfCapturaCveUnidadMax)) = 0 Then
                                  'Si no se capturó la unidad del máximo:
18                                fblnDatosValidos = False
19                                intMensaje = 2
20                                vsfArticulos.Row = lngCuenta
21                                vsfArticulos.Col = cintColVsfCapturaUnidadMax
                                  
22                            ElseIf (Val(.TextMatrix(lngCuenta, cintColVsfCapturaCveUnidadPun)) = -1 Or Val(.TextMatrix(lngCuenta, cintColVsfCapturaCveUnidadPun)) = 0) And Val(.TextMatrix(lngCuenta, cintColVsfCapturaPunto)) <> 0 Then
                                  'Si no se capturó la unidad del reorden y se escribió punto de reorden:
23                                fblnDatosValidos = False
24                                intMensaje = 2
25                                vsfArticulos.Row = lngCuenta
26                                vsfArticulos.Col = cintColVsfCapturaUnidadPun
                               
27                            ElseIf (Val(.TextMatrix(lngCuenta, cintColVsfCapturaCveUnidadMin)) = -1 Or Val(.TextMatrix(lngCuenta, cintColVsfCapturaCveUnidadMin)) = 0) And Val(.TextMatrix(lngCuenta, cintColVsfCapturaMinimo)) <> 0 Then
                                  'Si no se capturó la unidad del mínimo y se escribió mínimo:
28                                fblnDatosValidos = False
29                                intMensaje = 2
30                                vsfArticulos.Row = lngCuenta
31                                vsfArticulos.Col = cintColVsfCapturaUnidadMin
                              
32                            ElseIf (Val(.TextMatrix(lngCuenta, cintColVsfCapturaCveUnidadMin)) = -1 Or Val(.TextMatrix(lngCuenta, cintColVsfCapturaCveUnidadMin)) = 0) Then
33                                fblnDatosValidos = False
34                                intMensaje = 2
35                                vsfArticulos.Row = lngCuenta
36                                vsfArticulos.Col = cintColVsfCapturaUnidadMin
                                
37                            ElseIf Val(.TextMatrix(lngCuenta, cintColvsfCapturaTipoUniMax)) = Val(.TextMatrix(lngCuenta, cintColvsfCapturaTipoUniPun)) And _
                              Val(.TextMatrix(lngCuenta, cintColVsfCapturaMaximo)) <= Val(.TextMatrix(lngCuenta, cintColVsfCapturaPunto)) Then
                                  'Si el artículo se maneja en la misma unidad, pero es menor o igual el máximo que el punto de reorden
38                                fblnDatosValidos = False
39                                intMensaje = 26
40                                vsfArticulos.Row = lngCuenta
41                                vsfArticulos.Col = cintColVsfCapturaMaximo
                                 
42                            ElseIf Val(.TextMatrix(lngCuenta, cintColvsfCapturaTipoUniMax)) <> Val(.TextMatrix(lngCuenta, cintColvsfCapturaTipoUniPun)) And _
                                  Val(.TextMatrix(lngCuenta, cintColVsfCapturaPunto)) * IIf(Val(.TextMatrix(lngCuenta, cintColvsfCapturaTipoUniPun)) = 1, Val(.TextMatrix(lngCuenta, cintColVsfContenidoArt)), 1) >= _
                                  Val(.TextMatrix(lngCuenta, cintColVsfCapturaMaximo)) * IIf(Val(.TextMatrix(lngCuenta, cintColvsfCapturaTipoUniMax)) = 1, Val(.TextMatrix(lngCuenta, cintColVsfContenidoArt)), 1) Then
                                  'Si el artículo se maneja en unidades diferentes y es mayor el reorden que el máximo:
43                                fblnDatosValidos = False
44                                intMensaje = 26
45                                vsfArticulos.Row = lngCuenta
46                                vsfArticulos.Col = cintColVsfCapturaMaximo
                                
47                            ElseIf Val(.TextMatrix(lngCuenta, cintColvsfCapturaTipoUniPun)) = Val(.TextMatrix(lngCuenta, cintColvsfCapturaTipoUniMin)) And _
                              Val(.TextMatrix(lngCuenta, cintColVsfCapturaPunto)) < Val(.TextMatrix(lngCuenta, cintColVsfCapturaMinimo)) Then
                                  'Si el artículo se maneja en la misma unidad, pero es menor el reorden que el mínimo
48                                fblnDatosValidos = False
49                                intMensaje = 26
50                                vsfArticulos.Row = lngCuenta
51                                vsfArticulos.Col = cintColVsfCapturaPunto
                                 
52                            ElseIf Val(.TextMatrix(lngCuenta, cintColvsfCapturaTipoUniPun)) <> Val(.TextMatrix(lngCuenta, cintColvsfCapturaTipoUniMin)) And _
                                  Val(.TextMatrix(lngCuenta, cintColVsfCapturaMinimo)) * IIf(Val(.TextMatrix(lngCuenta, cintColvsfCapturaTipoUniMin)) = 1, Val(.TextMatrix(lngCuenta, cintColVsfContenidoArt)), 1) > _
                                  Val(.TextMatrix(lngCuenta, cintColVsfCapturaPunto)) * IIf(Val(.TextMatrix(lngCuenta, cintColvsfCapturaTipoUniPun)) = 1, Val(.TextMatrix(lngCuenta, cintColVsfContenidoArt)), 1) Then
                                  'Si el artículo se maneja en unidades diferentes, pero es menor el reorden que el mínimo
53                                fblnDatosValidos = False
54                                intMensaje = 26
55                                vsfArticulos.Row = lngCuenta
56                                vsfArticulos.Col = cintColVsfCapturaPunto
                                  
57                            End If
58                        End If
                     
                     ' se debe de validar que se cuente con max, min y punto de reorden para que se asigne al departamento
59                      End If
60               If fblnDatosValidos And Val(.TextMatrix(lngCuenta, cintColVsfCapturaMaximo)) <> 0 And Val(.TextMatrix(lngCuenta, cintColVsfCapturaMinimo)) <> 0 And Val(.TextMatrix(lngCuenta, cintColVsfCapturaCveUnidadPun)) <> 0 Then
61                          .TextMatrix(lngCuenta, 0) = "*"
62               Else ' si no cumple con el criterio de max,min y reorden, se debe desmarcar
63                          .TextMatrix(lngCuenta, 0) = ""
64               End If
                
                
                
65                    lngCuenta = lngCuenta + 1
66                Loop
67            End With
          
68            If Not fblnDatosValidos Then
                  ' 2  = ¡No ha ingresado datos!
                  ' 26 = ¡Rango incorrecto!
69                MsgBox SIHOMsg(intMensaje), vbOKOnly + vbExclamation, "Mensaje"
70                vsfArticulos.SetFocus
          
                  
71            End If
72        End If

73    Exit Function
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fblnDatosValidos" & " Linea:" & Erl()))
End Function

Private Sub pLlenaGridImportar()

On Error GoTo NotificaError
    Dim rs As ADODB.Recordset
    Dim rsConcepto As ADODB.Recordset
    Dim rsISR As ADODB.Recordset
    Dim rsiva As ADODB.Recordset
    Dim rsFormasPago As ADODB.Recordset
    Dim rsBancoSAT As ADODB.Recordset
    Dim X, Maximo, Minimo, Reorden, RenIniExcel As Long
    Dim CveMaximo, CveMinimo, CveReorden As Long
    Dim UniTipoMaximo, UniTipoMinimo, UniTipoReorden As Integer
    
    Dim vlCveAlmaDeptoReCo As String
    Dim vlstrIdArticulo As String
    Dim rowGrid As Integer
    Dim vlstrParametrosSP As String
    Dim vlstrsSQL As String
    Dim vlstrError As String
    Dim ContProduc As Long
    Dim BolValMaxMinReor As Boolean
    
    Dim vlstrAlmaDeptoReCo() As String
    
    
    'Para recorrer excel
    X = 6
    RenIniExcel = 0
    
    'Para grid
    rowGrid = 1
        
    'grdFacturas.Visible = False
    'vsfConcepto.Visible = False
    
    If oHoja.Cells(4, 1) = "" Or oHoja.Cells(X, 1) = "" Then
        vlstrError = "El documento está vacío o tiene un formato incorrecto."
        MsgBox vlstrError, vbExclamation + vbOKOnly, "Mensaje"
        pLimpiar
        Exit Sub
    End If
                                                           
    If Trim(oExcel.Cells(4, 1).Value) <> "DEPARTAMENTO " + Trim(CStr(cboDepartamento.List(cboDepartamento.ListIndex))) Then
         vlstrError = "La información no corresponde al departamento " + Trim(CStr(cboDepartamento.List(cboDepartamento.ListIndex))) + "."
         MsgBox vlstrError, vbExclamation + vbOKOnly, "Mensaje"
         pLimpiar
         Exit Sub
    End If
    
    lstrDeptoCompras = fAlmaComprasConf(Str(cboDepartamento.ItemData(cboDepartamento.ListIndex)), "CO")
    
     Do While oHoja.Cells(X, 1) <> "" Or oHoja.Cells(X, 2) <> ""
        
            'If X > 5 And Not IsNumeric(oHoja.Cells(X, 1)) Or Not IsNumeric(oHoja.Cells(X, 3)) Or Not IsNumeric(oHoja.Cells(X, 5)) Or Not IsNumeric(oHoja.Cells(X, 7)) Then
             '   vlstrError = "El documento tiene un formato incorrecto."
             ' GoTo NotificaError
            'End If
            
           If X > 6 Then
            
             'El identificador del artículo está vacío
                If (oHoja.Cells(X, 1) = "") Then
                   MsgBox "Renglón " + CStr(X - RenIniExcel) + ": El identificador del artículo está vacío", vbOKOnly + vbExclamation, "Mensaje"
                   pLimpiar
                        With vsfArticulos
                        .Redraw = True
                        .Visible = True
                        End With
                  Exit Sub
                End If
           
             If IsNumeric(oHoja.Cells(X, 1)) = True Then
             
               'Falta Maximo
                If (oHoja.Cells(X, 3) = "") Then
                   MsgBox "Renglón " + CStr(X - RenIniExcel) + ": Falta el Máximo", vbOKOnly + vbExclamation, "Mensaje"
                   pLimpiar
                        With vsfArticulos
                        .Redraw = True
                        .Visible = True
                        End With
                  Exit Sub
                End If
                
               'Falta reorden
                If (oHoja.Cells(X, 5) = "") Then
                   MsgBox "Renglón " + CStr(X - RenIniExcel) + ": Falta el Reorden", vbOKOnly + vbExclamation, "Mensaje"
                   pLimpiar
                        With vsfArticulos
                        .Redraw = True
                        .Visible = True
                        End With
                  Exit Sub
                End If
                
               'Falta Minimo
                If (oHoja.Cells(X, 7) = "") Then
                   MsgBox "Renglón " + CStr(X - RenIniExcel) + ": Falta el Mínimo", vbOKOnly + vbExclamation, "Mensaje"
                   pLimpiar
                        With vsfArticulos
                        .Redraw = True
                        .Visible = True
                        End With
                  Exit Sub
                End If
                
                'Maximo incorrecto
                If (IsNumeric(oHoja.Cells(X, 3)) = False) Then
                   MsgBox "Renglón " + CStr(X - RenIniExcel) + ": Máximo incorrecto", vbOKOnly + vbExclamation, "Mensaje"
                   pLimpiar
                        With vsfArticulos
                        .Redraw = True
                        .Visible = True
                        End With
                  Exit Sub
                End If
                
               'Reorden incorrecto
                If (IsNumeric(oHoja.Cells(X, 5)) = False) Then
                   MsgBox "Renglón " + CStr(X - RenIniExcel) + ": Reorden incorrecto", vbOKOnly + vbExclamation, "Mensaje"
                   pLimpiar
                        With vsfArticulos
                        .Redraw = True
                        .Visible = True
                        End With
                  Exit Sub
                End If
                
               'Minimo incorrecto
                If (IsNumeric(oHoja.Cells(X, 7)) = False) Then
                   MsgBox "Renglón " + CStr(X - RenIniExcel) + ": Mínimo incorrecto", vbOKOnly + vbExclamation, "Mensaje"
                   pLimpiar
                        With vsfArticulos
                        .Redraw = True
                        .Visible = True
                        End With
                  Exit Sub
                End If
                       
             
                Maximo = Val(oHoja.Cells(X, 3))
                Reorden = Val(oHoja.Cells(X, 5))
                Minimo = Val(oHoja.Cells(X, 7))
                
                lblnProductoExiste = False
                'Obtener el Id del Articulo
                vlstrIdArticulo = fCveArticuloMaxMin(CStr(oHoja.Cells(X, 1)))
                
                If vlstrIdArticulo <> "" Then
                    'Agregar articulo al grid
                    pAgreImportar (vlstrIdArticulo)
                End If
                              
               
                      'Identificador de artículo incorrecto
                      If lblnProductoExiste = False Then
                          MsgBox "Renglón " + CStr(X - RenIniExcel) + ": Identificador de artículo incorrecto", vbOKOnly + vbExclamation, "Mensaje"
                               pLimpiar
                              
                               With vsfArticulos
                               .Redraw = True
                               .Visible = True
                              End With
                          Exit Sub
                      End If
                
                 
               
                       If vsfArticulos.TextMatrix(rowGrid, 5) <> "" Then
                       
                           'vsfArticulos.TextMatrix(RowGrid, 5) = CStr(oHoja.Cells(X, 1))
                           'vsfArticulos.TextMatrix(RowGrid, 6) = IIf(IsNull(oHoja.Cells(X, 2)), "", CStr(oHoja.Cells(X, 2)))
                           ContProduc = vsfArticulos.TextMatrix(rowGrid, 21)
                           
                           If (ContProduc = 1) Then
                           
                           
                           'unidad cve maximo
                            If Trim(vsfArticulos.TextMatrix(rowGrid, 23)) = Trim(CStr(oHoja.Cells(X, 4))) Or Trim(vsfArticulos.TextMatrix(rowGrid, 25)) = Trim(CStr(oHoja.Cells(X, 4))) Or Trim(CStr(Trim(oHoja.Cells(X, 4)))) = "" Then
                                 
                                 If Trim(vsfArticulos.TextMatrix(rowGrid, 23)) = Trim(CStr(oHoja.Cells(X, 4))) Then
                                    CveMaximo = vsfArticulos.TextMatrix(rowGrid, 22)
                                    UniTipoMaximo = 1
                                Else
                                    CveMaximo = vsfArticulos.TextMatrix(rowGrid, 24)
                                    UniTipoMaximo = 0
                                End If
                             
                             Else
                                            
                                 MsgBox "Renglón " + CStr(X - RenIniExcel) + ": ¡Unidad del Máximo  incorrecto!", vbOKOnly + vbExclamation, "Mensaje"
                                 pLimpiar
                                     
                                 With vsfArticulos
                                 .Redraw = True
                                 .Visible = True
                                 End With
                                 
                                 Exit Sub
                               
                             End If
                             
                             
                              'unidad cve reorden
                            If Trim(vsfArticulos.TextMatrix(rowGrid, 23)) = Trim(CStr(oHoja.Cells(X, 6))) Or Trim(vsfArticulos.TextMatrix(rowGrid, 25)) = Trim(CStr(oHoja.Cells(X, 6))) Or Trim(CStr(Trim(oHoja.Cells(X, 6)))) = "" Then
                                 
                                 If Trim(vsfArticulos.TextMatrix(rowGrid, 23)) = Trim(CStr(oHoja.Cells(X, 6))) Then
                                    CveReorden = vsfArticulos.TextMatrix(rowGrid, 22)
                                    UniTipoReorden = 1
                                Else
                                    CveReorden = vsfArticulos.TextMatrix(rowGrid, 24)
                                    UniTipoReorden = 0
                                End If
                                
                             Else
                                            
                                 MsgBox "Renglón " + CStr(X - RenIniExcel) + ": ¡Unidad de Reorden incorrecto!", vbOKOnly + vbExclamation, "Mensaje"
                                 pLimpiar
                                     
                                 With vsfArticulos
                                 .Redraw = True
                                 .Visible = True
                                 End With
                                 
                                 Exit Sub
                               
                             End If
                             
                           
                           'unidad cve minimo
                         If Trim(vsfArticulos.TextMatrix(rowGrid, 23)) = Trim(CStr(oHoja.Cells(X, 8))) Or Trim(vsfArticulos.TextMatrix(rowGrid, 25)) = Trim(CStr(oHoja.Cells(X, 8))) Or Trim(CStr(Trim(oHoja.Cells(X, 8)))) = "" Then
                                 
                                 If vsfArticulos.TextMatrix(rowGrid, 23) = CStr(oHoja.Cells(X, 8)) Then
                                    Minimo = Minimo * ContProduc
                                     CveMinimo = vsfArticulos.TextMatrix(rowGrid, 22)
                                     UniTipoMinimo = 1
                                Else
                                    CveMinimo = vsfArticulos.TextMatrix(rowGrid, 24)
                                    UniTipoMinimo = 0
                                End If
                               
                             
                             Else
                                            
                                 MsgBox "Renglón " + CStr(X - RenIniExcel) + ": ¡Unidad de Mínimo incorrecto!", vbOKOnly + vbExclamation, "Mensaje"
                                 pLimpiar
                                     
                                 With vsfArticulos
                                 .Redraw = True
                                 .Visible = True
                                 End With
                                 
                                 Exit Sub
                               
                             End If
                                
                              'Validar Numeros
                             If Maximo = 0 And Minimo = 0 And Reorden = 0 Then
                                  BolValMaxMinReor = False
                             Else
                                  BolValMaxMinReor = True
                                 If Minimo > Reorden Or Minimo >= Maximo Then
                                     MsgBox "Renglón " + CStr(X - RenIniExcel) + ": ¡Rango incorrecto Mínimo!", vbOKOnly + vbExclamation, "Mensaje"
                                     pLimpiar
                                     
                                      With vsfArticulos
                                      .Redraw = True
                                      .Visible = True
                                     End With
                                     
                                 Exit Sub
                                End If
                                
                                 
                                If Reorden >= Maximo Or Reorden < Minimo Then
                                     MsgBox "Renglón " + CStr(X - RenIniExcel) + ": ¡Rango incorrecto Reorden!", vbOKOnly + vbExclamation, "Mensaje"
                                     pLimpiar
                                     
                                      With vsfArticulos
                                      .Redraw = True
                                      .Visible = True
                                     End With
                                     
                                 Exit Sub
                                End If
                        End If
                                    
                       ElseIf (ContProduc <> 1) Then
                                 
                                 
                                 '(RowGrid, 23) texto unidad Máximo  (RowGrid, 25) texto unidad Alterna
                                 '(X, 4) texto excel unidad Máximo
                             'Validar Maximo
                             If Trim(vsfArticulos.TextMatrix(rowGrid, 23)) = Trim(CStr(oHoja.Cells(X, 4))) Or Trim(vsfArticulos.TextMatrix(rowGrid, 25)) = Trim(CStr(oHoja.Cells(X, 4))) Or Trim(CStr(Trim(oHoja.Cells(X, 4)))) = "" Then
                                 
                                 If Trim(vsfArticulos.TextMatrix(rowGrid, 23)) = Trim(CStr(oHoja.Cells(X, 4))) Then
                                    Maximo = Maximo * ContProduc
                                    CveMaximo = vsfArticulos.TextMatrix(rowGrid, 22)
                                    UniTipoMaximo = 1
                                Else
                                    CveMaximo = vsfArticulos.TextMatrix(rowGrid, 24)
                                    UniTipoMaximo = 0
                                End If
                             
                             Else
                                            
                                 MsgBox "Renglón " + CStr(X - RenIniExcel) + ": ¡Unidad del Máximo  incorrecto!", vbOKOnly + vbExclamation, "Mensaje"
                                 pLimpiar
                                     
                                 With vsfArticulos
                                 .Redraw = True
                                 .Visible = True
                                 End With
                                 
                                 Exit Sub
                               
                             End If
                                 
                               
                                     
                              '(RowGrid, 23) texto unidad Máximo  (RowGrid, 25) texto unidad Alterna
                              '(X, 6) texto excel unidad Reorden
                              'Validar Reorden
                             If Trim(vsfArticulos.TextMatrix(rowGrid, 23)) = Trim(CStr(oHoja.Cells(X, 6))) Or Trim(vsfArticulos.TextMatrix(rowGrid, 25)) = Trim(CStr(oHoja.Cells(X, 6))) Or Trim(CStr(Trim(oHoja.Cells(X, 6)))) = "" Then
                                 
                                 If Trim(vsfArticulos.TextMatrix(rowGrid, 23)) = Trim(CStr(oHoja.Cells(X, 6))) Then
                                    Reorden = Reorden * ContProduc
                                    CveReorden = vsfArticulos.TextMatrix(rowGrid, 22)
                                    UniTipoReorden = 1
                                Else
                                    CveReorden = vsfArticulos.TextMatrix(rowGrid, 24)
                                    UniTipoReorden = 0
                                End If
                                
                             Else
                                            
                                 MsgBox "Renglón " + CStr(X - RenIniExcel) + ": ¡Unidad de Reorden incorrecto!", vbOKOnly + vbExclamation, "Mensaje"
                                 pLimpiar
                                     
                                 With vsfArticulos
                                 .Redraw = True
                                 .Visible = True
                                 End With
                                 
                                 Exit Sub
                               
                             End If
       
       
                              '(RowGrid, 23) texto unidad Máximo  (RowGrid, 25) texto unidad Alterna
                              '(X, 8) texto excel unidad Minimo
                              'Validar Minimo
                             If Trim(vsfArticulos.TextMatrix(rowGrid, 23)) = Trim(CStr(oHoja.Cells(X, 8))) Or Trim(vsfArticulos.TextMatrix(rowGrid, 25)) = Trim(CStr(oHoja.Cells(X, 8))) Or Trim(CStr(Trim(oHoja.Cells(X, 8)))) = "" Then
                                 
                                 If Trim(vsfArticulos.TextMatrix(rowGrid, 23)) = Trim(CStr(oHoja.Cells(X, 8))) Then
                                    Minimo = Minimo * ContProduc
                                     CveMinimo = vsfArticulos.TextMatrix(rowGrid, 22)
                                     UniTipoMinimo = 1
                                Else
                                    CveMinimo = vsfArticulos.TextMatrix(rowGrid, 24)
                                    UniTipoMinimo = 0
                                End If
                               
                             
                             Else
                                            
                                 MsgBox "Renglón " + CStr(X - RenIniExcel) + ": ¡Unidad de Mínimo incorrecto!", vbOKOnly + vbExclamation, "Mensaje"
                                 pLimpiar
                                     
                                 With vsfArticulos
                                 .Redraw = True
                                 .Visible = True
                                 End With
                                 
                                 Exit Sub
                               
                             End If
                            
                            
                        
                         If Maximo = 0 And Minimo = 0 And Reorden = 0 Then
                              BolValMaxMinReor = False
                         Else
                              BolValMaxMinReor = True
                             'Validar Numeros
                                 If Minimo > Reorden Or Minimo >= Maximo Then
                                     MsgBox "Renglón " + CStr(X - RenIniExcel) + ": ¡Rango incorrecto Mínimo!", vbOKOnly + vbExclamation, "Mensaje"
                                     pLimpiar
                                     
                                      With vsfArticulos
                                      .Redraw = True
                                      .Visible = True
                                     End With
                                     
                                 Exit Sub
                                End If
                                
                                 
                                If Reorden >= Maximo Or Reorden < Minimo Then
                                                 
                                     MsgBox "Renglón " + CStr(X - RenIniExcel) + ": ¡Rango incorrecto Reorden!", vbOKOnly + vbExclamation, "Mensaje"
                                     pLimpiar
                                     
                                      With vsfArticulos
                                      .Redraw = True
                                      .Visible = True
                                     End With
                                     
                                 Exit Sub
                                End If
                                
                        End If
                    End If
                                         
                           'Maximo
                           vsfArticulos.TextMatrix(rowGrid, 7) = CStr(oHoja.Cells(X, 3))
                           vsfArticulos.TextMatrix(rowGrid, 8) = IIf(IsNull(oHoja.Cells(X, 4)), "", CStr(oHoja.Cells(X, 4)))
                           vsfArticulos.TextMatrix(rowGrid, 9) = IIf(IsNull(oHoja.Cells(X, 4)), "", CveMaximo)
                           vsfArticulos.TextMatrix(rowGrid, 26) = UniTipoMaximo
                           
                           'Reorden
                           vsfArticulos.TextMatrix(rowGrid, 10) = CStr(oHoja.Cells(X, 5))
                           vsfArticulos.TextMatrix(rowGrid, 11) = IIf(IsNull(oHoja.Cells(X, 6)), "", CStr(oHoja.Cells(X, 6)))
                           vsfArticulos.TextMatrix(rowGrid, 12) = IIf(IsNull(oHoja.Cells(X, 6)), "", CveReorden)
                           vsfArticulos.TextMatrix(rowGrid, 27) = UniTipoReorden
                           
                           'Minimo
                           vsfArticulos.TextMatrix(rowGrid, 13) = CStr(oHoja.Cells(X, 7))
                           vsfArticulos.TextMatrix(rowGrid, 14) = IIf(IsNull(oHoja.Cells(X, 8)), "", CStr(oHoja.Cells(X, 8)))
                           vsfArticulos.TextMatrix(rowGrid, 15) = IIf(IsNull(oHoja.Cells(X, 8)), "", CveMinimo)
                           vsfArticulos.TextMatrix(rowGrid, 28) = UniTipoMinimo
                           
                           
                           If BolValMaxMinReor = True Then
                
                                'Validar Almacen reubica
                                 If Trim(oHoja.Cells(X, 9)) = "" Then
                                                              
                                    MsgBox "Renglón " + CStr(X - RenIniExcel) + ": Falta Almacén reubica", vbOKOnly + vbExclamation, "Mensaje"
                                    pLimpiar
                                    
                                    With vsfArticulos
                                    .Redraw = True
                                    .Visible = True
                                    End With
                                    Exit Sub
                                     
                                End If
                                                                       
                                'Departamento Consigna
                                If lbDeptoConSigna = True Then
                                    'validar que el Almacén reubica  este bien Depto consginación
                                    If CStr(Trim(oHoja.Cells(X, 9))) <> cboDepartamento.List(cboDepartamento.ListIndex) Then
                                         
                                         MsgBox "Renglón " + CStr(X - RenIniExcel) + ": ¡Almacén reubica incorrecto!", vbOKOnly + vbExclamation, "Mensaje"
                                         pLimpiar
                                     
                                         With vsfArticulos
                                         .Redraw = True
                                         .Visible = True
                                        End With
                                        Exit Sub
                                    
                                     Else
                                      vlCveAlmaDeptoReCo = "0,1"
                                    End If
                                    
                                Else
                                    vlCveAlmaDeptoReCo = fCveAlmaDeptoReCo(Str(cboDepartamento.ItemData(cboDepartamento.ListIndex)), IIf(IsNull(Trim(oHoja.Cells(X, 9))), "", CStr(Trim(oHoja.Cells(X, 9)))), "RE")
                                End If
                                
                              
                                If vlCveAlmaDeptoReCo = "0" Then
                                           
                                     MsgBox "Renglón " + CStr(X - RenIniExcel) + ": ¡Almacén reubica incorrecto!", vbOKOnly + vbExclamation, "Mensaje"
                                     pLimpiar
                                     
                                      With vsfArticulos
                                      .Redraw = True
                                      .Visible = True
                                     End With
                                     Exit Sub
                                     
                                   vsfArticulos.TextMatrix(rowGrid, cintColVsfCapturaCveAlmacen) = 0
                                   vsfArticulos.TextMatrix(rowGrid, cintColVsfCapturaAlmacen) = ""
                                Else
                                    
                                    
                                     If lbDeptoConSigna = True Then
                                        vsfArticulos.TextMatrix(rowGrid, cintColVsfCapturaCveAlmacen) = "0"
                                        vsfArticulos.TextMatrix(rowGrid, cintColVsfCapturaAlmacen) = CStr(oHoja.Cells(X, 9))
                                     Else
                                        vlstrAlmaDeptoReCo = Split(vlCveAlmaDeptoReCo, ",")
                                        vlCveAlmaDeptoReCo = vlstrAlmaDeptoReCo(0)
                                        
                                        vsfArticulos.TextMatrix(rowGrid, cintColVsfCapturaCveAlmacen) = vlCveAlmaDeptoReCo
                                        vsfArticulos.TextMatrix(rowGrid, cintColVsfCapturaAlmacen) = CStr(oHoja.Cells(X, 9))
                                                                                  
                                        lbConSigna = False
                                        If Trim(vlstrAlmaDeptoReCo(1)) = "1" Then
                                         lbConSigna = True
                                        End If
                                     End If
                                   
                                End If
                                
                                
                                'Validar  Departamento compras lstrDeptoCompras S esta configurado N no hay departamento
                                If lstrDeptoCompras = "S" Then
                                                                   
                                     If Trim(oHoja.Cells(X, 10)) = "" And lbConSigna = False Then
                                                                  
                                        MsgBox "Renglón " + CStr(X - RenIniExcel) + ": Falta Departamento compras", vbOKOnly + vbExclamation, "Mensaje"
                                        pLimpiar
                                        
                                        With vsfArticulos
                                        .Redraw = True
                                        .Visible = True
                                        End With
                                        Exit Sub
                                    Else
                                                vsfArticulos.TextMatrix(rowGrid, cintColVsfCapturaCveDeptoCompra) = 0
                                                vsfArticulos.TextMatrix(rowGrid, cintColVsfCapturaDeptoCompra) = ""
                                         
                                    End If
                                    
                                      If Trim(oHoja.Cells(X, 10)) <> "" Then
                                             vlCveAlmaDeptoReCo = fCveAlmaDeptoReCo(Str(cboDepartamento.ItemData(cboDepartamento.ListIndex)), IIf(IsNull(Trim(oHoja.Cells(X, 10))), "", CStr(Trim(oHoja.Cells(X, 10)))), "CO")
                                               
                                              If vlCveAlmaDeptoReCo = "0" Then
                                                     
                                                  MsgBox "Renglón " + CStr(X - RenIniExcel) + ": ¡Departamento compras incorrecto!", vbOKOnly + vbExclamation, "Mensaje"
                                                  pLimpiar
                                                  
                                                   With vsfArticulos
                                                   .Redraw = True
                                                   .Visible = True
                                                  End With
                                                 Exit Sub
                                             
                                                vsfArticulos.TextMatrix(rowGrid, cintColVsfCapturaCveDeptoCompra) = 0
                                                vsfArticulos.TextMatrix(rowGrid, cintColVsfCapturaDeptoCompra) = ""
                                             Else
                                             
                                                 vlstrAlmaDeptoReCo = Split(vlCveAlmaDeptoReCo, ",")
                                                 vlCveAlmaDeptoReCo = vlstrAlmaDeptoReCo(0)
                                   
                                                vsfArticulos.TextMatrix(rowGrid, cintColVsfCapturaCveDeptoCompra) = vlCveAlmaDeptoReCo
                                                vsfArticulos.TextMatrix(rowGrid, cintColVsfCapturaDeptoCompra) = CStr(oHoja.Cells(X, 10))
                                             End If
                                    End If
                                Else
                                 
                                 'Validar Departamento compras no esta configurado
                                 'El departamento que realiza la compra no ha sido configurado
                                     If Trim(oHoja.Cells(X, 10)) <> "" Then
                                                                  
                                        MsgBox "Renglón " + CStr(X - RenIniExcel) + ": El departamento que realiza la compra no ha sido configurado ", vbOKOnly + vbExclamation, "Mensaje"
                                        pLimpiar
                                        
                                        With vsfArticulos
                                        .Redraw = True
                                        .Visible = True
                                        End With
                                        Exit Sub
                                         
                                      End If
                                      
                                       vsfArticulos.TextMatrix(rowGrid, cintColVsfCapturaCveDeptoCompra) = 0
                                       vsfArticulos.TextMatrix(rowGrid, cintColVsfCapturaDeptoCompra) = ""
                                
                                End If
                                
                             
                                'Habilitar para guardar
                                vsfArticulos.TextMatrix(rowGrid, 29) = 1
                           Else
                               'No guardar Maximo = Minimo = Reorden = 0
                                vsfArticulos.TextMatrix(rowGrid, 29) = -1
                                                       
                               'Máximo:
                                   vsfArticulos.TextMatrix(rowGrid, cintColVsfCapturaCveUnidadMax) = 0
                                   vsfArticulos.TextMatrix(rowGrid, cintColvsfCapturaTipoUniMax) = ""
                              
                               'Punto reorden:
                                   vsfArticulos.TextMatrix(rowGrid, cintColVsfCapturaCveUnidadPun) = 0
                                   vsfArticulos.TextMatrix(rowGrid, cintColvsfCapturaTipoUniPun) = ""
                            
                               'Mínimo:
                                   vsfArticulos.TextMatrix(rowGrid, cintColVsfCapturaCveUnidadMin) = 0
                                   vsfArticulos.TextMatrix(rowGrid, cintColvsfCapturaTipoUniMin) = ""
                              
                               'Almacén reubica:
                                   vsfArticulos.TextMatrix(rowGrid, cintColVsfCapturaCveAlmacen) = 0
                                   vsfArticulos.TextMatrix(rowGrid, cintColVsfCapturaAlmacen) = ""
                             
                               'Departamento compra:
                                   vsfArticulos.TextMatrix(rowGrid, cintColVsfCapturaCveDeptoCompra) = 0
                                   vsfArticulos.TextMatrix(rowGrid, cintColVsfCapturaDeptoCompra) = ""
                                                  
                           End If
                         
                           
                            rowGrid = rowGrid + 1
                           
                            With vsfArticulos
                            '.Rows = RowGrid + 1
                             .Redraw = True
                             .Visible = True
                            End With
                    
                      End If
                 
              Else
               
                    MsgBox "Renglón " + CStr(X - RenIniExcel) + ": El identificador del artículo no es válido", vbOKOnly + vbExclamation, "Mensaje"
                         pLimpiar
                        
                         With vsfArticulos
                         .Redraw = True
                         .Visible = True
                        End With
                    Exit Sub
               
                
              End If
          End If
          
          X = X + 1
    Loop
   
    'Call grdFacturas_Click
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLlenaGridImportar"))
    MsgBox vlstrError, vbExclamation + vbOKOnly, "Mensaje"
    'pLimpia
End Sub
Private Sub cmdImportar_Click(Index As Integer)
    CommonDialog1.Filter = "Excel (*.xlsx)|*.xlsx|Todos los archivos (*.*)|*.*"
    CommonDialog1.DefaultExt = "xlsx"
    CommonDialog1.DialogTitle = "Seleccionar archivo"
    CommonDialog1.ShowOpen
    'True Then
    If CommonDialog1.FileName <> "" Then
        Set oExcel = CreateObject("Excel.Application")
        Set oLibro = oExcel.Workbooks.Open(FileName:=CommonDialog1.FileName)
        CommonDialog1.FileName = ""
        Set oHoja = oLibro.Worksheets(1)
        cmdImportar(0).Enabled = False
        Call pLlenaGridImportar
        
        If vsfArticulos.TextMatrix(1, 5) <> "" Then
           cmdGrabar.Enabled = True
           cmdExportar(2).Enabled = True
        Else
           cmdGrabar.Enabled = False
           cmdExportar(2).Enabled = False
        End If
        
        
    End If
    
     If Not oLibro Is Nothing Then
        oLibro.Close
        oExcel.Quit
        Set oExcel = Nothing
        Set oLibro = Nothing
        Set oHoja = Nothing
    End If
    
End Sub

Private Sub cmdImprimir_Click()
On Error GoTo NotificaError
    Dim i As Integer
    'Imprime requisiciones en caso de seleccionar imprimir localmente
    frmBotonera.Enabled = False
    For i = 1 To UBound(aRequisiciones(), 1)
        pImprimeLocal aRequisiciones(i).lgnNumRequisicion
    Next i
    frmBotonera.Enabled = True
    cmdImprimir.SetFocus
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdImprimir_Click"))
End Sub

Private Sub cmdInvertirSeleccion_Click()
On Error GoTo NotificaError
Dim lngContador As Long
    
    For lngContador = 1 To vsfArticulos.Rows - 1
        If Trim(vsfArticulos.TextMatrix(lngContador, cintColvsfCveArt)) = "*" Then
            vsfArticulos.Row = lngContador
            pMarca 0
        Else
            vsfArticulos.Row = lngContador
            pMarca 1
        End If
    Next lngContador

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdInvertirSeleccion_Click"))
End Sub

Private Sub cmdManejos_Click()
    frmManejoMedicamentos.blnSoloBusqueda = True
    frmManejoMedicamentos.Show vbModal, Me
End Sub

Private Sub cmdMarcar_Click()
On Error GoTo NotificaError

    If Trim(vsfArticulos.TextMatrix(vsfArticulos.Row, cintColvsfCveArt)) = "*" Then
        'Significa que el artículo no está ubicado:
        pMarca 0
    Else
        'El artículo ya está asignado a este departamento.
        'MsgBox SIHOMsg(872), vbOKOnly + vbInformation, "Mensaje"
        pMarca 1
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdMarcar_Click"))
End Sub

Private Sub pMarca(intTipo As Integer)
On Error GoTo NotificaError

    vsfArticulos.TextMatrix(vsfArticulos.Row, 0) = IIf(Trim(vsfArticulos.TextMatrix(vsfArticulos.Row, 0)) = "", "*", "")
    
    If vsfArticulos.TextMatrix(vsfArticulos.Row, 0) = "*" Then
        If intTipo = 0 Then
            llngTotalMarcados = llngTotalMarcados + 1
        Else
            llngTotalDesactivar = llngTotalDesactivar + 1
        End If
    Else
        If intTipo = 0 Then
            llngTotalMarcados = llngTotalMarcados - 1
        Else
            llngTotalDesactivar = llngTotalDesactivar - 1
        End If
    End If
    
    cmdAsignarDesactivar(0).Enabled = llngTotalMarcados > 0
    cmdAsignarDesactivar(1).Enabled = llngTotalDesactivar > 0

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pMarca"))
End Sub

Private Sub cmdRequisitar_Click()
1         On Error GoTo NotificaError
          Dim lngContador As Long
          Dim lngResiduo As Long
          Dim lngSinFaltante As Long
          Dim lngCantidadFaltante As Long
          Dim strArticulosControlados As String
          Dim strArticulosDepto As String
          Dim lngPendientes As Long
          Dim intControlado As Integer
          Dim lngCveDeptoSurte As Long
          Dim strNombreDeptoSurte As String
          Dim rsRequisicionCompra As New ADODB.Recordset
          Dim rsConSigna As ADODB.Recordset
          Dim vgstrParametrosSP As String
          Dim intCantidad As Integer
          Dim llngCveDeptoAutoriza As Long
          Dim llngCveDeptoRecibe As Long
          Dim lblnContinuar As Boolean
          Dim strListaArticulosNoIncluidos As String
          Dim rsValor As New ADODB.Recordset
          Dim strContinuarSubrogado As String
          Dim rsSubrogado As New ADODB.Recordset

          'Inicialización de variables
2         strListaArticulosNoIncluidos = ""

3         vgstrParametrosSP = 1 & "|" & vgintClaveEmpresaContable & "|" & CStr(cboDepartamento.ItemData(cboDepartamento.ListIndex))
4         Set rsConSigna = frsEjecuta_SP(vgstrParametrosSP, "SP_IVSELALMACENCONSIGNA")

5         pLimpiarVsfBusArticulos
6         VsfBusArticulos.Redraw = False
          
7         If rsConSigna.RecordCount > 0 Then
8             VsfBusArticulos.ColWidth(cintColBusAlmacen) = 0
9         End If

10        llngTotalRequerir = 0
11        lngSinFaltante = 0
12        strArticulosControlados = ""

13        lblnContinuar = True
14        If lstrTipoRequisicion = "C" Then
              'Departamento que autoriza y recibe la compra de los artículos:
15            Set rsRequisicionCompra = frsEjecuta_SP(Str(vglngNumeroLogin), "SP_IVSELREQUISDEPARTAMENTOCOMP")
16            If rsRequisicionCompra.RecordCount <> 0 Then
17                llngCveDeptoAutoriza = IIf(IsNull(rsRequisicionCompra!intCveDeptoAutoriza), 0, rsRequisicionCompra!intCveDeptoAutoriza)
18                llngCveDeptoRecibe = IIf(IsNull(rsRequisicionCompra!intCveDeptoRecibe), 0, rsRequisicionCompra!intCveDeptoRecibe)
19            Else
20                lblnContinuar = False
21                MsgBox SIHOMsg(869), vbOKOnly + vbInformation, "Mensaje"
22            End If
23            rsRequisicionCompra.Close
24        End If

25        If lblnContinuar Then
26            ldblAumento = 100 / (VsfBusArticulos.Rows - 1)

27            pgbBarra.Value = 0
28            fraBarra.Visible = True
29            lblTituloBarra.Caption = "Calculando cantidades faltantes, espere un momento"
30            strContinuarSubrogado = ""
31            For lngContador = 1 To vsfArticulos.Rows - 1
                  '*- Barra de avance
32                fraBarra.Refresh
33                pgbBarra.Value = IIf(pgbBarra.Value + ldblAumento > 100, 100, pgbBarra.Value + ldblAumento)
                  
                  '*- Barra de avance
34                If Val(vsfArticulos.TextMatrix(lngContador, cintColvsfRequisitar)) = 1 Then
                      'Quiere decir que está en REORDEN, FALTANTE O INSUFICIENTE
                      
35                    lngCantidadFaltante = 1
36                    If strClavesConsol <> "" Then
                          'vgstrParametrosSP = vsfArticulos.TextMatrix(lngContador, cintColVsfClaveArt) & "|" & Str(cboDepartamento.ItemData(cboDepartamento.ListIndex)) & "|" & _
                                              Str (cboDepartamento.ItemData(cboDepartamento.ListIndex)) & "_" & vsfArticulos.TextMatrix(lngContador, cintColVsfArtiDeptoEsclavos)
                          'Set rs = frsEjecuta_SP(vgstrParametrosSP, "Sp_IvintcantfaltaConsol", True, lngCantidadFaltante)
                          lngCantidadFaltante = Val(vsfArticulos.TextMatrix(lngContador, cintColVsfCapturaMaximo) - vsfArticulos.TextMatrix(lngContador, cintColVsfExistencia)) * Val(vsfArticulos.TextMatrix(lngContador, cintColVsfContenidoArt))
                      Else
                        vgstrParametrosSP = vsfArticulos.TextMatrix(lngContador, cintColVsfClaveArt) & "|" & Str(cboDepartamento.ItemData(cboDepartamento.ListIndex))
37                        Set rs = frsEjecuta_SP(vgstrParametrosSP, "SP_IVINTCANTIDADFALTANTE", True, lngCantidadFaltante)
                      End If
38                    lngPendientes = flngCantidadPendiente(vsfArticulos.TextMatrix(lngContador, cintColVsfClaveArt), cboDepartamento.ItemData(cboDepartamento.ListIndex))

39                    If lngCantidadFaltante - lngPendientes <> 0 Then
                          'Si existe faltante (Esta cantidad es en unidades mínimas para artículos que se manejan en dos unidades
                          'para los otros es la cantidad de alterna) y tiene asignado almacén que surte:

                          'Esto creo que está mal, estaban pasando articulos sin almacen que surtirá para compra pedido y reubicaciones
      '                    If (Trim(vsfArticulos.TextMatrix(lngContador, cintColVsfCapturaAlmacen)) <> "" Or rsConSigna.RecordCount > 0 And _
      '                        cboTipoRequisicion.ItemData(cboTipoRequisicion.ListIndex) = cintItemDataReubicacion) Or _
      '                        (Trim(vsfArticulos.TextMatrix(lngContador, cintColVsfCapturaDeptoCompra)) <> "" And _
      '                        cboTipoRequisicion.ItemData(cboTipoRequisicion.ListIndex) = cintItemDataCompra) Then
                              
40                        If ((Trim(vsfArticulos.TextMatrix(lngContador, cintColVsfCapturaAlmacen)) <> "" And _
                              cboTipoRequisicion.ItemData(cboTipoRequisicion.ListIndex) = cintItemDataReubicacion) Or rsConSigna.RecordCount > 0) Or _
                              (Trim(vsfArticulos.TextMatrix(lngContador, cintColVsfCapturaDeptoCompra)) <> "" And _
                              cboTipoRequisicion.ItemData(cboTipoRequisicion.ListIndex) = cintItemDataCompra) Then

                              'Se agregó esta validación para omitir los artículos que generan requisiciones de compra pedido por cantidades en 0's cuando es una requisición de tipo COMPRA
41                            If (lstrTipoRequisicion = "C" And Int((lngCantidadFaltante - lngPendientes) / Val(vsfArticulos.TextMatrix(lngContador, cintColVsfContenidoArt))) > 0) Or lstrTipoRequisicion = "R" Then
42                                intControlado = Val(vsfArticulos.TextMatrix(lngContador, cintColvsfControlado))
43                                lngCveDeptoSurte = Val(IIf(cboTipoRequisicion.ItemData(cboTipoRequisicion.ListIndex) = cintItemDataReubicacion, vsfArticulos.TextMatrix(lngContador, cintColVsfCapturaCveAlmacen), vsfArticulos.TextMatrix(lngContador, cintColVsfCapturaCveDeptoCompra)))
44                                strNombreDeptoSurte = IIf(cboTipoRequisicion.ItemData(cboTipoRequisicion.ListIndex) = cintItemDataReubicacion, vsfArticulos.TextMatrix(lngContador, cintColVsfCapturaAlmacen), vsfArticulos.TextMatrix(lngContador, cintColVsfCapturaDeptoCompra))

                                  'Si es medicamento controlado y se tiene permiso o no es controlado validar el tipo de artículo por tipo de requisición:
45                                If fblnPermisoControlado(intControlado, lngCveDeptoSurte, lstrTipoRequisicion) Then
                                      'validacion subrogados
                                      
46                                    If lintInterfazFarmaciaSJP = 1 And lstrTipoRequisicion = "R" Then
47                                        If llngDeptoSubrogado = lngCveDeptoSurte Then
48                                            Set rsSubrogado = frsRegresaRs("select count(*) total from IvArticulo inner join IvArticulosSubrogados on IVARTICULO.INTIDARTICULO = IVARTICULOSSUBROGADOS.INTIDARTICULO " & _
                                                              " where IVARTICULOSSUBROGADOS.VCHCVEARTICULOEXT is not null and IVARTICULO.CHRCVEARTICULO = '" & Trim(vsfArticulos.TextMatrix(lngContador, cintColVsfClaveArt)) & "'")
49                                            If rsSubrogado!Total = 0 Then
50                                                strContinuarSubrogado = strContinuarSubrogado & Trim(vsfArticulos.TextMatrix(lngContador, cintColVsfClaveArt)) & "  " & Trim(vsfArticulos.TextMatrix(lngContador, cintColVsfNombreComercial)) & Chr(13)
51                                            End If
52                                        End If
53                                    End If
54                                    llngTotalRequerir = llngTotalRequerir + 1
55                                    VsfBusArticulos.TextMatrix(VsfBusArticulos.Rows - 1, cintColBusClave) = vsfArticulos.TextMatrix(lngContador, cintColVsfClaveArt)
56                                    VsfBusArticulos.TextMatrix(VsfBusArticulos.Rows - 1, cintColBusNombre) = vsfArticulos.TextMatrix(lngContador, cintColVsfNombreComercial)
57                                    VsfBusArticulos.TextMatrix(VsfBusArticulos.Rows - 1, cintColBusLocalizacion) = vsfArticulos.TextMatrix(lngContador, cintColvsfLocalizacion)

                                      ' Valida renglon de consignación
58                                    If rsConSigna.RecordCount = 0 Then
59                                        VsfBusArticulos.TextMatrix(VsfBusArticulos.Rows - 1, cintColBusAlmacen) = strNombreDeptoSurte
60                                    End If

61                                    VsfBusArticulos.TextMatrix(VsfBusArticulos.Rows - 1, cintColBusIdAlmacen) = lngCveDeptoSurte
62                                    VsfBusArticulos.TextMatrix(VsfBusArticulos.Rows - 1, cintColBusExist) = vsfArticulos.TextMatrix(lngContador, cintColVsfExistencia)
63                                    VsfBusArticulos.TextMatrix(VsfBusArticulos.Rows - 1, cintColBusUnidadExist) = vsfArticulos.TextMatrix(lngContador, cintColVsfExistenciaUnidad)
64                                    VsfBusArticulos.TextMatrix(VsfBusArticulos.Rows - 1, cintColBusPendiente) = flngCantidadPendiente(vsfArticulos.TextMatrix(lngContador, cintColVsfClaveArt), cboDepartamento.ItemData(cboDepartamento.ListIndex))

65                                    If Val(vsfArticulos.TextMatrix(lngContador, cintColVsfContenidoArt)) = 1 Then
                                          'El articulo se maneja en una sola unidad:
66                                        VsfBusArticulos.TextMatrix(VsfBusArticulos.Rows - 1, cintColBusUnidadPendiente) = IIf(Val(VsfBusArticulos.TextMatrix(VsfBusArticulos.Rows - 1, cintColBusPendiente)) = 0, " ", vsfArticulos.TextMatrix(lngContador, cintColvsfUnidadAlternaArt))
67                                        VsfBusArticulos.TextMatrix(VsfBusArticulos.Rows - 1, cintColBusCantidad) = lngCantidadFaltante - lngPendientes
68                                        VsfBusArticulos.TextMatrix(VsfBusArticulos.Rows - 1, cintColBusUnidad) = vsfArticulos.TextMatrix(lngContador, cintColvsfUnidadAlternaArt)
69                                        VsfBusArticulos.TextMatrix(VsfBusArticulos.Rows - 1, cintColBusTipoUnidad) = "A"
70                                        VsfBusArticulos.TextMatrix(VsfBusArticulos.Rows - 1, cintColBusCantidadOriginal) = lngCantidadFaltante - lngPendientes
71                                    Else
                                          'El articulo se maneja en dos unidades:
72                                        VsfBusArticulos.TextMatrix(VsfBusArticulos.Rows - 1, cintColBusUnidadPendiente) = IIf(Val(VsfBusArticulos.TextMatrix(VsfBusArticulos.Rows - 1, cintColBusPendiente)) = 0, " ", vsfArticulos.TextMatrix(lngContador, cintColvsfUnidadMinimaArt))
73                                        lngResiduo = (lngCantidadFaltante - lngPendientes) Mod Val(vsfArticulos.TextMatrix(lngContador, cintColVsfContenidoArt))

74                                        If lngResiduo = 0 Then
                                              'Se completan unidades alternas, se pide en alterna:
75                                            VsfBusArticulos.TextMatrix(VsfBusArticulos.Rows - 1, cintColBusCantidad) = (lngCantidadFaltante - lngPendientes) / Val(vsfArticulos.TextMatrix(lngContador, cintColVsfContenidoArt))
76                                            VsfBusArticulos.TextMatrix(VsfBusArticulos.Rows - 1, cintColBusUnidad) = vsfArticulos.TextMatrix(lngContador, cintColvsfUnidadAlternaArt)
77                                            VsfBusArticulos.TextMatrix(VsfBusArticulos.Rows - 1, cintColBusTipoUnidad) = "A"
78                                            VsfBusArticulos.TextMatrix(VsfBusArticulos.Rows - 1, cintColBusCantidadOriginal) = (lngCantidadFaltante - lngPendientes) / Val(vsfArticulos.TextMatrix(lngContador, cintColVsfContenidoArt))
79                                        Else
80                                            If lstrTipoRequisicion = "R" Then
81                                                If rsConSigna.RecordCount = 0 Then
                                                      'REUBICACION:
                                                      'Se pide en unidad mínima, la cantidad exacta para alcanzar el máximo:
82                                                    VsfBusArticulos.TextMatrix(VsfBusArticulos.Rows - 1, cintColBusCantidad) = lngCantidadFaltante - lngPendientes
83                                                    VsfBusArticulos.TextMatrix(VsfBusArticulos.Rows - 1, cintColBusUnidad) = vsfArticulos.TextMatrix(lngContador, cintColvsfUnidadMinimaArt)
84                                                    VsfBusArticulos.TextMatrix(VsfBusArticulos.Rows - 1, cintColBusTipoUnidad) = "M"
85                                                    VsfBusArticulos.TextMatrix(VsfBusArticulos.Rows - 1, cintColBusCantidadOriginal) = lngCantidadFaltante - lngPendientes
86                                                Else
                                                      'COONSIGNACIÓN:
                                                      'Se pide en unidad alterna, la cantidad más cercana para alcanzar el máximo:
87                                                    intCantidad = Int((lngCantidadFaltante - lngPendientes) / Val(vsfArticulos.TextMatrix(lngContador, cintColVsfContenidoArt)))
88                                                    If intCantidad > 0 Then
89                                                        VsfBusArticulos.TextMatrix(VsfBusArticulos.Rows - 1, cintColBusCantidad) = Int((lngCantidadFaltante - lngPendientes) / Val(vsfArticulos.TextMatrix(lngContador, cintColVsfContenidoArt)))
90                                                        VsfBusArticulos.TextMatrix(VsfBusArticulos.Rows - 1, cintColBusCantidadOriginal) = Int((lngCantidadFaltante - lngPendientes) / Val(vsfArticulos.TextMatrix(lngContador, cintColVsfContenidoArt)))
91                                                    Else
92                                                        VsfBusArticulos.TextMatrix(VsfBusArticulos.Rows - 1, cintColBusCantidad) = Int((lngCantidadFaltante - lngPendientes)) '/ Val(vsfArticulos.TextMatrix(lngContador, cintColVsfContenidoArt)))
93                                                        VsfBusArticulos.TextMatrix(VsfBusArticulos.Rows - 1, cintColBusCantidadOriginal) = Int((lngCantidadFaltante - lngPendientes))
94                                                    End If

95                                                    VsfBusArticulos.TextMatrix(VsfBusArticulos.Rows - 1, cintColBusUnidad) = vsfArticulos.TextMatrix(lngContador, cintColvsfUnidadAlternaArt)
96                                                    VsfBusArticulos.TextMatrix(VsfBusArticulos.Rows - 1, cintColBusTipoUnidad) = "A"
97                                                End If
98                                            Else
                                                  'COMPRA - PEDIDO:
                                                  'Se pide en unidad alterna, la cantidad más cercana para alcanzar el máximo:
99                                                VsfBusArticulos.TextMatrix(VsfBusArticulos.Rows - 1, cintColBusCantidad) = Int((lngCantidadFaltante - lngPendientes) / Val(vsfArticulos.TextMatrix(lngContador, cintColVsfContenidoArt)))
100                                               VsfBusArticulos.TextMatrix(VsfBusArticulos.Rows - 1, cintColBusUnidad) = vsfArticulos.TextMatrix(lngContador, cintColvsfUnidadAlternaArt)
101                                               VsfBusArticulos.TextMatrix(VsfBusArticulos.Rows - 1, cintColBusTipoUnidad) = "A"
102                                               VsfBusArticulos.TextMatrix(VsfBusArticulos.Rows - 1, cintColBusCantidadOriginal) = Int((lngCantidadFaltante - lngPendientes) / Val(vsfArticulos.TextMatrix(lngContador, cintColVsfContenidoArt)))
103                                           End If
104                                       End If
105                                   End If

106                                   If lstrTipoRequisicion = "C" Then
                                          'Departamento que autoriza y recibe la compra de los artículos:
107                                       VsfBusArticulos.TextMatrix(VsfBusArticulos.Rows - 1, cintColBusCveDeptoAutoriza) = llngCveDeptoAutoriza
108                                       VsfBusArticulos.TextMatrix(VsfBusArticulos.Rows - 1, cintColBusCveDeptoRecibe) = llngCveDeptoRecibe
                                          If strClavesConsol <> "" Then
                                             VsfBusArticulos.TextMatrix(VsfBusArticulos.Rows - 1, cintColBusExisPrin) = vsfArticulos.TextMatrix(lngContador, cintColVsfExistencia)
                                             VsfBusArticulos.TextMatrix(VsfBusArticulos.Rows - 1, cintColBusMaxPrin) = vsfArticulos.TextMatrix(lngContador, cintColVsfCapturaMaximo)
                                                       
                                             VsfBusArticulos.TextMatrix(VsfBusArticulos.Rows - 1, cintColBusExisEsclavos) = vsfArticulos.TextMatrix(lngContador, cintColVsfArtiExisEsclavos)
                                             VsfBusArticulos.TextMatrix(VsfBusArticulos.Rows - 1, cintColBusMaxEsclavos) = vsfArticulos.TextMatrix(lngContador, cintColVsfArtiMaxEsclavos)
                                             VsfBusArticulos.TextMatrix(VsfBusArticulos.Rows - 1, cintColBusDeptoEsclavos) = vsfArticulos.TextMatrix(lngContador, cintColVsfArtiDeptoEsclavos)
                                          End If
109                                   End If
110                                   VsfBusArticulos.Rows = VsfBusArticulos.Rows + 1
111                               Else
                                      'Concatenar los artículos controlados:
112                                   strArticulosControlados = strArticulosControlados & vsfArticulos.TextMatrix(lngContador, cintColVsfNombreComercial) & Chr(13)
113                               End If
114                           Else
                                  'Listar los artículos que no se agregarán en la requisición automática de faltantes
115                               strListaArticulosNoIncluidos = strListaArticulosNoIncluidos & Trim(vsfArticulos.TextMatrix(lngContador, 20)) & " - " & Trim(vsfArticulos.TextMatrix(lngContador, 6)) & vbNewLine
116                           End If
117                       Else
                              'Concatenar los artículos que no tienen departamento asignado:
118                           strArticulosDepto = strArticulosDepto & vsfArticulos.TextMatrix(lngContador, cintColVsfNombreComercial) & Chr(13)
119                       End If
120                   Else
121                       lngSinFaltante = lngSinFaltante + 1
122                   End If
123               End If
124           Next lngContador
125           fraBarra.Visible = False
126       End If

127       pConfVsfBusArticulos
128       VsfBusArticulos.Redraw = True

129       If Trim(strArticulosDepto) <> "" Then
130           If cboTipoRequisicion.ItemData(cboTipoRequisicion.ListIndex) = cintItemDataReubicacion Then
131               If rsConSigna.RecordCount = 0 Then
132                   If Not lbFlag = True Then
                          'Existen artículos que no tienen configurado almacén que surte.
133                       MsgBox SIHOMsg(776) & Chr(13) & strArticulosDepto, vbOKOnly + vbInformation, "Mensaje"
134                   End If
135               End If
136               lbFlag = False
137           Else
                  'Existen artículos que no tienen configurado departamento que compra.
138               cboTipoRequisicion.ItemData(cboTipoRequisicion.ListIndex) = cintItemDataCompra
139               MsgBox SIHOMsg(960) & Chr(13) & strArticulosDepto, vbOKOnly + vbInformation, "Mensaje"
140           End If
141       End If
          
142       If Trim(strArticulosControlados) <> "" Then
              'No está autorizado para solicitar medicamento controlado.
143           MsgBox SIHOMsg(585) & Chr(13) & strArticulosControlados, vbOKOnly + vbExclamation, "Mensaje"
144       End If

145       If Trim(strListaArticulosNoIncluidos) <> "" And lstrTipoRequisicion = "C" Then
              'Mensaje que lista los artículos que no se incluirán en la requisición (debido a que generarían una requisisición por cantidad = 0)
146           MsgBox "Los siguientes artículos no serán incluidos ya que generan requisiciones con cantidad cero:" & vbNewLine & vbNewLine & strListaArticulosNoIncluidos, vbInformation, "Mensaje"
147       End If
          
148       If lstrTipoRequisicion = "R" Then
              ' si es el almacen subrogado que tenca codigo externo
149           If strContinuarSubrogado <> "" Then
150               MsgBox "Los siguientes artículos no tienen clave asignada de la farmacia subrogada. " & Chr(13) & "No es posible realizar la requisición. " & Chr(13) & Chr(13) & strContinuarSubrogado, vbCritical, "Mensaje"
151               Exit Sub
152           End If
153       End If

154       strListaArticulosNoIncluidos = ""

155       If llngTotalRequerir > 0 Then
156           VsfBusArticulos.Rows = VsfBusArticulos.Rows - 1
157           VsfBusArticulos.Row = 1
158           SSTab.Tab = 1
159           cmdEnviarRequisicion.Enabled = True
160           cmdDelete.Enabled = True
161           cmdDeshacer.Enabled = True
162           cmdImprimir.Enabled = False
              
163           Set rsValor = frsEjecuta_SP(0 & "|" & Me.Name & "|" & VsfBusArticulos.Name & "|Orden|" & vglngNumeroLogin & "|" & "1cintColBusNombre", "SP_GNSELULTIMACONFIGURACION")
164           If rsValor.RecordCount <> 0 Then
165               Select Case Mid(Trim(rsValor!VCHVALOR), 2, Len(Trim(rsValor!VCHVALOR)))
                      Case "cintColBusClave"
166                       VsfBusArticulos.Col = cintColBusClave
167                   Case "cintColBusNombre"
168                       VsfBusArticulos.Col = cintColBusNombre
169                   Case "cintColBusLocalizacion"
170                       VsfBusArticulos.Col = cintColBusLocalizacion
171                   Case "cintColBusAlmacen"
172                       VsfBusArticulos.Col = cintColBusAlmacen
173                   Case "cintColBusIdAlmacen"
174                       VsfBusArticulos.Col = cintColBusIdAlmacen
175                   Case "cintColBusExist"
176                       VsfBusArticulos.Col = cintColBusExist
177                   Case "cintColBusUnidadExist"
178                       VsfBusArticulos.Col = cintColBusUnidadExist
179                   Case "cintColBusPendiente"
180                       VsfBusArticulos.Col = cintColBusPendiente
181                   Case "cintColBusUnidadPendiente"
182                       VsfBusArticulos.Col = cintColBusUnidadPendiente
183                   Case "cintColBusCantidad"
184                       VsfBusArticulos.Col = cintColBusCantidad
185                   Case "cintColBusUnidad"
186                       VsfBusArticulos.Col = cintColBusUnidad
187                   Case "cintColBusTipoUnidad"
188                       VsfBusArticulos.Col = cintColBusTipoUnidad
189                   Case "cintColBusCveDeptoAutoriza"
190                       VsfBusArticulos.Col = cintColBusCveDeptoAutoriza
191                   Case "cintColBusCveDeptoRecibe"
192                       VsfBusArticulos.Col = cintColBusCveDeptoRecibe
193                   Case "cintColBusCantidadOriginal"
194                       VsfBusArticulos.Col = cintColBusCantidadOriginal
195                   Case "cintColsVsfBusArticulos"
196                       VsfBusArticulos.Col = cintColsVsfBusArticulos
197               End Select
198               If Not IsNull(rsValor!VCHVALOR) Then
199                   VsfBusArticulos.Sort = Val(Mid(Trim(rsValor!VCHVALOR), 1, 1))
200                   vlstrColSort = Trim(rsValor!VCHVALOR)
201               End If
202           Else
203               VsfBusArticulos.Col = cintColBusNombre
204               VsfBusArticulos.Sort = flexSortGenericAscending
205               vlstrColSort = flexSortGenericAscending & "cintColBusNombre"
206           End If
207           rsValor.Close

208           VsfBusArticulos.Col = cintColBusCantidad
209           VsfBusArticulos.SetFocus
210       Else
211           If lngSinFaltante <> 0 Then
                  'No se detectaron faltantes con respecto a máximos.
212               MsgBox SIHOMsg(497), vbOKOnly + vbInformation, "Mensaje"
213               If SSTab.Tab = 1 Then
214                   cmdAgregar_Click
215                   SSTab.Tab = 0
216               End If
217           End If
218       End If

219       rsConSigna.Close

220       Exit Sub
NotificaError:
       Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdRequisitar_Click" & " Linea:" & Erl()))
End Sub

Private Function fblnPermisoControlado(intControlado As Integer, lngCveAlmacen As Long, strTipoRequisicion As String) As Boolean
On Error GoTo NotificaError
Dim rsRequisicion As New ADODB.Recordset

    fblnPermisoControlado = True
    
    If intControlado = 1 Then
        vgstrParametrosSP = Str(vglngNumeroLogin) & "|" & Str(lngCveAlmacen) & "|" & IIf(strTipoRequisicion = "R", "RE", "CO")
        
        Set rsRequisicion = frsEjecuta_SP(vgstrParametrosSP, "SP_IVSELREQUISICIONDEPARTAMENT")
        
        If rsRequisicion.RecordCount <> 0 Then
            fblnPermisoControlado = rsRequisicion!bitControlado = 1
        End If
        rsRequisicion.Close
    End If
    
Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fblnPermisoControlado"))
End Function

Private Function flngCantidadPendiente(StrCveArticulo As String, lngCveDepto As Long) As Long
On Error GoTo NotificaError
Dim rsFaltanteMinimo As New ADODB.Recordset

    flngCantidadPendiente = 1
    
    vgstrParametrosSP = Str(vgintDiasRequisicion) & "|" & StrCveArticulo & "|" & Str(lngCveDepto)
    Set rsFaltanteMinimo = frsEjecuta_SP(vgstrParametrosSP, "SP_IVCANTIDADPENDIENTERECIBIR", True, flngCantidadPendiente)

Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":flngCantidadPendiente"))
End Function

Private Sub cmdSelecciona_Click(Index As Integer)
    
    If Index = 0 Then
        lstDisponibles_DblClick
    Else
        lstSeleccionadas_DblClick
    End If
    
End Sub

Private Sub cmdVistaPreliminar_Click()
1     On Error GoTo NotificaError
          Dim alstrParametros(12) As String
          Dim rsReporte As New ADODB.Recordset
          Dim lintTipoArt As Integer
          Dim lintControlado As Integer
          Dim lintRefrigerado As Integer
          Dim llngIdArticulo As Long
          Dim lstrLocalizaciones As String
          Dim lintCont As Integer
          Dim rsAux As New ADODB.Recordset
          Dim strManejos As String
          Dim bitManejo As Integer
          Dim intcontador As Integer
          Dim blnSinManejo As Boolean
          Dim strParamManejos As String
          
2         If chkVarias.Value Then
3             If lstSeleccionadas.ListCount = 0 Then
4                 MsgBox SIHOMsg(3), vbOKOnly + vbCritical, "Mensaje"
5                 fraLocalizaciones.Visible = True
6                 Exit Sub
7             End If
8         Else
9             lstrTotalLocalizaciones = IIf(cboLocalizacion.ItemData(cboLocalizacion.ListIndex) = -1, -1, "_" & cboLocalizacion.ItemData(cboLocalizacion.ListIndex) & "_")
10        End If
          
11        pInstanciaReporte vgrptReporte, "rptMaximosMinimos.rpt"
          
12        If optOpcion(0).Value Then
13            lintTipoArt = -1 'Todos
14        ElseIf optOpcion(1).Value Then
15            lintTipoArt = 0 'Artículos
16        ElseIf optOpcion(2).Value Then
17            lintTipoArt = 1 'Medicamentos
18        ElseIf optOpcion(3).Value Then
19            lintTipoArt = 2 'Insumos
20        End If
          
21        llngIdArticulo = -1
22        lstrCodigoBarras = ""
23        If cboNombreComercial.ListIndex <> -1 Then
24            llngIdArticulo = cboNombreComercial.ItemData(cboNombreComercial.ListIndex)
25        End If
          
26        If llngIdArticulo <> -1 Then
27            vgstrParametrosSP = fstrObtenCveArticulo(llngIdArticulo)
28            Set rsAux = frsEjecuta_SP(vgstrParametrosSP & "|" & "|", "SP_IVSELARTICULO")
29            If rsAux.RecordCount > 0 Then
30            lstrCodigoBarras = IIf(IsNull(rsAux!CodigoBarras), "", rsAux!CodigoBarras)
31            End If
32        End If
          
          'Manejos
33        strManejos = "_"
34        strParamManejos = ""
35        blnSinManejo = False
          
36        If Not optOpcion(3).Value Then
              
37            For intcontador = 0 To lstManejos.ListCount - 1
38                If lstManejos.Selected(intcontador) = True Or cboNombreComercial.ItemData(cboNombreComercial.ListIndex) > 0 Then
                      
39                    strManejos = strManejos & lstManejos.ItemData(intcontador) & "_"
40                    strParamManejos = strParamManejos & IIf(strParamManejos = "", "", ", ") & lstManejos.List(intcontador)
41                    If lstManejos.ItemData(intcontador) = 0 Then blnSinManejo = True
                      
42                End If
43            Next intcontador
              
44            If strManejos = "_" Then
45                bitManejo = 3
46            Else
47                bitManejo = IIf(blnSinManejo, 1, 2)
48            End If
          
49        Else
          
50            bitManejo = 1
51            strManejos = strTodosManejos
52            strParamManejos = ""
              
53        End If
          
54        If llngIdArticulo <> -1 Then
55            bitManejo = 1
56            strManejos = strTodosManejos
57        End If
          
58        vgstrParametrosSP = cboDepartamento.ItemData(cboDepartamento.ListIndex) _
                      & "|" & IIf(llngIdArticulo = -1, IIf(optOpcion(0).Value, 1, 0), 1) _
                      & "|" & IIf(llngIdArticulo = -1, lintTipoArt, -1)
59        vgstrParametrosSP = vgstrParametrosSP & "|" & IIf(llngIdArticulo = -1, IIf(cboFamilia.ItemData(cboFamilia.ListIndex) = -1, 1, 0), 1) _
                      & "|" & IIf(llngIdArticulo = -1, IIf(cboFamilia.ItemData(cboFamilia.ListIndex) = -1, -1, fstrFormateaComoFamilia(Str(cboFamilia.ItemData(cboFamilia.ListIndex)))), -1) _
                      & "|" & IIf(llngIdArticulo = -1, IIf(cboSubFamilia.ItemData(cboSubFamilia.ListIndex) = -1, 1, 0), 1) _
                      & "|" & IIf(llngIdArticulo = -1, IIf(cboSubFamilia.ItemData(cboSubFamilia.ListIndex) = -1, -1, cboSubFamilia.ItemData(cboSubFamilia.ListIndex)), -1) _
                      & "|" & IIf(llngIdArticulo = -1, 1, 0) _
                      & "|" & IIf(llngIdArticulo = -1, "*", fstrObtenCveArticulo(cboNombreComercial.ItemData(cboNombreComercial.ListIndex))) _
                      & "|" & "EXISTENCIAS ACTUALES POR PROVEEDOR" & "|2|-1" _
                      & "|" & IIf(llngIdArticulo = -1, lstrTotalLocalizaciones, "-1") _
                      & "|" & IIf(lblnAlmacenConsigna Or lstrTipoRequisicion = "C", "C", "R") _
                      & "|" & bitManejo _
                      & "|" & strManejos
60        If lstrModo = "C" Then
61            Set rsReporte = frsEjecuta_SP(vgstrParametrosSP, "SP_IVSELARTICULOSUBICACIONEXC", , , , True)
62        Else
63            Set rsReporte = frsEjecuta_SP(vgstrParametrosSP, "SP_IVSELARTICULOSUBICACION", , , , True)
64        End If
          
65        If rsReporte.RecordCount > 0 Then
              
66            If chkVarias.Value Then
67                lstrLocalizaciones = ""
68                For lintCont = 0 To lstSeleccionadas.ListCount - 1
69                    lstrLocalizaciones = lstrLocalizaciones & lstSeleccionadas.List(lintCont) & ","
70                Next lintCont
71                If lstrLocalizaciones <> "" Then lstrLocalizaciones = Mid(lstrLocalizaciones, 1, Len(lstrLocalizaciones) - 1)
72            End If
              
73            vgrptReporte.DiscardSavedData
            
74            alstrParametros(0) = "empresa;" & Trim(vgstrNombreHospitalCH)
75            alstrParametros(1) = "titulo;" & IIf(lblnAlmacenConsigna, "EXISTENCIAS ACTUALES POR PROVEEDOR", "MÁXIMOS, MÍNIMOS Y PUNTOS DE REORDEN")
76            alstrParametros(2) = "almacen;" & cboDepartamento.List(cboDepartamento.ListIndex)
77            alstrParametros(3) = "clasificacion;" & IIf(optOpcion(0).Value, "<TODOS>", IIf(optOpcion(1).Value, "ARTÍCULOS", IIf(optOpcion(2).Value, "MEDICAMENTOS", "INSUMOS")))
78            alstrParametros(4) = "localizacion;" & IIf(chkVarias.Value And llngIdArticulo = -1, lstrLocalizaciones, cboLocalizacion.List(cboLocalizacion.ListIndex))
79            alstrParametros(5) = "familia;" & cboFamilia.List(cboFamilia.ListIndex)
80            alstrParametros(6) = "nombrecomercial;" & cboNombreComercial.List(cboNombreComercial.ListIndex)
81            alstrParametros(7) = "nombregenerico;" & lblNombreGenerico.Caption
82            alstrParametros(8) = "subfamilia;" & cboSubFamilia.List(cboSubFamilia.ListIndex)
83            alstrParametros(9) = "codigodebarras;" & lstrCodigoBarras
84            alstrParametros(10) = "clave;" & Trim(txtClave)
85            alstrParametros(11) = "MostrarEstado;" & 1
86            alstrParametros(12) = "Manejos" & ";" & strParamManejos & ";TRUE"
              
87            pCargaParameterFields alstrParametros, vgrptReporte
            
88            pImprimeReporte vgrptReporte, rsReporte, "P", IIf(lblnAlmacenConsigna, "Existencias actuales por proveedor", "Máximos, mínimos y puntos de reorden")

89        Else
              'No existe información con esos parámetros.
90            MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje"
91        End If
92        If rsReporte.State <> adStateClosed Then rsReporte.Close
93        If rsAux.State <> adStateClosed Then rsAux.Close
          
94    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdVistaPreliminar_Click" & " Linea:" & Erl()))
End Sub

Private Sub Form_Activate()
On Error GoTo NotificaError

    If lblnEntrando Then
        If cboDepartamento.ListCount = 0 Then
            'El usuario no tiene departamentos asignados para asignar máximos y mínimos.
            MsgBox SIHOMsg(866), vbOKOnly + vbExclamation, "Mensaje"
            Unload Me
        Else
            pLimpiar
            If cboDepartamento.ListIndex = -1 And cboDepartamento.ListCount > 0 Then
                cboDepartamento.ListIndex = 0
            End If
        End If
        
        lblnEntrando = False
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_Activate"))
End Sub

Private Sub pLimpiar()
1     On Error GoTo NotificaError
      Dim intcontador As Integer

2         optOpcion(0).Value = True
3         optOpcion_Click 0

4         intMaxManejos = 0
          
5         For intcontador = 0 To lstManejos.ListCount - 1
6             lstManejos.Selected(intcontador) = True
7         Next intcontador
          
8         fraTipoAsignacion.Enabled = False
9         optconsumo(0).Value = True
10        optconsumo(0).Enabled = False
11        optconsumo(1).Enabled = False
          
12        pLimpiarVsfArticulos
13        pConfVsfArticulos
          
14        lblNombreComercialCompleto.Caption = ""
          
15        cmdGrabar.Enabled = False
16        cmdRequisitar.Enabled = False
17        cmdConsultarReq.Enabled = False
18        cmdVistaPreliminar.Enabled = False
          
19        chkVarias.Value = 0
           chkMostrarArti.Value = 0
20        For lintCiclos = 0 To lstSeleccionadas.ListCount - 1
21            lstDisponibles.AddItem lstSeleccionadas.List(lintCiclos)
22            lstDisponibles.ItemData(lstDisponibles.NewIndex) = lstSeleccionadas.ItemData(lintCiclos)
23        Next lintCiclos
24        lstSeleccionadas.Clear

25        cmdAsignarDesactivar(0).Enabled = False
26        cmdAsignarDesactivar(1).Enabled = False
          cmdImportar(0).Enabled = True
          cmdExportar(2).Enabled = False
          
27    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pLimpiar" & " Linea:" & Erl()))
End Sub

Private Sub pLimpiarVsfArticulos()
1     On Error GoTo NotificaError
      Dim rsConSigna As ADODB.Recordset
      Dim vgstrParametrosSP As String
          
2         If cboDepartamento.ListIndex > -1 Then
3             vgstrParametrosSP = 1 & "|" & vgintClaveEmpresaContable & "|" & CStr(cboDepartamento.ItemData(cboDepartamento.ListIndex))
4         Else
5             vgstrParametrosSP = 1 & "|" & vgintClaveEmpresaContable & "|" & CStr(cboDepartamento.ItemData(0))
6         End If
          
7         Set rsConSigna = frsEjecuta_SP(vgstrParametrosSP, "Sp_Ivselalmacenconsigna")

8         llngTotalMarcados = 0
9         llngTotalDesactivar = 0
          
10        lblNombreComercialCompleto.Caption = ""
          
11        cmdMarcar.Enabled = False
12        cmdInvertirSeleccion.Enabled = False
          
13       vsfArticulos.Clear
14       vsfArticulos.Rows = 2
         If strClavesConsol <> "" Then
            vsfArticulos.Cols = cintColsVsfArtiBusArtiConsol
         Else
            vsfArticulos.Cols = cintColsVsfArticulos
         End If
         
          
16        intMaxManejos = 0
          
17        If rsConSigna.RecordCount <> 0 Then
18            vsfArticulos.FormatString = "|.|.|.|.|" & cstrTitulosVsfArticulos2
19        Else
20            vsfArticulos.FormatString = "|.|.|.|.|" & cstrTitulosVsfArticulos
21        End If
22        rsConSigna.Close
          
23    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pLimpiarVsfArticulos" & " Linea:" & Erl()))
End Sub

Private Sub pLimpiarVsfBusArticulos()
On Error GoTo NotificaError

    llngTotalRequerir = 0

    VsfBusArticulos.Clear
    VsfBusArticulos.Rows = 2
    If strClavesConsol <> "" Then
        VsfBusArticulos.Cols = cintColsVsfBusArtiConsol
    Else
        VsfBusArticulos.Cols = cintColsVsfBusArticulos
    End If
    
    VsfBusArticulos.FormatString = cstrTitulosVsfBusArticulos

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pLimpiarVsfBusArticulos"))
End Sub

Private Sub pConfVsfBusArticulos()
On Error GoTo NotificaError

    With VsfBusArticulos
        .ColWidth(0) = 100
        .ColWidth(cintColBusClave) = 1300
        .AutoSize cintColBusNombre
        .ColWidth(cintColBusNombre) = IIf(.ColWidth(cintColBusNombre) > 4000, 4000, .ColWidth(cintColBusNombre))
        .AutoSize cintColBusLocalizacion
        .AutoSize cintColBusAlmacen
        .ColWidth(cintColBusIdAlmacen) = 0
        .AutoSize cintColBusExist
        .AutoSize cintColBusUnidadExist
        .ColWidth(cintColBusPendiente) = 1750
        .AutoSize cintColBusUnidadPendiente
        .ColWidth(cintColBusCantidad) = 1900
        .ColWidth(cintColBusUnidad) = 1500
        .ColWidth(cintColBusTipoUnidad) = 0
        .ColWidth(cintColBusCveDeptoAutoriza) = 0
        .ColWidth(cintColBusCveDeptoRecibe) = 0
        .ColWidth(cintColBusCantidadOriginal) = 0
        
        If strClavesConsol <> "" Then
            .ColWidth(cintColBusExisPrin) = 0
            .ColWidth(cintColBusMaxPrin) = 0
            .ColWidth(cintColBusExisEsclavos) = 0
            .ColWidth(cintColBusMaxEsclavos) = 0
            .ColWidth(cintColBusDeptoEsclavos) = 0
        End If
        
        .ColAlignment(cintColBusClave) = flexAlignLeftCenter
        .ColAlignment(cintColBusNombre) = flexAlignLeftCenter
        .ColAlignment(cintColBusLocalizacion) = flexAlignLeftCenter
        .ColAlignment(cintColBusAlmacen) = flexAlignLeftCenter
        .ColAlignment(cintColBusExist) = flexAlignRightCenter
        .ColAlignment(cintColBusUnidadExist) = flexAlignLeftCenter
        .ColAlignment(cintColBusPendiente) = flexAlignRightCenter
        .ColAlignment(cintColBusUnidadPendiente) = flexAlignLeftCenter
        .ColAlignment(cintColBusCantidad) = flexAlignRightCenter
        .ColAlignment(cintColBusUnidad) = flexAlignLeftCenter
        
        .FixedAlignment(cintColBusClave) = flexAlignCenterCenter
        .FixedAlignment(cintColBusNombre) = flexAlignCenterCenter
        .FixedAlignment(cintColBusLocalizacion) = flexAlignCenterCenter
        .FixedAlignment(cintColBusAlmacen) = flexAlignCenterCenter
        .FixedAlignment(cintColBusExist) = flexAlignCenterCenter
        .FixedAlignment(cintColBusUnidadExist) = flexAlignCenterCenter
        .FixedAlignment(cintColBusPendiente) = flexAlignCenterCenter
        .FixedAlignment(cintColBusUnidadPendiente) = flexAlignCenterCenter
        .FixedAlignment(cintColBusCantidad) = flexAlignCenterCenter
        .FixedAlignment(cintColBusUnidad) = flexAlignCenterCenter
    End With

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pConfVsfBusArticulos"))
End Sub

Private Sub pConfVsfArticulos()
On Error GoTo NotificaError
Dim intcontador As Integer

    With vsfArticulos
        .ColWidth(0) = 150
        
        For intcontador = 1 To cintColVsfManejo4
            .Col = intcontador
            .Row = 0
            .ColWidth(intcontador) = 0
            .ColAlignment(intcontador) = flexAlignCenterCenter
            .FixedAlignment(cintColVsfManejo1) = flexAlignCenterCenter
            .CellForeColor = &H8000000F
        Next intcontador
        
        .ColWidth(cintColVsfIdArticulo) = 0
        .ColWidth(cintColVsfNombreComercial) = 3200
        .ColWidth(cintColVsfCapturaMaximo) = 900
        .ColWidth(cintColVsfCapturaUnidadMax) = 1150
        .ColWidth(cintColVsfCapturaCveUnidadMax) = 0
        .ColWidth(cintColVsfCapturaPunto) = 900
        .ColWidth(cintColVsfCapturaUnidadPun) = 1200
        .ColWidth(cintColVsfCapturaCveUnidadPun) = 0
        .ColWidth(cintColVsfCapturaMinimo) = 900
        .ColWidth(cintColVsfCapturaUnidadMin) = 1200
        .ColWidth(cintColVsfCapturaCveUnidadMin) = 0
        .ColWidth(cintColVsfCapturaAlmacen) = IIf(lstrModo = "A", 1900, 0)
        .ColWidth(cintColVsfCapturaCveAlmacen) = 0
        .ColWidth(cintColVsfCapturaDeptoCompra) = IIf(lstrModo = "A", 1750, 0)
        .ColWidth(cintColVsfCapturaCveDeptoCompra) = 0
        .ColWidth(cintColVsfClaveArt) = 0
        .ColWidth(cintColVsfContenidoArt) = 0
        .ColWidth(cintColVsfCveUniAlternaArt) = 0
        .ColWidth(cintColvsfUnidadAlternaArt) = 0
        .ColWidth(cintColVsfCveUniMinimaArt) = 0
        .ColWidth(cintColvsfUnidadMinimaArt) = 0
        .ColWidth(cintColvsfCapturaTipoUniMax) = 0
        .ColWidth(cintColvsfCapturaTipoUniPun) = 0
        .ColWidth(cintColvsfCapturaTipoUniMin) = 0
        .ColWidth(cintcolvsfGuardar) = 0
        .ColWidth(cintColVsfExistencia) = IIf(lstrModo = "A", 0, 1000)
        .ColWidth(cintColVsfExistenciaUnidad) = IIf(lstrModo = "A", 0, 1200)
        .ColWidth(cintColvsfEstadoArt) = IIf(lstrModo = "C", 1600, 0)
        .ColWidth(cintColvsfRequisitar) = 0
        .ColWidth(cintColvsfCveArt) = 0
        .ColWidth(cintColvsfControlado) = 0
        .ColWidth(cintColvsfExistenciaAlterna) = 0
        .ColWidth(cintColvsfExistenciaMinima) = 0
        .ColWidth(cintColvsfLocalizacion) = 0
        
        If strClavesConsol <> "" Then
            .ColWidth(cintColVsfArtiExisEsclavos) = 0
            .ColWidth(cintColVsfArtiMaxEsclavos) = 0
            .ColWidth(cintColVsfArtiReordenEsclavos) = 0
            .ColWidth(cintColVsfArtiDeptoEsclavos) = 0
            .ColWidth(cintColVsfArtiMinEsclavos) = 0
        End If
        
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(cintColVsfNombreComercial) = flexAlignLeftCenter
        .ColAlignment(cintColVsfExistencia) = flexAlignRightCenter
        .ColAlignment(cintColVsfExistenciaUnidad) = flexAlignLeftCenter
        .ColAlignment(cintColVsfCapturaMaximo) = flexAlignRightCenter
        .ColAlignment(cintColVsfCapturaUnidadMax) = flexAlignLeftCenter
        .ColAlignment(cintColVsfCapturaPunto) = flexAlignRightCenter
        .ColAlignment(cintColVsfCapturaUnidadPun) = flexAlignLeftCenter
        .ColAlignment(cintColVsfCapturaMinimo) = flexAlignRightCenter
        .ColAlignment(cintColVsfCapturaUnidadMin) = flexAlignLeftCenter
        .ColAlignment(cintColVsfCapturaAlmacen) = flexAlignLeftCenter
        .ColAlignment(cintColVsfCapturaDeptoCompra) = flexAlignLeftCenter
        
        .FixedAlignment(cintColVsfNombreComercial) = flexAlignLeftCenter
        .FixedAlignment(cintColVsfExistencia) = flexAlignRightCenter
        .FixedAlignment(cintColVsfExistenciaUnidad) = flexAlignLeftCenter
        .FixedAlignment(cintColVsfCapturaMaximo) = flexAlignRightCenter
        .FixedAlignment(cintColVsfCapturaUnidadMax) = flexAlignLeftCenter
        .FixedAlignment(cintColVsfCapturaPunto) = flexAlignRightCenter
        .FixedAlignment(cintColVsfCapturaUnidadPun) = flexAlignLeftCenter
        .FixedAlignment(cintColVsfCapturaMinimo) = flexAlignRightCenter
        .FixedAlignment(cintColVsfCapturaUnidadMin) = flexAlignLeftCenter
        .FixedAlignment(cintColVsfCapturaAlmacen) = flexAlignLeftCenter
        .FixedAlignment(cintColVsfCapturaDeptoCompra) = flexAlignLeftCenter
        .FixedAlignment(cintColvsfEstadoArt) = flexAlignLeftCenter
        
        .MergeCells = flexMergeRestrictRows
        .MergeRow(0) = True
    End With

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pConfVsfArticulos"))
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError

    If KeyAscii = 27 Then
        If fraRequisiciones.Visible Then
            cmdCerrar_Click
        ElseIf fraLocalizaciones.Visible Then
            cmdAceptar_Click
        Else
            KeyAscii = 0
            Unload Me
        End If
    Else
        If KeyAscii = 13 Then
            If Me.ActiveControl.Name = "vsfArticulos" Then
                If vsfArticulos.Col = cintColVsfCapturaMaximo _
                Or vsfArticulos.Col = cintColVsfCapturaUnidadMax _
                Or vsfArticulos.Col = cintColVsfCapturaPunto _
                Or vsfArticulos.Col = cintColVsfCapturaUnidadPun _
                Or vsfArticulos.Col = cintColVsfCapturaMinimo _
                Or vsfArticulos.Col = cintColVsfCapturaUnidadMin _
                Or vsfArticulos.Col = cintColVsfCapturaAlmacen _
                Or vsfArticulos.Col = cintColVsfCapturaDeptoCompra Then
                    If vsfArticulos.Col = cintColVsfCapturaMaximo Then
                        vsfArticulos.Col = cintColVsfCapturaUnidadMax
                         KeyAscii = 0
                    ElseIf vsfArticulos.Col = cintColVsfCapturaUnidadMax Then
                        vsfArticulos.Col = cintColVsfCapturaPunto
                        KeyAscii = 0
                    ElseIf vsfArticulos.Col = cintColVsfCapturaPunto Then
                        vsfArticulos.Col = cintColVsfCapturaUnidadPun
                        KeyAscii = 0
                    ElseIf vsfArticulos.Col = cintColVsfCapturaUnidadPun Then
                        vsfArticulos.Col = cintColVsfCapturaMinimo
                        KeyAscii = 0
                    ElseIf vsfArticulos.Col = cintColVsfCapturaMinimo Then
                        vsfArticulos.Col = cintColVsfCapturaUnidadMin
                        KeyAscii = 0
                    ElseIf vsfArticulos.Col = cintColVsfCapturaUnidadMin Then
                        vsfArticulos.Col = cintColVsfCapturaAlmacen
                        KeyAscii = 0
                    ElseIf vsfArticulos.Col = cintColVsfCapturaAlmacen Then
                        KeyAscii = 0
                        vsfArticulos.Col = cintColVsfCapturaDeptoCompra
                    ElseIf vsfArticulos.Col = cintColVsfCapturaDeptoCompra Then
                        If vsfArticulos.Row < vsfArticulos.Rows - 1 Then
                            vsfArticulos.Col = cintColVsfCapturaMaximo
                            vsfArticulos.Row = vsfArticulos.Row + 1
                            KeyAscii = 0
                        End If
                    End If
                    vsfArticulos.SetFocus
                Else
                    SendKeys vbTab
                End If
            Else
                SendKeys vbTab
            End If
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_KeyPress"))
End Sub

Private Sub Form_Load()
1     On Error GoTo NotificaError
      Dim rsConSigna As New ADODB.Recordset
      Dim rs As New ADODB.Recordset
      Dim blnBandera As Boolean
      Dim intOpcionCompraDirecta As Long
      Dim rsParametro As New ADODB.Recordset
      Dim rsIvparametro As New ADODB.Recordset
2         Me.Icon = frmMenuPrincipal.Icon
          
          'Color de tabs
3         SetStyle SSTab.hwnd, 0
4         SetSolidColor SSTab.hwnd, 16777215
5         SSTabSubclass SSTab.hwnd
         
6         Me.Caption = IIf(lstrModo = "C", "Requisiciones automáticas de faltantes", "Máximos, mínimos y puntos de reorden")

          ' Revisa permiso para Compra directa
7         Select Case cgstrModulo
              Case "IM"
8                 intOpcionCompraDirecta = 2465
9             Case "PV"
10                intOpcionCompraDirecta = 2466
11            Case "CP"
12                intOpcionCompraDirecta = 2467
13            Case "LA"
14                intOpcionCompraDirecta = 2468
15            Case "EX"
16                intOpcionCompraDirecta = 2469
17            Case "CN"
18                intOpcionCompraDirecta = 2470
19            Case "CC"
20                intOpcionCompraDirecta = 2471
21            Case "NO"
22                intOpcionCompraDirecta = 2472
23            Case "CA"
24                intOpcionCompraDirecta = 2473
25            Case "IV"
26                intOpcionCompraDirecta = 2474
27            Case "AD"
28                intOpcionCompraDirecta = 2475
29            Case "CO"
30                intOpcionCompraDirecta = 2476
31            Case "SI"
32                intOpcionCompraDirecta = 2477
33            Case "BS"
34                intOpcionCompraDirecta = 2478
35            Case "SE"
36                intOpcionCompraDirecta = 2479
37            Case "DI"
38                intOpcionCompraDirecta = 2480
39            Case "TS"
40                intOpcionCompraDirecta = 2481
41        End Select
          
42        blnPermisoCompraDirecta = fblnRevisaPermiso(vglngNumeroLogin, intOpcionCompraDirecta, "E", True)

          'Números de opción
43        pCargarOpciones
          
44        lblnAlmacenConsigna = False
45        vgstrParametrosSP = 1 & "|" & vgintClaveEmpresaContable & "|" & CStr(vgintNumeroDepartamento)
46        Set rsConSigna = frsEjecuta_SP(vgstrParametrosSP, "Sp_Ivselalmacenconsigna")
47        If rsConSigna.RecordCount > 0 Then lblnAlmacenConsigna = True
          
48        lblnReubicarInsumos = False
49        pCargaDeptos
          
          'Manejos
50        blnBandera = True
51        strTodosManejos = "_"
52        lstManejos.Clear
53        Set rs = frsEjecuta_SP("-1|1|-1", "Sp_IvSelManejos")
54        If rs.RecordCount > 0 Then
              
55            Do While Not rs.EOF
56                lstManejos.AddItem Trim(rs!VCHDESCRIPCION)
57                lstManejos.ItemData(lstManejos.NewIndex) = rs!intCveManejo
58                lstManejos.Selected(lstManejos.NewIndex) = True
                  
59                strTodosManejos = strTodosManejos & CStr(rs!intCveManejo) & "_"
                  
60                rs.MoveNext
61                blnBandera = False
62            Loop
63            lstManejos.AddItem "SIN MANEJO"
64            lstManejos.ItemData(lstManejos.NewIndex) = 0
65            lstManejos.Selected(lstManejos.NewIndex) = True
              
66            strTodosManejos = strTodosManejos & "0_"
              
67            lstManejos.ListIndex = 0
68        End If
69        rs.Close
          
          'Modo consulta, no se ve nada de la asignación de máximos
70        fraTipoAsignacion.Visible = lstrModo = "A"
71        fraRangoConsumo.Visible = lstrModo = "A"
72        cmdCalcular.Visible = lstrModo = "A"
73        fraGrabar.Visible = lstrModo = "A"
74        cmdMarcar.Visible = lstrModo = "A"
75        cmdInvertirSeleccion.Visible = lstrModo = "A"
76        cmdAsignarDesactivar(0).Visible = lstrModo = "A"
77        cmdAsignarDesactivar(0).Visible = lstrModo = "A"
78        fraRequisitar.Visible = lstrModo = "C"
79        chkVarias.Visible = lstrModo = "C"
          chkMostrarArti.Visible = lstrModo = "A"
          
          'Apariencia de controles, en modo consulta o asignación
80        fraArticulos.Top = IIf(lstrModo = "A", clngTopModoAsignacion, clngTopModoConsulta)
81        fraArticulos.height = IIf(lstrModo = "A", clngHeightModoAsignacion, clngHeightModoConsulta)
82        vsfArticulos.height = IIf(lstrModo = "A", clngHeightVSFModoAsignacion, clngHeightVSFModoConsulta)
83        cboLocalizacion.width = IIf(lstrModo = "A", 5295, 2950)
          
          'Verificar si tiene permisos para asignar al departamento, solo en modo asignación
84        If lstrModo = "A" Then
85            lblnPermisoAsignar = fblnRevisaPermiso(vglngNumeroLogin, llngNumOpcionAsignar, "E", True)
86            cmdAsignarDesactivar(0).Enabled = False
87            cmdAsignarDesactivar(1).Enabled = False
88        End If
          
89        cmdAgregar.Caption = IIf(lstrModo = "C", "Consultar artículos", "Agregar artículos")
          
90        fraBarra.Visible = False
          
91        lblnEntrando = True
92        lblnBusquedaNombre = False
          
93        lintInterfazFarmaciaSJP = 0
94        llngDeptoSubrogado = 0
95        If lstrModo = "C" Then
96            Set rsParametro = frsSelParametros("IV", -1, "INTINTERFAZFARMACIASUBRROGADASJP")
97            If Not rsParametro.EOF Then
98                lintInterfazFarmaciaSJP = IIf(IsNull(rsParametro("Valor")), 0, rsParametro("Valor"))
99            Else
100               lintInterfazFarmaciaSJP = 0
101           End If
102           If lintInterfazFarmaciaSJP = 1 Then
                  'Consultar parametro de departamento subrogado
103               Set rsParametro = frsSelParametros("IV", -1, "INTDEPTOINTERFAZFARMACIA")
104               If Not rsParametro.EOF Then
105                   llngDeptoSubrogado = IIf(IsNull(rsParametro("Valor")), 0, rsParametro("Valor"))
106               Else
107                   llngDeptoSubrogado = 0
108               End If
109           End If
110       End If
          
111       grdRequisicion.Visible = False
          
112       fraRequisiciones.Visible = False
          
113       SSTab.Tab = 0
114       VsfBusArticulos.Editable = 0

115   Exit Sub
NotificaError:
       Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_Load" & " Linea:" & Erl()))
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo NotificaError

    If SSTab.Tab <> 0 Then
        Cancel = 1
        SSTab.Tab = 0
        VsfBusArticulos.Editable = 0
    
        If Not cmdEnviarRequisicion.Enabled And Not cmdDelete.Enabled And cmdImprimir.Enabled Then
            cmdAgregar_Click
        Else
            cmdRequisitar.SetFocus
        End If
    Else
        If fraRequisiciones.Visible Then
            cmdCerrar_Click
        End If
        If Val(Trim(vsfArticulos.TextMatrix(1, cintColVsfIdArticulo))) <> 0 Then
            Cancel = 1
            '¿Desea abandonar la operación?
            If MsgBox(SIHOMsg(17), vbQuestion + vbYesNo, "Mensaje") = vbYes Then
                pLimpiar
                txtClave.Text = ""
                If cboLocalizacion.ListCount > 0 Then cboLocalizacion.ListIndex = 0
                If cboNombreComercial.ListCount > 0 Then cboNombreComercial.ListIndex = 0
                If fblnCanFocus(cboDepartamento) Then cboDepartamento.SetFocus Else optOpcion(0).SetFocus
            End If
        End If
    End If

    cboUnidad.Visible = False
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_QueryUnload"))
End Sub

Private Sub Form_Unload(Cancel As Integer)
    vgstrVarIntercam = ""
End Sub

Private Sub lblNombreComercialCompleto_Change()
    If lblNombreComercialCompleto.Caption = "NombreR" Or lblNombreComercialCompleto.Caption = "R" Or lblNombreComercialCompleto.Caption = "Nombre" Then
        lblNombreComercialCompleto.Caption = ""
    End If
End Sub

Private Sub lstDisponibles_DblClick()
Dim lintPosicion As Integer

    If lstDisponibles.ListIndex <> -1 Then
        lstSeleccionadas.AddItem lstDisponibles.List(lstDisponibles.ListIndex)
        lstSeleccionadas.ItemData(lstSeleccionadas.NewIndex) = lstDisponibles.ItemData(lstDisponibles.ListIndex)
        lintPosicion = lstDisponibles.ListIndex
        lstDisponibles.RemoveItem (lstDisponibles.ListIndex)
        If lstDisponibles.ListCount > 0 Then
            If lintPosicion = lstDisponibles.ListCount Then lintPosicion = 0
            lstDisponibles.ListIndex = lintPosicion
        End If
        If lstSeleccionadas.ListCount > 0 Then lstSeleccionadas.ListIndex = 0
    End If
    
End Sub

Private Sub lstSeleccionadas_DblClick()
    
    If lstSeleccionadas.ListIndex <> -1 Then
        lstDisponibles.AddItem lstSeleccionadas.List(lstSeleccionadas.ListIndex)
        lstDisponibles.ItemData(lstDisponibles.NewIndex) = lstSeleccionadas.ItemData(lstSeleccionadas.ListIndex)
        lstSeleccionadas.RemoveItem (lstSeleccionadas.ListIndex)
        If lstDisponibles.ListCount > 0 Then lstDisponibles.ListIndex = 0
        If lstSeleccionadas.ListCount > 0 Then lstSeleccionadas.ListIndex = 0
    End If
    
End Sub

Private Sub mskFin_GotFocus()
On Error GoTo NotificaError

    pSelMkTexto mskFin

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFin_GotFocus"))
End Sub

Private Sub mskInicio_GotFocus()
On Error GoTo NotificaError

    pSelMkTexto mskInicio

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskInicio_GotFocus"))
End Sub

Private Sub optconsumo_Click(Index As Integer)
On Error GoTo NotificaError

    fraRangoConsumo.Enabled = Index = 1
    mskInicio.Enabled = Index = 1
    mskFin.Enabled = Index = 1
    lblAl.Enabled = Index = 1
    cmdCalcular.Enabled = Index = 1

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":optconsumo_Click"))
End Sub

Private Sub optOpcion_Click(Index As Integer)
On Error GoTo NotificaError
Dim intcontador As Integer

    If Not blnNoCargarFiltros Then
    
        lintTipoArticulo = -1
    
        If Index = 1 Then
            lintTipoArticulo = 0
        ElseIf Index = 2 Then
            lintTipoArticulo = 1
        ElseIf Index = 3 Then
            lintTipoArticulo = 2
        End If
    
        cboFamilia.Clear
        If Index <> 0 Then
            pCargaFamilias
        End If
        cboFamilia.AddItem "<TODAS>", 0
        cboFamilia.ItemData(cboFamilia.NewIndex) = -1
        cboFamilia.ListIndex = 0
        
        For intcontador = 0 To lstManejos.ListCount - 1
            lstManejos.Selected(intcontador) = True
        Next intcontador
        
        lstManejos.Enabled = Index <> 3
        lblFamilia.Enabled = Index <> 0
        cboFamilia.Enabled = Index <> 0
    
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":optOpcion_Click"))
End Sub

Private Sub pCargaFamilias()
On Error GoTo NotificaError

    Set rs = frsEjecuta_SP(CStr(lintTipoArticulo), "sp_IvSelFamilia")
    If rs.RecordCount <> 0 Then
        pLlenarCboRs_new cboFamilia, rs, 0, 1
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargaFamilias"))
End Sub

Private Sub txtClave_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo NotificaError

    If KeyCode = vbKeyReturn And Trim(txtClave.Text) <> "" Then
        
        lblnBusquedaNombre = True
        
        vgstrVarIntercam = UCase(txtClave.Text)
        vgstrVarIntercam2 = "Lista por clave"
        
        frmLista.gintEstatus = 1
        frmLista.gintFamilia = IIf(cboFamilia.ItemData(cboFamilia.ListIndex) = -1, 0, cboFamilia.ItemData(cboFamilia.ListIndex))
        frmLista.gintSubfamilia = IIf(cboSubFamilia.ItemData(cboSubFamilia.ListIndex) = -1, 0, cboSubFamilia.ItemData(cboSubFamilia.ListIndex))
        frmLista.Tag = ""
        If lstrModo = "C" Then
            frmLista.Tag = IIf(cboTipoRequisicion.List(cboTipoRequisicion.ListIndex) = "REUBICACION" And lblnReubicarInsumos, "T", Mid(cboTipoRequisicion.List(cboTipoRequisicion.ListIndex), 1, 1))
        End If
        frmLista.Show vbModal, Me
        If Len(vgstrVarIntercam) > 0 Then
            pDatosArticulo Trim(vgstrVarIntercam), -1
            cmdAgregar.SetFocus
        End If
        
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtClave_KeyDown"))
End Sub

Private Sub txtClave_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError

    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
        KeyAscii = 7
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtClave_KeyPress"))
End Sub

Private Sub vsfArticulos_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo NotificaError

    cmdGrabar.Enabled = True

    With vsfArticulos
        
        .TextMatrix(Row, cintcolvsfGuardar) = "1"
        'Para que deje un 0 al no teclearse nada
        If Col = cintColVsfCapturaMaximo Or Col = cintColVsfCapturaPunto Or Col = cintColVsfCapturaMinimo Then .Text = IIf(Trim(.Text) = "", "0", .Text)
        'Máximo:
        If Col = cintColVsfCapturaUnidadMax And .ComboIndex <> -1 Then
            .TextMatrix(.Row, cintColVsfCapturaCveUnidadMax) = .ComboData(.ComboIndex)
            .TextMatrix(.Row, cintColvsfCapturaTipoUniMax) = .ComboIndex
        End If
        'Punto reorden:
        If Col = cintColVsfCapturaUnidadPun And .ComboIndex <> -1 Then
            .TextMatrix(.Row, cintColVsfCapturaCveUnidadPun) = .ComboData(.ComboIndex)
            .TextMatrix(.Row, cintColvsfCapturaTipoUniPun) = .ComboIndex
        End If
        'Mínimo:
        If Col = cintColVsfCapturaUnidadMin And .ComboIndex <> -1 Then
            .TextMatrix(.Row, cintColVsfCapturaCveUnidadMin) = .ComboData(.ComboIndex)
            .TextMatrix(.Row, cintColvsfCapturaTipoUniMin) = .ComboIndex
        End If
        'Almacén reubica:
        If Col = cintColVsfCapturaAlmacen And .ComboIndex <> -1 Then
            .TextMatrix(.Row, cintColVsfCapturaCveAlmacen) = .ComboData(.ComboIndex)
        End If
        'Departamento compra:
        If Col = cintColVsfCapturaDeptoCompra And .ComboIndex <> -1 Then
            .TextMatrix(.Row, cintColVsfCapturaCveDeptoCompra) = .ComboData(.ComboIndex)
        End If
        .ComboList = ""
    End With

    Form_KeyPress 13

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":vsfArticulos_AfterEdit"))
End Sub

Private Sub vsfArticulos_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo NotificaError
Dim lngCveUnidad As Long
Dim rsConSigna As ADODB.Recordset
Dim vgstrParametrosSP As String

    vgstrParametrosSP = 1 & "|" & vgintClaveEmpresaContable & "|" & CStr(cboDepartamento.ItemData(cboDepartamento.ListIndex))
    Set rsConSigna = frsEjecuta_SP(vgstrParametrosSP, "Sp_Ivselalmacenconsigna")
    
    With vsfArticulos
        If lstrModo = "C" Then
            'En modo consulta no se permite la edición
            Cancel = True
        Else
            If Val(.TextMatrix(1, cintColVsfIdArticulo)) = 0 Or Col = cintColVsfNombreComercial Or Col = cintColVsfExistencia Or Col = cintColVsfExistenciaUnidad Then
                Cancel = True
            Else
                .ComboList = ""
                    
                If Col = cintColVsfCapturaUnidadMax Or Col = cintColVsfCapturaUnidadMin Or Col = cintColVsfCapturaUnidadPun Then
                    .ComboList = fstrUnidades(vsfArticulos, cintColVsfCveUniAlternaArt, cintColvsfUnidadAlternaArt, cintColVsfCveUniMinimaArt, cintColvsfUnidadMinimaArt, cintColVsfContenidoArt, True)
                    
                    If Col = cintColVsfCapturaUnidadMax Then
                        'Máximo:
                        lngCveUnidad = Val(.TextMatrix(Row, cintColVsfCapturaCveUnidadMax))
                    ElseIf Col = cintColVsfCapturaUnidadMin Then
                        'Punto de reorden:
                        lngCveUnidad = Val(.TextMatrix(Row, cintColVsfCapturaCveUnidadMin))
                    ElseIf Col = cintColVsfCapturaUnidadPun Then
                        'Mínimo:
                        lngCveUnidad = Val(.TextMatrix(Row, cintColVsfCapturaCveUnidadPun))
                    End If
                    
                    .ComboIndex = fintIndex(lngCveUnidad)
                End If
                
                If Col = cintColVsfCapturaAlmacen Then
                    If Trim(lstrDeptosReubican) = "" Then
                        'No existen almacenes que reubiquen a este departamento
                        
                        If rsConSigna.RecordCount = 0 Then
                            MsgBox SIHOMsg(870), vbOKOnly + vbInformation, "Mensaje"
                        End If
                        
                    Else
                        .ComboList = lstrDeptosReubican
                        .ComboIndex = fintIndex(Val(.TextMatrix(Row, cintColVsfCapturaCveAlmacen)))
                    End If
                End If
                
                If Col = cintColVsfCapturaDeptoCompra Then
                    If Trim(lstrDeptosCompran) = "" Then
                        'El departamento que realiza la compra no ha sido configurado.
                        MsgBox SIHOMsg(903), vbOKOnly + vbInformation, "Mensaje"
                    Else
                        .ComboList = lstrDeptosCompran
                        .ComboIndex = fintIndex(.TextMatrix(Row, cintColVsfCapturaCveDeptoCompra))
                    End If
                End If
            End If
        End If
        
    End With

    rsConSigna.Close
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":vsfArticulos_BeforeEdit"))
End Sub

Private Function fintIndex(lngId As Long) As Integer
On Error GoTo NotificaError
Dim intcontador As Integer

    If lngId = 0 Then
        fintIndex = 0
    Else
        intcontador = 0
        Do While intcontador <= vsfArticulos.ComboCount - 1
            If lngId = vsfArticulos.ComboData(intcontador) Then
                fintIndex = intcontador
            End If
            intcontador = intcontador + 1
        Loop
    End If

Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fintIndex"))
End Function

Private Function fstrUnidades(vsfFlexGrid As VSFlexGrid, intColCveUniAlternaArt As Integer, intColUnidadAlternaArt As Integer, intColCveUniMinimaArt As Integer, intColUnidadMinimaArt As Integer, intColContenidoArt As Integer, blnIncluirVacio As Boolean) As String
On Error GoTo NotificaError

    With vsfFlexGrid
    
        If blnIncluirVacio Then
            fstrUnidades = "|#" & Trim(Str(-1)) & ";" & " "
        End If
        
        'La unidad alterna:
        fstrUnidades = fstrUnidades & "|#" & Trim(.TextMatrix(.Row, intColCveUniAlternaArt)) & ";" & .TextMatrix(.Row, intColUnidadAlternaArt)
        
        If Val(.TextMatrix(vsfFlexGrid.Row, intColContenidoArt)) > 1 Then
            'La unidad mínima:
            fstrUnidades = fstrUnidades & "|#" & Trim(.TextMatrix(.Row, intColCveUniMinimaArt)) & ";" & .TextMatrix(.Row, intColUnidadMinimaArt)
        End If
    End With

Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fstrUnidades"))
End Function

Private Sub vsfArticulos_Click()
On Error GoTo NotificaError

    lblNombreComercialCompleto.Caption = vsfArticulos.TextMatrix(vsfArticulos.Row, cintColVsfNombreComercial)
    If lblNombreComercialCompleto.Caption = "Nombre Comercial" Then
        lblNombreComercialCompleto.Caption = ""
    Else
    
    End If
    If lblNombreComercialCompleto.Caption = "Nombre" Then
        lblNombreComercialCompleto.Caption = ""
    Else
    
    End If
    If lblNombreComercialCompleto.Caption = "NombreR" Then
        lblNombreComercialCompleto.Caption = ""
    Else
    
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":vsfArticulos_Click"))
End Sub

Private Sub vsfArticulos_DblClick()
On Error GoTo NotificaError

    If Val(vsfArticulos.TextMatrix(1, cintColVsfIdArticulo)) <> 0 Then
        If lblnPermisoAsignar Then cmdMarcar_Click
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":vsfArticulos_DblClick"))
End Sub

Private Sub vsfArticulos_GotFocus()
On Error GoTo NotificaError
    
    pHabilitaConsultaReq

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":vsfArticulos_GotFocus"))
End Sub

Private Sub vsfArticulos_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
On Error GoTo NotificaError

    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
        KeyAscii = 7
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":vsfArticulos_KeyPressEdit"))
End Sub

Private Sub vsfArticulos_RowColChange()
On Error GoTo NotificaError

    lblNombreComercialCompleto.Caption = vsfArticulos.TextMatrix(vsfArticulos.Row, cintColVsfNombreComercial)
    If lblNombreComercialCompleto.Caption = "Nombre" Then
        lblNombreComercialCompleto.Caption = ""
    Else
    
    End If

    pHabilitaConsultaReq
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":vsfArticulos_RowColChange"))
End Sub

Private Sub VsfBusArticulos_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    
    If Val(VsfBusArticulos.TextMatrix(Row, Col)) = 0 Then
        MsgBox SIHOMsg(452), vbOKOnly + vbInformation, "Mensaje"
        VsfBusArticulos.TextMatrix(Row, Col) = VsfBusArticulos.TextMatrix(Row, cintColBusCantidadOriginal)
        VsfBusArticulos.Editable = 0
    ElseIf Val(VsfBusArticulos.TextMatrix(Row, Col)) > Val(VsfBusArticulos.TextMatrix(Row, cintColBusCantidadOriginal)) Then
        MsgBox SIHOMsg(965), vbOKOnly + vbInformation, "Mensaje"
        VsfBusArticulos.TextMatrix(Row, Col) = VsfBusArticulos.TextMatrix(Row, cintColBusCantidadOriginal)
        VsfBusArticulos.Editable = 0
    End If
    
End Sub

Private Sub VsfBusArticulos_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    cboUnidad.Visible = False
End Sub

Private Sub VsfBusArticulos_AfterSort(ByVal Col As Long, Order As Integer)
    Select Case Col
        Case 1
            vlstrColSort = Order & "cintColBusClave"
        Case 2
            vlstrColSort = Order & "cintColBusNombre"
        Case 3
            vlstrColSort = Order & "cintColBusLocalizacion"
        Case 4
            vlstrColSort = Order & "cintColBusAlmacen"
        Case 5
            vlstrColSort = Order & "cintColBusIdAlmacen"
        Case 6
            vlstrColSort = Order & "cintColBusExist"
        Case 7
            vlstrColSort = Order & "cintColBusUnidadExist"
        Case 8
            vlstrColSort = Order & "cintColBusPendiente"
        Case 9
            vlstrColSort = Order & "cintColBusUnidadPendiente"
        Case 10
            vlstrColSort = Order & "cintColBusCantidad"
        Case 11
            vlstrColSort = Order & "cintColBusUnidad"
        Case 12
            vlstrColSort = Order & "cintColBusTipoUnidad"
        Case 13
            vlstrColSort = Order & "cintColBusCveDeptoAutoriza"
        Case 14
            vlstrColSort = Order & "cintColBusCveDeptoRecibe"
        Case 15
            vlstrColSort = Order & "cintColBusCantidadOriginal"
        Case 16
            vlstrColSort = Order & "cintColsVsfBusArticulos"
    End Select
End Sub

Private Sub VsfBusArticulos_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo NotificaError

    If Not Col = cintColBusCantidad Then
        Cancel = True
    Else
        Cancel = False
        If Not IsNumeric(VsfBusArticulos.TextMatrix(Row, Col)) Then
            liTemp = 0
        Else
            liTemp = VsfBusArticulos.TextMatrix(Row, Col)
            liTemp2 = liTemp
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":VsfBusArticulos_BeforeEdit"))
End Sub

Private Sub VsfBusArticulos_Click()
      Dim rsContenido As ADODB.Recordset
      Dim strParametros As String
      Dim strUnidadMinima As String
      Dim strUnidadAlterna As String
      Dim rsUnidadVenta As ADODB.Recordset
1     On Error GoTo NotificaError

2         If cboTipoRequisicion.Text <> "COMPRA - PEDIDO" And cboTipoRequisicion.Text <> "CONSIGNACION" And cmdImprimir.Enabled = False Then
3             Set rsUnidadVenta = frsEjecuta_SP(VsfBusArticulos.TextMatrix(VsfBusArticulos.Row, cintColBusClave) & "|" & "|", "Sp_Ivselarticulo")
4             If rsUnidadVenta.RecordCount > 0 Then
5                 strUnidadMinima = rsUnidadVenta!UnidadMinima
6                 strUnidadAlterna = rsUnidadVenta!UnidadAlterna
7             End If
8             rsUnidadVenta.Close
              
9             If VsfBusArticulos.Row > -1 Then
10                cboUnidad.Clear
11                cboUnidad.AddItem IIf(VsfBusArticulos.TextMatrix(VsfBusArticulos.Row, cintColBusTipoUnidad) = "A", strUnidadMinima, strUnidadAlterna)
12                strUnidad = IIf(VsfBusArticulos.TextMatrix(VsfBusArticulos.Row, cintColBusTipoUnidad) = "A", "Unidad mínima", "Unidad alterna")
13                If VsfBusArticulos.Col = cintColBusUnidad Then
14                    If VsfBusArticulos.TextMatrix(VsfBusArticulos.Row, cintColBusUnidad) <> "" Then
15                        Set rsContenido = frsEjecuta_SP(VsfBusArticulos.TextMatrix(VsfBusArticulos.Row, cintColBusClave), "SP_IVSELCONTENIDOARTICULO")
                          'Valida que el artículo tenga unidad alterna y unidad mínima para mostrar el combobox
16                        If rsContenido.RecordCount > 0 And rsContenido(0).Value > 1 Then
17                            cboUnidad.Visible = True
18                            VsfBusArticulos.RowHeight(VsfBusArticulos.Row) = cboUnidad.height
19                            cboUnidad.width = VsfBusArticulos.ColWidth(cintColBusUnidad)
20                            cboUnidad.Top = VsfBusArticulos.CellTop
21                            cboUnidad.Left = VsfBusArticulos.CellLeft
22                        End If
23                        rsContenido.Close
24                    End If
25                Else
26                    cboUnidad.Visible = False
27                    VsfBusArticulos.RowHeight(VsfBusArticulos.AutoResize) = 240
28                End If
29            End If
30        End If
          
31    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":VsfBusArticulos_Click" & " Linea:" & Erl()))
End Sub

Private Sub pConvierteUnidad()
      Dim rsContenido As ADODB.Recordset
      Dim lngCantidadFaltante As Long
      Dim lngPendientes As Long
      Dim strUnidadMinima As String
      Dim strUnidadAlterna As String
      Dim rsUnidadVenta As ADODB.Recordset
1     On Error GoTo NotificaError

2         If VsfBusArticulos.TextMatrix(VsfBusArticulos.Row, cintColBusUnidad) <> "" Then
3             Set rsUnidadVenta = frsEjecuta_SP(VsfBusArticulos.TextMatrix(VsfBusArticulos.Row, cintColBusClave) & "|" & "|", "Sp_Ivselarticulo")
4             If rsUnidadVenta.RecordCount > 0 Then
5                 strUnidadMinima = rsUnidadVenta!UnidadMinima
6                 strUnidadAlterna = rsUnidadVenta!UnidadAlterna
7             End If
8             rsUnidadVenta.Close

9             Set rsContenido = frsEjecuta_SP(VsfBusArticulos.TextMatrix(VsfBusArticulos.Row, cintColBusClave), "SP_IVSELCONTENIDOARTICULO")
10            lngCantidadFaltante = 1
11            lngPendientes = flngCantidadPendiente(VsfBusArticulos.TextMatrix(VsfBusArticulos.Row, cintColBusClave), cboDepartamento.ItemData(cboDepartamento.ListIndex))
12            frsEjecuta_SP VsfBusArticulos.TextMatrix(VsfBusArticulos.Row, cintColBusClave) & "|" & cboDepartamento.ItemData(cboDepartamento.ListIndex), "SP_IVINTCANTIDADFALTANTE", True, lngCantidadFaltante
              
13            If strUnidad = "Unidad mínima" Then
14                VsfBusArticulos.TextMatrix(VsfBusArticulos.Row, cintColBusCantidad) = lngCantidadFaltante - lngPendientes
15                VsfBusArticulos.TextMatrix(VsfBusArticulos.Row, cintColBusUnidad) = strUnidadMinima
16                VsfBusArticulos.TextMatrix(VsfBusArticulos.Row, cintColBusTipoUnidad) = "M"
17                VsfBusArticulos.TextMatrix(VsfBusArticulos.Row, cintColBusCantidadOriginal) = lngCantidadFaltante - lngPendientes
18            Else
19                If lngCantidadFaltante < rsContenido(0).Value Then
20                    VsfBusArticulos.TextMatrix(VsfBusArticulos.Row, cintColBusCantidad) = Int(((lngCantidadFaltante - lngPendientes) * rsContenido(0).Value) / rsContenido(0).Value)
21                    VsfBusArticulos.TextMatrix(VsfBusArticulos.Row, cintColBusCantidadOriginal) = Int(((lngCantidadFaltante - lngPendientes) * rsContenido(0).Value) / rsContenido(0).Value)
22                Else
23                    VsfBusArticulos.TextMatrix(VsfBusArticulos.Row, cintColBusCantidad) = Int((lngCantidadFaltante - lngPendientes) / rsContenido(0).Value)
24                    VsfBusArticulos.TextMatrix(VsfBusArticulos.Row, cintColBusCantidadOriginal) = Int((lngCantidadFaltante - lngPendientes) / rsContenido(0).Value)
25                End If
26                VsfBusArticulos.TextMatrix(VsfBusArticulos.Row, cintColBusUnidad) = strUnidadAlterna
27                VsfBusArticulos.TextMatrix(VsfBusArticulos.Row, cintColBusTipoUnidad) = "A"
28            End If
29            cboUnidad.Visible = False
30            VsfBusArticulos.RowHeight(VsfBusArticulos.AutoResize) = 240
31        End If
32        cboUnidad.Visible = False
33        VsfBusArticulos.RowHeight(VsfBusArticulos.AutoResize) = 240

34    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pConvierteUnidad" & " Linea:" & Erl()))
End Sub
Private Sub pMuestraRequisicion(strOrden As String)
Dim lngContador As Long
    'Pasar al grdRequisicion:
            For lngContador = 1 To VsfBusArticulos.Rows - 1
               If strOrden <> VsfBusArticulos.TextMatrix(lngContador, cintColBusClave) Then
                grdRequisicion.TextMatrix(grdRequisicion.Rows - 1, cintColGrdCveDeptoSurte) = VsfBusArticulos.TextMatrix(lngContador, cintColBusIdAlmacen)
                grdRequisicion.TextMatrix(grdRequisicion.Rows - 1, cintColGrdCveArticulo) = VsfBusArticulos.TextMatrix(lngContador, cintColBusClave)
                grdRequisicion.TextMatrix(grdRequisicion.Rows - 1, cintColGrdCantidad) = VsfBusArticulos.TextMatrix(lngContador, cintColBusCantidad)
                grdRequisicion.TextMatrix(grdRequisicion.Rows - 1, cintColGrdTipoUnidad) = VsfBusArticulos.TextMatrix(lngContador, cintColBusTipoUnidad)
                grdRequisicion.TextMatrix(grdRequisicion.Rows - 1, cintColGrdCveDeptoAutoriza) = VsfBusArticulos.TextMatrix(lngContador, cintColBusCveDeptoAutoriza)
                grdRequisicion.TextMatrix(grdRequisicion.Rows - 1, cintColGrdCveDeptoRecibe) = VsfBusArticulos.TextMatrix(lngContador, cintColBusCveDeptoRecibe)
                If strClavesConsol <> "" Then
                    grdRequisicion.TextMatrix(grdRequisicion.Rows - 1, cintColReqArtiExisEsclavos) = VsfBusArticulos.TextMatrix(lngContador, cintColBusExisEsclavos)
                    grdRequisicion.TextMatrix(grdRequisicion.Rows - 1, cintColReqArtiMaxEsclavos) = VsfBusArticulos.TextMatrix(lngContador, cintColBusMaxEsclavos)
                    grdRequisicion.TextMatrix(grdRequisicion.Rows - 1, cintColReqArtiDeptoEsclavos) = VsfBusArticulos.TextMatrix(lngContador, cintColBusDeptoEsclavos)
                End If
                grdRequisicion.Rows = grdRequisicion.Rows + 1
               End If
            Next lngContador
End Sub

Private Sub VsfBusArticulos_DblClick()
On Error GoTo NotificaError
    Dim CurrRow As Long
    Dim CurrCol As Long
    Dim Cancel As Boolean
    
    CurrRow = VsfBusArticulos.Row
    CurrCol = VsfBusArticulos.Col

    If Not CurrCol = cintColBusCantidad Then
        Cancel = True
    Else
        Cancel = False
        liTemp2 = VsfBusArticulos.TextMatrix(CurrRow, CurrCol)
        VsfBusArticulos.Editable = flexEDKbdMouse
    End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":VsfBusArticulos_DblClick"))
 End Sub

Private Sub VsfBusArticulos_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
On Error GoTo NotificaError

    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then KeyAscii = 7

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":VsfBusArticulos_KeyPressEdit"))
End Sub

Private Sub pHabilitaConsultaReq()
On Error GoTo NotificaError

    cmdConsultarReq.Enabled = Val(vsfArticulos.TextMatrix(vsfArticulos.Row, cintColVsfIdArticulo)) <> 0 And lstrModo = "C" And Trim(vsfArticulos.TextMatrix(vsfArticulos.Row, cintColvsfEstadoArt)) = cstrSolicitado

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pHabilitaConsultaReq"))
End Sub

Private Sub VsfBusArticulos_LeaveCell()
    VsfBusArticulos.RowHeight(VsfBusArticulos.AutoResize) = 240
    cboUnidad.Visible = False
End Sub
Private Function fCveArticuloMaxMin(strIdArticulo As String) As String
On Error GoTo NotificaError
    Dim strSentencia   As String

     fCveArticuloMaxMin = ""
     
     strSentencia = " SELECT NVL(INTIDARTICULO,'') AS INTIDARTICULO" & _
                    "  From IVARTICULO " & _
                    "  Where (CHRCVEARTICULO = " & CStr(strIdArticulo) & ")"
                    
     Set rs = frsRegresaRs(strSentencia, adLockReadOnly, adOpenForwardOnly)
     
     If rs.RecordCount <> 0 Then
      fCveArticuloMaxMin = rs!intIdArticulo
     End If
     
Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fCveArticuloMaxMin"))
End Function
Private Function fCveAlmaDeptoReCo(strCveDepto, strAlmaDeptoReq As String, TipoMov As String) As String
On Error GoTo NotificaError
    Dim strSentencia   As String

     fCveAlmaDeptoReCo = "0"
     
     strSentencia = " SELECT Distinct NODEPARTAMENTO.SMICVEDEPARTAMENTO, NODEPARTAMENTO.VCHDESCRIPCION, NODEPARTAMENTO.BITCONSIGNACION " & _
                    " From IVREQUISICIONDEPARTAMENTO " & _
                    " INNER JOIN LOGIN ON IVREQUISICIONDEPARTAMENTO.INTNUMEROLOGIN = LOGIN.INTNUMEROLOGIN " & _
                    " INNER JOIN NODEPARTAMENTO ON IVREQUISICIONDEPARTAMENTO.SMICVEDEPARTAMENTO = NODEPARTAMENTO.SMICVEDEPARTAMENTO " & _
                    " Where IVREQUISICIONDEPARTAMENTO.CHRTIPOREQUISICION = TRIM('" & TipoMov & "') AND LOGIN.SMICVEDEPARTAMENTO = " & strCveDepto & " " & _
                    " AND  TRIM(NODEPARTAMENTO.VCHDESCRIPCION) = '" & Trim(strAlmaDeptoReq) & "'"
                    
     Set rs = frsRegresaRs(strSentencia, adLockReadOnly, adOpenForwardOnly)
     
     If rs.RecordCount <> 0 Then
      fCveAlmaDeptoReCo = Str(rs!smicvedepartamento) & "," & Str(rs!BITCONSIGNACION)
      
     End If
     
Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fCveAlmaDeptoReCo"))
End Function
Private Function fAlmaComprasConf(strCveDepto, TipoMov As String) As String
' S existe almacen configurado de compras N no existe
On Error GoTo NotificaError
    Dim strSentencia   As String

     fAlmaComprasConf = "N"
     
     strSentencia = " SELECT Distinct NODEPARTAMENTO.SMICVEDEPARTAMENTO, NODEPARTAMENTO.VCHDESCRIPCION " & _
                    " From IVREQUISICIONDEPARTAMENTO " & _
                    " INNER JOIN LOGIN ON IVREQUISICIONDEPARTAMENTO.INTNUMEROLOGIN = LOGIN.INTNUMEROLOGIN " & _
                    " INNER JOIN NODEPARTAMENTO ON IVREQUISICIONDEPARTAMENTO.SMICVEDEPARTAMENTO = NODEPARTAMENTO.SMICVEDEPARTAMENTO " & _
                    " Where IVREQUISICIONDEPARTAMENTO.CHRTIPOREQUISICION = TRIM('" & TipoMov & "') AND LOGIN.SMICVEDEPARTAMENTO = " & strCveDepto
                              
     Set rs = frsRegresaRs(strSentencia, adLockReadOnly, adOpenForwardOnly)
     
     If rs.RecordCount <> 0 Then
      fAlmaComprasConf = "S"
     End If
     
Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fAlmaComprasConf"))
End Function


