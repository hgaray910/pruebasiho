VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.1#0"; "MyCommandButton.ocx"
Object = "{FF14BD24-9F8A-41E3-B5B8-7F0D45EE9F16}#15.0#0"; "hsflatcontrols.ocx"
Begin VB.Form frmMantoListasPrecios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listas de precios"
   ClientHeight    =   9930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14055
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9930
   ScaleMode       =   0  'User
   ScaleWidth      =   14055
   StartUpPosition =   2  'CenterScreen
   Begin HSFlatControls.MyCombo cboTipoCargo 
      Height          =   375
      Left            =   1800
      TabIndex        =   45
      Top             =   3000
      Width           =   5640
      _ExtentX        =   9948
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
      TabIndex        =   48
      Top             =   3810
      Width           =   5640
      _ExtentX        =   9948
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
   Begin SysInfoLib.SysInfo SysInfo1 
      Left            =   2040
      Top             =   10200
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Frame freBarra 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   810
      Left            =   70
      TabIndex        =   38
      Top             =   9840
      Visible         =   0   'False
      Width           =   13890
      Begin MSComctlLib.ProgressBar pgbBarra 
         Height          =   300
         Left            =   45
         TabIndex        =   39
         Top             =   480
         Width           =   13795
         _ExtentX        =   24342
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lblTextoBarra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Cargando datos, por favor espere..."
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   90
         TabIndex        =   40
         Top             =   180
         Width           =   11250
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   345
         Left            =   25
         Top             =   120
         Width           =   11000
      End
   End
   Begin VB.Frame cmdImprimir2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6135
      Left            =   70
      TabIndex        =   17
      Top             =   2760
      Width           =   13890
      Begin VB.ListBox lstPesos 
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
         Height          =   540
         ItemData        =   "frmMantoListasPrecios.frx":0000
         Left            =   7800
         List            =   "frmMantoListasPrecios.frx":000A
         TabIndex        =   70
         Top             =   2040
         Visible         =   0   'False
         Width           =   1000
      End
      Begin VB.CheckBox chkTabulador 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Utiliza tabulador"
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
         Height          =   255
         Left            =   3000
         TabIndex        =   60
         ToolTipText     =   "Permite capturar el detalle de la lista de precios"
         Top             =   4995
         Width           =   2415
      End
      Begin VB.CheckBox chkIncremento 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Incremento automático"
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
         Height          =   255
         Left            =   120
         TabIndex        =   69
         Top             =   4995
         Width           =   2655
      End
      Begin HSFlatControls.MyCombo cboConceptoFacturacion 
         Height          =   375
         Left            =   9705
         TabIndex        =   51
         Top             =   645
         Width           =   4065
         _ExtentX        =   7170
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
      Begin HSFlatControls.MyCombo cboClasificacionSA 
         Height          =   375
         Left            =   9705
         TabIndex        =   50
         Top             =   240
         Width           =   4065
         _ExtentX        =   7170
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
      Begin HSFlatControls.MyCombo cboSubFamilia 
         Height          =   375
         Left            =   1730
         TabIndex        =   49
         Top             =   1455
         Width           =   5640
         _ExtentX        =   9948
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
      Begin VB.TextBox txtClaveArticulo 
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
         Left            =   5890
         MaxLength       =   10
         TabIndex        =   47
         ToolTipText     =   "Clave del artículo"
         Top             =   650
         Width           =   1470
      End
      Begin VB.Timer tmrDespliega 
         Interval        =   1000
         Left            =   6240
         Top             =   2640
      End
      Begin VB.TextBox txtIniciales 
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
         Left            =   9705
         MaxLength       =   20
         TabIndex        =   52
         Top             =   1045
         Width           =   4065
      End
      Begin VB.ListBox UpDown1 
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
         Height          =   795
         ItemData        =   "frmMantoListasPrecios.frx":001E
         Left            =   480
         List            =   "frmMantoListasPrecios.frx":002B
         TabIndex        =   58
         Top             =   2280
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.TextBox txtPrecio 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   360
         MaxLength       =   15
         TabIndex        =   29
         Top             =   1920
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.Frame fraModificar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Modificar listas"
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
         Height          =   735
         Left            =   6863
         TabIndex        =   21
         Top             =   5280
         Width           =   6900
         Begin MyCommandButton.MyButton cmdMas 
            Height          =   375
            Left            =   3240
            TabIndex        =   26
            ToolTipText     =   "Iniciar el incremento de la lista en el porcentaje seleccionado"
            Top             =   255
            Width           =   1170
            _ExtentX        =   2064
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
            TransparentColor=   15790320
            Caption         =   "Aumentar"
            DepthEvent      =   1
            ShowFocus       =   -1  'True
         End
         Begin VB.OptionButton optModificar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Utilidad"
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
            Left            =   3960
            TabIndex        =   23
            Top             =   0
            Width           =   1095
         End
         Begin VB.OptionButton optModificar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Precio de venta"
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
            Left            =   1920
            TabIndex        =   22
            Top             =   0
            Value           =   -1  'True
            Width           =   1935
         End
         Begin VB.TextBox txtPorciento 
            Alignment       =   1  'Right Justify
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
            Left            =   1965
            MaxLength       =   5
            TabIndex        =   24
            ToolTipText     =   "Porcentaje de aumento o disminución de la lista"
            Top             =   250
            Width           =   615
         End
         Begin MyCommandButton.MyButton cmdInvertir 
            Height          =   375
            Left            =   120
            TabIndex        =   59
            ToolTipText     =   "Permite capturar el detalle de la lista de precios"
            Top             =   255
            Width           =   1815
            _ExtentX        =   3201
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
            TransparentColor=   15790320
            Caption         =   "Invertir selección"
            DepthEvent      =   1
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdMenos 
            Height          =   375
            Left            =   4440
            TabIndex        =   27
            ToolTipText     =   "Iniciar el decremento de la lista con el porcentaje seleccionado"
            Top             =   255
            Width           =   1035
            _ExtentX        =   1826
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
            TransparentColor=   15790320
            Caption         =   "Disminuir"
            DepthEvent      =   1
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdRec 
            Height          =   375
            Left            =   5520
            TabIndex        =   28
            ToolTipText     =   "Iniciar el decremento de la lista con el porcentaje seleccionado"
            Top             =   255
            Width           =   1185
            _ExtentX        =   2090
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
            TransparentColor=   15790320
            Caption         =   "Recalcular"
            DepthEvent      =   1
            ShowFocus       =   -1  'True
         End
         Begin VB.Label lblPorciento 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "%"
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
            Height          =   255
            Left            =   2640
            TabIndex        =   25
            Top             =   315
            Width           =   150
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Inicializar precios de "
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
         Height          =   735
         Left            =   135
         TabIndex        =   18
         Top             =   5280
         Width           =   6675
         Begin HSFlatControls.MyCombo cboOtrasListas 
            Height          =   375
            Left            =   120
            TabIndex        =   19
            ToolTipText     =   "Selcción de la lista de precios con que se desea inicializar"
            Top             =   255
            Width           =   5250
            _ExtentX        =   9260
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
         Begin MyCommandButton.MyButton cmdInicializa 
            Height          =   375
            Left            =   5475
            TabIndex        =   20
            ToolTipText     =   "Inicializar la lista de precios tomando como base la lista seleccionada"
            Top             =   255
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            BackColor       =   -2147483633
            Enabled         =   0   'False
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
            TransparentColor=   15790320
            Caption         =   "Inicializar"
            DepthEvent      =   1
            ShowFocus       =   -1  'True
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdPrecios 
         Height          =   3045
         Left            =   120
         TabIndex        =   16
         ToolTipText     =   "Captura de los precios"
         Top             =   1860
         Width           =   13665
         _ExtentX        =   24104
         _ExtentY        =   5371
         _Version        =   393216
         ForeColorSel    =   16777215
         BackColorBkg    =   -2147483643
         GridColor       =   -2147483638
         GridColorFixed  =   -2147483638
         GridColorUnpopulated=   -2147483638
         WordWrap        =   -1  'True
         HighLight       =   2
         GridLinesFixed  =   1
         GridLinesUnpopulated=   1
         Appearance      =   0
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
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin HSFlatControls.MyCombo cboArtMed 
         Height          =   375
         Left            =   1730
         TabIndex        =   46
         Top             =   650
         Width           =   2500
         _ExtentX        =   4419
         _ExtentY        =   661
         Style           =   1
         Enabled         =   -1  'True
         Text            =   "MyCombo2"
         Sorted          =   0   'False
         List            =   $"frmMantoListasPrecios.frx":0069
         ItemData        =   $"frmMantoListasPrecios.frx":008F
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
      Begin MyCommandButton.MyButton cmdFiltrar 
         Height          =   375
         Left            =   10370
         TabIndex        =   53
         ToolTipText     =   "Permite capturar el detalle de la lista de precios"
         Top             =   1455
         Width           =   1095
         _ExtentX        =   1931
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
         TransparentColor=   14215660
         Caption         =   "Filtrar"
         DepthEvent      =   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdExportar 
         Height          =   375
         Left            =   11520
         TabIndex        =   54
         ToolTipText     =   "Permite exportar la información filtrada en formato Excel"
         Top             =   1455
         Width           =   1095
         _ExtentX        =   1931
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
         TransparentColor=   14215660
         Caption         =   "Exportar"
         DepthEvent      =   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdImportar 
         Height          =   375
         Left            =   12675
         TabIndex        =   55
         ToolTipText     =   "Permite importar la información filtrada desde un formato Excel"
         Top             =   1455
         Width           =   1095
         _ExtentX        =   1931
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
         TransparentColor=   15790320
         Caption         =   "Importar"
         DepthEvent      =   1
         ShowFocus       =   -1  'True
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   9000
         Top             =   1320
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label lblClaveArticulo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Clave artículo"
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
         Height          =   255
         Left            =   4330
         TabIndex        =   66
         Top             =   720
         Width           =   1380
      End
      Begin VB.Label lblDescripcionCargo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Descripción cargo"
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
         Height          =   255
         Left            =   7515
         TabIndex        =   65
         Top             =   1110
         Width           =   1770
      End
      Begin VB.Label lblClasificacion 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Clasificación "
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
         Height          =   255
         Left            =   7515
         TabIndex        =   64
         Top             =   300
         Width           =   1305
      End
      Begin VB.Label lblFamilia 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   63
         Top             =   1110
         Width           =   690
      End
      Begin VB.Label lblSubFamilia 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   135
         TabIndex        =   62
         Top             =   1510
         Width           =   1005
      End
      Begin VB.Label lblTipoArticulo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Tipo artículo"
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
         Height          =   255
         Left            =   120
         TabIndex        =   61
         Top             =   700
         Width           =   1185
      End
      Begin VB.Label lblTipoCargo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Tipo de cargo"
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
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   300
         Width           =   1350
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Concepto de factura"
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
         Height          =   255
         Left            =   7515
         TabIndex        =   56
         Top             =   705
         Width           =   2085
      End
   End
   Begin TabDlg.SSTab sstListas 
      Height          =   10995
      Left            =   -45
      TabIndex        =   30
      Top             =   -570
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   19394
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmMantoListasPrecios.frx":009C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "freDatos"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmMantoListasPrecios.frx":00B8
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame6"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Descripción completa del cargo"
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
         Height          =   915
         Left            =   120
         TabIndex        =   42
         Top             =   9480
         Width           =   13890
         Begin VB.Label lblNombreCompleto 
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
            Height          =   510
            Left            =   135
            TabIndex        =   43
            Top             =   270
            Width           =   13635
         End
      End
      Begin VB.Frame Frame6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2730
         Left            =   -74860
         TabIndex        =   35
         Top             =   570
         Width           =   13870
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdHBusqueda 
            Height          =   2175
            Left            =   60
            TabIndex        =   36
            Top             =   165
            Width           =   13760
            _ExtentX        =   24262
            _ExtentY        =   3836
            _Version        =   393216
            ForeColorSel    =   16777215
            BackColorBkg    =   -2147483643
            GridColor       =   -2147483638
            GridColorFixed  =   -2147483638
            GridColorUnpopulated=   -2147483638
            GridLinesFixed  =   1
            GridLinesUnpopulated=   1
            Appearance      =   0
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
            _Band(0).Cols   =   2
         End
         Begin VB.Label lblPredeterminada 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Predeterminada"
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
            Height          =   255
            Left            =   60
            TabIndex        =   44
            Top             =   2400
            Width           =   1575
         End
      End
      Begin VB.Frame freDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1875
         Left            =   120
         TabIndex        =   32
         Top             =   600
         Width           =   13890
         Begin HSFlatControls.MyCombo cboTabulador 
            Height          =   375
            Left            =   1730
            TabIndex        =   68
            ToolTipText     =   "Tabulador aplicable"
            Top             =   1050
            Width           =   8800
            _ExtentX        =   15531
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
            Left            =   1730
            TabIndex        =   0
            Top             =   240
            Width           =   8800
            _ExtentX        =   15531
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
         Begin VB.CheckBox Chkbitcheckup 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Solo servicios para check up"
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
            Left            =   10750
            TabIndex        =   5
            ToolTipText     =   "Mostrar solamente los servicios seleccionados para check up"
            Top             =   700
            Width           =   3040
         End
         Begin VB.CheckBox chkImprimeCargosPrecioCero 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Imprimir con precio en cero"
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
            Left            =   10750
            TabIndex        =   6
            ToolTipText     =   "Imprimir cargos con precio en cero"
            Top             =   1110
            Value           =   1  'Checked
            Width           =   3000
         End
         Begin VB.CheckBox chkPredeterminada 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Lista predeterminada"
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
            Left            =   10750
            TabIndex        =   4
            Top             =   300
            Width           =   2385
         End
         Begin VB.CheckBox chkActivo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Activo"
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
            Left            =   1730
            TabIndex        =   3
            ToolTipText     =   "Estatus de activa o inactiva de la lista"
            Top             =   1500
            Width           =   960
         End
         Begin VB.TextBox txtClave 
            Alignment       =   1  'Right Justify
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
            Left            =   1730
            MaxLength       =   5
            TabIndex        =   1
            ToolTipText     =   "Clave de la lista de precios"
            Top             =   640
            Width           =   1215
         End
         Begin VB.TextBox txtDescripcion 
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
            Left            =   4970
            TabIndex        =   2
            ToolTipText     =   "Descripción de la lista de precios"
            Top             =   640
            Width           =   5560
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Tabulador aplicable"
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
            TabIndex        =   67
            Top             =   1110
            Width           =   1455
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
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
            ForeColor       =   &H80000008&
            Height          =   250
            Left            =   120
            TabIndex        =   41
            Top             =   300
            Width           =   1455
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Estado"
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
            Top             =   1510
            Width           =   660
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Clave "
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
            TabIndex        =   34
            Top             =   700
            Width           =   645
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Descripción"
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
            Height          =   255
            Left            =   3600
            TabIndex        =   33
            Top             =   700
            Width           =   1125
         End
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   780
         Left            =   3705
         TabIndex        =   31
         Top             =   2480
         Width           =   6735
         Begin MyCommandButton.MyButton cmdCapturaPrecios 
            Height          =   600
            Left            =   60
            TabIndex        =   7
            ToolTipText     =   "Permite capturar el detalle de la lista de precios"
            Top             =   130
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
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   "Capturar precios"
            DepthEvent      =   1
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdPrimerRegistro 
            Height          =   600
            Left            =   1870
            TabIndex        =   8
            ToolTipText     =   "Primer registro"
            Top             =   130
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
            Picture         =   "frmMantoListasPrecios.frx":00D4
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmMantoListasPrecios.frx":0A56
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdAnteriorRegistro 
            Height          =   600
            Left            =   2470
            TabIndex        =   9
            ToolTipText     =   "Anterior registro"
            Top             =   130
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
            Picture         =   "frmMantoListasPrecios.frx":13D8
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmMantoListasPrecios.frx":1D5A
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdBuscar 
            Height          =   600
            Left            =   3070
            TabIndex        =   10
            ToolTipText     =   "Búsqueda"
            Top             =   130
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
            Picture         =   "frmMantoListasPrecios.frx":26DC
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmMantoListasPrecios.frx":3060
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdSiguienteRegistro 
            Height          =   600
            Left            =   3670
            TabIndex        =   11
            ToolTipText     =   "Siguiente registro"
            Top             =   130
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
            Picture         =   "frmMantoListasPrecios.frx":39E4
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmMantoListasPrecios.frx":4366
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdUltimoRegistro 
            Height          =   600
            Left            =   4270
            TabIndex        =   12
            ToolTipText     =   "Último registro"
            Top             =   130
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
            Picture         =   "frmMantoListasPrecios.frx":4CE8
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmMantoListasPrecios.frx":566A
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdGrabarRegistro 
            Height          =   600
            Left            =   4870
            TabIndex        =   13
            ToolTipText     =   "Guardar el registro"
            Top             =   130
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
            Picture         =   "frmMantoListasPrecios.frx":5FEC
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmMantoListasPrecios.frx":6970
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdDelete 
            Height          =   600
            Left            =   5470
            TabIndex        =   14
            ToolTipText     =   "Borrar la clasificación"
            Top             =   130
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
            Picture         =   "frmMantoListasPrecios.frx":72F4
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmMantoListasPrecios.frx":7C76
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdImprimir 
            Height          =   600
            Left            =   6070
            TabIndex        =   15
            ToolTipText     =   "Permite imprimir el presupuesto"
            Top             =   130
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
            Picture         =   "frmMantoListasPrecios.frx":85FA
            BackColorDown   =   -2147483643
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmMantoListasPrecios.frx":8F7E
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
      End
   End
End
Attribute VB_Name = "frmMantoListasPrecios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------
'| Nombre del Proyecto      : Caja
'| Nombre del Formulario    : frmMantoListasPrecios
'-------------------------------------------------------------------------------------
'| Objetivo: Realiza el mantenimiento de listas de precios
'-------------------------------------------------------------------------------------
'| Análisis y Diseño        : Rodolfo Ramos G.
'| Autor                    : Rodolfo Ramos G.
'| Fecha de Creación        : 26/Diciembre/2000
'| Modificó                 : Nombre(s)
'| Fecha terminación        : 11/Enero/2001
'| Fecha última modificación: dd/mes/AAAA
'-------------------------------------------------------------------------------------
Option Explicit

'Index del cbo cboTipoCargo
Const cintIndexTodos = 0
Const cintIndexArticulo = 1
Const cintIndexEstudio = 2
Const cintIndexExamen = 3
Const cintIndexExamenGrupo = 4
Const cintIndexGrupo = 5
Const cintIndexOtro = 6
Const cintIndexPaquete = 7

'Columnas del grid de precios:
Const cintColClave = 1
Const cintColDescripcion = 2

Const cintColIncremetoAutomatico = 3
Const cintColTipoIncremento = 4
Const cintColUtilidad = 5
Const cintColTabulador = 6
Const cintColCosto = 7

Const cintColPrecio = 8
Const cintColMoneda = 9
Const cintColTipo = 10
Const cintColTipoDes = 11
Const cintColModificado = 12
Const cintColNuevo = 13
Const cintColCveFact = 14
Const cintColFacturacion = 15
Const cintColCostoUltimaEntrada = 16
Const cintColCostoMasAlto = 17
Const cintColPrecioMaximopublico = 18
Const cintColPrecioEspecifico = 19
Const cintColumnas = 20

Const clngRojo = &HC0&
Const clngAzul = &HC00000

Const cstrUltimaCompra = "ÚLTIMA COMPRA"
Const cstrCompraMasAlta = "COMPRA MÁS ALTA"
Const cstrPrecioMaximoPublico = "PRECIO MÁXIMO AL PÚBLICO"

Const clngLargoRenglon = 240

Public WithEvents grid As MSHFlexGrid
Attribute grid.VB_VarHelpID = -1
Public WithEvents txtEdit As TextBox
Attribute txtEdit.VB_VarHelpID = -1
Public vpblnEsSistemas As Boolean

Private vgrptReporte As CRAXDRT.report
Dim vgblnNoEditar As Boolean

Dim vgblnEditaPrecio As Boolean 'Para saber si se esta editando una cantidad
Dim vgblnEditaUtilidad As Boolean 'Para saber si se esta editando el margen de utilidad
Dim vgstrEstadoManto As String 'Estatus para saber donde ando en la pantalla
Dim vglngDesktop As Long     'Para saber el tamaño del desktop
Dim vlstrSentencia As String
Dim vlblnCancelCaptura As Boolean  ' Para no permitir capturar si no hay elementos en la lista
Dim vlblnRecalcularImportar As Boolean  ' Para saber si se esta importando

Dim lblnConsulta As Boolean
Dim llngMarcados As Long
Dim blnactivainicializa As Integer
Dim vlstrCveCargo As String

Dim rsListasDepto As New ADODB.Recordset 'Listas de precios del departamento

'-----------------------------------
' vgstrEstadoManto puede tener los siguientes valores
' ""  Nuevo registro, default
' "A" Una Alta de un Elemento
' "B" Que esta en la pantalla de búsqueda
' "M" Una modificacion o Consulta
' "ME" Esta editando un precio en una consulta
' "AE" Esta editando un precio en una alta
'-----------------------------------
Const cgintFactorMovVentana = 150
Const cgIntAltoVentanaMax = 10410
Const cgIntAltoVentanaMin = 3240

Public blnCatalogo As Boolean
Public StrCveArticulo As String
Dim lblnPermisoCosto As Boolean 'tiene permiso para modificar el costo

Private Sub cboArtMed_Click()
    On Error GoTo NotificaError

    Dim rs As ADODB.Recordset
    
    txtIniciales.Text = ""
    pconfiguragrid
    
    cboFamilia.Clear
    If cboArtMed.ListIndex > 0 Then
        '0 = Articulos
        '1 = Medicamentos
        Set rs = frsEjecuta_SP(IIf(cboArtMed.ItemData(cboArtMed.ListIndex) = 1, 0, 1), "SP_IVSELFAMILIA")
        pLlenarCboRs_new cboFamilia, rs, 0, 1
        rs.Close
    End If
    cboFamilia.AddItem "<TODOS>", 0
    cboFamilia.ItemData(cboFamilia.NewIndex) = 0
    cboFamilia.ListIndex = 0
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboArtMed_Click"))
End Sub

Private Sub cboArtMed_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    If KeyCode = vbKeyReturn Then SendKeys vbTab

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboArtMed_KeyDown"))
End Sub

Private Sub cboClasificacionSA_Click()
    On Error GoTo NotificaError

    txtIniciales.Text = ""
    pconfiguragrid

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboClasificacionSA_Click"))
End Sub

Private Sub cboClasificacionSA_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then SendKeys vbTab

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboClasificacionSA_KeyDown"))
End Sub

Private Sub cboConceptoFacturacion_Click()
    On Error GoTo NotificaError

    txtIniciales.Text = ""
    pconfiguragrid

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboConceptoFacturacion_Click"))
End Sub

Private Sub cboConceptoFacturacion_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then SendKeys vbTab

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboConceptoFacturacion_KeyDown"))
End Sub

Private Sub cboDepartamento_Click()
    On Error GoTo NotificaError

    If cboDepartamento.ListIndex <> -1 Then
        Set rsListasDepto = frsEjecuta_SP(CStr(cboDepartamento.ItemData(cboDepartamento.ListIndex)), "SP_PVSELLISTADEPTO")
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboDepartamento_Click"))
End Sub

Private Sub cboDepartamento_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError

   If KeyAscii = vbKeyReturn Then
      pEnfocaTextBox txtClave
   End If
   
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboDepartamento_KeyPress"))
End Sub

Private Sub cboDepartamento_LostFocus()
    On Error GoTo NotificaError

    pCargaOtrasListas -1
    pEnfocaTextBox txtClave
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboDepartamento_LostFocus"))
End Sub

Private Sub cboFamilia_Click()
    On Error GoTo NotificaError

    Dim rs As ADODB.Recordset
    
    txtIniciales.Text = ""
    pconfiguragrid
    
    cboSubFamilia.Clear
    If cboFamilia.ListIndex > 0 Then
        'TIPOS DE ARTÍCULO
        '0 = Articulos
        '1 = Medicamentos
        Set rs = frsEjecuta_SP(cboFamilia.ItemData(cboFamilia.ListIndex) & "|" & IIf(cboArtMed.ItemData(cboArtMed.ListIndex) = 1, 0, 1), "SP_IVSELSUBFAMILIAXFAMILIA")
        pLlenarCboRs_new cboSubFamilia, rs, 2, 3
        rs.Close
    End If
    cboSubFamilia.AddItem "<TODOS>", 0
    cboSubFamilia.ItemData(cboSubFamilia.NewIndex) = 0
    cboSubFamilia.ListIndex = 0
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboFamilia_Click"))
End Sub

Private Sub cboFamilia_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then SendKeys vbTab

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboFamilia_KeyDown"))
End Sub

Private Sub cboSubFamilia_Click()
    On Error GoTo NotificaError

    txtIniciales.Text = ""
    pconfiguragrid

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboSubfamilia_Click"))
End Sub

Private Sub cboSubfamilia_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then SendKeys vbTab

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboSubFamilia_KeyDown"))
End Sub

Private Sub cboTabulador_GotFocus()
    pHabilita 0, 0, 0, 0, 0, 1, 0, 0
End Sub

Private Sub cboTabulador_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        chkActivo.SetFocus
    End If
End Sub

Private Sub cboTipoCargo_Click()
    On Error GoTo NotificaError
       
    txtIniciales.Text = ""
    pconfiguragrid
    
    If cboTipoCargo.Text = "ARTÍCULOS" Then
        pCargaCbosArticulo
    End If
    If cboTipoCargo.Text = "ESTUDIOS" Then
        pCargaCboClasificacionSA "IM"
    End If
    If cboTipoCargo.Text = "EXÁMENES" Or cboTipoCargo.Text = "EXÁMENES Y GRUPOS" Or cboTipoCargo.Text = "GRUPOS DE EXÁMENES" Then
        pCargaCboClasificacionSA "LA"
    End If
           
    lblTipoArticulo.Enabled = cboTipoCargo.Text = "ARTÍCULOS"
    cboArtMed.Enabled = cboTipoCargo.Text = "ARTÍCULOS"
    txtClaveArticulo.Enabled = cboTipoCargo.Text = "ARTÍCULOS"
    If txtClaveArticulo.Enabled = False Then txtClaveArticulo.Text = ""
    lblClaveArticulo.Enabled = cboTipoCargo.Text = "ARTÍCULOS"
    
    lblFamilia.Enabled = cboTipoCargo.Text = "ARTÍCULOS"
    cboFamilia.Enabled = cboTipoCargo.Text = "ARTÍCULOS"
    
    lblSubFamilia.Enabled = cboTipoCargo.Text = "ARTÍCULOS"
    cboSubFamilia.Enabled = cboTipoCargo.Text = "ARTÍCULOS"
    
    lblClasificacion.Enabled = cboTipoCargo.Text = "EXÁMENES" Or cboTipoCargo.Text = "EXÁMENES Y GRUPOS" Or cboTipoCargo.Text = "GRUPOS DE EXÁMENES" Or cboTipoCargo.Text = "ESTUDIOS"
    cboClasificacionSA.Enabled = cboTipoCargo.Text = "EXÁMENES" Or cboTipoCargo.Text = "EXÁMENES Y GRUPOS" Or cboTipoCargo.Text = "GRUPOS DE EXÁMENES" Or cboTipoCargo.Text = "ESTUDIOS"
    
    If cboTipoCargo.Text <> "ARTÍCULOS" Then
        cboArtMed.ListIndex = 0
        cboFamilia.Text = ""
        cboSubFamilia.Text = ""
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboTipoCargo_Click"))
End Sub

Private Sub pCargaCbosArticulo()
    On Error GoTo NotificaError

    cboArtMed.ListIndex = -1
    cboArtMed.ListIndex = 0
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargaCbosArticulo"))
End Sub

Private Sub pCargaCboClasificacionSA(strTipo As String)
    On Error GoTo NotificaError

    Dim rs As ADODB.Recordset
    cboClasificacionSA.Clear
    If strTipo = "IM" Then
        Set rs = frsEjecuta_SP("1|-1|-1", "SP_IMSELCLASIFICACIONESTUDIO")
        pLlenarCboRs_new cboClasificacionSA, rs, 0, 1
        rs.Close
        cboClasificacionSA.AddItem "<TODOS>", 0
        cboClasificacionSA.ItemData(cboClasificacionSA.NewIndex) = 0
        cboClasificacionSA.ListIndex = 0
    End If
    If strTipo = "LA" Then
        Set rs = frsEjecuta_SP("", "SP_LASELCLASIFICACIONEXAMEN")
        pLlenarCboRs_new cboClasificacionSA, rs, 0, 1
        rs.Close
        cboClasificacionSA.AddItem "<TODOS>", 0
        cboClasificacionSA.ItemData(cboClasificacionSA.NewIndex) = 0
        cboClasificacionSA.ListIndex = 0
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargaCboClasificacionSA"))
End Sub

Private Sub cboTipoCargo_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then
        If cboTipoCargo.Text = "<TODOS>" Or cboTipoCargo.Text = "OTROS CONCEPTOS" Or cboTipoCargo.Text = "PAQUETES" Then
            cboConceptoFacturacion.SetFocus
        End If
        If cboTipoCargo.Text = "ARTÍCULOS" Then
            SendKeys vbTab
        End If
        If cboTipoCargo.Text = "EXÁMENES" Or cboTipoCargo.Text = "EXÁMENES Y GRUPOS" Or cboTipoCargo.Text = "GRUPOS DE EXÁMENES" Or cboTipoCargo.Text = "ESTUDIOS" Then
            cboClasificacionSA.SetFocus
        End If
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboTipoCargo_KeyDown"))
End Sub


Private Sub chkActivo_GotFocus()
    On Error GoTo NotificaError

    pHabilita 0, 0, 0, 0, 0, 1, 0, 0
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":chkActivo_GotFocus"))
End Sub

Private Sub chkActivo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cmdCapturaPrecios.Enabled Then
            cmdCapturaPrecios.SetFocus
        End If
    End If
End Sub

Private Sub Chkbitcheckup_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        If cmdCapturaPrecios.Enabled Then
            cmdCapturaPrecios.SetFocus
        End If
    End If

End Sub

Private Sub chkImprimeCargosPrecioCero_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cmdCapturaPrecios.Enabled Then
            cmdCapturaPrecios.SetFocus
        End If
    End If

End Sub

Private Sub chkIncremento_Click()
    On Error GoTo NotificaError

    Dim lngContador As Long
    With grdPrecios
        For lngContador = 1 To .Rows - 1
            If .TextMatrix(lngContador, cintColIncremetoAutomatico) = "*" Then
                If chkIncremento.Value = 0 Then
                   .TextMatrix(lngContador, cintColIncremetoAutomatico) = ""
                   .TextMatrix(lngContador, cintColModificado) = "*"
                End If
            Else
                If chkIncremento.Value = 1 Then
                   .TextMatrix(lngContador, cintColIncremetoAutomatico) = "*"
                   .TextMatrix(lngContador, cintColModificado) = "*"
                End If
            End If
        Next
    End With
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":chkIncremento_Click"))
End Sub

Private Sub chkPredeterminada_GotFocus()
    On Error GoTo NotificaError

    pHabilita 0, 0, 0, 0, 0, 1, 0, 0
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":chkPredeterminada_GotFocus"))
End Sub

Private Sub chkPredeterminada_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cmdCapturaPrecios.Enabled Then
            cmdCapturaPrecios.SetFocus
        End If
    End If
End Sub

Private Sub chkTabulador_Click()
    On Error GoTo NotificaError

    Dim lngContador As Long
    With grdPrecios
        For lngContador = 1 To .Rows - 1
            If .TextMatrix(lngContador, cintColTipo) = "AR" Then
               If .TextMatrix(lngContador, cintColTabulador) = "*" Then
                  If chkTabulador.Value = 0 Then
                     .TextMatrix(lngContador, cintColTabulador) = ""
                     .TextMatrix(lngContador, cintColModificado) = "*"
                  End If
               Else
                  If chkTabulador.Value = 1 Then
                     .TextMatrix(lngContador, cintColTabulador) = "*"
                     .TextMatrix(lngContador, cintColModificado) = "*"
                  End If
               End If
         
            End If
        Next
    End With
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":chkTabulador_Click"))
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo NotificaError
    Dim lngPersonaGraba As Long
    Dim lngError As Long
    
    If fblnRevisaPermiso(vglngNumeroLogin, cintNumOpcionListas, "E") Then
        ' Persona que graba
        lngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
        If lngPersonaGraba = 0 Then Exit Sub
        
        lngError = 1
        frsEjecuta_SP txtClave.Text, "sp_PvDelListaPrecio", False, lngError
        
        If lngError = 0 Then
            pGuardarLogTransaccion Me.Name, EnmBorrar, lngPersonaGraba, "LISTA DE PRECIOS", txtClave.Text
        Else
            'No se puede eliminar la información, ya ha sido utilizada.
            MsgBox SIHOMsg(771), vbOKOnly + vbCritical, "Mensaje"
        End If
        
        Set rsListasDepto = frsEjecuta_SP(CStr(cboDepartamento.ItemData(cboDepartamento.ListIndex)), "SP_PVSELLISTADEPTO")
        
        If chkPredeterminada.Value = 1 Then
            'Se ha eliminado la lista predeterminada, deberá seleccionar otra.
            MsgBox SIHOMsg(793), vbOKOnly + vbInformation, "Mensaje"
        End If
        
        txtClave.SetFocus
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdDelete_Click"))
End Sub

Private Sub pFiltraNombre()
    On Error GoTo NotificaError

    Dim lngContador As Long
    Dim blnHabilitaExportar As Boolean
    
    grdPrecios.Visible = False
    blnHabilitaExportar = False
    lngContador = 1
    Do While lngContador <= grdPrecios.Rows - 1
        If Trim(grdPrecios.TextMatrix(lngContador, cintColDescripcion)) >= Trim(txtIniciales.Text) Then blnHabilitaExportar = True
        grdPrecios.RowHeight(lngContador) = IIf(Trim(grdPrecios.TextMatrix(lngContador, cintColDescripcion)) >= Trim(txtIniciales.Text), clngLargoRenglon, 0)
        lngContador = lngContador + 1
    Loop
    cmdExportar.Enabled = blnHabilitaExportar
    'Se agrega para el botón importar
    cmdImportar.Enabled = blnHabilitaExportar
    grdPrecios.Visible = True

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pFiltraNombre"))
End Sub

Private Sub cmdBorrarRegistro_Click()

End Sub

Private Sub cmdDelete2_Click()

End Sub

Private Sub cmdExportar_Click()
On Error GoTo NotificaError
'    Dim rsAux As New ADODB.Recordset
    Dim o_Excel As Object
'    Dim o_ExcelAbrir As Object
    Dim o_Libro As Object
    Dim o_Sheet As Object
    Dim intRow As Long
    Dim intCol As Integer
    Dim dblAvance As Double
    Dim intRowExcel As Long
        
    If grdPrecios.Rows > 1 And grdPrecios.TextMatrix(1, 1) <> "" Then
        Set o_Excel = CreateObject("Excel.Application")
        Set o_Libro = o_Excel.Workbooks.Add
        Set o_Sheet = o_Libro.Worksheets(1)
        
        If Not IsObject(o_Excel) Then
            MsgBox "Necesitas Microsoft Excel para utilizar esta funcionalidad", _
               vbExclamation, "Mensaje"
            Exit Sub
        End If
        
        'Columnas titulos
        o_Excel.Cells(1, 1).Value = "Clave"
        o_Excel.Cells(1, 2).Value = "Descripción cargo"
        o_Excel.Cells(1, 3).Value = "Incremento automático"
        o_Excel.Cells(1, 4).Value = "Tipo incremento"
        o_Excel.Cells(1, 5).Value = "Margen utilidad"
        o_Excel.Cells(1, 6).Value = "Usar tabulador"
        o_Excel.Cells(1, 7).Value = "Costo base"
        o_Excel.Cells(1, 8).Value = "Precio"
        o_Excel.Cells(1, 9).Value = "Moneda"
        o_Excel.Cells(1, 10).Value = "Tipo"
        o_Excel.Cells(1, 11).Value = "Concepto de facturación"
        
        o_Sheet.Range("A1:K1").HorizontalAlignment = -4108
        o_Sheet.Range("A1:K1").VerticalAlignment = -4108
        o_Sheet.Range("A1:K1").WrapText = True
        o_Sheet.Range("A2").Select
        o_Excel.ActiveWindow.FreezePanes = True
        o_Sheet.Range("A1:K1").Interior.ColorIndex = 15 '15 48
        
        'o_Sheet.Range(o_Excel.Cells(grdPrecios.Rows + 7, 13), o_Excel.Cells(grdPrecios.Rows + 7, 23)).Interior.ColorIndex = 15
        o_Sheet.Range("A:A").ColumnWidth = 12
        o_Sheet.Range("B:B").ColumnWidth = 50
        o_Sheet.Range("C:C").ColumnWidth = 12
        o_Sheet.Range("D:D").ColumnWidth = 15
        o_Sheet.Range("E:E").ColumnWidth = 12
        o_Sheet.Range("G:G").ColumnWidth = 12
        o_Sheet.Range("H:H").ColumnWidth = 15
        o_Sheet.Range("I:I").ColumnWidth = 20
        o_Sheet.Range("J:J").ColumnWidth = 30
        o_Sheet.Range("K:K").ColumnWidth = 50
        
        'o_Sheet.Range(o_Excel.Cells(1, 1), o_Excel.Cells(grdPrecios.Rows, 15)).Borders(4).LineStyle = 1
        
        'info del rs
        o_Sheet.Range("A:K").Font.Size = 9
        o_Sheet.Range("A:K").Font.Name = "Times New Roman" '
        o_Sheet.Range("A:K").Font.Bold = False
        
        'o_Sheet.Range("A:A").NumberFormat = "0000000000"
        'titulos
        o_Sheet.Range("A1:X1").Font.Bold = True
        o_Sheet.Range(o_Excel.Cells(grdPrecios.Rows + 7, 1), o_Excel.Cells(grdPrecios.Rows + 7, 23)).Font.Bold = True
        'o_Sheet.Range(o_Excel.Cells(2, 1), o_Excel.Cells(5, 1)).Font.Bold = True
        'centrado, auto ajustar texto, alinear medio
        o_Sheet.Range("G:H").NumberFormat = "$ ###,###,###,##0.00"
        
        
        
        dblAvance = 100 / grdPrecios.Rows
        '------------------------
        ' Configuración de la Barra de estado
        '------------------------
        lblTextoBarra.Caption = "Exportando información, por favor espere..."
        freBarra.Visible = True
        freBarra.Top = 720
        pgbBarra.Value = 0
        freBarra.Refresh
        pgbBarra.Value = 0
        
        intRowExcel = 2
        'Recorre el grid y llena el Excel
        For intRow = 2 To grdPrecios.Rows '- 1

            ' Actualización de la barra de estado
            If pgbBarra.Value + dblAvance < 100 Then
                pgbBarra.Value = pgbBarra.Value + dblAvance
            Else
                pgbBarra.Value = 100
            End If
            pgbBarra.Refresh

            If grdPrecios.RowHeight(intRow - 1) > 0 Then
                With grdPrecios
                    
                    Dim strTipo As String
                    strTipo = .TextMatrix(intRow - 1, 10)
                    
                    If strTipo = "AR" Then
                        o_Sheet.Cells(intRowExcel, 1).NumberFormat = "0000000000"
                    Else
                        o_Sheet.Cells(intRowExcel, 1).NumberFormat = "0"
                    End If
                    
                    o_Sheet.Cells(intRowExcel, 1).Value = .TextMatrix(intRow - 1, 1) & " "
                    o_Sheet.Cells(intRowExcel, 2).Value = .TextMatrix(intRow - 1, 2) & " "
                    '
                    o_Sheet.Cells(intRowExcel, 3).Value = .TextMatrix(intRow - 1, 3) & ""
                    o_Sheet.Cells(intRowExcel, 4).Value = .TextMatrix(intRow - 1, 4) & " "
                    o_Sheet.Cells(intRowExcel, 5).Value = .TextMatrix(intRow - 1, 5) & " "
                    '
                    o_Sheet.Cells(intRowExcel, 6).Value = .TextMatrix(intRow - 1, 6) & ""
                    o_Sheet.Cells(intRowExcel, 7).Value = .TextMatrix(intRow - 1, 7) & " "
                    o_Sheet.Cells(intRowExcel, 8).Value = .TextMatrix(intRow - 1, 8) & " "
                    o_Sheet.Cells(intRowExcel, 9).Value = .TextMatrix(intRow - 1, 9) & " "
                    o_Sheet.Cells(intRowExcel, 10).Value = .TextMatrix(intRow - 1, 11) & " "
                    o_Sheet.Cells(intRowExcel, 11).Value = .TextMatrix(intRow - 1, 15) & " "
                End With
                intRowExcel = intRowExcel + 1
            End If
        Next
        
        'La información ha sido exportada exitosamente
        'MsgBox SIHOMsg(1185), vbOKOnly + vbInformation, "Mensaje"
        freBarra.Visible = False
        o_Excel.Visible = True
        
        Set o_Excel = Nothing
        cmdImportar.SetFocus
        
    Else
        'No existe información con esos parámetros
        MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje"
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

Private Sub cmdFiltrar_Click()
    On Error GoTo NotificaError

    Dim strTipoCargo As String
    
    If cboTipoCargo.ListIndex <> -1 Then
        If Trim(txtIniciales.Text) <> "" Then
            pFiltraNombre
        Else
            If cboTipoCargo.ListIndex = cintIndexTodos Then
                strTipoCargo = "*"
            ElseIf cboTipoCargo.ListIndex = cintIndexArticulo Then
                strTipoCargo = "AR"
            ElseIf cboTipoCargo.ListIndex = cintIndexExamen Then
                strTipoCargo = "EX"
            ElseIf cboTipoCargo.ListIndex = cintIndexExamenGrupo Then
                strTipoCargo = "EG"
            ElseIf cboTipoCargo.ListIndex = cintIndexGrupo Then
                strTipoCargo = "GE"
            ElseIf cboTipoCargo.ListIndex = cintIndexEstudio Then
                strTipoCargo = "ES"
            ElseIf cboTipoCargo.ListIndex = cintIndexOtro Then
                strTipoCargo = "OC"
            ElseIf cboTipoCargo.ListIndex = cintIndexPaquete Then
                strTipoCargo = "PA"
            End If
            pconfiguragrid
            pCargaLista Val(txtClave.Text), strTipoCargo, "", cboConceptoFacturacion.ItemData(cboConceptoFacturacion.ListIndex)
        End If
    End If

    pHabilitaModificar
    vlblnRecalcularImportar = False 'se agrego para inicializar variable para recalcular importar
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdFiltrar_Click"))
End Sub

Private Sub cmdImportar_Click()
On Error GoTo NotificaError
    Dim objXLApp As Object
    Dim txtRuta As String
    Dim vlstrValidador As String
    Dim intLoopCounter As Integer
    Dim intGridCounter As Integer
    Dim dblAvance As Double
    Dim vlblnModificaVal As Boolean
    Dim vlstrTipoCargo As String
    Dim vlRespuesta As Integer
    Dim intValorCero As Integer
    
   
    Set objXLApp = CreateObject("Excel.Application")

    CommonDialog1.DialogTitle = "Abrir archivo"
    CommonDialog1.Filter = "Documentos excel|*.xls;*.xlsx;"
    CommonDialog1.FilterIndex = 1
    CommonDialog1.Flags = cdlOFNFileMustExist
    CommonDialog1.CancelError = True
    On Error Resume Next
    CommonDialog1.ShowOpen
    If Err Then
        'Si se cancela el cuadro de diálogo
        Exit Sub
    End If
    txtRuta = CommonDialog1.FileName
    vlstrValidador = ""
    With objXLApp
        .Workbooks.Open txtRuta
        .Workbooks(1).Worksheets(1).Select
        
        dblAvance = 50 / CInt(.Cells.Find("*", SearchOrder:=1, SearchDirection:=2).Row)
        '------------------------
        ' Configuración de la Barra de estado
        '------------------------
        lblTextoBarra.Caption = "Importando información, por favor espere..."
        freBarra.Visible = True
        freBarra.Top = 720
        pgbBarra.Value = 0
        freBarra.Refresh
        pgbBarra.Value = 0
        
        '---VALIDACIONES DE EXCEL ANTES DE REALIZAR LA IMPORTACIÓN---
        '--- VALIDACIÓN PARA SABER SI NO HAY DATOS EN EL ARCHIVO ---
        If CInt(.Cells.Find("*", SearchOrder:=1, SearchDirection:=2).Row) > 1 Then
            '---VALIDACIÓN PARA SABER SI TIENEN LA MISMA CANTIDAD DE FILAS----
            If grdPrecios.Rows - 1 <> (CInt(.Cells.Find("*", SearchOrder:=1, SearchDirection:=2).Row) - 1) Then
                MsgBox "La cantidad de cargos que se intentan importar son diferentes a los cargos que ya se tienen en la cuadrícula", vbOKOnly + vbInformation, "Mensaje"
                freBarra.Visible = False
                pgbBarra.Value = 100
                .Workbooks(1).Close False
                .Quit
                objXLApp = Nothing
                Exit Sub
            End If
                        
            '---VALIDACIÓN PARA SABER SI TIENE FILAS REPETIDAS O QUE NO ESTEN EN LA CUADRICULA---
            For intLoopCounter = 2 To CInt(.Cells.Find("*", SearchOrder:=1, SearchDirection:=2).Row)
                vlblnModificaVal = False
                'For intGridCounter = 1 To grdPrecios.Rows - 1
                    If Trim(.Range("J" & intLoopCounter).Text) = "ARTICULO" Then
                        vlstrTipoCargo = "AR"
                    ElseIf Trim(.Range("J" & intLoopCounter).Text) = "EXAMEN" Then
                        vlstrTipoCargo = "EX"
                    ElseIf Trim(.Range("J" & intLoopCounter).Text) = "GRUPO EXAMEN" Then
                        vlstrTipoCargo = "GE"
                    ElseIf Trim(.Range("J" & intLoopCounter).Text) = "ESTUDIO" Then
                        vlstrTipoCargo = "ES"
                    ElseIf Trim(.Range("J" & intLoopCounter).Text) = "OTRO CONCEPTO" Then
                        vlstrTipoCargo = "OC"
                    ElseIf Trim(.Range("J" & intLoopCounter).Text) = "PAQUETE" Then
                        vlstrTipoCargo = "PA"
                    End If
                        
                    If (Trim(grdPrecios.TextMatrix(intLoopCounter - 1, 1)) = Trim(.Range("A" & intLoopCounter).Text)) And (Trim(grdPrecios.TextMatrix(intLoopCounter - 1, 2)) = Trim(.Range("B" & intLoopCounter).Text)) And (Trim(grdPrecios.TextMatrix(intLoopCounter - 1, 10)) = vlstrTipoCargo) And (Trim(grdPrecios.TextMatrix(intLoopCounter - 1, 15)) = Trim(.Range("K" & intLoopCounter).Text)) Then 'Clave, descripción, tipo y concepto de facturación
                        vlblnModificaVal = True
                        If grdPrecios.TextMatrix(intLoopCounter - 1, 0) = "*" Then
                            MsgBox "Renglón " & intLoopCounter & ". Clave, Descripción de cargo, Tipo y Concepto de facturación ya se encuentran en otra fila. ", vbOKOnly + vbInformation, "Mensaje"
                            pLimpiaSeleccionGrid
                            freBarra.Visible = False
                            pgbBarra.Value = 100
                            .Workbooks(1).Close False
                            .Quit
                            objXLApp = Nothing
                            Exit Sub
                        Else
                            '--- COSTO BASE ---
                            If Not (IsNumeric(Trim(.Range("G" & intLoopCounter)))) Then 'Costo base
                                MsgBox "Renglón " & intLoopCounter & ". Formato de la columna Costo base es incorrecto. ", vbOKOnly + vbInformation, "Mensaje"
                                pLimpiaSeleccionGrid
                                freBarra.Visible = False
                                pgbBarra.Value = 100
                                .Workbooks(1).Close False
                                .Quit
                                objXLApp = Nothing
                                Exit Sub
                            Else
                                If (Trim(.Range("J" & intLoopCounter).Text) = "ARTICULO") And (CLng(Trim(.Range("G" & intLoopCounter))) <> CLng(Trim(grdPrecios.TextMatrix(intLoopCounter - 1, 7)))) Then
                                    MsgBox "Renglón " & intLoopCounter & ". Valor de la columna Costo base es incorrecto. ", vbOKOnly + vbInformation, "Mensaje"
                                    pLimpiaSeleccionGrid
                                    freBarra.Visible = False
                                    pgbBarra.Value = 100
                                    .Workbooks(1).Close False
                                    .Quit
                                    objXLApp = Nothing
                                    Exit Sub
                                End If
                                
                                If CDbl(Trim(.Range("G" & intLoopCounter))) < 0 Then
                                    MsgBox "Renglón " & intLoopCounter & ". Valor de la columna Costo base es incorrecto. ", vbOKOnly + vbInformation, "Mensaje"
                                    pLimpiaSeleccionGrid
                                    freBarra.Visible = False
                                    pgbBarra.Value = 100
                                    .Workbooks(1).Close False
                                    .Quit
                                    objXLApp = Nothing
                                    Exit Sub
                                Else
                                    grdPrecios.TextMatrix(intLoopCounter - 1, 0) = "*"
                                    'Exit For
                                End If
                                
                            End If
                        End If
                    End If
                'Next intGridCounter
                If vlblnModificaVal = False Then
                    'MsgBox "Renglón " & intLoopCounter & ". Clave, Descripción de cargo, Tipo y Concepto de facturación no encontrados en la cuadrícula. ", vbOKOnly + vbInformation, "Mensaje"
                    MsgBox "Renglón " & intLoopCounter & ". Clave, Descripción de cargo, Tipo y Concepto de facturación no corresponde con la cuadrícula. ", vbOKOnly + vbInformation, "Mensaje"
                    pLimpiaSeleccionGrid
                    freBarra.Visible = False
                    pgbBarra.Value = 100
                    .Workbooks(1).Close False
                    .Quit
                    objXLApp = Nothing
                    Exit Sub
                End If
                
                '--- SI ENCUENTRA LA FILA EN LA CUADRÍCULA Y NO ESTA REPETIDA, SIGUE CON LAS DEMÁS VALIDACIONES ---
                '--- CLAVE ---
                If Trim(.Range("A" & intLoopCounter)) = "" Then 'Clave
                    MsgBox "Renglón " & intLoopCounter & ". Clave de cargo no encontrada. ", vbOKOnly + vbInformation, "Mensaje"
                    pLimpiaSeleccionGrid
                    freBarra.Visible = False
                    pgbBarra.Value = 100
                    .Workbooks(1).Close False
                    .Quit
                    objXLApp = Nothing
                    Exit Sub
                End If
                '--- DESCRIPCIÓN CARGO ---
                If Trim(.Range("B" & intLoopCounter)) = "" Then 'Descripción cargo
                    MsgBox "Renglón " & intLoopCounter & ". Descripción de cargo no encontrada. ", vbOKOnly + vbInformation, "Mensaje"
                    pLimpiaSeleccionGrid
                    freBarra.Visible = False
                    pgbBarra.Value = 100
                    .Workbooks(1).Close False
                    .Quit
                    objXLApp = Nothing
                    Exit Sub
                End If
                '---CONCEPTO DE FACTURACIÓN ---
                If Trim(.Range("K" & intLoopCounter)) = "" Then 'Concepto de facturación
                    MsgBox "Renglón " & intLoopCounter & ". Concepto de facturación no encontrado. ", vbOKOnly + vbInformation, "Mensaje"
                    pLimpiaSeleccionGrid
                    freBarra.Visible = False
                    pgbBarra.Value = 100
                    .Workbooks(1).Close False
                    .Quit
                    objXLApp = Nothing
                    Exit Sub
                End If
                '--- INCREMENTO AUTOMÁTICO ---
                If Not (Trim(.Range("C" & intLoopCounter)) = "*" Or Trim(.Range("C" & intLoopCounter)) = "") Then 'Incremento automático
                    MsgBox "Renglón " & intLoopCounter & ". Valor de la columna incremento automático es incorrecto. ", vbOKOnly + vbInformation, "Mensaje"
                    pLimpiaSeleccionGrid
                    freBarra.Visible = False
                    pgbBarra.Value = 100
                    .Workbooks(1).Close False
                    .Quit
                    objXLApp = Nothing
                    Exit Sub
                End If
                '---TIPO Y TIPO INCREMENTO ---
                If (Trim(.Range("J" & intLoopCounter)) = "ARTICULO") Then 'Tipo
                    If Not (Trim(.Range("D" & intLoopCounter)) = "PRECIO MÁXIMO AL PÚBLICO" Or Trim(.Range("D" & intLoopCounter)) = "COMPRA MÁS ALTA" Or Trim(.Range("D" & intLoopCounter)) = "ÚLTIMA COMPRA") Then  'Tipo incremento
                        MsgBox "Renglón " & intLoopCounter & ". Valor de la columna Tipo Incremento es incorrecto. ", vbOKOnly + vbInformation, "Mensaje"
                        pLimpiaSeleccionGrid
                        freBarra.Visible = False
                        pgbBarra.Value = 100
                        .Workbooks(1).Close False
                        .Quit
                        objXLApp = Nothing
                        Exit Sub
                    End If
                Else
                    If (Trim(.Range("J" & intLoopCounter)) = "EXAMEN" Or Trim(.Range("J" & intLoopCounter)) = "ESTUDIO" Or Trim(.Range("J" & intLoopCounter)) = "GRUPO EXAMEN" Or Trim(.Range("J" & intLoopCounter)) = "OTRO CONCEPTO" Or Trim(.Range("J" & intLoopCounter)) = "PAQUETE") Then 'Tipo
                        If Not (Trim(.Range("D" & intLoopCounter)) = "NA") Then  'Tipo incremento
                            MsgBox "Renglón " & intLoopCounter & ". Valor de la columna Tipo Incremento es incorrecto. ", vbOKOnly + vbInformation, "Mensaje"
                            pLimpiaSeleccionGrid
                            freBarra.Visible = False
                            pgbBarra.Value = 100
                            .Workbooks(1).Close False
                            .Quit
                            objXLApp = Nothing
                            Exit Sub
                        End If
                    Else
                        MsgBox "Renglón " & intLoopCounter & ". Valor de la columna Tipo es incorrecto. ", vbOKOnly + vbInformation, "Mensaje"
                        pLimpiaSeleccionGrid
                        freBarra.Visible = False
                        pgbBarra.Value = 100
                        .Workbooks(1).Close False
                        .Quit
                        objXLApp = Nothing
                        Exit Sub
                    End If
                End If
                '--- MARGEN DE UTILIDAD ---
                If Not (IsNumeric(Trim(.Range("E" & intLoopCounter)))) Then 'Margen de utilidad
                    MsgBox "Renglón " & intLoopCounter & ". Valor de la columna Margen de utilidad es incorrecto. ", vbOKOnly + vbInformation, "Mensaje"
                    pLimpiaSeleccionGrid
                    freBarra.Visible = False
                    pgbBarra.Value = 100
                    .Workbooks(1).Close False
                    .Quit
                    objXLApp = Nothing
                    Exit Sub
                Else
                    If CDbl(Trim(.Range("E" & intLoopCounter))) < 0 Then
                        MsgBox "Renglón " & intLoopCounter & ". Valor de la columna Margen de utilidad es incorrecto. ", vbOKOnly + vbInformation, "Mensaje"
                        pLimpiaSeleccionGrid
                        freBarra.Visible = False
                        pgbBarra.Value = 100
                        .Workbooks(1).Close False
                        .Quit
                        objXLApp = Nothing
                        Exit Sub
                    End If
                End If
                '--- USAR TABULADOR ---
                If Not (Trim(.Range("F" & intLoopCounter)) = "*" Or Trim(.Range("F" & intLoopCounter)) = "") Then 'Usar tabulador
                    MsgBox "Renglón " & intLoopCounter & ". Valor de la columna Usar tabulador es incorrecto. ", vbOKOnly + vbInformation, "Mensaje"
                    pLimpiaSeleccionGrid
                    freBarra.Visible = False
                    pgbBarra.Value = 100
                    .Workbooks(1).Close False
                    .Quit
                    objXLApp = Nothing
                    Exit Sub
                End If
                '--- PRECIO ---
                If Not (IsNumeric(Trim(.Range("H" & intLoopCounter)))) Then 'Precio
                   MsgBox "Renglón " & intLoopCounter & ". Formato de la columna Precio es incorrecto. ", vbOKOnly + vbInformation, "Mensaje"
                   pLimpiaSeleccionGrid
                   freBarra.Visible = False
                   pgbBarra.Value = 100
                   .Workbooks(1).Close False
                   .Quit
                   objXLApp = Nothing
                   Exit Sub
                Else
                   ' --- PRECIO IGUAL A 0 ---
                   If CDbl(Trim(.Range("H" & intLoopCounter))) = 0 Then
                        intValorCero = intValorCero + 1
                   Else
                        ' --- PRECIO MENOR A 0 ---
                        If CDbl(Trim(.Range("H" & intLoopCounter))) < 0 Then
                            MsgBox "Renglón " & intLoopCounter & ". Valor de la columna Precio es incorrecto. ", vbOKOnly + vbInformation, "Mensaje"
                            pLimpiaSeleccionGrid
                            freBarra.Visible = False
                            pgbBarra.Value = 100
                            .Workbooks(1).Close False
                            .Quit
                            objXLApp = Nothing
                            Exit Sub
                        End If
                   End If
                End If
                '--- MONEDA ---
                If Trim(.Range("J" & intLoopCounter)) = "PAQUETE" Then
                    If Not (Trim(.Range("I" & intLoopCounter)) = "PESOS" Or Trim(.Range("I" & intLoopCounter)) = "DÓLARES") Then 'Moneda
                        MsgBox "Renglón " & intLoopCounter & ". Valor de la columna Moneda es incorrecto. ", vbOKOnly + vbInformation, "Mensaje"
                        pLimpiaSeleccionGrid
                        freBarra.Visible = False
                        pgbBarra.Value = 100
                        .Workbooks(1).Close False
                        .Quit
                        objXLApp = Nothing
                        Exit Sub
                    End If
                Else
                    If Not (Trim(.Range("I" & intLoopCounter)) = "PESOS") Then 'Moneda
                        MsgBox "Renglón " & intLoopCounter & ". Valor de la columna Moneda es incorrecto. ", vbOKOnly + vbInformation, "Mensaje"
                        pLimpiaSeleccionGrid
                        freBarra.Visible = False
                        pgbBarra.Value = 100
                        .Workbooks(1).Close False
                        .Quit
                        objXLApp = Nothing
                        Exit Sub
                    End If
                End If
                
                ' Actualización de la barra de estado
                If pgbBarra.Value + dblAvance < 50 Then
                    pgbBarra.Value = pgbBarra.Value + dblAvance
                Else
                    pgbBarra.Value = 50
                End If
            Next intLoopCounter
            
            ' --- Existen filas con valor de precio cero. ¿Desea continuar con la importación? ---
            If intValorCero > 0 Then
                vlRespuesta = MsgBox("Existen filas con valor de precio cero. ¿Desea continuar con la importación? ", vbYesNo + vbQuestion, "Mensaje")
      
                If vlRespuesta = 7 Then
                    pLimpiaSeleccionGrid
                    freBarra.Visible = False
                    pgbBarra.Value = 100
                    .Workbooks(1).Close False
                    .Quit
                    objXLApp = Nothing
                    Exit Sub
                End If
            End If
            
            '---TERMINAN VALIDACIONES, INICIA IMPORTACIÓN DE DATOS ---
            For intLoopCounter = 2 To CInt(.Cells.Find("*", SearchOrder:=1, SearchDirection:=2).Row)
                'For intGridCounter = 1 To grdPrecios.Rows - 1
                    vlstrTipoCargo = ""
                    If Trim(.Range("J" & intLoopCounter).Text) = "ARTICULO" Then
                        vlstrTipoCargo = "AR"
                    ElseIf Trim(.Range("J" & intLoopCounter).Text) = "EXAMEN" Then
                        vlstrTipoCargo = "EX"
                    ElseIf Trim(.Range("J" & intLoopCounter).Text) = "GRUPO EXAMEN" Then
                        vlstrTipoCargo = "GE"
                    ElseIf Trim(.Range("J" & intLoopCounter).Text) = "ESTUDIO" Then
                        vlstrTipoCargo = "ES"
                    ElseIf Trim(.Range("J" & intLoopCounter).Text) = "OTRO CONCEPTO" Then
                        vlstrTipoCargo = "OC"
                    ElseIf Trim(.Range("J" & intLoopCounter).Text) = "PAQUETE" Then
                        vlstrTipoCargo = "PA"
                    End If
                        
                    If (Trim(grdPrecios.TextMatrix(intLoopCounter - 1, 1)) = Trim(.Range("A" & intLoopCounter).Text)) And (Trim(grdPrecios.TextMatrix(intLoopCounter - 1, 2)) = Trim(.Range("B" & intLoopCounter).Text)) And (Trim(grdPrecios.TextMatrix(intLoopCounter - 1, 10)) = vlstrTipoCargo) And (Trim(grdPrecios.TextMatrix(intLoopCounter - 1, 15)) = Trim(.Range("K" & intLoopCounter).Text)) Then  'Clave, descripción, tipo y concepto de facturación
                        If (Trim(.Range("J" & intLoopCounter)) = "ARTICULO") Then 'Tipo
                            If Trim(.Range("C" & intLoopCounter)) = "*" Or Trim(.Range("F" & intLoopCounter)) = "*" Then
                                Trim((grdPrecios.TextMatrix(intLoopCounter - 1, 0))) = "*"
                                vlblnRecalcularImportar = True
                            Else
                                grdPrecios.TextMatrix(intLoopCounter - 1, 0) = ""
                            End If
                            grdPrecios.TextMatrix(intLoopCounter - 1, 3) = Trim(.Range("C" & intLoopCounter).Text) 'Incremento autómatico
                            grdPrecios.TextMatrix(intLoopCounter - 1, 4) = Trim(.Range("D" & intLoopCounter).Text) 'Tipo incremento
                            grdPrecios.TextMatrix(intLoopCounter - 1, 5) = Format(Trim(.Range("E" & intLoopCounter)), "Percent") 'Margen de utilidad
                            'grdPrecios.TextMatrix(intLoopCounter - 1, 5) = Format(CDbl(.range("E" & intLoopCounter)) * 100, "0.0000") & "%"
                            grdPrecios.TextMatrix(intLoopCounter - 1, 6) = Trim(.Range("F" & intLoopCounter).Text) 'Usar tabulador
                            grdPrecios.TextMatrix(intLoopCounter - 1, 8) = Format(Trim(.Range("H" & intLoopCounter).Text), "$###,###,###,##0.00####") 'Precio
                            Select Case Trim(.Range("D" & intLoopCounter))
                                Case cstrUltimaCompra
                                    grdPrecios.TextMatrix(intLoopCounter - 1, cintColCosto) = Format(grdPrecios.TextMatrix(intLoopCounter - 1, cintColCostoUltimaEntrada), "$###,###,###,##0.0000##")
                                Case cstrCompraMasAlta
                                    grdPrecios.TextMatrix(intLoopCounter - 1, cintColCosto) = Format(grdPrecios.TextMatrix(intLoopCounter - 1, cintColCostoMasAlto), "$###,###,###,##0.0000##")
                                Case cstrPrecioMaximoPublico
                                    grdPrecios.TextMatrix(intLoopCounter - 1, cintColCosto) = Format(grdPrecios.TextMatrix(intLoopCounter - 1, cintColPrecioMaximopublico), "$###,###,###,##0.0000##")
                            End Select
                            grdPrecios.TextMatrix(intLoopCounter - 1, cintColModificado) = "*"
                            'Exit For
                        Else
                            If Trim(.Range("C" & intLoopCounter)) = "*" Then
                                grdPrecios.TextMatrix(intLoopCounter - 1, 0) = "*"
                                vlblnRecalcularImportar = True
                            Else
                                grdPrecios.TextMatrix(intLoopCounter - 1, 0) = ""
                            End If
                            grdPrecios.TextMatrix(intLoopCounter - 1, 3) = Trim(.Range("C" & intLoopCounter).Text) 'Incremento autómatico
                            grdPrecios.TextMatrix(intLoopCounter - 1, 5) = Format(Trim(.Range("E" & intLoopCounter)), "Percent") 'Margen de utilidad
                            'grdPrecios.TextMatrix(intLoopCounter - 1, 5) = Format(CDbl(.range("E" & intLoopCounter)) * 100, "0.0000") & "%"
                            grdPrecios.TextMatrix(intLoopCounter - 1, 7) = Format(Trim(.Range("G" & intLoopCounter).Text), "$###,###,###,##0.0000##") 'Costo base
                            grdPrecios.TextMatrix(intLoopCounter - 1, 8) = Format(Trim(.Range("H" & intLoopCounter).Text), "$###,###,###,##0.00####") 'Precio
                            If (Trim(.Range("J" & intLoopCounter)) = "PAQUETE") Then 'Tipo
                                grdPrecios.TextMatrix(intLoopCounter - 1, 9) = Trim(.Range("I" & intLoopCounter).Text) 'Moneda
                            End If
                            grdPrecios.TextMatrix(intLoopCounter - 1, cintColModificado) = "*"
                            'Exit For
                        End If
                        
                    End If
                'Next intGridCounter
                                      
                ' Actualización de la barra de estado
                If pgbBarra.Value + dblAvance < 100 Then
                    pgbBarra.Value = pgbBarra.Value + dblAvance
                Else
                    pgbBarra.Value = 100
                End If
            Next intLoopCounter
                
            If vlblnRecalcularImportar Then
                Call cmdRec_Click
            End If
               
            MsgBox "Se realizó la importación exitosa de los datos en la cuadrícula.", vbOKOnly + vbInformation, "Mensaje"

            freBarra.Visible = False
            pgbBarra.Value = 100
            .Workbooks(1).Close False
            .Quit
        Else
            MsgBox "No existen datos para importar", vbOKOnly + vbInformation, "Mensaje"
            freBarra.Visible = False
            pgbBarra.Value = 100
            .Workbooks(1).Close False
            .Quit
        End If
        
    End With
      
    Set objXLApp = Nothing
    cmdFiltrar.SetFocus
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdImportar_Click"))
    Unload Me
    
End Sub

Private Sub cmdImprimir_Click()
    On Error GoTo NotificaError

    Dim rs As New ADODB.Recordset
    Dim vlstrSentencia As String
    Dim vlstrNombreHospital As String
    Dim vlstrRegistro As String
    Dim vlstrDireccionHospital As String
    Dim vlstrTelefonoHospital As String
    Dim vlstrDepartamento As String
    Dim rspvselElementosListasPrecios As New ADODB.Recordset
    Dim alstrParametros(2) As String
    
    pInstanciaReporte vgrptReporte, "rptListaPrecios.rpt"
    
    vgstrParametrosSP = CStr(cboDepartamento.ItemData(cboDepartamento.ListIndex)) & "|" & txtClave.Text & "|" & "*|0|3|0|0|0||" & IIf(Chkbitcheckup.Value, 1, 0) & "|" & IIf(chkImprimeCargosPrecioCero.Value, 1, 0)
    ' Si está marcado el check "Imprimir cargos con precio en cero" se imprimirán los elementos con precio CERO, si no está marcado se imprimirán todos
    Set rspvselElementosListasPrecios = frsEjecuta_SP(vgstrParametrosSP, "sp_pvselElementosListasPrecios")
    If rspvselElementosListasPrecios.EOF Then
        MsgBox SIHOMsg(13), vbInformation, "Mensaje"
    Else
        vgrptReporte.DiscardSavedData
        alstrParametros(0) = "NombreHospital;" & Trim(vgstrNombreHospitalCH)
        alstrParametros(1) = "Lista;" & Me.txtDescripcion.Text
        alstrParametros(2) = "Departamento;" & Me.cboDepartamento.Text
        pCargaParameterFields alstrParametros, vgrptReporte
        pImprimeReporte vgrptReporte, rspvselElementosListasPrecios, "P", "Lista de precios"
    End If
    rspvselElementosListasPrecios.Close
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdImprimir_Click"))
End Sub

Private Sub cmdInvertir_Click()
    On Error GoTo NotificaError

    Dim lngContador As Long
    
    If grdPrecios.Rows > 1 And grdPrecios.TextMatrix(lngContador + 1, 1) <> "" Then
        For lngContador = 1 To grdPrecios.Rows - 1
            grdPrecios.TextMatrix(lngContador, 0) = IIf(Trim(grdPrecios.TextMatrix(lngContador, 0)) = "*", "", "*")
            
            If grdPrecios.TextMatrix(lngContador, 0) = "*" Then
                llngMarcados = llngMarcados + 1
            Else
                llngMarcados = llngMarcados - 1
            End If
        Next lngContador
        pHabilitaModificar
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdInvertir_Click"))
End Sub

Private Sub pHabilitaModificar()
    On Error GoTo NotificaError

    txtPorciento.Enabled = llngMarcados <> 0
    lblPorciento.Enabled = llngMarcados <> 0
    cmdMas.Enabled = llngMarcados <> 0
    cmdMenos.Enabled = llngMarcados <> 0
    cmdRec.Enabled = llngMarcados <> 0
    optModificar(0).Enabled = llngMarcados <> 0
    optModificar(1).Enabled = llngMarcados <> 0
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pHabilitaModificar"))
End Sub

Private Sub cmdRec_Click()
    On Error GoTo NotificaError

    Dim lngContador As Long
    Dim vlRespuesta As Integer
    
    vlRespuesta = 0
    If Not vlblnRecalcularImportar Then
        '¿Está seguro que desea recalcular los precios de la lista?
        vlRespuesta = MsgBox(SIHOMsg(990), vbYesNo + vbQuestion, "Mensaje")
    End If
        
    If vlRespuesta = 6 Or vlblnRecalcularImportar Then
    
        With grdPrecios
            For lngContador = 1 To .Rows - 1
                If .TextMatrix(lngContador, 0) = "*" Then
                    pCalcularPrecio lngContador
                   .TextMatrix(lngContador, cintColModificado) = "*"
                   .TextMatrix(lngContador, 0) = ""
                End If
            Next
        End With
        
        llngMarcados = 0
        pHabilitaModificar
        vlblnRecalcularImportar = False
        cmdInvertir.SetFocus
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdMenos_Click"))

End Sub

Private Sub Form_Load()
    On Error GoTo NotificaError

    Dim vlstrSentencia As String
    Dim rsConceptoFacturacion As New ADODB.Recordset
    Dim rsCveLista As New ADODB.Recordset
    
        'Color de Tab
    SetStyle sstListas.hwnd, 0
    SetSolidColor sstListas.hwnd, 16777215
    SSTabSubclass sstListas.hwnd
    
    Me.Icon = frmMenuPrincipal.Icon
    
    lblnPermisoCosto = fblnRevisaPermiso(vglngNumeroLogin, cintNumOpcionListas, "C")
    UpDown1.Clear
    UpDown1.AddItem cstrUltimaCompra
    UpDown1.AddItem cstrCompraMasAlta
    UpDown1.AddItem cstrPrecioMaximoPublico
    
    ' Definir el grid y el textbox con la misma letra
    txtPrecio.Font.Name = grdPrecios.Font.Name
    txtPrecio.Font.Size = grdPrecios.Font.Size
    txtPrecio.Font.Weight = grdPrecios.Font.Weight
    frmMantoListasPrecios.Height = cgIntAltoVentanaMin
    vglngDesktop = SysInfo1.WorkAreaHeight
    
    cmdExportar.Enabled = False
    'Se agrego para el botón importar
    cmdImportar.Enabled = False
       
    
    If Not fbCboLlenaDepartamento_new(cboDepartamento) Then
       MsgBox SIHOMsg(13), vbCritical, "Mensaje"
    Else
       cboDepartamento.ListIndex = 0
    End If
    
    If Not vpblnEsSistemas Then
       cboDepartamento.ListIndex = fintLocalizaCbo_new(cboDepartamento, STR(vgintNumeroDepartamento))
       cboDepartamento.Enabled = False
    End If
   
    vgstrEstadoManto = ""
    pconfiguragrid
    
    cboTipoCargo.AddItem "<TODOS>", cintIndexTodos
    cboTipoCargo.AddItem "ARTÍCULOS", cintIndexArticulo
    cboTipoCargo.AddItem "ESTUDIOS", cintIndexEstudio
    cboTipoCargo.AddItem "EXÁMENES", cintIndexExamen
    cboTipoCargo.AddItem "EXÁMENES Y GRUPOS", cintIndexExamenGrupo
    cboTipoCargo.AddItem "GRUPOS DE EXÁMENES", cintIndexGrupo
    cboTipoCargo.AddItem "OTROS CONCEPTOS", cintIndexOtro
    cboTipoCargo.AddItem "PAQUETES", cintIndexPaquete
    
    'Conceptos de faturación
    Set rsConceptoFacturacion = frsEjecuta_SP("0|1|-1", "sp_PvSelConceptoFactura")
    If rsConceptoFacturacion.recordCount > 0 Then
        pLlenarCboRs_new cboConceptoFacturacion, rsConceptoFacturacion, 0, 1, 3
        cboConceptoFacturacion.ListIndex = 0
    End If
    
    Set rsListasDepto = frsEjecuta_SP(CStr(cboDepartamento.ItemData(cboDepartamento.ListIndex)), "SP_PVSELLISTADEPTO")
    
    cboTipoCargo_Click
    
    ' Valida la opción cuando es llamada del catálogo de artículos
    
    If blnCatalogo = True Then
        Set rsCveLista = frsEjecuta_SP(CStr(cboDepartamento.ItemData(cboDepartamento.ListIndex)), "SP_PVCVELISTA")
        
        If Not rsCveLista.EOF Then
            pPreparaPantalla
             cboTipoCargo.ListIndex = cintIndexTodos
             pHabilita 0, 0, 0, 0, 0, 1, 0, 0
             
             txtClave.Text = rsCveLista(0).Value
            
             If fblnExiste() Then
                 'Consulta:
                 pModificaRegistro
                 pCargaOtrasListas Val(txtClave.Text)
                 pEnfocaTextBox txtDescripcion
             Else
                 'Alta:
                 vgstrEstadoManto = "A" 'Alta
                 chkPredeterminada.Enabled = fintListasDepto >= 1
                 If Not chkPredeterminada.Enabled Then
                     chkPredeterminada.Value = 1
                 End If
                 pCargaOtrasListas -1
                 pEnfocaTextBox txtDescripcion
             End If
             cmdCapturaPrecios_Click
        Else
            blnCatalogo = False
        End If
        rsCveLista.Close
    End If
    
    pCargaTabuladores
    
    If Screen.Height <= 12050 Then
        pcambiarT2
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_Load"))
End Sub

Private Sub pcambiarT2()

End Sub


Private Sub pCargaBusqueda(Optional nada As Integer)
    On Error GoTo NotificaError

    Dim vlstrSentencia As String
    Dim vlintContador As Integer
    
    'Para llenar el grid de la consulta de las listas
    pConfiguraGridBusqueda
    
    With grdHBusqueda
        If rsListasDepto.recordCount <> 0 Then
            rsListasDepto.MoveFirst
            Do While Not rsListasDepto.EOF
                .Row = .Rows - 1
                .TextMatrix(.Row, 1) = rsListasDepto!INTCVELISTA
                .TextMatrix(.Row, 2) = rsListasDepto!chrDescripcion
                .TextMatrix(.Row, 3) = IIf(rsListasDepto!bitestatusactivo = 1, "ACTIVO", "INACTIVO")
                
                If rsListasDepto!bitPredeterminada = 1 Then
                    For vlintContador = 1 To .Cols - 1
                        .Col = vlintContador
                        .CellForeColor = clngAzul
                    Next
                End If
                
                .Rows = .Rows + 1
                rsListasDepto.MoveNext
            Loop
            .Rows = .Rows - 1
        End If
    End With
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargaBusqueda"))
End Sub

Private Sub cmdBuscar_Click()
    On Error GoTo NotificaError

    pCargaBusqueda
    vgstrEstadoManto = "B"
    sstListas.Tab = 1
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdBuscar_Click"))
End Sub

Private Sub cmdCapturaPrecios_Click()
    On Error GoTo NotificaError

    If vgstrEstadoManto = "A" Or vgstrEstadoManto = "M" Then
        pHabilita 0, 0, 0, 0, 0, 1, 0, 0
        pconfiguragrid
        pPreparaPantalla
        
        lblTipoCargo.Enabled = lblnConsulta
        cboTipoCargo.Enabled = lblnConsulta
        cboTipoCargo.ListIndex = cintIndexTodos
        cboConceptoFacturacion.Enabled = lblnConsulta
        cboConceptoFacturacion.ListIndex = cintIndexTodos
        cmdFiltrar.Enabled = lblnConsulta
        
        cmdGrabarRegistro.Enabled = True
        cmdFiltrar.Enabled = True
        If blnCatalogo = False Then cmdFiltrar.SetFocus
    Else
        If Not vlblnCancelCaptura Then
            If MsgBox(SIHOMsg(17), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
                pEnfocaTextBox txtClave
                If blnCatalogo = True Then
                    pNuevoRegistro
                    frmMantoListasPrecios.Height = cgIntAltoVentanaMin
                    pHabilita 0, 0, 1, 0, 0, 0, 0, 0
                End If
            End If
        Else
            pEnfocaTextBox txtClave
        End If
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdCapturaPrecios_Click"))
End Sub

Private Sub pLimpiaSeleccionGrid()
On Error GoTo NotificaError

    Dim lngContador As Long
    
    For lngContador = 1 To grdPrecios.Rows - 1
        grdPrecios.TextMatrix(lngContador, 0) = ""
        
    Next lngContador

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & "pLimpiaSeleccionGrid"))
End Sub


Private Sub pPreparaPantalla()
    On Error GoTo NotificaError

    Dim vlintContador As Integer

    vgstrEstadoManto = vgstrEstadoManto & "E" 'Pantalla expandida
    
    '------------------------------
    'Abrir la forma
    '------------------------------
    frmMantoListasPrecios.Refresh
    For vlintContador = frmMantoListasPrecios.Height To cgIntAltoVentanaMax Step cgintFactorMovVentana
        frmMantoListasPrecios.Height = vlintContador
        frmMantoListasPrecios.Top = Int((vglngDesktop - frmMantoListasPrecios.Height) / 2)
    Next
    frmMantoListasPrecios.Height = cgIntAltoVentanaMax
    frmMantoListasPrecios.Top = Int((vglngDesktop - frmMantoListasPrecios.Height) / 2)
    cmdCapturaPrecios.Caption = "Cancelar captura"
    frmMantoListasPrecios.Refresh
    
    '------------------------------
    'Deshabilitar Botonera
    '------------------------------
    cmdBuscar.Enabled = False
    cmdGrabarRegistro.Enabled = True

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pPreparaPantalla"))
End Sub

Private Sub pCargaLista(lngCveLista As Long, strTipoCargo As String, strMarca As String, lngFacturacion As Long)
    On Error GoTo NotificaError

    Dim rs As New ADODB.Recordset
    Dim intcontador As Integer
    Dim blnNuevos As Boolean
    Dim dblAvance As Double
    Dim intRenglon As Integer
    
    vgstrParametrosSP = CStr(cboDepartamento.ItemData(cboDepartamento.ListIndex)) & "|" & CStr(lngCveLista) & "|" & strTipoCargo & "|" & CStr(lngFacturacion)
    If lblTipoArticulo.Enabled Then
        vgstrParametrosSP = vgstrParametrosSP & "|" & IIf(cboArtMed.ItemData(cboArtMed.ListIndex) = 0, 3, IIf(cboArtMed.ItemData(cboArtMed.ListIndex) = 1, 0, 1)) & "|" & cboFamilia.ItemData(cboFamilia.ListIndex) & "|" & cboSubFamilia.ItemData(cboSubFamilia.ListIndex)
    Else
        vgstrParametrosSP = vgstrParametrosSP & "|3|0|0"
    End If
    If lblClasificacion.Enabled Then
        vgstrParametrosSP = vgstrParametrosSP & "|" & cboClasificacionSA.ItemData(cboClasificacionSA.ListIndex)
    Else
        vgstrParametrosSP = vgstrParametrosSP & "|0"
    End If
        vgstrParametrosSP = vgstrParametrosSP & "|" & Trim(txtClaveArticulo.Text)
    
    If Chkbitcheckup.Value = 1 Then
        vgstrParametrosSP = vgstrParametrosSP & "|1"
    Else
        vgstrParametrosSP = vgstrParametrosSP & "|0"
    End If
    ' En la consulta se muestran todos los elementos de la lista (Tengan precio cero o no)
    vgstrParametrosSP = vgstrParametrosSP & "|1"
    Set rs = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELELEMENTOSLISTASPRECIOS")
    
    lblDescripcionCargo.Enabled = rs.recordCount <> 0
    txtIniciales.Enabled = rs.recordCount <> 0
    
    If rs.recordCount <> 0 Then
        dblAvance = 100 / rs.recordCount
        '------------------------
        ' Configuración de la Barra de estado
        '------------------------
        lblTextoBarra.Caption = SIHOMsg(280)
        freBarra.Visible = True
        freBarra.Top = 720
        pgbBarra.Value = 0
        freBarra.Refresh
    
        blnNuevos = False
        pgbBarra.Value = 0
        grdPrecios.Visible = False
        
        Do While Not rs.EOF
            ' Actualización de la barra de estado
            If pgbBarra.Value + dblAvance < 100 Then
                pgbBarra.Value = pgbBarra.Value + dblAvance
            Else
                pgbBarra.Value = 100
            End If
            pgbBarra.Refresh
       
            grdPrecios.TextMatrix(grdPrecios.Rows - 1, cintColClave) = rs!clave
            grdPrecios.TextMatrix(grdPrecios.Rows - 1, cintColDescripcion) = rs!Descripcion
            grdPrecios.TextMatrix(grdPrecios.Rows - 1, cintColPrecio) = Format(rs!precio, "$###,###,###,##0.00####")
            grdPrecios.TextMatrix(grdPrecios.Rows - 1, cintColMoneda) = rs!TipoMoneda
            grdPrecios.TextMatrix(grdPrecios.Rows - 1, cintColTipo) = rs!Tipo
            grdPrecios.TextMatrix(grdPrecios.Rows - 1, cintColTipoDes) = rs!TipoDescripcion
            grdPrecios.TextMatrix(grdPrecios.Rows - 1, cintColModificado) = strMarca
            grdPrecios.TextMatrix(grdPrecios.Rows - 1, cintColNuevo) = rs!nuevo
            grdPrecios.TextMatrix(grdPrecios.Rows - 1, cintColCveFact) = rs!CveConceptoFacturacion
            grdPrecios.TextMatrix(grdPrecios.Rows - 1, cintColFacturacion) = rs!ConceptoFacturacion
            grdPrecios.TextMatrix(grdPrecios.Rows - 1, cintColCosto) = Format(IIf(rs!tipoIncremento = "C", rs!costo, IIf(rs!tipoIncremento = "M", rs!PrecioMaximoPublico, rs!CostoMasAlto)), "$###,###,###,##0.0000##")
            grdPrecios.TextMatrix(grdPrecios.Rows - 1, cintColTipoIncremento) = IIf(rs!Tipo = "AR", IIf(rs!tipoIncremento = "C", cstrUltimaCompra, IIf(rs!tipoIncremento = "M", cstrPrecioMaximoPublico, cstrCompraMasAlta)), "NA")
            grdPrecios.TextMatrix(grdPrecios.Rows - 1, cintColUtilidad) = Format(rs!margenUtilidad, "0.0000") & "%"
            grdPrecios.TextMatrix(grdPrecios.Rows - 1, cintColCostoUltimaEntrada) = rs!costo
            grdPrecios.TextMatrix(grdPrecios.Rows - 1, cintColCostoMasAlto) = rs!CostoMasAlto
            grdPrecios.TextMatrix(grdPrecios.Rows - 1, cintColPrecioMaximopublico) = rs!PrecioMaximoPublico
            grdPrecios.TextMatrix(grdPrecios.Rows - 1, cintColIncremetoAutomatico) = IIf(rs!IncrementoAutomatico = 0, "", "*")
            grdPrecios.TextMatrix(grdPrecios.Rows - 1, cintColTabulador) = IIf(rs!tabulador = 0, "", "*")
            If IsNull(rs!precioespecifico) Then
                grdPrecios.TextMatrix(grdPrecios.Rows - 1, cintColPrecioEspecifico) = "0"
            Else
                grdPrecios.TextMatrix(grdPrecios.Rows - 1, cintColPrecioEspecifico) = Val(rs!precioespecifico)
            End If
            
            If rs!nuevo = 1 Then
                blnNuevos = True
                For intcontador = 1 To grdPrecios.Cols - 1
                    grdPrecios.Col = intcontador
                    grdPrecios.Row = grdPrecios.Rows - 1
                    grdPrecios.CellForeColor = clngRojo
                    grdPrecios.CellFontBold = True
                Next intcontador
                 
                 ' Valida si fue llamado del catálogo de artículos y se posiciona en el artículo nuevo
                 If blnCatalogo = True Then
                        grdPrecios.TextMatrix(grdPrecios.Row, 0) = IIf(rs!clave = StrCveArticulo, "*", "")
                        If grdPrecios.TextMatrix(grdPrecios.Row, 0) <> "" Then
                            intRenglon = grdPrecios.Row
                            llngMarcados = 1
                        End If
                End If
            End If
            
            grdPrecios.Rows = grdPrecios.Rows + 1
            rs.MoveNext
        Loop

        grdPrecios.Rows = grdPrecios.Rows - 1
        grdPrecios.Visible = True
        
        If blnNuevos Then
            'IMPORTANTE : Se encontró nuevos elementos en la lista (marcados con rojo), presione el botón de grabar para que estos precios se activen!
            MsgBox SIHOMsg(789), vbOKOnly + vbInformation, "Mensaje"
        End If
        
        freBarra.Visible = False
        
        grdPrecios.Row = 1
        grdPrecios.Col = cintColPrecio
        
        ' Valiada sí fue llamado del catálogo de artículos y pone el cursor en el elemento
        
        If blnCatalogo = True Then
            grdPrecios.Row = intRenglon
            Call pEditarColumna(32, txtPrecio, grdPrecios)
            txtPrecio.Visible = True
            txtPrecio.SetFocus
        End If
        cmdExportar.Enabled = True
        cmdExportar.SetFocus
        'Se agrega para el botón importar
        cmdImportar.Enabled = True
        pMuestraNomComercialCompleto
    Else
        cmdExportar.Enabled = False
        'Se agrega para el botón importar
        cmdImportar.Enabled = False
        'No existe información con esos parámetros.
        MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje"
        
        lblNombreCompleto.Caption = ""
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargaLista"))
End Sub

Private Sub cmdInicializa_Click()
    On Error GoTo NotificaError

    pconfiguragrid
    pPreparaPantalla
    pCargaLista cboOtrasListas.ItemData(cboOtrasListas.ListIndex), "*", "*", 0
    blnactivainicializa = 1
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdInicializa_Click"))
End Sub

Private Sub cmdMas_Click()
    On Error GoTo NotificaError

    Dim lngContador As Long
    '¿Está seguro que desea incrementar los precios de la lista en ese porcentaje?
    If MsgBox(SIHOMsg(IIf(optModificar(0).Value, 272, 905)), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
        With grdPrecios
            For lngContador = 1 To .Rows - 1
                If .TextMatrix(lngContador, 0) = "*" Then
                    If optModificar(0).Value Then
                        .TextMatrix(lngContador, cintColPrecio) = Format(.TextMatrix(lngContador, cintColPrecio) * (1 + (Val(txtPorciento.Text) / 100)), "$###,###,###,##0.00####")
                        .TextMatrix(lngContador, cintColIncremetoAutomatico) = ""
                    Else
                        .TextMatrix(lngContador, cintColUtilidad) = txtPorciento.Text & "%"
                        pCalcularPrecio lngContador
                    End If
                    .TextMatrix(lngContador, cintColModificado) = "*"
                    .TextMatrix(lngContador, 0) = ""
                End If
            Next lngContador
        End With
        llngMarcados = 0
        txtPorciento.Text = ""
        pHabilitaModificar
        cmdInvertir.SetFocus
    End If
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdMas_Click"))
End Sub

Private Sub cmdMenos_Click()
    On Error GoTo NotificaError

    Dim lngContador As Long
    
    '¿Está seguro que desea disminuir los precios de la lista en ese porcentaje?
    If MsgBox(SIHOMsg(275), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
    
        With grdPrecios
            For lngContador = 1 To .Rows - 1
                If .TextMatrix(lngContador, 0) = "*" Then
                   .TextMatrix(lngContador, cintColPrecio) = Format(.TextMatrix(lngContador, cintColPrecio) - (.TextMatrix(lngContador, cintColPrecio) * (Val(txtPorciento.Text) / 100)), "$###,###,###,##0.00####")
                   .TextMatrix(lngContador, cintColModificado) = "*"
                   .TextMatrix(lngContador, 0) = ""
                End If
            Next
        End With
        
        llngMarcados = 0
        txtPorciento.Text = ""
        pHabilitaModificar
        cmdInvertir.SetFocus
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdMenos_Click"))
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError

    If KeyAscii = vbKeyEscape And Me.ActiveControl.Name <> "UpDown1" And ActiveControl.Name <> "lstPrecios" Then
        KeyAscii = 0
        Unload Me
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_KeyPress"))
End Sub

Private Sub pconfiguragrid()
    On Error GoTo NotificaError
    
    lblDescripcionCargo.Enabled = False
    txtIniciales.Enabled = False
    cmdExportar.Enabled = False
    'Se agrega para el botón importar
    cmdImportar.Enabled = False
    With grdPrecios
        .Clear
        .Cols = cintColumnas
        .Rows = 2
        .FixedCols = 1
        .FixedRows = 1
        .FormatString = "|Clave|Descripción cargo|Incremento automático|Tipo incremento|Margen utilidad|Usar tabulador|Costo base|Precio|Moneda|Tipo|Tipo|Modificado|Nuevo||Concepto de facturación||||"
        .RowHeight(0) = 580
        .RowHeight(1) = 280
        .ColWidth(0) = 150  'Fix

        .ColWidth(cintColClave) = 1150
        .ColWidth(cintColIncremetoAutomatico) = 1300
        .ColWidth(cintColTipoIncremento) = 1300
        .ColWidth(cintColUtilidad) = 1000
        .ColWidth(cintColTabulador) = 1100
        .ColWidth(cintColDescripcion) = 4000
        .ColWidth(cintColCosto) = 1300
        .ColWidth(cintColPrecio) = 1100
        .ColWidth(cintColMoneda) = 1000
        .ColWidth(cintColTipo) = 0
        .ColWidth(cintColTipoDes) = 1900
        .ColWidth(cintColModificado) = 0
        .ColWidth(cintColNuevo) = 0
        .ColWidth(cintColCveFact) = 0
        .ColWidth(cintColFacturacion) = 3500
        .ColWidth(cintColCostoUltimaEntrada) = 0
        .ColWidth(cintColCostoMasAlto) = 0
        .ColWidth(cintColPrecioMaximopublico) = 0
        .ColWidth(cintColPrecioEspecifico) = 0
        
        .ColAlignmentFixed(cintColClave) = flexAlignCenterCenter
        .ColAlignmentFixed(cintColDescripcion) = flexAlignCenterCenter
        .ColAlignmentFixed(cintColIncremetoAutomatico) = flexAlignCenterCenter
        .ColAlignmentFixed(cintColTipoIncremento) = flexAlignCenterCenter
        .ColAlignmentFixed(cintColUtilidad) = flexAlignCenterCenter
        .ColAlignmentFixed(cintColTabulador) = flexAlignCenterCenter
        .ColAlignmentFixed(cintColPrecio) = flexAlignCenterCenter
        .ColAlignmentFixed(cintColMoneda) = flexAlignCenterCenter
        .ColAlignmentFixed(cintColCosto) = flexAlignCenterCenter
        .ColAlignmentFixed(cintColTipoDes) = flexAlignCenterCenter
        .ColAlignmentFixed(cintColFacturacion) = flexAlignCenterCenter
        
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(cintColClave) = flexAlignLeftCenter
        .ColAlignment(cintColDescripcion) = flexAlignLeftCenter
        .ColAlignment(cintColIncremetoAutomatico) = flexAlignCenterCenter
        .ColAlignment(cintColCosto) = flexAlignRightCenter
        .ColAlignment(cintColTabulador) = flexAlignCenterCenter
        .ColAlignment(cintColPrecio) = flexAlignRightCenter
        .ColAlignment(cintColFacturacion) = flexAlignLeftCenter
    End With

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pConfiguraGrid"))
End Sub

Private Sub cmdAnteriorRegistro_Click()
    On Error GoTo NotificaError

    If grdHBusqueda.Row > 1 Then
        grdHBusqueda.Row = grdHBusqueda.Row - 1
    End If
    txtClave.Text = grdHBusqueda.TextMatrix(grdHBusqueda.Row, 1)
    If fblnExiste() Then
        pModificaRegistro
        pCargaOtrasListas Val(txtClave.Text)
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdAnteriorRegistro_Click"))
End Sub

Private Sub cmdPrimerRegistro_Click()
    On Error GoTo NotificaError

    grdHBusqueda.Row = 1
    txtClave.Text = grdHBusqueda.TextMatrix(grdHBusqueda.Row, 1)
    If fblnExiste() Then
        pModificaRegistro
        pCargaOtrasListas Val(txtClave.Text)
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdPrimerRegistro_Click"))
End Sub

Private Sub cmdSiguienteRegistro_Click()
    On Error GoTo NotificaError

    If grdHBusqueda.Row < grdHBusqueda.Rows - 1 Then
        grdHBusqueda.Row = grdHBusqueda.Row + 1
    End If
    txtClave.Text = grdHBusqueda.TextMatrix(grdHBusqueda.Row, 1)
    If fblnExiste() Then
        pModificaRegistro
        pCargaOtrasListas Val(txtClave.Text)
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdSiguienteRegistro_Click"))
End Sub

Private Sub cmdUltimoRegistro_Click()
    On Error GoTo NotificaError

    grdHBusqueda.Row = grdHBusqueda.Rows - 1
    txtClave.Text = grdHBusqueda.TextMatrix(grdHBusqueda.Row, 1)
    If fblnExiste() Then
        pModificaRegistro
        pCargaOtrasListas Val(txtClave.Text)
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdUltimoRegistro_Click"))
End Sub

Private Sub pConfiguraGridBusqueda()
    On Error GoTo NotificaError

    With grdHBusqueda
        .Clear
        .Rows = 2
        .FormatString = "|Clave|Descripción|Estado"
        .ColWidth(0) = 100 'Fix
        .ColWidth(1) = 1000 'Clave
        .ColWidth(2) = 7000 'Descripcion
        .ColWidth(3) = 1500 'Estado
        .ColAlignment(1) = flexAlignRightCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignLeftCenter
        .ColAlignmentFixed(3) = flexAlignCenterCenter
        .ColAlignmentFixed(1) = flexAlignCenterCenter
        .ScrollBars = flexScrollBarVertical
        .TextMatrix(1, 1) = ""
        .TextMatrix(1, 2) = ""
    End With
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pConfiguraGridBusqueda"))
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo NotificaError

    Select Case vgstrEstadoManto
        Case "A", "M"
            Cancel = 1
            If MsgBox(SIHOMsg(17), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
                pEnfocaTextBox txtClave
            End If
        Case "AE", "ME"
            Cancel = 1
            cmdCapturaPrecios_Click
        Case "AEE", "MEE"
            Cancel = 1
            vgstrEstadoManto = Mid(vgstrEstadoManto, 1, 2)
        Case "B"
            Cancel = 1
            sstListas.Tab = 0
            pEnfocaTextBox txtClave
    End Select
    
    If blnCatalogo = True Then
        blnCatalogo = False
        pNuevoRegistro
        frmMantoListasPrecios.Height = cgIntAltoVentanaMin
        pHabilita 0, 0, 1, 0, 0, 0, 0, 0
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_QueryUnload"))
End Sub

Private Sub Frame2_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub grdHBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then
        grdHBusqueda_DblClick
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":grdHBusqueda_KeyDown"))
End Sub

Private Sub grdPrecios_LostFocus()
    On Error GoTo NotificaError

    vgstrAcumTextoBusqueda = ""
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":grdPrecios_LostFocus"))
End Sub

Private Sub lstPesos_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        lstPesos_MouseUp 0, 0, 0, 0
        If grdPrecios.Row < grdPrecios.Rows - 1 Then
            grdPrecios.Row = grdPrecios.Row + 1
        Else
            grdPrecios.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        grdPrecios.SetFocus
        lstPesos.Visible = False
    End If
End Sub

Private Sub lstPesos_LostFocus()
    lstPesos.Visible = False
End Sub

Private Sub lstPesos_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo NotificaError
    grdPrecios.Text = lstPesos.Text
    grdPrecios.TextMatrix(grdPrecios.Row, cintColModificado) = "*"
    grdPrecios.SetFocus
    lstPesos.Visible = False
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":lstPesos_MouseUp"))
End Sub

Private Sub lstPesos_Validate(Cancel As Boolean)
On Error GoTo NotificaError
    lstPesos_MouseUp 0, 0, 0, 0
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":lstPesos_Validate"))
End Sub

Private Sub optModificar_Click(Index As Integer)
    On Error GoTo NotificaError

    cmdMas.Caption = IIf(Index = 1, "Cambiar", "Aumentar")
    cmdMenos.Visible = IIf(Index = 1, False, True)
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":optModificar_Click"))
End Sub



Private Sub tmrDespliega_Timer()

    cmdFiltrar_Click
    tmrDespliega.Enabled = False
    
End Sub

Private Sub TxtClave_GotFocus()
    On Error GoTo NotificaError
    
    If blnCatalogo = False Then
        pNuevoRegistro
        pHabilita 0, 0, 1, 0, 0, 0, 0, 0
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtClave_GotFocus"))
End Sub

Private Sub pNuevoRegistro()
    On Error GoTo NotificaError
    Dim vlintContador As Integer
    
    lblnConsulta = False
    llngMarcados = 0
    
    txtClave.Text = frsEjecuta_SP("", "SP_PVSELCONSECUTIVOLISTAPRECIO").Fields(0)
    txtDescripcion.Text = ""
    txtDescripcion_Change
    txtIniciales.Text = ""
    cboTabulador.ListIndex = 0
    chkActivo.Value = 1
    
    chkPredeterminada.Enabled = fintListasDepto >= 1
    If Not chkPredeterminada.Enabled Then
        chkPredeterminada.Value = 1
    Else
        chkPredeterminada.Value = 0
    End If
    
    vgblnEditaPrecio = False 'Para saber si esta editando un precio
    
    If Len(vgstrEstadoManto) > 1 Then
        '------------------------------
        'Cerrar la forma
        '------------------------------
        For vlintContador = cgIntAltoVentanaMax To cgIntAltoVentanaMin Step (-1 * cgintFactorMovVentana)
            frmMantoListasPrecios.Height = vlintContador
            frmMantoListasPrecios.Top = Int((vglngDesktop - frmMantoListasPrecios.Height) / 2)
        Next
        frmMantoListasPrecios.Height = cgIntAltoVentanaMin
        
        cmdCapturaPrecios.Caption = "Capturar precios"
    End If
    vgstrEstadoManto = "" 'El inicio de la pantalla
    
    'Limpia grid precios:
    pconfiguragrid
    
    txtPorciento.Text = ""
    txtPorciento_Change
    
    pHabilitaModificar
    
    lblPredeterminada.ForeColor = clngAzul
    
    pEnfocaTextBox txtClave

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pNuevoRegistro"))
End Sub

Private Sub TxtClave_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    '-------------------------------------------------------------------------------------------
    'Validación para diferenciar cuando es una alta de un registro o cuando se va a consultar o
    'modificar uno que ya existe
    '-------------------------------------------------------------------------------------------
    If KeyCode = vbKeyReturn Then
        If fblnExiste() Then
            'Consulta:
            pModificaRegistro
            pCargaOtrasListas Val(txtClave.Text)
            pHabilita 1, 1, 1, 1, 1, 0, 1, 1
            'cmdBuscar.SetFocus
            cmdCapturaPrecios.SetFocus
            chkImprimeCargosPrecioCero.Value = 1
            Chkbitcheckup.Value = 0
            sstListas.Tab = 0
            
        Else
            'Alta:
            vgstrEstadoManto = "A" 'Alta
            chkPredeterminada.Enabled = fintListasDepto >= 1
            If Not chkPredeterminada.Enabled Then
                chkPredeterminada.Value = 1
            End If
            pCargaOtrasListas -1
            pEnfocaTextBox txtDescripcion
        End If
    End If
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtClave_KeyDown"))
End Sub

Private Sub pCargaOtrasListas(lngCveLista As Long)
    On Error GoTo NotificaError
    Dim intcontador As Integer

    cboOtrasListas.Clear

    If rsListasDepto.recordCount <> 0 Then
        rsListasDepto.MoveFirst
        intcontador = 0
        Do While Not rsListasDepto.EOF
            If lngCveLista <> rsListasDepto!INTCVELISTA Then
                cboOtrasListas.AddItem Trim(rsListasDepto!chrDescripcion), intcontador
                cboOtrasListas.ItemData(cboOtrasListas.NewIndex) = rsListasDepto!INTCVELISTA
                intcontador = intcontador + 1
            End If
            rsListasDepto.MoveNext
        Loop
    End If
    
    If cboOtrasListas.ListCount <> 0 Then
        cboOtrasListas.ListIndex = 0
    End If

    cmdInicializa.Enabled = cboOtrasListas.ListCount <> 0

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargaOtrasListas"))
End Sub

Private Function fintListasDepto()
    On Error GoTo NotificaError

    Dim rs As New ADODB.Recordset

    Set rs = frsEjecuta_SP(CStr(cboDepartamento.ItemData(cboDepartamento.ListIndex)), "SP_PVSELLISTADEPTO")
    fintListasDepto = rs.recordCount

    Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fintListasDepto"))
End Function

Private Sub pModificaRegistro()
    On Error GoTo NotificaError

    '-------------------------------------------------------------------------------------------
    ' Permite realizar la modificación de la descripción de un registro
    '-------------------------------------------------------------------------------------------
    
    lblnConsulta = True
    vgstrEstadoManto = "M" 'Modificacion
    
    txtClave.Text = rsListasDepto!INTCVELISTA
    txtDescripcion.Text = Trim(rsListasDepto!chrDescripcion)
    cboTabulador.ListIndex = fintLocalizaCbo_new(cboTabulador, rsListasDepto!intCveTabulador)
    
    chkActivo.Value = rsListasDepto!bitestatusactivo
    chkPredeterminada.Value = rsListasDepto!bitPredeterminada
    
    chkPredeterminada.Enabled = fintListasDepto >= 1
    If Not chkPredeterminada.Enabled Then
        chkPredeterminada.Value = 1
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pModificaRegistro"))
End Sub
Private Sub cmdGrabarRegistro_Click()
    On Error GoTo NotificaError
    Dim lngIdLista As Long
    Dim lngContador As Long
    Dim llngPersonaGraba  As Long
    Dim vlstrsql As String
    Dim X As Long
    Dim Y As Integer
    Dim cont As Integer
    Dim rs As New ADODB.Recordset
    Dim vlstrotro As String
    Dim vlstrCveCargocompleto As String
    Dim vlstrsqlpaquete As String
    Dim rsPaquete As New ADODB.Recordset
    Dim rsListaemp As New ADODB.Recordset
    Dim rsListapac As New ADODB.Recordset
    Dim vldblPrecio As Double
    Dim vldblPreciosuma As Double
    Dim vlstrsqlarticulo As String
    Dim rsArticulo As New ADODB.Recordset
    Dim vlstrSuma As String
    Dim rssuma As New ADODB.Recordset
    Dim vlstrtieneinccremento As String
    Dim rstieneincremento As New ADODB.Recordset

    If Trim(txtDescripcion.Text) = "" Then
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbOKOnly + vbInformation, "Mensaje"
        txtDescripcion.SetFocus
    Else
        If Not lblnConsulta And Trim(grdPrecios.TextMatrix(1, cintColClave)) = "" Then
            '¡No ha ingresado datos!
            MsgBox SIHOMsg(2), vbOKOnly + vbInformation, "Mensaje"
            cmdCapturaPrecios.SetFocus
        Else
            'Verificar que la lista no este asignada en caso de querer desactivarla
            If chkActivo.Value = 0 Then
                Set rsListaemp = frsRegresaRs("select * from pvlistaempresa pvle where pvle.INTCVELISTA = " & txtClave.Text)
                Set rsListapac = frsRegresaRs("select * from pvlistatipopaciente pvl where pvl.INTCVELISTA = " & txtClave.Text)
                If rsListaemp.recordCount > 0 Or rsListapac.recordCount > 0 Then
                    'No es posible desactivar una lista de precios que se encuentra asignada.
                    MsgBox SIHOMsg(1186), vbOKOnly + vbInformation, "Mensaje"
                    Exit Sub
                End If
            End If
            If fblnRevisaPermiso(vglngNumeroLogin, cintNumOpcionListas, "E") Then
                llngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
                If llngPersonaGraba <> 0 Then
                
                    EntornoSIHO.ConeccionSIHO.BeginTrans
                    
                        If chkPredeterminada.Value = 1 Then
                            frsEjecuta_SP CStr(cboDepartamento.ItemData(cboDepartamento.ListIndex)), "SP_PVUPDCAMBIAPREDETERMINADA"
                        End If
                        '--*--*--*--*--*--*--*--*--*--*--*
                        '1-- Graba o actualiza maestro
                        '--*--*--*--*--*--*--*--*--*--*--*
                        If Not lblnConsulta Then
                            vgstrParametrosSP = Trim(txtDescripcion.Text) & "|" & CStr(cboDepartamento.ItemData(cboDepartamento.ListIndex)) & "|" & CStr(chkActivo.Value) & "|" & CStr(chkPredeterminada.Value) & "|" & cboTabulador.ItemData(cboTabulador.ListIndex)
                            lngIdLista = 1
                            frsEjecuta_SP vgstrParametrosSP, "SP_PVINSLISTAPRECIO", True, lngIdLista
                            txtClave.Text = lngIdLista
                        Else
                            lngIdLista = Val(txtClave.Text)
                            vgstrParametrosSP = CStr(lngIdLista) & "|" & Trim(txtDescripcion.Text) & "|" & CStr(chkActivo.Value) & "|" & CStr(chkPredeterminada.Value) & "|" & cboTabulador.ItemData(cboTabulador.ListIndex)
                            frsEjecuta_SP vgstrParametrosSP, "SP_PVUPDLISTAPRECIO"
                        End If
                        
                        '--*--*--*--*--*--*--*--*--*--*--*
                        '2-- Graba detalle lista
                        '--*--*--*--*--*--*--*--*--*--*--*
                        freBarra.Visible = True
                        'Generando lista de precios, por favor espere...
                        lblTextoBarra.Caption = SIHOMsg(279)
                        pgbBarra.Value = 0
                        freBarra.Refresh
        
                        For lngContador = 1 To grdPrecios.Rows - 1
                            If ((lngContador / grdPrecios.Rows) * 100) Mod 2 Then
                                pgbBarra.Value = (lngContador / grdPrecios.Rows) * 100
                            End If
                            
                            If grdPrecios.TextMatrix(lngContador, cintColModificado) = "*" Then
                            
                                vgstrParametrosSP = CStr(lngIdLista) & "|" & grdPrecios.TextMatrix(lngContador, cintColClave) & "|" & grdPrecios.TextMatrix(lngContador, cintColTipo) & "|" & CStr(Val(Format(grdPrecios.TextMatrix(lngContador, cintColPrecio), "##########0.00####"))) & "|" & IIf(grdPrecios.TextMatrix(lngContador, cintColTipoIncremento) = cstrPrecioMaximoPublico, "M", IIf(grdPrecios.TextMatrix(lngContador, cintColTipoIncremento) = cstrCompraMasAlta, "A", "C")) & "|" & Replace(grdPrecios.TextMatrix(lngContador, cintColUtilidad), "%", "") & "|" & IIf(Trim(grdPrecios.TextMatrix(lngContador, cintColTabulador)) = "*", "1", "0") & "|" & IIf(Trim(grdPrecios.TextMatrix(lngContador, cintColIncremetoAutomatico)) = "*", "1", "0") & "|" & IIf(grdPrecios.TextMatrix(lngContador, cintColMoneda) = "PESOS", "1", "0")
                                If lblnConsulta And blnactivainicializa = 1 Then
                                    pEjecutaSentencia ("Delete From Pvdetallelista where intcvelista=" & lngIdLista & "And chrcvecargo='" & grdPrecios.TextMatrix(lngContador, cintColClave) & "' and chrtipocargo='" & grdPrecios.TextMatrix(lngContador, cintColTipo) & "'")
                                    grdPrecios.TextMatrix(lngContador, cintColNuevo) = 1
                                End If
                                frsEjecuta_SP vgstrParametrosSP, IIf(Not lblnConsulta, "SP_PVINSDETALLELISTAPRECIO", IIf(Val(grdPrecios.TextMatrix(lngContador, cintColNuevo)) = 1, "SP_PVINSDETALLELISTAPRECIO", "SP_PVUPDDETALLELISTAPRECIO"))
                                
                                'Actualiza costos base
                                If lblnPermisoCosto And (grdPrecios.TextMatrix(lngContador, cintColTipoIncremento) = "PRECIO MÁXIMO AL PÚBLICO" Or grdPrecios.TextMatrix(lngContador, cintColTipoIncremento) = "NA") Then
                                    vgstrParametrosSP = vgintClaveEmpresaContable & "|" & grdPrecios.TextMatrix(lngContador, cintColTipo) & "|" & grdPrecios.TextMatrix(lngContador, cintColClave) & "|" & Format(grdPrecios.TextMatrix(lngContador, cintColCosto), "##########0.0000##") & "|" & IIf(Format(grdPrecios.TextMatrix(lngContador, cintColCosto), "0.000000") = "0.000000", "1", "0")
                                    frsEjecuta_SP vgstrParametrosSP, "sp_PVActualizaCosto"
                                End If
                            End If

                        Next lngContador
                        
                        '--*--*--*--*--*--*--*--*--*--*--*
                        '3-- Graba registro de transacciones
                        '--*--*--*--*--*--*--*--*--*--*--*
                        pGuardarLogTransaccion Me.Name, IIf(Not lblnConsulta, EnmGrabar, EnmCambiar), llngPersonaGraba, "LISTA DE PRECIOS", txtClave.Text
                        X = 1
                        Y = 1
                        
                        If chkPredeterminada = 1 Then
'                            Do While X < 1000
'                                vlstrCveCargocompleto = Mid(vlstrCveCargo, X, IIf(vlstrCveCargo <> "", InStr(X, vlstrCveCargo, ",") - X, 0)) 'Identifica los cargos que fueron cambiados
'                                X = InStr(X, vlstrCveCargo, ",") + 1

                            For cont = 1 To grdPrecios.Rows - 1
'                                If Trim(grdPrecios.TextMatrix(cont, 1)) = vlstrCveCargocompleto And Trim(grdPrecios.TextMatrix(cont, 1)) <> "" Then
                                vlstrCveCargocompleto = Trim(grdPrecios.TextMatrix(cont, 1))
                                If Trim(grdPrecios.TextMatrix(cont, 1)) <> "" And grdPrecios.TextMatrix(cont, cintColModificado) = "*" Then
                                    vldblPrecio = grdPrecios.TextMatrix(cont, 8)
                                    If grdPrecios.TextMatrix(cont, 9) = "AR" Then
                                        Set rsArticulo = frsEjecuta_SP("'" & vlstrCveCargocompleto & "'", "sp_COArticulo")
                                        If rsArticulo.recordCount > 0 Then
                                            vlstrCveCargocompleto = rsArticulo!intIdArticulo
                                        End If
                                    End If
'                                    Exit For
'                                End If
'                                Next cont

                                vlstrsql = "select intnumpaquete from pvdetallepaquete where intcvecargo=" & Val(vlstrCveCargocompleto) & " and chrtipocargo = '" & grdPrecios.TextMatrix(cont, 9) & "'" ' Si el cargo pertenece a un paquete
                                Set rs = frsRegresaRs(vlstrsql)
                                If rs.recordCount > 0 Then
                                    Do While Not rs.EOF
                                                   vlstrsqlpaquete = "select bitpredeterminada from pvlistaprecio inner join " & _
                                                   "pvdetallelista on pvlistaprecio.intcvelista=pvdetallelista.intcvelista " & _
                                                   "where pvdetallelista.chrcvecargo  =" & rs!intnumpaquete & " and chrtipocargo='PA' "
                                                   Set rsPaquete = frsRegresaRs(vlstrsqlpaquete)
                                                   If rsPaquete.recordCount > 0 Then
                                                      Do While Not rsPaquete.EOF
                                                      If rsPaquete!bitPredeterminada = 1 Then ' si el paquete es de la lista predeterminada
                                                          vlstrtieneinccremento = "select nvl(bitincrementoautomatico,0) as incremento from pvpaquete where intnumpaquete=" & rs!intnumpaquete ' Si el paquete tiene el bit de incremento automatico
                                                          Set rstieneincremento = frsRegresaRs(vlstrtieneinccremento)
                                                          If rstieneincremento.recordCount > 0 Then
                                                             If rstieneincremento!incremento = 1 Then
                                                                vlstrSentencia = "update pvdetallepaquete set mnyprecio=" & vldblPrecio & " where intnumpaquete=" & rs!intnumpaquete & " And intcvecargo = " & vlstrCveCargocompleto & " and chrtipocargo = '" & grdPrecios.TextMatrix(cont, 9) & "'"
                                                                pEjecutaSentencia vlstrSentencia
                                                                vlstrSuma = "select sum(smicantidad* mnyprecio) as suma from pvdetallepaquete where intnumpaquete=" & rs!intnumpaquete
                                                                Set rssuma = frsRegresaRs(vlstrSuma)
                                                                If rssuma.recordCount > 0 Then
                                                                    vldblPreciosuma = rssuma!suma
                                                                    vgstrParametrosSP = Val(vldblPreciosuma) & "|" & Val(rs!intnumpaquete)
                                                                    frsEjecuta_SP vgstrParametrosSP, "SP_PVACTUALIZALISTAPRECIOS" 'Actualiza el precio en la lista de precios predeterminada
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                    rsPaquete.MoveNext
                                                    Loop
                                                End If
                                        rs.MoveNext
                                       
                                    Loop
                                End If
'                                If InStr(X, vlstrCveCargo, ",") = 0 Then
'                                    Exit Do
'                                End If
'                                Y = Y + 1
'                            Loop
                            End If
                            Next cont

                        End If

                        EntornoSIHO.ConeccionSIHO.CommitTrans
                    
                        pHabilita 0, 0, 1, 0, 0, 0, 0, 0
                    
                        freBarra.Visible = False
                    
                        ' Desabilita petición de catálogo de artículos
                        blnCatalogo = False
                    
                        'La lista de precios se actualizó satisfactoriamente.
                        MsgBox SIHOMsg(274), vbInformation, "Mensaje"
                        blnactivainicializa = 0
                        Set rsListasDepto = frsEjecuta_SP(CStr(cboDepartamento.ItemData(cboDepartamento.ListIndex)), "SP_PVSELLISTADEPTO")
                    
                        If rsListasDepto.recordCount = 1 And Not lblnConsulta Then
                        'Esta lista ha quedado como predeterminada por ser la primera, puede cambiarla cuando desee.
                        MsgBox SIHOMsg(347), vbOKOnly + vbInformation, "Mensaje"
                    End If
                    
                    pCargaOtrasListas -1
                    pNuevoRegistro
                End If
            End If
        End If
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdGrabarRegistro_Click"))
End Sub

Private Sub TxtClave_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError

    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
        KeyAscii = 7
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtClave_KeyPress"))
End Sub

Private Sub txtClave_LostFocus()
    On Error GoTo NotificaError

    If Trim(txtClave.Text) = "" Or Not fblnExiste() Then
        txtClave.Text = frsEjecuta_SP("", "SP_PVSELCONSECUTIVOLISTAPRECIO").Fields(0)
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtClave_LostFocus"))
End Sub

Private Function fblnExiste() As Boolean
    On Error GoTo NotificaError

    fblnExiste = False

    If rsListasDepto.recordCount <> 0 Then
        rsListasDepto.MoveFirst
        Do While Not rsListasDepto.EOF And Not fblnExiste
            If rsListasDepto!INTCVELISTA = Val(txtClave.Text) Then
                fblnExiste = True
            Else
                rsListasDepto.MoveNext
            End If
        Loop
    End If

    Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fblnExiste"))
End Function

Private Sub txtClaveArticulo_Change()
    txtIniciales.Text = ""
    llngMarcados = 0
    pHabilitaModificar
    If grdPrecios.Rows > 2 Or (grdPrecios.Rows = 2 And Trim(grdPrecios.TextMatrix(1, 1)) <> "") Then
        pconfiguragrid
    End If
End Sub

Private Sub txtClaveArticulo_GotFocus()
On Error GoTo NotificaError

    pSelTextBox txtClaveArticulo
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtClaveArticulo_GotFocus"))
    Unload Me
End Sub

Private Sub txtClaveArticulo_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then SendKeys vbTab

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtClaveArticulo_KeyDown"))
End Sub

Private Sub txtClaveArticulo_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError

    If KeyAscii = 13 Then
    Else
        If KeyAscii = 46 Then
            KeyAscii = 7
        Else
            pValidaNumero KeyAscii
            
'            Select Case KeyAscii
'                Case 8, 46
'                    txtIniciales.Text = ""
'                    If grdPrecios.Rows > 2 Or (grdPrecios.Rows = 2 And Trim(grdPrecios.TextMatrix(1, 1)) <> "") Then
'                        pConfiguraGrid
'                    End If
'                Case 48 To 57
'                    txtIniciales.Text = ""
'                    If grdPrecios.Rows > 2 Or (grdPrecios.Rows = 2 And Trim(grdPrecios.TextMatrix(1, 1)) <> "") Then
'                        pConfiguraGrid
'                    End If
'            End Select
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtClaveArticulo_KeyPress"))
    Unload Me
End Sub

Private Sub txtDescripcion_Change()
    On Error GoTo NotificaError

    cmdCapturaPrecios.Enabled = Trim(txtDescripcion.Text) <> ""
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtDescripcion_Change"))
End Sub

Private Sub txtDescripcion_GotFocus()
    On Error GoTo NotificaError

    pHabilita 0, 0, 0, 0, 0, 1, 0, 0
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtDescripcion_GotFocus"))
End Sub

Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then
        'If cmdCapturaPrecios.Enabled Then
        '    cmdCapturaPrecios.SetFocus
        'End If
        cboTabulador.SetFocus
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtDescripcion_KeyDown"))
End Sub

Private Sub TxtDescripcion_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError

    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtDescripcion_KeyPress"))
End Sub

Private Sub txtIniciales_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then SendKeys vbTab

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtIniciales_KeyDown"))
End Sub

Private Sub txtIniciales_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError

    KeyAscii = Asc(UCase(Chr(KeyAscii)))

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtIniciales_KeyPress"))
End Sub

Private Sub txtPorciento_Change()
    On Error GoTo NotificaError

    cmdMas.Enabled = Trim(txtPorciento.Text) <> ""
    cmdMenos.Enabled = Trim(txtPorciento.Text) <> ""
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtPorciento_Change"))
End Sub

Private Sub txtPorciento_GotFocus()
    On Error GoTo NotificaError

    pSelTextBox txtPorciento
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtPorciento_GotFocus"))
End Sub
Private Sub txtPorciento_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then
        cmdMas.SetFocus
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtPorciento_KeyDown"))
End Sub

Private Sub txtPorciento_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError

    If Not fblnFormatoCantidad(txtPorciento, KeyAscii, 2) Then
        KeyAscii = 7
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtPorciento_KeyPress"))
End Sub

Private Sub grdHBusqueda_DblClick()
    On Error GoTo NotificaError

    If Trim(grdHBusqueda.TextMatrix(1, 1)) <> "" Then
        txtClave.Text = grdHBusqueda.TextMatrix(grdHBusqueda.Row, 1)
        If fblnExiste() Then
            pModificaRegistro
            pCargaOtrasListas Val(txtClave.Text)
            pHabilita 1, 1, 1, 1, 1, 0, 1, 1
            'cmdBuscar.SetFocus
            cmdCapturaPrecios.SetFocus
            chkImprimeCargosPrecioCero.Value = 1
            Chkbitcheckup.Value = 0
            sstListas.Tab = 0
        End If
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":grdHBusqueda_DblClick"))
End Sub
Private Sub grdPrecios_Click()
    On Error GoTo NotificaError

    ' Despliega el textbox si el usuario presiona el click en una columna
    ' que no sea cabecera y edita el texto si es la columna que se puede editar.
    If Trim(grdPrecios.TextMatrix(grdPrecios.Row, cintColClave)) <> "" Then
        If grdPrecios.Col = cintColPrecio Then
            If Val(grdPrecios.TextMatrix(grdPrecios.Row, cintColPrecioEspecifico)) <> 1 Then
                Call pEditarColumna(32, txtPrecio, grdPrecios)
            Else
            '¡No se pueden modificar precios relacionados con cargos de paquetes!
            MsgBox SIHOMsg(1592) & ":" & Chr(13) & Trim(grdPrecios.TextMatrix(grdPrecios.Row, cintColDescripcion)), vbOKOnly + vbInformation, "Mensaje"
            Exit Sub
            End If
        End If
        If grdPrecios.Col = cintColUtilidad Then
            Call pEditarColumna(32, txtPrecio, grdPrecios)
        End If
        If grdPrecios.Col = cintColTipoIncremento Then
            pPonerUpDown grdPrecios
        End If
        
        If grdPrecios.Col = cintColCosto Then
            If lblnPermisoCosto And (grdPrecios.TextMatrix(grdPrecios.Row, cintColTipoIncremento) = "PRECIO MAXIMO AL PUBLICO" Or grdPrecios.TextMatrix(grdPrecios.Row, cintColTipoIncremento) = "NA") Then
                Call pEditarColumna(32, txtPrecio, grdPrecios)
            End If
        End If
        
        If grdPrecios.Col = cintColMoneda And grdPrecios.TextMatrix(grdPrecios.Row, cintColTipo) = "PA" Then
            pMostrarlstPesos grdPrecios
        End If
        
        If grdPrecios.MouseCol = 0 Then
            grdPrecios.TextMatrix(grdPrecios.Row, 0) = IIf(grdPrecios.TextMatrix(grdPrecios.Row, 0) = "*", "", "*")
            If grdPrecios.TextMatrix(grdPrecios.Row, 0) = "*" Then
                llngMarcados = llngMarcados + 1
            Else
                llngMarcados = llngMarcados - 1
            End If
            pHabilitaModificar
        End If
        
        If grdPrecios.TextMatrix(grdPrecios.Row, cintColTipo) = "AR" Then
            If grdPrecios.Col = cintColTabulador Then
                grdPrecios.TextMatrix(grdPrecios.Row, cintColTabulador) = IIf(grdPrecios.TextMatrix(grdPrecios.Row, cintColTabulador) = "*", "", "*")
                grdPrecios.TextMatrix(grdPrecios.Row, cintColModificado) = "*"
                'pCalcularPrecio grdPrecios.Row
            End If
        End If
        
        If grdPrecios.Col = cintColIncremetoAutomatico Then
            'Trim(grdPrecios.TextMatrix(grdPrecios.Row, cintColIncremetoAutomatico)) = IIf(Trim(grdPrecios.TextMatrix(grdPrecios.Row, cintColIncremetoAutomatico)) = "*", "", "*")
            grdPrecios.TextMatrix(grdPrecios.Row, cintColIncremetoAutomatico) = IIf(grdPrecios.TextMatrix(grdPrecios.Row, cintColIncremetoAutomatico) = "*", "", "*")
            grdPrecios.TextMatrix(grdPrecios.Row, cintColModificado) = "*"
            'pCalcularPrecio grdPrecios.Row
        End If
    End If
    pMuestraNomComercialCompleto
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":grdPrecios_Click"))
End Sub

Private Sub grdPrecios_GotFocus()
    On Error GoTo NotificaError

    If vgblnNoEditar Then Exit Sub
    
    'Copia el valor del textbox al grid y lo esconde
    If grdPrecios.Col = cintColPrecio Then
            Call pSetCellValueCol(grdPrecios, txtPrecio)
    End If
    If grdPrecios.Col = cintColUtilidad Then
        Call pSetCellValueCol(grdPrecios, txtPrecio)
    End If
    If grdPrecios.Col = cintColCosto Then
        Call pSetCellValueCol(grdPrecios, txtPrecio)
    End If
    If grdPrecios.Col = cintColTipoIncremento Then
        UpDown1.Visible = False
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":grdPrecios_GotFocus"))
End Sub

Private Sub grdPrecios_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    If grdPrecios.Col = cintColPrecio Or grdPrecios.Col = cintColUtilidad Or (lblnPermisoCosto And grdPrecios.Col = cintColCosto And (grdPrecios.TextMatrix(grdPrecios.Row, cintColTipoIncremento) = "PRECIO MÁXIMO AL PÚBLICO" Or grdPrecios.TextMatrix(grdPrecios.Row, cintColTipoIncremento) = "NA")) Then
        If KeyCode = vbKeyF2 And Trim(grdPrecios.TextMatrix(grdPrecios.Row, cintColClave)) <> "" Then 'para que se edite el contenido de la celda como en excel
         If Val(grdPrecios.TextMatrix(grdPrecios.Row, cintColPrecioEspecifico)) <> 1 Then
            Call pEditarColumna(13, txtPrecio, grdPrecios)
         Else
            '¡No se pueden modificar precios relacionados con cargos de paquetes!
            MsgBox SIHOMsg(1592) & ":" & Chr(13) & Trim(grdPrecios.TextMatrix(grdPrecios.Row, cintColDescripcion)), vbOKOnly + vbInformation, "Mensaje"
            Exit Sub
         End If
         
        End If
    ElseIf grdPrecios.Col = cintColTipoIncremento Then
    ElseIf grdPrecios.Col = cintColMoneda Then
    Else
        If KeyCode = vbKeyReturn Then
            If grdPrecios.Row - 1 < grdPrecios.Rows Then
                If grdPrecios.Row = grdPrecios.Rows - 1 Then
                    grdPrecios.Row = 1
                Else
                    grdPrecios.Row = grdPrecios.Row + 1
                End If
            End If
        End If
    End If
    
    If (KeyCode = vbKeyDown Or KeyCode = vbKeyUp) Then
        If UpDown1.Visible Then
            
        Else
            pMuestraNomComercialCompleto
            vgstrAcumTextoBusqueda = ""
        End If
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":grdPrecios_KeyDown"))
End Sub

Private Sub grdPrecios_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError

    If Trim(grdPrecios.TextMatrix(grdPrecios.Row, cintColClave)) <> "" Then
        If grdPrecios.Col = cintColPrecio Or grdPrecios.Col = cintColUtilidad Or (lblnPermisoCosto And grdPrecios.Col = cintColCosto And (grdPrecios.TextMatrix(grdPrecios.Row, cintColTipoIncremento) = "PRECIO MÁXIMO AL PÚBLICO" Or grdPrecios.TextMatrix(grdPrecios.Row, cintColTipoIncremento) = "NA")) Then 'Columna que puede ser editada
            If Val(grdPrecios.TextMatrix(grdPrecios.Row, cintColPrecioEspecifico)) <> 1 Then
                Call pEditarColumna(KeyAscii, txtPrecio, grdPrecios)
            Else
                '¡No se pueden modificar precios relacionados con cargos de paquetes!
                MsgBox SIHOMsg(1592) & ":" & Chr(13) & Trim(grdPrecios.TextMatrix(grdPrecios.Row, cintColDescripcion)), vbOKOnly + vbInformation, "Mensaje"
                Exit Sub
            End If
        End If
        If grdPrecios.Col = cintColIncremetoAutomatico Or grdPrecios.Col = cintColTabulador Then
            If KeyAscii = 32 Then
                grdPrecios_Click
            End If
        End If
        If grdPrecios.Col = cintColTipoIncremento Then
            If KeyAscii = 13 Then
                pPonerUpDown grdPrecios
            End If
        End If
        If grdPrecios.Col = cintColMoneda And grdPrecios.TextMatrix(grdPrecios.Row, cintColTipo) = "PA" Then
            If KeyAscii = 13 Then
                pMostrarlstPesos grdPrecios
            End If
        End If
    End If
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":grdPrecios_KeyPress"))
End Sub

Private Sub grdPrecios_LeaveCell()
    On Error GoTo NotificaError

    If vgblnNoEditar Then Exit Sub
    Call grdPrecios_GotFocus
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":grdPrecios_LeaveCell"))
End Sub

Private Sub grdPrecios_Scroll()
    On Error GoTo NotificaError

    Call grdPrecios_GotFocus
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":grdPrecios_Scroll"))
End Sub

Private Sub pSetCellValueCol(grid As MSHFlexGrid, txtEdit As TextBox)
    On Error GoTo NotificaError

    ' NOTA:
    ' Este código debe ser  llamado cada vez que el grid pierde el foco y su contenido puede cambiar.
    ' De otra manera, el nuevo valor de la celda se perdería.
    
    If grid.Col = cintColPrecio Then
        If txtPrecio.Visible Then
            If txtPrecio.Text <> "" Then
                If IsNumeric(txtPrecio.Text) Then
                    grid.Text = Format(txtPrecio.Text, "$###,###,###,##0.00####")
                    grid.TextMatrix(grid.Row, cintColModificado) = "*"
                    'grid.TextMatrix(grid.Row, cintColIncremetoAutomatico) = ""
                End If
                vgstrEstadoManto = Mid(vgstrEstadoManto, 1, 2)
            End If
            txtPrecio.Visible = False
        End If
    End If
    
    If grid.Col = cintColCosto Then
        If txtPrecio.Visible Then
            If txtPrecio.Text <> "" Then
                If IsNumeric(txtPrecio.Text) Then
                    grid.Text = Format(txtPrecio.Text, "$###,###,###,##0.0000##")
                    grid.TextMatrix(grid.Row, cintColModificado) = "*"
                    'grid.TextMatrix(grid.Row, cintColIncremetoAutomatico) = ""
                End If
                vgstrEstadoManto = Mid(vgstrEstadoManto, 1, 2)
            End If
            txtPrecio.Visible = False
        End If
    End If
    
    If grid.Col = cintColUtilidad Then
        If txtPrecio.Visible Then
            If txtPrecio.Text <> "" Then
                txtPrecio.Text = Replace(txtPrecio.Text, "%", "")
                If IsNumeric(txtPrecio.Text) Then
                    grid.Text = Format(txtPrecio.Text, "0.0000") & "%"
                    grid.TextMatrix(grid.Row, cintColModificado) = "*"
                    'pCalcularPrecio grid.Row
                End If
                vgstrEstadoManto = Mid(vgstrEstadoManto, 1, 2)
            End If
            txtPrecio.Visible = False
        End If
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pSetCellValueCol"))
End Sub

Public Sub pEditarColumna(KeyAscii As Integer, txtEdit As TextBox, grid As MSHFlexGrid)
    On Error GoTo NotificaError

    Dim vlintTexto As Integer

    With txtEdit
       .Text = Replace(grid, "%", "") 'Inicialización del Textbox
        Select Case KeyAscii
            Case 0 To 32
                'Edita el texto de la celda en la que está posicionado
                .SelStart = 0
                .SelLength = 1000
            Case 8, 48 To 57
                ' Reemplaza el texto actual solo si se teclean números
                vlintTexto = Chr(KeyAscii)
                .Text = vlintTexto
                .SelStart = 1
            Case 46
                ' Reemplaza el texto actual solo si se teclean números
                .Text = "."
                .SelStart = 1
        End Select
    End With
            
    ' Muestra el textbox en el lugar indicado
    With grid
        If .CellWidth < 0 Then Exit Sub
        txtEdit.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth - 48, .CellHeight - 8
    End With
    
    vgstrEstadoManto = vgstrEstadoManto & "E"
    txtEdit.Visible = True
    txtEdit.SetFocus
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pEditarColumna"))
End Sub

Private Sub pPonerUpDown(grid As MSHFlexGrid)
    On Error GoTo NotificaError

    Dim intIndex As Integer
    If grid.TextMatrix(grid.Row, cintColTipo) = "AR" Then
        UpDown1.ListIndex = -1
        For intIndex = 0 To UpDown1.ListCount - 1
            If UpDown1.List(intIndex) = grid.Text Then
                UpDown1.ListIndex = intIndex
                Exit For
            End If
        Next
        With grid
            UpDown1.Move .Left + .CellLeft, .Top + .CellTop, UpDown1.Width, UpDown1.Height
        End With
        UpDown1.Visible = True
        UpDown1.SetFocus
    End If
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pPonerUpDown"))
End Sub

Private Sub pMostrarlstPesos(grd As MSHFlexGrid)
    With grd
        lstPesos.Move .Left + .CellLeft, .Top + .CellTop, lstPesos.Width, lstPesos.Height
        If .TextMatrix(.Row, cintColMoneda) = "PESOS" Then
            lstPesos.ListIndex = 1
        Else
            lstPesos.ListIndex = 0
    End If
    End With
    
    lstPesos.Visible = True
    lstPesos.SetFocus
End Sub


Private Sub txtPorciento_LostFocus()
    On Error GoTo NotificaError

    If IsNumeric(txtPorciento.Text) Then
        txtPorciento.Text = Format(txtPorciento.Text, "0.00")
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtPorciento_LostFocus"))
End Sub

Private Sub txtPrecio_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    'Para verificar que tecla fue presionada en el textbox
    With grdPrecios
        Select Case KeyCode
            Case 27   'ESC
                 txtPrecio.Visible = False
                .SetFocus
            Case 38   'Flecha para arriba
                .SetFocus
                DoEvents
                If .Row > .FixedRows Then
                    vgblnNoEditar = True
                    .Row = .Row - 1
                    vgblnNoEditar = False
                End If
                vgblnEditaPrecio = False
            Case 40, 13
                .SetFocus
                DoEvents
                If .Row < .Rows - 1 Then
                    vgblnNoEditar = True
                    .Row = .Row + 1
                    vgblnNoEditar = False
                End If
                vgblnEditaPrecio = False
        End Select
    End With

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtPrecio_KeyDown"))
End Sub

Private Sub txtPrecio_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    Dim bytNumDecimales As Byte ' Solo permite números
    
    bytNumDecimales = IIf(grdPrecios.Col = cintColUtilidad, 4, 6)
    
    If Not fblnFormatoCantidad(txtPrecio, KeyAscii, bytNumDecimales) Then
        KeyAscii = 7
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtPrecio_KeyPress"))
End Sub

Sub pMuestraNomComercialCompleto()
    On Error GoTo NotificaError

    lblNombreCompleto.Caption = grdPrecios.TextMatrix(grdPrecios.Row, cintColDescripcion)
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pMuestraNomComercialCompleto"))
End Sub

Private Sub pHabilita(intTop As Integer, intBack As Integer, intlocate As Integer, intNext As Integer, intEnd As Integer, intSave As Integer, intDelete As Integer, intPrint As Integer)
    On Error GoTo NotificaError
    
    cmdPrimerRegistro.Enabled = intTop = 1
    cmdAnteriorRegistro.Enabled = intBack = 1
    cmdBuscar.Enabled = intlocate = 1
    cmdSiguienteRegistro.Enabled = intNext = 1
    cmdUltimoRegistro.Enabled = intEnd = 1
    cmdGrabarRegistro.Enabled = intSave = 1
    cmdDelete.Enabled = intDelete = 1
    cmdImprimir.Enabled = intPrint = 1
    chkImprimeCargosPrecioCero.Enabled = intPrint = 1
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pHabilita"))
End Sub

Private Sub txtPrecio_LostFocus()
    On Error GoTo NotificaError
    grdPrecios_GotFocus
    
    If chkPredeterminada.Value = 1 Then
        If grdPrecios.Col = 8 Then
            vlstrCveCargo = vlstrCveCargo & Trim(grdPrecios.TextMatrix(grdPrecios.Row, 1)) & ","
            grdPrecios.TextMatrix(grdPrecios.Row, cintColModificado) = "*"
        End If
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtPrecio_LostFocus"))
End Sub

Private Sub UpDown1_KeyPress(KeyAscii As Integer)
66666    On Error GoTo NotificaError

    If KeyAscii = 13 Then
        UpDown1_MouseUp 0, 0, 0, 0
        If grdPrecios.Row < grdPrecios.Rows - 1 Then
            grdPrecios.Row = grdPrecios.Row + 1
        End If
    End If
    If KeyAscii = 27 Then
        grdPrecios.SetFocus
        UpDown1.Visible = False
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":UpDown1_KeyPress"))
End Sub

Private Sub UpDown1_LostFocus()
    On Error GoTo NotificaError

    UpDown1.Visible = False
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":UpDown1_LostFocus"))
End Sub

Private Sub UpDown1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo NotificaError

    grdPrecios.Text = UpDown1.Text
    grdPrecios.TextMatrix(grdPrecios.Row, cintColModificado) = "*"
    Select Case grdPrecios.Text
        Case cstrUltimaCompra
            grdPrecios.TextMatrix(grdPrecios.Row, cintColCosto) = Format(grdPrecios.TextMatrix(grdPrecios.Row, cintColCostoUltimaEntrada), "$###,###,###,##0.0000##")
        Case cstrCompraMasAlta
            grdPrecios.TextMatrix(grdPrecios.Row, cintColCosto) = Format(grdPrecios.TextMatrix(grdPrecios.Row, cintColCostoMasAlto), "$###,###,###,##0.0000##")
        Case cstrPrecioMaximoPublico
            grdPrecios.TextMatrix(grdPrecios.Row, cintColCosto) = Format(grdPrecios.TextMatrix(grdPrecios.Row, cintColPrecioMaximopublico), "$###,###,###,##0.0000##")
    End Select
    grdPrecios.SetFocus
    UpDown1.Visible = False
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":UpDown1_MouseUp"))
End Sub

Private Sub UpDown1_Validate(Cancel As Boolean)
    On Error GoTo NotificaError

    UpDown1_MouseUp 0, 0, 0, 0
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":UpDown1_Validate"))
End Sub

Private Sub pCalcularPrecio(lngRow As Long)
    On Error GoTo NotificaError
    
    Dim dblPrecio As Double
    Dim dblAumentoTabulador As Double
    Dim rs As ADODB.Recordset
    Dim strParametros As String
    Dim dblCosto As Double
    Dim dblUtilidad As Double
    Dim strSql As String
    If Trim(grdPrecios.TextMatrix(lngRow, cintColIncremetoAutomatico)) = "*" Then
        dblCosto = CDbl(grdPrecios.TextMatrix(lngRow, cintColCosto))
        dblUtilidad = CDbl(Replace(grdPrecios.TextMatrix(lngRow, cintColUtilidad), "%", ""))
        dblAumentoTabulador = 0
        If Trim(grdPrecios.TextMatrix(lngRow, cintColTabulador)) = "*" Then
            strSql = "select sp_IVSelTabulador(" & dblCosto & ", '" & grdPrecios.TextMatrix(lngRow, cintColClave) & "', " & fintTabuladorListaPrecio(txtClave.Text) & ") aumento from dual"
            Set rs = frsRegresaRs(strSql)
            If Not rs.EOF Then
                dblAumentoTabulador = rs!aumento
            End If
            rs.Close
        End If
        dblCosto = dblCosto * (1 + (dblUtilidad / 100))
        dblCosto = dblCosto * (1 + (dblAumentoTabulador / 100))
        grdPrecios.TextMatrix(lngRow, cintColPrecio) = Format(dblCosto, "$###,###,###,##0.00####")
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCalcularPrecio"))
End Sub

Private Sub pCargaTabuladores()
    Dim rs As ADODB.Recordset
    cboTabulador.Clear
    cboTabulador.AddItem "<PREDETERMINADO>"
    cboTabulador.ItemData(cboTabulador.NewIndex) = -1
    Set rs = frsRegresaRs("select * from IVTabulador order by vchDescripcion", adLockReadOnly, adOpenForwardOnly)
    Do While Not rs.EOF
        cboTabulador.AddItem rs!VCHDESCRIPCION
        cboTabulador.ItemData(cboTabulador.NewIndex) = rs!intCveTabulador
        rs.MoveNext
    Loop
    rs.Close
    cboTabulador.ListIndex = 0
End Sub

