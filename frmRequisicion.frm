VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.1#0"; "MyCommandButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{FF14BD24-9F8A-41E3-B5B8-7F0D45EE9F16}#15.0#0"; "hsflatcontrols.ocx"
Begin VB.Form frmRequisicion 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Requisiciones de artículos"
   ClientHeight    =   10020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13830
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   10020
   ScaleWidth      =   13830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab sstObj 
      Height          =   10320
      Left            =   0
      TabIndex        =   39
      Top             =   0
      Width           =   14505
      _ExtentX        =   25585
      _ExtentY        =   18203
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   2
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      TabCaption(0)   =   "Ingreso/Revision"
      TabPicture(0)   =   "frmRequisicion.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraArticulos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraCabecera"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "frmBotonera"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraBarra"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fraFiltros"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "MyButton1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Busqueda"
      TabPicture(1)   =   "frmRequisicion.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(1)=   "cboERBusqueda"
      Tab(1).Control(2)=   "cboTRBusqueda"
      Tab(1).Control(3)=   "grdRequisiciones"
      Tab(1).Control(4)=   "Label4"
      Tab(1).Control(5)=   "Label5"
      Tab(1).ControlCount=   6
      Begin MyCommandButton.MyButton MyButton1 
         Height          =   375
         Left            =   11715
         TabIndex        =   92
         ToolTipText     =   "Observaciones de la requisición"
         Top             =   350
         Width           =   1935
         _ExtentX        =   3413
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
         BackColorDown   =   -2147483643
         BackColorOver   =   -2147483633
         BackColorFocus  =   -2147483633
         BackColorDisabled=   -2147483633
         BorderColor     =   -2147483627
         TransparentColor=   13160660
         Caption         =   "Ver observaciones"
         DepthEvent      =   1
         ShowFocus       =   -1  'True
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
         Height          =   3330
         Left            =   120
         TabIndex        =   51
         Top             =   2040
         Width           =   13635
         Begin HSFlatControls.MyCombo cboLocalizacion 
            Height          =   375
            Left            =   2600
            TabIndex        =   14
            ToolTipText     =   "Selección de la localización"
            Top             =   705
            Width           =   4335
            _ExtentX        =   7646
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
            Left            =   120
            TabIndex        =   10
            Top             =   300
            Width           =   930
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
            Left            =   1220
            TabIndex        =   11
            Top             =   300
            Width           =   1125
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
            Left            =   2590
            TabIndex        =   12
            Top             =   300
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
            Left            =   4680
            TabIndex        =   13
            Top             =   300
            Width           =   1080
         End
         Begin VB.TextBox txtCodigoBarras 
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
            Left            =   9315
            MaxLength       =   50
            TabIndex        =   18
            ToolTipText     =   "Cpodigo de barras del artículo"
            Top             =   1110
            Width           =   4215
         End
         Begin VB.TextBox txtNombreGenerico 
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
            Left            =   2600
            TabIndex        =   20
            ToolTipText     =   "Nombre generico del artículo"
            Top             =   1920
            Width           =   4335
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
            Left            =   9315
            TabIndex        =   19
            ToolTipText     =   "Clave del artículo"
            Top             =   1520
            Width           =   4215
         End
         Begin VB.TextBox txtCantidadSolicitada 
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
            Left            =   5960
            MaxLength       =   6
            TabIndex        =   22
            ToolTipText     =   "Cantidad de solicitada del artículo"
            Top             =   2320
            Width           =   975
         End
         Begin VB.TextBox txtExistencia 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
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
            Left            =   2600
            Locked          =   -1  'True
            TabIndex        =   21
            ToolTipText     =   "Existencia del artículo"
            Top             =   2320
            Width           =   975
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   320
            Left            =   7050
            TabIndex        =   61
            Top             =   2355
            Width           =   2085
            Begin VB.OptionButton optAlterna 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Alterna"
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
               Left            =   0
               TabIndex        =   23
               ToolTipText     =   "Manejar unidad venta"
               Top             =   50
               Width           =   990
            End
            Begin VB.OptionButton optMinima 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Mínima"
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
               Left            =   1000
               TabIndex        =   24
               ToolTipText     =   "Manejar unidad mínima"
               Top             =   50
               Width           =   1125
            End
         End
         Begin VB.CheckBox chkPedirMaximo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Pedir basándose en el punto de re orden"
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
            Left            =   9315
            TabIndex        =   78
            Top             =   2382
            Width           =   4260
         End
         Begin MyCommandButton.MyButton cmdAgregarNuevo 
            Height          =   495
            Left            =   13065
            TabIndex        =   81
            ToolTipText     =   "Agregar artículo nuevo."
            Top             =   2745
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   873
            BackColor       =   -2147483633
            Enabled         =   0   'False
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
            Picture         =   "frmRequisicion.frx":0038
            BackColorDown   =   -2147483643
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmRequisicion.frx":074C
            PictureAlignment=   5
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdAgregar 
            Height          =   495
            Left            =   12570
            TabIndex        =   80
            ToolTipText     =   "Agregar artículo"
            Top             =   2745
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   873
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
            Picture         =   "frmRequisicion.frx":0E60
            BackColorDown   =   -2147483643
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmRequisicion.frx":1574
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin HSFlatControls.MyCombo cboSubfamilia 
            Height          =   375
            Left            =   9315
            TabIndex        =   16
            ToolTipText     =   "Selección de la subfamilia"
            Top             =   705
            Width           =   4215
            _ExtentX        =   7435
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
            Left            =   2600
            TabIndex        =   17
            ToolTipText     =   "Selección del artículo"
            Top             =   1515
            Width           =   4335
            _ExtentX        =   7646
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
         Begin HSFlatControls.MyCombo cboFamilia 
            Height          =   375
            Left            =   2600
            TabIndex        =   15
            ToolTipText     =   "Selección de la familia"
            Top             =   1110
            Width           =   4335
            _ExtentX        =   7646
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
         Begin HSFlatControls.MyCombo cboCajaMaterial 
            Height          =   375
            Left            =   9315
            TabIndex        =   25
            ToolTipText     =   "Caja de material"
            Top             =   1920
            Width           =   4215
            _ExtentX        =   7435
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
         Begin VB.Label lblCajaMaterial 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Caja de material"
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
            Left            =   7050
            TabIndex        =   91
            Top             =   1980
            Width           =   1665
         End
         Begin VB.Label lblLocalizacion 
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
            Height          =   255
            Left            =   150
            TabIndex        =   52
            Top             =   760
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
            Height          =   255
            Left            =   150
            TabIndex        =   53
            Top             =   1170
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
            Height          =   255
            Left            =   7050
            TabIndex        =   54
            Top             =   765
            Width           =   1005
         End
         Begin VB.Label lblComercial 
            AutoSize        =   -1  'True
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
            Height          =   255
            Left            =   150
            TabIndex        =   55
            Top             =   1580
            Width           =   1830
         End
         Begin VB.Label lblCodigoBarras 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Código barras"
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
            Left            =   7050
            TabIndex        =   56
            Top             =   1170
            Width           =   1410
         End
         Begin VB.Label lblGenerico 
            AutoSize        =   -1  'True
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
            Left            =   150
            TabIndex        =   57
            Top             =   1980
            Width           =   1710
         End
         Begin VB.Label lblClave 
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
            Left            =   7050
            TabIndex        =   58
            Top             =   1580
            Width           =   585
         End
         Begin VB.Label lblExistencia 
            BackColor       =   &H80000005&
            Caption         =   "Existencia en el almacén"
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
            Left            =   150
            TabIndex        =   59
            Top             =   2380
            Width           =   3765
         End
         Begin VB.Label lblCantidad 
            BackColor       =   &H80000005&
            Caption         =   "Cantidad solicitada"
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
            Left            =   3900
            TabIndex        =   60
            Top             =   2385
            Width           =   1950
         End
      End
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
         Height          =   930
         Left            =   2715
         TabIndex        =   82
         Top             =   5640
         Visible         =   0   'False
         Width           =   8445
         Begin MSComctlLib.ProgressBar pgbBarra 
            Height          =   150
            Left            =   45
            TabIndex        =   83
            Top             =   600
            Width           =   8340
            _ExtentX        =   14711
            _ExtentY        =   265
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   60
            TabIndex        =   84
            Top             =   240
            Width           =   3540
         End
      End
      Begin VB.Frame frmBotonera 
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
         Left            =   3687
         TabIndex        =   38
         Top             =   9105
         Width           =   6500
         Begin MyCommandButton.MyButton cmdCierraReq 
            Height          =   600
            Left            =   4260
            TabIndex        =   88
            ToolTipText     =   "Cierra una requisición de reubicación, salida a departamento o compra - pedido"
            Top             =   200
            Width           =   2175
            _ExtentX        =   3836
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
            TransparentColor=   15790320
            Caption         =   "Cerrar requisición"
            DepthEvent      =   1
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdImprimir 
            Height          =   600
            Left            =   3660
            TabIndex        =   35
            ToolTipText     =   "Imprimir requisición"
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
            Picture         =   "frmRequisicion.frx":1C88
            BackColorDown   =   -2147483643
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmRequisicion.frx":260C
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdGrabar 
            Height          =   600
            Left            =   3060
            TabIndex        =   34
            ToolTipText     =   "Guardar el Registro"
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
            Picture         =   "frmRequisicion.frx":2F8E
            BackColorDown   =   -2147483643
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmRequisicion.frx":3912
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdUltimo 
            Height          =   600
            Left            =   2460
            TabIndex        =   33
            ToolTipText     =   "Último registro"
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
            Picture         =   "frmRequisicion.frx":4296
            BackColorDown   =   -2147483643
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmRequisicion.frx":4C18
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdSiguiente 
            Height          =   600
            Left            =   1860
            TabIndex        =   32
            ToolTipText     =   "Siguiente registro"
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
            Picture         =   "frmRequisicion.frx":559A
            BackColorDown   =   -2147483643
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmRequisicion.frx":5F1C
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdBuscar 
            Height          =   600
            Left            =   1260
            TabIndex        =   31
            ToolTipText     =   "Búsqueda"
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
            Picture         =   "frmRequisicion.frx":689E
            BackColorDown   =   -2147483643
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmRequisicion.frx":7222
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdAnterior 
            Height          =   600
            Left            =   660
            TabIndex        =   30
            ToolTipText     =   "Anterior registro"
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
            Picture         =   "frmRequisicion.frx":7BA6
            BackColorDown   =   -2147483643
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            CaptionPosition =   4
            DepthEvent      =   1
            ForeColorDisabled=   -2147483629
            ForeColorOver   =   13003064
            ForeColorFocus  =   13003064
            ForeColorDown   =   13003064
            PictureDisabled =   "frmRequisicion.frx":8528
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdPrimero 
            Height          =   600
            Left            =   60
            TabIndex        =   29
            ToolTipText     =   "Primer registro"
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
            Picture         =   "frmRequisicion.frx":8EAA
            BackColorDown   =   -2147483643
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmRequisicion.frx":982C
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000015&
            X1              =   0
            X2              =   6490
            Y1              =   30
            Y2              =   30
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Rango de fechas"
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
         Height          =   795
         Left            =   -74880
         TabIndex        =   63
         Top             =   10
         Width           =   3015
         Begin MSMask.MaskEdBox mskFecIni 
            Height          =   375
            Left            =   120
            TabIndex        =   64
            ToolTipText     =   "Fecha inicial"
            Top             =   300
            Width           =   1335
            _ExtentX        =   2355
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
         Begin MSMask.MaskEdBox mskFecFin 
            Height          =   375
            Left            =   1515
            TabIndex        =   65
            ToolTipText     =   "Fecha final"
            Top             =   300
            Width           =   1335
            _ExtentX        =   2355
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
      End
      Begin VB.Frame fraCabecera 
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
         Height          =   2010
         Left            =   120
         TabIndex        =   36
         Top             =   50
         Width           =   13635
         Begin VB.TextBox txtEstatus 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
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
            Left            =   9315
            Locked          =   -1  'True
            TabIndex        =   4
            ToolTipText     =   "El estatus de la requisición"
            Top             =   300
            Width           =   4215
         End
         Begin VB.CheckBox chkUrgente 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Urgente"
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
            Left            =   5960
            TabIndex        =   1
            ToolTipText     =   "La requisición es urgente o no"
            Top             =   360
            Width           =   1050
         End
         Begin VB.TextBox txtFecha 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
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
            Left            =   4550
            Locked          =   -1  'True
            TabIndex        =   3
            ToolTipText     =   "Fecha de la requisición"
            Top             =   300
            Width           =   1305
         End
         Begin VB.TextBox txtNumero 
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
            Left            =   2600
            TabIndex        =   0
            ToolTipText     =   "Número de requisición"
            Top             =   300
            Width           =   1200
         End
         Begin VB.TextBox txtEmpleadoSolicito 
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
            Left            =   2600
            Locked          =   -1  'True
            TabIndex        =   71
            Top             =   1110
            Width           =   4335
         End
         Begin VB.CheckBox chkIncluir 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Mostrar datos última compra"
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
            Left            =   9315
            TabIndex        =   7
            Top             =   1170
            Width           =   3130
         End
         Begin VB.CheckBox chkCompradirecta 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Compra directa"
            Enabled         =   0   'False
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
            Left            =   7050
            TabIndex        =   6
            ToolTipText     =   "Requisición marcada como compra directa"
            Top             =   1170
            Width           =   1920
         End
         Begin HSFlatControls.MyCombo cboAlmacenSurte 
            Height          =   375
            Left            =   9315
            TabIndex        =   8
            ToolTipText     =   "Almacén que surtirá o realizará el pedido a compras"
            Top             =   1515
            Width           =   4215
            _ExtentX        =   7435
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
         Begin HSFlatControls.MyCombo cboProveedor 
            Height          =   375
            Left            =   2600
            TabIndex        =   9
            ToolTipText     =   "Proveedor al que se realiza la compra directa"
            Top             =   1515
            Width           =   4335
            _ExtentX        =   7646
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
         Begin HSFlatControls.MyCombo cboTipoRequisicion 
            Height          =   375
            Left            =   9315
            TabIndex        =   5
            ToolTipText     =   "Tipo requisición (Salida a departamento, Reubicación, Compra)"
            Top             =   705
            Width           =   4215
            _ExtentX        =   7435
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
         Begin HSFlatControls.MyCombo cboDepartamentoSolicita 
            Height          =   375
            Left            =   2600
            TabIndex        =   2
            ToolTipText     =   "Selección del departamento que solicita la requisición"
            Top             =   705
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   661
            Style           =   1
            Enabled         =   0   'False
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
         Begin VB.Label lblAlmacenProv 
            BackColor       =   &H80000005&
            Caption         =   "Almacén surtirá"
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
            Left            =   7050
            TabIndex        =   85
            Top             =   1575
            Width           =   1680
         End
         Begin VB.Label lblTipo 
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
            Left            =   7050
            TabIndex        =   41
            Top             =   765
            Width           =   1470
         End
         Begin VB.Label lblEstatus 
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
            Height          =   255
            Left            =   7050
            TabIndex        =   42
            Top             =   360
            Width           =   660
         End
         Begin VB.Label lblEmpleado 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Empleado solicitó"
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
            Left            =   150
            TabIndex        =   43
            Top             =   1170
            Width           =   1740
         End
         Begin VB.Label lblDepartamento 
            BackColor       =   &H80000005&
            Caption         =   "Departamento solicita"
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
            Left            =   150
            TabIndex        =   44
            Top             =   765
            Width           =   2250
         End
         Begin VB.Label lblFecha 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Fecha"
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
            Left            =   3900
            TabIndex        =   45
            Top             =   360
            Width           =   585
         End
         Begin VB.Label lblNumReq 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Requisición"
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
            Left            =   150
            TabIndex        =   46
            Top             =   360
            Width           =   1080
         End
         Begin VB.Label lblProveedor 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Proveedor"
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
            Left            =   150
            TabIndex        =   40
            Top             =   1575
            Width           =   1005
         End
      End
      Begin HSFlatControls.MyCombo cboERBusqueda 
         Height          =   375
         Left            =   -65880
         TabIndex        =   69
         Top             =   320
         Width           =   3225
         _ExtentX        =   5689
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
      Begin HSFlatControls.MyCombo cboTRBusqueda 
         Height          =   375
         Left            =   -70200
         TabIndex        =   67
         Top             =   320
         Width           =   3225
         _ExtentX        =   5689
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdRequisiciones 
         Height          =   8985
         Left            =   -74880
         TabIndex        =   70
         ToolTipText     =   "Requisiciones"
         Top             =   840
         Width           =   13590
         _ExtentX        =   23971
         _ExtentY        =   15849
         _Version        =   393216
         ForeColor       =   0
         Cols            =   7
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
         FormatString    =   "|Número|Fecha|Tipo requisición|Estado|Almacén surte|Empleado realizó"
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
         _Band(0).Cols   =   7
      End
      Begin VB.Frame fraArticulos 
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
         Height          =   3765
         Left            =   120
         TabIndex        =   37
         Top             =   5385
         Width           =   13635
         Begin VB.TextBox txtCaptura 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
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
            Height          =   375
            Left            =   10320
            TabIndex        =   62
            Top             =   2120
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.Frame Frame6 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            Caption         =   "Nombre comercial completo"
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
            Left            =   90
            TabIndex        =   72
            Top             =   1950
            Width           =   10530
            Begin VB.TextBox txtnombrecompleto 
               Appearance      =   0  'Flat
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
               ForeColor       =   &H80000011&
               Height          =   490
               Left            =   60
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   73
               TabStop         =   0   'False
               Top             =   265
               Width           =   9660
            End
            Begin VB.Shape Shape1 
               BorderColor     =   &H80000015&
               Height          =   525
               Left            =   45
               Top             =   255
               Width           =   9690
            End
            Begin VB.Label Label11 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Nombre comercial completo"
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
               Left            =   130
               TabIndex        =   89
               Top             =   0
               Width           =   3015
            End
         End
         Begin VB.Frame fraTotales 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
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
            Height          =   880
            Left            =   9960
            TabIndex        =   74
            Top             =   2770
            Visible         =   0   'False
            Width           =   1785
            Begin VB.TextBox txtTotal 
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
               Left            =   70
               TabIndex        =   75
               Top             =   285
               Width           =   1620
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               Caption         =   "Costo estimado"
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
               Left            =   120
               TabIndex        =   76
               Top             =   30
               Width           =   1560
            End
         End
         Begin VB.Frame fraObservaciones 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000008&
            Height          =   850
            Left            =   90
            TabIndex        =   77
            Top             =   2805
            Width           =   10530
            Begin VB.TextBox txtObservaciones 
               Appearance      =   0  'Flat
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
               Height          =   495
               Left            =   60
               MaxLength       =   200
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   79
               Top             =   265
               Width           =   9660
            End
            Begin VB.Shape Shape2 
               BorderColor     =   &H80000015&
               Height          =   525
               Left            =   45
               Top             =   255
               Width           =   9690
            End
            Begin VB.Label Label12 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Observaciones"
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
               Left            =   130
               TabIndex        =   90
               Top             =   0
               Width           =   1575
            End
         End
         Begin MyCommandButton.MyButton cmdManejos 
            Height          =   375
            Left            =   12465
            TabIndex        =   86
            Top             =   3060
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
            BackColorDown   =   -2147483643
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   "Manejos"
            DepthEvent      =   1
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdBorrar 
            Height          =   495
            Left            =   13065
            TabIndex        =   28
            ToolTipText     =   "Eliminar artículo"
            Top             =   2040
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   873
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
            Picture         =   "frmRequisicion.frx":A1AE
            BackColorDown   =   -2147483643
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            CaptionPosition =   4
            DepthEvent      =   1
            ForeColorDisabled=   -2147483629
            ForeColorOver   =   13003064
            ForeColorFocus  =   13003064
            ForeColorDown   =   13003064
            PictureDisabled =   "frmRequisicion.frx":A8C2
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdCancelar 
            Height          =   495
            Left            =   12570
            TabIndex        =   27
            ToolTipText     =   "Cancelar artículo"
            Top             =   2040
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   873
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
            Picture         =   "frmRequisicion.frx":AFD6
            BackColorDown   =   -2147483643
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   15790320
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmRequisicion.frx":B6EA
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MSMask.MaskEdBox txtEdit 
            Height          =   390
            Left            =   3450
            TabIndex        =   50
            Top             =   5265
            Visible         =   0   'False
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   688
            _Version        =   393216
            BorderStyle     =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Century Gothic"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   " "
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdArticulos 
            Height          =   1710
            Left            =   90
            TabIndex        =   26
            ToolTipText     =   "Lista de artículos"
            Top             =   225
            Width           =   13470
            _ExtentX        =   23760
            _ExtentY        =   3016
            _Version        =   393216
            ForeColor       =   0
            Rows            =   0
            Cols            =   16
            FixedRows       =   0
            FixedCols       =   0
            ForeColorFixed  =   0
            ForeColorSel    =   0
            BackColorBkg    =   -2147483643
            BackColorUnpopulated=   -2147483643
            GridColor       =   -2147483638
            GridColorFixed  =   -2147483638
            GridColorUnpopulated=   -2147483638
            AllowBigSelection=   0   'False
            FocusRect       =   0
            GridLinesFixed  =   1
            GridLinesUnpopulated=   1
            Appearance      =   0
            FormatString    =   $"frmRequisicion.frx":BDFE
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
            _Band(0).Cols   =   16
         End
         Begin VB.Label lblArticuloNuevo 
            BackColor       =   &H80000005&
            Caption         =   "Artículo nuevo"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   300
            Left            =   12120
            TabIndex        =   87
            Top             =   2640
            Width           =   1425
         End
         Begin VB.Label lblPedidoSug 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Se sugiere"
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
            Height          =   300
            Left            =   150
            TabIndex        =   47
            Top             =   5595
            Width           =   8055
         End
         Begin VB.Label lblExisTotalAlm 
            BackColor       =   &H80000005&
            Caption         =   "Existencia total en almacenes"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   150
            TabIndex        =   48
            Top             =   5205
            Width           =   8025
         End
         Begin VB.Label lblExisTotalDpto 
            BackColor       =   &H80000005&
            Caption         =   "Existencia total en departamentos"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   150
            TabIndex        =   49
            Top             =   5400
            Width           =   8025
         End
      End
      Begin VB.Label Label4 
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
         Left            =   -71760
         TabIndex        =   66
         Top             =   380
         Width           =   1470
      End
      Begin VB.Label Label5 
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
         Height          =   255
         Left            =   -66720
         TabIndex        =   68
         Top             =   380
         Width           =   660
      End
   End
End
Attribute VB_Name = "frmRequisicion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------
'| Nombre del Proyecto      : Almacén
'| Nombre del Formulario    : frmRequisicion
'----------------------------------------------------------------------------------------
'| Objetivo: Realiza las requisiciones de tipo salida departamento, reubicacion y compra
'----------------------------------------------------------------------------------------

Option Explicit

Public vllngNumeroOpcionModulo As Long
Public intSubFamilia As Integer
Public intFamilia As Integer

Private vgrptReporte As CRAXDRT.Report

'Para las columnas del grdArticulos
Const intColManejo1 = 1
Const intColManejo2 = 2
Const intColManejo3 = 3
Const intColManejo4 = 4
Const intColCveArticulo = 5
Const intColNombreComercial = 6
Const intColCantidad = 7
Const intColDescripcionUnidad = 8
Const intColExistencia = 9
Const intColEstado = 10
Const intColUnidad = 11
Const intColProveedor = 12
Const intColFecha = 13
Const intColCantidadProveedor = 14
Const intColCosto = 15

Dim rsIvRequisicionMaestro As New ADODB.Recordset
Dim rsIvRequisicionDetalle As New ADODB.Recordset
Dim rsDepartamento As ADODB.Recordset
Dim rstipodepapel As ADODB.Recordset
Dim rs As New ADODB.Recordset

Dim vlintNumDepartamento As Integer
Dim lintAutorizaRequi As Integer
Dim lintSoloMaestro As Integer
Dim lintControlado As Integer
Dim intMaxManejos As Integer
Dim intcontador As Integer
Dim vlintCol As Integer

Dim lstrUnidadAlterna As String
Dim lstrUnidadMinima As String
Dim vlStrSQL As String
Dim vlstrx As String
Dim strSQL As String

Dim lblnAlmacenConsigna As Boolean
Dim vlblnControlado As Boolean
Dim vlblnConsulta As Boolean
Dim vlblnBusqueda As Boolean

Dim llngDeptoAutoriza As Long
Dim llngDeptoRecibe As Long
Dim lngRow As Long

Dim blnPermisoCompraDirecta As Boolean  'Si el usuario tiene permiso para realizar requisiciones compra directa
Dim vlintColCLick As Integer ' para grabar la columna sobre la que s da click

Dim rsAutorizaciones As New ADODB.Recordset
Dim vlstrEstadoMaestroIni As String
Dim vlstrEstadoMaestroFin As String
Dim vlblnNoFocus As Boolean
'Indica si se podrá seleccionar el proveedor cuando sea una compra directa, se imprimirá un reprote diferente
Dim lblnSeleccionarProveedor As Boolean
Dim ldblTipoCambio As Double
Dim lblnReubicarInsumos As Boolean  'Indica si se pueden reubicar insumos (cuando el departamento que solicita es almacén)
Dim lstrArticuloExcede As String
Dim lintDiasPendienteRequi As Integer   'Dias anteriores que se toman para validar la cantidad pendiente de surtir del articulo
Dim lblnValidarExcedeMaximo As Boolean  'indica si se debe validar que no se exceda la cantidad máxima configurada para el artículo
Dim llngCveCajaMaterialRequi As Long    'Indica la clave de la caja de material incluida en la requi (articulos agregados al grid), -1 ninguna
Dim lblnRecargarCajas As Boolean        'Indica si se consulta nuevamente las cajas de material
Dim llngDeptoSubrogado As Long          'Indica el depto subrogado configurado en el parametro de IV
Dim lintInterfazFarmaciaSJP As Integer     'Indica si se usa la interfaz de farmacia subrogada de SJP
Dim position As Integer

Private Sub cboAlmacenSurte_GotFocus()
1     On Error GoTo NotificaError
          
2         If vlblnConsulta Then
3             If fblnCanFocus(optOpcion(0)) Then
4                 optOpcion(0).SetFocus
5             Else
6                 If fblnCanFocus(grdArticulos) Then grdArticulos.SetFocus
7             End If
8         End If

9     Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboAlmacenSurte_GotFocus" & " Linea:" & Erl()))
        Unload Me
End Sub

Private Sub cboAlmacenSurte_KeyPress(KeyAscii As Integer)
1     On Error GoTo NotificaError
      Dim rs As New ADODB.Recordset
          
2         If KeyAscii = vbKeyReturn Then
          
3             If cboTipoRequisicion.ListCount = 0 Or cboTipoRequisicion.ListIndex = -1 Then
4                 cboTipoRequisicion.SetFocus
5             Else

6                 vlblnControlado = False
7                 Set rs = frsRegresaRs("SELECT bitControlado " & _
                    "FROM IvRequisicionDepartamento " & _
                    "WHERE intNumeroLogin=" & Str(vglngNumeroLogin) & _
                    " AND smiCveDepartamento=" & cboAlmacenSurte.ItemData(cboAlmacenSurte.ListIndex) & _
                    " AND chrTipoRequisicion='" & IIf(Trim(cboTipoRequisicion.List(cboTipoRequisicion.ListIndex)) = "SALIDA A DEPARTAMENTO", "SD", IIf(Trim(cboTipoRequisicion.List(cboTipoRequisicion.ListIndex)) = "REUBICACION", "RE", IIf(Trim(cboTipoRequisicion.List(cboTipoRequisicion.ListIndex)) = "ALMACEN ABASTECIMIENTO", "AG", "CO"))) & "'")
8                 If rs.RecordCount > 0 Then
9                   If ((rs!bitControlado = 1) Or (rs!bitControlado = True)) Then
10                    vlblnControlado = True
11                  End If
12                End If
13                pHabilita False, False, False, False, False, True, False
14                fraCabecera.Enabled = False
                  
15                cmdAgregarNuevo.Enabled = IIf(Trim(cboTipoRequisicion.List(cboTipoRequisicion.ListIndex)) = "COMPRA - PEDIDO", True, False)
16                fraFiltros.Enabled = True
17                fraArticulos.Enabled = True
18                If Trim(cboTipoRequisicion.List(cboTipoRequisicion.ListIndex)) = "REUBICACION" Then
19                    lblCajaMaterial.Enabled = True
20                    cboCajaMaterial.Enabled = True
21                    cboCajaMaterial.ListIndex = 0 '<NINGUNA>
22                End If
23                optOpcion(0).SetFocus
                 
24            End If
              
25        End If

26    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboAlmacenSurte_KeyPress" & " Linea:" & Erl()))
        Unload Me
End Sub

Private Sub cboCajaMaterial_Click()
1     On Error GoTo NotificaError

              'No se ha agregado ninguna caja
2             pLimpiaArticulo
3             If cboCajaMaterial.ListIndex = 0 Then
                  'Ninguno
4                 pHabilitaArticulo True
5                 chkPedirMaximo.Enabled = True
6             ElseIf cboCajaMaterial.ListIndex > 0 Then
                  'alguna caja
7                 chkPedirMaximo.Enabled = False
8                 pHabilitaArticulo False
9             End If
      '    End If
          
10    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboCajaMaterial_Click" & " Linea:" & Erl()))
End Sub

Private Sub pHabilitaArticulo(blnHabilita As Boolean)
1     On Error GoTo NotificaError
          
2         lblLocalizacion.Enabled = blnHabilita
3         cboLocalizacion.Enabled = blnHabilita
4         lblFamilia.Enabled = blnHabilita
5         cboFamilia.Enabled = blnHabilita
6         lblSubfamilia.Enabled = blnHabilita
7         cboSubfamilia.Enabled = blnHabilita
8         lblComercial.Enabled = blnHabilita
9         cboNombreComercial.Enabled = blnHabilita
10        lblGenerico.Enabled = blnHabilita
11        txtNombreGenerico.Enabled = blnHabilita
12        lblCodigoBarras.Enabled = blnHabilita
13        txtCodigoBarras.Enabled = blnHabilita
14        lblClave.Enabled = blnHabilita
15        txtClave.Enabled = blnHabilita
16        lblExistencia.Enabled = blnHabilita 'FIXME
17        txtExistencia.Enabled = blnHabilita
18        lblCantidad.Enabled = blnHabilita   'FIXME
19        txtCantidadSolicitada.Enabled = blnHabilita
20        OptAlterna.Enabled = blnHabilita
21        OptMinima.Enabled = blnHabilita
22        optOpcion(0).Enabled = blnHabilita
23        optOpcion(1).Enabled = blnHabilita
24        optOpcion(2).Enabled = blnHabilita
25        optOpcion(3).Enabled = blnHabilita
          
26    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pHabilitaArticulo" & " Linea:" & Erl()))
End Sub

Private Sub pLimpiaArticulo()
1     On Error GoTo NotificaError
2         cboLocalizacion.ListIndex = 0
          
3         cboFamilia.Clear
4         cboFamilia.AddItem "<TODAS>", 0
5         cboFamilia.ItemData(cboFamilia.NewIndex) = 0
6         cboFamilia.ListIndex = 0
          
7         cboSubfamilia.Clear
8         cboSubfamilia.AddItem "<TODAS>", 0
9         cboSubfamilia.ItemData(cboSubfamilia.NewIndex) = 0
10        cboSubfamilia.ListIndex = 0
          
11        cboNombreComercial.Clear
12        cboNombreComercial.AddItem "<TODOS>", 0
13        cboNombreComercial.ItemData(cboNombreComercial.NewIndex) = 0
14        cboNombreComercial.ListIndex = 0
          
15        txtNombreGenerico.Text = ""
16        txtCodigoBarras.Text = ""
17        txtClave.Text = ""
18        txtExistencia.Text = ""
19        txtCantidadSolicitada.Text = ""
20        OptAlterna.Value = True
21        OptMinima.Value = False
          
22    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLimpiaArticulo" & " Linea:" & Erl()))
End Sub

Private Sub cboCajaMaterial_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If fblnCanFocus(cmdAgregar) Then
            cmdAgregar.SetFocus
        End If
    End If
End Sub

Private Sub cboDepartamentoSolicita_Click()
On Error GoTo NotificaError
Dim rs As New ADODB.Recordset
    lblnReubicarInsumos = False
    If cboDepartamentoSolicita.ListIndex <> -1 Then
        Set rs = frsEjecuta_SP(CStr(cboDepartamentoSolicita.ItemData(cboDepartamentoSolicita.ListIndex)), "SP_GNSELDEPARTAMENTOS")
        If rs.RecordCount <> 0 Then lblnReubicarInsumos = rs!chrClasificacion = "A"
        rs.Close
    End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboDepartamentoSolicita_Click"))
End Sub

Private Sub cbodepartamentosolicita_GotFocus()
1     On Error GoTo NotificaError
          
2         If vlblnConsulta Then
3             If fblnCanFocus(optOpcion(0)) Then
4                 optOpcion(0).SetFocus
5             Else
6                 If fblnCanFocus(grdArticulos) Then grdArticulos.SetFocus
7             End If
8         End If
          
9     Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboDepartamentoSolicita_GotFocus" & " Linea:" & Erl()))
        Unload Me
End Sub

Private Sub cbodepartamentosolicita_KeyPress(KeyAscii As Integer)
1     On Error GoTo NotificaError
          
2         If KeyAscii = vbKeyReturn Then
3             pCargaLocalizacion
4             cboTipoRequisicion.SetFocus
5         End If

6     Exit Sub
NotificaError:
         Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboDepartamentoSolicita_KeyPress" & " Linea:" & Erl()))
         Unload Me
End Sub

Private Sub cboERBusqueda_Click()
On Error GoTo NotificaError
    
    If cboERBusqueda.ListIndex <> -1 Then pCargaRequisiciones

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboERBusqueda_Click"))
    Unload Me
End Sub

Private Sub cboERBusqueda_KeyPress(KeyAscii As Integer)
1     On Error GoTo NotificaError
          
2         If KeyAscii = vbKeyReturn Then
3             pCargaRequisiciones
4             grdRequisiciones.SetFocus
5         End If

6     Exit Sub
NotificaError:
         Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboERBusqueda_KeyPress" & " Linea:" & Erl()))
         Unload Me
End Sub

Private Sub cboFamilia_Click()
1     On Error GoTo NotificaError
          
2         If cboFamilia.ListIndex <> -1 Then
3             If cboFamilia.ItemData(cboFamilia.ListIndex) <> 0 Then
4                 vlstrx = "" & _
                  "select " & _
                      "vchDescripcion Descripcion," & _
                      "chrCveSubFamilia Clave " & _
                  "From " & _
                      "IvSubFamilia " & _
                  "Where " & _
                      "chrCveFamilia=" & cboFamilia.ItemData(cboFamilia.ListIndex) & " and " & _
                      "bitactivo = 1 " & _
                      "and chrCveArtMedicamen = " & Trim(Str(IIf(optOpcion(1).Value, 0, IIf(optOpcion(2).Value, 1, 2)))) & " " & _
                  "Union " & _
                  "select " & _
                      "'<TODAS>' Descripcion," & _
                      "0 Clave FROM dual"
                  
5                 Set rs = frsRegresaRs(vlstrx)
                  
6                 pLlenarCboRs_new cboSubfamilia, rs, 1, 0
                  
7                 cboSubfamilia.ListIndex = 0
8             End If
9         End If

10    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboFamilia_Click" & " Linea:" & Erl()))
        Unload Me
End Sub

Private Sub cboFamilia_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    If KeyAscii = vbKeyReturn Then cboSubfamilia.SetFocus
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboFamilia_KeyPress"))
    Unload Me
End Sub

Private Sub cboLocalizacion_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    If KeyAscii = vbKeyReturn Then cboFamilia.SetFocus
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboLocalizacion_KeyPress"))
    Unload Me
End Sub

Private Sub cboNombreComercial_Change()
1     On Error GoTo NotificaError

2         lblCantidad.Enabled = cboNombreComercial.ListIndex <> -1
3         txtCantidadSolicitada.Enabled = cboNombreComercial.ListIndex <> -1
4         OptAlterna.Enabled = cboNombreComercial.ListIndex <> -1
5         OptMinima.Enabled = cboNombreComercial.ListIndex <> -1

6     Exit Sub
NotificaError:
         Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboNombreComercial_Change" & " Linea:" & Erl()))
         Unload Me
End Sub

Private Sub cboNombreComercial_Click()
1     On Error GoTo NotificaError
          
2         If cboNombreComercial.ListIndex <> -1 Then
3             If Trim(cboNombreComercial.List(cboNombreComercial.ListIndex)) = "<TODOS>" Then
4                 If cboNombreComercial.ListCount > 1 Then
5                     cboNombreComercial.RemoveItem 1
6                 End If
              
7                 txtNombreGenerico.Text = ""
8                 txtCodigoBarras.Text = ""
9                 txtClave.Text = ""
10                txtExistencia.Text = ""
11                txtCantidadSolicitada.Text = ""
12                OptAlterna.Value = True
13                OptMinima.Value = False
14            End If
15        End If

16    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboNombreComercial_Click" & " Linea:" & Erl()))
        Unload Me
End Sub

Private Sub cboNombreComercial_KeyPress(KeyAscii As Integer)
1     On Error GoTo NotificaError

2         KeyAscii = Asc(StrConv(Chr$(KeyAscii), vbUpperCase))
          
3         If KeyAscii = vbKeyReturn Then
          
4             If cboNombreComercial.ListIndex = -1 Then
5                 If Len(cboNombreComercial.Text) > 0 Then
6                     txtNombreGenerico.Text = ""
7                     txtCodigoBarras.Text = ""
8                     txtClave.Text = ""
9                     txtExistencia.Text = ""
10                    txtCantidadSolicitada.Text = ""
                      
11                    vgstrVarIntercam = UCase(cboNombreComercial.Text)
12                    vgstrVarIntercam2 = "Lista por nombre comercial"
13                    vgstrNombreCbo = "Comercial"
                      
14                    frmLista.gintEstatus = 1
15                    frmLista.gintFamilia = cboFamilia.ItemData(cboFamilia.ListIndex)
16                    frmLista.gintSubfamilia = cboSubfamilia.ItemData(cboSubfamilia.ListIndex)
17                    frmLista.Tag = IIf(cboTipoRequisicion.List(cboTipoRequisicion.ListIndex) = "REUBICACION" And lblnReubicarInsumos, "T", Mid(cboTipoRequisicion.List(cboTipoRequisicion.ListIndex), 1, 1))
18                    frmLista.Show vbModal, Me
                      
19                    If Trim(vgstrVarIntercam) <> "" Then
20                        pCargaArticulo vgstrVarIntercam
21                    End If
22                Else
23                    cboNombreComercial.ListIndex = 0
24                    txtCodigoBarras.SetFocus
25                End If
26            Else
27                txtCodigoBarras.SetFocus
28            End If
          
29        End If

30    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboNombreComercial_KeyPress" & " Linea:" & Erl()))
        Unload Me
End Sub

Private Sub pCargaArticulo(vlstrCveArticulo As String)
1     On Error GoTo NotificaError
          Dim lstrDeptoConsigna As String, vlStrSentenciaSQL As String
          Dim rsAutorizaciones As New ADODB.Recordset
          
2         cboNombreComercial.Clear
3         cboNombreComercial.AddItem "<TODOS>", 0
4         cboNombreComercial.ItemData(cboNombreComercial.NewIndex) = 0
5         cboNombreComercial.ListIndex = 0

          'Set rs = frsEjecuta_SP(Trim(vlstrCveArticulo) & "|" & "|", "sp_IvSelArticulo")
6         If vlstrCveArticulo <> "" And frmLista.bvchDescripcion <> "" Then
7             vlstrCveArticulo = ""
8         End If
          
9         Set rs = frsEjecuta_SP(Trim(vlstrCveArticulo) & "|" & frmLista.bvchDescripcion & "|", "sp_IvSelArticulo")

10        If rs.RecordCount <> 0 Then
11            rs.MoveFirst
12            cboNombreComercial.AddItem vgstrVarIntercam2, 1
13            cboNombreComercial.ItemData(cboNombreComercial.NewIndex) = Val(rs!intIdArticulo)
14            cboNombreComercial.ListIndex = 1
              
15            txtNombreGenerico.Text = rs!NombreGenerico
16            txtCodigoBarras.Text = rs!CodigoBarras
17            txtClave.Text = rs!chrcvearticulo
                  
18            lintControlado = rs!bitControlado
19            lstrUnidadMinima = rs!UnidadMinima
20            lstrUnidadAlterna = rs!UnidadAlterna
              
21            lblCantidad.Enabled = True
22            txtCantidadSolicitada.Enabled = IIf(chkPedirMaximo.Value = 0, True, False)
23            OptAlterna.Enabled = True
              
24            If lstrUnidadMinima = lstrUnidadAlterna Then
25                OptMinima.Enabled = False
26            Else
27                OptMinima.Enabled = True
28            End If
              
29            vgstrParametrosSP = Str(IIf(Trim(cboTipoRequisicion.List(cboTipoRequisicion.ListIndex)) = "COMPRA - PEDIDO", cboDepartamentoSolicita.ItemData(cboDepartamentoSolicita.ListIndex), IIf(Trim(cboTipoRequisicion.List(cboTipoRequisicion.ListIndex)) = "ALMACEN ABASTECIMIENTO", cboDepartamentoSolicita.ItemData(cboDepartamentoSolicita.ListIndex), cboAlmacenSurte.ItemData(cboAlmacenSurte.ListIndex)))) & "|" & Trim(txtClave.Text)
30            Set rs = frsEjecuta_SP(vgstrParametrosSP, "sp_IvSelExistencia")
31            If rs.RecordCount <> 0 Then
              ''las compras pedido siempre son en unidades alternas''''''''''''''''''''''''''''''''''''''''
32                If Trim(cboTipoRequisicion.List(cboTipoRequisicion.ListIndex)) = "COMPRA - PEDIDO" Then
33                    If rs!intContenido = 1 Then
34                       txtExistencia.Text = rs!intExistenciaDeptouv
35                    Else
36                       txtExistencia.Text = Int(rs!intExistenciaDeptouv + (rs!intexistenciadeptoum / IIf(IsNull(rs!intContenido), 1, rs!intContenido)))
37                    End If
               ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
38                Else
39                txtExistencia.Text = rs!existencia
40                End If
41            Else
                    lstrDeptoConsigna = Str(cboDepartamentoSolicita.ItemData(cboDepartamentoSolicita.ListIndex))
                    vlStrSentenciaSQL = "select count(*) Articulo from IVUBICACION WHERE SMICVEDEPARTAMENTO = " & lstrDeptoConsigna & " AND CHRCVEARTICULO = " & Trim(txtClave.Text)
                    Set rsAutorizaciones = frsRegresaRs(vlStrSentenciaSQL, adLockOptimistic)
                    If rsAutorizaciones!Articulo = 0 Then
                    If Trim(cboTipoRequisicion.List(cboTipoRequisicion.ListIndex)) <> "COMPRA - PEDIDO" Then
                        lstrDeptoConsigna = Str(cboAlmacenSurte.ItemData(cboAlmacenSurte.ListIndex))
                        vlStrSentenciaSQL = "select count(*) Total from NoDepartamento where bitconsignacion = 0 and  SMICVEDEPARTAMENTO = " & lstrDeptoConsigna
                        Set rsAutorizaciones = frsRegresaRs(vlStrSentenciaSQL, adLockOptimistic)
                        If rsAutorizaciones!Total = 1 Then
                            MsgBox "El artículo se encuentra desasignado del departamento", vbOKOnly + vbInformation, "Mensaje"
                            txtCantidadSolicitada.Enabled = IIf(txtExistencia.Text = "0", False, True)
                            rsAutorizaciones.Close
                            Exit Sub
                        Else
42                        txtExistencia.Text = "0"
                        End If
                        rsAutorizaciones.Close
                    Else
                            txtExistencia.Text = "0"
                    End If
'43            End If
                Else
                    txtExistencia.Text = "0"
                    rsAutorizaciones.Close
                End If
                End If
                
44            If Trim(cboTipoRequisicion.List(cboTipoRequisicion.ListIndex)) = "REUBICACION" Then
45                cboCajaMaterial.ListIndex = 0
46            End If
              
47            If chkPedirMaximo.Value = 0 Then
48                txtCantidadSolicitada.SetFocus
49            Else
50                cmdAgregar.SetFocus
51            End If
              
52            Call Unidades
              
53        End If

54    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCargaArticulo" & " Linea:" & Erl()))
        Unload Me
End Sub

Private Sub cboProveedor_Click()
On Error GoTo NotificaError
    If cboProveedor.ListIndex > 0 Then
        chkIncluir.Value = 0
        chkIncluir.Enabled = False
    Else
        chkIncluir.Enabled = True
    End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboProveedor_Click"))
End Sub

Private Sub cboProveedor_GotFocus()
On Error GoTo NotificaError
    If vlblnConsulta Then
        If fblnCanFocus(optOpcion(0)) Then
            optOpcion(0).SetFocus
        Else
            If fblnCanFocus(grdArticulos) Then grdArticulos.SetFocus
        End If
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboProveedor_GotFocus"))
End Sub

Private Sub cboProveedor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cboAlmacenSurte.SetFocus
End Sub

Private Sub cboSubfamilia_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    If KeyAscii = vbKeyReturn Then cboNombreComercial.SetFocus
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboSubfamilia_KeyPress"))
    Unload Me
End Sub

Private Sub cboTipoRequisicion_Click()
1     On Error GoTo NotificaError
      Dim X As Long
          
2         If cboTipoRequisicion.ListIndex <> -1 Then
          
3             chkIncluir.Enabled = Trim(cboTipoRequisicion.List(cboTipoRequisicion.ListIndex)) = "COMPRA - PEDIDO"
4             chkIncluir.Value = IIf(Not chkIncluir.Enabled, 0, chkIncluir.Value)
5             chkCompradirecta.Enabled = IIf(Trim(cboTipoRequisicion.List(cboTipoRequisicion.ListIndex)) = "COMPRA - PEDIDO" And blnPermisoCompraDirecta, True, False)
6             chkCompradirecta.Value = IIf(Trim(cboTipoRequisicion.List(cboTipoRequisicion.ListIndex)) <> "COMPRA - PEDIDO", 0, chkCompradirecta.Value)
          
7             cboAlmacenSurte.Clear
8             cboAlmacenSurte.ListIndex = -1
9             If (Trim(cboTipoRequisicion.List(cboTipoRequisicion.ListIndex)) = "ALMACEN ABASTECIMIENTO") Then
                ' Extrae el valor de vchProveedorAlmacenGeneral, si es diferente a "Vacio" o "Nulo" entonces permite Almacén General
10              Set rs = frsRegresaRs("SELECT cp.vchNombre FROM coproveedor cp WHERE cp.intCveProveedor = " & vglngCveAlmacenGeneral)
11              If (rs.State <> adStateClosed) Then
12                If rs.RecordCount > 0 Then
13                  rs.MoveFirst
14                  If Not IsNull(rs!VCHNOMBRE) Then
15                    If (Trim(rs!VCHNOMBRE) <> "") Then
16                      cboAlmacenSurte.AddItem (Trim(rs!VCHNOMBRE))
17                      cboAlmacenSurte.ItemData(cboAlmacenSurte.NewIndex) = 0
18                      cboAlmacenSurte.ListIndex = 0
19                    End If
20                  End If
21                End If
22              End If
23            Else
              
24              vlstrx = "" & _
                "select " & _
                    "NoDepartamento.vchDescripcion," & _
                    "NoDepartamento.smiCveDepartamento " & _
                "From " & _
                    "IvRequisicionDepartamento " & _
                    "inner join NoDepartamento on IvRequisicionDepartamento.smiCveDepartamento = NoDepartamento.smiCveDepartamento " & _
                "Where " & _
                    "IvRequisicionDepartamento.chrTipoRequisicion='" & IIf(Trim(cboTipoRequisicion.List(cboTipoRequisicion.ListIndex)) = "SALIDA A DEPARTAMENTO", "SD", IIf(Trim(cboTipoRequisicion.List(cboTipoRequisicion.ListIndex)) = "REUBICACION", "RE", "CO")) & "' " & _
                    "and IvRequisicionDepartamento.intNumeroLogin = " & Str(vglngNumeroLogin)
                  
25              Set rs = frsRegresaRs(vlstrx)
                
26              If rs.RecordCount <> 0 Then
                    
27                  rs.MoveFirst
28                  Do While Not rs.EOF
29                      cboAlmacenSurte.AddItem rs!vchdescripcion
30                      cboAlmacenSurte.ItemData(cboAlmacenSurte.NewIndex) = rs!smicvedepartamento
31                      rs.MoveNext
32                  Loop
33                  cboAlmacenSurte.ListIndex = 0
34              End If
                
35            End If
36            If Trim(cboTipoRequisicion.List(cboTipoRequisicion.ListIndex)) = "COMPRA - PEDIDO" Or Trim(cboTipoRequisicion.List(cboTipoRequisicion.ListIndex)) = "REUBICACION" Or Trim(cboTipoRequisicion.List(cboTipoRequisicion.ListIndex)) = "ALMACEN ABASTECIMIENTO" Then
37                pFormatoArticulos True
38            Else
39                pFormatoArticulos False
40            End If
41        End If

42    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboTipoRequisicion_Click" & " Linea:" & Erl()))
        Unload Me
End Sub

Private Sub cboTipoRequisicion_GotFocus()
On Error GoTo NotificaError
    If vlblnConsulta Then
        If fblnCanFocus(optOpcion(0)) Then
            optOpcion(0).SetFocus
        Else
            If fblnCanFocus(grdArticulos) Then grdArticulos.SetFocus
        End If
    End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboTipoRequisicion_GotFocus"))
    Unload Me
End Sub

Private Sub cboTipoRequisicion_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    
    If KeyAscii = vbKeyReturn Then
        If chkCompradirecta.Enabled Then
            chkCompradirecta.SetFocus
        Else
            If chkIncluir.Enabled Then
                chkIncluir.SetFocus
            Else
                cboAlmacenSurte.SetFocus
            End If
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboTipoRequisicion_KeyPress"))
    Unload Me
End Sub

Private Sub cboTRBusqueda_Click()
On Error GoTo NotificaError
    If cboTRBusqueda.ListIndex <> -1 Then pCargaRequisiciones
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboTRBusqueda_Click"))
    Unload Me
End Sub

Private Sub pCargaRequisiciones()
1     On Error GoTo NotificaError
      Dim rs1 As New ADODB.Recordset
      Dim X As Long
      Dim vlstrProveedorAlmacenGeneral As String
      Dim strTipo As String
      Dim strEstatus As String
      Dim strParametros As String

2         vlstrProveedorAlmacenGeneral = ""
3         Set rs1 = frsRegresaRs("SELECT cp.vchNombre FROM coproveedor cp WHERE cp.intCveProveedor = " & vglngCveAlmacenGeneral)
4         If (rs1.State <> adStateClosed) Then
5           If rs1.RecordCount > 0 Then
6             rs1.MoveFirst
7             If Not IsNull(rs1!VCHNOMBRE) Then
8               If (Trim(rs1!VCHNOMBRE) <> "") Then
9                 vlstrProveedorAlmacenGeneral = Trim(rs1!VCHNOMBRE)
10              End If
11            End If
12          End If
13          rs1.Close
14        End If
          
15        With grdRequisiciones
16            .Rows = 2
17            .Cols = 7
18            For X = 1 To 6
19                .TextMatrix(1, X) = ""
20            Next X
21        End With

22        strTipo = IIf(cboTRBusqueda.ListIndex = 0, "*", IIf(Trim(cboTRBusqueda.List(cboTRBusqueda.ListIndex)) = "REUBICACION", "R", IIf(Trim(cboTRBusqueda.List(cboTRBusqueda.ListIndex)) = "SALIDA A DEPARTAMENTO", "D", IIf(Trim(cboTRBusqueda.List(cboTRBusqueda.ListIndex)) = "ALMACEN ABASTECIMIENTO", "A", IIf(Trim(cboTRBusqueda.List(cboTRBusqueda.ListIndex)) = "CONSIGNACION", "O", IIf(Trim(cboTRBusqueda.List(cboTRBusqueda.ListIndex)) = "PEDIDO SUGERIDO", "U", "C"))))))
23        strEstatus = IIf(Trim(cboERBusqueda.List(cboERBusqueda.ListIndex)) = "<TODOS>", "*", cboERBusqueda.List(cboERBusqueda.ListIndex))
24        strParametros = CStr(vgintNumeroDepartamento) & "|" & fstrFechaSQL(mskFecIni.Text) & "|" & fstrFechaSQL(mskFecFin.Text) & "|" & strTipo & "|" & strEstatus
          
25        Set rs = frsEjecuta_SP(strParametros, "SP_IVSELCARGAREQUISICIONES")
26        If rs.RecordCount <> 0 Then
              ' 3= TipoRequision, 5=AlmacenSurte
27            pLlenarMshFGrdRsSobreEscribe grdRequisiciones, rs, -1, 3, "ALMACEN ABASTECIMIENTO", 5, vlstrProveedorAlmacenGeneral
28        End If
29        With grdRequisiciones
30            .FixedCols = 1
31            .FixedRows = 1
32            .FormatString = "|Número|Fecha|Tipo requisición|Estado|Almacén surte|Empleado"
33            .ColWidth(0) = 100
34            .ColWidth(1) = 1000
35            .ColWidth(2) = 1000
36            .ColWidth(3) = 2000
37            .ColWidth(4) = 1500
38            .ColWidth(5) = 2500
39            .ColWidth(6) = 3500
40            .ColAlignment(1) = flexAlignRightCenter
41            .ColAlignment(2) = flexAlignLeftCenter
42            .ColAlignment(3) = flexAlignLeftCenter
43            .ColAlignment(4) = flexAlignLeftCenter
44            .ColAlignment(5) = flexAlignLeftCenter
45            .ColAlignment(6) = flexAlignLeftCenter
46            .ColAlignmentFixed(1) = flexAlignCenterCenter
47            .ColAlignmentFixed(2) = flexAlignCenterCenter
48            .ColAlignmentFixed(3) = flexAlignCenterCenter
49            .ColAlignmentFixed(4) = flexAlignCenterCenter
50            .ColAlignmentFixed(5) = flexAlignCenterCenter
51            .ColAlignmentFixed(6) = flexAlignCenterCenter
52        End With

53    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCargaRequisiciones" & " Linea:" & Erl()))
        Unload Me
End Sub

Private Sub pLlenarMshFGrdRsSobreEscribe(ObjGrid As MSHFlexGrid, ObjRS As Recordset, vlstrColumnaData, vlintColumna As Integer, vlstrColumnaValor As String, vlIntSobreEscribeColumna As Integer, vlstrSobreEscribeColumnaValor As String)
      '----------------------------------------------------------------------------------------
      'Procedimiento para llenar un grid con datos de un record set
      ' vlintColumna dice que columna del record set se debe de llenar en el RowData
      ' Si no quieres meter nada en el RowData mandale un -1 o una columna que no exista
      '----------------------------------------------------------------------------------------
1     On Error GoTo NotificaError
      Dim vlintNumCampos As Integer
      Dim vlintNumReg As Integer
      Dim vlintSeqFil As Integer
      Dim vlintSeqCol As Integer
          
          ' vlIntColumna Posicion de la columna en el Query (Fields= (Posicion-1))
          ' vlStrColumnaValor Valor a buscar
          ' vlIntSobreEscribeColumna Posicion de la columna en el Query (Fields= (Posicion-1)) a cambiar
          ' vlstrSobreEscribeColumnaValor Valor a cambiar
       
2         vlintNumCampos = ObjRS.Fields.Count
3         If vlintNumCampos > 0 Then
4             vlintNumReg = ObjRS.RecordCount
5             If vlintNumReg > 0 Then
                  
6                 ObjGrid.ClearStructure
7                 ObjGrid.Cols = vlintNumCampos + 1
8                 ObjGrid.Rows = vlintNumReg + 1
9                 ObjGrid.FixedCols = 1
10                ObjGrid.FixedRows = 1
11                ObjRS.MoveFirst
                  
12                For vlintSeqFil = 1 To vlintNumReg
13                    For vlintSeqCol = 1 To vlintNumCampos
14                        If IsNull(ObjRS.Fields(vlintSeqCol - 1).Value) = True Then
15                            ObjGrid.TextMatrix(vlintSeqFil, vlintSeqCol) = ""
16                        Else
17                            If vlstrColumnaData <> "" Then
18                                If vlintSeqCol - 1 = Val(vlstrColumnaData) Then
19                                    ObjGrid.RowData(vlintSeqFil) = ObjRS.Fields(vlintSeqCol - 1)
20                                End If
21                            End If
22                            ObjGrid.TextMatrix(vlintSeqFil, vlintSeqCol) = ObjRS.Fields(vlintSeqCol - 1).Value
23                        End If
24                    Next vlintSeqCol
25                    If ObjRS.Fields(vlintColumna - 1).Value = vlstrColumnaValor Then
26                      ObjGrid.TextMatrix(vlintSeqFil, vlIntSobreEscribeColumna) = vlstrSobreEscribeColumnaValor
27                    End If
28                    ObjRS.MoveNext
29                Next vlintSeqFil
30            End If
31        End If
          
32    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLlenarMshFGrdRsSobreEscribe" & " Linea:" & Erl()))
        Exit Sub
End Sub

Private Sub cboTRBusqueda_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    
    If KeyAscii = vbKeyReturn Then
        cboERBusqueda.SetFocus
        pCargaRequisiciones
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboTRBusqueda_KeyPress"))
    Unload Me
End Sub

Private Sub chkCompradirecta_Click()
On Error GoTo NotificaError
    If lblnSeleccionarProveedor Then
        If chkCompradirecta.Value = 1 Then
            lblProveedor.Enabled = True
            cboProveedor.Enabled = True
            cboProveedor.ListIndex = 0
        Else
            cboProveedor.ListIndex = -1
            lblProveedor.Enabled = False
            cboProveedor.Enabled = False
        End If
    End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":chkCompradirecta_Click"))
End Sub

Private Sub chkCompradirecta_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If chkIncluir.Enabled Then
            chkIncluir.SetFocus
        Else
            If cboProveedor.Enabled And cboProveedor.Visible Then
                cboProveedor.SetFocus
            Else: cboAlmacenSurte.SetFocus
            End If
        End If
    End If
End Sub

Private Sub chkIncluir_Click()
1     On Error GoTo NotificaError
      Dim lngContador As Long
      Dim dblTotal As Double
      Dim rsTemp As New ADODB.Recordset

2         fraTotales.Visible = chkIncluir.Value = 1

3         grdArticulos.ColWidth(intColProveedor) = IIf(chkIncluir.Value = 1, 3000, 0)
4         grdArticulos.ColWidth(intColFecha) = IIf(chkIncluir.Value = 1, 1100, 0)
5         grdArticulos.ColWidth(intColCantidadProveedor) = IIf(chkIncluir.Value = 1, 1500, 0)
6         grdArticulos.ColWidth(intColCosto) = IIf(chkIncluir.Value = 1, 1500, 0)
          
7         If vlblnConsulta And chkIncluir.Value = 1 Then
          
8             fraTotales.Visible = True
          
9             dblTotal = 0
10            For lngContador = 1 To grdArticulos.Rows - 1
              
11                vgstrParametrosSP = grdArticulos.TextMatrix(lngContador, intColCveArticulo)
12                Set rsTemp = frsEjecuta_SP(vgstrParametrosSP, "SP_IVSELULTIMACOMPRAARTICULO")
                  
13                If rsTemp.RecordCount <> 0 Then
                      
14                    With grdArticulos
15                      .TextMatrix(lngContador, intColProveedor) = rsTemp!NombreComercial
16                      .TextMatrix(lngContador, intColFecha) = rsTemp!fecha
17                      .TextMatrix(lngContador, intColCantidadProveedor) = rsTemp!cantidad
18                      .TextMatrix(lngContador, intColCosto) = FormatCurrency(rsTemp!Costo, 4)
19                    End With
                      
20                    dblTotal = dblTotal + rsTemp!Costo * Val(grdArticulos.TextMatrix(lngContador, intColCantidad))
21                End If
              
22            Next lngContador
          
23            txtTotal.Text = FormatCurrency(dblTotal, 4)
          
24        End If

25    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":chkIncluir_Click" & " Linea:" & Erl()))
        Unload Me
End Sub

Private Sub chkIncluir_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    If KeyAscii = vbKeyReturn Then
        If cboProveedor.Enabled And cboProveedor.Visible Then
            cboProveedor.SetFocus
        Else
            cboAlmacenSurte.SetFocus
        End If
    End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":chkIncluir_KeyPress"))
End Sub

Private Sub chkPedirMaximo_Click()
On Error GoTo NotificaError
    Call Unidades
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":chkPedirMaximo_Click"))
    Unload Me
End Sub

Private Sub chkPedirMaximo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If fblnCanFocus(cmdAgregar) Then cmdAgregar.SetFocus
    End If
End Sub

Private Sub chkUrgente_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    
    If KeyAscii = vbKeyReturn Then
        If cboDepartamentoSolicita.Enabled Then
            cboDepartamentoSolicita.SetFocus
        Else
            cboTipoRequisicion.SetFocus
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":chkUrgente_KeyPress"))
    Unload Me
End Sub

Private Function fblnValido() As Boolean
1     On Error GoTo NotificaError
          Dim lngCol As Long
          
2         fblnValido = True
          
3         If txtEstatus.Text <> "PENDIENTE" Then
4             fblnValido = False
5             cmdBuscar.SetFocus
              'La información ha cambiado, consulte de nuevo
6             MsgBox SIHOMsg(381), vbOKOnly + vbInformation, "Mensaje"
7         End If
          
8         If cboCajaMaterial.ListIndex > 0 Then
9             If cboCajaMaterial.ItemData(cboCajaMaterial.ListIndex) = llngCveCajaMaterialRequi Then
10                fblnValido = False
                  'Los artículos en la caja ya han sido agregados.
11                MsgBox SIHOMsg(1552), vbOKOnly + vbInformation, "Mensaje"
12            Else
13                If Trim(grdArticulos.TextMatrix(1, intColCveArticulo)) <> "" Then
14                    If MsgBox(SIHOMsg("1551"), (vbYesNo + vbQuestion), "Mensaje") = vbYes Then
15                        llngCveCajaMaterialRequi = cboCajaMaterial.ItemData(cboCajaMaterial.ListIndex)
16                        grdArticulos.Rows = 2
17                        For lngCol = intColManejo1 To grdArticulos.Cols - 1
18                            grdArticulos.Col = lngCol
19                            grdArticulos.TextMatrix(1, lngCol) = ""
20                            grdArticulos.CellForeColor = vbBlack
21                        Next lngCol
22                        pSumaCostos
23                    Else
24                        fblnValido = False
25                        If llngCveCajaMaterialRequi <> -1 Then
26                            cboCajaMaterial.ListIndex = fintLocalizaCbo_new(cboCajaMaterial, CStr(llngCveCajaMaterialRequi))
27                        Else
28                            cboCajaMaterial.ListIndex = 0
29                        End If
30                    End If
31                Else
32                    llngCveCajaMaterialRequi = cboCajaMaterial.ItemData(cboCajaMaterial.ListIndex)
33                End If
34            End If
35        ElseIf llngCveCajaMaterialRequi <> -1 Then
36            If MsgBox(SIHOMsg("1551"), (vbYesNo + vbQuestion), "Mensaje") = vbYes Then
37                llngCveCajaMaterialRequi = -1
38                grdArticulos.Rows = 2
39                For lngCol = intColManejo1 To grdArticulos.Cols - 1
40                    grdArticulos.Col = lngCol
41                    grdArticulos.TextMatrix(1, lngCol) = ""
42                    grdArticulos.CellForeColor = vbBlack
43                Next lngCol
44                pSumaCostos
45            Else
46                fblnValido = False
47                cboCajaMaterial.ListIndex = fintLocalizaCbo_new(cboCajaMaterial, CStr(llngCveCajaMaterialRequi))
48            End If
49        Else
50            llngCveCajaMaterialRequi = -1
51        End If
          
52        If llngCveCajaMaterialRequi = -1 Then
53            If cboNombreComercial.ListIndex = -1 Then
54                fblnValido = False
                  '¡Dato no válido, seleccione un valor de la lista!
55                MsgBox SIHOMsg(3), vbOKOnly + vbInformation, "Mensaje"
56                cboNombreComercial.SetFocus
57            End If
58            If fblnValido And Val(txtCantidadSolicitada.Text) = 0 And chkPedirMaximo.Value = 0 Then
59                fblnValido = False
                  'Debe indicar la forma en que va a solicitar.
60                'MsgBox SIHOMsg(898), vbOKOnly + vbInformation, "Mensaje"
                   MsgBox "La cantidad solicitada debe ser mayor que cero.", vbOKOnly + vbInformation, "Mensaje"
61            End If
62        End If

63    Exit Function
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnValido" & " Linea:" & Erl()))
        Unload Me
End Function

Private Sub cmdAgregar_Click()
1     On Error GoTo NotificaError
      Dim rs As New ADODB.Recordset
      Dim lngCantidadFaltante As Long
      Dim rsFaltante As New ADODB.Recordset
      Dim lngTotalAgregados As Long
      Dim dblAumento As Double
      Dim strArticulos As String
      Dim lngExistUnidadMinima As Long
      Dim lngReorden As Long
      Dim lngResiduo As Long
      Dim lstrTotalLocalizaciones As String
      Dim strTodosManejos As String
      Dim rsCaja As New ADODB.Recordset
      Dim lngCol As Long

         pEstatus
3         If fblnValido() Then
              'Validamos si estan usando presupuestos
4             If cboTipoRequisicion.Text = "SALIDA A DEPARTAMENTO" Then
5                 If Not fblnValidarSurtido() Then
6                     Exit Sub
7                 End If
8             End If
          
9             If llngCveCajaMaterialRequi > -1 Then
                  'Se seleccionó una caja de material
10                grdArticulos.Rows = 2
11                If Trim(grdArticulos.TextMatrix(1, intColCveArticulo)) <> "" Then
12                    For lngCol = intColManejo1 To grdArticulos.Cols - 1
13                        grdArticulos.Col = lngCol
14                        grdArticulos.TextMatrix(1, lngCol) = ""
15                        grdArticulos.CellForeColor = vbBlack
16                    Next lngCol
17                    pSumaCostos
18                End If
                              
19                Set rsCaja = frsEjecuta_SP(cboCajaMaterial.ItemData(cboCajaMaterial.ListIndex) & "|" & cboAlmacenSurte.ItemData(cboAlmacenSurte.ListIndex), "SP_IVSELCAJAMATERIAL")
20                If rsCaja.RecordCount <> 0 Then
21                        dblAumento = 100 / rsCaja.RecordCount
22                        pgbBarra.Value = 0
23                        fraBarra.Visible = True
          
24                        lngTotalAgregados = 0
              
25                        Do While Not rsCaja.EOF
26                            fraBarra.Refresh
27                            pgbBarra.Value = IIf(pgbBarra.Value + dblAumento > 100, 100, pgbBarra.Value + dblAumento)
28                            If (rsCaja!Controlado = 1 And vlblnControlado) Or rsCaja!Controlado = 0 Then
29                                pAgrega rsCaja!cveArticulo, rsCaja!descArticulo, rsCaja!cantidad, rsCaja!descUnidad, IIf(rsCaja!Unidad = "A", True, False), rsCaja!existencia, False
30                            Else
                                  'Concatenar los artículos controlados:
31                                strArticulos = strArticulos & rsCaja!descArticulo & Chr(13)
32                            End If
33                            rsCaja.MoveNext
34                        Loop
                      
35                        fraBarra.Visible = False
                          
36                        If Trim(strArticulos) <> "" Then
                              'No está autorizado para solicitar medicamento controlado.
37                            MsgBox SIHOMsg(585) & Chr(13) & strArticulos, vbOKOnly + vbExclamation, "Mensaje"
38                        End If
                          
39                        grdArticulos.TopRow = grdArticulos.Rows - 1
40                        grdArticulos.Row = grdArticulos.Rows - 1
41                        txtnombrecompleto.Text = grdArticulos.TextMatrix(grdArticulos.Row, intColNombreComercial)
                          
42                        If fblnCanFocus(cmdGrabar) Then cmdGrabar.SetFocus
43                Else
                      'No existe información con esos parámetros.
44                    MsgBox SIHOMsg(13), vbOKOnly + vbInformation, "Mensaje"
45                End If
46            Else
47                If chkPedirMaximo.Value = 0 Then
                      ' Pedido manual:
                      'Si es medicamento controlado y se tiene permiso o no es controlado:
48                    If (lintControlado = 1 And vlblnControlado) Or lintControlado = 0 Then
49                        pAgrega txtClave.Text, cboNombreComercial.Text, Val(txtCantidadSolicitada.Text), IIf(OptAlterna.Value, lstrUnidadAlterna, lstrUnidadMinima), OptAlterna.Value, txtExistencia.Text, OptAlterna.Value
                          
50                        cboNombreComercial.ListIndex = 0
51                        cboNombreComercial_Click
                          
52                    Else
                          'No está autorizado para solicitar medicamento controlado.
53                        MsgBox SIHOMsg(585) & Chr(13) & strArticulos, vbOKOnly + vbExclamation, "Mensaje"
54                    End If
                      
55                Else
                      
                      'Manejos
56                    strTodosManejos = "_"
57                    Set rs = frsEjecuta_SP("-1|1|-1", "Sp_IvSelManejos")
58                    If rs.RecordCount > 0 Then
                          
59                        Do While Not rs.EOF
60                            If ((vlblnControlado And rs!bitControlado = 1) Or rs!bitControlado = 0) Then
61                                strTodosManejos = strTodosManejos & CStr(rs!intCveManejo) & "_"
62                            End If
63                            rs.MoveNext
64                        Loop
                          
65                    End If
66                    rs.Close
67                    strTodosManejos = strTodosManejos & "0_"
                      
                      'Pedir con base en máximos y mínimos:
68                    lstrTotalLocalizaciones = ""
69                    vgstrParametrosSP = CStr(cboDepartamentoSolicita.ItemData(cboDepartamentoSolicita.ListIndex)) & "|" & _
                      CStr(IIf(cboNombreComercial.ItemData(cboNombreComercial.ListIndex) <> 0, 1, IIf(optOpcion(0).Value, 1, 0))) & "|" & _
                      CStr(IIf(cboNombreComercial.ItemData(cboNombreComercial.ListIndex) <> 0, "-1", IIf(optOpcion(0).Value, "-1", IIf(optOpcion(1).Value, "0", IIf(optOpcion(2).Value, "1", "2"))))) & "|" & _
                      CStr(IIf(cboNombreComercial.ItemData(cboNombreComercial.ListIndex) <> 0, 1, IIf(cboFamilia.ItemData(cboFamilia.ListIndex) = 0, 1, 0))) & "|" & _
                      CStr(IIf(cboNombreComercial.ItemData(cboNombreComercial.ListIndex) <> 0, "-1", IIf(cboFamilia.ItemData(cboFamilia.ListIndex) = 0, "-1", fstrFormateaComoFamilia(Str(cboFamilia.ItemData(cboFamilia.ListIndex)))))) & "|" & _
                      CStr(IIf(cboNombreComercial.ItemData(cboNombreComercial.ListIndex) <> 0, 1, IIf(cboSubfamilia.ItemData(cboSubfamilia.ListIndex) = 0, 1, 0))) & "|" & _
                      CStr(IIf(cboNombreComercial.ItemData(cboNombreComercial.ListIndex) <> 0, "-1", IIf(cboSubfamilia.ItemData(cboSubfamilia.ListIndex) = 0, "-1", fstrFormateaComoSubFamilia(Str(cboSubfamilia.ItemData(cboSubfamilia.ListIndex)))))) & "|" & _
                      CStr(IIf(cboNombreComercial.ItemData(cboNombreComercial.ListIndex) = 0, 1, 0)) & "|" & _
                      CStr(IIf(cboNombreComercial.ItemData(cboNombreComercial.ListIndex) = 0, "", fstrObtenCveArticulo(cboNombreComercial.ItemData(cboNombreComercial.ListIndex)))) & "|" & _
                      "X" & "|" & CStr(2) & "|" & CStr(-1) & "|" & _
                      IIf(lstrTotalLocalizaciones = "", IIf(cboLocalizacion.ListIndex = 0, -1, "_" & cboLocalizacion.ItemData(cboLocalizacion.ListIndex) & "_"), Trim(lstrTotalLocalizaciones)) & "|" & _
                      IIf(Trim(cboTipoRequisicion.List(cboTipoRequisicion.ListIndex)) = "COMPRA - PEDIDO", "C", "R") & "|" & _
                      "1|" & strTodosManejos
                      
70                    Set rs = frsEjecuta_SP(vgstrParametrosSP, "SP_IvSelArticulosUbicacion")
                      
71                    If rs.RecordCount <> 0 Then
          
                          '*- Inicio de barra
72                        dblAumento = 100 / rs.RecordCount
73                        pgbBarra.Value = 0
74                        fraBarra.Visible = True
          
75                        lngTotalAgregados = 0
              
76                        Do While Not rs.EOF
                          
                              '*- Barra de avance
77                            fraBarra.Refresh
78                            pgbBarra.Value = IIf(pgbBarra.Value + dblAumento > 100, 100, pgbBarra.Value + dblAumento)
                          
                              'Si artículo tiene asignados máximos y mínimos:
79                            If rs!IdArticuloMaximo <> -1 Then
                              
80                               If ((Trim(cboTipoRequisicion.List(cboTipoRequisicion.ListIndex)) = "SALIDA A DEPARTAMENTO" And rs!CostoGasto = "G") Or _
                                  (Trim(cboTipoRequisicion.List(cboTipoRequisicion.ListIndex)) = "REUBICACION" And rs!CostoGasto = "C") Or _
                                  Trim(cboTipoRequisicion.List(cboTipoRequisicion.ListIndex)) = "COMPRA - PEDIDO" Or _
                                  Trim(cboTipoRequisicion.List(cboTipoRequisicion.ListIndex)) = "ALMACEN ABASTECIMIENTO" Or _
                                  rs!CostoGasto = "A") Then
                                      
                                      'Existencia en unidades mínimas, en alterna si el artículo se maneja en una sola unidad
81                                    lngExistUnidadMinima = rs!ExistenciaUV * rs!Contenido + rs!ExistenciaUM
                                      'Punto de reorden definido en unidades mínimas, en alterna si el artículo se maneja en una sola unidad
82                                    lngReorden = rs!PuntoReorden * IIf(rs!INTUNIDADPUNTOREORDEN = 1, rs!Contenido, 1)
                                  
83                                    If lngExistUnidadMinima <= lngReorden Then
                                          'Si el artículo está en el mínimo o por debajo:
                                  
                                          'Determinar el faltante
84                                        lngCantidadFaltante = 1
85                                        vgstrParametrosSP = rs!clave & "|" & Str(cboDepartamentoSolicita.ItemData(cboDepartamentoSolicita.ListIndex))
86                                        Set rsFaltante = frsEjecuta_SP(vgstrParametrosSP, "Sp_Ivintcantidadfaltante", True, lngCantidadFaltante)
                                          
87                                        If lngCantidadFaltante > 0 Then
                                          
                                              'Si es medicamento controlado y se tiene permiso o no es controlado
                                              'Validar el tipo de artículo por tipo de requisición:
88                                            If (rs!Controlado = 1 And vlblnControlado) Or rs!Controlado = 0 Then
                                          
89                                                If Trim(cboTipoRequisicion.List(cboTipoRequisicion.ListIndex)) = "COMPRA - PEDIDO" And rs!Contenido <> 1 Then
                                                      'Convertir a unidades completas el faltante:
90                                                    lngCantidadFaltante = Int(lngCantidadFaltante / rs!Contenido)
91                                                    If lngCantidadFaltante <> 0 Then
                                                          'Agregar a la requisición en unidad alterna:
92                                                        pAgrega rs!clave, rs!NombreComercial, lngCantidadFaltante, rs!UnidadAlterna, True, Int(rs!ExistenciaUV + (rs!ExistenciaUM / rs!Contenido)), True
93                                                        lngTotalAgregados = lngTotalAgregados + 1
94                                                    End If
95                                                Else
96                                                    If rs!Contenido = 1 Then
                                                          'Agregar a la requisición: en unidad alterna:
97                                                        pAgrega rs!clave, rs!NombreComercial, lngCantidadFaltante, rs!UnidadAlterna, True, Int(rs!ExistenciaUV + (rs!ExistenciaUM / rs!Contenido)), True
98                                                        lngTotalAgregados = lngTotalAgregados + 1
99                                                    Else
                                                      
100                                                       lngResiduo = lngCantidadFaltante Mod rs!Contenido
                                      
101                                                       If lngResiduo = 0 Then
                                                              'Se solicita en unidad alterna:
                                                              'Agregar a la requisición: en unidad mínima:
102                                                           pAgrega rs!clave, rs!NombreComercial, lngCantidadFaltante / rs!Contenido, rs!UnidadAlterna, True, Int(rs!ExistenciaUV + (rs!ExistenciaUM / rs!Contenido)), True
103                                                           lngTotalAgregados = lngTotalAgregados + 1
104                                                       Else
                                                              'Se solicita en unidad mínima:
105                                                           pAgrega rs!clave, rs!NombreComercial, lngCantidadFaltante, rs!UnidadMinima, False, Int(rs!ExistenciaUV + (rs!ExistenciaUM / rs!Contenido)), False
106                                                           lngTotalAgregados = lngTotalAgregados + 1
107                                                       End If
108                                                   End If
109                                               End If
110                                           Else
                                                  'Concatenar los artículos controlados:
111                                               strArticulos = strArticulos & rs!NombreComercial & Chr(13)
112                                           End If
113                                       End If
114                                   End If
115                               End If
116                           End If
117                           rs.MoveNext
118                       Loop
                      
119                       fraBarra.Visible = False
                      
120                       If Trim(strArticulos) <> "" Then
                              'No está autorizado para solicitar medicamento controlado.
121                           MsgBox SIHOMsg(585) & Chr(13) & strArticulos, vbOKOnly + vbExclamation, "Mensaje"
122                       End If
123                       If lngTotalAgregados = 0 Then
                              'No se detectaron faltantes con respecto a máximos.
124                           MsgBox SIHOMsg(497), vbOKOnly + vbInformation, "Mensaje"
125                       End If
126                   Else
127                       If cboNombreComercial.ItemData(cboNombreComercial.ListIndex) = 0 Then
                              'No existe información con esos parámetros.
128                           MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje"
129                       Else
                              'Ese artículo no está reubicado a este departamento.
130                           MsgBox SIHOMsg(318), vbOKOnly + vbInformation, "Mensaje"
131                       End If
132                   End If
133               End If
                  
134               grdArticulos.TopRow = grdArticulos.Rows - 1
135               grdArticulos.Row = grdArticulos.Rows - 1
136               txtnombrecompleto.Text = grdArticulos.TextMatrix(grdArticulos.Row, intColNombreComercial)
137               cboNombreComercial.SetFocus
138           End If
139       End If

140   Exit Sub
NotificaError:
       Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdAgregar_Click" & " Linea:" & Erl()))
       Unload Me
End Sub

Private Sub pAgrega(vlstrCveArticulo As String, _
                    vlstrNombreComercial As String, _
                    vllngCantidadSolicitada As Long, _
                    vlstrUnidad As String, _
                    vlblnAlterna As Boolean, _
                    vllngExistencia As String, _
                    vlblnPregunta As Boolean)
1     On Error GoTo NotificaError
      Dim vllngRenglon As Long
      Dim lngContador As Long
      Dim vlblnEncontro As Boolean
      Dim sSQL As String
      Dim rsTemp As New ADODB.Recordset
          
2         If Trim(grdArticulos.TextMatrix(1, intColCveArticulo)) = "" Then
3             vllngRenglon = 1
4         Else
5             vlblnEncontro = False
6             lngContador = 1
7             Do While lngContador <= grdArticulos.Rows - 1 And Not vlblnEncontro
8                 If Trim(grdArticulos.TextMatrix(lngContador, intColCveArticulo)) = Trim(vlstrCveArticulo) Then
9                     vlblnEncontro = True
10                    vllngRenglon = lngContador
11                End If
12                lngContador = lngContador + 1
13            Loop
              
14            If Not vlblnEncontro Then
15                vllngRenglon = grdArticulos.Rows
16                grdArticulos.Rows = grdArticulos.Rows + 1
17            Else
18                If Trim(grdArticulos.TextMatrix(grdArticulos.Row, intColEstado)) = "PENDIENTE" Or Trim(grdArticulos.TextMatrix(grdArticulos.Row, intColEstado)) = "CANCELADA" Then
19                    If vlblnPregunta Then
                          'El artículo ya está incluído. ¿Desea actualizar los datos?
20                        If MsgBox(SIHOMsg(495), vbYesNo + vbQuestion, "Mensaje") = vbNo Then
21                            vllngRenglon = 0
22                        End If
23                    End If
24                Else
25                    vllngRenglon = 0
26                End If
27            End If
28        End If
          
29        If vllngRenglon <> 0 Then

30            If chkIncluir.Value = 1 Then
                  
31                vgstrParametrosSP = vlstrCveArticulo
32                Set rsTemp = frsEjecuta_SP(vgstrParametrosSP, "SP_IVSELULTIMACOMPRAARTICULO")
                  
33                If rsTemp.RecordCount <> 0 Then
                  
34                    With grdArticulos
35                      .TextMatrix(vllngRenglon, intColProveedor) = rsTemp!NombreComercial
36                      .TextMatrix(vllngRenglon, intColFecha) = rsTemp!fecha
37                      .TextMatrix(vllngRenglon, intColCantidadProveedor) = rsTemp!cantidad
38                      .TextMatrix(vllngRenglon, intColCosto) = FormatCurrency(rsTemp!Costo, 4)
39                    End With
40                End If
41                rsTemp.Close
                  
42            End If
              
43            If Not vlblnConsulta Then
44                If llngCveCajaMaterialRequi <> -1 Then
45                    cmdBorrar.Enabled = False
46                Else
47                    cmdBorrar.Enabled = True
48                End If
49            End If
              
              'Manejos
50            Set rsTemp = frsEjecuta_SP(cboNombreComercial.ItemData(cboNombreComercial.ListIndex), "Sp_IvSelArticuloManejos")
51            With rsTemp
52                intcontador = intColManejo4
53                Do While Not .EOF

54                    If intcontador >= intColManejo1 Then
55                        If Not IsNull(!vchSimbolo) Then
                          
56                            grdArticulos.Col = intcontador
57                            grdArticulos.Row = grdArticulos.Rows - 1
58                            grdArticulos.CellFontName = "Wingdings"
59                            grdArticulos.CellFontSize = 12
60                            grdArticulos.CellForeColor = CLng(!vchColor)
61                            grdArticulos.TextMatrix(grdArticulos.Row, intcontador) = !vchSimbolo
62                            intcontador = intcontador - 1
                              
63                        End If
64                    End If
65                    .MoveNext
                      
66                Loop
                  
67                If .RecordCount > intMaxManejos Then
68                    intMaxManejos = IIf(.RecordCount > 4, 4, .RecordCount)
69                End If
70                .Close
                  
71            End With

72            grdArticulos.TextMatrix(vllngRenglon, intColCveArticulo) = vlstrCveArticulo
73            grdArticulos.TextMatrix(vllngRenglon, intColNombreComercial) = vlstrNombreComercial
74            grdArticulos.TextMatrix(vllngRenglon, intColCantidad) = Str(vllngCantidadSolicitada)
75            grdArticulos.TextMatrix(vllngRenglon, intColDescripcionUnidad) = vlstrUnidad
76            grdArticulos.TextMatrix(vllngRenglon, intColExistencia) = vllngExistencia
77            grdArticulos.TextMatrix(vllngRenglon, intColEstado) = "PENDIENTE"
78            grdArticulos.TextMatrix(vllngRenglon, intColUnidad) = vlblnAlterna

79            grdArticulos.ColWidth(intColManejo1) = IIf(intMaxManejos = 4, 300, 0)
80            grdArticulos.ColWidth(intColManejo2) = IIf(intMaxManejos >= 3, 300, 0)
81            grdArticulos.ColWidth(intColManejo3) = IIf(intMaxManejos >= 2, 300, 0)
82            grdArticulos.ColWidth(intColManejo4) = IIf(intMaxManejos >= 1, 300, 0)
              
83            pSumaCostos
84            pHabilita False, False, False, False, False, True, False
          
85        End If

86    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pAgrega" & " Linea:" & Erl()))
        Unload Me
End Sub

Private Sub pSumaCostos()
Dim X As Integer
Dim vldbltotal As Double
  
  With grdArticulos
    If chkIncluir.Value = 1 Then
      For X = 1 To .Rows - 1
        vldbltotal = vldbltotal + (dVal(.TextMatrix(X, intColCantidad)) * dVal(.TextMatrix(X, intColCosto)))
      Next X
      txtTotal = FormatCurrency(vldbltotal, 4)
      If dVal(txtTotal) > 0 Then
        fraTotales.Visible = True
      Else
        fraTotales.Visible = False
      End If
    Else
      fraTotales.Visible = False
    End If
  End With

End Sub

Private Sub cmdAgregarNuevo_Click()
1     On Error GoTo NotificaError
          
2         If Trim(grdArticulos.TextMatrix(1, intColCveArticulo)) <> "" Then grdArticulos.Rows = grdArticulos.Rows + 1
          
3         txtCaptura.Text = ""
4         grdArticulos.Row = grdArticulos.Rows - 1
          
5         grdArticulos.TextMatrix(grdArticulos.Row, intColCveArticulo) = "NUEVO"
6         grdArticulos.TextMatrix(grdArticulos.Row, intColEstado) = "PENDIENTE"
7         grdArticulos.TextMatrix(grdArticulos.Row, intColUnidad) = True
          
8         For intcontador = intColCveArticulo To grdArticulos.Cols - 1
9             grdArticulos.Col = intcontador
10            grdArticulos.CellForeColor = &HC0C000
11        Next intcontador
          
12        If Not vlblnConsulta Then cmdBorrar.Enabled = True
          
13        grdArticulos.SetFocus
              
14        Exit Sub
          
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdAgregarNuevo_Click" & " Linea:" & Erl()))
        Unload Me
End Sub

Private Sub cmdAnterior_Click()
1     On Error GoTo NotificaError
          
2         If vlblnBusqueda Then
3             If grdRequisiciones.Row <> 1 Then
4                 grdRequisiciones.Row = grdRequisiciones.Row - 1
5             End If
6             If fintCargaReq(grdRequisiciones.TextMatrix(grdRequisiciones.Row, 1)) <> 0 Then
7                 pMuestra
8             End If
9         Else
10            rsIvRequisicionMaestro.MovePrevious
11            If rsIvRequisicionMaestro.BOF Then
12                rsIvRequisicionMaestro.MoveNext
13            End If
14            pMuestra
15        End If

16    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdAnterior_Click" & " Linea:" & Erl()))
        Unload Me
End Sub

Private Sub cmdBorrar_Click()
1     On Error GoTo NotificaError
      Dim X As Integer
      Dim intcontador As Integer
      Dim intContador2 As Integer
      Dim intColManejoInicial As Integer
      Dim intNewMaxManejo As Integer
      Dim intRenglon As Integer
      Dim lngColor As Long
          
2         If grdArticulos.Rows = 2 Then
          
3             For X = intColManejo1 To grdArticulos.Cols - 1
4                 grdArticulos.Col = X
5                 grdArticulos.TextMatrix(1, X) = ""
6                 grdArticulos.CellForeColor = vbBlack
7             Next X
8             cmdBorrar.Enabled = False
9             pSumaCostos
              
10        Else
              
11            grdArticulos.Redraw = False
12            intRenglon = grdArticulos.Row
              
13            If intRenglon = grdArticulos.Rows - 1 Then
14                grdArticulos.Rows = grdArticulos.Rows - 1
15            Else
              
16                For X = intRenglon To grdArticulos.Rows - 2

17                    For intcontador = 1 To grdArticulos.Cols - 1
                      
18                        grdArticulos.TextMatrix(X, intcontador) = grdArticulos.TextMatrix(X + 1, intcontador)
                                              
19                        Select Case intcontador
                              Case 1, 2, 3, 4
                              
20                                If grdArticulos.TextMatrix(X, intcontador) <> "" Then
21                                    grdArticulos.Col = intcontador
22                                    grdArticulos.Row = X + 1
23                                    lngColor = grdArticulos.CellForeColor
                                      
24                                    grdArticulos.Row = X
25                                    grdArticulos.CellFontName = "Wingdings"
26                                    grdArticulos.CellFontSize = 12
27                                    grdArticulos.CellForeColor = lngColor
28                                End If
                                  
29                            Case Else
                                  
30                                grdArticulos.Col = intcontador
31                                grdArticulos.Row = X
32                                If grdArticulos.TextMatrix(X, intColCveArticulo) = "NUEVO" Then
33                                    grdArticulos.CellForeColor = &HC0C000
34                                Else
35                                     grdArticulos.CellForeColor = &H80000012
36                                End If
                              
37                        End Select
                          
38                    Next intcontador
                      
39                Next X
40                grdArticulos.Rows = grdArticulos.Rows - 1
              
41            End If
              
42            grdArticulos.Redraw = True
43            pSumaCostos
              
44        End If
          
45        If grdArticulos.Rows > 0 And intMaxManejos > 0 Then
          
46            Select Case intMaxManejos
              Case 1
47                intColManejoInicial = 4
48            Case 2
49                intColManejoInicial = 3
50            Case 3
51                intColManejoInicial = 2
52            Case 4
53                intColManejoInicial = 1
54            End Select
              
55            intNewMaxManejo = 0
          
56            For intcontador = intColManejoInicial To intColManejo4
57                For intContador2 = 1 To grdArticulos.Rows - 1
                  
58                    If grdArticulos.TextMatrix(intContador2, intcontador) <> "" Then
59                        intNewMaxManejo = intcontador
60                        intcontador = intColManejo4
61                        Exit For
62                    End If
                          
63                Next intContador2
64            Next intcontador
              
65            Select Case intNewMaxManejo
              Case 0
66                intMaxManejos = 0
67            Case 1
68                intMaxManejos = 4
69            Case 2
70                intMaxManejos = 3
71            Case 3
72                intMaxManejos = 2
73            Case 4
74                intMaxManejos = 1
75            End Select
              
76            grdArticulos.ColWidth(intColManejo1) = IIf(intMaxManejos = 4, 300, 0)
77            grdArticulos.ColWidth(intColManejo2) = IIf(intMaxManejos >= 3, 300, 0)
78            grdArticulos.ColWidth(intColManejo3) = IIf(intMaxManejos >= 2, 300, 0)
79            grdArticulos.ColWidth(intColManejo4) = IIf(intMaxManejos >= 1, 300, 0)
              
80        End If
          
81        txtnombrecompleto.Text = ""
82        If grdArticulos.Row > 0 Then
83          txtnombrecompleto.Text = grdArticulos.TextMatrix(grdArticulos.Row, intColNombreComercial)
84        End If

85    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdBorrar_Click" & " Linea:" & Erl()))
        Unload Me
End Sub

Private Sub cmdBuscar_Click()
On Error GoTo NotificaError
    
    If cboTRBusqueda.ListIndex = -1 Then
        cboTRBusqueda.ListIndex = 0
    End If
    If cboERBusqueda.ListIndex = -1 Then
        cboERBusqueda.ListIndex = 0
    End If
    sstObj.Tab = 1
    If Trim(grdRequisiciones.TextMatrix(1, 1)) = "" Then
        mskFecIni.SetFocus
    Else
        grdRequisiciones.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdBuscar_Click"))
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
On Error GoTo NotificaError
    
   ' pEstatus
        
    If Trim(grdArticulos.TextMatrix(grdArticulos.Row, intColEstado)) = "PENDIENTE" Then
        grdArticulos.TextMatrix(grdArticulos.Row, intColEstado) = "CANCELADA"
        pHabilita False, False, False, False, False, True, False
        cmdCancelar.Enabled = False
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdCancelar_Click"))
    Unload Me
End Sub

Private Sub cmdCierraReq_Click()
          Dim vllngPersonaGraba As Long
          Dim vlStrSentenciaSQL As String
          
1         On Error GoTo NotificaError
          
          
2         If fblnRevisaPermiso(vglngNumeroLogin, vllngNumeroOpcionModulo, "E", False) Then
3             vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
4             If vllngPersonaGraba <> 0 Then
              
5                 vlstrEstadoMaestroIni = Trim(txtEstatus.Text)
              
6                 EntornoSIHO.ConeccionSIHO.BeginTrans

7                 frsEjecuta_SP txtNumero.Text, "sp_ivcierrarequisicion"
             
8                 If vlstrErrorLister > 0 Then
9                     GoTo NotificaError
10                End If
                  
11                Call pGuardarLogTransaccion(Me.Name, EnmGrabar, vllngPersonaGraba, "CIERRE REQUISICION REUBICACION/SALIDA DEPARTAMENTO", txtNumero.Text)
12                EntornoSIHO.ConeccionSIHO.CommitTrans
          
13                Select Case IIf(Trim(cboTipoRequisicion.List(cboTipoRequisicion.ListIndex)) = "REUBICACION", "R", IIf(Trim(cboTipoRequisicion.List(cboTipoRequisicion.ListIndex)) = "SALIDA A DEPARTAMENTO", "D", "C"))
                      Case "C"
14                        vlStrSentenciaSQL = "Select VCHVALOR from siparametro where VCHNOMBRE = 'BITAUTORIZAREQUISICION'"
15                        Set rsAutorizaciones = frsRegresaRs(vlStrSentenciaSQL, adLockOptimistic)
                        
16                        If rsAutorizaciones.RecordCount = 0 Then
17                           pImpresionRemota "RC", CLng(txtNumero.Text), cboAlmacenSurte.ItemData(cboAlmacenSurte.ListIndex)
18                        Else
19                           If Trim(rsAutorizaciones!VCHVALOR) = "1" Then
20                              If Trim(vlstrEstadoMaestroIni) <> "PENDIENTE" Then
21                                 pImpresionRemota "RC", CLng(txtNumero.Text), cboAlmacenSurte.ItemData(cboAlmacenSurte.ListIndex)
22                              End If
23                           Else
24                              pImpresionRemota "RC", CLng(txtNumero.Text), cboAlmacenSurte.ItemData(cboAlmacenSurte.ListIndex)
25                           End If
26                        End If
27                    Case "D"
28                        Set rsAutorizaciones = frsRegresaRs("SELECT NVL(COUNT(*),0) Autoriza FROM IVAUTORIZACIONREQUISICIONES WHERE intcvedepartamento = " & cboDepartamentoSolicita.ItemData(cboDepartamentoSolicita.ListIndex) & " AND chrtiporequisicion = 'S'", adLockOptimistic, adOpenDynamic)
29                        If rsAutorizaciones!Autoriza = 0 Then
30                            pImpresionRemota "RS", CLng(txtNumero.Text), cboAlmacenSurte.ItemData(cboAlmacenSurte.ListIndex)
31                        Else
32                            If Trim(vlstrEstadoMaestroIni) <> "PENDIENTE" Then
33                                pImpresionRemota "RS", CLng(txtNumero.Text), cboAlmacenSurte.ItemData(cboAlmacenSurte.ListIndex)
34                            End If
35                        End If
36                    Case "R"
37                        Set rsAutorizaciones = frsRegresaRs("SELECT NVL(COUNT(*),0) Autoriza FROM IVAUTORIZACIONREQUISICIONES WHERE intcvedepartamento = " & cboDepartamentoSolicita.ItemData(cboDepartamentoSolicita.ListIndex) & " AND chrtiporequisicion = 'R'", adLockOptimistic, adOpenDynamic)
38                        If rsAutorizaciones!Autoriza = 0 Then
39                            pImpresionRemota "RR", CLng(txtNumero.Text), cboAlmacenSurte.ItemData(cboAlmacenSurte.ListIndex)
40                        Else
41                            If Trim(vlstrEstadoMaestroIni) <> "PENDIENTE" Then
42                                pImpresionRemota "RR", CLng(txtNumero.Text), cboAlmacenSurte.ItemData(cboAlmacenSurte.ListIndex)
43                            End If
44                        End If
45                End Select
          
46                MsgBox SIHOMsg(420), vbOKOnly + vbInformation, "Mensaje"
47                txtNumero_GotFocus '' para que se limpie toda la forma
48            End If
49        End If
          
50    Exit Sub
NotificaError:
        If vlstrErrorLister > 0 Then
            vlstrErrorLister = 0
            EntornoSIHO.ConeccionSIHO.RollbackTrans
            fintCargaReq (Str(Val(txtNumero.Text)))
            pMuestra
        Else
            EntornoSIHO.ConeccionSIHO.RollbackTrans
            Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ": cmdCierraReq_Click" & " Linea:" & Erl()))
        End If
End Sub

Private Sub cmdGrabar_Click()
1         On Error GoTo NotificaError
          Dim X As Long
          Dim vllngNumRequisicion As Long
          Dim vllngPersonaGraba As Long
          Dim rsRequisicionCompra As New ADODB.Recordset
          Dim vlobjCommand As ADODB.Command 'Para hacer el save de detalles localmente y poder manejar el error
2         Set vlobjCommand = CreateObject("ADODB.Command")
          Dim strEstatusOld As String
           ''' agregadas en caso 6598
          Dim ObjRS As New ADODB.Recordset
          Dim vlstrSentencia As String
          Dim vlstrEstadoMaestro As String
          Dim intCont As Integer
          Dim strContinuarSubrogado As String
          Dim rsSubrogado As ADODB.Recordset
          
3         If vlblnConsulta Then
4             strEstatusOld = Trim(txtEstatus.Text)
              'pEstatus no se puede usar por que cambia la informacion del datagrid articulos, si se cancela un detalle , esta funcion la vuelve a poner pendiente.
              ''''''''sustituye a pEstaus'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
5              Set ObjRS = frsRegresaRs("SELECT * FROM ivRequisicionMAestro WHERE NumNumRequisicion = " & Me.txtNumero, adLockOptimistic, adOpenDynamic)
6              If Not ObjRS.EOF Then
7                 vlstrEstadoMaestro = IIf(IsNull(ObjRS!vchEstatusRequis), "", ObjRS!vchEstatusRequis)
8                 lintSoloMaestro = 1
9              End If
10             ObjRS.Close
              ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
11            If strEstatusOld <> vlstrEstadoMaestro Then
                  'La información ha cambiado, consulte de nuevo
12                MsgBox SIHOMsg(381), vbOKOnly + vbExclamation, "Mensaje"
13                cmdBuscar.SetFocus
14                Exit Sub
15            End If
16        End If

17        If Trim(cboTipoRequisicion.List(cboTipoRequisicion.ListIndex)) = "COMPRA - PEDIDO" Then
18            If Not fblnValidaNuevo Then Exit Sub
19        End If
          
20        lstrArticuloExcede = ""
21        If lblnValidarExcedeMaximo And (Trim(cboTipoRequisicion.List(cboTipoRequisicion.ListIndex)) = "COMPRA - PEDIDO" Or Trim(cboTipoRequisicion.List(cboTipoRequisicion.ListIndex)) = "REUBICACION") Then
22            If Not fblnCantidadMaximaValida Then
23                MsgBox SIHOMsg(900) & Chr(13) & Chr(13) & lstrArticuloExcede, vbOKOnly + vbExclamation, "Mensaje"
24                Exit Sub
25            End If
26        End If
          
27        strContinuarSubrogado = ""
28        If lintInterfazFarmaciaSJP = 1 Then
29            If Trim(cboTipoRequisicion.List(cboTipoRequisicion.ListIndex)) = "REUBICACION" Or Trim(cboTipoRequisicion.List(cboTipoRequisicion.ListIndex)) = "SALIDA A DEPARTAMENTO" Then
30                If llngDeptoSubrogado = cboAlmacenSurte.ItemData(cboAlmacenSurte.ListIndex) Then
31                    For intCont = 1 To grdArticulos.Rows - 1
32                        If Trim(grdArticulos.TextMatrix(intCont, intColEstado)) = "PENDIENTE" Then
33                            Set rsSubrogado = frsRegresaRs("select count(*) total from IvArticulo inner join IvArticulosSubrogados on IVARTICULO.INTIDARTICULO = IVARTICULOSSUBROGADOS.INTIDARTICULO " & _
                                              " where IVARTICULOSSUBROGADOS.VCHCVEARTICULOEXT is not null and IVARTICULO.CHRCVEARTICULO = '" & Trim(grdArticulos.TextMatrix(intCont, intColCveArticulo)) & "'")
34                            If rsSubrogado!Total = 0 Then
35                                strContinuarSubrogado = strContinuarSubrogado & Trim(grdArticulos.TextMatrix(intCont, intColCveArticulo)) & "  " & Trim(grdArticulos.TextMatrix(intCont, intColNombreComercial)) & Chr(13)
36                            End If
37                        End If
38                    Next intCont
39                End If
40            End If
41        End If
42        If strContinuarSubrogado <> "" Then
43            MsgBox "Los siguientes artículos no tienen clave asignada de la farmacia subrogada. " & Chr(13) & "No es posible realizar la requisición. " & Chr(13) & Chr(13) & strContinuarSubrogado, vbCritical, "Mensaje"
44            Exit Sub
45        End If
          
46        If fblnRevisaPermiso(vglngNumeroLogin, vllngNumeroOpcionModulo, "E", False) Then
47            If Trim(grdArticulos.TextMatrix(1, intColCveArticulo)) <> "" Then
                  
48                vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
              
49                If vllngPersonaGraba <> 0 Then
          
50                    EntornoSIHO.ConeccionSIHO.BeginTrans
                  
51                    With rsIvRequisicionMaestro
52                        If Not vlblnConsulta Then
53                            .AddNew
54                            !smiCveDeptoRequis = cboDepartamentoSolicita.ItemData(cboDepartamentoSolicita.ListIndex)
55                            !intCveEmpleaRequis = vllngPersonaGraba
56                            !smiCveDeptoAlmacen = cboAlmacenSurte.ItemData(cboAlmacenSurte.ListIndex)
57                            !dtmFechaRequisicion = fdtmServerFecha
58                            !DTMHORAREQUISICION = fdtmServerHora
59                            !vchEstatusRequis = "PENDIENTE"
60                            !chrDestino = IIf(Trim(cboTipoRequisicion.List(cboTipoRequisicion.ListIndex)) = "REUBICACION", "R", IIf(Trim(cboTipoRequisicion.List(cboTipoRequisicion.ListIndex)) = "SALIDA A DEPARTAMENTO", "D", IIf(Trim(cboTipoRequisicion.List(cboTipoRequisicion.ListIndex)) = "ALMACEN ABASTECIMIENTO", "A", "C")))
61                            !numNumCuenta = 0
62                            !bitUrgente = IIf(chkUrgente.Value = 1, 1, 0)
63                            !NUMNUMREQUISREL = 0
64                            !intNumeroLogin = vglngNumeroLogin
65                            !VCHOBSERVACIONES = Trim(txtObservaciones.Text)
66                            !smiCveDeptoGenera = vgintNumeroDepartamento
67                            !bitCompraDirecta = chkCompradirecta.Value
68                            If lblnSeleccionarProveedor And chkCompradirecta.Value = 1 And cboProveedor.ListIndex > 0 Then !intcveproveedor = cboProveedor.ItemData(cboProveedor.ListIndex)
69                            If llngCveCajaMaterialRequi <> -1 Then
70                                !intCajaMaterial = llngCveCajaMaterialRequi
71                            End If
72                            .Update
73                            txtNumero.Text = CStr(flngObtieneIdentity("SEC_IvRequisicionMaestro", rsIvRequisicionMaestro!numnumRequisicion))
74                            fintCargaReq (txtNumero.Text)
75                            lintSoloMaestro = 0
76                        Else
77                            If fintCargaReq(Str(Val(txtNumero.Text))) <> 0 Then
78                                !bitUrgente = IIf(chkUrgente.Value = 1, 1, 0)
79                                !intCveEmpleaRequis = vllngPersonaGraba
80                                If txtObservaciones.Enabled Then !VCHOBSERVACIONES = Trim(txtObservaciones.Text)
81                                .Update
82                            End If
83                        End If
84                    End With
                  
                      'If lintSoloMaestro = 0 Then ' se quita para que simpre que se encuentre una consulta y hay modificaciones se haga la actualizacion
                      'If vlblnConsulta Then ' solo borra los registros si la consulta esta activa, de lo contrario
                      '    vlstrx = "Delete FROM IvRequisicionDetalle where numNumRequisicion=" & txtNumero.Text
                      '    pEjecutaSentencia vlstrx
                      'End If ' se insertan los registros
85                        vlstrx = "DELETE FROM IvRequisicionDetalle WHERE numNumRequisicion = " & Trim(txtNumero.Text) & " AND CHRCVEARTICULO = 'NUEVO'"
                          'Metodo pEjecutaSentencia de modProcedimientos, para capturar el error localmente
86                        With vlobjCommand
87                          Set .ActiveConnection = EntornoSIHO.ConeccionSIHO
88                          .CommandText = vlstrx
89                          .Execute
90                        End With
                      
91                        With rsIvRequisicionDetalle
                          
92                            For X = 1 To grdArticulos.Rows - 1
                              
                                  Dim rsRequisicionDet As ADODB.Recordset 'Para validar los registros existentes en la tabla de detalles
93                                Set rsRequisicionDet = New ADODB.Recordset
                                  'If Trim(grdArticulos.TextMatrix(X, intColCveArticulo)) <> "NUEVO" Then
                                  
94                                vlstrx = "Select * From IvRequisicionDetalle Where numNumRequisicion = " & Trim(txtNumero.Text) & " AND CHRCVEARTICULO = '" & Trim(grdArticulos.TextMatrix(X, intColCveArticulo)) & "'"
95                                Set rsRequisicionDet = frsRegresaRs(vlstrx, adLockOptimistic, adOpenDynamic)
                                      
                                  'Else
                                  'vlstrx = "Select * From IvRequisicionDetalle Where numNumRequisicion = 1 AND CHRCVEARTICULO = '1'"
                                  'Set rsRequisicionDet = frsRegresaRs(vlstrx, adLockOptimistic, adOpenDynamic)
                                  'End If
                                  
96                                If rsRequisicionDet.RecordCount > 0 Then
97                                    If Trim(grdArticulos.TextMatrix(X, intColCveArticulo)) <> "NUEVO" Then
98                                        vlstrx = "UPDATE IvRequisicionDetalle SET " & _
                                          "IntCantidadSolicitada = " & Val(grdArticulos.TextMatrix(X, intColCantidad)) & ", " & _
                                          "CHRUNIDADCONTROL = '" & IIf(CBool(grdArticulos.TextMatrix(X, intColUnidad)), "A", "M") & "', " & _
                                          "vchEstatusDetRequis = '" & Trim(grdArticulos.TextMatrix(X, intColEstado)) & "'"
99                                        If Trim(grdArticulos.TextMatrix(X, intColCveArticulo)) = "NUEVO" Then
100                                           vlstrx = vlstrx & ", vchNombreArticuloNuevo = '" & Trim(grdArticulos.TextMatrix(X, intColNombreComercial)) & "', "
101                                           vlstrx = vlstrx & "VCHUNIDADARTICULONUEVO = '" & Trim(grdArticulos.TextMatrix(X, intColDescripcionUnidad)) & "'"
102                                       End If
103                                       vlstrx = vlstrx & " WHERE numNumRequisicion = " & Trim(txtNumero.Text) & " And CHRCVEARTICULO = '" & Trim(grdArticulos.TextMatrix(X, intColCveArticulo)) & "'"
                                          'Metodo pEjecutaSentencia de modProcedimientos, para capturar el error localmente
104                                       With vlobjCommand
105                                         Set .ActiveConnection = EntornoSIHO.ConeccionSIHO
106                                         .CommandText = vlstrx
107                                         .Execute
108                                       End With
109                                   Else
110                                       .AddNew
111                                       !numnumRequisicion = CDbl(txtNumero.Text)
112                                       !chrcvearticulo = Trim(grdArticulos.TextMatrix(X, intColCveArticulo))
113                                       !IntCantidadSolicitada = Val(grdArticulos.TextMatrix(X, intColCantidad))
114                                       !CHRUNIDADCONTROL = IIf(CBool(grdArticulos.TextMatrix(X, intColUnidad)), "A", "M")
115                                       !vchEstatusDetRequis = Trim(grdArticulos.TextMatrix(X, intColEstado))
116                                       If Trim(grdArticulos.TextMatrix(X, intColCveArticulo)) = "NUEVO" Then
117                                           !vchNombreArticuloNuevo = Trim(grdArticulos.TextMatrix(X, intColNombreComercial))
118                                           !VCHUNIDADARTICULONUEVO = Trim(grdArticulos.TextMatrix(X, intColDescripcionUnidad))
119                                       End If
120                                       .Update
121                                   End If
122                               Else
123                                   .AddNew
124                                   !numnumRequisicion = CDbl(txtNumero.Text)
125                                   !chrcvearticulo = Trim(grdArticulos.TextMatrix(X, intColCveArticulo))
126                                   !IntCantidadSolicitada = Val(grdArticulos.TextMatrix(X, intColCantidad))
127                                   !CHRUNIDADCONTROL = IIf(CBool(grdArticulos.TextMatrix(X, intColUnidad)), "A", "M")
128                                   !vchEstatusDetRequis = Trim(grdArticulos.TextMatrix(X, intColEstado))
129                                   If Trim(grdArticulos.TextMatrix(X, intColCveArticulo)) = "NUEVO" Then
130                                       !vchNombreArticuloNuevo = Trim(grdArticulos.TextMatrix(X, intColNombreComercial))
131                                       !VCHUNIDADARTICULONUEVO = Trim(grdArticulos.TextMatrix(X, intColDescripcionUnidad))
132                                   End If
133                                   .Update
134                               End If
                                  
135                           Next X
136                       End With
                      ''End If''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                  
                      'Si es Compra-Pedido graba el departamento que autoriza y el que recibe
137                   If Not vlblnConsulta Then
138                       If Trim(cboTipoRequisicion.List(cboTipoRequisicion.ListIndex)) = "COMPRA - PEDIDO" Then
139                           vlstrx = "Select * From IvRequisicionCompra Where numNumRequisicion = -1"
140                           Set rsRequisicionCompra = frsRegresaRs(vlstrx, adLockOptimistic, adOpenDynamic)
141                           rsRequisicionCompra.AddNew
142                           rsRequisicionCompra!numnumRequisicion = CDbl(txtNumero.Text)
143                           If lintAutorizaRequi = 1 Then rsRequisicionCompra!intCveDeptoAutoriza = llngDeptoAutoriza
144                           rsRequisicionCompra!intCveDeptoRecibe = llngDeptoRecibe
145                           rsRequisicionCompra.Update
146                       End If
147                   End If
                  
148                   vlstrEstadoMaestroIni = Trim(txtEstatus.Text)   'Estado antes del cambio
149                   frsEjecuta_SP txtNumero.Text, "Sp_IvUpdEstatusReqMaestro"
                      
150                   vllngNumRequisicion = CLng(txtNumero.Text)
151                   If Not vlblnConsulta Then
152                       Call pGuardarLogTransaccion(Me.Name, EnmGrabar, vllngPersonaGraba, "REQUISICION DE ARTICULOS", txtNumero.Text)
153                   Else
154                       Call pGuardarLogTransaccion(Me.Name, EnmCambiar, vllngPersonaGraba, "REQUISICION DE ARTICULOS", txtNumero.Text)
155                   End If
156                   vlblnConsulta = True
157                   cmdBorrar.Enabled = False
                      'cmdCancelar.Enabled = True
158                   EntornoSIHO.ConeccionSIHO.CommitTrans
                  
159                   rsIvRequisicionMaestro.Requery
160                   vlstrEstadoMaestroFin = Trim(rsIvRequisicionMaestro!vchEstatusRequis)   'Estado despues del cambio
                  
                      '¿Desea terminar la operación?
161                   If MsgBox(SIHOMsg(496), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
162                       If Not lblnAlmacenConsigna And chkCompradirecta.Value = 0 Then
163                           Select Case IIf(Trim(cboTipoRequisicion.List(cboTipoRequisicion.ListIndex)) = "REUBICACION", "R", IIf(Trim(cboTipoRequisicion.List(cboTipoRequisicion.ListIndex)) = "SALIDA A DEPARTAMENTO", "D", IIf(Trim(cboTipoRequisicion.List(cboTipoRequisicion.ListIndex)) = "ALMACEN ABASTECIMIENTO", "A", "C")))
                                  Case "D"
164                                   Set rsAutorizaciones = frsRegresaRs("SELECT NVL(COUNT(*),0) Autoriza FROM IVAUTORIZACIONREQUISICIONES WHERE intcvedepartamento = " & cboDepartamentoSolicita.ItemData(cboDepartamentoSolicita.ListIndex) & " AND chrtiporequisicion = 'S'", adLockOptimistic, adOpenDynamic)
165                                   If rsAutorizaciones!Autoriza = 0 Then
166                                       pImpresionRemota "RS", vllngNumRequisicion, cboAlmacenSurte.ItemData(cboAlmacenSurte.ListIndex)
167                                   Else
168                                       If Trim(vlstrEstadoMaestroIni) <> "PENDIENTE" Then
169                                           pImpresionRemota "RS", vllngNumRequisicion, cboAlmacenSurte.ItemData(cboAlmacenSurte.ListIndex)
170                                       End If
171                                   End If
172                               Case "R"
173                                   Set rsAutorizaciones = frsRegresaRs("SELECT NVL(COUNT(*),0) Autoriza FROM IVAUTORIZACIONREQUISICIONES WHERE intcvedepartamento = " & cboDepartamentoSolicita.ItemData(cboDepartamentoSolicita.ListIndex) & " AND chrtiporequisicion = 'R'", adLockOptimistic, adOpenDynamic)
174                                   If rsAutorizaciones!Autoriza = 0 Then
175                                       pImpresionRemota "RR", vllngNumRequisicion, cboAlmacenSurte.ItemData(cboAlmacenSurte.ListIndex)
176                                   Else
177                                       If Trim(vlstrEstadoMaestroIni) <> "PENDIENTE" Then
178                                           pImpresionRemota "RR", vllngNumRequisicion, cboAlmacenSurte.ItemData(cboAlmacenSurte.ListIndex)
179                                       End If
180                                   End If
181                               Case "C"
182                                   lintAutorizaRequi = 0
183                                   Set rsAutorizaciones = frsSelParametros("CO", -1, "BITAUTORIZAREQUISICION")
184                                   If Not rsAutorizaciones.EOF Then
185                                       lintAutorizaRequi = IIf(rsAutorizaciones!Valor = "0", 0, 1)
186                                   End If
187                                   rsAutorizaciones.Close
                                      
188                                   If lintAutorizaRequi = 0 Then
189                                       pImpresionRemota "RC", vllngNumRequisicion, cboAlmacenSurte.ItemData(cboAlmacenSurte.ListIndex)
190                                   Else
191                                       If Trim(vlstrEstadoMaestroIni) <> "PENDIENTE" Then
192                                           pImpresionRemota "RC", vllngNumRequisicion, cboAlmacenSurte.ItemData(cboAlmacenSurte.ListIndex)
193                                       End If
194                                   End If
195                           End Select
196                       End If
                              
197                       If Trim(cboTipoRequisicion.List(cboTipoRequisicion.ListIndex)) = "COMPRA - PEDIDO" _
                              Or Trim(cboTipoRequisicion.List(cboTipoRequisicion.ListIndex)) = "ALMACEN ABASTECIMIENTO" Then
198                           If fintCargaReq(Str(vllngNumRequisicion)) <> 0 Then
199                               txtEstatus.Text = Trim(rsIvRequisicionMaestro!vchEstatusRequis)
200                               pHabilita True, True, True, True, True, False, True
201                               cmdImprimir.SetFocus
202                           Else
203                               fraCabecera.Enabled = True
204                               txtNumero.SetFocus
205                           End If
206                       Else
207                           fraCabecera.Enabled = True
208                           txtNumero.SetFocus
209                       End If
210                   Else ' no deseo terminar la operacion entonces se debe actualizar la informacion de la pantalla con los datos ya guardados
                          
211                       pMuestra
                                                       
212                       If lblnAlmacenConsigna Then
213                           fraCabecera.Enabled = False
                              
214                           fraFiltros.Enabled = False
215                           fraArticulos.Enabled = True
                              'cmdCancelar.Enabled = True
216                       Else
217                           If fblnCanFocus(txtCodigoBarras) Then txtCodigoBarras.SetFocus
218                       End If
219                   End If
220               End If
221           End If
222       End If
          
223   Exit Sub
NotificaError:
       position = InStr(Err.Description, "TicketError")

       If position > 0 Then
         position = 0
         EntornoSIHO.ConeccionSIHO.RollbackTrans
         MsgBox "No es posible hacer esta operación, ya se ha generado un ticket en la farmacia subrogada.", vbExclamation, "Mensaje"
         fintCargaReq (Str(Val(txtNumero.Text)))
         pMuestra
         Exit Sub
       Else
         Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdGrabar_Click" & " Linea:" & Erl()))
         Unload Me
       End If
End Sub

Private Function fblnCantidadMaximaValida() As Boolean
1     On Error GoTo NotificaError
          Dim lngCont As Long
          Dim StrCveArticulo As String
          Dim lngDepto As Long
          Dim lngCantidadPendiente As Long
          Dim lngCantidadFaltante As Long
          Dim lngCantidadRequi As Long
          Dim rsContenido As New ADODB.Recordset
          Dim intContenido As Integer
          Dim rsAnterior As New ADODB.Recordset
          Dim lngAnteriorRequi As Long
          
2         lngDepto = cboDepartamentoSolicita.ItemData(cboDepartamentoSolicita.ListIndex)
3         fblnCantidadMaximaValida = True
4         intContenido = 1
5         For lngCont = 1 To grdArticulos.Rows - 1
6             StrCveArticulo = Trim(grdArticulos.TextMatrix(lngCont, intColCveArticulo))
7             If Trim(grdArticulos.TextMatrix(lngCont, intColCveArticulo)) <> "NUEVO" And Trim(grdArticulos.TextMatrix(lngCont, intColEstado)) = "PENDIENTE" Then
8                 If fblnMaximoConfigurado(StrCveArticulo, lngDepto) Then
9                     lngCantidadPendiente = 1    'pendiente de recibir
10                    frsEjecuta_SP Str(lintDiasPendienteRequi) & "|" & StrCveArticulo & "|" & Str(lngDepto), "SP_IVCANTIDADPENDIENTERECIBIR", True, lngCantidadPendiente
11                    lngCantidadFaltante = 1     'maximo - existencia
12                    frsEjecuta_SP StrCveArticulo & "|" & lngDepto, "SP_IVINTCANTIDADFALTANTE", True, lngCantidadFaltante
          
13                    Set rsContenido = frsEjecuta_SP(StrCveArticulo, "SP_IVSELCONTENIDOARTICULO")
14                    If rsContenido.RecordCount <> 0 Then
15                        intContenido = rsContenido!intContenido
16                    End If
                      
17                    lngAnteriorRequi = 0
18                    If vlblnConsulta Then
19                        Set rsAnterior = frsRegresaRs("select intCantidadSolicitada, chrUnidadControl from IvRequisicionDetalle where numnumrequisicion = " + Trim(txtNumero.Text) + " and chrcveArticulo = '" + StrCveArticulo + "'")
20                        If rsAnterior.RecordCount <> 0 Then
21                            lngAnteriorRequi = IIf(IsNull(rsAnterior!IntCantidadSolicitada), 0, rsAnterior!IntCantidadSolicitada) * IIf(rsAnterior!CHRUNIDADCONTROL = "A", intContenido, 1)
22                        End If
23                        rsAnterior.Close
24                    End If
                      
25                    If CBool(grdArticulos.TextMatrix(lngCont, intColUnidad)) Then
26                        lngCantidadRequi = Val(grdArticulos.TextMatrix(lngCont, intColCantidad)) * intContenido
27                    Else
28                        lngCantidadRequi = Val(grdArticulos.TextMatrix(lngCont, intColCantidad))
29                    End If
30                    If (lngCantidadPendiente + lngCantidadRequi - lngAnteriorRequi) > lngCantidadFaltante Then
31                        If lstrArticuloExcede = "" Then
32                            lstrArticuloExcede = Trim(grdArticulos.TextMatrix(lngCont, intColNombreComercial))
33                        Else
34                            lstrArticuloExcede = lstrArticuloExcede + Chr(13) + Trim(grdArticulos.TextMatrix(lngCont, intColNombreComercial))
35                        End If
36                        fblnCantidadMaximaValida = False
37                    End If
38                End If
39            End If
40        Next lngCont
41    Exit Function
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnCantidadMaximaValida" & " Linea:" & Erl()))
End Function

Private Function fblnMaximoConfigurado(cveArt As String, CveDepto As Long) As Boolean
    Dim rs As New ADODB.Recordset
    fblnMaximoConfigurado = False
    Set rs = frsRegresaRs("select mm.* from IvMaximoMinimo mm inner join ivArticulo a on mm.intIdArticulo = a.intIdArticulo where trim(a.chrCveArticulo) = '" + Trim(cveArt) + "' and mm.smiCveDepartamento = " + CStr(CveDepto))
    If rs.RecordCount <> 0 Then
        fblnMaximoConfigurado = True
    End If
End Function

Private Sub cmdImprimir_Click()
1     On Error GoTo NotificaError
      Dim rsReporte As New ADODB.Recordset
      Dim vlstrx As String
      Dim alstrParametros(1) As String
      Dim vlstrdoc As String
      Dim rsAux As New ADODB.Recordset
      Dim lstrTotal As String
          
2         If grdRequisiciones.TextMatrix(grdRequisiciones.Row, grdRequisiciones.Col) = "SALIDA A DEPARTAMENTO" Or cboTipoRequisicion.Text = "SALIDA A DEPARTAMENTO" Then
3             vlstrdoc = "RS"
4         End If

5         If grdRequisiciones.TextMatrix(grdRequisiciones.Row, grdRequisiciones.Col) = "REUBICACION" Or cboTipoRequisicion.Text = "REUBICACION" Then
6             vlstrdoc = "RR"
7         End If

8         If grdRequisiciones.TextMatrix(grdRequisiciones.Row, grdRequisiciones.Col) = "COMPRA - PEDIDO" Or cboTipoRequisicion.Text = "COMPRA - PEDIDO" Then
9             vlstrdoc = "RC"
10        End If
           
11        If vlstrdoc = "RC" And cboProveedor.ListIndex > 0 Then
12            pInstanciaReporte vgrptReporte, "rptrequisicionOrden.rpt"
13            Set rsReporte = frsEjecuta_SP(CStr(Val(txtNumero.Text)), "SP_IVRPTREQUISICIONORDEN")
14            If rsReporte.RecordCount > 0 Then
15                If Not IsNull(rsReporte!monedaProveedor) Then
16                    vlStrSQL = "SELECT NVL(SUM(CASE WHEN '" & Trim(rsReporte!monedaProveedor) & "' = 'DOLARES' AND TRIM(COLISTAPRECIOPROVEEDOR.VCHMONEDA) = 'PESOS' THEN " & _
                                 " NVL((IVREQUISICIONDETALLE.INTCANTIDADSOLICITADA * ROUND((NVL(COLISTAPRECIOPROVEEDOR.MNYCOSTOVIGENTE,0) / " & CStr(ldblTipoCambio) & "),2)) * (1 - (COLISTAPRECIOPROVEEDOR.RELDESCVIGENTE/100)) * (1 + COLISTAPRECIOPROVEEDOR.RELIMPUESTO/100),0) " & _
                                 " ELSE NVL(IVREQUISICIONDETALLE.INTCANTIDADSOLICITADA " & _
                                 " * ROUND(NVL(COLISTAPRECIOPROVEEDOR.MNYCOSTOVIGENTE,0)*CASE WHEN '" & Trim(rsReporte!monedaProveedor) & "' = TRIM(COLISTAPRECIOPROVEEDOR.VCHMONEDA) THEN 1 ELSE " & CStr(ldblTipoCambio) & " END,2) " & _
                                 " * (1 - (COLISTAPRECIOPROVEEDOR.RELDESCVIGENTE/100)) * (1 + COLISTAPRECIOPROVEEDOR.RELIMPUESTO/100),0) END),0) granTotal " & _
                                 " From IVREQUISICIONDETALLE INNER JOIN COLISTAPRECIOPROVEEDOR ON  COLISTAPRECIOPROVEEDOR.INTCVEPROVEEDOR = " & CStr(cboProveedor.ItemData(cboProveedor.ListIndex)) & _
                                 " AND COLISTAPRECIOPROVEEDOR.CHRCVEARTICULO = IVREQUISICIONDETALLE.CHRCVEARTICULO Where IVREQUISICIONDETALLE.NUMNUMREQUISICION = " & CStr(Val(txtNumero.Text))
17                    Set rsAux = frsRegresaRs(vlStrSQL, adLockOptimistic, adOpenDynamic)
18                    lstrTotal = ""
19                    If rsAux.RecordCount <> 0 Then
20                        If Trim(rsReporte!monedaProveedor) = "PESOS" Then
21                            lstrTotal = Trim(fstrNumeroenLetras(CDbl(Format(rsAux!granTotal, "############.00")), "PESOS", "M.N"))
22                        Else
23                            lstrTotal = Trim(fstrNumeroenLetras(CDbl(Format(rsAux!granTotal, "############.00")), "DÓLARES", "USD"))
24                        End If
25                    End If
26                End If
27                vgrptReporte.DiscardSavedData
28                alstrParametros(1) = "totalLetra;" & lstrTotal
29                alstrParametros(0) = "empresa;" & Trim(vgstrNombreHospitalCH)
                  
30                pCargaParameterFields alstrParametros, vgrptReporte
31                pImprimeReporte vgrptReporte, rsReporte, "P", "Requisición/orden de compra"
32            Else
                  'No existe información con esos parámetros.
33                MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje"
34            End If
35        Else
36            pInstanciaReporte vgrptReporte, "rptrequisicion.rpt"
              'Para la impresion de carta o media carta
37            vlStrSQL = "SELECT SMICVEDEPARTAMENTO FROM LOGIN WHERE INTNUMEROLOGIN=" & vglngNumeroLogin
38            Set rsDepartamento = frsRegresaRs(vlStrSQL, adLockOptimistic, adOpenDynamic)
                      
39            If Not rsDepartamento.EOF Then
40                vlintNumDepartamento = rsDepartamento!smicvedepartamento
41            End If
              
42            vlStrSQL = "SELECT INTWIDTH,INTHEIGHT FROM SIPAPELDOCDEPTOMODULO INNER JOIN SITIPODEPAPEL ON SIPAPELDOCDEPTOMODULO.INTIDPAPEL=SITIPODEPAPEL.INTCVETIPODEPAPEL   WHERE VCHMODULO= '" & Trim(cgstrModulo) & "' AND VCHTIPODOCUMENTO='" & Trim(vlstrdoc) & "' AND INTDEPTO=" & vlintNumDepartamento & ""
43            Set rstipodepapel = frsRegresaRs(vlStrSQL, adLockOptimistic, adOpenDynamic)
              
44            vgstrParametrosSP = CStr(Val(txtNumero.Text)) & "|" & CStr(vglngCveAlmacenGeneral)
45            Set rsReporte = frsEjecuta_SP(vgstrParametrosSP, "SP_IVRPTREQUISICIONES")
              
46            If rsReporte.RecordCount > 0 Then
47                  vgrptReporte.DiscardSavedData
                    
48                  If Not rstipodepapel.EOF Then
49                        vgrptReporte.SetUserPaperSize rstipodepapel!intHeight, rstipodepapel!intWidth
50                        vgrptReporte.PaperSize = crPaperUser
51                  End If
                    
52                  alstrParametros(0) = "tienehistorico;" & chkIncluir.Value
53                  alstrParametros(1) = "empresa;" & Trim(vgstrNombreHospitalCH)
54                  pCargaParameterFields alstrParametros, vgrptReporte
                  
55                  pImprimeReporte vgrptReporte, rsReporte, "P", "Requisición de artículos"
56            Else
                  'No existe información con esos parámetros.
57                MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje"
58            End If
59        End If
              
60        If rsReporte.State <> adStateClosed Then rsReporte.Close
          
61        fraCabecera.Enabled = True
62        txtNumero.SetFocus
63    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdImprimir_Click" & " Linea:" & Erl()))
        Unload Me
End Sub

Private Sub cmdManejos_Click()
    frmManejoMedicamentos.blnSoloBusqueda = True
    frmManejoMedicamentos.Show vbModal, Me
End Sub

Private Sub cmdPrimero_Click()
On Error GoTo NotificaError

    If vlblnBusqueda Then
        grdRequisiciones.Row = 1
        If fintCargaReq(grdRequisiciones.TextMatrix(grdRequisiciones.Row, 1)) <> 0 Then
            pMuestra
        End If
    Else
        rsIvRequisicionMaestro.MoveFirst
        pMuestra
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdPrimero_Click"))
    Unload Me
End Sub

Private Sub cmdSiguiente_Click()
On Error GoTo NotificaError
    
    If vlblnBusqueda Then
        If grdRequisiciones.Row <> grdRequisiciones.Rows - 1 Then
            grdRequisiciones.Row = grdRequisiciones.Row + 1
        End If
        If fintCargaReq(grdRequisiciones.TextMatrix(grdRequisiciones.Row, 1)) <> 0 Then
            pMuestra
        End If
    Else
        rsIvRequisicionMaestro.MoveNext
        If rsIvRequisicionMaestro.EOF Then
            rsIvRequisicionMaestro.MovePrevious
        End If
        pMuestra
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdSiguiente_Click"))
    Unload Me
End Sub

Private Sub cmdUltimo_Click()
On Error GoTo NotificaError
    
    If vlblnBusqueda Then
        grdRequisiciones.Row = grdRequisiciones.Rows - 1
        If fintCargaReq(grdRequisiciones.TextMatrix(grdRequisiciones.Row, 1)) <> 0 Then
            pMuestra
        End If
    Else
        rsIvRequisicionMaestro.MoveLast
        pMuestra
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdUltimo_Click"))
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    If KeyAscii = 27 Then
        If sstObj.Tab = 1 Then
            sstObj.Tab = 0
            If txtNumero.Enabled And txtNumero.Visible Then
              txtNumero.SetFocus
            End If
        Else
            If cmdGrabar.Enabled Or vlblnConsulta Then
                If txtCaptura.Visible Then
                    txtCaptura.Visible = False
                    grdArticulos.SetFocus
                Else
                    '¿Desea abandonar la operación?
                    If MsgBox(SIHOMsg(17), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
                        fraCabecera.Enabled = True
                        txtNumero.SetFocus
                    End If
                End If
            Else
                Unload Me
            End If
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyPress"))
    Unload Me
End Sub

Private Sub Form_Load()
1     On Error GoTo NotificaError
      Dim rsTemp As ADODB.Recordset
      Dim rsParametro As ADODB.Recordset
      Dim intOpcionCompraDirecta As Long

2         Me.Icon = frmMenuPrincipal.Icon
          
          'Color de tabs
3         SetStyle sstObj.hwnd, 0
4         SetSolidColor sstObj.hwnd, 16777215
5         SSTabSubclass sstObj.hwnd
          
          
          ' Revisa permiso para Compra directa
6         Select Case cgstrModulo
              Case "IV"
7                 intOpcionCompraDirecta = 2448
8             Case "IM"
9                 intOpcionCompraDirecta = 2449
10            Case "PV"
11                intOpcionCompraDirecta = 2450
12            Case "CP"
13                intOpcionCompraDirecta = 2451
14            Case "LA"
15                intOpcionCompraDirecta = 2452
16            Case "EX"
17                intOpcionCompraDirecta = 2453
18            Case "CN"
19                intOpcionCompraDirecta = 2454
20            Case "CC"
21                intOpcionCompraDirecta = 2455
22            Case "NO"
23                intOpcionCompraDirecta = 2456
24            Case "CA"
25                intOpcionCompraDirecta = 2457
26            Case "AD"
27                intOpcionCompraDirecta = 2458
28            Case "CO"
29                intOpcionCompraDirecta = 2459
30            Case "SI"
31                intOpcionCompraDirecta = 2460
32            Case "BS"
33                intOpcionCompraDirecta = 2461
34            Case "SE"
35                intOpcionCompraDirecta = 2462
36            Case "DI"
37                intOpcionCompraDirecta = 2463
38            Case "TS"
39                intOpcionCompraDirecta = 2464
40        End Select
          
          'Permitir seleccionar proveedor en compra directa, se modifica el tamano de la froma
41        pSeleccionarProveedor
42        ldblTipoCambio = fdblTipoCambio(fdtmServerFecha, "O")
43        blnPermisoCompraDirecta = fblnRevisaPermiso(vglngNumeroLogin, intOpcionCompraDirecta, "E", True)
          
44        vgstrNombreForm = Left(Me.Name, 50)
45        vlstrx = " " & _
            "select NUMNUMREQUISICION, SMICVEDEPTOREQUIS, INTCVEEMPLEAREQUIS, SMICVEDEPTOALMACEN, DTMFECHAREQUISICION, " & _
            "  DTMHORAREQUISICION, VCHESTATUSREQUIS, CHRDESTINO, NUMNUMCUENTA, BITURGENTE, NUMNUMREQUISREL," & _
            "  DTMFECHAREQUISAUTORI, DTMHORAREQUISAUTORI, CHRTIPOPACIENTE, CHRAPLICACIONMED,INTNUMEROLOGIN, " & _
            "  VCHOBSERVACIONES, SMICVEDEPTOGENERA, BITCOMPRADIRECTA, INTCVEPROVEEDOR, INTCAJAMATERIAL " & _
            " from IvRequisicionMaestro where smiCveDeptoGenera = " & vgintNumeroDepartamento & " and numNumRequisicion = -1"
46        Set rsIvRequisicionMaestro = frsRegresaRs(vlstrx, adLockOptimistic, adOpenDynamic)
          
47        vlstrx = " " & _
            "select NUMNUMREQUISICION, CHRCVEARTICULO, INTCANTIDADSOLICITADA, CHRUNIDADCONTROL, VCHESTATUSDETREQUIS, VCHNOMBREARTICULONUEVO, VCHUNIDADARTICULONUEVO " & _
            " from IvRequisicionDetalle where numNumRequisicion=null"
48        Set rsIvRequisicionDetalle = frsRegresaRs(vlstrx, adLockOptimistic, adOpenDynamic)
          
49        Set rsTemp = frsSelParametros("CO", -1, "BITAUTORIZAREQUISICION")
50        If Not rsTemp.EOF Then
51            lintAutorizaRequi = IIf(rsTemp!Valor = "0", 0, 1)
52        End If
53        rsTemp.Close
          
          'Clave del manejo configurado como controlado
54        Set rsTemp = frsEjecuta_SP("-1|1|1", "Sp_IvSelManejos")
55        If rsTemp.EOF Then
              '¡No se han configurado el manejo para los medicamentos controlados!
56            MsgBox SIHOMsg(1108), vbOKOnly + vbInformation, "Mensaje"
57        End If
58        rsTemp.Close
59        lblnReubicarInsumos = False
60        pCargaDepartamentoSolicita
61        pCargaLocalizacion
62        pCargaCajasMaterial
63        pLimpia
64        sstObj.Tab = 0
          ' No permitido para requerir medicamento controlado
65        vlblnControlado = False
          
66        fraBarra.Visible = False
          
67        pAlmacenConsigna
          
68        vlblnNoFocus = False
          
69        Set rsParametro = frsSelParametros("IV", vgintClaveEmpresaContable, "BITVALIDARMAXIMOOTRASENTRADAS")
70        If Not rsParametro.EOF Then
71            lblnValidarExcedeMaximo = IIf(IIf(IsNull(rsParametro("Valor")), 0, rsParametro("Valor")) = 0, False, True)
72        Else
73            lblnValidarExcedeMaximo = False
74        End If
75        Set rsParametro = frsSelParametros("SI", vgintClaveEmpresaContable, "INTDIASREQUISICION")
76        If Not rsParametro.EOF Then
77            lintDiasPendienteRequi = IIf(IsNull(rsParametro("Valor")), 0, rsParametro("Valor"))
78        Else
79            lintDiasPendienteRequi = 0
80        End If
81        rsParametro.Close
          
82        lintInterfazFarmaciaSJP = 0
83        Set rsParametro = frsSelParametros("IV", -1, "INTINTERFAZFARMACIASUBRROGADASJP")
84        If Not rsParametro.EOF Then
85            lintInterfazFarmaciaSJP = IIf(IsNull(rsParametro("Valor")), 0, rsParametro("Valor"))
86        Else
87            lintInterfazFarmaciaSJP = 0
88        End If
89        llngDeptoSubrogado = 0
90        If lintInterfazFarmaciaSJP = 1 Then
              'Consultar parametro de departamento subrogado
91            Set rsParametro = frsSelParametros("IV", -1, "INTDEPTOINTERFAZFARMACIA")
92            If Not rsParametro.EOF Then
93                llngDeptoSubrogado = IIf(IsNull(rsParametro("Valor")), 0, rsParametro("Valor"))
94            Else
95                llngDeptoSubrogado = 0
96            End If
97        End If
          
98        position = 0
          
99    Exit Sub
NotificaError:
       Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load" & " Linea:" & Erl()))
       Unload Me
End Sub

Private Sub pSeleccionarProveedor()
1     On Error GoTo NotificaError
      Dim rs As ADODB.Recordset
2         lblnSeleccionarProveedor = False
3         Set rs = frsSelParametros("IV", vgintClaveEmpresaContable, "BITSELECCIONARPROVEEDORREQUISICION")
4         If rs.RecordCount <> 0 Then
5             lblnSeleccionarProveedor = IIf(Trim(rs!Valor) = "1", True, False)
6         End If
          'proveedores
7         cboProveedor.Clear
8         cboProveedor.Enabled = False
9         lblProveedor.Enabled = False
          'se agregan los proveedores por si se consulta alguna requisicion con proveedor aunque el parámetro no esté prendido
10        Set rs = frsRegresaRs("SELECT intCveProveedor, vchNombre FROM COPROVEEDOR WHERE TRIM(vchTipoProveedor) IN ('PRODUCTOS','AMBOS') AND bitActivo = 1 ")
11        Call pLlenarCboRs_new(cboProveedor, rs, 0, 1, -1)
12        rs.Close
13        cboProveedor.AddItem "<NINGUNO>", 0
14        pMostrarProveedor (lblnSeleccionarProveedor)
          
15    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pSeleccionarProveedor" & " Linea:" & Erl()))
End Sub

Private Sub pMostrarProveedor(lblnMostrar As Boolean)
On Error GoTo NotificaError
    If lblnMostrar And Not cboProveedor.Visible Then
        cboProveedor.Visible = True
        lblProveedor.Visible = True
    ElseIf Not lblnMostrar And cboProveedor.Visible Then
        '-320
        cboProveedor.Visible = False
        lblProveedor.Visible = False
    End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pMostrarProveedor"))
End Sub

Private Sub pLimpia()
1     On Error GoTo NotificaError
      Dim X As Integer
      Dim lintAux As Integer

2         vlblnConsulta = False
3         vlblnBusqueda = False
          
4         pHabilitaArticulo True
5         txtCaptura.Visible = False
6         txtNumero.Text = frsRegresaRs("select isnull(max(numNumRequisicion),0)+1 from IvRequisicionMaestro").Fields(0)
7         txtFecha.Text = Format(fdtmServerFecha, "dd/mmm/yyyy")
8         chkUrgente.Value = 0
9         txtEstatus.Text = "PENDIENTE"
10        MyButton1.Enabled = False
11        cboDepartamentoSolicita.ListIndex = fintLocalizaCbo_new(cboDepartamentoSolicita, CStr(vgintNumeroDepartamento))
          
12        chkIncluir.Enabled = False
13        cboTipoRequisicion.Clear
14        vlstrx = "select count(*) from IvRequisicionDepartamento where  chrTipoRequisicion = 'SD' and intNumeroLogin = " & Str(vglngNumeroLogin)
15        If frsRegresaRs(vlstrx).Fields(0) > 0 Then
16            cboTipoRequisicion.AddItem "SALIDA A DEPARTAMENTO"
17        End If
          
18        vlstrx = "select count(*) from IvRequisicionDepartamento where  chrTipoRequisicion = 'RE' and intNumeroLogin = " & Str(vglngNumeroLogin)
19        If frsRegresaRs(vlstrx).Fields(0) > 0 Then
20            cboTipoRequisicion.AddItem "REUBICACION"
21        End If
          
22        llngDeptoRecibe = 0
23        vlstrx = "select count(*) from IvRequisicionDepartamento where  chrTipoRequisicion = 'CO' and intNumeroLogin = " & Str(vglngNumeroLogin)
24        If frsRegresaRs(vlstrx).Fields(0) > 0 Then
              'Obtiene el almacén que autoriza y el que recibe
25            vlstrx = "Select intCveDeptoAutoriza, intCveDeptoRecibe From IvRequisDepartamentoCompra Where intNumeroLogin = " & vglngNumeroLogin
26            Set rs = frsRegresaRs(vlstrx)
                  
27            If rs.RecordCount > 0 Then
28                llngDeptoRecibe = IIf(IsNull(rs!intCveDeptoRecibe), 0, rs!intCveDeptoRecibe)
29                If lintAutorizaRequi = 1 Then
30                    If Not IsNull(rs!intCveDeptoAutoriza) Then
31                        llngDeptoAutoriza = rs!intCveDeptoAutoriza
32                        cboTipoRequisicion.AddItem "COMPRA - PEDIDO"
33                    End If
34                Else
35                    cboTipoRequisicion.AddItem "COMPRA - PEDIDO"
36                End If
37            End If
38        End If
          
39        vlstrx = "select count(*) from IvRequisicionDepartamento where  chrTipoRequisicion = 'AG' and intNumeroLogin = " & Str(vglngNumeroLogin)
40        If frsRegresaRs(vlstrx).Fields(0) > 0 Then
41            cboTipoRequisicion.AddItem "ALMACEN ABASTECIMIENTO"
42        End If
          
43        cboTipoRequisicion.ListIndex = -1
44        chkCompradirecta.Enabled = False
45        chkCompradirecta.Value = 0
          
46        txtEmpleadoSolicito.Text = ""
          
47        cboAlmacenSurte.Clear
48        cboAlmacenSurte.ListIndex = -1
          
49        fraCabecera.Enabled = True
          
50        optOpcion(0).Value = False
51        optOpcion(1).Value = False
52        optOpcion(2).Value = False
53        optOpcion(3).Value = False
          
54        cboLocalizacion.ListIndex = -1
55        cboFamilia.ListIndex = -1
56        cboNombreComercial.ListIndex = -1
57        cboSubfamilia.ListIndex = -1
58        txtNombreGenerico.Text = ""
59        txtCodigoBarras.Text = ""
60        txtClave.Text = ""
61        txtExistencia.Text = ""
62        txtCantidadSolicitada.Text = ""
63        Me.txtnombrecompleto.Text = ""
          
64        lblCantidad.Enabled = False
65        txtCantidadSolicitada.Enabled = False
66        OptAlterna.Enabled = False
67        OptMinima.Enabled = False
          
68        OptAlterna.Value = True
69        OptMinima.Value = False
          
70        fraFiltros.Enabled = False
          
71        intMaxManejos = 0
72        pFormatoArticulos False
          
73        cmdAgregar.Enabled = True
74        cmdCancelar.Enabled = False
75        cmdBorrar.Enabled = False
          
76        mskFecIni.Mask = ""
77        mskFecIni.Text = DateAdd("d", -3, fdtmServerFecha)
78        mskFecIni.Mask = "##/##/####"
          
79        mskFecFin.Mask = ""
80        mskFecFin.Text = fdtmServerFecha
81        mskFecFin.Mask = "##/##/####"
          
82        vlstrx = "" & _
          "select " & _
              "Distinct " & _
              "Case chrDestino " & _
              "when 'R' then 'REUBICACION' " & _
              "when 'C' then 'COMPRA - PEDIDO' " & _
              "when 'D' then 'SALIDA A DEPARTAMENTO' " & _
              "when 'A' then 'ALMACEN ABASTECIMIENTO' " & _
              "when 'O' then 'CONSIGNACION' " & _
              "when 'U' then 'PEDIDO SUGERIDO' " & _
              "End " & _
              "TipoRequisicion," & _
              "0 Clave " & _
          "From " & _
              "IvRequisicionMaestro " & _
          "Where " & _
              "chrDestino <> 'P' " & _
              "and smiCveDeptoGenera = " & Str(vgintNumeroDepartamento) & " " & _
          "Union " & _
          "select " & _
              "'<TODAS>' TipoRequisicion," & _
              "0 Clave FROM dual"
83        Set rs = frsRegresaRs(vlstrx)
84        pLlenarCboRs_new cboTRBusqueda, rs, 1, 0
          
85        vlstrx = "SELECT DISTINCT TRIM(VCHESTATUSREQUIS) ESTATUS, 1 CLAVE FROM IVREQUISICIONMAESTRO WHERE SMICVEDEPTOGENERA = " & Str(vgintNumeroDepartamento) & " AND TRIM(VCHESTATUSREQUIS) IS NOT NULL ORDER BY ESTATUS ASC"
86        Set rs = frsRegresaRs(vlstrx)
87        pLlenarCboRs_new cboERBusqueda, rs, 1, 0, 3
88        cboERBusqueda.ListIndex = 0
89        txtTotal.Text = ""
90        fraTotales.Visible = False
          
91        txtObservaciones.Text = ""
92        txtObservaciones.Enabled = True
          
93        If lblnRecargarCajas Then
94            pCargaCajasMaterial
95        End If
96        lblCajaMaterial.Enabled = False
97        cboCajaMaterial.Enabled = False
98        cboCajaMaterial.ListIndex = -1
99        llngCveCajaMaterialRequi = -1
          
100       pHabilita False, False, False, False, False, False, False
          
101       cmdCierraReq.Enabled = False ' no entra dentro de la botonera por que requiere validacion especial para activarse o desactivarse
              
102       pMostrarProveedor (lblnSeleccionarProveedor)
          'Quitar los proveedores que no estén activos
103       lintAux = -1
104       For X = 0 To cboProveedor.ListCount - 1
105           If cboProveedor.ItemData(X) < 0 Then
106               lintAux = X
107               Exit For
108           End If
109       Next X
110       If lintAux > 0 Then cboProveedor.RemoveItem lintAux
111       If cboProveedor.ListCount > 0 Then cboProveedor.ListIndex = 0
          
112   Exit Sub
NotificaError:
       Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLimpia" & " Linea:" & Erl()))
End Sub

Private Sub pAlmacenConsigna()
1     On Error GoTo NotificaError
      Dim rsConSigna As New ADODB.Recordset
          
          'si el usuario pertenece a un almacen tipo consignacion no podrá modificar las requisiciones solo cancelar articulos, consultar e imprimir
2         lblnAlmacenConsigna = False
3         vgstrParametrosSP = "1|" & vgintClaveEmpresaContable & "|" & vgintNumeroDepartamento
4         Set rsConSigna = frsEjecuta_SP(vgstrParametrosSP, "Sp_Ivselalmacenconsigna")
5         If rsConSigna.RecordCount > 0 Then
6             lblnAlmacenConsigna = True
7         End If
8         rsConSigna.Close
          
9     Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pAlmacenConsigna" & " Linea:" & Erl()))
End Sub

Private Sub pFormatoArticulos(vlblnVerColumnaExistencia As Boolean)
1     On Error GoTo NotificaError
          
2         With grdArticulos
              
3             .Clear
4             .Rows = 2
5             .Cols = 16
6             .FixedCols = 1
7             .FixedRows = 1
8             .FormatString = "|.|.|.|.|Clave|Nombre comercial|Cantidad|Unidad|Existencia|Estado||Proveedor de ultima compra|Fecha|Cantidad|Costo"
              
9             For intcontador = 1 To intColManejo4
10                .Col = intcontador
11                .Row = 0
12                .ColAlignment(intcontador) = flexAlignCenterCenter
13                .CellForeColor = &H8000000F
14            Next intcontador
              
15            .ColWidth(0) = 100
16            .ColWidth(intColManejo1) = IIf(intMaxManejos >= 4, 300, 0)
17            .ColWidth(intColManejo2) = IIf(intMaxManejos >= 3, 300, 0)
18            .ColWidth(intColManejo3) = IIf(intMaxManejos >= 2, 300, 0)
19            .ColWidth(intColManejo4) = IIf(intMaxManejos >= 1, 300, 0)
20            .ColWidth(intColCveArticulo) = 1150         'Clave
21            .ColWidth(intColNombreComercial) = 5000     'Nombre comercial
22            .ColWidth(intColCantidad) = 1000            'Cantidad
23            .ColWidth(intColDescripcionUnidad) = 1800    'Unidad
24            .ColWidth(intColExistencia) = IIf(vlblnVerColumnaExistencia, 1000, 0) 'Existencia
25            .ColWidth(intColEstado) = 1800              'Estado
26            .ColWidth(intColUnidad) = 0                 'Alterna/Minima
27            .ColWidth(intColProveedor) = IIf(chkIncluir.Value = 1, 3000, 0)
28            .ColWidth(intColFecha) = IIf(chkIncluir.Value = 1, 1100, 0)
29            .ColWidth(intColCantidadProveedor) = IIf(chkIncluir.Value = 1, 1500, 0)
30            .ColWidth(intColCosto) = IIf(chkIncluir.Value = 1, 1500, 0)
              
31            .ColAlignmentFixed(intColCveArticulo) = flexAlignLeftCenter
32            .ColAlignmentFixed(intColNombreComercial) = flexAlignLeftCenter
33            .ColAlignmentFixed(intColCantidad) = flexAlignCenterCenter
34            .ColAlignmentFixed(intColDescripcionUnidad) = flexAlignLeftCenter
35            .ColAlignmentFixed(intColExistencia) = flexAlignRightCenter
36            .ColAlignmentFixed(intColEstado) = flexAlignLeftCenter
37            .ColAlignmentFixed(intColFecha) = flexAlignLeftCenter
38            .ColAlignmentFixed(intColCantidadProveedor) = flexAlignRightCenter
39            .ColAlignmentFixed(intColCosto) = flexAlignRightCenter
              
40            .ColAlignment(intColCveArticulo) = flexAlignLeftCenter
41            .ColAlignment(intColNombreComercial) = flexAlignLeftCenter
42            .ColAlignment(intColCantidad) = flexAlignRightCenter
43            .ColAlignment(intColDescripcionUnidad) = flexAlignLeftCenter
44            .ColAlignment(intColExistencia) = flexAlignRightCenter
45            .ColAlignment(intColEstado) = flexAlignLeftCenter
46            .ColAlignment(intColCantidadProveedor) = flexAlignRightCenter
47            .ColAlignment(intColCosto) = flexAlignRightCenter
              
48            pFormatoNumeroColumnaGrid grdArticulos, intColCantidadProveedor
49            pFormatoNumeroColumnaGrid grdArticulos, intColCosto, "$ "

             ' .MergeCells = flexMergeRestrictRows se comento por que no dejaba seleccionar los renglones
50            .MergeRow(0) = True
              
51        End With

52    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pFormatoArticulos" & " Linea:" & Erl()))
        Unload Me
End Sub

Private Sub pHabilita(vlbln1 As Boolean, vlbln2 As Boolean, vlbln3 As Boolean, vlbln4 As Boolean, vlbln5 As Boolean, vlbln6 As Boolean, vlbln7 As Boolean)
On Error GoTo NotificaError

    cmdPrimero.Enabled = IIf(rsIvRequisicionMaestro.RecordCount = 0, False, vlbln1)
    cmdAnterior.Enabled = IIf(rsIvRequisicionMaestro.RecordCount = 0, False, vlbln1)
    cmdBuscar.Enabled = Not vlbln6 'IIf(rsIvRequisicionMaestro.RecordCount = 0, False, vlbln1)
    cmdSiguiente.Enabled = IIf(rsIvRequisicionMaestro.RecordCount = 0, False, vlbln1)
    cmdUltimo.Enabled = IIf(rsIvRequisicionMaestro.RecordCount = 0, False, vlbln1)
    cmdGrabar.Enabled = vlbln6
    cmdImprimir.Enabled = vlbln7
 
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pHabilita"))
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then
    If sstObj.Tab = 1 Then
            sstObj.Tab = 0
            If txtNumero.Enabled And txtNumero.Visible Then
              txtNumero.SetFocus
            End If
             Cancel = True
        Else
            If cmdGrabar.Enabled Or vlblnConsulta Then
                If txtCaptura.Visible Then
                    txtCaptura.Visible = False
                    grdArticulos.SetFocus
                Else
                    '¿Desea abandonar la operación?
                    If MsgBox(SIHOMsg(17), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
                        fraCabecera.Enabled = True
                        txtNumero.SetFocus
                    End If
                End If
                 Cancel = True
            Else
             Cancel = False
                'Unload Me
            End If
        End If
End If

End Sub

Private Sub grdArticulos_Click()
If grdArticulos.TextMatrix(grdArticulos.Row, intColEstado) = "PENDIENTE" And vlblnConsulta And llngCveCajaMaterialRequi = -1 Then
    Me.cmdCancelar.Enabled = True
 Else
    Me.cmdCancelar.Enabled = False
End If

End Sub

Private Sub grdArticulos_GotFocus()

  txtnombrecompleto.Text = ""
  If grdArticulos.Row > 0 Then
    txtnombrecompleto.Text = grdArticulos.TextMatrix(grdArticulos.Row, intColNombreComercial)
  End If
  
End Sub

Private Sub grdArticulos_KeyPress(KeyAscii As Integer)
1     On Error GoTo NotificaError
2      If KeyAscii <> 13 Then
3      Else
4        If Trim(grdArticulos.TextMatrix(1, intColCveArticulo)) <> "" And Not lblnAlmacenConsigna Then
              'Nombre o Unidad
5             If Trim(grdArticulos.TextMatrix(grdArticulos.Row, intColCveArticulo)) = "NUEVO" And Trim(cboTipoRequisicion.List(cboTipoRequisicion.ListIndex)) = "COMPRA - PEDIDO" And (grdArticulos.Col = intColNombreComercial Or grdArticulos.Col = intColDescripcionUnidad Or grdArticulos.Col = intColCantidad) And Trim(grdArticulos.TextMatrix(grdArticulos.Row, intColEstado)) = "PENDIENTE" Then
6                 If grdArticulos.Col = intColCantidad Then
7                     txtCaptura.Move grdArticulos.Left + grdArticulos.CellLeft, grdArticulos.Top + grdArticulos.CellTop, grdArticulos.CellWidth - 8, grdArticulos.CellHeight - 8
8                     txtCaptura.Alignment = 1
9                     If IsNumeric(Chr(KeyAscii)) Then
10                        txtCaptura.Text = Chr(KeyAscii)
11                    Else
12                        txtCaptura.Text = grdArticulos.TextMatrix(grdArticulos.Row, grdArticulos.Col)
13                    End If
14                    txtCaptura.Visible = True
15                    txtCaptura.SelStart = Len(txtCaptura.Text)
16                    txtCaptura.SetFocus
17                Else
18                    txtCaptura.Move grdArticulos.Left + grdArticulos.CellLeft, grdArticulos.Top + grdArticulos.CellTop, grdArticulos.CellWidth - 8, grdArticulos.CellHeight - 8
19                    txtCaptura.Alignment = 0
                      'txtCaptura.Text = UCase(Chr(KeyAscii))  'Lo comenté porque inserta un enter que luego estorba y no aparece*
20                    txtCaptura.Visible = True
                      'txtCaptura.Text = ""
21                    txtCaptura.Text = grdArticulos.TextMatrix(grdArticulos.Row, grdArticulos.Col)
22                    txtCaptura.SelStart = Len(txtCaptura.Text)
23                    txtCaptura.SetFocus
                      
24                End If
25            End If
26        End If
27    End If
28    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdArticulos_KeyPress" & " Linea:" & Erl()))
        Unload Me
End Sub

Private Sub grdArticulos_RowColChange()
  
  txtnombrecompleto.Text = ""
  If grdArticulos.Row > 0 Then
    txtnombrecompleto.Text = grdArticulos.TextMatrix(grdArticulos.Row, intColNombreComercial)
  End If
  
End Sub

Private Sub grdRequisiciones_DblClick()
On Error GoTo NotificaError

    If Trim(grdRequisiciones.TextMatrix(1, 1)) <> "" Then
        If fintCargaReq(grdRequisiciones.TextMatrix(grdRequisiciones.Row, 1)) <> 0 Then
            vlblnBusqueda = True
            sstObj.Tab = 0
            txtNumero.Text = grdRequisiciones.TextMatrix(grdRequisiciones.Row, 1)
            pMuestra
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdRequisiciones_DblClick"))
    Unload Me
End Sub

Private Sub grdRequisiciones_KeyPress(KeyAscii As Integer)
1     On Error GoTo NotificaError
          
2         If KeyAscii = vbKeyReturn Then
          
3             If Trim(grdRequisiciones.TextMatrix(1, 1)) <> "" Then
4                 If fintCargaReq(grdRequisiciones.TextMatrix(grdRequisiciones.Row, 1)) <> 0 Then
                      
5                     vlblnBusqueda = True
6                     sstObj.Tab = 0
7                     txtNumero.Text = grdRequisiciones.TextMatrix(grdRequisiciones.Row, 1)
8                     pMuestra
9                 End If
10            End If
              
11        End If

12    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdRequisiciones_KeyPress" & " Linea:" & Erl()))
        Unload Me
End Sub

Private Sub mskFecFin_GotFocus()
On Error GoTo NotificaError
    pSelMkTexto mskFecFin
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFecFin_GotFocus"))
    Unload Me
End Sub

Private Sub mskFecFin_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    If KeyAscii = vbKeyReturn Then cboTRBusqueda.SetFocus
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFecFin_KeyPress"))
    Unload Me
End Sub

Private Sub mskFecFin_LostFocus()
On Error GoTo NotificaError

    If Trim(mskFecFin.ClipText) = "" Then
        mskFecFin.Mask = ""
        mskFecFin.Text = fdtmServerFecha
        mskFecFin.Mask = "##/##/####"
    End If
    If Not IsDate(mskFecFin.Text) Then
        mskFecFin.Mask = ""
        mskFecFin.Text = fdtmServerFecha
        mskFecFin.Mask = "##/##/####"
    End If

    pCargaRequisiciones
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFecFin_LostFocus"))
    Unload Me
End Sub

Private Sub mskFecini_GotFocus()
On Error GoTo NotificaError
    pSelMkTexto mskFecIni
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFecIni_GotFocus"))
    Unload Me
End Sub

Private Sub mskFecIni_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    If KeyAscii = vbKeyReturn Then mskFecFin.SetFocus
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFecIni_KeyPress"))
    Unload Me
End Sub

Private Sub mskFecIni_LostFocus()
On Error GoTo NotificaError

    If Trim(mskFecIni.ClipText) = "" Then
        mskFecIni.Mask = ""
        mskFecIni.Text = fdtmServerFecha
        mskFecIni.Mask = "##/##/####"
    End If
    If Not IsDate(mskFecIni.Text) Then
        mskFecIni.Mask = ""
        mskFecIni.Text = fdtmServerFecha
        mskFecIni.Mask = "##/##/####"
    End If
        
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFecIni_LostFocus"))
    Unload Me
End Sub

Private Sub MyButton1_Click()
Dim rsDescripcion As ADODB.Recordset
Dim mensaje As String
    Set rsDescripcion = frsRegresaRs("select vchDescripcionRequis from ivrequisicionmaestro where numnumrequisicion =" & txtNumero.Text)
       ' If rsEstatus.RecordCount > 0 Then txtEstatus.Text = rsEstatus!vchEstatusRequis
        'rsEstatus!vchEstatusRequis
        
    mensaje = IIf(IsNull(rsDescripcion!vchDescripcionRequis), "No tiene observación", (rsDescripcion!vchDescripcionRequis))
    MsgBox mensaje, vbOKOnly + vbInformation, "Observaciones"
    
End Sub

Private Sub optAlterna_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    If KeyAscii = vbKeyReturn Then cmdAgregar.SetFocus
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":optAlterna_KeyPress"))
    Unload Me
End Sub

Private Sub optMinima_GotFocus()
On Error GoTo NotificaError
    
    If Trim(cboTipoRequisicion.List(cboTipoRequisicion.ListIndex)) = "COMPRA - PEDIDO" _
      Or Trim(cboTipoRequisicion.List(cboTipoRequisicion.ListIndex)) = "ALMACEN ABASTECIMIENTO" Then
        OptAlterna.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":optMinima_GotFocus"))
    Unload Me
End Sub

Private Sub OptMinima_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    If KeyAscii = vbKeyReturn Then cmdAgregar.SetFocus
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":optMinima_KeyPress"))
    Unload Me
End Sub

Private Sub optOpcion_Click(Index As Integer)
On Error GoTo NotificaError
    pCargaFiltros
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":optOpcion_Click"))
    Unload Me
End Sub

Private Sub optOpcion_GotFocus(Index As Integer)
1     On Error GoTo NotificaError
          
2         cboLocalizacion.ListIndex = 0
          
3         cboFamilia.Clear
4         cboFamilia.AddItem "<TODAS>", 0
5         cboFamilia.ItemData(cboFamilia.NewIndex) = 0
6         cboFamilia.ListIndex = 0
          
7         cboSubfamilia.Clear
8         cboSubfamilia.AddItem "<TODAS>", 0
9         cboSubfamilia.ItemData(cboSubfamilia.NewIndex) = 0
10        cboSubfamilia.ListIndex = 0
          
11        cboNombreComercial.Clear
12        cboNombreComercial.AddItem "<TODOS>", 0
13        cboNombreComercial.ItemData(cboNombreComercial.NewIndex) = 0
14        cboNombreComercial.ListIndex = 0
          
15        txtNombreGenerico.Text = ""
16        txtCodigoBarras.Text = ""
17        txtClave.Text = ""
18        txtExistencia.Text = ""
19        txtCantidadSolicitada.Text = ""
20        OptAlterna.Value = True
21        OptMinima.Value = False

22    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":optOpcion_GotFocus" & " Linea:" & Erl()))
        Unload Me
End Sub

Private Sub optOpcion_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo NotificaError
    
    If KeyAscii = vbKeyReturn Then
        pCargaFiltros
        cboLocalizacion.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":optOpcion_KeyPress"))
    Unload Me
End Sub


Private Sub txtCantidadSolicitada_Change()
On Error GoTo NotificaError

    chkPedirMaximo.Enabled = Val(txtCantidadSolicitada.Text) = 0
    If Not chkPedirMaximo.Enabled Then
        chkPedirMaximo.Value = 0
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtCantidadSolicitada_Change"))
    Unload Me
End Sub

Private Sub txtCantidadSolicitada_GotFocus()
On Error GoTo NotificaError
    
    pSelTextBox txtCantidadSolicitada

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtCantidadSolicitada_GotFocus"))
    Unload Me
End Sub

Private Sub txtCantidadSolicitada_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    
    If KeyAscii = vbKeyReturn Then
        If Val(txtCantidadSolicitada.Text) = 0 Then
            pEnfocaTextBox txtCantidadSolicitada
        Else
            OptAlterna.SetFocus
        End If
    Else
        pValidaSoloNumero KeyAscii
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtCantidadSolicitada_KeyPress"))
    Unload Me
End Sub

Private Sub txtCaptura_KeyPress(KeyAscii As Integer)
1     On Error GoTo NotificaError
          
2         If grdArticulos.Col = intColCantidad Then
3             If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
4                 KeyAscii = 7
5             Else
6                 If KeyAscii = vbKeyReturn Then
7                     If Val(txtCaptura.Text) <> 0 Then
8                         grdArticulos.TextMatrix(grdArticulos.Row, grdArticulos.Col) = txtCaptura.Text
9                         grdArticulos.TextMatrix(grdArticulos.Row, intColEstado) = "PENDIENTE"
10                        txtCaptura.Visible = False
11                        pHabilita False, False, False, False, False, True, False
12                        grdArticulos.Col = IIf(Trim(grdArticulos.TextMatrix(grdArticulos.Row, intColCveArticulo)) = "NUEVO", intColDescripcionUnidad, intColCantidad)
13                        grdArticulos.SetFocus
14                    End If
15                End If
16            End If
17        Else
18            If KeyAscii = vbKeyReturn Then
19                grdArticulos.TextMatrix(grdArticulos.Row, grdArticulos.Col) = Trim(txtCaptura.Text)
20                txtCaptura.Visible = False
21                pHabilita False, False, False, False, False, True, False
22                grdArticulos.Col = IIf(grdArticulos.Col = intColNombreComercial, intColCantidad, intColDescripcionUnidad)
23                grdArticulos.SetFocus
24            Else
25                KeyAscii = Asc(UCase(Chr(KeyAscii)))
26            End If
27        End If

28    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtCaptura_KeyPress" & " Linea:" & Erl()))
        Unload Me
End Sub

Private Sub txtCaptura_LostFocus()
    txtCaptura.Visible = False
End Sub

Private Sub txtClave_KeyPress(KeyAscii As Integer)
1     On Error GoTo NotificaError
          
2         If KeyAscii = vbKeyReturn Then
3             If Trim(txtClave.Text) <> "" Then
4                 txtNombreGenerico.Text = ""
5                 txtExistencia.Text = ""
6                 txtCantidadSolicitada.Text = ""
7                 txtCodigoBarras.Text = ""
              
8                 vgstrVarIntercam = UCase(txtClave.Text) 'Variable global de entrada al frmlista para la busqueda
9                 vgstrVarIntercam2 = "Lista por clave" 'Variable global que es el titulo del formulario frmlista
                  
10                frmLista.gintEstatus = 1
11                frmLista.gintFamilia = cboFamilia.ItemData(cboFamilia.ListIndex)
12                frmLista.gintSubfamilia = cboSubfamilia.ItemData(cboSubfamilia.ListIndex)
13                frmLista.Tag = IIf(cboTipoRequisicion.List(cboTipoRequisicion.ListIndex) = "REUBICACION" And lblnReubicarInsumos, "T", Mid(cboTipoRequisicion.List(cboTipoRequisicion.ListIndex), 1, 1))
14                frmLista.Show vbModal, Me
15                If Len(vgstrVarIntercam) > 0 Then
16                    pCargaArticulo vgstrVarIntercam
17                End If
18            Else
19                cboNombreComercial.SetFocus
20            End If
21        Else
22            pValidaNumero KeyAscii
23        End If

24    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtClave_KeyPress" & " Linea:" & Erl()))
        Unload Me
End Sub

Private Sub txtCodigoBarras_KeyPress(KeyAscii As Integer)
1     On Error GoTo NotificaError
          
2         If KeyAscii = vbKeyReturn Then
3             If Trim(txtCodigoBarras.Text) <> "" Then
4                 txtNombreGenerico.Text = ""
5                 txtClave.Text = ""
6                 txtExistencia.Text = ""
7                 txtCantidadSolicitada.Text = ""
                  
8                 vgstrVarIntercam = UCase(txtCodigoBarras.Text) 'Variable global de entrada al frmlista para la busqueda
9                 vgstrVarIntercam2 = "Lista por código de barras" 'Variable global que es el titulo del formulario frmlista
                  
10                frmLista.gintEstatus = 1
11                frmLista.gintFamilia = cboFamilia.ItemData(cboFamilia.ListIndex)
12                frmLista.gintSubfamilia = cboSubfamilia.ItemData(cboSubfamilia.ListIndex)
13                frmLista.Tag = IIf(cboTipoRequisicion.List(cboTipoRequisicion.ListIndex) = "REUBICACION" And lblnReubicarInsumos, "T", Mid(cboTipoRequisicion.List(cboTipoRequisicion.ListIndex), 1, 1))
14                frmLista.Show vbModal, Me
15                If Len(vgstrVarIntercam) > 0 Then
16                    pCargaArticulo vgstrVarIntercam
17                End If
                  
18            Else
19                txtClave.SetFocus
20            End If
21        Else
22            KeyAscii = Asc(UCase(Chr(KeyAscii)))
23        End If

24    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtCodigoBarras_KeyPress" & " Linea:" & Erl()))
        Unload Me
End Sub

Private Sub txtEstatus_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    If KeyAscii = vbKeyReturn Then cboTipoRequisicion.SetFocus
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtEstatus_KeyPress"))
    Unload Me
End Sub


Private Sub txtNombreGenerico_Change()
On Error GoTo NotificaError

    lblCantidad.Enabled = cboNombreComercial.ListIndex <> -1
    txtCantidadSolicitada.Enabled = cboNombreComercial.ListIndex <> -1
    OptAlterna.Enabled = cboNombreComercial.ListIndex <> -1
    OptMinima.Enabled = cboNombreComercial.ListIndex <> -1

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboNombreComercial_Change"))
    Unload Me
End Sub

Private Sub txtNombreGenerico_KeyPress(KeyAscii As Integer)
1     On Error GoTo NotificaError
          
2         KeyAscii = Asc(StrConv(Chr$(KeyAscii), vbUpperCase))
          
3         If KeyAscii = vbKeyReturn Then
              
4                 If Len(txtNombreGenerico.Text) > 0 Then
5                     txtCodigoBarras.Text = ""
6                     txtClave.Text = ""
7                     txtExistencia.Text = ""
8                     txtCantidadSolicitada.Text = ""
                      
9                     vgstrVarIntercam = UCase(txtNombreGenerico.Text) 'Variable global de entrada al frmlista para la busqueda
10                    vgstrVarIntercam2 = "Lista por nombre genérico" 'Variable global que es el titulo del formulario frmlista
                      
11                    frmLista.gintEstatus = 1
12                    frmLista.gintFamilia = cboFamilia.ItemData(cboFamilia.ListIndex)
13                    frmLista.gintSubfamilia = cboSubfamilia.ItemData(cboSubfamilia.ListIndex)
14                    frmLista.Tag = IIf(cboTipoRequisicion.List(cboTipoRequisicion.ListIndex) = "REUBICACION" And lblnReubicarInsumos, "T", Mid(cboTipoRequisicion.List(cboTipoRequisicion.ListIndex), 1, 1))
15                    frmLista.Show vbModal, Me
                      
16                    If Trim(vgstrVarIntercam) <> "" Then
17                        pCargaArticulo vgstrVarIntercam
18                    End If
19                Else
20                    txtCodigoBarras.SetFocus
21                End If
             
22        End If

23    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtNombreGenerico_KeyPress" & " Linea:" & Erl()))
        Unload Me
End Sub

Private Sub txtNumero_GotFocus()
On Error GoTo NotificaError
    pLimpia
    pSelTextBox txtNumero

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtNumero_GotFocus"))
    Unload Me
End Sub

Private Sub txtNumero_KeyDown(KeyCode As Integer, Shift As Integer)
1     On Error GoTo NotificaError
2         If KeyCode = vbKeyReturn Then
3             vlblnNoFocus = True
4             If fintCargaReq(Str(Val(txtNumero.Text))) = 0 Then
5                 txtNumero.Text = frsRegresaRs("select isnull(max(numNumRequisicion),0)+1 from IvRequisicionMaestro").Fields(0)
6                 If Me.ActiveControl.Name = "txtNumero" Then chkUrgente.SetFocus
7                 fintCargaReq (Str(Val(txtNumero.Text)))
8             Else
9                 pMuestra
10            End If
11        End If
12    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtNumero_KeyDown" & " Linea:" & Erl()))
        Unload Me
End Sub

Private Sub pMuestra()
1     On Error GoTo NotificaError
      Dim X As Long
      Dim rsEstatus As ADODB.Recordset
      Dim rs1 As ADODB.Recordset
      Dim lblnTipo As Boolean
      Dim strFormatString As String
      Dim intcontador As Integer
      Dim intManejos As Integer
      Dim rsCaja As New ADODB.Recordset
      Dim strCmbTipoRequisicion As String
          
2         vlblnConsulta = True
3         intManejos = 0
          
4         MyButton1.Enabled = True
5         With rsIvRequisicionMaestro

6             txtNumero.Text = !numnumRequisicion
7             txtFecha.Text = Format(!dtmFechaRequisicion, "dd/mmm/yyyy")
8             chkUrgente.Value = IIf(!bitUrgente = 1, 1, 0)
              
9             Set rsEstatus = frsRegresaRs("select vchestatusrequis from ivrequisicionmaestro where numnumrequisicion =" & txtNumero.Text)
10            If rsEstatus.RecordCount > 0 Then txtEstatus.Text = rsEstatus!vchEstatusRequis
              
              
              
11               cboDepartamentoSolicita.ListIndex = fintLocalizaCbo_new(cboDepartamentoSolicita, CStr(!smiCveDeptoRequis))
              
12               If cboTipoRequisicion.ListCount = 0 Then
13                  If lblnAlmacenConsigna Then
14                     cboTipoRequisicion.AddItem "PEDIDO SUGERIDO", 0
15                     cboTipoRequisicion.ListIndex = 0
16                  Else
17                     If !chrDestino = "R" Then
18                        cboTipoRequisicion.AddItem "REUBICACION", 0
19                     ElseIf !chrDestino = "D" Then
20                             cboTipoRequisicion.AddItem "SALIDA A DEPARTAMENTO", 0
21                     ElseIf !chrDestino = "A" Then
22                            cboTipoRequisicion.AddItem "ALMACEN ABASTECIMIENTO", 0
23                     ElseIf !chrDestino = "C" Then
24                            cboTipoRequisicion.AddItem "COMPRA - PEDIDO", 0
25                     End If
26                     If cboTipoRequisicion.ListCount > 0 Then cboTipoRequisicion.ListIndex = 0
27                  End If
28               Else
29                  If !chrDestino = "O" Then cboTipoRequisicion.AddItem "CONSIGNACION"
                              
30                  lblnTipo = False
31                  For X = 0 To cboTipoRequisicion.ListCount - 1
                        'If IIf(!chrDestino = "R", "REUBICACION", IIf(!chrDestino = "D", "SALIDA A DEPARTAMENTO", IIf(!chrDestino = "A", "ALMACEN ABASTECIMIENTO", IIf(!chrDestino = "O", "CONSIGNACION", "COMPRA - PEDIDO")))) = cboTipoRequisicion.List(X) Then
32                      If IIf(!chrDestino = "R", "REUBICACION", IIf(!chrDestino = "D", "SALIDA A DEPARTAMENTO", IIf(!chrDestino = "A", "ALMACEN ABASTECIMIENTO", IIf(!chrDestino = "O", "CONSIGNACION", IIf(!chrDestino = "U", "PEDIDO SUGERIDO", "COMPRA - PEDIDO"))))) = cboTipoRequisicion.List(X) Then ' JASM 20220125
33                         cboTipoRequisicion.ListIndex = X
34                         lblnTipo = True
35                         strCmbTipoRequisicion = cboTipoRequisicion.List(X)
36                      End If
37                  Next X
38                  If IIf(!chrDestino = "R", "REUBICACION", IIf(!chrDestino = "D", "SALIDA A DEPARTAMENTO", IIf(!chrDestino = "A", "ALMACEN ABASTECIMIENTO", IIf(!chrDestino = "O", "CONSIGNACION", IIf(!chrDestino = "U", "PEDIDO SUGERIDO", "COMPRA - PEDIDO"))))) <> strCmbTipoRequisicion Then       ' JASM 20220125
39                          cboTipoRequisicion.AddItem IIf(!chrDestino = "R", "REUBICACION", IIf(!chrDestino = "D", "SALIDA A DEPARTAMENTO", IIf(!chrDestino = "A", "ALMACEN ABASTECIMIENTO", IIf(!chrDestino = "O", "CONSIGNACION", IIf(!chrDestino = "U", "PEDIDO SUGERIDO", "COMPRA - PEDIDO"))))), 0
40                          cboTipoRequisicion.ListIndex = 0
41                          lblnTipo = True
42                  End If
43                  If lblnTipo = False Then
44                     If !chrDestino = "R" Then
45                        cboTipoRequisicion.AddItem "REUBICACION", 0
46                     ElseIf !chrDestino = "D" Then
47                            cboTipoRequisicion.AddItem "SALIDA A DEPARTAMENTO", 0
48                     ElseIf !chrDestino = "A" Then
49                            cboTipoRequisicion.AddItem "ALMACEN ABASTECIMIENTO", 0
50                     ElseIf !chrDestino = "C" Then
51                            cboTipoRequisicion.AddItem "COMPRA - PEDIDO", 0
52                     End If
53                     cboTipoRequisicion.ListIndex = 0
54                   End If
55               End If
              
56               If IsNull(!bitCompraDirecta) Then
57                  chkCompradirecta.Value = 0
58               Else
59                  chkCompradirecta.Value = !bitCompraDirecta
60               End If
61               chkCompradirecta.Enabled = False
62               If chkCompradirecta.Value = 1 And Not IsNull(!intcveproveedor) Then
63                    lblProveedor.Enabled = True
64                    cboProveedor.Enabled = True
65                    cboProveedor.ListIndex = fintLocalizaCbo_new(cboProveedor, CStr(!intcveproveedor))
66                    If cboProveedor.ListIndex = -1 Then
                          'el proveedor ya no está activo
67                        Set rs1 = frsEjecuta_SP(CStr(!intcveproveedor) & "|-1", "SP_COSELPROVEEDOR")
68                        If rs1.RecordCount <> 0 Then
69                            cboProveedor.AddItem (Trim(rs1!vchNombreComercial)), 1
70                            cboProveedor.ItemData(cboProveedor.NewIndex) = -1 * rs1!intcveproveedor
71                            cboProveedor.ListIndex = 1
72                        End If
73                    End If
74                    pMostrarProveedor (True)
75               Else
76                    lblProveedor.Enabled = False
77                    cboProveedor.Enabled = False
78                    If lblnSeleccionarProveedor Then
                          'If chkCompradirecta.Value = 1 Then cboProveedor.ListIndex = 0 Else cboProveedor.ListIndex = -1
79                        cboProveedor.ListIndex = 0
80                        pMostrarProveedor (True)
81                    Else
82                        pMostrarProveedor (False)
83                    End If
84               End If
                 
85               txtEmpleadoSolicito.Text = frsRegresaRs("select ltrim(rtrim(vchApellidoPaterno))||' '||ltrim(rtrim(vchApellidoMaterno))||' '||ltrim(rtrim(vchNombre)) Nombre from NoEmpleado where intCveEmpleado=" & Str(!intCveEmpleaRequis)).Fields(0)
              
86               cboAlmacenSurte.ListIndex = -1
                 
87               For X = 0 To cboAlmacenSurte.ListCount - 1
88                   If !smiCveDeptoAlmacen = cboAlmacenSurte.ItemData(X) Then
89                      cboAlmacenSurte.ListIndex = X
90                   End If
91               Next X
              
92               If cboAlmacenSurte.ListIndex = -1 Then
93                  If (!chrDestino = "A") Then ' Almacén General
                        ' Extrae el valor de vchProveedorAlmacenGeneral, si es diferente a "Vacio" o "Nulo" entonces permite Almacén General
94                      Set rs = frsRegresaRs("SELECT cp.vchNombre FROM coproveedor cp WHERE cp.intCveProveedor = " & vglngCveAlmacenGeneral)
95                      If (rs.State <> adStateClosed) Then
96                         If rs.RecordCount > 0 Then
97                            rs.MoveFirst
98                            If Not IsNull(rs!VCHNOMBRE) Then
99                                If (Trim(rs!VCHNOMBRE) <> "") Then
100                                  cboAlmacenSurte.AddItem (Trim(rs!VCHNOMBRE))
101                                  cboAlmacenSurte.ItemData(cboAlmacenSurte.NewIndex) = 0
102                                  cboAlmacenSurte.ListIndex = 0
103                               End If
104                           End If
105                        End If
106                     End If
107                 ElseIf (!chrDestino <> "U") Then
108                        cboAlmacenSurte.AddItem frsRegresaRs("select ltrim(rtrim(vchDescripcion)) Nombre from NoDepartamento where smiCveDepartamento=" & Str(!smiCveDeptoAlmacen)).Fields(0)
109                        cboAlmacenSurte.ItemData(cboAlmacenSurte.NewIndex) = !smiCveDeptoAlmacen
110                        cboAlmacenSurte.ListIndex = cboAlmacenSurte.NewIndex
111                 End If
112              End If
                  
                 'cajas de material
113              cboCajaMaterial.Enabled = False
114              lblCajaMaterial.Enabled = False
115              llngCveCajaMaterialRequi = -1
116              cboCajaMaterial.ListIndex = -1
117              pHabilitaArticulo True
118              If !chrDestino = "R" Then
119                   If Not IsNull(!intCajaMaterial) Then
120                        If !intCajaMaterial <> 0 Then
121                            llngCveCajaMaterialRequi = !intCajaMaterial
122                            cboCajaMaterial.Enabled = True
123                            lblCajaMaterial.Enabled = True
124                            cboCajaMaterial.ListIndex = fintLocalizaCbo_new(cboCajaMaterial, CStr(!intCajaMaterial))
125                            If cboCajaMaterial.ListIndex = -1 Then
126                                lblnRecargarCajas = True
127                                Set rsCaja = frsRegresaRs("select trim(vchDescripcion) caja from ExCajaMedicamentoMaterial where intCve = " & CStr(!intCajaMaterial))
128                                If rsCaja.RecordCount <> 0 Then
129                                    cboCajaMaterial.AddItem rsCaja!caja
130                                    cboCajaMaterial.ItemData(cboCajaMaterial.NewIndex) = !intCajaMaterial
131                                    cboCajaMaterial.ListIndex = cboCajaMaterial.NewIndex
132                                End If
133                            End If
134                        End If
135                   Else
136                       cboCajaMaterial.ListIndex = 0
137                   End If
138               End If
                 'la opcion de cerrar requisicion sólo es para requisiciones de tipo Reubicacion y Salida a departamento
                 'que esten pendietes,autorizadas, autorizadas parciales o surtidas parciales
                 'se habilita también para compra pedido !chrDestino = "C"
139              If !chrDestino = "R" Or !chrDestino = "D" Or !chrDestino = "C" Then
140                 If txtEstatus.Text <> "CANCELADA" And txtEstatus.Text <> "NO AUTORIZADA" And txtEstatus.Text <> IIf(!chrDestino = "C", "RECIBIDA", "SURTIDA") Then
141                     cmdCierraReq.Enabled = True
142                 Else
143                     cmdCierraReq.Enabled = False
144                 End If
145              Else
146                  cmdCierraReq.Enabled = False
147              End If
                 '--------------------------------------------------------------------------------------
                          
148              txtObservaciones.Text = IIf(IsNull(Trim(!VCHOBSERVACIONES)), "", Trim(!VCHOBSERVACIONES))
              
149              If !vchEstatusRequis = "CANCELADA" Or !vchEstatusRequis = "NO AUTORIZADA" Or (!chrDestino <> "C" And !vchEstatusRequis = "SURTIDA") Or (!chrDestino = "C" And !vchEstatusRequis = "RECIBIDA") Then
150                  txtObservaciones.Enabled = False
151              Else
152                  txtObservaciones.Enabled = True
153                  If !vchEstatusRequis = "PENDIENTE" Then
154                     lintSoloMaestro = 0
155                     cmdAgregarNuevo.Enabled = IIf(cboTipoRequisicion.List(cboTipoRequisicion.ListIndex) = "COMPRA - PEDIDO", True, False)
156                  Else
157                     lintSoloMaestro = 1
158                  End If
159              End If
              
160       End With
          
           'saber si el login tiene permiso para solicitar medicamentos controlados
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      '    If cboAlmacenSurte.ListCount > 0 And cboTipoRequisicion.ListCount > 0 Then
      '        vlblnControlado = False
      '        Set rs1 = frsRegresaRs("SELECT bitControlado " & _
      '                               "FROM IvRequisicionDepartamento " & _
      '                              "WHERE intNumeroLogin=" & Str(vglngNumeroLogin) & _
      '                             " AND smiCveDepartamento=" & cboAlmacenSurte.ItemData(cboAlmacenSurte.ListIndex) & _
      '                             " AND chrTipoRequisicion='" & IIf(Trim(cboTipoRequisicion.List(cboTipoRequisicion.ListIndex)) = "SALIDA A DEPARTAMENTO", "SD", IIf(Trim(cboTipoRequisicion.List(cboTipoRequisicion.ListIndex)) = "REUBICACION", "RE", IIf(Trim(cboTipoRequisicion.List(cboTipoRequisicion.ListIndex)) = "ALMACEN ABASTECIMIENTO", "AG", "CO"))) & "'")
      '        If rs1.RecordCount > 0 Then
      '           If ((rs1!bitControlado = 1) Or (rs1!bitControlado = True)) Then vlblnControlado = True
      '        End If
      '    End If
161       If cboAlmacenSurte.ListCount > 0 And cboTipoRequisicion.ListCount > 0 Then  ' JASM 20211223
162           vlblnControlado = False
163           vlStrSQL = "SELECT bitControlado " & _
                                     "FROM IvRequisicionDepartamento " & _
                                    "WHERE intNumeroLogin=" & Str(vglngNumeroLogin) & _
                                   " AND chrTipoRequisicion='" & IIf(Trim(cboTipoRequisicion.List(cboTipoRequisicion.ListIndex)) = "SALIDA A DEPARTAMENTO", "SD", IIf(Trim(cboTipoRequisicion.List(cboTipoRequisicion.ListIndex)) = "REUBICACION", "RE", IIf(Trim(cboTipoRequisicion.List(cboTipoRequisicion.ListIndex)) = "ALMACEN ABASTECIMIENTO", "AG", "CO"))) & "'"
164           If cboAlmacenSurte.ListIndex <> -1 Then
165              vlStrSQL = vlStrSQL & " AND smiCveDepartamento=" & cboAlmacenSurte.ItemData(cboAlmacenSurte.ListIndex)
166           End If
167           Set rs1 = frsRegresaRs(vlStrSQL)
168           If rs1.RecordCount > 0 Then
169              If ((rs1!bitControlado = 1) Or (rs1!bitControlado = True)) Then vlblnControlado = True
170           End If
171       End If
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
             
172       grdArticulos.Redraw = False
173       grdArticulos.Visible = False
          
174       Set rs = frsEjecuta_SP(txtNumero.Text, "Sp_IvSelRequisicionDetalle")
175       If rs.RecordCount <> 0 Then
              
176           intMaxManejos = rs!Manejos
177           pFormatoArticulos True
              
178           With rs
179               Do While Not .EOF

180                   If grdArticulos.Rows > 2 Or grdArticulos.TextMatrix(1, intColCveArticulo) <> "" Then
181                       grdArticulos.Rows = grdArticulos.Rows + 1
182                   End If
183                   grdArticulos.Row = grdArticulos.Rows - 1
                      
184                   grdArticulos.TextMatrix(grdArticulos.Row, intColCveArticulo) = IIf(IsNull(!chrcvearticulo), "", !chrcvearticulo)
185                   grdArticulos.TextMatrix(grdArticulos.Row, intColNombreComercial) = IIf(IsNull(!vchNombreComercial), "", !vchNombreComercial)
186                   grdArticulos.TextMatrix(grdArticulos.Row, intColCantidad) = IIf(IsNull(!IntCantidadSolicitada), "", !IntCantidadSolicitada)
187                   grdArticulos.TextMatrix(grdArticulos.Row, intColDescripcionUnidad) = IIf(IsNull(!Unidad), "", !Unidad)
188                   grdArticulos.TextMatrix(grdArticulos.Row, intColExistencia) = IIf(IsNull(!existencia), "", !existencia)
189                   grdArticulos.TextMatrix(grdArticulos.Row, intColEstado) = IIf(IsNull(!vchEstatusDetRequis), "", !vchEstatusDetRequis)
190                   grdArticulos.TextMatrix(grdArticulos.Row, intColUnidad) = IIf(IsNull(!alterna), "", !alterna)
191                   grdArticulos.TextMatrix(grdArticulos.Row, intColProveedor) = IIf(IsNull(!VCHNOMBRE), "", !VCHNOMBRE)
192                   grdArticulos.TextMatrix(grdArticulos.Row, intColFecha) = IIf(IsNull(!dtmfecharecepcion), "", !dtmfecharecepcion)
193                   grdArticulos.TextMatrix(grdArticulos.Row, intColCantidadProveedor) = IIf(IsNull(!smiCantidadRecep), "", !smiCantidadRecep)
194                   grdArticulos.TextMatrix(grdArticulos.Row, intColCosto) = IIf(IsNull(!mnyCostoEntrada), "", !mnyCostoEntrada)
                      
                            
                      
                      'Manejos
195                   For intcontador = intColManejo4 To intColManejo1 Step -1
                      
196                       If Not IsNull(!vchSimbolo) Then
197                           grdArticulos.Col = intcontador
198                           grdArticulos.Row = grdArticulos.Rows - 1
199                           grdArticulos.CellFontName = "Wingdings"
200                           grdArticulos.CellFontSize = 12
201                           grdArticulos.CellForeColor = CLng(!vchColor)
202                           grdArticulos.TextMatrix(grdArticulos.Row, intcontador) = !vchSimbolo
203                       End If
204                       .MoveNext
                          
205                       If .EOF Then
206                           .MovePrevious
207                           Exit For
208                       ElseIf !chrcvearticulo <> grdArticulos.TextMatrix(grdArticulos.Row, intColCveArticulo) Or Trim(!chrcvearticulo) = "NUEVO" Then
209                           .MovePrevious
210                           Exit For
211                       End If
                          
212                   Next intcontador
213                   .MoveNext
214               Loop
215           End With
              
216       End If
          
217       grdArticulos.Redraw = True
218       grdArticulos.Visible = True
          
219       If rsIvRequisicionMaestro!chrDestino = "C" Then
220           For lngRow = 1 To grdArticulos.Rows - 1
221               grdArticulos.Row = lngRow
222               For vlintCol = intColCveArticulo To grdArticulos.Cols - 1
223                   grdArticulos.Col = vlintCol
224                   If Trim(grdArticulos.TextMatrix(lngRow, intColCveArticulo)) = "NUEVO" Then
225                       grdArticulos.CellForeColor = &HC0C000
226                   Else
227                       grdArticulos.CellForeColor = vbBlack
228                   End If
229               Next vlintCol
230           Next lngRow
231       ElseIf rsIvRequisicionMaestro!chrDestino = "U" Then
232           For lngRow = 1 To grdArticulos.Rows - 1
233               grdArticulos.Row = lngRow
234               If Trim(grdArticulos.TextMatrix(lngRow, intColEstado)) = "ORDENADA" Then
235                   grdArticulos.TextMatrix(lngRow, intColEstado) = "PENDIENTE"
236               End If
237           Next lngRow
238       End If
          
239       pSumaCostos
240       If txtEstatus.Text = "PENDIENTE" Then
241           If lblnAlmacenConsigna Then
242               fraCabecera.Enabled = False
                  
243               fraFiltros.Enabled = False
244               fraArticulos.Enabled = True
                  'cmdCancelar.Enabled = True
245           Else
246               fraCabecera.Enabled = True
247               fraFiltros.Enabled = True
248               fraArticulos.Enabled = True
249               cmdAgregar.Enabled = True
                  'cmdCancelar.Enabled = True
250           End If
251       Else
252         fraCabecera.Enabled = False
            
253         fraFiltros.Enabled = False
254         cmdCancelar.Enabled = False
255         cmdAgregar.Enabled = False
256       End If
257       cmdBorrar.Enabled = False
          
258       If optOpcion(0).Enabled And optOpcion(0).Visible Then
259          optOpcion(0).SetFocus
260       End If
          
261       pHabilita True, True, True, True, True, False, True
          
262       rsEstatus.Close

263       If llngCveCajaMaterialRequi > 0 Then
264           fraCabecera.Enabled = True
265           fraFiltros.Enabled = False
266           cmdCancelar.Enabled = False
267           cmdAgregar.Enabled = False
268       ElseIf txtEstatus.Text = "PENDIENTE" Then
269          If lblnAlmacenConsigna Then
270             fraCabecera.Enabled = False
                
271             fraFiltros.Enabled = False
272             fraArticulos.Enabled = True
273          Else
274             fraCabecera.Enabled = True
275             fraFiltros.Enabled = True
276             fraArticulos.Enabled = True
277          End If
278       Else
279             fraCabecera.Enabled = False
                
280             fraFiltros.Enabled = False
281       End If
282       cmdBorrar.Enabled = False
283       If optOpcion(0).Enabled And optOpcion(0).Visible Then
284           optOpcion(0).SetFocus
285       End If

286   Exit Sub
NotificaError:
       Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pMuestra" & " Linea:" & Erl()))
       Unload Me
End Sub

Private Sub txtNumero_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    pValidaSoloNumero KeyAscii
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtNumero_KeyPress"))
    Unload Me
End Sub

Private Sub pCargaLocalizacion()
On Error GoTo NotificaError
    
    vlstrx = "" & _
    "select " & _
        "vchDescripcion Descripcion," & _
        "smiCveLocalizacion Clave " & _
    "From " & _
        "IvLocalizacion " & _
    "Where " & _
        "smiCveDepartamento = " & Trim(Str(cboDepartamentoSolicita.ItemData(cboDepartamentoSolicita.ListIndex))) & " " & _
    "Union " & _
    "select " & _
        "'<TODAS>' Descripcion," & _
        "0 Clave FROM dual"
    Set rs = frsRegresaRs(vlstrx)
    
    pLlenarCboRs_new cboLocalizacion, rs, 1, 0
    cboLocalizacion.ListIndex = -1

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCargaLocalizacion"))
    Unload Me
End Sub

Private Sub pCargaCajasMaterial()
1     On Error GoTo NotificaError
          
2         cboCajaMaterial.Clear
3         lblCajaMaterial.Enabled = False
4         cboCajaMaterial.Enabled = False
5         Set rs = frsRegresaRs("select cm.intCve id, (cm.VCHDESCRIPCION) descripcion from ExCajamedicamentoMaterial cm " & _
                                "inner join ExCajamedicaMaterialElemento cme on cm.INTCVE = cme.INTCVECAJAMED where cm.bitActivo = 1 " & _
                                "group by cm.intCve, cm.VCHDESCRIPCION order by descripcion ")
6         pLlenarCboRs_new cboCajaMaterial, rs, 0, 1
7         cboCajaMaterial.AddItem "<NINGUNA>", 0
8         cboCajaMaterial.ListIndex = -1
9         lblnRecargarCajas = False
          
10    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCargaCajasMaterial" & " Linea:" & Erl()))
End Sub

Private Sub pCargaFiltros()
1     On Error GoTo NotificaError
          
          'Familia
2         If Not optOpcion(0).Value Then
3             vlstrx = "" & _
              "select " & _
                  "vchDescripcion Descripcion," & _
                  "chrCveFamilia Clave " & _
              "From " & _
                  "IvFamilia " & _
              "Where " & _
                  "bitactivo = 1 and chrCveArtMedicamen = " & Trim(Str(IIf(optOpcion(1).Value, 0, IIf(optOpcion(2).Value, 1, 2)))) & " " & _
              "Union " & _
              "select " & _
                  "'<TODAS>' Descripcion," & _
                  "0 Clave FROM dual"
          
4             Set rs = frsRegresaRs(vlstrx)
5             pLlenarCboRs_new cboFamilia, rs, 1, 0
          
6         Else
7             cboFamilia.Clear
8             cboFamilia.AddItem "<TODAS>", 0
9             cboFamilia.ItemData(cboFamilia.NewIndex) = 0
10            cboFamilia.ListIndex = 0
              
11            cboSubfamilia.Clear
12            cboSubfamilia.AddItem "<TODAS>", 0
13            cboSubfamilia.ItemData(cboSubfamilia.NewIndex) = 0
14            cboSubfamilia.ListIndex = 0
15        End If
          
16        cboFamilia.ListIndex = 0 'aqui se corre lo necesario para llenar la subfamilia

17    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCargaFiltros" & " Linea:" & Erl()))
        Unload Me
End Sub

Private Sub txtNumero_LostFocus()
If Not vlblnNoFocus Then
   txtNumero_KeyDown 13, 0
End If
 vlblnNoFocus = False
End Sub

Private Sub txtObservaciones_GotFocus()
    If vlblnConsulta Then cmdGrabar.Enabled = True
End Sub

Private Sub txtObservaciones_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If fblnCanFocus(cmdGrabar) Then cmdGrabar.SetFocus
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If


End Sub
Private Sub pCargaDepartamentoSolicita()
Dim rs As New ADODB.Recordset
Dim lngOpcion As Long
    
    Select Case cgstrModulo
        Case "IV"
            lngOpcion = 2118
        Case "IM"
            lngOpcion = 2119
        Case "PV"
            lngOpcion = 2120
        Case "CP"
            lngOpcion = 2121
        Case "LA"
            lngOpcion = 2122
        Case "EX"
            lngOpcion = 2123
        Case "CN"
            lngOpcion = 2124
        Case "CC"
            lngOpcion = 2125
        Case "NO"
            lngOpcion = 2126
        Case "CA"
            lngOpcion = 2127
        Case "AD"
            lngOpcion = 2128
        Case "CO"
            lngOpcion = 2129
        Case "SI"
            lngOpcion = 2130
        Case "BS"
            lngOpcion = 2131
        Case "SE"
            lngOpcion = 2132
        Case "DI"
            lngOpcion = 2133
        Case "TS"
            lngOpcion = 2134
    End Select
    
    cboDepartamentoSolicita.Clear
    Set rs = frsRegresaRs("Select smiCveDepartamento, Trim(vchDescripcion) From NoDepartamento Where bitEstatus = 1 And tnyClaveEmpresa = " & CStr(vgintClaveEmpresaContable), adLockReadOnly, adOpenForwardOnly)
    If rs.RecordCount > 0 Then
       pLlenarCboRs_new cboDepartamentoSolicita, rs, 0, 1
    End If
    cboDepartamentoSolicita.ListIndex = fintLocalizaCbo_new(cboDepartamentoSolicita, CStr(vgintNumeroDepartamento))
        
    If fblnRevisaPermiso(vglngNumeroLogin, lngOpcion, "E", True) Then cboDepartamentoSolicita.Enabled = True
   
End Sub

Private Function fblnValidaNuevo() As Boolean
1     On Error GoTo NotificaError
      Dim lngContador As Long
          
2         fblnValidaNuevo = True
          
3         If llngDeptoRecibe = 0 Then
              'El usuario no tiene asignado el departamento que recibe la requisición de tipo Compra-Pedido
4             MsgBox SIHOMsg(869), vbInformation + vbOKOnly, "Mensaje"
5             fblnValidaNuevo = False
6             Exit Function
7         End If
          
8         For lngContador = 1 To grdArticulos.Rows - 1
9             If Trim(grdArticulos.TextMatrix(lngContador, intColCveArticulo)) = "NUEVO" Then
                  'Nombre Cantidad ó Unidad
10                If Trim(grdArticulos.TextMatrix(lngContador, intColNombreComercial)) = "" Or Val(grdArticulos.TextMatrix(lngContador, intColCantidad)) <= 0 Or Trim(grdArticulos.TextMatrix(lngContador, intColDescripcionUnidad)) = "" Then
11                   fblnValidaNuevo = False
                     'Capture los datos para el artículo nuevo
12                   MsgBox SIHOMsg(911), vbInformation + vbOKOnly, "Mensaje"
13                   Exit Function
14                End If
15            End If
16        Next lngContador
          
17    Exit Function
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnValidaNuevo" & " Linea:" & Erl()))
        Unload Me
End Function

Private Sub pEstatus()
1     On Error GoTo NotificaError
      Dim rsEstatusArticulo As New ADODB.Recordset
      Dim lngContador As Long
      Dim strSQL As String

2         If vlblnConsulta Then
3             rsIvRequisicionMaestro.Requery
              
4             If fintCargaReq(Str(Val(txtNumero.Text))) <> 0 Then
5                 txtEstatus.Text = Trim(rsIvRequisicionMaestro!vchEstatusRequis)
                  
6                 If Trim(rsIvRequisicionMaestro!vchEstatusRequis) <> "PENDIENTE" Then
7                     For lngContador = 1 To grdArticulos.Rows - 1
8                         strSQL = "Select vchEstatusDetRequis Estatus From IvRequisicionDetalle Where numNumRequisicion = " & Trim(txtNumero.Text) & " and chrCveArticulo = '" & Trim(grdArticulos.TextMatrix(lngContador, intColCveArticulo)) & "'"
9                         If Trim(grdArticulos.TextMatrix(lngContador, intColCveArticulo)) = "NUEVO" Then strSQL = strSQL & " And vchNombreArticuloNuevo ='" & Trim(grdArticulos.TextMatrix(lngContador, intColNombreComercial)) & "'"
                          
10                        Set rsEstatusArticulo = frsRegresaRs(strSQL, adLockOptimistic, adOpenStatic)
                          
11                        If rsEstatusArticulo.RecordCount > 0 Then grdArticulos.TextMatrix(lngContador, intColEstado) = Trim(rsEstatusArticulo!Estatus)
12                    Next lngContador
13                End If
14            End If
              
15            If txtEstatus.Text = "PENDIENTE" Then
16                If lblnAlmacenConsigna Then
17                    fraCabecera.Enabled = False
                      
18                    fraFiltros.Enabled = False
19                    fraArticulos.Enabled = True
                      'cmdCancelar.Enabled = True
20                Else
21                    fraCabecera.Enabled = True
22                    fraFiltros.Enabled = True
23                    fraArticulos.Enabled = True
                      'cmdCancelar.Enabled = True
24                End If
25            Else
26                lintSoloMaestro = 1
27                fraCabecera.Enabled = False
                  
28                fraFiltros.Enabled = False
                  'cmdCancelar.Enabled = False
29                pHabilita True, True, True, True, True, False, True
30            End If
31        End If
              
32    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pEstatus" & " Linea:" & Erl()))
End Sub

Private Function fintCargaReq(strCveReq As String) As Integer
Dim strSQL As String
    If IsNumeric(strCveReq) Then
        'se agregó codigo para que no se muestren las requisiciones de cargo a paciente
        strSQL = "select NUMNUMREQUISICION, SMICVEDEPTOREQUIS, INTCVEEMPLEAREQUIS, SMICVEDEPTOALMACEN, DTMFECHAREQUISICION," & _
        " DTMHORAREQUISICION, VCHESTATUSREQUIS, CHRDESTINO, NUMNUMCUENTA, BITURGENTE, NUMNUMREQUISREL," & _
        " DTMFECHAREQUISAUTORI, DTMHORAREQUISAUTORI, CHRTIPOPACIENTE, CHRAPLICACIONMED,INTNUMEROLOGIN, VCHOBSERVACIONES, " & _
        " SMICVEDEPTOGENERA, BITCOMPRADIRECTA, INTCVEPROVEEDOR, INTCAJAMATERIAL " & _
        " from IvRequisicionMaestro where chrdestino <> 'P' and smiCveDeptoGenera = " & vgintNumeroDepartamento & " and numNumRequisicion = " & strCveReq
        
        Set rsIvRequisicionMaestro = frsRegresaRs(strSQL, adLockOptimistic, adOpenDynamic)
        If Not rsIvRequisicionMaestro.EOF Then
            fintCargaReq = 1
        Else
            fintCargaReq = 0
        End If
    Else
         fintCargaReq = 0
    End If
End Function

Private Sub Unidades()
1     On Error GoTo NotificaError
      Dim blnArticuloSel As Boolean 'Para saber si existe un artículo seleccionado

2         blnArticuloSel = True
          
3         If cboNombreComercial.ListIndex = -1 Then
4             blnArticuloSel = False
5         Else
6             If cboNombreComercial.ItemData(cboNombreComercial.ListIndex) = 0 Then
7                 blnArticuloSel = False
8             End If
9         End If

10        lblCantidad.Enabled = chkPedirMaximo.Value = 0 And blnArticuloSel
11        txtCantidadSolicitada.Enabled = IIf(chkPedirMaximo.Value = 0 And blnArticuloSel, True, False)
          
12        If blnArticuloSel Then
13            If chkPedirMaximo Then
14                OptAlterna.Enabled = False
15                OptMinima.Enabled = False
16            Else
17                Set rs = frsEjecuta_SP(Trim(UCase(txtClave.Text)) & "|" & "|", "sp_IvSelArticulo")
18                If rs.RecordCount <> 0 Then
19                    rs.MoveFirst
20                    OptAlterna.Enabled = chkPedirMaximo.Value = 0 And blnArticuloSel
21                    OptMinima.Enabled = IIf(chkPedirMaximo.Value = 0 And blnArticuloSel And rs!UnidadMinima = rs!UnidadAlterna, False, True)
22                Else
23                    OptAlterna.Enabled = True
24                    OptMinima.Enabled = True
25                End If
26            End If
27        Else
28            OptAlterna.Enabled = False
29            OptMinima.Enabled = False
30        End If
              
31    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pUnidades" & " Linea:" & Erl()))
End Sub

Private Function fblnValidarSurtido() As Boolean
    On Error GoTo NotificaError
    Dim rs As ADODB.Recordset
    Dim rsPresupuesto As ADODB.Recordset
    Dim rsTempPresupuesto As ADODB.Recordset
    Dim rsAño As ADODB.Recordset
    Dim rsTemp As ADODB.Recordset
    Dim vlintValidacion As String 'Variable para saber si el presupuesto es mensual,bimensual, etc.
    Dim vlpresupuesto1, vlpresupuesto2, vlpresupuesto3, vlpresupuesto4, vlpresupuesto5, vlpresupuesto6, vlpresupuesto7, vlpresupuesto8, vlpresupuesto9, vlpresupuesto10, vlpresupuesto11, vlpresupuesto12 As String
    Dim vlstrMes1, vlstrMes2, vlstrMes3, vlstrMes4, vlstrMes5, vlstrMes6, vlstrMes7, vlstrMes8, vlstrMes9, vlstrMes10, vlstrMes11, vlstrMes12 As Long
    Dim vlstrMes As String 'Variable para obtener el mes de la fecha final del presupuesto activo
    Dim vlstrDia As String 'Variable para obtener el ultimo dia del mes de la fecha final del mes del presupuesto activo
    Dim vlstrMesInicial, vlstrMesfinal As String 'Variable para guardar fechas
    Dim vlstrFechaFinal As String 'Variable para obtener la fecha final del presupuesto y validar que aun este activo
    Dim vlstrFechaDia As String 'Variable para obtener el ultimo del dia mes de la fecha final
    Dim vlstrPresupuestoActivo As String 'Variable que nos indica cual es el año del presupuesto activo
    Dim vlstrFechaInicial As String 'Variable que nos indica cuando inicia el presupuesto activo y validar que ya se inicio el presupuesto
    Dim vlintContador As Integer 'Variable para un for
    Dim vllngCantidad As Long 'Variable que nos indica la cnatidad solicitada del articulo validado en unidad minima
    Dim vllngTotal As Long 'Variable para obtener la cantidad solicitada en las requisiciones - el presupuesto configurado en presupuesto a salidas a departamento - cantidad solicitada
    Dim vllngExiste As Long 'Variable que nos retorna si, 0:todos, 1:por departamentos -1:no encontrado
    Dim vlintRequisicion As Long 'Clave de la requisicion
    Dim vlintAño As Integer 'Variable para un for
    Dim vlintContenidoArticulo As Integer 'Variable que indica el contenido del articulo

    'Primero validamos si esta usando presupuestos
    Set rsTempPresupuesto = frsRegresaRs("Select * from EXPRESUPUESTOSALIDADEPTOINICIO", adLockOptimistic, adOpenDynamic)
    If rsTempPresupuesto.RecordCount = 0 Then
        fblnValidarSurtido = True
        Exit Function
    End If
    
    'Agarramos el presupuesto actual activo, primero obtenemos los años de la tabla
    Set rsAño = frsRegresaRs("select MIN(INTANO) as minAno, MAX(INTANO) as maxAno from EXPRESUPUESTOSALIDADEPTOINICIO", adLockOptimistic, adOpenDynamic)
    vlstrPresupuestoActivo = ""
    
    For vlintAño = rsAño!minAno To rsAño!maxAno
        Set rsTempPresupuesto = frsRegresaRs("select * from EXPRESUPUESTOSALIDADEPTOINICIO where intano =" & vlintAño, adLockOptimistic, adOpenDynamic)
        Do While Not rsTempPresupuesto.EOF
            'Con el presupuesto del año minimo podemos checar si ya se tiene que cerrar el presupuesto, obteniendo el año y el mes de inicio y sumandole 12 meses.
            vlstrFechaFinal = CDate(vlintAño & "/" & rsTempPresupuesto!INTMESINICIO & "/01")
            vlstrFechaFinal = DateAdd("m", 11, vlstrFechaFinal)
            vlstrFechaDia = Day(DateSerial(Year(vlstrFechaFinal), Month(vlstrFechaFinal) + 1, 0))
            vlstrFechaFinal = CDate(Year(vlstrFechaFinal) & "/" & Month(vlstrFechaFinal) & "/" & vlstrFechaDia)
            vlstrFechaInicial = CDate(vlintAño & "/" & rsTempPresupuesto!INTMESINICIO & "/01")
            'Si la fecha final es mayor a la del servidor quiere decir que es el presupuesto activo
            If vlstrFechaFinal > fdtmServerFecha Then
                'Pero puede que el presupuesto inicie en otra fecha que no es la del momento, entonces no lo valida y permite seguir la requisicion
                If vlstrFechaInicial > fdtmServerFecha Then
                    If rsTempPresupuesto!intCveDepartamento = cboDepartamentoSolicita.ItemData(cboDepartamentoSolicita.ListIndex) Then
                        vlstrPresupuestoActivo = ""
                        Exit For
                    End If
                Else
                    vlstrPresupuestoActivo = vlintAño 'Nuestro presupuesto activo
                    Exit For
                End If
            Else
                vlstrPresupuestoActivo = ""
            End If
            rsTempPresupuesto.MoveNext
        Loop
    Next vlintAño
    
    'Si esta vacio significa que no hay presupuestos activos
    If vlstrPresupuestoActivo = "" Then
        fblnValidarSurtido = True
        Exit Function
    End If
    'Cerramos porque lo usamos mas abajo si pasa la validación
    rsTempPresupuesto.Close
        
    vllngExiste = 1
    vgstrParametrosSP = cboDepartamentoSolicita.ItemData(cboDepartamentoSolicita.ListIndex) & "|" & vlstrPresupuestoActivo
    frsEjecuta_SP vgstrParametrosSP, "SP_IVPPTODEPTOEXISTE", True, vllngExiste
        
    If vllngExiste = -1 Then
        fblnValidarSurtido = True
        Exit Function
    End If
    
    'Si esta usando validamos que tenga presupuesto para el articulo validado por departamento
    fblnValidarSurtido = True
    vllngCantidad = txtCantidadSolicitada
    If OptAlterna.Value Then
        Set rsTemp = frsRegresaRs("Select * from ivarticulo where CHRCVEARTICULO = " & txtClave.Text, adLockOptimistic, adOpenDynamic)
        vllngCantidad = rsTemp!intContenido * txtCantidadSolicitada
        vlintContenidoArticulo = rsTemp!intContenido
    End If
    'Obtenemos el presupuesto que este como autorizado, solo puedenn haber registros autorizados de 1 solo año, cuando el presupuesto se finaliza se cambia al estado "CERRADO" por eso no puede repetirse con otros creados
    Set rs = frsRegresaRs("select EXPRESUPUESTOSALIDADEPTOINICIO.*, EXPRESUPUESTOSALIDADEPTODET.*,nvl(EXPRESUPUESTOSALIDADEPTO.INTCVEMPLEADOAUTORIZA,0) autorizaDepartamento from EXPRESUPUESTOSALIDADEPTOINICIO " & _
        "inner join EXPRESUPUESTOSALIDADEPTO on EXPRESUPUESTOSALIDADEPTOINICIO.INTCVECREACION = EXPRESUPUESTOSALIDADEPTO.INTCVECREACION " & _
        "inner join EXPRESUPUESTOSALIDADEPTODET on EXPRESUPUESTOSALIDADEPTO.INTCVEPRESUPUESTO = EXPRESUPUESTOSALIDADEPTODET.INTCLAVEPRESUPUESTO " & _
        "inner join ivarticulo on ivarticulo.INTIDARTICULO = EXPRESUPUESTOSALIDADEPTODET.INTIDARTICULO " & _
        "where VCHESTADO = 'AUTORIZADO' and EXPRESUPUESTOSALIDADEPTO.INTCVEDEPARTAMENTO = " & cboDepartamentoSolicita.ItemData(cboDepartamentoSolicita.ListIndex) & " and ivarticulo.CHRCVEARTICULO = " & txtClave.Text & " and EXPRESUPUESTOSALIDADEPTODET.BITORIGINAL = 0", adLockOptimistic, adOpenDynamic)
        
    'Si no encuentra quiere decir que no hay presupuestos para ese articulo
    If Not rs.RecordCount > 0 Then
        fblnValidarSurtido = False
            MsgBox SIHOMsg(1643), vbOKOnly + vbInformation, "Mensaje"
        Exit Function
    Else
        If rs!autorizaDepartamento = 0 Then
            fblnValidarSurtido = False
            MsgBox "El presupuesto para el departamento no está autorizado.", vbOKOnly + vbInformation, "Mensaje"
            Exit Function
        End If
    End If
          
    vlintValidacion = rs!INTVALIDACION
    vlstrMesInicial = Format(CDate(rs!intAno & "/" & rs!INTMESINICIO & "/01"), "YYYY/MM/DD")
    vlstrMesfinal = Format(DateAdd("m", 11, vlstrMesInicial), "YYYY/MM/DD")
    vlstrMes = Month(vlstrMesfinal)
    vlstrMes = IIf((vlstrMes) < 10, "0" & vlstrMes, vlstrMes)
    vlstrDia = Day(DateSerial(rs!intAno, vlstrMes + 1, 0))
    vlstrMesfinal = Format(Format(Year(vlstrMesfinal) & "/" & vlstrMes & "/" & vlstrDia), "YYYY/MM/DD")
    
    If vlblnConsulta Then
        vlintRequisicion = txtNumero
    Else
        vlintRequisicion = 0
    End If
    
    Set rsPresupuesto = frsEjecuta_SP(txtClave.Text & "|" & cboDepartamentoSolicita.ItemData(cboDepartamentoSolicita.ListIndex) & "|" & vlstrMesInicial & "|" & vlstrMesfinal & "|" & vlintRequisicion, "SP_EXSELARTICULOSPPTOREQ")
    
    If rsPresupuesto.RecordCount <> 0 Then
            '19 al 30 porque son las columnas donde te indica en que  mes inicia el presupuesto
        For vlintContador = 18 To 29
            vlstrMes = "mes" & vlintContador
            'Hacemos un case para obtener la cantidad solicitada en el mes seleccionado de las requisiciones
            Select Case rsPresupuesto.Fields(vlintContador)
                Case "01"
                    vlstrMes1 = rsPresupuesto.Fields.Item("TOTAL" & rsPresupuesto.Fields(vlintContador).Name)
                Case "02"
                    vlstrMes2 = rsPresupuesto.Fields.Item("TOTAL" & rsPresupuesto.Fields(vlintContador).Name)
                Case "03"
                    vlstrMes3 = rsPresupuesto.Fields.Item("TOTAL" & rsPresupuesto.Fields(vlintContador).Name)
                Case "04"
                    vlstrMes4 = rsPresupuesto.Fields.Item("TOTAL" & rsPresupuesto.Fields(vlintContador).Name)
                Case "05"
                    vlstrMes5 = rsPresupuesto.Fields.Item("TOTAL" & rsPresupuesto.Fields(vlintContador).Name)
                Case "06"
                    vlstrMes6 = rsPresupuesto.Fields.Item("TOTAL" & rsPresupuesto.Fields(vlintContador).Name)
                Case "07"
                    vlstrMes7 = rsPresupuesto.Fields.Item("TOTAL" & rsPresupuesto.Fields(vlintContador).Name)
                Case "08"
                    vlstrMes8 = rsPresupuesto.Fields.Item("TOTAL" & rsPresupuesto.Fields(vlintContador).Name)
                Case "09"
                    vlstrMes9 = rsPresupuesto.Fields.Item("TOTAL" & rsPresupuesto.Fields(vlintContador).Name)
                Case "10"
                    vlstrMes10 = rsPresupuesto.Fields.Item("TOTAL" & rsPresupuesto.Fields(vlintContador).Name)
                Case "11"
                    vlstrMes11 = rsPresupuesto.Fields.Item("TOTAL" & rsPresupuesto.Fields(vlintContador).Name)
                Case "12"
                    vlstrMes12 = rsPresupuesto.Fields.Item("TOTAL" & rsPresupuesto.Fields(vlintContador).Name)
            End Select
        Next vlintContador
    Else
        vlstrMes1 = 0
        vlstrMes2 = 0
        vlstrMes3 = 0
        vlstrMes4 = 0
        vlstrMes5 = 0
        vlstrMes6 = 0
        vlstrMes7 = 0
        vlstrMes8 = 0
        vlstrMes9 = 0
        vlstrMes10 = 0
        vlstrMes11 = 0
        vlstrMes12 = 0
    End If
    
    'Obtenemos el mes a validar
    vlstrMes = Month(CDate(txtFecha.Text))
    'Checamos si es mensual,bimensual, etc.
    
    If vlintValidacion = 1 Then
        Select Case vlstrMes
            Case "1"
                vllngTotal = (rs!INTCANTIDADPRESUPUESTOMES1) - vlstrMes1
            Case "2"
                vllngTotal = (rs!INTCANTIDADPRESUPUESTOMES2) - vlstrMes2
            Case "3"
                vllngTotal = (rs!INTCANTIDADPRESUPUESTOMES3) - vlstrMes3
            Case "4"
                vllngTotal = (rs!INTCANTIDADPRESUPUESTOMES4) - vlstrMes4
            Case "5"
                vllngTotal = (rs!INTCANTIDADPRESUPUESTOMES5) - vlstrMes5
            Case "6"
                vllngTotal = (rs!INTCANTIDADPRESUPUESTOMES6) - vlstrMes6
            Case "7"
                vllngTotal = (rs!INTCANTIDADPRESUPUESTOMES7) - vlstrMes7
            Case "8"
                vllngTotal = (rs!INTCANTIDADPRESUPUESTOMES8) - vlstrMes8
            Case "9"
                vllngTotal = (rs!INTCANTIDADPRESUPUESTOMES9) - vlstrMes9
            Case "10"
                vllngTotal = (rs!INTCANTIDADPRESUPUESTOMES10) - vlstrMes10
            Case "11"
                vllngTotal = (rs!INTCANTIDADPRESUPUESTOMES11) - vlstrMes11
            Case "12"
                vllngTotal = (rs!INTCANTIDADPRESUPUESTOMES12) - vlstrMes12
        End Select
        
        vllngTotal = vllngTotal - vllngCantidad 'Total mensual
        
    ElseIf vlintValidacion = 2 Then
        vlpresupuesto1 = vlstrMes1 + vlstrMes2
        vlpresupuesto2 = vlstrMes3 + vlstrMes4
        vlpresupuesto3 = vlstrMes5 + vlstrMes6
        vlpresupuesto4 = vlstrMes7 + vlstrMes8
        vlpresupuesto5 = vlstrMes9 + vlstrMes10
        vlpresupuesto6 = vlstrMes11 + vlstrMes12
        Select Case vlstrMes
            Case "1"
                vllngTotal = (rs!INTCANTIDADPRESUPUESTOMES1 + rs!INTCANTIDADPRESUPUESTOMES2) - vlpresupuesto1
            Case "2"
                vllngTotal = (rs!INTCANTIDADPRESUPUESTOMES1 + rs!INTCANTIDADPRESUPUESTOMES2) - vlpresupuesto1
            Case "3"
                vllngTotal = (rs!INTCANTIDADPRESUPUESTOMES3 + rs!INTCANTIDADPRESUPUESTOMES4) - vlpresupuesto2
            Case "4"
                vllngTotal = (rs!INTCANTIDADPRESUPUESTOMES3 + rs!INTCANTIDADPRESUPUESTOMES4) - vlpresupuesto2
            Case "5"
                vllngTotal = (rs!INTCANTIDADPRESUPUESTOMES5 + rs!INTCANTIDADPRESUPUESTOMES6) - vlpresupuesto3
            Case "6"
                vllngTotal = (rs!INTCANTIDADPRESUPUESTOMES5 + rs!INTCANTIDADPRESUPUESTOMES6) - vlpresupuesto3
            Case "7"
                vllngTotal = (rs!INTCANTIDADPRESUPUESTOMES7 + rs!INTCANTIDADPRESUPUESTOMES8) - vlpresupuesto4
            Case "8"
                vllngTotal = (rs!INTCANTIDADPRESUPUESTOMES7 + rs!INTCANTIDADPRESUPUESTOMES8) - vlpresupuesto4
            Case "9"
                vllngTotal = (rs!INTCANTIDADPRESUPUESTOMES9 + rs!INTCANTIDADPRESUPUESTOMES10) - vlpresupuesto5
            Case "10"
                vllngTotal = (rs!INTCANTIDADPRESUPUESTOMES9 + rs!INTCANTIDADPRESUPUESTOMES10) - vlpresupuesto5
            Case "11"
                vllngTotal = (rs!INTCANTIDADPRESUPUESTOMES11 + rs!INTCANTIDADPRESUPUESTOMES12) - vlpresupuesto6
            Case "12"
                vllngTotal = (rs!INTCANTIDADPRESUPUESTOMES11 + rs!INTCANTIDADPRESUPUESTOMES12) - vlpresupuesto6
        End Select
        
        vllngTotal = vllngTotal - vllngCantidad 'Total bimensual
        
    ElseIf vlintValidacion = 3 Then
        vlpresupuesto1 = vlstrMes1 + vlstrMes2 + vlstrMes3
        vlpresupuesto2 = vlstrMes4 + vlstrMes5 + vlstrMes6
        vlpresupuesto3 = vlstrMes7 + vlstrMes8 + vlstrMes9
        vlpresupuesto4 = vlstrMes10 + vlstrMes11 + vlstrMes12
        Select Case vlstrMes
            Case "1"
                vllngTotal = (rs!INTCANTIDADPRESUPUESTOMES1 + rs!INTCANTIDADPRESUPUESTOMES2 + rs!INTCANTIDADPRESUPUESTOMES3) - vlpresupuesto1
            Case "2"
                vllngTotal = (rs!INTCANTIDADPRESUPUESTOMES1 + rs!INTCANTIDADPRESUPUESTOMES2 + rs!INTCANTIDADPRESUPUESTOMES3) - vlpresupuesto1
            Case "3"
                vllngTotal = (rs!INTCANTIDADPRESUPUESTOMES1 + rs!INTCANTIDADPRESUPUESTOMES2 + rs!INTCANTIDADPRESUPUESTOMES3) - vlpresupuesto1
            Case "4"
                vllngTotal = (rs!INTCANTIDADPRESUPUESTOMES4 + rs!INTCANTIDADPRESUPUESTOMES5 + rs!INTCANTIDADPRESUPUESTOMES6) - vlpresupuesto2
            Case "5"
                vllngTotal = (rs!INTCANTIDADPRESUPUESTOMES4 + rs!INTCANTIDADPRESUPUESTOMES5 + rs!INTCANTIDADPRESUPUESTOMES6) - vlpresupuesto2
            Case "6"
                vllngTotal = (rs!INTCANTIDADPRESUPUESTOMES4 + rs!INTCANTIDADPRESUPUESTOMES5 + rs!INTCANTIDADPRESUPUESTOMES6) - vlpresupuesto2
            Case "7"
                vllngTotal = (rs!INTCANTIDADPRESUPUESTOMES7 + rs!INTCANTIDADPRESUPUESTOMES8 + rs!INTCANTIDADPRESUPUESTOMES9) - vlpresupuesto3
            Case "8"
                vllngTotal = (rs!INTCANTIDADPRESUPUESTOMES7 + rs!INTCANTIDADPRESUPUESTOMES8 + rs!INTCANTIDADPRESUPUESTOMES9) - vlpresupuesto3
            Case "9"
                vllngTotal = (rs!INTCANTIDADPRESUPUESTOMES7 + rs!INTCANTIDADPRESUPUESTOMES8 + rs!INTCANTIDADPRESUPUESTOMES9) - vlpresupuesto3
            Case "10"
                vllngTotal = (rs!INTCANTIDADPRESUPUESTOMES10 + rs!INTCANTIDADPRESUPUESTOMES11 + rs!INTCANTIDADPRESUPUESTOMES12) - vlpresupuesto4
            Case "11"
                vllngTotal = (rs!INTCANTIDADPRESUPUESTOMES10 + rs!INTCANTIDADPRESUPUESTOMES11 + rs!INTCANTIDADPRESUPUESTOMES12) - vlpresupuesto4
            Case "12"
                vllngTotal = (rs!INTCANTIDADPRESUPUESTOMES10 + rs!INTCANTIDADPRESUPUESTOMES11 + rs!INTCANTIDADPRESUPUESTOMES12) - vlpresupuesto4
        End Select
        
        vllngTotal = vllngTotal - vllngCantidad 'Total Trimestral
        
    ElseIf vlintValidacion = 6 Then
        vlpresupuesto1 = vlstrMes1 + vlstrMes2 + vlstrMes3 + vlstrMes4 + vlstrMes5 + vlstrMes6
        vlpresupuesto2 = vlstrMes7 + vlstrMes8 + vlstrMes9 + vlstrMes10 + vlstrMes11 + vlstrMes12
        Select Case vlstrMes
            Case "1"
                vllngTotal = (rs!INTCANTIDADPRESUPUESTOMES1 + rs!INTCANTIDADPRESUPUESTOMES2 + rs!INTCANTIDADPRESUPUESTOMES3 + rs!INTCANTIDADPRESUPUESTOMES4 + rs!INTCANTIDADPRESUPUESTOMES5 + rs!INTCANTIDADPRESUPUESTOMES6) - vlpresupuesto1
            Case "2"
                vllngTotal = (rs!INTCANTIDADPRESUPUESTOMES1 + rs!INTCANTIDADPRESUPUESTOMES2 + rs!INTCANTIDADPRESUPUESTOMES3 + rs!INTCANTIDADPRESUPUESTOMES4 + rs!INTCANTIDADPRESUPUESTOMES5 + rs!INTCANTIDADPRESUPUESTOMES6) - vlpresupuesto1
            Case "3"
                vllngTotal = (rs!INTCANTIDADPRESUPUESTOMES1 + rs!INTCANTIDADPRESUPUESTOMES2 + rs!INTCANTIDADPRESUPUESTOMES3 + rs!INTCANTIDADPRESUPUESTOMES4 + rs!INTCANTIDADPRESUPUESTOMES5 + rs!INTCANTIDADPRESUPUESTOMES6) - vlpresupuesto1
            Case "4"
                vllngTotal = (rs!INTCANTIDADPRESUPUESTOMES1 + rs!INTCANTIDADPRESUPUESTOMES2 + rs!INTCANTIDADPRESUPUESTOMES3 + rs!INTCANTIDADPRESUPUESTOMES4 + rs!INTCANTIDADPRESUPUESTOMES5 + rs!INTCANTIDADPRESUPUESTOMES6) - vlpresupuesto1
            Case "5"
                vllngTotal = (rs!INTCANTIDADPRESUPUESTOMES1 + rs!INTCANTIDADPRESUPUESTOMES2 + rs!INTCANTIDADPRESUPUESTOMES3 + rs!INTCANTIDADPRESUPUESTOMES4 + rs!INTCANTIDADPRESUPUESTOMES5 + rs!INTCANTIDADPRESUPUESTOMES6) - vlpresupuesto1
            Case "6"
                vllngTotal = (rs!INTCANTIDADPRESUPUESTOMES1 + rs!INTCANTIDAPRESUPUESTOMES2 + rs!INTCANTIDADPRESUPUESTOMES3 + rs!INTCANTIDADPRESUPUESTOMES4 + rs!INTCANTIDADPRESUPUESTOMES5 + rs!INTCANTIDADPRESUPUESTOMES6) - vlpresupuesto1
            Case "7"
                vllngTotal = (rs!INTCANTIDADPRESUPUESTOMES7 + rs!INTCANTIDADPRESUPUESTOMES8 + rs!INTCANTIDADPRESUPUESTOMES9 + rs!INTCANTIDADPRESUPUESTOMES10 + rs!INTCANTIDADPRESUPUESTOMES11 + rs!INTCANTIDADPRESUPUESTOMES12) - vlpresupuesto2
            Case "8"
                vllngTotal = (rs!INTCANTIDADPRESUPUESTOMES7 + rs!INTCANTIDADPRESUPUESTOMES8 + rs!INTCANTIDADPRESUPUESTOMES9 + rs!INTCANTIDADPRESUPUESTOMES10 + rs!INTCANTIDADPRESUPUESTOMES11 + rs!INTCANTIDADPRESUPUESTOMES12) - vlpresupuesto2
            Case "9"
                vllngTotal = (rs!INTCANTIDADPRESUPUESTOMES7 + rs!INTCANTIDADPRESUPUESTOMES8 + rs!INTCANTIDADPRESUPUESTOMES9 + rs!INTCANTIDADPRESUPUESTOMES10 + rs!INTCANTIDADPRESUPUESTOMES11 + rs!INTCANTIDADPRESUPUESTOMES12) - vlpresupuesto2
            Case "10"
                vllngTotal = (rs!INTCANTIDADPRESUPUESTOMES7 + rs!INTCANTIDADPRESUPUESTOMES8 + rs!INTCANTIDADPRESUPUESTOMES9 + rs!INTCANTIDADPRESUPUESTOMES10 + rs!INTCANTIDADPRESUPUESTOMES11 + rs!INTCANTIDADPRESUPUESTOMES12) - vlpresupuesto2
            Case "11"
                vllngTotal = (rs!INTCANTIDADPRESUPUESTOMES7 + rs!INTCANTIDADPRESUPUESTOMES8 + rs!INTCANTIDADPRESUPUESTOMES9 + rs!INTCANTIDADPRESUPUESTOMES10 + rs!INTCANTIDADPRESUPUESTOMES11 + rs!INTCANTIDADPRESUPUESTOMES12) - vlpresupuesto2
            Case "12"
            vllngTotal = (rs!INTCANTIDADPRESUPUESTOMES7 + rs!INTCANTIDADPRESUPUESTOMES8 + rs!INTCANTIDADPRESUPUESTOMES9 + rs!INTCANTIDADPRESUPUESTOMES10 + rs!INTCANTIDADPRESUPUESTOMES11 + rs!INTCANTIDADPRESUPUESTOMES12) - vlpresupuesto2
        End Select
        
        vllngTotal = vllngTotal - txtCantidadSolicitada.Text 'Total Semestral
        
    ElseIf vlintValidacion = 12 Then
        vlpresupuesto1 = vlstrMes1 + vlstrMes2 + vlstrMes3 + vlstrMes4 + vlstrMes5 + vlstrMes6 + vlstrMes7 + vlstrMes8 + vlstrMes9 + vlstrMes10 + vlstrMes11 + vlstrMes12
        vllngTotal = (rs!INTCANTIDADPRESUPUESTOMES1 + rs!INTCANTIDADPRESUPUESTOMES2 + rs!INTCANTIDADPRESUPUESTOMES3 + rs!INTCANTIDADPRESUPUESTOMES4 + rs!INTCANTIDADPRESUPUESTOMES5 + rs!INTCANTIDADPRESUPUESTOMES6 + rs!INTCANTIDADPRESUPUESTOMES7 + rs!INTCANTIDADPRESUPUESTOMES8 + rs!INTCANTIDADPRESUPUESTOMES9 + rs!INTCANTIDADPRESUPUESTOMES10 + rs!INTCANTIDADPRESUPUESTOMES11 + rs!INTCANTIDADPRESUPUESTOMES12) - vlpresupuesto1
        vllngTotal = vllngTotal - vllngCantidad 'Total Anual
    End If
    
    If Not vllngTotal > -1 Then
        If OptAlterna.Value Then
            vllngTotal = (vllngTotal * -1) / vlintContenidoArticulo
        Else
            vllngTotal = (vllngTotal * -1)
        End If
        MsgBox SIHOMsg(1642) & " " & vllngTotal, vbOKOnly + vbInformation, "Mensaje"
        fblnValidarSurtido = False
        Exit Function
    End If
  
   Exit Function
NotificaError:
       Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnValidarSurtido" & " Linea:" & Erl()))
       Unload Me
End Function

