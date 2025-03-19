VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.1#0"; "MyCommandButton.ocx"
Object = "{FF14BD24-9F8A-41E3-B5B8-7F0D45EE9F16}#15.0#0"; "hsflatcontrols.ocx"
Begin VB.Form frmRequisicionCargoPac 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Requisiciones con cargo a paciente"
   ClientHeight    =   9570
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   12450
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   9570
   ScaleWidth      =   12450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab sstObj 
      Height          =   9600
      Left            =   -255
      TabIndex        =   50
      Top             =   0
      Width           =   13140
      _ExtentX        =   23178
      _ExtentY        =   16933
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   2
      TabCaption(0)   =   "cdlTest.ShowPrinter"
      TabPicture(0)   =   "frmRequisicionCargoPac.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cdlTest"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraTab0"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Busqueda"
      TabPicture(1)   =   "frmRequisicionCargoPac.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraTab1"
      Tab(1).ControlCount=   1
      Begin VB.Frame fraTab0 
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
         Height          =   9855
         Left            =   120
         TabIndex        =   51
         Top             =   -250
         Width           =   12900
         Begin VB.Frame fraIntExt 
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2160
            TabIndex        =   10
            Top             =   2180
            Width           =   4200
            Begin VB.OptionButton optInterno 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "&Interno"
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
               Left            =   90
               TabIndex        =   11
               Tag             =   "0"
               Top             =   50
               Width           =   975
            End
            Begin VB.OptionButton optExterno 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "&Externo"
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
               Left            =   1320
               TabIndex        =   12
               Tag             =   "1"
               Top             =   50
               Width           =   1020
            End
            Begin VB.OptionButton optAmbulatorio 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Am&bulatorio"
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
               Left            =   2520
               TabIndex        =   13
               Tag             =   "1"
               Top             =   50
               Width           =   1560
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
            Height          =   1635
            Left            =   240
            TabIndex        =   0
            Top             =   240
            Width           =   12240
            Begin VB.TextBox txtEmpleado 
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
               Left            =   2040
               Locked          =   -1  'True
               TabIndex        =   5
               ToolTipText     =   "Empleado que solicita la requisición"
               Top             =   1110
               Width           =   3975
            End
            Begin VB.TextBox txtEstatusReq 
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
               Left            =   7770
               Locked          =   -1  'True
               TabIndex        =   6
               ToolTipText     =   "El estatus de la requisición"
               Top             =   300
               Width           =   4350
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
               Left            =   5000
               TabIndex        =   3
               ToolTipText     =   "La requisición es urgente o no"
               Top             =   360
               Width           =   1050
            End
            Begin VB.TextBox txtFecha 
               Alignment       =   2  'Center
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
               Left            =   3645
               Locked          =   -1  'True
               TabIndex        =   2
               ToolTipText     =   "Fecha de la requisición"
               Top             =   300
               Width           =   1305
            End
            Begin VB.TextBox txtNumReq 
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
               Left            =   2040
               TabIndex        =   1
               ToolTipText     =   "Número de requisición"
               Top             =   300
               Width           =   840
            End
            Begin HSFlatControls.MyCombo cboAlmacenSurte 
               Height          =   375
               Left            =   7770
               TabIndex        =   8
               ToolTipText     =   "Almacén que surtirá la requisición ó realizará la compra"
               Top             =   1110
               Width           =   4350
               _ExtentX        =   7673
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
            Begin HSFlatControls.MyCombo cboTipoReq 
               Height          =   375
               Left            =   7770
               TabIndex        =   7
               ToolTipText     =   "Tipo de requisición"
               Top             =   700
               Width           =   4350
               _ExtentX        =   7673
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
               Left            =   2040
               TabIndex        =   4
               ToolTipText     =   "Departamento que solicita"
               Top             =   700
               Width           =   3975
               _ExtentX        =   7011
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
               Left            =   6120
               TabIndex        =   52
               Top             =   1170
               Width           =   1260
            End
            Begin VB.Label lblTipo 
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
               Left            =   6120
               TabIndex        =   53
               Top             =   760
               Width           =   1155
            End
            Begin VB.Label lblEstatus 
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
               Left            =   6120
               TabIndex        =   54
               Top             =   360
               Width           =   1020
            End
            Begin VB.Label lblEmpleado 
               BackColor       =   &H80000005&
               Caption         =   "Empleado solicita"
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
               TabIndex        =   55
               Top             =   1170
               Width           =   1245
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
               TabIndex        =   56
               Top             =   760
               Width           =   1680
            End
            Begin VB.Label lblFecha 
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
               Left            =   3000
               TabIndex        =   57
               Top             =   360
               Width           =   615
            End
            Begin VB.Label lblNumReq 
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
               TabIndex        =   58
               Top             =   360
               Width           =   1095
            End
         End
         Begin VB.Frame fraBuscaArticulo 
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
            Height          =   7090
            Left            =   240
            TabIndex        =   9
            Top             =   1850
            Width           =   12240
            Begin MSMask.MaskEdBox txtEdit 
               Height          =   390
               Left            =   3285
               TabIndex        =   59
               Top             =   4320
               Visible         =   0   'False
               Width           =   1410
               _ExtentX        =   2487
               _ExtentY        =   688
               _Version        =   393216
               BorderStyle     =   0
               Appearance      =   0
               PromptInclude   =   0   'False
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
            Begin MyCommandButton.MyButton cmdCancelar 
               Height          =   495
               Left            =   11110
               TabIndex        =   34
               ToolTipText     =   "Cancelar artículo"
               Top             =   6290
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
               Picture         =   "frmRequisicionCargoPac.frx":0038
               BackColorDown   =   -2147483643
               BackColorOver   =   -2147483633
               BackColorFocus  =   -2147483633
               BackColorDisabled=   -2147483633
               BorderColor     =   -2147483627
               TransparentColor=   16777215
               Caption         =   ""
               DepthEvent      =   1
               PictureDisabled =   "frmRequisicionCargoPac.frx":074C
               PictureAlignment=   5
               PictureDisabledEffect=   0
               ShowFocus       =   -1  'True
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
               Left            =   2040
               TabIndex        =   21
               ToolTipText     =   "Nombre genérico del artículo"
               Top             =   1110
               Width           =   3975
            End
            Begin VB.TextBox txtSubfamilia 
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
               Left            =   2040
               Locked          =   -1  'True
               TabIndex        =   23
               ToolTipText     =   "Nombre de la subfamilia a la que pertenece el artículo"
               Top             =   1920
               Width           =   3975
            End
            Begin VB.TextBox txtFamilia 
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
               Left            =   2040
               Locked          =   -1  'True
               TabIndex        =   22
               ToolTipText     =   "Nombre de la familia a la que pertenece el artículo"
               Top             =   1510
               Width           =   3975
            End
            Begin VB.TextBox txtIdArticulo 
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
               Left            =   10870
               TabIndex        =   31
               Top             =   2330
               Visible         =   0   'False
               Width           =   1230
            End
            Begin VB.TextBox txtUnidad 
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
               Left            =   8010
               Locked          =   -1  'True
               TabIndex        =   26
               Top             =   1920
               Width           =   4095
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
               Left            =   8010
               TabIndex        =   25
               ToolTipText     =   "Código de barras del artículo"
               Top             =   1520
               Width           =   4095
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
               Left            =   8010
               Locked          =   -1  'True
               TabIndex        =   24
               ToolTipText     =   "Clave del artículo"
               Top             =   1110
               Width           =   4095
            End
            Begin VB.Frame fraTipoArticulo 
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
               Height          =   375
               Left            =   6600
               TabIndex        =   15
               Top             =   720
               Width           =   3975
               Begin VB.OptionButton optMedicamento 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  Caption         =   "&Medicamento"
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
                  Left            =   2280
                  TabIndex        =   18
                  Tag             =   "1"
                  ToolTipText     =   "Sólo medicamentos"
                  Top             =   50
                  Width           =   1665
               End
               Begin VB.OptionButton optArticulo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  Caption         =   "&Artículo"
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
                  Left            =   1080
                  TabIndex        =   17
                  Tag             =   "0"
                  ToolTipText     =   "Sólo artículos"
                  Top             =   50
                  Width           =   1095
               End
               Begin VB.OptionButton optTodos 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  Caption         =   "&Todos"
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
                  TabIndex        =   16
                  Tag             =   "1"
                  ToolTipText     =   "Todos los artículos y medicamentos"
                  Top             =   50
                  Width           =   1020
               End
            End
            Begin VB.Frame fraUnidad 
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
               Height          =   375
               Left            =   6600
               TabIndex        =   60
               Top             =   2360
               Width           =   4095
               Begin VB.OptionButton optAlterna 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  Caption         =   "Unidad alterna"
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
                  TabIndex        =   29
                  ToolTipText     =   "Manejar unidad venta"
                  Top             =   50
                  Width           =   1755
               End
               Begin VB.OptionButton optMinima 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  Caption         =   "Unidad mínima"
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
                  Left            =   2040
                  TabIndex        =   30
                  ToolTipText     =   "Manejar unidad mínima"
                  Top             =   50
                  Width           =   1815
               End
            End
            Begin VB.TextBox txtCantidadArt 
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
               Left            =   5160
               TabIndex        =   28
               ToolTipText     =   "Cantidad solicitada del artículo"
               Top             =   2320
               Width           =   855
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
               Left            =   2040
               TabIndex        =   27
               ToolTipText     =   "Existencia del artículo en el almacén que surtirá"
               Top             =   2320
               Width           =   855
            End
            Begin VB.CheckBox chkAplicacion 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Aplicación"
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
               Left            =   10680
               TabIndex        =   19
               ToolTipText     =   "Medicamento de aplicación"
               Top             =   760
               Value           =   1  'Checked
               Width           =   1395
            End
            Begin MyCommandButton.MyButton cmdManejos 
               Height          =   375
               Left            =   120
               TabIndex        =   80
               Top             =   6290
               Width           =   1575
               _ExtentX        =   2778
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
               TransparentColor=   16777215
               Caption         =   "Manejos"
               DepthEvent      =   1
               ShowFocus       =   -1  'True
            End
            Begin MyCommandButton.MyButton cmdBorraExistencia 
               Height          =   495
               Left            =   11610
               TabIndex        =   35
               ToolTipText     =   "Eliminar artículo"
               Top             =   6290
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
               Picture         =   "frmRequisicionCargoPac.frx":0E60
               BackColorDown   =   -2147483643
               BackColorOver   =   -2147483633
               BackColorFocus  =   -2147483633
               BackColorDisabled=   -2147483633
               BorderColor     =   -2147483627
               TransparentColor=   16777215
               Caption         =   ""
               DepthEvent      =   1
               PictureDisabled =   "frmRequisicionCargoPac.frx":1574
               PictureAlignment=   4
               PictureDisabledEffect=   0
               ShowFocus       =   -1  'True
            End
            Begin MyCommandButton.MyButton cmdAgregarGrid 
               Height          =   495
               Left            =   10620
               TabIndex        =   33
               ToolTipText     =   "Agregar artículo"
               Top             =   6290
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
               Picture         =   "frmRequisicionCargoPac.frx":1C88
               BackColorDown   =   -2147483643
               BackColorOver   =   -2147483633
               BackColorFocus  =   -2147483633
               BackColorDisabled=   -2147483633
               BorderColor     =   -2147483627
               TransparentColor=   16777215
               Caption         =   ""
               DepthEvent      =   1
               PictureDisabled =   "frmRequisicionCargoPac.frx":239C
               PictureAlignment=   4
               PictureDisabledEffect=   0
               ShowFocus       =   -1  'True
            End
            Begin HSFlatControls.MyCombo cboNombreComercial 
               Height          =   375
               Left            =   2040
               TabIndex        =   20
               ToolTipText     =   "Nombre comercial del artículo"
               Top             =   700
               Width           =   3975
               _ExtentX        =   7011
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
            Begin HSFlatControls.MyCombo cboPaciente 
               Height          =   375
               Left            =   6330
               TabIndex        =   14
               ToolTipText     =   "Paciente relacionado a la requisición"
               Top             =   300
               Width           =   5775
               _ExtentX        =   10186
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
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdHArticulos 
               Height          =   3455
               Left            =   120
               TabIndex        =   32
               ToolTipText     =   "Artículos solicitados"
               Top             =   2730
               Width           =   11990
               _ExtentX        =   21140
               _ExtentY        =   6085
               _Version        =   393216
               ForeColor       =   0
               Rows            =   0
               Cols            =   16
               FixedRows       =   0
               FixedCols       =   0
               ForeColorFixed  =   0
               ForeColorSel    =   0
               BackColorBkg    =   -2147483643
               BackColorUnpopulated=   16777215
               GridColor       =   -2147483638
               GridColorFixed  =   -2147483638
               GridColorUnpopulated=   -2147483638
               AllowBigSelection=   0   'False
               HighLight       =   0
               GridLinesFixed  =   1
               GridLinesUnpopulated=   1
               Appearance      =   0
               FormatString    =   $"frmRequisicionCargoPac.frx":2AB0
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
            Begin VB.Label Label2 
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
               Height          =   250
               Left            =   150
               TabIndex        =   61
               Top             =   1170
               Width           =   1710
            End
            Begin VB.Label Label1 
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
               Left            =   150
               TabIndex        =   62
               Top             =   1980
               Width           =   1005
            End
            Begin VB.Label lblPedidoSug 
               BackColor       =   &H80000005&
               Caption         =   "Se sugiere"
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
               Left            =   6000
               TabIndex        =   63
               Top             =   6780
               Width           =   3030
            End
            Begin VB.Label lblExisTotalAlm 
               BackColor       =   &H80000005&
               Caption         =   "Existencia total en almacenes"
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
               Left            =   6000
               TabIndex        =   64
               Top             =   6280
               Width           =   3030
            End
            Begin VB.Label lblExisTotalDpto 
               BackColor       =   &H80000005&
               Caption         =   "Existencia total en departamentos"
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
               Left            =   6000
               TabIndex        =   65
               Top             =   6540
               Width           =   3030
            End
            Begin VB.Label lblPaciente 
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               Caption         =   "Paciente"
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
               TabIndex        =   66
               Top             =   360
               Width           =   855
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
               Left            =   150
               TabIndex        =   67
               Top             =   1580
               Width           =   690
            End
            Begin VB.Label Label3 
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
               Height          =   250
               Left            =   150
               TabIndex        =   68
               Top             =   760
               Width           =   1830
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
               Height          =   250
               Left            =   6120
               TabIndex        =   69
               Top             =   1170
               Width           =   585
            End
            Begin VB.Label lblCodigoBarras 
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               Caption         =   "Código de barras"
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
               Left            =   6120
               TabIndex        =   70
               Top             =   1580
               Width           =   1725
            End
            Begin VB.Label lblUnidadVta 
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               Caption         =   "Unidad"
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
               Left            =   6120
               TabIndex        =   71
               Top             =   1980
               Width           =   690
            End
            Begin VB.Label lblExistencia 
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               Caption         =   "Existencia"
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
               TabIndex        =   72
               Top             =   2390
               Width           =   930
            End
            Begin VB.Label lblCantidad 
               AutoSize        =   -1  'True
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
               Left            =   3120
               TabIndex        =   73
               Top             =   2390
               Width           =   1950
            End
         End
         Begin HSFlatControls.MyCombo cboManejoMedicamentos 
            Height          =   420
            Left            =   240
            TabIndex        =   79
            Top             =   8880
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            Style           =   1
            Enabled         =   -1  'True
            Text            =   "Combo1"
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
            Left            =   3900
            TabIndex        =   36
            Top             =   8880
            Width           =   4920
            Begin MyCommandButton.MyButton cmdImprimir 
               Height          =   600
               Left            =   4260
               TabIndex        =   44
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
               Picture         =   "frmRequisicionCargoPac.frx":2B5E
               BackColorDown   =   -2147483643
               BackColorOver   =   -2147483633
               BackColorFocus  =   -2147483633
               BackColorDisabled=   -2147483633
               BorderColor     =   -2147483627
               TransparentColor=   16777215
               Caption         =   ""
               DepthEvent      =   1
               PictureDisabled =   "frmRequisicionCargoPac.frx":34E2
               PictureAlignment=   4
               PictureDisabledEffect=   0
               ShowFocus       =   -1  'True
            End
            Begin MyCommandButton.MyButton cmdSuspender 
               Height          =   600
               Left            =   3660
               TabIndex        =   43
               ToolTipText     =   "Suspender requisición surtida parcial"
               Top             =   200
               Width           =   600
               _ExtentX        =   1058
               _ExtentY        =   1058
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
               Picture         =   "frmRequisicionCargoPac.frx":3E64
               BackColorDown   =   -2147483643
               BackColorOver   =   -2147483633
               BackColorFocus  =   -2147483633
               BackColorDisabled=   -2147483633
               BorderColor     =   -2147483627
               TransparentColor=   16777215
               Caption         =   ""
               DepthEvent      =   1
               PictureDisabled =   "frmRequisicionCargoPac.frx":47E8
               PictureAlignment=   4
               PictureDisabledEffect=   0
               ShowFocus       =   -1  'True
            End
            Begin MyCommandButton.MyButton cmdGrabarRegistro 
               Height          =   600
               Left            =   3060
               TabIndex        =   42
               ToolTipText     =   "Guardar el registro"
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
               Picture         =   "frmRequisicionCargoPac.frx":516C
               BackColorDown   =   -2147483643
               BackColorOver   =   -2147483633
               BackColorFocus  =   -2147483633
               BackColorDisabled=   -2147483633
               BorderColor     =   -2147483627
               TransparentColor=   16777215
               Caption         =   ""
               DepthEvent      =   1
               PictureDisabled =   "frmRequisicionCargoPac.frx":5AF0
               PictureAlignment=   4
               PictureDisabledEffect=   0
               ShowFocus       =   -1  'True
            End
            Begin MyCommandButton.MyButton cmdUltimoRegistro 
               Height          =   600
               Left            =   2460
               TabIndex        =   41
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
               Picture         =   "frmRequisicionCargoPac.frx":6474
               BackColorOver   =   -2147483633
               BackColorFocus  =   -2147483633
               BackColorDisabled=   -2147483633
               BorderColor     =   -2147483627
               TransparentColor=   16777215
               Caption         =   ""
               DepthEvent      =   1
               PictureDisabled =   "frmRequisicionCargoPac.frx":6DF6
               PictureAlignment=   4
               PictureDisabledEffect=   0
               ShowFocus       =   -1  'True
            End
            Begin MyCommandButton.MyButton cmdSiguienteRegistro 
               Height          =   600
               Left            =   1860
               TabIndex        =   40
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
               Picture         =   "frmRequisicionCargoPac.frx":7778
               BackColorDown   =   -2147483643
               BackColorOver   =   -2147483633
               BackColorFocus  =   -2147483633
               BackColorDisabled=   -2147483633
               BorderColor     =   -2147483627
               TransparentColor=   16777215
               Caption         =   ""
               CaptionPosition =   4
               DepthEvent      =   1
               ForeColorDisabled=   -2147483629
               ForeColorOver   =   13003064
               ForeColorFocus  =   13003064
               ForeColorDown   =   13003064
               PictureDisabled =   "frmRequisicionCargoPac.frx":80FA
               PictureAlignment=   4
               PictureDisabledEffect=   0
               ShowFocus       =   -1  'True
            End
            Begin MyCommandButton.MyButton cmdBuscar 
               Height          =   600
               Left            =   1260
               TabIndex        =   39
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
               Picture         =   "frmRequisicionCargoPac.frx":8A7C
               BackColorDown   =   -2147483643
               BackColorOver   =   -2147483633
               BackColorFocus  =   -2147483633
               BackColorDisabled=   -2147483633
               BorderColor     =   -2147483627
               TransparentColor=   16777215
               Caption         =   ""
               DepthEvent      =   1
               DropDownPicture =   "frmRequisicionCargoPac.frx":9400
               PictureDisabled =   "frmRequisicionCargoPac.frx":941C
               PictureAlignment=   4
               PictureDisabledEffect=   0
               ShowFocus       =   -1  'True
            End
            Begin MyCommandButton.MyButton cmdAnteriorRegistro 
               Height          =   600
               Left            =   660
               TabIndex        =   38
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
               Picture         =   "frmRequisicionCargoPac.frx":9DA0
               BackColorDown   =   -2147483643
               BackColorOver   =   -2147483633
               BackColorFocus  =   -2147483633
               BackColorDisabled=   -2147483633
               BorderColor     =   -2147483627
               TransparentColor=   16777215
               Caption         =   ""
               DepthEvent      =   1
               PictureDisabled =   "frmRequisicionCargoPac.frx":A722
               PictureAlignment=   4
               PictureDisabledEffect=   0
               ShowFocus       =   -1  'True
            End
            Begin MyCommandButton.MyButton cmdPrimerRegistro 
               Height          =   600
               Left            =   60
               TabIndex        =   37
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
               Picture         =   "frmRequisicionCargoPac.frx":B0A4
               BackColorDown   =   -2147483643
               BackColorOver   =   -2147483633
               BackColorFocus  =   -2147483633
               BackColorDisabled=   -2147483633
               BorderColor     =   -2147483627
               TransparentColor=   16777215
               Caption         =   ""
               DepthEvent      =   1
               PictureDisabled =   "frmRequisicionCargoPac.frx":BA26
               PictureAlignment=   4
               PictureDisabledEffect=   0
               ShowFocus       =   -1  'True
            End
         End
      End
      Begin VB.Frame fraTab1 
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
         Height          =   9750
         Left            =   -74760
         TabIndex        =   74
         Top             =   -160
         Width           =   12975
         Begin HSFlatControls.MyCombo cboEstatus 
            Height          =   375
            Left            =   1080
            TabIndex        =   45
            ToolTipText     =   "Tipos de estatus"
            Top             =   310
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   661
            Style           =   1
            Enabled         =   -1  'True
            Text            =   ""
            Sorted          =   0   'False
            List            =   $"frmRequisicionCargoPac.frx":C3A8
            ItemData        =   $"frmRequisicionCargoPac.frx":C3E8
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
         Begin MSMask.MaskEdBox mskFechaFin 
            Height          =   375
            Left            =   6720
            TabIndex        =   47
            Top             =   310
            Width           =   1575
            _ExtentX        =   2778
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
         Begin MSMask.MaskEdBox mskFechaInicio 
            Height          =   375
            Left            =   4800
            TabIndex        =   46
            Top             =   310
            Width           =   1575
            _ExtentX        =   2778
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
         Begin MSMask.MaskEdBox mskNumCuentaPac 
            Height          =   375
            Left            =   10440
            TabIndex        =   48
            Top             =   310
            Width           =   1920
            _ExtentX        =   3387
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
            PromptInclude   =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   " "
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdhBusqueda 
            Height          =   8920
            Left            =   120
            TabIndex        =   49
            TabStop         =   0   'False
            ToolTipText     =   "Listado de requisiciones, según el tipo de requisición y el estatus"
            Top             =   720
            Width           =   12240
            _ExtentX        =   21590
            _ExtentY        =   15743
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
            Rows            =   0
            Cols            =   16
            FixedRows       =   0
            FixedCols       =   0
            ForeColorFixed  =   0
            ForeColorSel    =   0
            BackColorBkg    =   16777215
            BackColorUnpopulated=   16777215
            GridColor       =   -2147483638
            GridColorFixed  =   -2147483638
            GridColorUnpopulated=   -2147483638
            AllowBigSelection=   0   'False
            FocusRect       =   2
            HighLight       =   0
            GridLinesFixed  =   1
            GridLinesUnpopulated=   1
            Appearance      =   0
            FormatString    =   $"frmRequisicionCargoPac.frx":C3FB
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
         Begin VB.Label Label8 
            BackColor       =   &H80000005&
            Caption         =   "Número de cuenta"
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
            Left            =   8520
            TabIndex        =   78
            Top             =   370
            Width           =   1935
         End
         Begin VB.Label Label7 
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
            Height          =   250
            Left            =   6480
            TabIndex        =   77
            Top             =   370
            Width           =   255
         End
         Begin VB.Label Label6 
            BackColor       =   &H80000005&
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
            Left            =   4440
            TabIndex        =   76
            Top             =   370
            Width           =   375
         End
         Begin VB.Label Label5 
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
            Height          =   250
            Left            =   195
            TabIndex        =   75
            Top             =   370
            Width           =   750
         End
      End
      Begin MSComDlg.CommonDialog cdlTest 
         Left            =   600
         Top             =   8760
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "frmRequisicionCargoPac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Requisiciones de cargos a pacientes para el Módulo de Inventarios'
'-------------------------------------------------------------------------------------
'| Nombre del Proyecto      : Inventario
'| Nombre del Formulario    : frmRequisicionCargoPac
'-------------------------------------------------------------------------------------
'| Objetivo: Realiza la requisición de cargos a pacientes
'-------------------------------------------------------------------------------------
'| Análisis y Diseño        : Eloy González - Inés Saláis
'| Autor                    : Eloy González - Inés Saláis
'| Fecha de Creación        : 14/Agosto/2000
'| Modificó                 : Nombre(s)
'| Fecha última modificación: dd/mes/AAAA
'-------------------------------------------------------------------------------------
Option Explicit 'Verifica que las variables y procedimiento sean creados y dimensionados para su uso

Private vgrptReporte As CRAXDRT.Report
Private ObjRs As New ADODB.Recordset
Private ObjRsAux As New ADODB.Recordset

Dim vllngCantidad As Long
Dim vlblnBuscar As Boolean
Dim vlblnNuevoReg As Boolean    ' Bandera para verificar que se ingresa un nuevo registro
Dim vlblnConsultaReg As Boolean ' Bandera para verificar que se esta consultando un registro
Dim vlstrDptoSolicita As String * 1
Dim vlstrDptoSurte As String * 1
Dim vlstrNumCuenta As String
Dim vlstrValidaUnidad As String
Dim vgStrTipoPaciente As String
Dim vlstrAplicacion As String   ' Variable que guarda si se aplica o no el medicamento de la requisicion
Dim vllngContenidoArtSeleccionado As Long
Dim vlstrx As String
Dim rsIvRequisicionMaestro As New ADODB.Recordset
Dim rsIvRequisicionDetalle As New ADODB.Recordset
Dim vllngLoginInicial As Long
Dim vlblnPrimeraVez As Boolean
Dim vlblnValidarContrasena As Boolean
Dim vlintMedicoTratante As Integer

Public Enum SeleccionArticulo
    saNOSeleccionado = 0    ' No se ha seleccionado el artículo
    saSeleccionado = 1      ' El artículo ya fue seleccionado
    saEnProceso = 2         ' El artículo se encuentra en proceso de selección
End Enum
Public saSeleccion As SeleccionArticulo ' Especifica el estado de selección del artículo en  el combo cboNombreComercial

Dim lblnAutorizarCargos As Boolean  ' Indica si el bit de Autorizar articulos no cubiertos en la requisición de artículos esta activo
Private Type Autorizacion
    blnExcluido As Boolean
    strCodigo As String
    StrCveArticulo As String
End Type
Dim arrautArticulos() As Autorizacion
Dim lintAutorizacion As Integer
Dim lblnPermitirMedicamentos As Boolean
Dim lstrComGen As String

Dim lblnManejaCuadroBasico As Boolean   ' Indica si el hospital maneja un cuadro básico de medicamentos
Dim lblnAutorizacion As Boolean ' Si el hospital maneja cuadro basico de medicamentos, indica si requiere de autorización cuando el medicamento no está en el cuadro básico
Private Type AutorizacionCuadroBasico
    lngPersona As Integer
    strTipoPersona As String
    strFechaAutorizacion As String
    strCveMedicamento As String
End Type
Dim arrAutorizadosCB() As AutorizacionCuadroBasico
Dim lintMedicamentosCB As Integer
Public llngCvePersonaAutoriza As Long
Public lstrTipoPersonaAutoriza As String
Public lstrFechaAutorizacion As String

'Variables que guardan el valor de las columnas del grid
Const lintColFixed = 0
Dim lintColsTotales As Integer    ' Columnas totales del grid de artículos
Dim lintColClave As Integer       ' Clave del producto
Dim lintColNombreArt As Integer   ' Nombre comercial
Dim lintColCantidad As Integer    ' Cantidad de requisicion
Dim lintColUnidad As Integer      ' Unidad en la que se pide el producto
Dim lintColEstatus As Integer     ' Estatus del producto
Dim lintColCveUnidad As Integer   ' Clave de la unidad en la que se pide el producto
Dim lintColTipoUnidad As Integer  ' Tipo de unidad A)lterna, M)inima
Dim lintColCantSurtida As Integer ' Cantidad surtida
Dim lintColEstatusOr As Integer   ' Estatus original
Dim lstrTitulos As String
Dim lintTotalManejos As Integer  'Variable que indica el total de manejos de medicamentos actualmente activos en el sistema
Dim position As Integer
Dim intDeptoIngreso As Integer
  
'Estructura para el manejo de medicamentos'
Private Type ManejoMedicamentos   'Para los colores del manejo de medicamentos
    intCveManejo As Integer
    strColor As String
    strSimbolo As String
End Type
Dim aManejoMedicamentos() As ManejoMedicamentos

Private Sub cboAlmacenSurte_Click()
On Error GoTo NotificaError
    Set ObjRsAux = frsRegresaRs("Select * from NoDepartamento where smiCveDepartamento = " & CStr(cboDepartamento.ItemData(cboDepartamento.ListIndex)))
    If ObjRsAux.RecordCount > 0 Then
        ObjRsAux.MoveFirst
        vlstrDptoSolicita = ObjRsAux!chrClasificacion
    End If
    ObjRsAux.Close
    Set ObjRsAux = frsRegresaRs("Select * from NoDepartamento where smiCveDepartamento = " & CStr(cboAlmacenSurte.ItemData(cboAlmacenSurte.ListIndex)))
    If ObjRsAux.RecordCount > 0 Then
        ObjRsAux.MoveFirst
        vlstrDptoSurte = ObjRsAux!chrClasificacion
    End If
    ObjRsAux.Close
    Exit Sub
NotificaError:
    Call pError
End Sub

Private Sub cboAlmacenSurte_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    Select Case KeyCode
        Case vbKeyReturn
            optInterno.Value = True
            If vlstrDptoSurte = "G" Then
                Call MsgBox(SIHOMsg("66"), vbExclamation, "Mensaje") 'No existe un almacen
            Else
                fraIntExt.Enabled = True
                If optInterno.Value Then
                    Optinterno_Click
                Else
                    Optexterno_Click
                End If
                cboPaciente_Click
                Optinterno_Click
                cboPaciente.Enabled = True
            End If
    End Select
Exit Sub
NotificaError:
    Call pError
End Sub

Private Sub pCargaPaciente(vlstrAreas As String, ObjCbo As MyCombo)
'Procedimiento para cargar los pacientes de una(s) areas determinadas
    Dim vlintseq, vlintSeq1, vlintLargo As Integer
    Dim vlstrSentencia As String
    Dim rsPacienteArea As New ADODB.Recordset
    
    ObjCbo.Clear
    vlstrSentencia = "I|" & Trim(vglngNumeroLogin) & "|0"
    Set rsPacienteArea = frsEjecuta_SP(vlstrSentencia, "sp_IVSelPacientesArea")
    
    If rsPacienteArea.RecordCount > 0 Then
        intDeptoIngreso = 0
        fraBuscaArticulo.Enabled = True
        cboPaciente.Enabled = True
        intDeptoIngreso = IIf(IsNull(rsPacienteArea!INTCVEDEPTOINGRESO), 0, rsPacienteArea!INTCVEDEPTOINGRESO)
        Do While Not rsPacienteArea.EOF
            ObjCbo.AddItem Trim(IIf(IsNull(rsPacienteArea!Cuarto), "", rsPacienteArea!Cuarto) & " " & rsPacienteArea!Nombre & " (" & IIf(IsNull(rsPacienteArea!PROCEDENCIA), "", rsPacienteArea!PROCEDENCIA)) & ")"
            ObjCbo.ItemData(ObjCbo.NewIndex) = rsPacienteArea!cuenta
            rsPacienteArea.MoveNext
        Loop
    End If
    rsPacienteArea.Close
End Sub

Private Sub pCargaPacienteAmbulatorio(ObjCbo As MyCombo)
'Procedimiento para cargar los pacientes de una(s) areas determinadas
    Dim vlintseq, vlintSeq1, vlintLargo As Integer
    Dim vlstrSentencia As String
    Dim rsPacienteArea As New ADODB.Recordset
    
    ObjCbo.Clear
    vlstrSentencia = "A|0|0"
    Set rsPacienteArea = frsEjecuta_SP(vlstrSentencia, "sp_IVSelPacientesArea")
    If rsPacienteArea.RecordCount > 0 Then
         intDeptoIngreso = 0
        fraBuscaArticulo.Enabled = True
        cboPaciente.Enabled = True
        intDeptoIngreso = IIf(IsNull(rsPacienteArea!INTCVEDEPTOINGRESO), 0, rsPacienteArea!INTCVEDEPTOINGRESO)
        Do While Not rsPacienteArea.EOF
            ObjCbo.AddItem Trim(IIf(IsNull(rsPacienteArea!Cuarto), "", rsPacienteArea!Cuarto) & " " & rsPacienteArea!Nombre & " (" & IIf(IsNull(rsPacienteArea!PROCEDENCIA), "", rsPacienteArea!PROCEDENCIA)) & ")"
            ObjCbo.ItemData(ObjCbo.NewIndex) = rsPacienteArea!cuenta
            rsPacienteArea.MoveNext
        Loop
    End If
    rsPacienteArea.Close
End Sub

Private Sub cboEstatus_Click()
    pLlenarBusqueda
End Sub

Private Sub pLlenarBusqueda()
    Dim vlstrSentencia As String
    Dim vlstrEstatusRequi As String
    
    If vlblnBuscar Then
        vlstrEstatusRequi = IIf(cboEstatus.ListIndex <> -1, Trim(cboEstatus.List(cboEstatus.ListIndex)), "")
        vlstrSentencia = cboDepartamento.ItemData(cboDepartamento.ListIndex) & "|" & vlstrEstatusRequi & "|" & Format(mskFechaInicio, "yyyy-MM-dd") & "|" & Format(mskFechaFin, "yyyy-MM-dd") & "|" & IIf(mskNumCuentaPac.ClipText = "", "-1", mskNumCuentaPac.ClipText) & "|A"
        Set ObjRs = frsEjecuta_SP(vlstrSentencia, "sp_IVSelRequisicionCargoPacie2")
        If ObjRs.RecordCount > 0 Then
            Call pIniciaMshFGrid(grdHBusqueda)
            Call pLlenarMshFGrdRs(grdHBusqueda, ObjRs)
            Call pConfBusq
        Else
            Call pIniciaMshFGrid(grdHBusqueda)
            Call pLimpiaMshFGrid(grdHBusqueda)
        End If
        ObjRs.Close
    End If
End Sub

Private Sub pLimpiaCaptura()
    cboNombreComercial.Text = ""
    txtFamilia.Text = ""
    txtSubfamilia.Text = ""
    txtNombreGenerico.Text = ""
    txtClaveArticulo.Text = ""
    txtCodigoBarras.Text = ""
    txtUnidad.Text = ""
    txtExistencia.Text = ""
    txtCantidadArt.Text = ""
    optAlterna.Value = False
    optMinima.Value = False
    optMinima.Enabled = True
    cmdAgregarGrid.Enabled = False
    Call Unidades
End Sub

Private Sub cboEstatus_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub cboNombreComercial_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
         If cboPaciente.Text <> "" Then
            If cboNombreComercial.ListIndex < 0 Then
                If Len(cboNombreComercial.Text) > 0 Then
                    saSeleccion = saEnProceso ' Se está en proceso de seleccionar un artículo
                    vgstrVarIntercam = UCase(cboNombreComercial.Text) ' Variable global de entrada al frmlista para la busqueda
                    vgintCveDeptoCargo = CStr(cboAlmacenSurte.ItemData(cboAlmacenSurte.ListIndex))
                    frmListaPaciente.llngCuentaPaciente = cboPaciente.ItemData(cboPaciente.ListIndex)
                    frmListaPaciente.lstrTipoPaciente = vgStrTipoPaciente
                    frmListaPaciente.vgblnNombreGenerico = 0
                    frmListaPaciente.vgstrCriterioParam = IIf(Not lblnPermitirMedicamentos, "0", IIf(optTodos.Value, "-1", IIf(optMedicamento.Value, "1", "0")))
                    frmListaPaciente.lblnCuadroBasico = lblnManejaCuadroBasico
                    frmListaPaciente.lblnAutorizacionMedicamento = lblnAutorizacion
                    llngCvePersonaAutoriza = 0
                    lstrTipoPersonaAutoriza = ""
                    lstrFechaAutorizacion = ""
                    frmListaPaciente.Show vbModal, Me
                    If Len(vgstrVarIntercam) > 0 Then
                        pCargaDatos Trim(vgstrVarIntercam), vgintCveDeptoCargo
                        If lblnAutorizacion And llngCvePersonaAutoriza <> 0 Then
                            pDatosAutorizacion llngCvePersonaAutoriza, lstrTipoPersonaAutoriza, Trim(vgstrVarIntercam), lstrFechaAutorizacion
                        End If
                    Else
                        If cboNombreComercial.ListCount > 0 Then
                            cboNombreComercial.ListIndex = 0
                        End If
                    End If
                Else
                    If cboNombreComercial.ListCount > 0 Then
                        cboNombreComercial.ListIndex = 0
                    End If
                    txtNombreGenerico.SetFocus
                End If
            Else
                cboNombreComercial.SetFocus
            End If
        Else
            cboPaciente.SetFocus
        End If
    End If
End Sub

Private Sub pDatosAutorizacion(lngAutoriza As Long, strTipo As String, strCveMedicamento As String, strFecha As String)
    Dim lintAux As Integer
    Dim lblnExiste As Boolean
    
    lblnExiste = False
    For lintAux = 0 To UBound(arrAutorizadosCB)
        If arrAutorizadosCB(lintAux).strCveMedicamento = strCveMedicamento Then
            lblnExiste = True
            Exit For
        End If
    Next lintAux
    If lblnExiste Then
        arrAutorizadosCB(lintAux).lngPersona = lngAutoriza
        arrAutorizadosCB(lintAux).strTipoPersona = strTipo
        arrAutorizadosCB(lintAux).strFechaAutorizacion = strFecha
    Else
        ReDim Preserve arrAutorizadosCB(lintMedicamentosCB)
        arrAutorizadosCB(lintMedicamentosCB).lngPersona = lngAutoriza
        arrAutorizadosCB(lintMedicamentosCB).strTipoPersona = lstrTipoPersonaAutoriza
        arrAutorizadosCB(lintMedicamentosCB).strFechaAutorizacion = strFecha
        arrAutorizadosCB(lintMedicamentosCB).strCveMedicamento = strCveMedicamento
        lintMedicamentosCB = lintMedicamentosCB + 1
    End If
End Sub

Private Sub pCargaDatos(vlstrCveArticulo As String, vlintCveDeptoSurte As Integer)
    Dim rsDatos As New ADODB.Recordset
    Dim vlstrx As String
    Dim vlrsSubrogado As ADODB.Recordset
    
    vlstrx = "SELECT " & _
                "IvArticulo.vchNombreComercial NombreComercial " & _
                ",IvFamilia.vchDescripcion Familia" & _
                ",IvSubFamilia.vchDescripcion Subfamilia" & _
                ",isnull(IvGenerico.vchDescripcion,' ') NombreGenerico" & _
                ",IvArticulo.chrCveArticulo ClaveArticulo " & _
                ",IvArticulo.intContenido Contenido " & _
                IIf(frmListaPaciente.vgblnNombreGenerico <> 2, ",isnull((SELECT MAX(vchCodigoBarras) FROM IvCodigoBarrasArticulo WHERE IvCodigoBarrasArticulo.chrCveArticulo=IvArticulo.chrCveArticulo),'') AS CodigoBarras ", ",ivcodigobarrasarticulo.vchcodigobarras AS CodigoBarras ") & _
                ",IvUnidadVenta.vchDescripcion UnidadVenta " & _
                ",IvArticulo.chrCveArtMedicamen " & _
                ", ivarticulo.INTIDARTICULO " & _
             "FROM IvArticulo " & _
                "INNER JOIN IvFamilia ON IvArticulo.chrCveArtMedicamen = IvFamilia.chrCveArtMedicamen AND IvArticulo.chrCveFamilia = IvFamilia.chrCveFamilia " & _
                "INNER JOIN IvSubFamilia ON IvArticulo.chrCveArtMedicamen = IvSubFamilia.chrCveArtMedicamen AND IvArticulo.chrCveFamilia = IvSubFamilia.chrCveFamilia AND IvArticulo.chrCveSubFamilia = IvSubFamilia.chrCveSubFamilia " & _
                "INNER JOIN IvUnidadVenta ON IvArticulo.intCveUniMinimaVta = IvUnidadVenta.intCveUnidadVenta " & _
                "LEFT OUTER JOIN IvArticuloGenerico ON IvArticulo.intIdArticulo = IvArticuloGenerico.intIdArticulo " & _
                "LEFT OUTER JOIN IvGenerico ON IvArticuloGenerico.intIdGenerico = IvGenerico.intIdGenerico " & _
                IIf(frmListaPaciente.vgblnNombreGenerico = 2, "LEFT OUTER JOIN ivcodigobarrasarticulo ON ivcodigobarrasarticulo.chrcvearticulo=ivarticulo.chrcvearticulo ", "") & _
             "WHERE IvArticulo.chrCveArticulo='" & Trim(vlstrCveArticulo) & "'"
        
    If frmListaPaciente.vgblnNombreGenerico = 2 And vgstrvarcodigodebarras <> "" Then
        vlstrx = vlstrx & " AND ivcodigobarrasarticulo.vchcodigobarras=" & vgstrvarcodigodebarras
    End If
    Set rsDatos = frsRegresaRs(vlstrx)
    
    If rsDatos.RecordCount <> 0 Then
        cboNombreComercial.Text = rsDatos!NombreComercial
        txtFamilia.Text = rsDatos!Familia
        txtSubfamilia.Text = rsDatos!Subfamilia
        txtNombreGenerico.Text = rsDatos!NombreGenerico
        txtClaveArticulo.Text = rsDatos!claveArticulo
        txtCodigoBarras.Text = IIf(IsNull(rsDatos!CodigoBarras), "", rsDatos!CodigoBarras)
        txtUnidad.Text = fstrUnidad(txtClaveArticulo.Text, IIf(optAlterna.Value, 1, 9))
        txtIdArticulo.Text = rsDatos!intIdArticulo
        vllngContenidoArtSeleccionado = rsDatos!Contenido
            
        Set vlrsSubrogado = frsEjecuta_SP("|" & rsDatos!claveArticulo & "|1", "sp_IVArticuloSub")
        If Not vlrsSubrogado.EOF Then
            optMinima.Enabled = False
            optAlterna.Enabled = True
            optAlterna.Value = True
        Else
            Call Unidades
        End If
        vlrsSubrogado.Close
            
        'Existencia
        vlstrx = "SELECT " & _
                    "IvUbicacion.intExistenciaDeptoUM," & _
                    "IvUbicacion.intExistenciaDeptoUV," & _
                    "IvArticulo.intContenido " & _
                 "FROM IvUbicacion " & _
                    "INNER JOIN IvArticulo ON " & _
                    "IvUbicacion.chrCveArticulo = IvArticulo.chrCveArticulo " & _
                 "WHERE IvUbicacion.chrCveArticulo='" & Trim(vlstrCveArticulo) & "' AND IvUbicacion.smiCveDepartamento=" & Str(cboAlmacenSurte.ItemData(cboAlmacenSurte.ListIndex))
        Set rsDatos = frsRegresaRs(vlstrx)
        txtExistencia.Text = ""
        If rsDatos.RecordCount <> 0 Then
            txtExistencia.Text = (rsDatos!intExistenciaDeptouv * rsDatos!intContenido) + rsDatos!intexistenciadeptoum
        End If
        txtCantidadArt.SetFocus
    End If
End Sub

Private Sub cboNombreComercial_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboNombreComercial_Validate(Cancel As Boolean)
    If saSeleccion <> saSeleccionado And cboNombreComercial.Text <> "" Then Cancel = True
    lstrComGen = "cboNombreComercial"
End Sub

Private Sub cboPaciente_Click()
    vlstrNumCuenta = cboDepartamento.ItemData(1)
End Sub

Private Sub cboPaciente_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    Select Case KeyCode
        Case vbKeyReturn
            If cboPaciente.Text <> "" Then
                If pRevisaCuentaCerrada Then
                    cboPaciente.Clear
                    Exit Sub ' Revisa si la cuenta esta cerrada
                End If
                If optInterno.Value Then pVerificarInformacionFaltantePaciente cboPaciente.ItemData(cboPaciente.ListIndex)
                If optTodos.Enabled Then
                    optTodos.Value = True
                ElseIf optArticulo.Enabled Then
                    optArticulo.Value = True
                ElseIf optMedicamento.Enabled Then
                    optMedicamento.Value = True
                End If
                cboNombreComercial.SetFocus
            Else
                cboPaciente.SetFocus
            End If
            If cboPaciente.ListIndex > -1 Then
                cboPaciente.Enabled = False
                saSeleccion = saNOSeleccionado
            End If
    End Select
    
Exit Sub
NotificaError:
    Call pError
End Sub

Private Sub cboPaciente_LostFocus()
    If cboPaciente.Text <> "" Then
        If pRevisaCuentaCerrada Then
            cboPaciente.Clear
            Exit Sub ' Revisa si la cuenta esta cerrada
        End If
    End If
End Sub

Private Sub cboTipoReq_Click()
    vlblnBuscar = False
    Call pEscogerCboTipoReq
    vlblnBuscar = True
End Sub

Private Sub cboTipoReq_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If cboAlmacenSurte.ListCount > 0 And cboAlmacenSurte.Enabled Then cboAlmacenSurte.SetFocus
    End If
End Sub

Private Sub chkUrgente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If cboDepartamento.ListCount > 0 And cboDepartamento.Enabled Then
            cboDepartamento.SetFocus
        Else
            If cboTipoReq.ListCount > 0 And cboTipoReq.Enabled Then cboTipoReq.SetFocus
        End If
    End If
    Exit Sub
NotificaError:
    Call pError
End Sub

Private Sub cmdAgregarGrid_Click()
    Dim vllngPosicion As Long
    Dim vlblnAgregar As Boolean
    Dim vlstrUnidad As String
    Dim vlstrUnidadControl As String
    Dim vllngCnt As Long
    
    If fblnCapturaValida() Then
        'Se revisa si es excluido antes de cargarse:
        vlblnAgregar = True
        
        If grdHArticulos.Rows <> 0 Then
            vllngPosicion = fintLocRegMshFGrd(grdHArticulos, txtClaveArticulo.Text, lintColClave)
            If vllngPosicion <> 0 Then
                'Desea actualizar datos
                If MsgBox(SIHOMsg("7"), vbExclamation + vbYesNo, "Mensaje") = vbNo Then
                    vlblnAgregar = False
                Else
                    If Not fblnContinuar Then vlblnAgregar = False
                End If
            Else
                If fblnContinuar Then
                    vllngPosicion = grdHArticulos.Rows
                    grdHArticulos.Rows = grdHArticulos.Rows + 1
                Else
                    vlblnAgregar = False
                End If
            End If
        Else
            If fblnContinuar Then
                vllngPosicion = 1
                grdHArticulos.Rows = 2
                pConfiguraGridArticulos
            Else
                vlblnAgregar = False
            End If
        End If
            
        If vlblnAgregar Then
            If optMinima.Value Then
                vlstrUnidad = IIf(vllngContenidoArtSeleccionado = 1, fstrUnidad(txtClaveArticulo.Text, 1), fstrUnidad(txtClaveArticulo.Text, 0))
                vlstrUnidadControl = IIf(vllngContenidoArtSeleccionado = 1, "A", "M")
            Else
                vlstrUnidad = fstrUnidad(txtClaveArticulo.Text, 1)
                vlstrUnidadControl = "A"
            End If
        
            With grdHArticulos
                .TextMatrix(vllngPosicion, lintColClave) = txtClaveArticulo.Text
                .TextMatrix(vllngPosicion, lintColNombreArt) = cboNombreComercial.Text
                .TextMatrix(vllngPosicion, lintColCantidad) = txtCantidadArt.Text
                .TextMatrix(vllngPosicion, lintColUnidad) = vlstrUnidad
                .TextMatrix(vllngPosicion, lintColEstatus) = "PENDIENTE"
                .TextMatrix(vllngPosicion, lintColCveUnidad) = vlstrUnidadControl
            End With
            
            'Se colorean todos los manejos en caso de que se llamara la función fintLocRegMshFGrd
            For vllngCnt = 1 To vllngPosicion
                pColorearManejo grdHArticulos, cboManejoMedicamentos, lintColFixed + 1, grdHArticulos.TextMatrix(vllngCnt, lintColClave), vllngCnt
            Next vllngCnt
            
            cmdBorraExistencia.Enabled = True
            pHabilita 0, 0, 0, 0, 0, 1, 0
            saSeleccion = saNOSeleccionado
        End If
    End If
    
    pLimpiaCaptura
    
    cboNombreComercial.SetFocus
End Sub

Private Function fblnContinuar() As Boolean
    ' Cargo excluido
    Dim blnAutoriza As Boolean
    Dim blnExcluido As Boolean
    Dim strCodigo As String
    Dim lintCont As Integer
    Dim lblnActualizar As Boolean
    
    fblnContinuar = True
    If Not fblnCargoExcluidoContinuar(cboPaciente.ItemData(cboPaciente.ListIndex), vgStrTipoPaciente, txtIdArticulo.Text, "AR", blnAutoriza) Then
        fblnContinuar = False
    Else
        If blnAutoriza Then
            If lblnAutorizarCargos Then ' Si parametro de autorizar articulos no cubiertos en la requisicion de articulos activo
                If fblnAceptarCargoExcluido(blnExcluido, strCodigo) Then
                    lblnActualizar = False
                    For lintCont = 0 To UBound(arrautArticulos) ' Si el articulo ya esta en el arreglo
                        If arrautArticulos(lintCont).StrCveArticulo = Trim(txtClaveArticulo.Text) Then
                            lblnActualizar = True
                            Exit For
                        End If
                    Next lintCont
                    If lblnActualizar Then
                        arrautArticulos(lintCont).blnExcluido = blnExcluido
                        arrautArticulos(lintCont).strCodigo = strCodigo
                    Else
                        ReDim Preserve arrautArticulos(lintAutorizacion)
                        arrautArticulos(lintAutorizacion).blnExcluido = blnExcluido
                        arrautArticulos(lintAutorizacion).strCodigo = strCodigo
                        arrautArticulos(lintAutorizacion).StrCveArticulo = Trim(txtClaveArticulo.Text)
                        lintAutorizacion = lintAutorizacion + 1
                    End If
                Else
                    fblnContinuar = False
                End If
            End If
        End If
    End If
    
End Function

Private Sub pHabilita(vlint1 As Integer, vlint2 As Integer, vlint3 As Integer, vlint4 As Integer, vlint5 As Integer, vlint6 As Integer, vlint7 As Integer)

    cmdPrimerRegistro.Enabled = False
    cmdAnteriorRegistro.Enabled = False
    cmdBuscar.Enabled = False
    cmdSiguienteRegistro.Enabled = False
    cmdUltimoRegistro.Enabled = False
    
    cmdGrabarRegistro.Enabled = False
    cmdImprimir.Enabled = False
    cmdSuspender.Enabled = False
    
    If vlint1 = 1 Then cmdPrimerRegistro.Enabled = True
    If vlint2 = 1 Then cmdAnteriorRegistro.Enabled = True
    If vlint3 = 1 Then cmdBuscar.Enabled = True
    If vlint4 = 1 Then cmdSiguienteRegistro.Enabled = True
    If vlint5 = 1 Then cmdUltimoRegistro.Enabled = True
    If vlint6 = 1 Then cmdGrabarRegistro.Enabled = True
    If vlint7 = 1 Then cmdImprimir.Enabled = True

End Sub

Private Function fstrUnidad(vlstrCveArticulo As String, vlintCveUnidad As Integer) As String
    Dim rsUnidad As New ADODB.Recordset

    fstrUnidad = ""
    
    vlstrx = "" & _
    "select " & _
        "IvUnidadVenta.vchDescripcion " & _
    "From " & _
        "IvArticulo " & _
        "inner join IvUnidadVenta on " & _
        "case when " & Str(vlintCveUnidad) & "=1 then " & _
            "IvArticulo.intCveUniAlternaVta " & _
        "Else " & _
            "IvArticulo.intCveUniMinimaVta " & _
        "End " & _
        "=IvUnidadVenta.intCveUnidadVenta " & _
    "Where " & _
        "IvArticulo.chrCveArticulo='" & Trim(vlstrCveArticulo) & "'"
        
    Set rsUnidad = frsRegresaRs(vlstrx)
    If rsUnidad.RecordCount <> 0 Then
        fstrUnidad = rsUnidad.Fields(0)
    End If
End Function

Private Sub pConfiguraGridArticulos()
    Dim vlintCnt As Integer
    
    With grdHArticulos
        .Redraw = False
        .Cols = lintColsTotales
        .FixedCols = 1
        .FixedRows = 1
        .FormatString = lstrTitulos & "|Clave|Nombre comercial|Cantidad|Unidad|Estado"
        
        For vlintCnt = lintColFixed + 1 To lintColClave - 1
            .ColWidth(vlintCnt) = 0
        Next vlintCnt
        
        .ColWidth(lintColFixed) = 100
        .ColWidth(lintColClave) = 1300
        .ColWidth(lintColNombreArt) = 5000
        .ColWidth(lintColCantidad) = 1200
        .ColWidth(lintColUnidad) = 1500
        .ColWidth(lintColEstatus) = 2000
        .ColWidth(lintColCveUnidad) = 0
        .ColAlignment(lintColClave) = flexAlignLeftCenter
        .ColAlignment(lintColNombreArt) = flexAlignLeftCenter
        .ColAlignment(lintColCantidad) = flexAlignRightCenter
        .ColAlignment(lintColUnidad) = flexAlignLeftCenter
        .ColAlignment(lintColEstatus) = flexAlignLeftCenter
        .ColAlignmentFixed(lintColClave) = flexAlignCenterCenter
        .ColAlignmentFixed(lintColNombreArt) = flexAlignCenterCenter
        .ColAlignmentFixed(lintColCantidad) = flexAlignCenterCenter
        .ColAlignmentFixed(lintColUnidad) = flexAlignCenterCenter
        .ColAlignmentFixed(lintColEstatus) = flexAlignCenterCenter
        
        .Redraw = True
    End With
End Sub

Private Function fblnCapturaValida() As Boolean
    Dim vlrsSubrogado As ADODB.Recordset
    fblnCapturaValida = True
    
    If Trim(cboNombreComercial.Text) = "" Then
        fblnCapturaValida = False
        cboNombreComercial.SetFocus
    End If
    If fblnCapturaValida And Val(txtExistencia.Text) = 0 Then
        Set vlrsSubrogado = frsEjecuta_SP("|" & txtClaveArticulo.Text & "|1", "sp_IVArticuloSub")
        If vlrsSubrogado.EOF Then
            If MsgBox(SIHOMsg(712), vbYesNo + vbQuestion, "Mensaje") = vbNo Then
                fblnCapturaValida = False
                txtCantidadArt.SetFocus
            End If
        End If
        vlrsSubrogado.Close
    End If
    If fblnCapturaValida And Val(txtCantidadArt.Text) = 0 Then
        fblnCapturaValida = False
        txtCantidadArt.SetFocus
    End If
End Function

Private Sub cmdAnteriorRegistro_Click()
    If grdHBusqueda.Rows - 1 > 0 Then
        If (grdHBusqueda.Row - 1) >= 1 Then
            grdHBusqueda.Row = grdHBusqueda.Row - 1
            Call pConsultaReq(grdHBusqueda.TextMatrix(grdHBusqueda.Row, 1))
        End If
    Else
        cmdPrimerRegistro_Click
    End If
End Sub

Private Sub cmdBorraExistencia_Click()
    If grdHArticulos.Rows - 1 > 0 Then
        If grdHArticulos.Row > 0 Then
            If MsgBox(SIHOMsg("6"), vbExclamation + vbYesNo, "Mensaje") = vbYes Then 'Desea eliminar los datos
                Call pBorrarRegMshFGrd(grdHArticulos, grdHArticulos.Row)
            End If
        End If

        cmdGrabarRegistro.Enabled = IIf(grdHArticulos.Rows - 1 > 0 And vlblnNuevoReg, True, False)
        cmdBorraExistencia.Enabled = IIf(grdHArticulos.Rows - 1 > 0 And vlblnNuevoReg, True, False)
        cboNombreComercial.SetFocus
    End If
End Sub

Private Sub cmdBuscar_Click()
    Dim vlintPasa As Integer
    vlintPasa = 0
    vlblnBuscar = False
    cboEstatus.ListIndex = 0
    vlblnBuscar = True
    cboEstatus.ListIndex = 0
    vlintPasa = 1
    sstObj.Tab = 1
    cboEstatus_Click
End Sub

Private Sub cmdBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            optMinima.Enabled = True
            fraBuscaArticulo.Enabled = True
            cboAlmacenSurte.Enabled = False
            cboTipoReq.Enabled = False
            txtNumReq.Enabled = False
    
            Select Case Mid(cboTipoReq.List(cboTipoReq.ListIndex), 1, 1)
                Case "P"    'Pedido
                    optMinima.Enabled = False
                Case Else
            End Select
    End Select
End Sub

Private Sub cmdCancelar_Click()
    If grdHArticulos.Rows - 1 > 0 Then
        If grdHArticulos.Row > 0 Then
            If grdHArticulos.TextMatrix(grdHArticulos.Row, lintColEstatusOr) = "PENDIENTE" Then
                If MsgBox("¿Desea cancelar?", vbExclamation + vbYesNo, "Mensaje") = vbYes Then 'Desea cancelar el registro
                    grdHArticulos.TextMatrix(grdHArticulos.Row, lintColEstatus) = "CANCELADA"
                    Call pActEstatusCab
                End If
            Else
                Call MsgBox("¡La requisición no se puede cancelar!", vbExclamation, "Mensaje")
            End If
        End If
    End If
End Sub

Private Sub cmdGrabarRegistro_Click()
    Dim rsMedicoTratante As New ADODB.Recordset
    Dim vlstrSentencia As String
    
    vgstrGrabaMedicoEmpleado = vgstrMedicoEnfermera
    vglngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento, vgstrGrabaMedicoEmpleado)
        
    If vglngPersonaGraba = 0 Then Exit Sub
        
    If vglngPersonaGraba <> vllngLoginInicial And vlblnValidarContrasena Then
        ' La persona que inició el proceso no corresponde a la persona que quiere grabar.
        MsgBox SIHOMsg(593), vbOKOnly + vbInformation, "Mensaje"
        Exit Sub
    End If
    
    If grdHArticulos.Rows - 1 > 0 Then
            If cboPaciente.ListIndex >= 0 Then
                vgintCveDepartamento = cboDepartamento.ItemData(cboDepartamento.ListIndex)
                vgblnContrasenaCorr = False
'               ******************** Obtener médico tratante
                vlstrSentencia = "SELECT exmedicoacargo.INTCVEMEDICO FROM exmedicoacargo WHERE INTCONSECUTIVO = (" & _
                                         "SELECT MAX(exmedicoacargo.INTCONSECUTIVO)  FROM exmedicoacargo " & _
                                         "Where NUMNUMCUENTA = " & cboPaciente.ItemData(cboPaciente.ListIndex) & _
                                         " AND exmedicoacargo.DTMFECHAHORATERMINO IS NULL " & _
                                         "AND exmedicoacargo.CHRESTATUSMEDICO = 'A' " & _
                                         "AND exmedicoacargo.BITMEDICORESP = 1)"
                Set rsMedicoTratante = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
                If rsMedicoTratante.RecordCount > 0 Then
                    If IsNull(rsMedicoTratante!INTCVEMEDICO) Or rsMedicoTratante!INTCVEMEDICO = 0 Then
                        rsMedicoTratante.Close
                       vlstrSentencia = "SELECT INTCVEMEDICOTRATANTE FROM EXPACIENTEINGRESO " & _
                                                 "WHERE INTNUMCUENTA = " & cboPaciente.ItemData(cboPaciente.ListIndex)
                        Set rsMedicoTratante = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
                        If rsMedicoTratante.RecordCount > 0 Then
                            If IsNull(rsMedicoTratante!INTCVEMEDICOTRATANTE) Or rsMedicoTratante!INTCVEMEDICOTRATANTE = 0 Then
                                Call MsgBox("¡No se le ha asignado un médico tratante al paciente!", vbExclamation, "Mensaje")
                                rsMedicoTratante.Close
                                Exit Sub
                            Else
                                vlintMedicoTratante = rsMedicoTratante!INTCVEMEDICO
                            End If
                        End If
                    Else
                        vlintMedicoTratante = rsMedicoTratante!INTCVEMEDICO
                    End If
                End If
                rsMedicoTratante.Close
                Set rsMedicoTratante = Nothing
'               *************************************************************************
                
                If vlblnNuevoReg Then
                    pGrabaRegistros
                End If
                If vlblnConsultaReg Then
                    pGrabarModRegistros
                End If
            Else
                Call MsgBox("¡No ha escogido un paciente!", vbExclamation + vbYesNo, "Mensaje")
                If cboPaciente.Enabled Then
                    cboPaciente.SetFocus
                End If
            End If
        
        'Set rsMedicoTratante = frsRegresaRs("SELECT vchApellidoPaterno ||' '|| vchApellidoMaterno ||' '|| vchNombre AS Nombre FROM noempleado WHERE intCveEmpleado = " & vglngPersonaGraba)
        'txtEmpleado.Text = IIf(rsEmpleado.RecordCount > 0, rsEmpleado!Nombre, "")
        'rsEmpleado.Close
    End If
End Sub

Private Sub cmdImprimir_Click()
    If vlblnConsultaReg Then Call pimprimir(txtNumReq.Text, "P")
End Sub

Private Sub cmdManejos_Click()
On Error GoTo NotificaError
    frmManejoMedicamentos.blnSoloBusqueda = True
    frmManejoMedicamentos.Show vbModal, Me
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdManejos_Click"))
End Sub

Private Sub cmdPrimerRegistro_Click()
    If grdHBusqueda.Rows - 1 > 0 Then
        grdHBusqueda.Row = 1
        Call pConsultaReq(grdHBusqueda.TextMatrix(grdHBusqueda.Row, 1))
    Else
        cboEstatus_Click
        If grdHBusqueda.Row > 0 Then
            grdHBusqueda.Row = 1
            Call pConsultaReq(grdHBusqueda.TextMatrix(grdHBusqueda.Row, 1))
        End If
    End If
End Sub

Private Sub cmdSiguienteRegistro_Click()
    If grdHBusqueda.Rows - 1 > 0 Then
        If (grdHBusqueda.Row + 1) <= (grdHBusqueda.Rows - 1) Then
            grdHBusqueda.Row = grdHBusqueda.Row + 1
            Call pConsultaReq(grdHBusqueda.TextMatrix(grdHBusqueda.Row, 1))
        End If
    Else
        cmdPrimerRegistro_Click
    End If
End Sub

Private Sub cmdSuspender_Click()
  Dim rsReqMaestro As New ADODB.Recordset
  Dim vlstrx As String
  
  vlstrx = ""
  vlstrx = vlstrx & "Select vchEstatusRequis,BITDETIENESALIDA from ivRequisicionMaestro where numNumRequisicion = " & txtNumReq.Text
  Set rsReqMaestro = frsRegresaRs(vlstrx)
  
  If rsReqMaestro.RecordCount > 0 Then
    If rsReqMaestro!vchEstatusRequis = "SURTIDA PARCIAL" And rsReqMaestro!BITDETIENESALIDA <> 1 Then
      If MsgBox("¿Desea suspender requisición " & txtNumReq.Text & "?", (vbYesNo + vbQuestion), "Mensaje") = vbYes Then
        If fblnRevisaPermiso(vglngNumeroLogin, flngObtenOpcion("cmdRequisicionPaciente"), "C") Then
          EntornoSIHO.ConeccionSIHO.BeginTrans
          vlstrx = "UPDATE IvRequisicionMaestro "
          vlstrx = vlstrx & " SET BITDETIENESALIDA = 1"
          vlstrx = vlstrx & " WHERE numNumRequisicion = " & txtNumReq.Text
          pEjecutaSentencia (vlstrx)
          Call pGuardarLogTransaccion(Me.Name, EnmCambiar, vglngNumeroLogin, "CARGO PACIENTES (SUSPENSION)", txtNumReq.Text)
          EntornoSIHO.ConeccionSIHO.CommitTrans
          pIniciatodo
          txtNumReq.SetFocus
        Else
          '¡El usuario no tiene permiso para grabar datos!
          MsgBox SIHOMsg(65), vbOKOnly + vbExclamation, "Actualización de máximos y mínimos"
        End If
      End If
    End If
  End If
  rsReqMaestro.Close
End Sub

Private Sub cmdUltimoRegistro_Click()
    If grdHBusqueda.Rows - 1 > 0 Then
        grdHBusqueda.Row = (grdHBusqueda.Rows - 1)
        Call pConsultaReq(grdHBusqueda.TextMatrix(grdHBusqueda.Row, 1))
    Else
        cboEstatus_Click
        If grdHBusqueda.Rows - 1 > 0 Then
            grdHBusqueda.Row = (grdHBusqueda.Rows - 1)
            Call pConsultaReq(grdHBusqueda.TextMatrix(grdHBusqueda.Row, 1))
        End If
    End If
End Sub

Private Sub dtpFechaFin_CloseUp()
    pLlenarBusqueda
End Sub

Private Sub grdHBusqueda_DblClick()
        If grdHBusqueda.Rows - 1 > 0 Then
        Call pConsultaReq(grdHBusqueda.TextMatrix(grdHBusqueda.Row, 1))
        sstObj.Tab = 0
    End If
End Sub

Private Sub mskFechaFin_GotFocus()
    mskFechaFin.Format = "dd/MM/yyyy"
    mskFechaFin.SelStart = 0
    mskFechaFin.SelLength = Len(mskFechaInicio.Text)
End Sub

Private Sub mskFechaFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub mskFechaFin_LostFocus()
    mskFechaFin.Format = "dd/MMM/yyyy"
End Sub

Private Sub mskFechaInicio_CloseUp()
    pLlenarBusqueda
End Sub

Private Sub mskFechaInicio_GotFocus()
    mskFechaInicio.Format = "dd/MM/yyyy"
    mskFechaInicio.SelStart = 0
    mskFechaInicio.SelLength = Len(mskFechaInicio.Text)
End Sub

Private Sub mskFechaInicio_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub mskFechaInicio_LostFocus()
    mskFechaInicio.Format = "dd/MMM/yyyy"
    pLlenarBusqueda
End Sub

Private Sub Form_Activate()
    If vlblnPrimeraVez And vlblnValidarContrasena Then
        vllngLoginInicial = flngPersonaGraba(vgintNumeroDepartamento, vgstrGrabaMedicoEmpleado)
        If vllngLoginInicial = 0 Then
            Unload Me
            Exit Sub
        Else
            vlblnPrimeraVez = False
        End If
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If sstObj.Tab = 1 Then
            sstObj.Tab = 0
            If txtNumReq.Enabled Then
                txtNumReq.SetFocus
            Else
                If cboTipoReq.Enabled Then
                    cboTipoReq.SetFocus
                End If
            End If
        Else
            If vlblnNuevoReg Then   'Se esta ingresando un nuevo registro
                If MsgBox(SIHOMsg("10"), vbQuestion + vbYesNo, "Mensaje") = vbYes Then  'Desea cancelar el nuevo ingreso
                    Call pIniciatodo
                    txtNumReq.SetFocus
                End If
            Else
                If vlblnConsultaReg Then    'Se esta consultando un registro
                    If MsgBox(SIHOMsg("17"), vbQuestion + vbYesNo, "Mensaje") = vbYes Then  'Desea cancelar la consulta
                        Call pIniciatodo
                        txtNumReq.SetFocus
                    End If
                Else
                    pCierraObjCmd
                    Unload Me
                End If
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim rsParametro As New ADODB.Recordset
    Dim rsParametro2 As New ADODB.Recordset
    Dim vlstrsql As String
    Dim vlintCnt As Integer
    
    Me.Icon = frmMenuPrincipal.Icon
    mskFechaFin = Date
    mskFechaInicio = DateAdd("d", -3, mskFechaFin)
    
    Set rsParametro = frsSelParametros("EX", -1, "BITPERMITEMEDICAMENTOS")
    If Not rsParametro.EOF Then
        If IsNull(rsParametro!Valor) Then
            lblnPermitirMedicamentos = False
            optMedicamento.Enabled = lblnPermitirMedicamentos
            optTodos.Enabled = lblnPermitirMedicamentos
        Else
            
            lblnPermitirMedicamentos = IIf(rsParametro!Valor = "1", True, False)
            optMedicamento.Enabled = lblnPermitirMedicamentos
            optTodos.Enabled = lblnPermitirMedicamentos
            
            Set rsParametro2 = frsSelParametros("EX", -1, "BITVALIDARCONTRASENAIFPROCESO")
            If Not rsParametro2.EOF Then
                vlblnValidarContrasena = IIf(IsNull(rsParametro2!Valor), False, IIf(rsParametro2!Valor = "1", True, False))
            Else
                vlblnValidarContrasena = False
            End If
            rsParametro2.Close
                
        End If
    Else
        lblnPermitirMedicamentos = False
        optMedicamento.Enabled = False
        optTodos.Enabled = False
        vlblnValidarContrasena = False
    End If
    rsParametro.Close
    
    'bit Autorizar articulos no cubiertos en la requisición de artículos
    lblnAutorizarCargos = False
    Set rsParametro = frsSelParametros("SI", vgintClaveEmpresaContable, "BITAUTORIZARARTICULOSREQUISICION")
    If rsParametro.RecordCount > 0 Then
        lblnAutorizarCargos = IIf(rsParametro!Valor = 1, True, False)
    End If
    rsParametro.Close
    
    'bit maneja cuadro básico
    lblnManejaCuadroBasico = False
    Set rsParametro = frsSelParametros("SI", vgintClaveEmpresaContable, "BITMANEJACUADROBASICO")
    If rsParametro.RecordCount > 0 Then
        lblnManejaCuadroBasico = IIf(Trim(rsParametro!Valor) = "1", True, False)
    End If
    rsParametro.Close
    lblnAutorizacion = False
    If lblnManejaCuadroBasico Then  'si maneja cuadro basico, checa si requiere autorización
        Set rsParametro = frsSelParametros("SI", vgintClaveEmpresaContable, "BITREQUIEREAUTORIZACION")
        If rsParametro.RecordCount > 0 Then
            lblnAutorizacion = IIf(Trim(rsParametro!Valor) = "1", True, False)
        End If
        rsParametro.Close
    End If
    
    vgstrNombreForm = Me.Name 'Nombre del Formulario
 
    vlblnPrimeraVez = True
    ' Recordsets dinámicos
    vlstrx = "select * from IvRequisicionMaestro where numNumRequisicion=null"
    Set rsIvRequisicionMaestro = frsRegresaRs(vlstrx, adLockOptimistic, adOpenDynamic)
    
    vlstrx = "select * from IvRequisicionDetalle where numNumRequisicion=null"
    Set rsIvRequisicionDetalle = frsRegresaRs(vlstrx, adLockOptimistic, adOpenDynamic)
    
    ' Configuración para el manejo de medicamentos
    lintTotalManejos = fintConfigManejosMedicamentos(cboManejoMedicamentos) 'Total de manejos de medicamentos utilizados actualmente
    lintColsTotales = 7 + lintTotalManejos    ' Columnas totales del grid de artículos
    lintColClave = 1 + lintTotalManejos       ' Clave del producto
    lintColNombreArt = 2 + lintTotalManejos   ' Nombre comercial
    lintColCantidad = 3 + lintTotalManejos    ' Cantidad de requisicion
    lintColUnidad = 4 + lintTotalManejos      ' Unidad en la que se pide el producto
    lintColEstatus = 5 + lintTotalManejos     ' Estatus del producto
    lintColCveUnidad = 6 + lintTotalManejos   ' Clave de la unidad en la que se pide el producto
    lintColTipoUnidad = 7 + lintTotalManejos  ' Tipo de unidad A)lterna, M)inima
    lintColCantSurtida = 8 + lintTotalManejos ' Cantidad surtida
    lintColEstatusOr = 9 + lintTotalManejos   ' Estatus original
    lstrTitulos = ""
    For vlintCnt = lintColFixed + 1 To lintColClave - 1
        lstrTitulos = lstrTitulos & "|"
    Next vlintCnt

    position = 0
    pIniciatodo
End Sub

Private Sub pIniciatodo()
    On Error GoTo NotificaError
    
    Dim vlstrSentencia As String
    Dim rs As New ADODB.Recordset
    
    vgstrNombreProcedimiento = "pIniciaTodo"
    
    txtNumReq.Text = CStr(fdblProxNum("IvRequisicionMaestro", "NUMNUMREQUISICION"))
    txtFecha.Text = UCase(CStr(Format(fdtmServerFecha, "dd/mmm/yyyy")))
    chkUrgente.Value = 0
    txtEstatusReq.Text = "PENDIENTE"
    
    pLlenarCboDpto cboDepartamento
    If cboDepartamento.ListCount > 0 Then
        cboDepartamento.ListIndex = fintLocalizaCbo_new(cboDepartamento, CStr(vgintNumeroDepartamento))
    End If
    cboDepartamento.Enabled = False
    txtEmpleado.Text = ""
    vglngPersonaGraba = 0
    
    cboTipoReq.Clear
    vlstrSentencia = "Select distinct chrTipoRequisicion Tipo from ivRequisicionDepartamento " & _
                    " Where intNumeroLogin = " & Trim(Str(vglngNumeroLogin))
    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    
    Do While Not rs.EOF
        If rs!tipo = "CA" Then
            cboTipoReq.AddItem "CARGO A PACIENTE"
        End If
        rs.MoveNext
    Loop
    
    cboPaciente.Clear
    
    chkAplicacion.Value = 1
    
    cboNombreComercial.Clear
    txtFamilia.Text = ""
    txtSubfamilia.Text = ""
    txtNombreGenerico.Text = ""
    optArticulo.Value = False
    optMedicamento.Value = False
    optTodos.Value = False
    txtClaveArticulo.Text = ""
    txtCodigoBarras.Text = ""
    txtUnidad.Text = ""
    
    optAlterna.Value = False
    optMinima.Value = False
    txtExistencia.Text = ""
    txtCantidadArt.Text = ""
    
    lblExisTotalAlm = ""
    lblExisTotalDpto = ""
    lblPedidoSug = ""
    
    vlblnNuevoReg = False
    vlblnConsultaReg = False
    vgStrTipoPaciente = "I"
    vlstrNumCuenta = ""
    vlstrAplicacion = ""
    vlblnBuscar = True
    
    If cboTipoReq.ListCount > 0 Then
        cboTipoReq.Enabled = True
        cboTipoReq.ListIndex = 0
        Call pEscogerCboTipoReq
    Else
        cboTipoReq.Enabled = False
        cboAlmacenSurte.Clear
        cboAlmacenSurte.Enabled = False
    End If
    
    cmdBorraExistencia.Enabled = False
    cmdAgregarGrid.Enabled = False
    cmdCancelar.Enabled = False
    
    cboPaciente.Enabled = False
    
    fraBuscaArticulo.Enabled = False
    fraIntExt.Enabled = False
    Call pIniciaMshFGrid(grdHArticulos)
    Call pLimpiaMshFGrid(grdHArticulos)
    
    txtNumReq.Enabled = True
    txtFecha.Enabled = False
    chkUrgente.Enabled = False
    txtEstatusReq.Enabled = False
    cboAlmacenSurte.Enabled = False
        
    optAlterna.Enabled = True
    optMinima.Enabled = True
    optArticulo.Enabled = True
    optMedicamento.Enabled = lblnPermitirMedicamentos
    optTodos.Enabled = lblnPermitirMedicamentos
    cboNombreComercial.Enabled = True
    txtCodigoBarras.Enabled = True

    txtExistencia.Enabled = True
    txtCantidadArt.Enabled = True
    
    pHabilita 1, 1, 1, 1, 1, 0, 0
    
    sstObj.Tab = 0
    
    lintAutorizacion = 0
    ReDim arrautArticulos(lintAutorizacion)
    
    lintMedicamentosCB = 0
    ReDim arrAutorizadosCB(lintMedicamentosCB)

Exit Sub
NotificaError:
    Call pError
End Sub

Private Sub pLlenarCboDpto(ObjCbo As MyCombo)
'Procedimiento para llenar el combo con los departamentos
On Error GoTo NotificaError
    Dim rsDepartamento As New ADODB.Recordset
    Dim vlstrsql As String
    
    vgstrNombreProcedimiento = "pLlenarCboDpto"
    vlstrsql = "SELECT * from NoDepartamento"
    Set rsDepartamento = frsRegresaRs(vlstrsql, adLockOptimistic, adOpenDynamic)
    pLlenarCboRs_new ObjCbo, rsDepartamento, 0, 1 ' - 1
    rsDepartamento.Close
    
NotificaError:
    Call pError
End Sub

Private Sub pActEstatusCab()
    'Procedimiento para Actualizar el estatus de la cabecera
    Dim vlintseq, vlintCancelada, vlintPendiente, vlintSurtida, vlintSurtidaParcial, vlintFila As Integer
    vlintCancelada = 0
    vlintPendiente = 0
    vlintSurtida = 0
    vlintSurtidaParcial = 0
    vlintFila = grdHArticulos.Row 'Fila actual
    
    grdHArticulos.Redraw = False
    
    For vlintseq = 1 To grdHArticulos.Rows - 1
        If (grdHArticulos.TextMatrix(vlintseq, lintColEstatusOr) = "SURTIDA PARCIAL" Or grdHArticulos.TextMatrix(vlintseq, lintColEstatusOr) = "PENDIENTE") And grdHArticulos.TextMatrix(vlintseq, lintColEstatus) <> "CANCELADA" Then
            Select Case grdHArticulos.TextMatrix(vlintseq, lintColCantSurtida)
                Case Is = 0
                    grdHArticulos.TextMatrix(vlintseq, lintColEstatus) = "PENDIENTE"
                Case Is > 0
                    If (grdHArticulos.TextMatrix(vlintseq, lintColCantidad) - grdHArticulos.TextMatrix(vlintseq, lintColCantSurtida)) = 0 Then
                        grdHArticulos.TextMatrix(vlintseq, lintColEstatus) = "SURTIDA"
                    Else
                        grdHArticulos.TextMatrix(vlintseq, lintColEstatus) = "SURTIDA PARCIAL"
                    End If
            End Select
        End If
        
        If grdHArticulos.TextMatrix(vlintseq, lintColEstatus) = "CANCELADA" Then
            vlintCancelada = vlintCancelada + 1
        ElseIf grdHArticulos.TextMatrix(vlintseq, lintColEstatus) = "PENDIENTE" Then
            vlintPendiente = vlintPendiente + 1
        ElseIf grdHArticulos.TextMatrix(vlintseq, lintColEstatus) = "SURTIDA" Then
            vlintSurtida = vlintSurtida + 1
        ElseIf grdHArticulos.TextMatrix(vlintseq, lintColEstatus) = "SURTIDA PARCIAL" Then
            vlintSurtidaParcial = vlintSurtidaParcial + 1
        End If
    Next vlintseq
    
    If vlintCancelada = (grdHArticulos.Rows - 1) Then
        txtEstatusReq = "CANCELADA"
    ElseIf vlintPendiente = (grdHArticulos.Rows - 1) Then
        txtEstatusReq = "PENDIENTE"
    ElseIf vlintSurtida + vlintCancelada = (grdHArticulos.Rows - 1) Then
        txtEstatusReq = "SURTIDA"
    ElseIf vlintSurtidaParcial = (grdHArticulos.Rows - 1) Then
        txtEstatusReq = "SURTIDA PARCIAL"
    ElseIf vlintSurtida > 0 Or vlintSurtidaParcial > 0 Then
        txtEstatusReq = "SURTIDA PARCIAL"
    ElseIf vlintSurtida > 0 Or vlintSurtidaParcial > 0 Then
        txtEstatusReq = "SURTIDA PARCIAL"
    ElseIf vlintPendiente > 0 And (vlintSurtidaParcial > 0 Or vlintSurtida > 0) Then
        txtEstatusReq = "PENDIENTE"
    End If
    
    grdHArticulos.Row = vlintFila
    grdHArticulos.Redraw = True
End Sub

Private Sub pEscogerCboTipoReq()
'Procedimiento para llenar el combo de tipos de requisicion
On Error GoTo NotificaError
    'El tipo de requisicion puede ser (S)alida a departamento,(R)eubicacion,(P)edido,(C)argo a paciente
    If cboTipoReq.ListCount > 0 And vlblnConsultaReg = False Then
        Select Case Mid(cboTipoReq.List(cboTipoReq.ListIndex), 1, 1)
            Case "C"    'Cargos a Pacientes
                Call pLlenarCboAlmSurte("C")
                lblAlmacenProv.Caption = "Almacén surtirá"
                cboAlmacenSurte.ToolTipText = "Almacén que surtirá"
            Case "S"    'Salida a departamento
                Call pLlenarCboAlmSurte("SD")
                cboAlmacenSurte.Enabled = True
                lblAlmacenProv.Caption = "Almacén surtirá"
                cboAlmacenSurte.ToolTipText = "Almacén que surtirá"
            Case "R"    'Reubicacion
                Call pLlenarCboAlmSurte("R")
                cboAlmacenSurte.Enabled = True
                lblAlmacenProv.Caption = "Almacén surtirá"
                cboAlmacenSurte.ToolTipText = "Almacén que surtirá"
            Case "C"    'Compra - Pedido
                Call pLlenarCboAlmSurte("P")
                cboAlmacenSurte.Enabled = True
                lblAlmacenProv.Caption = "Almacén compra"
                cboAlmacenSurte.ToolTipText = "Almacén que realizará la compra"
            Case Else
                cboAlmacenSurte.Clear
                cboAlmacenSurte.Enabled = False
                fraBuscaArticulo.Enabled = False
                lblAlmacenProv.Caption = "Almacén surtirá"
                cboAlmacenSurte.ToolTipText = "Almacén que surtirá"
        End Select
    Else
        cboTipoReq.Enabled = False
        cboAlmacenSurte.Clear
        cboAlmacenSurte.Enabled = False
    End If
    Exit Sub
NotificaError:
    Call pError
End Sub

Private Sub pLlenarCboAlmSurte(vlDptos As String)
'Procedimiento para llenar el combo de los almacenes que surten, o los que puedo pedir salidas a departament
'o con los almacenes que me pueden hacer alguna reubicacion de algun producto cualquiera

    Dim vlintseq As Integer
    Dim vlstrSentencia As String
    Dim vlstrTipoRequi As String
    Dim rs As New ADODB.Recordset

    Select Case vlDptos
        Case "SD"   'Salida a departamento
            vlstrTipoRequi = "SD"
        Case "R"    'Reubicación
            vlstrTipoRequi = "RE"
        Case "P"    'Compra-Pedido
            vlstrTipoRequi = "CO"
        Case "C"    'Cargo paciente
            vlstrTipoRequi = "CA"
    End Select
    cboAlmacenSurte.Clear
    vlstrSentencia = "Select Nodepartamento.smiCveDepartamento, Nodepartamento.vchDescripcion " & _
                    " from ivRequisicionDepartamento " & _
                    " inner join noDepartamento on ivRequisicionDepartamento.smiCveDepartamento = Nodepartamento.smiCveDepartamento " & _
                    " Where intNumeroLogin = " & Trim(Str(vglngNumeroLogin)) & _
                    " and chrTipoRequisicion = '" & vlstrTipoRequi & "'"
    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    
    Do While Not rs.EOF
        cboAlmacenSurte.AddItem rs!VCHDESCRIPCION
        cboAlmacenSurte.ItemData(cboAlmacenSurte.NewIndex) = rs!smicvedepartamento
        rs.MoveNext
    Loop
    rs.Close
    cboAlmacenSurte.Enabled = IIf(cboAlmacenSurte.ListCount > 0, True, False)
    cboAlmacenSurte.ListIndex = IIf(cboAlmacenSurte.ListCount > 0, 0, -1)
    
Exit Sub
    
NotificaError:
    Call pError
End Sub

Private Sub grdHArticulos_DblClick()
    If grdHArticulos.Rows <> 0 And vlblnNuevoReg Then pModificaCaptura
End Sub

Private Sub grdHArticulos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If grdHArticulos.Rows <> 0 And vlblnNuevoReg Then pModificaCaptura
    End If
End Sub

Private Sub pModificaCaptura()
    vgintCveDeptoCargo = CStr(cboAlmacenSurte.ItemData(cboAlmacenSurte.ListIndex))
    pCargaDatos grdHArticulos.TextMatrix(grdHArticulos.Row, lintColClave), vgintCveDeptoCargo
    txtCantidadArt.Text = grdHArticulos.TextMatrix(grdHArticulos.Row, lintColCantidad)
    If grdHArticulos.TextMatrix(grdHArticulos.Row, lintColCveUnidad) = "A" Then
        optAlterna.Value = True
    Else
        optMinima.Value = True
    End If
    txtCantidadArt.SetFocus
End Sub


Private Sub grdHBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If grdHBusqueda.Rows - 1 > 0 Then
            Call pConsultaReq(grdHBusqueda.TextMatrix(grdHBusqueda.Row, 1))
            sstObj.Tab = 0
        End If
    End If
End Sub

Private Sub mskNumCuentaPac_GotFocus()
    pSelMkTexto mskNumCuentaPac
End Sub

Private Sub mskNumCuentaPac_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        pEnfocaCbo_new cboEstatus
    Else
        If KeyAscii = 8 Then Exit Sub
        If Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9" Then KeyAscii = 0
    End If
End Sub

Private Sub mskNumCuentaPac_LostFocus()
    If Not IsNumeric(mskNumCuentaPac.ClipText) Then
        mskNumCuentaPac.Text = ""
    End If
    pLlenarBusqueda
End Sub

Private Sub optAlterna_Click()
    If txtClaveArticulo.Text <> "" Then txtUnidad.Text = fstrUnidad(txtClaveArticulo.Text, 1)
End Sub

Private Sub optAlterna_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdAgregarGrid.Enabled = True
        cmdAgregarGrid.SetFocus
    End If
End Sub

Private Sub optAmbulatorio_Click()
    If vlblnConsultaReg = False Then
        vgStrTipoPaciente = "I"
        cboPaciente.Clear
        fraBuscaArticulo.Enabled = False
        
        pCargaPacienteAmbulatorio cboPaciente
        
        If cboPaciente.Enabled = True Then
            cboPaciente.ListIndex = -1
            optMinima.Enabled = True
            fraBuscaArticulo.Enabled = True
            cboPaciente.Enabled = True
            cboAlmacenSurte.Enabled = False
            cboTipoReq.Enabled = False
            txtNumReq.Enabled = False
            cboPaciente.SetFocus
        End If
    End If
End Sub

Private Sub optArticulo_Click()
    If lstrComGen = "txtNombreGenerico" Then
        pEnfocaTextBox txtNombreGenerico
    Else
        pEnfocaCbo_new cboNombreComercial
    End If
End Sub

Private Sub Optexterno_Click()
    Dim vlstrCriterio As String
    
    ' Se mandan los parametros que requiere la forma de búsqueda de pacientes
    vgStrTipoPaciente = "E"
    vgstrEstatusPaciente = "E"
    
    
    If vlblnConsultaReg = False Then
    
        cboPaciente.Clear
        fraBuscaArticulo.Enabled = False
        With FrmBusquedaPacientes
            .vgStrTipoPaciente = "E"
            .Caption = .Caption & " Externos"
            .vgblnPideClave = False
            .vgIntMaxRecords = 100
            .vgstrMovCve = "M"
            .optSoloActivos.Enabled = True
            .optSoloActivos.Value = True
            .vgStrOtrosCampos = ", ccEmpresa.vchDescripcion as Empresa, " & _
            " (Select Trim(GnDomicilio.vchCalle)||' '||Trim(GnDomicilio.vchNumeroExterior)||Case When GnDomicilio.vchNumeroInterior Is Null Then '' Else ' Int. '||Trim(GnDomicilio.vchNumeroInterior) End " & _
            " From ExPacienteDomicilio " & _
            " Inner Join GnDomicilio ON ExPacienteDomicilio.intCveDomicilio = GnDomicilio.intCveDomicilio " & _
            " And GnDomicilio.intCveTipoDomicilio = 1 " & _
            " Where ExPacienteDomicilio.intNumPaciente = ExPaciente.intNumPaciente) as Dirección, " & _
            " ExPaciente.dtmFechaNacimiento as FechaNac, " & _
            " (Select GnTelefono.vchTelefono " & _
            " From ExPacienteTelefono " & _
            " Inner Join GnTelefono On ExPacienteTelefono.intCveTelefono = GnTelefono.intCveTelefono " & _
            " And GnTelefono.intCveTipoTelefono = 1 " & _
            " Where ExPacienteTelefono.intNumPaciente = ExPaciente.intNumpaciente) as Teléfono "
            .vgstrTamanoCampo = "950,3400,4100,990,980"
        End With
        vglngCvePaciente = FrmBusquedaPacientes.flngRegresaPaciente()
        
        If vglngCvePaciente > 0 Then
            intDeptoIngreso = 0
            vlstrCriterio = "E|0|" & vglngCvePaciente
            Set ObjRs = frsEjecuta_SP(vlstrCriterio, "sp_IVSelPacientesArea")
            If ObjRs.RecordCount > 0 Then
                intDeptoIngreso = IIf(IsNull(ObjRs!INTCVEDEPTOINGRESO), 0, ObjRs!INTCVEDEPTOINGRESO)
                Do While Not ObjRs.EOF
                    cboPaciente.AddItem Trim(IIf(IsNull(ObjRs!Cuarto), "", ObjRs!Cuarto) & " " & ObjRs!Nombre & " (" & IIf(IsNull(ObjRs!PROCEDENCIA), "", ObjRs!PROCEDENCIA)) & ")"
                    cboPaciente.ItemData(cboPaciente.NewIndex) = ObjRs!cuenta
                    ObjRs.MoveNext
                Loop
                cboPaciente.ListIndex = IIf(cboPaciente.ListCount > 0, 0, -1)
                optMinima.Enabled = True
                fraBuscaArticulo.Enabled = True
                cboPaciente.Enabled = True
                cboAlmacenSurte.Enabled = False
                cboTipoReq.Enabled = False
                txtNumReq.Enabled = False
                cboPaciente.SetFocus
            End If
            ObjRs.Close
        End If
    End If
End Sub

Private Sub Optexterno_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Optexterno_Click
    End Select
End Sub

Private Sub Optinterno_Click()
    If vlblnConsultaReg = False Then
        vgStrTipoPaciente = "I"
        cboPaciente.Clear
        fraBuscaArticulo.Enabled = False
        
        Call pCargaPaciente(vgstrArea, cboPaciente)
        
        If cboPaciente.Enabled = True Then
            cboPaciente.ListIndex = -1
            optMinima.Enabled = True
            fraBuscaArticulo.Enabled = True
            cboPaciente.Enabled = True
            cboAlmacenSurte.Enabled = False
            cboTipoReq.Enabled = False
            txtNumReq.Enabled = False
            cboPaciente.SetFocus
        End If
    End If
End Sub

Private Sub Optinterno_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Optinterno_Click
    End Select
End Sub

Private Sub optMedicamento_Click()
    If lstrComGen = "txtNombreGenerico" Then
        pEnfocaTextBox txtNombreGenerico
    Else
        pEnfocaCbo_new cboNombreComercial
    End If
End Sub

Private Sub optMinima_Click()
    If txtClaveArticulo.Text <> "" Then txtUnidad.Text = fstrUnidad(txtClaveArticulo.Text, 0)
End Sub

Private Sub optMinima_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdAgregarGrid.Enabled = True
        cmdAgregarGrid.SetFocus
    End If
End Sub

Private Sub optTodos_Click()
    If lstrComGen = "txtNombreGenerico" Then
        pEnfocaTextBox txtNombreGenerico
    Else
        pEnfocaCbo_new cboNombreComercial
    End If
End Sub

Private Sub SSTObj_Click(PreviousTab As Integer)
    Select Case sstObj.Tab
        Case 0
            fraTab0.Enabled = True
            fraTab1.Enabled = False
        Case 1
            fraTab0.Enabled = False
            fraTab1.Enabled = True
    End Select
End Sub

Private Sub txtCantidadArt_GotFocus()
    Call pSelTextBox(txtCantidadArt)
End Sub

Private Sub txtCantidadArt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If optAlterna.Value And optAlterna.Enabled Then
            optAlterna.SetFocus
        Else
            If optMinima.Value And optMinima.Enabled Then
                optMinima.SetFocus
            Else
                cboNombreComercial.SetFocus
            End If
        End If
    ElseIf KeyAscii = 46 Then
        KeyAscii = 0
    Else
        Call pValidaNumero(KeyAscii)
    End If
End Sub

Private Sub txtCodigoBarras_GotFocus()
    pSelTextBox txtCodigoBarras
End Sub

Private Sub txtCodigoBarras_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If cboPaciente.Text <> "" Then
            If txtCodigoBarras.Text <> "" Then
                If Len(txtCodigoBarras.Text) > 0 Then
                    saSeleccion = saEnProceso ' Se está en proceso de seleccionar un artículo
                    vgstrVarIntercam = UCase(txtCodigoBarras.Text) ' Variable global de entrada al frmlista para la busqueda
                    vgintCveDeptoCargo = CStr(cboAlmacenSurte.ItemData(cboAlmacenSurte.ListIndex))
                    frmListaPaciente.lblnCuadroBasico = lblnManejaCuadroBasico
                    frmListaPaciente.lblnAutorizacionMedicamento = lblnAutorizacion
                    frmListaPaciente.vgblnNombreGenerico = 2
                    frmListaPaciente.vgstrCriterioParam = IIf(Not lblnPermitirMedicamentos, "0", IIf(optTodos.Value, "-1", IIf(optMedicamento.Value, "1", "0")))
                    frmListaPaciente.lblnCuadroBasico = lblnManejaCuadroBasico
                    frmListaPaciente.lblnAutorizacionMedicamento = lblnAutorizacion
                    llngCvePersonaAutoriza = 0
                    lstrTipoPersonaAutoriza = ""
                    lstrFechaAutorizacion = ""
                    frmListaPaciente.Show vbModal, Me
                    If Len(vgstrVarIntercam) > 0 Then
                        pCargaDatos Trim(vgstrVarIntercam), vgintCveDeptoCargo
                        If lblnAutorizacion And llngCvePersonaAutoriza <> 0 Then
                            pDatosAutorizacion llngCvePersonaAutoriza, lstrTipoPersonaAutoriza, Trim(vgstrVarIntercam), lstrFechaAutorizacion
                        End If
                    Else
                        If cboNombreComercial.ListCount > 0 Then cboNombreComercial.ListIndex = 0
                    End If
                Else
                    If cboNombreComercial.ListCount > 0 Then
                        cboNombreComercial.ListIndex = 0
                    End If
                    cboNombreComercial.SetFocus
                End If
            Else
                cboNombreComercial.SetFocus
            End If
        Else
            cboPaciente.SetFocus
        End If
    End If
End Sub

Private Sub txtCodigoBarras_KeyPress(KeyAscii As Integer)
    pValidaSoloNumero KeyAscii
End Sub

Private Sub txtCodigoBarras_Validate(Cancel As Boolean)
    If saSeleccion <> saSeleccionado And txtCodigoBarras.Text <> "" Then Cancel = True
End Sub

Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
        Case Else
            Call pProcessMkTextMshFGrd(txtEdit, "Keydown", KeyCode, 0, grdHArticulos, 0, "", Me)
            If KeyCode = vbKeyReturn Then
                 If Len(txtEdit) > 0 Then
                    If CLng(txtEdit) = 0 Then
                        Call MsgBox("La cantidad solicitada debe ser mayor a cero", vbExclamation, "Mensaje")
                        grdHArticulos.TextMatrix(grdHArticulos.Row, lintColCantidad) = vllngCantidad
                        vllngCantidad = 0
                    Else
                        If CLng(txtEdit) < grdHArticulos.TextMatrix(grdHArticulos.Row, lintColCantSurtida) Then
                            Call MsgBox("La cantidad solicitada debe ser mayor o igual a la surtida", vbExclamation, "Mensaje")
                            grdHArticulos.TextMatrix(grdHArticulos.Row, lintColCantidad) = vllngCantidad
                            vllngCantidad = 0
                        End If
                    End If
                Else
                End If
            End If
    End Select
    Call pActEstatusCab
End Sub

Private Sub txtEdit_LostFocus()
    Call pProcessMkTextMshFGrd(txtEdit, "LostFocus", 13, 0, grdHArticulos, 0, "", Me)
End Sub

Private Sub txtExistencia_GotFocus()
    txtCantidadArt.SetFocus
End Sub

Private Sub txtNombreGenerico_GotFocus()
    pSelTextBox txtNombreGenerico
End Sub

Private Sub txtNombreGenerico_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If cboPaciente.Text <> "" Then
            If txtNombreGenerico.Text <> "" Then
                saSeleccion = saEnProceso
                vgstrVarIntercam = UCase(txtNombreGenerico.Text) ' Variable global de entrada al frmlista para la busqueda
                vgintCveDeptoCargo = CStr(cboAlmacenSurte.ItemData(cboAlmacenSurte.ListIndex))
                frmListaPaciente.lblnCuadroBasico = lblnManejaCuadroBasico
                frmListaPaciente.lblnAutorizacionMedicamento = lblnAutorizacion
                frmListaPaciente.vgblnNombreGenerico = 1
                frmListaPaciente.vgstrCriterioParam = IIf(Not lblnPermitirMedicamentos, "0", IIf(optTodos.Value, "-1", IIf(optMedicamento.Value, "1", "0")))
                frmListaPaciente.lblnCuadroBasico = lblnManejaCuadroBasico
                frmListaPaciente.lblnAutorizacionMedicamento = lblnAutorizacion
                llngCvePersonaAutoriza = 0
                lstrTipoPersonaAutoriza = ""
                lstrFechaAutorizacion = ""
                frmListaPaciente.Show vbModal, Me
                If Len(vgstrVarIntercam) > 0 Then
                    pCargaDatos Trim(vgstrVarIntercam), vgintCveDeptoCargo
                    If lblnAutorizacion And llngCvePersonaAutoriza <> 0 Then
                        pDatosAutorizacion llngCvePersonaAutoriza, lstrTipoPersonaAutoriza, Trim(vgstrVarIntercam), lstrFechaAutorizacion
                    End If
                End If
            Else
                If txtCodigoBarras.Enabled Then txtCodigoBarras.SetFocus
            End If
        Else
            If cboPaciente.Enabled Then cboPaciente.SetFocus
        End If
    End If
End Sub

Private Sub txtNombreGenerico_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNombreGenerico_Validate(Cancel As Boolean)
    If saSeleccion <> saSeleccionado And txtNombreGenerico.Text <> "" Then Cancel = True
    lstrComGen = "txtNombreGenerico"
End Sub

Private Sub txtNumReq_GotFocus()
    Call pSelTextBox(txtNumReq)
    cboAlmacenSurte.Enabled = False
End Sub

Private Sub txtNumReq_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo NotificaError
    Dim vllngNumReq As Long
    
    If KeyCode = vbKeyReturn Then
        If Len(txtNumReq.Text) > 0 Then
            Set ObjRs = frsRegresaRs("Select * from IvRequisicionMaestro where numNumRequisicion = " & txtNumReq & " And chrDestino = 'P'")
            vllngNumReq = ObjRs.RecordCount
            ObjRs.Close
            If vllngNumReq = 1 Then
                'Consultar la requisicion
                Call pConsultaReq(txtNumReq.Text)
                If grdHArticulos.Rows > 0 Then
                    Call pSelTextBox(txtNumReq)
                Else
                    txtNumReq.Text = CStr(fdblProxNum("IvRequisicionMaestro", "NumNumRequisicion"))
                    Call pIniciatodo
                    txtNumReq.SetFocus
                    Call pSelTextBox(txtNumReq)
                End If
            Else
                If CStr(fdblProxNum("IvRequisicionMaestro", "NumNumRequisicion")) = txtNumReq Then 'Se genera un nuevo registro
                    chkUrgente.Enabled = True
                    txtNumReq.Enabled = False
                    txtFecha.Enabled = False
                    txtEstatusReq.Enabled = False
                    cboTipoReq.Enabled = cboTipoReq.ListCount > 0
                    cboAlmacenSurte.Enabled = cboAlmacenSurte.ListCount > 0
                    chkUrgente.SetFocus
                    vlblnNuevoReg = True
                    vlblnConsultaReg = False
                Else
                    Call MsgBox(SIHOMsg("13"), vbExclamation, "Message")    'La informacion No existe
                    txtNumReq.Text = CStr(fdblProxNum("IvRequisicionMaestro", "NumNumRequisicion"))
                    Call pIniciatodo
                    txtNumReq.SetFocus
                    Call pSelTextBox(txtNumReq)
                End If
            End If
        Else
            txtNumReq.Text = CStr(fdblProxNum("IvRequisicionMaestro", "NumNumRequisicion"))
            Call pIniciatodo
            txtNumReq.SetFocus
            Call pSelTextBox(txtNumReq)
        End If
    End If
    Exit Sub
NotificaError:
    Call pError
End Sub

Private Sub txtNumReq_KeyPress(KeyAscii As Integer)
    Call pValidaSoloNumero(KeyAscii)
End Sub

Private Sub pimprimir(vlstrNumReq As String, vlstrDestinoImpresion As String)
    'Procedimiento para Imprimir
    Dim RSIvRptRequisicionCargoPaciente As New ADODB.Recordset
    Dim alstrParametros(1) As String

    If RSIvRptRequisicionCargoPaciente.State = 1 Then
        RSIvRptRequisicionCargoPaciente.Close
    End If
    
    pInstanciaReporte vgrptReporte, "rptRequisicionCargo.rpt"
    vgrptReporte.DiscardSavedData
    alstrParametros(0) = "muestracb;" & gintMuestraCodBar
    alstrParametros(1) = "empresa;" & Trim(vgstrNombreHospitalCH)
    pCargaParameterFields alstrParametros, vgrptReporte

    Set RSIvRptRequisicionCargoPaciente = frsEjecuta_SP(CLng(vlstrNumReq) & "|" & vgintNumeroDepartamento, "SP_IvRptRequisicioCargoPacient")
    
    If RSIvRptRequisicionCargoPaciente.RecordCount > 0 Then
        pImprimeReporte vgrptReporte, RSIvRptRequisicionCargoPaciente, IIf(vlstrDestinoImpresion = "P", "P", "I"), Me.Caption
    Else
        Call MsgBox(SIHOMsg("13"), (vbExclamation), "Mensaje") '!No existe información!
    End If
    
    RSIvRptRequisicionCargoPaciente.Close
End Sub

Private Sub pGrabaRegistros()
On Error GoTo NotificaError
    Dim cmd As New ADODB.Command
    Dim vllngDetalle As Long
    Dim X As Long
    Dim vllngNumReq As Long
    Dim lintCont As Integer
    
    vgstrNombreProcedimiento = "pGrabaRegistros"
    
    EntornoSIHO.ConeccionSIHO.BeginTrans

    ' Graba Registro Maestro
    With rsIvRequisicionMaestro
        .AddNew
        !smiCveDeptoRequis = cboDepartamento.ItemData(cboDepartamento.ListIndex)
        !intCveEmpleaRequis = vglngPersonaGraba
        !smiCveDeptoAlmacen = cboAlmacenSurte.ItemData(cboAlmacenSurte.ListIndex)
        !dtmFechaRequisicion = fdtmServerFecha
        !DTMHORAREQUISICION = fdtmServerHora
        !vchEstatusRequis = "PENDIENTE"
        !chrDestino = "P"
        !numNumCuenta = cboPaciente.ItemData(cboPaciente.ListIndex)
        !bitUrgente = IIf(chkUrgente.Value, 1, 0)
        !NUMNUMREQUISREL = 0
        !CHRTIPOPACIENTE = vgStrTipoPaciente
        !chrAplicacionMed = IIf(chkAplicacion.Value, 1, 0)
        !smiCveDeptoGenera = vgintNumeroDepartamento
        !INTCVEDEPTOINGRESO = intDeptoIngreso
        !INTCVESOLICITUDMEDICOREQUIS = vlintMedicoTratante
        .Update
        vllngNumReq = flngObtieneIdentity("SEC_IVREQUISICIONMAESTRO", !numnumRequisicion)
        Call pGuardarLogTransaccion(Me.Name, EnmGrabar, vglngNumeroLogin, "CARGO PACIENTES", Str(vllngNumReq))
    End With
    
    pEliminaDuplicados grdHArticulos
    
    With rsIvRequisicionDetalle
        For X = 1 To grdHArticulos.Rows - 1
            .AddNew
            !numnumRequisicion = CDbl(vllngNumReq)
            !CHRCVEARTICULO = grdHArticulos.TextMatrix(X, lintColClave)
            !IntCantidadSolicitada = Val(grdHArticulos.TextMatrix(X, lintColCantidad))
            !chrUnidadControl = grdHArticulos.TextMatrix(X, lintColCveUnidad)
            !vchEstatusDetRequis = "PENDIENTE"
            .Update
        
            If lblnAutorizarCargos Then ' Guarda en ExAutorizacionCargoRequi
                For lintCont = 0 To UBound(arrautArticulos)
                    If arrautArticulos(lintCont).StrCveArticulo = grdHArticulos.TextMatrix(X, lintColClave) Then
                        frsEjecuta_SP vllngNumReq & "|" & arrautArticulos(lintCont).StrCveArticulo & "|" & arrautArticulos(lintCont).strCodigo & "|" & IIf(arrautArticulos(lintCont).blnExcluido, "1", "0"), "Sp_Exinsautorizacioncargorequi", True
                        Exit For
                    End If
                Next lintCont
            End If
            ' Se maneja cuadro basico de medicamentos y requiere de autorización, guarda los datos de la autorización
            If lblnAutorizacion Then
                For lintCont = 0 To UBound(arrAutorizadosCB)
                    If Trim(arrAutorizadosCB(lintCont).strCveMedicamento) = Trim(grdHArticulos.TextMatrix(X, lintColClave)) Then
                        vgstrParametrosSP = "1|" & arrAutorizadosCB(lintCont).lngPersona & "|" & _
                                            Trim(arrAutorizadosCB(lintCont).strTipoPersona) & "|" & _
                                            Trim(grdHArticulos.TextMatrix(X, lintColClave)) & "|" & _
                                            vllngNumReq & "|" & Trim(arrAutorizadosCB(lintCont).strFechaAutorizacion) & "|RP"
                        frsEjecuta_SP vgstrParametrosSP, "SP_SIINSPROCESOAUTORIZADO"
                        Exit For
                    End If
                Next lintCont
            End If
        Next X
    End With
    
    ' Ejecuta grabada en ImpresionRemota
    pImpresionRemota "RP", vllngNumReq, cboAlmacenSurte.ItemData(cboAlmacenSurte.ListIndex)
    EntornoSIHO.ConeccionSIHO.CommitTrans
    pIniciatodo
    txtNumReq.SetFocus
Exit Sub
NotificaError:
    If Err.Number <> 0 Then
        pRegistraError Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":" & vgstrNombreProcedimiento)
    End If
End Sub

Private Sub pGrabarModRegistros()
On Error GoTo NotificaError
''cambioooooooooooooooooooooooooooooooooooooooooooooooo casos 7149-7147
    
    Dim vlstrSentencia As String
    Dim vlstrDestino As String
    Dim vlstrUrgente As String
    Dim vlstrNumReqRel As String
    Dim vlintseq As Integer
    Dim vllngIdentity As Long
    Dim cmd As New ADODB.Command
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim objMaestroRS As New ADODB.Recordset
    Dim objDetalleRS As New ADODB.Recordset
    Dim vlintcontador As Integer
    
    vgstrNombreProcedimiento = "pGrabarModRegistros"
    
    ''revision si la informacion en la base de datos coincide con la informacion del grid
    Set objMaestroRS = frsRegresaRs("Select * from ivrequisicionmaestro where numnumrequisicion = " & txtNumReq.Text)
    If Not objMaestroRS.EOF Then
           
        Select Case chkUrgente
            Case 1
                vlstrUrgente = "1"
            Case 0
                vlstrUrgente = "0"
        End Select
         'Si el articulo es aplicado o no
        Select Case chkAplicacion
            Case 1
                vlstrAplicacion = "1"
            Case 0
                vlstrAplicacion = "0"
        End Select

        EntornoSIHO.ConeccionSIHO.BeginTrans
                    
        
        ''  debemos revisar los estados de los detalles, si los estatus son iguales
            Set objDetalleRS = frsRegresaRs("select * from ivrequisiciondetalle where numnumrequisicion = " & txtNumReq.Text)
            objDetalleRS.MoveFirst
            
            Do While Not objDetalleRS.EOF
                For vlintcontador = 1 To grdHArticulos.Rows - 1
                    If Trim(objDetalleRS!CHRCVEARTICULO) = Trim(grdHArticulos.TextMatrix(vlintcontador, lintColClave)) Then
                       
                       If Trim(objDetalleRS!vchEstatusDetRequis) = "PENDIENTE" Then  ' si el estado de  la base de datos es pendiente(SI SE PUEDE MODIFICAR)
                       
                          If Trim(grdHArticulos.TextMatrix(vlintcontador, lintColEstatus)) = "CANCELADA" Then 'el estado cambio a cancelada
                          'solo actualizamos el estado de la requisicion
                             vlstrSentencia = " UPDATE  IvRequisicionDetalle "
                             vlstrSentencia = vlstrSentencia & " SET vchEstatusDetRequis = " & "'" & grdHArticulos.TextMatrix(vlintcontador, lintColEstatus) & "'" 'vchEstatusDetRequis
                             vlstrSentencia = vlstrSentencia & " WHERE IvRequisicionDetalle.numNumRequisicion = " & txtNumReq & " and "
                             vlstrSentencia = vlstrSentencia & " IvRequisicionDetalle.chrCveArticulo = '" & grdHArticulos.TextMatrix(vlintcontador, lintColClave) & "' "
                             Call pEjecutaSentencia2(vlstrSentencia) 'Ejecuta detalle
                             
                             If position <> 0 Then
                                position = 0
                                EntornoSIHO.ConeccionSIHO.RollbackTrans
                                Call pConsultaReq(txtNumReq.Text)
                                Exit Sub
                             End If
                             
                          ElseIf grdHArticulos.TextMatrix(vlintcontador, lintColEstatus) = "PENDIENTE" Then
                              
                              If grdHArticulos.TextMatrix(vlintcontador, lintColCantidad) <> objDetalleRS!IntCantidadSolicitada Then ' si cambiaron la cantidad
                                 'solo actualizamos la cantidad del detalle
                                 vlstrSentencia = " UPDATE  IvRequisicionDetalle "
                                 vlstrSentencia = vlstrSentencia & " SET intCantidadSolicitada = " & grdHArticulos.TextMatrix(vlintcontador, lintColCantidad) 'intCantidadSolicitada
                                 vlstrSentencia = vlstrSentencia & " WHERE IvRequisicionDetalle.numNumRequisicion = " & txtNumReq & " and "
                                 vlstrSentencia = vlstrSentencia & " IvRequisicionDetalle.chrCveArticulo = '" & grdHArticulos.TextMatrix(vlintcontador, lintColClave) & "' "
                                 Call pEjecutaSentencia(vlstrSentencia) 'Ejecuta detalle
                              End If
'
                          End If
                       Else ' si el estado del detalle en la base de datos no esta pendiente, revisamos cantidad y estado en el grid(NO DEBEN SER DIFERENTES)
                          If grdHArticulos.TextMatrix(vlintcontador, lintColCantidad) <> objDetalleRS!IntCantidadSolicitada Or _
                             Trim(grdHArticulos.TextMatrix(vlintcontador, lintColEstatus)) <> Trim(objDetalleRS!vchEstatusDetRequis) Then
                            ' si el estado es diferente o la cantidad es diferente pero el estado del detalle es diferente a pendiente
                            'en la base de datos pum mensaje de error, no se puede modificar
                             
                             'La información ha cambiado, consulte de nuevo
                              MsgBox SIHOMsg(33) & " " & SIHOMsg(381), vbOKOnly + vbExclamation, "Mensaje"
                             EntornoSIHO.ConeccionSIHO.RollbackTrans
                             Exit Sub
                          End If
                        End If
                    End If
                    
                Next
             objDetalleRS.MoveNext
         Loop
         
             'Actualiza Registro Maestro
             vlstrSentencia = " UPDATE  IvRequisicionMaestro "
             vlstrSentencia = vlstrSentencia & " SET intCveEmpleaRequis = " & vglngPersonaGraba & "," 'intCveEmpleaRequis
             vlstrSentencia = vlstrSentencia & " chrTipoPaciente ='" & vgStrTipoPaciente & "'," 'strTipoPaciente
             vlstrSentencia = vlstrSentencia & " chrAplicacionMed = '" & vlstrAplicacion & "'," 'bitAplicacionMed
             vlstrSentencia = vlstrSentencia & " bitUrgente = " & vlstrUrgente & "," 'bitUrgente
             vlstrSentencia = vlstrSentencia & " INTCVEDEPTOINGRESO = " & intDeptoIngreso & "," 'INTCVEDEPTOINGRESO
             vlstrSentencia = vlstrSentencia & " INTCVESOLICITUDMEDICOREQUIS = " & vlintMedicoTratante  'INTCVEDEPTOINGRESO
             vlstrSentencia = vlstrSentencia & " WHERE IvRequisicionMaestro.numNumRequisicion = " & txtNumReq
             Call pEjecutaSentencia(vlstrSentencia)
                            
             frsEjecuta_SP txtNumReq.Text, "sp_IVUpdEstatusReqMaestro", True ' SE MODIFICA EL ESTADO DEL MAESTRO DE LA REQUISICION
             EntornoSIHO.ConeccionSIHO.CommitTrans
            'pIniciatodo
            pConsultaReq txtNumReq.Text
              MsgBox SIHOMsg(358), vbOKOnly + vbInformation, "Mensaje"
       End If
          
    Exit Sub
NotificaError:
    If Err.Number <> 0 Then
        EntornoSIHO.ConeccionSIHO.RollbackTrans
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":" & vgstrNombreProcedimiento))
    End If
End Sub

Public Function pEjecutaSentencia2(vlstrSentencia As String)
  'Declaración de variables locales
  Dim vlobjCommand As ADODB.Command
  
  On Error GoTo NotificaError

  Set vlobjCommand = CreateObject("ADODB.Command")
  With vlobjCommand
    Set .ActiveConnection = EntornoSIHO.ConeccionSIHO
    .CommandText = vlstrSentencia
    .Execute
  End With

Exit Function
NotificaError:
  position = InStr(Err.Description, "TicketError")

    If position > 0 Then
      MsgBox "No es posible hacer esta operación, ya se ha generado un ticket en la farmacia subrogada.", vbExclamation, "Mensaje"
    Else
      Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & " :modProcedimientos " & ":pEjecutaSentencia"))
    End If
End Function

Private Sub pConfBusq()
    'Procedimiento para configurar el grid de busqueda
    Dim vllngseq As Long
    
    With grdHBusqueda
        .FormatString = "|Número||Almacén|Fecha|Hora|Estado|Urgente|Cuenta|Cuarto|Paciente"
        .Redraw = False
        .ColWidth(0) = 100
        .ColWidth(1) = 1000      ' Numero requisicion
        .ColWidth(2) = 0        ' Clave almacen
        .ColWidth(3) = 2000     ' Nombre del almacen
        .ColWidth(4) = 1500     ' Fecha
        .ColWidth(5) = 600      ' Hora
        .ColWidth(6) = 1600     ' Estatus
        .ColWidth(7) = 1000      ' Urgente
        .ColWidth(8) = 1000   ' Cuenta
        .ColWidth(9) = 800      ' Cuarto
        .ColWidth(10) = 3400    ' Nombre paciente
        
        .ColAlignment(1) = flexAlignRightCenter
        .ColAlignment(3) = flexAlignLeftCenter
        .ColAlignment(4) = flexAlignLeftCenter
        .ColAlignment(5) = flexAlignLeftCenter
        .ColAlignment(6) = flexAlignLeftCenter
        .ColAlignment(7) = flexAlignCenterCenter
        .ColAlignment(8) = flexAlignRightCenter
        .ColAlignment(9) = flexAlignLeftCenter
        .ColAlignment(10) = flexAlignLeftCenter
        
        .ColAlignmentFixed(1) = flexAlignCenterCenter
        .ColAlignmentFixed(2) = flexAlignCenterCenter
        .ColAlignmentFixed(3) = flexAlignCenterCenter
        .ColAlignmentFixed(4) = flexAlignCenterCenter
        .ColAlignmentFixed(5) = flexAlignCenterCenter
        .ColAlignmentFixed(6) = flexAlignCenterCenter
        .ColAlignmentFixed(7) = flexAlignCenterCenter
        .ColAlignmentFixed(8) = flexAlignCenterCenter
        .ColAlignmentFixed(9) = flexAlignCenterCenter
        .ColAlignmentFixed(10) = flexAlignCenterCenter
        
        For vllngseq = 1 To .Rows - 1
            .TextMatrix(vllngseq, 4) = Format(.TextMatrix(vllngseq, 4), "DD/MMM/YYYY")
            .TextMatrix(vllngseq, 5) = Format(.TextMatrix(vllngseq, 5), "HH:MM")
        Next vllngseq
        .Redraw = True
    End With

End Sub

Private Sub pConsultaReq(vlstrNumReq As String)
'Procedimiento para consultar una requisicion
On Error GoTo NotificaError

    Dim vlstrSentencia As String
    Dim vlstrCriterio As String
    Dim vllngNumReg As Long
    Dim vlintseq, vlintSumPend As Integer
    Dim rsEmpleado As New ADODB.Recordset
    Dim rsSelPaciente As New ADODB.Recordset
    Dim lvchNoReq As String  ' Parametro para el SoredProcedure
    
    lblExisTotalAlm = ""
    lblExisTotalDpto = ""
    lblPedidoSug = ""
    txtCantidadArt.Text = ""
    txtExistencia = ""
    cboPaciente.Clear
    
    If vlstrNumReq <> 0 Then
        Set ObjRs = frsRegresaRs("Select * from IvRequisicionMaestro where numNumRequisicion = " & vlstrNumReq & " and smiCveDeptoRequis = " & CStr(cboDepartamento.ItemData(cboDepartamento.ListIndex)))
    Else
        Set ObjRs = frsRegresaRs("Select * from IvRequisicionMaestro where smiCveDeptoRequis = " & CStr(cboDepartamento.ItemData(cboDepartamento.ListIndex)))
    End If
    
    vllngNumReg = ObjRs.RecordCount
    If vllngNumReg > 0 Then
        vlblnConsultaReg = True
        vlblnNuevoReg = False
        cmdImprimir.Enabled = True
        vlblnBuscar = False
    
        txtNumReq.Text = ObjRs!numnumRequisicion
        txtFecha.Text = UCase(Format(ObjRs!dtmFechaRequisicion, "dd/mmm/yyyy"))
        chkUrgente.Value = IIf(ObjRs!bitUrgente = 1 Or ObjRs!bitUrgente = True, 1, 0)
        
        txtEstatusReq.Text = ObjRs!vchEstatusRequis
        If txtEstatusReq.Text = "SURTIDA PARCIAL" And ObjRs!BITDETIENESALIDA = 1 Then
          txtEstatusReq.Text = txtEstatusReq.Text & " (SUSPENDIDA)"
        End If
        cboDepartamento.ListIndex = fintLocalizaCbo_new(cboDepartamento, CStr(ObjRs!smiCveDeptoRequis))

        If IsNull(ObjRs!intCveEmpleaRequis) Then
            Set rsEmpleado = frsRegresaRs("SELECT vchApellidoPaterno ||' '|| vchApellidoMaterno ||' '|| vchNombre AS Nombre FROM HoMedico WHERE intCveMedico = " & CStr(ObjRs!intCveMedicoRequis))
            txtEmpleado.Text = IIf(rsEmpleado.RecordCount > 0, Trim(rsEmpleado!Nombre), "")
            rsEmpleado.Close
        Else
            Set rsEmpleado = frsRegresaRs("SELECT vchApellidoPaterno ||' '|| vchApellidoMaterno ||' '|| vchNombre AS Nombre FROM noempleado WHERE intCveEmpleado = " & CStr(ObjRs!intCveEmpleaRequis))
            txtEmpleado.Text = IIf(rsEmpleado.RecordCount > 0, Trim(rsEmpleado!Nombre), "")
            rsEmpleado.Close
        End If
        
        cboTipoReq.Clear
        Select Case ObjRs!chrDestino
            Case "D"
                cboTipoReq.AddItem "SALIDA A DEPARTAMENTO"
            Case "R"
                cboTipoReq.AddItem "REUBICACION"
            Case "P"
                cboTipoReq.AddItem "CARGO A PACIENTE"
            Case "C"
                cboTipoReq.AddItem "COMPRA - PEDIDO"
        End Select
        cboTipoReq.ListIndex = 0
        
        vgStrTipoPaciente = IIf(IsNull(ObjRs!CHRTIPOPACIENTE), vgStrTipoPaciente, ObjRs!CHRTIPOPACIENTE)
        chkAplicacion.Value = IIf(ObjRs!chrAplicacionMed = "1", 1, 0)
        vlstrAplicacion = IIf(ObjRs!chrAplicacionMed = "1", "1", "0")
        
        cboAlmacenSurte.Clear
        Call pLlenarCboDpto(cboAlmacenSurte)
        vlblnBuscar = False
        cboAlmacenSurte.ListIndex = fintLocalizaCbo_new(cboAlmacenSurte, ObjRs!smiCveDeptoAlmacen)
        'Paciente
        vglngCvePaciente = ObjRs!numNumCuenta
        
        If vgStrTipoPaciente = "I" Then
            optInterno.Value = True
            Set rsSelPaciente = frsEjecuta_SP("P|0|" & vglngCvePaciente, "sp_IVSelPacientesArea")
            If rsSelPaciente.RecordCount > 0 Then
                Do While Not rsSelPaciente.EOF
                    cboPaciente.AddItem Trim(IIf(IsNull(rsSelPaciente!Cuarto), "", rsSelPaciente!Cuarto) & " " & rsSelPaciente!Nombre & " (" & IIf(IsNull(rsSelPaciente!PROCEDENCIA), "", rsSelPaciente!PROCEDENCIA)) & ")"
                    cboPaciente.ItemData(cboPaciente.NewIndex) = rsSelPaciente!cuenta
                    rsSelPaciente.MoveNext
                Loop
            End If
            rsSelPaciente.Close
        Else
            optExterno.Value = True
 
            Set rsSelPaciente = frsEjecuta_SP("E|0|" & vglngCvePaciente, "sp_IVSelPacientesArea")
            If rsSelPaciente.RecordCount > 0 Then
                Do While Not rsSelPaciente.EOF
                    cboPaciente.AddItem Trim(IIf(IsNull(rsSelPaciente!Cuarto), "", rsSelPaciente!Cuarto) & " " & rsSelPaciente!Nombre & " (" & IIf(IsNull(rsSelPaciente!PROCEDENCIA), "", rsSelPaciente!PROCEDENCIA)) & ")"
                    cboPaciente.ItemData(cboPaciente.NewIndex) = rsSelPaciente!cuenta
                    rsSelPaciente.MoveNext
                Loop
            End If
            rsSelPaciente.Close
        End If
        
        cboPaciente.ListIndex = IIf(cboPaciente.ListCount > 0, 0, -1)
        cboPaciente.Enabled = False
    
        vlblnBuscar = True
        fraBuscaArticulo.Enabled = True
        optArticulo.Value = False
        optMedicamento.Value = False
        optTodos.Value = False
        optAlterna.Enabled = False
        optMinima.Enabled = False
        optArticulo.Enabled = False
        optMedicamento.Enabled = False
        optTodos.Enabled = False
        cboNombreComercial.Enabled = False
        txtExistencia.Enabled = False
        txtCantidadArt.Enabled = False
        txtNumReq.Enabled = False
        txtCodigoBarras.Enabled = False
        chkUrgente.Enabled = IIf(txtEstatusReq = "PENDIENTE", True, False)
        
        cmdSuspender.Enabled = False
        If txtEstatusReq = "PENDIENTE" Or txtEstatusReq = "SURTIDA PARCIAL" Then
            If txtEstatusReq = "SURTIDA PARCIAL" Then cmdSuspender.Enabled = True
            cmdGrabarRegistro.Enabled = True
            chkAplicacion.Enabled = True
        Else
            cmdGrabarRegistro.Enabled = False
            chkAplicacion.Enabled = False
        End If
    Else
        Call MsgBox("La requisición que busca no pertenece a su departamento", vbExclamation, "Mensaje")
    End If
    ObjRs.Close
     
    If vllngNumReg > 0 Then
         lvchNoReq = txtNumReq.Text
         Set ObjRs = frsEjecuta_SP(lvchNoReq, "sp_CaConsultaRequisicion")
        If ObjRs.RecordCount > 0 Then
            Call pIniciaMshFGrid(grdHArticulos)
            'Call pLlenarMshFGrdRs(grdHArticulos, ObjRS)
            Call pLlenarMshFGrdRsManejos(grdHArticulos, ObjRs, lintTotalManejos, lintColClave)
            pConfGrdArt
            'Procedimiento para Activar el boton de cancelar
            vlintSumPend = 0
            For vlintseq = 1 To grdHArticulos.Rows - 1
                If grdHArticulos.TextMatrix(vlintseq, lintColEstatus) = "PENDIENTE" Then
                    vlintSumPend = vlintSumPend + 1
                End If
                
                If lintTotalManejos > 0 And vlintseq > lintColFixed And vlintseq < lintColClave Then
                     grdHArticulos.ColWidth(vlintseq) = 0
                End If
                
                pColorearManejo grdHArticulos, cboManejoMedicamentos, lintColFixed + 1, grdHArticulos.TextMatrix(vlintseq, lintColClave), vlintseq
            Next vlintseq
            cmdCancelar.Enabled = False
            If vlintSumPend > 0 Then
              If txtEstatusReq.Text <> "SURTIDA PARCIAL (SUSPENDIDA)" Then
                cmdCancelar.Enabled = True
              End If
            End If
        End If
        ObjRs.Close
    End If
    
    Exit Sub

NotificaError:
    Call pError
End Sub

Private Sub pConfGrdArt()
On Error GoTo NotificaError

    Dim vlintCnt As Integer
    
    With grdHArticulos
        .Redraw = False
        If vlblnConsultaReg Then
            .FormatString = lstrTitulos & "|Clave|Nombre comercial|Cantidad|Unidad|Estatus|||Recibido"
        Else
            .FormatString = lstrTitulos & "|Clave|Nombre comercial|Cantidad|Unidad|Estatus"
        End If
        
        '.MergeCells = flexMergeRestrictColumns
        For vlintCnt = lintColFixed + 1 To lintColClave - 1
            .ColWidth(vlintCnt) = 0
        Next vlintCnt
        '.MergeRow(0) = True
        
        .ColWidth(lintColFixed) = 100            ' Fixed row
        .ColWidth(lintColClave) = 1000           ' Clave del producto
        .ColWidth(lintColNombreArt) = 4100       ' Nombre comercial
        .ColWidth(lintColCantidad) = 1000        ' Cantidad de requisicion
        .ColWidth(lintColUnidad) = 1500          ' Unidad en la que se pide el producto
        .ColWidth(lintColEstatus) = 1450         ' Estatus del producto
        .ColWidth(lintColCveUnidad) = 0          ' Clave de la unidad en la que se pide el producto
        .ColWidth(lintColTipoUnidad) = 0         ' Tipo de unidad A)lterna, M)inima
        If vlblnConsultaReg Then
            .ColWidth(lintColCantSurtida) = 1000 ' Cantidad surtida
            .ColWidth(lintColEstatusOr) = 0      ' Estatus original
        End If
        .ColAlignment(lintColCantidad) = 7       ' Alinear la columna de cantidad
        
        .Redraw = True
    End With
    Call pDesSelMshFGrid(grdHArticulos)
    
Exit Sub
NotificaError:
    Call pError
End Sub

Private Function fblnAceptarCargoExcluido(ByRef blnExcluido As Boolean, ByRef strCodigo As String) As Boolean
    Dim frmConf As New frmAutorizacion
    fblnAceptarCargoExcluido = frmConf.fblnAceptarCargoExcluido(blnExcluido, strCodigo)
End Function

Private Sub pEliminaDuplicados(ObjGrid As MSHFlexGrid)
On Error GoTo NotificaError
    Dim i As Integer
    Dim j As Integer
    Dim strClaveArticulo As String
    
    i = 1
    Do While i < ObjGrid.Rows
        strClaveArticulo = grdHArticulos.TextMatrix(i, lintColClave)
        j = i + 1
        Do While j < ObjGrid.Rows
            If strClaveArticulo = grdHArticulos.TextMatrix(j, lintColClave) Then
                 ObjGrid.RemoveItem j
                 j = i
            End If
        j = j + 1
        Loop
    i = i + 1
    Loop
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pEliminaDuplicados"))
End Sub

Private Sub Unidades()
On Error GoTo NotificaError

    Dim blnArticuloSel As Boolean   ' Para saber si existe un artículo seleccionado
    Dim rsUnidades As New ADODB.Recordset

    blnArticuloSel = IIf(txtClaveArticulo.Text = "", False, True)
    If blnArticuloSel Then
        Set rsUnidades = frsEjecuta_SP(Trim(UCase(txtClaveArticulo.Text)) & "|" & "|", "sp_IvSelArticulo")
        If rsUnidades.RecordCount <> 0 Then
            rsUnidades.MoveFirst
            optAlterna.Enabled = blnArticuloSel
            optMinima.Enabled = IIf(blnArticuloSel And rsUnidades!UnidadMinima = rsUnidades!UnidadAlterna, False, True)
            If optMinima.Enabled Then
                optMinima.Value = True
            Else
                optAlterna.Value = True
            End If
        Else
            optAlterna.Enabled = True
            optMinima.Enabled = True
            optAlterna.Value = True
        End If
    Else
        optAlterna.Enabled = False
        optMinima.Enabled = False
    End If
        
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pUnidades"))
End Sub

Private Function pRevisaCuentaCerrada()
On Error GoTo NotificaError
    Dim rs As New ADODB.Recordset
    Dim vlstrSentencia As String
    
    pRevisaCuentaCerrada = False
    vgstrParametrosSP = cboPaciente.ItemData(cboPaciente.ListIndex) & "|" & IIf(optInterno.Value = True Or optAmbulatorio.Value = True, "'I'", "'E'") & "|" & CStr(vgintClaveEmpresaContable)
    Set rs = frsEjecuta_SP(vgstrParametrosSP, "Sp_Gnseldatospaciente")
    If rs.RecordCount <> 0 Then
        If rs!CuentaCerrada <> 0 Then
            'La cuenta del paciente está cerrada, no pueden realizarse modificaciones.
            MsgBox SIHOMsg(596), vbExclamation, "Mensaje"
            pRevisaCuentaCerrada = True
            If optInterno.Value Then optInterno.SetFocus
            If optExterno.Value Then optExterno.SetFocus
            If optAmbulatorio.Value Then optAmbulatorio.SetFocus
        End If
    End If
        
    Exit Function
    
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pRevisaCuentaCerrada"))
End Function

' Procedimiento para llenar el grid con las columnas del recordset más las de los manejos de medicamentos '
Private Sub pLlenarMshFGrdRsManejos(ObjGrid As MSHFlexGrid, ObjRs As Recordset, ByRef vlintTotalManejos As Integer, ByRef vlintColInicial As Integer, Optional vlstrColumnaData As String)
On Error GoTo NotificaError
    
    Dim vlintNumCampos As Long 'Total de Columnas
    Dim vlintNumReg As Long    'Total de Renglones
    Dim vlintSeqFil As Long    'Variable para el seguimiento de los renglones
    Dim vlintSeqCol As Long    'Variable para el seguimiento de las columnas
    Dim vlintSeqReg As Long    'Variable para el seguimiento de los registros del recordset
    
    vlintNumCampos = ObjRs.Fields.Count
    vlintNumReg = ObjRs.RecordCount
    
    If vlintNumReg > 0 Then
        With ObjGrid
            .Redraw = False
            .Visible = False
            .Clear
            .ClearStructure
            .Cols = vlintNumCampos + vlintTotalManejos + 1
            .Rows = vlintNumReg + 1
            .FixedCols = 1
            .FixedRows = 1
        
            ObjRs.MoveFirst
            For vlintSeqFil = 1 To vlintNumReg
                vlintSeqReg = 0
                For vlintSeqCol = vlintColInicial To vlintNumCampos + vlintTotalManejos
                    If IsNull(ObjRs.Fields(vlintSeqReg).Value) Then
                        .TextMatrix(vlintSeqFil, vlintSeqCol) = ""
                    Else
                        If vlstrColumnaData <> "" Then
                            If vlintSeqCol - 1 = Val(vlstrColumnaData) Then
                                .RowData(vlintSeqFil) = ObjRs.Fields(vlintSeqCol - 1)
                            End If
                        End If
                        .TextMatrix(vlintSeqFil, vlintSeqCol) = ObjRs.Fields(vlintSeqReg).Value
                    End If
                    vlintSeqReg = vlintSeqReg + 1
                Next vlintSeqCol
                ObjRs.MoveNext
            Next vlintSeqFil
            
            .Redraw = True
            .Visible = True
        End With
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLlenarMshFGrdRsManejos"))
End Sub

'- Función agregada para la configuración de manejos de medicamentos -'
Function fintConfigManejosMedicamentos(cboFuente As MyCombo) As Integer
On Error GoTo NotificaError

    Dim rsManejoMedicamentos As New ADODB.Recordset
    Dim lintCnt As Integer
    Dim lstrSql As String
    
    lintCnt = 0
    '- Traer el total de manejos de medicamentos activos -'
    lstrSql = "SELECT MAX(Manejos) as TotalManejos FROM " & _
               "(SELECT IVARTICULO.INTIDARTICULO, COUNT(IvManejoMedicamento.VCHSIMBOLO) Manejos " & _
               "FROM IVARTICULO " & _
               "LEFT JOIN IVARTICULOMANEJO ON IVARTICULO.INTIDARTICULO = IVARTICULOMANEJO.INTIDARTICULO " & _
               "LEFT JOIN IvManejoMedicamento ON IvManejoMedicamento.intCveManejo = IVARTICULOMANEJO.INTCVEMANEJO " & _
               "AND IvManejoMedicamento.BITACTIVO = 1 " & _
               "WHERE IVARTICULO.VCHESTATUS = 'ACTIVO' " & _
               "GROUP BY IVARTICULO.INTIDARTICULO " & _
               "ORDER BY Manejos DESC) "
    Set rsManejoMedicamentos = frsRegresaRs(lstrSql, adLockOptimistic, adOpenForwardOnly)
    If rsManejoMedicamentos.RecordCount > 0 Then
        lintCnt = rsManejoMedicamentos!TotalManejos
    End If
    rsManejoMedicamentos.Close
    fintConfigManejosMedicamentos = lintCnt
        
    '- Traer los manejos de medicamentos por cada artículo y llenar el combo -'
    lstrSql = "SELECT IVARTICULO.CHRCVEARTICULO, IVARTICULOMANEJO.INTCVEMANEJO " & _
              " FROM IVARTICULOMANEJO " & _
              " LEFT OUTER JOIN IVARTICULO ON IVARTICULOMANEJO.INTIDARTICULO = IVARTICULO.INTIDARTICULO " & _
              " ORDER BY IVARTICULO.CHRCVEARTICULO, IVARTICULOMANEJO.INTCVEMANEJO"
    Set rsManejoMedicamentos = frsRegresaRs(lstrSql, adLockOptimistic, adOpenForwardOnly)
    If rsManejoMedicamentos.RecordCount > 0 Then
        With cboFuente
            .Clear
            rsManejoMedicamentos.MoveFirst
            For lintCnt = 0 To rsManejoMedicamentos.RecordCount - 1
                .AddItem IIf(IsNull(rsManejoMedicamentos!CHRCVEARTICULO), "", rsManejoMedicamentos!CHRCVEARTICULO)
                .ItemData(lintCnt) = IIf(IsNull(rsManejoMedicamentos!intCveManejo), 0, rsManejoMedicamentos!intCveManejo)
                rsManejoMedicamentos.MoveNext
            Next
        End With
    End If
    rsManejoMedicamentos.Close
    
    '- Traer los manejos de medicamentos activos y llenar el arreglo -'
    lstrSql = "SELECT * FROM IvManejoMedicamento WHERE BITACTIVO = 1 ORDER BY 1"
    Set rsManejoMedicamentos = frsRegresaRs(lstrSql, adLockOptimistic, adOpenForwardOnly)
    If rsManejoMedicamentos.RecordCount > 0 Then
        ReDim aManejoMedicamentos(rsManejoMedicamentos.RecordCount)
        rsManejoMedicamentos.MoveFirst
        For lintCnt = 0 To rsManejoMedicamentos.RecordCount - 1
            aManejoMedicamentos(lintCnt).intCveManejo = rsManejoMedicamentos!intCveManejo
            aManejoMedicamentos(lintCnt).strColor = rsManejoMedicamentos!vchColor
            aManejoMedicamentos(lintCnt).strSimbolo = rsManejoMedicamentos!vchSimbolo
            rsManejoMedicamentos.MoveNext
        Next
    End If
    rsManejoMedicamentos.Close
    
    Exit Function
    
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fintConfigManejosMedicamentos"))
End Function

'- Procedimiento para el formato visual de los manejos de los medicamentos -'
Private Sub pColorearManejo(grdFuente As MSHFlexGrid, ByVal cboFuente As MyCombo, ByVal intColumna As Integer, ByVal strCveMedicamento As String, ByVal intRenglon As Integer, Optional intColWidth = 320)
On Error GoTo NotificaError
    
    Dim vlstrClave As String
    Dim vlintlista As Integer, vlintCnt As Integer, vlintCol As Integer
    
    If Trim(strCveMedicamento) = "" Then Exit Sub
    
    vlintCol = intColumna
    For vlintlista = 0 To cboFuente.ListCount - 1
        If (cboFuente.List(vlintlista) = strCveMedicamento) Then
            vlstrClave = cboFuente.ItemData(vlintlista)
            vlintCnt = 0
            Do While vlintCnt < UBound(aManejoMedicamentos)
                If aManejoMedicamentos(vlintCnt).intCveManejo = vlstrClave Then
                    With grdFuente
                        .TextMatrix(intRenglon, vlintCol) = aManejoMedicamentos(vlintCnt).strSimbolo
                        .Row = intRenglon
                        .Col = vlintCol
                        .CellFontBold = False
                        .CellFontName = "Wingdings"
                        .CellFontSize = 12
                        .CellForeColor = CLng(aManejoMedicamentos(vlintCnt).strColor)
                        .CellBackColor = .BackColor
                        .ColAlignment(vlintCol) = flexAlignCenterCenter
                        .ColWidth(vlintCol) = intColWidth
                        vlintCol = vlintCol + 1
                    End With
                End If
                vlintCnt = vlintCnt + 1
            Loop
        End If
    Next vlintlista
    
    Exit Sub
    
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pColorearManejo"))
End Sub

