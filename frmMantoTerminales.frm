VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.1#0"; "MyCommandButton.ocx"
Object = "{FF14BD24-9F8A-41E3-B5B8-7F0D45EE9F16}#15.0#0"; "hsflatcontrols.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMantoTerminales 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Terminales"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7440
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   7440
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   8895
      Left            =   -10
      TabIndex        =   37
      Top             =   -10
      Width           =   7600
      _ExtentX        =   13414
      _ExtentY        =   15690
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   -2147483630
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmMantoTerminales.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Winsock1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "TimerConexion"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmMantoTerminales.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grdTerminales"
      Tab(1).ControlCount=   1
      Begin VB.Timer TimerConexion 
         Left            =   360
         Top             =   7680
      End
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   6480
         Top             =   7680
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdTerminales 
         Height          =   8295
         Left            =   -75000
         TabIndex        =   36
         Top             =   0
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   14631
         _Version        =   393216
         Rows            =   3
         BackColorBkg    =   -2147483643
         BackColorUnpopulated=   -2147483643
         HighLight       =   2
         ScrollBars      =   2
         SelectionMode   =   1
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
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
         _Band(0).ColHeader=   1
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   7345
         Left            =   35
         TabIndex        =   38
         Top             =   30
         Width           =   7345
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
            Left            =   1200
            MaxLength       =   6
            TabIndex        =   0
            ToolTipText     =   "Clave terminal"
            Top             =   240
            Width           =   2295
         End
         Begin HSFlatControls.MyCombo cboCopias 
            Height          =   375
            Left            =   4920
            TabIndex        =   25
            ToolTipText     =   "No. copias por impresion"
            Top             =   6120
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   661
            Style           =   1
            Enabled         =   -1  'True
            Text            =   "MyCombo3"
            Sorted          =   0   'False
            List            =   $"frmMantoTerminales.frx":0038
            ItemData        =   $"frmMantoTerminales.frx":0045
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
         Begin HSFlatControls.MyCombo cboPP 
            Height          =   375
            Left            =   1200
            TabIndex        =   24
            ToolTipText     =   "Tipo de conexion o pind pad"
            Top             =   6120
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   661
            Style           =   1
            Enabled         =   -1  'True
            Text            =   "MyCombo2"
            Sorted          =   0   'False
            List            =   $"frmMantoTerminales.frx":0052
            ItemData        =   $"frmMantoTerminales.frx":0071
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
         Begin VB.TextBox txtPPPort 
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
            Left            =   4920
            TabIndex        =   23
            ToolTipText     =   "Puerto COM conectado a la terminal"
            Top             =   5700
            Width           =   2295
         End
         Begin VB.TextBox txtHTTPPort 
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
            Left            =   4920
            TabIndex        =   21
            ToolTipText     =   "Puerto de la direccion"
            Top             =   5280
            Width           =   2295
         End
         Begin VB.TextBox txtSoPort 
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
            Left            =   4920
            TabIndex        =   19
            ToolTipText     =   "Puerto conectado al servicio"
            Top             =   4860
            Width           =   2295
         End
         Begin VB.TextBox txtTIDUSD 
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
            Left            =   4920
            TabIndex        =   17
            ToolTipText     =   "Token USD"
            Top             =   4440
            Width           =   2295
         End
         Begin VB.TextBox txtTIDMXN 
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
            Left            =   4920
            TabIndex        =   15
            ToolTipText     =   "Token MXN"
            Top             =   4020
            Width           =   2295
         End
         Begin VB.TextBox txtKey 
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
            Left            =   4920
            TabIndex        =   13
            ToolTipText     =   "Token de seguridad"
            Top             =   3600
            Width           =   2295
         End
         Begin VB.TextBox txtStPort 
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
            Left            =   4920
            TabIndex        =   11
            ToolTipText     =   "Puerto del tunel conectado"
            Top             =   3120
            Width           =   2295
         End
         Begin VB.TextBox txtPPIP 
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
            Left            =   1200
            TabIndex        =   22
            ToolTipText     =   "IP o enlaze de la terminal"
            Top             =   5700
            Width           =   2295
         End
         Begin VB.TextBox txtHTTPIP 
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
            Left            =   1200
            TabIndex        =   20
            ToolTipText     =   "Direccion de la IP"
            Top             =   5280
            Width           =   2295
         End
         Begin VB.TextBox txtSoIP 
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
            Left            =   1200
            TabIndex        =   18
            ToolTipText     =   "Socket conectado al servicio"
            Top             =   4860
            Width           =   2295
         End
         Begin VB.TextBox txtNIDUSD 
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
            Left            =   1200
            TabIndex        =   16
            ToolTipText     =   "Identificador USD"
            Top             =   4440
            Width           =   2295
         End
         Begin VB.TextBox txtNIDMXN 
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
            Left            =   1200
            TabIndex        =   14
            ToolTipText     =   "Identificador MXN"
            Top             =   4020
            Width           =   2295
         End
         Begin VB.TextBox txtTimeOut 
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
            Left            =   1200
            TabIndex        =   12
            ToolTipText     =   "Tiempo de espera en transaccion pin pad"
            Top             =   3600
            Width           =   2295
         End
         Begin VB.TextBox txtStIP 
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
            Left            =   1200
            TabIndex        =   10
            ToolTipText     =   "IP del tunel conectado"
            Top             =   3180
            Width           =   2295
         End
         Begin VB.TextBox txtUSD 
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
            Left            =   6600
            MaxLength       =   3
            TabIndex        =   5
            ToolTipText     =   "Transacción de moneda USD"
            Top             =   1500
            Width           =   615
         End
         Begin VB.TextBox txtMXN 
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
            Left            =   5420
            MaxLength       =   3
            TabIndex        =   4
            ToolTipText     =   "Transacción de moneda MXN"
            Top             =   1500
            Width           =   615
         End
         Begin VB.TextBox txtPort 
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
            Left            =   4920
            MaxLength       =   6
            TabIndex        =   9
            ToolTipText     =   "Puerto del servidor"
            Top             =   2340
            Width           =   2295
         End
         Begin VB.TextBox txtPwd 
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
            IMEMode         =   3  'DISABLE
            Left            =   4920
            MaxLength       =   128
            PasswordChar    =   "*"
            TabIndex        =   7
            ToolTipText     =   "Contraseña del servicio"
            Top             =   1920
            Width           =   2295
         End
         Begin VB.TextBox txtIP 
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
            Left            =   1200
            MaxLength       =   128
            TabIndex        =   8
            ToolTipText     =   "IP o direccion de servidor"
            Top             =   2340
            Width           =   2295
         End
         Begin HSFlatControls.MyCombo cboProvider 
            Height          =   375
            Left            =   1200
            TabIndex        =   3
            ToolTipText     =   "Banco o proveedor de servicio de cobro"
            Top             =   1500
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   661
            Style           =   1
            Enabled         =   -1  'True
            Text            =   "MyCombo1"
            Sorted          =   0   'False
            List            =   $"frmMantoTerminales.frx":0081
            ItemData        =   $"frmMantoTerminales.frx":009D
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
         Begin VB.TextBox txtUsr 
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
            Left            =   1200
            MaxLength       =   128
            TabIndex        =   6
            ToolTipText     =   "Usuario del servicio"
            Top             =   1920
            Width           =   2295
         End
         Begin VB.TextBox txtURI 
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
            Left            =   1200
            MaxLength       =   1024
            TabIndex        =   2
            ToolTipText     =   "Dirección URL del socket o puente de comunicación"
            Top             =   1080
            Width           =   6015
         End
         Begin VB.TextBox txtNombre 
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
            Left            =   1200
            MaxLength       =   256
            TabIndex        =   1
            ToolTipText     =   "Nombre entidad bancaria o servicio de la terminal"
            Top             =   660
            Width           =   6015
         End
         Begin MyCommandButton.MyButton cmdConf 
            Height          =   495
            Left            =   4876
            TabIndex        =   28
            ToolTipText     =   "Cambiar conexion"
            Top             =   6720
            Width           =   2338
            _ExtentX        =   4128
            _ExtentY        =   873
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
            TransparentColor=   16777215
            Caption         =   "Cambiar configuración"
            DepthEvent      =   1
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdInitKeys 
            Height          =   495
            Left            =   120
            TabIndex        =   26
            ToolTipText     =   "Inicio de proceso de sincronización"
            Top             =   6720
            Width           =   2338
            _ExtentX        =   4128
            _ExtentY        =   873
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
            TransparentColor=   16777215
            Caption         =   "Inicializar llaves"
            DepthEvent      =   1
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdShow 
            Height          =   495
            Left            =   2498
            TabIndex        =   27
            ToolTipText     =   "Ver configuración de terminal "
            Top             =   6720
            Width           =   2340
            _ExtentX        =   4128
            _ExtentY        =   873
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
            TransparentColor=   16777215
            Caption         =   "Mostrar configuración"
            DepthEvent      =   1
            ShowFocus       =   -1  'True
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
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
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   26
            Left            =   120
            TabIndex        =   66
            Top             =   300
            Width           =   1215
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
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
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   18
            Left            =   120
            TabIndex        =   65
            Top             =   1560
            Width           =   1095
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Moneda"
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
            Index           =   19
            Left            =   3720
            TabIndex        =   64
            Top             =   1560
            Width           =   1215
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "MXN"
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
            Index           =   20
            Left            =   4920
            TabIndex        =   63
            Top             =   1560
            Width           =   495
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "USD"
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
            Index           =   21
            Left            =   6180
            TabIndex        =   62
            Top             =   1560
            Width           =   495
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Port"
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
            Index           =   25
            Left            =   3720
            TabIndex        =   61
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "IP"
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
            Index           =   24
            Left            =   120
            TabIndex        =   60
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000015&
            X1              =   0
            X2              =   7320
            Y1              =   2940
            Y2              =   2940
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Contraseña"
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
            Index           =   23
            Left            =   3720
            TabIndex        =   59
            Top             =   1980
            Width           =   1215
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Usuario"
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
            Index           =   22
            Left            =   120
            TabIndex        =   58
            Top             =   1980
            Width           =   1215
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "URI WebSocket"
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
            Index           =   17
            Left            =   120
            TabIndex        =   57
            Top             =   1140
            Width           =   1215
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Nombre"
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
            Index           =   16
            Left            =   120
            TabIndex        =   56
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Copias"
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
            Index           =   15
            Left            =   3720
            TabIndex        =   55
            Top             =   6180
            Width           =   1215
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Pinpad"
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
            Index           =   14
            Left            =   120
            TabIndex        =   54
            Top             =   6180
            Width           =   1215
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "COM port"
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
            Index           =   13
            Left            =   3720
            TabIndex        =   53
            Top             =   5760
            Width           =   1215
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Pinpad IP"
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
            Index           =   12
            Left            =   120
            TabIndex        =   52
            Top             =   5760
            Width           =   1215
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "HTTP port"
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
            Index           =   11
            Left            =   3720
            TabIndex        =   51
            Top             =   5340
            Width           =   1215
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "HTTP IP"
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
            Index           =   10
            Left            =   120
            TabIndex        =   50
            Top             =   5340
            Width           =   1215
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Socket port"
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
            Index           =   9
            Left            =   3720
            TabIndex        =   49
            Top             =   4920
            Width           =   1455
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Socket IP"
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
            Index           =   8
            Left            =   120
            TabIndex        =   48
            Top             =   4920
            Width           =   1215
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "TID USD"
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
            Index           =   7
            Left            =   3720
            TabIndex        =   47
            Top             =   4500
            Width           =   1215
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "NID USD"
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
            Index           =   6
            Left            =   120
            TabIndex        =   46
            Top             =   4500
            Width           =   1215
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "TID MXN"
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
            Index           =   5
            Left            =   3720
            TabIndex        =   45
            Top             =   4080
            Width           =   1215
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "NID MXN"
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
            Index           =   4
            Left            =   120
            TabIndex        =   44
            Top             =   4080
            Width           =   1215
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Key"
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
            Index           =   3
            Left            =   3720
            TabIndex        =   43
            Top             =   3660
            Width           =   1215
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Timeout"
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
            Index           =   2
            Left            =   120
            TabIndex        =   42
            Top             =   3660
            Width           =   1215
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Stunnel port"
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
            Index           =   1
            Left            =   3720
            TabIndex        =   41
            Top             =   3240
            Width           =   1575
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Stunnel IP"
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
            Index           =   0
            Left            =   120
            TabIndex        =   40
            Top             =   3240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
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
         Left            =   1560
         TabIndex        =   39
         Top             =   7320
         Width           =   4320
         Begin MyCommandButton.MyButton cmdTop 
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
            Picture         =   "frmMantoTerminales.frx":00AA
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   16777215
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmMantoTerminales.frx":0A2C
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdBack 
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
            Picture         =   "frmMantoTerminales.frx":13AE
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   16777215
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmMantoTerminales.frx":1D30
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdLocate 
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
            Picture         =   "frmMantoTerminales.frx":26B2
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   16777215
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmMantoTerminales.frx":3036
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdNext 
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
            Picture         =   "frmMantoTerminales.frx":39BA
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   16777215
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmMantoTerminales.frx":433C
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdEnd 
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
            Picture         =   "frmMantoTerminales.frx":4CBE
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   16777215
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmMantoTerminales.frx":5640
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdSave 
            Height          =   600
            Left            =   3060
            TabIndex        =   34
            ToolTipText     =   "Grabar"
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
            Picture         =   "frmMantoTerminales.frx":5FC2
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   16777215
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmMantoTerminales.frx":6946
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
         Begin MyCommandButton.MyButton cmdDelete 
            Height          =   600
            Left            =   3660
            TabIndex        =   35
            ToolTipText     =   "Eliminar registro"
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
            Picture         =   "frmMantoTerminales.frx":72CA
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   16777215
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmMantoTerminales.frx":7C4C
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
      End
   End
End
Attribute VB_Name = "frmMantoTerminales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents ws As WebSocketWrap.Client
Attribute ws.VB_VarHelpID = -1
Dim intSig As Long
Dim strRespuestaEsperar As String
Dim blnRespuestaEsperar As Boolean
Dim vllngNumeroOpcion As Long
Dim Conectado As Boolean
Dim intTimeout As Integer

Private Sub cboProvider_GotFocus()
    pEdicion
End Sub

Private Sub cmdBack_Click()
    If grdTerminales.Row > 1 Then
        grdTerminales.Row = grdTerminales.Row - 1
        grdTerminales_DblClick
    End If

End Sub

Private Sub cmdConf_Click()
    On Error GoTo Errs
    Dim rs As ADODB.Recordset
    Dim strReturn As String
    Dim intRespLen As Long
    Dim arrDatosPinPad() As String
    Set ws = New WebSocketWrap.Client
    ws.Timeout = intTimeout
    ws.Uri = txtURI.Text & "?host=" & txtIP.Text & "&port=" & txtPort.Text & "&prov=" & cboProvider.Text & "&usr=" & txtUsr.Text & "&pwd=" & txtPwd.Text
    blnRespuestaEsperar = False
    ws.SendMessage "UPDATECONF:" & txtStIP.Text & ":" & txtStPort.Text & ":" & txtTimeOut.Text & ":" & txtKey.Text & ":" & txtNIDMXN.Text & ":" & txtTIDMXN.Text & ":" & txtNIDUSD.Text & ":" & txtTIDUSD.Text & ":" & txtPPPort.Text & ":" & txtSoIP.Text & ":" & txtSoPort.Text & ":" & txtHTTPIP.Text & ":" & txtHTTPPort.Text & ":" & txtPPIP.Text & ":" & cboPP.Text & ":" & cboCopias.Text
    Do While Not blnRespuestaEsperar
        DoEvents
    Loop
    strReturn = strRespuestaEsperar
    If InStr(strReturn, "Error de socket") > 0 Then
        MsgBox "Error de conexión con el socket:" & vbCrLf & txtIP.Text & ":" & txtPort.Text, vbExclamation, "Mensaje"
    Else
        intRespLen = InStr(strReturn, "}") - 14
        strReturn = Mid(strReturn, 13, intRespLen)
        strReturn = Replace(strReturn, "|Respuesta=", "")
        strReturn = Replace(strReturn, "&", "|")
        arrDatosPinPad = Split(strReturn, "|")
        If fstrGetPPData(arrDatosPinPad, "dcs_form") = "UPDATECONF" Then
            Set rs = frsRegresaRs("select * from PVTerminal where intCveTerminal = " & txtClave.Text, adLockOptimistic, adOpenStatic)
            If Not rs.EOF Then
                rs!VCHSTIP = txtStIP.Text
                rs!VCHSTPORT = txtStPort.Text
                rs!intTimeout = txtTimeOut.Text
                rs!VCHKEY = txtKey.Text
                rs!VCHNMXN = txtNIDMXN.Text
                rs!VCHNUSD = txtNIDUSD.Text
                rs!VCHTMXN = txtTIDMXN.Text
                rs!VCHTUSD = txtTIDUSD.Text
                rs!VCHSCIP = txtSoIP.Text
                rs!VCHSCPORT = txtSoPort.Text
                rs!VCHHTIP = txtHTTPIP.Text
                rs!VCHHTPORT = txtHTTPPort.Text
                rs!VCHPPIP = txtPPIP.Text
                rs!VCHPPPORT = txtPPPort.Text
                rs!VCHPPTYPE = cboPP.Text
                rs!INTCOPIES = cboCopias.Text
                rs.Update
            End If
            rs.Close
            MsgBox "Configuración actualizada", vbInformation, "Mensaje"
            cmdShow.SetFocus
            pHabilitaConfig False
        End If
    End If
    Exit Sub
Errs:
    If InStr(Err.Description, "Error de conexión") > 0 Then
        MsgBox "Error de conexión con el Web Socket: " & vbCrLf & txtURI.Text, vbExclamation, "Mensaje"
    Else
        MsgBox Err.Description, vbExclamation, "Mensaje"
    End If

End Sub


Private Sub cmdDelete_Click()
    On Error GoTo Errs
    Dim vllngPersonaGraba As Long
    Dim strSentencia As String
    
    Dim rs As ADODB.Recordset
     If MsgBox("¿Esta seguro de eliminar la configuracion de la terminal?", vbYesNo + vbQuestion, "Mensaje") = vbYes Then
            
              
            
    If fblnRevisaPermiso(vglngNumeroLogin, vllngNumeroOpcion, "E", True) Or fblnRevisaPermiso(vglngNumeroLogin, vllngNumeroOpcion, "C", True) Then
        
         vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
         If vllngPersonaGraba <> 0 Then
         
         
    EntornoSIHO.ConeccionSIHO.BeginTrans
    strSentencia = "update PVFORMAPAGO set INTCVETERMINAL=null where INTCVETERMINAL=" & txtClave.Text
    pEjecutaSentencia strSentencia
    EntornoSIHO.ConeccionSIHO.CommitTrans
          EntornoSIHO.ConeccionSIHO.BeginTrans
    strSentencia = "update PVTERMINALLOG set INTCVETERMINAL=null where INTCVETERMINAL=" & txtClave.Text
    pEjecutaSentencia strSentencia
    EntornoSIHO.ConeccionSIHO.CommitTrans
         
         
         
         
        Set rs = frsRegresaRs("select * from PVTerminal where intCveTerminal = " & txtClave.Text, adLockOptimistic, adOpenStatic)
        'pEjecutaSentencia "delete from PVTerminal where intCveTerminal = " & txtClave.Text
        If Not rs.EOF Then
            rs.Delete
            pConfiguraGrid
            pCargaTerminales
            pGuardarLogTransaccion Me.Name, EnmBorrar, vglngNumeroLogin, "TERMINAL", Mid(txtNombre.Text & "|" & txtURI.Text & "|" & cboProvider.Text & "|" & txtMXN.Text & "|" & txtUSD.Text & "|" & txtUsr.Text & "|" & txtPwd.Text & "|" & txtIP.Text & "|" & txtPort.Text, 1, 2048), txtClave.Text
            MsgBox "El registro ha sido eliminado", vbOKOnly + vbExclamation, "Mensaje"
            txtClave.SetFocus
        End If
        rs.Close
    Else
        MsgBox SIHOMsg(65), vbOKOnly + vbExclamation, "Mensaje"
    End If
    Else
  
               
            End If
            End If
    Exit Sub
Errs:
    MsgBox Err.Description, vbExclamation, "Mensaje"
End Sub

Private Sub cmdInitKeys_Click()
    On Error GoTo Errs
    Dim strReturn As String
    Set ws = New WebSocketWrap.Client
    ws.Timeout = intTimeout
    ws.Uri = txtURI.Text & "?host=" & txtIP.Text & "&port=" & txtPort.Text & "&prov=" & cboProvider.Text & "&usr=" & txtUsr.Text & "&pwd=" & txtPwd.Text
    blnRespuestaEsperar = False
    
    
If cboProvider.ItemData(cboProvider.ListIndex) = 3 Then
    Dim strPublicaUrl As String
   If txtURI.Text <> "" And Conectado = False Then

 Winsock1.Connect txtURI.Text, txtPort.Text 'Conectamos el winsock
 TimerConexion.Interval = 100
 End If
 
    Else
    
    
    
    
    ws.SendMessage "INIKEY"
    
    End If
    
    Do While Not blnRespuestaEsperar
        DoEvents
    Loop
    strReturn = strRespuestaEsperar
    If InStr(strReturn, "Error de socket") > 0 Then
        MsgBox "Error de conexión con el socket:" & vbCrLf & txtIP.Text & ":" & txtPort.Text, vbExclamation, "Mensaje"
    Else
        MsgBox "Inicialización terminada", vbInformation, "Mensaje"
    End If
    Exit Sub
Errs:
    If InStr(Err.Description, "Error de conexión") > 0 Then
        MsgBox "Error de conexión con el Web Socket: " & vbCrLf & txtURI.Text, vbExclamation, "Mensaje"
    Else
        MsgBox Err.Description, vbExclamation, "Mensaje"
    End If
End Sub

Private Sub cmdLocate_Click()
    SSTab1.Tab = 1
    grdTerminales.SetFocus
    centerGrid Me, grdTerminales
End Sub

Private Sub cmdNext_Click()
    If grdTerminales.Row < grdTerminales.Rows - 1 Then
        grdTerminales.Row = grdTerminales.Row + 1
        grdTerminales_DblClick
    End If
End Sub


Private Sub cmdSave_Click()
    Dim rs As New ADODB.Recordset
    Dim blnConsulta As Boolean
    Dim vllngPersonaGraba As Long
    
    
    
    If fblnValida() Then
    
    
    
    
    
    If fblnRevisaPermiso(vglngNumeroLogin, vllngNumeroOpcion, "E", True) Or fblnRevisaPermiso(vglngNumeroLogin, vllngNumeroOpcion, "C", True) Then

        If fblnValida Then
         vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
         If vllngPersonaGraba <> 0 Then
            blnConsulta = True
            Set rs = frsRegresaRs("select * from PVTerminal where intCveTerminal = " & IIf(txtClave.Text = CStr(intSig), "0", txtClave.Text), adLockOptimistic, adOpenStatic)
            If rs.EOF Then
                blnConsulta = False
                rs.AddNew
            End If
            rs!vchNombre = Trim(txtNombre)
            rs!VCHURI = Trim(txtURI.Text)
            rs!INTPROVIDER = cboProvider.ItemData(cboProvider.ListIndex)
            rs!vchMXN = txtMXN.Text
            rs!vchUSD = txtUSD.Text
            rs!VCHUSR = txtUsr.Text
            rs!VCHPWD = txtPwd.Text
            rs!VCHIP = txtIP.Text
            rs!VCHPORT = txtPort.Text
            rs.Update
            If Not blnConsulta Then
                txtClave.Text = flngObtieneIdentity("SEC_PVTERMINAL", 1)
            End If
            rs.Close
            pConfiguraGrid
            pCargaTerminales
            pLimpia
            pGuardarLogTransaccion Me.Name, IIf(blnConsulta, EnmCambiar, EnmGrabar), vglngNumeroLogin, "TERMINAL", Mid(txtNombre.Text & "|" & txtURI.Text & "|" & cboProvider.Text & "|" & txtMXN.Text & "|" & txtUSD.Text & "|" & txtUsr.Text & "|" & txtPwd.Text & "|" & txtIP.Text & "|" & txtPort.Text, 1, 2048), txtClave.Text
            MsgBox "Información guardada satisfactoriamente", vbInformation, "Mensaje"
            pHabilita
            
            cmdLocate.SetFocus
            
        End If
         End If
    Else
        MsgBox SIHOMsg(65), vbOKOnly + vbExclamation, "Mensaje"
    End If
    End If
End Sub

Private Sub cmdShow_Click()
    On Error GoTo Errs
    Dim rs As ADODB.Recordset
    Dim strReturn As String
    Dim intRespLen As Long
    Dim arrDatosPinPad() As String
    Dim strMensajeEnvia As String
   
    
    
    Set ws = New WebSocketWrap.Client
    ws.Timeout = intTimeout
    ws.Uri = txtURI.Text & "?host=" & txtIP.Text & "&port=" & txtPort.Text & "&prov=" & cboProvider.Text & "&usr=" & txtUsr.Text & "&pwd=" & txtPwd.Text
    blnRespuestaEsperar = False
    
    
    'aqui intervine santander
    If cboProvider.ItemData(cboProvider.ListIndex) = 3 Then
    Dim strPublicaUrl As String
   If txtURI.Text <> "" And Conectado = False Then

 Winsock1.Connect txtURI.Text, txtPort.Text 'Conectamos el winsock
 TimerConexion.Interval = 100
 End If

    If txtIP.Text = "https://qa10.mitec.com.mx" Then
    strPublicaUrl = "https://qa3.mitec.com.mx"
    End If
    If txtIP.Text = "https://key.mitec.com.mx" Then
    strPublicaUrl = "https://ssl.e-pago.com.mx"
    End If
    
     strMensajeEnvia = Trim("LOGIN01|" & txtIP.Text & "|" & strPublicaUrl & "|3|" & txtUsr.Text & "|" & txtPwd.Text & "|*")
    
    Else
    strMensajeEnvia = "SHOWCONF"
    End If
    
    
      If cboProvider.ItemData(cboProvider.ListIndex) = 3 Then
      If Conectado Then Winsock1.SendData strMensajeEnvia
      If Conectado = False Then blnRespuestaEsperar = True
    Else
    ws.SendMessage strMensajeEnvia
    End If
    Do While Not blnRespuestaEsperar
        DoEvents
    Loop
    strReturn = strRespuestaEsperar
  
    If InStr(strReturn, "Error de socket") > 0 Then
        MsgBox "Error de conexión con el socket:" & vbCrLf & txtIP.Text & ":" & txtPort.Text, vbExclamation, "Mensaje"
    Else
        If cboProvider.ItemData(cboProvider.ListIndex) = 3 Then ' respuesta santander
         arrDatosPinPad = Split(strReturn, "|")
        If arrDatosPinPad(0) = "CONEXION" Then
        
        txtPPPort.Text = arrDatosPinPad(5)
        cboPP.ListIndex = 2
        txtTIDMXN.Text = "MXN"
        cboCopias.ListIndex = 1
        txtPPPort.Text = "COM9"
        End If
        If arrDatosPinPad(0) = "ERROR" Then
        MsgBox "Error conexion santander:" & arrDatosPinPad(1), vbExclamation, "Mensaje"
        End If
        
        
        
        Else
        
    
        intRespLen = InStr(strReturn, "}") - 14
        strReturn = Mid(strReturn, 13, intRespLen)
        strReturn = Replace(strReturn, "|Respuesta=", "")
        strReturn = Replace(strReturn, "&", "|")
        arrDatosPinPad = Split(strReturn, "|")
        txtStIP.Text = fstrGetPPData(arrDatosPinPad, "IP")
        txtStPort.Text = fstrGetPPData(arrDatosPinPad, "PUERTO")
        txtTimeOut.Text = fstrGetPPData(arrDatosPinPad, "TIMEOUT")
        txtKey.Text = fstrGetPPData(arrDatosPinPad, "LLAVE")
        txtNIDMXN.Text = fstrGetPPData(arrDatosPinPad, "NEGOCIO1")
        txtTIDMXN.Text = fstrGetPPData(arrDatosPinPad, "TERMINAL1")
        txtNIDUSD.Text = fstrGetPPData(arrDatosPinPad, "NEGOCIO2")
        txtTIDUSD.Text = fstrGetPPData(arrDatosPinPad, "TERMINAL2")
        txtSoIP.Text = fstrGetPPData(arrDatosPinPad, "SOCKETIP")
        txtSoPort.Text = fstrGetPPData(arrDatosPinPad, "SOCKETPORT")
        txtHTTPIP.Text = fstrGetPPData(arrDatosPinPad, "HTTP_ADDRESS")
        txtHTTPPort.Text = fstrGetPPData(arrDatosPinPad, "HTTP_PORT")
        txtPPIP.Text = fstrGetPPData(arrDatosPinPad, "SETTINGIPPINPAD")
        txtPPPort.Text = fstrGetPPData(arrDatosPinPad, "PUERTOCOM")
        cboPP.ListIndex = fintLocalizaCritCbo_new(cboPP, fstrGetPPData(arrDatosPinPad, "settingPinpad"))
        cboCopias.ListIndex = fintLocalizaCritCbo_new(cboCopias, fstrGetPPData(arrDatosPinPad, "SETTINGPRINTCOPY"))
            
        Set rs = frsRegresaRs("select * from PVTerminal where intCveTerminal = " & txtClave.Text, adLockOptimistic, adOpenStatic)
        If Not rs.EOF Then
            rs!VCHSTIP = txtStIP.Text
            rs!VCHSTPORT = txtStPort.Text
            rs!intTimeout = txtTimeOut.Text
            rs!VCHKEY = txtKey.Text
            rs!VCHNMXN = txtNIDMXN.Text
            rs!VCHNUSD = txtNIDUSD.Text
            rs!VCHTMXN = txtTIDMXN.Text
            rs!VCHTUSD = txtTIDUSD.Text
            rs!VCHSCIP = txtSoIP.Text
            rs!VCHSCPORT = txtSoPort.Text
            rs!VCHHTIP = txtHTTPIP.Text
            rs!VCHHTPORT = txtHTTPPort.Text
            rs!VCHPPIP = txtPPIP.Text
            rs!VCHPPPORT = txtPPPort.Text
            rs!VCHPPTYPE = cboPP.Text
            rs!INTCOPIES = cboCopias.Text
            rs.Update
        End If
        rs.Close
        pHabilitaConfig True
    End If
    
    End If
    Exit Sub
Errs:
    If InStr(Err.Description, "Error de conexión") > 0 Then
        MsgBox "Error de conexión con el Web Socket: " & vbCrLf & txtURI.Text, vbExclamation, "Mensaje"
    Else
        MsgBox Err.Description, vbExclamation, "Mensaje"
    End If

End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Me.ActiveControl.Name <> "txtClave" Then
            SendKeys vbTab
        End If
    End If
    If KeyAscii = 27 Then
        If SSTab1.Tab = 1 Then
        
            SSTab1.Tab = 0
        Else
        If MsgBox(SIHOMsg(17), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
                    pLimpia
                    txtClave.SetFocus
                    Unload Me
                End If
            
        End If
    End If
End Sub
Private Sub Form_Load()
    Dim rsTemp As ADODB.Recordset
   'Color de Tab
    SetStyle SSTab1.hwnd, 0
    SetSolidColor SSTab1.hwnd, 16777215
    SSTabSubclass SSTab1.hwnd
    Me.Icon = frmMenuPrincipal.Icon
    If cgstrModulo = "PV" Then
        vllngNumeroOpcion = 7046
        
    ElseIf cgstrModulo = "SI" Then
        vllngNumeroOpcion = 4164
       
    End If

    Set rsTemp = frsSelParametros("SI", vgintClaveEmpresaContable, "INTTERMINALSTIMEOUT")
    If Not rsTemp.EOF Then
        intTimeout = CInt(rsTemp!Valor)
    Else
        intTimeout = 300
    End If
    rsTemp.Close
    
    pConfiguraGrid
    pLimpia
    pCargaTerminales
     TimerConexion.Interval = 100
     'centerGrid Me, grdTerminales
     centerComponent Me, Frame1
     centerComponent Me, Frame2
     
     
   
End Sub
Private Sub centerComponent(frm As Form, grpBox As Frame)
    Dim leftnw As Single
    leftnw = (frm.ScaleWidth - grpBox.Width) / 2
    grpBox.Left = leftnw
    
    
    
End Sub
Private Sub centerGrid(frm As Form, grd As MSHFlexGrid)
    Dim leftnw As Single
    Dim topnw As Single
    Dim widthgrd As Single
    Dim heightgrd As Single
    
    widthgrd = frm.ScaleWidth + 50
    heightgrd = frm.ScaleHeight
    
    leftnw = (frm.ScaleWidth - grd.Width) / 2
    topnw = (frm.ScaleHeight - grd.Height) / 2
    
    grd.Height = heightgrd
    grd.Width = widthgrd
    grd.Left = leftnw
    grd.Top = topnw
    
    
    
    
End Sub



Private Sub pLimpia()
    txtClave.Text = intSig
    txtNombre.Text = ""
    txtNombre.Enabled = False
    txtURI.Text = ""
    txtURI.Enabled = False
    cboProvider.ListIndex = -1
    cboProvider.Enabled = False
    txtMXN.Text = ""
    txtMXN.Enabled = False
    txtUSD.Text = ""
    txtUSD.Enabled = False
    txtUsr.Text = ""
    txtUsr.Enabled = False
    txtPwd.Text = ""
    txtPwd.Enabled = False
    txtIP.Text = ""
    txtIP.Enabled = False
    txtPort.Text = ""
    txtPort.Enabled = False
    cmdTop.Enabled = False
    cmdBack.Enabled = False
    cmdLocate.Enabled = True
    cmdNext.Enabled = False
    cmdEnd.Enabled = False
    cmdSave.Enabled = False
    cmdDelete.Enabled = False
    cmdInitKeys.Enabled = False
    cmdShow.Enabled = False
    cmdConf.Enabled = False
    txtStIP.Text = ""
    txtStIP.Enabled = False
    txtStPort.Text = ""
    txtStPort.Enabled = False
    txtTimeOut.Text = ""
    txtTimeOut.Enabled = False
    txtKey.Text = ""
    txtKey.Enabled = False
    txtNIDMXN.Text = ""
    txtNIDMXN.Enabled = False
    txtTIDMXN.Text = ""
    txtTIDMXN.Enabled = False
    txtNIDUSD.Text = ""
    txtNIDUSD.Enabled = False
    txtTIDUSD.Text = ""
    txtTIDUSD.Enabled = False
    txtSoIP.Text = ""
    txtSoIP.Enabled = False
    txtSoPort.Text = ""
    txtSoPort.Enabled = False
    txtHTTPIP.Text = ""
    txtHTTPIP.Enabled = False
    txtHTTPPort.Text = ""
    txtHTTPPort.Enabled = False
    txtPPIP.Text = ""
    txtPPIP.Enabled = False
    txtPPPort.Text = ""
    txtPPPort.Enabled = False
    cboPP.ListIndex = -1
    cboPP.Enabled = False
    cboCopias.ListIndex = -1
    cboCopias.Enabled = False
End Sub

Private Sub pHabilita()
    txtNombre.Enabled = True
    txtURI.Enabled = True
    cboProvider.Enabled = True
    txtMXN.Enabled = True
    txtUSD.Enabled = True
    txtUsr.Enabled = True
    txtPwd.Enabled = True
    txtIP.Enabled = True
    txtPort.Enabled = True
    cmdDelete.Enabled = True
    cmdInitKeys.Enabled = True
    cmdShow.Enabled = True
    cmdLocate.Enabled = True
    cmdSave.Enabled = False
    cmdConf.Enabled = False
End Sub

Private Sub txtIP_KeyPress(KeyAscii As Integer)
  If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 46 Or KeyAscii = 8) Then
        ' Si no es un número, una letra, un punto o una tecla de control, cancelar el evento KeyPress
        KeyAscii = 0
        
    End If
End Sub

Private Sub txtURI_KeyPress(KeyAscii As Integer)
  If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 46 Or KeyAscii = 8) Then
        ' Si no es un número, una letra, un punto o una tecla de control, cancelar el evento KeyPress
        KeyAscii = 0
       
    End If
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

Dim Dato As String
Conectado = True
Winsock1.GetData strRespuestaEsperar
blnRespuestaEsperar = True
End Sub
Private Sub TimerConexion_Timer()
'Si el winsock está conectado, cambiamos la variable a true
DoEvents
If Winsock1.State <> sckConnected Then
Conectado = False
Else
Conectado = True
End If
End Sub
Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'En caso de error cerramos la conexión
Winsock1.Close
End Sub

Private Sub pHabilitaConfig(blnHabilita As Boolean)
    txtStIP.Enabled = blnHabilita
    txtStPort.Enabled = blnHabilita
    txtTimeOut.Enabled = blnHabilita
    txtKey.Enabled = blnHabilita
    txtNIDMXN.Enabled = blnHabilita
    txtTIDMXN.Enabled = blnHabilita
    txtNIDUSD.Enabled = blnHabilita
    txtTIDUSD.Enabled = blnHabilita
    txtSoIP.Enabled = blnHabilita
    txtSoPort.Enabled = blnHabilita
    txtHTTPIP.Enabled = blnHabilita
    txtHTTPPort.Enabled = blnHabilita
    txtPPIP.Enabled = blnHabilita
    txtPPPort.Enabled = blnHabilita
    cboPP.Enabled = blnHabilita
    cboCopias.Enabled = blnHabilita
    cmdConf.Enabled = blnHabilita
End Sub

Private Sub pEdicion()
    cmdTop.Enabled = False
    cmdBack.Enabled = False
    cmdLocate.Enabled = False
    cmdNext.Enabled = False
    cmdEnd.Enabled = False
    cmdSave.Enabled = True
    cmdDelete.Enabled = False
    cmdInitKeys.Enabled = False
    cmdShow.Enabled = False
    cmdConf.Enabled = False
'    txtStIP.Text = ""
'    txtStIP.Enabled = False
'    txtStPort.Text = ""
'    txtStPort.Enabled = False
'    txtTimeOut.Text = ""
'    txtTimeOut.Enabled = False
'    txtKey.Text = ""
'    txtKey.Enabled = False
'    txtNIDMXN.Text = ""
'    txtNIDMXN.Enabled = False
'    txtTIDMXN.Text = ""
'    txtTIDMXN.Enabled = False
'    txtNIDUSD.Text = ""
'    txtNIDUSD.Enabled = False
'    txtTIDUSD.Text = ""
'    txtTIDUSD.Enabled = False
'    txtSoIP.Text = ""
'    txtSoIP.Enabled = False
'    txtSoPort.Text = ""
'    txtSoPort.Enabled = False
'    txtHTTPIP.Text = ""
'    txtHTTPIP.Enabled = False
'    txtHTTPPort.Text = ""
'    txtHTTPPort.Enabled = False
'    txtPPIP.Text = ""
'    txtPPIP.Enabled = False
'    txtPPPort.Text = ""
'    txtPPPort.Enabled = False
'    cboPP.ListIndex = -1
'    cboPP.Enabled = False
'    cboCopias.ListIndex = -1
'    cboCopias.Enabled = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If txtNombre.Enabled Then
        txtClave.SetFocus
        Cancel = True
        
    End If
End Sub

Private Sub MyButton1_Click()

End Sub

Private Sub grdTerminales_DblClick()
    If grdTerminales.TextMatrix(grdTerminales.Row, 1) <> "" Then
        
        pCargaTerminal CLng(grdTerminales.TextMatrix(grdTerminales.Row, 1))
        
        cmdBack.Enabled = True
        cmdNext.Enabled = True
        If SSTab1.Tab = 1 Then
            
            SSTab1.Tab = 0
        End If
        cmdLocate.SetFocus
    End If
End Sub

Private Sub grdTerminales_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
 If grdTerminales.TextMatrix(grdTerminales.Row, 1) <> "" And grdTerminales.TextMatrix(grdTerminales.Row, 2) <> "" Then
        
        pCargaTerminal CLng(grdTerminales.TextMatrix(grdTerminales.Row, 1))
        
        cmdBack.Enabled = True
        cmdNext.Enabled = True
        If SSTab1.Tab = 1 Then
            
            SSTab1.Tab = 0
        End If
        cmdLocate.SetFocus
    End If
    cmdLocate.SetFocus
End If
End Sub


Private Sub txtClave_GotFocus()
    pLimpia
    pEnfocaTextBox txtClave
End Sub

Private Sub txtClave_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
        KeyAscii = 7
    End If
    If KeyAscii = 13 Then
        If txtClave.Text = CStr(intSig) Then
            pHabilita
            SendKeys vbTab
        Else
            If txtClave.Text = "" Then
                txtClave_GotFocus
            Else
                pCargaTerminal CLng(txtClave.Text)
                'SendKeys vbTab
            End If
            
        End If
    End If
End Sub

Private Sub txtIP_GotFocus()
    pEnfocaTextBox txtIP
    pEdicion
End Sub

Private Sub txtMXN_GotFocus()
    pEnfocaTextBox txtMXN
    pEdicion
End Sub

Private Sub txtMXN_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
        KeyAscii = 7
    End If

End Sub

Private Sub txtNombre_GotFocus()
    pEnfocaTextBox txtNombre
    pEdicion
End Sub

Private Sub pConfiguraGrid()
    With grdTerminales
        .Cols = 3
        .Rows = 2
        .ColWidth(0) = 100
        .ColWidth(1) = 800
        .ColWidth(2) = 7000
        .TextMatrix(0, 1) = "Clave"
        .TextMatrix(0, 2) = "Descripción"
    End With
End Sub

Private Sub pCargaTerminales()
    Dim rsTerminales As ADODB.Recordset
    Set rsTerminales = frsRegresaRs("select intCveTerminal, vchNombre from PVTerminal order by intCveTerminal", adLockReadOnly, adOpenForwardOnly)
    Do Until rsTerminales.EOF
        grdTerminales.TextMatrix(grdTerminales.Rows - 1, 1) = rsTerminales!intCveTerminal
        grdTerminales.TextMatrix(grdTerminales.Rows - 1, 2) = rsTerminales!vchNombre
        
        rsTerminales.MoveNext
        If Not rsTerminales.EOF Then
            grdTerminales.Rows = grdTerminales.Rows + 1
            intSig = rsTerminales!intCveTerminal + 1
        End If
    Loop
    If intSig = 0 Then
    
    
    intSig = 2
    End If
    
    rsTerminales.Close
End Sub

Private Sub pCargaTerminal(intCveTerminal As Long)
    Dim rs As ADODB.Recordset
    Set rs = frsRegresaRs("select * from PVTerminal where intCveTerminal = " & intCveTerminal, adLockReadOnly, adOpenForwardOnly)
    If Not rs.EOF Then
        txtClave.Text = rs!intCveTerminal
        txtNombre.Text = rs!vchNombre
        txtURI.Text = rs!VCHURI
        cboProvider.ListIndex = fintLocalizaCbo_new(cboProvider, IIf(IsNull(rs!INTPROVIDER), 0, rs!INTPROVIDER))
        txtMXN.Text = rs!vchMXN
        txtUSD.Text = IIf(IsNull(rs!vchUSD), "", rs!vchUSD)
        txtUsr.Text = IIf(IsNull(rs!VCHUSR), "", rs!VCHUSR)
        txtPwd.Text = IIf(IsNull(rs!VCHPWD), "", rs!VCHPWD)
        txtIP.Text = rs!VCHIP
        txtPort.Text = rs!VCHPORT
        
        txtStIP.Text = IIf(IsNull(rs!VCHSTIP), "", rs!VCHSTIP)
        txtStPort.Text = IIf(IsNull(rs!VCHSTPORT), "", rs!VCHSTPORT)
        txtTimeOut.Text = IIf(IsNull(rs!intTimeout), "", rs!intTimeout)
        txtKey.Text = IIf(IsNull(rs!VCHKEY), "", rs!VCHKEY)
        txtNIDMXN.Text = IIf(IsNull(rs!VCHNMXN), "", rs!VCHNMXN)
        txtTIDMXN.Text = IIf(IsNull(rs!VCHTMXN), "", rs!VCHTMXN)
        txtNIDUSD.Text = IIf(IsNull(rs!VCHNUSD), "", rs!VCHNUSD)
        txtTIDUSD.Text = IIf(IsNull(rs!VCHTUSD), "", rs!VCHTUSD)
        txtSoIP.Text = IIf(IsNull(rs!VCHSCIP), "", rs!VCHSCIP)
        txtSoPort.Text = IIf(IsNull(rs!VCHSCPORT), "", rs!VCHSCPORT)
        txtHTTPIP.Text = IIf(IsNull(rs!VCHHTIP), "", rs!VCHHTIP)
        txtHTTPPort.Text = IIf(IsNull(rs!VCHHTPORT), "", rs!VCHHTPORT)
        txtPPIP.Text = IIf(IsNull(rs!VCHPPIP), "", rs!VCHPPIP)
        txtPPPort.Text = IIf(IsNull(rs!VCHPPPORT), "", rs!VCHPPPORT)
        cboPP.ListIndex = fintLocalizaCritCbo_new(cboPP, IIf(IsNull(rs!VCHPPTYPE), "", rs!VCHPPTYPE))
        cboCopias.ListIndex = fintLocalizaCritCbo_new(cboCopias, IIf(IsNull(rs!INTCOPIES), 0, rs!INTCOPIES))
        pHabilita
        cmdLocate.SetFocus
    Else
        txtClave_GotFocus
    End If
    rs.Close
    
End Sub

Private Sub TxtNombre_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtPort_GotFocus()
    pEnfocaTextBox txtPort
    pEdicion
End Sub

Private Sub txtPort_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
        KeyAscii = 7
    End If

End Sub

Private Sub txtPwd_GotFocus()
    pEnfocaTextBox txtPwd
    pEdicion
End Sub

Private Sub txtURI_GotFocus()
    pEnfocaTextBox txtURI
    pEdicion
End Sub

Private Sub txtUSD_GotFocus()
    pEnfocaTextBox txtUSD
    pEdicion
End Sub

Private Sub txtUSD_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
        KeyAscii = 7
    End If

End Sub

Private Sub txtUsr_GotFocus()
    pEnfocaTextBox txtUsr
    pEdicion
End Sub

Private Sub ws_Answer(answ As String)
    strRespuestaEsperar = answ
    blnRespuestaEsperar = True
End Sub


Private Function fstrGetPPData(arrDatos() As String, strNombre As String) As String
    On Error GoTo Errs
    Dim intIndex As Integer
    For intIndex = 0 To UBound(arrDatos)
        If strNombre = Split(arrDatos(intIndex), "=")(0) Then
            fstrGetPPData = Replace(Split(arrDatos(intIndex), "=")(1), "_", " ")
            Exit Function
        End If
    Next
Errs:
    fstrGetPPData = ""
End Function

Private Function fblnValida() As Boolean
    
    
    fblnValida = True
    
    If Trim(txtNombre.Text) = "" Then
        fblnValida = False
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbOKOnly + vbInformation, "Mensaje"
        
        txtNombre.SetFocus
        Exit Function
    End If
    If Trim(txtURI.Text) = "" Then
        fblnValida = False
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbOKOnly + vbInformation, "Mensaje"
        txtURI.SetFocus
        Exit Function
    End If
   
    If cboProvider.ListIndex = -1 Then
        fblnValida = False
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbOKOnly + vbInformation, "Mensaje"
        cboProvider.SetFocus
        Exit Function
    End If
    
    If Trim(txtMXN.Text) = "" Then
        fblnValida = False
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbOKOnly + vbInformation, "Mensaje"
        txtMXN.SetFocus
        Exit Function
    End If

    If Trim(txtIP.Text) = "" Then
        fblnValida = False
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbOKOnly + vbInformation, "Mensaje"
        txtIP.SetFocus
        Exit Function
    End If
    If Trim(txtPort.Text) = "" Then
        fblnValida = False
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbOKOnly + vbInformation, "Mensaje"
        txtPort.SetFocus
        Exit Function
    End If


End Function

