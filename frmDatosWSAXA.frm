VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.1#0"; "MyCommandButton.ocx"
Object = "{FF14BD24-9F8A-41E3-B5B8-7F0D45EE9F16}#15.0#0"; "HSFlatControls.ocx"
Begin VB.Form frmDatosWSAXA 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datos de captura AXA"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6435
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   6435
   StartUpPosition =   2  'CenterScreen
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
      Height          =   735
      Left            =   50
      TabIndex        =   1
      Top             =   0
      Width           =   6345
      Begin HSFlatControls.MyCombo cboTipoUrgencia 
         Height          =   420
         Left            =   2160
         TabIndex        =   2
         Top             =   240
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   661
         Style           =   1
         Enabled         =   -1  'True
         Text            =   ""
         Sorted          =   -1  'True
         List            =   $"frmDatosWSAXA.frx":0000
         ItemData        =   $"frmDatosWSAXA.frx":0039
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
      Begin VB.Label lblTipoUrgencia 
         BackColor       =   &H80000005&
         Caption         =   "Tipo de urgencia"
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
         TabIndex        =   3
         Top             =   300
         Width           =   2175
      End
   End
   Begin MyCommandButton.MyButton cmdGrabar 
      Height          =   600
      Left            =   3000
      TabIndex        =   0
      ToolTipText     =   "Grabar"
      Top             =   3960
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
      Picture         =   "frmDatosWSAXA.frx":0046
      BackColorOver   =   -2147483633
      BackColorFocus  =   -2147483633
      BackColorDisabled=   -2147483633
      BorderColor     =   -2147483627
      TransparentColor=   16777215
      Caption         =   ""
      DepthEvent      =   1
      PictureDisabled =   "frmDatosWSAXA.frx":09CA
      PictureAlignment=   4
      PictureDisabledEffect=   0
      ShowFocus       =   -1  'True
   End
   Begin TabDlg.SSTab sstObj 
      Height          =   3975
      Left            =   -210
      TabIndex        =   4
      Top             =   760
      Width           =   6720
      _ExtentX        =   11853
      _ExtentY        =   7011
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   5
      TabHeight       =   26
      TabMaxWidth     =   26
      WordWrap        =   0   'False
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmDatosWSAXA.frx":134E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblSintomatologiaUS"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblFolioReceta"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdBuscarICDUS"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtSintomatologiaUS"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtDiagnosticoUS"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtFolioRecetaUS"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtICDUS"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmDatosWSAXA.frx":136A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtICDUR"
      Tab(1).Control(1)=   "txtDiagnosticoUR"
      Tab(1).Control(2)=   "txtFolioRecetaUR"
      Tab(1).Control(3)=   "cmdBuscarICDUR"
      Tab(1).Control(4)=   "Label8"
      Tab(1).Control(5)=   "Label7"
      Tab(1).Control(6)=   "Label2"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "frmDatosWSAXA.frx":1386
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtNumAutEspecialCP"
      Tab(2).Control(1)=   "txtNumAutGralCP"
      Tab(2).Control(2)=   "txtFolioRecetaCP"
      Tab(2).Control(3)=   "Label6"
      Tab(2).Control(4)=   "Label5"
      Tab(2).Control(5)=   "Label4"
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "Tab 3"
      TabPicture(3)   =   "frmDatosWSAXA.frx":13A2
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtDiagnosticoHO"
      Tab(3).Control(1)=   "txtNumAutGralHO"
      Tab(3).Control(2)=   "txtNumAutEspecialHO"
      Tab(3).Control(3)=   "txtFolioRecetaHO"
      Tab(3).Control(4)=   "txtICDHO"
      Tab(3).Control(5)=   "cmdBuscarICDHO"
      Tab(3).Control(6)=   "Label13"
      Tab(3).Control(7)=   "Label12"
      Tab(3).Control(8)=   "Label11"
      Tab(3).Control(9)=   "Label10"
      Tab(3).Control(10)=   "Label9"
      Tab(3).ControlCount=   11
      Begin VB.TextBox txtDiagnosticoHO 
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
         Left            =   -72840
         Locked          =   -1  'True
         TabIndex        =   19
         ToolTipText     =   "Descripción del código ICD"
         Top             =   980
         Width           =   4335
      End
      Begin VB.TextBox txtNumAutGralHO 
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
         Left            =   -70850
         MaxLength       =   7
         TabIndex        =   18
         ToolTipText     =   "Número de autorización general"
         Top             =   1380
         Width           =   2350
      End
      Begin VB.TextBox txtNumAutEspecialHO 
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
         Left            =   -70850
         MaxLength       =   3
         TabIndex        =   17
         ToolTipText     =   "Número de autorización especial"
         Top             =   1780
         Width           =   2350
      End
      Begin VB.TextBox txtFolioRecetaHO 
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
         Left            =   -72840
         MaxLength       =   20
         TabIndex        =   16
         ToolTipText     =   "Folio de receta"
         Top             =   160
         Width           =   2895
      End
      Begin VB.TextBox txtICDHO 
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
         Left            =   -72840
         Locked          =   -1  'True
         TabIndex        =   15
         ToolTipText     =   "Código ICD"
         Top             =   570
         Width           =   1815
      End
      Begin VB.TextBox txtICDUR 
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
         Left            =   -72840
         Locked          =   -1  'True
         TabIndex        =   14
         ToolTipText     =   "Código ICD"
         Top             =   570
         Width           =   1815
      End
      Begin VB.TextBox txtDiagnosticoUR 
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
         Left            =   -72840
         Locked          =   -1  'True
         TabIndex        =   13
         ToolTipText     =   "Descripción del código ICD"
         Top             =   980
         Width           =   4335
      End
      Begin VB.TextBox txtICDUS 
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
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   12
         ToolTipText     =   "Código ICD"
         Top             =   570
         Width           =   1815
      End
      Begin VB.TextBox txtNumAutEspecialCP 
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
         Left            =   -70920
         MaxLength       =   3
         TabIndex        =   11
         ToolTipText     =   "Número de autorización especial"
         Top             =   980
         Width           =   2415
      End
      Begin VB.TextBox txtNumAutGralCP 
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
         Left            =   -70920
         MaxLength       =   7
         TabIndex        =   10
         ToolTipText     =   "Número de autorización general"
         Top             =   570
         Width           =   2415
      End
      Begin VB.TextBox txtFolioRecetaCP 
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
         Left            =   -70920
         MaxLength       =   20
         TabIndex        =   9
         ToolTipText     =   "Folio de receta"
         Top             =   160
         Width           =   2415
      End
      Begin VB.TextBox txtFolioRecetaUR 
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
         Left            =   -72840
         MaxLength       =   20
         TabIndex        =   8
         ToolTipText     =   "Folio de receta"
         Top             =   160
         Width           =   2895
      End
      Begin VB.TextBox txtFolioRecetaUS 
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
         Left            =   2160
         MaxLength       =   20
         TabIndex        =   7
         ToolTipText     =   "Folio de receta"
         Top             =   160
         Width           =   2895
      End
      Begin VB.TextBox txtDiagnosticoUS 
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
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   6
         ToolTipText     =   "Descripción del código ICD"
         Top             =   980
         Width           =   4335
      End
      Begin VB.TextBox txtSintomatologiaUS 
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
         Height          =   975
         Left            =   360
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   5
         ToolTipText     =   "Sintomatología del paciente"
         Top             =   1800
         Width           =   6135
      End
      Begin MyCommandButton.MyButton cmdBuscarICDHO 
         Height          =   375
         Left            =   -69840
         TabIndex        =   20
         ToolTipText     =   "Buscar código ICD"
         Top             =   160
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
         Caption         =   "Buscar ICD"
         DepthEvent      =   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdBuscarICDUR 
         Height          =   375
         Left            =   -69840
         TabIndex        =   21
         ToolTipText     =   "Buscar código ICD"
         Top             =   160
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
         Caption         =   "Buscar ICD"
         DepthEvent      =   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdBuscarICDUS 
         Height          =   375
         Left            =   5160
         TabIndex        =   22
         ToolTipText     =   "Buscar código ICD"
         Top             =   160
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
         Caption         =   "Buscar ICD"
         DepthEvent      =   1
         ShowFocus       =   -1  'True
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Descripción "
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
         Left            =   -74670
         TabIndex        =   37
         Top             =   1040
         Width           =   1185
      End
      Begin VB.Label Label12 
         BackColor       =   &H80000005&
         Caption         =   "Número de autorización general"
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
         Left            =   -74670
         TabIndex        =   36
         Top             =   1440
         Width           =   3735
      End
      Begin VB.Label Label11 
         BackColor       =   &H80000005&
         Caption         =   "Número de autorización especial"
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
         Left            =   -74670
         TabIndex        =   35
         Top             =   1840
         Width           =   3735
      End
      Begin VB.Label Label10 
         BackColor       =   &H80000005&
         Caption         =   "Folio de receta"
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
         Left            =   -74670
         TabIndex        =   34
         Top             =   220
         Width           =   1695
      End
      Begin VB.Label Label9 
         BackColor       =   &H80000005&
         Caption         =   "Código ICD"
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
         Left            =   -74670
         TabIndex        =   33
         Top             =   630
         Width           =   1335
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000005&
         Caption         =   "Código ICD"
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
         Left            =   -74670
         TabIndex        =   32
         Top             =   630
         Width           =   1335
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Descripción "
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
         Left            =   -74670
         TabIndex        =   31
         Top             =   1040
         Width           =   1185
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000005&
         Caption         =   "Código ICD"
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
         Left            =   330
         TabIndex        =   30
         Top             =   630
         Width           =   1425
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000005&
         Caption         =   "Número de autorización especial"
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
         Left            =   -74670
         TabIndex        =   29
         Top             =   1040
         Width           =   3975
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000005&
         Caption         =   "Número de autorización general"
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
         Left            =   -74670
         TabIndex        =   28
         Top             =   630
         Width           =   3945
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000005&
         Caption         =   "Folio de receta"
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
         Left            =   -74670
         TabIndex        =   27
         Top             =   220
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000005&
         Caption         =   "Folio de receta"
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
         Left            =   -74670
         TabIndex        =   26
         Top             =   220
         Width           =   1935
      End
      Begin VB.Label lblFolioReceta 
         BackColor       =   &H80000005&
         Caption         =   "Folio de receta"
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
         Left            =   330
         TabIndex        =   25
         Top             =   220
         Width           =   1845
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Descripción "
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
         Left            =   330
         TabIndex        =   24
         Top             =   1040
         Width           =   1185
      End
      Begin VB.Label lblSintomatologiaUS 
         BackColor       =   &H80000005&
         Caption         =   "Sintomatología"
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
         Left            =   360
         TabIndex        =   23
         Top             =   1440
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmDatosWSAXA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public vgstrCodigoICD As String 'Se declaró como variables públicas para la selección de código ICD en pantalla
Public vgstrDescDiagnostico As String 'Se declaró como variables públicas para la selección de código ICD en pantalla
Public vgblnConexionCorrecta As Boolean 'Esta variable se activa cuando se recibe una conexión correcta, la cual es utilizada en la pantalla de ingreso de paciente para continuar con el ingreso
Public vglngFolioTrans As Long  'Variable que almacena el folio de la transacción para futuras modificaciones
Public vgblnTrasladoCargos As Boolean 'Indica si los datos provienen desde la forma de Traslado de cargos

Dim vglngNumCuenta As Long
Dim vglngNumPaciente As Long
Dim lngCveTipoIngreso As Long
Dim vgstrIP As String
Dim vgstrEquipo As String
Dim vlstrCodigoICD As String                         'Clave del diagnóstico seleccionado
Dim vlstrFolioReceta As String
Dim vlstrSintomatologia As String
Dim blnPrimeraVez As Boolean        'Variable que indica si la pantalla se está mostrando por primera vez

'Variables que tomarán el valor de la pantalla de admisión de pacientes a partir de las variables públicas.
Dim vgstrDatosProveedorAXA As String            'Variable que indica la clave del proveedor para AXA, configurado en el catálogo de equivalencias
Dim vgstrDatosContratoAXA As String            'Variable que indica la clave del contrato para AXA, configurado en el catálogo de equivalencias
Dim vgstrDatosControlAXA As String         'Variable que indica la clave del contrato para AXA
Dim vgstrDatosNumCuartoAXA As String         'Variable que indica el número de cuarto para la interfaz AXA
Dim vgstrDatosAutorizaGralAXA As String         'Variable que indica el número de autorización general para la interfaz AXA
Dim vgstrDatosAutorizaEspecialAXA As String         'Variable que indica el número de autorización especial para la interfaz AXA
Dim vgstrDatosMedicoTratanteAXA As String         'Variable que indica el nombre del médico tratante para la interfaz AXA (INTERNOS)
Dim vgstrDatosMedicoEmergenciasAXA As String         'Variable que indica el nombre del médico para emergencias para la interfaz AXA (URGENCIAS)
Dim vglngDatosCveTipoIngreso As Long                    'Variable que indica el tipo de ingreso del paciente, para saber que datos solicitar
Dim vllngDatosPersonaGraba As Long
Dim vgstrDatosTipoPaciente As String
Dim strDestino As String 'Ruta para el destino del archivo request del WS

Private Function fnDOMRequestXML(intTipoEvento As Integer, blnEnsobretadoSOAP As Boolean) As MSXML2.DOMDocument
    Dim DOMRequestXML As MSXML2.DOMDocument
    Set DOMRequestXML = New MSXML2.DOMDocument
        
    Dim strNombreArchivoRequest As String

    'Se forma el nombre del archivorequest (número de control axa + contrato + tipo evento + num. paciente + num cuenta + folio receta)
    strNombreArchivoRequest = Trim(vgstrDatosControlAXA) & Trim(vgstrDatosContratoAXA) & CStr(intTipoEvento) & CStr(vglngNumPaciente) & CStr(vglngNumCuenta) & Trim(vlstrFolioReceta)
        
    If blnEnsobretadoSOAP = True Then
        '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Dim DOMelementoSOAP As MSXML2.IXMLDOMElement
        Set DOMelementoSOAP = DOMRequestXML.createElement("soap:Envelope")
    
        ' Aquí van los Namespaces
        DOMelementoSOAP.setAttribute "xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance"
        DOMelementoSOAP.setAttribute "xmlns:xsd", "http://www.w3.org/2001/XMLSchema"
        DOMelementoSOAP.setAttribute "xmlns:soap", "http://schemas.xmlsoap.org/soap/envelope/"
        
        DOMRequestXML.appendChild DOMelementoSOAP
        '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Dim DOMelementoSOAPBody As MSXML2.IXMLDOMElement
        Set DOMelementoSOAPBody = DOMRequestXML.createElement("soap:Body")
        
        DOMelementoSOAP.appendChild DOMelementoSOAPBody
        '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    End If
        
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim DOMelementoRaiz As MSXML2.IXMLDOMElement
    
    If intTipoEvento = 4 Then 'Hospitalización
        Set DOMelementoRaiz = DOMRequestXML.createElement("HospitalizacionUrgenciaSentida")
    Else
        Set DOMelementoRaiz = DOMRequestXML.createElement("ObtenerAutorizacion")
    End If
    
    DOMelementoRaiz.setAttribute "xmlns", "http://www.axa-assistance-la.com/"
    
    If blnEnsobretadoSOAP = True Then
        DOMelementoSOAPBody.appendChild DOMelementoRaiz
    Else
        DOMRequestXML.appendChild DOMelementoRaiz
    End If
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    Dim DOMelementoAutoriza As MSXML2.IXMLDOMElement
    Dim DOMelementoHospitaliza As MSXML2.IXMLDOMElement
    Dim DOMelementoTipoUrgencia As MSXML2.IXMLDOMElement
    
    If intTipoEvento = 4 Then 'Hospitalización
    
        Set DOMelementoHospitaliza = DOMRequestXML.createElement("hospitaliza")
                    
        '================================================ ATRIBUTOS ===========================================
            
        DOMelementoHospitaliza.setAttribute "claveProveedor", vgstrDatosProveedorAXA
        DOMelementoHospitaliza.setAttribute "contrato", vgstrDatosContratoAXA
        DOMelementoHospitaliza.setAttribute "numeroNomina", vgstrDatosControlAXA
        DOMelementoHospitaliza.setAttribute "autorizacionGeneral", vgstrDatosAutorizaGralAXA
        DOMelementoHospitaliza.setAttribute "autorizacionEspecial", vgstrDatosAutorizaEspecialAXA
        DOMelementoHospitaliza.setAttribute "motivoIngreso", vlstrCodigoICD
        DOMelementoHospitaliza.setAttribute "noHabitacion", vgstrDatosNumCuartoAXA
        DOMelementoHospitaliza.setAttribute "medicoTratante", vgstrDatosMedicoTratanteAXA
        DOMelementoHospitaliza.setAttribute "xmlns", "http://www.axa-assistance-la.com/wsre/esquemas"
           
        DOMelementoRaiz.appendChild DOMelementoHospitaliza
        
        '===================================================================================================
        
    Else
    
        Set DOMelementoAutoriza = DOMRequestXML.createElement("autoriza")
            
        DOMelementoAutoriza.setAttribute "claveProveedor", vgstrDatosProveedorAXA '"ERCJD"
        DOMelementoAutoriza.setAttribute "contrato", vgstrDatosContratoAXA '"2740002"
        DOMelementoAutoriza.setAttribute "numeroNomina", vgstrDatosControlAXA '"849523"
        DOMelementoAutoriza.setAttribute "tipoEvento", intTipoEvento '"2"
        DOMelementoAutoriza.setAttribute "xmlns", "http://www.axa-assistance-la.com/wsre/esquemas"
        
        DOMelementoRaiz.appendChild DOMelementoAutoriza
        '============================================== TIPOS DE EVENTO =========================================
        If intTipoEvento = 1 Then 'Urgencia sentida
            Set DOMelementoTipoUrgencia = DOMRequestXML.createElement("urgenciaSentida")
            DOMelementoTipoUrgencia.setAttribute "servicio", "99281" 'Se manda por default el código de servicio de consulta
            DOMelementoTipoUrgencia.setAttribute "sintomatologia", Trim(txtSintomatologiaUS.Text)
            DOMelementoTipoUrgencia.setAttribute "motivoIngreso", vlstrCodigoICD
            DOMelementoTipoUrgencia.setAttribute "registroMedicoAtiende", vgstrDatosMedicoEmergenciasAXA
            DOMelementoTipoUrgencia.setAttribute "fechaHoraIngreso", Format(fdtmServerFechaHora, "YYYY-MM-DDTHH:MM:SS") 'Formateado en formato ISO (como lo requiere el WS)
            
            DOMelementoAutoriza.appendChild DOMelementoTipoUrgencia
        
        ElseIf intTipoEvento = 2 Then 'Urgencia real
            Set DOMelementoTipoUrgencia = DOMRequestXML.createElement("urgenciaReal")
            DOMelementoTipoUrgencia.setAttribute "motivoIngreso", vlstrCodigoICD
            
            DOMelementoAutoriza.appendChild DOMelementoTipoUrgencia
        
        ElseIf intTipoEvento = 3 Then 'Cirugía programada
            Set DOMelementoTipoUrgencia = DOMRequestXML.createElement("cirugiaProgramada")
            DOMelementoTipoUrgencia.setAttribute "autorizacionGeneral", vgstrDatosAutorizaGralAXA
            DOMelementoTipoUrgencia.setAttribute "autorizacionEspecial", vgstrDatosAutorizaEspecialAXA
            
            DOMelementoAutoriza.appendChild DOMelementoTipoUrgencia
        End If
    '===================================================================================================
    
    End If
        
    'Se regresa el valor de la función
    Set fnDOMRequestXML = DOMRequestXML

    'Se agrega un número ID para el archivo request para evitar duplicidades
    strNombreArchivoRequest = strNombreArchivoRequest & MakeTempFileName("xml")

    'Se graba el archivo Request Timbrado en la ruta especificada
    fnDOMRequestXML.Save strDestino + "\" + strNombreArchivoRequest
End Function


Private Sub pLimpia(Optional vlblnLimpiaSalida As Boolean)
On Error GoTo NotificaError
    
'Se reestablecen todos los campos y variables de la pantalla
txtFolioRecetaUS.Text = ""
txtFolioRecetaUR.Text = ""
txtICDUS.Text = ""
txtICDUR.Text = ""
txtICDHO.Text = ""
txtDiagnosticoUS = ""
txtDiagnosticoUR = ""
txtDiagnosticoHO = ""
txtSintomatologiaUS = ""
txtNumAutEspecialCP = ""
txtNumAutGralCP = ""
txtNumAutEspecialHO = ""
txtNumAutGralHO = ""
txtFolioRecetaCP = ""
txtFolioRecetaHO.Text = ""
vgblnConexionCorrecta = False

'Variables y botón grabar
vlstrCodigoICD = ""
vgstrDescDiagnostico = ""
vgstrCodigoICD = ""
cmdGrabar.Enabled = False

'Se limpian estas variables si se realizó una conexión exitosa
If vlblnLimpiaSalida Then
    vgstrDatosProveedorAXA = ""
    vgstrDatosContratoAXA = ""
    vgstrDatosControlAXA = ""
    vgstrDatosNumCuartoAXA = ""
    vgstrDatosAutorizaGralAXA = ""
    vgstrDatosAutorizaEspecialAXA = ""
    vgstrDatosMedicoTratanteAXA = ""
    vgstrDatosMedicoEmergenciasAXA = ""
    vglngDatosCveTipoIngreso = 0
    vglngNumCuenta = 0
    vglngNumPaciente = 0
    vgstrIP = ""
    vgstrEquipo = ""
    vlstrSintomatologia = ""
End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLimpia"))
End Sub
Private Function fblnValidaDatos() As Boolean
On Error GoTo NotificaError
    
'Se inicializa el resultado de la función
fblnValidaDatos = False
    
'Se valida la información capturada según la pestaña activa
Select Case Me.sstObj.Tab
    Case 0      'Urgencia sentida
        If Not txtFolioRecetaUS = "" And Not txtDiagnosticoUS = "" Then
            fblnValidaDatos = True
        End If
    Case 1      'Urgencia real
        If Not txtFolioRecetaUR = "" And Not txtDiagnosticoUR = "" Then
            fblnValidaDatos = True
        End If
    Case 2      'Cirugía programada
        If Not txtFolioRecetaCP = "" Then
            fblnValidaDatos = True
        End If
    Case 3      'Hospitalización (INTERNO FUE URGENCIAS)
        If Not txtFolioRecetaHO = "" And Not txtNumAutGralHO = "" Then
            fblnValidaDatos = True
        End If
End Select

If fblnValidaDatos Then
    cmdGrabar.Enabled = True
End If

Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblValidaDatos"))
End Function
Private Sub grdHBusqueda_Click()

End Sub

Private Sub txtIniciales_Change()

End Sub

Private Sub cboTipoUrgencia_Click()

    If cmdGrabar.Enabled And cboTipoUrgencia.ListIndex <> sstObj.Tab Then
        '¿Desea abandonar la operación?
        If MsgBox(SIHOMsg(17), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
            pLimpia
        Else
            'Se reestablece el valor el index del combo box
            Select Case sstObj.Tab
                Case 0
                cboTipoUrgencia.ListIndex = 0
                Case 1
                cboTipoUrgencia.ListIndex = 1
                Case 2
                cboTipoUrgencia.ListIndex = 2
            End Select
            
            Exit Sub
        End If
    ElseIf cmdGrabar.Enabled And cboTipoUrgencia.ListIndex = sstObj.Tab Then
        Exit Sub
    Else
        pLimpia
    End If

    If Not blnPrimeraVez Then
        Select Case cboTipoUrgencia.ListIndex
            Case 0
                sstObj.Tab = 0
                'Se ajusta el tamaño de la forma
                cmdGrabar.Top = 3960
                frmDatosWSAXA.Height = 5070
            Case 1
                sstObj.Tab = 1
                'Se ajusta el tamaño de la forma
                cmdGrabar.Top = 2450
                frmDatosWSAXA.Height = 3600
            Case 2
                sstObj.Tab = 2
                'Se ajusta el tamaño de la forma
                cmdGrabar.Top = 2520
                frmDatosWSAXA.Height = 3630
        End Select
    End If
    
End Sub


Private Sub cboTipoUrgencia_KeyDown(KeyCode As Integer, Shift As Integer)

  If KeyCode = 13 Then
    Select Case cboTipoUrgencia.ListIndex
        Case 0
            sstObj.Tab = 0
            txtFolioRecetaUS.SetFocus
        Case 1
            sstObj.Tab = 1
            txtFolioRecetaUR.SetFocus
        Case 2
            sstObj.Tab = 2
            txtFolioRecetaCP.SetFocus
    End Select
  End If

End Sub


Private Sub cmdBuscarICDHO_Click()
    On Error GoTo NotificaError
    
    Set frmBusquedaICD.vlfrmForma = frmDatosWSAXA
    
    frmBusquedaICD.vgstrTipoICD = "US"
    
    frmBusquedaICD.Show vbModal, Me

    If vgstrCodigoICD <> "" Then
        txtDiagnosticoHO.Text = vgstrDescDiagnostico
        vlstrCodigoICD = vgstrCodigoICD
        txtICDHO.Text = vgstrCodigoICD
    Else
       txtDiagnosticoHO.Text = ""
       vlstrCodigoICD = ""
    End If
    
    fblnValidaDatos
    
    txtNumAutGralHO.SetFocus


Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdBuscarICDHO_Click"))
    Unload Me

End Sub

Private Sub cmdBuscarICDUR_Click()

    Set frmBusquedaICD.vlfrmForma = frmDatosWSAXA
    
    frmBusquedaICD.vgstrTipoICD = "UR"
    
    frmBusquedaICD.Show vbModal, Me

    If vgstrCodigoICD <> "" Then
        txtDiagnosticoUR.Text = vgstrDescDiagnostico
        vlstrCodigoICD = vgstrCodigoICD
        txtICDUR.Text = vgstrCodigoICD
    Else
       txtDiagnosticoUR.Text = ""
       vlstrCodigoICD = ""
    End If
    
    fblnValidaDatos
    
    If cmdGrabar.Enabled = True Then
        cmdGrabar.SetFocus
    End If
        
End Sub

Private Sub cmdBuscarICDUS_Click()
    On Error GoTo NotificaError
    
    Set frmBusquedaICD.vlfrmForma = frmDatosWSAXA
    
    frmBusquedaICD.vgstrTipoICD = "US"
    
    frmBusquedaICD.Show vbModal, Me

    If vgstrCodigoICD <> "" Then
        txtDiagnosticoUS.Text = vgstrDescDiagnostico
        vlstrCodigoICD = vgstrCodigoICD
        txtICDUS.Text = vgstrCodigoICD
    Else
       txtDiagnosticoUS.Text = ""
       vlstrCodigoICD = ""
    End If
    
    fblnValidaDatos
    
    txtSintomatologiaUS.SetFocus


Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdBuscarICDUS_Click"))
    Unload Me
End Sub

Private Sub cmdGrabar_Click()

    Dim DOMRequestXML As MSXML2.DOMDocument
    Dim SerializerWS As SoapSerializer30 'Para serializar el XML
    Dim ReaderRespuestaWS As SoapReader30      'Para leer la respuesta del WebService
    Dim ConectorWS As ISoapConnector 'Para conectarse al WebService
    Dim rsConexion As New ADODB.Recordset
    Dim vlaryParametrosSalida() As String
    
    'Se inicializa la variable del folio
    vglngFolioTrans = 0
    Err.Clear
    
    If Not fblnValidaDatos Then
        MsgBox SIHOMsg(530), vbOKOnly + vbExclamation, "Mensaje"
        Exit Sub
    Else
    
        'Se especifican las rutas para la conexión con el servicio de timbrado
        Set rsConexion = frsEjecuta_SP(CStr(vgintClaveEmpresaContable), "SP_GNSELCONFIGINTERFAZWS")
    
        'Se elimina el contenido de la carpeta temporal
        strDestino = Environ$("temp") & "\fm-Axa27"
        pCreaDirectorio strDestino
        On Error Resume Next
        
        'Checa que el directorio se haya creado, para proceder a eliminar el archivo (evita el error 53)
        If Dir$(strDestino & "\") <> "" Then
            Kill strDestino & "\*.*"
            If Err.Number = 53 Or Err.Number = 0 Then '|  File not found
                Err.Clear
            Else
                Err.Raise Err.Number
            End If
        End If
        
        'Se asignan los valores de la pantalla de captura a las variables
        vgstrDatosProveedorAXA = Trim(rsConexion!CVEPROVEEDOR)
        
        Select Case Me.sstObj.Tab
            Case 0      'Urgencia sentida
                vlstrFolioReceta = Trim(txtFolioRecetaUS.Text)
                vlstrCodigoICD = Trim(txtICDUS.Text)
                vlstrSintomatologia = Trim(txtSintomatologiaUS.Text)
            Case 1      'Urgencia real
                vlstrFolioReceta = Trim(txtFolioRecetaUR.Text)
                vlstrCodigoICD = Trim(txtICDUR.Text)
            Case 2      'Cirugía programada
                vlstrFolioReceta = Trim(txtFolioRecetaCP.Text)
                vgstrDatosAutorizaGralAXA = Trim(txtNumAutGralCP.Text)
                vgstrDatosAutorizaEspecialAXA = Trim(txtNumAutEspecialCP.Text)
            Case 3      'Hospitalización (INTERNO FUE URGENCIAS)
                vlstrFolioReceta = Trim(txtFolioRecetaHO.Text)
                vlstrCodigoICD = Trim(txtICDHO.Text)
                vgstrDatosAutorizaGralAXA = Trim(txtNumAutGralHO.Text)
                vgstrDatosAutorizaEspecialAXA = Trim(txtNumAutEspecialHO.Text)
        End Select
    
        Set ConectorWS = New HttpConnector30
        ' La URL que atenderá nuestra solicitud
        ConectorWS.Property("EndPointURL") = Trim(rsConexion!URLWSConexion) '"http://www.axa-assistance-la.com:8082/wsre/eRecetario.asmx?WSDL"
                    
        ' Ruta del WebMethod según el tipo de ingreso
        ConectorWS.Property("SoapAction") = IIf(sstObj.Tab = 3, Trim(rsConexion!URLWMHospitalizacion), Trim(rsConexion!URLWMAutorizacion))  '"http://www.axa-assistance-la.com/HospitalizacionUrgenciaSentida" o "http://www.axa-assistance-la.com/ObtenerAutorizacion"
    
        'Se forma el archivo Request
        Set DOMRequestXML = fnDOMRequestXML(sstObj.Tab + 1, False)
        
        '###########################################################################################################################
        '########################################## INICIA CONEXIÓN CON EL SERVICIO DE TIMBRADO ##########################################
        '###########################################################################################################################
        ConectorWS.Connect
    
        ConectorWS.BeginMessage
            Set SerializerWS = New SoapSerializer30
            SerializerWS.Init ConectorWS.InputStream
            
            SerializerWS.StartEnvelope
                SerializerWS.StartBody
                    SerializerWS.WriteXml DOMRequestXML.xml
                SerializerWS.EndBody
            SerializerWS.EndEnvelope
            
        ConectorWS.EndMessage
        
        Set ReaderRespuestaWS = New SoapReader30
        ReaderRespuestaWS.Load ConectorWS.OutputStream
        '#############################################################################################################################
        '########################################## FINAL DE CONEXIÓN CON EL SERVICIO DE TIMBRADO ##########################################
        '#############################################################################################################################
        
        If Not ReaderRespuestaWS.Fault Is Nothing Then
            Dim strMensajeError As String
            If Not ReaderRespuestaWS.FaultDetail Is Nothing Then
                Dim DOMNodoCodigoError As MSXML2.IXMLDOMNode
                
                'Se obtiene el codigo del error (en caso de haberlo)
                Set DOMNodoCodigoError = ReaderRespuestaWS.FaultDetail.childNodes(0).Attributes.getNamedItem("codigo")
                
                'Se captura el error
                If DOMNodoCodigoError Is Nothing Then
                    Err.Raise 1000, "Comunicación AXA", "Error"
                Else
                    strMensajeError = "Ocurrió un error al comunicarse con AXA" & vbNewLine & vbNewLine & _
                                                "Número de error: " & DOMNodoCodigoError.Text & vbNewLine & _
                                                "Descripción: " & ReaderRespuestaWS.FaultString.Text
                    'Se muestra el mensaje de error en pantalla
                    MsgBox strMensajeError, vbCritical + vbOKOnly, "Mensaje"
                End If
            End If
        Else
        
            'Se valida si tiene un error en la conexión al WS
            If Err.Number = 5400 Or Err.Number = -2147024809 Then GoTo NotificaError
        
            'Se determina si se regresó un mensaje de error al solicitar información con AXA
            Dim DOMElementoRespuestaWS As MSXML2.IXMLDOMElement
            Set DOMElementoRespuestaWS = ReaderRespuestaWS.Body
             
            If Trim(DOMElementoRespuestaWS.selectSingleNode("//@mensajeError").Text) <> "" Then 'Si se regresó un mensaje de error...
            
                'Se almacena en el log de transacciones de la interfaz de AXA
                pCargaArreglo vlaryParametrosSalida, "|" & adDouble 'adBSTR 'adVarChar
                vgstrParametrosSP = "|" & vgintNumeroModulo & "|" & vglngDatosCveTipoIngreso & "|" & vglngNumCuenta & "|" & vglngNumPaciente & "|" & vgstrDatosTipoPaciente & "|" & Trim(vgstrIP) & "|" & IIf(sstObj.Tab = 0, "US", IIf(sstObj.Tab = 1, "UR", IIf(sstObj.Tab = 2, "CP", "HO"))) & "|NO|" & CStr(DOMRequestXML.xml) & "|" & CStr(ReaderRespuestaWS.Body.xml) & "|" & Trim(vgstrEquipo) & "|" & vllngDatosPersonaGraba & "|0|" & Trim(vlstrFolioReceta)
                frsEjecuta_SP vgstrParametrosSP, "sp_GNINSloginterfazaxa", , , vlaryParametrosSalida
                pObtieneValores vlaryParametrosSalida, vglngFolioTrans
            
                'Generación el MsgBox con hipervínculo al chat de AXA
                Set objShell = CreateObject("Wscript.Shell")
                intMessage = MsgBox("Información incorrecta: " & vbNewLine & vbNewLine & "- " & Trim(DOMElementoRespuestaWS.selectSingleNode("//@mensajeError").Text) & vbNewLine & vbNewLine & "¿Desea abrir el chat en línea con AXA?", vbYesNo + vbExclamation, "Mensaje")
                
                'Si se selecciona que sí, se abre la ventana del chat en línea con AXA
                If intMessage = vbYes And Left(Trim(DOMElementoRespuestaWS.selectSingleNode("//@mensajeError").Text), 6) = "Object" Then
                    MsgBox "No se proporcionó una liga de acceso al chat en línea de AXA", vbInformation, "Mensaje"
                    Exit Sub
                End If
                
                If intMessage = vbYes And Left(Trim(DOMElementoRespuestaWS.selectSingleNode("//@mensajeError").Text), 3) = "The" Then
                    MsgBox "No se proporcionó una liga de acceso al chat en línea de AXA", vbInformation, "Mensaje"
                    Exit Sub
                End If
                
                If intMessage = vbYes And Left(Trim(DOMElementoRespuestaWS.selectSingleNode("//@mensajeError").Text), 10) = "Conversion" Then
                    MsgBox "No se proporcionó una liga de acceso al chat en línea de AXA", vbInformation, "Mensaje"
                    Exit Sub
                End If
                
                If intMessage = vbYes And Trim(DOMElementoRespuestaWS.selectSingleNode("//@linkAyuda").Text) <> "" Then
                    objShell.Run DOMElementoRespuestaWS.selectSingleNode("//@linkAyuda").Text
                ElseIf intMessage = vbYes And Trim(DOMElementoRespuestaWS.selectSingleNode("//@linkAyuda").Text) = "" Then
                    MsgBox "No se proporcionó una liga de acceso al chat en línea de AXA", vbInformation, "Mensaje"
                    Exit Sub
                Else
                    Exit Sub
                End If
                
            Else
                'Se selecciona el mensaje de conexión exitosa correspondiente
                Select Case Me.sstObj.Tab
                    Case 0      'Urgencia sentida
                        MsgBox "Conexión exitosa. " & vbNewLine & vbNewLine & "Autorización general: " & Trim(DOMElementoRespuestaWS.selectSingleNode("//@autorizacionGeneral").Text) & vbNewLine & "Autorización especial: " & Trim(DOMElementoRespuestaWS.selectSingleNode("//@autorizacionEspecial").Text), vbInformation, "Mensaje"
                    Case 1      'Urgencia real
                        MsgBox "Conexión exitosa. " & vbNewLine & vbNewLine & "Autorización general: " & Trim(DOMElementoRespuestaWS.selectSingleNode("//@autorizacionGeneral").Text) & vbNewLine & "Autorización especial: " & Trim(DOMElementoRespuestaWS.selectSingleNode("//@autorizacionEspecial").Text), vbInformation, "Mensaje"
                    Case 2      'Cirugía programada
                        MsgBox "Conexión exitosa. " & vbNewLine & vbNewLine & "Nombre del derechohabiente: " & Trim(DOMElementoRespuestaWS.selectSingleNode("//@nombreDerechohabiente").Text) & vbNewLine & "Número de nómina: " & Trim(DOMElementoRespuestaWS.selectSingleNode("//@numeroNomina").Text) & vbNewLine & "Fecha de la cirugía: " & Trim(DOMElementoRespuestaWS.selectSingleNode("//@fechaCirugia").Text) & vbNewLine & "Proveedor: " & Trim(DOMElementoRespuestaWS.selectSingleNode("//@proveedor").Text) & vbNewLine & "Servicio: " & Trim(DOMElementoRespuestaWS.selectSingleNode("//@servicio").Text) & vbNewLine & "Motivo de ingreso: " & Trim(DOMElementoRespuestaWS.selectSingleNode("//@motivoIngreso").Text), vbInformation, "Mensaje"
                    Case 3      'Hospitalización (INTERNO FUE URGENCIAS)
                        MsgBox "Conexión exitosa. " & vbNewLine & vbNewLine & "Autorización especial: " & Trim(DOMElementoRespuestaWS.selectSingleNode("//@autorizacionEspecial").Text), vbInformation, "Mensaje"
                End Select
                                
                'Se almacena en el log de transacciones de la interfaz de AXA
                pCargaArreglo vlaryParametrosSalida, "|" & adDouble 'adBSTR 'adVarChar
                vgstrParametrosSP = "|" & vgintNumeroModulo & "|" & vglngDatosCveTipoIngreso & "|" & vglngNumCuenta & "|" & vglngNumPaciente & "|" & vgstrDatosTipoPaciente & "|" & Trim(vgstrIP) & "|" & IIf(sstObj.Tab = 0, "US", IIf(sstObj.Tab = 1, "UR", IIf(sstObj.Tab = 2, "CP", "HO"))) & "|SI|" & CStr(DOMRequestXML.xml) & "|" & CStr(ReaderRespuestaWS.Body.xml) & "|" & Trim(vgstrEquipo) & "|" & vllngDatosPersonaGraba & "|0|" & Trim(vlstrFolioReceta)
                frsEjecuta_SP vgstrParametrosSP, "sp_GNINSloginterfazaxa", , , vlaryParametrosSalida
                pObtieneValores vlaryParametrosSalida, vglngFolioTrans
                
                pLimpia (True)
                Unload Me
                
                'Se regresa variable de estatus de conexión correcta
                vgblnConexionCorrecta = True
                
            End If
            
            Set DOMElementoRespuestaWS = Nothing
        End If
                
    End If
        
Exit Sub
NotificaError:
    strMensajeError = "Ocurrió un error al comunicarse con AXA" & vbNewLine & vbNewLine & _
                                "Verifique que el equipo cuente con acceso a Internet y/o que la información enviada a AXA esté correcta."
    MsgBox strMensajeError, vbCritical + vbOKOnly, "Mensaje"
    Err.Clear
    Exit Sub
    
End Sub
     Private Function MakeTempFileName(Extension As String) As String
      On Error Resume Next
      Dim Isfile As Integer, FHandle As Integer, Cntr As Integer
      Dim WinTemp As String, TF As String
         Isfile = False
         FHandle = FreeFile

      Do
         'WinTemp = Environ("TEMP") & "\"
         WinTemp = ""
         For Cntr = 1 To 8
         WinTemp = WinTemp & Mid(LTrim(Str(CInt(Rnd * 10))), 1, 1)
         Next
            TF = Trim(WinTemp$) & "." & Extension

         Open TF For Output As #FHandle
      Debug.Print TF
         Print #FHandle, "This is a Temp file"
      Loop While Err > 0
      Close #FHandle
      MakeTempFileName = TF

   End Function
Private Sub Form_Activate()

    'Si se abre por primera vez la pantalla para US para posicionarse sobre el combo
    If sstObj.Tab = 0 And txtSintomatologiaUS = "" And txtDiagnosticoUS = "" And txtFolioRecetaUS.Text = "" And cmdGrabar.Enabled = False Then
        cboTipoUrgencia.SetFocus
    ElseIf sstObj.Tab = 0 And txtFolioRecetaUS.Visible = True And txtDiagnosticoUS.Text = "" Then
        pSelTextBox txtFolioRecetaUS
        txtFolioRecetaUS.SetFocus
    ElseIf sstObj.Tab = 0 And txtFolioRecetaUS.Visible = True And txtDiagnosticoUS.Text <> "" Then
        pSelTextBox txtSintomatologiaUS
        txtSintomatologiaUS.SetFocus
    End If

    If sstObj.Tab = 1 And txtFolioRecetaUR.Visible = True And txtDiagnosticoUR.Text = "" Then
        pSelTextBox txtFolioRecetaUR
        txtFolioRecetaUR.SetFocus
    ElseIf sstObj.Tab = 1 And txtFolioRecetaUR.Visible = True And txtDiagnosticoUR.Text <> "" Then
        pSelTextBox txtSintomatologiaUS
        If cmdGrabar.Enabled = True Then
            cmdGrabar.SetFocus
        Else
            pSelTextBox txtFolioRecetaUR
            txtFolioRecetaUR.SetFocus
        End If
    End If

    If sstObj.Tab = 2 And txtFolioRecetaCP.Visible = True And cmdGrabar.Enabled = False Then
        pSelTextBox txtFolioRecetaCP
        txtFolioRecetaCP.SetFocus
    End If

    If sstObj.Tab = 3 And txtFolioRecetaHO.Visible = True And txtDiagnosticoHO.Text = "" Then
        pSelTextBox txtFolioRecetaHO
        txtFolioRecetaHO.SetFocus
    ElseIf sstObj.Tab = 3 And txtFolioRecetaHO.Visible = True And txtDiagnosticoHO.Text <> "" Then
        pSelTextBox txtNumAutGralHO
        txtNumAutGralHO.SetFocus
    End If
        
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 27 Then
        If cmdGrabar.Enabled Then
            '¿Desea abandonar la operación?
            If MsgBox(SIHOMsg(17), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
                pLimpia
                
                'Se manda el foco
                If sstObj.Tab = 3 Then
                    txtFolioRecetaHO.SetFocus
                ElseIf sstObj.Tab = 2 Then
                    txtFolioRecetaCP.SetFocus
                Else
                    cboTipoUrgencia.SetFocus
                End If

            End If
        Else
            Unload Me
        End If
    End If
    
Exit Sub

NotificaError:
Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyPress"))

End Sub

Private Sub Form_Load()
    
    SetStyle sstObj.hwnd, 0
    SetSolidColor sstObj.hwnd, 16777215
    SSTabSubclass sstObj.hwnd
    
    Me.Icon = frmMenuPrincipal.Icon
    blnPrimeraVez = True
    cboTipoUrgencia.ListIndex = 0
    blnPrimeraVez = False
    
    If vgblnTrasladoCargos Then
    'Se asignan los valores de las variables según la información obtenida en la pantalla del traslado de cargos
        vgstrDatosContratoAXA = Trim(frmTrasladoCargos.vgstrContratoAXA)
        vgstrDatosControlAXA = Trim(frmTrasladoCargos.vgstrControlAXA)
        vgstrDatosNumCuartoAXA = Trim(frmTrasladoCargos.vgstrNumCuartoAXA)
        vgstrDatosAutorizaGralAXA = Trim(frmTrasladoCargos.vgstrAutorizaGralAXA)
        vgstrDatosAutorizaEspecialAXA = Trim(frmTrasladoCargos.vgstrAutorizaEspecialAXA)
        vgstrDatosMedicoTratanteAXA = Trim(frmTrasladoCargos.vgstrMedicoTratanteAXA)
        vgstrDatosMedicoEmergenciasAXA = Trim(frmTrasladoCargos.vgstrMedicoEmergenciasAXA)
        vglngDatosCveTipoIngreso = frmTrasladoCargos.vglngCveTipoIngresoAXA
        vllngDatosPersonaGraba = frmTrasladoCargos.vglngPersonaGrabaAXA
    Else
    'Se asignan los valores de las variables según la información obtenida en la pantalla de la admisión del paciente
'    vgstrDatosProveedorAXA = Trim(frmAdmisionPaciente.vgstrProveedorAXA)
        vgstrDatosContratoAXA = Trim(frmAdmisionPaciente.vgstrContratoAXA)
        vgstrDatosControlAXA = Trim(frmAdmisionPaciente.vgstrControlAXA)
        vgstrDatosNumCuartoAXA = Trim(frmAdmisionPaciente.vgstrNumCuartoAXA)
        vgstrDatosAutorizaGralAXA = Trim(frmAdmisionPaciente.vgstrAutorizaGralAXA)
        vgstrDatosAutorizaEspecialAXA = Trim(frmAdmisionPaciente.vgstrAutorizaEspecialAXA)
        vgstrDatosMedicoTratanteAXA = Trim(frmAdmisionPaciente.vgstrMedicoTratanteAXA)
        vgstrDatosMedicoEmergenciasAXA = Trim(frmAdmisionPaciente.vgstrMedicoEmergenciasAXA)
        vglngDatosCveTipoIngreso = frmAdmisionPaciente.vgCveTipoIngresoAXA
        vllngDatosPersonaGraba = frmAdmisionPaciente.vllngPersonaGrabaAXA
    End If
    
    vgblnConexionCorrecta = False
    vgstrDatosTipoPaciente = "E"
    
    'Se obtiene la IP de la máquina huesped
    Call ObtenerPCIP
    vgstrIP = vgstrNumeroIP
    vgstrEquipo = vgstrNombreMaquina
       
    'Se valida el tipo de ingreso, para saber que tipos de datos se deberán capturar.
    If vglngDatosCveTipoIngreso = 11 Or vglngDatosCveTipoIngreso = 1 Then 'PREVIO o INTERNAMIENTO NORMAL
        vgstrDatosTipoPaciente = "I"
        'Se ajusta la forma
        frmDatosWSAXA.Height = 5085
        sstObj.Top = 900
        cmdGrabar.Top = 3960
        cboTipoUrgencia.ListIndex = 2
        cboTipoUrgencia.Enabled = False
    ElseIf vglngDatosCveTipoIngreso = 4 Then 'INTERNO FUÉ URGENCIAS
        vgstrDatosTipoPaciente = "I"
        'Se ajusta la forma
        frmDatosWSAXA.Height = 4320
        sstObj.Top = 0
        cmdGrabar.Top = 3120
        cboTipoUrgencia.Enabled = False
        sstObj.Tab = 3
    End If
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    vgblnTrasladoCargos = False
End Sub

Private Sub txtDiagnosticoHO_Change()
fblnValidaDatos

End Sub

Private Sub txtDiagnosticoHO_GotFocus()
pSelTextBox txtDiagnosticoHO
End Sub


Private Sub txtDiagnosticoHO_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))

End Sub

Private Sub txtDiagnosticoUR_Change()
fblnValidaDatos
End Sub

Private Sub txtDiagnosticoUR_GotFocus()
    pSelTextBox txtDiagnosticoUR
End Sub

Private Sub txtDiagnosticoUR_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txtDiagnosticoUS_Change()
fblnValidaDatos
End Sub

Private Sub txtDiagnosticoUS_GotFocus()
    pSelTextBox txtDiagnosticoUS
End Sub

Private Sub txtDiagnosticoUS_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtDiagnosticoUS_KeyDown"))
End Sub

Private Sub txtDiagnosticoUS_KeyPress(KeyAscii As Integer)

On Error GoTo NotificaError
    
    KeyAscii = Asc(UCase(Chr(KeyAscii)))

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtIniciales_KeyPress"))

End Sub

Private Sub txtFolioRecetaCP_Change()
    fblnValidaDatos
End Sub

Private Sub txtFolioRecetaCP_GotFocus()
    pSelTextBox txtFolioRecetaCP
End Sub

Private Sub txtFolioRecetaCP_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If

End Sub


Private Sub txtFolioRecetaHO_Change()
    fblnValidaDatos
End Sub

Private Sub txtFolioRecetaHO_GotFocus()
    pSelTextBox txtFolioRecetaHO
End Sub

Private Sub txtFolioRecetaHO_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If Trim(txtDiagnosticoHO.Text) = "" Then
        If KeyAscii = 13 Then
            SendKeys vbTab
        End If
    Else
        If KeyAscii = 13 Then
            txtNumAutGralHO.SetFocus
        End If
    End If

End Sub

Private Sub txtFolioRecetaUR_Change()
    fblnValidaDatos
End Sub

Private Sub txtFolioRecetaUR_GotFocus()
    pSelTextBox txtFolioRecetaUR
End Sub

Private Sub txtFolioRecetaUR_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If Trim(txtDiagnosticoUR.Text) = "" Then
        If KeyAscii = 13 Then
            SendKeys vbTab
        End If
    Else
        If KeyAscii = 13 Then
            If cmdGrabar.Enabled = True Then
                cmdGrabar.SetFocus
            Else
                txtFolioRecetaUR.SetFocus
            End If
        End If
    End If
    
End Sub


Private Sub txtFolioRecetaUS_Change()
    fblnValidaDatos
End Sub

Private Sub txtFolioRecetaUS_GotFocus()
    pSelTextBox txtFolioRecetaUS
End Sub

Private Sub txtFolioRecetaUS_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If Trim(txtDiagnosticoUS.Text) = "" Then
        If KeyAscii = 13 Then
            SendKeys vbTab
        End If
    Else
        If KeyAscii = 13 Then
            txtSintomatologiaUS.SetFocus
        End If
    End If

End Sub


Private Sub txtICDHO_Change()
fblnValidaDatos

End Sub

Private Sub txtICDHO_GotFocus()
pSelTextBox txtICDHO
End Sub


Private Sub txtICDHO_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))

End Sub

Private Sub txtICDUR_Change()
fblnValidaDatos
End Sub

Private Sub txtICDUR_GotFocus()
    pSelTextBox txtICDUR
End Sub

Private Sub txtICDUR_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txtICDUS_Change()
fblnValidaDatos
End Sub

Private Sub txtICDUS_GotFocus()
    pSelTextBox txtICDUS
End Sub


Private Sub txtICDUS_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txtNumAutEspecialCP_Change()
fblnValidaDatos
End Sub

Private Sub txtNumAutEspecialCP_GotFocus()
    pSelTextBox txtNumAutEspecialCP

End Sub

Private Sub txtNumAutEspecialCP_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If KeyAscii = 13 Then
        If cmdGrabar.Enabled = True Then
            cmdGrabar.SetFocus
        End If
    End If
    
End Sub


Private Sub txtNumAutEspecialHO_Change()
fblnValidaDatos

End Sub

Private Sub txtNumAutEspecialHO_GotFocus()
pSelTextBox txtNumAutEspecialHO
End Sub


Private Sub txtNumAutEspecialHO_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If KeyAscii = 13 Then
        If cmdGrabar.Enabled = True Then
            cmdGrabar.SetFocus
        End If
    End If

End Sub


Private Sub txtNumAutGralCP_Change()
fblnValidaDatos
End Sub

Private Sub txtNumAutGralCP_GotFocus()
    pSelTextBox txtNumAutGralCP
End Sub

Private Sub txtNumAutGralCP_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
    
End Sub


Private Sub txtNumAutGralHO_Change()
fblnValidaDatos

End Sub

Private Sub txtNumAutGralHO_GotFocus()
pSelTextBox txtNumAutGralHO
End Sub


Private Sub txtNumAutGralHO_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If

End Sub

Private Sub txtSintomatologiaUS_Change()
fblnValidaDatos
End Sub

Private Sub txtSintomatologiaUS_GotFocus()
    pSelTextBox txtSintomatologiaUS
End Sub

Private Sub txtSintomatologiaUS_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        If cmdGrabar.Enabled = True Then
            cmdGrabar.SetFocus
        End If
    End If

End Sub

Private Sub txtSintomatologiaUS_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub




