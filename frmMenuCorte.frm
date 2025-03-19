VERSION 5.00
Object = "{6B45E0EA-D03D-4CBB-94F4-B6AD155551A1}#1.1#0"; "MyFramePanel.ocx"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.1#0"; "MyCommandButton.ocx"
Begin VB.Form frmMenuCorte 
   BackColor       =   &H00F7F3EC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menú de cortes de caja"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4200
   Icon            =   "frmMenuCorte.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   4200
   StartUpPosition =   2  'CenterScreen
   Begin MyFramePanel.MyFrame Frame1 
      Height          =   2490
      Left            =   120
      Top             =   120
      Width           =   3945
      _ExtentX        =   6959
      _ExtentY        =   4392
      BackColor       =   16777215
      ForeColor       =   11682635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   2
      BackgroundAlignment=   4
      BorderColor     =   8677689
      Caption         =   ""
      CaptionAlignment=   4
      CornerRadius    =   25
      CornerTopLeft   =   -1  'True
      CornerTopRight  =   -1  'True
      CornerBottomLeft=   -1  'True
      CornerBottomRight=   -1  'True
      HeaderHeight    =   32
      HeaderColorTopLeft=   13284230
      HeaderColorTopRight=   13284230
      HeaderColorBottomLeft=   16777215
      HeaderColorBottomRight=   16777215
      Begin MyCommandButton.MyButton cmdReImpR 
         Height          =   315
         Left            =   195
         TabIndex        =   4
         ToolTipText     =   "Consulta de registros de terminales"
         Top             =   2040
         Width           =   3545
         _ExtentX        =   6244
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   16777215
         AppearanceThemes=   1
         BackColorDown   =   15133676
         TransparentColor=   16249836
         Caption         =   "Registro de terminal"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdTransferencias 
         Height          =   315
         Left            =   195
         TabIndex        =   2
         ToolTipText     =   "Asignación de depósitos en tránsito a cuentas de banco"
         Top             =   1170
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   16777215
         AppearanceThemes=   1
         BackColorDown   =   15133676
         TransparentColor=   16249836
         Caption         =   "Transferencias"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdConsultaPolizas 
         Height          =   315
         Left            =   195
         TabIndex        =   3
         ToolTipText     =   "Consulta de pólizas del departamento"
         Top             =   1620
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   16777215
         AppearanceThemes=   1
         BackColorDown   =   15133676
         TransparentColor=   16249836
         Caption         =   "Consulta de pólizas"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdCorte 
         Height          =   315
         Left            =   195
         TabIndex        =   0
         ToolTipText     =   "Corte de caja"
         Top             =   270
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   16777215
         AppearanceThemes=   1
         BackColorDown   =   15133676
         TransparentColor=   16249836
         Caption         =   "Corte"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdCuentaChequeTransf 
         Height          =   315
         Left            =   195
         TabIndex        =   1
         ToolTipText     =   "Cuenta bancaria y RFC de pagos recibidos"
         Top             =   720
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   16777215
         AppearanceThemes=   1
         BackColorDown   =   15133676
         TransparentColor=   16249836
         Caption         =   "Cuenta bancaria y RFC de pagos recibidos"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
   End
End
Attribute VB_Name = "frmMenuCorte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConsultaPolizas_Click()
    frmConsultaPolizasDepartamento.vllngNumOpcion = IIf(cgstrModulo = "PV", 314, 630)
    frmConsultaPolizasDepartamento.Show vbModal
End Sub

Private Sub cmdCorte_Click()
    frmCorte.Show vbModal
End Sub

Private Sub cmdCuentaChequeTransf_Click()
    frmCuentaChequeTrans.Show vbModal
End Sub

Private Sub cmdTransferencias_Click()
    frmTransferencia.Show vbModal
End Sub

Private Sub Form_Activate()
    fblnHabilitaObjetos Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = frmMenuPrincipal.Icon
End Sub

Private Sub cmdReImpR_Click()
frmRegTerm.Show vbModal
End Sub
