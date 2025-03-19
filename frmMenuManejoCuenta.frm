VERSION 5.00
Object = "{6B45E0EA-D03D-4CBB-94F4-B6AD155551A1}#1.1#0"; "MyFramePanel.ocx"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.1#0"; "MyCommandButton.ocx"
Begin VB.Form frmMenuManejoCuenta 
   BackColor       =   &H00F7F3EC&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Menú de manejo de cuenta"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
   Icon            =   "frmMenuManejoCuenta.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MyFramePanel.MyFrame Frame1 
      Height          =   4275
      Left            =   120
      Top             =   120
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   7541
      BackColor       =   16777215
      ForeColor       =   8677689
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
      Begin MyCommandButton.MyButton cmdGpoFacturas 
         Height          =   315
         Left            =   240
         TabIndex        =   6
         ToolTipText     =   "Creación de grupos de cuentas"
         Top             =   2580
         Width           =   3210
         _ExtentX        =   5662
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
         Caption         =   "Grupo de cuentas"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdAsignacionPaquetes 
         Height          =   315
         Left            =   240
         TabIndex        =   1
         ToolTipText     =   "Asignación de paquetes"
         Top             =   660
         Width           =   3210
         _ExtentX        =   5662
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
         Caption         =   "Asignación paquetes"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdExclusion 
         Height          =   315
         Left            =   240
         TabIndex        =   3
         ToolTipText     =   "Exclusión de cargos"
         Top             =   1440
         Width           =   3210
         _ExtentX        =   5662
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
         Caption         =   "Exclusión"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdTraslado 
         Height          =   315
         Left            =   240
         TabIndex        =   9
         Top             =   3720
         Width           =   3210
         _ExtentX        =   5662
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
         Caption         =   "Traslado cargos"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdManejoCuenta 
         Height          =   315
         Left            =   240
         TabIndex        =   0
         Top             =   255
         Width           =   3210
         _ExtentX        =   5662
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
         Caption         =   "Abrir / Cerrar cuentas"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdCargoCuarto 
         Height          =   315
         Left            =   240
         TabIndex        =   2
         Top             =   1050
         Width           =   3210
         _ExtentX        =   5662
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
         Caption         =   "Cargo de cuartos"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdFacturacionDirecta 
         Height          =   315
         Left            =   240
         TabIndex        =   4
         ToolTipText     =   "Facturación directa a clientes"
         Top             =   1830
         Width           =   3210
         _ExtentX        =   5662
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
         Caption         =   "Facturación directa"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdNotasCredito 
         Height          =   315
         Left            =   240
         TabIndex        =   8
         ToolTipText     =   "Notas de crédito y cargo"
         Top             =   3330
         Width           =   3210
         _ExtentX        =   5662
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
         Caption         =   "Notas de crédito y cargo"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdInterfazVitamedica 
         Height          =   315
         Left            =   240
         TabIndex        =   7
         ToolTipText     =   "Interfaz con Vitamédica"
         Top             =   2955
         Width           =   3210
         _ExtentX        =   5662
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
         Caption         =   "Interfaz con Vitamédica"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdFacturacionMasiva 
         Height          =   315
         Left            =   240
         TabIndex        =   5
         ToolTipText     =   "Facturación masiva a clientes"
         Top             =   2200
         Width           =   3210
         _ExtentX        =   5662
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
         Caption         =   "Facturación masiva"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
   End
End
Attribute VB_Name = "frmMenuManejoCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdAsignacionPaquetes_Click()
    frmAsignacionPaquetesPaciente.HelpContextID = 12
    frmAsignacionPaquetesPaciente.Show vbModal
End Sub

Private Sub cmdCargoCuarto_Click()
    frmCargaCuartos.vllngNumeroOpcion = 339
    frmCargaCuartos.HelpContextID = 8
    frmCargaCuartos.Show vbModal
End Sub

Private Sub cmdFacturacionMasiva_Click()
    frmFacturacionMasiva.Show vbModal
End Sub

Private Sub cmdExclusion_Click()
    frmExclusionCargos.HelpContextID = 11
    frmExclusionCargos.Show vbModal
End Sub

Private Sub cmdFacturacionDirecta_Click()
    frmFacturacionDirecta.Show vbModal
End Sub

Private Sub cmdGpoFacturas_Click()
    frmGrupoFacturas.HelpContextID = 13
    frmGrupoFacturas.Show vbModal
End Sub

Private Sub cmdInterfazVitamedica_Click()

    frmInterfazVitamedica.Show vbModal

End Sub

Private Sub cmdManejoCuenta_Click()
    frmManejoCuenta.HelpContextID = 10
    frmManejoCuenta.Show vbModal
End Sub

Private Sub cmdNotasCredito_Click()
    
    frmNotas.Show vbModal
    
End Sub

Private Sub cmdTraslado_Click()
    frmTrasladoCargos.HelpContextID = 9
    frmTrasladoCargos.Show vbModal
End Sub

Private Sub Form_Activate()
    fblnHabilitaObjetos Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub
