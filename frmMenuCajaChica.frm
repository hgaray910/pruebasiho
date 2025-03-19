VERSION 5.00
Object = "{6B45E0EA-D03D-4CBB-94F4-B6AD155551A1}#1.1#0"; "MyFramePanel.ocx"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.1#0"; "MyCommandButton.ocx"
Begin VB.Form frmMenuCajaChica 
   BackColor       =   &H00F7F3EC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menú de caja chica"
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4110
   Icon            =   "frmMenuCajaChica.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   4110
   StartUpPosition =   2  'CenterScreen
   Begin MyFramePanel.MyFrame Frame1 
      Height          =   1260
      Left            =   120
      Top             =   120
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   2223
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
      Begin MyCommandButton.MyButton cmdCambioCheque 
         Height          =   315
         Left            =   240
         TabIndex        =   1
         ToolTipText     =   "Afectar el corte con la recepción de dinero"
         Top             =   675
         Width           =   3375
         _ExtentX        =   5953
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
         Caption         =   "Cambio de cheques / Transferencias"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdSalidaCaja 
         Height          =   315
         Left            =   240
         TabIndex        =   0
         ToolTipText     =   "Registrar facturas  o notas"
         Top             =   240
         Width           =   3375
         _ExtentX        =   5953
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
         Caption         =   "Salidas"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
   End
End
Attribute VB_Name = "frmMenuCajaChica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub cmdSalidaCaja_Click()
    frmCajaChica.HelpContextID = 14
    frmCajaChica.Show vbModal
End Sub


Private Sub cmdCambioCheque_Click()
    frmCambioCheque.Show vbModal
End Sub


Private Sub Form_Activate()
    fblnHabilitaObjetos Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub


