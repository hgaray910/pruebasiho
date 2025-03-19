VERSION 5.00
Object = "{6B45E0EA-D03D-4CBB-94F4-B6AD155551A1}#1.1#0"; "MyFramePanel.ocx"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.1#0"; "MyCommandButton.ocx"
Begin VB.Form frmMenuDescuentos 
   BackColor       =   &H00F7F3EC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menú de descuentos"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
   Icon            =   "frmMenuDescuentos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   3975
   StartUpPosition =   2  'CenterScreen
   Begin MyFramePanel.MyFrame Frame1 
      Height          =   1605
      Left            =   120
      Top             =   120
      Width           =   3765
      _ExtentX        =   6641
      _ExtentY        =   2831
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
      Begin MyCommandButton.MyButton cmdTarjeta 
         Height          =   315
         Left            =   210
         TabIndex        =   2
         ToolTipText     =   "Asignación de tarjeta de descuento"
         Top             =   1110
         Width           =   3345
         _ExtentX        =   5900
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
         Caption         =   "Descuentos por tarjeta"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdDescuentos 
         Height          =   315
         Left            =   210
         TabIndex        =   1
         ToolTipText     =   "Asignación de descuentos"
         Top             =   675
         Width           =   3345
         _ExtentX        =   5900
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
         Caption         =   "Descuentos normales"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdDescuentosEspeciales 
         Height          =   315
         Left            =   210
         TabIndex        =   0
         ToolTipText     =   "Asignación de descuentos especiales"
         Top             =   240
         Width           =   3345
         _ExtentX        =   5900
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
         Caption         =   "Descuentos especiales"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
   End
End
Attribute VB_Name = "frmMenuDescuentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDescuentos_Click()
    
    frmMantoDescuentos.HelpContextID = 18
    frmMantoDescuentos.Show vbModal
    
End Sub

Private Sub cmdDescuentosEspeciales_Click()
    frmDescuentosEspeciales.HelpContextID = 17
    frmDescuentosEspeciales.Show vbModal
    
End Sub

Private Sub cmdTarjeta_Click()
    
    frmTarjeta.HelpContextID = 19
    frmTarjeta.Show vbModal

End Sub
Private Sub Form_Activate()
    fblnHabilitaObjetos Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

