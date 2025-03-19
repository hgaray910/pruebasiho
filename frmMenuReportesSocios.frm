VERSION 5.00
Object = "{6B45E0EA-D03D-4CBB-94F4-B6AD155551A1}#1.1#0"; "MyFramePanel.ocx"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.1#0"; "MyCommandButton.ocx"
Begin VB.Form frmMenuReportesSocios 
   BackColor       =   &H00F7F3EC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reportes de socios"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4635
   Icon            =   "frmMenuReportesSocios.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   4635
   StartUpPosition =   2  'CenterScreen
   Begin MyFramePanel.MyFrame Frame1 
      Height          =   3780
      Left            =   120
      Top             =   120
      Width           =   4380
      _ExtentX        =   7726
      _ExtentY        =   6668
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
      Begin MyCommandButton.MyButton cmdBajasSocios 
         Height          =   315
         Left            =   240
         TabIndex        =   1
         ToolTipText     =   "Reporte de bajas de socios"
         Top             =   760
         Width           =   3855
         _ExtentX        =   6800
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
         Caption         =   "Bajas"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdReporteCuotas 
         Height          =   315
         Left            =   240
         TabIndex        =   2
         ToolTipText     =   "Reporte de cuotas de socios"
         Top             =   1155
         Width           =   3855
         _ExtentX        =   6800
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
         Caption         =   "Cuotas"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdSugerenciaBajas 
         Height          =   315
         Left            =   240
         TabIndex        =   7
         ToolTipText     =   "Reporte de sugerencias de bajas para socios"
         Top             =   3105
         Width           =   3855
         _ExtentX        =   6800
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
         Caption         =   "Sugerencias de baja"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdEgresadosHospital 
         Height          =   315
         Left            =   240
         TabIndex        =   3
         ToolTipText     =   "Reporte de socios egresados"
         Top             =   1545
         Width           =   3855
         _ExtentX        =   6800
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
         Caption         =   "Egresados"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdAltasSocios 
         Height          =   315
         Left            =   240
         TabIndex        =   0
         ToolTipText     =   "Reporte de altas de socios"
         Top             =   370
         Width           =   3855
         _ExtentX        =   6800
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
         Caption         =   "Altas"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdEstadoCuenta 
         Height          =   315
         Left            =   240
         TabIndex        =   4
         ToolTipText     =   "Reporte de estado de cuenta"
         Top             =   1935
         Width           =   3855
         _ExtentX        =   6800
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
         Caption         =   "Estado de cuenta"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdSEPOMEX 
         Height          =   315
         Left            =   240
         TabIndex        =   6
         ToolTipText     =   "Reporte de etiquetas SEPOMEX"
         Top             =   2715
         Width           =   3855
         _ExtentX        =   6800
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
         Caption         =   "Etiquetas SEPOMEX"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdCorreo 
         Height          =   315
         Left            =   240
         TabIndex        =   5
         ToolTipText     =   "Reporte de etiquetas para correo"
         Top             =   2325
         Width           =   3855
         _ExtentX        =   6800
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
         Caption         =   "Etiquetas para correo"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
   End
End
Attribute VB_Name = "frmMenuReportesSocios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim lblnSalir As Boolean '|  Indica si la forma se descará en el Activate

Private Sub cmdAltasSocios_Click()
On Error GoTo NotificaError

    frmRptAltasSocios.Show vbModal

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdAltasSocios_Click"))
    Unload Me

End Sub



Private Sub cmdBajasSocios_Click()
On Error GoTo NotificaError

    frmRptBajasSocios.Show vbModal

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdBajasSocios_Click"))
    Unload Me

End Sub

Private Sub cmdCorreo_Click()
On Error GoTo NotificaError
frmRptEtiquetas.vginttipo = 0
    frmRptEtiquetas.Caption = "Etiquetas para correo"
    frmRptEtiquetas.Show vbModal, Me
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmd_Click"))
End Sub

Private Sub cmdReporteCuotas_Click()
On Error GoTo NotificaError
    
    frmReporteSaldoSocios.Show vbModal, Me

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdReporteCuotas_Click"))
    Unload Me
End Sub


Private Sub cmdEgresadosHospital_Click()
On Error GoTo NotificaError
    
    frmRptSociosEgresados.Show vbModal

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdEgresadosHospital_Click"))
    Unload Me

End Sub

Private Sub cmdEstadoCuenta_Click()
On Error GoTo NotificaError
    
    frmReporteSociosEdoCuenta.Show vbModal

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdEstadoCuenta_Click"))
    Unload Me
End Sub


Private Sub cmdSEPOMEX_Click()
On Error GoTo NotificaError
frmRptEtiquetas.vginttipo = 1
    frmRptEtiquetas.Caption = "Etiquetas SEPOMEX"
    frmRptEtiquetas.Show vbModal, Me
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmd_Click"))
End Sub

Private Sub cmdSugerenciaBajas_Click()
On Error GoTo NotificaError

    frmRptSugerenciaBajas.Show vbModal, Me
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdSugerenciaBajas_Click"))

End Sub

Private Sub Form_Activate()
On Error GoTo NotificaError
    
    fblnHabilitaObjetos Me
    If lblnSalir Then Unload Me
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Activate"))
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    
    If KeyAscii = 27 Then Unload Me

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyPress"))
    Unload Me
End Sub

Private Sub Form_Load()
Dim strTitulo As String

    Me.Icon = frmMenuPrincipal.Icon
    lblnSalir = False
    
End Sub

