VERSION 5.00
Object = "{6B45E0EA-D03D-4CBB-94F4-B6AD155551A1}#1.1#0"; "MyFramePanel.ocx"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.1#0"; "MyCommandButton.ocx"
Begin VB.Form frmMenuSocios 
   BackColor       =   &H00F7F3EC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menú de socios"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   3975
   StartUpPosition =   2  'CenterScreen
   Begin MyFramePanel.MyFrame Frame1 
      Height          =   3015
      Left            =   120
      Top             =   120
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   5318
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
      Begin MyCommandButton.MyButton cmdDependientes 
         Height          =   315
         Left            =   240
         TabIndex        =   1
         ToolTipText     =   "Administración de dependientes"
         Top             =   720
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
         Caption         =   "Dependientes"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdSocios 
         Height          =   315
         Left            =   240
         TabIndex        =   5
         ToolTipText     =   "Administración de socios"
         Top             =   2460
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
         Caption         =   "Socios"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdCargos 
         Height          =   315
         Left            =   240
         TabIndex        =   0
         ToolTipText     =   "Cargos de socios"
         Top             =   285
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
         Caption         =   "Cargos"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdFacturacion 
         Height          =   315
         Left            =   240
         TabIndex        =   2
         ToolTipText     =   "Facturación de socios"
         Top             =   1155
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
         Caption         =   "Facturación"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdParametrosSocios 
         Height          =   315
         Left            =   240
         TabIndex        =   3
         ToolTipText     =   "Parámetros de socios"
         Top             =   1590
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
         Caption         =   "Parámetros"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdReportes 
         Height          =   315
         Left            =   240
         TabIndex        =   4
         ToolTipText     =   "Reportes de socios"
         Top             =   2040
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
         Caption         =   "Reportes"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
   End
End
Attribute VB_Name = "frmMenuSocios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCargos_Click()
    
    On Error GoTo NotificaError
    
    frmCargosSocios.Show vbModal, Me
        
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdCargos_Click"))
    
End Sub

Private Sub cmdDependientes_Click()

On Error GoTo NotificaError

    frmSocios.vlblnDependiente = True
    frmSocios.vlblnMostrarTabDependientes = True
    frmSocios.vlblnMostrarTabDictamenes = False
    frmSocios.vlblnMostrarTabDocumentacion = False
    frmSocios.vlblnMostrarTabDomicilios = False
    frmSocios.vlblnMostrarTabEstado = False
    frmSocios.Show vbModal, Me
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdDependientes_Click"))
    
End Sub

Private Sub cmdFacturacion_Click()

On Error GoTo NotificaError
        
    frmFacturacionMembresiaSocios.Show vbModal, Me
    Exit Sub

NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdFacturacion_Click_Click"))
    
End Sub

Private Sub cmdParametrosSocios_Click()
On Error GoTo NotificaError

    frmParametrosSocios.Show vbModal, Me
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdParametrosSocios_Click"))

End Sub

Private Sub cmdReportes_Click()
    
    On Error GoTo NotificaError
        
        frmMenuReportesSocios.Show vbModal, Me
        
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdReportes_Click"))
    
End Sub

Private Sub cmdSocios_Click()

On Error GoTo NotificaError

    frmSocios.vlblnDependiente = False
    frmSocios.vlblnMostrarTabDependientes = False
    frmSocios.vlblnMostrarTabDictamenes = True
    frmSocios.vlblnMostrarTabDocumentacion = True
    frmSocios.vlblnMostrarTabDomicilios = True
    frmSocios.vlblnMostrarTabEstado = True
    frmSocios.Show vbModal, Me
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdSocios_Click"))

End Sub

Private Sub Form_Activate()

On Error GoTo NotificaError

    fblnHabilitaObjetos Me
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Activate"))
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo NotificaError

        If KeyCode = vbKeyEscape Then
                KeyCode = 0
                Unload Me
        End If
        
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyDown"))

End Sub

Private Sub Form_Load()

On Error GoTo NotificaError

    Me.Icon = frmMenuPrincipal.Icon
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
    
End Sub

