VERSION 5.00
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.1#0"; "MyCommandButton.ocx"
Begin VB.Form frmMenuRequisicion 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menú de requisiciones"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4215
   HelpContextID   =   12
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   4215
   StartUpPosition =   2  'CenterScreen
   Begin MyCommandButton.MyButton cmdReqArticulo 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Requisiciones de artículos"
      Top             =   120
      Width           =   3975
      _ExtentX        =   7011
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
      MaskColor       =   16777215
      AppearanceThemes=   1
      BackColorOver   =   -2147483633
      BackColorFocus  =   -2147483633
      BackColorDisabled=   -2147483633
      BorderColor     =   -2147483627
      TransparentColor=   16777215
      Caption         =   "Requisiciones de artículos"
      CaptionPosition =   4
      DepthEvent      =   1
      PictureAlignment=   1
      ShowFocus       =   -1  'True
   End
   Begin MyCommandButton.MyButton cmdReqCargoDirecto 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Requisiciones de cargos directos"
      Top             =   600
      Width           =   3975
      _ExtentX        =   7011
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
      MaskColor       =   16777215
      AppearanceThemes=   1
      BackColorOver   =   -2147483633
      BackColorFocus  =   -2147483633
      BackColorDisabled=   -2147483633
      BorderColor     =   -2147483627
      TransparentColor=   16777215
      Caption         =   "Requisiciones de activo fijo"
      CaptionPosition =   4
      DepthEvent      =   1
      PictureAlignment=   1
      ShowFocus       =   -1  'True
   End
   Begin MyCommandButton.MyButton cmdRequisicionPaciente 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Requisiciones con cargo a la cuenta del paciente por materiales"
      Top             =   1080
      Width           =   3975
      _ExtentX        =   7011
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
      MaskColor       =   16777215
      AppearanceThemes=   1
      BackColorOver   =   -2147483633
      BackColorFocus  =   -2147483633
      BackColorDisabled=   -2147483633
      BorderColor     =   -2147483627
      TransparentColor=   16777215
      Caption         =   "Requisiciones con cargo a paciente"
      CaptionPosition =   4
      DepthEvent      =   1
      PictureAlignment=   1
      ShowFocus       =   -1  'True
   End
   Begin MyCommandButton.MyButton cmdReqServicioInterno 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Requisiciones de servicios internos"
      Top             =   1560
      Width           =   3975
      _ExtentX        =   7011
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
      MaskColor       =   16777215
      AppearanceThemes=   1
      BackColorOver   =   -2147483633
      BackColorFocus  =   -2147483633
      BackColorDisabled=   -2147483633
      BorderColor     =   -2147483627
      TransparentColor=   16777215
      Caption         =   "Requisiciones de servicios internos"
      CaptionPosition =   4
      DepthEvent      =   1
      PictureAlignment=   1
      ShowFocus       =   -1  'True
   End
   Begin MyCommandButton.MyButton cmdReqAutomatica 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "Requisiciones automáticas de faltantes en base al máximo y mínimo."
      Top             =   2040
      Width           =   3975
      _ExtentX        =   7011
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
      MaskColor       =   16777215
      AppearanceThemes=   1
      BackColorOver   =   -2147483633
      BackColorFocus  =   -2147483633
      BackColorDisabled=   -2147483633
      BorderColor     =   -2147483627
      TransparentColor=   16777215
      Caption         =   "Requisiciones automáticas de faltantes"
      CaptionPosition =   4
      DepthEvent      =   1
      PictureAlignment=   1
      ShowFocus       =   -1  'True
   End
End
Attribute VB_Name = "frmMenuRequisicion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Modifiqué para mandar llamar la forma de requisiciones automáticas de diferente manera
'-------------------------------------------------------------------------------------
'| Nombre del Formulario    : frmMenuRequisicion
'-------------------------------------------------------------------------------------
'| Objetivo: Llamar las diferentes requisiciones
'-------------------------------------------------------------------------------------


Private Sub cmdAutorizacionRequisiciones_Click()

On Error GoTo NotificaError
   
   frmConsultaRequisicion.HelpContextID = 7
   frmConsultaRequisicion.Show vbModal, Me

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdReqArticulo_Click"))
    Unload Me
    
End Sub

Private Sub cmdReqArticulo_Click()
    On Error GoTo NotificaError

   frmRequisicion.vllngNumeroOpcionModulo = flngObtenOpcion(cmdReqArticulo.Name)
   frmRequisicion.Show vbModal, Me

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdReqArticulo_Click"))
    Unload Me
End Sub

Private Sub cmdReqAutomatica_Click()
    On Error GoTo NotificaError
    
    Dim rsIvparametro As New ADODB.Recordset

    Set rsIvparametro = frsRegresaRs("SELECT  NVL(BITCONSOLIDARALMACEN,0) AS BITALMACEN, NVL(TNYALMACENPRINCIPAL,0) AS ALMPRINCIPAL, NVL(VCHCVESALMACENCONSOLIDAR,'-') as claves FROM IvParametro WHERE tnyClaveEmpresa = " & CStr(vgintClaveEmpresaContable), adLockOptimistic, adOpenKeyset)
    
    If rsIvparametro.RecordCount > 0 Then
        If rsIvparametro!BITALMACEN = 1 And (rsIvparametro!ALMPRINCIPAL <= 0 Or rsIvparametro!claves = "-") Then
         MsgBox "No se puede realizar la requisición automática de faltantes hasta que se realice la configuración o se desactive el parámetro consolidación de almacenes.", vbInformation + vbOKOnly, "Mensaje"
        Else
           frmMaximoMinimo.lstrModo = "C"
           frmMaximoMinimo.HelpContextID = 30
           frmMaximoMinimo.Show vbModal, Me
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdReqAutomatica_Click"))
    Unload Me
End Sub

Private Sub cmdReqCargoDirecto_Click()
    On Error GoTo NotificaError
    
    frmRequisCargoDirecto.vllngNumeroOpcion = flngObtenOpcion(cmdReqCargoDirecto.Name)
    frmRequisCargoDirecto.Show vbModal, Me

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdReqCargoDirecto_Click"))
    Unload Me
End Sub

Private Sub cmdReqServicioInterno_Click()
    On Error GoTo NotificaError

    frmSolicitudServicio.vllngNumeroOpcion = flngObtenOpcion(cmdReqServicioInterno.Name)
    frmSolicitudServicio.Show vbModal, Me

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdReqServicioInterno_Click"))
    Unload Me
End Sub

Private Sub cmdRequisicionPaciente_Click()
    frmRequisicionCargoPac.HelpContextID = 55
    frmRequisicionCargoPac.Show vbModal, Me
End Sub

Private Sub Form_Activate()
    vgstrNombreForm = Me.Name
    fblnHabilitaObjetos Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError

    If KeyAscii = 27 Then
        Unload Me
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyPress"))
    Unload Me
End Sub

Private Sub Form_Load()
    
    Me.Icon = frmMenuPrincipal.Icon

End Sub
