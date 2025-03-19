VERSION 5.00
Object = "{6B45E0EA-D03D-4CBB-94F4-B6AD155551A1}#1.1#0"; "MyFramePanel.ocx"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.1#0"; "MyCommandButton.ocx"
Begin VB.Form frmMenuCatalogosCaja 
   BackColor       =   &H00F7F3EC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menú de catálogos"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8160
   Icon            =   "frmMenuCatalogosCaja.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   8160
   StartUpPosition =   2  'CenterScreen
   Begin MyCommandButton.MyButton cmdTerminales 
      Height          =   315
      Left            =   4150
      TabIndex        =   15
      ToolTipText     =   "Mantenimiento terminales"
      Top             =   3310
      Width           =   3720
      _ExtentX        =   6562
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
      Caption         =   "Terminales"
      CaptionPosition =   4
      DepthEvent      =   1
      PictureAlignment=   1
      ShowFocus       =   -1  'True
   End
   Begin MyFramePanel.MyFrame Frame1 
      Height          =   4140
      Left            =   120
      Top             =   120
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   7303
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
      Begin MyCommandButton.MyButton cmdMantoCajaChica 
         Height          =   315
         Left            =   210
         TabIndex        =   4
         ToolTipText     =   "Mantenimiento a los conceptos de movimientos en caja chica"
         Top             =   1980
         Width           =   3720
         _ExtentX        =   6562
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
         Caption         =   "Conceptos de caja chica"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdMantoPaquetes 
         Height          =   315
         Left            =   4035
         TabIndex        =   14
         ToolTipText     =   "Mantenimiento de paquetes"
         Top             =   2790
         Width           =   3720
         _ExtentX        =   6562
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
         Caption         =   "Paquetes"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdMantoElementoFijoPresupuesto 
         Height          =   315
         Left            =   210
         TabIndex        =   7
         ToolTipText     =   "Mantenimiento de elementos fijos para presupuestos"
         Top             =   3195
         Width           =   3720
         _ExtentX        =   6562
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
         Caption         =   "Elementos fijos para presupuestos"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdMantoConceptosPago 
         Height          =   315
         Left            =   210
         TabIndex        =   5
         ToolTipText     =   "Mantenimiento de conceptos de pago"
         Top             =   2385
         Width           =   3720
         _ExtentX        =   6562
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
         Caption         =   "Conceptos de entradas y salidas de dinero"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdMantoConceptosFactura 
         Height          =   315
         Left            =   210
         TabIndex        =   6
         ToolTipText     =   "Mantenimiento de conceptos de factura"
         Top             =   2790
         Width           =   3720
         _ExtentX        =   6562
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
         Caption         =   "Conceptos de factura"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdMantoClientes 
         Height          =   315
         Left            =   210
         TabIndex        =   1
         ToolTipText     =   "Mantenimiento de clientes"
         Top             =   765
         Width           =   3720
         _ExtentX        =   6562
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
         Caption         =   "Clientes"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdMantoFormasPago 
         Height          =   315
         Left            =   4035
         TabIndex        =   8
         ToolTipText     =   "Mantenimiento de formas de pago"
         Top             =   360
         Width           =   3720
         _ExtentX        =   6562
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
         Caption         =   "Formas de pago"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdMantoOtrosConceptos 
         Height          =   315
         Left            =   4035
         TabIndex        =   12
         ToolTipText     =   "Mantenimiento de otros conceptos de cargo a pacientes"
         Top             =   1980
         Width           =   3720
         _ExtentX        =   6562
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
         Caption         =   "Otros conceptos de cargo"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdComisiones 
         Height          =   315
         Left            =   210
         TabIndex        =   3
         ToolTipText     =   "Mantenimiento de comisiones de honorarios médicos"
         Top             =   1575
         Width           =   3720
         _ExtentX        =   6562
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
         Caption         =   "Comisiones de honorarios médicos"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdExternos 
         Height          =   315
         Left            =   4035
         TabIndex        =   13
         ToolTipText     =   "Mantenimiento de pacientes externos"
         Top             =   2385
         Width           =   3720
         _ExtentX        =   6562
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
         Caption         =   "Pacientes externos"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdListaNegra 
         Height          =   315
         Left            =   4035
         TabIndex        =   11
         ToolTipText     =   "Mantenimiento de la lista de deudores incobrables"
         Top             =   1575
         Width           =   3720
         _ExtentX        =   6562
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
         Caption         =   "Lista de deudores incobrables"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdGruposCargos 
         Height          =   315
         Left            =   4035
         TabIndex        =   9
         ToolTipText     =   "Mantenimiento de los grupos de cargos"
         Top             =   765
         Width           =   3720
         _ExtentX        =   6562
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
         Caption         =   "Grupos de cargos para paquetes"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdComisionesBancarias 
         Height          =   315
         Left            =   210
         TabIndex        =   2
         ToolTipText     =   "Mantenimiento de comisiones bancarias"
         Top             =   1170
         Width           =   3720
         _ExtentX        =   6562
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
         Caption         =   "Comisiones bancarias"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdTipoCargoBancario 
         Height          =   315
         Left            =   4035
         TabIndex        =   16
         ToolTipText     =   "Mantenimiento de tipos de cargos bancarios"
         Top             =   3600
         Width           =   3720
         _ExtentX        =   6562
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
         Caption         =   "Tipos de cargos bancarios"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdGruposTiposClientes 
         Height          =   315
         Left            =   4035
         TabIndex        =   10
         ToolTipText     =   "Mantenimiento de los grupos de tipos de clientes"
         Top             =   1170
         Width           =   3720
         _ExtentX        =   6562
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
         Caption         =   "Grupos de tipos de clientes"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdAuditoriadeCargos 
         Height          =   315
         Left            =   210
         TabIndex        =   0
         ToolTipText     =   "Mantenimiento de auditoría de cargos"
         Top             =   360
         Width           =   3720
         _ExtentX        =   6562
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
         Caption         =   "Auditoría de cargos"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
   End
End
Attribute VB_Name = "frmMenuCatalogosCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
' Forma para registro y consulta del corte de caja, se muestran los cortes de todas
' las cajas, se puede cerrar el corte de una caja que no es del departamento que
' consulta, con advertencia.
' Fecha de programación: Jueves 1 de Febrero de 2001
'------------------------------------------------------------------------------------
' Ultimas modificaciones, especificar:
' 06/Febrero/2003 : Se agrega el catálogo de externos
'------------------------------------------------------------------------------------

Private Sub cmdAuditoriadeCargos_Click()
    If cgstrModulo = "PV" Then
        frmAuditoriaCargo.vllngNumeroOpcion = 7039
        frmAuditoriaCargo.HelpContextID = 26
        frmAuditoriaCargo.Show vbModal
    ElseIf cgstrModulo = "SI" Then
        frmAuditoriaCargo.vllngNumeroOpcion = 7040
        frmAuditoriaCargo.HelpContextID = 26
        frmAuditoriaCargo.Show vbModal
    End If
        
End Sub

Private Sub cmdComisiones_Click()
    If cgstrModulo = "PV" Then
        frmMtoComisiones.vllngNumeroOpcion = 348
        frmMtoComisiones.HelpContextID = 26
        frmMtoComisiones.Show vbModal
    ElseIf cgstrModulo = "SI" Then
        frmMtoComisiones.vllngNumeroOpcion = 1120
        frmMtoComisiones.HelpContextID = 26
        frmMtoComisiones.Show vbModal
    End If
End Sub

Private Sub cmdComisionesBancarias_Click()
    If cgstrModulo = "PV" Then
        frmMtoComisionesBancarias.vllngNumeroOpcion = 3060
        frmMtoComisionesBancarias.HelpContextID = 26
        frmMtoComisionesBancarias.Show vbModal
    ElseIf cgstrModulo = "SI" Then
        frmMtoComisionesBancarias.vllngNumeroOpcion = 3061
        frmMtoComisionesBancarias.HelpContextID = 26
        frmMtoComisionesBancarias.Show vbModal
    End If
End Sub

Private Sub cmdExternos_Click()
    
    frmAdmisionPaciente.vlblnMostrarTabGenerales = True
    frmAdmisionPaciente.vlblnMostrarTabInternamiento = True
    frmAdmisionPaciente.vlblnMostrarTabInternos = False
    frmAdmisionPaciente.vlblnMostrarTabPrepagos = False
    frmAdmisionPaciente.vlblnMostrarTabIngresosPrevios = False
    frmAdmisionPaciente.vlblnMostrarTabEgresados = False
    frmAdmisionPaciente.vlblnMostrarTabExternos = False
    frmAdmisionPaciente.vlintPestañaInicial = 0
    
    frmAdmisionPaciente.blnAbrirCuenta = False
    frmAdmisionPaciente.blnActivar = False
    frmAdmisionPaciente.blnHabilitarAbrirCuenta = True
    frmAdmisionPaciente.blnHabilitarActivar = False
    frmAdmisionPaciente.blnHabilitarReporte = True
    frmAdmisionPaciente.blnConsulta = False
    frmAdmisionPaciente.vglngExpedienteConsulta = 0
    frmAdmisionPaciente.blnHonorariosCC = False
    
    If cgstrModulo = "PV" Then
        frmAdmisionPaciente.vllngNumeroOpcionExterno = 352
    ElseIf cgstrModulo = "SI" Then
        frmAdmisionPaciente.vllngNumeroOpcionExterno = 2378
    End If
 
    frmAdmisionPaciente.Show vbModal, Me

End Sub

Private Sub cmdGruposCargos_Click()
   If cgstrModulo = "PV" Then
        'frmGruposCargos.vllngNumeroOpcion = 2054
        frmGruposCargos.Show vbModal, Me
    ElseIf cgstrModulo = "SI" Then
        'frmGruposCargos.vllngNumeroOpcion = 2379
        frmGruposCargos.Show vbModal, Me
    End If
End Sub

Private Sub cmdGruposTiposClientes_Click()
    FrmMtoGruposTiposCliente.HelpContextID = 26
    FrmMtoGruposTiposCliente.Show vbModal
End Sub

Private Sub cmdListaNegra_Click()
   If cgstrModulo = "PV" Then
        frmListaNegra.vllngNumeroOpcion = 2054
        frmListaNegra.Show vbModal, Me
    ElseIf cgstrModulo = "SI" Then
        frmListaNegra.vllngNumeroOpcion = 2379
        frmListaNegra.Show vbModal, Me
    End If
End Sub

Private Sub cmdMantoCajaChica_Click()
    frmMantoCajaChica.HelpContextID = 26
    frmMantoCajaChica.Show vbModal
End Sub

Private Sub cmdMantoClientes_Click()

    If cgstrModulo = "PV" Then
        frmMantoClientes.llngNumOpcion = 318
        frmMantoClientes.lblnTodosClientes = False
    ElseIf cgstrModulo = "SI" Then
        frmMantoClientes.llngNumOpcion = 1137
        frmMantoClientes.lblnTodosClientes = False
    End If

    frmMtoConceptos.HelpContextID = 26
    frmMantoClientes.Show vbModal

End Sub

Private Sub cmdMantoConceptosFactura_Click()
    If cgstrModulo = "PV" Then
        frmMtoConceptos.llngNumOpcion = 317
    ElseIf cgstrModulo = "SI" Then
        frmMtoConceptos.llngNumOpcion = 1121
    End If
    
    frmMtoConceptos.HelpContextID = 26
    frmMtoConceptos.Show vbModal
End Sub

Private Sub cmdMantoConceptosPago_Click()
    frmMantoConceptoPago.HelpContextID = 26
    frmMantoConceptoPago.Show vbModal
End Sub

Private Sub cmdMantoElementoFijoPresupuesto_Click()
    frmMantoElementoFijoPresupuesto.HelpContextID = 26
    frmMantoElementoFijoPresupuesto.Show vbModal
End Sub

Private Sub cmdMantoFormasPago_Click()
    frmMantoFormasPago.HelpContextID = 26
    frmMantoFormasPago.Show vbModal
End Sub

Private Sub cmdMantoOtrosConceptos_Click()
    FrmMtoOtrosConceptos.HelpContextID = 26
    FrmMtoOtrosConceptos.Show vbModal
End Sub

Private Sub cmdMantoPaquetes_Click()
    FrmMtoPaquetes.HelpContextID = 26
    FrmMtoPaquetes.Show vbModal
End Sub

Private Sub cmdTerminales_Click()
 
frmMantoTerminales.Show vbModal
End Sub

Private Sub cmdTipoCargoBancario_Click()
    If cgstrModulo = "PV" Then
        frmMtoTipoCargoBancario.vllngNumeroOpcion = 3062
        frmMtoTipoCargoBancario.HelpContextID = 26
        frmMtoTipoCargoBancario.Show vbModal
    ElseIf cgstrModulo = "SI" Then
        frmMtoTipoCargoBancario.vllngNumeroOpcion = 3063
        frmMtoTipoCargoBancario.HelpContextID = 26
        frmMtoTipoCargoBancario.Show vbModal
    End If
End Sub

Private Sub Form_Activate()
    fblnHabilitaObjetos Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Me.Icon = frmMenuPrincipal.Icon
End Sub


