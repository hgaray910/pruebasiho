VERSION 5.00
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.1#0"; "MyCommandButton.ocx"
Begin VB.Form frmPaquetesMedico 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Paquetes asignados al médico"
   ClientHeight    =   7080
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6405
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   6405
   StartUpPosition =   2  'CenterScreen
   Begin MyCommandButton.MyButton MyButton1 
      Height          =   735
      Left            =   6600
      TabIndex        =   4
      Top             =   2640
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   1296
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmPaquetesMedico.frx":0000
      BackColorDown   =   15133676
      TransparentColor=   16777215
      DepthEvent      =   1
      PictureDisabled =   "frmPaquetesMedico.frx":0712
      ShowFocus       =   -1  'True
   End
   Begin VB.Frame fraPaquetesMedico 
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
      Height          =   6495
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   6135
      Begin VB.ListBox lstPaquetesMedico 
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
         Height          =   4875
         Left            =   90
         Sorted          =   -1  'True
         TabIndex        =   0
         ToolTipText     =   "Listado de paquetes asignados al médico"
         Top             =   1530
         Width           =   5955
      End
      Begin VB.Label lblMensajePaquetes 
         BackColor       =   &H80000005&
         Caption         =   "Seleccione el paquete"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1260
         Left            =   90
         TabIndex        =   3
         Top             =   225
         Width           =   5955
      End
   End
   Begin MyCommandButton.MyButton cmdAceptar 
      Height          =   375
      Left            =   2100
      TabIndex        =   1
      Top             =   6590
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "Aceptar"
      DepthEvent      =   1
      ShowFocus       =   -1  'True
   End
End
Attribute VB_Name = "frmPaquetesMedico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public glngIdPaquete As Long    'Clave del paquete seleccionado
Public gdblPrecio As Double     'Precio del paquete
Dim ldblIncrementoTarifa As Double
Dim llngCveTipoPaciente As Long 'clave del tipo de paciente
Dim llngEmpresa As Long         'clave de la empresa con la que tiene convenio el paciente

Private Sub cmdAceptar_Click()
    Dim laryParametrosSalida() As String
    Dim dblPrecio As Double
    
    If lstPaquetesMedico.ListIndex <> -1 Then
        gdblPrecio = 0
        pCargaArreglo laryParametrosSalida, "|" & adDecimal & "||" & adDecimal
        frsEjecuta_SP lstPaquetesMedico.ItemData(lstPaquetesMedico.ListIndex) & "|" & "PA" & "|" & llngCveTipoPaciente & "|" & llngEmpresa & "|I|0|01/01/1900|" & vgintClaveEmpresaContable, "SP_PVSELOBTENERPRECIO", , , laryParametrosSalida
        pObtieneValores laryParametrosSalida, dblPrecio, ldblIncrementoTarifa
        If dblPrecio = -1 Or dblPrecio = 0 Then
            'El elemento seleccionado no cuenta con un precio capturado.
            MsgBox SIHOMsg(301), vbInformation, "Mensaje"
            lstPaquetesMedico.SetFocus
        Else
            glngIdPaquete = lstPaquetesMedico.ItemData(lstPaquetesMedico.ListIndex)
            gdblPrecio = dblPrecio
            Unload Me
        End If
    End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys vbTab
    ElseIf KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Me.Icon = frmMenuPrincipal.Icon
End Sub

Public Sub pIniciar(rs As ADODB.Recordset, strMensaje As String, lngTipoPaciente As Long, lngEmpresa As Long)
    lblMensajePaquetes.Caption = strMensaje
    llngCveTipoPaciente = lngTipoPaciente
    llngEmpresa = lngEmpresa
    gdblPrecio = 0
    If rs.RecordCount <> 0 Then
        pLlenarListRs lstPaquetesMedico, rs, 0, 1
        lstPaquetesMedico.ListIndex = 0
    End If
End Sub


