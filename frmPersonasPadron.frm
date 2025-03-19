VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmPersonasPadron 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Personas en el padrón"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10725
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   10725
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
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
      Height          =   3800
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   10575
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdPacientes 
         Height          =   3570
         Left            =   25
         TabIndex        =   1
         Top             =   180
         Width           =   10410
         _ExtentX        =   18362
         _ExtentY        =   6297
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         Cols            =   7
         ForeColorFixed  =   0
         ForeColorSel    =   16777215
         BackColorBkg    =   16777215
         BackColorUnpopulated=   16777215
         GridColor       =   -2147483638
         GridColorFixed  =   -2147483638
         GridColorUnpopulated=   16777215
         GridLinesFixed  =   1
         GridLinesUnpopulated=   1
         Appearance      =   0
         GridLineWidthFixed=   1
         FormatString    =   "|Expediente|Cuenta|Nombre|Fecha nacimiento|Domicilio|Estado"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   7
         _Band(0).GridLineWidthBand=   1
      End
   End
   Begin VB.Label lblMensaje558 
      BackColor       =   &H80000005&
      Caption         =   "Personas en padrón con el mismo nombre. Si está agregando una persona distinta presione la tecla <ESC>"
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
      Left            =   145
      TabIndex        =   2
      Top             =   120
      Width           =   10395
   End
End
Attribute VB_Name = "frmPersonasPadron"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public lngIdPersona As Long
Public strApPaterno As String
Public strApMaterno As String
Public strNombre As String
Public strNumAfiliacion As String
Public strSexo As String
Public dtmFechaNacimiento As Date
Public strDireccion As String
Public strDependencias As String
Public strVigencia As String

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        lngIdPersona = 0
        strNumAfiliacion = ""
        strApPaterno = ""
        strApMaterno = ""
        strNombre = ""
        strSexo = ""
        dtmFechaNacimiento = 0
        strDireccion = ""
        strDependencias = ""
        strVigencia = ""
        Me.Hide
    End If
End Sub

Private Sub Form_Load()
    Me.Icon = frmMenuPrincipal.Icon

End Sub

Private Sub grdPacientes_DblClick()
    lngIdPersona = Me.grdPacientes.TextMatrix(Me.grdPacientes.Row, 1)
    strNumAfiliacion = Me.grdPacientes.TextMatrix(Me.grdPacientes.Row, 5) & IIf(Trim(Me.grdPacientes.TextMatrix(Me.grdPacientes.Row, 6)) = "", "", "/" & Trim(Me.grdPacientes.TextMatrix(Me.grdPacientes.Row, 6)))
    strApPaterno = Me.grdPacientes.TextMatrix(Me.grdPacientes.Row, 3)
    strApMaterno = Me.grdPacientes.TextMatrix(Me.grdPacientes.Row, 4)
    strNombre = Me.grdPacientes.TextMatrix(Me.grdPacientes.Row, 2)
    strSexo = Me.grdPacientes.TextMatrix(Me.grdPacientes.Row, 12)
    dtmFechaNacimiento = Me.grdPacientes.TextMatrix(Me.grdPacientes.Row, 10)
    strDireccion = Me.grdPacientes.TextMatrix(Me.grdPacientes.Row, 11)
    strDependencias = Me.grdPacientes.TextMatrix(Me.grdPacientes.Row, 9)
    strVigencia = Me.grdPacientes.TextMatrix(Me.grdPacientes.Row, 8)
    Me.Hide
End Sub

Public Sub pConfigura()
    
    With grdPacientes
        .FixedCols = 1
        .FixedRows = 1
        .FormatString = "||Nombre|Apellido paterno|Apellido Materno|Afiliación|Benef.|Tipo|Vigente||Fecha nac.|Dirección|Sexo"
        
        .ColWidth(0) = 100   'fixed
        .ColWidth(1) = 0     'id persona
        .ColWidth(2) = 2500  'nombre
        .ColWidth(3) = 2500  'paterno
        .ColWidth(4) = 2500  'materno
        .ColWidth(5) = 1000  'afiliacion
        .ColWidth(6) = 800   'beneficiario
        .ColWidth(7) = 500   'Tipo
        .ColWidth(8) = 800   'vigencia
        .ColWidth(9) = 0     'dependencias
        .ColWidth(10) = 1000 'fecha nacimiento
        .ColWidth(11) = 3000 'direccion
        .ColWidth(12) = 500  'sexo
       
        
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignLeftCenter
        .ColAlignment(4) = flexAlignLeftCenter
        .ColAlignment(5) = flexAlignCenterCenter
        .ColAlignment(6) = flexAlignCenterCenter
        .ColAlignment(7) = flexAlignCenterCenter
        .ColAlignment(8) = flexAlignCenterCenter
        .ColAlignment(10) = flexAlignCenterCenter
        .ColAlignment(11) = flexAlignLeftCenter
        .ColAlignment(12) = flexAlignCenterCenter
        
        .ColAlignmentFixed(2) = flexAlignCenterCenter
        .ColAlignmentFixed(3) = flexAlignCenterCenter
        .ColAlignmentFixed(4) = flexAlignCenterCenter
        .ColAlignmentFixed(5) = flexAlignCenterCenter
        .ColAlignmentFixed(6) = flexAlignCenterCenter
        .ColAlignmentFixed(7) = flexAlignCenterCenter
        .ColAlignmentFixed(8) = flexAlignCenterCenter
        .ColAlignmentFixed(10) = flexAlignCenterCenter
        .ColAlignmentFixed(11) = flexAlignCenterCenter
        .ColAlignmentFixed(12) = flexAlignCenterCenter
    
    End With
End Sub

Private Sub grdPacientes_GotFocus()
    Me.grdPacientes.Col = 2
End Sub

Private Sub grdPacientes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        grdPacientes_DblClick
    End If
End Sub

Private Sub imgImagen_Click()

End Sub
