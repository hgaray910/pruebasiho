VERSION 5.00
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.1#0"; "MyCommandButton.ocx"
Begin VB.Form frmAutorizacion 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mensaje"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4950
   Icon            =   "frmAutorizacion.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
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
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   1120
      Width           =   4695
      Begin VB.OptionButton optExcluir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Cargar a la cuenta de la empresa"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   4215
      End
      Begin VB.OptionButton optExcluir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Cargar a la cuenta del paciente"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   500
         Width           =   4215
      End
   End
   Begin VB.TextBox txtCodigo 
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
      Height          =   375
      Left            =   120
      MaxLength       =   4000
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   2280
      Width           =   4695
   End
   Begin MyCommandButton.MyButton cmdCancelar 
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   2690
      Width           =   1095
      _ExtentX        =   1931
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
      Caption         =   "Cancelar"
      DepthEvent      =   1
      ShowFocus       =   -1  'True
   End
   Begin MyCommandButton.MyButton cmdAceptar 
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   2690
      Width           =   1215
      _ExtentX        =   2143
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
   Begin VB.Label lblMsg 
      BackColor       =   &H80000005&
      Caption         =   $"frmAutorizacion.frx":000C
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1305
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   4695
   End
   Begin VB.Label lblCodigo 
      BackColor       =   &H80000005&
      Caption         =   "Código de autorización"
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
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   2655
   End
End
Attribute VB_Name = "frmAutorizacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim blnResult As Boolean

'Etiqueta de prueba por cambio de VSS a la nube

Private Sub cmdAceptar_Click()
    If Me.ActiveControl.Name = "optExcluir" Then
        SendKeys vbTab
        Exit Sub
    End If
    If optExcluir(0).Value Then
        If Trim(txtCodigo.Text) = "" Then
            MsgBox SIHOMsg(929), vbExclamation, "Mensaje"
            txtCodigo.SetFocus
            Exit Sub
        End If
    End If
    blnResult = True
    Me.Hide
End Sub

Private Sub cmdCancelar_Click()
    Me.Hide
End Sub

Private Sub Form_Activate()
    Me.Icon = frmMenuPrincipal.Icon
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
        Me.Hide
    End If
    
End Sub

Public Function fblnAceptarCargoExcluido(ByRef blnExcluido As Boolean, ByRef strCodigo As String, Optional strDescripcionCargo As String = "") As Boolean
    blnResult = False
    Me.lblMsg.Caption = IIf(strDescripcionCargo = "", "", strDescripcionCargo & vbCrLf) & Me.lblMsg.Caption
    optExcluir(1).Value = blnExcluido
    txtCodigo.Text = strCodigo
    Me.Show vbModal
    blnExcluido = optExcluir(1).Value
    strCodigo = txtCodigo.Text
    fblnAceptarCargoExcluido = blnResult
    Unload Me
End Function

Private Sub optExcluir_Click(Index As Integer)
    If Index = 0 Then
        lblCodigo.Caption = "Código de autorización"
    Else
        lblCodigo.Caption = "Notas"
    End If
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

