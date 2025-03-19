VERSION 5.00
Begin VB.Form frmMensajeAlmacen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mensaje"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6630
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   6630
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      ToolTipText     =   "Aceptar"
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label lblLineaTres 
      AutoSize        =   -1  'True
      Caption         =   "Linea tres"
      Height          =   195
      Left            =   915
      TabIndex        =   3
      Top             =   780
      Width           =   690
   End
   Begin VB.Label lblLineaDos 
      AutoSize        =   -1  'True
      Caption         =   "Linea dos"
      Height          =   195
      Left            =   915
      TabIndex        =   2
      Top             =   480
      Width           =   690
   End
   Begin VB.Label lblLineaUno 
      AutoSize        =   -1  'True
      Caption         =   "Linea uno"
      Height          =   195
      Left            =   915
      TabIndex        =   1
      Top             =   180
      Width           =   705
   End
   Begin VB.Image Image1 
      Height          =   630
      Left            =   90
      Picture         =   "frmMensajeAlmacen.frx":0000
      Top             =   240
      Width           =   600
   End
End
Attribute VB_Name = "frmMensajeAlmacen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAceptar_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    On Error GoTo NotificaError
        
    frmMensajeAlmacen.Refresh
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Activate"))
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()

    Me.Icon = frmMenuPrincipal.Icon

End Sub

