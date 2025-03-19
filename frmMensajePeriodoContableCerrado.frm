VERSION 5.00
Begin VB.Form frmMensajePeriodoContableCerrado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mensaje"
   ClientHeight    =   870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   870
   ScaleWidth      =   3675
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "El periodo contable está cerrado"
      Height          =   195
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   585
      Left            =   120
      Picture         =   "frmMensajePeriodoContableCerrado.frx":0000
      Top             =   120
      Width           =   600
   End
End
Attribute VB_Name = "frmMensajePeriodoContableCerrado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    On Error GoTo NotificaError
    
    Dim X As Long
    
    frmMensajePeriodoContableCerrado.Refresh
    
    For X = 1 To 200000000
    Next X
    
    Unload Me


Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Activate"))

End Sub

