VERSION 5.00
Begin VB.Form frmEsperaAbranPeriodoContable 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mensaje"
   ClientHeight    =   990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   990
   ScaleWidth      =   4725
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Este periodo contable está cerrado, el sistema está esperando que sea abierto y continuará normalmente, avise a contabilidad."
      Height          =   585
      Left            =   645
      TabIndex        =   0
      Top             =   90
      Width           =   3705
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   585
      Left            =   30
      Picture         =   "frmEsperaAbranPeriodoContable.frx":0000
      Top             =   15
      Width           =   600
   End
End
Attribute VB_Name = "frmEsperaAbranPeriodoContable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------'
' Pantalla que muestra mensaje cuando se intenta guardar una póliza en un
' periodo cerrado, esta pantalla se muestra hasta que el periodo haya sido
' abierto.
' Fecha de programación: 13 de Diciembre del 2000
'--------------------------------------------------------------------------'
Public vlintEjercicio As Integer
Public vlintMes As Integer


Private Sub Form_Activate()
    On Error GoTo NotificaError
    
    Dim X As Long
    frmEsperaAbranPeriodoContable.Refresh
        
    X = 1
    Do While fblnPeriodoCerrado(vgintClaveEmpresaContable, vlintEjercicio, vlintMes)
        X = X + 1
    Loop
    Unload Me


Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Activate"))
End Sub

