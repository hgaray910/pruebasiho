VERSION 5.00
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.1#0"; "MyCommandButton.ocx"
Object = "{FF14BD24-9F8A-41E3-B5B8-7F0D45EE9F16}#15.0#0"; "HSFlatControls.ocx"
Begin VB.Form frmSeleccionProveedorSubrogado 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Selección de proveedor de servicios subrogados"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MyCommandButton.MyButton cmdAceptar 
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   2640
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
   Begin HSFlatControls.MyCombo cboProveedores 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   661
      Style           =   1
      Enabled         =   -1  'True
      Text            =   "MyCombo1"
      Sorted          =   0   'False
      List            =   ""
      ItemData        =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MyCommandButton.MyButton cmdCancelar 
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   2640
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
      Caption         =   "Cancelar"
      DepthEvent      =   1
      ShowFocus       =   -1  'True
   End
   Begin VB.Label lblCargo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "[ Descripción del cargo ]"
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
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   4935
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Proveedor"
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
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   1005
   End
   Begin VB.Label lblMsg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Importante. Seleccione el proveedor de servicios subrogados que realizará el siguiente cargo:"
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
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "frmSeleccionProveedorSubrogado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboProveedores_Change()
'Actualizar label

End Sub

Private Sub cboProveedores_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then cmdAceptar.SetFocus
End Sub

Private Sub cmdAceptar_Click()
    If cboProveedores.ItemData(cboProveedores.ListIndex) <> -1 Then
        Me.Hide
    Else
        If MsgBox(SIHOMsg(1217), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
            Me.Hide
        End If
    End If
    
End Sub

Private Sub cmdCancelar_Click()
    cboProveedores.ListIndex = -1
    Me.Hide
End Sub

Private Sub Form_Load()
    Me.Icon = frmMenuPrincipal.Icon
    'pLlenaComboProveedor
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    If Cancel <> 1 Then
'        cboProveedores.ListIndex = 0
'        Cancel = 1
'        'Me.Hide
'    End If
End Sub
