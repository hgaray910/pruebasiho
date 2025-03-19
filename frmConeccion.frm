VERSION 5.00
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.1#0"; "MyCommandButton.ocx"
Object = "{FF14BD24-9F8A-41E3-B5B8-7F0D45EE9F16}#15.0#0"; "HSFlatControls.ocx"
Begin VB.Form frmConeccion 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuración de la conexión SIHO"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4995
   Icon            =   "frmConeccion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   4995
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
      Height          =   3735
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   4740
      Begin VB.CheckBox chkMuestra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "&No mostrar cuando se pierda la conexión"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   120
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   3360
         Width           =   4395
      End
      Begin VB.TextBox txtDriver 
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
         Height          =   420
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Text            =   "Text1"
         ToolTipText     =   "Driver utilizado para la base de datos"
         Top             =   1080
         Width           =   2460
      End
      Begin VB.TextBox txtServidor 
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
         Height          =   420
         Left            =   2160
         TabIndex        =   2
         Text            =   "Text2"
         ToolTipText     =   "Servidor de la base de datos"
         Top             =   1530
         Width           =   2460
      End
      Begin VB.TextBox txtUsuario 
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
         Height          =   420
         Left            =   2160
         TabIndex        =   3
         Text            =   "Text3"
         ToolTipText     =   "Usuario de base datos"
         Top             =   1980
         Width           =   2460
      End
      Begin VB.TextBox txtPsw 
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
         Height          =   420
         IMEMode         =   3  'DISABLE
         Left            =   2160
         PasswordChar    =   "•"
         TabIndex        =   4
         ToolTipText     =   "Contraseña del usuario de base de datos"
         Top             =   2430
         Width           =   2460
      End
      Begin VB.TextBox txtBD 
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
         Height          =   420
         Left            =   2160
         TabIndex        =   5
         Text            =   "Text5"
         ToolTipText     =   "Nombre de la base de datos utilizada"
         Top             =   2880
         Width           =   2460
      End
      Begin HSFlatControls.MyCombo cboBd 
         Height          =   375
         Left            =   2160
         TabIndex        =   0
         Top             =   300
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   661
         Style           =   1
         Enabled         =   -1  'True
         Text            =   ""
         Sorted          =   -1  'True
         List            =   $"frmConeccion.frx":030A
         ItemData        =   $"frmConeccion.frx":031D
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
      Begin VB.CheckBox chkSeguridadIntegrada 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Seguridad integrada"
         Enabled         =   0   'False
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
         Left            =   2160
         TabIndex        =   16
         Top             =   720
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000005&
         Caption         =   "Manejador de base de datos"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   2025
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Driver"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1140
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Servidor"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1590
         Width           =   765
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Usuario"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2040
         Width           =   690
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Contraseña"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   2430
         Width           =   1140
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000005&
         Caption         =   "Nombre de base de datos"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   120
         TabIndex        =   9
         Top             =   2820
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
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
      Height          =   840
      Left            =   2160
      TabIndex        =   15
      Top             =   3680
      Width           =   720
      Begin MyCommandButton.MyButton cmdGrabarRegistro 
         Height          =   600
         Left            =   60
         TabIndex        =   7
         ToolTipText     =   "Guardar el registro"
         Top             =   200
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   1058
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         MaskColor       =   16777215
         Picture         =   "frmConeccion.frx":0327
         BackColorOver   =   -2147483633
         BackColorFocus  =   -2147483633
         BackColorDisabled=   -2147483633
         BorderColor     =   -2147483627
         TransparentColor=   16777215
         Caption         =   ""
         DepthEvent      =   1
         PictureDisabled =   "frmConeccion.frx":0CAB
         PictureAlignment=   4
         PictureDisabledEffect=   0
         ShowFocus       =   -1  'True
      End
   End
End
Attribute VB_Name = "frmConeccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public vgblnIndependiente As Boolean
Dim vlblnBand As Boolean
Dim vlstrKey As String

Private Sub cboBd_Click()
  
    If cboBd.Text <> "ORACLE" Then
      txtDriver = "SQLOLEDB.1"
      chkSeguridadIntegrada.Enabled = True
      chkSeguridadIntegrada_Click
    Else
      txtDriver = "ORAOLEDB.ORACLE.1"
      txtBD = ""
      chkSeguridadIntegrada.Enabled = True
      chkSeguridadIntegrada.Value = 1
      chkSeguridadIntegrada.Enabled = False
      txtPsw.Enabled = True
      txtUsuario.Enabled = True
    End If
End Sub


Private Sub chkMuestra_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdGrabarRegistro.SetFocus
    End If
End Sub

Private Sub chkSeguridadIntegrada_Click()
    If chkSeguridadIntegrada.Value Then
        txtPsw.Enabled = False
        txtUsuario.Enabled = False
    Else
        txtPsw.Enabled = True
        txtUsuario.Enabled = True
    End If
End Sub

Private Sub cmdGrabarRegistro_Click()
  If txtDriver <> "" And txtServidor <> "" And (txtUsuario <> "" Or chkSeguridadIntegrada.Value = 1) And (txtPsw <> "" Or chkSeguridadIntegrada.Value = 1) Then
    SaveSetting vlstrKey, "CONEXION", "TYPE", UCase(cboBd.Text)
    vgstrBaseDatosUtilizada = cboBd.Text
    If cboBd.Text = "MSSQL" Then
      SaveSetting vlstrKey & "\CONEXION", "SQL", "DRIVER", txtDriver
      SaveSetting vlstrKey & "\CONEXION", "SQL", "SERVERBD", txtServidor
      SaveSetting vlstrKey & "\CONEXION", "SQL", "USERBD", txtUsuario
      SaveSetting vlstrKey & "\CONEXION", "SQL", "PSWBD", fstrConvierteAsc(fstrEncrypt2(txtPsw, txtUsuario))
      SaveSetting vlstrKey & "\CONEXION", "SQL", "NAMEBD", txtBD
      SaveSetting vlstrKey & "\CONEXION", "SQL", "SHOWSCR", chkMuestra.Value
      SaveSetting vlstrKey & "\CONEXION", "SQL", "SEGINTEGRADA", chkSeguridadIntegrada.Value
    End If
    If cboBd.Text = "ORACLE" Then
      SaveSetting vlstrKey & "\CONEXION", "ORACLE", "DRIVER", txtDriver
      SaveSetting vlstrKey & "\CONEXION", "ORACLE", "SERVERBD", txtServidor
      SaveSetting vlstrKey & "\CONEXION", "ORACLE", "USERBD", txtUsuario
      SaveSetting vlstrKey & "\CONEXION", "ORACLE", "PSWBD", fstrConvierteAsc(fstrEncrypt2(txtPsw, txtUsuario))
      SaveSetting vlstrKey & "\CONEXION", "ORACLE", "NAMEBD", txtBD
      SaveSetting vlstrKey & "\CONEXION", "ORACLE", "SHOWSCR", chkMuestra.Value
    End If
    MsgBox "La información se actualizo satisfactoriamente.", vbInformation, "Mensaje"
    vlblnBand = True
    Unload Me
  Else
    MsgBox "Los datos estan incompletos.", vbExclamation, "Mensaje"
  End If
End Sub

Private Sub Form_Activate()

Dim vlstrBD As String
Dim vlstrAux As String
  
  fblnLimpiaForma Me
  vlstrBD = GetSetting(vlstrKey, "CONEXION", "TYPE")
  If vlstrBD <> "DESCONOCIDO" Then
    On Error Resume Next
    If vlstrBD = "ORACLE" Then
        cboBd.ListIndex = 1
    Else
        cboBd.ListIndex = 0
        chkSeguridadIntegrada.Value = GetSetting(vlstrKey & "\CONEXION", vlstrBD, "SEGINTEGRADA")
    End If
    txtDriver = GetSetting(vlstrKey & "\CONEXION", vlstrBD, "DRIVER")
    txtServidor = GetSetting(vlstrKey & "\CONEXION", vlstrBD, "SERVERBD")
    txtUsuario = GetSetting(vlstrKey & "\CONEXION", vlstrBD, "USERBD")
    txtPsw = UCase(fstrEncrypt2(fstrConvierteChr(GetSetting(vlstrKey & "\CONEXION", vlstrBD, "PSWBD")), txtUsuario))
    txtBD = GetSetting(vlstrKey & "\CONEXION", vlstrBD, "NAMEBD")
    chkMuestra.Value = GetSetting(vlstrKey & "\CONEXION", vlstrBD, "SHOWSCR")
    
  End If
  
  If chkMuestra.Value = 1 And vgblnVieneError Then
    Unload Me
  End If
  
  vlblnBand = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    pFocusNextControl Me, ActiveControl.TabIndex
  Else
    If KeyAscii = vbKeyEscape Then
      Unload Me
    Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
  End If
End Sub

Private Sub Form_Load()
    
    'Me.Icon = frmMenuPrincipal.Icon (se eliminó el icono porque se ejecuta el load del principal y ya no sale esta pantalla)
    
    If vgblnIndependiente Then
        vlstrKey = "MESSENGER"
    Else
        vlstrKey = "SIHO"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not vlblnBand Then
    vgblnTerminate = True
  End If
End Sub
