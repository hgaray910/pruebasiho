VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmPermisosBusquedaPerfil 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8250
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   8250
   StartUpPosition =   2  'CenterScreen
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdPerfiles 
      Height          =   6480
      Left            =   105
      TabIndex        =   1
      Top             =   360
      Width           =   7980
      _ExtentX        =   14076
      _ExtentY        =   11430
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      FocusRect       =   0
      SelectionMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Búsqueda de perfiles"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   2235
   End
End
Attribute VB_Name = "frmPermisosBusquedaPerfil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------
' Programa para realizar una busqueda de los perfiles
' Fecha de programación: 23 de Septiembre del 2003
'-----------------------------------------------------------------------------
' Ultimas modificaciones:
' Fecha:
' Descripción del cambio
'-----------------------------------------------------------------------------
Option Explicit

Public vllngPerfilSeleccionado As Long
Dim rsPerfiles As New ADODB.Recordset
Dim vlstrsql As String


Private Sub Form_Activate()
    On Error GoTo NotificaError
    
    If rsPerfiles.RecordCount = 0 Then
        '¡No existe información!
        MsgBox SIHOMsg(13), vbOKOnly + vbInformation, "Mensaje"
        Unload Me
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Activate"))

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 27 Then
        Unload Me
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyPress"))

End Sub

Private Sub Form_Load()
    On Error GoTo NotificaError
    
    vllngPerfilSeleccionado = 0
    
    vlstrsql = "select intPerfil as Numero,vchdescripcionPerfil,vchNombreModulo, intPerfil from siPerfiles order by intperfil"
    Set rsPerfiles = frsRegresaRs(vlstrsql)
    If rsPerfiles.RecordCount <> 0 Then
        pLlenarMshFGrdRs grdPerfiles, rsPerfiles, 3
        
        With grdPerfiles
            .FormatString = "|Perfil|Nombre del perfil|Modulo|Nombre del modulo"
            .ColWidth(0) = 100
            .ColWidth(1) = 500 'Perfil
            .ColWidth(2) = 4300 'Nombre del perfil
            .ColWidth(3) = 3000 'Nombre del modulo
            .ColWidth(4) = 0
        End With
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))

End Sub

Private Sub grdPerfiles_DblClick()
    On Error GoTo NotificaError
    
    vllngPerfilSeleccionado = grdPerfiles.RowData(grdPerfiles.Row)
    Unload Me

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdPerfiles_DblClick"))

End Sub

Private Sub grdPerfiles_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        grdPerfiles_DblClick
    End If
End Sub
