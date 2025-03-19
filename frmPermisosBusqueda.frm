VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmPermisosBusqueda 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Búsqueda de usuarios"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9630
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   9630
   StartUpPosition =   2  'CenterScreen
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdUsuarios 
      Height          =   7425
      Left            =   60
      TabIndex        =   0
      Top             =   80
      Width           =   9500
      _ExtentX        =   16748
      _ExtentY        =   13097
      _Version        =   393216
      ForeColor       =   0
      Rows            =   0
      FixedRows       =   0
      ForeColorFixed  =   0
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorUnpopulated=   -2147483643
      GridColor       =   -2147483638
      GridColorFixed  =   -2147483638
      GridColorUnpopulated=   -2147483638
      FocusRect       =   0
      GridLinesFixed  =   1
      GridLinesUnpopulated=   1
      SelectionMode   =   1
      Appearance      =   0
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
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frmPermisosBusqueda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------
' Programa para realizar una busqueda de los usuarios
' Fecha de programación: 1 de Diciembre del 2000
'-----------------------------------------------------------------------------
' Ultimas modificaciones:
' Fecha:
' Descripción del cambio
'-----------------------------------------------------------------------------
Option Explicit

Public vllngLoginSeleccionado As Long
Dim rsLogines As New ADODB.Recordset
Dim vlstrsql As String


Private Sub Form_Activate()
    On Error GoTo NotificaError
    
    If rsLogines.RecordCount = 0 Then
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
    
    Me.Icon = frmMenuPrincipal.Icon
    
    vllngLoginSeleccionado = 0
    
    vlstrsql = "select intNumeroLogin as Numero,vchUsuario,vchNombreDepartamento,dtmFechaInicial,dtmFechaFinal,intNumeroLogin from Login order by vchUsuario"
    Set rsLogines = frsRegresaRs(vlstrsql)
    If rsLogines.RecordCount <> 0 Then
        pLlenarMshFGrdRs grdUsuarios, rsLogines, 5
        
        With grdUsuarios
            .FormatString = "|Clave|Usuario|Nombre departamento|Fecha inicial|Fecha final"
            .ColWidth(0) = 100
            .ColWidth(1) = 800 'Clave
            .ColWidth(2) = 2330 'Usuario
            .ColWidth(3) = 3300 'Nombre departamento
            .ColWidth(4) = 1300 'Fecha inicial
            .ColWidth(5) = 1300 'Fecha final
            .ColWidth(6) = 0
        End With
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))

End Sub

Private Sub grdUsuarios_DblClick()
    On Error GoTo NotificaError
    
    vllngLoginSeleccionado = grdUsuarios.RowData(grdUsuarios.Row)
    Unload Me

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdUsuarios_DblClick"))

End Sub

Private Sub grdUsuarios_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        grdUsuarios_DblClick
    End If
End Sub

