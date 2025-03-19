VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmPacientes 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Búsqueda de pacientes"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11325
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   11325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmOpcion 
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
      Height          =   5500
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   11085
      Begin VB.OptionButton optTipo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Paterno"
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
         Height          =   250
         Index           =   0
         Left            =   120
         TabIndex        =   0
         ToolTipText     =   "Búsqueda por apellido paterno"
         Top             =   360
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton optTipo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Materno"
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
         Height          =   250
         Index           =   1
         Left            =   1320
         TabIndex        =   1
         ToolTipText     =   "Búsqueda por apellido materno"
         Top             =   360
         Width           =   1110
      End
      Begin VB.OptionButton optTipo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Nombre"
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
         Height          =   250
         Index           =   2
         Left            =   2520
         TabIndex        =   2
         ToolTipText     =   "Búsqueda por nombre"
         Top             =   360
         Width           =   1090
      End
      Begin VB.OptionButton optTipo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "CURP"
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
         Height          =   250
         Index           =   3
         Left            =   3720
         TabIndex        =   3
         ToolTipText     =   "Búsqueda por CURP"
         Top             =   360
         Width           =   820
      End
      Begin VB.OptionButton optTipo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Expediente"
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
         Height          =   250
         Index           =   4
         Left            =   4680
         TabIndex        =   4
         ToolTipText     =   "Búsqueda por expediente electrónico"
         Top             =   360
         Width           =   1395
      End
      Begin VB.TextBox txtBusqueda 
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
         TabIndex        =   5
         ToolTipText     =   "Criterio de búsqueda"
         Top             =   700
         Width           =   6555
      End
      Begin VB.Timer tmrCargaPacientes 
         Enabled         =   0   'False
         Interval        =   200
         Left            =   7200
         Top             =   1080
      End
      Begin VSFlex7LCtl.VSFlexGrid grdPacientes 
         Height          =   4275
         Left            =   120
         TabIndex        =   6
         Top             =   1110
         Width           =   10845
         _cx             =   19129
         _cy             =   7541
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483638
         GridColorFixed  =   -2147483638
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   0
         Cols            =   0
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   -1  'True
         WordWrap        =   -1  'True
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
   End
End
Attribute VB_Name = "frmPacientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public vglngExpediente As Long
'Indica si la búsqueda de pacientes es de los que tuvieron cuenta como egresados
Public gblnSoloEgresados As Boolean

Private Sub Form_Activate()
On Error GoTo NotificaError

    txtBusqueda.SetFocus
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Activate"))
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        SendKeys vbTab
    ElseIf KeyAscii = vbKeyEscape Then
        vglngExpediente = -1
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()
On Error GoTo NotificaError

    vgstrNombreForm = Me.Name
    
    pConfiguraGrid
    
    vglngExpediente = -1
    
    txtBusqueda.Text = ""

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
    Unload Me
End Sub

Private Sub grdPacientes_DblClick()
On Error GoTo NotificaError
    
    If Trim(grdPacientes.TextMatrix(grdPacientes.Row, 1)) <> "" Then
        vglngExpediente = Val(grdPacientes.TextMatrix(grdPacientes.Row, 1))
        Unload Me
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdPacientes_DblClick"))
End Sub

Private Sub grdPacientes_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then
        grdPacientes_DblClick
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdPacientes_KeyDown"))
End Sub

Private Sub tmrCargaPacientes_Timer()

    Call pCargaPacientes(txtBusqueda.Text)
    tmrCargaPacientes.Enabled = False

End Sub

Private Sub pCargaPacientes(strCadena As String)
On Error GoTo NotificaError
    Dim rsPacientes As New ADODB.Recordset
    Dim lngContador As Long
    
    vgstrParametrosSP = IIf(OptTipo(0).Value, 1, IIf(OptTipo(1).Value, 2, IIf(OptTipo(2).Value, 3, IIf(OptTipo(3).Value, 4, 5)))) & _
                        "|" & Trim(strCadena)
    If gblnSoloEgresados = True Then vgstrParametrosSP = vgstrParametrosSP & "|-1"
    Set rsPacientes = frsEjecuta_SP(vgstrParametrosSP, IIf(gblnSoloEgresados = True, "Sp_GnSelPacientesEgresados", "Sp_AdSelPacientes"))
    If rsPacientes.RecordCount > 0 Then
    
        grdPacientes.Clear
        pConfiguraGrid
    
        With rsPacientes
        
            For lngContador = 1 To .RecordCount
                grdPacientes.Rows = lngContador + 1
                grdPacientes.TextMatrix(lngContador, 1) = !intnumpaciente
                grdPacientes.TextMatrix(lngContador, 2) = !Nombre
                grdPacientes.TextMatrix(lngContador, 3) = UCase(Format(!dtmFechaNacimiento, "DD/MMM/YYYY"))
                grdPacientes.TextMatrix(lngContador, 4) = IIf(IsNull(!RFC), "", !RFC)
                grdPacientes.TextMatrix(lngContador, 5) = IIf(IsNull(!CURP), "", !CURP)
                grdPacientes.TextMatrix(lngContador, 6) = IIf(IsNull(!Domicilio), "", !Domicilio)
                grdPacientes.TextMatrix(lngContador, 7) = IIf(IsNull(!Telefono), "", !Telefono)
                .MoveNext
            Next lngContador
        
        End With
    End If
    rsPacientes.Close
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCargaPacientes"))
    Unload Me
End Sub

Private Sub txtBusqueda_GotFocus()
On Error GoTo NotificaError

    pSelTextBox txtBusqueda
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtBusqueda_GotFocus"))
    Unload Me
End Sub

Private Sub txtBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        If Trim(txtBusqueda.Text) <> "" Then
            If tmrCargaPacientes.Enabled Then
                Call pCargaPacientes(txtBusqueda.Text)
            End If
        End If
        grdPacientes.Col = 1
        grdPacientes.Row = 1
    End If

End Sub

Private Sub txtBusqueda_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    
    If KeyAscii = 39 Then KeyAscii = 0
    tmrCargaPacientes.Enabled = False
    KeyAscii = Asc(UCase(Chr(KeyAscii)))

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtBusqueda_KeyPress"))
    Unload Me
End Sub

Private Sub txtBusqueda_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyDown Then
        grdPacientes.Col = 1
        grdPacientes.Row = 1
        Exit Sub
    End If
    
    If txtBusqueda.Text = "" Then
        grdPacientes.Clear
        pConfiguraGrid
    Else
        tmrCargaPacientes.Enabled = True
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtBusqueda_KeyUp"))
    Unload Me
End Sub

Private Sub pConfiguraGrid()
On Error GoTo NotificaError

    With grdPacientes
        .Cols = 8
        .Rows = 2
        .FixedCols = 1
        .FixedRows = 1
        
        .FormatString = "|Expediente|Nombre|Fecha nacimiento|RFC|CURP|Domicilio|Teléfono"

        .ColWidth(0) = 150
        .ColWidth(1) = 1200
        .ColWidth(2) = 4500
        .ColWidth(3) = 1500
        .ColWidth(4) = 1500
        .ColWidth(5) = 1500
        .ColWidth(6) = 5000
        .ColWidth(7) = 1000
        
    End With

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pConfiguraGrid"))
    Unload Me
End Sub



