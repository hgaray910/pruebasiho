VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmPrecioHorario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parámetros para precios especiales por horario"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11805
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   11805
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   10365
      TabIndex        =   31
      Top             =   6615
      Visible         =   0   'False
      Width           =   1005
      Begin VB.TextBox txtTipo 
         Height          =   315
         Left            =   255
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   240
         Width           =   525
      End
      Begin VB.TextBox txtClave 
         Height          =   315
         Left            =   90
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   210
         Width           =   525
      End
      Begin VB.TextBox txtTipoPaciente 
         Height          =   315
         Left            =   375
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   180
         Width           =   525
      End
      Begin VB.TextBox txtTipoCargo 
         Height          =   315
         Left            =   150
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   240
         Width           =   525
      End
      Begin VB.TextBox txtClaveCargo 
         Height          =   315
         Left            =   405
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   255
         Width           =   525
      End
      Begin VB.TextBox txtDia 
         Height          =   315
         Left            =   90
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   240
         Width           =   525
      End
   End
   Begin VB.ComboBox cboMoveableCargo 
      ForeColor       =   &H80000015&
      Height          =   315
      Left            =   2730
      TabIndex        =   28
      Text            =   "cboMoveableCargo"
      Top             =   6945
      Visible         =   0   'False
      Width           =   1845
   End
   Begin VB.TextBox txtmoveable 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   105
      MaxLength       =   2
      TabIndex        =   27
      Top             =   6630
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox cboMoveable 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000015&
      Height          =   315
      Left            =   75
      TabIndex        =   26
      Text            =   "cboMoveable"
      Top             =   6945
      Visible         =   0   'False
      Width           =   2640
   End
   Begin VB.Frame fraBorrar 
      Height          =   720
      Left            =   5640
      TabIndex        =   22
      Top             =   6570
      Width           =   630
      Begin VB.CommandButton cmdBorrar 
         Enabled         =   0   'False
         Height          =   495
         Left            =   60
         MaskColor       =   &H80000005&
         Picture         =   "frmPrecioHorario.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   165
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdParametros 
      Height          =   3840
      Left            =   75
      TabIndex        =   30
      Top             =   2730
      Width           =   11640
      _ExtentX        =   20532
      _ExtentY        =   6773
      _Version        =   393216
      Cols            =   11
      ForeColorSel    =   -2147483643
      FormatString    =   "|Consecutivo|Tipo|Subtipo|Tipo paciente|Dia semana|Hora inicio|Duracion|Tipo cargo|Descripcion|Porcentaje"
      _NumberOfBands  =   1
      _Band(0).Cols   =   11
   End
   Begin VB.Frame fraFiltros 
      Caption         =   "Mostrar la información con los siguientes filtros "
      Height          =   2595
      Left            =   75
      TabIndex        =   14
      Top             =   75
      Width           =   11640
      Begin VB.CommandButton cmdCargar 
         Caption         =   "Cargar la información"
         Height          =   405
         Left            =   8805
         TabIndex        =   29
         Top             =   2010
         Width           =   1905
      End
      Begin VB.TextBox txtHoraFin 
         Height          =   315
         Left            =   9720
         MaxLength       =   2
         TabIndex        =   12
         Top             =   690
         Width           =   630
      End
      Begin VB.TextBox txtHoraIni 
         Height          =   315
         Left            =   8520
         MaxLength       =   2
         TabIndex        =   11
         Top             =   690
         Width           =   630
      End
      Begin VB.ComboBox cboDia 
         Height          =   315
         Left            =   8520
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   330
         Width           =   2790
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   1470
         Left            =   3720
         TabIndex        =   20
         Top             =   1005
         Width           =   2400
         Begin VB.OptionButton optTipoCargo 
            Caption         =   "Todos"
            Height          =   225
            Index           =   5
            Left            =   0
            TabIndex        =   5
            Top             =   150
            Width           =   1830
         End
         Begin VB.OptionButton optTipoCargo 
            Caption         =   "OC = Otros conceptos"
            Height          =   225
            Index           =   4
            Left            =   0
            TabIndex        =   9
            Top             =   930
            Width           =   2070
         End
         Begin VB.OptionButton optTipoCargo 
            Caption         =   "EX = Exámenes"
            Height          =   225
            Index           =   3
            Left            =   0
            TabIndex        =   8
            Top             =   735
            Width           =   1830
         End
         Begin VB.OptionButton optTipoCargo 
            Caption         =   "GE = Grupo de exámenes"
            Height          =   225
            Index           =   2
            Left            =   0
            TabIndex        =   7
            Top             =   540
            Width           =   2235
         End
         Begin VB.OptionButton optTipoCargo 
            Caption         =   "ES = Estudios"
            Height          =   225
            Index           =   1
            Left            =   0
            TabIndex        =   6
            Top             =   345
            Width           =   1950
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   1125
         Left            =   1545
         TabIndex        =   18
         Top             =   1005
         Width           =   1140
         Begin VB.OptionButton optTipoPaciente 
            Caption         =   "Externos"
            Height          =   225
            Index           =   2
            Left            =   0
            TabIndex        =   4
            Top             =   555
            Width           =   960
         End
         Begin VB.OptionButton optTipoPaciente 
            Caption         =   "Internos"
            Height          =   225
            Index           =   1
            Left            =   0
            TabIndex        =   3
            Top             =   345
            Width           =   870
         End
         Begin VB.OptionButton optTipoPaciente 
            Caption         =   "Todos"
            Height          =   225
            Index           =   0
            Left            =   0
            TabIndex        =   2
            Top             =   150
            Width           =   795
         End
      End
      Begin VB.ComboBox cboSubTipo 
         Height          =   315
         Left            =   1545
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   690
         Width           =   4245
      End
      Begin VB.ComboBox cboTipo 
         Height          =   315
         Left            =   1545
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   330
         Width           =   4245
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "horas"
         Height          =   195
         Left            =   10575
         TabIndex        =   25
         Top             =   750
         Width           =   390
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "a"
         Height          =   195
         Left            =   9390
         TabIndex        =   24
         Top             =   750
         Width           =   90
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Horario de"
         Height          =   195
         Left            =   7425
         TabIndex        =   23
         Top             =   750
         Width           =   735
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Día semana"
         Height          =   195
         Left            =   7440
         TabIndex        =   21
         Top             =   405
         Width           =   870
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo cargo"
         Height          =   195
         Left            =   2790
         TabIndex        =   19
         Top             =   1155
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de paciente"
         Height          =   195
         Left            =   165
         TabIndex        =   17
         Top             =   1155
         Width           =   1200
      End
      Begin VB.Label lblSubtipo 
         AutoSize        =   -1  'True
         Caption         =   "Subtipo"
         Height          =   195
         Left            =   165
         TabIndex        =   16
         Top             =   750
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
         Height          =   195
         Left            =   165
         TabIndex        =   15
         Top             =   390
         Width           =   315
      End
   End
End
Attribute VB_Name = "frmPrecioHorario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------
'| Nombre del Proyecto      : Caja                                                   -
'| Nombre del Formulario    : frmPrecioHorario
'-------------------------------------------------------------------------------------
'| Objetivo: Registrar los parámetros en PvHorarioEmpresa
'-------------------------------------------------------------------------------------
'| Fecha de Creación        : 30/Oct/2002
'| Fecha Terminación        : 04/Nov/2002
'| Modificó                 :
'| Fecha última modificación:
'| Descripción de la modificación:
'-------------------------------------------------------------------------------------

Dim vlstrx As String
Dim rs As New ADODB.Recordset
Dim vllngMarcados As Long

Dim vllngUltimaHora As Long
Dim vllngUltimaDuracion As Long



Private Sub cboDia_Click()
    On Error GoTo NotificaError

    If cboDia.ListIndex = 0 Then
        txtHoraIni.Enabled = False
        txtHoraIni.Enabled = False
        txtHoraFin.Enabled = False
        txtHoraFin.Enabled = False
    Else
        txtHoraIni.Enabled = True
        txtHoraIni.Enabled = True
        txtHoraFin.Enabled = True
        txtHoraFin.Enabled = True
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboDia_Click"))
End Sub

Private Sub cboDia_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError

    If keyacii = 13 Then
        If txtHoraIni.Enabled Then
            txtHoraFin.SetFocus
        Else
            cmdCargar.SetFocus
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboDia_KeyPress"))
End Sub

Private Sub cboMoveable_Click()
    On Error GoTo NotificaError

    If cboMoveable.ListIndex <> -1 Then
        If grdParametros.Col = 2 Then
            cboTipo.ListIndex = cboMoveable.ListIndex + 1
        End If
    End If
    

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboMoveable_Click"))
End Sub

Private Sub cboMoveable_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError

    If KeyAscii = 13 Then
        If cboMoveable.ListIndex <> -1 Then
    
            With grdParametros
                .TextMatrix(.Row, .Col) = cboMoveable.List(cboMoveable.ListIndex)
                ' Tipo
                If .Col = 2 Then
                    txtTipo.Text = IIf(cboMoveable.ListIndex = 0, "EM", IIf(cboMoveable.ListIndex = 1, "TC", "TP"))
                End If
                ' Subtipo
                If .Col = 3 Then
                    txtClave.Text = cboMoveable.ItemData(cboMoveable.ListIndex)
                    cboSubTipo.ListIndex = cboMoveable.ListIndex + 1
                    .ColWidth(3) = 2500
                End If
                ' Tipo paciente
                If .Col = 4 Then
                    txtTipoPaciente.Text = IIf(cboMoveable.ListIndex = 0, "A", IIf(cboMoveable.ListIndex = 1, "I", "E"))
                    OptTipoPaciente(0).Value = True
                End If
                ' Dia
                If .Col = 5 Then
                    txtDia.Text = IIf(cboMoveable.ListIndex = 0, -1, cboMoveable.ListIndex)
                    cboDia.ListIndex = 0
                End If
                If .Col = 8 Then
                    txtTipoCargo.Text = IIf(cboMoveable.ListIndex = 0, "ES", IIf(cboMoveable.ListIndex = 1, "GE", IIf(cboMoveable.ListIndex = 2, "EX", "OC")))
                    .TextMatrix(.Row, 9) = "Todos"
                    txtClaveCargo.Text = -1
                    .Col = .Col + 1
                End If
                
                cboMoveable.Visible = False
                .Col = .Col + 1
                grdParametros_Click
            End With
        Else
            cboMoveable.Visible = False
            
            grdParametros_Click
        End If
    End If
    
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboMoveable_KeyPress"))
End Sub

Private Sub cboMoveableCargo_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError

    Dim vlstrIniciales As String

    If KeyAscii = 13 Then
        If cboMoveableCargo.ListIndex <> -1 Then
            If cboMoveableCargo.ItemData(cboMoveableCargo.ListIndex) = 0 Then
                txtClaveCargo.Text = -1
            Else
                txtClaveCargo.Text = cboMoveableCargo.ItemData(cboMoveableCargo.ListIndex)
            End If
            
            With grdParametros
                .ColWidth(.Col) = 1200
                .TextMatrix(.Row, .Col) = IIf(cboMoveableCargo.ItemData(cboMoveableCargo.ListIndex) = 0, "Todos", cboMoveableCargo.List(cboMoveableCargo.ListIndex))
                cboMoveableCargo.Visible = False
                .Col = .Col + 1
                grdParametros_Click
            End With
        Else
            If Trim(txtTipoCargo.Text) = "AR" Then
                vlstrIniciales = cboMoveableCargo.Text
                                
                vlstrx = "" & _
                "select " & _
                    "IvArticulo.vchNombreComercial Descripcion," & _
                    "IvArticulo.intIDArticulo Clave " & _
                "From IvArticulo " & _
                "Where " & _
                    "vchEstatus='ACTIVO' " & _
                    "and chrCostoGasto <> 'G' " & _
                    "and vchNombreComercial>='" & Trim(vlstrIniciales) & "' " & _
                "Order By " & _
                    "Descripcion "
                
                Set rs = frsRegresaRs(vlstrx, adLockOptimistic, adOpenDynamic)
                
                pLlenarCboRs cboMoveableCargo, rs, 1, 0, 3
                
                cboMoveableCargo.ListIndex = 1
                
                pEnfocaCbo cboMoveableCargo
            Else
                grdParametros_Click
            End If
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboMoveableCargo_KeyPress"))
End Sub

Private Sub cboSubTipo_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError

    If KeyAscii = 13 Then
        If OptTipoPaciente(0).Value Then
            OptTipoPaciente(0).SetFocus
        Else
            If OptTipoPaciente(1).Value Then
                OptTipoPaciente(1).SetFocus
            Else
                 OptTipoPaciente(2).SetFocus
            End If
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboSubTipo_KeyPress"))
End Sub

Private Sub CboTipo_Click()
    On Error GoTo NotificaError

    If cboTipo.ListIndex = 0 Then
        lblSubtipo.Visible = False
    Else
        lblSubtipo.Caption = cboTipo.List(cboTipo.ListIndex)
        lblSubtipo.Visible = True
        
        If cboTipo.ListIndex = 1 Then
            vlstrx = "" & _
            "select " & _
                "Distinct " & _
                "CcEmpresa.vchDescripcion Descripcion," & _
                "CcEmpresa.intCveEmpresa Clave " & _
            "From CcEmpresa "
        Else
            If cboTipo.ListIndex = 2 Then
                vlstrx = "" & _
                "select " & _
                    "Distinct " & _
                    "CcTipoConvenio.vchDescripcion Descripcion," & _
                    "CcTipoConvenio.tnyCveTipoConvenio Clave " & _
                "From CcTipoConvenio "
            Else
                vlstrx = "" & _
                "select " & _
                    "Distinct " & _
                    "AdTipoPaciente.vchDescripcion Descripcion," & _
                    "AdTipoPaciente.tnyCveTipoPaciente Clave " & _
                "From AdTipoPaciente "
            End If
        End If
        
        Set rs = frsRegresaRs(vlstrx)
        If rs.RecordCount <> 0 Then
            pLlenarCboRs cboSubTipo, rs, 1, 0, 3
        End If
'       cboSubTipo.AddItem "<TODOS>", 0
'       cboSubTipo.ItemData(cboSubTipo.NewIndex) = 0
'       cboSubTipo.ListIndex = 0
        
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboTipo_Click"))
End Sub

Private Sub cboTipo_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        cboSubTipo.SetFocus
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboTipo_KeyPress"))
End Sub

Private Sub pIniciaGrid()
    On Error GoTo NotificaError

    With grdParametros
        .Cols = 11
        .Rows = 2
        .FixedCols = 1
    End With
    
    
    pConfiguraGrid "Subtipo"

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pIniciaGrid"))
End Sub

Private Sub pConfiguraGrid(vlstrTituloColumna As String)
    On Error GoTo NotificaError
    
    With grdParametros
        .FormatString = "|Número|Tipo|" & vlstrTituloColumna & "|Tipo paciente|Día|Hora|Duración|Tipo cargo|Descripción|%"
        
        .RowHeightMin = cboMoveable.Height
        .ColWidth(0) = 180
        .ColWidth(1) = 0        'Número
        .ColWidth(2) = 1600     'Tipo
        .ColWidth(3) = 2500     'Subtipo
        .ColWidth(4) = 1100     'Tipo paciente
        .ColWidth(5) = 1100      'Dia
        .ColWidth(6) = 600      'Hora inicio
        .ColWidth(7) = 750      'Duracion
        .ColWidth(8) = 1400     'Tipo cargo
        .ColWidth(9) = 1200     'Descripcion cargo
        .ColWidth(10) = 700     'Porcentaje
    
        .ColAlignment(3) = flexAlignLeftCenter
        .ColAlignment(5) = flexAlignLeftCenter
        .ColAlignment(9) = flexAlignLeftCenter
        
        .ColAlignmentFixed(1) = flexAlignCenterCenter
        .ColAlignmentFixed(2) = flexAlignCenterCenter
        .ColAlignmentFixed(3) = flexAlignCenterCenter
        .ColAlignmentFixed(4) = flexAlignCenterCenter
        .ColAlignmentFixed(5) = flexAlignCenterCenter
        .ColAlignmentFixed(6) = flexAlignCenterCenter
        .ColAlignmentFixed(7) = flexAlignCenterCenter
        .ColAlignmentFixed(8) = flexAlignCenterCenter
        .ColAlignmentFixed(9) = flexAlignCenterCenter
        .ColAlignmentFixed(10) = flexAlignCenterCenter
    End With

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pConfiguraGrid"))
End Sub

Private Sub cmdBorrar_Click()
    On Error GoTo NotificaError

    Dim X As Long
    
    For X = 1 To grdParametros.Rows - 2
        If Trim(grdParametros.TextMatrix(X, 0)) = "*" Then
            vlstrx = "" & _
            "Delete from  PvHorarioEmpresa " & _
            "where intConsecutivo=" & grdParametros.TextMatrix(X, 1)
            
            pEjecutaSentencia vlstrx
            Call pGuardarLogTransaccion(Me.Name, EnmBorrar, vglngNumeroLogin, "PRECIO DE VENTA", grdParametros.TextMatrix(X, 1))
        End If
    Next X
        
    cmdCargar_Click

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdBorrar_Click"))
End Sub

Private Sub cmdCargar_Click()
    On Error GoTo NotificaError
    Dim rsPvSelHorarioEmpresa As New ADODB.Recordset
    Dim vlintSubTipo As Integer
    
    Dim X As Long
    Dim Y As Long
    Dim vlstrTituloColumna As String

    With grdParametros
        .Cols = 11
        .Rows = 2
        .FixedCols = 1
        
        For X = 1 To .Cols - 1
            For Y = 1 To .Rows - 1
                .Col = X
                .Row = Y
                .CellBackColor = &H80000014
                .TextMatrix(Y, X) = ""
            Next Y
        Next X
    End With
    
    If cboSubTipo.ListIndex = -1 Then
       vlintSubTipo = -1
    Else
       vlintSubTipo = IIf(cboSubTipo.ItemData(cboSubTipo.ListIndex) = 0, -1, cboSubTipo.ItemData(cboSubTipo.ListIndex))
    End If

    Me.MousePointer = 11
    vgstrParametrosSP = "" & _
    IIf(cboTipo.ListIndex = 0, "*", IIf(cboTipo.ListIndex = 1, "EM", IIf(cboTipo.ListIndex = 2, "TC", "TP"))) & "|" & vlintSubTipo & "|" & _
    IIf(OptTipoPaciente(0).Value, "*", IIf(OptTipoPaciente(1).Value, "I", "E")) & "|" & _
    IIf(optTipoCargo(5).Value, "*", IIf(optTipoCargo(1).Value, "ES", IIf(optTipoCargo(2).Value, "GE", IIf(optTipoCargo(3).Value, "EX", "OC")))) & "|" & _
    IIf(cboDia.ListIndex = 0, 0, cboDia.ListIndex) & "|" & _
    IIf(cboDia.ListIndex = 0, 1, 0) & "|" & _
    IIf(cboDia.ListIndex = 0, 0, Val(txtHoraIni.Text)) & "|" & _
    IIf(cboDia.ListIndex = 0, 0, Val(txtHoraFin.Text)) & "|" & _
    vgintClaveEmpresaContable
    
    Set rsPvSelHorarioEmpresa = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELHORARIOEMPRESA")
    
    If rsPvSelHorarioEmpresa.RecordCount <> 0 Then
        grdParametros.Redraw = False
        
        pLlenarMshFGrdRs grdParametros, rsPvSelHorarioEmpresa
        pFormatoPorcentaje
        grdParametros.Redraw = True
        grdParametros.Rows = grdParametros.Rows + 1
    
    End If
    rsPvSelHorarioEmpresa.Close
    
    Me.MousePointer = 0
    
    vlstrTituloColumna = IIf(cboTipo.ListIndex = 0, "Todos", cboTipo.List(cboTipo.ListIndex))
    
    pConfiguraGrid vlstrTituloColumna
    
    If grdParametros.Rows > 2 Then
        grdParametros.Redraw = False
        For X = 1 To grdParametros.Cols - 1
            For Y = 1 To grdParametros.Rows - 2
                grdParametros.Col = X
                grdParametros.Row = Y
                grdParametros.CellBackColor = &H80000018
            Next Y
        Next X
        grdParametros.Redraw = True
    End If
    
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdCargar_Click"))
End Sub

Private Sub Form_Activate()
    vgstrNombreForm = Me.Name
End Sub
   Private Sub pFormatoPorcentaje()
    Dim X As Long
       For X = 1 To grdParametros.Rows - 1
        grdParametros.TextMatrix(X, 10) = Format(Val(grdParametros.TextMatrix(X, 10)), "###.00")
    Next X
    End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError

    If KeyAscii = 27 Then
        If Not fraFiltros.Enabled Then
            '¿Desea abandonar la operación?
            If MsgBox(SIHOMsg(17), vbYesNo + vbInformation, "Mensaje") = vbYes Then
                pLimpia
                
                fraFiltros.Enabled = True
                fraBorrar.Enabled = True
                
                cboMoveable.Visible = False
                cboMoveableCargo.Visible = False
                txtmoveable.Visible = False
                
                cmdCargar_Click
            End If
        Else
            Unload Me
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyPress"))
End Sub

Private Sub pLimpia()
    On Error GoTo NotificaError
    
    txtClave.Text = ""
    txtClaveCargo.Text = ""
    txtDia.Text = ""
    txtTipo.Text = ""
    txtTipoCargo.Text = ""
    txtTipoPaciente.Text = ""
    vllngUltimaDuracion = 0
    vllngUltimaHora = 0

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLimpia"))
End Sub

Private Sub Form_Load()
    On Error GoTo NotificaError

    Me.Icon = frmMenuPrincipal.Icon

    vllngMarcados = 0

    cboTipo.AddItem "<TODOS>", 0
    cboTipo.ItemData(cboTipo.NewIndex) = 0
    cboTipo.AddItem "Empresa", 1
    cboTipo.ItemData(cboTipo.NewIndex) = 0
    cboTipo.AddItem "Tipo de convenio", 2
    cboTipo.ItemData(cboTipo.NewIndex) = 0
    cboTipo.AddItem "Tipo de paciente", 3
    cboTipo.ItemData(cboTipo.NewIndex) = 0
    cboTipo.ListIndex = 0
    
    cboSubTipo.AddItem "<TODOS>", 0
    cboSubTipo.ListIndex = 0
    
    cboDia.AddItem "<TODOS>", 0
    cboDia.ItemData(cboDia.NewIndex) = 0
    cboDia.AddItem "Domingo", 1
    cboDia.ItemData(cboDia.NewIndex) = 0
    cboDia.AddItem "Lunes", 2
    cboDia.ItemData(cboDia.NewIndex) = 0
    cboDia.AddItem "Martes", 3
    cboDia.ItemData(cboDia.NewIndex) = 0
    cboDia.AddItem "Miércoles", 4
    cboDia.ItemData(cboDia.NewIndex) = 0
    cboDia.AddItem "Jueves", 5
    cboDia.ItemData(cboDia.NewIndex) = 0
    cboDia.AddItem "Viernes", 6
    cboDia.ItemData(cboDia.NewIndex) = 0
    cboDia.AddItem "Sábado", 7
    cboDia.ItemData(cboDia.NewIndex) = 0
    cboDia.ListIndex = 0
    
    OptTipoPaciente(0).Value = True
    optTipoCargo(5).Value = True
    
    txtHoraIni.Text = ""
    txtHoraFin.Text = ""
    
    pIniciaGrid
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
End Sub


Private Sub grdParametros_Click()
    On Error GoTo NotificaError
    
    With grdParametros

        If Trim(.TextMatrix(.Row, 1)) = "" Then
            fraFiltros.Enabled = False
            fraBorrar.Enabled = False
            
            If .Col = 2 Or .Col = 3 Or .Col = 5 Or .Col = 4 Or .Col = 8 Or .Col = 9 Then
                If .Col = 2 Then                       'Tipo
                    pllenacbo cboTipo, 1
                    'If Trim(.TextMatrix(.Row, .Col)) <> "" Then
                    If Trim(txtTipo.Text) <> "" Then
                        cboMoveable.ListIndex = IIf(Trim(txtTipo.Text) = "EM", 0, IIf(Trim(txtTipo.Text) = "TC", 1, 2))
                    End If
                Else
                    If .Col = 3 Then                    'Subtipo
                        If Trim(.TextMatrix(.Row, 2)) = "" Then
                            .Col = 2
                            grdParametros_Click
                        Else
                            .ColWidth(3) = 4000
                            pllenacbo cboSubTipo, 1
                            'If Trim(.TextMatrix(.Row, .Col)) <> "" Then
                            If Val(txtClave.Text) <> 0 Then
                                pPosiciona Val(txtClave.Text)
                            End If
                        End If
                    Else
                        If .Col = 5 Then                'Dia semana
                            pllenacbo cboDia, 0
                            'If Trim(.TextMatrix(.Row, .Col)) <> "" Then
                            If Val(txtDia.Text) <> 0 Then
                                cboMoveable.ListIndex = IIf(Val(txtDia.Text) = -1, 0, Val(txtDia.Text))
                            End If
                        Else
                            If .Col = 4 Then            'Tipo paciente
                                txtmoveable.Visible = False
                                cboMoveable.Clear
                                cboMoveable.AddItem "Todos", 0
                                cboMoveable.AddItem "Internos", 1
                                cboMoveable.AddItem "Externos", 2
                                cboMoveable.ListIndex = 0
                                cboMoveable.Move .Left + .CellLeft - 15, .Top + .CellTop - 15, .CellWidth
                                cboMoveable.Visible = True
                                cboMoveable.SetFocus
                                'If Trim(.TextMatrix(.Row, .Col)) <> "" Then
                                If Trim(txtTipoPaciente.Text) <> "" Then
                                    cboMoveable.ListIndex = IIf(Trim(txtTipoPaciente.Text) = "A", 0, IIf(Trim(txtTipoPaciente.Text) = "I", 1, 2))
                                End If
                            Else
                                If .Col = 8 Then        'Tipo cargo
                                    txtmoveable.Visible = False
                                    cboMoveable.Clear
                                    'cboMoveable.AddItem "Artículos", 0
                                    cboMoveable.AddItem "Estudios", 0
                                    cboMoveable.AddItem "Grupo de exámenes", 1
                                    cboMoveable.AddItem "Exámenes", 2
                                    cboMoveable.AddItem "Otros conceptos", 3
                                    cboMoveable.ListIndex = 1
                                    cboMoveable.Move .Left + .CellLeft - 15, .Top + .CellTop - 15, .CellWidth
                                    cboMoveable.Visible = True
                                    cboMoveable.SetFocus
                                    If Trim(.TextMatrix(.Row, .Col)) <> "" Then
                                        cboMoveable.ListIndex = IIf(Trim(txtTipoCargo.Text) = "ES", 1, IIf(Trim(txtTipoCargo.Text) = "GE", 2, IIf(Trim(txtTipoCargo.Text) = "EX", 3, 4)))
                                    End If
                                Else
                                                        'Descripcion cargo
                                    If Trim(.TextMatrix(.Row, 8)) = "" Then
                                        .Col = 8
                                        grdParametros_Click
                                    Else
                                        txtmoveable.Visible = False
                                        cboMoveable.Visible = False
                                        
                                        .ColWidth(9) = 4000
                                        cboMoveableCargo.Clear
                                        pllenacboMoveableCargo
                                        cboMoveableCargo.Move .Left + .CellLeft - 15, .Top + .CellTop - 15, .CellWidth
                                        cboMoveableCargo.Visible = True
                                        cboMoveableCargo.SetFocus
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If .Col = 6 Or .Col = 7 Or .Col = 10 Then
                    ' cambia la cantidad de caracteres que le puedes introducir al campo, para el % del campo descuento
                    If .Col = 10 Then
                    txtmoveable.MaxLength = 3
                    Else
                    txtmoveable.MaxLength = 2
                    End If
                    
                    cboMoveable.Visible = False
                    cboMoveableCargo.Visible = False
                    
                    txtmoveable.Move .Left + .CellLeft - 15, .Top + .CellTop - 15, .CellWidth - 8, .CellHeight - 6
                    txtmoveable.Text = .TextMatrix(.Row, .Col)
                    txtmoveable.Visible = True
                    txtmoveable.SetFocus
                    
                    If .Col = 6 Then
                        If vllngUltimaHora <> 0 Then
                            txtmoveable.Text = vllngUltimaHora
                        End If
                    End If
                    If .Col = 7 Then
                        If vllngUltimaDuracion <> 0 Then
                            txtmoveable.Text = vllngUltimaDuracion
                        End If
                    End If
                    
                End If
            End If
        Else
            cboMoveable.Visible = False
            cboMoveableCargo.Visible = False
        End If
    
    End With

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdParametros_Click"))
End Sub

Private Sub pPosiciona(vllngClave As Long)
    On Error GoTo NotificaError
    
    Dim X As Long
    Dim vlblnTermina As Boolean
    
    X = 0
    vlblnTermina = False
    Do While X <= cboMoveable.ListCount - 1 And Not vlblnTermina
        If cboMoveable.ItemData(X) = vllngClave Then
            cboMoveable.ListIndex = X
            vlblnTermina = True
        End If
        X = X + 1
    Loop

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pPosiciona"))
End Sub

Private Sub pllenacboMoveableCargo()
    On Error GoTo NotificaError

    If Trim(txtTipoCargo.Text) = "AR" Then
       cboMoveableCargo.AddItem "<TODOS>", 0
       cboMoveableCargo.ItemData(cboMoveableCargo.NewIndex) = 0
    End If
    If Trim(txtTipoCargo.Text) = "ES" Then
        vlstrx = "" & _
        "select " & _
            "ImEstudio.vchNombre Descripcion," & _
            "ImEstudio.intCveEstudio Clave " & _
        "From ImEstudio " & _
        "Where bitStatusActivo=1 " & _
        "Order By Descripcion "
    End If
    If Trim(txtTipoCargo.Text) = "GE" Then
        vlstrx = "" & _
        "select " & _
            "LaGrupoExamen.chrNombre Descripcion," & _
            "LaGrupoExamen.intCveGrupo Clave " & _
        "From LaGrupoExamen " & _
        "Where bitEstatusActivo=1 " & _
        "Order By Descripcion "
    End If
    If Trim(txtTipoCargo.Text) = "EX" Then
        vlstrx = "" & _
        "select " & _
            "LaExamen.chrNombre Descripcion," & _
            "LaExamen.intCveExamen Clave " & _
        "From LaExamen " & _
        "Where bitEstatusActivo=1 " & _
        "Order By Descripcion "
    End If
    If Trim(txtTipoCargo.Text) = "OC" Then
        vlstrx = "" & _
        "select " & _
            "PvOtroConcepto.chrDescripcion Descripcion," & _
            "PvOtroConcepto.intCveConcepto Clave " & _
        "From PvOtroConcepto " & _
        "Where bitEstatus=1 " & _
        "Order By Descripcion "
    End If
    
    Me.MousePointer = 11
    Set rs = frsRegresaRs(vlstrx)
    Me.MousePointer = 0
    
    If Trim(txtTipoCargo.Text) <> "AR" Then
       pLlenarCboRs cboMoveableCargo, rs, 1, 0, 3
    End If
    
    cboMoveableCargo.ListIndex = 0

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pllenacboMoveableCargo"))
End Sub

Private Sub pllenacbo(cboFuente As ComboBox, vllngPosicion As Long)
    On Error GoTo NotificaError
    
    'vllngPosicion - 1
    
    cboMoveable.Clear
    
    Do While vllngPosicion <= cboFuente.ListCount - 1
        cboMoveable.AddItem cboFuente.List(vllngPosicion), cboMoveable.ListCount
        cboMoveable.ItemData(cboMoveable.NewIndex) = cboFuente.ItemData(vllngPosicion)
        vllngPosicion = vllngPosicion + 1
    Loop
    cboMoveable.ListIndex = 0
    cboMoveable.Move grdParametros.Left + grdParametros.CellLeft - 15, grdParametros.Top + grdParametros.CellTop - 15, grdParametros.CellWidth
    cboMoveable.Visible = True
    cboMoveable.SetFocus
    
    txtmoveable.Visible = False
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pllenacbo"))
End Sub

Private Sub grdParametros_DblClick()
    On Error GoTo NotificaError
    
    Dim vllngColumnaActual As Long

    If Trim(grdParametros.TextMatrix(grdParametros.Row, 1)) <> "" Then
        If Trim(grdParametros.TextMatrix(grdParametros.Row, 0)) = "*" Then
            grdParametros.TextMatrix(grdParametros.Row, 0) = ""
            vllngMarcados = vllngMarcados - 1
        Else
            vllngColumnaActual = grdParametros.Col
            grdParametros.Col = 0
            grdParametros.CellFontBold = True
            grdParametros.TextMatrix(grdParametros.Row, 0) = "*"
            vllngMarcados = vllngMarcados + 1
            grdParametros.Col = vllngColumnaActual
        End If
        If vllngMarcados <> 0 Then
            cmdBorrar.Enabled = True
        Else
            cmdBorrar.Enabled = False
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdParametros_DblClick"))
End Sub

Private Sub grdParametros_Scroll()
    On Error GoTo NotificaError
    
    cboMoveable.Visible = False
    cboMoveableCargo.Visible = False
    txtmoveable.Visible = False

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdParametros_Scroll"))
End Sub

Private Sub optTipoCargo_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        cboDia.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":optTipoCargo_KeyPress"))
End Sub

Private Sub optTipoPaciente_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        If optTipoCargo(1).Value Then
            optTipoCargo(1).SetFocus
        Else
            If optTipoCargo(2).Value Then
                optTipoCargo(2).SetFocus
            Else
                If optTipoCargo(3).Value Then
                    optTipoCargo(3).SetFocus
                Else
                    If optTipoCargo(4).Value Then
                        optTipoCargo(4).SetFocus
                    Else
                        optTipoCargo(5).SetFocus
                    End If
                End If
            End If
        End If
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":optTipoPaciente_KeyPress"))
End Sub



Private Sub txtHoraFin_GotFocus()
    On Error GoTo NotificaError
    
    pSelTextBox txtHoraFin

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtHoraFin_GotFocus"))
End Sub

Private Sub txtHoraFin_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        cmdCargar.SetFocus
    Else
        If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack Then
            KeyAscii = 7
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtHoraFin_KeyPress"))
End Sub

Private Sub txtHoraIni_GotFocus()
    On Error GoTo NotificaError
    
    pSelTextBox txtHoraIni

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtHoraIni_GotFocus"))
End Sub

Private Sub txtHoraIni_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        txtHoraFin.SetFocus
    Else
        If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack Then
            KeyAscii = 7
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtHoraIni_KeyPress"))
End Sub

Private Sub txtmoveable_Change()
If grdParametros.Col = 10 Then
    If Len(txtmoveable.Text) = 3 Then
        If InStr(1, txtmoveable.Text, ".") = 3 Then
        txtmoveable.MaxLength = 5
        ElseIf InStr(1, txtmoveable.Text, ".") = 2 Then
        txtmoveable.MaxLength = 4
        Else
       txtmoveable.MaxLength = 3
         End If
    End If
Else
txtmoveable.MaxLength = 2
End If
End Sub

Private Sub txtmoveable_GotFocus()
    On Error GoTo NotificaError
    
    pSelTextBox txtmoveable

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtmoveable_GotFocus"))
End Sub

    
Private Sub txtmoveable_KeyPress(KeyAscii As Integer)

    On Error GoTo NotificaError
    'Private Sub txtmoveable_KeyPress(KeyAscii As Integer)
    Dim vllngDatoFaltante As Long
    Dim rsPvHorarioEmpresa As New ADODB.Recordset
    
     
    If KeyAscii = 13 Then
        If grdParametros.Col = 6 Then
            'Hora inicio
            If Val(txtmoveable.Text) > 23 Then
                Call Beep
                pEnfocaTextBox txtmoveable
            Else
                grdParametros.TextMatrix(grdParametros.Row, grdParametros.Col) = txtmoveable.Text
                txtHoraIni.Text = "1"
                txtHoraFin.Text = "23"
                
                txtmoveable.Visible = False
                grdParametros.Col = grdParametros.Col + 1
                grdParametros_Click
            End If
        Else
            grdParametros.TextMatrix(grdParametros.Row, grdParametros.Col) = txtmoveable.Text
            txtmoveable.Visible = False
            If grdParametros.Col <> 10 Then
                'Duracion
                grdParametros.Col = grdParametros.Col + 1
                grdParametros_Click
            Else
                vllngDatoFaltante = flngDatoFaltante()
                If vllngDatoFaltante = -1 Then
                    If Val(grdParametros.TextMatrix(grdParametros.Row, 7)) <> 0 Then
                        If Not fblnDuplicado() Then
                            vlstrx = "select * from PvHorarioEmpresa where intConsecutivo=" & Trim(Str(Val(grdParametros.TextMatrix(grdParametros.Row, 1))))
                            Set rsPvHorarioEmpresa = frsRegresaRs(vlstrx, adLockOptimistic, adOpenDynamic)
                            
                            With rsPvHorarioEmpresa
                                If .RecordCount = 0 Then
                                    .AddNew
                                End If
                                !chrTipo = IIf(Trim(txtTipo.Text) = "", " ", Trim(txtTipo.Text))
                                !intcveempresa = IIf(Trim(txtTipo.Text) = "EM", Val(txtClave.Text), 0)
                                !intTipoConvenio = IIf(Trim(txtTipo.Text) = "TC", Val(txtClave.Text), 0)
                                !intTipoPacienteInterno = IIf(Trim(txtTipo.Text) = "TP", Val(txtClave.Text), 0)
                                !chrtipopaciente = IIf(Trim(txtTipoPaciente.Text) = "", " ", Trim(txtTipoPaciente.Text))
                                !chrTipoCargo = IIf(Trim(txtTipoCargo.Text) = "", " ", Trim(txtTipoCargo.Text))
                                !intCveCargo = Val(txtClaveCargo.Text)
                                !intHoraInicio = Val(grdParametros.TextMatrix(grdParametros.Row, 6))
                                !intDuracion = Val(grdParametros.TextMatrix(grdParametros.Row, 7))
                                !intDiaSemana = Val(txtDia.Text)
                                !mnyPorcentaje = Val(grdParametros.TextMatrix(grdParametros.Row, 10))
                                !tnyclaveempresa = vgintClaveEmpresaContable
                                .Update
                            End With
                            
                            rsPvHorarioEmpresa.Close
                            
                            vllngUltimaHora = Val(grdParametros.TextMatrix(grdParametros.Row, 6))
                            vllngUltimaDuracion = Val(grdParametros.TextMatrix(grdParametros.Row, 7))
                                                        
                            'La información se actualizó satisfactoriamente.
                            MsgBox SIHOMsg(284), vbOKOnly + vbInformation, "Mensaje"
                            
                            cmdCargar_Click
                            
                            grdParametros.Row = grdParametros.Rows - 1
                            grdParametros.Col = 2
                            grdParametros_Click
                        Else
                            'Existe información con el mismo contenido
                            MsgBox SIHOMsg(19), vbOKOnly + vbExclamation, "Mensaje"
                            
                            grdParametros.Col = 2
                            grdParametros_Click
                        End If
                    Else
                        grdParametros.Col = 7
                        grdParametros_Click
                    End If
                Else
                    grdParametros.Col = vllngDatoFaltante
                    grdParametros_Click
                End If
            End If
        End If
    Else
        If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack Then
                    
            'KeyAscii = 7
            KeyAscii = fintPDecimal(KeyAscii)
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtmoveable_KeyPress"))
End Sub
Private Function fintPDecimal(tecla As Integer) As Integer ' valida en el campo de % el formato de 3 numeros, un punto y 2 decimales
Dim vlintposicionpunto As Integer
If tecla = 46 Then
        If grdParametros.Col = 10 Then ' si no se tiene enfocado el campo % no realiza nada esta funcion solo devuelve 7(beep)
                If txtmoveable.Text <> "" Then
                    vlintposicionpunto = InStr(1, txtmoveable.Text, ".")
                    If vlintposicionpunto > 0 Then ' si ya hay un punto entonces no se pude poner otro
                         fintPDecimal = 7
                        Exit Function
                    Else ' si no hay punto decimal entonces se coloca y se ajusta el tamaño del campo
                     If Len(txtmoveable.Text) = 6 Or Len(txtmoveable.Text) = 5 Or Len(txtmoveable.Text) = 4 Then
                       fintPDecimal = 7
                       Exit Function
                    Else
                        txtmoveable.MaxLength = Len(txtmoveable.Text) + 3
                        fintPDecimal = 46
                        Exit Function
                   End If
                End If
                Else
                  txtmoveable.MaxLength = Len(txtmoveable.Text) + 3
                  fintPDecimal = 46
                     Exit Function
                End If
        Else
        fintPDecimal = 7
        Exit Function
        End If
Else
fintPDecimal = 7
Exit Function
End If
End Function








Private Function fblnDuplicado() As Boolean
Dim vlintHoraInicio As Integer
Dim vlintHoraFin As Integer
Dim vlintDia As Integer

    On Error GoTo NotificaError

    fblnDuplicado = False
    
    vlintHoraInicio = Val(grdParametros.TextMatrix(grdParametros.Row, 6))
    vlintHoraFin = Val(grdParametros.TextMatrix(grdParametros.Row, 6)) + Val(grdParametros.TextMatrix(grdParametros.Row, 7))
    vlintDia = Val(txtDia.Text)
    
    vlstrx = "" & _
    "select " & _
        "count(*) " & _
    "From " & _
        "PvHorarioEmpresa " & _
    "Where " & _
        "intConsecutivo<>" & Trim(Str(Val(grdParametros.TextMatrix(grdParametros.Row, 1)))) & " " & _
        "and chrTipo='" & Trim(txtTipo.Text) & "' " & _
        "and " & _
            "(" & _
            "intCveEmpresa = " & Trim(txtClave.Text) & " " & _
            "or intTipoConvenio=" & Trim(txtClave.Text) & " " & _
            "or intTipoPacienteInterno=" & Trim(txtClave.Text) & " " & _
            ") " & _
        "and chrTipoPaciente='" & Trim(txtTipoPaciente.Text) & "' " & _
        "and intDiaSemana=" & Trim(txtDia.Text) & " " & _
        "and chrTipoCargo='" & Trim(txtTipoCargo) & "'" & _
        "and intDiaSemana*24+intHoraInicio>= (" & vlintDia & "*24+" & vlintHoraInicio & ")" & _
        "and intDiaSemana*24+intHoraInicio<= (" & vlintDia & "*24+" & vlintHoraFin & ")" & _
        "and intCveCargo = " & txtClaveCargo.Text & _
        "and tnyclaveempresa = " & vgintClaveEmpresaContable
        
    Set rs = frsRegresaRs(vlstrx)
    
    If rs.Fields(0) <> 0 Then
        fblnDuplicado = True
    End If

Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnDuplicado"))
End Function

Private Function flngDatoFaltante() As Long
    On Error GoTo NotificaError
    
    Dim X As Long
    Dim vlblnTermina As Boolean
    
    flngDatoFaltante = -1
    vlblnTermina = False
    
    X = 2
    Do While X <= grdParametros.Cols - 1 And Not vlblnTermina
        If Trim(grdParametros.TextMatrix(grdParametros.Row, X)) = "" Then
            flngDatoFaltante = X
            vlblnTermina = True
        End If
        X = X + 1
    Loop
    
Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":flngDatoFaltante"))
End Function

