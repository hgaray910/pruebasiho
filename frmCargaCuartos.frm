VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCargaCuartos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cargo automático de cuartos"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11160
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   11160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame freTrabajando 
      Height          =   1335
      Left            =   3720
      TabIndex        =   11
      Top             =   7680
      Visible         =   0   'False
      Width           =   4560
      Begin VB.Label lblTextoTrabajando 
         Caption         =   "Consultando facturas, por favor espere..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   1395
         TabIndex        =   12
         Top             =   345
         Width           =   3090
      End
   End
   Begin VB.Frame freBarra 
      Height          =   1290
      Left            =   1860
      TabIndex        =   9
      Top             =   9030
      Visible         =   0   'False
      Width           =   7680
      Begin MSComctlLib.ProgressBar pgbBarra 
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   600
         Visible         =   0   'False
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblTextoBarra 
         BackColor       =   &H80000002&
         Caption         =   "Cargando cuartos, por favor espere..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   90
         TabIndex        =   10
         Top             =   180
         Width           =   7410
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000002&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   345
         Left            =   30
         Top             =   120
         Width           =   7620
      End
   End
   Begin VB.CheckBox chkSoloNoCargados 
      Caption         =   "Cargar sólo los que no tengan este cargo hoy"
      Height          =   555
      Left            =   420
      TabIndex        =   8
      Top             =   6210
      Value           =   1  'Checked
      Width           =   2040
   End
   Begin VB.Frame Frame2 
      Height          =   990
      Left            =   3000
      TabIndex        =   1
      Top             =   6000
      Width           =   6465
      Begin VB.Frame Frame4 
         Height          =   675
         Left            =   4800
         TabIndex        =   7
         Top             =   195
         Width           =   105
      End
      Begin VB.Frame Frame3 
         Height          =   675
         Left            =   3045
         TabIndex        =   6
         Top             =   210
         Width           =   105
      End
      Begin VB.CommandButton cmdSeleccion 
         Caption         =   "&Invertir selección"
         Height          =   360
         Index           =   4
         Left            =   150
         TabIndex        =   5
         Top             =   555
         Width           =   2685
      End
      Begin VB.CommandButton cmdSeleccion 
         Caption         =   "&Seleccionar / Quitar Selección"
         Height          =   360
         Index           =   0
         Left            =   150
         TabIndex        =   4
         Top             =   195
         Width           =   2685
      End
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "&Cerrar"
         Height          =   465
         Left            =   5025
         TabIndex        =   3
         Top             =   345
         Width           =   1290
      End
      Begin VB.CommandButton cmdCargar 
         Caption         =   "&Aplicar cuartos"
         Height          =   465
         Left            =   3315
         TabIndex        =   2
         Top             =   345
         Width           =   1290
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pacientes internos actualmente"
      Height          =   5850
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10905
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdPacientes 
         Height          =   5460
         Left            =   135
         TabIndex        =   13
         Top             =   270
         Width           =   10650
         _ExtentX        =   18785
         _ExtentY        =   9631
         _Version        =   393216
         GridColor       =   -2147483638
         FocusRect       =   2
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
End
Attribute VB_Name = "frmCargaCuartos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------
'| Nombre del Proyecto      : Caja                                                   -
'| Nombre del Formulario    : frmCargaCuartos                                        -
'-------------------------------------------------------------------------------------
'| Objetivo: Carga los cuartos automaticos
'-------------------------------------------------------------------------------------
'| Análisis y Diseño        : Rodolfo Ramos G.
'| Autor                    : Rodolfo Ramos G.
'| Fecha de Creación        : 08/Ene/2002
'| Modificó                 : Nombre(s)
'| Fecha Terminación        : 09/Ene/2002
'| Fecha última modificación: 07/Feb/2002
'-------------------------------------------------------------------------------------
Option Explicit

Public vllngNumeroOpcion As Long

Private Sub pCargaPacientes()
    Dim vlstrSentencia As String
    Dim rs As New ADODB.Recordset
    grdPacientes.MousePointer = flexHourglass
    
    '-----------------------------
    ' Letrero de "Cargando datos..."
    '-----------------------------
    freTrabajando.Top = 2500
    freTrabajando.Visible = True
    frmCargaCuartos.Refresh
    '-----------------------------
    
    pInicioGridPacientes
    pConfiguraGridPacientes
    vlstrSentencia = "Select ad.numNumCuenta Cuenta, " & _
                " ad.tnyCveTipoPaciente, " & _
                " ad.intCveEmpresa Empresa, " & _
                " Case when TP.bitUtilizaConvenio = 1 then Empre.vchDescripcion else tp.vchDescripcion end as Convenio, " & _
                " ad.vchNumCuarto Cuarto, " & _
                " rtrim(Pac.vchApellidoPaterno)|| ' ' || rtrim(pac.vchApellidoMaterno)|| ' ' || rtrim(pac.vchNombre) Nombre, " & _
                " AdCuarto.intOtroConcepto ConceptoCargo, " & _
                " ADAREA.TNYCVEDEPTO DeptoArea" & _
                " FROM adAdmision ad " & _
                " INNER JOIN adPaciente Pac ON ad.numCvePaciente = Pac.numCvePaciente " & _
                " INNER JOIN adTipoPaciente TP ON Ad.tnyCveTipoPaciente = TP.tnyCveTipoPaciente " & _
                " INNER JOIN adCuarto ON ad.vchNumCuarto = adCuarto.vchNumCuarto " & _
                " INNER JOIN ADAREA ON ADCUARTO.TNYCVEAREA = ADAREA.TNYCVEAREA " & _
                " INNER JOIN NODEPARTAMENTO ON ad.INTCVEDEPARTAMENTO = NODEPARTAMENTO.SMICVEDEPARTAMENTO " & _
                " LEFT OUTER JOIN ccEmpresa empre ON Ad.intCveEmpresa = Empre.intCveEmpresa " & _
                " where ad.chrEstatusAdmision = 'A' " & _
                " and ad.vchNumCuarto is not null and ADAREA.TNYCVEDEPTO is not null " & _
                " and nodepartamento.tnyclaveempresa = " & vgintClaveEmpresaContable & _
                " ORDER BY ad.vchNumCuarto "

    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    grdPacientes.Redraw = False
    Do While Not rs.EOF
        With grdPacientes
            If .RowData(1) <> -1 Then
                 .Rows = .Rows + 1
                 .Row = .Rows - 1
            End If
            
            .RowData(.Row) = rs!Cuenta
            .Col = 0
            .CellFontBold = True
            .TextMatrix(.Row, 0) = ""
            .TextMatrix(.Row, 1) = IIf(IsNull(rs!Cuenta), "", rs!Cuenta)
            .TextMatrix(.Row, 2) = IIf(IsNull(rs!Nombre), "", rs!Nombre)
            .TextMatrix(.Row, 3) = IIf(IsNull(rs!Cuarto), "", rs!Cuarto)
            .TextMatrix(.Row, 4) = IIf(IsNull(rs!convenio), "", rs!convenio)
            .TextMatrix(.Row, 5) = IIf(IsNull(rs!ConceptoCargo), "", rs!ConceptoCargo)
            .TextMatrix(.Row, 6) = IIf(IsNull(rs!DeptoArea), "", rs!DeptoArea)
        End With
        rs.MoveNext
    Loop
    grdPacientes.Row = 1
    grdPacientes.Redraw = True
    freTrabajando.Visible = False
    grdPacientes.MousePointer = flexArrow
    rs.Close
    
End Sub

Private Sub pInicioGridPacientes()
    With grdPacientes
        .Clear
        .ClearStructure
        .Cols = 7
        .Rows = 2
        .RowData(1) = -1
    End With
End Sub

Private Sub pConfiguraGridPacientes()
    With grdPacientes
        .Cols = 7
        .FixedCols = 1
        .FixedRows = 1
        .SelectionMode = flexSelectionByRow
        .FormatString = "|Cuenta|Nombre|Cuarto|Convenio"
        .ColWidth(0) = 200  'Fix
        .ColWidth(1) = 1000 'Cuenta
        .ColWidth(2) = 5000 'Nombre
        .ColWidth(3) = 600  'Cuarto
        .ColWidth(4) = 3450 'Convenio
        .ColWidth(5) = 0    'Clave del Cargo
        .ColWidth(6) = 0    'Depto de area del cuarto
        .ColAlignment(0) = flexAlignCenterBottom
        .ColAlignment(1) = flexAlignCenterCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignCenterCenter
        .ColAlignment(4) = flexAlignLeftCenter
        .ColAlignmentFixed(0) = flexAlignCenterCenter
        .ColAlignmentFixed(1) = flexAlignCenterCenter
        .ColAlignmentFixed(2) = flexAlignCenterCenter
        .ColAlignmentFixed(3) = flexAlignCenterCenter
        .ColAlignmentFixed(4) = flexAlignCenterCenter
        .ScrollBars = flexScrollBarBoth
    End With
End Sub

Private Sub chkIncluyeCargados_Click()
    pCargaPacientes
End Sub
Private Sub cmdCargar_Click()
    Dim vllngContador As Long
    Dim vllngResultado As Long
    Dim vllngPersonaGraba As Long
    Dim vlblnAlmenosuno As Boolean
    
    If fblnRevisaPermiso(vglngNumeroLogin, vllngNumeroOpcion, "E") Then
      ' Persona que graba
      vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
      If vllngPersonaGraba <> 0 Then
        freBarra.Top = 2500
        pgbBarra.Value = 0
        freBarra.Visible = True
        freBarra.Refresh
        vlblnAlmenosuno = False
        For vllngContador = 1 To grdPacientes.Rows - 1
          pgbBarra.Value = vllngContador / grdPacientes.Rows * 100
          If grdPacientes.TextMatrix(vllngContador, 0) = "*" Then
            If fbolNoCargadoHoy(CLng(Val(grdPacientes.TextMatrix(vllngContador, 1))), CLng(Val(grdPacientes.TextMatrix(vllngContador, 5)))) Then
                vllngResultado = 1
                frsEjecuta_SP CStr(CLng(Val(grdPacientes.TextMatrix(vllngContador, 5)))) & "|" & grdPacientes.TextMatrix(vllngContador, 6) & "|D|0|" & CStr(CLng(Val(grdPacientes.TextMatrix(vllngContador, 1)))) & "|I|OC|0|1|" & vllngPersonaGraba & "|0||0|2", "SP_PVUPDCARGOS", True, vllngResultado
                If vllngResultado < 1 Then
                    MsgBox "No se pudo hacer el cargo automático a la cuenta número " & Trim(grdPacientes.TextMatrix(vllngContador, 1)) & " debido a: " & Chr(13) & SIHOMsg(CInt(vllngResultado) * -1), vbExclamation, "Mensaje"
                Else
                  vlblnAlmenosuno = True
                  Call pGuardarLogTransaccion(Me.Name, EnmGrabar, vllngPersonaGraba, "CARGO AUTOMATICO DE CUARTOS", Trim(grdPacientes.TextMatrix(vllngContador, 1)))
                End If
            End If
          End If
        Next
        freBarra.Visible = False
        If vlblnAlmenosuno Then
          MsgBox SIHOMsg(420), vbInformation, "Mensaje"
        Else
          MsgBox SIHOMsg(33), vbExclamation, "Mensaje"
        End If
        cmdCargar.Enabled = False
        pCargaPacientes
      End If
    End If
End Sub
Private Function fbolNoCargadoHoy(vllngMovPaciente As Long, lngCveCargo As Long) As Boolean
    
    Dim vlStrParametroSP As String
    Dim vlIntResultadoSP As Long

    If chkSoloNoCargados.Value = 1 Then
        'Osea hay que validar
        vlIntResultadoSP = 1
        vlStrParametroSP = fstrFechaSQL(fdtmServerFecha) & "|" & Trim(Str(vllngMovPaciente)) & "|" & Trim(Str(lngCveCargo))
        frsEjecuta_SP vlStrParametroSP, "fn_PvSelPuedeCargarCuarto", True, vlIntResultadoSP
        'vlIntResultadoSP = 2 (si se puede hacer el cargo)
        'vlIntResultadoSP = 3 (no se puede hacer el cargo)
        fbolNoCargadoHoy = vlIntResultadoSP = 2
    Else
        'Sin validacion
        fbolNoCargadoHoy = True
    End If

End Function

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdSeleccion_Click(Index As Integer)
    Select Case Index
    Case 0
        pPonQuitaLetra "*", grdPacientes.Row
    Case 4
        pPonQuitaLetra "*", -1
    End Select
    
End Sub

Private Sub pPonQuitaLetra(Caracter As String, vllngCual As Long)
    Dim vlCont As Long
    Dim vllngTemp As Long
    With grdPacientes
        vllngTemp = .Row
        If vllngCual < 1 Then 'Todos o Invertir selección
            If .RowData(1) <> -1 Then ' Que esta vacio
                For vlCont = 1 To .Rows - 1
                    .TextMatrix(vlCont, 0) = IIf(vllngCual = 0, Caracter, IIf(.TextMatrix(vlCont, 0) = Caracter, "", Caracter))
                    .Col = 0
                    .Row = vlCont
                    .CellFontBold = vbBlackness
                Next vlCont
            End If
        Else
            .TextMatrix(vllngCual, 0) = IIf(.TextMatrix(vllngCual, 0) = Caracter, "", Caracter)
            .Col = 0
            .Row = vllngCual
            .CellFontBold = vbBlackness
        End If
        .Row = vllngTemp
    End With
    cmdCargar.Enabled = True
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    
    Me.Icon = frmMenuPrincipal.Icon
    
    cmdCargar.Enabled = False
    pInicioGridPacientes
    pConfiguraGridPacientes
    frmCargaCuartos.Refresh
    pCargaPacientes
End Sub

Private Sub grdPacientes_Click()
    cmdSeleccion_Click 0
End Sub

Private Sub grdPacientes_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cmdSeleccion_Click 0
    End If
End Sub
