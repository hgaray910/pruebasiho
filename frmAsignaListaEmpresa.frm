VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmAsignaLista 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asignación de listas de precios a empresas y por procedencia"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10950
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   10950
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Listas de precios asignadas"
      Height          =   2565
      Left            =   5145
      TabIndex        =   11
      Top             =   3510
      Width           =   5730
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdAsignadas 
         Height          =   2175
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   5490
         _ExtentX        =   9684
         _ExtentY        =   3836
         _Version        =   393216
         Enabled         =   0   'False
         ScrollBars      =   2
         SelectionMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame freTipo 
      Height          =   690
      Left            =   105
      TabIndex        =   5
      Top             =   75
      Width           =   4935
      Begin VB.OptionButton optProcedencia 
         Caption         =   "Asignación por procedencia"
         Height          =   270
         Left            =   2400
         TabIndex        =   7
         Top             =   270
         Width           =   2340
      End
      Begin VB.OptionButton OptEmpresa 
         Caption         =   "Asignación a empresas"
         Height          =   270
         Left            =   240
         TabIndex        =   6
         Top             =   270
         Value           =   -1  'True
         Width           =   2040
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Listas de precios disponibles"
      Height          =   2685
      Left            =   5145
      TabIndex        =   3
      Top             =   75
      Width           =   5730
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDisponibles 
         Height          =   2295
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   5490
         _ExtentX        =   9684
         _ExtentY        =   4048
         _Version        =   393216
         ScrollBars      =   2
         SelectionMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame1 
      Height          =   700
      Left            =   2262
      TabIndex        =   2
      Top             =   5340
      Width           =   650
      Begin VB.CommandButton cmdGrabar 
         Height          =   480
         Left            =   70
         Picture         =   "frmAsignaListaEmpresa.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Permite grabar la asignación de la empresa"
         Top             =   160
         Width           =   495
      End
   End
   Begin VB.Frame Frame4 
      Height          =   765
      Left            =   7095
      TabIndex        =   13
      Top             =   2740
      Width           =   1830
      Begin VB.CommandButton cmdSelecciona 
         Caption         =   "Incluir"
         Height          =   510
         Index           =   0
         Left            =   120
         MaskColor       =   &H80000014&
         Picture         =   "frmAsignaListaEmpresa.frx":0342
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Incluir un cargo al paquete"
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   780
      End
      Begin VB.CommandButton cmdSelecciona 
         Caption         =   "Excluir"
         Enabled         =   0   'False
         Height          =   510
         Index           =   1
         Left            =   945
         MaskColor       =   &H80000014&
         Picture         =   "frmAsignaListaEmpresa.frx":049C
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Excluir un cargo al paquete"
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   780
      End
   End
   Begin VB.Frame fraTipoPaciente 
      Caption         =   "Tipo de ingreso"
      Height          =   690
      Left            =   120
      TabIndex        =   16
      Top             =   820
      Width           =   4935
      Begin VB.OptionButton optTipoPaciente 
         Caption         =   "Urgencias"
         Height          =   195
         Index           =   3
         Left            =   3720
         TabIndex        =   20
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton optTipoPaciente 
         Caption         =   "Externos"
         Height          =   195
         Index           =   2
         Left            =   2520
         TabIndex        =   19
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton optTipoPaciente 
         Caption         =   "Internos"
         Height          =   195
         Index           =   1
         Left            =   1320
         TabIndex        =   18
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton optTipoPaciente 
         Caption         =   "Todos"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.Frame freEmpresas 
      Caption         =   "Empresas"
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   4950
      Begin VB.ListBox lstEmpresas 
         Height          =   3375
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   1
         ToolTipText     =   "Selección de la empresa para asignarle una lista de precios"
         Top             =   240
         Width           =   4650
      End
   End
   Begin VB.Frame fraProcedencia 
      Caption         =   "Tipos de Paciente"
      Height          =   3735
      Left            =   120
      TabIndex        =   8
      Top             =   1580
      Visible         =   0   'False
      Width           =   4950
      Begin VB.ListBox lstTipoPaciente 
         Height          =   3375
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   9
         ToolTipText     =   "Selección de la empresa para asignarle una lista de precios"
         Top             =   240
         Width           =   4650
      End
   End
End
Attribute VB_Name = "frmAsignaLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------
'| Nombre del Proyecto      : Caja
'| Nombre del Formulario    : frmAsignaListaEmpresa
'-------------------------------------------------------------------------------------
'| Objetivo: Realiza la asignación de la lista de precios a la empresa
'-------------------------------------------------------------------------------------
'| Análisis y Diseño        : Rodolfo Ramos G.
'| Autor                    : Rodolfo Ramos G.
'| Fecha de Creación        : 12/Enero/2001
'| Modificó                 : Nombre(s)
'| Fecha terminación        : 12/Enero/2001
'| Fecha última modificación: dd/mes/AAAA
'-------------------------------------------------------------------------------------
Option Explicit
Dim vgstrEstadoManto As String
Public vllngNumeroOpcion As Long
Dim vgblnDatosModificados As Boolean  '|  Sirve para conocer si los datos fueron modificaron o no, para mandar una alerta
Dim vgintIndexTipoPaciente As Integer '|  Indica el índice que tenía optTipoPaciente antes de cambiar
Dim vgblnEmpresas As Boolean          '|  Indica si esta seleccionado el OptEmpresa

Private Sub cmdSelecciona_Click(Index As Integer)
    On Error GoTo NotificaError
    Dim vlintContador As Integer
    Dim vlblnDepartamentoAsignado As Boolean   'Bandera para saber si esta ya asignado ese departamento
    Dim blnModificado As Boolean
    
    vgblnDatosModificados = True
    If Index = 0 Then 'Selecciona
        vlblnDepartamentoAsignado = True
        For vlintContador = 1 To grdAsignadas.Rows - 1
            If Val(grdAsignadas.TextMatrix(vlintContador, 3)) = grdDisponibles.TextMatrix(grdDisponibles.Row, 3) Then
                vlblnDepartamentoAsignado = False
            End If
        Next
        If vlblnDepartamentoAsignado Then
            pSeleccionaGrid grdDisponibles.Row, grdDisponibles, grdAsignadas, cmdSelecciona(0), cmdSelecciona(1)
        Else
            '|  No se pueden seleccionar dos listas de precios de un mismo departamento.
            MsgBox SIHOMsg(289), vbOKOnly + vbExclamation, "Mensaje"
            vgblnDatosModificados = False
        End If
    Else
        pSeleccionaGrid grdAsignadas.Row, grdAsignadas, grdDisponibles, cmdSelecciona(1), cmdSelecciona(0)
    End If
    grdDisponibles.Col = 2
    grdDisponibles.Sort = 1
    grdDisponibles.Col = 1

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdSelecciona_Click"))
End Sub

Private Sub Form_Activate()
    On Error GoTo NotificaError
    
    Me.Icon = frmMenuPrincipal.Icon
    If grdDisponibles.RowData(1) = -1 Then
        MsgBox "No hay listas de precios, capturadas"
        vgstrEstadoManto = "S"
        Unload Me
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Activate"))
End Sub

Private Sub Form_Load()
    On Error GoTo NotificaError
    Dim rs As New ADODB.Recordset
    Dim vlstrSentencia As String
    
    vgblnDatosModificados = False '|  Inicialmente la información no han sido modificada
    vgstrEstadoManto = "" 'Estatus inicial
    
    '-------------
    'Empresas
    '-------------
    vlstrSentencia = "Select intCveEmpresa as Clave, vchDescripcion as Descripcion " & _
                     "from CCEmpresa where  BITACTIVO = 1  order by Descripcion"
    
    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    
    With lstEmpresas
        .Clear
        Do While Not rs.EOF
            .AddItem rs!Descripcion
            .ItemData(.newIndex) = rs!Clave
            rs.MoveNext
        Loop
    End With
    rs.Close
    
    '-------------
    'Listas de precios
    '-------------
    pLlenaGrid
    lstEmpresas.ListIndex = 0
    
    '-------------
    'Tipos de Paciente
    '-------------
    vlstrSentencia = "Select tnyCveTipoPaciente as Clave, vchDescripcion as Descripcion " & _
                    "from AdTipoPaciente"
    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    
    With lstTipoPaciente
        .Clear
        Do While Not rs.EOF
            .AddItem rs!Descripcion
            .ItemData(.newIndex) = rs!Clave
            rs.MoveNext
        Loop
    End With
    lstTipoPaciente.ListIndex = 0
    vgintIndexTipoPaciente = 0
    vgblnEmpresas = True
    rs.Close

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
End Sub

Private Sub pLlenaGrid()
    On Error GoTo NotificaError
    
    Dim rs As New ADODB.Recordset
    Dim vlstrSentencia As String
    
    pConfiguraGrid
    vlstrSentencia = "SELECT intCveLista, " & _
                     "       chrDescripcion, " & _
                     "       smiCveDepartamento, " & _
                     "       vchDescripcion as Departamento " & _
                     "FROM   pvListaPrecio " & _
                     "       INNER JOIN NoDepartamento ON ( pvListaPrecio.smiDepartamento = NoDepartamento.smiCveDepartamento) " & _
                     "WHERE  bitEstatusActivo = 1 " & _
                     "ORDER     " & _
                     "BY     Departamento, chrDescripcion"
    
    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    
    With grdDisponibles
        Do While Not rs.EOF
            If .RowData(1) <> -1 Then
                .Rows = .Rows + 1
                .Row = .Rows - 1
            End If
            .RowData(.Row) = rs!intcvelista
            .TextMatrix(.Row, 1) = rs!chrDescripcion
            .TextMatrix(.Row, 2) = rs!departamento
            .TextMatrix(.Row, 3) = rs!SMICVEDEPARTAMENTO
            rs.MoveNext
        Loop
        .Row = 1
    End With
    rs.Close

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pllenaGrid"))
End Sub

Private Sub pConfiguraGrid()
    On Error GoTo NotificaError
    
    With grdDisponibles
        .Cols = 4
        .FormatString = "|Lista|Departamento"
        .FixedCols = 1
        .ColWidth(0) = 100   'Fix
        .ColWidth(1) = 3600  'Nombre de la lista 3600
        .ColWidth(2) = 1450  'Nombre Departamento
        .ColWidth(3) = 0     'Cve del Departamento
        .ColAlignment(1) = flexAlignLeftBottom
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignmentFixed(1) = flexAlignCenterCenter
        .ColAlignmentFixed(2) = flexAlignCenterCenter
        .ScrollBars = flexScrollBarBoth
        .TextMatrix(1, 2) = ""
        .RowData(1) = -1
    End With
    With grdAsignadas
        .Cols = 4
        .FormatString = "|Lista|Departamento"
        .FixedCols = 1
        .ColWidth(0) = 100  'Fix
        .ColWidth(1) = 3600 'Nombre de la lista
        .ColWidth(2) = 1450 'Nombre Departamento
        .ColWidth(3) = 0    'Cve Departamento
        .ColAlignment(1) = flexAlignLeftBottom
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignmentFixed(1) = flexAlignCenterCenter
        .ColAlignmentFixed(2) = flexAlignCenterCenter
        .ScrollBars = flexScrollBarBoth
        .TextMatrix(1, 2) = ""
        .RowData(1) = -1
    End With

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pConfiguraGrid"))
End Sub

Private Sub cmdGrabar_Click()
    On Error GoTo NotificaError
    
    Dim rs             As New ADODB.Recordset
    Dim vlstrSentencia As String
    Dim vlintContador  As Integer
    Dim vllngPersonaGraba As Long
    Dim strListaDepartamentos As String '|  Lista de los departamentos que tienen listas de precios que se eliminarán
    Dim strTPEliminar As String '|  Tipos de paciente que se van a eliminar
    
    If fblnRevisaPermiso(vglngNumeroLogin, vllngNumeroOpcion, "E") Then
      ' Persona que graba
      vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
      If vllngPersonaGraba <> 0 Then
        EntornoSIHO.ConeccionSIHO.BeginTrans 'Inicio de la Transacción
        
        '#####################################################################################################################
        '#####################################################################################################################
        '############################################### ASIGNACIÓN POR EMPRESA ################################################
        '#####################################################################################################################
        '#####################################################################################################################
        If OptEmpresa.Value Then
            '|  Si no está vacio el grid
            If grdAsignadas.RowData(1) <> -1 Then
                '|  Tipos de paciente de las listas que serán eliminadas
                strTPEliminar = IIf(optTipoPaciente(0).Value, "'I', 'E', 'U'", "'A'")
                '|  Lista de departamentos de las listas de precios seleccionadas
                strListaDepartamentos = -1
                For vlintContador = 1 To grdAsignadas.Rows - 1
                    strListaDepartamentos = strListaDepartamentos & ", " & grdAsignadas.TextMatrix(vlintContador, 3)
                Next
                '--------------------------------------------------------------------------------------------------------------------------------
                '||  Muestra una advertencia en caso de que se vayan a eliminar listas de precios de otros tipos de pacientes de la empresa
                '--------------------------------------------------------------------------------------------------------------------------------
                vlstrSentencia = "SELECT COUNT(*) co " & _
                                 "FROM   PvListaEmpresa " & _
                                 "WHERE  PvListaEmpresa.intcvelista IN ( SELECT PvListaEmpresa.intcvelista " & _
                                 "                                       FROM   PvListaEmpresa " & _
                                 "                                              INNER JOIN PvListaPrecio ON ( PvListaEmpresa.intcvelista = PvListaPrecio.intcvelista) " & _
                                 "                                       WHERE  PvListaEmpresa.intCveEmpresa = " & lstEmpresas.ItemData(lstEmpresas.ListIndex) & _
                                 "                                              AND PvListaPrecio.smidepartamento in (" & strListaDepartamentos & ") " & _
                                 "                                              AND PvListaEmpresa.chrtipopaciente IN (" & strTPEliminar & ")) " & _
                                 "       And PvListaEmpresa.intCveEmpresa = " & lstEmpresas.ItemData(lstEmpresas.ListIndex)
                If frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)!CO > 0 Then
                    '|  Si existen listas para "I" y/o "E"
                    If optTipoPaciente(0).Value Then
                        '|  El sistema ha detectado al menos una lista asignada por tipo de paciente (Interno, Externo y/o Urgencias),
                        '|  ¿Desea eliminar dicha asignación y establecer una sola lista de precios para todos los tipos de paciente?
                        If MsgBox(SIHOMsg(735), vbYesNo + vbExclamation, "Mensaje") = vbNo Then
                            EntornoSIHO.ConeccionSIHO.RollbackTrans
                            Exit Sub
                        End If
                    Else '|  Existen listas para "A"
                        '|  El sistema ha detectado una asignación para todos los tipos de paciente,
                        '|  ¿Desea eliminar dicha asignación y establecer diferentes listas de precios para los tipos de paciente?
                        If MsgBox(SIHOMsg(736), vbYesNo + vbExclamation, "Mensaje") = vbNo Then
                            EntornoSIHO.ConeccionSIHO.RollbackTrans
                            Exit Sub
                        End If
                    End If
                End If
                '---------------------------------------------------------------------------------------------------------------------------------------------
                '||  Elimina todas las listas de precios que se duplicarían de la empresa seleccionada
                '||  Ej. Si está seleccionado TODOS, se eliminarán las listas para INTERNOS, EXTERNOS y URGENCIAS, de los departamentos de las listas seleccionadas
                '||      Si esta seleccionado INTERNOS, EXTERNOS y/o URGENCIAS, se eliminarán las listas para TODOS, de los departamentos de las listas seleccionadas
                '||  PD. Esto es porque si existe una lista para TODOS, no es correcto que exista ni para INTERNOS, ni para EXTERNOS, ni para URGENCIAS y viceversa
                '---------------------------------------------------------------------------------------------------------------------------------------------
                vlstrSentencia = "DELETE " & _
                                 "FROM   PvListaEmpresa " & _
                                 "WHERE  PvListaEmpresa.intcvelista IN ( SELECT PvListaEmpresa.intcvelista " & _
                                 "                                       FROM   PvListaEmpresa " & _
                                 "                                              INNER JOIN PvListaPrecio ON ( PvListaEmpresa.intcvelista = PvListaPrecio.intcvelista) " & _
                                 "                                       WHERE  PvListaEmpresa.intCveEmpresa = " & lstEmpresas.ItemData(lstEmpresas.ListIndex) & _
                                 "                                              AND PvListaPrecio.smidepartamento in (" & strListaDepartamentos & ") " & _
                                 "                                              AND PvListaEmpresa.chrtipopaciente IN (" & strTPEliminar & ")) " & _
                                 "       And PvListaEmpresa.intCveEmpresa = " & lstEmpresas.ItemData(lstEmpresas.ListIndex)
                pEjecutaSentencia vlstrSentencia
            End If
            '-----------------------------------------------------------------------------------------------------------------------
            '||  Elimina todas las listas de precios del tipo de paciente seleccionado ("I", "E" ó "A") de la empresa seleccionada
            '-----------------------------------------------------------------------------------------------------------------------
            vlstrSentencia = "DELETE " & _
                             "FROM   PvListaEmpresa " & _
                             "WHERE  PvListaEmpresa.intcveempresa = " & lstEmpresas.ItemData(lstEmpresas.ListIndex) & _
                             "       AND PvListaEmpresa.chrtipopaciente = '" & IIf(optTipoPaciente(0).Value, "A", IIf(optTipoPaciente(1).Value, "I", IIf(optTipoPaciente(2).Value, "E", "U"))) & "'"
            pEjecutaSentencia vlstrSentencia
        Else
        
        '#####################################################################################################################
        '#####################################################################################################################
        '############################################# ASIGNACIÓN POR PROCEDENCIA ###############################################
        '#####################################################################################################################
        '#####################################################################################################################
'            vlstrSentencia = "Delete from pvlistaTipoPaciente " & _
'                             "where tnyCveTipoPaciente = " & lstTipoPaciente.ItemData(lstTipoPaciente.ListIndex)
'            pEjecutaSentencia vlstrSentencia

            '|  Si no está vacio el grid
            If grdAsignadas.RowData(1) <> -1 Then
                '|  Tipos de paciente de las listas que serán eliminadas
                strTPEliminar = IIf(optTipoPaciente(0).Value, "'I', 'E', 'U'", "'A'")
                '|  Lista de departamentos de las listas de precios seleccionadas
                strListaDepartamentos = -1
                For vlintContador = 1 To grdAsignadas.Rows - 1
                    strListaDepartamentos = strListaDepartamentos & ", " & grdAsignadas.TextMatrix(vlintContador, 3)
                Next
                '--------------------------------------------------------------------------------------------------------------------------------
                '||  Muestra una advertencia en caso de que se vayan a eliminar listas de precios de otros tipos de pacientes
                '--------------------------------------------------------------------------------------------------------------------------------
                vlstrSentencia = "SELECT COUNT(*) co " & _
                                 "FROM   PvListaTipoPaciente " & _
                                 "WHERE  PvListaTipoPaciente.intcvelista IN ( SELECT PvListaTipoPaciente.intcvelista " & _
                                 "                                       FROM   PvListaTipoPaciente " & _
                                 "                                              INNER JOIN PvListaPrecio ON ( PvListaTipoPaciente.intcvelista = PvListaPrecio.intcvelista) " & _
                                 "                                       WHERE  PvListaTipoPaciente.TNYCVETIPOPACIENTE = " & lstTipoPaciente.ItemData(lstTipoPaciente.ListIndex) & _
                                 "                                              AND PvListaPrecio.smidepartamento in (" & strListaDepartamentos & ") " & _
                                 "                                              AND PvListaTipoPaciente.chrtipopaciente IN (" & strTPEliminar & ")) " & _
                                 "       And PvListaTipoPaciente.TNYCVETIPOPACIENTE = " & lstTipoPaciente.ItemData(lstTipoPaciente.ListIndex)
                If frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)!CO > 0 Then
                    '|  Si existen listas para "I" y/o "E"
                    If optTipoPaciente(0).Value Then
                        '|  El sistema ha detectado al menos una lista asignada por tipo de paciente (Interno, Externo y/o Urgencias),
                        '|  ¿Desea eliminar dicha asignación y establecer una sola lista de precios para todos los tipos de paciente?
                        If MsgBox(SIHOMsg(735), vbYesNo + vbExclamation, "Mensaje") = vbNo Then
                            EntornoSIHO.ConeccionSIHO.RollbackTrans
                            Exit Sub
                        End If
                    Else '|  Existen listas para "A"
                        '|  El sistema ha detectado una asignación para todos los tipos de paciente,
                        '|  ¿Desea eliminar dicha asignación y establecer diferentes listas de precios para los tipos de paciente?
                        If MsgBox(SIHOMsg(736), vbYesNo + vbExclamation, "Mensaje") = vbNo Then
                            EntornoSIHO.ConeccionSIHO.RollbackTrans
                            Exit Sub
                        End If
                    End If
                End If
                '---------------------------------------------------------------------------------------------------------------------------------------------
                '||  Elimina todas las listas de precios que se duplicarían para el tipo de paciente seleccionado
                '||  Ej. Si está seleccionado TODOS, se eliminarán las listas para INTERNOS, EXTERNOS y URGENCIAS, de los departamentos de las listas seleccionadas
                '||      Si esta seleccionado INTERNOS, EXTERNOS y/o URGENCIAS, se eliminarán las listas para TODOS, de los departamentos de las listas seleccionadas
                '||  PD. Esto es porque si existe una lista para TODOS, no es correcto que exista ni para INTERNOS, ni para EXTERNOS, ni para URGENCIAS y viceversa
                '---------------------------------------------------------------------------------------------------------------------------------------------
                vlstrSentencia = "DELETE " & _
                                 "FROM   PvListaTipoPaciente " & _
                                 "WHERE  PvListaTipoPaciente.intcvelista IN ( SELECT PvListaTipoPaciente.intcvelista " & _
                                 "                                       FROM   PvListaTipoPaciente " & _
                                 "                                              INNER JOIN PvListaPrecio ON ( PvListaTipoPaciente.intcvelista = PvListaPrecio.intcvelista) " & _
                                 "                                       WHERE  PvListaTipoPaciente.TNYCVETIPOPACIENTE = " & lstTipoPaciente.ItemData(lstTipoPaciente.ListIndex) & _
                                 "                                              AND PvListaPrecio.smidepartamento in (" & strListaDepartamentos & ") " & _
                                 "                                              AND PvListaTipoPaciente.chrtipopaciente IN (" & strTPEliminar & ")) " & _
                                 "       And PvListaTipoPaciente.TNYCVETIPOPACIENTE = " & lstTipoPaciente.ItemData(lstTipoPaciente.ListIndex)
                pEjecutaSentencia vlstrSentencia
            End If
            '-----------------------------------------------------------------------------------------------------------------------
            '||  Elimina todas las listas de precios del tipo de paciente seleccionado ("I", "E" ó "A") de la empresa seleccionada
            '-----------------------------------------------------------------------------------------------------------------------
            vlstrSentencia = "DELETE " & _
                             "FROM   PvListaTipoPaciente " & _
                             "WHERE  PvListaTipoPaciente.TNYCVETIPOPACIENTE = " & lstTipoPaciente.ItemData(lstTipoPaciente.ListIndex) & _
                             "       AND PvListaTipoPaciente.chrtipopaciente = '" & IIf(optTipoPaciente(0).Value, "A", IIf(optTipoPaciente(1).Value, "I", IIf(optTipoPaciente(2).Value, "E", "U"))) & "'"
            pEjecutaSentencia vlstrSentencia

        End If
        
        If grdAsignadas.RowData(1) <> -1 Then
            For vlintContador = 1 To grdAsignadas.Rows - 1
                
                If OptEmpresa.Value Then
                    vlstrSentencia = "INSERT INTO pvListaEmpresa ( intCveLista, " & _
                                     "                             intCveEmpresa, " & _
                                     "                             chrTipoPaciente) " & _
                                     "                    VALUES ( " & grdAsignadas.RowData(vlintContador) & ", " & _
                                                                   lstEmpresas.ItemData(lstEmpresas.ListIndex) & ", '" & _
                                                                   IIf(optTipoPaciente(0).Value, "A", IIf(optTipoPaciente(1).Value, "I", IIf(optTipoPaciente(2).Value, "E", "U"))) & "')"
                Else
                    vlstrSentencia = "INSERT INTO pvListaTipoPaciente ( intCveLista, " & _
                                     "                                  tnyCveTipoPaciente, " & _
                                     "                                   chrTipoPaciente) " & _
                                     "                         VALUES (" & grdAsignadas.RowData(vlintContador) & "," & _
                                                                        lstTipoPaciente.ItemData(lstTipoPaciente.ListIndex) & ", '" & _
                                                                   IIf(optTipoPaciente(0).Value, "A", IIf(optTipoPaciente(1).Value, "I", IIf(optTipoPaciente(2).Value, "E", "U"))) & "')"
                End If
                pEjecutaSentencia vlstrSentencia
            Next
        End If
        If OptEmpresa.Value Then
            Call pGuardarLogTransaccion(Me.Name, EnmGrabar, vllngPersonaGraba, "ASIGNACION DE LISTAS DE PRECIOS A EMPRESAS", lstEmpresas.ItemData(lstEmpresas.ListIndex))
        Else
            Call pGuardarLogTransaccion(Me.Name, EnmGrabar, vllngPersonaGraba, "ASIGNACION DE LISTAS DE PRECIOS A TIPOS PACIENTE", lstTipoPaciente.ItemData(lstTipoPaciente.ListIndex))
        End If
        EntornoSIHO.ConeccionSIHO.CommitTrans 'Commit de la transacción
        '|  La información se actualizó satisfactoriamente.
        MsgBox SIHOMsg(284), vbInformation, "Mensaje"
        vgblnDatosModificados = False
        
        If OptEmpresa.Value Then
            OptEmpresa.SetFocus
        Else
            optProcedencia.SetFocus
        End If
      End If 'De la seguridad
    End If
Exit Sub

NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdGrabar_Click"))
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
        If KeyAscii = vbKeyEscape Then Unload Me
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyPress"))
End Sub

Private Sub grdAsignadas_DblClick()
    On Error GoTo NotificaError
        cmdSelecciona_Click 1
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdAsignadas_DblClick"))
End Sub

Private Sub grdDisponibles_DblClick()
    On Error GoTo NotificaError
        cmdSelecciona_Click 0
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdDisponibles_DblClick"))
End Sub

Private Sub grdDisponibles_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
        If KeyCode = vbKeyReturn Then cmdSelecciona_Click 0
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdDisponibles_KeyDown"))
End Sub

Private Sub lstEmpresas_Click()
    On Error GoTo NotificaError
        Screen.MousePointer = vbHourglass
        pPonTipoPaciente
        lstEmpresas_KeyUp vbKey0, 0
        Screen.MousePointer = vbNormal
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":lstEmpresas_Click"))
End Sub

Private Sub lstEmpresas_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
        If KeyCode = vbKeyReturn Then
            If grdDisponibles.Enabled Then grdDisponibles.SetFocus
        End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":lstEmpresas_KeyDown"))
End Sub

Private Sub pConsultElementos(vlstrSentencia As String)
    On Error GoTo NotificaError
    
    Dim rs As New ADODB.Recordset
    Dim vlintContador As Integer
    
    'Inicializar grids
    grdAsignadas.Redraw = False
    grdDisponibles.Redraw = False
    cmdSelecciona(0).Enabled = True
    cmdSelecciona(1).Enabled = False
    grdAsignadas.Enabled = False
    grdDisponibles.Enabled = True
    
    grdAsignadas.Clear
    grdAsignadas.Rows = 2
    grdAsignadas.Cols = 0
    grdDisponibles.Clear
    grdDisponibles.Rows = 2
    grdDisponibles.Cols = 0
    pConfiguraGrid
    pLlenaGrid
    
    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    Do While Not rs.EOF
        For vlintContador = 1 To grdDisponibles.Rows - 1
            If grdDisponibles.RowData(vlintContador) = rs!intcvelista Then
                pSeleccionaGrid vlintContador, grdDisponibles, grdAsignadas, cmdSelecciona(0), cmdSelecciona(1)
                Exit For
            End If
        Next
        rs.MoveNext
    Loop
    grdAsignadas.Redraw = True
    grdDisponibles.Redraw = True

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pConsultElementos"))
End Sub

Private Sub lstTipoPaciente_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    Dim vlstrSentencia As String
    
    vlstrSentencia = "Select intCveLista from PvListaTipoPaciente " & _
                     "where tnyCveTipoPaciente = " & lstTipoPaciente.ItemData(lstTipoPaciente.ListIndex) & _
                    "       AND chrtipopaciente = '" & IIf(optTipoPaciente(0).Value, "A", IIf(optTipoPaciente(1).Value, "I", IIf(optTipoPaciente(2).Value, "E", "U"))) & "'"
    pConsultElementos vlstrSentencia

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":lstTipoPaciente_KeyUp"))
End Sub

Private Sub lstEmpresas_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    Dim vlstrSentencia As String
    
    vlstrSentencia = "SELECT intCveLista " & _
                     "FROM   PvListaEmpresa " & _
                     "WHERE  intCveEmpresa = " & lstEmpresas.ItemData(lstEmpresas.ListIndex) & _
                     "       AND chrtipopaciente = '" & IIf(optTipoPaciente(0).Value, "A", IIf(optTipoPaciente(1).Value, "I", IIf(optTipoPaciente(2).Value, "E", "U"))) & "'"
    pConsultElementos vlstrSentencia

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pIniciaCaptura"))
End Sub

Private Sub lstTipoPaciente_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
        If KeyCode = vbKeyReturn Then
            If grdDisponibles.Enabled Then grdDisponibles.SetFocus
        End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":lstEmpresas_KeyDown"))
End Sub

Private Sub lstTipoPaciente_Click()
    On Error GoTo NotificaError
        Screen.MousePointer = vbHourglass
        pPonTipoPaciente
        lstTipoPaciente_KeyUp vbKey0, 0
        Screen.MousePointer = vbNormal
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":lstTipoPaciente_Click"))
End Sub

Private Sub OptEmpresa_Click()
    Dim blnPasa As Boolean
    On Error GoTo NotificaError
    
    blnPasa = False
    '|  Si se dió click sobre el elemento que ya estaba seleccionado
    If vgblnEmpresas Then Exit Sub
    If vgblnDatosModificados Then
        '|  ¿Desea abandonar la operación?
        blnPasa = IIf(MsgBox(SIHOMsg(17), vbInformation + vbYesNo, "Mensaje") = vbYes, True, False)
    Else
        blnPasa = True
    End If
    
    If blnPasa Then
        freEmpresas.Visible = True
        fraProcedencia.Visible = False
        optTipoPaciente(0).Value = True
        lstEmpresas.SetFocus
        lstEmpresas_Click
        vgblnEmpresas = True
        vgblnDatosModificados = False
    Else
        '|  Deja el Option como estaba
        optProcedencia.Value = True
        optProcedencia.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":OptEmpresa_Click"))
End Sub

Private Sub optProcedencia_Click()
    Dim blnPasa As Boolean
    On Error GoTo NotificaError
    
    '|  Si se dió click sobre el elemento que ya estaba seleccionado
    If Not vgblnEmpresas Then Exit Sub
    If vgblnDatosModificados Then
        '|  ¿Desea abandonar la operación?
        blnPasa = IIf(MsgBox(SIHOMsg(17), vbInformation + vbYesNo, "Mensaje") = vbYes, True, False)
    Else
        blnPasa = True
    End If
    
    If blnPasa Then
        freEmpresas.Visible = False
        fraProcedencia.Visible = True
        lstTipoPaciente.SetFocus
        lstTipoPaciente_Click
        vgblnEmpresas = False
        vgblnDatosModificados = False
    Else
        '|  Deja el Option como estaba
        OptEmpresa.Value = True
        OptEmpresa.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":optProcedencia_Click"))
End Sub

Private Sub optTipoPaciente_Click(Index As Integer)
    Dim blnPasa As Boolean
    
On Error GoTo NotificaError
    '|  Si se dió click sobre el elemento que ya estaba seleccionado
    If vgintIndexTipoPaciente = Index Then Exit Sub
    If vgblnDatosModificados Then
        '|  ¿Desea abandonar la operación?
        blnPasa = IIf(MsgBox(SIHOMsg(17), vbInformation + vbYesNo, "Mensaje") = vbYes, True, False)
    Else
        blnPasa = True
    End If
    
    If blnPasa Then
        If OptEmpresa.Value = True Then
            'lstEmpresas.SetFocus
            lstEmpresas_KeyUp vbKey0, 0
            vgblnDatosModificados = False
            vgintIndexTipoPaciente = Index
        Else
            lstTipoPaciente.SetFocus
            lstTipoPaciente_KeyUp vbKey0, 0
            vgblnDatosModificados = False
            vgintIndexTipoPaciente = Index
        End If
    Else
        '|  Deja option como estaba
        optTipoPaciente(vgintIndexTipoPaciente).Value = True
        optTipoPaciente(vgintIndexTipoPaciente).SetFocus
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":optTipoPaciente_Click"))
End Sub

Private Sub pPonTipoPaciente()
    Dim strSentencia As String
    Dim rsTipoPaciente As New ADODB.Recordset
    
    If OptEmpresa.Value = True Then
        strSentencia = " SELECT DISTINCT chrTipoPaciente" & _
                       " FROM   PVLISTAEMPRESA " & _
                       " WHERE  PVLISTAEMPRESA.intcveempresa = " & lstEmpresas.ItemData(lstEmpresas.ListIndex) & _
                       " ORDER BY chrTipoPaciente"
    Else
        strSentencia = " SELECT DISTINCT chrTipoPaciente" & _
                       " FROM   PVLISTATIPOPACIENTE " & _
                       " WHERE  PVLISTATIPOPACIENTE.TNYCVETIPOPACIENTE = " & lstTipoPaciente.ItemData(lstTipoPaciente.ListIndex) & _
                       " ORDER BY chrTipoPaciente"
    End If
    
    Set rsTipoPaciente = frsRegresaRs(strSentencia, adLockOptimistic, adOpenDynamic)
    If rsTipoPaciente.RecordCount > 0 Then
        Select Case rsTipoPaciente!CHRTIPOPACIENTE
            Case "A"
                optTipoPaciente(0).Value = True
            Case "I"
                optTipoPaciente(1).Value = True
            Case "E"
                optTipoPaciente(2).Value = True
            Case "U"
                optTipoPaciente(3).Value = True
        End Select
    Else
        optTipoPaciente(0).Value = True
    End If
End Sub
