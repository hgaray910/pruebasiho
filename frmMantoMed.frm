VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Begin VB.Form frmMantoNombreMed 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nombre genérico por tipo de paciente/convenio"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13005
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   13005
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraBotones 
      Height          =   2745
      Left            =   6240
      TabIndex        =   22
      Top             =   480
      Width           =   705
      Begin VB.CommandButton cmdDeleteAll 
         Caption         =   "Todos"
         Height          =   615
         Left            =   75
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMantoMed.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Elimina uno"
         Top             =   2050
         UseMaskColor    =   -1  'True
         Width           =   570
      End
      Begin VB.CheckBox chkTodos 
         Caption         =   "Todos"
         Height          =   615
         Left            =   75
         MaskColor       =   &H80000005&
         Picture         =   "frmMantoMed.frx":0482
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Seleccionar todos"
         Top             =   1430
         UseMaskColor    =   -1  'True
         Width           =   570
      End
      Begin VB.CommandButton cmdSelecciona 
         Caption         =   "Excluir"
         Height          =   615
         Index           =   1
         Left            =   75
         MaskColor       =   &H80000014&
         Picture         =   "frmMantoMed.frx":0734
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Excluir de la lista"
         Top             =   780
         UseMaskColor    =   -1  'True
         Width           =   570
      End
      Begin VB.CommandButton cmdSelecciona 
         Caption         =   "Incluir"
         Height          =   615
         Index           =   0
         Left            =   75
         MaskColor       =   &H80000014&
         Picture         =   "frmMantoMed.frx":08AE
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Seleccionar"
         Top             =   150
         UseMaskColor    =   -1  'True
         Width           =   570
      End
   End
   Begin VB.Frame freActualizando 
      Height          =   1335
      Left            =   3345
      TabIndex        =   14
      Top             =   8670
      Visible         =   0   'False
      Width           =   4560
      Begin VB.Label Label10 
         Caption         =   "Actualizando descuentos, por favor espere..."
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
         Left            =   390
         TabIndex        =   15
         Top             =   315
         Width           =   4020
      End
   End
   Begin SysInfoLib.SysInfo SysInfo1 
      Left            =   720
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Frame freBarra 
      Height          =   1275
      Left            =   1620
      TabIndex        =   10
      Top             =   82220
      Visible         =   0   'False
      Width           =   8205
      Begin MSComctlLib.ProgressBar pgbCargando 
         Height          =   360
         Left            =   165
         TabIndex        =   11
         Top             =   675
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblTextoBarra 
         BackColor       =   &H80000002&
         Caption         =   "Cargando datos, por favor espere..."
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
         Height          =   240
         Left            =   105
         TabIndex        =   12
         Top             =   150
         Width           =   7890
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000002&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   285
         Left            =   30
         Top             =   120
         Width           =   8145
      End
   End
   Begin VB.Frame FreConDescuento 
      Caption         =   "Mostrar el nombre genérico (I)=Interno (E)=Externo (A)=Ambos"
      Height          =   2880
      Left            =   7010
      TabIndex        =   13
      Top             =   360
      Width           =   5895
      Begin VB.ListBox lstConDescuento 
         Height          =   2205
         Left            =   60
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   330
         Width           =   5775
      End
   End
   Begin TabDlg.SSTab SSTDescuentos 
      Height          =   3585
      Left            =   0
      TabIndex        =   23
      Top             =   -15
      Width           =   13275
      _ExtentX        =   23416
      _ExtentY        =   6324
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Tipo de Paciente"
      TabPicture(0)   =   "frmMantoMed.frx":0A28
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "freTipoPaciente"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Convenio"
      TabPicture(1)   =   "frmMantoMed.frx":0A44
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FreEmpresa"
      Tab(1).ControlCount=   1
      Begin VB.Frame FreEmpresa 
         Caption         =   "Empresas activas"
         Height          =   2880
         Left            =   -74880
         TabIndex        =   17
         Top             =   360
         Width           =   6060
         Begin VB.ListBox lstEmpresas 
            Height          =   2010
            Left            =   50
            Sorted          =   -1  'True
            TabIndex        =   21
            Top             =   555
            Width           =   4965
         End
         Begin VB.ComboBox cboTipoConvenio 
            Height          =   315
            Left            =   50
            Style           =   2  'Dropdown List
            TabIndex        =   1
            ToolTipText     =   "Tipos de convenio disponibles"
            Top             =   225
            Width           =   4965
         End
         Begin VB.OptionButton OptTipoPac 
            Caption         =   "&Interno"
            Height          =   195
            Index           =   0
            Left            =   5040
            TabIndex        =   20
            ToolTipText     =   "Tipo paciente"
            Top             =   675
            Value           =   -1  'True
            Width           =   825
         End
         Begin VB.OptionButton OptTipoPac 
            Caption         =   "&Externo"
            Height          =   195
            Index           =   1
            Left            =   5040
            TabIndex        =   19
            ToolTipText     =   "Tipo paciente"
            Top             =   975
            Width           =   975
         End
         Begin VB.OptionButton OptTipoPac 
            Caption         =   "&Ambos"
            Height          =   195
            Index           =   2
            Left            =   5040
            TabIndex        =   18
            ToolTipText     =   "Tipo paciente"
            Top             =   1260
            Width           =   855
         End
      End
      Begin VB.Frame freTipoPaciente 
         Caption         =   "Tipos de paciente"
         Height          =   2880
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   6055
         Begin VB.ListBox lstTiposPaciente 
            Height          =   2205
            Left            =   45
            Sorted          =   -1  'True
            TabIndex        =   0
            Top             =   345
            Width           =   4965
         End
         Begin VB.OptionButton optTipoPac2 
            Caption         =   "&Interno"
            Height          =   195
            Index           =   0
            Left            =   5040
            TabIndex        =   2
            Top             =   675
            Value           =   -1  'True
            Width           =   795
         End
         Begin VB.OptionButton optTipoPac2 
            Caption         =   "&Externo"
            Height          =   195
            Index           =   1
            Left            =   5040
            TabIndex        =   3
            Top             =   975
            Width           =   855
         End
         Begin VB.OptionButton optTipoPac2 
            Caption         =   "&Ambos"
            Height          =   195
            Index           =   2
            Left            =   5040
            TabIndex        =   4
            Top             =   1260
            Width           =   855
         End
      End
   End
End
Attribute VB_Name = "frmMantoNombreMed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------
'| Nombre del Proyecto      : Caja
'| Nombre del Formulario    : frmMantoNombreMed
'-------------------------------------------------------------------------------------
'| Objetivo: Realiza el mantenimiento del formato de los nombres de los medicamentos
'-------------------------------------------------------------------------------------
'| Análisis y Diseño        :
'| Autor                    :
'| Fecha de Creación        : 17/Marzo/2009
'| Modificó                 : Nombre(s)
'| Fecha última modificación:

Option Explicit
Dim vgstrEstadoManto As String
Dim vglngDesktop As Long     'Para saber el tamaño del desktop
Const cgintFactorVentana = 200
Dim rsDescuentos As New ADODB.Recordset
    
Private Sub pCargaGuardados(vlstrCual As String)
    On Error GoTo NotificaError

    Dim rs As New ADODB.Recordset
    Dim vlstrSentencia As String
    
    If vlstrCual = "T" Then 'Pacientes
        vlstrSentencia = "SELECT DISTINCT pvgpoNombreGenerico.CHRTIPOGRUPO, " & _
                "pvgpoNombreGenerico.INTCVEAFECTADA, pvgpoNombreGenerico.CHRTIPOPACIENTE, " & _
                "AdTipoPaciente.vchDescripcion as Dato " & _
                "FROM pvgpoNombreGenerico LEFT OUTER JOIN " & _
                "AdTipoPaciente ON " & _
                "pvgpoNombreGenerico.intCveAfectada = AdTipoPaciente.tnyCveTipoPaciente " & _
                "Where CHRTIPOGRUPO = 'T' and pvgpoNombreGenerico.tnyclaveempresa = " & vgintClaveEmpresaContable
   
    ElseIf vlstrCual = "E" Then 'Tipos de Paciente

    vlstrSentencia = "SELECT DISTINCT pvgpoNombreGenerico.CHRTIPOGRUPO, " & _
                "pvgpoNombreGenerico.INTCVEAFECTADA, " & _
                "pvgpoNombreGenerico.CHRTIPOPACIENTE, " & _
                "CcEmpresa.vchDescripcion as Dato " & _
                "FROM pvgpoNombreGenerico LEFT OUTER JOIN " & _
                "CcEmpresa ON " & _
                "pvgpoNombreGenerico.intCveAfectada = CcEmpresa.INTCVEEMPRESA " & _
                "Where CHRTIPOGRUPO = 'E' and pvgpoNombreGenerico.tnyclaveempresa = " & vgintClaveEmpresaContable
             

    End If
    
    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    lstConDescuento.Clear
    Do While Not rs.EOF
            lstConDescuento.AddItem RTrim(rs!Dato) & IIf(rs!chrTipoPaciente = " ", "", "  (" & rs!chrTipoPaciente & ")")
        lstConDescuento.ItemData(lstConDescuento.NewIndex) = rs!intCveAfectada
        rs.MoveNext
    Loop
    
    rs.Close

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCargaGuardados"))
    Unload Me
End Sub

Private Sub pCargaTiposPaciente()
    On Error GoTo NotificaError
    
    Dim vlstrSentencia As String
    Dim rs As New ADODB.Recordset
    
    vlstrSentencia = "Select tnyCveTipoPaciente, vchDescripcion from AdTipoPaciente order by vchDescripcion"
    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    Do While Not rs.EOF
        lstTiposPaciente.AddItem rs!vchDescripcion
        lstTiposPaciente.ItemData(lstTiposPaciente.NewIndex) = rs!tnyCveTipoPaciente
        rs.MoveNext
    Loop
    rs.Close

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCargaTiposPaciente"))
    Unload Me
End Sub

Private Sub pCargaTipoConvenio()
    On Error GoTo NotificaError
    
    Dim vlstrSentencia As String
    Dim rs As New ADODB.Recordset
    
    vlstrSentencia = "Select tnyCveTipoConvenio, vchDescripcion from ccTipoConvenio order by vchDescripcion"
    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    cboTipoConvenio.AddItem "<Todas>"
    cboTipoConvenio.ItemData(cboTipoConvenio.NewIndex) = 0
    Do While Not rs.EOF
        cboTipoConvenio.AddItem rs!vchDescripcion
        cboTipoConvenio.ItemData(cboTipoConvenio.NewIndex) = rs!tnyCveTipoConvenio
        rs.MoveNext
    Loop
    rs.Close
    If cboTipoConvenio.ListCount > 0 Then
        cboTipoConvenio.ListIndex = 0
    Else
        cboTipoConvenio.Enabled = False
        lstEmpresas.Enabled = False
    End If


Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCargaTipoConvenio"))
    Unload Me
End Sub

Private Sub pCargaEmpresas()
    On Error GoTo NotificaError
    
    Dim vlstrSentencia As String
    Dim rs As New ADODB.Recordset
    
    If cboTipoConvenio.ListIndex > 0 Then
        vlstrSentencia = "Select intCveEmpresa, vchDescripcion " & _
                            " from ccEmpresa " & _
                            "Where tnyCveTipoConvenio = " & RTrim(Str(cboTipoConvenio.ItemData(cboTipoConvenio.ListIndex))) & _
                            " order by vchDescripcion "
    Else
        vlstrSentencia = "Select intCveEmpresa, vchDescripcion " & _
                            " from ccEmpresa " & _
                            " order by vchDescripcion "
    End If
    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    lstEmpresas.Clear
    If rs.RecordCount > 0 Then
        Do While Not rs.EOF
            lstEmpresas.AddItem rs!vchDescripcion
            lstEmpresas.ItemData(lstEmpresas.NewIndex) = rs!intcveempresa
            rs.MoveNext
        Loop
        lstEmpresas.Enabled = True
        
    Else
        lstEmpresas.Enabled = False
    End If
    rs.Close

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCargaEmpresas"))
    Unload Me
End Sub

Private Sub cbotipoconvenio_Click()
    On Error GoTo NotificaError

    pCargaEmpresas

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboTipoConvenio_Click"))
    Unload Me
End Sub

Private Sub cboTipoConvenio_KeyDown(KeyCode As Integer, shift As Integer)
  If KeyCode = vbKeyReturn Then
    lstEmpresas.SetFocus
       
    End If
End Sub

Private Sub chkTodos_KeyDown(KeyCode As Integer, shift As Integer)
 If KeyCode = vbKeyReturn Then
   chkTodos.Value = 1
   chkTodos_Click
   lstConDescuento.SetFocus
       
    End If
End Sub

Private Sub cmdDeleteAll_Click()
      
    Dim vlintContador As Integer
    Dim vlstrSentencia As String
    Dim vlstrTipoGrupo As String
    Dim vlstrCadena As String
    Dim vllngCveAfectada As Long
    Dim vlstrTipoPaciente As String
    Dim vllngPersonaGraba As Long
    Dim vlblYa As Boolean
    Dim SQL As String
                          Select Case SSTDescuentos.Tab
                            Case 0 'Tipos de pacientes
                                vlstrTipoGrupo = "T"
                               
                               
                            Case 1 'Empresa
                                vlstrTipoGrupo = "E"
                                
                                
                           End Select
                           
                            If lstConDescuento.ListCount > 0 Then
                             
                            Else
                            Exit Sub
                            End If
                            
                          EntornoSIHO.ConeccionSIHO.BeginTrans
                                    vlstrSentencia = "Delete from pvgpoNombreGenerico " & _
                                     "where chrTipoGrupo = " & "'" & vlstrTipoGrupo & "'" & _
                                     " and tnyclaveempresa = " & vgintClaveEmpresaContable
                                     
                                    pEjecutaSentencia vlstrSentencia
                                  '--------------------------------------------------------------
                   
                        Call pGuardarLogTransaccion(Me.Name, EnmCambiar, vllngPersonaGraba, "ASIGNACION DE NOMBRES DE MEDICAMENTOS", vlstrTipoGrupo & " " & CStr(vllngCveAfectada) & " " & vlstrTipoPaciente)
                          EntornoSIHO.ConeccionSIHO.CommitTrans

                     If SSTDescuentos.Tab = 0 Then
                        frmMantoNombreMed.Refresh
                      End If
                      pCancelar 'Solamente para limpiar la pantalla
                      ' Cargar la lista  de los guardados
                      If SSTDescuentos.Tab = 0 Then
                            pCargaGuardados "T"
                      ElseIf SSTDescuentos.Tab = 1 Then
                           pCargaGuardados "E"
                      End If
      

End Sub
Private Sub cmdSelecciona_Click(Index As Integer)
    On Error GoTo NotificaError
    
    Dim vlintContador As Integer
    Dim vlstrSentencia As String
    Dim vlstrTipoGrupo As String
    Dim vlstrCadena As String
    Dim vllngCveAfectada As Long
    Dim vlstrTipoPaciente As String
    Dim vllngPersonaGraba As Long
    Dim vlblYa As Boolean
    Dim SQL As String
     
    If Index = 0 Then
                   
                    Set rsDescuentos = frsRegresaRs("SELECT * FROM pvgpoNombreGenerico where tnyclaveempresa = " & vgintClaveEmpresaContable, adLockOptimistic, adOpenDynamic)
                
                    With rsDescuentos
                        If .State = 0 Then
                            .Open
                        End If
                                                                    
                      
                        Select Case SSTDescuentos.Tab
                            Case 0 'Tipos de pacientes
                                vlstrTipoGrupo = "T"
                                If lstTiposPaciente.ListIndex > -1 Then
                                    vllngCveAfectada = lstTiposPaciente.ItemData(lstTiposPaciente.ListIndex)
                                Else
                                    Exit Sub
                                End If
                                vlstrTipoPaciente = IIf(optTipoPac2(0).Value, "I", IIf(optTipoPac2(1).Value, "E", "A"))
                            Case 1 'Empresa
                                vlstrTipoGrupo = "E"
                                If lstEmpresas.ListIndex > -1 Then
                                    vllngCveAfectada = lstEmpresas.ItemData(lstEmpresas.ListIndex)
                                Else
                                    Exit Sub
                                End If
                                vlstrTipoPaciente = IIf(OptTipoPac(0).Value, "I", IIf(OptTipoPac(1).Value, "E", "A"))
                        End Select
                
                        ' validar que no exista ya esa informacion en la bd
                        
                        Do While Not .EOF
                           If (!chrTipoGrupo = vlstrTipoGrupo And !intCveAfectada = vllngCveAfectada And !chrTipoPaciente = vlstrTipoPaciente And !tnyclaveempresa = vgintClaveEmpresaContable) Then
                             vlblYa = True
                           End If
                           .MoveNext
                        Loop
                        
                         If Not (vlblYa) Then
                         
                         
'                                '--------------------------------------------------------
'                                ' Persona que graba
'                                '--------------------------------------------------------
'                                vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
'                                If vllngPersonaGraba = 0 Then Exit Sub
'                                '--------------------------------------------------------
                         
                         
                               EntornoSIHO.ConeccionSIHO.BeginTrans
                                 .AddNew
                                        !chrTipoGrupo = vlstrTipoGrupo
                                        !intCveAfectada = vllngCveAfectada
                                        !chrTipoPaciente = vlstrTipoPaciente
                        
                                         !tnyclaveempresa = vgintClaveEmpresaContable
                                        .Update
                        
                                Call pGuardarLogTransaccion(Me.Name, EnmCambiar, vllngPersonaGraba, "ASIGNACION DE NOMBRES DE MEDICAMENTOS", vlstrTipoGrupo & " " & CStr(vllngCveAfectada) & " " & vlstrTipoPaciente)
                                  EntornoSIHO.ConeccionSIHO.CommitTrans
                                  
                               Else
                               '  MsgBox ("El registro ya tiene asignado el parámetro de nombre genérico"), vbOKOnly, "Aviso"
                              
                          End If
                                                  
                   End With
                    
                      If SSTDescuentos.Tab = 0 Then
                        frmMantoNombreMed.Refresh
                      End If
                      pCancelar 'Solamente para limpiar la pantalla
                      ' Cargar la lista  de los guardados
                      If SSTDescuentos.Tab = 0 Then
                            pCargaGuardados "T"
                      ElseIf SSTDescuentos.Tab = 1 Then
                           pCargaGuardados "E"
                      End If
    ElseIf Index = 1 Then
                        
                          Select Case SSTDescuentos.Tab
                            Case 0 'Tipos de pacientes
                                vlstrTipoGrupo = "T"
                            Case 1 'Empresa
                                vlstrTipoGrupo = "E"
                           End Select
                           
                            If lstConDescuento.ListIndex > -1 Then
                             vllngCveAfectada = lstConDescuento.ItemData(lstConDescuento.ListIndex)
                            Else
                            Exit Sub
                            End If
                            
                        
                        If (InStr(1, lstConDescuento.Text, "(E)", vbBinaryCompare)) Then
                         vlstrTipoPaciente = "E"
                         ElseIf (InStr(1, lstConDescuento.Text, "(I)", vbBinaryCompare)) Then
                          vlstrTipoPaciente = "I"
                         Else
                          vlstrTipoPaciente = "A"
                          
                         End If
                        
                                  
'                                    '--------------------------------------------------------
'                                    ' Persona que graba
'                                    '--------------------------------------------------------
'                                    vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
'                                    If vllngPersonaGraba = 0 Then Exit Sub
'                                    '--------------------------------------------------------
'                                    '--------------------------------------------------------------
                                 ' Borrado del elemento  guardado
                          EntornoSIHO.ConeccionSIHO.BeginTrans
                                    vlstrSentencia = "Delete from pvgpoNombreGenerico " & _
                                     "where chrTipoGrupo = " & "'" & vlstrTipoGrupo & "'" & _
                                     " and intCveAfectada = " & vllngCveAfectada & _
                                     " and chrTipoPaciente = " & "'" & vlstrTipoPaciente & "'" & _
                                     " and tnyclaveempresa = " & vgintClaveEmpresaContable
                                     
                                    pEjecutaSentencia vlstrSentencia
                                  '--------------------------------------------------------------
                   
                        Call pGuardarLogTransaccion(Me.Name, EnmCambiar, vllngPersonaGraba, "ASIGNACION DE NOMBRES DE MEDICAMENTOS", vlstrTipoGrupo & " " & CStr(vllngCveAfectada) & " " & vlstrTipoPaciente)
                          EntornoSIHO.ConeccionSIHO.CommitTrans

                     If SSTDescuentos.Tab = 0 Then
                        frmMantoNombreMed.Refresh
                      End If
                      pCancelar 'Solamente para limpiar la pantalla
                      ' Cargar la lista  de los guardados
                      If SSTDescuentos.Tab = 0 Then
                            pCargaGuardados "T"
                      ElseIf SSTDescuentos.Tab = 1 Then
                           pCargaGuardados "E"
                      End If
      
  End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdSelecciona_Click"))
    Unload Me
End Sub

Private Function FintBuscaEnRowData(grdHBusca As MSHFlexGrid, vlintCriterio As Long, vlstrTipoElemento)
    On Error GoTo NotificaError
    
    Dim vlintContador As Long
    
    FintBuscaEnRowData = -1
    With grdHBusca
    For vlintContador = 1 To .Rows - 1
        If .RowData(vlintContador) = vlintCriterio And vlstrTipoElemento = .TextMatrix(vlintContador, 3) Then
            FintBuscaEnRowData = vlintContador
            Exit For
        End If
    Next
    End With

Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":FintBuscaEnRowData"))
    Unload Me
End Function

Private Sub cmdSelecciona_KeyDown(Index As Integer, KeyCode As Integer, shift As Integer)
 If KeyCode = vbKeyReturn Then
   lstConDescuento.SetFocus
       
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo NotificaError
    
    Select Case vgstrEstadoManto
        Case "A"
            Cancel = True
            If MsgBox(SIHOMsg(17), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
                pCancelar
            End If
        Case "AS", "MS"
            Cancel = True
            chkTodos.Value = 0

    End Select

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_QueryUnload"))
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo NotificaError

    If rsDescuentos.State = 1 Then
        rsDescuentos.Close
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Unload"))
    Unload Me
End Sub
'Private Sub pPideDescuento()
'    On Error GoTo NotificaError
'
'    'La opción del descuento por costo aplica sólo para Artículos
' '   optTipoDescuento(2).Enabled = sstElementos.Tab = 0 Or sstElementos.Tab = 5
'    '-------------------------------------------------------------
'
''    freDescuento.Visible = True
''    FreElementos.Enabled = False
''    freElementosIncuidos.Enabled = False
'    cmdGrabarRegistro.Enabled = False
''    cmdVerDescuentos.Enabled = False
''    pEnfocaTextBox txtDescuento
'
'Exit Sub
'NotificaError:
'    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pPideDescuento"))
'    Unload Me
'End Sub

Private Sub chkTodos_Click()
    On Error GoTo NotificaError
    Dim vlintContador As Integer
    Dim vlstrSentencia As String
    Dim vlstrTipoGrupo As String
    Dim vlstrCadena As String
    Dim vllngCveAfectada As Long
    Dim vlstrTipoPaciente As String
    Dim vllngPersonaGraba As Long
    Dim vlblYa As Boolean
    Dim SQL As String
    
    If chkTodos.Value = 1 Then
   
   
'       '--------------------------------------------------------
'       ' Persona que graba
'       '--------------------------------------------------------
'                  vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
'                  If vllngPersonaGraba = 0 Then Exit Sub
'       '--------------------------------------------------------
'
                    
                    Set rsDescuentos = frsRegresaRs("SELECT * FROM pvgpoNombreGenerico where tnyclaveempresa = " & vgintClaveEmpresaContable, adLockOptimistic, adOpenDynamic)
                
                    With rsDescuentos
                        If .State = 0 Then
                            .Open
                        End If
                                                                    
                      
                     Select Case SSTDescuentos.Tab
                        Case 0 'Tipos de pacientes
                                vlstrTipoPaciente = IIf(optTipoPac2(0).Value, "I", IIf(optTipoPac2(1).Value, "E", "A"))
                                vlstrTipoGrupo = "T"
                                If lstTiposPaciente.ListCount > 0 Then
                                
                                Else
                                    Exit Sub
                                End If
                            lstTiposPaciente.ListIndex = 0
                            
                              For vlintContador = 0 To lstTiposPaciente.ListCount - 1
                                vllngCveAfectada = lstTiposPaciente.ItemData(vlintContador)
                                 If .RecordCount > 0 Then
                                 ' Que no exista esa informacion en la BD
                                 '-------------------------------------------------------------------------------------------------
                                       .MoveFirst
                                        Do While Not .EOF And vlblYa = False
                                                    If (!chrTipoGrupo = vlstrTipoGrupo And !intCveAfectada = vllngCveAfectada And !chrTipoPaciente = vlstrTipoPaciente And !tnyclaveempresa = vgintClaveEmpresaContable) Then
                                                      vlblYa = True
                                                    End If
                                                    .MoveNext
                                         Loop
                                '-------------------------------------------------------------------------------------------------
                                End If
                                 If Not (vlblYa) Then
                                    EntornoSIHO.ConeccionSIHO.BeginTrans
                                        .AddNew
                                        !chrTipoGrupo = vlstrTipoGrupo
                                        !intCveAfectada = vllngCveAfectada
                                        !chrTipoPaciente = vlstrTipoPaciente
                                        !tnyclaveempresa = vgintClaveEmpresaContable
                                       .Update
                                      Call pGuardarLogTransaccion(Me.Name, EnmCambiar, vllngPersonaGraba, "ASIGNACION DE NOMBRES DE MEDICAMENTOS", vlstrTipoGrupo & " " & CStr(vllngCveAfectada) & " " & vlstrTipoPaciente)
                                     EntornoSIHO.ConeccionSIHO.CommitTrans
                                       
                                      Else
                                      vlblYa = False
                                  End If
                                Next vlintContador
                               
                               
                          Case 1 'Empresa
                                 vlstrTipoPaciente = IIf(OptTipoPac(0).Value, "I", IIf(OptTipoPac(1).Value, "E", "A"))
                                 vlstrTipoGrupo = "E"
                                 If lstEmpresas.ListCount > 0 Then
                                      Else
                                      chkTodos.Value = 0
                                          Exit Sub
                                      End If
                                
                                 For vlintContador = 0 To lstEmpresas.ListCount - 1
                                      vllngCveAfectada = lstEmpresas.ItemData(vlintContador)
                                      
                                      If .RecordCount > 0 Then
                                       ' Que no exista esa informacion en la BD
                                       '-------------------------------------------------------------------------------------------------
                                             .MoveFirst
                                              Do While Not .EOF And vlblYa = False
                                                          If (!chrTipoGrupo = vlstrTipoGrupo And !intCveAfectada = vllngCveAfectada And !chrTipoPaciente = vlstrTipoPaciente And !tnyclaveempresa = vgintClaveEmpresaContable) Then
                                                            vlblYa = True
                                                          End If
                                                          .MoveNext
                                               Loop
                                      '-------------------------------------------------------------------------------------------------
                                      End If
                                       If Not (vlblYa) Then
                                           EntornoSIHO.ConeccionSIHO.BeginTrans
                                              .AddNew
                                              !chrTipoGrupo = vlstrTipoGrupo
                                              !intCveAfectada = vllngCveAfectada
                                              !chrTipoPaciente = vlstrTipoPaciente
                                              !tnyclaveempresa = vgintClaveEmpresaContable
                                           .Update
                                           Call pGuardarLogTransaccion(Me.Name, EnmCambiar, vllngPersonaGraba, "ASIGNACION DE NOMBRES DE MEDICAMENTOS", vlstrTipoGrupo & " " & CStr(vllngCveAfectada) & " " & vlstrTipoPaciente)
                                           EntornoSIHO.ConeccionSIHO.CommitTrans
                                              
                                        Else
                                          vlblYa = False
                                       End If
                                      Next vlintContador
                                
                                     
                        End Select
                
                        
                                                  
                   End With
                    
                      If SSTDescuentos.Tab = 0 Then
                        frmMantoNombreMed.Refresh
                      End If
                      pCancelar 'Solamente para limpiar la pantalla
                      ' Cargar la lista  de los guardados
                      If SSTDescuentos.Tab = 0 Then
                            pCargaGuardados "T"
                      ElseIf SSTDescuentos.Tab = 1 Then
                           pCargaGuardados "E"
                      End If
        chkTodos.Value = 0
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":chkTodos_Click"))
    Unload Me
End Sub
Private Sub lstConDescuento_DblClick()
    On Error GoTo NotificaError
    
    If SSTDescuentos.Tab = 0 Then
        lstTiposPaciente.ListIndex = fintLocalizaEnLista(lstTiposPaciente, lstConDescuento.ItemData(lstConDescuento.ListIndex))
        optTipoPac2(0).Value = Mid(Trim(lstConDescuento.List(lstConDescuento.ListIndex)), Len(Trim(lstConDescuento.List(lstConDescuento.ListIndex))) - 1, 1) = "I"
        optTipoPac2(1).Value = Mid(Trim(lstConDescuento.List(lstConDescuento.ListIndex)), Len(Trim(lstConDescuento.List(lstConDescuento.ListIndex))) - 1, 1) = "E"
        optTipoPac2(2).Value = Mid(Trim(lstConDescuento.List(lstConDescuento.ListIndex)), Len(Trim(lstConDescuento.List(lstConDescuento.ListIndex))) - 1, 1) = "A"
      '  lstTiposPaciente_DblClick
        
       
        
    ElseIf SSTDescuentos.Tab = 1 Then
        cboTipoConvenio.ListIndex = 0
        lstEmpresas.ListIndex = fintLocalizaEnLista(lstEmpresas, lstConDescuento.ItemData(lstConDescuento.ListIndex))
        OptTipoPac(0).Value = Mid(Trim(lstConDescuento.List(lstConDescuento.ListIndex)), Len(Trim(lstConDescuento.List(lstConDescuento.ListIndex))) - 1, 1) = "I"
        OptTipoPac(1).Value = Mid(Trim(lstConDescuento.List(lstConDescuento.ListIndex)), Len(Trim(lstConDescuento.List(lstConDescuento.ListIndex))) - 1, 1) = "E"
        OptTipoPac(2).Value = Mid(Trim(lstConDescuento.List(lstConDescuento.ListIndex)), Len(Trim(lstConDescuento.List(lstConDescuento.ListIndex))) - 1, 1) = "A"
     '   lstEmpresas_DblClick
        
   
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":lstConDescuento_DblClick"))
    Unload Me
End Sub

Private Function fintLocalizaEnLista(lstLista As ListBox, intClave As Integer) As Integer
    On Error GoTo NotificaError
    
    Dim vlintContador As Integer
    fintLocalizaEnLista = -1   'Regresa un -1 si no lo encuentra
    For vlintContador = 0 To lstLista.ListCount - 1
        If lstLista.ItemData(vlintContador) = intClave Then
            fintLocalizaEnLista = vlintContador
            vlintContador = lstLista.ListCount + 1
        End If
    Next vlintContador

Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fintLocalizaEnLista"))
    Unload Me
End Function



Private Sub lstConDescuento_KeyDown(KeyCode As Integer, shift As Integer)
 If KeyCode = vbKeyReturn Then
   cmdSelecciona(1).SetFocus
       
    End If
End Sub

Private Sub lstEmpresas_KeyDown(KeyCode As Integer, shift As Integer)
    On Error GoTo NotificaError
    
    If KeyCode = vbKeyReturn Then
    OptTipoPac(0).SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":lstEmpresas_KeyDown"))
    Unload Me
End Sub


Private Sub Form_Load()
    On Error GoTo NotificaError
    
    Me.Icon = frmMenuPrincipal.Icon
    
    SSTDescuentos.Tab = 0
    vglngDesktop = SysInfo1.WorkAreaHeight
    vgstrEstadoManto = ""
 
    pCargaTiposPaciente
    pCargaTipoConvenio
    pCargaGuardados ("T")
    lstTiposPaciente.Enabled = True
    lstTiposPaciente.ListIndex = 0
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
    Unload Me
End Sub

Private Sub Form_Activate()
    On Error GoTo NotificaError
    vgstrNombreForm = Me.Name
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Activate"))
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = vbKeyEscape Then
        KeyAscii = 0
        Unload Me
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyPress"))
    Unload Me
End Sub


Private Sub lstTiposPaciente_DblClick()
    On Error GoTo NotificaError
    
'    cmdGrabarRegistro.Enabled = True
'    If cmdGrabarRegistro.Visible And cmdGrabarRegistro.Enabled Then cmdGrabarRegistro.SetFocus

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":lstTiposPaciente_DblClick"))
    Unload Me
End Sub

Private Sub lstTiposPaciente_KeyDown(KeyCode As Integer, shift As Integer)
    On Error GoTo NotificaError
    
    If KeyCode = vbKeyReturn Then
    optTipoPac2(0).SetFocus
       
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":lstTiposPaciente_KeyDown"))
    Unload Me
End Sub


Private Sub optTipoPac_KeyDown(Index As Integer, KeyCode As Integer, shift As Integer)
  If KeyCode = vbKeyReturn Then
        cmdSelecciona(0).SetFocus
    End If

End Sub


Private Sub OptTipoPaciente_Click(Index As Integer)
    On Error GoTo NotificaError
    
    pCancelar

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":OptTipoPaciente_Click"))
    Unload Me
End Sub

Private Sub pCancelar()
    On Error GoTo NotificaError
    
    Dim vlintContador As Integer
    
    SSTDescuentos.TabEnabled(0) = True
    SSTDescuentos.TabEnabled(1) = True

    vgstrEstadoManto = ""
    Select Case SSTDescuentos.Tab
        Case 0
            If lstTiposPaciente.Visible And lstTiposPaciente.Enabled Then lstTiposPaciente.SetFocus
        Case 1
            If lstEmpresas.Visible And lstEmpresas.Enabled Then lstEmpresas.SetFocus
    End Select

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pcancelar"))
    Unload Me
End Sub

Private Sub optTipoPac2_KeyDown(Index As Integer, KeyCode As Integer, shift As Integer)
 If KeyCode = vbKeyReturn Then
    cmdSelecciona(0).SetFocus
       
    End If
End Sub

Private Sub SSTDescuentos_Click(PreviousTab As Integer)
    On Error GoTo NotificaError
    
    If SSTDescuentos.Tab = 0 Then
        If lstTiposPaciente.Visible And lstTiposPaciente.Enabled Then lstTiposPaciente.SetFocus
        FreConDescuento.Caption = "Mostrar el nombre genérico"
        pCargaGuardados "T"
    ElseIf SSTDescuentos.Tab = 1 Then
       If cboTipoConvenio.Visible And cboTipoConvenio.Enabled Then cboTipoConvenio.SetFocus
        FreConDescuento.Caption = "Mostrar el nombre genérico"
        pCargaGuardados "E"
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":SSTDescuentos_Click"))
    Unload Me
End Sub
