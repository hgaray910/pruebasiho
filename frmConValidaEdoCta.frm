VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmConValidaEdoCta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conexión para validación de estado de cuenta"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7185
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   7185
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab sstObj 
      Height          =   6015
      Left            =   -120
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   -360
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   10610
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmConValidaEdoCta.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraBotonera"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraInterfaz"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmConValidaEdoCta.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grdHBusqueda"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "frmConValidaEdoCta.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdDesasignar"
      Tab(2).Control(1)=   "cmdAsignar"
      Tab(2).Control(2)=   "Frame1"
      Tab(2).Control(3)=   "lstEmpresasAsignadas"
      Tab(2).Control(4)=   "lstEmpresasTodas"
      Tab(2).Control(5)=   "Label8"
      Tab(2).Control(6)=   "Label7"
      Tab(2).ControlCount=   7
      Begin VB.CommandButton cmdDesasignar 
         Caption         =   "<"
         Height          =   375
         Left            =   -71490
         TabIndex        =   36
         ToolTipText     =   "Quitar empresa"
         Top             =   2400
         Width           =   375
      End
      Begin VB.CommandButton cmdAsignar 
         Caption         =   ">"
         Height          =   375
         Left            =   -71490
         TabIndex        =   35
         ToolTipText     =   "Asignar empresa"
         Top             =   1920
         Width           =   375
      End
      Begin VB.Frame Frame1 
         Height          =   540
         Left            =   -72555
         TabIndex        =   40
         Top             =   5280
         Width           =   2295
         Begin VB.CommandButton cmdCancelar 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   1140
            TabIndex        =   39
            ToolTipText     =   "Cancelar los cambios"
            Top             =   120
            Width           =   1095
         End
         Begin VB.CommandButton cmdAceptar 
            Caption         =   "Aceptar"
            Height          =   375
            Left            =   50
            TabIndex        =   38
            ToolTipText     =   "Aceptar los cambios"
            Top             =   120
            Width           =   1095
         End
      End
      Begin VB.ListBox lstEmpresasAsignadas 
         Height          =   4350
         Left            =   -70965
         Sorted          =   -1  'True
         TabIndex        =   37
         ToolTipText     =   "Empresas asignadas"
         Top             =   840
         Width           =   3135
      End
      Begin VB.ListBox lstEmpresasTodas 
         Height          =   4350
         Left            =   -74760
         Sorted          =   -1  'True
         TabIndex        =   34
         ToolTipText     =   "Empresas disponibles"
         Top             =   840
         Width           =   3135
      End
      Begin VB.Frame fraInterfaz 
         Height          =   4725
         Left            =   240
         TabIndex        =   24
         Top             =   360
         Width           =   6960
         Begin VB.CheckBox chkCataCargosEmpresa 
            Caption         =   "Usar catálogo de cargos por empresa para códigos y descripciones"
            BeginProperty DataFormat 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "True"
               FalseValue      =   "False"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   7
            EndProperty
            Enabled         =   0   'False
            Height          =   255
            Left            =   1680
            TabIndex        =   11
            ToolTipText     =   "Usar catálogo de cargos por empresa para códigos y descripciones"
            Top             =   4080
            Width           =   5115
         End
         Begin VB.CheckBox chkUsaFiltros 
            Caption         =   "Usa filtros"
            BeginProperty DataFormat 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "True"
               FalseValue      =   "False"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   7
            EndProperty
            Height          =   255
            Left            =   1680
            TabIndex        =   10
            ToolTipText     =   "Usa filtros"
            Top             =   3790
            Width           =   1275
         End
         Begin VB.CheckBox chkConSeguridad 
            Caption         =   "Con seguridad"
            BeginProperty DataFormat 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "True"
               FalseValue      =   "False"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   7
            EndProperty
            Height          =   255
            Left            =   1680
            TabIndex        =   9
            ToolTipText     =   "Con seguridad"
            Top             =   3510
            Width           =   1395
         End
         Begin MSMask.MaskEdBox txtCveInterfaz 
            Height          =   315
            Left            =   1680
            TabIndex        =   0
            ToolTipText     =   "Clave de la conexión"
            Top             =   240
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            AllowPrompt     =   -1  'True
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin VB.TextBox txtPass 
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   1680
            MaxLength       =   50
            PasswordChar    =   "*"
            TabIndex        =   4
            ToolTipText     =   "Contraseña del usuario"
            Top             =   1680
            Width           =   3585
         End
         Begin VB.TextBox txtUsr 
            Height          =   315
            Left            =   1680
            MaxLength       =   50
            TabIndex        =   3
            ToolTipText     =   "Usuario para el servicio FTP"
            Top             =   1320
            Width           =   3585
         End
         Begin VB.TextBox txtClaveHosp 
            Height          =   315
            Left            =   1680
            MaxLength       =   9
            TabIndex        =   7
            ToolTipText     =   "Clave del hospital, proporcionada por la aseguradora y que será enviada en el estado de cuenta"
            Top             =   2760
            Width           =   1065
         End
         Begin VB.TextBox txtClaveAseg 
            Height          =   315
            Left            =   1680
            MaxLength       =   6
            TabIndex        =   8
            ToolTipText     =   "Clave de la aseguradora, proporcionada por la misma aseguradora y que será enviada en el estado de cuenta"
            Top             =   3120
            Width           =   1065
         End
         Begin VB.TextBox txtDirectorio 
            Height          =   315
            Left            =   1680
            MaxLength       =   30
            TabIndex        =   6
            ToolTipText     =   "Directorio remoto donde se colocarán los archivos"
            Top             =   2400
            Width           =   3585
         End
         Begin VB.TextBox txtURL 
            Height          =   315
            Left            =   1680
            MaxLength       =   255
            TabIndex        =   2
            ToolTipText     =   "URL del servicio FTP"
            Top             =   960
            Width           =   4785
         End
         Begin VB.TextBox txtDescripcion 
            Height          =   315
            Left            =   1680
            MaxLength       =   255
            TabIndex        =   1
            ToolTipText     =   "Descripción de la conexión"
            Top             =   600
            Width           =   4785
         End
         Begin VB.CheckBox chkActivo 
            Caption         =   "Activo"
            BeginProperty DataFormat 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "True"
               FalseValue      =   "False"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   7
            EndProperty
            Height          =   255
            Left            =   1680
            TabIndex        =   12
            ToolTipText     =   "Estado"
            Top             =   4370
            Width           =   915
         End
         Begin MSMask.MaskEdBox txtPuerto 
            Height          =   315
            Left            =   1680
            TabIndex        =   5
            ToolTipText     =   "Puerto del servicio FTP"
            Top             =   2040
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            AllowPrompt     =   -1  'True
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Clave aseguradora"
            Height          =   195
            Left            =   120
            TabIndex        =   33
            Top             =   3180
            Width           =   1335
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Clave hospital"
            Height          =   195
            Left            =   120
            TabIndex        =   32
            Top             =   2820
            Width           =   1335
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Directorio"
            Height          =   195
            Left            =   120
            TabIndex        =   31
            Top             =   2460
            Width           =   1335
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Puerto"
            Height          =   195
            Left            =   120
            TabIndex        =   30
            Top             =   2100
            Width           =   1335
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Contraseña"
            Height          =   195
            Left            =   120
            TabIndex        =   29
            Top             =   1740
            Width           =   1335
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Usuario"
            Height          =   195
            Left            =   120
            TabIndex        =   28
            Top             =   1380
            Width           =   1335
         End
         Begin VB.Label lblNombre 
            AutoSize        =   -1  'True
            Caption         =   "Descripción"
            Height          =   195
            Left            =   120
            TabIndex        =   27
            Top             =   660
            Width           =   1335
         End
         Begin VB.Label lblClave 
            AutoSize        =   -1  'True
            Caption         =   "Clave"
            Height          =   195
            Left            =   120
            TabIndex        =   26
            Top             =   300
            Width           =   1335
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "URL"
            Height          =   195
            Left            =   120
            TabIndex        =   25
            Top             =   1020
            Width           =   1335
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdHBusqueda 
         Height          =   5295
         Left            =   -74775
         TabIndex        =   23
         ToolTipText     =   "Lista de conexiones"
         Top             =   480
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   9340
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Frame fraBotonera 
         Height          =   660
         Left            =   1322
         TabIndex        =   22
         Top             =   5160
         Width           =   4540
         Begin VB.CommandButton cmdPrimerRegistro 
            Height          =   480
            Left            =   45
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmConValidaEdoCta.frx":0054
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Primer registro"
            Top             =   135
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdAnteriorRegistro 
            Height          =   480
            Left            =   540
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmConValidaEdoCta.frx":01C6
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Anterior registro"
            Top             =   135
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdBuscar 
            Height          =   480
            Left            =   1035
            Picture         =   "frmConValidaEdoCta.frx":0338
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Búsqueda"
            Top             =   135
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdSiguienteRegistro 
            Height          =   480
            Left            =   1530
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmConValidaEdoCta.frx":04AA
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Siguiente registro"
            Top             =   135
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdUltimoRegistro 
            Height          =   480
            Left            =   2025
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmConValidaEdoCta.frx":061C
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Último registro"
            Top             =   135
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdGrabarRegistro 
            Height          =   480
            Left            =   2520
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmConValidaEdoCta.frx":078E
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Guardar el registro"
            Top             =   135
            Width           =   495
         End
         Begin VB.CommandButton cmdDelete 
            Height          =   480
            Left            =   3015
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmConValidaEdoCta.frx":0AD0
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Borrar el registro"
            Top             =   135
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdEmpresas 
            Caption         =   "Empresas"
            Height          =   480
            Left            =   3520
            TabIndex        =   20
            ToolTipText     =   "Asignación de empresas"
            Top             =   135
            Width           =   975
         End
      End
      Begin VB.Label Label8 
         Caption         =   "Empresas asignadas"
         Height          =   255
         Left            =   -70960
         TabIndex        =   42
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label Label7 
         Caption         =   "Empresas disponibles"
         Height          =   255
         Left            =   -74760
         TabIndex        =   41
         Top             =   600
         Width           =   3255
      End
   End
End
Attribute VB_Name = "frmConValidaEdoCta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsInterfaz As ADODB.Recordset
Dim llngSigCve As Long
Dim lblnCargandoReg As Boolean
Dim lblnUltimoTabAceptado As Boolean
Dim colEmpresasDisponibles As Collection
Dim colEmpresasAsignadas As Collection

Private Sub chkActivo_Click()
    pModoEditaDatos
End Sub

Private Sub chkConSeguridad_Click()
    pModoEditaDatos
End Sub

Private Sub chkUsaFiltros_Click()
    pModoEditaDatos
End Sub

Private Sub chkCataCargosEmpresa_Click()
    pModoEditaDatos
End Sub

Private Sub cmdAceptar_Click()
    Dim varItm(1) As Variant
    Dim intCount As Integer
    Set colEmpresasAsignadas = New Collection
    For intCount = 0 To lstEmpresasAsignadas.ListCount - 1
        varItm(0) = lstEmpresasAsignadas.List(intCount)
        varItm(1) = lstEmpresasAsignadas.ItemData(intCount)
        colEmpresasAsignadas.Add varItm, "K" & lstEmpresasAsignadas.ItemData(intCount)
    Next
    lblnUltimoTabAceptado = True
    sstObj.Tab = 0
End Sub

Private Sub cmdAnteriorRegistro_Click()
    If Not rsInterfaz.BOF Then
        rsInterfaz.MovePrevious
        If rsInterfaz.BOF Then
           rsInterfaz.MoveFirst
        End If
        pMuestraRegistro
    End If
End Sub

Private Sub cmdAsignar_Click()
    lstEmpresasTodas_DblClick
End Sub

Private Sub cmdBuscar_Click()
On Error GoTo NotificaError
    
    If Not rsInterfaz.EOF Then
        sstObj.Tab = 1
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdBuscar_Click"))
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    sstObj.Tab = 0
End Sub

Private Sub cmdDelete_Click()
    Dim blnTransaccionIniciada As Boolean
    Dim lngPersonaGraba As Long
    
    blnTransaccionIniciada = False
    On Error GoTo NotificaError
    '| Antes de grabar revisa que el usuario tenga permiso de ESCRITURA o CONTROL TOTAL
    If fblnRevisaPermiso(vglngNumeroLogin, IIf(cgstrModulo = "SI", 3058, 3057), "C", True) Then
        If MsgBox(SIHOMsg(6), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
            lngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
            If lngPersonaGraba = 0 Then Exit Sub
            
            EntornoSIHO.ConeccionSIHO.BeginTrans
            blnTransaccionIniciada = True
            pEjecutaSentencia ("Delete PVFTPEstadoCuentaDetalle where intCveInterfaz = " & txtCveInterfaz.Text)
            pEjecutaSentencia ("Delete PVFTPEstadoCuenta where intCveInterfaz = " & txtCveInterfaz.Text)
            EntornoSIHO.ConeccionSIHO.CommitTrans
            pGuardarLogTransaccion Me.Name, EnmBorrar, lngPersonaGraba, "CONEXIÓN PARA VALIDACIÓN DE ESTADO DE CUENTA", txtCveInterfaz.Text
            
            blnTransaccionIniciada = False
            pCargaDatos
            txtCveInterfaz.SetFocus
        End If
    Else
        '| El usuario no tiene permiso para realizar esta operación.
        MsgBox SIHOMsg(635), vbOKOnly + vbExclamation, "Mensaje"
        If fblnCanFocus(txtCveInterfaz) Then txtCveInterfaz.SetFocus
    End If
    Exit Sub
NotificaError:
    If blnTransaccionIniciada Then EntornoSIHO.ConeccionSIHO.RollbackTrans
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdDelete_Click"))
End Sub

Private Sub cmdDesasignar_Click()
    lstEmpresasAsignadas_DblClick
End Sub

Private Sub cmdEmpresas_Click()
    Dim intCount As Integer
    Dim varItm As Variant
    Dim blnTest As Boolean
    lstEmpresasAsignadas.Clear
    lstEmpresasTodas.Clear
    For Each varItm In colEmpresasAsignadas
        lstEmpresasAsignadas.AddItem varItm(0)
        lstEmpresasAsignadas.ItemData(lstEmpresasAsignadas.newIndex) = CLng(varItm(1))
    Next
    
    For Each varItm In colEmpresasDisponibles
        On Error Resume Next
        blnTest = IsNull(colEmpresasAsignadas("K" & varItm(1)))
        If (Err <> 0) Then
            Err.Clear
            lstEmpresasTodas.AddItem varItm(0)
            lstEmpresasTodas.ItemData(lstEmpresasTodas.newIndex) = CLng(varItm(1))
        End If
    Next
    sstObj.Tab = 2
End Sub

Private Sub cmdGrabarRegistro_Click()
    On Error GoTo NotificaError
    Dim varItm As Variant
    Dim lngCveInterfaz As Long
    Dim vllngPersonaGraba As Long
    Dim vlClaveHosp As String
    Dim vlClaveAseg As String
    
    vlClaveHosp = Trim(txtClaveHosp.Text)
    vlClaveAseg = Trim(txtClaveAseg.Text)
    '| Antes de grabar revisa que el usuario tenga permiso de ESCRITURA o CONTROL TOTAL
    If fblnRevisaPermiso(vglngNumeroLogin, IIf(cgstrModulo = "SI", 3058, 3057), "E", True) Or fblnRevisaPermiso(vglngNumeroLogin, IIf(cgstrModulo = "SI", 3058, 3057), "C", True) Then
        If fblnValida Then
            vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
            If vllngPersonaGraba <> 0 Then
                If txtCveInterfaz.Text = llngSigCve Then
                    pEjecutaSentencia "insert into PVFTPEstadoCuenta (vchDescripcion, vchURL, intPuerto, vchUsuario, vchContrasena, vchDirectorio, bitActivo, vchClaveHospital, vchClaveAseguradora, bitConSeguridad, BITUSAFILTROS, BITCARGOSPOREMPRESA) values ('" & fstrParseo(txtDescripcion.Text) & "', '" & fstrParseo(txtURL.Text) & "', " & txtPuerto.Text & ", '" & fstrParseo(txtUsr.Text) & "', '" & fstrParseo(txtPass.Text) & "', '" & fstrParseo(txtDirectorio.Text) & "', " & IIf(chkActivo.Value = vbChecked, 1, 0) & ", '" & fstrParseo(vlClaveHosp) & "', '" & fstrParseo(vlClaveAseg) & "'," & chkConSeguridad.Value & "," & chkUsaFiltros.Value & "," & chkCataCargosEmpresa.Value & ")"
                    lngCveInterfaz = flngObtieneIdentity("SEC_PVFTPESTADOCUENTA", 0)
                    For Each varItm In colEmpresasAsignadas
                        pEjecutaSentencia "insert into PVFTPEstadoCuentaDetalle (intCveInterfaz, intCveEmpresa) values(" & lngCveInterfaz & ", " & varItm(1) & ")"
                    Next
                    pGuardarLogTransaccion Me.Name, EnmGrabar, vllngPersonaGraba, "CONEXIÓN PARA VALIDACIÓN DE ESTADO DE CUENTA", CStr(lngCveInterfaz)
                Else
                    pEjecutaSentencia "update PVFTPEstadoCuenta set vchDescripcion = '" & fstrParseo(txtDescripcion.Text) & "', vchURL = '" & fstrParseo(txtURL.Text) & "', intPuerto = " & txtPuerto.Text & ", vchUsuario = '" & fstrParseo(txtUsr.Text) & "', vchContrasena = '" & fstrParseo(txtPass.Text) & "', vchDirectorio = '" & fstrParseo(txtDirectorio.Text) & "', bitActivo = " & IIf(chkActivo.Value = vbChecked, 1, 0) & ", vchClaveHospital = '" & fstrParseo(vlClaveHosp) & "', vchClaveAseguradora = '" & fstrParseo(vlClaveAseg) & "', bitConSeguridad = " & chkConSeguridad.Value & ", BITUSAFILTROS = " & chkUsaFiltros.Value & ", BITCARGOSPOREMPRESA = " & chkCataCargosEmpresa.Value & " where intCveInterfaz = " & txtCveInterfaz.Text
                    pEjecutaSentencia "delete PVFTPEstadoCuentaDetalle where intCveInterfaz = " & txtCveInterfaz.Text
                    For Each varItm In colEmpresasAsignadas
                        pEjecutaSentencia "insert into PVFTPEstadoCuentaDetalle (intCveInterfaz, intCveEmpresa) values(" & txtCveInterfaz.Text & ", " & varItm(1) & ")"
                    Next
                    pGuardarLogTransaccion Me.Name, EnmCambiar, vllngPersonaGraba, "CONEXIÓN PARA VALIDACIÓN DE ESTADO DE CUENTA", CStr(txtCveInterfaz.Text)
                End If
                pCargaDatos
                txtCveInterfaz.SetFocus
            End If
        End If
    Else
        '| El usuario no tiene permiso para realizar esta operación.
        MsgBox SIHOMsg(635), vbOKOnly + vbExclamation, "Mensaje"
        If fblnCanFocus(txtCveInterfaz) Then txtCveInterfaz.SetFocus
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdGrabarRegistro_Click"))
End Sub

Private Function fblnValida() As Boolean
    Dim strMensaje As String
    Dim intFoco As Integer
    
    fblnValida = True
    strMensaje = ""
    intFoco = 0
    
    If Trim(txtPuerto.Text) = "" Then
        strMensaje = "Número de puerto" & vbCrLf
        intFoco = 3
    End If
    
    If Trim(txtURL.Text) = "" Then
        strMensaje = "Dirección URL" & vbCrLf
        intFoco = 2
    End If
    
    If Trim(txtDescripcion.Text) = "" Then
        strMensaje = "Descripción" & vbCrLf
        intFoco = 1
    End If
    
    If strMensaje <> "" Then
        fblnValida = False
        MsgBox SIHOMsg(2) & vbCrLf & strMensaje, vbExclamation, "Mensaje"
        Select Case intFoco
            Case 1
                If fblnCanFocus(txtDescripcion) Then txtDescripcion.SetFocus
            Case 2
                If fblnCanFocus(txtURL) Then txtURL.SetFocus
            Case 3
                If fblnCanFocus(txtPuerto) Then txtPuerto.SetFocus
        End Select
    End If
End Function

Private Sub cmdPrimerRegistro_Click()
    If Not rsInterfaz.BOF Then
        rsInterfaz.MoveFirst
        pMuestraRegistro
    End If
End Sub

Private Sub cmdSiguienteRegistro_Click()
    If Not rsInterfaz.EOF Then
        rsInterfaz.MoveNext
        If rsInterfaz.EOF Then
            rsInterfaz.MoveLast
        End If
        pMuestraRegistro
    End If
End Sub

Private Sub cmdUltimoRegistro_Click()
    If Not rsInterfaz.EOF Then
        rsInterfaz.MoveLast
        pMuestraRegistro
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
    If KeyAscii = 13 Then
        If sstObj.Tab = 0 Then
            If Me.ActiveControl.Name <> "txtCveInterfaz" Then
                SendKeys vbTab
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    Me.Icon = frmMenuPrincipal.Icon
    lblnCargandoReg = False
    llngSigCve = 1
    pCargaDatos
    lblnUltimoTabAceptado = True
    sstObj.Tab = 0
    
    chkCataCargosEmpresa.Enabled = fblnATCConCargosEmpresa
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If sstObj.Tab > 0 Then
        sstObj.Tab = 0
        Cancel = True
    Else
        If cmdGrabarRegistro.Enabled Or cmdDelete.Enabled Then
            txtCveInterfaz.SetFocus
            Cancel = True
        End If
    End If
End Sub

Private Sub pLlenaGrid()
On Error GoTo NotificaError
    
    Dim intcontador As Integer
    GrdHBusqueda.Clear
    GrdHBusqueda.Rows = 2
    GrdHBusqueda.Cols = 4
    
    If Not rsInterfaz.EOF Then
        intcontador = 1
        Do Until rsInterfaz.EOF
            llngSigCve = rsInterfaz!intcveinterfaz + 1
            GrdHBusqueda.TextMatrix(intcontador, 1) = rsInterfaz!intcveinterfaz
            GrdHBusqueda.TextMatrix(intcontador, 2) = rsInterfaz!vchDescripcion
            GrdHBusqueda.TextMatrix(intcontador, 3) = IIf(rsInterfaz!bitactivo = 1, "*", "")
            intcontador = intcontador + 1
            GrdHBusqueda.Rows = GrdHBusqueda.Rows + 1
            rsInterfaz.MoveNext
        Loop
        rsInterfaz.MoveFirst
        GrdHBusqueda.Rows = GrdHBusqueda.Rows - 1
    End If
    
    pConfiguraGrid

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pllenaGrid"))
    Unload Me
End Sub

Private Sub pConfiguraGrid()
On Error GoTo NotificaError
    
    With GrdHBusqueda
        .FormatString = "|Clave|Descripción|Activo"
        .ColWidth(0) = 100  'Fix
        .ColWidth(1) = 600 'Clave
        .ColWidth(2) = 5200 'Descripción
        .ColWidth(3) = 600 'Activo
        .ColAlignment(3) = flexAlignCenterCenter
        .ScrollBars = flexScrollBarVertical
    End With

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pConfiguraGrid"))
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    rsInterfaz.Close
End Sub

Private Sub pMuestraRegistro()
    lblnCargandoReg = True
    With rsInterfaz
        If Not .EOF Then
            txtCveInterfaz.Text = !intcveinterfaz
            txtURL.Text = !vchURL
            txtDescripcion.Text = !vchDescripcion
            txtUsr.Text = IIf(IsNull(!vchUsuario), "", !vchUsuario)
            txtPass.Text = IIf(IsNull(!vchContrasena), "", !vchContrasena)
            txtPuerto.Text = !intPuerto
            txtDirectorio = IIf(IsNull(!vchDirectorio), "", !vchDirectorio)
            txtClaveHosp = IIf(IsNull(!vchClaveHospital), "", !vchClaveHospital)
            txtClaveAseg = IIf(IsNull(!vchClaveAseguradora), "", !vchClaveAseguradora)
            chkActivo.Value = IIf(!bitactivo = 1, vbChecked, vbUnchecked)
            chkConSeguridad.Value = IIf(!bitConSeguridad = 1, vbChecked, vbUnchecked)
            chkUsaFiltros.Value = IIf(!BITUSAFILTROS = 1, vbChecked, vbUnchecked)
            chkCataCargosEmpresa.Value = IIf(!BITCARGOSPOREMPRESA = 1, vbChecked, vbUnchecked)
            pModoMuestraDatos
            pCargaEmpresasAsignadas
        Else
            pLimpiaDatos
        End If
    End With
    lblnCargandoReg = False
End Sub

Private Sub grdHBusqueda_DblClick()
    lblnUltimoTabAceptado = True
    sstObj.Tab = 0
    If GrdHBusqueda.TextMatrix(GrdHBusqueda.Row, 1) <> "" Then
        If fintLocalizaPkRs(rsInterfaz, 0, GrdHBusqueda.TextMatrix(GrdHBusqueda.Row, 1)) > 0 Then
            pMuestraRegistro
            pEnfocaTextBox txtDescripcion
        Else
            txtCveInterfaz.SetFocus
        End If
    Else
        txtCveInterfaz.SetFocus
    End If
End Sub

Private Sub lstEmpresasAsignadas_DblClick()
    If lstEmpresasAsignadas.ListIndex > -1 Then
        lstEmpresasTodas.AddItem lstEmpresasAsignadas.List(lstEmpresasAsignadas.ListIndex)
        lstEmpresasTodas.ItemData(lstEmpresasTodas.newIndex) = lstEmpresasAsignadas.ItemData(lstEmpresasAsignadas.ListIndex)
        lstEmpresasAsignadas.RemoveItem lstEmpresasAsignadas.ListIndex
    End If
End Sub

Private Sub lstEmpresasTodas_DblClick()
    If lstEmpresasTodas.ListIndex > -1 Then
        lstEmpresasAsignadas.AddItem lstEmpresasTodas.List(lstEmpresasTodas.ListIndex)
        lstEmpresasAsignadas.ItemData(lstEmpresasAsignadas.newIndex) = lstEmpresasTodas.ItemData(lstEmpresasTodas.ListIndex)
        lstEmpresasTodas.RemoveItem lstEmpresasTodas.ListIndex
    End If
End Sub

Private Sub SSTObj_Click(PreviousTab As Integer)
On Error GoTo NotificaError
    
    If sstObj.Tab = 2 Then
        lblnUltimoTabAceptado = False
        lstEmpresasTodas.Enabled = True
        lstEmpresasAsignadas.Enabled = True
        cmdAceptar.Enabled = True
        cmdCancelar.Enabled = True
        lstEmpresasTodas.SetFocus
        fraInterfaz.Enabled = False
        FraBotonera.Enabled = False
    End If
    
    If sstObj.Tab = 1 Then
        lblnUltimoTabAceptado = False
        GrdHBusqueda.Enabled = True
        GrdHBusqueda.SetFocus
        fraInterfaz.Enabled = False
        FraBotonera.Enabled = False
    End If
    
    If sstObj.Tab = 0 Then
        fraInterfaz.Enabled = True
        FraBotonera.Enabled = True
        If Not lblnUltimoTabAceptado Then
            If PreviousTab = 1 Then cmdBuscar.SetFocus
            If PreviousTab = 2 Then
                cmdEmpresas.SetFocus
            End If
        Else
            If PreviousTab = 2 Then
                pModoEditaDatos
                cmdEmpresas.SetFocus
            End If
        End If
        lstEmpresasTodas.Enabled = False
        lstEmpresasAsignadas.Enabled = False
        GrdHBusqueda.Enabled = False
        cmdAceptar.Enabled = False
        cmdCancelar.Enabled = False
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":SSTObj_Click"))
    Unload Me
End Sub

Private Sub pCargaDatos()
    Dim strSentencia As String
    strSentencia = "select * from PVFTPEstadocuenta order by intCveInterfaz"
    Set rsInterfaz = frsRegresaRs(strSentencia, adLockOptimistic, adOpenStatic)
    pCargaEmpresas
    pLlenaGrid
End Sub

Private Sub pCargaEmpresas()
    Dim rsEmpresas As ADODB.Recordset
    Dim varItm(1) As Variant
    Set colEmpresasDisponibles = New Collection
    Set rsEmpresas = frsRegresaRs("select intCveEmpresa, vchDescripcion from CCEmpresa where bitActivo = 1 and CCEmpresa.INTCVEEMPRESA NOT IN ( Select PVFTPESTADOCUENTADETALLE.INTCVEEMPRESA From PVFTPESTADOCUENTADETALLE) order by vchDescripcion")
    Do Until rsEmpresas.EOF
        varItm(0) = rsEmpresas.Fields("vchDescripcion").Value
        varItm(1) = rsEmpresas.Fields("intCveEmpresa").Value
        colEmpresasDisponibles.Add varItm, "K" & varItm(1)
        rsEmpresas.MoveNext
    Loop
    rsEmpresas.Close
End Sub

Private Sub pCargaEmpresasAsignadas()
    Dim rsEmpresas As ADODB.Recordset
    Dim varItm(1) As Variant
    Set colEmpresasAsignadas = New Collection
    Set rsEmpresas = frsRegresaRs("select CCEmpresa.intCveEmpresa, CCEmpresa.vchDescripcion from CCEmpresa inner join PVFTPEstadoCuentaDetalle on PVFTPEstadoCuentaDetalle.intCveEmpresa = CCEmpresa.intCveEmpresa and intCveInterfaz = " & txtCveInterfaz.Text & " order by vchDescripcion")
    Do Until rsEmpresas.EOF
        varItm(0) = rsEmpresas.Fields("vchDescripcion").Value
        varItm(1) = rsEmpresas.Fields("intCveEmpresa").Value
        colEmpresasAsignadas.Add varItm, "K" & varItm(1)
        rsEmpresas.MoveNext
    Loop
    rsEmpresas.Close
End Sub

Private Sub pLimpiaDatos()
    lblnCargandoReg = True
    Set colEmpresasAsignadas = New Collection
    txtCveInterfaz.Text = llngSigCve
    txtURL.Text = ""
    txtDescripcion.Text = ""
    txtUsr.Text = ""
    txtPass.Text = ""
    txtPuerto.Text = ""
    txtDirectorio = ""
    txtClaveHosp = ""
    txtClaveAseg = ""
    chkActivo.Value = vbUnchecked
    chkConSeguridad = vbUnchecked
    chkUsaFiltros = vbUnchecked
    chkCataCargosEmpresa = vbUnchecked
    cmdAnteriorRegistro.Enabled = True
    cmdPrimerRegistro.Enabled = True
    cmdBuscar.Enabled = True
    cmdSiguienteRegistro.Enabled = True
    cmdUltimoRegistro.Enabled = True
    cmdGrabarRegistro.Enabled = False
    cmdDelete.Enabled = False
    cmdEmpresas.Enabled = False
    lblnCargandoReg = False
End Sub

Private Sub pModoEditaDatos()
    If Not lblnCargandoReg Then
        cmdAnteriorRegistro.Enabled = False
        cmdPrimerRegistro.Enabled = False
        cmdBuscar.Enabled = False
        cmdSiguienteRegistro.Enabled = False
        cmdUltimoRegistro.Enabled = False
        cmdGrabarRegistro.Enabled = True
        cmdDelete.Enabled = False
        cmdEmpresas.Enabled = True
    End If
End Sub

Private Sub pModoMuestraDatos()
    cmdAnteriorRegistro.Enabled = True
    cmdPrimerRegistro.Enabled = True
    cmdBuscar.Enabled = True
    cmdSiguienteRegistro.Enabled = True
    cmdUltimoRegistro.Enabled = True
    cmdGrabarRegistro.Enabled = False
    cmdDelete.Enabled = True
    cmdEmpresas.Enabled = True
End Sub

Private Sub txtClaveAseg_Change()
    pModoEditaDatos
End Sub

Private Sub txtClaveAseg_GotFocus()
    pSelTextBox txtClaveAseg
End Sub

Private Sub txtClaveHosp_Change()
    pModoEditaDatos
End Sub

Private Sub txtClaveHosp_GotFocus()
    pSelTextBox txtClaveHosp
End Sub

Private Sub txtCveInterfaz_GotFocus()
    pLimpiaDatos
    pSelMkTexto txtCveInterfaz
End Sub

Private Sub txtCveInterfaz_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtDescripcion.SetFocus
    End If
End Sub

Private Sub txtCveInterfaz_LostFocus()
    If txtCveInterfaz.Text <> "" Then
        If fintLocalizaPkRs(rsInterfaz, 0, txtCveInterfaz.Text) > 0 Then
            pMuestraRegistro
            pSelTextBox txtDescripcion
        Else
            If Not rsInterfaz.BOF Then
                rsInterfaz.MoveLast
            End If
            pLimpiaDatos
        End If
    Else
        pLimpiaDatos
    End If
End Sub

Private Sub txtDescripcion_Change()
    pModoEditaDatos
End Sub

Private Sub txtDescripcion_GotFocus()
    pSelTextBox txtDescripcion
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtDirectorio_Change()
    pModoEditaDatos
End Sub

Private Sub txtDirectorio_GotFocus()
    pSelTextBox txtDirectorio
End Sub

Private Sub txtPass_Change()
    pModoEditaDatos
End Sub

Private Sub txtPass_GotFocus()
    pSelTextBox txtPass
End Sub

Private Sub txtPuerto_Change()
    pModoEditaDatos
End Sub

Private Sub txtPuerto_GotFocus()
    pSelMkTexto txtPuerto
End Sub

Private Sub txtURL_Change()
    pModoEditaDatos
End Sub

Private Sub txtURL_GotFocus()
    pSelTextBox txtURL
End Sub

Private Sub txtUsr_Change()
    pModoEditaDatos
End Sub

Private Sub txtUsr_GotFocus()
    pSelTextBox txtUsr
End Sub
 
Private Function fblnATCConCargosEmpresa() As Boolean
    Dim strSQL As String
    Dim strEncriptado As String
    Dim rsTemp As ADODB.Recordset
    
    fblnATCConCargosEmpresa = False
    
    strSQL = "SELECT TRIM(SIPARAMETRO.VCHVALOR) AS VALOR FROM SIPARAMETRO WHERE VCHNOMBRE = 'VCHATCCATCARGOSEMPRESA'"
    Set rsTemp = frsRegresaRs(strSQL)
    If Not rsTemp.EOF Then
        fblnATCConCargosEmpresa = IIf(IIf(IsNull(rsTemp!Valor), 0, rsTemp!Valor) = "1", True, False)
    End If
End Function
