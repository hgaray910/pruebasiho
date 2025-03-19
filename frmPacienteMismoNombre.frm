VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmPacienteMismoNombre 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pacientes con el mismo nombre"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10845
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   10845
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
      Height          =   3765
      Left            =   120
      TabIndex        =   0
      Top             =   700
      Width           =   10610
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdPacientes 
         Height          =   3615
         Left            =   0
         TabIndex        =   1
         Top             =   150
         Width           =   10600
         _ExtentX        =   18706
         _ExtentY        =   6376
         _Version        =   393216
         ForeColor       =   0
         Cols            =   7
         ForeColorFixed  =   0
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorUnpopulated=   -2147483643
         GridColor       =   -2147483638
         GridColorFixed  =   -2147483638
         GridColorUnpopulated=   -2147483638
         GridLinesFixed  =   1
         GridLinesUnpopulated=   1
         Appearance      =   0
         GridLineWidthFixed=   1
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
         _Band(0).Cols   =   7
         _Band(0).GridLineWidthBand=   1
      End
   End
   Begin VB.Label lblMensaje558 
      BackColor       =   &H80000005&
      Caption         =   "Mensaje 558"
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
      TabIndex        =   2
      Top             =   60
      Width           =   10605
   End
End
Attribute VB_Name = "frmPacienteMismoNombre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Modificaciones para el registro único de pacientes
' Eliminé el uso de la variable <gblnCatCentralizados>

'-------------------------------------------------------------------------------------
'| Nombre del Proyecto      : prjAdmision
'| Nombre del Formulario    : frmPacienteMismoNombre
'-------------------------------------------------------------------------------------
'| Objetivo:    Forma usada por los procesos de admision normal y express para
'|              mostrar un listado de pacientes cuando el nombre del paciente que
'|              se está ingresando es igual a pacientes registrados.
'-------------------------------------------------------------------------------------

Option Explicit

Public vllngNumeroExpediente As Long        'Número de expediente del paciente seleccionado
Public vllngNumeroCuenta As Long            'Número de cuenta del paciente seleccionado

Public vlblnIngresoPrevio As Boolean        'Estado de ingreso previo de AdPaciente del paciente seleccionado (bitIngresoPrevio)

Public vlstrNombre As String                'Nombre del paciente seleccionado
Public vlstrApellidoPaterno As String       'Apellido paterno del paciente seleccionado
Public vlstrApellidoMaterno As String       'Apellido materno del paciente seleccionado


Public vlstrSexo As String                  'Sexo del paciente seleccionado
Public vlstrFechaNacimiento As String       'Fecha nacimiento, paciente seleccionado
Public vlstrRFC As String                   'RFC paciente seleccionado
Public vlstrCurp As String                  'Curp paciente seleccionado
Public vlstrDomicilio As String             'Domicilio paciente seleccionado
Public vlstrNumeroExterior As String 'Número exterior del paciente
Public vlstrNumeroInterior As String 'Número interior del paciente
Public vlstrColonia As String               'Colonia paciente seleccionado
Public vlstrCveCiudad As String             'Clave de la ciudad paciente seleccionado
Public vlstrTelefono As String              'Telefono paciente seleccionado
Public vlstrEstadoAdmision As String        'Estado de la admisión
                                            '"A" = actualmente interno
                                            '"P" = prepago
                                            '"E" = paciente egresado
                                            '<> anteriores = puede ser por ingreso previo o algun otro
Public vlintmostrardatos As Integer
Public vlstrNombreBuscar As String
Public vlstrPaternoBuscar As String
Public vlstrMaternoBuscar As String

Public lblnMotrarPreregistro As Boolean

Const cintColNombreCompleto = 1
Const cintColNumeroPaciente = 2
Const cintColPaterno = 3
Const cintColMaterno = 4
Const cintColNombre = 5
Const cintColSexo = 6
Const cintColFechaNacimiento = 7
Const cintColRFC = 8
Const cintColCURP = 9
Const cintColDomicilioCompleto = 10
Const cintColCalleNumero = 11
Const cintColColonia = 12
Const cintColCveCiudad = 13
Const cintColTelefono = 14
Const cintColPrevio = 15
Const cintColNumExterior = 16
Const cintColNuminterior = 17
Private Const cintColCvePreRegistro As Integer = 18
Private Const cintColumnas As Integer = 20
Private Const cintColTipoIngreso = 19

Dim vlstrSentencia As String
Dim blnFirst As Boolean

Private Sub Form_Activate()
    On Error GoTo NotificaError
    
    If blnFirst Then
        blnFirst = False
        pCargaPacientes
        If Trim(grdPacientes.TextMatrix(1, 1)) = "" Then
            'No encontró pacientes con el mismo nombre
            Me.Hide
        End If
    Else
        pCargaPacientes
        If Trim(grdPacientes.TextMatrix(1, 1)) = "" Then
            'No encontró pacientes con el mismo nombre
            Me.Hide
        End If
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Activate"))
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 27 Then
        Me.Hide
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyPress"))
    Unload Me
End Sub

Private Sub Form_Load()
    On Error GoTo NotificaError
    
    Me.Icon = frmMenuPrincipal.Icon
    blnFirst = True
    
    lblMensaje558.Caption = SIHOMsg(558)
    
    vllngNumeroExpediente = 0
    vllngNumeroCuenta = 0

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
    Unload Me
End Sub

Private Sub pCargaPacientes()
    On Error GoTo NotificaError
    
    Dim rsPacientes As New ADODB.Recordset
    Dim intcontador As Long
    Dim rsPreRegistro  As New ADODB.Recordset
    Dim vlstrSentencia As String
    Dim rsTipo As ADODB.Recordset
    Dim strTipo As String
    
With grdPacientes
        .Clear
        .Rows = 2
        .Cols = 20 'cintColumnas
        .FixedCols = 1
        .FixedRows = 1
        .FormatString = "|Nombre|Número|Paterno|Materno|Nombre|Sexo|Fecha nacimiento|RFC|CURP|Domicilio|CalleNumero|Colonia|CveCiudad|Teléfono|Previo||||Tipo de paciente"
        For intcontador = 1 To .Cols - 1
            .TextMatrix(1, intcontador) = ""
        Next intcontador
    End With

    vgstrParametrosSP = Trim(vlstrPaternoBuscar) & "|" & Trim(vlstrMaternoBuscar)
    Set rsPacientes = frsEjecuta_SP(vgstrParametrosSP, "SP_GNSELPACIENTEDATOSSIMILARES")
    
    With rsPacientes
        If .RecordCount <> 0 Then
            grdPacientes.Row = 1
            
            Do While Not .EOF
                grdPacientes.TextMatrix(grdPacientes.Rows - 1, cintColNombreCompleto) = IIf(IsNull(!nombreCompleto), "", !nombreCompleto)
                grdPacientes.TextMatrix(grdPacientes.Rows - 1, cintColNumeroPaciente) = IIf(!NumeroPaciente > 0, !NumeroPaciente, "")
                
                If !NumeroPaciente > 0 Then
                    vlstrSentencia = "SELECT INTNUMPACIENTE, MAX(INTNUMCUENTA), EXPACIENTEINGRESO.INTCVETIPOINGRESO, SITIPOINGRESO.VCHNOMBRE As " & Chr(34) & "Tipo" & Chr(34) & " FROM EXPACIENTEINGRESO JOIN SITIPOINGRESO ON SITIPOINGRESO.INTCVETIPOINGRESO = EXPACIENTEINGRESO.INTCVETIPOINGRESO WHERE INTNUMPACIENTE = " & !NumeroPaciente & " and rownum = 1 GROUP BY EXPACIENTEINGRESO.INTCVETIPOINGRESO, INTNUMPACIENTE, SITIPOINGRESO.VCHNOMBRE"
                    Set rsTipo = frsRegresaRs(vlstrSentencia)
                End If
                
                If cgstrModulo = "LA" Or cgstrModulo = "IM" Or cgstrModulo = "PV" Or cgstrModulo = "CC" Or cgstrModulo = "BS" Then
                    strTipo = "EXTERNO"
                Else
                    strTipo = "URGENCIAS"
                End If
                
                If Not rsTipo.EOF Then
                    If rsTipo!Tipo = "EXTERNO" Then
                        If cgstrModulo = "LA" Or cgstrModulo = "IM" Or cgstrModulo = "PV" Or cgstrModulo = "CC" Or cgstrModulo = "BS" Then
                            strTipo = "EXTERNO"
                        Else
                            strTipo = "URGENCIAS"
                        End If
                    Else
                        strTipo = rsTipo!Tipo
                    End If
                End If
                
                grdPacientes.TextMatrix(grdPacientes.Rows - 1, cintColPaterno) = IIf(IsNull(!Paterno), "", !Paterno)
                grdPacientes.TextMatrix(grdPacientes.Rows - 1, cintColMaterno) = IIf(IsNull(!Materno), "", !Materno)
                grdPacientes.TextMatrix(grdPacientes.Rows - 1, cintColNombre) = IIf(IsNull(!Nombre), "", !Nombre)
                grdPacientes.TextMatrix(grdPacientes.Rows - 1, cintColSexo) = IIf(IsNull(!Sexo), "", !Sexo)
                grdPacientes.TextMatrix(grdPacientes.Rows - 1, cintColFechaNacimiento) = IIf(IsNull(!FechaNacimiento), "", Format(!FechaNacimiento, "dd/mmm/yyyy"))
                grdPacientes.TextMatrix(grdPacientes.Rows - 1, cintColRFC) = IIf(IsNull(!RFC), "", !RFC)
                grdPacientes.TextMatrix(grdPacientes.Rows - 1, cintColCURP) = IIf(IsNull(!CURP), "", !CURP)
                grdPacientes.TextMatrix(grdPacientes.Rows - 1, cintColDomicilioCompleto) = IIf(IsNull(!DomicilioCompleto), "", !DomicilioCompleto)
                grdPacientes.TextMatrix(grdPacientes.Rows - 1, cintColCalleNumero) = IIf(IsNull(!Domicilio), "", !Domicilio)
                grdPacientes.TextMatrix(grdPacientes.Rows - 1, cintColColonia) = IIf(IsNull(!Colonia), "", !Colonia)
                grdPacientes.TextMatrix(grdPacientes.Rows - 1, cintColCveCiudad) = IIf(IsNull(!ClaveCiudad), 0, !ClaveCiudad)
                grdPacientes.TextMatrix(grdPacientes.Rows - 1, cintColTelefono) = IIf(IsNull(!Telefono), "", !Telefono)
                grdPacientes.TextMatrix(grdPacientes.Rows - 1, cintColPrevio) = !Previo
                grdPacientes.TextMatrix(grdPacientes.Rows - 1, cintColNumExterior) = IIf(IsNull(!NumeroExterior), "", !NumeroExterior)
                grdPacientes.TextMatrix(grdPacientes.Rows - 1, cintColNuminterior) = IIf(IsNull(!NumeroInterior), "", !NumeroInterior)
                grdPacientes.TextMatrix(grdPacientes.Rows - 1, cintColCvePreRegistro) = ""
                grdPacientes.TextMatrix(grdPacientes.Rows - 1, cintColTipoIngreso) = strTipo
                grdPacientes.Rows = grdPacientes.Rows + 1
                .MoveNext
            Loop
            grdPacientes.Rows = grdPacientes.Rows - 1
        End If
    End With
    If lblnMotrarPreregistro Then
        vgstrParametrosSP = Trim(vlstrPaternoBuscar) & "|" & Trim(vlstrMaternoBuscar)
        Set rsPreRegistro = frsEjecuta_SP(vgstrParametrosSP, "SP_GNSELDATOSPREREGISTRO")
        If rsPreRegistro.RecordCount <> 0 Then
            grdPacientes.Row = grdPacientes.Rows - 1
            With rsPreRegistro
                Do While Not .EOF
                    grdPacientes.TextMatrix(grdPacientes.Rows - 1, cintColNombreCompleto) = IIf(IsNull(!nombreCompleto), "", !nombreCompleto)
                    grdPacientes.TextMatrix(grdPacientes.Rows - 1, cintColNumeroPaciente) = ""
                    grdPacientes.TextMatrix(grdPacientes.Rows - 1, cintColPaterno) = IIf(IsNull(!Paterno), "", !Paterno)
                    grdPacientes.TextMatrix(grdPacientes.Rows - 1, cintColMaterno) = IIf(IsNull(!Materno), "", !Materno)
                    grdPacientes.TextMatrix(grdPacientes.Rows - 1, cintColNombre) = IIf(IsNull(!Nombre), "", !Nombre)
                    grdPacientes.TextMatrix(grdPacientes.Rows - 1, cintColSexo) = IIf(IsNull(!Sexo), "", !Sexo)
                    grdPacientes.TextMatrix(grdPacientes.Rows - 1, cintColFechaNacimiento) = IIf(IsNull(!FechaNacimiento), "", Format(!FechaNacimiento, "dd/mmm/yyyy"))
                    grdPacientes.TextMatrix(grdPacientes.Rows - 1, cintColRFC) = IIf(IsNull(!RFC), "", !RFC)
                    grdPacientes.TextMatrix(grdPacientes.Rows - 1, cintColCURP) = IIf(IsNull(!CURP), "", !CURP)
                    grdPacientes.TextMatrix(grdPacientes.Rows - 1, cintColDomicilioCompleto) = IIf(IsNull(!DomicilioCompleto), "", !DomicilioCompleto)
                    grdPacientes.TextMatrix(grdPacientes.Rows - 1, cintColCalleNumero) = IIf(IsNull(!Domicilio), "", !Domicilio)
                    grdPacientes.TextMatrix(grdPacientes.Rows - 1, cintColColonia) = IIf(IsNull(!Colonia), "", !Colonia)
                    grdPacientes.TextMatrix(grdPacientes.Rows - 1, cintColCveCiudad) = IIf(IsNull(!ClaveCiudad), 0, !ClaveCiudad)
                    grdPacientes.TextMatrix(grdPacientes.Rows - 1, cintColTelefono) = IIf(IsNull(!Telefono), "", !Telefono)
                    grdPacientes.TextMatrix(grdPacientes.Rows - 1, cintColPrevio) = "0"
                    grdPacientes.TextMatrix(grdPacientes.Rows - 1, cintColNumExterior) = IIf(IsNull(!NumeroExterior), "", !NumeroExterior)
                    grdPacientes.TextMatrix(grdPacientes.Rows - 1, cintColNuminterior) = IIf(IsNull(!NumeroInterior), "", !NumeroInterior)
                    grdPacientes.TextMatrix(grdPacientes.Rows - 1, cintColCvePreRegistro) = !NumeroPaciente
                    grdPacientes.TextMatrix(grdPacientes.Rows - 1, cintColTipoIngreso) = ""
                    grdPacientes.Rows = grdPacientes.Rows + 1
                    .MoveNext
                Loop
                grdPacientes.Rows = grdPacientes.Rows - 1
            End With
        End If
    End If
    
    pOrdColMshFGrid grdPacientes, cintColNombreCompleto
    
    
    With grdPacientes
        .ColWidth(0) = 100
        .ColWidth(cintColNombreCompleto) = 3500
        .ColWidth(cintColNumeroPaciente) = 1000
        .ColWidth(cintColPaterno) = 0
        .ColWidth(cintColMaterno) = 0
        .ColWidth(cintColNombre) = 0
        .ColWidth(cintColSexo) = 600
        .ColWidth(cintColFechaNacimiento) = 2000
        .ColWidth(cintColRFC) = 1400
        .ColWidth(cintColCURP) = 1400
        .ColWidth(cintColDomicilioCompleto) = 6150
        .ColWidth(cintColCalleNumero) = 0
        .ColWidth(cintColColonia) = 0
        .ColWidth(cintColCveCiudad) = 0
        .ColWidth(cintColTelefono) = 0
        .ColWidth(cintColPrevio) = 0
        .ColWidth(cintColNumExterior) = 0
        .ColWidth(cintColNuminterior) = 0
        .ColWidth(cintColCvePreRegistro) = 0
        .ColWidth(cintColTipoIngreso) = 6150
        
        .ColAlignment(cintColNombreCompleto) = flexAlignLeftCenter
        .ColAlignment(cintColNumeroPaciente) = flexAlignRightCenter
        .ColAlignment(cintColSexo) = flexAlignCenterCenter
        .ColAlignment(cintColFechaNacimiento) = flexAlignLeftCenter
        .ColAlignment(cintColRFC) = flexAlignLeftCenter
        .ColAlignment(cintColCURP) = flexAlignLeftCenter
        .ColAlignment(cintColDomicilioCompleto) = flexAlignLeftCenter
        .ColAlignment(cintColTipoIngreso) = flexAlignLeftCenter
        
        .ColAlignmentFixed(cintColNombreCompleto) = flexAlignCenterCenter
        .ColAlignmentFixed(cintColNumeroPaciente) = flexAlignCenterCenter
        .ColAlignmentFixed(cintColSexo) = flexAlignCenterCenter
        .ColAlignmentFixed(cintColFechaNacimiento) = flexAlignCenterCenter
        .ColAlignmentFixed(cintColRFC) = flexAlignCenterCenter
        .ColAlignmentFixed(cintColCURP) = flexAlignCenterCenter
        .ColAlignmentFixed(cintColDomicilioCompleto) = flexAlignCenterCenter
        .ColAlignmentFixed(cintColTipoIngreso) = flexAlignCenterCenter
    End With
    
    grdPacientes.Row = 1
    grdPacientes.Col = cintColNombreCompleto
    

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCargaPacientes"))
    Unload Me
End Sub

Private Sub pValidaIngreso()
    On Error GoTo NotificaError
        
    Dim rsUltimoInternamiento As New ADODB.Recordset
    
    vlintmostrardatos = 1
    
    If Trim(grdPacientes.TextMatrix(1, 1)) <> "" Then
        
        If Val(grdPacientes.TextMatrix(grdPacientes.Row, cintColNumeroPaciente)) > 0 Then
            vllngNumeroExpediente = CLng(grdPacientes.TextMatrix(grdPacientes.Row, cintColNumeroPaciente))
        Else
            vllngNumeroExpediente = CLng(grdPacientes.TextMatrix(grdPacientes.Row, cintColCvePreRegistro))
        End If
        vlblnIngresoPrevio = CBool(grdPacientes.TextMatrix(grdPacientes.Row, cintColPrevio))
        
        vlstrNombre = Trim(grdPacientes.TextMatrix(grdPacientes.Row, cintColNombre))
        vlstrApellidoPaterno = Trim(grdPacientes.TextMatrix(grdPacientes.Row, cintColPaterno))
        vlstrApellidoMaterno = Trim(grdPacientes.TextMatrix(grdPacientes.Row, cintColMaterno))
        
        vlstrSexo = Trim(grdPacientes.TextMatrix(grdPacientes.Row, cintColSexo))
        vlstrFechaNacimiento = IIf(Trim(grdPacientes.TextMatrix(grdPacientes.Row, cintColFechaNacimiento)) = "", "", Format(Trim(grdPacientes.TextMatrix(grdPacientes.Row, cintColFechaNacimiento)), "dd/mm/yyyy"))
        vlstrRFC = Trim(grdPacientes.TextMatrix(grdPacientes.Row, cintColRFC))
        vlstrCurp = Trim(grdPacientes.TextMatrix(grdPacientes.Row, cintColCURP))
        vlstrDomicilio = Trim(grdPacientes.TextMatrix(grdPacientes.Row, cintColCalleNumero))
        vlstrColonia = Trim(grdPacientes.TextMatrix(grdPacientes.Row, cintColColonia))
        vlstrCveCiudad = Trim(grdPacientes.TextMatrix(grdPacientes.Row, cintColCveCiudad))
        vlstrTelefono = Trim(grdPacientes.TextMatrix(grdPacientes.Row, cintColTelefono))
        vlstrNumeroExterior = Trim(grdPacientes.TextMatrix(grdPacientes.Row, cintColNumExterior))
        vlstrNumeroInterior = Trim(grdPacientes.TextMatrix(grdPacientes.Row, cintColNuminterior))
                
        If vllngNumeroExpediente > 0 Then 'negativo cuando es de pre registro de paciente, es decir no existe en exPaciente
            'Buscar el último internamiento que haya tenido el paciente:
            Set rsUltimoInternamiento = frsEjecuta_SP(Str(vllngNumeroExpediente) & "|" & vgintNumeroDepartamento, "SP_ADSELULTIMOINTERNAMIENTO")
            If rsUltimoInternamiento.RecordCount <> 0 Then
                vllngNumeroCuenta = rsUltimoInternamiento!intNumCuenta 'numNumCuenta
                vlstrEstadoAdmision = rsUltimoInternamiento!ChrEstatus 'chrEstatusAdmision
            Else
                vllngNumeroCuenta = 0
                vlstrEstadoAdmision = "X"
            End If
        Else
            vllngNumeroCuenta = 0
            vlstrEstadoAdmision = "X"
        End If
        
        Me.Hide
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pValidaIngreso"))
    Unload Me
End Sub

Private Sub grdPacientes_DblClick()
    On Error GoTo NotificaError
    
    pValidaIngreso

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdPacientes_DblClick"))
    Unload Me
End Sub

Private Sub grdPacientes_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    If KeyCode = vbKeyReturn Then
        grdPacientes_DblClick
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdPacientes_KeyDown"))
    Unload Me
End Sub


