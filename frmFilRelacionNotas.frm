VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmFilRelacionNotas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relación de notas de cargo y crédito"
   ClientHeight    =   5730
   ClientLeft      =   -15
   ClientTop       =   -15
   ClientWidth     =   9300
   Icon            =   "frmFilRelacionNotas.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   9300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4890
      Left            =   90
      TabIndex        =   26
      Top             =   -45
      Width           =   9135
      Begin VB.CheckBox chkTipo 
         Caption         =   "Concentrado"
         Height          =   255
         Left            =   6720
         TabIndex        =   22
         Top             =   4440
         Width           =   1215
      End
      Begin VB.Frame FraPaciente 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   1450
         TabIndex        =   40
         Top             =   1200
         Width           =   4455
         Begin VB.OptionButton optTipoPaciente 
            Caption         =   "Todos"
            Height          =   255
            Index           =   0
            Left            =   280
            TabIndex        =   9
            Top             =   120
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optTipoPaciente 
            Caption         =   "Internos"
            Height          =   255
            Index           =   1
            Left            =   1320
            TabIndex        =   10
            Top             =   120
            Width           =   975
         End
         Begin VB.OptionButton optTipoPaciente 
            Caption         =   "Externos"
            Height          =   255
            Index           =   2
            Left            =   2400
            TabIndex        =   11
            Top             =   120
            Width           =   975
         End
      End
      Begin VB.Frame FraCliente 
         BorderStyle     =   0  'None
         Height          =   540
         Left            =   1450
         TabIndex        =   39
         Top             =   650
         Width           =   7560
         Begin VB.OptionButton optTipoCliente 
            Caption         =   "Todos"
            Height          =   220
            Index           =   5
            Left            =   280
            TabIndex        =   3
            ToolTipText     =   "Selección del tipo de cliente"
            Top             =   200
            Value           =   -1  'True
            Width           =   885
         End
         Begin VB.OptionButton optTipoCliente 
            Caption         =   "Convenio"
            Height          =   195
            Index           =   0
            Left            =   1320
            TabIndex        =   4
            Top             =   220
            Width           =   1005
         End
         Begin VB.OptionButton optTipoCliente 
            Caption         =   "Médico"
            Height          =   195
            Index           =   1
            Left            =   2400
            TabIndex        =   5
            Top             =   220
            Width           =   885
         End
         Begin VB.OptionButton optTipoCliente 
            Caption         =   "Empleado"
            Height          =   195
            Index           =   2
            Left            =   3360
            TabIndex        =   6
            Top             =   220
            Width           =   1035
         End
         Begin VB.OptionButton optTipoCliente 
            Caption         =   "Paciente interno"
            Height          =   330
            Index           =   3
            Left            =   4440
            TabIndex        =   7
            Top             =   180
            Width           =   1575
         End
         Begin VB.OptionButton optTipoCliente 
            Caption         =   "Paciente externo"
            Height          =   330
            Index           =   4
            Left            =   6000
            TabIndex        =   8
            ToolTipText     =   "Selección del tipo de cliente"
            Top             =   180
            Width           =   1605
         End
      End
      Begin VB.Frame fraNota 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   1450
         TabIndex        =   38
         Top             =   120
         Width           =   4455
         Begin VB.OptionButton optTipoNota 
            Caption         =   "Todos"
            Height          =   195
            Index           =   0
            Left            =   280
            TabIndex        =   0
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optTipoNota 
            Caption         =   "Cliente"
            Height          =   195
            Index           =   1
            Left            =   1320
            TabIndex        =   1
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton optTipoNota 
            Caption         =   "Paciente"
            Height          =   195
            Index           =   2
            Left            =   2400
            TabIndex        =   2
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame fraOrdenFecha 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   735
         Left            =   3550
         TabIndex        =   37
         Top             =   3960
         Width           =   1935
         Begin VB.OptionButton optOrdenFecha 
            Caption         =   "Factura"
            Height          =   225
            Index           =   1
            Left            =   0
            TabIndex        =   21
            Top             =   480
            Width           =   915
         End
         Begin VB.OptionButton optOrdenFecha 
            Caption         =   "Cliente / Paciente"
            Height          =   225
            Index           =   0
            Left            =   0
            TabIndex        =   20
            Top             =   120
            Width           =   1635
         End
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   1785
         Left            =   1755
         TabIndex        =   32
         Top             =   2955
         Width           =   3825
         Begin VB.Frame fraOrdenEmpresa 
            BorderStyle     =   0  'None
            Height          =   855
            Left            =   1680
            TabIndex        =   36
            Top             =   0
            Width           =   1215
            Begin VB.OptionButton optOrdenEmpresa 
               Caption         =   "Fecha"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   17
               Top             =   360
               Value           =   -1  'True
               Width           =   855
            End
            Begin VB.OptionButton optOrdenEmpresa 
               Caption         =   "Factura"
               Height          =   315
               Index           =   1
               Left            =   120
               TabIndex        =   18
               Top             =   600
               Width           =   975
            End
         End
         Begin VB.OptionButton OptTipo 
            Caption         =   "Cliente / Paciente"
            Height          =   315
            Index           =   0
            Left            =   -15
            TabIndex        =   16
            ToolTipText     =   "Reporte agrupado por cliente"
            Top             =   100
            Value           =   -1  'True
            Width           =   1590
         End
         Begin VB.OptionButton OptTipo 
            Caption         =   "Fecha"
            Height          =   200
            Index           =   1
            Left            =   -15
            TabIndex        =   19
            ToolTipText     =   "Reporte agrupado por fecha"
            Top             =   1035
            Width           =   1215
         End
      End
      Begin VB.ComboBox cboConcepto 
         Height          =   315
         Left            =   1755
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   13
         ToolTipText     =   "Concepto de la nota"
         Top             =   2190
         Width           =   6200
      End
      Begin VB.ComboBox cboCliente 
         Height          =   315
         Left            =   1755
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   12
         ToolTipText     =   "Selección del cliente"
         Top             =   1770
         Width           =   6200
      End
      Begin MSMask.MaskEdBox mskFecIni 
         Height          =   315
         Left            =   1755
         TabIndex        =   14
         ToolTipText     =   "Fecha inicial"
         Top             =   2580
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mmm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskFecFin 
         Height          =   315
         Left            =   3495
         TabIndex        =   15
         ToolTipText     =   "Fecha final"
         Top             =   2580
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mmm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de paciente"
         Height          =   195
         Left            =   180
         TabIndex        =   35
         Top             =   1320
         Width           =   1200
      End
      Begin VB.Label lbTipoNota 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de nota"
         Height          =   195
         Left            =   180
         TabIndex        =   34
         Top             =   360
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de cliente"
         Height          =   195
         Left            =   180
         TabIndex        =   33
         Top             =   840
         Width           =   1050
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Agrupación"
         Height          =   195
         Left            =   180
         TabIndex        =   31
         Top             =   3100
         Width           =   810
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "a"
         Height          =   195
         Left            =   3240
         TabIndex        =   30
         Top             =   2640
         Width           =   90
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Rango de fechas"
         Height          =   195
         Left            =   180
         TabIndex        =   29
         Top             =   2640
         Width           =   1230
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Concepto"
         Height          =   195
         Left            =   180
         TabIndex        =   28
         Top             =   2250
         Width           =   690
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         Height          =   195
         Left            =   180
         TabIndex        =   27
         Top             =   1830
         Width           =   480
      End
   End
   Begin VB.Frame Frame2 
      Height          =   705
      Left            =   4080
      TabIndex        =   25
      Top             =   4920
      Width           =   1110
      Begin VB.CommandButton cmdVistaPrevia 
         Height          =   495
         Left            =   60
         Picture         =   "frmFilRelacionNotas.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Vista previa"
         Top             =   150
         Width           =   495
      End
      Begin VB.CommandButton cmdPrint 
         Height          =   495
         Left            =   555
         Picture         =   "frmFilRelacionNotas.frx":070E
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Imprimir"
         Top             =   150
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmFilRelacionNotas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------------
' Programa para reporte de las notas de crédito y cargo
' Fecha de programación: Marzo, 2003
' Por:                   José Manuel Tórres Sáenz
'-----------------------------------------------------------------------------------
' Ultimas modificaciones, especificar:
'-----------------------------------------------------------------------------------
' Fecha:                Febrero, 2004
' Por:                  Samantha Delgado
' Descripción:          Migración de código
'-----------------------------------------------------------------------------------
' Fecha:                Marzo, 2004
' Por:                  Rosenda Hernández Anaya
' Descripción:          Control de errores, pruebas grales.
'-----------------------------------------------------------------------------------

Dim vlstrsentencia As String
Dim rs As New ADODB.Recordset
Private vgrptReporte As CRAXDRT.report


Private Sub cboConcepto_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then
        pEnfocaMkTexto mskFecIni
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboConcepto_KeyDown"))
    Unload Me
End Sub

Private Sub chkTipo_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        cmdVistaPrevia.SetFocus
    End If
    
End Sub


Private Sub cmdPrint_Click()
    On Error GoTo NotificaError
    
    
    'Osea que no an seleccionado un cliente
    If cboCliente.ListIndex = -1 Then Exit Sub

    pImprime "I"

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdPrint_Click"))
    Unload Me
End Sub

Private Sub pImprime(vlstrDestino As String)
    On Error GoTo NotificaError
    
    Dim vlintOrdenDetalle As Integer
    Dim vlstrsentencia As String
    Dim vlrsResultado As New ADODB.Recordset
    Dim strTipoNota As String
    Dim strTipoPaciente As String
    Dim strParametros As String
    Dim alstrParametros(2) As String
    Dim strTipoCliente As String
    

    If optTipoNota(0).Value = True Then strTipoNota = "A"
    If optTipoNota(1).Value = True Then strTipoNota = "C"
    If optTipoNota(2).Value = True Then strTipoNota = "P"
    
    
    If optTipoPaciente(0).Value = True Then strTipoPaciente = "A"
    If optTipoPaciente(1).Value = True Then strTipoPaciente = "I"
    If optTipoPaciente(2).Value = True Then strTipoPaciente = "E"
    
    If optTipoCliente(0).Value = True Then strTipoCliente = "CO"
    If optTipoCliente(1).Value = True Then strTipoCliente = "ME"
    If optTipoCliente(2).Value = True Then strTipoCliente = "EM"
    If optTipoCliente(3).Value = True Then strTipoCliente = "PI"
    If optTipoCliente(4).Value = True Then strTipoCliente = "PE"
    If optTipoCliente(5).Value = True Then strTipoCliente = "TO"
       
    If OptTipo(0).Value Then
        If optOrdenEmpresa(0).Value Then
            vlintOrdenDetalle = 1
        Else
            vlintOrdenDetalle = 2
        End If
    Else
        If optOrdenFecha(0).Value Then
            vlintOrdenDetalle = 3
        Else
            vlintOrdenDetalle = 4
        End If
    End If
       
'    If OptTipo(0).Value Then
'        vlintOrdenDetalle = IIf(optOrdenEmpresa(0).Value, 0, 1)
'    Else
'        vlintOrdenDetalle = IIf(optOrdenFecha(0).Value, 0, 1)
'    End If
        
    strParametros = fstrFechaSQL(mskFecIni, "00:00:00", True) & "|" & _
                                            fstrFechaSQL(mskFecFin, "23:59:59", True) & "|" & _
                                            strTipoCliente & "|" & _
                                            cboCliente.ItemData(cboCliente.ListIndex) & "|" & _
                                            cboConcepto.ItemData(cboConcepto.ListIndex) & "|" & _
                                            IIf(OptTipo(0).Value, 0, 1) & "|" & _
                                            Trim(vgstrNombreDepartamento) & "|" & _
                                            vlintOrdenDetalle & "|" & strTipoPaciente & "|" & strTipoNota & "|" & vgintClaveEmpresaContable
                                            
    alstrParametros(0) = "Concentrado" & ";" & IIf(chkTipo.Value, 1, 0) & ";BOOLEAN"
    alstrParametros(1) = "NombreHospital" & ";" & Trim(vgstrNombreHospitalCH) & ";TRUE"
    
    If OptTipo(0).Value Then
        If optOrdenEmpresa(0).Value Then
            alstrParametros(2) = "Orden" & ";" & 1 & ";NUMBER"
        Else
            alstrParametros(2) = "Orden" & ";" & 2 & ";NUMBER"
        End If
    Else
        If optOrdenFecha(0).Value Then
            alstrParametros(2) = "Orden" & ";" & 3 & ";NUMBER"
        Else
            alstrParametros(2) = "Orden" & ";" & 4 & ";NUMBER"
        End If
    End If

    If IsDate(mskFecIni) And IsDate(mskFecFin) Then
        Set vlrsResultado = frsEjecuta_SP(strParametros, "SP_CCRPTRELACIONNOTASCARGOYCRE")
        
        If vlrsResultado.RecordCount > 0 Then
            vgrptReporte.DiscardSavedData

            pCargaParameterFields alstrParametros, vgrptReporte
            
            pImprimeReporte vgrptReporte, vlrsResultado, vlstrDestino, "Relación de notas de cargo y crédito"
        Else
            MsgBox SIHOMsg(13), vbInformation + vbOKOnly, "Mensaje"
        End If
        vlrsResultado.Close
    Else
        '¡Fecha no válida!, formato de fecha dd/mm/aaaa
        MsgBox SIHOMsg(29), vbInformation, "Mensaje"
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pImprime"))
    Unload Me
End Sub


Private Sub cmdVistaPrevia_Click()
    On Error GoTo NotificaError
    
    
    'Osea que no an seleccionado un cliente
    If cboCliente.ListIndex = -1 Then Exit Sub

    pImprime "P"

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdVistaPrevia_Click"))
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
   
   
    If KeyAscii = vbKeyEscape Then
       Unload Me
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyPress"))
    Unload Me
End Sub

Private Sub Form_Load()
    On Error GoTo NotificaError
   
    Me.Icon = frmMenuPrincipal.Icon
   
   vlstrsentencia = "" & _
   "select " & _
       "chrDescripcion Descripcion," & _
       "smiCveConcepto Clave " & _
   "From " & _
       "PVConceptoFacturacion "


   Set rs = frsRegresaRs(vlstrsentencia)
   If rs.RecordCount <> 0 Then
        pLlenarCboRs cboConcepto, rs, 1, 0
   End If
   rs.Close
   cboConcepto.AddItem "<TODOS>", 0
   cboConcepto.ItemData(cboConcepto.newIndex) = -1
   cboConcepto.ListIndex = 0
   mskFecIni.Mask = ""
   mskFecIni.Text = fdtmServerFecha - 31
   mskFecIni.Mask = "##/##/####"
   
   mskFecFin.Mask = ""
   mskFecFin.Text = fdtmServerFecha
   mskFecFin.Mask = "##/##/####"
   
   OptTipo(0).Value = True
   optTipoCliente_Click 0
   
   
   optTipoNota(0).Value = True
   optTipoPaciente(0).Value = True
   fraOrdenFecha.Enabled = False
   
   optTipoCliente(5).Value = True
   
   optTipoNota_Click (0)
   
   pInstanciaReporte vgrptReporte, "rptrelacionnotascargoabono.rpt"
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
    Unload Me
End Sub



Private Sub mskFecFin_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    If KeyCode = vbKeyReturn Then
        OptTipo(0).SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFecFin_KeyDown"))
    Unload Me
End Sub

Private Sub mskFecIni_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    If KeyCode = vbKeyReturn Then
        pEnfocaMkTexto mskFecFin
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFecIni_KeyDown"))
    Unload Me
End Sub




Private Sub optOrdenEmpresa_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then
        cmdVistaPrevia.SetFocus
    End If


Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":optOrdenEmpresa_KeyDown"))
    Unload Me
End Sub

Private Sub optOrdenFecha_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then
        cmdVistaPrevia.SetFocus
    End If


Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":optOrdenFecha_KeyDown"))
    Unload Me
End Sub

Private Sub optTipo_Click(Index As Integer)
    On Error GoTo NotificaError
    
    If Index = 0 Then
    
        optOrdenFecha(0).Value = False
        optOrdenFecha(1).Value = False
        fraOrdenFecha.Enabled = False
        
        optOrdenEmpresa(0).Value = True
        
        fraOrdenEmpresa.Enabled = True
        
    Else
        optOrdenEmpresa(0).Value = False
        optOrdenEmpresa(1).Value = False
        fraOrdenEmpresa.Enabled = False
        
        optOrdenFecha(0).Value = True
        
        fraOrdenFecha.Enabled = True
        
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":optTipo_Click"))
    Unload Me
End Sub

Private Sub Opttipo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    If KeyCode = vbKeyReturn Then
        chkTipo.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":optTipo_KeyDown"))
    Unload Me
End Sub

Private Sub optTipoCliente_Click(Index As Integer)
    On Error GoTo NotificaError
    
    optTipoCliente_MouseDown Index, 1, 0, 0, 0

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":optTipoCliente_Click"))
    Unload Me
End Sub

Private Sub optTipoCliente_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    If KeyCode = vbKeyReturn Then
    
        optTipoCliente_Click Index
        
        If optTipoPaciente(0).Enabled = True Then
        
            optTipoPaciente(0).SetFocus
            
        Else
            
            cboCliente.SetFocus
            
        End If
        
        
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":optTipoCliente_KeyDown"))
    Unload Me
End Sub

Private Sub optTipoCliente_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo NotificaError
   
   cboCliente.Clear
   vlTipoCli = ""
   If Index = 0 Then
      vlTipoCli = "CO"
   ElseIf Index = 1 Then
      vlTipoCli = "ME"
   ElseIf Index = 2 Then
      vlTipoCli = "EM"
   ElseIf Index = 3 Then
      vlTipoCli = "PI"
   ElseIf Index = 4 Then
      vlTipoCli = "PE"
   End If
      vlstrsentencia = "" & _
         "select (CASE Cccliente.chrTipoCliente " & _
         "when 'CO' then isnull(CcEmpresa.vchDescripcion,' ') " & _
         "when 'ME' then isnull(HoMedico.vchApellidoPaterno,' ') ||' '|| isnull(HoMedico.vchApellidoMaterno,' ') ||' '|| isnull(HoMedico.vchNombre,' ') " & _
         "when 'EM' then isnull(NoEmpleado.vchApellidoPaterno,' ') ||' '|| isnull(NoEmpleado.vchApellidoPaterno,' ') ||' '|| isnull(NoEmpleado.vchApellidoMaterno,' ') " & _
         "when 'PI' then isnull(AdPaciente.vchApellidoPaterno,' ') ||' '|| isnull(AdPaciente.vchApellidoMaterno,' ') ||' '|| isnull(AdPaciente.vchNombre,' ') " & _
         "when 'PE' then isnull(rtrim(Externo.chrApePaterno),' ') ||' '|| isnull(rtrim(Externo.chrApeMaterno),' ') ||' '|| isnull(rtrim(Externo.chrNombre),' ') " & _
         "end) Descripcion, " & _
         "CcCliente.intNumCliente cliente "
      vlstrsentencia = vlstrsentencia & _
         "FROM CcCliente " & _
         "INNER JOIN NODEPARTAMENTO ON CCCLIENTE.SMICVEDEPARTAMENTO = NODEPARTAMENTO.SMICVEDEPARTAMENTO " & _
         "LEFT OUTER JOIN HoMedico ON CcCliente.intNumReferencia = HoMedico.intCveMedico " & _
         "left outer join AdAdmision on CcCliente.intNumReferencia = AdAdmision.numNumCuenta " & _
         "LEFT OUTER JOIN AdPaciente ON AdAdmision.numCvePaciente = AdPaciente.numCvePaciente " & _
         "left outer join RegistroExterno on CcCliente.intNumReferencia = RegistroExterno.intNumCuenta " & _
         "LEFT OUTER JOIN Externo ON RegistroExterno.intNumPaciente = Externo.intNumPaciente " & _
         "LEFT OUTER JOIN CcEmpresa ON CcCliente.intNumReferencia = CcEmpresa.intCveEmpresa " & _
         "LEFT OUTER JOIN NoEmpleado ON CcCliente.intNumReferencia = NoEmpleado.intCveEmpleado " & _
         "where Cccliente.chrTipoCliente = '" & vlTipoCli & "' And Nodepartamento.tnyclaveempresa = " & vgintClaveEmpresaContable
   
   Set rs = frsRegresaRs(vlstrsentencia)
   If rs.RecordCount > 0 Then
      pLlenarCboRs cboCliente, rs, 1, 0
   End If
   cboCliente.AddItem "<TODOS>", 0
   cboCliente.ItemData(cboCliente.newIndex) = -1
   cboCliente.ListIndex = 0
   If cboCliente.Visible And cboCliente.Enabled Then
    cboCliente.SetFocus
   End If
   optTipoCliente(Index).Value = True

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":optTipoCliente_MouseDown"))
    Unload Me
End Sub

Private Sub cboCliente_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    If KeyCode = vbKeyReturn Then
        cboConcepto.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboCliente_KeyDown"))
    Unload Me
End Sub

Private Sub optTipoNota_Click(Index As Integer)

On Error GoTo NotificaError

    If optTipoNota(0).Value = True Then
    
        optTipoPaciente(0).Enabled = False
        optTipoPaciente(0).Value = True
        
        optTipoPaciente(1).Enabled = False
        optTipoPaciente(2).Enabled = False
        
        optTipoCliente(0).Enabled = False
        optTipoCliente(1).Enabled = False
        optTipoCliente(2).Enabled = False
        optTipoCliente(3).Enabled = False
        optTipoCliente(4).Enabled = False
        optTipoCliente(5).Enabled = False
        
        optTipoCliente(5).Value = True
        
        cboCliente.Enabled = True
        
        
        FraCliente.Enabled = False
        FraPaciente.Enabled = False
    
    End If
    
    If optTipoNota(1).Value = True Then
    
        optTipoPaciente(0).Enabled = False
        optTipoPaciente(1).Enabled = False
        optTipoPaciente(2).Enabled = False
        
        optTipoCliente(0).Enabled = True
        optTipoCliente(1).Enabled = True
        optTipoCliente(2).Enabled = True
        optTipoCliente(3).Enabled = True
        optTipoCliente(4).Enabled = True
        optTipoCliente(5).Enabled = True
        
        cboCliente.Enabled = True
        
        FraCliente.Enabled = True
        FraPaciente.Enabled = True
    
    End If
    
    If optTipoNota(2).Value = True Then
    
        optTipoPaciente(0).Enabled = True
        optTipoPaciente(1).Enabled = True
        optTipoPaciente(2).Enabled = True
        
        optTipoCliente(0).Enabled = False
        optTipoCliente(1).Enabled = False
        optTipoCliente(2).Enabled = False
        optTipoCliente(3).Enabled = False
        optTipoCliente(4).Enabled = False
        optTipoCliente(5).Enabled = False
        
        cboCliente.Enabled = False
        
        FraCliente.Enabled = True
        FraPaciente.Enabled = True
            
    End If
     
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":optTipoNota_Click"))
    Unload Me
    
End Sub


Private Sub optTipoNota_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
    
        If optTipoCliente(0).Enabled = True Then
        
            optTipoCliente(5).SetFocus
            
        Else
        
            If optTipoPaciente(0).Enabled = True Then
            
                optTipoPaciente(0).SetFocus
                
            Else
            
                If cboCliente.Enabled = True Then cboCliente.SetFocus
            
            End If
                
            
        End If
        
    End If
    
End Sub


Private Sub optTipoPaciente_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
    
        If cboCliente.Enabled = True Then
        
            cboCliente.SetFocus
            
        Else
            
            cboConcepto.SetFocus
        
        End If
    
    End If
    
End Sub


