VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmFilComisionesPromotores 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Comisiones para promotores"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6285
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   720
      Left            =   2587
      TabIndex        =   24
      Top             =   6120
      Width           =   1110
      Begin VB.CommandButton cmdImprimir 
         Height          =   495
         Left            =   555
         MaskColor       =   &H80000014&
         Picture         =   "frmFilComisionesPromotores.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Impresión física"
         Top             =   150
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdVistaPrevia 
         Height          =   495
         Left            =   60
         MaskColor       =   &H80000014&
         Picture         =   "frmFilComisionesPromotores.frx":0672
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Impresión en pantalla"
         Top             =   150
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.Frame fraTipos 
      Height          =   3375
      Left            =   120
      TabIndex        =   20
      Top             =   2760
      Width           =   6015
      Begin VB.Frame fraFormapago 
         Caption         =   "Forma de pago"
         Height          =   615
         Left            =   120
         TabIndex        =   26
         Top             =   1080
         Width           =   5655
         Begin VB.CheckBox chkCredito 
            Caption         =   "A crédito al facturar"
            Height          =   255
            Left            =   3480
            TabIndex        =   9
            ToolTipText     =   "A crédito al facturar"
            Top             =   240
            Width           =   1695
         End
         Begin VB.CheckBox chkEfectivo 
            Caption         =   "En efectivo al facturar"
            Height          =   255
            Left            =   1320
            TabIndex        =   8
            ToolTipText     =   "En efectivo al facturar"
            Top             =   240
            Value           =   1  'Checked
            Width           =   1935
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Presentación"
         Height          =   780
         Left            =   120
         TabIndex        =   25
         Top             =   2400
         Width           =   5655
         Begin VB.OptionButton optTipoReporte 
            Caption         =   "Concentrado"
            Height          =   195
            Index           =   0
            Left            =   960
            TabIndex        =   12
            ToolTipText     =   "Concentrado"
            Top             =   360
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.OptionButton optTipoReporte 
            Caption         =   "Detallado"
            Height          =   195
            Index           =   1
            Left            =   4080
            TabIndex        =   13
            ToolTipText     =   "Detallado"
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.ComboBox cboTipoPacienteEmpresa 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   7
         ToolTipText     =   "Tipo de paciente / empresa"
         Top             =   600
         Width           =   5535
      End
      Begin MSMask.MaskEdBox mskFechaInicial 
         Height          =   315
         Left            =   1440
         TabIndex        =   10
         ToolTipText     =   "Fecha de inicio"
         Top             =   1860
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mmm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskFechaFinal 
         Height          =   315
         Left            =   4390
         TabIndex        =   11
         ToolTipText     =   "Fecha final"
         Top             =   1860
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mmm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Fecha final"
         Height          =   195
         Left            =   3360
         TabIndex        =   23
         Top             =   1920
         Width           =   780
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha inicial"
         Height          =   195
         Left            =   240
         TabIndex        =   22
         Top             =   1920
         Width           =   885
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de paciente / empresa"
         Height          =   195
         Left            =   240
         TabIndex        =   21
         Top             =   300
         Width           =   1965
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Áreas de productividad"
      Height          =   1335
      Left            =   120
      TabIndex        =   19
      Top             =   1320
      Width           =   6015
      Begin VB.CheckBox chkAreaProductividad 
         Caption         =   "Pacientes ingresados al hospital"
         Height          =   375
         Index           =   1
         Left            =   240
         MaskColor       =   &H8000000F&
         TabIndex        =   3
         ToolTipText     =   "Pacientes ingresados al hospital"
         Top             =   540
         UseMaskColor    =   -1  'True
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin VB.CheckBox chkAreaProductividad 
         Caption         =   "Pacientes referidos a laboratorio"
         Height          =   375
         Index           =   4
         Left            =   3000
         TabIndex        =   6
         ToolTipText     =   "Pacientes referidos a laboratorio"
         Top             =   540
         Value           =   1  'Checked
         Width           =   2895
      End
      Begin VB.CheckBox chkAreaProductividad 
         Caption         =   "Pacientes referidos a imagenología"
         Height          =   375
         Index           =   3
         Left            =   3000
         TabIndex        =   5
         ToolTipText     =   "Pacientes referidos a imagenología"
         Top             =   240
         Value           =   1  'Checked
         Width           =   2895
      End
      Begin VB.CheckBox chkAreaProductividad 
         Caption         =   "Pacientes referidos a farmacia"
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   4
         ToolTipText     =   "Pacientes referidos a farmacia"
         Top             =   840
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin VB.CheckBox chkAreaProductividad 
         Caption         =   "Pacientes externos atendidos"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   2
         ToolTipText     =   "Pacientes externos atendidos"
         Top             =   240
         Value           =   1  'Checked
         Width           =   2655
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   6015
      Begin VB.ComboBox cboMedico 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Médico"
         Top             =   600
         Width           =   4695
      End
      Begin VB.ComboBox cboPromotor 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Promotor"
         Top             =   240
         Width           =   4695
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Médico"
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   660
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Promotor"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   300
         Width           =   630
      End
   End
End
Attribute VB_Name = "frmFilComisionesPromotores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private vgrptReporte As CRAXDRT.Report

Private Sub cboTipoPacienteEmpresa_Click()
    
    If cboTipoPacienteEmpresa.ItemData(cboTipoPacienteEmpresa.ListIndex) = 0 And cboTipoPacienteEmpresa.ListIndex > 0 Then
        cboTipoPacienteEmpresa.ListIndex = cboTipoPacienteEmpresa.ListIndex + 1
    End If
    
End Sub

Private Sub chkCredito_Click()

    If chkCredito.Value = 0 Then
        
        chkEfectivo.Value = 1
        
    End If
        
End Sub

Private Sub chkEfectivo_Click()
    
    If chkEfectivo.Value = 0 Then
    
        chkCredito.Value = 1
        
    End If
    
End Sub

Private Sub cmdImprimir_Click()
    pImprime "I"
End Sub

Private Sub cmdVistaPrevia_Click()
    pImprime "P"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            SendKeys vbTab
        Case vbKeyEscape
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
    Me.Icon = frmMenuPrincipal.Icon
    pCargaPromotor
    pCargaTipoPacienteEmpresa
    mskFechaInicial.Text = fdtmServerFecha
    mskFechaFinal.Text = fdtmServerFecha
End Sub

Private Sub pCargaPromotor()
    Dim strParametros As String
    Dim rsPromotor As New ADODB.Recordset
    
On Error GoTo NotificaError
    
    strParametros = "-1|1|" & vgintClaveEmpresaContable
    Set rsPromotor = frsEjecuta_SP(strParametros, "SP_GNSELEMPLEADO")
    pLlenarCboRs cboPromotor, rsPromotor, 0, 1, 3
    cboPromotor.ListIndex = 0

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargaPromotor"))
End Sub

Private Sub cboPromotor_Click()
    Dim strParametros As String
    Dim rsMedicoPromotor As New ADODB.Recordset
    
On Error GoTo NotificaError
    
    '|  Si están seleccionados todos los promotores se cargan todos los médicos
    If cboPromotor.ListIndex = 0 Then
        pCargaMedico
    Else '|  Si están seleccionado un promotor se cargan los médicos asignados al promotor
        strParametros = cboPromotor.ItemData(cboPromotor.ListIndex)
        Set rsMedicoPromotor = frsEjecuta_SP(strParametros, "SP_PVSELMEDICOPROMOTOR")
        pLlenarCboRs cboMedico, rsMedicoPromotor, 0, 1, 3
        cboMedico.ListIndex = 0
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboPromotor_Click"))
End Sub


Private Sub pCargaTipoPacienteEmpresa()
    Dim strParametros As String
    Dim rsTipoPacienteEmpresa As New ADODB.Recordset
    
    With cboTipoPacienteEmpresa
        .Clear
        .AddItem "<TODOS>"
        .ItemData(.NewIndex) = 0
        .AddItem "------------------------------------------------------------------------------------- Tipo de paciente"
        .ItemData(.NewIndex) = 0
        '|  Carga los tipos de paciente ordenados alfabeticamente
        strParametros = "-1"
        Set rsTipoPacienteEmpresa = frsEjecuta_SP(strParametros, "SP_ADSELTIPOPACIENTE")
        Do Until rsTipoPacienteEmpresa.EOF
            .AddItem rsTipoPacienteEmpresa!vchDescripcion
            .ItemData(.NewIndex) = rsTipoPacienteEmpresa!tnyCveTipoPaciente * -1
            rsTipoPacienteEmpresa.MoveNext
        Loop
        .AddItem "------------------------------------------------------------------------------------------------- Empresas"
        .ItemData(.NewIndex) = 0
        '|  Carga las empresas ordenadas alfabeticamente
        Set rsTipoPacienteEmpresa = frsEjecuta_SP("", "SP_ADEMPRESA")
        Do Until rsTipoPacienteEmpresa.EOF
            .AddItem rsTipoPacienteEmpresa!vchDescripcion
            .ItemData(.NewIndex) = rsTipoPacienteEmpresa!intcveempresa
            rsTipoPacienteEmpresa.MoveNext
        Loop
        .ListIndex = 0
    End With
End Sub


Private Sub pCargaMedico()
    Dim strParametros As String
    Dim rsMedico As New ADODB.Recordset
    
On Error GoTo NotificaError
    
    strParametros = "-1|1"
    Set rsMedico = frsEjecuta_SP(strParametros, "SP_EXSELMEDICO")
    pLlenarCboRs cboMedico, rsMedico, 0, 1, 3
    cboMedico.ListIndex = 0

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargaMedico"))
End Sub

Private Sub mskFechaFinal_GotFocus()
    pEnfocaMkTexto mskFechaFinal
End Sub

Private Sub mskFechaInicial_GotFocus()
    pEnfocaMkTexto mskFechaInicial
End Sub

Private Sub mskFechaInicial_LostFocus()
    On Error GoTo NotificaError


    If Not IsDate(mskFechaInicial.Text) Then
        mskFechaInicial.Mask = ""
        mskFechaInicial.Text = fdtmServerFecha
        mskFechaInicial.Mask = "##/##/####"
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFechaInicial_LostFocus"))
End Sub

Private Sub mskFechaFinal_LostFocus()
    On Error GoTo NotificaError


    If Not IsDate(mskFechaFinal.Text) Then
        mskFechaFinal.Mask = ""
        mskFechaFinal.Text = fdtmServerFecha
        mskFechaFinal.Mask = "##/##/####"
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFechaFinal_LostFocus"))
End Sub

Private Sub pImprime(strDestino As String)

On Error GoTo NotificaError

    If Not fblnValidos Then Exit Sub
    
    Dim fecha1 As Date
    Dim fecha2 As Date
    Dim alstrParametros(5) As String
    Dim rsReporte As ADODB.Recordset
    Dim vlstrSPPar As String
    Dim rs As New ADODB.Recordset
    Dim strFormadepago As String
           
        
        vlstrSPPar = cboPromotor.ItemData(cboPromotor.ListIndex) & "|" & cboMedico.ItemData(cboMedico.ListIndex)
        
        If Val(cboTipoPacienteEmpresa.ItemData(cboTipoPacienteEmpresa.ListIndex)) < 0 Then
            vlstrSPPar = vlstrSPPar & "|0|" & Abs(Val(cboTipoPacienteEmpresa.ItemData(cboTipoPacienteEmpresa.ListIndex)))
        Else
            If Val(cboTipoPacienteEmpresa.ItemData(cboTipoPacienteEmpresa.ListIndex)) > 0 Then
                vlstrSPPar = vlstrSPPar & "|" & cboTipoPacienteEmpresa.ItemData(cboTipoPacienteEmpresa.ListIndex) & "|0"
            Else
                vlstrSPPar = vlstrSPPar & "|0|0"
            End If
        End If
        
        vlstrSPPar = vlstrSPPar & "|" & fstrFechaSQL(mskFechaInicial.Text, "00:00:00") & "|" & fstrFechaSQL(mskFechaFinal.Text, "23:59:59")
        
        
        If chkEfectivo.Value = 1 And chkCredito.Value = 1 Then
            
            strFormadepago = "*"
        
        Else
            
            If chkEfectivo.Value = 1 Then
            
                strFormadepago = "E"
            
            Else
            
                strFormadepago = "C"
                
            End If
            
            
        End If
        
        
        'Tipo de Reporte
        If optTipoReporte(0).Value = True Then
        
                vlstrSPPar = vlstrSPPar & "|" & Str(vgintClaveEmpresaContable) _
                & "|" & IIf(chkAreaProductividad(0).Value = 1, 1, 0) & "|" & IIf(chkAreaProductividad(1).Value = 1, 1, 0) _
                & "|" & IIf(chkAreaProductividad(2) = 1, 1, 0) & "|" & IIf(chkAreaProductividad(3).Value = 1, 1, 0) _
                & "|" & IIf(chkAreaProductividad(4).Value = 1, 1, 0) & "|" & strFormadepago
                Set rs = frsEjecuta_SP(vlstrSPPar, "Sp_PvrptComisionesPromotores")
                pInstanciaReporte vgrptReporte, "rptComisionesPromotoresCon.rpt"
                
        Else
                
                vlstrSPPar = vlstrSPPar & "|" & Str(vgintClaveEmpresaContable) _
                & "|" & IIf(chkAreaProductividad(0).Value = 1, 1, 0) & "|" & IIf(chkAreaProductividad(1).Value = 1, 1, 0) _
                & "|" & IIf(chkAreaProductividad(2).Value = 1, 1, 0) & "|" & IIf(chkAreaProductividad(3).Value = 1, 1, 0) _
                & "|" & IIf(chkAreaProductividad(4).Value = 1, 1, 0) & "|" & strFormadepago
                Set rs = frsEjecuta_SP(vlstrSPPar, "Sp_PvrptComisionesPromotores")
                pInstanciaReporte vgrptReporte, "rptComisionesPromotoresDet.rpt"
                
       End If
            
       If rs.EOF Then
            MsgBox SIHOMsg(13), vbInformation, "Mensaje"
       Else
            vgrptReporte.DiscardSavedData
            
            alstrParametros(0) = "NombreHospital;" & Trim(vgstrNombreHospitalCH)
            alstrParametros(1) = "FechaIni;" & UCase(Format(mskFechaInicial.Text, "dd/mmm/yyyy"))
            alstrParametros(2) = "FechaFin;" & UCase(Format(mskFechaFinal.Text, "dd/mmm/yyyy"))
            alstrParametros(3) = "Promotor;" & cboPromotor.Text
            alstrParametros(4) = "Medico;" & cboMedico.Text
            
            pCargaParameterFields alstrParametros, vgrptReporte
    
            pImprimeReporte vgrptReporte, rs, strDestino, "Comisiones de promotores"
                        
                        
      End If
      
      rs.Close
   
    
Exit Sub

NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pImprime"))
    
End Sub

Private Function fblnValidos() As Boolean
    On Error GoTo NotificaError
    
    fblnValidos = True
    
    If CDate(mskFechaInicial.Text) > CDate(mskFechaFinal.Text) Then
        fblnValidos = False
        '¡Rango de fechas no válido!
        MsgBox SIHOMsg(64), vbOKOnly + vbError, "Mensaje"
        mskFechaInicial.SetFocus
    End If

    Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fblnValidos"))
End Function

