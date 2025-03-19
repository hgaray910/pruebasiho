VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRptSociosEgresados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Socios egresados"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4410
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   4410
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraReporte 
      Height          =   5295
      Left            =   0
      TabIndex        =   8
      Top             =   -120
      Width           =   4425
      Begin VB.Frame Frame2 
         Caption         =   "Fecha de egreso"
         Height          =   810
         Left            =   120
         TabIndex        =   12
         Top             =   1920
         Width           =   4200
         Begin MSComCtl2.DTPicker dtpFechaInicial 
            Height          =   300
            Left            =   675
            TabIndex        =   4
            ToolTipText     =   "Fecha de inicio"
            Top             =   310
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "dd/MMM/yyy"
            Format          =   99155971
            CurrentDate     =   40882
         End
         Begin MSComCtl2.DTPicker dtpFechaFinal 
            Height          =   300
            Left            =   2700
            TabIndex        =   5
            ToolTipText     =   "Fecha final"
            Top             =   310
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "dd/MMM/yyy"
            Format          =   99155971
            CurrentDate     =   40882
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            Height          =   195
            Left            =   2205
            TabIndex        =   14
            Top             =   405
            Width           =   420
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   405
            Width           =   465
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Departamento"
         Height          =   810
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   4200
         Begin VB.ComboBox cboDepartamento 
            Height          =   315
            Left            =   120
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   315
            Width           =   3915
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Tipo de paciente"
         Height          =   810
         Left            =   120
         TabIndex        =   10
         Top             =   1065
         Width           =   4200
         Begin VB.OptionButton optPaciente 
            Caption         =   "Externos"
            Height          =   195
            Index           =   2
            Left            =   3100
            TabIndex        =   3
            ToolTipText     =   "Pacientes externos"
            Top             =   380
            Width           =   930
         End
         Begin VB.OptionButton optPaciente 
            Caption         =   "Internos"
            Height          =   195
            Index           =   1
            Left            =   1635
            TabIndex        =   2
            ToolTipText     =   "Pacientes internos"
            Top             =   380
            Width           =   1170
         End
         Begin VB.OptionButton optPaciente 
            Caption         =   "Todos"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   1
            ToolTipText     =   "Todos los tipos de paciente"
            Top             =   380
            Value           =   -1  'True
            Width           =   930
         End
      End
      Begin VB.Frame fraImprime 
         Height          =   735
         Left            =   1680
         TabIndex        =   9
         Top             =   2760
         Width           =   1155
         Begin VB.CommandButton cmdPreview 
            Height          =   495
            Left            =   80
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmRptSociosEgresados.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Vista previa"
            Top             =   150
            Width           =   495
         End
         Begin VB.CommandButton cmdPrint 
            Height          =   495
            Left            =   580
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmRptSociosEgresados.frx":0403
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Imprimir"
            Top             =   150
            Width           =   495
         End
      End
   End
End
Attribute VB_Name = "frmRptSociosEgresados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private vgrptReporte As CRAXDRT.Report


Private Sub cboDepartamentos_Click()

End Sub

Private Sub cboDepartamento_Click()
   Dim vlstr As String
   Dim rs As New ADODB.Recordset
   
   vlstrsql = "SELECT SMICVELOCALIZACION, VCHDESCRIPCION, SMICVEDEPARTAMENTO, BITACTIVO From IvLocalizacion WHERE (bitActivo = 1) "
   If cboDepartamento.ItemData(cboDepartamento.ListIndex) <> -1 Then
      vlstr = "select SMICVELOCALIZACION, VCHDESCRIPCION, SMICVEDEPARTAMENTO, BITACTIVO from ivlocalizacion where smiCveDepartamento = " & cboDepartamento.ItemData(cboDepartamento.ListIndex)
      Set rs = frsRegresaRs(vlstr, adLockReadOnly, adOpenForwardOnly)
      If rs.RecordCount > 0 Then
         vlstrsql = vlstrsql & " and smiCveDepartamento = " & cboDepartamento.ItemData(cboDepartamento.ListIndex)
      End If
      rs.Close
   End If
End Sub

Private Sub cboDepartamento_KeyPress(KeyAscii As Integer)
  If KeyCode = 13 Then
    SendKeys vbTab
  End If
End Sub

Private Sub cmdPreview_Click()
    
On Error GoTo NotificaError
    
    pImprime "P"

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdPreview_Click"))
    Unload Me
    
End Sub

Private Sub cmdPrint_Click()

On Error GoTo NotificaError
    
    pImprime "I"

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdPrint_Click"))
    Unload Me
    
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo NotificaError

    Select Case KeyCode
        
        Case 27
                
                    Unload Me
                    
        Case 13
            
                SendKeys vbTab
                        
    End Select
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyDown"))
    Unload Me
    
End Sub

Private Sub Form_Load()

On Error GoTo NotificaError

    Me.Icon = frmMenuPrincipal.Icon
    
    Call pLlenarCboDepartamento
    
    dtpFechaInicial.Value = DateSerial(Year(Date), Month(Date), 1) 'fdtmServerFecha
    dtpFechaFinal.Value = DateSerial(Year(Date), IIf(Month(Date) = 12, 1, Month(Date) + 1), 1) 'fdtmServerFecha
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
    Unload Me
    
End Sub
Private Sub pLlenarCboDepartamento()
'----------------------------------------------------------------------------------------
' Llena el combo del departamento para poder filtrar las entradas/salidas
'----------------------------------------------------------------------------------------
    On Error GoTo NotificaError
        
    Dim rsDepartamento As New ADODB.Recordset
    Dim vlstrsql As String
    
    vlstrsql = "Select SMICVEDEPARTAMENTO, VCHDESCRIPCION from NoDepartamento "
    Set rsDepartamento = frsRegresaRs(vlstrsql, adLockOptimistic, adOpenDynamic)
    
    Call pLlenarCboRs(cboDepartamento, rsDepartamento, 0, 1, -1)
    
    rsDepartamento.Close
    
    cboDepartamento.AddItem "<TODOS>", 0
    cboDepartamento.ItemData(0) = -1
    cboDepartamento.ListIndex = 0
'    cboDepartamento.ListIndex = fintLocalizaCbo(cboDepartamento, CStr(vgintNumeroDepartamento)) 'se posiciona en el depto con el que se dio login
    
NotificaError:
    If vgblnExistioError Then
        Exit Sub
    Else
        If Err.Number <> 0 Then
            Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLlenarCboDepartamento"))
            On Error GoTo 0
            vgblnExistioError = True
            Exit Sub
        End If
    End If
End Sub
Private Sub pImprime(vlstrTipo As String)

On Error GoTo NotificaError

    Dim rsSocios As New ADODB.Recordset
    Dim alstrParametros(4) As String
    Dim strParametros As String
    
        
    pInstanciaReporte vgrptReporte, "rptSociosEgresados.rpt"
        
    vgrptReporte.DiscardSavedData
    
    strParametros = IIf(cboDepartamento.ItemData(cboDepartamento.ListIndex) = 0, -1, cboDepartamento.ItemData(cboDepartamento.ListIndex)) & "|" & IIf(optPaciente(0).Value, "A", IIf(optPaciente(1).Value, "I", IIf(optPaciente(2).Value, "E", "E"))) & "|" & CStr(dtpFechaInicial.Value) & "|" & CStr(dtpFechaFinal.Value)
        
    Set rsSocios = frsEjecuta_SP(strParametros, "Sp_SORPTEGRESADOS")
    
    If rsSocios.RecordCount <> 0 Then
            
            alstrParametros(0) = "lblNombreHospital;" & Trim(vgstrNombreHospitalCH)
            alstrParametros(1) = UCase("FechaIni;" & CStr(Format(dtpFechaInicial.Value, "dd/MMM/yyyy")))
            alstrParametros(2) = UCase("FechaFin;" & CStr(Format(dtpFechaFinal.Value, "dd/MMM/yyyy")))
            If cboDepartamento.ItemData(cboDepartamento.ListIndex) = -1 Then
            alstrParametros(3) = "strDepartamento;" & "TODOS LOS DEPARTAMENTOS"
            Else
                alstrParametros(3) = "strDepartamento;" & cboDepartamento.List(cboDepartamento.ListIndex)
            End If
            alstrParametros(4) = "strTipoPaciente;" & IIf(optPaciente(0).Value, "PACIENTES INTERNOS Y EXTERNOS", IIf(optPaciente(1).Value, "PACIENTES INTERNOS", IIf(optPaciente(2).Value, "PACIENTES EXTERNOS", "PACIENTES EXTERNOS")))
            pCargaParameterFields alstrParametros, vgrptReporte
            pImprimeReporte vgrptReporte, rsSocios, IIf(vlstrTipo = "P", "P", "I"), "Socios egresados"
    
    Else
        'No existe información con esos parámetros
        MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje"
    End If
    
    rsSocios.Close

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pImprime"))
End Sub

