VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmRptVentaClientes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ventas a crédito"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7290
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   7290
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Height          =   675
      Left            =   60
      TabIndex        =   14
      Top             =   -30
      Width           =   7125
      Begin VB.ComboBox cboHospital 
         Height          =   315
         Left            =   915
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Seleccione la empresa"
         Top             =   240
         Width           =   6015
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Empresa"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   300
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   60
      TabIndex        =   13
      Top             =   600
      Width           =   7125
      Begin VB.ComboBox cboTipoCte 
         Height          =   315
         ItemData        =   "frmRptVentaClientes.frx":0000
         Left            =   1290
         List            =   "frmRptVentaClientes.frx":0013
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   630
         Width           =   5640
      End
      Begin VB.ComboBox cboDepto 
         Height          =   315
         Left            =   1290
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   225
         Width           =   5640
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   690
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Departamento"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   270
         Width           =   1005
      End
   End
   Begin VB.CheckBox chkCuantaSaldada 
      Caption         =   "Incluir cuenta saldada"
      Height          =   195
      Left            =   60
      TabIndex        =   8
      Top             =   2715
      Width           =   2010
   End
   Begin VB.Frame Frame6 
      Height          =   780
      Left            =   4515
      TabIndex        =   12
      Top             =   1815
      Width           =   2670
      Begin VB.CommandButton cmdExportar 
         Caption         =   "Exportar"
         Height          =   495
         Left            =   1080
         TabIndex        =   7
         ToolTipText     =   "Exportar los datos"
         Top             =   165
         Width           =   1500
      End
      Begin VB.CommandButton cmdPrint 
         Height          =   495
         Left            =   570
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmRptVentaClientes.frx":0056
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Imprimir"
         Top             =   165
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdPreview 
         Height          =   495
         Left            =   75
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmRptVentaClientes.frx":0423
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Vista previa"
         Top             =   165
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Rango de fechas"
      Height          =   780
      Left            =   60
      TabIndex        =   9
      Top             =   1815
      Width           =   4425
      Begin MSComCtl2.DTPicker dtmFecFin 
         Height          =   330
         Left            =   2850
         TabIndex        =   4
         ToolTipText     =   "Fecha final"
         Top             =   285
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   582
         _Version        =   393216
         Format          =   57999361
         CurrentDate     =   38058
      End
      Begin MSComCtl2.DTPicker dtmFecIni 
         Height          =   330
         Left            =   825
         TabIndex        =   3
         ToolTipText     =   "Fecha inicial"
         Top             =   270
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   582
         _Version        =   393216
         Format          =   57999361
         CurrentDate     =   38058
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         Height          =   195
         Left            =   2310
         TabIndex        =   11
         Top             =   330
         Width           =   420
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Left            =   135
         TabIndex        =   10
         Top             =   330
         Width           =   465
      End
   End
   Begin MSComDlg.CommonDialog CDgArchivo 
      Left            =   6765
      Top             =   1500
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   ".txt"
      DialogTitle     =   "Exportación de ventas"
      FileName        =   "Ventas.txt"
      Filter          =   "Texto (*.txt)|*.txt| Todos los archivos (*.*)|*.*"
   End
End
Attribute VB_Name = "frmRptVentaClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private vgrptReporte As CRAXDRT.Report

Dim lFecIni As String
Dim lFecFin As String
Dim lTipoCliente As String

Private Sub cboHospital_Click()
    On Error GoTo NotificaError
    
    Dim rs As New ADODB.Recordset

    If cboHospital.ListIndex <> -1 Then
        cboDepto.Clear
        vgstrParametrosSP = "-1|1|*|" & CStr(cboHospital.ItemData(cboHospital.ListIndex))
        Set rs = frsEjecuta_SP(vgstrParametrosSP, "Sp_Gnseldepartamento")
        If rs.RecordCount <> 0 Then
            pLlenarCboRs cboDepto, rs, 0, 1
        End If
        cboDepto.AddItem "<TODOS>", 0
        cboDepto.ItemData(cboDepto.NewIndex) = -1
        cboDepto.ListIndex = 0
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboHospital_Click"))
    Unload Me
End Sub

Private Sub cmdExportar_Click()
  On Error GoTo ErrHandler
  CDgArchivo.CancelError = True
  CDgArchivo.InitDir = App.Path
  CDgArchivo.Flags = cdlOFNOverwritePrompt
  CDgArchivo.ShowSave
  pExportar
Exit Sub
ErrHandler:
  'El usuario presionó el botón de cancelar
  Err.Number = 0
End Sub

Private Sub cmdPreview_Click()
  pImprime "P"
End Sub

Private Sub cmdPrint_Click()
  pImprime "I"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
    SendKeys vbTab
  ElseIf KeyCode = vbKeyEscape Then
    Unload Me
  End If
End Sub

Private Sub Form_Load()

    Dim lngNumOpcion As Long

    Me.Icon = frmMenuPrincipal.Icon
    
    Select Case cgstrModulo
    Case "PV"
         lngNumOpcion = 1786
    Case "SE"
         lngNumOpcion = 2010
    End Select
    
    pCargaHospital lngNumOpcion
  
    dtmFecIni = fdtmServerFecha
    dtmFecFin = dtmFecIni
    
    cboTipoCte.ListIndex = 0


End Sub

Private Sub pImprime(vlstrDispositivo As String)
    Dim rsReporte As New ADODB.Recordset
    Dim alstrParametros(1) As String
     
    lFecIni = fstrFechaSQL(dtmFecIni, " 00:00:00")
    lFecFin = fstrFechaSQL(dtmFecFin, " 23:59:59")
    
    lTipoCliente = fstrRegresaTipoCliente
    
    vgstrParametrosSP = lFecIni & "|" & lFecFin & "|" & lTipoCliente & _
                        "|" & cboDepto.ItemData(cboDepto.ListIndex) & _
                        "|" & IIf(chkCuantaSaldada.Value = 0, 0, -1) & _
                        "|" & Str(cboHospital.ItemData(cboHospital.ListIndex))
    
    Set rsReporte = frsEjecuta_SP(vgstrParametrosSP, "Sp_PvSelVentaClientes")
    
    If rsReporte.RecordCount > 0 Then
        pInstanciaReporte vgrptReporte, "rptPvVentasCredito.rpt"
        vgrptReporte.DiscardSavedData
        alstrParametros(0) = "Empresa;" & Trim(cboHospital.List(cboHospital.ListIndex))
        alstrParametros(1) = "Filtro;" & "DEL " & Format(dtmFecIni, "DD/MMM/YYYY") & " A " & Format(dtmFecFin, "DD/MMM/YYYY") & " DEL DEPARTAMENTO " & cboDepto.Text & " DEL TIPO DE CLIENTE " & cboTipoCte.Text
        pCargaParameterFields alstrParametros, vgrptReporte
        pImprimeReporte vgrptReporte, rsReporte, vlstrDispositivo, "Ventas a crédito"
    Else
        MsgBox SIHOMsg(13), vbInformation + vbOKOnly
    End If
    
End Sub
Private Function fstrRegresaTipoCliente() As String
  Select Case cboTipoCte
  Case "EMPLEADO"
    fstrRegresaTipoCliente = "'EM'"
  Case "EMPRESA"
    fstrRegresaTipoCliente = "'CO'"
  Case "PACIENTE INTERNO"
    fstrRegresaTipoCliente = "'PI'"
  Case "PACIENTE EXTERNO"
    fstrRegresaTipoCliente = "'PE'"
  Case "MEDICO"
    fstrRegresaTipoCliente = "'ME'"
  End Select
End Function

Private Sub pExportar()
  Dim rs As New ADODB.Recordset
  'Set rs = frsCargaRs
  Dim vlstrCadena As String
  
  
  lFecIni = fstrFechaSQL(dtmFecIni, " 00:00:00")
  lFecFin = fstrFechaSQL(dtmFecFin, " 23:59:59")
    
  lTipoCliente = fstrRegresaTipoCliente
    
  vgstrParametrosSP = lFecIni & "|" & lFecFin & "|" & lTipoCliente & _
                        "|" & cboDepto.ItemData(cboDepto.ListIndex) & _
                        "|" & IIf(chkCuantaSaldada.Value = 0, 0, -1)
    
  Set rs = frsEjecuta_SP(vgstrParametrosSP, "Sp_PvSelVentaClientes")
  
  
  With rs
    If .State <> adStateClosed Then
      If .RecordCount = 0 Then
          ' No existe información
        MsgBox SIHOMsg(12), vbOKOnly + vbExclamation, "Exportación de poliza(s)"
      Else
        Open CDgArchivo.FileName For Output As #1  ' Open file for output.
          .MoveFirst
          While Not .EOF
            vlstrCadena = ""                                                                                                ' NADA
            vlstrCadena = vlstrCadena & "" & Format(!Fecha, "YYYY/MM/DD") & " " & Format(!Fecha, "HH:MM")                   ' FECHA y HORA
            vlstrCadena = vlstrCadena & " " & Format(!NumCliente, "0000000000")                                             ' NUMCTE
            vlstrCadena = vlstrCadena & " " & !Cliente & String(60 - Len(!Cliente), " ")                                    ' CLIENTE
            vlstrCadena = vlstrCadena & " " & String(12 - Len(Format(!Monto, "###0.00")), " ") & Format(!Monto, "###0.00")  ' MONTO
            vlstrCadena = vlstrCadena & " " & Format(!Referencia, "0000000000")                                             ' NUMCTE
            vlstrCadena = vlstrCadena & " " '
            Print #1, vlstrCadena ' Detalle o Movimientos de la Póliza
            .MoveNext
          Wend
          Close #1 ' Close file.
          '¡Los datos han sido guardados satisfactoriamente!
          MsgBox SIHOMsg(358), vbOKOnly + vbInformation, "Exportación de poliza(s)"
        End If
        .Close
    End If
  End With
End Sub

Private Sub pCargaHospital(lngNumOpcion As Long)
    On Error GoTo NotificaError
    
    Dim rs As New ADODB.Recordset
    
    Set rs = frsEjecuta_SP("-1", "Sp_Gnselempresascontable")
    If rs.RecordCount <> 0 Then
        pLlenarCboRs cboHospital, rs, 1, 0
        cboHospital.ListIndex = flngLocalizaCbo(cboHospital, Str(vgintClaveEmpresaContable))
    End If
    
    cboHospital.Enabled = fblnRevisaPermiso(vglngNumeroLogin, lngNumOpcion, "C")
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCargaHospital"))
    Unload Me
End Sub


