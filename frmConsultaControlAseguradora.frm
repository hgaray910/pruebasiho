VERSION 5.00
Begin VB.Form frmConsultaControlAseguradora 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control de aseguradora"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10935
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   10935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGrabarRegistro 
      Height          =   495
      Left            =   5220
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmConsultaControlAseguradora.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Guardar la información"
      Top             =   3900
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Height          =   690
      Left            =   5160
      TabIndex        =   30
      Top             =   3750
      Width           =   615
   End
   Begin VB.Frame freControlAseguradora 
      Height          =   3750
      Left            =   75
      TabIndex        =   12
      Top             =   -20
      Width           =   10800
      Begin VB.TextBox txtCartaAutorización 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1860
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   35
         Top             =   240
         Width           =   4755
      End
      Begin VB.TextBox txtNumeroControl 
         Height          =   315
         Left            =   1860
         MaxLength       =   100
         TabIndex        =   34
         ToolTipText     =   "Número de control"
         Top             =   2250
         Width           =   4755
      End
      Begin VB.TextBox txtNumeroPoliza 
         Height          =   315
         Left            =   1860
         MaxLength       =   100
         TabIndex        =   33
         ToolTipText     =   "Número de póliza"
         Top             =   1920
         Width           =   4755
      End
      Begin VB.ComboBox cboTipoPoliza 
         Height          =   315
         Left            =   1860
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   32
         ToolTipText     =   "Tipo de póliza"
         Top             =   1590
         Width           =   4755
      End
      Begin VB.TextBox txtCoaseguroMedico 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   8820
         MaxLength       =   20
         TabIndex        =   28
         Top             =   1935
         Width           =   1845
      End
      Begin VB.TextBox txtCopago 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   8820
         MaxLength       =   20
         TabIndex        =   10
         Top             =   2625
         Width           =   1845
      End
      Begin VB.TextBox txtObservaciones 
         Height          =   1095
         Left            =   1860
         MaxLength       =   1000
         MultiLine       =   -1  'True
         TabIndex        =   3
         ToolTipText     =   "Observaciones"
         Top             =   2580
         Width           =   4755
      End
      Begin VB.TextBox txtTotalSeguro 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8820
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   3075
         Width           =   1845
      End
      Begin VB.TextBox txtParentesco 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1860
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   1
         Top             =   915
         Width           =   4755
      End
      Begin VB.TextBox txtNombreAsegurado 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1860
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   0
         Top             =   570
         Width           =   4755
      End
      Begin VB.TextBox txtControlPersonaAutoriza 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1860
         MaxLength       =   100
         TabIndex        =   2
         Top             =   1260
         Width           =   4755
      End
      Begin VB.TextBox txtDeducible 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   8820
         MaxLength       =   20
         TabIndex        =   7
         Top             =   1245
         Width           =   1845
      End
      Begin VB.TextBox txtCoaseguro 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   8820
         MaxLength       =   20
         TabIndex        =   8
         Top             =   1590
         Width           =   1845
      End
      Begin VB.TextBox txtSumaAsegurada 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   8820
         MaxLength       =   20
         TabIndex        =   5
         Top             =   555
         Width           =   1845
      End
      Begin VB.TextBox txtHonorarios 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   8820
         MaxLength       =   20
         TabIndex        =   4
         Top             =   210
         Width           =   1845
      End
      Begin VB.TextBox txtExcedenteSuma 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   8820
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   900
         Width           =   1845
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Carta de autorización"
         Height          =   195
         Left            =   150
         TabIndex        =   36
         Top             =   300
         Width           =   1500
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Coaseguro médico"
         Height          =   195
         Left            =   6840
         TabIndex        =   29
         Top             =   1995
         Width           =   1320
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Coaseguro adicional"
         Height          =   195
         Left            =   6825
         TabIndex        =   27
         Top             =   2340
         Width           =   1440
      End
      Begin VB.Label lblCoaseguroAdicional 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   315
         Left            =   8820
         TabIndex        =   9
         Top             =   2280
         Width           =   1845
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Número de control"
         Height          =   195
         Left            =   150
         TabIndex        =   26
         Top             =   2355
         Width           =   1305
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Número de póliza"
         Height          =   195
         Left            =   150
         TabIndex        =   25
         Top             =   2010
         Width           =   1230
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de póliza"
         Height          =   195
         Left            =   150
         TabIndex        =   24
         Top             =   1665
         Width           =   990
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Copago"
         Height          =   195
         Left            =   6825
         TabIndex        =   23
         Top             =   2685
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones"
         Height          =   195
         Left            =   150
         TabIndex        =   22
         Top             =   2700
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total conceptos de seguro"
         Height          =   195
         Left            =   6825
         TabIndex        =   21
         Top             =   3135
         Width           =   1905
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "Parentesco"
         Height          =   195
         Left            =   150
         TabIndex        =   20
         Top             =   975
         Width           =   810
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "Nombre del asegurado"
         Height          =   195
         Left            =   150
         TabIndex        =   19
         Top             =   630
         Width           =   1605
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         Caption         =   "Persona que autoriza"
         Height          =   195
         Left            =   150
         TabIndex        =   18
         Top             =   1320
         Width           =   1500
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         Caption         =   "Deducible"
         Height          =   195
         Left            =   6825
         TabIndex        =   17
         Top             =   1305
         Width           =   720
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         Caption         =   "Coaseguro"
         Height          =   195
         Left            =   6825
         TabIndex        =   16
         Top             =   1650
         Width           =   765
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         Caption         =   "Suma asegurada"
         Height          =   195
         Left            =   6825
         TabIndex        =   15
         Top             =   615
         Width           =   1200
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         Caption         =   "Honorarios médicos"
         Height          =   195
         Left            =   6825
         TabIndex        =   14
         Top             =   270
         Width           =   1395
      End
      Begin VB.Label Label75 
         AutoSize        =   -1  'True
         Caption         =   "Excedente"
         Height          =   195
         Left            =   6825
         TabIndex        =   13
         Top             =   960
         Width           =   765
      End
   End
End
Attribute VB_Name = "frmConsultaControlAseguradora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------
'| Nombre del Proyecto      : Caja                                                   -
'| Nombre del Formulario    : frmConsultaControlAseguradora                          -
'-------------------------------------------------------------------------------------
'| Objetivo: Consultar el control de la aseguradora despues de facturar
'-------------------------------------------------------------------------------------
'| Análisis y Diseño        : Rodolfo Ramos G.
'| Autor                    : Rodolfo Ramos G.
'| Fecha de Creación        : 14/May/2003
'| Modificó                 : Nombre(s)
'| Fecha última modificación: 15/May/2003
'-------------------------------------------------------------------------------------
Option Explicit
Public vglngMovPaciente As Long
Public vgstrInternoExterno As String
Public vgstrTipoFactura As String
Public vgstrFolioFactura As String
Public vllngNumeroOpcionAseguradora As Long
Public vglngCveCarta As Long

Dim intNumCuenta As Long




Private Sub pLlenarTipoPoliza()
On Error GoTo NotificaError
    
    Dim vlstrSentencia As String
    Dim rsTipoPoliza As New ADODB.Recordset

    vlstrSentencia = "SELECT * FROM ADTIPOPOLIZASEGURO ORDER BY VCHDESCRIPCION"
    Set rsTipoPoliza = frsRegresaRs(vlstrSentencia)
    
    pLlenarCboRs cboTipoPoliza, rsTipoPoliza, 0, 1
    cboTipoPoliza.AddItem " "
    cboTipoPoliza.ListIndex = 0


Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLlenaClasificacion"))
End Sub
Private Sub pMayusculas(ByRef KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboTipoPoliza_KeyPress(KeyAscii As Integer)
    pMayusculas KeyAscii
    If KeyAscii = 13 Then txtNumeroPoliza.SetFocus
End Sub


Private Sub cmdGrabarRegistro_Click()
      On Error GoTo NotificaError
      
      Dim rs As New ADODB.Recordset
      Dim rsAseguradora As New ADODB.Recordset
      Dim rsPoliza As New ADODB.Recordset
      Dim rsControlAseguradora As New ADODB.Recordset
      Dim strSentencia As String
      Dim lngPersonaGraba As Long
      Dim intCveTipoPoliza As Integer
      Dim lblntipoPoliza As Boolean
      Dim lngCartaDefault As Long
      
     If fblnRevisaPermiso(vglngNumeroLogin, IIf(cgstrModulo = "PV", 4059, 4114), "C", True) Or fblnRevisaPermiso(vglngNumeroLogin, IIf(cgstrModulo = "PV", 4059, 4114), "E", True) Then
   
         lngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
      
         strSentencia = "select ADTIPOPOLIZASEGURO.INTIDTIPOPOLIZA from ADTIPOPOLIZASEGURO where ADTIPOPOLIZASEGURO.VCHDESCRIPCION = " & "'" & cboTipoPoliza.Text & "'"
         Set rsPoliza = frsRegresaRs(strSentencia, adLockReadOnly, adOpenForwardOnly)
         If rsPoliza.RecordCount > 0 Then
             intCveTipoPoliza = rsPoliza!intIdTipoPoliza
             lblntipoPoliza = True
         Else
             lblntipoPoliza = False
         End If
         rsPoliza.Close
      
        lngCartaDefault = flngCartaDefault(vglngMovPaciente)
        If lngCartaDefault = vglngCveCarta Then
            strSentencia = "select expacienteingreso.* from expacienteingreso where expacienteingreso.intNumCuenta = " & intNumCuenta
            Set rsAseguradora = frsRegresaRs(strSentencia, adLockOptimistic, adOpenDynamic)
            With rsAseguradora
                !VCHNUMPOLIZA = txtNumeroPoliza.Text
                !VCHNUMAFILIACION = txtNumeroControl.Text
                If lblntipoPoliza = True Then
                    !intCveTipoPoliza = intCveTipoPoliza
                Else
                    !intCveTipoPoliza = Null
                End If
                .Update
            End With
            rsAseguradora.Close
        End If
        
         strSentencia = "select chrComentarios, VCHNUMPOLIZA, VCHNUMAFILIACION, intCveTipoPoliza from pvControlAseguradora " & _
                        " where intMovPaciente = " & Trim(Str(vglngMovPaciente)) & _
                        " and chrTipoPaciente = '" & Trim(vgstrInternoExterno) & "'" & _
                        " and intCveCarta = " & IIf(vglngCveCarta = 0, "Is Null", " = " & vglngCveCarta)

        Set rsControlAseguradora = frsRegresaRs(strSentencia, adLockOptimistic, adOpenDynamic)
        With rsControlAseguradora
            !chrComentarios = txtObservaciones.Text
            !VCHNUMPOLIZA = Trim(txtNumeroPoliza.Text)
            !VCHNUMAFILIACION = txtNumeroControl.Text
            If lblntipoPoliza = True Then
                !intCveTipoPoliza = intCveTipoPoliza
            Else
                !intCveTipoPoliza = Null
            End If
            .Update
        End With
        rsControlAseguradora.Close
        pGuardarLogTransaccion Me.Name, EnmCambiar, vglngNumeroLogin, "CONTROL DE ASEGURADORA", vgstrFolioFactura
    Else
        MsgBox SIHOMsg(65), vbOKOnly + vbExclamation, "Mensaje"
    End If


Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdSave_Click"))
End Sub


Private Sub Form_Activate()
    cboTipoPoliza.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Dim vlstrSentencia As String
    Dim rs As New ADODB.Recordset
    Dim rsPaciente As New ADODB.Recordset
    Dim lngCartaDefault As Long
   
    Me.Icon = frmMenuPrincipal.Icon
    pLlenarTipoPoliza

    vlstrSentencia = "select * from pvCartaControlSeguro " & _
                        " where intCveCarta " & IIf(vglngCveCarta = 0, "Is Null", " = " & vglngCveCarta)

    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    If rs.RecordCount > 0 Then
        txtCartaAutorización.Text = rs!vchDescripcion
    End If
    
    vlstrSentencia = "select * from pvControlAseguradora " & _
                        " where intMovPaciente = " & Trim(Str(vglngMovPaciente)) & _
                        " and chrTipoPaciente = '" & Trim(vgstrInternoExterno) & "'" & _
                        " and intCveCarta " & IIf(vglngCveCarta = 0, "Is Null", " = " & vglngCveCarta)

    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    If rs.RecordCount > 0 Then
        txtNombreAsegurado.Text = IIf(IsNull(rs!chrNombreAsegurado), " ", rs!chrNombreAsegurado)
        txtParentesco.Text = rs!chrParentesco
        txtSumaAsegurada.Text = FormatCurrency(rs!MNYSUMAASEGURADA, 2)
        txtExcedenteSuma.Text = FormatCurrency(rs!MNYEXCEDENTESUMAASEGURADA, 2)
        txtHonorarios.Text = FormatCurrency(rs!mnyHonorarios, 2)
        txtDeducible.Text = FormatCurrency(rs!MNYCANTIDADDEDUCIBLE, 2)
        txtCoaseguro.Text = FormatCurrency(rs!MNYCANTIDADCOASEGURO, 2)
        txtCoaseguroMedico.Text = FormatCurrency(rs!MNYCANTIDADCOASEGUROMEDICO, 2)
        lblCoaseguroAdicional.Caption = FormatCurrency(rs!MNYCANTIDADCOASEGUROADICIONAL, 2)
        txtCopago.Text = FormatCurrency(rs!MNYCANTIDADCOPAGO, 2)
        txtTotalSeguro.Text = FormatCurrency(rs!MNYCANTIDADDEDUCIBLE + rs!MNYCANTIDADCOASEGURO + rs!MNYCANTIDADCOASEGUROMEDICO + rs!MNYCANTIDADCOASEGUROADICIONAL + rs!MNYCANTIDADCOPAGO + rs!MNYEXCEDENTESUMAASEGURADA, 2)
        txtObservaciones.Text = IIf(IsNull(rs!chrComentarios), " ", rs!chrComentarios)
    
        If vgstrInternoExterno = "I" Then
            vgstrParametrosSP = Str(vglngMovPaciente) & "|" & Str(vgintClaveEmpresaContable)
            Set rsPaciente = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELINTERNOFACTURA")
        Else
            vgstrParametrosSP = Str(vglngMovPaciente) & "|" & Str(vgintClaveEmpresaContable)
            Set rsPaciente = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELEXTERNOFACTURA")
        End If
        If rsPaciente.RecordCount <> 0 Then
            intNumCuenta = CLng(rsPaciente!intNumCuenta)
            txtControlPersonaAutoriza.Text = IIf(IsNull(rsPaciente!Autoriza), "", rsPaciente!Autoriza)
            If Not IsNull(rsPaciente!NombrePoliza) Then cboTipoPoliza.Text = rsPaciente!NombrePoliza
            txtNumeroPoliza.Text = IIf(IsNull(rsPaciente!NumeroPoliza), "", rsPaciente!NumeroPoliza)
            txtNumeroControl.Text = IIf(IsNull(rsPaciente!NumeroControl), "", rsPaciente!NumeroControl)
        End If
        rsPaciente.Close
        
        lngCartaDefault = flngCartaDefault(vglngMovPaciente)
        If lngCartaDefault <> vglngCveCarta Then
            txtControlPersonaAutoriza.Text = IIf(IsNull(rs!VCHAUTORIZACION), "", rs!VCHAUTORIZACION)
            If Not IsNull(rs!intCveTipoPoliza) Then cboTipoPoliza.ListIndex = flngLocalizaCbo(cboTipoPoliza, Str(rs!intCveTipoPoliza))
            txtNumeroPoliza.Text = IIf(IsNull(rs!VCHNUMPOLIZA), "", rs!VCHNUMPOLIZA)
            txtNumeroControl.Text = IIf(IsNull(rs!VCHNUMAFILIACION), "", rs!VCHNUMAFILIACION)
            
            
'            !VCHNUMPOLIZA = Trim(txtNumeroPoliza.Text)
'            !VCHNUMAFILIACION = txtNumeroControl.Text
'            If lblntipoPoliza = True Then
'                !intCveTipoPoliza = intCveTipoPoliza
'            Else
'                !intCveTipoPoliza = Null
'            End If
            
        End If
    End If
    rs.Close
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '¿Desea abandonar la operación?
    If MsgBox(SIHOMsg(17), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
        Unload Me
    Else
        Cancel = True
        cboTipoPoliza.SetFocus
    End If
End Sub



Private Sub txtNumeroControl_GotFocus()
    pSelTextBox txtNumeroControl
End Sub

Private Sub txtNumeroControl_KeyPress(KeyAscii As Integer)
    pMayusculas KeyAscii
    If KeyAscii = 13 Then txtObservaciones.SetFocus
End Sub


Private Sub txtNumeroPoliza_GotFocus()
    pSelTextBox txtNumeroPoliza
End Sub

Private Sub txtNumeroPoliza_KeyPress(KeyAscii As Integer)
    pMayusculas KeyAscii
    If KeyAscii = 13 Then txtNumeroControl.SetFocus
End Sub


Private Sub txtObservaciones_GotFocus()
    pSelTextBox txtObservaciones
End Sub

Private Sub txtObservaciones_KeyPress(KeyAscii As Integer)
    pMayusculas KeyAscii
    If KeyAscii = 13 Then cmdGrabarRegistro.SetFocus
End Sub


