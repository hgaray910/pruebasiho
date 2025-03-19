VERSION 5.00
Begin VB.Form frmCartaSeguro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Carta de autorización"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7890
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   7890
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   7655
      Begin VB.ComboBox cboEmpresa 
         Height          =   315
         Left            =   1080
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Empresa a la que corresponde la carta de autorización"
         Top             =   720
         Width           =   6375
      End
      Begin VB.TextBox txtNombre 
         Height          =   315
         Left            =   1080
         MaxLength       =   100
         TabIndex        =   0
         ToolTipText     =   "Nombre de la carta de autorización"
         Top             =   240
         Width           =   6375
      End
      Begin VB.Label Label2 
         Caption         =   "Empresa"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   735
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   270
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Height          =   675
      Left            =   3660
      TabIndex        =   7
      Top             =   1560
      Width           =   600
      Begin VB.CommandButton cmdImportacion 
         Caption         =   "I&mportar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2475
         TabIndex        =   3
         ToolTipText     =   "Importar póliza"
         Top             =   135
         Width           =   1380
      End
      Begin VB.CommandButton cmdGrabar 
         Height          =   495
         Left            =   50
         Picture         =   "frmCartaSeguro.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Guardar carta"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmCartaSeguro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------
'| Nombre del Proyecto      : Caja
'| Nombre del Formulario    : frmCartaSeguro
'-------------------------------------------------------------------------------------
'| Objetivo:    Generación de cartas de autorización para las aseguradoras del paciente
'|
'-------------------------------------------------------------------------------------
'| Análisis y Diseño        : Betsi de la Torre y Mayra Armendáriz
'| Autor                    : Mayra Armendáriz
'| Fecha de Creación        : 07/Julio/2022
'-------------------------------------------------------------------------------------
'| Últimas modificaciones, especificar:
'-------------------------------------------------------------------------------------
' Fecha:
' Descripción:
' Autor:
'-------------------------------------------------------------------------------------
Option Explicit

Dim vlstrSentencia As String
Public vgNumCuentaPaciente As String
Dim rs As ADODB.Recordset


Private Sub cmdGrabar_Click()
On Error GoTo NotificaError
    Dim vllngPersonaGraba As Long
    Dim vllngCveCarta As Long
    
    If fblnDatosValidos Then
        vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
        
        If vllngPersonaGraba <> 0 Then
            vlstrSentencia = "Select * from PVCARTACONTROLSEGURO where INTCVECARTA = -1"
            Set rs = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
            With rs
                .AddNew
                !VCHDESCRIPCION = txtNombre.Text
                !intNumCuenta = vgNumCuentaPaciente
                !intcveempresa = cboEmpresa.ItemData(cboEmpresa.ListIndex)
                !BITDEFAULT = 0
                .Update
                vllngCveCarta = flngObtieneIdentity(UCase("SEQ_PVCARTACONTROLSEGURO"), !intCveCarta)
            End With
        '    If chkDefault.Value = 1 Then
        '
        '        pEjecutaSentencia ("UPDATE PvCargo set intcvecarta = " & vllngCveCarta & " where intmovpaciente = " & vgNumCuentaPaciente)
        '    End If
            
            MsgBox SIHOMsg(358), vbOKOnly + vbInformation, "Mensaje"
            'Limpiar información
            txtNombre.Text = ""
            txtNombre.SetFocus
            cboEmpresa.ListIndex = -1
            Call pGuardarLogTransaccion(Me.Name, EnmGrabar, vllngPersonaGraba, "CARTA DE AUTORIZACIÓN DE ASEGURADORA", vgNumCuentaPaciente & " " & vllngCveCarta)
        '    chkDefault.Value = 0
            'pCartaDefecto
        End If
    End If
       
    
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
End Sub


Private Sub Form_Load()
On Error GoTo NotificaError
    Dim rsEmpresa As ADODB.Recordset
    Me.Icon = frmMenuPrincipal.Icon
    vlstrSentencia = "select distinct ccEmpresa.intCveEmpresa, ccEmpresa.vchDescripcion from CcEmpresa " & _
                 " inner join CcTipoConvenio on CcEmpresa.TNYCVETIPOCONVENIO = CcTipoConvenio.TNYCVETIPOCONVENIO " & _
                              " where CcTipoConvenio.bitAseguradora =1 "

    Set rsEmpresa = frsRegresaRs(vlstrSentencia)
    pLlenarCboRs cboEmpresa, rsEmpresa, 0, 1
    cboEmpresa.ListIndex = flngLocalizaCbo(cboEmpresa, frmFacturacion.vgintEmpresa)
'    pCartaDefecto
   
   ' vgintEmpresa
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    If KeyAscii = 27 Then Unload Me
    If KeyAscii = 13 Then SendKeys vbTab
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyPress"))
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
     KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

'Private Sub pCartaDefecto()
'     'revisar si el paciente ya tiene una carta por defecto, solo puede agregarse 1
'    vlstrSentencia = "Select Count(*) cartasDefault  from PVCARTACONTROLSEGURO where BITDEFAULT = 1 and intNumCuenta=" & vgNumCuentaPaciente
'    Set rs = frsRegresaRs(vlstrSentencia)
'    If rs!cartasDefault > 0 Then
'        chkDefault.Enabled = False
'    End If
'End Sub
Private Function fblnDatosValidos() As Boolean
    fblnDatosValidos = True
    'Que se haya agregado descripcion:
    If txtNombre.Text = "" Then
        fblnDatosValidos = False
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbExclamation, "Mensaje"
        pEnfocaTextBox txtNombre
    End If
    
    If fblnDatosValidos And cboEmpresa.ListIndex = -1 Then
        fblnDatosValidos = False
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbExclamation, "Mensaje"
        cboEmpresa.SetFocus
    End If
End Function

