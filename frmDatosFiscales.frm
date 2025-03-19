VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmDatosFiscales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datos fiscales"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8850
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   8850
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab sstDatos 
      Height          =   5630
      Left            =   0
      TabIndex        =   17
      Top             =   -600
      Width           =   8910
      _ExtentX        =   15716
      _ExtentY        =   9922
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmDatosFiscales.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "freDatosFiscales"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmDatosFiscales.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "freBuscarDatos"
      Tab(1).ControlCount=   1
      Begin VB.Frame freDatosFiscales 
         Height          =   4115
         Left            =   135
         TabIndex        =   18
         Top             =   600
         Width           =   8640
         Begin VB.CheckBox chkRazonSocial 
            Caption         =   "Usar razón social"
            Enabled         =   0   'False
            Height          =   255
            Left            =   6200
            TabIndex        =   47
            ToolTipText     =   "Usar razón social"
            Top             =   270
            Width           =   1695
         End
         Begin VB.CheckBox chkRFCgenerico 
            Caption         =   "RFC genérico"
            Height          =   255
            Left            =   3360
            TabIndex        =   46
            ToolTipText     =   "Capturar RFC genérico"
            Top             =   270
            Width           =   1455
         End
         Begin VB.ComboBox cboRegimenFiscal 
            Height          =   315
            Left            =   1560
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   2
            ToolTipText     =   "Selección del régimen fiscal"
            Top             =   630
            Width           =   6930
         End
         Begin VB.TextBox txtCorreo 
            Height          =   315
            Left            =   1560
            MaxLength       =   100
            TabIndex        =   12
            ToolTipText     =   "Correo electrónico"
            Top             =   3330
            Width           =   6930
         End
         Begin VB.ComboBox cboUsoCFDI 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   3720
            Width           =   4815
         End
         Begin VB.CheckBox chkIEPS 
            Caption         =   "Sujeto a IEPS"
            Height          =   255
            Left            =   6840
            TabIndex        =   11
            Top             =   2970
            Width           =   1335
         End
         Begin VB.CheckBox chkExtranjero 
            Caption         =   "Extranjero"
            Height          =   195
            Left            =   4920
            TabIndex        =   1
            Top             =   300
            Width           =   1095
         End
         Begin VB.TextBox txtNumInterior 
            Height          =   315
            Left            =   4800
            MaxLength       =   30
            TabIndex        =   6
            Top             =   1800
            Width           =   1560
         End
         Begin VB.TextBox txtNumExterior 
            Height          =   315
            Left            =   1560
            MaxLength       =   30
            TabIndex        =   5
            Top             =   1800
            Width           =   1560
         End
         Begin VB.TextBox txtColonia 
            Height          =   315
            Left            =   1560
            MaxLength       =   100
            TabIndex        =   7
            Top             =   2180
            Width           =   6930
         End
         Begin VB.TextBox txtCP 
            Height          =   315
            Left            =   1560
            MaxLength       =   5
            TabIndex        =   9
            Top             =   2940
            Width           =   1560
         End
         Begin VB.ComboBox cboCiudad 
            Height          =   315
            Left            =   1560
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   8
            ToolTipText     =   "Seleccione la ciudad"
            Top             =   2550
            Width           =   4800
         End
         Begin VB.TextBox txtRFC 
            Height          =   315
            Left            =   1560
            MaxLength       =   13
            TabIndex        =   0
            Top             =   240
            Width           =   1560
         End
         Begin VB.TextBox txtNombreFactura 
            Height          =   315
            Left            =   1560
            MaxLength       =   300
            TabIndex        =   3
            Top             =   1020
            Width           =   6930
         End
         Begin VB.TextBox txtDireccionFactura 
            Height          =   315
            Left            =   1560
            MaxLength       =   250
            TabIndex        =   4
            Top             =   1410
            Width           =   6930
         End
         Begin VB.TextBox txtTelefonoFactura 
            Height          =   315
            Left            =   4800
            MaxLength       =   20
            TabIndex        =   10
            Top             =   2940
            Width           =   1560
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Régimen fiscal"
            Height          =   195
            Left            =   225
            TabIndex        =   45
            Top             =   690
            Width           =   1035
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Correo electrónico"
            Height          =   195
            Left            =   225
            TabIndex        =   44
            Top             =   3390
            Width           =   1290
         End
         Begin VB.Label lblUsoCFDI 
            Caption         =   "Uso del CFDI"
            Height          =   195
            Left            =   225
            TabIndex        =   43
            Top             =   3780
            Width           =   975
         End
         Begin VB.Label Label5 
            Caption         =   "Número interior"
            Height          =   255
            Left            =   3600
            TabIndex        =   42
            Top             =   1890
            Width           =   1335
         End
         Begin VB.Label Label4 
            Caption         =   "Número exterior"
            Height          =   255
            Left            =   225
            TabIndex        =   41
            Top             =   1890
            Width           =   1215
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Código postal"
            Height          =   195
            Left            =   225
            TabIndex        =   40
            Top             =   3000
            Width           =   960
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Colonia"
            Height          =   195
            Left            =   225
            TabIndex        =   39
            Top             =   2280
            Width           =   525
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Ciudad"
            Height          =   195
            Left            =   225
            TabIndex        =   38
            Top             =   2640
            Width           =   495
         End
         Begin VB.Label Label55 
            AutoSize        =   -1  'True
            Caption         =   "Razón social"
            Height          =   195
            Left            =   225
            TabIndex        =   22
            Top             =   1080
            Width           =   915
         End
         Begin VB.Label Label54 
            AutoSize        =   -1  'True
            Caption         =   "Calle"
            Height          =   195
            Left            =   225
            TabIndex        =   21
            Top             =   1480
            Width           =   345
         End
         Begin VB.Label Label53 
            AutoSize        =   -1  'True
            Caption         =   "Teléfono"
            Height          =   195
            Left            =   3960
            TabIndex        =   20
            Top             =   3000
            Width           =   630
         End
         Begin VB.Label Label52 
            AutoSize        =   -1  'True
            Caption         =   "RFC"
            Height          =   195
            Left            =   225
            TabIndex        =   19
            Top             =   300
            Width           =   315
         End
      End
      Begin VB.Frame Frame2 
         Height          =   690
         Left            =   3600
         TabIndex        =   37
         Top             =   4770
         Width           =   1620
         Begin VB.CommandButton cmdSave 
            Height          =   495
            Left            =   60
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmDatosFiscales.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Guardar"
            Top             =   135
            Width           =   495
         End
         Begin VB.CommandButton cmdDelete 
            Height          =   495
            Left            =   555
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmDatosFiscales.frx":037A
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Borrar"
            Top             =   135
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdLocate 
            Height          =   495
            Left            =   1050
            Picture         =   "frmDatosFiscales.frx":051C
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Búsqueda"
            Top             =   135
            UseMaskColor    =   -1  'True
            Width           =   495
         End
      End
      Begin VB.Frame freBuscarDatos 
         Height          =   4875
         Left            =   -74880
         TabIndex        =   23
         Top             =   600
         Width           =   8655
         Begin VB.Frame Frame1 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   375
            Left            =   210
            TabIndex        =   33
            Top             =   600
            Width           =   1815
            Begin VB.OptionButton optClave 
               Caption         =   "Nombre"
               Height          =   285
               Index           =   1
               Left            =   960
               Style           =   1  'Graphical
               TabIndex        =   36
               Top             =   0
               Value           =   -1  'True
               Width           =   855
            End
            Begin VB.OptionButton optClave 
               Caption         =   "Clave"
               Height          =   285
               Index           =   0
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   35
               Top             =   0
               Width           =   855
            End
         End
         Begin VB.Frame fraBusquedaClave 
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   1890
            TabIndex        =   31
            Top             =   600
            Width           =   6615
            Begin VB.TextBox txtNumeroReferencia 
               Height          =   285
               Left            =   255
               MaxLength       =   100
               TabIndex        =   32
               Top             =   0
               Width           =   6240
            End
         End
         Begin VB.OptionButton optBusqueda 
            Caption         =   "Médicos"
            Height          =   285
            Index           =   5
            Left            =   5700
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   240
            Width           =   1330
         End
         Begin VB.TextBox txtBuscaDatos 
            Height          =   285
            Left            =   2025
            MaxLength       =   100
            TabIndex        =   29
            Top             =   600
            Width           =   5985
         End
         Begin VB.OptionButton optBusqueda 
            Caption         =   "Empleados"
            Height          =   285
            Index           =   4
            Left            =   4350
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   240
            Width           =   1330
         End
         Begin VB.OptionButton optBusqueda 
            Caption         =   "Internos"
            Height          =   285
            Index           =   3
            Left            =   330
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   240
            Width           =   1330
         End
         Begin VB.OptionButton optBusqueda 
            Caption         =   "Otros"
            Height          =   285
            Index           =   2
            Left            =   7035
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   240
            Width           =   1330
         End
         Begin VB.OptionButton optBusqueda 
            Caption         =   "Empresa"
            Height          =   285
            Index           =   1
            Left            =   3015
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   240
            Width           =   1330
         End
         Begin VB.OptionButton optBusqueda 
            Caption         =   "Externos"
            Height          =   285
            Index           =   0
            Left            =   1665
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   240
            Value           =   -1  'True
            Width           =   1330
         End
         Begin VB.ListBox lstBuscaDatos 
            Height          =   3375
            Left            =   300
            TabIndex        =   34
            Top             =   1080
            Width           =   8070
         End
      End
   End
End
Attribute VB_Name = "frmDatosFiscales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------
'| Nombre del Proyecto      : Caja                                                   -
'| Nombre del Formulario    : frmDatosFiscales                                       -
'-------------------------------------------------------------------------------------
'| Análisis y Diseño        : Rodolfo Ramos G.
'| Autor                    : Rodolfo Ramos G.
'| Fecha de Creación        : 31/Ene/2002
'| Modificó                 : Nombre(s)
'| Fecha Terminación        : hoy
'| Fecha última modificación: 31/Ene/2002
'-------------------------------------------------------------------------------------
Option Explicit

Public vgstrRFC As String
Public vgstrNombre As String
Public vgstrDireccion As String
Public vgstrNumExterior As String
Public vgstrNumInterior As String
Public vgBitExtranjero As Integer
Public vgstrTelefono As String
Public llngCveCiudad As Long
Public vgstrColonia As String
Public vgstrCP As String

Public vlstrRegimenFiscal As String 'Regimen fiscal

Public vglngNumRef As Long, vgstrTipo As String
Public vlstrNumRef As String, vlstrTipo As String
Public vgblnModalResult As Boolean
Public vglngCveDatosFiscales As Long
Public vglngDatosParametro As Boolean 'para indentificar cuando es una venta al publico en general
Public vgblnMostrarUsoCFDI As Boolean
Public vgintUsoCFDI As Long
Public vgstrTipoUsoCFDI As String
Public vgintTipoPacEmp As Integer
Public vgstrCorreo As String

Dim vlintIDConsulta As Long ' si este ID tiene informacion es por que los datos de pantalla vienen de una consulta
Dim vlstrTipoPac As String  ' Tipo de paciente
Dim vlstrNumPac As String   ' Número de paciente
Public vgBitSujetoaIEPS As Integer
Public vgActivaSujetoaIEPS As Boolean
Dim vlblnLicenciaIEPS As Boolean
Dim vlblnLicenciaContaElectronica As Boolean
' Variable de almacenamiento del RFC cuando se revierte el genérico
Dim vlstrRFCprovisional As String

Private Sub chkExtranjero_Click()
    '' caso 6857 para que el sistema pueda capturar un rfc extranjero de forma automatica
    If chkExtranjero.Value = 1 Then ' solo cuando el rfc esta vacio
        chkRazonSocial.Enabled = False
        chkRazonSocial.Value = 0
        If txtRFC.Text = "" Then
            txtRFC.Text = "XEXX010101000"
        End If
    Else
        If chkRFCgenerico.Value = 1 Then
            chkRazonSocial.Enabled = True
        End If
        If txtRFC.Text = "XEXX010101000" Then
            txtRFC.Text = ""
        ElseIf txtRFC.Text = "" Then
               pEnfocaTextBox txtRFC
        Else

        End If
        If vlintIDConsulta = 0 Then
           vlintIDConsulta = -2
        End If
    End If
End Sub

Private Sub chkRazonSocial_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cboRegimenFiscal.SetFocus
End Sub

Private Sub chkRFCgenerico_Click()
    If chkRFCgenerico.Value = vbChecked And chkExtranjero.Value = vbChecked Then
        txtRFC.Enabled = False
        vlstrRFCprovisional = txtRFC.Text
        txtRFC.Text = "XEXX010101000"
'        pValidaBloqueoRegimen
        
    ElseIf chkRFCgenerico.Value = vbChecked And chkExtranjero.Value = vbUnchecked Then
        txtRFC.Enabled = False
        vlstrRFCprovisional = txtRFC.Text
        txtRFC.Text = "XAXX010101000"
        chkRazonSocial.Enabled = True
'        pValidaBloqueoRegimen
    Else
        txtRFC.Enabled = True
        txtRFC.Text = vlstrRFCprovisional
        chkRazonSocial.Enabled = False
        chkRazonSocial.Value = 0
    End If
    pValidaBloqueoRegimen
End Sub

Private Sub chkRFCgenerico_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cboRegimenFiscal.SetFocus
End Sub

Private Sub cmdDelete_Click()
    pEjecutaSentencia "Delete from PvDatosFiscales Where chrRFC = '" & Trim(txtRFC.Text) & "'"
    MsgBox SIHOMsg(420), vbInformation, "Mensaje"
    pLimpiaForma
End Sub
Private Sub cmdLocate_Click()
    sstDatos.Tab = 1
    freBuscarDatos.Enabled = True
    txtBuscaDatos.SetFocus
    Select Case vlstrTipo
      Case "PI"
        optBusqueda(3).Value = True
      Case "PE"
        optBusqueda(0).Value = True
      Case "EM"
        optBusqueda(4).Value = True
      Case "CO"
        optBusqueda(1).Value = True
      Case "ME"
        optBusqueda(5).Value = True
      Case Else
        optBusqueda(2).Value = True
    End Select
End Sub
Private Sub cmdSave_Click()
    Dim rs As New ADODB.Recordset
    Dim vlrsCveDatosFiscales As New ADODB.Recordset
    Dim ObjRS As New ADODB.Recordset
    Dim vlstrsql As String
    Dim vlblnAddUpd As Boolean ' FALSE = inserta PVDATOSFISCALES, TRUE = actualiza PVDATOSFISCALES
    Dim vlblnBanAddUpd As Boolean ' auxiliar de la variable vlblnBanAddUpd cuando se intenta actualizar un registro pero los cambios coinciden con otro registro existente
    Dim vllngRegimenFiscal As Long
    
    If Not fnValidaDatos Then Exit Sub ' Validacion de los datos
    
    If vlintIDConsulta > 0 Then ' Tiene un ID de una consulta en PVDATOSFISCALES
        ' Revisar si coincide el RFC con el ID que se tiene capturado
        vlstrsql = "SELECT IntID FROM PvDatosFiscales WHERE chrRFC = '" & Trim(Me.txtRFC.Text) & "' AND intID = " & vlintIDConsulta & " ORDER BY intID ASC"
        Set ObjRS = frsRegresaRs(vlstrsql, adLockOptimistic)
        If ObjRS.RecordCount = 0 Then ' El RFC cambió
            vlblnBanAddUpd = True
        Else ' Corresponden RFC e ID en la tabla PVDATOSFISCALES
            vlblnBanAddUpd = False
        End If
        
        ' Búsqueda por registro para ver si lo que intentamos insertar no se encuentra en la base de datos
        vlstrsql = "SELECT IntID FROM PvDatosFiscales WHERE chrRFC = '" & Trim(Me.txtRFC.Text) & "'" '& "' AND intID = " & vlintIDConsulta '& " ORDER BY intID ASC"
        
        '----- AGREGADO PARA CASO 7931 -----'
        If Trim(vlstrTipoPac) <> "" Then  ' Pacientes Internos y Externos
            vlstrsql = vlstrsql & " AND chrTipoPaciente = '" & vlstrTipoPac & "'"
        ElseIf Trim(vlstrTipo) <> "" And Trim(vlstrTipo) <> "OT" Then ' Empleados, Médicos y Convenios
            'vlstrsql = vlstrsql & " AND chrTipoCliente = '" & vlstrTipo & "'"  '------ Comentado por caso 8595
        End If
        
        If Trim(vlstrNumPac) <> "" Then   ' Pacientes Internos y Externos
            vlstrsql = vlstrsql & " AND intNumCuenta = " & vlstrNumPac
        ElseIf Trim(vlstrNumRef) <> "NULL" And Val(vlstrNumRef) <> 0 Then ' Empleados, Médicos y Convenios
            'vlstrsql = vlstrsql & " AND intNumReferencia = " & vlstrNumRef   '------ Comentado por caso 8595
        End If
        vlstrsql = vlstrsql & " ORDER BY intID ASC"
        '-----------------------------------'
        
    Else ' No hay ID de PVDATOSFISCALES(NO LO ENCONTRÓ O ES UN REGISTRO NUEVO)
        vlstrsql = "SELECT IntID FROM PvDatosFiscales WHERE chrRFC = '" & fStrRFCValido(txtRFC.Text) & "'"
        vlstrsql = vlstrsql & IIf(chkExtranjero.Value = 1, " AND chrTipoPaciente IS NULL AND intNumCuenta IS NULL", "")
        
        '----- AGREGADO PARA CASO 7931 -----'
        If Trim(vlstrTipoPac) <> "" And chkExtranjero.Value = 0 Then ' Pacientes Internos y Externos
            vlstrsql = vlstrsql & " AND chrTipoPaciente = '" & vlstrTipoPac & "'"
        ElseIf Trim(vlstrTipo) <> "" And Trim(vlstrTipo) <> "OT" Then ' Empleados, Médicos y Convenios
            'vlstrsql = vlstrsql & " AND chrTipoCliente = '" & vlstrTipo & "'"    '------ Comentado por caso 8595
        End If
        
        If Trim(vlstrNumPac) <> "" And chkExtranjero.Value = 0 Then  ' Pacientes Internos y Externos
            vlstrsql = vlstrsql & " AND intNumCuenta = " & vlstrNumPac
        ElseIf Trim(vlstrNumRef) <> "NULL" And Val(vlstrNumRef) <> 0 Then ' Empleados, Médicos y Convenios
            'vlstrsql = vlstrsql & " AND intNumReferencia = " & vlstrNumRef     '------ Comentado por caso 8595
        End If
        '-----------------------------------'
        
        vlstrsql = vlstrsql & " ORDER BY intID ASC"
    End If
     
    Set ObjRS = frsRegresaRs(vlstrsql, adLockOptimistic)
    ' Si encontramos registros activamos la variable vlblnAddUPD para que se haga una actualizacion, si no hay registros deben de insertarse Registros
    vlblnAddUpd = IIf(vlintIDConsulta > 0, IIf(vlblnBanAddUpd, IIf(ObjRS.RecordCount > 0, True, False), True), IIf(ObjRS.RecordCount > 0, True, False))
    ' Si entoncontro INformación en la base, atrapamos el Id del registro a modificar, si no, la varible IDCONSULTA SE QUEDA CON SU MISMO VALOR
    vlintIDConsulta = IIf(ObjRS.RecordCount > 0, ObjRS!intid, vlintIDConsulta)
    ObjRS.Close 'cerramos el recordset por si las moscas
     

        vlstrsql = "DELETE FROM PvDatosFiscales WHERE CHRRFC = '" & fStrRFCValido(txtRFC.Text) & "'"
        If vlstrNumPac <> "" Then
            vlstrsql = vlstrsql & " AND INTNUMCUENTA = " & vlstrNumPac
        Else
            vlstrsql = vlstrsql & " AND INTNUMCUENTA IS NULL"
        End If
        If vlstrTipoPac <> "" Then
            vlstrsql = vlstrsql & " AND CHRTIPOPACIENTE = '" & vlstrTipoPac & "'"
        Else
            vlstrsql = vlstrsql & " AND CHRTIPOPACIENTE IS NULL"
        End If
        
        
        pEjecutaSentencia vlstrsql
        vglngCveDatosFiscales = vlintIDConsulta ' esta variable no se usa en esta pantalla pero puede que en otras si, asi que la agregamos
        vllngRegimenFiscal = cboRegimenFiscal.ItemData(cboRegimenFiscal.ListIndex) 'Regimen fiscal
                             
' insertamos el registro de los datos fiscales
        vlstrsql = "INSERT INTO PvDatosFiscales (CHRRFC, " & "CHRNOMBRE, " & "CHRCALLE, " & _
                                                "VCHNUMEROEXTERIOR, " & "VCHNUMEROINTERIOR, " & _
                                                "CHRTELEFONO, " & _
                                                "CHRTIPOPACIENTE, " & _
                                                "INTNUMCUENTA, " & _
                                                "intNumReferencia, " & _
                                                "chrTipoCliente, " & _
                                                "INTCVECIUDAD, " & _
                                                "vchColonia, " & _
                                                "vchCodigoPostal, bitExtranjero" & IIf(Me.chkIEPS.Enabled = False, "", " ,bitSujetoAIEPS ") & _
                                                ", vchCorreoElectronico,VCHREGIMENFISCAL) " & _
                                  " VALUES('" & IIf(fStrRFCValido(txtRFC.Text) = "", " ", fStrRFCValido(txtRFC.Text)) & "','" & _
                                                Replace(Trim(txtNombreFactura.Text), "'", "''") & "','" & Trim(txtDireccionFactura.Text) & "','" & _
                                                Trim(txtNumExterior.Text) & "','" & _
                                                Trim(txtNumInterior.Text) & "','" & Trim(txtTelefonoFactura.Text) & "'," & _
                                                IIf(vlstrTipoPac = "", " NULL ", "'" & vlstrTipoPac & "'") & ", " & _
                                                IIf(vlstrNumPac = "", " NULL ", vlstrNumPac) & ", " & _
                                                IIf(vlstrNumRef = "", "NULL", vlstrNumRef) & ",'" & _
                                                IIf(vlstrTipo = "", "OT", vlstrTipo) & "'," & _
                                                str(cboCiudad.ItemData(cboCiudad.ListIndex)) & _
                                                ",'" & Trim(txtColonia.Text) & "'," & _
                                                "'" & Trim(txtCP.Text) & "'," & IIf((chkExtranjero.Value = vbChecked), 1, 0) & _
                                                IIf(Me.chkIEPS.Enabled = False, "", "," & IIf((chkIEPS.Value = vbChecked), 1, 0)) & _
                                                " ,'" & Trim(txtCorreo.Text) & "' ,'" & IIf(vllngRegimenFiscal = 0, "", vllngRegimenFiscal) & "' ) "
        pEjecutaSentencia vlstrsql
        vglngCveDatosFiscales = flngObtieneIdentity("SEC_PvDatosFiscales", vglngCveDatosFiscales)  ' esta variable no se usa en esta pantalla pero puede que en otras si, asi que la agregamos
      
    vgstrRFC = fStrRFCValido(txtRFC.Text)
    vgstrNombre = txtNombreFactura.Text
    vgstrDireccion = txtDireccionFactura.Text
    vgstrNumExterior = txtNumExterior.Text
    vgstrNumInterior = txtNumInterior.Text
    vgBitExtranjero = IIf(chkExtranjero.Value = vbChecked, 1, 0)
    vgBitSujetoaIEPS = IIf(Me.chkIEPS.Enabled = True, IIf(chkIEPS.Value = vbChecked, 1, 0), 0)
    vgstrTelefono = txtTelefonoFactura.Text
    vgstrColonia = txtColonia.Text
    vgstrCP = txtCP.Text
    'Regimen fiscal agregado para el caso 181
    vlstrRegimenFiscal = cboRegimenFiscal.ItemData(cboRegimenFiscal.ListIndex)
    
    llngCveCiudad = cboCiudad.ItemData(cboCiudad.ListIndex)
    If cboUsoCFDI.ListIndex > -1 Then
        vgintUsoCFDI = cboUsoCFDI.ItemData(cboUsoCFDI.ListIndex)
    End If
    vgstrCorreo = Trim(txtCorreo.Text)
    
    vglngNumRef = 0
    If Not (Trim(vlstrNumRef) = "" Or Trim(vlstrNumRef) = "NULL") Then
        vglngNumRef = CLng(Trim(vlstrNumRef))
    ElseIf Not Trim(vlstrNumPac) = "" Then
        vglngNumRef = CLng(Trim(vlstrNumPac))
    End If
    
    If Trim(vlstrTipo) <> "" And Trim(vlstrTipo) <> "OT" Then
        vgstrTipo = vlstrTipo
    ElseIf Trim(vlstrTipoPac) <> "" Then
        vgstrTipo = IIf(vlstrTipoPac = "I", "PI", "PE")
    End If
    vlblnUsarRazonSocial = IIf(chkRazonSocial = 1, 1, 0)
    
    vgblnModalResult = True
    Me.Hide
    
End Sub
Private Function fnValidaDatos() As Boolean
    fnValidaDatos = True
    
    If Trim(txtRFC.Text) = "" Then
        'Favor de registrar el RFC
        MsgBox SIHOMsg(1013), vbExclamation + vbOKOnly, "Mensaje"
        txtRFC.SetFocus
        fnValidaDatos = False
        Exit Function
    End If
    
    If Trim(txtRFC.Text) <> "" Then
        If vlblnLicenciaContaElectronica Then
            If Len(Trim(txtRFC.Text)) <> 12 And Len(Trim(txtRFC.Text)) <> 13 Then
                'El RFC ingresado no tiene un tamaño válido, favor de verificar.
                MsgBox SIHOMsg(1345), vbOKOnly + vbInformation, "Mensaje"
                txtRFC.SetFocus
                fnValidaDatos = False
                Exit Function
            End If
        End If
    End If
    'Regimen fiscal
    'Si la version de CFDI es 4.0, es obligatoria la captura del regimen
    If vgstrVersionCFDI = "4.0" Then
        If cboRegimenFiscal.ListIndex = 0 Then
            fnValidaDatos = False
            '¡No ha ingresado datos!
            MsgBox SIHOMsg(2), vbExclamation + vbOKOnly, "Mensaje"
            If cboRegimenFiscal.Enabled Then
                cboRegimenFiscal.SetFocus
            End If
            Exit Function
        End If
    End If
        
    If Trim(txtNombreFactura.Text) = "" Then
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbExclamation + vbOKOnly, "Mensaje"
        txtNombreFactura.SetFocus
        fnValidaDatos = False
        Exit Function
    End If
        
    If cboCiudad.ListCount = 0 Or cboCiudad.ListIndex = -1 Then
        '¡Dato no válido, seleccione un valor de la lista!
        MsgBox SIHOMsg(3), vbExclamation + vbOKOnly, "Mensaje"
        cboCiudad.SetFocus
        fnValidaDatos = False
        Exit Function
    End If
    'Si la version de CFDI es 4.0, es obligatoria la captura del regimen
    If vgstrVersionCFDI = "4.0" Then
        If Trim(txtCP.Text) = "" Or Len(Me.txtCP.Text) < 5 Then ' el código postal tiene que ser de  5 dígitos
            '¡Dato no válido, el código postal debe ser de 5 dígitos!
            MsgBox SIHOMsg(1181), vbExclamation + vbOKOnly, "Mensaje"
            Me.txtCP.SetFocus
            fnValidaDatos = False
            Exit Function
        End If
    Else
        If Trim(txtCP.Text) <> "" And Len(Me.txtCP.Text) <> 5 Then ' el código postal puede ser null o de 5 dígitos
            '¡Dato no válido, el código postal debe ser de 5 dígitos!
            MsgBox SIHOMsg(1181), vbExclamation + vbOKOnly, "Mensaje"
            Me.txtCP.SetFocus
            fnValidaDatos = False
            Exit Function
        End If
    End If
    
    If cboUsoCFDI.Visible Then
        If cboUsoCFDI.ListIndex = -1 Then 'Uso del CFDI
            MsgBox "Seleccione el uso del CFDI", vbExclamation + vbOKOnly, "Mensaje"
            cboUsoCFDI.SetFocus
            fnValidaDatos = False
            Exit Function
        End If
    End If
    
    
End Function
Private Sub Form_Activate()
    txtNombreFactura.Text = vgstrNombre
    txtDireccionFactura.Text = vgstrDireccion
    txtNumExterior.Text = vgstrNumExterior
    txtNumInterior.Text = vgstrNumInterior
    chkExtranjero.Value = IIf(vgBitExtranjero = 1, vbChecked, vbUnchecked)
    Me.chkIEPS.Enabled = vlblnLicenciaIEPS And vgActivaSujetoaIEPS
    If vlblnLicenciaIEPS And vgActivaSujetoaIEPS Then chkIEPS.Value = IIf(vgBitSujetoaIEPS = 1, vbChecked, vbUnchecked)
    txtTelefonoFactura.Text = vgstrTelefono
    txtColonia.Text = vgstrColonia
    txtCP.Text = vgstrCP
    txtRFC.Text = Trim(Replace(Replace(Replace(vgstrRFC, "-", ""), "_", ""), " ", ""))
    txtCorreo.Text = vgstrCorreo
    
    vlintIDConsulta = -1
    If sstDatos.Tab = 1 Then
        pEnfocaTextBox txtBuscaDatos
        optBusqueda(2).Value = True
        Me.lstBuscaDatos.Clear
    Else
        If Trim(txtNombreFactura.Text) = "" Then
            optBusqueda(2).Value = True
            pEnfocaTextBox txtRFC
        Else
            pEnfocaTextBox txtNombreFactura
        End If
    End If
    vgblnModalResult = False
    txtBuscaDatos.Text = ""
    freBuscarDatos.Enabled = True
    If vgblnMostrarUsoCFDI Then
        pCargaUsosCFDI
        lblUsoCFDI.Visible = True
        cboUsoCFDI.Visible = True
        cboUsoCFDI.ListIndex = flngLocalizaCbo(cboUsoCFDI, flngCatalogoSATIdByNombreTipo("c_UsoCFDI", CLng(vgintTipoPacEmp), vgstrTipoUsoCFDI, 1))
        If cboUsoCFDI.ListIndex = -1 Then
            cboUsoCFDI.ListIndex = flngLocalizaCbo(cboUsoCFDI, 64)
        End If
    Else
        lblUsoCFDI.Visible = False
        cboUsoCFDI.Visible = False
    End If
    'Regimen fiscal
    'pCargaRegimenFiscal
    If vlstrRegimenFiscal <> "" Then
        cboRegimenFiscal.ListIndex = flngLocalizaCbo(cboRegimenFiscal, vlstrRegimenFiscal)
    End If
    vlblnUsarRazonSocial = 0
    chkRazonSocial.Visible = False
    Dim frm As Form
    For Each frm In Forms
        If frm.Name = "frmPOS" Then
            chkRazonSocial.Visible = True
        End If
    Next frm
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then ' si se preciona Esc
        If sstDatos.Tab = 0 Then ' esta en la pantalla de la captura
            If vlintIDConsulta = 0 Then ' no hay consulta activa
                Unload Me
            Else ' si hay consulta activa
                 '¿Desea abandonar la operación?
               If MsgBox(SIHOMsg(17), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
               ' se limpia toda la forma
                  If vglngDatosParametro = True Then ' aqui ya dejan de ser los parametros venta al público
                     vglngDatosParametro = False
                     vgActivaSujetoaIEPS = True
                  End If
                 pLimpiaForma
               End If
            End If
        Else ' si esta en la pantalla de la consulta
            pLimpiaForma
        End If
    ElseIf KeyAscii = 13 Then
        SendKeys vbTab
'    ElseIf KeyAscii = 39 Then
'        KeyAscii = 0
    End If
End Sub
Private Sub Form_Load()
    Me.Icon = frmMenuPrincipal.Icon
    vlblnLicenciaIEPS = fblLicenciaIEPS
    vlblnLicenciaContaElectronica = fblnLicenciaContaElectronica
    pLimpiaForma
End Sub

Private Sub pLimpiaForma()
    Dim rs As New ADODB.Recordset
    
    
    ' Inicialización
    vgstrRFC = ""
    vgstrNombre = ""
    vgstrDireccion = ""
    vgstrNumExterior = ""
    vgstrNumInterior = ""
    vgBitExtranjero = 0
    vgstrTelefono = ""
    vgstrColonia = ""
    vgstrCP = ""
    vglngNumRef = 0
    vgstrTipo = "OT"
    txtRFC.Text = ""
    txtDireccionFactura.Text = ""
    txtNumExterior.Text = ""
    txtNumInterior.Text = ""
    Me.txtColonia.Text = ""
    Me.txtCP.Text = ""
    chkExtranjero.Value = vbUnchecked
    txtNombreFactura.Text = ""
    txtTelefonoFactura.Text = ""
    vgblnModalResult = False
    fraBusquedaClave.Visible = False
    vlintIDConsulta = 0
    vgblnMostrarUsoCFDI = False
    Me.vglngDatosParametro = False ' se desactiva la variable para que no afecte
    Me.chkIEPS.Value = vbUnchecked
    Me.chkIEPS.Enabled = vlblnLicenciaIEPS And vgActivaSujetoaIEPS
    Me.chkIEPS.Visible = vlblnLicenciaIEPS
    'Cargar las ciudades:
    vgstrParametrosSP = "-1|-1|1"
    ' (Modificación de Keyla para llenar el combo de ciudades)
    'Set rs = frsEjecuta_SP(vgstrParametrosSP, "SP_GNSELCIUDAD")
    Set rs = frsEjecuta_SP("", "Sp_GnSelCiudadEstado")
    If rs.RecordCount <> 0 Then
        pLlenarCboRs cboCiudad, rs, 0, 1
    End If
    vgstrCorreo = ""
    txtCorreo.Text = ""
    vlstrTipoPac = ""  ' Tipo de paciente
    vlstrNumPac = ""   ' Número de paciente

    optBusqueda(0).Value = False
    optBusqueda(1).Value = False
    optBusqueda(2).Value = False
    optBusqueda(3).Value = False
    optBusqueda(4).Value = False
    optBusqueda(5).Value = False
    sstDatos.Tab = 0
    pEnfocaTextBox txtRFC
    pCargaRegimenFiscal
    chkRazonSocial.Enabled = False
    chkRazonSocial.Value = 0
    chkRFCgenerico.Value = 0
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = True
    If sstDatos.Tab = 0 Then
        If vlintIDConsulta = 0 Then ' no hay consulta activa
            If frmDatosFiscales.Visible Then
               frmDatosFiscales.Hide
            End If
        Else ' si hay consulta activa
            If frmDatosFiscales.Visible Then
                '¿Desea abandonar la operación?
                If MsgBox(SIHOMsg(17), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
                    ' se limpia toda la forma
                    pLimpiaForma
                End If
            End If
        End If
    Else
        pLimpiaForma
    End If
End Sub

Private Sub lstBuscaDatos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        lstBuscaDatos_DblClick
    End If
End Sub
Private Sub optBusqueda_Click(Index As Integer)
    freBuscarDatos.Enabled = True
    txtBuscaDatos.SetFocus
    lstBuscaDatos.Clear
    lstBuscaDatos.Enabled = False
    txtBuscaDatos_KeyUp 0, 0
End Sub
Private Sub optClave_GotFocus(Index As Integer)
    If optClave(0).Value = True Then
        fraBusquedaClave.Visible = True
        pEnfocaTextBox txtNumeroReferencia
    Else
        fraBusquedaClave.Visible = False
        pEnfocaTextBox txtBuscaDatos
    End If
End Sub

Private Sub txtBuscaDatos_LostFocus()
If lstBuscaDatos.ListCount > 0 Then
    lstBuscaDatos.SetFocus
End If



End Sub

Private Sub txtColonia_GotFocus()
    pSelTextBox txtColonia
End Sub

Private Sub txtColonia_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtCorreo_GotFocus()
    pSelTextBox txtCorreo
End Sub

Private Sub txtCP_GotFocus()
    pSelTextBox txtCP
End Sub
Private Sub txtCP_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError

    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then KeyAscii = 7

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtCP_KeyPress"))
End Sub
Private Sub txtDireccionFactura_GotFocus()
    pSelTextBox txtDireccionFactura
End Sub
Private Sub txtNombreFactura_GotFocus()
    pSelTextBox txtNombreFactura
End Sub
Private Sub txtNumeroReferencia_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyReturn Or KeyCode = vbKeyRight Or KeyCode = vbKeyDown) And lstBuscaDatos.Enabled Then
        lstBuscaDatos.SetFocus
    End If
End Sub
Private Sub txtNumeroReferencia_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim vlstrSentencia As String
    
    If optBusqueda(0).Value Then 'Externos
        vlstrSentencia = "SELECT intNumPaciente Clave, rtrim(chrApePaterno) || ' ' || rtrim(chrApeMaterno) || ' ' || rtrim(chrNombre) Nombre FROM Externo"
        PSuperBusqueda txtNumeroReferencia, vlstrSentencia, lstBuscaDatos, "intNumPaciente", 20, , "intNumPaciente"
    ElseIf optBusqueda(1).Value Then  'Empresas
        vlstrSentencia = "SELECT intCveEmpresa Clave, vchDescripcion Descripcion FROM CCEmpresa"
        PSuperBusqueda txtNumeroReferencia, vlstrSentencia, lstBuscaDatos, "intCveEmpresa", 20, , "intCveEmpresa"
    ElseIf optBusqueda(2).Value Then  'Otros PVDatosFiscales
        vlstrSentencia = "SELECT intID Clave, chrNombre Descripcion,intnumcuenta FROM pvDatosFiscales"
        PSuperBusqueda txtNumeroReferencia, vlstrSentencia, lstBuscaDatos, "intID", 20, , "intID"
    ElseIf optBusqueda(3).Value Then 'Internos
        vlstrSentencia = "SELECT NUMCVEPACIENTE Clave, rtrim(vchApellidoPaterno) || ' ' || rtrim(vchApellidoMaterno) || ' ' || rtrim(vchNombre) Nombre FROM AdPaciente"
        PSuperBusqueda txtNumeroReferencia, vlstrSentencia, lstBuscaDatos, "NUMCVEPACIENTE", 20, , "NUMCVEPACIENTE"
    ElseIf optBusqueda(4).Value Then 'Empleados
        vlstrSentencia = "SELECT INTCVEEMPLEADO Clave, rtrim(vchApellidoPaterno) || ' ' || rtrim(vchApellidoMaterno) || ' ' || rtrim(vchNombre) Nombre FROM NoEmpleado"
        PSuperBusqueda txtNumeroReferencia, vlstrSentencia, lstBuscaDatos, "INTCVEEMPLEADO", 20, , "INTCVEEMPLEADO"
    ElseIf optBusqueda(5).Value Then 'Médicos
        vlstrSentencia = "SELECT INTCVEMEDICO Clave, rtrim(vchApellidoPaterno) || ' ' || rtrim(vchApellidoMaterno) || ' ' || rtrim(vchNombre) Nombre FROM HoMedico"
        PSuperBusqueda txtNumeroReferencia, vlstrSentencia, lstBuscaDatos, "INTCVEMEDICO", 20, , "INTCVEMEDICO"
    End If
End Sub

Private Sub txtNumeroReferencia_LostFocus()
    If lstBuscaDatos.ListCount > 0 Then
        lstBuscaDatos.SetFocus
        
    End If
End Sub

Private Sub txtNumExterior_GotFocus()
    pSelTextBox txtNumExterior
End Sub
Private Sub txtNumExterior_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub txtNumInterior_GotFocus()
    pSelTextBox txtNumInterior
End Sub
Private Sub txtNumInterior_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtRFC_Change()
    If txtRFC = "XAXX010101000" Or txtRFC = "XEXX010101000" Then
        pValidaBloqueoRegimen
    End If
End Sub

Private Sub txtRFC_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim vlstrSentencia As String
    Dim rs As New ADODB.Recordset
    Dim strtempRFC As String

    vgstrRFC = txtRFC.Text
    If KeyCode = vbKeyReturn Then
        If vlintIDConsulta = 0 Then
           'Me.vglngDatosParametro = False ' aqui ya no importa si viene de parametros por que se esta eligiendo otros
                  
           If vglngDatosParametro = True Then ' aqui ya dejan de ser los parametros venta al público
              vglngDatosParametro = False
              vgActivaSujetoaIEPS = True '<---------------------------------------------------------
           End If
                    
            vlstrSentencia = "SELECT * FROM PvDatosFiscales WHERE chrRFC = '" & Trim(txtRFC.Text) & "'"
            Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
            If rs.RecordCount > 0 Then
                txtNombreFactura.Text = IIf(IsNull(rs!CHRNOMBRE), "", Trim(rs!CHRNOMBRE))
                txtDireccionFactura.Text = IIf(IsNull(rs!chrCalle), "", Trim(rs!chrCalle))
                txtNumExterior.Text = IIf(IsNull(rs!VCHNUMEROEXTERIOR), "", Trim(rs!VCHNUMEROEXTERIOR))
                txtNumInterior.Text = IIf(IsNull(rs!VCHNUMEROINTERIOR), "", Trim(rs!VCHNUMEROINTERIOR))
                chkExtranjero.Value = IIf(rs!bitExtranjero = 0, vbUnchecked, vbChecked)
                'pActivaCHKIEPS
                chkIEPS.Value = IIf(Me.chkIEPS.Enabled = False, vbUnchecked, IIf(vlblnLicenciaIEPS, IIf(rs!bitsujetoaieps = 0, vbUnchecked, vbChecked), vbUnchecked))
                cboCiudad.ListIndex = flngLocalizaCbo(cboCiudad, str(IIf(IsNull(rs!intCveCiudad), 0, rs!intCveCiudad)))
                txtTelefonoFactura.Text = IIf(IsNull(rs!chrTelefono), "", Trim(rs!chrTelefono))
                txtColonia.Text = IIf(IsNull(rs!VCHCOLONIA), "", Trim(rs!VCHCOLONIA))
                txtCP.Text = IIf(IsNull(rs!VCHCODIGOPOSTAL), "", Trim(rs!VCHCODIGOPOSTAL))
                vlstrTipo = IIf(IsNull(rs!chrTipoCliente), "OT", rs!chrTipoCliente)
                vlstrNumRef = IIf(IsNull(rs!intNumReferencia), "NULL", rs!intNumReferencia)
                vlintIDConsulta = rs!intid '  si encuentra datos, cacha el ID del renglon a modificar
                txtCorreo.Text = IIf(IsNull(rs!vchCorreoElectronico), "", Trim(rs!vchCorreoElectronico))
            Else
                txtNombreFactura.Text = ""
                txtDireccionFactura.Text = ""
                txtNumExterior.Text = ""
                txtNumInterior.Text = ""
                If chkExtranjero.Value = vbChecked Then
                    If txtRFC.Text = "" Then chkExtranjero.Value = vbUnchecked
                End If
                'pActivaCHKIEPS
                chkIEPS.Value = vbUnchecked
                cboCiudad.ListIndex = -1
                txtTelefonoFactura.Text = ""
                txtColonia.Text = ""
                txtCP.Text = ""
                vlstrTipo = "OT"
                vlstrNumRef = "NULL"
                vlintIDConsulta = -2 ' si no hay datos ID a -2 quiere decir que se ingresa un nuevo registro
                txtCorreo.Text = ""
            End If
            rs.Close
        
          '----- COMENTAREADO PARA CASO 7931 -----'
'         ElseIf vlintIDConsulta = -1 And Not vglngDatosParametro Then ' hay una consulta que no proviene de pvdatosfiscales
'                vlstrSentencia = "SELECT * FROM PvDatosFiscales WHERE chrRFC = '" & Trim(txtRFC.Text) & "'"
'                Set rs = frsRegresaRs(vlstrSentencia, adLockOptimistic)
'                If rs.RecordCount > 0 Then
'                   txtNombreFactura.Text = IIf(IsNull(rs!chrNombre), "", Trim(rs!chrNombre))
'                   txtDireccionFactura.Text = IIf(IsNull(rs!chrCalle), "", Trim(rs!chrCalle))
'                   txtNumExterior.Text = IIf(IsNull(rs!vchNumeroExterior), "", Trim(rs!vchNumeroExterior))
'                   txtNumInterior.Text = IIf(IsNull(rs!vchNumeroInterior), "", Trim(rs!vchNumeroInterior))
'                   chkExtranjero.Value = IIf(rs!bitExtranjero = 0, vbUnchecked, vbChecked)
'                   cboCiudad.ListIndex = flngLocalizaCbo(cboCiudad, Str(IIf(IsNull(rs!intCveCiudad), 0, rs!intCveCiudad)))
'                   txtTelefonoFactura.Text = IIf(IsNull(rs!chrTelefono), "", Trim(rs!chrTelefono))
'                   txtColonia.Text = IIf(IsNull(rs!vchcolonia), "", Trim(rs!vchcolonia))
'                   txtCP.Text = IIf(IsNull(rs!vchcodigopostal), "", Trim(rs!vchcodigopostal))
'                   vlstrTipo = IIf(IsNull(rs!chrTipoCliente), "OT", rs!chrTipoCliente)
'                   vlstrNumRef = IIf(IsNull(rs!intNumReferencia), "NULL", rs!intNumReferencia)
'                   vlintIDConsulta = rs!IntID '  si encuentra datos, cacha el ID del renglon a modificar
'                End If
         End If
         pActivaCHKIEPS
    End If
End Sub
Private Sub pActivaCHKIEPS()
If vglngDatosParametro = False And fStrRFCValido(Me.txtRFC.Text) <> "XAXX010101000" And fStrRFCValido(Me.txtRFC.Text) <> "XEXX010101000" And Len(fStrRFCValido(Me.txtRFC.Text)) >= 12 And Len(fStrRFCValido(Me.txtRFC.Text)) <= 13 Then
   Me.chkIEPS.Enabled = vlblnLicenciaIEPS And vgActivaSujetoaIEPS
ElseIf vglngDatosParametro = True And Len(fStrRFCValido(Me.txtRFC.Text)) >= 12 And Len(fStrRFCValido(Me.txtRFC.Text)) <= 13 Then
   Me.chkIEPS.Enabled = vlblnLicenciaIEPS And vgActivaSujetoaIEPS
Else
   Me.chkIEPS.Value = vbUnchecked
   Me.chkIEPS.Enabled = False
End If

End Sub
Private Sub txtRFC_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) And Not (KeyAscii >= 65 And KeyAscii <= 90) And Not (KeyAscii >= 97 And KeyAscii <= 122) And Not (KeyAscii = 209 Or KeyAscii = 241 Or KeyAscii = 127 Or KeyAscii = 8 Or KeyAscii = 38) Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub
Private Sub txtBuscaDatos_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyRight Or KeyCode = vbKeyDown) And lstBuscaDatos.Enabled Then
        lstBuscaDatos.SetFocus
    End If
End Sub

Private Sub txtBuscaDatos_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNumeroReferencia_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub lstBuscaDatos_DblClick()
    Dim stropcion As String
    Dim vlstrSentencia As String
    Dim rs As New ADODB.Recordset
    Dim i As Integer
    Dim ban As Integer
    
    ban = 0
    vlstrTipoPac = ""
    vlstrNumPac = ""
    stropcion = ""
    If optBusqueda(0).Value Then 'Externos
        stropcion = "EX"
        If frmFacturacion.txtCantidadFP > 0 Then
            vlstrSentencia = "SELECT vchRFC RFC, ExPacienteIngreso.INTNUMCUENTA Clave, RTRIM(vchNombre) || ' ' || RTRIM(vchApellidoPaterno) || ' ' || RTRIM(vchApellidoMaterno) Nombre, trim(domicilio.vchCALLE) chrCalle,  domicilio.VCHNUMEROEXTERIOR, domicilio.VCHNUMEROINTERIOR, trim(telefono.vchTelefono) Telefono, 'PE' Tipo, domicilio.intCveCiudad IdCiudad, domicilio.vchColonia Colonia, domicilio.vchCodigoPostal CP, exPaciente.vchCorreoElectronico correo " & _
                             " FROM ExPaciente INNER JOIN ExPacienteIngreso ON (ExPaciente.INTNUMPACIENTE = ExPacienteIngreso.INTNUMPACIENTE)  " & _
                             " left join (select gnDomicilio.*, expacientedomicilio.intnumpaciente from gndomicilio inner join exPacienteDomicilio on expacientedomicilio.intcvedomicilio = gndomicilio.intcvedomicilio AND gndomicilio.intcvetipodomicilio = 1 " & _
                             "            ) domicilio on ExPaciente.intnumpaciente = domicilio.intnumpaciente " & _
                             " left join (select gntelefono.*, exPAcientetelefono.intnumPaciente from gnTelefono inner join exPAcientetelefono on exPAcientetelefono.intCveTelefono = gnTelefono.intCveTelefono and gnTelefono.intCveTipoTelefono = 1 " & _
                             "            ) telefono on ExPaciente.intnumpaciente = telefono.intnumpaciente" & _
                             " where ExPacienteIngreso.intCuentaFacturada = 0 AND ExPacienteIngreso.chrTipoIngreso='E' AND ExPaciente.intNumPaciente = " & Trim(str(lstBuscaDatos.ItemData(lstBuscaDatos.ListIndex)))
        Else
            vlstrSentencia = "SELECT vchRFC RFC, ExPacienteIngreso.INTNUMCUENTA Clave, RTRIM(vchNombre) || ' ' || RTRIM(vchApellidoPaterno) || ' ' || RTRIM(vchApellidoMaterno) Nombre, trim(domicilio.vchCALLE) chrCalle,  domicilio.VCHNUMEROEXTERIOR, domicilio.VCHNUMEROINTERIOR, trim(telefono.vchTelefono) Telefono, 'PE' Tipo, domicilio.intCveCiudad IdCiudad, domicilio.vchColonia Colonia, domicilio.vchCodigoPostal CP, exPaciente.vchCorreoElectronico correo " & _
                             " FROM ExPaciente LEFT OUTER JOIN ExPacienteIngreso ON (ExPaciente.INTNUMPACIENTE = ExPacienteIngreso.INTNUMPACIENTE AND ExPacienteIngreso.DTMFECHAHORAEGRESO IS NULL) AND ExPacienteIngreso.chrTipoIngreso='E' " & _
                             " left join (select gnDomicilio.*, expacientedomicilio.intnumpaciente from gndomicilio inner join exPacienteDomicilio on expacientedomicilio.intcvedomicilio = gndomicilio.intcvedomicilio AND gndomicilio.intcvetipodomicilio = 1 " & _
                             "            ) domicilio on ExPaciente.intnumpaciente = domicilio.intnumpaciente " & _
                             " left join (select gntelefono.*, exPAcientetelefono.intnumPaciente from gnTelefono inner join exPAcientetelefono on exPAcientetelefono.intCveTelefono = gnTelefono.intCveTelefono and gnTelefono.intCveTipoTelefono = 1 " & _
                             "            ) telefono on ExPaciente.intnumpaciente = telefono.intnumpaciente" & _
                             " WHERE ExPaciente.intNumPaciente = " & Trim(str(lstBuscaDatos.ItemData(lstBuscaDatos.ListIndex)))
        End If
        optBusqueda(0).Value = False
        ban = 0
        vlstrTipoPac = "E"
        
    ElseIf optBusqueda(1).Value Then 'Empresa
        stropcion = "EM"
        vlstrSentencia = "SELECT chrRfcEmpresa RFC, intCveEmpresa Clave, vchRazonSocial Nombre, chrCalle, vchNumeroExterior,vchNumeroInterior, chrTelefonoEmpresa Telefono, 'CO' Tipo, CcEmpresa.intCveCiudad IdCiudad, CCEmpresa.vchColonia Colonia, TO_CHAR(CCEmpresa.vchCodigoPostal) CP, trim(vchCorreo) correo,nvl(CCEmpresa.VCHREGIMENFISCAL,0) VCHREGIMENFISCAL FROM CCEmpresa  WHERE intCveEmpresa = " & Trim(str(lstBuscaDatos.ItemData(lstBuscaDatos.ListIndex)))
        optBusqueda(1).Value = False
        ban = 0
        
    ElseIf optBusqueda(2).Value Then 'Otros
        stropcion = "OT"
        '- CASO 7931: Se agregó a la consulta el número de cuenta y el tipo del paciente -'
        vlstrSentencia = "SELECT PvDatosFiscales.intID ,PvDatosFiscales.chrRFC RFC, PvDatosFiscales.intNumReferencia Clave, PvDatosFiscales.chrNombre Nombre, PvDatosFiscales.chrCalle, " & _
                            " PvDatosFiscales.vchNumeroExterior, PvDatosFiscales.vchNumeroInterior, PvDatosFiscales.bitextranjero, PvDatosFiscales.chrTelefono Telefono, " & _
                            " PvDatosFiscales.chrTipoCliente Tipo, PvDatosFiscales.intCveCiudad IdCiudad, PvDatosFiscales.vchColonia Colonia, PvDatosFiscales.vchCodigoPostal CP, " & _
                            " PvDatosFiscales.intNumCuenta Cuenta, PvDatosFiscales.chrTipoPaciente TipoPac, PvDatosFiscales.bitsujetoaIEPS, " & _
                            " CASE WHEN trim(NVL(PvDatosFiscales.vchCorreoElectronico,'')) = '' THEN exPaciente.vchCorreoElectronico ELSE PvDatosFiscales.vchCorreoElectronico END correo,nvl(PvDatosFiscales.VCHREGIMENFISCAL,0) VCHREGIMENFISCAL " & _
                                " FROM PvDatosFiscales " & _
                                " LEFT JOIN EXPACIENTEINGRESO ON EXPACIENTEINGRESO.INTNUMCUENTA = PvDatosFiscales.INTNUMCUENTA " & _
                                    " AND EXPACIENTEINGRESO.INTCVETIPOINGRESO IN (SELECT SITIPOINGRESO.INTCVETIPOINGRESO FROM SITIPOINGRESO WHERE SITIPOINGRESO.CHRTIPOINGRESO = PVDATOSFISCALES.CHRTIPOPACIENTE ) " & _
                                " LEFT JOIN EXPACIENTE ON EXPACIENTE.INTNUMPACIENTE = EXPACIENTEINGRESO.INTNUMPACIENTE " & _
                            " WHERE PvDatosFiscales.intID = " & Trim(str(lstBuscaDatos.ItemData(lstBuscaDatos.ListIndex)))
        optBusqueda(2).Value = False
        ban = 1
        
    ElseIf optBusqueda(3).Value Then 'Paciente Interno
        stropcion = "IN"
        If frmFacturacion.txtCantidadFP > 0 Then
            vlstrSentencia = "SELECT vchRFC RFC, ExPacienteIngreso.INTNUMCUENTA Clave, RTRIM(vchNombre) || ' ' || RTRIM(vchApellidoPaterno) || ' ' || RTRIM(vchApellidoMaterno) Nombre, trim(domicilio.vchCALLE) chrCalle, domicilio.VCHNUMEROEXTERIOR, domicilio.VCHNUMEROINTERIOR, trim(telefono.vchTelefono) Telefono, 'PI' Tipo, domicilio.intCveCiudad IdCiudad, domicilio.vchColonia Colonia, domicilio.vchCodigoPostal CP, exPaciente.vchCorreoElectronico correo " & _
                             " FROM ExPaciente INNER JOIN ExPacienteIngreso ON (ExPaciente.INTNUMPACIENTE = ExPacienteIngreso.INTNUMPACIENTE)  " & _
                             " left join (select gnDomicilio.*, expacientedomicilio.intnumpaciente from gndomicilio inner join exPacienteDomicilio on expacientedomicilio.intcvedomicilio = gndomicilio.intcvedomicilio AND gndomicilio.intcvetipodomicilio = 1 " & _
                             "            ) domicilio on ExPaciente.intnumpaciente = domicilio.intnumpaciente " & _
                             " left join (select gntelefono.*, exPAcientetelefono.intnumPaciente from gnTelefono inner join exPAcientetelefono on exPAcientetelefono.intCveTelefono = gnTelefono.intCveTelefono and gnTelefono.intCveTipoTelefono = 1 " & _
                             "            ) telefono on ExPaciente.intnumpaciente = telefono.intnumpaciente" & _
                             " where ExPacienteIngreso.intCuentaFacturada = 0 AND ExPacienteIngreso.chrTipoIngreso='I' AND ExPaciente.intNumPaciente = " & Trim(str(lstBuscaDatos.ItemData(lstBuscaDatos.ListIndex)))
        Else
            vlstrSentencia = "SELECT vchRFC RFC, ExPacienteIngreso.INTNUMCUENTA Clave, RTRIM(vchNombre) || ' ' || RTRIM(vchApellidoPaterno) || ' ' || RTRIM(vchApellidoMaterno) Nombre, trim(domicilio.vchCALLE) chrCalle, domicilio.VCHNUMEROEXTERIOR, domicilio.VCHNUMEROINTERIOR, trim(telefono.vchTelefono) Telefono, 'PI' Tipo, domicilio.intCveCiudad IdCiudad, domicilio.vchColonia Colonia, domicilio.vchCodigoPostal CP, exPaciente.vchCorreoElectronico correo " & _
                             " FROM ExPaciente LEFT OUTER JOIN ExPacienteIngreso ON (ExPaciente.INTNUMPACIENTE = ExPacienteIngreso.INTNUMPACIENTE AND ExPacienteIngreso.DTMFECHAHORAEGRESO IS NULL) AND ExPacienteIngreso.chrTipoIngreso='I' " & _
                             " left join (select gnDomicilio.*, expacientedomicilio.intnumpaciente from gndomicilio inner join exPacienteDomicilio on expacientedomicilio.intcvedomicilio = gndomicilio.intcvedomicilio AND gndomicilio.intcvetipodomicilio = 1 " & _
                             "            ) domicilio on ExPaciente.intnumpaciente = domicilio.intnumpaciente " & _
                             " left join (select gntelefono.*, exPAcientetelefono.intnumPaciente from gnTelefono inner join exPAcientetelefono on exPAcientetelefono.intCveTelefono = gnTelefono.intCveTelefono and gnTelefono.intCveTipoTelefono = 1 " & _
                             "            ) telefono on ExPaciente.intnumpaciente = telefono.intnumpaciente" & _
                             " WHERE ExPaciente.intNumPaciente = " & Trim(str(lstBuscaDatos.ItemData(lstBuscaDatos.ListIndex)))
        End If
        optBusqueda(3).Value = False
        ban = 0
        vlstrTipoPac = "I"
        
    ElseIf optBusqueda(4).Value Then 'Empleados
        stropcion = "E"
        vlstrSentencia = "SELECT chrRFC RFC, INTCVEEMPLEADO Clave, RTRIM(vchNombre) || ' ' || RTRIM(vchApellidoPaterno) || ' ' || RTRIM(vchApellidoMaterno) Nombre, chrCalle, vchNumeroExterior, vchNumeroInterior, chrTelefono Telefono, 'EM' Tipo,NoEmpleado.intCveCiudad IdCiudad, NoEmpleado.chrColonia Colonia, NoEmpleado.chrcodigopostal CP, trim(vchCorreo) correo, nvl(NoEmpleado.VCHREGIMENFISCAL,0) VCHREGIMENFISCAL FROM NoEmpleado WHERE INTCVEEMPLEADO = " & Trim(str(lstBuscaDatos.ItemData(lstBuscaDatos.ListIndex)))
        optBusqueda(4).Value = False
        ban = 0
        
    ElseIf optBusqueda(5).Value Then 'Medicos
        stropcion = "ME"
        vlstrSentencia = "SELECT vchRFCMedico RFC, INTCVEMEDICO Clave, RTRIM(vchNombre) || ' ' || RTRIM(vchApellidoPaterno) || ' ' || RTRIM(vchApellidoMaterno) Nombre, VCHCONSULCALLE AS chrCalle, VCHCONSULNUMEROEXTERIOR AS vchNumeroExterior, VCHCONSULNUMEROINTERIOR as vchNumeroInterior, '' Telefono, 'ME' Tipo, HoMedico.intCveCiudad IdCiudad, HoMedico.vchconsulcolonia Colonia, HoMedico.vchconsulcodpostal CP, trim(vchemail) correo, nvl(HoMedico.VCHREGIMENFISCAL,0) VCHREGIMENFISCAL FROM HoMedico WHERE INTCVEMEDICO = " & Trim(str(lstBuscaDatos.ItemData(lstBuscaDatos.ListIndex)))
        optBusqueda(5).Value = False
        ban = 0
    End If
    
    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    If rs.RecordCount > 0 Then
        If vglngDatosParametro = True Then ' aqui ya dejan de ser los parametros venta al público
           vglngDatosParametro = False
           vgActivaSujetoaIEPS = True
        End If
        If IsNull(rs!RFC) Then
           txtRFC.Text = ""
        Else
           txtRFC.Text = Trim(Replace(Replace(Replace(rs!RFC, "-", ""), "_", ""), " ", ""))
        End If
        txtNombreFactura.Text = IIf(IsNull(rs!Nombre), "", Trim(rs!Nombre))
        txtDireccionFactura.Text = IIf(IsNull(rs!chrCalle), "", Trim(rs!chrCalle))
        txtNumExterior.Text = IIf(IsNull(rs!VCHNUMEROEXTERIOR), "", Trim(rs!VCHNUMEROEXTERIOR))
        txtNumInterior.Text = IIf(IsNull(rs!VCHNUMEROINTERIOR), "", Trim(rs!VCHNUMEROINTERIOR))
        If ban = 1 Then
            chkExtranjero.Value = IIf((rs!bitExtranjero = 1), vbChecked, vbUnchecked)
            vlintIDConsulta = rs!intid ' cacha el ID de la tabla pvdatos fiscales
        Else
            chkExtranjero.Value = vbUnchecked
            vlintIDConsulta = -1 ' pone id en -1 para que no se active
        End If
        'ban = 0
        cboCiudad.ListIndex = flngLocalizaCbo(cboCiudad, str(IIf(IsNull(rs!IdCiudad), 0, rs!IdCiudad)))
        txtTelefonoFactura.Text = IIf(IsNull(rs!Telefono), "", Trim(rs!Telefono))
        txtColonia.Text = IIf(IsNull(rs!Colonia), "", Trim(rs!Colonia))
        txtCP.Text = IIf(IsNull(rs!CP), "", Trim(rs!CP))
        txtCorreo.Text = IIf(IsNull(rs!CORREO), "", Trim(rs!CORREO))
        
        vlstrTipo = IIf(IsNull(rs!tipo), "OT", rs!tipo)
        vlstrNumRef = IIf(vlstrTipo = "OT", "NULL", IIf(IsNull(rs!clave), 0, rs!clave))
        
        pActivaCHKIEPS '<-----------------------------------------------------------------------------------------------------------------------------------
        If ban = 1 Then
        chkIEPS.Value = IIf(Me.chkIEPS.Enabled = False, vbUnchecked, IIf(vlblnLicenciaIEPS, IIf(rs!bitsujetoaieps = 1, vbChecked, vbUnchecked), vbUnchecked))
        Else
        chkIEPS.Value = vbUnchecked
        End If
        
        '- CASO 7931: Agregado para tener la referencia del paciente -'
        If Trim(vlstrTipoPac) <> "" Then
            vlstrNumPac = IIf(IsNull(rs!clave), "", rs!clave)
        ElseIf ban = 1 Then ' Se buscó en tabla PvDatosFiscales
            If vlstrTipo = "PI" Or vlstrTipo = "PE" Then
                vlstrTipoPac = IIf(vlstrTipo = "PI", "I", "E")
                vlstrNumPac = vlstrNumRef
            Else
                vlstrTipoPac = IIf(IsNull(rs!TipoPac), "", rs!TipoPac)
                vlstrNumPac = IIf(IsNull(rs!cuenta), "", rs!cuenta)
            End If
        End If
        
        '-------------------------------------------------------------
        '*                     Regimen fiscal                        *
        '-------------------------------------------------------------
        
        If stropcion = "EX" Then
            cboRegimenFiscal.ListIndex = 0
        ElseIf stropcion = "IN" Then
            cboRegimenFiscal.ListIndex = 0
        ElseIf stropcion = "EM" Then
            cboRegimenFiscal.ListIndex = flngLocalizaCbo(cboRegimenFiscal, rs!VCHREGIMENFISCAL)
        ElseIf stropcion = "OT" Then
            cboRegimenFiscal.ListIndex = flngLocalizaCbo(cboRegimenFiscal, rs!VCHREGIMENFISCAL)
        ElseIf stropcion = "E" Then
            cboRegimenFiscal.ListIndex = flngLocalizaCbo(cboRegimenFiscal, rs!VCHREGIMENFISCAL)
        ElseIf stropcion = "ME" Then
            cboRegimenFiscal.ListIndex = flngLocalizaCbo(cboRegimenFiscal, rs!VCHREGIMENFISCAL)
        End If
        
        
    End If
    rs.Close
    
    txtBuscaDatos.Text = ""
    sstDatos.Tab = 0
    freBuscarDatos.Enabled = False
    pEnfocaTextBox txtNombreFactura
End Sub

Private Sub txtBuscaDatos_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim vlstrSentencia As String
    
    'If optClave.Value = False Then
    If optBusqueda(0).Value Then 'Externos
        vlstrSentencia = "SELECT intNumPaciente Clave, RTRIM(chrApePaterno) || ' ' || RTRIM(chrApeMaterno) || ' ' || RTRIM(chrNombre) Nombre FROM Externo"
        PSuperBusqueda txtBuscaDatos, vlstrSentencia, lstBuscaDatos, "RTRIM(chrApePaterno) || ' ' || RTRIM(chrApeMaterno) || ' ' || RTRIM(chrNombre)", 20, , "chrApePaterno, chrApeMaterno, chrNombre"
    ElseIf optBusqueda(1).Value Then  'Empresas
        vlstrSentencia = "SELECT intCveEmpresa Clave, vchDescripcion Descripcion from CCEmpresa"
        PSuperBusqueda txtBuscaDatos, vlstrSentencia, lstBuscaDatos, "vchDescripcion", 20, , "vchDescripcion"
    ElseIf optBusqueda(2).Value Then  'Otros PVDatosFiscales
        vlstrSentencia = "SELECT intID Clave, chrNombre Descripcion FROM pvDatosFiscales df"
        pBuscaDatosFiscales txtBuscaDatos, vlstrSentencia, lstBuscaDatos, "chrNombre", 20, , "chrNombre"
    ElseIf optBusqueda(3).Value Then 'Internos
        vlstrSentencia = "SELECT NUMCVEPACIENTE Clave, RTRIM(vchApellidoPaterno) || ' ' || RTRIM(vchApellidoMaterno) || ' ' || RTRIM(vchNombre) Nombre FROM AdPaciente"
        PSuperBusqueda txtBuscaDatos, vlstrSentencia, lstBuscaDatos, "RTRIM(vchApellidoPaterno) || ' ' || RTRIM(vchApellidoMaterno) || ' ' || RTRIM(vchNombre)", 20, , "vchApellidoPaterno, vchApellidoMaterno, vchNombre"
    ElseIf optBusqueda(4).Value Then 'Empleados
        vlstrSentencia = "SELECT INTCVEEMPLEADO Clave, RTRIM(vchApellidoPaterno) || ' ' || RTRIM(vchApellidoMaterno) || ' ' || RTRIM(vchNombre) Nombre FROM NoEmpleado"
        PSuperBusqueda txtBuscaDatos, vlstrSentencia, lstBuscaDatos, "RTRIM(vchApellidoPaterno) || ' ' || RTRIM(vchApellidoMaterno) || ' ' || RTRIM(vchNombre)", 20, , "vchApellidoPaterno, vchApellidoMaterno, vchNombre"
    ElseIf optBusqueda(5).Value Then 'Médicos
        vlstrSentencia = "SELECT INTCVEMEDICO Clave, RTRIM(vchApellidoPaterno) || ' ' || RTRIM(vchApellidoMaterno) || ' ' || RTRIM(vchNombre) Nombre FROM HoMedico"
        PSuperBusqueda txtBuscaDatos, vlstrSentencia, lstBuscaDatos, "RTRIM(vchApellidoPaterno) || ' ' || RTRIM(vchApellidoMaterno) || ' ' || RTRIM(vchNombre)", 20, , "vchApellidoPaterno, vchApellidoMaterno, vchNombre"
    End If
End Sub

Private Sub txtDireccionFactura_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNombreFactura_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtRFC_LostFocus()
    txtRFC_KeyDown 13, 0
End Sub

Private Sub txtTelefonoFactura_GotFocus()
    pSelTextBox txtTelefonoFactura
End Sub

Private Sub txtTelefonoFactura_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn And Not KeyAscii = 45 Then
        KeyAscii = 7
    End If
End Sub

Public Sub pBuscaDatosFiscales(txtTexto As TextBox, vlstrInstruccion As String, lstResultados As ListBox, vlstrCampoFiltro As String, vlintMaxRecords As Integer, Optional vlstrOtroFiltro As String, Optional vlstrOrderBy As String)

    Dim vlintRenglones As Integer
    Dim rsDatos As New ADODB.Recordset
    
    If EntornoSIHO.ConeccionSIHO.State = 0 Then
        EntornoSIHO.ConeccionSIHO.Open
    End If
    
        
     vlstrInstruccion = vlstrInstruccion & " where (intid,chrnombre) in " & _
                                           "(select max(intid),chrnombre" & _
                                           " from pvdatosfiscales" & _
                                           " where " & vlstrCampoFiltro & " like '" & txtTexto.Text & "%' " & vlstrOtroFiltro & _
                                           " group by chrnombre)"
        
    If vlstrOrderBy <> "" Then
        vlstrInstruccion = vlstrInstruccion & " Order by " & vlstrOrderBy
    End If
    
    ' Instruccion para que Deshabilitar el Filtro de la SuperBusqueda con ROWCOUNT
    ' vlstrInstruccion = vlstrInstruccion & " set rowCount 0"
    '-----------------------------------------------------------------------------
    
    lstResultados.Clear
    If txtTexto.Text <> "" Then
        lstResultados.Visible = False
        
        Set rsDatos = frsRegresaRs(vlstrInstruccion, adLockOptimistic, adOpenForwardOnly, vlintMaxRecords)
        For vlintRenglones = 0 To rsDatos.RecordCount - 1
            lstResultados.AddItem rsDatos.Fields(1)
            lstResultados.ItemData(lstResultados.newIndex) = rsDatos.Fields(0)
            rsDatos.MoveNext
        Next vlintRenglones
        
        If rsDatos.RecordCount > 0 Then
            lstResultados.ListIndex = 0
            lstResultados.Enabled = True
        Else
            lstResultados.Enabled = False
        End If
        rsDatos.Close
        lstResultados.Visible = True
    Else
        lstResultados.Enabled = False
    End If
End Sub

Private Sub pCargaUsosCFDI()
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = frsCatalogoSAT("c_UsoCFDI")
    If Not rsTmp.EOF Then
        pLlenarCboRs cboUsoCFDI, rsTmp, 0, 1
        cboUsoCFDI.ListIndex = -1
    End If
End Sub
Private Sub pCargaRegimenFiscal()
    Dim rsTmp As ADODB.Recordset
    cboRegimenFiscal.Clear
    Set rsTmp = frsRegresaRs("SELECT VCHCLAVE, VCHDESCRIPCION FROM GNCATALOGOSATDETALLE WHERE INTIDCATALOGOSAT = 1", adLockReadOnly, adOpenForwardOnly)
        If Not rsTmp.EOF Then
            pLlenarCboRs cboRegimenFiscal, rsTmp, 0, 1
            cboRegimenFiscal.AddItem "<NINGUNO>", 0
            cboRegimenFiscal.ItemData(cboRegimenFiscal.newIndex) = 0
            cboRegimenFiscal.ListIndex = 0
        End If
    rsTmp.Close
End Sub


'Función prototipo para cambiar el valor de un combo a uno en específico
Private Sub pValidaBloqueoRegimen()
    If (txtRFC = "XAXX010101000" Or txtRFC = "XEXX010101000") And vgstrVersionCFDI = "4.0" Then
        SetComboBoxToItem cboRegimenFiscal, "SIN OBLIGACIONES FISCALES"
        'cboUsoCFDI(0).Enabled = False
    'Else
    '    cboUsoCFDI(0).Enabled = Enabled
    End If
End Sub

Private Sub SetComboBoxToItem(Box As ComboBox, Itm As String)
  Dim i%
  For i = 0 To Box.ListCount - 1
    If Box.List(i) = Itm Then
       Box.ListIndex = i
       Exit Sub
    End If
  Next
End Sub

