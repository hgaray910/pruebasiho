VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Begin VB.Form frmMantoFolios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Folios"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7845
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   7845
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTabFolios 
      Height          =   4950
      Left            =   -60
      TabIndex        =   28
      Top             =   0
      Width           =   7950
      _ExtentX        =   14023
      _ExtentY        =   8731
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Registro"
      TabPicture(0)   =   "frmMantoFolios.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Consulta"
      TabPicture(1)   =   "frmMantoFolios.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grdFolios"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Cancelación"
      TabPicture(2)   =   "frmMantoFolios.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame5"
      Tab(2).Control(1)=   "Frame3"
      Tab(2).ControlCount=   2
      Begin VB.Frame Frame5 
         Height          =   1935
         Left            =   -74835
         TabIndex        =   38
         Top             =   570
         Width           =   7530
         Begin VB.TextBox txtMotivo 
            Height          =   315
            Left            =   1260
            TabIndex        =   25
            ToolTipText     =   "Motivo de la cancelación"
            Top             =   1155
            Width           =   5790
         End
         Begin VB.TextBox txtFinalCancelar 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1260
            MaxLength       =   10
            TabIndex        =   24
            ToolTipText     =   "Folio final a cancelar"
            Top             =   795
            Width           =   1485
         End
         Begin VB.TextBox txtInicialCancelar 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1260
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   23
            ToolTipText     =   "Folio inicial a cancelar"
            Top             =   435
            Width           =   1485
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Motivo"
            Height          =   195
            Left            =   180
            TabIndex        =   41
            Top             =   1230
            Width           =   480
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Folio final"
            Height          =   195
            Left            =   180
            TabIndex        =   40
            Top             =   855
            Width           =   660
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Folio inicial"
            Height          =   195
            Left            =   180
            TabIndex        =   39
            Top             =   495
            Width           =   765
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2955
         Left            =   180
         TabIndex        =   0
         Top             =   540
         Width           =   7605
         Begin VB.Frame fraCFD 
            Height          =   795
            Left            =   4920
            TabIndex        =   48
            Top             =   930
            Width           =   2510
            Begin VB.CheckBox chkComprobanteFiscal 
               Caption         =   "Comprobante fiscal digital"
               Enabled         =   0   'False
               Height          =   315
               Left            =   200
               TabIndex        =   51
               ToolTipText     =   "Indíca si es un formato físico o digital"
               Top             =   120
               Width           =   2175
            End
            Begin VB.OptionButton optCFD 
               Caption         =   "CFD"
               Height          =   255
               Index           =   0
               Left            =   360
               TabIndex        =   50
               ToolTipText     =   "Indica si es de tipo Comprobante Fiscal Digital"
               Top             =   470
               Width           =   735
            End
            Begin VB.OptionButton optCFD 
               Caption         =   "CFDi"
               Height          =   255
               Index           =   1
               Left            =   1440
               TabIndex        =   49
               ToolTipText     =   "Indica si es de tipo Comprobante Fiscal Digital por Internet"
               Top             =   470
               Width           =   735
            End
         End
         Begin VB.TextBox txtNumAprobacion 
            Enabled         =   0   'False
            Height          =   315
            Left            =   5955
            Locked          =   -1  'True
            MaxLength       =   14
            TabIndex        =   44
            ToolTipText     =   "Número de aprobación proporcionado por el SAT"
            Top             =   2140
            Width           =   1455
         End
         Begin VB.ComboBox cboDepartamento 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   43
            ToolTipText     =   "Departamento"
            Top             =   1050
            Width           =   3090
         End
         Begin VB.Frame Frame4 
            BorderStyle     =   0  'None
            Height          =   825
            Left            =   1400
            TabIndex        =   30
            Top             =   165
            Width           =   5625
            Begin VB.OptionButton optTipoAplAnt 
               Caption         =   "Aplicación de anticipos"
               Height          =   195
               Left            =   3480
               TabIndex        =   9
               ToolTipText     =   "Tipo de folio aplicación de anticipos"
               Top             =   490
               Width           =   1935
            End
            Begin VB.OptionButton optTipoDonativos 
               Caption         =   "Donativo"
               Height          =   195
               Left            =   3480
               TabIndex        =   8
               ToolTipText     =   "Tipo donativo"
               Top             =   280
               Width           =   1095
            End
            Begin VB.OptionButton optTipoRecibo 
               Caption         =   "Recibo"
               Height          =   195
               Left            =   120
               TabIndex        =   2
               ToolTipText     =   "Tipo de folio recibo"
               Top             =   280
               Width           =   825
            End
            Begin VB.OptionButton optTipoTicket 
               Caption         =   "Ticket"
               Height          =   195
               Left            =   1755
               TabIndex        =   5
               ToolTipText     =   "Tipo ticket"
               Top             =   280
               Width           =   825
            End
            Begin VB.OptionButton optTipoNotaCreditoCargo 
               Caption         =   "Nota de cargo/crédito"
               Height          =   195
               Left            =   3480
               TabIndex        =   7
               ToolTipText     =   "Tipo nota de cargo/crédito"
               Top             =   80
               Width           =   2025
            End
            Begin VB.OptionButton optTipoSalidaDinero 
               Caption         =   "Salidas de dinero"
               Height          =   195
               Left            =   1755
               TabIndex        =   6
               ToolTipText     =   "Tipo salidas de dinero"
               Top             =   490
               Width           =   1620
            End
            Begin VB.OptionButton optTipoNotaCargo 
               Caption         =   "Nota de cargo"
               Height          =   195
               Left            =   120
               TabIndex        =   3
               ToolTipText     =   "Tipo de folio nota de cargo"
               Top             =   490
               Width           =   1395
            End
            Begin VB.OptionButton optTipoNotaCredito 
               Caption         =   "Nota de crédito"
               Height          =   195
               Left            =   1755
               TabIndex        =   4
               ToolTipText     =   "Tipo de folio nota de crédito"
               Top             =   80
               Width           =   1560
            End
            Begin VB.OptionButton optTipoFactura 
               Caption         =   "Factura"
               Height          =   195
               Left            =   120
               TabIndex        =   1
               ToolTipText     =   "Tipo de folio factura"
               Top             =   80
               Width           =   1065
            End
         End
         Begin VB.TextBox txtIdentificador 
            Height          =   315
            Left            =   5955
            MaxLength       =   10
            TabIndex        =   10
            ToolTipText     =   "Identificador del folio con otros departamentos"
            Top             =   1770
            Width           =   1455
         End
         Begin VB.TextBox txtFolioInicial 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   11
            ToolTipText     =   "Folio inicial"
            Top             =   1425
            Width           =   1930
         End
         Begin VB.TextBox txtFolioFinal 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   12
            ToolTipText     =   "Folio final"
            Top             =   1800
            Width           =   1930
         End
         Begin VB.TextBox txtFolioActual 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1560
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   13
            ToolTipText     =   "Folio a imprimir"
            Top             =   2175
            Width           =   1930
         End
         Begin VB.TextBox txtMensaje 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2340
            MaxLength       =   3
            TabIndex        =   14
            ToolTipText     =   "Mensaje de aviso al faltar folios"
            Top             =   2550
            Width           =   660
         End
         Begin MSMask.MaskEdBox mskFechaAprobacion 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   3
            EndProperty
            Height          =   315
            Left            =   5955
            TabIndex        =   47
            ToolTipText     =   "Fecha de aprobación proporcionada por el SAT"
            Top             =   2520
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Format          =   "dd/mmm/yyyy "
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin VB.Label Label13 
            Caption         =   "Fecha de aprobación"
            Enabled         =   0   'False
            Height          =   255
            Left            =   4260
            TabIndex        =   46
            Top             =   2595
            Width           =   1575
         End
         Begin VB.Label Label12 
            Caption         =   "Número de aprobación"
            Enabled         =   0   'False
            Height          =   195
            Left            =   4260
            TabIndex        =   45
            Top             =   2230
            Width           =   1695
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "folios"
            Height          =   195
            Left            =   3105
            TabIndex        =   42
            Top             =   2610
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tipo"
            Height          =   195
            Left            =   180
            TabIndex        =   37
            Top             =   220
            Width           =   315
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Departamento"
            Height          =   195
            Left            =   180
            TabIndex        =   36
            Top             =   1110
            Width           =   1005
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Identificador"
            Height          =   210
            Left            =   4260
            TabIndex        =   35
            Top             =   1845
            Width           =   870
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Folio inicial"
            Height          =   195
            Left            =   180
            TabIndex        =   34
            Top             =   1485
            Width           =   765
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Folio final"
            Height          =   195
            Left            =   180
            TabIndex        =   33
            Top             =   1860
            Width           =   660
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Folio actual"
            Height          =   195
            Left            =   180
            TabIndex        =   32
            Top             =   2235
            Width           =   810
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Enviar mensaje cuando falten"
            Height          =   195
            Left            =   165
            TabIndex        =   31
            Top             =   2610
            Width           =   2100
         End
      End
      Begin VB.Frame Frame2 
         Height          =   705
         Left            =   2190
         TabIndex        =   29
         Top             =   3600
         Width           =   3585
         Begin VB.CommandButton cmdTop 
            Height          =   495
            Left            =   60
            Picture         =   "frmMantoFolios.frx":0054
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Primer folio"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdBack 
            Height          =   495
            Left            =   555
            Picture         =   "frmMantoFolios.frx":0456
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Anterior folio"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdLocate 
            Height          =   495
            Left            =   1050
            Picture         =   "frmMantoFolios.frx":0948
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Búsqueda de folios"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdNext 
            Height          =   495
            Left            =   1545
            Picture         =   "frmMantoFolios.frx":0E3A
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Siguiente folio"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdEnd 
            Height          =   495
            Left            =   2040
            Picture         =   "frmMantoFolios.frx":132C
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Ultimo folio"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdDelete 
            Height          =   495
            Left            =   3030
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmMantoFolios.frx":181E
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Ultimo folio"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdSave 
            Height          =   495
            Left            =   2535
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmMantoFolios.frx":1F20
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Grabar folio"
            Top             =   150
            Width           =   495
         End
      End
      Begin VB.Frame Frame3 
         Height          =   675
         Left            =   -71370
         TabIndex        =   27
         Top             =   2805
         Width           =   585
         Begin VB.CommandButton cmdCancelar 
            Height          =   495
            Left            =   45
            Picture         =   "frmMantoFolios.frx":2262
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   135
            UseMaskColor    =   -1  'True
            Width           =   495
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdFolios 
         Height          =   3810
         Left            =   -74870
         TabIndex        =   22
         Top             =   510
         Width           =   7680
         _ExtentX        =   13547
         _ExtentY        =   6720
         _Version        =   393216
         Rows            =   0
         FixedRows       =   0
         SelectionMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
End
Attribute VB_Name = "frmMantoFolios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------
' Programa para dar mantenimiento a la tabla "RegistroFolio" (registro de folios de
' documentos) y "FolioCancelado", este programa muestra los folios de la clave de
' departamento que se trae en memoria (vgintNumeroDepartamento)
' Fecha de programación: Lunes 15 de Enero de 2001
'-------------------------------------------------------------------------------------
' Ultimas modificaciones, especificar:
' 30-May-2001 : Se cambio para que grabara en RegistroFolio en vez de RegistroFolio igual
' para los cancelados FolioCancelado en vez de FolioCancelado
' 29-Abr-2003 : Se incluyo el concepto de "Salida de dinero"
'-------------------------------------------------------------------------------------
Option Explicit

Dim vlstrSentencia As String
Dim rsDepartamentos As New ADODB.Recordset
Dim rsRegistroFolio As New ADODB.Recordset
Dim rsFolioCancelado As New ADODB.Recordset
Dim rsFolioUnicoNotas As New ADODB.Recordset
Dim vlblnConsulta As Boolean
Dim vllngCveRegistro As Long
Dim vlstrTipoActivo As String


Private Sub pMuestraFolio()
    On Error GoTo NotificaError
    Dim rsNombreDepartamento As New ADODB.Recordset
    
    vlblnConsulta = True
    
    cboDepartamento.ListIndex = flngLocalizaCbo(cboDepartamento, CStr(rsRegistroFolio!SMIDEPARTAMENTO))
    
    If rsRegistroFolio!intNumeroActual > rsRegistroFolio!intNumeroFinal Then
        
        'Deshabilita todo
        pHabilitaTextBox 0, 0, 0, 0, 0, 0, 0
        Label12.Enabled = False
        Label13.Enabled = False
        txtNumAprobacion.Enabled = False
        mskFechaAprobacion.Enabled = False
        optCFD(0).Enabled = False
        optCFD(1).Enabled = False
        
    Else
        
        If rsRegistroFolio!intNumeroInicial <> rsRegistroFolio!intNumeroActual Then
            'Deshabilita número inicial e identificador
            pHabilitaTextBox 0, 0, 1, 1, 1, 1, 1
            Label12.Enabled = True
            Label13.Enabled = True
            optCFD(0).Enabled = False
            optCFD(1).Enabled = False
        Else
            'Habilita todo (verificando el tipo de CFD)
            If optCFD(1).Value = True Then
                pHabilitaTextBox 1, 1, 1, 1, 1, 0, 0
            Else
                pHabilitaTextBox 1, 1, 1, 1, 1, 1, 1
            End If
            Label12.Enabled = True
            Label13.Enabled = True
            optCFD(0).Enabled = True
            optCFD(1).Enabled = True
        End If
        
    End If
    
    txtIdentificador.Text = Trim(rsRegistroFolio!chrCveDocumento)
    txtFolioInicial.Text = Trim(rsRegistroFolio!intNumeroInicial)
    txtFolioFinal.Text = Trim(rsRegistroFolio!intNumeroFinal)
    txtFolioActual.Text = Trim(rsRegistroFolio!intNumeroActual)
    txtMensaje.Text = Trim(rsRegistroFolio!smiFoliosAviso)
    
'---------------------------------------------- FACTURA
    If rsRegistroFolio!chrTipoDocumento = "FA" Then
        
        optTipoFactura.Value = True
        If rsRegistroFolio!BitTipo = 1 Then
        
            chkComprobanteFiscal.Value = vbChecked
            txtNumAprobacion.Text = IIf(IsNull(rsRegistroFolio!vchNumAprobacion), "", Trim(rsRegistroFolio!vchNumAprobacion))
            mskFechaAprobacion.Text = IIf(IsNull(rsRegistroFolio!dtmFechaAprobacion), "  /  /    ", Trim(rsRegistroFolio!dtmFechaAprobacion))
            chkComprobanteFiscal.Enabled = False
            Label12.Enabled = True
            Label13.Enabled = True

            'Se verifica el tipo de CFD
            optCFD(0).Value = IIf(rsRegistroFolio!BitCFDi = 0, True, False)
            optCFD(1).Value = IIf(rsRegistroFolio!BitCFDi = 1, True, False)

            If optCFD(0).Value = True Then
                txtNumAprobacion.Enabled = True
                mskFechaAprobacion.Enabled = True
            Else
                txtNumAprobacion.Enabled = False
                mskFechaAprobacion.Enabled = False
            End If

            If rsRegistroFolio!intNumeroActual > rsRegistroFolio!intNumeroFinal Then
                'Deshabilita todo
                Label12.Enabled = False
                Label13.Enabled = False
                optCFD(0).Enabled = False
                optCFD(1).Enabled = False
                txtNumAprobacion.Enabled = False
                mskFechaAprobacion.Enabled = False
            End If
            
        Else
        
            chkComprobanteFiscal.Value = vbUnchecked
            txtNumAprobacion.Text = ""
            mskFechaAprobacion.Mask = ""
            mskFechaAprobacion.Text = ""
            chkComprobanteFiscal.Enabled = False
            txtNumAprobacion.Enabled = False
            mskFechaAprobacion.Enabled = False
            Label12.Enabled = False
            Label13.Enabled = False
            optCFD(0).Enabled = False
            optCFD(1).Enabled = False
            optCFD(0).Value = False
            optCFD(1).Value = False
            
            If rsRegistroFolio!intNumeroActual > rsRegistroFolio!intNumeroFinal Then
                'Deshabilita todo
                Label12.Enabled = False
                Label13.Enabled = False
                optCFD(0).Enabled = False
                optCFD(1).Enabled = False
                optCFD(0).Value = False
                optCFD(1).Value = False
                txtNumAprobacion.Enabled = False
                mskFechaAprobacion.Enabled = False
            End If
            
        End If
        
    '---------------------------------------------- RECIBO
    ElseIf rsRegistroFolio!chrTipoDocumento = "RE" Then
    
        optTipoRecibo.Value = True
        If rsRegistroFolio!BitTipo = 1 Then
            chkComprobanteFiscal.Value = vbChecked
            txtNumAprobacion.Text = IIf(IsNull(rsRegistroFolio!vchNumAprobacion), "", Trim(rsRegistroFolio!vchNumAprobacion))
            mskFechaAprobacion.Text = IIf(IsNull(rsRegistroFolio!dtmFechaAprobacion), "  /  /    ", Trim(rsRegistroFolio!dtmFechaAprobacion))
            chkComprobanteFiscal.Enabled = False
            Label12.Enabled = True
            Label13.Enabled = True

            'Se verifica el tipo de CFD
            optCFD(0).Value = IIf(rsRegistroFolio!BitCFDi = 0, True, False)
            optCFD(1).Value = IIf(rsRegistroFolio!BitCFDi = 1, True, False)

            If optCFD(0).Value = True Then
                txtNumAprobacion.Enabled = True
                mskFechaAprobacion.Enabled = True
            Else
                txtNumAprobacion.Enabled = False
                mskFechaAprobacion.Enabled = False
            End If

            If rsRegistroFolio!intNumeroActual > rsRegistroFolio!intNumeroFinal Then
                'Deshabilita todo
                Label12.Enabled = False
                Label13.Enabled = False
                optCFD(0).Enabled = False
                optCFD(1).Enabled = False
                txtNumAprobacion.Enabled = False
                mskFechaAprobacion.Enabled = False
            End If
        Else
            chkComprobanteFiscal.Value = vbUnchecked
            txtNumAprobacion.Text = ""
            mskFechaAprobacion.Mask = ""
            mskFechaAprobacion.Text = ""
            chkComprobanteFiscal.Enabled = False
            txtNumAprobacion.Enabled = False
            mskFechaAprobacion.Enabled = False
            Label12.Enabled = False
            Label13.Enabled = False
            optCFD(0).Enabled = False
            optCFD(1).Enabled = False
            optCFD(0).Value = False
            optCFD(1).Value = False
            
            If rsRegistroFolio!intNumeroActual > rsRegistroFolio!intNumeroFinal Then
                'Deshabilita todo
                Label12.Enabled = False
                Label13.Enabled = False
                optCFD(0).Enabled = False
                optCFD(1).Enabled = False
                optCFD(0).Value = False
                optCFD(1).Value = False
                txtNumAprobacion.Enabled = False
                mskFechaAprobacion.Enabled = False
            End If
        End If
        
    '---------------------------------------------- NOTA DE CREDITO
    ElseIf rsRegistroFolio!chrTipoDocumento = "NC" Then
        
        optTipoNotaCredito.Value = True
        
        If rsRegistroFolio!BitTipo = 1 Then
        
            chkComprobanteFiscal.Value = vbChecked
            txtNumAprobacion.Text = IIf(IsNull(rsRegistroFolio!vchNumAprobacion), "", Trim(rsRegistroFolio!vchNumAprobacion))
            mskFechaAprobacion.Text = IIf(IsNull(rsRegistroFolio!dtmFechaAprobacion), "  /  /    ", Trim(rsRegistroFolio!dtmFechaAprobacion))
            chkComprobanteFiscal.Enabled = False
            Label12.Enabled = True
            Label13.Enabled = True

            'Se verifica el tipo de CFD
            optCFD(0).Value = IIf(rsRegistroFolio!BitCFDi = 0, True, False)
            optCFD(1).Value = IIf(rsRegistroFolio!BitCFDi = 1, True, False)
            
            If optCFD(0).Value = True Then
                txtNumAprobacion.Enabled = True
                mskFechaAprobacion.Enabled = True
            Else
                txtNumAprobacion.Enabled = False
                mskFechaAprobacion.Enabled = False
            End If
            
            If rsRegistroFolio!intNumeroActual > rsRegistroFolio!intNumeroFinal Then
                'Deshabilita todo
                Label12.Enabled = False
                Label13.Enabled = False
                optCFD(0).Enabled = False
                optCFD(1).Enabled = False
                txtNumAprobacion.Enabled = False
                mskFechaAprobacion.Enabled = False
            End If
            
        Else
        
            chkComprobanteFiscal.Value = vbUnchecked
            txtNumAprobacion.Text = ""
            mskFechaAprobacion.Mask = ""
            mskFechaAprobacion.Text = ""
            chkComprobanteFiscal.Enabled = False
            txtNumAprobacion.Enabled = False
            mskFechaAprobacion.Enabled = False
            Label12.Enabled = False
            Label13.Enabled = False
            optCFD(0).Enabled = False
            optCFD(1).Enabled = False
            optCFD(0).Value = False
            optCFD(1).Value = False
            
            If rsRegistroFolio!intNumeroActual > rsRegistroFolio!intNumeroFinal Then
                'Deshabilita todo
                Label12.Enabled = False
                Label13.Enabled = False
                optCFD(0).Enabled = False
                optCFD(1).Enabled = False
                optCFD(0).Value = False
                optCFD(1).Value = False
                txtNumAprobacion.Enabled = False
                mskFechaAprobacion.Enabled = False
            End If
            
        End If

    '---------------------------------------------- NOTA DE CARGO
    ElseIf rsRegistroFolio!chrTipoDocumento = "NA" Then
        optTipoNotaCargo.Value = True
        
        If rsRegistroFolio!BitTipo = 1 Then
        
            chkComprobanteFiscal.Value = vbChecked
            txtNumAprobacion.Text = IIf(IsNull(rsRegistroFolio!vchNumAprobacion), "", Trim(rsRegistroFolio!vchNumAprobacion))
            mskFechaAprobacion.Text = IIf(IsNull(rsRegistroFolio!dtmFechaAprobacion), "  /  /    ", Trim(rsRegistroFolio!dtmFechaAprobacion))
            chkComprobanteFiscal.Enabled = False
            Label12.Enabled = True
            Label13.Enabled = True

            'Se verifica el tipo de CFD
            optCFD(0).Value = IIf(rsRegistroFolio!BitCFDi = 0, True, False)
            optCFD(1).Value = IIf(rsRegistroFolio!BitCFDi = 1, True, False)
            
            If optCFD(0).Value = True Then
                txtNumAprobacion.Enabled = True
                mskFechaAprobacion.Enabled = True
            Else
                txtNumAprobacion.Enabled = False
                mskFechaAprobacion.Enabled = False
            End If
            
            If rsRegistroFolio!intNumeroActual > rsRegistroFolio!intNumeroFinal Then
                'Deshabilita todo
                Label12.Enabled = False
                Label13.Enabled = False
                optCFD(0).Enabled = False
                optCFD(1).Enabled = False
                txtNumAprobacion.Enabled = False
                mskFechaAprobacion.Enabled = False
            End If
            
        Else
        
            chkComprobanteFiscal.Value = vbUnchecked
            txtNumAprobacion.Text = ""
            mskFechaAprobacion.Mask = ""
            mskFechaAprobacion.Text = ""
            chkComprobanteFiscal.Enabled = False
            txtNumAprobacion.Enabled = False
            mskFechaAprobacion.Enabled = False
            Label12.Enabled = False
            Label13.Enabled = False
            optCFD(0).Enabled = False
            optCFD(1).Enabled = False
            optCFD(0).Value = False
            optCFD(1).Value = False
            
            If rsRegistroFolio!intNumeroActual > rsRegistroFolio!intNumeroFinal Then
                'Deshabilita todo
                Label12.Enabled = False
                Label13.Enabled = False
                optCFD(0).Enabled = False
                optCFD(1).Enabled = False
                optCFD(0).Value = False
                optCFD(1).Value = False
                txtNumAprobacion.Enabled = False
                mskFechaAprobacion.Enabled = False
            End If
            
        End If

    '---------------------------------------------- TICKET
    ElseIf rsRegistroFolio!chrTipoDocumento = "TI" Then
        
        optTipoTicket.Value = True
        txtNumAprobacion.Enabled = False
        mskFechaAprobacion.Enabled = False
        chkComprobanteFiscal.Enabled = False
        optCFD(0).Enabled = False
        optCFD(1).Enabled = False
        optCFD(0).Value = False
        optCFD(1).Value = False
        Label12.Enabled = False
        Label13.Enabled = False
        chkComprobanteFiscal.Value = vbUnchecked
        mskFechaAprobacion.Mask = ""
        mskFechaAprobacion.Text = ""
        txtNumAprobacion.Text = ""
        If rsRegistroFolio!intNumeroActual > rsRegistroFolio!intNumeroFinal Then
            'Deshabilita todo
            Label12.Enabled = False
            Label13.Enabled = False
            txtNumAprobacion.Enabled = False
            mskFechaAprobacion.Enabled = False
            chkComprobanteFiscal.Value = vbUnchecked
            mskFechaAprobacion.Mask = ""
            mskFechaAprobacion.Text = ""
            txtNumAprobacion.Text = ""
            optCFD(0).Enabled = False
            optCFD(1).Enabled = False
            optCFD(0).Value = False
            optCFD(1).Value = False
        End If
    
    '---------------------------------------------- NOTA DE CREDITO/CARGO
    ElseIf rsRegistroFolio!chrTipoDocumento = "CC" Then
        optTipoNotaCreditoCargo.Value = True
        
        If rsRegistroFolio!BitTipo = 1 Then
        
            chkComprobanteFiscal.Value = vbChecked
            txtNumAprobacion.Text = IIf(IsNull(rsRegistroFolio!vchNumAprobacion), "", Trim(rsRegistroFolio!vchNumAprobacion))
            mskFechaAprobacion.Text = IIf(IsNull(rsRegistroFolio!dtmFechaAprobacion), "  /  /    ", Trim(rsRegistroFolio!dtmFechaAprobacion))
            chkComprobanteFiscal.Enabled = False
            Label12.Enabled = True
            Label13.Enabled = True

            'Se verifica el tipo de CFD
            optCFD(0).Value = IIf(rsRegistroFolio!BitCFDi = 0, True, False)
            optCFD(1).Value = IIf(rsRegistroFolio!BitCFDi = 1, True, False)
            
            If optCFD(0).Value = True Then
                txtNumAprobacion.Enabled = True
                mskFechaAprobacion.Enabled = True
            Else
                txtNumAprobacion.Enabled = False
                mskFechaAprobacion.Enabled = False
            End If
            
            If rsRegistroFolio!intNumeroActual > rsRegistroFolio!intNumeroFinal Then
                'Deshabilita todo
                Label12.Enabled = False
                Label13.Enabled = False
                optCFD(0).Enabled = False
                optCFD(1).Enabled = False
                txtNumAprobacion.Enabled = False
                mskFechaAprobacion.Enabled = False
            End If
            
        Else
        
            chkComprobanteFiscal.Value = vbUnchecked
            txtNumAprobacion.Text = ""
            mskFechaAprobacion.Mask = ""
            mskFechaAprobacion.Text = ""
            chkComprobanteFiscal.Enabled = False
            txtNumAprobacion.Enabled = False
            mskFechaAprobacion.Enabled = False
            Label12.Enabled = False
            Label13.Enabled = False
            optCFD(0).Enabled = False
            optCFD(1).Enabled = False
            optCFD(0).Value = False
            optCFD(1).Value = False
            
            If rsRegistroFolio!intNumeroActual > rsRegistroFolio!intNumeroFinal Then
                'Deshabilita todo
                Label12.Enabled = False
                Label13.Enabled = False
                optCFD(0).Enabled = False
                optCFD(1).Enabled = False
                optCFD(0).Value = False
                optCFD(1).Value = False
                txtNumAprobacion.Enabled = False
                mskFechaAprobacion.Enabled = False
            End If
            
        End If
        
    '---------------------------------------------- DONATIVO
    ElseIf rsRegistroFolio!chrTipoDocumento = "DO" Then
        
        optTipoDonativos.Value = True
        
        If rsRegistroFolio!BitTipo = 1 Then
        
            chkComprobanteFiscal.Value = vbChecked
            txtNumAprobacion.Text = IIf(IsNull(rsRegistroFolio!vchNumAprobacion), "", Trim(rsRegistroFolio!vchNumAprobacion))
            mskFechaAprobacion.Text = IIf(IsNull(rsRegistroFolio!dtmFechaAprobacion), "  /  /    ", Trim(rsRegistroFolio!dtmFechaAprobacion))
            chkComprobanteFiscal.Enabled = False
            Label12.Enabled = True
            Label13.Enabled = True

            'Se verifica el tipo de CFD
            optCFD(0).Value = IIf(rsRegistroFolio!BitCFDi = 0, True, False)
            optCFD(1).Value = IIf(rsRegistroFolio!BitCFDi = 1, True, False)
            
            If optCFD(0).Value = True Then
                txtNumAprobacion.Enabled = True
                mskFechaAprobacion.Enabled = True
            Else
                txtNumAprobacion.Enabled = False
                mskFechaAprobacion.Enabled = False
            End If
            
            If rsRegistroFolio!intNumeroActual > rsRegistroFolio!intNumeroFinal Then
                'Deshabilita todo
                Label12.Enabled = False
                Label13.Enabled = False
                optCFD(0).Enabled = False
                optCFD(1).Enabled = False
                txtNumAprobacion.Enabled = False
                mskFechaAprobacion.Enabled = False
            End If
            
        Else
        
            chkComprobanteFiscal.Value = vbUnchecked
            txtNumAprobacion.Text = ""
            mskFechaAprobacion.Mask = ""
            mskFechaAprobacion.Text = ""
            chkComprobanteFiscal.Enabled = False
            txtNumAprobacion.Enabled = False
            mskFechaAprobacion.Enabled = False
            Label12.Enabled = False
            Label13.Enabled = False
            optCFD(0).Enabled = False
            optCFD(1).Enabled = False
            optCFD(0).Value = False
            optCFD(1).Value = False
            
            If rsRegistroFolio!intNumeroActual > rsRegistroFolio!intNumeroFinal Then
                'Deshabilita todo
                Label12.Enabled = False
                Label13.Enabled = False
                optCFD(0).Enabled = False
                optCFD(1).Enabled = False
                optCFD(0).Value = False
                optCFD(1).Value = False
                txtNumAprobacion.Enabled = False
                mskFechaAprobacion.Enabled = False
            End If
            
        End If
    ElseIf rsRegistroFolio!chrTipoDocumento = "AA" Then
        optTipoAplAnt.Value = True
        optTipoAplAnt_Click
        

    Else
    '---------------------------------------------- SALIDA DE DINERO
        optTipoSalidaDinero.Value = True
        txtNumAprobacion.Enabled = False
        mskFechaAprobacion.Enabled = False
        chkComprobanteFiscal.Enabled = False
        optCFD(0).Enabled = False
        optCFD(1).Enabled = False
        optCFD(0).Value = False
        optCFD(1).Value = False
        Label12.Enabled = False
        Label13.Enabled = False
        chkComprobanteFiscal.Value = vbUnchecked
        mskFechaAprobacion.Mask = ""
        mskFechaAprobacion.Text = ""
        txtNumAprobacion.Text = ""
        If rsRegistroFolio!intNumeroActual > rsRegistroFolio!intNumeroFinal Then
            'Deshabilita todo
            Label12.Enabled = False
            Label13.Enabled = False
            txtNumAprobacion.Enabled = False
            mskFechaAprobacion.Enabled = False
            chkComprobanteFiscal.Value = vbUnchecked
            mskFechaAprobacion.Mask = ""
            mskFechaAprobacion.Text = ""
            txtNumAprobacion.Text = ""
            optCFD(0).Enabled = False
            optCFD(1).Enabled = False
            optCFD(0).Value = False
            optCFD(1).Value = False
        End If
    End If
    
    'Validación final para las etiquetas
    Label12.Enabled = txtNumAprobacion.Enabled
    Label13.Enabled = mskFechaAprobacion.Enabled

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pMuestraFolio"))
End Sub

Private Sub pHabilitaTextBox(b As Integer, c As Integer, d As Integer, e As Integer, f As Integer, g As Integer, h As Integer)
    On Error GoTo NotificaError
    
    If b = 1 Then
        txtIdentificador.Enabled = True
    Else
        txtIdentificador.Enabled = False
    End If
    If c = 1 Then
        txtFolioInicial.Enabled = True
    Else
        txtFolioInicial.Enabled = False
    End If
    If d = 1 Then
        txtFolioFinal.Enabled = True
    Else
        txtFolioFinal.Enabled = False
    End If
    If e = 1 Then
        txtFolioActual.Enabled = True
    Else
        txtFolioActual.Enabled = False
    End If
    If f = 1 Then
        txtMensaje.Enabled = True
    Else
        txtMensaje.Enabled = False
    End If
    If g = 1 Then
        txtNumAprobacion.Enabled = True
    Else
        txtNumAprobacion.Enabled = False
    End If
    If h = 1 Then
        mskFechaAprobacion.Enabled = True
    Else
        mskFechaAprobacion.Enabled = False
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pHabilitaTextBox"))
End Sub

Private Sub pHabilita(vlb1 As Integer, vlb2 As Integer, vlb3 As Integer, vlb4 As Integer, vlb5 As Integer, vlb6 As Integer, vlb7 As Integer)
    On Error GoTo NotificaError
    
    If vlb1 = 1 Then
        cmdTop.Enabled = True
    Else
        cmdTop.Enabled = False
    End If
    If vlb2 = 1 Then
        cmdBack.Enabled = True
    Else
        cmdBack.Enabled = False
    End If
    If vlb3 = 1 Then
        cmdLocate.Enabled = True
    Else
        cmdLocate.Enabled = False
    End If
    If vlb4 = 1 Then
        cmdNext.Enabled = True
    Else
        cmdNext.Enabled = False
    End If
    If vlb5 = 1 Then
        cmdEnd.Enabled = True
    Else
        cmdEnd.Enabled = False
    End If
    If vlb6 = 1 Then
        cmdSave.Enabled = True
    Else
        cmdSave.Enabled = False
    End If
    If vlb7 = 1 Then
        cmdDelete.Enabled = True
    Else
        cmdDelete.Enabled = False
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pHabilita"))
End Sub

Private Sub pLimpia()
    On Error GoTo NotificaError
    Dim rsConsultaFolios As New ADODB.Recordset
    
    vlblnConsulta = False
    
    If vlstrTipoActivo = "Factura" Then
        optTipoFactura.Value = True
    ElseIf vlstrTipoActivo = "Recibo" Then
        optTipoRecibo.Value = True
    ElseIf vlstrTipoActivo = "NCargo" Then
        optTipoNotaCargo.Value = True
    ElseIf vlstrTipoActivo = "NCredito" Then
        optTipoNotaCredito.Value = True
    ElseIf vlstrTipoActivo = "Ticket" Then
        optTipoTicket.Value = True
    ElseIf vlstrTipoActivo = "SDinero" Then
        optTipoSalidaDinero.Value = True
    ElseIf vlstrTipoActivo = "NCC" Then
        optTipoNotaCreditoCargo.Value = True
    ElseIf vlstrTipoActivo = "Donativo" Then
        optTipoDonativos.Value = True
    Else: optTipoFactura.Value = True
    End If
    
    If cboDepartamento.ListCount = 0 Then
        Call MsgBox(SIHOMsg("12") & Chr(13) & "Dato:" & cboDepartamento.ToolTipText, vbExclamation, "Mensaje")
        Unload Me
        Exit Sub
    End If
    
    If cgstrModulo = "SI" Then
        cboDepartamento.ListIndex = 0
    Else
        cboDepartamento.ListIndex = flngLocalizaCbo(cboDepartamento, CStr(vgintNumeroDepartamento))
    End If
    
    chkComprobanteFiscal.Value = vbUnchecked
    optCFD(0).Value = False
    optCFD(1).Value = False
    optCFD(0).Enabled = False
    optCFD(1).Enabled = False
    txtIdentificador.Text = ""
    txtFolioFinal.Text = ""
    txtFolioInicial.Text = ""
    txtFolioActual.Text = ""
    txtMensaje.Text = ""
    txtInicialCancelar.Text = ""
    txtFinalCancelar.Text = ""
    txtMotivo.Text = ""
    txtNumAprobacion.Text = ""
    txtNumAprobacion.Locked = True
    mskFechaAprobacion.Enabled = False
    mskFechaAprobacion.Mask = ""
    mskFechaAprobacion.Text = ""
    grdFolios.Rows = 0
    
    If rsRegistroFolio.RecordCount = 0 Then
        SSTabFolios.TabEnabled(1) = False
        pHabilita 0, 0, 0, 0, 0, 0, 0
    Else
        pHabilita 1, 1, 1, 1, 1, 0, 0
        SSTabFolios.TabEnabled(1) = True
        
        vgstrParametrosSP = vgintClaveEmpresaContable
        Set rsConsultaFolios = frsEjecuta_SP(vgstrParametrosSP, "SP_GNSELFOLIOS")
        If rsConsultaFolios.RecordCount > 0 Then
            pLlenarMshFGrdRs grdFolios, rsConsultaFolios, 9
            With grdFolios
                .FormatString = "|Departamento|Documento|Identificador|Tipo|Fecha|Inicial|Final|Actual|Estado"
                .ColWidth(0) = 100
                .ColWidth(1) = 2000 'Departamento
                .ColWidth(2) = 2250 'Documento
                .ColWidth(3) = 1200 'Identificador
                .ColWidth(4) = 600 'Tipo
                .ColWidth(5) = 1000 'Fecha
                .ColWidth(6) = 1200 'Inicial
                .ColWidth(7) = 1200 'Final
                .ColWidth(8) = 1200 'Actual
                .ColWidth(9) = 850 'Estatus
                .ColWidth(10) = 0
            End With

        End If
    End If
    pHabilitaTextBox 1, 1, 1, 1, 1, 1, 1
    
    If optTipoFactura.Value = True Or optTipoRecibo.Value = True Or optTipoNotaCargo.Value = True Or optTipoNotaCredito.Value = True Or optTipoNotaCreditoCargo.Value = True Or optTipoDonativos.Value = True Then
        chkComprobanteFiscal.Enabled = True
    End If
    
    SSTabFolios.TabEnabled(2) = False
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pLimpia"))
End Sub


Private Sub cboDepartamento_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        txtIdentificador.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboDepartamentos_KeyPress"))
End Sub

Private Sub cboDepartamento_GotFocus()
    On Error GoTo NotificaError
    
    pHabilita 0, 0, 0, 0, 0, 1, 0

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboDepartamento_GotFocus"))
End Sub

Private Sub chkComprobanteFiscal_Click()

    If chkComprobanteFiscal.Value = vbChecked Then
        'Se habilitan los campos de opción de CFD
        
            'Se verifica si se habilitan o no....
            If Trim(txtFolioInicial.Text) <> Trim(txtFolioActual.Text) Then
                optCFD(0).Enabled = False
                optCFD(1).Enabled = False
            Else
                optCFD(0).Enabled = True
                optCFD(1).Enabled = True
            End If
        
        'Se habilita automáticamente la opción de CFD
        optCFD(0).Value = True
        
        If optCFD(0).Value = True Then
            'Se habilitan los campos si la opción CFD está activada
            Label12.Enabled = True
            Label13.Enabled = True
            txtNumAprobacion.Enabled = True
            mskFechaAprobacion.Enabled = True
            mskFechaAprobacion.Mask = "##/##/####"
        Else
            Label12.Enabled = False
            Label13.Enabled = False
            mskFechaAprobacion.Mask = ""
            mskFechaAprobacion.Text = ""
            txtNumAprobacion.Text = ""
        End If
    Else
        optCFD(0).Enabled = False
        optCFD(1).Enabled = False
        optCFD(0).Value = False
        optCFD(1).Value = False
        Label12.Enabled = False
        Label13.Enabled = False
        mskFechaAprobacion.Mask = ""
        mskFechaAprobacion.Text = ""
        txtNumAprobacion.Text = ""
        mskFechaAprobacion.Enabled = False
        txtNumAprobacion.Enabled = False
    End If

End Sub

Private Sub chkComprobanteFiscal_GotFocus()
    On Error GoTo NotificaError
    
    pHabilita 0, 0, 0, 0, 0, 1, 0
    
    If chkComprobanteFiscal.Value = vbChecked Then
    
        'Se habilitan los campos de opción de CFD
        optCFD(0).Enabled = True
        optCFD(1).Enabled = True
        
        'Se habilita automáticamente la opción de CFD
        optCFD(0).Value = True
        
        If optCFD(0).Value = True Then
        'Se habilitan los campos si la opción CFD está activada
            Label12.Enabled = True
            Label13.Enabled = True
            txtNumAprobacion.Enabled = True
            mskFechaAprobacion.Enabled = True
            mskFechaAprobacion.Mask = "##/##/####"
        Else
            Label12.Enabled = False
            Label13.Enabled = False
            mskFechaAprobacion.Mask = ""
            mskFechaAprobacion.Text = ""
            txtNumAprobacion.Text = ""
        End If
    Else
        optCFD(0).Enabled = False
        optCFD(1).Enabled = False
        optCFD(0).Value = False
        optCFD(1).Value = False
        Label12.Enabled = False
        Label13.Enabled = False
        mskFechaAprobacion.Mask = ""
        mskFechaAprobacion.Text = ""
        txtNumAprobacion.Text = ""
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":chkComprobanteFiscal_GotFocus"))
End Sub

Private Sub chkComprobanteFiscal_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        pHabilita 0, 0, 0, 0, 0, 1, 0
        If chkComprobanteFiscal.Value = vbChecked Then
            optCFD(0).SetFocus
        Else
            txtIdentificador.SetFocus
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":optTipoRecibo_KeyPress"))
End Sub

Private Sub chkComprobanteFiscal_LostFocus()
On Error GoTo NotificaError

    If chkComprobanteFiscal.Value = vbChecked Then
        'Se habilitan los campos de opción de CFD
        optCFD(0).Enabled = True
        optCFD(1).Enabled = True
        
        'Se habilita automáticamente la opción de CFD
        optCFD(0).Value = True
        
        If optCFD(0).Value = True Then
        'Se habilitan los campos si la opción CFD está activada
            Label12.Enabled = True
            Label13.Enabled = True
            txtNumAprobacion.Enabled = True
            mskFechaAprobacion.Enabled = True
        Else
            Label12.Enabled = False
            Label13.Enabled = False
            mskFechaAprobacion.Enabled = False
            txtNumAprobacion.Enabled = False
        End If
    Else
        optCFD(0).Enabled = False
        optCFD(1).Enabled = False
        optCFD(0).Value = False
        optCFD(1).Value = False
        Label12.Enabled = False
        Label13.Enabled = False
        mskFechaAprobacion.Mask = ""
        mskFechaAprobacion.Text = ""
        txtNumAprobacion.Text = ""
        mskFechaAprobacion.Enabled = False
        txtNumAprobacion.Enabled = False
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":chkComprobanteFiscal_LostFocus"))
End Sub

Private Sub cmdBack_Click()
    On Error GoTo NotificaError
    
    If rsRegistroFolio.RecordCount = 0 Then
        ' No existe información.
        MsgBox SIHOMsg(13), vbOKOnly + vbInformation, "Mensaje"
    Else
        If Not rsRegistroFolio.BOF Then
            rsRegistroFolio.MovePrevious
            If rsRegistroFolio.BOF Then
                rsRegistroFolio.MoveNext
            End If
        Else
            rsRegistroFolio.MoveNext
        End If
        pMuestraFolio
        pHabilita 1, 1, 1, 1, 1, 0, IIf((rsRegistroFolio!SMIDEPARTAMENTO <> vgintNumeroDepartamento And cgstrModulo <> "SI") Or rsRegistroFolio!intNumeroActual > rsRegistroFolio!intNumeroInicial, 0, 1)
        pMuestraCancelar
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdBack_Click"))
End Sub

Private Sub cmdCancelar_Click()
    On Error GoTo NotificaError
    Dim vllngPersonaGraba As Long
    Dim strParametros As String
    Dim lnCveRegistro As Long
    
    If Val(txtFinalCancelar.Text) < Val(txtInicialCancelar.Text) Then
        '¡Rango invalido!
        MsgBox SIHOMsg(26), vbOKOnly + vbInformation, "Mensaje"
        txtFinalCancelar.Text = rsRegistroFolio!intNumeroFinal
        txtFinalCancelar.SetFocus
    Else
        If Val(txtFinalCancelar.Text) > rsRegistroFolio!intNumeroFinal Then
            '¡Rango invalido!
            MsgBox SIHOMsg(26), vbOKOnly + vbInformation, "Mensaje"
            txtFinalCancelar.Text = rsRegistroFolio!intNumeroFinal
            txtFinalCancelar.SetFocus
        Else
            '¿Desea guardar los datos?
            If MsgBox(SIHOMsg(4), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
                '--------------------------------------------------------
                ' Persona que graba
                '--------------------------------------------------------
                vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
                If vllngPersonaGraba = 0 Then Exit Sub
                EntornoSIHO.ConeccionSIHO.BeginTrans
                    With rsFolioCancelado

                         strParametros = CStr(rsRegistroFolio!intcveregistro) & "|" & CStr(rsRegistroFolio!chrTipoDocumento) & "|" & CStr(Trim(rsRegistroFolio!chrCveDocumento)) _
                        & "|" & CStr(vgintNumeroDepartamento) & "|" & CStr(fdtmServerFecha) & "|" & Trim(txtMotivo.Text) _
                        & "|" & txtInicialCancelar.Text & "|" & txtFinalCancelar.Text & "|" & CStr(vglngNumeroEmpleado)
                    
                    frsEjecuta_SP strParametros, "SP_PVINSFOLIOCANCELADO"
                                                           
                    lnCveRegistro = rsRegistroFolio!intcveregistro
                    rsRegistroFolio.Requery
                    rsRegistroFolio.Find ("intCveRegistro=" & lnCveRegistro)
                    
                    End With
                    
                Call pGuardarLogTransaccion(Me.Name, EnmCancelacion, vllngPersonaGraba, "FOLIO", CStr(rsRegistroFolio!intcveregistro))
                EntornoSIHO.ConeccionSIHO.CommitTrans
                pMuestraFolio
                pHabilita 1, 1, 1, 1, 1, 0, IIf(rsRegistroFolio!SMIDEPARTAMENTO <> vgintNumeroDepartamento And cgstrModulo <> "SI", 0, IIf(rsRegistroFolio!intNumeroActual > rsRegistroFolio!intNumeroFinal, 0, 1))
                pMuestraCancelar
                SSTabFolios.Tab = 0
                cmdLocate.SetFocus
            End If
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdCancelar_Click"))
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo NotificaError
    Dim rsFoliosCancelados As New ADODB.Recordset
    Dim vllngPersonaGraba As Long
    
    '--------------------------------------------------------
    ' Persona que graba
    '--------------------------------------------------------
    vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
    If vllngPersonaGraba = 0 Then Exit Sub
    vlstrSentencia = "select count(*) as Total from FolioCancelado inner join Nodepartamento on FolioCancelado.smidepartamento = NoDepartamento.SMICVEDEPARTAMENTO and NoDepartamento.TnyClaveEmpresa = " & vgintClaveEmpresaContable & " where intCveRegistro=" + Str(rsRegistroFolio!intcveregistro)
    Set rsFoliosCancelados = frsRegresaRs(vlstrSentencia)
    If rsFoliosCancelados!Total = 0 Then
        Call pGuardarLogTransaccion(Me.Name, EnmBorrar, vllngPersonaGraba, "FOLIO", CStr(rsRegistroFolio!intcveregistro))
        rsRegistroFolio.Delete
        rsRegistroFolio.Update
        rsRegistroFolio.Requery
        optTipoFactura.SetFocus
    Else
        MsgBox SIHOMsg(257), vbOKOnly + vbCritical, "Mensaje"
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdDelete_Click"))
End Sub

Private Sub cmdEnd_Click()
    On Error GoTo NotificaError
    
    If rsRegistroFolio.RecordCount = 0 Then
        ' No existe información.
        MsgBox SIHOMsg(13), vbOKOnly + vbInformation, "Mensaje"
    Else
        rsRegistroFolio.MoveLast
        pMuestraFolio
        pHabilita 1, 1, 1, 1, 1, 0, IIf((rsRegistroFolio!SMIDEPARTAMENTO <> vgintNumeroDepartamento And cgstrModulo <> "SI") Or rsRegistroFolio!intNumeroActual > rsRegistroFolio!intNumeroInicial, 0, 1)
        pMuestraCancelar
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdEnd_Click"))
End Sub

Private Sub cmdLocate_Click()
    On Error GoTo NotificaError
    
    If rsRegistroFolio.RecordCount = 0 Then
        ' No existe información.
        MsgBox SIHOMsg(13), vbOKOnly + vbInformation, "Mensaje"
    Else
        SSTabFolios.Tab = 1
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdLocate_Click"))
End Sub

Private Sub cmdNext_Click()
    On Error GoTo NotificaError
    
    If rsRegistroFolio.RecordCount = 0 Then
        ' No existe información.
        MsgBox SIHOMsg(13), vbOKOnly + vbInformation, "Mensaje"
    Else
        If Not rsRegistroFolio.EOF Then
            rsRegistroFolio.MoveNext
            If rsRegistroFolio.EOF Then
                rsRegistroFolio.MovePrevious
            End If
        Else
            rsRegistroFolio.MovePrevious
        End If
        pMuestraFolio
        pHabilita 1, 1, 1, 1, 1, 0, IIf((rsRegistroFolio!SMIDEPARTAMENTO <> vgintNumeroDepartamento And cgstrModulo <> "SI") Or rsRegistroFolio!intNumeroActual > rsRegistroFolio!intNumeroInicial, 0, 1)
        pMuestraCancelar
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdNext_Click"))
End Sub

Private Sub cmdSave_Click()
    On Error GoTo NotificaError
    Dim vllngPersonaGraba As Long                'Persona que esta generando la factura
    Dim vllngSecuencia As Long
    Dim strTipoDocumento As String
    Dim intTipo As Integer
    Dim intNumeroActual As Long
    Dim strParametros As String
    Dim strdtmFechaAprobacion As String
    Dim strvchNumAprobacion As String
    Dim intCFDi As Integer
    
    If fblnDatosValidos() Then
        
        ' Persona que graba
        vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
        
        If vllngPersonaGraba <> 0 Then
        
            With rsRegistroFolio
                If optTipoFactura.Value Then
                    strTipoDocumento = "FA"
                ElseIf optTipoRecibo.Value Then
                    strTipoDocumento = "RE"
                ElseIf optTipoNotaCredito.Value Then
                    strTipoDocumento = "NC"
                ElseIf optTipoNotaCargo.Value Then
                    strTipoDocumento = "NA"
                ElseIf optTipoTicket.Value Then
                    strTipoDocumento = "TI"
                ElseIf optTipoNotaCreditoCargo.Value Then
                    strTipoDocumento = "CC"
                ElseIf optTipoDonativos.Value Then
                    strTipoDocumento = "DO"
                ElseIf optTipoAplAnt.Value Then
                    strTipoDocumento = "AA"
                Else
                    strTipoDocumento = "SD"
                End If
                
                If chkComprobanteFiscal.Value = vbChecked Then
                    intTipo = 1
                    If Trim(mskFechaAprobacion.Text) = "" Then
                        strdtmFechaAprobacion = ""
                    Else
                        strdtmFechaAprobacion = Trim(CDate(mskFechaAprobacion.Text))
                    End If
                    strvchNumAprobacion = Trim(txtNumAprobacion.Text)
                    intCFDi = IIf(optCFD(0).Value, 0, IIf(optCFD(1), 1, 2))
                Else
                    intTipo = 0
                    intCFDi = 2
                End If

                If txtFolioInicial.Enabled = False Then
                    intNumeroActual = !intNumeroActual
                Else
                    intNumeroActual = Val(txtFolioInicial.Text)
                End If
                     
                EntornoSIHO.ConeccionSIHO.BeginTrans
                
                If Not vlblnConsulta Then
                    
                    strParametros = "" & "|" & strTipoDocumento & "|" & IIf(Trim(txtIdentificador.Text) = "", " ", Trim(txtIdentificador.Text)) & "|" & cboDepartamento.ItemData(cboDepartamento.ListIndex) _
                    & "|" & CStr(vglngNumeroEmpleado) & "|" & CStr(fdtmServerFecha) & "|" & CStr(Val(txtFolioInicial.Text)) _
                    & "|" & CStr(Val(txtFolioFinal.Text)) & "|" & CStr(intNumeroActual) & "|" & Val(txtMensaje.Text) _
                    & "|" & strvchNumAprobacion & "|" & strdtmFechaAprobacion & "|" & CStr(intTipo) & "|" & "I" & "|" & intCFDi
                    
                    frsEjecuta_SP strParametros, "SP_PVUPDINSFOLIOS"
                
                Else
                
                    strParametros = CStr(!intcveregistro) & "|" & strTipoDocumento & "|" & Trim(txtIdentificador.Text) & "|" & cboDepartamento.ItemData(cboDepartamento.ListIndex) _
                    & "|" & CStr(vglngNumeroEmpleado) & "|" & CStr(!dtmFechaRegistro) & "|" & CStr(Val(txtFolioInicial.Text)) _
                    & "|" & CStr(Val(txtFolioFinal.Text)) & "|" & CStr(intNumeroActual) & "|" & Val(txtMensaje.Text) _
                    & "|" & strvchNumAprobacion & "|" & strdtmFechaAprobacion & "|" & CStr(intTipo) & "|" & "A" & "|" & intCFDi
                    
                    frsEjecuta_SP strParametros, "SP_PVUPDINSFOLIOS"
                
                End If
                
                         
                If Not vlblnConsulta Then ' si es un registro que se agrega
                  vllngSecuencia = flngObtieneIdentity("SEC_RegistroFolio", vllngSecuencia)
                Else ' si es un registro que se modifico
                  vllngSecuencia = !intcveregistro
                End If
                .Requery
                .Find ("intCveRegistro=" & vllngSecuencia)
                
                EntornoSIHO.ConeccionSIHO.CommitTrans
                
                If Not vlblnConsulta Then
                    Call pGuardarLogTransaccion(Me.Name, EnmGrabar, vllngPersonaGraba, "FOLIO", CStr(vllngSecuencia))
                Else
                    Call pGuardarLogTransaccion(Me.Name, EnmCambiar, vllngPersonaGraba, "FOLIO", CStr(vllngSecuencia))
                End If
            End With
            optTipoFactura.SetFocus
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdSave_Click"))
End Sub

Private Function fblnDatosValidos() As Boolean
    On Error GoTo NotificaError
    Dim rsFoliosActivos As New ADODB.Recordset
    Dim rsFoliosMismoIdentificador As New ADODB.Recordset
    Dim vlstrTipoDocumento As String
    Dim vlstrIndicador As String
    
    'Validar el campo del indicador
    If txtIdentificador.Text = "" Then
        vlstrIndicador = "        "
    Else
        vlstrIndicador = txtIdentificador.Text
    End If
    
    If optTipoFactura.Value Then
        vlstrTipoDocumento = "FA"
    ElseIf optTipoRecibo.Value Then
        vlstrTipoDocumento = "RE"
    ElseIf optTipoNotaCredito.Value Then
        vlstrTipoDocumento = "NC"
    ElseIf optTipoNotaCargo.Value Then
        vlstrTipoDocumento = "NA"
    ElseIf optTipoTicket.Value Then
        vlstrTipoDocumento = "TI"
    ElseIf optTipoNotaCreditoCargo.Value Then
        vlstrTipoDocumento = "CC"
    ElseIf optTipoDonativos.Value Then
        vlstrTipoDocumento = "DO"
    ElseIf optTipoAplAnt.Value Then
        vlstrTipoDocumento = "AA"
    Else
        vlstrTipoDocumento = "SD"
    End If

    fblnDatosValidos = True
    
    If fblnDatosValidos And chkComprobanteFiscal.Value = vbUnchecked And optTipoDonativos.Value = True Then
        fblnDatosValidos = False
        'Los folios para donativo solo pueden ser de tipo digital
        MsgBox SIHOMsg(1088), vbOKOnly + vbExclamation, "Mensaje"
        cmdSave.SetFocus
    End If
    
    If fblnDatosValidos And chkComprobanteFiscal.Value = vbChecked And optCFD(0).Value = True And Trim(txtNumAprobacion.Text) = "" Then
        fblnDatosValidos = False
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbOKOnly + vbInformation, "Mensaje"
        txtNumAprobacion.SetFocus
    End If
        
    If fblnDatosValidos And chkComprobanteFiscal.Value = vbChecked And optCFD(0).Value = True And Trim(mskFechaAprobacion.ClipText) = "" Then
        fblnDatosValidos = False
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbOKOnly + vbInformation, "Mensaje"
        mskFechaAprobacion.SetFocus
    End If
    
    If fblnDatosValidos And (Trim(txtFolioFinal.Text) > 2147483647) Then
        fblnDatosValidos = False
        'Por favor, introduzca un valor menor a 2147483647
        MsgBox SIHOMsg(1011), vbOKOnly + vbInformation, "Mensaje"
        txtFolioFinal.SetFocus
    End If
    
    If fblnDatosValidos And chkComprobanteFiscal.Value = vbChecked And optCFD(0).Value = True And Trim(mskFechaAprobacion.ClipText) <> "" Then
        If Not IsDate(mskFechaAprobacion.Text) Then
            fblnDatosValidos = False
            '¡Fecha no válida!, formato de fecha dd/mm/aaaa
            MsgBox SIHOMsg(29), vbOKOnly + vbInformation, "Mensaje"
            mskFechaAprobacion.SetFocus
        End If
    End If
    
    If Trim(txtFolioInicial.Text) = "" Then
        fblnDatosValidos = False
        MsgBox SIHOMsg(2), vbOKOnly + vbInformation, "Mensaje"
        If txtFolioInicial.Enabled And txtFolioInicial.Visible Then
          txtFolioInicial.SetFocus
        Else
          If txtFolioFinal.Enabled And txtFolioFinal.Visible Then txtFolioFinal.SetFocus
        End If
    End If
    If fblnDatosValidos And Trim(txtFolioFinal.Text) = "" Then
        fblnDatosValidos = False
        MsgBox SIHOMsg(2), vbOKOnly + vbInformation, "Mensaje"
        txtFolioFinal.SetFocus
    End If
    If fblnDatosValidos And Trim(txtMensaje.Text) = "" Then
        fblnDatosValidos = False
        MsgBox SIHOMsg(2), vbOKOnly + vbInformation, "Mensaje"
        txtMensaje.SetFocus
    End If
    If fblnDatosValidos And Val(txtFolioInicial.Text) <= 0 Then
        fblnDatosValidos = False
        'Dato incorrecto: El valor debe ser
        MsgBox SIHOMsg(36) & " mayor a cero", vbOKOnly + vbInformation, "Mensaje"
        If txtFolioInicial.Enabled And txtFolioInicial.Visible Then
          txtFolioInicial.SetFocus
        Else
          If txtFolioFinal.Enabled And txtFolioFinal.Visible Then txtFolioFinal.SetFocus
        End If
    End If
    If fblnDatosValidos And Val(txtFolioFinal.Text) < Val(txtFolioInicial.Text) Then
        fblnDatosValidos = False
        'El folio inicial debe ser menor o igual al final!
        MsgBox SIHOMsg(201), vbOKOnly + vbInformation, "Mensaje"
        If txtFolioInicial.Enabled And txtFolioInicial.Visible Then
          txtFolioInicial.SetFocus
        Else
          If txtFolioFinal.Enabled And txtFolioFinal.Visible Then txtFolioFinal.SetFocus
        End If
    End If
    If fblnDatosValidos And Val(txtFolioFinal.Text) < Val(txtFolioActual.Text) Then
        fblnDatosValidos = False
        'Dato incorrecto: El valor debe ser
        MsgBox SIHOMsg(36) & " mayor o igual al folio actual", vbOKOnly + vbInformation, "Mensaje"
        txtFolioFinal.SetFocus
    End If
    If fblnDatosValidos And (Val(txtFolioFinal.Text) - Val(txtFolioInicial.Text) + 1) < Val(txtMensaje.Text) Then
        fblnDatosValidos = False
        'El número de folios para aviso está incorrecto.
        MsgBox SIHOMsg(285), vbOKOnly + vbInformation, "Mensaje"
        txtMensaje.SetFocus
    End If
    If fblnDatosValidos Then
        If Not vlblnConsulta Then
            vlstrSentencia = "" & _
            "select " & _
                "count(*) as Total " & _
            "from " & _
                "RegistroFolio inner join Nodepartamento on RegistroFolio.smidepartamento = NoDepartamento.SMICVEDEPARTAMENTO and NoDepartamento.TnyClaveEmpresa = " & vgintClaveEmpresaContable & _
            " where " & _
                "intNumeroFinal >= intNumeroActual " & _
                "and smiDepartamento = " & Str(cboDepartamento.ItemData(cboDepartamento.ListIndex)) & " " & _
                "and chrTipoDocumento = " & " '" & vlstrTipoDocumento & "'"
            Set rsFoliosActivos = frsRegresaRs(vlstrSentencia)
            If rsFoliosActivos!Total <> 0 Then
                fblnDatosValidos = False
                'Aún existen folios activos de este tipo de documento para este departamento.
                MsgBox SIHOMsg(286), vbOKOnly + vbInformation, "Mensaje"
                cmdSave.SetFocus
            End If
        End If
    End If
    
    If fblnDatosValidos Then
        vlstrSentencia = "select count(*) as Total " & _
                           "from RegistroFolio " & _
                     "inner join Nodepartamento " & _
                             "on RegistroFolio.smidepartamento = NoDepartamento.SMICVEDEPARTAMENTO and NoDepartamento.TnyClaveEmpresa = " & vgintClaveEmpresaContable & _
                         " where chrCveDocumento=" + "'" + vlstrIndicador + "' " & _
                            "and chrTipoDocumento = " & " '" & vlstrTipoDocumento & "' " & _
                            "and (" & _
                                "(intNumeroInicial >= " & Val(txtFolioInicial.Text) & " and intNumeroInicial <=" & Val(txtFolioFinal.Text) & " ) or " & _
                                "(intnumeroInicial <= " & Val(txtFolioInicial.Text) & " and intNumeroFinal >=" & Val(txtFolioInicial.Text) & " ) or " & _
                                "(intnumeroInicial >= " & Val(txtFolioInicial.Text) & " and intNumeroFinal <=" & Val(txtFolioFinal.Text) & " ))"
        If vlblnConsulta Then
           vlstrSentencia = vlstrSentencia & " and intcveregistro <> " & rsRegistroFolio!intcveregistro
        End If
        
        Set rsFoliosMismoIdentificador = frsRegresaRs(vlstrSentencia)
        If rsFoliosMismoIdentificador!Total <> 0 Then
            fblnDatosValidos = False
            '!Existe duplicidad en los folios!
            MsgBox SIHOMsg(287), vbOKOnly + vbInformation, "Mensaje"
            cmdSave.SetFocus
        End If
    End If
    
    If fblnDatosValidos And txtFolioFinal.Enabled Then
        If Len(Trim(txtFolioFinal.Text) & Trim(txtIdentificador.Text)) > 12 Then
            fblnDatosValidos = False
            'El folio final e identificador no deben exceder de 12 caracteres.
            MsgBox SIHOMsg(1511), vbOKOnly + vbInformation, "Mensaje"
            txtFolioFinal.SetFocus
        End If
    End If
       
'    If fblnDatosValidos Then
'        vlstrSentencia = "select count(*) as Total from RegistroFolio inner join Nodepartamento on RegistroFolio.smidepartamento = NoDepartamento.SMICVEDEPARTAMENTO and NoDepartamento.TnyClaveEmpresa = " & vgintClaveEmpresaContable
'        vlstrSentencia = vlstrSentencia + " where chrCveDocumento=" + "'" + vlstrIndicador + "'" + " "
'        vlstrSentencia = vlstrSentencia + "and chrTipoDocumento = " & " '" & vlstrTipoDocumento & "'" & " "
'        vlstrSentencia = vlstrSentencia + "and smiDepartamento<>" + Str(cboDepartamento.ItemData(cboDepartamento.ListIndex)) + " "
'        vlstrSentencia = vlstrSentencia + "and intNumeroInicial<=" + txtFolioInicial.Text + " "
'        vlstrSentencia = vlstrSentencia + "and intNumeroFinal>=" + txtFolioInicial.Text
'        Set rsFoliosMismoIdentificador = frsRegresaRs(vlstrSentencia)
'        If rsFoliosMismoIdentificador!Total <> 0 Then
'            fblnDatosValidos = False
'            '!Existe duplicidad en los folios!
'            MsgBox SIHOMsg(287), vbOKOnly + vbInformation, "Mensaje"
'            cmdSave.SetFocus
'        End If
'    End If
'    If fblnDatosValidos Then
'        vlstrSentencia = "select count(*) as Total from RegistroFolio inner join Nodepartamento on RegistroFolio.smidepartamento = NoDepartamento.SMICVEDEPARTAMENTO and NoDepartamento.TnyClaveEmpresa = " & vgintClaveEmpresaContable
'        vlstrSentencia = vlstrSentencia + " where chrCveDocumento=" + "'" + vlstrIndicador + "'" + " "
'        vlstrSentencia = vlstrSentencia + "and chrTipoDocumento = " + " '" + vlstrTipoDocumento + "'"
'        vlstrSentencia = vlstrSentencia + "and smiDepartamento<>" + Str(cboDepartamento.ItemData(cboDepartamento.ListIndex)) + " "
'        vlstrSentencia = vlstrSentencia + "and intNumeroInicial<=" + txtFolioFinal.Text + " "
'        vlstrSentencia = vlstrSentencia + "and intNumeroFinal>=" + txtFolioFinal.Text
'        Set rsFoliosMismoIdentificador = frsRegresaRs(vlstrSentencia)
'        If rsFoliosMismoIdentificador!Total <> 0 Then
'            fblnDatosValidos = False
'            '!Existe duplicidad en los folios!
'            MsgBox SIHOMsg(287), vbOKOnly + vbInformation, "Mensaje"
'            cmdSave.SetFocus
'        End If
'    End If
'    If fblnDatosValidos Then
'        If Not vlblnConsulta Then
'            If fblnDatosValidos Then
'                vlstrSentencia = "select count(*) as Total from RegistroFolio inner join Nodepartamento on RegistroFolio.smidepartamento = NoDepartamento.SMICVEDEPARTAMENTO and NoDepartamento.TnyClaveEmpresa = " & vgintClaveEmpresaContable
'                vlstrSentencia = vlstrSentencia + " where chrCveDocumento=" + "'" + vlstrIndicador + "'" + " "
'                vlstrSentencia = vlstrSentencia + "and chrTipoDocumento = " + " '" + vlstrTipoDocumento + "'" + " "
'                vlstrSentencia = vlstrSentencia + "and smiDepartamento=" + Str(cboDepartamento.ItemData(cboDepartamento.ListIndex)) + " "
'                vlstrSentencia = vlstrSentencia + "and intNumeroInicial<=" + txtFolioInicial.Text + " "
'                vlstrSentencia = vlstrSentencia + "and intNumeroFinal>=" + txtFolioInicial.Text
'                Set rsFoliosMismoIdentificador = frsRegresaRs(vlstrSentencia)
'                If rsFoliosMismoIdentificador!Total <> 0 Then
'                    fblnDatosValidos = False
'                    '!Existe duplicidad en los folios!
'                    MsgBox SIHOMsg(287), vbOKOnly + vbInformation, "Mensaje"
'                    cmdSave.SetFocus
'                End If
'            End If
'            If fblnDatosValidos Then
'                vlstrSentencia = "select count(*) as Total from RegistroFolio inner join Nodepartamento on RegistroFolio.smidepartamento = NoDepartamento.SMICVEDEPARTAMENTO and NoDepartamento.TnyClaveEmpresa = " & vgintClaveEmpresaContable
'                vlstrSentencia = vlstrSentencia + " where chrCveDocumento=" + "'" + vlstrIndicador + "'" + " "
'                vlstrSentencia = vlstrSentencia + "and chrTipoDocumento = " + " '" + vlstrTipoDocumento + "'"
'                vlstrSentencia = vlstrSentencia + "and smiDepartamento=" + Str(cboDepartamento.ItemData(cboDepartamento.ListIndex)) + " "
'                vlstrSentencia = vlstrSentencia + "and intNumeroInicial<=" + txtFolioFinal.Text + " "
'                vlstrSentencia = vlstrSentencia + "and intNumeroFinal>=" + txtFolioFinal.Text
'                Set rsFoliosMismoIdentificador = frsRegresaRs(vlstrSentencia)
'                If rsFoliosMismoIdentificador!Total <> 0 Then
'                    fblnDatosValidos = False
'                    '!Existe duplicidad en los folios!
'                    MsgBox SIHOMsg(287), vbOKOnly + vbInformation, "Mensaje"
'                    cmdSave.SetFocus
'                End If
'            End If
'        End If
'    End If

Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fblnDatosValidos"))
End Function

Private Sub cmdTop_Click()
    On Error GoTo NotificaError
    
    If rsRegistroFolio.RecordCount = 0 Then
        ' No existe información.
        MsgBox SIHOMsg(13), vbOKOnly + vbInformation, "Mensaje"
    Else
        rsRegistroFolio.MoveFirst
        pMuestraFolio
        pHabilita 1, 1, 1, 1, 1, 0, IIf((rsRegistroFolio!SMIDEPARTAMENTO <> vgintNumeroDepartamento And cgstrModulo <> "SI") Or rsRegistroFolio!intNumeroActual > rsRegistroFolio!intNumeroInicial, 0, 1)
        pMuestraCancelar
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdTop_Click"))
End Sub

Private Sub pMuestraCancelar()
    On Error GoTo NotificaError
    
    If rsRegistroFolio!SMIDEPARTAMENTO = cboDepartamento.ItemData(cboDepartamento.ListIndex) Then
        If rsRegistroFolio!intNumeroActual <= rsRegistroFolio!intNumeroFinal Then
            SSTabFolios.TabEnabled(2) = True
            txtInicialCancelar.Text = rsRegistroFolio!intNumeroActual
            txtFinalCancelar.Text = rsRegistroFolio!intNumeroFinal
            txtMotivo.Text = ""
        Else
            SSTabFolios.TabEnabled(2) = False
        End If
    Else
        SSTabFolios.TabEnabled(2) = False
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pMuestraCancelar"))
End Sub

Private Sub Form_Activate()
    On Error GoTo NotificaError
    
    If rsDepartamentos.RecordCount = 0 Then
        'No se encontró información del departamento.
        MsgBox SIHOMsg(233), vbOKOnly + vbInformation, "Mensaje"
        Unload Me
    End If
    
    If rsFolioUnicoNotas.RecordCount <> 0 Then
        If rsFolioUnicoNotas!folios <> 0 Then
            optTipoNotaCargo.Enabled = False
            optTipoNotaCredito.Enabled = False
        Else
            optTipoNotaCreditoCargo.Enabled = False
        End If
    Else
        'No se ha definido si se manejará un folio único para notas de cargo y crédito en los parámetros del módulo de cuentas por cobrar
        MsgBox SIHOMsg(1257), vbOKOnly + vbInformation, "Mensaje"
        Unload Me
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_Activate"))
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 27 Then
        If SSTabFolios.Tab = 0 Then
            If cmdSave.Enabled Then
                '¿Desea abandonar la operación?
                If MsgBox(SIHOMsg(17), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
                    optTipoFactura.SetFocus
                End If
            Else
                Unload Me
            End If
        Else
            KeyAscii = 0
            SSTabFolios.Tab = 0
            optTipoFactura.SetFocus
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_KeyPress"))
End Sub

Private Sub Form_Load()
    On Error GoTo NotificaError
    
    Me.Icon = frmMenuPrincipal.Icon
    '-----------------------------------------------
    ' Recordsets tipo tabla RegistroFolio, FolioCancelado
    '-----------------------------------------------
    vlstrSentencia = "select * From RegistroFolio where RegistroFolio.smidepartamento in (select NoDepartamento.SMICVEDEPARTAMENTO from Nodepartamento where NoDepartamento.TnyClaveEmpresa = " & vgintClaveEmpresaContable & " )"
    Set rsRegistroFolio = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
    vlstrSentencia = "select * From FolioCancelado where FolioCancelado.smidepartamento in (select NoDepartamento.SMICVEDEPARTAMENTO from Nodepartamento where NoDepartamento.TnyClaveEmpresa = " & vgintClaveEmpresaContable & " )"
    Set rsFolioCancelado = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
    
    vlstrSentencia = "select * from Nodepartamento where bitestatus = 1 and tnyclaveempresa = " & vgintClaveEmpresaContable
    If cgstrModulo <> "SI" Then
        cboDepartamento.Enabled = False
    End If
    
    Set rsFolioUnicoNotas = frsRegresaRs(" Select intFolioUnicoNotas folios from CcParametro ", adLockOptimistic, adOpenDynamic)
    If rsFolioUnicoNotas.RecordCount <> 0 Then
        If rsFolioUnicoNotas!folios <> 0 Then
            optTipoNotaCargo.Enabled = False
            optTipoNotaCredito.Enabled = False
        Else
            optTipoNotaCreditoCargo.Enabled = False
        End If
    End If
    
    Set rsDepartamentos = frsRegresaRs(vlstrSentencia)
    If rsDepartamentos.RecordCount > 0 Then
        Call pLlenarCboRs(cboDepartamento, rsDepartamentos, 0, 1)
        If cgstrModulo <> "SI" Then
            cboDepartamento.ListIndex = flngLocalizaCbo(cboDepartamento, CStr(vgintNumeroDepartamento))
        End If
        SSTabFolios.Tab = 0
    End If

    optTipoFactura.Value = True

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_Load"))
End Sub

Private Sub grdFolios_DblClick()
    On Error GoTo NotificaError
    
    fintLocalizaPkRs rsRegistroFolio, 0, Str(grdFolios.RowData(grdFolios.Row))
    pMuestraFolio
    pHabilita 1, 1, 1, 1, 1, 0, IIf((rsRegistroFolio!SMIDEPARTAMENTO <> vgintNumeroDepartamento And cgstrModulo <> "SI") Or rsRegistroFolio!intNumeroActual > rsRegistroFolio!intNumeroInicial, 0, 1)
    pMuestraCancelar
    SSTabFolios.Tab = 0
    cmdLocate.SetFocus

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":grdFolios_DblClick"))
End Sub

Private Sub grdFolios_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    If KeyCode = 13 Then
        fintLocalizaPkRs rsRegistroFolio, 0, Str(grdFolios.RowData(grdFolios.Row))
        pMuestraFolio
        pHabilita 1, 1, 1, 1, 1, 0, IIf((rsRegistroFolio!SMIDEPARTAMENTO <> vgintNumeroDepartamento And cgstrModulo <> "SI") Or rsRegistroFolio!intNumeroActual > rsRegistroFolio!intNumeroInicial, 0, 1)
        pMuestraCancelar
        SSTabFolios.Tab = 0
        cmdLocate.SetFocus
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":grdFolios_KeyDown"))
End Sub

Private Sub mskFechaAprobacion_GotFocus()
On Error GoTo NotificaError

    If txtNumAprobacion.Locked = True Then
    
        If txtIdentificador.Enabled = True Then txtIdentificador.SetFocus
        
    Else
        mskFechaAprobacion.Mask = "##/##/####"
        pSelMkTexto mskFechaAprobacion
    End If
    
    If chkComprobanteFiscal.Value = vbChecked Then
        pHabilita 0, 0, 0, 0, 0, 1, 0
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFechaAprobacion_GotFocus"))
End Sub

Private Sub mskFechaAprobacion_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        If cmdSave.Enabled = True Then
            cmdSave.SetFocus
        Else
            txtIdentificador.SetFocus
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFechaAprobacion_KeyPress"))
End Sub

Private Sub optCFD_Click(Index As Integer)
        If optCFD(0).Value = True Then 'Se habilitan los campos si la opción CFD está activada
            Label12.Enabled = True
            Label13.Enabled = True
            txtNumAprobacion.Enabled = True
            mskFechaAprobacion.Enabled = True
            mskFechaAprobacion.Mask = "##/##/####"
        Else
            Label12.Enabled = False
            Label13.Enabled = False
            mskFechaAprobacion.Mask = ""
            mskFechaAprobacion.Text = ""
            txtNumAprobacion.Text = ""
            txtNumAprobacion.Enabled = False
            mskFechaAprobacion.Enabled = False
        End If
End Sub

Private Sub optCFD_GotFocus(Index As Integer)
    pHabilita 0, 0, 0, 0, 0, 1, 0
End Sub


Private Sub optCFD_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtIdentificador.SetFocus
    End If
End Sub

Private Sub optTipoAplAnt_Click()
    chkComprobanteFiscal.Enabled = True
    chkComprobanteFiscal.Value = vbChecked
    optCFD(1).Value = True
    fraCFD.Enabled = False
End Sub

Private Sub optTipoAplAnt_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
      pEnfocaTextBox txtIdentificador
    End If
End Sub

Private Sub optTipoDonativos_Click()
    fraCFD.Enabled = True
    If optTipoDonativos.Value = True Then
        chkComprobanteFiscal.Enabled = True
        txtNumAprobacion.Enabled = True
        mskFechaAprobacion.Enabled = True
        mskFechaAprobacion.Mask = "##/##/####"
    End If
    
End Sub

Private Sub optTipoDonativos_GotFocus()
    On Error GoTo NotificaError
    
    vlstrTipoActivo = "Donativo"
    pLimpia

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":optTipoDonativos_GotFocus"))
End Sub

Private Sub optTipoDonativos_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        optTipoDonativos.Value = True
        pHabilita 0, 0, 0, 0, 0, 1, 0
        chkComprobanteFiscal.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":optTipoRecibo_KeyPress"))
End Sub

Private Sub optTipoFactura_Click()
    fraCFD.Enabled = True
    If optTipoFactura.Value = True Then
        chkComprobanteFiscal.Enabled = True
        txtNumAprobacion.Enabled = True
        mskFechaAprobacion.Enabled = True
        mskFechaAprobacion.Mask = "##/##/####"
    End If
    
End Sub

Private Sub optTipoFactura_GotFocus()
    On Error GoTo NotificaError
    
    vlstrTipoActivo = "Factura"
    pLimpia

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":optTipoFactura_GotFocus"))
End Sub

Private Sub optTipoFactura_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        optTipoFactura.Value = True
        pHabilita 0, 0, 0, 0, 0, 1, 0
        chkComprobanteFiscal.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":optTipoFactura_KeyPress"))
End Sub

Private Sub optTipoNotaCargo_Click()
    fraCFD.Enabled = True
    If optTipoNotaCargo.Value = True Then
        chkComprobanteFiscal.Enabled = True
        txtNumAprobacion.Enabled = True
        mskFechaAprobacion.Enabled = True
        mskFechaAprobacion.Mask = "##/##/####"
    End If
    
End Sub

Private Sub optTipoNotaCargo_GotFocus()
    On Error GoTo NotificaError
    
    vlstrTipoActivo = "NCargo"
    pLimpia

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":optTipoNotaCargo_GotFocus"))
End Sub

Private Sub optTipoNotaCargo_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        optTipoNotaCargo.Value = True
        pHabilita 0, 0, 0, 0, 0, 1, 0
        chkComprobanteFiscal.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":optTipoNotaCargo_KeyPress"))
End Sub

Private Sub optTipoNotaCredito_Click()
    fraCFD.Enabled = True
    If optTipoNotaCredito.Value = True Then
        chkComprobanteFiscal.Enabled = True
        txtNumAprobacion.Enabled = True
        mskFechaAprobacion.Enabled = True
        mskFechaAprobacion.Mask = "##/##/####"
    End If
    
End Sub

Private Sub optTipoNotaCredito_GotFocus()
    On Error GoTo NotificaError
    
    vlstrTipoActivo = "NCredito"
    pLimpia

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":optTipoNotaCredito_GotFocus"))
End Sub

Private Sub optTipoNotaCredito_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        optTipoNotaCredito.Value = True
        pHabilita 0, 0, 0, 0, 0, 1, 0
        chkComprobanteFiscal.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":optTipoNotaCredito_KeyPress"))
End Sub

Private Sub optTipoNotaCreditoCargo_Click()
    fraCFD.Enabled = True
    If optTipoNotaCreditoCargo.Value = True Then
        chkComprobanteFiscal.Enabled = True
        txtNumAprobacion.Enabled = True
        mskFechaAprobacion.Enabled = True
        mskFechaAprobacion.Mask = "##/##/####"
    End If
    
End Sub

Private Sub optTipoNotaCreditoCargo_GotFocus()
    On Error GoTo NotificaError
    
    vlstrTipoActivo = "NCC"
    pLimpia
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":optTipoNotaCreditoCargo_GotFocus"))
End Sub

Private Sub optTipoNotaCreditoCargo_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        optTipoNotaCreditoCargo.Value = True
        pHabilita 0, 0, 0, 0, 0, 1, 0
        chkComprobanteFiscal.SetFocus
    
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":optTipoNotaCreditoCargo_KeyPress"))
End Sub

Private Sub optTipoRecibo_Click()
    fraCFD.Enabled = True
    If optTipoRecibo.Value = True Then
        chkComprobanteFiscal.Enabled = True
        optCFD(0).Enabled = True
        optCFD(1).Enabled = True
        optCFD(0).Value = False
        optCFD(1).Value = True
        txtNumAprobacion.Enabled = True
        mskFechaAprobacion.Enabled = True
        mskFechaAprobacion.Mask = "##/##/####"
    End If
    
End Sub

Private Sub optTipoRecibo_GotFocus()
    On Error GoTo NotificaError
    
    vlstrTipoActivo = "Recibo"
    pLimpia

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":optTipoRecibo_GotFocus"))
End Sub

Private Sub optTipoRecibo_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        optTipoRecibo.Value = True
        pHabilita 0, 0, 0, 0, 0, 1, 0
        chkComprobanteFiscal.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":optTipoRecibo_KeyPress"))
End Sub

Private Sub optTipoSalidaDinero_Click()
    fraCFD.Enabled = True
    If optTipoSalidaDinero.Value = True Then
        chkComprobanteFiscal.Enabled = False
        optCFD(0).Enabled = False
        optCFD(1).Enabled = False
        optCFD(0).Value = False
        optCFD(1).Value = False
        txtNumAprobacion.Enabled = False
        mskFechaAprobacion.Enabled = False
        mskFechaAprobacion.Mask = "##/##/####"
    End If
    
End Sub

Private Sub optTipoSalidaDinero_GotFocus()
    On Error GoTo NotificaError
    
    vlstrTipoActivo = "SDinero"
    pLimpia

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":optTipoSalidaDinero_GotFocus"))
End Sub

Private Sub optTipoSalidaDinero_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        optTipoSalidaDinero.Value = True
        pHabilita 0, 0, 0, 0, 0, 1, 0
        txtIdentificador.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":optTipoSalidaDinero_KeyPress"))

End Sub

Private Sub optTipoTicket_Click()
    fraCFD.Enabled = True
    If optTipoTicket.Value = True Then
        chkComprobanteFiscal.Enabled = False
        optCFD(0).Enabled = False
        optCFD(1).Enabled = False
        optCFD(0).Value = False
        optCFD(1).Value = False
        txtNumAprobacion.Enabled = False
        mskFechaAprobacion.Enabled = False
        mskFechaAprobacion.Mask = "##/##/####"
    End If
    
End Sub

Private Sub optTipoTicket_GotFocus()
    On Error GoTo NotificaError

    vlstrTipoActivo = "Ticket"
    pLimpia

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":optTipoTicket_GotFocus"))
End Sub

Private Sub optTipoTicket_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        optTipoTicket.Value = True
        pHabilita 0, 0, 0, 0, 0, 1, 0
        txtIdentificador.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":optTipoTicket_KeyPress"))
End Sub

Private Sub SSTabFolios_Click(PreviousTab As Integer)
    On Error GoTo NotificaError
    
    If SSTabFolios.Tab = 2 Then
        txtFinalCancelar.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":SSTabFolios_Click"))
End Sub

Private Sub txtFinalCancelar_GotFocus()
    On Error GoTo NotificaError
    
    pSelTextBox txtFinalCancelar

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtFinalCancelar_GotFocus"))
End Sub

Private Sub txtFinalCancelar_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        txtMotivo.SetFocus
    Else
        If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
            KeyAscii = 7
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtFinalCancelar_KeyPress"))
End Sub

Private Sub txtFolioActual_GotFocus()
    On Error GoTo NotificaError
    
    pHabilita 0, 0, 0, 0, 0, 1, 0
    pSelTextBox txtFolioActual

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtFolioActual_GotFocus"))
End Sub

Private Sub txtFolioFinal_GotFocus()
    On Error GoTo NotificaError
    
    pHabilita 0, 0, 0, 0, 0, 1, 0
    pSelTextBox txtFolioFinal

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtFolioFinal_GotFocus"))
End Sub

Private Sub txtFolioFinal_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        txtMensaje.SetFocus
    Else
        If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
            KeyAscii = 7
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtFolioFinal_KeyPress"))
End Sub

Private Sub txtFolioInicial_GotFocus()
    On Error GoTo NotificaError
    
    pHabilita 0, 0, 0, 0, 0, 1, 0
    pSelTextBox txtFolioInicial

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtFolioInicial_GotFocus"))
End Sub

Private Sub txtFolioInicial_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        txtFolioFinal.SetFocus
    Else
        If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
            KeyAscii = 7
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtFolioInicial_KeyPress"))
End Sub

Private Sub txtFolioInicial_LostFocus()
    On Error GoTo NotificaError
    
    txtFolioActual.Text = txtFolioInicial.Text

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtFolioInicial_LostFocus"))
End Sub

Private Sub txtIdentificador_GotFocus()
    On Error GoTo NotificaError
    
    pHabilita 0, 0, 0, 0, 0, 1, 0

    pSelTextBox txtIdentificador

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtIdentificador_GotFocus"))
End Sub

Private Sub txtIdentificador_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        txtFolioInicial.SetFocus
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtIdentificador_KeyPress"))
End Sub

Private Sub txtMensaje_GotFocus()
    On Error GoTo NotificaError
    
    pHabilita 0, 0, 0, 0, 0, 1, 0
    pSelTextBox txtMensaje

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtMensaje_GotFocus"))
End Sub

Private Sub txtMensaje_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        If chkComprobanteFiscal.Value = vbChecked And chkComprobanteFiscal.Value = vbChecked Then
            If txtNumAprobacion.Enabled = True Then
                txtNumAprobacion.SetFocus
            Else
                cmdSave.SetFocus
            End If
        Else
            cmdSave.SetFocus
        End If
    Else
        If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
            KeyAscii = 7
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtMensaje_KeyPress"))
End Sub

Private Sub txtMotivo_GotFocus()
    On Error GoTo NotificaError
    
    pSelTextBox txtMotivo

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtMotivo_GotFocus"))
End Sub

Private Sub txtMotivo_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        cmdCancelar.SetFocus
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtMotivo_KeyPress"))
End Sub

Private Sub txtNumAprobacion_GotFocus()

    If chkComprobanteFiscal.Value = vbChecked Then
        pHabilita 0, 0, 0, 0, 0, 1, 0
    End If
    
    pSelTextBox txtNumAprobacion
        
    If chkComprobanteFiscal.Value = vbChecked Then
        txtNumAprobacion.Locked = False
    End If
    
End Sub

Private Sub txtNumAprobacion_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        mskFechaAprobacion.SetFocus
    Else
        If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
            KeyAscii = 7
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtNumAprobacion_KeyPress"))
End Sub

