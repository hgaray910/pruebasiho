VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmBusquedaPacientes 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Búsqueda de pacientes"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10920
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   10920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame freEstadoPaciente 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Estados del paciente"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   6360
      TabIndex        =   10
      Top             =   0
      Width           =   4440
      Begin VB.OptionButton optSoloActivos 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Solo activos"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   180
         TabIndex        =   13
         Top             =   280
         Width           =   1500
      End
      Begin VB.OptionButton optSinFacturar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Sin facturar"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1800
         TabIndex        =   15
         Top             =   280
         Width           =   1380
      End
      Begin VB.OptionButton optTodos 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3360
         TabIndex        =   16
         Top             =   280
         Width           =   855
      End
   End
   Begin VB.Frame freClavePaciente 
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
      Height          =   680
      Left            =   0
      TabIndex        =   8
      Top             =   -150
      Width           =   4710
      Begin VB.TextBox txtClavePaciente 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   0
         Top             =   230
         Width           =   1815
      End
      Begin VB.Label lblMoviClave 
         BackColor       =   &H80000005&
         Caption         =   "Movimiento del paciente"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   9
         Top             =   290
         Width           =   2895
      End
   End
   Begin VB.Frame frmOpcion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Tipo de búsqueda"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   75
      TabIndex        =   11
      Top             =   0
      Width           =   6210
      Begin VB.OptionButton optTipo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   2
         Left            =   2520
         TabIndex        =   3
         ToolTipText     =   "Búsqueda por nombre"
         Top             =   280
         Width           =   1090
      End
      Begin VB.OptionButton optTipo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Paterno"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   0
         Left            =   165
         TabIndex        =   1
         ToolTipText     =   "Búsqueda por apellido paterno"
         Top             =   280
         Width           =   1095
      End
      Begin VB.OptionButton optTipo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Materno"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   1
         Left            =   1320
         TabIndex        =   2
         ToolTipText     =   "Búsqueda por apellido materno"
         Top             =   280
         Width           =   1110
      End
      Begin VB.OptionButton optTipo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "CURP"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   3
         Left            =   3720
         TabIndex        =   4
         ToolTipText     =   "Búsqueda por CURP"
         Top             =   280
         Width           =   820
      End
      Begin VB.OptionButton optTipo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Expediente"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   4
         Left            =   4680
         TabIndex        =   5
         ToolTipText     =   "Búsqueda por expediente electrónico"
         Top             =   280
         Width           =   1380
      End
   End
   Begin VB.Timer tmrCarga 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   4800
      Top             =   3120
   End
   Begin VB.Frame freBusqueda 
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
      Height          =   5970
      Left            =   75
      TabIndex        =   12
      Top             =   550
      Width           =   10725
      Begin VB.TextBox txtBusqueda 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2925
         TabIndex        =   6
         ToolTipText     =   "Criterio de búsqueda"
         Top             =   240
         Width           =   7680
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdBusqueda 
         Height          =   4770
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   "Búsqueda de pacientes"
         Top             =   1080
         Width           =   10485
         _ExtentX        =   18494
         _ExtentY        =   8414
         _Version        =   393216
         ForeColor       =   0
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
         _Band(0).Cols   =   2
      End
      Begin VB.Label lblTeclea 
         BackColor       =   &H80000005&
         Caption         =   "Teclee el no. de expediente"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   17
         Top             =   285
         Width           =   2970
      End
      Begin VB.Label lblCargando 
         BackColor       =   &H80000005&
         Caption         =   "Cargando datos, por favor espere..."
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3360
         TabIndex        =   14
         Top             =   720
         Visible         =   0   'False
         Width           =   3615
      End
   End
End
Attribute VB_Name = "FrmBusquedaPacientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'-------------------------------------------------------------------------------------
'| Nombre del Proyecto      : Caja                                                   -
'| Nombre del Formulario    : frmBusquedaPacientes                                         -
'-------------------------------------------------------------------------------------
'| Objetivo: Tener una SuperBusqueda de Pacientes
'-------------------------------------------------------------------------------------
'| Análisis y Diseño        : Rodolfo Ramos G.
'| Autor                    : Rodolfo Ramos G.
'| Fecha de Creación        : 22/Ene/2001
'| Modificó                 : Nombre(s)
'| Fecha Terminación        :
'| Fecha última modificación: 22/Ene/2001
'-------------------------------------------------------------------------------------
Option Explicit
Public vgIntMaxRecords As Integer 'Propiedad para configurar el numero de registros a devolver
Public vgStrOtrosCampos As String 'Propiedad de la forma para incluir otros campos al grid
Public vgstrTipoPaciente As String 'Tipo de paciente "I"nterno o "E"xterno
Public vgblnSoloInternos As Boolean 'Si es internos poder mostrar solo los activos
Public vgstrTamanoCampo As String 'Contiene los tamaños de los campos
Public vgstrMovCve As String 'Para saber si la pantalla regresa el movimiento o la clave del paciente
Public vglngClavePaciente As Long 'Regresa la clave del paciente seleccionado en la busqueda
Public vglngMovtoPaciente As Long 'Regresa el movimiento del paciente seleccionado en la busqueda
Public vgblnPideClave As Boolean 'Bandera para que pida primero la clave del Paciente
Public vgblndecredito As Integer
Public vgstrForma As String 'Para el nombre de la forma desde la que se mandó llamar la búsqueda
Public vgblnCreditoActivo As Boolean 'Se utiliza en el POS cuando se selecciona un medico o un empleado para identificar si tiene credito activo

Dim vlblnCargaDatos As Boolean 'Bandera para cargar cuando deje de teclear
Dim vlvarColorCreditoActivo As Variant 'Color de la columna "Estado crédito", cuando el crédito está activo
Dim vlvarColorCreditoInactivo As Variant 'Color de la columna "Estado crédito", cuando el crédito está inactivo
Dim vlvarColorSinCredito As Variant 'Color de la columna "Estado crédito", cuando no tiene crédito
Dim vlvarColorNormal As Variant 'Color predeterminado

Public Function flngRegresaPaciente() As Long
    FrmBusquedaPacientes.Show vbModal
    flngRegresaPaciente = vglngMovtoPaciente
End Function

Function FConfiguraBusqueda(vlstrCriterio As String)
    Dim vlstrInstruccion As String 'Temporal para utilizar la instruccion
    Dim vlstrFiltroEstado As String 'Filtro de solo activos o facturados o todos
    Dim vlstrMovCve As String 'Instruccion para los movimientos o claves de pacientes
    Dim vlstrFrom As String 'Instrucción para almacenar la parte de From
    Dim vlstrFiltro As String
    
    If vgstrMovCve = "M" Then
        vlstrFiltroEstado = " "
        
        If vgstrTipoPaciente = "I" Then  'Internos
            If optSoloActivos.Value Then
                vlstrFiltroEstado = " and ExPacienteIngreso.chrEstatus = 'A' and nodepartamento.TNYCLAVEEMPRESA = " & vgintClaveEmpresaContable & " "
            ElseIf optSinFacturar.Value Then
                vlstrFiltroEstado = " and ExPacienteIngreso.intCuentaFacturada = 0 and ExPacienteIngreso.chrEstatus <> 'C' and nodepartamento.TNYCLAVEEMPRESA = " & vgintClaveEmpresaContable & " "
            End If
            
            vlstrMovCve = " ExPacienteIngreso.intNumCuenta as Cuenta, "
            vlstrFrom = " From ExPacienteIngreso " & _
                        " Inner Join SiTipoIngreso On ExPacienteIngreso.intCveTipoIngreso = SiTipoIngreso.intCveTipoIngreso And SiTipoIngreso.chrTipoIngreso = 'I' " & _
                        " LEFT OUTER JOIN ExPaciente ON ExPacienteIngreso.intNumPaciente = ExPaciente.intNumPaciente " & _
                        " LEFT OUTER JOIN CCEmpresa ON ExPacienteIngreso.intCveEmpresa = CCEmpresa.intCveEmpresa " & _
                        " LEFT OUTER JOIN AdTipoPaciente ON ExPacienteIngreso.intCveTipoPaciente = AdTipoPaciente.tnyCveTipoPaciente " & _
                        " INNER JOIN NODEPARTAMENTO ON ExPacienteIngreso.intCveDepartamento = NODEPARTAMENTO.SMICVEDEPARTAMENTO "
                        
                        
        Else  'Externos
            If optSoloActivos.Value Then
                vlstrFiltroEstado = " and ExPacienteIngreso.dtmFechaHoraEgreso is null and nodepartamento.TNYCLAVEEMPRESA = " & vgintClaveEmpresaContable & " "
            ElseIf optSinFacturar.Value Then
                vlstrFiltroEstado = " and ExPacienteIngreso.intCuentaFacturada = 0 and nodepartamento.TNYCLAVEEMPRESA = " & vgintClaveEmpresaContable & " "
            End If
            
            vlstrMovCve = " ExPacienteIngreso.intNumCuenta as Cuenta, "
            vlstrFrom = " from ExPacienteIngreso " & _
                        " Inner Join SiTipoIngreso On ExPacienteIngreso.intCveTipoIngreso = SiTipoIngreso.intCveTipoIngreso And SiTipoIngreso.chrTipoIngreso = 'E' " & _
                        " LEFT OUTER JOIN ExPaciente On ExPacienteIngreso.intNumPaciente = ExPaciente.intNumPaciente " & _
                        " LEFT OUTER join CCempresa on ExPacienteIngreso.intCveEmpresa = CCEmpresa.intCveEmpresa " & _
                        " LEFT OUTER JOIN AdTipoPaciente ON ExPacienteIngreso.intCveTipoPaciente = AdTipoPaciente.tnyCveTipoPaciente " & _
                        " INNER JOIN NODEPARTAMENTO ON ExPacienteIngreso.intCveDepartamento = NODEPARTAMENTO.SMICVEDEPARTAMENTO "
        
        End If
    Else
    
    

        
        '' validacion solo para SA y LA  caso 6778'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If vgstrForma = "frmSolicitudEstudio" Or vgstrForma = "frmSolicitudExamen" Or vgstrForma = "frmPOS" Or vgstrForma = "frmPacientesAtendidos" Then

            vlstrFiltro = ""
            vlstrMovCve = " distinct(ExPaciente.intNumPaciente) as Clave, "
        
            vlstrFrom = " From ExPaciente " & _
                        " left join AdTipoPaciente on ExPaciente.intCveTipoPaciente = AdTipoPaciente.tnyCveTipoPaciente "
            vlstrFrom = vlstrFrom & " left join Expacienteingreso on Expacienteingreso.intnumpaciente = Expaciente.intNumPaciente " & _
                                    " left join Sitipoingreso on Sitipoingreso.intcvetipoingreso = Expacienteingreso.intcvetipoingreso "
        Else
            
             ' este código es como estaba inicialmente la consulta(ANTES DEL CASO 6778), de esta manera se comporta igual en los demas módulos
            vlstrFiltro = ""
            vlstrMovCve = " ExPaciente.intNumPaciente as Clave, "
        
             vlstrFrom = " From ExPaciente " & _
                         " left join AdTipoPaciente on ExPaciente.intCveTipoPaciente = AdTipoPaciente.tnyCveTipoPaciente "
              vlstrFrom = vlstrFrom & " join Expacienteingreso on Expacienteingreso.intnumpaciente = Expaciente.intNumPaciente " & _
                                      " join Sitipoingreso on Sitipoingreso.intcvetipoingreso = Expacienteingreso.intcvetipoingreso "

        End If
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        vlstrFrom = vlstrFrom & " and Sitipoingreso.chrtipoingreso = " & "'" & vgstrTipoPaciente & "'" & _
                                          " left join CcEmpresa on ExPaciente.intCveEmpresa = CcEmpresa.intCveEmpresa "
                      
        If vgstrForma = "frmPOS" Then
          
            vlstrFrom = vlstrFrom & vlstrFiltro & _
                        " left join (Select Max(ExPacienteIngreso.intNumCuenta) intNumCuenta, ExPacienteIngreso.intNumPaciente, ExPacienteIngreso.intCveMedicoRelacionado, ExPacienteIngreso.intCveEmpleadoRelacionado " & _
                        "            From ExPacienteIngreso " & _
                        "            Inner Join SiTipoIngreso On ExPacienteIngreso.intCveTipoIngreso = SiTipoIngreso.intCveTipoIngreso And SiTipoIngreso.chrTipoIngreso = '" & vgstrTipoPaciente & "' " & _
                        "            Group By ExPacienteIngreso.intNumPaciente, intcveMEdicoRelacionado, intCveEmpleadoRelacionado ) Cuenta " & _
                        "            On ExPaciente.intNumPaciente = Cuenta.intNumPaciente " & _
                        " left outer join (Select * from cccliente " & _
                        " inner join nodepartamento on cccliente.smicvedepartamento =  nodepartamento.smicvedepartamento " & _
                        " and nodepartamento.tnyclaveempresa =  " & vgintClaveEmpresaContable & " ) Clientes " & _
                        " On Case AdTipoPaciente.chrTipo  " & _
                        "    when 'CO' then ExPaciente.intCveEmpresa " & _
                        "    when 'EM' then Cuenta.intCveEmpleadoRelacionado " & _
                        "    when 'ME' then Cuenta.intCveMedicoRelacionado " & _
                        "    End  = Clientes.intNumReferencia  and AdTipoPaciente.chrTipo = Clientes.chrTipoCliente "
        End If
          
    End If
    
        If optTipo(0) Then    'Paterno
            vlstrInstruccion = vlstrInstruccion & "Select" & vlstrMovCve & "trim(ExPaciente.vchApellidoPaterno)||' '||trim(ExPaciente.vchApellidoMaterno)||' '||trim(ExPaciente.vchNombre) as Nombre " & vgStrOtrosCampos & ", trim(ExPaciente.vchCurp) CURP " & IIf(vgstrMovCve = "M", " ,ExPaciente.intNumPaciente as Expediente ", "") & _
                               vlstrFrom
            vlstrInstruccion = vlstrInstruccion & "where trim(ExPaciente.vchApellidoPaterno)||' '||trim(ExPaciente.vchApellidoMaterno)||' '||trim(ExPaciente.vchNombre) like '" & vlstrCriterio & "%' " & _
                                vlstrFiltroEstado '&
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''cambio caso 6778
            If vgstrForma = "frmSolicitudEstudio" Or vgstrForma = "frmSolicitudExamen" Or vgstrForma = "frmPOS" Or vgstrForma = "frmPacientesAtendidos" Then
               vlstrInstruccion = vlstrInstruccion & " Order by Nombre " ' para que no truene el DISTINCT
            Else
                ' estas son las lineas originales que contenia la busqueda
                vlstrInstruccion = vlstrInstruccion & " Order by ExPaciente.vchApellidoPaterno, ExPaciente.vchApellidoMaterno, ExPaciente.vchNombre "
            
                If vgstrMovCve = "M" Then
                    vlstrInstruccion = vlstrInstruccion & ", ExPacienteIngreso.dtmFechaHoraIngreso Desc"
                End If
            End If
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
    
        ElseIf optTipo(1) Then    ' Materno
            vlstrInstruccion = vlstrInstruccion & "Select" & vlstrMovCve & "trim(ExPaciente.vchApellidoPaterno)||' '||trim(ExPaciente.vchApellidoMaterno)||' '||trim(ExPaciente.vchNombre) as Nombre " & vgStrOtrosCampos & ", trim(ExPaciente.vchCurp) CURP " & IIf(vgstrMovCve = "M", " ,ExPaciente.intNumPaciente  as Expediente ", "") & _
                               vlstrFrom
            vlstrInstruccion = vlstrInstruccion & "where trim(ExPaciente.vchApellidoMaterno)||' '||trim(ExPaciente.vchNombre) like '" & vlstrCriterio & "%' " & _
                                vlstrFiltroEstado '&
                                
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''cambio caso 6778
            If vgstrForma = "frmSolicitudEstudio" Or vgstrForma = "frmSolicitudExamen" Or vgstrForma = "frmPOS" Or vgstrForma = "frmPacientesAtendidos" Then
               vlstrInstruccion = vlstrInstruccion & " Order by Nombre" ' para que no truene el DISTINCT
            Else
                ' estas son las lineas originales que contenia la busqueda
                vlstrInstruccion = vlstrInstruccion & " Order by ExPaciente.vchApellidoPaterno, ExPaciente.vchApellidoMaterno, ExPaciente.vchNombre "
                If vgstrMovCve = "M" Then
                   vlstrInstruccion = vlstrInstruccion & ", ExPacienteIngreso.dtmFechaHoraIngreso Desc"
                End If
            End If
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                
          
        ElseIf optTipo(2) Then    ' Nombre
            vlstrInstruccion = vlstrInstruccion & "Select" & vlstrMovCve & "trim(ExPaciente.vchNombre)||' '||trim(ExPaciente.vchApellidoPaterno)||' '||trim(ExPaciente.vchApellidoMaterno) as Nombre " & vgStrOtrosCampos & ", trim(ExPaciente.vchCurp) CURP " & IIf(vgstrMovCve = "M", " ,ExPaciente.intNumPaciente as Expediente ", "") & _
                               vlstrFrom
            vlstrInstruccion = vlstrInstruccion & "where trim(ExPaciente.vchNombre)||' '|| trim(ExPaciente.vchApellidoPaterno)||' '||trim(ExPaciente.vchApellidoMaterno) like '" & vlstrCriterio & "%' " & _
                               vlstrFiltroEstado '&

            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''cambio caso 6778
            If vgstrForma = "frmSolicitudEstudio" Or vgstrForma = "frmSolicitudExamen" Or vgstrForma = "frmPOS" Or vgstrForma = "frmPacientesAtendidos" Then
               vlstrInstruccion = vlstrInstruccion & " Order by Nombre" ' para que no truene el DISTINCT
            Else
                ' estas son las lineas originales que contenia la busqueda
                vlstrInstruccion = vlstrInstruccion & "Order by ExPaciente.vchNombre, ExPaciente.vchApellidoPaterno,  ExPaciente.vchApellidoMaterno "
            
                If vgstrMovCve = "M" Then
                    vlstrInstruccion = vlstrInstruccion & ", ExPacienteIngreso.dtmFechaHoraIngreso Desc"
                End If
            End If
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
           
        ElseIf optTipo(3) Then    ' CURP
            vlstrInstruccion = vlstrInstruccion & "Select" & vlstrMovCve & " trim(ExPaciente.vchApellidoPaterno)||' '||trim(ExPaciente.vchApellidoMaterno)||' '||trim(ExPaciente.vchNombre) as Nombre " & vgStrOtrosCampos & ", trim(ExPaciente.vchCurp) CURP " & IIf(vgstrMovCve = "M", " ,ExPaciente.intNumPaciente as Expediente ", "") & _
                               vlstrFrom
            vlstrInstruccion = vlstrInstruccion & "where trim(ExPaciente.vchCurp) like '" & vlstrCriterio & "%' " & _
                                vlstrFiltroEstado '&
                                
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''cambio caso 6778
            If vgstrForma = "frmSolicitudEstudio" Or vgstrForma = "frmSolicitudExamen" Or vgstrForma = "frmPOS" Or vgstrForma = "frmPacientesAtendidos" Then
               vlstrInstruccion = vlstrInstruccion & " Order by Nombre" ' para que no truene el DISTINCT
            Else
                  ' estas son las lineas originales que contenia la busqueda
                vlstrInstruccion = vlstrInstruccion & " Order by ExPaciente.vchApellidoPaterno, ExPaciente.vchApellidoMaterno, ExPaciente.vchNombre "
                
                If vgstrMovCve = "M" Then
                    vlstrInstruccion = vlstrInstruccion & ", ExPacienteIngreso.dtmFechaHoraIngreso Desc"
                End If
            End If
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                
            
                
        ElseIf optTipo(4) Then    ' No. de Expediente electrónico
            vlstrInstruccion = vlstrInstruccion & "Select" & vlstrMovCve & " trim(ExPaciente.vchApellidoPaterno)||' '||trim(ExPaciente.vchApellidoMaterno)||' '||trim(ExPaciente.vchNombre) as Nombre " & vgStrOtrosCampos & ", trim(ExPaciente.vchCurp) CURP " & IIf(vgstrMovCve = "M", " ,ExPaciente.intNumPaciente as Expediente ", "") & _
                               vlstrFrom
            vlstrInstruccion = vlstrInstruccion & "where ExPaciente.intNumPaciente like '" & vlstrCriterio & "%' " & _
                                vlstrFiltroEstado '&
                                
                                
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''cambio caso 6778
            If vgstrForma = "frmSolicitudEstudio" Or vgstrForma = "frmSolicitudExamen" Or vgstrForma = "frmPOS" Or vgstrForma = "frmPacientesAtendidos" Then
               vlstrInstruccion = vlstrInstruccion & " Order by Nombre" ' para que no truene el DISTINCT
            Else
                ' estas son las lineas originales que contenia la busqueda
                vlstrInstruccion = vlstrInstruccion & " Order by ExPaciente.vchApellidoPaterno, ExPaciente.vchApellidoMaterno, ExPaciente.vchNombre "
            
                  If vgstrMovCve = "M" Then
                       vlstrInstruccion = vlstrInstruccion & ", ExPacienteIngreso.dtmFechaHoraIngreso Desc"
                  End If
            
            End If
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                      
      
            
        End If
    
    FConfiguraBusqueda = vlstrInstruccion
End Function

Function fCalculaColumnas(vlintCuantosRenglones As Integer) As Integer()
    '----------------------------------------------------------------------------
    'Regresa un arreglo con numeros de los tamaños de las columnas del grid
    'tomando como base el string que nos manda como propiedad de la forma
    'Ej. "300,4000,5000"
    '----------------------------------------------------------------------------
    Dim vlintContComas As Integer 'Recorrer todo el String caracter por caracter
    Dim vlchrTemporal As String 'Guarda los numeros concatenados para obtener la cantidad
    Dim vlintCaracter As Integer  'Validar si es numero el caracter
    Dim vlintContRenglones As Byte 'Controlar los renglones del array
    Dim alintArreglo() As Integer  'Arreglo donde se guardan los datos
    
    ReDim alintArreglo(vlintCuantosRenglones)
    
    vlchrTemporal = ""
    vlintContRenglones = 0
    
    For vlintContComas = 1 To Len(vgstrTamanoCampo)
        vlintCaracter = Asc(Mid(vgstrTamanoCampo, vlintContComas, 1))
        If (vlintCaracter >= 48 And vlintCaracter <= 57) Then
            'Sí es numero
            vlchrTemporal = vlchrTemporal & Chr(vlintCaracter)
        Else 'Se encontro otro caracter que no es numero
            
            alintArreglo(vlintContRenglones) = Int(vlchrTemporal)
            vlintContRenglones = vlintContRenglones + 1
            vlchrTemporal = ""
        End If
    Next vlintContComas
    If vlintContRenglones > 0 Then
        alintArreglo(vlintContRenglones) = Int(vlchrTemporal)
    End If
    fCalculaColumnas = alintArreglo
    
End Function

Sub pConfiguraGrid(ObjGrd As MSHFlexGrid)
    With ObjGrd
        .FixedCols = 1
        .FixedRows = 1
        .ColWidth(0) = 100
        .ColWidth(1) = 3200
        .Row = 1
        .Col = 1
        .Rows = 2
        .TextMatrix(1, 1) = ""
        .ScrollBars = flexScrollBarBoth
    End With
End Sub

Sub pLimpiaGrid(ObjGrd As MSHFlexGrid)
    Dim vlbytColumnas As Byte
    With ObjGrd
        .FormatString = ""
        .Row = 1
        .Col = 1
        .Rows = 2
        For vlbytColumnas = 1 To .Cols - 1
            .TextMatrix(1, vlbytColumnas) = ""
        Next vlbytColumnas
        .TextMatrix(1, 1) = ""
    End With
End Sub

Sub pTamanoColumnas(alintTamanos() As Integer)
    Dim vlContador As Integer
    With grdBusqueda
      For vlContador = 0 To UBound(alintTamanos)
        If .Cols - 1 > vlContador Then
          If .TextMatrix(0, vlContador + 1) <> "" Then
            .ColWidth(vlContador + 1) = IIf(alintTamanos(vlContador) <> 0, alintTamanos(vlContador) + 300, 2000)
            'Para ampliar la columna de las fechas de ingreso y egreso
            If .TextMatrix(0, vlContador + 1) = "Fecha ing." Then
                .ColWidth(vlContador + 1) = IIf(alintTamanos(vlContador) <> 0, alintTamanos(vlContador) + 1500, 2500)
                .TextMatrix(0, vlContador + 1) = "Fecha ingreso"
            End If
            If .TextMatrix(0, vlContador + 1) = "Fecha egr." Then
                .ColWidth(vlContador + 1) = IIf(alintTamanos(vlContador) <> 0, alintTamanos(vlContador) + 1500, 2500)
                .TextMatrix(0, vlContador + 1) = "Fecha egreso"
            End If
            If .TextMatrix(0, vlContador + 1) = "Fecha" Then
                .ColAlignment(vlContador + 1) = vbAlignLeft
                .ColWidth(vlContador + 1) = 2500
            End If
            If .TextMatrix(0, vlContador + 1) = "FECHA NACIMIENTO" Then
                .ColAlignment(vlContador + 1) = vbAlignLeft
                .ColWidth(vlContador + 1) = 2000
            End If
            If .TextMatrix(0, vlContador + 1) = "FECHA INGRESO" Then
                .ColAlignment(vlContador + 1) = vbAlignLeft
                .ColWidth(vlContador + 1) = 2500
            End If
            If .TextMatrix(0, vlContador + 1) = "FECHA EGRESO" Then
                .ColAlignment(vlContador + 1) = vbAlignLeft
                .ColWidth(vlContador + 1) = 2500
            End If
            If .TextMatrix(0, vlContador + 1) = "EGRESO" Then
                .ColAlignment(vlContador + 1) = vbAlignLeft
                .ColWidth(vlContador + 1) = 2500
            End If
            If .TextMatrix(0, vlContador + 1) = "INGRESO" Then
                .ColAlignment(vlContador + 1) = vbAlignLeft
                .ColWidth(vlContador + 1) = 2500
            End If
            If .TextMatrix(0, vlContador + 1) = "FECHA" Then
                .ColAlignment(vlContador + 1) = vbAlignLeft
                .ColWidth(vlContador + 1) = 2500
            End If
            .ColAlignmentFixed(vlContador + 1) = 4
          Else
            .ColWidth(vlContador + 1) = 0
          End If
        End If
      Next vlContador
    End With
End Sub

Sub PSuperBusqueda(vlstrCriterio As String, vlintTipo As Integer)
    Dim vlstrInstruccion As String
    Dim vlintRenglones As Integer
    Dim vlintColumnas As Integer
    Dim rsDatos As New ADODB.Recordset
    Dim alintTamanos() As Integer
    
    vlstrInstruccion = FConfiguraBusqueda(vlstrCriterio)
    
    Set rsDatos = frsRegresaRs(vlstrInstruccion, adLockReadOnly, adOpenForwardOnly, vgIntMaxRecords)
    lblCargando.Visible = True
    lblCargando.Refresh
    grdBusqueda.Redraw = False
    
    With rsDatos
        ReDim alintTamanos(.Fields.Count)
        alintTamanos() = fCalculaColumnas(.Fields.Count)
    End With
    
    Call pLimpiaGrid(grdBusqueda)
    
    grdBusqueda.Rows = IIf(rsDatos.RecordCount = 0, 2, rsDatos.RecordCount + 1)
    With grdBusqueda
       For vlintColumnas = 1 To rsDatos.Fields.Count
            If rsDatos.Fields(vlintColumnas - 1).Name = "FECHA_NAC" Then
                .FormatString = .FormatString & "|" & "FECHA NACIMIENTO"
            ElseIf rsDatos.Fields(vlintColumnas - 1).Name = "FECHA_ING" Then
                .FormatString = .FormatString & "|" & "FECHA INGRESO"
            ElseIf rsDatos.Fields(vlintColumnas - 1).Name = "FECHA_EGR" Then
                .FormatString = .FormatString & "|" & "FECHA EGRESO"
            Else
                .FormatString = .FormatString & "|" & rsDatos.Fields(vlintColumnas - 1).Name
            End If
       Next vlintColumnas
       For vlintRenglones = 1 To rsDatos.RecordCount
            .Row = vlintRenglones
            For vlintColumnas = 0 To rsDatos.Fields.Count - 1
                .Col = vlintColumnas + 1
                If Trim(rsDatos.Fields(vlintColumnas).Name) = "Estado crédito" Then
                    If rsDatos.Fields(vlintColumnas).Value = "Sin crédito" Then
                        .CellForeColor = vlvarColorSinCredito
                    End If
                    If rsDatos.Fields(vlintColumnas).Value = "Activo" Then
                        .CellForeColor = vlvarColorCreditoActivo
                    End If
                    If rsDatos.Fields(vlintColumnas).Value = "Inactivo" Then
                        .CellForeColor = vlvarColorCreditoInactivo
                    End If
                Else
                    .CellForeColor = vlvarColorNormal
                End If
                If IsNull(rsDatos.Fields(vlintColumnas).Value) Then
                    .TextMatrix(vlintRenglones, vlintColumnas + 1) = ""
                Else
                    .TextMatrix(vlintRenglones, vlintColumnas + 1) = rsDatos.Fields(vlintColumnas).Value
                End If
            Next vlintColumnas
             rsDatos.MoveNext
       Next vlintRenglones
       .Col = 1
       .Row = 1
    End With
    Call pTamanoColumnas(alintTamanos())
    lblCargando.Visible = False
    grdBusqueda.Redraw = True
    rsDatos.Close
    
End Sub

Private Sub Form_Activate()
   
    If vgstrMovCve = "C" Then optTipo(4).Enabled = False
    frmOpcion.Visible = Not vgblnPideClave
    freBusqueda.Visible = Not vgblnPideClave
    freEstadoPaciente.Visible = Not vgblnPideClave
    freClavePaciente.Visible = vgblnPideClave
    vglngMovtoPaciente = -1
    vglngClavePaciente = -1
    FrmBusquedaPacientes.Height = IIf(vgblnPideClave, 920, 6975)
    FrmBusquedaPacientes.Width = IIf(vgblnPideClave, 4800, 10995)
    
    'Verificar que tipo de paciente es para cambiar el titulo
    If vgstrTipoPaciente = "I" Then
        Me.Caption = "Búsqueda de pacientes internos"
    End If
    If vgstrTipoPaciente = "E" Then
        Me.Caption = "Búsqueda de pacientes externos"
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyEscape Then
        vglngMovtoPaciente = -1
        vglngClavePaciente = vglngMovtoPaciente
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()

    vlvarColorCreditoActivo = &HFF0000
    vlvarColorCreditoInactivo = &H80&
    vlvarColorSinCredito = &H80000012
    vlvarColorNormal = &H80000008

End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   If RTrim(txtBusqueda.Text) = "" And RTrim(txtClavePaciente.Text) = "" Then
        If vgstrMovCve = "M" Then
            vglngMovtoPaciente = -1
        Else
            vglngClavePaciente = -1
        End If
   End If
   vgstrTipoPaciente = ""
   
End Sub

Private Sub grdBusqueda_DblClick()
    Call grdBusqueda_KeyDown(vbKeyReturn, 0)
End Sub

Private Sub grdBusqueda_GotFocus()
    If txtBusqueda.Text = "" Then txtBusqueda.SetFocus
End Sub

Private Sub grdBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
        
    Select Case KeyCode
        Case vbKeyReturn
            vgblnCreditoActivo = False
            If RTrim(grdBusqueda.TextMatrix(grdBusqueda.Row, 1)) = "" Then
                vglngMovtoPaciente = -1
            Else
                vglngMovtoPaciente = IIf(CLng(grdBusqueda.TextMatrix(grdBusqueda.Row, 1)) = 0, -1, CLng(grdBusqueda.TextMatrix(grdBusqueda.Row, 1)))
                If vgstrForma = "frmPOS" Then
                    vgblnCreditoActivo = IIf(Trim(grdBusqueda.TextMatrix(grdBusqueda.Row, 4)) = "Activo", True, False)
                End If
            End If
            vglngClavePaciente = vglngMovtoPaciente
            Unload Me
        Case vbKeyEscape
            vgblnCreditoActivo = False
            vglngMovtoPaciente = -1
            vglngClavePaciente = vglngMovtoPaciente
            Unload Me
    End Select
    
End Sub

Private Sub optSinFacturar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtBusqueda.SetFocus
    txtBusqueda_KeyUp 0, 0
End Sub

Private Sub optSoloActivos_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtBusqueda.SetFocus
    txtBusqueda_KeyUp 0, 0
End Sub

Private Sub optTipo_Click(Index As Integer)
    
    Call pLimpiaGrid(grdBusqueda)
    Call pConfiguraGrid(grdBusqueda)
    txtBusqueda.Text = ""
    If txtBusqueda.Enabled Then
        txtBusqueda.SetFocus
    End If
    If optTipo(0) Then
        lblTeclea.Caption = "Teclee el Apellido Paterno"
        txtBusqueda.Left = 2760
        txtBusqueda.Width = 7840
    ElseIf optTipo(1) Then
        lblTeclea.Caption = "Teclee el Apellido Materno"
        txtBusqueda.Left = 2760
        txtBusqueda.Width = 7840
    ElseIf optTipo(2) Then
        lblTeclea.Caption = "Teclee el Nombre"
        txtBusqueda.Left = 1920
        txtBusqueda.Width = 8685
    ElseIf optTipo(3) Then
        lblTeclea.Caption = "Teclee el CURP"
        txtBusqueda.Left = 1600
        txtBusqueda.Width = 9000
    ElseIf optTipo(4) Then
        lblTeclea.Caption = "Teclee el no. de expediente"
        txtBusqueda.Left = 2925
        txtBusqueda.Width = 7680
    End If
        
End Sub

Private Sub optTodos_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtBusqueda.SetFocus
    txtBusqueda_KeyUp 0, 0
End Sub

Private Sub tmrCarga_Timer()
    Call PSuperBusqueda(txtBusqueda.Text, 1)
    tmrCarga.Enabled = False
End Sub

Private Sub txtBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyReturn
            If Trim(txtBusqueda.Text) <> "" Then
                If tmrCarga.Enabled Then
                    Call PSuperBusqueda(txtBusqueda.Text, 1)
                End If
            End If
            grdBusqueda.SetFocus
            
            If vgblndecredito = 1 Then
                If grdBusqueda.TextMatrix(1, 1) = "" Then
                    Unload Me
                End If
            End If
        
        
        Case vbKeyEscape
            Unload Me
    End Select
    
End Sub

Private Sub txtBusqueda_KeyPress(KeyAscii As Integer)
    
    If optTipo(4) And Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If

    If KeyAscii = 39 Then KeyAscii = 0
    tmrCarga.Enabled = False
    KeyAscii = Asc(UCase(Chr(KeyAscii)))

End Sub

Private Sub txtBusqueda_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyDown Then
        grdBusqueda.SetFocus
        Exit Sub
    End If
    If txtBusqueda.Text = "" Then
        Call pLimpiaGrid(grdBusqueda)
    Else
        tmrCarga.Enabled = True
    End If
    
End Sub

Private Sub txtClavePaciente_GotFocus()
    lblMoviClave.Caption = IIf(vgstrMovCve = "C", "Clave ", "Cuenta ") & "del paciente"
End Sub

Private Sub pAgrandaForma()
    Dim vlintContRow As Integer
    Dim vlintContCol As Integer
    
    For vlintContRow = FrmBusquedaPacientes.Width To 10995 Step 200
        'FrmBusquedaPacientes.Height = 6195
        'FrmBusquedaPacientes.Width = 9810
        If FrmBusquedaPacientes.Height < 6975 Then
            FrmBusquedaPacientes.Height = vlintContRow
        End If
        FrmBusquedaPacientes.Width = vlintContRow
    Next
    FrmBusquedaPacientes.Height = 6975
    FrmBusquedaPacientes.Width = 10995
    
End Sub

Private Sub txtClavePaciente_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rs As New ADODB.Recordset
    Dim vlstrSentencia As String
    Dim vlstrFiltroEstado As String
    
    If KeyCode = vbKeyReturn And RTrim(txtClavePaciente) = "" Then
        
        frmOpcion.Visible = True
        freBusqueda.Visible = True
        freEstadoPaciente.Visible = True
        freClavePaciente.Visible = False
        
        pAgrandaForma
        
    ElseIf KeyCode = vbKeyReturn Then
        vlstrFiltroEstado = " "
        If vgstrMovCve = "M" Then 'Movimiento del pac
            If vgstrTipoPaciente = "I" Then  'Internos
                If optSoloActivos.Value Then
                    vlstrFiltroEstado = " and adAdmision.chrEstatusAdmision = 'A' "
                ElseIf optSinFacturar.Value Then
                    vlstrFiltroEstado = " and adAdmision.bitFacturado = 0 and adadmision.CHRESTATUSADMISION <> 'C' "
                End If
                vlstrSentencia = "Select count(numNumCuenta) from Adadmision inner join NODepartamento on NODepartamento.smiCveDepartamento = Adadmision.intCveDepartamento where numNumCuenta = " & txtClavePaciente & vlstrFiltroEstado & " and tnyClaveEmpresa = " & vgintClaveEmpresaContable
            
            Else 'Externos
                If optSoloActivos.Value Or optSinFacturar.Value Then
                    vlstrFiltroEstado = " and RegistroExterno.bitFacturado = 0 "
                End If
                vlstrSentencia = "Select count(IntNumCuenta) from RegistroExterno inner join NODepartamento on NODepartamento.smiCveDepartamento = RegistroExterno.intCveDepartamento where IntNumCuenta = " & txtClavePaciente & vlstrFiltroEstado & " and tnyClaveEmpresa = " & vgintClaveEmpresaContable
            End If
        Else
            If vgstrTipoPaciente = "I" Then  'Internos
                vlstrSentencia = "Select count(numCvePaciente) from AdPaciente where numCvePaciente = " & txtClavePaciente & vlstrFiltroEstado
            Else 'Externos
                vlstrSentencia = "Select count(IntNumPaciente) from Externo where IntNumPaciente = " & txtClavePaciente & vlstrFiltroEstado
            End If
        End If
        
        Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
        If rs.Fields(0) = 0 Then
            MsgBox SIHOMsg(12), vbExclamation, "Mensaje"
            pEnfocaTextBox txtClavePaciente
        Else
            vglngMovtoPaciente = -1
            vglngClavePaciente = -1
            vglngMovtoPaciente = IIf(CLng(txtClavePaciente.Text) = 0, -1, CLng(txtClavePaciente.Text))
            vglngClavePaciente = vglngMovtoPaciente
            FrmBusquedaPacientes.Hide
            Unload FrmBusquedaPacientes
        End If
    End If
    
End Sub

Private Sub txtClavePaciente_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
        KeyAscii = 7
    End If
End Sub
