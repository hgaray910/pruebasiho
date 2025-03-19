VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmCambioCheque 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambio de cheques / Transferencias"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11730
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   11730
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame7 
      Height          =   765
      Left            =   9480
      TabIndex        =   27
      Top             =   5355
      Width           =   2175
   End
   Begin VB.Frame Frame6 
      Height          =   765
      Left            =   6090
      TabIndex        =   26
      Top             =   5355
      Width           =   2175
   End
   Begin VB.Frame Frame2 
      Height          =   315
      Left            =   6090
      TabIndex        =   25
      Top             =   4980
      Width           =   5565
   End
   Begin VB.Frame frmChequeAplicado 
      Height          =   3840
      Left            =   60
      TabIndex        =   23
      Top             =   1110
      Width           =   11595
      Begin VB.CommandButton cmdSeleccionar 
         Caption         =   "Seleccionar / Quitar"
         Height          =   315
         Left            =   9120
         TabIndex        =   24
         Top             =   3400
         Width           =   2280
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdCheque 
         Height          =   3195
         Left            =   45
         TabIndex        =   6
         Top             =   135
         Width           =   11505
         _ExtentX        =   20294
         _ExtentY        =   5636
         _Version        =   393216
         Cols            =   7
         GridColor       =   12632256
         FormatString    =   "|Fecha|Banco|Número|Beneficiario|Estado|Persona aplicó al corte"
         _NumberOfBands  =   1
         _Band(0).Cols   =   7
      End
      Begin VB.Shape Shape3 
         FillColor       =   &H00808000&
         FillStyle       =   0  'Solid
         Height          =   180
         Left            =   1080
         Top             =   3480
         Width           =   180
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Transferencias"
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   1305
         TabIndex        =   32
         Top             =   3480
         Width           =   1050
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   180
         Left            =   120
         Top             =   3480
         Width           =   180
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Cheques"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   345
         TabIndex        =   31
         Top             =   3480
         Width           =   630
      End
   End
   Begin VB.Frame fraChequeAplicado 
      Caption         =   "Información de la aplicación"
      Height          =   1140
      Left            =   90
      TabIndex        =   19
      Top             =   4980
      Width           =   5970
      Begin VB.Label lbPersonaAplico 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1365
         TabIndex        =   9
         Top             =   660
         Width           =   4470
      End
      Begin VB.Label lblFechaAplicacion 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   4455
         TabIndex        =   8
         Top             =   300
         Width           =   1380
      End
      Begin VB.Label lblNumeroCorte 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1365
         TabIndex        =   7
         Top             =   300
         Width           =   1380
      End
      Begin VB.Label lblTituloPersonaAplico 
         AutoSize        =   -1  'True
         Caption         =   "Persona aplicó"
         Height          =   195
         Left            =   105
         TabIndex        =   22
         Top             =   720
         Width           =   1050
      End
      Begin VB.Label lblTituloFecha 
         AutoSize        =   -1  'True
         Caption         =   "Fecha la aplicación"
         Height          =   195
         Left            =   2880
         TabIndex        =   21
         Top             =   360
         Width           =   1380
      End
      Begin VB.Label lblTituloCorte 
         AutoSize        =   -1  'True
         Caption         =   "Número corte"
         Height          =   195
         Left            =   105
         TabIndex        =   20
         Top             =   360
         Width           =   960
      End
   End
   Begin VB.Frame Frame3 
      Height          =   765
      Left            =   8310
      TabIndex        =   18
      Top             =   5355
      Width           =   1125
      Begin VB.CommandButton cmdCancelar 
         Height          =   495
         Left            =   555
         Picture         =   "frmCambioCheque.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Cancelar la aplicación del cheque o transferencia"
         Top             =   165
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdGuardar 
         Height          =   495
         Left            =   60
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmCambioCheque.frx":04F2
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Aplicar cheque o transferencia al corte"
         Top             =   165
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1065
      Left            =   60
      TabIndex        =   12
      Top             =   45
      Width           =   11595
      Begin VB.Frame Frame18 
         Caption         =   "Medio de pago"
         Height          =   715
         Left            =   3480
         TabIndex        =   28
         Top             =   185
         Width           =   3390
         Begin VB.OptionButton optMediopago 
            Caption         =   "Cheques"
            Height          =   195
            Index           =   1
            Left            =   960
            TabIndex        =   30
            ToolTipText     =   "Selección de cheques"
            Top             =   340
            Width           =   925
         End
         Begin VB.OptionButton optMediopago 
            Caption         =   "Transferencias"
            Height          =   195
            Index           =   2
            Left            =   1940
            TabIndex        =   29
            ToolTipText     =   "Selección de transferencias"
            Top             =   340
            Width           =   1355
         End
         Begin VB.OptionButton optMediopago 
            Caption         =   "Todos"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   1
            ToolTipText     =   "Selección del medio de pago (Todos)"
            Top             =   340
            Value           =   -1  'True
            Width           =   765
         End
      End
      Begin MSMask.MaskEdBox mskFecIni 
         Height          =   315
         Left            =   7620
         TabIndex        =   2
         Top             =   270
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mmm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Frame Frame4 
         Height          =   825
         Left            =   3315
         TabIndex        =   17
         Top             =   120
         Width           =   85
      End
      Begin VB.CommandButton cmdCargar 
         Caption         =   "Cargar"
         Height          =   315
         Left            =   9225
         TabIndex        =   5
         Top             =   605
         Width           =   2280
      End
      Begin VB.ComboBox cboEstado 
         Height          =   315
         Left            =   9825
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   235
         Width           =   1695
      End
      Begin MSMask.MaskEdBox mskFecFin 
         Height          =   315
         Left            =   7620
         TabIndex        =   3
         Top             =   600
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mmm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Left            =   7035
         TabIndex        =   16
         Top             =   300
         Width           =   465
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         Height          =   195
         Left            =   7035
         TabIndex        =   15
         Top             =   645
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Estado"
         Height          =   195
         Left            =   9210
         TabIndex        =   14
         Top             =   285
         Width           =   495
      End
      Begin VB.Label lblDepartamento 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   3105
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Departamento"
         Height          =   195
         Left            =   135
         TabIndex        =   13
         Top             =   240
         Width           =   1005
      End
   End
End
Attribute VB_Name = "frmCambioCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Grid cheques:
Const cintColIdCheque = 1
Const cintColFecha = 2
Const cintColBanco = 3
Const cintColNumero = 4
Const cintColBeneficiario = 5
Const cintColCantidad = 6
Const cintColbitPesos = 7
Const cintColMoneda = 8
Const cintColEstatus = 9
Const cIntColEstado = 10
Const cintColIdAplicacion = 11
Const cintColIdCorteAfectado = 12
Const cintColFechaAplicado = 13
Const cintColPersonaAplico = 14
Const cintColMedioPago = 15
Const cintColumnas = 16
Const cstrTitulos = "|IdCheque|Fecha|Banco|Número|Beneficiario|Cantidad|bitPesos|Moneda|Estatus|Estado|IdAplicacion|IdCorteAfectado|Fecha aplicado|Persona que aplicó"

Const cintIdTodos = -1
Const cintIdSinAplicar = 1
Const cintIdAplicados = 2

Const cstrFormato = "############.##"

Dim aFormasPago() As FormasPago
Dim ldtmFecha As Date 'Fecha actual
Dim lblnMensaje As Boolean 'Para saber si se muestra mensaje al cargar la información
Dim ldblTipoCambio As Double 'Tipo de cambio al que se recibe un cheque en dlls.
Dim llngPersonaGraba As Long 'Persona que guarda o cancela datos
Dim llngNumCorte As Long 'Corte en el que se guarda la información
Dim llngCorteGrabando As Long 'Estado del corte
Dim llngRenglonSel As Long 'Renglón seleccionado


Dim rs As New ADODB.Recordset
Dim vlstrMedioPago As String


Private Sub pLimpiaGrid()
    On Error GoTo NotificaError

    Dim intcontador As Integer

    cmdGuardar.Enabled = False
    cmdCancelar.Enabled = False

    llngRenglonSel = 0

    With grdCheque
        .Cols = cintColumnas
        .Rows = 2
        .FixedCols = 1
        .FixedRows = 1
        .FormatString = cstrTitulos
        
        .ColWidth(0) = 150
        .ColWidth(cintColIdCheque) = 0
        .ColWidth(cintColFecha) = 1100
        .ColWidth(cintColBanco) = 2700
        .ColWidth(cintColNumero) = 1000
        .ColWidth(cintColBeneficiario) = 2700
        .ColWidth(cintColCantidad) = 1100
        .ColWidth(cintColbitPesos) = 0
        .ColWidth(cintColMoneda) = 1000
        .ColWidth(cintColEstatus) = 0
        .ColWidth(cIntColEstado) = 1400
        .ColWidth(cintColIdAplicacion) = 0
        .ColWidth(cintColIdCorteAfectado) = 0
        .ColWidth(cintColFechaAplicado) = 0
        .ColWidth(cintColPersonaAplico) = 0
        .ColWidth(cintColMedioPago) = 0
        
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(cintColFecha) = flexAlignLeftCenter
        .ColAlignment(cintColBanco) = flexAlignLeftCenter
        .ColAlignment(cintColNumero) = flexAlignRightCenter
        .ColAlignment(cintColBeneficiario) = flexAlignLeftCenter
        .ColAlignment(cintColCantidad) = flexAlignRightCenter
        .ColAlignment(cintColMoneda) = flexAlignLeftCenter
        .ColAlignment(cIntColEstado) = flexAlignLeftCenter
        .ColAlignment(cintColFechaAplicado) = flexAlignLeftCenter
        .ColAlignment(cintColPersonaAplico) = flexAlignLeftCenter
        
        For intcontador = 0 To .Cols - 1
            .TextMatrix(1, intcontador) = ""
            .ColAlignmentFixed(intcontador) = flexAlignCenterCenter
        Next intcontador
            
    End With

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pLimpiaGrid"))
End Sub

Private Sub cmdCancelar_Click()
    On Error GoTo NotificaError

    Dim lngIdCancelacion As Long
    Dim blnerror As Boolean
    Dim intcontador As Integer
    
    If fblnCancelacionValida() Then
    
        EntornoSIHO.ConeccionSIHO.BeginTrans
        
        llngNumCorte = flngNumeroCorte(vgintNumeroDepartamento, vglngNumeroEmpleado, "C")
        
        If llngNumCorte = 0 Then
            EntornoSIHO.ConeccionSIHO.RollbackTrans
            'No se encontró un corte abierto.
            MsgBox SIHOMsg(659), vbExclamation + vbOKOnly, "Mensaje"
        Else
            '------------------------------------------------------------------
            ' Bloquear el corte
            '------------------------------------------------------------------
            llngCorteGrabando = 1
            frsEjecuta_SP CStr(llngNumCorte) & "|Grabando", "Sp_PvUpdEstatusCorte", True, llngCorteGrabando
            If llngCorteGrabando <> 2 Then
                EntornoSIHO.ConeccionSIHO.RollbackTrans
                
                'En este momento se está afectando el corte, espere un momento e intente de nuevo.
                MsgBox SIHOMsg(779), vbExclamation + vbOKOnly, "Mensaje"
            Else
                '------------------------------------------------------
                ' Desaplicar los cheques
                '------------------------------------------------------
                vgstrParametrosSP = "-1" _
                & "|" & "0" _
                & "|" & fstrFechaSQL(Format(ldtmFecha, "dd/mm/yyyy")) _
                & "|" & fstrFechaSQL(Format(ldtmFecha, "dd/mm/yyyy")) _
                & "|" & "*" _
                & "|" & grdCheque.TextMatrix(llngRenglonSel, cintColIdCheque) _
                & "|" & grdCheque.TextMatrix(llngRenglonSel, cintColMedioPago)
                        
                Set rs = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELCHEQUECAJACHICA")
                blnerror = rs.RecordCount = 0
                        
                If Not blnerror Then
                
                    blnerror = rs!Estatus <> "R"
                    
                    If Not blnerror Then
                    
                        vgstrParametrosSP = grdCheque.TextMatrix(llngRenglonSel, cintColIdAplicacion) & "|" & CStr(llngPersonaGraba) & "|" & CStr(llngNumCorte) & "|" & grdCheque.TextMatrix(llngRenglonSel, cintColIdCheque) & "|'" & grdCheque.TextMatrix(llngRenglonSel, cintColMedioPago) & "'"
                        frsEjecuta_SP vgstrParametrosSP, "SP_PVINSCHEQUECANCELADO"
                        
                        '------------------------------------------------------
                        ' Afectar el corte con negativos:
                        '------------------------------------------------------
                        vgstrParametrosSP = grdCheque.TextMatrix(llngRenglonSel, cintColIdAplicacion) & "|" & IIf(grdCheque.TextMatrix(llngRenglonSel, cintColMedioPago) = "CH", "EC", "ET") & "|" & grdCheque.TextMatrix(llngRenglonSel, cintColIdCorteAfectado)
                        Set rs = frsEjecuta_SP(vgstrParametrosSP, "sp_PvSelFormaDoctoCorte")
                        If rs.RecordCount <> 0 Then
                        
                            Do While Not rs.EOF
                                vgstrParametrosSP = CStr(llngNumCorte) _
                                & "|" & fstrFechaSQL(fdtmServerFecha, fdtmServerHora) _
                                & "|" & grdCheque.TextMatrix(llngRenglonSel, cintColIdAplicacion) _
                                & "|" & IIf(grdCheque.TextMatrix(llngRenglonSel, cintColMedioPago) = "CH", "EC", "ET") _
                                & "|" & CStr(rs!intFormaPago) _
                                & "|" & CStr(rs!mnyCantidadPagada * -1) _
                                & "|" & CStr(rs!MNYTIPOCAMBIO) _
                                & "|" & CStr(rs!intfoliocheque) _
                                & "|" & grdCheque.TextMatrix(llngRenglonSel, cintColIdCorteAfectado)
                                
                                frsEjecuta_SP vgstrParametrosSP, "sp_PvInsDetalleCorte"
                                rs.MoveNext
                            Loop
                        End If
                        '------------------------------------------------------
                        ' Registro de transacciones:
                        '------------------------------------------------------
                        pGuardarLogTransaccion Me.Name, EnmCancelacion, llngPersonaGraba, "CAMBIO CHEQUES", CStr(lngIdCancelacion)
                        '------------------------------------------------------
                        ' Libera corte:
                        '------------------------------------------------------
                        pLiberaCorte llngNumCorte
                    End If
                End If
                
                If blnerror Then
                    EntornoSIHO.ConeccionSIHO.RollbackTrans
                    'La información ha cambiado, consulte de nuevo.
                    MsgBox SIHOMsg(381), vbExclamation + vbOKOnly, "Mensaje"
                Else
                    EntornoSIHO.ConeccionSIHO.CommitTrans
                    'La operación se realizó satisfactoriamente.
                    MsgBox SIHOMsg(420), vbInformation + vbOKOnly, "Mensaje"
                    
                    cmdCargar_Click
                End If
            End If
        End If
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdCancelar_Click"))
End Sub


Private Function fblnCancelacionValida() As Boolean
    On Error GoTo NotificaError

    fblnCancelacionValida = True
    
    '*--*--*---*--*--*---*--*--*---
    ' Que tenga permisos
    '*--*--*---*--*--*---*--*--*---
    fblnCancelacionValida = fblnRevisaPermiso(vglngNumeroLogin, 1647, "E")
    '*--*--*---*--*--*---*--*--*---
    ' Que exista el tipo de cambio cuando se reciben dlls.
    '*--*--*---*--*--*---*--*--*---
    If fblnCancelacionValida Then
        If Val(grdCheque.TextMatrix(llngRenglonSel, cintColbitPesos)) = 0 <> 0 And ldblTipoCambio = 0 Then
            fblnCancelacionValida = False
            'Registre el tipo de cambio del día.
            MsgBox SIHOMsg(335), vbExclamation + vbOKOnly, "Mensaje"
        End If
    End If
    '*--*--*---*--*--*---*--*--*---
    ' Que la contraseña sea válida
    '*--*--*---*--*--*---*--*--*---
    If fblnCancelacionValida Then
        llngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
        fblnCancelacionValida = llngPersonaGraba <> 0
    End If

    Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fblnCancelacionValida"))
End Function

Private Sub cmdCargar_Click()
    On Error GoTo NotificaError

    Dim strEstado As String
    Dim Y As Integer
    
    pLimpiaGrid
    
    Select Case cboEstado.ItemData(cboEstado.ListIndex)
        Case cintIdTodos
            strEstado = "*"
        Case cintIdSinAplicar
            strEstado = "A"
        Case cintIdAplicados
            strEstado = "R"
    End Select
  
    vgstrParametrosSP = CStr(vgintNumeroDepartamento) _
    & "|" & "1" _
    & "|" & fstrFechaSQL(mskFecIni.Text) _
    & "|" & fstrFechaSQL(mskFecFin.Text) _
    & "|" & strEstado _
    & "|" & "-1" _
    & "|" & vlstrMedioPago

    Set rs = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELCHEQUECAJACHICA")
    
    If rs.RecordCount <> 0 Then
        With grdCheque
            .Visible = False
            Do While Not rs.EOF
                .TextMatrix(.Rows - 1, cintColIdCheque) = rs!IdChequeTransferencia
                .TextMatrix(.Rows - 1, cintColFecha) = Format(rs!FechaChequeTransferencia, "dd/mmm/yyyy")
                .TextMatrix(.Rows - 1, cintColBanco) = rs!NombreBanco
                .TextMatrix(.Rows - 1, cintColNumero) = rs!NumeroChequeTransferencia
                .TextMatrix(.Rows - 1, cintColBeneficiario) = rs!NombreBeneficiario
                .TextMatrix(.Rows - 1, cintColCantidad) = FormatCurrency(rs!Cantidad, 2)
                .TextMatrix(.Rows - 1, cintColbitPesos) = rs!BITPESOS
                .TextMatrix(.Rows - 1, cintColMoneda) = rs!Moneda
                .TextMatrix(.Rows - 1, cintColEstatus) = rs!Estatus
                .TextMatrix(.Rows - 1, cIntColEstado) = rs!Estado
                .TextMatrix(.Rows - 1, cintColIdAplicacion) = rs!IdAplicacion
                .TextMatrix(.Rows - 1, cintColIdCorteAfectado) = rs!IdCorteAfectado
                .TextMatrix(.Rows - 1, cintColFechaAplicado) = Format(rs!FechaAplicado, "dd/mmm/yyyy")
                .TextMatrix(.Rows - 1, cintColPersonaAplico) = rs!NombrePersonaAplico
                .TextMatrix(.Rows - 1, cintColMedioPago) = rs!medioPago
                If rs!medioPago = "CH" Then
                    For Y = 1 To .Cols - 1
                        .Row = .Rows - 1
                        .Col = Y
                        'Negro
                        .CellForeColor = &H80000012
                    Next Y
                ElseIf rs!medioPago = "TR" Then
                    For Y = 1 To .Cols - 1
                        .Row = .Rows - 1
                        .Col = Y
                        'Verde
                        .CellForeColor = &H808000
                    Next Y
                End If
                
                .Rows = .Rows + 1
                rs.MoveNext
            Loop
            .Rows = .Rows - 1
            .Visible = True
        End With
        
    Else
        If lblnMensaje Then
            'No existe información con esos parámetros.
            MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje"
        Else
            lblnMensaje = True
        End If
    End If
    
    cmdSeleccionar.Enabled = rs.RecordCount <> 0
    
    rs.Close
    
    grdCheque.Row = 1
    grdCheque.Col = cintColFecha
    grdCheque_Click

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdCargar_Click"))
End Sub

Private Sub cmdGuardar_Click()
    On Error GoTo NotificaError

    Dim rs As New ADODB.Recordset
    Dim intcontador As Integer
    Dim blnerror As Boolean
    Dim lngIdAplicacion As Long
    Dim vllngNumDetalleCorte As Long

    If fblnDatosValidos() Then
    
        EntornoSIHO.ConeccionSIHO.BeginTrans
        
        llngNumCorte = flngNumeroCorte(vgintNumeroDepartamento, vglngNumeroEmpleado, "C")
        
        If llngNumCorte = 0 Then
            EntornoSIHO.ConeccionSIHO.RollbackTrans
            'No se encontró un corte abierto.
            MsgBox SIHOMsg(659), vbExclamation + vbOKOnly, "Mensaje"
        Else
            '------------------------------------------------------------------
            ' Bloquear el corte
            '------------------------------------------------------------------
            llngCorteGrabando = 1
            frsEjecuta_SP CStr(llngNumCorte) & "|Grabando", "Sp_PvUpdEstatusCorte", True, llngCorteGrabando
            If llngCorteGrabando <> 2 Then
                EntornoSIHO.ConeccionSIHO.RollbackTrans
                
                'En este momento se está afectando el corte, espere un momento e intente de nuevo.
                MsgBox SIHOMsg(779), vbExclamation + vbOKOnly, "Mensaje"
            Else
                '------------------------------------------------------
                ' Registrar el cheque como aplicado:
                '------------------------------------------------------
                vgstrParametrosSP = "-1" _
                & "|" & "0" _
                & "|" & fstrFechaSQL(Format(ldtmFecha, "dd/mm/yyyy")) _
                & "|" & fstrFechaSQL(Format(ldtmFecha, "dd/mm/yyyy")) _
                & "|" & "*" _
                & "|" & grdCheque.TextMatrix(llngRenglonSel, cintColIdCheque) _
                & "|" & grdCheque.TextMatrix(llngRenglonSel, cintColMedioPago)
                        
                Set rs = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELCHEQUECAJACHICA")
                blnerror = rs.RecordCount = 0
                        
                If Not blnerror Then
                    blnerror = rs!Estatus <> "A" 'si continúa vigente
                    If Not blnerror Then
                        vgstrParametrosSP = CStr(llngPersonaGraba) & "|" & CStr(vgintNumeroDepartamento) & "|" & CStr(llngNumCorte) & "|" & grdCheque.TextMatrix(llngRenglonSel, cintColIdCheque) & "|'" & grdCheque.TextMatrix(llngRenglonSel, cintColMedioPago) & "'"
                        lngIdAplicacion = 1
                        frsEjecuta_SP vgstrParametrosSP, "SP_PVINSCHEQUEAPLICADO", True, lngIdAplicacion
                        
                        '------------------------------------------------------
                        ' Afectar el corte:
                        '------------------------------------------------------
                        intcontador = 0
                        Do While intcontador <= UBound(aFormasPago(), 1)
                            vgstrParametrosSP = CStr(llngNumCorte) _
                            & "|" & fstrFechaSQL(fdtmServerFecha, fdtmServerHora) _
                            & "|" & CStr(lngIdAplicacion) _
                            & "|" & IIf(grdCheque.TextMatrix(llngRenglonSel, cintColMedioPago) = "CH", "EC", "ET") _
                            & "|" & CStr(aFormasPago(intcontador).vlintNumFormaPago) _
                            & "|" & CStr(IIf(aFormasPago(intcontador).vldblTipoCambio = 0, aFormasPago(intcontador).vldblCantidad, aFormasPago(intcontador).vldblDolares)) _
                            & "|" & CStr(aFormasPago(intcontador).vldblTipoCambio) _
                            & "|" & IIf(Trim(aFormasPago(intcontador).vlstrFolio) = "", "0", Trim(aFormasPago(intcontador).vlstrFolio)) _
                            & "|" & CStr(llngNumCorte)
                        
                            frsEjecuta_SP vgstrParametrosSP, "sp_PvInsDetalleCorte"
                            
                            vllngNumDetalleCorte = flngObtieneIdentity("SEC_PVDETALLECORTE", 0)
                    
                            If Not aFormasPago(intcontador).vlbolEsCredito Then
                                If Trim(aFormasPago(intcontador).vlstrRFC) <> "" And Trim(aFormasPago(intcontador).vlstrBancoSAT) <> "" Then
                                    frsEjecuta_SP llngNumCorte & "|" & vllngNumDetalleCorte & "|'" & Trim(aFormasPago(intcontador).vlstrRFC) & "'|'" & Trim(aFormasPago(intcontador).vlstrBancoSAT) & "'|'" & Trim(aFormasPago(intcontador).vlstrCuentaBancaria) & "'|'" & IIf(Trim(aFormasPago(intcontador).vlstrCuentaBancaria) = "", Null, fstrFechaSQL(Trim(aFormasPago(intcontador).vldtmFecha))) & "'|'" & Trim(aFormasPago(intcontador).vlstrBancoExtranjero) & "'", "SP_PVINSCORTECHEQUETRANSCTA"
                                End If
                            End If
                            
                            intcontador = intcontador + 1
                        Loop
                        '------------------------------------------------------
                        ' Registro de transacciones:
                        '------------------------------------------------------
                        pGuardarLogTransaccion Me.Name, EnmGrabar, llngPersonaGraba, "CAMBIO CHEQUES/TRANSFERENCIAS", CStr(lngIdAplicacion)
                        '------------------------------------------------------
                        ' Libera corte:
                        '------------------------------------------------------
                        pLiberaCorte llngNumCorte
                    End If
                End If
                
                If blnerror Then
                    EntornoSIHO.ConeccionSIHO.RollbackTrans
                    'La información ha cambiado, consulte de nuevo.
                    MsgBox SIHOMsg(381), vbExclamation + vbOKOnly, "Mensaje"
                Else
                    EntornoSIHO.ConeccionSIHO.CommitTrans
                    'La operación se realizó satisfactoriamente.
                    MsgBox SIHOMsg(420), vbInformation + vbOKOnly, "Mensaje"
                    
                    cmdCargar_Click
                End If
                
            End If
        End If
        
    
    End If


    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdGuardar_Click"))
End Sub

Private Function fblnDatosValidos() As Boolean
    On Error GoTo NotificaError
    Dim dblCantidad As Double

    fblnDatosValidos = True
    
    '*--*--*---*--*--*---*--*--*---
    ' Que tenga permisos
    '*--*--*---*--*--*---*--*--*---
    fblnDatosValidos = fblnRevisaPermiso(vglngNumeroLogin, 1647, "E")
    '*--*--*---*--*--*---*--*--*---
    ' Que exista el tipo de cambio cuando se reciben dlls.
    '*--*--*---*--*--*---*--*--*---
    If fblnDatosValidos Then
        If Val(grdCheque.TextMatrix(llngRenglonSel, cintColbitPesos)) = 0 And ldblTipoCambio = 0 Then
            fblnDatosValidos = False
            'Registre el tipo de cambio del día.
            MsgBox SIHOMsg(335), vbExclamation + vbOKOnly, "Mensaje"
        End If
    End If
    '*--*--*---*--*--*---*--*--*---
    ' Que la contraseña sea válida
    '*--*--*---*--*--*---*--*--*---
    If fblnDatosValidos Then
        llngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
        fblnDatosValidos = llngPersonaGraba <> 0
    End If
    '*-*-*-
    ' Que se hayan seleccionado formas de pago
    '*-*-*-
    If fblnDatosValidos Then
        If Val(grdCheque.TextMatrix(llngRenglonSel, cintColbitPesos)) = 0 Then
            dblCantidad = Val(Format(grdCheque.TextMatrix(llngRenglonSel, cintColCantidad), cstrFormato)) * ldblTipoCambio
        Else
            dblCantidad = Val(Format(grdCheque.TextMatrix(llngRenglonSel, cintColCantidad), cstrFormato))
        End If
        fblnDatosValidos = fblnFormasPagoPos(aFormasPago(), dblCantidad, True, ldblTipoCambio, False, 0, "", Trim(Replace(Replace(Replace(vgstrRfCCH, "-", ""), "_", ""), " ", "")))
    End If

    Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fblnDatosValidos"))
End Function



Private Sub cmdSeleccionar_Click()
    On Error GoTo NotificaError

    
    grdCheque_DblClick

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdSeleccionar_Click"))
End Sub

Private Sub Form_Activate()
    
    Dim lngMensaje As Long

    lngMensaje = flngCorteValido(vgintNumeroDepartamento, vglngNumeroEmpleado, "C")

    If lngMensaje <> 0 Then
        'Cierre el corte actual.
        MsgBox SIHOMsg(Str(lngMensaje)), vbOKOnly + vbInformation, "Mensaje"
        Unload Me
    End If


End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError


    If KeyAscii = 27 Then
        Unload Me
    Else
        If KeyAscii = 13 Then
            SendKeys vbTab
        End If
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_KeyPress"))
End Sub

Private Sub Form_Load()
    On Error GoTo NotificaError


    Me.Icon = frmMenuPrincipal.Icon

    lblDepartamento.Caption = vgstrNombreDepartamento

    ldtmFecha = fdtmServerFecha
    ldblTipoCambio = fdblTipoCambio(ldtmFecha, "O")

    mskFecIni.Mask = ""
    mskFecIni.Text = ldtmFecha
    mskFecIni.Mask = "##/##/####"

    mskFecFin.Mask = ""
    mskFecFin.Text = ldtmFecha
    mskFecFin.Mask = "##/##/####"

    cboEstado.AddItem "<Todos>", 0
    cboEstado.ItemData(cboEstado.newIndex) = cintIdTodos
    cboEstado.AddItem "Sin aplicar", 1
    cboEstado.ItemData(cboEstado.newIndex) = cintIdSinAplicar
    cboEstado.AddItem "Aplicados", 2
    cboEstado.ItemData(cboEstado.newIndex) = cintIdAplicados
    cboEstado.ListIndex = 0
    
    lblnMensaje = False
    vlstrMedioPago = "T"
    cmdCargar_Click

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_Load"))
End Sub

Private Sub grdCheque_Click()
    On Error GoTo NotificaError


    With grdCheque
        If Val(.TextMatrix(.Row, cintColIdCheque)) <> 0 Then
            lblNumeroCorte.Caption = IIf(Val(.TextMatrix(.Row, cintColIdCorteAfectado)) = 0, "", .TextMatrix(.Row, cintColIdCorteAfectado))
            lblFechaAplicacion.Caption = Format(.TextMatrix(.Row, cintColFechaAplicado), "dd/mmm/yyyy")
            lbPersonaAplico.Caption = .TextMatrix(.Row, cintColPersonaAplico)
        End If
        
        fraChequeAplicado.Enabled = .TextMatrix(.Row, cintColEstatus) = "R"
        
        lblTituloCorte.Enabled = .TextMatrix(.Row, cintColEstatus) = "R"
        lblTituloFecha.Enabled = .TextMatrix(.Row, cintColEstatus) = "R"
        lblTituloPersonaAplico.Enabled = .TextMatrix(.Row, cintColEstatus) = "R"
        
    End With
    
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":grdCheque_Click"))
End Sub

Private Sub grdCheque_DblClick()
    On Error GoTo NotificaError

    
    If Val(grdCheque.TextMatrix(grdCheque.Row, cintColIdCheque)) <> 0 Then
        grdCheque.TextMatrix(grdCheque.Row, 0) = IIf(Trim(grdCheque.TextMatrix(grdCheque.Row, 0)) = "", "*", "")
        
        If Trim(grdCheque.TextMatrix(grdCheque.Row, 0)) = "*" Then
            
            llngRenglonSel = grdCheque.Row
            pLimpiaSeleccion grdCheque.Row
        
        Else
            llngRenglonSel = 0
        End If
            
        cmdGuardar.Enabled = Trim(grdCheque.TextMatrix(grdCheque.Row, 0)) = "*" And Trim(grdCheque.TextMatrix(grdCheque.Row, cintColEstatus)) = "A"
        cmdCancelar.Enabled = Trim(grdCheque.TextMatrix(grdCheque.Row, 0)) = "*" And Trim(grdCheque.TextMatrix(grdCheque.Row, cintColEstatus)) = "R"
    End If
    
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":grdCheque_DblClick"))
End Sub

Private Sub pLimpiaSeleccion(intRenglon As Integer)
    On Error GoTo NotificaError

    Dim intcontador As Integer
    
    For intcontador = 1 To grdCheque.Rows - 1
        If intcontador <> intRenglon Then
            grdCheque.TextMatrix(intcontador, 0) = ""
        End If
    Next intcontador

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pLimpiaSeleccion"))
End Sub



Private Sub grdCheque_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError


    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
        grdCheque_Click
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":grdCheque_KeyDown"))
End Sub

Private Sub lblTituloTotal_Click()
    On Error GoTo NotificaError


    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":lblTituloTotal_Click"))
End Sub

Private Sub mskFecFin_Change()
    On Error GoTo NotificaError


    cmdCargar.Enabled = IsDate(mskFecFin.Text)

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFecFin_Change"))
End Sub

Private Sub mskFecFin_GotFocus()
    On Error GoTo NotificaError


    pSelMkTexto mskFecFin

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFecFin_GotFocus"))
End Sub

Private Sub mskFecFin_LostFocus()
    On Error GoTo NotificaError


    If Not IsDate(mskFecFin.Text) Then
        mskFecFin.Mask = ""
        mskFecFin.Text = ldtmFecha
        mskFecFin.Mask = "##/##/####"
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFecFin_LostFocus"))
End Sub

Private Sub mskFecIni_Change()
    On Error GoTo NotificaError

    
    cmdCargar.Enabled = IsDate(mskFecIni.Text)

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFecIni_Change"))
End Sub

Private Sub mskFecIni_GotFocus()
    On Error GoTo NotificaError


    pSelMkTexto mskFecIni
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFecIni_GotFocus"))
End Sub

Private Sub mskFecIni_LostFocus()
    On Error GoTo NotificaError


    If Not IsDate(mskFecIni.Text) Then
        mskFecIni.Mask = ""
        mskFecIni.Text = ldtmFecha
        mskFecIni.Mask = "##/##/####"
    End If


    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFecIni_LostFocus"))
End Sub

Private Sub optMediopago_Click(Index As Integer)

    If optMediopago(0).Value Then
        vlstrMedioPago = "T" 'Mostrar todos los pagos
    ElseIf optMediopago(1).Value Then
        vlstrMedioPago = "CH"  'Mostrar solo los pagos con cheques
    ElseIf optMediopago(2).Value Then
        vlstrMedioPago = "TR"  'Mostrar solo los pagos con transferencia
    End If
    
End Sub


