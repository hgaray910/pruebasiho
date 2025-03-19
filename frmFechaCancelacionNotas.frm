VERSION 5.00
Begin VB.Form frmFechaCancelacionNotas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fecha de cancelación"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3480
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   3480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   675
      Left            =   1440
      TabIndex        =   4
      Top             =   1680
      Width           =   585
      Begin VB.CommandButton cmdGrabarRegistro 
         Height          =   495
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmFechaCancelacionNotas.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Guardar el registro"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.Frame fraFechaCancelacion 
      Caption         =   "Aplicar cancelación de documentos"
      Height          =   1425
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   3015
      Begin VB.OptionButton OptCancelaTipoActual 
         Caption         =   "En la fecha actual"
         Height          =   255
         Left            =   280
         TabIndex        =   3
         Top             =   840
         Width           =   2535
      End
      Begin VB.OptionButton OptCancelaTipoDocu 
         Caption         =   "En la fecha del documento"
         Height          =   255
         Left            =   280
         TabIndex        =   0
         Top             =   480
         Value           =   -1  'True
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmFechaCancelacionNotas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************** NOTA *********************************************
Option Explicit

Const cintTipoDocumento = 8 'Nota de crédito

Private Type ResumenFactura
    vlstrFolioFactura As String
    vldblSubtotal As Double
    vldblDescuento As Double
    vldblIVA As Double
    vldtmFecha As Date
End Type

'Para el detalle de la nota:
Const vlintColFactura = 1
Const vlintColDescripcionConcepto = 2
Const vlintColCantidad = 3
Const vlintColDescuento = 4
Const vlintColIVA = 5
Const vlintColTotal = 6
Const vlintColCveConcepto = 7
Const vlintColCtaIngresos = 8
Const vlintColCtaDescuentos = 9
Const vlintColCtaIVA = 10
Const vlintColTipoCargo = 11
Const vlintColTipoNotaFARCDetalle = 12

'Para la búsqueda de notas:
Const vlintColFechaNota = 1
Const vlintColFolioNota = 2
Const vlintColTipoNota = 3
Const vlintColNumCliente = 4
Const vlintColNombreCliente = 5
Const vlintColEstadoNota = 6
Const vlintColMotivoNota = 7
Const vlintColDomicilioCliente = 8
Const vlintColRFCCliente = 9
Const vlintColSubtotalNota = 10
Const vlintColDescuentoNota = 11
Const vlintColIVANota = 12
Const vlintColTotalNota = 13
Const vlintColchrTipo = 14
Const vlintColchrEstatus = 15
Const vlintColintNumPoliza = 16
Const vlintColdtmFechaRegistro = 17
Const vlintColCuentaContable = 18
Const vlintColTipoNotaFACR = 19
Const vlintColIdNota = 20

Dim vlstrFormato As String 'Formato de moneda
Dim vlintBitSaldarCuentas As Long               'Variable que indica el valor del bit pvConceptoFacturacion.BitSaldarCuentas, que nos dice si la cuenta del ingreso se salda con la del descuento
Dim vlblnCuentaIngresoSaldada As Boolean        'Variable que indica si la cuenta del ingreso fue saldada con la cuenta del descuento
Public vlblnActivaMotivo As Boolean

Private Function fintErrorCancelarNota(vllngClaveCliente As Long, vlstrFolio As String) As Integer
'----------------------------------------------------------------------------------------------------------------------'
' Función que revisa que la nota de cargo no estés incluida en un paquete de cobranza o que no tenga pagos registrados '
'----------------------------------------------------------------------------------------------------------------------'
    Dim rs As New ADODB.Recordset
    Dim rsPagos As New ADODB.Recordset
    
    fintErrorCancelarNota = 0
    
    'Que el o los créditos de la factura no tengan pagos registrados
    vgstrParametrosSP = fstrFechaSQL(fdtmServerFecha) & _
                        "|" & fstrFechaSQL(fdtmServerFecha) & _
                        "|" & vllngClaveCliente & _
                        "|" & "0" & _
                        "|" & "CA" & _
                        "|" & "0" & _
                        "|" & vlstrFolio & _
                        "|" & "0" & _
                        "|" & "0" & _
                        "|" & "*" & _
                        "|" & "0"
    Set rs = frsEjecuta_SP(vgstrParametrosSP, "sp_CcSelCredito")
    If rs.RecordCount <> 0 Then
        If IsDate(rs!fechaEnvio) Then
            'No se puede cancelar el documento, los créditos fueron incluídos en un paquete de cobranza.
            fintErrorCancelarNota = 718
        Else
            vgstrParametrosSP = Str(rs!Movimiento) & "|" & "0" & "|" & fstrFechaSQL(fdtmServerFecha) & "|" & "I"
            Set rsPagos = frsEjecuta_SP(vgstrParametrosSP, "sp_CcSelPagosCredito")
            If rsPagos.RecordCount <> 0 Then
                'No se puede cancelar el documento  el crédito tiene pagos registrados.
                fintErrorCancelarNota = 368
            End If
            rsPagos.Close
        End If
    Else
        '¡La información no existe!
        fintErrorCancelarNota = 12
        Exit Function
    End If
    rs.Close
End Function
Public Sub pCancelaNota(vlstrFolioNota As String, vlstrCliente As String, strTipo As String, vlblnCancelaSiHO As Boolean, vlblnMuestramensaje As Boolean, Optional blnNoPersonagraba As Boolean = False, Optional PG As Long)
On Error GoTo NotificaError

    Dim rsNotasFacturas As New ADODB.Recordset
    Dim vllngPersonaGraba As Long
    Dim vllngNumPoliza  As Long
    Dim vllngResultado As Long
    Dim vllngContador As Long
    Dim vllngDetallePoliza As Long
    Dim vlstrEstadoNota As String
    Dim vlstrSentencia As String
    Dim vldblCantidad As Double
    Dim vldblDescuento As Double
    Dim vldblIVA As Double
    Dim dblTotalIVA As Double 'Para hacer el movimiento a IVA no cobrado
    Dim intMensaje As Integer 'Mensaje que regresa la función fintErrorCancelarNota()
    Dim intNumeroCuenta As Long
    Dim rsDetalleNotaElectronica As ADODB.Recordset
    Dim vlstrTipoCancelacion As String
    Dim lngFacturaPagada As Long
    Dim vlstrFechaCancelacion As String
    Dim rsDepartamentoConcepto As ADODB.Recordset
    'Dim dblPorcentajeNota As Double
    'Dim rsDetallePoliza As ADODB.Recordset
    Dim vllngNumPolizaNota As Long
    Dim rsCnDetallePoliza As New ADODB.Recordset
    Dim rsDatosPoliza As New ADODB.Recordset
    Dim vldblTipoCambio As Double
    
    Dim rsNota As New Recordset
    Dim vllngconsecutivoNota As Long
    Dim vlstrTipoNota As String
    Dim vldtmFechaNota As String
    Dim vllngPoliza As Long
    Dim vllngClaveCliente As Long
    Dim rsConceptoPolizaCancelada As New ADODB.Recordset
    Dim vlObtieneConceptoPoliza As String
    
    vlstrTipoCancelacion = Trim(strTipo)

    '|-----------------------------------------------------------------------------------------------------------------------
    '|  Consulta toda la información que necesita de la nota
    '|-----------------------------------------------------------------------------------------------------------------------
    vlstrSentencia = "Select intConsecutivo, chrTipo, dtmFecha, intNumPoliza, intCliente " & _
                     "  From CcNota " & _
                     " Where CcNota.chrFolioNota = '" & vlstrFolioNota & "'"
    Set rsNota = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
    
    If rsNota.RecordCount > 0 Then
        vllngconsecutivoNota = rsNota!intConsecutivo
        vlstrTipoNota = rsNota!chrTipo
        vldtmFechaNota = rsNota!dtmfecha
        vllngPoliza = IIf(IsNull(rsNota!intNumPoliza), 0, rsNota!intNumPoliza)
        vllngClaveCliente = IIf(IsNull(rsNota!intCliente), 0, rsNota!intCliente)
    Else
        '| No existe la nota
        Exit Sub
    End If


    ' Inicialización del formato
    vlstrFormato = "###############.00"
    
    If Not fblnRevisaPermiso(vglngNumeroLogin, IIf(cgstrModulo = "PV", 2296, IIf(cgstrModulo = "CC", 634, -1)), "E") Then Exit Sub
    
    If blnNoPersonagraba = False Then
        vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
        If vllngPersonaGraba = 0 Then
           frmNotas.blnCancelaNota = False
           Exit Sub
        End If
    Else
        vllngPersonaGraba = PG
    End If
    If vlblnActivaMotivo Then
        frmMotivosCancelacion.blnActivaUUID = False
        frmMotivosCancelacion.Show vbModal, Me
        If vgMotivoCancelacion = "" Then Exit Sub
        vlblnActivaMotivo = False
    End If
        
    
    '-----------------------------------------------------------------------------------------------------------------------
    'Cancelacion ante la SAT sólo cuando es una nota electronica, no se por que se válida de esta forma pero asi estaba ya
    '-----------------------------------------------------------------------------------------------------------------------
    Set rsDetalleNotaElectronica = frsRegresaRs("SELECT * FROM GnComprobanteFiscalDigital INNER JOIN CCnota ON GnComprobanteFiscalDigital.INTCOMPROBANTE = CCNota.INTCONSECUTIVO AND GnComprobanteFiscalDigital.CHRTIPOCOMPROBANTE = CCNota.CHRTIPO WHERE CCNota.ChrFolioNota = '" & Trim(vlstrFolioNota) & "'")
    If rsDetalleNotaElectronica.RecordCount <> 0 Then
           
           If IIf(IsNull(rsDetalleNotaElectronica!VCHUUID), "", rsDetalleNotaElectronica!VCHUUID) <> "" Then
                '----------------------------------
                'Cancelar el CFDi por medio del PAC
                '----------------------------------
                If Not fblnCancelaCFDi(vllngconsecutivoNota, vlstrTipoNota) Then
                   frmNotas.blnCancelaNota = False
                   If vlstrMensajeErrorCancelacionCFDi <> "" Then MsgBox vlstrMensajeErrorCancelacionCFDi, vbOKOnly + vbCritical, "Mensaje"
                   Exit Sub
                End If
           End If
           
           If (frmFechaCancelacionNotas.OptCancelaTipoDocu.Value = True And vlstrTipoCancelacion = "ELEGIR") Or vlstrTipoCancelacion = "DOCUMENTO" Then 'Si se va a cancelar con la fecha del documento (Método actual)....
              vgstrParametrosSP = vllngconsecutivoNota & "|" & vlstrTipoNota & "|" & fstrFechaSQL(vldtmFechaNota, "00:00:00", False) & "|'" & vgMotivoCancelacion & "'"
           ElseIf (frmFechaCancelacionNotas.OptCancelaTipoActual.Value = True And vlstrTipoCancelacion = "ELEGIR") Or vlstrTipoCancelacion = "ACTUAL" Then 'Si se va a cancelar con la fecha actual (Método viejo)....
                vgstrParametrosSP = vllngconsecutivoNota & "|" & vlstrTipoNota & "|" & fstrFechaSQL(fdtmServerFecha, fdtmServerHora) & "|'" & vgMotivoCancelacion & "'"
           End If
            
           frsEjecuta_SP vgstrParametrosSP, "SP_GNUPDCANCELACOMPROBANTEFIS"
    End If
    
    If vlblnCancelaSiHO Then
        '-----------------------------------------------------------------------------------------------------------------------
        '|  Cancelación en el SiHO
        '-----------------------------------------------------------------------------------------------------------------------
        
        EntornoSIHO.ConeccionSIHO.BeginTrans
        
        lngFacturaPagada = 1
        ' Función que valida el número de cliente ' no cambia nada en la base de datos
        frsEjecuta_SP CStr(vllngconsecutivoNota), "FN_CCSELFACTURAPAGADA", True, lngFacturaPagada
        
        ' Se valida el estado actual de la nota
        vlstrEstadoNota = frsRegresaRs("SELECT chrEstatus FROM CcNota WHERE intConsecutivo = " & vllngconsecutivoNota).Fields(0)
        If vlstrEstadoNota = "C" Then
            EntornoSIHO.ConeccionSIHO.RollbackTrans
            MsgBox "Hay una reciente versión del estado de la nota, favor de cargar nuevamente los datos para actualizar la información.", vbOKOnly + vbExclamation, "Mensaje"
            frmNotas.blnCancelaNota = False
            Exit Sub
        End If
        
        '---------------------------------------------------------------------------------'
        ' Que no tenga pagos registrados y que no esté incluida en un paquete de cobranza '
        '---------------------------------------------------------------------------------'
        If vlstrTipoNota = "CA" Then
            intMensaje = fintErrorCancelarNota(vllngClaveCliente, Trim(vlstrFolioNota))
            If intMensaje <> 0 Then
                EntornoSIHO.ConeccionSIHO.RollbackTrans
                MsgBox SIHOMsg(intMensaje), vbOKOnly + vbExclamation, "Mensaje"
                frmNotas.blnCancelaNota = False
                Exit Sub
            End If
        End If
        
        '--------------------------------------------------------------------'
        ' Revisar que no se esté haciendo un cierre contable en este momento '
        '--------------------------------------------------------------------'
        vllngResultado = 1
        vgstrParametrosSP = vgintClaveEmpresaContable & "|" & "Grabando poliza"
        frsEjecuta_SP vgstrParametrosSP, "Sp_CnUpdEstatusCierre", True, vllngResultado
        If vllngResultado <> 1 Then
            EntornoSIHO.ConeccionSIHO.RollbackTrans
            'En este momento se está realizando un cierre contable, espere un momento e intente de nuevo.
            MsgBox SIHOMsg(714), vbOKOnly + vbInformation, "Mensaje"
            frmNotas.blnCancelaNota = False
            Exit Sub
        End If
        
        '----------------------------------------------'
        ' Revisar que el periodo contable esté abierto '
        '----------------------------------------------'
        ' Valida el parametro para aplicar la fecha de cancelación en la póliza
        If (frmFechaCancelacionNotas.OptCancelaTipoDocu.Value = True And vlstrTipoCancelacion = "ELEGIR") Or vlstrTipoCancelacion = "DOCUMENTO" Then 'Si se va a cancelar con la fecha del documento (Método actual)....
            vlstrFechaCancelacion = fstrFechaSQL(CDate(vldtmFechaNota), "00:00:00", True)
        ElseIf (frmFechaCancelacionNotas.OptCancelaTipoActual.Value = True And vlstrTipoCancelacion = "ELEGIR") Or vlstrTipoCancelacion = "ACTUAL" Then 'Si se va a cancelar con la fecha actual (Método viejo)....
            vlstrFechaCancelacion = fstrFechaSQL(fdtmServerFecha, fdtmServerHora, True)
        End If
    '    If fblnPeriodoCerrado(vgintClaveEmpresaContable, Year(CDate(vldtmFechaNota)), Month(CDate(vldtmFechaNota))) Then
        If fblnPeriodoCerrado(vgintClaveEmpresaContable, Year(CDate(vlstrFechaCancelacion)), Month(CDate(vlstrFechaCancelacion))) Then
            EntornoSIHO.ConeccionSIHO.RollbackTrans
            'El periodo contable esta cerrado.
            MsgBox SIHOMsg(209), vbOKOnly + vbInformation, "Mensaje"
            frmNotas.blnCancelaNota = False
            Exit Sub
        End If
        
        '---------------------------------------------'
        ' Actualizar a cancelado el estado de la nota '
        '---------------------------------------------'
        vlstrSentencia = "UPDATE CcNota SET chrEstatus = 'C', intPersonaBorra = " & Str(vllngPersonaGraba) & " WHERE intConsecutivo = " & vllngconsecutivoNota
        pEjecutaSentencia vlstrSentencia
        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        '+++++++++++++++++++++++++ INICIO +++++++++++++++++++++++++++
        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        '-----------------------------------------------------------'
        '   Registra el recibo cancelado en documentos cancelados   '
        '-----------------------------------------------------------'
        If (frmFechaCancelacionNotas.OptCancelaTipoDocu.Value = True And vlstrTipoCancelacion = "ELEGIR") Or vlstrTipoCancelacion = "DOCUMENTO" Then 'Si se va a cancelar con la fecha del documento (Método actual)....
            vgstrParametrosSP = IIf(vlstrTipoNota = "CR", "NC", "NA") & "|" & Trim(vlstrFolioNota) & "|" & Str(vgintNumeroDepartamento) & "|" & Str(vllngPersonaGraba) & "|" & fstrFechaSQL(vldtmFechaNota, "00:00:00", False)
        ElseIf (frmFechaCancelacionNotas.OptCancelaTipoActual.Value = True And vlstrTipoCancelacion = "ELEGIR") Or vlstrTipoCancelacion = "ACTUAL" Then 'Si se va a cancelar con la fecha actual (Método viejo)....
            vgstrParametrosSP = IIf(vlstrTipoNota = "CR", "NC", "NA") & "|" & Trim(vlstrFolioNota) & "|" & Str(vgintNumeroDepartamento) & "|" & Str(vllngPersonaGraba) & "|" & fstrFechaSQL(fdtmServerFecha, fdtmServerHora)
        End If
        
        frsEjecuta_SP vgstrParametrosSP, "SP_CCUPDCANCELANOTA", True
        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        '+++++++++++++++++++++++++ FIN ++++++++++++++++++++++++++++++
        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                
        '--------------------------'
        ' Hacer una póliza inversa '
        '--------------------------'
        
        vlObtieneConceptoPoliza = ""
        
        Set rsConceptoPolizaCancelada = frsRegresaRs("select vchconceptopoliza from cnpoliza Where intnumeropoliza = (select INTNUMPOLIZA from ccnota Where chrfolionota = '" & frmNotas.lblFolio & "')", adLockOptimistic, adOpenDynamic)
        If rsConceptoPolizaCancelada.RecordCount > 0 Then vlObtieneConceptoPoliza = rsConceptoPolizaCancelada!vchConceptoPoliza
       
        If CDate(vldtmFechaNota) <= fdtmServerFecha Then
            
            '|----------------------------------------------------------------------
            '|             N O T A S     A U T O M Á T I C A S
            '|----------------------------------------------------------------------
            If fblnNotaAutomatica(vlstrFolioNota) Then
                ' Valida el parametro para aplicar la fecha de cancelación en la póliza
                If (frmFechaCancelacionNotas.OptCancelaTipoDocu.Value = True And vlstrTipoCancelacion = "ELEGIR") Or vlstrTipoCancelacion = "DOCUMENTO" Then 'Si se va a cancelar con la fecha del documento (Método actual)....
                    vllngNumPoliza = flngInsertarPoliza(CDate(vldtmFechaNota), "D", "CANCELACION DE " & vlObtieneConceptoPoliza, vllngPersonaGraba)
                    vlstrFechaCancelacion = fstrFechaSQL(CDate(vldtmFechaNota), "00:00:00")
                ElseIf (frmFechaCancelacionNotas.OptCancelaTipoActual.Value = True And vlstrTipoCancelacion = "ELEGIR") Or vlstrTipoCancelacion = "ACTUAL" Then 'Si se va a cancelar con la fecha actual (Método viejo)....
                    vllngNumPoliza = flngInsertarPoliza(fdtmServerFecha, "D", "CANCELACION DE " & vlObtieneConceptoPoliza, vllngPersonaGraba)
                    vlstrFechaCancelacion = fstrFechaSQL(fdtmServerFecha, fdtmServerHora)
                End If
                
                '***** (CR) Guardar información de la póliza fuera del corte para reporte en Corte de caja *****'
                vgstrParametrosSP = flngNumeroCorte(vgintNumeroDepartamento, vglngNumeroEmpleado, "P") & "|" & vlstrFechaCancelacion & "|" & IIf(vlstrTipoNota = "CR", "NC", "NA") & "|" & CStr(vllngNumPoliza) & "|" & CStr(vllngPersonaGraba) & "|" & CStr(vgintNumeroDepartamento)
                frsEjecuta_SP vgstrParametrosSP, "SP_PVINSMOVIMIENTOFUERACORTE"
                '***********************************************************************************************'
                vllngNumPolizaNota = vllngPoliza
                
                Set rsDatosPoliza = frsRegresaRs("Select mnyCantidad ImporteIngreso, intCuentaDescuento CtaIngreso, mnyIva ImporteIvaCobrado, intCuentaIVA CtaIvaCobrado, (mnyCantidad + mnyIva) ImporteCtaPteNotaCredPac, intCuentaIngreso CtaPteNotaCreditoPaciente  From CcNotaDetalle Where CcNotaDetalle.INTCONSECUTIVO = " & vllngconsecutivoNota)
                If rsDatosPoliza.RecordCount <> 0 Then
                    Set rsCnDetallePoliza = frsRegresaRs("select * from cnDetallePoliza where intNumeroPoliza=-1", adLockOptimistic, adOpenDynamic)
                    With rsCnDetallePoliza
                        rsDatosPoliza.MoveFirst
                        Do While Not rsDatosPoliza.EOF
                            '|  Asiento contable de la  cuenta del ingreso
                            .AddNew
                            !intNumeroPoliza = vllngNumPoliza
                            !intNumeroCuenta = rsDatosPoliza!CtaIngreso
                            !bitNaturalezaMovimiento = 0
                            !mnyCantidadMovimiento = rsDatosPoliza!ImporteIngreso
                            !vchConcepto = vlstrFolioNota
                            !vchReferencia = " "
                            .Update
                            
                            '|  Asiento contable de la cuenta de IVA Cobrado
                            If rsDatosPoliza!ImporteIvaCobrado > 0 Then
                                .AddNew
                                !intNumeroPoliza = vllngNumPoliza
                                !intNumeroCuenta = rsDatosPoliza!CtaIvaCobrado
                                !bitNaturalezaMovimiento = 0
                                !mnyCantidadMovimiento = rsDatosPoliza!ImporteIvaCobrado
                                !vchConcepto = vlstrFolioNota
                                !vchReferencia = " "
                                .Update
                            End If
                            
                            '|  Asiento contable de la cuenta puente para notas de crédito del paciente
                            .AddNew
                            !intNumeroPoliza = vllngNumPoliza
                            !intNumeroCuenta = rsDatosPoliza!CtaPteNotaCreditoPaciente
                            !bitNaturalezaMovimiento = 1
                            !mnyCantidadMovimiento = rsDatosPoliza!ImporteCtaPteNotaCredPac
                            !vchConcepto = vlstrFolioNota
                            !vchReferencia = " "
                            .Update
                            
                            rsDatosPoliza.MoveNext
                        Loop
                        rsDatosPoliza.Close
                    End With
                    rsCnDetallePoliza.Close
                End If
            Else
                ' Valida el parametro para aplicar la fecha de cancelación en la póliza
                If (frmFechaCancelacionNotas.OptCancelaTipoDocu.Value = True And vlstrTipoCancelacion = "ELEGIR") Or vlstrTipoCancelacion = "DOCUMENTO" Then 'Si se va a cancelar con la fecha del documento (Método actual)....
                    vllngNumPoliza = flngInsertarPoliza(CDate(vldtmFechaNota), "D", "CANCELACION DE " & vlObtieneConceptoPoliza, vllngPersonaGraba)
                    vlstrFechaCancelacion = fstrFechaSQL(CDate(vldtmFechaNota), "00:00:00")
                ElseIf (frmFechaCancelacionNotas.OptCancelaTipoActual.Value = True And vlstrTipoCancelacion = "ELEGIR") Or vlstrTipoCancelacion = "ACTUAL" Then 'Si se va a cancelar con la fecha actual (Método viejo)....
                    vllngNumPoliza = flngInsertarPoliza(fdtmServerFecha, "D", "CANCELACION DE " & vlObtieneConceptoPoliza, vllngPersonaGraba)
                    vlstrFechaCancelacion = fstrFechaSQL(fdtmServerFecha, fdtmServerHora)
                End If
                
                '***** (CR) Guardar información de la póliza fuera del corte para reporte en Corte de caja *****'
                vgstrParametrosSP = flngNumeroCorte(vgintNumeroDepartamento, vglngNumeroEmpleado, "P") & "|" & vlstrFechaCancelacion & "|" & IIf(vlstrTipoNota = "CR", "NC", "NA") & "|" & CStr(vllngNumPoliza) & "|" & CStr(vllngPersonaGraba) & "|" & CStr(vgintNumeroDepartamento)
                frsEjecuta_SP vgstrParametrosSP, "SP_PVINSMOVIMIENTOFUERACORTE"
                '***********************************************************************************************'
                vllngNumPolizaNota = vllngPoliza
                
                Set rsDatosPoliza = frsRegresaRs("select * from cndetallepoliza where intnumeropoliza = " & vllngNumPolizaNota & " order by intnumeroregistro")
                If rsDatosPoliza.RecordCount <> 0 Then
                    Set rsCnDetallePoliza = frsRegresaRs("select * from cnDetallePoliza where intNumeroPoliza=-1", adLockOptimistic, adOpenDynamic)
                    With rsCnDetallePoliza
                        rsDatosPoliza.MoveFirst
                        Do While Not rsDatosPoliza.EOF
                            .AddNew
                            !intNumeroPoliza = vllngNumPoliza
                            !intNumeroCuenta = rsDatosPoliza!intNumeroCuenta
                            !bitNaturalezaMovimiento = IIf(rsDatosPoliza!bitNaturalezaMovimiento, 0, 1)
                            !mnyCantidadMovimiento = rsDatosPoliza!mnyCantidadMovimiento
                            !vchConcepto = rsDatosPoliza!vchConcepto
                            !vchReferencia = rsDatosPoliza!vchReferencia
                            .Update
                            rsDatosPoliza.MoveNext
                        Loop
                        rsDatosPoliza.Close
                    End With
                    rsCnDetallePoliza.Close
                End If
            End If
        Else
            vllngNumPoliza = vllngPoliza
            pCancelarPoliza vllngNumPoliza, "CANCELACION DE NOTA DE " & IIf(vlstrTipoNota = "CR", "CREDITO ", "CARGO ") & Trim(vlstrFolioNota) & " (REUTILIZAR POLIZA) "
        End If
        
        vlstrSentencia = "UPDATE CcNota SET intNumPolizaCancelacion = " & Str(vllngNumPoliza) & " WHERE intConsecutivo = " & vllngconsecutivoNota
        pEjecutaSentencia vlstrSentencia
        
        '------------------------------------'
        ' Afectar los movimientos de crédito '
        '------------------------------------'
    '    vlstrSentencia = "SELECT CcNotaFactura.*, PvFactura.bitpesos, PvFactura.mnytipocambio FROM CcNotaFactura, PvFactura WHERE CcNotaFactura.intConsecutivo = " & vllngconsecutivoNota & " and trim(CcNotaFactura.chrfoliofactura) = trim(PvFactura.chrfoliofactura)"
        
        vlstrSentencia = "SELECT CcNotaFactura.*, CASE WHEN CcNotaFactura.CHRTIPOFOLIO = 'FA' THEN CASE WHEN PvFactura.bitpesos IS NULL THEN 1 ELSE PvFactura.bitpesos END ELSE 1 END bitpesos, CASE WHEN CcNotaFactura.CHRTIPOFOLIO = 'FA' THEN CASE WHEN PvFactura.mnytipocambio IS NULL THEN 0 ELSE PvFactura.mnytipocambio END ELSE 0 END mnytipocambio  " & _
                         "From CcNotaFactura " & _
                            "LEFT JOIN PvFactura ON trim(CcNotaFactura.chrfoliofactura) = trim(PvFactura.chrfoliofactura) " & _
                         "WHERE CcNotaFactura.intConsecutivo = " & vllngconsecutivoNota
        Set rsNotasFacturas = frsRegresaRs(vlstrSentencia)
        Do While Not rsNotasFacturas.EOF
            vldblTipoCambio = IIf(rsNotasFacturas!BITPESOS = 1, 1, rsNotasFacturas!MNYTIPOCAMBIO)
            If vlstrTipoNota = "CR" And lngFacturaPagada = 0 Then
                vlstrSentencia = " UPDATE CCMOVIMIENTOCREDITO " & _
                                 " SET CCMOVIMIENTOCREDITO.MNYCANTIDADPAGADA = CCMOVIMIENTOCREDITO.MNYCANTIDADPAGADA - " & Str(((rsNotasFacturas!MNYSUBTOTAL * vldblTipoCambio) - (rsNotasFacturas!MNYDESCUENTO * vldblTipoCambio) + Format((rsNotasFacturas!MNYIVA * vldblTipoCambio), vlstrFormato))) & _
                                 " WHERE CCMOVIMIENTOCREDITO.intNumMovimiento = " & rsNotasFacturas!intNumMovimientoCredito
                pEjecutaSentencia vlstrSentencia
            Else
              '  vlstrSentencia = "DELETE FROM CcMovimientoCredito WHERE intNumMovimiento = " & rsNotasFacturas!intNumMovimientoCredito
                vlstrSentencia = "UPDATE CCMOVIMIENTOCREDITO " & _
                                         "SET  CCMOVIMIENTOCREDITO.BITCANCELADO =" & 1 & "," & _
                                         "CCMOVIMIENTOCREDITO.DTMFECHACANCELACION = " & fstrFechaSQL(fdtmServerFecha) & _
                                         "WHERE CCMOVIMIENTOCREDITO.intNumMovimiento = " & rsNotasFacturas!intNumMovimientoCredito
                                         'Modificado para el caso 12095 SC
                pEjecutaSentencia vlstrSentencia
            End If
            rsNotasFacturas.MoveNext
        Loop
    
        Call pGuardarLogTransaccion(Me.Name, EnmCancelacion, vllngPersonaGraba, IIf(vlstrTipoNota = "CR", "NOTAS DE CREDITO", "NOTAS DE CARGO"), Trim(vlstrFolioNota))
    
        pEjecutaSentencia "UPDATE CnEstatusCierre SET vchEstatus = 'Libre' WHERE tnyClaveEmpresa = " + Str(vgintClaveEmpresaContable)
        
       
        EntornoSIHO.ConeccionSIHO.CommitTrans
        
        If vlblnMuestramensaje Then pMensajeCanelacionCFDi vllngconsecutivoNota, vlstrTipoNota
    End If

Exit Sub
NotificaError:
    frmNotas.blnCancelaNota = False
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCancelaNota"))
    Unload Me
End Sub
Private Sub cmdGrabarRegistro_Click()
On Error GoTo NotificaError

    If Trim(frmNotas.txtCveCliente.Text) <> "" Then 'Si es nota hacer...
    '    vgblnAuxNotas = False 'Inicialización del Auxiliar
        pCancelaNota frmNotas.lblFolio, frmNotas.txtCveCliente.Text, "ELEGIR", True, True, False
    End If
    
    Unload Me

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdGrabarRegistro_Click"))
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Me.Icon = frmMenuPrincipal.Icon
End Sub

Private Sub OptCancela_Click(Index As Integer)

End Sub

Private Sub OptCancela_GotFocus(Index As Integer)

End Sub

Private Sub OptCancelaTipoActual_GotFocus()
    If MsgBox(SIHOMsg("1046"), vbYesNo + vbExclamation, "Mensaje") = vbYes Then
        OptCancelaTipoActual.Value = True
        If cmdGrabarRegistro.Enabled = True And cmdGrabarRegistro.Visible = True Then
            cmdGrabarRegistro.SetFocus
        End If
    Else
        If cmdGrabarRegistro.Enabled = True And cmdGrabarRegistro.Visible = True Then
            cmdGrabarRegistro.SetFocus
        End If
    End If
    
    Exit Sub
End Sub

Private Sub OptCancelaTipoActual_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdGrabarRegistro.SetFocus
    End If
End Sub

Private Sub OptCancelaTipoDocu_Click()
    OptCancelaTipoDocu.Value = True
End Sub

Private Sub OptCancelaTipoDocu_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdGrabarRegistro.SetFocus
    End If
End Sub
