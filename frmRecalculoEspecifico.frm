VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRecalculoEspecifico 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recálculo de existencias y costos promedios"
   ClientHeight    =   525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7740
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   525
   ScaleWidth      =   7740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtArticulo 
      BorderStyle     =   0  'None
      Height          =   232
      Left            =   45
      TabIndex        =   6
      Text            =   "txtArticulo"
      Top             =   10
      Width           =   7575
   End
   Begin VB.Frame Frame2 
      Caption         =   "Almacén"
      Height          =   660
      Left            =   105
      TabIndex        =   3
      Top             =   675
      Visible         =   0   'False
      Width           =   5055
      Begin VB.ComboBox cboAlmacen 
         Height          =   315
         Left            =   135
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Selección del almacén"
         Top             =   225
         Visible         =   0   'False
         Width           =   4770
      End
   End
   Begin VB.Frame Frame1 
      Height          =   765
      Left            =   6120
      TabIndex        =   2
      Top             =   6600
      Visible         =   0   'False
      Width           =   1490
      Begin VB.CommandButton cmdRecalculo 
         Caption         =   "Recalcular"
         Enabled         =   0   'False
         Height          =   525
         Left            =   60
         TabIndex        =   1
         ToolTipText     =   "Iniciar el proceso de recálculo"
         Top             =   165
         Visible         =   0   'False
         Width           =   1350
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdArticulos 
      Height          =   5145
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "Lista de artículos modificados"
      Top             =   1380
      Visible         =   0   'False
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   9075
      _Version        =   393216
      Cols            =   7
      GridColor       =   12632256
      FormatString    =   "|Movimiento|Fecha|Referencia|Consecutivo|Clave|Artículo"
      _NumberOfBands  =   1
      _Band(0).Cols   =   7
   End
   Begin MSComctlLib.ProgressBar pgbProceso 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   5
      Top             =   240
      Width           =   7740
      _ExtentX        =   13653
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "frmRecalculoEspecifico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************
'* Pantalla para recalcular un articulo desde algun proceso sin mostrar la pantalla de recalculo
'* Fecha de programación: 24/04/2013
'* Ultima modificación:
'*******************************************************************************

Option Explicit

Public vlstrCveArticulo As String
Public vlblnnegativos As Boolean
Public vlstrProceso As String
Public vlcveMovimiento As Long
Public vllngKardexMovimiento As Long
Public vlblnBorroRecepcion As Boolean
Public vllngCuentaPagarABorrar As Long
Public vlintNumeroDepto As Integer
Public vllngPolizaConsignaBorrar As Long

Private Type MovtosKardex
    Movimiento As Long
    fecha As Date
    Referencia As String
    Consecutivo As Double
    clave As String
    Articulo As String
    Contenido As Long
    CveDepto As Integer
    PosicionGrid As Long
End Type

Dim aMovtoKardex() As MovtosKardex

Dim vlstrx As String

Dim rsMovimientosAfectados As New ADODB.Recordset

Dim vllngTotalUV As Long
Dim vllngTotalUM As Long
Dim vllngCantMovtoUM As Long
Dim vllngCantMovtoUV As Long

Private Type Datospoliza
    NumerodeCuenta As Long
    CantidadMovimiento As Double
    NaturalezaMovimiento As Integer
End Type

Dim apoliza() As Datospoliza

Dim intBitDescuentoCostoPromedio As Integer

Private Function fdblCostoEntradaSalidaFS(vldblNumRef As Double, vlstrCveArticulo As String, vlintCveDepto As Integer, vllngMovKardex As Long) As Double
    Dim rsTemp As ADODB.Recordset
    ' ENTRADA / SALIDA DE LA FARMACIA SUBROGADA
    
    fdblCostoEntradaSalidaFS = 0
    vlstrx = ""
    vlstrx = vlstrx & _
            "SELECT NVL(IVREQUISICIONESSURTIDAS.NUMPRECIOSURTIDO, 0) NUMPRECIOSURTIDO " & _
            "FROM IVREQUISICIONDETALLESUB " & _
            "INNER JOIN IVARTICULO ON IVREQUISICIONDETALLESUB.CHRCVEARTICULO = IVARTICULO.CHRCVEARTICULO " & _
            "INNER JOIN IVREQUISICIONESSURTIDAS ON IVREQUISICIONESSURTIDAS.NUMNUMREQUISICION = IVREQUISICIONDETALLESUB.NUMNUMREQUISICION AND IVREQUISICIONESSURTIDAS.INTIDARTICULOSURTIDO = IVARTICULO.INTIDARTICULO " & _
            "WHERE IVREQUISICIONDETALLESUB.NUMNUMREQUISICION = " & Str(vldblNumRef) & " AND IVREQUISICIONDETALLESUB.CHRCVEARTICULO = '" & Trim(vlstrCveArticulo) & "'"
        
    Set rsTemp = frsRegresaRs(vlstrx)
    If rsTemp.RecordCount > 0 Then
        fdblCostoEntradaSalidaFS = rsTemp!NUMPRECIOSURTIDO
    Else
        rsTemp.Close
        vlstrx = ""
        vlstrx = vlstrx & _
        "SELECT NVL(IVREQUISICIONESSURTIDAS.NUMPRECIOSURTIDO, 0) NUMPRECIOSURTIDO From IVREUBICACIONMAESTRO " & _
        "INNER JOIN IVREQUISICIONESSURTIDAS ON IVREUBICACIONMAESTRO.NUMNUMREQUISICION = IVREQUISICIONESSURTIDAS.NUMNUMREQUISICION " & _
        "AND IVREUBICACIONMAESTRO.NUMNUMREUBICACION = IVREQUISICIONESSURTIDAS.NUMNUMCARGO " & _
        "INNER JOIN IVARTICULO ON IVREQUISICIONESSURTIDAS.INTIDARTICULOSURTIDO = IVARTICULO.INTIDARTICULO " & _
        "Where IVREUBICACIONMAESTRO.NUMNUMREUBICACION = " & Str(vldblNumRef) & " AND IVARTICULO.CHRCVEARTICULO = '" & Trim(vlstrCveArticulo) & "'"
        Set rsTemp = frsRegresaRs(vlstrx)
        If rsTemp.RecordCount > 0 Then
            fdblCostoEntradaSalidaFS = rsTemp!NUMPRECIOSURTIDO
        Else
            rsTemp.Close
            vlstrx = ""
            vlstrx = vlstrx & _
            "SELECT NVL(IVREQUISICIONESSURTIDAS.NUMPRECIOSURTIDO, 0) NUMPRECIOSURTIDO From IVSALIDADEPTOMAESTRO " & _
            "INNER JOIN IVREQUISICIONESSURTIDAS ON IVSALIDADEPTOMAESTRO.NUMNUMREQUISICION = IVREQUISICIONESSURTIDAS.NUMNUMREQUISICION " & _
            "AND IVSALIDADEPTOMAESTRO.NUMNUMSALIDADEPTO = IVREQUISICIONESSURTIDAS.NUMNUMCARGO " & _
            "INNER JOIN IVARTICULO ON IVREQUISICIONESSURTIDAS.INTIDARTICULOSURTIDO = IVARTICULO.INTIDARTICULO " & _
            "Where IVSALIDADEPTOMAESTRO.NUMNUMSALIDADEPTO = " & Str(vldblNumRef) & " AND IVARTICULO.CHRCVEARTICULO = '" & Trim(vlstrCveArticulo) & "'"
            Set rsTemp = frsRegresaRs(vlstrx)
            If rsTemp.RecordCount > 0 Then
                fdblCostoEntradaSalidaFS = rsTemp!NUMPRECIOSURTIDO
            Else
                rsTemp.Close
                vlstrx = ""
                vlstrx = vlstrx & _
                "SELECT Distinct NVL(IVREQUISICIONESSURTIDAS.NUMPRECIOSURTIDO, 0) NUMPRECIOSURTIDO From IVSALIDADEPTOMAESTRO " & _
                "INNER JOIN IVSALIDADEPTODETALLE ON IVSALIDADEPTOMAESTRO.NUMNUMSALIDADEPTO = IVSALIDADEPTODETALLE.NUMNUMSALIDADEPTO " & _
                "INNER JOIN IVREQUISICIONESSURTIDAS ON IVSALIDADEPTOMAESTRO.NUMNUMREQUISICION = IVREQUISICIONESSURTIDAS.NUMNUMREQUISICION " & _
                "AND IVSALIDADEPTOMAESTRO.NUMNUMSALIDADEPTO = IVREQUISICIONESSURTIDAS.NUMNUMCARGO " & _
                "INNER JOIN IVDEVOLUCIONDEPTOMAESTRO ON IVSALIDADEPTOMAESTRO.NUMNUMSALIDADEPTO = IVDEVOLUCIONDEPTOMAESTRO.NUMNUMREFERENCIA " & _
                "Where IVDEVOLUCIONDEPTOMAESTRO.NUMNUMDEVOLUCION = " & Str(vldblNumRef) & " AND IVSALIDADEPTODETALLE.CHRCVEARTICULO = '" & Trim(vlstrCveArticulo) & "'"
                Set rsTemp = frsRegresaRs(vlstrx)
                If rsTemp.RecordCount > 0 Then
                    fdblCostoEntradaSalidaFS = rsTemp!NUMPRECIOSURTIDO
                Else
                    rsTemp.Close
                    vlstrx = ""
                    vlstrx = vlstrx & _
                    "SELECT Distinct NVL(IVREQUISICIONESSURTIDAS.NUMPRECIOSURTIDO, 0) NUMPRECIOSURTIDO From IVREUBICACIONMAESTRO " & _
                    "INNER JOIN IVREUBICACIONDETALLE ON IVREUBICACIONMAESTRO.NUMNUMREUBICACION = IVREUBICACIONDETALLE.NUMNUMREUBICACION " & _
                    "INNER JOIN IVREQUISICIONESSURTIDAS ON IVREUBICACIONMAESTRO.NUMNUMREQUISICION = IVREQUISICIONESSURTIDAS.NUMNUMREQUISICION " & _
                    "AND IVREUBICACIONMAESTRO.NUMNUMREUBICACION = IVREQUISICIONESSURTIDAS.NUMNUMCARGO " & _
                    "INNER JOIN IVDEVOLUCIONDEPTOMAESTRO ON IVREUBICACIONMAESTRO.NUMNUMREUBICACION = IVDEVOLUCIONDEPTOMAESTRO.NUMNUMREFERENCIA " & _
                    "Where IVDEVOLUCIONDEPTOMAESTRO.NUMNUMDEVOLUCION = " & Str(vldblNumRef) & " And IVREUBICACIONDETALLE.CHRCVEARTICULO = '" & Trim(vlstrCveArticulo) & "'"
                    Set rsTemp = frsRegresaRs(vlstrx)
                    If rsTemp.RecordCount > 0 Then
                        fdblCostoEntradaSalidaFS = rsTemp!NUMPRECIOSURTIDO
                    Else
                        fdblCostoEntradaSalidaFS = 0
                    End If
                End If
            End If
        End If
    End If
    rsTemp.Close
    
    If fdblCostoEntradaSalidaFS = 0 Then
        Set rsTemp = frsRegresaRs("SELECT ISNULL(mnyCostoPromedio, 0) mnyCostoPromedio FROM IvUbicacion WHERE smiCveDepartamento = " & vlintCveDepto & " AND chrCveArticulo = '" & Trim(vlstrCveArticulo) & "'")
        If rsTemp.RecordCount > 0 Then
            fdblCostoEntradaSalidaFS = rsTemp!MNYCOSTOPROMEDIO
        End If
        rsTemp.Close
    End If
End Function

Private Function fdblCostoEntradaECO(vldblNumConsecutivo As Double, vlstrCveArticulo As String, CveDepto As Integer) As Double
    Dim rsTemp As ADODB.Recordset
    
    fdblCostoEntradaECO = 0
    
    vlstrx = ""
    vlstrx = vlstrx & _
    "select " & _
        "mnyprecioproveedor " & _
    "From " & _
        "IvEntradaSalidaConsignaDetalle " & _
    "Where " & _
        "numEntradaSalidaConsignacion =" & Str(vldblNumConsecutivo) & " and " & _
        "chrCveArticulo = '" & Trim(vlstrCveArticulo) & "' "
            
    Set rsTemp = frsRegresaRs(vlstrx)
    If rsTemp.RecordCount > 0 Then
        fdblCostoEntradaECO = rsTemp!mnyprecioproveedor
    End If
    rsTemp.Close
    
    If fdblCostoEntradaECO = 0 Then
        Set rsTemp = frsRegresaRs("SELECT ISNULL(mnyCostoPromedio, 0) mnyCostoPromedio FROM IvUbicacion WHERE smiCveDepartamento = " & CveDepto & " AND chrCveArticulo = '" & Trim(vlstrCveArticulo) & "'")
        If rsTemp.RecordCount > 0 Then
            fdblCostoEntradaECO = rsTemp!MNYCOSTOPROMEDIO
        End If
        rsTemp.Close
    End If
End Function

Private Function fdblCostoEntradaEAJ(vldblNumCaptura As Double, vlstrCveArticulo As String) As Double

    vlstrx = ""
    vlstrx = vlstrx & _
    "select " & _
        "IvInvFisicoHistorico.mnyCostoPromedio " & _
    "From " & _
        "IvInvFisicoHistorico " & _
    "Where " & _
        "IvInvFisicoHistorico.intcvecaptura=" & Str(vldblNumCaptura) & " and " & _
        "IvInvFisicoHistorico.chrCveAriculo='" & Trim(vlstrCveArticulo) & "' " & _
    "union " & _
    "select " & _
        "IvInvFisico.mnyCostoPromedio " & _
    "From " & _
        "IvInvFisico " & _
    "Where " & _
        "IvInvFisico.intcvecaptura=" & Str(vldblNumCaptura) & " and " & _
        "IvInvFisico.chrCveArticulo='" & Trim(vlstrCveArticulo) & "' "

    fdblCostoEntradaEAJ = frsRegresaRs(vlstrx).Fields(0)
    
End Function


 Private Function fdblCostoEntradaERE_SD(vldblNumRecepcion As Double, vlstrCveArticulo As String) As Double

    vlstrx = ""
    vlstrx = vlstrx & _
    "select mnyCostoEntrada as Costo " & _
    "From " & _
        "IvRecepcionDetalle " & _
    "Where " & _
        "numNumRecepcion=" & Str(vldblNumRecepcion) & " and " & _
        "chrCveArticulo='" & Trim(vlstrCveArticulo) & "' "
        
    fdblCostoEntradaERE_SD = frsRegresaRs(vlstrx, adLockOptimistic, adOpenDynamic).Fields(0)

End Function

Private Sub pLlenapoliza(vllngxNumeroCuenta As Long, vldblxCantidad As Double, vlintxTipoMovto As Integer)
On Error GoTo NotificaError

    Dim vlintTamaño As Integer
    Dim vlintPosicion As Integer
    Dim vlblnEstaCuenta As Boolean
    Dim X As Integer
    
    If apoliza(0).NumerodeCuenta = 0 Then
        apoliza(0).NumerodeCuenta = vllngxNumeroCuenta
        apoliza(0).CantidadMovimiento = vldblxCantidad
        apoliza(0).NaturalezaMovimiento = vlintxTipoMovto
    Else
        vlblnEstaCuenta = False
        vlintTamaño = UBound(apoliza(), 1)
        For X = 0 To vlintTamaño
            If apoliza(X).NumerodeCuenta = vllngxNumeroCuenta Then
                vlblnEstaCuenta = True
                vlintPosicion = X
            End If
        Next X
        If vlblnEstaCuenta Then
            If apoliza(vlintPosicion).NaturalezaMovimiento = vlintxTipoMovto Then
                apoliza(vlintPosicion).CantidadMovimiento = apoliza(vlintPosicion).CantidadMovimiento + vldblxCantidad
            Else
                If apoliza(vlintPosicion).CantidadMovimiento > vldblxCantidad Then
                    apoliza(vlintPosicion).CantidadMovimiento = apoliza(vlintPosicion).CantidadMovimiento - vldblxCantidad
                Else
                    If apoliza(vlintPosicion).CantidadMovimiento < vldblxCantidad Then
                        apoliza(vlintPosicion).CantidadMovimiento = vldblxCantidad - apoliza(vlintPosicion).CantidadMovimiento
                        If apoliza(vlintPosicion).NaturalezaMovimiento = 1 Then
                            apoliza(vlintPosicion).NaturalezaMovimiento = 0
                        Else
                            apoliza(vlintPosicion).NaturalezaMovimiento = 1
                        End If
                    Else
                        apoliza(vlintPosicion).NumerodeCuenta = 0
                        apoliza(vlintPosicion).CantidadMovimiento = 0
                        apoliza(vlintPosicion).NaturalezaMovimiento = 0
                    End If
                End If
            End If
        Else
            ReDim Preserve apoliza(vlintTamaño + 1)
            apoliza(vlintTamaño + 1).NumerodeCuenta = vllngxNumeroCuenta
            apoliza(vlintTamaño + 1).CantidadMovimiento = vldblxCantidad
            apoliza(vlintTamaño + 1).NaturalezaMovimiento = vlintxTipoMovto
        End If
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pLlenaPoliza"))
End Sub



Private Function fBlnExtraeParametropoliza(vlstrTipoCA As String, vlstrTemp As String, vlstrArticulo As String, ByRef vllngCuenta As Long) As Boolean
On Error GoTo NotificaError
  Dim rs As New ADODB.Recordset
  Dim vlstrx As String
  Dim vlblnError As Boolean
  Dim rsDescripcionArt As New ADODB.Recordset
  Dim rsdatosrequisicion As New ADODB.Recordset
  Dim vlintEmpresa As Integer
  Dim vlintDepartamento As Integer
  
  '--------------------------------------------------------------------------------------------------------
  'No se utilizan las variables globales , ya que los parametros y cuentas contables dependen de la empresa
  'que requisito los artículos y no de la empresa que esta realizando la recepción
  '--------------------------------------------------------------------------------------------------------
  
'    Set rsdatosrequisicion = frsEjecuta_SP(Str(vlintNumOrden), "sp_IvSelDatosRequisicion")
'    If rsdatosrequisicion.RecordCount > 0 Then
'        vlintempresa = rsdatosrequisicion!tnyclaveempresa
'        vlintDepartamento = rsdatosrequisicion!smiCveDeptoRequis
'    Else
        vlintEmpresa = vgintClaveEmpresaContable
        vlintDepartamento = vgintNumeroDepartamento
'    End If
  
  fBlnExtraeParametropoliza = False
  vlblnError = False
  If Not vlblnError Then
    If vlstrTipoCA = "A" Or vlstrTipoCA = "C" Or vlstrTipoCA = "G" Then
      vlstrx = ""
      vlstrx = vlstrx & "SELECT intCveDepartamento, chrTipoActivoGastoCosto, intNumCuentaContable, chrTipoArticuloMedicamentoInsu, chrCveFamilia, chrCveSubFamilia, chrTipoGastoCostoAmbos "
      vlstrx = vlstrx & " FROM IvParametroPoliza "
      vlstrx = vlstrx & " WHERE chrTipoActivoGastoCosto ='" & vlstrTipoCA & "' "
      vlstrx = vlstrx & " AND TNYCLAVEEMPRESA = " & vlintEmpresa
      If (Mid(vlstrTemp, 1, 1) = "1") Then
        vlstrx = vlstrx & " AND intCveDepartamento=" & CStr(vlintDepartamento)
      End If
      If ((Mid(vlstrTemp, 2, 1) > "1") And (Mid(vlstrTemp, 2, 1) < "5")) Then
        vlstrx = vlstrx & " AND chrTipoArticuloMedicamentoInsu=" & fstrTipoFamiliaSubFamilia(vlstrArticulo, 1) & " "
        If ((Mid(vlstrTemp, 2, 1) = "3") Or (Mid(vlstrTemp, 2, 1) = "4")) Then
          vlstrx = vlstrx & " AND chrCveFamilia=" & fstrTipoFamiliaSubFamilia(vlstrArticulo, 2) & " "
          If (Mid(vlstrTemp, 2, 1) = "4") Then
            vlstrx = vlstrx & " AND chrCveSubFamilia=" & fstrTipoFamiliaSubFamilia(vlstrArticulo, 3) & " "
          End If
        End If
      End If
      If (Mid(vlstrTemp, 3, 1) = "5") Then
        vlstrx = vlstrx & " AND chrTipoGastoCostoAmbos='" & frsRegresaRs("SELECT chrCostoGasto FROM IvArticulo WHERE chrCveArticulo='" & vlstrArticulo & "' ").Fields(0) & "' "
      End If
      Screen.MousePointer = vbHourglass
      Set rs = frsRegresaRs(vlstrx)
      Screen.MousePointer = vbDefault
      If (rs.State <> adStateClosed) Then
        If rs.RecordCount > 0 Then 'Existe cuenta asignada a esa combinación, o.k.
          rs.MoveFirst
          vllngCuenta = rs!intNumCuentaContable
          fBlnExtraeParametropoliza = True
        Else
          'Kardex del artículo X para el tipo de poliza Y no tiene asignada la cuenta contable para generar la póliza
          Set rsDescripcionArt = frsRegresaRs("SELECT vchNombreComercial FROM IvArticulo WHERE chrCveArticulo='" & vlstrArticulo & "' ")
          MsgBox "Recepción del artículo " & vlstrArticulo & " (" & rsDescripcionArt!vchNombreComercial & ") para el tipo de poliza de " & IIf(vlstrTipoCA = "A", "ACTIVO", IIf(vlstrTipoCA = "C", "COSTO", "GASTO")) & Chr(13) & "no tiene asignada la cuenta contable para generar la póliza", vbOKOnly + vbInformation, "Mensaje"
        End If
        rs.Close
      Else
        'Error al tratar de extraer cuenta contable para generar la póliza
        MsgBox "Error al tratar de extraer cuenta contable para generar la póliza", vbOKOnly + vbInformation, "Mensaje"
      End If
    End If
  End If

Exit Function
NotificaError:
    Screen.MousePointer = vbDefault
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fBlnExtraeParametroPoliza"))
End Function



Private Function fblnParametropoliza(vlstrArticulo As String, ByRef vlLngCuentaAbono As Long) As Boolean
On Error GoTo NotificaError
    Dim vllngCuenta As Long
  
    fblnParametropoliza = False
    vllngCuenta = 0
    If fBlnExtraeParametropoliza("A", gstrConfiguracionActivo, vlstrArticulo, vllngCuenta) Then
        vlLngCuentaAbono = vllngCuenta ' Sólo la cuenta de Cargo
        fblnParametropoliza = True
    End If

Exit Function
NotificaError:
  Screen.MousePointer = vbDefault
  Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnParametroPoliza"))
End Function

Private Sub pRegistroPolizaAjuste(vlintCveDepartamento As Integer, vlstrCveArticulo As String, vlintCantidadArtDev As Integer, vlintCtoProm As Double, vlintNumDevolucion As Integer, vllngPersonaGraba As Long, vllngCuentaAjuste As Long, vldtmFecha As Date, vldblSubtotal As Double, vldblIVA As Double, vldblDescuento As Double, vlstrFolio As String, vldblCostoEntrada As Double)
    On Error GoTo NotificaError
    Dim rsCantidadCuentaPago As New ADODB.Recordset
    Dim vlLngCuentaAbono As Long
    Dim vlCargos As Double
    Dim vlAbonos As Double
    Dim lstrConceptoPoliza As String
    Dim vlintNumeroProceso As Integer
    Dim i As Integer
    Dim vlintAbonoAlmacen As Integer
    Dim vldblAjuste As Double
    Dim vlstrActivoGastoCosto As String
    Dim rs As New ADODB.Recordset
    Dim vlstrNombreComercialArticulo As String
    Dim vldblMovimientoCostoProm As Double
    Dim vlstrSentencia As String
    Dim vlintnumeropolizaajuste As Integer
    Dim vllngNumeroPolizaAjuste As Long
    Dim vllngnoerror  As Long
    Dim rsCnPoliza As ADODB.Recordset
    Dim rsCnDetallePoliza As ADODB.Recordset
    Dim X As Integer
    Dim rsPolizaAjuste As ADODB.Recordset
    Dim vldblMovimientoCostoEntrada As Double
    Dim ldblMovimientoCostoProm As Double
    
    vldtmFecha = fdtmServerFecha
    
    'Cuenta de costo para ajuste por valuación
    '------------------------------------------
    vlstrActivoGastoCosto = "C"
    vlstrSentencia = ""
    vlstrSentencia = vlstrSentencia & "SELECT intCveDepartamento, chrTipoActivoGastoCosto, intNumCuentaContable, chrTipoArticuloMedicamentoInsu, chrCveFamilia, chrCveSubFamilia, chrTipoGastoCostoAmbos "
    vlstrSentencia = vlstrSentencia & " FROM IvParametropoliza "
    vlstrSentencia = vlstrSentencia & " WHERE chrTipoActivoGastoCosto ='" & vlstrActivoGastoCosto & "' "
    vlstrSentencia = vlstrSentencia & " and tnyclaveempresa = " & vgintClaveEmpresaContable
    If (Mid(gstrConfiguracionCosto, 1, 1) = "1") Then
      vlstrSentencia = vlstrSentencia & " AND intCveDepartamento=" & CStr(vlintCveDepartamento)
    End If
    If ((Mid(gstrConfiguracionCosto, 2, 1) > "1") And (Mid(gstrConfiguracionCosto, 2, 1) < "5")) Then
      vlstrSentencia = vlstrSentencia & " AND chrTipoArticuloMedicamentoInsu=" & fstrTipoFamiliaSubFamilia(Trim(vlstrCveArticulo), 1) & " "
      If ((Mid(gstrConfiguracionCosto, 2, 1) = "3") Or (Mid(gstrConfiguracionCosto, 2, 1) = "4")) Then
        vlstrSentencia = vlstrSentencia & " AND chrCveFamilia=" & fstrTipoFamiliaSubFamilia(Trim(vlstrCveArticulo), 2) & " "
        If (Mid(gstrConfiguracionCosto, 2, 1) = "4") Then
          vlstrSentencia = vlstrSentencia & " AND chrCveSubFamilia=" & fstrTipoFamiliaSubFamilia(Trim(vlstrCveArticulo), 3) & " "
        End If
      End If
    End If
    If (Mid(gstrConfiguracionCosto, 3, 1) = "5") Then
      vlstrSentencia = vlstrSentencia & " AND chrTipoGastoCostoAmbos='" & frsRegresaRs("SELECT chrCostoGasto FROM IvArticulo WHERE chrCveArticulo='" & Trim(vlstrCveArticulo) & "' ").Fields(0) & "' "
    End If
    Screen.MousePointer = vbHourglass
    Set rs = frsRegresaRs(vlstrSentencia)
    Screen.MousePointer = vbDefault
    If (rs.State <> adStateClosed) Then
      If rs.RecordCount > 0 Then 'Existe cuenta asignada a esa combinación, o.k.
        rs.MoveFirst
        vllngCuentaAjuste = rs!intNumCuentaContable
      Else
        vlstrNombreComercialArticulo = Trim(frsRegresaRs("Select vchnombrecomercial from ivarticulo where chrCveArticulo = '" & Trim(vlstrCveArticulo) & "'").Fields(0))
        'Kardex del artículo X para el tipo de poliza Y no tiene asignada la cuenta contable para generar la poliza
        MsgBox "Kardex del artículo " & Trim(vlstrCveArticulo) & " / " & vlstrNombreComercialArticulo & " para el tipo de póliza " & IIf(vlstrActivoGastoCosto = "A", "ACTIVOS", IIf(vlstrActivoGastoCosto = "C", "COSTO", "GASTO")) & " no tiene asignada la cuenta contable del departamento " & frsRegresaRs("SELECT vchDescripcion FROM noDepartamento WHERE smiCveDepartamento=" & CStr(vlintCveDepartamento)).Fields(0) & " para generar la póliza", vbOKOnly + vbInformation, "Mensaje"
        vgblnErrorIngreso = True
        Exit Sub
      End If
      rs.Close
    Else
      vlstrNombreComercialArticulo = Trim(frsRegresaRs("Select vchnombrecomercial from ivarticulo where chrCveArticulo = '" & Trim(vlstrCveArticulo) & "'").Fields(0))
      'Error al tratar de extraer cuenta contable para generar la poliza
      MsgBox "Error al tratar de extraer cuenta contable para generar la póliza", vbOKOnly + vbInformation, "Mensaje"
      vgblnErrorIngreso = True
      Exit Sub
    End If
    '------------------------------------------

    vldblMovimientoCostoProm = vlintCantidadArtDev * vlintCtoProm
    vldblMovimientoCostoEntrada = vlintCantidadArtDev * vldblCostoEntrada
    
    If vldblMovimientoCostoEntrada > vldblMovimientoCostoProm Then
        vlintAbonoAlmacen = 1
        vldblAjuste = vldblMovimientoCostoEntrada - vldblMovimientoCostoProm
    ElseIf vldblMovimientoCostoEntrada < vldblMovimientoCostoProm Then
        vlintAbonoAlmacen = 0
        vldblAjuste = vldblMovimientoCostoProm - vldblMovimientoCostoEntrada
    ElseIf (vldblMovimientoCostoEntrada - vldblMovimientoCostoProm) = 0 Then
        vldblAjuste = vldblMovimientoCostoEntrada - vldblMovimientoCostoProm
    End If
    
    ReDim apoliza(0)
    vllngNumeroPolizaAjuste = 0
    apoliza(0).NumerodeCuenta = 0
    
    vlstrx = " select po.dtmfechapoliza, dp.mnycantidadmovimiento, nf.intnumeropolizaajuste " & _
             " from cpnotafactura nf, ivdevoluproveedmaestro dpo, cnpoliza po, cndetallepoliza dp " & _
             " where dpo.intnumdevolucion = " & vlintNumDevolucion & _
             " and dpo.intnumnota = nf.intnumnota and nf.intnumeropolizaajuste = po.intnumeropoliza " & _
             " and po.intnumeropoliza = dp.intnumeropoliza and dp.intnumerocuenta = " & vllngCuentaAjuste
    Set rsPolizaAjuste = frsRegresaRs(vlstrx, adLockOptimistic, adOpenDynamic)
    If rsPolizaAjuste.RecordCount > 0 Then  'Si existe póliza de ajuste
        'Si existe póliza de ajuste, verifica si el período contable al que pertenece no se encuentre cerrado
        If fblnPeriodoCerrado(vgintClaveEmpresaContable, Year(rsPolizaAjuste!dtmFechaPoliza), Month(rsPolizaAjuste!dtmFechaPoliza)) Then
            frmEsperaAbranPeriodoContable.vlintEjercicio = Year(rsPolizaAjuste!dtmFechaPoliza)
            frmEsperaAbranPeriodoContable.vlintMes = Month(rsPolizaAjuste!dtmFechaPoliza)
            frmEsperaAbranPeriodoContable.Show vbModal
        End If
        If vldblAjuste = 0 Then
            'Cancela póliza, si después de recalcular no existe ajuste por diferencia en la valuación del inventario
            pEjecutaSentencia "DELETE cndetallepoliza WHERE intnumeropoliza = " & rsPolizaAjuste!intNumeroPolizaAjuste
            pEjecutaSentencia "UPDATE cnpoliza SET vchconceptopoliza = 'CANCELADA' WHERE intnumeropoliza = " & rsPolizaAjuste!intNumeroPolizaAjuste
        Else
            If vldblAjuste <> rsPolizaAjuste!mnyCantidadMovimiento Then
                'Si el ajuste antes del recalculo es diferente al ajuste después del recalculo, agrega nuevo valor del ajuste
                'Movimiento al almacen
                pEjecutaSentencia "DELETE cndetallepoliza WHERE intnumeropoliza = " & rsPolizaAjuste!intNumeroPolizaAjuste
                '-------------------------------------------
                ' Recordset tipo tabla CnDetallePoliza
                '-------------------------------------------
                vlstrx = " " & _
                  "select INTNUMEROPOLIZA, INTNUMEROCUENTA, BITNATURALEZAMOVIMIENTO, MNYCANTIDADMOVIMIENTO,VCHREFERENCIA, VCHCONCEPTO, INTNUMEROREGISTRO " & _
                  " from CnDetallePoliza where intNumeroPoliza = -1"
                Set rsCnDetallePoliza = frsRegresaRs(vlstrx, adLockOptimistic, adOpenDynamic)
                
                If fblnParametropoliza(Trim(vlstrCveArticulo), vlLngCuentaAbono) Then
                    pLlenapoliza vlLngCuentaAbono, vldblAjuste, IIf(vlintAbonoAlmacen = 0, 0, 1)
                    pLlenapoliza vllngCuentaAjuste, vldblAjuste, IIf(vlintAbonoAlmacen = 0, 1, 0)
                    For X = 0 To UBound(apoliza(), 1)
                        If apoliza(X).NumerodeCuenta <> 0 Then
                            With rsCnDetallePoliza
                                .AddNew
                                !INTNUMEROPOLIZA = rsPolizaAjuste!intNumeroPolizaAjuste
                                !INTNUMEROCUENTA = apoliza(X).NumerodeCuenta
                                !bitNaturalezaMovimiento = apoliza(X).NaturalezaMovimiento
                                !mnyCantidadMovimiento = apoliza(X).CantidadMovimiento
                                !vchConcepto = " "
                                !vchReferencia = " "
                                .Update
                            End With
                        End If
                    Next X
                Else
                    'Todo para atras pues esto es un error ya que falta la definición de la cuenta contable de activos
                    vgblnErrorIngreso = True
                    Exit Sub
                End If
            End If
        End If
    Else
        If vldblAjuste > 0 Then
            'Movimiento al almacen
            If fblnParametropoliza(Trim(vlstrCveArticulo), vlLngCuentaAbono) Then
                pLlenapoliza vlLngCuentaAbono, vldblAjuste, IIf(vlintAbonoAlmacen = 0, 0, 1)
                pLlenapoliza vllngCuentaAjuste, vldblAjuste, IIf(vlintAbonoAlmacen = 0, 1, 0)
            Else
                'Todo para atras pues esto es un error ya que falta la definición de la cuenta contable de activos
                vgblnErrorIngreso = True
                Exit Sub
            End If
                    
            lstrConceptoPoliza = "Ajuste por diferencia en valuación de inventario en devolución número: " & vlintNumDevolucion         'fstrConceptoPoliza
            
            '-------------------------------------------
            ' Recordset tipo tabla CnPoliza
            '-------------------------------------------
            vlstrx = " " & _
              "select INTNUMEROPOLIZA, TNYCLAVEEMPRESA, SMIEJERCICIO, TNYMES, INTCLAVEPOLIZA, DTMFECHAPOLIZA, CHRTIPOPOLIZA, " & _
              "  VCHCONCEPTOPOLIZA, SMICVEDEPARTAMENTO, INTCVEEMPLEADO, VCHNUMERO, BITASENTADA " & _
              " from CnPoliza where intNumeroPoliza = -1"
            Set rsCnPoliza = frsRegresaRs(vlstrx, adLockOptimistic, adOpenDynamic)
            
            '-------------------------------------------
            ' Recordset tipo tabla CnDetallePoliza
            '-------------------------------------------
            vlstrx = " " & _
              "select INTNUMEROPOLIZA, INTNUMEROCUENTA, BITNATURALEZAMOVIMIENTO, MNYCANTIDADMOVIMIENTO,VCHREFERENCIA, VCHCONCEPTO, INTNUMEROREGISTRO " & _
              " from CnDetallePoliza where intNumeroPoliza = -1"
            Set rsCnDetallePoliza = frsRegresaRs(vlstrx, adLockOptimistic, adOpenDynamic)
            
            With rsCnPoliza
                .AddNew
                !tnyclaveempresa = vgintClaveEmpresaContable
                !smiEjercicio = Year(vldtmFecha)
                !tnyMes = Month(vldtmFecha)
                !intClavePoliza = flngFolioPoliza(vgintClaveEmpresaContable, "D", Year(vldtmFecha), Month(vldtmFecha), False)
                !dtmFechaPoliza = vldtmFecha
                !chrTipoPoliza = "D"
                !vchConceptoPoliza = lstrConceptoPoliza
                !SMICVEDEPARTAMENTO = vlintCveDepartamento
                !intCveEmpleado = vllngPersonaGraba
                .Update
            End With
            
            vllngNumeroPolizaAjuste = flngObtieneIdentity("SEC_CNPOLIZA", rsCnPoliza!INTNUMEROPOLIZA)
            
            For X = 0 To UBound(apoliza(), 1)
                If apoliza(X).NumerodeCuenta <> 0 Then
                    With rsCnDetallePoliza
                        .AddNew
                        !INTNUMEROPOLIZA = vllngNumeroPolizaAjuste
                        !INTNUMEROCUENTA = apoliza(X).NumerodeCuenta
                        !bitNaturalezaMovimiento = apoliza(X).NaturalezaMovimiento
                        !mnyCantidadMovimiento = apoliza(X).CantidadMovimiento
                        !vchConcepto = " "
                        !vchReferencia = " "
                        .Update
                    End With
                End If
            Next X
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pRegistroPolizaAjuste"))
End Sub

Private Sub pActualizaCostoPromedioSalidaDevReubica(vldblNumReferencia As Double, vlstrCveArticulo As String, vldblCostoPromedio As Double)
    vlstrx = ""
    vlstrx = vlstrx & _
    "update IvDevolucionDeptoDetalle set " & _
        "mnyCostoPromedio = " & Str(vldblCostoPromedio) & " " & _
    "Where " & _
        "numNumDevolucion=" & Str(vldblNumReferencia) & " and " & _
        "chrCveArticulo='" & Trim(vlstrCveArticulo) & "' "
    pEjecutaSentencia vlstrx
End Sub



Private Function fdblCostoEntradaEDRA(vldblNumReubicacion As Double, vlstrCveArticulo As String) As Double
    
    vlstrx = ""
    vlstrx = vlstrx & _
    "select " & _
        "mnyCostoPromedio " & _
    "From " & _
        "IvDevolucionDeptoDetalle " & _
    "Where " & _
        "NumNumDevolucion =" & Str(vldblNumReubicacion) & " and " & _
        "chrCveArticulo = '" & Trim(vlstrCveArticulo) & "' "
        
    fdblCostoEntradaEDRA = frsRegresaRs(vlstrx).Fields(0)

End Function


Private Function fdblRecalculaCostoPromedioSal(vllngCantMovtoUV As Long, vllngCantMovtoUM As Long, vllngContenido As Long, vldblCostoEntrada As Double, vllngTotalUV As Long, vllngTotalUM As Long, vldblCostoPromedio As Double) As Double
    
    If ((vllngTotalUV + (vllngTotalUM / vllngContenido)) - (vllngCantMovtoUV + (vllngCantMovtoUM / vllngContenido))) <= 0 Then
        fdblRecalculaCostoPromedioSal = vldblCostoPromedio
        Exit Function
    End If
    
    fdblRecalculaCostoPromedioSal = _
    Round(( _
    (vllngTotalUV + (vllngTotalUM / vllngContenido)) * vldblCostoPromedio _
    - _
    (vllngCantMovtoUV + (vllngCantMovtoUM / vllngContenido)) * vldblCostoEntrada _
    ) / _
    ( _
    (vllngTotalUV + (vllngTotalUM / vllngContenido)) _
    - _
    (vllngCantMovtoUV + (vllngCantMovtoUM / vllngContenido)) _
    ), 4)

End Function


Private Sub pArticulosRecalculo(vlintCveDepto As Integer)

    Dim rsArticulosRecalculo As New ADODB.Recordset
    Dim X As Long
    Dim vllngRenglon As Long
    Dim vlblnHuboNegativos As Boolean

    grdArticulos.Cols = 9
    grdArticulos.Rows = 2
    
    For X = 0 To grdArticulos.Cols - 1
        grdArticulos.TextMatrix(1, X) = ""
    Next X
    
    vgstrParametrosSP = vlstrCveArticulo & "|" & Trim(Str(vlintCveDepto)) & "|" & vllngKardexMovimiento
    Set rsArticulosRecalculo = frsEjecuta_SP(vgstrParametrosSP, "SP_IVSELARTICULORECALCULO")
    If rsArticulosRecalculo.RecordCount <> 0 Then
        rsArticulosRecalculo.MoveFirst
        Do While Not rsArticulosRecalculo.EOF
            If Not fblnEstaArticulo(rsArticulosRecalculo!CHRCVEARTICULO) Then
                If Trim(grdArticulos.TextMatrix(1, 1)) = "" Then
                    vllngRenglon = 1
                Else
                    grdArticulos.Rows = grdArticulos.Rows + 1
                    vllngRenglon = grdArticulos.Rows - 1
                End If
                
                X = 0
                Do While X <= rsArticulosRecalculo.Fields.Count - 1
                    grdArticulos.TextMatrix(vllngRenglon, X + 1) = rsArticulosRecalculo.Fields(X)
                    X = X + 1
                Loop
            End If
            rsArticulosRecalculo.MoveNext
        Loop
    End If
    
    With grdArticulos
        .FormatString = "|Movimiento|Fecha|Referencia|Consecutivo|Clave|Artículo"
        .ColWidth(0) = 350
        .ColWidth(1) = 0
        .ColWidth(2) = 1000
        .ColWidth(3) = 1800
        .ColWidth(4) = 1000
        .ColWidth(5) = 1000
        .ColWidth(6) = 3900
        .ColWidth(7) = 0
        .ColWidth(8) = 0
        .ColAlignment(1) = flexAlignRightCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignLeftCenter
        .ColAlignment(4) = flexAlignRightCenter
        .ColAlignment(5) = flexAlignLeftCenter
        .ColAlignment(6) = flexAlignLeftCenter
        .ColAlignmentFixed(0) = flexAlignCenterCenter
        .ColAlignmentFixed(1) = flexAlignCenterCenter
        .ColAlignmentFixed(2) = flexAlignCenterCenter
        .ColAlignmentFixed(3) = flexAlignCenterCenter
        .ColAlignmentFixed(4) = flexAlignCenterCenter
        .ColAlignmentFixed(5) = flexAlignCenterCenter
        .ColAlignmentFixed(6) = flexAlignCenterCenter
    End With
    
    'Marca mara recalculo todos los articulos del grid
    pSelecciona

End Sub

Private Function fblnEstaArticulo(vlstrCveArticulo As String) As Boolean
    Dim X As Long
    Dim vlblnSalir As Boolean
    
    fblnEstaArticulo = False
    
    If Trim(grdArticulos.TextMatrix(1, 1)) <> "" Then
        X = 1
        vlblnSalir = False
        
        Do While X <= grdArticulos.Rows - 1 And Not vlblnSalir
            If Trim(grdArticulos.TextMatrix(X, 5)) = Trim(vlstrCveArticulo) Then
                fblnEstaArticulo = True
                vlblnSalir = True
            End If
            X = X + 1
        Loop
    End If
End Function

Private Sub cboAlmacen_Click()
    pArticulosRecalculo cboAlmacen.ItemData(cboAlmacen.ListIndex)
End Sub

Public Sub cmdRecalculo_Click()
    On Error GoTo NotificaError
    Dim vllngMovtoIniPgb As Long
    Dim vllngMovtoFinPgb As Long
    Dim X As Long
    Dim vldblCostoPromedio As Double
    Dim vldblCostoEntrada As Double
    Dim vlstrNombreDepartamento As String
    Dim rsIvSelCostoPromedioRecalculo As New ADODB.Recordset
    Dim rsIvSelExistenciaRecalculo As New ADODB.Recordset
    Dim vlstrTablaRelacion As String
    Dim strSentencia As String
    Dim lngTotalMovtos As Long
    Dim vldbldiferenciasum As Double
    Dim vldbldiferenciasuv As Double
    Dim rsAux As New ADODB.Recordset
    Dim vllngnoerror As Long
    Dim vldblNumRecepcion As Double
    Dim rs As ADODB.Recordset
    
    vllngTotalUM = 0
    vllngTotalUV = 0
    
    pgbProceso.Value = 0
    
    pCargaArregloMovtos
    
    X = 0
    vlblnnegativos = False
    
    Do While X <= UBound(aMovtoKardex())
    
        pgbProceso.Value = 0.5
    
        '------------------------------------------------------------
        'Iniciar las variables de costo promedio y totales de UV y UM para recalcular.
        
        'Costo promedio
        vgstrParametrosSP = aMovtoKardex(X).CveDepto & "|" & aMovtoKardex(X).clave & "|" & Format(aMovtoKardex(X).fecha, "YYYY-MM-DD HH:MM:SS") & "|" & aMovtoKardex(X).Movimiento
        Set rsAux = frsEjecuta_SP(vgstrParametrosSP, "SP_IvSelCostoPromedio")
        
        'Obtener costo promedio del movimiento anterior
        If rsAux.RecordCount <> 0 Then
            vldblCostoPromedio = rsAux!MNYCOSTOPROMEDIO
        Else
            'Asigna el costo promedio del movimiento actual (aplica para 'EAJ')
            vldblCostoPromedio = frsRegresaRs("select mnyCostoPromedio From IvKardexInventario Where intNumMovimiento=" & CLng(aMovtoKardex(X).Movimiento)).Fields(0)
            vlstrTablaRelacion = Trim(frsRegresaRs("select vchTablaRelacion From IvKardexInventario Where intNumMovimiento=" & CStr(CLng(aMovtoKardex(X).Movimiento)) & " ").Fields(0))
            
            If vlstrTablaRelacion = "EOE" Then
                vldblCostoPromedio = frsRegresaRs("select mnyCosto From IvOtrasEntSalDetalle Where numNumEntradaSalida=" & CLng(aMovtoKardex(X).Consecutivo) & " and chrCveArticulo='" & aMovtoKardex(X).clave & "' ").Fields(0)
            End If
            
            If vlstrTablaRelacion = "ERE" Then
                'Revisa la forma de calcular el costo promedio
                If intBitDescuentoCostoPromedio = 0 Then
                    'Costo sin incluir el descuento
                    vldblCostoPromedio = Round(frsRegresaRs("select mnyCostoEntrada From IvRecepcionDetalle Where numNumRecepcion=" & CLng(aMovtoKardex(X).Consecutivo) & " and chrCveArticulo='" & aMovtoKardex(X).clave & "' ").Fields(0), 4)
                Else
                    'Costo incluyendo el descuento
                    vldblCostoPromedio = Round(frsRegresaRs("select mnyCostoEntrada*(1-relDescuentoEnt/100) From IvRecepcionDetalle Where numNumRecepcion=" & CLng(aMovtoKardex(X).Consecutivo) & " and chrCveArticulo='" & aMovtoKardex(X).clave & "' ").Fields(0), 4)
                End If
            End If
            
            If vlstrTablaRelacion = "EBO" Then ' Bonificaciones
                vldblCostoPromedio = Round(frsRegresaRs("SELECT MNYCOSTOENTRADA FROM IVRECEPCIONBONIFICACION WHERE NUMNUMRECEPCION=" & CLng(aMovtoKardex(X).Consecutivo) & " and CHRCVEARTICULO='" & aMovtoKardex(X).clave & "' ").Fields(0), 4)
            End If
            If vlstrTablaRelacion = "ECO" Then ' Entrada de consignación
                'vldblCostoPromedio = 0
                Set rs = frsRegresaRs("SELECT mnyprecioproveedor FROM IvEntradaSalidaConsignaDetalle WHERE numEntradaSalidaConsignacion=" & CLng(aMovtoKardex(X).Consecutivo) & " and CHRCVEARTICULO='" & aMovtoKardex(X).clave & "' ", adLockOptimistic, adOpenForwardOnly)
                If rs.RecordCount > 0 Then
                    vldblCostoPromedio = Round(rs!mnyprecioproveedor, 4)
                End If
                rs.Close
            End If
            If vlstrTablaRelacion = "EFS" Then ' Entrada de consignación
                'vldblCostoPromedio = 0
                vlstrx = "SELECT COUNT(*) " & _
                        "FROM IVREQUISICIONDETALLESUB " & _
                        "WHERE IVREQUISICIONDETALLESUB.NUMNUMREQUISICION = " & CLng(aMovtoKardex(X).Consecutivo) & " AND IVREQUISICIONDETALLESUB.CHRCVEARTICULO = '" & aMovtoKardex(X).clave & "'"
                    
                Set rs = frsRegresaRs(vlstrx)
                If rs.RecordCount > 0 Then
                    vldblCostoPromedio = frsRegresaRs("select mnyCostoPromedio From IvKardexInventario Where intNumMovimiento=" & CLng(aMovtoKardex(X).Movimiento)).Fields(0)
                Else
                    rs.Close
                    vlstrx = "SELECT IVREUBICACIONDETALLE.MNYCOSTOPROMEDIO FROM IVREUBICACIONDETALLE " & _
                            "INNER JOIN IVREUBICACIONMAESTRO ON IVREUBICACIONDETALLE.NUMNUMREUBICACION = IVREUBICACIONMAESTRO.NUMNUMREUBICACION " & _
                            "WHERE IVREUBICACIONDETALLE.CHRCVEARTICULO = '" & aMovtoKardex(X).clave & "' AND IVREUBICACIONMAESTRO.NUMNUMREQUISICION = " & CLng(aMovtoKardex(X).Consecutivo)
                    Set rs = frsRegresaRs(vlstrx)
                    If rs.RecordCount > 0 Then
                        vldblCostoPromedio = rs!MNYCOSTOPROMEDIO
                    Else
                        rs.Close
                        vlstrx = "SELECT DISTINCT NVL(IVSALIDADEPTOMAESTRO.NUMNUMREQUISICION, 0) " & _
                                 "From IVSALIDADEPTOMAESTRO " & _
                                 "INNER JOIN IVSALIDADEPTODETALLE ON IVSALIDADEPTOMAESTRO.NUMNUMSALIDADEPTO = IVSALIDADEPTODETALLE.NUMNUMSALIDADEPTO " & _
                                 "Where IVSALIDADEPTOMAESTRO.NUMNUMSALIDADEPTO = " & CLng(aMovtoKardex(X).Consecutivo) & " And IVSALIDADEPTODETALLE.CHRCVEARTICULO = '" & aMovtoKardex(X).clave & "'"
                        Set rs = frsRegresaRs(vlstrx)
                        If rs.RecordCount > 0 Then
                            vldblCostoPromedio = rs!MNYCOSTOPROMEDIO
                        Else
                            rs.Close
                            vlstrx = "SELECT IVREUBICACIONDETALLE.MNYCOSTOPROMEDIO " & _
                                    "From IVREUBICACIONMAESTRO " & _
                                    "INNER JOIN IVREUBICACIONDETALLE ON IVREUBICACIONMAESTRO.NUMNUMREUBICACION = IVREUBICACIONDETALLE.NUMNUMREUBICACION " & _
                                    "INNER JOIN IVDEVOLUCIONDEPTOMAESTRO ON IVREUBICACIONMAESTRO.NUMNUMREUBICACION = IVDEVOLUCIONDEPTOMAESTRO.NUMNUMREFERENCIA " & _
                                    "Where IVDEVOLUCIONDEPTOMAESTRO.NUMNUMDEVOLUCION = " & CLng(aMovtoKardex(X).Consecutivo) & " And IVREUBICACIONDETALLE.CHRCVEARTICULO = '" & aMovtoKardex(X).clave & "'"
                            Set rs = frsRegresaRs(vlstrx)
                            If rs.RecordCount > 0 Then
                                vldblCostoPromedio = rs!MNYCOSTOPROMEDIO
                            Else
                                rs.Close
                                vlstrx = "SELECT IVSALIDADEPTODETALLE.MNYCOSTOPROMEDIO From IVSALIDADEPTOMAESTRO " & _
                                        "INNER JOIN IVSALIDADEPTODETALLE ON IVSALIDADEPTOMAESTRO.NUMNUMSALIDADEPTO = IVSALIDADEPTODETALLE.NUMNUMSALIDADEPTO " & _
                                        "INNER JOIN IVDEVOLUCIONDEPTOMAESTRO ON IVSALIDADEPTOMAESTRO.NUMNUMSALIDADEPTO = IVDEVOLUCIONDEPTOMAESTRO.NUMNUMREFERENCIA " & _
                                        "Where IVDEVOLUCIONDEPTOMAESTRO.NUMNUMDEVOLUCION = " & CLng(aMovtoKardex(X).Consecutivo) & " And IVSALIDADEPTODETALLE.CHRCVEARTICULO = '" & aMovtoKardex(X).clave & "'"
                                Set rs = frsRegresaRs(vlstrx)
                                If rs.RecordCount > 0 Then
                                    vldblCostoPromedio = rs!MNYCOSTOPROMEDIO
                                Else
                                    vldblCostoPromedio = 0
                                End If
                            End If
                        End If
                    End If
                End If
                rs.Close
              End If
        End If
        
        Set rsIvSelExistenciaRecalculo = frsEjecuta_SP(CLng(aMovtoKardex(X).Movimiento) & "|" & aMovtoKardex(X).CveDepto & "|" & aMovtoKardex(X).clave & "|" & Format(aMovtoKardex(X).fecha, "DD/MM/YYYY HH:MM:SS") & "|", "SP_IvSelExistenciaRecalculo")
        vllngTotalUV = rsIvSelExistenciaRecalculo!UV
        vllngTotalUM = rsIvSelExistenciaRecalculo!UM
    
        rsIvSelExistenciaRecalculo.Close
    
        '------------------------------------------------------------
        'Extraer los movimientos de kárdex que se afectarán
        '------------------------------------------------------------
        pMovimientosAfectados CDate(aMovtoKardex(X).fecha), aMovtoKardex(X).clave, aMovtoKardex(X).CveDepto
    
        rsMovimientosAfectados.MoveFirst
    
        Do While Not rsMovimientosAfectados.EOF And vllngTotalUV >= 0 And vllngTotalUM >= 0
        
            vllngCantMovtoUM = rsMovimientosAfectados!relCantidadUM
            vllngCantMovtoUV = rsMovimientosAfectados!relCantidadUV
        
            If fblnEsEntrada(Trim(rsMovimientosAfectados!vchTablaRelacion)) Then
                '----------------------------------------------------------
                ' E N T R A D A S
                '----------------------------------------------------------
                If vllngCantMovtoUM <> 0 Or vllngCantMovtoUV <> 0 Or Trim(rsMovimientosAfectados!vchTablaRelacion) = "EAJ" Then
                    If Trim(rsMovimientosAfectados!vchTablaRelacion) = "EAJ" Then
                        If vllngCantMovtoUM <> 0 Then
                            vldbldiferenciasum = rsMovimientosAfectados!RELEXISTENCIAUM - vllngTotalUM
                            vldbldiferenciasuv = 0
                        Else
                            vldbldiferenciasum = 0
                            vldbldiferenciasuv = rsMovimientosAfectados!RELEXISTENCIAUV - vllngTotalUV
                        End If
                
                        'Caso 13969
                        'Si el movimiento anterior tiene costo cero, se recalcula el costo promedio
                        If vldblCostoPromedio = 0 Then
                            vldblCostoEntrada = fdblCostoEntradaEAJ(rsMovimientosAfectados!numnumreferencia, aMovtoKardex(X).clave)
                        
                            vgstrParametrosSP = IIf(vldbldiferenciasum >= 0, vldbldiferenciasum, vldbldiferenciasum * -1) & "|" & IIf(vldbldiferenciasuv >= 0, vldbldiferenciasuv, vldbldiferenciasuv * -1) & "|" & _
                            rsMovimientosAfectados!INTNUMMOVIMIENTO & "|" & aMovtoKardex(X).clave & "|" & aMovtoKardex(X).CveDepto & "|" & vldblCostoEntrada & "|" & _
                            IIf((vldbldiferenciasum < 0 Or vldbldiferenciasuv < 0), "'SAJ'", "'EAJ'")
                                
                            frsEjecuta_SP vgstrParametrosSP, "SP_IVUPDKARDEXINVENTARIO"
                            
                            vldblCostoPromedio = fdblRecalculaCostoPromedio( _
                                                vllngCantMovtoUV, _
                                                vllngCantMovtoUM, _
                                                aMovtoKardex(X).Contenido, _
                                                vldblCostoEntrada, _
                                                vllngTotalUV, _
                                                vllngTotalUM, _
                                                vldblCostoPromedio)
                        Else
                            vgstrParametrosSP = IIf(vldbldiferenciasum >= 0, vldbldiferenciasum, vldbldiferenciasum * -1) & "|" & IIf(vldbldiferenciasuv >= 0, vldbldiferenciasuv, vldbldiferenciasuv * -1) & "|" & _
                            rsMovimientosAfectados!INTNUMMOVIMIENTO & "|" & aMovtoKardex(X).clave & "|" & aMovtoKardex(X).CveDepto & "|" & vldblCostoPromedio & "|" & _
                            IIf((vldbldiferenciasum < 0 Or vldbldiferenciasuv < 0), "'SAJ'", "'EAJ'")
                            
                            frsEjecuta_SP vgstrParametrosSP, "SP_IVUPDKARDEXINVENTARIO"
                        End If
                
                        '---------------------------------- ----------------
                        'AJUSTE DE INVENTARIO (los saldos en unidad minima y alterna se reinician según la captura física.
                        '--------------------------------------------------
                        If vllngCantMovtoUM <> 0 Then
                            vllngTotalUM = rsMovimientosAfectados!RELEXISTENCIAUM
                            vllngTotalUV = vllngTotalUV + vllngCantMovtoUV
                        ElseIf vllngCantMovtoUV <> 0 Then
                            vllngTotalUM = vllngTotalUM + vllngCantMovtoUM
                            vllngTotalUV = rsMovimientosAfectados!RELEXISTENCIAUV
                        Else
                            vllngTotalUM = rsMovimientosAfectados!RELEXISTENCIAUM
                            vllngTotalUV = rsMovimientosAfectados!RELEXISTENCIAUV
                        End If
                    Else
                        '--------------------------------------------------
                        'OTRAS ENTRADAS
                        '--------------------------------------------------
                        If Trim(rsMovimientosAfectados!vchTablaRelacion) = "EOE" Then
                    
                            pCostoPromedioEntradaSalida rsMovimientosAfectados!numnumreferencia, aMovtoKardex(X).clave, vldblCostoPromedio
                        
                            'La entrada afecta el costo promedio?
                            If fblnEntradaAfectaCostoPromedio(rsMovimientosAfectados!numnumreferencia, aMovtoKardex(X).clave, vldblCostoPromedio) Then
                                'Recalcular costo promedio
                                vldblCostoEntrada = fdblCostoEntradaEOE(rsMovimientosAfectados!numnumreferencia, aMovtoKardex(X).clave)

                                vldblCostoPromedio = fdblRecalculaCostoPromedio( _
                                                    vllngCantMovtoUV, _
                                                    vllngCantMovtoUM, _
                                                    aMovtoKardex(X).Contenido, _
                                                    vldblCostoEntrada, _
                                                    vllngTotalUV, _
                                                    vllngTotalUM, _
                                                    vldblCostoPromedio)
                            
                                pActualizaCostoPromedioSalida rsMovimientosAfectados!numnumreferencia, aMovtoKardex(X).clave, vldblCostoPromedio
                            End If
                        End If
                        '--------------------------------------------------
                        'RECEPCIONES
                        '--------------------------------------------------
                        If Trim(rsMovimientosAfectados!vchTablaRelacion) = "ERE" Then
                            vldblCostoEntrada = fdblCostoEntradaERE(rsMovimientosAfectados!numnumreferencia, aMovtoKardex(X).clave)
                            vldblCostoPromedio = fdblRecalculaCostoPromedio( _
                                                vllngCantMovtoUV, _
                                                vllngCantMovtoUM, _
                                                aMovtoKardex(X).Contenido, _
                                                vldblCostoEntrada, _
                                                vllngTotalUV, _
                                                vllngTotalUM, _
                                                vldblCostoPromedio)
                        End If
                        '--------------------------------------------------
                        'BONIFICACION
                        '--------------------------------------------------
                        If Trim(rsMovimientosAfectados!vchTablaRelacion) = "EBO" Then
                            'La Bonificación afecta el costo promedio?
                            'If fblnBonificacionAfectaCostoPromedio(rsMovimientosAfectados!numnumreferencia, aMovtoKardex(X).clave, vldblCostoPromedio) Then
                            '    'Recalcular costo promedio
                            '    vldblCostoEntrada = fdblCostoEntradaEBO(rsMovimientosAfectados!numnumreferencia, aMovtoKardex(X).clave)
                            '
                            '    If vllngCantMovtoUM <> 0 Then
                            '        'Fue una entrada minima, por tanto debe elevarse el costo de entrada de unidad minima a alterna
                            '        vldblCostoEntrada = vldblCostoEntrada * aMovtoKardex(X).Contenido
                            '    End If
                                
                            If rsMovimientosAfectados!bitBonificacionCostoCero = 1 And (vllngTotalUV > 0 Or vllngTotalUM > 0) Then
                                vldblCostoEntrada = 0
                            Else
                                vldblCostoEntrada = vldblCostoPromedio
                            End If
                            vldblCostoPromedio = fdblRecalculaCostoPromedio( _
                                                vllngCantMovtoUV, _
                                                vllngCantMovtoUM, _
                                                aMovtoKardex(X).Contenido, _
                                                vldblCostoEntrada, _
                                                vllngTotalUV, _
                                                vllngTotalUM, _
                                                vldblCostoPromedio)
                                                
                            vlstrx = "update IvKardexInventario set bitBonificacionCostoCero = " & IIf(vldblCostoEntrada = 0, 1, 0) & " where intNumMovimiento=" & Str(rsMovimientosAfectados!INTNUMMOVIMIENTO)
                            pEjecutaSentencia vlstrx
                            'End If
                        End If
                    
                        '--------------------------------------------------
                        'REUBICACION
                        '--------------------------------------------------
                        If Trim(rsMovimientosAfectados!vchTablaRelacion) = "ERA" Then
                            vldblCostoEntrada = fdblCostoEntradaERA(rsMovimientosAfectados!numnumreferencia, aMovtoKardex(X).clave)
                            vldblCostoPromedio = fdblRecalculaCostoPromedio( _
                                                vllngCantMovtoUV, _
                                                vllngCantMovtoUM, _
                                                aMovtoKardex(X).Contenido, _
                                                vldblCostoEntrada, _
                                                vllngTotalUV, _
                                                vllngTotalUM, _
                                                vldblCostoPromedio)
                        End If
    
                        '--------------------------------------------------
                        'DEVOLUCIÓN REUBICACION
                        '--------------------------------------------------
                        If Trim(rsMovimientosAfectados!vchTablaRelacion) = "EDRA" Or Trim(rsMovimientosAfectados!vchTablaRelacion) = "EDSD" Then
                            
                            vldblCostoEntrada = fdblCostoEntradaEDRA(rsMovimientosAfectados!numnumreferencia, aMovtoKardex(X).clave)
                            
                            vldblCostoPromedio = fdblRecalculaCostoPromedio( _
                                                vllngCantMovtoUV, _
                                                vllngCantMovtoUM, _
                                                aMovtoKardex(X).Contenido, _
                                                vldblCostoEntrada, _
                                                vllngTotalUV, _
                                                vllngTotalUM, _
                                                vldblCostoPromedio)
                        End If
                    
                        '--------------------------------------------------
                        ' CONSIGNACIÓN
                        '--------------------------------------------------
                        If Trim(rsMovimientosAfectados!vchTablaRelacion) = "ECO" Then
                            
                            vldblCostoEntrada = fdblCostoEntradaECO(rsMovimientosAfectados!numnumreferencia, aMovtoKardex(X).clave, aMovtoKardex(X).CveDepto)
                            
                            vldblCostoPromedio = fdblRecalculaCostoPromedio( _
                                                vllngCantMovtoUV, _
                                                vllngCantMovtoUM, _
                                                aMovtoKardex(X).Contenido, _
                                                vldblCostoEntrada, _
                                                vllngTotalUV, _
                                                vllngTotalUM, _
                                                vldblCostoPromedio)
                        End If
                        
                        '--------------------------------------------------
                        ' ENTRADA POR FARMACIA SUBROGADA
                        '--------------------------------------------------
                        If Trim(rsMovimientosAfectados!vchTablaRelacion) = "EFS" Then
                            
                            vldblCostoEntrada = fdblCostoEntradaSalidaFS(rsMovimientosAfectados!numnumreferencia, aMovtoKardex(X).clave, aMovtoKardex(X).CveDepto, aMovtoKardex(X).Movimiento)
                            
                            vldblCostoPromedio = fdblRecalculaCostoPromedio( _
                                                vllngCantMovtoUV, _
                                                vllngCantMovtoUM, _
                                                aMovtoKardex(X).Contenido, _
                                                vldblCostoEntrada, _
                                                vllngTotalUV, _
                                                vllngTotalUM, _
                                                vldblCostoPromedio)
                        End If
                    
                        vllngTotalUM = vllngTotalUM + vllngCantMovtoUM
                        vllngTotalUV = vllngTotalUV + vllngCantMovtoUV
                    End If
                
                    pCorrigeKardex rsMovimientosAfectados!INTNUMMOVIMIENTO, vllngTotalUV, vllngTotalUM, vldblCostoPromedio
                Else
                    '--------------------------------------------------
                    'CANCELAR EL MOVIMIENTO
                    '--------------------------------------------------
                    '* Otra entrada
                    If Trim(rsMovimientosAfectados!vchTablaRelacion) = "EOE" Then
                        pBorraEntradaSalida rsMovimientosAfectados!numnumreferencia, aMovtoKardex(X).clave
                    End If
                    
                    '* Recepción
                    If Trim(rsMovimientosAfectados!vchTablaRelacion) = "ERE" Then
                        pBorraRecepcion rsMovimientosAfectados!numnumreferencia, aMovtoKardex(X).clave
                    End If
                    
                    pBorraKardex rsMovimientosAfectados!INTNUMMOVIMIENTO
                End If
            Else
            '----------------------------------------------------------
            ' S A L I D A S
            '----------------------------------------------------------
                If vllngCantMovtoUM <> 0 Or vllngCantMovtoUV <> 0 Or Trim(rsMovimientosAfectados!vchTablaRelacion) = "SAJ" Then
                    If Trim(rsMovimientosAfectados!vchTablaRelacion) = "SAJ" Then
                        If vllngCantMovtoUM <> 0 Then
                            vldbldiferenciasum = rsMovimientosAfectados!RELEXISTENCIAUM - vllngTotalUM
                            vldbldiferenciasuv = 0
                        Else
                            vldbldiferenciasum = 0
                            vldbldiferenciasuv = rsMovimientosAfectados!RELEXISTENCIAUV - vllngTotalUV
                        End If
                    
                        vgstrParametrosSP = IIf(vldbldiferenciasum >= 0, vldbldiferenciasum, vldbldiferenciasum * -1) & "|" & IIf(vldbldiferenciasuv >= 0, vldbldiferenciasuv, vldbldiferenciasuv * -1) & "|" & rsMovimientosAfectados!INTNUMMOVIMIENTO & "|" & aMovtoKardex(X).clave & "|" & aMovtoKardex(X).CveDepto & "|" & vldblCostoPromedio & "|" & IIf((vldbldiferenciasum < 0 Or vldbldiferenciasuv < 0), "'SAJ'", "'EAJ'")
                        frsEjecuta_SP vgstrParametrosSP, "SP_IVUPDKARDEXINVENTARIO"
                        '---------------------------------- ----------------
                        'AJUSTE DE INVENTARIO (los saldos en unidad minima y alterna se reinician según la captura física.
                        '--------------------------------------------------
                        If vllngCantMovtoUM <> 0 Then
                           vllngTotalUM = rsMovimientosAfectados!RELEXISTENCIAUM
                           vllngTotalUV = vllngTotalUV
                        ElseIf vllngCantMovtoUV <> 0 Then
                           vllngTotalUM = vllngTotalUM
                           vllngTotalUV = rsMovimientosAfectados!RELEXISTENCIAUV
                        Else
                            vllngTotalUM = rsMovimientosAfectados!RELEXISTENCIAUM
                            vllngTotalUV = rsMovimientosAfectados!RELEXISTENCIAUV
                        End If
                    Else
                        '--------------------------------------------------
                        ' DEVOLUCIÓN A PROVEEDOR
                        '--------------------------------------------------
                        If Trim(rsMovimientosAfectados!vchTablaRelacion) = "SDPR" Then
                            'Obtiene el numero de recepción por la cual se hace la devolución
                            vldblNumRecepcion = frsRegresaRs("select numnumrecepcion From IvDevoluProveedMaestro Where intnumdevolucion = " & rsMovimientosAfectados!numnumreferencia).Fields(0)
                            vldblCostoEntrada = fdblCostoEntradaERE_SD(vldblNumRecepcion, aMovtoKardex(X).clave)
                        
                            'Recalcular costo promedio con el costo de entrada de la recepción por la cual se hace la devolución
'                            vldblCostoPromedio = fdblRecalculaCostoPromedioSal( _
'                                                 vllngCantMovtoUV, _
'                                                 vllngCantMovtoUM, _
'                                                 aMovtoKardex(X).Contenido, _
'                                                 vldblCostoEntrada, _
'                                                 vllngTotalUV, _
'                                                 vllngTotalUM, _
'                                                 vldblCostoPromedio)
                            pRegistroPolizaAjuste rsMovimientosAfectados!SMICVEDEPARTAMENTO, rsMovimientosAfectados!CHRCVEARTICULO, rsMovimientosAfectados!relCantidadUV, vldblCostoPromedio, rsMovimientosAfectados!numnumreferencia, vglngNumeroEmpleado, 0, rsMovimientosAfectados!DTMFECHAHORAMOV, 0, 0, 0, "NOTA", vldblCostoEntrada
                            If vgblnErrorIngreso = True Then
                                EntornoSIHO.ConeccionSIHO.RollbackTrans
                                If aMovtoKardex(X).PosicionGrid <> 0 Then
                                    If cboAlmacen.ItemData(cboAlmacen.ListIndex) = aMovtoKardex(X).CveDepto Then
                                        grdArticulos.TextMatrix(aMovtoKardex(X).PosicionGrid, 0) = "X"
                                        grdArticulos.Refresh
                                    End If
                                End If
                                pgbProceso.Visible = False
                                Exit Sub
                            End If
                        End If
                        
                        pRestaMovimiento aMovtoKardex(X).Contenido
                    End If
                
                    If Trim(rsMovimientosAfectados!vchTablaRelacion) = "SOS" Then
                        '--------------------------------------------------
                        'OTRAS SALIDAS
                        '--------------------------------------------------
                        pActualizaCostoPromedioSalida rsMovimientosAfectados!numnumreferencia, aMovtoKardex(X).clave, vldblCostoPromedio
                    End If
                    
                    If Trim(rsMovimientosAfectados!vchTablaRelacion) = "SRA" Then
                        '--------------------------------------------------
                        'REUBICACION
                        '--------------------------------------------------
                        pActualizaCostoPromedioSalidaReubica rsMovimientosAfectados!numnumreferencia, aMovtoKardex(X).clave, vldblCostoPromedio
                    End If
                    
                    If Trim(rsMovimientosAfectados!vchTablaRelacion) = "SD" Then
                        '--------------------------------------------------
                        'SALIDA A DEPARTAMENTO
                        '--------------------------------------------------
                        pActualizaCostoPromedioSalidaDepartamento rsMovimientosAfectados!numnumreferencia, aMovtoKardex(X).clave, vldblCostoPromedio
                    End If
                    If Trim(rsMovimientosAfectados!vchTablaRelacion) = "SDRA" Then
                        '--------------------------------------------------
                        'DEVOLUCIÓN REUBICACION
                        '--------------------------------------------------
                        pActualizaCostoPromedioSalidaDevReubica rsMovimientosAfectados!numnumreferencia, aMovtoKardex(X).clave, vldblCostoPromedio
                    End If
                
                    pCorrigeKardex rsMovimientosAfectados!INTNUMMOVIMIENTO, vllngTotalUV, vllngTotalUM, vldblCostoPromedio
                
                    'Si es una salida por reubicación entonces se marca el movimiento de entrada del almacen para recalcular en el almacen reubicado
                    If Trim(rsMovimientosAfectados!vchTablaRelacion) = "SRA" Or Trim(rsMovimientosAfectados!vchTablaRelacion) = "SDRA" Then
                        pIncluyeMovtoArreglo _
                        rsMovimientosAfectados!DTMFECHAHORAMOV, _
                        aMovtoKardex(X).Referencia, _
                        rsMovimientosAfectados!numnumreferencia, _
                        aMovtoKardex(X).clave, _
                        aMovtoKardex(X).Articulo, _
                        aMovtoKardex(X).Contenido, _
                        Trim(rsMovimientosAfectados!vchTablaRelacion), _
                        X
                    End If
                Else
                    '--------------------------------------------------
                    'CANCELAR EL MOVIMIENTO
                    '--------------------------------------------------
                    '* Otra entrada
                    If Trim(rsMovimientosAfectados!vchTablaRelacion) = "EOE" Then
                        pBorraEntradaSalida rsMovimientosAfectados!numnumreferencia, aMovtoKardex(X).clave
                    End If
                
                    pBorraKardex rsMovimientosAfectados!INTNUMMOVIMIENTO
                End If
            End If
        
            pgbProceso.Value = rsMovimientosAfectados.Bookmark / rsMovimientosAfectados.RecordCount * 100
        
            rsMovimientosAfectados.MoveNext
        Loop
    
        pActualizaUbicacion aMovtoKardex(X).clave, aMovtoKardex(X).CveDepto, vldblCostoPromedio, vllngTotalUV, vllngTotalUM
    
        If vllngTotalUV >= 0 And vllngTotalUM >= 0 Then
            If aMovtoKardex(X).PosicionGrid <> 0 Then
                grdArticulos.TextMatrix(aMovtoKardex(X).PosicionGrid, 0) = "Ok"
                grdArticulos.Refresh
            End If
        Else
            If aMovtoKardex(X).PosicionGrid <> 0 Then
                If cboAlmacen.ItemData(cboAlmacen.ListIndex) = aMovtoKardex(X).CveDepto Then
                    grdArticulos.TextMatrix(aMovtoKardex(X).PosicionGrid, 0) = "X"
                    grdArticulos.Refresh
                End If
            End If
            
            vlstrNombreDepartamento = frsRegresaRs("select vchDescripcion from NoDepartamento where smiCveDepartamento=" & Str(aMovtoKardex(X).CveDepto)).Fields(0)
            vlblnnegativos = True
        End If
        X = X + 1
    Loop
    
    Me.Visible = False
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdRecalculo_Click"))
End Sub

Private Sub pActualizaCostoPromedioSalidaReubica(vldblNumReferencia As Double, vlstrCveArticulo As String, vldblCostoPromedio As Double)
    vlstrx = ""
    vlstrx = vlstrx & _
    "update IvReubicacionDetalle set " & _
    "mnyCostoPromedio = " & Str(vldblCostoPromedio) & " " & _
    "Where " & _
    "numNumReubicacion=" & Str(vldblNumReferencia) & " and " & _
    "chrCveArticulo='" & Trim(vlstrCveArticulo) & "' "
    pEjecutaSentencia vlstrx
End Sub

Private Sub pActualizaCostoPromedioSalidaDepartamento(vldblNumReferencia As Double, vlstrCveArticulo As String, vldblCostoPromedio As Double)
    vlstrx = ""
    vlstrx = vlstrx & _
    "update IvSalidaDeptoDetalle set " & _
        "mnyCostoPromedio = " & Str(vldblCostoPromedio) & " " & _
    "Where " & _
        "NumNumSalidaDepto=" & Str(vldblNumReferencia) & " and " & _
        "chrCveArticulo='" & Trim(vlstrCveArticulo) & "'"
    pEjecutaSentencia vlstrx
End Sub

Private Sub pActualizaCostoPromedioSalida(vldblNumReferencia As Double, vlstrCveArticulo As String, vldblCostoPromedio As Double)
    vlstrx = ""
    vlstrx = vlstrx & _
    "update IvOtrasEntSalDetalle set " & _
        "mnyCostoPromedio = " & Str(vldblCostoPromedio) & " " & _
    "Where " & _
        "numNumEntradaSalida=" & Str(vldblNumReferencia) & " and " & _
        "chrCveArticulo='" & Trim(vlstrCveArticulo) & "' "
    pEjecutaSentencia vlstrx
End Sub

Private Function fblnEsEntrada(vlstrTablaRelacion As String) As Boolean
On Error GoTo NotificaError

    Dim rsIvKardextiporelacion As New ADODB.Recordset
    Set rsIvKardextiporelacion = frsEjecuta_SP("", "sp_IvSelKardexTipoRelacion")

    fblnEsEntrada = False
    
    If rsIvKardextiporelacion.RecordCount <> 0 Then
        rsIvKardextiporelacion.MoveFirst
        Do While Not rsIvKardextiporelacion.EOF
            If vlstrTablaRelacion = rsIvKardextiporelacion!vchTablaRelacion Then
                fblnEsEntrada = True
            End If
            rsIvKardextiporelacion.MoveNext
        Loop
    End If
    rsIvKardextiporelacion.Close
        'vlstrTablaRelacion = "EII" Or _
        'vlstrTablaRelacion = "EAJ" Or _
        'vlstrTablaRelacion = "EOE" Or _
        'vlstrTablaRelacion = "ERE" Or _
        'vlstrTablaRelacion = "EDCP" Or _
        vlstrTablaRelacion = "EDCL" Or _
        vlstrTablaRelacion = "ECP" Or _
        'vlstrTablaRelacion = "ERA" Or _
        'vlstrTablaRelacion = "EDRA" Or _
        vlstrTablaRelacion = "EDRS" Or _
        vlstrTablaRelacion = "DPD" Or _
        'vlstrTablaRelacion = "EVP" Or _
        'vlstrTablaRelacion = "EHC" Or _
        'vlstrTablaRelacion = "EBO" Then
        
Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnEsEntrada"))
End Function

Private Sub pRestaMovimiento(vllngContenido As Long)
    Dim vllngTotalTempUV As Long
    Dim vllngTotalTempUM As Long

    If vllngCantMovtoUV <> 0 Then
        If vllngTotalUV >= vllngCantMovtoUV Then
            vllngTotalUV = vllngTotalUV - vllngCantMovtoUV
        Else
            If vllngTotalUV <> 0 Then
                vllngCantMovtoUV = vllngCantMovtoUV - vllngTotalUV
                vllngTotalUV = 0
            End If
            If vllngTotalUM / vllngContenido >= vllngCantMovtoUV Then
                vllngTotalUM = vllngTotalUM - vllngCantMovtoUV * vllngContenido
            Else
                'Se genera cantidad negativa
                vllngTotalUV = vllngTotalUV - vllngCantMovtoUV
            End If
        End If
    Else
        If vllngTotalUM >= vllngCantMovtoUM Then
            vllngTotalUM = vllngTotalUM - vllngCantMovtoUM
        Else
            If vllngTotalUM > 0 Then
                vllngCantMovtoUM = vllngCantMovtoUM - vllngTotalUM
                vllngTotalUM = 0
            End If
            If vllngTotalUV * vllngContenido >= vllngCantMovtoUM Then
                vllngTotalTempUV = Int((vllngTotalUV * vllngContenido - vllngCantMovtoUM) / vllngContenido)
                vllngTotalTempUM = _
                ( _
                (vllngTotalUV * vllngContenido - vllngCantMovtoUM) / vllngContenido _
                - _
                Int((vllngTotalUV * vllngContenido - vllngCantMovtoUM) / vllngContenido) _
                ) _
                * _
                vllngContenido
                
                vllngTotalUM = vllngTotalTempUM
                vllngTotalUV = vllngTotalTempUV
            Else
                'Se genera cantidad negativa
                vllngTotalUM = vllngTotalUM - vllngCantMovtoUM
            End If
        End If
    End If

End Sub

Private Sub pActualizaUbicacion(vlstrCveArticulo As String, vlintCveDepto As Integer, vldblCostoPromedio As Double, vllngUV As Long, vllngUM As Long)
    vlstrx = ""
    vlstrx = vlstrx & _
    "update IvUbicacion set " & _
        "mnyCostoPromedio=" & Str(vldblCostoPromedio) & "," & _
        "intExistenciaDeptoUV=" & Str(vllngUV) & "," & _
        "intExistenciaDeptoUM=" & Str(vllngUM) & " " & _
    "Where " & _
        "smiCveDepartamento=" & Str(vlintCveDepto) & " and " & _
        "chrCveArticulo='" & Trim(vlstrCveArticulo) & "' "
    pEjecutaSentencia vlstrx
End Sub

Private Sub pIncluyeMovtoArreglo(ByVal vldtmFecha As Date, ByVal vlstrNumReferencia As String, ByVal vldblConsecutivo As Double, ByVal vlstrCveArticulo As String, ByVal vlstrArticulo As String, ByVal vllngContenido As Long, vlstrTipoMovimiento As String, vllngRegActualArreglo As Long)
    
    Dim rsMovimientoEntrada As New ADODB.Recordset
    Dim vllngNumArticulos As Long
    Dim vllngCiclo As Long
    Dim vlblnEstaMovto As Boolean
    Dim vlintTamaño As Long
    
    If vlstrTipoMovimiento = "SRA" Then
        vlstrx = ""
        vlstrx = vlstrx & _
        "select " & _
            "intNumMovimiento," & _
            "smiCveDepartamento " & _
        "From " & _
            "IvKardexInventario " & _
        "Where " & _
            "numNumReferencia=" & Str(vldblConsecutivo) & " and " & _
            "chrCveArticulo='" & Trim(vlstrCveArticulo) & "' and " & _
            "vchTablaRelacion='ERA' "
    ElseIf vlstrTipoMovimiento = "SDRA" Then
        vlstrx = ""
        vlstrx = vlstrx & _
        "select " & _
            "intNumMovimiento," & _
            "smiCveDepartamento " & _
        "From " & _
            "IvKardexInventario " & _
        "Where " & _
            "numNumReferencia=" & Str(vldblConsecutivo) & " and " & _
            "chrCveArticulo='" & Trim(vlstrCveArticulo) & "' and " & _
            "vchTablaRelacion='EDRA' "
    End If
    Set rsMovimientoEntrada = frsRegresaRs(vlstrx)
    
    'Revisar si el movimiento ya existe en el arreglo de los registros del arreglo que faltan por revisar para que no los agregue repetidos y se cicle el proceso
    vlblnEstaMovto = False
    vlintTamaño = UBound(aMovtoKardex(), 1)
    For vllngCiclo = vllngRegActualArreglo To vlintTamaño
        If aMovtoKardex(vllngCiclo).Movimiento = rsMovimientoEntrada!INTNUMMOVIMIENTO _
            And aMovtoKardex(vllngCiclo).fecha = vldtmFecha _
            And aMovtoKardex(vllngCiclo).Referencia = vlstrNumReferencia _
            And aMovtoKardex(vllngCiclo).Consecutivo = vldblConsecutivo _
            And aMovtoKardex(vllngCiclo).clave = vlstrCveArticulo _
            And aMovtoKardex(vllngCiclo).Articulo = vlstrArticulo _
            And aMovtoKardex(vllngCiclo).Contenido = vllngContenido _
            And aMovtoKardex(vllngCiclo).CveDepto = rsMovimientoEntrada!SMICVEDEPARTAMENTO Then
                vlblnEstaMovto = True
                Exit For
        End If
    Next vllngCiclo
            
    If Not vlblnEstaMovto Then
        vllngNumArticulos = UBound(aMovtoKardex()) + 1
        
        ReDim Preserve aMovtoKardex(vllngNumArticulos) As MovtosKardex
        
        aMovtoKardex(vllngNumArticulos).Movimiento = rsMovimientoEntrada!INTNUMMOVIMIENTO
        aMovtoKardex(vllngNumArticulos).fecha = vldtmFecha
        aMovtoKardex(vllngNumArticulos).Referencia = vlstrNumReferencia
        aMovtoKardex(vllngNumArticulos).Consecutivo = vldblConsecutivo
        aMovtoKardex(vllngNumArticulos).clave = vlstrCveArticulo
        aMovtoKardex(vllngNumArticulos).Articulo = vlstrArticulo
        aMovtoKardex(vllngNumArticulos).Contenido = vllngContenido
        aMovtoKardex(vllngNumArticulos).CveDepto = rsMovimientoEntrada!SMICVEDEPARTAMENTO
        aMovtoKardex(vllngNumArticulos).PosicionGrid = 0
    
        vlstrx = "update IvKardexInventario set bitRecalculo=1 where intNumMovimiento=" & Str(rsMovimientoEntrada!INTNUMMOVIMIENTO)
        pEjecutaSentencia vlstrx
    End If
    
End Sub

Private Sub pCostoPromedioEntradaSalida(vldblNumEntradaSalida As Double, vlstrCveArticulo As String, vldblCostoPromedio As Double)
    'Actualizar el costo promedio que estaba antes de esta entrada

    vlstrx = ""
    vlstrx = vlstrx & _
    "update IvOtrasEntSalDetalle set " & _
        "mnyCostoPromedio=" & Str(vldblCostoPromedio) & " " & _
    "Where " & _
        "numNumEntradaSalida=" & Str(vldblNumEntradaSalida) & " and " & _
        "chrCveArticulo='" & Trim(vlstrCveArticulo) & "' "
        
    pEjecutaSentencia vlstrx
End Sub

Private Sub pBorraKardex(vllngNumMovimiento As Long)
    vlstrx = ""
    vlstrx = vlstrx & _
    "Delete FROM " & _
        "IvKardexInventario " & _
    "Where " & _
        "intNumMovimiento=" & Str(vllngNumMovimiento) & " "

    pEjecutaSentencia vlstrx
End Sub

Private Sub pBorraEntradaSalida(vldblNumEntradaSalida As Double, vlstrCveArticulo As String)
    vlstrx = _
    "Delete FROM " & _
        "IvOtrasEntSalDetalle " & _
    "Where " & _
        "numNumEntradaSalida=" & Str(vldblNumEntradaSalida) & " and " & _
        "chrCveArticulo='" & Trim(vlstrCveArticulo) & "' " & _
    " "
    pEjecutaSentencia vlstrx
    
    vlstrx = _
            "Delete FROM " & _
                "IvOtrasEntSalMaestro " & _
            "Where " & _
                "numNumEntradaSalida=" & Str(vldblNumEntradaSalida) & " " & _
                " AND (SELECT COUNT(oesd.numNumEntradaSalida) FROM IvOtrasEntSalDetalle oesd" & _
                " WHERE oesd.numNumEntradaSalida=" & Str(vldblNumEntradaSalida) & " ) = 0"
    pEjecutaSentencia vlstrx
End Sub

Private Sub pBorraRecepcion(vldblNumRecepcion As Double, vlstrCveArticulo As String)
Dim rs As New ADODB.Recordset
Dim rsNormal As New ADODB.Recordset
Dim rsConSigna As New ADODB.Recordset
Dim rsRecepcionDetalleBorrable As ADODB.Recordset
    
    vlstrx = ""
    vlstrx = vlstrx & _
    "Delete FROM " & _
        "IvRecepcionDetalle " & _
    "Where " & _
        "numNumRecepcion=" & Str(vldblNumRecepcion) & " and " & _
        "chrCveArticulo='" & Trim(vlstrCveArticulo) & "' " & _
    " "
    pEjecutaSentencia vlstrx
       
'    vlstrx = "SELECT * FROM IVRECEPCIONMAESTRO " & _
'             "WHERE " & _
'                "numNumRecepcion=" & Str(vldblNumRecepcion) & _
'                " AND (SELECT COUNT(rd.numNumRecepcion) FROM IvRecepcionDetalle rd " & _
'                " WHERE rd.numNumRecepcion=" & Str(vldblNumRecepcion) & ")=0"
     vlstrx = "SELECT * FROM IVRECEPCIONMAESTRO " & _
              "WHERE " & _
                 "numNumRecepcion=" & Str(vldblNumRecepcion) & _
                 " AND (SELECT COUNT(rd.numNumRecepcion) FROM IvRecepcionDetalle rd " & _
                 " WHERE rd.numNumRecepcion=" & Str(vldblNumRecepcion) & " AND rd.SMICANTIDADRECEP <> 0)=0"

    Set rs = frsRegresaRs(vlstrx)
    If rs.RecordCount > 0 Then
        vlblnBorroRecepcion = True
    
        vllngCuentaPagarABorrar = 0
        vlstrx = "SELECT * FROM IVRECEPCIONMAESTROFACTURA WHERE intidrecepcion = " & Str(vldblNumRecepcion)
        Set rsNormal = frsRegresaRs(vlstrx)
        If rsNormal.RecordCount > 0 Then
            vllngCuentaPagarABorrar = rsNormal!intIdCuentaPagar
            pEjecutaSentencia "DELETE FROM IVRECEPCIONMAESTROFACTURA WHERE intidrecepcion = " & Str(vldblNumRecepcion)
        End If

        vllngPolizaConsignaBorrar = 0
        vlstrx = "SELECT * FROM IVRECEPCIONCONSIGNA WHERE numnumrecepcion = " & Str(vldblNumRecepcion)
        Set rsConSigna = frsRegresaRs(vlstrx)
        If rsConSigna.RecordCount > 0 Then
            vllngPolizaConsignaBorrar = rsConSigna!intNumPoliza
            pEjecutaSentencia "DELETE FROM IVRECEPCIONCONSIGNA WHERE numnumrecepcion = " & Str(vldblNumRecepcion)
        End If
        
        '--> BS
        'Relacion con poliza y factura de caja chica
        pEjecutaSentencia "DELETE FROM IVRECEPCIONFACTURACAJACHICA WHERE INTCVERECEPCION = " & Str(vldblNumRecepcion)
        '<--
        
        vlstrx = "SELECT * FROM IvRecepcionDetalle " & _
            "WHERE " & _
                "numNumRecepcion=" & Str(vldblNumRecepcion) & " and " & _
                "SMICANTIDADRECEP = 0"
        Set rsRecepcionDetalleBorrable = frsRegresaRs(vlstrx)
        If rsRecepcionDetalleBorrable.RecordCount = 0 Then
            vlstrx = ""
            vlstrx = vlstrx & _
                    "Delete FROM " & _
                        "IvRecepcionMaestro " & _
                    "Where " & _
                        "numNumRecepcion=" & Str(vldblNumRecepcion) & _
                        " AND (SELECT COUNT(rd.numNumRecepcion) FROM IvRecepcionDetalle rd " & _
                        " WHERE rd.numNumRecepcion=" & Str(vldblNumRecepcion) & ")=0"
            pEjecutaSentencia vlstrx
        End If
    End If
    
End Sub

Private Sub pCorrigeKardex(vllngNumMovimiento As Long, vllngTotalUV As Long, vllngTotalUM As Long, vldblCostoPromedio As Double)
    vlstrx = ""
    vlstrx = vlstrx & _
    "update IvKardexInventario set " & _
        "relExistenciaUV=" & Str(vllngTotalUV) & "," & _
        "relExistenciaUM=" & Str(vllngTotalUM) & "," & _
        "mnyCostoPromedio=" & Str(vldblCostoPromedio) & "," & _
        "bitRecalculo = 0 " & _
    "Where " & _
        "intNumMovimiento=" & Str(vllngNumMovimiento) & " "
    pEjecutaSentencia vlstrx
End Sub

Private Function fdblCostoEntradaERA(vldblNumReubicacion As Double, vlstrCveArticulo As String) As Double
'    vlstrx = ""
'    vlstrx = vlstrx & _
'    "select " & _
'        "mnyCostoPromedio " & _
'    "From " & _
'        "IvReubicacionDetalle " & _
'    "Where " & _
'        "NumNumReubicacion=" & Str(vldblNumReubicacion) & " and " & _
'        "chrCveArticulo='" & Trim(vlstrCveArticulo) & "' "
        
    vlstrx = ""
    vlstrx = vlstrx & _
    "select " & _
        "case when nodepartamento.bitconsignacion = 1 then " & _
        " ivreubicaciondetalle.mnyprecioproveedor " & _
        "Else " & _
        " ivreubicaciondetalle.mnycostopromedio " & _
        "end costoaux " & _
    "From " & _
        "ivreubicacionmaestro " & _
        "inner join ivreubicaciondetalle on ivreubicacionmaestro.numnumreubicacion = ivreubicaciondetalle.numnumreubicacion " & _
        "inner join nodepartamento on nodepartamento.smicvedepartamento = ivreubicacionmaestro.smicvealmacenreubica " & _
    "Where " & _
        "ivreubicaciondetalle.NumNumReubicacion=" & Str(vldblNumReubicacion) & " and " & _
        "ivreubicaciondetalle.chrCveArticulo='" & Trim(vlstrCveArticulo) & "' "
        
    fdblCostoEntradaERA = frsRegresaRs(vlstrx).Fields(0)
End Function

Private Function fdblCostoEntradaERE(vldblNumRecepcion As Double, vlstrCveArticulo As String) As Double
    vlstrx = ""
    
    'Revisa la forma de calcular el costo promedio
    If intBitDescuentoCostoPromedio = 0 Then
        'Costo sin incluir el descuento
        vlstrx = vlstrx & _
        "select " & _
            "Case intCostoMasIVa when 1 then " & _
                "((mnyCostoEntrada))*(1+relIvaEntrada/100)" & _
            "Else " & _
                "(mnyCostoEntrada) " & _
            "End as Costo " & _
        "From " & _
            "IvRecepcionDetalle " & _
        "Where " & _
            "numNumRecepcion=" & Str(vldblNumRecepcion) & " and " & _
            "chrCveArticulo='" & Trim(vlstrCveArticulo) & "' "
    Else
        'Costo incluyendo el descuento
        vlstrx = vlstrx & _
        "select " & _
            "Case intCostoMasIVa when 1 then " & _
                "((mnyCostoEntrada)*(1-relDescuentoEnt/100))*(1+relIvaEntrada/100)" & _
            "Else " & _
                "(mnyCostoEntrada)*(1-relDescuentoEnt/100) " & _
            "End as Costo " & _
        "From " & _
            "IvRecepcionDetalle " & _
        "Where " & _
            "numNumRecepcion=" & Str(vldblNumRecepcion) & " and " & _
            "chrCveArticulo='" & Trim(vlstrCveArticulo) & "' "
    End If
        
    fdblCostoEntradaERE = frsRegresaRs(vlstrx, adLockOptimistic, adOpenDynamic).Fields(0)
End Function

Private Function fdblCostoEntradaEOE(vldblNumEntradaSalida As Double, vlstrCveArticulo As String) As Double
    vlstrx = ""
    vlstrx = vlstrx & _
    "select " & _
        "IvOtrasEntSalDetalle.mnyCosto " & _
    "From " & _
        "IvOtrasEntSalDetalle " & _
    "Where " & _
        "IvOtrasEntSalDetalle.numNumEntradaSalida=" & Str(vldblNumEntradaSalida) & " and " & _
        "IvOtrasEntSalDetalle.chrCveArticulo='" & Trim(vlstrCveArticulo) & "' "

    fdblCostoEntradaEOE = frsRegresaRs(vlstrx).Fields(0)
End Function

Private Function fdblCostoEntradaEBO(vldblNumEntradaSalida As Double, vlstrCveArticulo As String) As Double
    vlstrx = ""
    vlstrx = vlstrx & _
    "select " & _
        "IVRECEPCIONBONIFICACION.MNYCOSTOENTRADA " & _
    "From " & _
        "IVRECEPCIONBONIFICACION " & _
    "Where " & _
        "IVRECEPCIONBONIFICACION.NUMNUMRECEPCION=" & Str(vldblNumEntradaSalida) & " and " & _
        "IVRECEPCIONBONIFICACION.chrCveArticulo='" & Trim(vlstrCveArticulo) & "' "

    fdblCostoEntradaEBO = frsRegresaRs(vlstrx).Fields(0)
End Function

Private Function fdblRecalculaCostoPromedio(vllngCantMovtoUV As Long, vllngCantMovtoUM As Long, vllngContenido As Long, vldblCostoEntrada As Double, vllngTotalUV As Long, vllngTotalUM As Long, vldblCostoPromedio As Double) As Double
    fdblRecalculaCostoPromedio = _
    Round(( _
    (vllngCantMovtoUV + (vllngCantMovtoUM / vllngContenido)) * vldblCostoEntrada _
    + _
    (vllngTotalUV + (vllngTotalUM / vllngContenido)) * vldblCostoPromedio _
    ) / _
    ( _
    vllngCantMovtoUV + (vllngCantMovtoUM / vllngContenido) _
    + _
    vllngTotalUV + (vllngTotalUM / vllngContenido) _
    ), 4)
End Function

Private Function fblnEntradaAfectaCostoPromedio(vldblNumReferencia As Double, vlstrCveArticulo As String, vldblCostoPromedio As Double) As Boolean
    Dim rs As New ADODB.Recordset

    vlstrx = ""
    vlstrx = vlstrx & _
    "select " & _
        "IvTipoOtrasEntSal.bitAfectaCtoProm, IvTipoOtrasEntSal.BITBONIFICACION, IvParametro.BITBONIFICACIONCOSTOCERO " & _
    "From " & _
        "IvOtrasEntSalMaestro " & _
        "inner join IvTipoOtrasEntSal on " & _
        "IvOtrasEntSalMaestro.smiCveTipoEntSal = IvTipoOtrasEntSal.smiCveTipoEntSal, IvParametro " & _
    "Where " & _
        "IvOtrasEntSalMaestro.numNumEntradaSalida=" & Str(vldblNumReferencia) & _
        " and ivparametro.tnyclaveempresa = " & vgintClaveEmpresaContable & " "
        
    fblnEntradaAfectaCostoPromedio = True
    Set rs = frsRegresaRs(vlstrx)
    If rs.State <> adStateClosed Then
        If rs.RecordCount > 0 Then
            rs.MoveFirst
            If rs!bitBonificacion = 1 Then
                If rs!bitBonificacionCostoCero = 1 Then
                    fblnEntradaAfectaCostoPromedio = True
                Else
                    fblnEntradaAfectaCostoPromedio = False
                    vlstrx = ""
                    vlstrx = vlstrx & _
                    "UPDATE IvOtrasEntSalDetalle SET " & _
                    "  mnyCosto=" & Str(vldblCostoPromedio) & " " & _
                    " WHERE " & _
                    "  numNumEntradaSalida=" & Str(vldblNumReferencia) & " and " & _
                    "  chrCveArticulo='" & Trim(vlstrCveArticulo) & "' "
                    pEjecutaSentencia vlstrx
                End If
            Else
                If rs!bitAfectaCtoProm = 1 Then
                    fblnEntradaAfectaCostoPromedio = True
                Else
                    fblnEntradaAfectaCostoPromedio = False
                    vlstrx = ""
                    vlstrx = vlstrx & _
                    "UPDATE IvOtrasEntSalDetalle SET " & _
                    "  mnyCosto=" & Str(vldblCostoPromedio) & " " & _
                    " WHERE " & _
                    "  numNumEntradaSalida=" & Str(vldblNumReferencia) & " and " & _
                    "  chrCveArticulo='" & Trim(vlstrCveArticulo) & "' "
                    pEjecutaSentencia vlstrx
                End If
            End If
        End If
        rs.Close
    End If

End Function

Private Function fblnBonificacionAfectaCostoPromedio(vldblNumReferencia As Double, vlstrCveArticulo As String, vldblCostoPromedio As Double) As Boolean
    
    fblnBonificacionAfectaCostoPromedio = True
    
    If gintBonificacionCostoCero = 0 Then
        fblnBonificacionAfectaCostoPromedio = False
        vlstrx = ""
        vlstrx = vlstrx & _
        "UPDATE IVRECEPCIONBONIFICACION SET " & _
        "  MNYCOSTOENTRADA=" & Str(vldblCostoPromedio) & " " & _
        " WHERE " & _
        "  NUMNUMRECEPCION=" & Str(vldblNumReferencia) & " and " & _
        "  CHRCVEARTICULO='" & Trim(vlstrCveArticulo) & "' "
        pEjecutaSentencia vlstrx
    End If
End Function

Private Sub pMovimientosAfectados(strFechaMovto As String, vlstrCveArticulo As String, vlintCveDepto As Integer)
On Error GoTo NotificaError

    vgstrParametrosSP = Format(strFechaMovto, "DD/MM/YYYY HH:MM:SS") & "|" & Trim(vlstrCveArticulo) & "|" & Str(vlintCveDepto)
    Set rsMovimientosAfectados = frsEjecuta_SP(vgstrParametrosSP, "sp_IvSelKardexInventario")
   
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pMovimientosAfectados"))
    Unload Me
End Sub

Private Sub pCargaArregloMovtos()
    Dim X As Long
    Dim a As Long
    
    ReDim aMovtoKardex(0) As MovtosKardex
    aMovtoKardex(0).Movimiento = 0
    
    a = 0
    For X = 1 To grdArticulos.Rows - 1
        If Trim(grdArticulos.TextMatrix(X, 0)) = "*" Then
            If aMovtoKardex(0).Movimiento = 0 Then
                a = 0
            Else
                ReDim Preserve aMovtoKardex(a) As MovtosKardex
            End If
            aMovtoKardex(a).Movimiento = Val(grdArticulos.TextMatrix(X, 1))
            aMovtoKardex(a).fecha = grdArticulos.TextMatrix(X, 2)
            aMovtoKardex(a).Referencia = grdArticulos.TextMatrix(X, 3)
            aMovtoKardex(a).Consecutivo = Val(grdArticulos.TextMatrix(X, 4))
            aMovtoKardex(a).clave = Trim(grdArticulos.TextMatrix(X, 5))
            aMovtoKardex(a).Articulo = Trim(grdArticulos.TextMatrix(X, 6))
            aMovtoKardex(a).Contenido = Val(grdArticulos.TextMatrix(X, 7))
            aMovtoKardex(a).CveDepto = Val(grdArticulos.TextMatrix(X, 8))
            aMovtoKardex(a).PosicionGrid = X
            a = a + 1
        End If
    Next X

End Sub

Private Sub Form_Activate()
    pgbProceso.Value = 0
    cmdRecalculo_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Me.Icon = frmMenuPrincipal.Icon
    
    intBitDescuentoCostoPromedio = fRegresaParametro("BITDESCUENTOCOSTOPROMEDIO", "IvParametro", 0)
    
    vlblnBorroRecepcion = False
    vllngCuentaPagarABorrar = 0
    vllngPolizaConsignaBorrar = 0
    
    pCargaAlmacenes
End Sub

Private Sub pCargaAlmacenes()
    Dim rsAlmacenes As New ADODB.Recordset
    
    cmdRecalculo.Enabled = False
    cboAlmacen.Clear
    grdArticulos.Clear
    
    Set rsAlmacenes = frsRegresaRs("SELECT vchDescripcion, smiCveDepartamento FROM NODEPARTAMENTO WHERE smiCveDepartamento = " & vlintNumeroDepto)
    If rsAlmacenes.RecordCount > 0 Then
        pLlenarCboRs cboAlmacen, rsAlmacenes, 1, 0
        cboAlmacen.ListIndex = 0
    End If
End Sub

Private Sub pSelecciona()
    Dim X As Long
    X = 1
    
    'Recorre el grid para marcar todo
    Do While X <= grdArticulos.Rows - 1
        If Trim(grdArticulos.TextMatrix(X, 0)) = "" Then
            grdArticulos.TextMatrix(X, 0) = "*"
            grdArticulos.CellFontBold = True
            cmdRecalculo.Enabled = True
        End If
        X = X + 1
    Loop
End Sub

