VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFacturaParcial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Public colFormasPago As Collection ' Para contener los arreglos (aFormasPago) que regresa frmFormasPago para facturar parcialmente
Public intIdFacturaParcial As Long
Public intCveClienteCredito As Long
Public strTipoClienteCredito As String
Public intCveDatosFiscales As Long
Public strDFNombre As String
Public strDFDireccion As String
Public strDFNumeroExterior As String
Public strDFNumeroInterior As String
Public strDFRFC As String
Public strDFTelefono As String
Public llngCveCiudad As Long
Public strDFColonia As String
Public strDFCP As String

Private Sub Class_Initialize()
    Set colFormasPago = New Collection
End Sub

Public Sub AgregaFormaPago(intNumFormaPago As Integer, dblCantidad As Double, strFolio As String, lngCuentaContable As Long, dblTipoCambio As Double, bolEsCredito As Boolean, dblDolares As Double, Optional lngIdBanco As Long = 0, Optional lngCuentaComisionBancaria As Long = 0, Optional dblCantidadComisionBancaria As Double = 0, Optional dblIvaComisionBancaria As Double = 0)
    Dim fpaFormaPago As New clsFormaPago
    
    With fpaFormaPago
        .bolEsCredito = bolEsCredito
        .dblCantidad = dblCantidad
        .dblDolares = dblDolares
        .dblTipoCambio = dblTipoCambio
        .intNumFormaPago = intNumFormaPago
        .lngCuentaContable = lngCuentaContable
        .strFolio = strFolio
        .lngIdBanco = lngIdBanco
        .lngCuentaComisionBancaria = lngCuentaComisionBancaria
        .dblCantidadComisionBancaria = dblCantidadComisionBancaria
        .dblIvaComisionBancaria = dblIvaComisionBancaria
    End With
    colFormasPago.Add fpaFormaPago
End Sub
