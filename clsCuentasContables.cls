VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCuentasContables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Devuelve las cuentas de : Conceptos de facturación, de su IVA y de su descuento, del Exedente en suma asegurada y de su iva, de las formas de pago y de su comision."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Enum enmTipoCuenta
    CuentaConceptoFacturacion = 1
    CuentaExcedenteSumaAsegurada = 2
    CuentaFormaPago = 3
End Enum

Public Sub ConsultaCuenta(ByVal TipoCuenta As enmTipoCuenta)
Attribute ConsultaCuenta.VB_Description = "Permite leer una cuenta contable de algun tipo"
Attribute ConsultaCuenta.VB_UserMemId = 0
    'ConsultaCuenta = 1
End Sub

