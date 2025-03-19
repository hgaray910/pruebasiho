VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFormaPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public intNumFormaPago As Integer
Public dblCantidad As Double
Public strFolio As String
Public dblComision As Double
Public lngCuentaContable As Long
Public dblTipoCambio As Double
Public bolEsCredito As Boolean
Public dblDolares As Double
Public lngIdBanco As Long
Public lngCuentaComisionBancaria As Long
Public dblCantidadComisionBancaria As Double
Public dblIvaComisionBancaria As Double
