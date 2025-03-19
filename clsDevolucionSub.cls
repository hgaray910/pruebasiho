VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDevolucionSub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mvardepartamento As Long
Private mvaridentity As Long
Private mvardepartamentoReq As Long

Public Property Let identity(ByVal vData As Long)
    mvaridentity = vData
End Property

Public Property Get identity() As Long
    identity = mvaridentity
End Property

Public Property Let departamento(ByVal vData As Long)
    mvardepartamento = vData
End Property

Public Property Get departamento() As Long
    departamento = mvardepartamento
End Property

Public Property Let departamentoReq(ByVal vData As Long)
    mvardepartamentoReq = vData
End Property

Public Property Get departamentoReq() As Long
    departamentoReq = mvardepartamentoReq
End Property



