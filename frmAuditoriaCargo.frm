VERSION 5.00
Begin VB.Form frmAuditoriaCargo 
   Caption         =   "Auditoría de cargos"
   ClientHeight    =   4710
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12090
   LinkTopic       =   "Form1"
   ScaleHeight     =   4710
   ScaleWidth      =   12090
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Height          =   4920
      Left            =   0
      TabIndex        =   0
      Top             =   -100
      Width           =   12135
      Begin VB.Frame Frame1 
         Caption         =   "Cargos disponibles"
         Height          =   800
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   5415
         Begin VB.OptionButton optTipoCargo 
            Caption         =   "Otros conceptos"
            Height          =   400
            Index           =   4
            Left            =   4200
            TabIndex        =   11
            Top             =   240
            Value           =   -1  'True
            Width           =   1100
         End
         Begin VB.OptionButton optTipoCargo 
            Caption         =   "Grupo de exámenes"
            Height          =   400
            Index           =   3
            Left            =   3075
            TabIndex        =   10
            Top             =   240
            Width           =   1100
         End
         Begin VB.OptionButton optTipoCargo 
            Caption         =   "Examen"
            Height          =   255
            Index           =   2
            Left            =   2040
            TabIndex        =   9
            Top             =   315
            Width           =   975
         End
         Begin VB.OptionButton optTipoCargo 
            Caption         =   "Estudio"
            Height          =   255
            Index           =   1
            Left            =   1080
            TabIndex        =   8
            Top             =   315
            Width           =   975
         End
         Begin VB.OptionButton optTipoCargo 
            Caption         =   "Artículo"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   7
            Top             =   315
            Width           =   975
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Cargos asignados"
         Height          =   800
         Left            =   6550
         TabIndex        =   5
         Top             =   240
         Width           =   5415
      End
      Begin VB.ListBox lstCargossDisponibles 
         Height          =   3570
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   1095
         Width           =   5415
      End
      Begin VB.ListBox lstCargosAsignados 
         Height          =   3570
         Left            =   6550
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   1095
         Width           =   5415
      End
      Begin VB.CommandButton cmdAgrega 
         Caption         =   ">"
         Height          =   495
         Left            =   5800
         TabIndex        =   1
         ToolTipText     =   "Agregar"
         Top             =   2000
         Width           =   495
      End
      Begin VB.CommandButton cmdElimina 
         Caption         =   "<"
         Height          =   495
         Left            =   5800
         TabIndex        =   2
         ToolTipText     =   "Eliminar"
         Top             =   2500
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmAuditoriaCargo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public vllngNumeroOpcion As Long

Private Sub cmdAgrega_Click()
On Error GoTo NotificaError
    Dim vlrsCargos As New ADODB.Recordset
    Dim vlrsCargosAud As New ADODB.Recordset
    Dim vlstrTipoCargo As String
    
    If lstCargossDisponibles.ListIndex <> -1 Then
        If optTipoCargo(0).Value = True Then
            vlstrTipoCargo = "AR"
        ElseIf optTipoCargo(1).Value = True Then
                vlstrTipoCargo = "ES"
            ElseIf optTipoCargo(2).Value = True Then
                    vlstrTipoCargo = "EX"
                ElseIf optTipoCargo(3).Value = True Then
                        vlstrTipoCargo = "GE"
                    Else
                        vlstrTipoCargo = "OC"
        End If
            
        EntornoSIHO.ConeccionSIHO.BeginTrans
        
        vgstrParametrosSP = vlstrTipoCargo & "|" & lstCargossDisponibles.ItemData(lstCargossDisponibles.ListIndex)
        frsEjecuta_SP vgstrParametrosSP, "SP_PVINSPVPRECIOSAUDITORIA"
        
        EntornoSIHO.ConeccionSIHO.CommitTrans
        
        Set vlrsCargos = frsEjecuta_SP(vlstrTipoCargo, "SP_PVSELAUDCARGOSDISPONIBLES")
        pLlenarListRs lstCargossDisponibles, vlrsCargos, 0, 1
        lstCargossDisponibles.ListIndex = IIf(lstCargossDisponibles.ListCount > 0, 0, -1)
        
        Set vlrsCargosAud = frsEjecuta_SP(vlstrTipoCargo, "SP_PVSELAUDCARGOS")
        pLlenarListRs lstCargosAsignados, vlrsCargosAud, 0, 1
        lstCargosAsignados.ListIndex = IIf(lstCargosAsignados.ListCount > 0, 0, -1)
    Else
        MsgBox "No hay selección", vbInformation, "Mensaje"
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdAgrega_Click"))

End Sub

Private Sub cmdElimina_Click()
On Error GoTo NotificaError
    Dim vlrsCargos As New ADODB.Recordset
    Dim vlrsCargosAud As New ADODB.Recordset
    Dim vlstrTipoCargo As String
    Dim vlstrSentencia As String
    
    If lstCargosAsignados.ListIndex <> -1 Then
        If optTipoCargo(0).Value = True Then
            vlstrTipoCargo = "AR"
        ElseIf optTipoCargo(1).Value = True Then
                vlstrTipoCargo = "ES"
            ElseIf optTipoCargo(2).Value = True Then
                    vlstrTipoCargo = "EX"
                ElseIf optTipoCargo(3).Value = True Then
                        vlstrTipoCargo = "GE"
                    Else
                        vlstrTipoCargo = "OC"
        End If
            
        EntornoSIHO.ConeccionSIHO.BeginTrans
        
        vlstrSentencia = "DELETE FROM PVPRECIOSAUDITORIA WHERE chrtipocargo = '" & vlstrTipoCargo & "'  AND chrcvecargo = '" & lstCargosAsignados.ItemData(lstCargosAsignados.ListIndex) & "' "
        pEjecutaSentencia vlstrSentencia
        
        'vgstrParametrosSP = vlstrTipoCargo & "|" & lstCargosAsignados.ItemData(lstCargosAsignados.ListIndex)
        'frsEjecuta_SP vgstrParametrosSP, "SP_PVDELPVPRECIOSAUDITORIA"
        
        EntornoSIHO.ConeccionSIHO.CommitTrans
        
        Set vlrsCargos = frsEjecuta_SP(vlstrTipoCargo, "SP_PVSELAUDCARGOSDISPONIBLES")
        pLlenarListRs lstCargossDisponibles, vlrsCargos, 0, 1
        lstCargossDisponibles.ListIndex = IIf(lstCargossDisponibles.ListCount > 0, 0, -1)
        
        Set vlrsCargosAud = frsEjecuta_SP(vlstrTipoCargo, "SP_PVSELAUDCARGOS")
        pLlenarListRs lstCargosAsignados, vlrsCargosAud, 0, 1
        lstCargosAsignados.ListIndex = IIf(lstCargosAsignados.ListCount > 0, 0, -1)
    Else
        MsgBox "No hay selección", vbInformation, "Mensaje"
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdElimina_Click"))

End Sub

Private Sub Form_Load()
On Error GoTo NotificaError
    Dim vlrsCargos As New ADODB.Recordset
    Dim vlrsCargosAud As New ADODB.Recordset
    Dim vlstrTipoCargo As String
    
    Me.Icon = frmMenuPrincipal.Icon
    
    optTipoCargo(0).Value = True
    vlstrTipoCargo = "AR"
    
    Set vlrsCargos = frsEjecuta_SP(vlstrTipoCargo, "SP_PVSELAUDCARGOSDISPONIBLES")
    pLlenarListRs lstCargossDisponibles, vlrsCargos, 0, 1
    lstCargossDisponibles.ListIndex = IIf(lstCargossDisponibles.ListCount > 0, 0, -1)
    
    Set vlrsCargosAud = frsEjecuta_SP(vlstrTipoCargo, "SP_PVSELAUDCARGOS")
    pLlenarListRs lstCargosAsignados, vlrsCargosAud, 0, 1
    lstCargosAsignados.ListIndex = IIf(lstCargosAsignados.ListCount > 0, 0, -1)
    
    If lstCargossDisponibles.ListCount = 0 Then
        cmdAgrega.Enabled = False
        MsgBox SIHOMsg(1568), vbOKOnly + vbExclamation, "Mensaje"
    Else
        cmdAgrega.Enabled = True
    End If
        
    If lstCargosAsignados.ListCount = 0 Then
        cmdElimina.Enabled = False
    Else
        cmdElimina.Enabled = True
    End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_Load"))

End Sub

Private Sub lstCargossDisponibles_DblClick()
    cmdAgrega_Click
End Sub

Private Sub optTipoCargo_Click(Index As Integer)
On Error GoTo NotificaError
Dim vlrsCargos As New ADODB.Recordset
    Dim vlrsCargosAud As New ADODB.Recordset
    Dim vlstrTipoCargo As String
    Dim vlstrSentencia As String
    
        If optTipoCargo(0).Value = True Then
            vlstrTipoCargo = "AR"
        ElseIf optTipoCargo(1).Value = True Then
                vlstrTipoCargo = "ES"
            ElseIf optTipoCargo(2).Value = True Then
                    vlstrTipoCargo = "EX"
                ElseIf optTipoCargo(3).Value = True Then
                        vlstrTipoCargo = "GE"
                    Else
                        vlstrTipoCargo = "OC"
        End If
    
        Set vlrsCargos = frsEjecuta_SP(vlstrTipoCargo, "SP_PVSELAUDCARGOSDISPONIBLES")
        pLlenarListRs lstCargossDisponibles, vlrsCargos, 0, 1
        lstCargossDisponibles.ListIndex = IIf(lstCargossDisponibles.ListCount > 0, 0, -1)
        
        Set vlrsCargosAud = frsEjecuta_SP(vlstrTipoCargo, "SP_PVSELAUDCARGOS")
        pLlenarListRs lstCargosAsignados, vlrsCargosAud, 0, 1
        lstCargosAsignados.ListIndex = IIf(lstCargosAsignados.ListCount > 0, 0, -1)
        
        If lstCargossDisponibles.ListCount = 0 Then
            cmdAgrega.Enabled = False
            MsgBox SIHOMsg(1568), vbOKOnly + vbExclamation, "Mensaje"
        Else
           cmdAgrega.Enabled = True
        End If
        
        If lstCargosAsignados.ListCount = 0 Then
            cmdElimina.Enabled = False
        Else
            cmdElimina.Enabled = True
        End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":optTipoCargo_Click"))
End Sub
