VERSION 5.00
Begin VB.Form FrmGestionCajaDesdeServicios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   9420
   ClientLeft      =   4260
   ClientTop       =   1230
   ClientWidth     =   13185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9420
   ScaleWidth      =   13185
   StartUpPosition =   2  'CenterScreen
   Begin SISGalenPlus.ucGestionCaja ucGestionCaja1 
      Height          =   9375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   16536
   End
End
Attribute VB_Name = "FrmGestionCajaDesdeServicios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mb_abrioCaja As Boolean
Dim mo_AdminCaja As New SIGHNegocios.ReglasCaja
Dim ml_IdUsuarioAuditoria As Long
Dim lc_NombrePc As String
Dim lNroOrdenPago As Long
Dim mb_CerrarAlGuardar As Boolean

Property Let lcNombrePc(lValue As String)
   lc_NombrePc = lValue
End Property

Property Let lIdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property

Property Let lNumeroOrden(lValue As Long)
   lNroOrdenPago = lValue
End Property

Property Let CerrarAlGuardar(bValue As Boolean)
   mb_CerrarAlGuardar = bValue
End Property


Private Sub Form_Activate()
'    If (mb_abrioCaja = False) Then
        Dim oDOCajaGestion As New DOCajaGestion
        
        Set oDOCajaGestion = Principal.oDOCajaGestion
        If oDOCajaGestion.IdTurno = 0 Then oDOCajaGestion.IdTurno = 1   'debb-16/03/2016
        mb_abrioCaja = Me.ucGestionCaja1.RealizarAperturaDeCaja(ml_IdUsuarioAuditoria, _
                                oDOCajaGestion.IdCaja, oDOCajaGestion.IdTurno, Principal.bCajeroEmiteSoloServicios)
'    End If
    ucGestionCaja1.NombreCajero = Principal.status.Panels(2).Text
    ucGestionCaja1.Visible = True
    ucGestionCaja1.ActivarTabGestionCaja 1
    ucGestionCaja1.ActivarOrdenExistenteFS
    
    ucGestionCaja1.AsignarNroOrden CStr(lNroOrdenPago)
    ucGestionCaja1.BuscarOrdenExistente
'    oDOCajaGestion.
End Sub

Private Function CompartamientoCaja()
    
End Function

Private Sub Form_Load()
'    ucGestionCaja1.NombreCajero = Principal.status.Panels(2).Text
    ucGestionCaja1.Visible = True
    ucGestionCaja1.idUsuario = ml_IdUsuarioAuditoria
    ucGestionCaja1.NombreCajero = Principal.status.Panels(2).Text
    ucGestionCaja1.lnIdTablaLISTBARITEMS = 702
    ucGestionCaja1.lcNombrePc = lc_NombrePc
    'toolbar.Toolbars("Gestión de Caja").Visible = True
    ucGestionCaja1.inicializar
End Sub

Private Sub ucGestionCaja1_GuardoComprobante(bGuardo As Boolean)
    If bGuardo = True Then
        If mb_CerrarAlGuardar = True Then
            Unload Me
        End If
    End If
End Sub
