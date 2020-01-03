VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoNotaCreditoDebito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: NotaCreditoDebito
'        Autor: Daniel Enrique Barrantes
'        Fecha: 13/08/2015 10:54:51
'        Auspiciado por: Usaid/Politicas en Salud
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------

Dim ml_IdUsuarioAuditoria As Long
Dim ml_IdNota As Long
Dim ml_IdComprobantePago As Long
Dim ml_IdTipoNota As Long
Dim ms_NroSerie As String
Dim ms_NroDocumento As String
Dim ms_RazonSocial As String
Dim ms_RUC As String
Dim mc_SubTotal As Currency
Dim mc_IGV As Currency
Dim mc_Total As Currency
Dim ml_IdUsuarioAutoriza As Long
Dim mda_FechaAprueba As Date
Dim mc_TipoCambio As Currency
Dim ms_Observaciones As String
Dim ml_IdEstadoNota As Long
Dim mda_FechaPagado As Date
Dim ml_IdGestionCaja As Long
Dim ml_IdPaciente As Long
Dim ml_IdCajero As Long
Dim ml_idTurno As Long
Dim ml_idCaja As Long
Dim ml_idFarmacia As Long
Dim ml_idMotivo As Long
Dim ms_Direccion As String
Dim mb_TipoAnulacion As Boolean

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let IdNota(iValue As Long)
   ml_IdNota = iValue
End Property
Property Get IdNota() As Long
   IdNota = ml_IdNota
End Property
Property Let IdComprobantePago(iValue As Long)
   ml_IdComprobantePago = iValue
End Property
Property Get IdComprobantePago() As Long
   IdComprobantePago = ml_IdComprobantePago
End Property
Property Let IdTipoNota(iValue As Long)
   ml_IdTipoNota = iValue
End Property
Property Get IdTipoNota() As Long
   IdTipoNota = ml_IdTipoNota
End Property
Property Let NroSerie(sValue As String)
   ms_NroSerie = sValue
End Property
Property Get NroSerie() As String
   NroSerie = ms_NroSerie
End Property
Property Let NroDocumento(sValue As String)
   ms_NroDocumento = sValue
End Property
Property Get NroDocumento() As String
   NroDocumento = ms_NroDocumento
End Property
Property Let RazonSocial(sValue As String)
   ms_RazonSocial = sValue
End Property
Property Get RazonSocial() As String
   RazonSocial = ms_RazonSocial
End Property
Property Let RUC(sValue As String)
   ms_RUC = sValue
End Property
Property Get RUC() As String
   RUC = ms_RUC
End Property
Property Let SubTotal(cValue As Currency)
   mc_SubTotal = cValue
End Property
Property Get SubTotal() As Currency
   SubTotal = mc_SubTotal
End Property
Property Let IGV(cValue As Currency)
   mc_IGV = cValue
End Property
Property Get IGV() As Currency
   IGV = mc_IGV
End Property
Property Let Total(cValue As Currency)
   mc_Total = cValue
End Property
Property Get Total() As Currency
   Total = mc_Total
End Property
Property Let IdUsuarioAutoriza(iValue As Long)
   ml_IdUsuarioAutoriza = iValue
End Property
Property Get IdUsuarioAutoriza() As Long
   IdUsuarioAutoriza = ml_IdUsuarioAutoriza
End Property
Property Let FechaAprueba(daValue As Date)
   mda_FechaAprueba = daValue
End Property
Property Get FechaAprueba() As Date
   FechaAprueba = mda_FechaAprueba
End Property
Property Let TipoCambio(cValue As Currency)
   mc_TipoCambio = cValue
End Property
Property Get TipoCambio() As Currency
   TipoCambio = mc_TipoCambio
End Property
Property Let Observaciones(sValue As String)
   ms_Observaciones = sValue
End Property
Property Get Observaciones() As String
   Observaciones = ms_Observaciones
End Property
Property Let IdEstadoNota(iValue As Long)
   ml_IdEstadoNota = iValue
End Property
Property Get IdEstadoNota() As Long
   IdEstadoNota = ml_IdEstadoNota
End Property
Property Let FechaPagado(daValue As Date)
   mda_FechaPagado = daValue
End Property
Property Get FechaPagado() As Date
   FechaPagado = mda_FechaPagado
End Property
Property Let IdGestionCaja(iValue As Long)
   ml_IdGestionCaja = iValue
End Property
Property Get IdGestionCaja() As Long
   IdGestionCaja = ml_IdGestionCaja
End Property
Property Let IdPaciente(iValue As Long)
   ml_IdPaciente = iValue
End Property
Property Get IdPaciente() As Long
   IdPaciente = ml_IdPaciente
End Property
Property Let IdCajero(iValue As Long)
   ml_IdCajero = iValue
End Property
Property Get IdCajero() As Long
   IdCajero = ml_IdCajero
End Property
Property Let idTurno(iValue As Long)
   ml_idTurno = iValue
End Property
Property Get idTurno() As Long
   idTurno = ml_idTurno
End Property
Property Let idCaja(iValue As Long)
   ml_idCaja = iValue
End Property
Property Get idCaja() As Long
   idCaja = ml_idCaja
End Property
Property Let idFarmacia(iValue As Long)
   ml_idFarmacia = iValue
End Property
Property Get idFarmacia() As Long
   idFarmacia = ml_idFarmacia
End Property
Property Let idMotivo(iValue As Long)
   ml_idMotivo = iValue
End Property
Property Get idMotivo() As Long
   idMotivo = ml_idMotivo
End Property
Property Let Direccion(sValue As String)
   ms_Direccion = sValue
End Property
Property Get Direccion() As String
   Direccion = ms_Direccion
End Property
Property Let TipoAnulacion(bValue As Boolean)
   mb_TipoAnulacion = bValue
End Property
Property Get TipoAnulacion() As Boolean
   TipoAnulacion = mb_TipoAnulacion
End Property
