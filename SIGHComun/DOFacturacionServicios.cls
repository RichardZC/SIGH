VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOFacturacionServicios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla FacturacionServicios
'        Programado por: Barrantes D
'        Fecha: Abril 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_Auditoria As Long
Dim mda_FechaCreacion As Date
Dim ml_IdProducto As Long
Dim mc_PrecioUnitario As Currency
Dim ml_IdEstadoFacturacion As Long
Dim ml_idFuenteFinanciamiento As Long
Dim ml_IdTipoFinanciamiento As Long
Dim ml_IdPartidaPresupuestal As Long
Dim ml_IdCentroCosto As Long
Dim ml_IdServicioInternamiento As Long
Dim ml_idComprobantePago As Long
Dim mc_TotalPorPagar As Currency
Dim ml_IdAtencion As Long
Dim ml_IdFacturacionServicio As Long
Dim ml_idUsuarioModifica As Long
Dim ml_idPuntoCarga As Long
Dim mda_FechaModificacion As Date
Dim ml_idOrden As Long
Dim ml_IdUsuarioAutorizaSeguro As Long
Dim ml_IdUsuarioAutorizaPendiente As Long
Dim ml_IdUsuarioAutorizaDevolucion As Long
Dim ml_IdCajero As Long
Dim mda_FechaAutorizaSeguro As Date
Dim mda_FechaAutorizaPendiente As Date
Dim mda_FechaAutorizaDevolucion As Date
Dim mda_FechaCajero As Date
Dim ml_IdComprobantePagoDevolucion As Long
Dim md_Cantidad As Double
Dim ml_IdUsuarioCrea As Long

'***************daniel barrantes**************
'***************Campos creado para SIS/Soat/Exoneraciones
'***************
Dim mc_ImporteSIS As Currency
Dim mc_ImporteEXO As Currency
Dim ml_idUsuarioAutorizaEXO As Long
Dim mda_FechaAutorizaEXO As Date
Dim ml_idProductoLocal As Long
Dim mc_ImporteSOAT As Currency
Dim mc_TotalPagar As Currency
Dim mc_PrecioSIS As Currency
Dim ml_CantidadSIS As Long
Dim mc_PrecioSOAT As Currency
Dim ml_CantidadSOAT As Long
Dim ml_CantidadPagar As Long
Dim ml_cantidadDev As Long
Dim mda_FechaAutorizaConv As Long
Dim ml_cantidadConv As Long
Dim ml_precConv As Long
Dim ml_ImporteConv As Long


Property Let ImporteConv(lValue As Long)
   ml_ImporteConv = lValue
End Property
Property Get ImporteConv() As Long
   ImporteConv = ml_ImporteConv
End Property
Property Let precConv(lValue As Long)
   ml_precConv = lValue
End Property
Property Get precConv() As Long
   precConv = ml_precConv
End Property
Property Let cantidadConv(lValue As Long)
   ml_cantidadConv = lValue
End Property
Property Get cantidadConv() As Long
   cantidadConv = ml_cantidadConv
End Property
Property Let FechaAutorizaConv(lValue As Date)
   mda_FechaAutorizaConv = lValue
End Property
Property Get FechaAutorizaConv() As Date
   FechaAutorizaConv = mda_FechaAutorizaConv
End Property
Property Let cantidadDev(lValue As Long)
   ml_cantidadDev = lValue
End Property
Property Get cantidadDev() As Long
   cantidadDev = ml_cantidadDev
End Property


Property Let CantidadSIS(lValue As Long)
   ml_CantidadSIS = lValue
End Property
Property Get CantidadSIS() As Long
   CantidadSIS = ml_CantidadSIS
End Property
Property Let CantidadSOAT(lValue As Long)
   ml_CantidadSOAT = lValue
End Property
Property Get CantidadSOAT() As Long
   CantidadSOAT = ml_CantidadSOAT
End Property
Property Let CantidadPagar(lValue As Long)
   ml_CantidadPagar = lValue
End Property
Property Get CantidadPagar() As Long
   CantidadPagar = ml_CantidadPagar
End Property
Property Let PrecioSIS(cValue As Currency)
   mc_PrecioSIS = cValue
End Property
Property Get PrecioSIS() As Currency
   PrecioSIS = mc_PrecioSIS
End Property
Property Let PrecioSOAT(cValue As Currency)
   mc_PrecioSOAT = cValue
End Property
Property Get PrecioSOAT() As Currency
   PrecioSOAT = mc_PrecioSOAT
End Property



Property Let TotalPagar(cValue As Currency)
   mc_TotalPagar = cValue
End Property
Property Get TotalPagar() As Currency
   TotalPagar = mc_TotalPagar
End Property


Property Let ImporteSOAT(cValue As Currency)
   mc_ImporteSOAT = cValue
End Property
Property Get ImporteSOAT() As Currency
   ImporteSOAT = mc_ImporteSOAT
End Property
Property Let idProductoLocal(lValue As Long)
   ml_idProductoLocal = lValue
End Property
Property Get idProductoLocal() As Long
   idProductoLocal = ml_idProductoLocal
End Property
Property Let FechaAutorizaEXO(daValue As Date)
   mda_FechaAutorizaEXO = daValue
End Property
Property Get FechaAutorizaEXO() As Date
   FechaAutorizaEXO = mda_FechaAutorizaEXO
End Property
Property Let idUsuarioAutorizaEXO(lValue As Long)
   ml_idUsuarioAutorizaEXO = lValue
End Property
Property Get idUsuarioAutorizaEXO() As Long
   idUsuarioAutorizaEXO = ml_idUsuarioAutorizaEXO
End Property
Property Let ImporteEXO(cValue As Currency)
   mc_ImporteEXO = cValue
End Property
Property Get ImporteEXO() As Currency
   ImporteEXO = mc_ImporteEXO
End Property
Property Let ImporteSIS(cValue As Currency)
   mc_ImporteSIS = cValue
End Property
Property Get ImporteSIS() As Currency
   ImporteSIS = mc_ImporteSIS
End Property


Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let FechaCreacion(daValue As Date)
   mda_FechaCreacion = daValue
End Property
Property Get FechaCreacion() As Date
   FechaCreacion = mda_FechaCreacion
End Property
Property Let IdProducto(lValue As Long)
   ml_IdProducto = lValue
End Property
Property Get IdProducto() As Long
   IdProducto = ml_IdProducto
End Property
Property Let PrecioUnitario(cValue As Currency)
   mc_PrecioUnitario = cValue
End Property
Property Get PrecioUnitario() As Currency
   PrecioUnitario = mc_PrecioUnitario
End Property
Property Let IdEstadoFacturacion(lValue As Long)
   ml_IdEstadoFacturacion = lValue
End Property
Property Get IdEstadoFacturacion() As Long
   IdEstadoFacturacion = ml_IdEstadoFacturacion
End Property
Property Let idFuenteFinanciamiento(lValue As Long)
   ml_idFuenteFinanciamiento = lValue
End Property
Property Get idFuenteFinanciamiento() As Long
   idFuenteFinanciamiento = ml_idFuenteFinanciamiento
End Property
Property Let IdTipoFinanciamiento(lValue As Long)
   ml_IdTipoFinanciamiento = lValue
End Property
Property Get IdTipoFinanciamiento() As Long
   IdTipoFinanciamiento = ml_IdTipoFinanciamiento
End Property
Property Let IdPartidaPresupuestal(lValue As Long)
   ml_IdPartidaPresupuestal = lValue
End Property
Property Get IdPartidaPresupuestal() As Long
   IdPartidaPresupuestal = ml_IdPartidaPresupuestal
End Property
Property Let IdCentroCosto(lValue As Long)
   ml_IdCentroCosto = lValue
End Property
Property Get IdCentroCosto() As Long
   IdCentroCosto = ml_IdCentroCosto
End Property
Property Let IdServicioInternamiento(lValue As Long)
   ml_IdServicioInternamiento = lValue
End Property
Property Get IdServicioInternamiento() As Long
   IdServicioInternamiento = ml_IdServicioInternamiento
End Property
Property Let idComprobantePago(lValue As Long)
   ml_idComprobantePago = lValue
End Property
Property Get idComprobantePago() As Long
   idComprobantePago = ml_idComprobantePago
End Property
Property Let TotalPorPagar(cValue As Currency)
   mc_TotalPorPagar = cValue
End Property
Property Get TotalPorPagar() As Currency
   TotalPorPagar = mc_TotalPorPagar
End Property
Property Let IdAtencion(lValue As Long)
   ml_IdAtencion = lValue
End Property
Property Get IdAtencion() As Long
   IdAtencion = ml_IdAtencion
End Property
Property Let IdFacturacionServicio(lValue As Long)
   ml_IdFacturacionServicio = lValue
End Property
Property Get IdFacturacionServicio() As Long
   IdFacturacionServicio = ml_IdFacturacionServicio
End Property
Property Let idUsuarioModifica(lValue As Long)
   ml_idUsuarioModifica = lValue
End Property
Property Get idUsuarioModifica() As Long
   idUsuarioModifica = ml_idUsuarioModifica
End Property
Property Let idPuntoCarga(lValue As Long)
   ml_idPuntoCarga = lValue
End Property
Property Get idPuntoCarga() As Long
   idPuntoCarga = ml_idPuntoCarga
End Property
Property Let FechaModificacion(daValue As Date)
   mda_FechaModificacion = daValue
End Property
Property Get FechaModificacion() As Date
   FechaModificacion = mda_FechaModificacion
End Property
Property Let idOrden(lValue As Long)
   ml_idOrden = lValue
End Property
Property Get idOrden() As Long
   idOrden = ml_idOrden
End Property
Property Let IdUsuarioAutorizaSeguro(lValue As Long)
   ml_IdUsuarioAutorizaSeguro = lValue
End Property
Property Get IdUsuarioAutorizaSeguro() As Long
   IdUsuarioAutorizaSeguro = ml_IdUsuarioAutorizaSeguro
End Property
Property Let IdUsuarioAutorizaPendiente(lValue As Long)
   ml_IdUsuarioAutorizaPendiente = lValue
End Property
Property Get IdUsuarioAutorizaPendiente() As Long
   IdUsuarioAutorizaPendiente = ml_IdUsuarioAutorizaPendiente
End Property
Property Let IdUsuarioAutorizaDevolucion(lValue As Long)
   ml_IdUsuarioAutorizaDevolucion = lValue
End Property
Property Get IdUsuarioAutorizaDevolucion() As Long
   IdUsuarioAutorizaDevolucion = ml_IdUsuarioAutorizaDevolucion
End Property
Property Let IdCajero(lValue As Long)
   ml_IdCajero = lValue
End Property
Property Get IdCajero() As Long
   IdCajero = ml_IdCajero
End Property
Property Let FechaAutorizaSeguro(daValue As Date)
   mda_FechaAutorizaSeguro = daValue
End Property
Property Get FechaAutorizaSeguro() As Date
   FechaAutorizaSeguro = mda_FechaAutorizaSeguro
End Property
Property Let FechaAutorizaPendiente(daValue As Date)
   mda_FechaAutorizaPendiente = daValue
End Property
Property Get FechaAutorizaPendiente() As Date
   FechaAutorizaPendiente = mda_FechaAutorizaPendiente
End Property
Property Let FechaAutorizaDevolucion(daValue As Date)
   mda_FechaAutorizaDevolucion = daValue
End Property
Property Get FechaAutorizaDevolucion() As Date
   FechaAutorizaDevolucion = mda_FechaAutorizaDevolucion
End Property
Property Let FechaCajero(daValue As Date)
   mda_FechaCajero = daValue
End Property
Property Get FechaCajero() As Date
   FechaCajero = mda_FechaCajero
End Property
Property Let IdComprobantePagoDevolucion(lValue As Long)
   ml_IdComprobantePagoDevolucion = lValue
End Property
Property Get IdComprobantePagoDevolucion() As Long
   IdComprobantePagoDevolucion = ml_IdComprobantePagoDevolucion
End Property
Property Let Cantidad(dValue As Double)
   md_Cantidad = dValue
End Property
Property Get Cantidad() As Double
   Cantidad = md_Cantidad
End Property
Property Let IdUsuarioCrea(lValue As Long)
   ml_IdUsuarioCrea = lValue
End Property
Property Get IdUsuarioCrea() As Long
   IdUsuarioCrea = ml_IdUsuarioCrea
End Property


