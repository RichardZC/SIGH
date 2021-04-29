VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoFacturacionServicioFin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla FacturacionServicioFinanciamientos
'        Programado por: Barrantes D
'        Fecha: Abril 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_IdUsuarioAuditoria As Long
Dim ml_IdOrden As Long
Dim ml_idProducto As Long
Dim ml_IdTipoFinanciamiento As Long
Dim ml_CantidadFinanciada As Long
Dim mc_PrecioFinanciado As Currency
Dim mc_TotalFinanciado As Currency
Dim mda_FechaAutoriza As Date
Dim ml_IdUsuarioAutoriza As Long
Dim ml_idFuenteFinanciamiento As Long
Dim ml_IdEstadoFacturacion As Long

'SCCQ 26-04-2021 Cambio 64 Inicio (Reversion Cambio 46)
'RHA 17/11/2020 Cambio 46 Inicio
'Dim ml_CodPres As String
'Property Get CodPres() As String
'   CodPres = ml_CodPres
'End Property
'Property Let CodPres(lValue As String)
'   ml_CodPres = lValue
'End Property
'RHA 17/11/2020 Cambio 46 Fin
'SCCQ 26-04-2021 Cambio 64 Fin (Reversion Cambio 46)


Property Let IdEstadoFacturacion(iValue As Long)
   ml_IdEstadoFacturacion = iValue
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


Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property

Property Let IdOrden(iValue As Long)
   ml_IdOrden = iValue
End Property
Property Get IdOrden() As Long
   IdOrden = ml_IdOrden
End Property
Property Let idProducto(iValue As Long)
   ml_idProducto = iValue
End Property
Property Get idProducto() As Long
   idProducto = ml_idProducto
End Property
Property Let IdTipoFinanciamiento(iValue As Long)
   ml_IdTipoFinanciamiento = iValue
End Property
Property Get IdTipoFinanciamiento() As Long
   IdTipoFinanciamiento = ml_IdTipoFinanciamiento
End Property
Property Let CantidadFinanciada(iValue As Long)
   ml_CantidadFinanciada = iValue
End Property
Property Get CantidadFinanciada() As Long
   CantidadFinanciada = ml_CantidadFinanciada
End Property
Property Let PrecioFinanciado(cValue As Currency)
   mc_PrecioFinanciado = cValue
End Property
Property Get PrecioFinanciado() As Currency
   PrecioFinanciado = mc_PrecioFinanciado
End Property
Property Let TotalFinanciado(cValue As Currency)
   mc_TotalFinanciado = cValue
End Property
Property Get TotalFinanciado() As Currency
   TotalFinanciado = mc_TotalFinanciado
End Property
Property Let FechaAutoriza(daValue As Date)
   mda_FechaAutoriza = daValue
End Property
Property Get FechaAutoriza() As Date
   FechaAutoriza = mda_FechaAutoriza
End Property
Property Let IdUsuarioAutoriza(iValue As Long)
   ml_IdUsuarioAutoriza = iValue
End Property
Property Get IdUsuarioAutoriza() As Long
   IdUsuarioAutoriza = ml_IdUsuarioAutoriza
End Property


