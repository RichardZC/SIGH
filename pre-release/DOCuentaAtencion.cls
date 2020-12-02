VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOCuentaAtencion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla FacturacionCuentasAtencion
'        Programado por: Barrantes D
'        Fecha: Agosto 2010
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_Auditoria As Long
Dim mc_TotalPorPagar As Currency
Dim ml_idEstado As Long
Dim mc_TotalPagado As Currency
Dim mc_TotalAsegurado As Currency
Dim mc_TotalExonerado As Currency
Dim ms_HoraCierre As String
Dim mda_FechaCierre As Date
Dim ms_HoraApertura As String
Dim mda_FechaApertura As Date
Dim ml_IdPaciente As Long
Dim ml_idCuentaAtencion As Long
Dim mda_FechaCreacion As Date
Property Let FechaCreacion(daValue As Date)
   mda_FechaCreacion = daValue
End Property
Property Get FechaCreacion() As Date
   FechaCreacion = mda_FechaCreacion
End Property
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let TotalPorPagar(cValue As Currency)
   mc_TotalPorPagar = cValue
End Property
Property Get TotalPorPagar() As Currency
   TotalPorPagar = mc_TotalPorPagar
End Property
Property Let idEstado(lValue As Long)
   ml_idEstado = lValue
End Property
Property Get idEstado() As Long
   idEstado = ml_idEstado
End Property
Property Let TotalPagado(cValue As Currency)
   mc_TotalPagado = cValue
End Property
Property Get TotalPagado() As Currency
   TotalPagado = mc_TotalPagado
End Property
Property Let TotalAsegurado(cValue As Currency)
   mc_TotalAsegurado = cValue
End Property
Property Get TotalAsegurado() As Currency
   TotalAsegurado = mc_TotalAsegurado
End Property
Property Let TotalExonerado(cValue As Currency)
   mc_TotalExonerado = cValue
End Property
Property Get TotalExonerado() As Currency
   TotalExonerado = mc_TotalExonerado
End Property
Property Let HoraCierre(sValue As String)
   ms_HoraCierre = sValue
End Property
Property Get HoraCierre() As String
   HoraCierre = ms_HoraCierre
End Property
Property Let FechaCierre(daValue As Date)
   mda_FechaCierre = daValue
End Property
Property Get FechaCierre() As Date
   FechaCierre = mda_FechaCierre
End Property
Property Let HoraApertura(sValue As String)
   ms_HoraApertura = sValue
End Property
Property Get HoraApertura() As String
   HoraApertura = ms_HoraApertura
End Property
Property Let FechaApertura(daValue As Date)
   mda_FechaApertura = daValue
End Property
Property Get FechaApertura() As Date
   FechaApertura = mda_FechaApertura
End Property
Property Let IdPaciente(lValue As Long)
   ml_IdPaciente = lValue
End Property
Property Get IdPaciente() As Long
   IdPaciente = ml_IdPaciente
End Property
Property Let idCuentaAtencion(lValue As Long)
   ml_idCuentaAtencion = lValue
End Property
Property Get idCuentaAtencion() As Long
   idCuentaAtencion = ml_idCuentaAtencion
End Property


