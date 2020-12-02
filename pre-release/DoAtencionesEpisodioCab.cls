VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoAtencionesEpisodiosCab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla EpisodiosCabecera
'        Programado por: Cachay F
'        Fecha: Enero 2014
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_IdUsuarioAuditoria As Long
Dim ml_IdPaciente As Long
Dim ml_idEpisodio As Long
Dim mda_FechaApertura As Date
Dim mda_FechaCierre As Date
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let IdPaciente(iValue As Long)
   ml_IdPaciente = iValue
End Property
Property Get IdPaciente() As Long
   IdPaciente = ml_IdPaciente
End Property
Property Let idEpisodio(iValue As Long)
   ml_idEpisodio = iValue
End Property
Property Get idEpisodio() As Long
   idEpisodio = ml_idEpisodio
End Property
Property Let FechaApertura(daValue As Date)
   mda_FechaApertura = daValue
End Property
Property Get FechaApertura() As Date
   FechaApertura = mda_FechaApertura
End Property
Property Let FechaCierre(daValue As Date)
   mda_FechaCierre = daValue
End Property
Property Get FechaCierre() As Date
   FechaCierre = mda_FechaCierre
End Property

