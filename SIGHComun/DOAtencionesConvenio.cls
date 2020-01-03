VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOAtencionesConvenio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla AtencionesConvenio
'        Programado por: Castro W
'        Fecha: Enero 2007
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_Auditoria As Long
Dim ms_NroOficio As String
Dim ml_IdPaciente As Long
Dim ml_IdAtencionesConvenio As Long
Dim mc_ImporteSesion As Currency
Dim md_FechaSesion As Date
Dim ms_NombrePaciente As String
Dim ms_NombreProducto  As String
Dim ml_IdProducto As Long
Dim ms_NroHistoria  As String
Dim ms_CodProducto  As String
Property Let CodProducto(sValue As String)
   ms_CodProducto = sValue
End Property
Property Get CodProducto() As String
   CodProducto = ms_CodProducto
End Property
Property Let NroHistoria(sValue As String)
   ms_NroHistoria = sValue
End Property
Property Get NroHistoria() As String
   NroHistoria = ms_NroHistoria
End Property

Property Let IdProducto(lValue As Long)
   ml_IdProducto = lValue
End Property
Property Get IdProducto() As Long
   IdProducto = ml_IdProducto
End Property

Property Let NombreProducto(sValue As String)
   ms_NombreProducto = sValue
End Property
Property Get NombreProducto() As String
   NombreProducto = ms_NombreProducto
End Property
Property Let NombrePaciente(sValue As String)
   ms_NombrePaciente = sValue
End Property
Property Get NombrePaciente() As String
   NombrePaciente = ms_NombrePaciente
End Property
Property Let FechaSesion(cValue As Date)
   md_FechaSesion = cValue
End Property
Property Get FechaSesion() As Date
   FechaSesion = md_FechaSesion
End Property

Property Let ImporteSesion(cValue As Currency)
   mc_ImporteSesion = cValue
End Property
Property Get ImporteSesion() As Currency
   ImporteSesion = mc_ImporteSesion
End Property

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let NroOficio(sValue As String)
   ms_NroOficio = sValue
End Property
Property Get NroOficio() As String
   NroOficio = ms_NroOficio
End Property
Property Let IdPaciente(lValue As Long)
   ml_IdPaciente = lValue
End Property
Property Get IdPaciente() As Long
   IdPaciente = ml_IdPaciente
End Property
Property Let IdAtencionesConvenio(lValue As Long)
   ml_IdAtencionesConvenio = lValue
End Property
Property Get IdAtencionesConvenio() As Long
   IdAtencionesConvenio = ml_IdAtencionesConvenio
End Property


