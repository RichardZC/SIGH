VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOAtencionNacimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla AtencionesNacimiento
'        Programado por: Barrantes D
'        Fecha: Enero 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_Auditoria As Long
Dim ml_IdAtencion As Long
Dim ml_IdCondicionRN As Long
Dim ml_IdTipoSexo As Long
Dim md_Peso As Double
Dim md_Talla As Double
Dim ml_EdadSemanas As Long
Dim mda_FechaNacimiento As Date
Dim ml_IdNacimiento As Long
Dim ml_Apgar_1 As Long
Dim ml_Apgar_5 As Long
Dim mda_ClamplajeFecha As Date
Dim ml_NroOrdenHijoEnParto As Long
Dim ml_NroOrdenHijo As Long
Dim ml_idPacienteNacido As Long
Dim ml_IdDocIdentidad As Long
Dim ml_docIdentidad As String

Property Let DocIdentidad(lValue As String)
   ml_docIdentidad = lValue
End Property
Property Get DocIdentidad() As String
   DocIdentidad = ml_docIdentidad
End Property
Property Let IdDocIdentidad(lValue As Long)
   ml_IdDocIdentidad = lValue
End Property
Property Get IdDocIdentidad() As Long
   IdDocIdentidad = ml_IdDocIdentidad
End Property

Property Let idPacienteNacido(lValue As Long)
   ml_idPacienteNacido = lValue
End Property
Property Get idPacienteNacido() As Long
   idPacienteNacido = ml_idPacienteNacido
End Property
Property Let NroOrdenHijo(lValue As Long)
   ml_NroOrdenHijo = lValue
End Property
Property Get NroOrdenHijo() As Long
   NroOrdenHijo = ml_NroOrdenHijo
End Property

Property Let NroOrdenHijoEnParto(lValue As Long)
   ml_NroOrdenHijoEnParto = lValue
End Property
Property Get NroOrdenHijoEnParto() As Long
   NroOrdenHijoEnParto = ml_NroOrdenHijoEnParto
End Property
Property Let ClamplajeFecha(daValue As Date)
   mda_ClamplajeFecha = daValue
End Property
Property Get ClamplajeFecha() As Date
   ClamplajeFecha = mda_ClamplajeFecha
End Property



Property Let Apgar_5(lValue As Long)
   ml_Apgar_5 = lValue
End Property
Property Get Apgar_5() As Long
   Apgar_5 = ml_Apgar_5
End Property
Property Let Apgar_1(lValue As Long)
   ml_Apgar_1 = lValue
End Property
Property Get Apgar_1() As Long
   Apgar_1 = ml_Apgar_1
End Property

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let IdAtencion(lValue As Long)
   ml_IdAtencion = lValue
End Property
Property Get IdAtencion() As Long
   IdAtencion = ml_IdAtencion
End Property
Property Let IdCondicionRN(lValue As Long)
   ml_IdCondicionRN = lValue
End Property
Property Get IdCondicionRN() As Long
   IdCondicionRN = ml_IdCondicionRN
End Property
Property Let IdTipoSexo(lValue As Long)
   ml_IdTipoSexo = lValue
End Property
Property Get IdTipoSexo() As Long
   IdTipoSexo = ml_IdTipoSexo
End Property
Property Let Peso(dValue As Double)
   md_Peso = dValue
End Property
Property Get Peso() As Double
   Peso = md_Peso
End Property
Property Let Talla(dValue As Double)
   md_Talla = dValue
End Property
Property Get Talla() As Double
   Talla = md_Talla
End Property
Property Let EdadSemanas(lValue As Long)
   ml_EdadSemanas = lValue
End Property
Property Get EdadSemanas() As Long
   EdadSemanas = ml_EdadSemanas
End Property
Property Let FechaNacimiento(daValue As Date)
   mda_FechaNacimiento = daValue
End Property
Property Get FechaNacimiento() As Date
   FechaNacimiento = mda_FechaNacimiento
End Property
Property Let IdNacimiento(lValue As Long)
   ml_IdNacimiento = lValue
End Property
Property Get IdNacimiento() As Long
   IdNacimiento = ml_IdNacimiento
End Property


