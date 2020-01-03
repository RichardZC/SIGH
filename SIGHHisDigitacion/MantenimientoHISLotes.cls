VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MantenimientoHISLotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para mantenimiento de HisLotes
'        Programado por: Cachay F
'        Fecha: Febrero 2014
'
'------------------------------------------------------------------------------------

Option Explicit
Dim mo_lcNombrePc As String
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_FormularioLotesHIS As New MantenimientoLotesHIS
Dim mi_Opcion As sghOpciones
Dim ml_IdEstablecimiento As Long
Dim ml_IdUsuario As Long
Dim ml_IdRegistroLote As Long

Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property

Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property

Property Get Opcion() As sghOpciones
    Opcion = mi_Opcion
End Property

Property Let Opcion(lValue As sghOpciones)
   mi_Opcion = lValue
End Property

Property Let IdUsuario(lValue As Long)
   ml_IdUsuario = lValue
End Property

Property Let IdEstablecimiento(lValue As Long)
   ml_IdEstablecimiento = lValue
End Property


Property Get IdRegistroLote() As Long
    IdRegistroLote = ml_IdRegistroLote
End Property

Property Let IdRegistroLote(lValue As Long)
   ml_IdRegistroLote = lValue
End Property

Sub MostrarFormulario()
    mo_FormularioLotesHIS.Opcion = mi_Opcion
    mo_FormularioLotesHIS.IdUsuario = ml_IdUsuario
    mo_FormularioLotesHIS.IdEstablecimiento = ml_IdEstablecimiento
    'mo_FormularioLotesHIS.lnIdTablaLISTBARITEMS = mo_lnIdTablaLISTBARITEMS
    'mo_FormularioLotesHIS.lcNombrePc = mo_lcNombrePc
    mo_FormularioLotesHIS.IdLote = ml_IdRegistroLote
    mo_FormularioLotesHIS.Show 1
    Set mo_FormularioLotesHIS = Nothing
End Sub

