VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSeleccionaAnexoFUA2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: MINSA - Oficina Informatica y Telecomunicaciones
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para seleccionar el anexo del FUA2
'        Programado por: CACHAY F
'        Fecha: AGOSTO 2015
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mc_SeleccionaAnexoFua2 As New SIGHSis.SeleccionaAnexoFua2
Dim mi_respuesta As Integer
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
Dim ml_IdServicio As Long

Property Get Respuesta() As Integer
    Respuesta = mi_respuesta
End Property
Property Let Respuesta(lValue As Integer)
   mi_respuesta = lValue
End Property
Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property
Property Let IdServicio(lValue As Long)
   ml_IdServicio = lValue
End Property

Sub MostrarFormulario()
    mc_SeleccionaAnexoFua2.lcNombrePc = mo_lcNombrePc
    mc_SeleccionaAnexoFua2.IdServicio = ml_IdServicio
    mc_SeleccionaAnexoFua2.lnIdTablaLISTBARITEMS = mo_lnIdTablaLISTBARITEMS
    mc_SeleccionaAnexoFua2.Respuesta = mi_respuesta
    mc_SeleccionaAnexoFua2.Show 1
    mi_respuesta = mc_SeleccionaAnexoFua2.Respuesta
    Set mc_SeleccionaAnexoFua2 = Nothing
End Sub

