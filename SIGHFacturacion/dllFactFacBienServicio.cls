VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "dllFactFacBienServicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para Facturación
'        Programado por: Barrantes D
'        Fecha: Enero 2011
'
'------------------------------------------------------------------------------------
Option Explicit
Function RetornaDatosDeCabecera(lnIdCuentaAtencion As Long) As ADODB.Recordset
    Dim oRs As New ADODB.Recordset
    Dim lcSql As String
    Dim lcBuscaParametro As New SIGHDatos.Parametros
    Dim oReglasComunes As New ReglasComunes
    Dim oReglasFarmacia As New ReglasFarmacia
    Set RetornaDatosDeCabecera = Nothing
    '******************* Asegura que la "Atencion" sea CREDITO HOSPITALARIO**************
    If lcBuscaParametro.SeleccionaFilaParametro(231) = "N" Then
        oReglasComunes.AtencionesActualizaCreditoHospitalario lnIdCuentaAtencion
    End If
    '******************* Asegura que la "Atencion" sea CREDITO HOSPITALARIO**************
    Set RetornaDatosDeCabecera = oReglasFarmacia.FacturacionCuentasAtencionFiltrarXcuenta(lnIdCuentaAtencion)
    Set lcBuscaParametro = Nothing
    Set oReglasComunes = Nothing
    Set oReglasFarmacia = Nothing
End Function

