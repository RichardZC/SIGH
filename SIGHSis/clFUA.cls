VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clFUA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para llamar al Formato FUA
'        Programado por: Barrantes D
'        Fecha: Enero 2013
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mi_opcion As sghOpciones
Dim ml_idUsuario As Long
Dim mo_lcNombrePc As String
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim ml_IdCuentaAtencion As Long
Dim mo_SISFuaV1 As New SisFua
Dim mo_SISFuaV2 As New SisFuaVersion2
Dim mo_lbEsAltaMedica As Boolean
Dim ml_IdServicio As Long
Dim mo_ReglasSISgalenhos As New ReglasSISgalenhos
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mc_FuaVersionFormato As String
Dim mi_FuaTipoAnexo2015 As Integer
Dim mo_SoloImprimeFUAyaGrabado As Boolean
Dim ml_CodigoPrestacion As String

Property Let CodigoPrestacion(lValue As String)
    ml_CodigoPrestacion = lValue
End Property

Property Let SoloImprimeFUAyaGrabado(lValue As Boolean)
   mo_SoloImprimeFUAyaGrabado = lValue
End Property
Property Let EsAltaMedica(lValue As Boolean)
   mo_lbEsAltaMedica = lValue
End Property

Property Let IdCuentaAtencion(lValue As Long)
   ml_IdCuentaAtencion = lValue
End Property

Property Get Opcion() As sghOpciones
    Opcion = mi_opcion
End Property
Property Let Opcion(lValue As sghOpciones)
   mi_opcion = lValue
End Property

Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
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

Property Let FuaVersionFormato(lValue As String)
   mc_FuaVersionFormato = lValue
End Property

Property Let FuaTipoAnexo2015(lValue As Integer)
   mi_FuaTipoAnexo2015 = lValue
End Property

Sub MostrarFormulario()
    If ml_IdServicio <> 0 Then
        If Not SeRegistroFua Then
            mc_FuaVersionFormato = lcBuscaParametro.SeleccionaFilaParametro(358)
            ObtenerValidarFuaTipoAnexo2015 mi_FuaTipoAnexo2015
            If mi_FuaTipoAnexo2015 = 0 Then Exit Sub
        End If
    End If

    If mc_FuaVersionFormato = "A" Then '1er Formato FUA
        
        mo_SISFuaV1.IdCuentaAtencion = ml_IdCuentaAtencion
        mo_SISFuaV1.Opcion = mi_opcion
        mo_SISFuaV1.idUsuario = ml_idUsuario
        mo_SISFuaV1.lcNombrePc = mo_lcNombrePc
        mo_SISFuaV1.lnIdTablaLISTBARITEMS = mo_lnIdTablaLISTBARITEMS
        mo_SISFuaV1.EsAltaMedica = mo_lbEsAltaMedica
        mo_SISFuaV1.Show 1
           
        Set mo_SISFuaV1 = Nothing
    ElseIf mc_FuaVersionFormato = "B" Then '2do Formato FUA
        mo_SISFuaV2.CodigoPrestacion = ml_CodigoPrestacion
        mo_SISFuaV2.SoloImprimeFUAyaGrabado = mo_SoloImprimeFUAyaGrabado
        mo_SISFuaV2.IdCuentaAtencion = ml_IdCuentaAtencion
        mo_SISFuaV2.Opcion = mi_opcion
        mo_SISFuaV2.idUsuario = ml_idUsuario
        mo_SISFuaV2.lcNombrePc = mo_lcNombrePc
        mo_SISFuaV2.lnIdTablaLISTBARITEMS = mo_lnIdTablaLISTBARITEMS
        mo_SISFuaV2.EsAltaMedica = mo_lbEsAltaMedica
        mo_SISFuaV2.FuaVersionFormato = mc_FuaVersionFormato
        mo_SISFuaV2.FuaTipoAnexo2015 = mi_FuaTipoAnexo2015
        mo_SISFuaV2.Show 1
        
        Set mo_SISFuaV2 = Nothing
    End If
End Sub

'HRA 10/12/2020 Cambio 47 Inicio
Sub MostrarFormularioF()
    If ml_IdServicio <> 0 Then
        If Not SeRegistroFua Then
            mc_FuaVersionFormato = lcBuscaParametro.SeleccionaFilaParametro(358)
            ObtenerValidarFuaTipoAnexo2015 mi_FuaTipoAnexo2015
            If mi_FuaTipoAnexo2015 = 0 Then Exit Sub
        End If
    End If
    If mc_FuaVersionFormato = "A" Then '1er Formato FUA
        mo_SISFuaV1.IdCuentaAtencion = ml_IdCuentaAtencion
        mo_SISFuaV1.Opcion = mi_opcion
        mo_SISFuaV1.idUsuario = ml_idUsuario
        mo_SISFuaV1.lcNombrePc = mo_lcNombrePc
        mo_SISFuaV1.lnIdTablaLISTBARITEMS = mo_lnIdTablaLISTBARITEMS
        mo_SISFuaV1.EsAltaMedica = mo_lbEsAltaMedica
        mo_SISFuaV1.Show 1
        Set mo_SISFuaV1 = Nothing
    ElseIf mc_FuaVersionFormato = "B" Then '2do Formato FUA
        mo_SISFuaV2.CodigoPrestacion = ml_CodigoPrestacion
        mo_SISFuaV2.SoloImprimeFUAyaGrabado = mo_SoloImprimeFUAyaGrabado
        mo_SISFuaV2.IdCuentaAtencion = ml_IdCuentaAtencion
        mo_SISFuaV2.Opcion = mi_opcion
        mo_SISFuaV2.idUsuario = ml_idUsuario
        mo_SISFuaV2.lcNombrePc = mo_lcNombrePc
        mo_SISFuaV2.lnIdTablaLISTBARITEMS = mo_lnIdTablaLISTBARITEMS
        mo_SISFuaV2.EsAltaMedica = mo_lbEsAltaMedica
        mo_SISFuaV2.FuaVersionFormato = mc_FuaVersionFormato
        mo_SISFuaV2.FuaTipoAnexo2015 = mi_FuaTipoAnexo2015
        mo_SISFuaV2.GuardarFua = "S"
        mo_SISFuaV2.Show 1
        Set mo_SISFuaV2 = Nothing
    End If
End Sub
'HRA 10/12/2020 Cambio 47 Fin

Function SeRegistroFua() As Boolean
    Dim oRsTmp1 As New Recordset
    Set oRsTmp1 = mo_ReglasSISgalenhos.SisFuaAtencionSeleccionarPorCuenta(ml_IdCuentaAtencion)
    SeRegistroFua = False
    If oRsTmp1.RecordCount > 0 Then
        mc_FuaVersionFormato = oRsTmp1.Fields!FuaVersionFormato
        mi_FuaTipoAnexo2015 = oRsTmp1.Fields!FuaTipoAnexo2015
        SeRegistroFua = True
    End If
    Set oRsTmp1 = Nothing
End Function

Sub ObtenerValidarFuaTipoAnexo2015(ByRef IdFuaTipoAnexo As Integer)
    IdFuaTipoAnexo = 0
    If lcBuscaParametro.SeleccionaFilaParametro(358) = "B" Then
        If lcBuscaParametro.SeleccionaFilaParametro(359) = "3" Then
            Dim doServicio As New doServicio
            Set doServicio = RetornaServicio(ml_IdServicio)
            If IsNull(doServicio.FuaTipoAnexo2015) Or doServicio.FuaTipoAnexo2015 = 0 Or doServicio.FuaTipoAnexo2015 = 3 Then
                Dim oSeleccionFUA As New SIGHSis.clsSeleccionaAnexoFUA2
                oSeleccionFUA.lcNombrePc = mo_lcNombrePc
                oSeleccionFUA.lnIdTablaLISTBARITEMS = mo_lnIdTablaLISTBARITEMS
                oSeleccionFUA.IdServicio = ml_IdServicio
                oSeleccionFUA.MostrarFormulario
                IdFuaTipoAnexo = oSeleccionFUA.Respuesta
                Set oSeleccionFUA = Nothing
            Else
                IdFuaTipoAnexo = doServicio.FuaTipoAnexo2015
            End If
        Else
            IdFuaTipoAnexo = lcBuscaParametro.SeleccionaFilaParametro(359)
        End If
    Else
        IdFuaTipoAnexo = 3
    End If
End Sub

Function RetornaServicio(ml_IdServ As Long) As doServicio
    Dim oConexion As New Connection
    Dim oDoServicio As New doServicio
    Dim oServicio As New Servicios
    
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    
    Set oServicio.Conexion = oConexion
    oDoServicio.IdServicio = ml_IdServ
    If oServicio.SeleccionarPorId(oDoServicio) Then
    End If
    
    oConexion.Close
    Set oConexion = Nothing
    Set oServicio = Nothing
    Set RetornaServicio = oDoServicio
End Function
