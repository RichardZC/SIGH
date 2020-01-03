VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Farm_FormDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: farm_formdet
'        Autor: Daniel Enrique Barrantes
'        Fecha: 11/08/2017 12:39:34 p.m.
'        Auspiciado por: Usaid/Politicas en Salud
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------

Dim mo_Conexion As ADODB.Connection
Dim ms_MensajeError As String
Property Set Conexion(oValue As ADODB.Connection)
   Set mo_Conexion = oValue
End Property
Property Get Conexion() As ADODB.Connection
   Set Conexion = mo_Conexion
End Property
Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property
Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property
'------------------------------------------------------------------------------------
'   Función:        Insertar
'   Descripción:    Inserta un registro a la tabla farm_formdet
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As Dofarm_formDet) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "farm_formdetAgregar"
   Set oParameter = .CreateParameter("@CODIGO_EJE", adVarChar, adParamInput, 3, IIf(oTabla.CODIGO_EJE = "", Null, oTabla.CODIGO_EJE))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CODIGO_PRE", adVarChar, adParamInput, 11, IIf(oTabla.CODIGO_PRE = "", Null, oTabla.CODIGO_PRE))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TIPSUM", adVarChar, adParamInput, 1, IIf(oTabla.TIPSUM = "", Null, oTabla.TIPSUM))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ANNOMES", adVarChar, adParamInput, 6, IIf(oTabla.ANNOMES = "", Null, oTabla.ANNOMES))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CODIGO_MED", adVarChar, adParamInput, 7, IIf(oTabla.CODIGO_MED = "", Null, oTabla.CODIGO_MED))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@SALDO", adInteger, adParamInput, 0, oTabla.SALDO)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PRECIO", adCurrency, adParamInput, 0, oTabla.precio)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@INGRE", adInteger, adParamInput, 0, oTabla.INGRE)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@REINGRE", adInteger, adParamInput, 0, oTabla.REINGRE)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@VENTA", adInteger, adParamInput, 0, oTabla.VENTA)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@SIS", adInteger, adParamInput, 0, oTabla.SIS)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@INTERSAN", adInteger, adParamInput, 0, oTabla.INTERSAN)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FAC_PERD", adInteger, adParamInput, 0, oTabla.FAC_PERD)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DEFNAC", adInteger, adParamInput, 0, oTabla.DEFNAC)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EXO", adInteger, adParamInput, 0, oTabla.EXO)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@SOAT", adInteger, adParamInput, 0, oTabla.SOAT)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CREDHOSP", adInteger, adParamInput, 0, oTabla.CREDHOSP)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@OTR_CONV", adInteger, adParamInput, 0, oTabla.OTR_CONV)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DEVOL", adInteger, adParamInput, 0, oTabla.DEVOL)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@VENCIDO", adInteger, adParamInput, 0, oTabla.VENCIDO)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@MERMA", adInteger, adParamInput, 0, oTabla.MERMA)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DISTRI", adInteger, adParamInput, 0, oTabla.DISTRI)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TRANSF", adInteger, adParamInput, 0, oTabla.TRANSF)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@VENTAINST", adInteger, adParamInput, 0, oTabla.VENTAINST)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DEV_VEN", adInteger, adParamInput, 0, oTabla.DEV_VEN)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DEV_MERMA", adInteger, adParamInput, 0, oTabla.DEV_MERMA)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@OTRAS_SAL", adInteger, adParamInput, 0, oTabla.OTRAS_SAL)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@STOCK_FIN", adInteger, adParamInput, 0, oTabla.STOCK_FIN)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@STOCK_FIN1", adInteger, adParamInput, 0, oTabla.STOCK_FIN1)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@REQ", adInteger, adParamInput, 0, oTabla.REQ)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TOTAL", adInteger, adParamInput, 0, oTabla.Total)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FEC_EXP", adDBTimeStamp, adParamInput, 0, oTabla.FEC_EXP)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DO_SALDO", adInteger, adParamInput, 0, oTabla.DO_SALDO)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DO_INGRE", adInteger, adParamInput, 0, oTabla.DO_INGRE)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DO_CON", adInteger, adParamInput, 0, oTabla.DO_CON)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DO_OTR", adInteger, adParamInput, 0, oTabla.DO_OTR)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DO_TOT", adInteger, adParamInput, 0, oTabla.DO_TOT)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DO_STK", adInteger, adParamInput, 0, oTabla.DO_STK)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DO_FECEXP", adDBTimeStamp, adParamInput, 0, IIf(oTabla.DO_FECEXP = 0, Null, oTabla.DO_FECEXP))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FECHA", adDBTimeStamp, adParamInput, 0, IIf(oTabla.Fecha = 0, Null, oTabla.Fecha))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@USUARIO", adVarChar, adParamInput, 15, IIf(oTabla.Usuario = "", Null, oTabla.Usuario))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@INDIPROC", adVarChar, adParamInput, 1, IIf(oTabla.INDIPROC = "", Null, oTabla.INDIPROC))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@SIT", adVarChar, adParamInput, 1, IIf(oTabla.SIT = "", Null, oTabla.SIT))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@INDISIGA", adVarChar, adParamInput, 1, IIf(oTabla.INDISIGA = "", Null, oTabla.INDISIGA))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DSTKCERO", adInteger, adParamInput, 0, oTabla.DSTKCERO)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@MPTOREPO", adInteger, adParamInput, 0, oTabla.MPTOREPO)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
End With
 
Insertar = True
ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function







'------------------------------------------------------------------------------------
'   Función:        Modificar
'   Descripción:    modifica un registro a la tabla farm_formdet
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As Dofarm_formDet) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "farm_formdetModificar"
   Set oParameter = .CreateParameter("@CODIGO_EJE", adVarChar, adParamInput, 3, IIf(oTabla.CODIGO_EJE = "", Null, oTabla.CODIGO_EJE))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CODIGO_PRE", adVarChar, adParamInput, 11, IIf(oTabla.CODIGO_PRE = "", Null, oTabla.CODIGO_PRE))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TIPSUM", adVarChar, adParamInput, 1, IIf(oTabla.TIPSUM = "", Null, oTabla.TIPSUM))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ANNOMES", adVarChar, adParamInput, 6, IIf(oTabla.ANNOMES = "", Null, oTabla.ANNOMES))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CODIGO_MED", adVarChar, adParamInput, 7, IIf(oTabla.CODIGO_MED = "", Null, oTabla.CODIGO_MED))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@SALDO", adInteger, adParamInput, 0, oTabla.SALDO)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PRECIO", adCurrency, adParamInput, 0, oTabla.precio)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@INGRE", adInteger, adParamInput, 0, oTabla.INGRE)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@REINGRE", adInteger, adParamInput, 0, oTabla.REINGRE)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@VENTA", adInteger, adParamInput, 0, oTabla.VENTA)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@SIS", adInteger, adParamInput, 0, oTabla.SIS)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@INTERSAN", adInteger, adParamInput, 0, oTabla.INTERSAN)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FAC_PERD", adInteger, adParamInput, 0, oTabla.FAC_PERD)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DEFNAC", adInteger, adParamInput, 0, oTabla.DEFNAC)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EXO", adInteger, adParamInput, 0, oTabla.EXO)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@SOAT", adInteger, adParamInput, 0, oTabla.SOAT)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CREDHOSP", adInteger, adParamInput, 0, oTabla.CREDHOSP)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@OTR_CONV", adInteger, adParamInput, 0, oTabla.OTR_CONV)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DEVOL", adInteger, adParamInput, 0, oTabla.DEVOL)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@VENCIDO", adInteger, adParamInput, 0, oTabla.VENCIDO)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@MERMA", adInteger, adParamInput, 0, oTabla.MERMA)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DISTRI", adInteger, adParamInput, 0, oTabla.DISTRI)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TRANSF", adInteger, adParamInput, 0, oTabla.TRANSF)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@VENTAINST", adInteger, adParamInput, 0, oTabla.VENTAINST)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DEV_VEN", adInteger, adParamInput, 0, oTabla.DEV_VEN)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DEV_MERMA", adInteger, adParamInput, 0, oTabla.DEV_MERMA)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@OTRAS_SAL", adInteger, adParamInput, 0, oTabla.OTRAS_SAL)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@STOCK_FIN", adInteger, adParamInput, 0, oTabla.STOCK_FIN)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@STOCK_FIN1", adInteger, adParamInput, 0, oTabla.STOCK_FIN1)
   .Parameters.Append oParameter
   
   
   Set oParameter = .CreateParameter("@REQ", adInteger, adParamInput, 0, IIf(oTabla.REQ = 0, Null, oTabla.REQ))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TOTAL", adInteger, adParamInput, 0, oTabla.Total)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FEC_EXP", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FEC_EXP = 0, Null, oTabla.FEC_EXP))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DO_SALDO", adInteger, adParamInput, 0, IIf(oTabla.DO_SALDO = 0, 0, oTabla.DO_SALDO))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DO_INGRE", adInteger, adParamInput, 0, IIf(oTabla.DO_INGRE = 0, 0, oTabla.DO_INGRE))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DO_CON", adInteger, adParamInput, 0, IIf(oTabla.DO_CON = 0, 0, oTabla.DO_CON))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DO_OTR", adInteger, adParamInput, 0, IIf(oTabla.DO_OTR = 0, 0, oTabla.DO_OTR))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DO_TOT", adInteger, adParamInput, 0, IIf(oTabla.DO_TOT = 0, 0, oTabla.DO_TOT))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DO_STK", adInteger, adParamInput, 0, IIf(oTabla.DO_STK = 0, 0, oTabla.DO_STK))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DO_FECEXP", adDBTimeStamp, adParamInput, 0, IIf(oTabla.DO_FECEXP = 0, Null, oTabla.DO_FECEXP))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FECHA", adDBTimeStamp, adParamInput, 0, IIf(oTabla.Fecha = 0, Null, oTabla.Fecha))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@USUARIO", adVarChar, adParamInput, 15, IIf(oTabla.Usuario = "", Null, oTabla.Usuario))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@INDIPROC", adVarChar, adParamInput, 1, IIf(oTabla.INDIPROC = "", Null, oTabla.INDIPROC))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@SIT", adVarChar, adParamInput, 1, IIf(oTabla.SIT = "", Null, oTabla.SIT))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@INDISIGA", adVarChar, adParamInput, 1, IIf(oTabla.INDISIGA = "", Null, oTabla.INDISIGA))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DSTKCERO", adInteger, adParamInput, 0, IIf(oTabla.DSTKCERO = 0, 0, oTabla.DSTKCERO))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@MPTOREPO", adInteger, adParamInput, 0, IIf(oTabla.MPTOREPO = 0, 0, oTabla.MPTOREPO))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
End With
 
Modificar = True
ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
Resume
End Function

'------------------------------------------------------------------------------------
'   Función:        Eliminar
'   Descripción:    elimia o borra un registro a la tabla farm_formdet
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As Dofarm_formDet) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "farm_formdetEliminar"
   Set oParameter = .CreateParameter("@CODIGO_EJE", adVarChar, adParamInput, 3, IIf(oTabla.CODIGO_EJE = "", Null, oTabla.CODIGO_EJE))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CODIGO_PRE", adVarChar, adParamInput, 11, IIf(oTabla.CODIGO_PRE = "", Null, oTabla.CODIGO_PRE))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TIPSUM", adVarChar, adParamInput, 1, IIf(oTabla.TIPSUM = "", Null, oTabla.TIPSUM))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ANNOMES", adVarChar, adParamInput, 6, IIf(oTabla.ANNOMES = "", Null, oTabla.ANNOMES))
   .Parameters.Append oParameter
   .Execute
End With
 
Eliminar = True
ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

'------------------------------------------------------------------------------------
'   Función:        SeleccionarPorId
'   Descripción:    Selecciona un unico registro de la tabla farm_formdet
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As Dofarm_formDet) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "farm_formdetSeleccionarPorId"
     Set oParameter = .CreateParameter("@CODIGO_EJE", adVarChar, adParamInput, 3, oTabla.CODIGO_EJE)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       CargaTabla oTabla, oRecordset
       SeleccionarPorId = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As Dofarm_formDet)
   oTabla.CODIGO_EJE = ""
   oTabla.CODIGO_PRE = ""
   oTabla.TIPSUM = ""
   oTabla.ANNOMES = ""
   oTabla.CODIGO_MED = ""
   oTabla.SALDO = 0
   oTabla.precio = 0
   oTabla.INGRE = 0
   oTabla.REINGRE = 0
   oTabla.VENTA = 0
   oTabla.SIS = 0
   oTabla.INTERSAN = 0
   oTabla.FAC_PERD = 0
   oTabla.DEFNAC = 0
   oTabla.EXO = 0
   oTabla.SOAT = 0
   oTabla.CREDHOSP = 0
   oTabla.OTR_CONV = 0
   oTabla.DEVOL = 0
   oTabla.VENCIDO = 0
   oTabla.MERMA = 0
   oTabla.DISTRI = 0
   oTabla.TRANSF = 0
   oTabla.VENTAINST = 0
   oTabla.DEV_VEN = 0
   oTabla.DEV_MERMA = 0
   oTabla.OTRAS_SAL = 0
   oTabla.STOCK_FIN = 0
   oTabla.STOCK_FIN1 = 0
   oTabla.REQ = 0
   oTabla.Total = 0
   oTabla.FEC_EXP = Now
   oTabla.DO_SALDO = 0
   oTabla.DO_INGRE = 0
   oTabla.DO_CON = 0
   oTabla.DO_OTR = 0
   oTabla.DO_TOT = 0
   oTabla.DO_STK = 0
   oTabla.DO_FECEXP = Now
   oTabla.Fecha = Now
   oTabla.Usuario = ""
   oTabla.INDIPROC = ""
   oTabla.SIT = ""
   oTabla.INDISIGA = ""
   oTabla.DSTKCERO = 0
   oTabla.MPTOREPO = 0
End Sub

Function EliminarPorCodigo(ByVal oTabla As Dofarm_formDet) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
EliminarPorCodigo = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "Farm_formDetEliminarXproducto"
    Set oParameter = .CreateParameter("@Producto", adVarChar, adParamInput, 7, oTabla.CODIGO_MED): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@Almacen", adVarChar, adParamInput, 11, oTabla.CODIGO_PRE): .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@FechaMov", adVarChar, adParamInput, 6, oTabla.ANNOMES): .Parameters.Append oParameter
   .Execute
End With
 
EliminarPorCodigo = True
ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function


Function SeleccionarPorCodigo(ByVal oTabla As Dofarm_formDet) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorCodigo = False
   With oCommand
         .CommandType = adCmdStoredProc
         Set .ActiveConnection = mo_Conexion
        .CommandText = "Farm_formDetSeleccionarUltimoSaldoPorIdproductoXmes"
        Set oParameter = .CreateParameter("@Producto", adVarChar, adParamInput, 7, oTabla.CODIGO_MED): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Almacen", adVarChar, adParamInput, 11, oTabla.CODIGO_PRE): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaMov", adVarChar, adParamInput, 6, oTabla.ANNOMES): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
        CargaTabla oTabla, oRecordset
        SeleccionarPorCodigo = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function

End Function

Sub CargaTabla(ByRef oTabla As Dofarm_formDet, oRecordset As Recordset)
       oTabla.CODIGO_EJE = IIf(IsNull(oRecordset!CODIGO_EJE), "", oRecordset!CODIGO_EJE)
       oTabla.CODIGO_PRE = IIf(IsNull(oRecordset!CODIGO_PRE), "", oRecordset!CODIGO_PRE)
       oTabla.TIPSUM = IIf(IsNull(oRecordset!TIPSUM), "", oRecordset!TIPSUM)
       oTabla.ANNOMES = IIf(IsNull(oRecordset!ANNOMES), "", oRecordset!ANNOMES)
       oTabla.CODIGO_MED = IIf(IsNull(oRecordset!CODIGO_MED), "", oRecordset!CODIGO_MED)
       oTabla.SALDO = IIf(IsNull(oRecordset!SALDO), 0, oRecordset!SALDO)
       oTabla.precio = IIf(IsNull(oRecordset!precio), 0, oRecordset!precio)
       oTabla.INGRE = IIf(IsNull(oRecordset!INGRE), 0, oRecordset!INGRE)
       oTabla.REINGRE = IIf(IsNull(oRecordset!REINGRE), 0, oRecordset!REINGRE)
       oTabla.VENTA = IIf(IsNull(oRecordset!VENTA), 0, oRecordset!VENTA)
       oTabla.SIS = IIf(IsNull(oRecordset!SIS), 0, oRecordset!SIS)
       oTabla.INTERSAN = IIf(IsNull(oRecordset!INTERSAN), 0, oRecordset!INTERSAN)
       oTabla.FAC_PERD = IIf(IsNull(oRecordset!FAC_PERD), 0, oRecordset!FAC_PERD)
       oTabla.DEFNAC = IIf(IsNull(oRecordset!DEFNAC), 0, oRecordset!DEFNAC)
       oTabla.EXO = IIf(IsNull(oRecordset!EXO), 0, oRecordset!EXO)
       oTabla.SOAT = IIf(IsNull(oRecordset!SOAT), 0, oRecordset!SOAT)
       oTabla.CREDHOSP = IIf(IsNull(oRecordset!CREDHOSP), 0, oRecordset!CREDHOSP)
       oTabla.OTR_CONV = IIf(IsNull(oRecordset!OTR_CONV), 0, oRecordset!OTR_CONV)
       oTabla.DEVOL = IIf(IsNull(oRecordset!DEVOL), 0, oRecordset!DEVOL)
       oTabla.VENCIDO = IIf(IsNull(oRecordset!VENCIDO), 0, oRecordset!VENCIDO)
       oTabla.MERMA = IIf(IsNull(oRecordset!MERMA), 0, oRecordset!MERMA)
       oTabla.DISTRI = IIf(IsNull(oRecordset!DISTRI), 0, oRecordset!DISTRI)
       oTabla.TRANSF = IIf(IsNull(oRecordset!TRANSF), 0, oRecordset!TRANSF)
       oTabla.VENTAINST = IIf(IsNull(oRecordset!VENTAINST), 0, oRecordset!VENTAINST)
       oTabla.DEV_VEN = IIf(IsNull(oRecordset!DEV_VEN), 0, oRecordset!DEV_VEN)
       oTabla.DEV_MERMA = IIf(IsNull(oRecordset!DEV_MERMA), 0, oRecordset!DEV_MERMA)
       oTabla.OTRAS_SAL = IIf(IsNull(oRecordset!OTRAS_SAL), 0, oRecordset!OTRAS_SAL)
       oTabla.STOCK_FIN = IIf(IsNull(oRecordset!STOCK_FIN), 0, oRecordset!STOCK_FIN)
       oTabla.STOCK_FIN1 = IIf(IsNull(oRecordset!STOCK_FIN1), 0, oRecordset!STOCK_FIN1)
       oTabla.REQ = IIf(IsNull(oRecordset!REQ), 0, oRecordset!REQ)
       oTabla.Total = IIf(IsNull(oRecordset!Total), 0, oRecordset!Total)
       oTabla.FEC_EXP = IIf(IsNull(oRecordset!FEC_EXP), 0, oRecordset!FEC_EXP)
       oTabla.DO_SALDO = IIf(IsNull(oRecordset!DO_SALDO), 0, oRecordset!DO_SALDO)
       oTabla.DO_INGRE = IIf(IsNull(oRecordset!DO_INGRE), 0, oRecordset!DO_INGRE)
       oTabla.DO_CON = IIf(IsNull(oRecordset!DO_CON), 0, oRecordset!DO_CON)
       oTabla.DO_OTR = IIf(IsNull(oRecordset!DO_OTR), 0, oRecordset!DO_OTR)
       oTabla.DO_TOT = IIf(IsNull(oRecordset!DO_TOT), 0, oRecordset!DO_TOT)
       oTabla.DO_STK = IIf(IsNull(oRecordset!DO_STK), 0, oRecordset!DO_STK)
       oTabla.DO_FECEXP = IIf(IsNull(oRecordset!DO_FECEXP), 0, oRecordset!DO_FECEXP)
       oTabla.Fecha = IIf(IsNull(oRecordset!Fecha), 0, oRecordset!Fecha)
       oTabla.Usuario = IIf(IsNull(oRecordset!Usuario), "", oRecordset!Usuario)
       oTabla.INDIPROC = IIf(IsNull(oRecordset!INDIPROC), "", oRecordset!INDIPROC)
       oTabla.SIT = IIf(IsNull(oRecordset!SIT), "", oRecordset!SIT)
       oTabla.INDISIGA = IIf(IsNull(oRecordset!INDISIGA), "", oRecordset!INDISIGA)
       oTabla.DSTKCERO = IIf(IsNull(oRecordset!DSTKCERO), 0, oRecordset!DSTKCERO)
       oTabla.MPTOREPO = IIf(IsNull(oRecordset!MPTOREPO), 0, oRecordset!MPTOREPO)

End Sub


'------------------------------------------------------------------------------------
'        Fin de código autogenerado para la clase: farm_formdet
'        Autor:   Daniel Enrique Barrantes
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'        Inicio de código manual para la clase: farm_formdet
'        Autor:   DAniel Barrantes B.
'        Fecha: 11/08/2017 12:39:35 p.m.
'------------------------------------------------------------------------------------
 

