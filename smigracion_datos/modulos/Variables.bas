Attribute VB_Name = "Variables"
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Módulo para declaración de variables públicas
'        Programado por: Barrantes D
'        Fecha: Enero 2010
'
'------------------------------------------------------------------------------------
Public Const SW_SHOWNORMAL = 1
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const wxTipoBD As String = "1"         '1->Access, 2->SQL Server
Public Const wxSistema As String = "CR1111"    'CR<-para generar ConfRegional.exe
Public wxImagen As PictureBox
Public wxArchivoElegido As String       'Archivo Imagen elegido desde Disco
Public wxProceso As Boolean
Public wxConexion As New ADODB.Connection         'Conexion a la BD LolCli
Public wxConexionRed As New ADODB.Connection      'Conexion a la BD GalenHos
Public wxOpcMenu As String     'Opcion elegida del Menu Principal
Public wxMant As String          'Mantenimiento (0-Nuevo, 1-Modificar, 2-Borrar)

Public wxCodigo As Variant, wxDescripcion As Variant, wxFecha As Date, wxPlaca As String, wxPrecio As Double  'Codigo,Descripcion,Fecha, Precio Unitario  elegidos desde Busqueda de Tabla
Public wxUsuarioSist As String, wxCUsuaSist As String, wxUsuarioGalenhos As Long     'Usuario del Sistema
Public wxUsuarioExterno As Boolean 'Usuario diferente a ADMIN,RICARDO,EDICA
Public wxEmpresa As String, wxOficina  As String, wxRuc As String, wxDireccion As String, wxNum As String, wxDistrito As String, wxProvincia As String, wxDpto As String
Public Const wxSki As String = "winaqua.skn"
Public wxImporteHoraLibre As Double
Public lcFechaTrabajo As String
Public wxNumMaxHoraLibre As Double
Public wxFechaInicioCompetencia As Date, wxPremio As String 'Para Premios por feriados
Public wxVersionEnRed As Integer            'Version del SISTEMA   0<-MonoUsuario  1<-En RED
Public wxMinUltimoAvisoParaAlarma As Long   'Nro MINUTOS antes que acabe donde la ALARMA aparece
Public wxTime As String                     'HORA para todas las CABINAS
Public wxDate As Date                       'FECHA para todas las CABINAS
Public wxRutaBaseDatos As String            'Ruta donde se encuentra la BD en el SERVIDOR
Public wxBDGalenhos As String
Public wxBDLolcli As String
Public wxVersionBDactualizada As String
Public wxVersionSQL As String
