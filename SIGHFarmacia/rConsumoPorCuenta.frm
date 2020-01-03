VERSION 5.00
Begin VB.Form rConsumoPorCuenta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consumo de Pacientes por N° Cuenta"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9330
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "rConsumoPorCuenta.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   9330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   45
      TabIndex        =   4
      Top             =   3540
      Width           =   9180
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "rConsumoPorCuenta.frx":0CCA
         DownPicture     =   "rConsumoPorCuenta.frx":118E
         Height          =   700
         Left            =   4740
         Picture         =   "rConsumoPorCuenta.frx":167A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "rConsumoPorCuenta.frx":1B66
         DownPicture     =   "rConsumoPorCuenta.frx":1FC6
         Height          =   700
         Left            =   3210
         Picture         =   "rConsumoPorCuenta.frx":243B
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1365
      End
   End
   Begin VB.Frame fraDatosHistoria 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3555
      Left            =   15
      TabIndex        =   3
      Top             =   0
      Width           =   9195
      Begin VB.CheckBox chkConsideraItemsPaquetes 
         Caption         =   "Los PAQUETES saldrán en Medicamentos/Insumos"
         Height          =   255
         Left            =   2940
         TabIndex        =   18
         Top             =   1620
         Visible         =   0   'False
         Width           =   6015
      End
      Begin VB.Frame Frame1 
         Caption         =   "Reporte"
         Height          =   2055
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   2775
         Begin VB.ComboBox CmbMes 
            Enabled         =   0   'False
            Height          =   330
            ItemData        =   "rConsumoPorCuenta.frx":28B0
            Left            =   600
            List            =   "rConsumoPorCuenta.frx":28D8
            TabIndex        =   16
            Top             =   1440
            Width           =   2070
         End
         Begin VB.OptionButton optConsumoConsolidado 
            Caption         =   "Consumo Mensual"
            Height          =   375
            Left            =   255
            TabIndex        =   15
            Top             =   840
            Width           =   1815
         End
         Begin VB.OptionButton optConsumoDet 
            Caption         =   "Consumo Detallado"
            Height          =   375
            Left            =   240
            TabIndex        =   14
            Top             =   360
            Value           =   -1  'True
            Width           =   2295
         End
         Begin VB.Label Lblmes 
            Caption         =   "Mes"
            Height          =   375
            Left            =   240
            TabIndex        =   17
            Top             =   1440
            Width           =   735
         End
      End
      Begin VB.CheckBox chkFF 
         Caption         =   "Mostrar el Reporte en orden de FORMA DE PAGO"
         Height          =   255
         Left            =   2940
         TabIndex        =   12
         Top             =   1320
         Width           =   6015
      End
      Begin VB.CheckBox chkPagadasEnCaja 
         Caption         =   "Sólo cuentas PAGADAS EN CAJA"
         Height          =   495
         Left            =   8010
         TabIndex        =   11
         Top             =   2040
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.CommandButton cmdBuscaCuentaPorApellidos 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2610
         TabIndex        =   10
         ToolTipText     =   "Busca Cuenta por Apellidos y Nombres"
         Top             =   240
         Width           =   315
      End
      Begin VB.CheckBox chkExcel 
         Caption         =   "En Excel"
         Height          =   315
         Left            =   2940
         Picture         =   "rConsumoPorCuenta.frx":295A
         TabIndex        =   9
         Top             =   990
         Width           =   1995
      End
      Begin VB.TextBox txtNombrePaciente 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4440
         TabIndex        =   8
         Top             =   630
         Width           =   4545
      End
      Begin VB.TextBox txtNhistoria 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2940
         MaxLength       =   30
         TabIndex        =   7
         ToolTipText     =   "Ingrese el Nro de Historia Clínica"
         Top             =   630
         Width           =   1425
      End
      Begin VB.TextBox txtNcuenta 
         Height          =   315
         Left            =   1470
         MaxLength       =   30
         TabIndex        =   0
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtDatosDeCuenta 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2940
         TabIndex        =   5
         Top             =   240
         Width           =   6045
      End
      Begin VB.Label lblNcuenta 
         AutoSize        =   -1  'True
         Caption         =   "N° Cuenta"
         Height          =   210
         Left            =   120
         TabIndex        =   6
         Top             =   270
         Width           =   855
      End
   End
End
Attribute VB_Name = "rConsumoPorCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Reporte Consumo por Cuenta
'        Programado por: Barrantes D
'        Fecha: Febrero 2009
'
'------------------------------------------------------------------------------------
Option Explicit
'***************daniel barrantes**************
'***************Registro de datos de filtro para el Reporte
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
Dim ms_MensajeError As String
Dim mo_Teclado As New SIGHEntidades.Teclado
Dim ml_TextoDelFiltro As String
Dim ml_IdPaciente As Long
Dim ml_idFuenteFinanciamiento As Long
Dim ml_idAtencion As Long

Private Sub btnAceptar_Click()
'editado por mariano 14112014
    If ValidaDatosObligatorios Then
         Me.MousePointer = 11
         If chkFF.Value = 1 Then
            Dim oRepConsumoPorCuenta As New RepConsumoPorCuenta
            oRepConsumoPorCuenta.ReporteXformaFarmaceutica Val(txtNcuenta.Text), ml_TextoDelFiltro, Me.hwnd
         Else
            If optConsumoDet.Value = True Then 'detallado
                Dim oRptClase As New rCrystal
                oRptClase.EnArchivoExcel = IIf(chkExcel.Value = 1, True, False)
                oRptClase.IdCuenta = Val(txtNcuenta.Text)
                oRptClase.TextoDelFiltro = ml_TextoDelFiltro
                oRptClase.TipoReporte = Me.Name
                oRptClase.idFuenteFinanciamiento = ml_idFuenteFinanciamiento
                oRptClase.SoloPagados = IIf(chkPagadasEnCaja.Value = 1, True, False)
               'oRptClase.SoloConsolidado = IIf(optConsumoDet.Value = True, True, False)
                oRptClase.ConsideraItemsDePaquetes = IIf(Me.chkConsideraItemsPaquetes.Value = 1, True, False)
                oRptClase.Show vbModal
                Set oRptClase = Nothing
            Else 'mensualizado
                Dim oRptClase1 As New rCrytalInventario
                oRptClase1.EnArchivoExcel = IIf(chkExcel.Value = 1, True, False)
                oRptClase1.IdCuenta = Val(txtNcuenta.Text)
                oRptClase1.Mes = Left(Val(CmbMes.Text), 2)
                oRptClase1.IdAnio = Right((CmbMes.Text), 4)
                oRptClase1.TextoDelFiltro = ml_TextoDelFiltro
                oRptClase1.TipoReporte = Me.Name
                oRptClase1.idFuenteFinanciamiento = ml_idFuenteFinanciamiento
                oRptClase1.SoloPagados = IIf(chkPagadasEnCaja.Value = 1, True, False)
                'oRptClase1.ConsideraItemsDePaquetes = IIf(Me.chkConsideraItemsPaquetes.Value = 1, True, False)
                oRptClase1.Show vbModal
                Set oRptClase1 = Nothing
            End If
         End If
        Me.MousePointer = 1
    End If
End Sub

'kike 2017
Function DxPorCuenta() As String
    Dim oRsTmp As New Recordset
    DxPorCuenta = ""
    Set oRsTmp = mo_ReglasAdmision.AtencionesDiagnosticosSeleccionarTodosPorIdAtencion(ml_idAtencion)
    If oRsTmp.RecordCount > 0 Then
         DxPorCuenta = "  (Dx= " & Trim(oRsTmp.Fields!CodigoCIE2004) & " " & Trim(oRsTmp.Fields!Descripcion) & ")"
    End If
    Set oRsTmp = Nothing
End Function
'kike 2017
Function ValidaDatosObligatorios() As Boolean
    ms_MensajeError = ""
    ml_TextoDelFiltro = "FILTROS:   N° Cuenta: " & Trim(txtNcuenta.Text) & " - " & Trim(txtDatosDeCuenta.Text) & _
                        "    Paciente: " & HCigualDNI_DevuelveHistoriaConCerosIzquierda(txtNhistoria.Text, False) & " - " & txtNombrePaciente.Text & _
                        DxPorCuenta
    If txtDatosDeCuenta.Text = "" Then
        ms_MensajeError = ms_MensajeError + "Por favor ingrese el N° Cuenta"
        txtNcuenta.SetFocus
    End If
    'mariano 0711014
    If CmbMes.ListIndex < 0 And optConsumoConsolidado.Value Then
        ms_MensajeError = ms_MensajeError + "Por favor Seleccione mes"
        CmbMes.SetFocus
    End If
    
    If ms_MensajeError <> "" Then
       MsgBox ms_MensajeError, vbInformation, Me.Caption
       ValidaDatosObligatorios = False
    Else
       ValidaDatosObligatorios = True
    End If
End Function

Private Sub btnCancelar_Click()
    Me.Visible = False
    LimpiarVariablesDeMemoria
End Sub

Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
        Case vbKeyF6
'           ucListaProductos1.RealizarBusqueda
        Case vbKeyEscape
'           btnCancelar_Click
        Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub




Private Sub cmdBuscaCuentaPorApellidos_Click()
    Dim oBusqueda As New SIGHNegocios.BuscaPacientes
    Dim oDOPaciente As New DOPaciente
    Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
    Dim oConexion As New Connection
    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    oBusqueda.TipoFiltro = sghFiltrarTodos
    oBusqueda.MostrarFormulario
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDOPaciente = mo_AdminAdmision.PacientesSeleccionarPorId(oBusqueda.IdRegistroSeleccionado, oConexion)
        If Not oDOPaciente Is Nothing Then
            ml_IdPaciente = oDOPaciente.IdPaciente
            txtNhistoria.Text = oDOPaciente.NroHistoriaClinica
            txtNombrePaciente.Text = Trim(oDOPaciente.ApellidoPaterno) + " " + Trim(oDOPaciente.ApellidoMaterno) + " " + oDOPaciente.PrimerNombre
            Dim oRsTmp As New Recordset
            Set oRsTmp = mo_ReglasFarmacia.FacturacionCuentasAtencionSeleccionarPorIdPaciente(ml_IdPaciente, oConexion, True)
            If oRsTmp.RecordCount > 0 Then
               txtNcuenta.Text = oRsTmp.Fields!idCuentaAtencion
            End If
            oRsTmp.Close
            Set oRsTmp = Nothing
            txtNcuenta_LostFocus
        End If
    End If
    oConexion.Close
    Set oConexion = Nothing
End Sub




Private Sub Form_Unload(Cancel As Integer)
    LimpiarVariablesDeMemoria
End Sub

Private Sub optConsumoConsolidado_Click()
  CmbMes.Enabled = True 'Mariano 07112014
  chkConsideraItemsPaquetes.Visible = False
End Sub

Private Sub optConsumoDet_Click()
    CmbMes.Enabled = False 'Mariano 07112014
    CmbMes.Text = ""
    'chkConsideraItemsPaquetes.Visible = True
End Sub

Private Sub txtNcuenta_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtNcuenta

End Sub



Private Sub txtNcuenta_LostFocus()
    If mo_Teclado.TextoEsSoloNumeros(txtNcuenta.Text) Then
       Dim oRsTmp As New Recordset
       Dim oConexion As New Connection
       oConexion.Open SIGHEntidades.CadenaConexion
       oConexion.CursorLocation = adUseClient
       Set oRsTmp = mo_ReglasFarmacia.AtencionesSelecionarPorCuenta(txtNcuenta.Text, oConexion)
       If oRsTmp.RecordCount > 0 Then
          txtDatosDeCuenta.Text = "F.Ing: " & oRsTmp.Fields!fechaingreso & " - F.Egreso: " & oRsTmp.Fields!fechaEgreso & " - " & IIf(oRsTmp.Fields!IdTipoServicio = 1, "Consultorios Externos", IIf(oRsTmp.Fields!IdTipoServicio = 3, "Hospitalización", "Emergencia")) & " - (Est:" & Trim(oRsTmp.Fields!estadoCta) & ")   IAFA: " & oRsTmp.Fields!dFuenteFinanciamiento
          ml_IdPaciente = oRsTmp.Fields!IdPaciente
          txtNombrePaciente.Text = Trim(oRsTmp.Fields!ApellidoPaterno) & " " & Trim(oRsTmp.Fields!ApellidoMaterno) & " " & oRsTmp.Fields!PrimerNombre
          txtNhistoria.Text = oRsTmp.Fields!NroHistoriaClinica
          ml_idFuenteFinanciamiento = oRsTmp.Fields!idFuenteFinanciamiento
          ml_idAtencion = oRsTmp!idAtencion
       Else
          txtDatosDeCuenta.Text = ""
          ml_IdPaciente = 0
          txtNombrePaciente.Text = ""
          txtNhistoria.Text = ""
          ml_idFuenteFinanciamiento = 0
          ml_idAtencion = 0
       End If
       oRsTmp.Close
       Set oRsTmp = Nothing
       oConexion.Close
       Set oConexion = Nothing
       'mariano 07112014
       CargarComboMesesConsumo
    End If

End Sub
Sub CargarComboMesesConsumo()
Dim oRsTmp As New Recordset
       Dim oConexion As New Connection
        oConexion.Open SIGHEntidades.CadenaConexion
        oConexion.CursorLocation = adUseClient
                        Set oRsTmp = oConexion.Execute("SELECT DISTINCT MONTH(dbo.farmMovimiento.fechaCreacion) AS meses,year(dbo.farmMovimiento.fechaCreacion) as ann, dbo.farmMovimientoVentas.idCuentaAtencion FROM dbo.farmMovimiento INNER JOIN dbo.farmMovimientoVentas ON dbo.farmMovimiento.MovNumero = dbo.farmMovimientoVentas.movNumero AND dbo.farmMovimiento.MovTipo = dbo.farmMovimientoVentas.MovTipo Where (dbo.farmMovimientoVentas.idCuentaAtencion = '" & Me.txtNcuenta.Text & "')")
                        CmbMes.Clear
                        Do While Not oRsTmp.EOF
                            CmbMes.AddItem oRsTmp!meses & " " & MonthName(oRsTmp!meses) & " " & (oRsTmp!ann)
                            oRsTmp.MoveNext
                        Loop
                            oRsTmp.Close: Set oRsTmp = Nothing
                            oConexion.Close: Set oConexion = Nothing
End Sub


Sub LimpiarVariablesDeMemoria()
    On Error Resume Next
    Set mo_ReglasFarmacia = Nothing
    Set mo_Teclado = Nothing
End Sub


