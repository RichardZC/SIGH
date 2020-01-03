VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Begin VB.Form EnvioHistoriaDetalle 
   Caption         =   "Form1"
   ClientHeight    =   8400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13335
   Icon            =   "EnvioHCDetalle.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8400
   ScaleWidth      =   13335
   StartUpPosition =   2  'CenterScreen
   Begin UltraGrid.SSUltraGrid grdHistoriasSolicitadas 
      Height          =   5895
      Left            =   60
      TabIndex        =   29
      Top             =   1260
      Width           =   13185
      _ExtentX        =   23257
      _ExtentY        =   10398
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   67108864
      Caption         =   "Lista de historias"
   End
   Begin VB.Frame frmParametros 
      Caption         =   "Parámetros"
      Height          =   1155
      Left            =   10590
      TabIndex        =   23
      Top             =   60
      Width           =   2655
      Begin VB.CheckBox chkActivarRefrescoAtomatico 
         Caption         =   "Activar refresco automatico"
         Height          =   195
         Left            =   150
         TabIndex        =   28
         Top             =   870
         Width           =   2295
      End
      Begin VB.TextBox txtIntervaloRefresco 
         Height          =   315
         Left            =   2160
         TabIndex        =   26
         Text            =   "10"
         Top             =   540
         Width           =   315
      End
      Begin VB.TextBox txtTamanioLetra 
         Height          =   315
         Left            =   2160
         TabIndex        =   24
         Text            =   "9"
         Top             =   180
         Width           =   315
      End
      Begin VB.Label Label7 
         Caption         =   "Intervalo refresco (seg)"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   570
         Width           =   1725
      End
      Begin VB.Label Label3 
         Caption         =   "Tamaño letra"
         Height          =   255
         Left            =   90
         TabIndex        =   25
         Top             =   270
         Width           =   2115
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1095
      Left            =   9930
      TabIndex        =   19
      Top             =   7230
      Width           =   3345
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "EnvioHCDetalle.frx":08CA
         DownPicture     =   "EnvioHCDetalle.frx":0D8E
         Height          =   700
         Left            =   1785
         Picture         =   "EnvioHCDetalle.frx":127A
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "EnvioHCDetalle.frx":1766
         DownPicture     =   "EnvioHCDetalle.frx":1BC6
         Height          =   700
         Left            =   240
         Picture         =   "EnvioHCDetalle.frx":203B
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   240
         Width           =   1365
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Datos de envío"
      Height          =   1065
      Left            =   90
      TabIndex        =   10
      Top             =   7230
      Width           =   9765
      Begin MSMask.MaskEdBox txtHoraEnvio 
         Height          =   315
         Left            =   1350
         TabIndex        =   22
         Top             =   540
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin VB.CommandButton btnBusquedaRespRecepcion 
         Caption         =   "..."
         Height          =   315
         Left            =   5160
         TabIndex        =   8
         Top             =   570
         Width           =   315
      End
      Begin VB.TextBox txtIdResponsableRecepcion 
         Height          =   315
         Left            =   4080
         TabIndex        =   7
         Top             =   570
         Width           =   1000
      End
      Begin VB.CommandButton btnBusquedaRespEnvio 
         Caption         =   "..."
         Height          =   315
         Left            =   5160
         TabIndex        =   6
         Top             =   210
         Width           =   315
      End
      Begin VB.TextBox txtIdResponsableEnvio 
         Height          =   315
         Left            =   4080
         TabIndex        =   5
         Top             =   210
         Width           =   1000
      End
      Begin MSMask.MaskEdBox txtFechaEnvio 
         Height          =   315
         Left            =   150
         TabIndex        =   4
         Top             =   540
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblNombreRespRecepcion 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   5520
         TabIndex        =   21
         Top             =   570
         Width           =   4095
      End
      Begin VB.Label Label8 
         Caption         =   "Resp. de recepción"
         Height          =   315
         Left            =   2040
         TabIndex        =   20
         Top             =   630
         Width           =   1965
      End
      Begin VB.Label lblNombreRespEnvio 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   5520
         TabIndex        =   18
         Top             =   210
         Width           =   4095
      End
      Begin VB.Label Label6 
         Caption         =   "Resp. de transporte"
         Height          =   315
         Left            =   2040
         TabIndex        =   17
         Top             =   270
         Width           =   1875
      End
      Begin VB.Label Label5 
         Caption         =   "Fecha de envío"
         Height          =   285
         Left            =   150
         TabIndex        =   16
         Top             =   270
         Width           =   1425
      End
   End
   Begin VB.Frame fraBusqueda 
      Caption         =   "Búsqueda"
      Height          =   1155
      Left            =   60
      TabIndex        =   9
      Top             =   60
      Width           =   10485
      Begin VB.ComboBox cmbComparadorFechas 
         Height          =   315
         Left            =   1440
         TabIndex        =   33
         Top             =   660
         Width           =   1410
      End
      Begin VB.ComboBox cmbIdTipoServicio 
         Height          =   315
         Left            =   1440
         TabIndex        =   32
         Top             =   285
         Width           =   3105
      End
      Begin VB.Timer TimerDeRefresco 
         Enabled         =   0   'False
         Interval        =   6000
         Left            =   9930
         Top             =   720
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   5490
         Picture         =   "EnvioHCDetalle.frx":24B0
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   690
         Width           =   585
      End
      Begin MSMask.MaskEdBox txtFechaPrestamoRequeridaHasta 
         Height          =   315
         Left            =   4080
         TabIndex        =   2
         Top             =   660
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFechaPrestamoRequeridaDesde 
         Height          =   315
         Left            =   2910
         TabIndex        =   1
         Top             =   660
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.CommandButton btnBuscarServicios 
         Caption         =   "..."
         Height          =   315
         Left            =   6360
         TabIndex        =   13
         Top             =   300
         Width           =   315
      End
      Begin VB.TextBox txtIdServicio 
         Height          =   315
         Left            =   5490
         TabIndex        =   0
         Top             =   300
         Width           =   825
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha requerida"
         Height          =   315
         Left            =   150
         TabIndex        =   15
         Top             =   750
         Width           =   1515
      End
      Begin VB.Label lblNombreServicio 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   6720
         TabIndex        =   14
         Top             =   300
         Width           =   3615
      End
      Begin VB.Label Label2 
         Caption         =   "Servicio"
         Height          =   285
         Left            =   4650
         TabIndex        =   12
         Top             =   330
         Width           =   1125
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo de servicio"
         Height          =   255
         Left            =   150
         TabIndex        =   11
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "EnvioHistoriaDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: POEnviosHistoriaClinica
'        Autor: William Castro Grijalva
'        Fecha: 18/08/2004 12:26:56 a.m.
'        Empresa: Digital Works Corporation
'        Todos los derechos reservados
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------

Dim mo_Teclado As New SIGHComun.Teclado
Dim mo_Formulario As New SIGHComun.Formulario
Dim ml_IdUsuario As Long
Dim ms_MensajeError As String
Dim mi_Opcion As sghOpciones
Dim mb_ExistenDatos As Boolean
Dim mo_EnviosHistoriaClinica As New DOEnvioHistoriaClinica
Dim ml_IdEnvio As Long
Dim mo_AdminArchivoClinico As New ReglasArchivoClinico
Dim mo_AdminServiciosHosp As New ReglasServiciosHosp
Dim mo_AdminComun As New ReglasComunes
Dim mrs_HistoriasPorEnviar As New ADODB.Recordset
Dim mo_Prestamos As Collection
Dim mo_Apariencia As New SIGHComun.GridInfragistic
Dim mo_cmbIdTipoServicio As New SIGHComun.ListaDespleglable
Dim mo_cmbComparadorFechas As New SIGHComun.ListaDespleglable

Property Let ExistenDatos(bValue As Boolean)
   mb_ExistenDatos = bValue
End Property
Property Get ExistenDatos() As Boolean
   ExistenDatos = mb_ExistenDatos
End Property
Property Let Opcion(iValue As sghOpciones)
   mi_Opcion = iValue
End Property
Property Get Opcion() As sghOpciones
   Opcion = mi_Opcion
End Property
Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property
Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property
Property Let IdUsuario(lValue As Long)
   ml_IdUsuario = lValue
End Property
Property Get IdUsuario() As Long
   IdUsuario = ml_IdUsuario
End Property
Property Let IdEnvio(lValue As Long)
   ml_IdEnvio = lValue
End Property
Property Get IdEnvio() As Long
   IdEnvio = ml_IdEnvio
End Property
Sub CargarComboBoxes()
Dim sSQL As String
Dim sMensaje As String

        mo_cmbIdTipoServicio.BoundColumn = "IdTipoServicio"
        mo_cmbIdTipoServicio.ListField = "DescripcionLarga"
        Set mo_cmbIdTipoServicio.RowSource = mo_AdminServiciosHosp.TiposServicioSeleccionarTodos()
        sMensaje = sMensaje + mo_AdminServiciosHosp.MensajeError

        mo_cmbComparadorFechas.BoundColumn = "IdParametro"
        mo_cmbComparadorFechas.ListField = "Codigo"
        Set mo_cmbComparadorFechas.RowSource = mo_AdminComun.ParametrosComparadorFechas()
        sMensaje = sMensaje + mo_AdminServiciosHosp.MensajeError

        If sMensaje <> "" Then
            MsgBox mo_AdminServiciosHosp.MensajeError, vbCritical, Me.Caption
        End If

End Sub

Private Sub btnAgregar_Click()
Dim sListaAExcluir As String

    Dim oRow As SSRow
    Set oRow = Me.grdHistoriasSolicitadas.ActiveRow

    With mrs_HistoriasPorEnviar
        .AddNew
        .Fields!IdPrestamo = oRow.Cells(0)
        .Fields!HistoriaClinica = oRow.Cells(1)
        .Fields!Nombres = oRow.Cells(2)
        .Fields!FechaPrestamoRequerida = oRow.Cells(3)
        .Fields!NroFolios = 0
    End With
    
    RefrescarHistoriasSolicitadas

End Sub

Private Sub btnBuscar_Click()
    RefrescarHistoriasSolicitadas
    
End Sub
Sub RefrescarHistoriasSolicitadas()
Dim oBusqueda As SIGHComun.sghBusquedaPrestamoHistorias
Dim oPrestamos As Collection
Dim rsPrestamos As New Recordset
Dim oDOPrestamo As New DOPrestamoHistoriaClinica

    'Guarda las historias seleccionadas
    Set oPrestamos = New Collection
    Dim oRow As SSRow
    Dim oPrestamo As DOPrestamoHistoriaClinica
    Set oRow = Me.grdHistoriasSolicitadas.GetRow(ssChildRowFirst)
    If Not oRow Is Nothing Then
        'Para el primero
        If oRow.Cells("Enviar") Then
            Set oPrestamo = New DOPrestamoHistoriaClinica
            oPrestamo.IdPrestamo = "" & oRow.Cells("IdPrestamo").Value
            oPrestamo.NroFolios = Val("" & oRow.Cells("NroFolios").Value)
            oPrestamos.Add oPrestamo
        End If
        'Para los siguientes
        Do While oRow.HasNextSibling
            Set oRow = oRow.GetSibling(ssSiblingRowNext)
            If oRow.Cells("Enviar") Then
                Set oPrestamo = New DOPrestamoHistoriaClinica
                oPrestamo.IdPrestamo = oRow.Cells("IdPrestamo").Value
                oPrestamo.NroFolios = Val("" & oRow.Cells("NroFolios").Value)
                oPrestamos.Add oPrestamo
            End If
        Loop
    End If
    
    oBusqueda.IdServicio = Val(Me.txtIdServicio.Tag)
    oBusqueda.IdTipoServicio = Val(mo_cmbIdTipoServicio.BoundText)
    oBusqueda.ComparadorFecha = mo_cmbComparadorFechas.BoundText
    If IsDate(Me.txtFechaPrestamoRequeridaDesde.Text) Then
        oBusqueda.FechaPrestamoRequeridaDesde = IIf(Me.txtFechaPrestamoRequeridaDesde.Text <> SIGHComun.FECHA_VACIA_DMY, Me.txtFechaPrestamoRequeridaDesde, 0)
    End If
    If IsDate(Me.txtFechaPrestamoRequeridaHasta.Text) Then
        oBusqueda.FechaPrestamoRequeridaHasta = IIf(Me.txtFechaPrestamoRequeridaHasta <> SIGHComun.FECHA_VACIA_DMY, Me.txtFechaPrestamoRequeridaHasta, 0)
    End If
    oBusqueda.IdEnvio = 0
    oBusqueda.IdEstadoPrestamo = 1
    
    'eliminar todas las filas
    On Error Resume Next
    mrs_HistoriasPorEnviar.MoveFirst
    Do While Not mrs_HistoriasPorEnviar.EOF
           mrs_HistoriasPorEnviar.Delete
           mrs_HistoriasPorEnviar.Update
           mrs_HistoriasPorEnviar.MoveNext
    Loop
    
    Set rsPrestamos = mo_AdminArchivoClinico.PrestamosHistoriaClinicaFiltrarParaEnvio(oBusqueda)
    
    Do While Not rsPrestamos.EOF
        With mrs_HistoriasPorEnviar
            .AddNew
            .Fields!IdPrestamo = rsPrestamos!IdPrestamo
            .Fields!HistoriaClinica = rsPrestamos!HistoriaClinica
            .Fields!Nombres = rsPrestamos!Nombres
            .Fields!FechaPrestamoRequerida = Format(rsPrestamos!FechaPrestamoRequerida, "dd/mm/yyyy")
            .Fields!Servicio = rsPrestamos!Servicio
            .Fields!NroFolios = rsPrestamos!NroFolios
            .Fields!Enviar = False
        End With
        rsPrestamos.MoveNext
    Loop
    mo_Apariencia.ConfigurarFilasBiColores grdHistoriasSolicitadas, SIGHComun.GrillaConFilasBicolor
    
    'recupera los valores de las historias
    Set oRow = Me.grdHistoriasSolicitadas.GetRow(ssChildRowFirst)
    If Not oRow Is Nothing Then
        'Para el primero
        If BuscarEnLaColeccion(oRow.Cells("IdPrestamo"), oPrestamos, oDOPrestamo) Then
            oRow.Cells("Enviar") = True
            oRow.Cells("NroFolios") = oDOPrestamo.NroFolios
        End If
        'Para los siguientes
        Do While oRow.HasNextSibling
            Set oRow = oRow.GetSibling(ssSiblingRowNext)
            If BuscarEnLaColeccion(oRow.Cells("IdPrestamo"), oPrestamos, oDOPrestamo) Then
                oRow.Cells("Enviar") = True
                oRow.Cells("NroFolios") = oDOPrestamo.NroFolios
            End If
        Loop
    End If

End Sub
Function BuscarEnLaColeccion(lIdPrestamo As Long, oPrestamos As Collection, oDOPrestamo As DOPrestamoHistoriaClinica) As Boolean
Dim oPrestamo As DOPrestamoHistoriaClinica

    BuscarEnLaColeccion = False
    
    For Each oPrestamo In oPrestamos
        If lIdPrestamo = oPrestamo.IdPrestamo Then
            Set oDOPrestamo = oPrestamo
            BuscarEnLaColeccion = True
            Exit For
        End If
    Next

End Function

Private Sub btnBuscarServicios_Click()
Dim oBusqueda As New ServiciosBusqueda
Dim oDOServicio As New DOServicio

    oBusqueda.IdTipoServicio = Val(mo_cmbIdTipoServicio.BoundText)
    oBusqueda.Show 1
    
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDOServicio = mo_AdminServiciosHosp.ServiciosSeleccionarPorId(oBusqueda.IdRegistroSeleccionado)
        If Not oDOServicio Is Nothing Then
            Me.txtIdServicio.Text = oDOServicio.Codigo
            Me.txtIdServicio.Tag = oDOServicio.IdServicio
            Me.lblNombreServicio = oDOServicio.Nombre
        Else
            Me.txtIdServicio.Text = ""
            Me.txtIdServicio.Tag = ""
            Me.lblNombreServicio = ""
        End If
    End If
    
    
    
End Sub

Private Sub btnBusquedaRespEnvio_Click()
Dim oBusqueda As New EmpleadosBusqueda
Dim oDoEmpleado As New DOEmpleado

    oBusqueda.Show 1
    
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDoEmpleado = mo_AdminComun.EmpleadosSeleccionarPorId(oBusqueda.IdRegistroSeleccionado)
        If Not oDoEmpleado Is Nothing Then
            Me.txtIdResponsableEnvio.Tag = oDoEmpleado.IdEmpleado
            Me.txtIdResponsableEnvio.Text = oDoEmpleado.CodigoPlanilla
            Me.lblNombreRespEnvio = oDoEmpleado.ApellidoPaterno + " " + oDoEmpleado.ApellidoMaterno + " " + oDoEmpleado.Nombres
        End If
    End If
End Sub

Private Sub btnBusquedaRespRecepcion_Click()
Dim oBusqueda As New EmpleadosBusqueda
Dim oDoEmpleado As New DOEmpleado

    oBusqueda.Show 1
    
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDoEmpleado = mo_AdminComun.EmpleadosSeleccionarPorId(oBusqueda.IdRegistroSeleccionado)
        If Not oDoEmpleado Is Nothing Then
            Me.txtIdResponsableRecepcion.Tag = oDoEmpleado.IdEmpleado
            Me.txtIdResponsableRecepcion.Text = oDoEmpleado.CodigoPlanilla
            Me.lblNombreRespRecepcion = oDoEmpleado.ApellidoPaterno + " " + oDoEmpleado.ApellidoMaterno + " " + oDoEmpleado.Nombres
        End If
    End If
End Sub

Private Sub btnQuitar_Click()
    On Error Resume Next
    With mrs_HistoriasPorEnviar
        If Not .EOF And Not .BOF Then
           .Delete
           .Update
        End If
    End With
    RefrescarHistoriasSolicitadas
End Sub

Private Sub chkActivarRefrescoAtomatico_Click()
    SIGHComun.ActivarRefrescoGrillaEnvios = Me.chkActivarRefrescoAtomatico.Value
    Me.TimerDeRefresco.Enabled = Val(Me.chkActivarRefrescoAtomatico.Value)
End Sub

Private Sub cmbComparadorFechas_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbComparadorFechas
    AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbComparadorFechas_LostFocus()
   'If cmbComparadorFechas.Text <> "" Then
   '    mo_cmbComparadorFechas.BoundText = Trim(cmbComparadorFechas.Text)
   'End If
End Sub

Private Sub cmbComparadorFechas_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsComparador(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub cmbIdTipoServicio_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdTipoServicio
    AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbIdTipoServicio_LostFocus()
   If cmbIdTipoServicio.Text <> "" Then
       mo_cmbIdTipoServicio.BoundText = Val(Split(cmbIdTipoServicio.Text, " = ")(0))
   End If
End Sub

Private Sub cmbIdTipoServicio_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub



Private Sub Form_Initialize()
    
    Set mo_cmbIdTipoServicio.MiComboBox = cmbIdTipoServicio
    Set mo_cmbComparadorFechas.MiComboBox = cmbComparadorFechas
    
End Sub

Private Sub grdHistoriasSolicitadas_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    
    Me.grdHistoriasSolicitadas.Bands(0).Columns("IdPrestamo").Hidden = True
    
    Me.grdHistoriasSolicitadas.Bands(0).Columns("HistoriaClinica").Header.Caption = "Nº HC"
    Me.grdHistoriasSolicitadas.Bands(0).Columns("HistoriaClinica").Width = 1000
    
    Me.grdHistoriasSolicitadas.Bands(0).Columns("Nombres").Header.Caption = "Nombres y apellidos"
    Me.grdHistoriasSolicitadas.Bands(0).Columns("Nombres").Width = 4000
    
    Me.grdHistoriasSolicitadas.Bands(0).Columns("FechaPrestamoRequerida").Header.Caption = "Fec. Req."
    Me.grdHistoriasSolicitadas.Bands(0).Columns("FechaPrestamoRequerida").Width = 2000
    
    Me.grdHistoriasSolicitadas.Bands(0).Columns("Servicio").Header.Caption = "Servicio"
    Me.grdHistoriasSolicitadas.Bands(0).Columns("Servicio").Width = 4000
    
    Me.grdHistoriasSolicitadas.Bands(0).Columns("NroFolios").Header.Caption = "Nº Folios"
    Me.grdHistoriasSolicitadas.Bands(0).Columns("NroFolios").Width = 1200
    
    Me.grdHistoriasSolicitadas.Bands(0).Columns("Enviar").Header.Caption = "Enviar?"
    Me.grdHistoriasSolicitadas.Bands(0).Columns("Enviar").Style = ssStyleCheckBox
    Me.grdHistoriasSolicitadas.Bands(0).Columns("Enviar").Width = 800
    
'    Dim Col As SSColumn
'    Set Col = grdHistoriasSolicitadas.Bands(0).Columns.Add("Enviar", "Enviar?")
'    Col.Style = ssStyleCheckBox
'    Col.DataType = ssDataTypeBoolean
'    Col.Width = 800
    
    Select Case mi_Opcion
    Case sghModificar, sghEliminar, sghConsultar
            'Guarda las historias seleccionadas
            Dim oRow As SSRow
            Set oRow = Me.grdHistoriasSolicitadas.GetRow(ssChildRowFirst)
            If Not oRow Is Nothing Then
                'Para el primero
                oRow.Cells("Enviar") = True
                'Para los siguientes
                Do While oRow.HasNextSibling
                    Set oRow = oRow.GetSibling(ssSiblingRowNext)
                    oRow.Cells("Enviar") = True
                Loop
            End If
    End Select

End Sub

Private Sub TimerDeRefresco_Timer()
    RefrescarHistoriasSolicitadas
End Sub

Private Sub txtIdResponsableRecepcion_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtIdResponsableRecepcion
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtIdResponsableRecepcion_LostFocus()
    
    If txtIdResponsableRecepcion.Text <> "" Then
        Dim oDoEmpleado As New DOEmpleado
        If mo_AdminComun.EmpleadosSeleccionarPorCodigo(txtIdResponsableRecepcion.Text, oDoEmpleado) Then
            txtIdResponsableRecepcion.Tag = oDoEmpleado.IdEmpleado
            txtIdResponsableRecepcion.Text = oDoEmpleado.CodigoPlanilla
            Me.lblNombreRespRecepcion = oDoEmpleado.ApellidoPaterno + " " + oDoEmpleado.ApellidoMaterno + " " + oDoEmpleado.Nombres
        Else
            txtIdResponsableRecepcion.Tag = ""
            txtIdResponsableRecepcion = ""
            Me.lblNombreRespRecepcion = ""
        End If
    End If
   
   mo_Formulario.MarcarComoVacio txtIdResponsableRecepcion
End Sub

Private Sub txtIdResponsableRecepcion_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtIdResponsableEnvio_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtIdResponsableEnvio
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtIdResponsableEnvio_LostFocus()
    
    If txtIdResponsableEnvio <> "" Then
        Dim oDoEmpleado As New DOEmpleado
        If mo_AdminComun.EmpleadosSeleccionarPorCodigo(txtIdResponsableEnvio.Text, oDoEmpleado) Then
            txtIdResponsableEnvio.Tag = oDoEmpleado.IdEmpleado
            txtIdResponsableEnvio.Text = oDoEmpleado.CodigoPlanilla
            Me.lblNombreRespEnvio = oDoEmpleado.ApellidoPaterno + " " + oDoEmpleado.ApellidoMaterno + " " + oDoEmpleado.Nombres
        Else
            txtIdResponsableEnvio.Tag = ""
            txtIdResponsableEnvio = ""
            Me.lblNombreRespEnvio = ""
        End If
    End If
    
   mo_Formulario.MarcarComoVacio txtIdResponsableEnvio
End Sub

Private Sub txtIdResponsableEnvio_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtFechaEnvio_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtFechaEnvio
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtFechaEnvio_LostFocus()
   mo_Formulario.MarcarComoVacio txtFechaEnvio
End Sub

Private Sub txtFechaEnvio_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsParaFecha(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub
Private Sub txtHoraEnvio_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtHoraEnvio
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtHoraEnvio_LostFocus()
   mo_Formulario.MarcarComoVacio txtHoraEnvio
End Sub

Private Sub txtHoraEnvio_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsParaHora(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla EnviosHistoriaClinica
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosAlFormulario()

    mo_Formulario.HabilitarDeshabilitar Me.lblNombreServicio, False
    mo_Formulario.HabilitarDeshabilitar Me.lblNombreRespEnvio, False
    mo_Formulario.HabilitarDeshabilitar Me.lblNombreRespRecepcion, False
    
    Select Case mi_Opcion
     Case sghAgregar
        Me.txtFechaEnvio = Format(Now, "dd/mm/yyyy")
        Me.txtHoraEnvio = Format(Now, "hh:mm")
        
     Case sghModificar
         CargarDatosALosControles
     Case sghConsultar
         CargarDatosALosControles
     Case sghEliminar
         CargarDatosALosControles
 End Select
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla EnviosHistoriaClinica
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Load()
        
        Me.chkActivarRefrescoAtomatico.Value = 0
        Me.TimerDeRefresco.Enabled = False
        Me.frmParametros.Enabled = False
        Me.fraBusqueda.Enabled = False
        GenerarRecordsetTemporal
        
        Select Case mi_Opcion
        Case sghAgregar
            Me.Caption = "Agregar EnviosHistoriaClinica"
            Me.chkActivarRefrescoAtomatico.Value = 0
            Me.frmParametros.Enabled = True
            Me.fraBusqueda.Enabled = True
            
            Me.txtIntervaloRefresco = SIGHComun.IntervaloRefrescoGrillaEnvios
            Me.chkActivarRefrescoAtomatico.Value = Val(SIGHComun.ActivarRefrescoGrillaEnvios)
        
        Case sghModificar
            Me.Caption = "Modificar EnviosHistoriaClinica"
        Case sghConsultar
            Me.Caption = "Consultar EnviosHistoriaClinica"
        Case sghEliminar
            Me.Caption = "Eliminar EnviosHistoriaClinica"
        End Select

        CargarComboBoxes
        CargarDatosAlFormulario
        mo_Formulario.ConfigurarTipoLetra "Tahoma", "9", Me
        
        Me.grdHistoriasSolicitadas.Font.Name = "Tahoma"
        Me.txtTamanioLetra = SIGHComun.TamanioLetraDeGrillaEnvios
        Me.grdHistoriasSolicitadas.Font.Size = SIGHComun.TamanioLetraDeGrillaEnvios
                
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla EnviosHistoriaClinica
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Activate()
   If mi_Opcion <> sghAgregar Then
       If Not mb_ExistenDatos Then
           Me.Visible = False
       End If
   End If
End Sub
Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
       Case vbKeyEscape
           btnCancelar_Click
       Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   AdministrarKeyPreview KeyCode
End Sub

Private Sub btnAceptar_Click()
   
    
   Select Case mi_Opcion
   Case sghAgregar
       If ValidarDatosObligatorios() Then
            CargaDatosAlObjetosDeDatos
            If ValidarReglas() Then
               If AgregarDatos() Then
                    MsgBox "Los datos se agregaron correctamente", vbInformation, Me.Caption
                    LimpiarFormulario
                Else
                    MsgBox "No se pudo agregar los datos" + Chr(13) + mo_AdminArchivoClinico.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghModificar
       If ValidarDatosObligatorios() Then
            CargaDatosAlObjetosDeDatos
            If ValidarReglas() Then
               If ModificarDatos() Then
                    MsgBox "Los datos se modificaron correctamente", vbInformation, Me.Caption
                    Me.Visible = False
                Else
                    MsgBox "No se pudo modificar los datos" + Chr(13) + mo_AdminArchivoClinico.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghEliminar
            CargaDatosAlObjetosDeDatos
            If ValidarReglas() Then
               If EliminarDatos() Then
                    MsgBox "Los datos se eliminaron correctamente", vbInformation, Me.Caption
                    Me.Visible = False
                Else
                    MsgBox "No se pudo eliminar los datos" + Chr(13) + mo_AdminArchivoClinico.MensajeError, vbExclamation, Me.Caption
               End If
           End If
   End Select
End Sub

Private Sub btnCancelar_Click()
    
    Me.TimerDeRefresco.Enabled = False
   Me.Visible = False
End Sub

Function ValidarDatosObligatorios() As Boolean
Dim sMensaje As String

   ValidarDatosObligatorios = False
   If Me.txtIdResponsableRecepcion.Text = "" Then
       sMensaje = sMensaje + "Ingrese el responsable de la recepción de las historias clínicas" + Chr(13)
   End If
   If Me.txtIdResponsableEnvio.Text = "" Then
       sMensaje = sMensaje + "Ingrese el responsable del envío (transporte) de las historias clinicas" + Chr(13)
   End If
   If Me.txtFechaEnvio.Text = "" Then
       sMensaje = sMensaje + "Ingrese la fecha y hora del envio" + Chr(13)
   End If
    
   If sMensaje <> "" Then
       MsgBox sMensaje, vbInformation, Me.Caption
       Exit Function
   End If
   ValidarDatosObligatorios = True
End Function
Function ValidarReglas() As Boolean
   ValidarReglas = False
    
    If mo_Prestamos.Count = 0 Then
        MsgBox "Debe seleccionar las historias clinicas a enviar", vbExclamation, Me.Caption
        Exit Function
    End If
    
   ValidarReglas = True
End Function
'------------------------------------------------------------------------------------
'   Cargar datos al objetos de datos
'   Descripción:    Seleccionar un registro unico de la tabla EnviosHistoriaClinica
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargaDatosAlObjetosDeDatos()
Dim oPrestamos As New Collection

   With mo_EnviosHistoriaClinica
           .IdResponsableRecepcion = Me.txtIdResponsableRecepcion.Tag
           .IdResponsableEnvio = Me.txtIdResponsableEnvio.Tag
           .FechaPrestamoReal = Me.txtFechaEnvio.Text
           .HoraPrestamoReal = Me.txtHoraEnvio.Text
           .IdEnvio = Me.IdEnvio
   End With
        
    'Guarda las historias seleccionadas
    Set oPrestamos = New Collection
    Dim oRow As SSRow
    Dim oPrestamo As DOPrestamoHistoriaClinica
    Set oRow = Me.grdHistoriasSolicitadas.GetRow(ssChildRowFirst)
    If Not oRow Is Nothing Then
        'Para el primero
        If oRow.Cells("Enviar") Then
            Set oPrestamo = New DOPrestamoHistoriaClinica
            oPrestamo.IdPrestamo = oRow.Cells("IdPrestamo").Value
            oPrestamo.NroFolios = IIf(IsNull(oRow.Cells("NroFolios").Value), 0, oRow.Cells("NroFolios").Value)
            oPrestamos.Add oPrestamo
        End If
        'Para los siguientes
        Do While oRow.HasNextSibling
            Set oRow = oRow.GetSibling(ssSiblingRowNext)
            If oRow.Cells("Enviar") Then
                Set oPrestamo = New DOPrestamoHistoriaClinica
                oPrestamo.IdPrestamo = oRow.Cells("IdPrestamo").Value
                oPrestamo.NroFolios = IIf(IsNull(oRow.Cells("NroFolios").Value), 0, oRow.Cells("NroFolios").Value)
                oPrestamos.Add oPrestamo
            End If
        Loop
    End If
   
    Set mo_Prestamos = oPrestamos
   
End Sub

'------------------------------------------------------------------------------------
'        Agregar Datos
'------------------------------------------------------------------------------------

Function AgregarDatos() As Boolean
    AgregarDatos = mo_AdminArchivoClinico.EnviosHistoriaClinicaAgregar(mo_EnviosHistoriaClinica, mo_Prestamos)
End Function

'------------------------------------------------------------------------------------
'        Modificar Datos
'------------------------------------------------------------------------------------

Function ModificarDatos() As Boolean
    ModificarDatos = mo_AdminArchivoClinico.EnviosHistoriaClinicaModificar(mo_EnviosHistoriaClinica, mo_Prestamos)
End Function

'------------------------------------------------------------------------------------
'        Eliminar Datos
'------------------------------------------------------------------------------------
Function EliminarDatos() As Boolean
    EliminarDatos = mo_AdminArchivoClinico.EnviosHistoriaClinicaEliminar(mo_EnviosHistoriaClinica)
End Function

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla EnviosHistoriaClinica
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosALosControles()
Dim oDOPrestamo As New DOPrestamoHistoriaClinica
Dim oDOPaciente As New doPaciente
Dim rsPrestamos As Recordset
Dim oDoEmpleado As New DOEmpleado

        Set mo_EnviosHistoriaClinica = mo_AdminArchivoClinico.EnvioHistoriaClinicaSeleccionarPorId(Me.IdEnvio)
        If mo_AdminArchivoClinico.MensajeError <> "" Then
             MsgBox "No se pudo obtener los datos + Chr(13) + mo_AdminServiciosComunes.MensajeError, vbCritical, Me.Caption"
             mb_ExistenDatos = False
             Exit Sub
        End If
        
        If Not mo_EnviosHistoriaClinica Is Nothing Then
            With mo_EnviosHistoriaClinica
                
                Me.txtIdResponsableRecepcion.Tag = .IdResponsableRecepcion
                Me.txtIdResponsableEnvio.Tag = .IdResponsableEnvio
                
                Me.txtFechaEnvio.Text = IIf(.FechaPrestamoReal <> 0, Format(.FechaPrestamoReal, "dd/mm/yyyy"), SIGHComun.FECHA_VACIA_DMY)
                Me.txtHoraEnvio.Text = IIf(.HoraPrestamoReal <> "", Format(.HoraPrestamoReal, "hh:mm"), "__:__")
                Me.IdEnvio = .IdEnvio
                
                 Set oDoEmpleado = mo_AdminComun.EmpleadosSeleccionarPorId(.IdResponsableEnvio)
                 If Not oDoEmpleado Is Nothing Then
                     Me.txtIdResponsableEnvio = oDoEmpleado.CodigoPlanilla
                     Me.lblNombreRespEnvio = oDoEmpleado.ApellidoPaterno + " " + oDoEmpleado.ApellidoMaterno + " " + oDoEmpleado.Nombres
                 End If
                
                 Set oDoEmpleado = mo_AdminComun.EmpleadosSeleccionarPorId(.IdResponsableRecepcion)
                 If Not oDoEmpleado Is Nothing Then
                     Me.txtIdResponsableRecepcion = oDoEmpleado.CodigoPlanilla
                     Me.lblNombreRespRecepcion = oDoEmpleado.ApellidoPaterno + " " + oDoEmpleado.ApellidoMaterno + " " + oDoEmpleado.Nombres
                 End If
                
                mb_ExistenDatos = True
            End With
            
            Dim oBusqueda As SIGHComun.sghBusquedaPrestamoHistorias
            oBusqueda.IdEnvio = Me.IdEnvio
            oBusqueda.IdEstadoPrestamo = 2
            Set rsPrestamos = mo_AdminArchivoClinico.PrestamosHistoriaClinicaFiltrarParaEnvio(oBusqueda)
            Do While Not rsPrestamos.EOF
                With mrs_HistoriasPorEnviar
                    .AddNew
                    .Fields!IdPrestamo = rsPrestamos!IdPrestamo
                    .Fields!HistoriaClinica = rsPrestamos!HistoriaClinica
                    .Fields!Nombres = rsPrestamos!Nombres
                    .Fields!FechaPrestamoRequerida = Format(rsPrestamos!FechaPrestamoRequerida, "dd/mm/yyyy")
                    .Fields!Servicio = rsPrestamos!Servicio
                    .Fields!NroFolios = rsPrestamos!NroFolios
                End With
                rsPrestamos.MoveNext
            Loop
            mo_Apariencia.ConfigurarFilasBiColores grdHistoriasSolicitadas, SIGHComun.GrillaConFilasBicolor
       Else
           mb_ExistenDatos = False
           Exit Sub
       End If
   
End Sub

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla EnviosHistoriaClinica
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub LimpiarFormulario()

           Me.txtIdResponsableRecepcion.Text = ""
           Me.txtIdResponsableEnvio.Text = ""
           Me.txtFechaEnvio.Text = SIGHComun.FECHA_VACIA_DMY
           Me.txtHoraEnvio.Text = SIGHComun.HORA_VACIA_HM
           Me.lblNombreRespEnvio = ""
           Me.lblNombreRespRecepcion = ""
           Me.IdEnvio = 0
   
End Sub

Private Sub txtIdServicio_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtIdServicio
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtIdServicio_LostFocus()
    
    Me.txtIdServicio.Text = UCase(Me.txtIdServicio.Text)

   If Me.txtIdServicio.Text <> "" Then
    Dim oDOServicio As DOServicio
        Set oDOServicio = mo_AdminServiciosHosp.ServiciosSeleccionarPorCodigo(Me.txtIdServicio.Text)
        If Not oDOServicio Is Nothing Then
            Me.txtIdServicio.Tag = oDOServicio.IdServicio
            Me.lblNombreServicio = oDOServicio.Nombre
            mo_cmbIdTipoServicio.BoundText = oDOServicio.IdTipoServicio
        Else
            Me.txtIdServicio.Tag = ""
            Me.lblNombreServicio = ""
            mo_cmbIdTipoServicio.BoundText = ""
        End If
    Else
            Me.txtIdServicio.Tag = ""
            Me.lblNombreServicio = ""
            mo_cmbIdTipoServicio.BoundText = ""
   End If
   
End Sub

Private Sub txtIdServicio_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtIntervaloRefresco_LostFocus()
    If Val(txtIntervaloRefresco) < 1 Then
        MsgBox "El mínimo tamaño permitido es 1 seg", vbExclamation, Me.Caption
        txtIntervaloRefresco = 1
    End If
    If Val(txtIntervaloRefresco) > 300 Then
        MsgBox "El máximo tamaño permitido es 300 seg", vbExclamation, Me.Caption
        txtIntervaloRefresco = 5
    End If
    
    SIGHComun.IntervaloRefrescoGrillaEnvios = Me.txtIntervaloRefresco
    Me.TimerDeRefresco.Interval = Val(Me.txtIntervaloRefresco) * 1000
End Sub
Private Sub txtIntervaloRefresco_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtIntervaloRefresco
End Sub

Private Sub txtIntervaloRefresco_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub
Private Sub txtTamanioLetra_lostfocus()
    If Val(txtTamanioLetra) < 8 Then
        MsgBox "El mínimo tamaño permitido es 8", vbExclamation, Me.Caption
        txtTamanioLetra = 8
    End If
    If Val(txtTamanioLetra) > 20 Then
        MsgBox "El máximo tamaño permitido es 20", vbExclamation, Me.Caption
        txtTamanioLetra = 20
    End If
    
    SIGHComun.TamanioLetraDeGrillaEnvios = txtTamanioLetra
    Me.grdHistoriasSolicitadas.Font.Size = txtTamanioLetra
End Sub
Private Sub txtTamanioLetra_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtTamanioLetra
End Sub

Private Sub txtTamanioLetra_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Sub GenerarRecordsetTemporal()
    
    With mrs_HistoriasPorEnviar
          .Fields.Append "IdPrestamo", adInteger
          .Fields.Append "HistoriaClinica", adInteger
          .Fields.Append "Nombres", adVarChar, 255
          .Fields.Append "FechaPrestamoRequerida", adChar, 10
          .Fields.Append "Servicio", adVarChar, 100
          .Fields.Append "NroFolios", adInteger, 4, adFldIsNullable
          .Fields.Append "Enviar", adBoolean
          .CursorType = adOpenKeyset
          .LockType = adLockOptimistic
          .Open
    End With
    Set Me.grdHistoriasSolicitadas.DataSource = mrs_HistoriasPorEnviar
    
End Sub


