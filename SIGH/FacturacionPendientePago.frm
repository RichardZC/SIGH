VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FacturacionPendientePago 
   Caption         =   "Pendientes de Pago"
   ClientHeight    =   10230
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   Icon            =   "FacturacionPendientePago.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10230
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame fraDatos 
      Caption         =   "Datos del paciente"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   60
      TabIndex        =   6
      Top             =   30
      Width           =   16290
      Begin VB.TextBox txtIdNroHistoria 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3915
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   240
         Width           =   1410
      End
      Begin VB.ComboBox cmbIdTipoGenHistoriaClinica 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5370
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   240
         Width           =   3135
      End
      Begin VB.TextBox lblNroCuenta 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1125
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   255
         Width           =   1740
      End
      Begin VB.TextBox lblPaciente 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9390
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   240
         Width           =   4440
      End
      Begin VB.TextBox lblFechaIngreso 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3930
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   630
         Width           =   1395
      End
      Begin VB.TextBox lblServicioIngreso 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9390
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   630
         Width           =   4455
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Cuenta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   17
         Top             =   300
         Width           =   1065
      End
      Begin VB.Label Label2 
         Caption         =   "Paciente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8580
         TabIndex        =   16
         Top             =   300
         Width           =   1005
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Ingreso"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2670
         TabIndex        =   15
         Top             =   630
         Width           =   1155
      End
      Begin VB.Label Label7 
         Caption         =   "Nº historia"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3000
         TabIndex        =   14
         Top             =   285
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Servicio Ingreso"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8025
         TabIndex        =   13
         Top             =   675
         Width           =   1305
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1035
      Left            =   90
      TabIndex        =   0
      Top             =   9120
      Width           =   16275
      Begin VB.CommandButton btnImprimir 
         Caption         =   "Imprimir [F3]"
         Enabled         =   0   'False
         Height          =   705
         Left            =   120
         Picture         =   "FacturacionPendientePago.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   210
         Width           =   1245
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "FacturacionPendientePago.frx":11A3
         DownPicture     =   "FacturacionPendientePago.frx":1603
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   6645
         Picture         =   "FacturacionPendientePago.frx":1A78
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "FacturacionPendientePago.frx":1EED
         DownPicture     =   "FacturacionPendientePago.frx":23B1
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   8190
         Picture         =   "FacturacionPendientePago.frx":289D
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   225
         Width           =   1365
      End
   End
   Begin TabDlg.SSTab tabExoneracion 
      Height          =   7935
      Left            =   60
      TabIndex        =   4
      Top             =   1140
      Width           =   16275
      _ExtentX        =   28707
      _ExtentY        =   13996
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Servicios"
      TabPicture(0)   =   "FacturacionPendientePago.frx":2D89
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "grdServicios"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Bienes e Insumos"
      TabPicture(1)   =   "FacturacionPendientePago.frx":2DA5
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grdBienes"
      Tab(1).ControlCount=   1
      Begin UltraGrid.SSUltraGrid grdServicios 
         Height          =   7395
         Left            =   120
         TabIndex        =   5
         Top             =   420
         Width           =   16035
         _ExtentX        =   28284
         _ExtentY        =   13044
         _Version        =   131072
         GridFlags       =   17040388
         UpdateMode      =   2
         LayoutFlags     =   67108884
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Lista de Servicios"
      End
      Begin UltraGrid.SSUltraGrid grdBienes 
         Height          =   7395
         Left            =   -74880
         TabIndex        =   18
         Top             =   420
         Width           =   16035
         _ExtentX        =   28284
         _ExtentY        =   13044
         _Version        =   131072
         GridFlags       =   17040388
         UpdateMode      =   2
         LayoutFlags     =   67108884
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Lista de Bienes"
      End
   End
End
Attribute VB_Name = "FacturacionPendientePago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: POAtencionesInterconsultas
'        Autor: William Castro Grijalva
'        Fecha: 31/10/2004 09:32:29 a.m.
'        Empresa: Digital Works Corporation
'        Todos los derechos reservados
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------
Dim mo_Teclado As New SIGHCOmun.Teclado
Dim mo_Formulario As New SIGHCOmun.Formulario

Dim mo_AdminFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_AdminArchivoClinico As New SIGHNegocios.ReglasArchivoClinico

Dim ml_IdUsuario As Long
Dim ms_MensajeError As String
Dim ml_IdCuentaAtencion As Long
Dim mo_cmbIdTipoGenHistoriaClinica As New ListaDespleglable

Property Let IdCuentaAtencion(Value As Long)
    ml_IdCuentaAtencion = Value
End Property
Property Get IdCuentaAtencion() As Long
    IdCuentaAtencion = ml_IdCuentaAtencion
End Property
Property Let IdUsuario(Value As Long)
    ml_IdUsuario = Value
End Property
Property Get IdUsuario() As Long
    IdUsuario = ml_IdUsuario
End Property

Private Sub btnAceptar_Click()
    
    If MsgBox("Por favor confirmar, ¿Realmente desea grabar los cambios que ha realizado?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
        Me.grdServicios.Update
        Me.grdBienes.Update
        
        Set Me.grdServicios.DataSource = mo_AdminFacturacion.FacturacionServiciosObtenerParaPendientePago(ml_IdCuentaAtencion)
        Set Me.grdBienes.DataSource = mo_AdminFacturacion.FacturacionBienesInsumosObtenerParaPendientePago(ml_IdCuentaAtencion)
    End If
    
End Sub

Private Sub btnCancelar_Click()
    If MsgBox("Por favor confirmar, ¿Realmente desea salir?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
        Me.Visible = False
    End If
End Sub

Private Sub Form_Load()

    CargarComboBoxes
    ObtenerDatosDePaciente

    'Cargar datos de servicios
    Set Me.grdServicios.DataSource = mo_AdminFacturacion.FacturacionServiciosObtenerParaPendientePago(ml_IdCuentaAtencion)
    Set Me.grdBienes.DataSource = mo_AdminFacturacion.FacturacionBienesInsumosObtenerParaPendientePago(ml_IdCuentaAtencion)
    
    ConfigurarListaDesplegablesDeServicio
    ConfigurarListaDesplegablesDeBienes
    
    mo_Formulario.HabilitarDeshabilitar lblNroCuenta, False
    mo_Formulario.HabilitarDeshabilitar txtIdNroHistoria, False
    mo_Formulario.HabilitarDeshabilitar cmbIdTipoGenHistoriaClinica, False
    mo_Formulario.HabilitarDeshabilitar lblPaciente, False
    mo_Formulario.HabilitarDeshabilitar lblFechaIngreso, False
    mo_Formulario.HabilitarDeshabilitar lblServicioIngreso, False
    
    
End Sub
Private Sub ConfigurarListaDesplegablesDeServicio()
Dim oValueList As SSValueList
Dim i As Long
Dim rsEstado As ADODB.Recordset

    'Crea lista de estadps de facturacion
    Set oValueList = Me.grdServicios.ValueLists.Add("EstadoFacturacion")
    Set rsEstado = mo_AdminFacturacion.EstadosFacturacionObtenerTodos()
    rsEstado.MoveFirst
    For i = 0 To rsEstado.RecordCount - 1
        oValueList.ValueListItems.Add rsEstado.Fields("IdEstadoFacturacion").Value, rsEstado.Fields("Descripcion").Value
        rsEstado.MoveNext
    Next i
    Me.grdServicios.Bands(1).Columns("IdEstadoFacturacion").ValueList = oValueList
    Me.grdServicios.Bands(1).Columns("IdEstadoFacturacion").ValueList.DisplayStyle = ssValueListDisplayStyleDisplayText
    Me.grdServicios.Bands(1).Columns("IdEstadoFacturacion").Style = ssStyleDropDownList
    rsEstado.Close

    'Crea lista de empleados que han autorizado pendiente incluyendo el nuevo empleado
    Set oValueList = Me.grdServicios.ValueLists.Add("Empleados")
    Set rsEstado = mo_AdminFacturacion.EmpleadosSeleccionarParaPendientePagoServicio(ml_IdCuentaAtencion, ml_IdUsuario)
    rsEstado.MoveFirst
    For i = 0 To rsEstado.RecordCount - 1
        oValueList.ValueListItems.Add rsEstado.Fields("IdEmpleado").Value, rsEstado.Fields("Nombre").Value
        rsEstado.MoveNext
    Next i
    Me.grdServicios.Bands(1).Columns("IdEmpleadoAutorizaPendiente").ValueList = oValueList
    Me.grdServicios.Bands(1).Columns("IdEmpleadoAutorizaPendiente").ValueList.DisplayStyle = ssValueListDisplayStyleDisplayText
    Me.grdServicios.Bands(1).Columns("IdEmpleadoAutorizaPendiente").Style = ssStyleDropDownList
    rsEstado.Close

End Sub
Private Sub ConfigurarListaDesplegablesDeBienes()
Dim oValueList As SSValueList
Dim i As Long
Dim rsEstado As ADODB.Recordset

    'Crea lista de estadps de facturacion
    Set oValueList = Me.grdBienes.ValueLists.Add("EstadoFacturacion")
    Set rsEstado = mo_AdminFacturacion.EstadosFacturacionObtenerTodos()
    rsEstado.MoveFirst
    For i = 0 To rsEstado.RecordCount - 1
        oValueList.ValueListItems.Add rsEstado.Fields("IdEstadoFacturacion").Value, rsEstado.Fields("Descripcion").Value
        rsEstado.MoveNext
    Next i
    Me.grdBienes.Bands(1).Columns("IdEstadoFacturacion").ValueList = oValueList
    Me.grdBienes.Bands(1).Columns("IdEstadoFacturacion").ValueList.DisplayStyle = ssValueListDisplayStyleDisplayText
    Me.grdBienes.Bands(1).Columns("IdEstadoFacturacion").Style = ssStyleDropDownList
    rsEstado.Close

    'Crea lista de empleados que han autorizado pendiente incluyendo el nuevo empleado
    Set oValueList = Me.grdBienes.ValueLists.Add("Empleados")
    Set rsEstado = mo_AdminFacturacion.EmpleadosSeleccionarParaPendientePagoBienInsumo(ml_IdCuentaAtencion, ml_IdUsuario)
    rsEstado.MoveFirst
    For i = 0 To rsEstado.RecordCount - 1
        oValueList.ValueListItems.Add rsEstado.Fields("IdEmpleado").Value, rsEstado.Fields("Nombre").Value
        rsEstado.MoveNext
    Next i
    Me.grdBienes.Bands(1).Columns("IdEmpleadoAutorizaPendiente").ValueList = oValueList
    Me.grdBienes.Bands(1).Columns("IdEmpleadoAutorizaPendiente").ValueList.DisplayStyle = ssValueListDisplayStyleDisplayText
    Me.grdBienes.Bands(1).Columns("IdEmpleadoAutorizaPendiente").Style = ssStyleDropDownList
    rsEstado.Close

End Sub

Private Sub Form_Resize()

    On Error Resume Next
    Me.tabExoneracion.Width = Me.Width - 240
    Me.tabExoneracion.Height = Me.Height - Me.Frame4.Height - Me.fraDatos.Height - 640
    
    Me.grdServicios.Width = Me.tabExoneracion.Width - 240
    Me.grdServicios.Height = Me.tabExoneracion.Height - 560
    
    Me.grdBienes.Width = Me.tabExoneracion.Width - 240
    Me.grdBienes.Height = Me.tabExoneracion.Height - 560
    
    Me.fraDatos.Width = Me.tabExoneracion.Width
    
    Me.Frame4.Width = Me.tabExoneracion.Width
    Me.Frame4.Left = Me.tabExoneracion.Left
    Me.Frame4.Top = Me.tabExoneracion.Top + Me.tabExoneracion.Height
End Sub

Function CalculaTotalPendientePorCategoria(oRowParent As SSRow)
Dim oRow As SSRow

    Dim cTotal As Currency
    Set oRow = oRowParent.GetChild(ssChildRowFirst)
    cTotal = oRow.Cells("SubTotalPendientePago").Value
    Do While oRow.HasNextSibling
        Set oRow = oRow.GetSibling(ssSiblingRowNext)
        cTotal = cTotal + oRow.Cells("SubTotalPendientePago").Value
    Loop
    CalculaTotalPendientePorCategoria = cTotal
End Function


Private Sub grdServicios_AfterCellListCloseUp(ByVal Cell As UltraGrid.SSCell)
Dim oRow As SSRow
Dim oRowParent As SSRow

    If Cell.Column.BaseColumnName = "IdEstadoFacturacion" Then
    
        Set oRow = Cell.Row
        Select Case Cell.GetText
        Case "Emitido"
            oRow.Cells("SubTotalPendientePago").Value = 0
            oRow.Cells("IdEmpleadoAutorizaPendiente").Value = Null
            oRow.Cells("FechaAutorizaPendiente").Value = Null
        Case "Pendiente Pago"
            oRow.Cells("SubTotalPendientePago").Value = oRow.Cells("SubTotalPorPagar").Value
            oRow.Cells("IdEmpleadoAutorizaPendiente").Value = ml_IdUsuario
            oRow.Cells("FechaAutorizaPendiente").Value = Format(Now, "dd/MM/yyyy hh:mm:ss")
        Case Else
            MsgBox "Ud solo puede modificar el estado si esta en Emitido o Pendiente de Pago", vbInformation, Me.Caption
            Me.grdServicios.PerformAction ssKeyActionUndoCell
            Exit Sub
        End Select
    
        oRow.Cells("IdEmpleadoModifica").Value = ml_IdUsuario
        oRow.Cells("FechaModificacion").Value = Format(Now, "dd/MM/yyyy hh:mm:ss")
    
        oRow.Cells("SubTotalPendientePago").Refresh
        oRow.Cells("IdEmpleadoAutorizaPendiente").Refresh
        oRow.Cells("FechaAutorizaPendiente").Refresh
        Me.grdServicios.PerformAction ssKeyActionExitEditMode
        
        Dim oFirstRow As SSRow
        Set oRowParent = oRow.GetParent()
        Set oFirstRow = oRowParent.GetSibling(ssSiblingRowFirst)
        Set oRow = oFirstRow
        Dim cTotal As Currency
        Dim cTotalCategoria As Currency
        cTotal = 0
        cTotalCategoria = 0
        
        Do While oRow.HasNextSibling
            Set oRow = oRow.GetSibling(ssSiblingRowNext)
            cTotalCategoria = CalculaTotalPendientePorCategoria(oRow)
            cTotal = cTotal + cTotalCategoria
            
            oRow.Cells("SubTotalPendientePagoAux").Value = cTotalCategoria
            oRow.Cells("SubTotalPendientePagoAux").Refresh
        Loop
        
        oFirstRow.Cells("SubTotalPendientePagoAux").Value = cTotal
        oFirstRow.Cells("SubTotalPendientePagoAux").Refresh
    End If

End Sub

Private Sub grdBienes_AfterCellListCloseUp(ByVal Cell As UltraGrid.SSCell)
Dim oRow As SSRow
Dim oRowParent As SSRow

    If Cell.Column.BaseColumnName = "IdEstadoFacturacion" Then
    
        Set oRow = Cell.Row
        Select Case Cell.GetText
        Case "Emitido"
            oRow.Cells("SubTotalPendientePago").Value = 0
            oRow.Cells("IdEmpleadoAutorizaPendiente").Value = Null
            oRow.Cells("FechaAutorizaPendiente").Value = Null
        Case "Pendiente Pago"
            oRow.Cells("SubTotalPendientePago").Value = oRow.Cells("SubTotalPorPagar").Value
            oRow.Cells("IdEmpleadoAutorizaPendiente").Value = ml_IdUsuario
            oRow.Cells("FechaAutorizaPendiente").Value = Format(Now, "dd/MM/yyyy hh:mm:ss")
        Case Else
            MsgBox "Ud solo puede modificar el estado si esta en Emitido o Pendiente de Pago", vbInformation, Me.Caption
            Me.grdBienes.PerformAction ssKeyActionUndoCell
            Exit Sub
        End Select
    
        oRow.Cells("IdEmpleadoModifica").Value = ml_IdUsuario
        oRow.Cells("FechaModificacion").Value = Format(Now, "dd/MM/yyyy hh:mm:ss")
    
        oRow.Cells("SubTotalPendientePago").Refresh
        oRow.Cells("IdEmpleadoAutorizaPendiente").Refresh
        oRow.Cells("FechaAutorizaPendiente").Refresh
        Me.grdBienes.PerformAction ssKeyActionExitEditMode
        
        Dim oFirstRow As SSRow
        Set oRowParent = oRow.GetParent()
        Set oFirstRow = oRowParent.GetSibling(ssSiblingRowFirst)
        Set oRow = oFirstRow
        Dim cTotal As Currency
        Dim cTotalCategoria As Currency
        cTotal = 0
        cTotalCategoria = 0
        
        Do While oRow.HasNextSibling
            Set oRow = oRow.GetSibling(ssSiblingRowNext)
            cTotalCategoria = CalculaTotalPendientePorCategoria(oRow)
            cTotal = cTotal + cTotalCategoria
            
            oRow.Cells("SubTotalPendientePagoAux").Value = cTotalCategoria
            oRow.Cells("SubTotalPendientePagoAux").Refresh
        Loop
        
        oFirstRow.Cells("SubTotalPendientePagoAux").Value = cTotal
        oFirstRow.Cells("SubTotalPendientePagoAux").Refresh
    End If

End Sub

Private Sub grdServicios_BeforeCellListDropDown(ByVal Cell As UltraGrid.SSCell, ByVal Cancel As UltraGrid.SSReturnBoolean)
Dim oRow As SSRow

    Set oRow = Cell.Row
    Select Case oRow.Cells("IdEstadoFacturacion").Value
    Case 1, 3
    Case Else
        MsgBox "Ud solo puede modificar el estado si esta en Emitido o Pendiente de Pago ", vbInformation, Me.Caption
        Cancel = True
    End Select


End Sub
Private Sub grdBienes_BeforeCellListDropDown(ByVal Cell As UltraGrid.SSCell, ByVal Cancel As UltraGrid.SSReturnBoolean)
Dim oRow As SSRow

    Set oRow = Cell.Row
    Select Case oRow.Cells("IdEstadoFacturacion").Value
    Case 1, 3
    Case Else
        MsgBox "Ud solo puede modificar el estado si esta en Emitido o Pendiente de Pago ", vbInformation, Me.Caption
        Cancel = True
    End Select


End Sub

Private Sub grdServicios_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)

    Layout.ViewStyleBand = ssViewStyleBandVertical
    Layout.Override.ExpandRowsOnLoad = ssExpandOnLoadNo
    Layout.Override.FetchRows = ssFetchRowsPreloadWithParent
    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti


    With Layout.Override
        .ExpandRowsOnLoad = ssExpandOnLoadNo
        .CellClickAction = ssClickActionEdit
        '.RowSelectors = ssRowSelectorsOff
        .CellSpacing = 75
        .CellPadding = 45
        .RowAppearance.BackColor = &H44F4F9 '&HCDEBFF
        .CellAppearance.BackColor = vbWhite
        .BorderStyleCell = ssBorderStyleNone
        .BorderStyleRow = ssBorderStyleNone
        
        .RowAppearance.AlphaLevel = 192
        .RowAppearance.BackColorAlpha = ssAlphaUseAlphaLevel
        .CellAppearance.AlphaLevel = 192
        .CellAppearance.BackColorAlpha = ssAlphaUseAlphaLevel
        
        .ActiveRowAppearance.BackColorAlpha = ssAlphaOpaque
        .ActiveCellAppearance.BackColorAlpha = ssAlphaOpaque
        
    End With
    
    InitializePendientePago
    
End Sub

Sub InitializePendientePago()
    
    'Banda 0
    Me.grdServicios.Bands(0).Override.HeaderAppearance.Font.Name = "Tahoma"
    Me.grdServicios.Bands(0).Override.HeaderAppearance.Font.Size = 10
    Me.grdServicios.Bands(0).Override.HeaderAppearance.Font.Bold = True
    
    Me.grdServicios.Bands(0).Override.RowAppearance.Font.Name = "Tahoma"
    Me.grdServicios.Bands(0).Override.RowAppearance.Font.Size = 10
    Me.grdServicios.Bands(0).Override.RowAppearance.BackColor = &HDEB59E
    
    Me.grdServicios.Bands(0).Columns("IdCategoriaProducto").Hidden = True
    Me.grdServicios.Bands(0).Columns("Descripcion").Width = 5000
    
    Me.grdServicios.Bands(0).Columns("SubTotalPorPagar").Header.Caption = "Por Pagar S/."
    Me.grdServicios.Bands(0).Columns("SubTotalPorPagar").Width = 1500
    Me.grdServicios.Bands(0).Columns("SubTotalPorPagar").Activation = ssActivationActivateOnly
    
    Me.grdServicios.Bands(0).Columns("SubTotalPendientePago").Hidden = True
    
    Me.grdServicios.Bands(0).Columns.Add "SubTotalPendientePagoAux"
    Me.grdServicios.Bands(0).Columns("SubTotalPendientePagoAux").Header.Caption = "Pendiente S/."
    Me.grdServicios.Bands(0).Columns("SubTotalPendientePagoAux").Width = 1500
    Me.grdServicios.Bands(0).Columns("SubTotalPendientePagoAux").CellAppearance.TextAlign = ssAlignRight
    Me.grdServicios.Bands(0).Columns("SubTotalPendientePagoAux").Activation = ssActivationActivateOnly
    
    'Banda 1
    Me.grdServicios.Bands(1).Override.HeaderAppearance.Font.Name = "Tahoma"
    Me.grdServicios.Bands(1).Override.HeaderAppearance.Font.Size = 10
    Me.grdServicios.Bands(1).Override.HeaderAppearance.Font.Bold = True
    
    Me.grdServicios.Bands(1).Override.RowAppearance.Font.Name = "Tahoma"
    Me.grdServicios.Bands(1).Override.RowAppearance.Font.Size = 10

    Me.grdServicios.Bands(1).Columns("IdCategoriaProducto").Hidden = True
    Me.grdServicios.Bands(1).Columns("Nombre").Width = 5000
    Me.grdServicios.Bands(1).Columns("Nombre").DisplayEllipses = ssDisplayEllipsesYes
    Me.grdServicios.Bands(1).Columns("Nombre").Activation = ssActivationActivateOnly
    
    Me.grdServicios.Bands(1).Columns("PrecioUnitario").Header.Caption = "P.U. S/."
    Me.grdServicios.Bands(1).Columns("PrecioUnitario").Width = 1000
    Me.grdServicios.Bands(1).Columns("PrecioUnitario").Activation = ssActivationActivateNoEdit
    
    Me.grdServicios.Bands(1).Columns("Cantidad").Header.Caption = "Cantidad S/."
    Me.grdServicios.Bands(1).Columns("Cantidad").Width = 1000
    Me.grdServicios.Bands(1).Columns("Cantidad").Activation = ssActivationActivateNoEdit
    

    Me.grdServicios.Bands(1).Columns("SubTotalPorPagar").Header.Caption = "Por Pagar S/."
    Me.grdServicios.Bands(1).Columns("SubTotalPorPagar").Width = 1500
    Me.grdServicios.Bands(1).Columns("SubTotalPorPagar").Activation = ssActivationActivateOnly
    
    Me.grdServicios.Bands(1).Columns("SubTotalPendientePago").Header.Caption = "Pendiente S/."
    Me.grdServicios.Bands(1).Columns("SubTotalPendientePago").Width = 1500
    Me.grdServicios.Bands(1).Columns("SubTotalPendientePago").Activation = ssActivationActivateOnly
    
    Me.grdServicios.Bands(1).Columns("IdEstadoFacturacion").Header.Caption = "Estado"
    Me.grdServicios.Bands(1).Columns("IdEstadoFacturacion").Width = 1800
        
    Me.grdServicios.Bands(1).Columns("IdEmpleadoAutorizaPendiente").Header.Caption = "Resp.de Pend."
    Me.grdServicios.Bands(1).Columns("IdEmpleadoAutorizaPendiente").Width = 3000
    Me.grdServicios.Bands(1).Columns("IdEmpleadoAutorizaPendiente").Activation = ssActivationActivateOnly
    
    Me.grdServicios.Bands(1).Columns("FechaAutorizaPendiente").Header.Caption = "Fecha Pendiente"
    Me.grdServicios.Bands(1).Columns("FechaAutorizaPendiente").Width = 2500
    Me.grdServicios.Bands(1).Columns("FechaAutorizaPendiente").Activation = ssActivationActivateOnly
    Me.grdServicios.Bands(1).Columns("FechaAutorizaPendiente").Format = "dd/MM/yyyy hh:mm:ss"

    Me.grdServicios.Bands(1).Columns("IdEmpleadoModifica").Hidden = True
    Me.grdServicios.Bands(1).Columns("FechaModificacion").Hidden = True

End Sub

Private Sub grdBienes_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)

    Layout.ViewStyleBand = ssViewStyleBandVertical
    Layout.Override.ExpandRowsOnLoad = ssExpandOnLoadNo
    Layout.Override.FetchRows = ssFetchRowsPreloadWithParent
    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti


    With Layout.Override
        .ExpandRowsOnLoad = ssExpandOnLoadNo
        .CellClickAction = ssClickActionEdit
        '.RowSelectors = ssRowSelectorsOff
        .CellSpacing = 75
        .CellPadding = 45
        .RowAppearance.BackColor = &H44F4F9 '&HCDEBFF
        .CellAppearance.BackColor = vbWhite
        .BorderStyleCell = ssBorderStyleNone
        .BorderStyleRow = ssBorderStyleNone
        
        .RowAppearance.AlphaLevel = 192
        .RowAppearance.BackColorAlpha = ssAlphaUseAlphaLevel
        .CellAppearance.AlphaLevel = 192
        .CellAppearance.BackColorAlpha = ssAlphaUseAlphaLevel
        
        .ActiveRowAppearance.BackColorAlpha = ssAlphaOpaque
        .ActiveCellAppearance.BackColorAlpha = ssAlphaOpaque
        
    End With
    
    InitializePendientePagoBienes
    
End Sub

Sub InitializePendientePagoBienes()
    
    'Banda 0
    Me.grdBienes.Bands(0).Override.HeaderAppearance.Font.Name = "Tahoma"
    Me.grdBienes.Bands(0).Override.HeaderAppearance.Font.Size = 10
    Me.grdBienes.Bands(0).Override.HeaderAppearance.Font.Bold = True
    
    Me.grdBienes.Bands(0).Override.RowAppearance.Font.Name = "Tahoma"
    Me.grdBienes.Bands(0).Override.RowAppearance.Font.Size = 10
    Me.grdBienes.Bands(0).Override.RowAppearance.BackColor = &HDEB59E
    
    Me.grdBienes.Bands(0).Columns("IdTipoBienInsumo").Hidden = True
    Me.grdBienes.Bands(0).Columns("Descripcion").Width = 5000
    
    Me.grdBienes.Bands(0).Columns("SubTotalPorPagar").Header.Caption = "Por Pagar S/."
    Me.grdBienes.Bands(0).Columns("SubTotalPorPagar").Width = 1500
    Me.grdBienes.Bands(0).Columns("SubTotalPorPagar").Activation = ssActivationActivateOnly
    
    Me.grdBienes.Bands(0).Columns("SubTotalPendientePago").Hidden = True
    
    Me.grdBienes.Bands(0).Columns.Add "SubTotalPendientePagoAux"
    Me.grdBienes.Bands(0).Columns("SubTotalPendientePagoAux").Header.Caption = "Pendiente S/."
    Me.grdBienes.Bands(0).Columns("SubTotalPendientePagoAux").Width = 1500
    Me.grdBienes.Bands(0).Columns("SubTotalPendientePagoAux").CellAppearance.TextAlign = ssAlignRight
    Me.grdBienes.Bands(0).Columns("SubTotalPendientePagoAux").Activation = ssActivationActivateOnly
    
    'Banda 1
    Me.grdBienes.Bands(1).Override.HeaderAppearance.Font.Name = "Tahoma"
    Me.grdBienes.Bands(1).Override.HeaderAppearance.Font.Size = 10
    Me.grdBienes.Bands(1).Override.HeaderAppearance.Font.Bold = True
    
    Me.grdBienes.Bands(1).Override.RowAppearance.Font.Name = "Tahoma"
    Me.grdBienes.Bands(1).Override.RowAppearance.Font.Size = 10

    Me.grdBienes.Bands(1).Columns("IdTipoBienInsumo").Hidden = True
    Me.grdBienes.Bands(1).Columns("Nombre").Width = 5000
    Me.grdBienes.Bands(1).Columns("Nombre").DisplayEllipses = ssDisplayEllipsesYes
    Me.grdBienes.Bands(1).Columns("Nombre").Activation = ssActivationActivateOnly
    
    Me.grdBienes.Bands(1).Columns("PrecioUnitario").Header.Caption = "P.U. S/."
    Me.grdBienes.Bands(1).Columns("PrecioUnitario").Width = 1000
    Me.grdBienes.Bands(1).Columns("PrecioUnitario").Activation = ssActivationActivateNoEdit
    
    Me.grdBienes.Bands(1).Columns("Cantidad").Header.Caption = "Cantidad S/."
    Me.grdBienes.Bands(1).Columns("Cantidad").Width = 1000
    Me.grdBienes.Bands(1).Columns("Cantidad").Activation = ssActivationActivateNoEdit
    

    Me.grdBienes.Bands(1).Columns("SubTotalPorPagar").Header.Caption = "Por Pagar S/."
    Me.grdBienes.Bands(1).Columns("SubTotalPorPagar").Width = 1500
    Me.grdBienes.Bands(1).Columns("SubTotalPorPagar").Activation = ssActivationActivateOnly
    
    Me.grdBienes.Bands(1).Columns("SubTotalPendientePago").Header.Caption = "Pendiente S/."
    Me.grdBienes.Bands(1).Columns("SubTotalPendientePago").Width = 1500
    Me.grdBienes.Bands(1).Columns("SubTotalPendientePago").Activation = ssActivationActivateOnly
    
    Me.grdBienes.Bands(1).Columns("IdEstadoFacturacion").Header.Caption = "Estado"
    Me.grdBienes.Bands(1).Columns("IdEstadoFacturacion").Width = 1800
        
    Me.grdBienes.Bands(1).Columns("IdEmpleadoAutorizaPendiente").Header.Caption = "Resp.de Pend."
    Me.grdBienes.Bands(1).Columns("IdEmpleadoAutorizaPendiente").Width = 3000
    Me.grdBienes.Bands(1).Columns("IdEmpleadoAutorizaPendiente").Activation = ssActivationActivateOnly
    
    Me.grdBienes.Bands(1).Columns("FechaAutorizaPendiente").Header.Caption = "Fecha Pendiente"
    Me.grdBienes.Bands(1).Columns("FechaAutorizaPendiente").Width = 2500
    Me.grdBienes.Bands(1).Columns("FechaAutorizaPendiente").Activation = ssActivationActivateOnly
    Me.grdBienes.Bands(1).Columns("FechaAutorizaPendiente").Format = "dd/MM/yyyy hh:mm:ss"

    Me.grdBienes.Bands(1).Columns("IdEmpleadoModifica").Hidden = True
    Me.grdBienes.Bands(1).Columns("FechaModificacion").Hidden = True

End Sub


Private Sub grdServicios_InitializeRow(ByVal Context As UltraGrid.Constants_Context, ByVal Row As UltraGrid.SSRow, ByVal ReInitialize As Boolean)

    If Row.Band.Index = 0 Then
        Row.Cells("SubTotalPendientePagoAux").Value = Row.Cells("SubTotalPendientePago").Value
    Else
        Row.Cells("SubTotalPendientePago").Appearance.Font.Size = 11
        Row.Cells("SubTotalPendientePago").Appearance.Font.Bold = True
        Row.Cells("SubTotalPendientePago").Appearance.ForeColor = RGB(255, 0, 0)
        
        If Row.Cells("IdEstadoFacturacion").Value <> 1 And Row.Cells("IdEstadoFacturacion").Value <> 2 Then
            Row.Cells("IdEstadoFacturacion").Activation = ssActivationDisabled
        End If
        
    End If
    
    
    
End Sub
Private Sub grdBienes_InitializeRow(ByVal Context As UltraGrid.Constants_Context, ByVal Row As UltraGrid.SSRow, ByVal ReInitialize As Boolean)

    If Row.Band.Index = 0 Then
        Row.Cells("SubTotalPendientePagoAux").Value = Row.Cells("SubTotalPendientePago").Value
    Else
        Row.Cells("SubTotalPendientePago").Appearance.Font.Size = 11
        Row.Cells("SubTotalPendientePago").Appearance.Font.Bold = True
        Row.Cells("SubTotalPendientePago").Appearance.ForeColor = RGB(255, 0, 0)
        
        If Row.Cells("IdEstadoFacturacion").Value <> 1 And Row.Cells("IdEstadoFacturacion").Value <> 2 Then
            Row.Cells("IdEstadoFacturacion").Activation = ssActivationDisabled
        End If
        
    End If
    
    
    
End Sub


Sub ObtenerDatosDePaciente()
Dim rsPaciente  As New Recordset

    Screen.MousePointer = vbHourglass
    Set rsPaciente = mo_AdminAdmision.CuentasAtencionDatosPacientePorIdCuentaAtencion(ml_IdCuentaAtencion)
    Screen.MousePointer = vbDefault
    
    'Si hay una sola coincidencia
    If rsPaciente.RecordCount = 1 Then
        rsPaciente.MoveFirst
        Me.txtIdNroHistoria.Text = rsPaciente!NroHistoriaClinica
        mo_cmbIdTipoGenHistoriaClinica.BoundText = rsPaciente!IdTipoNumeracion
        Me.lblFechaIngreso = rsPaciente!FechaIngreso
        Me.lblServicioIngreso = rsPaciente!ServicioIngreso
        Me.lblPaciente = rsPaciente!ApellidoPaterno + " " + rsPaciente!ApellidoMaterno + " " + rsPaciente!PrimerNombre + " " + ("" & rsPaciente!SegundoNombre)
        Me.lblNroCuenta = rsPaciente!IdCuentaAtencion
    ElseIf rsPaciente.RecordCount = 0 Then
        MsgBox "No se encontraron atenciones para el nro de cuenta ingresado", vbInformation, Me.Caption
    End If

End Sub

Private Sub Form_Initialize()
    Set mo_cmbIdTipoGenHistoriaClinica.MiComboBox = cmbIdTipoGenHistoriaClinica
End Sub

Sub CargarComboBoxes()
Dim sSQL As String
Dim sMensaje As String
       
       mo_cmbIdTipoGenHistoriaClinica.BoundColumn = "IdTipoNumeracion"
       mo_cmbIdTipoGenHistoriaClinica.ListField = "DescripcionLarga"
       Set mo_cmbIdTipoGenHistoriaClinica.RowSource = mo_AdminArchivoClinico.TiposGeneracionHistoriasSeleccionarTodos()

End Sub

