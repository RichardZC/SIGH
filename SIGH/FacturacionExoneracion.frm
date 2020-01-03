VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FacturacionExoneracion 
   Caption         =   "Exoneración"
   ClientHeight    =   10230
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   Icon            =   "FacturacionExoneracion.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   10230
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
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
      TabIndex        =   5
      Top             =   30
      Width           =   16290
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
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   630
         Width           =   4455
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
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   630
         Width           =   1395
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
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   255
         Width           =   1740
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
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   240
         Width           =   3135
      End
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
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   240
         Width           =   1410
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
         TabIndex        =   16
         Top             =   675
         Width           =   1305
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
         TabIndex        =   15
         Top             =   285
         Width           =   975
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
         TabIndex        =   14
         Top             =   630
         Width           =   1155
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
         TabIndex        =   13
         Top             =   300
         Width           =   1005
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
         TabIndex        =   12
         Top             =   300
         Width           =   1065
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1035
      Left            =   90
      TabIndex        =   1
      Top             =   9120
      Width           =   16215
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "FacturacionExoneracion.frx":0CCA
         DownPicture     =   "FacturacionExoneracion.frx":118E
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
         Picture         =   "FacturacionExoneracion.frx":167A
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "FacturacionExoneracion.frx":1B66
         DownPicture     =   "FacturacionExoneracion.frx":1FC6
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
         Picture         =   "FacturacionExoneracion.frx":243B
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnImprimir 
         Caption         =   "Imprimir [F3]"
         Enabled         =   0   'False
         Height          =   705
         Left            =   120
         Picture         =   "FacturacionExoneracion.frx":28B0
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   210
         Width           =   1245
      End
   End
   Begin TabDlg.SSTab tabExoneracion 
      Height          =   7935
      Left            =   60
      TabIndex        =   0
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
      TabPicture(0)   =   "FacturacionExoneracion.frx":2D89
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "grdServicios"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Bienes e Insumos"
      TabPicture(1)   =   "FacturacionExoneracion.frx":2DA5
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grdBienes"
      Tab(1).ControlCount=   1
      Begin UltraGrid.SSUltraGrid grdServicios 
         Height          =   7395
         Left            =   120
         TabIndex        =   17
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
         Caption         =   "Lista de Bienes e Insumos"
      End
   End
End
Attribute VB_Name = "FacturacionExoneracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mo_Teclado As New SIGHComun.Teclado
Dim mo_Formulario As New SIGHComun.Formulario

Dim mo_AdminFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_AdminArchivoClinico As New SIGHNegocios.ReglasArchivoClinico

Dim ml_IdUsuario As Long
Dim ms_MensajeError As String
Dim ml_IdCuentaAtencion As Long
Dim mo_cmbIdTipoGenHistoriaClinica As New ListaDespleglable
Dim mrs_FacturacionServicios As ADODB.Recordset
Dim mrs_FacturacionBienes As ADODB.Recordset
Dim mo_Apariencia As New SIGHComun.GridInfragistic

Dim mo_FacturacionServicios As Collection
Dim mo_FacturacionBienes As Collection


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
    If MsgBox("Por favor confirmar, ¿Realmente desea grabar los cambios que ha realizado?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
        Exit Sub
    End If
    CargaDatosAlObjetosDeDatos
    If ValidarReglas() Then
        If ModificarDatos() Then
             MsgBox "Los datos se modificaron correctamente", vbInformation, Me.Caption
             Me.Visible = False
         Else
             MsgBox "No se pudo agregar los datos" + Chr(13) + mo_AdminFacturacion.MensajeError, vbExclamation, Me.Caption
        End If
    End If
End Sub

Private Sub btnCancelar_Click()
    'If MsgBox("Por favor confirmar, ¿Realmente desea salir?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
        Me.Visible = False
    'End If
End Sub

Private Sub Form_Load()

    CargarComboBoxes
    ObtenerDatosDePaciente
    GenerarRecordsetTemporal
    CargarDatosEstadoCuenta
    
    
    mo_Formulario.HabilitarDeshabilitar lblNroCuenta, False
    mo_Formulario.HabilitarDeshabilitar txtIdNroHistoria, False
    mo_Formulario.HabilitarDeshabilitar cmbIdTipoGenHistoriaClinica, False
    mo_Formulario.HabilitarDeshabilitar lblPaciente, False
    mo_Formulario.HabilitarDeshabilitar lblFechaIngreso, False
    mo_Formulario.HabilitarDeshabilitar lblServicioIngreso, False
    
    
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

Private Sub grdBienes_AfterCellUpdate(ByVal Cell As UltraGrid.SSCell)
    If Cell.Column.Key = "SubTotalExonerado" Then
        mrs_FacturacionBienes.Fields!EstadoRegistro = "M"
    End If
End Sub

Private Sub grdBienes_BeforeCellUpdate(ByVal Cell As UltraGrid.SSCell, NewValue As Variant, ByVal Cancel As UltraGrid.SSReturnBoolean)
    If Cell.Column.Key <> "SubTotalExonerado" Then
        Cancel = True
    Else
        'Validamos que el monto ingresado no sea mayor a la deuda que queda
        Dim dPorPagar As Double
        dPorPagar = IIf(IsNull(mrs_FacturacionServicios.Fields!SubTotalPorPagar), 0, mrs_FacturacionServicios.Fields!SubTotalPorPagar)
        If (NewValue - Cell.Value) > dPorPagar Then
            MsgBox "No puede exonerar más de lo que está pendiente de pago", vbExclamation, Me.Caption
            Cancel = True
            Exit Sub
        End If
    End If
End Sub

Private Sub grdServicios_AfterCellUpdate(ByVal Cell As UltraGrid.SSCell)
    If Cell.Column.Key = "SubTotalExonerado" Then
        mrs_FacturacionServicios.Fields!EstadoRegistro = "M"
    End If
End Sub

Private Sub grdServicios_BeforeCellUpdate(ByVal Cell As UltraGrid.SSCell, NewValue As Variant, ByVal Cancel As UltraGrid.SSReturnBoolean)
    If Cell.Column.Key <> "SubTotalExonerado" Then
        Cancel = True
    Else
        'Validamos que el monto ingresado no sea mayor a la deuda que queda
        Dim dPorPagar As Double
        dPorPagar = IIf(IsNull(mrs_FacturacionServicios.Fields!SubTotalPorPagar), 0, mrs_FacturacionServicios.Fields!SubTotalPorPagar)
        If (NewValue - Cell.Value) > dPorPagar Then
            MsgBox "No puede exonerar más de lo que está pendiente de pago", vbExclamation, Me.Caption
            Cancel = True
            Exit Sub
        End If
    End If
End Sub

Private Sub grdServicios_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)

    Dim rs As New Recordset
    
    Set rs = mo_AdminFacturacion.EstadosFacturacionObtenerTodos()
    With grdServicios.ValueLists.Add("IdEstadoFacturacion").ValueListItems
        Do Until rs.EOF
            .Add Trim(Str(rs.Fields!IdEstadoFacturacion)), rs.Fields!Descripcion
            rs.MoveNext
        Loop
    End With
    rs.Close
    
    Set rs = mo_AdminFacturacion.EstadosAtencionObtenerTodos()
    With grdServicios.ValueLists.Add("IdEstadoAtencion").ValueListItems
        Do Until rs.EOF
            .Add Trim(Str(rs.Fields!IdEstadoAtencion)), rs.Fields!Descripcion
            rs.MoveNext
        Loop
    End With
    rs.Close
        
    grdServicios.Bands(0).Columns("IdFacturacionServicio").Hidden = True
    grdServicios.Bands(0).Columns("IdProducto").Hidden = True
    grdServicios.Bands(0).Columns("EstadoRegistro").Hidden = True
    
    grdServicios.Bands(0).Columns("NroOrden").Header.Caption = "Nº Orden"
    grdServicios.Bands(0).Columns("NroOrden").Width = 800
    
    grdServicios.Bands(0).Columns("FechaOrden").Header.Caption = "Fecha Ord."
    grdServicios.Bands(0).Columns("FechaOrden").Width = 1200
    
    grdServicios.Bands(0).Columns("CodProducto").Header.Caption = "Cod.Serv."
    grdServicios.Bands(0).Columns("CodProducto").Width = 1000
   
    grdServicios.Bands(0).Columns("NombreServicio").Header.Caption = "Servicio"
    grdServicios.Bands(0).Columns("NombreServicio").Width = 3000
    
    grdServicios.Bands(0).Columns("Cantidad").Header.Caption = "Cantidad"
    grdServicios.Bands(0).Columns("Cantidad").Width = 800
    
    grdServicios.Bands(0).Columns("PrecioUnitario").Header.Caption = "P.U.(S/.)"
    grdServicios.Bands(0).Columns("PrecioUnitario").Width = 1000
    
    grdServicios.Bands(0).Columns("SubTotalExonerado").Header.Caption = "Exonerar(S/.)"
    grdServicios.Bands(0).Columns("SubTotalExonerado").Width = 1500
    
    grdServicios.Bands(0).Columns("SubTotalPagadoACuenta").Header.Caption = "PagoACuenta(S/.)"
    grdServicios.Bands(0).Columns("SubTotalPagadoACuenta").Width = 1200
    
    grdServicios.Bands(0).Columns("SubTotalPorPagar").Header.Caption = "PorPagar(S/.)"
    grdServicios.Bands(0).Columns("SubTotalPorPagar").Width = 1300
    
    grdServicios.Bands(0).Columns("IdEstadoAtencion").Header.Caption = "EstadoAtención"
    grdServicios.Bands(0).Columns("IdEstadoAtencion").Width = 1600
    grdServicios.Bands(0).Columns("IdEstadoAtencion").ValueList = "IdEstadoAtencion"
    grdServicios.Bands(0).Columns("IdEstadoAtencion").ButtonDisplayStyle = ssButtonDisplayStyleOnCellActivate
    
    grdServicios.Bands(0).Columns("IdEstadoFacturacion").Header.Caption = "EstadoFacturación"
    grdServicios.Bands(0).Columns("IdEstadoFacturacion").Width = 1600
    grdServicios.Bands(0).Columns("IdEstadoFacturacion").ValueList = "IdEstadoFacturacion"
    grdServicios.Bands(0).Columns("IdEstadoFacturacion").ButtonDisplayStyle = ssButtonDisplayStyleOnCellActivate

End Sub
Private Sub grdBienes_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)


    Dim rs As New Recordset
    
    Set rs = mo_AdminFacturacion.EstadosFacturacionObtenerTodos()
    With grdBienes.ValueLists.Add("IdEstadoFacturacion").ValueListItems
        Do Until rs.EOF
            .Add Trim(Str(rs.Fields!IdEstadoFacturacion)), rs.Fields!Descripcion
            rs.MoveNext
        Loop
    End With
    rs.Close
    
    Set rs = mo_AdminFacturacion.EstadosAtencionObtenerTodos()
    With grdBienes.ValueLists.Add("IdEstadoAtencion").ValueListItems
        Do Until rs.EOF
            .Add Trim(Str(rs.Fields!IdEstadoAtencion)), rs.Fields!Descripcion
            rs.MoveNext
        Loop
    End With
    rs.Close
        
    

    grdBienes.Bands(0).Columns("IdFacturacionBienes").Hidden = True
    grdBienes.Bands(0).Columns("IdProducto").Hidden = True
    grdBienes.Bands(0).Columns("EstadoRegistro").Hidden = True
    
    
    grdBienes.Bands(0).Columns("NroReceta").Header.Caption = "Nº Orden"
    grdBienes.Bands(0).Columns("NroReceta").Width = 800
    
    grdBienes.Bands(0).Columns("FechaReceta").Header.Caption = "Fecha Ord."
    grdBienes.Bands(0).Columns("FechaReceta").Width = 1200
    
    grdBienes.Bands(0).Columns("CodProducto").Header.Caption = "Cod.Serv."
    grdBienes.Bands(0).Columns("CodProducto").Width = 1000
   
    grdBienes.Bands(0).Columns("NombreProducto").Header.Caption = "Servicio"
    grdBienes.Bands(0).Columns("NombreProducto").Width = 3000
    
    grdBienes.Bands(0).Columns("Cantidad").Header.Caption = "Cantidad"
    grdBienes.Bands(0).Columns("Cantidad").Width = 800
    
    grdBienes.Bands(0).Columns("PrecioUnitario").Header.Caption = "P.U.(S/.)"
    grdBienes.Bands(0).Columns("PrecioUnitario").Width = 1000
    
    grdBienes.Bands(0).Columns("SubTotalExonerado").Header.Caption = "Exonerar(S/.)"
    grdBienes.Bands(0).Columns("SubTotalExonerado").Width = 1500
    
    grdBienes.Bands(0).Columns("SubTotalPagadoACuenta").Header.Caption = "PagadoACuenta(S/.)"
    grdBienes.Bands(0).Columns("SubTotalPagadoACuenta").Width = 1200
    
    grdBienes.Bands(0).Columns("SubTotalPorPagar").Header.Caption = "PorPagar(S/.)"
    grdBienes.Bands(0).Columns("SubTotalPorPagar").Width = 1300
    
    grdBienes.Bands(0).Columns("IdEstadoAtencion").Header.Caption = "EstadoAtención"
    grdBienes.Bands(0).Columns("IdEstadoAtencion").Width = 1600
    grdBienes.Bands(0).Columns("IdEstadoAtencion").ValueList = "IdEstadoAtencion"
    grdBienes.Bands(0).Columns("IdEstadoAtencion").ButtonDisplayStyle = ssButtonDisplayStyleOnCellActivate
    
    grdBienes.Bands(0).Columns("IdEstadoFacturacion").Header.Caption = "EstadoFacturación"
    grdBienes.Bands(0).Columns("IdEstadoFacturacion").Width = 1600
    grdBienes.Bands(0).Columns("IdEstadoFacturacion").ValueList = "IdEstadoFacturacion"
    grdBienes.Bands(0).Columns("IdEstadoFacturacion").ButtonDisplayStyle = ssButtonDisplayStyleOnCellActivate
    
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

Sub GenerarRecordsetTemporal()
    Set mrs_FacturacionServicios = New Recordset
    With mrs_FacturacionServicios
            '
          .Fields.Append "IdFacturacionServicio", adInteger, 4, adFldIsNullable
          .Fields.Append "IdProducto", adInteger, 4, adFldIsNullable
          .Fields.Append "NroOrden", adVarChar, 20, adFldIsNullable
          .Fields.Append "FechaOrden", adDate, , adFldIsNullable
          .Fields.Append "CodProducto", adVarChar, 20, adFldIsNullable
          .Fields.Append "NombreServicio", adVarChar, 200, adFldIsNullable
          .Fields.Append "Cantidad", adCurrency, 8, adFldIsNullable
          .Fields.Append "PrecioUnitario", adCurrency, 8, adFldIsNullable
          .Fields.Append "SubTotalExonerado", adCurrency, 8, adFldIsNullable
          .Fields.Append "SubTotalPagadoACuenta", adCurrency, 8, adFldIsNullable
          .Fields.Append "SubTotalPorPagar", adCurrency, 8, adFldIsNullable
          .Fields.Append "IdEstadoAtencion", adInteger, 4, adFldIsNullable
          .Fields.Append "IdEstadoFacturacion", adInteger, 4, adFldIsNullable
          .Fields.Append "EstadoRegistro", adVarChar, 1, adFldIsNullable
          
          .LockType = adLockOptimistic
          .Open
    End With
    Set Me.grdServicios.DataSource = mrs_FacturacionServicios
    
    
    Set mrs_FacturacionBienes = New Recordset
    With mrs_FacturacionBienes
            '
          .Fields.Append "IdFacturacionBienes", adInteger, 4, adFldIsNullable
          .Fields.Append "IdProducto", adInteger, 4, adFldIsNullable
          .Fields.Append "NroReceta", adVarChar, 20, adFldIsNullable
          .Fields.Append "FechaReceta", adDate, , adFldIsNullable
          .Fields.Append "CodProducto", adVarChar, 20, adFldIsNullable
          .Fields.Append "NombreProducto", adVarChar, 200, adFldIsNullable
          .Fields.Append "Cantidad", adCurrency, 8, adFldIsNullable
          .Fields.Append "PrecioUnitario", adCurrency, 8, adFldIsNullable
          .Fields.Append "SubTotalExonerado", adCurrency, 8, adFldIsNullable
          .Fields.Append "SubTotalPagadoACuenta", adCurrency, 8, adFldIsNullable
          .Fields.Append "SubTotalPorPagar", adCurrency, 8, adFldIsNullable
          .Fields.Append "IdEstadoAtencion", adInteger, 4, adFldIsNullable
          .Fields.Append "IdEstadoFacturacion", adInteger, 4, adFldIsNullable
          .Fields.Append "EstadoRegistro", adVarChar, 1, adFldIsNullable
          
          .LockType = adLockOptimistic
          .Open
    End With
    Set Me.grdBienes.DataSource = mrs_FacturacionBienes
    
End Sub
Sub CargarDatosEstadoCuenta()
Dim rsDetalle As New Recordset

    Set rsDetalle = mo_AdminFacturacion.FacturacionServiciosObtenerParaEstadoCuenta(ml_IdCuentaAtencion, sghFacturacionServicioTotal)
    Do While Not rsDetalle.EOF
        With mrs_FacturacionServicios
            .AddNew
            
            .Fields!IdFacturacionServicio = rsDetalle!IdFacturacionServicio
            .Fields!IdProducto = rsDetalle!IdProducto
            .Fields!NroOrden = rsDetalle!NroOrden
            .Fields!FechaOrden = rsDetalle!FechaOrden
            .Fields!CodProducto = rsDetalle!CodProducto
            .Fields!NombreServicio = rsDetalle!NombreServicio
            .Fields!Cantidad = rsDetalle!Cantidad
            .Fields!PrecioUnitario = rsDetalle!PrecioUnitario
            .Fields!SubTotalExonerado = rsDetalle!SubTotalExonerado
            .Fields!SubTotalPagadoACuenta = rsDetalle!SubTotalPagadoACuenta
            .Fields!SubTotalPorPagar = rsDetalle!SubTotalPorPagar
            .Fields!IdEstadoAtencion = rsDetalle!IdEstadoAtencion
            .Fields!IdEstadoFacturacion = rsDetalle!IdEstadoFacturacion
            .Fields!EstadoRegistro = "-"
            
        End With
        rsDetalle.MoveNext
    Loop
    mo_Apariencia.ConfigurarFilasBiColores Me.grdServicios, SIGHComun.GrillaConFilasBicolor

    'Cargamos los Bienes e Insumos
    Set rsDetalle = mo_AdminFacturacion.FacturacionBienesInsumosObtenerParaEstadoCuenta(ml_IdCuentaAtencion)
    Do While Not rsDetalle.EOF
        With mrs_FacturacionBienes
            .AddNew
            
            .Fields!IdFacturacionBienes = rsDetalle!IdFacturacionBienes
            .Fields!IdProducto = rsDetalle!IdProducto
            .Fields!NroReceta = rsDetalle!NroReceta
            .Fields!FechaReceta = rsDetalle!FechaReceta
            .Fields!CodProducto = rsDetalle!CodProducto
            .Fields!NombreProducto = rsDetalle!NombreProducto
            .Fields!Cantidad = rsDetalle!Cantidad
            .Fields!PrecioUnitario = rsDetalle!PrecioUnitario
            .Fields!SubTotalExonerado = rsDetalle!SubTotalExonerado
            .Fields!SubTotalPagadoACuenta = rsDetalle!SubTotalPagadoACuenta
            .Fields!SubTotalPorPagar = rsDetalle!SubTotalPorPagar
            .Fields!IdEstadoAtencion = rsDetalle!IdEstadoAtencion
            .Fields!IdEstadoFacturacion = rsDetalle!IdEstadoFacturacion
            .Fields!EstadoRegistro = "-"
            
        End With
        rsDetalle.MoveNext
    Loop
    mo_Apariencia.ConfigurarFilasBiColores Me.grdBienes, SIGHComun.GrillaConFilasBicolor
End Sub
Sub CargaDatosAlObjetosDeDatos()
    Set mo_FacturacionServicios = New Collection
    Set mo_FacturacionBienes = New Collection
    Dim oDOFacturacionServicios As DOFacturacionServicios
    Dim odoFacturacionBienesInsumos As doFacturacionBienesInsumos
        
    If Not (mrs_FacturacionServicios.EOF And mrs_FacturacionServicios.BOF) Then
        mrs_FacturacionServicios.MoveFirst
        Do While Not mrs_FacturacionServicios.EOF
            If mrs_FacturacionServicios!EstadoRegistro = "M" Then
                Set oDOFacturacionServicios = New DOFacturacionServicios
                oDOFacturacionServicios.IdFacturacionServicio = mrs_FacturacionServicios!IdFacturacionServicio
                oDOFacturacionServicios.IdUsuarioAuditoria = Me.IdUsuario
                
                mo_FacturacionServicios.Add oDOFacturacionServicios
            
            End If
            
            mrs_FacturacionServicios.MoveNext
        Loop
        mrs_FacturacionServicios.MoveFirst
    End If
    
    '
    If Not (mrs_FacturacionBienes.EOF And mrs_FacturacionBienes.BOF) Then
        mrs_FacturacionBienes.MoveFirst
        Do While Not mrs_FacturacionBienes.EOF
            If mrs_FacturacionBienes!EstadoRegistro = "M" Then
                Set odoFacturacionBienesInsumos = New doFacturacionBienesInsumos
                odoFacturacionBienesInsumos.IdFacturacionBienes = mrs_FacturacionBienes!IdFacturacionBienes
                'odoFacturacionBienesInsumos.SubTotalExonerado = mrs_FacturacionBienes!SubTotalExonerado
                odoFacturacionBienesInsumos.IdUsuarioAuditoria = Me.IdUsuario
                
                mo_FacturacionBienes.Add odoFacturacionBienesInsumos
            End If
            mrs_FacturacionBienes.MoveNext
        Loop
        mrs_FacturacionBienes.MoveFirst
    End If

End Sub
Function ValidarReglas() As Boolean
    ValidarReglas = False
    
    
    ValidarReglas = True
End Function
Function ModificarDatos() As Boolean
    ModificarDatos = mo_AdminFacturacion.ActualizarExoneracionItemsFacturacion(mo_FacturacionServicios, mo_FacturacionBienes)
End Function

