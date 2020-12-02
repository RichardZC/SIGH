VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form SelecccionProductos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleccion de Productos "
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   12060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1155
      Left            =   90
      TabIndex        =   4
      Top             =   5940
      Width           =   11895
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "frmSelecccionProductos.frx":0000
         DownPicture     =   "frmSelecccionProductos.frx":0460
         Height          =   700
         Left            =   4290
         Picture         =   "frmSelecccionProductos.frx":08D5
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "frmSelecccionProductos.frx":0D4A
         DownPicture     =   "frmSelecccionProductos.frx":120E
         Height          =   700
         Left            =   5820
         Picture         =   "frmSelecccionProductos.frx":16FA
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1365
      End
   End
   Begin TabDlg.SSTab tabCuentas 
      Height          =   5865
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   11925
      _ExtentX        =   21034
      _ExtentY        =   10345
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
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
      TabPicture(0)   =   "frmSelecccionProductos.frx":1BE6
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "grdServicios"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Bienes e Insumos"
      TabPicture(1)   =   "frmSelecccionProductos.frx":1C02
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "grdBienes"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "grillaBusqueda"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin UltraGrid.SSUltraGrid grillaBusqueda 
         Height          =   2655
         Left            =   360
         TabIndex        =   1
         Top             =   2280
         Visible         =   0   'False
         Width           =   10680
         _ExtentX        =   18838
         _ExtentY        =   4683
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   67108884
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "grillaBusqueda"
      End
      Begin UltraGrid.SSUltraGrid grdServicios 
         Height          =   5295
         Left            =   -74880
         TabIndex        =   2
         Top             =   450
         Width           =   11640
         _ExtentX        =   20532
         _ExtentY        =   9340
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   67174420
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ValueLists      =   "frmSelecccionProductos.frx":1C1E
         Caption         =   "Servicios"
      End
      Begin UltraGrid.SSUltraGrid grdBienes 
         Height          =   5265
         Left            =   150
         TabIndex        =   3
         Top             =   450
         Width           =   11610
         _ExtentX        =   20479
         _ExtentY        =   9287
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   67174420
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ValueLists      =   "frmSelecccionProductos.frx":1C89
         Caption         =   "Bienes e Insumos"
      End
   End
End
Attribute VB_Name = "SelecccionProductos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Seleccionar Servicios y/o Medicamentos
'        Programado por: Barrantes D
'        Fecha: Enero 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mrs_FacturacionServicios As New ADODB.Recordset
Dim mrs_FacturacionBienes As New ADODB.Recordset

Dim mb_Acepta As Boolean

Dim gridInfra As New GridInfragistic

Property Let ServiciosDataSource(dValue As Recordset)
   Set mrs_FacturacionServicios = dValue
End Property
Property Get ServiciosDataSource() As Recordset
   Set ServiciosDataSource = mrs_FacturacionServicios
End Property
Property Let BienesDataSource(dValue As Recordset)
   Set mrs_FacturacionBienes = dValue
End Property
Property Get BienesDataSource() As Recordset
   Set BienesDataSource = mrs_FacturacionBienes
End Property
Property Get Acepta() As Boolean
   Acepta = mb_Acepta
End Property

Public Function Inicializar()
    Set grdBienes.DataSource = Nothing
    Set grdServicios.DataSource = Nothing
   ConfiguraGrilla ("Servicios")
   ConfiguraGrilla ("Bienes")
    
End Function

Private Sub ConfiguraGrilla(tipo As String)
    
    Dim rsFacturacion As New ADODB.Recordset
    Dim i As Integer
    With rsFacturacion
        .Fields.Append "Id", adInteger
        .Fields.Append "IdProducto", adInteger
        .Fields.Append "Codigo", adVarChar, 100
        .Fields.Append "Descripcion", adVarChar, 255
        .Fields.Append "Cantidad", adInteger
        .Fields.Append "PrecioUnitario", adDouble
        .Fields.Append "totalporpagar", adDouble
        .Fields.Append "IdEstadoFacturacion", adInteger
        .Fields.Append "IdAtencion", adInteger
        .Fields.Append "Seleccionar", adBoolean
    End With
    
    If tipo = "Bienes" Then
        If Not mrs_FacturacionBienes Is Nothing Then
            If mrs_FacturacionBienes.RecordCount > 0 Then
                mrs_FacturacionBienes.MoveFirst
                rsFacturacion.Open
                For i = 0 To mrs_FacturacionBienes.RecordCount - 1
                    With rsFacturacion
                        .AddNew
                        .Fields("Seleccionar").Value = True
                        .Fields("Id").Value = mrs_FacturacionBienes.Fields("Id").Value
                        .Fields("IdProducto").Value = mrs_FacturacionBienes.Fields("IdProducto").Value
                        .Fields("Codigo").Value = mrs_FacturacionBienes.Fields("Codigo").Value
                        .Fields("Descripcion").Value = mrs_FacturacionBienes.Fields("Descripcion").Value
                        .Fields("Cantidad").Value = mrs_FacturacionBienes.Fields("Cantidad").Value
                        .Fields("PrecioUnitario").Value = mrs_FacturacionBienes.Fields("PrecioUnitario").Value
                        .Fields("totalporpagar").Value = mrs_FacturacionBienes.Fields("totalporpagar").Value
                        .Fields("IdEstadoFacturacion").Value = mrs_FacturacionBienes.Fields("IdEstadoFacturacion").Value
                        .Fields("IdAtencion").Value = mrs_FacturacionBienes.Fields("IdAtencion").Value
                    End With
                    mrs_FacturacionBienes.MoveNext
                Next i
                Set grdBienes.DataSource = rsFacturacion
            End If
        End If
    Else
        If Not mrs_FacturacionServicios Is Nothing Then
            If mrs_FacturacionServicios.RecordCount > 0 Then
                mrs_FacturacionServicios.MoveFirst
                rsFacturacion.Open
                For i = 0 To mrs_FacturacionServicios.RecordCount - 1
                    With rsFacturacion
                        .AddNew
                        .Fields("Seleccionar").Value = True
                        .Fields("Id").Value = Val(mrs_FacturacionServicios.Fields("Id").Value)
                        .Fields("IdProducto").Value = mrs_FacturacionServicios.Fields("IdProducto").Value
                        .Fields("Codigo").Value = mrs_FacturacionServicios.Fields("Codigo").Value
                        .Fields("Descripcion").Value = mrs_FacturacionServicios.Fields("Descripcion").Value
                        .Fields("Cantidad").Value = mrs_FacturacionServicios.Fields("Cantidad").Value
                        .Fields("PrecioUnitario").Value = mrs_FacturacionServicios.Fields("PrecioUnitario").Value
                        .Fields("totalporpagar").Value = mrs_FacturacionServicios.Fields("totalporpagar").Value
                        .Fields("IdEstadoFacturacion").Value = mrs_FacturacionServicios.Fields("IdEstadoFacturacion").Value
                        .Fields("IdAtencion").Value = mrs_FacturacionServicios.Fields("IdAtencion").Value
                    End With
                    mrs_FacturacionServicios.MoveNext
                Next i
                Set grdServicios.DataSource = rsFacturacion
            End If
        End If
    End If
End Sub


Private Sub FormatoGrilla(oGrilla As SSUltraGrid)
Dim oColumnProducto As SSColumn
Dim oColumnId As SSColumn

    If oGrilla.Bands.Count <= 0 Then
         Exit Sub
    End If

    Set oColumnId = oGrilla.Bands(0).Columns("Seleccionar")
    oColumnId.Style = ssStyleCheckBox
    oColumnId.Activation = ssActivationAllowEdit
    
    oGrilla.Bands(0).Columns("IdProducto").Hidden = True
    
    oGrilla.Bands(0).Columns("id").Hidden = True
    
    oGrilla.Bands(0).Columns("Codigo").Header.Caption = "Codigo"
    oGrilla.Bands(0).Columns("Codigo").Activation = ssActivationActivateNoEdit
    oGrilla.Bands(0).Columns("Codigo").Width = 800
    
    Set oColumnProducto = oGrilla.Bands(0).Columns("Descripcion")
    oColumnProducto.Header.Caption = "Descripcion"
    oColumnProducto.Width = 3000
    oColumnProducto.Activation = ssActivationActivateNoEdit
    
    oGrilla.Bands(0).Columns("Cantidad").Header.Caption = "Cantidad"
    oGrilla.Bands(0).Columns("Cantidad").Activation = ssActivationActivateNoEdit
    
    oGrilla.Bands(0).Columns("preciounitario").Header.Caption = "P.U.(s/.)"
    oGrilla.Bands(0).Columns("preciounitario").Format = "#0.00"
    oGrilla.Bands(0).Columns("preciounitario").Activation = ssActivationActivateNoEdit
    
    oGrilla.Bands(0).Columns("totalporpagar").Header.Caption = "Subtotal"
    oGrilla.Bands(0).Columns("totalporpagar").Format = "#0.00"
    oGrilla.Bands(0).Columns("totalporpagar").Activation = ssActivationActivateNoEdit
    
    oGrilla.Bands(0).Columns("Idestadofacturacion").Hidden = True
    
    oGrilla.Bands(0).Columns("IdAtencion").Hidden = True
    
    gridInfra.ConfigurarFilasBiColores oGrilla, sighEntidades.GrillaConFilasBicolor
End Sub



Private Sub cmdAceptar_Click()

    Dim rsFacturacion As ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim i As Integer
    
   
    With rs
        
        .Fields.Append "Id", adInteger
        .Fields.Append "IdProducto", adInteger
        .Fields.Append "Codigo", adVarChar, 100
        .Fields.Append "Descripcion", adVarChar, 255
        .Fields.Append "Cantidad", adInteger
        .Fields.Append "PrecioUnitario", adDouble
        .Fields.Append "totalporpagar", adDouble
        .Fields.Append "IdEstadoFacturacion", adInteger
        .Fields.Append "IdAtencion", adInteger
    End With
    
    Set rsFacturacion = grdBienes.DataSource
    
    If Not rsFacturacion Is Nothing Then
       
        If rsFacturacion.RecordCount > 0 Then
            rsFacturacion.MoveFirst
            rs.Open
        End If
        For i = 0 To rsFacturacion.RecordCount - 1
        With rs
            If rsFacturacion.Fields("Seleccionar").Value = True Then
                .AddNew
                .Fields("Id").Value = rsFacturacion.Fields("Id").Value
                .Fields("IdProducto").Value = rsFacturacion.Fields("IdProducto").Value
                .Fields("Codigo").Value = rsFacturacion.Fields("Codigo").Value
                .Fields("Descripcion").Value = rsFacturacion.Fields("Descripcion").Value
                .Fields("Cantidad").Value = rsFacturacion.Fields("Cantidad").Value
                .Fields("PrecioUnitario").Value = rsFacturacion.Fields("PrecioUnitario").Value
                .Fields("totalporpagar").Value = rsFacturacion.Fields("totalporpagar").Value
                .Fields("IdEstadoFacturacion").Value = rsFacturacion.Fields("IdEstadoFacturacion").Value
                .Fields("IdAtencion").Value = rsFacturacion.Fields("IdAtencion").Value
            End If
            rsFacturacion.MoveNext
        End With
    Next i
    Set mrs_FacturacionBienes.DataSource = rs
    End If
    
    
    
    
    
    Set rsFacturacion = grdServicios.DataSource
    Set rs = New ADODB.Recordset
    With rs
        
        .Fields.Append "Id", adInteger
        .Fields.Append "IdProducto", adInteger
        .Fields.Append "Codigo", adVarChar, 100
        .Fields.Append "Descripcion", adVarChar, 255
        .Fields.Append "Cantidad", adInteger
        .Fields.Append "PrecioUnitario", adDouble
        .Fields.Append "totalporpagar", adDouble
        .Fields.Append "IdEstadoFacturacion", adInteger
        .Fields.Append "IdAtencion", adInteger
    End With
    
    If Not rsFacturacion Is Nothing Then
    If rsFacturacion.RecordCount > 0 Then
        rsFacturacion.MoveFirst
        rs.Open
    End If
    For i = 0 To rsFacturacion.RecordCount - 1
        With rs
            If rsFacturacion.Fields("Seleccionar").Value = True Then
                .AddNew
                .Fields("Id").Value = rsFacturacion.Fields("Id").Value
                .Fields("IdProducto").Value = rsFacturacion.Fields("IdProducto").Value
                .Fields("Codigo").Value = rsFacturacion.Fields("Codigo").Value
                .Fields("Descripcion").Value = rsFacturacion.Fields("Descripcion").Value
                .Fields("Cantidad").Value = rsFacturacion.Fields("Cantidad").Value
                .Fields("PrecioUnitario").Value = rsFacturacion.Fields("PrecioUnitario").Value
                .Fields("totalporpagar").Value = rsFacturacion.Fields("totalporpagar").Value
                .Fields("IdEstadoFacturacion").Value = rsFacturacion.Fields("IdEstadoFacturacion").Value
                .Fields("IdAtencion").Value = rsFacturacion.Fields("IdAtencion").Value
            End If
            rsFacturacion.MoveNext
        End With
    Next i
        Set mrs_FacturacionServicios.DataSource = rs
    End If
    mb_Acepta = True
    'Unload Me
    Me.Hide
End Sub

Private Sub cmdSalir_Click()
    mb_Acepta = False
    Me.Hide
    'Unload Me
End Sub

Private Sub Command1_Click()
    FormatoGrilla grdServicios
End Sub

Private Sub grdBienes_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    FormatoGrilla grdBienes
End Sub

Private Sub grdServicios_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    FormatoGrilla grdServicios
End Sub

