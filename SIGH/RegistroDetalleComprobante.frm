VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Begin VB.Form RegistroDetalleComprobante 
   Caption         =   "Agregar Detalle Comprobante Pago"
   ClientHeight    =   5835
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7320
   Icon            =   "RegistroDetalleComprobante.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5835
   ScaleWidth      =   7320
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "&Cerrar [ESC]"
      Height          =   375
      Left            =   5790
      TabIndex        =   10
      Top             =   5310
      Width           =   1395
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "&Agregar [F2]"
      Height          =   375
      Left            =   4290
      TabIndex        =   9
      Top             =   5310
      Width           =   1395
   End
   Begin VB.Frame Frame2 
      Caption         =   "Items encontrados"
      Height          =   4635
      Left            =   60
      TabIndex        =   8
      Top             =   1140
      Width           =   7215
      Begin VB.TextBox txtCantidad 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   900
         TabIndex        =   6
         Text            =   "1"
         Top             =   4140
         Width           =   1005
      End
      Begin UltraGrid.SSUltraGrid grdItems 
         Height          =   3885
         Left            =   60
         TabIndex        =   4
         Top             =   180
         Width           =   7065
         _ExtentX        =   12462
         _ExtentY        =   6853
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
         Caption         =   "Items Comprobante"
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "&Cantidad"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   4200
         Width           =   630
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Busqueda"
      Height          =   1035
      Left            =   60
      TabIndex        =   7
      Top             =   60
      Width           =   7215
      Begin VB.OptionButton optBienesInsumos 
         Caption         =   "&Bienes Insumos      [F11]"
         Height          =   255
         Left            =   3000
         TabIndex        =   1
         Top             =   240
         Width           =   2655
      End
      Begin VB.OptionButton optServicios 
         Caption         =   "&Servicios     [F10]"
         Height          =   255
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.TextBox txtNombre 
         Height          =   315
         Left            =   1020
         TabIndex        =   3
         Top             =   540
         Width           =   4875
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Nombre"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   555
      End
   End
End
Attribute VB_Name = "RegistroDetalleComprobante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'MZD Ini 06/06/2005 [Todo el archivo]
Option Explicit
Dim mrs_ItemsFacturables As New ADODB.Recordset
Dim mo_Teclado As New SIGHComun.Teclado
Dim mo_Formulario As New SIGHComun.Formulario
Dim ms_MensajeError As String
Dim mo_AdminSeguridad As New SIGHNegocios.ReglasDeSeguridad
Dim mo_AdminCaja As New SIGHNegocios.ReglasCaja
Dim mo_AdminServiciosComunes As New SIGHNegocios.ReglasComunes

Dim mo_Apariencia As New SIGHComun.GridInfragistic

Sub GenerarRecordsetTemporal()

    InitRecordSetItems
    
'    With mrs_ItemsFacturables
'          .Fields.Append "Tipo", adVarChar, 2, adFldIsNullable
'          .Fields.Append "IdProducto", adVarChar, 2, adFldIsNullable
'          .Fields.Append "Codigo", adVarChar, 20, adFldIsNullable
'          .Fields.Append "Nombre", adVarChar, 250, adFldIsNullable
'          .Fields.Append "PrecioUnitario", adCurrency, 8, adFldIsNullable
'
'          .LockType = adLockOptimistic
'          .Open
'    End With
'
'    Set Me.grdItems.DataSource = mrs_ItemsFacturables
End Sub

Private Sub cmdAgregar_Click()
    Dim CompDetalleSeleccionado As New DOCajaComprobantesDetalle
    'Validamos que haya ingresado una cantidad
    If Val(txtCantidad) <= 0 Then
        MsgBox "Debe ingresar una cantidad mayor a cero", vbExclamation, Me.Caption
        txtCantidad.SetFocus
        Exit Sub
    End If
    On Error Resume Next
    Dim rs As ADODB.Recordset
    Set rs = Me.grdItems.DataSource
    
    With rs
        If Not .EOF And Not .BOF Then
            CompDetalleSeleccionado.TipoDetalle = .Fields!tipo
            CompDetalleSeleccionado.CodigoProducto = .Fields!codigo
            CompDetalleSeleccionado.IdProducto = .Fields!IdProducto
            CompDetalleSeleccionado.PrecioUnitario = .Fields!PrecioUnitario
            CompDetalleSeleccionado.cantidad = Val(Me.txtCantidad)
            CompDetalleSeleccionado.NombreProducto = .Fields!Nombre
            
            Set mo_CajaComprobantesDetalleSeleccionado = CompDetalleSeleccionado
        End If
    End With
    Unload Me
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    BuscarProductos
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyEscape
        cmdCerrar.Value = True
    Case vbKeyF2
        cmdAgregar.Value = True
    Case vbKeyF10
        optServicios.Value = True
     Case vbKeyF11
        optBienesInsumos.Value = True
    End Select
End Sub
Sub AdministrarKeyPreview(KeyCode As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            cmdCerrar.Value = True
        Case vbKeyF2
            cmdAgregar.Value = True
    End Select
End Sub


Private Sub Form_Load()
    GenerarRecordsetTemporal
    mo_Formulario.ConfigurarTipoLetra "Tahoma", "9", Me
    Set mo_CajaComprobantesDetalleSeleccionado = Nothing
End Sub

Private Sub grdItems_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    grdItems.Bands(0).Columns("Tipo").Hidden = True
    grdItems.Bands(0).Columns("IdProducto").Hidden = True
    
    grdItems.Bands(0).Columns("Codigo").Header.Caption = "Código"
    grdItems.Bands(0).Columns("Codigo").Width = 1200

    grdItems.Bands(0).Columns("Nombre").Header.Caption = "Nombre"
    grdItems.Bands(0).Columns("Nombre").Width = 4000
    
    grdItems.Bands(0).Columns("PrecioUnitario").Header.Caption = "C.U (S/.)"
    grdItems.Bands(0).Columns("PrecioUnitario").Width = 1200

End Sub

Private Sub optBienesInsumos_Click()
    BuscarProductos
    txtNombre.SetFocus
End Sub

Private Sub optServicios_Click()
    BuscarProductos
    txtNombre.SetFocus
End Sub

Private Sub txtCantidad_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmdAgregar
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsDinero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtNombre_Change()
    BuscarProductos
End Sub

Public Function GetComprobantesDetalleSeleccionado() As DOCajaComprobantesDetalle

    Set GetComprobantesDetalleSeleccionado = mo_CajaComprobantesDetalleSeleccionado
End Function
Private Sub BuscarProductos()
    Dim iTipo As Integer
    Dim rsItems As New Recordset
    Dim iCount As Integer
    
    InitRecordSetItems
    
    If Len(Trim(txtNombre.Text)) < 3 Then
        Exit Sub
    End If
    
    
    If Me.optServicios.Value Then
        iTipo = SIGHComun.sghTipoDetalleComprobante.sghDetalleComprobanteServicios
    Else
        iTipo = SIGHComun.sghTipoDetalleComprobante.sghDetalleComprobanteInsumos
    End If
    
    Set rsItems = mo_AdminCaja.BuscarCoincidenciasEnCatalogos(iTipo, Trim(txtNombre.Text))
    
    Set Me.grdItems.DataSource = rsItems
    
'    iCount = 0
'    Do While Not rsItems.EOF
'        With mrs_ItemsFacturables
'            iCount = iCount + 1
'            .AddNew
'            .Fields!Tipo = rsItems!Tipo
'            .Fields!IdProducto = rsItems!IdProducto
'            .Fields!Codigo = rsItems!Codigo
'            .Fields!Nombre = rsItems!Nombre
'            .Fields!PrecioUnitario = rsItems!PrecioUnitario
'            If iCount > 10 Then
'                Exit Do
'            End If
'        End With
'        rsItems.MoveNext
'    Loop
'    rsItems.Close
    mo_Apariencia.ConfigurarFilasBiColores Me.grdItems, SIGHComun.GrillaConFilasBicolor
    'Set grdItems.DataSource = mrs_ItemsFacturables
End Sub
Private Sub InitRecordSetItems()
    
    Set mrs_ItemsFacturables = New Recordset
    With mrs_ItemsFacturables
          .Fields.Append "Tipo", adVarChar, 2, adFldIsNullable
          .Fields.Append "IdProducto", adVarChar, 2, adFldIsNullable
          .Fields.Append "Codigo", adVarChar, 20, adFldIsNullable
          .Fields.Append "Nombre", adVarChar, 250, adFldIsNullable
          .Fields.Append "PrecioUnitario", adCurrency, 8, adFldIsNullable

          .LockType = adLockOptimistic
          .Open
    End With
    Set Me.grdItems.DataSource = mrs_ItemsFacturables
End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtCantidad
    AdministrarKeyPreview KeyCode
End Sub
