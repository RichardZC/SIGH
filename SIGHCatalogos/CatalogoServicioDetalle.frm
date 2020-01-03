VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.Form CatalogoServiciosDetalle 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CatalogoServiciosDetalle"
   ClientHeight    =   8490
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12030
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   12030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Height          =   1065
      Left            =   90
      TabIndex        =   7
      Top             =   7350
      Width           =   11850
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "CatalogoServicioDetalle.frx":0000
         DownPicture     =   "CatalogoServicioDetalle.frx":04C4
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
         Left            =   5970
         Picture         =   "CatalogoServicioDetalle.frx":09B0
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "CatalogoServicioDetalle.frx":0E9C
         DownPicture     =   "CatalogoServicioDetalle.frx":12FC
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
         Left            =   4425
         Picture         =   "CatalogoServicioDetalle.frx":1771
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   225
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Catalogo Base"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7305
      Left            =   90
      TabIndex        =   5
      Top             =   30
      Width           =   11835
      Begin VB.CommandButton btnLimpiar 
         Height          =   315
         Left            =   7650
         Picture         =   "CatalogoServicioDetalle.frx":1BE6
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   570
         Width           =   1215
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   6270
         Picture         =   "CatalogoServicioDetalle.frx":47C2
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   570
         Width           =   1305
      End
      Begin VB.TextBox txtNombre 
         Height          =   345
         Left            =   1620
         TabIndex        =   1
         Top             =   540
         Width           =   4545
      End
      Begin VB.TextBox txtCodigo 
         Height          =   345
         Left            =   150
         TabIndex        =   0
         Top             =   540
         Width           =   1395
      End
      Begin UltraGrid.SSUltraGrid grdServiciosSeleccionados 
         Height          =   6165
         Left            =   150
         TabIndex        =   4
         Top             =   990
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   10874
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
         Caption         =   "Lista de servicios aun no agregados"
      End
      Begin VB.Label Label1 
         Caption         =   "Código                     Nombre    "
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
         Left            =   180
         TabIndex        =   6
         Top             =   300
         Width           =   4875
      End
   End
End
Attribute VB_Name = "CatalogoServiciosDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Busca Procedimiento CPT
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'
'------------------------------------------------------------------------------------

Dim mo_Teclado As New sighentidades.Teclado
Dim mo_Formulario As New sighentidades.Formulario
Dim mo_CatalogoServicios As New DOCatalogoServicio
Dim ml_idUsuario As Long
Dim ms_MensajeError As String
Dim mi_Opcion As sghOpciones
Dim mb_ExistenDatos As Boolean
Dim ml_IdProducto As Long
Dim mo_AdminFacturacion As New ReglasFacturacion
Dim ml_TipoCatalogo As Long
Dim mrs_ServiciosSeleccionados As Recordset
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim mrs_Servicios As New Recordset

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
Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property
Property Get idUsuario() As Long
   idUsuario = ml_idUsuario
End Property
Property Let TipoCatalogo(lValue As Long)
    ml_TipoCatalogo = lValue
End Property
Property Get TipoCatalogo() As Long
    TipoCatalogo = ml_TipoCatalogo
End Property


Private Sub btnAceptar_Click()

    Select Case mi_Opcion
    Case sghAgregar
        If ValidarServicios() Then
            If AgregarServicios() Then
                MsgBox "Los servicios se agregaron correctamente", vbInformation, Me.Caption
                btnBuscar_Click
            End If
        End If
    Case sghModificar
    Case sghEliminar
    Case sghConsultar
    End Select

End Sub

Function ValidarServicios() As Boolean
    
    ValidarServicios = False
    
    Dim oRow As SSRow
    
    Set oRow = Me.grdServiciosSeleccionados.GetRow(ssChildRowFirst)
    If Not oRow Is Nothing Then
        
        If oRow.Cells("Agregar").Value Then
            If oRow.Cells("PrecioUnitario").Value <= 0 Then
                MsgBox "Ingrese el precio unitario del producto: <" & oRow.Cells("Nombre").Value & ">", vbInformation, Me.Caption
                Exit Function
            End If
        End If
        
        Do While oRow.HasNextSibling
            Set oRow = oRow.GetSibling(ssSiblingRowNext)
            If oRow.Cells("Agregar").Value Then
                If oRow.Cells("PrecioUnitario").Value <= 0 Then
                    MsgBox "Ingrese el precio unitario del producto: <" & oRow.Cells("Nombre").Value & ">", vbInformation, Me.Caption
                    Exit Function
                End If
            End If
        Loop
    End If
    
    ValidarServicios = True

End Function

Function AgregarServicios() As Boolean
    
    AgregarServicios = False
    
    Dim oRow As SSRow
    
    Set oRow = Me.grdServiciosSeleccionados.GetRow(ssChildRowFirst)
    If Not oRow Is Nothing Then
        AgregarNuevoServicio oRow
        Do While oRow.HasNextSibling
            Set oRow = oRow.GetSibling(ssSiblingRowNext)
            AgregarNuevoServicio oRow
        Loop
    End If
    
    If mo_AdminFacturacion.AgregarServiciosAlCatalogo(ml_TipoCatalogo, mrs_Servicios) Then
        AgregarServicios = True
    End If
    

End Function

'***************Barrantes D**************
'***************Al GRABAR el SERVICIO para CONVENIO/SIS/SOAT/NORMAL/...
'***************el Precio tome tambien DECIMALES
Sub AgregarNuevoServicio(oRow As SSRow)
        
        If oRow.Cells("Agregar").Value Then
            Dim lnPrecio As Double
            lnPrecio = oRow.Cells("PrecioUnitario").Value
            mrs_Servicios.AddNew
            mrs_Servicios.Fields!IdProducto = oRow.Cells("IdProducto").Value
            mrs_Servicios.Fields("precioUnitario").Value = lnPrecio
            mrs_Servicios.Fields!idUsuario = ml_idUsuario
            mrs_Servicios.Update
        End If
        
End Sub

Sub CrearRecordsetAgregarServicio()

    With mrs_Servicios
          .Fields.Append "IdProducto", adInteger, 4
          .Fields.Append "PrecioUnitario", adCurrency
          .Fields.Append "Agregar", adBoolean
          .Fields.Append "IdUsuario", adInteger, 4
          .CursorType = adOpenStatic
          .LockType = adLockOptimistic
          .Open
    End With

End Sub

Private Sub btnBuscar_Click()
    RealizarBusqueda

End Sub

Private Sub btnCancelar_Click()
Me.Hide
End Sub

Private Sub Form_Load()
    
    RealizarBusqueda
    Select Case mi_Opcion
    Case sghAgregar
        CrearRecordsetAgregarServicio
    Case sghModificar
    Case sghEliminar
    Case sghConsultar
    End Select
    
End Sub

Sub RealizarBusqueda()
Dim oDoCatalogoServicio As New DOCatalogoServicio
Dim oRecorset As Recordset

    oDoCatalogoServicio.Codigo = Trim(txtCodigo)
    oDoCatalogoServicio.Nombre = Trim(txtNombre)
    
    Set grdServiciosSeleccionados.DataSource = mo_AdminFacturacion.CatalogoServiciosSeleccionarPorTipoCatalogo(oDoCatalogoServicio, ml_TipoCatalogo)
    
    mo_Apariencia.ConfigurarFilasBiColores Me.grdServiciosSeleccionados, sighentidades.GrillaConFilasBicolor

End Sub

Private Sub grdServiciosSeleccionados_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    
    grdServiciosSeleccionados.Bands(0).Columns("IdProducto").Hidden = True
    
    grdServiciosSeleccionados.Bands(0).Columns("Codigo").Header.Caption = "Código"
    grdServiciosSeleccionados.Bands(0).Columns("Codigo").Width = 1000

    grdServiciosSeleccionados.Bands(0).Columns("Nombre").Header.Caption = "Nombre"
    grdServiciosSeleccionados.Bands(0).Columns("Nombre").Width = 7500

    grdServiciosSeleccionados.Bands(0).Columns.Add "PrecioUnitario", "PrecioUnitario"
    grdServiciosSeleccionados.Bands(0).Columns("PrecioUnitario").Header.Caption = "Precio Unit. S/."
    grdServiciosSeleccionados.Bands(0).Columns("PrecioUnitario").Width = 1500

    grdServiciosSeleccionados.Bands(0).Columns.Add "Agregar", "Agregar"
    grdServiciosSeleccionados.Bands(0).Columns("Agregar").DataType = ssDataTypeBoolean
    grdServiciosSeleccionados.Bands(0).Columns("Agregar").Header.Caption = "¿Agregar?"
    grdServiciosSeleccionados.Bands(0).Columns("Agregar").Width = 1500
    grdServiciosSeleccionados.Bands(0).Columns("Agregar").Style = ssStyleCheckBox

End Sub
