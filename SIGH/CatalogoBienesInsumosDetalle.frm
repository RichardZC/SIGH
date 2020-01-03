VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.Form CatalogoBienesInsumosDetalle 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CatalogosBienesInsumosDetalle"
   ClientHeight    =   8490
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   12030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
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
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   11835
      Begin VB.TextBox txtCodigo 
         Height          =   345
         Left            =   150
         TabIndex        =   0
         Top             =   540
         Width           =   1395
      End
      Begin VB.TextBox txtNombre 
         Height          =   345
         Left            =   1620
         TabIndex        =   1
         Top             =   540
         Width           =   4545
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   6270
         Picture         =   "CatalogoBienesInsumosDetalle.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   570
         Width           =   1305
      End
      Begin VB.CommandButton btnLimpiar 
         Height          =   315
         Left            =   7650
         Picture         =   "CatalogoBienesInsumosDetalle.frx":2C49
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   570
         Width           =   1215
      End
      Begin UltraGrid.SSUltraGrid grdBienesInsumosSeleccionados 
         Height          =   6165
         Left            =   150
         TabIndex        =   4
         Top             =   960
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
         Caption         =   "Lista de bienes e insumos aun no agregados"
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
         TabIndex        =   9
         Top             =   300
         Width           =   4875
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1065
      Left            =   0
      TabIndex        =   5
      Top             =   7320
      Width           =   11850
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "CatalogoBienesInsumosDetalle.frx":5825
         DownPicture     =   "CatalogoBienesInsumosDetalle.frx":5C85
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
         Picture         =   "CatalogoBienesInsumosDetalle.frx":60FA
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "CatalogoBienesInsumosDetalle.frx":656F
         DownPicture     =   "CatalogoBienesInsumosDetalle.frx":6A33
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
         Picture         =   "CatalogoBienesInsumosDetalle.frx":6F1F
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   225
         Width           =   1365
      End
   End
End
Attribute VB_Name = "CatalogoBienesInsumosDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Busca Medicamentos e Insumos
'        Programado por: Castro W
'        Fecha: Agosto 2004
'------------------------------------------------------------------------------------

Dim mo_Teclado As New sighentidades.Teclado
Dim mo_Formulario As New sighentidades.Formulario
Dim mo_CatalogoBienesInsumos As New DOCatalogoBienesInsumos
Dim ml_idUsuario As Long
Dim ms_MensajeError As String
Dim mi_Opcion As sghOpciones
Dim mb_ExistenDatos As Boolean
Dim ml_IdProducto As Long
Dim mo_AdminFacturacion As New ReglasFacturacion
Dim ml_TipoCatalogo As Long
Dim mrs_BienesInsumosSeleccionados As Recordset
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim mrs_BienesInsumos As New Recordset

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
        If ValidarBienesInsumos() Then
            If AgregarBienesInsumos() Then
                MsgBox "Los Bienes e Insumos se agregaron correctamente", vbInformation, Me.Caption
                btnBuscar_Click
            Else
               MsgBox "No se pudo agregar los datos" + Chr(13) + mo_AdminFacturacion.MensajeError, vbExclamation, Me.Caption
            End If
        End If
    Case sghModificar
    Case sghEliminar
    Case sghConsultar
    End Select

End Sub

Function ValidarBienesInsumos() As Boolean
    
    ValidarBienesInsumos = False
    
    Dim oRow As SSRow
    
    Set oRow = Me.grdBienesInsumosSeleccionados.GetRow(ssChildRowFirst)
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
    
    ValidarBienesInsumos = True

End Function

Function AgregarBienesInsumos() As Boolean
    
    AgregarBienesInsumos = False
    
    Dim oRow As SSRow
    
    Set oRow = Me.grdBienesInsumosSeleccionados.GetRow(ssChildRowFirst)
    If Not oRow Is Nothing Then
        AgregarNuevoBienInsumo oRow
        Do While oRow.HasNextSibling
            Set oRow = oRow.GetSibling(ssSiblingRowNext)
            AgregarNuevoBienInsumo oRow
        Loop
    End If
    
    If mo_AdminFacturacion.AgregarBienesInsumosAlCatalogo(ml_TipoCatalogo, mrs_BienesInsumos) Then
        AgregarBienesInsumos = True
    End If
    

End Function

'***************Barrantes D**************
'***************Registra PRECIO UNITARIO incluyendo decimales
'***************para los MEDICAMENTOS ELEGIDOS
Sub AgregarNuevoBienInsumo(oRow As SSRow)
        
        If oRow.Cells("Agregar").Value Then
            Dim lnPrecio As Double
            On Error Resume Next
            If oRow.Cells("PrecioUnitario").Value <> "" Then
               lnPrecio = oRow.Cells("PrecioUnitario").Value
            End If
            If lnPrecio > 0 Then
                mrs_BienesInsumos.AddNew
                mrs_BienesInsumos.Fields!IdProducto = oRow.Cells("IdProducto").Value
                mrs_BienesInsumos.Fields!PrecioUnitario = lnPrecio
                mrs_BienesInsumos.Fields!idUsuario = ml_idUsuario
                mrs_BienesInsumos.Update
            End If
        End If
        
End Sub

Sub CrearRecordsetAgregarBienInsumo()

    With mrs_BienesInsumos
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
        CrearRecordsetAgregarBienInsumo
    Case sghModificar
    Me.Hide
    Case sghEliminar
    Me.Hide
    Case sghConsultar
    Me.Hide
    End Select
    
End Sub

Sub RealizarBusqueda()
Dim oDoCatalogoBienInsumo As New DOCatalogoBienesInsumos
Dim oRecorset As Recordset

    oDoCatalogoBienInsumo.codigo = Trim(txtCodigo)
    oDoCatalogoBienInsumo.Nombre = Trim(txtNombre)
    
    Set grdBienesInsumosSeleccionados.DataSource = mo_AdminFacturacion.CatalogoBienesInsumosSeleccionarPorTipoCatalogo(oDoCatalogoBienInsumo, ml_TipoCatalogo)
    
    mo_Apariencia.ConfigurarFilasBiColores Me.grdBienesInsumosSeleccionados, sighentidades.GrillaConFilasBicolor

End Sub

Private Sub grdBienesInsumosSeleccionados_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    
    grdBienesInsumosSeleccionados.Bands(0).Columns("IdProducto").Hidden = True
    
    grdBienesInsumosSeleccionados.Bands(0).Columns("Codigo").Header.Caption = "Código"
    grdBienesInsumosSeleccionados.Bands(0).Columns("Codigo").Width = 1000

    grdBienesInsumosSeleccionados.Bands(0).Columns("Nombre").Header.Caption = "Nombre"
    grdBienesInsumosSeleccionados.Bands(0).Columns("Nombre").Width = 7500

    grdBienesInsumosSeleccionados.Bands(0).Columns.Add "PrecioUnitario", "PrecioUnitario"
    grdBienesInsumosSeleccionados.Bands(0).Columns("PrecioUnitario").Header.Caption = "Precio Unit. S/."
    grdBienesInsumosSeleccionados.Bands(0).Columns("PrecioUnitario").Width = 1500

    grdBienesInsumosSeleccionados.Bands(0).Columns.Add "Agregar", "Agregar"
    grdBienesInsumosSeleccionados.Bands(0).Columns("Agregar").DataType = ssDataTypeBoolean
    grdBienesInsumosSeleccionados.Bands(0).Columns("Agregar").Header.Caption = "¿Agregar?"
    grdBienesInsumosSeleccionados.Bands(0).Columns("Agregar").Width = 1500
    grdBienesInsumosSeleccionados.Bands(0).Columns("Agregar").Style = ssStyleCheckBox

    grdBienesInsumosSeleccionados.Override.ExpandRowsOnLoad = ssExpandOnLoadNo

End Sub

