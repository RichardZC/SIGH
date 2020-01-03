VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form HerrActualizaSaldo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Actualiza Saldos"
   ClientHeight    =   8385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11820
   Icon            =   "HerrActualizaSaldo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8385
   ScaleWidth      =   11820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   8325
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   14684
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Actualiza Saldos"
      TabPicture(0)   =   "HerrActualizaSaldo.frx":0CCA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame6"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Actualiza Fecha Vencimiento"
      TabPicture(1)   =   "HerrActualizaSaldo.frx":0CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(1)=   "grdFarmSaldoDetallado"
      Tab(1).Control(2)=   "cmdActualizaFVencimiento"
      Tab(1).Control(3)=   "txtCodigoSismed"
      Tab(1).ControlCount=   4
      Begin VB.TextBox txtCodigoSismed 
         Height          =   345
         Left            =   -72930
         TabIndex        =   23
         Top             =   2760
         Width           =   1425
      End
      Begin VB.CommandButton cmdActualizaFVencimiento 
         Caption         =   "Actualiza columna:  ""F.Venc.Nueva""   en las tablas"
         Height          =   705
         Left            =   -74910
         TabIndex        =   22
         Top             =   4020
         Width           =   11415
      End
      Begin VB.Frame Frame6 
         Enabled         =   0   'False
         Height          =   7125
         Left            =   120
         TabIndex        =   10
         Top             =   1140
         Width           =   11655
         Begin VB.CommandButton Command2 
            Caption         =   "Limpia Lista"
            Height          =   345
            Left            =   8280
            TabIndex        =   18
            ToolTipText     =   "Pone a todos los MEDICAMENTOS el saldo"
            Top             =   5670
            Width           =   1335
         End
         Begin VB.TextBox txtBuscaMed 
            Height          =   345
            Left            =   2040
            TabIndex        =   17
            Top             =   5670
            Width           =   5685
         End
         Begin VB.CheckBox chkSoloInventario 
            Caption         =   "Solo graba INVENTARIO anulado"
            Height          =   285
            Left            =   7650
            TabIndex        =   16
            Top             =   6150
            Value           =   1  'Checked
            Width           =   3615
         End
         Begin VB.CommandButton cmdSaldosT 
            Caption         =   "..."
            Height          =   345
            Left            =   10770
            TabIndex        =   15
            ToolTipText     =   "Pone a todos los MEDICAMENTOS el saldo"
            Top             =   5670
            Width           =   465
         End
         Begin VB.TextBox txtSaldoT 
            Height          =   315
            Left            =   9630
            TabIndex        =   14
            Text            =   "0"
            Top             =   5670
            Width           =   1065
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Del"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   11160
            TabIndex        =   13
            Top             =   450
            Width           =   435
         End
         Begin VB.CommandButton Command4 
            Caption         =   "New"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   11160
            TabIndex        =   12
            ToolTipText     =   "New"
            Top             =   180
            Width           =   435
         End
         Begin VB.CommandButton cmdIgualaSaldos 
            Caption         =   $"HerrActualizaSaldo.frx":0D02
            Height          =   885
            Left            =   60
            TabIndex        =   11
            Top             =   6090
            Width           =   7215
         End
         Begin MSDataGridLib.DataGrid grdSaldosAjuste 
            Height          =   5445
            Left            =   30
            TabIndex        =   19
            Top             =   150
            Width           =   11025
            _ExtentX        =   19447
            _ExtentY        =   9604
            _Version        =   393216
            Enabled         =   -1  'True
            HeadLines       =   1
            RowHeight       =   15
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   3
            BeginProperty Column00 
               DataField       =   "codigo"
               Caption         =   "Código"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   "medicamento"
               Caption         =   "Medicamento"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column02 
               DataField       =   "saldo"
               Caption         =   "saldo a la F.ajuste"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "#,##0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   1
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
                  Locked          =   -1  'True
                  ColumnWidth     =   975.118
               EndProperty
               BeginProperty Column01 
                  Locked          =   -1  'True
                  ColumnWidth     =   8100.284
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   1395.213
               EndProperty
            EndProperty
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "....."
            Height          =   255
            Left            =   180
            TabIndex        =   20
            Top             =   5730
            Width           =   1725
         End
      End
      Begin VB.Frame Frame4 
         Height          =   735
         Left            =   90
         TabIndex        =   1
         Top             =   390
         Width           =   11655
         Begin VB.CommandButton cmdBuscarS 
            Caption         =   "Buscar"
            Height          =   375
            Left            =   10620
            TabIndex        =   3
            Top             =   240
            Width           =   945
         End
         Begin VB.TextBox txtDctoAjuste 
            Height          =   315
            Left            =   5820
            MaxLength       =   4
            TabIndex        =   2
            Text            =   "AJ01"
            Top             =   240
            Width           =   645
         End
         Begin MSMask.MaskEdBox txtFAjuste 
            Height          =   315
            Left            =   8010
            TabIndex        =   4
            Top             =   240
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSDataListLib.DataCombo cmbAlmacenA 
            Height          =   345
            Left            =   810
            TabIndex        =   5
            Top             =   180
            Width           =   3195
            _ExtentX        =   5636
            _ExtentY        =   609
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            Text            =   "DataCombo1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSMask.MaskEdBox txtHajuste 
            Height          =   315
            Left            =   9240
            TabIndex        =   6
            Top             =   240
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   5
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin VB.Label Label13 
            Caption         =   "N° Dcto Ajuste:"
            Height          =   255
            Left            =   4620
            TabIndex        =   9
            Top             =   270
            Width           =   1185
         End
         Begin VB.Label Label15 
            Caption         =   "F.Ajuste:"
            Height          =   225
            Left            =   7380
            TabIndex        =   8
            Top             =   270
            Width           =   705
         End
         Begin VB.Label Label14 
            Caption         =   "Almacen Destino:"
            Height          =   465
            Left            =   60
            TabIndex        =   7
            Top             =   210
            Width           =   855
         End
      End
      Begin MSDataGridLib.DataGrid grdFarmSaldoDetallado 
         Height          =   2115
         Left            =   -74910
         TabIndex        =   21
         Top             =   450
         Width           =   11445
         _ExtentX        =   20188
         _ExtentY        =   3731
         _Version        =   393216
         Enabled         =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "codigo"
            Caption         =   "Código"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "nombre"
            Caption         =   "Medicamento"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "cantidad"
            Caption         =   "saldo "
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "lote"
            Caption         =   "Lote"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "fechaVencimiento"
            Caption         =   "F.Venc.Actual"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "fechaVencimientoN"
            Caption         =   "F.Venc.Nueva"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   3
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               Locked          =   -1  'True
               ColumnWidth     =   975.118
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               ColumnWidth     =   4965.166
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               ColumnWidth     =   720
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
               ColumnWidth     =   1065.26
            EndProperty
            BeginProperty Column04 
               Locked          =   -1  'True
               ColumnWidth     =   1425.26
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1755.213
            EndProperty
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "Ingrese el CODIGO sismed y pulse ENTER"
         Height          =   465
         Left            =   -74910
         TabIndex        =   24
         Top             =   2670
         Width           =   1905
      End
   End
End
Attribute VB_Name = "HerrActualizaSaldo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Muestra Saldos de Almacén
'        Programado por: Barrantes D
'        Fecha: Diciembre 2013
'
'------------------------------------------------------------------------------------
Option Explicit
Dim oConexionFox As New ADODB.Connection
Dim oRsAlmacenes As New ADODB.Recordset
Dim oRsAlmacenes1 As New ADODB.Recordset
Dim oRsUsuario As New ADODB.Recordset
Dim oRsSaldosAjuste As New ADODB.Recordset
Dim oRsFarmSaldoDetallado As New ADODB.Recordset
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
Dim mo_farmMovimiento As New DoFarmMovimiento
Dim mo_farmMovimientoNotaIngreso As New DOfarmMovimientoNotaIngreso
Dim mo_FarmInventario As New DoFarmInventario
Dim oDoProveedores As New DoProveedores
Dim lcSql As String
Dim lbEsFarmacia As Boolean
Dim mo_mensajeError As String
Const ml_idUsuario As Long = 738
Dim wxConexionRed As String









Private Sub cmbAlmacenA_Click(Area As Integer)
    txtDctoAjuste.Text = "AJ0" & cmbAlmacenA.BoundText
End Sub

Private Sub cmdActualizaPrecios_Click()
    Dim oRsTmp11 As New Recordset
    Dim oRsTmp12 As New Recordset
    Dim oRsFox As New ADODB.Recordset
    Dim lnIdProducto As Long
    Dim lcSql As String
    '
    oRsFox.Open "SELECT * from mProducto", oConexionFox, adOpenKeyset, adLockOptimistic
    oRsFox.MoveFirst
    Do While Not oRsFox.EOF
       oRsTmp12.Open "select * from FactCatalogoBienesInsumos where codigo='" & Trim(oRsFox.Fields!medCod) & "'", wxConexionRed, adOpenKeyset, adLockOptimistic
       If oRsTmp12.RecordCount > 0 Then
          lcSql = "update FactCatalogoBienesInsumosHosp set PrecioUnitario=" & oRsFox.Fields!prdPreOpe & " where idProducto=" & oRsTmp12.Fields!idProducto
          oRsTmp11.Open lcSql, wxConexionRed, adOpenKeyset, adLockOptimistic
       End If
       oRsTmp12.Close
       oRsFox.MoveNext
    Loop
    oRsFox.Close
    Unload Me
End Sub

Private Sub cmdActualizaFVencimiento_Click()
        On Error GoTo errActFV
        If oRsFarmSaldoDetallado.RecordCount = 0 Then
           MsgBox "No hay Items en la LISTA"
        End If
        If MsgBox("Esta seguro?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
           Dim oRsTmp As New ADODB.Recordset
           oRsFarmSaldoDetallado.MoveFirst
           Do While Not oRsFarmSaldoDetallado.EOF
              If Not IsNull(oRsFarmSaldoDetallado.Fields!fechaVencimientoN) Then
                    lcSql = "update FarmInventarioDetalle set FechaVencimiento='" & oRsFarmSaldoDetallado.Fields!fechaVencimientoN & _
                          "' where idProducto=" & oRsFarmSaldoDetallado.Fields!idProducto & " and Lote='" & oRsFarmSaldoDetallado.Fields!Lote & _
                          "' and fechaVencimiento='" & oRsFarmSaldoDetallado.Fields!FechaVencimiento & "'"
                    oRsTmp.Open lcSql, sighEntidades.CadenaConexion, adOpenKeyset, adLockOptimistic
                    lcSql = "update FarmMovimientoDetalle set FechaVencimiento='" & oRsFarmSaldoDetallado.Fields!fechaVencimientoN & _
                          "' where idProducto=" & oRsFarmSaldoDetallado.Fields!idProducto & " and Lote='" & oRsFarmSaldoDetallado.Fields!Lote & _
                          "' and fechaVencimiento='" & oRsFarmSaldoDetallado.Fields!FechaVencimiento & "'"
                    oRsTmp.Open lcSql, sighEntidades.CadenaConexion, adOpenKeyset, adLockOptimistic
                    lcSql = "update FarmSaldoDetallado set FechaVencimiento='" & oRsFarmSaldoDetallado.Fields!fechaVencimientoN & _
                          "' where idProducto=" & oRsFarmSaldoDetallado.Fields!idProducto & " and Lote='" & oRsFarmSaldoDetallado.Fields!Lote & _
                          "' and fechaVencimiento='" & oRsFarmSaldoDetallado.Fields!FechaVencimiento & "'"
                    oRsTmp.Open lcSql, sighEntidades.CadenaConexion, adOpenKeyset, adLockOptimistic
                    
              End If
              oRsFarmSaldoDetallado.MoveNext
           Loop
           Unload Me
        End If
        Exit Sub
errActFV:
    MsgBox Err.Description
End Sub

Private Sub cmdBuscarS_Click()
    If cmbAlmacenA.Text = "" Then
       MsgBox "Elija el Almacen/Farmacia", vbCritical, Me.Caption
       Exit Sub
    End If
    If txtDctoAjuste.Text = "" Then
       MsgBox "Ingrese el Nro Documento de AJUSTE INVENTARIO", vbCritical, Me.Caption
       Exit Sub
    End If
    If txtFAjuste.Text = sighEntidades.FECHA_VACIA_DMY Then
       MsgBox "Ingrese la Fecha ", vbCritical, Me.Caption
       Exit Sub
    End If
    If txtHajuste.Text = sighEntidades.HORA_VACIA_HM Then
       MsgBox "Ingrese la Hora ", vbCritical, Me.Caption
       Exit Sub
    End If
    Dim oRsTmp As New Recordset
    
    lcSql = "SELECT     dbo.farmInventarioCabecera.*, dbo.FactCatalogoBienesInsumos.Codigo, dbo.FactCatalogoBienesInsumos.Nombre, " & _
            "          dbo.farmInventario.numeroInventario , dbo.farmInventario.FechaCreacion" & _
            " FROM         dbo.farmInventarioCabecera LEFT OUTER JOIN" & _
            "          dbo.FactCatalogoBienesInsumos ON dbo.farmInventarioCabecera.idProducto = dbo.FactCatalogoBienesInsumos.IdProducto LEFT OUTER JOIN" & _
            "          dbo.farmInventario ON dbo.farmInventarioCabecera.idInventario = dbo.farmInventario.idInventario" & _
            " WHERE    dbo.farmInventario.numeroInventario='" & Trim(txtDctoAjuste.Text) & "'" & _
            "          and dbo.farmInventario.idEstadoInventario=0 and dbo.farmInventario.idAlmacen=" & cmbAlmacenA.BoundText    'este inventario debe estar ANULADO
    oRsTmp.Open lcSql, wxConexionRed, adOpenKeyset, adLockOptimistic
    If oRsTmp.RecordCount > 0 Then
       '************ya esta registrado el Documento
       
       oRsTmp.MoveFirst
       Do While Not oRsTmp.EOF
          oRsSaldosAjuste.AddNew
          oRsSaldosAjuste.Fields!idProducto = oRsTmp.Fields!idProducto
          oRsSaldosAjuste.Fields!codigo = oRsTmp.Fields!codigo
          oRsSaldosAjuste.Fields!medicamento = oRsTmp.Fields!nombre
          oRsSaldosAjuste.Fields!saldo = oRsTmp.Fields!Cantidad
          oRsSaldosAjuste.Fields!RegistroSanitario = "."
          oRsSaldosAjuste.Update
          oRsTmp.MoveNext
       Loop
    Else
       '************registran por  primera vez
       oRsTmp.Close
       lcSql = "select * from FactCatalogoBienesInsumos order by Nombre"
       oRsTmp.Open lcSql, wxConexionRed, adOpenKeyset, adLockOptimistic
       oRsTmp.MoveFirst
       Do While Not oRsTmp.EOF
          oRsSaldosAjuste.AddNew
          oRsSaldosAjuste.Fields!idProducto = oRsTmp.Fields!idProducto
          oRsSaldosAjuste.Fields!codigo = oRsTmp.Fields!codigo
          oRsSaldosAjuste.Fields!medicamento = oRsTmp.Fields!nombre
          oRsSaldosAjuste.Fields!RegistroSanitario = "."
          oRsSaldosAjuste.Fields!saldo = 0
          oRsSaldosAjuste.Update
          oRsTmp.MoveNext
       Loop
    End If
    Set oRsTmp = Nothing
    Frame4.Enabled = False
    Frame6.Enabled = True
End Sub

Private Sub cmdElimina_Click()

        Dim oRsInvCab As New ADODB.Recordset
        'actualiza correlativos a CERO
        oRsInvCab.Open "update FarmTipoDocumentos set correlativo=0", wxConexionRed, adOpenKeyset, adLockOptimistic
        oRsInvCab.Open "update FarmRelMod set DocumentoUltimoNumero='0'", wxConexionRed, adOpenKeyset, adLockOptimistic
        'ELIMINA DATOS
        oRsInvCab.Open "DELETE FROM FARMINVENTARIOCABECERA", wxConexionRed, adOpenKeyset, adLockOptimistic
        oRsInvCab.Open "DELETE FROM FARMINVENTARIODETALLE", wxConexionRed, adOpenKeyset, adLockOptimistic
        oRsInvCab.Open "DELETE FROM FARMINVENTARIO", wxConexionRed, adOpenKeyset, adLockOptimistic
        oRsInvCab.Open "DELETE FROM FARMMOVIMIENTONOTAINGRESO", wxConexionRed, adOpenKeyset, adLockOptimistic
        oRsInvCab.Open "DELETE FROM FARMMOVIMIENTOPROGRAMAS", wxConexionRed, adOpenKeyset, adLockOptimistic
        oRsInvCab.Open "DELETE FROM FacturacionBienesPagos", wxConexionRed, adOpenKeyset, adLockOptimistic
        oRsInvCab.Open "DELETE FROM FactOrdenesBienes", wxConexionRed, adOpenKeyset, adLockOptimistic
        oRsInvCab.Open "DELETE FROM FacturacionBienesFinanciamientos", wxConexionRed, adOpenKeyset, adLockOptimistic
        oRsInvCab.Open "DELETE FROM FARMMOVIMIENTOVENTASDETALLE", wxConexionRed, adOpenKeyset, adLockOptimistic
        oRsInvCab.Open "DELETE FROM FARMMOVIMIENTOVENTAS", wxConexionRed, adOpenKeyset, adLockOptimistic
        oRsInvCab.Open "DELETE FROM FARMMOVIMIENTODETALLE", wxConexionRed, adOpenKeyset, adLockOptimistic
        oRsInvCab.Open "DELETE FROM FARMMOVIMIENTO", wxConexionRed, adOpenKeyset, adLockOptimistic
        oRsInvCab.Open "DELETE FROM FARMPREVENTADETALLE", wxConexionRed, adOpenKeyset, adLockOptimistic
        oRsInvCab.Open "DELETE FROM FARMPREVENTA", wxConexionRed, adOpenKeyset, adLockOptimistic
        oRsInvCab.Open "DELETE FROM FARMSALDODETALLADO", wxConexionRed, adOpenKeyset, adLockOptimistic
        oRsInvCab.Open "DELETE FROM FARMSALDO", wxConexionRed, adOpenKeyset, adLockOptimistic
        oRsInvCab.Open "delete from CajaComprobantesPago where idTipoOrden>1", wxConexionRed, adOpenKeyset, adLockOptimistic
        oRsInvCab.Open "select * from factCatalogoBienesInsumosHosp", wxConexionRed, adOpenKeyset, adLockOptimistic
        Unload Me
End Sub

Private Sub cmdIgualaSaldos_Click()
    If MsgBox("Esta seguro?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
       Dim oRsTmp As New ADODB.Recordset
       Dim oRsTmp1 As New ADODB.Recordset
       Dim oRsTmp2 As New ADODB.Recordset
       Dim lnTotal As Long, lnIdProducto As Long
       Dim lcLote As String, ldFvencimiento As Date, ldFechaProceso As Date
       Dim lnCantAjusEnt As Long, lnCantAjusSal As Long
       Dim lbSigue As Boolean, lnTotalNI As Double, lnTotalNS As Double
       Dim lnSaldoGalenHos As Long, lnPrecioAjuste As Double
       Dim lcErrores As String
       Dim lcDocumento As String
       On Error GoTo ErrSald
       ldFechaProceso = CDate(txtFAjuste.Text & " " & txtHajuste.Text)
       lcDocumento = txtDctoAjuste.Text & txtFAjuste.Text & txtHajuste.Text
       'Elimina detalle Anterior
       lcSql = "select * from FarmMovimiento where documentoNumero='" & lcDocumento & _
                    "' and documentoIdTipo=10 and idTipoConcepto=20"
       oRsTmp1.Open lcSql, wxConexionRed, adOpenKeyset, adLockOptimistic
       If oRsTmp1.RecordCount > 0 Then
          oRsTmp1.MoveFirst
          Do While Not oRsTmp1.EOF
            lbSigue = True
            If oRsTmp1.Fields!movTipo = "E" Then
               If oRsTmp1.Fields!idAlmacenDestino <> Val(cmbAlmacenA.BoundText) Then
                  lbSigue = False
               End If
            Else
               If oRsTmp1.Fields!idAlmacenOrigen <> Val(cmbAlmacenA.BoundText) Then
                  lbSigue = False
               End If
            End If
            If lbSigue = True And chkSoloInventario.Value = 0 Then
                lcSql = "delete from FarmMovimientoDetalle where MovNumero='" & oRsTmp1.Fields!movNumero & "' and MovTipo='" & oRsTmp1.Fields!movTipo & "'"
                oRsTmp2.Open lcSql, wxConexionRed, adOpenKeyset, adLockOptimistic
                lcSql = "delete from FarmMovimientoNotaIngreso where MovNumero='" & oRsTmp1.Fields!movNumero & "' and MovTipo='" & oRsTmp1.Fields!movTipo & "'"
                oRsTmp2.Open lcSql, wxConexionRed, adOpenKeyset, adLockOptimistic
                lcSql = "delete from FarmMovimiento where MovNumero='" & oRsTmp1.Fields!movNumero & "' and MovTipo='" & oRsTmp1.Fields!movTipo & "'"
                oRsTmp2.Open lcSql, wxConexionRed, adOpenKeyset, adLockOptimistic
            End If
            oRsTmp1.MoveNext
          Loop
       End If
       oRsTmp1.Close
       'Proceso
       
       Set oRsTmp = mo_ReglasFarmacia.farmRegeneraSaldos
       lnTotal = oRsTmp.RecordCount
       lnTotalNI = 0: lnTotalNS = 0: lcErrores = ""
       If lnTotal > 0 Then
            'Actualiza por cada Producto el AJUSTE DE INGRESO, AJUSTE DE SALIDA
            oRsTmp.MoveFirst
            Do While Not oRsTmp.EOF
               lnIdProducto = oRsTmp.Fields!idProducto
If oRsTmp.Fields!idProducto = 8 Then
   lnSaldoGalenHos = 0
End If
               ldFvencimiento = oRsTmp.Fields!FechaVencimiento
               lcLote = oRsTmp.Fields!Lote
               lnSaldoGalenHos = 0
               lnPrecioAjuste = 0
               Do While Not oRsTmp.EOF And lnIdProducto = oRsTmp.Fields!idProducto
                  If oRsTmp.Fields!FechaCreacion <= ldFechaProceso Then
                        If oRsTmp.Fields!movTipo = "E" Then
                           If oRsTmp.Fields!idAlmacenDestino = Val(cmbAlmacenA.BoundText) Then
                                lnSaldoGalenHos = lnSaldoGalenHos + oRsTmp.Fields!Cantidad
                                If oRsTmp.Fields!FechaVencimiento > ldFvencimiento Then
                                  ldFvencimiento = oRsTmp.Fields!FechaVencimiento
                                  lcLote = oRsTmp.Fields!Lote
                                End If
                                lnPrecioAjuste = oRsTmp.Fields!Precio
                           End If
                        Else
                           If oRsTmp.Fields!idAlmacenOrigen = Val(cmbAlmacenA.BoundText) Then
                                lnSaldoGalenHos = lnSaldoGalenHos - oRsTmp.Fields!Cantidad
                           End If
                        End If
                  End If
                  oRsTmp.MoveNext
                  If oRsTmp.EOF Then
                     Exit Do
                  End If
               Loop
               oRsSaldosAjuste.MoveFirst
               oRsSaldosAjuste.Find "idProducto=" & lnIdProducto
               If Not oRsSaldosAjuste.EOF Then
                  If lnPrecioAjuste = 0 Then
                     lnPrecioAjuste = 1
                  End If
                  If oRsSaldosAjuste.Fields!saldo > lnSaldoGalenHos Then
                     oRsSaldosAjuste.Fields!Cantidad = oRsSaldosAjuste.Fields!saldo - lnSaldoGalenHos
                     oRsSaldosAjuste.Fields!movTipo = "E"
                     oRsSaldosAjuste.Fields!Precio = lnPrecioAjuste
                     oRsSaldosAjuste.Fields!Total = Round((oRsSaldosAjuste.Fields!saldo - lnSaldoGalenHos) * lnPrecioAjuste, 2)
                     oRsSaldosAjuste.Fields!Lote = lcLote
                     oRsSaldosAjuste.Fields!FechaVencimiento = ldFvencimiento
                     oRsSaldosAjuste.Update
                     lnTotalNI = lnTotalNI + oRsSaldosAjuste.Fields!Total
                  ElseIf oRsSaldosAjuste.Fields!saldo <= lnSaldoGalenHos Then
                     oRsSaldosAjuste.Fields!Cantidad = lnSaldoGalenHos - oRsSaldosAjuste.Fields!saldo
                     oRsSaldosAjuste.Fields!movTipo = "S"
                     oRsSaldosAjuste.Fields!Precio = lnPrecioAjuste
                     oRsSaldosAjuste.Fields!Total = Round((lnSaldoGalenHos - oRsSaldosAjuste.Fields!saldo) * lnPrecioAjuste, 2)
                     oRsSaldosAjuste.Fields!Lote = lcLote
                     oRsSaldosAjuste.Fields!FechaVencimiento = ldFvencimiento
                     oRsSaldosAjuste.Update
                     lnTotalNS = lnTotalNS + oRsSaldosAjuste.Fields!Total
                  End If
               End If
            Loop
            '
            'Si no  hay LOTE y F.Vencimiento le asigna lote=1234 y FV=01/02/2010
            oRsSaldosAjuste.MoveFirst
            Do While Not oRsSaldosAjuste.EOF
               If IsNull(oRsSaldosAjuste.Fields!Lote) Or oRsSaldosAjuste.Fields!Lote = "" Then
                  oRsSaldosAjuste.Fields!Lote = "1234"
               End If
               If IsNull(oRsSaldosAjuste.Fields!FechaVencimiento) Then
                  oRsSaldosAjuste.Fields!FechaVencimiento = CDate("01/02/2010")
               End If
               If IsNull(oRsSaldosAjuste.Fields!movTipo) Or oRsSaldosAjuste.Fields!movTipo = "" Then
                     'Producto Nuevo
                     lnPrecioAjuste = 1
                     lnSaldoGalenHos = oRsSaldosAjuste.Fields!saldo
                     oRsSaldosAjuste.Fields!Cantidad = oRsSaldosAjuste.Fields!saldo
                     oRsSaldosAjuste.Fields!movTipo = "E"
                     oRsSaldosAjuste.Fields!Precio = lnPrecioAjuste
                     oRsSaldosAjuste.Fields!Total = Round(lnSaldoGalenHos * lnPrecioAjuste, 2)
                     oRsSaldosAjuste.Fields!Lote = lcLote
                     oRsSaldosAjuste.Fields!FechaVencimiento = ldFvencimiento
                     lnTotalNI = lnTotalNI + oRsSaldosAjuste.Fields!Total
               End If
               oRsSaldosAjuste.MoveNext
            Loop
            'Genera Archivos Cabecera/Detalle de NI-Ajuste
            If lnTotalNI > 0 Then
                With mo_farmMovimiento
                    .DocumentoIdtipo = 10    'ajuste inventario
                    .DocumentoNumero = lcDocumento
                    .FechaCreacion = ldFechaProceso
                    .idAlmacenDestino = Val(cmbAlmacenA.BoundText)
                    .idAlmacenOrigen = 0
                    .idEstadoMovimiento = 1   'registrado
                    .idTipoConcepto = 20     'invenario
                    .idUsuario = ml_idUsuario
                    .IdUsuarioAuditoria = ml_idUsuario
                    .movTipo = "E"
                    .Observaciones = ""
                    .Total = lnTotalNI
                End With
                With mo_farmMovimientoNotaIngreso
                    .DocumentoFechaRecepcion = ldFechaProceso
                    .IdPaciente = 0
                    .IdComprobantePago = 0
                    .idProveedor = 0
                    .idTipoCompra = 0
                    .idTipoProceso = 0
                    .IdUsuarioAuditoria = ml_idUsuario
                    .movTipo = "E"
                    .NumeroProceso = ""
                    .OrigenFecha = 0
                    .OrigenIdTipo = 22
                    .OrigenNumero = ""
                    .IdCuentaAtencion = 0
                    .idFuenteFinanciamiento = 0
                End With
                oRsSaldosAjuste.Filter = "movTipo='E'"
                If chkSoloInventario.Value = 0 Then
                    If Not mo_ReglasFarmacia.AgregaDatosDeNotaIngreso(mo_farmMovimiento, mo_farmMovimientoNotaIngreso, oDoProveedores, oRsSaldosAjuste, 0, 1304, "Olidata") = True Then
                       MsgBox "Grabo mal Nota Ingreso por Ajuste" + Chr(13) + mo_ReglasFarmacia.MensajeError, vbCritical, Me.Caption
                    End If
                End If
            End If
            'Genera Archivos Cabecera/Detalle de NS-Ajuste
            If lnTotalNS > 0 Then
                With mo_farmMovimiento
                    .DocumentoIdtipo = 10    'ajuste inventario
                    .DocumentoNumero = lcDocumento
                    .FechaCreacion = ldFechaProceso
                    .idAlmacenDestino = 0
                    .idAlmacenOrigen = Val(cmbAlmacenA.BoundText)
                    .idEstadoMovimiento = 1   'registrado
                    .idTipoConcepto = 20     'invenario
                    .idUsuario = ml_idUsuario
                    .IdUsuarioAuditoria = ml_idUsuario
                    .movTipo = "S"
                    .Observaciones = ""
                    .Total = lnTotalNS
                End With
                oRsSaldosAjuste.Filter = "movTipo='S'"
                If chkSoloInventario.Value = 0 Then
                    If Not AgregaDatosDeNotaSalidaAI(mo_farmMovimiento, oRsSaldosAjuste, 1305, "Olidata") = True Then
                       MsgBox "Grabo mal Nota Salida por Ajuste" + Chr(13) + mo_ReglasFarmacia.MensajeError, vbCritical, Me.Caption
                    End If
                 End If
            End If
        End If
        'Graba el Documento registrado como INVENTARIO ANULADO, con la finalidad que se
        'pueda seguir ingresando datos la proxima vez
        lcSql = "select * from FarmInventario where idEstadoInventario=0 and NumeroInventario='" & txtDctoAjuste.Text & "' and idAlmacen=" & cmbAlmacenA.BoundText
        oRsTmp2.Open lcSql, wxConexionRed, adOpenKeyset, adLockOptimistic
        If oRsTmp2.RecordCount > 0 Then
           oRsTmp2.MoveFirst
           Do While Not oRsTmp2.EOF
              lcSql = "delete from FarmInventarioCabecera where idInventario=" & oRsTmp2.Fields!idInventario
              oRsTmp1.Open lcSql, wxConexionRed, adOpenKeyset, adLockOptimistic
              lcSql = "delete from FarmInventarioDetalle where idInventario=" & oRsTmp2.Fields!idInventario
              oRsTmp1.Open lcSql, wxConexionRed, adOpenKeyset, adLockOptimistic
              oRsTmp2.Delete
              oRsTmp2.Update
              oRsTmp2.MoveNext
           Loop
        End If
        oRsTmp2.Close
        With mo_FarmInventario
            .FechaCreacion = ldFechaProceso
            .numeroInventario = txtDctoAjuste.Text
            .idAlmacen = Val(cmbAlmacenA.BoundText)
            .idEstadoInventario = 0   'registrado
            .idUsuario = ml_idUsuario
            .IdUsuarioAuditoria = ml_idUsuario
        End With
        oRsSaldosAjuste.Filter = ""
        oRsSaldosAjuste.MoveFirst
        Do While Not oRsSaldosAjuste.EOF
           oRsSaldosAjuste.Fields!Cantidad = oRsSaldosAjuste.Fields!saldo
           oRsSaldosAjuste.Update
           oRsSaldosAjuste.MoveNext
        Loop
        If Not mo_ReglasFarmacia.AgregaDatosDeInventario(mo_FarmInventario, oRsSaldosAjuste, oRsSaldosAjuste, 801, "Olidata", False) = True Then
           MsgBox "Grabo mal el DOCUMENTO", vbCritical, Me.Caption
        End If
       
       '
       Unload Me
    End If
    Exit Sub
ErrSald:
    MsgBox Err.Description
    'Resume
End Sub



Private Sub cmdSaldosT_Click()
  If Val(txtSaldoT.Text) > -1 Then
    oRsSaldosAjuste.MoveFirst
    Do While Not oRsSaldosAjuste.EOF
       oRsSaldosAjuste.Fields!saldo = Val(txtSaldoT.Text)
       oRsSaldosAjuste.Update
       oRsSaldosAjuste.MoveNext
    Loop
  End If
End Sub

Private Sub Command1_Click()
    On Error GoTo ErrDel
    oRsSaldosAjuste.Delete
    oRsSaldosAjuste.Update
ErrDel:
End Sub

Private Sub Command2_Click()
    If MsgBox("Elimina toda la Lista de Medicamentos?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
       LimpiaLista
    End If
End Sub
Sub LimpiaLista()
        If oRsSaldosAjuste.RecordCount > 0 Then
          oRsSaldosAjuste.MoveFirst
          Do While Not oRsSaldosAjuste.EOF
             oRsSaldosAjuste.Delete
             oRsSaldosAjuste.Update
             oRsSaldosAjuste.MoveNext
          Loop
        End If

End Sub



Private Sub Command4_Click()
    Dim lcCodigoS As String
    Dim oRsTmp1 As New Recordset
    lcCodigoS = InputBox("ingrese Codigo del Producto (SISMED): ")
    lcSql = "select * from FactCatalogoBienesInsumos where codigo='" & lcCodigoS & "'"
    oRsTmp1.Open lcSql, wxConexionRed, adOpenKeyset, adLockOptimistic
    If oRsTmp1.RecordCount > 0 Then
       If oRsSaldosAjuste.RecordCount > 0 Then
            oRsSaldosAjuste.MoveFirst
            oRsSaldosAjuste.Find "idProducto=" & oRsTmp1.Fields!idProducto
       End If
       If oRsSaldosAjuste.EOF Then
          oRsSaldosAjuste.AddNew
          oRsSaldosAjuste.Fields!idProducto = oRsTmp1.Fields!idProducto
          oRsSaldosAjuste.Fields!codigo = oRsTmp1.Fields!codigo
          oRsSaldosAjuste.Fields!medicamento = oRsTmp1.Fields!nombre
          oRsSaldosAjuste.Fields!saldo = 0
          oRsSaldosAjuste.Fields!RegistroSanitario = "."
          oRsSaldosAjuste.Update
       Else
          MsgBox "Ese Código ya existe", vbCritical, Me.Caption
       End If
    End If
    Set oRsTmp1 = Nothing
End Sub


Private Sub Form_Load()
    wxConexionRed = sighEntidades.CadenaConexion
    
   
    
    oRsAlmacenes1.Open "select * from farmAlmacen where idTipoLocales<>'X' and idEstado=1 order by descripcion", wxConexionRed, adOpenKeyset, adLockOptimistic
    Set cmbAlmacenA.RowSource = oRsAlmacenes1
    cmbAlmacenA.ListField = "descripcion"
    cmbAlmacenA.BoundColumn = "idAlmacen"
    '
    '
    GenerarRecordsetTemporal
End Sub


Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    oConexionFox.Close
    oRsAlmacenes.Close
    oRsAlmacenes1.Close
    oRsUsuario.Close
    oRsSaldosAjuste.Close
    oRsFarmSaldoDetallado.Close
End Sub


Sub GenerarRecordsetTemporal()
    With oRsSaldosAjuste
          .Fields.Append "IdProducto", adInteger
          .Fields.Append "Codigo", adVarChar, 20, adFldIsNullable
          .Fields.Append "Medicamento", adVarChar, 150, adFldIsNullable
          .Fields.Append "Saldo", adInteger
          .Fields.Append "Lote", adChar, 15
          .Fields.Append "FechaVencimiento", adDate, , adFldIsNullable
          .Fields.Append "Cantidad", adInteger
          .Fields.Append "Precio", adDouble
          .Fields.Append "Total", adDouble
          .Fields.Append "MovNumeroS", adChar, 9, adFldIsNullable
          .Fields.Append "MovTipo", adVarChar, 1, adFldIsNullable
          .Fields.Append "RegistroSanitario", adVarChar, 10, adFldIsNullable
          .LockType = adLockOptimistic
          .Open
    End With
    Set grdSaldosAjuste.DataSource = oRsSaldosAjuste
    '
    With oRsFarmSaldoDetallado
          .Fields.Append "IdProducto", adInteger
          .Fields.Append "Codigo", adVarChar, 20, adFldIsNullable
          .Fields.Append "nombre", adVarChar, 150, adFldIsNullable
          .Fields.Append "Cantidad", adInteger
          .Fields.Append "Lote", adChar, 15
          .Fields.Append "FechaVencimiento", adDate, , adFldIsNullable
          .Fields.Append "FechaVencimientoN", adDate, , adFldIsNullable
          .LockType = adLockOptimistic
          .Open
    End With
    Set grdFarmSaldoDetallado.DataSource = oRsFarmSaldoDetallado
End Sub


Private Sub grdSaldosAjuste_HeadClick(ByVal ColIndex As Integer)
    Select Case ColIndex
    Case 0   'ordenado por codigo
         oRsSaldosAjuste.Sort = "codigo asc"
         Label1.Caption = "Por Codigo"
    Case 1   'ordenado por descripcion
         oRsSaldosAjuste.Sort = "medicamento asc"
         Label1.Caption = "Por Descripción"
    End Select
End Sub




Function farmRegeneraSaldos() As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set farmRegeneraSaldos = Nothing
    ms_MensajeError = ""
    oConexion.Open wxConexionRed
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmRegeneraSaldos"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set farmRegeneraSaldos = oRecordset
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


Function ActualizaSaldosPorProducto(lcEntradaOsalida As String, lnIdAlmacen As Long, lnIdProducto As Long, lcLote As String, ldFechaVencimiento As Date, lnCantidad As Long, lnPrecio As Double) As Boolean
    On Error GoTo ManejadorDeError
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    ActualizaSaldosPorProducto = False
    oConexion.Open wxConexionRed
    oConexion.CursorLocation = adUseClient
    With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = oConexion
       .CommandText = "FarmActualizaSaldosPorProducto"
       Set oParameter = .CreateParameter("@lcEntradaSalida", adVarChar, adParamInput, 1, lcEntradaOsalida): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdAlmacen", adInteger, adParamInput, 0, lnIdAlmacen): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@Lote", adVarChar, adParamInput, 15, lcLote): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@FechaVencimiento", adDate, adParamInput, 10, ldFechaVencimiento): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@Cantidad", adInteger, adParamInput, 0, lnCantidad): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@Precio", adDouble, adParamInput, 0, lnPrecio): .Parameters.Append oParameter
       .Execute
    End With
    ActualizaSaldosPorProducto = True
    Exit Function
ManejadorDeError:

End Function



Function AgregaDatosDeNotaSalidaAI(oDoMovimiento As DoFarmMovimiento, oRsDetalleProductos As ADODB.Recordset, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oMovimiento As New farmMovimiento
    Dim oMovimientoDetalle As New farmMovimientoDetalle
    Dim oDoMovimientoDetalle As New DoFarmMovimientoDetalle
    Dim lcCorrelativo As String
    Dim lnItem As Long
    Dim bProcesoOK As Boolean
    oConexion.Open sighEntidades.CadenaConexion
    oConexion.BeginTrans
    bProcesoOK = True
    Set oMovimiento.Conexion = oConexion
    Set oMovimientoDetalle.Conexion = oConexion
    '*********  graba tabla correlativos farmTipoDocumentos  ***************
    lcCorrelativo = oMovimiento.DevuelveYactualizaCorrelativosDeDocumentosES(2)
    '*********  graba tabla Movimientos  ***************
    With oDoMovimiento
       .movNumero = lcCorrelativo
    End With
    
    If Not oMovimiento.Insertar(oDoMovimiento) Then
            bProcesoOK = False: mo_mensajeError = oMovimiento.MensajeError: GoTo TerminarNS
    End If
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoMovimiento.IdUsuarioAuditoria, "A", 0, "FarmMovimiento/" & oDoMovimiento.movTipo & "/" & oDoMovimiento.movNumero, oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "")            'ListBarItems.idListItem
    '*********  graba tabla farmMovimientosDetalle,farmSaldo,farmSaldoDetalle  ***************
    
    oDoMovimientoDetalle.IdUsuarioAuditoria = oDoMovimiento.IdUsuarioAuditoria
    oDoMovimientoDetalle.movNumero = oDoMovimiento.movNumero
    oDoMovimientoDetalle.movTipo = oDoMovimiento.movTipo
    lnItem = 1
    oRsDetalleProductos.MoveFirst
    Do While Not oRsDetalleProductos.EOF
       oDoMovimientoDetalle.Cantidad = oRsDetalleProductos.Fields!Cantidad
       oDoMovimientoDetalle.FechaVencimiento = oRsDetalleProductos.Fields!FechaVencimiento
       oDoMovimientoDetalle.idProducto = oRsDetalleProductos.Fields!idProducto
       oDoMovimientoDetalle.Item = lnItem
       oDoMovimientoDetalle.Lote = oRsDetalleProductos.Fields!Lote
       oDoMovimientoDetalle.Precio = oRsDetalleProductos.Fields!Precio
       oDoMovimientoDetalle.RegistroSanitario = ""
       oDoMovimientoDetalle.Total = oRsDetalleProductos.Fields!Total
       If Not oMovimientoDetalle.Insertar(oDoMovimientoDetalle) Then
                bProcesoOK = False: mo_mensajeError = oMovimientoDetalle.MensajeError: GoTo TerminarNS
       End If
       lnItem = lnItem + 1
       oRsDetalleProductos.MoveNext
    Loop
TerminarNS:
    If bProcesoOK Then
        AgregaDatosDeNotaSalidaAI = True
        oConexion.CommitTrans
    Else
        AgregaDatosDeNotaSalidaAI = False
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oMovimiento = Nothing
    Set oMovimientoDetalle = Nothing
    Set oDoMovimientoDetalle = Nothing
End Function


Private Sub txtBuscaMed_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And txtBuscaMed.Text <> "" Then
        Dim lnLen As Integer
        lnLen = Len(Trim(txtBuscaMed.Text))
        oRsSaldosAjuste.MoveFirst
        Do While Not oRsSaldosAjuste.EOF
           If UCase(Right(Label1.Caption, 6)) = "CODIGO" Then
                If UCase(Left(txtBuscaMed.Text, lnLen)) = UCase(Left(oRsSaldosAjuste.Fields!codigo, lnLen)) Then
                   Exit Do
                End If
           Else
                If UCase(Left(txtBuscaMed.Text, lnLen)) = UCase(Left(oRsSaldosAjuste.Fields!medicamento, lnLen)) Then
                   Exit Do
                End If
           End If
           oRsSaldosAjuste.MoveNext
        Loop
        If oRsSaldosAjuste.EOF Then
           oRsSaldosAjuste.MoveFirst
        End If
    End If
End Sub



Private Sub txtCodigoSismed_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And txtCodigoSismed.Text <> "" Then
       Dim oRsTmp As New ADODB.Recordset
       lcSql = "SELECT     dbo.FactCatalogoBienesInsumos.Codigo, dbo.FactCatalogoBienesInsumos.Nombre, dbo.farmSaldoDetallado.*" & _
               " FROM         dbo.farmSaldoDetallado LEFT OUTER JOIN" & _
               "       dbo.FactCatalogoBienesInsumos ON dbo.farmSaldoDetallado.idProducto = dbo.FactCatalogoBienesInsumos.IdProducto" & _
               " where dbo.FactCatalogoBienesInsumos.Codigo='" & txtCodigoSismed.Text & "'"
       oRsTmp.Open lcSql, sighEntidades.CadenaConexion, adOpenKeyset, adLockOptimistic
       If oRsFarmSaldoDetallado.RecordCount > 0 Then
          oRsFarmSaldoDetallado.MoveFirst
          Do While Not oRsFarmSaldoDetallado.EOF
             oRsFarmSaldoDetallado.Delete
             oRsFarmSaldoDetallado.Update
             oRsFarmSaldoDetallado.MoveNext
          Loop
       End If
       If oRsTmp.RecordCount > 0 Then
          oRsTmp.MoveFirst
          Do While Not oRsTmp.EOF
             oRsFarmSaldoDetallado.AddNew
             oRsFarmSaldoDetallado.Fields!idProducto = oRsTmp.Fields!idProducto
             oRsFarmSaldoDetallado.Fields!codigo = oRsTmp.Fields!codigo
             oRsFarmSaldoDetallado.Fields!nombre = oRsTmp.Fields!nombre
             oRsFarmSaldoDetallado.Fields!Cantidad = oRsTmp.Fields!Cantidad
             oRsFarmSaldoDetallado.Fields!Lote = oRsTmp.Fields!Lote
             oRsFarmSaldoDetallado.Fields!FechaVencimiento = oRsTmp.Fields!FechaVencimiento
             oRsTmp.Update
             oRsTmp.MoveNext
          Loop
          Set grdFarmSaldoDetallado.DataSource = oRsFarmSaldoDetallado
       Else
          MsgBox "No hay datos para ese CODIGO"
          Set grdFarmSaldoDetallado.DataSource = Nothing
       End If
    End If
End Sub
