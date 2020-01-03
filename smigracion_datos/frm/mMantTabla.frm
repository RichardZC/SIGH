VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form mMantTabla 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de la tabla: LOLCLI"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   18030
   Icon            =   "mMantTabla.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   18030
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab 
      Height          =   7665
      Left            =   15
      TabIndex        =   1
      Top             =   -15
      Width           =   17970
      _ExtentX        =   31697
      _ExtentY        =   13520
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Varios"
      TabPicture(0)   =   "mMantTabla.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame23"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Actualiza ICI"
      TabPicture(1)   =   "mMantTabla.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame"
      Tab(1).Control(1)=   "grdICI"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame2 
         BackColor       =   &H80000002&
         Caption         =   "Lee Archivo Excel y graba nuevos Dx en Galenhos"
         ForeColor       =   &H000000FF&
         Height          =   2385
         Left            =   6720
         TabIndex        =   31
         Top             =   2895
         Width           =   6165
         Begin VB.CommandButton cmdProcesaDx 
            Caption         =   "Procesar"
            Height          =   405
            Left            =   165
            TabIndex        =   34
            Top             =   1800
            Width           =   5895
         End
         Begin VB.TextBox txtDx1 
            Height          =   315
            Left            =   1320
            TabIndex        =   33
            Text            =   "c:\dx.xls"
            Top             =   210
            Width           =   4725
         End
         Begin VB.TextBox Text1 
            Height          =   1200
            Left            =   240
            MultiLine       =   -1  'True
            TabIndex        =   32
            Text            =   "mMantTabla.frx":047A
            Top             =   525
            Width           =   5835
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Archivo Excel:"
            Height          =   285
            Left            =   210
            TabIndex        =   35
            Top             =   270
            Width           =   1245
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H80000009&
         Caption         =   "Lee Archivo Excel y graba nuevos CPT en Galenhos"
         ForeColor       =   &H000000FF&
         Height          =   2385
         Left            =   6720
         TabIndex        =   26
         Top             =   480
         Width           =   6165
         Begin VB.TextBox Text6 
            Height          =   1200
            Left            =   255
            MultiLine       =   -1  'True
            TabIndex        =   29
            Text            =   "mMantTabla.frx":056B
            Top             =   525
            Width           =   5835
         End
         Begin VB.TextBox txtCpt1 
            Height          =   315
            Left            =   1320
            TabIndex        =   28
            Text            =   "c:\cpt.xls"
            Top             =   210
            Width           =   4725
         End
         Begin VB.CommandButton cmdAgregaCpt 
            Caption         =   "Procesar"
            Height          =   405
            Left            =   180
            TabIndex        =   27
            Top             =   1800
            Width           =   5895
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Archivo Excel:"
            Height          =   195
            Left            =   210
            TabIndex        =   30
            Top             =   270
            Width           =   1020
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Filtro"
         Height          =   705
         Left            =   -74925
         TabIndex        =   13
         Top             =   435
         Width           =   17805
         Begin VB.CheckBox chkConLotes 
            Caption         =   "ICI solo LOTES/F.VENCIMIENTO"
            Height          =   270
            Left            =   11160
            TabIndex        =   25
            Top             =   300
            Width           =   2790
         End
         Begin VB.CommandButton cmdLimpiar 
            Caption         =   "Limpiar"
            Height          =   300
            Left            =   16455
            TabIndex        =   24
            Top             =   255
            Width           =   1200
         End
         Begin VB.CheckBox chkSismed 
            Caption         =   "Sismed"
            Height          =   255
            Left            =   9810
            TabIndex        =   23
            Top             =   300
            Value           =   1  'Checked
            Width           =   900
         End
         Begin VB.CommandButton cmdFiltro 
            Caption         =   "Buscar"
            Height          =   300
            Left            =   15135
            TabIndex        =   22
            Top             =   270
            Width           =   1200
         End
         Begin VB.TextBox txtCodigo1 
            Height          =   285
            Left            =   7740
            TabIndex        =   21
            Top             =   270
            Width           =   1410
         End
         Begin VB.TextBox txtAnio1 
            Height          =   285
            Left            =   5685
            TabIndex        =   19
            Top             =   270
            Width           =   585
         End
         Begin VB.TextBox txtMes1 
            Height          =   285
            Left            =   4260
            TabIndex        =   17
            Top             =   270
            Width           =   585
         End
         Begin VB.TextBox txtFarmacia1 
            Height          =   285
            Left            =   945
            TabIndex        =   15
            Top             =   270
            Width           =   2655
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Código Item"
            Height          =   195
            Left            =   6840
            TabIndex        =   20
            Top             =   330
            Width           =   840
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Año"
            Height          =   195
            Left            =   5355
            TabIndex        =   18
            Top             =   330
            Width           =   285
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Mes"
            Height          =   195
            Left            =   3930
            TabIndex        =   16
            Top             =   330
            Width           =   300
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Farmacia"
            Height          =   195
            Left            =   255
            TabIndex        =   14
            Top             =   330
            Width           =   645
         End
      End
      Begin VB.Frame Frame23 
         BackColor       =   &H8000000D&
         Caption         =   "Lee Archivo Excel y graba datos en Galenhos (CPT corto)"
         ForeColor       =   &H000000FF&
         Height          =   4215
         Left            =   165
         TabIndex        =   2
         Top             =   465
         Width           =   6165
         Begin VB.CheckBox chkSoloSisSoat 
            Caption         =   "solo actualiza SIS y SOAT"
            Height          =   285
            Left            =   240
            TabIndex        =   10
            Top             =   3120
            Value           =   1  'Checked
            Width           =   3015
         End
         Begin VB.CheckBox chkNohallado 
            Caption         =   "marca en el EXCEL como NO HALLADO"
            Height          =   285
            Left            =   2520
            TabIndex        =   9
            Top             =   2640
            Value           =   1  'Checked
            Width           =   3495
         End
         Begin VB.TextBox txtCodigoESSALUD 
            Height          =   345
            Left            =   2010
            TabIndex        =   8
            Text            =   "17"
            Top             =   2190
            Width           =   4035
         End
         Begin VB.TextBox Text4 
            Height          =   345
            Left            =   210
            TabIndex        =   7
            Text            =   "g->Precio ESSALUD........"
            Top             =   2190
            Width           =   1815
         End
         Begin VB.CheckBox chkSoloActualiza 
            Caption         =   "Agrega nuevos CPT"
            Height          =   285
            Left            =   210
            TabIndex        =   6
            Top             =   2640
            Value           =   1  'Checked
            Width           =   2055
         End
         Begin VB.CommandButton cmdGrabaDescripcionCortaCPT 
            Caption         =   "Procesar"
            Height          =   405
            Left            =   120
            TabIndex        =   5
            Top             =   3600
            Width           =   5895
         End
         Begin VB.TextBox txtExcel1 
            Height          =   315
            Left            =   1320
            TabIndex        =   4
            Text            =   "c:\cpt.xls"
            Top             =   210
            Width           =   4725
         End
         Begin VB.TextBox Text3 
            Height          =   1665
            Left            =   210
            MultiLine       =   -1  'True
            TabIndex        =   3
            Text            =   "mMantTabla.frx":0652
            Top             =   540
            Width           =   5835
         End
         Begin VB.Label Label28 
            Caption         =   "Archivo Excel:"
            Height          =   285
            Left            =   210
            TabIndex        =   11
            Top             =   270
            Width           =   1245
         End
      End
      Begin MSDataGridLib.DataGrid grdICI 
         Height          =   6315
         Left            =   -74895
         TabIndex        =   12
         Top             =   1245
         Width           =   17805
         _ExtentX        =   31406
         _ExtentY        =   11139
         _Version        =   393216
         Enabled         =   -1  'True
         HeadLines       =   1
         RowHeight       =   18
         AllowAddNew     =   -1  'True
         AllowDelete     =   -1  'True
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
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Formato ICI"
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
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
            DataField       =   ""
            Caption         =   ""
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   285
      Left            =   60
      TabIndex        =   0
      Top             =   7725
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   503
      _Version        =   327682
      Appearance      =   1
   End
End
Attribute VB_Name = "mMantTabla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oRsICI As New Recordset
Dim lcBuscaParametro As New SIGHDatos.Parametros

Private Sub cmdAgregaCpt_Click()
    Dim lnPrecioSIS As Double, lnPrecioSOAT As Double, lnPrecioConvenio As Double, lnPrecioESSSALUD As Double
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    Dim EXL As Excel.Application
    Set EXL = New Excel.Application
    Dim W As Excel.Workbook
    Set W = EXL.Workbooks.Open(txtCpt1.Text)
    Dim s As Excel.Worksheet
    Set s = W.Sheets("Hoja1")
    Dim lnFor As Integer, lnFila As Integer, lcRango As String, lnFilaFinal As Integer, oRsTmp As New Recordset, lnIdCpt As Long, lcSql As String, lcCodigo As String
    Dim lcCPTcorta As String, lnPrecioPagante As Double, lnIdProducto As Long, lbContinuar As Boolean
    Dim lbCont2 As Boolean
    lnFila = 1
    lnFilaFinal = 30000
    Me.ProgressBar1.Min = lnFila
    Me.ProgressBar1.Max = lnFilaFinal
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open SIGHEntidades.CadenaConexion
  
    For lnFor = lnFila To lnFilaFinal
        DoEvents: Me.ProgressBar1.Value = lnFor: Me.Refresh
        lbCont2 = True
        lcRango = "A" + Trim(Str(lnFor))
        lcCodigo = Trim(s.Range(lcRango).Value)
        lcRango = "B" + Trim(Str(lnFor))
        lcCPTcorta = Trim(s.Range(lcRango).Value)
        If Len(Trim(lcCodigo)) > 0 And Len(lcCodigo) < 21 And Trim(lcCPTcorta) <> "" Then
            If oRsTmp.State = 1 Then
               oRsTmp.Close
            End If
            
            With oCommand
                .CommandType = adCmdStoredProc
                Set .ActiveConnection = oConexion
                .CommandTimeout = 150
                .CommandText = "FactCatalogoServiciosXcodigo"
                Set oParameter = .CreateParameter("@lcCodigo", adVarChar, adParamInput, 20, lcCodigo): .Parameters.Append oParameter
                Set oRsTmp = .Execute
                Set oRsTmp.ActiveConnection = Nothing
            End With
            Set oCommand = Nothing
            Set oParameter = Nothing
            
            If oRsTmp.RecordCount = 0 Then
                lcRango = "C" + Trim(Str(lnFor))
                s.Range(lcRango).Value = "*"
                With oCommand
                    .CommandType = adCmdStoredProc
                    Set .ActiveConnection = oConexion
                    .CommandTimeout = 150
                    .CommandText = "FactCatalogoServiciosAgregarInformacion"
                    Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamOutput, 0, 1): .Parameters.Append oParameter
                    Set oParameter = .CreateParameter("@Codigo", adVarChar, adParamInput, 20, lcCodigo): .Parameters.Append oParameter
                    Set oParameter = .CreateParameter("@Nombre", adVarChar, adParamInput, 250, Left(lcCPTcorta, 250)): .Parameters.Append oParameter
                    Set oParameter = .CreateParameter("@IdServicioGrupo", adInteger, adParamInput, 0, 5): .Parameters.Append oParameter
                    Set oParameter = .CreateParameter("@IdServicioSubGrupo", adInteger, adParamInput, 0, 24): .Parameters.Append oParameter
                    Set oParameter = .CreateParameter("@IdServicioSeccion", adInteger, adParamInput, 0, 78): .Parameters.Append oParameter
                    Set oParameter = .CreateParameter("@EsCPT", adInteger, adParamInput, 0, 1): .Parameters.Append oParameter
                    Set oParameter = .CreateParameter("@NombreMINSA", adVarChar, adParamInput, 250, Left(lcCPTcorta, 250)): .Parameters.Append oParameter
                    .Execute
                    lnIdProducto = .Parameters("@IdProducto")
                End With
                Set oCommand = Nothing
                Set oParameter = Nothing
            End If
        End If
    Next
    Set s = Nothing
    W.Save
    W.Close
    Set W = Nothing
    Set EXL = Nothing
    oConexion.Close
    Set oConexion = Nothing
    Unload Me

End Sub

Private Sub cmdFiltro_Click()
    On Error GoTo eRRCarga2
    If txtFarmacia1.Text = "" Then
       MsgBox "Ingrese el CODIGO DE FARMACIA"
       Exit Sub
    ElseIf Me.txtAnio1.Text = "" Then
       MsgBox "ingrese el AÑO"
       Exit Sub
    ElseIf Me.txtMes1.Text = "" Then
       MsgBox "ingrese el MES"
       Exit Sub
    End If
    If Len(Me.txtMes1.Text) = 1 Then
       Me.txtMes1.Text = "0" & Me.txtMes1.Text
    End If
    
    Dim txt As String
    If Me.chkConLotes.Value = 1 Then
    txt = "select * from farm_formdetL where CODIGO_PRE='" & txtFarmacia1.Text & _
        "' and TIPSUM='" & IIf(Me.chkSismed.Value = 1, "S", "D") & _
        "' and ANNOMES='" & Trim(Me.txtAnio1.Text) & Trim(Me.txtMes1.Text) & _
        "'" & IIf(Me.txtCodigo1.Text = "", "", " and CODIGO_MED='" & Me.txtCodigo1.Text & "'")
    Else
    txt = "select * from farm_formdet where CODIGO_PRE='" & txtFarmacia1.Text & _
        "' and TIPSUM='" & IIf(Me.chkSismed.Value = 1, "S", "D") & _
        "' and ANNOMES='" & Trim(Me.txtAnio1.Text) & Trim(Me.txtMes1.Text) & _
        "'" & IIf(Me.txtCodigo1.Text = "", "", " and CODIGO_MED='" & Me.txtCodigo1.Text & "'")
    End If
    If oRsICI.State = 1 Then oRsICI.Close
    oRsICI.Open txt, SIGHEntidades.CadenaConexionShape, adOpenKeyset, adLockOptimistic
    Set grdICI.DataSource = oRsICI
    Exit Sub
eRRCarga2:
    If Err.Number = 3705 Then
       oRsICI.Close
       Resume
    Else
       MsgBox Err.Description
    End If

End Sub

Private Sub cmdGrabaDescripcionCortaCPT_Click()
    Dim lnPrecioSIS As Double, lnPrecioSOAT As Double, lnPrecioConvenio As Double, lnPrecioESSSALUD As Double
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    Dim EXL As Excel.Application
    Set EXL = New Excel.Application
    Dim W As Excel.Workbook
    Set W = EXL.Workbooks.Open(txtExcel1.Text)
    Dim s As Excel.Worksheet
    Set s = W.Sheets("Hoja1")
    Dim lnFor As Integer, lnFila As Integer, lcRango As String, lnFilaFinal As Integer, oRsTmp As New Recordset, lnIdCpt As Long, lcSql As String, lcCodigo As String
    Dim lcCPTcorta As String, lnPrecioPagante As Double, lnIdProducto As Long, lbContinuar As Boolean
    Dim lbCont2 As Boolean
    lnFila = 1
    lnFilaFinal = 10000
    Me.ProgressBar1.Min = lnFila
    Me.ProgressBar1.Max = lnFilaFinal
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open SIGHEntidades.CadenaConexion
  
    For lnFor = lnFila To lnFilaFinal
        DoEvents: Me.ProgressBar1.Value = lnFor: Me.Refresh
        lbCont2 = True
        lcRango = "A" + Trim(Str(lnFor))
        lcCodigo = Trim(s.Range(lcRango).Value)
        lcRango = "B" + Trim(Str(lnFor))
        lcCPTcorta = Trim(s.Range(lcRango).Value)
        lcRango = "C" + Trim(Str(lnFor))
        lnPrecioPagante = Val(s.Range(lcRango).Value)
        lcRango = "D" + Trim(Str(lnFor))
        lnPrecioSIS = Val(s.Range(lcRango).Value)
        lcRango = "E" + Trim(Str(lnFor))
        lnPrecioSOAT = Val(s.Range(lcRango).Value)
        lcRango = "F" + Trim(Str(lnFor))
        lnPrecioConvenio = Val(s.Range(lcRango).Value)
        lcRango = "G" + Trim(Str(lnFor))
        lnPrecioESSSALUD = Val(s.Range(lcRango).Value)
        If Len(Trim(lcCodigo)) > 0 And Len(lcCodigo) < 8 And Trim(lcCPTcorta) <> "" Then
            If oRsTmp.State = 1 Then
               oRsTmp.Close
            End If
            
            With oCommand
                .CommandType = adCmdStoredProc
                Set .ActiveConnection = oConexion
                .CommandTimeout = 150
                .CommandText = "FactCatalogoServiciosXcodigo"
                Set oParameter = .CreateParameter("@lcCodigo", adVarChar, adParamInput, 20, lcCodigo): .Parameters.Append oParameter
                Set oRsTmp = .Execute
                Set oRsTmp.ActiveConnection = Nothing
            End With
            Set oCommand = Nothing
            Set oParameter = Nothing
            
            lbContinuar = True
            If chkSoloActualiza.Value = 0 And oRsTmp.RecordCount = 0 Then
               lbContinuar = False
            End If

            
            
            If lbContinuar = True Then
            
                lbCont2 = True
                If chkSoloSisSoat.Value = 1 Then
                   lbCont2 = False
                End If
            
                If oRsTmp.RecordCount = 0 Then
                    If chkNohallado.Value = 1 Then
                       lcRango = "G" + Trim(Str(lnFor))
                       s.Range(lcRango).Value = "NO HALLADO"
                    End If
                    With oCommand
                        .CommandType = adCmdStoredProc
                        Set .ActiveConnection = oConexion
                        .CommandTimeout = 150
                        .CommandText = "FactCatalogoServiciosAgregarInformacion"
                        Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamOutput, 0, 1): .Parameters.Append oParameter
                        Set oParameter = .CreateParameter("@Codigo", adVarChar, adParamInput, 7, lcCodigo): .Parameters.Append oParameter
                        Set oParameter = .CreateParameter("@Nombre", adVarChar, adParamInput, 255, Left(lcCPTcorta, 255)): .Parameters.Append oParameter
                        Set oParameter = .CreateParameter("@IdServicioGrupo", adInteger, adParamInput, 0, 5): .Parameters.Append oParameter
                        Set oParameter = .CreateParameter("@IdServicioSubGrupo", adInteger, adParamInput, 0, 24): .Parameters.Append oParameter
                        Set oParameter = .CreateParameter("@IdServicioSeccion", adInteger, adParamInput, 0, 78): .Parameters.Append oParameter
                        Set oParameter = .CreateParameter("@EsCPT", adInteger, adParamInput, 0, 1): .Parameters.Append oParameter
                        Set oParameter = .CreateParameter("@NombreMINSA", adVarChar, adParamInput, 255, Left(lcCPTcorta, 255)): .Parameters.Append oParameter
                        .Execute
                        lnIdProducto = .Parameters("@IdProducto")
                    End With
                    Set oCommand = Nothing
                    Set oParameter = Nothing
                Else
                    With oCommand
                        .CommandType = adCmdStoredProc
                        Set .ActiveConnection = oConexion
                        .CommandTimeout = 150
                        .CommandText = "FactCatalogoServiciosActualizarInformacion"
                        Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, oRsTmp.Fields!idProducto): .Parameters.Append oParameter
                        Set oParameter = .CreateParameter("@Nombre", adVarChar, adParamInput, 255, Left(lcCPTcorta, 255)): .Parameters.Append oParameter
                        .Execute
                    End With
                    Set oCommand = Nothing
                    Set oParameter = Nothing
                    lnIdProducto = oRsTmp.Fields!idProducto
                End If
                
                oRsTmp.Close
                
                If lbCont2 = True Then                  'particular
               
                    With oCommand
                        .CommandType = adCmdStoredProc
                        Set .ActiveConnection = oConexion
                        .CommandTimeout = 150
                        .CommandText = "FactCatalogoServiciosHospActualizarInformacionPorIdTipoFinanciamientoIdProducto"
                        Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
                        Set oParameter = .CreateParameter("@IdTipoFinanciamiento", adInteger, adParamInput, 0, 1): .Parameters.Append oParameter
                        Set oParameter = .CreateParameter("@PrecioUnitario", adCurrency, adParamInput, 0, FormatCurrency(lnPrecioPagante, 2, vbTrue, vbTrue)): .Parameters.Append oParameter
                        .Execute
                    End With
                    Set oCommand = Nothing
                    Set oParameter = Nothing
                    
               End If
               ' If lnPrecioSIS > 0 Then
                
                    With oCommand
                        .CommandType = adCmdStoredProc
                        Set .ActiveConnection = oConexion
                        .CommandTimeout = 150
                        .CommandText = "FactCatalogoServiciosHospActualizarInformacionPorIdTipoFinanciamientoIdProductoSIS"
                        Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
                        Set oParameter = .CreateParameter("@IdTipoFinanciamiento", adInteger, adParamInput, 0, 2): .Parameters.Append oParameter
                        Set oParameter = .CreateParameter("@PrecioUnitario", adCurrency, adParamInput, 0, FormatCurrency(lnPrecioSIS, 2, vbTrue, vbTrue)): .Parameters.Append oParameter
                        .Execute
                    End With
                    Set oCommand = Nothing
                    Set oParameter = Nothing
                    
                'End If
                'If lnPrecioSOAT > 0 Then
                    With oCommand
                        .CommandType = adCmdStoredProc
                        Set .ActiveConnection = oConexion
                        .CommandTimeout = 150
                        .CommandText = "FactCatalogoServiciosHospActualizarInformacionPorIdTipoFinanciamientoIdProducto"
                        Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
                        Set oParameter = .CreateParameter("@IdTipoFinanciamiento", adInteger, adParamInput, 0, 3): .Parameters.Append oParameter
                        Set oParameter = .CreateParameter("@PrecioUnitario", adCurrency, adParamInput, 0, FormatCurrency(lnPrecioSOAT, 2, vbTrue, vbTrue)): .Parameters.Append oParameter
                        .Execute
                    End With
                    Set oCommand = Nothing
                    Set oParameter = Nothing
                    
               ' End If
               If lbCont2 = True Then       'convenio
                    With oCommand
                        .CommandType = adCmdStoredProc
                        Set .ActiveConnection = oConexion
                        .CommandTimeout = 150
                        .CommandText = "FactCatalogoServiciosHospActualizarInformacionPorIdTipoFinanciamientoIdProducto"
                        Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
                        Set oParameter = .CreateParameter("@IdTipoFinanciamiento", adInteger, adParamInput, 0, 4): .Parameters.Append oParameter
                        Set oParameter = .CreateParameter("@PrecioUnitario", adCurrency, adParamInput, 0, FormatCurrency(lnPrecioConvenio, 2, vbTrue, vbTrue)): .Parameters.Append oParameter
                        .Execute
                    End With
                    Set oCommand = Nothing
                    Set oParameter = Nothing
                    
               End If
               If lnPrecioESSSALUD > 0 And Val(txtCodigoESSALUD.Text) > 10 And lbCont2 = True Then
                    With oCommand
                        .CommandType = adCmdStoredProc
                        Set .ActiveConnection = oConexion
                        .CommandTimeout = 150
                        .CommandText = "FactCatalogoServiciosHospActualizarInformacionPorIdTipoFinanciamientoIdProducto"
                        Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
                        Set oParameter = .CreateParameter("@IdTipoFinanciamiento", adInteger, adParamInput, 0, CLng(txtCodigoESSALUD.Text)): .Parameters.Append oParameter
                        Set oParameter = .CreateParameter("@PrecioUnitario", adCurrency, adParamInput, 0, FormatCurrency(lnPrecioESSSALUD, 2, vbTrue, vbTrue)): .Parameters.Append oParameter
                        .Execute
                    End With
                    Set oCommand = Nothing
                    Set oParameter = Nothing
                    
                End If
            End If
        End If
    Next
    Set s = Nothing
    'W.Save
    W.Close
    Set W = Nothing
    Set EXL = Nothing
    Unload Me

End Sub


Sub LimpiaFiltroICI()
     Me.txtAnio1.Text = Year(Date)
     Me.txtMes1.Text = Month(Date) - 1
     Me.txtFarmacia1.Text = lcBuscaParametro.SeleccionaFilaParametro(208) & "F01"
     chkConLotes.Value = 0
     txtCodigo1.Text = ""
     chkSismed.Value = 1
End Sub

Private Sub cmdLimpiar_Click()
    LimpiaFiltroICI
End Sub

Private Sub cmdProcesaDx_Click()
    Dim lnPrecioSIS As Double, lnPrecioSOAT As Double, lnPrecioConvenio As Double, lnPrecioESSSALUD As Double
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    Dim EXL As Excel.Application
    Set EXL = New Excel.Application
    Dim W As Excel.Workbook
    Set W = EXL.Workbooks.Open(txtDx1.Text)
    Dim s As Excel.Worksheet
    Set s = W.Sheets("Hoja1")
    Dim lnFor As Integer, lnFila As Integer, lcRango As String, lnFilaFinal As Integer, oRsTmp As New Recordset, lnIdCpt As Long, lcSql As String, lcCodigo As String
    Dim lcCPTcorta As String, lnPrecioPagante As Double, lnIdProducto As Long, lbContinuar As Boolean
    Dim lbCont2 As Boolean
    lnFila = 1
    lnFilaFinal = 30000
    Me.ProgressBar1.Min = lnFila
    Me.ProgressBar1.Max = lnFilaFinal
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open SIGHEntidades.CadenaConexion
  
    For lnFor = lnFila To lnFilaFinal
        DoEvents: Me.ProgressBar1.Value = lnFor: Me.Refresh
        lbCont2 = True
        lcRango = "A" + Trim(Str(lnFor))
        lcCodigo = Trim(s.Range(lcRango).Value)
        lcRango = "B" + Trim(Str(lnFor))
        lcCPTcorta = Trim(s.Range(lcRango).Value)
        If Len(Trim(lcCodigo)) > 0 And Len(lcCodigo) < 7 And Trim(lcCPTcorta) <> "" Then
            lcCodigo = Trim(Left(lcCodigo, 3) & "." & Mid(lcCodigo, 4, 10))
            If oRsTmp.State = 1 Then
               oRsTmp.Close
            End If
            
            With oCommand
                .CommandType = adCmdStoredProc
                Set .ActiveConnection = oConexion
                .CommandTimeout = 150
                .CommandText = "DiagnosticosSeleccionarTodoCamposPorCodigoCie2004"
                Set oParameter = .CreateParameter("@CodigoCie2004", adVarChar, adParamInput, 7, lcCodigo): .Parameters.Append oParameter
                Set oRsTmp = .Execute
                Set oRsTmp.ActiveConnection = Nothing
            End With
            Set oCommand = Nothing
            Set oParameter = Nothing
            
            If oRsTmp.RecordCount = 0 Then
                lcRango = "C" + Trim(Str(lnFor))
                s.Range(lcRango).Value = "*"
                With oCommand
                    .CommandType = adCmdStoredProc
                    Set .ActiveConnection = oConexion
                    .CommandTimeout = 150
                    .CommandText = "DiagnosticosAgregarPorCodigoDescripcionDatosCompletos"
                    Set oParameter = .CreateParameter("@CodigoCie2004", adVarChar, adParamInput, 7, lcCodigo): .Parameters.Append oParameter
                    Set oParameter = .CreateParameter("@Descripcion", adVarChar, adParamInput, 250, Left(lcCPTcorta, 250)): .Parameters.Append oParameter
                    Set oParameter = .CreateParameter("@EdadMaxDias", adInteger, adParamInput, 0, Null): .Parameters.Append oParameter
                    Set oParameter = .CreateParameter("@EdadMinDias", adInteger, adParamInput, 0, Null): .Parameters.Append oParameter
                    Set oParameter = .CreateParameter("@IdTipoSexo", adInteger, adParamInput, 0, Null): .Parameters.Append oParameter
                    Set oParameter = .CreateParameter("@FechaInicioVigencia", adDBTimeStamp, adParamInput, 0, CDate("31/12/2022")): .Parameters.Append oParameter 'Actualizado 23092014
                    Set oParameter = .CreateParameter("@EsActivo", adBoolean, adParamInput, 0, 1): .Parameters.Append oParameter
                    .Execute
                    'lnIdProducto = .Parameters("@IdDiagnostico")
                End With
                Set oCommand = Nothing
                Set oParameter = Nothing
            End If
        End If
    Next
    Set s = Nothing
    W.Save
    W.Close
    Set W = Nothing
    Set EXL = Nothing
    oConexion.Close
    Set oConexion = Nothing
    Unload Me

End Sub

Private Sub Form_Load()
     LimpiaFiltroICI
     LimpiaDBF
End Sub

Sub LimpiaDBF()
     oRsICI.Open "delete from farm_formdet where sit='9'", SIGHEntidades.CadenaConexionShape, adOpenKeyset, adLockOptimistic
     oRsICI.Open "delete from farm_formdetL where sit='9'", SIGHEntidades.CadenaConexionShape, adOpenKeyset, adLockOptimistic

End Sub
