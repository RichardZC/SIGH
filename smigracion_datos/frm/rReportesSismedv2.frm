VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form rReportesSismedv2 
   Caption         =   "Reportes desde el Formato ICI"
   ClientHeight    =   7155
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9840
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   9840
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame7 
      Height          =   1080
      Left            =   60
      TabIndex        =   21
      Top             =   6030
      Width           =   9735
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "rReportesSismedv2.frx":0000
         DownPicture     =   "rReportesSismedv2.frx":0460
         Height          =   700
         Left            =   3465
         Picture         =   "rReportesSismedv2.frx":08D5
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "rReportesSismedv2.frx":0D4A
         DownPicture     =   "rReportesSismedv2.frx":120E
         Height          =   700
         Left            =   5010
         Picture         =   "rReportesSismedv2.frx":16FA
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   225
         Width           =   1365
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Reportes e Indicadores"
      Height          =   2475
      Left            =   60
      TabIndex        =   6
      Top             =   3510
      Width           =   9735
      Begin VB.Frame Frame4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   390
         TabIndex        =   11
         Top             =   1710
         Width           =   8625
         Begin VB.OptionButton optMenor30 
            Caption         =   "Menor o igual a 30 días"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   180
            Value           =   -1  'True
            Width           =   2775
         End
         Begin VB.OptionButton optEntre31y60 
            Caption         =   "Entre 31 y 60 días"
            Height          =   255
            Left            =   3330
            TabIndex        =   13
            Top             =   180
            Width           =   2115
         End
         Begin VB.OptionButton optEntre61y90 
            Caption         =   "Entre 61 y 90 días"
            Height          =   255
            Left            =   6480
            TabIndex        =   12
            Top             =   180
            Width           =   1995
         End
      End
      Begin VB.OptionButton optValorStock 
         Caption         =   "Reporte por Valor de Stock en riesgo de expiración"
         Height          =   225
         Left            =   120
         TabIndex        =   10
         Top             =   1440
         Width           =   5385
      End
      Begin VB.OptionButton optDevolucionXvencimiento 
         Caption         =   "Reporte por 'Devolución x Vencimiento'"
         Height          =   225
         Left            =   120
         TabIndex        =   9
         Top             =   1065
         Width           =   4905
      End
      Begin VB.OptionButton optStockFinal 
         Caption         =   "Reporte por 'Stock Final'"
         Height          =   225
         Left            =   120
         TabIndex        =   8
         Top             =   675
         Width           =   5145
      End
      Begin VB.OptionButton optSalidaTotal 
         Caption         =   "Reporte por 'Salida Total'"
         Height          =   225
         Left            =   120
         TabIndex        =   7
         Top             =   300
         Value           =   -1  'True
         Width           =   4785
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Filtros para Reportes (solo se tomará datos de FARMACIA)"
      Height          =   1335
      Left            =   60
      TabIndex        =   2
      Top             =   2160
      Width           =   9735
      Begin VB.Frame Frame6 
         Caption         =   "Solo para Reportes"
         Height          =   1065
         Left            =   5730
         TabIndex        =   16
         Top             =   180
         Width           =   3855
         Begin VB.OptionButton optMuestraImporte 
            Caption         =   "Se muestra 'Importes'"
            Height          =   225
            Left            =   180
            TabIndex        =   18
            Top             =   660
            Width           =   3495
         End
         Begin VB.OptionButton optMuestraCantidad 
            Caption         =   "Se muestra 'Cantidades'"
            Height          =   225
            Left            =   150
            TabIndex        =   17
            Top             =   300
            Value           =   -1  'True
            Width           =   3525
         End
      End
      Begin MSDataListLib.DataCombo cmdPrograma 
         Height          =   330
         Left            =   990
         TabIndex        =   15
         Top             =   330
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   582
         _Version        =   393216
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txtAnio 
         Height          =   315
         Left            =   990
         TabIndex        =   5
         Text            =   "2009"
         Top             =   720
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Año"
         Height          =   210
         Left            =   150
         TabIndex        =   4
         Top             =   780
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Programa"
         Height          =   210
         Left            =   150
         TabIndex        =   3
         Top             =   390
         Width           =   765
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Consideraciones:"
      Height          =   2025
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   9735
      Begin VB.OptionButton optGalenhos 
         Caption         =   "Usando datos desde GalenHos"
         Height          =   255
         Left            =   5250
         TabIndex        =   20
         Top             =   1650
         Width           =   2985
      End
      Begin VB.OptionButton optSismedv2 
         Caption         =   "Usando datos desde el Sistema SISMEDV2"
         Height          =   255
         Left            =   180
         TabIndex        =   19
         Top             =   1650
         Value           =   -1  'True
         Width           =   4275
      End
      Begin VB.TextBox Text1 
         Height          =   1305
         Left            =   150
         MultiLine       =   -1  'True
         TabIndex        =   1
         Text            =   "rReportesSismedv2.frx":1BE6
         Top             =   270
         Width           =   9465
      End
   End
End
Attribute VB_Name = "rReportesSismedv2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Reportes Sismed
'        Programado por: Barrantes D
'        Fecha: Enero 2010
'
'------------------------------------------------------------------------------------
Option Explicit
Dim oRsFoxPrograma As New Recordset
Dim oRsFox1 As New Recordset
Dim oRsFox2 As New Recordset
Dim oRsFox3 As New Recordset
Dim oConexionFox As New Connection
Dim lcSql As String
Dim lcBuscaParametro As New SIGHDatos.Parametros


Private Sub btnAceptar_Click()
    If Val(cmdPrograma.BoundText) = 0 Then
       MsgBox "Elija el Programa", vbCritical, Me.Caption
       Exit Sub
    End If
    
    Dim oRsCubo As New Recordset
    Dim lbNuevo As Boolean, lcDescripcion As String
    Dim lcPrograma As String, lcSubPrograma As String
    Dim lbStockEnRiesgo As Boolean, lnDiasRiesgo As Integer
    Dim lcTitulo As String, lnSalidasIS As Long
    lcPrograma = Left(Me.cmdPrograma.BoundText, 2)
    lcSubPrograma = Right(Me.cmdPrograma.BoundText, 2)
    '
    With oRsCubo
      .Fields.Append "Codigo", adVarChar, 7
      .Fields.Append "Producto", adVarChar, 300
      .Fields.Append "Precio", adDouble
      .Fields.Append "ValorEne", adDouble
      .Fields.Append "ValorFeb", adDouble
      .Fields.Append "ValorMar", adDouble
      .Fields.Append "ValorAbr", adDouble
      .Fields.Append "ValorMay", adDouble
      .Fields.Append "ValorJun", adDouble
      .Fields.Append "ValorJul", adDouble
      .Fields.Append "ValorAgo", adDouble
      .Fields.Append "ValorSet", adDouble
      .Fields.Append "ValorOct", adDouble
      .Fields.Append "ValorNov", adDouble
      .Fields.Append "ValorDic", adDouble
      .Fields.Append "ValorTot", adDouble
      .LockType = adLockOptimistic
      .Open
    End With
    '
    lcSql = "select * from tFormDet where left(annomes,4)='" & Me.txtAnio.Text & "'"
    oRsFox1.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
    If oRsFox1.RecordCount > 0 Then
       oRsFox1.MoveFirst
       '
       lcTitulo = oRsFox1.Fields!codigo_pre
       If InStr(oRsFox1.Fields!codigo_pre, "F") > 0 Then
          lcSql = "select * from  mAlmacen where almCod='" & Left(Left(lcTitulo, InStr(lcTitulo, "F") - 1) & Space(10), 10) & "'"
       ElseIf InStr(oRsFox1.Fields!codigo_pre, "A") > 0 Then
          lcSql = "select * from  mAlmacen where almCod='" & Left(Left(lcTitulo, InStr(lcTitulo, "A") - 1) & Space(10), 10) & "'"
       Else
          lcSql = "select * from  mAlmacen where almCod='" & Left(lcTitulo & Space(10), 10) & "'"
       End If
       oRsFox2.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
       lcTitulo = ""
       If oRsFox2.RecordCount > 0 Then
          lcTitulo = "<<" & Trim(oRsFox2.Fields!almDes) & ">>     "
       End If
       oRsFox2.Close
       '
       Do While Not oRsFox1.EOF
If Trim(oRsFox1.Fields!codigo_med) = "01467" Then
lbNuevo = False
End If
          If InStr(oRsFox1.Fields!codigo_pre, "F") > 0 Then    'solo farmacia
                lcSql = "select * from mMedxProg where mCodPrg='" & lcPrograma & "' and " & _
                                                     " MCodSub='" & lcSubPrograma & "' and " & _
                                                     " mMedCod='" & oRsFox1.Fields!codigo_med & "'"
                oRsFox2.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
                If oRsFox2.RecordCount > 0 Then
                   lbNuevo = True
                   If oRsCubo.RecordCount > 0 Then
                      oRsCubo.MoveFirst
                      oRsCubo.Find "codigo='" & oRsFox1.Fields!codigo_med & "'"
                      If Not oRsCubo.EOF Then
                         lbNuevo = False
                      End If
                   End If
                   If lbNuevo = True Then
                      lcSql = "select * from Mproducto where medCod='" & oRsFox1.Fields!codigo_med & "'"
                      oRsFox3.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
                      lcDescripcion = ""
                      If oRsFox3.RecordCount > 0 Then
                         lcDescripcion = Left(Trim(oRsFox3.Fields!medNom) & " " & Trim(oRsFox3.Fields!medPres) & " " & Trim(oRsFox3.Fields!medcnc) & " " & Trim(oRsFox3.Fields!medFF), 300)
                      End If
                      oRsFox3.Close
                      oRsCubo.AddNew
                      oRsCubo.Fields!Codigo = oRsFox1.Fields!codigo_med
                      oRsCubo.Fields!Producto = lcDescripcion
                      oRsCubo.Fields!Precio = oRsFox1.Fields!Precio
                      oRsCubo.Fields!ValorEne = 0
                      oRsCubo.Fields!ValorFeb = 0
                      oRsCubo.Fields!ValorMar = 0
                      oRsCubo.Fields!ValorAbr = 0
                      oRsCubo.Fields!ValorMay = 0
                      oRsCubo.Fields!ValorJun = 0
                      oRsCubo.Fields!ValorAgo = 0
                      oRsCubo.Fields!ValorSet = 0
                      oRsCubo.Fields!ValorOct = 0
                      oRsCubo.Fields!ValorNov = 0
                      oRsCubo.Fields!ValorDic = 0
                      oRsCubo.Fields!ValorTot = 0
                      oRsCubo.Update
                   End If
                   '
                   lbStockEnRiesgo = False
                   If optValorStock.Value = True Then
                      lnDiasRiesgo = 1000
                      If sighEntidades.EsFecha(Format(oRsFox1.Fields!fec_exp, "dd/mm/yyyy"), "DD/MM/AAAA") And Year(oRsFox1.Fields!fec_exp) > 2000 Then
                         lnDiasRiesgo = DateDiff("d", Date, oRsFox1.Fields!fec_exp)
                      End If
                      If optMenor30.Value = True Then
                         If lnDiasRiesgo <= 30 Then
                            lbStockEnRiesgo = True
                         End If
                      ElseIf optEntre31y60.Value = True Then
                         If lnDiasRiesgo > 30 And lnDiasRiesgo <= 60 Then
                            lbStockEnRiesgo = True
                         End If
                      ElseIf optEntre61y90.Value = True Then
                         If lnDiasRiesgo > 60 And lnDiasRiesgo <= 90 Then
                            lbStockEnRiesgo = True
                         End If
                      End If
                   End If
                   '
                   lnSalidasIS = IIf(lcPrograma = "20", oRsFox1.Fields!interSan, 0)    'Salud Materna, disminuye Intervencion Sanitaria a Ventas y Stock
                   Select Case Val(Right(Trim(oRsFox1.Fields!annoMes), 2))
                   Case 1
                        If optSalidaTotal.Value = True Then
                           oRsCubo.Fields!ValorEne = oRsCubo.Fields!ValorEne + IIf(optMuestraCantidad.Value = True, (oRsFox1.Fields!Total - lnSalidasIS), (oRsFox1.Fields!Total - lnSalidasIS) * oRsFox1.Fields!Precio)
                        ElseIf optStockFinal.Value = True Then
                           oRsCubo.Fields!ValorEne = oRsCubo.Fields!ValorEne + IIf(optMuestraCantidad.Value = True, (oRsFox1.Fields!stock_fin - lnSalidasIS), (oRsFox1.Fields!stock_fin - lnSalidasIS) * oRsFox1.Fields!Precio)
                        ElseIf optDevolucionXvencimiento.Value = True Then
                           oRsCubo.Fields!ValorEne = oRsCubo.Fields!ValorEne + IIf(optMuestraCantidad.Value = True, oRsFox1.Fields!dev_ven, oRsFox1.Fields!dev_ven * oRsFox1.Fields!Precio)
                        ElseIf optValorStock.Value = True And lbStockEnRiesgo = True Then
                           oRsCubo.Fields!ValorEne = oRsCubo.Fields!ValorEne + IIf(optMuestraCantidad.Value = True, (oRsFox1.Fields!stock_fin - lnSalidasIS), (oRsFox1.Fields!stock_fin - lnSalidasIS) * oRsFox1.Fields!Precio)
                        End If
                   Case 2
                        If optSalidaTotal.Value = True Then
                           oRsCubo.Fields!ValorFeb = oRsCubo.Fields!ValorFeb + IIf(optMuestraCantidad.Value = True, (oRsFox1.Fields!Total - lnSalidasIS), (oRsFox1.Fields!Total - lnSalidasIS) * oRsFox1.Fields!Precio)
                        ElseIf optStockFinal.Value = True Then
                           oRsCubo.Fields!ValorFeb = oRsCubo.Fields!ValorFeb + IIf(optMuestraCantidad.Value = True, (oRsFox1.Fields!stock_fin - lnSalidasIS), (oRsFox1.Fields!stock_fin - lnSalidasIS) * oRsFox1.Fields!Precio)
                        ElseIf optDevolucionXvencimiento.Value = True Then
                           oRsCubo.Fields!ValorFeb = oRsCubo.Fields!ValorFeb + IIf(optMuestraCantidad.Value = True, oRsFox1.Fields!dev_ven, oRsFox1.Fields!dev_ven * oRsFox1.Fields!Precio)
                        ElseIf optValorStock.Value = True And lbStockEnRiesgo = True Then
                           oRsCubo.Fields!ValorFeb = oRsCubo.Fields!ValorFeb + IIf(optMuestraCantidad.Value = True, (oRsFox1.Fields!stock_fin - lnSalidasIS), (oRsFox1.Fields!stock_fin - lnSalidasIS) * oRsFox1.Fields!Precio)
                        End If
                   Case 3
                        If optSalidaTotal.Value = True Then
                           oRsCubo.Fields!ValorMar = oRsCubo.Fields!ValorMar + IIf(optMuestraCantidad.Value = True, (oRsFox1.Fields!Total - lnSalidasIS), (oRsFox1.Fields!Total - lnSalidasIS) * oRsFox1.Fields!Precio)
                        ElseIf optStockFinal.Value = True Then
                           oRsCubo.Fields!ValorMar = oRsCubo.Fields!ValorMar + IIf(optMuestraCantidad.Value = True, (oRsFox1.Fields!stock_fin - lnSalidasIS), (oRsFox1.Fields!stock_fin - lnSalidasIS) * oRsFox1.Fields!Precio)
                        ElseIf optDevolucionXvencimiento.Value = True Then
                           oRsCubo.Fields!ValorMar = oRsCubo.Fields!ValorMar + IIf(optMuestraCantidad.Value = True, oRsFox1.Fields!dev_ven, oRsFox1.Fields!dev_ven * oRsFox1.Fields!Precio)
                        ElseIf optValorStock.Value = True And lbStockEnRiesgo = True Then
                           oRsCubo.Fields!ValorMar = oRsCubo.Fields!ValorMar + IIf(optMuestraCantidad.Value = True, (oRsFox1.Fields!stock_fin - lnSalidasIS), (oRsFox1.Fields!stock_fin - lnSalidasIS) * oRsFox1.Fields!Precio)
                        End If
                   Case 4
                        If optSalidaTotal.Value = True Then
                           oRsCubo.Fields!ValorAbr = oRsCubo.Fields!ValorAbr + IIf(optMuestraCantidad.Value = True, (oRsFox1.Fields!Total - lnSalidasIS), (oRsFox1.Fields!Total - lnSalidasIS) * oRsFox1.Fields!Precio)
                        ElseIf optStockFinal.Value = True Then
                           oRsCubo.Fields!ValorAbr = oRsCubo.Fields!ValorAbr + IIf(optMuestraCantidad.Value = True, (oRsFox1.Fields!stock_fin - lnSalidasIS), (oRsFox1.Fields!stock_fin - lnSalidasIS) * oRsFox1.Fields!Precio)
                        ElseIf optDevolucionXvencimiento.Value = True Then
                           oRsCubo.Fields!ValorAbr = oRsCubo.Fields!ValorAbr + IIf(optMuestraCantidad.Value = True, oRsFox1.Fields!dev_ven, oRsFox1.Fields!dev_ven * oRsFox1.Fields!Precio)
                        ElseIf optValorStock.Value = True And lbStockEnRiesgo = True Then
                           oRsCubo.Fields!ValorAbr = oRsCubo.Fields!ValorAbr + IIf(optMuestraCantidad.Value = True, (oRsFox1.Fields!stock_fin - lnSalidasIS), (oRsFox1.Fields!stock_fin - lnSalidasIS) * oRsFox1.Fields!Precio)
                        End If
                   Case 5
                        If optSalidaTotal.Value = True Then
                           oRsCubo.Fields!ValorMay = oRsCubo.Fields!ValorMay + IIf(optMuestraCantidad.Value = True, (oRsFox1.Fields!Total - lnSalidasIS), (oRsFox1.Fields!Total - lnSalidasIS) * oRsFox1.Fields!Precio)
                        ElseIf optStockFinal.Value = True Then
                           oRsCubo.Fields!ValorMay = oRsCubo.Fields!ValorMay + IIf(optMuestraCantidad.Value = True, (oRsFox1.Fields!stock_fin - lnSalidasIS), (oRsFox1.Fields!stock_fin - lnSalidasIS) * oRsFox1.Fields!Precio)
                        ElseIf optDevolucionXvencimiento.Value = True Then
                           oRsCubo.Fields!ValorMay = oRsCubo.Fields!ValorMay + IIf(optMuestraCantidad.Value = True, oRsFox1.Fields!dev_ven, oRsFox1.Fields!dev_ven * oRsFox1.Fields!Precio)
                        ElseIf optValorStock.Value = True And lbStockEnRiesgo = True Then
                           oRsCubo.Fields!ValorMay = oRsCubo.Fields!ValorMay + IIf(optMuestraCantidad.Value = True, (oRsFox1.Fields!stock_fin - lnSalidasIS), (oRsFox1.Fields!stock_fin - lnSalidasIS) * oRsFox1.Fields!Precio)
                        End If
                   Case 6
                        If optSalidaTotal.Value = True Then
                           oRsCubo.Fields!ValorJun = oRsCubo.Fields!ValorJun + IIf(optMuestraCantidad.Value = True, (oRsFox1.Fields!Total - lnSalidasIS), (oRsFox1.Fields!Total - lnSalidasIS) * oRsFox1.Fields!Precio)
                        ElseIf optStockFinal.Value = True Then
                           oRsCubo.Fields!ValorJun = oRsCubo.Fields!ValorJun + IIf(optMuestraCantidad.Value = True, (oRsFox1.Fields!stock_fin - lnSalidasIS), (oRsFox1.Fields!stock_fin - lnSalidasIS) * oRsFox1.Fields!Precio)
                        ElseIf optDevolucionXvencimiento.Value = True Then
                           oRsCubo.Fields!ValorJun = oRsCubo.Fields!ValorJun + IIf(optMuestraCantidad.Value = True, oRsFox1.Fields!dev_ven, oRsFox1.Fields!dev_ven * oRsFox1.Fields!Precio)
                        ElseIf optValorStock.Value = True And lbStockEnRiesgo = True Then
                           oRsCubo.Fields!ValorJun = oRsCubo.Fields!ValorJun + IIf(optMuestraCantidad.Value = True, (oRsFox1.Fields!stock_fin - lnSalidasIS), (oRsFox1.Fields!stock_fin - lnSalidasIS) * oRsFox1.Fields!Precio)
                        End If
                   Case 7
                        If optSalidaTotal.Value = True Then
                           oRsCubo.Fields!ValorJul = oRsCubo.Fields!ValorJul + IIf(optMuestraCantidad.Value = True, (oRsFox1.Fields!Total - lnSalidasIS), (oRsFox1.Fields!Total - lnSalidasIS) * oRsFox1.Fields!Precio)
                        ElseIf optStockFinal.Value = True Then
                           oRsCubo.Fields!ValorJul = oRsCubo.Fields!ValorJul + IIf(optMuestraCantidad.Value = True, (oRsFox1.Fields!stock_fin - lnSalidasIS), (oRsFox1.Fields!stock_fin - lnSalidasIS) * oRsFox1.Fields!Precio)
                        ElseIf optDevolucionXvencimiento.Value = True Then
                           oRsCubo.Fields!ValorJul = oRsCubo.Fields!ValorJul + IIf(optMuestraCantidad.Value = True, oRsFox1.Fields!dev_ven, oRsFox1.Fields!dev_ven * oRsFox1.Fields!Precio)
                        ElseIf optValorStock.Value = True And lbStockEnRiesgo = True Then
                           oRsCubo.Fields!ValorJul = oRsCubo.Fields!ValorJul + IIf(optMuestraCantidad.Value = True, (oRsFox1.Fields!stock_fin - lnSalidasIS), (oRsFox1.Fields!stock_fin - lnSalidasIS) * oRsFox1.Fields!Precio)
                        End If
                   Case 8
                        If optSalidaTotal.Value = True Then
                           oRsCubo.Fields!ValorAgo = oRsCubo.Fields!ValorAgo + IIf(optMuestraCantidad.Value = True, (oRsFox1.Fields!Total - lnSalidasIS), (oRsFox1.Fields!Total - lnSalidasIS) * oRsFox1.Fields!Precio)
                        ElseIf optStockFinal.Value = True Then
                           oRsCubo.Fields!ValorAgo = oRsCubo.Fields!ValorAgo + IIf(optMuestraCantidad.Value = True, (oRsFox1.Fields!stock_fin - lnSalidasIS), (oRsFox1.Fields!stock_fin - lnSalidasIS) * oRsFox1.Fields!Precio)
                        ElseIf optDevolucionXvencimiento.Value = True Then
                           oRsCubo.Fields!ValorAgo = oRsCubo.Fields!ValorAgo + IIf(optMuestraCantidad.Value = True, oRsFox1.Fields!dev_ven, oRsFox1.Fields!dev_ven * oRsFox1.Fields!Precio)
                        ElseIf optValorStock.Value = True And lbStockEnRiesgo = True Then
                           oRsCubo.Fields!ValorAgo = oRsCubo.Fields!ValorAgo + IIf(optMuestraCantidad.Value = True, (oRsFox1.Fields!stock_fin - lnSalidasIS), (oRsFox1.Fields!stock_fin - lnSalidasIS) * oRsFox1.Fields!Precio)
                        End If
                   Case 9
                        If optSalidaTotal.Value = True Then
                           oRsCubo.Fields!ValorSet = oRsCubo.Fields!ValorSet + IIf(optMuestraCantidad.Value = True, (oRsFox1.Fields!Total - lnSalidasIS), (oRsFox1.Fields!Total - lnSalidasIS) * oRsFox1.Fields!Precio)
                        ElseIf optStockFinal.Value = True Then
                           oRsCubo.Fields!ValorSet = oRsCubo.Fields!ValorSet + IIf(optMuestraCantidad.Value = True, (oRsFox1.Fields!stock_fin - lnSalidasIS), (oRsFox1.Fields!stock_fin - lnSalidasIS) * oRsFox1.Fields!Precio)
                        ElseIf optDevolucionXvencimiento.Value = True Then
                           oRsCubo.Fields!ValorSet = oRsCubo.Fields!ValorSet + IIf(optMuestraCantidad.Value = True, oRsFox1.Fields!dev_ven, oRsFox1.Fields!dev_ven * oRsFox1.Fields!Precio)
                        ElseIf optValorStock.Value = True And lbStockEnRiesgo = True Then
                           oRsCubo.Fields!ValorSet = oRsCubo.Fields!ValorSet + IIf(optMuestraCantidad.Value = True, (oRsFox1.Fields!stock_fin - lnSalidasIS), (oRsFox1.Fields!stock_fin - lnSalidasIS) * oRsFox1.Fields!Precio)
                        End If
                   Case 10
                        If optSalidaTotal.Value = True Then
                           oRsCubo.Fields!ValorOct = oRsCubo.Fields!ValorOct + IIf(optMuestraCantidad.Value = True, (oRsFox1.Fields!Total - lnSalidasIS), (oRsFox1.Fields!Total - lnSalidasIS) * oRsFox1.Fields!Precio)
                        ElseIf optStockFinal.Value = True Then
                           oRsCubo.Fields!ValorOct = oRsCubo.Fields!ValorOct + IIf(optMuestraCantidad.Value = True, (oRsFox1.Fields!stock_fin - lnSalidasIS), (oRsFox1.Fields!stock_fin - lnSalidasIS) * oRsFox1.Fields!Precio)
                        ElseIf optDevolucionXvencimiento.Value = True Then
                           oRsCubo.Fields!ValorOct = oRsCubo.Fields!ValorOct + IIf(optMuestraCantidad.Value = True, oRsFox1.Fields!dev_ven, oRsFox1.Fields!dev_ven * oRsFox1.Fields!Precio)
                        ElseIf optValorStock.Value = True And lbStockEnRiesgo = True Then
                           oRsCubo.Fields!ValorOct = oRsCubo.Fields!ValorOct + IIf(optMuestraCantidad.Value = True, (oRsFox1.Fields!stock_fin - lnSalidasIS), (oRsFox1.Fields!stock_fin - lnSalidasIS) * oRsFox1.Fields!Precio)
                        End If
                   Case 11
                        If optSalidaTotal.Value = True Then
                           oRsCubo.Fields!ValorNov = oRsCubo.Fields!ValorNov + IIf(optMuestraCantidad.Value = True, (oRsFox1.Fields!Total - lnSalidasIS), (oRsFox1.Fields!Total - lnSalidasIS) * oRsFox1.Fields!Precio)
                        ElseIf optStockFinal.Value = True Then
                           oRsCubo.Fields!ValorNov = oRsCubo.Fields!ValorNov + IIf(optMuestraCantidad.Value = True, (oRsFox1.Fields!stock_fin - lnSalidasIS), (oRsFox1.Fields!stock_fin - lnSalidasIS) * oRsFox1.Fields!Precio)
                        ElseIf optDevolucionXvencimiento.Value = True Then
                           oRsCubo.Fields!ValorNov = oRsCubo.Fields!ValorNov + IIf(optMuestraCantidad.Value = True, oRsFox1.Fields!dev_ven, oRsFox1.Fields!dev_ven * oRsFox1.Fields!Precio)
                        ElseIf optValorStock.Value = True And lbStockEnRiesgo = True Then
                           oRsCubo.Fields!ValorNov = oRsCubo.Fields!ValorNov + IIf(optMuestraCantidad.Value = True, (oRsFox1.Fields!stock_fin - lnSalidasIS), (oRsFox1.Fields!stock_fin - lnSalidasIS) * oRsFox1.Fields!Precio)
                        End If
                   Case 12
                        If optSalidaTotal.Value = True Then
                           oRsCubo.Fields!ValorDic = oRsCubo.Fields!ValorDic + IIf(optMuestraCantidad.Value = True, (oRsFox1.Fields!Total - lnSalidasIS), (oRsFox1.Fields!Total - lnSalidasIS) * oRsFox1.Fields!Precio)
                        ElseIf optStockFinal.Value = True Then
                           oRsCubo.Fields!ValorDic = oRsCubo.Fields!ValorDic + IIf(optMuestraCantidad.Value = True, (oRsFox1.Fields!stock_fin - lnSalidasIS), (oRsFox1.Fields!stock_fin - lnSalidasIS) * oRsFox1.Fields!Precio)
                        ElseIf optDevolucionXvencimiento.Value = True Then
                           oRsCubo.Fields!ValorDic = oRsCubo.Fields!ValorDic + IIf(optMuestraCantidad.Value = True, oRsFox1.Fields!dev_ven, oRsFox1.Fields!dev_ven * oRsFox1.Fields!Precio)
                        ElseIf optValorStock.Value = True And lbStockEnRiesgo = True Then
                           oRsCubo.Fields!ValorDic = oRsCubo.Fields!ValorDic + IIf(optMuestraCantidad.Value = True, (oRsFox1.Fields!stock_fin - lnSalidasIS), (oRsFox1.Fields!stock_fin - lnSalidasIS) * oRsFox1.Fields!Precio)
                        End If
                   End Select
                   If optSalidaTotal.Value = True Then
                      oRsCubo.Fields!ValorTot = oRsCubo.Fields!ValorTot + IIf(optMuestraCantidad.Value = True, (oRsFox1.Fields!Total - lnSalidasIS), (oRsFox1.Fields!Total - lnSalidasIS) * oRsFox1.Fields!Precio)
                   ElseIf optStockFinal.Value = True Then
                      oRsCubo.Fields!ValorTot = oRsCubo.Fields!ValorTot + IIf(optMuestraCantidad.Value = True, (oRsFox1.Fields!stock_fin - lnSalidasIS), (oRsFox1.Fields!stock_fin - lnSalidasIS) * oRsFox1.Fields!Precio)
                   ElseIf optDevolucionXvencimiento.Value = True Then
                      oRsCubo.Fields!ValorTot = oRsCubo.Fields!ValorTot + IIf(optMuestraCantidad.Value = True, oRsFox1.Fields!dev_ven, oRsFox1.Fields!dev_ven * oRsFox1.Fields!Precio)
                   ElseIf optValorStock.Value = True And lbStockEnRiesgo = True Then
                      oRsCubo.Fields!ValorTot = oRsCubo.Fields!ValorTot + IIf(optMuestraCantidad.Value = True, (oRsFox1.Fields!stock_fin - lnSalidasIS), (oRsFox1.Fields!stock_fin - lnSalidasIS) * oRsFox1.Fields!Precio)
                   End If
                   oRsCubo.Update
                End If
                oRsFox2.Close
          Else
lbNuevo = False
          End If
          oRsFox1.MoveNext
       Loop
    End If
    oRsFox1.Close
    If oRsCubo.RecordCount = 0 Then
       MsgBox "No existen datos", vbCritical, Me.Caption
    Else
        oRsCubo.Sort = "producto"
        Dim oExcel As Excel.Application
        Dim oWorkBookPlantilla As Workbook
        Dim oWorkBook As Workbook
        Dim oWorkSheet As Worksheet
        Dim mo_ReporteUtil As New ReporteUtil
        Dim iFila As Integer
        Dim lnTotEne As Double, lnTotFeb As Double, lnTotMar As Double, lnTotAbr As Double
        Dim lnTotMay As Double, lnTotJun As Double, lnTotJul As Double, lnTotAgo As Double
        Dim lnTotSet As Double, lnTotOct As Double, lnTotNov As Double, lnTotDic As Double
        Dim lnTot As Double
        '
        Set oExcel = GalenhosExcelApplication()  'New Excel.Application
        Set oWorkBook = oExcel.Workbooks.Add
        Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Sismedv2.xls")
        oWorkBookPlantilla.Worksheets("ICI").Copy Before:=oWorkBook.Sheets(1)
        oWorkBookPlantilla.Close
        Set oWorkSheet = oWorkBook.Sheets(1)
        lcTitulo = lcTitulo & "(Programa: " & Trim(Me.cmdPrograma.Text) & ")     (Año: " & txtAnio.Text & ")     (" & _
                                     IIf(optMuestraCantidad.Value = True, "Muestra Cantidades", "Muestra Importes") & ")     (" & _
                                     "Reporte: " & IIf(optSalidaTotal.Value = True, optSalidaTotal.Caption, IIf(optStockFinal.Value = True, optStockFinal.Caption, IIf(optDevolucionXvencimiento.Value = True, optDevolucionXvencimiento.Caption, optValorStock.Caption))) & ")"
        If optValorStock.Value = True Then
           lcTitulo = lcTitulo & "     (" & IIf(optMenor30.Value = True, optMenor30.Caption, IIf(optEntre31y60.Value = True, optEntre31y60.Caption, optEntre61y90.Caption)) & ")"
        End If
        oWorkSheet.Cells(3, 2).Value = lcTitulo
        iFila = 6
        lnTotEne = 0: lnTotFeb = 0: lnTotMar = 0: lnTotAbr = 0
        lnTotMay = 0: lnTotJun = 0: lnTotJul = 0: lnTotAgo = 0
        lnTotSet = 0: lnTotOct = 0: lnTotNov = 0: lnTotDic = 0
        lnTot = 0
        oRsCubo.MoveFirst
        Do While Not oRsCubo.EOF
           If oRsCubo.Fields!ValorTot > 0 Then
                oWorkSheet.Cells(iFila, 2).Value = oRsCubo.Fields!Codigo
                oWorkSheet.Cells(iFila, 3).Value = oRsCubo.Fields!Producto
                oWorkSheet.Cells(iFila, 4).Value = oRsCubo.Fields!Precio
                oWorkSheet.Cells(iFila, 5).Value = oRsCubo.Fields!ValorEne
                oWorkSheet.Cells(iFila, 6).Value = oRsCubo.Fields!ValorFeb
                oWorkSheet.Cells(iFila, 7).Value = oRsCubo.Fields!ValorMar
                oWorkSheet.Cells(iFila, 8).Value = oRsCubo.Fields!ValorAbr
                oWorkSheet.Cells(iFila, 9).Value = oRsCubo.Fields!ValorMay
                oWorkSheet.Cells(iFila, 10).Value = oRsCubo.Fields!ValorJun
                oWorkSheet.Cells(iFila, 11).Value = oRsCubo.Fields!ValorJul
                oWorkSheet.Cells(iFila, 12).Value = oRsCubo.Fields!ValorAgo
                oWorkSheet.Cells(iFila, 13).Value = oRsCubo.Fields!ValorSet
                oWorkSheet.Cells(iFila, 14).Value = oRsCubo.Fields!ValorOct
                oWorkSheet.Cells(iFila, 15).Value = oRsCubo.Fields!ValorNov
                oWorkSheet.Cells(iFila, 16).Value = oRsCubo.Fields!ValorDic
                oWorkSheet.Cells(iFila, 17).Value = oRsCubo.Fields!ValorTot
                If optMuestraCantidad.Value = True Then
                   oWorkSheet.Cells(iFila, 18).Value = oRsCubo.Fields!ValorTot * oRsCubo.Fields!Precio
                End If
                iFila = iFila + 1
                If optMuestraCantidad.Value = True Then
                    lnTotEne = lnTotEne + oRsCubo.Fields!ValorEne * oRsCubo.Fields!Precio
                    lnTotFeb = lnTotFeb + oRsCubo.Fields!ValorFeb * oRsCubo.Fields!Precio
                    lnTotMar = lnTotMar + oRsCubo.Fields!ValorMar * oRsCubo.Fields!Precio
                    lnTotAbr = lnTotAbr + oRsCubo.Fields!ValorAbr * oRsCubo.Fields!Precio
                    lnTotMay = lnTotMay + oRsCubo.Fields!ValorMay * oRsCubo.Fields!Precio
                    lnTotJun = lnTotJun + oRsCubo.Fields!ValorJun * oRsCubo.Fields!Precio
                    lnTotJul = lnTotJul + oRsCubo.Fields!ValorJul * oRsCubo.Fields!Precio
                    lnTotAgo = lnTotAgo + oRsCubo.Fields!ValorAgo * oRsCubo.Fields!Precio
                    lnTotSet = lnTotSet + oRsCubo.Fields!ValorSet * oRsCubo.Fields!Precio
                    lnTotOct = lnTotOct + oRsCubo.Fields!ValorOct * oRsCubo.Fields!Precio
                    lnTotNov = lnTotNov + oRsCubo.Fields!ValorNov * oRsCubo.Fields!Precio
                    lnTotDic = lnTotDic + oRsCubo.Fields!ValorDic * oRsCubo.Fields!Precio
                    lnTot = lnTot + oRsCubo.Fields!ValorTot * oRsCubo.Fields!Precio
                Else
                    lnTotEne = lnTotEne + oRsCubo.Fields!ValorEne
                    lnTotFeb = lnTotFeb + oRsCubo.Fields!ValorFeb
                    lnTotMar = lnTotMar + oRsCubo.Fields!ValorMar
                    lnTotAbr = lnTotAbr + oRsCubo.Fields!ValorAbr
                    lnTotMay = lnTotMay + oRsCubo.Fields!ValorMay
                    lnTotJun = lnTotJun + oRsCubo.Fields!ValorJun
                    lnTotJul = lnTotJul + oRsCubo.Fields!ValorJul
                    lnTotAgo = lnTotAgo + oRsCubo.Fields!ValorAgo
                    lnTotSet = lnTotSet + oRsCubo.Fields!ValorSet
                    lnTotOct = lnTotOct + oRsCubo.Fields!ValorOct
                    lnTotNov = lnTotNov + oRsCubo.Fields!ValorNov
                    lnTotDic = lnTotDic + oRsCubo.Fields!ValorDic
                    lnTot = lnTot + oRsCubo.Fields!ValorTot
                End If
           End If
           oRsCubo.MoveNext
        Loop
        oWorkSheet.Cells(iFila, 5).Value = lnTotEne
        oWorkSheet.Cells(iFila, 6).Value = lnTotFeb
        oWorkSheet.Cells(iFila, 7).Value = lnTotMar
        oWorkSheet.Cells(iFila, 8).Value = lnTotAbr
        oWorkSheet.Cells(iFila, 9).Value = lnTotMay
        oWorkSheet.Cells(iFila, 10).Value = lnTotJun
        oWorkSheet.Cells(iFila, 11).Value = lnTotJul
        oWorkSheet.Cells(iFila, 12).Value = lnTotAgo
        oWorkSheet.Cells(iFila, 13).Value = lnTotSet
        oWorkSheet.Cells(iFila, 14).Value = lnTotOct
        oWorkSheet.Cells(iFila, 15).Value = lnTotNov
        oWorkSheet.Cells(iFila, 16).Value = lnTotDic
        oWorkSheet.Cells(iFila, 17).Value = lnTot
        If optMuestraCantidad.Value = True Then
           oWorkSheet.Cells(iFila, 18).Value = lnTot
        End If
        '
        oExcel.Visible = True
        oWorkSheet.PrintPreview
    End If
End Sub

Private Sub btnCancelar_Click()
    On Error Resume Next
    oRsFoxPrograma.Close
    oConexionFox.Close
    oRsFox1.Close
    oRsFox2.Close
    oRsFox3.Close
    Unload Me
End Sub


Private Sub Form_Load()
    'Año
    txtAnio.Text = Year(Date)
    '
    If optGalenhos.Value = True Then
       optGalenhos_Click
    Else
       optSismedv2_Click
    End If
End Sub



Private Sub optGalenhos_Click()
    AbreConexion False
End Sub

Private Sub optSismedv2_Click()
    AbreConexion True
End Sub

Sub AbreConexion(lbDesdeSismedv2 As Boolean)
    On Error GoTo errConexion
    If oConexionFox.State = 1 Then
       Set oConexionFox = Nothing
    End If
    oConexionFox.CommandTimeout = 300
    If lbDesdeSismedv2 = True Then
       oConexionFox.Open "DSN=HIS"
    Else
       oConexionFox.Open "DSN=GalenHosSql2008" 'lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
    End If
    'Carga Combo PROGRAMA
    If oRsFoxPrograma.State = 1 Then
       Set oRsFoxPrograma = Nothing
    End If
    lcSql = "select subCodPrg+Subcod as Codigo,subDes from mSubComponente order by subDes"
    oRsFoxPrograma.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
    cmdPrograma.ListField = "subDes"
    cmdPrograma.BoundColumn = "Codigo"
    Set cmdPrograma.RowSource = oRsFoxPrograma
    cmdPrograma.Text = ""
    Exit Sub
errConexion:
    MsgBox Err.Description
End Sub
