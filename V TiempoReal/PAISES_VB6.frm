VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form PAISES_VB6 
   Caption         =   "FLETE01_DAT"
   ClientHeight    =   4230
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   3660
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4230
   ScaleWidth      =   3660
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   5340
      TabIndex        =   0
      Top             =   3960
      Width           =   1080
   End
   Begin MSAdodcLib.Adodc datPrimaryRS 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   3900
      Width           =   3660
      _ExtentX        =   6456
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "PROVIDER=MSDASQL;dsn=DESICA806;uid=ops$desfle00;pwd=ops$desfle00;"
      OLEDBString     =   "PROVIDER=MSDASQL;dsn=DESICA806;uid=ops$desfle00;pwd=ops$desfle00;"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select C1_COD_PAIS,C1_NOMBRE_PAIS from FLETE01_DAT Order by C1_COD_PAIS"
      Caption         =   " "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Bindings        =   "PAISES_VB6.frx":0000
      DragIcon        =   "PAISES_VB6.frx":001B
      Height          =   3840
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   6360
      _ExtentX        =   11218
      _ExtentY        =   6773
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      Rows            =   3
      FixedCols       =   0
      BackColorFixed  =   8421376
      ForeColorFixed  =   16777215
      GridColor       =   12632256
      GridColorFixed  =   0
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   0
      GridLinesFixed  =   1
      AllowUserResizing=   1
      FormatString    =   "C1_COD_PAIS|C1_NOMBRE_PAIS"
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLineWidthBand=   1
      _Band(0).TextStyleBand=   0
   End
   Begin VB.Image imgSort 
      Height          =   480
      Index           =   1
      Left            =   2730
      Top             =   2355
      Width           =   1200
   End
End
Attribute VB_Name = "PAISES_VB6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************************************
'*   Creado x el Asistente de Formularios via ODBC
'*   Desarrollado x Henry J. Pulgar B.
'*
'*********************************************************************
Private Const MARGIN_SIZE = 60     ' en twips
' variables para permitir el orden de columnas
Private m_iSortCol As Integer
Private m_iSortType As Integer

' variables para arrastrar columnas
Private m_bDragOK As Boolean
Private m_iDragCol As Integer
Private xdn As Integer, ydn As Integer

Private Sub Form_Load()
    Dim i As Integer

    datPrimaryRS.Visible = False

    With MSHFlexGrid1

        .Redraw = False
        ' establecer anchos de columna de cuadrícula
        .ColWidth(0) = -1
        .ColWidth(1) = -1

        ' establecer tipo de cuadrícula
        .AllowBigSelection = True
        .FillStyle = flexFillRepeat

        ' encabezado en negrita
        .Row = 0
        .Col = 0
        .RowSel = .FixedRows - 1
        .ColSel = .Cols - 1
        .CellFontBold = True

        ' atenuar otra columna
        For i = .FixedCols To .Cols() - 1 Step 2
            .Col = i
            .Row = .FixedRows
            .RowSel = .Rows - 1
            .CellBackColor = &HC0C0C0   ' gris claro
        Next i

        .AllowBigSelection = False
        .FillStyle = flexFillSingle
        .Redraw = True

    End With

End Sub

Private Sub MSHFlexGrid1_DragDrop(Source As Control, X As Single, Y As Single)
'-------------------------------------------------------------------------------------------
' el código de los eventos DragDrop, MouseDown, MouseMove y MouseUp permite arrastrar columnas
'-------------------------------------------------------------------------------------------

    If m_iDragCol = -1 Then Exit Sub    ' no se estaba arrastrando
    If MSHFlexGrid1.MouseRow <> 0 Then Exit Sub

    With MSHFlexGrid1
        .Redraw = False
        .ColPosition(m_iDragCol) = .MouseCol

        .FillStyle = flexFillRepeat
        .Col = 0
        .Row = .FixedRows
        .RowSel = .Rows - 1
        .ColSel = .Cols - 1
        .CellBackColor = &HFFFFFF
        Dim iLoop As Integer
        For iLoop = .FixedCols To .Cols() - 1 Step 2
            .Col = iLoop
            .Row = .FixedRows
            .RowSel = .Rows - 1
            .CellBackColor = &HC0C0C0
        Next iLoop
        .FillStyle = flexFillSingle

        .Redraw = True
    End With

End Sub

Private Sub MSHFlexGrid1_MouseDown(Button As Integer, shift As Integer, X As Single, Y As Single)
'-------------------------------------------------------------------------------------------
' el código de los eventos DragDrop, MouseDown, MouseMove y MouseUp permite arrastrar columnas
'-------------------------------------------------------------------------------------------

    If MSHFlexGrid1.MouseRow <> 0 Then Exit Sub

    xdn = X
    ydn = Y
    m_iDragCol = -1     ' borrar indicador de arrastre
    m_bDragOK = True

End Sub

Private Sub MSHFlexGrid1_MouseMove(Button As Integer, shift As Integer, X As Single, Y As Single)
'-------------------------------------------------------------------------------------------
' el código de los eventos DragDrop, MouseDown, MouseMove y MouseUp permite arrastrar columnas
'-------------------------------------------------------------------------------------------

    ' probar si se debe iniciar el arrastre
    If Not m_bDragOK Then Exit Sub
    If Button <> 1 Then Exit Sub                        ' botón incorrecto
    If m_iDragCol <> -1 Then Exit Sub                   ' ya se está arrastrando
    If Abs(xdn - X) + Abs(ydn - Y) < 50 Then Exit Sub   ' no se ha movido suficiente
    If MSHFlexGrid1.MouseRow <> 0 Then Exit Sub         ' hay que arrastrar el encabezado

    ' si se llega aquí, iniciar el arrastre
    m_iDragCol = MSHFlexGrid1.MouseCol
    MSHFlexGrid1.Drag vbBeginDrag

End Sub

Private Sub MSHFlexGrid1_MouseUp(Button As Integer, shift As Integer, X As Single, Y As Single)
'-------------------------------------------------------------------------------------------
' el código de los eventos DragDrop, MouseDown, MouseMove y MouseUp permite arrastrar columnas
'-------------------------------------------------------------------------------------------

    m_bDragOK = False

End Sub

Private Sub MSHFlexGrid1_DblClick()
'-------------------------------------------------------------------------------------------
' el código del evento DblClick de la cuadrícula permite ordenar columnas
'-------------------------------------------------------------------------------------------

    Dim i As Integer

    ' sólo ordena cuando se hace clic en una fila
    If MSHFlexGrid1.MouseRow >= MSHFlexGrid1.FixedRows Then Exit Sub

    i = m_iSortCol                  ' guarda la columna antigua
    m_iSortCol = MSHFlexGrid1.Col   ' establece la nueva columna

    ' incrementa el tipo de orden
    If i <> m_iSortCol Then
        ' si hace clic en una columna nueva, inicia con orden ascendente
        m_iSortType = 1
    Else
        ' si hace clic en la misma columna, alterna entre orden ascendente y orden descendente
        m_iSortType = m_iSortType + 1
    If m_iSortType = 3 Then m_iSortType = 1
    End If

    DoColumnSort

End Sub

Sub DoColumnSort()
'-------------------------------------------------------------------------------------------
' orden de tipo intercambio en la columna m_iSortCol
'-------------------------------------------------------------------------------------------

    With MSHFlexGrid1
        .Redraw = False
        .Row = 1
        .RowSel = .Rows - 1
        .Col = m_iSortCol
        .Sort = m_iSortType
        .Redraw = True
    End With

End Sub

Private Sub Form_Resize()

    Dim sngButtonTop As Single
    Dim sngScaleWidth As Single
    Dim sngScaleHeight As Single

    On Error GoTo Form_Resize_Error
    With Me
        sngScaleWidth = .ScaleWidth
        sngScaleHeight = .ScaleHeight

        ' mueve el botón Cerrar a la esquina superior derecha
        With .cmdClose
                sngButtonTop = sngScaleHeight - (.Height + MARGIN_SIZE)
                .Move sngScaleWidth - (.Width + MARGIN_SIZE), sngButtonTop
        End With

        .MSHFlexGrid1.Move MARGIN_SIZE, _
            MARGIN_SIZE, _
            sngScaleWidth - (2 * MARGIN_SIZE), _
            sngButtonTop - (2 * MARGIN_SIZE)

    End With
    Exit Sub

Form_Resize_Error:
    ' evita errores en valores negativos
    Resume Next

End Sub
Private Sub cmdClose_Click()

    Unload Me

End Sub


