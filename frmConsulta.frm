VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmConsulta 
   Caption         =   "Form1"
   ClientHeight    =   6825
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11535
   LinkTopic       =   "Form1"
   ScaleHeight     =   6825
   ScaleWidth      =   11535
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdProximo 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4800
      TabIndex        =   2
      Top             =   720
      Width           =   315
   End
   Begin VB.CommandButton cmdAnterior 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3135
      TabIndex        =   1
      Top             =   720
      Width           =   315
   End
   Begin MSFlexGridLib.MSFlexGrid MSHFlexGrid1 
      Height          =   4140
      Left            =   1200
      TabIndex        =   0
      Top             =   1440
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   7303
      _Version        =   393216
      Rows            =   1
      Cols            =   6
      FixedCols       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblPagina 
      Caption         =   "Label1"
      Height          =   735
      Left            =   6840
      TabIndex        =   3
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "frmConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim paginaAtual As Integer
Dim totalPaginas As Integer
Dim totalRegistros As Integer
Dim registrosPorPagina As Integer




Private Sub ConectarBanco()
    On Error Resume Next
    Set conn = New ADODB.Connection
    conn.ConnectionString = "Provider=SQLOLEDB;Data Source=DESKTOP-CGCLBFQ;Initial Catalog=XYZ;User ID=sa;Password=ouro18;"
    conn.Open

End Sub

Private Sub CarregarGrid()
    Dim i As Integer, j As Integer
    Dim SQL As String
    Dim rsTotal As ADODB.Recordset
    Dim inicio As Integer

    
    ConectarBanco

    
    registrosPorPagina = 2

    
    inicio = (paginaAtual - 1) * registrosPorPagina

    
    SQL = "SELECT * FROM ( " & _
          "SELECT *, ROW_NUMBER() OVER (ORDER BY ID_Transacao) AS RowNum FROM Transacao) AS T " & _
          "WHERE T.RowNum BETWEEN " & inicio + 1 & " AND " & inicio + registrosPorPagina

    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open SQL, conn, adOpenStatic, adLockReadOnly

    
    If rs.EOF Then
        MsgBox "Nenhum registro encontrado.", vbInformation
        Exit Sub
    End If

    
    SQL = "SELECT COUNT(*) AS Total FROM Transacao"
    Set rsTotal = New ADODB.Recordset
    rsTotal.Open SQL, conn, adOpenStatic, adLockReadOnly
    totalRegistros = rsTotal!Total
    rsTotal.Close
    Set rsTotal = Nothing

    totalPaginas = (totalRegistros + registrosPorPagina - 1) \ registrosPorPagina

    
    MSHFlexGrid1.Rows = rs.RecordCount + 1
    MSHFlexGrid1.Cols = rs.Fields.Count

   
    For i = 0 To rs.Fields.Count - 1
        MSHFlexGrid1.TextMatrix(0, i) = rs.Fields(i).Name
    Next i

    
    i = 1
    Do While Not rs.EOF
        For j = 0 To rs.Fields.Count - 1
            MSHFlexGrid1.TextMatrix(i, j) = rs.Fields(j).Value
        Next j
        rs.MoveNext
        i = i + 1
    Loop

    lblPagina.Caption = "Página " & paginaAtual & " de " & totalPaginas

    'MsgBox "Dados carregados na página " & paginaAtual & "!", vbInformation
End Sub


Private Function Ceiling(ByVal X As Double) As Integer
    If X = Int(X) Then
        Ceiling = X
    Else
        Ceiling = Int(X) + 1
    End If
End Function

Private Sub cmdProximo_Click()
    If paginaAtual < totalPaginas Then
        paginaAtual = paginaAtual + 1
        CarregarGrid
    End If
End Sub

Private Sub cmdAnterior_Click()
    If paginaAtual > 1 Then
        paginaAtual = paginaAtual - 1
        CarregarGrid
    End If
End Sub


' Inicializa a conexão e carrega a primeira página
Private Sub Form_Load()
    paginaAtual = 1
    CarregarGrid
End Sub




