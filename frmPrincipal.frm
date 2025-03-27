VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPrincipal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Projeto XYZ"
   ClientHeight    =   6960
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   9690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   9690
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   6735
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   11880
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Gravar/Alterar"
      TabPicture(0)   =   "frmPrincipal.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Pesquisar"
      TabPicture(1)   =   "frmPrincipal.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frmGrid"
      Tab(1).Control(1)=   "frmPesquisa"
      Tab(1).ControlCount=   2
      Begin VB.Frame frmPesquisa 
         Height          =   6255
         Left            =   -74880
         TabIndex        =   31
         Top             =   360
         Width           =   9375
         Begin VB.TextBox txtValorP1 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   240
            TabIndex        =   39
            Text            =   "0,00"
            Top             =   1320
            Width           =   1305
         End
         Begin VB.TextBox txtNumCartaoP 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1560
            MaxLength       =   16
            TabIndex        =   38
            Top             =   600
            Width           =   2535
         End
         Begin VB.TextBox txtValorP2 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   2040
            TabIndex        =   37
            Text            =   "0,00"
            Top             =   1320
            Width           =   1305
         End
         Begin VB.TextBox txtDescricaoP 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1245
            Left            =   240
            MultiLine       =   -1  'True
            TabIndex        =   36
            Top             =   2760
            Width           =   6855
         End
         Begin VB.ListBox lstStatus 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   1605
            Left            =   4470
            Style           =   1  'Checkbox
            TabIndex        =   35
            Top             =   645
            Width           =   1455
         End
         Begin VB.CommandButton cmdPesquisar 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Pesquisar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   4320
            Width           =   1275
         End
         Begin VB.TextBox txtIdTransacaoP 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   240
            TabIndex        =   33
            Top             =   600
            Width           =   1095
         End
         Begin VB.CommandButton cmdLimparP 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Limpar Pesquisa"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            Left            =   5040
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   4320
            Width           =   1995
         End
         Begin MSMask.MaskEdBox MaskDataP1 
            Height          =   375
            Left            =   240
            TabIndex        =   40
            Top             =   2040
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskDataP2 
            Height          =   375
            Left            =   2400
            TabIndex        =   41
            Top             =   2040
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label11 
            Caption         =   "Número do Cartão"
            Height          =   375
            Left            =   1560
            TabIndex        =   49
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label12 
            Caption         =   "Valor"
            Height          =   375
            Left            =   240
            TabIndex        =   48
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label Label13 
            Caption         =   "Dt. Transação"
            Height          =   375
            Left            =   240
            TabIndex        =   47
            Top             =   1800
            Width           =   1335
         End
         Begin VB.Label Label14 
            Caption         =   "até"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1630
            TabIndex        =   46
            Top             =   1440
            Width           =   375
         End
         Begin VB.Label Label15 
            Caption         =   "até"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2000
            TabIndex        =   45
            Top             =   2160
            Width           =   375
         End
         Begin VB.Label Label16 
            Caption         =   "Descrição"
            Height          =   375
            Left            =   240
            TabIndex        =   44
            Top             =   2520
            Width           =   1335
         End
         Begin VB.Label Label17 
            Caption         =   "Status"
            Height          =   375
            Left            =   4440
            TabIndex        =   43
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label18 
            Caption         =   "ID"
            Height          =   375
            Left            =   240
            TabIndex        =   42
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame frmGrid 
         Height          =   6255
         Left            =   -74880
         TabIndex        =   23
         Top             =   360
         Width           =   9375
         Begin VB.CommandButton cmdProximo 
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "MS Reference Sans Serif"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   7080
            TabIndex        =   27
            Top             =   5400
            Width           =   435
         End
         Begin VB.CommandButton cmdAnterior 
            Caption         =   "<"
            BeginProperty Font 
               Name            =   "MS Reference Sans Serif"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   1800
            TabIndex        =   26
            Top             =   5400
            Width           =   435
         End
         Begin VB.CommandButton cmdAtvPesquisar 
            Caption         =   "Pesquisar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   7800
            TabIndex        =   25
            Top             =   5400
            Width           =   1395
         End
         Begin VB.CommandButton cmdExportar 
            Caption         =   "Exportar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   120
            TabIndex        =   24
            Top             =   5400
            Width           =   1395
         End
         Begin MSFlexGridLib.MSFlexGrid Grid1 
            Height          =   4980
            Left            =   120
            TabIndex        =   28
            Top             =   360
            Width           =   9075
            _ExtentX        =   16007
            _ExtentY        =   8784
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
            Alignment       =   2  'Center
            Caption         =   "Paginação"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2400
            TabIndex        =   30
            Top             =   5520
            Width           =   4695
         End
         Begin VB.Label Label10 
            Caption         =   "* Para alterar um cadastro, basta dar um duplo clique na tabela (grid)."
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   120
            Width           =   6375
         End
      End
      Begin VB.Frame Frame1 
         Height          =   6255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   9375
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   5520
            Top             =   5520
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.TextBox txtValor 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   4320
            TabIndex        =   12
            Text            =   "0,00"
            Top             =   480
            Width           =   1305
         End
         Begin VB.TextBox txtIdTransacao 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   240
            TabIndex        =   11
            Top             =   480
            Width           =   1095
         End
         Begin VB.TextBox txtNumCartao 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1560
            MaxLength       =   16
            TabIndex        =   10
            Top             =   480
            Width           =   2535
         End
         Begin VB.ComboBox cboStatus 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   2040
            TabIndex        =   9
            Text            =   "Combo1"
            Top             =   1320
            Width           =   2055
         End
         Begin VB.TextBox txtDescricao 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1245
            Left            =   240
            MultiLine       =   -1  'True
            TabIndex        =   8
            Top             =   2160
            Width           =   7815
         End
         Begin VB.Frame frmBotoes 
            Height          =   1095
            Left            =   0
            TabIndex        =   2
            Top             =   3960
            Width           =   9375
            Begin VB.CommandButton cmdExcluir 
               BackColor       =   &H008080FF&
               Caption         =   "Excluir"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   570
               Left            =   3480
               Style           =   1  'Graphical
               TabIndex        =   7
               Top             =   360
               Width           =   1275
            End
            Begin VB.CommandButton cmdLimpar 
               BackColor       =   &H00C0FFFF&
               Caption         =   "Limpar"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   570
               Left            =   1680
               Style           =   1  'Graphical
               TabIndex        =   6
               Top             =   360
               Width           =   1275
            End
            Begin VB.CommandButton cmdAlterar 
               BackColor       =   &H00FFFFC0&
               Caption         =   "Alterar"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   570
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   5
               Top             =   360
               Visible         =   0   'False
               Width           =   1275
            End
            Begin VB.CommandButton cmdGravar 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Gravar"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   570
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   4
               Top             =   360
               Width           =   1275
            End
            Begin VB.CommandButton cmdSair 
               BackColor       =   &H00FFC0C0&
               Caption         =   "Sair"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   570
               Left            =   7920
               Style           =   1  'Graphical
               TabIndex        =   3
               Top             =   360
               Width           =   1275
            End
         End
         Begin MSMask.MaskEdBox MaskData 
            Height          =   375
            Left            =   240
            TabIndex        =   13
            Top             =   1320
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label8 
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   4710
            TabIndex        =   15
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label9 
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   2520
            TabIndex        =   16
            Top             =   1080
            Width           =   255
         End
         Begin VB.Label Label2 
            Caption         =   "Número do Cartão"
            Height          =   375
            Left            =   1560
            TabIndex        =   22
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "ID"
            Height          =   375
            Left            =   240
            TabIndex        =   21
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "Valor"
            Height          =   375
            Left            =   4320
            TabIndex        =   20
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label4 
            Caption         =   "Dt. Transação"
            Height          =   375
            Left            =   240
            TabIndex        =   19
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label Label5 
            Caption         =   "Status"
            Height          =   375
            Left            =   2040
            TabIndex        =   18
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label Label6 
            Caption         =   "Descrição"
            Height          =   375
            Left            =   240
            TabIndex        =   17
            Top             =   1920
            Width           =   1335
         End
         Begin VB.Label Label7 
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   2880
            TabIndex        =   14
            Top             =   240
            Width           =   255
         End
      End
   End
End
Attribute VB_Name = "frmPrincipal"
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
Dim SQL As String
Dim SQLPesquisa As String
Dim i As Integer
Dim detalhesErro As String




Private Sub cmdAlterar_Click()
    Alterar
End Sub

Private Sub Alterar()
On Error GoTo TratarErro

    Screen.MousePointer = vbHourglass

    If ValidaDados = False Then Exit Sub
    
    
    Dim SQL As String

            SQL = "UPDATE Transacao SET Id_Status = " & cboStatus.ItemData(cboStatus.ListIndex) & _
            ", Numero_Cartao = '" & Trim(txtNumCartao.Text) & "', Descricao = '" & Trim(txtDescricao.Text) & _
            "', Data_Transacao = '" & Format(MaskData, "yyyymmdd") & "', Valor_Transacao = '" & Trim(Replace(txtValor.Text, ",", ".")) & _
            "' WHERE Id_Transacao = " & txtIdTransacao & ""
        conn.Execute SQL
        
       
        MsgBox "Cadastro Atualizado Com Sucesso!", vbInformation
        
        Limpar
        Screen.MousePointer = vbNormal
        Exit Sub


TratarErro:
    detalhesErro = "Número do erro: " & Err.Number & " - Descrição do erro: " & IIf(Err.Description = "", "Descrição indisponível", Err.Description)
    RegistrarErro "Erro na subrotina Limpar: " & detalhesErro
    MsgBox "Ocorreu um erro. Consulte o arquivo de log para mais detalhes.", vbExclamation, "Erro"
    Screen.MousePointer = vbNormal
End Sub

Private Sub cmdAtvPesquisar_Click()
    frmPesquisa.Visible = True
End Sub

Private Sub cmdExcluir_Click()
On Error GoTo TratarErro

    Screen.MousePointer = vbHourglass

    Dim Pergunta As Variant
    Pergunta = MsgBox("Deseja realmente excluir a transação?", vbYesNo, "Confirma")
    If (Pergunta = vbYes) Then

        conn.Execute "DELETE FROM Transacao where Id_Transacao =  " & txtIdTransacao & ""
        SQL = "INSERT INTO Log_Exclusao (Id_Transacao, Numero_Cartao, Valor_Transacao, Data_Transacao, Id_Status, Descricao, Data_Acao)"
        SQL = SQL & " VALUES (" & txtIdTransacao.Text & _
            ", '" & Trim(txtNumCartao.Text) & _
            "', '" & Trim(Replace(txtValor.Text, ",", ".")) & _
            "', '" & Format(MaskData, "yyyymmdd") & _
            "', " & cboStatus.ItemData(cboStatus.ListIndex) & _
            ", '" & Trim(txtDescricao.Text) & _
            "', '" & Format(Now, "yyyymmdd") & "')"
        conn.Execute SQL
        MsgBox "A transação foi excluída com sucesso!", vbInformation + vbOKOnly, "Informação Importante"
        Limpar
    Else
        MsgBox "Ação cancelada", vbInformation, "Informação"
    End If

    Screen.MousePointer = vbNormal
    Exit Sub

TratarErro:
    detalhesErro = "Número do erro: " & Err.Number & " - Descrição do erro: " & IIf(Err.Description = "", "Descrição indisponível", Err.Description)
    RegistrarErro "Erro na subrotina Limpar: " & detalhesErro
    MsgBox "Ocorreu um erro. Consulte o arquivo de log para mais detalhes.", vbExclamation, "Erro"
    Screen.MousePointer = vbNormal
End Sub


Private Sub cmdExportar_Click()
    ExportarParaExcel
End Sub

Private Sub cmdGravar_Click()

    Cadastrar
    
End Sub

Private Sub Cadastrar()


    Screen.MousePointer = vbHourglass


    If ValidaDados = False Then Exit Sub

On Error GoTo TratarErro

            SQL = "INSERT INTO Transacao (Numero_Cartao, Valor_Transacao, Data_Transacao, Id_Status,Descricao)"
            SQL = SQL & " VALUES ('" & Trim(txtNumCartao.Text) & _
                "', '" & Trim(Replace(txtValor.Text, ",", ".")) & _
                "', '" & Format(MaskData, "yyyymmdd") & _
                "', " & cboStatus.ItemData(cboStatus.ListIndex) & _
                ", '" & Trim(txtDescricao.Text) & "')"
            conn.Execute SQL
            
            MsgBox "Cadastro Realizado Com Sucesso!", vbInformation, "Informação"
            Limpar
  
            
       Screen.MousePointer = vbNormal
        Exit Sub


TratarErro:
    detalhesErro = "Número do erro: " & Err.Number & " - Descrição do erro: " & IIf(Err.Description = "", "Descrição indisponível", Err.Description)
    RegistrarErro "Erro na subrotina Limpar: " & detalhesErro
    MsgBox "Ocorreu um erro. Consulte o arquivo de log para mais detalhes.", vbExclamation, "Erro"
    Screen.MousePointer = vbNormal
End Sub

Private Sub cmdLimpar_Click()


    Limpar
    cmdAlterar.Visible = False
    cmdGravar.Visible = True


End Sub

Private Sub cmdLimparP_Click()


    LimparPesquisa
    

End Sub

Private Sub cmdPesquisar_Click()

    CarregarGrid
    
End Sub





Private Sub Grid1_DblClick()
On Error GoTo TratarErro

    If Grid1.Row = 0 Then Exit Sub
    
     Screen.MousePointer = vbHourglass

    cmdAlterar.Visible = True
    cmdExcluir.Visible = True
    cmdGravar.Visible = False
    txtIdTransacao.Text = Grid1.TextMatrix(Grid1.Row, 0)
    txtNumCartao.Text = Grid1.TextMatrix(Grid1.Row, 1)
    txtValor.Text = Replace(Grid1.TextMatrix(Grid1.Row, 2), "R$ ", "")
    MaskData.Text = Grid1.TextMatrix(Grid1.Row, 3)
    txtDescricao.Text = Grid1.TextMatrix(Grid1.Row, 5)

    
    Dim statusTexto As String
    statusTexto = Trim(Grid1.TextMatrix(Grid1.Row, 4))

    
    Dim i As Integer
    Dim encontrado As Boolean
    encontrado = False

    For i = 0 To cboStatus.ListCount - 1

        If Trim(cboStatus.List(i)) = statusTexto Then
            cboStatus.ListIndex = i
            encontrado = True
            Exit For
        End If
    Next i

    If Not encontrado Then
        MsgBox "Status não encontrado na ComboBox!", vbExclamation, "Informação"
    End If
    
    If cboStatus.Text = "Aprovado" Then
        cmdAlterar.Visible = False
        cmdGravar.Visible = False
        cmdExcluir.Visible = False
        MsgBox "Alterações não são permitidas para transações com status 'Aprovado'.", vbInformation, "Aviso"
    End If
    
     SSTab1.Tab = 0
        
     Screen.MousePointer = vbNormal
     Exit Sub
TratarErro:
    detalhesErro = "Número do erro: " & Err.Number & " - Descrição do erro: " & IIf(Err.Description = "", "Descrição indisponível", Err.Description)
    RegistrarErro "Erro na subrotina Grid1_DblClick: " & detalhesErro
    MsgBox "Ocorreu um erro. Consulte o arquivo de log para mais detalhes.", vbExclamation, "Erro"
    Screen.MousePointer = vbNormal
End Sub






Private Sub MontarStatus()
    On Error GoTo erro
    Dim qr1 As ADODB.Recordset
    Set qr1 = New ADODB.Recordset

    With lstStatus
        .Clear
        qr1.Open "SELECT id_status, status FROM Status", conn, adOpenStatic, adLockReadOnly
        If Not (qr1.BOF And qr1.EOF) Then
            Do While Not qr1.EOF
                If Not IsNull(qr1!Status) And Not IsNull(qr1!id_status) Then
                    .AddItem qr1!Status
                    .ItemData(.NewIndex) = qr1!id_status
                End If
                qr1.MoveNext
            Loop
        End If
        qr1.Close
    End With

    Exit Sub

erro:
    MsgBox "Ocorreu um erro: " & Err.Description, vbCritical, "Erro"
    If Not qr1 Is Nothing Then qr1.Close
    Set qr1 = Nothing
End Sub






Private Sub txtValor_KeyPress(KeyAscii As Integer)
    KeyAscii = PermitirSoNumero(KeyAscii, ",", True)
End Sub

Private Sub txtValor_LostFocus()
    txtValor.Text = ValorReal(txtValor.Text)
End Sub

Private Sub Limpar()
On Error GoTo TratarErro

    txtNumCartao.Text = Empty
    txtValor.Text = Empty
    txtDescricao.Text = Empty
    MaskData.Mask = Format(Now, "dd/mm/yyyy")
    cboStatus.ListIndex = 0
    txtIdTransacao.Text = Empty

    Exit Sub

TratarErro:
    detalhesErro = "Número do erro: " & Err.Number & " - Descrição do erro: " & IIf(Err.Description = "", "Descrição indisponível", Err.Description)
    RegistrarErro "Erro na subrotina Limpar: " & detalhesErro
    MsgBox "Ocorreu um erro. Consulte o arquivo de log para mais detalhes.", vbExclamation, "Erro"
End Sub


Private Sub LimparPesquisa()
On Error GoTo TratarErro
    
    txtNumCartaoP.Text = Empty
    txtValorP1.Text = Empty
    txtValorP2.Text = Empty
    txtDescricaoP.Text = Empty
    MaskDataP1.Mask = "__/__/____"
    MaskDataP2.Mask = "__/__/____"
    txtIdTransacaoP.Text = Empty
    For i = 0 To lstStatus.ListCount - 1
        lstStatus.Selected(i) = True
    Next
    
    Exit Sub
    
TratarErro:
    detalhesErro = "Número do erro: " & Err.Number & " - Descrição do erro: " & IIf(Err.Description = "", "Descrição indisponível", Err.Description)
    RegistrarErro "Erro na subrotina LimparPesquisa: " & detalhesErro
    MsgBox "Ocorreu um erro. Consulte o arquivo de log para mais detalhes.", vbExclamation, "Erro"
    
End Sub

Private Function ValidaDados() As Boolean
    Dim valor As Double
    ValidaDados = True


        If txtNumCartao.Text = "" Or txtNumCartao.Text = " " Or Len(txtNumCartao.Text) <> 16 Or Not IsNumeric(txtNumCartao.Text) Then
            MsgBox "Insira um número do cartão valido!", vbInformation, "Informação"
            txtNumCartao.SetFocus
            Screen.MousePointer = vbNormal
            ValidaDados = False
            Exit Function
        End If
        
        If Not IsNumeric(txtValor.Text) Or txtValor.Text = "" Or txtValor.Text = "0,00" Then
            MsgBox "Insira o valor correto!", vbInformation, "Informação"
            txtValor.SetFocus
            Screen.MousePointer = vbNormal
            ValidaDados = False
            Exit Function
        End If
        
        
        valor = CDbl(txtValor.Text)
        
        If valor <= 0 Then
            MsgBox "O valor não pode ser negativo.", vbExclamation, "Informação"
            txtValor.SetFocus
            Screen.MousePointer = vbNormal
            ValidaDados = False
            Exit Function
        End If
        
        If cboStatus.ListIndex = 0 Then
            MsgBox "Selecione um Status", vbInformation, "Informação"
            cboStatus.SetFocus
            Screen.MousePointer = vbNormal
            ValidaDados = False
            Exit Function
        End If
        

End Function

Private Function ValidaDadosPesquisa() As Integer
    Dim valor1 As Double
    Dim valor2 As Double
    Dim ValidaData As Integer
    ValidaData = 0
    
        If Not IsNumeric(txtIdTransacaoP.Text) Or txtIdTransacaoP.Text = "" Or txtIdTransacaoP.Text = "0" Then
            ValidaDadosPesquisa = ValidaDadosPesquisa + 1
        End If

        If txtNumCartaoP.Text = "" Or txtNumCartaoP.Text = " " Or Len(txtNumCartaoP.Text) <> 16 Or Not IsNumeric(txtNumCartaoP.Text) Then
            ValidaDadosPesquisa = ValidaDadosPesquisa + 1
        End If
        
        If txtDescricaoP.Text = "" Then
            ValidaDadosPesquisa = ValidaDadosPesquisa + 1
        End If
        
        If Not IsNumeric(txtValorP1.Text) Or txtValorP1.Text = "" Or txtValorP1.Text = "0,00" Then
            ValidaDadosPesquisa = ValidaDadosPesquisa + 1
        End If
        
        If txtValorP1.Text = "" Then
            txtValorP1.Text = "0,00"
            valor1 = CDbl(txtValorP1.Text)
        End If
        
        If txtValorP2.Text = "" Then
            txtValorP2.Text = "0,00"
            valor2 = CDbl(txtValorP2.Text)
        End If
        
        If (valor1 <= 0 Or valor2 <= 0) Or (valor2 <= valor1) Then
            ValidaDadosPesquisa = ValidaDadosPesquisa + 1
        End If
        
        
            ''''
            If (Trim(MaskDataP1.Text) <> "__/__/____" And Trim(MaskDataP2.Text) = "__/__/____") Or _
               (Trim(MaskDataP2.Text) <> "__/__/____" And Trim(MaskDataP1.Text) = "__/__/____") Or _
               (Trim(MaskDataP1.Text) = "__/__/____" And Trim(MaskDataP2.Text) = "__/__/____") Then
                ValidaData = 1
            End If
            ' Verifique se MaskDataP1 não tem uma data maior que MaskDataP2
            If Trim(MaskDataP1.Text) <> "__/__/____" And Trim(MaskDataP2.Text) <> "__/__/____" Then
                Dim dataP1 As Date
                Dim dataP2 As Date
                
                ' Converter para datas
                On Error Resume Next
                dataP1 = CDate(MaskDataP1.Text)
                dataP2 = CDate(MaskDataP2.Text)
        
                If dataP1 > dataP2 Then
                    MsgBox "A data inicial não pode ser maior que a data final.", vbExclamation, "Pesquisa"
                   ValidaData = 1
                End If
            End If
            
            ''''
        
        ValidaDadosPesquisa = ValidaDadosPesquisa + ValidaData
        
        If ValidaDadosPesquisa = 6 Then
            MsgBox "Pelo menos um filtro deve ser utilizado para realizar a pesquisa!", vbInformation, "Informação"
        End If

End Function


Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Carregarcombo()
    Dim SQL As String
    Dim rs As ADODB.Recordset
    

    ConectarBanco
       
    SQL = "SELECT ID_status, status FROM Status"
       
       
    Set rs = New ADODB.Recordset
    rs.Open SQL, conn
       
    cboStatus.Clear
    
    cboStatus.AddItem "Todos"
    cboStatus.ItemData(cboStatus.ListCount - 1) = -1
       
      
    If Not rs.EOF Then
        
        Do While Not rs.EOF
            cboStatus.AddItem rs.Fields("status").Value
            cboStatus.ItemData(cboStatus.ListCount - 1) = rs.Fields("ID_status").Value
            rs.MoveNext
        Loop
    Else
        MsgBox "Dados não encontrados.", vbInformation, "Informação"
    End If
    cboStatus.ListIndex = 0
       
    rs.Close
    Set rs = Nothing
    
End Sub


Private Sub Form_Load()
    

    If Me.MDIChild Then
        ' Calcula a posição central
        Me.Left = (MDIForm1.ScaleWidth - Me.Width) / 2
        Me.Top = (MDIForm1.ScaleHeight - Me.Height) / 2
    End If
  
    Carregarcombo
    paginaAtual = 1
    MaskData.Text = Format(Now, "dd/mm/yyyy")
    MontarStatus
    
    For i = 0 To lstStatus.ListCount - 1
        lstStatus.Selected(i) = True
    Next

    
    
End Sub


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
    Dim rs As ADODB.Recordset
    Dim inicio As Integer
    Dim FiltroStatus As String
    Dim totalRegistros As Integer

    
    ConectarBanco

    
    registrosPorPagina = 5 ' Coloquei 5 para poder ter como testar devido ao baixo volume de dados.
    inicio = (paginaAtual - 1) * registrosPorPagina

    
    SQL = "SELECT COUNT(*) AS Total FROM VW_Transacao WHERE Id_Transacao > 0"

    If Trim(txtIdTransacaoP.Text) <> "" Then
        SQL = SQL & " AND Id_Transacao = " & txtIdTransacaoP.Text
    End If

    If IsNumeric(txtNumCartaoP.Text) Then
        SQL = SQL & " AND Numero_Cartao = " & txtNumCartaoP.Text
    End If

    If Len(txtDescricaoP.Text) > 0 Then
        SQL = SQL & " AND Descricao LIKE '%" & txtDescricaoP.Text & "%'"
    End If

    FiltroStatus = ""
    If lstStatus.SelCount > 0 Then
        For i = 0 To lstStatus.ListCount - 1
            If lstStatus.Selected(i) Then
                FiltroStatus = FiltroStatus & lstStatus.ItemData(i) & ","
            End If
        Next
        If Len(FiltroStatus) > 0 Then
            FiltroStatus = Left(FiltroStatus, Len(FiltroStatus) - 1) ' Remove vírgula extra
            SQL = SQL & " AND id_Status IN (" & FiltroStatus & ")"
        End If
    End If

    If txtValorP1.Text <> "0,00" Then
        SQL = SQL & " AND Valor_Transacao BETWEEN " & fnValorSQL(txtValorP1.Text) & " AND " & fnValorSQL(txtValorP2.Text)
    End If

    If MaskDataP1.Text <> "__/__/____" Then
        SQL = SQL & " AND Data_Transacao BETWEEN '" & DataSQL(MaskDataP1.Text) & "' AND '" & DataSQL(MaskDataP2.Text) & "'"
    End If

    
    Set rsTotal = New ADODB.Recordset
    rsTotal.Open SQL, conn, adOpenStatic, adLockReadOnly
    If Not rsTotal.EOF Then
        totalRegistros = rsTotal.Fields("Total").Value
    End If
    rsTotal.Close
    Set rsTotal = Nothing

    
    SQLPesquisa = "SELECT Id_Transacao, Numero_Cartao, Valor_Transacao, Data_Transacao, Status, Descricao FROM ( " & _
                  "SELECT *, ROW_NUMBER() OVER (ORDER BY Id_Transacao) AS RowNum FROM VW_Transacao WHERE Id_Transacao > 0"

    If Trim(txtIdTransacaoP.Text) <> "" Then
        SQLPesquisa = SQLPesquisa & " AND Id_Transacao = " & txtIdTransacaoP.Text
    End If

    If IsNumeric(txtNumCartaoP.Text) Then
        SQLPesquisa = SQLPesquisa & " AND Numero_Cartao = " & txtNumCartaoP.Text
    End If

    If Len(txtDescricaoP.Text) > 0 Then
        SQLPesquisa = SQLPesquisa & " AND Descricao LIKE '%" & txtDescricaoP.Text & "%'"
    End If

    FiltroStatus = ""
    If lstStatus.SelCount > 0 Then
        For i = 0 To lstStatus.ListCount - 1
            If lstStatus.Selected(i) Then
                FiltroStatus = FiltroStatus & lstStatus.ItemData(i) & ","
            End If
        Next
        If Len(FiltroStatus) > 0 Then
            FiltroStatus = Left(FiltroStatus, Len(FiltroStatus) - 1) ' Remove vírgula extra
            SQLPesquisa = SQLPesquisa & " AND id_Status IN (" & FiltroStatus & ")"
        End If
    End If

    If txtValorP1.Text <> "0,00" Then
        SQLPesquisa = SQLPesquisa & " AND Valor_Transacao BETWEEN " & fnValorSQL(txtValorP1.Text) & " AND " & fnValorSQL(txtValorP2.Text)
    End If

    If MaskDataP1.Text <> "__/__/____" Then
        SQLPesquisa = SQLPesquisa & " AND Data_Transacao BETWEEN '" & DataSQL(MaskDataP1.Text) & "' AND '" & DataSQL(MaskDataP2.Text) & "'"
    End If

    SQLPesquisa = SQLPesquisa & " ) AS T WHERE T.RowNum BETWEEN " & (inicio + 1) & " AND " & (inicio + registrosPorPagina)

    
    Debug.Print SQLPesquisa
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open SQLPesquisa, conn, adOpenStatic, adLockReadOnly

  
    If rs.EOF Then
        MsgBox "Nenhum registro encontrado.", vbInformation
        Exit Sub
    End If

    
    Grid1.Rows = rs.RecordCount + 1
    Grid1.Cols = rs.Fields.Count

    Dim tamanhosColunas As Variant
    tamanhosColunas = Array(1000, 2100, 1500, 1800, 1200, 2500)

    For i = LBound(tamanhosColunas) To UBound(tamanhosColunas)
        Grid1.ColWidth(i) = tamanhosColunas(i)
    Next i

    Dim cabeçalho As Variant
    cabeçalho = Array("ID", "NÚMERO CARTÃO", "VALOR", "DATA", "STATUS", "DESCRIÇÃO")

    For i = LBound(cabeçalho) To UBound(cabeçalho)
        Grid1.TextMatrix(0, i) = cabeçalho(i)
    Next i

    i = 1
    Do While Not rs.EOF
        For j = 0 To rs.Fields.Count - 1
            If j = 2 Then
                Grid1.TextMatrix(i, j) = Format(rs.Fields(j).Value, "R$ #,##0.00")
                Grid1.ColAlignment(j) = flexAlignRightCenter
            Else
                Grid1.TextMatrix(i, j) = rs.Fields(j).Value
            End If
        Next j
        rs.MoveNext
        i = i + 1
    Loop

    totalPaginas = (totalRegistros + registrosPorPagina - 1) \ registrosPorPagina
    If totalPaginas < 1 Then totalPaginas = 1
    
    frmPesquisa.Visible = False
    SSTab1 = 0

    lblPagina.Caption = "Página " & paginaAtual & " de " & totalPaginas
End Sub


'Private Sub ExportarParaExcel()
'    On Error GoTo TratarErro
'
'    ' Declaração de variáveis
'    Dim ExcelApp As Object
'    Dim Workbook As Object
'    Dim Worksheet As Object
'    Dim i As Integer, j As Integer
'
'    ' Inicia o Excel
'    Set ExcelApp = CreateObject("Excel.Application")
'    Set Workbook = ExcelApp.Workbooks.Add
'    Set Worksheet = Workbook.Sheets(1)
'
'    ' Copia cabeçalhos para o Excel
'    For j = 0 To Grid1.Cols - 1
'        Worksheet.Cells(1, j + 1).Value = Grid1.TextMatrix(0, j) ' Cabeçalhos
'    Next j
'
'    ' Copia dados da Grid para o Excel
'    For i = 1 To Grid1.Rows - 1
'        For j = 0 To Grid1.Cols - 1
'            Worksheet.Cells(i + 1, j + 1).Value = Grid1.TextMatrix(i, j) ' Dados
'        Next j
'    Next i
'
'    ' Ajusta a largura das colunas automaticamente
'    Worksheet.Columns("A:Z").AutoFit
'
'    ' Salva o arquivo Excel
'    Dim CaminhoArquivo As String
'    CaminhoArquivo = App.Path & "\Exportacao_Grid.xlsx" ' Salvar no diretório da aplicação
'    Workbook.SaveAs CaminhoArquivo
'
'    ' Exibe mensagem de sucesso
'    MsgBox "Exportação concluída com sucesso!" & vbCrLf & "Arquivo salvo em: " & CaminhoArquivo, vbInformation
'
'    ' Finaliza objetos
'    Workbook.Close
'    ExcelApp.Quit
'    Set Worksheet = Nothing
'    Set Workbook = Nothing
'    Set ExcelApp = Nothing
'    Exit Sub
'
'TratarErro:
'    MsgBox "Erro ao exportar para Excel: " & Err.Description, vbCritical
'    If Not ExcelApp Is Nothing Then ExcelApp.Quit
'    Set Worksheet = Nothing
'    Set Workbook = Nothing
'    Set ExcelApp = Nothing
'End Sub



Sub ExportarParaExcel()
    Dim excelApp As Object
    Dim excelBook As Object
    Dim excelSheet As Object
    Dim i As Integer, j As Integer
    Dim rsExport As ADODB.Recordset
    Dim SQLExport As String
    Dim filePath As String

    ' Conectando ao banco
    ConectarBanco
    
    ' Carregando todos os registros
    SQLExport = "SELECT Id_Transacao, Numero_Cartao, Valor_Transacao, Data_Transacao, Status, Descricao FROM VW_Transacao WHERE Id_Transacao > 0"
    
    Set rsExport = New ADODB.Recordset
    rsExport.Open SQLExport, conn, adOpenStatic, adLockReadOnly
    
    If rsExport.EOF Then
        MsgBox "Nenhum registro encontrado para exportar.", vbInformation
        Exit Sub
    End If
    
    ' Usando Common Dialog para escolher o caminho do arquivo
    On Error Resume Next
    CommonDialog1.DialogTitle = "Salvar Arquivo Excel"
    CommonDialog1.Filter = "Arquivos do Excel (*.xlsx)|*.xlsx"
    CommonDialog1.FileName = "DadosExportados.xlsx"
    CommonDialog1.ShowSave
    
    filePath = CommonDialog1.FileName
    If filePath = "" Then
        MsgBox "Exportação cancelada pelo usuário.", vbInformation
        Exit Sub
    End If
    
    ' Iniciando o Excel
    Set excelApp = CreateObject("Excel.Application")
    Set excelBook = excelApp.Workbooks.Add
    Set excelSheet = excelBook.Sheets(1)
    
    ' Criando Cabeçalhos
    For j = 0 To rsExport.Fields.Count - 1
        excelSheet.Cells(1, j + 1).Value = rsExport.Fields(j).Name
    Next j
    
    ' Preenchendo os Dados
    i = 2
    Do While Not rsExport.EOF
        For j = 0 To rsExport.Fields.Count - 1
            If j = 1 Then ' Segunda coluna (número do cartão)
                excelSheet.Cells(i, j + 1).Value = "'" & rsExport.Fields(j).Value ' Força o formato texto
            Else
                excelSheet.Cells(i, j + 1).Value = rsExport.Fields(j).Value
            End If
        Next j
        rsExport.MoveNext
        i = i + 1
    Loop
    
    ' Salvando o Arquivo
    excelBook.SaveAs filePath
    excelApp.Quit
    
    rsExport.Close
    Set rsExport = Nothing
    MsgBox "Dados exportados para Excel com sucesso no arquivo: " & filePath, vbInformation
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


