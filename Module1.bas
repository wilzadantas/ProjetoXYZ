Attribute VB_Name = "Module1"
Global conn As ADODB.Connection
Global rs As ADODB.Recordset
Global strConn As String



Function fnValorSQL(valor As String) As String

    valor = Replace(valor, ".", "")
    valor = Replace(valor, ",", ".")
    fnValorSQL = valor

End Function

Function DataSQL(strData As String, Optional ByVal aspas As Boolean = False) As String

    If aspas Then
        DataSQL = "'" & Format(strData, "yyyymmdd") & "'"
    Else
    
        DataSQL = Format(strData, "yyyymmdd")
    End If
End Function

Sub RegistrarErro(emDetalhes As String)
    On Error Resume Next

    Dim dataHoraAtual As String
    dataHoraAtual = Format(Now, "yyyy-mm-dd hh:nn:ss")

    Dim caminhoArquivoLog As String
    caminhoArquivoLog = "C:\Logs\ErroLog.txt" ' Altere o caminho conforme necessário

    
    Dim arquivoLog As Integer
    arquivoLog = FreeFile
    Open caminhoArquivoLog For Append As arquivoLog
    Print #arquivoLog, dataHoraAtual & " - " & emDetalhes
    Close arquivoLog
End Sub

Public Function PermitirSoNumero(ByVal qKeyAscii As Integer, Optional Exceção As String, Optional TrocarPontoPorVirgula As Boolean) As Integer
    
    If TrocarPontoPorVirgula Then
        If qKeyAscii = 46 Then qKeyAscii = 44
    End If
    
    If Exceção = "" Then
        If Not Chr(qKeyAscii) Like "[0-9]" And qKeyAscii <> vbKeyReturn And qKeyAscii <> vbKeyBack Then
            PermitirSoNumero = 0
        Else
            PermitirSoNumero = qKeyAscii
        End If
    Else
        If Not Chr(qKeyAscii) Like "[0-9" & Exceção & "]" And qKeyAscii <> vbKeyReturn And qKeyAscii <> vbKeyBack Then
            PermitirSoNumero = 0
        Else
            PermitirSoNumero = qKeyAscii
        End If
    End If
End Function




Public Function ValorReal(qValor As String) As String
  Dim qC As Integer
  Dim qC2 As Integer
  If Trim(qValor) <> "" Then
    For qC = 1 To Len(qValor)
      If Mid(qValor, qC, 1) = "," Then
        qC2 = qC2 + 1
        If qC2 >= 2 Then
          qValor = "0,00"
          Exit For
        End If
      End If
    Next
  Else
    qValor = 0
  End If
  ValorReal = Format(qValor, "##,##0.00")
End Function
