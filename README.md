# Projeto: Geração de Contratos em PDF

## Descrição do Projeto

Este projeto tem como objetivo automatizar a geração de contratos em PDF utilizando dados de clientes armazenados em uma planilha Excel. A macro em VBA lê informações como CPF/CNPJ, telefone, e-mail e data de inclusão de uma planilha e preenche um modelo de contrato, gerando um arquivo PDF para cada cliente.

## Estrutura do Projeto

### 1. Arquivos Utilizados

- **Dados.xlsx**: Planilha contendo os dados dos clientes, incluindo:
  - CPF/CNPJ
  - Nome
  - Endereço
  - Bairro
  - Estado
  - CEP
  - Cidade
  - Telefone
  - E-mail
  - Data de Inclusão adicionando 15 dias a data do Registro

- **Modelo.xlsx**: Modelo do contrato que contém placeholders para os dados dos clientes:
  - `[CNPJ]`
  - `[TELEFONE]`
  - `[EMAIL]`
  - `[RAZAO_SOCIAL]`
  - `[ENDERECO]`
  - `[BAIRRO]`
  - `[ESTADO]`
  - `[CEP]`
  - `[CIDADE]`
  - `[DATA]`

### 2. Resultado Esperado

- Geração de arquivos PDF com o contrato preenchido para cada cliente, salvos na pasta `Contratos/`.

## Funcionalidades da Macro VBA

A macro `GerarContratos()` implementa as seguintes funcionalidades:

1. **Leitura de Dados**: Abre a planilha `Dados.xlsx` e lê as informações a partir da segunda linha (ignorando o cabeçalho).

2. **Substituição de Placeholders**: Para cada cliente, substitui os placeholders no modelo de contrato pelos dados correspondentes.

3. **Geração de Nome do Arquivo**: Gera o nome do arquivo PDF a partir do CPF/CNPJ, removendo caracteres inválidos.

4. **Verificação de Existência de Arquivos**: Antes de salvar, verifica se um arquivo com o mesmo nome já existe e modifica o nome, se necessário.

5. **Exportação para PDF**: Exporta a nova cópia do modelo como um arquivo PDF, utilizando o método `ExportAsFixedFormat`.

6. **Notificação**: Exibe uma mensagem informando que os contratos foram gerados com sucesso.

7. **Exclusão de Campo**: Se o campo "Cidade" estiver vazio, não exibe "N/D" no contrato.

## Código da Macro com Geração de Arquivos que contém CPF

```vba
Sub GerarContratos()
    Dim wbDados As Workbook
    Dim wbModelo As Workbook
    Dim wbNovo As Workbook
    Dim wsDados As Worksheet
    Dim wsModelo As Worksheet
    Dim i As Integer
    Dim clienteCNPJ As String
    Dim clienteTelefone As String
    Dim clienteEmail As String
    Dim clienteEndereco As String
    Dim clienteBairro As String
    Dim clienteCEP As String
    Dim clienteCidade As String
    Dim clienteEstado As String
    Dim clienteRazaoSocial As String
    Dim clienteData As String
    Dim outputPath As String
    Dim fileName As String
    Dim pdfFileName As String

    ' Abra os arquivos de dados e modelo
    Set wbDados = Workbooks.Open(ThisWorkbook.Path & "\Dados.xlsx")
    Set wbModelo = Workbooks.Open(ThisWorkbook.Path & "\Modelo.xlsx")

    ' Defina as planilhas
    Set wsDados = wbDados.Sheets(1) ' Primeira aba de Dados
    Set wsModelo = wbModelo.Sheets(1) ' Primeira aba do Modelo

    ' Caminho para salvar os contratos
    outputPath = ThisWorkbook.Path & "\Contratos\"
    If Dir(outputPath, vbDirectory) = "" Then MkDir outputPath

    ' Loop através dos clientes
    For i = 2 To wsDados.Cells(wsDados.Rows.Count, 1).End(xlUp).Row
        clienteCNPJ = wsDados.Cells(i, 1).Value
        clienteTelefone = wsDados.Cells(i, 6).Value
        clienteEmail = wsDados.Cells(i, 7).Value
        clienteEndereco = wsDados.Cells(i, 9).Value
        clienteBairro = wsDados.Cells(i, 3).Value
        clienteCEP = wsDados.Cells(i, 8).Value
        clienteCidade = wsDados.Cells(i, 4).Value
        clienteEstado = wsDados.Cells(i, 5).Value
        clienteRazaoSocial = wsDados.Cells(i, 2).Value
        clienteData = wsDados.Cells(i, 10).Text

        ' Limpa CNPJ e Telefone
        fileName = Replace(clienteCNPJ, ".", "")
        fileName = Replace(fileName, "-", "")
        fileName = Replace(fileName, "/", "")
        
        clienteTelefone = Replace(clienteTelefone, "-", "")
        clienteTelefone = Replace(clienteTelefone, " ", "")
        
        ' Verifica se o CNPJ não está vazio
        If fileName <> "ND" And fileName <> "" Then
            ' Cria uma nova cópia do modelo
            wsModelo.Copy
            Set wbNovo = ActiveWorkbook ' A nova planilha copiada será o workbook ativo

            ' Substitui os placeholders no novo workbook
            With wbNovo.Sheets(1)
                .Cells.Replace "[CNPJ]", clienteCNPJ
                .Cells.Replace "[TELEFONE]", clienteTelefone
                .Cells.Replace "[EMAIL]", clienteEmail
                .Cells.Replace "[ENDERECO]", clienteEndereco
                .Cells.Replace "[BAIRRO]", clienteBairro
                .Cells.Replace "[CEP]", clienteCEP
                ' Substituir Cidade apenas se não for "N/D"
                If clienteCidade <> "N/D" Then
                    .Cells.Replace "[CIDADE]", clienteCidade
                Else
                    .Cells.Replace "[CIDADE]", ""
                End If
                .Cells.Replace "[ESTADO]", clienteEstado
                .Cells.Replace "[RAZAO_SOCIAL]", clienteRazaoSocial
                .Cells.Replace "[DATA]", Format(clienteData, "dd/mm/yyyy")
            End With

            ' Define o nome do arquivo PDF
            pdfFileName = outputPath & "Contrato_" & fileName & ".pdf"

            ' Verifica se o arquivo já existe e modifica o nome, se necessário
            Dim fileCounter As Integer
            fileCounter = 1
            While Dir(pdfFileName) <> ""
                pdfFileName = outputPath & "Contrato_" & fileName & "_" & fileCounter & ".pdf"
                fileCounter = fileCounter + 1
            Wend

            ' Exporta o workbook como PDF
            wbNovo.ExportAsFixedFormat Type:=xlTypePDF, fileName:=pdfFileName, Quality:=xlQualityStandard, _
                IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
            
            ' Fecha o novo workbook sem salvar alterações
            wbNovo.Close SaveChanges:=False
        End If
    Next i

    ' Fecha os arquivos originais
    wbDados.Close SaveChanges:=False
    wbModelo.Close SaveChanges:=False

    MsgBox "Contratos gerados com sucesso!", vbInformation
End Sub
```
## Geração de arquivos gerais (Independente se tem ou não CPF na planilha)

```vba

Sub GerarContratos()
    Dim wbDados As Workbook
    Dim wbModelo As Workbook
    Dim wbNovo As Workbook
    Dim wsDados As Worksheet
    Dim wsModelo As Worksheet
    Dim i As Integer
    Dim clienteCNPJ As String
    Dim clienteTelefone As String
    Dim clienteEmail As String
    Dim clienteEndereco As String
    Dim clienteBairro As String
    Dim clienteCEP As String
    Dim clienteCidade As String
    Dim clienteEstado As String
    Dim clienteRazaoSocial As String
    Dim clienteData As String
    Dim outputPath As String
    Dim fileName As String
    Dim pdfFileName As String

    ' Abra os arquivos de dados e modelo
    Set wbDados = Workbooks.Open(ThisWorkbook.Path & "\Dados.xlsx")
    Set wbModelo = Workbooks.Open(ThisWorkbook.Path & "\Modelo.xlsx")

    ' Defina as planilhas
    Set wsDados = wbDados.Sheets(1) ' Primeira aba de Dados
    Set wsModelo = wbModelo.Sheets(1) ' Primeira aba do Modelo

    ' Caminho para salvar os contratos
    outputPath = ThisWorkbook.Path & "\Contratos\"
    If Dir(outputPath, vbDirectory) = "" Then MkDir outputPath

    ' Loop através dos clientes
    For i = 2 To wsDados.Cells(wsDados.Rows.Count, 1).End(xlUp).Row
        clienteCNPJ = wsDados.Cells(i, 1).Value
        clienteTelefone = wsDados.Cells(i, 6).Value
        clienteEmail = wsDados.Cells(i, 7).Value
        clienteEndereco = wsDados.Cells(i, 9).Value
        clienteBairro = wsDados.Cells(i, 3).Value
        clienteCEP = wsDados.Cells(i, 8).Value
        clienteCidade = wsDados.Cells(i, 4).Value
        clienteEstado = wsDados.Cells(i, 5).Value
        clienteRazaoSocial = wsDados.Cells(i, 2).Value
        clienteData = wsDados.Cells(i, 10).Text

        ' Limpa CNPJ e Telefone
        fileName = Replace(clienteCNPJ, ".", "")
        fileName = Replace(fileName, "-", "")
        fileName = Replace(fileName, "/", "")
        
        clienteTelefone = Replace(clienteTelefone, "-", "")
        clienteTelefone = Replace(clienteTelefone, " ", "")
        
        ' Cria uma nova cópia do modelo
        wsModelo.Copy
        Set wbNovo = ActiveWorkbook ' A nova planilha copiada será o workbook ativo

        ' Substitui os placeholders no novo workbook
        With wbNovo.Sheets(1)
            .Cells.Replace "[CNPJ]", clienteCNPJ
            .Cells.Replace "[TELEFONE]", clienteTelefone
            .Cells.Replace "[EMAIL]", clienteEmail
            .Cells.Replace "[ENDERECO]", clienteEndereco
            .Cells.Replace "[BAIRRO]", clienteBairro
            .Cells.Replace "[CEP]", clienteCEP
            ' Substituir Cidade apenas se não for "N/D"
            If clienteCidade <> "N/D" Then
                .Cells.Replace "[CIDADE]", clienteCidade
            Else
                .Cells.Replace "[CIDADE]", ""
            End If
            .Cells.Replace "[ESTADO]", clienteEstado
            .Cells.Replace "[RAZAO_SOCIAL]", clienteRazaoSocial
            .Cells.Replace "[DATA]", Format(clienteData, "dd/mm/yyyy")
        End With

        ' Define o nome do arquivo PDF
        If fileName = "" Then
            fileName = "Contrato_Cliente_" & i ' Nome alternativo se CNPJ estiver vazio
        End If
        pdfFileName = outputPath & "Contrato_" & fileName & ".pdf"

        ' Verifica se o arquivo já existe e modifica o nome, se necessário
        Dim fileCounter As Integer
        fileCounter = 1
        While Dir(pdfFileName) <> ""
            pdfFileName = outputPath & "Contrato_" & fileName & "_" & fileCounter & ".pdf"
            fileCounter = fileCounter + 1
        Wend

        ' Exporta o workbook como PDF
        wbNovo.ExportAsFixedFormat Type:=xlTypePDF, fileName:=pdfFileName, Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
        
        ' Fecha o novo workbook sem salvar alterações
        wbNovo.Close SaveChanges:=False
    Next i

    ' Fecha os arquivos originais
    wbDados.Close SaveChanges:=False
    wbModelo.Close SaveChanges:=False

    MsgBox "Contratos gerados com sucesso!", vbInformation
End Sub
