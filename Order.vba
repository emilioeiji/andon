Private Sub UserForm_Initialize()
    Dim ws As Worksheet
    Dim i As Integer
    Dim imgPath As String
    Dim imgControl As Object
    Dim txt1 As Object, txt2 As Object, txt3 As Object
    Dim topPos As Single, leftPos As Single
    Dim colCount As Integer
    Dim lastRow As Long
    Dim items() As Variant
    Dim temp As Variant
    Dim j As Integer
    
    Set ws = ThisWorkbook.Sheets("Itens")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row ' Conta a quantidade de itens
    
    ' Carregar os itens em um array
    ReDim items(1 To lastRow - 1, 1 To 4) ' Dimensões do array para armazenar os dados
    
    For i = 2 To lastRow
        items(i - 1, 1) = ws.Cells(i, 1).Value ' Coluna A (nome do item)
        items(i - 1, 2) = ws.Cells(i, 2).Value ' Coluna B
        items(i - 1, 3) = ws.Cells(i, 3).Value ' Coluna C
        items(i - 1, 4) = ws.Cells(i, 4).Value ' Coluna D (caminho da imagem)
    Next i
    
    ' Ordenar o array pelo nome do item (primeira coluna)
    For i = 1 To UBound(items, 1) - 1
        For j = i + 1 To UBound(items, 1)
            If items(i, 1) > items(j, 1) Then
                ' Troca as linhas
                temp = items(i, 1)
                items(i, 1) = items(j, 1)
                items(j, 1) = temp
                
                temp = items(i, 2)
                items(i, 2) = items(j, 2)
                items(j, 2) = temp
                
                temp = items(i, 3)
                items(i, 3) = items(j, 3)
                items(j, 3) = temp
                
                temp = items(i, 4)
                items(i, 4) = items(j, 4)
                items(j, 4) = temp
            End If
        Next j
    Next i
    
    ' Preparar o formulário para apresentar os itens ordenados
    colCount = 0
    topPos = 10
    leftPos = 10
    
    ' Configurar o Frame para permitir scroll
    With Me.Frame1
        .ScrollBars = fmScrollBarsBoth
        .ScrollHeight = ((lastRow \ 10) + 1) * 200 ' Ajusta a altura do scroll
        .ScrollWidth = ((lastRow \ 10) + 1) * 150 ' Ajusta a largura do scroll
    End With
    
    ' Adicionar os itens ao UserForm na ordem correta
    For i = 1 To UBound(items, 1)
        If (i - 1) Mod 10 = 0 And i > 1 Then
            colCount = colCount + 1
            topPos = 10
            leftPos = leftPos + 150
        End If
        
        ' Adicionar Imagem
        imgPath = items(i, 4)
        Set imgControl = Me.Frame1.Controls.Add("Forms.Image.1", "img" & i)
        With imgControl
            .Picture = LoadPicture(imgPath)
            .Top = topPos
            .Left = leftPos
            .Width = 100
            .Height = 100
        End With
        
        ' Adicionar TextBox1
        Set txt1 = Me.Frame1.Controls.Add("Forms.TextBox.1", "txt1_" & i)
        With txt1
            .Top = topPos + 110
            .Left = leftPos
            .Width = 100
            .Text = items(i, 1) ' Nome do item (Coluna A)
        End With
        
        ' Adicionar TextBox2
        Set txt2 = Me.Frame1.Controls.Add("Forms.TextBox.1", "txt2_" & i)
        With txt2
            .Top = topPos + 140
            .Left = leftPos
            .Width = 100
            .Text = items(i, 2) ' Descrição (Coluna B)
        End With
        
        ' Adicionar TextBox3
        Set txt3 = Me.Frame1.Controls.Add("Forms.TextBox.1", "txt3_" & i)
        With txt3
            .Top = topPos + 170
            .Left = leftPos
            .Width = 100
            .Text = items(i, 3) ' Preço (Coluna C)
        End With
        
        topPos = topPos + 200
    Next i
    
    ' Ajustar a largura do scroll após adicionar todos os itens
    Me.Frame1.ScrollWidth = (colCount + 1) * 150
End Sub
