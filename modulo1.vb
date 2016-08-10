Sub matrizRecebeNota()
cmdMatrizRecebe.Show

End Sub

Sub atualizaNotas()
'este algoritmo atualiza as notas na planilha CONTROLE_GERAL

Dim loja, nota As Integer 'declara variável loja e nota

Dim notaJaAdicionada As Boolean ' declara variável notaJadicionada
Dim ultCelula As String

'vai para planilha NOTAS DETALHE
Sheets("NOTAS_DETALHE").Select
Range("A1").Select
ultCelula = ActiveCell.AddressLocal 'guarda o endereço da ult celula usada

Do While (ActiveCell.Offset(1, 0) <> "")
    
    Range(ultCelula).Activate 'ativa ultima célula
    ActiveCell.Offset(1, 0).Select 'pula uma célula
    ultCelula = ActiveCell.AddressLocal 'atualiza o endereço da ult celula selecionada
    
    'se na planilha NOTA_DETALHE, coluna M (Geral), não estiver escrito SIM, o programa verificará se a nota já está referenciada
    'no controle geral, caso esteja, adiciona o sim, para aumentar o desempenho na próxima atualização
    'se a nota não estiver na planilha CONTROLE_GERAL, adiciona-a, e coloca o SIM (o SIM significa que a nota já está na no CONTROLE_GERAL)
    
    'início IF, caso esteja sem o SIM
    If ActiveCell.Offset(0, 12).FormulaR1C1 = "" Then
    
        loja = ActiveCell.FormulaR1C1
        nota = ActiveCell.Offset(0, 1).FormulaR1C1
        data = ActiveCell.Offset(0, 2).FormulaR1C1
        rds = ActiveCell.Offset(0, 3).FormulaR1C1
        posicao = ActiveCell.Address 'local da célula verificada
    
        notaJaAdicionada = False 'supõe que a nota fiscal não foi adicionada
    
        Sheets("CONTROLE_GERAL").Select
        Range("a1").Select
    
        'compara a nota fiscal com planilha CONTROLE_GERAL
        Do While (ActiveCell.Offset(1, 0).FormulaR1C1 <> "")
            ActiveCell.Offset(1, 0).Select
            If loja = ActiveCell.FormulaR1C1 And nota = ActiveCell.Offset(0, 1).FormulaR1C1 Then
                notaJaAdicionada = True
                Exit Do ' se ja houver nota sai do loop
            End If
        Loop
    
        ActiveCell.Offset(1, 0).Select
        
        'adiciona loja e nota à planilha controle geral
        If (notaJaAdicionada = False) Then
        
            ActiveCell.FormulaR1C1 = loja
            ActiveCell.Offset(0, 1).FormulaR1C1 = nota
            ActiveCell.Offset(0, 2).FormulaR1C1 = data
            ActiveCell.Offset(0, 3).FormulaR1C1 = rds
            
            Sheets("NOTAS_DETALHE").Select
            'coloca sim na coluna geral, isso serve para aumentar o desempenho do programa
            Range(posicao).Select
            ActiveCell.Offset(0, 12).FormulaR1C1 = "Sim"
        
        End If
        
        If notaJaAdicionada = True Then
            Sheets("NOTAS_DETALHE").Select
            'coloca sim na coluna geral
            Range(posicao).Select
            ActiveCell.Offset(0, 12).FormulaR1C1 = "Sim"
        End If
    
    End If
    'fim IF do SIM
    Sheets("NOTAS_DETALHE").Select
    
Loop

Sheets("CONTROLE_GERAL").Select
Range("a1").Select

End Sub

Sub receberItem()
    frmRecItem.Show
    
End Sub

Sub NFControle1()
    addNoControle ("A7")
End Sub

Sub NFControle2()
    addNoControle ("A8")
End Sub

Sub NFControle3()
    addNoControle ("A9")
End Sub

Sub NFControle4()
    addNoControle ("A10")
End Sub



Function addNoControle(ind)

Dim notaRepete As Boolean

Sheets("CAPA_EQUIP").Select
Range(ind).Select
 lj = ActiveCell.Offset(0, 1).FormulaR1C1
 nf = ActiveCell.Offset(0, 2).FormulaR1C1
 dt = ActiveCell.Offset(0, 3).FormulaR1C1
 rds = ActiveCell.Offset(0, 4).FormulaR1C1
 item = ActiveCell.Offset(0, 6).Value
 descricao = ActiveCell.Offset(0, 7).Value
 serie = ActiveCell.Offset(0, 8).Value
 pat = ActiveCell.Offset(0, 9).Value
 notaHyperlink = lj + "." + nf
'Application.Visible = True
'Workbooks.Application.Visible = True



Sheets("CONTROLE_GERAL").Select
Range("a1").Select

notaRepete = False
'compara com as notas
Do While (ActiveCell.Offset(1, 0).FormulaR1C1 <> "")
        ActiveCell.Offset(1, 0).Select
       
        If lj = ActiveCell.FormulaR1C1 And nf = ActiveCell.Offset(0, 1).FormulaR1C1 Then
            notaRepete = True
            Exit Do ' se ja houver nota sai do loop
        End If
Loop

If notaRepete = False Then
    ActiveCell.Offset(1, 0).Select
    ActiveCell.FormulaR1C1 = lj
    'adiciona hyperlink da nota
    ActiveCell.Offset(0, 1).Select
    ActiveCell.FormulaR1C1 = nf
    ActiveCell.Hyperlinks.Add Anchor:=Selection, Address:="NFenviadas\" + notaHyperlink + ".pdf"
    ActiveCell.Offset(0, 1).FormulaR1C1 = dt
    ActiveCell.Offset(0, 2).FormulaR1C1 = rds
    
End If

Sheets("NOTAS_DETALHE").Select
Range("a1").Select
'procura o ultimo campo
Do While (ActiveCell.Offset(1, 0).FormulaR1C1 <> "")
        ActiveCell.Offset(1, 0).Select
Loop
ActiveCell.Offset(1, 0).Select
ActiveCell.FormulaR1C1 = lj
ActiveCell.Offset(0, 1).FormulaR1C1 = nf
ActiveCell.Offset(0, 2).FormulaR1C1 = dt
ActiveCell.Offset(0, 3).FormulaR1C1 = rds
ActiveCell.Offset(0, 4).FormulaR1C1 = item
ActiveCell.Offset(0, 5).FormulaR1C1 = descricao
ActiveCell.Offset(0, 6).FormulaR1C1 = serie
ActiveCell.Offset(0, 7).FormulaR1C1 = pat
Sheets("CAPA_EQUIP").Select
Range(ind).Select
MsgBox ("REGISTRO EFETUADO")

End Function

Sub auto_Open()
   MsgBox ("Contato: lazarofdsilva@gmail.com")
   Sheets("CAPA_EQUIP").Select
End Sub
