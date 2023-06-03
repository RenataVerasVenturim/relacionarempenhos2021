This document exist just to GitHub understand VBA language in the project

Sub atualizar()

Sheets("EMPENHO2021").Select

sublinharate = Cells(10450, 1).End(xlUp).Row + 1

Range(Cells(2, 1), Cells(sublinharate, 16)).Clear


For Each ABA In ThisWorkbook.Sheets
If ABA.Name <> "EMPENHO2021" Then

ABA.Select

    linha = Cells(10000, 1).End(xlUp).Row + 1
    
    Range(Cells(2, 1), Cells(linha, 16)).Copy

    Sheets("EMPENHO2021").Select

    proximalinha = Cells(10000, 1).End(xlUp).Row + 1
    Cells(proximalinha, 1).PasteSpecial xlPasteValues
Application.CutCopyMode = False
        
    
End If

Next

MsgBox ("Listagem atualizada com sucesso!")

End Sub

