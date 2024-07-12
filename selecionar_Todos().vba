Sub selecionar_Todos()
'
' Select all
'
Sheets("1103").Range("B:F").AutoFilter Field:=5 'FILTER PRIORITY VIEW ALL
Sheets("1109").Range("B:F").AutoFilter Field:=5 'FILTER PRIORITY VIEW ALL
    
Call moverMenu
    
End Sub

Sub Selecionar_Vazios()

Sheets("1103").Range("B:F").AutoFilter Field:=5, Criteria1:="=" 'FILTER PRIORITY VIEW EMPTY
Sheets("1109").Range("B:F").AutoFilter Field:=5, Criteria1:="=" 'FILTER PRIORITY VIEW EMPTY
    
Call moverMenu

End Sub

Sub prioridade_0()

Sheets("1103").Range("B:F").AutoFilter Field:=4, Criteria1:="0" 'FILTER PRIORITY ZERO
Sheets("1109").Range("B:F").AutoFilter Field:=4, Criteria1:="0" 'FILTER PRIORITY ZERO

Call moverMenu

End Sub

Sub prioridade_1()

Sheets("1103").Range("B:F").AutoFilter Field:=4, Criteria1:="1" 'FILTER PRIORITY ONE
Sheets("1109").Range("B:F").AutoFilter Field:=4, Criteria1:="1" 'FILTER PRIORITY ONE

Call moverMenu

End Sub

Sub prioridade_2()

Sheets("1103").Range("B:F").AutoFilter Field:=4, Criteria1:="2" 'FILTER PRIORITY TWO
Sheets("1109").Range("B:F").AutoFilter Field:=4, Criteria1:="2" 'FILTER PRIORITY TWO

Call moverMenu

End Sub

Sub prioridade_3()

Sheets("1103").Range("B:F").AutoFilter Field:=4, Criteria1:="3" 'FILTER PRIORITY THREE
Sheets("1109").Range("B:F").AutoFilter Field:=4, Criteria1:="3" 'FILTER PRIORITY THREE

Call moverMenu

End Sub

Sub prioridade_4()

Sheets("1103").Range("B:F").AutoFilter Field:=4, Criteria1:="4" 'FILTER PRIORITY FOUR
Sheets("1109").Range("B:F").AutoFilter Field:=4, Criteria1:="4" 'FILTER PRIORITY FOUR

Call moverMenu

End Sub

Sub prioridade_5()

Sheets("1103").Range("B:F").AutoFilter Field:=4, Criteria1:="5" 'FILTER PRIORITY FIVE
Sheets("1109").Range("B:F").AutoFilter Field:=4, Criteria1:="5" 'FILTER PRIORITY FIVE

Call moverMenu

End Sub

Sub prioridade_X()

Sheets("1103").Range("B:F").AutoFilter Field:=4, Criteria1:="??" 'FILTER PRIORITY URGENT
Sheets("1109").Range("B:F").AutoFilter Field:=4, Criteria1:="??" 'FILTER PRIORITY URGENT

Call moverMenu

End Sub

Sub prioridade_tds()

Sheets("1103").Range("B:F").AutoFilter Field:=4 'FILTER PRIORITY ALL FILTERS
Sheets("1109").Range("B:F").AutoFilter Field:=4 'FILTER PRIORITY ALL FILTERS

Call moverMenu

End Sub
