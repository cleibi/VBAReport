Sub Report()
'
' Report Macro
'
' Keyboard Shortcut: Ctrl+r
'
    
    '
    ' Deleta linhas e colunas desnecessarias
    '
    Cells(Rows.Count, 9).End(xlUp).EntireRow.Delete
    Rows("1:3").Delete
    Columns("A:H").Delete
    Columns("C").Delete
    Columns("E:F").Delete
    Columns("F:H").Delete
    Columns("G").Delete
    Columns("H:R").Delete
    Rows(1).EntireRow.Delete
    
    '
    ' Deleta as linhas sem "arrive at"
    '

    Dim lRow As Long
    Dim iCntr As Long
    
    '
    ' *****ATENCAO***** Mudar o numero de linhas de acordo
    '

    lRow = 1000
    For iCntr = lRow To 1 Step -1
        If Trim(Cells(iCntr, 1)) = “” Then
            Rows(iCntr).Delete
        End If
    Next

    '
    ' Converte data para texto
    '

    Columns("A:A").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlFixedWidth, _
        OtherChar:="-", FieldInfo:=Array(0, 2), TrailingMinusNumbers:=True

    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("A:A").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="/", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1)), _
        TrailingMinusNumbers:=True
    Columns("D:D").Select
    Selection.Delete Shift:=xlToLeft

    Columns("B").Delete

    '
    ' *****ATENCAO***** Mudar o numero de linhas de acordo, se necessario, na linha abaixo. Esse numero tem que ser maior do que numero de linhas no arquivo. O numero a ser mudado e o numero que segue a letra H
    '
    Range("A1:H1000").Select
    For i = Selection.Rows.Count To 1 Step -1
        '
        ' *****ATENCAO***** Mudar o ANO aqui, se necessario, na linha abaixo. E o numero entre aspas duplas depois de .Value <>
        '
        If Cells(i, 2).Value <> "2022" Then
            Cells(i, 2).EntireRow.Delete
         End If
    Next i

    '
    ' *****ATENCAO***** Mudar o numero de linhas de acordo, se necessario, na linha abaixo. Esse numero tem que ser maior do que numero de linhas no arquivo. O numero a ser mudado e o numero que segue a letra H
    '
    Range("A1:H1000").Select
    For i = Selection.Rows.Count To 1 Step -1
        '
        ' *****ATENCAO***** Mudar o MES aqui, na linha abaixo. E o numero entre aspas duplas depois de .Value <>
        '
        If Cells(i, 1).Value <> "2" Then
            Cells(i, 1).EntireRow.Delete
        End If
    Next i

    Columns("B").Delete
    Columns("A").Delete

    '
    ' Insere os textos na coluna H
    '

    Columns("H:H").ColumnWidth = 26.5
    Columns("I:I").ColumnWidth = 17.38
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "Numero de Novos Clientes"
    Range("H3").Select
    ActiveCell.FormulaR1C1 = "Residencia"
    Range("I3").Select
    ActiveCell.FormulaR1C1 = "Qtde"
    Range("H4").Select
    ActiveCell.FormulaR1C1 = "Cidadão canadense"
    Range("H5").Select
    ActiveCell.FormulaR1C1 = "PR"
    Range("H6").Select
    ActiveCell.FormulaR1C1 = "Processo de PR"
    Range("H7").Select
    ActiveCell.FormulaR1C1 = "Work Permit"
    Range("H8").Select
    ActiveCell.FormulaR1C1 = "Outros"
    Range("H10").Select
    ActiveCell.FormulaR1C1 = "Renda Annual"
    Range("I10").Select
    ActiveCell.FormulaR1C1 = "Qtde"
    Range("H11").Select
    ActiveCell.FormulaR1C1 = "$40-60 mil"
    Range("H12").Select
    ActiveCell.FormulaR1C1 = "$60-80 mil"
    Range("H13").Select
    ActiveCell.FormulaR1C1 = "$80-100 mil"
    Range("H14").Select
    ActiveCell.FormulaR1C1 = "$100-150 mil"
    Range("H15").Select
    ActiveCell.FormulaR1C1 = "$150-200 mil"
    Range("H16").Select
    ActiveCell.FormulaR1C1 = "'+ $200 mil"
    Range("H18").Select
    ActiveCell.FormulaR1C1 = "Downpayment"
    Range("I18").Select
    ActiveCell.FormulaR1C1 = "Qtde"
    Range("H19").Select
    ActiveCell.FormulaR1C1 = "Até $10 mil"
    Range("H20").Select
    ActiveCell.FormulaR1C1 = "$10-20 mil"
    Range("H21").Select
    ActiveCell.FormulaR1C1 = "$20-30 mil"
    Range("H22").Select
    ActiveCell.FormulaR1C1 = "$30-40 mil"
    Range("H23").Select
    ActiveCell.FormulaR1C1 = "$40-50 mil"
    Range("H24").Select
    ActiveCell.FormulaR1C1 = "$50-60 mil"
    Range("H25").Select
    ActiveCell.FormulaR1C1 = "'+ $60 mil"
    Range("H26").Select
    ActiveCell.FormulaR1C1 = "Quando"
    Range("I26").Select
    ActiveCell.FormulaR1C1 = "Qtde"
    Range("H27").Select
    ActiveCell.FormulaR1C1 = "Agora (1-3 meses)"
    Range("H28").Select
    ActiveCell.FormulaR1C1 = "Futuro breve (6-12 meses))"
    Range("H29").Select
    ActiveCell.FormulaR1C1 = "Quero apenas informações"
    Range("H31").Select
    ActiveCell.FormulaR1C1 = "Origem"
    Range("I31").Select
    ActiveCell.FormulaR1C1 = "Qtde"
    Range("H32").Select
    ActiveCell.FormulaR1C1 = "Indicação"
    Range("H33").Select
    ActiveCell.FormulaR1C1 = "Facebook"
    Range("H34").Select
    ActiveCell.FormulaR1C1 = "Instagram"
    Range("H35").Select
    ActiveCell.FormulaR1C1 = "Blogs, sites e conteúdos online"
    Range("H36").Select
    ActiveCell.FormulaR1C1 = "Palestras e lives"
    Range("H37").Select
    ActiveCell.FormulaR1C1 = "Google"
    Range("H38").Select
    ActiveCell.FormulaR1C1 = "Youtube"
    Range("H40").Select
    ActiveCell.FormulaR1C1 = "Origem - High Stakes"
    Range("I40").Select
    ActiveCell.FormulaR1C1 = "Qtde"
    Range("I41").Select
    ActiveCell.FormulaR1C1 = "Qtde"
    Range("H41").Select
    ActiveCell.FormulaR1C1 = "Indicação"
    Range("H42").Select
    ActiveCell.FormulaR1C1 = "Facebook"
    Range("H43").Select
    ActiveCell.FormulaR1C1 = "Instagram"
    Range("H44").Select
    ActiveCell.FormulaR1C1 = "Blogs, sites e conteúdos online"
    Range("H45").Select
    ActiveCell.FormulaR1C1 = "Palestras e lives"
    Range("H46").Select
    ActiveCell.FormulaR1C1 = "Google"
    Range("H47").Select
    ActiveCell.FormulaR1C1 = "Youtube"
    Range("H49").Select
    ActiveCell.FormulaR1C1 = "Origem - Medium Stakes"
    Range("I49").Select
    ActiveCell.FormulaR1C1 = "Qtde"
    Range("H50").Select
    ActiveCell.FormulaR1C1 = "Indicação"
    Range("H51").Select
    ActiveCell.FormulaR1C1 = "Facebook"
    Range("H52").Select
    ActiveCell.FormulaR1C1 = "Instagram"
    Range("H53").Select
    ActiveCell.FormulaR1C1 = "Blogs, sites e conteúdos online"
    Range("H54").Select
    ActiveCell.FormulaR1C1 = "Palestras e lives"
    Range("H55").Select
    ActiveCell.FormulaR1C1 = "Google"
    Range("H56").Select
    ActiveCell.FormulaR1C1 = "Youtube"
    Range("H58").Select
    ActiveCell.FormulaR1C1 = "Origem - LowStakes"
    Range("I58").Select
    ActiveCell.FormulaR1C1 = "Qtde"
    Range("H59").Select
    ActiveCell.FormulaR1C1 = "Indicação"
    Range("H60").Select
    ActiveCell.FormulaR1C1 = "Facebook"
    Range("H61").Select
    ActiveCell.FormulaR1C1 = "Instagram"
    Range("H62").Select
    ActiveCell.FormulaR1C1 = "Blogs, sites e conteúdos online"
    Range("H63").Select
    ActiveCell.FormulaR1C1 = "Palestras e lives"
    Range("H64").Select
    ActiveCell.FormulaR1C1 = "Google"
    Range("H65").Select
    ActiveCell.FormulaR1C1 = "Youtube"

    '
    ' Insere os valores de Residencia
    '

    Range("I4").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(C[-7],RC[-1])"
    Range("I5").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(C[-7],RC[-1])"
    Range("I6").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(C[-7],RC[-1])"
    Range("I7").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(C[-7],RC[-1])"
    Range("I8").Select

    '
    ' Insere os valores de Renda Anual
    '

    ActiveCell.FormulaR1C1 = "=COUNTIF(C[-7],RC[-1])"
    Range("I11").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(C[-6],RC[-1])"
    Range("I12").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(C[-6],RC[-1])"
    Range("I13").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(C[-6],RC[-1])"
    Range("I14").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(C[-6],RC[-1])"
    Range("I15").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(C[-6],RC[-1])"
    Range("I16").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(C[-6],""'+ $200 mil"")"

    '
    ' Insere os valores de Downpayment
    '

    Range("I19").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(C[-5],RC[-1])"
    Range("I20").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(C[-5],RC[-1])"
    Range("I21").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(C[-5],RC[-1])"
    Range("I22").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(C[-5],RC[-1])"
    Range("I23").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(C[-5],RC[-1])"
    Range("I24").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(C[-5],RC[-1])"
    Range("I25").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(C[-5],""'+ $60 mil"")"

    '
    ' Insere os valores de Quando
    '

    Range("I27").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(C[-4],RC[-1])"
    Range("I28").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(C[-4],RC[-1])"
    Range("I29").Select

    '
    ' Insere os valores de Origem
    '

    ActiveCell.FormulaR1C1 = "=COUNTIF(C[-4],RC[-1])"
    Range("I32").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(C[-3],RC[-1])"
    Range("I33").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(C[-3],RC[-1])"
    Range("I34").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(C[-3],RC[-1])"
    Range("I35").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(C[-3],RC[-1])"
    Range("I36").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(C[-3],RC[-1])"
    Range("I37").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(C[-3],RC[-1])"
    Range("I38").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(C[-3],RC[-1])"

    '
    ' Insere os valores de Origem - High Stakes
    '

    Range("I41").Select
    ActiveCell.FormulaR1C1 = _
        "=COUNTIFS(C[-7],R[-37]C[-1],C[-6],R[-27]C[-1],C[-3],RC[-1]) + COUNTIFS(C[-7],R[-37]C[-1],C[-6],R[-26]C[-1],C[-3],RC[-1]) + COUNTIFS(C[-7],R[-37]C[-1],C[-6],""'+ $200 mil"",C[-3],RC[-1]) + COUNTIFS(C[-7],R[-36]C[-1],C[-6],R[-27]C[-1],C[-3],RC[-1]) + COUNTIFS(C[-7],R[-36]C[-1],C[-6],R[-26]C[-1],C[-3],RC[-1]) + COUNTIFS(C[-7],R[-36]C[-1],C[-6],""'+ $200 mil"",C[-3],RC[" & _
        "-1]) + COUNTIFS(C[-7],R[-35]C[-1],C[-6],R[-27]C[-1],C[-3],RC[-1]) + COUNTIFS(C[-7],R[-35]C[-1],C[-6],R[-26]C[-1],C[-3],RC[-1]) + COUNTIFS(C[-7],R[-35]C[-1],C[-6],""'+ $200 mil"",C[-3],RC[-1])" & _
        ""
    Range("I42").Select
    ActiveCell.FormulaR1C1 = _
        "=COUNTIFS(C[-7],R[-38]C[-1],C[-6],R[-28]C[-1],C[-3],RC[-1]) + COUNTIFS(C[-7],R[-38]C[-1],C[-6],R[-27]C[-1],C[-3],RC[-1]) + COUNTIFS(C[-7],R[-38]C[-1],C[-6],""'+ $200 mil"",C[-3],RC[-1]) + COUNTIFS(C[-7],R[-37]C[-1],C[-6],R[-28]C[-1],C[-3],RC[-1]) + COUNTIFS(C[-7],R[-37]C[-1],C[-6],R[-27]C[-1],C[-3],RC[-1]) + COUNTIFS(C[-7],R[-37]C[-1],C[-6],""'+ $200 mil"",C[-3],RC[" & _
        "-1]) + COUNTIFS(C[-7],R[-36]C[-1],C[-6],R[-28]C[-1],C[-3],RC[-1]) + COUNTIFS(C[-7],R[-36]C[-1],C[-6],R[-27]C[-1],C[-3],RC[-1]) + COUNTIFS(C[-7],R[-36]C[-1],C[-6],""'+ $200 mil"",C[-3],RC[-1])" & _
        ""
    Range("I43").Select
    ActiveCell.FormulaR1C1 = _
        "=COUNTIFS(C[-7],R[-39]C[-1],C[-6],R[-29]C[-1],C[-3],RC[-1]) + COUNTIFS(C[-7],R[-39]C[-1],C[-6],R[-28]C[-1],C[-3],RC[-1]) + COUNTIFS(C[-7],R[-39]C[-1],C[-6],""'+ $200 mil"",C[-3],RC[-1]) + COUNTIFS(C[-7],R[-38]C[-1],C[-6],R[-29]C[-1],C[-3],RC[-1]) + COUNTIFS(C[-7],R[-38]C[-1],C[-6],R[-28]C[-1],C[-3],RC[-1]) + COUNTIFS(C[-7],R[-38]C[-1],C[-6],""'+ $200 mil"",C[-3],RC[" & _
        "-1]) + COUNTIFS(C[-7],R[-37]C[-1],C[-6],R[-29]C[-1],C[-3],RC[-1]) + COUNTIFS(C[-7],R[-37]C[-1],C[-6],R[-28]C[-1],C[-3],RC[-1]) + COUNTIFS(C[-7],R[-37]C[-1],C[-6],""'+ $200 mil"",C[-3],RC[-1])" & _
        ""
    Range("I44").Select
    ActiveCell.FormulaR1C1 = _
        "=COUNTIFS(C[-7],R[-40]C[-1],C[-6],R[-30]C[-1],C[-3],RC[-1]) + COUNTIFS(C[-7],R[-40]C[-1],C[-6],R[-29]C[-1],C[-3],RC[-1]) + COUNTIFS(C[-7],R[-40]C[-1],C[-6],""'+ $200 mil"",C[-3],RC[-1]) + COUNTIFS(C[-7],R[-39]C[-1],C[-6],R[-30]C[-1],C[-3],RC[-1]) + COUNTIFS(C[-7],R[-39]C[-1],C[-6],R[-29]C[-1],C[-3],RC[-1]) + COUNTIFS(C[-7],R[-39]C[-1],C[-6],""'+ $200 mil"",C[-3],RC[" & _
        "-1]) + COUNTIFS(C[-7],R[-38]C[-1],C[-6],R[-30]C[-1],C[-3],RC[-1]) + COUNTIFS(C[-7],R[-38]C[-1],C[-6],R[-29]C[-1],C[-3],RC[-1]) + COUNTIFS(C[-7],R[-38]C[-1],C[-6],""'+ $200 mil"",C[-3],RC[-1])" & _
        ""
    Range("I45").Select
    ActiveCell.FormulaR1C1 = _
        "=COUNTIFS(C[-7],R[-41]C[-1],C[-6],R[-31]C[-1],C[-3],RC[-1]) + COUNTIFS(C[-7],R[-41]C[-1],C[-6],R[-30]C[-1],C[-3],RC[-1]) + COUNTIFS(C[-7],R[-41]C[-1],C[-6],""'+ $200 mil"",C[-3],RC[-1]) + COUNTIFS(C[-7],R[-40]C[-1],C[-6],R[-31]C[-1],C[-3],RC[-1]) + COUNTIFS(C[-7],R[-40]C[-1],C[-6],R[-30]C[-1],C[-3],RC[-1]) + COUNTIFS(C[-7],R[-40]C[-1],C[-6],""'+ $200 mil"",C[-3],RC[" & _
        "-1]) + COUNTIFS(C[-7],R[-39]C[-1],C[-6],R[-31]C[-1],C[-3],RC[-1]) + COUNTIFS(C[-7],R[-39]C[-1],C[-6],R[-30]C[-1],C[-3],RC[-1]) + COUNTIFS(C[-7],R[-39]C[-1],C[-6],""'+ $200 mil"",C[-3],RC[-1])" & _
        ""
    Range("I46").Select
    ActiveCell.FormulaR1C1 = _
        "=COUNTIFS(C[-7],R[-42]C[-1],C[-6],R[-32]C[-1],C[-3],RC[-1]) + COUNTIFS(C[-7],R[-42]C[-1],C[-6],R[-31]C[-1],C[-3],RC[-1]) + COUNTIFS(C[-7],R[-42]C[-1],C[-6],""'+ $200 mil"",C[-3],RC[-1]) + COUNTIFS(C[-7],R[-41]C[-1],C[-6],R[-32]C[-1],C[-3],RC[-1]) + COUNTIFS(C[-7],R[-41]C[-1],C[-6],R[-31]C[-1],C[-3],RC[-1]) + COUNTIFS(C[-7],R[-41]C[-1],C[-6],""'+ $200 mil"",C[-3],RC[" & _
        "-1]) + COUNTIFS(C[-7],R[-40]C[-1],C[-6],R[-32]C[-1],C[-3],RC[-1]) + COUNTIFS(C[-7],R[-40]C[-1],C[-6],R[-31]C[-1],C[-3],RC[-1]) + COUNTIFS(C[-7],R[-40]C[-1],C[-6],""'+ $200 mil"",C[-3],RC[-1])" & _
        ""
    Range("I47").Select
    ActiveCell.FormulaR1C1 = _
        "=COUNTIFS(C[-7],R[-43]C[-1],C[-6],R[-33]C[-1],C[-3],RC[-1]) + COUNTIFS(C[-7],R[-43]C[-1],C[-6],R[-32]C[-1],C[-3],RC[-1]) + COUNTIFS(C[-7],R[-43]C[-1],C[-6],""'+ $200 mil"",C[-3],RC[-1]) + COUNTIFS(C[-7],R[-42]C[-1],C[-6],R[-33]C[-1],C[-3],RC[-1]) + COUNTIFS(C[-7],R[-42]C[-1],C[-6],R[-32]C[-1],C[-3],RC[-1]) + COUNTIFS(C[-7],R[-42]C[-1],C[-6],""'+ $200 mil"",C[-3],RC[" & _
        "-1]) + COUNTIFS(C[-7],R[-41]C[-1],C[-6],R[-33]C[-1],C[-3],RC[-1]) + COUNTIFS(C[-7],R[-41]C[-1],C[-6],R[-32]C[-1],C[-3],RC[-1]) + COUNTIFS(C[-7],R[-41]C[-1],C[-6],""'+ $200 mil"",C[-3],RC[-1])" & _
        ""
    
    '
    ' Insere os valores de Origem - Medium Stakes
    '

    Range("I50").Select
    ActiveCell.FormulaR1C1 = _
        "=COUNTIFS(C[-6],R[-39]C[-1],C[-3],RC[-1]) + COUNTIFS(C[-6],R[-38]C[-1],C[-3],RC[-1]) + COUNTIFS(C[-6],R[-37]C[-1],C[-3],RC[-1])"
    Range("I51").Select
    ActiveCell.FormulaR1C1 = _
        "=COUNTIFS(C[-6],R[-40]C[-1],C[-3],RC[-1]) + COUNTIFS(C[-6],R[-39]C[-1],C[-3],RC[-1]) + COUNTIFS(C[-6],R[-38]C[-1],C[-3],RC[-1])"
    Range("I52").Select
    ActiveCell.FormulaR1C1 = _
        "=COUNTIFS(C[-6],R[-41]C[-1],C[-3],RC[-1]) + COUNTIFS(C[-6],R[-40]C[-1],C[-3],RC[-1]) + COUNTIFS(C[-6],R[-39]C[-1],C[-3],RC[-1])"
    Range("I53").Select
    ActiveCell.FormulaR1C1 = _
        "=COUNTIFS(C[-6],R[-42]C[-1],C[-3],RC[-1]) + COUNTIFS(C[-6],R[-41]C[-1],C[-3],RC[-1]) + COUNTIFS(C[-6],R[-40]C[-1],C[-3],RC[-1])"
    Range("I54").Select
    ActiveCell.FormulaR1C1 = _
        "=COUNTIFS(C[-6],R[-43]C[-1],C[-3],RC[-1]) + COUNTIFS(C[-6],R[-42]C[-1],C[-3],RC[-1]) + COUNTIFS(C[-6],R[-41]C[-1],C[-3],RC[-1])"
    Range("I55").Select
    ActiveCell.FormulaR1C1 = _
        "=COUNTIFS(C[-6],R[-44]C[-1],C[-3],RC[-1]) + COUNTIFS(C[-6],R[-43]C[-1],C[-3],RC[-1]) + COUNTIFS(C[-6],R[-42]C[-1],C[-3],RC[-1])"
    Range("I56").Select
    ActiveCell.FormulaR1C1 = _
        "=COUNTIFS(C[-6],R[-45]C[-1],C[-3],RC[-1]) + COUNTIFS(C[-6],R[-44]C[-1],C[-3],RC[-1]) + COUNTIFS(C[-6],R[-43]C[-1],C[-3],RC[-1])"

    '
    ' Insere os valores de Origem - Low Stakes
    '

    Range("I59").Select
    ActiveCell.FormulaR1C1 = "=COUNTIFS(C[-6],R[-48]C[-1],C[-3],RC[-1])"
    Range("I60").Select
    ActiveCell.FormulaR1C1 = "=COUNTIFS(C[-6],R[-49]C[-1],C[-3],RC[-1])"
    Range("I61").Select
    ActiveCell.FormulaR1C1 = "=COUNTIFS(C[-6],R[-50]C[-1],C[-3],RC[-1])"
    Range("I62").Select
    ActiveCell.FormulaR1C1 = "=COUNTIFS(C[-6],R[-51]C[-1],C[-3],RC[-1])"
    Range("I63").Select
    ActiveCell.FormulaR1C1 = "=COUNTIFS(C[-6],R[-52]C[-1],C[-3],RC[-1])"
    Range("I64").Select
    ActiveCell.FormulaR1C1 = "=COUNTIFS(C[-6],R[-53]C[-1],C[-3],RC[-1])"
    Range("I65").Select
    ActiveCell.FormulaR1C1 = "=COUNTIFS(C[-6],R[-54]C[-1],C[-3],RC[-1])"

    '
    ' Coloring the headers
    '

    Range("H1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("H3:I3").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("H10:I10").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("H18:I18").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("H26:I26").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("H31:I31").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("H40:I40").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("H49:I49").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("H58:I58").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

    Range("I1").Select
   
'
End Sub