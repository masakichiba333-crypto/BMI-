Attribute VB_Name = "Module1"
Private Sub cmdEnd_Click()
    frmBMI.Hide         'ユーザーフォームを隠す'
End Sub

Private Sub cmdHantei_Click()
    
    Dim sngSintyo As Single       '身長[cm]'
    Dim sngTaizyu As Single       '体重[kg]'
    Dim sngBmi As Single          'BMI'
    Dim strHantei As String       '判定結果'
    
    sngSintyo = Val(TxtSintyo.Text)
    sngTaizyu = Val(TxtTaizyu.Text)
    
    
    'BMIの計算'
    sngBmi = sngTaizyu / (sngSintyo * 0.01) ^ 2
    
    'BMIの判定'
    Select Case sngBmi
        Case Is >= 40: strHantei = "肥満度4"
        Case 35 To 40: strHantei = "肥満度3"
        Case 30 To 35: strHantei = "肥満度2"
        Case 25 To 30: strHantei = "肥満度1"
        Case 18.5 To 25: strHantei = "標準"
        Case Is < 18.5: strHantei = "やせ"
    End Select
        
    
    'BMIの表示'
    TxtBmi.Text = Format(sngBmi, "00.00")

    '判定結果の表示'
    TxtHantei.Text = Format(strHantei)

End Sub



Private Sub cmdEnd_Click()
        
        frmSeirekiwareki.Hide   'ユーザーフォームを隠す'
        
End Sub

Private Sub TxtSeireki_Change()

    Dim intSeireki As Integer     '西暦'
    Dim intWareki As Integer      '和暦'
    Dim strWareki As String       '元号'
    Const strGannenn As String = "元年"        '元年'
    
    
    'テキストボックスの文字列を数値化して変数に代入'
    intSeireki = Val(TxtSeireki.Text)
    
    '西暦を和暦に変換'
    If intSeireki < 1868 Then  '計算範囲外'
        TxtWareki.Text = "計算範囲外"
    ElseIf intSeireki = 1868 Or intSeireki = 1912 Or intSeireki = 1926 Or intSeireki = 1989 Or intSeireki = 2019 Then  '元年の計算'
        If intSeireki = 1868 Then
            strWareki = "明治"
        ElseIf intSeireki = 1912 Then
            strWareki = "大正"
        ElseIf intSeireki = 1926 Then
            strWareki = "昭和"
        ElseIf intSeireki = 1989 Then
            strWareki = "平成"
        ElseIf intSeireki = 2019 Then
            strWareki = "令和"
        End If
        TxtWareki = strWareki & strGannenn
    
    Else
        If intSeireki <= 1911 Then
            strWareki = "明治"
            intWareki = intSeireki - 1867
        ElseIf intSeireki <= 1925 Then
            strWareki = "大正"
            intWareki = intSeireki - 1911
        ElseIf intSeireki <= 1988 Then
            strWareki = "昭和"
            intWareki = intSeireki - 1925
        ElseIf intSeireki <= 2018 Then
            strWareki = "平成"
            intWareki = intSeireki - 1988
        Else
            strWareki = "令和"
            intWareki = intSeireki - 2018
        End If
        TxtWareki = strWareki & intWareki & "年"
        
    End If
    
    
End Sub

Private Sub UserForm_Click()

End Sub

