  '主函数
Sub fun()
    
    Dim projectno As String, s0, s1, s2, s3, s4
    Dim excelobject As Object, wb As Object, r As Long, i As Long, arr
    
    '获取word中的数据
    projectno = ActiveDocument
    
    '启动Excel程序并打开对应的Excel表
    Set excelobject = CreateObject("excel.application")
    excelobject.Visible = False   '不可见
    Set wb = excelobject.Workbooks.Open("H:\list.xlsx")
    
    s0 = "FIRSTNAME##"
    s1 = "LASTNAME##"
    s2 = "COMPANY##"
    s3 = "COUNTRY##"
    s4 = "PRESENTATIONTITLE##"
    s5 = "PCIM"
    
    
    '循环获取Excel中的每行数据，并在word中的指定的位置替换掉对应的内容
    For i = 2 To 10
        
        x1 = wb.Sheets("contributor info").Range("a" & i) 'First Name
        x2 = wb.Sheets("contributor info").Range("b" & i) 'Last Name
        x3 = wb.Sheets("contributor info").Range("c" & i) 'Organization
        x4 = wb.Sheets("contributor info").Range("d" & i) 'Country
        x5 = wb.Sheets("contributor info").Range("e" & i) 'Title of the Presentation
    
        '调用“查找和替换的代码”
        Call find(UCase(s0), UCase(x1))
        Call find(UCase(s1), UCase(x2))
        Call find(UCase(s2), UCase(x3))
        Call find(UCase(s3), UCase(x4))
        Call find(UCase(s4), UCase(x5))
        Call find(s5, s5)
        
        ActiveDocument.SaveAs FileName:="liu" & i & ".doc" '生成新的word页面
        
        '需要注意，在word里面自动变成了大写，所以需要转格式
        
        '调用“查找和替换的代码”
        Call find_(UCase(x5), UCase(s4))
        Call find_(UCase(x4), UCase(s3))
        Call find_(UCase(x3), UCase(s2))
        Call find_(UCase(x2), UCase(s1))
        Call find_(UCase(x1), UCase(s0))
        
        Next
    
    '测试对应的数据
    'MsgBox (x2)
    excelobject.Quit
End Sub

  '查找并替换函数1
Sub find(x1, x2)
   'Selection.HomeKey Unit:=wdLine
       With Selection.find
        .Text = x1
        .Replacement.Text = x2
        .Forward = True '向下查找
        .Wrap = 1
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
  Selection.find.Execute Replace:=wdReplaceOne
  
End Sub
  '查找并替换函数2
Sub find_(x1, x2)
   'Selection.HomeKey Unit:=wdLine
       With Selection.find
        .Text = x1
        .Replacement.Text = x2
        .Forward = False '向上查找
        .Wrap = 1
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
  Selection.find.Execute Replace:=wdReplaceOne
  
End Sub








