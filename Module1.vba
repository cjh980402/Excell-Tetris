'(3,33):플레이 시간, (8,33):레벨, (8,35):경험치(1줄에 10%)
'(10,33):점수, (12,33):콤보, (14,33):최고점수, (14,34):최고점수 때 플레이시간
'(3,42):도형 행, (3,43):도형 열, (3,44):도형 종류, (3,45):회전 종류, (3,46):그림자 행
'(3,47):쌓인 블록의 최상단 행, (5,42):게임 재생 여부
'(5,43):다음 호출 예약 시간, (5,44):다음 블록, (5,45):보관 블록, (5,46):줄 올라온 시간

Sub Selectarea(r, c, btype, bturn)
'그리거나 지우는 영역 선택
    Set center = Cells(r, c)
    '블록의 기준점
    Select Case btype
    Case 1
        If bturn = 1 Then
            Range(center.Offset(0, -1), center.Offset(0, 2)).Select
        ElseIf bturn = 2 Then
            Range(center.Offset(-1, 0), center.Offset(2, 0)).Select
        End If
    Case 2
        If bturn = 1 Then
            Union(Range(center.Offset(1, -1), center.Offset(1, 1)), center.Offset(0, -1)).Select
        ElseIf bturn = 2 Then
            Union(Range(center.Offset(-1, 1), center.Offset(1, 1)), center.Offset(1, 0)).Select
        ElseIf bturn = 3 Then
            Union(Range(center.Offset(-1, -1), center.Offset(-1, 1)), center.Offset(0, 1)).Select
        ElseIf bturn = 4 Then
            Union(Range(center.Offset(-1, -1), center.Offset(1, -1)), center.Offset(-1, 0)).Select
        End If
    Case 3
        If bturn = 1 Then
            Union(Range(center.Offset(1, -1), center.Offset(1, 1)), center.Offset(0, 1)).Select
        ElseIf bturn = 2 Then
            Union(Range(center.Offset(-1, 1), center.Offset(1, 1)), center.Offset(-1, 0)).Select
        ElseIf bturn = 3 Then
            Union(Range(center.Offset(-1, -1), center.Offset(-1, 1)), center.Offset(0, -1)).Select
        ElseIf bturn = 4 Then
            Union(Range(center.Offset(-1, -1), center.Offset(1, -1)), center.Offset(1, 0)).Select
        End If
    Case 4
        If bturn = 1 Then
            Union(center.Offset(0, -1), center, center.Offset(1, 0), center.Offset(1, 1)).Select
        ElseIf bturn = 2 Then
            Union(center, center.Offset(1, 0), center.Offset(-1, 1), center.Offset(0, 1)).Select
        End If
    Case 5
        If bturn = 1 Then
            Union(center.Offset(0, 1), center, center.Offset(1, 0), center.Offset(1, -1)).Select
        ElseIf bturn = 2 Then
            Union(center, center.Offset(-1, 0), center.Offset(1, 1), center.Offset(0, 1)).Select
        End If
    Case 6
        If bturn = 1 Then
            Union(Range(center.Offset(1, -1), center.Offset(1, 1)), center).Select
        ElseIf bturn = 2 Then
            Union(Range(center.Offset(-1, 1), center.Offset(1, 1)), center).Select
        ElseIf bturn = 3 Then
            Union(Range(center.Offset(-1, -1), center.Offset(-1, 1)), center).Select
        ElseIf bturn = 4 Then
            Union(Range(center.Offset(-1, -1), center.Offset(1, -1)), center).Select
        End If
    Case 7
        Range(center.Offset(0, -1), center.Offset(1, 0)).Select
    End Select
End Sub

Sub drawblock(r, c, btype, bturn)
'블록 그리기
    Dim color(6)
    color(0) = RGB(0, 255, 255)
    color(1) = RGB(101, 101, 255)
    color(2) = RGB(255, 165, 0)
    color(3) = RGB(255, 0, 0)
    color(4) = RGB(0, 255, 0)
    color(5) = RGB(170, 0, 255)
    color(6) = RGB(229, 229, 0)
    
    Call Selectarea(r, c, btype, bturn)
    
    On Error Resume Next
    Selection.Interior.ColorIndex = xlNone
'    Selection.Interior.color = color(Cells(3, 44) - 1)
    With Selection.Interior
        .Pattern = xlPatternRectangularGradient
        .Gradient.RectangleLeft = 0.5
        .Gradient.RectangleRight = 0.5
        .Gradient.RectangleTop = 0.5
        .Gradient.RectangleBottom = 0.5
        .Gradient.ColorStops.Clear
    End With
    With Selection.Interior.Gradient.ColorStops.Add(0)
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    With Selection.Interior.Gradient.ColorStops.Add(1)
        .color = color(btype - 1)
        .TintAndShade = 0
    End With
    Cells(48, 48).Select
End Sub

Sub drawother(i, btype)
'i=0이면 다음 블록, i=1이면 보관 블록 그리기
    If btype > 0 Then
        
        If btype = 1 Then
            r = 7 + i * 8
            c = 25
        ElseIf btype = 7 Then
            r = 6 + i * 8
            c = 26
        Else
            r = 6 + i * 8
            c = 25
        End If
        
        Call drawblock(r, c, btype, 1)
        
    End If
End Sub

Sub Eraseblock()
'블록 지우는 함수
    Call Selectarea(Cells(3, 42), Cells(3, 43), Cells(3, 44), Cells(3, 45))
    
    On Error Resume Next
    Selection.Interior.ColorIndex = xlNone
    Cells(48, 48).Select
End Sub

Function checkleft()
'왼쪽 이동 가능한지 검사
    Set center = Cells(Cells(3, 42), Cells(3, 43))
    Select Case Cells(3, 44)
    Case 1
        If Cells(3, 45) = 1 Then
            center.Offset(0, -2).Select
        ElseIf Cells(3, 45) = 2 Then
            Range(center.Offset(-1, -1), center.Offset(2, -1)).Select
        End If
    Case 2
        If Cells(3, 45) = 1 Then
            Range(center.Offset(0, -2), center.Offset(1, -2)).Select
        ElseIf Cells(3, 45) = 2 Then
            Union(center.Offset(-1, 0), center, center.Offset(1, -1)).Select
        ElseIf Cells(3, 45) = 3 Then
            Union(center.Offset(-1, -2), center).Select
        ElseIf Cells(3, 45) = 4 Then
            Range(center.Offset(-1, -2), center.Offset(1, -2)).Select
        End If
    Case 3
        If Cells(3, 45) = 1 Then
            Union(center, center.Offset(1, -2)).Select
        ElseIf Cells(3, 45) = 2 Then
            Union(center, center.Offset(1, 0), center.Offset(-1, -1)).Select
        ElseIf Cells(3, 45) = 3 Then
            Range(center.Offset(-1, -2), center.Offset(0, -2)).Select
        ElseIf Cells(3, 45) = 4 Then
            Range(center.Offset(-1, -2), center.Offset(1, -2)).Select
        End If
    Case 4
        If Cells(3, 45) = 1 Then
            Union(center.Offset(0, -2), center.Offset(1, -1)).Select
        ElseIf Cells(3, 45) = 2 Then
            Union(center.Offset(0, -1), center.Offset(1, -1), center.Offset(-1, 0)).Select
        End If
    Case 5
        If Cells(3, 45) = 1 Then
            Union(center.Offset(0, -1), center.Offset(1, -2)).Select
        ElseIf Cells(3, 45) = 2 Then
            Union(center.Offset(-1, -1), center.Offset(0, -1), center.Offset(1, 0)).Select
        End If
    Case 6
        If Cells(3, 45) = 1 Then
            Union(center.Offset(0, -1), center.Offset(1, -2)).Select
        ElseIf Cells(3, 45) = 2 Then
            Union(center.Offset(-1, 0), center.Offset(0, -1), center.Offset(1, 0)).Select
        ElseIf Cells(3, 45) = 3 Then
            Union(center.Offset(-1, -2), center.Offset(0, -1)).Select
        ElseIf Cells(3, 45) = 4 Then
            Range(center.Offset(-1, -2), center.Offset(1, -2)).Select
        End If
    Case 7
        Range(center.Offset(0, -2), center.Offset(1, -2)).Select
    End Select
    
    If Selection.Interior.ColorIndex = xlNone Then
        checkleft = True
    Else
        checkleft = False
    End If
    Cells(48, 48).Select
End Function

Function checkright()
'오른쪽 이동 가능한지 검사
    Set center = Cells(Cells(3, 42), Cells(3, 43))
    Select Case Cells(3, 44)
    Case 1
        If Cells(3, 45) = 1 Then
            center.Offset(0, 3).Select
        ElseIf Cells(3, 45) = 2 Then
            Range(center.Offset(-1, 1), center.Offset(2, 1)).Select
        End If
    Case 2
        If Cells(3, 45) = 1 Then
            Union(center, center.Offset(1, 2)).Select
        ElseIf Cells(3, 45) = 2 Then
            Range(center.Offset(-1, 2), center.Offset(1, 2)).Select
        ElseIf Cells(3, 45) = 3 Then
            Range(center.Offset(-1, 2), center.Offset(0, 2)).Select
        ElseIf Cells(3, 45) = 4 Then
            Union(center, center.Offset(1, 0), center.Offset(-1, 1)).Select
        End If
    Case 3
        If Cells(3, 45) = 1 Then
            Range(center.Offset(0, 2), center.Offset(1, 2)).Select
        ElseIf Cells(3, 45) = 2 Then
            Range(center.Offset(-1, 2), center.Offset(1, 2)).Select
        ElseIf Cells(3, 45) = 3 Then
            Union(center, center.Offset(-1, 2)).Select
        ElseIf Cells(3, 45) = 4 Then
            Union(center, center.Offset(-1, 0), center.Offset(1, 1)).Select
        End If
    Case 4
        If Cells(3, 45) = 1 Then
            Union(center.Offset(0, 1), center.Offset(1, 2)).Select
        ElseIf Cells(3, 45) = 2 Then
            Union(center.Offset(-1, 2), center.Offset(0, 2), center.Offset(1, 1)).Select
        End If
    Case 5
        If Cells(3, 45) = 1 Then
            Union(center.Offset(0, 2), center.Offset(1, 1)).Select
        ElseIf Cells(3, 45) = 2 Then
            Union(center.Offset(0, 2), center.Offset(1, 2), center.Offset(-1, 1)).Select
        End If
    Case 6
        If Cells(3, 45) = 1 Then
            Union(center.Offset(0, 1), center.Offset(1, 2)).Select
        ElseIf Cells(3, 45) = 2 Then
            Range(center.Offset(-1, 2), center.Offset(1, 2)).Select
        ElseIf Cells(3, 45) = 3 Then
            Union(center.Offset(-1, 2), center.Offset(0, 1)).Select
        ElseIf Cells(3, 45) = 4 Then
            Union(center.Offset(-1, 0), center.Offset(0, 1), center.Offset(1, 0)).Select
        End If
    Case 7
        Range(center.Offset(0, 1), center.Offset(1, 1)).Select
    End Select
    
    If Selection.Interior.ColorIndex = xlNone Then
        checkright = True
    Else
        checkright = False
    End If
    Cells(48, 48).Select
End Function

Function checkdown(r)
'아래쪽 이동 가능한지 검사
    Set center = Cells(r, Cells(3, 43))
    Select Case Cells(3, 44)
    Case 1
        If Cells(3, 45) = 1 Then
            Range(center.Offset(1, -1), center.Offset(1, 2)).Select
        ElseIf Cells(3, 45) = 2 Then
            center.Offset(3, 0).Select
        End If
    Case 2
        If Cells(3, 45) = 1 Then
            Range(center.Offset(2, -1), center.Offset(2, 1)).Select
        ElseIf Cells(3, 45) = 2 Then
            Range(center.Offset(2, 0), center.Offset(2, 1)).Select
        ElseIf Cells(3, 45) = 3 Then
            Union(center.Offset(0, -1), center, center.Offset(1, 1)).Select
        ElseIf Cells(3, 45) = 4 Then
            Union(center, center.Offset(2, -1)).Select
        End If
    Case 3
        If Cells(3, 45) = 1 Then
            Range(center.Offset(2, -1), center.Offset(2, 1)).Select
        ElseIf Cells(3, 45) = 2 Then
            Union(center, center.Offset(2, 1)).Select
        ElseIf Cells(3, 45) = 3 Then
            Union(center.Offset(0, 1), center, center.Offset(1, -1)).Select
        ElseIf Cells(3, 45) = 4 Then
            Range(center.Offset(2, -1), center.Offset(2, 0)).Select
        End If
    Case 4
        If Cells(3, 45) = 1 Then
            Union(center.Offset(2, 0), center.Offset(2, 1), center.Offset(1, -1)).Select
        ElseIf Cells(3, 45) = 2 Then
            Union(center.Offset(2, 0), center.Offset(1, 1)).Select
        End If
    Case 5
        If Cells(3, 45) = 1 Then
            Union(center.Offset(2, -1), center.Offset(2, 0), center.Offset(1, 1)).Select
        ElseIf Cells(3, 45) = 2 Then
            Union(center.Offset(1, 0), center.Offset(2, 1)).Select
        End If
    Case 6
        If Cells(3, 45) = 1 Then
        Range(center.Offset(2, -1), center.Offset(2, 1)).Select
        ElseIf Cells(3, 45) = 2 Then
            Union(center.Offset(1, 0), center.Offset(2, 1)).Select
        ElseIf Cells(3, 45) = 3 Then
            Union(center.Offset(0, -1), center.Offset(1, 0), center.Offset(0, 1)).Select
        ElseIf Cells(3, 45) = 4 Then
            Union(center.Offset(2, -1), center.Offset(1, 0)).Select
        End If
    Case 7
        Range(center.Offset(2, -1), center.Offset(2, 0)).Select
    End Select
    
    emptycnt = 0
    allcnt = 0
    For Each x In Selection
        'for each를 selection에 적용해 개별셀 접근
        If x.Interior.ColorIndex = xlNone Or x.Interior.Pattern = xlPatternDown Then
            '그림자랑 겹칠수도 있으므로 그림자 패턴에도 개수 증가
            emptycnt = emptycnt + 1
        End If
        allcnt = allcnt + 1
    Next x
    
    If emptycnt = allcnt Then
        checkdown = True
        '아래쪽 영역이 그림자 또는 빈칸이면 이동가능 지역
    Else
        checkdown = False
    End If
    Cells(48, 48).Select
End Function

Function checkrclockturn()
'반시계 방향 회전 가능한지 검사
    If Cells(3, 42) = 3 Then
        Call godown
        Call ActiveSheet.Unprotect("Tkdlqjrj")
    End If
    Set center = Cells(Cells(3, 42), Cells(3, 43))
    beforec = Cells(3, 43)
    Select Case Cells(3, 44)
    Case 1
        If Cells(3, 45) = 1 Then
        Union(center.Offset(1, 0), center.Offset(2, 0), center.Offset(-1, 0)).Select
        ElseIf Cells(3, 45) = 2 Then
            If Cells(3, 43) = 3 Then
                Call goright
                Call ActiveSheet.Unprotect("Tkdlqjrj")
                Set center = Cells(Cells(3, 42), Cells(3, 43))
            ElseIf Cells(3, 43) > 16 Then
                For i = 1 To (Cells(3, 43) - 16)
                    Call goleft
                    Call ActiveSheet.Unprotect("Tkdlqjrj")
                Next i
                Set center = Cells(Cells(3, 42), Cells(3, 43))
            End If
            Union(center.Offset(0, 1), center.Offset(0, 2), center.Offset(0, -1)).Select
        End If
    Case 2
        If Cells(3, 45) = 1 Then
            Range(center.Offset(-1, 1), center.Offset(0, 1)).Select
        ElseIf Cells(3, 45) = 2 Then
            If Cells(3, 43) = 3 Then
                Call goright
                Call ActiveSheet.Unprotect("Tkdlqjrj")
                Set center = Cells(Cells(3, 42), Cells(3, 43))
            End If
            Range(center.Offset(-1, -1), center.Offset(-1, 0)).Select
        ElseIf Cells(3, 45) = 3 Then
            Range(center.Offset(0, -1), center.Offset(1, -1)).Select
        ElseIf Cells(3, 45) = 4 Then
            If Cells(3, 43) = 18 Then
                Call goleft
                Call ActiveSheet.Unprotect("Tkdlqjrj")
                Set center = Cells(Cells(3, 42), Cells(3, 43))
            End If
            Range(center.Offset(1, 0), center.Offset(1, 1)).Select
        End If
    Case 3
        If Cells(3, 45) = 1 Then
            Range(center.Offset(-1, 0), center.Offset(-1, 1)).Select
        ElseIf Cells(3, 45) = 2 Then
            If Cells(3, 43) = 3 Then
                Call goright
                Call ActiveSheet.Unprotect("Tkdlqjrj")
                Set center = Cells(Cells(3, 42), Cells(3, 43))
            End If
            Range(center.Offset(-1, -1), center.Offset(0, -1)).Select
        ElseIf Cells(3, 45) = 3 Then
            Range(center.Offset(1, -1), center.Offset(1, 0)).Select
        ElseIf Cells(3, 45) = 4 Then
            If Cells(3, 43) = 18 Then
                Call goleft
                Call ActiveSheet.Unprotect("Tkdlqjrj")
                Set center = Cells(Cells(3, 42), Cells(3, 43))
            End If
            Range(center.Offset(0, 1), center.Offset(1, 1)).Select
        End If
    Case 4
        If Cells(3, 45) = 1 Then
            Range(center.Offset(-1, 1), center.Offset(0, 1)).Select
        ElseIf Cells(3, 45) = 2 Then
            If Cells(3, 43) = 3 Then
                Call goright
                Call ActiveSheet.Unprotect("Tkdlqjrj")
                Set center = Cells(Cells(3, 42), Cells(3, 43))
            End If
            Union(center.Offset(0, -1), center.Offset(1, 1)).Select
        End If
    Case 5
        If Cells(3, 45) = 1 Then
            Union(center.Offset(-1, 0), center.Offset(1, 1)).Select
        ElseIf Cells(3, 45) = 2 Then
            If Cells(3, 43) = 3 Then
                Call goright
                Call ActiveSheet.Unprotect("Tkdlqjrj")
                Set center = Cells(Cells(3, 42), Cells(3, 43))
            End If
            Range(center.Offset(1, -1), center.Offset(1, 0)).Select
        End If
    Case 6
        If Cells(3, 45) = 1 Then
            Range(center.Offset(-1, 1), center.Offset(0, 1)).Select
        ElseIf Cells(3, 45) = 2 Then
            If Cells(3, 43) = 3 Then
                Call goright
                Call ActiveSheet.Unprotect("Tkdlqjrj")
                Set center = Cells(Cells(3, 42), Cells(3, 43))
            End If
            Range(center.Offset(-1, -1), center.Offset(-1, 0)).Select
        ElseIf Cells(3, 45) = 3 Then
            Range(center.Offset(0, -1), center.Offset(1, -1)).Select
        ElseIf Cells(3, 45) = 4 Then
            If Cells(3, 43) = 18 Then
                Call goleft
                Call ActiveSheet.Unprotect("Tkdlqjrj")
                Set center = Cells(Cells(3, 42), Cells(3, 43))
            End If
            Range(center.Offset(1, 0), center.Offset(1, 1)).Select
        End If
    End Select
    
    emptycnt = 0
    allcnt = 0
    For Each x In Selection
        'for each를 selection에 적용해 개별셀 접근
        If x.Interior.ColorIndex = xlNone Or x.Interior.Pattern = xlPatternDown Then
            '그림자랑 겹칠수도 있으므로 그림자 패턴에도 개수 증가
            emptycnt = emptycnt + 1
        End If
        allcnt = allcnt + 1
    Next x
    
    If emptycnt = allcnt Then
        checkrclockturn = True
    Else
        If beforec <> Cells(3, 43) Then
            Call Eraseshadow
            Call Eraseblock
            Cells(3, 43) = beforec
            Call drawshadow
            Call drawblock(Cells(3, 42), Cells(3, 43), Cells(3, 44), Cells(3, 45))
        End If
        checkrclockturn = False
    End If
    
    Cells(48, 48).Select
End Function

Function checkclockturn()
'시계 방향 회전 가능한지 검사
    If Cells(3, 42) = 3 Then
        Call godown
        Call ActiveSheet.Unprotect("Tkdlqjrj")
    End If
    
    Set center = Cells(Cells(3, 42), Cells(3, 43))
    beforec = Cells(3, 43)
    Select Case Cells(3, 44)
    Case 1, 4, 5
        checkclockturn = checkrclockturn
        Exit Function
    Case 2
        If Cells(3, 45) = 1 Then
            Range(center.Offset(-1, -1), center.Offset(-1, 0)).Select
        ElseIf Cells(3, 45) = 2 Then
            If Cells(3, 43) = 3 Then
                Call goright
                Call ActiveSheet.Unprotect("Tkdlqjrj")
                Set center = Cells(Cells(3, 42), Cells(3, 43))
            End If
            Range(center.Offset(0, -1), center.Offset(1, -1)).Select
        ElseIf Cells(3, 45) = 3 Then
            Range(center.Offset(1, 0), center.Offset(1, 1)).Select
        ElseIf Cells(3, 45) = 4 Then
            If Cells(3, 43) = 18 Then
                Call goleft
                Call ActiveSheet.Unprotect("Tkdlqjrj")
                Set center = Cells(Cells(3, 42), Cells(3, 43))
            End If
            Range(center.Offset(-1, 1), center.Offset(0, 1)).Select
        End If
    Case 3
        If Cells(3, 45) = 1 Then
            Range(center.Offset(-1, -1), center.Offset(0, -1)).Select
        ElseIf Cells(3, 45) = 2 Then
            If Cells(3, 43) = 3 Then
                Call goright
                Call ActiveSheet.Unprotect("Tkdlqjrj")
                Set center = Cells(Cells(3, 42), Cells(3, 43))
            End If
            Range(center.Offset(1, -1), center.Offset(1, 0)).Select
        ElseIf Cells(3, 45) = 3 Then
            Range(center.Offset(0, 1), center.Offset(1, 1)).Select
        ElseIf Cells(3, 45) = 4 Then
            If Cells(3, 43) = 18 Then
                Call goleft
                Call ActiveSheet.Unprotect("Tkdlqjrj")
                Set center = Cells(Cells(3, 42), Cells(3, 43))
            End If
            Range(center.Offset(-1, 1), center.Offset(-1, 0)).Select
        End If
    Case 6
        If Cells(3, 45) = 1 Then
            Range(center.Offset(-1, -1), center.Offset(0, -1)).Select
        ElseIf Cells(3, 45) = 2 Then
            If Cells(3, 43) = 3 Then
                Call goright
                Call ActiveSheet.Unprotect("Tkdlqjrj")
                Set center = Cells(Cells(3, 42), Cells(3, 43))
            End If
            Range(center.Offset(1, -1), center.Offset(1, 0)).Select
        ElseIf Cells(3, 45) = 3 Then
            Range(center.Offset(0, 1), center.Offset(1, 1)).Select
        ElseIf Cells(3, 45) = 4 Then
            If Cells(3, 43) = 18 Then
                Call goleft
                Call ActiveSheet.Unprotect("Tkdlqjrj")
                Set center = Cells(Cells(3, 42), Cells(3, 43))
            End If
            Range(center.Offset(-1, 0), center.Offset(-1, 1)).Select
        End If
    End Select
    
    emptycnt = 0
    allcnt = 0
    For Each x In Selection
        'for each를 selection에 적용해 개별셀 접근
        If x.Interior.ColorIndex = xlNone Or x.Interior.Pattern = xlPatternDown Then
            '그림자랑 겹칠수도 있으므로 그림자 패턴에도 개수 증가
            emptycnt = emptycnt + 1
        End If
        allcnt = allcnt + 1
    Next x
    
    If emptycnt = allcnt Then
        checkclockturn = True
    Else
        If beforec <> Cells(3, 43) Then
            Call Eraseshadow
            Call Eraseblock
            Cells(3, 43) = beforec
            Call drawshadow
            Call drawblock(Cells(3, 42), Cells(3, 43), Cells(3, 44), Cells(3, 45))
        End If
        checkclockturn = False
    End If
    
    Cells(48, 48).Select
End Function

Sub hold()
    Call ActiveSheet.Unprotect("Tkdlqjrj")
    Application.ScreenUpdating = False
    Cells(3, 33) = Cells(3, 33) + 6 / 100
    Call makingline
    
    If Cells(5, 45) = 0 Then '블록 보관하기
        Call Eraseshadow
        Call Eraseblock
        
        For i = 0 To 1
            Range(Cells(4 + 8 * i, 23), Cells(9 + 8 * i, 28)).Interior.ColorIndex = xlNone
        Next i
        
        '현재 블록 보관하기
        Cells(5, 45) = -Cells(3, 44)
        Call drawother(1, -Cells(5, 45))
        
        '다음 블록 등장시키기
        Cells(3, 42) = 3
        Cells(3, 43) = 11
        Cells(3, 45) = 1
        Call proceedprd
        Call drawother(0, Cells(5, 44))
        Call drawshadow
        Call drawblock(Cells(3, 42), Cells(3, 43), Cells(3, 44), Cells(3, 45))
        
    ElseIf (Cells(5, 45) > 0 And Cells(5, 48) < 8) And (Cells(5, 45) <> Cells(3, 44)) Then '블록 꺼내기
        Call Eraseshadow
        Call Eraseblock
        Range(Cells(4 + 8, 23), Cells(9 + 8, 28)).Interior.ColorIndex = xlNone
        
        '보관 블록과 현재 블록 교환하기
        tmp = Cells(3, 44)
        Cells(3, 44) = Cells(5, 45)
        Cells(5, 45) = -tmp
        Call drawother(1, -Cells(5, 45))
        
        '보관 블록 등장시키기
        Cells(3, 42) = 3
        Cells(3, 43) = 11
        Cells(3, 45) = 1
        Call drawshadow
        Call drawblock(Cells(3, 42), Cells(3, 43), Cells(3, 44), Cells(3, 45))
    End If
    Call ActiveSheet.Protect("Tkdlqjrj", False, True)
End Sub

Sub goleft()
    Call ActiveSheet.Unprotect("Tkdlqjrj")
    Application.ScreenUpdating = False
    Cells(3, 33) = Cells(3, 33) + 6 / 100
    Call makingline
    
    If checkleft Then
        Call Eraseshadow
        Call Eraseblock
        
        Cells(3, 43) = Cells(3, 43) - 1
        
        Call drawshadow
        Call drawblock(Cells(3, 42), Cells(3, 43), Cells(3, 44), Cells(3, 45))
    End If
    
    Call ActiveSheet.Protect("Tkdlqjrj", False, True)
End Sub

Sub goright()
    Call ActiveSheet.Unprotect("Tkdlqjrj")
    Application.ScreenUpdating = False
    Cells(3, 33) = Cells(3, 33) + 6 / 100
    Call makingline
    
    If checkright Then
        Call Eraseshadow
        Call Eraseblock
        
        Cells(3, 43) = Cells(3, 43) + 1
        
        Call drawshadow
        Call drawblock(Cells(3, 42), Cells(3, 43), Cells(3, 44), Cells(3, 45))
    End If
    Call ActiveSheet.Protect("Tkdlqjrj", False, True)
End Sub

Sub godown()
    Call ActiveSheet.Unprotect("Tkdlqjrj")
    Application.ScreenUpdating = False
    Cells(3, 33) = Cells(3, 33) + 6 / 100
    Call makingline
    
    If checkdown(Cells(3, 42)) Then
        Call Eraseblock
        
        Cells(3, 42) = Cells(3, 42) + 1
        
        Call drawblock(Cells(3, 42), Cells(3, 43), Cells(3, 44), Cells(3, 45))
    End If
    Call ActiveSheet.Protect("Tkdlqjrj", False, True)
End Sub

Sub gobottom()
    '하드드롭 함수
    Call ActiveSheet.Unprotect("Tkdlqjrj")
    Application.ScreenUpdating = False
    Cells(3, 33) = Cells(3, 33) + 6 / 100
    Call makingline
    
    If Cells(3, 42) > 2 Then
        Call Eraseshadow
        Call Eraseblock
    End If
    
    Cells(3, 42) = Cells(3, 46)
    '그림자 위치로 옮기는 것과 같음

    Call drawblock(Cells(3, 42), Cells(3, 43), Cells(3, 44), Cells(3, 45))
    Call nextstatus(0)
    
    Call ActiveSheet.Protect("Tkdlqjrj", False, True)
End Sub

Sub clockturn()
    '시계방향 회전
    Call ActiveSheet.Unprotect("Tkdlqjrj")
    Application.ScreenUpdating = False
    Cells(3, 33) = Cells(3, 33) + 6 / 100
    Call makingline
    If Cells(3, 44) <> 7 Then
        
        If checkclockturn Then
            Call Eraseshadow
            Call Eraseblock
            
            Cells(3, 45) = Cells(3, 45) - 1
            
            If (Cells(3, 44) = 1 Or Cells(3, 44) = 4 Or Cells(3, 44) = 5) And Cells(3, 45) = 0 Then
                Cells(3, 45) = 2
            ElseIf (Cells(3, 44) = 2 Or Cells(3, 44) = 3 Or Cells(3, 44) = 6) And Cells(3, 45) = 0 Then
                Cells(3, 45) = 4
            End If
            
            Call drawshadow
            Call drawblock(Cells(3, 42), Cells(3, 43), Cells(3, 44), Cells(3, 45))
        End If
    End If
    Call ActiveSheet.Protect("Tkdlqjrj", False, True)
End Sub

Sub rclockturn()
    '반시계방향 회전
    Call ActiveSheet.Unprotect("Tkdlqjrj")
    Application.ScreenUpdating = False
    Cells(3, 33) = Cells(3, 33) + 6 / 100
    Call makingline
    If Cells(3, 44) <> 7 Then
        
        If checkrclockturn Then
            Call Eraseshadow
            Call Eraseblock
            
            Cells(3, 45) = Cells(3, 45) + 1
            
            If ((Cells(3, 44) = 1 Or Cells(3, 44) = 4 Or Cells(3, 44) = 5) And Cells(3, 45) = 3) Or _
             ((Cells(3, 44) = 2 Or Cells(3, 44) = 3 Or Cells(3, 44) = 6) And Cells(3, 45) = 5) Then
                Cells(3, 45) = 1
            End If
            
            Call drawshadow
            Call drawblock(Cells(3, 42), Cells(3, 43), Cells(3, 44), Cells(3, 45))
        End If
    End If
    Call ActiveSheet.Protect("Tkdlqjrj", False, True)
End Sub


Sub makeperiod()
    Dim blockarr(6)
    For i = 0 To 6
        blockarr(i) = i + 1
    Next i
    
    For i = 0 To 6
        r = Int(Rnd * 7)
        tmp = blockarr(i)
        blockarr(i) = blockarr(r)
        blockarr(r) = tmp
    Next i
    
    If blockarr(0) = Cells(3, 44) Then
        tmp = blockarr(0)
        blockarr(0) = blockarr(1)
        blockarr(1) = tmp
    End If
    
    For i = 0 To 6
        Cells(5 + i, 44) = blockarr(i)
    Next i
End Sub

Sub proceedprd()
    Cells(3, 44) = Cells(5, 44)
    If Cells(6, 44) = "" Then
        Call makeperiod
    Else
        For i = 0 To 6
            If Cells(5 + i, 44) = "" Then
                Exit For
            End If
            Cells(5 + i, 44) = Cells(6 + i, 44)
        Next i
    End If
    
End Sub

Sub nextstatus(ismain)
    Call calhigh
    Call Eraseline
    
    Cells(3, 42) = 3
    Cells(3, 43) = 11
    Cells(3, 45) = 1
    Call proceedprd
    
    If Cells(5, 45) < 0 Then
        Cells(5, 45) = -Cells(5, 45)
    End If
    
    Range(Cells(4, 23), Cells(9, 28)).Interior.ColorIndex = xlNone
    
    If Cells(3, 47) > 3 And checkdown(Cells(3, 42)) Then
        Call drawother(0, Cells(5, 44))
        Call drawshadow
        Call drawblock(Cells(3, 42), Cells(3, 43), Cells(3, 44), Cells(3, 45))
        Application.ScreenUpdating = True
        
        
        If ismain = 1 Then
            veltmp = 31 - Cells(8, 33)
            If veltmp < 5 Then
                veltmp = 5
            End If
            '가변속도 알고리즘
            Cells(5, 43) = "=NOW()+""00:00:0" & CStr(veltmp / 100) & """"
            Cells(3, 33) = Cells(3, 33) + veltmp / 100
            Application.OnTime Cells(5, 43), "gameloop"
        End If
        Exit Sub
    Else
        If Cells(3, 47) > 3 Then
            Cells(3, 42) = 2
            If checkdown(Cells(3, 42)) Then
                Cells(3, 42) = 3
            End If
            Call drawblock(Cells(3, 42), Cells(3, 43), Cells(3, 44), Cells(3, 45))
        End If
        Application.ScreenUpdating = True
        
        MsgBox "GAME OVER!!!!" & vbNewLine & Cells(10, 33) & "점을 달성하였습니다." & _
        vbNewLine & "플레이 시간 : " & Int(Cells(3, 33) + 0.5) & "초", 64, "테트리스"
        
        'Range(Cells(2, 3), Cells(2, 18)).Interior.color = RGB(128, 128, 128)
        With Range(Cells(2, 3), Cells(2, 18)).Interior
            .Pattern = xlPatternLinearGradient
            .Gradient.Degree = 90
            .Gradient.ColorStops.Clear
        End With
        With Range(Cells(2, 3), Cells(2, 18)).Interior.Gradient.ColorStops.Add(0)
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
        End With
        With Range(Cells(2, 3), Cells(2, 18)).Interior.Gradient.ColorStops.Add(1)
            .color = RGB(128, 128, 128)
            .TintAndShade = 0
        End With
        
        Call resetgame
        Exit Sub
    End If
End Sub
Sub Eraseline()
    Call ActiveSheet.Unprotect("Tkdlqjrj")
    ercnt = 0
    '지운 줄의 수
    score = 0
    '얻은 점수
    
    For i = WorksheetFunction.Min(30, Cells(3, 42) + 2) To Cells(3, 47) Step -1
        cnt = 0
        '줄에서 채워진 블록의 수
        For j = 3 To 18
            If Cells(i, j).Interior.ColorIndex <> xlNone Then
                cnt = cnt + 1
            End If
        Next
        
        If cnt = 16 Then
            Range(Cells(Cells(3, 47) - 1, 3), Cells(i - 1, 18)).Copy Range(Cells(Cells(3, 47), 3), Cells(i, 18))
            Application.CutCopyMode = False
            
            ercnt = ercnt + 1
            Cells(3, 47) = Cells(3, 47) + 1
            i = i + 1
        End If
    Next i
    
    If Cells(12, 33) > 0 Then
        score = Cells(8, 33) * 10 * (Cells(12, 30) + 50)
    End If
    '콤보점수 반영
    
    If ercnt = 0 Then
        Cells(12, 33) = 0
    Else
        Cells(12, 33) = Cells(12, 33) + 1
        Cells(8, 35) = Cells(8, 35) + ercnt
        If ercnt = 1 Then
            score = score + Cells(8, 33) * 1000
        ElseIf ercnt = 2 Then
            score = score + Cells(8, 33) * 3000
        ElseIf ercnt = 3 Then
            score = score + Cells(8, 33) * 5000
        ElseIf ercnt = 4 Then
            score = score + Cells(8, 33) * 8000
        End If
        
        '퍼펙트 클리어 판정
        If Range(Cells(3, 3), Cells(30, 18)).Interior.ColorIndex = xlNone Then
            score = score + Cells(8, 33) * 30000
        End If
        
        If Cells(8, 35) > 9 Then
            Cells(8, 33) = Cells(8, 33) + 1
            Cells(8, 35) = Cells(8, 35) - 10
        End If
        '레벨업 시스템
    End If
    
    Cells(10, 33) = Cells(10, 33) + score
    If Cells(14, 33) < Cells(10, 33) Then
        Cells(14, 33) = Cells(10, 33)
        Cells(14, 34) = Cells(3, 33)
        ActiveWorkbook.Save
    End If
End Sub

Sub calhigh()
    '최상단 행 계산
    hightmp = 0
    Select Case Cells(3, 44)
    Case 1
        Select Case Cells(3, 45)
        Case 1
            hightmp = Cells(3, 42)
        Case 2
            hightmp = Cells(3, 42) - 1
        End Select
    Case 2
        Select Case Cells(3, 45)
        Case 1
            hightmp = Cells(3, 42)
        Case 2, 3, 4
            hightmp = Cells(3, 42) - 1
        End Select
    Case 3
        Select Case Cells(3, 45)
        Case 1
            hightmp = Cells(3, 42)
        Case 2, 3, 4
            hightmp = Cells(3, 42) - 1
        End Select
    Case 4
        Select Case Cells(3, 45)
        Case 1
            hightmp = Cells(3, 42)
        Case 2
            hightmp = Cells(3, 42) - 1
        End Select
    Case 5
        Select Case Cells(3, 45)
        Case 1
            hightmp = Cells(3, 42)
        Case 2
            hightmp = Cells(3, 42) - 1
        End Select
    Case 6
    Select Case Cells(3, 45)
        Case 1
            hightmp = Cells(3, 42)
        Case 2, 3, 4
            hightmp = Cells(3, 42) - 1
        End Select
    Case 7
        hightmp = Cells(3, 42)
    End Select
    
    If Cells(3, 47) > hightmp Then
        Cells(3, 47) = hightmp
    End If
End Sub

Sub makingline()
    Timeterm = 81 - Cells(8, 33)
    If Timeterm < 60 Then
        Timeterm = 60
    End If
    
    '줄 올라온 시간이 됐을 때
    If Cells(3, 42) > 2 And Int(Cells(3, 33) - Cells(5, 46)) >= Timeterm Then
        Cells(5, 46) = Int(Cells(3, 33))
        
        Call Eraseshadow
        Call Eraseblock
        Call panaltyline
        
        
        If Cells(3, 47) > 3 Then
            If Cells(3, 45) = 1 Then
                Cells(3, 46) = 3
            Else
                Cells(3, 46) = 4
            End If
            Call drawshadow
            
            If Cells(3, 42) < Cells(3, 46) Then
                '원래 위치
                Call drawblock(Cells(3, 42), Cells(3, 43), Cells(3, 44), Cells(3, 45))
            Else
                If checkdown(Cells(3, 42)) Then
                    Cells(3, 42) = Cells(3, 46)
                    '올라온 줄에 의해 바닥에 붙은 위치
                    Call drawblock(Cells(3, 42), Cells(3, 43), Cells(3, 44), Cells(3, 45))
                Else
                    '게임오버 경우임
                    Call drawblock(Cells(3, 42), Cells(3, 43), Cells(3, 44), Cells(3, 45))
                End If
            End If
        End If
    End If
End Sub

Sub gameloop()
    '메인 재귀(자동 하강) 함수
    If Cells(5, 42) = 0 Then
        Exit Sub
    Else
        Call ActiveSheet.Unprotect("Tkdlqjrj")
        Application.ScreenUpdating = False
        
        Call makingline
        
        If checkdown(Cells(3, 42)) Then
            Call Eraseblock
            
            Cells(3, 42) = Cells(3, 42) + 1
            
            Call drawblock(Cells(3, 42), Cells(3, 43), Cells(3, 44), Cells(3, 45))
            
            'Cells(5, 43) = "=NOW()+""00:00:00.25"""
            'Cells(3, 33) = Cells(3, 33) + 25 / 85
            veltmp = 31 - Cells(8, 33)
            If veltmp < 5 Then
                veltmp = 5
            End If
            '가변속도 알고리즘
            Cells(5, 43) = "=NOW()+""00:00:0" & CStr(veltmp / 100) & """"
            Cells(3, 33) = Cells(3, 33) + veltmp / 90
            Application.OnTime Cells(5, 43), "gameloop"
        Else
            Call nextstatus(1)
        End If
        Call ActiveSheet.Protect("Tkdlqjrj", False, True)
    End If
End Sub

Sub keyset()
    For i = 0 To 24
        Application.OnKey CStr(Chr(65 + i)), ""
    Next i
    
    For i = 0 To 9
        Application.OnKey CStr(i), ""
    Next i
    
    Application.OnKey "{LEFT}", "goleft"
    Application.OnKey "{RIGHT}", "goright"
    Application.OnKey "{DOWN}", "godown"
    Application.OnKey "{UP}", "clockturn"
    Application.OnKey "Z", "rclockturn"
    Application.OnKey "z", "rclockturn"
    Application.OnKey " ", "gobottom"
    Application.OnKey "X", "hold"
    Application.OnKey "x", "hold"
    Application.OnKey "q", "pausegame"
    Application.OnKey "Q", "pausegame"
End Sub

Sub unkeyset()
    For i = 0 To 24
        Application.OnKey CStr(Chr(65 + i))
    Next i
    
    For i = 0 To 9
        Application.OnKey CStr(i)
    Next i
    
    Application.OnKey "{LEFT}"
    Application.OnKey "{RIGHT}"
    Application.OnKey "{DOWN}"
    Application.OnKey "{UP}"
    Application.OnKey " "
    Application.OnKey "q"
    Application.OnKey "Q"
End Sub

Sub gamestart()
    '게임 시작
    If Cells(5, 42) = 0 Then
        Call ActiveSheet.Unprotect("Tkdlqjrj")
        Cells(5, 42) = 1
        
        Call keyset
        
        Application.ScreenUpdating = False
        If Range(Cells(4, 23), Cells(9, 28)).Interior.ColorIndex = xlNone Then
            '게임이 처음 시작하는 상태인 경우
            Call drawother(0, Cells(5, 44))
            Call drawshadow
            Call drawblock(Cells(3, 42), Cells(3, 43), Cells(3, 44), Cells(3, 45))
        End If
        Application.ScreenUpdating = True
        
        Cells(5, 43) = "=NOW()+""00:00:00.30"""
        Cells(3, 33) = Cells(3, 33) + 30 / 90
        Application.OnTime Cells(5, 43), "gameloop"
        Call ActiveSheet.Protect("Tkdlqjrj", False, True)
    End If
End Sub

Sub pausegame()
    '게임 일시정지
    If Cells(5, 42) = 1 Then
        Call ActiveSheet.Unprotect("Tkdlqjrj")
        Cells(5, 42) = 0
        
        Call unkeyset
        Call ActiveSheet.Protect("Tkdlqjrj", False, True)
    End If
End Sub

Sub resetgame()
    Call ActiveSheet.Unprotect("Tkdlqjrj")
    
    Call makeperiod
    
    Cells(3, 42) = 3
    Cells(3, 43) = 11
    Cells(3, 45) = 1
    Cells(3, 47) = 31
    
    Call proceedprd
    
    Range(Cells(3, 3), Cells(30, 18)).Interior.ColorIndex = xlNone
    For i = 0 To 1
        Range(Cells(4 + 8 * i, 23), Cells(9 + 8 * i, 28)).Interior.ColorIndex = xlNone
    Next i
    
    Cells(8, 33) = 1
    Cells(8, 35) = 0
    Cells(3, 33) = 0
    Cells(10, 33) = 0
    Cells(12, 33) = 0
    Cells(3, 46) = 29 - Int(Cells(3, 44) = 1)
    Cells(5, 45) = 0
    Cells(5, 46) = 0
    
    
    Call pausegame
    Call ActiveSheet.Protect("Tkdlqjrj", False, True)
End Sub

Sub drawshadow()
    Cells(3, 46) = Cells(3, 42)
    '하드드롭 알고리즘 이용
    Do While True
        If checkdown(Cells(3, 46)) Then
            Cells(3, 46) = Cells(3, 46) + 1
        Else
            Exit Do
        End If
    Loop
    
    Call Selectarea(Cells(3, 46), Cells(3, 43), Cells(3, 44), Cells(3, 45))
    
    On Error Resume Next
    With Selection.Interior
        .ColorIndex = xlNone
        .Pattern = xlPatternDown
        .PatternColor = RGB(0, 0, 0)
    End With
    
    Cells(48, 48).Select
End Sub

Sub Eraseshadow()
    
    Call Selectarea(Cells(3, 46), Cells(3, 43), Cells(3, 44), Cells(3, 45))
    
    On Error Resume Next
    Selection.Interior.ColorIndex = xlNone
    
    Cells(48, 48).Select
End Sub

Sub panaltyline()
    Dim flag(15)
    For i = 0 To 15
        flag(i) = i + 3
    Next i
    
    linenum = Int(Cells(8, 33) / 2 + 1)
    If linenum > 6 Then
        linenum = 6
    End If
    
    If Cells(3, 47) - linenum <= 2 Then
        linenum = Cells(3, 47) - 3
    End If
    
    If Cells(3, 47) < 31 Then
        Range(Cells(Cells(3, 47), 3), Cells(30, 18)).Copy Range(Cells(Cells(3, 47) - linenum, 3), Cells(30 - linenum, 18))
        Application.CutCopyMode = False
    End If
    
    Cells(3, 47) = Cells(3, 47) - linenum
    For a = 0 To linenum - 1
        For i = 0 To 15
            r = Int(Rnd * 2)
            tmp = flag(r)
            flag(r) = flag(i)
            flag(i) = tmp
        Next i
        
        For i = 3 To 18
            If emptycell(flag, i) Then
                Cells(30 - a, i).Interior.ColorIndex = xlNone
            Else
'                Cells(30 - a, i).Interior.color = RGB(0, 0, 0)
               With Cells(30 - a, i).Interior
                   .Pattern = xlPatternRectangularGradient
                    .Gradient.RectangleLeft = 0.5
                    .Gradient.RectangleRight = 0.5
                    .Gradient.RectangleTop = 0.5
                    .Gradient.RectangleBottom = 0.5
                    .Gradient.ColorStops.Clear
                End With
                With Cells(30 - a, i).Interior.Gradient.ColorStops.Add(0)
                    .ThemeColor = xlThemeColorDark1
                    .TintAndShade = 0
                End With
                With Cells(30 - a, i).Interior.Gradient.ColorStops.Add(1)
                    .color = RGB(0, 0, 0)
                    .TintAndShade = 0
                End With
            End If
        Next i
    Next a
End Sub

Function emptycell(flag, row)
    emtcnt = Int(Rnd * 4)
    
    For i = 0 To emtcnt
        If row = flag(i) Then
            emptycell = True
            Exit Function
        End If
    Next i
    emptycell = False
End Function

Sub help()
    MsgBox "조작방법" & vbNewLine & "X : 블록 보관하기" & vbNewLine & "Z : 반시계 방향 회전" & vbNewLine & "Q : 일시 정지" & _
    vbNewLine & "위쪽 방향키 : 시계 방향 회전" & vbNewLine & "아래쪽 방향키 : 낙하 속도 증가" & _
    vbNewLine & "왼쪽, 오른쪽 방향키 : 좌우 이동" & vbNewLine & "스페이스바 : 하드 드롭", _
    64, "테트리스"
End Sub
