Sub TextPartitionAndInsertion()
    Dim text As String
    ' Input text here
    text = "10余年前无明显诱因出现口干、多饮，每日具体饮水量不详,1个月内体重下降约10公斤，无明显多尿，自觉泡沫尿。多次测量血糖大于7mmol/L，外院诊断为:“2型糖尿病”，给予口服“二甲双胍+注射门冬胰岛素（具体剂量不详）”，经治疗后，上述症状好转。10年来，规律服药，血糖波动较大，随机血糖最高达20mmol/L。近1个月以来，自觉上述症状加重，伴乏力、手指及脚趾远端蚁行感，体重在1个月内下降2.5kg，随机血糖最高达17mmol/L。无心悸、气短、胸闷、怕热，无咳嗽咳痰、头晕头痛，无视物模糊、肌痛、手抖、跛行，无明显皮肤瘙痒，无满月脸、皮肤紫纹，无恶心呕吐，无尿频、尿急、尿痛。目前应用“达格列净10mg Qd+糖适平 30mg Tid+二甲双胍缓释片 0.5g Bid”治疗，测随机血糖波动10~12mmol/L，监测空腹血糖波动8~12mmol/L，餐后血糖17mmol/L，近期每日饮水量约1500ml，门诊拟“2型糖尿病”收入住院。发病以来，白天精力欠佳，体重变化如前所述，食欲食量正常，二便正常。"
 
    Dim maxWidth As Double
    maxWidth = 30 ' Set the maximum width for each partition.
    
    Dim partitions() As String
    Dim currentWidth As Double
    Dim currentChunk As String
    Dim i As Integer
    Dim charWidth As Double
    Dim currentChar As String
    Dim nextChar As String
    
    currentWidth = 0
    currentChunk = ""
    Dim partitionCount As Integer
    partitionCount = 1
    
    ReDim partitions(1 To 1)
    
    ' Loop through each character in the text.
    For i = 1 To Len(text)
        currentChar = Mid(text, i, 1)
        
        ' Get the next character (if it exists).
        If i < Len(text) Then
            nextChar = Mid(text, i + 1, 1)
        Else
            nextChar = " " ' No next character for the last character.
        End If
        
        ' Determine the width of the current character based on the next character.
        If AscW(currentChar) > 127 Then
            ' Current char is non-ASCII
            If AscW(nextChar) > 127 Then
                charWidth = 1 ' Next char is non-ASCII
            Else
                charWidth = 1.33 ' Next char is ASCII
            End If
        Else
            ' Current char is ASCII
            If AscW(nextChar) > 127 Then
                charWidth = 0.66 ' Next char is non-ASCII
            Else
                charWidth = 0.44 ' Next char is ASCII
            End If
        End If
        
        ' Check if adding the character exceeds the maximum width.
        If currentWidth + charWidth > maxWidth Then
            ' Save the current chunk to the partitions array.
            partitions(partitionCount) = currentChunk
            partitionCount = partitionCount + 1
            ReDim Preserve partitions(1 To partitionCount)
            
            ' Start a new chunk.
            currentChunk = currentChar
            currentWidth = charWidth
        Else
            ' Add the character to the current chunk.
            currentChunk = currentChunk & currentChar
            currentWidth = currentWidth + charWidth
        End If
    Next i
    
    ' Add the last chunk.
    If currentChunk <> "" Then
        partitions(partitionCount) = currentChunk
    End If
    

    
    
    Dim doc As Document
    Set doc = ActiveDocument
    Dim tbl As Table
    Set tbl = doc.Tables(1)
    
    
    Dim cellText As String
    Dim t As Integer
    Dim rowIndex As Integer
    rowIndex = 1
    
    ' Find the "Start_here" cell noted by ###
    For rowIndex = 1 To tbl.Rows.Count
        cellText = tbl.Cell(rowIndex, 2).Range.text
        If InStr(cellText, "###") > 0 Then
            Exit For
        End If
    Next rowIndex
    
    ' Insert the subtexts starting from the "Start_here" cell
    For t = LBound(partitions) To UBound(partitions)
        If rowIndex > tbl.Rows.Count Then Exit For
        tbl.Cell(rowIndex, 2).Range.text = partitions(t)
        rowIndex = rowIndex + 1
    Next t
End Sub


