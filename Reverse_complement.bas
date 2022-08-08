Function complement(seq)
    
    Dim seq_out As String
    
    seq = LCase(seq)
    
    For i = 1 To Len(seq):
        If Mid(seq, i, 1) = "a" Then
            seq_out = seq_out + "t"
        
        ElseIf Mid(seq, i, 1) = "t" Then
            seq_out = seq_out + "a"
        
        ElseIf Mid(seq, i, 1) = "c" Then
            seq_out = seq_out + "g"
        
        ElseIf Mid(seq, i, 1) = "g" Then
            seq_out = seq_out + "c"
        
        Else
            MsgBox ("Error! Incorrect letter is contained")
            MsgBox (i & "th letter is not in atcg. that letter is '" & Mid(seq, i, 1) & "'.")
            
        End If
        
    Next
    complement = seq_out

End Function
----------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function reverse_complement(seq)
    
    seq = VBA.StrReverse(seq_out)
    reverse_complement = seq_out

End Function

