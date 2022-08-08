Attribute VB_Name = "Module1"
Function reverse_complement(seq)
    
    Dim seq_out As String
    
    seq = LCase(seq)
    
    For i = 1 To Len(seq):
        If Mid(seq, i, 1) = "a" Then
            seq_out = seq_out + "t"
        
        ElseIf Mid(seq, i, 1) = "t" Then
            seq_out = seq_out + "a"
        
        ElseIf Mid(seq, i, 1) = "c" Then
            seq_out = seq_out + "g"
        
        Else
            seq_out = seq_out + "c"
        
        End If
        
    Next
    seq_out = VBA.StrReverse(seq_out)
    reverse_complement = seq_out

End Function

