´´´´

    Sub impuesto()
        p = Int(InputBox("valor a pagar anual"))
        
        Total = p * ip
        
        If p > 0 And p < 1000 Then
        MsgBox " no pagar impuesto"
        
        Else
        If p > 1001 And 10000 Then
            ip = 0.05
            Total = p * ip
            
            MsgBox " el resultado es: " & Total
            
        Else
            If p > 10001 And p < 100000 Then
            ip = 0.1
            Total = p * ip
            
            MsgBox " el resultado es: " & Total
            
            Else
                If p > 100001 And p < 1000000 Then
                    ip = 0.15
                    Total = p * ip
                    
                    MsgBox " el resultado es: " & Total
                
                Else
                    If p > 1000001 And p < 10000000 Then
                    ip = 0.2
                    Total = p * ip
                    
                    MsgBox " el resultado es: " & Total
                        
                    Else
                    If p > 10000001 And p < 100000000 Then
                        ip = 0.25
                        Total = p * ip
                    
                    MsgBox " el resultado es: " & Total
                    
                    End If
                    End If
                End If
            End If
        End If
        End If
    End Sub


´´´´