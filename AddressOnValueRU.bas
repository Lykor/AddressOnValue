Attribute VB_Name = "Module1"
Sub AddressOnValueRU()
Attribute AddressOnValueRU.VB_ProcData.VB_Invoke_Func = "q\n14"
Dim c As Range
Dim a As String


For Each c In ActiveCell
    If IsEmpty(c.FormulaLocal) Then
    End
    Else
    
        c.FormulaLocal = Replace(c.FormulaLocal, "+", "&""+""&")
        c.FormulaLocal = Replace(c.FormulaLocal, "-", "&""-""&")
        c.FormulaLocal = Replace(c.FormulaLocal, "/", "&""/""&")
        c.FormulaLocal = Replace(c.FormulaLocal, "*", "&""*""&")
        c.FormulaLocal = Replace(c.FormulaLocal, "йнпемэ", "в(""йнпемэ"")&")
        c.FormulaLocal = Replace(c.FormulaLocal, "яреоемэ", "яслл")
        c.FormulaLocal = Replace(c.FormulaLocal, ";2", ";в(""яреоемэ"")")
        c.FormulaLocal = Replace(c.FormulaLocal, "йнпемэ(3)", "в(""йнпемэРПХ"")")
        
        
        
             
        
        MsgBox (ActiveCell)
        c.FormulaLocal = Replace(c.FormulaLocal, "&""+""&", "+")
        c.FormulaLocal = Replace(c.FormulaLocal, "&""-""&", "-")
        c.FormulaLocal = Replace(c.FormulaLocal, "&""/""&", "/")
        c.FormulaLocal = Replace(c.FormulaLocal, "&""*""&", "*")
        c.FormulaLocal = Replace(c.FormulaLocal, "в(""йнпемэ"")&", "йнпемэ")
        c.FormulaLocal = Replace(c.FormulaLocal, "яслл", "яреоемэ")
        c.FormulaLocal = Replace(c.FormulaLocal, ";в(""яреоемэ"")", ";2")
        c.FormulaLocal = Replace(c.FormulaLocal, "в(""йнпемэРПХ"")", "йнпемэ(3)")
        
               
    End If
Next
End Sub


