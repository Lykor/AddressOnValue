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
        c.FormulaLocal = Replace(c.FormulaLocal, "������", "�(""������"")&")
        c.FormulaLocal = Replace(c.FormulaLocal, "�������", "����")
        c.FormulaLocal = Replace(c.FormulaLocal, ";2", ";�(""�������"")")
        c.FormulaLocal = Replace(c.FormulaLocal, "������(3)", "�(""���������"")")
        
        
        
             
        
        MsgBox (ActiveCell)
        c.FormulaLocal = Replace(c.FormulaLocal, "&""+""&", "+")
        c.FormulaLocal = Replace(c.FormulaLocal, "&""-""&", "-")
        c.FormulaLocal = Replace(c.FormulaLocal, "&""/""&", "/")
        c.FormulaLocal = Replace(c.FormulaLocal, "&""*""&", "*")
        c.FormulaLocal = Replace(c.FormulaLocal, "�(""������"")&", "������")
        c.FormulaLocal = Replace(c.FormulaLocal, "����", "�������")
        c.FormulaLocal = Replace(c.FormulaLocal, ";�(""�������"")", ";2")
        c.FormulaLocal = Replace(c.FormulaLocal, "�(""���������"")", "������(3)")
        
               
    End If
Next
End Sub


