Attribute VB_Name = "Modul1"
' Relation for Excel
' Version 3.0 9.8.2022
' matti@belle-nuit.com
'This module provides functions to make simple relational algebra
'The relational model is simplified.
'A relation is defined as a 2d-table, columns can be adressed by name or position (1-based
'However, rows are not ordered (except with relOrder function) and do not have duplicates
'Unlike other Excel solutions, this module is purely functional, not using macros.
'Relations are saved as text in one cell with :: als field and newline as row separator
'Note that in a cell, the text cannot be more than 32k characters.
'Version 3.0 adds relSql, cache of big results and tables and accepts newlines in cells

Option Explicit

Dim prelstamps As New Collection
Dim prellaststamp As Double
Dim prellaststampcode As String
Dim prelevaluationcache As New Collection
Dim prelcaches As New Collection
 







Private Function prelCacheGet(ByVal key As String)
        Dim last As String
        last = prelStamp("prelcacheget")
    If prelInCollection(prelcaches, key) Then
        prelCacheGet = prelcaches(key)
    Else
        prelCacheGet = ""
    End If
    prelStamp (last)
End Function

Private Function prelCacheSet(ByVal s As String)
    Dim last As String
    last = prelStamp("prelcacheset")
    Dim key As String
    key = relRCR64(s)
    If Not prelInCollection(prelcaches, key) Then
        prelcaches.Add s, key
    End If
    prelStamp (last)
    prelCacheSet = key
End Function

Function prelCompareSql(v1 As Variant, v2 As Variant)
' empty < numbers < strings
Dim d1, d2 As Double

d1 = Val(v1)
d2 = Val(v2)

If d1 <> 0 Then
    If d2 <> 0 Then
        If d1 > d2 Then
            prelCompareSql = 1
        ElseIf d1 = d2 Then
            prelCompareSql = 0
        Else
            prelCompareSql = -1
        End If
    ElseIf Left(v2, 1) = "0" Then
        If d1 > 0 Then
            prelCompareSql = 1
        Else
            prelCompareSql = -1
        End If
    ElseIf v2 = "" Then
        prelCompareSql = 1
    Else 'non empty string
        prelCompareSql = -1
    End If
ElseIf Left(v1, 1) = "0" Then
    If d2 <> 0 Then
        If d2 > 0 Then
            prelCompareSql = -1
        Else
            prelCompareSql = 1
        End If
    ElseIf Left(v2, 1) = "0" Then
        prelCompareSql = 0
    ElseIf v2 = "" Then
        prelCompareSql = 1
    Else 'non empty string
        prelCompareSql = -1
    End If
ElseIf v1 = "" Then
    If v2 = "" Then
        prelCompareSql = 0
    Else
        prelCompareSql = -1
    End If
Else
    If d2 <> 0 Then
        prelCompareSql = 1
    ElseIf Left(v2, 1) = "0" Then
        prelCompareSql = 1
    ElseIf v2 = "" Then
        prelCompareSql = 1
    Else
        prelCompareSql = StrComp(v1, v2) ' to do compare constant
    End If
End If


End Function

Private Function prelRPN(ByVal expression As String, ByRef values() As String)
Dim list() As String
Dim elem As String
Dim key As String
Dim stack() As Variant
Dim ind As Long
Dim arg1, arg2, arg3 As Variant
Dim i, c, j As Long
Dim pattern As String
Dim argc As Long
Dim arglist() As String
Dim test As Long

On Error GoTo errHandler

   
list = Split(expression, vbTab)
c = UBound(list)
ReDim stack(0) ' cannot start with empty
For i = 0 To c
    elem = list(i)
    If Len(elem) > 0 Then
        key = Left(elem, 1)
    Else
        key = ""
    End If
    Select Case key
    Case ""
        ' pass
    Case "?"
        ' pass
    Case "$"
        ReDim Preserve stack(UBound(stack) + 1)
        stack(UBound(stack)) = Mid(elem, 3, Len(elem) - 3) ' $"..."
    Case "#"
        ReDim Preserve stack(UBound(stack) + 1)
        stack(UBound(stack)) = Mid(elem, 2)
    Case "@"
        ind = Val(Mid(elem, 2)) - 1 '
        ReDim Preserve stack(UBound(stack) + 1)
        If ind >= 0 And ind <= UBound(values) Then
            stack(UBound(stack)) = values(ind)
        Else
            prelRPN = "#ERROR illegal value " + elem
            Exit Function
        End If
    Case "%"
        ' argument count for next elem
        argc = Val(Mid(elem, 2))
    Case "!"
        Select Case Trim(Mid(elem, 2))
           Case "+"
                arg1 = stack(UBound(stack) - 1)
                arg2 = stack(UBound(stack))
                ReDim Preserve stack(UBound(stack) - 1)
                stack(UBound(stack)) = Str(Val(arg1) + Val(arg2))
           Case "-"
                arg1 = stack(UBound(stack) - 1)
                arg2 = stack(UBound(stack))
                ReDim Preserve stack(UBound(stack) - 1)
                stack(UBound(stack)) = Str(Val(arg1) - Val(arg2))
           Case "*"
                arg1 = stack(UBound(stack) - 1)
                arg2 = stack(UBound(stack))
                ReDim Preserve stack(UBound(stack) - 1)
                stack(UBound(stack)) = Str(Val(arg1) * Val(arg2))
           Case "/"
                arg1 = stack(UBound(stack) - 1)
                arg2 = stack(UBound(stack))
                ReDim Preserve stack(UBound(stack) - 1)
                If Val(arg2) <> 0 Then
                    stack(UBound(stack)) = Str(Val(arg1) * Val(arg2))
                Else
                    prelRPN = "#ERROR zero division "
                    Exit Function
                End If
            Case "="
                arg1 = stack(UBound(stack) - 1)
                arg2 = stack(UBound(stack))
                ReDim Preserve stack(UBound(stack) - 1)
                stack(UBound(stack)) = 1 - Abs(prelCompareSql(arg1, arg2)) ' 1 = equal
            Case "<>"
                arg1 = stack(UBound(stack) - 1)
                arg2 = stack(UBound(stack))
                ReDim Preserve stack(UBound(stack) - 1)
                stack(UBound(stack)) = Abs(prelCompareSql(arg1, arg2)) ' equal
            Case ">"
                arg1 = stack(UBound(stack) - 1)
                arg2 = stack(UBound(stack))
                ReDim Preserve stack(UBound(stack) - 1)
                If prelCompareSql(arg1, arg2) = 1 Then
                    stack(UBound(stack)) = 1
                Else
                    stack(UBound(stack)) = 0
                End If
            Case ">="
                arg1 = stack(UBound(stack) - 1)
                arg2 = stack(UBound(stack))
                ReDim Preserve stack(UBound(stack) - 1)
                If prelCompareSql(arg1, arg2) >= 0 Then
                    stack(UBound(stack)) = 1
                Else
                    stack(UBound(stack)) = 0
                End If
           Case "<"
                arg1 = stack(UBound(stack) - 1)
                arg2 = stack(UBound(stack))
                ReDim Preserve stack(UBound(stack) - 1)
                If prelCompareSql(arg1, arg2) = -1 Then
                    stack(UBound(stack)) = 1
                Else
                    stack(UBound(stack)) = 0
                End If
            Case "<="
                arg1 = stack(UBound(stack) - 1)
                arg2 = stack(UBound(stack))
                ReDim Preserve stack(UBound(stack) - 1)
                If prelCompareSql(arg1, arg2) <= 0 Then
                    stack(UBound(stack)) = 1
                Else
                    stack(UBound(stack)) = 0
                End If
            Case "LIKE"
                arg1 = stack(UBound(stack) - 1)
                arg2 = stack(UBound(stack))
                
                ' we have to modify pattern
                ' excel ? is /./, * is /.*/, # is /\d/, [abc] is /[abc], [!abc] is [^abc]
                ' sql is % is /.*/, _ is /./
                
                arg2 = Replace(arg2, "%", "*")
                arg2 = Replace(arg2, "_", "?")
                
                ReDim Preserve stack(UBound(stack) - 1)
                stack(UBound(stack)) = relLike(arg1, arg2)
            Case "AND"
                arg1 = stack(UBound(stack) - 1)
                arg2 = stack(UBound(stack))
                ReDim Preserve stack(UBound(stack) - 1)
                stack(UBound(stack)) = Val(arg1) * Val(arg2)
           Case "OR"
                arg1 = stack(UBound(stack) - 1)
                arg2 = stack(UBound(stack))
                ReDim Preserve stack(UBound(stack) - 1)
                stack(UBound(stack)) = 1 - (1 - Val(arg1)) * (1 - Val(arg2))
           Case "IN"
                ReDim arglist(argc - 1)
                For j = 0 To UBound(arglist)
                    arglist(j) = stack(UBound(stack) - j)
                Next j
                ReDim Preserve stack(UBound(stack) - argc)
                arg1 = stack(UBound(stack))
                test = 0
                For j = 0 To UBound(arglist)
                    If prelCompareSql(arg1, arglist(j)) = 0 Then
                        test = 1
                    End If
                Next j
                stack(UBound(stack)) = test
        Case Else
            prelRPN = "#ERROR RPN not implented " + elem
            Exit Function
        End Select
    Case "<"
        Select Case Mid(elem, 2)
            Case "NOT"
                arg1 = stack(UBound(stack))
                stack(UBound(stack)) = 1 - Val(arg1)
        Case Else
            prelRPN = "#ERROR RPN not implented " + elem
            Exit Function
        End Select
    Case "*"
        Select Case Mid(elem, 2)
            Case "ABS"
                arg1 = stack(UBound(stack))
                stack(UBound(stack)) = Math.Abs(Val(arg1))
            Case "COS"
                arg1 = stack(UBound(stack))
                stack(UBound(stack)) = Math.Cos(Val(arg1))
           Case "EXP"
                arg1 = stack(UBound(stack))
                stack(UBound(stack)) = Math.Exp(Val(arg1))
            Case "LEFT"
                arg1 = stack(UBound(stack) - 1)
                arg2 = stack(UBound(stack))
                ReDim Preserve stack(UBound(stack) - 1)
                stack(UBound(stack)) = Left(arg1, Val(arg2))
           Case "LEN"
                arg1 = stack(UBound(stack))
                stack(UBound(stack)) = Len(arg1)
           Case "INT"
                arg1 = stack(UBound(stack))
                stack(UBound(stack)) = Int(Val(arg1))
           Case "LN"
                arg1 = stack(UBound(stack))
                If Val(arg1) <= 0 Then
                    prelRPN = "#ERROR log on negative or zero value " & arg1
                    Exit Function
                End If
                stack(UBound(stack)) = Math.Log(Val(arg1))
            Case "LOG"
                arg1 = stack(UBound(stack))
                If Val(arg1) <= 0 Then
                    prelRPN = "#ERROR log on negative or zero value " & arg1
                    Exit Function
                End If
                stack(UBound(stack)) = Math.Log(Val(arg1)) / Math.Log(10)
           Case "LOWER"
                arg1 = stack(UBound(stack))
                stack(UBound(stack)) = LCase(arg1)
           Case "MID"
                arg1 = stack(UBound(stack) - 2)
                arg2 = stack(UBound(stack) - 1)
                arg3 = stack(UBound(stack))
                ReDim Preserve stack(UBound(stack) - 1)
                stack(UBound(stack)) = Mid(arg1, Val(arg2), Val(arg3))
           Case "MOD"
                arg1 = stack(UBound(stack) - 1)
                arg2 = stack(UBound(stack))
                ReDim Preserve stack(UBound(stack) - 1)
               If Val(arg2) <> 0 Then
                    stack(UBound(stack)) = Str(Val(arg1) Mod Val(arg2))
                Else
                    prelRPN = "#ERROR zero division "
                    Exit Function
                End If
            Case "POW"
                arg1 = stack(UBound(stack) - 1)
                arg2 = stack(UBound(stack))
                ReDim Preserve stack(UBound(stack) - 1)
                stack(UBound(stack)) = Val(arg1) ^ Val(arg2)
            Case "RIGHT"
                arg1 = stack(UBound(stack) - 1)
                arg2 = stack(UBound(stack))
                ReDim Preserve stack(UBound(stack) - 1)
                stack(UBound(stack)) = Right(arg1, Val(arg2))
            Case "ROUND"
                arg1 = stack(UBound(stack) - 1)
                arg2 = stack(UBound(stack))
                ReDim Preserve stack(UBound(stack) - 1)
                stack(UBound(stack)) = Math.Round(Val(arg1), Val(arg2))
            Case "SGN"
                arg1 = stack(UBound(stack))
                stack(UBound(stack)) = Math.Sgn(Val(arg1))
            Case "SIN"
                arg1 = stack(UBound(stack))
                stack(UBound(stack)) = Math.Sin(Val(arg1))
            Case "SQRT"
                arg1 = stack(UBound(stack))
                If Val(arg1) < 0 Then
                    prelRPN = "#ERROR sqrt on negative value " & arg1
                    Exit Function
                End If
                stack(UBound(stack)) = Math.Sqr(Val(arg1))
            Case "TAN"
                arg1 = stack(UBound(stack))
                stack(UBound(stack)) = Math.Tan(Val(arg1))
            Case "TRIM"
                arg1 = stack(UBound(stack))
                stack(UBound(stack)) = Trim(arg1)
            Case "UPPER"
                arg1 = stack(UBound(stack))
                stack(UBound(stack)) = UCase(arg1)
        Case Else
            prelRPN = "#ERROR RPN not implented " + elem
            Exit Function
        End Select
    Case Else
        prelRPN = "#ERROR RPN not implented " + elem
        Exit Function
    End Select
Next i

prelRPN = stack(UBound(stack))
Exit Function

   
errHandler:
  prelRPN = "#Error RPN " & Err.Number & ": " & Err.Description
   
   
   
End Function

Public Function relCRC64(s As String)

Dim l As Integer, l3 As Integer
Dim s1 As String, s2 As String, s3 As String, s4 As String

l = Int(Len(s) / 4)
s1 = Mid(s, 1, l)
s2 = Mid(s, 2 * l + 1, l)
s3 = Mid(s, 3 * l + 1, l)
s4 = Mid(s, 4 * l + 1, l + 4)

relCRC64 = prelHash4(s1) + prelHash4(s2) + prelHash4(s3) + prelHash4(s4)

End Function

Public Function relCRC32(s As String)

Dim l As Integer, l3 As Integer
Dim s1 As String, s2 As String

l = Int(Len(s) / 2)
s1 = Mid(s, 1, l)
s2 = Mid(s, l + 1, l + 1)

relCRC32 = prelHash4(s1) + prelHash4(s2)

End Function

Private Function prelHash4(txt)
Dim x As Long
Dim mask, i, j, nC, crc As Integer
Dim c As String

crc = &HFFFF

For nC = 1 To Len(txt)
    j = Asc(Mid(txt, nC))
    crc = crc Xor j
    For j = 1 To 8
        mask = 0
        If crc / 2 <> Int(crc / 2) Then mask = &HA001
        crc = Int(crc / 2) And &H7FFF: crc = crc Xor mask
    Next j
Next nC

c = Hex$(crc)

While Len(c) < 4
  c = "0" & c
Wend

prelHash4 = c

End Function


Public Function relCacheSize(rn As Range)
    relCacheSize = prelcaches.Count
End Function

Private Function prelEvaluate(ByVal expression As String)
    Dim last As String
    If Left(expression, 1) <> "=" Then
        prelEvaluate = expression
        Exit Function
    End If
    
    last = prelStamp("prelevaluate")
       
    If prelInCollection(prelevaluationcache, expression) Then
        prelEvaluate = prelevaluationcache(expression)
    Else
        prelEvaluate = Application.Evaluate(expression)
        prelevaluationcache.Add prelEvaluate, expression
    End If
    prelStamp (last)
End Function

Public Function prelFunctionArity(fn As String)
Select Case fn
Case "COUNT", "SUM", "AVG", "MIN", "MAX", "MEDIAN", "STDEV"
    prelFunctionArity = 1
Case "ABS", "COS", "EXP", "INT", "LN", "LOG", "MOD", "SGN", "SIN", "SQRT", "TAN"
    prelFunctionArity = 1
Case "LEN", "LOWER", "TRIM", "UPPER"
    prelFunctionArity = 2
Case "MOD", "POW", "ROUND"
    prelFunctionArity = 2
Case "LEFT", "RIGHT"
    prelFunctionArity = 2
Case "MID", "REPLACE"
    prelFunctionArity = 3
Case "NOP"
    prelFunctionArity = 0
End Select

End Function

Private Function prelCompileExpression(ByVal tokenstring As String, Optional ByVal maxaggregator As Long = 0)
Dim tokens(), token, state, ttype, tvalue, pexpression As String
Dim deeplevel, aggregatorcount As Long
Dim stack(), functionstack() As String
Dim functioncommas() As Long
Dim list() As String
Dim found As Boolean

list = Split(tokenstring, vbTab)

ReDim functionstack(0)
ReDim functioncommas(0)

state = "start"
deeplevel = 0
aggregatorcount = 0
pexpression = ""

For Each token In list
    ttype = Left(token, 1)
    tvalue = Mid(token, 2)
    
    Select Case state
    Case "start", "comma"
        Select Case ttype
        Case "("
            deeplevel = deeplevel + 1
            ReDim Preserve functionstack(UBound(functionstack) + 1)
            functionstack(UBound(functionstack)) = ""
            ReDim Preserve functioncommas(UBound(functioncommas) + 1)
            functioncommas(UBound(functioncommas)) = 0
            pexpression = pexpression + vbTab + token
            state = "expression"
        Case "#", "$", "@"
            pexpression = pexpression + vbTab + token
            state = "value"
        Case "+"
            aggregatorcount = aggregatorcount + 1
            If aggregatorcount > maxaggregator Then
                prelCompileExpression = "#ERROR invalid aggregator " + token
                Exit Function
            End If
            ReDim Preserve functionstack(UBound(functionstack) + 1)
            functionstack(UBound(functionstack)) = tvalue
            ReDim Preserve functioncommas(UBound(functioncommas) + 1)
            functioncommas(UBound(functioncommas)) = 0
            pexpression = pexpression + vbTab + token
            state = "function"
        Case "*"
            ReDim Preserve functionstack(UBound(functionstack) + 1)
            functionstack(UBound(functionstack)) = tvalue
            ReDim Preserve functioncommas(UBound(functioncommas) + 1)
            functioncommas(UBound(functioncommas)) = 0
            pexpression = pexpression + vbTab + token
            state = "function"
        Case "<"
            pexpression = pexpression + vbTab + token
           state = "expression"
        Case Else
            prelCompileExpression = "#ERROR invalid token " + token
            Exit Function
        End Select
    Case "expression"
        Select Case ttype
        Case "("
            deeplevel = deeplevel + 1
            ReDim Preserve functionstack(UBound(functionstack) + 1)
            functionstack(UBound(functionstack)) = ""
            ReDim Preserve functioncommas(UBound(functioncommas) + 1)
            functioncommas(UBound(functioncommas)) = 0
            pexpression = pexpression + vbTab + token
        Case ")"
            If deeplevel = 0 Then
                prelCompileExpression = "#ERROR invalid ) " + pexpression + " " + token
                Exit Function
            End If
            deeplevel = deeplevel - 1
            If functionstack(UBound(functionstack)) <> "" Then
                If functioncommas(UBound(functioncommas)) <> prelFunctionArity(functionstack(UBound(functionstack))) + 1 Then
                    prelCompileExpression = "#ERROR function arity " + pexpression + " " + token
                    Exit Function
                End If
            Else
                ' show argument count
                pexpression = pexpression + vbTab + "%" + Trim(Str(functioncommas(UBound(functioncommas)) + 1))
            End If
            ReDim Preserve functionstack(UBound(functionstack) - 1)
            ReDim Preserve functioncommas(UBound(functioncommas) - 1)
            functioncommas(UBound(functioncommas)) = 0
            pexpression = pexpression + vbTab + token
            state = "value"
        Case "#", "$", "@"
            pexpression = pexpression + vbTab + token
            state = "value"
        Case "+"
            aggregatorcount = aggregatorcount + 1
            If aggregatorcount > maxaggregator Then
                prelCompileExpression = "#ERROR invalid aggregator " + pexpression + " " + token
                Exit Function
            End If
            ReDim Preserve functionstack(UBound(functionstack) + 1)
            functionstack(UBound(functionstack)) = tvalue
            ReDim Preserve functioncommas(UBound(functioncommas) + 1)
            functioncommas(UBound(functioncommas)) = 0
            pexpression = pexpression + vbTab + token
            state = "function"
        Case "*"
            pexpression = pexpression + vbTab + token
            ReDim Preserve functionstack(UBound(functionstack) + 1)
            functionstack(UBound(functionstack)) = tvalue
            ReDim Preserve functioncommas(UBound(functioncommas) + 1)
            functioncommas(UBound(functioncommas)) = 0
            state = "function"
        Case "<"
            pexpression = pexpression + vbTab + token
        Case Else
            prelCompileExpression = "#ERROR invalid token " + pexpression + " " + token
            Exit Function
        End Select
    Case "value"
        Select Case ttype
        Case "!"
            pexpression = pexpression + vbTab + token
            state = "expression"
        Case ","
            If deeplevel = 0 Then
                prelCompileExpression = "#ERROR comma on root level " + pexpression + " " + token
                Exit Function
            End If
            functioncommas(deeplevel) = functioncommas(deeplevel) + 1
            ' pexpression = pexpression + vbTab + token
            state = "comma"
        Case ")"
            If deeplevel = 0 Then
                prelCompileExpression = "#ERROR invalid ) " + pexpression + " " + token
                Exit Function
            End If
            deeplevel = deeplevel - 1
            If functionstack(UBound(functionstack)) <> "" Then
                If functioncommas(UBound(functioncommas)) + 1 <> prelFunctionArity(functionstack(UBound(functionstack))) Then
                    prelCompileExpression = "#ERROR function arity " + pexpression + " " + token
                    Exit Function
                End If
            Else
                ' show argument count
                pexpression = pexpression + vbTab + "%" + Trim(Str(functioncommas(UBound(functioncommas)) + 1))
            End If
            ReDim Preserve functionstack(UBound(functionstack) - 1)
            ReDim Preserve functioncommas(UBound(functioncommas) - 1)
            pexpression = pexpression + vbTab + token
            state = "value"
        Case Else
            prelCompileExpression = "#ERROR invalid token " + pexpression + " " + token
            Exit Function
        End Select
    Case "function"
        Select Case ttype
        Case "("
            deeplevel = deeplevel + 1
            pexpression = pexpression + vbTab + token
            state = "expression"
        Case Else
            prelCompileExpression = "#ERROR ( expected " + pexpression + " " + token
            Exit Function
        End Select
    Case Else
        prelCompileExpression = "#ERROR invalid state " + state + " " + pexpression + " " + token
        Exit Function
    End Select
Next token

If deeplevel <> 0 Then
    prelCompileExpression = "#ERROR ) expected " + pexpression
    Exit Function
End If

Select Case state
Case "start", "expression", "value"
    ' ok
Case "comma", "function"
    prelCompileExpression = "#ERROR invalid state " + state + " " + pexpression
    Exit Function
End Select

' expression is syntactally correct, now we can compile it to RPN
' firstletter is vbTab
prelCompileExpression = "RPN(" + Mid(pexpression, 2) + ")"

' shunting yard
' outstack and waitstack
Dim rpnlist() As String
Dim operatorstack() As String
Dim pr As Long
Dim r, o As Long

list = Split(pexpression, vbTab)

ReDim rpnlist(UBound(list))
ReDim operatorstack(UBound(list))
r = -1
o = -1

For Each token In list
    ttype = Left(token, 1)
    tvalue = Mid(token, 2)
    
    Select Case ttype
    
    Case "$", "#", "@", "%"
        r = r + 1
        rpnlist(r) = token
    Case "+", "*", "("
        o = o + 1
        operatorstack(o) = token
    Case ")"
        found = False
        While o > 0 And Not found
            If Left(operatorstack(o), 1) <> "(" Then
                r = r + 1
                rpnlist(r) = operatorstack(o)
                o = o - 1
            Else
                found = True
            End If
        Wend
        o = o - 1 ' ((
        If o > -1 Then
            If Left(operatorstack(o), 1) = "+" Or Left(operatorstack(o), 1) = "*" Then
                r = r + 1
                rpnlist(r) = operatorstack(o)
                o = o - 1
            End If
        End If
    Case "!", "<", ","
        pr = prelPrecedence(token)
        found = False
        While o > -1 And Not found ' cannot take next condition on the line, would be executed (-1)
            If Left(operatorstack(o), 1) = "!" And prelPrecedence(operatorstack(o)) >= pr Then
                r = r + 1
                rpnlist(r) = operatorstack(o)
                o = o - 1
            Else
                found = True ' stop
            End If
        Wend
        o = o + 1
        operatorstack(o) = token
    Case Else
        prelCompileExpression = "#ERROR rpn generation " + token
    End Select

Next token

While o > -1
   r = r + 1
   rpnlist(r) = operatorstack(o)
   o = o - 1
Wend
    

ReDim Preserve rpnlist(r)

prelCompileExpression = "?" + vbTab + Join(rpnlist, vbTab)
        
        
        
        
        
        
    
        
    
   
    
    
    



End Function

Public Function relEvaluationCacheSize(rn As Range)
    relEvaluationCacheSize = prelevaluationcache.Count
End Function

Private Function prelPrecedence(ByVal t As String)
Dim ttype, tvalue As String

ttype = Left(t, 1)
tvalue = Mid(t, 2)

prelPrecedence = 0

Select Case ttype
Case "!", "<"
    Select Case tvalue
    Case "*", "/"
        prelPrecedence = 52
    Case "+", "-"
        prelPrecedence = 51
    Case "<", ">", "<>", ">=", "<=", "="
        prelPrecedence = 42
    Case "LIKE"
        prelPrecedence = 41
    Case "IN"
        prelPrecedence = 41
    Case "NOT"
        prelPrecedence = 31
    Case "AND"
        prelPrecedence = 22
    Case "OR"
        prelPrecedence = 21
    End Select
Case ","
    prelPrecedence = 11
End Select

    

    


End Function

Public Function relProfile(rn As Range)
    Dim s As String
    Dim st As Variant
    Dim tsa() As String
    Dim totaltime As Double
    Dim p As Long
    For Each st In prelstamps
        tsa = Split(st, " ")
        totaltime = totaltime + CDbl(tsa(1))
    Next st
    If totaltime = 0 Then totaltime = 1
    For Each st In prelstamps
      tsa = Split(st, " ")
      p = Val(tsa(1)) / totaltime * 100
       If p > 9 Then
        s = s + Str(p) + " " + tsa(2)
       End If
    Next st
    relProfile = s
End Function

Function prelStamp(s As String)
    Dim t As Double
    Dim t0 As Double
    Dim ts As String
    Dim tsa() As String
    
    On Error GoTo errHandler
    prelStamp = prellaststampcode
    
    If prellaststamp = 0 Then
        prellaststamp = VBA.Timer
        prellaststampcode = s
    Else
    
     t = VBA.Timer - prellaststamp
     prellaststamp = t
     If prelInCollection(prelstamps, prellaststampcode) Then
        ts = prelstamps.Item(prellaststampcode)
        tsa = Split(ts, " ")
        t0 = Val(tsa(1))
        prelstamps.Remove prellaststampcode
        prelstamps.Add Str(t0 + t) + " " + prellaststampcode, prellaststampcode
        
     Else
        prelstamps.Add Str(t) + " " + prellaststampcode, prellaststampcode
     End If
    prellaststampcode = s
    End If
    
    Exit Function
    
errHandler:
    MsgBox ("error stamp " + s)
    
End Function

Public Sub relResetStamps()
    Set prelstamps = Nothing
End Sub

Public Function relRange(rn As Range, Optional hasheader As Long = True)

relRange = prelRange(rn, hasheader, False, False)

End Function

Private Function prelRange(rn As Range, hasheader As Long, noError As Boolean, lazy As Boolean) As String

' Calculates a relation from a range
' A relation is a table where rows are separated by newline and columns by ::
' We use a simplified model where tuples can have no named properties, but by position (1-based)
' If header is false, default number header will be used
' If header is true, first line of rn is considered header
' optional noerror is necessary for relFilter
' lazy is necessary for relFilter, removing duplicates only when necessary


Dim arr() As Variant
Dim hd() As Variant
Dim tuples() As String
Dim fields() As String
Dim r, c, i, j, first As Long
Dim v As Variant
Dim l As String
Dim found As Boolean
Dim last As String

last = prelStamp("prelrange")

arr = rn


'range returns an array which is 1-based
r = UBound(arr, 1)
c = UBound(arr, 2)

ReDim tuples(r - 1)
ReDim fields(c - 1)


If hasheader Then
    For j = 0 To c - 1
        fields(j) = arr(1, j + 1)
        If Val(fields(j)) > 0 Then
            prelRange = "#ERROR NUMERIC NAME"
            Exit Function
        End If
        For i = 0 To c - 1
            If j <> i And fields(j) = fields(i) Then
                prelRange = "#ERROR DUPLICATE COLUMN " & fields(j)
                Exit Function
            End If
        Next i
    Next j
    first = 1
 Else
    ReDim hd(c)
    For i = 1 To c
        hd(i) = "c" & Trim(Str(i))
    Next
    For j = 0 To c - 1
        fields(j) = hd(j + 1)
    Next j
    first = 0
End If
tuples(0) = Join(fields, "::")

For i = 1 To r - 1
    For j = 0 To c - 1
        v = arr(i + first, j + 1)
        If IsError(v) Then
            prelRange = "#ERROR CELL "
            Exit Function
        End If
        fields(j) = Replace(Replace(v, vbCrLf, ", "), vbCr, ", ")
        If fields(j) <> "" Then found = True
    Next j
    tuples(i) = Join(fields, "::")
Next i

If lazy Then
    l = Join(tuples, prelNewline())
Else
    l = prelString(tuples)
End If

If Len(l) > 32768 And Not noError Then
        prelRange = "#ERROR LONG RESULT " & Str(Len(l))
        Exit Function
End If
 
prelRange = l
prelStamp (last)

End Function

Public Function relParseSql(ByVal code As String, Optional ByVal multiline As Long = 0)
Dim tokenstring, token, ttype, tvalue, state, test As String
Dim elem, elem2 As Variant
Dim pexpression, pid As String
Dim jexpression, jid, jmode As String
Dim wexpression As String
Dim hexpression As String
Dim omode, oid As String
Dim tokens() As String
Dim relation As String
Dim joins As New Collection
Dim selects As New Collection
Dim selects2 As New Collection
Dim havings As New Collection
Dim projects As New Collection
Dim orders As New Collection
Dim lines() As String
Dim fields() As String
Dim extensions As New Collection
Dim projectcols As New Collection
Dim projectfinalcols As New Collection
Dim afterproject As New Collection
Dim ordercols As New Collection
Dim deeplevel, aggregatorcount As Long
Dim i, j, k As Long
Dim even As Boolean
Dim vtest As Variant
Dim found As Boolean
Dim tempname As String

If Left(code, 6) = "#ERROR" Then
    relParseSql = code
    Exit Function
End If

tokenstring = relTokenizeSql(code, 1)

If Left(tokenstring, 6) = "#ERROR" Then
    relParseSql = tokenstring
    Exit Function
End If


tokens = Split(tokenstring, vbCrLf)


state = "start"
For Each token In tokens
    If token <> "" Then
    ttype = Left(token, 1)
    If InStr("$#@!<+*(,)", ttype) Then
        tvalue = Mid(token, 2)
    Else
        ttype = token
        tvalue = ""
    End If
        
    Select Case state
    Case "start"
        Select Case ttype
        Case "SELECT"
            state = "select"
        Case Else
            relParseSql = "#ERROR missing select " + token
            Exit Function
        End Select
    Case "select"
        deeplevel = 0
        aggregatorcount = 0
        Select Case ttype
        Case "@"
            pexpression = token
            pid = tvalue
            state = "selectid"
        Case "(", "#", "$", "+", "*", "<"
            pexpression = token
            pid = tvalue
            state = "selectexpression"
        Case Else
            If tvalue = "*" Then
                pexpression = token
                state = "selectall"
            Else
                relParseSql = "#ERROR invalid token " + token
                Exit Function
            End If
        End Select
    Case "selectid"
        Select Case ttype
        Case "!"
            pexpression = pexpression + vbTab + token
            state = "selectexpression"
        Case "AS"
            state = "selectas"
        Case ","
            projects.Add pexpression + vbCrLf + pid
            state = "select"
        Case "FROM"
            projects.Add pexpression + vbCrLf + pid
            state = "from"
        Case Else
            relParseSql = "#ERROR invalid token " + pexpression + " " + token
            Exit Function
        End Select
     Case "selectexpression"
        Select Case ttype
        Case "$", "#", "@", "!", "<", "+", "*", "(", ",", ")"
            pexpression = pexpression + vbTab + token
            state = "selectexpression"
        Case "AS"
            state = "selectas"
        Case Else
            relParseSql = "#ERROR invalid token " + pexpression + " " + token
            Exit Function
        End Select
   Case "selectas"
        Select Case ttype
        Case "@"
            pid = tvalue
            test = prelCompileExpression(pexpression, 1)
            If Left(test, 6) = "#ERROR" Then
                relParseSql = test
                Exit Function
            End If
            projects.Add test + vbCrLf + pid
            state = "selectasid"
        Case Else
            relParseSql = "#ERROR missing alias id " + token
            Exit Function
        End Select
    Case "selectasid"
        Select Case ttype
        Case ","
            state = "select"
        Case "FROM"
            state = "from"
        Case Else
            relParseSql = "#ERROR missing COMMA or FROM " + token
            Exit Function
        End Select
    Case "selectall"
        Select Case ttype
        Case "FROM"
            state = "from"
        Case Else
            relParseSql = "#ERROR missing FROM " + token
            Exit Function
        End Select
    Case "from"
        Select Case ttype
        Case "@"
            relation = tvalue
            state = "fromid"
        Case Else
            relParseSql = "#ERROR missing table after from " + token
            Exit Function
        End Select
    Case "fromid"
        Select Case ttype
        Case "NATURAL", "LEFT", "RIGHT", "OUTER"
            jmode = ttype
            state = "joinmode"
        Case "JOIN"
            state = "join"
        Case "WHERE"
            state = "where"
        Case "HAVING"
            state = "having"
        Case "ORDER"
            state = "order"
        Case Else
            relParseSql = "#ERROR invalid token after table " + token
            Exit Function
        End Select
    Case "joinmode"
        Select Case ttype
        Case "JOIN"
            state = "join"
        Case Else
            relParseSql = "#ERROR missing JOIN " + ttype
            Exit Function
        End Select
    Case "join"
        Select Case ttype
        Case "@"
            jid = tvalue
            If jmode = "NATURAL" Then ' special join
                joins.Add jid + vbCrLf + jmode + vbCrLf + ""
                jmode = ""
                state = "fromid"
            Else
                state = "joinid"
            End If
        Case Else
            relParseSql = "#ERROR missing join table " + token
            Exit Function
        End Select
    Case "joinid"
        Select Case ttype
        Case "ON"
            state = "joinon"
        Case Else
            relParseSql = "#ERROR missing ON " + token
            Exit Function
        End Select
    Case "joinon"
        Select Case ttype
        Case "$", "#", "@", "<", "+", "*", "("
            jexpression = token
            state = "joinonexpression"
        Case Else
            relParseSql = "#ERROR missing expression after ON " + token
            Exit Function
        End Select
    Case "joinonexpression"
        Select Case ttype
        Case "$", "#", "@", "!", "<", "+", "*", "(", ",", ")"
            jexpression = jexpression + vbTab + token
        Case "NATURAL", "LEFT", "RIGHT", "OUTER", "JOIN", "WHERE", "ORDER"
            test = prelCompileExpression(jexpression)
            If Left(test, 6) = "#ERROR" Then
                relParseSql = test
                Exit Function
            End If
            joins.Add jid + vbCrLf + jmode + vbCrLf + test

            jexpression = ""
            jmode = ttype
            Select Case ttype
            Case "NATURAL", "LEFT", "RIGHT", "OUTER"
                state = "joinmode"
            Case "JOIN"
                state = "join"
            Case "WHERE"
                state = "where"
            Case "HAVING"
                state = "having"
            Case "ORDER"
                state = "order"
            End Select
        Case Else
            relParseSql = "#ERROR invalid token " + token
            Exit Function
        End Select
     Case "where"
        Select Case ttype
        Case "$", "#", "@", "<", "+", "*", "("
            wexpression = token
            state = "whereexpression"
            ' optimizing
            If ttype = "(" Then
                deeplevel = 1
            Else
                deeplevel = 0
            End If
        Case Else
            relParseSql = "#ERROR missing expression after WHERE " + token
            Exit Function
        End Select
     Case "whereexpression"
        Select Case ttype
        Case "$", "#", "@", "!", "<", "+", "*", "(", ",", ")"
            ' optimizing
            If ttype = "!" And tvalue = "AND" And deeplevel = 0 Then
                test = prelCompileExpression(wexpression)
                If Left(test, 6) = "#ERROR" Then
                    relParseSql = test
                    Exit Function
                End If
                selects.Add test, test
                state = "where"
            Else
                If ttype = "(" Then deeplevel = deeplevel + 1
                If ttype = ")" Then deeplevel = deeplevel - 1
                wexpression = wexpression + vbTab + token
            End If
        Case "ORDER"
            test = prelCompileExpression(wexpression)
            If Left(test, 6) = "#ERROR" Then
                relParseSql = test
                Exit Function
            End If
            selects.Add test, test
            state = "order"
        Case "HAVING"
            test = prelCompileExpression(wexpression)
            If Left(test, 6) = "#ERROR" Then
                relParseSql = test
                Exit Function
            End If
            selects.Add test, test
            state = "having"
        Case Else
            relParseSql = "#ERROR invalid token " + token
            Exit Function
        End Select
    Case "having"
        Select Case ttype
        Case "$", "#", "@", "<", "+", "*", "("
            hexpression = token
            state = "havingexpression"
        Case Else
            relParseSql = "#ERROR missing expression after HAVING " + token
            Exit Function
        End Select
     Case "havingexpression"
        Select Case ttype
        Case "$", "#", "@", "!", "<", "+", "*", "(", ",", ")"
            ' optimizing
            hexpression = hexpression + vbTab + token
        Case "ORDER"
            test = prelCompileExpression(hexpression)
            If Left(test, 6) = "#ERROR" Then
                relParseSql = test
                Exit Function
            End If
            havings.Add test
            state = "order"
        Case Else
            relParseSql = "#ERROR invalid token " + token
            Exit Function
        End Select
    Case "order"
        Select Case ttype
        Case "BY"
            state = "orderby"
        Case Else
            relParseSql = "#ERROR invalid token after ORDER " + token
            Exit Function
        End Select
    Case "orderby"
        Select Case ttype
        Case "@"
            oid = tvalue
            omode = "ASC"
            state = "orderid"
        Case Else
            relParseSql = "#ERROR invalid token after ORDER BY " + token
            Exit Function
        End Select
    Case "orderid"
        Select Case ttype
        Case "ASC", "DESC"
            omode = ttype
            state = "ordermode"
        Case ","
            orders.Add oid + vbTab + omode
            state = "orderby"
        Case Else
            relParseSql = "#ERROR invalid token ORDER BY <id> " + token
            Exit Function
        End Select
    Case "ordermode"
        Select Case ttype
        Case ","
            orders.Add oid + vbTab + omode
            state = "orderby"
        Case Else
            relParseSql = "#ERROR invalid token after ORDER BY <id> <mode> " + token
            Exit Function
        End Select
    Case Else
        relParseSql = "#ERROR invalid state " + state
        Exit Function
    End Select
    End If
Next token

Select Case state
    Case "orderid"
        orders.Add oid + vbTab + omode
    Case "ordermode"
        orders.Add oid + vbTab + omode
    Case "start", "fromid"
        ' pass
    Case "joinexpression", "joinonexpression"
        test = prelCompileExpression(jexpression)
         If Left(test, 6) = "#ERROR" Then
            relParseSql = test
            Exit Function
         End If
        joins.Add jid + vbCrLf + jmode + vbCrLf + test
    Case "whereexpression"
         test = prelCompileExpression(wexpression)
         If Left(test, 6) = "#ERROR" Then
            relParseSql = test
            Exit Function
         End If
         selects.Add test, test
     Case "havingexpression"
         test = prelCompileExpression(hexpression)
         If Left(test, 6) = "#ERROR" Then
            relParseSql = test
            Exit Function
         End If
         havings.Add test
    Case Else
        relParseSql = "#ERROR invalid state " + state
End Select

relParseSql = "_ " + relation
relParseSql = relParseSql + vbCrLf + "R * " + relation

' take all selects with columns that start with relation
For Each elem In selects
    fields = Split(elem, vbTab)
    found = True
    For i = 0 To UBound(fields)
        If Left(fields(i), 1) = "@" Then
            If Left(fields(i), Len("@" + relation + ".")) <> "@" + relation + "." Then
                found = False
            End If
        End If
    Next i
    If found Then
        relParseSql = relParseSql + vbCrLf + "S " + elem
        selects.Remove (elem)
    End If
Next elem

For Each elem In joins
    lines = Split(elem, vbCrLf)
    relParseSql = relParseSql + vbCrLf + "_ " + lines(0)
    relParseSql = relParseSql + vbCrLf + "R * " + lines(0)
    
    ' take all selects with columns that start with relation
    For Each elem2 In selects
        fields = Split(elem2, vbTab)
        found = True
        For i = 0 To UBound(fields)
            If Left(fields(i), 1) = "@" Then
                If Left(fields(i), Len("@" + lines(0) + ".")) <> "@" + lines(0) + "." Then
                    found = False
                End If
            End If
        Next i
        If found Then
            relParseSql = relParseSql + vbCrLf + "S " + elem2
            selects.Remove (elem2)
        End If
    Next elem2
    
    If Trim(lines(1)) <> "NATURAL" Then
    ' try to optimize equi join
    ' if the expression has only AND and =, equi join is possible
        found = False
        fields = Split(lines(2), vbTab)
        For Each vtest In fields
            Select Case test
            Case "!AND", "!="
             'pass
            Case Else
                If Left(vtest, 1) = "@" Then
                    'pass
                Else
                    found = True
                End If
            End Select
        Next vtest

        If Not found Then
            'optimize
            ' create for each identifier a temp column for natural
            ' rename all other columns so that they are not alike for natural join
            ' make natural join
            ' to do
        End If
     End If
     
    relParseSql = relParseSql + vbCrLf + "J " + Trim(lines(1) + " " + lines(2))
Next elem

For Each elem In selects
    relParseSql = relParseSql + vbCrLf + "S " + elem
Next elem


For Each elem In projects
    lines = Split(elem, vbCrLf)
    fields = Split(lines(0), vbTab)
    If UBound(fields) > 0 Then ' expression
        ' is there an aggregator, then split in 3
        k = -1
        found = False
        For i = 0 To UBound(fields)
            If Left(fields(i), 1) = "+" Then
                found = True
                If i > 1 Or Left(fields(0), 1) <> "@" Then ' expression
                   relParseSql = relParseSql + vbCrLf + "E " + lines(1) + "_temp "
                    For k = 0 To i - 1
                        relParseSql = relParseSql + fields(k) + vbTab
                    Next k
                    tempname = lines(1) + "_temp"
                Else ' id
                    tempname = fields(0)
                End If
                projectcols.Add tempname + " " + Mid(fields(i), 2)
                If i < UBound(fields) Then
                    
                    test = "E " + lines(1) + " ?" + vbTab + "@" + tempname + "_" + LCase(Mid(fields(i), 2)) + vbTab
                    For k = i + 1 To UBound(fields)
                       test = test + fields(k) + vbTab
                    Next k
                    afterproject.Add test
                Else
                    afterproject.Add "R " + tempname + "_" + LCase(Mid(fields(i), 2)) + " " + lines(1)
                End If
                i = UBound(fields)
            End If
        Next i
        If Not found Then
            extensions.Add "E " + lines(1) + " " + lines(0), lines(1)
            projectcols.Add lines(1)
        End If
    Else
        tvalue = Mid(lines(0), 2)
        If tvalue <> lines(1) Then  ' expression is token, rename is value
            relParseSql = relParseSql + vbCrLf + "R " + lines(0) + " " + lines(1)
        End If
        projectcols.Add lines(1)
    End If
    projectfinalcols.Add lines(1)
Next elem


' now check if we need to extend earlier
' we check for each extend line if it is mentioned before
lines = Split(relParseSql, vbCrLf)
For i = 0 To UBound(lines)
    elem = lines(i)
    Select Case Left(elem, 1)
        Case "S", "J"
            If extensions.Count > 0 Then
            For Each elem2 In extensions
                fields = Split(elem2, " ")
                test = fields(1) ' 0 E 1 id 2... formula
                If InStr(elem + vbTab, "@" + test + vbTab) Then
                    ReDim Preserve lines(UBound(lines) + 1)
                    For j = i To UBound(lines) - 1
                        lines(j + 1) = lines(j)
                    Next j
                    lines(i) = elem2
                    extensions.Remove (test)
                End If
            Next elem2
            End If
    End Select
Next i
relParseSql = Join(lines, vbCrLf)
If extensions.Count > 0 Then
For Each elem In extensions
    relParseSql = relParseSql + vbCrLf + elem
Next
End If




If projectcols.Count > 0 Then
    test = ""
    For Each elem In projectcols
        test = test + elem + "::"
    Next elem
    relParseSql = relParseSql + vbCrLf + "P " + Left(test, Len(test) - 2)
End If

If afterproject.Count > 0 Then
    For Each elem In afterproject
        relParseSql = relParseSql + vbCrLf + elem
    Next elem
End If

For Each elem In havings
    relParseSql = relParseSql + vbCrLf + "S " + elem
Next elem

If afterproject.Count > 0 Then
    If projectfinalcols.Count > 0 Then
        test = ""
        For Each elem In projectfinalcols
            test = test + elem + "::"
        Next elem
        relParseSql = relParseSql + vbCrLf + "P " + Left(test, Len(test) - 2)
    End If
End If

' reconnecting following selects
lines = Split(relParseSql, vbCrLf)
test = ""
k = 0
For i = 0 To UBound(lines)
    elem = lines(i - k) ' VBA does not update the upper limit in for loop!
    If Left(elem, 1) = "S" And test = "S" Then
        lines(i - 1 - k) = lines(i - 1 - k) + Mid(lines(i), 4) + vbTab + "!AND" 'S ?
        For j = i To UBound(lines) - 1
            lines(j) = lines(j + 1)
        Next j
        ReDim Preserve lines(UBound(lines) - 1)
        k = k + 1
    End If
    test = Left(elem, 1)
Next i
relParseSql = Join(lines, vbCrLf)

test = ""
For Each elem In orders
    test = test + Replace(elem, vbTab, " ") + "::"
Next elem
If test <> "" Then
    relParseSql = relParseSql + vbCrLf + "O " + Left(test, Len(test) - 2)
End If

relParseSql = relParseSql + vbCrLf + "R -" ' clean up

End Function



Function relUnion(ByVal rel1 As String, ByVal rel2 As String, Optional noError As Boolean = False, Optional lazy As Boolean = False) As String
Dim first1, first2, r As String
Dim fields1() As String
Dim fields2() As String
Dim rows1() As String
Dim rows2() As String
Dim header1list() As String
Dim header2list() As String
Dim afields() As String
Dim nfields() As String
Dim ub11, ub12, ub2 As Long
Dim s, header1, header2, l As String
Dim c1, c2, r1, r2, i, j, n As Long
Dim columns() As Long

prelStamp ("relunion")

If rel1 = "" Then
    relUnion = rel2
    Exit Function
End If

If rel2 = "" Then
    relUnion = rel1
    Exit Function
End If

' both relations must have the same fields
' order of first relation is retained

rows1 = Split(rel1, prelNewline())
header1 = rows1(0)
rows2 = Split(rel2, prelNewline())
header2 = rows2(0)

header1list = Split(header1, "::")
header2list = Split(header2, "::")

r1 = UBound(rows1)
r2 = UBound(rows2)
c1 = UBound(header1list)
c2 = UBound(header2list)

ReDim columns(c2)

If c1 <> c2 Then
    relUnion = "#ERROR ARITY : " & Str(c1 + 1) & " <>" & Str(c2 + 1)
    Exit Function
End If

For i = 0 To c2
    n = prelNameToColumn(header1, header2list(i)) - 1
    If n < 0 Then
        relUnion = "#ERROR FIELD NOT MATCH " & header2list(i)
        Exit Function
    End If
    columns(i) = n
Next

ReDim Preserve rows1(r1 + r2)

For i = 1 To r2
    afields = Split(rows2(i), "::")
    ReDim nfields(c2)
    For j = 0 To c2
       nfields(columns(j)) = afields(j)
    Next
    rows1(r1 + i) = Join(nfields, "::")
Next i

If lazy Then
    l = Join(rows1, prelNewline())
Else
    l = prelString(rows1)
End If

If Len(l) > 32768 And Not noError Then
        relUnion = "#ERROR LONG RESULT " & Str(Len(l))
        Exit Function
End If
 
relUnion = l


End Function

Function relDifference(rel1 As String, rel2 As String) As String

Dim first1, first2, r As String
Dim fields1() As String
Dim fields2() As String
Dim rows1() As String
Dim rows2() As String
Dim header1list() As String
Dim header2list() As String
Dim arr() As String
Dim ub11, ub12, ub2 As Long
Dim s, header1, header2 As String
Dim found As Boolean
Dim r1, r2, c1, c2, n, i, j, k, l, offset As Long

If rel1 = "" Then
    relDifference = ""
    Exit Function
End If

If rel2 = "" Then
    relDifference = rel1
    Exit Function
End If

' both relations must have the same fields
' order of first relation is retained

rows1 = Split(rel1, prelNewline())
header1 = rows1(0)
rows2 = Split(rel2, prelNewline())
header2 = rows2(0)

header1list = Split(header1, "::")
header2list = Split(header2, "::")

r1 = UBound(rows1)
r2 = UBound(rows2)
c1 = UBound(header1list)
c2 = UBound(header2list)

ReDim columns(c2)

If c1 <> c2 Then
    relDifference = "#ERROR ARITY : " & Str(c1 + 1) + " <>" & Str(c2 + 1)
    Exit Function
End If

For i = 0 To c2
    n = prelNameToColumn(header1, header2list(i)) - 1
    If n < 0 Then
        relDifference = "#ERROR FIELD NOT MATCH " & header2list(i)
        Exit Function
    End If
    columns(i) = n
Next

rows1 = Split(rel1, prelNewline())
rows2 = Split(rel2, prelNewline())

r1 = UBound(rows1)
r2 = UBound(rows2)

'reorganize columnns in second relation so that they have the same order
For i = 0 To r2
    fields1 = Split(rows2(i), "::")
    ReDim fields2(c2)
    For j = 0 To c2
        fields2(columns(j)) = fields1(j)
    Next j
    rows2(i) = Join(fields2, "::")
Next i

ReDim rows(r1)
rows(0) = rows1(0)
offset = 1
For i = 1 To r1
    found = False
    ' for each tuple in the first relation we check if it is in the second relation
    For j = 0 To r2
        If rows1(i) = rows2(j) Then
            found = True
            Exit For
        End If
    Next j
    If Not found Then
        rows(offset) = rows1(i)
        offset = offset + 1
    End If
Next i

'If offset = 1 Then
'    relDifference = ""
'    Exit Function
'End If

ReDim Preserve rows(offset - 1)

relDifference = Join(rows, prelNewline()) 'no duplicate elimination needed


End Function

Function relIntersect(rel1 As String, rel2 As String) As String

Dim first1, first2, r As String
Dim fields1() As String
Dim fields2() As String
Dim arr() As String
Dim rows1() As String
Dim rows2() As String
Dim header1list() As String
Dim header2list() As String
Dim ub11, ub12, ub2 As Long
Dim s, header1, header2 As String
Dim found As Boolean
Dim r1, r2, c1, c2, i, j, k, l, offset, n As Long

If rel1 = "" Then
    relIntersect = ""
    Exit Function
End If

If rel2 = "" Then
    relIntersect = ""
    Exit Function
End If

' both relations must have the same fields
' order of first relation is retained

rows1 = Split(rel1, prelNewline())
header1 = rows1(0)
rows2 = Split(rel2, prelNewline())
header2 = rows2(0)

header1list = Split(header1, "::")
header2list = Split(header2, "::")

r1 = UBound(rows1)
r2 = UBound(rows2)
c1 = UBound(header1list)
c2 = UBound(header2list)

ReDim columns(c2)

If c1 <> c2 Then
    relIntersect = "#ERROR ARITY : " & Str(c1 + 1) & "<>" & Str(c2 + 1)
    Exit Function
End If

For i = 0 To c2
    n = prelNameToColumn(header1, header2list(i)) - 1
    If n < 0 Then
        relIntersect = "#ERROR FIELD NOT MATCH " & header2list(i)
        Exit Function
    End If
    columns(i) = n
Next

rows1 = Split(rel1, prelNewline())
rows2 = Split(rel2, prelNewline())

r1 = UBound(rows1)
r2 = UBound(rows2)

'reorganize columnns in second relation so that they have the same order
For i = 0 To r2
    fields1 = Split(rows2(i), "::")
    ReDim fields2(c2)
    For j = 0 To c2
        fields2(columns(j)) = fields1(j)
    Next j
    rows2(i) = Join(fields2, "::")
Next i

ReDim rows(r1)
rows(0) = rows1(0)
offset = 1
For i = 1 To r1
    found = False
    ' for each tuple in the first relation we check if it is in the second relation
    For j = 0 To r2
        If rows1(i) = rows2(j) Then
            found = True
            rows(offset) = rows1(i)
            offset = offset + 1
            Exit For
        End If
    Next j
    If Not found Then
        
    End If
Next i

'If offset = 1 Then
'    relIntersect = ""
'    Exit Function
'End If

ReDim Preserve rows(offset - 1)

relIntersect = Join(rows, prelNewline()) 'no duplicate elimination needed


End Function

Function relSelect(rel As String, condition As String) As String

Dim arr()  As Variant
Dim values() As String
Dim rows() As String
Dim cond, header, field, row, key, s As String
Dim r, c, i, j, offset As Long
Dim eval As Variant

prelStamp ("relselect")

If rel = "" Then
    relSelect = ""
    Exit Function
End If

rows = Split(rel, prelNewline())
r = UBound(rows)
header = rows(0)

relSelect = prelCheckHeader(header)
If relSelect <> "" Then Exit Function

condition = prelSubstituteNames(condition, header)

offset = 0
For i = 1 To r
    row = rows(i)
    values = Split(row, "::")
    cond = prelParseExpression(condition, values)
    eval = prelEvaluate(cond)
    If IsError(eval) Then
        relSelect = "#ERROR CONDITION LINE " & Str(i + 1) & " : " & cond
    Exit Function
    End If
    If eval Then
        offset = offset + 1
        rows(offset) = row
    End If
Next i

ReDim Preserve rows(offset)

relSelect = Join(rows, prelNewline()) 'no duplicate elimination needed


End Function

Function relExtend(rel As String, ByVal calculation As String, Optional ByVal name As String, Optional noError As Boolean = False) As String

Dim arr()  As Variant
Dim values() As String
Dim rows() As String
Dim newlist() As String
Dim cond, header, field, l As String
Dim r, c, i, j, offset As Long
Dim result As Variant

prelStamp ("relextend")

On Error GoTo errHandler

If rel = "" Then
    relExtend = ""
    Exit Function
End If

rows = Split(rel, prelNewline())
r = UBound(rows)
header = rows(0)

relExtend = prelCheckHeader(header)
If relExtend <> "" Then Exit Function

values = Split(header, "::")
c = UBound(values)

calculation = prelSubstituteNames(calculation, header)

If name = "" Then name = "c" & Format(c + 1, "0")
rows(0) = rows(0) + "::" + name
'no duplicate columns allowed
newlist = Split(rows(0), "::")
For i = 0 To UBound(newlist)
    For j = 0 To i - 1
        If newlist(i) = newlist(j) Then
            relExtend = "#ERROR DUPLICATE COLUMN " & newlist(j)
            Exit Function
        End If
    Next j
Next i

For i = 1 To r
    values = Split(rows(i), "::")
    cond = prelParseExpression(calculation, values)
    If cond = "=(0)" Then
        result = "0"
    Else
        result = prelEvaluate(cond)
    End If
    If IsError(result) Then
        relExtend = "#ERROR CALCULATION LINE " & Str(i + 1) & " : " & cond
        Exit Function
    End If
    If IsNumeric(result) Then result = Trim(Str(result))
    If result = 0 Then result = Trim(Str(result))
    If result = True Then result = Trim(Str(result))
    rows(i) = rows(i) & "::" & result
Next i

l = Join(rows, prelNewline()) 'no duplicate elimination needed

If Len(l) > 32768 And Not noError Then
        relExtend = "#ERROR LONG RESULT " & Str(Len(l))
        Exit Function
End If
 
relExtend = l


Exit Function

errHandler:
  relExtend = "#Error relExtend " & Err.Number & ": " & Err.Description


End Function


Function relProject(ByVal rel As String, ByVal list As String) As String

' project is also used for group aggregation, as relational algebra does not know duplicates

Dim arr()  As Variant
Dim aggregators() As String
Dim rows() As String
Dim rowkey As String
Dim values() As String
Dim test As Variant
Dim cols() As String
Dim vfields() As String
Dim headerfields() As String
Dim cstring, vstring As String
Dim r, c, c2, i, j, excluded, cval As Long
Dim v1, v2, v3 As Double
Dim s1, s2, s3 As String
Dim found, hasaverage As Boolean
Dim header, newheader As String
Dim newlist() As String
Dim usecollection As Boolean
Dim medianlines As Long

prelStamp ("relproject")

Dim aggregator() As String
Dim dict As New Collection
Dim cc As Long


On Error GoTo errHandler

If list = "" Then
    relProject = ""
    Exit Function
End If

If list = " " Then
    relProject = "#ERROR LIST EMPTY"
    Exit Function
End If

If rel = "" And InStr(list, "COUNT") Then
    relProject = Replace(list, " COUNT", "_count") & prelNewline() & "0"
    Exit Function
End If

If rel = "" And (list Like "SUM #" Or list Like "SUM ##") Then
    relProject = list + prelNewline() & "0"
    Exit Function
End If

If rel = "" Then
    relProject = ""
    Exit Function
End If


If rel = "" Then
    relProject = "#ERROR COLUMN: " & list
    Exit Function
End If

arr = prelArray(rel)
rows = Split(rel, prelNewline())
header = rows(0)

relProject = prelCheckHeader(header)
If relProject <> "" Then Exit Function

headerfields = Split(header, "::")

r = UBound(arr, 1)
c = UBound(arr, 2)
excluded = 0


'check list and build aggregators
cols = Split(list, "::")
c2 = UBound(cols) ' onebased
ReDim aggregator(c2)
ReDim newlist(c2)
' check limit
For i = 0 To c2
    cstring = cols(i)
    aggregator(i) = ""
    If InStr(cstring, " ") Then
        Dim words() As String
        words = Split(cstring, " ")
        Select Case words(1)
            Case "SUM"
                aggregator(i) = "SUM"
            Case "COUNT"
                aggregator(i) = "COUNT"
            Case "MAX"
                aggregator(i) = "MAX"
            Case "MIN"
                aggregator(i) = "MIN"
            Case "AVG"
                aggregator(i) = "AVG"
                hasaverage = True
            Case "MEDIAN"
                aggregator(i) = "MEDIAN"
                hasaverage = True
            Case "STDEV"
                aggregator(i) = "STDEV"
                hasaverage = True
            Case Else
                relProject = "#ERROR AGGREGATOR: " & cstring
                Exit Function
        End Select
        cval = prelNameToColumn(header, words(0))
        If cval < 1 Or cval > c + 1 Then
            relProject = "#ERROR COLUMN: " & cstring
            Exit Function
        End If
        newlist(i) = headerfields(cval - 1) & "_" & LCase(aggregator(i))
        usecollection = True
    Else
        cval = prelNameToColumn(header, cstring)
        If cval < 1 Or cval > c + 1 Then
            relProject = "#ERROR COLUMN: " & cstring
            Exit Function
        End If
        newlist(i) = headerfields(cval - 1)
    End If
    
    cols(i) = cval
Next i

'no duplicate columns allowed
For i = 0 To UBound(newlist)
    For j = 0 To i - 1
        If newlist(i) = newlist(j) Then
            relProject = "#ERROR DUPLICATE COLUMN " & newlist(j)
            Exit Function
        End If
    Next j
Next i

' project and aggregate
' we use a collection to get unique keys
ReDim rows(r)




For i = 1 To r
    ReDim values(c2)
    For j = 0 To c2
        If aggregator(j) = "" Then
            cc = cols(j) - 1
            If cc < 0 Or cc > c Then
                relProject = "#ERROR COLUMN: " & Str(j + 1)
                Exit Function
            End If
            values(j) = arr(i, cc)
        End If
    Next j
    rowkey = Join(values, "::")
    If prelInCollection(dict, rowkey) Then
        ' you cannot change a collection, so we have to remove and add it later.
        values = Split(dict.Item(rowkey), "::")
        dict.Remove (rowkey)
    End If
     For j = 0 To c2
        
        cc = cols(j) - 1
        If cc < 0 Or cc > c Then
            relProject = "#ERROR AGGREGATOR: " & Str(j + 1)
            Exit Function
        End If
        If arr(i, cc) = "" Then
          ' ignore empty values
        Else
            Select Case aggregator(j)
            Case "SUM"
                
                v1 = prelDouble(values(j))
                v2 = prelDouble(arr(i, cc))
                values(j) = Trim(Str(prelDouble(values(j)) + prelDouble(arr(i, cc))))
            Case "COUNT"
                
                values(j) = Trim(Str(prelDouble(values(j)) + 1))
            Case "MAX"
                v1 = prelDouble(values(j))
                v2 = prelDouble(arr(i, cc))
                If values(j) = "" Or v2 > v1 Then values(j) = arr(i, cc)
            Case "MIN"
              v1 = prelDouble(values(j))
              v2 = prelDouble(arr(i, cc))
             If values(j) = "" Or v2 < v1 Then values(j) = arr(i, cc)
            Case "AVG"
                vstring = values(j)
                If vstring = "" Then vstring = "0/0"
                vfields = Split(vstring, "/")
                v1 = prelDouble(vfields(0))
                v2 = prelDouble(vfields(1))
                v1 = v1 + prelDouble(arr(i, cc))
                v2 = v2 + 1
                ' we pack the sum and the count into one value
                vstring = Trim(Str(v1)) & "/" & Trim(Str(v2))
                values(j) = vstring
            Case "MEDIAN"
                'we build a relation we will order later to take median value
                If values(j) = "" Then values(j) = "vmedian"
                values(j) = values(j) + prelNewline() + arr(i, cc)
            Case "STDEV"
                vstring = values(j)
                If vstring = "" Then vstring = "0/0/0"
                vfields = Split(vstring, "/")
                v1 = prelDouble(vfields(0))
                v2 = prelDouble(vfields(1))
                v3 = prelDouble(vfields(2))
                v1 = v1 + prelDouble(arr(i, cc)) * prelDouble(arr(i, cc))
                v2 = v2 + prelDouble(arr(i, cc))
                v3 = v3 + 1
                ' we pack the sum and the count into one value
                vstring = Trim(Str(v1)) & "/" & Trim(Str(v2)) & "/" & Trim(Str(v3))
                values(j) = vstring
            End Select
        End If
       
    Next j
    dict.Add Join(values, "::"), rowkey
Next i

ReDim rows(dict.Count)
rows(0) = Join(newlist, "::")
For i = 1 To dict.Count
    rows(i) = dict.Item(i)
    'If hasaverage Then
        values = Split(rows(i), "::")
        For j = 0 To c2
            Select Case aggregator(j)
            Case "SUM", "COUNT"
                If values(j) = "" Then values(j) = 0
            Case "AVG"
                 'we need to make the division of sum/count
                 vstring = values(j)
                 vfields = Split(vstring, "/")
                 v1 = prelDouble(vfields(0))
                 v2 = prelDouble(vfields(1))
                 ' we never have 0 division here, haven't we
                 If v2 = 0 Then
                    values(j) = ""
                 Else
                     vstring = Trim(Str(v1 / v2))
                    values(j) = vstring
                End If
            Case "MEDIAN"
                s1 = values(j)
                s1 = relOrder(s1, "vmedian 9")
                vfields = Split(s1, prelNewline())
                medianlines = UBound(vfields)
                If medianlines < 0 Then
                    values(j) = ""
                ElseIf Round(medianlines / 2, 0) = medianlines / 2 Then
                    'even
                    v1 = prelDouble(vfields(medianlines / 2))
                    v2 = prelDouble(vfields(medianlines / 2 + 1))
                    values(j) = Str((v1 + v2) / 2)
                Else
                    values(j) = vfields(Round(medianlines / 2 + 0.1))
                    
                End If
           Case "STDEV"
                 'we need to make the division of sum/count
                 vstring = values(j)
                 vfields = Split(vstring, "/")
                 v1 = prelDouble(vfields(0))
                 v2 = prelDouble(vfields(1))
                 v3 = prelDouble(vfields(2))
                 
                 If v3 = 0 Then
                    values(j) = ""
                 Else
                 ' we never have 0 division here, haven't we
                    vstring = Trim(Str(VBA.Sqr(v1 / v3 - (v2 * v2) / (v3 * v3))))
                    values(j) = vstring
                End If
            End Select
        Next j
        rows(i) = Join(values, "::")
    'End If
Next i



'special case no row, we still need count and sum
If UBound(rows) = 0 Then
    ReDim values(c2)
    found = False
    For j = 0 To c2
        If aggregator(j) = "SUM" Or aggregator(j) = "COUNT" Then
            values(j) = 0
            found = True
        End If
    Next j
    If found Then
        ReDim rows(1)
        rows(0) = Join(newlist, "::")
        rows(1) = Join(values, "::")
    End If
    
End If

relProject = Join(rows, prelNewline()) 'no duplicate elimination needed
Exit Function

errHandler:
  Application.StatusBar = "Error relProject" & Err.Number & ": " & Err.Description



End Function

Private Function prelSpecialJoin(rel1 As String, rel2 As String, Optional flag As String = "", Optional noError As Boolean = False, Optional lazy As Boolean = False) As String

'this is a natural join on common fields
'flags: "NATURAL" (default), "LEFT"
'"LEFTSEMI", "LEFTANTISEMI"

On Error GoTo errHandler

prelStamp ("prelspecialjoin")

Dim rows1() As String
Dim rows2() As String
Dim rows() As String
Dim common1() As Long
Dim common2() As Long
Dim other1() As Long
Dim other2() As Long
Dim values1() As String
Dim values2() As String
Dim values() As String
Dim fields1() As String
Dim fields2() As String
Dim fields() As String
Dim keys1() As String
Dim keys2() As String
Dim hexkey As String
Dim rest1() As String
Dim rest2() As String
Dim list() As String
Dim row, first1, first2, empty1, empty2, elem As String
Dim r, r1, r2, c, c1, c2, o1, o2, i, j, k, l, offset, commoncolumns As Long
Dim eval, elev As Variant
Dim found As Boolean

flag = Trim(flag)

If rel1 = "" Or rel2 = "" Then
    Select Case flag
    Case "NATURAL", "LEFTSEMI"
        prelSpecialJoin = ""
    Case "LEFT", "LEFTANTISEMI"
        prelSpecialJoin = rel1
    End Select
    Exit Function
End If



 


rows1 = Split(rel1, prelNewline())
rows2 = Split(rel2, prelNewline())

r1 = UBound(rows1)
r2 = UBound(rows2)

'find common columns
fields1 = Split(rows1(0), "::")
fields2 = Split(rows2(0), "::")
c1 = UBound(fields1)
c2 = UBound(fields2)

commoncolumns = 0
ReDim common1(c1 + c2)
ReDim common2(c1 + c2)
ReDim other1(c1)
ReDim other2(c2)
ReDim fields(c1 + 1 + c2)

For i = 0 To c1
    If Mid(fields1(i), 3, 1) = "." Then
        fields1(i) = Mid(fields1(i), 4)
    End If
    For j = 0 To c2
        If Mid(fields2(j), 3, 1) = "." Then
            fields2(j) = Mid(fields2(j), 4)
        End If
        If fields1(i) = fields2(j) Then
            common1(commoncolumns) = i
            common2(commoncolumns) = j
            fields(commoncolumns) = fields1(i)
            commoncolumns = commoncolumns + 1
        End If
    Next j
Next i
commoncolumns = commoncolumns - 1

If commoncolumns < 0 Then
    prelSpecialJoin = ""
    Exit Function
End If

ReDim Preserve common1(commoncolumns)
ReDim Preserve common2(commoncolumns)
ReDim Preserve fields(commoncolumns)

c = 0
For i = 0 To c1
    found = False
    For j = 0 To commoncolumns
        If fields1(i) = fields(j) Then
            found = True
            Exit For
        End If
    Next j
    If Not found Then
       other1(c) = i
       c = c + 1
    End If
Next i
If c > 0 Then
    ReDim Preserve other1(c - 1)
End If
o1 = c - 1

c = 0
For i = 0 To c2
    found = False
    For j = 0 To commoncolumns
        If fields2(i) = fields(j) Then
            found = True
            Exit For
        End If
    Next j
    If Not found Then
       other2(c) = i
       c = c + 1
    End If
Next i
If c > 0 Then
    ReDim Preserve other2(c - 1)
End If
o2 = c - 1


' now prepare all rows, first part is common, second is other

ReDim values1(commoncolumns)
ReDim values2(UBound(other1))
ReDim keys1(r1)
ReDim keys2(r2)
ReDim rest1(r1)
ReDim rest2(r2)

For i = 0 To r1
    If rows1(i) = "" Then
        ReDim values(0)
    Else
        values = Split(rows1(i), "::")
    End If
    For j = 0 To commoncolumns
        values1(j) = values(common1(j))
        
    Next j
    keys1(i) = Join(values1, "::")
    For j = 0 To o1
        values2(j) = values(other1(j))
    Next j
    If o1 < 0 Then
        rest1(i) = ""
    Else
        rest1(i) = "::" & Join(values2, "::")
    End If
Next i

ReDim values2(UBound(other2))
Dim keyhash As New Collection

For i = 0 To r2
    If rows2(i) = "" Then
        ReDim values(0)
    Else
        values = Split(rows2(i), "::")
    End If
    For j = 0 To commoncolumns
        values1(j) = values(common2(j))
    Next j
    keys2(i) = Join(values1, "::")
    For j = 0 To o2
        values2(j) = values(other2(j))
    Next j
    If o2 < 0 Then
        rest2(i) = ""
    Else
        rest2(i) = "::" & Join(values2, "::")
    End If
    hexkey = prelAsciiToHexString(keys2(i))
    If prelInCollection(keyhash, hexkey) Then
        rest2(i) = keyhash.Item(hexkey) + prelNewline() + rest2(i)
        keyhash.Remove hexkey
    End If
    keyhash.Add rest2(i), hexkey
Next i


'create empty rows for outer
empty1 = ""
empty2 = ""
If rest1(0) <> "" Then
ReDim values1(UBound(other1))
empty1 = "::" & Join(values1, "::")
End If
If rest2(0) <> "" Then
ReDim values2(UBound(other2))
empty2 = "::" & Join(values2, "::")
End If


' now we can cross compare all rows
' header
ReDim rows(r + 1)
Select Case flag
    Case "NATURAL", "LEFT", "RIGHT", "OUTER", ""
        rows(0) = keys1(0) & rest1(0) & rest2(0)
    Case "LEFTSEMI", "LEFTANTISEMI"
        rows(0) = keys1(0) & rest1(0)
    Case "RIGHTSEMI", "RIGHTANTISEMI"
        rows(0) = keys1(0) & rest2(0)
End Select


'rows
offset = 1
For i = 1 To r1
    found = False
    hexkey = prelAsciiToHexString(keys1(i))
    If prelInCollection(keyhash, hexkey) Then
        found = True
        Select Case flag
        Case "NATURAL", "LEFT", ""
            Dim rest As String
            rest = keyhash.Item(hexkey)
            Dim restlist() As String
            restlist = Split(rest, prelNewline())
            r2 = UBound(restlist)
            If r2 = -1 Then
                rows(offset) = keys1(i) & rest1(i)
                offset = offset + 1
                If offset > UBound(rows) Then
                    ReDim Preserve rows(offset * 2)
                End If
            Else
                For j = 0 To r2
                    rows(offset) = keys1(i) & rest1(i) & restlist(j)
                    offset = offset + 1
                    If offset > UBound(rows) Then
                        ReDim Preserve rows(offset * 2)
                     End If
                Next j
            End If
        Case "LEFTSEMI"
            rows(offset) = keys1(i) & rest1(i)
            offset = offset + 1
            If offset > UBound(rows) Then
                    ReDim Preserve rows(offset * 2)
            End If
        End Select
    End If
    ' add outer rows on the left
    If Not found Then
        Select Case flag
            Case "LEFT"
                rows(offset) = keys1(i) & rest1(i) & empty2
                offset = offset + 1
            Case "LEFTANTISEMI"
                rows(offset) = keys1(i) & rest1(i)
                offset = offset + 1
        End Select
        If offset > UBound(rows) Then
            ReDim Preserve rows(offset * 2)
        End If
    End If
    If offset > UBound(rows) Then
        ReDim Preserve rows(offset * 2)
    End If
Next i


ReDim Preserve rows(offset - 1)

If lazy Then
    l = Join(rows, prelNewline()) 'no duplicate elimination needed
Else
    l = prelString(rows)
End If
 

If Len(l) > 32768 And Not noError Then
        prelSpecialJoin = "#ERROR LONG RESULT " & Str(Len(l))
        Exit Function
End If
 
prelSpecialJoin = l


Exit Function

errHandler:
  prelSpecialJoin = "Error prelSpecialJoin " & Err.Number & ": " & Err.Description



End Function


Function relJoin(rel1 As String, rel2 As String, condition As String, Optional noError As Boolean = False, Optional lazy As Boolean = False) As String

'this is a theta join
'for cross just set conditon "true"
'if you set equal condition on all columns, you get an intersection

Dim rows1() As String
Dim rows2() As String
Dim rows() As String
Dim values1() As String
Dim values2() As String
Dim values() As String
Dim fields() As String
Dim header, cond, row, first1, first2 As String
Dim r1, r2, c1, c2, i, j, k, l, offset As Long
Dim eval As Variant
Dim leftflag, rightflag As Boolean
Dim leftfound As Boolean
Dim rightfound() As Boolean
Dim emptyleft As String
Dim emptyright As String



prelStamp ("reljoin")

Select Case condition
Case "CROSS"
    relJoin = prelSpecialJoin(rel1, rel2, condition, noError)
    Exit Function
Case "RIGHT"
    relJoin = prelSpecialJoin(rel2, rel1, "LEFT", noError, lazy)
    Exit Function
Case "RIGHTSEMI"
    relJoin = prelSpecialJoin(rel2, rel1, "LEFTSEMI", noError, lazy)
    Exit Function
Case "RIGHTANTISEMI"
    relJoin = prelSpecialJoin(rel2, rel1, "LEFTANTISEMI", noError, lazy)
    Exit Function
Case "OUTER"
    Dim left1 As String
    Dim left2 As String
    left1 = prelSpecialJoin(rel1, rel2, "LEFT", noError, True)
    left2 = prelSpecialJoin(rel2, rel1, "LEFT", noError, True)
    relJoin = relUnion(left1, left2, noError, lazy)
    Exit Function
Case "NATURAL", "LEFT", "LEFTSEMI", "LEFTANTISEMI"
    relJoin = prelSpecialJoin(rel1, rel2, condition, noError, lazy)
    Exit Function
End Select


If rel1 = "" Or rel2 = "" Then
    relJoin = ""
    Exit Function
End If

rows1 = Split(rel1, prelNewline())
rows2 = Split(rel2, prelNewline())

first1 = rows1(0)
first2 = rows2(0)

r1 = UBound(rows1)
r2 = UBound(rows2)


values1 = Split(first1, "::")
values2 = Split(first2, "::")
c1 = UBound(values1)
c2 = UBound(values2)

ReDim fields(c1)
emptyleft = Join(fields, "::")
ReDim fields(c2)
emptyright = Join(fields, "::")


For i = 0 To UBound(values1)
    For j = 0 To UBound(values2)
        If values1(i) = values2(j) Then
            values1(i) = values1(i) & "_1"
            values2(j) = values2(j) & "_2"
        End If
    Next j
Next i
rows1(0) = Join(values1, "::")
rows2(0) = Join(values2, "::")
header = rows1(0) & "::" & rows2(0)



If Left(condition, 4) = "LEFT" Then
    leftflag = True
    condition = Mid(condition, 6)
ElseIf Left(condition, 5) = "RIGHT" Then
    rightflag = True
    condition = Mid(condition, 7)
ElseIf Left(condition, 5) = "OUTER" Then
    leftflag = True
    rightflag = True
    condition = Mid(condition, 7)
End If


condition = prelSubstituteNames(condition, header)

offset = 1
ReDim rows(r1 + 1) 'we will make it bigger later when needed
rows(0) = header
ReDim rightfound(r2)
For i = 1 To r1
    leftfound = False
    For j = 1 To r2
        row = rows1(i) & "::" & rows2(j)
        values = Split(row, "::")
        
        cond = prelParseExpression(condition, values)
        eval = prelEvaluate(cond)
        If IsError(eval) Then
            relJoin = "#ERROR CONDITION LINE " & Str(i + 1) & "/" & Str(j + 1) & " : " & cond
            Exit Function
        End If
        If Left(eval, 6) = "#ERROR" Then
            relJoin = "#ERROR CONDITION LINE " & Str(i + 1) & "/" & Str(j + 1) & " : " & eval
            Exit Function
        End If
        If eval Then
            If offset > UBound(rows) Then
                'we grow the array only as much as needed
                ReDim Preserve rows(2 * offset)
            End If
            rows(offset) = row
            offset = offset + 1
            leftfound = True
            rightfound(j) = True
        End If
     Next j
     If Not leftfound And leftflag Then
        If offset > UBound(rows) Then
            'we grow the array only as much as needed
            ReDim Preserve rows(2 * offset)
        End If
        rows(offset) = rows1(i) + "::" + emptyright
        offset = offset + 1
     End If
Next i
For j = 1 To r2
    If offset > UBound(rows) Then
        'we grow the array only as much as needed
        ReDim Preserve rows(2 * offset)
    End If
    If Not rightfound(j) And rightflag Then
        rows(offset) = emptyleft & "::" & rows2(j)
        offset = offset + 1
    End If
Next j


ReDim Preserve rows(offset - 1)

If lazy Then
    l = Join(rows, prelNewline()) 'no duplicate elimination needed
Else
    l = prelString(rows)
End If

If Len(l) > 32768 And Not noError Then
        relJoin = "#ERROR LONG RESULT " & Str(Len(l))
        Exit Function
End If
 
relJoin = l



End Function


Function relOrder(ByVal rel As String, list As String) As String
Dim arr()  As String
Dim values() As String
Dim rows() As String
Dim orderlist() As String
Dim order As String
Dim cols() As Long
Dim modes() As String
Dim fields() As String
Dim cond As String
Dim field As String
Dim handrow As String
Dim other As String
Dim bigger As Variant
Dim r, c, c2, i, j, offset As Long

prelStamp ("relorder")

If rel = "" Then
 relOrder = ""
 Exit Function
End If

arr = Split(rel, prelNewline())


relOrder = prelCheckHeader(arr(0))
If relOrder <> "" Then Exit Function


r = UBound(arr, 1)

orderlist = Split(list, "::")
c2 = UBound(orderlist)
ReDim cols(c2)
ReDim modes(c2)

' order is a list separated with ::
' each item is column and modifier
' modifier influences order
' A alphabetic (default if omitted)
' Z alphabetic reverse
' 1 numeric bottom top
' 9 numeric top bottom

For j = 0 To c2
    order = orderlist(j)
    fields = Split(order, " ")
    cols(j) = prelNameToColumn(arr(0), fields(0)) - 1
    If cols(j) < 0 Then
        relOrder = "ERROR COLUMN " & fields(0)
        Exit Function
    End If
    If UBound(fields) > 0 Then
        modes(j) = fields(1)
    Else
      modes(j) = "A"
    End If
Next

ReDim values(c)

'insertion sort
'all rows before the ith row are ordered (row 0 is sorted at beginning)
'we take the ith row in to hand, compare to all precedent rows, move them forward
'until the row in the hand is bigger, then we insert
'the comparison in costum function prelBigger
For i = 2 To r
    handrow = arr(i)
    j = i
    Do
        j = j - 1
        other = arr(j)
        bigger = prelBigger(other, handrow, cols, modes)
        
        If bigger <> True And bigger <> False Then
            relOrder = "#ERROR " & bigger
            Exit Function
        End If
        If bigger Then
            arr(j + 1) = other
        Else
              j = j + 1
              Exit Do
        End If
    Loop Until j = 1 'top row is header and must stay on top
    arr(j) = handrow
Next i

relOrder = Join(arr, prelNewline()) 'no duplicate elimination needed

End Function


Private Function prelBigger(row1 As String, row2 As String, cols() As Long, modes() As String) As Boolean
    Dim v1s() As String
    Dim v2s() As String
    Dim v1, v2 As String
    Dim i, c2 As Long
    Dim test As Long
    
    'trivial
    If row1 = "" Then
      prelBigger = False 'smaller or equal, should not happe
      Exit Function
    End If
        
    If row2 = "" Then
      prelBigger = True 'bigger, as row1 cannot be ""
      Exit Function
    End If
    
    v1s = Split(row1, "::")
    v2s = Split(row2, "::")
    
    c2 = UBound(cols)
    
    For i = 0 To c2
        If cols(i) < 0 Or cols(i) > UBound(v1s) Then
        'error
        prelBigger = "COLUMN " & Str(cols(i) + 1)
           Exit Function
        End If
        v1 = v1s(cols(i))
        v2 = v2s(cols(i))
        
        test = 0
        Select Case modes(i)
        Case "A"
            test = StrComp(v1, v2)
        Case "Z"
             test = StrComp(v2, v1)
        Case "1"
          'sgn is -1 0 1
           test = Sgn(prelDouble(v1) - prelDouble(v2))
        Case "9"
           test = Sgn(prelDouble(v2) - prelDouble(v1))
        Case "ASC"
           test = prelCompareSql(v1, v2)
        Case "DESC"
           test = prelCompareSql(v2, v1)
        Case Else
           'error
           prelBigger = "MODE " & modes(i)
           Exit Function
       End Select
       
       Select Case test
        Case 1
            prelBigger = True
            Exit Function
        Case -1
            prelBigger = False
            Exit Function
        Case 0
            'both are equal at this level, we need to go to the next on the list
            prelBigger = False
        End Select
  Next i
    'both are equal (possible, if not all fields are in order or if numeric value equal)
    prelBigger = False
End Function



Private Function prelArray(s As String) As Variant
Dim rows() As String
Dim fields() As String
Dim cells() As Variant
Dim r, c, i, j As Long
Dim last As String
last = prelStamp("prelarray")
' converts a relation to a 0-based 2-dimensional array

rows = Split(s, prelNewline())
fields = Split(rows(0), "::")

r = UBound(rows)
c = UBound(fields)

ReDim cells(r, c)

For i = 0 To r
    fields = Split(rows(i), "::")
    ReDim Preserve fields(c) 'format all rows to first one
For j = 0 To c
    cells(i, j) = fields(j)
Next j
Next i

prelArray = cells
prelStamp (last)

End Function

Private Function prelString(arr() As String) As String
Dim tuples() As String
Dim tuple As String
Dim fields() As String
Dim duplicates As Long
Dim r As Long
Dim i As Long
Dim j As Long
Dim found As Boolean
 
' converts an array to a relation and eliminates duplicates and empty lines
 
r = UBound(arr, 1)
 
Dim hash As New Collection
Dim key As String

 
For i = 0 To r
    If Replace(arr(i), "::", "") <> "" Then
        key = prelAsciiToHexString(arr(i))
        If prelInCollection(hash, key) Then
            ' do nothing
        Else
            hash.Add arr(i), key
        End If
    End If
Next i
 
Dim c As Long
c = hash.Count
ReDim tuples(c - 1)
 
For i = 1 To c
     tuples(i - 1) = hash.Item(i)
Next
 
 
prelString = Join(tuples, prelNewline())
 
 
 
End Function
 
 
 
Private Function prelAsciiToHexString(ByVal asciiText As String, _
                                 Optional ByVal hexPrefix As String = "") As String
    prelAsciiToHexString = asciiText  'default failure return value
    If Not (asciiText = vbNullString) Then
        Dim asciiChars() As Byte
        asciiChars = StrConv(asciiText, vbFromUnicode)
        ReDim hexChars(LBound(asciiChars) To UBound(asciiChars)) As String
        Dim char As Long
        For char = LBound(asciiChars) To UBound(asciiChars)
            hexChars(char) = Right$("00" & Hex$(asciiChars(char)), 2)
        Next char
        prelAsciiToHexString = hexPrefix & Join(hexChars, "")
    End If
End Function


Function relCell(rel As String, r As Long, c As Long, Optional Numeric As Boolean = False, Optional noError As Boolean = False) As Variant

Dim tuples() As String
Dim tuple As Variant
Dim fields() As String

prelStamp ("relcell")

If rel = "" Then
    relCell = ""
    Exit Function
End If

'user 1-based
c = c - 1

tuples = Split(rel, prelNewline())

If r < 0 Or r > UBound(tuples) Then
    If noError Then
        relCell = ""
    Else
        relCell = "#ERROR BOUNDS ROW: " & Str(r)
    End If
    Exit Function
End If

tuple = tuples(r)

fields = Split(tuple, "::")

If c < 0 Or c > UBound(fields) Then
    If noError Then
        relCell = ""
    Else
        relCell = "#ERROR BOUNDS COLUMN: " & Str(c + 1)
    End If
    Exit Function
End If

relCell = fields(c)
If Numeric And IsNumeric(relCell) Then
   relCell = CDbl(relCell)
End If


End Function







Static Function prelNewline() As String
    
    Dim platform As Long
    
    Select Case platform
    Case 1
        prelNewline = " " & vbLf
    Case 2
        prelNewline = " " & vbCrLf
    Case Else
        'only evaluate once may cost CPU time
        If Application.OperatingSystem Like "*Mac*" Then
            platform = 1
        Else
            platform = 2
        End If
        prelNewline = prelNewline()
    End Select


End Function


Private Function prelInCollection(col As Collection, key As String) As Boolean

' it is not possible to get a list of keys from a collection.
' so we just try to get the value and catch the error

On Error GoTo incol

col.Item key
prelInCollection = True
Exit Function

incol:
  prelInCollection = False

End Function

Function relLike(ByVal s As String, ByVal pattern As String) As Boolean
  
    'expose like to excel
    relLike = s Like pattern
    
End Function



Public Function relCellArray(rel As String, Optional noHeader As Boolean = False)
Dim c, r, i, j As Long
Dim r1(), r2() As Variant
Dim start As Long

'relCell for a complete range
'empty if out of range
'control shift enter to use this function

If noHeader Then
    start = 1
Else
    start = 0
End If


With Application.Caller
        r = .rows.Count
        c = .columns.Count
End With
ReDim r2(r - start, c)

If rel <> "" Then
    r1 = prelArray(rel)
Else
    ReDim r1(0, 0)
    r1(0, 0) = ""
End If



For i = start To r
    For j = 0 To c
         If i <= UBound(r1, 1) And j <= UBound(r1, 2) Then
            r2(i - start, j) = r1(i, j)
         Else
            r2(i - start, j) = ""
         End If
    Next j
Next i
relCellArray = r2


End Function

Public Function relFilter(ParamArray list()) As Variant
Dim vlist() As Variant
Dim i As Long
Dim c As Long
Dim elem As Variant
Dim rn As Range
Dim s As String
c = 0
For Each elem In list
    c = c + 1
Next elem
If c < 1 Then
    relFilter = ""
    Exit Function
End If
ReDim vlist(c - 1)
i = 0
For Each elem In list
    If TypeName(elem) = "Range" Then
        Set rn = elem
        If rn.Count > 1 Then
            vlist(i) = prelRange(rn, True, True, False)
        Else
            s = rn.Value2 'cast variant as string
            vlist(i) = s
        End If
    ElseIf IsError(elem) Then
        relFilter = "#ERROR invalid parameter"
        Exit Function
        
    Else
        vlist(i) = elem
    End If
    i = i + 1
Next elem

relFilter = prelFilter(vlist)

End Function



Private Function prelFilter(ByRef list() As Variant)
    Dim elem As Variant
    Dim test, body As String
    Dim fields() As String
    Dim v1, v2 As Long
    Dim stack() As String
    Dim stackpointer As Long
    Dim rn As Range
    Dim arr() As Variant
    Dim done As Boolean
    Dim s As String
    Dim extendname, extendbody As String
    
    prelStamp ("relfilter")
    
    Application.ScreenUpdating = False
    
    On Error GoTo errHandler
    
    ReDim stack(0)
    stack(0) = ""
    stackpointer = 0
    
    For Each elem In list
        'test for range bigger than one dimension
        done = False
        prelStamp ("relfilter")
        If VarType(elem) = vbError Then
            done = True
        End If
        If done Then
            ' do nothing
        ElseIf InStr(elem, prelNewline()) Then
            stackpointer = stackpointer + 1
            ReDim Preserve stack(stackpointer)
            stack(stackpointer) = elem
        ElseIf elem = "" Then
            stackpointer = stackpointer + 1
            ReDim Preserve stack(stackpointer)
            stack(stackpointer) = elem
        ElseIf Left(elem, 6) = "#CACHE" Then
            stackpointer = stackpointer + 1
            ReDim Preserve stack(stackpointer)
            stack(stackpointer) = prelCacheGet(Mid(elem, 8))
            stackpointer = stackpointer
        ElseIf Len(elem) > 1 And Mid(elem, 2, 1) <> " " Then 'empty relation with header only
            stackpointer = stackpointer + 1
            ReDim Preserve stack(stackpointer)
            stack(stackpointer) = elem
        Else
        
            test = Left(elem, 1)
            body = Mid(elem, 3) 'space second character ignored
            Select Case test
                Case "S"
                    stack(stackpointer) = relSelect(stack(stackpointer), body)
                Case "E"
                    fields = Split(body, " ")
                    extendname = fields(0)
                    extendbody = Mid(body, Len(extendname) + 2)
                    stack(stackpointer) = relExtend(stack(stackpointer), extendbody, extendname, True)
                Case "P"
                    'as we are lazy, we need to remove duplicates first
                    fields = Split(stack(stackpointer), prelNewline())
                    stack(stackpointer) = relProject(prelString(fields), body)
                Case "U"
                   If stackpointer < 1 Then
                        prelFilter = "#EMPTY STACK UNION"
                        Exit Function
                   End If
                   stackpointer = stackpointer - 1
                   stack(stackpointer) = relUnion(stack(stackpointer), stack(stackpointer + 1), True, True)
                 Case "D"
                   If stackpointer < 1 Then
                        prelFilter = "#EMPTY STACK DIFFERENCE"
                        Exit Function
                   End If
                   stackpointer = stackpointer - 1
                   stack(stackpointer) = relDifference(stack(stackpointer), stack(stackpointer + 1))
                 Case "I"
                   If stackpointer < 1 Then
                        prelFilter = "#EMPTY STACK INTERSECT"
                        Exit Function
                   End If
                   stackpointer = stackpointer - 1
                   stack(stackpointer) = relIntersect(stack(stackpointer), stack(stackpointer + 1))
                Case "J"
                   If stackpointer < 1 Then
                        prelFilter = "#EMPTY STACK JOIN"
                        Exit Function
                   End If
                   stackpointer = stackpointer - 1
                   stack(stackpointer) = relJoin(stack(stackpointer), stack(stackpointer + 1), body, True, True)
                 Case "O"
                    stack(stackpointer) = relOrder(stack(stackpointer), body)
                Case "L"
                    'we have two parameters, separated by space
                    fields = Split(body, " ")
                    
                    If UBound(fields) < 1 Then
                        prelFilter = "#MISSING ARGUMENT LIMIT"
                        Exit Function
                    End If
                    v1 = Val(fields(0))
                    v2 = Val(fields(1))
                    stack(stackpointer) = relLimit(stack(stackpointer), v1, v2)
                Case "R"
                    stack(stackpointer) = relRename(stack(stackpointer), body)
                Case "Q"
                    stack(stackpointer) = relRotate(stack(stackpointer))
                Case "C" ' single cell
                    s = stack(stackpointer)
                    fields = Split(s, prelNewline())
                    If UBound(fields) > 0 Then
                        prelFilter = fields(1)
                        If InStr(fields(1), "::") Then
                            fields = Split(fields(1), "::")
                            prelFilter = fields(0)
                        End If
                        If prelDouble(prelFilter) > 0 Then prelFilter = prelDouble(prelFilter)
                        If prelDouble(prelFilter) < 0 Then prelFilter = prelDouble(prelFilter)
                        If prelFilter = 0 Then prelFilter = 0
                        Exit Function
                    Else
                        prelFilter = ""
                        Exit Function
                    End If
                Case "K" ' single cell forced text
                    s = stack(stackpointer)
                    fields = Split(s, prelNewline())
                    If UBound(fields) > 0 Then
                        prelFilter = fields(1)
                        If InStr(fields(1), "::") Then
                            fields = Split(fields(1), "::")
                            prelFilter = fields(0)
                        End If
                        Exit Function
                    Else
                        prelFilter = ""
                        Exit Function
                    End If
                Case "Z" ' single cell forced number
                    s = stack(stackpointer)
                    fields = Split(s, prelNewline())
                    If UBound(fields) > 0 Then
                        prelFilter = fields(1)
                        If InStr(fields(1), "::") Then
                            fields = Split(fields(1), "::")
                            prelFilter = fields(0)
                        End If
                        prelFilter = prelDouble(prelFilter)
                        Exit Function
                    Else
                        prelFilter = 0
                        Exit Function
                    End If
                Case "!" 'cut
                    prelFilter = stack(stackpointer)
                    Exit Function
                Case "#"
                   
                   ' ignore
             Case Else
                      prelFilter = "#INVALID OPERATOR " & test
                      Exit Function
             End Select
        End If

        If Left(stack(stackpointer), 6) = "#ERROR" Then
            prelFilter = stack(stackpointer)
            Exit Function
        End If
        
    Next elem
    
    If Len(stack(stackpointer)) > 32768 Then
        prelFilter = "#CACHE " + prelCacheSet(stack(stackpointer))
        'relFilter = "#ERROR LONG RESULT " & Str(Len(stack(stackpointer)))
        Exit Function
    End If

    
    prelFilter = stack(stackpointer)
    
    
    ReDim stack(0)
    
    'MsgBox relFilter
    Application.ScreenUpdating = True
    
    
    Exit Function
    
errHandler:
    prelFilter = "Error relFilter " & Err.Number & ": " & Err.Description


    
 
End Function



Public Function relLimit(rel As String, ByVal start As Long, ByVal n As Long) As String
    Dim rows() As String
    Dim i As Long
    Dim result() As String
    
    prelStamp ("rellimit")
    
    If rel = "" Then
        relLimit = "#ERROR EMPTY"
        Exit Function
    End If
    
    rows = Split(rel, prelNewline())
    
    relLimit = prelCheckHeader(rows(0))
    If relLimit <> "" Then Exit Function


        
    If start < 1 Then
       relLimit = "#ERROR OUT OF BOUNDS START"
        Exit Function
    End If
    
    If n < -1 Then
       relLimit = "#ERROR OUT OF BOUNDS N"
        Exit Function
    End If

    'If n = 0 Then
    '    relLimit = ""
    '    Exit Function
    'End If
        
    ReDim result(n)
    
    result(0) = rows(0)
    
    For i = 1 To n
        If start + i - 1 <= UBound(rows) Then
            result(i) = rows(start + i - 1)
        Else
            ReDim Preserve result(i - 1)
            Exit For
        End If
    Next i
    
    
    
    relLimit = Join(result, prelNewline())

End Function

Public Function relRotate(rel As String) As String
   
   Dim arr1() As Variant
   Dim tuple() As String
   Dim arr2() As String
   Dim r As Long
   Dim c As Long
   Dim i As Long
   Dim j As Long
   Dim start As Long
   Dim start2 As Long
    
   ' columns to rows, row to columns
   ' first column will be column names
   ' first row will be first column
        
    If rel = "" Then
        relRotate = ""
        Exit Function
    End If
   
   arr1 = prelArray(rel)
    
    r = UBound(arr1, 1)
    c = UBound(arr1, 2)
    
    ReDim tuple(r)
    ReDim arr2(c)
    
    For i = 0 To c
        For j = 0 To r
            tuple(j) = arr1(j, i)
        Next j
        arr2(i) = Join(tuple, "::")
    
    Next i
    
    relRotate = prelString(arr2)

End Function


Private Function prelDouble(ByVal v As Variant) As Double
   
   'accepts both , and . comma for fractions
     
    If IsNumeric(v) Then
        prelDouble = CDbl(v)
        Exit Function
    End If
    
    v = Replace(v, ",", ".")
    
    If IsNumeric(v) Then
        prelDouble = CDbl(v)
        Exit Function
    End If
    
    v = Replace(v, ".", ",")
    
    If IsNumeric(v) Then
        prelDouble = CDbl(v)
        Exit Function
    End If
    
    prelDouble = 0
    
    
    
End Function


Private Function prelParseExpression(condition As String, values() As String)
Dim cond, field As String
Dim j, c As Long
Dim last As String
Dim test As String

test = Left(condition, 1)
If Left(condition, 1) = "?" Then
    prelParseExpression = prelRPN(condition, values())
    Exit Function
End If
last = prelStamp("prelparsexpression")

cond = condition

c = UBound(values)

' going top down to avoid ambiguities $1 $10
For j = c To 0 Step -1
    field = Format(j + 1, "$00")
    If InStr(cond, field) Then
        cond = Replace(cond, field, """" & values(j) & """")
    End If
    field = Format(j + 1, "%00")
    If InStr(cond, field) Then
        cond = Replace(cond, field, Trim(Str(prelDouble(values(j)))))
    End If
Next j

'put expression in container to have always legal expressions
'note that expression must have english and not local syntax (, instead of ;)

cond = "=(" & cond & ")"

prelParseExpression = cond
prelStamp (last)
End Function

Function prelNameToColumn(ByVal header As String, ByVal name As String) As Long
Dim fields() As String
Dim shortname As String
Dim c, i As Long

    If Val(name) > 0 Then
        prelNameToColumn = Val(name)
        Exit Function
    End If
    fields = Split(header, "::")
    c = UBound(fields)
    For i = 0 To c
        If Trim(LCase(fields(i))) = Trim(LCase(name)) Then
            prelNameToColumn = i + 1
            Exit Function
        End If
    Next i
    
    'if not found search again without prefix
    For i = 0 To c
        If Mid(fields(i), 3, 1) = "." Then
            shortname = Mid(fields(i), 4)
            If Trim(LCase(shortname)) = Trim(LCase(name)) Then
                prelNameToColumn = i + 1
                Exit Function
            End If
        End If
    Next i
   
    
    prelNameToColumn = 0

End Function

Private Function prelSubstituteNames(ByVal expression As String, ByVal header As String) As String
Dim headerlist() As String
Dim i, c, n As Long
Dim field, afield, nfield As String

headerlist = Split(header, "::") 'to do this list should be sorted by length to sort out ambiguities
Dim last As String

last = prelStamp("prelsubstitutenames")

c = UBound(headerlist)

For i = 0 To c
    field = headerlist(i)
    n = prelNameToColumn(header, field)
    afield = "$" & field
    nfield = Format(n, "$00")
    If InStr(expression, afield) Then
        expression = Replace(expression, afield, nfield)
    End If
    afield = "%" & field
    nfield = Format(n, "%00")
    If InStr(expression, afield) Then
        expression = Replace(expression, afield, nfield)
    End If
    afield = "@" & field
    nfield = "@" & Format(n, "00")
    If InStr(expression, afield) Then
        expression = Replace(expression, afield, nfield)
    ElseIf Mid(field, 3, 1) = "." Then
        afield = "@" & Mid(field, 4)
        If InStr(expression, afield) Then
            expression = Replace(expression, afield, nfield)
        End If
    End If
   

Next i

prelSubstituteNames = Trim(expression)
prelStamp (last)

End Function

Function relRename(ByVal rel As String, ByVal list As String) As String

Dim arr()  As Variant
Dim values() As String
Dim rows() As String
Dim newlist() As String
Dim fields() As String
Dim words() As String
Dim cond, header, field, prefix, shortname As String
Dim r, c, i, j, offset, n As Long
Dim result As Variant
Dim dict As New Collection

prelStamp ("relrename")

If rel = "" Then
    relRename = ""
    Exit Function
End If

If list = "" Then
    relRename = rel
    Exit Function
End If



rows = Split(rel, prelNewline())

header = rows(0)
newlist = Split(header, "::")

If Left(list, 1) = "*" Then 'prefix
    prefix = Mid(list, 3) + "."
    For i = 0 To UBound(newlist)
        newlist(i) = prefix + newlist(i)
    Next i
ElseIf Left(list, 1) = "-" Then 'clean
    For i = 0 To UBound(newlist)
        If Mid(newlist(i), 3, 1) = "." Then
            shortname = Mid(newlist(i), 4)
            If prelInCollection(dict, shortname) Then
                dict.Add dict(shortname) + 1, shortname
            Else
               dict.Add 1, shortname
            End If
        End If
    Next i
    For i = 0 To UBound(newlist)
        shortname = Mid(newlist(i), 4)
        If Mid(newlist(i), 3, 1) = "." Then
            If prelInCollection(dict, shortname) Then
                If dict(shortname) = 1 Then
                    newlist(i) = shortname
                End If
            Else
               newlist(i) = shortname
            End If
        End If
    Next i
    
Else

    fields = Split(list, "::")
    c = UBound(fields)
    
    For i = 0 To c
        If InStr(fields(i), " ") Then
            words = Split(fields(i), " ")
            n = prelNameToColumn(header, words(0))
            If n > 0 Then
                If Val(words(1)) > 0 Then
                     relRename = "#ERROR NUMERIC NAME " & words(1)
                     Exit Function
                End If
                
                newlist(n - 1) = words(1)
            End If
        End If
    Next i
    'no duplicate columns allowed
    For i = 0 To UBound(newlist)
        For j = 0 To i - 1
            If newlist(i) = newlist(j) Then
                relRename = "#ERROR DUPLICATE COLUMN " & newlist(j)
                Exit Function
            End If
        Next j
    Next i

End If

rows(0) = Join(newlist, "::")

relRename = Join(rows, prelNewline())


End Function


Function relAssert(ByVal rel As String, ByVal constraint As String, ByVal expression As String)

' possible asserts
' ALL expression
' EXISTS expression
' UNIQUE expression
' COLUMNS expression
' returns true or error tuple

Dim values() As String
Dim rows() As String
Dim fields() As String
Dim cond, header, field, condition As String
Dim r, c, i, j, offset As Long
Dim eval As Variant
Dim found As Boolean

If rel = "" Then
    relAssert = "#ASSERTION EMPTY RELATION"
    Exit Function
End If

rows = Split(rel, prelNewline())
r = UBound(rows)
header = rows(0)

If constraint = "COLUMNS" Then
    fields = Split(expression, "::")
    values = Split(header, "::")
    For i = 0 To UBound(fields)
        found = False
        For j = 0 To UBound(values)
            If values(j) = fields(i) Then
                found = True
                Exit For
            End If
        Next j
        If Not found Then
            relAssert = "#ASSERTION COLUMNS " & fields(i)
            Exit Function
        End If
    Next i
    relAssert = True
    Exit Function
End If


condition = prelSubstituteNames(expression, header)

offset = 0
For i = 1 To r
    values = Split(rows(i), "::")
    cond = prelParseExpression(condition, values)
    eval = Application.Evaluate(cond)
    If IsError(eval) Then
        relAssert = "#ERROR CONDITION LINE " & Str(i + 1) & " : " & cond
        Exit Function
    End If
    
    Select Case constraint
    Case "ALL"
        If Not eval Then
            relAssert = "#ASSERTION ALL " & Str(i + 1) & " : " & cond
            Exit Function
        End If
    Case "EXISTS"
        If eval Then
            relAssert = True
            Exit Function
        End If
    Case "UNIQUE"
        rows(i) = eval
        For j = 1 To i - 1
            If rows(j) = rows(i) Then
                relAssert = "#ASSERTION UNIQUE " & Str(i + 1) & " : " & rows(i) & " = " & Str(j + 1) & " : " & rows(j)
                Exit Function
            End If
        Next j
    Case Else
        relAssert = "#ERROR INVALID CONSTRAINT " & constraint
    End Select
Next i

Select Case constraint
Case "ALL"
     relAssert = True
Case "EXISTS"
     relAssert = "#ASSERTION EXISTS " & expression
Case "UNIQUE"
     relAssert = True
End Select


End Function

Function relFixpoint(ByVal rel As String, fixpoint As String, ByVal start As String, connect As String) As String

Dim rows1() As String
Dim rows2() As String
Dim values1() As String
Dim values2() As String
Dim list() As String
Dim header, tuple1, tuple2, header0 As String
Dim r, col1, col2, offset1, offset2, level As Long
Dim found As Boolean

If rel = "" Then
    relFixpoint = ""
End If

rows1 = Split(rel, prelNewline())
r = UBound(rows1)
header = rows1(0)


col1 = prelNameToColumn(header, fixpoint) - 1
col2 = prelNameToColumn(header, connect) - 1

If col1 < 0 Then
    relFixpoint = "#ERROR INVALID FIXPOINT " & fixpoint
    Exit Function
End If

If col2 < 0 Then
    relFixpoint = "#ERROR INVALID CONNECT " & connect
    Exit Function
End If

start = fixpoint & prelNewline() & start

'level = 1
'start = relExtend(start, , Trim(Str(level)), "level")

relFixpoint = relJoin(rel, start, "NATURAL")

Do
    offset1 = Len(relFixpoint)
    start = relRename(relProject(relFixpoint, connect), connect & " " & fixpoint)
    
    'level = level + 1
    'start = relExtend(start, , Trim(Str(level)), "level")

    relFixpoint = relUnion(relFixpoint, relJoin(rel, start, "NATURAL"))
    relFixpoint = relProject(relFixpoint, header)
    offset2 = Len(relFixpoint)

Loop Until offset2 = offset1


End Function

Function prelCheckHeader(ByVal header As String)

Dim fields() As String
Dim c, i, j As Long

If header = "" Then
    prelCheckHeader = "#ERROR EMPTY HEADER"
    Exit Function
End If

fields = Split(header, "::")
c = UBound(fields)

For i = 0 To c
    If Val(fields(i)) > 0 Or fields(i) = "0" Then
        prelCheckHeader = "#ERROR NUMERIC HEADER"
        Exit Function
    End If
    If Trim(fields(i)) = "" Then
        prelCheckHeader = "#ERROR EMPTY HEADER"
        Exit Function
    End If
    For j = i + 1 To c
        If fields(i) = fields(j) Then
            prelCheckHeader = "#ERROR DUPLICATE HEADER"
            Exit Function
        End If
    Next j
Next i

prelCheckHeader = ""

End Function




Public Function relTokenizeSql(ByVal code As String, Optional ByVal multiline As Long = 0)
    Dim state, buffer, letters, digits, followletters, followdigits, whitespace As String
    Dim operators, multioperators, letteroperators, leftoperators, comma, newline, quote, keywords As String
    Dim functions, aggregators As String
    Dim c, code2, test As String
    Dim offset, elem, start As Long
    Dim tokens() As String
    state = "start"
    buffer = ""
    offset = 0
    elem = 0
    letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
    digits = "0123456789"
    followletters = letters + digits + "._"
    followdigits = digits + "."
    newline = vbCr + vbLf
    whitespace = " " + vbTab + newline
    operators = "+-*/="
    multioperators = " < <= <> > >= " ' keep space after last item
    letteroperators = " AND OR LIKE IN "
    leftoperators = " NOT "
    
    comma = ","
    quote = "'"
    keywords = " SELECT AS FROM NATURAL OUTER LEFT RIGHT JOIN ON WHERE HAVING ORDER BY ASC DESC " ' keep space after last item
    functions = " ABS COS EXP INT LEFT LEN LN LOG LOWER MID MOD POW REPLACE RIGHT ROUND SIN SQRT TAN TRIM UPPER "
    aggregators = " COUNT SUM AVG MEDIAN MIN MAX STDEV " ' keep space after last item
    ReDim tokens(Len(code)) ' force terminating tokens
    code2 = code + " "
    For offset = 1 To Len(code2)
        c = Mid(code2, offset, 1)
        Select Case state
        Case "start"
            start = offset
            If InStr(quote, c) Then
                buffer = """"
                state = "quote"
            ElseIf InStr(letters, c) Then
                buffer = c
                state = "id"
            ElseIf InStr(digits, c) Then
                buffer = c
                state = "number"
            ElseIf InStr(operators, c) Then
                elem = elem + 1
                tokens(elem) = "!" + c
            ElseIf InStr(multioperators, " " + c + " ") Then
                buffer = c
                state = "multioperators"
            ElseIf InStr(comma, c) Then
                elem = elem + 1
                tokens(elem) = "," + c
            ElseIf c = "(" Then
                elem = elem + 1
                tokens(elem) = "(" + c
            ElseIf InStr(")", c) Then
                elem = elem + 1
                tokens(elem) = ")" + c
            ElseIf InStr(whitespace, c) Then
                ' pass
            Else
                relTokenizeSql = "#ERROR invalid token at offset " + Str(offset) + " " + Left(code, offset)
                Exit Function
            End If
        Case "quote"
            If InStr(quote, c) Then
                state = "quotesuspend"
            ElseIf c = """" Then
                buffer = buffer + c + c
            Else
                buffer = buffer + c
            End If
        Case "quotesuspend"
            If InStr(quote, c) Then
                buffer = buffer + "'"
                state = "quote"
            Else
                buffer = buffer + """"
                elem = elem + 1
                tokens(elem) = "$" + buffer
                offset = offset - 1
                state = "start"
            End If
         Case "id"
            If InStr(followletters, c) Then
                buffer = buffer + c
            Else
                elem = elem + 1
                If InStr(keywords, " " + buffer + " ") Then
                    tokens(elem) = buffer
                ElseIf InStr(functions, " " + buffer + " ") Then
                    tokens(elem) = "*" + buffer
                ElseIf InStr(aggregators, " " + buffer + " ") Then
                    tokens(elem) = "+" + buffer
                ElseIf InStr(letteroperators, " " + buffer + " ") Then
                    tokens(elem) = "!" + buffer
                ElseIf InStr(leftoperators, " " + buffer + " ") Then
                    tokens(elem) = "<" + buffer
                Else
                    tokens(elem) = "@" + buffer
                End If
                
                offset = offset - 1
                state = "start"
            End If
        Case "number"
            If InStr(followdigits, c) Then
                buffer = buffer + c
            Else
                elem = elem + 1
                tokens(elem) = "#" + buffer
                offset = offset - 1
                state = "start"
            End If
        Case "multioperators"
            test = buffer + c
            If InStr(multioperators, " " + test + " ") Then ' 2 letters
                elem = elem + 1
                tokens(elem) = "!" + test
                state = "start"
            ElseIf InStr(multioperators, " " + buffer + " ") Then ' 1 letter
                elem = elem + 1
                tokens(elem) = "!" + test
                offset = offset - 1
                state = "start"
            End If
        Case Else
            relTokenizeSql = "#ERROR invalid state " + state + " at offset " + Str(offset) + " " + Left(code, offset)
            Exit Function
            
            
        
        End Select
    Next offset
        
    ReDim Preserve tokens(elem)
     
     
    relTokenizeSql = ""
    
    If multiline = 1 Then
    For offset = 1 To UBound(tokens)
        relTokenizeSql = relTokenizeSql + tokens(offset) + vbCrLf
    Next offset
    Else
    For offset = 1 To UBound(tokens)
        relTokenizeSql = relTokenizeSql + tokens(offset) + " "
    Next offset
    End If
    
    

                
                
            
       
        
    
    
End Function

Public Function relSql(ByVal code As String, ParamArray list())
Dim elem As Variant
Dim tvalues() As String
Dim i As Long
Dim k As Long
Dim n As Long
Dim rn As Range
Dim lines() As String
Dim vlines() As Variant

Dim compiled As String

' first element must be code
' following elements must be values (Range or string)

ReDim tvalues(0)
i = 0

For Each elem In list
    i = i + 1
    ReDim Preserve tvalues(i)
    If TypeName(elem) = "Range" Then
            Set rn = elem
            If rn.Count > 1 Then
                tvalues(i) = prelRange(rn, True, True, False)
            Else
                tvalues(i) = rn.Value2 'what is value2?
            End If
        Else
            tvalues(i) = elem
        End If
Next elem

compiled = relParseSql(code)

lines = Split(compiled, vbCrLf)

n = UBound(lines)

For i = 0 To n
    Select Case Left(lines(i), 1)
    Case "_"
        k = Val(Right(Trim(lines(i)), 1))
        If k > 0 And k <= UBound(tvalues) Then
            lines(i) = tvalues(k)
        Else
            relSql = "#ERROR wrong id " + lines(i)
            Exit Function
        End If
    End Select
Next i

ReDim vlines(UBound(lines))
For i = 0 To UBound(vlines)
    vlines(i) = lines(i)
Next i
relSql = prelFilter(vlines)
i = i
End Function





