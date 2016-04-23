Attribute VB_Name = "Module1"
' Relation for Excel
' Version 1.0 20.4.2016
' matti@belle-nuit.com
'This module provides functions to make simple relational algebra
'The relational model is simplified. A relation is defined as a 2d-table, columns have neither name nor type and are adressed by position (1-based)
'However, rows are not ordered (except with relOrder function) and do not have duplicates
'Unlike other Excel solutions, this module is purely functional, not using macros.
'Relations are saved as text in one cell with :: als field and newline as row separator
'Note that in a cell, the text cannot be more than 32k characters.

Function relRange(rn As Range)
Attribute relRange.VB_Description = "Creates a relation from a range. A relation is a table where rows are separated by newline and columns by ::"
Attribute relRange.VB_ProcData.VB_Invoke_Func = " \n14"

' Calculates a relation from a range
' A relation is a table where rows are separated by newline and columns by ::
' We use a simplified model where tuples have no named properties, but by position (1-based)

Dim arr() As Variant
Dim tuples() As String
Dim fields() As String
Dim r, c, i, j As Integer

arr = rn

'range returns an array which is 1-based
r = UBound(arr, 1) - 1
c = UBound(arr, 2) - 1

ReDim tuples(r)
ReDim fields(c)

For i = 0 To r
For j = 0 To c
    fields(j) = arr(i + 1, j + 1)
Next j
    tuples(i) = Join(fields, "::")
Next i

relRange = relString(tuples)

End Function

Function relUnion(rel1 As String, rel2 As String)
Attribute relUnion.VB_Description = "Calculates the union of two relations. Both relations must have the same arity (column count)."
Attribute relUnion.VB_ProcData.VB_Invoke_Func = " \n14"
Dim first1, first2, r As String
Dim fields1() As String
Dim fields2() As String
Dim arr() As String
Dim ub11, ub12, ub2 As Integer
Dim s As String
Dim c1, c2 As Integer

If rel1 = "" Then
    relUnion = rel2
    Exit Function
End If

If rel2 = "" Then
    relUnion = rel1
    Exit Function
End If

' both relations must have the same arity
' we check the arity on the first tuple of each relation

first1 = relRow(rel1, 0)
first2 = relRow(rel2, 0)

fields1 = Split(first1, "::")
fields2 = Split(first2, "::")


c1 = UBound(fields1)
c2 = UBound(fields2)
If c1 <> c2 Then
    relUnion = "#ERROR ARITY : " + Str(c1 + 1) + " <>" + Str(c2 + 1)
    Exit Function
End If

r = rel1 + relNewline() + rel2

'remove duplicates
arr = Split(r, relNewline())
relUnion = relString(arr)


End Function

Function relDifference(rel1 As String, rel2 As String)
Attribute relDifference.VB_Description = "Calculates the difference of the first minus the second relation. Both relations must have the same arity."
Attribute relDifference.VB_ProcData.VB_Invoke_Func = " \n14"

Dim first1, first2, r As String
Dim fields1() As String
Dim fields2() As String
Dim arr() As String
Dim ub11, ub12, ub2 As Integer
Dim s As String
Dim found As Boolean
Dim r1, r2, c1, c2, i, j, k, l, offset As Integer

If rel1 = "" Then
    relDifference = ""
    Exit Function
End If

If rel2 = "" Then
    relDifference = rel1
    Exit Function
End If

' both relations must have the same arity
' we check the arity on the first tuple of each relation

first1 = relRow(rel1, 0)
first2 = relRow(rel2, 0)

fields1 = Split(first1, "::")
fields2 = Split(first2, "::")

c1 = UBound(fields1)
c2 = UBound(fields2)
If c1 <> c2 Then
    relDifference = "#ERROR ARITY : " + Str(c1 + 1) + " <>" + Str(c2 + 1)
    Exit Function
End If

rows1 = Split(rel1, relNewline())
rows2 = Split(rel2, relNewline())

r1 = UBound(rows1)
r2 = UBound(rows2)

ReDim rows(r1)
offset = 0
For i = 0 To r1
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

If offset = 0 Then
    relDifference = ""
    Exit Function
End If

ReDim Preserve rows(offset - 1)

relDifference = Join(rows, relNewline()) 'no duplicate elimination needed


End Function

Function relSelect(condition As String, rel As String)
Attribute relSelect.VB_Description = "Filters a relation based on a condition (any Excel expression). Use $ for string column, % for number and , as separator."
Attribute relSelect.VB_ProcData.VB_Invoke_Func = " \n14"
Dim arr()  As Variant
Dim values() As String
Dim rows() As String
Dim cond As String
Dim field As String
Dim r, c, i, j, offset As Integer
Dim eval As Variant

If rel = "" Then
    relSelect = ""
    Exit Function
End If

arr = relArray(rel)

r = UBound(arr, 1)
c = UBound(arr, 2)

ReDim values(c)
ReDim rows(r)
offset = 0
For i = 0 To r
    cond = condition
    For j = 0 To c
        values(j) = arr(i, j)
       'user is 1-based, internal is 0-based
       'variables with $ are quoted as string values
        field = Format(j + 1, "$0")
        cond = Replace(cond, field, """" + values(j) + """")
        'variables with # are interpreted as numeric values
        field = Format(j + 1, "%0")
        'only replace when needed, as relDouble costs
        If InStr(cond, field) Then
            cond = Replace(cond, field, Str(relDouble(values(j))))
        End If
    Next j
    'put expression in container to have always legal expressions
    cond = "=(" + cond + ")"
    'not expression must have english and not local syntax (, instead of ;)
    eval = Application.Evaluate(cond)
    If IsError(eval) Then
        relSelect = "#ERROR CONDITION LINE " + Str(i + 1) + " : " + cond
    Exit Function
    End If
    If eval Then
        rows(offset) = Join(values, "::")
        offset = offset + 1
    End If
Next i

If offset = 0 Then
    relSelect = ""
    Exit Function
End If

ReDim Preserve rows(offset - 1)

relSelect = Join(rows, relNewline()) 'no duplicate elimination needed

End Function

Function relExtend(calculation As String, rel As String)
Attribute relExtend.VB_Description = "Adds a column to a relation based on a calculation (any Excel expresssion). Use $ for string columns, % for number."
Attribute relExtend.VB_ProcData.VB_Invoke_Func = " \n14"

Dim arr()  As Variant
Dim values() As String
Dim rows() As String
Dim cond As String
Dim field As String
Dim r, c, i, j, offset As Integer
Dim result As Variant

If rel = "" Then
    relExtend = ""
    Exit Function
End If

arr = relArray(rel)

r = UBound(arr, 1)
c = UBound(arr, 2)

ReDim values(c + 1)
ReDim rows(r)
offset = 0
For i = 0 To r
    cond = calculation
    For j = 0 To c
        values(j) = arr(i, j)
       'user indexes 1-based, internal 0-based
       'variables with $ are quoted as string values
        field = Format(j + 1, "$0")
        cond = Replace(cond, field, """" + values(j) + """")
        'variables with # are interpreted as numeric values
        field = Format(j + 1, "%0")
        If InStr(cond, field) Then
            cond = Replace(cond, field, Str(relDouble(values(j))))
        End If
    Next j
    'container to get valid expressions
    cond = "=(" + cond + ")"
    result = Application.Evaluate(cond)
    If IsError(result) Then
        relExtend = "#ERROR CALCULATION LINE " + Str(i + 1) + " : " + cond
        Exit Function
    End If
    values(c + 1) = result
     rows(offset) = Join(values, "::")
     offset = offset + 1
Next i


ReDim Preserve rows(offset - 1)

relExtend = Join(rows, relNewline()) 'no duplicate elimination needed

End Function


Function relProject(list As String, rel As String)
Attribute relProject.VB_Description = "Filters columns by list, separated by :: You can aggregate columns with SUM COUNT MAX MIN AVG"
Attribute relProject.VB_ProcData.VB_Invoke_Func = " \n14"

' project is also used for group aggregation, as relational algebra does not know duplicates

Dim arr()  As Variant
Dim aggregators() As String
Dim rows() As String
Dim rowkey As String
Dim values() As String
Dim test As Variant
Dim cols() As String
Dim vfields() As String
Dim cstring, vstring As String
Dim r, c, i, j, excluded, cval As Integer
Dim v1, v2 As Double
Dim found, hasaverage As Boolean

Dim aggregator() As String
Dim dict As New Collection
Dim cc As Integer

If list = "" Then
    relProject = ""
    Exit Function
End If

If list = " " Then
    relProject = "#ERROR LIST EMPTY"
    Exit Function
End If

If rel = "" Then
    relProject = "#ERROR COLUMN: " + list
    Exit Function
End If


arr = relArray(rel)

r = UBound(arr, 1)
c = UBound(arr, 2)
excluded = 0


'check list and build aggregators
cols = Split(list, "::")
c2 = UBound(cols) ' onebased
ReDim aggregator(c2)
' check limit
For i = 0 To c2
    cstring = cols(i)
    aggregator(i) = ""
    If InStr(cstring, " ") Then
        Dim words() As String
        words = Split(cstring, " ")
        Select Case words(0)
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
            Case Else
                relProject = "#ERROR AGGREGATOR: " + cstring
                Exit Function
        End Select
        cstring = words(1)
        usecollection = True
    End If
    cval = Val(cstring)
    cols(i) = cval
    If cval < 1 Or cval > c + 1 Then
            relProject = "#ERROR COLUMN: " + cstring
        Exit Function
    End If
Next i

' project and aggregate
' we use a collection to get unique keys
ReDim rows(r)
For i = 0 To r
    ReDim values(c2)
    For j = 0 To c2
        If aggregator(j) = "" Then
            cc = cols(j) - 1
            If cc < 0 Or cc > c Then
                relProject = "#ERROR COLUMN: " + Str(j + 1)
                Exit Function
            End If
            values(j) = arr(i, cc)
        End If
    Next j
    rowkey = Join(values, "::")
    If relInCollection(dict, rowkey) Then
        ' you cannot change a collection, so we have to remove and add it later.
        values = Split(dict.Item(rowkey), "::")
        dict.Remove (rowkey)
    End If
     For j = 0 To c2
        cc = cols(j) - 1
        If cc < 0 Or cc > c Then
            relProject = "#ERROR AGGREGATOR: " + Str(j + 1)
            Exit Function
        End If
        Select Case aggregator(j)
        Case "SUM"
            v1 = relDouble(values(j))
            v2 = relDouble(arr(i, cc))
            values(j) = Trim(Str(relDouble(values(j)) + relDouble(arr(i, cc))))
        Case "COUNT"
            values(j) = Trim(Str(relDouble(values(j)) + 1))
        Case "MAX"
            v1 = relDouble(values(j))
            v2 = relDouble(arr(i, cc))
            If values(j) = "" Or v2 > v1 Then values(j) = arr(i, cc)
        Case "MIN"
          v1 = relDouble(values(j))
          v2 = relDouble(arr(i, cc))
         If values(j) = "" Or v2 < v1 Then values(j) = arr(i, cc)
        Case "AVG"
            vstring = values(j)
            If vstring = "" Then vstring = "0/0"
            vfields = Split(vstring, "/")
            v1 = relDouble(vfields(0))
            v2 = relDouble(vfields(1))
            v1 = v1 + relDouble(arr(i, cc))
            v2 = v2 + 1
            ' we pack the sum and the count into one value
            vstring = Trim(Str(v1)) + "/" + Trim(Str(v2))
            values(j) = vstring
        End Select
       
    Next j
    dict.Add Join(values, "::"), rowkey
Next i

ReDim rows(dict.Count - 1)
For i = 1 To dict.Count
    rows(i - 1) = dict.Item(i)
    If hasaverage Then
        values = Split(rows(i - 1), "::")
        For j = 0 To c2
            If aggregator(j) = "AVG" Then
                 'we need to make the division of sum/count
                 vstring = values(j)
                 vfields = Split(vstring, "/")
                 v1 = relDouble(vfields(0))
                 v2 = relDouble(vfields(1))
                 ' we never have 0 division here, haven't we
                 vstring = Trim(Str(v1 / v2))
                 values(j) = vstring
            End If
        Next j
        rows(i - 1) = Join(values, "::")
    End If
Next i

relProject = Join(rows, relNewline()) 'no duplicate elimination needed
Exit Function



End Function

Function relJoin(condition As String, rel1 As String, rel2 As String)
Attribute relJoin.VB_Description = "Calculates a theta join of two relations based on a condition (any Excel expression). Use $ for string col, % for number"
Attribute relJoin.VB_ProcData.VB_Invoke_Func = " \n14"

'this is a theta join
'for cross just set conditon "true"
'for natural join set condition with column equality like "AND($1=$4,$2=$5)"
'if you set equal condition on all columns, you get an intersection

Dim rows1() As String
Dim rows2() As String
Dim rows() As String
Dim values1() As String
Dim values2() As String
Dim row As String
Dim first1, first2 As String
Dim r1, r2, c1, c2, i, j, k, l, offset As Integer
Dim eval As Variant

If rel1 = "" Or rel2 = "" Then
    relJoin = ""
    Exit Function
End If

rows1 = Split(rel1, relNewline())
rows2 = Split(rel2, relNewline())

r1 = UBound(rows1)
r2 = UBound(rows2)

first1 = relRow(rel1, 0)
first2 = relRow(rel2, 0)

values1 = Split(first1, "::")
values2 = Split(first2, "::")
c1 = UBound(values1)
c2 = UBound(values2)


offset = 0
ReDim rows(r1 + 1) 'we will make it bigger later when needed
For i = 0 To r1
    For j = 0 To r2
        row = rows1(i) + "::" + rows2(j)
        values = Split(row, "::")
        c = UBound(values)
    
        cond = condition
        For k = 0 To c
            'user index 1-based, internal 0-based
            'variables with $ are quoted as string values
            field = Format(k + 1, "$0")
            cond = Replace(cond, field, """" + values(k) + """")
            'variables with # are interpreted as numeric values
             field = Format(k + 1, "%0")
             If InStr(cond, field) Then
                cond = Replace(cond, field, Str(relDouble(values(k))))
            End If
        Next k
        cond = "=(" + cond + ")"
        eval = Application.Evaluate(cond)
        If IsError(eval) Then
            relJoin = "#ERROR CONDITION LINE " + Str(i + 1) + "/" + Str(j + 1) + " : " + cond
            Exit Function
        End If
        If eval Then
            If offset > UBound(rows) Then
                'we grow the array only as much as needed
                ReDim Preserve rows(2 * offset)
            End If
            rows(offset) = row
            offset = offset + 1
        End If
     Next j
Next i

If offset = 0 Then
    relJoin = ""
    Exit Function
End If

ReDim Preserve rows(offset - 1)

relJoin = relString(rows)


End Function

Function relOrder(list As String, rel As String)
Attribute relOrder.VB_Description = "Orders the relation by list, separated by :: Each column has modifier A Z (alphabetic) 1 9 (numeric)."
Attribute relOrder.VB_ProcData.VB_Invoke_Func = " \n14"
Dim arr()  As String
Dim values() As String
Dim rows() As String
Dim orderlist() As String
Dim order As String
Dim cols() As Integer
Dim modes() As String
Dim fields() As String
Dim cond As String
Dim field As String
Dim handrow As String
Dim other As String
Dim bigger As Variant
Dim r, c, c2, i, j, offset As Integer

arr = Split(rel, relNewline())

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
    order = orderlist(i)
    fields = Split(order, " ")
    cols(j) = fields(0) - 1
    modes(j) = fields(1)
    If fields(1) = "" Then modes(j) = "A"
Next

ReDim values(c)

'insertion sort
'all rows before the ith row are ordered (row 0 is sorted at beginning)
'we take the ith row in to hand, compare to all precedent rows, move them forward
'until the row in the hand is bigger, then we insert
'the comparison in costum function relBigger
For i = 1 To r
    handrow = arr(i)
    j = i
    Do
        j = j - 1
        other = arr(j)
        bigger = relBigger(other, handrow, cols, modes)
        
        If bigger <> True And bigger <> False Then
            relOrder = "#ERROR " + bigger
            Exit Function
        End If
        If bigger Then
            arr(j + 1) = other
        Else
              j = j + 1
              Exit Do
        End If
    Loop Until j = 0
    arr(j) = handrow
Next i

relOrder = Join(arr, relNewline()) 'no duplicate elimination needed

End Function


Private Function relBigger(row1 As String, row2 As String, cols() As Integer, modes() As String)
    Dim v1s() As String
    Dim v2s() As String
    Dim v1, v2 As String
    Dim i, c2 As Integer
    Dim test As Integer
    
    
    
    v1s = Split(row1, "::")
    v2s = Split(row2, "::")
    
    c2 = UBound(cols)
    
    For i = 0 To c2
        If cols(i) < 0 Or cols(i) > UBound(v1s) Then
        'error
        relBigger = "COLUMN " + Str(cols(i) + 1)
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
           test = Sgn(relDouble(v1) - relDouble(v2))
        Case "9"
           test = Sgn(relDouble(v2) - relDouble(v1))
        Case Else
           'error
           relBigger = "MODE " + modes(i)
           Exit Function
       End Select
       
       Select Case test
        Case 1
            relBigger = True
            Exit Function
        Case -1
            relBigger = False
            Exit Function
        Case 0
            'both are equal at this level, we need to go to the next on the list
            relBigger = False
        End Select
  Next i
    'both are equal (possible, if not all fields are in order or if numeric value equal)
    relBigger = False
End Function



Private Function relArray(s As String) As Variant
Dim rows() As String
Dim fields() As String
Dim cells() As Variant
Dim r, c, i, j As Integer

' converts a relation to a 0-based 2-dimensional array

rows = Split(s, relNewline())
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

relArray = cells

End Function

Private Function relString(arr() As String)
Dim tuples(), tuple, fields() As String
Dim duplicates, r As Integer
Dim found As Boolean

' converts an array to a relation and eliminates duplicates

duplicates = 0 ' we need to count them to redim the array

r = UBound(arr, 1)

ReDim tuples(r)
ReDim fields(c)

For i = 0 To r
    tuple = arr(i)
    found = False
    
    ' find duplicates in lower tuples
    For j = 0 To i - 1
        If tuples(j) = tuple Then
            found = True
            Exit For
        End If
    Next
    If found Then
        'one result less
        duplicates = duplicates + 1
    Else
       'index takes into account duplicates
        tuples(i - duplicates) = tuple
    End If
Next i
ReDim Preserve tuples(r - duplicates)
relString = Join(tuples, relNewline())

End Function




Function relCell(s As String, r As Integer, c As Integer, Optional noerror As Boolean = False)
Attribute relCell.VB_Description = "Recovers a cell of a relation."
Attribute relCell.VB_ProcData.VB_Invoke_Func = " \n14"

Dim tuple As Variant
Dim fields() As String

'user 1-based
c = c - 1
r = r - 1

tuple = relRow(s, r)

If IsNumeric(tuple) Then
    If noerror Then
        relCell = ""
    Else
        relCell = "#ERROR BOUNDS ROW: " + Str(r + 1)
    End If
    Exit Function
End If

fields = Split(tuple, "::")

If c < 0 Or c > UBound(fields) Then
    If noerror Then
        relCell = ""
    Else
        relCell = "#ERROR BOUNDS COLUMN: " + Str(c + 1)
    End If
    Exit Function
End If

relCell = fields(c)


End Function



Private Function relRow(s As String, r As Integer, Optional noerror As Boolean = False)
Dim tuples() As String

tuples = Split(s, relNewline())

If r < 0 Or r > UBound(tuples) Then
    If noerror Then
        relRow = ""
    Else
        relRow = -1 'error
    End If
    Exit Function
End If

relRow = tuples(r)


End Function


Private Function relNewline()
    'to be consistent between platfoms use the same separator
    relNewline = vbCr
End Function


Private Function relInCollection(col As Collection, key As String) As Boolean

' it is not possible to get a list of keys from a collection.
' so we just try to get the value and catch the error

On Error GoTo incol
  col.Item key

incol:
  relInCollection = (Err.Number = 0)

End Function

Function relLike(s As String, pattern As String)
Attribute relLike.VB_Description = "Exposes like to Excel. Operators: ? (any once), * (any zero or more), # (number), [] list [!] exclude list"
Attribute relLike.VB_ProcData.VB_Invoke_Func = " \n14"
  
    'expose like to excel
    relLike = s Like pattern
    
End Function



Public Function relCellArray(rel As String)
Attribute relCellArray.VB_Description = "Recovers multiple cells at once, using the {Ê} formula for a range."
Attribute relCellArray.VB_ProcData.VB_Invoke_Func = " \n14"
Dim c, r, i, j As Integer
Dim r1(), r2() As Variant

'relCell for a complete range
'empty if out of range

With Application.Caller
        r = .rows.Count
        c = .columns.Count
End With

r1 = relArray(rel)
ReDim r2(r, c)
For i = 0 To r
    For j = 0 To c
         If i < UBound(r1, 1) And j <= UBound(r1, 2) Then
            r2(i, j) = r1(i, j)
         Else
            r2(i, j) = ""
         End If
    Next j
Next i
relCellArray = r2


End Function

Public Function relFilter(list As String, condition As String, rn As Range, Optional listorder As String = "", Optional start As Integer = -1, Optional n As Integer = -1)
Attribute relFilter.VB_Description = "Shortcut for relLimit(relOrder(relProject(list, relSelect(condition, relRange(rn)))))"
Attribute relFilter.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim r As String
    Dim s As String
    Dim p As String
    Dim o As String
    Dim l As String
    Dim cond As String
    
    'short cut to keep result < 32K (limit in Excel, but not VBA)
    
    r = relRange(rn)
    cond = condition
    s = relSelect(cond, r)
    p = relProject(list, s)
    If listorder <> "" Then
        o = relOrder(listorder, p)
    Else
        o = p
    End If
     If start <> -1 Then
        l = relLimit(start, n, o)
    Else
        l = o
    End If
    
    relFilter = l
    
End Function

Public Function relLimit(start As Integer, n As Integer, rel As String)
Attribute relLimit.VB_Description = "Limits (an ordered) relation with start and limit parameter."
Attribute relLimit.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim rows() As String
    Dim i As Integer
    Dim result() As String
    
    If rel = "" Then
        relLimit = "#ERROR EMPTY"
        Exit Function
    End If
    
    rows = Split(rel, relNewline())
    
    If UBound(rows) <= start Then
       relLimit = "#ERROR OUT OF BOUNDS"
        Exit Function
    End If
    
    If start < 1 Then
       relLimit = "#ERROR OUT OF BOUNDS"
        Exit Function
    End If

    If n = 0 Then
        relLimit = ""
        Exit Function
    End If
    
    If n = -1 Then
        n = UBound(rows) - start
    End If
    
    ReDim result(n - 1)
    
    For i = 0 To n - 1
        If start + i > UBound(rows) Then
            result(i) = ""
        Else
            result(i) = rows(start + i)
        End If
    Next i
    
    relLimit = Join(result, relNewline())

End Function

Public Function relRotate(rel As String)
Attribute relRotate.VB_Description = "Rotates a relation by exchanging columns and rows"
Attribute relRotate.VB_ProcData.VB_Invoke_Func = " \n14"
   Dim arr1() As Variant
   Dim tuple() As String
   Dim arr2() As String
   Dim r, c, i, j As Integer
   
   ' columns to rows, row to columns
    
    If rel = "" Then
        relRotate = ""
        Exit Function
    End If
    
    arr1 = relArray(rel)
    
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
    
    relRotate = relString(arr2)
    
    

End Function


Private Function relDouble(v As Variant) As Double
   
   'accepts both , and . comma for fractions
   
    If IsNumeric(v) Then
        relDouble = CDbl(v)
        Exit Function
    End If
    
    v = Replace(v, ",", ".")
    
    If IsNumeric(v) Then
        relDouble = CDbl(v)
        Exit Function
    End If
    
    v = Replace(v, ".", ",")
    
    If IsNumeric(v) Then
        relDouble = CDbl(v)
        Exit Function
    End If
    
    relDouble = 0
    
    
    
End Function
