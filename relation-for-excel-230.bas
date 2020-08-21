Attribute VB_Name = "Module1"
' Relation for Excel
' Version 2.3 20.8.2020
' matti@belle-nuit.com
'This module provides functions to make simple relational algebra
'The relational model is simplified.
'A relation is defined as a 2d-table, columns can be adressed by name or position (1-based
'However, rows are not ordered (except with relOrder function) and do not have duplicates
'Unlike other Excel solutions, this module is purely functional, not using macros.
'Relations are saved as text in one cell with :: als field and newline as row separator
'Note that in a cell, the text cannot be more than 32k characters.
'Version history 2.2: Added MEDIAN, optimisation expensive duplicate removal only when needed (before projection and at end)
'... dim correct one at a time, long replacing integer
 
Option Explicit
 
 
 
Function relRange(rn As Range, Optional hasheader As Long = True)
 
relRange = prelRange(rn, hasheader, False, False)
 
End Function
 
Private Function prelRange(rn As Range, hasheader As Long, noError As Boolean, lazy As Boolean) As String
 
' Calculates a relation from a range
' A relation is a table where rows are separated by newline and columns by ::
' We use a simplified model where tuples can have no named properties, but by position (1-based)
' If header is false, default number header will be used
' If header is true, first line of rn is considered header
' noerror is necessary for relFilter
' lazy is necessary for relFilter, removing duplicates only when necessary
 
Dim arr() As Variant
Dim hd() As Variant
Dim tuples() As String
Dim fields() As String
Dim r As Long
Dim c As Long
Dim i As Long
Dim j As Long
Dim first As Long
Dim l As String
Dim v As Variant
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
        'prelRange = "#ERROR CELL " + Str(i + first) + " " + Str(j + 1)
        Exit Function
    End If
   
    fields(j) = arr(i + first, j + 1)
Next j
    tuples(i) = Join(fields, "::")
Next i
 
If lazy Then
    l = Join(tuples, prelNewline()) 'no duplicate elimination needed
Else
    l = prelString(tuples)
End If
 
If Len(l) > 32768 And Not noError Then
        prelRange = "#ERROR LONG RESULT " & Str(Len(l))
        Exit Function
End If
prelRange = l
 
End Function
 
Function relUnion(ByVal rel1 As String, ByVal rel2 As String, Optional noError As Boolean = False, Optional lazy As Boolean = False) As String
Dim first1 As String
Dim first2 As String
Dim r As String
Dim fields1() As String
Dim fields2() As String
Dim rows1() As String
Dim rows2() As String
Dim header1list() As String
Dim header2list() As String
Dim afields() As String
Dim nfields() As String
Dim ub11 As Long
Dim ub12 As Long
Dim ub2 As Long
Dim s As String
Dim header1 As String
Dim header2 As String
Dim l As String
Dim c1 As Long
Dim c2 As Long
Dim r1 As Long
Dim r2 As Long
Dim i As Long
Dim j As Long
Dim n As Long
Dim columns() As Long
 
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
    l = Join(rows1, prelNewline()) 'no duplicate elimination needed
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
 
Dim first1 As String
Dim first2 As String
Dim r As String
Dim fields1() As String
Dim fields2() As String
Dim rows1() As String
Dim rows2() As String
Dim header1list() As String
Dim header2list() As String
Dim arr() As String
Dim ub11 As Long
Dim ub12 As Long
Dim ub2 As Long
Dim s As String
Dim header1 As String
Dim header2 As String
Dim found As Boolean
Dim r1 As Long
Dim r2 As Long
Dim c1 As Long
Dim c2 As Long
Dim n As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l  As Long
Dim offset As Long
 
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
    If (rows2(i)) <> "" Then
    fields1 = Split(rows2(i), "::")
    ReDim fields2(c2)
    For j = 0 To c2
        fields2(columns(j)) = fields1(j)
    Next j
    rows2(i) = Join(fields2, "::")
    End If
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
 
Dim first1 As String
Dim first2 As String
Dim r As String
Dim fields1() As String
Dim fields2() As String
Dim arr() As String
Dim rows1() As String
Dim rows2() As String
Dim header1list() As String
Dim header2list() As String
Dim ub11 As Long
Dim ub12 As Long
Dim ub2 As Long
Dim s As String
Dim header1 As String
Dim header2 As String
Dim found As Boolean
Dim r1 As Long
Dim r2 As Long
Dim c1 As Long
Dim c2 As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long
Dim offset As Long
Dim n As Long
 
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
Dim cond As String
Dim header As String
Dim field As String
Dim r As Long
Dim c As Long
Dim i As Long
Dim j As Long
Dim offset As Long
Dim eval As Variant
 
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
    values = Split(rows(i), "::")
    cond = prelParseExpression(condition, values)
    eval = Application.Evaluate(cond)
    If IsError(eval) Then
        relSelect = "#ERROR CONDITION LINE " & Str(i + 1) & " : " & cond
    Exit Function
    End If
    If eval Then
        offset = offset + 1
        rows(offset) = Join(values, "::")
    End If
Next i
 
'If offset = 0 Then
'  relSelect = ""
'  Exit Function
'End If
 
ReDim Preserve rows(offset)
 
relSelect = Join(rows, prelNewline()) 'no duplicate elimination needed
 
End Function
 
Function relExtend(rel As String, ByVal calculation As String, Optional ByVal name As String, Optional noError As Boolean = False) As String
 
Dim arr()  As Variant
Dim values() As String
Dim rows() As String
Dim newlist() As String
Dim cond As String
Dim header As String
Dim field As String
Dim l As String
Dim r As Long
Dim c As Long
Dim i As Long
Dim j As Long
Dim offset As Long
Dim result As Variant
 
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
        result = Application.Evaluate(cond)
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
Dim r As Long
Dim c As Long
Dim c2 As Long
Dim i As Long
Dim j As Long
Dim excluded As Long
Dim cval As Long
Dim v1 As Double
Dim v2 As Double
Dim v3 As Double
Dim s1 As String
Dim found, hasaverage As Boolean
Dim header, newheader As String
Dim newlist() As String
Dim usecollection As Boolean
Dim medianlines As Long
 
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
        End Select
      
    Next j
    dict.Add Join(values, "::"), rowkey
Next i
 
ReDim rows(dict.Count)
rows(0) = Join(newlist, "::")
For i = 1 To dict.Count
    rows(i) = dict.Item(i)
    If hasaverage Then
        values = Split(rows(i), "::")
        For j = 0 To c2
            If aggregator(j) = "AVG" Then
                 'we need to make the division of sum/count
                 vstring = values(j)
                 vfields = Split(vstring, "/")
                 v1 = prelDouble(vfields(0))
                 v2 = prelDouble(vfields(1))
                 ' we never have 0 division here, haven't we
                 vstring = Trim(Str(v1 / v2))
                 values(j) = vstring
            End If
            If aggregator(j) = "STDEV" Then
                 'we need to make the division of sum/count
                 vstring = values(j)
                 vfields = Split(vstring, "/")
                 v1 = prelDouble(vfields(0))
                 v2 = prelDouble(vfields(1))
                 v3 = prelDouble(vfields(2))
                 
                 
                 ' we never have 0 division here, haven't we
                 vstring = Trim(Str(VBA.Sqr(v1 / v3 - (v2 * v2) / (v3 * v3))))
                 values(j) = vstring
            End If
            If aggregator(j) = "MEDIAN" Then
                s1 = values(j)
                s1 = relOrder(s1, "vmedian 9")
                vfields = Split(s1, prelNewline())
                medianlines = UBound(vfields)
                If Round(medianlines / 2, 0) = medianlines / 2 Then
                    'even
                    v1 = prelDouble(vfields(medianlines / 2))
                    v2 = prelDouble(vfields(medianlines / 2 + 1))
                    values(j) = Str((v1 + v2) / 2)
                Else
                    values(j) = vfields(Round(medianlines / 2 + 0.1))
                   
                End If
            End If
        Next j
        rows(i) = Join(values, "::")
    End If
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
 
Private Function prelExists(coll As Collection, key As String) As Boolean
 
    On Error GoTo EH
 
    IsObject (coll.Item(key))
   
    prelExists = True
EH:
End Function
 
Private Function prelSpecialJoin(rel1 As String, rel2 As String, Optional flag As String = "", Optional noError As Boolean = False, Optional lazy As Boolean = False) As String
 
'this is a natural join on common fields
'flags: "NATURAL" (default), "LEFT",
'"LEFTSEMI", "LEFTANTISEMI"
 
On Error GoTo errHandler
 
 
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
Dim row As String
Dim first1 As String
Dim first2 As String
Dim empty1 As String
Dim empty2 As String
Dim r As Long
Dim r1 As Long
Dim r2 As Long
Dim c As Long
Dim c1 As Long
Dim c2 As Long
Dim o1 As Long
Dim o2 As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As String
Dim offset As Long
Dim commoncolumns As Long
Dim eval As Variant
Dim found As Boolean
 
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
    For j = 0 To c2
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
    If prelExists(keyhash, hexkey) Then
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
    Case "NATURAL", "LEFT", ""
        rows(0) = keys1(0) & rest1(0) & rest2(0)
    Case "LEFTSEMI", "LEFTANTISEMI"
        rows(0) = keys1(0) & rest1(0)
End Select
 
'rows
offset = 1
For i = 1 To r1
    found = False
    hexkey = prelAsciiToHexString(keys1(i))
    If prelExists(keyhash, hexkey) Then
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
Dim header As String
Dim cond As String
Dim row As String
Dim first1 As String
Dim first2 As String
Dim r1 As Long
Dim r2 As Long
Dim c1 As Long
Dim c2 As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As String
Dim offset As Long
Dim eval As Variant
 
Select Case condition
Case "CROSS"
    relJoin = relJoin(rel1, rel2, "TRUE", noError, lazy)
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
 
condition = prelSubstituteNames(condition, header)
 
offset = 1
ReDim rows(r1 + 1) 'we will make it bigger later when needed
rows(0) = header
For i = 1 To r1
    For j = 1 To r2
        row = rows1(i) & "::" & rows2(j)
        values = Split(row, "::")
       
        cond = prelParseExpression(condition, values)
       
        eval = Application.Evaluate(cond)
        If IsError(eval) Then
            relJoin = "#ERROR CONDITION LINE " & Str(i + 1) & "/" & Str(j + 1) & " : " & cond
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
 
'If offset = 1 Then
'    relJoin = ""
'    Exit Function
'End If
 
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
 
 
 
 
Function relOrder(rel As String, list As String) As String
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
Dim r As Long
Dim c As Long
Dim c2 As Long
Dim i As Long
Dim j As Long
Dim offset As Long
 
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
    Dim v1 As String
    Dim v2 As String
    Dim i As Long
    Dim c2 As Long
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
Dim r As Long
Dim c As Long
Dim i As Long
Dim j As Long
 
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
 
' converts an array to a relation and eliminates duplicates
 
r = UBound(arr, 1)
 
Dim hash As New Collection
 
For i = 0 To r
    If prelExists(hash, arr(i)) Then
        ' do nothing
    Else
        hash.Add arr(i), arr(i)
        End If
Next i
 
Dim c As Integer
c = hash.Count
ReDim tuples(c - 1)
 
For i = 1 To c
     tuples(i - 1) = hash.Item(i)
Next
 
 
prelString = Join(tuples, prelNewline())
 
 
 
End Function
 
 
 
 
Function relCell(rel As String, r As Long, c As Long, Optional Numeric As Boolean = False, Optional noError As Boolean = True) As Variant
 
Dim tuples() As String
Dim tuple As Variant
Dim fields() As String
 
 
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
 
 
 
 
 
 
 
Private Static Function prelNewline() As String
   
    Dim platform As Long
   
    Select Case platform
    Case 1
        prelNewline = " " & vbCrLf
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
 
incol:
  prelInCollection = (Err.Number = 0)
 
End Function
 
Function relLike(s As String, pattern As String)
 
    'expose like to excel
    relLike = s Like pattern
   
End Function
 
 
 
Public Function relCellArray(rel As String, Optional noHeader As Boolean = False)
Dim c As Long
Dim r As Long
Dim i As Long
Dim j As Long
Dim r1() As Variant
Dim r2() As Variant
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
    Dim elem As Variant
    Dim test As String
    Dim body As String
    Dim fields() As String
    Dim v1 As Long
    Dim v2 As Long
    Dim stack() As String
    Dim stackpointer As Long
    Dim rn As Range
    Dim arr() As Variant
    Dim done As Boolean
    Dim s As String
    Dim extendname As String
    Dim extendbody As String
   
    On Error GoTo errHandler
   
    ReDim stack(0)
    stack(0) = ""
    stackpointer = 0
   
    For Each elem In list
        'test for range bigger than one dimension
        done = False
        If VarType(elem) = vbError Then
            done = True
        ElseIf TypeName(elem) = "Range" Then
            Set rn = elem
            If rn.Count > 1 Then
                stackpointer = stackpointer + 1
                ReDim Preserve stack(stackpointer)
                stack(stackpointer) = prelRange(rn, True, True, True)
                done = True
            Else
                elem = rn.Value2
            End If
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
                        relFilter = "#EMPTY STACK UNION"
                        Exit Function
                   End If
                   stackpointer = stackpointer - 1
                   stack(stackpointer) = relUnion(stack(stackpointer), stack(stackpointer + 1), True, True)
                 Case "D"
                   If stackpointer < 1 Then
                        relFilter = "#EMPTY STACK DIFFERENCE"
                        Exit Function
                   End If
                   stackpointer = stackpointer - 1
                   stack(stackpointer) = relDifference(stack(stackpointer), stack(stackpointer + 1))
                 Case "I"
                   If stackpointer < 1 Then
                        relFilter = "#EMPTY STACK INTERSECT"
                        Exit Function
                   End If
                   stackpointer = stackpointer - 1
                   stack(stackpointer) = relIntersect(stack(stackpointer), stack(stackpointer + 1))
                Case "J"
                   If stackpointer < 1 Then
                        relFilter = "#EMPTY STACK JOIN"
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
                        relFilter = "#MISSING ARGUMENT LIMIT"
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
                        relFilter = fields(1)
                        If InStr(fields(1), "::") Then
                            fields = Split(fields(1), "::")
                            relFilter = fields(0)
                        End If
                        If prelDouble(relFilter) > 0 Then relFilter = prelDouble(relFilter)
                        If prelDouble(relFilter) < 0 Then relFilter = prelDouble(relFilter)
                        If relFilter = 0 Then relFilter = 0
                        Exit Function
                    Else
                        relFilter = ""
                        Exit Function
                    End If
                Case "K" ' single cell forced text
                    s = stack(stackpointer)
                    fields = Split(s, prelNewline())
                    If UBound(fields) > 0 Then
                        relFilter = fields(1)
                        If InStr(fields(1), "::") Then
                            fields = Split(fields(1), "::")
                            relFilter = fields(0)
                        End If
                        Exit Function
                    Else
                        relFilter = ""
                        Exit Function
                    End If
                Case "Z" ' single cell forced number
                    s = stack(stackpointer)
                    fields = Split(s, prelNewline())
                    If UBound(fields) > 0 Then
                        relFilter = fields(1)
                        If InStr(fields(1), "::") Then
                            fields = Split(fields(1), "::")
                            relFilter = fields(0)
                        End If
                        relFilter = prelDouble(relFilter)
                        Exit Function
                    Else
                        relFilter = 0
                        Exit Function
                    End If
                Case "!" 'cut
                    relFilter = stack(stackpointer)
                    Exit Function
                Case "#"
                   ' ignore
             Case Else
                      relFilter = "#INVALID OPERATOR " & test
                      Exit Function
             End Select
        End If
       
    Next elem
   
    If Len(stack(stackpointer)) > 32768 Then
        relFilter = "#ERROR LONG RESULT " & Str(Len(stack(stackpointer)))
        Exit Function
    End If
 
    
    relFilter = stack(stackpointer)
   
    'as we are lazy, we need to remove duplicates now
    fields = Split(relFilter, prelNewline())
    relFilter = prelString(fields)
   
    ReDim stack(0)
   
    
    
    Exit Function
   
errHandler:
    relFilter = "Error relFilter " & Err.Number & ": " & Err.Description
 
 
   
 
End Function
 
 
 
Public Function relLimit(rel As String, ByVal start As Long, ByVal n As Long)
    Dim rows() As String
    Dim i As Long
    Dim result() As String
   
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
 
Public Function relRotate(rel As String)
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
Dim cond As String
Dim field As String
Dim j As Long
Dim c As Long
 
cond = condition
 
c = UBound(values)
 
' going top down to avoid ambiguities $1 $10
For j = c To 0 Step -1
    field = Format(j + 1, "$00")
    If InStr(cond, field) Then
        cond = Replace(cond, field, """" & values(j) & """")
    End If
    field = Format(j + 1, "$0")
    If InStr(cond, field) Then
        cond = Replace(cond, field, """" & values(j) & """")
    End If
    field = Format(j + 1, "%00")
    If InStr(cond, field) Then
        cond = Replace(cond, field, Trim(Str(prelDouble(values(j)))))
    End If
    field = Format(j + 1, "%0")
    If InStr(cond, field) Then
        cond = Replace(cond, field, Trim(Str(prelDouble(values(j)))))
    End If
 
Next j
 
'put expression in container to have always legal expressions
'note that expression must have english and not local syntax (, instead of ;)
 
cond = "=(" & cond & ")"
 
prelParseExpression = cond
 
End Function
 
Function prelNameToColumn(ByVal header As String, ByVal name As String) As Long
Dim fields() As String
Dim c As Long
Dim i As Long
 
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
    prelNameToColumn = 0
 
End Function
 
Private Function prelSubstituteNames(ByVal expression As String, ByVal header As String) As String
Dim headerlist() As String
Dim i As Long
Dim c As Long
Dim n As Long
Dim field As String
Dim afield As String
Dim nfield As String
 
headerlist = Split(header, "::") 'to do this list should be sorted by length to sort out ambiguities
 
 
 
c = UBound(headerlist)
 
For i = 0 To c
    field = headerlist(i)
    n = prelNameToColumn(header, field)
    afield = "$" & field
    nfield = Format(n, "$0")
    If InStr(expression, afield) Then
        expression = Replace(expression, afield, nfield)
    End If
    afield = "%" & field
    nfield = Format(n, "%0")
    If InStr(expression, afield) Then
        expression = Replace(expression, afield, nfield)
    End If
Next i
 
prelSubstituteNames = expression
 
 
End Function
 
Function relRename(ByVal rel As String, ByVal list As String) As String
 
Dim arr()  As Variant
Dim values() As String
Dim rows() As String
Dim newlist() As String
Dim fields() As String
Dim words() As String
Dim cond As String
Dim header As String
Dim field As String
Dim r As Long
Dim c As Long
Dim i As Long
Dim j As Long
Dim offset As Long
Dim n As Long
Dim result As Variant
 
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
Dim cond As String
Dim header As String
Dim field As String
Dim condition As String
Dim r As Long
Dim c As Long
Dim i As Long
Dim j As Long
Dim offset As Long
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
 
Function relFixpoint(ByVal rel As String, fixpoint As String, ByVal start As String, connect As String)
 
Dim rows1() As String
Dim rows2() As String
Dim values1() As String
Dim values2() As String
Dim list() As String
Dim header As String
Dim tuple1 As String
Dim tuple2 As String
Dim header0 As String
Dim r As Long
Dim col1 As Long
Dim col2 As Long
Dim offset1 As Long
Dim offset2 As Long
Dim level As Long
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
Dim c As Long
Dim i As Long
Dim j As Long
 
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



