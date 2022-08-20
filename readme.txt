Relation for Excel

This macro provides functions to make simple relational algebra
The relational model is simplified. A relation is defined as a 2d-table, columns have names but not type. Results are always sets, so distinct and no row can be empty.

Unlike other Excel solutions, this one is purely functional, not using macros.
Relations are saved as text in one cell with :: als field and space+newline as row separator

Relation for Excel 3.0 adds SQL syntax. You can now make SQL queries in Excel using any range in a worksheet including joins.

Valid SQL queries are

SELECT * FROM t1
SELECT CustomerName, City FROM t1
SELECT * FROM t1 WHERE City='Berlin' OR City='München'
SELECT country, POW(2,3) AS c FROM t1
SELECT Country, COUNT(Country) AS n FROM t1 WHERE Country LIKE '%land%' HAVING n < 2 ORDER BY Country
SELECT t1.OrderID, t2.CustomerName, t1.OrderDate FROM t1 NATURAL JOIN t2
SELECT t1.CustomerName AS p, t2.OrderID AS ok FROM t1 JOIN t2 ON t1.CustomerID = t2.CustomerID ORDER BY p

Note that:

Valid instructions are SELECT AS FROM NATURAL LEFT RIGHT OUTER JOIN ON WHERE HAVING ORDER BY and the must be in capitals

Ranges are referred as t1 to t9 and used as arguments in the function

Expressions are allowed in columns, but they must have a qualified name with AS

Functions and aggregators can be mixed, but there can be only one aggregator

Strings use single quotes

Columns are not typed: context defines type. On comparision string bigger number bigger empty string

As relation, results are always distinct and grouping is auto, so there is no group by instruction

Operators: + - * / > < >= <= = <> LIKE IN
Aggregators: AVG COUNT MAX MEDIAN MIN STDEV SUM
Numerical functions: ABS COS EXP INT LN LOG MOD POW ROUND SGN SIN SQRT TAN
Text functions: LEFT LEN LOWER MID REPLACE RIGHT TRIM UPPER

Documentation

http://www.belle-nuit.com/relation-for-excel

matti@belle-nuit.com
21.8.202