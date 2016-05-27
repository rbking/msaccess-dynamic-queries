msaccess-dynamic-queries
========================

About
-----

Have you ever created a query, made copies of it with different WHERE statements, and then made multiple copies of a report based off of those queries.
It becomes very difficult to manage if you ever need to change the fields, change the joins, or add new tables.
One solution is you could hard code the query in VBA and make the where statement dynamic.  This too can be cumbersome when updates need to be made.

CSQLCompiler
-------

The CSQLCompiler Class Module uses SQL statements from existing queries and inserts a custom WHERE statement and then returns the new SQL statement.
The CSQLCompiler has 3 Properties and 3 Methods.

Properties
-------

* **rMainQueryName** - Required String - Set the name of the query to use and insert the WHERE statement into.
* **oMasterQueryName** - Optional String - Set the name of the query that the MainQuery is nested into.
* **oWhereStatement** - Optional String - Set a Custom WHERE statement to insert into the MainQuery.

Methods
-------        

* **gCompleteSQL** - Returns the Generated SQL statement with a semicolon (;) at the end.
* **gInnerSQL** - Returns the Generated SQL statement without the simicolon (;).  Useful if using the generated statement as a subquery.
* **ClearAll** - Resets the Properties.  Use this when reusing the object in a loop or later in the routine.

Example
-------
A database is included and it demonstrates how the class can be used with a form and report.

Caveats
-------
* If you have a SELECT Query of a GROUP BY Query all in one SQL Statement it will add the WHERE statement to the subquery.  

EXAMPLE:  
Query1:  

```
#!vba

	SQL = SELECT s.*, MSysQueries.Expression 
              FROM (
                    SELECT MSysObjects.Connect, MSysObjects.Database, MSysObjects.Id 
                    FROM MSysObjects 
                    GROUP BY MSysObjects.Connect, MSysObjects.Database, MSysObjects.Id
                    )  AS s INNER JOIN MSysQueries ON s.Id = MSysQueries.ObjectId;
```


VBA: 

```
#!vba

		Dim clssql As CSQLCompiler
		Set clssql = New CSQLCompiler
		With clssql
			.rMainQueryName = "Query1"
			.oWhereStatement = "WHERE MSysQueries.Expression = 1"
			
			Debug.Print .gCompleteSQL
		End With
```


RETURNS:

```
#!vba

	SELECT s.*, MSysQueries.Expression 
        FROM (
              SELECT MSysObjects.Connect, MSysObjects.Database, MSysObjects.Id 
              FROM MSysObjects WHERE MSysQueries.Expression = 1 
              GROUP BY MSysObjects.Connect, MSysObjects.Database, MSysObjects.Id
              )  AS s INNER JOIN MSysQueries ON s.Id = MSysQueries.ObjectId;
```


Solution:
     Make sure there is a WHERE Statement at the END of the top Query.

Query1:  

```
#!vba

	SQL = SELECT s.*, MSysQueries.Expression 
              FROM (
                    SELECT MSysObjects.Connect, MSysObjects.Database, MSysObjects.Id 
                    FROM MSysObjects 
                    GROUP BY MSysObjects.Connect, MSysObjects.Database, MSysObjects.Id
                    )  AS s INNER JOIN MSysQueries ON s.Id = MSysQueries.ObjectId
              WHERE MSysQueries.Expression = 0;
```



Contributing
============

Pull requests, issue reports etc welcomed.