VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSQLCompiler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'// Class Name:      CSQLCompiler
'// Created By:      Robert King
'// Created Date:    3/17/2015
'// Description:     This Class is Used to Generate SQL Statements
'// Dependencies:    NONE

Option Explicit

Private sdInfo As Object

'// Let Properties
Public Property Let rMainQueryName(ByVal sValue As String):   sdInfo.Item("sQry") = sValue: End Property
Public Property Let oWhereStatement(ByVal sValue As String):  sdInfo.Item("sWhr") = sValue: End Property
Public Property Let oMasterQueryName(ByVal sValue As String): sdInfo.Item("mQry") = sValue: End Property

'// Get Properties
Public Property Get rMainQueryName() As String:   rMainQueryName = sdInfo.Item("sQry"):     End Property
Public Property Get oWhereStatement() As String:  oWhereStatement = sdInfo.Item("sWhr"):    End Property
Public Property Get oMasterQueryName() As String: oMasterQueryName = sdInfo.Item("mQry"):   End Property

Public Property Get gCompleteSQL() As String
    gCompleteSQL = gInnerSQL & ";"          '// Add a semicolon (;) to the end of the Generated SQL statement
End Property
Public Property Get gInnerSQL() As String
                                            '// If an SQL Statement has not been generated already, then generate it
    If Not sdInfo.Item("bRn") <> False _
        Then CompileSQL
    gInnerSQL = sdInfo.Item("cSql")         '// Return the SQL Statement without a semicolon (;)
End Property

'// Public Functions
Public Function ClearAll() As Boolean
    sdInfo.RemoveAll
End Function

'// Private Subs
Private Sub CompileSQL()
    If Len(Trim$(Me.rMainQueryName)) = 0 Then                                   '// Make sure rMainQueryName has been provided
                                                                                '// Tell User that rMainQueryName is required
        MsgBox "rMainQueryName is required." _
               , vbCritical + vbOKOnly _
               , "Required Field"
        Exit Sub                                                                '// Exit the sub
    End If
    With sdInfo
        .Item("sSql") = CurrentDb.QueryDefs(rMainQueryName).sql                 '// Bring in the Main Query's SQL to Start Parsing
        
        If InStr(.Item("sSql"), ";") > 0 Then _
            .Item("sSql") = Replace(.Item("sSql"), ";", vbNullString)           '// Take out the Semicolon a the end of the SQL
            
        .Item("fPh") = InStrRev(.Item("sSql"), "WHERE ")                        '// Find the last Where Statement in the Main Query
        .Item("sPh") = InStr(.Item("fPh") + 1, .Item("sSql"), "GROUP BY ")      '// See if there is a Group By in behind the Where Statement.
                                                                                '// .Item("fPh") + 1  is needed so the function doesn't error
                                                                                '// out if there is no where statement.
                                                                            
                                                                                '// Check for a Where Statement and split the string apart
        If (.Item("fPh") <> 0) And (InStr(.Item("fPh") + 1, .Item("sSql"), " as ") = 0 _
        Or Len(Trim$(oMasterQueryName)) = 0) _
        Then                                                                    '// There is a where clause and it is not part of a subquery
            .Item("fHlf") = Left$(.Item("sSql"), (.Item("fPh") - 1))            '// Hold the SQL part before the Where for future substitution
        Else                                                                    '// There is no Where Clause so the whole statement can be used
            .Item("fHlf") = .Item("sSql")
        End If
                                                                            
        If Not .Item("sPh") = 0 Then                                            '// Seperate if there is a Group By Section and hold the end of the Sql Statement
            .Item("sHlf") = ChrW(32) _
                            & Mid$(.Item("sSql"), .Item("sPh"))                 '// Add a space infront of the GROUP BY Statement
            .Item("fHlf") = Left$(.Item("fHlf"), (.Item("sPh") - 1))            '// Seperate the First Half of the Statement
        End If

        .Item("sSql") = .Item("fHlf") _
                        & oWhereStatement _
                        & .Item("sHlf")                                         '// Insert WHERE statement provided into the query
                    
        On Error GoTo CmplTestQryError
        If Not oMasterQueryName = vbNullString Then                             '// If Master Query is Given
            .Item("mSql") = CurrentDb.QueryDefs(oMasterQueryName).sql           '// Get the SQL of the Master Query
            
            If InStr(.Item("mSql"), ";") > 0 Then _
                .Item("mSql") = Replace(.Item("mSql"), ";", vbNullString)       '// Replace the ; if it exists
                
            .Item("fPh") = InStr(1, .Item("mSql"), "FROM ")                     '// Find the first from in the Master Query.  This will be the Sub Query
            .Item("fPh") = InStr(.Item("fPh"), .Item("mSql"), rMainQueryName)   '// Find the first instance of the Sub Query's Name
            .Item("fHlf") = Left$(.Item("mSql"), (.Item("fPh") - 1))            '// Hold the first part of the Master Query's Sql
            .Item("sHlf") = Mid$(.Item("mSql"), .Item("fPh"))                   '// Hold the last part of the Master Query's Sql Statement.
            
                                                                                '// Insert the Sub Query's SQL Statement,
                                                                                '// Inclose in paraenthesis while nameing it the Subquery's name,
                                                                                '// and then add the end of the Master Query's SQL.
            .Item("cSQL") = .Item("fHlf") _
                          & "(" _
                          & .Item("sSql") _
                          & ") AS " _
                          & .Item("sHlf")
                          
        Else
            .Item("cSQL") = .Item("sSql")                                       '// Only compile the subquery and return the sql.
        End If
        On Error GoTo 0
                                                                                        
        .Item("cSql") = Trim$(.Item("cSQL"))                                    '// Clean up the Query
        .Item("bRn") = True                                                     '// Only need to Generate SQL once and hold it in Memory
    End With
    
CmplTestQryResume:
        
Exit Sub

CmplTestQryError:
                                                                                '// Send a Message Box to tell the user
                                                                                '// that the Main Query is not Part of the Master Query
    MsgBox rMainQueryName _
           & " was not found in " _
           & oMasterQueryName _
           , vbOKOnly + vbCritical _
           , "Subquery not found"
    GoTo CmplTestQryResume
End Sub

Private Sub Class_Initialize()
    Set sdInfo = CreateObject("Scripting.Dictionary")
End Sub

Private Sub Class_Terminate()
    Me.ClearAll
    Set sdInfo = Nothing
End Sub
