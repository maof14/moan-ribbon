Imports System.Data

''' <summary>
''' Define interface for the two database classes; MySQL and SQLite. 
''' </summary>
''' <remarks>Is interface the way to go, or Abstract (MustInherit)?</remarks>
Public Interface IDatabase

    ''' <summary>
    ''' Execute a query, that returns no data. 
    ''' </summary>
    ''' <param name="sql">The SQL statement.</param>
    ''' <returns>Integer on success</returns>
    ''' <remarks></remarks>
    Function executeQuery(sql As String) As Integer

    ''' <summary>
    ''' Insert data to the database. 
    ''' </summary>
    ''' <param name="table">The table to be affected by the query.</param>
    ''' <param name="data">The data to be inserted, column and value.</param>
    ''' <returns>Boolean indicating success.</returns>
    ''' <remarks></remarks>
    Function insert(table As String, data As Dictionary(Of String, String)) As Boolean

    ''' <summary>
    ''' Update data in the database.
    ''' </summary>
    ''' <param name="table">The table to be affected by the query.</param>
    ''' <param name="data">The data to be updated, column and value.</param>
    ''' <returns>Boolean indicating success.</returns>
    ''' <remarks></remarks>
    Function update(table As String, data As Dictionary(Of String, String)) As Boolean

    ''' <summary>
    ''' Delete data from the database.
    ''' </summary>
    ''' <param name="table">The table to be affected by the query.</param>
    ''' <param name="data">The data to be deleted, column and value.</param>
    ''' <returns>Boolena indicating success.</returns>
    ''' <remarks></remarks>
    Function delete(table As String, data As Dictionary(Of String, String)) As Boolean

    ''' <summary>
    ''' Select data from the database, returning a datatable with the results. 
    ''' </summary>
    ''' <param name="sql">The SELECT query.</param>
    ''' <returns>DataTable with the results.</returns>
    ''' <remarks></remarks>
    Function getDataTable(ByVal sql As String) As DataTable

End Interface
