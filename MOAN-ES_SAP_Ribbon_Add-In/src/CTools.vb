''' <summary>
''' A wrapper class to encapsulate the tools under the Tools section in the Ribbon. 
''' </summary>
''' <remarks></remarks>
Public Class CTools

    ''' <summary>
    ''' "Convert" a value to string - appending a apostrophe to the value. 
    ''' </summary>
    ''' <param name="value">The value to be converted.</param>
    ''' <returns>The value appended by a apostrophe.</returns>
    ''' <remarks></remarks>
    Public Function convertToString(ByVal value) As String
        Return "'" & value
    End Function

    ''' <summary>
    ''' Return the current date in the specified format. 
    ''' </summary>
    ''' <param name="dateFormat">The desired format.</param>
    ''' <returns>The current date in selected format.</returns>
    ''' <remarks></remarks>
    Public Function getDate(ByVal dateFormat As String) As String
        dateFormat = dateFormat.Replace("Y", "y").Replace("D", "d")
        Return Format(Now(), dateFormat)
    End Function

End Class
