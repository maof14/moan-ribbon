Public Class CTools

    ' Class to encapsulate a few tools that is called from the Ribbon. 

    ' Function to convert a value to string in Excel. That is, prepending a apostrophe. 
    ' Return String value appended by a apostrophe. 
    Public Function convertToString(ByVal value) As String
        Return "'" & value
    End Function

    ' Function to get the current date in the format specified in the settings. 
    ' Return String the current date. 
    Public Function getDate(ByVal dateFormat As String) As String
        dateFormat = dateFormat.Replace("Y", "y").Replace("D", "d")
        Return Format(Now(), dateFormat)
    End Function

End Class
