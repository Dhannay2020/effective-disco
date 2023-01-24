Public Class CUtility
    Public WarehouseRef, ShopRef, SupplierRef, StockCode As String
    Public AddressLine1, AddressLine2, AddressLine3, AddressLine4, PostCode, Telephone, Fax, ContactName, EMail, WebsiteAddress, Memo As String


    Public Function GetConnString(ByVal Id As Integer) As String
        Dim StringConn As String
        If Id = 1 Then
            StringConn = "Initial Catalog=DMHStockv2;Data Source=.\SQLEXPRESS;Persist Security Info=False;Integrated Security=true;"
        Else
            StringConn = "Initial Catalog=master;Data Source=.\SQLEXPRESS;Persist Security Info=False;Integrated Security=true;"
        End If
        Return StringConn
    End Function
    Public Function GetDateRequired(ByRef Dte As Date, ByRef DateType As Integer) As Date
        Dim DateToReturn As Date
        If DateType = 1 Then
            DateToReturn = Dte.AddDays(0 - Dte.DayOfWeek)
        Else
            DateToReturn = Dte.AddDays(0 - Dte.DayOfWeek + 7)
        End If
        Return DateToReturn
    End Function
    Public Function GetUserName() As String
        Return "Admin"
    End Function
End Class
