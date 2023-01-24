Imports System.Data.SqlClient


Public Class CSupplier
    Inherits CUtility
    Public Function SaveSupplier(SupplierRef As String, SupplierName As String, Address1 As String, Address2 As String, Address3 As String, Address4 As String, PostCode As String, ContactName As String, Telephone As String, Telephone2 As String, Fax As String, eMail As String, Memo As String) As Boolean
        Using conn As New SqlConnection(GetConnString(1))
            Dim cmd As New SqlCommand
            With cmd
                .Connection = conn
                .Connection.Open()
                .CommandType = CommandType.Text
                .CommandText = "INSERT INTO tblSuppliers (SupplierRef, SupplierName, Address1, Address2, Address3, Address4, PostCode, ContactName,Telephone, Telephone2, Fax, eMail, Memo, CreatedBy, CreatedDate) VALUES (@SupplierRef, @SupplierName, @Address1, @Address2, @Address3, @Address4, @PostCode, @ContactName, @Telephone, @Telephone2, @Fax, @eMail, @Memo, @CreatedBy, @CreatedDate)"
                With .Parameters
                    .AddWithValue("@SupplierRef", SupplierRef)
                    .AddWithValue("@SupplierName", SupplierName)
                    .AddWithValue("@Address1", Address1)
                    .AddWithValue("@Address2", Address2)
                    .AddWithValue("@Address3", Address3)
                    .AddWithValue("@Address4", Address4)
                    .AddWithValue("@PostCode", PostCode)
                    .AddWithValue("@ContactName", ContactName)
                    .AddWithValue("@Telephone", Telephone)
                    .AddWithValue("@Telephone2", Telephone2)
                    .AddWithValue("@Fax", Fax)
                    .AddWithValue("@eMail", eMail)
                    .AddWithValue("@Memo", Memo)
                    .AddWithValue("@CreatedBy", GetUserName())
                    .AddWithValue("@CreatedDate", Date.Now)
                End With
                .ExecuteNonQuery()
            End With
        End Using
        Return True
    End Function
End Class
