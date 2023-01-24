Imports System.Data.SqlClient

Public Class CWarehouse
    Inherits CUtility
    Public Function SaveWarehouse(WarehouseRef As String, WarehouseName As String, WarehouseType As String) As Boolean
        Using conn As New SqlConnection(GetConnString(1))
            Dim cmd As New SqlCommand
            With cmd
                .Connection = conn
                .Connection.Open()
                .CommandType = CommandType.Text
                .CommandText = "INSERT INTO tblWarehouses (WarehouseRef, WarehouseName, Address1, Address2, Address3, Address4, PostCode, ContactName, Telephone, Telephone2, Fax, eMail, WarehouseType, Memo, CreatedBy, CreatedDate) VALUES (@WarehouseRef, @WarehouseName, @Address1, @Address2, @Address3, @Address4, @PostCode, @ContactName, @Telephone, @Telephone2, @Fax, @eMail, @WarehouseType, @Memo, @CreatedBy, @CreatedDate)"
                With .Parameters
                    .AddWithValue("@WarehouseRef", WarehouseRef)
                    .AddWithValue("@WarehouseName", WarehouseName)
                    .AddWithValue("@Address1", AddressLine1)
                    .AddWithValue("@Address2", AddressLine2)
                    .AddWithValue("@Address3", AddressLine3)
                    .AddWithValue("@Address4", AddressLine4)
                    .AddWithValue("@PostCode", PostCode)
                    .AddWithValue("@ContactName", ContactName)
                    .AddWithValue("@Telephone", Telephone)
                    .AddWithValue("@Telephone2", Telephone2)
                    .AddWithValue("@Fax", Fax)
                    .AddWithValue("@eMail", EMail)
                    .AddWithValue("@Memo", Memo)
                    .AddWithValue("@WarehouseType", WarehouseType)
                    .AddWithValue("@CreatedBy", GetUserName())
                    .AddWithValue("@CreatedDate", Date.Now)
                End With
                .ExecuteNonQuery()
            End With
        End Using
        Return True
    End Function
End Class
