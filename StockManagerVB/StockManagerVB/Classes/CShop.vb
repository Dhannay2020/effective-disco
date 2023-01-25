Imports System.Data.SqlClient

Public Class CShop
    Inherits CUtility
    Public Function SaveShops(ShopRef As String, ShopName As String, Address1 As String, Address2 As String, Address3 As String, Address4 As String, PostCode As String, ContactName As String, Telephone As String, Telephone2 As String, Fax As String, eMail As String, Memo As String, ShopType As String, VATPayable As Boolean) As Boolean
        Using conn As New SqlConnection(GetConnString(1))
            Dim cmd As New SqlCommand
            With cmd
                .Connection = conn
                .Connection.Open()
                .CommandType = CommandType.Text
                .CommandText = "INSERT INTO tblShops (ShopRef, ShopName, ContactName, Street, Area, Town, County, Country, PostCode, Telephone, Telephone2, Fax, eMail, ShopType, Memo, CreatedBy, CreatedDate, VATPayable) VALUES (@ShopRef, @ShopName, @ContactName, @Street, @Area, @Town, @County, @Country, @PostCode, @Telephone, @Telephone2, @Fax, @eMail, @ShopType, @Memo, @CreatedBy, @CreatedDate, @VATPayable)"
                With .Parameters
                    .AddWithValue("@ShopRef", ShopRef)
                    .AddWithValue("@ShopName", ShopName)
                    .AddWithValue("@Street", Address1)
                    .AddWithValue("@Area", Address2)
                    .AddWithValue("@Town", Address3)
                    .AddWithValue("@County", Address4)
                    .AddWithValue("@Country", "UK")
                    .AddWithValue("@PostCode", PostCode)
                    .AddWithValue("@ContactName", ContactName)
                    .AddWithValue("@Telephone", Telephone)
                    .AddWithValue("@Telephone2", Telephone2)
                    .AddWithValue("@Fax", Fax)
                    .AddWithValue("@eMail", eMail)
                    .AddWithValue("@Memo", Memo)
                    .AddWithValue("@ShopType", ShopType)
                    .AddWithValue("@CreatedBy", GetUserName())
                    .AddWithValue("@CreatedDate", Date.Now)
                    .AddWithValue("@VATPayable", VATPayable)
                End With
                .ExecuteNonQuery()
            End With
        End Using
        Return True
    End Function
    Public Function SaveShop() As Boolean
        Return True
    End Function
    Public Function UpdateShop() As Boolean
        Return True
    End Function
    Public Function DeleteShop() As Boolean
        Return True
    End Function
    Public Function GetShopName() As String
        Return ""
    End Function
    Public Function GetDataGridViewDBCmdString() As String
        Return ""
    End Function
End Class
