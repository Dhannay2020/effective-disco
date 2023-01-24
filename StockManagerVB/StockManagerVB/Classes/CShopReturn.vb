Imports System.Data.SqlClient
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel

Public Class CShopReturn
    Inherits CUtility
    Public Function SaveShopReturnsHead(ShopRef As String, WarehouseRef As String, Reference As String, TotalItems As Integer, TransactionDate As Date) As Boolean
        Using conn As New SqlConnection(GetConnString(1))
            Dim inscmd As New SqlCommand
            With inscmd
                .Connection = conn
                .Connection.Open()
                .CommandType = CommandType.Text
                .CommandText = "INSERT INTO tblReturns(ShopRef,WarehouseRef,Reference,TotalItems,TransactionDate,CreatedBy,CreatedDate) VALUES (@ShopRef,@WarehouseRef,@Reference,@TotalItems,@TransactionDate,@CreatedBy,@CreatedDate)"
                With .Parameters
                    .AddWithValue("@ShopRef", ShopRef)
                    .AddWithValue("@WarehouseRef", WarehouseRef)
                    .AddWithValue("@Reference", Reference)
                    .AddWithValue("@TotalItems", CInt(TotalItems))
                    .AddWithValue("@TransactionDate", CDate(TransactionDate))
                    .AddWithValue("@CreatedBy", username)
                    .AddWithValue("@CreatedDate", Date.Now)
                End With
                .ExecuteNonQuery()
            End With
        End Using
        Return True
    End Function
    Public Function SaveShopReturnsLines(ReturnID As Integer, StockCode As String, Qty As Integer, Value As Decimal) As Boolean
        Using conn As New SqlConnection(GetConnString(1))
            Dim inscmd As New SqlCommand
            With inscmd
                .Connection = conn
                .Connection.Open()
                .CommandType = CommandType.Text
                .CommandText = "INSERT INTO tblReturnLines (ReturnID,StockCode,Qty,Value) VALUES (@ReturnID,@StockCode,@Qty,@Value)"
                With .Parameters
                    .AddWithValue("@ReturnID", ReturnID)
                    .AddWithValue("@StockCode", StockCode)
                    .AddWithValue("@Qty", Qty)
                    .AddWithValue("@Value", Value)
                End With
                .ExecuteNonQuery()
            End With
        End Using
        Return True
    End Function
End Class
