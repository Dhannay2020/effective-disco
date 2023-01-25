Imports System.Data.SqlClient

Public Class CShopDelivery
    Inherits CUtility
    Public Function SaveShopDelHead(ShopRef As String, ShopName As String, WarehouseRef As String, WarehousesName As String, Reference As String, TotalItems As Integer, DeliveryDate As Date) As Boolean
        Using conn As New SqlConnection(GetConnString(1))
            Dim inscmd As New SqlCommand
            With inscmd
                .Connection = conn
                .Connection.Open()
                .CommandType = CommandType.Text
                .CommandText = "INSERT INTO tblShopDeliveries(ShopRef,ShopName,WarehouseRef,WarehouseName,Reference,TotalItems,DeliveryDate,DeliveryType,ConfirmedDate,Notes,CreatedBy,CreatedDate) VALUES (@ShopRef,@ShopName,@WarehouseRef,@WarehouseName,@Reference,@TotalItems,@DeliveryDate,@DeliveryType,@ConfirmedDate,@Notes,@CreatedBy,@CreatedDate)"
                With .Parameters
                    .AddWithValue("@ShopRef", ShopRef)
                    .AddWithValue("@ShopName", ShopName)
                    .AddWithValue("@WarehouseRef", WarehouseRef)
                    .AddWithValue("@WarehouseName", WarehousesName)
                    .AddWithValue("@Reference", Reference)
                    .AddWithValue("@TotalItems", CInt(TotalItems))
                    .AddWithValue("@DeliveryDate", CDate(DeliveryDate))
                    .AddWithValue("@DeliveryType", "Confirmed")
                    .AddWithValue("@ConfirmedDate", CDate(DeliveryDate))
                    .AddWithValue("@Notes", "")
                    .AddWithValue("@CreatedBy", username)
                    .AddWithValue("@CreatedDate", Date.Now)
                End With
                .ExecuteNonQuery()
            End With
        End Using
        Return True
    End Function
    Public Function SaveShopDelLines(SDeliveriesID As Integer, SStockCode As String, DeliveredQty As Integer) As Boolean
        Using conn As New SqlConnection(GetConnString(1))
            Dim inscmd As New SqlCommand
            With inscmd
                .Connection = conn
                .Connection.Open()
                .CommandType = CommandType.Text
                .CommandText = "INSERT INTO tblShopDeliveriesLines (SDeliveriesID,SStockCode,DeliveredQty) VALUES (@SDeliveriesID,@SStockCode,@DeliveredQty)"
                With .Parameters
                    .AddWithValue("@SDeliveriesID", CInt(SDeliveriesID))
                    .AddWithValue("@SStockCode", SStockCode)
                    .AddWithValue("@DeliveredQty", CInt(DeliveredQty))
                End With
                .ExecuteNonQuery()
            End With
        End Using
        Return True
    End Function
    Public Function SaveShopDeliveryHead() As Integer

        Return Result
    End Function
    Public Function SaveShopDeliveryLine() As Boolean

        Return True
    End Function
    Public Function UpdateShopDeliveryHead() As Boolean

        Return True
    End Function
    Public Function UpdateShopDeliveryLine() As Boolean

        Return True
    End Function
    Public Function DeleteShopDelivery() As Boolean

        Return True
    End Function
End Class
