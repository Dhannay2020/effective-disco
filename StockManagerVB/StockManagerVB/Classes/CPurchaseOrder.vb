Imports System.Data.SqlClient
Public Class CPurchaseOrder
    Inherits CUtility

    Public Function SavePOrdersHead(notes As String, invoice As String, shipper As String, shipperinv As String) As Integer

        Using conn As New SqlConnection(GetConnString(1))
            Dim insCmd As New SqlCommand
            With insCmd
                .Connection = conn
                .CommandText = "INSERT INTO tblPurchaseOrder (OurRef, SupplierRef, LocationRef,  TotalGarments, TotalBoxes, TotalHangers, NetAmount, DeliveryCharge, Commission, TotalAmount, DeliveryDate, DeliveryType, ConfirmedDate, Notes, InvoiceNo, Shipper, ShipperInvoice, CreatedBy, CreatedDate) VALUES (@OurRef, @SupplierRef, @LocationRef, @TotalGarments, @TotalBoxes, @TotalHangers, @NetAmount, @DeliveryCharge, @Commission, @TotalAmount, @DeliveryDate, @DeliveryType, @ConfirmedDate, @Notes, @InvoiceNo, @Shipper, @ShipperInvoice, @CreatedBy, @CreatedDate)"
                .CommandType = CommandType.Text
                With .Parameters
                    .AddWithValue("@OurRef", StockCode)
                    .AddWithValue("@SupplierRef", SupplierRef)
                    .AddWithValue("@LocationRef", WarehouseRef)
                    .AddWithValue("@TotalGarments", TotalGarments)
                    .AddWithValue("@TotalBoxes", TotalBoxes)
                    .AddWithValue("@TotalHangers", TotalHangers)
                    .AddWithValue("@NetAmount", NetAmount)
                    .AddWithValue("@DeliveryCharge", DeliveryCharge)
                    .AddWithValue("@Commission", Commission)
                    .AddWithValue("@TotalAmount", TotalAmount)
                    .AddWithValue("@DeliveryDate", NewDate)
                    .AddWithValue("@DeliveryType", "Confirmed")
                    .AddWithValue("@ConfirmedDate", NewDate)
                    .AddWithValue("@Notes", notes)
                    .AddWithValue("@InvoiceNo", invoice)
                    .AddWithValue("@Shipper", shipper)
                    .AddWithValue("@ShipperInvoice", shipperinv)
                    .AddWithValue("@CreatedBy", GetUserName())
                    .AddWithValue("@CreatedDate", Now)
                End With
                .Connection.Open()
                Result = .ExecuteNonQuery()
            End With
        End Using
        Return Result
    End Function
    Public Function SavePOrdersLine() As Boolean
        Using conn As New SqlConnection(GetConnString(1))
            Dim insCmd As New SqlCommand
            With insCmd
                .Connection = conn
                .CommandText = "INSERT INTO tblPurchaseOrderLine (DeliveryID, StockCode, DQtyGarments, DQtyBoxes, DQtyHangers, NetAmount) VALUES (@DeliveryID, @StockCode, @DQtyGarments, @DQtyBoxes, @DQtyHangers, @NetAmount)"
                .CommandType = CommandType.Text
                With .Parameters
                    .AddWithValue("@DeliveryID", DeliveryID)
                    .AddWithValue("@StockCode", StockCode)
                    .AddWithValue("@DQtyGarments", DeliveredQtyGarments)
                    .AddWithValue("@DQtyBoxes", DeliveredQtyBoxes)
                    .AddWithValue("@DQtyHangers", DeliveredQtyHangers)
                    .AddWithValue("@NetAmount", LineAmount)
                End With
                .Connection.Open()
                .ExecuteNonQuery()
            End With
        End Using
        Return True
    End Function
    Public Function GetDataGridViewDBCmdString() As String
        Return ""
    End Function
End Class
