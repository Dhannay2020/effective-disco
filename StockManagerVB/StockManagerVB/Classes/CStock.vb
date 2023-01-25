Imports System.Data.SqlClient
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel

Public Class CStock
    Inherits CUtility
    Public Function SaveStock(stockcode As String, supplierref As String, costvalue As Decimal, totalG As Integer) As Boolean
        Using conn As New SqlConnection(GetConnString(1))
            Dim insertCmd As New SqlCommand
            With insertCmd
                .Connection = conn
                .CommandText = " INSERT INTO tblStock (StockCode,SupplierRef,DeadCode,DeliveredQtyHangers,AmountTaken,CostValue,PCMarkUp,ZeroQty,CreatedBy,CreatedDate,Season) VALUES (@StockCode,@SupplierRef,@DeadCode,@DeliveredQtyHangers,@AmountTaken,@CostValue,@PCMarkUp,@ZeroQty,@CreatedBy,@CreatedDate,@Season)"
                .CommandType = CommandType.Text
                With .Parameters
                    .AddWithValue("@ZeroQty", "0")
                    .AddWithValue("@StockCode", stockcode)
                    .AddWithValue("@SupplierRef", supplierref)
                    .AddWithValue("@DeadCode", "0")
                    .AddWithValue("@AmountTaken", "0.00")
                    .AddWithValue("@CostValue", CDec(costvalue))
                    .AddWithValue("@PCMarkUp", "0")
                    .AddWithValue("@Season", "ALL")
                    .AddWithValue("@CreatedBy", username)
                    .AddWithValue("@CreatedDate", Date.Now)
                    .AddWithValue("@DeliveredQtyHangers", CInt(totalG))
                End With
                .Connection.Open()
                .ExecuteNonQuery()
            End With
        End Using
        Return True
    End Function
    Public Function SaveStockCode() As Boolean
        Return True
    End Function
    Public Function UpdateStockCode() As Boolean
        Return True
    End Function
    Public Function DeleteStockCode() As Boolean
        Return True
    End Function
    Public Function GetDataGridViewDBCmdString() As String
        Return ""
    End Function
End Class
