Imports System.Data.SqlClient
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel

Public Class CWarehouseReturn
    Inherits CUtility

    Public Function SaveWHReturnHead(WarehouseRef As String, SWarehouseRef As String, Reference As String, TotalItems As Integer, TransactionDate As Date) As Boolean
        Using conn As New SqlConnection(GetConnString(1))
            Dim inscmd As New SqlCommand
            With inscmd
                .Connection = conn
                .Connection.Open()
                .CommandType = CommandType.Text
                .CommandText = "INSERT INTO tblWReturns (WarehouseRef,SWarehouseRef,Reference,TotalItems,TransactionDate,CreatedBy,CreatedDate) VALUES (@WarehouseRef,@SWarehouseRef,@Reference,@TotalItems,@TransactionDate,@CreatedBy,@CreatedDate)"
                With .Parameters
                    .AddWithValue("@WarehouseRef", WarehouseRef)
                    .AddWithValue("@SWarehouseRef", SWarehouseRef)
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
    Public Function SaveWhReturnLines(WReturnID As Integer, StockCode As String, Qty As Integer, Value As Decimal) As Boolean
        Using conn As New SqlConnection(GetConnString(1))
            Dim inscmd As New SqlCommand
            With inscmd
                .Connection = conn
                .Connection.Open()
                .CommandType = CommandType.Text
                .CommandText = "INSERT INTO tblWReturnLines (WReturnID,StockCode,Qty,Value) VALUES (@WReturnID,@StockCode,@Qty,@Value)"
                With .Parameters
                    .AddWithValue("@WReturnID", CInt(WReturnID))
                    .AddWithValue("@StockCode", StockCode)
                    .AddWithValue("@Qty", CInt(Qty))
                    .AddWithValue("@Value", CDec(Value))
                End With
                .ExecuteNonQuery()
            End With
        End Using
        Return True
    End Function
End Class
