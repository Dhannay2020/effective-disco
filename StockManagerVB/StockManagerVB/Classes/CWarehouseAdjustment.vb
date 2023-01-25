Imports System.Data.SqlClient
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel

Public Class CWarehouseAdjustment
    Inherits CUtility
    Public Function SaveWHAdjustHead(WarehouseRef As String, Reference As String, TotalLossItems As Integer, TotalGainItems As Integer, MovementDate As Date) As Boolean
        Using conn As New SqlConnection(GetConnString(1))
            Dim inscmd As New SqlCommand
            With inscmd
                .Connection = conn
                .Connection.Open()
                .CommandType = CommandType.Text
                .CommandText = "INSERT INTO tblWarehouseAdjustments(WarehouseRef,Reference,TotalLossItems,TotalGainItems,MovementDate,CreatedBy,CreatedDate) VALUES (@WarehouseRef,@Reference,@TotalLossItems,@TotalGainItems,@MovementDate,@CreatedBy,@CreatedDate)"
                With .Parameters
                    .AddWithValue("@WarehouseRef", WarehouseRef)
                    .AddWithValue("@Reference", Reference)
                    .AddWithValue("@TotalLossItems", CInt(TotalLossItems))
                    .AddWithValue("@TotalGainItems", CInt(TotalGainItems))
                    .AddWithValue("@MovementDate", CDate(MovementDate))
                    .AddWithValue("@CreatedBy", username)
                    .AddWithValue("@CreatedDate", Date.Now)
                End With
                .ExecuteNonQuery()
            End With
        End Using
        Return True
    End Function
    Public Function SaveWhAdjustLines(WarehouseAdjustID As Integer, StockCode As String, MovementType As String, qty As Integer, Value As Decimal) As Boolean
        Using conn As New SqlConnection(GetConnString(1))
            Dim inscmd As New SqlCommand
            With inscmd
                .Connection = conn
                .Connection.Open()
                .CommandType = CommandType.Text
                .CommandText = "INSERT INTO tblWarehouseAdjustmentsLines(WarehouseAdjustID,StockCode,MovementType,Qty,Value) VALUES (@WarehouseAdjustID,@StockCode,@MovementType,@Qty,@Value)"
                With .Parameters
                    .AddWithValue("@WarehouseAdjustID", CInt(WarehouseAdjustID))
                    .AddWithValue("@StockCode", StockCode)
                    .AddWithValue("@MovementType", MovementType)
                    .AddWithValue("@Qty", CInt(qty))
                    .AddWithValue("@Value", CDec(Value))
                End With
                .ExecuteNonQuery()
            End With
        End Using
        Return True
    End Function
    Public Function SaveWarehouseAdjustmentHead() As Integer

        Return Result
    End Function
    Public Function SaveWarehouseAdjustmentLine() As Boolean

        Return True
    End Function
    Public Function UpdateWarehouseAdjustmentHead() As Boolean

        Return True
    End Function
    Public Function UpdateWarehouseAdjustmentLine() As Boolean

        Return True
    End Function
    Public Function DeleteWarehouseAdjustment() As Boolean

        Return True
    End Function
    Public Function GetDataGridViewDBCmdString() As String
        Return ""
    End Function
End Class
