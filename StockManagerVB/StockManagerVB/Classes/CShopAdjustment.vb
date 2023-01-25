Imports System.Data.SqlClient
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel

Public Class CShopAdjustment
    Inherits CUtility
    Public Function SaveSHAdjustHead(ShopRef As String, Reference As String, TotalLossItems As Integer, TotalGainItems As Integer, MovementDate As Date) As Boolean
        Using conn As New SqlConnection(GetConnString(1))
            Dim inscmd As New SqlCommand
            With inscmd
                .Connection = conn
                .Connection.Open()
                .CommandType = CommandType.Text
                .CommandText = "INSERT INTO tblShopAdjustments(ShopRef,Reference,TotalLossItems,TotalGainItems,MovementDate,CreatedBy,CreatedDate) VALUES (@ShopRef,@Reference,@TotalLossItems,@TotalGainItems,@MovementDate,@CreatedBy,@CreatedDate)"
                With .Parameters
                    .AddWithValue("@ShopRef", ShopRef)
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
    Public Function SaveSHAdjustLines(ShopAdjustID As Integer, StockCode As String, MovementType As String, qty As Integer, Value As Decimal) As Boolean
        Using conn As New SqlConnection(GetConnString(1))
            Dim inscmd As New SqlCommand
            With inscmd
                .Connection = conn
                .Connection.Open()
                .CommandType = CommandType.Text
                .CommandText = "INSERT INTO tblShopAdjustmentsLines(ShopAdjustID,StockCode,MovementType,Qty,Value) VALUES (@ShopAdjustID,@StockCode,@MovementType,@Qty,@Value)"
                With .Parameters
                    .AddWithValue("@ShopAdjustID", CInt(ShopAdjustID))
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
    Public Function SaveShopAdjustmentHead() As Integer

        Return Result
    End Function
    Public Function SaveShopAdjustmentLine() As Boolean

        Return True
    End Function
    Public Function UpdateShopAdjustmentHead() As Boolean

        Return True
    End Function
    Public Function UpdateShopAdjustmentLine() As Boolean

        Return True
    End Function
    Public Function DeleteShopAdjustment() As Boolean

        Return True
    End Function
    Public Function GetDataGridViewDBCmdString() As String
        Return ""
    End Function
End Class
