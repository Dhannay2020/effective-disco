Imports System.Data.SqlClient
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel

Public Class CShopSale
    Inherits CUtility
    Public Function SaveShopSalesHead(shopref As String, shopname As String, TransactionDate As Date, TotalQty As Integer, TotalValue As Decimal, Vatvalue As Decimal) As Boolean
        Try
            Using conn As New SqlConnection(GetConnString(1))
                Dim inscmd As New SqlCommand
                With inscmd
                    .Connection = conn
                    .Connection.Open()
                    .CommandType = CommandType.Text
                    .CommandText = "INSERT INTO tblSales (ShopRef,ShopName,SAReference,TransactionDate,TotalQty,TotalValue,CreatedBy,CreatedDate,VATAmount) VALUES (@ShopRef,@ShopName,@SAReference,@TransactionDate,@TotalQty,@TotalValue,@CreatedBy,@CreatedDate,@VATAmount)"
                    With .Parameters
                        .AddWithValue("@ShopRef", shopref)
                        .AddWithValue("@ShopName", shopname)
                        .AddWithValue("@SAReference", "0")
                        .AddWithValue("@TransactionDate", TransactionDate)
                        .AddWithValue("@TotalQty", TotalQty)
                        .AddWithValue("@TotalValue", TotalValue)
                        .AddWithValue("@CreatedBy", username)
                        .AddWithValue("@CreatedDate", Date.Now)
                        .AddWithValue("@VATAmount", CDec(Vatvalue))
                    End With
                    .ExecuteNonQuery()
                End With
            End Using
            Return True
        Catch ex As SqlException
            MsgBox("Record Creation Failed because of" & vbCrLf & ex.ErrorCode & "  " & ex.Message, MsgBoxStyle.Information, Application.ProductName)
        End Try
        Return True
    End Function
    Public Function SaveShopSalesLines(salesID As Integer, stockcode As String, CurrentQty As Integer, QtySold As Integer, SalesAmount As Decimal, StockID As Integer) As Boolean
        Using conn As New SqlConnection(GetConnString(1))
            Dim inscmd As New SqlCommand
            With inscmd
                .Connection = conn
                .Connection.Open()
                .CommandType = CommandType.Text
                .CommandText = "INSERT INTO tblSalesLines (SalesID,StockCode,CurrentQty,QtySold,SalesAmount,StockMovementID) VALUES (@SalesID,@StockCode,@CurrentQty,@QtySold,@SalesAmount,@StockMovementID)"
                With .Parameters
                    .AddWithValue("@SalesID", salesID)
                    .AddWithValue("@StockCode", stockcode)
                    .AddWithValue("@CurrentQty", CurrentQty)
                    .AddWithValue("@QtySold", QtySold)
                    .AddWithValue("@SalesAmount", SalesAmount)
                    .AddWithValue("@StockMovementID", StockID)
                End With
                .ExecuteNonQuery()
            End With
        End Using
        Return True
    End Function

    Public Function deleteShopSalesLines(salesID As Integer) As Boolean
        Using conn As New SqlConnection(GetConnString(1))
            Dim inscmd As New SqlCommand
            With inscmd
                .Connection = conn
                .Connection.Open()
                .CommandType = CommandType.Text
                .CommandText = "DELETE from tblSalesLines WHERE SalesID = @SalesID AND QtySold = @QtySold AND SalesAmount = @SalesAmount;"
                With .Parameters
                    .AddWithValue("@SalesID", salesID)
                    .AddWithValue("@QtySold", "0")
                    .AddWithValue("@SalesAmount", "0.00")

                End With
                .ExecuteNonQuery()
            End With
        End Using
        Return True

    End Function
End Class
