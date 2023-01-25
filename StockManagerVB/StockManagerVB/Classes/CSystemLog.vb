Imports System.Data.SqlClient

Public Class CSystemLog
    Inherits CUtility
    Public Sub SaveSystemLog(stockcode As String, supplierref As String, location As String, qty As Integer, movementtype As String, recordtype As String, mdate As Date, reference As String)
        Using conn As New SqlConnection(GetConnString(1))
            Using InsertCmd As New SqlCommand()
                With InsertCmd
                    .Connection = conn
                    .CommandType = CommandType.Text
                    .CommandText = "INSERT INTO tblSystemLog (StockCode, SupplierRef, Location, Qty, MovementType, RecordType, MovementDate, Timestamp, Reference, CreatedBy) VALUES (@StockCode, @SupplierRef, @Location, @Qty, @MovementType, @RecordType, @MovementDate, @Timestamp, @Reference, @CreatedBy)"
                    .Connection.Open()
                    .Parameters.AddWithValue("@StockCode", stockcode)
                    .Parameters.AddWithValue("@SupplierRef", supplierref)
                    .Parameters.AddWithValue("@Location", location)
                    .Parameters.AddWithValue("@Qty", qty)
                    .Parameters.AddWithValue("@MovementType", movementtype)
                    .Parameters.AddWithValue("@RecordType", recordtype)
                    .Parameters.AddWithValue("@MovementDate", mdate)
                    .Parameters.AddWithValue("@Timestamp", Date.Now)
                    .Parameters.AddWithValue("@Reference", reference)
                    .Parameters.AddWithValue("@CreatedBy", GetUserName)
                    .ExecuteNonQuery()
                End With
            End Using
        End Using
    End Sub
    Public Sub SaveStockMovements(stockcode As String, supplierref As String, location As String, locationtype As String, hqty As Integer, bqty As Integer, movementtype As String, movedate As Date, value As Decimal, IReference As String, TRef As Integer)
        Try
            Using conn As New SqlConnection(GetConnString(1))
                Dim insertcmd As New SqlCommand
                With insertcmd
                    .CommandText = "INSERT INTO tblStockMovements(StockCode, SupplierRef, LocationRef, LocationType, MovementQtyHangers, MovementType, MovementDate, MovementValue, Reference, TransferReference, CreatedBy, CreatedDate) VALUES(@StockCode, @SupplierRef, @LocationRef, @LocationType, @MovementQtyHangers, @MovementType, @MovementDate, @MovementValue, @Reference, @TransferReference, @CreatedBy, @CreatedDate)"
                    .CommandType = CommandType.Text
                    .Connection = conn
                    .Parameters.AddWithValue("@StockCode", stockcode)
                    .Parameters.AddWithValue("@SupplierRef", supplierref)
                    .Parameters.AddWithValue("@LocationRef", location)
                    .Parameters.AddWithValue("@LocationType", CInt(locationtype))
                    .Parameters.AddWithValue("@MovementQtyHangers", CInt(hqty))
                    .Parameters.AddWithValue("@MovementType", CInt(movementtype))
                    .Parameters.AddWithValue("@MovementDate", CDate(movedate))
                    .Parameters.AddWithValue("@MovementValue", CDec(value))
                    .Parameters.AddWithValue("@Reference", IReference)
                    .Parameters.AddWithValue("@TransferReference", CInt(TRef))
                    .Parameters.AddWithValue("@CreatedBy", GetUserName())
                    .Parameters.AddWithValue("@CreatedDate", Date.Now)
                    .Connection.Open()
                    .ExecuteNonQuery()
                End With
            End Using
        Catch ex As SqlException
            MsgBox("Record Creation Failed because of" & vbCrLf & ex.ErrorCode & "  " & ex.Message, MsgBoxStyle.Information, Application.ProductName)
        End Try
    End Sub
    Public Sub DeleteZeroStockMovements(movementtype As String, transferref As Integer, reference As String, dte As Date)
        Using conn As New SqlConnection(GetConnString(1))
            Dim deleteCmd As New SqlCommand
            With deleteCmd
                .Connection = conn
                .Connection.Open()
                .CommandType = CommandType.Text
                .CommandText = "DELETE FROM tblStockMovements WHERE movementtype = @MovementType AND TransferReference = @TransferReference AND Reference = @Reference AND MovementDate = @MovementDate AND MovementQtyHangers = '0' AND MovementValue = '0'"
                With .Parameters
                    .AddWithValue("@MovementType", movementtype)
                    .AddWithValue("@TransferReference", CInt(transferref))
                    .AddWithValue("@Reference", reference)
                    .AddWithValue("@MovementDate", dte)
                End With
                .ExecuteNonQuery()
            End With
        End Using
    End Sub
    Public Sub DeleteStockMovements(movementtype As String, transferref As Integer, reference As String, dte As Date)
        Using conn As New SqlConnection(GetConnString(1))
            Dim deleteCmd As New SqlCommand
            With deleteCmd
                .Connection = conn
                .Connection.Open()
                .CommandType = CommandType.Text
                .CommandText = "DELETE FROM tblStockMovements WHERE movementtype = @MovementType AND TransferReference = @TransferReference AND Reference = @Reference AND MovementDate = @MovementDate"
                With .Parameters
                    .AddWithValue("@MovementType", movementtype)
                    .AddWithValue("@TransferReference", CInt(transferref))
                    .AddWithValue("@Reference", reference)
                    .AddWithValue("@MovementDate", dte)
                End With
                .ExecuteNonQuery()
            End With
        End Using
    End Sub
End Class
