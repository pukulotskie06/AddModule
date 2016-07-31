Public Class DatabaseClass

    Private con As New OleDb.OleDbConnection
    Private dataAdapter As OleDb.OleDbDataAdapter
    Public dataSet As New DataSet
    Private dataProvider As String = "Provider=Microsoft.ACE.OLEDB.12.0;"
    Private dataSource As String = "Data Source = " & System.Environment.CurrentDirectory & "\Inventory.accdb"
    'Provider=Microsoft.ACE.OLEDB.12.0;Data Source="C:\Users\pukulot\Documents\Visual Studio 2010\Projects\POS\POS\bin\Debug\inventory.accdb";Persist Security Info=True;User ID=admin

    Public Sub DatabaseClass()
        con.Close()

        con.ConnectionString = dataProvider & dataSource
        con.Open()
        dataAdapter = New OleDb.OleDbDataAdapter("Select * From Items Where Qty >= 3", con)
        dataAdapter.Fill(dataSet, "Inventory")



        ''MsgBox(dataSource)

    End Sub


    'Add function for the
    Public Sub Add(ByVal items() As String)

        con.ConnectionString = dataProvider & dataSource
        con.Open()

        Dim sqlCmd As String = "INSERT INTO Items ([Product_Name],[Description],[Brand],[Type],[Size],[Qty],[Unit],[Price]) Values (?,?,?,?,?,?,?,?)"

        Dim cmd As New OleDb.OleDbCommand(sqlCmd, con)

        cmd.Parameters.Add(New OleDb.OleDbParameter("Product_Name", items(0)))
        cmd.Parameters.Add(New OleDb.OleDbParameter("Description", items(1)))
        cmd.Parameters.Add(New OleDb.OleDbParameter("Brand", items(2)))
        cmd.Parameters.Add(New OleDb.OleDbParameter("Type", items(3)))
        cmd.Parameters.Add(New OleDb.OleDbParameter("Size", items(4)))
        cmd.Parameters.Add(New OleDb.OleDbParameter("Qty", items(5)))
        cmd.Parameters.Add(New OleDb.OleDbParameter("Unit", items(6)))
        cmd.Parameters.Add(New OleDb.OleDbParameter("Price", items(7)))

        Try
            cmd.ExecuteNonQuery()
            cmd.Dispose()
        Catch e As Exception
            MsgBox(e.ToString)
        End Try

        MsgBox("You have successfully added a new item")

    End Sub
    ''end of add function


End Class
