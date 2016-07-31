Public Class Add

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub
    Dim dbs As New DatabaseClass

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim items() As String =
            {txtProdName.Text, txtDescription.Text, txtBrand.Text, txtType.Text,
             txtSize.Text, txtQty.Text, txtUnit.Text, txtPrice.Text}

        If (FillIt(items) = True) Then
            dbs.Add(items)
        End If

    End Sub

    Private Sub priceQty(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtQty.KeyPress, txtPrice.KeyPress

        If (Asc(e.KeyChar) >= 48 And Asc(e.KeyChar) <= 57 Or Asc(e.KeyChar) = 8) Then
            e.Handled = False

        Else
            e.Handled = True
        End If

        'txtPrice.Text = FormatNumber(txtPrice.Text, 2, , , TriState.True)
    End Sub


    Private Sub txtPrice_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtPrice.KeyUp
        If (e.KeyCode = 190) Then
            If txtPrice.Text.IndexOf(".") = -1 Then
                txtPrice.Text &= "."
            End If

        End If

    End Sub

    Private Function FillIt(ByRef items() As String)

        If items(0) = "" Then
            MsgBox("Required field!!! Fill in the Product_Name!!")
            txtProdName.Focus()
            Return False
        End If

        If items(1) = "" Then
            items(1) = "N/A"
        End If

        If items(2) = "" Then
            items(2) = "N/A"
        End If

        If items(3) = "" Then
            items(3) = "N/A"
        End If

        If items(4) = "" Then
            items(4) = "N/A"
        End If

        If items(6) = "" Then
            items(6) = "N/A"
        End If

        Return True
    End Function

    Private Sub btnClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClear.Click

        txtPrice.Clear()
        txtBrand.Clear()
        txtQty.Clear()
        txtUnit.Clear()
        txtDescription.Clear()
        txtType.Clear()
        txtSize.Clear()
        txtProdName.Clear()

    End Sub
End Class
