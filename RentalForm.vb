
Option Explicit On
Option Strict On
Option Compare Binary
Public Class RentalForm
    Dim valueCheck As Boolean
    Dim totalMiles As Integer
    Dim totalCustomers As Integer
    Dim totalCharge As Decimal
    Function Validate() As Boolean
        Dim zipValue, beginOdometer, endOdometer, checkDay As Integer

        If NameTextBox.Text = String.Empty Then
            ActiveControl = NameTextBox
            MsgBox("All Values must be entered before Proceeding.")
            Exit Function

        ElseIf ZipCodeTextBox.Text = String.Empty Then
            ActiveControl = ZipCodeTextBox
            MsgBox("All Values must be entered before Proceeding.")
            Exit Function

        ElseIf CityTextBox.Text = String.Empty Then
            ActiveControl = CityTextBox
            MsgBox("All Values must be entered before Proceeding.")
            Exit Function

        ElseIf AddressTextBox.Text = String.Empty Then
            ActiveControl = AddressTextBox
            MsgBox("All Values must be entered before Proceeding.")
            Exit Function

        ElseIf StateTextBox.Text = String.Empty Then
            ActiveControl = StateTextBox
            MsgBox("All Values must be entered before Proceeding.")
            Exit Function

        ElseIf BeginOdometerTextBox.Text = String.Empty Then
            ActiveControl = BeginOdometerTextBox
            MsgBox("All Values must be entered before Proceeding.")
            Exit Function

        ElseIf EndOdometerTextBox.Text = String.Empty Then
            ActiveControl = EndOdometerTextBox
            MsgBox("All Values must be entered before Proceeding.")
            Exit Function

        ElseIf DaysTextBox.Text = String.Empty Then
            ActiveControl = DaysTextBox
            MsgBox("All Values must be entered before Proceeding.")
            Exit Function
        End If

        Try
            zipValue = CInt(ZipCodeTextBox.Text)
        Catch
            ActiveControl = ZipCodeTextBox
            MsgBox("Zipcode value must be a whole number")
            Exit Function
        End Try

        Try
            beginOdometer = CInt(BeginOdometerTextBox.Text)
        Catch
            ActiveControl = BeginOdometerTextBox
            BeginOdometerTextBox.Text = String.Empty
            MsgBox("Whole number must be used for Odometer reading")
            Exit Function
        End Try

        Try
            endOdometer = CInt(EndOdometerTextBox.Text)
        Catch
            ActiveControl = EndOdometerTextBox
            EndOdometerTextBox.Text = String.Empty
            MsgBox("Whole number must be used for Odometer reading")
            Exit Function
        End Try

        Try
            checkDay = CInt(DaysTextBox.Text)
        Catch
            ActiveControl = DaysTextBox
            DaysTextBox.Text = String.Empty
            MsgBox("Number of days must be a whole number")
            Exit Function
        End Try

        If beginOdometer > endOdometer Then
            ActiveControl = BeginOdometerTextBox
            BeginOdometerTextBox.Text = String.Empty
            EndOdometerTextBox.Text = String.Empty
            MsgBox("Beginning Odometer Reading can't be higher than end odometer reading")
            Exit Function
        End If

        If checkDay > 45 Or checkDay <= 0 Then
            ActiveControl = DaysTextBox
            DaysTextBox.Text = String.Empty
            MsgBox("Renting limit is 45 days")
            Exit Function
        End If
        valueCheck = True
        Return valueCheck
    End Function
    Private Sub CalculateButton_Click(sender As Object, e As EventArgs) Handles CalculateButton.Click, CalculateToolStripMenuItem.Click, ContextMenuCalculate.Click
        valueCheck = False
        valueCheck = Validate()
        If valueCheck = False Then
            Exit Sub
        End If
        totalCustomers += 1
        'allows value to add to itself 
        Calculate()
    End Sub
    Sub Calculate()
        Dim days As Integer
        Dim chargeAmount As Integer
        Dim milesCharge, totalMiles, discount, roundAmount As Decimal

        days = CInt(DaysTextBox.Text)
        chargeAmount = days * 15
        'DayChargeTextBox.Text = "$" & CStr(chargeAmount) & ".00"
        DayChargeTextBox.Text = $"${CStr(chargeAmount)}.00"
        'accomplishes same outcome where "$" indicates the start of a string where "{" denotes an "&"
        'and "cstring" denotes converting the value in parantheses into a string 

        If KilometersradioButton.Checked = True Then
            totalMiles = CInt((0.62 * CInt(EndOdometerTextBox.Text)) - (0.62 * CInt(BeginOdometerTextBox.Text)))
            TotalMilesTextBox.Text = CStr(CInt(EndOdometerTextBox.Text) - CInt(BeginOdometerTextBox.Text)) & " Miles"
        End If

        If MilesradioButton.Checked = True Then
            totalMiles = CInt(EndOdometerTextBox.Text) - CInt(BeginOdometerTextBox.Text)
            TotalMilesTextBox.Text = CStr(CInt(EndOdometerTextBox.Text) - CInt(BeginOdometerTextBox.Text)) & " Miles"
        End If

        If totalMiles > 500 Then
            milesCharge += 300 * 0.12D
            milesCharge += (totalMiles - 500) * 0.1D
            '+= allows the value given to add to "milesCharge" itself
        ElseIf totalMiles < 501 And totalMiles > 201 Then
            milesCharge = (totalMiles - 200) * 0.12D
        End If
        MileageChargeTextBox.Text = $"${CStr(milesCharge)}"
        'MileageChargeTextBox.Text = "$" & CStr(milesCharge)

        If AAAcheckbox.Checked = True And Seniorcheckbox.Checked = True Then
            discount = ((chargeAmount + milesCharge) * 0.03D) + ((chargeAmount + milesCharge) * 0.05D)
        ElseIf AAAcheckbox.Checked = True And Seniorcheckbox.Checked = False Then
            discount = (chargeAmount + milesCharge) * 0.05D
        ElseIf Seniorcheckbox.Checked = True And AAAcheckbox.Checked = False Then
            discount = (chargeAmount + milesCharge) * 0.03D
        End If

        roundAmount = Math.Round(discount, 2, MidpointRounding.AwayFromZero)
        TotalDiscountTextBox.Text = "$" & CStr(roundAmount)
        TotalChargeTextBox.Text = CStr((chargeAmount + milesCharge) - roundAmount)
        totalMiles += CInt(totalMiles)
        totalCharge += (chargeAmount + milesCharge) - roundAmount
        SummaryButton.Enabled = True
    End Sub
    Private Sub ClearButton_Click(sender As Object, e As EventArgs) Handles ClearButton.Click, ClearToolStripMenuItem1.Click, ContextMenuClear.Click
        valueClear()
    End Sub
    'Clearbutton_click summons the clear sub. This is to allow other functions to call upon this sub 
    'such as the summaryButton_Click. This makes it so I dont have to type all these values in again.
    Sub valueClear()
        BeginOdometerTextBox.Text = String.Empty
        EndOdometerTextBox.Text = String.Empty
        MileageChargeTextBox.Text = String.Empty
        TotalMilesTextBox.Text = String.Empty
        DaysTextBox.Text = String.Empty
        DayChargeTextBox.Text = String.Empty
        NameTextBox.Text = String.Empty
        AddressTextBox.Text = String.Empty
        CityTextBox.Text = String.Empty
        StateTextBox.Text = String.Empty
        ZipCodeTextBox.Text = String.Empty
        TotalDiscountTextBox.Text = String.Empty
        TotalChargeTextBox.Text = String.Empty
        AAAcheckbox.Checked = False
        Seniorcheckbox.Checked = False
        MilesradioButton.Checked = True
        KilometersradioButton.Checked = False
    End Sub
    Private Sub ExitButton_Click(sender As Object, e As EventArgs) Handles ExitButton.Click, ExitToolStripMenuItem1.Click, ContextMenuExit.Click
        Dim areuSure As MsgBoxResult
        areuSure = MsgBox("Would you like to exit?", MsgBoxStyle.YesNo)
        If areuSure = vbYes Then
            Me.Close()
        End If
    End Sub

    Private Sub SummaryButton_Click(sender As Object, e As EventArgs) Handles SummaryButton.Click, SummaryToolStripMenuItem1.Click, ContextMenuSummary.Click
        MsgBox("Number of Customers: " & totalCustomers & vbNewLine &
               "Total Charges: $" & totalCharge &
               "Total Miles Driven: " & totalMiles & vbNewLine)
        valueClear()
    End Sub

    Private Sub RentalForm_click(sender As Object, e As MouseEventArgs) Handles Me.MouseClick
        '"e" is dimed as mouse event argument where mouse event arguments contain mouse info like: position, click, button pressed ..etc.
        If e.Button = MouseButtons.Right Then
            'If e.Button.ToString = "right" Then 
            'if mouse button clicked is the right mouse button.
            ContextMenuStrip.Show()
            ContextMenuStrip.Location = MousePosition
            'allows the context menu strip to appear where the position of the mouse was when right mouse button was clicked

        End If

    End Sub
End Class


