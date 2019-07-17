Public Class baseform

    Private Sub baseform_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    End Sub

    'Forming a collection of textboxes, to convinently print months in the calendar

    Dim all As New Collection                   ' all is collection of textboxes that stores month
    Private Sub a1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        all.Add(a1)
        all.Add(a2)
        all.Add(a3)
        all.Add(a4)
        all.Add(a5)
        all.Add(a6)
        all.Add(a7)
        all.Add(b1)
        all.Add(b2)
        all.Add(b3)
        all.Add(b4)
        all.Add(b5)
        all.Add(b6)
        all.Add(b7)
        all.Add(c1)
        all.Add(c2)
        all.Add(c3)
        all.Add(c4)
        all.Add(c5)
        all.Add(c6)
        all.Add(c7)
    End Sub

    Dim boxes As New Collection                 ' boxes is collection of text boxes that stores days
    Private Sub box1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        boxes.Add(box1)
        boxes.Add(box2)
        boxes.Add(box3)
        boxes.Add(box4)
        boxes.Add(box5)
        boxes.Add(box6)
        boxes.Add(box7)
        boxes.Add(box8)
        boxes.Add(box9)
        boxes.Add(box10)
        boxes.Add(box11)
        boxes.Add(box12)
        boxes.Add(box13)
        boxes.Add(box14)
        boxes.Add(box15)
        boxes.Add(box16)
        boxes.Add(box17)
        boxes.Add(box18)
        boxes.Add(box19)
        boxes.Add(box20)
        boxes.Add(box21)
        boxes.Add(box22)
        boxes.Add(box23)
        boxes.Add(box24)
        boxes.Add(box25)
        boxes.Add(box26)
        boxes.Add(box27)
        boxes.Add(box28)
        boxes.Add(box29)
        boxes.Add(box30)
        boxes.Add(box31)
        boxes.Add(box32)
        boxes.Add(box33)
        boxes.Add(box34)
        boxes.Add(box35)
        boxes.Add(box36)
        boxes.Add(box37)
        boxes.Add(box38)
        boxes.Add(box39)
        boxes.Add(box40)
        boxes.Add(box41)
        boxes.Add(box42)
        boxes.Add(box43)
        boxes.Add(box44)
        boxes.Add(box45)
        boxes.Add(box46)
        boxes.Add(box47)
        boxes.Add(box48)
        boxes.Add(box49)

    End Sub

    'Describing the exception handling and functionality of the GO button
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles gobutton.Click
        '  If input string contains anything except 0-9 or is more than 4 characters long then generate error
        Try
            If year.Text = "" Then
                MessageBox.Show("Can't leave input year as empty", "Null input")
                For i As Integer = 1 To 49
                    boxes(i).BackColor = Color.LightGoldenrodYellow
                Next
            Else

                year.Text = year.Text.Trim(" ")                                                  'First of all trimming any leading 
                Dim temp As String = year.Text.Remove(0, 1)                                      'or trailing white spaces and  
                If Not temp.StartsWith("+") Then                                                 'leading 0's in the input string
                    year.Text = year.Text.TrimStart("+")
                End If
                year.Text = year.Text.TrimStart("0")
                ' Checking various instances when the input string is not a valid year from 0001 to 9999 
                If year.Text.Length > 4 Or year.Text.Contains(",") Or year.Text.Contains(".") Or year.Text.Contains("(") Or year.Text.Contains("-") Or year.Text.Contains(" ") Or Not IsNumeric(year.Text) Or year.Text.Contains("+") Then
                    year.Text = year.Text.Remove(year.Text.Length - 1, 1)
                    ' Resetting the textboxes displaying months as empty
                    For i As Integer = 1 To 21
                        all(i).ResetText()
                    Next
                    ' Since the year is not valid, therefore disabling the monthdrop and datedrop comboboxes
                    monthdrop.Enabled = False
                    datedrop.Enabled = False
                    ' Unhighlighting the highlighted cell   
                    For i As Integer = 1 To 49
                        boxes(i).BackColor = Color.LightGoldenrodYellow
                    Next

                    ' We could show the following message in case of error if needed.
                    ' But in case of our program, since the input is year, which is a well known entity that it is a number, and for all 
                    ' practical uses, it is less than 9999, therefore wrong input is generally due to typing error.
                    ' Hence, displaying error box is not of any use as the user would most likely know his/her mistake.

                    '  MessageBox.Show("Please enter year as a whole number without any special symbol (or sign) between 0001 to 9999", "Invalid Input")
                End If
                Dim entry As Integer = CInt(year.Text)
                '  If valid input is there, then clear the text which may be there from earlier result 
                For i As Integer = 1 To 21
                    all(i).ResetText()
                Next
                '  Print the new result
                printcalendar(entry)
                '  Since the year has been selected, user can now enter the month.
                monthdrop.Enabled = True
                ComboBox2_SelectedIndexChanged(sender, e)

            End If
        Catch ex As Exception
        End Try

    End Sub

    '  Zeller's Algorithm to find day for a given date (source:Wikipedia- https://en.wikipedia.org/wiki/Zeller%27s_congruence)
    Private Function algo(year As Integer, month As Integer, inputdate As Integer)
        '  Special case for Jan and Feb
        If month = 1 Or month = 2 Then
            month = month + 12
            year = year - 1
        End If

        Dim h As Integer = (inputdate + year + Math.Floor(year / 4) + Math.Floor(year / 400) - Math.Floor(year / 100) + Math.Floor(13 * (month + 1) / 5)) Mod 7
        Dim d As Integer = ((h + 5) Mod 7) + 1
        Return d

    End Function

    '  Function returning the name of month from month number as input
    Private Function getMonth(month As Integer)
        Select Case month
            Case 1
                Return "Jan"
            Case 2
                Return "Feb"
            Case 3
                Return "Mar"
            Case 4
                Return "Apr"
            Case 5
                Return "May"
            Case 6
                Return "June"
            Case 7
                Return "July"
            Case 8
                Return "Aug"
            Case 9
                Return "Sep"
            Case 10
                Return "Oct"
            Case 11
                Return "Nov"
            Case 12
                Return "Dec"
        End Select
    End Function

    '  Function to display the required calendar
    Private Sub printcalendar(year As Integer)

        Dim counter() As Integer = {0, 0, 0, 0, 0, 0, 0}                          ' counter(<index>-1) is used to check which row we are in presently, for column number <index> 
        Dim dayofweek As Integer

        For i As Integer = 1 To 12                                                ' Calculating for the first of each month is sufficient as it covers all the columns of days
            dayofweek = algo(year, i, 1)
            all((counter(dayofweek - 1)) * 7 + dayofweek).Text = getMonth(i)      ' Printing the calculated month in the desired cell
            counter(dayofweek - 1) = counter(dayofweek - 1) + 1                   ' Incrementing the counter for that column by 1, so that next time a month of that column is printed in the next row
        Next

    End Sub

    '  Various Sub routines for all the textboxes and table layout used for the display purpose
    Private Sub TableLayoutPanel1_Paint(sender As Object, e As PaintEventArgs) Handles layouttable.Paint

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub TextBox32_TextChanged(sender As Object, e As EventArgs) Handles TextBox32.TextChanged

    End Sub

    Private Sub TextBox31_TextChanged(sender As Object, e As EventArgs) Handles TextBox31.TextChanged

    End Sub

    Private Sub TextBox30_TextChanged(sender As Object, e As EventArgs) Handles TextBox30.TextChanged

    End Sub

    Private Sub TextBox29_TextChanged(sender As Object, e As EventArgs) Handles TextBox29.TextChanged

    End Sub

    Private Sub TextBox28_TextChanged(sender As Object, e As EventArgs) Handles TextBox28.TextChanged

    End Sub

    Private Sub TextBox27_TextChanged(sender As Object, e As EventArgs) Handles TextBox27.TextChanged

    End Sub

    Private Sub TextBox26_TextChanged(sender As Object, e As EventArgs) Handles TextBox26.TextChanged

    End Sub

    Private Sub TextBox25_TextChanged(sender As Object, e As EventArgs) Handles TextBox25.TextChanged

    End Sub

    Private Sub TextBox24_TextChanged(sender As Object, e As EventArgs) Handles TextBox24.TextChanged

    End Sub

    Private Sub TextBox23_TextChanged(sender As Object, e As EventArgs) Handles TextBox23.TextChanged

    End Sub

    Private Sub TextBox22_TextChanged(sender As Object, e As EventArgs) Handles TextBox22.TextChanged

    End Sub

    Private Sub TextBox21_TextChanged(sender As Object, e As EventArgs) Handles TextBox21.TextChanged

    End Sub

    Private Sub TextBox20_TextChanged(sender As Object, e As EventArgs) Handles TextBox20.TextChanged

    End Sub

    Private Sub TextBox19_TextChanged(sender As Object, e As EventArgs) Handles TextBox19.TextChanged

    End Sub

    Private Sub TextBox18_TextChanged(sender As Object, e As EventArgs) Handles TextBox18.TextChanged

    End Sub

    Private Sub TextBox17_TextChanged(sender As Object, e As EventArgs) Handles TextBox17.TextChanged

    End Sub

    Private Sub TextBox16_TextChanged(sender As Object, e As EventArgs) Handles TextBox16.TextChanged

    End Sub

    Private Sub TextBox15_TextChanged(sender As Object, e As EventArgs) Handles TextBox15.TextChanged

    End Sub

    Private Sub TextBox14_TextChanged(sender As Object, e As EventArgs) Handles TextBox14.TextChanged

    End Sub

    Private Sub TextBox13_TextChanged(sender As Object, e As EventArgs) Handles TextBox13.TextChanged

    End Sub

    Private Sub TextBox12_TextChanged(sender As Object, e As EventArgs) Handles TextBox12.TextChanged

    End Sub

    Private Sub TextBox11_TextChanged(sender As Object, e As EventArgs) Handles TextBox11.TextChanged

    End Sub

    Private Sub TextBox10_TextChanged(sender As Object, e As EventArgs) Handles TextBox10.TextChanged

    End Sub

    Private Sub TextBox9_TextChanged(sender As Object, e As EventArgs) Handles TextBox9.TextChanged

    End Sub

    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged

    End Sub

    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged

    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged

    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged

    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged

    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged

    End Sub

    Private Sub TextBox61_TextChanged(sender As Object, e As EventArgs) Handles box22.TextChanged

    End Sub

    Private Sub TextBox41_TextChanged(sender As Object, e As EventArgs) Handles box16.TextChanged

    End Sub

    Private Sub TextBox58_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub TextBox39_TextChanged(sender As Object, e As EventArgs) Handles box9.TextChanged

    End Sub

    Private Sub TextBox83_TextChanged(sender As Object, e As EventArgs) Handles TextBox83.TextChanged

    End Sub


    Private Sub year_TextChanged(sender As Object, e As EventArgs) Handles year.TextChanged
        ' If year starts with 0 then trim it first 
        year.Text = year.Text.TrimStart("0")

        ' If input is not empty then call Button1_Click function 
        If year.Text <> "" Then
            Button1_Click(sender, e)
        End If
        ' Now if input is not empty, it means it is valid as checked in the Button1_Click function
        If year.Text <> "" Then
            '  Now if some month is selected then enable the monthdrop combobox and call combobox2_selectedindexchanged function
            If monthdrop.Text <> "" Then
                monthdrop.Enabled = True
                ComboBox2_SelectedIndexChanged(sender, e)
            End If
            ' If input is empty, disable the monthdrop and datedrop combobox
        Else
            monthdrop.Enabled = False
            datedrop.Enabled = False
        End If

    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs)
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles monthdrop.SelectedIndexChanged
        '  Since month has been selected then datedrop is enabled so that user can select date
        datedrop.Enabled = True
        If monthdrop.Text <> "" And datedrop.Text <> "" Then
            ' If both date and month are selected from previous instance, then check (for feb) whether the date selected is valid or not 
            ' If not, then clear that date
            If CInt(datedrop.Text) > NoOfDays(CInt(year.Text), CInt(monthdrop.Text)) Then
                datedrop.Text = ""
            End If
        End If

        ' Retaining what date had been entered by the user, in case the user wants to see that date for multiple years
        ' It would otherwise get erased when clear is called later in the program
        Dim temp As Integer = -1
        If datedrop.Text <> "" Then
            temp = CInt(datedrop.Text)
        End If

        ' Clearing the contents of the datedrop combobox so that we can recalculate and add the correct number of days according to the chosen month
        datedrop.Items.Clear()
        Dim enter As String = monthdrop.Text

        ' If month has not been selected, then disable datedrop combobox
        If enter = "" Then
            datedrop.Enabled = False
        End If

        ' The items in the datedrop combobox are modified according to the number of days in the month selected
        For i As Integer = 1 To NoOfDays(CInt(year.Text), enter)
            datedrop.Items.Add(i)
        Next
        If temp <> -1 Then
            datedrop.Text = temp
        End If

        ' Calling the highlighter function to highlight the selected date
        highlighter()
    End Sub

    Private Sub datedrop_SelectedIndexChanged(sender As Object, e As EventArgs) Handles datedrop.SelectedIndexChanged
        ' Calling the highlighter function to highlight the selected date
        highlighter()
    End Sub

    ' Auxiliary function to return no of days in a given month taking year and month as input 
    Private Function NoOfDays(year As Integer, month As Integer)
        Select Case month
            Case 1, 3, 5, 7, 8, 10, 12
                Return 31
            Case 4, 6, 9, 11
                Return 30
            Case 2
                If year Mod 400 = 0 Then
                    Return 29
                Else
                    If year Mod 4 = 0 And year Mod 100 <> 0 Then
                        Return 29
                    Else
                        Return 28
                    End If
                End If
        End Select
    End Function

    ' Function to highlight the selected date
    Private Function highlighter()
        ' First refreshing the screen to remove the earlier highlighted date
        For i As Integer = 1 To 49
            boxes(i).BackColor = Color.LightGoldenrodYellow
        Next

        ' Only highlight if year, month and date are all valid and not empty 
        If monthdrop.Text <> "" And datedrop.Text <> "" And year.Text <> "" Then
            Dim month As Integer = CInt(monthdrop.Text)
            Dim inputdate As Integer = CInt(datedrop.Text)
            Dim years As Integer = CInt(year.Text)
            Dim d As Integer = algo(years, month, inputdate)
            ' Calculating which box to highlight
            boxes(((inputdate - 1) Mod 7) * 7 + 1 + (7 - ((inputdate) Mod 7) + d) Mod 7).BackColor = Color.GreenYellow    ' Changing the colour of the calculated textbox
        End If

    End Function

    ' Auxiliary function to calculate day of week by taking d as input
    Private Function dayofweek(d As Integer)
        Dim day As String
        Select Case d
            Case 1
                day = "Mon"
            Case 2
                day = "Tue"
            Case 3
                day = "Wed"
            Case 4
                day = "Thu"
            Case 5
                day = "Fri"
            Case 6
                day = "Sat"
            Case 7
                day = "Sun"
        End Select
        Return day
    End Function

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub TableLayoutPanel1_Paint_1(sender As Object, e As PaintEventArgs) Handles TableLayoutPanel1.Paint

    End Sub

    Private Sub c5_TextChanged(sender As Object, e As EventArgs) Handles c5.TextChanged

    End Sub

    Private Sub a3_TextChanged(sender As Object, e As EventArgs) Handles a3.TextChanged

    End Sub

    Private Sub a2_TextChanged(sender As Object, e As EventArgs) Handles a2.TextChanged

    End Sub

    Private Sub b2_TextChanged(sender As Object, e As EventArgs) Handles b2.TextChanged

    End Sub

    Private Sub c2_TextChanged(sender As Object, e As EventArgs) Handles c2.TextChanged

    End Sub

    Private Sub c1_TextChanged(sender As Object, e As EventArgs) Handles c1.TextChanged

    End Sub

    Private Sub b1_TextChanged(sender As Object, e As EventArgs) Handles b1.TextChanged

    End Sub

    Private Sub a1_TextChanged(sender As Object, e As EventArgs) Handles a1.TextChanged

    End Sub

    Private Sub TextBox80_TextChanged(sender As Object, e As EventArgs) Handles TextBox80.TextChanged

    End Sub

    Private Sub box45_TextChanged(sender As Object, e As EventArgs) Handles box45.TextChanged

    End Sub

    Private Sub box38_TextChanged(sender As Object, e As EventArgs) Handles box38.TextChanged

    End Sub

    Private Sub box31_TextChanged(sender As Object, e As EventArgs) Handles box31.TextChanged

    End Sub

    Private Sub box24_TextChanged(sender As Object, e As EventArgs) Handles box24.TextChanged

    End Sub

    Private Sub box17_TextChanged(sender As Object, e As EventArgs) Handles box17.TextChanged

    End Sub

    Private Sub box10_TextChanged(sender As Object, e As EventArgs) Handles box10.TextChanged

    End Sub

    Private Sub box3_TextChanged(sender As Object, e As EventArgs) Handles box3.TextChanged

    End Sub

    Private Sub box2_TextChanged(sender As Object, e As EventArgs) Handles box2.TextChanged

    End Sub

    Private Sub box23_TextChanged(sender As Object, e As EventArgs) Handles box23.TextChanged

    End Sub

    Private Sub box30_TextChanged(sender As Object, e As EventArgs) Handles box30.TextChanged

    End Sub

    Private Sub box37_TextChanged(sender As Object, e As EventArgs) Handles box37.TextChanged

    End Sub

    Private Sub box44_TextChanged(sender As Object, e As EventArgs) Handles box44.TextChanged

    End Sub

    Private Sub box43_TextChanged(sender As Object, e As EventArgs) Handles box43.TextChanged

    End Sub

    Private Sub box36_TextChanged(sender As Object, e As EventArgs) Handles box36.TextChanged

    End Sub

    Private Sub box29_TextChanged(sender As Object, e As EventArgs) Handles box29.TextChanged

    End Sub

    Private Sub box15_TextChanged(sender As Object, e As EventArgs) Handles box15.TextChanged

    End Sub

    Private Sub box8_TextChanged(sender As Object, e As EventArgs) Handles box8.TextChanged

    End Sub

    Private Sub box1_TextChanged(sender As Object, e As EventArgs) Handles box1.TextChanged

    End Sub
End Class

