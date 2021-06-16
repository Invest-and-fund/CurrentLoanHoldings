'####
'this extract is designed to run on the first of the month to pick up all current loan holding positions as at the end of the previous month. 
'Obviously it could be run on any date for the day previous.
'In order to make it work for a point in the past there are a few code changes required. these have been highlighted with #### 
'####


Imports System.Configuration
Imports System.Collections.Specialized
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient



Public Class Form1
    Public Class Extract
        Property AccountID As Integer
        Property Firstname As String
        Property LastName As String
        Property CompanyName As String
        Property LoanID As Integer
        Property LHBalID As Integer
        Property LHID As Integer
        Property OrderID As Integer
        Property Outstanding As Decimal
        Property Loanstatus As Integer
        Property DateOfAcquisition As Date
        Property DateOfLoan As Date
        Property DateOfRepayment As Date
        Property TotalFacilityAmount As Decimal
    End Class

    Public Class Summary
        Property Fullname As String
        Property LHAmount As Decimal
    End Class

    Public Shared ReadOnly Property FBSQLEnv As String = System.Configuration.ConfigurationManager.AppSettings("RunFBSQL")
    Public Shared ReadOnly Property UpdateFBSQLEnv As String = System.Configuration.ConfigurationManager.AppSettings("UpdateFBSQL")

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim dateentered As String = Today.ToString("yyyy/MM/01")
        ExtRunDate.Text = dateentered
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ' run sql to set up table for extract

        '##### if running against a specific date instead of running against the current date then change this to the date of the next day (i.e. pick up everything less than this date) ####
        Dim dExtractDate As Date = #2019-05-01#

        Dim environ As String = "L"
        RunningMsg.Text = "The extract is running. This can take a while. Please be patient"
        RunningMsg.MaximumSize = New Size(200, 0)
        RunningMsg.AutoSize = True
        Me.Refresh()

        Dim dateentered As Date = ExtRunDate.Text
        Dim DisplayDate As String = ExtRunDate.Text

        Dim nLoanStatus As Int16 = 0
        Dim dLoanEndDate As Date
        Dim nLoanid As Integer



        DisplayDate = DisplayDate.Replace("/", "-")

        dExtractDate = dateentered
        Dim ExtractList As New List(Of Extract)
        Dim SummaryList As New List(Of Summary)


        Dim sSQL As String
            Dim sErrorStr As String = ""
            Using con As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("SQLConnectionString").ConnectionString)

                con.Open()

                Dim command As SqlCommand = con.CreateCommand()
                command.Connection = con
                command.CommandType = CommandType.Text
                
                Try
                sSQL = "CREATE OR ALTER VIEW LH_POS_BALANCES(
                   ACCOUNTID,
                    LH_ID,
                    LH_BALS_ID,
                    NUM_UNITS)
                    AS
                    SELECT   t.ACCOUNTID, t.LH_ID, vt.maxlhbalsid AS LH_BALS_ID, t.NUM_UNITS
                        FROM      
                    (SELECT   MAX(LH_BALS_ID) AS maxlhbalsid, LHID_ACCID, LH_ID, ACCOUNTID
                       FROM      dbo.LH_BALS AS l
                         WHERE   (LH_ID >= 0) AND (DATECREATED < '" & DisplayDate & "')
                      GROUP BY LHID_ACCID, LH_ID, ACCOUNTID) AS vt INNER JOIN
                     dbo.LH_BALS AS t ON t.LH_BALS_ID = vt.maxlhbalsid
                    WHERE   (t.NUM_UNITS > 0) ; "
                command.CommandText = sSQL
                    command.ExecuteNonQuery()

                Catch ex As Exception

                End Try

                Try
                    sSQL = "CREATE OR ALTER VIEW LH_POS_BALANCES_SUSPENSE(
                   ACCOUNTID,
                    LH_ID,
                    LH_BALS_ID,
                    NUM_UNITS)
                    AS
                     SELECT   vt.ACCOUNTID, t.LH_ID, vt.max_lh_bals_sus_id AS LH_BALS_ID, t.NUM_UNITS
                       FROM     
                 (SELECT   ACCOUNTID, MAX(LH_BALS_SUSPENSE_ID) AS max_lh_bals_sus_id
                    FROM      dbo.LH_BALS_SUSPENSE AS s
                     WHERE   (LH_ID > 0) AND (DATECREATED < '" & DisplayDate & "')
                         GROUP BY ACCOUNTID, LH_ID) AS vt INNER JOIN
                     dbo.LH_BALS_SUSPENSE AS t ON t.LH_BALS_SUSPENSE_ID = vt.max_lh_bals_sus_id
                    WHERE   (t.NUM_UNITS > 0) ;"
                command.CommandText = sSQL
                    command.ExecuteNonQuery()
                Catch ex As Exception

                End Try
                con.Close()
                Dim Rdr, Rdr2, Rdr3 As SqlDataReader

                Try
                    con.Open()
                    sSQL = "  SELECT
                                lhb.accountid as TheAccountID,
                                Trim(u.firstname) as TheFirstName,
                                Trim(u.lastname) as TheLastName,
                                Trim(l.business_name) AS TheCompanyName,
                                l.loanid AS TheLoan,
                                lhb.lh_bals_id as Thelhbalsid,
                                o.lh_id as Thelhid,
                                0 as TheOrder,
                                COALESCE(Cast(lhb.num_units AS FLOAT)/100, 0)
                                    + COALESCE(Cast(lhS.num_units AS FLOAT)/100, 0) AS Theoutstanding,
                                l.loanstatus,
                                l.ipo_end as TheDateOfLoan,
                                l.DATE_OF_LAST_PAYMENT as TheDateOfRepayment,
                                l.facility_amount as TheTotalFacilityAmount

                            FROM orders o,
                                loans l,
                                loan_holdings lh,
                                lh_pos_balances lhb

                            INNER JOIN accounts a on a.accountid = lhb.accountid
                            inner join users u on u.userid = a.userid

                            FULL JOIN lh_pos_balances_suspense lhs
                                    ON lhb.accountid = lhs.accountid
                                   AND lhb.lh_id = lhs.lh_id
                                   
                                 WHERE lh.loan_holdings_id = lhb.lh_id
                                   AND l.loanid = lh.loanid
                                   AND o.lh_id = lh.loan_holdings_id 
                    
                                   AND l.loanstatus in (2, 4)
                                   AND a.accountid = lhb.accountid
                                   AND u.userid = a.userid
                                   AND u.usertype = 0
                                   AND l.loanid not in (54, 58, 60, 89, 293)

                                GROUP BY
                                    lhb.accountid,
                                    u.firstname,
                                    u.lastname,
                                    l.business_name,
                                    l.loanid,
                                    lhb.lh_bals_id,
                                    o.lh_id,
                                    lhb.num_units,
									lhS.num_units,
                                    l.loanstatus,
                                    l.ipo_end,
                                    l.DATE_OF_LAST_PAYMENT,
                                    l.facility_amount"

                    Dim cmd As SqlCommand = New SqlCommand(sSQL, con)
                    cmd.Parameters.Clear()



                    Rdr = cmd.ExecuteReader


                    Do While Rdr.Read()
                        Dim newExtract As New Extract
                        newExtract.AccountID = IIf(IsDBNull(Rdr(0)), "", Rdr(0))
                        newExtract.Firstname = IIf(IsDBNull(Rdr(1)), "", Rdr(1))
                        newExtract.LastName = IIf(IsDBNull(Rdr(2)), "", Rdr(2))
                        newExtract.CompanyName = IIf(IsDBNull(Rdr(3)), "", Rdr(3))
                        newExtract.LoanID = IIf(IsDBNull(Rdr(4)), "", Rdr(4))
                        nLoanid = IIf(IsDBNull(Rdr(4)), "", Rdr(4))
                        newExtract.LHBalID = IIf(IsDBNull(Rdr(5)), "", Rdr(5))
                        newExtract.LHID = IIf(IsDBNull(Rdr(6)), "", Rdr(6))
                        newExtract.OrderID = IIf(IsDBNull(Rdr(7)), "", Rdr(7))
                        newExtract.Outstanding = IIf(IsDBNull(Rdr(8)), 0, Rdr(8))
                        newExtract.Loanstatus = IIf(IsDBNull(Rdr(9)), "", Rdr(9))
                        nLoanStatus = IIf(IsDBNull(Rdr(9)), "", Rdr(9))
                        newExtract.DateOfAcquisition = "#01/01/0001#"
                        newExtract.DateOfLoan = IIf(IsDBNull(Rdr(10)), "#01/01/0001#", Rdr(10))
                        newExtract.DateOfRepayment = IIf(IsDBNull(Rdr(11)), "#01/01/0001#", Rdr(11))
                        dLoanEndDate = IIf(IsDBNull(Rdr(11)), "#01/01/0001#", Rdr(11))
                        newExtract.TotalFacilityAmount = IIf(IsDBNull(Rdr(12)), 0, Rdr(12) / 100)

                        Dim TheEnd As String = "break here"

                        If nLoanid = 370 Then
                            TheEnd = "370"
                        End If

                        If nLoanStatus = 4 Then

                            If dLoanEndDate < dExtractDate Then
                            Else
                                ExtractList.Add(newExtract)
                            End If
                        Else
                            ExtractList.Add(newExtract)
                        End If

                    Loop


                    Rdr.Close()
                    'con.Close()
                    Rdr = Nothing
                    Rdr2 = Nothing
                    Rdr3 = Nothing
                    'thi is where the date of acquisition is added to the table

                    'Dim TransList As New List(Of Trans)
                    Dim nAccountID As Integer
                    Dim nOrderID As Integer
                    Dim nOrderPrev As Integer
                    Dim nLHID As Integer


                    Dim sAccountID As Integer = 0
                    Dim sOrderID As Integer = 0


                    Dim xPrevAccount As String = ""
                    Dim xLatestAccount As String

                    Dim newSummary As Summary
                    Dim nAccountTotal As Decimal = 0

                    sAccountID = 0
                    sOrderID = 0


                    For Each Extract In ExtractList
                        nAccountID = Extract.AccountID
                        nLHID = Extract.LHID
                        nOrderID = Extract.OrderID
                        nLoanid = Extract.LoanID
                        nOrderPrev = 0

                        sSQL = "Select o.lh_id, o.accountid, o.loanid, o.orderprev, o.orderid
                                From orders  o
                                WHERE  o.accountid = @iaccountid
                                And (o.lh_id = @ilhid
                                 or o.loanid = @iloanid)
                                Order By o.lh_id desc, o.loanid desc,  o.orderprev desc, o.accountid  "

                        'con.Open()
                        cmd = New SqlCommand(sSQL, con)
                        cmd.Parameters.Clear()
                        With cmd.Parameters
                            .Add(New SqlParameter("@iaccountid", nAccountID))
                            .Add(New SqlParameter("@ilhid", nLHID))
                            .Add(New SqlParameter("@iloanid", nLoanid))
                        End With


                        Rdr2 = cmd.ExecuteReader

                        If Rdr2.Read() Then
                            Extract.OrderID = IIf(IsDBNull(Rdr2(4)), "", Rdr2(4))
                            nOrderID = Extract.OrderID
                            nOrderPrev = IIf(IsDBNull(Rdr2(3)), 0, Rdr2(3))
                        End If
                        Rdr2.Close()
                        'con.Close()

                        sSQL = "SELECT
                                ft.datecreated
                              FROM
                                fin_trans ft
                              WHERE ft.transtype in (1206, 1401)
                                AND ft.accountid = @iaccountid
                                AND (ft.orderid = @iorderid
                                Or ft.orderid = @iorderprev)"

                        'con.Open()
                        cmd = New SqlCommand(sSQL, con)
                        cmd.Parameters.Clear()
                        With cmd.Parameters
                            .Add(New SqlParameter("@iaccountid", nAccountID))
                            .Add(New SqlParameter("@iorderid", nOrderID))
                            .Add(New SqlParameter("@iorderprev", nOrderPrev))
                        End With


                        Rdr3 = cmd.ExecuteReader
                        If Rdr3.Read() Then
                            Extract.DateOfAcquisition = IIf(IsDBNull(Rdr3(0)), "", Rdr3(0))

                        End If
                        Rdr3.Close()
                        'con.Close()

                        xLatestAccount = Extract.AccountID & " - " & Extract.Firstname.Trim & " " & Extract.LastName.Trim
                        nAccountID = Extract.AccountID
                        nOrderID = Extract.OrderID
                        '      Reader = Nothing
                        sAccountID = 0
                        sOrderID = 0

                        If xLatestAccount <> xPrevAccount Then

                            If xPrevAccount <> "" Then
                                newSummary = New Summary With {
                                    .Fullname = xPrevAccount,
                                    .LHAmount = nAccountTotal}

                                SummaryList.Add(newSummary)
                                nAccountTotal = 0
                            End If
                        End If
                        nAccountTotal = Extract.Outstanding + nAccountTotal
                        xPrevAccount = xLatestAccount
                    Next
                    'once we have read and processed the final record we need to write it away
                    newSummary = New Summary With {
                                    .Fullname = xPrevAccount,
                                    .LHAmount = nAccountTotal}

                    SummaryList.Add(newSummary)
                    nAccountTotal = 0
                Catch ex As Exception
                    sErrorStr = sErrorStr & vbNewLine & ex.Message
                Finally

                End Try

                con.Close()

            End Using


        DataGridView1.DataSource = ExtractList
        DataGridView2.DataSource = SummaryList

        TextBox2.Text = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        TextBox1.Text = "csvloanholdings"

        RunningMsg.Text = ""

    End Sub


    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim xFileName As String
        Dim xFolderPath As String

        xFileName = TextBox1.Text
        xFolderPath = TextBox2.Text
        If xFileName = "" Then
            xFileName = "csvloanholdings"
        End If
        If xFolderPath = "" Then
            xFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        End If

        'export the files as csv using the variable name
        'Build the CSV file data as a Comma separated string.
        Dim csv As String = String.Empty

        'Add the Header row for CSV file.
        For Each column As DataGridViewColumn In DataGridView1.Columns
            csv += column.HeaderText & ","c
        Next

        'Add new line.
        csv += vbCr & vbLf

        'Adding the Rows
        For Each row As DataGridViewRow In DataGridView1.Rows
            For Each cell As DataGridViewCell In row.Cells
                'Add the Data rows.
                csv += cell.Value.ToString().Replace(",", ";") & ","c
            Next

            'Add new line.
            csv += vbCr & vbLf
        Next

        'Exporting to Excel

        Dim xFilePath As String
        xFilePath = xFolderPath & "\" & xFileName & "extract.csv"
        File.WriteAllText(xFilePath, csv)
        MessageBox.Show("CSV file written to " & xFilePath)


        'now write the second file
        'Add the Header row for CSV file.
        csv = String.Empty

        For Each column As DataGridViewColumn In DataGridView2.Columns
            csv += column.HeaderText & ","c
        Next

        'Add new line.
        csv += vbCr & vbLf

        'Adding the Rows
        For Each row As DataGridViewRow In DataGridView2.Rows
            For Each cell As DataGridViewCell In row.Cells
                'Add the Data rows.
                csv += cell.Value.ToString().Replace(",", ";") & ","c
            Next

            'Add new line.
            csv += vbCr & vbLf
        Next

        'Exporting to Excel

        xFilePath = xFolderPath & "\" & xFileName & "summary.csv"
        File.WriteAllText(xFilePath, csv)
        MessageBox.Show("CSV file written to " & xFilePath)
    End Sub

    Private Sub Label5_Click(sender As Object, e As EventArgs) Handles Label5.Click

    End Sub
End Class
