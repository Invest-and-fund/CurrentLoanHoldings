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





        DisplayDate = DisplayDate.Replace("/", "-")

        dExtractDate = dateentered



        Dim connection As String = "FBConnectionString"


        Dim Cmd As New FirebirdSql.Data.FirebirdClient.FbCommand With {
            .Connection = New FirebirdSql.Data.FirebirdClient.FbConnection(ConfigurationManager.ConnectionStrings(environ & connection).ConnectionString),
            .CommandType = CommandType.Text
        }
        Cmd.CommandText = "CREATE OR ALTER VIEW LH_POS_BALANCES(
                ACCOUNTID,
               LH_ID,
            LH_BALS_ID,
            NUM_UNITS)
            AS
            select t.accountid, t.lh_id, maxlhbalsid, t.num_units
            from
            (
              select  max(l.lh_bals_id) as maxlhbalsid,l.lhid_accid, lh_id, accountid
                from lh_bals l
           where lh_id >= 0
            and l.datecreated < '" & DisplayDate & "'
            group by l.lhid_accid, lh_id, accountid
           order by lhid_accid
            ) vt
            inner join lh_bals t on t.lh_bals_id = vt.maxlhbalsid
            where t.num_units > 0
            ;"
        Cmd.Connection.Open()

        Cmd.ExecuteNonQuery()

        Cmd = New FirebirdSql.Data.FirebirdClient.FbCommand With {
            .Connection = New FirebirdSql.Data.FirebirdClient.FbConnection(ConfigurationManager.ConnectionStrings(environ & connection).ConnectionString),
            .CommandType = CommandType.Text
        }
        Cmd.CommandText = "CREATE OR ALTER VIEW LH_POS_BALANCES_SUSPENSE(
        ACCOUNTID,
        LH_ID,
        LH_BALS_ID,
        NUM_UNITS)
        AS
        select vt.accountid, lh_id, max_lh_bals_sus_id, t.num_units
        from 
        ( 
        select accountid, max(lh_bals_suspense_id) as max_lh_bals_sus_id
        from lh_bals_suspense s
        where lh_id > 0
        and s.datecreated < '" & DisplayDate & "'
        group by accountid, lh_id
        ) vt
        inner join lh_bals_suspense t on t.lh_bals_suspense_id = vt.max_lh_bals_sus_id
        where t.num_units > 0
        ;"
        Cmd.Connection.Open()

        Cmd.ExecuteNonQuery()

        Cmd.Connection.Close()




        connection = "FBConnectionString"
        Dim Reader As FirebirdSql.Data.FirebirdClient.FbDataReader = Nothing

        Cmd = New FirebirdSql.Data.FirebirdClient.FbCommand With {
            .Connection = New FirebirdSql.Data.FirebirdClient.FbConnection(ConfigurationManager.ConnectionStrings(environ & connection).ConnectionString),
            .CommandType = CommandType.Text
        }
        Cmd.CommandText = "SELECT
                                lhb.accountid as AccountID,
                                Trim(u.firstname) as FirstName,
                                Trim(u.lastname) as LastName,
                                Trim(l.business_name) AS CompanyName,
                                l.loanid AS TheLoan,
                                lhb.lh_bals_id,
                                o.lh_id,
                                0 as TheOrder,
                                COALESCE(Cast(lhb.num_units AS FLOAT)/100, 0)
                                    + COALESCE(Cast(lhS.num_units AS FLOAT)/100, 0) AS outstanding,
                                l.loanstatus,
                                l.ipo_end as DateOfLoan,
                                l.DATE_OF_LAST_PAYMENT as DateOfRepayment,
                                l.facility_amount as TotalFacilityAmount

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
                                    AccountID,
                                    FirstName,
                                    LastName,
                                    CompanyName,
                                    TheLoan,
                                    lhb.lh_bals_id,
                                    lh_id,
                                    outstanding,
                                    l.loanstatus,
                                    DateOfLoan,
                                    DateOfRepayment,
                                    TotalFacilityAmount"
        Cmd.Connection.Open()
        Reader = Cmd.ExecuteReader()

        '####
        'if running against a specific date instead of running against the current date then change the references in the above SQL for lh_balances and lh_balances_suspense  
        '                                                                                                                            to lh_pos_balances and lh_pos_balances_suspense
        'there also needs to be views created on Shadow - (run the Replicator then run this against Shadow) - which are copies of the lh_balances and lh_balances_suspense  
        '      called lh_pos_balances and lh_pos_balances_suspense. these need to have the line --->     and l.datecreated <= '2019-01-15'  
        '                                                                                   or  --->     and s.datecreated <= '2019-01-15' (change dates accordingly) added to them
        '####
        Dim ExtractList As New List(Of Extract)

        Dim nLoanStatus As Int16 = 0
        Dim dLoanEndDate As Date
        Dim nLoanid As Integer

        Do While Reader.Read()
            Dim newExtract As New Extract
            newExtract.AccountID = IIf(IsDBNull(Reader(0)), "", Reader(0))
            newExtract.Firstname = IIf(IsDBNull(Reader(1)), "", Reader(1))
            newExtract.LastName = IIf(IsDBNull(Reader(2)), "", Reader(2))
            newExtract.CompanyName = IIf(IsDBNull(Reader(3)), "", Reader(3))
            newExtract.LoanID = IIf(IsDBNull(Reader(4)), "", Reader(4))
            nLoanid = IIf(IsDBNull(Reader(4)), "", Reader(4))
            newExtract.LHBalID = IIf(IsDBNull(Reader(5)), "", Reader(5))
            newExtract.LHID = IIf(IsDBNull(Reader(6)), "", Reader(6))
            newExtract.OrderID = IIf(IsDBNull(Reader(7)), "", Reader(7))
            newExtract.Outstanding = IIf(IsDBNull(Reader(8)), 0, Reader(8))
            newExtract.Loanstatus = IIf(IsDBNull(Reader(9)), "", Reader(9))
            nLoanStatus = IIf(IsDBNull(Reader(9)), "", Reader(9))
            newExtract.DateOfAcquisition = "#01/01/0001#"
            newExtract.DateOfLoan = IIf(IsDBNull(Reader(10)), "#01/01/0001#", Reader(10))
            newExtract.DateOfRepayment = IIf(IsDBNull(Reader(11)), "#01/01/0001#", Reader(11))
            dLoanEndDate = IIf(IsDBNull(Reader(11)), "#01/01/0001#", Reader(11))
            newExtract.TotalFacilityAmount = IIf(IsDBNull(Reader(12)), 0, Reader(12) / 100)

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




        Reader = Nothing
 

        'thi is where the date of acquisition is added to the table

        'Dim TransList As New List(Of Trans)
        Dim nAccountID As Integer
        Dim nOrderID As Integer
        Dim nOrderPrev As Integer
        Dim nLHID As Integer


        Dim sAccountID As Integer = 0
        Dim sOrderID As Integer = 0

        Dim SummaryList As New List(Of Summary)
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

            
            Cmd.CommandText = "Select o.lh_id, o.accountid, o.loanid, o.orderprev, o.orderid
                                From orders  o
                                WHERE  o.accountid = @iaccountid
                                And (o.lh_id = @ilhid
                                 or o.loanid = @iloanid)
                                Order By o.lh_id desc, o.loanid desc,  o.orderprev desc, o.accountid  "
            Cmd.Parameters.Clear()

            Cmd.Parameters.Add("@iaccountid", FirebirdSql.Data.FirebirdClient.FbDbType.Integer).Value = nAccountID
            Cmd.Parameters.Add("@ilhid", FirebirdSql.Data.FirebirdClient.FbDbType.Integer).Value = nLHID
            Cmd.Parameters.Add("@iloanid", FirebirdSql.Data.FirebirdClient.FbDbType.Integer).Value = nLoanID

            Reader = Cmd.ExecuteReader()

            If Reader.Read() Then
                Extract.OrderID = IIf(IsDBNull(Reader(4)), "", Reader(4))
                nOrderID = Extract.OrderID
                nOrderPrev = IIf(IsDBNull(Reader(3)), 0, Reader(3))
            End If
            Reader.Close()


            Cmd.CommandText = "SELECT
                                ft.datecreated
                              FROM
                                fin_trans ft
                              WHERE ft.transtype in (1206, 1401)
                                AND ft.accountid = @iaccountid
                                AND (ft.orderid = @iorderid
                                Or ft.orderid = @iorderprev)"
            Cmd.Parameters.Clear()

            Cmd.Parameters.Add("@iaccountid", FirebirdSql.Data.FirebirdClient.FbDbType.Integer).Value = nAccountID
            Cmd.Parameters.Add("@iorderid", FirebirdSql.Data.FirebirdClient.FbDbType.Integer).Value = nOrderID
            Cmd.Parameters.Add("@iorderprev", FirebirdSql.Data.FirebirdClient.FbDbType.Integer).Value = nOrderPrev

            Reader = Cmd.ExecuteReader


            If Reader.Read() Then
                Extract.DateOfAcquisition = IIf(IsDBNull(Reader(0)), "", Reader(0))

            End If
            Reader.Close()

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

        Cmd.Connection.Close()
        Reader = Nothing
        Cmd.Connection = Nothing
        Cmd = Nothing

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
