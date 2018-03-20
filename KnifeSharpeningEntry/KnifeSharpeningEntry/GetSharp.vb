Imports System.Data.Sql
Imports System.Data.SqlClient
Public Module GlobalVariables
    Public SOpen As String
    Public connectionString As String = "Data Source=cfcsql02\inst02;Initial Catalog=ULTIPRO_CFC;Integrated Security=False;User ID=jobber;Password=runthatjob"
    Public conn As New SqlConnection(connectionString)
    Public tm As String
    Public FirstName As String
    Public LastName As String
    Public Department As String
    Public JobDescription As String
    Public SetNum As String
    Public operdate As DateTime
    Public comment As String
    Public knife As String
    Public operation As String
    Public issue As String
    Public knifecomment As String
    Public thirtydaysago As DateTime
End Module
Module GetSharp
    Sub Main()
        If SOpen = "" Then
            conn.Open()
            SOpen = "Y"
        End If
        Console.BackgroundColor = ConsoleColor.White
        Console.WindowWidth = 120
        Console.WindowHeight = 35
        Console.Clear()
        Console.BackgroundColor = ConsoleColor.White
        Console.ForegroundColor = ConsoleColor.Blue
        If StrConv(Environment.UserName, VbStrConv.Lowercase) <> "gschell" And StrConv(Environment.UserName, VbStrConv.Lowercase) <> "tvelez" And StrConv(Environment.UserName, VbStrConv.Lowercase) <> "tleece" And StrConv(Environment.UserName, VbStrConv.Lowercase) <> "jims" Then
            Console.ForegroundColor = ConsoleColor.Red
            Console.WriteLine("Hello {0}.  You are not an authorized knife sharpener.", Environment.UserName)
            Console.ReadLine()
            Exit Sub
        End If
        Console.WriteLine("Hello {0}.  This is KnifeSharpeningEntry version 1.5", Environment.UserName)
        thirtydaysago = System.DateTime.Now.AddDays(-30)
GetTm:
        tm = ""
        comment = ""
        knife = ""
        operation = ""
        issue = ""
        knifecomment = ""
        Console.WriteLine()
        Console.ForegroundColor = ConsoleColor.Black
        Console.Write("Enter Team Member Number> ")
        tm = Console.ReadLine()
        If tm.Length = 0 Or tm = "." Then
            If SOpen = "Y" Then
                conn.Close()
            End If
            Exit Sub
        End If
        If IsNumeric(tm) = 0 Then
            TMSearch()
            GoTo GetTm
        End If
        tm = Right("000000" & tm, 6)
        DisplayTM()
        DisplayHist()
GetDate:
        operdate = System.DateTime.Now
        Dim xdate As String
        Console.WriteLine()
        Console.ForegroundColor = ConsoleColor.Black
        Console.Write("Enter Operation Date or Enter for " + operdate.ToString("MM/dd/yy HH:mm") + "> ")
        xdate = Console.ReadLine()
        Dim isValid = False
        Dim format() = {"M/d/yy HH:mm", "M/d/yy H:mm", "M/d/yyyy HH:mm", "M/d/yyyy H:mm"}
        isValid = Date.TryParseExact(xdate, format, System.Globalization.DateTimeFormatInfo.InvariantInfo, Globalization.DateTimeStyles.None, operdate)
        If xdate = "." Then
            GoTo GetTm
        End If
        If xdate = "T" Or xdate = "t" Then
            DisplayTM()
            GoTo GetDate
        End If
        If xdate = "H" Or xdate = "h" Then
            DisplayHist()
            GoTo GetDate
        End If
        If xdate = "." Then
            GoTo GetTm
        End If
        If xdate.Length = 0 Then
            operdate = System.DateTime.Now
            GoTo GetComment
        End If
        If isValid = False Then
            Console.WriteLine()
            Console.ForegroundColor = ConsoleColor.Red
            Console.WriteLine("Invalid date entered")
            GoTo GetDate
        End If
        operdate = Convert.ToDateTime(xdate)
        'Console.WriteLine(operdate.ToString("MM/dd/yy HH:mm"))
GetComment:
        comment = ""
        Console.WriteLine()
        Console.ForegroundColor = ConsoleColor.Black
        Console.Write("Enter Comment> ")
        'Console.ResetColor()
        comment = Console.ReadLine()
        If comment.Length = 0 Then
            GoTo GetKnife
        End If
        If comment = "." Then
            GoTo GetTm
        End If
        If comment = "T" Or comment = "t" Then
            DisplayTM()
            GoTo GetComment
        End If
        If comment = "H" Or comment = "h" Then
            DisplayHist()
            GoTo GetComment
        End If
        operation = ""
        issue = ""
        knifecomment = ""
        WriteKnifeSharp()
        GoTo GetTm
GetKnife:
        knife = ""
        Console.WriteLine()
        Console.ForegroundColor = ConsoleColor.Black
        Console.Write("Enter knife number> ")
        'Console.ResetColor()
        knife = Console.ReadLine()
        If knife.Length = 0 And comment <> "" Then
            WriteKnifeSharp()
            GoTo GetTm
        End If
        If knife.Length = 0 And comment.Length = 0 Then
            GoTo GetTm
        End If
        If knife = "." Then
            GoTo GetTm
        End If
        If StrConv(knife, VbStrConv.Lowercase) = "set" Then
            UpdateSet()
            DisplayTM()
            GoTo GetKnife
        End If

        If knife = "T" Or knife = "t" Then
            DisplayTM()
            GoTo GetKnife
        End If
        If knife = "H" Or knife = "h" Then
            knife = ""
            DisplayHist()
            GoTo GetKnife
        End If
        If knife = "?" Then
            knife = ""
            DisplayTM()
            DisplayHist()
            GoTo GetKnife
        End If
        If IsNumeric(knife) = 0 Then
            Console.WriteLine()
            Console.ForegroundColor = ConsoleColor.Red
            Console.WriteLine("Knife number must be numeric")
            'Console.ResetColor()
            GoTo GetKnife
        End If
GetOperation:
        operation = ""
        Console.WriteLine()
        Console.ForegroundColor = ConsoleColor.Blue
        Console.WriteLine("H = Hollow Grinding")
        Console.WriteLine("E = Edging/Buffing")
        Console.WriteLine()
        Console.ForegroundColor = ConsoleColor.Black
        Console.Write("Enter Operation> ")
        'Console.ResetColor()
        operation = Console.ReadLine()
        If operation.Length = 0 Then
            GoTo GetIssue
        End If
        If operation = "." Then
            GoTo GetKnife
        End If
        If operation = "T" Or operation = "t" Or operation = "?" Then
            DisplayTM()
            DisplayHist()
            GoTo GetOperation
        End If
        If operation <> "h" And operation <> "H" And operation <> "e" And operation <> "E" Then
            Console.WriteLine()
            Console.ForegroundColor = ConsoleColor.Red
            Console.WriteLine("Invalid Operation")
            'Console.ResetColor()
            GoTo GetOperation
        End If
        If operation = "h" Or operation = "H" Then
            operation = "Hollow Grinding"
        End If
        If operation = "e" Or operation = "E" Then
            operation = "Edging/Buffing"
        End If
GetIssue:
        issue = ""
        Console.WriteLine()
        Console.ForegroundColor = ConsoleColor.Blue
        Console.WriteLine("D = Dirty")
        Console.WriteLine("K = Knicks")
        Console.WriteLine("R = Rusty")
        Console.WriteLine("O = Other")
        Console.WriteLine()
        Console.ForegroundColor = ConsoleColor.Black
        Console.Write("Enter knife issue> ")
        'Console.ResetColor()
        issue = Console.ReadLine()
        If issue = "." Then
            GoTo GetOperation
        End If
        If issue = "T" Or issue = "t" Or issue = "?" Then
            DisplayTM()
            DisplayHist()
            GoTo GetIssue
        End If
        If issue <> "d" And issue <> "k" And issue <> "r" And issue <> "o" And issue <> "D" And issue <> "K" And issue <> "R" And issue <> "O" And issue.Length <> 0 Then
            Console.WriteLine()
            Console.ForegroundColor = ConsoleColor.Red
            Console.WriteLine("Invalid Issue")
            'Console.ResetColor()
            GoTo GetIssue
        End If
        If issue = "d" Or issue = "D" Then
            issue = "Dirty"
        End If
        If issue = "k" Or issue = "K" Then
            issue = "Knicks"
        End If
        If issue = "r" Or issue = "R" Then
            issue = "Rusty"
        End If
        If issue = "o" Or issue = "O" Then
            issue = "Other"
        End If
        If issue.Length = 0 And operation.Length = 0 Then
            Console.WriteLine()
            Console.ForegroundColor = ConsoleColor.Red
            Console.WriteLine("Operation and/or Issue entry is required")
            'Console.ResetColor()
            GoTo GetOperation
        End If
GetKnifeComment:
        knifecomment = ""
        Console.WriteLine()
        Console.ForegroundColor = ConsoleColor.Black
        Console.Write("Enter Knife Comment> ")
        'Console.ResetColor()
        knifecomment = Console.ReadLine()
        If knifecomment = "." Then
            GoTo GetIssue
        End If
        If knifecomment.Length = 0 And issue = "Other" Then
            Console.WriteLine()
            Console.ForegroundColor = ConsoleColor.Red
            Console.WriteLine("Comment required for 'Other' issue")
            'Console.ResetColor()
            GoTo GetKnifeComment
        End If
        WriteKnifeSharp()
        knife = ""
        operation = ""
        issue = ""
        knifecomment = ""
        GoTo GetKnife
    End Sub
    Sub DisplayTM()
        Console.WriteLine()
        ' Console.WriteLine("Display team member " + tm)
        Dim sqlcmd As String
        sqlcmd = "Select DISTINCT EC.EeCEmpNo,EP.eepNamefirst,EP.eepnamelast,O.OrgCode + ' ' + O.OrgDesc AS Department,J.JbcDesc,Sup.eepNameLast + ', ' + Sup.eepNamefirst as SupervisorFullName, isnull(ks.KSSetNum,0) as KSSetNum" _
                + " From EmpComp EC" _
                + " INNER Join EmpPers EP ON EP.EepEEID = EC.EecEEID" _
                + " INNER Join Company C ON c.CmpCoID = EC.EEcCoID" _
                + " Left Join JobCode J ON J.jbcJObCode = eecJobCode" _
                + " LEFT JOIN OrgLevel O ON O.OrgCode = EC.EecOrgLvl2" _
                + " Left Join EmpPers Sup on Sup.EEPEEID = ec.eecsupervisorid" _
                + " left join knifeset ks on ks.KSTM = EC.EeCEmpNo" _
                + " WHERE EC.eecemplstatus <> 'T' And EC.EeCEmpNo = '" + tm + "'"
        Dim comm As New SqlCommand(sqlcmd, conn)
        Dim reader As SqlDataReader = comm.ExecuteReader
        Dim i As Integer
        i = 0
        Console.ForegroundColor = ConsoleColor.Blue
        While reader.Read()
            Console.WriteLine("{0} {1} {2}, Dept. - {3}, Job - {4}, Srpv. - {5}, Set# - {6}", Trim(reader("EeCEmpNo").ToString), Trim(reader("eepNamefirst").ToString), Trim(reader("eepnamelast").ToString),
                              reader("Department").ToString, reader("JbcDesc").ToString, reader("SupervisorFullName").ToString, reader("KSSetNum").ToString)
            'Console.WriteLine(Trim(reader("EeCEmpNo").ToString) + " " + Trim(reader("eepNamefirst").ToString) + " " + Trim(reader("eepnamelast").ToString) + " | " _
            '                  + reader("Department").ToString + " | " + reader("JbcDesc").ToString + " | " + reader("SupervisorFullName").ToString + " | " + reader("KSSetNum").ToString)
            FirstName = Trim(reader("eepNamefirst").ToString)
            LastName = Trim(reader("eepnamelast").ToString)
            Department = Trim(reader("Department").ToString)
            JobDescription = Trim(reader("JbcDesc").ToString)
            SetNum = reader("KSSetNum")
            i = i + 1
        End While
        'Console.ResetColor()
        If i = 0 Then
            Console.WriteLine()
            Console.ForegroundColor = ConsoleColor.Blue
            Console.WriteLine(" Invalid Team Member ")
            'Console.ResetColor()
        End If
        reader.Close()
    End Sub
    Sub UpdateSet()
        Console.WriteLine()
        Dim sqlcmd As String
        sqlcmd = "Select isnull(KSTM,0) as KSTM, isnull(KSSetNum,0) as KSSetNum From KnifeSet WHERE KSTM = '" + tm + "'"
        Dim comm As New SqlCommand(sqlcmd, conn)
        Dim reader As SqlDataReader = comm.ExecuteReader
        Dim i As Integer
        i = 0
        Console.ForegroundColor = ConsoleColor.Blue
        While reader.Read()
            Console.WriteLine("TM# - {0}, Set# - {1}", Trim(reader("KSTM").ToString), Trim(reader("KSSetNum").ToString))
            SetNum = reader("KSSetNum")
            i = i + 1
        End While
        If i = 0 Then
            Console.WriteLine()
            Console.ForegroundColor = ConsoleColor.Blue
            Console.WriteLine(" Adding Knife Set for new TM")
        End If
        reader.Close()
        Dim newset As String
GetSet:
        newset = ""
        Console.WriteLine()
        Console.ForegroundColor = ConsoleColor.Black
        Console.Write("Enter a new knife set number> ")
        'Console.ResetColor()
        newset = Console.ReadLine()
        If newset = "." Or newset.Length = 0 Then
            Exit Sub
        End If
        If IsNumeric(newset) = False Then
            Console.WriteLine()
            Console.ForegroundColor = ConsoleColor.Red
            Console.WriteLine("Invalid set number.  Numbers only.")
            'Console.ResetColor()
            GoTo GetSet
        End If
        If i = 0 Then
            sqlcmd = "insert into knifeset (KSTM, KSSetNum) values('" + tm + "','" + newset + "')"
        Else
            sqlcmd = "Update knifeset set KSSetNum = '" + newset + "' where KSTM = '" + tm + "'"
        End If
        comm = New SqlCommand(sqlcmd, conn)
        comm.ExecuteNonQuery()
        Console.ForegroundColor = ConsoleColor.Blue
        Console.WriteLine()
        Console.WriteLine("Set number updated")
    End Sub
    Sub TMSearch()
        Console.WriteLine()
        ' Console.WriteLine("Display team member " + tm)
        Dim sqlcmd As String
        sqlcmd = "Select DISTINCT EC.EeCEmpNo,EP.eepNamefirst,EP.eepnamelast,O.OrgCode + ' ' + O.OrgDesc AS Department,J.JbcDesc" _
                + " From EmpComp EC" _
                + " INNER Join EmpPers EP ON EP.EepEEID = EC.EecEEID" _
                + " INNER Join Company C ON c.CmpCoID = EC.EEcCoID" _
                + " Left Join JobCode J ON J.jbcJObCode = eecJobCode" _
                + " LEFT JOIN OrgLevel O ON O.OrgCode = EC.EecOrgLvl2" _
                + " WHERE EC.eecemplstatus <> 'T' And (EP.eepNamefirst like '%" + tm + "%' or EP.eepnamelast like '%" + tm + "%')"
        Dim comm As New SqlCommand(sqlcmd, conn)
        Dim reader As SqlDataReader = comm.ExecuteReader
        Dim i As Integer
        i = 0
        Console.ForegroundColor = ConsoleColor.Blue
        While reader.Read()
            Console.WriteLine("{0} {1} {2}, Dept. - {3}, Job - {4}", Trim(reader("EeCEmpNo").ToString), Trim(reader("eepNamefirst").ToString), Trim(reader("eepnamelast").ToString),
                              reader("Department").ToString, reader("JbcDesc").ToString)
            i = i + 1
        End While
        'Console.ResetColor()
        If i = 0 Then
            Console.WriteLine()
            Console.ForegroundColor = ConsoleColor.Blue
            Console.WriteLine(" Invalid Team Member ")
            'Console.ResetColor()
        End If
        reader.Close()
    End Sub
    Sub DisplayHist()
        Console.WriteLine()
        ' Console.WriteLine("Display team member " + tm + " name, dept, job desc., last 30 day entries, set #")
        Dim sqlcmd As String
        If knife.Length = 0 Then
            sqlcmd = "SELECT * FROM KnifeSharpening where KSTM = '" + tm + "' and KSCreatedDate > '" + thirtydaysago.ToString("MM/dd/yyyy") + "' order by KSOperDT"
        Else
            sqlcmd = "SELECT * FROM KnifeSharpening where KSTM = '" + tm + "' and KSKnifeNum = '" + knife + "' and KSCreatedDate > '" + thirtydaysago.ToString("MM/dd/yyyy") + "' order by KSOperDT"
        End If
        'Console.WriteLine()
        'Console.WriteLine(sqlcmd)
        Dim comm As New SqlCommand(sqlcmd, conn)
        Dim reader As SqlDataReader = comm.ExecuteReader
        Dim i As Integer
        i = 0
        Console.ForegroundColor = ConsoleColor.Blue
        Console.WriteLine("Operation DT   | Comment    | Set# | Knife# | Operation       | Issue  | Opr Comment | Created By | Created Date")
        Console.ForegroundColor = ConsoleColor.Blue
        Dim d1 As DateTime
        Dim ds1 As String
        Dim d2 As DateTime
        Dim ds2 As String
        Dim c1 As String
        Dim c2 As String
        While reader.Read()
            d1 = reader("KSOperDT")
            ds1 = d1.ToString("MM/dd/yy HH:mm")
            d2 = reader("KSCreatedDate")
            ds2 = d2.ToString("MM/dd/yy HH:mm")
            c1 = reader("KSComment")
            c1 = Left(c1, 10)
            c2 = reader("KSOperComment")
            c2 = Left(c2, 11)
            Console.WriteLine("{0} | {1,-10} | {2,4} | {3,6} | {4,-15} | {5,-6} | {6,-11} | {7,-10} | {8}", ds1, c1, reader("KSSetNum").ToString, reader("KSKnifeNum").ToString _
                              , reader("KSOperation").ToString, reader("KSKnifeIssue").ToString, c2, reader("KSCreatedBy").ToString, ds2)
            'Console.WriteLine(reader("KSOperDT").ToString + " | " + reader("KSComment").ToString + " | " + reader("KSSetNum").ToString + " | " + reader("KSKnifeNum").ToString _
            '                  + " | " + reader("KSOperation").ToString + " | " + reader("KSKnifeIssue").ToString + " | " + reader("KSOperComment").ToString + " | " + reader("KSCreatedBy").ToString + " | " + reader("KSCreatedDate").ToString)
            i = i + 1
        End While
        'Console.ResetColor()
        If i = 0 Then
            Console.WriteLine()
            Console.ForegroundColor = ConsoleColor.Blue
            Console.WriteLine(" No history found ")
            'Console.ResetColor()
        End If
        reader.Close()
    End Sub
    Sub WriteKnifeSharp()
        'Console.WriteLine()
        'Console.WriteLine("Write record " + tm + " " + comment + " " + knife + " " + operation + " " + issue + " " + knifecomment)
        'SetNum = "10"
        Dim nSetNum As Integer
        Dim nKnife As Integer
        nSetNum = Convert.ToDecimal(SetNum)
        If knife.Length = 0 Then
            nKnife = 0
        Else
            nKnife = Convert.ToDecimal(knife)
        End If
        Dim username As String
        If TypeOf My.User.CurrentPrincipal Is Security.Principal.WindowsPrincipal Then
            Dim parts() As String = Split(My.User.Name, "\")
            username = parts(1)
        Else
            username = My.User.Name
        End If
        Dim sqlcmd As String
        Dim thedate As DateTime = System.DateTime.Now
        sqlcmd = "Insert into KnifeSharpening (KSTM, KSFirstName, KSLastName, KSDepartment, KSJobDescription, KSComment, " _
                        + "KSOperDT, KSSetNum, KSKnifeNum, KSOperation, KSKnifeIssue, KSOperComment, KSCreatedBy, KSCreatedDate) " _
                        + "values ('" + tm + "','" + FirstName + "','" + LastName + "','" + Department + "','" + JobDescription + "','" + comment _
                        + "','" + operdate.ToString("MM/dd/yyyy HH:mm") + "'," + nSetNum.ToString + "," + nKnife.ToString + ",'" + operation + "','" + issue + "','" + knifecomment + "','" + Environment.UserName + "','" + thedate.ToString("MM/dd/yyyy HH:mm") + "')"
        'Console.WriteLine("sqlcmd = " + sqlcmd)
        Dim comm As New SqlCommand(sqlcmd, conn)
        comm.ExecuteNonQuery()
        Console.ForegroundColor = ConsoleColor.Blue
        Console.WriteLine()
        Console.WriteLine("Operation Saved")
        'Console.ResetColor()
    End Sub
End Module
