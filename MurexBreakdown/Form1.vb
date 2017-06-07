Imports System.Globalization

Public Class Form1
    Private fr As FileReader

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        'Dim dt As New DataTable()

        'Try
        '    fr = New FileReader

        '    dt = fr.ReadExcelFile().Tables(0)

        '    dgv.DataSource = dt

        'Catch ex As Exception
        '    MsgBox(ex.Message)
        'End Try

        ''Dim dm As New dataManagement
        ''For Each el In dm.getListOfBreakdowns(dt)
        ''    lst.Items.Add(el)
        ''Next

        'Dim c As New dataManagement
        'For Each el In c.getListOfCurrencies(dt)
        '    lst.Items.Add(el)
        'Next

        ''Dim d As Date = Date.FromOADate(dt.Rows(1).Item(0))
        ''MsgBox(d)


        Try
            fr = New FileReader




        Catch ex As Exception
            MsgBox(ex.Message, vbOKOnly + vbCritical, "Error !")
        End Try

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim d As Date = Date.FromOADate(41498)
        MsgBox(d)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Dim studentList = New List(Of Student) From {
            New Student() With {.StudentID = 1, .StudentName = "John", .Age = 13},
            New Student() With {.StudentID = 2, .StudentName = "Moin", .Age = 21},
            New Student() With {.StudentID = 3, .StudentName = "Bill", .Age = 18},
            New Student() With {.StudentID = 4, .StudentName = "Ram", .Age = 20},
            New Student() With {.StudentID = 5, .StudentName = "Ron", .Age = 15}
        }

        Dim ts As IList(Of Student) = (From s In studentList
                                       Where s.Age >= 20 Select s).ToList

        Dim ts1 As IList(Of Student) = studentList.Where(Function(s) s.Age > 12 And s.Age < 20).ToList


        For Each t In ts
            lst.Items.Add(t.StudentID & "-" & t.StudentName)
        Next


        For Each t In ts1
            lst.Items.Add(t.StudentID & "-" & t.StudentName)
        Next

        ' Dim tt As SourceDataStructure
        fr = New FileReader()

        Dim los As List(Of SourceDataStructure) = fr.getListOfRows(fr.ReadExcelFile().Tables(0))






    End Sub


End Class

Public Class Student
    Private _studentID As Integer
    Private _studentName As String
    Private _age As Integer

    Property StudentID() As Integer
        Get
            Return _studentID
        End Get
        Set(ByVal Value As Integer)
            _studentID = Value
        End Set
    End Property

    Property StudentName() As String
        Get
            Return _studentName
        End Get
        Set(ByVal Value As String)
            _studentName = Value
        End Set
    End Property

    Property Age() As Integer
        Get
            Return _age
        End Get
        Set(ByVal Value As Integer)
            _age = Value
        End Set
    End Property


End Class
