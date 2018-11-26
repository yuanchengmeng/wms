Imports Intermec.DataCollection2
Imports System.Data.SqlTypes
Imports System.Data.SqlClient
Imports System.Data.SqlServerCe


Imports System.Net.Sockets
Imports System.Net
Imports System.Text
Imports System.Threading
Imports System.String
Imports System.IO
Imports System.Data

Public Class Form1

    Friend WithEvents bcr As Intermec.DataCollection2.BarcodeReader


    'Private oOneInstance As SingleInstanceApplication = Nothing

    Dim SQL As New SqlConnection


    Dim SQLServer As String
    Dim SQLDatabase As String
    Dim SQLUser As String
    Dim SQLPassword As String

    Dim SQLK3 As New SqlConnection
    Dim SQLServerK3 As String
    Dim SQLDatabaseK3 As String
    Dim SQLUserK3 As String
    Dim SQLPasswordK3 As String

    Dim SQLMES As New SqlConnection
    Dim SQLServerMES As String
    Dim SQLDatabaseMES As String
    Dim SQLUserMES As String
    Dim SQLPasswordMES As String

    Dim NowUser As String
    Dim NowJobNo As String
    Dim NowMachine As String
    Dim NowClass As String
    Dim BoxCode As String
    Dim NowDate As String

    Dim OrderID As Long
    Dim DetailID As Long
    Dim ArrP(,)
    Dim ArrO(,)

    Dim PageNo As Long
    Dim BoxNum As Long
    Dim BillNo As String
    Dim FBillNo As String
    Dim FBillID As Long
    Dim FInterID As Long
    Dim OldFInterID As Long
    Dim FItemID As Long
    Dim FCustID As Long
    Dim FUnitID As Long
    Dim Product As String
    Dim k3User As String
    Dim vehicleID As Long
    Dim flag As Long
    Dim Fcount As Long
    Dim FaultLoc As Long
    Dim ck As Long
    Dim FExchangeRate As Decimal

    Dim FAllCount As Long
    Dim FAllEntryID As Long
    Dim FAllFInterID As Long
    '20180602新增其它入库
    Dim QtInType As Long '入库类型：1.正常入库 2.外购入库 3.其它入库

    Dim StockID As Long '库区id
    Dim stockArr(,)  '全部库区
    Dim ReasonID As Long '取消原因id
    Dim ReasonArr(,)  '全部取消入库原因

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Me.Width = 200
        Me.Height = 200

        Try
            bcr = New Intermec.DataCollection2.BarcodeReader()
            bcr.PostRead = True
            bcr.Symbology.Code39.Enable = CSymbology.CCode39.EEnable.Enable
            bcr.Symbology.Code128.EnableCode128 = CSymbology.CCode128.EEnableCode128.Enable
        Catch bcrexp As BarcodeReaderException
            MessageBox.Show(bcrexp.Message)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        Timer1.Enabled = True

    End Sub

    Private Sub bcr_BarcodeRead(ByVal sender As Object, ByVal bre As Intermec.DataCollection2.BarcodeReadEventArgs) Handles bcr.BarcodeRead
        Try
            TextBox5.Visible = True
            TextBox5.Text = bre.StrBarcodeData

        Catch bcrexp As BarcodeReaderException
            MessageBox.Show(bcrexp.Message)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Public Sub New()

        '' 此调用是 Windows 窗体设计器所必需的。
        InitializeComponent()


        '' 在 InitializeComponent() 调用之后添加任何初始化。

    End Sub



    Private Sub Form1_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Disposed
        SQL.Close()
        SQLK3.Close()
    End Sub


    Sub ShowPanel(ByVal ID As Integer)
        Dim PP As Panel
        PageNo = ID

        Panel1.Visible = False
        Panel2.Visible = False
        Panel3.Visible = False
        Panel4.Visible = False
        Panel5.Visible = False
        Panel6.Visible = False
        Panel7.Visible = False
        Panel8.Visible = False
        Panel9.Visible = False
        Panel10.Visible = False
        Panel11.Visible = False
        Panel12.Visible = False
        Panel13.Visible = False
        Panel14.Visible = False
        Panel15.Visible = False
        Panel16.Visible = False
        Panel17.Visible = False
        Panel18.Visible = False
        Panel19.Visible = False
        Panel20.Visible = False
        Panel21.Visible = False
        Panel22.Visible = False
        Panel23.Visible = False
        Panel24.Visible = False
        Panel25.Visible = False
        Panel26.Visible = False
        Select Case ID
            Case 1 : PP = Panel1
            Case 2 : PP = Panel2
            Case 3 : PP = Panel3
            Case 4 : PP = Panel4
            Case 5 : PP = Panel5
            Case 6 : PP = Panel6
            Case 7 : PP = Panel7
            Case 8 : PP = Panel8
            Case 9 : PP = Panel9
            Case 10 : PP = Panel10
            Case 11 : PP = Panel11
            Case 12 : PP = Panel12
            Case 13 : PP = Panel13
            Case 14 : PP = Panel14
            Case 15 : PP = Panel15
            Case 16 : PP = Panel16
            Case 17 : PP = Panel17
            Case 18 : PP = Panel18
            Case 19 : PP = Panel19
            Case 20 : PP = Panel20
            Case 21 : PP = Panel21
            Case 22 : PP = Panel22
            Case 23 : PP = Panel23
            Case 24 : PP = Panel24
            Case 25 : PP = Panel25
            Case 26 : PP = Panel26
            Case Else : PP = Panel1
        End Select
        PP.Left = 0
        PP.Top = 0
        PP.Visible = True
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If NowUser = "" Then MsgBox("请先登录") : Exit Sub
        ShowPanel(22)
    End Sub '入库管理

    Sub RefUserShow()
        Label2.Text = "工号：" & NowJobNo
        Label26.Text = "姓名：" & NowUser
        Label51.Text = "日期:" & NowDate
        Label52.Text = "班别:" & NowClass
    End Sub

    Sub SetLog(ByVal LogStr As String)
        Dim Path As String = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase)
        Dim sw As IO.StreamWriter = New StreamWriter(Path & "\Log" & Int(Now.ToOADate).ToString & ".txt", True)  'true是指以追加的方式打开指定文件  
        sw.WriteLine(LogStr)
        sw.Flush()
        sw.Close()
        sw = Nothing
    End Sub

    Sub GetCombo(ByVal ComboString As String, Optional ByVal Para1 As String = "")

        Dim Cb As ComboBox
        Dim t As Long
        Dim StrSQL As String
        Dim StrErr As String
        StrSQL = ""
        flag = 0
        Select Case ComboString
            Case "作业班组" : Cb = ComboBox5 : StrSQL = "select name from hand_class"
            Case Else
                MsgBox("未知Combo")
        End Select

        Dim Arr(,)

        If StrSQL <> "" Then
            Cb.Items.Clear()

            StrErr = GetRst(StrSQL, Arr, SQL)
            If StrErr <> "" Then
                flag = 1
                SetLog("获取" & ComboString & "失败" & vbCrLf & StrSQL)
                MsgBox(StrErr)
                Exit Sub
            End If

            For t = 1 To UBound(Arr, 2)
                If Arr(1, t).ToString <> "" Then Cb.Items.Add(Arr(1, t).ToString)
            Next
            'End Select

        End If
    End Sub

    Sub GetOrder()
        Dim t As Long
        Dim StrSQL As String
        Dim StrErr As String 

        StrSQL = "select FBillNo,FInterID from SEOrder where  FClosed=0"
        ComboBox18.Items.Clear()

        StrErr = GetRst(StrSQL, ArrP, SQLK3)
        If StrErr <> "" Then
            SetLog("获取订单号失败" & vbCrLf & StrSQL)
            MsgBox("获取订单号失败" & vbCrLf & StrErr)
            Exit Sub
        End If

        For t = 1 To UBound(ArrP, 2)
            If ArrP(1, t).ToString <> "" Then ComboBox18.Items.Add(ArrP(1, t).ToString)

        Next

    End Sub


    Function GetRst(ByVal StrSQL, ByRef Arr(,), ByRef SqlCon) As String
        Dim StrErr As String = ""
        Dim CMD As New SqlCommand
        Dim RA As SqlDataReader
        Dim t As Integer
        GetRst = ""
        StrErr = ConSQL(SqlCon)
        If StrErr <> "" Then GetRst = StrErr : Exit Function
        Try
            CMD.CommandText = StrSQL
            CMD.CommandTimeout = 20
            CMD.Connection = SqlCon
            RA = CMD.ExecuteReader
            ReDim Arr(RA.FieldCount, 0)
            While RA.Read
                ReDim Preserve Arr(RA.FieldCount, UBound(Arr, 2) + 1)
                For t = 1 To RA.FieldCount
                    Arr(t, UBound(Arr, 2)) = RA(t - 1).ToString()
                Next
            End While
            RA.Close()
        Catch ex As Exception
            GetRst = "获取数据失败" & ex.Message & vbCrLf & ex.StackTrace
        End Try

    End Function

    Function ExeSQLS(ByVal StrSQL() As String, ByRef SqlCon As SqlConnection) As String
        Dim StrErr As String = ""
        Dim Ts As SqlTransaction
        Dim CMD As New SqlCommand
        Dim t As Integer

        ExeSQLS = ""

        StrErr = ConSQL(SqlCon)
        If StrErr <> "" Then ExeSQLS = StrErr : Exit Function

        Ts = SqlCon.BeginTransaction
        CMD.Transaction = Ts
        CMD.Connection = SqlCon

        Try
            For t = 1 To UBound(StrSQL)
                CMD.CommandText = StrSQL(t)
                CMD.ExecuteNonQuery()
            Next
            Ts.Commit()

        Catch ex As Exception
            Ts.Rollback()
            ExeSQLS = "执行数据失败" & ex.Message & vbCrLf & ex.StackTrace
        End Try
    End Function

    Function ConSQL(ByRef SQL As SqlConnection) As String
        ConSQL = ""
        Try
            If SQL.State <> Data.ConnectionState.Open Then SQL.Open()
        Catch ex As Exception
            'ConSQL = "网络连接失败,无法连接数据库!!"
            'SetLog(ex.Message & vbCrLf & ex.StackTrace)
            ConSQL = ex.Message & vbCrLf & ex.StackTrace
            'MsgBox(ex.Message)
            'MsgBox(ex.StackTrace)
        End Try
    End Function

    Private Sub readXMl()
        Try
            Dim Dir As String = System.Reflection.Assembly.GetExecutingAssembly.GetName.CodeBase
            Dir = Dir.Substring(0, Dir.LastIndexOf("\"))
            Dim doc As New Xml.XmlDocument
            doc.Load(Dir & "\Set.XML")
            Dim re As Xml.XmlNodeReader = New Xml.XmlNodeReader(doc)
            Dim tmpStr As String = ""
            Dim name As String = ""
            While re.Read
                Select Case re.NodeType
                    Case Xml.XmlNodeType.Element
                        name = re.Name
                    Case Xml.XmlNodeType.Text
                        If name.Equals("DataSource") Then SQLServer = re.Value
                        If name.Equals("InitialCatalog") Then SQLDatabase = re.Value
                        If name.Equals("UserID") Then SQLUser = re.Value
                        If name.Equals("Password") Then SQLPassword = re.Value
                End Select
            End While

        Catch ex As Exception
            MsgBox(ex.Message & vbCrLf & ex.StackTrace)
        End Try

    End Sub

    Private Sub readXMlK3()
        Try
            Dim Dir As String = System.Reflection.Assembly.GetExecutingAssembly.GetName.CodeBase
            Dir = Dir.Substring(0, Dir.LastIndexOf("\"))

            Dim doc As New Xml.XmlDocument
            doc.Load(Dir & "\SetK3.XML")
            Dim re As Xml.XmlNodeReader = New Xml.XmlNodeReader(doc)
            Dim tmpStr As String = ""
            Dim name As String = ""
            While re.Read
                Select Case re.NodeType
                    Case Xml.XmlNodeType.Element
                        name = re.Name
                    Case Xml.XmlNodeType.Text
                        If name.Equals("DataSource") Then SQLServerK3 = re.Value
                        If name.Equals("InitialCatalog") Then SQLDatabaseK3 = re.Value
                        If name.Equals("UserID") Then SQLUserK3 = re.Value
                        If name.Equals("Password") Then SQLPasswordK3 = re.Value
                End Select
            End While

        Catch ex As Exception
            MsgBox(ex.Message & vbCrLf & ex.StackTrace)
        End Try

    End Sub

    Private Sub readXMlMES()
        Try
            Dim Dir As String = System.Reflection.Assembly.GetExecutingAssembly.GetName.CodeBase
            Dir = Dir.Substring(0, Dir.LastIndexOf("\"))

            Dim doc As New Xml.XmlDocument
            doc.Load(Dir & "\SetMES.XML")
            Dim re As Xml.XmlNodeReader = New Xml.XmlNodeReader(doc)
            Dim tmpStr As String = ""
            Dim name As String = ""
            While re.Read
                Select Case re.NodeType
                    Case Xml.XmlNodeType.Element
                        name = re.Name
                    Case Xml.XmlNodeType.Text
                        If name.Equals("DataSource") Then SQLServerMES = re.Value
                        If name.Equals("InitialCatalog") Then SQLDatabaseMES = re.Value
                        If name.Equals("UserID") Then SQLUserMES = re.Value
                        If name.Equals("Password") Then SQLPasswordMES = re.Value
                End Select
            End While

        Catch ex As Exception
            MsgBox("读取xmlmes.xml失败！")
        End Try

    End Sub


    Private Sub createXML()
        Try
            Dim Dir As String = System.Reflection.Assembly.GetExecutingAssembly.GetName.CodeBase
            Dir = Dir.Substring(0, Dir.LastIndexOf("\"))

            Dim writer As New Xml.XmlTextWriter(Dir & "\Set.XML", System.Text.Encoding.GetEncoding("utf-8"))
            '使用自动缩进便于阅读
            writer.Formatting = Xml.Formatting.Indented
            writer.WriteRaw("<?xml version=""1.0"" encoding=""utf-8"" ?>")
            '书写根元素()
            writer.WriteStartElement("Config")
            '添加次级元素
            writer.WriteStartElement("DatabaseSetting")
            '添加子元素()
            writer.WriteElementString("DataSource", "1")
            writer.WriteElementString("InitialCatalog", "2")
            writer.WriteElementString("UserID", "3")
            writer.WriteElementString("Password", "4")
            '关闭次级元素DatabaseSetting
            writer.WriteEndElement()
            '关闭根元素
            writer.WriteFullEndElement()
            '将XML写入文件并关闭writer
            writer.Close()
        Catch ex As Exception
            MsgBox(ex.Message & vbCrLf & ex.StackTrace)
        End Try
    End Sub

    Private Sub createXMLK3()
        Try
            Dim Dir As String = System.Reflection.Assembly.GetExecutingAssembly.GetName.CodeBase
            Dir = Dir.Substring(0, Dir.LastIndexOf("\"))

            Dim writer As New Xml.XmlTextWriter(Dir & "\SetK3.XML", System.Text.Encoding.GetEncoding("utf-8"))
            '使用自动缩进便于阅读
            writer.Formatting = Xml.Formatting.Indented
            writer.WriteRaw("<?xml version=""1.0"" encoding=""utf-8"" ?>")
            '书写根元素()
            writer.WriteStartElement("Config")
            '添加次级元素
            writer.WriteStartElement("DatabaseSetting")
            '添加子元素()
            writer.WriteElementString("DataSource", "1")
            writer.WriteElementString("InitialCatalog", "2")
            writer.WriteElementString("UserID", "3")
            writer.WriteElementString("Password", "4")
            '关闭次级元素DatabaseSetting
            writer.WriteEndElement()
            '关闭根元素
            writer.WriteFullEndElement()
            '将XML写入文件并关闭writer
            writer.Close()
        Catch ex As Exception
            MsgBox(ex.Message & vbCrLf & ex.StackTrace)
        End Try
    End Sub

    Private Sub SaveXML()
        Dim Dir As String = System.Reflection.Assembly.GetExecutingAssembly.GetName.CodeBase
        Dir = Dir.Substring(0, Dir.LastIndexOf("\")) & "\Set.XML"

        Dim Arr1(3) As String
        Dim Arr2(3) As String

        Arr1(0) = "DataSource"
        Arr1(1) = "InitialCatalog"
        Arr1(2) = "UserID"
        Arr1(3) = "Password"

        Arr2(0) = SQLServer
        Arr2(1) = SQLDatabase
        Arr2(2) = SQLUser
        Arr2(3) = SQLPassword

        modifXML(Dir, "Config", Arr1, Arr2)

    End Sub

    Private Sub SaveXMLK3()
        Dim Dir As String = System.Reflection.Assembly.GetExecutingAssembly.GetName.CodeBase
        Dir = Dir.Substring(0, Dir.LastIndexOf("\")) & "\SetK3.XML"

        Dim Arr1(3) As String
        Dim Arr2(3) As String

        Arr1(0) = "DataSource"
        Arr1(1) = "InitialCatalog"
        Arr1(2) = "UserID"
        Arr1(3) = "Password"

        Arr2(0) = SQLServerK3
        Arr2(1) = SQLDatabaseK3
        Arr2(2) = SQLUserK3
        Arr2(3) = SQLPasswordK3

        modifXMLK3(Dir, "Config", Arr1, Arr2)

    End Sub

    Public Sub modifXML(ByVal xmlFileName As String, ByVal rootName As String, ByVal elementNameArry() As String, ByVal innerTextArry() As String)
        Try


            Dim doc As New Xml.XmlDocument
            doc.Load(xmlFileName)
            Dim list As Xml.XmlNodeList
            list = doc.SelectSingleNode(rootName).ChildNodes

            For Each xn As Xml.XmlNode In list
                Dim xe As Xml.XmlElement
                xe = xn
                Dim nls As Xml.XmlNodeList = xe.ChildNodes
                For Each xn1 As Xml.XmlNode In nls
                    Dim xe2 As Xml.XmlElement
                    xe2 = xn1
                    For i As Integer = 0 To elementNameArry.Length - 1
                        If xe2.Name = elementNameArry(i) Then
                            xe2.InnerText = innerTextArry(i)
                        End If
                    Next
                Next
            Next
            doc.Save(xmlFileName)
        Catch ex As Exception
            MsgBox(ex.Message & vbCrLf & ex.StackTrace)
        End Try
    End Sub

    Public Sub modifXMLK3(ByVal xmlFileName As String, ByVal rootName As String, ByVal elementNameArry() As String, ByVal innerTextArry() As String)
        Try
            Dim doc As New Xml.XmlDocument
            doc.Load(xmlFileName)
            Dim list As Xml.XmlNodeList
            list = doc.SelectSingleNode(rootName).ChildNodes

            For Each xn As Xml.XmlNode In list
                Dim xe As Xml.XmlElement
                xe = xn
                Dim nls As Xml.XmlNodeList = xe.ChildNodes
                For Each xn1 As Xml.XmlNode In nls
                    Dim xe2 As Xml.XmlElement
                    xe2 = xn1
                    For i As Integer = 0 To elementNameArry.Length - 1
                        If xe2.Name = elementNameArry(i) Then
                            xe2.InnerText = innerTextArry(i)
                        End If
                    Next
                Next
            Next
            doc.Save(xmlFileName)
        Catch ex As Exception
            MsgBox(ex.Message & vbCrLf & ex.StackTrace)
        End Try
    End Sub


    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Me.Dispose()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        'BoxMessage()
        ShowPanel(8)
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        GetCombo("作业班组")
        If flag = 1 Then Exit Sub
        Dim StrErr As String
        Dim Arr(,)
        StrErr = GetRst("select convert(varchar(20),getdate(),120)", Arr, SQL)
        If StrErr <> "" Then MsgBox("获取时间失败！！") : Exit Sub
        TextBox8.Text = Arr(1, 1)
        TextBox6.Text = ""
        ShowPanel(4)
    End Sub '人员登录

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        SQLServer = TextBox1.Text
        SQLDatabase = TextBox2.Text
        SQLUser = TextBox3.Text
        SQLPassword = TextBox4.Text
        Try
            If SQL.State <> Data.ConnectionState.Closed Then SQL.Close()
        Catch ex As Exception
        End Try
        SQL.ConnectionString = "server=" & SQLServer & ";database=" & SQLDatabase & ";user id=" & SQLUser & ";pwd=" & SQLPassword & ""
        SaveXML()
        MsgBox("保存成功！")
        ShowPanel(3)

    End Sub '保存数据库设定


    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        ShowPanel(11)
    End Sub

    Private Sub Button12_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        TextBox5.Visible = False
        ShowPanel(1)
    End Sub

    Function GetInBar(ByRef StrShow As String)
        Dim LogStr As String = ""
        GetInBar = ""
        If Val(BoxCode) < 99979 Or Val(BoxCode) > 99999 Then
            Dim Arr6(,)
            Dim StrErr6 As String
            StrErr6 = GetRst("select top 1 barcode from hand_store where StoreState='在库' and boxcode='" & BoxCode & "'", Arr6, SQL)
            If StrErr6 <> "" Then GetInBar = StrErr6 : Exit Function

            If UBound(Arr6, 2) > 0 Then
                ''''2016-05-26  无需判断重复，因为新规格没有重复的，已经在上步加以判断
                'Dim Arr1(,)
                'Dim StrErr1 As String
                'StrErr1 = GetRst("select barcode from hand_store where StoreState='在库' and  barcode ='" & TextBox7.Text.Trim & "' and boxcode='" & BoxCode & "'", Arr1, SQL)
                'If StrErr1 <> "" Then GetInBar = StrErr1 : Exit Function
                'If UBound(Arr1, 2) > 0 Then GetInBar = "同一笼框条码" & TextBox7.Text.Trim & "，条码不能重复！！" : Exit Function

                Dim Arr2(,)
                Dim StrErr2 As String
                ''''2016-05-26 新条码规则改变，前两位和后四位为流水号，3-5为规格，6-7为品牌   改动
                'StrErr2 = GetRst("select barcode from hand_store where StoreState='在库' and  barcode like '" & Mid(TextBox7.Text.Trim, 1, 7) & "%' and boxcode='" & BoxCode & "'", Arr2, SQL)

                '2016-08-08 10位条码
                'StrErr2 = GetRst("select top 1 barcode from hand_store where StoreState='在库' and  barcode like '__" & Mid(TextBox7.Text.Trim, 3, 5) & "%' and boxcode='" & BoxCode & "'", Arr2, SQL)
                StrErr2 = GetRst("select top 1 barcode from hand_store where StoreState='在库' and  ProductID = " & FItemID & " and boxcode='" & BoxCode & "'", Arr2, SQL)

                If StrErr2 <> "" Then GetInBar = StrErr2 : Exit Function
                If UBound(Arr2, 2) <= 0 Then GetInBar = "同一笼框条码,规格必须一致！！" : Exit Function

            End If
        End If


        Dim SQLStr() As String
        ReDim SQLStr(1)
        Dim StrErr4 As String
        SQLStr(1) = "UPDATE KTMSSQL.dbo.sync_bcwl SET kq_nm =" & FaultLoc & " WHERE barcode = '" & TextBox7.Text.Trim & "'"
        StrErr4 = ExeSQLS(SQLStr, SQLMES)
        If StrErr4 <> "" Then GetInBar = StrErr4 : Exit Function

        Dim StrErr5 As String
        Dim SQL1() As String
        ReDim SQL1(1)
        SQL1(1) = "insert into hand_store (StoreState,InClass,ProductID,inman,intime,indate,oldrk_time,oldrk_class,oldrk_man,barcode,boxcode,flag,FaultLoc,instore_type) values " & _
                        "('在库','" & NowClass & "'," & FItemID & ",'" & NowUser & "',convert(datetime,convert(varchar(20),getdate(),120)),convert(varchar(10),getdate(),120),convert(datetime,convert(varchar(20),getdate(),120)),'" & NowClass & "','" & NowUser & "','" & TextBox7.Text.Trim & "','" & BoxCode & "',0," & FaultLoc & "," & QtInType & ")"

        StrErr5 = ExeSQLS(SQL1, SQL)
        If StrErr5 <> "" Then MsgBox(StrErr5) : Exit Function
        StrShow = "扫码成功！！"
    End Function

    Function GetQtInBar(ByRef StrShow As String)

        Dim SQLStr() As String
        ReDim SQLStr(1)
        Dim StrErr4 As String
        SQLStr(1) = "UPDATE KTMSSQL.dbo.sync_bcwl SET kq_nm =" & FaultLoc & " WHERE barcode = '" & TextBox28.Text.Trim & "'"
        StrErr4 = ExeSQLS(SQLStr, SQLMES)
        If StrErr4 <> "" Then GetQtInBar = StrErr4 : Exit Function

        Dim StrErr5 As String
        Dim SQL1() As String
        ReDim SQL1(1)

        SQL1(1) = "insert into hand_store (StoreState,InClass,ProductID,inman,incode,intime,indate,oldrk_time,oldrk_class,oldrk_man,barcode,boxcode,flag,FaultLoc,instore_type,k3_time) values " & _
                        "('在库','" & NowClass & "'," & FItemID & ",'" & NowUser & "'," & FInterID & ",convert(datetime,convert(varchar(20),getdate(),120)),convert(varchar(10),getdate(),120),convert(datetime,convert(varchar(20),getdate(),120)),'" & NowClass & "','" & NowUser & "','" & TextBox28.Text.Trim & "','" & BoxCode & "',1," & FaultLoc & "," & QtInType & ",convert(datetime,convert(varchar(20),getdate(),120)))"

        StrErr5 = ExeSQLS(SQL1, SQL)
        If StrErr5 <> "" Then MsgBox(StrErr5) : Exit Function
        StrShow = "扫码成功,已上传金蝶！！"
    End Function

    '2018-6-2增加如果入库条码是【入库取消】的话，直接修改入库标志即可
    Function updateQtInBar(ByRef StrShow As String)

        Dim StrErr5 As String
        Dim SQL1() As String
        ReDim SQL1(1)

        SQL1(1) = "update hand_store set incode=" & FInterID & ",k3_time=convert(datetime,convert(varchar(20),getdate(),120)),oldrk_time=intime,oldrk_class=InClass,oldrk_man=inman,StoreState='在库',InClass='" & NowClass & "',inman='" & NowUser & "',intime=convert(datetime,convert(varchar(20),getdate(),120)),indate=convert(varchar(10),getdate(),120),boxcode='" & BoxCode & "',flag=1,instore_type=" & QtInType & " where barcode='" & TextBox7.Text.Trim & "'"

        StrErr5 = ExeSQLS(SQL1, SQL)
        If StrErr5 <> "" Then MsgBox(StrErr5) : Exit Function
        StrShow = "扫码成功,已上传金蝶！！"
    End Function

    '2018-4-16增加如果入库条码是【入库取消】的话，直接修改入库标志即可
    Function updateInBar(ByRef StrShow As String)
        Dim LogStr As String = ""
        updateInBar = ""
        If Val(BoxCode) < 99979 Or Val(BoxCode) > 99999 Then
            Dim Arr6(,)
            Dim StrErr6 As String
            StrErr6 = GetRst("select top 1 barcode from hand_store where StoreState='在库' and boxcode='" & BoxCode & "'", Arr6, SQL)
            If StrErr6 <> "" Then updateInBar = StrErr6 : Exit Function

            If UBound(Arr6, 2) > 0 Then
                Dim Arr2(,)
                Dim StrErr2 As String
                StrErr2 = GetRst("select top 1 barcode from hand_store where StoreState='在库' and  ProductID = " & FItemID & " and boxcode='" & BoxCode & "'", Arr2, SQL)
                If StrErr2 <> "" Then updateInBar = StrErr2 : Exit Function
                If UBound(Arr2, 2) <= 0 Then updateInBar = "同一笼框条码,规格必须一致！！" : Exit Function
            End If
        End If

        Dim StrErr5 As String
        Dim SQL1() As String
        ReDim SQL1(1)

        SQL1(1) = "update hand_store set FaultLoc=" & FaultLoc & ",oldrk_time=intime,oldrk_class=InClass,oldrk_man=inman,StoreState='在库',InClass='" & NowClass & "',inman='" & NowUser & "',intime=convert(datetime,convert(varchar(20),getdate(),120)),indate=convert(varchar(10),getdate(),120),boxcode='" & BoxCode & "',flag=0,instore_type=" & QtInType & " where barcode='" & TextBox7.Text.Trim & "'"

        StrErr5 = ExeSQLS(SQL1, SQL)
        If StrErr5 <> "" Then MsgBox(StrErr5) : Exit Function
        StrShow = "扫码成功！！"
    End Function

    Function GetOutBar(ByRef StrShow As String)

        SetLog("准备扫码操作")
        Dim LogStr As String = ""
        GetOutBar = ""
        Dim StrErr2 As String
        Dim Arr1(,)
        Dim StrErr1 As String
        StrErr1 = GetRst("select top 1 barcode,ProductID,FaultLoc from hand_store where StoreState='在库' and  barcode ='" & TextBox9.Text & "' ", Arr1, SQL)
        If StrErr1 <> "" Then GetOutBar = StrErr1 : Exit Function
        If UBound(Arr1, 2) <= 0 Then GetOutBar = "条码" & TextBox9.Text & "未入库" : Exit Function

        '2016-10-18修改，出库也获取中间表规格 ---start
        'StrErr1 = GetMesFItemID(TextBox9.Text)
        'If StrErr1 <> "" Then GetOutBar = StrErr1 : Exit Function

        FItemID = Arr1(2, 1)
        If StockID <= 0 Then
            FaultLoc = Arr1(3, 1)
        Else
            FaultLoc = StockID
        End If

        '----------------------------------------end

        '''''''检查订单数量
        Dim flag As Long
        Dim FPrice As Decimal
        For p = 1 To UBound(ArrO, 2)

            If ArrO(6, p) = FItemID Then
                FUnitID = ArrO(7, p)
                DetailID = ArrO(1, p)
                flag = flag + 1
                FPrice = ArrO(8, p)
                FAllCount = ArrO(3, p)
                FAllEntryID = ArrO(9, p)
                FAllFInterID = ArrO(10, p)
                ArrO(5, p) = ArrO(5, p) + 1
                If ArrO(5, p) + ArrO(4, p) > ArrO(3, p) Then
                    GetOutBar = "规格数量超出预定量,无法出库！"
                    Exit Function
                End If
            End If
        Next

        If flag = 0 Then GetOutBar = "本订单没有此规格轮胎！" : Exit Function

        Dim Arr5(,)
        Dim StrErr5 As String
        StrErr5 = GetRst("select a.FBillNo, a.FInterID,a.FTranType,b.FEntryID from SEOrder a left join SEOrderEntry b on a.FInterID=b.FInterID where a.FBillNo = '" & FBillNo & "'", Arr5, SQLK3)
        If StrErr5 <> "" Then GetOutBar = StrErr5 : Exit Function
        If UBound(Arr5, 2) < 0 Then GetOutBar = "订单无此单号！！" : Exit Function
        GetOutBill() '生成出库单

        Dim Arr4(,)
        Dim StrErr4 As String
        StrErr4 = GetRst("select count(*) from ICStockBillEntry where FInterID = " & FInterID, Arr4, SQLK3)
        If StrErr4 <> "" Then GetOutBar = StrErr4 : Exit Function

        Dim Arr3(,)
        Dim StrErr3 As String
        Dim OutNo As Long
        StrErr3 = GetRst("select FQty,FEntryID from ICStockBillEntry where FInterID = " & FInterID & " and FItemID=" & FItemID, Arr3, SQLK3)
        If StrErr3 <> "" Then GetOutBar = StrErr3 : Exit Function
        If UBound(Arr3, 2) > 0 Then
            Dim SQL2() As String
            ReDim SQL2(1)
            OutNo = Arr3(2, 1)
            'SQL2(1) = "update ICStockBillEntry set FQty=" & (Arr3(1, 1) + 1) & ",FAuxQty=" & (Arr3(1, 1) + 1) & ",FAmount=" & (Arr3(1, 1) + 1) & "*FPrice,FConsignPrice= FPrice /" & FExchangeRate & ",FConsignAmount=" & (Arr3(1, 1) + 1) & " * FPrice  where FInterID = " & FInterID & " and FItemID=" & FItemID
            SQL2(1) = "update ICStockBillEntry set FOrderEntryID = " & FAllEntryID & ",FOrderBillNo= '" & FBillNo & "',FOrderInterID=" & FAllFInterID & ", FAuxQtyMust= " & FAllCount & ",FQty=" & (Arr3(1, 1) + 1) & ",FAuxQty=" & (Arr3(1, 1) + 1) & ",FConsignPrice= FPrice,FConsignAmount=" & (Arr3(1, 1) + 1) & " * FPrice  where FInterID = " & FInterID & " and FItemID=" & FItemID
            StrErr2 = ExeSQLS(SQL2, SQLK3)
            If StrErr2 <> "" Then MsgBox(StrErr2) : Exit Function
        Else
            Dim SQL2() As String
            ReDim SQL2(1)
            OutNo = Arr4(1, 1) + 1
            SQL2(1) = "insert into ICStockBillEntry (FBrNo,FInterID,FEntryID,FItemID,FQty,FAuxQty,FBatchNo,FSourceBillNo,FContractBillNo,FICMOBillNo,FMTONo,FClientOrderNo,FItemSize,FItemSuite,FPositionNo,FSEOutBillNo,FConfirmMemEntry,FReturnNoticeBillNO,FSourceEntryID,FSourceInterId,FSourceTranType,FNote,FUnitID,FDCSPID,FSnListID,FDCStockID,FChkPassItem,FPrice,FAuxPrice,FAmount,FConsignPrice,FConsignAmount,FAuxQtyMust,FOrderBillNo,FOrderEntryID,FOrderInterID) values ('0'," & FInterID & "," & (Arr4(1, 1) + 1) & "," & FItemID & ",1,1,'','" & Arr5(1, 1) & "','','','','','','','','','',''," & Arr5(4, 1) & "," & Arr5(2, 1) & "," & Arr5(3, 1) & ",''," & FUnitID & ",0,0," & FaultLoc & ",1058 ,Convert(decimal(18,10)," & FPrice & " * " & FExchangeRate & "),0,0,Convert(decimal(18,10)," & FPrice & " * " & FExchangeRate & "),Convert(decimal(18,10)," & FPrice & " * " & FExchangeRate & ")," & FAllCount & ",'" & FBillNo & "'," & FAllEntryID & "," & FAllFInterID & ")"
            StrErr2 = ExeSQLS(SQL2, SQLK3)
            If StrErr2 <> "" Then MsgBox(StrErr2) : Exit Function
        End If

        Dim SQL1() As String
        ReDim SQL1(1)
        SQL1(1) = "update hand_store set OutStockID=" & FaultLoc & ",vehicleID=" & vehicleID & ", OrderID=" & OrderID & ",DetailID=" & DetailID & ",outcode=" & FInterID & ",outno=" & OutNo & ", OutTime=convert(datetime,convert(varchar(20),getdate(),120)), OutMan='" & NowUser & "', OutClass='" & NowClass & "', OutDate=convert(varchar(10),getdate(),120),StoreState = '已出库' where id =(select top 1 id from hand_store where StoreState = '在库' and barcode = '" & TextBox9.Text & "') "
        StrErr1 = ExeSQLS(SQL1, SQL)
        If StrErr1 <> "" Then MsgBox(StrErr1) : Exit Function

        'AddOperLog(TextBox9.Text, "出库", OrderID, vehicleID, FInterID)

        SQL1(1) = "update vehicle set outNo='" & BillNo & "' where vehicleNo = '" & ComboBox1.Text.Trim & "' and orderID=" & OrderID
        StrErr1 = ExeSQLS(SQL1, SQL)
        If StrErr1 <> "" Then MsgBox(StrErr1) : Exit Function

        SQL1(1) = "update KTMSSQL.dbo.sync_bcwl set ck_no='" & BillNo & "',cph='" & ComboBox1.Text.Trim & "',ck_time=convert(varchar(10),getdate(),120),crkbz=2 where barcode='" & TextBox9.Text & "'"
        StrErr1 = ExeSQLS(SQL1, SQL)
        If StrErr1 <> "" Then MsgBox(StrErr1) : Exit Function

        StrShow = "扫码成功"
    End Function

    Function GetOutBoxBar(ByRef StrShow As String)
        SetLog("准备扫码操作")
        Dim LogStr As String = ""
        GetOutBoxBar = ""

        Dim StrErr2 As String

        Dim Arr1(,)
        Dim StrErr1 As String
        StrErr1 = GetRst("select barcode from hand_store where StoreState='在库' and  boxcode ='" & TextBox9.Text & "' ", Arr1, SQL)
        If StrErr1 <> "" Then GetOutBoxBar = StrErr1 : Exit Function
        If UBound(Arr1, 2) <= 0 Then GetOutBoxBar = "此笼框没有入库的轮胎，请换一个笼框！" : Exit Function
        BoxNum = UBound(Arr1, 2)

        '2016-08-07 增加10位条码 ---start
        'StrErr1 = GetFItemID(Arr1(1, 1))
        If TextBox9.Text.Length = 10 Then
            StrErr1 = GetMesFItemID(Arr1(1, 1))
        Else
            StrErr1 = GetFItemID(Arr1(1, 1))
        End If
        '2016-08-07 增加10位条码 --end

        If StrErr1 <> "" Then GetOutBoxBar = "网络连接失败，无法连接数据库！！" : Exit Function

        '''''''检查订单数量
        Dim flag As Long
        Dim FPrice As Decimal
        For p = 1 To UBound(ArrO, 2)

            If ArrO(6, p) = FItemID Then
                FUnitID = ArrO(7, p)
                DetailID = ArrO(1, p)
                flag = flag + 1
                FPrice = ArrO(8, p)
                FAllCount = ArrO(3, p)
                FAllEntryID = ArrO(9, p)
                FAllFInterID = ArrO(10, p)
                ArrO(5, p) = ArrO(5, p) + BoxNum
                If ArrO(5, p) + ArrO(4, p) > ArrO(3, p) Then
                    GetOutBoxBar = "规格数量超出预定量,无法出库，请单独出库！"
                    Exit Function
                End If
            End If
        Next

        If flag = 0 Then GetOutBoxBar = "本订单没有此规格轮胎！" : Exit Function


        Dim Arr5(,)
        Dim StrErr5 As String
        StrErr5 = GetRst("select a.FBillNo, a.FInterID,a.FTranType,b.FEntryID from SEOrder a left join SEOrderEntry b on a.FInterID=b.FInterID where a.FBillNo = '" & FBillNo & "'", Arr5, SQLK3)
        If StrErr5 <> "" Then GetOutBoxBar = StrErr5 : Exit Function
        If UBound(Arr5, 2) < 0 Then GetOutBoxBar = "订单无此单号！！" : Exit Function

        GetOutBill() '生成出库单

        Dim Arr4(,)
        Dim StrErr4 As String
        StrErr4 = GetRst("select count(*) from ICStockBillEntry where FInterID = " & FInterID, Arr4, SQLK3)
        If StrErr4 <> "" Then GetOutBoxBar = StrErr4 : Exit Function

        Dim Arr3(,)
        Dim StrErr3 As String
        Dim OutNo As Long
        StrErr3 = GetRst("select FQty,FEntryID from ICStockBillEntry where FInterID = " & FInterID & " and FItemID=" & FItemID, Arr3, SQLK3)
        If StrErr3 <> "" Then GetOutBoxBar = StrErr3 : Exit Function
        If UBound(Arr3, 2) > 0 Then
            Dim SQL2() As String
            ReDim SQL2(1)
            OutNo = Arr3(2, 1)
            SQL2(1) = "update ICStockBillEntry set FOrderEntryID = " & FAllEntryID & ",FOrderBillNo= '" & FBillNo & "',FOrderInterID=" & FAllFInterID & ", FAuxQtyMust = " & FAllCount & ",FQty=" & (Arr3(1, 1) + BoxNum) & ",FAuxQty=" & (Arr3(1, 1) + BoxNum) & ",FPrice=" & FPrice & ",FAuxPrice=0,FAmount=0,FConsignPrice=" & FPrice & ",FConsignAmount=" & (Arr3(1, 1) + BoxNum) & "*" & FPrice & " where FInterID = " & FInterID & " and FItemID=" & FItemID
            StrErr2 = ExeSQLS(SQL2, SQLK3)
            If StrErr2 <> "" Then MsgBox(StrErr2) : Exit Function
        Else
            Dim SQL2() As String
            ReDim SQL2(1)
            OutNo = Arr4(1, 1) + 1
            SQL2(1) = "insert into ICStockBillEntry (FBrNo,FInterID,FEntryID,FItemID,FQty,FAuxQty,FBatchNo,FSourceBillNo,FContractBillNo,FICMOBillNo,FMTONo,FClientOrderNo,FItemSize,FItemSuite,FPositionNo,FSEOutBillNo,FConfirmMemEntry,FReturnNoticeBillNO,FSourceEntryID,FSourceInterId,FSourceTranType,FNote,FUnitID,FDCSPID,FSnListID,FDCStockID,FChkPassItem,FPrice,FAuxPrice,FAmount,FConsignPrice,FConsignAmount,FAuxQtyMust,FOrderBillNo,FOrderEntryID,FOrderInterID) values ('0'," & FInterID & "," & (Arr4(1, 1) + 1) & "," & FItemID & "," & BoxNum & "," & BoxNum & ",'','" & Arr5(1, 1) & "','','','','','','','','','',''," & Arr5(4, 1) & "," & Arr5(2, 1) & "," & Arr5(3, 1) & ",''," & FUnitID & ",0,0," & FaultLoc & ",1058 ,Convert(decimal(18,10)," & FPrice & " * " & FExchangeRate & "),0,0,Convert(decimal(18,10)," & FPrice & " * " & FExchangeRate & "),Convert(decimal(18,10)," & BoxNum & "*" & FPrice & " * " & FExchangeRate & ")," & FAllCount & ",'" & FBillNo & "'," & FAllEntryID & "," & FAllFInterID & ")"
            StrErr2 = ExeSQLS(SQL2, SQLK3)
            If StrErr2 <> "" Then MsgBox(StrErr2) : Exit Function
        End If


        Dim SQL1() As String
        ReDim SQL1(1)

        SQL1(1) = "update hand_store set OrderID=" & OrderID & ",DetailID=" & DetailID & ",outcode=" & FInterID & ",outno=" & OutNo & ", OutTime=convert(datetime,convert(varchar(20),getdate(),120)), OutMan='" & NowUser & "', OutClass='" & NowClass & "', OutDate=convert(varchar(10),getdate(),120),StoreState = '已出库' where StoreState='在库' and boxcode = '" & TextBox9.Text & "' "

        LogStr = LogStr & vbCrLf & "扫码扫描:[" & TextBox9.Text & "]"
        StrErr1 = ExeSQLS(SQL1, SQL)
        If StrErr1 <> "" Then MsgBox(StrErr1) : Exit Function
        SetLog("扫码完成" & LogStr)
        StrShow = "扫码成功"

    End Function

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click

        Dim Arr(,)
        If TextBox18.Text = "" Then MsgBox("请选择人员") : Exit Sub
        If ComboBox5.Text = "" Then MsgBox("请选择班组") : Exit Sub
        If TextBox8.Text = "" Then MsgBox("请选择日期") : Exit Sub

        Dim StrErr As String
        StrErr = GetRst("select jobno,password,id,name,emp_id from sys_user", Arr, SQL)
        If StrErr <> "" Then MsgBox(StrErr) : Exit Sub
        Dim t As Long
        Dim Have As Boolean
        Dim Str1 As String
        Dim SQL1() As String
        For t = 1 To UBound(Arr, 2)
            If Arr(1, t) = TextBox18.Text And Arr(2, t) = TextBox6.Text Then
                NowUser = Arr(4, t)
                NowJobNo = TextBox18.Text
                NowClass = ComboBox5.Text
                NowDate = TextBox8.Text
                k3User = Arr(5, t)
                ReDim SQL1(1)
                SQL1(1) = "insert into loginlog (luser,logintime,address) values ('" & Arr(1, t) & Arr(4, t) & "',convert(varchar(20),getdate(),120), '手持端')"
                Str1 = ExeSQLS(SQL1, SQL)
                If Str1 <> "" Then MsgBox("数据库连接失败！！") : Exit Sub
                
                RefUserShow()
                Have = True
                ShowPanel(1)
                Exit For
            End If
        Next

        If Have = False Then MsgBox("密码不正确")

    End Sub  ''''人员登录

    Private Sub TextBox5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox5.TextChanged
        SetLog("开始扫描条码:" & TextBox5.Text)
        If TextBox5.Text = "" Then Exit Sub
        Dim Str As String
        Str = TextBox5.Text
        Str = Str.Replace(vbCr, "")
        Str = Str.Replace(vbLf, "")
        TextBox5.Text = ""
        Select Case PageNo
            Case 2
                TextBox9.Text = Str
                OutStore()
            Case 5
                TextBox7.Text = Str
                InStore()
            Case 8
                TextBox13.Text = Str
                BoxCode = Str
                BoxMessage()
            Case 9
                TextBox21.Text = Str
                InBoxMessage()
            Case 13
                TextBox20.Text = Str
                ChageBox()
            Case 14
                TextBox22.Text = Str
                CancelBarcode()
            Case 15
                TextBox23.Text = Str
                BarcodeMessage()
            Case 16
                TextBox24.Text = Str
                OldBoxMessage()
            Case 17
                TextBox25.Text = Str
                OldInStore()
            Case 18
                TextBox26.Text = Str
                CancelOut()
            Case 21
                TextBox27.Text = Str
                BoxCode = Str
                BoxMesMessage()
            Case 23
                TextBox28.Text = Str
                QtInStore()
            Case 24
                TextBox29.Text = Str
                BoxCode = Str
                QtInMessage()
            Case 26
                TextBox30.Text = Str
                ReturnOutStore()
        End Select
    End Sub

    Sub ReturnOutStore()
        If TextBox30.Text.Length <> 10 And TextBox30.Text.Length <> 11 Then MsgBox("请扫描10位或11位条码！！") : Exit Sub

        Dim StrErr As String = ""
        Dim StrShow As String = ""
        If TextBox30.Text.Length = 10 Or TextBox30.Text.Length = 11 Then
            StrErr = GetReturnOutBar(StrShow)
        End If

        If StrErr <> "" Then MsgBox(StrErr) : Exit Sub
        GetRdataByProductId(FItemID)
    End Sub

    Sub GetRdataByProductId(ByVal FInID As Long)  '''''''''获取单据信息

        Dim StrErr As String
        Dim Arr(,)
        Dim Arr2(,)

        Dim StrErr2 As String
        StrErr2 = GetRst("select FInterID,FCustID from SEOutStock where FBillNo ='" & ComboBox2.Text & "'", Arr2, SQLK3)
        If StrErr2 <> "" Then MsgBox("数据库连接失败！！") : Exit Sub
        If UBound(Arr2, 2) < 1 Then MsgBox("没有该订单信息！！") : Exit Sub
        FBillNo = ComboBox18.Text
        OrderID = Arr2(1, 1)
        FCustID = Arr2(2, 1)
        StrErr = Me.GetRst("select a.FDetailID,b.Fname,a.FAuxQty,0,0 ,b.FItemID,a.FUnitID,a.FPrice,a.FEntryID,a.FInterID from SEOutStockEntry a left join t_icitem b on a.FItemID=b.FItemID where a.FInterID =" & OrderID & " and a.FItemID=" & FInID, ArrO, SQLK3)

        If StrErr <> "" Then MsgBox("获取订单信息错误" & vbCrLf & StrErr) : Exit Sub
        If UBound(ArrO, 2) = 0 Then MsgBox("无订单信息") : Exit Sub

        StrErr = Me.GetRst("select count(*) from hand_store where storestate = '在库' and instore_type =4 and OrderID = " & OrderID & " and ProductID=" & FInID, Arr, SQL)

        If StrErr <> "" Then MsgBox("获取发货信息错误" & vbCrLf & StrErr) : Exit Sub
        Dim Have As Boolean
        Dim NoneCount As Long = 0
        Dim AllCount As Long
        Dim OutCount As Long

        For p = 1 To UBound(ArrO, 2)
            If ArrO(6, p) = FInID Then
                Have = True
                ArrO(4, p) = Arr(1, 1)
            End If
        Next

        For t = 1 To UBound(ArrO, 2)
            AllCount = AllCount + ArrO(3, t)
            OutCount = OutCount + ArrO(4, t)
        Next


        Dim ds As New DataSet
        Dim dt As New DataTable
        dt.Columns.Add("规格型号", Type.GetType("System.String"))
        dt.Columns.Add("应扫", Type.GetType("System.Int32"))
        dt.Columns.Add("实扫", Type.GetType("System.Int32"))
        For t = 1 To UBound(ArrO, 2)
            Dim dw = dt.NewRow
            dw.Item(0) = ArrO(2, t)
            dw.Item(1) = ArrO(3, t) - 0
            dw.Item(2) = ArrO(4, t)
            dt.Rows.Add(dw)

        Next
        ds.Tables.Add(dt)
        DataGrid5.DataSource = ds.Tables(0)

        '''''''''''''修改列宽

        DataGrid5.TableStyles.Clear()
        DataGrid5.TableStyles.Add(New DataGridTableStyle)
        DataGrid5.TableStyles.Item(0).MappingName = dt.TableName
        DataGrid5.TableStyles(0).GridColumnStyles.Item(0).Width = 140
        DataGrid5.TableStyles(0).GridColumnStyles.Item(1).Width = 35
        DataGrid5.TableStyles(0).GridColumnStyles.Item(2).Width = 35

    End Sub

    Function GetReturnOutBar(ByRef StrShow As String)
 
        Dim LogStr As String = ""
        GetReturnOutBar = ""
        Dim StrErr2 As String
        Dim Arr1(,)
        Dim StrErr1 As String
        StrErr1 = GetRst("select top 1 barcode,ProductID,FaultLoc from hand_store where StoreState='已出库' and  barcode ='" & TextBox30.Text & "' ", Arr1, SQL)
        If StrErr1 <> "" Then GetReturnOutBar = StrErr1 : Exit Function
        If UBound(Arr1, 2) <= 0 Then GetReturnOutBar = "条码" & TextBox30.Text & "未出库" : Exit Function

        FItemID = Arr1(2, 1)
        If StockID <= 0 Then
            FaultLoc = Arr1(3, 1)
        Else
            FaultLoc = StockID
        End If
        '----------------------------------------end

        '''''''检查订单数量
        Dim flag As Long
        Dim FPrice As Decimal
        For p = 1 To UBound(ArrO, 2)

            If ArrO(6, p) = FItemID Then
                FUnitID = ArrO(7, p)
                DetailID = ArrO(1, p)
                flag = flag + 1
                FPrice = ArrO(8, p)
                FAllCount = ArrO(3, p)
                FAllEntryID = ArrO(9, p)
                FAllFInterID = ArrO(10, p)
                ArrO(5, p) = ArrO(5, p) + 1
                If ArrO(5, p) + ArrO(4, p) > ArrO(3, p) Then
                    GetReturnOutBar = "规格数量超出预定量,无法出库！"
                    Exit Function
                End If
            End If
        Next

        If flag = 0 Then GetReturnOutBar = "本订单没有此规格轮胎！" : Exit Function
        Dim Arr5(,)
        Dim StrErr5 As String
        StrErr5 = GetRst("select a.FBillNo, a.FInterID,a.FTranType,b.FEntryID from SEOutStock a left join SEOutStockEntry b on a.FInterID=b.FInterID where a.FBillNo = '" & FBillNo & "'", Arr5, SQLK3)
        If StrErr5 <> "" Then GetReturnOutBar = StrErr5 : Exit Function
        If UBound(Arr5, 2) < 0 Then GetReturnOutBar = "订单无此单号！！" : Exit Function
        GetReturnOutBill() '生成出库单

        Dim Arr4(,)
        Dim StrErr4 As String
        StrErr4 = GetRst("select count(*) from ICStockBillEntry where FInterID = " & FInterID, Arr4, SQLK3)
        If StrErr4 <> "" Then GetReturnOutBar = StrErr4 : Exit Function

        Dim Arr3(,)
        Dim StrErr3 As String
        Dim OutNo As Long
        StrErr3 = GetRst("select FQty,FEntryID from ICStockBillEntry where FInterID = " & FInterID & " and FItemID=" & FItemID, Arr3, SQLK3)
        If StrErr3 <> "" Then GetReturnOutBar = StrErr3 : Exit Function
        If UBound(Arr3, 2) > 0 Then
            Dim SQL2() As String
            ReDim SQL2(1)
            OutNo = Arr3(2, 1)
            SQL2(1) = "update ICStockBillEntry set FOrderEntryID = " & FAllEntryID & ",FOrderBillNo= '" & FBillNo & "',FOrderInterID=" & FAllFInterID & ", FAuxQtyMust= " & FAllCount & ",FQty=" & (Arr3(1, 1) - 1) & ",FAuxQty=" & (Arr3(1, 1) - 1) & ",FConsignPrice= FPrice,FConsignAmount=" & (Arr3(1, 1) + 1) & " * FPrice  where FInterID = " & FInterID & " and FItemID=" & FItemID
            StrErr2 = ExeSQLS(SQL2, SQLK3)
            If StrErr2 <> "" Then MsgBox(StrErr2) : Exit Function
        Else
            Dim SQL2() As String
            ReDim SQL2(1)
            OutNo = Arr4(1, 1) + 1
            SQL2(1) = "insert into ICStockBillEntry (FBrNo,FInterID,FEntryID,FItemID,FQty,FAuxQty,FBatchNo,FSourceBillNo,FContractBillNo,FICMOBillNo,FMTONo,FClientOrderNo,FItemSize,FItemSuite,FPositionNo,FSEOutBillNo,FConfirmMemEntry,FReturnNoticeBillNO,FSourceEntryID,FSourceInterId,FSourceTranType,FNote,FUnitID,FDCSPID,FSnListID,FDCStockID,FChkPassItem,FPrice,FAuxPrice,FAmount,FConsignPrice,FConsignAmount,FAuxQtyMust,FOrderBillNo,FOrderEntryID,FOrderInterID) values ('0'," & FInterID & "," & (Arr4(1, 1) + 1) & "," & FItemID & ",-1,-1,'','" & Arr5(1, 1) & "','','','','','','','','','',''," & Arr5(4, 1) & "," & Arr5(2, 1) & "," & Arr5(3, 1) & ",''," & FUnitID & ",0,0," & FaultLoc & ",1058 ,Convert(decimal(18,10)," & FPrice & " * " & FExchangeRate & "),0,0,Convert(decimal(18,10)," & FPrice & " * " & FExchangeRate & "),Convert(decimal(18,10)," & FPrice & " * " & FExchangeRate & ")," & FAllCount & ",'" & FBillNo & "'," & FAllEntryID & "," & FAllFInterID & ")"
            StrErr2 = ExeSQLS(SQL2, SQLK3)
            If StrErr2 <> "" Then MsgBox(StrErr2) : Exit Function
        End If

        Dim SQL1() As String
        ReDim SQL1(1)
        SQL1(1) = "update hand_store set StoreState='在库',instore_type=4,oldrk_time=intime,oldrk_class=InClass,oldrk_man=inman,OrderID='" & OrderID & "',InClass='" & NowClass & "',inman='" & NowUser & "',intime=convert(datetime,convert(varchar(20),getdate(),120)),indate=convert(varchar(10),getdate(),120) where id=(select top 1 id from hand_store where Barcode='" & TextBox30.Text.Trim & "' and StoreState = '已出库' order by InTime desc)"
 
        StrErr1 = ExeSQLS(SQL1, SQL)
        If StrErr1 <> "" Then MsgBox(StrErr1) : Exit Function

        StrShow = "扫码成功"
    End Function


    Sub GetReturnOutBill() '生成取消出库单

        '2016-11-23 修改 汇率放在最上边
        Dim StrErr6 As String
        Dim Arr6(,)
        StrErr6 = GetRst("select FDeptID,FEmpID,FExchangeRate,FCurrencyID from SEOutStock where FBillNo='" & FBillNo & "'", Arr6, SQLK3)
        If UBound(Arr6, 2) < 0 Then MsgBox("该订单没有相关数据！！") : Exit Sub
        FExchangeRate = Arr6(3, 1) '获取汇率

        Dim StrErr7 As String
        Dim Arr7(,)
        StrErr7 = GetRst("select a.FInterID,a.FBillNo,a.FHeadSelfB0154 from ICStockBill a left join ICStockBillEntry b on a.FInterID=b.FInterID where b.FSourceBillNo='" & FBillNo & "'", Arr7, SQLK3)
        If UBound(Arr7, 2) > 0 Then
            FInterID = Arr7(1, 1)
            BillNo = Arr7(2, 1)
            Exit Sub
        End If

        '2016-11-23 修改 汇率放在最上边

        Dim StrErr1 As String
        StrErr1 = GetBillNo()
        If StrErr1 <> "" Then MsgBox(StrErr1) : Exit Sub
        'Label61.Text = "出库单号：" & BillNo

        Dim StrErr3 As String
        StrErr3 = GetFInterID(SQLK3)
        If StrErr3 <> "" Then MsgBox(StrErr3) : Exit Sub
        '2016-11-16 修改 start
        StrErr6 = GetRst("select FDeptID,FEmpID,FExchangeRate,FCurrencyID from SEOutStock where FBillNo='" & FBillNo & "'", Arr6, SQLK3)
        If UBound(Arr6, 2) < 0 Then MsgBox("该订单没有相关数据！！") : Exit Sub
        '2016-11-16 修改 end
        Dim Str As String
        Dim SQL2() As String
        ReDim SQL2(1)
        SQL2(1) = "insert into ICStockBill (FBrNo,FTranType,FDate,FBillNo,FExplanation,FFetchAdd,FPOSName,FConfirmMem,FYearPeriod,FInterID,FDCStockID,FFManagerID,FSManagerID,FBillerID,FVchInterID,FUpStockWhenSave,FManageType,FDeptID,FEmpID,FSupplyID,FSaleStyle,FRelateBrID,FBrID,FSettleDate,FOrderAffirm,FConsignee,FReceiver,FHeadSelfB0154,FCurrencyID,FROB) values " & _
                  "('0'," & FBillID & ",convert(datetime,convert(varchar(20),getdate(),120)),'" & BillNo & "','','','','',''," & FInterID & "," & FaultLoc & "," & k3User & "," & k3User & ",16427,0,0,0," & Arr6(1, 1) & "," & Arr6(2, 1) & "," & FCustID & ",101,0,0,convert(datetime,convert(varchar(20),getdate(),120)),0,0,'',1," & Arr6(4, 1) & ",-1)"

        Str = ExeSQLS(SQL2, SQLK3)
        If Str <> "" Then MsgBox(Str) : Exit Sub

    End Sub

    Sub GetNoOutBill()  '''''''''生成出库单
        Dim StrSQL As String
        Dim StrErr As String
        StrSQL = "select FDCStockID,FDeptID,FEmpID,FSupplyID,FHeadSelfB0154,FCurrencyID from ICStockBill where FInterID=" & OldFInterID
        Dim Arr(,)
        StrErr = GetRst(StrSQL, Arr, SQLK3)
        If UBound(Arr, 2) <= 0 Then ShowCancelOutLabel("无法识别原始出库单！！", Color.Red) : Exit Sub

        Dim StrErr3 As String
        StrErr3 = GetFInterID(SQLK3)
        If StrErr3 <> "" Then ShowCancelOutLabel(StrErr3, Color.Red) : Exit Sub

        Dim Str As String
        Dim SQL1() As String
        ReDim SQL1(1)
        'SQL1(1) = "insert into ICStockBill (FBrNo,FTranType,FDate,FBillNo,FExplanation,FFetchAdd,FPOSName,FConfirmMem,FYearPeriod,FInterID,FDCStockID,FFManagerID,FSManagerID,FBillerID,FVchInterID,FUpStockWhenSave,FManageType,FROB,FDeptID) values " & _
        '                "('0'," & FBillID & ",convert(datetime,convert(varchar(20),getdate(),120)),'" & BillNo & "','','','','',''," & FInterID & "," & ck & "," & k3User & "," & k3User & ",16427 ,0,1,0,-1,2913)"

        SQL1(1) = "insert into ICStockBill (FBrNo,FTranType,FDate,FBillNo,FExplanation,FFetchAdd,FPOSName,FConfirmMem,FYearPeriod,FInterID,FDCStockID,FFManagerID,FSManagerID,FBillerID,FVchInterID,FUpStockWhenSave,FManageType,FDeptID,FEmpID,FSupplyID,FSaleStyle,FRelateBrID,FBrID,FSettleDate,FOrderAffirm,FConsignee,FReceiver,FHeadSelfB0154,FCurrencyID,FROB ) values " & _
                 "('0'," & FBillID & ",convert(datetime,convert(varchar(20),getdate(),120)),'" & BillNo & "','','','','',''," & FInterID & "," & Arr(1, 1) & ",'" & k3User & "','" & k3User & "',16427,0,0,0," & Arr(2, 1) & "," & Arr(3, 1) & "," & Arr(4, 1) & ",101,0,0,convert(datetime,convert(varchar(20),getdate(),120)),0,0,'',1," & Arr(6, 1) & ",-1)"

        Str = ExeSQLS(SQL1, SQLK3)
        If Str <> "" Then MsgBox(Str) : Exit Sub

    End Sub

    Function GetMesFItemID(ByVal Barcode As String)
        ''2016-08-07  修改---通过 mes中间表获取规格信息
        GetMesFItemID = ""

        Dim Arr2(,)
        Dim StrErr2 As String
        StrErr2 = GetRst("select b.wl_nm,b.wl_na,b.wl_kw from sync_bcwl a left join sync_wlxx b on a.wl_nm=b.wl_nm where a.barcode='" & Barcode.Trim & "'", Arr2, SQLMES)
        If StrErr2 <> "" Then GetMesFItemID = StrErr2 : Exit Function
        If UBound(Arr2, 2) = 0 Then GetMesFItemID = "不识别MES规格,请确认MES规格！！" : Exit Function
        If Arr2(1, 1) = "" Then GetMesFItemID = "MES无对应规格,请确认MES规格！！" : Exit Function

        Dim Arr1(,)
        Dim StrErr1 As String
        StrErr1 = GetRst("select a.FItemID from t_ICItem a where a.FItemID=" & Arr2(1, 1), Arr1, SQLK3)
        If StrErr1 <> "" Then GetMesFItemID = StrErr1 : Exit Function
        If UBound(Arr1, 2) = 0 Then GetMesFItemID = "K3规格发生修改，请联系管理员！！" : Exit Function
        If Arr1(1, 1) = "" Then GetMesFItemID = "K3规格发生修改，请联系管理员！！" : Exit Function

        FItemID = Arr2(1, 1)
        Product = Arr2(2, 1)

        If StockID <= 0 Then
            FaultLoc = Arr2(3, 1)
        Else
            FaultLoc = StockID
        End If

    End Function

    Sub InStore()
        If ComboBox4.Text = "" Then MsgBox("请选择库区！！") : Exit Sub
        If TextBox7.Text = "" Then ShowInLabel("请扫描条码", Color.Red) : Exit Sub
        ''''''格式判断
        If TextBox7.Text.Length <> 10 And TextBox7.Text.Length <> 11 Then ShowInLabel("请扫描10位或11位条码！", Color.Red) : Exit Sub

        '2016-08-07 增加10条码  --start
        If TextBox7.Text.Length = 10 Then
            Dim StrErr3 As String
            Dim Arr3(,)
            StrErr3 = GetRst("select * from sync_bcwl where barcode='" & TextBox7.Text.Trim & "'", Arr3, SQLMES)
            If StrErr3 <> "" Then ShowInLabel(StrErr3, Color.Red) : Exit Sub
            If UBound(Arr3, 2) <= 0 Then ShowInLabel("MES中不存在该条码，不允许入库！！", Color.Red) : Exit Sub
            '2016-08-07 增加10条码  --end
        End If
        Dim StrErr2 As String
        Dim Arr2(,)
        ''''2016-05-26 新规则不允许重复条码     改动
        ' StrErr2 = GetRst("select Barcode from hand_store where Barcode='" & TextBox7.Text & "' and InTime>=DATEADD(n,-30,convert(datetime,convert(varchar(20),getdate(),120))) and InTime<=convert(datetime,convert(varchar(20),getdate(),120)) and StoreState='在库'", Arr2, SQL)
        StrErr2 = GetRst("select top 1 Barcode from hand_store where StoreState='在库' and Barcode='" & TextBox7.Text & "'", Arr2, SQL)
        If StrErr2 <> "" Then ShowInLabel(StrErr2, Color.Red) : Exit Sub
        If UBound(Arr2, 2) > 0 Then ShowInLabel("该条码的轮胎已入库，不能重复入库！！", Color.Red) : Exit Sub
        StrErr2 = GetRst("select top 1 Barcode from hand_store where StoreState='已出库' and Barcode='" & TextBox7.Text & "'", Arr2, SQL)
        If StrErr2 <> "" Then ShowInLabel(StrErr2, Color.Red) : Exit Sub
        If UBound(Arr2, 2) > 0 Then ShowInLabel("该条码的轮胎已出库，不能重复入库！！", Color.Red) : Exit Sub
        Dim StrErr1 As String
        If TextBox7.Text.Length = 10 Then
            StrErr1 = GetMesFItemID(TextBox7.Text)
        Else
            StrErr1 = GetFItemID(TextBox7.Text)
        End If
        If StrErr1 <> "" Then ShowInLabel(StrErr1, Color.Red) : Exit Sub

        Dim StrErr As String
        Dim StrShow As String = ""
        '2018-4-16增加--如果条码的当前状态为【取消入库】的话，则直接修改条码状态即可；入库不是的话新增一条记录
        StrErr2 = GetRst("select top 1 Barcode from hand_store where StoreState='取消入库' and Barcode='" & TextBox7.Text & "'", Arr2, SQL)
        If StrErr2 <> "" Then ShowInLabel(StrErr2, Color.Red) : Exit Sub
        If UBound(Arr2, 2) > 0 Then
            StrErr = updateInBar(StrShow)
        Else
            StrErr = GetInBar(StrShow)
        End If
        If StrErr <> "" Then ShowInLabel(StrErr, Color.Red) : Exit Sub

        StrShow = "规格：" & vbCrLf & Product & vbCrLf & vbCrLf & "条码：" & TextBox7.Text & vbCrLf & vbCrLf & StrShow
        ShowInLabel(StrShow, Color.Green)

        Dim StrErr5 As String
        Dim Arr5(,)
        StrErr5 = GetRst("select count(*) from hand_store where StoreState='在库' and boxcode='" & TextBox13.Text & "'", Arr5, SQL)
        Label22.Text = Arr5(1, 1)
        ShowBoxLabel("当前容量： " & Arr5(1, 1), Color.Green)
    End Sub

    Sub ShowCount()
        Dim Arr(,)
        Dim StrErr As String
        Dim StrSql As String
        StrSql = "select 1,count(*) from hand_store where InClass='" & NowClass & "' and InDate=convert(varchar(10),getdate(),120) UNION select 2,count(*) from hand_store where InMan='" & NowUser & "' and InDate=convert(varchar(10),getdate(),120)"
        StrErr = GetRst(StrSql, Arr, SQL)
        If StrErr <> "" Then ShowInLabel(StrErr, Color.Red) : Exit Sub

        Label24.Text = Arr(2, 2)
        Label25.Text = Arr(2, 1)

    End Sub

    Sub GetOutBox(ByRef barcode As String)
        SetLog("准备扫码操作")
        Dim LogStr As String = ""
        Dim StrErr2 As String
        Dim Arr1(,)
        Dim StrErr1 As String
        StrErr1 = GetRst("select barcode from hand_store where StoreState='在库' and  barcode ='" & barcode & "' and boxcode='" & TextBox9.Text & "'", Arr1, SQL)
        If StrErr1 <> "" Then MsgBox(StrErr1) : Exit Sub
        'If UBound(Arr1, 2) <= 0 Then MsgBox("条码" & barcode & "未入库") : Exit Sub
        If UBound(Arr1, 2) <= 0 Then Exit Sub

        '2016-08-07 增加10位条码 ---start
        'StrErr2 = GetFItemID(barcode)
        If barcode.Length = 10 Then
            StrErr2 = GetMesFItemID(barcode)
        Else
            StrErr2 = GetFItemID(barcode)
        End If
        '2016-08-07 增加10位条码 ---end

        If StrErr2 <> "" Then MsgBox(StrErr2) : Exit Sub

        '''''''检查订单数量
        Dim flag As Long
        Dim FPrice As Decimal
        For p = 1 To UBound(ArrO, 2)

            If ArrO(6, p) = FItemID Then
                FUnitID = ArrO(7, p)
                DetailID = ArrO(1, p)
                flag = flag + 1
                FPrice = ArrO(8, p)
                FAllCount = ArrO(3, p)
                FAllEntryID = ArrO(9, p)
                FAllFInterID = ArrO(10, p)

                ArrO(5, p) = ArrO(5, p) + 1
                If ArrO(5, p) + ArrO(4, p) > ArrO(3, p) Then Exit Sub
                'If ArrO(5, p) + ArrO(4, p) > ArrO(3, p) Then
                '    MsgBox("规格数量超出预定量,无法出库！")
                '    Exit Sub
                'End If
            End If
        Next

        'If flag = 0 Then MsgBox("本订单没有此规格轮胎！") : Exit Sub
        If flag = 0 Then Exit Sub

        Dim Arr5(,)
        Dim StrErr5 As String
        StrErr5 = GetRst("select a.FBillNo, a.FInterID,a.FTranType,b.FEntryID from SEOrder a left join SEOrderEntry b on a.FInterID=b.FInterID where a.FBillNo = '" & FBillNo & "'", Arr5, SQLK3)
        If StrErr5 <> "" Then MsgBox(StrErr5) : Exit Sub
        If UBound(Arr5, 2) < 0 Then MsgBox("订单无此单号！！") : Exit Sub
        GetOutBill() '生成出库单

        Dim Arr4(,)
        Dim StrErr4 As String
        StrErr4 = GetRst("select count(*) from ICStockBillEntry where FInterID = " & FInterID, Arr4, SQLK3)
        If StrErr4 <> "" Then MsgBox(StrErr4) : Exit Sub

        Dim Arr3(,)
        Dim StrErr3 As String
        Dim OutNo As Long
        StrErr3 = GetRst("select FQty,FEntryID from ICStockBillEntry where FInterID = " & FInterID & " and FItemID=" & FItemID, Arr3, SQLK3)
        If StrErr3 <> "" Then MsgBox(StrErr3) : Exit Sub
        If UBound(Arr3, 2) > 0 Then
            Dim SQL2() As String
            ReDim SQL2(1)
            OutNo = Arr3(2, 1)
            SQL2(1) = "update ICStockBillEntry set FOrderEntryID = " & FAllEntryID & ",FOrderBillNo= '" & FBillNo & "',FOrderInterID=" & FAllFInterID & ",FAuxQtyMust = " & FAllCount & ", FQty=" & (Arr3(1, 1) + 1) & ",FAuxQty=" & (Arr3(1, 1) + 1) & ",FPrice=" & FPrice & ",FAuxPrice=0,FAmount=" & (Arr3(1, 1) + 1) & "*" & FPrice & ",FConsignPrice=" & FPrice & ",FConsignAmount=" & (Arr3(1, 1) + 1) & "*" & FPrice & " where FInterID = " & FInterID & " and FItemID=" & FItemID
            StrErr2 = ExeSQLS(SQL2, SQLK3)
            If StrErr2 <> "" Then MsgBox(StrErr2) : Exit Sub
        Else
            Dim SQL2() As String
            ReDim SQL2(1)
            OutNo = Arr4(1, 1) + 1
            SQL2(1) = "insert into ICStockBillEntry (FBrNo,FInterID,FEntryID,FItemID,FQty,FAuxQty,FBatchNo,FSourceBillNo,FContractBillNo,FICMOBillNo,FMTONo,FClientOrderNo,FItemSize,FItemSuite,FPositionNo,FSEOutBillNo,FConfirmMemEntry,FReturnNoticeBillNO,FSourceEntryID,FSourceInterId,FSourceTranType,FNote,FUnitID,FDCSPID,FSnListID,FDCStockID,FChkPassItem,FPrice,FAuxPrice,FAmount,FConsignPrice,FConsignAmount,FAuxQtyMust,FOrderBillNo,FOrderEntryID,FOrderInterID) values ('0'," & FInterID & "," & (Arr4(1, 1) + 1) & "," & FItemID & ",1,1,'','" & Arr5(1, 1) & "','','','','','','','','','',''," & Arr5(4, 1) & "," & Arr5(2, 1) & "," & Arr5(3, 1) & ",''," & FUnitID & ",0,0," & FaultLoc & ",1058 ,Convert(decimal(18,10)," & FPrice & " * " & FExchangeRate & "),0,0,Convert(decimal(18,10)," & FPrice & " * " & FExchangeRate & "),Convert(decimal(18,10)," & FPrice & " * " & FExchangeRate & ")," & FAllCount & ",'" & FBillNo & "'," & FAllEntryID & "," & FAllFInterID & ")"
            StrErr2 = ExeSQLS(SQL2, SQLK3)
            If StrErr2 <> "" Then MsgBox(StrErr2) : Exit Sub
        End If

        Dim SQL1() As String
        ReDim SQL1(1)
        SQL1(1) = "update hand_store set OrderID=" & OrderID & ",DetailID=" & DetailID & ",outcode=" & FInterID & ",outno=" & OutNo & ", OutTime=convert(datetime,convert(varchar(20),getdate(),120)), OutMan='" & NowUser & "', OutClass='" & NowClass & "', OutDate=convert(varchar(10),getdate(),120),StoreState = '已出库' where id =(select top 1 id from hand_store where StoreState = '在库' and barcode = '" & barcode & "' and boxcode='" & TextBox9.Text & "') "
        LogStr = LogStr & vbCrLf & "扫码扫描:[" & barcode & "]"
        StrErr1 = ExeSQLS(SQL1, SQL)
        If StrErr1 <> "" Then MsgBox(StrErr1) : Exit Sub
        SetLog("扫码完成" & LogStr)
    End Sub

    Sub OutStore()
        If ComboBox3.Text = "" Then MsgBox("请选择库区！！") : Exit Sub
        If ComboBox1.Text = "" Then MsgBox("请选择车牌号！！") : Exit Sub

        Dim Arr7(,)
        Dim StrErr7 As String
        StrErr7 = GetRst("select id from vehicle where vehicleNo = '" & ComboBox1.Text.Trim & "' and orderID=" & OrderID, Arr7, SQL)
        If StrErr7 <> "" Then MsgBox("网络无法连接!!") : Exit Sub
        If UBound(Arr7, 2) = 0 Then MsgBox("无法识别车牌号!!") : Exit Sub
        vehicleID = Arr7(1, 1)
        If TextBox9.Text.Length <> 10 And TextBox9.Text.Length <> 11 And TextBox9.Text.Length <> 5 Then MsgBox("请扫描5位、10位或11位条码！！") : Exit Sub

        Dim StrErr As String = ""
        Dim StrShow As String = ""
        If TextBox9.Text.Length = 10 Or TextBox9.Text.Length = 11 Then
            Dim StrErr2 As String
            Dim Arr2(,)
            StrErr2 = GetRst("select top 1 Barcode from hand_store where Barcode='" & TextBox9.Text & "' and OutTime>=DATEADD(n,-10,convert(datetime,convert(varchar(20),getdate(),120))) and OutTime<=convert(datetime,convert(varchar(20),getdate(),120)) and StoreState='已出库'", Arr2, SQL)
            If StrErr2 <> "" Then MsgBox(StrErr2) : Exit Sub
            If UBound(Arr2, 2) > 0 Then MsgBox("10分钟不能重复出库" & TextBox9.Text & "的条码！！") : Exit Sub
            StrErr = GetOutBar(StrShow)
        End If

        If TextBox9.Text.Length = 5 Then
            MsgBox("轮框号暂时未启用！！")
            'Dim StrErr2 As String
            'Dim Arr2(,)
            'StrErr2 = GetRst("select top 1 Barcode from hand_store where boxcode='" & TextBox9.Text & "' and OutTime>=DATEADD(n,-60,convert(datetime,convert(varchar(20),getdate(),120))) and OutTime<=convert(datetime,convert(varchar(20),getdate(),120)) and StoreState='已出库'", Arr2, SQL)
            'If StrErr2 <> "" Then MsgBox(StrErr2) : Exit Sub
            'If UBound(Arr2, 2) > 0 Then MsgBox("1小时内不能出库重复笼框的轮胎！！") : Exit Sub
            'If Val(TextBox9.Text.Trim) < 99979 Or Val(TextBox9.Text.Trim) > 99999 Then
            '    StrErr = GetOutBoxBar(StrShow)
            'Else
            '    Dim StrErr3 As String
            '    Dim Arr3(,)
            '    StrErr3 = GetRst("select Barcode from hand_store where boxcode='" & TextBox9.Text & "' and StoreState='在库'", Arr3, SQL)
            '    If StrErr3 <> "" Then MsgBox(StrErr3) : Exit Sub

            '    For t = 1 To UBound(Arr3, 2)
            '        GetOutBox(Arr3(1, t))
            '    Next

            'End If
        End If

        If StrErr <> "" Then MsgBox(StrErr) : Exit Sub
        GetOdataByProductId(FItemID)
        'GetOdata()
    End Sub

    Function GetFItemID(ByVal Barcode As String)
        ''2016-05-27  修改---金蝶规格表有“助记符”=“条码中间5位”
        GetFItemID = ""

        Dim Arr2(,)
        Dim StrErr2 As String
        ''2016-05-27  修改---金蝶规格表有“助记符”=“条码中间5位”
        'StrErr2 = GetRst("select FItemID,Fname,fdefaultloc from t_icitem where Fname like '" & Arr1(1, 1) & "%" & Arr1(3, 1) & "%" & Arr1(4, 1) & "%" & Arr1(2, 1) & "'", Arr2, SQLK3)

        StrErr2 = GetRst("select top 1 FItemID,Fname,fdefaultloc from t_icitem where FHelpCode = '" & Mid(Barcode.Trim, 3, 5) & "'", Arr2, SQLK3)
        If StrErr2 <> "" Then GetFItemID = StrErr2 : Exit Function
        If UBound(Arr2, 2) = 0 Then GetFItemID = "不识别K3规格,请确认K3规格！！" : Exit Function
        If Arr2(1, 1) = 0 Then GetFItemID = "不识别K3规格,请确认K3规格！！" : Exit Function
        FItemID = Arr2(1, 1)
        Product = Arr2(2, 1)
 
        If StockID <= 0 Then
            FaultLoc = Arr2(3, 1)
        Else
            FaultLoc = StockID
        End If
    End Function

    Sub BoxMessage()
        If TextBox13.Text.Length <> 5 Then ShowBoxLabel("请扫描5位条码！", Color.Red) : Exit Sub
        Dim StrErr As String
        Dim Arr(,)
        StrErr = GetRst("select barcode from hand_store where StoreState='在库' and boxcode='" & TextBox13.Text & "'", Arr, SQL)
        If StrErr <> "" Then
            SetLog(StrErr)
            MsgBox("连接数据库失败！！")
            Exit Sub
        End If
        BoxNum = UBound(Arr, 2)
        ShowBoxLabel("当前容量： " & BoxNum, Color.Green)
    End Sub

    Sub QtInMessage()
        If TextBox29.Text.Length <> 5 Then ShowBoxLabel("请扫描5位条码！", Color.Red) : Exit Sub
        Dim StrErr As String
        Dim Arr(,)
        StrErr = GetRst("select barcode from hand_store where StoreState='在库' and boxcode='" & TextBox29.Text & "'", Arr, SQL)
        If StrErr <> "" Then
            SetLog(StrErr)
            MsgBox("连接数据库失败！！")
            Exit Sub
        End If
        BoxNum = UBound(Arr, 2)
        ShowBoxLabel("当前容量： " & BoxNum, Color.Green)
    End Sub

    Sub BoxMesMessage()
        If TextBox27.Text.Length <> 5 Then ShowMesBoxLabel("请扫描5位条码！", Color.Red) : Exit Sub
        Dim StrErr As String
        Dim Arr(,)
        StrErr = GetRst("select barcode from hand_store where StoreState='在库' and boxcode='" & TextBox27.Text & "'", Arr, SQL)
        If StrErr <> "" Then
            MsgBox("连接数据库失败！！")
            Exit Sub
        End If
        BoxNum = UBound(Arr, 2)
        ShowMesBoxLabel("当前容量： " & BoxNum, Color.Green)
    End Sub

    Sub ShowMesBoxLabel(ByVal Str As String, ByVal CC As Color)

        Label61.Text = Str
        Label61.BackColor = CC
        'If CC = Color.Red Then
        '    MsgBox(Str)
        'End If
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        ShowPanel(15)
    End Sub

    Private Sub Button19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button19.Click

        Dim StrErr As String
        Dim Arr(,)

        StrErr = GetRst("select name,password from sys_user where jobno = '" & NowJobNo & "'", Arr, SQL)
        If StrErr <> "" Then MsgBox(StrErr) : Exit Sub

        If UBound(Arr, 2) = 0 Then MsgBox("人员不存在") : Exit Sub

        If Arr(2, 1) <> TextBox12.Text Then MsgBox("旧密码不正确") : Exit Sub
        If TextBox11.Text <> TextBox10.Text Then MsgBox("两次密码不同") : Exit Sub

        Dim SQL1() As String
        ReDim SQL1(1)
        SQL1(1) = "update sys_user set password = '" & TextBox11.Text.ToString & "' where jobno = '" & NowJobNo & "'"
        StrErr = ExeSQLS(SQL1, SQL)
        If StrErr <> "" Then MsgBox(StrErr) : Exit Sub

        MsgBox("密码修改成功")

        TextBox12.Text = ""
        TextBox11.Text = ""
        TextBox10.Text = ""

        ShowPanel(1)
    End Sub

    Private Sub Button17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button17.Click
        ShowPanel(10)
    End Sub

    Sub ShowInLabel(ByVal Str As String, ByVal CC As Color)

        Label35.Text = Str
        Label35.BackColor = CC
        'If CC = Color.Red Then 
        '    MsgBox(Str)
        'End If
    End Sub

    Sub ShowBoxLabel(ByVal Str As String, ByVal CC As Color)

        Label38.Text = Str
        Label38.BackColor = CC
        'If CC = Color.Red Then
        '    MsgBox(Str)
        'End If
    End Sub

    Sub ShowNoLabel(ByVal Str As String, ByVal CC As Color)

        Label16.Text = Str
        Label16.BackColor = CC
        'If CC = Color.Red Then
        '    MsgBox(Str)
        'End If
    End Sub

    Private Sub Button25_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button25.Click
        If TextBox7.Text = "" Then MsgBox("条码为空！！") : Exit Sub
        InStore()
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        readXMl()
        readXMlK3()
        readXMlMES()
        SQL.ConnectionString = "server=" & SQLServer & ";database=" & SQLDatabase & ";user id=" & SQLUser & ";pwd=" & SQLPassword & ""
        SQLK3.ConnectionString = "server=" & SQLServerK3 & ";database=" & SQLDatabaseK3 & ";user id=" & SQLUserK3 & ";pwd=" & SQLPasswordK3 & ""
        SQLMES.ConnectionString = "server=" & SQLServerMES & ";database=" & SQLDatabaseMES & ";user id=" & SQLUserMES & ";pwd=" & SQLPasswordMES & ""
        ShowPanel(1)
        Timer1.Enabled = False

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If NowUser = "" Then MsgBox("请先登录") : Exit Sub
        ShowPanel(7)
        Dim StrSQL As String
        Dim StrErr As String
        StrSQL = "select FBillID from ICBillNo where FBillName LIKE '销售出库%' "
        Dim Arr(,)
        StrErr = GetRst(StrSQL, Arr, SQLK3)
        FBillID = Arr(1, 1)
        GetOrder()
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        ShowPanel(7)
        GetOdata()
    End Sub

    Private Sub Button28_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button28.Click
        If ComboBox18.Text = "" Then MsgBox("请选择出库订单！！") : Exit Sub
        GetVehicleNo()
        GetStockList("out")
        ShowPanel(2)
    End Sub

    Sub GetVehicleNo()
        Dim t As Long
        Dim StrSQL As String
        Dim StrErr As String

        StrSQL = "select vehicleNo from vehicle where  orderID=" & OrderID
        ComboBox1.Items.Clear()

        StrErr = GetRst(StrSQL, ArrP, SQL)
        If StrErr <> "" Then
            MsgBox("网络连接失败！！")
            Exit Sub
        End If

        For t = 1 To UBound(ArrP, 2)
            If ArrP(1, t).ToString <> "" Then ComboBox1.Items.Add(ArrP(1, t).ToString)
        Next

    End Sub

    '获取库位
    Sub GetStockList(ByVal stockType As String)
        Dim t As Long
        Dim StrSQL As String
        Dim StrErr As String

        StrSQL = "select t.FName, t.FItemID from  t_stock t where t.FParentID =430 order by t.FNumber"
        StrErr = GetRst(StrSQL, stockArr, SQLK3)
        If StrErr <> "" Then
            MsgBox("网络连接失败！！")
            Exit Sub
        End If

        StockID = 0
        Select Case stockType
            Case "in"
                ComboBox4.Items.Clear()
                ComboBox4.Items.Add("") '默认为空
                For t = 1 To UBound(stockArr, 2)
                    If stockArr(1, t).ToString <> "" Then ComboBox4.Items.Add(stockArr(1, t).ToString)
                Next
            Case "out"
                ComboBox3.Items.Clear()
                ComboBox3.Items.Add("")
                For t = 1 To UBound(stockArr, 2)
                    If stockArr(1, t).ToString <> "" Then ComboBox3.Items.Add(stockArr(1, t).ToString)
                Next
            Case "elseIn"  '其它入库
                ComboBox6.Items.Clear()
                ComboBox6.Items.Add("")
                For t = 1 To UBound(stockArr, 2)
                    If stockArr(1, t).ToString <> "" Then ComboBox6.Items.Add(stockArr(1, t).ToString)
                Next
            Case "returnOut" '销售退货
                ComboBox7.Items.Clear()
                ComboBox7.Items.Add("")
                For t = 1 To UBound(stockArr, 2)
                    If stockArr(1, t).ToString <> "" Then ComboBox7.Items.Add(stockArr(1, t).ToString)
                Next

        End Select
    End Sub

    Private Sub Button27_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button27.Click
        ShowPanel(11)
    End Sub

    Sub GetOdata()  '''''''''获取单据信息

        Dim StrErr As String
        Dim Arr(,)
        Dim Arr2(,)

        Dim StrErr2 As String
        StrErr2 = GetRst("select FInterID,FCustID from SEOrder where FBillNo ='" & ComboBox18.Text & "'", Arr2, SQLK3)
        If StrErr2 <> "" Then MsgBox("数据库连接失败！！") : Exit Sub
        If UBound(Arr2, 2) < 1 Then MsgBox("没有该订单信息！！") : Exit Sub
        FBillNo = ComboBox18.Text
        OrderID = Arr2(1, 1)
        FCustID = Arr2(2, 1)
        StrErr = Me.GetRst("select a.FDetailID,b.Fname,a.FAuxQty,0,0 ,b.FItemID,a.FUnitID,a.FPrice,a.FEntryID,a.FInterID from SEOrderEntry a left join t_icitem b on a.FItemID=b.FItemID where a.FInterID =" & OrderID, ArrO, SQLK3)

        If StrErr <> "" Then MsgBox("获取订单信息错误" & vbCrLf & StrErr) : Exit Sub
        If UBound(ArrO, 2) = 0 Then MsgBox("无订单信息") : Exit Sub

        '2016-07-20 修改熟读
        ' StrErr = Me.GetRst("select Barcode,DetailID,ProductID from hand_store where storestate = '已出库' and OrderID = " & OrderID, Arr, SQL)
        StrErr = Me.GetRst("select ProductID,COUNT(*) from hand_store where storestate = '已出库' and OrderID = " & OrderID & " GROUP BY ProductID", Arr, SQL)

        If StrErr <> "" Then MsgBox("获取发货信息错误" & vbCrLf & StrErr) : Exit Sub
        'Dim Have As Boolean
        Dim NoneCount As Long = 0
        Dim AllCount As Long
        Dim OutCount As Long

        For t = 1 To UBound(Arr, 2)
            For p = 1 To UBound(ArrO, 2)
                If Arr(1, t) = ArrO(6, p) Then
                    'Have = True
                    ArrO(4, p) = Arr(2, t)
                End If
            Next
        Next


        For t = 1 To UBound(ArrO, 2)
            AllCount = AllCount + ArrO(3, t)
            OutCount = OutCount + ArrO(4, t)
        Next


        Dim ds As New DataSet
        Dim dt As New DataTable
        dt.Columns.Add("规格型号", Type.GetType("System.String"))
        dt.Columns.Add("应扫", Type.GetType("System.Int32"))
        dt.Columns.Add("实扫", Type.GetType("System.Int32"))
        For t = 1 To UBound(ArrO, 2)
            Dim dw = dt.NewRow
            dw.Item(0) = ArrO(2, t)
            dw.Item(1) = ArrO(3, t) - 0
            dw.Item(2) = ArrO(4, t)
            dt.Rows.Add(dw)

        Next
        ds.Tables.Add(dt)
        DataGrid1.DataSource = ds.Tables(0)
        'DataGrid2.DataSource = ds.Tables(0)

        '''''''''''''修改列宽

        DataGrid1.TableStyles.Clear()
        DataGrid1.TableStyles.Add(New DataGridTableStyle)
        DataGrid1.TableStyles.Item(0).MappingName = dt.TableName
        DataGrid1.TableStyles(0).GridColumnStyles.Item(0).Width = 140
        DataGrid1.TableStyles(0).GridColumnStyles.Item(1).Width = 35
        DataGrid1.TableStyles(0).GridColumnStyles.Item(2).Width = 35

    End Sub

    Sub GetOdataByProductId(ByVal FInID As Long)  '''''''''获取单据信息

        Dim StrErr As String
        Dim Arr(,)
        Dim Arr2(,)

        Dim StrErr2 As String
        StrErr2 = GetRst("select FInterID,FCustID from SEOrder where FBillNo ='" & ComboBox18.Text & "'", Arr2, SQLK3)
        If StrErr2 <> "" Then MsgBox("数据库连接失败！！") : Exit Sub
        If UBound(Arr2, 2) < 1 Then MsgBox("没有该订单信息！！") : Exit Sub
        FBillNo = ComboBox18.Text
        OrderID = Arr2(1, 1)
        FCustID = Arr2(2, 1)
        StrErr = Me.GetRst("select a.FDetailID,b.Fname,a.FAuxQty,0,0 ,b.FItemID,a.FUnitID,a.FPrice,a.FEntryID,a.FInterID from SEOrderEntry a left join t_icitem b on a.FItemID=b.FItemID where a.FInterID =" & OrderID & " and a.FItemID=" & FInID, ArrO, SQLK3)

        If StrErr <> "" Then MsgBox("获取订单信息错误" & vbCrLf & StrErr) : Exit Sub
        If UBound(ArrO, 2) = 0 Then MsgBox("无订单信息") : Exit Sub

        'StrErr = Me.GetRst("select Barcode,DetailID,ProductID from hand_store where storestate = '已出库' and OrderID = " & OrderID & " and ProductID=" & FInID, Arr, SQL)
        StrErr = Me.GetRst("select count(*) from hand_store where storestate = '已出库' and OrderID = " & OrderID & " and ProductID=" & FInID, Arr, SQL)

        If StrErr <> "" Then MsgBox("获取发货信息错误" & vbCrLf & StrErr) : Exit Sub
        Dim Have As Boolean
        Dim NoneCount As Long = 0
        Dim AllCount As Long
        Dim OutCount As Long
 
        For p = 1 To UBound(ArrO, 2)
            If ArrO(6, p) = FInID Then
                Have = True
                ArrO(4, p) = Arr(1, 1)
            End If
        Next

        For t = 1 To UBound(ArrO, 2)
            AllCount = AllCount + ArrO(3, t)
            OutCount = OutCount + ArrO(4, t)
        Next


        Dim ds As New DataSet
        Dim dt As New DataTable
        dt.Columns.Add("规格型号", Type.GetType("System.String"))
        dt.Columns.Add("应扫", Type.GetType("System.Int32"))
        dt.Columns.Add("实扫", Type.GetType("System.Int32"))
        For t = 1 To UBound(ArrO, 2)
            Dim dw = dt.NewRow
            dw.Item(0) = ArrO(2, t)
            dw.Item(1) = ArrO(3, t) - 0
            dw.Item(2) = ArrO(4, t)
            dt.Rows.Add(dw)

        Next
        ds.Tables.Add(dt)
        DataGrid2.DataSource = ds.Tables(0)

        '''''''''''''修改列宽

        DataGrid2.TableStyles.Clear()
        DataGrid2.TableStyles.Add(New DataGridTableStyle)
        DataGrid2.TableStyles.Item(0).MappingName = dt.TableName
        DataGrid2.TableStyles(0).GridColumnStyles.Item(0).Width = 140
        DataGrid2.TableStyles(0).GridColumnStyles.Item(1).Width = 35
        DataGrid2.TableStyles(0).GridColumnStyles.Item(2).Width = 35

    End Sub

    Sub GetOutBill()  '''''''''生成出库单

        '2016-11-23 修改 汇率放在最上边
        Dim StrErr6 As String
        Dim Arr6(,)
        StrErr6 = GetRst("select FDeptID,FEmpID,FExchangeRate,FCurrencyID from SEOrder where FBillNo='" & FBillNo & "'", Arr6, SQLK3)
        If UBound(Arr6, 2) < 0 Then MsgBox("该订单没有相关数据！！") : Exit Sub
        FExchangeRate = Arr6(3, 1) '获取汇率

        Dim StrErr7 As String
        Dim Arr7(,)
        'StrErr5 = GetRst("select a.FInterID,a.FBillNo,a.FHeadSelfB0154 from ICStockBill a left join ICStockBillEntry b on a.FInterID=b.FInterID where a.FDate >=convert(datetime,convert(varchar(10),getdate(),120)) and a.FDate < convert(datetime,convert(varchar(11),dateadd(day,1,getdate()),120)) and a.FCheckerID is null and b.FSourceBillNo='" & FBillNo & "'", Arr5, SQLK3)
        StrErr7 = GetRst("select outNo from vehicle where outNo is not null and outNo<>'' and vehicleNo = '" & ComboBox1.Text.Trim & "' and orderID=" & OrderID, Arr7, SQL)
        If UBound(Arr7, 2) > 0 Then
            Dim StrErr5 As String
            Dim Arr5(,)
            'StrErr5 = GetRst("select a.FInterID,a.FBillNo,a.FHeadSelfB0154 from ICStockBill a left join ICStockBillEntry b on a.FInterID=b.FInterID where a.FDate >=convert(datetime,convert(varchar(10),getdate(),120)) and a.FDate < convert(datetime,convert(varchar(11),dateadd(day,1,getdate()),120)) and a.FCheckerID is null and b.FSourceBillNo='" & FBillNo & "'", Arr5, SQLK3)
            StrErr5 = GetRst("select a.FInterID,a.FBillNo,a.FHeadSelfB0154 from ICStockBill a  where a.FCheckerID is null and a.FBillNo='" & Arr7(1, 1) & "'", Arr5, SQLK3)
            If UBound(Arr5, 2) > 0 Then
                FInterID = Arr5(1, 1)
                BillNo = Arr5(2, 1)
                'FExchangeRate = Arr5(3, 1) '获取汇率
                ShowPanel(2)
                Exit Sub
            End If
        End If

        '2016-11-23 修改 汇率放在最上边

        Dim StrErr1 As String
        StrErr1 = GetBillNo()
        If StrErr1 <> "" Then MsgBox(StrErr1) : Exit Sub
        'Label61.Text = "出库单号：" & BillNo

        Dim StrErr3 As String
        StrErr3 = GetFInterID(SQLK3)
        If StrErr3 <> "" Then MsgBox(StrErr3) : Exit Sub
        '2016-11-16 修改 start
        StrErr6 = GetRst("select FDeptID,FEmpID,FExchangeRate,FCurrencyID from SEOrder where FBillNo='" & FBillNo & "'", Arr6, SQLK3)
        If UBound(Arr6, 2) < 0 Then MsgBox("该订单没有相关数据！！") : Exit Sub
        '2016-11-16 修改 end
        Dim Str As String
        Dim SQL2() As String
        ReDim SQL2(1)
        SQL2(1) = "insert into ICStockBill (FBrNo,FTranType,FDate,FBillNo,FExplanation,FFetchAdd,FPOSName,FConfirmMem,FYearPeriod,FInterID,FDCStockID,FFManagerID,FSManagerID,FBillerID,FVchInterID,FUpStockWhenSave,FManageType,FDeptID,FEmpID,FSupplyID,FSaleStyle,FRelateBrID,FBrID,FSettleDate,FOrderAffirm,FConsignee,FReceiver,FHeadSelfB0154,FCurrencyID ) values " & _
                  "('0'," & FBillID & ",convert(datetime,convert(varchar(20),getdate(),120)),'" & BillNo & "','','','','',''," & FInterID & "," & FaultLoc & "," & k3User & "," & k3User & ",16427,0,0,0," & Arr6(1, 1) & "," & Arr6(2, 1) & "," & FCustID & ",101,0,0,convert(datetime,convert(varchar(20),getdate(),120)),0,0,'',1," & Arr6(4, 1) & ")"
       
        Str = ExeSQLS(SQL2, SQLK3)
        If Str <> "" Then MsgBox(Str) : Exit Sub

    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        TextBox5.Visible = False
        ShowPanel(1)
    End Sub

    Private Sub Button16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        TextBox5.Visible = False
        ShowPanel(1)
    End Sub

    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
        If TextBox13.Text = "" Then ShowBoxLabel("请扫描5位条码!", Color.Red) : Exit Sub
        If TextBox13.Text.Length <> 5 Then ShowBoxLabel("请扫描5位条码！", Color.Red) : Exit Sub

        BoxCode = TextBox13.Text.Trim
        Label22.Text = BoxNum
        Dim StrErr As String
        Dim Arr(,)
        StrErr = GetRst("select count(*) from hand_store where StoreState='在库' and boxcode='" & TextBox13.Text & "'", Arr, SQL)
        Label22.Text = Arr(1, 1)

        GetStockList("in")
        ShowPanel(5)
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If TextBox9.Text = "" Then Exit Sub
        OutStore()
    End Sub

    Private Sub Button18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        TextBox5.Visible = False
        ShowPanel(1)
    End Sub

    Private Sub Button21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        TextBox5.Visible = False
        ShowPanel(1)
    End Sub

    Private Sub Button20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        TextBox5.Visible = False
        ShowPanel(1)
    End Sub

    Function GetBillNo() As String
        Dim StrErr As String = ""
        Dim Ts As SqlTransaction
        Dim CMD As New SqlCommand
        GetBillNo = ""

        StrErr = ConSQL(SQLK3)
        If StrErr <> "" Then GetBillNo = StrErr : Exit Function

        Ts = SQLK3.BeginTransaction
        CMD.Transaction = Ts
        CMD.Connection = SQLK3

        Try
            CMD.CommandText = "p_BM_GetBillNo"
            CMD.CommandType = CommandType.StoredProcedure
            CMD.Parameters.Add("@ClassType", SqlDbType.Int).Value = FBillID
            CMD.Parameters.Add("@BillNo", SqlDbType.VarChar, 200).Direction = ParameterDirection.Output
            CMD.ExecuteNonQuery()
            Ts.Commit()
            BillNo = CMD.Parameters("@BillNo").Value '存储过程式的返回值
        Catch ex As Exception
            Ts.Rollback()
            GetBillNo = "执行数据失败"
        End Try
    End Function

    Function GetFInterID(ByRef SqlCon As SqlConnection) As String
        Dim StrErr As String = ""
        Dim Ts As SqlTransaction
        Dim CMD As New SqlCommand
        GetFInterID = ""

        StrErr = ConSQL(SqlCon)
        If StrErr <> "" Then GetFInterID = StrErr : Exit Function

        Ts = SqlCon.BeginTransaction
        CMD.Transaction = Ts
        CMD.Connection = SqlCon

        Try
            CMD.CommandText = "GetICMaxNum"
            CMD.CommandType = CommandType.StoredProcedure
            CMD.Parameters.Add("@Increment", SqlDbType.Int).Value = 1
            CMD.Parameters.Add("@UserID", SqlDbType.Int).Value = 16427
            CMD.Parameters.Add("@TableName", SqlDbType.VarChar, 200).Value = "ICStockBill"
            CMD.Parameters.Add("@FInterID", SqlDbType.VarChar, 200).Direction = ParameterDirection.Output
            CMD.ExecuteNonQuery()
            Ts.Commit()
            FInterID = CMD.Parameters("@FInterID").Value '存储过程式的返回值
        Catch ex As Exception
            Ts.Rollback()
            GetFInterID = "执行数据失败！！"
        End Try
    End Function

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        If TextBox13.Text = "" Then MsgBox("请输入条码！！") : Exit Sub
        BoxCode = TextBox13.Text
        BoxMessage()
    End Sub

    Private Sub Button34_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button34.Click
        TextBox5.Visible = False
        ShowPanel(1)
    End Sub

    Private Sub Button31_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button31.Click
        ShowPanel(11)
    End Sub

    Private Sub Button33_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button33.Click
        SQLServerK3 = TextBox17.Text
        SQLDatabaseK3 = TextBox16.Text
        SQLUserK3 = TextBox15.Text
        SQLPasswordK3 = TextBox14.Text
        Try
            If SQLK3.State <> Data.ConnectionState.Closed Then SQLK3.Close()
        Catch ex As Exception
        End Try
        SQLK3.ConnectionString = "server=" & SQLServerK3 & ";database=" & SQLDatabaseK3 & ";user id=" & SQLUserK3 & ";pwd=" & SQLPasswordK3 & ""
        SaveXMLK3()
        MsgBox("保存成功！")
        ShowPanel(12)
    End Sub

    Private Sub Button29_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button29.Click
        'ShowPanel(11) 数据库设置功能屏蔽掉
        If NowUser = "" Then MsgBox("请先登录") : Exit Sub
        ShowPanel(25)
        Dim StrSQL As String
        Dim StrErr As String
        StrSQL = "select FBillID from ICBillNo where FBillName LIKE '退货通知单据%' "
        Dim Arr(,)
        StrErr = GetRst(StrSQL, Arr, SQLK3)
        FBillID = Arr(1, 1)
        GetReturnOrder()
    End Sub

    Sub GetReturnOrder()
        Dim t As Long
        Dim StrSQL As String
        Dim StrErr As String

        StrSQL = "select t.FBillNo,t.FInterID from SEOutStock t WHERE t.FBillNo like 'SEIN%'"
        ComboBox2.Items.Clear()

        StrErr = GetRst(StrSQL, ArrP, SQLK3)
        If StrErr <> "" Then
            MsgBox("获取订单号失败" & vbCrLf & StrErr)
            Exit Sub
        End If

        For t = 1 To UBound(ArrP, 2)
            If ArrP(1, t).ToString <> "" Then ComboBox2.Items.Add(ArrP(1, t).ToString)
        Next

    End Sub

    Private Sub Button30_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If NowUser = "" Then MsgBox("请先登录") : Exit Sub
        ShowPanel(11)
    End Sub

    'Private Sub Button32_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    TextBox1.Text = SQLServer
    '    TextBox2.Text = SQLDatabase
    '    TextBox3.Text = SQLUser
    '    TextBox4.Text = SQLPassword
    '    ShowPanel(3)
    'End Sub

    Private Sub TextBox18_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox18.TextChanged
        TextBox19.ForeColor = Color.Red
        If TextBox18.Text.Length <> 4 Then TextBox19.Text = "工号为4位！！" : Exit Sub
        Dim Arr(,)
        Dim StrErr As String
        StrErr = GetRst("select name from sys_user where jobno ='" & TextBox18.Text & "'", Arr, SQL)
        If StrErr <> "" Then TextBox19.Text = "工号不存在！" : Exit Sub
        If UBound(Arr, 2) <= 0 Then TextBox19.Text = "工号不存在！" : Exit Sub
        TextBox19.Text = Arr(1, 1)
        TextBox19.ForeColor = Color.Black
    End Sub
    '
    Private Sub Button6_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        TextBox5.Visible = False
        ShowPanel(1)
    End Sub

    Private Sub Button18_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button18.Click
        TextBox5.Visible = False
        ShowPanel(19)
    End Sub

    Private Sub Button16_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button16.Click
        OutStore()
    End Sub

    Private Sub ComboBox18_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox18.SelectedIndexChanged
        GetOdata()
        TextBox9.Text = ""
        DataGrid2.DataSource = Nothing
    End Sub

    Private Sub ComboBox18_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox18.TextChanged

        If ComboBox18.Text.Length > 0 Then Exit Sub

        Dim Str As String

        ComboBox18.Items.Clear()

        Dim t As Long
        For t = 1 To UBound(ArrP, 2)
            Str = ArrP(1, t)
            If InStr(Str, ComboBox18.Text) Then
                ComboBox18.Items.Add(Str)
            End If
        Next

    End Sub

    Private Sub Button20_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button20.Click
        If NowUser = "" Then MsgBox("请先登录") : Exit Sub
        ShowPanel(10)
    End Sub

    Private Sub Button36_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button36.Click
        TextBox21.Text = ""
        Label40.Text = ""
        Label40.BackColor = Color.White
        ShowPanel(9)
    End Sub

    Private Sub Button21_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button21.Click
        TextBox5.Visible = False
        ShowPanel(1)
    End Sub

    Private Sub Button26_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button26.Click
        ShowPanel(10)
    End Sub

    Private Sub Button22_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button22.Click
        If TextBox21.Text = "" Then ShowBoxLabel("请扫描5位笼框条码!", Color.Red) : Exit Sub
        If TextBox21.Text.Length <> 5 Then ShowBoxLabel("请扫描5位笼框条码！", Color.Red) : Exit Sub

        TextBox20.Text = ""
        Label3.Text = ""
        Label3.BackColor = Color.White
        ShowPanel(13)
    End Sub

    Private Sub Button37_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button37.Click
        TextBox5.Visible = False
        ShowPanel(1)
    End Sub

    Private Sub Button39_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button39.Click
        ShowPanel(9)
        OutBoxMessage()
    End Sub

    Sub OutBoxMessage()
        If TextBox20.Text.Length <> 5 And TextBox20.Text.Length <> 10 And TextBox20.Text.Length <> 11 Then ShowBoxLabel("请扫描5位、10位或11位条码！", Color.Red) : Exit Sub
        Dim StrErr As String
        Dim Arr(,)
        If TextBox20.Text.Length = 5 Then
            StrErr = GetRst("select barcode from hand_store where StoreState='在库' and boxcode='" & TextBox20.Text.Trim & "'", Arr, SQL)
            If StrErr <> "" Then ShowOutBoxLabel("网络连接失败！获取获取数据库！！", Color.Red) : Exit Sub
            BoxNum = UBound(Arr, 2)
            ShowOutBoxLabel("当前容量： " & BoxNum, Color.Green)
            Exit Sub
        End If

        StrErr = GetRst("select boxcode from hand_store where StoreState='在库' and barcode='" & TextBox20.Text.Trim & "'", Arr, SQL)
        If StrErr <> "" Then ShowOutBoxLabel("网络连接失败！获取获取数据库！！", Color.Red) : Exit Sub
        BoxNum = UBound(Arr, 2)
        If BoxNum <= 0 Then ShowOutBoxLabel("该轮胎没有在库中！！ ", Color.Red) : Exit Sub
        ShowOutBoxLabel("该轮胎所在笼框： " & Arr(1, 1), Color.Green)

        ChageBox()
    End Sub

    Sub ShowOutBoxLabel(ByVal Str As String, ByVal CC As Color)
        Label3.Text = Str
        Label3.BackColor = CC
        'If CC = Color.Red Then
        '    MsgBox(Str)
        'End If
    End Sub

    Sub InBoxMessage()
        If TextBox21.Text.Length <> 5 Then ShowBoxLabel("请扫描5位条码！", Color.Red) : Exit Sub
        Dim StrErr As String
        Dim Arr(,)
        StrErr = GetRst("select barcode from hand_store where StoreState='在库' and boxcode='" & TextBox21.Text & "'", Arr, SQL)
        If StrErr <> "" Then
            SetLog(StrErr)
            MsgBox("连接数据库失败！！")
            Exit Sub
        End If
        BoxNum = UBound(Arr, 2)
        ShowInBoxLabel("笼框当前容量： " & BoxNum, Color.Green)
    End Sub

    Sub ShowInBoxLabel(ByVal Str As String, ByVal CC As Color)
        Label40.Text = Str
        Label40.BackColor = CC
        'If CC = Color.Red Then
        '    MsgBox(Str)
        'End If
    End Sub

    Private Sub Button23_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button23.Click
        ChageBox()
    End Sub

    Sub ChageBox()
        If TextBox20.Text = "" Then ShowOutBoxLabel("请输入5位或11位条码！！", Color.Red) : Exit Sub
        If TextBox20.Text.Length <> 5 And TextBox20.Text.Length <> 10 And TextBox20.Text.Length <> 11 Then ShowOutBoxLabel("请扫描5位或11位条码！", Color.Red) : Exit Sub
        Dim Arr1(,)
        Dim StrErr1 As String

        If TextBox20.Text.Length = 5 Then
            StrErr1 = GetRst("select ProductID from hand_store where StoreState ='在库' and boxcode= '" & TextBox20.Text.Trim & "'", Arr1, SQL)
            If StrErr1 <> "" Then ShowOutBoxLabel(StrErr1, Color.Red) : Exit Sub
            If UBound(Arr1, 2) <= 0 Then ShowOutBoxLabel("转出笼框" & TextBox20.Text.Trim & "内没有轮胎！！", Color.Red) : Exit Sub
        Else
            StrErr1 = GetRst("select ProductID from hand_store where StoreState ='在库' and barcode= '" & TextBox20.Text.Trim & "'", Arr1, SQL)
            If StrErr1 <> "" Then ShowOutBoxLabel(StrErr1, Color.Red) : Exit Sub
            If UBound(Arr1, 2) <= 0 Then ShowOutBoxLabel("转出轮胎" & TextBox20.Text.Trim & "没有在库中！！", Color.Red) : Exit Sub
        End If

        Dim Arr2(,)
        Dim StrErr2 As String
        StrErr2 = GetRst("select ProductID from hand_store where StoreState ='在库' and boxcode= '" & TextBox21.Text.Trim & "'", Arr2, SQL)
        If StrErr2 <> "" Then ShowOutBoxLabel(StrErr2, Color.Red) : Exit Sub
        If UBound(Arr2, 2) <> 0 Then
            If Arr1(1, 1) <> Arr2(1, 1) Then ShowOutBoxLabel("转出笼框和转入笼框规格不匹配！！", Color.Red) : Exit Sub
        End If

        Dim SQL1() As String
        Dim StrErr3 As String
        ReDim SQL1(1)
        If TextBox20.Text.Length = 5 Then
        	'2018-4-16增加改变笼框的时间changebox_time
            SQL1(1) = "update hand_store set changebox_time=convert(datetime,convert(varchar(20),getdate(),120)),changebox_class='" & NowClass & "',changebox_man='" & NowUser & "', nochange_boxcode=boxcode,boxcode='" & TextBox21.Text.Trim & "' where StoreState = '在库' and boxcode = '" & TextBox20.Text.Trim & "'"
        End If
        If TextBox20.Text.Length = 10 Or TextBox20.Text.Length = 11 Then
        	'2018-4-16增加改变笼框的时间changebox_time
            SQL1(1) = "update hand_store set changebox_time=convert(datetime,convert(varchar(20),getdate(),120)),changebox_class='" & NowClass & "', changebox_man='" & NowUser & "', nochange_boxcode=boxcode, boxcode='" & TextBox21.Text.Trim & "' where StoreState = '在库' and barcode = '" & TextBox20.Text.Trim & "'"
        End If
        StrErr3 = ExeSQLS(SQL1, SQL)
        If StrErr3 <> "" Then ShowOutBoxLabel(StrErr3, Color.Red) : Exit Sub
        ShowOutBoxLabel("转出成功！！", Color.Green)
        'InBoxMessage()
    End Sub

    Private Sub Button38_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button38.Click
        InBoxMessage()
    End Sub

    Private Sub Button42_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button42.Click
        ShowPanel(22)
    End Sub

    Private Sub Button40_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button40.Click
        TextBox5.Visible = False
        ShowPanel(1)
    End Sub

    Private Sub Button35_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button35.Click
        TextBox22.Text = ""
        Label63.Text = ""
        Label63.BackColor = Color.White
        Dim StrSQL As String
        Dim StrErr As String
        StrSQL = "select FBillID from ICBillNo where FBillName LIKE '产品入库%' "
        Dim Arr(,)
        StrErr = GetRst(StrSQL, Arr, SQLK3)
        FBillID = Arr(1, 1)


        Dim StrErr1 As String
        Dim SQLStr1() As String
        ReDim SQLStr1(1)
        SQLStr1(1) = "DELETE FROM ICStockBillEntry  WHERE FQty=0 and FAmount=0"
        StrErr1 = ExeSQLS(SQLStr1, SQLK3)

        GetReasonList()
        ShowPanel(14)
    End Sub

    '获取取消入库原因
    Sub GetReasonList()
        Dim t As Long
        Dim StrSQL As String
        Dim StrErr As String

        StrSQL = "select t.dic_name, t.id  from sys_dic t where t.parent_id =1"
        StrErr = GetRst(StrSQL, ReasonArr, SQL)
        If StrErr <> "" Then
            MsgBox("网络连接失败！！")
            Exit Sub
        End If

        ReasonID = 0
        ComboBox8.Items.Clear()
        ComboBox8.Items.Add("") '默认为空
        For t = 1 To UBound(ReasonArr, 2)
            If ReasonArr(1, t).ToString <> "" Then ComboBox8.Items.Add(ReasonArr(1, t).ToString)
        Next
    End Sub

    Sub ShowCancelLabel(ByVal Str As String, ByVal CC As Color)
        Label63.Text = Str
        Label63.BackColor = CC
    End Sub

    Sub GetNoInBill()  '''''''''生成入库单
        Dim StrSQL1 As String
        Dim StrErr1 As String
        StrSQL1 = "select FInterID from ICStockBill where FBillNo='" & BillNo & "'"
        Dim Arr1(,)
        StrErr1 = GetRst(StrSQL1, Arr1, SQLK3)
        If UBound(Arr1, 2) > 0 Then FInterID = Arr1(1, 1) : Exit Sub

        Dim StrErr3 As String
        StrErr3 = GetFInterID(SQLK3)
        If StrErr3 <> "" Then ShowCancelLabel(StrErr3, Color.Red) : Exit Sub

        Dim Str As String
        Dim SQL1() As String
        ReDim SQL1(1)
        SQL1(1) = "insert into ICStockBill (FBrNo,FTranType,FDate,FBillNo,FExplanation,FFetchAdd,FPOSName,FConfirmMem,FYearPeriod,FInterID,FDCStockID,FFManagerID,FSManagerID,FBillerID,FVchInterID,FUpStockWhenSave,FManageType,FROB,FDeptID) values " & _
                        "('0'," & FBillID & ",convert(datetime,convert(varchar(20),getdate(),120)),'" & BillNo & "','','','','',''," & FInterID & "," & ck & "," & k3User & "," & k3User & ",16427 ,0,0,0,-1,2913)"
        Str = ExeSQLS(SQL1, SQLK3)
        If Str <> "" Then MsgBox(Str) : Exit Sub

    End Sub

    Sub ShowQtInLabel(ByVal Str As String, ByVal CC As Color)
        Label83.Text = Str
        Label83.BackColor = CC
    End Sub
    '20180602新增其它入库
    Sub QtInStore()
        If ComboBox6.Text = "" Then MsgBox("请选择库区！！") : Exit Sub
        If TextBox28.Text = "" Then ShowQtInLabel("请扫描11位条码!!", Color.Red) : Exit Sub
        If TextBox28.Text.Length <> 10 And TextBox28.Text.Length <> 11 Then ShowQtInLabel("请扫描10或者11位条码！！", Color.Red) : Exit Sub
        If TextBox28.Text.Length = 10 Then
            Dim StrErr0 As String
            Dim Arr0(,)
            StrErr0 = GetRst("select * from sync_bcwl where barcode='" & TextBox28.Text.Trim & "'", Arr0, SQLMES)
            If StrErr0 <> "" Then ShowQtInLabel(StrErr0, Color.Red) : Exit Sub
            If UBound(Arr0, 2) <= 0 Then ShowQtInLabel("MES中不存在该条码，不允许入库！！", Color.Red) : Exit Sub
        End If

        Dim StrErr1 As String
        Dim Arr1(,)

        StrErr1 = GetRst("select top 1 Barcode from hand_store where StoreState='在库' and Barcode='" & TextBox28.Text & "'", Arr1, SQL)
        If StrErr1 <> "" Then ShowQtInLabel(StrErr1, Color.Red) : Exit Sub
        If UBound(Arr1, 2) > 0 Then ShowQtInLabel("该条码的轮胎已入库，不能重复入库！！", Color.Red) : Exit Sub

        StrErr1 = GetRst("select top 1 Barcode from hand_store where StoreState='已出库' and Barcode='" & TextBox28.Text & "'", Arr1, SQL)
        If StrErr1 <> "" Then ShowQtInLabel(StrErr1, Color.Red) : Exit Sub
        If UBound(Arr1, 2) > 0 Then ShowQtInLabel("该条码的轮胎已出库，不能重复入库！！", Color.Red) : Exit Sub
        '插入仓储系统
        If TextBox28.Text.Length = 10 Then
            StrErr1 = GetMesFItemID(TextBox28.Text)
        Else
            StrErr1 = GetFItemID(TextBox28.Text)
        End If
        If StrErr1 <> "" Then ShowQtInLabel(StrErr1, Color.Red) : Exit Sub

        If Val(BoxCode) < 99979 Or Val(BoxCode) > 99999 Then
            Dim Arr6(,)
            Dim StrErr6 As String
            StrErr6 = GetRst("select top 1 barcode from hand_store where StoreState='在库' and boxcode='" & BoxCode & "'", Arr6, SQL)
            If StrErr6 <> "" Then ShowQtInLabel(StrErr1, Color.Red) : Exit Sub

            If UBound(Arr6, 2) > 0 Then
                Dim Arr2(,)
                Dim StrErr2 As String
                StrErr2 = GetRst("select top 1 barcode from hand_store where StoreState='在库' and  ProductID = " & FItemID & " and boxcode='" & BoxCode & "'", Arr2, SQL)

                If StrErr2 <> "" Then ShowQtInLabel(StrErr1, Color.Red) : Exit Sub
                If UBound(Arr2, 2) <= 0 Then ShowQtInLabel("同一笼框条码,规格必须一致！！", Color.Red) : Exit Sub

            End If
        End If

        '生成其他入库单
        CreateQtInBill()

        '插入仓储系统
        Dim StrShow As String = ""
        StrErr1 = GetRst("select top 1 Barcode from hand_store where StoreState='取消入库' and Barcode='" & TextBox28.Text & "'", Arr1, SQL)
        If StrErr1 <> "" Then ShowQtInLabel(StrErr1, Color.Red) : Exit Sub
        If UBound(Arr1, 2) > 0 Then
            StrErr1 = updateQtInBar(StrShow)
        Else
            StrErr1 = GetQtInBar(StrShow)
        End If
        If StrErr1 <> "" Then ShowQtInLabel(StrErr1, Color.Red) : Exit Sub

        StrShow = "规格：" & vbCrLf & Product & vbCrLf & vbCrLf & "条码：" & TextBox28.Text & vbCrLf & vbCrLf & StrShow
        ShowQtInLabel(StrShow, Color.Green)
    End Sub
    '20180602新增其它入库
    Sub CreateQtInBill() '生成其他入库单
        Dim StrErr1 As String
        Dim Arr6(,)
        Dim StrErr6 As String
        StrErr6 = GetRst("select FInterID from ICStockBill where FCheckerID is null and FTranType=" & FBillID & " and FBillerID=16427 and FROB=1 and FDate >=convert(datetime,convert(varchar(10),getdate(),120)) and FDate < convert(datetime,convert(varchar(11),dateadd(day,1,getdate()),120))", Arr6, SQLK3)
        If StrErr6 <> "" Then ShowQtInLabel(StrErr6, Color.Red) : Exit Sub
        If UBound(Arr6, 2) <= 0 Then
            StrErr1 = GetBillNo()
            If StrErr1 <> "" Then ShowQtInLabel(StrErr1, Color.Red) : Exit Sub
            GetQtInBill() '生成其他入库主单
        Else
            FInterID = Arr6(1, 1)
        End If

        Dim Arr4(,)
        Dim StrErr4 As String
        StrErr4 = GetRst("select count(*) from ICStockBillEntry where FInterID = " & FInterID, Arr4, SQLK3)
        If StrErr4 <> "" Then ShowQtInLabel(StrErr4, Color.Red) : Exit Sub

        Dim Arr5(,)
        Dim StrErr5 As String
        StrErr5 = GetRst("select FQty from ICStockBillEntry where FInterID = " & FInterID & " and FItemID=" & FItemID, Arr5, SQLK3)
        If StrErr5 <> "" Then ShowQtInLabel(StrErr5, Color.Red) : Exit Sub
        If UBound(Arr5, 2) > 0 Then
            Dim SQL2() As String
            ReDim SQL2(1)
            SQL2(1) = "update ICStockBillEntry set FQty=" & (Arr5(1, 1) + 1) & ",FAuxQty=" & (Arr5(1, 1) + 1) & ",FConsignPrice= FPrice ,FConsignAmount=" & (Arr5(1, 1) + 1) & " * FPrice  where FInterID = " & FInterID & " and FItemID=" & FItemID
            StrErr5 = ExeSQLS(SQL2, SQLK3)
            If StrErr5 <> "" Then ShowQtInLabel(StrErr5, Color.Red) : Exit Sub
        Else
            Dim SQL2() As String
            ReDim SQL2(1)
            SQL2(1) = "insert into ICStockBillEntry (FBrNo,FInterID,FEntryID,FItemID,FQty,FAuxQty,FBatchNo,FSourceBillNo,FContractBillNo,FICMOBillNo,FOrderBillNo,FMTONo,FClientOrderNo,FItemSize,FItemSuite,FPositionNo,FSEOutBillNo,FConfirmMemEntry,FReturnNoticeBillNO,FNote,FUnitID,FDCSPID,FSnListID,FDCStockID,FChkPassItem) values " & _
                      "('0'," & FInterID & "," & (Arr4(1, 1) + 1) & "," & FItemID & ",1,1,'','','','','','','','','','','','','','',254,0,0," & FaultLoc & ",1058)"
            StrErr5 = ExeSQLS(SQL2, SQLK3)
            If StrErr5 <> "" Then ShowQtInLabel(StrErr5, Color.Red) : Exit Sub
        End If

    End Sub
    '20180602新增其它入库
    Sub GetQtInBill()  '''''''''生成其他入库主单
        Dim StrErr3 As String
        StrErr3 = GetFInterID(SQLK3)
        If StrErr3 <> "" Then ShowQtInLabel(StrErr3, Color.Red) : Exit Sub

        Dim Str As String
        Dim SQL1() As String
        ReDim SQL1(1)
        SQL1(1) = "insert into ICStockBill (FBrNo,FTranType,FDate,FBillNo,FExplanation,FFetchAdd,FPOSName,FConfirmMem,FYearPeriod,FInterID,FDCStockID,FFManagerID,FSManagerID,FBillerID,FVchInterID,FUpStockWhenSave,FManageType,FDeptID) values " & _
                "('0'," & FBillID & ",convert(datetime,convert(varchar(20),getdate(),120)),'" & BillNo & "','','','','',''," & FInterID & "," & FaultLoc & "," & k3User & "," & k3User & ",16427 ,0,0,0,2913)"
        Str = ExeSQLS(SQL1, SQLK3)
        If Str <> "" Then ShowQtInLabel(Str, Color.Red) : Exit Sub
    End Sub

    Sub CancelOut()
        If TextBox26.Text = "" Then ShowCancelOutLabel("请扫描11位条码!!", Color.Red) : Exit Sub
        If TextBox26.Text.Length <> 10 And TextBox26.Text.Length <> 11 Then ShowCancelOutLabel("请扫描10或者11位条码！！", Color.Red) : Exit Sub

        Dim Arr1(,)
        Dim StrErr1 As String
        StrErr1 = GetRst("select ProductID,outcode,outno from hand_store where id=(select top 1 id from hand_store where Barcode='" & TextBox26.Text.Trim & "' and StoreState = '已出库' order by InTime desc)", Arr1, SQL)

        If StrErr1 <> "" Then ShowCancelOutLabel(StrErr1, Color.Red) : Exit Sub
        If UBound(Arr1, 2) = 0 Then ShowCancelOutLabel("没有该条码轮胎！！", Color.Red) : Exit Sub


        Dim Arr2(,)
        Dim StrErr2 As String
        StrErr2 = GetRst("select ProductID barcode from hand_store where StoreState='在库' and boxcode='" & BoxCode & "'", Arr2, SQL)
        If StrErr2 <> "" Then ShowCancelOutLabel(StrErr2, Color.Red) : Exit Sub
        If UBound(Arr2, 2) > 0 Then
            If Arr2(1, 1) <> Arr1(1, 1) Then ShowCancelOutLabel("同一笼框条码,规格必须一致！！", Color.Red) : Exit Sub
        End If
        '2016-08-06修改 反入库时判断是否审核，如果审核生成红字入库单，，如果没有什么直接修改出库单 --start
        Dim Arr4(,)
        Dim StrErr4 As String
        StrErr4 = GetRst("select FCheckerID from ICStockBill where FInterID =" & Arr1(2, 1), Arr4, SQLK3)
        If StrErr4 <> "" Then ShowCancelOutLabel(StrErr4, Color.Red) : Exit Sub
        If UBound(Arr4, 2) > 0 And Arr4(1, 1) <> "" Then '已审核
            '2017-01-13 已审核的出库单，不允许反入库
            ' OldFInterID = Arr1(2, 1)
            'CreateNoOutBill(Arr1(1, 1))
            ShowCancelOutLabel("已审核的出库单，不允许反入库!!", Color.Red)
            Exit Sub
        Else
            ModifyNoOutBill(TextBox26.Text.Trim, Arr1(2, 1), Arr1(1, 1)) '未审核
        End If
        '2016-08-06  修改end

        'ModifyNoOutBill(Arr1(2, 1), Arr1(1, 1))
        Dim SQL1() As String
        Dim StrErr3 As String
        ReDim SQL1(1)
        '2018-4-16增加反入库的时间frk_time
        SQL1(1) = "update hand_store set StoreState='在库',frk_flag=1,nofrk_boxcode=boxcode,boxcode='" & BoxCode & "',frk_class='" & NowClass & "',frk_man='" & NowUser & "',frk_time=convert(datetime,convert(varchar(20),getdate(),120)),flag=1 where id=(select top 1 id from hand_store where Barcode='" & TextBox26.Text.Trim & "' and StoreState = '已出库' order by InTime desc)"
        StrErr3 = ExeSQLS(SQL1, SQL)
        If StrErr3 <> "" Then ShowCancelOutLabel(StrErr3, Color.Red) : Exit Sub

        If BillNo Is Nothing Then '已审核
            SQL1(1) = "update KTMSSQL.dbo.sync_bcwl set operatetype = 'update',crkbz=3,tongbu=0,rk_time=convert(varchar(10),getdate(),120)  where barcode='" & TextBox26.Text & "'"
            '未审核
        Else

            SQL1(1) = "update KTMSSQL.dbo.sync_bcwl set operatetype = 'update',crkbz=3,tongbu=0,rk_time=convert(varchar(10),getdate(),120),ck_no='" & BillNo & "' where barcode='" & TextBox26.Text & "'"

        End If
        ' SQL1(1) = "update KTMSSQL.dbo.sync_bcwl set operatetype = 'update',crkbz=3,tongbu=0 where barcode='" & TextBox26.Text & "'"
        StrErr1 = ExeSQLS(SQL1, SQLMES)
        If StrErr1 <> "" Then MsgBox(StrErr1) : Exit Sub
        ShowCancelOutLabel("取消出库成功!!", Color.Green)
    End Sub


    Sub ModifyNoOutBill(ByVal barcode As String, ByVal FInID As Long, ByVal FItID As Long) '修改出库单
        Dim StrErr1 As String
        Dim SQL1() As String
        ReDim SQL1(1)
        SQL1(1) = "update ICStockBillEntry set FQty=FQty-1,FAuxQty=FAuxQty-1,FAmount=0,FConsignAmount=(FQty-1)*FPrice/(FQty*FPrice/FConsignAmount) where FInterID = " & FInID & " and FItemID=" & FItID
        StrErr1 = ExeSQLS(SQL1, SQLK3)
        If StrErr1 <> "" Then ShowCancelOutLabel(StrErr1, Color.Red) : Exit Sub
        'AddOperLog(barcode, "反入库", FInID, BoxCode)
    End Sub

    Sub ShowCancelOutLabel(ByVal Str As String, ByVal CC As Color)
        Label70.Text = Str
        Label70.BackColor = CC
        If CC = Color.Red Then
            MsgBox(Str)
        End If
    End Sub

    Sub CancelBarcode()
        If ComboBox8.Text = "" Then MsgBox("请选择取消原因！！") : Exit Sub
        If TextBox22.Text = "" Then ShowCancelLabel("请扫描5位或11位条码!!", Color.Red) : Exit Sub
        If TextBox22.Text.Length <> 10 And TextBox22.Text.Length <> 11 And TextBox22.Text.Length <> 5 Then ShowCancelLabel("请扫描5位、10位或11位条码！！", Color.Red) : Exit Sub

        Dim Arr1(,)
        Dim StrErr1 As String
        If TextBox22.Text.Length = 10 Or TextBox22.Text.Length = 11 Then
            StrErr1 = GetRst("select ProductID,incode,inno from hand_store where id=(select top 1 id from hand_store where Barcode='" & TextBox22.Text.Trim & "' and StoreState = '在库' order by InTime desc)", Arr1, SQL)
        Else
            StrErr1 = GetRst("select ProductID,incode,inno from hand_store where StoreState='在库' and Boxcode='" & TextBox22.Text.Trim & "'", Arr1, SQL)
        End If

        If StrErr1 <> "" Then ShowCancelLabel(StrErr1, Color.Red) : Exit Sub
        If UBound(Arr1, 2) = 0 Then ShowCancelLabel("没有该条码轮胎！！", Color.Red) : Exit Sub

        Dim Arr4(,)
        Dim StrErr4 As String
        Dim t As Long
        For t = 1 To UBound(Arr1, 2)
            If Arr1(2, t) <> "" Then

                StrErr4 = GetRst("select FCheckerID from ICStockBill where FInterID =" & Arr1(2, 1), Arr4, SQLK3)
                If StrErr4 <> "" Then ShowCancelLabel(StrErr4, Color.Red) : Exit Sub
                If UBound(Arr4, 2) > 0 And Arr4(1, 1) <> "" Then '已审核
                    CreateNoInBill(TextBox22.Text, Arr1(1, 1))
                Else
                    ModifyNoInBill(TextBox22.Text, Arr1(2, 1), Arr1(1, 1)) '未审核
                End If

            Else
                deleteInBarcode(TextBox22.Text.Trim)
                'AddOperLog(TextBox22.Text.Trim, "入库取消(未上传)")
            End If
        Next

        If ReasonID <= 0 Then ReasonID = 3

        Dim SQL1() As String
        Dim StrErr3 As String
        ReDim SQL1(1)
        If TextBox22.Text.Length = 10 Or TextBox22.Text.Length = 11 Then
            '2018-4-16增加取消入库的时间frk_time,flag=1为了不上传K3
            'SQL1(1) = "delete from hand_store where id=(select top 1 id from hand_store where Barcode='" & TextBox22.Text.Trim & "' and StoreState = '在库' order by InTime desc)"
            SQL1(1) = "update hand_store set incode = null, inno = null, flag=1,StoreState = '取消入库',qxrk_time=convert(datetime,convert(varchar(20),getdate(),120)),qxrk_class='" & NowClass & "',qxrk_man='" & NowUser & "',qxrk_reason=" & ReasonID & " where id=(select top 1 id from hand_store where Barcode='" & TextBox22.Text.Trim & "' and StoreState = '在库' order by InTime desc)"
            StrErr3 = ExeSQLS(SQL1, SQL)
            If StrErr3 <> "" Then ShowCancelLabel(StrErr3, Color.Red) : Exit Sub
            ShowCancelLabel("取消入库成功!!", Color.Green)
            Exit Sub
        End If
        '2018-4-16增加取消入库的时间frk_time,flag=1为了不上传K3
        'SQL1(1) = "delete from hand_store where StoreState = '在库' and boxcode = '" & TextBox22.Text.Trim & "'"
        SQL1(1) = "update hand_store set incode = null, inno = null, flag=1,StoreState = '取消入库',qxrk_time=convert(datetime,convert(varchar(20),getdate(),120)),qxrk_class='" & NowClass & "',qxrk_man='" & NowUser & "',qxrk_reason=" & ReasonID & " where StoreState = '在库' and boxcode = '" & TextBox22.Text.Trim & "'"
        StrErr3 = ExeSQLS(SQL1, SQL)
        If StrErr3 <> "" Then ShowCancelLabel(StrErr3, Color.Red) : Exit Sub
        ShowCancelLabel("取消入库成功!!", Color.Green)
    End Sub
    
    
    Sub deleteInBarcode(ByVal Str As String) '没有生成入库单，直接删除

        Dim StrErr1 As String
        Dim SQL1() As String
        ReDim SQL1(1)
        '2018-4-16增加取消入库的时间frk_time,flag=1为了不上传K3
        SQL1(1) = "update hand_store set flag=1,StoreState = '取消入库',qxrk_time=convert(datetime,convert(varchar(20),getdate(),120)),qxrk_class='" & NowClass & "',qxrk_man='" & NowUser & "' WHERE Barcode = '" & Str & "'"
        StrErr1 = ExeSQLS(SQL1, SQL)
        If StrErr1 <> "" Then ShowCancelLabel(StrErr1, Color.Red) : Exit Sub
    End Sub

    Sub CreateNoInBill(ByVal barcode As String, ByVal FItemID As Long) '生成取消入库单
        Dim StrSQL As String
        Dim StrErr As String
        StrSQL = "select FBillID from ICBillNo where FBillName LIKE '产品入库%' "
        Dim Arr(,)
        StrErr = GetRst(StrSQL, Arr, SQLK3)
        If StrErr <> "" Then ShowCancelLabel(StrErr, Color.Red) : Exit Sub
        FBillID = Arr(1, 1)


        Dim Arr7(,)
        Dim StrErr7 As String

        StrErr7 = GetRst("select FItemID,Fname,fdefaultloc from t_icitem where FItemID=" & FItemID, Arr7, SQLK3)
        If StrErr7 <> "" Then ShowCancelLabel(StrErr7, Color.Red) : Exit Sub
        ck = Arr7(3, 1)

        Dim Arr6(,)
        Dim StrErr6 As String
        StrErr6 = GetRst("select FInterID from ICStockBill where FCheckerID is null and FTranType=2 and FBillerID=" & k3User & " and FROB=-1 and FDate >=convert(datetime,convert(varchar(10),getdate(),120)) and FDate < convert(datetime,convert(varchar(11),dateadd(day,1,getdate()),120))", Arr6, SQLK3)
        If StrErr6 <> "" Then ShowCancelLabel(StrErr6, Color.Red) : Exit Sub
        If UBound(Arr6, 2) <= 0 Then

            Dim StrErr1 As String
            StrErr1 = GetBillNo()
            If StrErr1 <> "" Then ShowCancelLabel(StrErr1, Color.Red) : Exit Sub
            GetNoInBill() '生成取消入库单
        Else
            FInterID = Arr6(1, 1)
        End If


        Dim Arr4(,)
        Dim StrErr4 As String
        StrErr4 = GetRst("select count(*) from ICStockBillEntry where FInterID = " & FInterID, Arr4, SQLK3)
        If StrErr4 <> "" Then ShowCancelLabel(StrErr4, Color.Red) : Exit Sub

        Dim Arr5(,)
        Dim StrErr5 As String
        StrErr5 = GetRst("select FQty from ICStockBillEntry where FInterID = " & FInterID & " and FItemID=" & FItemID, Arr5, SQLK3)
        If StrErr5 <> "" Then ShowCancelLabel(StrErr5, Color.Red) : Exit Sub
        If UBound(Arr5, 2) > 0 Then
            Dim SQL2() As String
            ReDim SQL2(1)
            SQL2(1) = "update ICStockBillEntry set FQty=" & (Arr5(1, 1) - 1) & ",FAuxQty=" & (Arr5(1, 1) - 1) & " where FInterID = " & FInterID & " and FItemID=" & FItemID
            StrErr5 = ExeSQLS(SQL2, SQLK3)
            If StrErr5 <> "" Then ShowCancelLabel(StrErr5, Color.Red) : Exit Sub
        Else
            Dim SQL2() As String
            ReDim SQL2(1)
            SQL2(1) = "insert into ICStockBillEntry (FBrNo,FInterID,FEntryID,FItemID,FQty,FAuxQty,FBatchNo,FContractBillNo,FICMOBillNo,FMTONo,FClientOrderNo,FItemSize,FItemSuite,FPositionNo,FSEOutBillNo,FConfirmMemEntry,FReturnNoticeBillNO,FNote,FUnitID,FDCSPID,FSnListID,FChkPassItem,FDCStockID) values ('0'," & FInterID & "," & (Arr4(1, 1) + 1) & "," & FItemID & "," & -1 & "," & -1 & ",'','','','','','','','','','','','',254,0,0,1058," & ck & ")"
            StrErr5 = ExeSQLS(SQL2, SQLK3)
            If StrErr5 <> "" Then ShowCancelLabel(StrErr5, Color.Red) : Exit Sub
        End If

        'AddOperLog(barcode, "入库取消", FInterID)

    End Sub

    Sub ModifyNoInBill(ByVal barcode As String, ByVal FInID As Long, ByVal FItID As Long) '修改入库单
        Dim StrErr1 As String
        Dim SQL1() As String
        ReDim SQL1(1)
        SQL1(1) = "update ICStockBillEntry set FQty=FQty-1,FAuxQty=FAuxQty-1 where FInterID = " & FInID & " and FItemID=" & FItID
        StrErr1 = ExeSQLS(SQL1, SQLK3)
        If StrErr1 <> "" Then ShowCancelLabel(StrErr1, Color.Red) : Exit Sub
        'AddOperLog(barcode, "入库取消", FInID)
    End Sub

    Sub BarcodeMessage()
        Dim str1 As String
        str1 = ""
        ''''''格式判断
        If TextBox23.Text = "" Then ShowBarcodeMessage("请扫描11位条码！！", Color.Red) : Exit Sub
        If TextBox23.Text.Length <> 10 And TextBox23.Text.Length <> 11 Then ShowBarcodeMessage("请扫描11位条码！！", Color.Red) : Exit Sub

        Dim StrErr1 As String
        Dim Arr1(,)
        '2018-06-10 修改轮胎信息只显示【在库】和【已出库】
        'StrErr1 = GetRst("select ProductID from hand_store where Barcode='" & TextBox23.Text.Trim & "'", Arr1, SQL)
        StrErr1 = GetRst("select ProductID from hand_store where StoreState='在库' and Barcode='" & TextBox23.Text.Trim & _
                "' UNION select ProductID from hand_store where StoreState='已出库' and Barcode='" & TextBox23.Text.Trim & "'", Arr1, SQL)

        If StrErr1 <> "" Then ShowBarcodeMessage(StrErr1, Color.Red) : Exit Sub
        If UBound(Arr1, 2) <= 0 Then ShowBarcodeMessage("没有该轮胎信息！！", Color.Red) : Exit Sub


        Dim StrErr2 As String
        Dim Arr2(,)
        StrErr2 = GetRst("select Fname from t_icitem where FItemID = " & Arr1(1, 1), Arr2, SQLK3)
        If StrErr2 <> "" Then ShowBarcodeMessage(StrErr2, Color.Red) : Exit Sub
        If UBound(Arr2, 2) = 0 Then ShowBarcodeMessage("不识别K3规格,请确认K3规格！！", Color.Red) : Exit Sub
        ShowBarcodeMessage("轮胎规格：" & vbLf & Arr2(1, 1), Color.Green)

    End Sub

    Sub ShowBarcodeMessage(ByVal Str As String, ByVal CC As Color)
        Label67.Text = Str
        Label67.BackColor = CC
    End Sub

    Private Sub Button45_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button45.Click
        BarcodeMessage()
    End Sub

    Private Sub Button46_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button46.Click
        ShowPanel(10)
    End Sub

    Private Sub Button44_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        ShowPanel(16)
    End Sub

    Private Sub Button50_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button50.Click
        ShowPanel(10)
    End Sub

    Private Sub Button48_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button48.Click
        If TextBox24.Text = "" Then ShowBoxLabel("请扫描5位条码!", Color.Red) : Exit Sub
        If TextBox24.Text.Length <> 5 Then ShowBoxLabel("请扫描5位条码！", Color.Red) : Exit Sub
        GetBoxData() '获取笼框规格信息
        ShowPanel(17)
    End Sub

    Private Sub Button53_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button53.Click
        Dim StrErr As String
        Dim Arr(,)
        StrErr = GetRst("select barcode from hand_store where StoreState='在库' and boxcode='" & TextBox24.Text & "'", Arr, SQL)
        If StrErr <> "" Then
            SetLog(StrErr)
            MsgBox("连接数据库失败！！")
            Exit Sub
        End If
        BoxNum = UBound(Arr, 2)
        ShowOldBoxLabel("当前容量： " & BoxNum, Color.Green)
        ShowPanel(16)
    End Sub

    Private Sub Button51_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button51.Click
        ShowPanel(10)
    End Sub

    Sub OldBoxMessage()
        If TextBox24.Text.Length <> 5 Then ShowOldBoxLabel("请扫描5位条码！", Color.Red) : Exit Sub
        Dim StrErr As String
        Dim Arr(,)
        StrErr = GetRst("select barcode from hand_store where StoreState='在库' and boxcode='" & TextBox24.Text & "'", Arr, SQL)
        If StrErr <> "" Then
            SetLog(StrErr)
            MsgBox("连接数据库失败！！")
            Exit Sub
        End If
        BoxNum = UBound(Arr, 2)
        ShowOldBoxLabel("当前容量： " & BoxNum, Color.Green)
    End Sub

    Sub ShowOldBoxLabel(ByVal Str As String, ByVal CC As Color)

        Label75.Text = Str
        Label75.BackColor = CC
        'If CC = Color.Red Then
        '    MsgBox(Str)
        'End If
    End Sub

    Sub OldInStore()
        If TextBox25.Text = "" Then ShowOldInLabel("请扫描条码", Color.Red) : Exit Sub
        ''''''格式判断
        If TextBox25.Text.Length <> 10 And TextBox25.Text.Length <> 11 Then ShowOldInLabel("请扫描11位条码！", Color.Red) : Exit Sub

        Dim StrErr2 As String
        Dim Arr2(,)
        StrErr2 = GetRst("select Barcode from hand_store where Barcode='" & TextBox25.Text & "' and InTime>=DATEADD(n,-30,convert(datetime,convert(varchar(20),getdate(),120))) and InTime<=convert(datetime,convert(varchar(20),getdate(),120)) and StoreState='在库'", Arr2, SQL)
        If StrErr2 <> "" Then ShowOldInLabel(StrErr2, Color.Red) : Exit Sub
        If UBound(Arr2, 2) > 0 Then ShowOldInLabel("半个小时内不能入库重复条码的轮胎！！", Color.Red) : Exit Sub

        Dim StrErr1 As String

        StrErr1 = GetOldFItemID(TextBox25.Text)
        If StrErr1 <> "" Then ShowOldInLabel(StrErr1, Color.Red) : Exit Sub

        Dim StrErr As String
        Dim StrShow As String = ""
        StrErr = GetOldInBar(StrShow)

        If StrErr <> "" Then ShowOldInLabel(StrErr, Color.Red) : Exit Sub
        GetBoxData() '获取笼框规格信息
        ShowOldInLabel(StrShow, Color.Green)

    End Sub


    Sub GetBoxData()  '''''''''获取笼框信息

        Dim StrErr As String
        Dim Arr(,)
        StrErr = Me.GetRst("select ProductID,count(ProductID) from hand_store where boxcode='" & TextBox24.Text.Trim & "' and StoreState='在库' group by ProductID", Arr, SQL)
        If StrErr <> "" Then MsgBox(StrErr) : Exit Sub

        Dim ds As New DataSet
        Dim dt As New DataTable
        dt.Columns.Add("规格型号", Type.GetType("System.String"))
        dt.Columns.Add("数量", Type.GetType("System.Int32"))
        For t = 1 To UBound(Arr, 2)

            Dim StrErr1 As String
            Dim Arr1(,)
            StrErr1 = Me.GetRst("select Fname from t_icitem  where  FItemID=" & Arr(1, t), Arr1, SQLK3)
            If StrErr1 <> "" Then MsgBox(StrErr1) : Exit Sub

            Dim dw = dt.NewRow
            dw.Item(0) = Arr1(1, 1)
            dw.Item(1) = Arr(2, t) - 0
            dt.Rows.Add(dw)

        Next
        ds.Tables.Add(dt)
        DataGrid3.DataSource = ds.Tables(0)

        'DataGrid1.Font = New Font("宋体", 8, FontStyle.Regular )

        '''''''''''''修改列宽

        DataGrid3.TableStyles.Clear()
        DataGrid3.TableStyles.Add(New DataGridTableStyle)
        DataGrid3.TableStyles.Item(0).MappingName = dt.TableName
        DataGrid3.TableStyles(0).GridColumnStyles.Item(0).Width = 175
        DataGrid3.TableStyles(0).GridColumnStyles.Item(1).Width = 45

    End Sub

    Function GetOldInBar(ByRef StrShow As String)
        SetLog("准备扫码操作")
        Dim LogStr As String = ""
        GetOldInBar = ""

        Dim Arr6(,)
        Dim StrErr6 As String
        StrErr6 = GetRst("select barcode from hand_store where StoreState='在库' and boxcode='" & TextBox24.Text.Trim & "'", Arr6, SQL)
        If StrErr6 <> "" Then GetOldInBar = StrErr6 : Exit Function

        If Val(TextBox24.Text.Trim) < 99979 Or Val(TextBox24.Text.Trim) > 99999 Then
            If UBound(Arr6, 2) > 0 Then

                Dim Arr1(,)
                Dim StrErr1 As String
                StrErr1 = GetRst("select barcode from hand_store where StoreState='在库' and  barcode ='" & TextBox25.Text.Trim & "' and boxcode='" & BoxCode & "'", Arr1, SQL)
                If StrErr1 <> "" Then GetOldInBar = StrErr1 : Exit Function
                If UBound(Arr1, 2) > 0 Then GetOldInBar = "同一笼框条码" & TextBox25.Text.Trim & "，条码不能重复！！" : Exit Function

                Dim Arr2(,)
                Dim StrErr2 As String
                StrErr2 = GetRst("select barcode from hand_store where StoreState='在库' and boxcode='" & TextBox24.Text.Trim & "'", Arr2, SQL)
                If StrErr2 <> "" Then GetOldInBar = StrErr2 : Exit Function
                If UBound(Arr2, 2) > 0 Then
                    If Mid(Arr2(1, 1), 3, 5) <> Mid(TextBox25.Text.Trim, 3, 5) Then
                        GetOldInBar = "同一笼框条码" & TextBox25.Text.Trim & "，规格必须一致！！"
                        Exit Function
                    End If
                End If

            End If
        End If

        Dim StrErr5 As String
        Dim SQL1() As String
        ReDim SQL1(1)
        '2018-4-16老库存不上传k3，标识改为1
        SQL1(1) = "insert into hand_store (StoreState,InClass,ProductID,inman,intime,indate,barcode,boxcode,flag) values " & _
                        "('在库','" & NowClass & "'," & FItemID & ",'" & NowUser & "',convert(datetime,convert(varchar(20),getdate(),120)),convert(varchar(10),getdate(),120),'" & TextBox25.Text.Trim & "','" & TextBox24.Text.Trim & "',1)"
        LogStr = LogStr & vbCrLf & "扫码扫描:[" & TextBox25.Text & "]"
        StrErr5 = ExeSQLS(SQL1, SQL)
        If StrErr5 <> "" Then MsgBox(StrErr5) : Exit Function
        SetLog("扫码完成" & LogStr)
        StrShow = "扫码入库成功！！"
    End Function

    Sub ShowOldInLabel(ByVal Str As String, ByVal CC As Color)

        Label88.Text = Str
        Label88.BackColor = CC
        'If CC = Color.Red Then
        '    MsgBox(Str)
        'End If
    End Sub

    Private Sub Button49_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button49.Click
        OldBoxMessage()
    End Sub

    Private Sub Button52_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button52.Click
        OldInStore()
    End Sub

    Private Sub Button41_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button41.Click
        CancelBarcode()
    End Sub
    Function GetOldFItemID(ByVal Barcode As String)
        Dim Arr1(,)
        Dim StrErr1 As String
        GetOldFItemID = ""
        StrErr1 = GetRst("select a.name,b.name,a.speed,a.tread from standard a,brand b where a.code='" & Mid(Barcode, 3, 3) & "' and b.code='" & Mid(Barcode, 6, 2) & "'", Arr1, SQL)

        If StrErr1 <> "" Then ShowOldInLabel(StrErr1, Color.Red) : Exit Function
        If UBound(Arr1, 2) < 1 Then GetOldFItemID = "没有该规格,请先添加！！" : Exit Function

        Dim Arr2(,)
        Dim StrErr2 As String
        StrErr2 = GetRst("select FItemID,Fname,fdefaultloc from t_icitem where Fname like '" & Arr1(1, 1) & "%" & Arr1(3, 1) & "%" & Arr1(4, 1) & "%" & Arr1(2, 1) & "'", Arr2, SQLK3)
        If StrErr2 <> "" Then GetOldFItemID = StrErr2 : Exit Function
        If UBound(Arr2, 2) = 0 Then GetOldFItemID = "不识别K3规格,请确认K3规格！！" : Exit Function
        If Arr2(1, 1) = 0 Then GetOldFItemID = "不识别K3规格,请确认K3规格！！" : Exit Function
        FItemID = Arr2(1, 1)
        Product = Arr2(2, 1)
        FaultLoc = Arr2(3, 1)
    End Function

    Private Sub Button47_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim StrErr As String
        Dim SQLStr() As String
        ReDim SQLStr(1)
        SQLStr(1) = "DELETE FROM ICStockBillEntry  WHERE FQty=0 and FAmount=0"
        StrErr = ExeSQLS(SQLStr, SQLK3)
        ShowPanel(21)
    End Sub

    Private Sub Button55_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button55.Click
        CancelOut()
    End Sub

    Private Sub Button56_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button56.Click
        ShowPanel(21)
    End Sub

    Private Sub Button54_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button54.Click
        ShowPanel(1)
    End Sub

    Private Sub Button57_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button57.Click
        ShowCount()
        Label30.Visible = True
        Label24.Visible = True
        Label29.Visible = True
        Label25.Visible = True
    End Sub

    Private Sub Button63_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button63.Click
        ShowPanel(8)
    End Sub

    Private Sub Button60_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        ShowICStockK3Label("正在生成.....", Color.Green)
        'Dim Arr1(,)
        'Dim StrErr1 As String
        Dim Arr2(,)
        Dim StrErr2 As String
        'StrErr1 = GetRst("SELECT count(*) from hand_store where InMan='" & NowUser & "' and InDate >=convert(varchar(10),DATEADD(day,-2,convert(varchar(10),getdate(),120)),120) and InDate<=convert(varchar(10),getdate(),120) and flag=0", Arr1, SQL)
        'If StrErr1 <> "" Then ShowICStockK3Label(StrErr1, Color.Red) : Exit Sub
        Dim StrErr As String
        Dim SQLStr() As String
        ReDim SQLStr(2)
        SQLStr(1) = "insert into ICStocks (ProductID,FaultLoc,FQty,flag,InMan,InDate) select ProductID,FaultLoc,count(*),flag,InMan,InDate from hand_store where " & _
                    "InMan='" & NowUser & "' and  InDate >=convert(varchar(10),DATEADD(day,-2,convert(varchar(10),getdate(),120)),120) and InDate<=convert(varchar(10),getdate(),120)  and flag=0 group by ProductID,FaultLoc,flag,InMan,InDate"

        StrErr2 = GetRst("SELECT count(*) from hand_store where InMan='" & NowUser & "' and InDate >=convert(varchar(10),DATEADD(day,-2,convert(varchar(10),getdate(),120)),120) and InDate<=convert(varchar(10),getdate(),120) and flag=0", Arr2, SQL)
        If StrErr2 <> "" Then ShowICStockK3Label(StrErr2, Color.Red) : Exit Sub
        	'2016-4-16增加上传k3的时间
        SQLStr(2) = "update hand_store set flag=1,k3_time=convert(datetime,convert(varchar(20),getdate(),120)) where InMan='" & NowUser & "' and  InDate >=convert(varchar(10),DATEADD(day,-2,convert(varchar(10),getdate(),120)),120) and InDate<=convert(varchar(10),getdate(),120) and flag=0"
        StrErr = ExeSQLS(SQLStr, SQL)
        If StrErr <> "" Then ShowICStockK3Label(StrErr, Color.Red) : Exit Sub
        If Arr2(1, 1) = 0 Then
            ShowICStockK3Label("没有要生成单据的轮胎，请先入库之后进行此操作", Color.Red)
            Exit Sub
        End If
        ShowICStockK3Label("生成成功，共" & Arr2(1, 1) & "条", Color.Green)
    End Sub

    'Sub ShowICStockLabel(ByVal Str As String, ByVal CC As Color)

    '    Label80.Text = Str
    '    Label80.BackColor = CC
    '    'If CC = Color.Red Then
    '    '    MsgBox(Str)
    '    'End If
    'End Sub

    Sub ShowICStockK3Label(ByVal Str As String, ByVal CC As Color)

        Label78.Text = Str
        Label78.BackColor = CC
        'If CC = Color.Red Then
        '    MsgBox(Str)
        'End If
    End Sub

    Private Sub Button62_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button62.Click
        ShowPanel(20)
    End Sub

    Function ToK3Storage(ByRef Arr1(,), ByRef FBillID, ByRef flag) As String
        ToK3Storage = ""

        Dim StrErr2 As String
        StrErr2 = GetBillNo() '生成入库单据号
        If StrErr2 <> "" Then ShowICStockK3Label(StrErr2, Color.Red) : Exit Function

        Dim StrErr3 As String
        StrErr3 = GetFInterID(SQLK3) '生成入库单内码
        If StrErr3 <> "" Then ShowICStockK3Label(StrErr3, Color.Red) : Exit Function

        Dim Str As String
        Dim SQL1() As String
        ReDim SQL1(1) '生成入库单主表
        SQL1(1) = "insert into ICStockBill (FBrNo,FTranType,FDate,FBillNo,FExplanation,FFetchAdd,FPOSName,FConfirmMem,FYearPeriod,FInterID,FDCStockID,FFManagerID,FSManagerID,FBillerID,FVchInterID,FUpStockWhenSave,FManageType,FDeptID) values " & _
                        "('0'," & FBillID & ",DATEADD(day," & flag & ",convert(varchar(20),getdate(),120)),'" & BillNo & "','','','','',''," & FInterID & "," & Arr1(2, 1) & "," & k3User & "," & k3User & ",16427 ,0,0,0,2913)"
        Str = ExeSQLS(SQL1, SQLK3)
        If Str <> "" Then ShowICStockK3Label(Str, Color.Red) : Exit Function

        Dim SQLStr() As String
        ReDim SQLStr(1)
        Dim StrErr4 As String
        SQLStr(1) = "UPDATE KTMSSQL.dbo.sync_bcwl SET rk_no ='" & BillNo & "',operatetype = 'update', rjh =(SELECT boxcode from wms.dbo.hand_store where wms.dbo.hand_store.Barcode=KTMSSQL.dbo.sync_bcwl.barcode and wms.dbo.hand_store.InMan = '" & NowUser & "' AND wms.dbo.hand_store.InDate = CONVERT (VARCHAR (10),DATEADD(DAY," & flag & ",CONVERT (VARCHAR(10), getdate(), 120)),120) AND wms.dbo.hand_store.flag = 0),crkbz = 1,rk_time='" & Arr1(4, 1) & "' WHERE barcode IN (SELECT barcode from wms.dbo.hand_store where InMan = '" & NowUser & "' AND InDate = CONVERT (VARCHAR (10),DATEADD(DAY," & flag & ",CONVERT (VARCHAR(10), getdate(), 120)),120)AND flag = 0)"
        StrErr4 = ExeSQLS(SQLStr, SQLMES)
        If StrErr4 <> "" Then ShowICStockK3Label(StrErr4, Color.Red) : Exit Function
        '2016-4-16增加上传k3的时间 
        SQLStr(1) = "update hand_store set incode=" & FInterID & ",flag=1,k3_time=convert(datetime,convert(varchar(20),getdate(),120)) where InMan='" & NowUser & "' and  InDate =convert(varchar(10),DATEADD(day," & flag & ",convert(varchar(10),getdate(),120)),120) and flag=0"
        StrErr4 = ExeSQLS(SQLStr, SQL)
        If StrErr4 <> "" Then ShowICStockK3Label(StrErr4, Color.Red) : Exit Function

        'AddOperLog(flag)

        Dim SQL2() As String
        ReDim SQL2(1)
        Dim t As Long
        Dim count As Long = 0
        For t = 1 To UBound(Arr1, 2)

            SQL2(1) = "insert into ICStockBillEntry (FBrNo,FInterID,FEntryID,FItemID,FQty,FAuxQty,FBatchNo,FSourceBillNo,FContractBillNo,FICMOBillNo,FOrderBillNo,FMTONo,FClientOrderNo,FItemSize,FItemSuite,FPositionNo,FSEOutBillNo,FConfirmMemEntry,FReturnNoticeBillNO,FNote,FUnitID,FDCSPID,FSnListID,FDCStockID,FChkPassItem) values " & _
                       "('0'," & FInterID & "," & t & "," & Arr1(1, t) & "," & Arr1(3, t) & "," & Arr1(3, t) & ",'','','','','','','','','','','','','','',254,0,0," & Arr1(2, t) & ",1058)"
            StrErr3 = ExeSQLS(SQL2, SQLK3)
            If StrErr3 <> "" Then ShowICStockK3Label(StrErr3, Color.Red) : Exit Function

            'SQL1(1) = "update ICStocks set flag=1,BillNo='" & BillNo & "' where id=" & Arr1(4, t)
            'Str = ExeSQLS(SQL1, SQL)
            'If Str <> "" Then ShowICStockK3Label(Str, Color.Red) : Exit Function
            count = count + Arr1(3, t)
        Next

        ToK3Storage = BillNo & "共" & count & "条" & vbCrLf
    End Function

    Private Sub Button58_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button58.Click
        Dim StrSQL As String
        Dim StrErr As String
        StrSQL = "select FBillID from ICBillNo where FBillName LIKE '产品入库%' "
        Dim Arr(,)
        StrErr = GetRst(StrSQL, Arr, SQLK3)
        If StrErr <> "" Then ShowICStockK3Label(StrErr, Color.Red) : Exit Sub
        FBillID = Arr(1, 1)

        ShowICStockK3Label("正在生成.....", Color.Green)

        Dim Arr1(,)
        Dim StrErr1 As String
        Dim message As String = ""
        Dim message1 As String = ""

        Dim t As Long
        Dim count As Long = 0
        For t = -2 To 0 '天数
            'StrErr1 = GetRst("select ProductID,FaultLoc,FQty,id from ICStocks where InMan='" & NowUser & "' and InDate =convert(varchar(10),DATEADD(day," & t & ",convert(varchar(10),getdate(),120)),120) and flag=0", Arr1, SQL)
            '人库修改20160725 ---start
            'date1 = Now
            StrErr1 = GetRst("select ProductID,FaultLoc,count(*),InDate from hand_store where InMan='" & NowUser & "' and InDate =convert(varchar(10),DATEADD(day," & t & ",convert(varchar(10),getdate(),120)),120) and flag=0 group by ProductID,FaultLoc,flag,InMan,InDate order by ProductID", Arr1, SQL)
            If StrErr1 <> "" Then ShowICStockK3Label(NowUser & "DAY," & t & ",CONVERT" & StrErr1, Color.Red) : Exit Sub

          
            '人库修改20160725 ---end
            If UBound(Arr1, 2) > 0 Then
                message1 = ToK3Storage(Arr1, FBillID, t)
                message = message + message1
            Else
                count = count + 1
            End If
        Next


        If count = 3 Then
            ShowICStockK3Label("没有要上传的单据，请先生成手持入库再进行此操作！！", Color.Red)
            Exit Sub
        End If
        ShowICStockK3Label("上传成功！！" & vbCrLf & message, Color.Green)
    End Sub

    Private Sub Button64_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button64.Click
        ShowPanel(19)
    End Sub

    Private Sub Button59_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button59.Click
        ShowPanel(8)
    End Sub

    Private Sub Button24_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button24.Click
        ShowPanel(20)
    End Sub

    Private Sub Button61_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button61.Click
        ShowPanel(22)
    End Sub

    Private Sub Button43_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button43.Click
        ShowPanel(6)
    End Sub

    Private Sub Button60_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button60.Click
        If TextBox27.Text = "" Then ShowBoxLabel("请扫描5位条码!", Color.Red) : Exit Sub
        If TextBox27.Text.Length <> 5 Then ShowBoxLabel("请扫描5位条码！", Color.Red) : Exit Sub

        BoxCode = TextBox27.Text.Trim
        ShowPanel(18)
    End Sub

    Private Sub Button66_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button66.Click
        ShowPanel(11)
    End Sub

    Private Sub Button65_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button65.Click
        If TextBox27.Text = "" Then MsgBox("请输入条码！！") : Exit Sub
        BoxCode = TextBox27.Text
        BoxMesMessage()
    End Sub

    '出库日志
    'Sub AddOperLog(ByVal barcode As String, ByVal opertype As String, ByVal operorder As Long, ByVal plateno As Long, ByVal operno As Long)
    '    Dim StrErr As String
    '    Dim SQL1() As String
    '    ReDim SQL1(1)
    '    SQL1(1) = "insert into oper_log (barcode,operator,oper_time,oper_type,oper_order,oper_class,plate_no,oper_no) values " & _
    '                  "('" & barcode & "','" & NowUser & "',convert(varchar(20),getdate(),120),'" & opertype & "'," & operorder & ",'" & NowClass & "'," & plateno & "," & operno & ")"
    '    StrErr = ExeSQLS(SQL1, SQL)
    '    If StrErr <> "" Then MsgBox("网络连接失败！！") : Exit Sub
    'End Sub

    '入库日志
    'Sub AddOperLog(ByVal barcode As String, ByVal opertype As String, ByVal boxno As String)
    '    Dim StrErr As String
    '    Dim SQL1() As String
    '    ReDim SQL1(1)
    '    SQL1(1) = "insert into oper_log (barcode,operator,oper_time,oper_type,oper_class,box_no) values " & _
    '                  "('" & barcode & "','" & NowUser & "',convert(varchar(20),getdate(),120),'" & opertype & "','" & NowClass & "','" & boxno & "')"
    '    StrErr = ExeSQLS(SQL1, SQL)
    '    If StrErr <> "" Then MsgBox("网络连接失败！！") : Exit Sub
    'End Sub

    '入库取消日志
    'Sub AddOperLog(ByVal barcode As String, ByVal opertype As String, ByVal operno As Long)
    '    Dim StrErr As String
    '    Dim SQL1() As String
    '    ReDim SQL1(1)
    '    SQL1(1) = "insert into oper_log (barcode,operator,oper_time,oper_type,oper_class,oper_no) values " & _
    '                  "('" & barcode & "','" & NowUser & "',convert(varchar(20),getdate(),120),'" & opertype & "','" & NowClass & "'," & operno & ")"
    '    StrErr = ExeSQLS(SQL1, SQL)
    '    If StrErr <> "" Then MsgBox("网络连接失败！！") : Exit Sub
    'End Sub

    '返入库日志
    'Sub AddOperLog(ByVal barcode As String, ByVal opertype As String, ByVal operno As Long, ByVal boxno As String)
    '    Dim StrErr As String
    '    Dim SQL1() As String
    '    ReDim SQL1(1)
    '    SQL1(1) = "insert into oper_log (barcode,operator,oper_time,oper_type,oper_class,oper_no,box_no) values " & _
    '                  "('" & barcode & "','" & NowUser & "',convert(varchar(20),getdate(),120),'" & opertype & "','" & NowClass & "'," & operno & ",'" & boxno & "')"
    '    StrErr = ExeSQLS(SQL1, SQL)
    '    If StrErr <> "" Then MsgBox("网络连接失败！！") : Exit Sub
    'End Sub

    '入库取消日志
    'Sub AddOperLog(ByVal barcode As String, ByVal opertype As String)
    '    Dim StrErr As String
    '    Dim SQL1() As String
    '    ReDim SQL1(1)
    '    SQL1(1) = "insert into oper_log (barcode,operator,oper_time,oper_type,oper_class) values " & _
    '                  "('" & barcode & "','" & NowUser & "',convert(varchar(20),getdate(),120),'" & opertype & "','" & NowClass & "')"
    '    StrErr = ExeSQLS(SQL1, SQL)
    '    If StrErr <> "" Then MsgBox("网络连接失败！！") : Exit Sub
    'End Sub

    '上传金蝶日志
    'Sub AddOperLog(ByVal flag As String)
    '    Dim StrErr As String
    '    Dim SQL1() As String
    '    ReDim SQL1(1)
    '    SQL1(1) = "insert into oper_log (barcode,operator,oper_time,oper_type,oper_class,box_no,oper_no) " & _
    '                  " select barcode,InMan,convert(varchar(20),getdate(),120),'上传金蝶','" & NowClass & "',boxcode,incode from hand_store where InMan='" & NowUser & "' and  InDate =convert(varchar(10),DATEADD(day," & flag & ",convert(varchar(10),getdate(),120)),120) and flag=0"
    '    StrErr = ExeSQLS(SQL1, SQL)
    '    If StrErr <> "" Then MsgBox("网络连接失败！！") : Exit Sub
    'End Sub

    Private Sub Button67_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            If SQL.State <> Data.ConnectionState.Open Then SQL.Open()
        Catch ex As Exception
            'ConSQL = "网络连接失败,无法连接数据库!!"
            SetLog(ex.Message & vbCrLf & ex.StackTrace)
            MsgBox(ex.Message)
            MsgBox(ex.StackTrace)
        End Try
    End Sub

    Private Sub Button73_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button73.Click
        ShowPanel(19)
        Dim StrSQL As String
        Dim StrErr As String
        StrSQL = "select FBillID from ICBillNo where FBillName LIKE '产品入库%' "
        Dim Arr(,)
        StrErr = GetRst(StrSQL, Arr, SQLK3)
        FBillID = Arr(1, 1)
        QtInType = 1
    End Sub
    '20180602新增其它入库
    Private Sub Button68_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        ShowPanel(24)
        Label93.Text = "外购入库>"
        Label82.Text = "外购入库>"
        Dim StrSQL As String
        Dim StrErr As String
        StrSQL = "select FBillID from ICBillNo where FBillName LIKE '外购入库%' "
        Dim Arr(,)
        StrErr = GetRst(StrSQL, Arr, SQLK3)
        FBillID = Arr(1, 1)
        QtInType = 2
    End Sub
    '20180602新增其它入库
    Private Sub Button76_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button76.Click
        ShowPanel(22)
    End Sub
    '20180602新增其它入库
    Private Sub Button74_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button74.Click
        If TextBox29.Text = "" Then ShowBoxLabel("请扫描5位条码!", Color.Red) : Exit Sub
        If TextBox29.Text.Length <> 5 Then ShowBoxLabel("请扫描5位条码！", Color.Red) : Exit Sub
        GetStockList("elseIn")
        BoxCode = TextBox29.Text.Trim
        ShowPanel(23)
    End Sub
    '20180602新增其它入库
    Private Sub Button72_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button72.Click
        ShowPanel(24)
    End Sub
    '20180602新增其它入库
    Private Sub Button69_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button69.Click
        ShowPanel(1)
    End Sub
    '20180602新增其它入库
    Private Sub Button67_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button67.Click
        ShowPanel(24)
        Label93.Text = "其它入库>"
        Label82.Text = "其它入库>"
        Dim StrSQL As String
        Dim StrErr As String
        StrSQL = "select FBillID from ICBillNo where FBillName LIKE '其他入库%' "
        Dim Arr(,)
        StrErr = GetRst(StrSQL, Arr, SQLK3)
        FBillID = Arr(1, 1)
        QtInType = 3
    End Sub

    Private Sub Button71_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button71.Click
        QtInStore()
    End Sub

    Private Sub Button75_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button75.Click
        QtInMessage()
    End Sub

    Private Sub Button70_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button70.Click
        ShowPanel(1)
    End Sub

    Private Sub Button78_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button78.Click
        ReturnOutStore()
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        GetRdata()
        TextBox30.Text = ""
        DataGrid5.DataSource = Nothing
    End Sub

    Sub GetRdata()  '''''''''获取单据信息

        Dim StrErr As String
        Dim Arr(,)
        Dim Arr2(,)

        Dim StrErr2 As String
        StrErr2 = GetRst("select FInterID,FCustID from SEOutStock where FBillNo ='" & ComboBox2.Text & "'", Arr2, SQLK3)
        If StrErr2 <> "" Then MsgBox("数据库连接失败！！") : Exit Sub
        If UBound(Arr2, 2) < 1 Then MsgBox("没有该订单信息！！") : Exit Sub
        FBillNo = ComboBox18.Text
        OrderID = Arr2(1, 1)
        FCustID = Arr2(2, 1)
        StrErr = Me.GetRst("select a.FDetailID,b.Fname,a.FAuxQty,0,0 ,b.FItemID,a.FUnitID,a.FPrice,a.FEntryID,a.FInterID from SEOutStockEntry a left join t_icitem b on a.FItemID=b.FItemID where a.FInterID =" & OrderID, ArrO, SQLK3)

        If StrErr <> "" Then MsgBox("获取订单信息错误" & vbCrLf & StrErr) : Exit Sub
        If UBound(ArrO, 2) = 0 Then MsgBox("无订单信息") : Exit Sub

        '2016-07-20 修改熟读
        StrErr = Me.GetRst("select ProductID,COUNT(*) from hand_store where storestate = '在库' and instore_type=4 and OrderID = " & OrderID & " GROUP BY ProductID", Arr, SQL)

        If StrErr <> "" Then MsgBox("获取发货信息错误" & vbCrLf & StrErr) : Exit Sub
        'Dim Have As Boolean
        Dim NoneCount As Long = 0
        Dim AllCount As Long
        Dim OutCount As Long

        For t = 1 To UBound(Arr, 2)
            For p = 1 To UBound(ArrO, 2)
                If Arr(1, t) = ArrO(6, p) Then
                    'Have = True
                    ArrO(4, p) = Arr(2, t)
                End If
            Next
        Next


        For t = 1 To UBound(ArrO, 2)
            AllCount = AllCount + ArrO(3, t)
            OutCount = OutCount + ArrO(4, t)
        Next


        Dim ds As New DataSet
        Dim dt As New DataTable
        dt.Columns.Add("规格型号", Type.GetType("System.String"))
        dt.Columns.Add("应扫", Type.GetType("System.Int32"))
        dt.Columns.Add("实扫", Type.GetType("System.Int32"))
        For t = 1 To UBound(ArrO, 2)
            Dim dw = dt.NewRow
            dw.Item(0) = ArrO(2, t)
            dw.Item(1) = ArrO(3, t) - 0
            dw.Item(2) = ArrO(4, t)
            dt.Rows.Add(dw)

        Next
        ds.Tables.Add(dt)
        DataGrid4.DataSource = ds.Tables(0)
        'DataGrid2.DataSource = ds.Tables(0)

        '''''''''''''修改列宽

        DataGrid4.TableStyles.Clear()
        DataGrid4.TableStyles.Add(New DataGridTableStyle)
        DataGrid4.TableStyles.Item(0).MappingName = dt.TableName
        DataGrid4.TableStyles(0).GridColumnStyles.Item(0).Width = 140
        DataGrid4.TableStyles(0).GridColumnStyles.Item(1).Width = 35
        DataGrid4.TableStyles(0).GridColumnStyles.Item(2).Width = 35

    End Sub

    Private Sub Button77_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button77.Click
        TextBox5.Visible = False
        ShowPanel(22)
    End Sub

    Private Sub Button68_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button68.Click
        If ComboBox2.Text = "" Then MsgBox("请选择订单！！") : Exit Sub
        GetStockList("returnOut")
        ShowPanel(26)
    End Sub

    Private Sub Button80_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button80.Click
        ShowPanel(25)
        GetRdata()
    End Sub

    Private Sub Button79_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button79.Click
        TextBox5.Visible = False
        ShowPanel(19)
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged
        Dim Have As Boolean = False

        For t = 1 To UBound(stockArr, 2)
            If stockArr(1, t) = ComboBox2.Text Then
                Have = True
                StockID = stockArr(2, t)
            End If
        Next

        If Have = False Then MsgBox("请选择出库库区！！") : Exit Sub
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox4.SelectedIndexChanged
        Dim Have As Boolean = False

        For t = 1 To UBound(stockArr, 2)
            If stockArr(1, t) = ComboBox4.Text Then
                Have = True
                StockID = stockArr(2, t)
            End If
        Next

        If Have = False Then MsgBox("请选择入库库区！！") : Exit Sub
    End Sub

    Private Sub ComboBox6_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox6.SelectedIndexChanged
        Dim Have As Boolean = False

        For t = 1 To UBound(stockArr, 2)
            If stockArr(1, t) = ComboBox6.Text Then
                Have = True
                StockID = stockArr(2, t)
            End If
        Next

        If Have = False Then MsgBox("请选择入库库区！！") : Exit Sub
    End Sub

    Private Sub ComboBox7_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox7.SelectedIndexChanged
        Dim Have As Boolean = False

        For t = 1 To UBound(stockArr, 2)
            If stockArr(1, t) = ComboBox7.Text Then
                Have = True
                StockID = stockArr(2, t)
            End If
        Next

        If Have = False Then MsgBox("请选择退库库区！！") : Exit Sub
    End Sub

    Private Sub ComboBox8_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox8.SelectedIndexChanged
        Dim Have As Boolean = False

        For t = 1 To UBound(ReasonArr, 2)
            If ReasonArr(1, t) = ComboBox8.Text Then
                Have = True
                ReasonID = ReasonArr(2, t)
            End If
        Next

        If Have = False Then MsgBox("请选择取消原因！！") : Exit Sub
    End Sub 

    Private Sub Button32_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button32.Click
        Dim Arr(,)
        Dim StrErr As String
        Dim StrSql As String
        StrSql = "select 1,count(*) from hand_store where InClass='" & NowClass & "' and InDate=convert(varchar(10),getdate(),120) and instore_type = 3 UNION select 2,count(*) from hand_store where InMan='" & NowUser & "' and InDate=convert(varchar(10),getdate(),120) and instore_type = 3 UNION select 3,count(*) from hand_store where boxcode='" & BoxCode & "'"
        StrErr = GetRst(StrSql, Arr, SQL)
        If StrErr <> "" Then ShowInLabel(StrErr, Color.Red) : Exit Sub

        Label109.Text = Arr(2, 2)
        Label108.Text = Arr(2, 1)
        Label106.Text = Arr(2, 3)

        Label110.Visible = True
        Label111.Visible = True
        Label107.Visible = True
    End Sub
End Class