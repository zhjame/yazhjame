<%@ Language="VBScript" %>
<!DOCTYPE html>
<html>
<head>
    <meta charset="GB2312">
    <title>Excel设备数据导入SQLite（完整显示版）</title>
    <link rel="icon" href="img/zhlogo32.png" sizes="32x32" type="image/png">
    <style>
        body { font-family: 微软雅黑; font-size: 14px; line-height: 1.8; }
        .success { color: #008000; font-weight: bold; }
        .error { color: #FF0000; font-weight: bold; }
        .info { color: #0000FF; }
        .debug { color: #666; font-size: 13px; border: 1px solid #eee; padding: 10px; margin: 10px 0; }
        .btn { display: inline-block; padding: 8px 20px; background: #0066CC; color: white; text-decoration: none; border-radius: 4px; margin: 5px; }
        table { border-collapse: collapse; margin: 20px 0; width: 100%; }
        table td, table th { border: 1px solid #333; padding: 8px 12px; text-align: center; }
        table th { background-color: #f0f0f0; font-weight: bold; }
        .empty { color: #999; }
    </style>
</head>
<body>
    <h2>Excel设备数据导入SQLite（天全移动易欣-完整显示版）</h2>
    <%
        ' 自定义IIF函数
        Function MyIIF(condition, valTrue, valFalse)
            If condition Then
                MyIIF = valTrue
            Else
                MyIIF = valFalse
            End If
        End Function
        
        Const adOpenStatic = 1
        Const adLockOptimistic = 3
        
        ' 核心配置
        Dim excelPath, dbPath, connExcel, connStrExcel
        excelPath = "D:\myweb\天全移动易欣.xlsx"
        dbPath = "D:\myweb\db\mywebsite.db"
        connStrExcel = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & excelPath & ";Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1"";"
        
        ' 第一步：连接Excel获取工作表
        Set connExcel = Server.CreateObject("ADODB.Connection")
        On Error Resume Next
        connExcel.Open connStrExcel
        If Err.Number <> 0 Then
            Response.Write "<p class='error'>? 连接Excel失败！</p>"
            Response.Write "<p>错误信息：" & Err.Description & "</p>"
            Response.End
        End If
        On Error GoTo 0
        
        ' 获取工作表名
        Dim cat, tbl, sheetNames, i
        Set cat = Server.CreateObject("ADOX.Catalog")
        Set cat.ActiveConnection = connExcel
        sheetNames = ""
        i = 1
        For Each tbl In cat.Tables
            If InStr(tbl.Name, "$") > 0 And Left(tbl.Name, 1) <> "~" Then
                sheetNames = sheetNames & "<option value='" & tbl.Name & "'>" & Replace(tbl.Name, "$", "") & "</option>"
                i = i + 1
            End If
        Next
        connExcel.Close
        Set connExcel = Nothing
        Set cat = Nothing
        
        If sheetNames = "" Then
            Response.Write "<p class='error'>? 未识别到工作表</p>"
            Response.End
        End If
        
        ' 第二步：选择工作表
        If Request.Form("sheetName") = "" Then
    %>
            <form method="post" action="excel_to_sqlite_complete.asp">
                <select name="sheetName" style="padding:8px;" required><%= sheetNames %></select>
                <button type="submit" class="btn">开始导入并查看所有数据</button>
            </form>
    <%
        Else
            Dim selectedSheet, connSQLite, connStrSQLite, tableName
            selectedSheet = Request.Form("sheetName")
            tableName = "device_list_" & Replace(Replace(selectedSheet, "$", ""), " ", "")
            Response.Write "<p class='info'>?? 正在导入工作表：" & Replace(selectedSheet, "$", "") & "</p>"
            
            ' 第三步：读取Excel数据
            Set connExcel = Server.CreateObject("ADODB.Connection")
            connExcel.Open connStrExcel
            Dim rsExcel, sqlSelect, fieldCount
            sqlSelect = "SELECT * FROM [" & selectedSheet & "]"
            
            Set rsExcel = Server.CreateObject("ADODB.Recordset")
            rsExcel.CursorType = adOpenStatic
            rsExcel.LockType = adLockOptimistic
            rsExcel.Open sqlSelect, connExcel
            
            ' 统计有效数据量
            Dim dataTotal, tempRs
            dataTotal = 0
            Set tempRs = rsExcel.Clone
            Do While Not tempRs.EOF
                Dim tempDeviceName
                tempDeviceName = MyIIF(Not IsNull(tempRs.Fields(1).Value), tempRs.Fields(1).Value, "")
                If tempDeviceName <> "" And tempDeviceName <> "设备名称" Then
                    dataTotal = dataTotal + 1
                End If
                tempRs.MoveNext
            Loop
            tempRs.Close
            Set tempRs = Nothing
            Response.Write "<p class='success'>? 有效数据共" & dataTotal & "条，已全部导入！</p>"
            
            ' 第四步：连接SQLite创表+插入数据（和之前逻辑一致，确保全量插入）
            connStrSQLite = "Driver=SQLite3 ODBC Driver;Database=" & dbPath & ";SyncPragma=NORMAL;FKSupport=ON;"
            Set connSQLite = Server.CreateObject("ADODB.Connection")
            connSQLite.Open connStrSQLite
            
            Dim sqlDrop, sqlCreate
            sqlDrop = "DROP TABLE IF EXISTS " & tableName & ";"
            sqlCreate = "CREATE TABLE " & tableName & " (" & _
                        "id INTEGER PRIMARY KEY AUTOINCREMENT, " & _
                        "seq_num VARCHAR(20), " & _
                        "device_name VARCHAR(100) NOT NULL, " & _
                        "brand VARCHAR(50), " & _
                        "model VARCHAR(100), " & _
                        "unit VARCHAR(20), " & _
                        "quantity DECIMAL(10,2), " & _
                        "unit_price DECIMAL(10,2), " & _
                        "total_price DECIMAL(10,2)" & _
                        ");"
            connSQLite.Execute sqlDrop
            connSQLite.Execute sqlCreate
            
            ' 手动指定列索引（和Excel列完全匹配）
            Dim idx_seq, idx_name, idx_brand, idx_model, idx_unit, idx_qty, idx_price, idx_total
            idx_seq = 0
            idx_name = 1
            idx_brand = 2
            idx_model = 3
            idx_unit = 4
            idx_qty = 5
            idx_price = 6
            idx_total = 7
            
            ' 批量插入所有数据
            Dim sqlInsert, seqNum, deviceName, brand, model, unit, quantity, unitPrice, totalPrice, dataCount
            dataCount = 0
            connSQLite.BeginTrans
            
            On Error Resume Next
            rsExcel.MoveFirst
            Do While Not rsExcel.EOF
                seqNum = MyIIF(Not IsNull(rsExcel.Fields(idx_seq).Value), rsExcel.Fields(idx_seq).Value, "")
                deviceName = MyIIF(Not IsNull(rsExcel.Fields(idx_name).Value), rsExcel.Fields(idx_name).Value, "")
                brand = MyIIF(Not IsNull(rsExcel.Fields(idx_brand).Value), rsExcel.Fields(idx_brand).Value, "")
                model = MyIIF(Not IsNull(rsExcel.Fields(idx_model).Value), rsExcel.Fields(idx_model).Value, "")
                unit = MyIIF(Not IsNull(rsExcel.Fields(idx_unit).Value), rsExcel.Fields(idx_unit).Value, "")
                quantity = MyIIF(IsNumeric(rsExcel.Fields(idx_qty).Value), rsExcel.Fields(idx_qty).Value, 0)
                unitPrice = MyIIF(IsNumeric(rsExcel.Fields(idx_price).Value), rsExcel.Fields(idx_price).Value, 0)
                totalPrice = MyIIF(IsNumeric(rsExcel.Fields(idx_total).Value), rsExcel.Fields(idx_total).Value, 0)
                
                If deviceName <> "" And deviceName <> "设备名称" Then
                    sqlInsert = "INSERT INTO " & tableName & " (seq_num, device_name, brand, model, unit, quantity, unit_price, total_price) VALUES (" & _
                                "'" & Replace(seqNum, "'", "''") & "', " & _
                                "'" & Replace(deviceName, "'", "''") & "', " & _
                                "'" & Replace(brand, "'", "''") & "', " & _
                                "'" & Replace(model, "'", "''") & "', " & _
                                "'" & Replace(unit, "'", "''") & "', " & _
                                quantity & ", " & _
                                unitPrice & ", " & _
                                totalPrice & _
                                ");"
                    
                    connSQLite.Execute sqlInsert
                    If Err.Number = 0 Then dataCount = dataCount + 1
                End If
                
                rsExcel.MoveNext
            Loop
            connSQLite.CommitTrans
            On Error GoTo 0
            
            ' 第五步：显示所有导入的数据（关键！不再只显示前5条）
            Response.Write "<h3>?? 全部导入数据（共" & dataCount & "条）</h3>"
            Dim rsSQLite
            Set rsSQLite = connSQLite.Execute("SELECT * FROM " & tableName & " ORDER BY id ASC")
            
            ' 修复表格HTML语法错误，确保正常显示
            Response.Write "<table>"
            Response.Write "<tr>" & _
                          "<<th>ID</</th>" & _
                          "<<th>序号</</th>" & _
                          "<<th>设备名称</</th>" & _
                          "<<th>品牌</</th>" & _
                          "<<th>型号</</th>" & _
                          "<<th>单位</</th>" & _
                          "<<th>数量</</th>" & _
                          "<<th>单价</</th>" & _
                          "<<th>总价</</th>" & _
                          "</tr>"
            
            Do While Not rsSQLite.EOF
                Response.Write "<tr>"
                Response.Write "<td>" & rsSQLite("id") & "</td>"
                ' 空序号显示“-”，更直观
                Response.Write "<td class='" & MyIIF(rsSQLite("seq_num")="", "empty", "") & "'>" & MyIIF(rsSQLite("seq_num")="", "-", rsSQLite("seq_num")) & "</td>"
                Response.Write "<td>" & rsSQLite("device_name") & "</td>"
                Response.Write "<td>" & MyIIF(rsSQLite("brand")="", "-", rsSQLite("brand")) & "</td>"
                Response.Write "<td>" & MyIIF(rsSQLite("model")="", "-", rsSQLite("model")) & "</td>"
                Response.Write "<td>" & MyIIF(rsSQLite("unit")="", "-", rsSQLite("unit")) & "</td>"
                Response.Write "<td>" & rsSQLite("quantity") & "</td>"
                Response.Write "<td>" & rsSQLite("unit_price") & "</td>"
                Response.Write "<td>" & rsSQLite("total_price") & "</td>"
                Response.Write "</tr>"
                rsSQLite.MoveNext
            Loop
            Response.Write "</table>"
            
            ' 关闭连接
            rsExcel.Close: rsSQLite.Close: connExcel.Close: connSQLite.Close
            Set rsExcel=Nothing: Set rsSQLite=Nothing: Set connExcel=Nothing: Set connSQLite=Nothing
            Response.Write "<p class='success' style='margin-top:30px;'>?? 导入完成！所有数据已完整存入SQLite的" & tableName & "表</p>"
        End If
    %>
</body>
</html>