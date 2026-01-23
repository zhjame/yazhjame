<%@ Language="VBScript" %>
<% Response.CodePage = 936 %> <!-- 强制GB2312编码（关键修复乱码） -->
<% Response.Charset = "GB2312" %>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="GB2312"> <!-- 与ASP编码统一，彻底解决乱码 -->
    <title>设备数据查询与统计 - 雅安市纵横计算机网络有限公司</title>
    <link rel="icon" href="img/zhlogo32.png" sizes="32x32" type="image/png">
    <style>
        /* 完全复用官网样式，无任何修改 */
        body { margin: 0; padding: 0; font-family: "Microsoft YaHei"; }
        .header { background: #0066CC; color: white; padding: 20px; text-align: center; }
        .nav { background: #F5F5F5; padding: 10px; text-align: center; }
        .nav a { margin: 0 15px; color: #333; text-decoration: none; }
        .nav a:hover { color: #0066CC; }
        .content { width: 1000px; margin: 30px auto; }
        .section { margin-bottom: 50px; }
        .section h2 { color: #0066CC; border-bottom: 2px solid #E0E0E0; padding-bottom: 10px; }
        .footer { background: #333; color: white; text-align: center; padding: 15px; margin-top: 50px; }
        
        /* 仅保留数据页面必要样式（贴合官网） */
        .search-bar { background: #F9F9F9; padding: 20px; margin-bottom: 30px; border: 1px solid #E0E0E0; }
        .search-bar input, .search-bar select, .search-bar button { padding: 8px; font-size: 14px; border: 1px solid #ddd; margin-right: 10px; font-family: "Microsoft YaHei"; }
        .search-bar button { background: #0066CC; color: white; border: none; cursor: pointer; padding: 8px 15px; }
        .search-bar a { color: #0066CC; text-decoration: none; margin-left: 10px; }
        
        .stats { display: flex; gap: 30px; margin-bottom: 30px; }
        .stat-card { flex: 1; background: #F9F9F9; padding: 20px; border: 1px solid #E0E0E0; text-align: center; }
        .stat-card h4 { color: #333; margin: 0 0 10px 0; }
        .stat-card .num { color: #0066CC; font-size: 24px; font-weight: bold; }
        
        .data-table { width: 100%; border-collapse: collapse; margin-bottom: 30px; }
        .data-table th, .data-table td { border: 1px solid #E0E0E0; padding: 12px; text-align: center; }
        .data-table th { background: #F5F5F5; color: #333; font-weight: bold; }
        .data-table tr:hover { background: #F9F9F9; }
        .empty { color: #999; }
        .table-title { color: #333; font-size: 16px; margin-bottom: 15px; }
    </style>
</head>
<body>
    <!-- 完全复用官网头部导航 -->
    <div class="header">
        <h1>雅安市纵横计算机网络有限公司</h1>
        <p>诚信为本 · 您的满意是我们最大的荣誉</p>
    </div>
  <div class="nav">
        <a href="index.asp">首页</a>
        <a href="about.asp">公司简介</a>
        <a href="business.asp">业务范围</a>
        <a href="cooperate.asp">合作客户</a>
        <a href="mytool.asp">税价计算器</a>
        <a href="contact.asp">联系我们</a>
        <a href="cyaddr.asp" target="_blank">常用网址</a>
    </div>

    <div class="content">
        <!-- 页面标题（贴合官网h2样式） -->
        <div class="section">
            <h2>设备数据查询与统计</h2>
        </div>
        
        <!-- 搜索+筛选功能 -->
        <div class="search-bar">
            <form method="get" action="use_device_data_fix.asp">
                <input type="text" name="searchName" placeholder="输入设备名称搜索（如：摄像头）" value="<%= Request.QueryString("searchName") %>">
                <select name="brand">
                    <option value="">所有品牌</option>
                    <option value="国产" <% If Request.QueryString("brand")="国产" Then Response.Write "selected" End If %>>国产</option>
                    <option value="TP" <% If Request.QueryString("brand")="TP" Then Response.Write "selected" End If %>>TP</option>
                    <option value="大华" <% If Request.QueryString("brand")="大华" Then Response.Write "selected" End If %>>大华</option>
                    <option value="华三" <% If Request.QueryString("brand")="华三" Then Response.Write "selected" End If %>>华三</option>
                    <option value="希捷" <% If Request.QueryString("brand")="希捷" Then Response.Write "selected" End If %>>希捷</option>
                </select>
                <button type="submit">查询筛选</button>
                <a href="use_device_data_fix.asp">重置</a>
            </form>
        </div>
        
        <%
            ' 自定义MyIIF函数
            Function MyIIF(condition, valTrue, valFalse)
                If condition Then
                    MyIIF = valTrue
                Else
                    MyIIF = valFalse
                End If
            End Function
            
            ' 核心配置（复用原有路径）
            Dim dbPath, conn, connStr, sql, rs
            dbPath = "D:\myweb\db\mywebsite.db"
            connStr = "Driver=SQLite3 ODBC Driver;Database=" & dbPath & ";SyncPragma=NORMAL;FKSupport=ON;"
            
            ' 连接数据库
            Set conn = Server.CreateObject("ADODB.Connection")
            conn.Open connStr
            
            ' 数据统计（总设备数、总金额、平均单价）
            Set rsStat = conn.Execute("SELECT COUNT(*) AS total_count, SUM(total_price) AS total_money, AVG(unit_price) AS avg_price FROM device_list_市场")
        %>
        
        <!-- 统计结果 -->
        <div class="stats">
            <div class="stat-card">
                <h4>总设备数量</h4>
                <div class="num"><%= rsStat("total_count") %></div>
            </div>
            <div class="stat-card">
                <h4>设备总金额（元）</h4>
                <div class="num"><%= FormatNumber(rsStat("total_money"), 2) %></div>
            </div>
            <div class="stat-card">
                <h4>平均单价（元）</h4>
                <div class="num"><%= FormatNumber(rsStat("avg_price"), 2) %></div>
            </div>
        </div>
        
        <%
            ' 条件查询逻辑
            sql = "SELECT * FROM device_list_市场 WHERE 1=1"
            If Request.QueryString("searchName") <> "" Then
                sql = sql & " AND device_name LIKE '%" & Replace(Request.QueryString("searchName"), "'", "''") & "%'"
            End If
            If Request.QueryString("brand") <> "" Then
                sql = sql & " AND brand = '" & Replace(Request.QueryString("brand"), "'", "''") & "'"
            End If
            sql = sql & " ORDER BY id ASC"
            
            ' 执行查询
            Set rs = conn.Execute(sql)
        %>
        
        <!-- 数据表格 -->
        <div class="table-title">设备数据列表（共<%= rs.RecordCount %>条）</div>
        <table class="data-table">
            <tr>
                <<th>ID</</th>
                <<th>序号</</th>
                <<th>设备名称</</th>
                <<th>品牌</</th>
                <<th>型号</</th>
                <<th>单位</</th>
                <<th>数量</</th>
                <<th>单价（元）</</th>
                <<th>总价（元）</</th>
            </tr>
            <% If rs.EOF Then %>
                <tr>
                    <td colspan="9" class="empty">暂无匹配数据</td>
                </tr>
            <% Else %>
                <% Do While Not rs.EOF %>
                <tr>
                    <td><%= rs("id") %></td>
                    <td><%= MyIIF(rs("seq_num")="", "-", rs("seq_num")) %></td>
                    <td><%= rs("device_name") %></td>
                    <td><%= rs("brand") %></td>
                    <td><%= MyIIF(rs("model")="", "-", rs("model")) %></td>
                    <td><%= rs("unit") %></td>
                    <td><%= rs("quantity") %></td>
                    <td><%= FormatNumber(rs("unit_price"), 2) %></td>
                    <td><%= FormatNumber(rs("total_price"), 2) %></td>
                </tr>
                <% 
                    rs.MoveNext 
                Loop %>
            <% End If %>
        </table>
        
        <%
            ' 关闭连接
            rs.Close: rsStat.Close
            conn.Close
            Set rs = Nothing: Set rsStat = Nothing: Set conn = Nothing
        %>
    </div>

    <!-- 完全复用官网页脚 -->
    <div class="footer">
        <p>地址：四川省雅安市健康路112号、126号 | 联系电话：0835-2232136 | 技术支持热线：0835-6208811，13881609876</p>
        <p>? 2025 雅安市纵横计算机网络有限公司 版权所有</p>
    </div>
</body>
</html>