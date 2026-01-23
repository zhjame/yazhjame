<%@ Language="VBScript" %>  <!-- ASP基础声明（指定脚本语言，静态页面可省略但建议保留） -->
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>雅安市纵横计算机网络有限公司 - 专业IT解决方案服务商</title>
    <link rel="icon" href="img/zhlogo32.png" sizes="32x32" type="image/png">
    <style>
        /* 原有样式完全保留 */
        body { margin: 0; padding: 0; font-family: "Microsoft YaHei"; }
        .header { background: #0066CC; color: white; padding: 20px; text-align: center; }
        .nav { background: #F5F5F5; padding: 10px; text-align: center; }
        .nav a { margin: 0 15px; color: #333; text-decoration: none; }
        .nav a:hover { color: #0066CC; }
        .content { width: 1000px; margin: 30px auto; }
        .section { margin-bottom: 50px; }
        .section h2 { color: #0066CC; border-bottom: 2px solid #E0E0E0; padding-bottom: 10px; }
        .footer { background: #333; color: white; text-align: center; padding: 15px; margin-top: 50px; }

        /* 税价计算器样式保留 */
        .tax-calculator {
            background: white;
            padding: 2rem;
            border-radius: 10px;
            box-shadow: 0 0 10px rgba(0,0,0,0.1);
            margin-top: 20px;
        }
        .input-group { margin-bottom: 1.5rem; }
        .input-group label { display: block; margin-bottom: 0.5rem; color: #333; font-weight: 500; }
        .input-group input { width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px; box-sizing: border-box; font-family: "Microsoft YaHei"; }
        .tax-btn { background: #0066CC; color: white; border: none; padding: 10px 20px; border-radius: 4px; cursor: pointer; transition: background 0.3s; font-family: "Microsoft YaHei"; }
        .tax-btn:hover { background: #0052a3; }
        .precision-switch { margin: 1rem 0; display: flex; gap: 1rem; }
        .precision-btn { padding: 5px 10px; border: 1px solid #ddd; cursor: pointer; border-radius: 4px; background: #f5f5f5; }
        .precision-btn.active { background: #0066CC; color: white; }
        .tax-result { margin-top: 1.5rem; padding: 1rem; background: #f8f9fa; border-radius: 4px; }
        .verify-result { margin-top: 1rem; padding: 10px; border-radius: 4px; }
        .valid { background: #d4edda; color: #155724; padding: 5px; border-radius: 3px; }
        .invalid { background: #f8d7da; color: #721c24; padding: 5px; border-radius: 3px; }
        .tax-error { color: #dc3545; font-size: 0.9rem; margin-top: 0.5rem; display: none; }
    </style>
</head>
<body>
    <div class="header">
        <h1>雅安市纵横计算机网络有限公司</h1>
        <p>诚信为本 · 您的满意是我们最大的荣誉</p>
    </div>

    <!-- 导航链接：所有.html后缀改为.asp -->
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
        <div class="section">
            <h2>图片浏览（选盘符→选文件夹→预览）</h2>
        </div>
        
        <!-- 提示信息 -->
        <div class="tip">
            <strong>操作说明：</strong> 1. 选择要浏览的盘符 → 2. 选择该盘符下的文件夹 → 3. 页面将显示文件夹内所有图片（支持JPG、PNG、GIF、BMP格式）；点击图片可查看大图。
        </div>
        
        <%
            ' ===================== 核心功能逻辑 =====================
            ' 自定义函数：获取系统可用盘符（排除软驱、网络驱动器）
            Function GetAvailableDrives()
                Dim fso, drive, drivesHTML
                Set fso = Server.CreateObject("Scripting.FileSystemObject")
                drivesHTML = "<option value=''>请选择盘符</option>"
                
                For Each drive In fso.Drives
                    ' 只显示本地硬盘（DriveType=2）和可移动设备（DriveType=1，如U盘）
                    If drive.IsReady And (drive.DriveType = 2 Or drive.DriveType = 1) Then
                        drivesHTML = drivesHTML & "<option value='" & drive.Path & "' " & IIF(Request.Form("drive")=drive.Path, "selected", "") & ">" & drive.Path & " (" & drive.VolumeName & ")</option>"
                    End If
                Next
                Set fso = Nothing
                GetAvailableDrives = drivesHTML
            End Function
            
            ' 自定义函数：根据盘符获取文件夹列表
            Function GetFoldersByDrive(drivePath)
                Dim fso, folder, foldersHTML
                Set fso = Server.CreateObject("Scripting.FileSystemObject")
                foldersHTML = "<option value=''>请选择文件夹</option>"
                
                If fso.FolderExists(drivePath) Then
                    For Each folder In fso.GetFolder(drivePath).SubFolders
                        ' 排除系统文件夹和隐藏文件夹
                        If Not (Left(folder.Name, 1) = "." Or folder.Attributes And 2) Then
                            foldersHTML = foldersHTML & "<option value='" & folder.Path & "' " & IIF(Request.Form("folder")=folder.Path, "selected", "") & ">" & folder.Name & "</option>"
                        End If
                    Next
                End If
                Set fso = Nothing
                GetFoldersByDrive = foldersHTML
            End Function
            
            ' 自定义函数：根据文件夹路径获取图片列表
            Function GetImagesByFolder(folderPath)
                Dim fso, file, imageExt, imagesHTML
                imageExt = Array("jpg", "jpeg", "png", "gif", "bmp") ' 支持的图片格式
                Set fso = Server.CreateObject("Scripting.FileSystemObject")
                imagesHTML = ""
                
                If fso.FolderExists(folderPath) Then
                    On Error Resume Next
                    For Each file In fso.GetFolder(folderPath).Files
                        ' 筛选图片格式
                        If UBound(Filter(imageExt, LCase(fso.GetExtensionName(file.Name)))) <> -1 Then
                            ' 图片卡片HTML
                            imagesHTML = imagesHTML & "<div class='image-card' onclick='openModal('" & Server.URLEncode(file.Path) & "')'>"
                            imagesHTML = imagesHTML & "<img src='file:///" & Replace(file.Path, "\", "/") & "' class='image-preview' alt='" & file.Name & "' onerror='this.src=''data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTIwIiBoZWlnaHQ9IjkwIiB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCAxMjAgOTAiPgogIDxnIGZpbGw9Im5vbmUiIGZpbGwtcnVsZT0iZXZlbm9kZCI+CiAgICA8cmVjdCB3aWR0aD0iMTAwIiBoZWlnaHQ9IjgwIiByeD0iMTUiIGZpbGw9IiNlMGUwZTAiLz4KICAgIDxjaXJjbGUgY3g9IjUwIiBjeT0iNTAiIHI9IjE1IiBmaWxsPSIjMDA2NkNDIi8+CiAgICA8cGF0aCBkPSJNMzAgNzBDMzAgNzAgNDUgNjAgNTAgNjBDNTUgNjAgNzAgNzAgNzAgNzAiIHN0cm9rZT0iIzAwNjZDQyIgc3Ryb2tlLXdpZHRoPSIyIi8+CiAgPC9nPgo8L3N2Zz4=';'>"
                            imagesHTML = imagesHTML & "<div class='image-info'>"
                            imagesHTML = imagesHTML & "<div class='image-name'>" & file.Name & "</div>"
                            imagesHTML = imagesHTML & "<div class='image-path'>" & Left(folderPath, 30) & "..." & "</div>"
                            imagesHTML = imagesHTML & "</div></div>"
                        End If
                    Next
                    On Error GoTo 0
                End If
                
                ' 无图片提示
                If imagesHTML = "" Then
                    imagesHTML = "<div class='error-tip'>当前文件夹中未找到支持的图片文件（支持JPG、PNG、GIF、BMP格式）</div>"
                End If
                
                Set fso = Nothing
                GetImagesByFolder = imagesHTML
            End Function
            
            ' 自定义IIF函数
            Function IIF(condition, valTrue, valFalse)
                If condition Then IIF = valTrue Else IIF = valFalse
            End Function
            
            ' 权限检查提示
            Dim permissionTip
            permissionTip = "<div class='tip'><strong>权限提示：</strong>如果无法显示文件夹/图片，请确保IIS匿名用户（IUSR）拥有该盘符/文件夹的读取权限（右键文件夹→属性→安全→编辑→添加→输入IUSR→授予读取权限）。</div>"
        %>
        
        <!-- 盘符+文件夹选择表单 -->
        <form method="post" action="image_browser.asp" class="file-selector">
            <label for="drive">1. 选择盘符：</label>
            <select id="drive" name="drive" required onchange="this.form.submit()">
                <%= GetAvailableDrives() %>
            </select>
            
            <label for="folder">2. 选择文件夹：</label>
            <select id="folder" name="folder" required onchange="this.form.submit()">
                <% If Request.Form("drive") <> "" Then %>
                    <%= GetFoldersByDrive(Request.Form("drive")) %>
                <% Else %>
                    <option value="">请先选择盘符</option>
                <% End If %>
            </select>
            
            <button type="submit">刷新图片列表</button>
            <button type="reset" class="btn-reset" onclick="window.location.href='image_browser.asp'">重置选择</button>
        </form>
        
        <!-- 权限提示 -->
        <%= permissionTip %>
        
        <!-- 图片展示区域 -->
        <div class="image-container">
            <% If Request.Form("folder") <> "" Then %>
                <h3 style="color: #333; margin-bottom: 15px;">当前路径：<%= Request.Form("folder") %></h3>
                <div class="image-grid">
                    <%= GetImagesByFolder(Request.Form("folder")) %>
                </div>
            <% Else %>
                <div class="tip">请选择盘符和文件夹，即可显示该文件夹下的所有图片</div>
            <% End If %>
        </div>
        
        <!-- 大图预览模态框 -->
        <div class="modal" id="imageModal">
            <span class="modal-close" onclick="closeModal()">&times;</span>
            <img class="modal-content" id="modalImage">
        </div>
    </div>

    <!-- 复用官网页脚 -->
    <div class="footer">
        <p>地址：四川省雅安市健康路112号、126号 | 联系电话：0835-2232136 | 技术支持热线：0835-6208811，13881609876</p>
        <p>? 2025 雅安市纵横计算机网络有限公司 版权所有</p>
    </div>

    <!-- 图片预览JS -->
    <script>
        // 模态框操作
        const modal = document.getElementById("imageModal");
        const modalImage = document.getElementById("modalImage");
        
        function openModal(imagePath) {
            modal.style.display = "flex";
            // 解码路径并转换为本地文件URL
            const decodedPath = decodeURIComponent(imagePath);
            modalImage.src = "file:///" + decodedPath.replace(/\\/g, "/");
        }
        
        function closeModal() {
            modal.style.display = "none";
            modalImage.src = ""; // 清空图片，避免缓存
        }
        
        // 点击模态框背景关闭
        window.onclick = function(event) {
            if (event.target == modal) {
                closeModal();
            }
        }
        
        // 键盘ESC关闭
        document.addEventListener("keydown", function(event) {
            if (event.key === "Escape" && modal.style.display === "flex") {
                closeModal();
            }
        });
    </script>
</body>
</html>