<%@ Language="VBScript" %>  <!-- ASP基础声明（指定脚本语言，静态页面可省略但建议保留） -->
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>雅安市纵横计算机网络有限公司 - 专业IT解决方案服务商</title>
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
        <a href="mytool.asp">实用工具</a>
        <a href="contact.asp">联系我们</a>
        <a href="use_device_data_fix.asp">数据库</a>
      <a href="https://taitoubiao.com/website" target="_blank" >常用网址</a>
    </div>

    <div class="content">
        <!-- 原有首页内容完全保留 -->
        <div class="section">
            <h2>关于我们</h2>
            <p>成立于1999年，注册资本100万元，专注于为政府、军队、企业提供全方位IT解决方案，现有员工18人，技术团队8人，是联想电脑雅安指定经销商、富士施乐授权经销商。</p>
        </div>

        <div class="section">
            <h2>核心业务</h2>
            <ul>
                <li>计算机硬件销售与维修</li>
                <li>网络安全与系统集成</li>
                <li>监控报警系统与弱电工程</li>
               <a href="http://192.168.1.182"> <li>视频会议及投影产品服务</li> </a>
            </ul>
        </div>

        <div class="section">
            <h2>荣誉资质</h2>
            <p>雅安全星级联想专卖店 | 雅安家电下乡指定店面 | 富士施乐（中国）认定经销商</p>
        </div>

       
    <div class="footer">
        <p>地址：四川省雅安市健康路112号、126号 | 联系电话：0835-2232136 | 技术支持热线：0835-6208811，13881609876</p>
        <p>© 2025 雅安市纵横计算机网络有限公司 版权所有</p>
    </div>

   
</body>
</html>