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
        <h2 class="section-title">联系我们</h2>
        <p>无论您有设备采购、技术支持或项目合作需求，欢迎通过以下方式与我们联系，我们将在24小时内回复您的咨询。</p>
        <div class="contact-container">
            <div class="contact-info">
                <h3>联系方式</h3>
                <div class="info-item">
                    <div class="icon">📍</div>
                    <div class="text">
                        <h4>公司地址</h4>
                        <p>主店：四川省雅安市健康路112号</p>
                        <div class="nav-btns">
                            <a href="https://uri.amap.com/marker?position=103.004052,29.979613&name=雅安市纵横计算机网络有限公司（主店）&address=四川省雅安市健康路112号&coordinate=gaode&callnative=1" class="nav-btn amap-btn" target="_blank">高德导航</a>
                            <a href="https://api.map.baidu.com/marker?location=29.979613,103.004052&title=雅安市纵横计算机网络有限公司（主店）&content=四川省雅安市健康路112号&output=html&src=公司官网" class="nav-btn bmap-btn" target="_blank">百度导航</a>
                        </div>
                        <p style="margin-top: 8px;">分店：四川省雅安市健康路126号</p>
                        <div class="nav-btns">
                            <a href="https://uri.amap.com/marker?position=103.004301,29.979448&name=雅安市纵横计算机网络有限公司（分店）&address=四川省雅安市健康路126号&coordinate=gaode&callnative=1" class="nav-btn amap-btn" target="_blank">高德导航</a>
                            <a href="https://api.map.baidu.com/marker?location=29.979448,103.004301&title=雅安市纵横计算机网络有限公司（分店）&content=四川省雅安市健康路126号&output=html&src=公司官网" class="nav-btn bmap-btn" target="_blank">百度导航</a>
                        </div>
                    </div>
                </div>
                <div class="info-item">
                    <div class="icon">📞</div>
                    <div class="text">
                        <h4>联系电话</h4>
                        <p>业务咨询：0835-2232136<br>技术支持：0835-6208811，13881609876</p>
                    </div>
                </div>
                <div class="info-item">
                    <div class="icon">⏰</div>
                    <div class="text">
                        <h4>营业时间</h4>
                        <p>周一至周五：9:00 - 18:00<br>周六至周日：10:00 - 16:00（节假日除外）</p>
                    </div>
                </div>
                <div class="info-item">
                    <div class="icon">✉️</div>
                    <div class="text">
                        <h4>电子邮箱</h4>
                        <p>业务合作：10520778@qq.com<br>技术咨询：248769886@qq.com</p>
                    </div>
                </div>
                <!-- 核心修改：替换占位符为百度地图（定位到健康路126号） -->
                
            </div>
            <div class="contact-form">
                <h3>在线留言</h3>
                <form action="#" method="POST">
                    <div class="form-group">
                        <label for="name">您的姓名</label>
                        <input type="text" id="name" name="name" placeholder="请输入您的姓名" required>
                    </div>
                    <div class="form-group">
                        <label for="phone">联系电话</label>
                        <input type="tel" id="phone" name="phone" placeholder="请输入您的联系电话" required>
                    </div>
                    <div class="form-group">
                        <label for="email">电子邮箱</label>
                        <input type="email" id="email" name="email" placeholder="请输入您的邮箱地址" required>
                    </div>
                    <div class="form-group">
                        <label for="message">留言内容</label>
                        <textarea id="message" name="message" placeholder="请描述您的需求（如设备采购、技术支持、项目合作等）" required></textarea>
                    </div>
                    <button type="submit" class="submit-btn">提交留言</button>
                </form>
            </div>
        </div>
    </div>
    <div class="footer">
        <p>地址：四川省雅安市健康路112号、126号 | 联系电话：0835-2232136 | 技术支持热线：0835-6208811，13881609876</p>
        <p>© 2025 雅安市纵横计算机网络有限公司 版权所有</p>
    </div>
</body>
</html>