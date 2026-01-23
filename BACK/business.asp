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



    <!-- 原有业务范围内容完全保留 -->
    <div class="content">
        <h2 class="section-title">核心业务板块</h2>
        <p>我们专注于为客户提供全方位IT解决方案，涵盖硬件销售、系统集成、网络安全等多个领域。</p>

        <div class="business-grid">
            <div class="business-card">
                <h3>系统集成与网络工程</h3>
                <p>提供从网络规划、设计到实施的全流程服务。</p>
                <ul class="service-list">
                    <li>企业局域网（LAN）设计与部署</li>
                    <li>跨区域广域网（WAN）连接方案</li>
                    <li>无线网络（WiFi）全覆盖方案</li>
                    <li>机房建设与服务器部署</li>
                </ul>
            </div>
            <div class="business-card">
                <h3>网络安全与数据防护</h3>
                <p>针对企业数据安全需求，提供多层次防护方案。</p>
                <ul class="service-list">
                    <li>防火墙与入侵检测系统部署</li>
                    <li>数据备份与灾难恢复方案</li>
                    <li>病毒防护与终端安全管理</li>
                    <li>安全漏洞评估与修复建议</li>
                </ul>
            </div>
            <div class="business-card">
                <h3>安防监控与弱电工程</h3>
                <p>结合现代安防技术，提供智能化监控及弱电系统。</p>
                <ul class="service-list">
                    <li>高清视频监控系统安装调试</li>
                    <li>门禁考勤系统与一卡通集成</li>
                    <li>综合布线（电话/网络/视频线）</li>
                    <li>公共广播与会议音响系统</li>
                </ul>
            </div>
            <div class="business-card">
                <h3>计算机硬件销售与维修</h3>
                <p>作为联想授权经销商，提供品牌电脑及周边设备销售。</p>
                <ul class="service-list">
                    <li>联想台式机、笔记本电脑销售</li>
                    <li>打印机、复印机等办公设备销售（富士施乐授权）</li>
                    <li>电脑硬件故障检测与维修</li>
                    <li>办公设备日常维护与耗材更换</li>
                </ul>
            </div>
            <div class="business-card">
                <h3>办公自动化与软件服务</h3>
                <p>助力企业实现办公数字化，提供软件部署、定制及培训。</p>
                <ul class="service-list">
                    <li>Office办公软件安装与培训</li>
                    <li>企业OA系统部署与定制</li>
                    <li>财务软件与ERP系统实施</li>
                    <li>数据恢复与系统重装</li>
                </ul>
            </div>
            <div class="business-card">
                <h3>视频会议与投影系统</h3>
                <p>为远程协作提供解决方案，支持高清视频会议。</p>
                <ul class="service-list">
                    <li>高清视频会议终端部署</li>
                    <li>投影设备与幕布安装调试</li>
                    <li>多媒体教室与会议室系统集成</li>
                    <li>远程会议软件配置与维护</li>
                </ul>
            </div>
        </div>

        <div class="advantages">
            <h3>我们的服务优势</h3>
            <div class="advantage-list">
                <div class="advantage-item">
                    <div class="icon">✦</div>
                    <h4>专业认证团队</h4>
                    <p>8名持证技术人员，平均10年以上行业经验</p>
                </div>
                <div class="advantage-item">
                    <div class="icon">✦</div>
                    <h4>品牌授权保障</h4>
                    <p>联想、富士施乐等官方授权合作伙伴</p>
                </div>
                <div class="advantage-item">
                    <div class="icon">✦</div>
                    <h4>本地化服务</h4>
                    <p>2小时响应，市区内4小时到场维修</p>
                </div>
                <div class="advantage-item">
                    <div class="icon">✦</div>
                    <h4>定制化方案</h4>
                    <p>根据客户需求提供专属解决方案</p>
                </div>
            </div>
        </div>
    </div>

    <div class="footer">
        <p>地址：四川省雅安市健康路112号、126号 | 联系电话：0835-2232136 | 技术支持热线：0835-6208811，13881609876</p>
        <p>© 2025 雅安市纵横计算机网络有限公司 版权所有</p>
    </div>
</body>
</html>