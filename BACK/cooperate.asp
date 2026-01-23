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


  

    <!-- 原有合作客户内容完全保留 -->
    <div class="content">
        <h2 class="section-title">合作客户一览</h2>
        <p>自1999年成立以来，我们已为政府机关、企事业单位等5000余家客户提供专业IT服务。</p>

        <div class="client-category">
            <h3>政府机关与事业单位</h3>
            <div class="client-grid">
                <div class="client-card">
                    <div class="client-name">雅安市地税局</div>
                    <div class="合作项目">局域网升级改造<br>安防监控系统部署<br>设备定期维护服务</div>
                </div>
                <div class="client-card">
                    <div class="client-name">雅安市教育局</div>
                    <div class="合作项目">校园网络集成方案<br>多媒体教室设备安装<br>教师办公设备采购</div>
                </div>
                <div class="client-card">
                    <div class="client-name">雅安市人民医院</div>
                    <div class="合作项目">医疗设备网络搭建<br>数据备份系统部署<br>终端安全防护服务</div>
                </div>
            </div>
        </div>

        <div class="client-category">
            <h3>大型企业与集团</h3>
            <div class="client-grid">
                <div class="client-card">
                    <div class="client-name">中国移动雅安分公司</div>
                    <div class="合作项目">办公系统集成<br>视频会议设备安装<br>年度IT运维服务</div>
                </div>
                <div class="client-card">
                    <div class="client-name">雅安电力集团</div>
                    <div class="合作项目">电力监控网络搭建<br>服务器机房建设<br>网络安全防护方案</div>
                </div>
                <div class="client-card">
                    <div class="client-name">雅安市商业银行</div>
                    <div class="合作项目">网点设备采购与部署<br>数据加密系统实施<br>应急故障响应服务</div>
                </div>
            </div>
        </div>

        <div>
            <h2 class="section-title">典型合作案例</h2>
            <div class="case-showcase">
                <div class="case-item">
                    <div class="case-img">项目现场图</div>
                    <div class="case-content">
                        <h4>中国移动雅安分公司全业务运维项目</h4>
                        <p><strong>项目内容：</strong>为12个营业网点提供全方位IT运维服务，包括网络设备管理、终端故障处理、数据备份与恢复等。</p>
                        <p><strong>实施成果：</strong>系统故障率降低65%，故障响应时间缩短至2小时内，年度运维成本降低30%。</p>
                    </div>
                </div>
                <div class="case-item">
                    <div class="case-img">项目现场图</div>
                    <div class="case-content">
                        <h4>雅安市地税局网络安全升级工程</h4>
                        <p><strong>项目内容：</strong>部署新一代防火墙、入侵检测系统，建立数据加密传输通道，实施安全审计机制。</p>
                        <p><strong>实施成果：</strong>通过国家网络安全等级保护三级认证，成功抵御12次网络攻击尝试。</p>
                    </div>
                </div>
            </div>
        </div>

        <div class="cooperation-model">
            <h2 class="section-title">我们的合作模式</h2>
            <div class="model-list">
                <div class="model-item">
                    <h4>年度运维服务</h4>
                    <p>提供固定周期的设备巡检、故障处理、系统优化等全包服务，适合长期稳定需求的客户。</p>
                </div>
                <div class="model-item">
                    <h4>项目制合作</h4>
                    <p>针对特定IT建设项目提供从设计到实施的全流程服务，按项目验收结算。</p>
                </div>
                <div class="model-item">
                    <h4>设备采购+服务捆绑</h4>
                    <p>结合硬件采购提供安装调试、培训及质保期内免费维护，实现"一站式"解决方案。</p>
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