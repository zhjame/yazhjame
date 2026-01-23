<%@ Language="VBScript" %>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>雅安市纵横计算机网络有限公司 - 专业IT解决方案服务商</title>
    <link rel="icon" href="img/zhlogo32.png" sizes="32x32" type="image/png">
    <style>
        /* 全局样式 */
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: "Microsoft YaHei", "Helvetica Neue", sans-serif;
            color: #333;
            line-height: 1.6;
            background-color: #f9f9f9;
            scroll-behavior: smooth; /* 平滑滚动 */
        }
        
        /* 头部样式 */
        .header {
            background: #0066CC;
            color: white;
            padding: 40px 0;
            text-align: center;
            box-shadow: 0 2px 15px rgba(0,0,0,0.15);
            position: relative;
            overflow: hidden;
        }
        
        /* 头部背景装饰 */
        .header::before {
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: url("data:image/svg+xml,%3Csvg width='60' height='60' viewBox='0 0 60 60' xmlns='http://www.w3.org/2000/svg'%3E%3Cg fill='none' fill-rule='evenodd'%3E%3Cg fill='%23ffffff' fill-opacity='0.05'%3E%3Cpath d='M36 34v-4h-2v4h-4v2h4v4h2v-4h4v-2h-4zm0-30V0h-2v4h-4v2h4v4h2V6h4V4h-4zM6 34v-4H4v4H0v2h4v4h2v-4h4v-2H6zM6 4V0H4v4H0v2h4v4h2V6h4V4H6z'/%3E%3C/g%3E%3C/g%3E%3C/svg%3E");
        }
        
        .header-content {
            position: relative;
            z-index: 1;
        }
        
        .header h1 {
            font-size: 2.5rem;
            margin-bottom: 12px;
            letter-spacing: 2px;
            text-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        
        .header p {
            font-size: 1.2rem;
            opacity: 0.9;
            max-width: 800px;
            margin: 0 auto;
        }
        
        /* 导航样式 */
        .nav {
            background: white;
            padding: 15px 0;
            text-align: center;
            box-shadow: 0 2px 8px rgba(0,0,0,0.08);
            position: sticky;
            top: 0;
            z-index: 100;
            transition: all 0.3s ease;
        }
        
        /* 滚动时导航栏变化 */
        .nav.scrolled {
            padding: 10px 0;
            box-shadow: 0 4px 12px rgba(0,0,0,0.12);
        }
        
        .nav-container {
            max-width: 1200px;
            margin: 0 auto;
            padding: 0 20px;
        }
        
        .nav a {
            margin: 0 18px;
            color: #333;
            text-decoration: none;
            font-size: 1rem;
            transition: all 0.3s ease;
            position: relative;
            padding: 5px 0;
            display: inline-block;
        }
        
        .nav a:hover {
            color: #0066CC;
        }
        
        .nav a:after {
            content: '';
            position: absolute;
            width: 0;
            height: 2px;
            bottom: 0;
            left: 0;
            background-color: #0066CC;
            transition: width 0.3s ease;
        }
        
        .nav a:hover:after {
            width: 100%;
        }
        
        .nav a:last-child {
            color: #0066CC;
            font-weight: 500;
        }
        
        /* 内容区域样式 */
        .content {
            max-width: 1200px;
            margin: 50px auto;
            padding: 0 20px;
        }
        
        .section {
            background: white;
            border-radius: 10px;
            padding: 35px;
            margin-bottom: 45px;
            box-shadow: 0 5px 15px rgba(0,0,0,0.05);
            transition: transform 0.3s ease, box-shadow 0.3s ease;
            position: relative;
            overflow: hidden;
        }
        
        .section::before {
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            width: 100%;
            height: 4px;
            background: linear-gradient(90deg, #0066CC 0%, #4da6ff 100%);
            opacity: 0;
            transition: opacity 0.3s ease;
        }
        
        .section:hover {
            transform: translateY(-5px);
            box-shadow: 0 8px 25px rgba(0,0,0,0.1);
        }
        
        .section:hover::before {
            opacity: 1;
        }
        
        .section h2 {
            color: #0066CC;
            border-bottom: 2px solid #E0E0E0;
            padding-bottom: 15px;
            margin-bottom: 25px;
            font-size: 1.8rem;
            position: relative;
            display: inline-block;
        }
        
        .section h2:after {
            content: '';
            position: absolute;
            width: 60px;
            height: 2px;
            background-color: #0066CC;
            bottom: -2px;
            left: 0;
        }
        
        .section p {
            margin-bottom: 18px;
            font-size: 1.05rem;
            line-height: 1.8;
            color: #555;
        }
        
        .section h3 {
            color: #444;
            margin: 25px 0 15px;
            font-size: 1.3rem;
        }
        
        /* 列表样式 */
        .fact-list {
            list-style: none;
            margin: 25px 0;
        }
        
        .fact-list li {
            background: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='16' height='16' viewBox='0 0 24 24' fill='none' stroke='%230066CC' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'%3E%3Cpolyline points='20 6 9 17 4 12'%3E%3C/polyline%3E%3C/svg%3E") no-repeat left center;
            padding-left: 30px;
            margin-bottom: 15px;
            font-size: 1.05rem;
            transition: transform 0.2s ease;
        }
        
        .fact-list li:hover {
            transform: translateX(5px);
        }
        
        /* 时间线样式优化 */
        .timeline {
            position: relative;
            margin: 35px 0;
            padding-left: 40px;
        }
        
        .timeline:before {
            content: '';
            position: absolute;
            left: 10px;
            top: 0;
            bottom: 0;
            width: 2px;
            background: #0066CC;
            border-radius: 2px;
        }
        
        .timeline-item {
            position: relative;
            margin-bottom: 35px;
            padding-bottom: 10px;
            transition: transform 0.3s ease;
        }
        
        .timeline-item:hover {
            transform: translateX(5px);
        }
        
        .timeline-year {
            position: absolute;
            left: -38px;
            top: 0;
            width: 32px;
            height: 32px;
            background: #0066CC;
            color: white;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            font-weight: bold;
            font-size: 0.9rem;
            box-shadow: 0 2px 8px rgba(0,102,204,0.3);
            z-index: 1;
        }
        
        .timeline-item p {
            padding-left: 10px;
            background: #f8f9fa;
            padding: 15px;
            border-radius: 6px;
            border-left: 3px solid #0066CC;
        }
        
        /* 荣誉资质样式 */
        .honor-list {
            display: flex;
            flex-wrap: wrap;
            gap: 20px;
            margin-top: 25px;
        }
        
        .honor-item {
            flex: 1;
            min-width: 200px;
            background: #f8f9fa;
            padding: 25px 20px;
            border-radius: 8px;
            text-align: center;
            transition: all 0.3s ease;
            border: 1px solid transparent;
            position: relative;
        }
        
        .honor-item::before {
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            height: 3px;
            background: #0066CC;
            transform: scaleX(0);
            transition: transform 0.3s ease;
        }
        
        .honor-item:hover {
            background: #f0f5ff;
            transform: translateY(-5px);
            box-shadow: 0 5px 15px rgba(0,102,204,0.1);
            border-color: #e6f0ff;
        }
        
        .honor-item:hover::before {
            transform: scaleX(1);
        }
        
        .honor-item p {
            margin: 0;
            font-weight: 500;
            color: #333;
        }
        
        /* 资质图片表格样式 */
        .certificate-table {
            width: 100%;
            border-collapse: separate;
            border-spacing: 15px;
            margin-top: 25px;
        }
        
        .certificate-table td {
            border: 1px solid #f0f0f0;
            padding: 15px;
            text-align: center;
            transition: all 0.3s ease;
            border-radius: 8px;
            background: #fff;
        }
        
        .certificate-table td:hover {
            background-color: #f8f9fa;
            transform: translateY(-5px);
            box-shadow: 0 5px 15px rgba(0,0,0,0.08);
            border-color: #e6e6e6;
        }
        
        .certificate-table img {
            width: 100%;
            max-width: 200px;
            height: auto;
            border-radius: 6px;
            box-shadow: 0 3px 10px rgba(0,0,0,0.1);
            transition: transform 0.4s ease, box-shadow 0.4s ease;
            cursor: pointer;
        }
        
        .certificate-table img:hover {
            transform: scale(1.08);
            box-shadow: 0 8px 20px rgba(0,0,0,0.15);
        }
        
        .certificate-table p {
            margin-top: 12px;
            color: #333;
            font-weight: 500;
            font-size: 1rem;
        }
        
        /* 页脚样式 */
        .footer {
            background: #333;
            color: white;
            text-align: center;
            padding: 50px 20px 25px;
            margin-top: 80px;
            position: relative;
        }
        
        .footer::before {
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            width: 100%;
            height: 5px;
            background: linear-gradient(90deg, #0066CC 0%, #4da6ff 100%);
        }
        
        .footer-content {
            max-width: 1200px;
            margin: 0 auto;
        }
        
        .footer p {
            margin-bottom: 12px;
            opacity: 0.9;
            font-size: 1.05rem;
        }
        
        .footer p:last-child {
            margin-top: 30px;
            font-size: 0.9rem;
            opacity: 0.7;
            padding-top: 20px;
            border-top: 1px solid rgba(255,255,255,0.1);
        }
        
        /* 响应式设计优化 */
        @media (max-width: 768px) {
            .header h1 {
                font-size: 2rem;
            }
            
            .nav a {
                margin: 0 6px;
                font-size: 0.9rem;
                padding: 3px 0;
            }
            
            .section {
                padding: 25px 20px;
                margin-bottom: 35px;
            }
            
            .section h2 {
                font-size: 1.5rem;
            }
            
            .certificate-table {
                border-spacing: 10px;
            }
            
            .certificate-table td {
                display: block;
                width: 100%;
                border: none;
                padding: 15px;
            }
            
            .timeline {
                padding-left: 30px;
            }
            
            .honor-list {
                gap: 15px;
            }
            
            .honor-item {
                min-width: 100%;
            }
        }
        
        /* 图片查看器样式 */
        .image-viewer {
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(0,0,0,0.9);
            display: none;
            justify-content: center;
            align-items: center;
            z-index: 1000;
            padding: 20px;
        }
        
        .image-viewer.active {
            display: flex;
        }
        
        .viewer-content {
            position: relative;
            max-width: 90%;
            max-height: 90%;
        }
        
        .viewer-image {
            max-width: 100%;
            max-height: 80vh;
            border: 5px solid white;
            border-radius: 5px;
        }
        
        .close-viewer {
            position: absolute;
            top: -40px;
            right: -40px;
            color: white;
            font-size: 30px;
            cursor: pointer;
            transition: transform 0.3s ease;
        }
        
        .close-viewer:hover {
            transform: rotate(90deg);
        }
        
        /* 滚动动画 */
        @keyframes fadeInUp {
            from {
                opacity: 0;
                transform: translateY(20px);
            }
            to {
                opacity: 1;
                transform: translateY(0);
            }
        }
        
        .fade-in {
            animation: fadeInUp 0.6s ease forwards;
            opacity: 0;
        }
        
        .delay-1 { animation-delay: 0.1s; }
        .delay-2 { animation-delay: 0.2s; }
        .delay-3 { animation-delay: 0.3s; }
    </style>
</head>
<body>
    <div class="header">
        <div class="header-content">
            <h1>雅安市纵横计算机网络有限公司</h1>
            <p>诚信为本 · 您的满意是我们最大的荣誉</p>
        </div>
    </div>

    <div class="nav" id="mainNav">
        <div class="nav-container">
            <a href="index.asp">首页</a>
            <a href="about.asp">公司简介</a>
            <a href="business.asp">业务范围</a>
            <a href="cooperate.asp">合作客户</a>
            <a href="mytool.asp">税价计算器</a>
            <a href="contact.asp">联系我们</a>
            <a href="cyaddr.asp" target="_blank">常用网址</a>
        </div>
    </div>

    <div class="content">
        <div class="section fade-in">
            <h2>公司概况</h2>
            <p>雅安市纵横计算机网络有限公司成立于1999年，是经雅安市工商行政管理局批准注册的专业IT服务企业，注册资本100万元人民币。公司坐落于雅安市健康路核心商圈，现有两处经营场所（健康路112号、126号），总面积达200余平方米。</p>
            
            <h3>核心数据</h3>
            <ul class="fact-list">
                <li>成立时间：1999年（深耕IT行业25年）</li>
                <li>注册资本：100万元人民币</li>
                <li>员工规模：18人（其中技术团队8人，均持有专业认证）</li>
                <li>经营面积：200余平方米（两处门店）</li>
                <li>年服务客户：超800家次</li>
            </ul>
        </div>

        <div class="section fade-in delay-1">
            <h2>发展历程</h2>
            <div class="timeline">
                <div class="timeline-item">
                    <div class="timeline-year">1999</div>
                    <p>公司正式成立，初期以计算机销售及维修为核心业务。</p>
                </div>
                <div class="timeline-item">
                    <div class="timeline-year">2005</div>
                    <p>拓展网络工程业务，完成雅安市多个政府部门局域网搭建项目。</p>
                </div>
                <div class="timeline-item">
                    <div class="timeline-year">2010</div>
                    <p>成为联想电脑授权经销商，同年荣获"雅安全星级联想专卖店"称号。</p>
                </div>
                <div class="timeline-item">
                    <div class="timeline-year">2015</div>
                    <p>业务升级至系统集成领域，承接雅安电力集团等大型企业IT运维项目。</p>
                </div>
                <div class="timeline-item">
                    <div class="timeline-year">2020</div>
                    <p>拓展网络安全、视频会议等增值服务，与中国移动雅安分公司等建立长期合作。</p>
                </div>
            </div>
        </div>

        <div class="section fade-in delay-2">
            <h2>团队与理念</h2>
            <p>公司核心团队由8名资深技术人员组成，平均从业年限超10年，持有CCNA、MCSE等专业认证。</p>
            <ul class="fact-list">
                <li>使命：以专业技术助力客户数字化转型</li>
                <li>愿景：成为雅安地区最值得信赖的IT服务品牌</li>
                <li>价值观：诚信、专业、创新、共赢</li>
            </ul>
        </div>

        <div class="section fade-in delay-3">
            <h2>荣誉资质</h2>
            <div class="honor-list">
                <div class="honor-item"><p>联想电脑全星级专卖店</p></div>
                <div class="honor-item"><p>雅安市家电下乡指定销售网点</p></div>
                <div class="honor-item"><p>富士施乐（中国）授权经销商</p></div>
                <div class="honor-item"><p>雅安市中小企业IT服务示范单位</p></div>
            </div>
        </div>

        <div class="section fade-in">
            <h2>资质证书展示</h2>
            <!-- 资质图片表格 -->
            <table class="certificate-table">
                <tr>
                    <td>
                        <img src="img/01.jpg" alt="雅安消协认定诚信单位" class="cert-img">
                        <p>雅安消协认定诚信单位</p>
                    </td>
                    <td>
                        <img src="img/02.jpg" alt="联想售后服务站" class="cert-img">
                        <p>联想售后服务站</p>
                    </td>
                    <td>
                        <img src="img/03.jpg" alt="联想专卖店" class="cert-img">
                        <p>联想专卖店</p>
                    </td>
                </tr>
                <tr>
                    <td>
                        <img src="img/04.jpg" alt="联想经销商" class="cert-img">
                        <p>联想经销商</p>
                    </td>
                    <td>
                        <img src="img/05.jpg" alt="富士施乐经销商" class="cert-img">
                        <p>富士施乐认定经销商</p>
                    </td>
                    <td>
                        <img src="img/06.jpg" alt="家电下乡指定经销商" class="cert-img">
                        <p>家电下乡指定经销商</p>
                    </td>
                </tr>
                <tr>
                    <td>
                        <img src="img/07.jpg" alt="奥图码经销商" class="cert-img">
                        <p>奥图码经销商</p>
                    </td>
                    <td>
                        <img src="img/08.jpg" alt="浪潮服务器经销商" class="cert-img">
                        <p>浪潮服务器经销商</p>
                    </td>
                    <td>
                        <img src="img/09.jpg" alt="公司彩页" class="cert-img">
                        <p>公司彩页</p>
                    </td>
                </tr>
                <tr>
                    <td>
                        <img src="img/10.jpg" alt="公司证书" class="cert-img">
                        <p>公司证书</p>
                    </td>
                    <td>
                        <img src="img/11.jpg" alt="门店外观" class="cert-img">
                        <p>门店外观</p>
                    </td>
                    <td></td>
                </tr>
            </table>
        </div>
    </div>

    <!-- 图片查看器 -->
    <div class="image-viewer" id="imageViewer">
        <div class="viewer-content">
            <span class="close-viewer" id="closeViewer">&times;</span>
            <img src="" alt="证书大图" class="viewer-image" id="viewerImage">
        </div>
    </div>

    <div class="footer">
        <div class="footer-content">
            <p>地址：四川省雅安市健康路112号、126号</p>
            <p>联系电话：0835-2232136 | 技术支持热线：0835-6208811，13881609876</p>
            <p>© 2025 雅安市纵横计算机网络有限公司 版权所有</p>
        </div>
    </div>

    <script>
        // 导航栏滚动效果
        window.addEventListener('scroll', function() {
            const nav = document.getElementById('mainNav');
            if (window.scrollY > 50) {
                nav.classList.add('scrolled');
            } else {
                nav.classList.remove('scrolled');
            }
        });

        // 图片查看器功能
        const imageViewer = document.getElementById('imageViewer');
        const viewerImage = document.getElementById('viewerImage');
        const closeViewer = document.getElementById('closeViewer');
        const certImages = document.querySelectorAll('.cert-img');

        certImages.forEach(img => {
            img.addEventListener('click', function() {
                viewerImage.src = this.src;
                imageViewer.classList.add('active');
                document.body.style.overflow = 'hidden'; // 防止背景滚动
            });
        });

        closeViewer.addEventListener('click', function() {
            imageViewer.classList.remove('active');
            document.body.style.overflow = 'auto'; // 恢复滚动
        });

        // 点击查看器背景关闭
        imageViewer.addEventListener('click', function(e) {
            if (e.target === imageViewer) {
                imageViewer.classList.remove('active');
                document.body.style.overflow = 'auto';
            }
        });

        // 滚动动画
        function checkFade() {
            const fadeElements = document.querySelectorAll('.fade-in');
            fadeElements.forEach(element => {
                const elementTop = element.getBoundingClientRect().top;
                const elementVisible = 150;
                if (elementTop < window.innerHeight - elementVisible) {
                    element.style.opacity = '1';
                    element.style.transform = 'translateY(0)';
                }
            });
        }

        // 初始检查
        checkFade();
        // 滚动时检查
        window.addEventListener('scroll', checkFade);
    </script>
</body>
</html>