package com.example.myapplication

import android.os.Bundle
import android.webkit.WebView
import androidx.activity.enableEdgeToEdge
import androidx.appcompat.app.AppCompatActivity
import androidx.core.view.ViewCompat
import androidx.core.view.WindowInsetsCompat

class MainActivityztai : AppCompatActivity() {
    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        enableEdgeToEdge()
        setContentView(R.layout.activity_main_activityztai)
        ViewCompat.setOnApplyWindowInsetsListener(findViewById(R.id.main)) { v, insets ->
            val systemBars = insets.getInsets(WindowInsetsCompat.Type.systemBars())
            v.setPadding(systemBars.left, systemBars.top, systemBars.right, systemBars.bottom)
            insets
        }

        // 初始化 WebView 和加载数据
        val webView = findViewById<WebView>(R.id.webView)
        val text = """
            <html>
            <head>
                <style type="text/css">
                    body {
                        margin: 16px; /* 边距设置为16px */
                        font-family: 'Serif'; /* 设置字体为Serif，类似宋体 */
                        font-size: 16px; /* 字号设置为16px */
                    }
                    p {
                        text-indent: 2em; /* 首行缩进设置为两个字符的空间 */
                        text-align: justify; /* 文本两端对齐 */
                        line-height: 1.6; /* 行间距设置为1.6倍 */
                    }
                    img {
                        width: 100%; /* 使图片宽度适应容器宽度 */
                        height: auto; /* 保持图片的纵横比 */
                    }
                    
                </style>
            </head>
            <body>
                <p>
                    众泰汽车 (000980) 作为以汽车整车研发、制造及销售为核心业务的汽车整车制造企业，拥有众泰、江南等自主品牌，产品覆盖SUV、轿车、MPV 和新能源汽车四个细分市场。
                    凭借复制豪华车的外观造型，众泰曾经受到消费者追捧。2016年，众泰推出外形酷似保时捷Macan的车型——众泰SR9，虽然市场意见褒贬不一，但也获得不错的销量成绩。2016年，众泰汽车全年累计销量达33.31万辆。但是，随着2018年我国汽车产业出现十几年来的首次增速下滑，特别是新能源汽车补贴的退坡及补贴门槛的提高，让众泰等多家以低端新能运汽车为主、依赖补贴的车企立刻变得举步维艰。2018年众泰销量跌至15.48万辆，2019年销量仅为11.66万辆。自2019年以来，众泰一直处于亏损中。自2020年起，众泰长期处于半停产或停产状态，产能长时间利用率为零，随后大规模人员开始流失。众泰汽车于2020年6月22日召开的第七届董事会第八次会议审议通过《关于公司2019年度计提资产减值准备的议案》，公司2019年度需计提各类资产减值准备总额为84.3亿元。2020年6月24日，公司因2019年度的财务会计报告被天职国际会计师事务所（特殊普通合伙）出具了无法表示意见的审计报告。 根据《深圳证券交易所股票上市规则》第13.2.1条的相关规定，深圳证券交易所将对众泰汽车股票交易实行“退市风险警示”处理，股票简称由“众泰汽车”变更为“*ST众泰”。2020年9月，严重资不抵债的ST众泰进入破产重整，2020年12月，当时的控股股东铁牛集团也被裁定终止重整程序，并被宣告破产。2021年，根据公司 2018 年、2019 年和 2020 年年度报告，最近三个会计年度扣除非经常性损益前后净利润孰低者均为负值，且公司持续经营能力存在不确定性，根据深圳证券交易所《上市规则》第 13.3 条第（四）、（六）项的相关规定，公司股票自 2021 年 4 月 29 日起将继续叠加实施“其他风险警示”特别处理。
                    2021年10月，江苏深商控股集团有限公司向ST众泰投入重整投资款20亿元，接替铁牛集团成为控股股东。2021年12月，ST众泰重整计划执行完毕。ST众泰自2021年底完成重整后致力于汽车整车的复产工作，2022年6月，众泰汽车曾宣布拟定增募资不超60亿元，用于加码新能源项目。2022年10月20日，众泰汽车永康基地下线第一批复产的T300车型并举行下线仪式。因2021年度经审计公司的净资产为正值，且中兴财光华会计师事务所（特殊普通合伙） 对公司2021年年度财务报告出具了标准无保留意见的审计报告，众泰汽车于 2022年4月25日向深圳证券交易所申请撤销退市风险警示。根据相关规定，公司股票自 2022年5月20日开市起撤销退市风险警示并继续实施其他风险警示，股票简称由“*ST众泰”变更为“ST众泰”；公司股票自2022年11月3日开市起撤销其他风险警示，股票简称由“ST 众泰”变更为“众泰汽车”，成功摘帽。
                    系统数据显示，2021年年主营业务毛利率从-4.11%上升到10%，2022年存货周转率减少42%。
                </p>
                <img src="file:///android_res/drawable/picture_ztai_1.png" alt="示例图片" />
                <img src="file:///android_res/drawable/picture_ztai_2.png" alt="示例图片" />
                <p>
                    1）存货周转率下降的主要原因是营业成本下降53%造成的-找到原因。
                    2）2021年年主营业务毛利率从-4.11%上升到10%-找到原因。
                </p>
                <img src="file:///android_res/drawable/picture_ztai_3.png" alt="示例图片" />
                <p>
                    对公司年报数据进行分析后，得出主要其交通运输设备制造业的毛利率在收入下降的年份，毛利率还从-6.25%增长到18.04%，并且在之后的年度毛利率直线下降。
                </p>
                <img src="file:///android_res/drawable/picture_ztai_4.png" alt="示例图片" />
                <p>
                    进一步关注2021年年报：
                    2021年，经过众泰公司一年的努力，公司顺利完成了重整。报告期内公司完成销售收入 825, 170,423.45元元，同比下降38.34%，实现利润总额 -716,354,708.22元，同比减亏92.79% ，归属于上市公司股东净利润-705,532, 147.28 元，同比减93.18% 。主要原因是虽然公司2021年度已完成重整，重整计划已执行完毕，产生了重整收益，但因公司下属各汽车生产基地基本处于停产状态，2021年公司设计产能68.5万辆，报告期内因公司整车业务处于停产的状态，产能利用率0%。公司的汽车整车没有销量，销售收入总额较低，所以造成公司 2021年度经营业绩仍为亏损。同时，公司主要业务 汽车整车业务均处于停产状态，公司计提大额的资产减值准备和坏账准备等，因此公司2021年度整体业绩亏损。2021年因汽车零配件利润率高，汽车利润率为负，在汽车不生产后，反倒提升了企业的整体毛利率。
                </p>
                <img src="file:///android_res/drawable/picture_ztai_5.png" alt="示例图片" />
                <p>
                    2023年众泰汽车年报监事娄国海以公司通知时间较晚为由，对公司2023 年度报告内容存在异议或无法保证其真实、 准确、完整。
                </p>

            </body>
            </html>
        """
        webView.settings.setSupportZoom(true)
        webView.settings.builtInZoomControls = true
        webView.settings.displayZoomControls = false
        webView.loadDataWithBaseURL(null, text, "text/html", "utf-8", null)
    }
}