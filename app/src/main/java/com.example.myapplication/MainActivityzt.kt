package com.example.myapplication

import android.os.Bundle
import android.webkit.WebView
import androidx.activity.enableEdgeToEdge
import androidx.appcompat.app.AppCompatActivity
import androidx.core.view.ViewCompat
import androidx.core.view.WindowInsetsCompat

class MainActivityzt : AppCompatActivity() {
    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        enableEdgeToEdge()
        setContentView(R.layout.activity_main_activityzt)
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
                    中通客车 (000957) 过去5年未ST过。
                    系统数据异常值显示2022、2023年营业收入毛利率增速明显，2023年主营业务成本率下降明显：
                </p>
                <img src="file:///android_res/drawable/picture_zt_1.png" alt="示例图片" />
                <p>
                    2023 年年报披露公司抢抓出口机遇，重点加强成本控制，实现了产品盈利能力进一步提升。主营业务利润较去年同期大幅增长。
                    全年销售客车 7,531 辆，实现营业收入 42.44 亿元。净利润下降的主要原因是上年度处置子公司新疆中通客车有限公司，产生较大投资收益，导致比较基数过大影响所致。
                    进一步数据年报分析，是营业成本中的原材料下降导致的毛利率增加，及主营业务成本率下降明显。行业平均原材料价格波动比率需要系统进一步开发来进行确认。整体看行业的毛利率2023年与2022年、2021年是持平的，都是10%-11%之间，没有明显波动。中通客车的原材料占比连年下降是非常特殊的。
                </p>
                <img src="file:///android_res/drawable/picture_zt_2.png" alt="示例图片" />
                <p>
                    系统定位出：2023年中通客车原材料下降存疑，需要了解具体情况。
                    进一步进行分析：
                    2023年车型变化与2022年持平，
                </p>
                <img src="file:///android_res/drawable/picture_zt_3.png" alt="示例图片" />
                <p>
                    2023年新能源车占比远远低于2022年，是否是新能源车成本率比传统客车要高造成的？
                </p>
                <img src="file:///android_res/drawable/picture_zt_4.png" alt="示例图片" />
                <p>
                    查询其他新能源客车的上市公司年报，其他公司披露是不分是否新能源车，无法找到同类可比公司去验证。再进一步搜索行业数据及新能源车成本的专业文章，无法在公开资源里找到依据。因为2023年新能源车销量下降明显，可以关注2024年年报，类似比例下原材料成本是否保持在2023年的原材料成本率。
                    这点存疑，但目前无法进行验证。
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