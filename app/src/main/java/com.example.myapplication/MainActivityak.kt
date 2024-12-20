package com.example.myapplication

import android.os.Bundle
import androidx.activity.enableEdgeToEdge
import androidx.appcompat.app.AppCompatActivity
import androidx.core.view.ViewCompat
import androidx.core.view.WindowInsetsCompat
import android.webkit.WebView

class MainActivityak : AppCompatActivity() {
    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        enableEdgeToEdge()
        setContentView(R.layout.activity_main_activityak)
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
                    2019年4月，安凯客车(000868)因净利润连续两年亏损，被实行“退市风险警示”处理，股票简称由“安凯客车”变更为“*ST安凯”。2019年，安凯客车实现扭亏为盈，不过主要利润来源为依靠处置非流动资产。2021年4月，*ST安凯摘星脱帽。系统数据异常值显示：2023年资本保值增值率比2022年增长了近10倍，且销售费用率、管理费用率、研发费用率比2022年都呈减少趋势等异常变动。
                </p>
                <img src="file:///android_res/drawable/picture1.png" alt="示例图片" />
                <p>
                    对年报进一步分析：
                    1）2023年所有者权益大幅增加原因如下:
                </p>   
                <img src="file:///android_res/drawable/picture2.png" alt="示例图片" />
                <p>
                    2）2023年销售费用、管理费用、研发费用占收入的比例较2022年均下降
                    国内市场受旅游市场强势复苏等因素影响，公路客车市场销量大幅提升；国际市场对客车的需求逐步恢复，市场需求持续增长。2023 年，全年累计实现客车销量 4,328 台，同比增长 40.89%；实现营业收入 21.46 亿元，同比增长 44.25%。利润同比实现减亏。
                </p>
                <img src="file:///android_res/drawable/picture3.png" alt="示例图片" />
                <p>
                    2023年管理费用比2022年下降是不合理的，进一步拆解管理费用明细，折旧与摊销额及安全生产费比2022年明显下降。
                </p>
                <img src="file:///android_res/drawable/picture4.png" alt="示例图片" />
                <p>
                    2023年固定资产、无形资产数据变动不大。
                </p>
                <img src="file:///android_res/drawable/picture5.png" alt="示例图片" />
                <p>
                    进一步分析，安凯客车现金流量表中披露折旧费用并未明显减少：
                </p>
                <img src="file:///android_res/drawable/picture6.png" alt="示例图片" />
                <p>
                    营业收入从2022年14.87亿增长到21.46亿，但折旧没有增加还减少了，划分上是折旧更多划分给成本，成本折旧与营业收入同比例增加，所以管理费用折旧就明显减少了。
                </p>
                <img src="file:///android_res/drawable/picture7.png" alt="示例图片" />
                
                
            </body>
            </html>
        """
        webView.settings.setSupportZoom(true)
        webView.settings.builtInZoomControls = true
        webView.settings.displayZoomControls = false
        webView.loadDataWithBaseURL(null, text, "text/html", "utf-8", null)

    }


}

