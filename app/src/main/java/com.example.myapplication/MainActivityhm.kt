package com.example.myapplication

import android.os.Bundle
import android.webkit.WebView
import androidx.activity.enableEdgeToEdge
import androidx.appcompat.app.AppCompatActivity
import androidx.core.view.ViewCompat
import androidx.core.view.WindowInsetsCompat

class MainActivityhm : AppCompatActivity() {
    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        enableEdgeToEdge()
        setContentView(R.layout.activity_main_activityhm)
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
                    海马汽车 (000572) 2016年，海马年销量21.6万辆，此后销量逐年下滑，2020年，年销量仅为17773辆。另外海马汽车2017年度、2018年度连续两个会计年度经审计的净利润为负值，股票于2019年4月24日开市起被深圳证券交易所施行退市风险警示处理，海马汽车股票名称由"海马汽车"变更为"*ST海马"。2019年，海马汽车出售344套房产，此外，作为小鹏汽车的代工厂，海马也收入颇丰，其2019年净利润以0.85亿元的成绩顺利转正，代工成为除卖房外又一重要营收来源。2020年，海马再次出售海口市住宅楼145套房产。2020年6月19日，海马汽车实现了摘星，从"*ST海马"变为"ST海马"。2021年，海马汽车控股子公司向中国铁路投资出售了其持有的海南银行7%股权，转让价格为3.297亿元。终于，一系列操作之下，"ST海马"彻底摘帽，赢下了第二阶段，成功归来。
                </p>
                <p>
                    系统数据异常值显示： 系统数据显示海马汽车2020年、2021年毛利率高出其他年度一倍。
                </p>
                <img src="file:///android_res/drawable/picture_hm_1.png" alt="示例图片" />
                <p>
                    进一步查询年报信息，可以看到三年收入主要来源于汽车制造收入：
                </p>
                <img src="file:///android_res/drawable/picture_hm_2.png" alt="示例图片" />
                <p>
                    2019-2021年海马汽车产销量变化，车型MPV、SUV、交叉型是大型车，轿车是小型车量，2020年大型车占比没有明显变化。
                </p>
                <img src="file:///android_res/drawable/picture_hm_3.png" alt="示例图片" />
                <p>
                    而2020年、2021年毛利率增加是源于汽车制造原材料成本下降导致的，2020年原材料与收入比2019年下降了17.38%， 2022、2023年比率为85.4%、82.06%，就2020年、2021年原材料与收入比降到了71-72%。
                </p>
                <img src="file:///android_res/drawable/picture_hm_4.png" alt="示例图片" />
                <p>
                    通过对营业成本进行分详分析，也可以看出2020年、2021年是原材料在营业成本中占比大幅下降导致的毛利率的升高。
                </p>
                <img src="file:///android_res/drawable/picture_hm_5.png" alt="示例图片" />
                <p>
                    根据行业平均数据分析，2020年、2021年行业平均毛利率并没有大幅升高，海马汽车是特殊的。另外去查询钢铁网2020年钢铁价格是下降的，但2021年钢铁价格又上涨上来。海马汽车2020、2021年原材料价格的大幅下降可以确定是不符合常理的，需要企业进一步说明。
                </p>
                <img src="file:///android_res/drawable/picture_hm_6.png" alt="示例图片" />
                
           
            </body>
            </html>
        """
        webView.settings.setSupportZoom(true)
        webView.settings.builtInZoomControls = true
        webView.settings.displayZoomControls = false
        webView.loadDataWithBaseURL(null, text, "text/html", "utf-8", null)

    }
}