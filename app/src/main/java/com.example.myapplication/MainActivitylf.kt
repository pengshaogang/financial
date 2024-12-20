package com.example.myapplication

import android.os.Bundle
import android.webkit.WebView
import androidx.activity.enableEdgeToEdge
import androidx.appcompat.app.AppCompatActivity
import androidx.core.view.ViewCompat
import androidx.core.view.WindowInsetsCompat

class MainActivitylf : AppCompatActivity() {
    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        enableEdgeToEdge()
        setContentView(R.layout.activity_main_activitylf)
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
                   力帆 (601777) 财报显示，其已经连续多年亏损。2016-2018年，连续三年扣非后净利润为负，2018年亏损21亿，2019年亏损47亿元，2020年上半年亏损26亿。2020年6月29日，债权人以力帆股份不能清偿到期债务，明显缺乏清偿能力，但仍具有重整价值为由，向重庆市第五中级人民法院申请重整；该院于2020年8月21日裁定受理力帆股份重整，并于同日指定清算组担任管理人。根据《上海证券交易所股票上市规则》的相关规定，力帆股份股票于2020年8月25日被实施退市风险警示，股票简称改为“*ST力帆”。吉利进入成为重整投资人，2021年2月8日，力帆重整计划已执行完毕，并终结重整程序。鉴于公司重整计划执行完毕涉及退市风险警示的情形已经消除，2021年3月2日，上交所同意撤销退市风险警示。2021年4月10日，力帆公司披露了2020年报实现营业收入36.37亿元，归属于母公司股东的净利润0.58亿元。最终被认定，"公司涉及其他风险警示的情形已经消除，符合申请撤销股票其他风险警示的条件。"上交所于2021年4月22日同意申请并撤销其他风险警示。
                   系统数据异常值显示-力帆科技2020财务费用率变化异常；2019、2020、2021年三年毛利率为-3.02%、9.84%、16.07%，增长过于迅速，远远超越行业平均数：
                </p>
                <img src="file:///android_res/drawable/picture_lf_1.png" alt="示例图片" />
                <p>
                    1）2020年财务费用率51.91%-找到原因。
                </p>
                <img src="file:///android_res/drawable/picture_lf_2.png" alt="示例图片" />
                <p>
                    2）力帆科技2019、2020、2021年三年毛利率为-3.02%、9.84%、16.07%，增长过于迅速，远远超越行业平均数。重组后力帆科技产业有所调整，从2019年至2021年年报“2. 收入和成本分析”显示， 2019年当年该产品毛利为-17.27%，且随后两年该产品毛利逐年增加。
                    2020年、2021年乘用车及配件毛利率增长到15.78%。
                </p>
                <img src="file:///android_res/drawable/picture_lf_3.png" alt="示例图片" />
                <p>
                    再进一步细化，主要是原材料成本下降造成的，可以问询企业：为什么乘用车原材料成本2021年降低17.39%，并且三种产品的原材料成本变化不同的原因，排除钢铁原材料成本下降原因。
                </p>
                <img src="file:///android_res/drawable/picture_lf_4.png" alt="示例图片" />
                <p>
                    我们把乘用车单独列示，可以明显看到2021年毛利率在吉利入主当年大幅增加，可是在以后年度的毛利率仍然是很低的。
                </p>
                <img src="file:///android_res/drawable/picture_lf_5.png" alt="示例图片" />
                <p>
                    按照车型进行分析，2021年新能源车占比达86%，以后年度新能源车占比逐年下滑，新能源车直接材料所用更少，从2021年、2022年、2023年数据上看，新能源车比率每年比上一年占比正好下降14%，而毛利率下降10.56%，7.41%，毛利率下降比率与新能源车型下降比率不成正比。所以新能源车型原材料成本更低不能完全解释存疑的点。
                </p>
                <img src="file:///android_res/drawable/picture_lf_6.png" alt="示例图片" />
                <p>
                    另外，进一步分析乘用车的营业成本，可以看出2021年乘用车的原材料、直接人工、制造费用比率都是前无古人、后无来者的低点，是非常明显的存疑点：
                </p>
                <img src="file:///android_res/drawable/picture_lf_7.png" alt="示例图片" />
           
            </body>
            </html>
        """
        webView.settings.setSupportZoom(true)
        webView.settings.builtInZoomControls = true
        webView.settings.displayZoomControls = false
        webView.loadDataWithBaseURL(null, text, "text/html", "utf-8", null)
    }
}