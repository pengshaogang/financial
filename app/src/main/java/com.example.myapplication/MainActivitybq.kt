package com.example.myapplication

import android.os.Bundle
import android.webkit.WebView
import androidx.activity.enableEdgeToEdge
import androidx.appcompat.app.AppCompatActivity
import androidx.core.view.ViewCompat
import androidx.core.view.WindowInsetsCompat

class MainActivitybq : AppCompatActivity() {
    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        enableEdgeToEdge()
        setContentView(R.layout.activity_main_activitybq)
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
                    北汽蓝谷(600733)过去5年未ST过。
                    系统数据异常值显示-2020年毛利率奇低、收入未明显减少的情况下应收账款却逐年减少。
                </p>
                <img src="file:///android_res/drawable/picture_bq_1.png" alt="示例图片" />
                <img src="file:///android_res/drawable/picture_bq_2.png" alt="示例图片" />
                <p>
                    1）北汽蓝谷2020年毛利率奇低，找到原因。
                    2020年毛利奇低--找到部分原因，如下图。是2020年执行新收入准则，对当年收入、成本造成了一定影响。在不考虑该因素影响的情况下，计算出2020年主营业务毛利率为-23.04%。
                </p>   
                <img src="file:///android_res/drawable/picture_bq_3.png" alt="示例图片" />
                <img src="file:///android_res/drawable/picture_bq_4.png" alt="示例图片" />
                <p>
                    另外，2020年，受新冠肺炎疫情等因素影响，北汽蓝谷的产销量未达预期，尤其是占比较高的对公销量受疫情影响严重，导致收入和毛利大幅下降，对北汽蓝谷业绩影响金额约为30亿元。另外还因市场压力、补贴退坡及高端乏力等因素。2019年年报显示，北汽蓝谷净利润为9201.01万元，这其中计算了非经营性损益项目中的10.4亿元政府补助。也就是说，如果没有补贴政策，北汽蓝谷恐怕早就陷入亏损危机。随着市场逐渐成熟，补贴退坡成为必然，北汽蓝谷其2020年收到的政府补助对比2019年大幅下降，对公司业绩影响金额约为9亿元。2020年北汽蓝谷累计销售新能源汽车2.59万辆，同比暴跌82.79%，甚至还不及2019年销量（15.06万辆）的零头，产量也仅为1.32万辆，产销数据均大幅下滑，业绩惨淡。2020年销量下降是导致了主营业务毛利率为负数的主要原因。
                </p>
                <p>
                    2）近年在收入未明显减少的情况下应收账款逐年减少，关联方还款造成的-找到原因。
                    北汽蓝谷近五年收入及应收账款余额未同向变动
                </p>
                <img src="file:///android_res/drawable/picture_bq_5.png" alt="示例图片" />
                <p>
                    收入主要来源于商品车销售、材料销售收入，账期如果一致的话，为什么收入增加应收账款反倒减少？2021年报P123，61 营业收入明细项目：
                </p>
                <img src="file:///android_res/drawable/picture_bq_6.png" alt="示例图片" />
                <p>
                    行业的应收账款周转率并无明显改善，而北汽蓝谷的数据特别特殊，系统可以帮助精准定位问题点。深入分析找到原因：北汽蓝谷应收账款来自关联方的部分，在快速清偿收款，北汽关联企业应收账款从2019年的101亿，逐年缩减到2023年的6亿。
                </p>
                <img src="file:///android_res/drawable/picture_bq_7.png" alt="示例图片" />
                
            </body>
            </html>
        """
        webView.settings.setSupportZoom(true)
        webView.settings.builtInZoomControls = true
        webView.settings.displayZoomControls = false
        webView.loadDataWithBaseURL(null, text, "text/html", "utf-8", null)


    }
}