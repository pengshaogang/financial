package com.example.myapplication

import android.os.Bundle
import androidx.activity.enableEdgeToEdge
import androidx.appcompat.app.AppCompatActivity
import androidx.core.view.ViewCompat
import androidx.core.view.WindowInsetsCompat
import android.widget.Toast
import android.content.Intent
import android.app.AlertDialog

class MainActivity2 : AppCompatActivity() {
    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        enableEdgeToEdge()
        setContentView(R.layout.activity_main2)
//        ViewCompat.setOnApplyWindowInsetsListener(findViewById(R.id.main)) { v, insets ->
//            val systemBars = insets.getInsets(WindowInsetsCompat.Type.systemBars())
//            v.setPadding(systemBars.left, systemBars.top, systemBars.right, systemBars.bottom)
//            insets
//        }
    }

    fun onButtonClick(view: android.view.View) {
        Toast.makeText(this, "未发现异常！", Toast.LENGTH_SHORT).show()
    }

    fun onButtonClick1(view: android.view.View) {
        Toast.makeText(this, "2024年8月9日公告退市！", Toast.LENGTH_SHORT).show()
    }

    fun onButtonClick2(view: android.view.View) {
        Toast.makeText(this, "出具保留意见! ", Toast.LENGTH_SHORT).show()
    }

    fun onButtonClick3(view: android.view.View) {
        Toast.makeText(this, "目前在重整！ ", Toast.LENGTH_SHORT).show()
    }

    fun onButtonClick_errak(view: android.view.View) {
        val builder = AlertDialog.Builder(this)
        // 设置对话框的标题和消息
//        builder.setTitle("异常通知")
        builder.setMessage("异常，点击进一步了解原因")

        // 设置对话框的“确定”按钮
        builder.setPositiveButton("进一步了解") { dialog, which ->
            // 创建 Intent 来启动 MainActivityak
            val intent = Intent(this, MainActivityak::class.java)
            // 启动 MainActivityak
            startActivity(intent)
            dialog.dismiss()
        }

        // 设置对话框的“取消”按钮
        builder.setNegativeButton("取消") { dialog, which ->
            // 用户取消了操作，对话框被关闭
            dialog.dismiss()
        }

        // 创建并显示对话框
        builder.create().show()
    }

    fun onButtonClick_errbq(view: android.view.View) {
        val builder = AlertDialog.Builder(this)
        // 设置对话框的标题和消息
//        builder.setTitle("异常通知")
        builder.setMessage("异常，点击进一步了解原因")

        // 设置对话框的“确定”按钮
        builder.setPositiveButton("进一步了解") { dialog, which ->
            // 创建 Intent 来启动 MainActivityak
            val intent = Intent(this, MainActivitybq::class.java)
            // 启动 MainActivityak
            startActivity(intent)
            dialog.dismiss()
        }

        // 设置对话框的“取消”按钮
        builder.setNegativeButton("取消") { dialog, which ->
            // 用户取消了操作，对话框被关闭
            dialog.dismiss()
        }

        // 创建并显示对话框
        builder.create().show()
    }

    fun onButtonClick_errhm(view: android.view.View) {
        val builder = AlertDialog.Builder(this)
        // 设置对话框的标题和消息
//        builder.setTitle("异常通知")
        builder.setMessage("异常，点击进一步了解原因")

        // 设置对话框的“确定”按钮
        builder.setPositiveButton("进一步了解") { dialog, which ->
            // 创建 Intent 来启动 MainActivityak
            val intent = Intent(this, MainActivityhm::class.java)
            // 启动 MainActivityak
            startActivity(intent)
            dialog.dismiss()
        }

        // 设置对话框的“取消”按钮
        builder.setNegativeButton("取消") { dialog, which ->
            // 用户取消了操作，对话框被关闭
            dialog.dismiss()
        }

        // 创建并显示对话框
        builder.create().show()
    }

    fun onButtonClick_errlf(view: android.view.View) {
        val builder = AlertDialog.Builder(this)
        // 设置对话框的标题和消息
//        builder.setTitle("异常通知")
        builder.setMessage("异常，点击进一步了解原因")

        // 设置对话框的“确定”按钮
        builder.setPositiveButton("进一步了解") { dialog, which ->
            // 创建 Intent 来启动 MainActivityak
            val intent = Intent(this, MainActivitylf::class.java)
            // 启动 MainActivityak
            startActivity(intent)
            dialog.dismiss()
        }

        // 设置对话框的“取消”按钮
        builder.setNegativeButton("取消") { dialog, which ->
            // 用户取消了操作，对话框被关闭
            dialog.dismiss()
        }

        // 创建并显示对话框
        builder.create().show()
    }

    fun onButtonClick_erryq(view: android.view.View) {
        val builder = AlertDialog.Builder(this)
        // 设置对话框的标题和消息
//        builder.setTitle("异常通知")
        builder.setMessage("异常，点击进一步了解原因")

        // 设置对话框的“确定”按钮
        builder.setPositiveButton("进一步了解") { dialog, which ->
            // 创建 Intent 来启动 MainActivityak
            val intent = Intent(this, MainActivityyq::class.java)
            // 启动 MainActivityak
            startActivity(intent)
            dialog.dismiss()
        }

        // 设置对话框的“取消”按钮
        builder.setNegativeButton("取消") { dialog, which ->
            // 用户取消了操作，对话框被关闭
            dialog.dismiss()
        }

        // 创建并显示对话框
        builder.create().show()
    }

    fun onButtonClick_errzt(view: android.view.View) {
        val builder = AlertDialog.Builder(this)
        // 设置对话框的标题和消息
//        builder.setTitle("异常通知")
        builder.setMessage("异常，点击进一步了解原因")

        // 设置对话框的“确定”按钮
        builder.setPositiveButton("进一步了解") { dialog, which ->
            // 创建 Intent 来启动 MainActivityak
            val intent = Intent(this, MainActivityzt::class.java)
            // 启动 MainActivityak
            startActivity(intent)
            dialog.dismiss()
        }

        // 设置对话框的“取消”按钮
        builder.setNegativeButton("取消") { dialog, which ->
            // 用户取消了操作，对话框被关闭
            dialog.dismiss()
        }

        // 创建并显示对话框
        builder.create().show()
    }

    fun onButtonClick_errztai(view: android.view.View) {
        val builder = AlertDialog.Builder(this)
        // 设置对话框的标题和消息
//        builder.setTitle("异常通知")
        builder.setMessage("异常，点击进一步了解原因")

        // 设置对话框的“确定”按钮
        builder.setPositiveButton("进一步了解") { dialog, which ->
            // 创建 Intent 来启动 MainActivityak
            val intent = Intent(this, MainActivityztai::class.java)
            // 启动 MainActivityak
            startActivity(intent)
            dialog.dismiss()
        }

        // 设置对话框的“取消”按钮
        builder.setNegativeButton("取消") { dialog, which ->
            // 用户取消了操作，对话框被关闭
            dialog.dismiss()
        }

        // 创建并显示对话框
        builder.create().show()
    }

}