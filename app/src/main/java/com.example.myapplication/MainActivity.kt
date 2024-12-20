package com.example.myapplication
import android.content.Intent

import android.os.Bundle
import android.widget.Toast
import androidx.activity.ComponentActivity
import androidx.activity.compose.setContent
import androidx.activity.enableEdgeToEdge
import androidx.compose.foundation.layout.fillMaxSize
import androidx.compose.foundation.layout.padding
import androidx.compose.material3.Scaffold
import androidx.compose.material3.Text
import androidx.compose.material3.Button
import androidx.compose.runtime.Composable
import androidx.compose.foundation.layout.Column
import androidx.compose.foundation.layout.Box
import androidx.compose.foundation.layout.Arrangement
import androidx.compose.foundation.layout.Spacer
import androidx.compose.foundation.layout.height
import androidx.compose.ui.Modifier
import androidx.compose.ui.Alignment
import androidx.compose.ui.unit.dp
import androidx.compose.ui.tooling.preview.Preview
import androidx.compose.ui.platform.LocalContext
import androidx.compose.ui.text.style.TextAlign
import com.example.myapplication.ui.theme.MyApplicationTheme


class MainActivity : ComponentActivity() {
    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        enableEdgeToEdge()
        setContent {
            MyApplicationTheme {
                Scaffold(modifier = Modifier.fillMaxSize()) { innerPadding ->
                    Greeting(
                        name = "Android",
                        modifier = Modifier.padding(innerPadding)
                    )
                }
            }
        }
    }
}

@Composable
fun Greeting(name: String, modifier: Modifier = Modifier) {
    val context = LocalContext.current//----psg

    Box(
        modifier = Modifier
            .fillMaxSize() // Fill the available space
            .padding(16.dp) // Provide padding
    ) {
        Column(
            horizontalAlignment = Alignment.CenterHorizontally, // Center content horizontally
            verticalArrangement = Arrangement.Center, // Center content vertically
            modifier = Modifier.align(Alignment.Center) // Align the column to the center of the Box
        ) {
            Text(
                text = "整车制造行业上市公司分析系统-分析报告",
                textAlign = TextAlign.Center // Center the text
            )
            Spacer(modifier = Modifier.height(8.dp)) // Space between text and button
            Button(onClick = {
                context.startActivity(Intent(context, MainActivity2::class.java))
//                Toast.makeText(context, "欢迎, $name!", Toast.LENGTH_SHORT).show()
            }) {
                Text("查看更多信息")
            }
        }
    }

//    Column(modifier = modifier) {//------psg
//        Text(
////            text = "证券集团，开始开发, going, $name!",
//            text = "整车制造行业上市公司分析系统-分析报告",
//            modifier = modifier
//        )
//        Button(onClick = {
//            Toast.makeText(context, "欢迎, $name!", Toast.LENGTH_SHORT).show()
//        }) {
//            Text("查看更多信息")
//        }
//    }

//    Text(
//        text = "证券集团，开始开发, going, $name!",
//        modifier = modifier
//    )
}

@Preview(showBackground = true)
@Composable
fun GreetingPreview() {
    MyApplicationTheme {
        Greeting("Android")
    }
}