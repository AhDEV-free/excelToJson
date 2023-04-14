package com.example.exceltojson;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.serializer.SerializerFeature;
import com.example.exceltojson.util.ExcelToJSONUtil;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

@SpringBootTest
class ExcelToJsonApplicationTests {


    /**
     * 数据转换
     * */
    @Test
    void contextLoads() {
        try {
            InputStream resourceAsStream = null;
             resourceAsStream = this.getClass().getClassLoader().getResourceAsStream("excelToJSONTest.xlsx");
//            resourceAsStream = new FileInputStream("/Users/yzh/home/skkj/dev/唐客中台/05数据采集/数据协议/承载网/cznet_table_data_模板.xlsx");
            // 本地文件
            // resourceAsStream = new FileInputStream("");
            JSONArray jsonArray = ExcelToJSONUtil.exportData(resourceAsStream);
            System.out.println(JSON.toJSONString(jsonArray,SerializerFeature.DisableCircularReferenceDetect));
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

}
