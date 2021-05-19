package com.cas.simple;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.EasyExcelFactory;
import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.read.metadata.ReadSheet;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.cas.config.DemoDataListener;
import com.cas.po.DemoData;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.test.context.junit4.SpringRunner;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * @author: xianglong[1391086179@qq.com]
 * @date: 上午10:59 2021/5/18
 * @version: V1.0
 * @review:
 */
@RunWith(SpringRunner.class)
public class SimpleTest {

    private static final String rootPath = "/Users/xianglong/IdeaProjects/cas-easyExcel/src/test/java/com/cas/simple";

    /**
     * 最简单的读
     */
    @Test
    public void simpleRead() {
        String fileName = rootPath + "/demo.xlsx";
        /**
         * 文件名｜第几页｜头类型｜转换实体｜
         */
        List<DemoData> list = EasyExcelFactory.read(fileName).sheet(0).headRowNumber(1).head(DemoData.class).doReadSync();
        list.forEach(a -> {
            System.out.println(a.toString());
        });
    }

    /**
     * 调用监听器
     */
    @Test
    public void listener() {
        String fileName = rootPath + "/demo.xlsx";
        ExcelReader excelReader = null;
        try {
            excelReader = EasyExcelFactory.read(fileName, DemoData.class, new DemoDataListener()).build();
            ReadSheet readSheet = EasyExcel.readSheet(0).build();
            excelReader.read(readSheet);
        } finally {
            if (excelReader != null) {
                // 这里千万别忘记关闭，读的时候会创建临时文件，到时磁盘会崩的
                excelReader.finish();
            }
        }
    }

    @Test
    public void simpleWrite() {
        // 写法1
        String fileName = rootPath + "simpleWrite" + System.currentTimeMillis() + ".xlsx";
        // 这里 需要指定写用哪个class去写，然后写到第一个sheet，名字为模板 然后文件流会自动关闭
        // 如果这里想使用03 则 传入excelType参数即可
//        EasyExcel.write(fileName, DemoData.class).sheet("模板").doWrite(data());

        // 写法2
        fileName = rootPath + "simpleWrite" + System.currentTimeMillis() + ".xlsx";
        // 这里 需要指定写用哪个class去写
        ExcelWriter excelWriter = null;
        try {
            excelWriter = EasyExcel.write(fileName, DemoData.class).build();
            WriteSheet writeSheet = EasyExcel.writerSheet("模板").build();
            excelWriter.write(data(), writeSheet);
        } finally {
            // 千万别忘记finish 会帮忙关闭流
            if (excelWriter != null) {
                excelWriter.finish();
            }
        }
    }

    private List<DemoData> data() {
        List<DemoData> list = new ArrayList<>();
        list.add(new DemoData("xl", new Date(), 12.0));
        list.add(new DemoData("xl2", new Date(), 12.0));
        list.add(new DemoData("xl3", new Date(), 12.0));
        return list;
    }





}
