package com.cas.util;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.cas.po.ExcelBase;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.net.URLEncoder;
import java.util.List;

/**
 * @author: xianglong[1391086179@qq.com]
 * @date: 下午7:46 2021/5/18
 * @version: V1.0
 * @review:
 */
public class ExcelUtil {

    public static <T extends ExcelBase> void write(HttpServletResponse response, String excelName, List<T> list, Class<T> clz) {
        ServletOutputStream out = null;
        ExcelWriter excelWriter = null;
        try {
            out = response.getOutputStream();
            excelWriter = EasyExcel.write(out, clz).build();
            WriteSheet writeSheet = EasyExcel.writerSheet("sheet1").build();
            excelWriter.write(list, writeSheet);
            response.setHeader("Content-disposition", "attachment;filename=" + URLEncoder.encode(excelName, "UTF-8") + ".xlsx");
            response.setContentType("application/x-msdownload");
            response.setCharacterEncoding("utf-8");
            out.flush();
        } catch (Exception e) {
        } finally {
            if (excelWriter != null) {
                excelWriter.finish();
            }
            if (out != null) {
                try {
                    out.close();
                } catch (Exception e) {

                }
            }
        }
    }

}
