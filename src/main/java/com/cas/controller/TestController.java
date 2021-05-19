package com.cas.controller;

import com.cas.po.DemoData;
import com.cas.util.ExcelUtil;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;

import javax.servlet.http.HttpServletResponse;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * @author: xianglong[1391086179@qq.com]
 * @date: 下午6:07 2021/5/18
 * @version: V1.0
 * @review:
 */
@Controller
public class TestController {

    @GetMapping("/export")
    public void test(HttpServletResponse response) {
        ExcelUtil.write(response, "向龙", data(), DemoData.class);
    }

    private List<DemoData> data() {
        List<DemoData> list = new ArrayList<>();
        list.add(new DemoData("xl", new Date(), 12.0));
        list.add(new DemoData("xl2", new Date(), 12.0));
        list.add(new DemoData("xl3", new Date(), 12.0));
        return list;
    }


}
