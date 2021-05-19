package com.cas.po;

import com.alibaba.excel.annotation.ExcelProperty;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.Date;

@Data
public class DemoData extends ExcelBase{
    @ExcelProperty(value = "姓名", index = 0)
    private String string;
    @ExcelProperty(value = "时间", index = 1)
    private Date date;
    @ExcelProperty(value = "价格", index = 2)
    private Double doubleData;

    public DemoData() {}

    public DemoData(String string, Date date, Double doubleData) {
        this.string = string;
        this.date = date;
        this.doubleData = doubleData;
    }

    public String getString() {
        return string;
    }

    public void setString(String string) {
        this.string = string;
    }

    public Date getDate() {
        return date;
    }

    public void setDate(Date date) {
        this.date = date;
    }

    public Double getDoubleData() {
        return doubleData;
    }

    public void setDoubleData(Double doubleData) {
        this.doubleData = doubleData;
    }

    @Override
    public String toString() {
        return "DemoData{" +
                "name='" + string + '\'' +
                ", date=" + date +
                ", doubleData=" + doubleData +
                '}';
    }
}
