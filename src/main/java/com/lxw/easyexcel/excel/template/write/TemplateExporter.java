package com.lxw.easyexcel.excel.template.write;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.ColumnWidth;
import lombok.Data;

import java.util.List;


@Data
public class TemplateExporter{
    @ColumnWidth(20)
    @ExcelProperty({"ID"})
    private int id;
    @ExcelProperty({"名称"})
    @ColumnWidth(20)
    private String nameT;

    private List<TemplateExporter> tem;
}
