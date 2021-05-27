package com.lxw.easyexcel.excel.template.read;

import com.alibaba.excel.annotation.ExcelProperty;
import lombok.Data;

/**
 * @author liang.xiongwei
 * @Title: TemplateReader
 * @Package com.intellif.vesionbook.vesionbookstatement.excel.template.read
 * @Description
 * @date 2021/5/22 11:23
 */
@Data
public class TemplateReader {

    @ExcelProperty(index = 0)
    private int id;

    @ExcelProperty(index = 1)
    private String name;
}
