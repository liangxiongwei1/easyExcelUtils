package com.lxw.easyexcel.excel;

import lombok.Data;

import java.util.List;
import java.util.Map;

/**
 * @author liang.xiongwei
 * @Title: ExcelSource
 * @Package com.intellif.vesionbook.vesionbookstatement.excel
 * @Description
 * @date 2021/5/23 15:48
 */
@Data
public class ExcelSource<R extends ExcelModelInterface> {

    /**模板路径*/
    private String templatePath;

    /**excel数据 one excel one key*/
    private Map<String,List<R>> excelData;
}
