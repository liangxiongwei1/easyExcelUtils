package com.lxw.easyexcel.excel;

import lombok.Data;

import java.util.List;
import java.util.Map;
import java.util.function.Function;

/**
 * @author liang.xiongwei
 * @Title: ExcelSource
 * @Package com.intellif.vesionbook.vesionbookstatement.excel
 * @Description
 * @date 2021/5/23 15:48
 */
@Data
public class ExcelSource<T,R extends ExcelModelInterface> {

    private String templatePath;

    private Function<T,R> sheetFunction;

    private Map<String,List<T>> excelParam;
}
