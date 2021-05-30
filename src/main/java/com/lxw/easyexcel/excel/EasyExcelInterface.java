package com.lxw.easyexcel.excel;

/**
 * @author liang.xiongwei
 * @Title: EasyExcelInterface
 * @Package com.intellif.vesionbook.vesionbookstatement.excel.template
 * @Description 创建excel源数据
 * @date 2021/5/22 21:17
 */
public interface EasyExcelInterface<T> {

    /**
     * 创建ExcelSource
     * @param t
     * @return sheetName
     */
    ExcelSource createExcelSource(T t);
}
