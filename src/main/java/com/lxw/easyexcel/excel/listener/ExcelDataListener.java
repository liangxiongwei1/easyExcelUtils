package com.lxw.easyexcel.excel.listener;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.fastjson.JSON;
import lombok.extern.slf4j.Slf4j;

import java.util.ArrayList;
import java.util.List;
import java.util.function.Consumer;


/**
 * @author liang.xiongwei
 * @Title: ExcelDataListener
 * @Package com.intellif.vesionbook.vesionbookstatement.excel.listener
 * @Description
 * @date 2021/5/22 10:59
 */
@Slf4j
public class ExcelDataListener<T> extends AnalysisEventListener<T> {

    /**每次读取100条数据就进行保存操作*/
    private Integer batchCount;

    List<T> list = new ArrayList<>();

    Consumer<List<T>> saveWork;

    public ExcelDataListener(Integer batchCount,Consumer<List<T>> saveWork){
        this.batchCount = batchCount;
        this.saveWork = saveWork;
    }


    @Override
    public void invoke(T data, AnalysisContext analysisContext) {
        log.info("解析到一条数据:{}", JSON.toJSONString(data));
        list.add(data);
        if (list.size() >= batchCount) {
            //存储数据
            saveData();
            // 存储完成清理 list
            list.clear();
        }
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext analysisContext) {
        // 这里也要保存数据，确保最后遗留的数据也存储到数据库
        saveData();
        log.info("所有数据解析完成！");
    }

    private void saveData() {
        log.info("{}条数据，开始存储数据库！", list.size());
        //保存数据
        this.saveWork.accept(list);
        log.info("存储数据库成功！");
    }

}
