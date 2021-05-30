package com.lxw.easyexcel.excel;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.read.metadata.ReadSheet;
import com.alibaba.excel.util.StringUtils;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.metadata.fill.FillConfig;
import com.alibaba.excel.write.metadata.fill.FillWrapper;
import com.lxw.easyexcel.excel.listener.ExcelDataListener;
import com.lxw.easyexcel.excel.template.write.TemplateExporter;
import com.lxw.easyexcel.utils.FileUtils;
import com.lxw.easyexcel.utils.ZipUtil;
import lombok.Cleanup;
import lombok.Data;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.mock.web.MockMultipartFile;
import org.springframework.util.CollectionUtils;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.*;
import java.util.function.Consumer;
import java.util.stream.Collectors;

/**
 * @author liang.xiongwei
 * @Title: EasyExcelUtils
 * @Package com.intellif.vesionbook.vesionbookstatement.excel
 * @Description
 * @date 2021/5/22 11:12
 */
@Slf4j
public class EasyExcelUtils{

    public static final String EXCEL_XLS = "xls";
    public static final String EXCEL_XLSX = "xlsx";

    /**
     * 读单个sheet excel
     * @param file excel文件
     * @param work 读取后工作方法
     * @param myClass 模型类
     * @param <T>
     * @throws Exception
     */
    public static <T> void readSingleSheetExcel(MultipartFile file, Consumer<List<T>> work, Class myClass) throws Exception{
        readSingleSheetExcel(file,0,100,work,myClass);
    }

    public static <T> void readSingleSheetExcel(MultipartFile file,Integer sheetNum, Integer batchCount, Consumer<List<T>> work,Class myClass) throws Exception{
        @Cleanup InputStream fis = file.getInputStream();
        ExcelDataListener listener = new ExcelDataListener(batchCount,work);
        ExcelReader excelReader = EasyExcel.read(fis, myClass, listener).build();
        ReadSheet readSheet = EasyExcel.readSheet(sheetNum).build();
        excelReader.read(readSheet);
        // 这里千万别忘记关闭，读的时候会创建临时文件，到时磁盘会崩的
        excelReader.finish();
    }

    /**
     * 导出单sheet excel
     * @param data
     * @param sheetName
     * @param myClass
     * @return
     */
    public static MultipartFile writeSingleSheetExcel(List<?> data,String sheetName,Class myClass){
        String temporaryFile = null;
        try {
            temporaryFile = createNewExcelPath();
            ExcelWriter excelWriter = EasyExcel.write(temporaryFile, myClass).build();
            excelWriter.write(data,EasyExcel.writerSheet(sheetName).build());
            excelWriter.finish();
            byte[] fileData = FileUtils.fileConvertToByteArray(temporaryFile);
            return new MockMultipartFile("excel","excel.xls","xls",fileData);
        } catch (Exception e) {
            log.error("导出单sheet excel异常",e);
            return null;
        }finally {
            if(temporaryFile!=null){
                FileUtils.delFile(temporaryFile);
            }
        }
    }

    /**
     * 导出多个sheet excel到本地
     * @param data
     * @param myClass
     * @return
     */
    public static MultipartFile writeMultiSheetExcel(List<MultipleSheet> data,Class myClass){
        String temporaryFile = null;
        try {
            temporaryFile = createNewExcelPath();
            ExcelWriter excelWriter = EasyExcel.write(temporaryFile, myClass).build();
            for(int i = 0 ; i < data.size() ; i++){
                MultipleSheet multipleSheet = data.get(i);
                excelWriter.write(multipleSheet.getData(),EasyExcel.writerSheet(i,multipleSheet.getSheetName()).build());
            }
            excelWriter.finish();
            byte[] fileData = FileUtils.fileConvertToByteArray(temporaryFile);
            return new MockMultipartFile("excel","excel.xls","xls",fileData);
        } catch (Exception e) {
            log.error("导出多个sheet excel异常",e);
            return null;
        }finally {
            if(temporaryFile!=null){
                FileUtils.delFile(temporaryFile);
            }
        }
    }

    /**
     * 按模板导出多个sheet excel
     * @param data
     * @param multipartFile
     * @return
     * @throws Exception
     */
    public static MultipartFile writeTemplateExcel(List<MultipleSheet> data,MultipartFile multipartFile,String excelName){
        String temporaryFile = null;
        try {
            Workbook templateWorkbook = getWorkbook(multipartFile);
            for (int i = 0; i < data.size()-1; i++) {
                templateWorkbook.cloneSheet(0);
            }
            for (int i = 0; i < data.size(); i++) {
                templateWorkbook.setSheetName(i, data.get(i).getSheetName());
            }
            @Cleanup ByteArrayOutputStream outStream = new ByteArrayOutputStream();
            templateWorkbook.write(outStream);
            @Cleanup ByteArrayInputStream templateInputStream = new ByteArrayInputStream(outStream.toByteArray());
            temporaryFile = createNewExcelPath();
            ExcelWriter excelWriter = EasyExcel.write(temporaryFile).withTemplate(templateInputStream).build();
            FillConfig fillConfig = FillConfig.builder().forceNewRow(Boolean.TRUE).build();
            for (int i = 0; i < data.size(); i++) {
                WriteSheet writeSheet = EasyExcel.writerSheet(i).build();
                List<MultipleTable> tableList = data.get(i).getTableDataList();
                if(!CollectionUtils.isEmpty(tableList)){
                    for(int j = 0;j<tableList.size();j++){
                        MultipleTable multipleTable = tableList.get(j);
                        excelWriter.fill(new FillWrapper(multipleTable.getTableName(), multipleTable.getData()), fillConfig,writeSheet);
                    }
                }
            }
            excelWriter.finish();
            byte[] fileData = FileUtils.fileConvertToByteArray(temporaryFile);
            return new MockMultipartFile(excelName,excelName+".xls","xls",fileData);
        } catch (Exception e) {
            log.error("按模板导出多个sheet excel异常",e);
            return null;
        }finally {
            if(temporaryFile!=null){
                FileUtils.delFile(temporaryFile);
            }
        }
    }


    /**
     * 导出多文件 多sheet excel
     * @param excelSource Excel需要导出源数据
     * @return
     * @throws Exception
     */
    public static MultipartFile statementExport(ExcelSource excelSource) throws Exception{
        if(excelSource == null || StringUtils.isEmpty(excelSource.getTemplatePath())){
            log.warn("获取excel 导出源异常");
            return null;
        }
        return exportToZIP(excelSource);
    }


    /**
     * 导出excel 并打包
     * @param excelSource
     * @return
     */
    private static <R extends ExcelModelInterface> MultipartFile exportToZIP(ExcelSource<R> excelSource) throws Exception{
        Map<String,List<MultipleSheet>> map = new HashMap<>(8);
        excelSource.getExcelData().forEach((k,v)->{
            List<EasyExcelUtils.MultipleSheet> multipleSheets = v.stream().map(p->p.buildSheet()).collect(Collectors.toList());
            map.put(k,multipleSheets);
        });

        MultipartFile template = FileUtils.getResourcesFile(excelSource.getTemplatePath(),EasyExcelUtils.EXCEL_XLS);
        List<MultipartFile> zipFiles = new ArrayList<>();
        map.forEach((k,v)->{
            MultipartFile multipartFile = writeTemplateExcel(v,template,k);
            if(multipartFile != null){
                zipFiles.add(multipartFile);
            }
        });

        if(!CollectionUtils.isEmpty(zipFiles)){
            if(zipFiles.size() == 1){
                return new MockMultipartFile("file","file.xls", EasyExcelUtils.EXCEL_XLS,zipFiles.get(0).getBytes());
            }else {
                return ZipUtil.toZip(zipFiles);
            }
        }
        return null;
    }

    /**
     * @Description 判断Excel的版本, 获取Workbook
     */
    public static Workbook getWorkbook(MultipartFile file) throws IOException {
        Workbook wb = null;
        if (file.getOriginalFilename().endsWith(EXCEL_XLS)) {
            //Excel 2003
            wb = new HSSFWorkbook(file.getInputStream());
        } else if (file.getOriginalFilename().endsWith(EXCEL_XLSX)) {
            // Excel 2007/2010
            wb = new XSSFWorkbook(file.getInputStream());
        }
        return wb;
    }

    public static String createNewExcelPath(){
        return UUID.randomUUID().toString() + ".xls";
    }

    @Data
    public static class MultipleSheet{

        private List<?> data;

        private List<MultipleTable> tableDataList;

        private String sheetName;
    }

    @Data
    public static class MultipleTable{

        private List<?> data;

        private String tableName;
    }
}
