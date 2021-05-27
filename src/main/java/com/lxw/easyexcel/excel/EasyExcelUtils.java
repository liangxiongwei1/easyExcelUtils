package com.lxw.easyexcel.excel;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.read.metadata.ReadSheet;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.metadata.fill.FillConfig;
import com.alibaba.excel.write.metadata.fill.FillWrapper;
import com.lxw.easyexcel.excel.listener.ExcelDataListener;
import com.lxw.easyexcel.excel.template.write.TemplateExporter;
import com.lxw.easyexcel.utils.FileUtils;
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
import java.net.URL;
import java.util.ArrayList;
import java.util.List;
import java.util.UUID;
import java.util.function.Consumer;

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
     * @param file
     * @param work
     * @param myClass
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
        } catch (IOException e) {
            log.error("按模板导出多个sheet excel异常",e);
            return null;
        }finally {
            if(temporaryFile!=null){
                FileUtils.delFile(temporaryFile);
            }
        }
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

    public static MultipartFile getResourcesFile(String filePath,String type){
        ClassLoader classLoader = FileUtils.class.getClassLoader();
        URL resource = classLoader.getResource(filePath);
        String path = resource.getPath();
        return new MockMultipartFile("file","file"+type,type,FileUtils.fileConvertToByteArray(path));
    }

    public static void main(String[] args) throws Exception{
        MultipleSheet multipleSheet1 = new MultipleSheet();
        List<TemplateExporter> list1 = new ArrayList<>();
        TemplateExporter  templateExporter = new TemplateExporter();
        templateExporter.setId(1);
        templateExporter.setNameT("one");
        TemplateExporter  templateExporter1 = new TemplateExporter();
        templateExporter1.setId(2);
        templateExporter1.setNameT("two");
        list1.add(templateExporter);
        list1.add(templateExporter1);
        multipleSheet1.setSheetName("2021-5-20");
        List<MultipleTable> listtemp1 = new ArrayList<>();
        MultipleTable multipleTable = new MultipleTable();
        multipleTable.setData(list1);
        multipleTable.setTableName("data0");
        MultipleTable multipleTable2 = new MultipleTable();
        multipleTable2.setData(list1);
        multipleTable2.setTableName("data1");
        listtemp1.add(multipleTable);
        listtemp1.add(multipleTable2);
        multipleSheet1.setTableDataList(listtemp1);

        List<MultipleSheet> multipleSheets = new ArrayList<>();
        multipleSheets.add(multipleSheet1);
//
        writeSingleSheetExcel(list1,"2021-5-20",TemplateExporter.class);
//        writeMultiSheetExcel(multipleSheets,TemplateExporter.class);
//        System.out.println(System.getProperty("user.dir"));
//        MultipartFile multipartFile = FileUtils.getResourcesFile("exceltempalte/0a8ada19-26d9-4aba-91dd-1e04be2fde61.xls","xls");
//        writeTemplateExcel(multipleSheets,multipartFile,"test");
    }
}
