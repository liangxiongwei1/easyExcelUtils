package com.lxw.easyexcel.utils;

import lombok.Cleanup;
import lombok.extern.slf4j.Slf4j;
import org.springframework.core.io.ClassPathResource;
import org.springframework.mock.web.MockMultipartFile;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;


@Slf4j
public class FileUtils {
    /**
     * 把一个文件转化为byte字节数组。
     */
    public static byte[] fileConvertToByteArray(String filePath) throws Exception {
        @Cleanup FileInputStream fis = new FileInputStream(filePath);
        return fileStreamConvertToByteArray(fis);
    }


    /**
     * 把一个文件转化为byte字节数组。
     */
    public static byte[] fileStreamConvertToByteArray(InputStream inputStream) throws Exception{
        @Cleanup ByteArrayOutputStream baos = new ByteArrayOutputStream();
        int len;
        byte[] buffer = new byte[1024];
        while ((len = inputStream.read(buffer)) != -1) {
            baos.write(buffer, 0, len);
        }
        return baos.toByteArray();
    }

    public static MultipartFile getResourcesFile(String filePath, String type) throws Exception{
        ClassPathResource classPathResource = new ClassPathResource(filePath);
        InputStream inputStream =classPathResource.getInputStream();
        return new MockMultipartFile("file","file"+type,type,FileUtils.fileStreamConvertToByteArray(inputStream));
    }


    /**
     * 删除文件
     * @param path 包含文件名
     */
    public static void delFile(String path){
        File file=new File(path);
        if(file.exists()&&file.isFile())
            file.delete();
    }


}
