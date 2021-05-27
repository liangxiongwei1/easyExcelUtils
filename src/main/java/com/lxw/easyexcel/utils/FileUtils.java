package com.lxw.easyexcel.utils;


import com.alibaba.excel.util.StringUtils;
import lombok.Cleanup;
import lombok.extern.slf4j.Slf4j;
import org.apache.tools.zip.ZipEntry;
import org.apache.tools.zip.ZipFile;
import org.springframework.mock.web.MockMultipartFile;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.net.URL;
import java.util.Enumeration;

@Slf4j
public class FileUtils {

    /**
     * 解压
     * @param sourceFile 带解压的文件
     * @param destPath 解压后的路径
     */
    public static File unZip(String sourceFile, String destPath) throws Exception {
        ZipFile zipFile = new ZipFile(sourceFile,"gbk");
        Enumeration zList = zipFile.getEntries();  //得到zip包里的所有元素
        ZipEntry ze;
        File srcFile = null;
        byte[] buf = new byte[1024];
        while (zList.hasMoreElements()) {
            ze = (ZipEntry) zList.nextElement();
            if (ze.isDirectory()) {
                System.out.println(ze.getName());
                if("封面/".equalsIgnoreCase(ze.getName())) {
                    log.info("打开zip文件里的文件夹:" + ze.getName() + "skipped...");
                    continue;
                }else{
                    log.error("非法的模板格式");
                    throw new IllegalArgumentException("非法的模板格式");
                }
            }
            OutputStream outputStream = null;
            InputStream inputStream = null;
            try {
                //以ZipEntry为参数得到一个InputStream，并写到OutputStream中
                File realFile = getRealFileName(destPath, ze.getName());
                System.out.println(realFile.getAbsolutePath());
                if (ze.getName().lastIndexOf(".xlsx") != -1) {
                    srcFile = realFile;
                }
                outputStream = new BufferedOutputStream(new FileOutputStream(realFile));
                inputStream = new BufferedInputStream(zipFile.getInputStream(ze));
                int readLen;
                while ((readLen = inputStream.read(buf, 0, 1024)) != -1) {
                    outputStream.write(buf, 0, readLen);
                }
                inputStream.close();
                outputStream.close();
            } catch (Exception e) {
                log.error("解压失败", e);
                throw new IOException("解压失败：" + e.toString());
            } finally {
                if (inputStream != null) {
                    inputStream.close();
                }
                if (outputStream != null) {
                     outputStream.close();
                }
            }

        }
        zipFile.close();
        return srcFile;
    }

    /**
     * 给定根目录，返回一个相对路径所对应的实际文件名.
     *
     * @param path     指定根目录
     * @param absFileName 相对路径名，来自于ZipEntry中的name
     * @return java.io.File 实际的文件
     */
    private static File getRealFileName(String path, String absFileName) {
        String[] dirs = absFileName.split("/", absFileName.length());
        File ret = new File(path);// 创建文件对象
        if (dirs.length > 1) {
            for (int i = 0; i < dirs.length - 1; i++) {
                ret = new File(ret, dirs[i]);
            }
        }
        if (!ret.exists()) {         // 检测文件是否存在
            ret.mkdirs(); // 创建此抽象路径名指定的目录
        }
        ret = new File(ret, dirs[dirs.length - 1]);// 根据 ret 抽象路径名和 child 路径名字符串创建一个新 File 实例

        return ret;
    }

    public static void deleteFile(File file) {
        if (file.exists()) {
            if (file.isDirectory()) {
                File[] files = file.listFiles();
                for (int i = 0; i < files.length; i++) {
                    deleteFile(files[i]);
                }
            }
        }
        file.delete();
    }

    /**
     * 把一个文件转化为byte字节数组。
     */
    public static byte[] fileConvertToByteArray(String filePath){
        byte[] data = null;
        try {
            @Cleanup FileInputStream fis = new FileInputStream(filePath);
            @Cleanup ByteArrayOutputStream baos = new ByteArrayOutputStream();
            int len;
            byte[] buffer = new byte[1024];
            while ((len = fis.read(buffer)) != -1) {
                baos.write(buffer, 0, len);
            }
            data = baos.toByteArray();
        } catch (IOException e) {
            log.error("读取文件异常",e);
        }
        return data;
    }

    public static  void inputStreamToFile(InputStream ins,File file) {
        try {
            OutputStream os = new FileOutputStream(file);
            int bytesRead = 0;
            byte[] buffer = new byte[8192];
            while ((bytesRead = ins.read(buffer, 0, 8192)) != -1) {
                os.write(buffer, 0, bytesRead);
            }
            os.close();
            ins.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 判断文件大小
     * @param len
     * 文件长度
     * @param size
     *  限制大小
     * @param unit
     * 限制单位（B,K,M,G）
     * @return
     */
    public static boolean checkFileSize(Long len, int size, String unit) {
        double fileSize = 0;
        if ("B".equals(unit.toUpperCase())) {
            fileSize = (double) len;
        } else if ("K".equals(unit.toUpperCase())) {
            fileSize = (double) len / 1024;
        } else if ("M".equals(unit.toUpperCase())) {
            fileSize = (double) len / 1048576;
        } else if ("G".equals(unit.toUpperCase())) {
            fileSize = (double) len / 1073741824;
        }
        if (fileSize > size) {
            return false;
        }
        return true;
    }

    /**
     * 利用StringBuffer写文件
     * @param path
     * @param content
     */
    public static void StringBuff(String path,String content) throws IOException{
        File file=new File(path);
        if(!file.exists()){
            file.getParentFile().mkdirs();
            file.createNewFile();
        }
        FileOutputStream out=new FileOutputStream(file,true);
        StringBuffer sb=new StringBuffer();
        if(StringUtils.isEmpty(content)){
            sb.append(content);
            out.write(sb.toString().getBytes("utf-8"));
        }
        out.close();
    }

    /**
     * 文件重命名
     */
    public static void renameFile(String path,String oldname,String newname){
        if(!oldname.equals(newname)){//新的文件名和以前文件名不同时,才有必要进行重命名
            File oldfile=new File(path+"/"+oldname);
            File newfile=new File(path+"/"+newname);
            if(newfile.exists())//若在该目录下已经有一个文件和新文件名相同，则不允许重命名
                System.out.println(newname+"已经存在！");
            else{
                oldfile.renameTo(newfile);
            }
        }
    }

    /**
     * 文件转移目录
     * 转移文件目录不等同于复制文件，
     * 复制文件是复制后两个目录都存在该文件，而转移文件目录则是转移后，只有新目录中存在该文件
     * @param filename
     * @param oldpath
     * @param newpath
     * @param cover
     * @return
     */
    public String changeDirectory(String filename, String oldpath, String newpath, boolean cover) throws IOException {
        if(!oldpath.equals(newpath)){
            File oldfile=new File(oldpath+"/"+filename);
            File newfile=new File(newpath+"/"+filename);
            if(newfile.exists()){//若在待转移目录下，已经存在待转移文件
                if(cover) {//覆盖
                    oldfile.renameTo(newfile);
                }
                else{
                    System.out.println("在新目录下已经存在："+filename);
                }
            }
         else{
                oldfile.renameTo(newfile);
            }
        }
        return null;
    }



    /**
     * 利用FileInputStream读取文件
     * @param path
     * @return
     * @throws IOException
     */
        public static String FileInputStreamDemo(String path) throws IOException{
        File file=new File(path);
        if(!file.exists()||file.isDirectory())
            throw new FileNotFoundException();
        FileInputStream fis=new FileInputStream(file);
        byte[] buf = new byte[1024];
        StringBuffer sb=new StringBuffer();
        while((fis.read(buf))!=-1){
            sb.append(new String(buf));
            buf=new byte[1024];//重新生成，避免和上次读取的数据重复
        }
        return sb.toString();
    }

    /**
     * 利用bufferReader读取文件
     * @param path
     * @return
     * @throws IOException
     */
    public static String BufferedReaderDemo(String path) throws IOException{
        File file=new File(path);
        if(!file.exists()||file.isDirectory())
            throw new FileNotFoundException();
        BufferedReader br=new BufferedReader(new FileReader(file));
        String temp=null;
        StringBuffer sb=new StringBuffer();
        temp=br.readLine();
        while(temp!=null){
            sb.append(temp+" ");
            temp=br.readLine();
        }
        return sb.toString();
    }

    /**
     * 创建文件（夹）
     * @param path
     * @return
     * @throws IOException
     */
    public static File createDirFile(String path) throws IOException {
        File dir=new File(path);
        if(!dir.exists()){
            dir.getParentFile().mkdirs();
             dir.createNewFile();
        }
        return dir;
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

    // 获取文件后缀
    public static String getExt(String originalFilename){
        return originalFilename.substring(originalFilename.lastIndexOf(".") + 1);
    }


    public static MultipartFile getResourcesFile(String filePath, String type){
        ClassLoader classLoader = FileUtils.class.getClassLoader();
        URL resource = classLoader.getResource(filePath);
        String path = resource.getPath();
        return new MockMultipartFile("file","file"+type,type,fileConvertToByteArray(path));
    }

}
