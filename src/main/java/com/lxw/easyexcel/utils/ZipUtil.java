package com.lxw.easyexcel.utils;

import lombok.extern.slf4j.Slf4j;
import org.apache.tools.zip.ZipEntry;
import org.apache.tools.zip.ZipFile;
import org.apache.tools.zip.ZipOutputStream;
import org.springframework.mock.web.MockMultipartFile;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.util.Enumeration;
import java.util.List;

import static org.springframework.util.StreamUtils.BUFFER_SIZE;

/**
 * @author liang.xiongwei
 * @Title: ZipUtil
 * @Package com.intellif.vesionbook.admin.utils
 * @Description
 * @date 2021/1/8 11:12
 */
@Slf4j
public class ZipUtil {
    /**
     * 把文件集合打成zip压缩包
     * @param srcFiles 压缩文件集合
     * @throws RuntimeException 异常
     */
    public static MultipartFile toZip(List<MultipartFile> srcFiles) throws RuntimeException {
        long start = System.currentTimeMillis();
        ZipOutputStream zos = null;
        try {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            zos = new ZipOutputStream(out);
            for (MultipartFile srcFile : srcFiles) {
                byte[] buf = new byte[BUFFER_SIZE];
                zos.putNextEntry(new ZipEntry(srcFile.getOriginalFilename()));
                int len;
                InputStream in = srcFile.getInputStream();
                while ((len = in.read(buf)) != -1) {
                    zos.write(buf, 0, len);
                }
                zos.setComment("我是注释");
                zos.closeEntry();
                in.close();
            }
            zos.close();
            out.close();
            long end = System.currentTimeMillis();
            log.info("压缩完成，耗时：" + (end - start) + " ms");
            MultipartFile resultFile = new MockMultipartFile("file", "file.zip", "zip", out.toByteArray());
            return  resultFile;
        } catch (Exception e) {
            log.error("ZipUtil toZip exception, ", e);
            throw new RuntimeException("zipFile error from ZipUtils", e);
        } finally {
            if (zos != null) {
                try {
                    zos.close();
                } catch (IOException e) {
                    log.error("ZipUtil toZip close exception, ", e);
                }
            }
        }
    }
}
