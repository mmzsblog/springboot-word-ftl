package com.mmzsblog.word.utils;

import com.documents4j.api.DocumentType;
import com.documents4j.api.IConverter;
import com.documents4j.job.LocalConverter;
import sun.misc.BASE64Encoder;

import java.io.*;
import java.util.Random;

/**
 * @author ：mmzsblog.cn
 * @description：
 * @date ：2023/6/8 19:56
 */
public class ToolsUtil {

    public static void main(String[] args){
        word2pdf("E:\\tmp\\离职证明.doc", "E:\\\\tmp\\离职证明.pdf");
    }

    /**
     * word文档转成PDF文档(注意：要使用这个word转pdf的话，运行这个代码的服务器上需要安装WPS或者office软件)
     * @param wordPath word文档路径
     * @param pdfPath pdf文档路径
     */
    public static void word2pdf(String wordPath, String pdfPath){
        File inputWord = new File(wordPath);
        File outputFile = new File(pdfPath);
        InputStream docxInputStream = null;
        OutputStream outputStream = null;
        IConverter converter = null;
        try  {
            docxInputStream = new FileInputStream(inputWord);
            outputStream = new FileOutputStream(outputFile);
            converter = LocalConverter.builder().build();
            converter.convert(docxInputStream).as(DocumentType.DOCX).to(outputStream).as(DocumentType.PDF).execute();
            outputStream.close();
        } catch (Exception e) {
            e.printStackTrace();
//            log.error("word文档转成PDF文档时，发生异常：{}", e.getLocalizedMessage());
        } finally {
            //关闭资源
            try {
                if(docxInputStream != null){
                    docxInputStream.close();
                }
                if(outputStream != null){
                    outputStream.close();
                }
                if(converter != null){
                    converter.shutDown();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    //生成指定位数的随机数
    public static String getRandomLength(int length) {
        String val = "";
        Random random = new Random();
        for (int i = 0; i < length; i++) {
            val += String.valueOf(random.nextInt(10));
        }
        return val;
    }

    /**
     * 获得指定图片文件的base64编码数据
     * @param filePath 文件路径
     * @return base64编码数据
     */
    public static String getBase64ByPath(String filePath) {
        if(!hasLength(filePath)){
            return "";
        }
        File file = new File(filePath);
        if(!file.exists()) {
            return "";
        }
        InputStream in = null;
        byte[] data = null;
        try {
            in = new FileInputStream(file);
        } catch (FileNotFoundException e1) {
            e1.printStackTrace();
        }
        try {
            assert in != null;
            data = new byte[in.available()];
            in.read(data);
            in.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        BASE64Encoder encoder = new BASE64Encoder();
        return encoder.encode(data);
    }

    /**
     * @desc 判断字符串是否有长度
     */
    public static boolean hasLength(String str) {
        return org.springframework.util.StringUtils.hasLength(str);
    }

}
