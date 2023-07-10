package com.mmzsblog.word.docx;

import freemarker.template.Configuration;
import freemarker.template.Template;
import org.springframework.util.ResourceUtils;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.util.HashMap;
import java.util.Map;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

import static freemarker.template.Configuration.DEFAULT_INCOMPATIBLE_IMPROVEMENTS;


/**
 * @author ：mmzsblog.cn
 * @description：模板word导出docx格式，工具类
 * @date ：2023/7/10 22:37
 */
public class DocxUtil {

    /**
     * 生成docx文档： 依赖resources资源文件夹下的templates/docxTemplate中的zip和xml文件
     */
    public static void main(String[] args) throws Exception {
        String fileDirectory = ResourceUtils.getFile("classpath:templates/docxTemplate").getPath();
        String docxZipPath = ResourceUtils.getFile("classpath:templates/docxTemplate/测试docx.zip").getPath();
        String outputDocxName = "E:/tmp/测试docx.docx";
        String itemName = "word/document.xml";
        Map<String, Object> dataMap = new HashMap<>();
        dataMap.put("dateTime", "2023-06-11");
        dataMap.put("testResult", "测试通过");
        dataMap.put("userName", "王大锤");
        DocxUtil.ftlToDocx(fileDirectory, docxZipPath, outputDocxName, itemName, dataMap);
    }



    /**
     * @param fileDirectory  模板路径 "/templates/docxTemplate/word";
     * @param docxZipPath    zip文件的zip输入流
     * @param outputDocxName 输出的zip输出流
     * @param itemName       要替换的 item 名称 一般固定"word/document.xml"
     * @param dataMap
     * @throws Exception
     */
    public static void ftlToDocx(String fileDirectory, String docxZipPath,
                                 String outputDocxName, String itemName,
                                 Map<String, Object> dataMap) throws Exception {
        /** 初始化配置文件 **/
        Configuration configuration = new Configuration(DEFAULT_INCOMPATIBLE_IMPROVEMENTS);
        /** 加载文件 **/
        configuration.setDirectoryForTemplateLoading(new File(fileDirectory));
        /** 加载模板 **/
        Template template = configuration.getTemplate("document.xml");

        /** 指定输出word文件的路径 **/
        String outFilePath = fileDirectory + File.separator + "data.xml";
        File docFile = new File(outFilePath);
        FileOutputStream fos = new FileOutputStream(docFile);
//        OutputStreamWriter oWriter = new OutputStreamWriter(fos);
        Writer out = new BufferedWriter(new OutputStreamWriter(fos), 10240);
        template.process(dataMap, out);

        if (out != null) {
            out.close();
        }
        // ZipUtils 是一个工具类
        ZipInputStream zipInputStream = ZipUtils.wrapZipInputStream(new FileInputStream(new File(docxZipPath)));
        ZipOutputStream zipOutputStream = ZipUtils.wrapZipOutputStream(new FileOutputStream(new File(outputDocxName)));
        ZipUtils.replaceItem(zipInputStream, zipOutputStream, itemName, new FileInputStream(new File(outFilePath)));
    }

}
