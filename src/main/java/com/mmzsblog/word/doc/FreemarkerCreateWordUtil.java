package com.mmzsblog.word.doc;

import freemarker.template.Configuration;
import freemarker.template.Template;
import freemarker.template.Version;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.Map;

/**
 * 模板word创建生成，工具类
 */
public class FreemarkerCreateWordUtil {

    private static Configuration configuration = null;

    static {
        configuration = new Configuration(new Version("2.3.0"));
        configuration.setDefaultEncoding(StandardCharsets.UTF_8.toString());
        //获取模板路径    setClassForTemplateLoading 这个方法默认路径是webRoot 路径下
        configuration.setClassForTemplateLoading(FreemarkerCreateWordUtil.class, "/templates");
    }

    private FreemarkerCreateWordUtil() {
        throw new AssertionError();
    }

    /**
     * 根据 /resources/templates 目录下的ftl模板文件生成文件并写到客户端
     *
     * @param map         数据集合
     * @param fileName    用户下载到的文件名称
     * @param ftlFileName ftl模板文件名称
     * @throws IOException
     */
    public static void createWordByFtl(Map map, String fileName, String ftlFileName) throws IOException {
        Template freemarkerTemplate = configuration.getTemplate(ftlFileName);
        // 生成的文件路径
        String filePath = "E:/tmp/" + fileName;
        // 调用工具类的createDoc方法生成Word文档
        File file = createDoc(map, freemarkerTemplate, filePath);
    }

    private static File createDoc(Map<?, ?> dataMap, Template template, String filePath) {
        File f = new File(filePath);
        Template t = template;
        try {
            // 这个地方不能使用FileWriter因为需要指定编码类型否则生成的Word文档会因为有无法识别的编码而无法打开
            Writer w = new OutputStreamWriter(new FileOutputStream(f), StandardCharsets.UTF_8.toString());
            t.process(dataMap, w);
            w.close();
        } catch (Exception ex) {
            ex.printStackTrace();
            throw new RuntimeException(ex);
        }
        return f;
    }

}