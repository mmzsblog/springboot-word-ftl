package com.mmzsblog.word.other;

import com.mmzsblog.word.docx.ZipUtils;
import freemarker.template.Configuration;
import freemarker.template.Template;
import org.springframework.core.io.ClassPathResource;
import org.springframework.util.ResourceUtils;
import sun.misc.BASE64Encoder;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.util.Map;
import java.util.Objects;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

import static freemarker.template.Configuration.DEFAULT_INCOMPATIBLE_IMPROVEMENTS;

/**
 * 模板word导出，工具类
 */
public class FreemarkerCreateDocxWordUtil {

    /**
     * 本地图片转换Base64的方法
     *
     * @param imgPath
     */
    public static String imageToBase64(String imgPath) {
        byte[] data = null;
        // 读取图片字节数组
        try {
            // 打包后Spring试图访问文件系统路径，但无法访问JAR中的路径。 必须使用resource.getInputStream()
            ClassPathResource classPathResource = new ClassPathResource(imgPath);
            InputStream in = classPathResource.getInputStream();
            data = new byte[in.available()];
            in.read(data);
            in.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        // 对字节数组Base64编码
        BASE64Encoder encoder = new BASE64Encoder();
        // 返回Base64编码过的字节数组字符串
        return encoder.encode(Objects.requireNonNull(data));
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
        // ZipUtils 是一个工具类，主要用来替换具体可以看github工程
        ZipInputStream zipInputStream = ZipUtils.wrapZipInputStream(new FileInputStream(new File(docxZipPath)));
        ZipOutputStream zipOutputStream = ZipUtils.wrapZipOutputStream(new FileOutputStream(new File(outputDocxName)));
        ZipUtils.replaceItem(zipInputStream, zipOutputStream, itemName, new FileInputStream(new File(outFilePath)));
    }

    /**
     * 导出签名的doc文档
     */
    public static void exportMillCertificateWordDoc(HttpServletResponse response, Map map, String ftlFile, String
            fileName, String path) throws IOException {
        /** 初始化配置文件 **/
        Configuration configuration = new Configuration(DEFAULT_INCOMPATIBLE_IMPROVEMENTS);
        /** 加载文件 **/
        configuration.setDirectoryForTemplateLoading(new File("xxxxxxx"));

        Template freemarkerTemplate = configuration.getTemplate(ftlFile);
        File file = null;
        InputStream fin = null;


        OutputStream os = null;

        try {
            // 调用工具类的createDoc方法生成Word文档
            file = createDoc(map, freemarkerTemplate);


//            ClassPathResource classPathResource = new ClassPathResource("files");
            // 保存word至bash文件夹
            // 1K的数据缓冲
            byte[] bs = new byte[1024];
            // 读取到的数据长度
            int len;

            File tempFile = new File(path);
            if (!tempFile.exists()) {
                tempFile.mkdirs();
            }


            fin = new FileInputStream(file);
            os = new FileOutputStream(tempFile.getPath() + File.separator + fileName);
            // 开始读取
            while ((len = fin.read(bs)) != -1) {
                os.write(bs, 0, len);
            }

        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (fin != null) {
                fin.close();
            }
            if (file != null) {
                file.delete(); // 删除临时文件
            }
        }
    }


    private static File createDoc(Map<?, ?> dataMap, Template template) {
        String name = "sellPlan.doc";
        File f = new File(name);
        Template t = template;
        try {
            // 这个地方不能使用FileWriter因为需要指定编码类型否则生成的Word文档会因为有无法识别的编码而无法打开
            Writer w = new OutputStreamWriter(new FileOutputStream(f), "utf-8");
            t.process(dataMap, w);
            w.close();
        } catch (Exception ex) {
            ex.printStackTrace();
            throw new RuntimeException(ex);
        }
        return f;
    }


}