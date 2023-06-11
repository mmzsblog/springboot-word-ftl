package com.mmzsblog.word.controller;

import com.mmzsblog.word.utils.FreemarkerCreateWordUtil;
import com.mmzsblog.word.utils.FreemarkerExportWordUtil;
import com.mmzsblog.word.utils.ToolsUtil;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

/**
 * @author ：mmzsblog.cn
 * @description：
 * @date ：2023/6/8 19:48
 */
@RestController
public class TestController {
    /**
     * 根据ftl模板模板生成word文档到指定目录
     */
    @GetMapping("/createWord")
    public void createWord(){
        Map<String, Object> map = new HashMap<>();
        map.put("dateTime", "2023-06-11");
        map.put("testResult", "测试通过");
        map.put("userName", "王大锤");
        try {
            FreemarkerCreateWordUtil.createWordByFtl(map, "测试word.doc", "testWord.ftl");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    /**
     * 根据ftl模板模板下载word文档
     */
    @GetMapping("/downloadWord")
    public void downloadWord(HttpServletResponse response){
        Map<String, Object> map = new HashMap<>();
        map.put("dateTime", "2023-06-11");
        map.put("testResult", "测试通过");
        map.put("userName", "王大锤");
        map.put("imageData", ToolsUtil.getBase64ByPath("E:\\tmp\\mmzsblog.jpg"));//测试图片Base64数据（不包含头信息）
        try {
            FreemarkerExportWordUtil.exportWord(response, map, "测试wordImage.doc", "testWordImage.ftl");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}
