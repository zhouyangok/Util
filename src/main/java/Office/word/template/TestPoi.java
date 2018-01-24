package Office.word.template;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import java.util.UUID;

/**
 * Created with IntelliJ IDEA.
 *
 * @Author medo_zy
 * @Desciption:
 * @Date 2017-12-15 9:59
 */
public class TestPoi {

    public static void main(String[] args) throws IOException {
        Map<String, Object> param=new HashMap<String, Object>();

        Map<String,Object> header = new HashMap<String, Object>();
        Map<String,Object> header2 = new HashMap<String, Object>();
        Map<String,Object> header3 = new HashMap<String, Object>();
        param.put("${typeName}", "神舟一号");
        param.put("${seriesName}", "神舟系列");
        param.put("${measureStatus}", "振前测量");
        param.put("${txx}", "2.123334444443");
        param.put("${txy}", "-2.32355555555");
        param.put("${txz}", "2.123");
        param.put("${tyx}", "5.45555523");
        param.put("${tyy}", "7.123");
        param.put("${tyz}", "3.823");
        param.put("${tzx}", "-1.623");
        param.put("${tzy}", "6.177777743");
        param.put("${tzz}", "5.462");
        param.put("${tx}", "8.123");
        param.put("${ty}", "1.76777777773");
        param.put("${tz}", "-2.873");
        param.put("${xx}", "9.323");
        param.put("${xy}", "2.12443");
        param.put("${xz}", "2.12888888883");
        param.put("${yx}", "123.5553");
        param.put("${yy}", "12.2443");
        param.put("${yz}", "1.437777777723");
        param.put("${zx}", "7.4355523");
        param.put("${zy}", "13.4666323");
        param.put("${zz}", "2.4325553");
        param.put("${x}", "2.43266666663");
        param.put("${y}", "2.43288888883");
        param.put("${z}", "2.4323");
        param.put("${rxx}", "9.323");
        param.put("${rxy}", "2.123");
        param.put("${rxz}", "2.123");
        param.put("${ryx}", "123.3");
        param.put("${ryy}", "12.23");
        param.put("${ryz}", "1.4323");
        param.put("${rzx}", "7.4323");
        param.put("${rzy}", "13.433333323");
        param.put("${rzz}", "2.4323");
        param.put("${rx}", "-2.4323");
        param.put("${ry}", "-2.4323444444");
        param.put("${rz}", "-2.4355555523");
        header.put("width", 650);
        header.put("height", 200);
        header.put("type", "png");
        header.put("content",  WordUtil.inputStream2ByteArray(new FileInputStream("D:\\generateWord\\echarts1.png"),true));
        param.put("${pic1}",header);
        header2.put("width", 650);
        header2.put("height", 200);
        header2.put("type", "png");
        header2.put("content",  WordUtil.inputStream2ByteArray(new FileInputStream("D:\\generateWord\\echarts2.png"),true));
         header3.put("width", 650);
        header3.put("height", 200);
        header3.put("type", "png");
        header3.put("content",  WordUtil.inputStream2ByteArray(new FileInputStream("D:\\generateWord\\echarts3.png"),true));
        
        param.put("${pic2}",header2);
        param.put("${pic3}",header3);
        
        String originPath = "D:\\generateWord\\model\\model2.docx";
        String finalPath = "D:\\generateWord\\";
        String fileName = UUID.randomUUID().toString()+".docx";
        generWord(param,originPath,finalPath,fileName);
       
    }
    public static void generWord(Map param,String originPath,String finalPath,String fileName)throws IOException{
        CustomXWPFDocument doc=WordUtil.generateWord(param, originPath);
        FileOutputStream fopts = new FileOutputStream(finalPath+fileName);
        doc.write(fopts);
        fopts.close();
    }
}
