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
        
        String originPath = "D:\\generateWord\\model\\model.docx";
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
