package Office.word.itextword;

import java.util.Map;

/**
 * Created with IntelliJ IDEA.
 *
 * @Author medo_zy
 * @Desciption:此类用于在浏览器端输出word，生成word后直接在浏览器端输出
 * @Date 2018-1-24 10:15
 */
public class ExportWord {

    public Map<String, Object> operateExcelFile(HttpServletResponse response, HttpServletRequest request) {
        try {
            doSomething();
            String fileName = map.get("fileName").toString();
            String fileUrl= map.get("fileUrl").toString();
            ResponseUtils.renderExcelFile(response,request,fileName, fileUrl);
            map.remove("fileUrl");
            map.put("success", true);
        } catch (Exception e) {
            e.printStackTrace();
            LOG.error("operateFile fail!", e);
            map.put("success", false);
            map.remove("fileUrl");
            map.put("data", getExcptionMessage(e));
        }
        return map;
    }
}
