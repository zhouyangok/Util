package Office.word.itextword;

import com.alibaba.fastjson.JSON;
import com.medo.core.utils.logging.Log;
import com.medo.core.utils.logging.LogFactory;
import org.apache.commons.lang3.StringUtils;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.URLEncoder;
import java.util.Map;

public class ResponseUtils {

    // header 常量定义
    private static final String ENCODING_PREFIX = "encoding";
    private static final String NOCACHE_PREFIX = "no-cache";
    private static final String ENCODING_DEFAULT = "UTF-8";
    private static final boolean NOCACHE_DEFAULT = true;

    private static final Log LOG = LogFactory.getLog(ResponseUtils.class);

    /**
     * 直接输出内容的简便函数.
     * 
     * eg. render("text/plain", "hello", "encoding:GBK"); render("text/plain",
     * "hello", "no-cache:false"); render("text/plain", "hello", "encoding:GBK",
     * "no-cache:false");
     * 
     * @param headers
     *            可变的header数组，目前接受的值为"encoding:"或"no-cache:",默认值分别为UTF-8和true.
     */
    public static void render(HttpServletResponse response,
            final String contentType, final String content,
            final String... headers) {
        try {
            // 分析headers参数
            String encoding = ENCODING_DEFAULT;
            boolean noCache = NOCACHE_DEFAULT;
            for (String header : headers) {
                String headerName = StringUtils.substringBefore(header, ":");
                String headerValue = StringUtils.substringAfter(header, ":");

                if (StringUtils.equalsIgnoreCase(headerName, ENCODING_PREFIX)) {
                    encoding = headerValue;
                } else if (StringUtils.equalsIgnoreCase(headerName,
                        NOCACHE_PREFIX)) {
                    noCache = Boolean.parseBoolean(headerValue);
                } else {
                    throw new IllegalArgumentException(headerName
                            + "不是一个合法的header类型");
                }
            }

            // 设置headers参数
            String fullContentType = contentType + ";charset=" + encoding;
            response.setContentType(fullContentType);
            if (noCache) {
                response.setHeader("Pragma", "No-cache");
                response.setHeader("Cache-Control", "no-cache");
                response.setDateHeader("Expires", 0);
            }

            response.getWriter().write(content);
            response.getWriter().flush();

        } catch (IOException e) {
            LOG.error(e.getMessage(), e);
        }
    }

    /**
     * 直接输出文本.
     *
     */
    public static void renderText(HttpServletResponse response,
            final String text, final String... headers) {
        render(response, "text/plain", text, headers);
    }

    /**
     * 直接输出HTML.
     * 
     */
    public static void renderHtml(HttpServletResponse response,
            final String html, final String... headers) {
        render(response, "text/html", html, headers);
    }

    /**
     * 直接输出XML.
     * 
     */
    public static void renderXml(HttpServletResponse response,
            final String xml, final String... headers) {
        render(response, "text/xml", xml, headers);
    }

    /**
     * 直接输出JSON.
     * 
     * @param string
     *            json字符串.
     */
    public static void renderJson(HttpServletResponse response,
            final String string, final String... headers) {
        render(response, "application/json", string, headers);
    }

    /**
     * 直接输出JSON.
     * 
     * @param map
     *            Map对象,将被转化为json字符串.
     */
    @SuppressWarnings("unchecked")
    public static void renderJson(HttpServletResponse response, final Map map,
            final String... headers) {
        String jsonString = JSON.toJSONString(map);
        renderJson(response, jsonString, headers);
    }

    /**
     * 直接输出JSON.
     * 
     * @param object
     *            Java对象,将被转化为json字符串.
     */
    public static void renderJson(HttpServletResponse response,
            final Object object, final String... headers) {
        String jsonString = JSON.toJSONString(object);
        renderJson(response, jsonString, headers);
    }

    /**
     * 输出excel
     * 
     * @param response
     * @param excel
     * @param headers
     */
    /*
     * public static void renderExcel(HttpServletResponse response, final String
     * excel, final String... headers) { render(response, "excel.xls", excel,
     * headers); }
     */

    /**
     * 直接输出文件.
     * 
     * @throws IOException
     */
    public static void renderFile(HttpServletResponse response,
            String fileName, byte[] data) throws IOException {
        String fileNameUTF8 = URLEncoder.encode(fileName, "UTF-8");
        response.reset();
        response.setHeader("Content-Disposition", "attachment; filename=\""
                + fileNameUTF8 + "\"");
        response.addHeader("Content-Length", "" + data.length);
        response.setContentType("application/octet-stream;charset=UTF-8");
        OutputStream outputStream = new BufferedOutputStream(
                response.getOutputStream());
        outputStream.write(data);
        outputStream.flush();
        outputStream.close();
    }

    /**
     * 直接excel文件.
     *
     * @param
     * @throws IOException
     * @see #(String, String, String...)
     */
    public static void renderExcelFile(HttpServletResponse response, HttpServletRequest request,
                                       String fileName, String fileUrl) throws IOException {
        File file = new File(fileUrl);
        //文件输入流
        FileInputStream inputStream = new FileInputStream(file);
        //缓冲文件输入流
        BufferedInputStream bufferedInputStream = new BufferedInputStream(inputStream);

        byte[] buff = new byte[inputStream.available()];
        inputStream.read(buff);
        bufferedInputStream.close();
        inputStream.close();
        String fileNameUTF8 = URLEncoder.encode(fileName, "UTF8");//其他浏览器
        response.reset();
        response.setHeader("Content-disposition", fileNameUTF8);
        response.setContentType("application/msword");
        response.setHeader("Content-disposition", "attachment; filename=\""
                + fileNameUTF8 + "\"");
        response.addHeader("Content-Length", "" + file.length());
        response.setHeader("Pragma", "No-cache");
        OutputStream outputStream = new BufferedOutputStream(response.getOutputStream());

        outputStream.write(buff);
        //流的关闭
        outputStream.flush();//强制清除缓冲区的内容

        outputStream.close();

    }

}
