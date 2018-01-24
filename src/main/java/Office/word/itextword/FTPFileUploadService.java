package Office.word.itextword;

import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;
import sun.net.ftp.FtpClient;
import sun.net.ftp.FtpProtocolException;

import java.io.*;
import java.net.InetSocketAddress;
import java.net.SocketAddress;
import java.text.Format;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * Created by Administrator on 2016/12/6.
 */
@SuppressWarnings("unchecked")
@Service("fTPFileUploadService")
public class FTPFileUploadService {
    private static FtpClient ftpClient;
    @Value("${ftp.ip}")
    private String ip;
    @Value("${ftp.port}")
    private int port;
    @Value("${ftp.user}")
    private String user;
    @Value("${ftp.password}")
    private String password;
    @Value("${ftp.path}")
    private String path;
    private String localfilename;
    private String remotefilename;

    public void connectServer() {
        try {
            /*******连接服务器的两种方法*******/
            ftpClient = FtpClient.create();
            try {
                // 第一种方法
                SocketAddress addr = new InetSocketAddress(ip, port);
                ftpClient.connect(addr);
                ftpClient.login(user, password.toCharArray());
                  /* ******连接服务器的两种方法****** */
                // 第二种方法
                  /* ftpClient = FtpClient.create(ip);
                    ftpClient.login(user, null, password);*/
                // 设置成2进制传输
                ftpClient.setBinaryType();
                Date date = new Date();
                Format format = new SimpleDateFormat("yyyyMMdd");
                String urlPath = path + "/" + format.format(date);
                try {
//                    if(!isDirExist(urlPath))  判断是否存在，sun的api好像没
                    ftpClient.makeDirectory(urlPath);
                } catch (FtpProtocolException ex) {
                    ftpClient.changeDirectory(urlPath);
                }

            } catch (FtpProtocolException e) {
                e.printStackTrace();
            }
        } catch (IOException ex) {
            ex.printStackTrace();
            throw new RuntimeException(ex);
        }
    }

    /** 验证文件夹是否存在
     *@author flychen
     *@date 2016-12-6
     */
/*    public boolean isDirExist(String dir) {
        try {
            int a = ftpClient.cd(dir);
        } catch (Exception e) {
            return false;
        }
        return true;
    }*/

    /**
     * 关闭连接
     *
     * @author flychen
     * @date 2016-12-6
     */
    public void closeConnect() {
        try {
            ftpClient.close();
        } catch (IOException ex) {
            ex.printStackTrace();
            throw new RuntimeException(ex);
        }
    }

    /**
     * 上传文件
     *
     * @param multipartFile 本地文件
     * @author flychen
     * @date 2016-12-6
     */
    public String upload(MultipartFile multipartFile) {
        connectServer();
        OutputStream os = null;
        InputStream is = null;
        Date date = new Date();
        Format format = new SimpleDateFormat("yyyyMMdd");
        Format formatTime = new SimpleDateFormat("yyyyMMddHHmmss");
        //获取文件后缀
        String originFileName = multipartFile.getOriginalFilename().substring(multipartFile.getOriginalFilename().lastIndexOf("."));
        //根据时间重新定义文件名称
        remotefilename = formatTime.format(date) + originFileName;
        localfilename = format.format(date) + "/" + remotefilename;
        try {
            //将远程文件加入输出流中
            try {
                os = ftpClient.putFileStream(remotefilename);
            } catch (FtpProtocolException e) {
                e.printStackTrace();
            }
            is = multipartFile.getInputStream();
            byte[] bytes = new byte[1024];
            int c;
            while ((c = is.read(bytes)) != -1) {
                os.write(bytes, 0, c);
            }
        } catch (IOException ex) {
            ex.printStackTrace();
            throw new RuntimeException(ex);
        } finally {
            try {
                if (is != null) {
                    is.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            } finally {
                try {
                    if (os != null) {
                        os.close();
                        closeConnect();
                    }
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
        return localfilename;
    }

    /**
     * 下载文件
     *
     * @param remoteFile 远程文件路径(服务器端)
     * @param localFile  本地文件路径(客户端)
     * @author flychen
     * @date 2016-12-6
     */
    public void download(String remoteFile, String localFile) {
        connectServer();
        InputStream is = null;
        FileOutputStream os = null;
        try {
            //获取远程机器上的文件filename，借助TelnetInputStream把该文件传送到本地。
            try {
                is = ftpClient.getFileStream(remoteFile);
            } catch (FtpProtocolException e) {
                e.printStackTrace();
            }
            File file_in = new File(localFile);
            os = new FileOutputStream(file_in);
            byte[] bytes = new byte[1024];
            int c;
            while ((c = is.read(bytes)) != -1) {
                os.write(bytes, 0, c);
            }
        } catch (IOException ex) {
            ex.printStackTrace();
            throw new RuntimeException(ex);
        } finally {
            try {
                if (is != null) {
                    is.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            } finally {
                try {
                    if (os != null) {
                        os.close();
                        closeConnect();
                    }
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

}
