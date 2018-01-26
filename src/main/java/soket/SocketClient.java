package soket;

import java.io.BufferedReader;
import java.io.InputStreamReader;
import java.io.PrintWriter;
import java.net.Socket;

/**
 * Created with IntelliJ IDEA.
 *
 * @Author zy
 * @Desciption:
 * @Date 2018-1-26 11:10
 */
public class SocketClient {
    public static void main(String[] args) {
       try{
           Socket socket = new Socket("127.0.0.1",2018);
           socket.setSoTimeout(60000);

           PrintWriter printWriter = new PrintWriter(socket.getOutputStream(),true);
           BufferedReader buff = new BufferedReader(new InputStreamReader(System.in));
           
           printWriter.print(buff.readLine());
           printWriter.flush();
           
           BufferedReader bufferedReader = new BufferedReader(new InputStreamReader(socket.getInputStream()));
           
           String result = bufferedReader.readLine();
           System.out.println("Server say:" + result);
           printWriter.close();
           buff.close();
           bufferedReader.close();
           socket.close();
       }catch (Exception e){
           System.out.println("Exception:" + e);
       }
        
    }
}
