package soket;

import java.io.BufferedReader;
import java.io.InputStreamReader;
import java.io.PrintWriter;
import java.net.Socket;

/**
 * Created with IntelliJ IDEA.
 *
 * @Author medo_zy
 * @Desciption:
 * @Date 2018-1-26 14:07
 */
public class MultiSocketClient {

    public static void main(String[] args) {
        try{
            
            Socket socket = new Socket("127.0.0.1",2018);
            socket.setSoTimeout(60000);

            PrintWriter printWriter = new PrintWriter(socket.getOutputStream(),true);
            BufferedReader bufferedReader = new BufferedReader(new InputStreamReader(socket.getInputStream()));
            
            String result = "";
            while(result.indexOf("bye") == -1){
                BufferedReader bufferedReader1 = new BufferedReader(new InputStreamReader(System.in));
                printWriter.println(bufferedReader1.readLine());
                printWriter.flush();
                
                result = bufferedReader.readLine();
                System.out.println("Server say:" + result);
                
            }
            printWriter.close();
            bufferedReader.close();
            socket.close();
            
        }catch (Exception e){
            e.printStackTrace();
        }
    }
}
