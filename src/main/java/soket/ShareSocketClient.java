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
 * @Date 2018-1-26 15:26
 */
public class ShareSocketClient extends Socket {

    private static final String SERVER_IP ="127.0.0.1";
    private static final int SERVER_PORT =2018;

    private Socket client;
    private PrintWriter out;
    private BufferedReader in;
    
    public ShareSocketClient() throws Exception{
        super(SERVER_IP,SERVER_PORT);
        client = this;
        out = new PrintWriter(this.getOutputStream(),true);
        in = new BufferedReader(new InputStreamReader(this.getInputStream()));
        new ReadLineThread();
        
        while(true){
            in = new BufferedReader(new InputStreamReader(System.in));
            String input = in.readLine();
            out.println(input);
        }
    }
    
    class ReadLineThread extends Thread{
        
        private BufferedReader buff;
        public ReadLineThread(){
            try {
                buff = new BufferedReader(new InputStreamReader(client.getInputStream()));
                start();
            }catch (Exception e){
                e.printStackTrace();
            }
        }
        
        @Override
        public void  run(){
            try{
               while (true){
                   String result = buff.readLine();
                   if("byeClinet".equals(result)){
                       break;
                   }else {
                       System.out.println(result);
                   }
               }
                in.close();
                out.close();
                client.close();
            }catch ( Exception e){
                e.printStackTrace();
            }
        }
    }

    public static void main(String[] args) {
        try {
            new ShareSocketClient();
        }catch (Exception e){
            e.printStackTrace();
        }
    }
    
}
