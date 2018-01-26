package soket;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.PrintWriter;
import java.net.ServerSocket;
import java.net.Socket;

/**
 * Created with IntelliJ IDEA.
 *
 * @Author medo_zy
 * @Desciption:
 * @Date 2018-1-26 14:18
 */
public class MultiSocketServer extends ServerSocket {
    
    private static final int SERVER_PORT =2018;

    public MultiSocketServer() throws IOException {
        super(SERVER_PORT);
        try{
            while(true){
                Socket socket = accept();
                new CreateServerThread(socket);
            }
        }catch (Exception e){
            
        }finally {
            close();
        }
    }

    class CreateServerThread extends Thread{
        private Socket client;
        private BufferedReader bufferedReader;
        private PrintWriter printWriter;
        
        public CreateServerThread(Socket s) throws IOException{
            
            client = s ;
            bufferedReader = new BufferedReader(new InputStreamReader(client.getInputStream()));
            
            printWriter = new PrintWriter(client.getOutputStream(),true);
            System.out.println("Client("+getName()+") come in...");
            
            start();
        }
        public synchronized void run(){
            try{
                for(int i = 0;;i++){
                    System.out.println(i);
                    String line = bufferedReader.readLine();
                    while(!line.equals("bye")){
                        printWriter.println("continue,clinet("+getName()+")!");
                        line = bufferedReader.readLine();
                        System.out.println("Client("+getName()+")!");
                    }
                    printWriter.println("bye,Clinet("+getName()+")!");

                    System.out.println("Client("+getName()+") exit!");
                    printWriter.close();
                    bufferedReader.close();
                    client.close();
                }
            }catch ( Exception e){
                
            }
        }
    }
    public static void main(String[] args) throws IOException {
        new MultiSocketServer();
    }
}
