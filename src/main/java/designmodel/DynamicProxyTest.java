package designmodel;

import java.lang.reflect.InvocationHandler;
import java.lang.reflect.Method;
import java.lang.reflect.Proxy;

/**
 * Created with IntelliJ IDEA.
 *
 * @Author zy
 * @Desciption:
 * @Date 2018-1-26 17:48
 */
interface AbstractClass {

    public void show();

}

// 具体类A  
class ClassA implements AbstractClass {

    @Override
    public void show() {
        // TODO Auto-generated method stub  
        System.out.println("我是A类！");
    }
}

// 具体类B  
class ClassB implements AbstractClass {

    @Override
    public void show() {
        // TODO Auto-generated method stub  
        System.out.println("我是B类！");
    }
}

//动态代理类，实现InvocationHandler接口  
class Invoker implements InvocationHandler {
    AbstractClass ac;

    public Invoker(AbstractClass ac) {
        this.ac = ac;
    }

    @Override
    public Object invoke(Object proxy, Method method, Object[] arg)
            throws Throwable {
        //调用之前可以做一些处理  
        method.invoke(ac, arg);
        //调用之后也可以做一些处理  
        return null;
    }
}

/**
 * test
 */
public class DynamicProxyTest {

    public static void main(String[] args) {
        //创建具体类ClassB的处理对象  
        Invoker invoker1=new Invoker(new ClassA());
        //获得具体类ClassA的代理  
        AbstractClass ac1 = (AbstractClass) Proxy.newProxyInstance(
                AbstractClass.class.getClassLoader(),
                new Class[] { AbstractClass.class }, invoker1);
        //调用ClassA的show方法。  
        ac1.show();


        //创建具体类ClassB的处理对象  
        Invoker invoker2=new Invoker(new ClassB());
        //获得具体类ClassB的代理  
        AbstractClass ac2 = (AbstractClass) Proxy.newProxyInstance(
                AbstractClass.class.getClassLoader(),
                new Class[] { AbstractClass.class }, invoker2);
        //调用ClassB的show方法。  
        ac2.show();

    }
}  

