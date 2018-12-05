public class sl4_11 {                                  //创建类
    public static void main(String[] args) {           //创建主方法
        String str="";
        long starTime=System.currentTimeMillis();
        for (int i=0;i<10000;i++){
            str=str+i;
        }
        long endTime=System.currentTimeMillis();
        long time=endTime-starTime;
        System.out.println("String运算花费："+time+"s");
        StringBuilder builder=new StringBuilder("");
        starTime=System.currentTimeMillis();
        for (int j=0;j<10000;j++){
            builder.append(j);
        }
        endTime=System.currentTimeMillis();
        time=endTime-starTime;
        System.out.println("String运算花费："+time+"s");
    }
}



