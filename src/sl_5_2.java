public class sl_5_2 {
    public static void main(String[] args){
        String str1="HelloJava";
        String str2="Helloworld";
        String newstr1=str1.substring(0,4);
        String newstr2=str2.substring(0,4);
        if (newstr1.equals(newstr2)){
            System.out.println("两个字符串相同");
        }
        else{
            System.out.println("两个字符串不相同");
        }
    }
}
