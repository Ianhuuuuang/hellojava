public class sl_5_3 {
    public static void main(String[] args){
        String regex="\\d{11}";
        String str="15950550785";
        if(str.matches(regex)){
            System.out.println("字符串是合法手机号码");
        }
        else{
            System.out.println("字符串是非法手机号码");
        }
    }
}
