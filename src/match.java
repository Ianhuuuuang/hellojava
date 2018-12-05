public class match {
    public static void main(String[] args) {
        String str1 = "ababcabcacbab";
        String str2 = "abcac";
        int s = str1.length();
        int t = str2.length();
        if (t==0||s<t){
            System.out.println("str1中未找到str2匹配项");
        }else{
            int i=0;int j=0;
            while (i<s&&j<t){
                if (str1.charAt(i)==str2.charAt(j)){
                    i++;j++;}
                else{
                    i=i-j+1;j=0;}
            if (j>=t){
                    int loc=i-t;
                    System.out.println("str1中第"+loc+"个字符开始与str2匹配");}
            else{
                    System.out.println("str1中未找到str2匹配项");
                }
            }
        }
    }
}
