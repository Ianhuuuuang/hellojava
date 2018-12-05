import java.util.Scanner;

public class kmpStringMatch{
    public static int kmpmatch(String source,String pattern){
        char[]s=source.toCharArray();
        char[]p=pattern.toCharArray();
        int sl=s.length;
        int pl=p.length;
        int[]next=getnext(p);
        int i=0;
        int j=0;
        if (pl==0||sl<pl)
            return-1;
        else
            while (i < sl && j < pl)
            {
                if (j==-1 || s[i] == p[j])
                {
                    ++i;
                    ++j;
                }
                else
                    j = next[j];

            }
            if (j==pl)
                return i-j;
            else
                return -1;
    }
    public static int[] getnext(char[]p){
        int j=0;
        int k=-1;
        int pl=p.length;
        int[]next=new int[pl];
        next[0]=-1;
        while (j<pl-1)
        {
            if (k==-1||p[j]==p[k])
            {
                ++k;
                ++j;
                next[j]=k;
            }
            else k=next[k];
        }
        System.out.println("next数组：");
        for (int ii=0;ii<next.length;ii++)
        {
            System.out.print(next[ii] + " ");
        }
        System.out.println();
        return next;
    }
    public static void main(String[]args){
        Scanner sc=new Scanner(System.in);
        String a=sc.nextLine();
        String b=sc.nextLine();
        System.out.println(kmpmatch(a,b));
    }
}
