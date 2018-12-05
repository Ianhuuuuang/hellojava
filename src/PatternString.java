import java.util.Scanner;

public class PatternString {

    public static int bruteForceStringMatch(String source, String pattern)
    {
        int slen = source.length();
        int plen = pattern.length();
        char[] s = source.toCharArray();
        char[] p = pattern.toCharArray();
        int i = 0;
        int j = 0;

        if(slen < plen)
            return -1;                        //如果主串长度小于模式串，直接返回-1，匹配失败
        else
        {
            while(i < slen && j < plen)
            {
                if(s[i] == p[j])            //如果i,j位置上的字符匹配成功就继续向后匹配
                {
                    ++i;
                    ++j;
                }
                else
                {
                    i = i - (j -1);            //i回溯到主串上一次开始匹配下一个位置的地方
                    j = 0;                    //j重置，模式串从开始再次进行匹配
                }
            }
            if(j == plen)                    //匹配成功
                return i - j;
            else
                return -1;                    //匹配失败
        }
    }


    public static int kmpStringMatch(String source, String pattern)
    {
        int i = 0;
        int j = 0;
        char[] s = source.toCharArray();
        char[] p = pattern.toCharArray();
        int slen = s.length;
        int plen = p.length;
        int[] next = getNext(p);
        while(i < slen && j < plen)
        {
            if(j == -1 || s[i] == p[j])
            {
                ++i;
                ++j;
            }
            else
            {
                //如果j != -1且当前字符匹配失败，则令i不变，
                //j = next[j],即让pattern模式串右移j - next[j]个单位
                j = next[j];
            }
        }


        if(j == plen)
            return i - j;
        else
            return -1;
    }

    private static int[] getNext(char[] p)
    {
        /**
         * 已知next[j] = k, 利用递归的思想求出next[j+1]的值
         * 1.如果p[j] = p[k]，则next[j+1] = next[k] + 1;
         * 2.如果p[j] != p[k],则令k = next[k],如果此时p[j] == p[k],则next[j+1] = k+1
         * 如果不相等，则继续递归前缀索引，令k=next[k],继续判断，直至k=-1(即k=next[0])或者p[j]=p[k]为止
         */
        int plen = p.length;
        int[] next = new int[plen];
        int k = -1;
        int j = 0;
        next[0] = -1;                //这里采用-1做标识
        while(j < plen -1)
        {
            if(k == -1 || p[j] == p[k])
            {
                ++k;
                ++j;
                next[j] = k;
            }
            else
            {
                k = next[k];
            }
        }
        System.out.println("next函数值：");
        for(int ii = 0;ii<next.length;ii++)
        {

            System.out.print(next[ii]+ " ");
        }
        System.out.println();
        return next;
    }

    public static void main(String[] args) {
        Scanner sc = new Scanner(System.in);
        String a = sc.nextLine();
        String b = sc.nextLine();
        System.out.println(bruteForceStringMatch(a, b));
        System.out.println(kmpStringMatch(a, b));
    }

}