package common;

import java.math.BigDecimal;

public class NumberProcessor {
    public static int getIndexOfLargest( int[] array )
    {
        if ( array == null || array.length == 0 ) return -1; // null or empty

        int largest = 0;
        for ( int i = 1; i < array.length; i++ )
        {
            if ( array[i] > array[largest] ) largest = i;
        }
        return largest; // position of the first largest found
    }

    //計算數字位數
    public static int countDecimalPlace(String Number){
        BigDecimal a = new BigDecimal(Number);
        return a.scale();
    }
    public static void main(String[] args){
        System.out.println( countDecimalPlace("100.0"));
    }
}
