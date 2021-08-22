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

    public static  BigDecimal dividByPowerOf10(int decimalPlace) {

        int size = Integer.valueOf(decimalPlace);
        BigDecimal digit = new BigDecimal(1);
        BigDecimal ten = new BigDecimal(10);
        for (int i = 0; i < size; i++) {
            digit = digit.divide(ten);
        }

        return digit;
    }
    //傳入10 => 回傳 1。傳入1 => 回傳 1。傳入數字位數 0.3 =>回傳0.1。傳入數字位數 0.03 =>回傳0.01。
    public static BigDecimal OneInLastDigitPoint(int decimalplace) {
        if( decimalplace>0){
           return dividByPowerOf10(decimalplace);
        }


        else return  new BigDecimal(1);
    }
    public static void main(String[] args){
        System.out.println( countDecimalPlace("100.0"));
    }
}
