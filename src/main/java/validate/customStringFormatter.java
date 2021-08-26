package validate;

import common.NumberProcessor;
import common.StringProcessor;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;

public class customStringFormatter {


    //2.傳出Acceptable Range 字串
    public static String getAccecptableRangeString(MyRange myRange) {
        String minRange = "";
        String maxRange = "";
        String andSign = " & ";

        //若有 Not inverse
        if (myRange.hasMax()) {
            maxRange = OperatorConvertor.getAcceptableSymbol(false, myRange.getMaxEqualTo()) + myRange.getMax();
        }

        if (myRange.hasMin()) {

            minRange = OperatorConvertor.getAcceptableSymbol(true, myRange.getMinEqualTo())+myRange.getMin();
        }

        if (!minRange.isEmpty() && !maxRange.isEmpty()) {
            return minRange + andSign + maxRange;

        } else return minRange + maxRange;

    }

    public static List<String> sperateStringByKeywords(String string, List<String> keywords) {
        int beginIndex = 0;
        int endIndex = 0;
        List<String> subStrings = new ArrayList<>();
        for (int i = 0; i < keywords.size(); i++) {
            String keyword = keywords.get(i);
            endIndex = string.indexOf(keyword);
            String subString = string.substring(beginIndex, endIndex);
            subStrings.add(subString);
            beginIndex = endIndex;

        }
        return subStrings;
    }

    public static List<Double> setOOS (MyRange range ) throws RangeException {
        List<Double> OOSNum =new ArrayList<>();


        // 有上下限
        if (range.hasMax()&&range.hasMin()){


            BigDecimal bd1 = range.getMin();
            //若數字有等於符號，OOS數字: 上限需加1， 下限需減1

            if(range.getMinEqualTo()){
                BigDecimal oneForMin = NumberProcessor.OneInLastDigitPoint(range.getMinDecimalPlace());
                //若為isGreater的operator，減一使數字(bd1)落在range外
                bd1=bd1.subtract(oneForMin);
            }
            OOSNum.add(bd1.doubleValue());

            BigDecimal bd2 = range.getMax();
            if(range.getMaxEqualTo()){
                BigDecimal oneForMax = NumberProcessor.OneInLastDigitPoint(range.getMaxDecimalPlace());
                bd2=bd2.add(oneForMax);
            }
            OOSNum.add(bd2.doubleValue());


        }

//只有上限或下限
        else if (!range.hasMax()&&range.hasMin()){
            BigDecimal bd1 = range.getMin();

            //若數字有等於符號，OOS數字
            if(range.getMinEqualTo()){
                BigDecimal oneForMin = NumberProcessor.OneInLastDigitPoint(range.getMinDecimalPlace());
                //若為isGreater的operator，減一使數字(bd1)落在range外
                bd1=bd1.subtract(oneForMin);
            }
            OOSNum.add(bd1.doubleValue());

        }
        else if (range.hasMax()&&!range.hasMin()){
            BigDecimal bd2 = range.getMax();
            if(range.getMaxEqualTo()){
                BigDecimal oneForMax = NumberProcessor.OneInLastDigitPoint(range.getMaxDecimalPlace());
                bd2=bd2.add(oneForMax);
            }
            OOSNum.add(bd2.doubleValue());



        }
        else {
            throw new RangeException("neither max nor min");
        }

        return OOSNum;
    }

    public static boolean isNumeric(String strNum) {
        if (strNum == null) {
            return false;
        }
        try {
            double d = Double.parseDouble(strNum);
        } catch (NumberFormatException nfe) {
            return false;
        }
        return true;
    }
    public static MyRange MyRangeGenerator(String specification) {

        MyRange range = new MyRange();
        BigDecimal min;
        BigDecimal max;
        Boolean minEqualTo;
        Boolean maxEqualTo;
        List<String> extractdigitFromString = StringProcessor.extractStringByPattern(specification, StringProcessor.extract_digit_rex);
        List<BigDecimal> extractdigits = new ArrayList<>();
        for (String s : extractdigitFromString) {
            if (isNumeric(s)){
                extractdigits.add(new BigDecimal(s));
            }
        }
        List<String> subStrings = sperateStringByKeywords(specification, extractdigitFromString);


        //有上下限，會傳入兩個數字
        if (extractdigitFromString.size() == 2) {

int position_max;
int position_min;

            BigDecimal num0 = extractdigits.get(0);
            BigDecimal num1 = extractdigits.get(1);
            if(num0.compareTo(num1)==1){
                //num0 is greater than num1
                position_max =0;
                position_min =1;
            }else if (num0.compareTo(num1)==-1){
                //num0 is greater than num1
                position_max =1;
                position_min =0;
            }else{//==0
                //error actually 兩數字不太可能相等
                position_max =0;
                position_min =1;
            }

            min =extractdigits.get(position_min);
            max = extractdigits.get(position_max);

            minEqualTo = OperatorConvertor.isEqual(subStrings.get(position_min));
            maxEqualTo = OperatorConvertor.isEqual(subStrings.get(position_max));

            range.setMinRange(min, minEqualTo);
            range.setMaxRange(max, maxEqualTo);

            int minDecimalplace = NumberProcessor.countDecimalPlace(extractdigitFromString.get(position_min));
            int maxDecimalplace = NumberProcessor.countDecimalPlace(extractdigitFromString.get(position_max));


            range.setMinDecimalPlace(minDecimalplace);
            range.setMaxDecimalPlace(maxDecimalplace);
        }
        //只有上下限其中一個，會傳入一個數字
        if (extractdigitFromString.size() == 1) {

            //判斷是大於或小於
            if (OperatorConvertor.isGreater(subStrings.get(0))) {
                min = extractdigits.get(0);
                minEqualTo = OperatorConvertor.isEqual(subStrings.get(0));
                range.setMinRange(min, minEqualTo);
                int minDecimalplace = NumberProcessor.countDecimalPlace(extractdigitFromString.get(0));
                range.setMinDecimalPlace(minDecimalplace);



            }
            else {
                max = extractdigits.get(0);
                maxEqualTo = OperatorConvertor.isEqual(subStrings.get(0));
                range.setMaxRange(max, maxEqualTo);
                int maxDecimalplace = NumberProcessor.countDecimalPlace(extractdigitFromString.get(0));
                range.setMaxDecimalPlace(maxDecimalplace);
            }
        }

        return range;
    }



    public static void main(String[] args) {
//        MyRange myRange =MyRangeGenerator("  not less than 0.7 not more than 0.9 " );
//        MyRange myRange =MyRangeGenerator("  not less than 0.7 " );
//        MyRange myRange =MyRangeGenerator(" not more than 0.9" );
        MyRange myRange =MyRangeGenerator("within 100" );
        System.out.println(myRange.getMaxDecimalPlace());
        System.out.println(customStringFormatter.getAccecptableRangeString(myRange));
        try {
            List<Double> result = setOOS(myRange);
            for(Double d:result){
                System.out.println(d);
            }
        } catch (RangeException e) {
            e.printStackTrace();
        }

    }

}
