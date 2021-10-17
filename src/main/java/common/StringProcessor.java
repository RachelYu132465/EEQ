
package common;
import org.apache.commons.lang.StringUtils;

import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class StringProcessor {

    //    final static String spe ="`!@#$%^&*()_+':\"-=\\[\\]{};\\|,.<>/?~";
    final static String any_special_symbol = "`|!|@|#|$|%|^|&|*|(|)|_|+|'|:|\"|\\-|=|\\[|\\]|{|}|;|\\||,|.|<|>|/|?|~";
    final static String custom_special_symbol = "`|!|@|#|%|^|&|*|(|)|_|+|'|\"|\\-|=|\\[|\\]|{|}|;|\\||,|.|<|>|/|?|~";
    final static String contain_special_symbol = "[.*" + any_special_symbol + ".*]+";
    public final static String extract_digit_rex = "(\\d+(?:\\.\\d+)?)";


    public static List<String> sortStringByNumericValue(List<String> String) {
        Collections.sort(String, new Comparator<String>() {
            public int compare(String o1, String o2) {

                String o1StringPart = o1.replaceAll("\\d", "");
                String o2StringPart = o2.replaceAll("\\d", "");


                if(o1StringPart.equalsIgnoreCase(o2StringPart))
                {
                    return extractInt(o1) - extractInt(o2);
                }
                return o1.compareTo(o2);
            }

            int extractInt(String s) {
                String num = s.replaceAll("\\D", "");
                // return 0 if no digits found
                return num.isEmpty() ? 0 : Integer.parseInt(num);
            }
        });
        return String;
    }
    public static List<String> extractStringByPattern(String string,String pattern) {

        Pattern regex = Pattern.compile(pattern);
        Matcher matcher = regex.matcher(string);

        List<String> extractedStrings =new ArrayList<>();

        for (int i = 0; matcher.find() ; i++) {
            String s = matcher.group();
            extractedStrings.add(s);

        }
        return extractedStrings;
    }

    //IF(ISBLANK(C7),"",(C8-C9)*100/(C8-C7)) ---> ISBLANK(C7),"",(C8-C9)*100/(C8-C7)
    public static  String removeProrgrammingMethodFromMostOuter(String formula){
        //find the first occurence of (
        if(StringUtils.isNotBlank(formula)){
            int leftParentheses = formula.indexOf("(");
            int rightParentheses = formula.lastIndexOf(")");
            return formula.substring(leftParentheses +1 ,rightParentheses);
        }
        else return "";

    }

    //傳入array的時候，contain字彙是第二個參數!
    public static boolean ifContain(String toBeCompared, String CellValue) {
        if (StringUtils.isNotBlank(CellValue) && CellValue.toLowerCase().contains(toBeCompared.toLowerCase())) return true;
        else return false;
    }

    public static boolean ifContainStrings(String CellValue, String[] toBeCompared) {
        for  (String s: toBeCompared ) {
            if (StringUtils.isNotBlank(CellValue) && CellValue.toLowerCase().contains(s.toLowerCase())) {
                return true;
            }
        }
             return false;

    }
    public static boolean ifNotContainStrings(String CellValue ,String[] notContain) {
        for  (String s: notContain ) {
            if (StringUtils.isNotBlank(CellValue) && !CellValue.toLowerCase().contains(s.toLowerCase())) {

                return true;
            }
        }
        return false;

    }

    //check if the row =the selected cells contain keyword A[], and/or not contain keyword B[]
    public static boolean ifContainA_AND_NotContainB(String CellValue ,String[] A,String[] B) {

           if(ifContainStrings(CellValue,A) && ifNotContainStrings(CellValue,B))
                return true;

        return false;
    }
    public static boolean ifContainA_OR_NotContainB(String CellValue ,String[] A,String[] B) {

        if(ifContainStrings(CellValue,A) || ifNotContainStrings(CellValue,B))
            return true;

        return false;
    }

    public static boolean ifEqual(String toBeCompared, String CellValue) {
        if (StringUtils.isNotBlank(CellValue) && replaceSpecialSymbol(CellValue.toLowerCase(), "").equals(toBeCompared))
            return true;
        else return false;
    }
    public static boolean ifContainSpecialSymbol(String strToCheck){
        return Pattern.matches(contain_special_symbol, strToCheck);
    }

    public static String replaceSpecialSymbol(String strToCheck, String replacement){
        return  strToCheck.replaceAll("["+any_special_symbol+"]+", replacement);
    }
    public static String replaceCustomSymbol(String strToCheck, String replacement){
        return  strToCheck.replaceAll("["+custom_special_symbol+"]+", replacement);
    }

    /*"ifBlank(C3)"-->true
      "C3*17+O2"-->false */
    public static boolean ifContainMethod(String StrToChecl){
        return Pattern.matches("[a-zA-Z]{2,}.*",replaceSpecialSymbol(StrToChecl,""));
    }

    public static String CapitalCharToLowerCaseWithDelimitor(String strToCheck, String Delimitor){

        List<Integer> indexOfCapital =IndexOfMatches(strToCheck,"[A-Z]");
        String result =strToCheck;
        for(Integer i: indexOfCapital){
            String capitalStr = String.valueOf(strToCheck.charAt(i));
            result = result.replace(capitalStr, Delimitor+capitalStr.toLowerCase());
        }
        return  result.trim();
    }

public static List<Integer> IndexOfMatches(String StrToCheck, String regularExp){
    List<Integer> result = new ArrayList<>();
    Pattern pattern = Pattern.compile(regularExp);
    Matcher matcher = pattern.matcher(StrToCheck);
    while (matcher.find()){
        result.add(matcher.start());//this will give you index
    }
    return result;
}
//    public final static String micro_biotin_rex = "^([a-zA-Z])(\bbiotin\\b)$";
    public final static String micro_biotin_rex_pre = "([0-9.-]*[\\s]*[a-zA-Z0-9]+([\\s]+))+(biotin)";
    public final static String micro_biotin_rex_suf = "(biotin)(([\\s]+)[a-zA-Z0-9]+[0-9.-]*)+";
    public final static String micro_biotin_rex_ = "([a-zA-Z0-9.-]+([\\s]+))*(biotin)(([\\s]*)[a-zA-Z0-9.-]*)*";
    public final static String micro_biotin_rex_p = "(biotin)(([\\s]*)[^\\s]*)*";
    public final static String micro_biotin_rex_s = "([^\\s]+([\\s]*))+(biotin)";
    public static boolean isMatch (String input, String... rexs){
        input = input.trim();
//        List<String>  regularExp = new ArrayList<>();
        for (String rex  :rexs) {
            Pattern mypattern = Pattern.compile(rex, Pattern.CASE_INSENSITIVE);
            Matcher mymatcher= mypattern.matcher(input);
            if (mymatcher.matches()) return true;
        }
    return false;
    }

    public static void main(String... arg) {
//        System.out.println(replaceSpecialSymbol("C15"," "));
//        System.out.println(Pattern.matches(".*[A-Z]+.*","0.154"));
        String input =" Biotin  AVFF RF5G";
//        Pattern mypattern = Pattern.compile(micro_biotin_rex_pre, Pattern.CASE_INSENSITIVE);
//        Pattern mypattern2 = Pattern.compile(micro_biotin_rex_suf, Pattern.CASE_INSENSITIVE);
//        Matcher mymatcher= mypattern.matcher(input);
//        Matcher mymatcher2= mypattern2.matcher(input);
        System.out.println( isMatch(input,micro_biotin_rex_pre,micro_biotin_rex_suf));
//        System.out.println( isMatch(input,micro_biotin_rex_pre,micro_biotin_rex_suf,micro_biotin_rex_));
//        System.out.println( isMatch(input,micro_biotin_rex_p,micro_biotin_rex_s));

//        System.out.println(Pattern.matches(micro_biotin_rex_pre,input )|| Pattern.matches(micro_biotin_rex_suf, input));
//        System.out.println(replaceCustomSymbol("^RSQ(D15:D19,B15:B19)"," "));
//        System.out.println(CapitalCharToLowerCaseWithDelimitor("ActKind"," "));
//        System.out.println(ifContainMethod("C3*17+O2"));//false
//        System.out.println(ifContainMethod("ifBlank(C3)"));//true


    }
}
