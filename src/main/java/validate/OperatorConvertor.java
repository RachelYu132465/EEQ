package validate;

import common.StringProcessor;

public class OperatorConvertor {




    public static String getAcceptableSymbol(boolean lessOrGreater, boolean equalSign) {
//        boolean lessOrGreater, boolean equalSig
        String symbol_1 = "";
        String symbol_2 = "";
        if (lessOrGreater) symbol_1 = greaterSymbol;
        if (!lessOrGreater) symbol_1 = lessSymbol;

        if (equalSign) symbol_2 = equalSymbol;
        if (!equalSign) symbol_2 = "";

        return symbol_1 + symbol_2;

    }

    public static String getUnAcceptableSymbol(boolean lessOrGreater, boolean equalSign) {
        String symbol_1 = "";
        String symbol_2 = "";
        if (!lessOrGreater) symbol_1 = greaterSymbol;
        if (lessOrGreater) symbol_1 = lessSymbol;

        if (!equalSign) symbol_2 = equalSymbol;
        if (equalSign) symbol_2 = "";
        return symbol_1 + symbol_2;


    }

    public static boolean isGreater(String StringOperator) {
//        1.有大於字眼 且沒有‘不’ 字眼 2.有小於字眼且有‘不’字眼
        if ((StringProcessor.ifContainStrings(StringOperator, greaterStrs) &&
                (!StringProcessor.ifContainStrings(StringOperator, inverseStrs))) ||
                (StringProcessor.ifContainStrings(StringOperator, lessStrs) &&
                        (StringProcessor.ifContainStrings(StringOperator, inverseStrs))))
            return true;
        else if ((StringProcessor.ifContainStrings(StringOperator, lessStrs) &&
                (!StringProcessor.ifContainStrings(StringOperator, inverseStrs))) ||
                (StringProcessor.ifContainStrings(StringOperator, greaterStrs) &&
                        (StringProcessor.ifContainStrings(StringOperator, inverseStrs))))
            return false;
        else System.out.println("neither greater nor less,return false");
        return false;

    }

    public static boolean isEqual(String StringOperator) {
        if (StringProcessor.ifContainStrings(StringOperator, inverseStrs)||
         (StringProcessor.ifContainStrings(StringOperator, containEqualStrs))) return true;
        else return false;
    }

    static String lessSymbol = "<";
    static String greaterSymbol = ">";
    static String equalSymbol = "=";
    static String greater_and_equalSymbol ="≧";
    static String less_and_equalSymbol ="≦";

    static String NOT = "not";
    static String NO = "no";
    static String EXCEED = "exceed";
    static String LESS = "less";
    static String MORE = "more";
    static String GREATER = "Greater";

    static String NLT = "NLT";
    static String NMT = "NMT";
    static String WITHIN = "within";//尚未寫入邏輯，此為同時包含小於等於的一個字彙
    static String PASSES_THROUGH = "passes through";
    static String EQUAL = "equal";

    static String[] inverseStrs = {NOT, NO};
    static String[] lessStrs = {LESS,lessSymbol,NMT,WITHIN,less_and_equalSymbol};
    static String[] greaterStrs = {GREATER, MORE, EXCEED,greaterSymbol,NLT,greater_and_equalSymbol};
    static String[] containEqualStrs = {equalSymbol,NMT,WITHIN,NLT,less_and_equalSymbol,greater_and_equalSymbol};
    static String[] equalStrs = {PASSES_THROUGH,EQUAL};

}
