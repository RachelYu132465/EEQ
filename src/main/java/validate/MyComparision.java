//package validate;
//
//import common.StringProcessor;
//
//import java.util.Arrays;
//import java.util.regex.Matcher;
//import java.util.regex.Pattern;
//
//import static common.NumberProcessor.countDecimalPlace;
//
//
//public class MyComparision {
//
//
//
//    class mycomparison{
//       boolean static compare(NUmber a1, myRange a2){
//            if(a1 > a2.get()){
//                //
//            }
//        }
//    }
//
//
//
//
//
//
//    public boolean check() {
//
//        if ((StringProcessor.ifContainStrings(specification, targetInBetween)
//                || (StringProcessor.ifContainStrings(specification, targetBig) &&
//                StringProcessor.ifContainStrings(specification, targetSmall)))
//                &&
//
//                !(StringProcessor.ifContain("Standard", specification))) {
//            this.targetSmall = true;
//            this.tragetBig = true;
//
//
//        } else if (StringProcessor.ifContainStrings(specification, targetBig)) {
//            this.targetSmall = false;
//            this.tragetBig = true;
//
//        } else if (StringProcessor.ifContainStrings(specification, targetSmall)) {
//            this.targetSmall = true;
//            this.tragetBig = false;
//
//        } else System.out.println("in class SpecificationProcessor, operator not set correctly");
//        return this;
//
//    }
//
//
//    public MyComparision(String specification) {
//        this.targetSmall = false;
//        this.tragetBig = false;
//        this.specification = specification;
//        setOperator(specification);
//        saveNum(specification);
//        setOOSNum();
////        setNumberInSpec(specification);
//    }
//
//    public static void main(String[] args) {
////        MyComparision m = new MyComparision("The  Recovery is not less than 97.5% and not  more than 102.5%.");
//        MyComparision m = new MyComparision(" Not more than 3.0% is retained on No.80 U.S. standard sieve.");
//
//        System.out.println(m.getAccecptableRangeString());
//        m.setOOSNum();
//        for (double d : m.getOOSNum())
//            System.out.println(d);
////        System.out.println(m.decimalPlace("3"));
////        m.setSpecNum("Specification:  Recovery 90.0% ~ 110.0%");
//    }
//
//
//    //1. 將 spec的數字取出，只取字串前面數字
//    public MyComparision saveNum(String spec) {
//
//        if (targetSmall && tragetBig){
//            saveFirstTwoNum(spec);
//        }
//        else if (targetSmall ^ tragetBig){
//
//            saveFirstNum(spec);
//        } else {
//            this.num = new double[2];
//            this.numS = new String[2];
//        }
//        return this;
//    }
//
//    //case 1:in bwtween
//    public MyComparision saveFirstTwoNum(String spec) {
//        Pattern regex = Pattern.compile("(\\d+(?:\\.\\d+)?)");
//        Matcher matcher = regex.matcher(spec);
////        double max;
////        double min;
////        Double[] num;
////        String maxS;
////        String minS;
////        List<Double> temp = new ArrayList<Double>();
//        double[] arr = new double[2];
//        String[] arrS = new String[2];
//        for (int i = 0; matcher.find() && i < 2; i++) {
//            String s = matcher.group();
//            arrS[i] = s;
//            arr[i] = Double.valueOf(s);
//        }
//        Arrays.sort(arr);
//        this.numS = arrS;
//        this.num = arr;
//        return this;
//    }
////case 2: max & min
//
//    public MyComparision saveFirstNum(String spec) {
//        Pattern regex = Pattern.compile("(\\d+(?:\\.\\d+)?)");
//        Matcher matcher = regex.matcher(spec);
////        double num;
////        String numS;
//        double[] arr = new double[2];
//        String[] arrS = new String[2];
//        for (int i = 0; matcher.find() && i < 1; i++) {
//            String s = matcher.group();
//            arr[0] = Double.valueOf(s);
//            arrS[0] = s;
//
//        }
//
//
//        this.numS = arrS;
//        this.num = arr;
//        return this;
//    }
//
//
//
//
//
//
//
//    public double dividByPowerOf10(double decimalPlace) {
//
//        double size = Double.valueOf(decimalPlace);
//        double one = (double) 1;
//        for (int i = 0; i < size; i++) {
//            one = one / 10;
//        }
//        return one;
//    }
//    //傳入10 => 回傳 1。傳入1 => 回傳 1。傳入數字位數 0.3 =>回傳0.1。傳入數字位數 0.03 =>回傳0.01。
//    public double OneInLastDigitPoint(String number) {
//       double numberD = Double.valueOf(number);
//        if(countDecimalPlace(number) >0){
//           return dividByPowerOf10(numberD);
//        }
//
//
//        else return 1;
//    }
//    //3.存入OOS數字
//    public MyComparision setOOSNum () {
//        OOSNum = new double[2];
//        // 有上下限
//        if (targetSmall && tragetBig){
//            double oneForMin = OneInLastDigitPoint(this.numS[0]);
//            this.OOSNum[0] = this.num[0] -oneForMin;
//            double oneForMax = OneInLastDigitPoint(this.numS[1]);
//            this.OOSNum[1] = this.num[1] +oneForMax;
//        }
//
//        else if (!targetSmall && tragetBig){
//            double oneForMin = OneInLastDigitPoint(this.numS[0]);
//            this.OOSNum[0] = this.num[0] -oneForMin;
//        }
//        else if (targetSmall && !tragetBig){
//            double oneForMax = OneInLastDigitPoint(this.numS[0]);
//            this.OOSNum[0] = this.num[0] +oneForMax;
//        }
//        else this.OOSNum[0] =-1;
//
//        return this;
//
//    }
//
//
//
//
//
//    public String getSpecification() {
//        return specification;
//
//    }
//
//    public void setSpecification(String specification) {
//        this.specification = specification;
//    }
//
//    public double[] getNum() {
//        return num;
//    }
//
//    public void setNum(double[] num) {
//        this.num = num;
//    }
//
//    public String[] getNumS() {
//        return numS;
//    }
//
//
//    public void setNumS(String[] numS) {
//        this.numS = numS;
//    }
//
//    public Boolean getTargetSmall() {
//        return targetSmall;
//    }
//
////    public void setMyOperator(Operator myOperator) {
////        this.myOperator = myOperator;
////    }
//
//
//    public void setTargetSmall(Boolean targetSmall) {
//        this.targetSmall = targetSmall;
//    }
//
//    public Boolean getTragetBig() {
//        return tragetBig;
//    }
//
//    public void setTragetBig(Boolean tragetBig) {
//        this.tragetBig = tragetBig;
//    }
//
//
//
////    public void setMyOperator(Operator myOperator) {
////        this.myOperator = myOperator;
////    }
//
////    public double getOOSNumber() {
////        return OOSNumber;
////
////    }
////
////    public void setOOSNumber(double OOSNumber) {
////        this.OOSNumber = OOSNumber;
////    }
//
//    public double[] getOOSNum() {
//        return OOSNum;
//    }
//
//    public void setOOSNum(double[] OOSNum) {
//        this.OOSNum = OOSNum;
//    }
//
//    public double setNumberInSpec(String specification) {
//
////       if(specification.contains("%")){
////           return Double.valueOf(specification.replaceAll("[^(0-9|.)]", ""));
////       }
//
//        return Double.valueOf(specification.replaceAll("[^(0-9|.)]", ""));
//    }
//
//
//    public double getNumberInSpec() {
//        return numberInSpec;
//    }
//
//    public void setNumberInSpec(double numberInSpec) {
//        this.numberInSpec = numberInSpec;
//    }
//
//    public String getAccepatbleOperator() {
//        return accepatbleOperator;
//    }
//
//
//    public String getUnAccepatbleOperator() {
//        return unAccepatbleOperator;
//
//    }
//
//    public void setAccepatbleOperator(String accepatbleOperator) {
//        this.accepatbleOperator = accepatbleOperator;
//    }
//
//    public void setUnAccepatbleOperator(String unAccepatbleOperator) {
//        this.unAccepatbleOperator = unAccepatbleOperator;
//    }
//
//    static String[] targetBig = {"not less than", "≧"};
//    static String[] targetSmall = {"not more than", "≦"};
//    static String[] targetInBetween = {"within", "and", "~", "-"};
//
//}
