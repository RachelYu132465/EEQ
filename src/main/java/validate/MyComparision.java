package validate;

import common.StringProcessor;

import java.math.BigDecimal;
import java.util.Arrays;
import java.util.regex.Matcher;
import java.util.regex.Pattern;


public class MyComparision {

    private final String specification;
//    public double max;
//    public double min;
    public double[] num;
    public String[] numS;
    Boolean hasMax;
    Boolean hasMin;
    public String accepatbleOperator;
    public String unAccepatbleOperator;
    public double numberInSpec;
    public double OOSNumber;
    public Operator myOperator;
    public double[] OOSNum;

    public MyComparision(String specification) {

        this.specification = specification;
        setOperator(specification);
        saveNum(specification);
//        setNumberInSpec(specification);
    }

    public static void main(String[] args) {
//        MyComparision m = new MyComparision("The  Recovery is not less than 97.5% and not  more than 102.5%.");
        MyComparision m = new MyComparision(" Not more than 3.0% is retained on No.80 U.S. standard sieve.");

        System.out.println(m.getAccecptableRangeString());
      m.setOOSNum();
      for (double d :m.getOOSNum())
        System.out.println(d);
//        System.out.println(m.decimalPlace("3"));
//        m.setSpecNum("Specification:  Recovery 90.0% ~ 110.0%");
    }




    //1. 將 spec的數字取出，只取字串前面數字
    public MyComparision saveNum(String spec) {
        if (hasMax&&hasMin){
            saveFirstTwoNum(spec);
        }
        else if (hasMax ^ hasMin){
            saveFirstNum(spec);
        }

        else {
            this.num = new double[2];
            this.numS= new String[2];
        }
        return this;
    }
    //case 1:in bwtween
    public MyComparision saveFirstTwoNum(String spec) {
        Pattern regex = Pattern.compile("(\\d+(?:\\.\\d+)?)");
        Matcher matcher = regex.matcher(spec);
//        double max;
//        double min;
//        Double[] num;
//        String maxS;
//        String minS;
//        List<Double> temp = new ArrayList<Double>();
        double[] arr = new double[2];
        String[] arrS = new String[2];
        for (int i = 0; matcher.find()&&i < 2; i++) {
            String s =matcher.group();
                arrS[i] =s;
                        arr[i] = Double.valueOf(s);
        }
        Arrays.sort(arr);
        this.numS = arrS;
       this.num = arr;
        return this;
    }
//case 2: max & min

    public MyComparision saveFirstNum(String spec) {
        Pattern regex = Pattern.compile("(\\d+(?:\\.\\d+)?)");
        Matcher matcher = regex.matcher(spec);
//        double num;
//        String numS;
        double[] arr = new double[2];
        String[] arrS = new String[2];
        for (int i = 0; matcher.find()&&i < 1; i++) {
            String s =matcher.group();
            arr[0] = Double.valueOf(s);
            arrS[0] =s;

        }


        this.numS = arrS;
        this.num = arr;
        return this;
    }


    private MyComparision setOperator(String specification) {
//List<String> targetBig = new ArrayList<String>();
        String[] targetBig = {"not less than", "≧"};
        String[] targetSmall = {"not more than", "≦"};
        String[] targetInBetween = {"within", "and", "~"};
        if ((StringProcessor.ifContainStrings(specification, targetInBetween)
        ||  (StringProcessor.ifContainStrings(specification, targetBig)&&
                StringProcessor.ifContainStrings(specification, targetSmall)))
                &&
                !(StringProcessor.ifContain("Standard",specification)))
         {
            this.hasMax = true;
            this.hasMin = true;
        } else if (StringProcessor.ifContainStrings(specification, targetBig)) {
            this.hasMax = false;
            this.hasMin = true;

        } else if (StringProcessor.ifContainStrings(specification, targetSmall)) {
            this.hasMax = true;
            this.hasMin = false;

        } else System.out.println("in class SpecificationProcessor, operator not set correctly");
        return this;

    }
    //2.傳出Acceptable Range 字串
    public String getAccecptableRangeString() {

        if (hasMax&&hasMin){
            return ">=" +numS[0] +" & "+ "<="+numS[1];
        }
        else if (!hasMax && hasMin){
            return ">="+numS[0];
        }
        else if (hasMax && !hasMin){
            return "<="+numS[0];
        }
        else return  "";
    }
//3.存入OOS數字
    public int decimalPlace (String Number){
        BigDecimal a = new BigDecimal(Number);
        return a.scale();
    }
    public double oneInDecimalPlace (int decimalPlace) {
        int size = Integer.valueOf(decimalPlace);
        double one=(double)1;
        for (int i = 0; i < size; i++) {
            one=one/10;
        }
        return one;
    }

    public MyComparision setOOSNum () {
        OOSNum = new double[2];
        if (hasMax&&hasMin){
            double oneForMin = oneInDecimalPlace(decimalPlace(this.numS[0]));
            this.OOSNum[0] = this.num[0] -oneForMin;
            double oneForMax = oneInDecimalPlace(decimalPlace(this.numS[1]));
            this.OOSNum[1] = this.num[1] +oneForMax;
        }
        else if (!hasMax && hasMin){
            double oneForMin = oneInDecimalPlace(decimalPlace(this.numS[0]));
            this.OOSNum[0] = this.num[0] -oneForMin;
        }
        else if (hasMax && !hasMin){
            double oneForMax = oneInDecimalPlace(decimalPlace(this.numS[0]));
            this.OOSNum[0] = this.num[0] +oneForMax;
        }
        else this.OOSNum[0] =-1;
        return this;

    }
    public String getUnaccecptableOperator(MyComparision MyComparision) {

       if (hasMax==true){
           return ">";
       }
       else if (hasMax==true){
            return "<";
        }
       else return  "";
    }
    public Boolean getHasMax() {
        return hasMax;
    }

    public void setHasMax(Boolean hasMax) {
        this.hasMax = hasMax;
    }

    public Boolean getHasMin() {
        return hasMin;
    }

    public void setHasMin(Boolean hasMin) {
        this.hasMin = hasMin;
    }



    public void setMyOperator(Operator myOperator) {
        this.myOperator = myOperator;
    }

    public double getOOSNumber() {
        return OOSNumber;

    }

    public void setOOSNumber(double OOSNumber) {
        this.OOSNumber = OOSNumber;
    }

    public double[] getOOSNum() {
        return OOSNum;
    }

    public void setOOSNum(double[] OOSNum) {
        this.OOSNum = OOSNum;
    }

    public double setNumberInSpec(String specification) {

//       if(specification.contains("%")){
//           return Double.valueOf(specification.replaceAll("[^(0-9|.)]", ""));
//       }

        return Double.valueOf(specification.replaceAll("[^(0-9|.)]", ""));
    }


    public double getNumberInSpec() {
        return numberInSpec;
    }

    public void setNumberInSpec(double numberInSpec) {
        this.numberInSpec = numberInSpec;
    }

    public String getAccepatbleOperator() {
        return accepatbleOperator;
    }



    public String getUnAccepatbleOperator() {
        return unAccepatbleOperator;

    }

    public void setAccepatbleOperator(String accepatbleOperator) {
        this.accepatbleOperator = accepatbleOperator;
    }

    public void setUnAccepatbleOperator(String unAccepatbleOperator) {
        this.unAccepatbleOperator = unAccepatbleOperator;
    }


}
