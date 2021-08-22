package validate;

import java.math.BigDecimal;

public class MyRange {

    private Boolean maxEqualTo;
    private Boolean minEqualTo; //  false for can't equal, true for can equal;

    BigDecimal min;
    BigDecimal max;

    int minDecimalPlace;
    int maxDecimalPlace;


    //e.g. myRange.setRange(3).SetEqualTo(true)
    //e.g. new myRange(3,MyRange.max).setToLess().SetEqualTo(true) => <=3

    public static void main(String[] args) {
        MyRange range1 = new MyRange();
//        range1.setMaxRange(3., false);
//        range1.setMinRange(2., true);
        System.out.println(range1.toString());
//        System.out.println("getAcceptableSymbol" + range1.maxSymbol.getAcceptableSymbol());
//        System.out.println("getUnAcceptableSymbol" + range1.maxSymbol.getUnAcceptableSymbol());
//        System.out.println("getUnAcceptableSymbol" + range1.maxSymbol.getAccecptableRangeString(range1));
    }

    @Override
    public String toString() {
        return "MyRange{" +
                "maxEqualTo=" + maxEqualTo +
                ", minEqualTo=" + minEqualTo +

                ", min=" + min +
                ", max=" + max +
                '}';
    }

    public Boolean getMaxEqualTo() {
        return maxEqualTo;
    }

    public void setMaxEqualTo(Boolean maxEqualTo) {
        this.maxEqualTo = maxEqualTo;
    }

    public Boolean getMinEqualTo() {
        return minEqualTo;
    }

    public void setMinEqualTo(Boolean minEqualTo) {
        this.minEqualTo = minEqualTo;
    }

    private void setMin(BigDecimal bd) {
        this.min = bd;

    }

    public int getMinDecimalPlace() {
        return minDecimalPlace;
    }

    public void setMinDecimalPlace(int minDecimalPlace) {
        this.minDecimalPlace = minDecimalPlace;
    }

    public int getMaxDecimalPlace() {
        return maxDecimalPlace;
    }

    public void setMaxDecimalPlace(int maxDecimalPlace) {
        this.maxDecimalPlace = maxDecimalPlace;
    }

    public BigDecimal getMin() {
        return min;
    }

    public BigDecimal getMax() {
        return max;
    }

    public MyRange setMinRange(BigDecimal number, boolean hasEqualSign) {
//        BigDecimal bd = new BigDecimal(number);
        setMin(number);
        setMinEqualTo(hasEqualSign);


        return this;
    }

    public void setMax(BigDecimal bd) {
        this.max = bd;

    }

    public MyRange setMaxRange(BigDecimal number, boolean hasEqualSign) {
        //設定數值
//        BigDecimal bd = new BigDecimal(number);
        setMax(number);
        setMaxEqualTo(hasEqualSign);

        return this;
    }


    public MyRange() {

        this.min = default_min;
        this.max = default_max;

    }

    public boolean hasMax() {

        if (max.compareTo(default_max) == 0) {
            return false;
        }
        return true;
    }

    public boolean hasMin() {

        if (min.compareTo(default_min) == 0) {
            return false;
        }
        return true;
    }

    //    public MyRange setMinRange (String number){
//        BigDecimal bd = new BigDecimal(number);
//        return setMin(bd);
//    }
//    public MyRange setMaxRange (String number){
//        BigDecimal bd = new BigDecimal(number);
//        return setMax(bd);
//    }
    static BigDecimal default_max = new BigDecimal(1000000000.);
    static BigDecimal default_min = new BigDecimal(-1000000000.);
}
