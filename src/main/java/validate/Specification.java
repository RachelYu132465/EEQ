package validate;

import common.StringProcessor;

public class Specification {
    public String specification;
    public String accepatbleOperator;
    public String unAccepatbleOperator;
   public double numberInSpec;
    public double OOSNumber;

    public double getOOSNumber() {
        return OOSNumber;
    }

    public void setOOSNumber(double OOSNumber) {
        this.OOSNumber = OOSNumber;
    }

    public  double setNumberInSpec(String specification){

//       if(specification.contains("%")){
//           return Double.valueOf(specification.replaceAll("[^(0-9|.)]", ""));
//       }

       return Double.valueOf(specification.replaceAll("[^(0-9|.)]", ""));
   }

    private Specification setOperator(String specification) {

        if (StringProcessor.ifContain(specification, "not less than")
||      StringProcessor.ifContain(specification, "≧")
        ) {
            setAccepatbleOperator(">=");
            setUnAccepatbleOperator("<");

        }

        else if (StringProcessor.ifContain(specification, "not more than")
                ||      StringProcessor.ifContain(specification, "≦")) {
            setAccepatbleOperator("<=");
            setUnAccepatbleOperator(">");

        }
        else System.out.println("in class SpecificationProcessor, operator not set correctly");
            return this;

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

    public Specification(String specification) {
        this.specification = specification;
        setOperator(specification);
        setNumberInSpec(specification);
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
