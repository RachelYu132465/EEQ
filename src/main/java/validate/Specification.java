package validate;

import common.StringProcessor;
enum Operator{
    targetSmall,targetBig,targetInBetween;

    Operator() {
    }
}
public class Specification {
    public String specification;
    public String accepatbleOperator;
    public String unAccepatbleOperator;
   public double numberInSpec;
    public double OOSNumber;
public Operator myOperator;
Boolean hasMax;
Boolean hasMin;
    public Operator getMyOperator() {
        return myOperator;
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

    public  double setNumberInSpec(String specification){

//       if(specification.contains("%")){
//           return Double.valueOf(specification.replaceAll("[^(0-9|.)]", ""));
//       }

       return Double.valueOf(specification.replaceAll("[^(0-9|.)]", ""));
   }

    private Specification setOperator(String specification) {
//List<String> targetBig = new ArrayList<String>();
        String[] targetBig ={"not less than" ,"≧"};
        String[] targetSmall ={"not more than","≦"};
        String[] targetInBetween ={"within","and","~"};
        if(StringProcessor.ifContainStrings(specification,targetInBetween)){

        }
        else if (StringProcessor.ifContainStrings(specification,targetBig))
        {
            setAccepatbleOperator(">=");
            setUnAccepatbleOperator("<");
setMyOperator(Operator.targetBig);
        }

        else if (StringProcessor.ifContainStrings(specification,targetSmall))
             {
            setAccepatbleOperator("<=");
            setUnAccepatbleOperator(">");
                 setMyOperator(Operator.targetSmall);
        }
        else if (StringProcessor.ifContainStrings(specification,targetSmall))
        {
            setAccepatbleOperator("<=");
            setUnAccepatbleOperator(">");
            setMyOperator(Operator.targetSmall);
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
