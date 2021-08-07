package dataStructure;

import msexcel.ExcelCell;
import org.apache.poi.ss.usermodel.Cell;
import validate.Specification;

import java.util.HashSet;

public class ValidGoal {
    ExcelCell input;
    ExcelCell output;
//    double targetOutput;
    Specification specification;
    HashSet<ExcelCell> allInputs;

    public ValidGoal() {
        input = new ExcelCell();
        output = new ExcelCell();
        allInputs = new HashSet<ExcelCell>();
    }

    public ValidGoal(ExcelCell input, ExcelCell output, HashSet<ExcelCell> allInputs) {
        this.input = input;
        this.output = output;
        this.allInputs = allInputs;
    }

    public ValidGoal(Cell input, Cell output, Specification targetOutput,HashSet<ExcelCell> allInputs) {
        this.input = new ExcelCell(input);
        this.output = new ExcelCell(output);
        this.specification = targetOutput;
        this.allInputs = allInputs;
    }

    public ValidGoal(Cell input, Cell output, String specification,HashSet<ExcelCell> allInputs) {
        this.input = new ExcelCell(input);
        this.output = new ExcelCell(output);
        this.specification = new Specification(specification);
        this.allInputs = allInputs;
    }

    public HashSet<ExcelCell> getAllInputs() {
        return allInputs;
    }

    public void setAllInputs(HashSet<ExcelCell> allInputs) {
        this.allInputs = allInputs;
    }

    public ExcelCell getOutput() {
        return output;
    }

    public void setOutput(ExcelCell output) {
        this.output = output;
    }

    public Specification getSpecification() {
        return specification;
    }

    public void setSpecification(String specification) {
        Specification s =new Specification(specification);
        this.specification = s;
    }
//    public double getTargetOutput() {
//        return targetOutput;
//    }
//
//    public void setTargetOutput(double targetOutput) {
//        this.targetOutput = targetOutput;
//    }

    public ExcelCell getInput() {
        return input;
    }

    public void setInput(ExcelCell input) {
        this.input = input;
    }

    @Override
    public String toString() {
        return "input:" + input + System.getProperty("line.separator")
                + "inputs:" + allInputs + System.getProperty("line.separator")
                + "output:" + output + System.getProperty("line.separator");
//                + "target:" + targetOutput + System.getProperty("line.separator");
    }
}
