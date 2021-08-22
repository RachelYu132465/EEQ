package dataStructure;

import msexcel.ExcelCell;
import org.apache.poi.ss.usermodel.Cell;
import validate.MyRange;

import java.util.HashSet;

public class ValidGoal {
    ExcelCell input;
    ExcelCell output;
    String title;

    //    double targetOutput;
    MyRange myComparision;
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

    public ValidGoal(Cell input, Cell output, MyRange targetOutput, HashSet<ExcelCell> allInputs) {
        this(new ExcelCell(input), new ExcelCell(output), allInputs);
        this.myComparision = targetOutput;
    }

    public ValidGoal(Cell input, Cell output, String specification, HashSet<ExcelCell> allInputs) {
        this(input, output, new MyRange(), allInputs);
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

    public MyRange getMyRange() {
        return myComparision;
    }

    public void setMyRange( ) {

    }

    public String getTitle() {
        return title;
    }

    public void setTitle(String title) {
        this.title = title;
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
                + "output:" + output + System.getProperty("line.separator")
                + "target:" + myComparision + System.getProperty("line.separator");
    }
}
