package dataStructure;

import common.FileHandler;
import mainFlow.VBS;
import mainFlow.extract;

import msexcel.Excel;
import msexcel.ExcelCell;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import validate.MyRange;

import javax.swing.filechooser.FileSystemView;
import java.io.IOException;
import java.util.*;

import static mainFlow.ProduceWordFile.writeToWord_testCase;
import static mainFlow.VBS.execAllVBSFiles;
import static msexcel.customExcelStyle.setMyConditionalFormatting;
import static validate.customStringFormatter.MyRangeGenerator;
import static validate.excelFormulaValidator.getValidatedValues;

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

//    public ValidGoal(Cell input, Cell output,  HashSet<ExcelCell> allInputs) {
//        this(new ExcelCell(input), new ExcelCell(output), allInputs);
//
//    }

//    public ValidGoal(Cell input, Cell output, String specification, HashSet<ExcelCell> allInputs) {
//        this(input, output, allInputs);
//    }

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

    public void setMyRange(String specification) {
        this.myComparision = MyRangeGenerator(specification);

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



    public static final String proj_path = FileSystemView.getFileSystemView().getHomeDirectory().getPath() + "/";

    public static void main(String[] args) throws IOException, InterruptedException
//            ,RangeException
    {
//        String fileName =  "格式_RT30397_LABSpreadsheet數字(1).xlsx";
//        String fileName = "C- RT30358-LAB Spreadsheet-數字版.xlsx";
//        String fileName = "final_RT30358_LAB_Spreadsheet複製.xlsx";
//        String output = "final_RT30358_LAB_Spreadsheet_2.xlsx";
//        // String output =  "originSpec_RT30358_LAB_Spreadsheet.xlsx";
////        String fileName = "R000012383-LAB Spreadsheet數字.xlsx";
//        Excel outputExcel = Excel.loadExcel(proj_path + output);
//
//        outputExcel.assignSheet(0);
        String s[] ={"C20","D30","C30","C8"};

//       HashSet<String> non_Formula = new HashSet<>(sortStringByNumericValue (Arrays.asList(s)));
//       for (String ss : non_Formula)
//        System.out.println(ss);
//             sortCell(extract.extractData(outputExcel));

    }
}
