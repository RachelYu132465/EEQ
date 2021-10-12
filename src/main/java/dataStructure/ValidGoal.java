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


//    public static HashSet<String> getAllInputs(HashMap<String, ValidGoal> goals) {
//        HashSet<String> AllInputs = new HashSet<>();
//        //save all inputs from goals
//        for (Map.Entry<String, ValidGoal> goal : goals.entrySet()) {
//            HashSet<ExcelCell> input = goal.getValue().getAllInputs();
//            for (ExcelCell e : input) {
//                AllInputs.add(e.getR1c1());
//            }
//
//        }
//        return AllInputs;
//    }
//
    public static List <String>  getOutputCellAddress(HashMap<String, ValidGoal> goals) {
        HashSet<String> OutputCellAddress = new HashSet<>();

        for (Map.Entry<String, ValidGoal> goal : goals.entrySet()) {
            String outputR1C1 = goal.getKey();
            OutputCellAddress.add(outputR1C1);
        }
        List <String> list = new ArrayList<>(OutputCellAddress);

        return list;
    }

    public static String gettCellValueByR1C1 (String targetR1C1, HashMap<String, ValidGoal> goals) {
      //  HashSet<String> FormulaCells = new HashSet<>();
        // HashSet<String> AllIuputs =getAllInputs(goals) ;
      //  for (String C1R1 : getAllCells) {
            for (Map.Entry<String, ValidGoal> goal : goals.entrySet()) {
                if (goal.getKey().equals(targetR1C1)) return goal.getValue().getOutput().getValue();

                for (ExcelCell input_c : goal.getValue().getAllInputs()) {
                    if (input_c.getR1c1().equals(targetR1C1)) return input_c.getValue();
                }
            }


      //  }
        return "";
    }

    public static HashSet<String> getCellByFormulaType(HashSet<ExcelCell> AllInputs, boolean ifFormula) {
        HashSet<String> Formula = new HashSet<>();
        HashSet<String> non_Formula = new HashSet<>();
        HashSet<String> returnFormula;
        if (ifFormula) {
            for (ExcelCell input_c : AllInputs) {
                Cell formula_cells = input_c.getCell();

                if (formula_cells.getCellType().equals(CellType.FORMULA)) {
                    Formula.add(input_c.getR1c1());
                }
            }
            returnFormula = Formula;
        } else {
            for (ExcelCell input_c : AllInputs) {
                Cell formula_cells = input_c.getCell();

                if (!formula_cells.getCellType().equals(CellType.FORMULA)) {
                    non_Formula.add(input_c.getR1c1());
                }

            }
            returnFormula = non_Formula;
        }


        return returnFormula;
    }

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

//    public static void

    //    public static HashMap<String, ValidGoal> sortCell(HashMap<String, ValidGoal> goals) {
//
//        List<String> r1c1List = new ArrayList<String>();
//        HashMap<String, ValidGoal> sortedGoal = new HashMap<>();
//        for (Map.Entry<String, ValidGoal> goal : goals.entrySet()) {
//            r1c1List.add(goal.getKey().toString());
//        }
//        Collections.sort(r1c1List);
//
//        Collections.sort(r1c1List, new Comparator<String>() {
//            public int compare(String o1, String o2) {
//                return extractInt(o1) - extractInt(o2);
//            }
//
//            int extractInt(String s) {
//                String num = s.replaceAll("\\D", "");
//                // return 0 if no digits found
//                return num.isEmpty() ? 0 : Integer.parseInt(num);
//            }
//        });
//        for (String s : r1c1List)
//        System.out.println(s);
//
//
//        for (String s : r1c1List) {
//
//            inner:
//            for (Map.Entry<String, ValidGoal> goal : goals.entrySet()) {
//
//                if (s.equals(goal.getKey())) {
//                    sortedGoal.put(s, goal.getValue());
//
//                    break inner;
//                }
//            }
//        }
//
//        for (Map.Entry<String, ValidGoal> goal : sortedGoal.entrySet()){
//
//            HashSet<ExcelCell> allInputs  =goal.getValue().getAllInputs();
//            HashSet<ExcelCell> sortedInputs = sortCell(allInputs);
//            goal.getValue().setAllInputs(sortedInputs);
//        }
//return sortedGoal;
//    }
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

       HashSet<String> non_Formula = new HashSet<>(sortStringByNumericValue (Arrays.asList(s)));
       for (String ss : non_Formula)
        System.out.println(ss);
//             sortCell(extract.extractData(outputExcel));

    }
}
