package msword;

import dataStructure.ValidGoal;
import msexcel.Excel;
import msexcel.ExcelCell;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;

import java.util.*;

import static common.StringProcessor.sortStringByNumericValue;

public class CustomTableText {
    String alltestitems;
    List<String> FormulaList;
    List<String> nonFormulaList;
    List<String> outputList;
    List<String> inputFormulaList;


    public CustomTableText(Excel excel, ValidGoal goal) {

        HashMap<String, ValidGoal> goals =new HashMap<>();

        goals.put(goal.getOutput().toString(),goal);
        new CustomTableText(excel,goals);
    }
    public CustomTableText(){}
    public CustomTableText(HashMap<String, ValidGoal> goals) {


        this.FormulaList = getFormulaCellAddress(goals);
        this.nonFormulaList = getNonFormulaCellAddress(goals);
        this.outputList = getOutputCellAddress(goals);

        this.inputFormulaList = new ArrayList<>();
        inputFormulaList.addAll(FormulaList);
        //刪除在formulalist中，output的address，只保留input欄位有formula的address
        inputFormulaList.removeAll(outputList);


    }
    public CustomTableText(Excel excel, HashMap<String, ValidGoal> goals) {

        setAlltestitems(excel);

        this.FormulaList = getFormulaCellAddress(goals);
        this.nonFormulaList = getNonFormulaCellAddress(goals);
        this.outputList = getOutputCellAddress(goals);

        this.inputFormulaList = new ArrayList<>();
        inputFormulaList.addAll(FormulaList);
        //刪除在formulalist中，output的address，只保留input欄位有formula的address
        inputFormulaList.removeAll(outputList);


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

    public String getAlltestitems() {
        return alltestitems;
    }

    public void setAlltestitems(Excel excel) {
        List<String> TestItemR1C1 = excel.findAllkeyWordR1C1InWb("Test Item", excel);

        for (String testitemR1C1 : TestItemR1C1) {
            String thisTestItem = excel.findfirstWordAtRight(excel.getCell(testitemR1C1).getRowIndex(), excel.getCell(testitemR1C1).getColumnIndex());
           if (!alltestitems.isEmpty()){alltestitems = alltestitems + " & ";}
            alltestitems = alltestitems + thisTestItem;
        }
        this.alltestitems = alltestitems;
    }

    public List<String> getFormulaList() {
        return FormulaList;
    }


    public List<String> getNonFormulaList() {
        return nonFormulaList;
    }


    public List<String> getOutputList() {
        return outputList;
    }


    public List<String> getInputFormulaList() {
        return inputFormulaList;
    }

    public static List<String> getOutputCellAddress(HashMap<String, ValidGoal> goals) {
        HashSet<String> OutputCellAddress = new HashSet<>();

        for (Map.Entry<String, ValidGoal> goal : goals.entrySet()) {
            String outputR1C1 = goal.getKey();
            OutputCellAddress.add(outputR1C1);
        }
        List<String> list = new ArrayList<>(OutputCellAddress);
        list = sortStringByNumericValue(list);
        return list;
    }

    public static String getInPutByGoalR1C1(String targetR1C1, HashMap<String, ValidGoal> goals) {

        for (Map.Entry<String, ValidGoal> goal : goals.entrySet()) {
            if (goal.getKey().equals(targetR1C1)) return goal.getValue().getInput().getR1c1();


        }


        //  }
        return "";
    }

    public static String gettCellValueByR1C1(String targetR1C1, HashMap<String, ValidGoal> goals) {
        //  HashSet<String> FormulaCells = new HashSet<>();
        // HashSet<String> AllIuputs =getAllInputs(goals) ;
        //  for (String C1R1 : getAllCells) {
        for (Map.Entry<String, ValidGoal> goal : goals.entrySet()) {
            if (goal.getKey().equals(targetR1C1)) return goal.getValue().getOutput().getValue();

            for (ExcelCell input_c : goal.getValue().getAllInputs()) {
                if (input_c.getR1c1().equals(targetR1C1)) return input_c.getValue();
            }
        }
        return "";
    }

        public static ExcelCell getCellByR1C1(String targetR1C1, HashMap<String, ValidGoal> goals) {
            //  HashSet<String> FormulaCells = new HashSet<>();
            // HashSet<String> AllIuputs =getAllInputs(goals) ;
            //  for (String C1R1 : getAllCells) {
            for (Map.Entry<String, ValidGoal> goal : goals.entrySet()) {
                if (goal.getKey().equals(targetR1C1)) return goal.getValue().getOutput();

                for (ExcelCell input_c : goal.getValue().getAllInputs()) {
                    if (input_c.getR1c1().equals(targetR1C1)) return input_c;
                }
            }

        //  }
        return null;
    }

    public static HashSet<String> getCellByFormulaType(HashSet<ExcelCell> AllInputs, boolean ifFormula) {
        HashSet<String> Formula = new HashSet<>();
        HashSet<String> non_Formula = new HashSet<>();
        HashSet<String> returnFormula;
        if (ifFormula) {
            for (ExcelCell input_c : AllInputs) {
                Cell formula_cells = input_c.getCell();

                if (formula_cells.getCellType() == CellType.FORMULA) {
                    System.out.println(formula_cells.getCellFormula());
                }

                else if (formula_cells.getCellType() == CellType.STRING) {
                    System.out.println(formula_cells.getStringCellValue());
                }
//               else if (formula_cells.getCellType() == CellType.FORMULA) {
//                    System.out.println(formula_cells.getCellFormula());
//                }
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


    public static List<String> getFormulaCellAddress(HashMap<String, ValidGoal> goals) {
        HashSet<String> Formula = new HashSet<>();

        for (Map.Entry<String, ValidGoal> entry : goals.entrySet()) {
            Formula.add(entry.getKey());
            HashSet<String> allInput = getCellByFormulaType(entry.getValue().getAllInputs(), true);
            Formula.addAll(allInput);

        }

        List<String> FormulaList = new ArrayList<>(Formula);
        FormulaList = sortStringByNumericValue(FormulaList);
//        Formula = new HashSet<>(temp);
        return FormulaList;
    }

    public static List<String> getNonFormulaCellAddress(HashMap<String, ValidGoal> goals) {
        HashSet<String> non_Formula = new HashSet<>();
        for (Map.Entry<String, ValidGoal> entry : goals.entrySet()) {
            non_Formula.addAll(getCellByFormulaType(entry.getValue().getAllInputs(), false));
        }

        List<String> nonFormulaList = new ArrayList<>(non_Formula);
        nonFormulaList = sortStringByNumericValue(nonFormulaList);
//        non_Formula = new HashSet<>(temp);
        return nonFormulaList;
    }

}
