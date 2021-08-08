package mainFlow;

import dataStructure.ValidGoal;
import msexcel.Excel;
import msexcel.ExcelCell;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import validate.excelFormulaValidator;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;

import static msexcel.Excel.getR1C1Idx;
import static test.ExcelForRu.proj_path;
import static validate.excelFormulaProcessor.*;
import static validate.excelFormulaValidator.scanForSpecificationOption;

public class extract {
    public static Object extractData(String FileName) {

        List<String> SpecificationList = new ArrayList<String>();
        HashMap<String, ValidGoal> TobeProcessed = new HashMap<String, ValidGoal>();
        Excel excel = Excel.loadExcel(proj_path + FileName);
        excel.assignSheet(0);

        List<Cell> allFomulaCell = findAllBlueFormulaCall(excel);

        //需求一
        for (Cell formulaCell : allFomulaCell) {
            System.out.println("Formula cell:" + formulaCell.getAddress().formatAsString());
            String specificationR1C1 = scanForSpecificationOption();
            String specification = "";
            if (specificationR1C1.trim().isBlank()){
                specification= findSpecification(excel, formulaCell.getRowIndex(), formulaCell.getColumnIndex());
            }else{
                specification =  excel.getCellValue(specificationR1C1).toString().replace(SPECIFICATION+":","");
            }
            HashSet<ExcelCell> inputCells = excelFormulaValidator.getInputCells(excel, formulaCell);
            ExcelCell inputCell = null;
            ValidGoal goal = new ValidGoal();
            goal.setOutput(new ExcelCell(formulaCell));
//            goal.setOutput(new ExcelCell(getR1C1Idx(formulaCell),
//                    findFormulaForValidate(excel.getCellValue_OriginalFormula(formulaCell).toString()), formulaCell));

            for (ExcelCell c : inputCells) {
                if (c.getCell().getCellType() != CellType.FORMULA) {
                    inputCell = c;
                    break;
                }
            }
            if (inputCell != null) {
                goal.setInput(inputCell);
            }
            goal.setMyComparision(specification);
            goal.setAllInputs(inputCells);
            SpecificationList.add(specification);
            TobeProcessed.put(getR1C1Idx(formulaCell), goal);
        }

//        excel.saveToFile();
        return new Object[]{SpecificationList, TobeProcessed};
    }
}
