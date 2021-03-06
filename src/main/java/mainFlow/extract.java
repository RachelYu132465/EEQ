package mainFlow;

import dataStructure.ValidGoal;
import msexcel.Excel;
import msexcel.ExcelCell;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import validate.excelFormulaValidator;

import java.util.*;
import java.util.stream.Stream;

import static msexcel.Excel.getR1C1Idx;
import static validate.excelFormulaProcessor.findSpecification;
import static validate.excelFormulaProcessor.getBlueCells;
import static validate.excelFormulaValidator.scanByTypingSpec;
import static validate.excelFormulaValidator.scanForSpecificationOption;


public class extract {

    public static   HashMap<String, ValidGoal> extractData(Excel excel) {

        HashMap<String, ValidGoal> TobeProcessed = new HashMap<String, ValidGoal>();

            //檢查是次否為需要驗證的試算表
            //if (mySheet.checkSheet(excel.getSheet())) {

            List<Cell> allFomulaCell = getBlueCells(excel.getWorkbook(),excel.getSheet(),"blue",CellType.FORMULA);

            //需求一
            for (Cell formulaCell : allFomulaCell) {
                System.out.println("sheet: "+excel.getSheet().getSheetName());

                String specification = findSpecification(excel, formulaCell.getRowIndex(), formulaCell.getColumnIndex());
                System.out.println("Formula cell:" + formulaCell.getAddress().formatAsString());
                System.out.println("specification:" + specification);
//               // String r1c1ForSpec =  scanForSpecificationOption();
                String r1c1ForSpec =  scanByTypingSpec();
////                String r1c1ForSpec =  "";
                if(!StringUtils.isBlank(r1c1ForSpec)){

                    specification = r1c1ForSpec;
                }


//                String specificationR1C1 = scanForSpecificationOption();
//
//                String specification = "";
//                if (specificationR1C1.trim().isBlank()) {
//                    specification = findSpecification(excel, formulaCell.getRowIndex(), formulaCell.getColumnIndex());
//                } else {
//                    specification = excel.getCellValue(specificationR1C1).toString().replace(SPECIFICATION + ":", "");
//                }

                HashSet<ExcelCell> inputCells = excelFormulaValidator.getInputCells(excel, formulaCell);
                ExcelCell inputCell = null;
                ValidGoal goal = new ValidGoal();
                goal.setOutput(new ExcelCell(excel,formulaCell));

//            goal.setOutput(new ExcelCell(getR1C1Idx(formulaCell),
//                    findFormulaForValidate(excel.getCellValue_OriginalFormula(formulaCell).toString()), formulaCell));

                for (ExcelCell c : inputCells) {
                    if (inputCell==null && !c.getCell().getCellType().equals(CellType.FORMULA)) {
                        inputCell = c;
                        c.setNote(excel,c.getCell());
                    }
                }
                if (inputCell != null) {
                    goal.setInput(inputCell);
                }
                goal.setMyRange(specification);
                goal.setAllInputs(inputCells);
                goal.setTitle(excel.getSheet().getSheetName());
                TobeProcessed.put(getR1C1Idx(formulaCell), goal);
            }


//            }
//        else break;

        return  TobeProcessed;
    }
}
