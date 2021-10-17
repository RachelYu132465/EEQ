package msword;

import dataStructure.ValidGoal;
import msexcel.ExcelCell;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xwpf.usermodel.*;
import validate.MyRange;
import validate.customStringFormatter;

import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.*;

import static common.StringProcessor.sortStringByNumericValue;
import static dataStructure.ValidGoal.*;
import static msword.CustomTableText.*;
import static msword.CustomWordStyle.getTableTitleStyle;
import static msword.CustomWordStyle.types;
import static msword.ManageTable.*;
import static msword.ManageText.getParagraph;
import static msword.StrOnCustomWord.*;

public class CustomTableStyle {

    //using reflection to invoke static method:
    //addTitle_Style1(table.getRow(2).getCell(0).addParagraph(), INPUT_CELL);
    // <cellValue, cellWidth>
    public static void addToTable(types type, XWPFTable table, int rowIdx, HashMap<String, Integer>... values) {
        int numOfcells = 0;
        for (int realNumOfCell = 0; realNumOfCell < table.getRow(rowIdx).getTableCells().size(); realNumOfCell++) {
            if (table.getRow(rowIdx).getCell(realNumOfCell) != null) ++numOfcells;
        }
        if (numOfcells != values.length) {
            System.out.println("String arg size cannot match current table cells");
            return;
        }

        for (int cellIdx = 0; cellIdx < numOfcells; ++cellIdx) {
            HashMap<String, Integer> cellSetting = values[cellIdx];
            if (cellSetting != null) {
                XWPFTableCell cell = table.getRow(rowIdx).getCell(cellIdx);
                XWPFParagraph paragraph = getParagraph(cell);
                for (Map.Entry<String, Integer> en : cellSetting.entrySet()) {
                    if (en.getValue() != null) setCellW(cell, en.getValue());
                    try {
                        Method method = CustomWordStyle.class.getDeclaredMethod("add" + type.getValue() + "_Style", XWPFParagraph.class, String.class);
                        method.invoke(null, paragraph, en.getKey());
                    } catch (NoSuchMethodException e) {
                        e.printStackTrace();
                    } catch (InvocationTargetException e) {
                        e.printStackTrace();
                    } catch (IllegalAccessException e) {
                        e.printStackTrace();
                    }
                    break;
                }

            }


        }


    }

    //return real cell number (after merging table)
    static int addToTableByForce(XWPFTable table, int rowIdx, int cellIdx) {
        if (table.getRow(rowIdx) == null) {
            if (rowIdx > table.getNumberOfRows() - 1) {
                int numOfRowsToCreate = rowIdx + 1 - table.getNumberOfRows();
                for (int createdRow = 0; createdRow < numOfRowsToCreate; ++createdRow) {
                    table.createRow();
                }
            }
        }
        XWPFTableRow row = table.getRow(rowIdx);
        int numOfcells = 0;
        for (int realNumOfCell = 0; realNumOfCell < row.getTableCells().size(); realNumOfCell++) {
            if (table.getRow(rowIdx).getCell(realNumOfCell) != null) ++numOfcells;
        }
        //若實際上的cell(合併儲存格之後)，和傳進的值的數量相比，不足，則再新建cell
        if (row.getCell(cellIdx) == null) {
            if (cellIdx > numOfcells - 1) {
                int numOfCellsToCreate = cellIdx + 1 - numOfcells;
                for (int createdCells = 0; createdCells < numOfCellsToCreate; ++createdCells) {
                    row.createCell();++numOfcells;
                }
            }
        }
        return numOfcells;
    }

    public static void addToTable(types type, XWPFTable table, int rowIdx, String... values) {
        //因為資料表格要事後補充值，若已經超過該table列數量，另建列
        int numOfcells = addToTableByForce(table, rowIdx, values.length - 1);

        if (numOfcells < values.length) {
            System.out.println("Length of string arg more than num of current table cells");
            return;
        }
        for (int cellIdx = 0; cellIdx < values.length; ++cellIdx) {
            try {
                XWPFTableCell cell = table.getRow(rowIdx).getCell(cellIdx);
                XWPFParagraph paragraph = getParagraph(cell);
                Method method = CustomWordStyle.class.getDeclaredMethod("add" + type.getValue() + "_Style", XWPFParagraph.class, String.class);
                method.invoke(null, paragraph, values[cellIdx]);
            } catch (NoSuchMethodException e) {
                e.printStackTrace();
            } catch (InvocationTargetException e) {
                e.printStackTrace();
            } catch (IllegalAccessException e) {
                e.printStackTrace();
            }
        }


    }

    public static HashMap<String, Integer> getHashMap(String cellValue, Integer cellSize) {
        HashMap<String, Integer> result = new HashMap<String, Integer>();
        result.put(cellValue, cellSize);
        return result;
    }

    public static XWPFTable getCustomTable(XWPFDocument doc, int row, int col) {
        XWPFTable table = doc.createTable(row, col);
        table.setCellMargins(100, 50, 90, 5);
        return table;
    }




        //Excel.getCellValue_OriginalFormula(input).toString()
        public static void modifiedAppendToTable1(XWPFTable table, HashMap<String, ValidGoal> goals,CustomTableText CustomTableText){

            //加所有非公式的input儲存格 到column 1 & 2
    for (String s : CustomTableText.getNonFormulaList()) {

                addToTable(types.Content, table, table.getNumberOfRows(),
                        s
                        , GENERAL_FORMAT
                        , ""
                        , ""
                );
            }
            int rowIdx = 2;

List<String> InputFormulaList = CustomTableText.getInputFormulaList();
List<String> outputList =CustomTableText.getOutputList();
            for (String sortedFormulaCellAddress : CustomTableText.getFormulaList()) {

                    if (InputFormulaList.contains(sortedFormulaCellAddress)) {
                        addToTable(types.Content, table, ++rowIdx,
                                ""
                                , ""
                                , sortedFormulaCellAddress
                                , GENERAL_FORMAT);

                    }

                    else if (outputList.contains(sortedFormulaCellAddress)) {
                            //加入output儲存格 到column 3 & 4

                            for (Map.Entry<String, ValidGoal> entry : goals.entrySet()) {
                                if (sortedFormulaCellAddress.equals(entry.getKey())) {
                                    ValidGoal goal = entry.getValue();
                                    ExcelCell output = goal.getOutput();
                                    String format = "";
                                    int MaxDecimalPlace = 0;
                                    int MinDecimalPlace = 0;
                                    MyRange range = goal.getMyRange();
                                    if (range.hasMax()) MaxDecimalPlace = range.getMaxDecimalPlace();
                                    if (range.hasMin()) MinDecimalPlace = range.getMinDecimalPlace();
                                    if (MaxDecimalPlace == 0 && MinDecimalPlace == 0) format = INTEGER;
                                    else if (MaxDecimalPlace != 0)
                                        format = NUMERIC_DECIMAL + " " + range.getMaxDecimalPlace();
                                    else if (MinDecimalPlace != 0)
                                        format = NUMERIC_DECIMAL + " " + range.getMinDecimalPlace();
                                    addToTable(types.Content, table, ++rowIdx,
                                            ""
                                            , ""
                                            , output.getR1c1()
                                            , format);
                                    break;
                                }
                            }



                }

            }

        }



    public static void getTable_Style1(XWPFDocument doc, String worksheetName, HashMap<String, ValidGoal> goals,CustomTableText CustomTableText ) {
        XWPFTable table = getCustomTable(doc, 3, 4);
        table.getRow(1).setHeight(700);
        table.getRow(2).setHeight(600);
        setTableW(table);
        mergeCellHorizontally(table, 0, 0, 3);
        mergeCellHorizontally(table, 1, 0, 3);
        addToTable(types.Title, table, 0, WORKSHEET + Colon + worksheetName);
        addToTable(types.Title, table, 1, ITEM_NAME + Colon + CustomTableText.getAlltestitems());
        addToTable(types.Content, table, 2, INPUT_CELL, FORMAT_DESC, OUTPUT_CELL, FORMAT_DESC);

        modifiedAppendToTable1(table,goals,CustomTableText);


        endTable(table);
    }

    public static void appendToTable2(XWPFTable table, HashMap<String, ValidGoal> goals,CustomTableText CustomTableText) {
//for (String formatCellAddress : CustomTableText.getFormulaList()){
//    if
//}
        for (Map.Entry<String, ValidGoal> goal : goals.entrySet()) {
            XWPFTableRow row = table.createRow();

            addToTable(types.Content, table, table.getNumberOfRows() - 1,
                    goal.getValue().getOutput().getCell().getCellFormula(),
                    goal.getValue().getOutput().getR1c1(),
                    goal.getValue().getOutput().getNote()
            );
        }
    }

    public static void getTable_Style2(XWPFDocument doc, HashMap<String, ValidGoal> goals,CustomTableText CustomTableText) {
        XWPFRun title_run = doc.createParagraph().createRun();
        getTableTitleStyle(title_run);
        title_run.setText(FORMULA + " used");
        XWPFTable table = getCustomTable(doc, 1, 3);
        setTableW(table);//FORMULA, CELL_REF, DESCRIPTION
        addToTable(types.Content, table, 0,
                getHashMap(FORMULA, 5500),
                getHashMap(CELL_REF, 2000),
                getHashMap(DESCRIPTION, 5000));
        appendToTable2(table, goals,CustomTableText);
        endTable(table);
    }

    public static void appendToTable3(XWPFTable table, HashMap<String, ValidGoal> goals,CustomTableText CustomTableText) {
        for (Map.Entry<String, ValidGoal> goal : goals.entrySet()) {
            XWPFTableRow row = table.createRow();
            row.createCell();
            addToTable(types.Content, table, table.getNumberOfRows() - 1,
                    goal.getValue().getOutput().getR1c1(),
                    customStringFormatter.getAccecptableRangeString(goal.getValue().getMyRange()));

        }
    }

    public static void getTable_Style3(XWPFDocument doc, HashMap<String, ValidGoal> goals,CustomTableText CustomTableText) {
        XWPFRun title_run = doc.createParagraph().createRun();
        getTableTitleStyle(title_run);
        title_run.setText("Range validation");
        XWPFTable table = getCustomTable(doc, 2, 2);
        setTableW(table);
        mergeCellHorizontally(table, 0, 0, 1);
        addToTable(types.Content, table, 0, range_valid_txt);
        addToTable(types.Content, table, 1,
                getHashMap(CELL, 1000),
                getHashMap(ACCEPT_R, 6000));
        appendToTable3(table, goals, CustomTableText);
        endTable(table);
    }


    //印output
    public static void modifiedAppendToTable4_2 (XWPFTable table, HashMap<String, ValidGoal> goals,CustomTableText CustomTableText) {
        int rowIdx = 0;

        for (String s : CustomTableText.getNonFormulaList()) {
            String value = gettCellValueByR1C1(s, goals);

                addToTable(types.Content, table, ++rowIdx,
                        s
                        , value
                        , ""
                        , ""
                        , "", "", ""

                );


        }
        rowIdx = 0;

            for (String sortedFormulaCellAddress : CustomTableText.getFormulaList()) {

                    String value = gettCellValueByR1C1(sortedFormulaCellAddress, goals);

                    addToTable(types.Content, table, ++rowIdx,
                            "", "", sortedFormulaCellAddress, "", "", value);
                }
            }




        public static void getHeadStyle_tbl4 (XWPFTable table){
            table.getRow(0).setHeight(700);
            table.getRow(3).setHeight(600);
            table.getRow(2).setHeight(700);
            setTableW(table);
            mergeCellHorizontally(table, 0, 1, 4);
            mergeCellHorizontally(table, 0, 2, 3);
            mergeCellHorizontally(table, 1, 0, 6);
            mergeCellHorizontally(table, 2, 0, 6);
            mergeCellHorizontally(table, 3, 0, 6);
        }

        public static void modifiedAppendToTable4_3 (XWPFTable table, HashMap<String, ValidGoal> oneValidGoal, CustomTableText CustomTableText){

            table.createRow();
            int newRowIdx = table.getNumberOfRows() - 1;
            if (table.getRow(newRowIdx).getTableCells().size() > 6)
                mergeCellHorizontally(table, newRowIdx, 0, 6);
           String outputCellAddress ="";
            String changedValue = getInPutByGoalR1C1(outputCellAddress, oneValidGoal);

            for (Map.Entry<String, ValidGoal> goal : oneValidGoal.entrySet()){
                outputCellAddress = goal.getKey();
            }
                table.getRow(newRowIdx).getCell(0).setText("Range Validation: OOS " + oneValidGoal.get(outputCellAddress).getOutput().getNote());

            int FormulaIndex = newRowIdx;
            int nonFormulaIndex = newRowIdx;


            for (String s : CustomTableText.getNonFormulaList()) {
                String value = gettCellValueByR1C1(s, oneValidGoal);

                if (value.equals(changedValue)) {
                    addToTable(types.Title, table, ++nonFormulaIndex,
                            s
                            , value
                            , ""
                            , ""
                            , "", "", ""

                    );
                } else {
                    addToTable(types.Content, table, ++nonFormulaIndex,
                            s
                            , value
                            , ""
                            , ""
                            , "", "", ""

                    );
                }


            }
            for (String sortedFormulaCellAddress : CustomTableText.getFormulaList()) {
                String value = gettCellValueByR1C1(sortedFormulaCellAddress, oneValidGoal);

                addToTable(types.Content, table, ++FormulaIndex,
                        "", "", sortedFormulaCellAddress, "", "", value);
            }
        }



        public static void getTable_Style4 (XWPFDocument doc, HashMap < String, ValidGoal > goals, HashMap < String, ValidGoal > newGoals,
        int testCaseIdx, CustomTableText CustomTableText){
            XWPFTable headTable = getCustomTable(doc, 4, 7);
            getHeadStyle_tbl4(headTable);

            addToTable(types.Content, headTable, 0, "Case " + testCaseIdx, "To Test the " + CustomTableText.getAlltestitems() + "  percentage", "Test Case No.: " + CustomTableText.getAlltestitems() + "#" + testCaseIdx);
            addToTable(types.Content, headTable, 1, instruction_txt);
            addToTable(types.Content, headTable, 2, accep_criteria_txt);
            addToTable(types.Content, headTable, 3, "in spec (Case " + testCaseIdx + "-Attachment 1)");
            endTable(headTable);
            XWPFTable table = getCustomTable(doc, 1, 7);
            addToTable(types.Content, table, 0, INPUT_CELL, "Input value",
                    OUTPUT_CELL, "Output " + RESULT, "Calculated " + RESULT, "Expected " + RESULT, "Test " + RESULT + "(Pass/Fail)");


            //第一個in spec table
            modifiedAppendToTable4_2(table, goals,CustomTableText);
            int count = 1;

            List<String> newDataR1C1 = getOutputCellAddress(newGoals);
//         newlist = new ArrayList<>(newDataR1C1);
            newDataR1C1 = sortStringByNumericValue(newDataR1C1);


            //test validation的table
            for (String s : newDataR1C1) {

                for (Map.Entry<String, ValidGoal> goal : newGoals.entrySet()) {
                    if (s.equals(goal.getKey())) {
                        System.out.println("test validation = " + s);
                        //加所有公式的、非公式的儲存格
                        if (goal.getKey() != null && goal.getValue() != null && goal.getKey().split("_").length > 1)// (goal.getKey().split("_")[1]+1)
                            goal.getValue().getOutput().setNote(goal.getValue().getOutput().getNote() + " (Case " + testCaseIdx + "- Attachment " + ++count + ")");
//            appendToTable4_3(table, goal.getValue());
                        HashMap<String, ValidGoal> oneValidGoal = new HashMap<>();
                        oneValidGoal.put(goal.getValue().getOutput().getR1c1(), goal.getValue());
                        CustomTableText rangeValidation = new CustomTableText(oneValidGoal);


                        modifiedAppendToTable4_3(table, oneValidGoal,rangeValidation);
                    }
                }

            }
            endTable(table);
        }

    }
