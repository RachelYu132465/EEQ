package msexcel;


import com.gembox.spreadsheet.*;
import org.apache.commons.lang.StringUtils;
import org.apache.commons.lang.math.NumberUtils;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;


import javax.swing.filechooser.FileSystemView;
import java.io.*;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.*;

public class Excel {

    public static void main(String[] args) {
        Excel n = createExcel(ExcelType.EXCEL_XLSX);
        n.createSheet("test");

        n.assignSheet(0).assignRow(0).assignCell(0);
        Cell c0 = n.getCurCell();

        CellStyle style = n.getWorkbook().createCellStyle();
//        CellStyle style = c0.getCellStyle();
        Font defaultFont = n.getWorkbook().createFont();
        defaultFont.setItalic(true);
        style.setFont(defaultFont);
        c0.setCellStyle(style);
        c0.setCellValue("AA");


        n.assignSheet(0).assignRow(0).assignCell(1);
        Cell c1 = n.getCurCell();
        c1.setCellValue("AA");
        CellStyle style1 = n.getWorkbook().createCellStyle();
        Font defaultFont1 = n.getWorkbook().createFont();
        defaultFont1.setItalic(false);
        style1.setFont(defaultFont1);
        c1.setCellStyle(style1);

        n.saveToFile( FileSystemView.getFileSystemView().getHomeDirectory().getPath() + File.separator + "testTina.xlsx");

    }

    public static final String CellAddrR1C1Rex = "^[A-Z][0-9]+";
    File file;
    Workbook curWb;
    Sheet curSheet;
    Row curRow;
    Cell curCell;
    String fileName;

    public File getFile() {
        return file;
    }

    public void setFile(File file) {
        this.file = file;
    }

    Font font;
    HSSFColor slxColor;
    XSSFColor xlsxColor;
    CellStyle cellStyle;
    ExcelType excelType;

    public Sheet getSheet() {
        return curSheet;
    }


    public Row getCurRow() {
        return curRow;
    }

    public enum ExcelType {

        EXCEL_XLS(".xls"), EXCEL_XLSX(".xlsx");

        private String value;

        private ExcelType(String type) {
            this.value = type;
        }

        public String getValue() {
            return value;
        }

        public void setValue(String type) {
            this.value = type;
        }
    }

    private Excel() {
    }

    public Excel(String fileName) throws IOException {
        this(new FileInputStream(new File(fileName)), fileName);
        this.fileName = fileName;
    }

    private Excel(InputStream in, String fileName) {
        try {
            if (fileName.toString().toLowerCase().endsWith(ExcelType.EXCEL_XLS.getValue())) {// Excel 2003
                curWb = new HSSFWorkbook(in);
                this.excelType = ExcelType.EXCEL_XLS;
            } else if (fileName.toString().toLowerCase().endsWith(ExcelType.EXCEL_XLSX.getValue())) {// Excel 2007/2010
                curWb = new XSSFWorkbook(in);
                this.excelType = ExcelType.EXCEL_XLSX;
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private Excel(ExcelType type, byte[] bytes) {
        try {
            this.excelType = type;
            ByteArrayInputStream in = new ByteArrayInputStream(bytes);
            switch (type) {
                case EXCEL_XLS:
                    curWb = new HSSFWorkbook(in);
                    break;
                case EXCEL_XLSX:
                    curWb = new XSSFWorkbook(in);
                    break;
                default:
                    break;
            }
        } catch (IOException e) {

        }
    }

    private Excel(ExcelType type) {
        this.excelType = type;
        switch (type) {
            case EXCEL_XLS:
                curWb = new HSSFWorkbook();
                break;
            case EXCEL_XLSX:
                curWb = new XSSFWorkbook();
                break;
            default:
                break;
        }
    }

    public String getExtName() {
        return excelType.value;
    }

    public ExcelType getExcelType() {
        return excelType;
    }

    public static Excel loadExcel(String uri) {
        return loadExcel(new File(uri));
    }

    public static Excel loadExcel(File file) {
        try {
//            return new Excel(new FileInputStream(file), file.getName());
            Excel excel = new Excel(new FileInputStream(file), file.getName());
            excel.setFile(file);
            excel.setFileName(file.getName());
            return excel;
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        return null;
    }

    public static Excel loadExcel(InputStream is, String name) {
        return new Excel(is, name);
    }

    public static Excel loadExcel(ExcelType type, byte[] bytes) {
        return new Excel(type, bytes);
    }

    public static Excel loadExcel(byte[] bytes) {
        return new Excel(ExcelType.EXCEL_XLSX, bytes);
    }

    public static Excel createExcel(ExcelType type) {
        return new Excel(type);
    }

    /*
     * sheet operation
     */

    /**
     * If the name is null,
     *
     * @param sheetname The name to set for the sheet.
     * @return New sheet will be created.
     */
    public Excel createSheet(String sheetname) {

        if (StringUtils.isNotEmpty(sheetname)) {
            curSheet = curWb.createSheet(sheetname);
            return this;
        }
        curSheet = curWb.createSheet();
        return this;
    }

    public Excel assignSheet(int index) {

        curSheet = curWb.getSheetAt(index);
//        if (curSheet == null) {
//            curSheet = curWb.createSheet();
//        }
        return this;
    }

    public Excel assignSheet(String name) {
        curSheet = curWb.getSheet(name);
        if (curSheet == null) {
            curSheet = curWb.createSheet(name);
        }

        return this;
    }

    //excel represent new line in this way
    public static boolean ifContainNewLine(String s) {
        char LF = (char) 0x0A;
        if (s.contains(String.valueOf(LF))) {

            return true;
        }
        return false;
    }

    public Sheet copySheet(int sheetidx) {
        Sheet s = curWb.cloneSheet(sheetidx);
        return s;
    }

    public String getSheetName(int index) {
        return curWb.getSheetName(index);
    }

    @SuppressWarnings("unchecked")
    public List<Sheet> getSheets() {
        Iterator<Sheet> i = curWb.iterator();
        List<Sheet> sheetList = new ArrayList<>();
        while (i.hasNext()) {
            sheetList.add(i.next());
        }
        return sheetList;
    }

    public List<String> getAllSheetsNames() {
        int sheetcount = curWb.getNumberOfSheets();
        List<String> sheetName = new ArrayList<>();
        for (int i = 0; i < sheetcount; i++) {
            String s = curWb.getSheetName(i);
            sheetName.add(s);
        }

        return sheetName;
    }

    public static List<String> getSheetsNameBykeywordsIgnoreCase(Excel excel, String... keywords) {
        List<Sheet> sheets = excel.getSheets();
        List<String> containKeywordsheets = new ArrayList<>();
        List<String> SheetsNames = excel.getAllSheetsNames();
        for (String s : SheetsNames) {
            for (String keyword : keywords) {
//                System.out.println(s.getSheetName());
                if (s.toLowerCase().contains(keyword.toLowerCase(Locale.ROOT))) {
                    containKeywordsheets.add(s);
//                    System.out.println(s.getSheetName());

                }
            }

        }
        return containKeywordsheets;
    }

    public static List<Sheet> getSheetsBykeywordsIgnoreCase(Excel excel, String... keywords) {
        List<Sheet> sheets = excel.getSheets();

        List<Sheet> Sheets = new ArrayList<>();
        for (Sheet s : sheets) {
            for (String keyword : keywords) {
                if (s.getSheetName().toLowerCase().contains(keyword.toLowerCase(Locale.ROOT))) {

                    Sheets.add(s);

                }
            }
        }


        return Sheets;
    }

    public int getSheetSize() {
        return this.getSheets().size();
    }

    public int getNumberOfSheets() {
        return curWb.getNumberOfSheets();
    }

    public Workbook getWorkbook() {
        return curWb;
    }


    /*
     * row operation
     */
    public Excel assignRow(int idx) {
        if (curSheet == null) {
            assignSheet(curWb.getNumberOfSheets() - 1);
        }
        this.curRow = curSheet.getRow(idx);
        if (curRow == null) {
            this.curRow = curSheet.createRow(idx);
        }
        this.curCell = null;
        return this;
    }

    public void removeSheet(int index) {
        int sheetCount = curWb.getNumberOfSheets();
        if (index < sheetCount) {
            curWb.removeSheetAt(index);
        }
    }

    public void removeRow(int index) {
        int lastRowNum = curSheet.getLastRowNum();
        if (index >= 0 && index < lastRowNum)
            curSheet.shiftRows(index + 1, lastRowNum, -1);
        // 将行号为rowIndex+1一直到行号为lastRowNum的单元格全部上移一行，以便删除rowIndex行
        if (index == lastRowNum) {
            Row removingRow = curSheet.getRow(index);
            if (removingRow != null)
                curSheet.removeRow(removingRow);
        }
    }

    public int getLastRowNum() {
        if (curSheet == null) {
            return 0;
        }
        return curSheet.getLastRowNum();
    }

    public boolean checkRowIsBlank(int idx) throws Exception {
        if (curSheet == null) {
            throw new Exception("current sheet is null!");
        }
        this.curRow = curSheet.getRow(idx);
        if (curRow != null) {
            for (int i = curRow.getFirstCellNum(); i < curRow.getLastCellNum(); i++) {
                Cell cell = curRow.getCell(i, MissingCellPolicy.RETURN_NULL_AND_BLANK);
                if (cell != null) {
                    if (StringUtils.isNotBlank(cell.toString())) {
                        return false;
                    }
                }
            }
        }
        return true;
    }

    /*
     * cell operation
     */
    public Excel getCell(int idx) {
        if (curCell == null) {
            this.assignCell(idx);
            if (curCell == null)
                this.curCell = curRow.createCell(idx);
        }
        this.curCell = this.curRow.getCell(idx);

        return this;
    }

    public CellStyle getCellStyle() {
        return this.curCell.getCellStyle();
    }

    public void setCellStyle(CellStyle style) {
        this.curCell.setCellStyle(style);
    }

    public Cell getCurCell() {
        return this.curCell;

    }

    public void setAutoFilter(int firstRow, int lastRow, int firstCol, int lastCol) {
        this.curSheet.setAutoFilter(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
    }

    /**
     * @param index
     * @return true if the Cell is empty or null
     */
    public boolean checkCellIsNull(int index) {
        if (curSheet != null && curRow != null) {
            return curRow.getCell(index) == null;
        }
        return false;
    }

    public Excel assignCell(int index) {
        return assignCell(index, false);
    }

    public Excel assignCell(int index, boolean toCreate) {
        if (curRow != null) {
            curCell = curRow.getCell(index,
                    toCreate ? MissingCellPolicy.CREATE_NULL_AS_BLANK : MissingCellPolicy.RETURN_NULL_AND_BLANK);
        } else {
            this.assignRow(this.getLastRowNum());
        }
        if (curCell == null) {
            curCell = curRow.createCell(index);
        }
        return this;
    }

    public int getLastCellNum() {
        if (curRow == null) {
            return 0;
        }
        return curRow.getLastCellNum();
    }

//    public Excel setCellValue(Object context, CellType type) {
//        this.curCell.setCellType(type);
//
//        switch (type) {
//        case NUMERIC:
//            this.curCell.setCellValue((double) context);
//            break;
//        case ERROR:
//
//            break;
//        case BOOLEAN:
//            this.curCell.setCellValue((boolean) context);
//            break;
//        default:
//            this.curCell.setCellValue("" + context);
//            break;
//        }
//        return this;
//    }

    public Excel setCellValue(double dou) {
        this.curCell.setCellValue(dou);
        return this;
    }

    public Excel setCellValue(String str) {
        this.curCell.setCellValue(str);
        return this;
    }

    public Excel setCellValue(String r1c1, String str) {
        this.getCell(r1c1).setCellValue(str);
        return this;
    }

    public Excel setCellValue(Date date) {
        this.curCell.setCellValue(date);
        return this;
    }

    public Object getCellValue() {
        return getCellValue(this.curCell);
    }

    public List<String> getCellsStringValue(int startRow, int endRow, int startcol, int endcol) {
        List<String> stringList = new ArrayList<>();
        for (int j = startRow; j < endRow + 1; j++) {
            this.assignRow(j);

            for (int k = startcol; k < endcol + 1; k++) {
                this.assignCell(k);
                String cellString = this.getAbsoluteStringCellValue();
                if (!cellString.isBlank()) {
//                    System.out.print ("R"+j+"C"+k+cellString);

                    stringList.add(cellString);

                }
            }
        }
        return stringList;
    }

    public static Object getCellValue_OriginalFormula(Cell cell) {
        switch (cell.getCellType()) {
            case NUMERIC:
                if (HSSFDateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue();
                } else {
                    String value = cell.getNumericCellValue() + "";
                    return new BigDecimal(value).toPlainString();
                }
            case STRING:
                return cell.getStringCellValue();
            case FORMULA:
                return cell.getCellFormula();
            case BOOLEAN:
                return cell.getBooleanCellValue();
            case ERROR:
                return cell.getErrorCellValue();
            case BLANK:
                return "";
            default:
                break;
        }
        return "";
    }

    public Object getCellValue(int rowIdx, int columnIdx) {
        Cell cell = this.getSheet().getRow(rowIdx).getCell(columnIdx);
        return getCellValue(cell);
    }

    public static Object getCellValue(Cell cell) {
        if (cell != null) {
            switch (cell.getCellType()) {
                case NUMERIC:
                    if (HSSFDateUtil.isCellDateFormatted(cell)) {
                        return cell.getDateCellValue();
                    } else {
                        return cell.getNumericCellValue();
                    }

                case STRING:
                    return cell.getStringCellValue();
                case FORMULA:
                    switch (cell.getCachedFormulaResultType()) {
                        case NUMERIC:
                            return cell.getNumericCellValue();
                        case STRING:
                            return cell.getRichStringCellValue();
                        default:
                            return cell.getCellFormula();
                    }
                case BOOLEAN:
                    return cell.getBooleanCellValue();
                case ERROR:
                    return cell.getErrorCellValue();
                case BLANK:
//                return this.curCell.getCellComment();
                    return "";
                default:
                    break;
            }
        }
        return "";
    }

    public String findNextCellStrValue() {
        int currentColumnIdx = this.getCurCell().getColumnIndex();
        return this.getCell(currentColumnIdx + 1).getAbsoluteStringCellValue();
    }

    public BigDecimal getNumericCellValue() {
        String _val = getStringCellValue();
        return StringUtils.isEmpty(_val) ? new BigDecimal(0) : new BigDecimal(_val);
    }

    public String getAbsoluteStringCellValue() {
        if (curCell != null) {
            switch (this.excelType.name()) {
                case "EXCEL_XLS":
                    curCell.setCellType(CellType.STRING);
                    break;
                case "EXCEL_XLSX":
                    curCell.setCellType(CellType.STRING);
                    break;
                default:
                    break;
            }
            return this.curCell.getStringCellValue();
        }
        return "";
    }

    public String getStringCellValue() {
        Object _val = getCellValue();
        if (NumberUtils.isNumber("" + _val)) {
            String[] rs = String.valueOf(_val).split("\\.");
            if (rs.length > 1 && rs[1].equals("0")) {
                return rs[0];
            } else {
                return "" + _val;
            }
        } else {
            return "" + _val;
        }

    }

    public Date getDateCellValue() {
        return this.curCell.getDateCellValue();
    }

    /*
     * cell style
     */

    public Excel createCellStyle(String sfmt) {

        short fmt = curWb.getCreationHelper().createDataFormat().getFormat(sfmt);
        if (cellStyle == null) {
            cellStyle = curWb.createCellStyle();
        }
        cellStyle.setDataFormat(fmt);

        return this;
    }

    public Excel applyCellStyle(String sfmt) {
        if (curCell != null) {
            if (cellStyle != null) {
                short fmt = curWb.getCreationHelper().createDataFormat().getFormat(sfmt);
                cellStyle.setDataFormat(fmt);
            } else {
                cellStyle = curWb.createCellStyle();
            }
            this.curCell.setCellStyle(cellStyle);
        }
        return this;
    }

    public Excel applyCellStyle() {
        if (curCell != null && cellStyle != null) {
            this.curCell.setCellStyle(cellStyle);
        }
        return this;
    }

    /*
     * other operation
     */
    public byte[] toArray() {
        ByteArrayOutputStream obs = new ByteArrayOutputStream();
        try {
            curWb.write(obs);
        } catch (IOException e) {

        }
        return obs.toByteArray();
    }

    public void outputFile(String fileName) {
        try {
            String excelFileName = String.format("%s%s", fileName, getExtName());
            FileOutputStream fileOut = new FileOutputStream(excelFileName);
            curWb.write(fileOut);
            curWb.close();
            fileOut.flush();
            fileOut.close();
        } catch (IOException e) {

        }
    }

    public int getCurSheetRowCnt() {
        if (curSheet != null) {
            return curSheet.getPhysicalNumberOfRows();
        }
        System.out.println("Excel.getCurSheetRowCnt() curSheet is null");
        return 0;
    }

    public void autoSizeColumns() {
        int numberOfSheets = curWb.getNumberOfSheets();
        for (int i = 0; i < numberOfSheets; i++) {
            Sheet sheet = curWb.getSheetAt(i);
            if (sheet.getPhysicalNumberOfRows() > 0) {
                Row row = sheet.getRow(sheet.getFirstRowNum());
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    int columnIndex = cell.getColumnIndex();
                    sheet.autoSizeColumn(columnIndex);
                }
            }
        }
    }


    public static int getNbOfMergedRegions(Sheet sheet, int row) {
        int count = 0;
        for (int i = 0; i < sheet.getNumMergedRegions(); ++i) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            if (range.getFirstRow() <= row && range.getLastRow() >= row)
                ++count;
        }
        return count;
    }

    public Excel turnToNoMergedCellExcel(Excel excel, int rowFrom, int rowTo, int colFrom, int colTo) {
        CellRangeAddress region = new CellRangeAddress(1, 1, 1, 4);
        curSheet.addMergedRegion(region);
        return excel;
    }

    @SuppressWarnings("unchecked")
    public Excel turnToNoMergedCellExcel(Excel excel) {

        for (int sheetCnt = 0; sheetCnt < excel.getNumberOfSheets(); ++sheetCnt) {
            Sheet sheet = excel.assignSheet(sheetCnt).getSheet();
            for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
                CellRangeAddress region = sheet.getMergedRegion(i); // Region of merged cells
                int colIndex = region.getFirstColumn(); // number of columns merged
                int rowNum = region.getFirstRow(); // number of rows merged
                Cell firstCellInMergedArea = sheet.getRow(rowNum).getCell(colIndex);

                Object cellValue = getCellValue(firstCellInMergedArea);

                for (Row row : sheet) {
                    for (Cell cell : row) {
                        if (region.isInRange(cell)) {
                            if (cellValue instanceof Date)
                                cell.setCellValue((Date) cellValue);
                            else if (cellValue instanceof Boolean)
                                cell.setCellValue((Boolean) cellValue);
                            else if (cellValue instanceof Double)
                                cell.setCellValue((Double) cellValue);
                            else
                                cell.setCellValue(cellValue.toString());
                        }
                    }
                }
                sheet.removeMergedRegion(i);
            }
        }
        return excel;
    }

    public void removeLastBlankRow(Sheet sheet){
        int lastRowIdx = sheet.getLastRowNum();
        Check1:
        for(int idx=lastRowIdx; idx>0;idx--){
            Row row=sheet.getRow(idx);
            for(int cellIdx =0; cellIdx<row.getLastCellNum(); cellIdx++){
                if(row.getCell(cellIdx)!=null && StringUtils.isNotBlank(getCellValueAsString(row.getCell(cellIdx)))){
                    break Check1;
                }
            }
            sheet.removeRow(row);
        }
    }

    public static boolean removeMergedArea(Sheet sheet, int rowStartIdxOfMerged, int rowEndIdxOfMerged) {
        boolean result = false;
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress region = sheet.getMergedRegion(i); // Region of merged cells
            int colIndex = region.getFirstColumn(); // number of columns merged
            int rowNum = region.getFirstRow(); // number of rows merged
//                if(rowIdxOfMerged == rowNum && colIndex==colIndex){
            if (rowStartIdxOfMerged <= rowNum && rowEndIdxOfMerged >= rowNum) {
                Cell firstCellInMergedArea = sheet.getRow(rowNum).getCell(colIndex);
                Object cellValue = getCellValue(firstCellInMergedArea);

                for (Row row : sheet) {
                    for (Cell cell : row) {
                        if (region.isInRange(cell)) {
                            if (cellValue instanceof Date)
                                cell.setCellValue((Date) cellValue);
                            else if (cellValue instanceof Boolean)
                                cell.setCellValue((Boolean) cellValue);
                            else if (cellValue instanceof Double)
                                cell.setCellValue((Double) cellValue);
                            else
                                cell.setCellValue(cellValue.toString());
                        }
                    }
                }
                sheet.removeMergedRegion(i);
                result = true;
            }

        }

        return result;
    }

    public static String getCellIndex(Cell cell) {
        return cell.getRowIndex() + "," + cell.getColumnIndex();
    }

    public Cell getCell(String r1c1) {
        CellReference cellReference = new CellReference(r1c1);
        Row row = this.curSheet.getRow(cellReference.getRow());
        if (row != null) {
            return row.getCell(cellReference.getCol());
        }

        return null;
    }

    public String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return null;
        } else if (cell.getCellType().equals(CellType.STRING)) {
//            logger.info("TEST: " + cell.toString());
            return cell.toString();
        } else if (cell.getCellType().equals(CellType.NUMERIC)) {
            DataFormatter formatter = new DataFormatter();
            String formattedValue = formatter.formatCellValue(cell);
            return formattedValue;
        } else if (cell.getCellType().equals(CellType.FORMULA)) {
            FormulaEvaluator evaluator = this.curWb.getCreationHelper().createFormulaEvaluator();
            DataFormatter formatter = new DataFormatter();
            String formattedValue = formatter.formatCellValue(cell, evaluator);
            return formattedValue;
        } else {
//            throw new InvalidCSVException(
//                    "Value stored in cell is invalid! Valid types are Numbers or Strings.");
            return "";
        }
    }

    public Object getCellValue(String r1c1) {
        return getCellValue(getCell(r1c1));
    }

    public void setCellValue(Cell cell, Object value) {
        if (value instanceof Number) {
            cell.setCellValue(Double.valueOf(value.toString()));
        } else if (value instanceof String) {
            cell.setCellValue((String) value);
        } else if (value instanceof Date) {
            cell.setCellValue((Date) value);
        } else if (value instanceof Boolean) {
            cell.setCellValue((Boolean) value);
        } else {
            System.out.println("Set cell fails:" +
                    cell.getRowIndex() + "," + cell.getColumnIndex() + value.toString());
        }
    }

    public Excel save() {
        try (FileOutputStream outputStream = new FileOutputStream(this.file.getPath())) {
            this.getWorkbook().write(outputStream);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return this;
    }

    public Excel saveToFile(String pathName) {
        try (FileOutputStream outputStream = new FileOutputStream(pathName)) {
            this.getWorkbook().write(outputStream);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return loadExcel(pathName);
    }

    //this method use api from third party
    //if range passed in is D11:D13 --> return ["D11","D12","D13"]
    public List<String> getCellsAddrByRange(String range) throws IOException {
        List<String> result = new ArrayList<>();
        SpreadsheetInfo.setLicense("FREE-LIMITED-KEY");

        SpreadsheetInfo.addFreeLimitReachedListener(eventArguments -> eventArguments.setFreeLimitReachedAction(FreeLimitReachedAction.CONTINUE_AS_TRIAL));
        ExcelFile workbook;
        if (this.getFile() != null) {
//            saveToFile(this.getFile().getAbsolutePath());
            workbook = ExcelFile.load(new FileInputStream(this.getFile()));
            if (workbook.getWorksheet(this.curSheet.getSheetName()) != null) {
                ExcelWorksheet worksheet = workbook.getWorksheet(this.curSheet.getSheetName());
                String cells[] = range.split(":");
                if (cells.length > 1) {
                    com.gembox.spreadsheet.CellRange cellRange = worksheet.getCells().getSubrange(cells[0], cells[1]);
                    CellRangeIterator ir = cellRange.getReadIterator();
                    while (ir.hasNext()) {
                        result.add(ir.next().getName());
                    }
                }
            }

        }

        return result;
    }

    public static String getR1C1Idx(Cell cell) {
        return cell.getAddress().formatAsString();
    }

    public String getR1C1Idx() {
        return getCurCell().getAddress().formatAsString();
    }

    public void setCellValue(String r1c1, Object value) {
        CellReference cellReference = new CellReference(r1c1);
        Row row = this.curSheet.getRow(cellReference.getRow());
        Cell cell = row.getCell(cellReference.getRow());
        if (cell != null) setCellValue(cell, value);
    }

    public String getfontColor() {
        // CellStyle styles= this.getCurCell().getCellStyle();
        // int fontIndex = styles.getFontIndex();
        // IndexedColors color =IndexedColors.fromInt(fontIndex);

        XSSFCellStyle cs = ((XSSFCellStyle) this.getCurCell().getCellStyle());
        XSSFFont font = cs.getFont();
        // Getting Font color
        XSSFColor color = font.getXSSFColor();
        if (color == null) {
            System.out.println("NULL");
            return "NULL";
        } else {
            // System.out.println("Font color : " + color.getARGBHex());
            return color.getARGBHex().toString();
        }


    }


    public String findFirstWord(String searchword, int startRow, int endRow, int startcol, int endcol) {

        for (int j = startRow; j < endRow; j++) {
            this.assignRow(j);

            for (int k = startcol; k < endcol; k++) {
                this.assignCell(k);
                String cellString = this.getAbsoluteStringCellValue();
                if (StringUtils.containsIgnoreCase(cellString, searchword.trim())) {
                    String s = getR1C1Idx(curCell);
                    return s;
                }
            }
        }
        return "";
    }

    /* loop excel
   Excel excel  =Excel.createExcel(EXCEL_XLSX);
            Excel templateExcel = Excel.loadExcel(template_absolute_path);

            int sheetSize =excel.getNumberOfSheets();
            for (int sheetCnt = 0; sheetCnt < sheetSize; ++sheetCnt) {
                excel.assignSheet(sheetCnt);
                Sheet sheet = excel.assignSheet(sheetCnt).getSheet();
                int rowSize = excel.getLastRowNum();
                for (int j = 0; j < rowSize; j++) {
                    excel.assignRow(j);
                    int colSize = excel.getLastCellNum();
                    for (int k = 0; k < colSize; k++) {
                        excel.assignCell(k);
                    }}}
    *
    * */
    public String findFirstWordInWb(String searchword, Excel excel) {


        int rowSize = excel.getLastRowNum();
        for (int j = 0; j < rowSize; j++) {
            this.assignRow(j);
            int colSize = excel.getLastCellNum();
            for (int k = 0; k < colSize; k++) {
                this.assignCell(k);
                String cellString = excel.getAbsoluteStringCellValue();
                if (StringUtils.containsIgnoreCase(cellString, searchword.trim())) {
                    String s = getR1C1Idx(curCell);
                    return s;
                }
            }

        }
        return "";
    }

    public String findfirstWordAtRight(int rowIdx, int colIdx) {
        int size = this.getLastCellNum();

        for (int k = colIdx + 1; k < size; k++) {
            this.assignRow(rowIdx);
            this.assignCell(k);
            if (StringUtils.isNotBlank(this.getAbsoluteStringCellValue()))
                return this.getCellValue().toString();


        }
        return "";
    }

    public String findfirstBlackWordAtRight(int rowIdx, int colIdx) {
        int size = this.getLastCellNum();
        for (int k = colIdx + 1; k < size; k++) {
            this.assignRow(rowIdx);
            this.assignCell(k);
            System.out.println("  rowidx:" + k);
            XSSFCellStyle cs = ((XSSFCellStyle) this.getCurCell().getCellStyle());
            XSSFFont font = cs.getFont();
            XSSFColor color = font.getXSSFColor();
            System.out.println("Font color : " + color.getARGBHex());
            if ((this.getCellValue().toString().isEmpty()))
                return "";

            else if (!(this.getCellValue().toString().isEmpty())) {
                if (this.getfontColor().equals("NULL") || this.getfontColor().equals("FF000000")) {
                    System.out.println("if loop" + k);
                    return this.getCellValue().toString();

                }
            }
        }
        return "";
    }

    public String findfirstBlackWordAtLeft(int rowIdx, int colIdx) {
        for (int k = colIdx - 1; k >= 0; k--) {
            this.assignRow(rowIdx);
            this.assignCell(k);
            System.out.println("  rowidx:" + k);
            XSSFCellStyle cs = ((XSSFCellStyle) this.getCurCell().getCellStyle());
            XSSFFont font = cs.getFont();
            XSSFColor color = font.getXSSFColor();
            System.out.println("Font color : " + color.getARGBHex());
            if ((this.getCellValue().toString().isEmpty()))
                return "";

            else if (!(this.getCellValue().toString().isEmpty())) {
                if (this.getfontColor().equals("NULL") || this.getfontColor().equals("FF000000")) {
                    System.out.println("if loop" + k);
                    return this.getCellValue().toString();

                }
            }
        }
        return "";
    }

    public String findFirstBlackWordAtTop(int rowIdx, int colIdx) {
        for (int k = rowIdx - 1; k >= 0; k--) {
            this.assignRow(k);
            this.assignCell(colIdx);

            System.out.println("  rowidx:" + k);
            // XSSFCellStyle cs = ((XSSFCellStyle)
            // this.getCurCell().getCellStyle());
            // XSSFFont font = cs.getFont();
            // XSSFColor color = font.getXSSFColor();
            // System.out.println("Font color : " + color.getARGBHex());
            if (!(this.getCellValue().toString().isEmpty())) {
                if (this.getfontColor().equals("NULL") || this.getfontColor().equals("FF000000")) {

                    // System.out.println("if loop" + k);
                    return this.getCellValue().toString();

                }
            }
        }
        return "";
    }

    public void reEvaluateFormula() {
        if (this.getExtName().equals(ExcelType.EXCEL_XLS.value))
            HSSFFormulaEvaluator.evaluateAllFormulaCells(this.getWorkbook());
        if (this.getExtName().equals(ExcelType.EXCEL_XLSX.value))
            XSSFFormulaEvaluator.evaluateAllFormulaCells(this.getWorkbook());
    }


    public String getFileName() {
        return fileName;
    }

    public void setFileName(String fileName) {
        this.fileName = fileName;
    }


    /*
   excel view setting function
   */
    public Sheet isDisplayGridlines(Sheet sheet, boolean DisplayGridlines) {
        //set display grid lines or not 是否顯示儲存格框線
        sheet.setDisplayGridlines(true);
        return sheet;
    }

    public Sheet setMargin(Sheet sheet, double TopMargin, double BottomMargin, double LeftMargin, double RightMargin) {
        //Set printing parameters
        sheet.setMargin(Sheet.TopMargin, (double) TopMargin);//  Margin (top)
        sheet.setMargin(Sheet.BottomMargin, (double) BottomMargin);//  Margin (bottom)
        sheet.setMargin(Sheet.LeftMargin, (double) LeftMargin);//  Margin (left)
        sheet.setMargin(Sheet.RightMargin, (double) RightMargin);//  Margin (right)
        return sheet;
    }

    /*
   excel view setting function block
   */


    /*
    excel print setting function
    */
    public Excel setPrintArea(Excel excel, int sheetidx, int start_column, int end_column, int start_row, int end_row) {
        excel.getWorkbook().setPrintArea(sheetidx, start_column, end_column, start_row, end_row);
        return excel;
    }

    public Sheet print_setting_on_sheet_PrintSetup(Sheet sheet, int setFitWidth, int setFitHeight, short papersize) {
        if (sheet != null) {
            PrintSetup ps = sheet.getPrintSetup();
            //一個page要印幾個sheet進去
            ps.setFitWidth((short) setFitWidth);
            ps.setFitHeight((short) setFitHeight);
            // e.g. A4 B4
            ps.setPaperSize(papersize);
        }
        return sheet;
    }

    public Sheet print_setting_on_sheet(Sheet sheet, boolean VerticallyCenter, boolean HorizontallyCenter, boolean PrintGridlines, boolean setFitToPage, boolean PrintRowAndColumnHeadings) {

        if (sheet != null) {

            //將sheet水平/垂直置中影印
            sheet.setVerticallyCenter(VerticallyCenter);
            sheet.setHorizontallyCenter(HorizontallyCenter);

            //set print grid lines or not 是否影印儲存格框線
            sheet.setPrintGridlines(PrintGridlines);

            sheet.setFitToPage(setFitToPage);

            //影印欄&列的名稱 i.e.ABC &123
            sheet.setPrintRowAndColumnHeadings(PrintRowAndColumnHeadings);
        }
        return sheet;
    }

//    public Sheet print_setting_on_workbook (Sheet newsheet, boolean VerticallyCenter,boolean HorizontallyCenter,boolean DisplayGridlines,boolean PrintGridlines, boolean PrintRowAndColumnHeadings ){
//
//
//    }
    /*
    end block: excel print setting function
    */
}// end of class
