package copyExcel;

import java.io.*;
import java.util.*;

import msexcel.Excel;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTMarker;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTTwoCellAnchor;

import javax.swing.filechooser.FileSystemView;

import static common.FileHandler.getFilesByKeywords;
import static msexcel.Excel.getSheetsBykeywordsIgnoreCase;

public class converter {

    @SuppressWarnings("unused")
    private static class FormulaInfo {
        private String sheetName;
        private Integer rowIndex;
        private Integer cellIndex;
        private String formula;

        private FormulaInfo(String sheetName, Integer rowIndex, Integer cellIndex, String formula) {
            this.sheetName = sheetName;
            this.rowIndex = rowIndex;
            this.cellIndex = cellIndex;
            this.formula = formula;
        }

        public String getSheetName() {
            return sheetName;
        }

        public void setSheetName(String sheetName) {
            this.sheetName = sheetName;
        }

        public Integer getRowIndex() {
            return rowIndex;
        }

        public void setRowIndex(Integer rowIndex) {
            this.rowIndex = rowIndex;
        }

        public Integer getCellIndex() {
            return cellIndex;
        }

        public void setCellIndex(Integer cellIndex) {
            this.cellIndex = cellIndex;
        }

        public String getFormula() {
            return formula;
        }

        public void setFormula(String formula) {
            this.formula = formula;
        }
    }

    static List<FormulaInfo> formulaInfoList = new ArrayList<FormulaInfo>();

    public static void refreshFormula(XSSFWorkbook workbook) {
        for (FormulaInfo formulaInfo : formulaInfoList) {
            workbook.getSheet(formulaInfo.getSheetName()).getRow(formulaInfo.getRowIndex())
                    .getCell(formulaInfo.getCellIndex()).setCellFormula(formulaInfo.getFormula());
        }
        formulaInfoList.removeAll(formulaInfoList);
    }

    public static XSSFWorkbook convertWorkbookHSSFToXSSF(HSSFWorkbook source) {
        XSSFWorkbook retVal = new XSSFWorkbook();

        for (int i = 0; i < source.getNumberOfSheets(); i++) {

            HSSFSheet hssfsheet = source.getSheetAt(i);
            XSSFSheet xssfSheet = retVal.createSheet(hssfsheet.getSheetName());

            copySheetSettings(xssfSheet, hssfsheet);
            copySheet(hssfsheet, xssfSheet);
            copyPictures(xssfSheet, hssfsheet);
        }

        refreshFormula(retVal);

        return retVal;
    }

    private static void copySheet(HSSFSheet source, XSSFSheet destination) {
        copySheet(source, destination, true);
    }

    /**
     * @param destination the sheet to create from the copy.
     * @param source      sheet to copy.
     * @param copyStyle   true copy the style.
     */
    private static void copySheet(HSSFSheet source, XSSFSheet destination, boolean copyStyle) {
        int maxColumnNum = 0;
        List<CellStyle> styleMap2 = (copyStyle) ? new ArrayList<CellStyle>() : null;
        for (int i = source.getFirstRowNum(); i <= source.getLastRowNum(); i++) {
            HSSFRow srcRow = source.getRow(i);
            XSSFRow destRow = destination.createRow(i);
            if (srcRow != null) {
                // copyRow(source, destination, srcRow, destRow, styleMap);
                copyRow(source, destination, srcRow, destRow, styleMap2);
                if (srcRow.getLastCellNum() > maxColumnNum) {
                    maxColumnNum = srcRow.getLastCellNum();
                }
            }
        }
        for (int i = 0; i <= maxColumnNum; i++) {
            destination.setColumnWidth(i, source.getColumnWidth(i));
        }
    }

    /**
     * @param srcSheet  the sheet to copy.
     * @param destSheet the sheet to create.
     * @param srcRow    the row to copy.
     * @param destRow   the row to create.
     * @param styleMap  -
     */
    private static void copyRow(HSSFSheet srcSheet, XSSFSheet destSheet, HSSFRow srcRow, XSSFRow destRow,
                                // Map<Integer, HSSFCellStyle> styleMap) {
                                List<CellStyle> styleMap) {
        // manage a list of merged zone in order to not insert two times a
        // merged zone
        Set<CellRangeAddressWrapper> mergedRegions = new TreeSet<CellRangeAddressWrapper>();
        short dh = srcSheet.getDefaultRowHeight();
        if (srcRow.getHeight() != dh) {
            destRow.setHeight(srcRow.getHeight());
        }
        // pour chaque row
        // for (int j = srcRow.getFirstCellNum(); j <= srcRow.getLastCellNum();
        // j++) {
        int j = srcRow.getFirstCellNum();
        if (j < 0) {
            j = 0;
        }
        for (; j <= srcRow.getLastCellNum(); j++) {
            HSSFCell oldCell = srcRow.getCell(j); // ancienne cell
            XSSFCell newCell = destRow.getCell(j); // new cell
            if (oldCell != null) {
                if (newCell == null) {
                    newCell = destRow.createCell(j);
                }
                // copy chaque cell
                copyCell(oldCell, newCell, styleMap);
                // copy les informations de fusion entre les cellules
                // System.out.println("row num: " + srcRow.getRowNum() +
                // " , col: " + (short)oldCell.getColumnIndex());
                CellRangeAddress mergedRegion = getMergedRegion(srcSheet, srcRow.getRowNum(),
                        (short) oldCell.getColumnIndex());

                if (mergedRegion != null) {
                    // System.out.println("Selected merged region: " +
                    // mergedRegion.toString());
                    CellRangeAddress newMergedRegion = new CellRangeAddress(mergedRegion.getFirstRow(),
                            mergedRegion.getLastRow(), mergedRegion.getFirstColumn(), mergedRegion.getLastColumn());
                    // System.out.println("New merged region: " +
                    // newMergedRegion.toString());
                    CellRangeAddressWrapper wrapper = new CellRangeAddressWrapper(newMergedRegion);
                    if (isNewMergedRegion(wrapper, mergedRegions)) {
                        mergedRegions.add(wrapper);
                        destSheet.addMergedRegion(wrapper.range);
                    }

                }
            }
        }

    }

    /**
     * Récupère les informations de fusion des cellules dans la sheet source
     * pour les appliquer à la sheet destination... Récupère toutes les zones
     * merged dans la sheet source et regarde pour chacune d'elle si elle se
     * trouve dans la current row que nous traitons. Si oui, retourne l'objet
     * CellRangeAddress.
     *
     * @param sheet   the sheet containing the data.
     * @param rowNum  the num of the row to copy.
     * @param cellNum the num of the cell to copy.
     * @return the CellRangeAddress created.
     */
    public static CellRangeAddress getMergedRegion(HSSFSheet sheet, int rowNum, short cellNum) {
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress merged = sheet.getMergedRegion(i);
            if (merged.isInRange(rowNum, cellNum)) {
                return merged;
            }
        }
        return null;
    }

    /**
     * Check that the merged region has been created in the destination sheet.
     *
     * @param newMergedRegion the merged region to copy or not in the destination sheet.
     * @param mergedRegions   the list containing all the merged region.
     * @return true if the merged region is already in the list or not.
     */
    public static boolean isNewMergedRegion(CellRangeAddressWrapper newMergedRegion,
                                            Set<CellRangeAddressWrapper> mergedRegions) {
        return !mergedRegions.contains(newMergedRegion);
    }

    private static void copyPictures(Sheet newSheet, Sheet sheet) {
        Drawing drawingOld = sheet.createDrawingPatriarch();
        Drawing drawingNew = newSheet.createDrawingPatriarch();
        CreationHelper helper = newSheet.getWorkbook().getCreationHelper();

        // if (drawingNew instanceof HSSFPatriarch) {
        if (drawingOld instanceof HSSFPatriarch) {
            List<HSSFShape> shapes = ((HSSFPatriarch) drawingOld).getChildren();
            for (int i = 0; i < shapes.size(); i++) {
                System.out.println(shapes.size());
                if (shapes.get(i) instanceof HSSFPicture) {
                    HSSFPicture pic = (HSSFPicture) shapes.get(i);
                    HSSFPictureData picdata = pic.getPictureData();
                    int pictureIndex = newSheet.getWorkbook().addPicture(picdata.getData(), picdata.getFormat());
                    ClientAnchor anchor = null;
                    if (pic.getAnchor() != null) {
                        anchor = helper.createClientAnchor();
                        anchor.setDx1(((HSSFClientAnchor) pic.getAnchor()).getDx1());
                        anchor.setDx2(((HSSFClientAnchor) pic.getAnchor()).getDx2());
                        anchor.setDy1(((HSSFClientAnchor) pic.getAnchor()).getDy1());
                        anchor.setDy2(((HSSFClientAnchor) pic.getAnchor()).getDy2());
                        anchor.setCol1(((HSSFClientAnchor) pic.getAnchor()).getCol1());
                        anchor.setCol2(((HSSFClientAnchor) pic.getAnchor()).getCol2());
                        anchor.setRow1(((HSSFClientAnchor) pic.getAnchor()).getRow1());
                        anchor.setRow2(((HSSFClientAnchor) pic.getAnchor()).getRow2());
                        anchor.setAnchorType(((HSSFClientAnchor) pic.getAnchor()).getAnchorType());
                    }
                    drawingNew.createPicture(anchor, pictureIndex);
                }
            }
        } else {
            if (drawingNew instanceof XSSFDrawing) {
                List<XSSFShape> shapes = ((XSSFDrawing) drawingOld).getShapes();
                for (int i = 0; i < shapes.size(); i++) {
                    if (shapes.get(i) instanceof XSSFPicture) {
                        XSSFPicture pic = (XSSFPicture) shapes.get(i);
                        XSSFPictureData picdata = pic.getPictureData();
                        int pictureIndex = newSheet.getWorkbook().addPicture(picdata.getData(),
                                picdata.getPictureType());
                        XSSFClientAnchor anchor = null;
                        CTTwoCellAnchor oldAnchor = ((XSSFDrawing) drawingOld).getCTDrawing().getTwoCellAnchorArray(i);
                        if (oldAnchor != null) {
                            anchor = (XSSFClientAnchor) helper.createClientAnchor();
                            CTMarker markerFrom = oldAnchor.getFrom();
                            CTMarker markerTo = oldAnchor.getTo();
                            anchor.setDx1((int) markerFrom.getColOff());
                            anchor.setDx2((int) markerTo.getColOff());
                            anchor.setDy1((int) markerFrom.getRowOff());
                            anchor.setDy2((int) markerTo.getRowOff());
                            anchor.setCol1(markerFrom.getCol());
                            anchor.setCol2(markerTo.getCol());
                            anchor.setRow1(markerFrom.getRow());
                            anchor.setRow2(markerTo.getRow());
                        }
                        drawingNew.createPicture(anchor, pictureIndex);
                    }
                }
            }
        }
    }


    private static CellStyle getSameCellStyle(Cell oldCell, Cell newCell, List<CellStyle> styleList) {
        CellStyle styleToFind = oldCell.getCellStyle();
        CellStyle currentCellStyle = null;
        CellStyle returnCellStyle = null;
        Iterator<CellStyle> iterator = styleList.iterator();
        Font oldFont = null;
        Font newFont = null;
        while (iterator.hasNext() && returnCellStyle == null) {
            currentCellStyle = iterator.next();

            if (currentCellStyle.getAlignment() != styleToFind.getAlignment()) {
                continue;
            }
            if (currentCellStyle.getHidden() != styleToFind.getHidden()) {
                continue;
            }
            if (currentCellStyle.getLocked() != styleToFind.getLocked()) {
                continue;
            }
            if (currentCellStyle.getWrapText() != styleToFind.getWrapText()) {
                continue;
            }
            if (currentCellStyle.getBorderBottom() != styleToFind.getBorderBottom()) {
                continue;
            }
            if (currentCellStyle.getBorderLeft() != styleToFind.getBorderLeft()) {
                continue;
            }
            if (currentCellStyle.getBorderRight() != styleToFind.getBorderRight()) {
                continue;
            }
            if (currentCellStyle.getBorderTop() != styleToFind.getBorderTop()) {
                continue;
            }
            if (currentCellStyle.getBottomBorderColor() != styleToFind.getBottomBorderColor()) {
                continue;
            }
            if (currentCellStyle.getFillBackgroundColor() != styleToFind.getFillBackgroundColor()) {
                continue;
            }
            if (currentCellStyle.getFillForegroundColor() != styleToFind.getFillForegroundColor()) {
                continue;
            }
            if (currentCellStyle.getFillPattern() != styleToFind.getFillPattern()) {
                continue;
            }
            if (currentCellStyle.getIndention() != styleToFind.getIndention()) {
                continue;
            }
            if (currentCellStyle.getLeftBorderColor() != styleToFind.getLeftBorderColor()) {
                continue;
            }
            if (currentCellStyle.getRightBorderColor() != styleToFind.getRightBorderColor()) {
                continue;
            }
            if (currentCellStyle.getRotation() != styleToFind.getRotation()) {
                continue;
            }
            if (currentCellStyle.getTopBorderColor() != styleToFind.getTopBorderColor()) {
                continue;
            }
            if (currentCellStyle.getVerticalAlignment() != styleToFind.getVerticalAlignment()) {
                continue;
            }

            oldFont = oldCell.getSheet().getWorkbook().getFontAt(oldCell.getCellStyle().getFontIndex());
            newFont = newCell.getSheet().getWorkbook().getFontAt(currentCellStyle.getFontIndex());

            if (newFont.getBold() == oldFont.getBold()) {
                continue;

            }
            if (newFont.getColor() != oldFont.getColor()) {
                continue;
            }
            if (newFont.getFontHeight() != oldFont.getFontHeight()) {
                continue;
            }
            if (!newFont.getFontName().equals(oldFont.getFontName())) {
                continue;
            }
            if (newFont.getItalic() != oldFont.getItalic()) {
                continue;
            }
            if (newFont.getStrikeout() != oldFont.getStrikeout()) {
                continue;
            }
            if (newFont.getTypeOffset() != oldFont.getTypeOffset()) {
                continue;
            }
            if (newFont.getUnderline() != oldFont.getUnderline()) {
                continue;
            }
            if (newFont.getCharSet() != oldFont.getCharSet()) {
                continue;
            }
            if (oldCell.getCellStyle().getDataFormatString().equals(currentCellStyle.getDataFormatString())) {
                continue;
            }

            returnCellStyle = currentCellStyle;
        }
        return returnCellStyle;
    }

    public static void copySheetSettings(Sheet sheetToCopy, Sheet newSheet) {

        newSheet.setPrintRowAndColumnHeadings(sheetToCopy.isPrintRowAndColumnHeadings());
        newSheet.setAutobreaks(sheetToCopy.getAutobreaks());
        newSheet.setDefaultColumnWidth(sheetToCopy.getDefaultColumnWidth());
        newSheet.setDefaultRowHeight(sheetToCopy.getDefaultRowHeight());
        newSheet.setDefaultRowHeightInPoints(sheetToCopy.getDefaultRowHeightInPoints());
        newSheet.setDisplayGuts(sheetToCopy.getDisplayGuts());
        newSheet.setFitToPage(sheetToCopy.getFitToPage());

        newSheet.setForceFormulaRecalculation(sheetToCopy.getForceFormulaRecalculation());

        PrintSetup sheetToCopyPrintSetup = sheetToCopy.getPrintSetup();
        PrintSetup newSheetPrintSetup = newSheet.getPrintSetup();

        newSheetPrintSetup.setPaperSize(sheetToCopyPrintSetup.getPaperSize());
        newSheetPrintSetup.setScale(sheetToCopyPrintSetup.getScale());
        newSheetPrintSetup.setPageStart(sheetToCopyPrintSetup.getPageStart());
        newSheetPrintSetup.setFitWidth(sheetToCopyPrintSetup.getFitWidth());
        newSheetPrintSetup.setFitHeight(sheetToCopyPrintSetup.getFitHeight());
        newSheetPrintSetup.setLeftToRight(sheetToCopyPrintSetup.getLeftToRight());
        newSheetPrintSetup.setLandscape(sheetToCopyPrintSetup.getLandscape());
        newSheetPrintSetup.setValidSettings(sheetToCopyPrintSetup.getValidSettings());
        newSheetPrintSetup.setNoColor(sheetToCopyPrintSetup.getNoColor());
        newSheetPrintSetup.setDraft(sheetToCopyPrintSetup.getDraft());
        newSheetPrintSetup.setNotes(sheetToCopyPrintSetup.getNotes());
        newSheetPrintSetup.setNoOrientation(sheetToCopyPrintSetup.getNoOrientation());
        newSheetPrintSetup.setUsePage(sheetToCopyPrintSetup.getUsePage());
        newSheetPrintSetup.setHResolution(sheetToCopyPrintSetup.getHResolution());
        newSheetPrintSetup.setVResolution(sheetToCopyPrintSetup.getVResolution());
        newSheetPrintSetup.setHeaderMargin(sheetToCopyPrintSetup.getHeaderMargin());
        newSheetPrintSetup.setFooterMargin(sheetToCopyPrintSetup.getFooterMargin());
        newSheetPrintSetup.setCopies(sheetToCopyPrintSetup.getCopies());

        Header sheetToCopyHeader = sheetToCopy.getHeader();
        Header newSheetHeader = newSheet.getHeader();
        newSheetHeader.setCenter(sheetToCopyHeader.getCenter());
        newSheetHeader.setLeft(sheetToCopyHeader.getLeft());
        newSheetHeader.setRight(sheetToCopyHeader.getRight());

        Footer sheetToCopyFooter = sheetToCopy.getFooter();
        Footer newSheetFooter = newSheet.getFooter();
        newSheetFooter.setCenter(sheetToCopyFooter.getCenter());
        newSheetFooter.setLeft(sheetToCopyFooter.getLeft());
        newSheetFooter.setRight(sheetToCopyFooter.getRight());

        newSheet.setHorizontallyCenter(sheetToCopy.getHorizontallyCenter());
        newSheet.setMargin(Sheet.LeftMargin, sheetToCopy.getMargin(Sheet.LeftMargin));
        newSheet.setMargin(Sheet.RightMargin, sheetToCopy.getMargin(Sheet.RightMargin));
        newSheet.setMargin(Sheet.TopMargin, sheetToCopy.getMargin(Sheet.TopMargin));
        newSheet.setMargin(Sheet.BottomMargin, sheetToCopy.getMargin(Sheet.BottomMargin));

        newSheet.setPrintGridlines(sheetToCopy.isPrintGridlines());
        newSheet.setRowSumsBelow(sheetToCopy.getRowSumsBelow());
        newSheet.setRowSumsRight(sheetToCopy.getRowSumsRight());
        newSheet.setVerticallyCenter(sheetToCopy.getVerticallyCenter());
        newSheet.setDisplayFormulas(sheetToCopy.isDisplayFormulas());
        newSheet.setDisplayGridlines(sheetToCopy.isDisplayGridlines());
        newSheet.setDisplayRowColHeadings(sheetToCopy.isDisplayRowColHeadings());
        newSheet.setDisplayZeros(sheetToCopy.isDisplayZeros());
        newSheet.setPrintGridlines(sheetToCopy.isPrintGridlines());
        newSheet.setRightToLeft(sheetToCopy.isRightToLeft());
        newSheet.setZoom(85);
        copyPrintTitle(newSheet, sheetToCopy);
    }

    public static void copyHSSFRow(HSSFSheet destSheet, HSSFRow srcRow, HSSFRow destRow, Map<Integer, HSSFCellStyle> styleMap, List<CellRangeAddress> sheetMergedRegions, Set<String> mergedRegions) {
        destRow.setHeight(srcRow.getHeight());
        // reckoning delta rows
        int deltaRows = destRow.getRowNum() - srcRow.getRowNum();
        // pour chaque row
        for (int j = srcRow.getFirstCellNum(); j <= srcRow.getLastCellNum(); j++) {
            HSSFCell oldCell = srcRow.getCell(j);   // ancienne cell
            HSSFCell newCell = destRow.getCell(j);  // new cell
            if (oldCell != null) {
                if (newCell == null) {
                    newCell = destRow.createCell(j);
                }
                // copy chaque cell
                copyHSSFCell(oldCell, newCell, styleMap);
                // copy les informations de fusion entre les cellules
                CellRangeAddress mergedRegion = getMergedRegion(destSheet, srcRow.getRowNum(), (short) oldCell.getColumnIndex());

                if (mergedRegion != null) {
                    CellRangeAddress newMergedRegion = new CellRangeAddress(mergedRegion.getFirstRow(), mergedRegion.getLastRow(), mergedRegion.getFirstColumn(), mergedRegion.getLastColumn());
                    if (isNewMergedRegion(newMergedRegion, sheetMergedRegions)) {
                        mergedRegions.add(newMergedRegion.formatAsString());
                        destSheet.addMergedRegion(newMergedRegion);
                    }
                }
            }
        }

    }

    private static void copySheet(XSSFSheet source, XSSFSheet destination) {
        copySheet(source, destination, true);
    }

    private static void copySheet(XSSFSheet source, XSSFSheet destination, boolean copyStyle) {
        int maxColumnNum = 0;
        List<CellStyle> styleMap2 = (copyStyle) ? new ArrayList<CellStyle>() : null;
        for (int i = source.getFirstRowNum(); i <= source.getLastRowNum(); i++) {
            XSSFRow srcRow = source.getRow(i);
            XSSFRow destRow = destination.createRow(i);
            if (srcRow != null) {
                copyRow(source, destination, srcRow, destRow, styleMap2);
                if (srcRow.getLastCellNum() > maxColumnNum) {
                    maxColumnNum = srcRow.getLastCellNum();
                }
            }
        }
        for (int i = 0; i <= maxColumnNum; i++) {
            destination.setColumnWidth(i, source.getColumnWidth(i));
        }
    }

    /**
     * @param srcSheet  the sheet to copy.
     * @param destSheet the sheet to create.
     * @param srcRow    the row to copy.
     * @param destRow   the row to create.
     * @param styleMap  -
     */
    private static void copyRow(XSSFSheet srcSheet, XSSFSheet destSheet, XSSFRow srcRow, XSSFRow destRow,
                                List<CellStyle> styleMap) {
        Set<CellRangeAddressWrapper> mergedRegions = new TreeSet<CellRangeAddressWrapper>();
        short dh = srcSheet.getDefaultRowHeight();
        if (srcRow.getHeight() != dh) {
            destRow.setHeight(srcRow.getHeight());
        }


        int j = srcRow.getFirstCellNum();
        if (j < 0) {
            j = 0;
        }
        for (; j <= srcRow.getLastCellNum(); j++) {
            XSSFCell oldCell = srcRow.getCell(j); // ancienne cell
            XSSFCell newCell = destRow.getCell(j); // new cell
            if (oldCell != null) {
                if (newCell == null) {
                    newCell = destRow.createCell(j);
                }

                copyCell(oldCell, newCell, styleMap);
                CellRangeAddress mergedRegion = getMergedRegion(srcSheet, srcRow.getRowNum(),
                        (short) oldCell.getColumnIndex());

                if (mergedRegion != null) {
                    CellRangeAddress newMergedRegion = new CellRangeAddress(mergedRegion.getFirstRow(),
                            mergedRegion.getLastRow(), mergedRegion.getFirstColumn(), mergedRegion.getLastColumn());
                    CellRangeAddressWrapper wrapper = new CellRangeAddressWrapper(newMergedRegion);
//modify: the CellRangeAddress of a merged region is contained in a set (i.e.mergedRegions), we should add merged region before check

                    if (isNewMergedRegion(wrapper, mergedRegions)) {
                        mergedRegions.add(wrapper);
                        destSheet.addMergedRegion(wrapper.range);
                    }
                }
            }
        }


    }

    public static CellRangeAddress getMergedRegion(XSSFSheet sheet, int rowNum, short cellNum) {
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress merged = sheet.getMergedRegion(i);
            if (merged.isInRange(rowNum, cellNum)) {
                return merged;
            }
        }
        return null;
    }


    public static void copyPrintTitle(Sheet newSheet, Sheet sheetToCopy) {
        int nbNames = sheetToCopy.getWorkbook().getNumberOfNames();
        Name name = null;
        String formula = null;

        String part1S = null;
        String part2S = null;
        String formS = null;
        String formF = null;
        String part1F = null;
        String part2F = null;
        int rowB = -1;
        int rowE = -1;
        int colB = -1;
        int colE = -1;

        for (int i = 0; i < nbNames; i++) {
            name = sheetToCopy.getWorkbook().getNameAt(i);
            if (name.getSheetIndex() == sheetToCopy.getWorkbook().getSheetIndex(sheetToCopy)) {
                if (name.getNameName().equals("Print_Titles") || name.getNameName().equals(XSSFName.BUILTIN_PRINT_TITLE)) {
                    formula = name.getRefersToFormula();
                    int indexComma = formula.indexOf(",");
                    if (indexComma == -1) {
                        indexComma = formula.indexOf(";");
                    }
                    String firstPart = null;
                    ;
                    String secondPart = null;
                    if (indexComma == -1) {
                        firstPart = formula;
                    } else {
                        firstPart = formula.substring(0, indexComma);
                        secondPart = formula.substring(indexComma + 1);
                    }

                    formF = firstPart.substring(firstPart.indexOf("!") + 1);
                    part1F = formF.substring(0, formF.indexOf(":"));
                    part2F = formF.substring(formF.indexOf(":") + 1);

                    if (secondPart != null) {
                        formS = secondPart.substring(secondPart.indexOf("!") + 1);
                        part1S = formS.substring(0, formS.indexOf(":"));
                        part2S = formS.substring(formS.indexOf(":") + 1);
                    }

                    rowB = -1;
                    rowE = -1;
                    colB = -1;
                    colE = -1;
                    String rowBs, rowEs, colBs, colEs;
                    if (part1F.lastIndexOf("$") != part1F.indexOf("$")) {
                        rowBs = part1F.substring(part1F.lastIndexOf("$") + 1, part1F.length());
                        rowEs = part2F.substring(part2F.lastIndexOf("$") + 1, part2F.length());
                        rowB = Integer.parseInt(rowBs);
                        rowE = Integer.parseInt(rowEs);
                        if (secondPart != null) {
                            colBs = part1S.substring(part1S.lastIndexOf("$") + 1, part1S.length());
                            colEs = part2S.substring(part2S.lastIndexOf("$") + 1, part2S.length());
                            colB = CellReference.convertColStringToIndex(colBs);// CExportExcelHelperPoi.convertColumnLetterToInt(colBs);
                            colE = CellReference.convertColStringToIndex(colEs);// CExportExcelHelperPoi.convertColumnLetterToInt(colEs);
                        }
                    } else {
                        colBs = part1F.substring(part1F.lastIndexOf("$") + 1, part1F.length());
                        colEs = part2F.substring(part2F.lastIndexOf("$") + 1, part2F.length());
                        colB = CellReference.convertColStringToIndex(colBs);// CExportExcelHelperPoi.convertColumnLetterToInt(colBs);
                        colE = CellReference.convertColStringToIndex(colEs);// CExportExcelHelperPoi.convertColumnLetterToInt(colEs);
                        if (secondPart != null) {
                            rowBs = part1S.substring(part1S.lastIndexOf("$") + 1, part1S.length());
                            rowEs = part2S.substring(part2S.lastIndexOf("$") + 1, part2S.length());
                            rowB = Integer.parseInt(rowBs);
                            rowE = Integer.parseInt(rowEs);
                        }
                    }

//                    newSheet.getWorkbook().setRepeatingRowsAndColumns(newSheet.getWorkbook().getSheetIndex(newSheet), colB, colE, rowB - 1, rowE - 1);
                }
            }
        }
    }


    public static void copyHSSFSheet(HSSFSheet newSheet, HSSFSheet sheet) {
        int maxColumnNum = 0;
        Map<Integer, HSSFCellStyle> styleMap = new HashMap<>();
        // manage a list of merged zone in order to not insert two times a merged zone
        Set<String> mergedRegions = new TreeSet<>();
        List<CellRangeAddress> sheetMergedRegions = sheet.getMergedRegions();
        for (int i = sheet.getFirstRowNum(); i <= sheet.getLastRowNum(); i++) {
            HSSFRow srcRow = sheet.getRow(i);
            HSSFRow destRow = newSheet.createRow(i);
            if (srcRow != null) {
                copyHSSFRow(newSheet, srcRow, destRow, styleMap, sheetMergedRegions, mergedRegions);
                if (srcRow.getLastCellNum() > maxColumnNum) {
                    maxColumnNum = srcRow.getLastCellNum();
                }
            }
        }
        for (int i = 0; i <= maxColumnNum; i++) {
            newSheet.setColumnWidth(i, sheet.getColumnWidth(i));
        }
    }


    private static boolean areAllTrue(boolean... values) {
        for (int i = 0; i < values.length; ++i) {
            if (values[i] != true) {
                return false;
            }
        }
        return true;
    }


    private static boolean isNewMergedRegion(CellRangeAddress newMergedRegion,
                                             Collection<CellRangeAddress> mergedRegions) {

        boolean isNew = true;

        // we want to check if newMergedRegion is contained inside our collection
        for (CellRangeAddress add : mergedRegions) {

            boolean r1 = (add.getFirstRow() == newMergedRegion.getFirstRow());
            boolean r2 = (add.getLastRow() == newMergedRegion.getLastRow());
            boolean c1 = (add.getFirstColumn() == newMergedRegion.getFirstColumn());
            boolean c2 = (add.getLastColumn() == newMergedRegion.getLastColumn());
            if (areAllTrue(r1, r2, c1, c2)) {
                isNew = false;
            }
        }

        return isNew;
    }

    public static void copyCell(Cell oldCell, Cell newCell, List<CellStyle> styleList) {
        if (styleList != null) {
            if (oldCell.getSheet().getWorkbook() == newCell.getSheet().getWorkbook()) {
                newCell.setCellStyle(oldCell.getCellStyle());
            } else {
                DataFormat newDataFormat = newCell.getSheet().getWorkbook().createDataFormat();

                CellStyle newCellStyle = getSameCellStyle(oldCell, newCell, styleList);
                if (newCellStyle == null) {
                    //Create a new cell style
                    Font oldFont = oldCell.getSheet().getWorkbook().getFontAt(oldCell.getCellStyle().getFontIndex());
                    //Find a existing font corresponding to avoid to create a new one
                    Font newFont = newCell.getSheet().getWorkbook().findFont(oldFont.getBold(), oldFont.getColor(), oldFont.getFontHeight(), oldFont.getFontName(), oldFont.getItalic(), oldFont.getStrikeout(), oldFont.getTypeOffset(), oldFont.getUnderline());
                    if (newFont == null) {
                        newFont = newCell.getSheet().getWorkbook().createFont();
                        newFont.setBold(oldFont.getBold());
                        newFont.setColor(oldFont.getColor());
                        newFont.setFontHeight(oldFont.getFontHeight());
                        newFont.setFontName(oldFont.getFontName());
                        newFont.setItalic(oldFont.getItalic());
                        newFont.setStrikeout(oldFont.getStrikeout());
                        newFont.setTypeOffset(oldFont.getTypeOffset());
                        newFont.setUnderline(oldFont.getUnderline());
                        newFont.setCharSet(oldFont.getCharSet());
                    }

                    short newFormat = newDataFormat.getFormat(oldCell.getCellStyle().getDataFormatString());
                    newCellStyle = newCell.getSheet().getWorkbook().createCellStyle();
                    newCellStyle.setFont(newFont);
                    newCellStyle.setDataFormat(newFormat);

                    newCellStyle.setAlignment(oldCell.getCellStyle().getAlignment());
                    newCellStyle.setHidden(oldCell.getCellStyle().getHidden());
                    newCellStyle.setLocked(oldCell.getCellStyle().getLocked());
                    newCellStyle.setWrapText(oldCell.getCellStyle().getWrapText());
                    newCellStyle.setBorderBottom(oldCell.getCellStyle().getBorderBottom());
                    newCellStyle.setBorderLeft(oldCell.getCellStyle().getBorderLeft());
                    newCellStyle.setBorderRight(oldCell.getCellStyle().getBorderRight());
                    newCellStyle.setBorderTop(oldCell.getCellStyle().getBorderTop());
                    newCellStyle.setBottomBorderColor(oldCell.getCellStyle().getBottomBorderColor());
//                    if(oldCell.getCellStyle().getFillBackgroundColorColor()!=null)  System.out.println(oldCell.getAddress().toString());
                    newCellStyle.setFillBackgroundColor(oldCell.getCellStyle().getFillBackgroundColor());

                    newCellStyle.setFillForegroundColor(oldCell.getCellStyle().getFillForegroundColor());
                    newCellStyle.setFillPattern(oldCell.getCellStyle().getFillPattern());
                    newCellStyle.setIndention(oldCell.getCellStyle().getIndention());
                    newCellStyle.setLeftBorderColor(oldCell.getCellStyle().getLeftBorderColor());
                    newCellStyle.setRightBorderColor(oldCell.getCellStyle().getRightBorderColor());
                    newCellStyle.setRotation(oldCell.getCellStyle().getRotation());
                    newCellStyle.setTopBorderColor(oldCell.getCellStyle().getTopBorderColor());
                    newCellStyle.setVerticalAlignment(oldCell.getCellStyle().getVerticalAlignment());

                    styleList.add(newCellStyle);
                }
                newCell.setCellStyle(newCellStyle);
            }
        }
        switch (oldCell.getCellType()) {
            case STRING:
                newCell.setCellValue(oldCell.getStringCellValue());
                break;
            case NUMERIC:
                newCell.setCellValue(oldCell.getNumericCellValue());
                break;
            case BLANK:
                newCell.setCellValue("");
                break;
            case BOOLEAN:
                newCell.setCellValue(oldCell.getBooleanCellValue());
                break;
            case ERROR:
                newCell.setCellErrorValue(oldCell.getErrorCellValue());
                break;
            case FORMULA:
                newCell.setCellFormula(oldCell.getCellFormula());
                break;
            default:
                break;
        }

    }

    public static XSSFCell copyHSSFCellstyle(Cell oldCell, Cell newCell) {


        Map<Integer, XSSFCellStyle> styleMap = new HashMap<>();
        XSSFCell cell = (XSSFCell) newCell;
        int stHashCode = oldCell.getCellStyle().hashCode();
        XSSFCellStyle newCellStyle = styleMap.get(stHashCode);
        if (newCellStyle == null) {
            newCellStyle = cell.getSheet().getWorkbook().createCellStyle();
            newCellStyle.cloneStyleFrom(oldCell.getCellStyle());
            styleMap.put(stHashCode, newCellStyle);
        }
        cell.setCellStyle(newCellStyle);
        return cell;


    }

    public static void copyHSSFCell(HSSFCell oldCell, HSSFCell newCell, Map<Integer, HSSFCellStyle> styleMap) {
        if (styleMap != null) {
            if (oldCell.getSheet().getWorkbook() == newCell.getSheet().getWorkbook()) {
                newCell.setCellStyle(oldCell.getCellStyle());
            } else {
                int stHashCode = oldCell.getCellStyle().hashCode();
                HSSFCellStyle newCellStyle = styleMap.get(stHashCode);
                if (newCellStyle == null) {
                    newCellStyle = newCell.getSheet().getWorkbook().createCellStyle();
                    newCellStyle.cloneStyleFrom(oldCell.getCellStyle());
                    styleMap.put(stHashCode, newCellStyle);
                }
                newCell.setCellStyle(newCellStyle);
            }
        }
        switch (oldCell.getCellTypeEnum()) {
            case STRING:
                newCell.setCellValue(oldCell.getStringCellValue());
                break;
            case NUMERIC:
                newCell.setCellValue(oldCell.getNumericCellValue());
                break;
            case BLANK:
                newCell.setCellType(CellType.BLANK);
                break;
            case BOOLEAN:
                newCell.setCellValue(oldCell.getBooleanCellValue());
                break;
            case ERROR:
                newCell.setCellErrorValue(oldCell.getErrorCellValue());
                break;
            case FORMULA:
                newCell.setCellFormula(oldCell.getCellFormula());
                formulaInfoList.add(new FormulaInfo(oldCell.getSheet().getSheetName(), oldCell.getRowIndex(), oldCell.getColumnIndex(), oldCell.getCellFormula()));
                break;
            default:
                break;
        }

    }

    public static XSSFSheet copyRowStyle(XSSFSheet fromSheet, XSSFSheet toSheet, int fromRow, int toRow) {

        int maxColumnNum = 0;

        List<CellStyle> styleMap2 = new ArrayList<CellStyle>();

        XSSFRow srcRow = fromSheet.getRow(fromRow);
        XSSFRow destRow = toSheet.getRow(toRow);
//        XSSFRow destRow = toSheet.createRow(toRow);
        if (srcRow != null && destRow != null) {

            Set<CellRangeAddressWrapper> mergedRegions = new TreeSet<CellRangeAddressWrapper>();
            short dh = fromSheet.getDefaultRowHeight();
            if (srcRow.getHeight() != dh) {
                destRow.setHeight(srcRow.getHeight());
            }

            int j = srcRow.getFirstCellNum();
            if (j < 0) {
                j = 0;
            }
            for (; j <= srcRow.getLastCellNum(); j++) {
                XSSFCell oldCell = srcRow.getCell(j); // ancienne cell
                XSSFCell newCell = destRow.getCell(j); // new cell
                if (oldCell != null) {
                    if (newCell == null) {
                        newCell = destRow.createCell(j);
                    }

                    copyCell(oldCell, newCell, styleMap2);
                    CellRangeAddress mergedRegion = getMergedRegion(fromSheet, srcRow.getRowNum(),
                            (short) oldCell.getColumnIndex());
                    CellRangeAddress newMergedRegion = getMergedRegion(toSheet, destRow.getRowNum(),
                            (short) oldCell.getColumnIndex());

//                    CellRangeAddress toMerge = new CellRangeAddress(newMergedRegion.getFirstRow(),
//                            newMergedRegion.getLastRow(), mergedRegion.getFirstColumn(), mergedRegion.getLastColumn());
//
//                    toSheet.addMergedRegion(toMerge);
                    if (mergedRegion != null) {
//                        CellRangeAddress newMergedRegion = new CellRangeAddress(mergedRegion.getFirstRow(),
//                                mergedRegion.getLastRow(), mergedRegion.getFirstColumn(), mergedRegion.getLastColumn());
//                        System.out.println("getFirstRow"+newMergedRegion.getFirstRow());

                        CellRangeAddress toMerge = new CellRangeAddress(newMergedRegion.getFirstRow(),
                                newMergedRegion.getLastRow(), mergedRegion.getFirstColumn(), mergedRegion.getLastColumn());

                        CellRangeAddressWrapper wrapper = new CellRangeAddressWrapper(toMerge);
                        mergedRegions.add(wrapper);
                        if (isNewMergedRegion(wrapper, mergedRegions)) {

                            toSheet.addMergedRegion(toMerge);
                        }

                    }
                }
            }
            if (srcRow.getLastCellNum() > maxColumnNum) {
                maxColumnNum = srcRow.getLastCellNum();
            }
        }

        for (int i = 0; i <= maxColumnNum; i++) {
            toSheet.setColumnWidth(i, fromSheet.getColumnWidth(i));
        }

        return toSheet;
    }

    public static XSSFSheet copyRowStyle2(XSSFSheet fromSheet, XSSFSheet toSheet, int fromRow, int toRow) {

        int maxColumnNum = 0;

        List<CellStyle> styleMap2 = new ArrayList<CellStyle>();

        XSSFRow srcRow = fromSheet.getRow(fromRow);
        XSSFRow destRow = toSheet.getRow(toRow);
//        XSSFRow destRow = toSheet.createRow(toRow);
        if (srcRow != null && destRow != null) {

            Set<CellRangeAddressWrapper> mergedRegions = new TreeSet<CellRangeAddressWrapper>();
            short dh = fromSheet.getDefaultRowHeight();
            if (srcRow.getHeight() != dh) {
                destRow.setHeight(srcRow.getHeight());
            }

            int j = srcRow.getFirstCellNum();
            if (j < 0) {
                j = 0;
            }
            for (; j <= srcRow.getLastCellNum(); j++) {
                XSSFCell oldCell = srcRow.getCell(j); // ancienne cell
                XSSFCell newCell = destRow.getCell(j); // new cell
                if (oldCell != null) {
                    if (newCell == null) {
                        newCell = destRow.createCell(j);
                    }

                    copyCell(oldCell, newCell, styleMap2);
                    CellRangeAddress mergedRegion = getMergedRegion(fromSheet, srcRow.getRowNum(),
                            (short) oldCell.getColumnIndex());
//                  CellRangeAddress newMergedRegion = getMergedRegion(toSheet, destRow.getRowNum(),
//                            (short) oldCell.getColumnIndex());

                    if (mergedRegion != null) {
                        CellRangeAddress newMergedRegion = new CellRangeAddress(toRow,
                                toRow, mergedRegion.getFirstColumn(), mergedRegion.getLastColumn());

                        CellRangeAddressWrapper wrapper = new CellRangeAddressWrapper(newMergedRegion);

                        if (isNewMergedRegion(wrapper, mergedRegions)) {
//                            System.out.println ("!!!");
                            mergedRegions.add(wrapper);
                            toSheet.addMergedRegion(newMergedRegion);
                        }

                    }
                }
            }
            if (srcRow.getLastCellNum() > maxColumnNum) {
                maxColumnNum = srcRow.getLastCellNum();
            }
        }

        for (int i = 0; i <= maxColumnNum; i++) {
            toSheet.setColumnWidth(i, fromSheet.getColumnWidth(i));
        }

        return toSheet;
    }

    public static final String dir_1 =
            "C:" + File.separator
                    + "Users" + File.separator
                    + "x" + File.separator
                    + "x" + File.separator
                    + "x" + File.separator
                    + "x.xlsx";

    public static final String dir_2 =
            "C:" + File.separator
                    + "Users" + File.separator
                    + "x" + File.separator
                    + "x" + File.separator
                    + "x" + File.separator
                    + "x2.xlsx";
    public static final String micro = "micro";
    public static final String proj_path = FileSystemView.getFileSystemView().getHomeDirectory().getPath() + File.separator + micro;
    public static String template = "template.xlsx";
    public static final String template_absolute_path = FileSystemView.getFileSystemView().getHomeDirectory().getPath() + File.separator + template;

    public static final String mrcroorganism_arr[] = {"e.", "salmonella", "aureus", "escherichia", "staphylococcus aureus", "pseudomonas", "aeruginosa"};
    //    public static String destination = "copy.xlsx";
//    public static final String destination_absolute_path = FileSystemView.getFileSystemView().getHomeDirectory().getPath() + File.separator + destination;

    //1 找尋所有舊表格需用字串
    //2 計算 test item &spec
    //3 刪除舊表格
    //4 新增新表格
    //5 存檔
    public static void main(String[] args) throws IOException, InvalidFormatException {

        List<String> mrcroorganism = Arrays.asList(mrcroorganism_arr);
//        File destination = new File(destination_absolute_path);
//        File template = new File(template_absolute_path);
//        String micro = "micro";
        List<File> f = getFilesByKeywords(proj_path, "");
        for (File ff : f) {
            String destination_absolute_path = ff.getAbsolutePath();

            Workbook wb_destination = new XSSFWorkbook(OPCPackage.open(destination_absolute_path));
            Workbook wb_template = new XSSFWorkbook(OPCPackage.open(template_absolute_path));

            Excel destination = Excel.loadExcel(destination_absolute_path);
            Excel template = Excel.loadExcel(template_absolute_path);

            Sheet oldSheet = getSheetsBykeywordsIgnoreCase(destination, "micro").get(0);
            template.assignSheet(0);
            destination.assignSheet(wb_destination.getSheetIndex(oldSheet.getSheetName()));

            System.out.println("SheetName :" + oldSheet.getSheetName() + ", index: " + wb_destination.getSheetIndex(oldSheet.getSheetName()));
            int countfile = 1;
//1 找尋所有舊表格需用字串
            if (!StringUtils.isBlank(destination.findFirstWordInWb("Inspection Notebook", destination)) ||
                    !StringUtils.isBlank(destination.findFirstWordInWb("plate 1", destination))) {
                System.out.println(countfile + " " + "fileName :" + ff.getName());

                String ProductName = destination.findFirstWordInWb("Product Name", destination);
                String thisProductName = destination.findfirstWordAtRight(destination.getCell(ProductName).getRowIndex(), destination.getCell(ProductName).getColumnIndex());

                //標題Test Item
                String TestItem_0 = destination.findFirstWordInWb("Test Item", destination);
                String thisTestItem_0 = destination.findfirstWordAtRight(destination.getCell(TestItem_0).getRowIndex(), destination.getCell(TestItem_0).getColumnIndex());

                String Code = destination.findFirstWordInWb("Code", destination);
                String thisCode = destination.findfirstWordAtRight(destination.getCell(Code).getRowIndex(), destination.getCell(Code).getColumnIndex());

                //Test Item子項目
//        String TestItem = excel.findFirstWordInWb("Test Item", excel);
                String TestItem = destination.findFirstWord("Test Item", 6, destination.getLastRowNum(), 0, destination.getLastCellNum());
                String result = destination.findFirstWord("result", 6, destination.getLastRowNum(), 0, destination.getLastCellNum());

                String Specification = destination.findFirstWordInWb("Specification", destination);

                if (!StringUtils.isBlank(TestItem)) {
                    List<String> testItemList = destination.getCellsStringValue(destination.getCell(TestItem).getRowIndex(),
                            destination.getCell(Specification).getRowIndex() - 1,
                            0,
                            destination.getCell(Specification).getColumnIndex());
                    testItemList.removeAll(Arrays.asList("test item", "Test Item", "Test item"));
                    System.out.println(testItemList.size());
                    for (String s : testItemList)
                        System.out.println(s);
                }
                List<String> testItemList = new ArrayList<>();
                if (!StringUtils.isBlank(result)) {
                    testItemList = destination.getCellsStringValue(destination.getCell(result).getRowIndex(),
                            destination.getCell(Specification).getRowIndex() - 1,
                            0,
                            destination.getCell(Specification).getColumnIndex());
                    testItemList.removeAll(Arrays.asList("test item", "Test Item", "Test item"));
                    System.out.println(testItemList.size());
                    for (String s : testItemList)
                        System.out.println(s);
                }

                List<String> specificationList = destination.getCellsStringValue(destination.getCell(Specification).getRowIndex(),
                        destination.getLastRowNum() - 1,
                        destination.getCell(Specification).getColumnIndex(),
                        destination.getLastCellNum()
                );
                List<String> newSpecificationList = specificationList;
                for (int i = 0; i < specificationList.size(); i++) {

                    if (Excel.ifContainNewLine(specificationList.get(i))) {
//                System.out.println("!!");
                        String[] specSperate = specificationList.get(i).split("\n");
                        for (String sss : specSperate) {
                            newSpecificationList.add(sss);
                        }
                        newSpecificationList.remove(specificationList.get(i));
                    }

                }

                for (int i = 0; i < newSpecificationList.size(); i++) {
                    String s = newSpecificationList.get(i);
                    if (s.toLowerCase(Locale.ROOT).contains("specification")) {
                        newSpecificationList.remove(s);

                    }
                }
//
                for (String s : newSpecificationList) {
                    System.out.println(s);
                }
                System.out.println("size: " + newSpecificationList.size());
                //2 計算 test item &spec

                int testItemRowStartIdx = 6;
                int testItemNum = testItemList.size();
                int specNum = newSpecificationList.size();

                //3 刪除舊表格
                wb_destination.removeSheetAt(wb_destination.getSheetIndex(oldSheet.getSheetName()));
                //4 新增新表格
                wb_destination.createSheet(oldSheet.getSheetName());
                /*Get sheets from the temp file*/
                XSSFSheet sheet_1 = ((XSSFWorkbook) wb_destination).getSheet(oldSheet.getSheetName());
                XSSFSheet sheet_2 = ((XSSFWorkbook) wb_template).getSheetAt(0);
                copySheet(sheet_2, sheet_1);
                copySheetSettings(sheet_2, sheet_1);

                //4-1 填值入新表格

                sheet_1.getRow(0).getCell(1).setCellValue(thisProductName);
                sheet_1.getRow(1).getCell(1).setCellValue(thisCode);
                sheet_1.getRow(3).getCell(1).setCellValue(thisTestItem_0);


                for (int i = testItemRowStartIdx + 1; i < testItemRowStartIdx + testItemNum; i++) {
                    int lastRow = sheet_1.getPhysicalNumberOfRows();
                    System.out.print(i);
                    try {
                        sheet_1.shiftRows(i, lastRow, 1, true, true);
                    } catch (Exception e) {
                        System.out.println(i + ".REMOVE合併結果:" + Excel.removeMergedArea(sheet_1, lastRow + 1, lastRow + 1));
                        for (int idx_test = 0; idx_test < sheet_1.getNumMergedRegions(); idx_test++) {
                            CellRangeAddress region = sheet_1.getMergedRegion(idx_test);
                            System.out.println("合併區域:" + region.formatAsString());
                        }
                        System.out.println("第一航:" + i + ";最後:" + lastRow);
                        sheet_1.shiftRows(i, lastRow, 1, true, true);
                    }

                    sheet_1.createRow(i);
                    int lastCell = sheet_1.getRow(i).getLastCellNum();
                    for (int j = 0; j < lastCell; j++) {
                        sheet_1.getRow(i).createCell(j);
                    }
                    copyRowStyle2(sheet_1, sheet_1, 6, i);
                }

                int count = testItemRowStartIdx;
                for (int i = testItemRowStartIdx; i < testItemRowStartIdx + testItemNum; i++) {
                    if (sheet_1.getRow(i).getCell(0) == null) {
                        sheet_1.getRow(i).createCell(0);
                    }
                    //negative control的微生物側項要斜體
                    Cell cell = sheet_1.getRow(i).getCell(0);
                    String testitem_w = testItemList.get(count - testItemRowStartIdx);

                    CheckWords:
                    for (String testitem : testitem_w.split(" ")) {

//                        CellStyle style = cell.getCellStyle();
                        CellStyle style = wb_destination.createCellStyle();
                        style.setBorderBottom(BorderStyle.THIN);

                        cell.setCellStyle(style);
                        cell.setCellValue(testitem_w);

                        if (mrcroorganism.contains(testitem.toLowerCase().trim())) {
                            Font defaultFont = wb_destination.createFont();
                            defaultFont.setItalic(true);
                            defaultFont.setFontName("Times New Roman");
                            defaultFont.setFontHeightInPoints((short) 12);
                            style.setFont(defaultFont);
                            break CheckWords;
                        }
                    }

                    count++;
                }

                for (int i = count + 1, y = 0; y < newSpecificationList.size(); i++) {
                    System.out.println(count);
                    if (sheet_1.getRow(i) == null) {
                        sheet_1.createRow(i);
                    }
                    if (sheet_1.getRow(i).getCell(1) == null) {
                        sheet_1.getRow(i).createCell(1);
                    }
                    String Specification_1 = newSpecificationList.get(y);
                    Cell specCell = sheet_1.getRow(i).getCell(1);
                    specCell.setCellValue(Specification_1);
                    CheckWord:
                    for (String spec_1 : Specification_1.split(" ")) {


                        CellStyle style = wb_destination.createCellStyle();


                        specCell.setCellStyle(style);
                        specCell.setCellValue(Specification_1);
                        Font defaultFont = wb_destination.createFont();
                        defaultFont.setBold(true);
                        defaultFont.setUnderline(Font.U_SINGLE);
                        defaultFont.setFontHeightInPoints((short) 14);
                        defaultFont.setFontName("Times New Roman");
                        if (mrcroorganism.contains(spec_1.toLowerCase().trim())) {

                            defaultFont.setItalic(true);

                            style.setFont(defaultFont);

                            break CheckWord;
                        }
                        style.setFont(defaultFont);
//                        else {
//
//                        }
                    }

                    count++;
                    y++;
                }

                //重設border粗體
                CellRangeAddress region = CellRangeAddress.valueOf("A5:E5");
//        short borderStyle = CellStyle.BORDER_MEDIUM;
                RegionUtil.setBorderBottom(BorderStyle.THICK, region, sheet_1);
                destination.removeLastBlankRow(sheet_1);
                //5 存檔
                OutputStream os = new FileOutputStream(ff.getName());
                System.out.print("If you arrived here, it means you're good boy");
                wb_destination.write(os);
                os.flush();
                os.close();
                wb_destination.close();
            }
            countfile++;
        }

    }
}