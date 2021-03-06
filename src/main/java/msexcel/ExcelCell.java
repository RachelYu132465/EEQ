package msexcel;

import org.apache.commons.lang.StringUtils;
import org.apache.poi.ss.usermodel.Cell;

import static msexcel.Excel.getCellValue;
import static validate.excelFormulaProcessor.findBlackTitleAtLeft;
import static validate.excelFormulaProcessor.findBlackTitleAtTop;

public class ExcelCell {
    String note;
    String r1c1;
    //formula cell will store its result value
    String value;
    Cell cell;

    public String getNote() {
        return note;
    }

    public void setNote(String note) {
        this.note = note;
    }

    public void setNote(Excel excel,Cell c) {
        String description2 ="";
         String description1 = findBlackTitleAtLeft(excel,c);
        if (StringUtils.isBlank(description1)){findBlackTitleAtTop(excel,c);}
        //to-do: logic of appending description
        if (StringUtils.isBlank(description1)&&StringUtils.isBlank(description2))
        this.setNote(cell.getAddress().toString());
       else this.setNote(description1 +" "+description2);
    }

    public ExcelCell copyNote(ExcelCell c) {
        this.setNote(c.getNote());
        return this;
    }

    public String getR1c1() {
        return r1c1;
    }

    public void setR1c1(String r1c1) {
        this.r1c1 = r1c1;
    }

    public String getValue() {
        return value;
    }

    public void setValue(String value) {
        this.value = value;
    }

    public Cell getCell() {
        return cell;
    }

    public ExcelCell(){

    }
//    public ExcelCell(Cell cell) {
//        if(cell!=null){
//            this.r1c1 = Excel.getR1C1Idx(cell);
//            this.value = getCellValue(cell).toString();
//            this.cell = cell;
//        }else{
//            return;
//        }
//    }

    public ExcelCell( Excel excel,Cell cell) {
        if(cell!=null){
            this.r1c1 = Excel.getR1C1Idx(cell);
//            this.value = getCellValue(cell).toString();
            System.out.println("in ExcelCell" +excel.getSheet().getSheetName());
            this.value = excel.getCellValueAsString(cell);
            this.cell = cell;
            setNote(excel,cell);
        }else{
            return;
        }
    }
    public ExcelCell(String r1c1, String value, Cell cell) {
        this.r1c1 = r1c1;
        this.value = value;
        this.cell = cell;
    }

    public ExcelCell setCell(Cell cell) {
        this.cell = cell;
        return null;
    }


    @Override
    public  String toString(){
        return r1c1 + ":" + value;
    }

    @Override
    public boolean equals(Object o){
        if (o == this)
            return true;
        if (!(o instanceof ExcelCell))
            return false;
        else {
//            if(((ExcelCell)(o)).getR1c1().equalsIgnoreCase(this.getR1c1())){
//                return true;
            if(((ExcelCell)(o)).getCell().equals(this.getCell())){
                return true;
            }
        }
       return  false;
    }

    @Override
    public final int hashCode() {
        int result = 17;
        result = 31 * result + this.getCell().hashCode();
        return  result;
    }
}
