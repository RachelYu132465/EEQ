package dataStructure;


import org.apache.poi.ss.usermodel.Cell;

public class myCell {


public Cell cell;
	String lefttitle;
	String uptittle;


	int[] position = new int[2];
	String ecxcelFormatName;
	String value;
	String Formula;

	public myCell(){}
	public myCell(String ecxcelFormatName, String value, String Formula,

			String lefttitle, String uptittle) {
		this.ecxcelFormatName = ecxcelFormatName;
		this.value = value;
		this.Formula = Formula;

		this.lefttitle = lefttitle;
		this.uptittle = uptittle;
	}

	public Cell getCell() {
		return cell;
	}

	public void setCell(Cell cell) {
		this.cell = cell;
	}
}
