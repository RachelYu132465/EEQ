package dataStructure;


import org.apache.poi.ss.usermodel.Cell;

public class myCell {

	int[] position = new int[2];
	String ecxcelFormatName;
	String value;
	String Formula;
public Cell cell;
	String lefttitle;
	String uptittle;

	public myCell(){}
	public myCell(String ecxcelFormatName, String value, String Formula,

			String lefttitle, String uptittle) {
		this.ecxcelFormatName = ecxcelFormatName;
		this.value = value;
		this.Formula = Formula;

		this.lefttitle = lefttitle;
		this.uptittle = uptittle;
	}

}
