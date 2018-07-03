package utils;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.WorkbookUtil;

public class LibroExcel {

	private File file = null;
	//private File tempFile = null;
	private HSSFWorkbook workbook = null;
	private HSSFSheet sheet = null;

	public LibroExcel(String documentName){
		try {
			file = new File(documentName+".xls");

			if(!file.exists()){
				file.createNewFile();
				workbook = new HSSFWorkbook();
			}else
				workbook = (HSSFWorkbook) WorkbookFactory.create(file);

			if(workbook.getNumberOfSheets() == 0)
				sheet = workbook.createSheet(WorkbookUtil.createSafeSheetName(documentName));
			else
				sheet = workbook.getSheetAt(0);

		} catch (EncryptedDocumentException e) { e.printStackTrace();
		} catch (InvalidFormatException e) { e.printStackTrace();
		} catch (FileNotFoundException e) {e.printStackTrace();
		} catch (IOException e) { e.printStackTrace();
		}
	}
	
	
	public LibroExcel(String pathFile, String documentName){
		try {
			this.file = new File(documentName+".xls");

			if(!this.file.exists()){
				this.file.createNewFile();
				this.workbook = new HSSFWorkbook();
			}else
				this.workbook = (HSSFWorkbook) WorkbookFactory.create(this.file);

			if(this.workbook.getNumberOfSheets() == 0)
				this.sheet = this.workbook.createSheet(WorkbookUtil.createSafeSheetName(documentName));
			else
				this.sheet = this.workbook.getSheetAt(0);

		} catch (EncryptedDocumentException e) { e.printStackTrace();
		} catch (InvalidFormatException e) { e.printStackTrace();
		} catch (FileNotFoundException e) {e.printStackTrace();
		} catch (IOException e) { e.printStackTrace();
		}
	}
	

	public LibroExcel(File file){
		try{
			this.file = file;
			if(file.length() == 0)
				this.workbook = new HSSFWorkbook();
			else
				this.workbook = (HSSFWorkbook) WorkbookFactory.create(file);

			if(this.workbook.getNumberOfSheets() == 0)
				this.sheet = this.workbook.createSheet(WorkbookUtil.createSafeSheetName("hoja"));
			else
				this.sheet = this.workbook.getSheetAt(0);
		} catch (EncryptedDocumentException e) { e.printStackTrace();
		} catch (InvalidFormatException e) { e.printStackTrace();
		} catch (FileNotFoundException e) {e.printStackTrace();
		} catch (IOException e) { e.printStackTrace();
		}
	}

	/**
	 * save: Guarda el documento.
	 */
	public void save(){
		try {
			File file = this.file;
			HSSFWorkbook workbook = this.workbook;
			FileOutputStream fos = new FileOutputStream(file);
			workbook.write(fos);
			fos.close();
		} catch (FileNotFoundException e) {e.printStackTrace();
		} catch (IOException e) { e.printStackTrace();
		}
	}

	/**
	 * save: Guarda el documento en un archivo especifico.
	 * @param pathname
	 */
	public void save(String pathname){
		try {
			File file = new File(pathname+".xls");
			HSSFWorkbook workbook = this.workbook;
			FileOutputStream fos = new FileOutputStream(file);
			workbook.write(fos);
			fos.close();
		} catch (FileNotFoundException e) {e.printStackTrace();
		} catch (IOException e) { e.printStackTrace();
		}
	}

	public void delete(){
		this.file.delete();
	}

	/**
	 * isEmpty: Determina si la hoja no tiene ningun registro.
	 * @return
	 */
	public boolean isEmpty(){
		boolean noRows = firstEmptyRow(0) == 0;
		boolean noColumns = firstEmptyColumn(0) == 0;
		if(noRows && noColumns)
			return true;
		return false;
	}

	/**
	 * setActiveSheet: Selecciona una hoja.
	 * @param sheetIndex
	 */
	public void setActiveSheet(int sheetIndex){
		this.sheet = workbook.getSheetAt(sheetIndex);
	}

	/**
	 * setActiveSheet: Selecciona una hoja.
	 * @param sheetName
	 */
	public void setActiveSheet(String sheetName){
		this.sheet = workbook.getSheet(sheetName);
	}

	/**
	 * addSheet: Crea una nueva hoja, si ya existe no la agrega.
	 * @param sheetname
	 */
	public void addSheet(String sheetname){
		try {			
			for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
				if(sheetname.equals(workbook.getSheetName(i)))
					return;
			}
			workbook.createSheet(sheetname);
		} catch (EncryptedDocumentException e) { e.printStackTrace();
		}
	}

	/**
	 * addInACell: Agrega un registro en una posicion especifica.
	 * @param rowIndex
	 * @param columnIndex
	 * @param registro
	 */
	public void addInACell(int rowIndex, int columnIndex, Object registro){
		try {
			HSSFRow row = sheet.getRow(rowIndex);
			if(row == null)
				row = sheet.createRow(rowIndex);
			HSSFCell cell = row.getCell(columnIndex);
			if(cell == null)
				cell = row.createCell(columnIndex);

			String aux = registro.getClass().getName();
			if(aux.equals(String.class.getName())) //String
				cell.setCellValue((String)registro);
			if(aux.equals(Integer.class.getName())) //Interger
				cell.setCellValue((Integer)registro);
			if(aux.equals(Double.class.getName())) //Double
				cell.setCellValue((Double)registro);
			if(aux.equals(Boolean.class.getName())) //Boolean
				cell.setCellValue((Boolean)registro);

		} catch (EncryptedDocumentException e) { e.printStackTrace();
		}
	}
	/**
	 * searchInRow: Busca un registro en una fila. Retorna el indice de la columna donde encuentra el registro y -1 si no lo encuentra.
	 * @param rowIndex
	 * @param registro
	 * @return columnIndex
	 */
	public int searchInRow(int rowIndex, Object registro){
		HSSFRow row = sheet.getRow(rowIndex);
		int columnIndex = 0;
		HSSFCell cell  = row.getCell(columnIndex);
		while(cell.getCellTypeEnum() != CellType.BLANK){
			String aux = registro.getClass().getName();

			//Celda String
			if(cell.getCellTypeEnum() == CellType.STRING && aux.equals(String.class.getName())){
				if(cell.getStringCellValue().equals((String)registro))
					return cell.getColumnIndex();
			}
			//Celda Numerica
			boolean registroIsNumeric = aux.equals(Integer.class.getName()) || aux.equals(Double.class.getName());
			if(cell.getCellTypeEnum() == CellType.NUMERIC &&  registroIsNumeric ){
				if(cell.getNumericCellValue() == (Integer)registro)
					return cell.getColumnIndex();
			}

			//Celda Booleana
			if(cell.getCellTypeEnum() == CellType.BOOLEAN && aux.equals(Boolean.class.getName())){
				if(cell.getBooleanCellValue() == (Boolean)registro)
					return cell.getColumnIndex();
			}

			//Si no es igual, busca en la siguiente columna
			columnIndex++;
			cell = row.getCell(columnIndex);
		}
		return (Integer) null;
	}

	/**
	 * searchInColumn: Busca un registro en una columna. Retorna el indice de la fila donde se encuentra el registro y -1 si no lo encuentra. 
	 * @param columnIndex
	 * @param registro
	 * @return rowIndex
	 */
	public int searchInColumn(int columnIndex, Object registro){
		int rowIndex = 0;
		HSSFRow row = sheet.getRow(rowIndex);
		HSSFCell cell  = row.getCell(columnIndex);
		while(cell.getCellTypeEnum() != CellType.BLANK){
			String aux = registro.getClass().getName();

			//Celda String
			if(cell.getCellTypeEnum() == CellType.STRING && aux.equals(String.class.getName())){
				if(cell.getStringCellValue().equals((String)registro))
					return cell.getColumnIndex();
			}
			//Celda Numerica
			boolean registroIsNumeric = aux.equals(Integer.class.getName()) || aux.equals(Double.class.getName());
			if(cell.getCellTypeEnum() == CellType.NUMERIC &&  registroIsNumeric ){
				if(cell.getNumericCellValue() == (Integer)registro)
					return cell.getColumnIndex();
			}

			//Celda Booleana
			if(cell.getCellTypeEnum() == CellType.BOOLEAN && aux.equals(Boolean.class.getName())){
				if(cell.getBooleanCellValue() == (Boolean)registro)
					return cell.getColumnIndex();
			}
			
			//Si no es igual, busca en la siguiente fila
			rowIndex++;
			row = sheet.getRow(rowIndex);
			cell = row.getCell(columnIndex);
		}
		return (Integer) null;
	}

	/**
	 * firstEmptyRow: Retorna el indice de primera columna vacia.
	 * @param columnIndex
	 * @return firstEmptyRow
	 */
	public int firstEmptyRow(int columnIndex){
		int rowIndex = 0;
		while(true){
			HSSFRow row = sheet.getRow(rowIndex);
			if(row == null)
				row = sheet.createRow(rowIndex);
			HSSFCell cell  = row.getCell(columnIndex);
			if(cell == null || cell.getCellTypeEnum() == CellType.BLANK){
				cell  = row.createCell(columnIndex);
				return rowIndex;
			}else
				rowIndex++;
		}
	}

	/**
	 * firstEmptyColumn: Retorna el indice de primera columna vacia.
	 * @param rowIndex
	 * @return firstEmptyColumn
	 */
	public int firstEmptyColumn(int rowIndex){
		int columnIndex = 0;
		while(true){
			HSSFRow row = sheet.getRow(rowIndex);
			if(row == null)
				row = sheet.createRow(rowIndex);
			HSSFCell cell  = row.getCell(columnIndex);
			if(cell == null || cell.getCellTypeEnum() == CellType.BLANK){
				cell  = row.createCell(columnIndex);
				return columnIndex;
			}else
				columnIndex++;
		}
	}

	/**
	 * addInRow: Agrega un registro en la primera celda vacia de una fila.
	 * @param indexRow
	 * @param registro
	 */
	public void addInRow(int indexRow, Object registro){
		addInACell(indexRow, firstEmptyColumn(indexRow), registro);
	}

	/**
	 * addInColumn: Agrega un registro en la primera celda vacia de una columna.
	 * @param columnIndex
	 * @param registro
	 */
	public void addInColumn(int columnIndex, Object registro){
		addInACell(firstEmptyRow(columnIndex), columnIndex, registro);
	}

	
	/**
	 * getCellValue: Obtiene el valor de una celda espefifica.
	 * @param rowIndex
	 * @param columnIndex
	 * @return object
	 */
	public Object getCellValue(int rowIndex, int columnIndex){
		HSSFRow row = sheet.getRow(rowIndex);
		if(row == null)
			row = sheet.createRow(rowIndex);
		HSSFCell cell  = row.getCell(columnIndex);
		if(cell == null)
			return null;

		if(cell.getCellTypeEnum() == CellType.STRING)
			return cell.getStringCellValue();
		if(cell.getCellTypeEnum() == CellType.NUMERIC){
			double numero = cell.getNumericCellValue();
			if(numero-(long)numero == 0){
				return (int)numero;
			}else
				return cell.getNumericCellValue();
		}
		if(cell.getCellTypeEnum() == CellType.BOOLEAN)
			return cell.getBooleanCellValue();

		return null;

	}
	
	//TEST
	public void removeRow(int rowIndex){
		int columnIndex = 0;
		HSSFRow row = sheet.getRow(rowIndex);
		HSSFCell cell = row.getCell(columnIndex);
		
		int lastRow = firstEmptyRow(columnIndex) - 1;
		
		for (int i = rowIndex; i < lastRow; i++) {
			Object aux = getCellValue(rowIndex + 1, columnIndex);
			addInACell(rowIndex, columnIndex, aux);
			
			rowIndex = i;
			row = sheet.getRow(rowIndex);
			cell = row.getCell(columnIndex);
		}
	}

}
