package br.com.leitor.xls;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class EditaExcel {

	private static final String fileName = "C:/testes/clientes.xls";

	public static void main(String[] args) throws IOException {

		try {
			FileInputStream file = new FileInputStream(new File(EditaExcel.fileName));

			HSSFWorkbook workbook = new HSSFWorkbook(file);
			HSSFSheet sheetClientes = workbook.getSheetAt(0);

			for (int i = 0; i < sheetClientes.getPhysicalNumberOfRows(); i++) {

				Row row = sheetClientes.getRow(i);
				Cell cellValorFinal = row.getCell(5);
				Cell cellAprovado = row.getCell(6);
				cellAprovado.setCellValue(cellValorFinal.getNumericCellValue() <= 300);
			}
			file.close();

			FileOutputStream outFile = new FileOutputStream(new File(EditaExcel.fileName));
			workbook.write(outFile);
			outFile.close();
			System.out.println("Arquivo Excel editado com sucesso!");

		} catch (FileNotFoundException e) {
			e.printStackTrace();
			System.out.println("Arquivo Excel não encontrado!");
		} catch (IOException e) {
			e.printStackTrace();
			System.out.println("Erro na edição do arquivo!");
		}

	}

}
