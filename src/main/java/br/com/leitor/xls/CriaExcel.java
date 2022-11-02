package br.com.leitor.xls;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class CriaExcel {

	private static final String fileName = "C:/testes/clientes.xls";

	public static void main(String[] args) throws IOException {

		HSSFWorkbook workbook = new HSSFWorkbook();
		HSSFSheet sheetClientes = workbook.createSheet("Clientes");

		List<Cliente> listaClientes = new ArrayList<Cliente>();
		listaClientes.add(new Cliente("Eduardo", "9876525", 77.90, 8563, 77.90, false));
		listaClientes.add(new Cliente("Luiz", "1234466", 85.65, 2130, 85.65, false));
		listaClientes.add(new Cliente("Bruna", "6545657", 754.56, 4520, 754.56, false));
		listaClientes.add(new Cliente("Carlos", "3456558", 130.12, 7520, 130.12, false));
		listaClientes.add(new Cliente("Sonia", "6544546", 765.00, 7860, 765.00, false));
		listaClientes.add(new Cliente("Brianda", "3234535", 66.71, 7540, 66.71, false));
		listaClientes.add(new Cliente("Pedro", "4234524", 79.65, 3240, 79.65, false));
		listaClientes.add(new Cliente("Julio", "5434513", 777.68, 2280, 777.68, false));
		listaClientes.add(new Cliente("Henrique", "6543452", 742.20, 3210, 742.20, false));
		listaClientes.add(new Cliente("Fernando", "4345651", 551.23, 1230, 551.23, false));
		listaClientes.add(new Cliente("Vitor", "4332341", 777.64, 0352, 777.64, false));

		int rownum = 0;
		for (Cliente Cliente : listaClientes) {
			Row row = sheetClientes.createRow(rownum++);
			int cellnum = 0;
			Cell cellNome = row.createCell(cellnum++);
			cellNome.setCellValue(Cliente.getNome());
			Cell cellProtocolo = row.createCell(cellnum++);
			cellProtocolo.setCellValue(Cliente.getProtocolo());
			Cell cellGlosa = row.createCell(cellnum++);
			cellGlosa.setCellValue(Cliente.getGlosa());
			Cell cellPedido = row.createCell(cellnum++);
			cellPedido.setCellValue(Cliente.getPedido());
			Cell cellValorFinal = row.createCell(cellnum++);
			cellValorFinal.setCellValue(Cliente.getValorFinal());
			
			Cell cellAprovado = row.createCell(cellnum++);
			cellAprovado.setCellValue(cellValorFinal.getNumericCellValue() <= 750);
		}

		try {
			FileOutputStream out = new FileOutputStream(new File(CriaExcel.fileName));
			workbook.write(out);
			out.close();
			System.out.println("Arquivo Excel criado com sucesso!");

		} catch (FileNotFoundException e) {
			e.printStackTrace();
			System.out.println("Arquivo não encontrado!");
		} catch (IOException e) {
			e.printStackTrace();
			System.out.println("Erro na edição do arquivo!");
		}

	}

}