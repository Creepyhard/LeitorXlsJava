package br.com.leitor.xls;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class AbreExcel {

	private static final String fileName = "C:/testes/clientes.xls";

	public static void main(String[] args) throws IOException {

		List<Cliente> listaClientes = new ArrayList<Cliente>();

		try {
			FileInputStream arquivo = new FileInputStream(new File(AbreExcel.fileName));

			HSSFWorkbook workbook = new HSSFWorkbook(arquivo);

			HSSFSheet sheetClientes = workbook.getSheetAt(0);

			Iterator<Row> rowIterator = sheetClientes.iterator();

			while (rowIterator.hasNext()) {
				Row row = rowIterator.next();
				Iterator<Cell> cellIterator = row.cellIterator();

				Cliente cliente = new Cliente();
				listaClientes.add(cliente);
				while (cellIterator.hasNext()) {
					Cell cell = cellIterator.next();
					switch (cell.getColumnIndex()) {
					case 0:
						cliente.setNome(cell.getStringCellValue());
						break;
					case 1:
						cliente.setProtocolo(cell.getStringCellValue());
						break;
					case 2:
						cliente.setGlosa(cell.getNumericCellValue());
						break;
					case 3:
						cliente.setPedido(cell.getNumericCellValue());
						break;
					case 4:
						cliente.setValorFinal(cell.getNumericCellValue());
						break;
					case 5:
						cliente.setAprovado(cell.getBooleanCellValue());
						break;
					}
				}
			}
			arquivo.close();

		} catch (FileNotFoundException e) {
			e.printStackTrace();
			System.out.println("Arquivo Excel não encontrado!");
		}

		if (listaClientes.size() == 0) {
			System.out.println("Nenhum cliente encontrado!");
		} else {
			double somaValorFinal = 0;
			double somaValorRecusado = 0;
			int aprovados =0;
			int reprovados = 0;
			
			for(Cliente cliente : listaClientes) {
                System.out.println("Cliente: " + cliente.getNome() + " || Valor Final: "
                        + cliente.getValorFinal() + " || Aprovado/Reprovado: " + cliente.getAprovado());
				if(cliente.getValorFinal() <= 750) {
					somaValorFinal += cliente.getValorFinal();
					aprovados++;
				} else {
					somaValorRecusado += cliente.getValorFinal();
					reprovados++;
				}
			}
		
		System.out.println("Valor total Recusado: R$" + somaValorRecusado);
		System.out.println("Valor total gasto: R$" + somaValorFinal);
		System.out.println("Número total de clientes: " + listaClientes.size());
		System.out.println("O numero de clientes aprovados é: " + aprovados);
		System.out.println("O numero de clientes reprovados é: " + reprovados);
		}
	}

}
