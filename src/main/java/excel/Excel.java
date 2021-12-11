package excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import excel.model.Producto;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

public class Excel {

    public static void main(String[] args) {
        try {
            ArrayList<Producto> productos = cargarArchivoPlano("C:\\Users\\samad\\Documents\\Geral\\PrimerProyecto\\",
                    "healthcare10.xlsx");

            for (Producto producto : productos) {
                System.out.println(producto.getNombreProducto());
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static ArrayList<Producto> cargarArchivoPlano(String path, String fileName) throws IOException {
        ArrayList<Producto> list = new ArrayList<>();
        // read excel file from file system
        FileInputStream excelFile = new FileInputStream(new File(path + fileName));
        // Access the workbook
        XSSFWorkbook workbook = new XSSFWorkbook(excelFile);
        // Access the sheet
        XSSFSheet sheet = workbook.getSheetAt(0);
        // Iterate through each rows one by one
        Iterator<Row> rowIterator = sheet.iterator();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            // For each row, iterate through all the columns
            Iterator<Cell> cellIterator = row.cellIterator();
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                // Check the cell type and format accordingly
                switch (cell.getCellTypeEnum()) {
                    case STRING:
                        // System.out.print("STRING: " + cell.getStringCellValue() + "\t");
                        if (!cell.getStringCellValue().equals("producto")) {
                            Producto prod = new Producto(cell.getStringCellValue());
                            list.add(prod);
                        }
                        break;
                    case NUMERIC:
                        // System.out.print("NUMERIC: " + cell.getNumericCellValue() + "\t");
                        break;
                    case BLANK:
                        // System.out.print("\t");
                        break;
                    default:
                }
            }
            // System.out.println();
        }
        excelFile.close();
        workbook.close();
        return list;
    }

    public static ArrayList<Map<String, String>> leerDatosDeHojaDeExcel(String rutaDeExcel, String hojaDeExcel)
            throws IOException {
        ArrayList<Map<String, String>> arrayListDatoPlanTrabajo = new ArrayList<Map<String, String>>();
        Map<String, String> informacionProyecto = new HashMap<String, String>();
        File file = new File(rutaDeExcel);
        FileInputStream inputStream = new FileInputStream(file);
        XSSFWorkbook newWorkbook = new XSSFWorkbook(inputStream);
        XSSFSheet newSheet = newWorkbook.getSheet(hojaDeExcel);
        Iterator<Row> rowIterator = newSheet.iterator();
        Row titulos = rowIterator.next();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Iterator<Cell> cellIterator = row.cellIterator();
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                cell.getColumnIndex();
                switch (cell.getCellTypeEnum()) {
                    case STRING:
                        informacionProyecto.put(titulos.getCell(cell.getColumnIndex()).toString(),
                                cell.getStringCellValue());
                        break;
                    case NUMERIC:
                        informacionProyecto.put(titulos.getCell(cell.getColumnIndex()).toString(),
                                String.valueOf((long) cell.getNumericCellValue()));
                        break;
                    case BLANK:
                        informacionProyecto.put(titulos.getCell(cell.getColumnIndex()).toString(), "");
                        break;
                    default:
                }
            }
            arrayListDatoPlanTrabajo.add(informacionProyecto);
            informacionProyecto = new HashMap<String, String>();
        }
        newWorkbook.close();
        return arrayListDatoPlanTrabajo;
    }

}
