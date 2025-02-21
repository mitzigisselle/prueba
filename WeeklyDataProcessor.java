
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class WeeklyDataProcessor {
    public static void main(String[] args) throws IOException {
        // Ruta del archivo Excel
        String filePath = "Prueba de Conocimiento - Data Scientist Jr - Perfil Procesamiento de Datos_Automatizacioìn.xlsx";
        FileInputStream file = new FileInputStream(new File(filePath));

        // Cargar el archivo Excel en un objeto Workbook
        Workbook workbook = new XSSFWorkbook(file);
        Sheet sheet = workbook.getSheet("Raw Data"); // Obtener la hoja "Raw Data"

        // Iterar sobre las filas del archivo
        Iterator<Row> rowIterator = sheet.iterator();
        rowIterator.next(); // Saltar la fila de encabezado

        // Imprimir encabezado de los datos procesados en el mismo formato que "Resultado Final"
        System.out.println("FECHA_SEMANA, FUERZA_VENTAS_ARGENTINA_COLORITO, FUERZA_VENTAS_BRASIL_COLORITO, FUERZA_VENTAS_CHILE_COLORITO, FUERZA_VENTAS_PERU_COLORITO, FUERZA_VENTAS_TOTAL_COLORITO");

        // Formateador de fecha para convertir la fecha en un string legible
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");

        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Cell dateCell = row.getCell(0); // Obtener la celda de fecha

            // Validar que la celda contenga una fecha válida
            if (dateCell != null && dateCell.getCellType() == CellType.NUMERIC) {
                Date date = dateCell.getDateCellValue(); // Obtener el valor de la fecha
                String dateStr = sdf.format(date); // Convertir la fecha a formato String

                // Obtener valores numéricos de las demás columnas
                double argentina = row.getCell(1).getNumericCellValue();
                double brazil = row.getCell(2).getNumericCellValue();
                double chile = row.getCell(3).getNumericCellValue();
                double peru = row.getCell(4).getNumericCellValue();
                double total = row.getCell(5).getNumericCellValue();

                // Imprimir los valores procesados en el formato correcto
                System.out.printf("%s, %.0f, %.0f, %.0f, %.0f, %.0f%n", dateStr, argentina, brazil, chile, peru, total);
            }
        }

        // Cerrar recursos
        file.close();
        workbook.close();
    }
}
