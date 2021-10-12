package com.xtesoft.samples.excelreader.services;

import com.xtesoft.samples.excelreader.entities.Persona;
import com.xtesoft.samples.excelreader.repositories.PersonaRepository;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

@Service
public class ExcelService {
    private static final String DATETIME_FORMAT = "d/m/yyyy h:mm";
    private static final String DATE_FORMAT = "d/m/yyyy";
    private static final String PERCENTAGE_FORMAT = "0.00%";
    public static final String EXCEL = "registro.xlsx";

    @Autowired
    private PersonaRepository personaRepository;

    @Value("${excel.service.templates}")
    private String excel_service_templates;

    private XSSFCell createCell(XSSFRow row, int index, CellType cellType) {
        XSSFCell c = row.createCell(index);
        c.setCellType(cellType);
        return c;
    }

    private Cell createCell(Row row, int index, CellType cellType) {
        Cell c = row.createCell(index);
        if (cellType.equals(CellType.BLANK)) {
            c.setBlank();
        }
        return c;
    }

    private XSSFCell createCell(XSSFRow row, int index) {
        XSSFCell c = row.createCell(index);
        return c;
    }

    private Cell createCell(Row row, int index) {
        Cell c = row.createCell(index);
        return c;
    }

    private Cell createStringCell(Row row, int index) {
        return createCell(row, index, CellType.STRING);
    }

    private XSSFCell createStringCell(XSSFRow row, int index) {
        return createCell(row, index, CellType.STRING);
    }

    private Cell createNumericCell(Row row, int index) {
        return createCell(row, index, CellType.NUMERIC);
    }

    private XSSFCell createNumericCell(XSSFRow row, int index) {
        return createCell(row, index, CellType.NUMERIC);
    }

    private Cell createBlankCell(Row row, int index) {
        return createCell(row, index, CellType.BLANK);
    }

    private XSSFCell createBlankCell(XSSFRow row, int index) {
        return createCell(row, index, CellType.BLANK);
    }

    private Cell createDateCell(Row row, int index, Workbook wb, String format) {
        CreationHelper createHelper = wb.getCreationHelper();
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setDataFormat(
                createHelper.createDataFormat().getFormat(format));
        Cell datecell = createCell(row, index);
        datecell.setCellStyle(cellStyle);
        return datecell;
    }

    private XSSFCell createDateCell(XSSFRow row, int index, XSSFWorkbook wb, String format) {
        XSSFCreationHelper createHelper = wb.getCreationHelper();
        XSSFCellStyle cellStyle = wb.createCellStyle();
        XSSFDataFormat cformat = wb.createDataFormat();
        cellStyle.setDataFormat(cformat.getFormat(format));
        XSSFCell dateCell = createCell(row, index);
        dateCell.setCellStyle(cellStyle);
        return dateCell;
    }

    private Cell createPercentageCell(Row row, int index, Workbook wb) {
        CreationHelper createHelper = wb.getCreationHelper();
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setDataFormat(
                createHelper.createDataFormat().getFormat(PERCENTAGE_FORMAT));
        Cell percentageCell = createCell(row, index);
        percentageCell.setCellStyle(cellStyle);
        return percentageCell;
    }

    private XSSFCell createPercentageCell(XSSFRow row, int index, XSSFWorkbook wb) {
        XSSFCreationHelper createHelper = wb.getCreationHelper();
        XSSFCellStyle cellStyle = wb.createCellStyle();
        XSSFDataFormat cformat = wb.createDataFormat();
        cellStyle.setDataFormat(cformat.getFormat(PERCENTAGE_FORMAT));
        XSSFCell percentageCell = createCell(row, index);
        percentageCell.setCellStyle(cellStyle);
        return percentageCell;
    }

    private void setCellStringValue(Row row, int index, String value) {
        if (value != null && !value.equals("") && !value.equals(" ")) {
            createStringCell(row, index).setCellValue(value);
        } else {
            createBlankCell(row, index).setCellValue("");
        }
    }

    private void setCellStringValue(XSSFRow row, int index, String value) {
        if (value != null && !value.equals("") && !value.equals(" ")) {
            createStringCell(row, index).setCellValue(value);
        } else {
            createBlankCell(row, index).setCellValue("");
        }
    }

    private void setCellIntegerValue(Row row, int index, Integer value) {
        if (value != null) {
            createNumericCell(row, index).setCellValue(value);
        } else {
            createBlankCell(row, index).setCellValue("");
        }
    }

    private void setCellIntegerValue(XSSFRow row, int index, Integer value) {
        if (value != null) {
            createNumericCell(row, index).setCellValue(value);
        } else {
            createBlankCell(row, index).setCellValue("");
        }
    }

    private void setCellBlankValue(Row row, int index) {
        createBlankCell(row, index).setCellValue("");
    }

    private void setCellBlankValue(XSSFRow row, int index) {
        createBlankCell(row, index).setCellValue("");
    }

    private XSSFCell getCellFormula(XSSFRow row, int index, String formula) {
        XSSFCell c = row.createCell(index);

        c.setCellFormula(formula);
        return c;
    }

    private void setCellDoubleValue(Row row, int index, Double value) {
        if (value != null) {
            createNumericCell(row, index).setCellValue(value);
        } else {
            createBlankCell(row, index).setCellValue("");
        }
    }

    private void setCellDoubleValue(XSSFRow row, int index, Double value) {
        if (value != null) {
            createNumericCell(row, index).setCellValue(value);
        } else {
            createBlankCell(row, index).setCellValue("");
        }
    }

    private void setCellPercentageValue(Row row, int index, Workbook wb, Double value) {
        if (value != null) {
            createPercentageCell(row, index, wb).setCellValue(value);
        } else {
            createBlankCell(row, index).setCellValue("");
        }
    }

    private void setCellPercentageValue(XSSFRow row, int index, Workbook wb, Double value) {
        if (value != null) {
            createPercentageCell(row, index, wb).setCellValue(value);
        } else {
            createBlankCell(row, index).setCellValue("");
        }
    }

    private void setCellDateValue(Row row, int index, Workbook wb, Date value, String format) {
        if (value != null) {
            createDateCell(row, index, wb, format).setCellValue(value);
        } else {
            createBlankCell(row, index).setCellValue("");
        }
    }

    private void setCellDateValue(XSSFRow row, int index, Workbook wb, Date value, String format) {
        if (value != null) {
            createDateCell(row, index, wb, format).setCellValue(value);
        } else {
            createBlankCell(row, index).setCellValue("");
        }
    }

    public List<Persona> readXLSXCalificacionFile() {
        List<Persona> personas = new ArrayList<>();
        File xlsx = new File(excel_service_templates, EXCEL);
        try (InputStream inp = new FileInputStream(xlsx)) {
            Workbook wb = WorkbookFactory.create(inp);
            Sheet sheet = wb.getSheetAt(0);
            Iterator<Row> iterator = sheet.iterator();
            while (iterator.hasNext()) {
                Row currentRow = iterator.next();
                Iterator<Cell> cellIterator = currentRow.cellIterator();
                while (cellIterator.hasNext()) {
                    Cell currentCell = cellIterator.next();
                    switch (currentCell.getCellType()) {
                        case STRING:
                            System.out.print(currentCell.getStringCellValue());
                            break;
                        case BOOLEAN:
                            System.out.print(currentCell.getBooleanCellValue());
                            break;
                        case NUMERIC:
                            System.out.print(currentCell.getNumericCellValue());
                            break;
                        default:
                            break;
                    }
                    System.out.print(" | ");
                }
                System.out.println();
            }
            // Closing the workbook and input stream
            wb.close();
            //inp.close();
        } catch (Exception e) {
            System.err.println("Error reading the XLSX File: " + xlsx.getPath());
        }
        return personas;
    }

    private Object getValue(Cell cell) {
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return String.valueOf((int) cell.getNumericCellValue());
            case BOOLEAN:
                return cell.getBooleanCellValue();
            case ERROR:
                return cell.getErrorCellValue();
            case FORMULA:
                return cell.getCellFormula();
            case BLANK:
                return null;
        }
        return null;
    }

    private Long getLongValue(Cell cell) {
        return Long.valueOf(
                (long) cell.getNumericCellValue()
        );
    }

    private String getStringValue(Cell cell) {
        return cell.getStringCellValue();
    }

    public void saveToDatabase() {
        List<Persona> personas = readXLSXCalificacionFile1();
        personas.stream().forEach(p ->
                personaRepository.save(p)
        );
    }

    public List<Persona> readXLSXCalificacionFile1() {
        List<Persona> personas = new ArrayList<>();
        File xlsx = new File(excel_service_templates, EXCEL);
        try (InputStream inp = new FileInputStream(xlsx)) {
            Workbook wb = WorkbookFactory.create(inp);
            Sheet sheet = wb.getSheetAt(0);
            int startRow = 1;
            for (int rowIndex = startRow; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                if (row != null) {
                    Persona currentPerson = new Persona();
                    Cell cell0 = row.getCell(0); //id
                    if (cell0 != null) {
                        currentPerson.setId(getLongValue(cell0));
                    }
                    Cell cell1 = row.getCell(1); //nombres
                    if (cell1 != null) {
                        currentPerson.setNombres(getStringValue(cell1));
                    }
                    Cell cell2 = row.getCell(2); //apellidos
                    if (cell2 != null) {
                        currentPerson.setApellidos(getStringValue(cell2));
                    }
                    Cell cell3 = row.getCell(3); //fecha de nacimiento
                    if (cell3 != null) {
                        currentPerson.setFechaNacimiento(cell3.getDateCellValue());
                    }

                    Cell cell4 = row.getCell(4);
                    if (cell4 != null) {
                        currentPerson.setSalario(BigDecimal.valueOf(cell4.getNumericCellValue()));
                    }

                    personas.add(currentPerson);

                }
            }

            // Closing the workbook and input stream
            wb.close();
            //inp.close();
        } catch (Exception e) {
            System.err.println("Error reading the XLSX File: " + xlsx.getPath());
        }
        return personas;
    }

}
