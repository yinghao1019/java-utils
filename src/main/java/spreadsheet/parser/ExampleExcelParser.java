package com.tp.asset_ap.spreadsheet.parser;

import com.tp.asset_ap.exception.BadRequestException;
import com.tp.asset_ap.model.ExcelExampleDataDTO;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.web.multipart.MultipartFile;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.InvocationTargetException;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.List;

import static org.apache.poi.xssf.usermodel.XSSFWorkbookType.XLSX;

public class ExampleExcelParser {
    private final Logger logger = LoggerFactory.getLogger(this.getClass());
    private final List<ExcelExampleDataDTO> dataList = new ArrayList<>();

    public void parse(MultipartFile file) throws InvocationTargetException, IllegalAccessException {
        Workbook wb = getWorkbook(file);
        parseExcel(wb);
    }

    public List<ExcelExampleDataDTO> getDataList() {
        return dataList;
    }

    private Workbook getWorkbook(MultipartFile file) {
        Workbook wb;
        if (file == null) {
            return null;
        }

        try (InputStream is = file.getInputStream()) {
            wb = new XSSFWorkbook(is);
            return wb;
        } catch (IOException e) {
            logger.info(e.getMessage());
        }
        return null;
    }

    private void parseExcel(Workbook workbook) throws InvocationTargetException, IllegalAccessException {
        // 遍歷每一個sheet
        for (int sheetNum = 0; sheetNum < workbook.getNumberOfSheets(); sheetNum++) {
            Sheet sheet = workbook.getSheetAt(sheetNum);

            if (sheet == null) {
                continue;
            }

            int firstRowNum = sheet.getFirstRowNum();
            Row firstRow = sheet.getRow(firstRowNum);
            if (null == firstRow) {
                throw new BadRequestException("解析Excel失敗");
            }

            int rowStart = firstRowNum + 1;
            int rowEnd = sheet.getPhysicalNumberOfRows();

            parseEachRowData(rowStart, rowEnd, sheet);
        }
    }

    private void parseEachRowData(int rowStart, int rowEnd, Sheet sheet)
        throws InvocationTargetException, IllegalAccessException {
        for (int rowNum = rowStart; rowNum < rowEnd; rowNum++) {
            Row row = sheet.getRow(rowNum);
            if (null == row) {
                continue;
            }

            ExcelExampleDataDTO dto = convertRowToPDTO(row);
            dataList.add(dto);
        }
    }

    private ExcelExampleDataDTO convertRowToPDTO(Row row) {
        ExcelExampleDataDTO excelData = new ExcelExampleDataDTO();
        int cellNum = 0;
        // 讀取識別碼
        String id = convertCellValueToString(getCell(row, cellNum++));
        excelData.setId(id);
        // 讀取年度號
        String name = convertCellValueToString(getCell(row, cellNum++));
        excelData.setName(name);

        return excelData;
    }

    private Cell getCell(Row row, int cellNum) {
        return row.getCell(cellNum);
    }

    private String convertCellValueToString(Cell cell) {
        if (cell == null) {
            return null;
        }
        String returnValue = null;
        switch (cell.getCellType()) {
            case NUMERIC:
                Double doubleValue = cell.getNumericCellValue();
                DecimalFormat df = new DecimalFormat("0");
                returnValue = df.format(doubleValue);
                break;
            case STRING:
                returnValue = cell.getStringCellValue();
                break;
            default:
                break;
        }
        return returnValue;
    }
}
