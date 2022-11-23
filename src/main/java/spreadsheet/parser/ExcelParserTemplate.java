package com.tp.asset_ap.spreadsheet.parser;

import com.tp.asset_ap.exception.BadRequestException;
import com.tp.asset_ap.model.dto.excel.ExcelSheetDTO;
import com.tp.asset_ap.util.TPStringUtils;
import com.tp.asset_ap.util.TimeUtils;

import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.InvocationTargetException;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.time.ZoneOffset;
import java.util.Date;

import org.apache.commons.lang3.EnumUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.web.multipart.MultipartFile;

public abstract class ExcelParserTemplate<T> {

    protected final String ASSET_CLASS = "公規資產";
    private final Logger logger = LoggerFactory.getLogger(this.getClass());
    protected final ExcelSheetDTO<T> sheetData = new ExcelSheetDTO<>();

    public void parse(MultipartFile file, int sheetNum)
        throws InvocationTargetException, IllegalAccessException, ParseException {
        Workbook wb = getWorkbook(file);
        parseExcel(wb, sheetNum);
    }

    public void parse(Workbook wb, int sheetNum)
        throws InvocationTargetException, IllegalAccessException, ParseException {
        parseExcel(wb, sheetNum);
    }

    public ExcelSheetDTO<T> getSheetData() {
        return sheetData;
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

    private void parseExcel(Workbook workbook, int sheetNum)
        throws InvocationTargetException, IllegalAccessException, ParseException {
        Sheet sheet = workbook.getSheetAt(sheetNum);
        int firstRowNum = sheet.getFirstRowNum();
        int rowEnd = sheet.getPhysicalNumberOfRows();
        Row firstRow = sheet.getRow(firstRowNum);
        if (null == firstRow) {
            throw new BadRequestException("解析Excel失敗");
        }
        sheetData.setData(parseEachRowData(firstRowNum, rowEnd, sheet));
    }

    protected Cell getCell(Row row, int cellNum) {
        return row.getCell(cellNum);
    }

    protected boolean cellIsEmpty(Cell cell) {
        return cell == null || cell.getCellType() == CellType.BLANK;
    }

    protected String convertCellValueToString(Cell cell) {
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

    protected Integer convertCellValueToInteger(Cell cell) {
        if (cellIsEmpty(cell)) {
            return null;
        }
        Integer returnValue = null;
        switch (cell.getCellType()) {
            case NUMERIC:
                Double value = cell.getNumericCellValue();
                returnValue = value.intValue();
                break;
            case STRING:
                returnValue = TPStringUtils.toInteger(cell.getStringCellValue());
                break;
        }
        return returnValue;
    }

    protected Long convertCellValueToTimeStamp(Cell cell) {
        if (cellIsEmpty(cell)) {
            return null;
        }
        Date date = null;
        switch (cell.getCellType()) {
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    date = cell.getDateCellValue();
                }
                break;
            case STRING:
                date = TimeUtils.toDate(cell.getStringCellValue());
                break;
        }
        return TimeUtils.toUTCMilliseconds(date.toInstant().atOffset(ZoneOffset.UTC));
    }

    protected boolean isValidDateFormat(Cell cell) {
        if (cellIsEmpty(cell)) {
            return true;
        }
        if (cell.getCellType().equals(CellType.STRING)) {
            return TimeUtils.isValidDateString(cell.getStringCellValue());
        }
        return true;
    }

    protected boolean isValidInteger(Cell cell) {
        if (cellIsEmpty(cell)) {
            return true;
        }

        if (cell.getCellType().equals(CellType.STRING)) {
            return TPStringUtils.isNumeric(cell.getStringCellValue());
        }
        return true;
    }

    protected Boolean convertCellValueToStandardSpecification(Cell cell) {
        String standardSpecification = convertCellValueToString(cell);
        return ASSET_CLASS.equals(standardSpecification);
    }

    protected boolean isRowEmpty(Row row) {
        if (row == null) {
            return true;
        }
        if (row.getLastCellNum() <= 0) {
            return true;
        }
        for (int cellNum = row.getFirstCellNum(); cellNum < row.getLastCellNum(); cellNum++) {
            Cell cell = row.getCell(cellNum);
            if (!cellIsEmpty(cell)) {
                return false;
            }
        }
        return true;
    }

    protected abstract T parseEachRowData(int rowStart, int rowEnd, Sheet sheet)
        throws ParseException;

}

