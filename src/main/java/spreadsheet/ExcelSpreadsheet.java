package com.tp.asset_ap.spreadsheet;

import com.tp.asset_ap.model.dto.excel.ExcelComputeDateDTO;
import com.tp.asset_ap.util.TimeUtils;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.time.OffsetDateTime;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.reflect.FieldUtils;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.core.io.Resource;
import org.springframework.web.multipart.MultipartFile;
import com.tp.asset_ap.constant.MediaTypes;
import com.tp.asset_ap.exception.BadRequestException;
import com.tp.asset_ap.exception.InternalServerErrorException;
import com.tp.asset_ap.report.SheetType;
import com.tp.asset_ap.util.TPStringUtils;

public class ExcelSpreadsheet implements Spreadsheet {

    private static final Logger LOGGER = LoggerFactory.getLogger(ExcelSpreadsheet.class);
    private static final String DEFAULT_SHEET_NAME = "Sheet1";
    private final Workbook workbook;
    private final Map<SpreadsheetStyle, CellStyle> styleMap = new HashMap<>();
    private Sheet workingSheet;
    private Row workingRow;
    private Cell workingCell;

    public ExcelSpreadsheet() {
        workbook = new SXSSFWorkbook();
    }

    public ExcelSpreadsheet(int rowAccessWindowSize) {
        workbook = new SXSSFWorkbook(rowAccessWindowSize);
    }

    public ExcelSpreadsheet(MultipartFile excelFile, ExcelType excelType) {
        try {
            XSSFWorkbook xssfWorkbook =
                (XSSFWorkbook) WorkbookFactory.create(excelFile.getInputStream());
            switch (excelType) {
                case XSS:
                    workbook = xssfWorkbook;
                    break;
                case SXSS:
                default:
                    workbook = new SXSSFWorkbook(xssfWorkbook, SXSSFWorkbook.DEFAULT_WINDOW_SIZE);
                    break;
            }
        } catch (IOException ex) {
            throw new InternalServerErrorException(ex);
        }
    }

    public ExcelSpreadsheet(File file, ExcelType excelType) {
        try {
            XSSFWorkbook xssfWorkbook = (XSSFWorkbook) WorkbookFactory.create(file);
            switch (excelType) {
                case XSS:
                    workbook = xssfWorkbook;
                    break;
                case SXSS:
                default:
                    workbook = new SXSSFWorkbook(xssfWorkbook, SXSSFWorkbook.DEFAULT_WINDOW_SIZE);
                    break;
            }
        } catch (EncryptedDocumentException | IOException ex) {
            throw new InternalServerErrorException(ex);
        }
    }

    public ExcelSpreadsheet(InputStream is, int rowAccessWindowSize) {
        try {
            XSSFWorkbook xssfWorkbook = (XSSFWorkbook) WorkbookFactory.create(is);
            workbook = new SXSSFWorkbook(xssfWorkbook, rowAccessWindowSize);
        } catch (EncryptedDocumentException | IOException ex) {
            throw new InternalServerErrorException(ex);
        }
    }

    public static ExcelSpreadsheet createWorkbook() {
        return createWorkbook(DEFAULT_SHEET_NAME);
    }

    public static ExcelSpreadsheet createWorkbook(String sheetName) {
        ExcelSpreadsheet xlsx = new ExcelSpreadsheet();
        xlsx.createSheet(sheetName);
        return xlsx;
    }

    public static ExcelSpreadsheet createWorkbook(int rowAccessWindowSize) {
        ExcelSpreadsheet xlsx = new ExcelSpreadsheet(rowAccessWindowSize);
        xlsx.createSheet(DEFAULT_SHEET_NAME);
        return xlsx;
    }

    // 會修改原檔案
    public static ExcelSpreadsheet loadWorkbook(File excelFile) {
        return loadWorkbook(excelFile, ExcelType.SXSS);
    }

    // 會修改原檔案
    public static ExcelSpreadsheet loadWorkbook(File excelFile, ExcelType excelType) {
        return new ExcelSpreadsheet(excelFile, excelType);
    }

    // 不會修改原檔案
    public static ExcelSpreadsheet loadWorkbook(InputStream input) {
        return new ExcelSpreadsheet(input, SXSSFWorkbook.DEFAULT_WINDOW_SIZE);
    }

    // 不會修改原檔案 自訂AccessWindowSize
    public static ExcelSpreadsheet loadWorkbook(InputStream input, int rowAccessWindowSize) {
        return new ExcelSpreadsheet(input, rowAccessWindowSize);
    }

    /*
     * cell style
     */
    public static SpreadsheetStyle cellCenterStyle() {
        SpreadsheetStyle style = SpreadsheetStyle.builder()
            .borderPosition(SpreadsheetBorderPosition.builder().build()).build();
        style.setHAlign(TpHorizontalAlignment.CENTER);
        style.setVAlign(TpVerticalAlignment.CENTER);
        style.setWrapText(true);
        style.setFontStyle(cellFont());
        return style;
    }

    public static SpreadsheetStyle cellCenterStyle(SpreadsheetFontStyle fontStyle) {
        SpreadsheetStyle style = SpreadsheetStyle.builder()
            .borderPosition(SpreadsheetBorderPosition.builder().build()).build();
        style.setHAlign(TpHorizontalAlignment.CENTER);
        style.setVAlign(TpVerticalAlignment.CENTER);
        style.setWrapText(true);
        style.setFontStyle(fontStyle);
        return style;
    }

    public static SpreadsheetStyle cellLeftStyle() {
        SpreadsheetStyle style = SpreadsheetStyle.builder()
            .borderPosition(SpreadsheetBorderPosition.builder().build()).build();
        style.setHAlign(TpHorizontalAlignment.LEFT);
        style.setVAlign(TpVerticalAlignment.CENTER);
        style.setWrapText(true);
        style.setFontStyle(cellFont());

        return style;
    }

    public static SpreadsheetStyle cellRightStyle() {
        SpreadsheetStyle style = SpreadsheetStyle.builder()
            .borderPosition(SpreadsheetBorderPosition.builder().build()).build();
        style.setHAlign(TpHorizontalAlignment.RIGHT);
        style.setVAlign(TpVerticalAlignment.CENTER);
        style.setWrapText(true);
        style.setFontStyle(cellFont());

        return style;
    }

    public static SpreadsheetStyle buildHeaderStyle() {
        return SpreadsheetStyle.builder().hAlign(TpHorizontalAlignment.CENTER)
            .vAlign(TpVerticalAlignment.CENTER).wrapText(true)
            .backGroundColor(new XSSFColor(new java.awt.Color(63, 63, 63), new DefaultIndexedColorMap()))
            .fontStyle(
                SpreadsheetFontStyle.builder().fontSize(10).fontName("Arial").bold(true).color(IndexedColors.WHITE)
                    .build())
            .build();
    }

    public static SpreadsheetStyle cellLeftBoldStyle() {
        SpreadsheetStyle style = SpreadsheetStyle.builder()
            .borderPosition(SpreadsheetBorderPosition.builder().build()).build();
        style.setHAlign(TpHorizontalAlignment.LEFT);
        style.setVAlign(TpVerticalAlignment.CENTER);
        style.setWrapText(true);
        style.setFontStyle(cellFontBold());

        return style;
    }

    public static SpreadsheetFontStyle cellFont() {
        SpreadsheetFontStyle font = SpreadsheetFontStyle.builder().build();
        font.setFontName("標楷體");
        font.setFontSize(12);
        font.setBold(false);
        font.setItalic(false);
        return font;
    }

    public static SpreadsheetFontStyle cellFontBold() {
        SpreadsheetFontStyle font = SpreadsheetFontStyle.builder().build();
        font.setFontName("標楷體");
        font.setFontSize(14);
        font.setBold(true);
        font.setItalic(false);
        return font;
    }

    public static String safeGetStringCellValue(Cell cell) {
        // 有可能因為格式問題，導致讀取不到資料，所以先轉成 String 格式，讀完再轉回原格式
        CellType originalCellType = cell.getCellType();
        String formula = originalCellType == CellType.FORMULA ? cell.getCellFormula() : null;
        try {
            cell.setCellType(CellType.STRING);
            return cell.getStringCellValue();
        } catch (Exception ex) {
            return "";
        } finally {
            if (originalCellType == CellType.FORMULA) {
                cell.setCellFormula(formula);
            } else {
                cell.setCellType(originalCellType);
            }
        }
    }

    public static double safeGetNumericCellValue(Cell cell) {
        try {
            return cell.getNumericCellValue();
        } catch (Exception ex) {
            return 0.0;
        }
    }

    public <T> ExcelSpreadsheet generateSheet(List<T> dataList) throws IllegalAccessException {
        return this.generateSheet(dataList, 0, 0, null, true, 1);
    }

    public <T> ExcelSpreadsheet generateSheet(
        List<T> dataList, int startRowIndex,
        int startColIndex, String title, boolean showColumnHeader, int titleShiftRowSize)
        throws IllegalAccessException {
        return this.generateSheet(dataList, startRowIndex, startColIndex, title, showColumnHeader,
            titleShiftRowSize, false);
    }

    public <T> ExcelSpreadsheet generateSheet(
        List<T> dataList, int startRowIndex,
        int startColIndex, String title, boolean showColumnHeader, int titleShiftRowSize,
        boolean hasCreateDate) throws IllegalAccessException {
        if (CollectionUtils.isEmpty(dataList)) {
            throw new IllegalArgumentException("No data to generate!");
        }
        Map<Integer, ExcelColumn> map = parseExcelColumn(dataList.get(0).getClass());
        if (map.size() == 0) {
            throw new IllegalArgumentException(dataList.get(0).getClass().getName()
                + " have no fields annotated with @ExcelColumn!");
        }
        int shiftRowSize = hasCreateDate ? 1 + titleShiftRowSize : titleShiftRowSize;
        buildTitle(map, startRowIndex, startColIndex, title, hasCreateDate);
        if (showColumnHeader) {
            shiftRowSize++;
            buildHeader(map, startRowIndex + shiftRowSize, startColIndex);
        }
        shiftRowSize++;
        buildContent(dataList, map, startRowIndex + shiftRowSize);
        setColumnWidth(map, startColIndex);
        return this;
    }

    public <T> ExcelSpreadsheet generateRowSpanSheet(
        Map<Integer, List<T>> dataMap,
        int startRowIndex, int startColIndex, int setIdColIndex) throws IllegalAccessException {
        int rowIndex = startRowIndex;
        for (Map.Entry<Integer, List<T>> data : dataMap.entrySet()) {
            List<T> dataList = data.getValue();
            generateSheet(dataList, rowIndex, startColIndex, null, false, 0);
            int mergeRowCount = dataList.size();
            //            合併欄位的列(僅有一筆資料時不用合併)
            if (mergeRowCount > 1) {
                for (int colNum = 0; colNum < setIdColIndex; colNum++) {
                    mergeRows(rowIndex + 1, mergeRowCount, colNum);
                }
            }
            //            設置子項目的編號
            Field field = dataList.get(0).getClass().getDeclaredFields()[0];
            for (int dataId = 1; dataId < mergeRowCount + 1; dataId++) {
                setCellValue(rowIndex + dataId, setIdColIndex, dataId, getStyleByField(field));
            }
            rowIndex += mergeRowCount;
        }
        return this;
    }

    public ExcelSpreadsheet generateDailyRentSheet(
        Map<Integer, String> departmentCol,
        Map<Integer, String> assetTypeRow,
        Map<Integer, Map<Integer, BigDecimal>> dailyRent,
        String title,
        OffsetDateTime createDate,
        ExcelComputeDateDTO dateDTO) {
        buildDailyRentTitle(departmentCol.size() + 2, title, createDate, dateDTO);
        int startRowIndex = 4;
        int startColIndex = 0;
        buildDailyRentHeader(departmentCol, startRowIndex, startColIndex);
        startRowIndex++;
        buildSideBar(assetTypeRow, startRowIndex, startColIndex);
        startColIndex++;

        Map<Integer, BigDecimal> assetTotalMap = new HashMap<>();
        SpreadsheetStyle valueStyle = cellCenterStyle();
        SpreadsheetStyle amountStyle = cellCenterStyle();
        amountStyle.setBackGroundColor(new XSSFColor(new java.awt.Color(112, 173, 71), new DefaultIndexedColorMap()));
        //        設定日租金
        for (Map.Entry<Integer, String> department : departmentCol.entrySet()) {
            int rowIndex = startRowIndex;
            BigDecimal departmentTotal = new BigDecimal(0);
            Map<Integer, BigDecimal> assetRent = dailyRent.get(department.getKey());
            for (Map.Entry<Integer, String> assetType : assetTypeRow.entrySet()) {
                BigDecimal rentValue = assetRent.get(assetType.getKey());
                departmentTotal = departmentTotal.add(rentValue);
                setCellValue(rowIndex, startColIndex, rentValue, valueStyle);
                //               加總各資產項目的總和
                BigDecimal assetTotal = assetTotalMap.getOrDefault(assetType.getKey(), new BigDecimal(0));
                assetTotalMap.put(assetType.getKey(), assetTotal.add(rentValue));
                rowIndex++;
            }
            //            設置處別總計
            setCellValue(rowIndex, 0, "總計", amountStyle);
            setCellValue(rowIndex, startColIndex, departmentTotal, amountStyle);
            startColIndex++;
        }
        //        設置各資產小計
        BigDecimal total = new BigDecimal(0);
        int rowIndex = startRowIndex;
        for (Map.Entry<Integer, String> assetType : assetTypeRow.entrySet()) {
            setCellValue(rowIndex, startColIndex, assetTotalMap.get(assetType.getKey()), valueStyle);
            total = total.add(assetTotalMap.get(assetType.getKey()));
            rowIndex++;
        }
        setCellValue(rowIndex, startColIndex, total, amountStyle);
        return this;
    }

    public void close() throws IOException {
        workbook.close();
    }

    /**
     * sheet
     */
    public ExcelSpreadsheet createSheet(String sheetName) {
        workingSheet = workbook.createSheet(sheetName);
        return this;
    }

    public ExcelSpreadsheet getSheetAt(int index) {
        workingSheet = workbook.getSheetAt(index);
        if (workingSheet == null) {
            createSheet("sheet" + index);
        }
        return this;
    }

    public ExcelSpreadsheet useSheet(String sheetName) {
        workingSheet = workbook.getSheet(sheetName);
        if (workingSheet == null) {
            createSheet(sheetName);
        }
        return this;
    }

    public void removeSheet(int sheetNum) {
        if (this.isSheetExist(sheetNum)) {
            workbook.removeSheetAt(sheetNum);
        }
    }

    public SXSSFSheet cloneSheet(int sheetNum) {
        return (SXSSFSheet) workbook.cloneSheet(sheetNum);
    }

    public ExcelSpreadsheet setSheetColumnWidth(List<Integer> columnWidth) {
        int columnNum = 0;
        for (Integer column : columnWidth) {
            workingSheet.setColumnWidth(columnNum, column * 4 * 256);
            columnNum++;
        }
        return this;
    }

    public boolean isSheetExist(String sheetName) {
        workingSheet = workbook.getSheet(sheetName);
        return workingSheet != null;
    }

    public boolean isSheetExist(int sheetNum) {
        return workbook.getSheetAt(sheetNum) != null;
    }

    /**
     * row
     */
    public ExcelSpreadsheet createRow(int rowIndex) {
        if (workingSheet == null) {
            createSheet(DEFAULT_SHEET_NAME);
        }
        workingRow = workingSheet.createRow(rowIndex);
        return this;
    }

    public ExcelSpreadsheet useRow(int rowIndex) {
        workingRow = workingSheet.getRow(rowIndex);
        if (workingRow == null) {
            createRow(rowIndex);
        }
        return this;
    }

    /**
     * cell
     */
    public ExcelSpreadsheet createCell(int columnIndex) {
        if (workingRow == null) {
            throw new BadRequestException("workingRow must not be null");
        }
        workingCell = workingRow.createCell(columnIndex);
        return this;
    }

    public ExcelSpreadsheet useCell(int columnIndex) {
        workingCell = workingRow.getCell(columnIndex);
        if (workingCell == null) {
            createCell(columnIndex);
        }
        return this;
    }

    public ExcelSpreadsheet setCellValue(int rowIndex, int columnIndex, Object value) {
        setCellValue(rowIndex, columnIndex, value, null);
        return this;
    }

    public ExcelSpreadsheet setCellValue(
        int rowIndex, int columnIndex, Object value,
        SpreadsheetStyle style) {
        workingRow = workingSheet.getRow(rowIndex);
        if (workingRow == null) {
            createRow(rowIndex);
        }
        workingCell = workingRow.getCell(columnIndex);
        if (workingCell == null) {
            createCell(columnIndex);
        }

        this.setCellValueByType(value);

        buildAndSetCellStyle(style);
        return this;
    }

    public ExcelSpreadsheet setCellValueByType(Object value) {
        if (value == null) {
            workingCell.setCellValue("");
        } else if (value instanceof Integer) {
            workingCell.setCellValue((int) value);
        } else if (value instanceof Long) {
            workingCell.setCellValue((long) value);
        } else {
            workingCell.setCellValue(value.toString());
        }

        return this;
    }

    public List<List<String>> readFields() {
        return readFields(DEFAULT_SHEET_NAME);
    }

    public List<List<String>> readFields(String sheetName) {
        return readFields(sheetName, null);
    }

    public List<List<String>> readFields(String sheetName, Integer maxReadCellNum) {
        useSheet(sheetName);
        List<List<String>> data = new ArrayList<>();
        int rowNum = workingSheet.getLastRowNum();
        for (int i = 0; i <= rowNum; i++) {
            List<String> cellData = new ArrayList<>();
            workingRow = workingSheet.getRow(i);
            if (workingRow != null) {
                int cellNum = workingRow.getLastCellNum() - 1;
                cellNum = maxReadCellNum == null || cellNum < maxReadCellNum ? cellNum
                    : maxReadCellNum;
                for (int j = 0; j <= cellNum; j++) {
                    String val = "";
                    workingCell = workingRow.getCell(j);
                    if (workingCell != null) {
                        val = getCellValue();
                    }
                    cellData.add(val);
                }
            }
            data.add(cellData);
        }
        return data;
    }

    public String getCellValue() {
        switch (workingCell.getCellType()) {
            case STRING:
                return workingCell.getRichStringCellValue().getString();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(workingCell)) {
                    return workingCell.getDateCellValue().toString();
                } else {
                    return String.valueOf(workingCell.getNumericCellValue());
                }
            case FORMULA:
                return workingCell.getCellFormula();
            case BOOLEAN:
                return String.valueOf(workingCell.getBooleanCellValue());
            case BLANK:
            case ERROR:
            default:
                return StringUtils.EMPTY;
        }
    }

    /**
     * ["yyyy/mm/dd", "yyyy/m/d hh:mm", ...]
     */
    public ExcelSpreadsheet setCellValue(Date value) {
        workingCell.setCellValue(value);
        return this;
    }

    /**
     * ["0.0", "#,##0.0000", ...]
     */
    public ExcelSpreadsheet setCellValue(double value) {
        workingCell.setCellValue(value);
        return this;
    }

    private int getCellValueLength(Cell cell) {
        if (cell == null) {
            return 0;
        }
        switch (cell.getCellType()) {
            case STRING:
                return cell.getRichStringCellValue().getString().length();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString().length();
                } else {
                    return String.valueOf(cell.getNumericCellValue()).length();
                }
            case FORMULA:
                return cell.getCellFormula().length();
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue()).length();
            case BLANK:
            case ERROR:
            default:
                return 0;
        }
    }

    public ExcelSpreadsheet mergeCells(
        int rowStartIndex, int mergedRowCount, int colStartIndex,
        int mergedColumnCount) {
        return mergeCells(rowStartIndex, mergedRowCount, colStartIndex, mergedColumnCount, null,
            false);
    }

    public ExcelSpreadsheet mergeCells(
        int rowStartIndex, int mergedRowCount, int colStartIndex,
        int mergedColumnCount, SpreadsheetStyle style) {
        return mergeCells(rowStartIndex, mergedRowCount, colStartIndex, mergedColumnCount, style,
            false);
    }

    public ExcelSpreadsheet mergeCells(
        int rowStartIndex, int mergedRowCount, int colStartIndex,
        int mergedColumnCount, SpreadsheetStyle style, boolean isWorkingCellAutoRowHeight) {
        int rowEndIndex =
            (mergedRowCount == 0) ? rowStartIndex : (rowStartIndex + mergedRowCount - 1);
        int colEndIndex =
            (mergedColumnCount == 0) ? colStartIndex : (colStartIndex + mergedColumnCount - 1);
        CellRangeAddress range =
            new CellRangeAddress(rowStartIndex, rowEndIndex, colStartIndex, colEndIndex);

        buildAndSetCellStyle(style, range);
        if (mergedRowCount != 1 || mergedColumnCount != 1) {
            workingSheet.addMergedRegion(range);
        }

        if (isWorkingCellAutoRowHeight) {
            int textSize = getCellValueLength(workingCell);
            XSSFFont cellFont = this.getWorkingCellFont();
            if (cellFont != null) {
                int fontSize = cellFont.getFontHeightInPoints();
                setRowHeight(rowStartIndex, this.getColWidthSum(colStartIndex, colEndIndex),
                    textSize, fontSize);
            }
        }
        return this;
    }

    public ExcelSpreadsheet mergeRows(int rowStartIndex, int mergedRowCount, int colStartIndex) {
        return mergeCells(rowStartIndex, mergedRowCount, colStartIndex, 0);
    }

    private XSSFFont getWorkingCellFont() {
        XSSFCellStyle cellStyle = (XSSFCellStyle) workingCell.getCellStyle();
        if (cellStyle != null) {
            return cellStyle.getFont();
        } else {
            return null;
        }
    }

    private int getColWidthSum(int colStartIndex, int colEndIndex) {
        int count = 0;
        if (colEndIndex - colStartIndex >= 0) {
            for (int i = colStartIndex; i < colEndIndex + 1; i++) {
                count += (workingSheet.getColumnWidth(i) / 256);
            }
        } else {
            throw new InternalServerErrorException("合併儲存格不得為負值");
        }
        return count;
    }

    public void setRowHeight(int rowIndex, short height) {
        Row row = workingSheet.getRow(rowIndex);
        if (row == null) {
            return;
        }
        row.setHeight(height);
    }

    private void buildAndSetCellStyle(SpreadsheetStyle style) {
        if (style != null) {
            CellStyle cellStyle = putStyleToMap(style);
            workingCell.setCellStyle(cellStyle);
        }
    }

    private void buildAndSetCellStyle(SpreadsheetStyle style, CellRangeAddress region) {
        if (style != null) {
            CellStyle cellStyle = putStyleToMap(style);
            RegionUtil.setBorderBottom(cellStyle.getBorderBottom(), region, workingSheet);
            RegionUtil.setBorderTop(cellStyle.getBorderTop(), region, workingSheet);
            RegionUtil.setBorderLeft(cellStyle.getBorderLeft(), region, workingSheet);
            RegionUtil.setBorderRight(cellStyle.getBorderRight(), region, workingSheet);
        }
    }

    private CellStyle putStyleToMap(SpreadsheetStyle style) {
        return styleMap.computeIfAbsent(style, this::buildCellStyle);
    }

    public CellStyle buildCellStyle(SpreadsheetStyle style) {
        XSSFCellStyle cellStyle = (XSSFCellStyle) workbook.createCellStyle();

        SpreadsheetFontStyle fontStyle = style.getFontStyle();
        if (fontStyle != null) {
            Font font = workbook.createFont();
            font.setFontName(fontStyle.getFontName());
            font.setFontHeightInPoints((short) fontStyle.getFontSize());
            font.setBold(style.getFontStyle().isBold());
            font.setColor(fontStyle.getColor().getIndex());
            cellStyle.setFont(font);
        }
        cellStyle.setVerticalAlignment(verticalAlignmentTypeConverter(style.getVAlign()));
        cellStyle.setAlignment(horizontalAlignmentTypeConverter(style.getHAlign()));
        cellStyle.setWrapText(style.isWrapText());
        if (style.getBackGroundColor() != null) {
            cellStyle.setFillForegroundColor(style.getBackGroundColor());
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        }
        if (style.getBorderPosition().isBottom()) {
            cellStyle.setBorderBottom(borderConverter(style.getBorderStyle()));
        }
        if (style.getBorderPosition().isLeft()) {
            cellStyle.setBorderLeft(borderConverter(style.getBorderStyle()));
        }
        if (style.getBorderPosition().isRight()) {
            cellStyle.setBorderRight(borderConverter(style.getBorderStyle()));
        }
        if (style.getBorderPosition().isTop()) {
            cellStyle.setBorderTop(borderConverter(style.getBorderStyle()));
        }
        return cellStyle;
    }

    // getDefaultRowHeightInPoints 為15的情況適用
    public void autoSetRowHeight() {
        int rowLastNum = workingSheet.getLastRowNum() + 1;
        for (int rowIndex = 0; rowIndex < rowLastNum; rowIndex++) {
            Row row = workingSheet.getRow(rowIndex);
            if (row == null) {
                continue;
            }
            int cellLastNum = row.getLastCellNum();
            int neededRowsMax = 1;

            for (int cellIndex = 0; cellIndex < cellLastNum; cellIndex++) {
                Cell cell = row.getCell(cellIndex);
                if (cell == null) {
                    continue;
                }
                // 欄寬 *4.8 (計算出來的可能有更好的?)
                double columnWidth =
                    4.8 * (workingSheet.getColumnWidth(cell.getColumnIndex()) / 256D);
                int fontSize =
                    ((XSSFCellStyle) cell.getCellStyle()).getFont().getFontHeightInPoints();
                // 字數*大小
                int chartSize = this.getCellValueLength(cell) * fontSize;

                // 無條件進位
                int neededRows = (int) Math
                    .ceil(Math.ceil(chartSize / columnWidth) * Math.ceil(fontSize / 12.0));

                if (neededRows > neededRowsMax) {
                    neededRowsMax = neededRows;
                }
            }
            float defaultRowHeight = workingSheet.getDefaultRowHeightInPoints();
            row.setHeightInPoints(neededRowsMax * defaultRowHeight);
        }
    }

    public void setRowHeight(int rowIndex, int totalWidth, int textSize, int fontSize) {
        // 欄寬 *4.8 (計算出來的可能有更好的?)
        double columnWidth = 4.8 * totalWidth;
        // 字數*大小
        int chartSize = textSize * fontSize;

        // 無條件進位
        int neededRows =
            (int) Math.ceil(Math.ceil(chartSize / columnWidth) * Math.ceil(fontSize / 12.0));
        float defaultRowHeight = workingSheet.getDefaultRowHeightInPoints();
        if (neededRows * defaultRowHeight > workingSheet.getRow(rowIndex).getHeightInPoints()) {
            workingSheet.getRow(rowIndex).setHeightInPoints(neededRows * defaultRowHeight);
        }
    }

    public ExcelSpreadsheet autoCellWidth() {
        if (workingSheet instanceof SXSSFSheet) {
            ((SXSSFSheet) workingSheet).trackAllColumnsForAutoSizing();
        }
        Row header = workingSheet.getRow(0);
        int cellCnt = header.getLastCellNum();
        for (int i = 0; i < cellCnt; i++) {
            workingSheet.autoSizeColumn(i);
        }
        return this;
    }

    public void autoAllCellWidth(int rowIndex) {
        Row row = workingSheet.getRow(rowIndex);
        if (row != null) {
            int cellCnt = row.getLastCellNum();
            if (workingSheet instanceof SXSSFSheet) {
                ((SXSSFSheet) workingSheet).trackAllColumnsForAutoSizing();
            }
            for (int j = 0; j < cellCnt; j++) {
                workingSheet.autoSizeColumn(j);
            }
        }
    }

    public Resource toResource() throws IOException {
        try (ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();) {
            workbook.write(byteArrayOutputStream);
            return new ByteArrayResource(byteArrayOutputStream.toByteArray());
        }
    }

    public byte[] getBytes() throws IOException {
        try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
            workbook.write(bos);
            return bos.toByteArray();
        }
    }

    public ExcelSpreadsheet exportFile(String path, String fileName) throws IOException {
        final String fileNameWithExtension = appendExtensionIfNot(fileName);
        File file = new File(path, fileNameWithExtension);
        return exportFile(file);
    }

    public ExcelSpreadsheet exportFile(File path, String fileName) throws IOException {
        final String fileNameWithExtension = appendExtensionIfNot(fileName);
        File file = new File(path, fileNameWithExtension);
        return exportFile(file);
    }

    private ExcelSpreadsheet exportFile(File file) throws IOException {
        try (FileOutputStream fileOut = new FileOutputStream(file)) {
            exportFile(fileOut);
        } catch (FileNotFoundException ex) {
            LOGGER.error("No file found with file name: {}", file.getAbsoluteFile(), ex);
        }
        return this;
    }

    private String appendExtensionIfNot(String fileName) {
        if (TPStringUtils.isNullOrEmpty(fileName)) {
            throw new BadRequestException("非法的檔案名稱");
        }
        String extension = SheetType.EXCEL.getExtension();
        if (!fileName.endsWith(extension)) {
            fileName = fileName + extension;
        }
        return fileName;
    }

    public ExcelSpreadsheet exportFile(OutputStream outputStream) throws IOException {
        try {
            workbook.write(outputStream);
        } catch (IOException ex) {
            LOGGER.error(ex.getMessage(), ex);
        } finally {
            outputStream.close();
            workbook.close();
        }
        return this;
    }

    @Override
    public String getExtension() {
        return SheetType.EXCEL.getExtension();
    }

    private <T> Map<Integer, ExcelColumn> parseExcelColumn(Class<T> clazz) {
        Map<Integer, ExcelColumn> map = new HashMap<>();

        Field[] fields = clazz.getDeclaredFields();
        int idx = 0;
        for (Field field : fields) {
            ExcelColumn annotation = field.getAnnotation(ExcelColumn.class);
            if (annotation != null) {
                map.put(idx++, annotation);
            }
        }
        return map;
    }

    private ExcelSpreadsheet buildTitle(
        Map<Integer, ExcelColumn> map, int startRowIndex,
        int startColIndex, String title, boolean hasCreateDate) {
        if (TPStringUtils.isNullOrEmpty(title)) {
            return this;
        }
        int mergedColumnCount = map.entrySet().size();
        SpreadsheetFontStyle titleStyle = new SpreadsheetFontStyle("標楷體", 20, false, false, IndexedColors.BLACK);
        SpreadsheetStyle style = SpreadsheetStyle.builder().build();
        style.setFontStyle(titleStyle);
        style.setHAlign(TpHorizontalAlignment.CENTER);
        style.setWrapText(true);
        mergeCells(startRowIndex, 1, startColIndex, mergedColumnCount, style);
        setCellValue(startRowIndex, startColIndex, title, style);
        setRowHeight(startRowIndex, 20 * mergedColumnCount, title.length(),
            titleStyle.getFontSize());

        if (hasCreateDate) {
            SpreadsheetFontStyle dateFontStyle = new SpreadsheetFontStyle("標楷體", 12, false, false,
                IndexedColors.BLACK);
            SpreadsheetStyle dateStyle = SpreadsheetStyle.builder().build();
            dateStyle.setFontStyle(dateFontStyle);
            dateStyle.setHAlign(TpHorizontalAlignment.RIGHT);
            dateStyle.setWrapText(true);

            mergeCells(startRowIndex + 1, 1, startColIndex, mergedColumnCount, dateStyle);
            OffsetDateTime now = OffsetDateTime.now();
            setCellValue(startRowIndex + 1, startColIndex, "產生日期:" + now.getYear() + "年"
                + now.getMonthValue() + "月" + now.getDayOfMonth() + "日", dateStyle);
        }

        return this;
    }

    private void buildDailyRentTitle(
        int mergedColumnCount, String title, OffsetDateTime createDate,
        ExcelComputeDateDTO dateDTO) {
        int rowIndex = 0;
        int startColIndex = 0;
        SpreadsheetStyle style = cellCenterStyle(
            SpreadsheetFontStyle.builder().fontSize(12).fontName("標楷體").build());
        setCellValue(rowIndex, startColIndex, "昕力資訊股份有限公司", style);
        mergeCells(rowIndex, 0, startColIndex, mergedColumnCount, style);
        rowIndex++;
        setCellValue(rowIndex, startColIndex, title, style);
        mergeCells(rowIndex, 0, startColIndex, mergedColumnCount, style);
        rowIndex++;
        setCellValue(rowIndex, startColIndex, "下載日期", cellLeftStyle());
        setCellValue(rowIndex, startColIndex + 1, TimeUtils.toString(createDate));
        mergeCells(rowIndex, 0, startColIndex + 1, mergedColumnCount - 1, style);
        rowIndex++;
        int colIndex = startColIndex + 1;
        setCellValue(rowIndex, startColIndex, "計算起日");
        setCellValue(rowIndex, colIndex, TimeUtils.toString(dateDTO.getComputeStartDate()));
        int mergeColumn = mergedColumnCount - 2;
        int startDateColRange = mergeColumn % 2 == 0 ? mergeColumn / 2 : mergeColumn % 2;
        mergeCells(rowIndex, 0, colIndex, startDateColRange, style);
        colIndex += startDateColRange;
        setCellValue(rowIndex, colIndex, "計算迄日");
        colIndex++;
        setCellValue(rowIndex, colIndex, TimeUtils.toString(dateDTO.getComputeEndDate()));
        mergeCells(rowIndex, 0, colIndex, mergeColumn - startDateColRange, style);
        //        設置欄位寬度
        for (int index = 0; index < mergedColumnCount; index++) {
            this.workingSheet.setColumnWidth(index, 10 * 256);
        }
    }

    public ExcelSpreadsheet buildDailyRentHeader(
        Map<Integer, String> map, int startRowIndex,
        int startColIndex) {
        SpreadsheetStyle style = buildHeaderStyle();
        int columnIndex = startColIndex + 1;
        int rowIndex = startRowIndex;
        setCellValue(startRowIndex, startColIndex, "資產項目", style);
        for (Map.Entry<Integer, String> entry : map.entrySet()) {
            String colName = entry.getValue();
            setCellValue(rowIndex, columnIndex++, colName, style);
        }
        setCellValue(rowIndex, columnIndex, "小計", style);
        return this;
    }

    public ExcelSpreadsheet buildSideBar(
        Map<Integer, String> map, int startRowIndex,
        int colIndex) {
        SpreadsheetStyle style = cellCenterStyle();
        int rowIndex = startRowIndex;
        for (Map.Entry<Integer, String> entry : map.entrySet()) {
            String colName = entry.getValue();
            setCellValue(rowIndex++, colIndex, colName, style);
        }
        return this;
    }

    private ExcelSpreadsheet buildHeader(
        Map<Integer, ExcelColumn> map, int startRowIndex,
        int startColIndex) {
        int columnIndex = startColIndex;
        Map<Integer, SpreadsheetStyle> headerStyleMap = buildStyleMapByExcelColumn(map);

        for (Map.Entry<Integer, ExcelColumn> entry : map.entrySet()) {
            ExcelColumn annotation = entry.getValue();
            String colName = annotation.colName();
            if (!TPStringUtils.isNullOrEmpty(annotation.mergeGroup())) {
                mergeCells(startRowIndex - annotation.rowSize(), 1, columnIndex,
                    annotation.mergeGroupSize(), headerStyleMap.get(entry.getKey()));
                setCellValue(startRowIndex - annotation.rowSize(), columnIndex,
                    annotation.mergeGroup(), headerStyleMap.get(entry.getKey()));
                setRowHeight(startRowIndex - annotation.rowSize(),
                    annotation.columnWidth() * annotation.mergeGroupSize(),
                    annotation.mergeGroup().length(),
                    headerStyleMap.get(entry.getKey()).getFontStyle().getFontSize());
            }

            if (annotation.rowSize() > 1) {
                // 包含此格往上找size格
                mergeCells(startRowIndex - (annotation.rowSize() - 1), annotation.rowSize(),
                    columnIndex, 1, headerStyleMap.get(entry.getKey()));
                setCellValue(startRowIndex - (annotation.rowSize() - 1), columnIndex++, colName,
                    headerStyleMap.get(entry.getKey()));
            } else {
                setCellValue(startRowIndex, columnIndex++, colName,
                    headerStyleMap.get(entry.getKey()));
            }
        }
        return this;
    }

    private <T> ExcelSpreadsheet buildContent(
        List<T> dataList, Map<Integer, ExcelColumn> map,
        int startRowIndex) throws IllegalAccessException {
        int rowIndex = startRowIndex;
        Map<Integer, SpreadsheetStyle> contentStyleMap = buildStyleMapByExcelColumn(map);

        for (T data : dataList) {
            Field[] fields = data.getClass().getDeclaredFields();
            for (Map.Entry<Integer, ExcelColumn> entry : map.entrySet()) {
                int idx = entry.getKey();
                int colIndex = entry.getValue().colIndex();
                Object value = FieldUtils.readField(fields[idx], data, true);
                setCellValue(rowIndex, colIndex, value, contentStyleMap.get(entry.getKey()));
            }
            rowIndex++;
        }
        return this;
    }

    public <T> ExcelSpreadsheet buildColumnValue(
        List<T> dataList, int startRowIndex,
        int startColIndex) throws IllegalAccessException {
        int rowIndex = startRowIndex;
        SpreadsheetStyle contentStyle = SpreadsheetStyle.builder().build();
        for (T data : dataList) {
            Field[] fields = data.getClass().getDeclaredFields();
            int idx = 0;
            for (Field field : fields) {
                int colIndex = startColIndex + idx;
                Object value = FieldUtils.readField(field, data, true);
                setCellValue(rowIndex, colIndex, value, contentStyle);
                idx++;
            }
            rowIndex++;
        }
        return this;
    }

    private Map<Integer, SpreadsheetStyle> buildStyleMapByExcelColumn(
        Map<Integer, ExcelColumn> map) {
        Map<Integer, SpreadsheetStyle> contentStyleMap = new HashMap<>();
        for (Map.Entry<Integer, ExcelColumn> entry : map.entrySet()) {
            contentStyleMap.put(entry.getKey(), getColumnStyle(entry.getValue()));
        }
        return contentStyleMap;
    }

    private SpreadsheetStyle getStyleByField(Field field) {
        return getColumnStyle(field.getAnnotation(ExcelColumn.class));
    }

    private SpreadsheetStyle getColumnStyle(ExcelColumn excelColumn) {
        SpreadsheetStyle style = SpreadsheetStyle.builder().hAlign(excelColumn.hAlign())
            .vAlign(excelColumn.vAlign()).wrapText(excelColumn.isWrap())
            .fontStyle(SpreadsheetFontStyle.builder()
                .fontName(excelColumn.fontName())
                .fontSize(excelColumn.fontSize()).build())
            .build();
        return style;
    }

    private VerticalAlignment verticalAlignmentTypeConverter(
        TpVerticalAlignment verticalAlignment) {
        switch (verticalAlignment) {
            case TOP:
                return VerticalAlignment.TOP;
            case CENTER:
                return VerticalAlignment.CENTER;
            case BOTTOM:
            default:
                return VerticalAlignment.BOTTOM;
        }
    }

    private HorizontalAlignment horizontalAlignmentTypeConverter(
        TpHorizontalAlignment horizontalAlignment) {
        switch (horizontalAlignment) {
            case LEFT:
                return HorizontalAlignment.LEFT;
            case RIGHT:
                return HorizontalAlignment.RIGHT;
            case CENTER:
                return HorizontalAlignment.CENTER;
            case JUSTIFY:
                return HorizontalAlignment.JUSTIFY;
            case FILL:
                return HorizontalAlignment.FILL;
            case GENERAL:
            default:
                return HorizontalAlignment.GENERAL;
        }
    }

    private BorderStyle borderConverter(TpBorderStyle tpBorderStyle) {
        switch (tpBorderStyle) {
            case THIN:
                return BorderStyle.THIN;
            case MEDIUM:
                return BorderStyle.MEDIUM;
            case THICK:
                return BorderStyle.THICK;
            case DOUBLE:
                return BorderStyle.DOUBLE;
            case NONE:
            default:
                return BorderStyle.NONE;
        }
    }

    private ExcelSpreadsheet setColumnWidth(Map<Integer, ExcelColumn> map, int colIndex) {
        for (Map.Entry<Integer, ExcelColumn> entry : map.entrySet()) {
            ExcelColumn annotation = entry.getValue();
            int colWidth = annotation.columnWidth();
            this.workingSheet.setColumnWidth(entry.getKey() + colIndex, colWidth * 256);
        }

        return this;
    }

    public Workbook getWorkbook() {
        return this.workbook;
    }

    public Sheet getWorkingSheet() {
        return this.workingSheet;
    }

    public void setWorkingCell(Cell workingCell) {
        this.workingCell = workingCell;
    }

    public String getMediaTypeValue() {
        return MediaTypes.APPLICATION_XLSX_VALUE;
    }

    // SXSS：適合大量寫入，但寫過的欄位不能重複寫入，且不能讀取內容
    // XSS：適合讀取內容，或是需要重複寫入用
    public enum ExcelType {
        SXSS, XSS
    }
}
