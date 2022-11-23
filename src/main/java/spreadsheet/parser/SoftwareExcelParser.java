package com.tp.asset_ap.spreadsheet.parser;

import com.tp.asset_ap.exception.BadRequestException;
import com.tp.asset_ap.model.bo.HardwareExcelRowBO;
import com.tp.asset_ap.model.bo.SoftwareExcelRowBO;
import com.tp.asset_ap.util.ExcelUtils;
import com.tp.asset_ap.util.LocaleUtils;
import com.tp.asset_ap.util.TPStringUtils;

import java.text.ParseException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class SoftwareExcelParser extends ExcelParserTemplate<Map<Integer, SoftwareExcelRowBO>> {
    private static final Integer DEFINE_CELL_INDEX = 0;
    private static final Integer ASSET_CELL_INDEX = 1;
    private static final Integer NAME_CELL_INDEX = 2;
    private static final Integer VERSION_CELL_INDEX = 3;
    private static final Integer SPEC_CELL_INDEX = 4;
    private static final Integer PURCHASER_CELL_INDEX = 5;
    private static final Integer PUCRHASE_DATE_CELL_INDEX = 6;
    private static final Integer SUPPLIER_CELL_INDEX = 7;
    private static final Integer PUCRHASE_ORDER_CELL_INDEX = 8;
    private static final Integer PRICE_CELL_INDEX = 9;
    private static final Integer LOCATION_CELL_INDEX = 10;
    private static final Integer PROJECT_CELL_INDEX = 11;
    private static final Integer DEPARTMENT_CELL_INDEX = 12;
    private static final Integer LICENSE_TYPE_CELL_INDEX = 13;
    private static final Integer LICENSE_YEAR_CELL_INDEX = 14;
    private static final Integer LICENSE_START_CELL_INDEX = 15;
    private static final Integer LICENSE_END_CELL_INDEX = 16;
    private static final Integer LICENSE_COUNT_CELL_INDEX = 17;
    private static final Integer LINK_CELL_INDEX = 18;
    private static final Integer ACCOUNT_CELL_INDEX = 19;
    private static final Integer PASSWORD_CELL_INDEX = 20;
    private static final Integer KEY_CELL_INDEX = 21;
    private static final Integer OTHER_CELL_INDEX = 22;

    @Override
    protected Map<Integer, SoftwareExcelRowBO> parseEachRowData(
        int rowStart, int rowEnd,
        Sheet sheet) throws ParseException {
        Map<Integer, SoftwareExcelRowBO> dataMap = new HashMap<>();
        int firstRow = rowStart + 3;

        for (int rowNum = firstRow; rowNum < rowEnd; rowNum++) {
            Row row = sheet.getRow(rowNum);
            if (isRowEmpty(row)) {
                continue;
            }
            //            驗證並設置錯誤訊息至對應的row
            try {
                validRowData(row);
            } catch (BadRequestException e) {
                row.createCell(ExcelUtils.SOFTWARE_IMPORT_ERROR_CELL, CellType.STRING);
                getCell(row, ExcelUtils.SOFTWARE_IMPORT_ERROR_CELL).setCellValue(e.getMessage());
                super.sheetData.setValid(false);
                continue;
            }

            SoftwareExcelRowBO excelRowBO = convertRowToPDTO(row);
            dataMap.put(rowNum, excelRowBO);
        }
        return dataMap;
    }

    private SoftwareExcelRowBO convertRowToPDTO(Row row) {
        SoftwareExcelRowBO excelRowBO = new SoftwareExcelRowBO();

        String defineCode = convertCellValueToString(getCell(row, DEFINE_CELL_INDEX));
        excelRowBO.setAssetTypeCode(defineCode);

        String assetTypeCode = convertCellValueToString(getCell(row, ASSET_CELL_INDEX));
        excelRowBO.setAssetTypeCode(assetTypeCode);

        String name = convertCellValueToString(getCell(row, NAME_CELL_INDEX));
        excelRowBO.setName(name);

        String version = convertCellValueToString(getCell(row, VERSION_CELL_INDEX));
        excelRowBO.setVersion(version);

        String specification = convertCellValueToString(getCell(row, SPEC_CELL_INDEX));
        excelRowBO.setSpecification(specification);

        String purchaser = convertCellValueToString(getCell(row, PURCHASER_CELL_INDEX));
        excelRowBO.setPurchaser(purchaser);

        Long purchaseDate = convertCellValueToTimeStamp(getCell(row, PUCRHASE_DATE_CELL_INDEX));
        excelRowBO.setPurchaseDate(purchaseDate);

        String supplierCode = convertCellValueToString(getCell(row, SUPPLIER_CELL_INDEX));
        excelRowBO.setSupplierCode(supplierCode);

        String purchaseOrderId = convertCellValueToString(getCell(row, PUCRHASE_ORDER_CELL_INDEX));
        excelRowBO.setPurchaseOrderId(purchaseOrderId);

        Integer price = convertCellValueToInteger(getCell(row, PRICE_CELL_INDEX));
        excelRowBO.setPrice(price);

        String locationCode = convertCellValueToString(getCell(row, LOCATION_CELL_INDEX));
        excelRowBO.setLocationCode(locationCode);

        String projectCode = convertCellValueToString(getCell(row, PROJECT_CELL_INDEX));
        excelRowBO.setProjectCode(projectCode);

        Integer departmentErpId = convertCellValueToInteger(getCell(row, DEPARTMENT_CELL_INDEX));
        excelRowBO.setDepartmentErpId(departmentErpId);

        String licenseTypeCode = convertCellValueToString(getCell(row, LICENSE_TYPE_CELL_INDEX));
        excelRowBO.setLicenseTypeCode(licenseTypeCode);

        String licenseYearCode = convertCellValueToString(getCell(row, LICENSE_YEAR_CELL_INDEX));
        excelRowBO.setLicenseYearCode(licenseYearCode);

        Long licenseStartDate = convertCellValueToTimeStamp(getCell(row, LICENSE_START_CELL_INDEX));
        excelRowBO.setLicenseStartDate(licenseStartDate);

        Long licenseEndDate = convertCellValueToTimeStamp(getCell(row, LICENSE_END_CELL_INDEX));
        excelRowBO.setLicenseEndDate(licenseEndDate);

        Integer licenseCount = convertCellValueToInteger(getCell(row, LICENSE_COUNT_CELL_INDEX));
        excelRowBO.setLicenseCount(licenseCount);

        String link = convertCellValueToString(getCell(row, LINK_CELL_INDEX));
        excelRowBO.setLink(link);

        String account = convertCellValueToString(getCell(row, ACCOUNT_CELL_INDEX));
        excelRowBO.setAccount(account);

        String password = convertCellValueToString(getCell(row, PASSWORD_CELL_INDEX));
        excelRowBO.setPassword(password);

        String key = convertCellValueToString(getCell(row, KEY_CELL_INDEX));
        excelRowBO.setKey(key);

        String other = convertCellValueToString(getCell(row, OTHER_CELL_INDEX));
        excelRowBO.setOther(other);
        return excelRowBO;
    }

    public void validRowData(Row row) throws BadRequestException {
        List<String> errorMsgList = new ArrayList<>();
        String errMsg = "";
        // 檢查資料不可為空
        if (cellIsEmpty(row.getCell(DEFINE_CELL_INDEX))) {
            errorMsgList.add(LocaleUtils.get("SoftwareEntity.define.name"));
        }
        if (cellIsEmpty(row.getCell(ASSET_CELL_INDEX))) {
            errorMsgList.add(LocaleUtils.get("SoftwareEntity.assetType.name"));
        }
        if (cellIsEmpty(row.getCell(NAME_CELL_INDEX))) {
            errorMsgList.add(LocaleUtils.get("SoftwareEntity.name"));
        }
        if (cellIsEmpty(row.getCell(VERSION_CELL_INDEX))) {
            errorMsgList.add(LocaleUtils.get("SoftwareEntity.version.name"));
        }
        if (cellIsEmpty(row.getCell(SPEC_CELL_INDEX))) {
            errorMsgList.add(LocaleUtils.get("SoftwareEntity.spec.name"));
        }
        if (cellIsEmpty(row.getCell(LICENSE_TYPE_CELL_INDEX))) {
            errorMsgList.add(LocaleUtils.get("SoftwareEntity.licenseType.name"));
        }
        if (cellIsEmpty(row.getCell(LICENSE_YEAR_CELL_INDEX))) {
            errorMsgList.add(LocaleUtils.get("SoftwareEntity.licenseYear.name"));
        }
        if (cellIsEmpty(row.getCell(LICENSE_START_CELL_INDEX))) {
            errorMsgList.add(LocaleUtils.get("SoftwareEntity.licenseStart.name"));
        }
        if (cellIsEmpty(row.getCell(LICENSE_END_CELL_INDEX))) {
            errorMsgList.add(LocaleUtils.get("SoftwareEntity.licenseYear.name"));
        }
        if (cellIsEmpty(row.getCell(LICENSE_COUNT_CELL_INDEX))) {
            errorMsgList.add(LocaleUtils.get("SoftwareEntity.licenseCount.name"));
        }
        if (!errorMsgList.isEmpty()) {
            errMsg = LocaleUtils.get("ExcelParser.required",
                String.join(",", errorMsgList)) + "\n";
        }

        //        檢查是否合法的整數(價格,年限,處別)
        if (!isValidInteger(row.getCell(PRICE_CELL_INDEX)) || !isValidInteger(
            row.getCell(LICENSE_COUNT_CELL_INDEX)) || !isValidInteger(
            row.getCell(DEPARTMENT_CELL_INDEX))) {
            errMsg += LocaleUtils.get("ExcelParser.invalid.integer") + "\n";
        }
        //        檢查日期格式(購買日期,保固期限起迄日)
        if (!isValidDateFormat(row.getCell(PUCRHASE_DATE_CELL_INDEX)) || !isValidDateFormat(
            row.getCell(LICENSE_START_CELL_INDEX)) || !isValidDateFormat(
            row.getCell(LICENSE_END_CELL_INDEX))) {
            errMsg += LocaleUtils.get("ExcelParser.invalid.dateTime") + "\n";
        }
        if (TPStringUtils.isNotEmpty(errMsg)) {
            throw new BadRequestException(errMsg);
        }
    }
}
