package com.tp.asset_ap.spreadsheet.parser;

import com.tp.asset_ap.exception.BadRequestException;
import com.tp.asset_ap.model.bo.HardwareExcelRowBO;

import com.tp.asset_ap.util.LocaleUtils;
import com.tp.asset_ap.util.TPStringUtils;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class HardwareExcelParser extends ExcelParserTemplate<Map<Integer, HardwareExcelRowBO>> {

    private static final Integer DEFINE_CELL_INDEX = 0;
    private static final Integer ASSET_CELL_INDEX = 1;
    private static final Integer BRAND_CELL_INDEX = 2;
    private static final Integer NAME_CELL_INDEX = 3;
    private static final Integer MODEL_CELL_INDEX = 4;
    private static final Integer SPEC_CELL_INDEX = 5;
    private static final Integer PRODUCT_CELL_INDEX = 6;
    private static final Integer PURCHASER_CELL_INDEX = 7;
    private static final Integer PUCRHASE_DATE_CELL_INDEX = 8;
    private static final Integer SUPPLIER_CELL_INDEX = 9;
    private static final Integer PUCRHASE_ORDER_CELL_INDEX = 10;
    private static final Integer PRICE_CELL_INDEX = 11;
    private static final Integer WARRANTY_START_CELL_INDEX = 12;
    private static final Integer WARRANTY_END_CELL_INDEX = 13;
    private static final Integer LOCATION_CELL_INDEX = 14;
    private static final Integer PROJECT_CELL_INDEX = 15;
    private static final Integer DEPARTMENT_CELL_INDEX = 16;
    private static final Integer DURABLE_CELL_INDEX = 17;
    private static final Integer ASSET_TYPE_CELL_INDEX = 18;
    private static final Integer ERROR_MSG_CELL_INDEX = 19;

    @Override
    protected Map<Integer, HardwareExcelRowBO> parseEachRowData(
        int fisrtRow, int rowEnd,
        Sheet sheet) {
        Map<Integer, HardwareExcelRowBO> dataMap = new HashMap<>();
        int rowStart = fisrtRow + 3;

        for (int rowNum = rowStart; rowNum < rowEnd; rowNum++) {
            Row row = sheet.getRow(rowNum);
            if (null == row || "".equals(convertCellValueToString(row.getCell(0)))
                || null == convertCellValueToString(row.getCell(0))) {
                break;
            }
            //            驗證並設置錯誤訊息至對應的row
            try {
                validRowData(row);
            } catch (BadRequestException e) {
                row.createCell(ERROR_MSG_CELL_INDEX, CellType.STRING);
                getCell(row, ERROR_MSG_CELL_INDEX).setCellValue(e.getMessage());
                super.sheetData.setValid(false);
                continue;
            }

            HardwareExcelRowBO excelRowBO = convertRowToPDTO(row);
            dataMap.put(rowNum, excelRowBO);
        }
        return dataMap;
    }

    private HardwareExcelRowBO convertRowToPDTO(Row row) {
        HardwareExcelRowBO excelRowBO = new HardwareExcelRowBO();
        String defineCode = convertCellValueToString(getCell(row, DEFINE_CELL_INDEX));
        excelRowBO.setDefineCode(defineCode);
        // 讀取資產項目
        String assetTypeCode = convertCellValueToString(getCell(row, ASSET_CELL_INDEX));
        excelRowBO.setAssetTypeCode(assetTypeCode);
        // 讀取廠牌
        String brandCode = convertCellValueToString(getCell(row, BRAND_CELL_INDEX));
        excelRowBO.setBrandCode(brandCode);
        // 讀取資產名稱
        String name = convertCellValueToString(getCell(row, NAME_CELL_INDEX));
        excelRowBO.setName(name);
        // 讀取資產型號
        String modelNumber = convertCellValueToString(getCell(row, MODEL_CELL_INDEX));
        excelRowBO.setModelNumber(modelNumber);
        // 讀取資產規格
        String specification = convertCellValueToString(getCell(row, SPEC_CELL_INDEX));
        excelRowBO.setSpecification(specification);
        // 讀取產品序號
        String productSerialNumber = convertCellValueToString(getCell(row, PRODUCT_CELL_INDEX));
        excelRowBO.setProductSerialNumber(productSerialNumber);
        // 讀取購買人
        String purchaser = convertCellValueToString(getCell(row, PURCHASER_CELL_INDEX));
        excelRowBO.setPurchaser(purchaser);
        // 讀取購買日期
        Long purchaseDate = convertCellValueToTimeStamp(getCell(row, PUCRHASE_DATE_CELL_INDEX));
        excelRowBO.setPurchaseDate(purchaseDate);
        // 讀取供應商
        String supplierCode = convertCellValueToString(getCell(row, SUPPLIER_CELL_INDEX));
        excelRowBO.setSupplierCode(supplierCode);
        // 讀取採購單號
        String purchaseOrderId = convertCellValueToString(getCell(row, PUCRHASE_ORDER_CELL_INDEX));
        excelRowBO.setPurchaseOrderId(purchaseOrderId);
        // 讀取取得成本
        Integer price = convertCellValueToInteger(getCell(row, PRICE_CELL_INDEX));
        excelRowBO.setPrice(price);
        // 讀取保固起日
        Long warrantyStartDate = convertCellValueToTimeStamp(
            getCell(row, WARRANTY_START_CELL_INDEX));
        excelRowBO.setWarrantyStartDate(warrantyStartDate);
        // 讀取保固迄日
        Long warrantyEndDate = convertCellValueToTimeStamp(getCell(row, WARRANTY_END_CELL_INDEX));
        excelRowBO.setWarrantyEndDate(warrantyEndDate);
        // 讀取地點
        String locationCode = convertCellValueToString(getCell(row, LOCATION_CELL_INDEX));
        excelRowBO.setLocationCode(locationCode);
        // 讀取所屬專案
        String projectCode = convertCellValueToString(getCell(row, PROJECT_CELL_INDEX));
        excelRowBO.setProjectCode(projectCode);
        // 讀取認列處別
        Integer departmentErpId = convertCellValueToInteger(getCell(row, DEPARTMENT_CELL_INDEX));
        excelRowBO.setDepartmentErpId(departmentErpId);
        // 耐用年限
        Integer durableYear = convertCellValueToInteger(getCell(row, DURABLE_CELL_INDEX));
        excelRowBO.setDurableYear(durableYear);
        //        資產別(公規資產,一般資產)
        Boolean standardSpec = convertCellValueToStandardSpecification(
            getCell(row, ASSET_TYPE_CELL_INDEX));
        excelRowBO.setStandardSpecification(standardSpec);
        return excelRowBO;
    }

    public void validRowData(Row row) throws BadRequestException {
        List<String> errorMsgList = new ArrayList<>();
        String errMsg = "";
        // 檢查資料不可為空
        //        資產定義
        if (cellIsEmpty(row.getCell(DEFINE_CELL_INDEX))) {
            errorMsgList.add(LocaleUtils.get("AdminHardwareEntity.define.name"));
        }
        //資產項目
        if (cellIsEmpty(row.getCell(ASSET_CELL_INDEX))) {
            errorMsgList.add(LocaleUtils.get("AdminHardwareEntity.assetType.name"));
        }
        //廠牌
        if (cellIsEmpty(row.getCell(BRAND_CELL_INDEX))) {
            errorMsgList.add(LocaleUtils.get("AdminHardwareEntity.brand.name"));
        }
        // 名稱
        if (cellIsEmpty(row.getCell(NAME_CELL_INDEX))) {
            errorMsgList.add(LocaleUtils.get("AdminHardwareEntity.assetName"));
        }
        // 型號
        if (cellIsEmpty(row.getCell(MODEL_CELL_INDEX))) {
            errorMsgList.add(LocaleUtils.get("AdminHardwareEntity.modelNumber.name"));
        }
        // 規格
        if (cellIsEmpty(row.getCell(SPEC_CELL_INDEX))) {
            errorMsgList.add(LocaleUtils.get("AdminHardwareEntity.spec.name"));
        }
        // 產品序號
        if (cellIsEmpty(row.getCell(PRODUCT_CELL_INDEX))) {
            errorMsgList.add(LocaleUtils.get("AdminHardwareEntity.serialNumber.name"));
        }
        // 資產別
        if (cellIsEmpty(row.getCell(ASSET_TYPE_CELL_INDEX))) {
            errorMsgList.add(LocaleUtils.get("AdminHardwareEntity.assetClass.name"));
        }
        if (!errorMsgList.isEmpty()) {
            errMsg = LocaleUtils.get("ExcelParser.required",
                String.join(",", errorMsgList)) + "\n";
        }

        //        檢查是否合法的整數(價格,年限,處別)
        if (!isValidInteger(row.getCell(PRICE_CELL_INDEX)) || !isValidInteger(
            row.getCell(DURABLE_CELL_INDEX)) || !isValidInteger(
            row.getCell(DEPARTMENT_CELL_INDEX))) {
            errMsg += LocaleUtils.get("ExcelParser.invalid.integer") + "\n";
        }
        //        檢查日期格式(購買日期,保固期限起迄日)
        if (!isValidDateFormat(row.getCell(PUCRHASE_DATE_CELL_INDEX)) || !isValidDateFormat(
            row.getCell(WARRANTY_START_CELL_INDEX)) || !isValidDateFormat(
            row.getCell(WARRANTY_END_CELL_INDEX))) {
            errMsg += LocaleUtils.get("ExcelParser.invalid.dateTime") + "\n";
        }
        if (TPStringUtils.isNotEmpty(errMsg)) {
            throw new BadRequestException(errMsg);
        }
    }
}
