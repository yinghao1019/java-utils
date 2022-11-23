package com.tp.asset_ap.spreadsheet.parser;

import com.tp.asset_ap.exception.BadRequestException;
import com.tp.asset_ap.model.bo.HardwareBindExcelBO;
import com.tp.asset_ap.util.ExcelUtils;
import com.tp.asset_ap.util.LocaleUtils;
import com.tp.asset_ap.util.TPStringUtils;
import com.tp.asset_ap.util.TokenUtils;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class HardwareBindExcelParser extends ExcelParserTemplate<Map<Integer, HardwareBindExcelBO>> {
    private static final Integer ASSET_ID_CELL_INDEX = 0;
    private static final Integer AGENT_CELL_INDEX = 1;
    private static final Integer USER_CELL_INDEX = 2;
    private static final Integer LOCATION_CELL_INDEX = 3;
    private static final Integer PROJECT_CELL_INDEX = 4;
    private static final Integer DEPARTMENT_CELL_INDEX = 5;
    private static final Integer PICK_DATE_CELL_INDEX = 6;
    private static final Integer RETURN_DATE_CELL_INDEX = 7;

    @Override
    protected Map<Integer, HardwareBindExcelBO> parseEachRowData(
        int fisrtRow, int rowEnd,
        Sheet sheet) {
        Map<Integer, HardwareBindExcelBO> dataMap = new HashMap<>();
        int rowStart = fisrtRow + 3;

        for (int rowNum = rowStart; rowNum < rowEnd; rowNum++) {
            Row row = sheet.getRow(rowNum);
            if (isRowEmpty(row)) {
                continue;
            }
            //            驗證並設置錯誤訊息至對應的row
            try {
                validRowData(row);
            } catch (BadRequestException e) {
                row.createCell(ExcelUtils.HARDWARE_BIND_ERROR_CELL, CellType.STRING);
                getCell(row, ExcelUtils.HARDWARE_BIND_ERROR_CELL).setCellValue(e.getMessage());
                super.sheetData.setValid(false);
                continue;
            }

            HardwareBindExcelBO excelRowBO = convertRowData(row);
            dataMap.put(rowNum, excelRowBO);
        }
        return dataMap;
    }

    private HardwareBindExcelBO convertRowData(Row row) {
        HardwareBindExcelBO rowBO = new HardwareBindExcelBO();
        String assetId = convertCellValueToString(getCell(row, ASSET_ID_CELL_INDEX));
        rowBO.setAssetId(assetId);

        String agent = convertCellValueToString(getCell(row, AGENT_CELL_INDEX));
        rowBO.setAgent(TPStringUtils.isNullOrEmpty(agent) ? TokenUtils.getCode() : agent);

        String user = convertCellValueToString(getCell(row, USER_CELL_INDEX));
        rowBO.setUser(user);

        String locationCode = convertCellValueToString(getCell(row, LOCATION_CELL_INDEX));
        rowBO.setLocationCode(locationCode);

        String projectCode = convertCellValueToString(getCell(row, PROJECT_CELL_INDEX));
        rowBO.setProjectCode(projectCode);

        Integer departmentErpId = convertCellValueToInteger(getCell(row, DEPARTMENT_CELL_INDEX));
        rowBO.setDepartmentErpId(departmentErpId);

        Long pickupDate = convertCellValueToTimeStamp(getCell(row, PICK_DATE_CELL_INDEX));
        rowBO.setActualPickupDate(pickupDate);

        Long returnDate = convertCellValueToTimeStamp(getCell(row, RETURN_DATE_CELL_INDEX));
        rowBO.setEstimatedReturnDate(returnDate);
        return rowBO;
    }

    public void validRowData(Row row) throws BadRequestException {
        List<String> errorMsgList = new ArrayList<>();
        String errMsg = "";
        // 檢查資料不可為空
        if (cellIsEmpty(getCell(row, ASSET_ID_CELL_INDEX))) {
            errorMsgList.add(LocaleUtils.get("AdminHardwareEntity.assetType.name"));
        }
        if (cellIsEmpty(getCell(row, USER_CELL_INDEX))) {
            errorMsgList.add(LocaleUtils.get("AdminHardwareEntity.user.name"));
        }
        if (cellIsEmpty(getCell(row, LOCATION_CELL_INDEX))) {
            errorMsgList.add(LocaleUtils.get("AdminHardwareEntity.location.name"));
        }
        if (cellIsEmpty(getCell(row, PROJECT_CELL_INDEX))) {
            errorMsgList.add(LocaleUtils.get("AdminHardwareEntity.project.name"));
        }
        if (cellIsEmpty(getCell(row, DEPARTMENT_CELL_INDEX))) {
            errorMsgList.add(LocaleUtils.get("AdminHardwareEntity.department.name"));
        }
        if (cellIsEmpty(getCell(row, PICK_DATE_CELL_INDEX))) {
            errorMsgList.add(LocaleUtils.get("AdminHardwareEntity.actualPickupDate.name"));
        }

        if (!errorMsgList.isEmpty()) {
            errMsg = LocaleUtils.get("ExcelParser.required",
                String.join(",", errorMsgList)) + "\n";
        }

        if (!isValidInteger(row.getCell(DEPARTMENT_CELL_INDEX))) {
            errMsg += LocaleUtils.get("ExcelParser.invalid.integer") + "\n";
        }

        if (!isValidDateFormat(row.getCell(PICK_DATE_CELL_INDEX)) || !isValidDateFormat(
            row.getCell(RETURN_DATE_CELL_INDEX))) {
            errMsg += LocaleUtils.get("ExcelParser.invalid.dateTime") + "\n";
        }
        if (TPStringUtils.isNotEmpty(errMsg)) {
            throw new BadRequestException(errMsg);
        }
    }
}
