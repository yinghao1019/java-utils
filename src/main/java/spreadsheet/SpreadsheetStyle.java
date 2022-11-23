package com.tp.asset_ap.spreadsheet;

import lombok.Builder;
import lombok.Getter;
import lombok.Setter;
import org.apache.poi.xssf.usermodel.XSSFColor;

@Builder
@Setter
@Getter
public class SpreadsheetStyle {
    private SpreadsheetFontStyle fontStyle;
    @Builder.Default
    private TpHorizontalAlignment hAlign = TpHorizontalAlignment.GENERAL;
    @Builder.Default
    private TpVerticalAlignment vAlign = TpVerticalAlignment.CENTER;
    @Builder.Default
    private TpBorderStyle borderStyle = TpBorderStyle.THIN;
    @Builder.Default
    private SpreadsheetBorderPosition borderPosition = SpreadsheetBorderPosition.builder().build();
    private boolean wrapText;
    private XSSFColor backGroundColor;
}
