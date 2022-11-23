package com.tp.asset_ap.spreadsheet;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Getter;
import lombok.Setter;
import org.apache.poi.ss.usermodel.IndexedColors;

@Builder
@Setter
@Getter
@AllArgsConstructor
public class SpreadsheetFontStyle {
    private String fontName;
    private int fontSize;
    private boolean bold;
    private boolean italic;
    @Builder.Default
    private IndexedColors color = IndexedColors.BLACK;
}
