package com.tp.asset_ap.spreadsheet;

import lombok.Builder;
import lombok.Getter;
import lombok.Setter;

@Builder
@Setter
@Getter
public class SpreadsheetBorderPosition {
    @Builder.Default
    private boolean left = true;
    @Builder.Default
    private boolean right = true;
    @Builder.Default
    private boolean top = true;
    @Builder.Default
    private boolean bottom = true;
}
