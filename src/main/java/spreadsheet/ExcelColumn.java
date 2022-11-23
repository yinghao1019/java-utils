package com.tp.asset_ap.spreadsheet;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Retention(RetentionPolicy.RUNTIME)
@Target(value = { ElementType.FIELD })
public @interface ExcelColumn {

    int colIndex();

    String colName();

    int columnWidth() default 20;

    String fontName() default "標楷體";

    int fontSize() default 12;

    TpHorizontalAlignment hAlign() default TpHorizontalAlignment.GENERAL;

    TpVerticalAlignment vAlign() default TpVerticalAlignment.CENTER;

    int rowSize() default 1;

    boolean isWrap() default true;

    String mergeGroup() default "";

    int mergeGroupSize() default 1;

    boolean mergeGroupAutoColWidth() default true;
}