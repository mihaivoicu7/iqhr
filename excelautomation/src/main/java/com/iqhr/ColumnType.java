package com.iqhr;

import org.apache.poi.ss.usermodel.CellStyle;

public class ColumnType {
    private String columnName;
    private CellTypeCustom cellType = CellTypeCustom.STRING;
    private CellStyle cellStyle = null;
    private int trials = 0;

    public String getColumnName() {
        return columnName;
    }

    public void setColumnName(String columnName) {
        this.columnName = columnName;
    }

    public CellTypeCustom getCellType() {
        return cellType;
    }

    public void setCellType(CellTypeCustom cellType) {
        this.cellType = cellType;
    }

    public CellStyle getCellStyle() {
        return cellStyle;
    }

    public void setCellStyle(CellStyle cellStyle) {
        this.cellStyle = cellStyle;
    }

    public int getTrials() {
        return trials;
    }

    public void setTrials(int trials) {
        this.trials = trials;
    }
}
