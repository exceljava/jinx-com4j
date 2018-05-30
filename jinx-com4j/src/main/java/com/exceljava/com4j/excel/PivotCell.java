package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface PivotCell extends Com4jObject {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Application"
   * </p>
   */

  @DISPID(148)
  @PropGet
  com.exceljava.com4j.excel._Application getApplication();


  /**
   * <p>
   * Getter method for the COM property "Creator"
   * </p>
   */

  @DISPID(149)
  @PropGet
  com.exceljava.com4j.excel.XlCreator getCreator();


  /**
   * <p>
   * Getter method for the COM property "Parent"
   * </p>
   */

  @DISPID(150)
  @PropGet
  com4j.Com4jObject getParent();


  /**
   * <p>
   * Getter method for the COM property "PivotCellType"
   * </p>
   */

  @DISPID(2189)
  @PropGet
  com.exceljava.com4j.excel.XlPivotCellType getPivotCellType();


  /**
   * <p>
   * Getter method for the COM property "PivotTable"
   * </p>
   */

  @DISPID(716)
  @PropGet
  com.exceljava.com4j.excel.PivotTable getPivotTable();


  /**
   * <p>
   * Getter method for the COM property "DataField"
   * </p>
   */

  @DISPID(2091)
  @PropGet
  com.exceljava.com4j.excel.PivotField getDataField();


  /**
   * <p>
   * Getter method for the COM property "PivotField"
   * </p>
   */

  @DISPID(731)
  @PropGet
  com.exceljava.com4j.excel.PivotField getPivotField();


  /**
   * <p>
   * Getter method for the COM property "PivotItem"
   * </p>
   */

  @DISPID(740)
  @PropGet
  com.exceljava.com4j.excel.PivotItem getPivotItem();


  /**
   * <p>
   * Getter method for the COM property "RowItems"
   * </p>
   */

  @DISPID(2190)
  @PropGet
  com.exceljava.com4j.excel.PivotItemList getRowItems();


  /**
   * <p>
   * Getter method for the COM property "ColumnItems"
   * </p>
   */

  @DISPID(2191)
  @PropGet
  com.exceljava.com4j.excel.PivotItemList getColumnItems();


  /**
   * <p>
   * Getter method for the COM property "Range"
   * </p>
   */

  @DISPID(197)
  @PropGet
  com.exceljava.com4j.excel.Range getRange();


  /**
   * <p>
   * Getter method for the COM property "Dummy18"
   * </p>
   */

  @DISPID(2295)
  @PropGet
  java.lang.String getDummy18();


  /**
   * <p>
   * Getter method for the COM property "CustomSubtotalFunction"
   * </p>
   */

  @DISPID(2193)
  @PropGet
  com.exceljava.com4j.excel.XlConsolidationFunction getCustomSubtotalFunction();


  /**
   * <p>
   * Getter method for the COM property "PivotRowLine"
   * </p>
   */

  @DISPID(2673)
  @PropGet
  com.exceljava.com4j.excel.PivotLine getPivotRowLine();


  /**
   * <p>
   * Getter method for the COM property "PivotColumnLine"
   * </p>
   */

  @DISPID(2674)
  @PropGet
  com.exceljava.com4j.excel.PivotLine getPivotColumnLine();


  /**
   */

  @DISPID(2928)
  void allocateChange();


  /**
   */

  @DISPID(2929)
  void discardChange();


  /**
   * <p>
   * Getter method for the COM property "DataSourceValue"
   * </p>
   */

  @DISPID(2930)
  @PropGet
  java.lang.Object getDataSourceValue();


  /**
   * <p>
   * Getter method for the COM property "CellChanged"
   * </p>
   */

  @DISPID(2931)
  @PropGet
  com.exceljava.com4j.excel.XlCellChangedState getCellChanged();


  /**
   * <p>
   * Getter method for the COM property "MDX"
   * </p>
   */

  @DISPID(2123)
  @PropGet
  java.lang.String getMDX();


  /**
   * <p>
   * Getter method for the COM property "ServerActions"
   * </p>
   */

  @DISPID(2491)
  @PropGet
  com.exceljava.com4j.excel.Actions getServerActions();


  // Properties:
}
