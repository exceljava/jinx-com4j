package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00024458-0001-0000-C000-000000000046}")
public interface IPivotCell extends Com4jObject {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Application"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel._Application
   */

  @VTID(7)
  com.exceljava.com4j.excel._Application getApplication();


  /**
   * <p>
   * Getter method for the COM property "Creator"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlCreator
   */

  @VTID(8)
  com.exceljava.com4j.excel.XlCreator getCreator();


  /**
   * <p>
   * Getter method for the COM property "Parent"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(9)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getParent();


  /**
   * <p>
   * Getter method for the COM property "PivotCellType"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlPivotCellType
   */

  @VTID(10)
  com.exceljava.com4j.excel.XlPivotCellType getPivotCellType();


  /**
   * <p>
   * Getter method for the COM property "PivotTable"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotTable
   */

  @VTID(11)
  com.exceljava.com4j.excel.PivotTable getPivotTable();


  /**
   * <p>
   * Getter method for the COM property "DataField"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotField
   */

  @VTID(12)
  com.exceljava.com4j.excel.PivotField getDataField();


  /**
   * <p>
   * Getter method for the COM property "PivotField"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotField
   */

  @VTID(13)
  com.exceljava.com4j.excel.PivotField getPivotField();


  /**
   * <p>
   * Getter method for the COM property "PivotItem"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotItem
   */

  @VTID(14)
  com.exceljava.com4j.excel.PivotItem getPivotItem();


  /**
   * <p>
   * Getter method for the COM property "RowItems"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotItemList
   */

  @VTID(15)
  com.exceljava.com4j.excel.PivotItemList getRowItems();


  /**
   * <p>
   * Getter method for the COM property "ColumnItems"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotItemList
   */

  @VTID(16)
  com.exceljava.com4j.excel.PivotItemList getColumnItems();


  /**
   * <p>
   * Getter method for the COM property "Range"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(17)
  com.exceljava.com4j.excel.Range getRange();


  /**
   * <p>
   * Getter method for the COM property "Dummy18"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(18)
  java.lang.String getDummy18();


  /**
   * <p>
   * Getter method for the COM property "CustomSubtotalFunction"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlConsolidationFunction
   */

  @VTID(19)
  com.exceljava.com4j.excel.XlConsolidationFunction getCustomSubtotalFunction();


  /**
   * <p>
   * Getter method for the COM property "PivotRowLine"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotLine
   */

  @VTID(20)
  com.exceljava.com4j.excel.PivotLine getPivotRowLine();


  /**
   * <p>
   * Getter method for the COM property "PivotColumnLine"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.PivotLine
   */

  @VTID(21)
  com.exceljava.com4j.excel.PivotLine getPivotColumnLine();


  /**
   */

  @VTID(22)
  void allocateChange();


  /**
   */

  @VTID(23)
  void discardChange();


  /**
   * <p>
   * Getter method for the COM property "DataSourceValue"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(24)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getDataSourceValue();


  /**
   * <p>
   * Getter method for the COM property "CellChanged"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlCellChangedState
   */

  @VTID(25)
  com.exceljava.com4j.excel.XlCellChangedState getCellChanged();


  /**
   * <p>
   * Getter method for the COM property "MDX"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(26)
  java.lang.String getMDX();


  /**
   * <p>
   * Getter method for the COM property "ServerActions"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Actions
   */

  @VTID(27)
  com.exceljava.com4j.excel.Actions getServerActions();


  // Properties:
}
