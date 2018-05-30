package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{000244CE-0001-0000-C000-000000000046}")
public interface ITableObject extends Com4jObject {
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
   * Getter method for the COM property "RowNumbers"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(10)
  boolean getRowNumbers();


  /**
   * <p>
   * Setter method for the COM property "RowNumbers"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(11)
  void setRowNumbers(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "FetchedRowOverflow"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(12)
  boolean getFetchedRowOverflow();


  /**
   * <p>
   * Getter method for the COM property "RefreshStyle"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlCellInsertionMode
   */

  @VTID(13)
  com.exceljava.com4j.excel.XlCellInsertionMode getRefreshStyle();


  /**
   * <p>
   * Setter method for the COM property "RefreshStyle"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlCellInsertionMode parameter.
   */

  @VTID(14)
  void setRefreshStyle(
    com.exceljava.com4j.excel.XlCellInsertionMode rhs);


  /**
   * <p>
   * Getter method for the COM property "EnableRefresh"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(15)
  boolean getEnableRefresh();


  /**
   * <p>
   * Setter method for the COM property "EnableRefresh"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(16)
  void setEnableRefresh(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "Destination"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(17)
  com.exceljava.com4j.excel.Range getDestination();


  /**
   * <p>
   * Getter method for the COM property "ResultRange"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(18)
  com.exceljava.com4j.excel.Range getResultRange();


  /**
   */

  @VTID(19)
  void delete();


  /**
   * @return  Returns a value of type boolean
   */

  @VTID(20)
  boolean refresh();


  /**
   * <p>
   * Getter method for the COM property "EnableEditing"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(21)
  boolean getEnableEditing();


  /**
   * <p>
   * Setter method for the COM property "EnableEditing"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(22)
  void setEnableEditing(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "PreserveColumnInfo"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(23)
  boolean getPreserveColumnInfo();


  /**
   * <p>
   * Setter method for the COM property "PreserveColumnInfo"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(24)
  void setPreserveColumnInfo(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "PreserveFormatting"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(25)
  boolean getPreserveFormatting();


  /**
   * <p>
   * Setter method for the COM property "PreserveFormatting"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(26)
  void setPreserveFormatting(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "AdjustColumnWidth"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(27)
  boolean getAdjustColumnWidth();


  /**
   * <p>
   * Setter method for the COM property "AdjustColumnWidth"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(28)
  void setAdjustColumnWidth(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "ListObject"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.ListObject
   */

  @VTID(29)
  com.exceljava.com4j.excel.ListObject getListObject();


  /**
   * <p>
   * Getter method for the COM property "WorkbookConnection"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.WorkbookConnection
   */

  @VTID(30)
  com.exceljava.com4j.excel.WorkbookConnection getWorkbookConnection();


  // Properties:
}
