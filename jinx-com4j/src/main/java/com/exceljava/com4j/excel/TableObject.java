package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface TableObject extends Com4jObject {
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
   * Getter method for the COM property "RowNumbers"
   * </p>
   */

  @DISPID(1585)
  @PropGet
  boolean getRowNumbers();


  /**
   * <p>
   * Setter method for the COM property "RowNumbers"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1585)
  @PropPut
  void setRowNumbers(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "FetchedRowOverflow"
   * </p>
   */

  @DISPID(1588)
  @PropGet
  boolean getFetchedRowOverflow();


  /**
   * <p>
   * Getter method for the COM property "RefreshStyle"
   * </p>
   */

  @DISPID(1590)
  @PropGet
  com.exceljava.com4j.excel.XlCellInsertionMode getRefreshStyle();


  /**
   * <p>
   * Setter method for the COM property "RefreshStyle"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlCellInsertionMode parameter.
   */

  @DISPID(1590)
  @PropPut
  void setRefreshStyle(
    com.exceljava.com4j.excel.XlCellInsertionMode rhs);


  /**
   * <p>
   * Getter method for the COM property "EnableRefresh"
   * </p>
   */

  @DISPID(1477)
  @PropGet
  boolean getEnableRefresh();


  /**
   * <p>
   * Setter method for the COM property "EnableRefresh"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1477)
  @PropPut
  void setEnableRefresh(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "Destination"
   * </p>
   */

  @DISPID(681)
  @PropGet
  com.exceljava.com4j.excel.Range getDestination();


  /**
   * <p>
   * Getter method for the COM property "ResultRange"
   * </p>
   */

  @DISPID(1592)
  @PropGet
  com.exceljava.com4j.excel.Range getResultRange();


  /**
   */

  @DISPID(117)
  void delete();


  /**
   */

  @DISPID(1417)
  boolean refresh();


  /**
   * <p>
   * Getter method for the COM property "EnableEditing"
   * </p>
   */

  @DISPID(1595)
  @PropGet
  boolean getEnableEditing();


  /**
   * <p>
   * Setter method for the COM property "EnableEditing"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1595)
  @PropPut
  void setEnableEditing(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "PreserveColumnInfo"
   * </p>
   */

  @DISPID(1867)
  @PropGet
  boolean getPreserveColumnInfo();


  /**
   * <p>
   * Setter method for the COM property "PreserveColumnInfo"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1867)
  @PropPut
  void setPreserveColumnInfo(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "PreserveFormatting"
   * </p>
   */

  @DISPID(1500)
  @PropGet
  boolean getPreserveFormatting();


  /**
   * <p>
   * Setter method for the COM property "PreserveFormatting"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1500)
  @PropPut
  void setPreserveFormatting(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "AdjustColumnWidth"
   * </p>
   */

  @DISPID(1868)
  @PropGet
  boolean getAdjustColumnWidth();


  /**
   * <p>
   * Setter method for the COM property "AdjustColumnWidth"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(1868)
  @PropPut
  void setAdjustColumnWidth(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "ListObject"
   * </p>
   */

  @DISPID(2257)
  @PropGet
  com.exceljava.com4j.excel.ListObject getListObject();


  /**
   * <p>
   * Getter method for the COM property "WorkbookConnection"
   * </p>
   */

  @DISPID(2544)
  @PropGet
  com.exceljava.com4j.excel.WorkbookConnection getWorkbookConnection();


  // Properties:
}
