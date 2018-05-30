package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface ModelTable extends Com4jObject {
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
   * Getter method for the COM property "Name"
   * </p>
   */

  @DISPID(110)
  @PropGet
  java.lang.String getName();


  /**
   * <p>
   * Getter method for the COM property "SourceName"
   * </p>
   */

  @DISPID(721)
  @PropGet
  java.lang.String getSourceName();


  /**
   * <p>
   * Getter method for the COM property "ModelTableColumns"
   * </p>
   */

  @DISPID(3119)
  @PropGet
  com.exceljava.com4j.excel.ModelTableColumns getModelTableColumns();


  /**
   * <p>
   * Getter method for the COM property "SourceWorkbookConnection"
   * </p>
   */

  @DISPID(3120)
  @PropGet
  com.exceljava.com4j.excel.WorkbookConnection getSourceWorkbookConnection();


  /**
   */

  @DISPID(65543)
  void _Dummy7();


  /**
   */

  @DISPID(1417)
  void refresh();


  /**
   * <p>
   * Getter method for the COM property "RecordCount"
   * </p>
   */

  @DISPID(1478)
  @PropGet
  int getRecordCount();


  /**
   * @param newName Mandatory java.lang.String parameter.
   */

  @DISPID(1782)
  void dummy1(
    java.lang.String newName);


  // Properties:
}
