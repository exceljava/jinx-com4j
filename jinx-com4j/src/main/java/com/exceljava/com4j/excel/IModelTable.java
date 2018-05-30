package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{000244D7-0001-0000-C000-000000000046}")
public interface IModelTable extends Com4jObject {
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
   * Getter method for the COM property "Name"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(10)
  java.lang.String getName();


  /**
   * <p>
   * Getter method for the COM property "SourceName"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(11)
  java.lang.String getSourceName();


  /**
   * <p>
   * Getter method for the COM property "ModelTableColumns"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.ModelTableColumns
   */

  @VTID(12)
  com.exceljava.com4j.excel.ModelTableColumns getModelTableColumns();


  /**
   * <p>
   * Getter method for the COM property "SourceWorkbookConnection"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.WorkbookConnection
   */

  @VTID(13)
  com.exceljava.com4j.excel.WorkbookConnection getSourceWorkbookConnection();


  /**
   */

  @VTID(14)
  void _Dummy7();


  /**
   */

  @VTID(15)
  void refresh();


  /**
   * <p>
   * Getter method for the COM property "RecordCount"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(16)
  int getRecordCount();


  /**
   * @param newName Mandatory java.lang.String parameter.
   */

  @VTID(17)
  void dummy1(
    java.lang.String newName);


  // Properties:
}
