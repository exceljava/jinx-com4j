package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{000244D9-0001-0000-C000-000000000046}")
public interface IModelRelationship extends Com4jObject {
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
   * Getter method for the COM property "ForeignKeyTable"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.ModelTable
   */

  @VTID(10)
  com.exceljava.com4j.excel.ModelTable getForeignKeyTable();


  /**
   * <p>
   * Getter method for the COM property "ForeignKeyColumn"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.ModelTableColumn
   */

  @VTID(11)
  com.exceljava.com4j.excel.ModelTableColumn getForeignKeyColumn();


  /**
   * <p>
   * Getter method for the COM property "PrimaryKeyTable"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.ModelTable
   */

  @VTID(12)
  com.exceljava.com4j.excel.ModelTable getPrimaryKeyTable();


  /**
   * <p>
   * Getter method for the COM property "PrimaryKeyColumn"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.ModelTableColumn
   */

  @VTID(13)
  com.exceljava.com4j.excel.ModelTableColumn getPrimaryKeyColumn();


  /**
   * <p>
   * Getter method for the COM property "Active"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(14)
  boolean getActive();


  /**
   * <p>
   * Setter method for the COM property "Active"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(15)
  void setActive(
    boolean rhs);


  /**
   */

  @VTID(16)
  void delete();


  // Properties:
}
