package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface ModelRelationship extends Com4jObject {
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
   * Getter method for the COM property "ForeignKeyTable"
   * </p>
   */

  @DISPID(3122)
  @PropGet
  com.exceljava.com4j.excel.ModelTable getForeignKeyTable();


  /**
   * <p>
   * Getter method for the COM property "ForeignKeyColumn"
   * </p>
   */

  @DISPID(3123)
  @PropGet
  com.exceljava.com4j.excel.ModelTableColumn getForeignKeyColumn();


  /**
   * <p>
   * Getter method for the COM property "PrimaryKeyTable"
   * </p>
   */

  @DISPID(3124)
  @PropGet
  com.exceljava.com4j.excel.ModelTable getPrimaryKeyTable();


  /**
   * <p>
   * Getter method for the COM property "PrimaryKeyColumn"
   * </p>
   */

  @DISPID(3125)
  @PropGet
  com.exceljava.com4j.excel.ModelTableColumn getPrimaryKeyColumn();


  /**
   * <p>
   * Getter method for the COM property "Active"
   * </p>
   */

  @DISPID(2312)
  @PropGet
  boolean getActive();


  /**
   * <p>
   * Setter method for the COM property "Active"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2312)
  @PropPut
  void setActive(
    boolean rhs);


  /**
   */

  @DISPID(117)
  void delete();


  // Properties:
}
