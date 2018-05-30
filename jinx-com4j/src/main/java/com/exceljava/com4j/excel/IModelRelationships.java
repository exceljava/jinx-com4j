package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{000244DA-0001-0000-C000-000000000046}")
public interface IModelRelationships extends Com4jObject,Iterable<Com4jObject> {
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
   * Getter method for the COM property "Count"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(10)
  int getCount();


  /**
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.ModelRelationship
   */

  @VTID(11)
  com.exceljava.com4j.excel.ModelRelationship item(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "_Default"
   * </p>
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.ModelRelationship
   */

  @VTID(12)
  @DefaultMethod
  com.exceljava.com4j.excel.ModelRelationship get_Default(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "_NewEnum"
   * </p>
   */

  @VTID(13)
  java.util.Iterator<Com4jObject> iterator();

  /**
   * @param foreignKeyColumn Mandatory com.exceljava.com4j.excel.ModelTableColumn parameter.
   * @param primaryKeyColumn Mandatory com.exceljava.com4j.excel.ModelTableColumn parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.ModelRelationship
   */

  @VTID(14)
  com.exceljava.com4j.excel.ModelRelationship add(
    com.exceljava.com4j.excel.ModelTableColumn foreignKeyColumn,
    com.exceljava.com4j.excel.ModelTableColumn primaryKeyColumn);


  /**
   * @param pivotTable Mandatory com.exceljava.com4j.excel.PivotTable parameter.
   */

  @VTID(15)
  void detectRelationships(
    com.exceljava.com4j.excel.PivotTable pivotTable);


  // Properties:
}
