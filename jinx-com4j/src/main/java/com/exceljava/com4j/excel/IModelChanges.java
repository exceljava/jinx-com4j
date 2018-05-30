package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{000244E4-0001-0000-C000-000000000046}")
public interface IModelChanges extends Com4jObject {
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
   * Getter method for the COM property "TablesAdded"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.ModelTableNames
   */

  @VTID(10)
  com.exceljava.com4j.excel.ModelTableNames getTablesAdded();


  /**
   * <p>
   * Getter method for the COM property "TablesDeleted"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.ModelTableNames
   */

  @VTID(11)
  com.exceljava.com4j.excel.ModelTableNames getTablesDeleted();


  /**
   * <p>
   * Getter method for the COM property "TablesModified"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.ModelTableNames
   */

  @VTID(12)
  com.exceljava.com4j.excel.ModelTableNames getTablesModified();


  /**
   * <p>
   * Getter method for the COM property "TableNamesChanged"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.ModelTableNameChanges
   */

  @VTID(13)
  com.exceljava.com4j.excel.ModelTableNameChanges getTableNamesChanged();


  /**
   * <p>
   * Getter method for the COM property "RelationshipChange"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(14)
  boolean getRelationshipChange();


  /**
   * <p>
   * Getter method for the COM property "ColumnsAdded"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.ModelColumnNames
   */

  @VTID(15)
  com.exceljava.com4j.excel.ModelColumnNames getColumnsAdded();


  /**
   * <p>
   * Getter method for the COM property "ColumnsDeleted"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.ModelColumnNames
   */

  @VTID(16)
  com.exceljava.com4j.excel.ModelColumnNames getColumnsDeleted();


  /**
   * <p>
   * Getter method for the COM property "ColumnsChanged"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.ModelColumnChanges
   */

  @VTID(17)
  com.exceljava.com4j.excel.ModelColumnChanges getColumnsChanged();


  /**
   * <p>
   * Getter method for the COM property "MeasuresAdded"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.ModelMeasureNames
   */

  @VTID(18)
  com.exceljava.com4j.excel.ModelMeasureNames getMeasuresAdded();


  /**
   * <p>
   * Getter method for the COM property "UnknownChange"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(19)
  boolean getUnknownChange();


  /**
   * <p>
   * Getter method for the COM property "Source"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlModelChangeSource
   */

  @VTID(20)
  com.exceljava.com4j.excel.XlModelChangeSource getSource();


  // Properties:
}
