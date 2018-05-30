package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface ModelChanges extends Com4jObject {
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
   * Getter method for the COM property "TablesAdded"
   * </p>
   */

  @DISPID(3145)
  @PropGet
  com.exceljava.com4j.excel.ModelTableNames getTablesAdded();


  /**
   * <p>
   * Getter method for the COM property "TablesDeleted"
   * </p>
   */

  @DISPID(3146)
  @PropGet
  com.exceljava.com4j.excel.ModelTableNames getTablesDeleted();


  /**
   * <p>
   * Getter method for the COM property "TablesModified"
   * </p>
   */

  @DISPID(3147)
  @PropGet
  com.exceljava.com4j.excel.ModelTableNames getTablesModified();


  /**
   * <p>
   * Getter method for the COM property "TableNamesChanged"
   * </p>
   */

  @DISPID(3148)
  @PropGet
  com.exceljava.com4j.excel.ModelTableNameChanges getTableNamesChanged();


  /**
   * <p>
   * Getter method for the COM property "RelationshipChange"
   * </p>
   */

  @DISPID(3149)
  @PropGet
  boolean getRelationshipChange();


  /**
   * <p>
   * Getter method for the COM property "ColumnsAdded"
   * </p>
   */

  @DISPID(3150)
  @PropGet
  com.exceljava.com4j.excel.ModelColumnNames getColumnsAdded();


  /**
   * <p>
   * Getter method for the COM property "ColumnsDeleted"
   * </p>
   */

  @DISPID(3151)
  @PropGet
  com.exceljava.com4j.excel.ModelColumnNames getColumnsDeleted();


  /**
   * <p>
   * Getter method for the COM property "ColumnsChanged"
   * </p>
   */

  @DISPID(3152)
  @PropGet
  com.exceljava.com4j.excel.ModelColumnChanges getColumnsChanged();


  /**
   * <p>
   * Getter method for the COM property "MeasuresAdded"
   * </p>
   */

  @DISPID(3153)
  @PropGet
  com.exceljava.com4j.excel.ModelMeasureNames getMeasuresAdded();


  /**
   * <p>
   * Getter method for the COM property "UnknownChange"
   * </p>
   */

  @DISPID(3154)
  @PropGet
  boolean getUnknownChange();


  /**
   * <p>
   * Getter method for the COM property "Source"
   * </p>
   */

  @DISPID(222)
  @PropGet
  com.exceljava.com4j.excel.XlModelChangeSource getSource();


  // Properties:
}
