package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface CalculatedMember extends Com4jObject {
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
   * Getter method for the COM property "Formula"
   * </p>
   */

  @DISPID(261)
  @PropGet
  java.lang.String getFormula();


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
   * Getter method for the COM property "SolveOrder"
   * </p>
   */

  @DISPID(2187)
  @PropGet
  int getSolveOrder();


  /**
   * <p>
   * Getter method for the COM property "IsValid"
   * </p>
   */

  @DISPID(2188)
  @PropGet
  boolean getIsValid();


  /**
   * <p>
   * Getter method for the COM property "_Default"
   * </p>
   */

  @DISPID(0)
  @PropGet
  @DefaultMethod
  java.lang.String get_Default();


  /**
   */

  @DISPID(117)
  void delete();


  /**
   * <p>
   * Getter method for the COM property "Type"
   * </p>
   */

  @DISPID(108)
  @PropGet
  com.exceljava.com4j.excel.XlCalculatedMemberType getType();


  /**
   * <p>
   * Getter method for the COM property "Dynamic"
   * </p>
   */

  @DISPID(2926)
  @PropGet
  boolean getDynamic();


  /**
   * <p>
   * Getter method for the COM property "DisplayFolder"
   * </p>
   */

  @DISPID(2927)
  @PropGet
  java.lang.String getDisplayFolder();


  /**
   * <p>
   * Getter method for the COM property "HierarchizeDistinct"
   * </p>
   */

  @DISPID(2925)
  @PropGet
  boolean getHierarchizeDistinct();


  /**
   * <p>
   * Setter method for the COM property "HierarchizeDistinct"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2925)
  @PropPut
  void setHierarchizeDistinct(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "FlattenHierarchies"
   * </p>
   */

  @DISPID(2924)
  @PropGet
  boolean getFlattenHierarchies();


  /**
   * <p>
   * Setter method for the COM property "FlattenHierarchies"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2924)
  @PropPut
  void setFlattenHierarchies(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "MeasureGroup"
   * </p>
   */

  @DISPID(3092)
  @PropGet
  java.lang.String getMeasureGroup();


  /**
   * <p>
   * Getter method for the COM property "ParentHierarchy"
   * </p>
   */

  @DISPID(3093)
  @PropGet
  java.lang.String getParentHierarchy();


  /**
   * <p>
   * Getter method for the COM property "ParentMember"
   * </p>
   */

  @DISPID(3094)
  @PropGet
  java.lang.String getParentMember();


  /**
   * <p>
   * Getter method for the COM property "NumberFormat"
   * </p>
   */

  @DISPID(193)
  @PropGet
  com.exceljava.com4j.excel.XlCalcMemNumberFormatType getNumberFormat();


  // Properties:
}
