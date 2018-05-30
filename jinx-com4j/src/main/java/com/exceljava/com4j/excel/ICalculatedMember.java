package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00024455-0001-0000-C000-000000000046}")
public interface ICalculatedMember extends Com4jObject {
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
   * Getter method for the COM property "Formula"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(11)
  java.lang.String getFormula();


  /**
   * <p>
   * Getter method for the COM property "SourceName"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(12)
  java.lang.String getSourceName();


  /**
   * <p>
   * Getter method for the COM property "SolveOrder"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(13)
  int getSolveOrder();


  /**
   * <p>
   * Getter method for the COM property "IsValid"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(14)
  boolean getIsValid();


  /**
   * <p>
   * Getter method for the COM property "_Default"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(15)
  @DefaultMethod
  java.lang.String get_Default();


  /**
   */

  @VTID(16)
  void delete();


  /**
   * <p>
   * Getter method for the COM property "Type"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlCalculatedMemberType
   */

  @VTID(17)
  com.exceljava.com4j.excel.XlCalculatedMemberType getType();


  /**
   * <p>
   * Getter method for the COM property "Dynamic"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(18)
  boolean getDynamic();


  /**
   * <p>
   * Getter method for the COM property "DisplayFolder"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(19)
  java.lang.String getDisplayFolder();


  /**
   * <p>
   * Getter method for the COM property "HierarchizeDistinct"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(20)
  boolean getHierarchizeDistinct();


  /**
   * <p>
   * Setter method for the COM property "HierarchizeDistinct"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(21)
  void setHierarchizeDistinct(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "FlattenHierarchies"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(22)
  boolean getFlattenHierarchies();


  /**
   * <p>
   * Setter method for the COM property "FlattenHierarchies"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(23)
  void setFlattenHierarchies(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "MeasureGroup"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(24)
  java.lang.String getMeasureGroup();


  /**
   * <p>
   * Getter method for the COM property "ParentHierarchy"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(25)
  java.lang.String getParentHierarchy();


  /**
   * <p>
   * Getter method for the COM property "ParentMember"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(26)
  java.lang.String getParentMember();


  /**
   * <p>
   * Getter method for the COM property "NumberFormat"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlCalcMemNumberFormatType
   */

  @VTID(27)
  com.exceljava.com4j.excel.XlCalcMemNumberFormatType getNumberFormat();


  // Properties:
}
