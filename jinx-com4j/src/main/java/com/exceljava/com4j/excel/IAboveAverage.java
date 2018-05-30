package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{0002449E-0001-0000-C000-000000000046}")
public interface IAboveAverage extends Com4jObject {
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
   * Getter method for the COM property "Priority"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(10)
  int getPriority();


  /**
   * <p>
   * Setter method for the COM property "Priority"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(11)
  void setPriority(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "StopIfTrue"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(12)
  boolean getStopIfTrue();


  /**
   * <p>
   * Setter method for the COM property "StopIfTrue"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(13)
  void setStopIfTrue(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "AppliesTo"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(14)
  com.exceljava.com4j.excel.Range getAppliesTo();


  /**
   * <p>
   * Getter method for the COM property "AboveBelow"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlAboveBelow
   */

  @VTID(15)
  com.exceljava.com4j.excel.XlAboveBelow getAboveBelow();


  /**
   * <p>
   * Setter method for the COM property "AboveBelow"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlAboveBelow parameter.
   */

  @VTID(16)
  void setAboveBelow(
    com.exceljava.com4j.excel.XlAboveBelow rhs);


  /**
   * <p>
   * Getter method for the COM property "Interior"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Interior
   */

  @VTID(17)
  com.exceljava.com4j.excel.Interior getInterior();


  /**
   * <p>
   * Getter method for the COM property "Borders"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Borders
   */

  @VTID(18)
  com.exceljava.com4j.excel.Borders getBorders();


  /**
   * <p>
   * Getter method for the COM property "Font"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Font
   */

  @VTID(19)
  com.exceljava.com4j.excel.Font getFont();


  /**
   * <p>
   * Getter method for the COM property "Type"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(20)
  int getType();


  /**
   * <p>
   * Getter method for the COM property "NumberFormat"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(21)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getNumberFormat();


  /**
   * <p>
   * Setter method for the COM property "NumberFormat"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(22)
  void setNumberFormat(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   */

  @VTID(23)
  void setFirstPriority();


  /**
   */

  @VTID(24)
  void setLastPriority();


  /**
   */

  @VTID(25)
  void delete();


  /**
   * @param range Mandatory com.exceljava.com4j.excel.Range parameter.
   */

  @VTID(26)
  void modifyAppliesToRange(
    com.exceljava.com4j.excel.Range range);


  /**
   * <p>
   * Getter method for the COM property "PTCondition"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(27)
  boolean getPTCondition();


  /**
   * <p>
   * Getter method for the COM property "ScopeType"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlPivotConditionScope
   */

  @VTID(28)
  com.exceljava.com4j.excel.XlPivotConditionScope getScopeType();


  /**
   * <p>
   * Setter method for the COM property "ScopeType"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlPivotConditionScope parameter.
   */

  @VTID(29)
  void setScopeType(
    com.exceljava.com4j.excel.XlPivotConditionScope rhs);


  /**
   * <p>
   * Getter method for the COM property "CalcFor"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlCalcFor
   */

  @VTID(30)
  com.exceljava.com4j.excel.XlCalcFor getCalcFor();


  /**
   * <p>
   * Setter method for the COM property "CalcFor"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlCalcFor parameter.
   */

  @VTID(31)
  void setCalcFor(
    com.exceljava.com4j.excel.XlCalcFor rhs);


  /**
   * <p>
   * Getter method for the COM property "NumStdDev"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(32)
  int getNumStdDev();


  /**
   * <p>
   * Setter method for the COM property "NumStdDev"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(33)
  void setNumStdDev(
    int rhs);


  // Properties:
}
