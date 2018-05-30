package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00024497-0001-0000-C000-000000000046}")
public interface IIconSetCondition extends Com4jObject {
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
   * Getter method for the COM property "AppliesTo"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(13)
  com.exceljava.com4j.excel.Range getAppliesTo();


  /**
   * <p>
   * Getter method for the COM property "Type"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(14)
  int getType();


  /**
   * @param range Mandatory com.exceljava.com4j.excel.Range parameter.
   */

  @VTID(15)
  void modifyAppliesToRange(
    com.exceljava.com4j.excel.Range range);


  /**
   * <p>
   * Getter method for the COM property "PTCondition"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(16)
  boolean getPTCondition();


  /**
   * <p>
   * Getter method for the COM property "ScopeType"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlPivotConditionScope
   */

  @VTID(17)
  com.exceljava.com4j.excel.XlPivotConditionScope getScopeType();


  /**
   * <p>
   * Setter method for the COM property "ScopeType"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlPivotConditionScope parameter.
   */

  @VTID(18)
  void setScopeType(
    com.exceljava.com4j.excel.XlPivotConditionScope rhs);


  /**
   */

  @VTID(19)
  void setFirstPriority();


  /**
   */

  @VTID(20)
  void setLastPriority();


  /**
   */

  @VTID(21)
  void delete();


  /**
   * <p>
   * Getter method for the COM property "ReverseOrder"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(22)
  boolean getReverseOrder();


  /**
   * <p>
   * Setter method for the COM property "ReverseOrder"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(23)
  void setReverseOrder(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "PercentileValues"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(24)
  boolean getPercentileValues();


  /**
   * <p>
   * Setter method for the COM property "PercentileValues"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(25)
  void setPercentileValues(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "ShowIconOnly"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(26)
  boolean getShowIconOnly();


  /**
   * <p>
   * Setter method for the COM property "ShowIconOnly"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(27)
  void setShowIconOnly(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "Formula"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(28)
  java.lang.String getFormula();


  /**
   * <p>
   * Setter method for the COM property "Formula"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(29)
  void setFormula(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "IconSet"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(30)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getIconSet();


  /**
   * <p>
   * Setter method for the COM property "IconSet"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(31)
  void setIconSet(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "IconCriteria"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.IconCriteria
   */

  @VTID(32)
  com.exceljava.com4j.excel.IconCriteria getIconCriteria();


  // Properties:
}
