package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface Top10 extends Com4jObject {
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
   * Getter method for the COM property "Priority"
   * </p>
   */

  @DISPID(985)
  @PropGet
  int getPriority();


  /**
   * <p>
   * Setter method for the COM property "Priority"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(985)
  @PropPut
  void setPriority(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "StopIfTrue"
   * </p>
   */

  @DISPID(2625)
  @PropGet
  boolean getStopIfTrue();


  /**
   * <p>
   * Setter method for the COM property "StopIfTrue"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2625)
  @PropPut
  void setStopIfTrue(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "AppliesTo"
   * </p>
   */

  @DISPID(2626)
  @PropGet
  com.exceljava.com4j.excel.Range getAppliesTo();


  /**
   * <p>
   * Getter method for the COM property "TopBottom"
   * </p>
   */

  @DISPID(2728)
  @PropGet
  com.exceljava.com4j.excel.XlTopBottom getTopBottom();


  /**
   * <p>
   * Setter method for the COM property "TopBottom"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlTopBottom parameter.
   */

  @DISPID(2728)
  @PropPut
  void setTopBottom(
    com.exceljava.com4j.excel.XlTopBottom rhs);


  /**
   * <p>
   * Getter method for the COM property "Rank"
   * </p>
   */

  @DISPID(1290)
  @PropGet
  int getRank();


  /**
   * <p>
   * Setter method for the COM property "Rank"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(1290)
  @PropPut
  void setRank(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "Percent"
   * </p>
   */

  @DISPID(2729)
  @PropGet
  boolean getPercent();


  /**
   * <p>
   * Setter method for the COM property "Percent"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2729)
  @PropPut
  void setPercent(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "Interior"
   * </p>
   */

  @DISPID(129)
  @PropGet
  com.exceljava.com4j.excel.Interior getInterior();


  /**
   * <p>
   * Getter method for the COM property "Borders"
   * </p>
   */

  @DISPID(435)
  @PropGet
  com.exceljava.com4j.excel.Borders getBorders();


  /**
   * <p>
   * Getter method for the COM property "Font"
   * </p>
   */

  @DISPID(146)
  @PropGet
  com.exceljava.com4j.excel.Font getFont();


  /**
   * <p>
   * Getter method for the COM property "Type"
   * </p>
   */

  @DISPID(108)
  @PropGet
  int getType();


  /**
   * <p>
   * Getter method for the COM property "NumberFormat"
   * </p>
   */

  @DISPID(193)
  @PropGet
  java.lang.Object getNumberFormat();


  /**
   * <p>
   * Setter method for the COM property "NumberFormat"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(193)
  @PropPut
  void setNumberFormat(
    java.lang.Object rhs);


  /**
   */

  @DISPID(2629)
  void setFirstPriority();


  /**
   */

  @DISPID(2630)
  void setLastPriority();


  /**
   */

  @DISPID(117)
  void delete();


  /**
   * @param range Mandatory com.exceljava.com4j.excel.Range parameter.
   */

  @DISPID(2627)
  void modifyAppliesToRange(
    com.exceljava.com4j.excel.Range range);


  /**
   * <p>
   * Getter method for the COM property "PTCondition"
   * </p>
   */

  @DISPID(2631)
  @PropGet
  boolean getPTCondition();


  /**
   * <p>
   * Getter method for the COM property "ScopeType"
   * </p>
   */

  @DISPID(2615)
  @PropGet
  com.exceljava.com4j.excel.XlPivotConditionScope getScopeType();


  /**
   * <p>
   * Setter method for the COM property "ScopeType"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlPivotConditionScope parameter.
   */

  @DISPID(2615)
  @PropPut
  void setScopeType(
    com.exceljava.com4j.excel.XlPivotConditionScope rhs);


  /**
   * <p>
   * Getter method for the COM property "CalcFor"
   * </p>
   */

  @DISPID(2730)
  @PropGet
  com.exceljava.com4j.excel.XlCalcFor getCalcFor();


  /**
   * <p>
   * Setter method for the COM property "CalcFor"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlCalcFor parameter.
   */

  @DISPID(2730)
  @PropPut
  void setCalcFor(
    com.exceljava.com4j.excel.XlCalcFor rhs);


  // Properties:
}
