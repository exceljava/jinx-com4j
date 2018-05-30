package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface IconSetCondition extends Com4jObject {
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
   * Getter method for the COM property "AppliesTo"
   * </p>
   */

  @DISPID(2626)
  @PropGet
  com.exceljava.com4j.excel.Range getAppliesTo();


  /**
   * <p>
   * Getter method for the COM property "Type"
   * </p>
   */

  @DISPID(108)
  @PropGet
  int getType();


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
   * <p>
   * Getter method for the COM property "ReverseOrder"
   * </p>
   */

  @DISPID(2723)
  @PropGet
  boolean getReverseOrder();


  /**
   * <p>
   * Setter method for the COM property "ReverseOrder"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2723)
  @PropPut
  void setReverseOrder(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "PercentileValues"
   * </p>
   */

  @DISPID(2724)
  @PropGet
  boolean getPercentileValues();


  /**
   * <p>
   * Setter method for the COM property "PercentileValues"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2724)
  @PropPut
  void setPercentileValues(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "ShowIconOnly"
   * </p>
   */

  @DISPID(2725)
  @PropGet
  boolean getShowIconOnly();


  /**
   * <p>
   * Setter method for the COM property "ShowIconOnly"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(2725)
  @PropPut
  void setShowIconOnly(
    boolean rhs);


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
   * Setter method for the COM property "Formula"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(261)
  @PropPut
  void setFormula(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "IconSet"
   * </p>
   */

  @DISPID(2726)
  @PropGet
  java.lang.Object getIconSet();


  /**
   * <p>
   * Setter method for the COM property "IconSet"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @DISPID(2726)
  @PropPut
  void setIconSet(
    java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "IconCriteria"
   * </p>
   */

  @DISPID(2727)
  @PropGet
  com.exceljava.com4j.excel.IconCriteria getIconCriteria();


  // Properties:
}
