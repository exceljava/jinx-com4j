package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface Trendline extends Com4jObject {
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
   * Getter method for the COM property "Backward"
   * </p>
   */

  @DISPID(185)
  @PropGet
  int getBackward();


  /**
   * <p>
   * Setter method for the COM property "Backward"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(185)
  @PropPut
  void setBackward(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "Border"
   * </p>
   */

  @DISPID(128)
  @PropGet
  com.exceljava.com4j.excel.Border getBorder();


  /**
   */

  @DISPID(112)
  java.lang.Object clearFormats();


  /**
   * <p>
   * Getter method for the COM property "DataLabel"
   * </p>
   */

  @DISPID(158)
  @PropGet
  com.exceljava.com4j.excel.DataLabel getDataLabel();


  /**
   */

  @DISPID(117)
  java.lang.Object delete();


  /**
   * <p>
   * Getter method for the COM property "DisplayEquation"
   * </p>
   */

  @DISPID(190)
  @PropGet
  boolean getDisplayEquation();


  /**
   * <p>
   * Setter method for the COM property "DisplayEquation"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(190)
  @PropPut
  void setDisplayEquation(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "DisplayRSquared"
   * </p>
   */

  @DISPID(189)
  @PropGet
  boolean getDisplayRSquared();


  /**
   * <p>
   * Setter method for the COM property "DisplayRSquared"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(189)
  @PropPut
  void setDisplayRSquared(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "Forward"
   * </p>
   */

  @DISPID(191)
  @PropGet
  int getForward();


  /**
   * <p>
   * Setter method for the COM property "Forward"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(191)
  @PropPut
  void setForward(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "Index"
   * </p>
   */

  @DISPID(486)
  @PropGet
  int getIndex();


  /**
   * <p>
   * Getter method for the COM property "Intercept"
   * </p>
   */

  @DISPID(186)
  @PropGet
  double getIntercept();


  /**
   * <p>
   * Setter method for the COM property "Intercept"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @DISPID(186)
  @PropPut
  void setIntercept(
    double rhs);


  /**
   * <p>
   * Getter method for the COM property "InterceptIsAuto"
   * </p>
   */

  @DISPID(187)
  @PropGet
  boolean getInterceptIsAuto();


  /**
   * <p>
   * Setter method for the COM property "InterceptIsAuto"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(187)
  @PropPut
  void setInterceptIsAuto(
    boolean rhs);


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
   * Setter method for the COM property "Name"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @DISPID(110)
  @PropPut
  void setName(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "NameIsAuto"
   * </p>
   */

  @DISPID(188)
  @PropGet
  boolean getNameIsAuto();


  /**
   * <p>
   * Setter method for the COM property "NameIsAuto"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(188)
  @PropPut
  void setNameIsAuto(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "Order"
   * </p>
   */

  @DISPID(192)
  @PropGet
  int getOrder();


  /**
   * <p>
   * Setter method for the COM property "Order"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(192)
  @PropPut
  void setOrder(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "Period"
   * </p>
   */

  @DISPID(184)
  @PropGet
  int getPeriod();


  /**
   * <p>
   * Setter method for the COM property "Period"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @DISPID(184)
  @PropPut
  void setPeriod(
    int rhs);


  /**
   */

  @DISPID(235)
  java.lang.Object select();


  /**
   * <p>
   * Getter method for the COM property "Type"
   * </p>
   */

  @DISPID(108)
  @PropGet
  com.exceljava.com4j.excel.XlTrendlineType getType();


  /**
   * <p>
   * Setter method for the COM property "Type"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlTrendlineType parameter.
   */

  @DISPID(108)
  @PropPut
  void setType(
    com.exceljava.com4j.excel.XlTrendlineType rhs);


  /**
   * <p>
   * Getter method for the COM property "Backward2"
   * </p>
   */

  @DISPID(2650)
  @PropGet
  double getBackward2();


  /**
   * <p>
   * Setter method for the COM property "Backward2"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @DISPID(2650)
  @PropPut
  void setBackward2(
    double rhs);


  /**
   * <p>
   * Getter method for the COM property "Forward2"
   * </p>
   */

  @DISPID(2651)
  @PropGet
  double getForward2();


  /**
   * <p>
   * Setter method for the COM property "Forward2"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @DISPID(2651)
  @PropPut
  void setForward2(
    double rhs);


  /**
   * <p>
   * Getter method for the COM property "Format"
   * </p>
   */

  @DISPID(116)
  @PropGet
  com.exceljava.com4j.excel.ChartFormat getFormat();


  /**
   * @param id Mandatory java.lang.String parameter.
   * @param value Mandatory java.lang.Object parameter.
   */

  @DISPID(3253)
  void setProperty(
    java.lang.String id,
    java.lang.Object value);


  /**
   * @param id Mandatory java.lang.String parameter.
   */

  @DISPID(3256)
  java.lang.Object getProperty(
    java.lang.String id);


  // Properties:
}
