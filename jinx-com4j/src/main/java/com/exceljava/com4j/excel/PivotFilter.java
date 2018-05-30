package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface PivotFilter extends Com4jObject {
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
   * Getter method for the COM property "FilterType"
   * </p>
   */

  @DISPID(2686)
  @PropGet
  com.exceljava.com4j.excel.XlPivotFilterType getFilterType();


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
   * Getter method for the COM property "Description"
   * </p>
   */

  @DISPID(218)
  @PropGet
  java.lang.String getDescription();


  /**
   */

  @DISPID(117)
  void delete();


  /**
   * <p>
   * Getter method for the COM property "Active"
   * </p>
   */

  @DISPID(2312)
  @PropGet
  boolean getActive();


  /**
   * <p>
   * Getter method for the COM property "PivotField"
   * </p>
   */

  @DISPID(731)
  @PropGet
  com.exceljava.com4j.excel.PivotField getPivotField();


  /**
   * <p>
   * Getter method for the COM property "DataField"
   * </p>
   */

  @DISPID(2091)
  @PropGet
  com.exceljava.com4j.excel.PivotField getDataField();


  /**
   * <p>
   * Getter method for the COM property "DataCubeField"
   * </p>
   */

  @DISPID(2687)
  @PropGet
  com.exceljava.com4j.excel.CubeField getDataCubeField();


  /**
   * <p>
   * Getter method for the COM property "Value1"
   * </p>
   */

  @DISPID(2688)
  @PropGet
  java.lang.Object getValue1();


  /**
   * <p>
   * Getter method for the COM property "Value2"
   * </p>
   */

  @DISPID(1388)
  @PropGet
  java.lang.Object getValue2();


  /**
   * <p>
   * Getter method for the COM property "MemberPropertyField"
   * </p>
   */

  @DISPID(2689)
  @PropGet
  com.exceljava.com4j.excel.PivotField getMemberPropertyField();


  /**
   * <p>
   * Getter method for the COM property "IsMemberPropertyFilter"
   * </p>
   */

  @DISPID(2690)
  @PropGet
  boolean getIsMemberPropertyFilter();


  /**
   * <p>
   * Getter method for the COM property "WholeDayFilter"
   * </p>
   */

  @DISPID(3099)
  @PropGet
  boolean getWholeDayFilter();


  /**
   * <p>
   * Setter method for the COM property "WholeDayFilter"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(3099)
  @PropPut
  void setWholeDayFilter(
    boolean rhs);


  // Properties:
}
