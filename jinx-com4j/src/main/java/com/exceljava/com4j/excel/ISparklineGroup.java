package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{000244B7-0001-0000-C000-000000000046}")
public interface ISparklineGroup extends Com4jObject,Iterable<Com4jObject> {
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
   * Getter method for the COM property "Count"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(10)
  int getCount();


  /**
   * <p>
   * Getter method for the COM property "Item"
   * </p>
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.Sparkline
   */

  @VTID(11)
  com.exceljava.com4j.excel.Sparkline getItem(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "_NewEnum"
   * </p>
   */

  @VTID(12)
  java.util.Iterator<Com4jObject> iterator();

  /**
   * <p>
   * Getter method for the COM property "Location"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Range
   */

  @VTID(13)
  com.exceljava.com4j.excel.Range getLocation();


  /**
   * <p>
   * Setter method for the COM property "Location"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.Range parameter.
   */

  @VTID(14)
  void setLocation(
    com.exceljava.com4j.excel.Range rhs);


  /**
   * <p>
   * Getter method for the COM property "SourceData"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(15)
  java.lang.String getSourceData();


  /**
   * <p>
   * Setter method for the COM property "SourceData"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(16)
  void setSourceData(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "DateRange"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(17)
  java.lang.String getDateRange();


  /**
   * <p>
   * Setter method for the COM property "DateRange"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(18)
  void setDateRange(
    java.lang.String rhs);


  /**
   * @param location Mandatory com.exceljava.com4j.excel.Range parameter.
   */

  @VTID(19)
  void modifyLocation(
    com.exceljava.com4j.excel.Range location);


  /**
   * @param sourceData Mandatory java.lang.String parameter.
   */

  @VTID(20)
  void modifySourceData(
    java.lang.String sourceData);


  /**
   * @param location Mandatory com.exceljava.com4j.excel.Range parameter.
   * @param sourceData Mandatory java.lang.String parameter.
   */

  @VTID(21)
  void modify(
    com.exceljava.com4j.excel.Range location,
    java.lang.String sourceData);


  /**
   * @param dateRange Mandatory java.lang.String parameter.
   */

  @VTID(22)
  void modifyDateRange(
    java.lang.String dateRange);


  /**
   */

  @VTID(23)
  void delete();


  /**
   * <p>
   * Getter method for the COM property "Type"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlSparkType
   */

  @VTID(24)
  com.exceljava.com4j.excel.XlSparkType getType();


  /**
   * <p>
   * Setter method for the COM property "Type"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlSparkType parameter.
   */

  @VTID(25)
  void setType(
    com.exceljava.com4j.excel.XlSparkType rhs);


  /**
   * <p>
   * Getter method for the COM property "SeriesColor"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.FormatColor
   */

  @VTID(26)
  com.exceljava.com4j.excel.FormatColor getSeriesColor();


  /**
   * <p>
   * Getter method for the COM property "Points"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.SparkPoints
   */

  @VTID(27)
  com.exceljava.com4j.excel.SparkPoints getPoints();


  /**
   * <p>
   * Getter method for the COM property "Axes"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.SparkAxes
   */

  @VTID(28)
  com.exceljava.com4j.excel.SparkAxes getAxes();


  /**
   * <p>
   * Getter method for the COM property "DisplayBlanksAs"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlDisplayBlanksAs
   */

  @VTID(29)
  com.exceljava.com4j.excel.XlDisplayBlanksAs getDisplayBlanksAs();


  /**
   * <p>
   * Setter method for the COM property "DisplayBlanksAs"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlDisplayBlanksAs parameter.
   */

  @VTID(30)
  void setDisplayBlanksAs(
    com.exceljava.com4j.excel.XlDisplayBlanksAs rhs);


  /**
   * <p>
   * Getter method for the COM property "DisplayHidden"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(31)
  boolean getDisplayHidden();


  /**
   * <p>
   * Setter method for the COM property "DisplayHidden"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(32)
  void setDisplayHidden(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "LineWeight"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(33)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getLineWeight();


  /**
   * <p>
   * Setter method for the COM property "LineWeight"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(34)
  void setLineWeight(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "PlotBy"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlSparklineRowCol
   */

  @VTID(35)
  com.exceljava.com4j.excel.XlSparklineRowCol getPlotBy();


  /**
   * <p>
   * Setter method for the COM property "PlotBy"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlSparklineRowCol parameter.
   */

  @VTID(36)
  void setPlotBy(
    com.exceljava.com4j.excel.XlSparklineRowCol rhs);


  // Properties:
}
