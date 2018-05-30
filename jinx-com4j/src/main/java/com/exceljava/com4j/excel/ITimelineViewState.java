package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{000244E0-0001-0000-C000-000000000046}")
public interface ITimelineViewState extends Com4jObject {
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
   * Getter method for the COM property "ShowHeader"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(10)
  boolean getShowHeader();


  /**
   * <p>
   * Setter method for the COM property "ShowHeader"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(11)
  void setShowHeader(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "ShowSelectionLabel"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(12)
  boolean getShowSelectionLabel();


  /**
   * <p>
   * Setter method for the COM property "ShowSelectionLabel"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(13)
  void setShowSelectionLabel(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "ShowTimeLevel"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(14)
  boolean getShowTimeLevel();


  /**
   * <p>
   * Setter method for the COM property "ShowTimeLevel"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(15)
  void setShowTimeLevel(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "ShowHorizontalScrollbar"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(16)
  boolean getShowHorizontalScrollbar();


  /**
   * <p>
   * Setter method for the COM property "ShowHorizontalScrollbar"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(17)
  void setShowHorizontalScrollbar(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "Level"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlTimelineLevel
   */

  @VTID(18)
  com.exceljava.com4j.excel.XlTimelineLevel getLevel();


  /**
   * <p>
   * Setter method for the COM property "Level"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlTimelineLevel parameter.
   */

  @VTID(19)
  void setLevel(
    com.exceljava.com4j.excel.XlTimelineLevel rhs);


  // Properties:
}
