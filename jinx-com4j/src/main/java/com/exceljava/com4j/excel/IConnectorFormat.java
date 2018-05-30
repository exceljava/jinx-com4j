package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{0002443E-0001-0000-C000-000000000046}")
public interface IConnectorFormat extends Com4jObject {
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
   * @param connectedShape Mandatory com.exceljava.com4j.excel.Shape parameter.
   * @param connectionSite Mandatory int parameter.
   */

  @VTID(10)
  void beginConnect(
    com.exceljava.com4j.excel.Shape connectedShape,
    int connectionSite);


  /**
   */

  @VTID(11)
  void beginDisconnect();


  /**
   * @param connectedShape Mandatory com.exceljava.com4j.excel.Shape parameter.
   * @param connectionSite Mandatory int parameter.
   */

  @VTID(12)
  void endConnect(
    com.exceljava.com4j.excel.Shape connectedShape,
    int connectionSite);


  /**
   */

  @VTID(13)
  void endDisconnect();


  /**
   * <p>
   * Getter method for the COM property "BeginConnected"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTriState
   */

  @VTID(14)
  com.exceljava.com4j.office.MsoTriState getBeginConnected();


  /**
   * <p>
   * Getter method for the COM property "BeginConnectedShape"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Shape
   */

  @VTID(15)
  com.exceljava.com4j.excel.Shape getBeginConnectedShape();


  /**
   * <p>
   * Getter method for the COM property "BeginConnectionSite"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(16)
  int getBeginConnectionSite();


  /**
   * <p>
   * Getter method for the COM property "EndConnected"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTriState
   */

  @VTID(17)
  com.exceljava.com4j.office.MsoTriState getEndConnected();


  /**
   * <p>
   * Getter method for the COM property "EndConnectedShape"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Shape
   */

  @VTID(18)
  com.exceljava.com4j.excel.Shape getEndConnectedShape();


  /**
   * <p>
   * Getter method for the COM property "EndConnectionSite"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(19)
  int getEndConnectionSite();


  /**
   * <p>
   * Getter method for the COM property "Type"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoConnectorType
   */

  @VTID(20)
  com.exceljava.com4j.office.MsoConnectorType getType();


  /**
   * <p>
   * Setter method for the COM property "Type"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.office.MsoConnectorType parameter.
   */

  @VTID(21)
  void setType(
    com.exceljava.com4j.office.MsoConnectorType rhs);


  // Properties:
}
