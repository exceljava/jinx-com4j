package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{000C0311-0000-0000-C000-000000000046}")
public interface CalloutFormat extends com.exceljava.com4j.office._IMsoDispObj {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Parent"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(1) //= 0x1. The runtime will prefer the VTID if present
  @VTID(9)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getParent();


  /**
   */

  @DISPID(10) //= 0xa. The runtime will prefer the VTID if present
  @VTID(10)
  void automaticLength();


  /**
   * @param drop Mandatory float parameter.
   */

  @DISPID(11) //= 0xb. The runtime will prefer the VTID if present
  @VTID(11)
  void customDrop(
    float drop);


  /**
   * @param length Mandatory float parameter.
   */

  @DISPID(12) //= 0xc. The runtime will prefer the VTID if present
  @VTID(12)
  void customLength(
    float length);


  /**
   * @param dropType Mandatory com.exceljava.com4j.office.MsoCalloutDropType parameter.
   */

  @DISPID(13) //= 0xd. The runtime will prefer the VTID if present
  @VTID(13)
  void presetDrop(
    com.exceljava.com4j.office.MsoCalloutDropType dropType);


  /**
   * <p>
   * Getter method for the COM property "Accent"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTriState
   */

  @DISPID(100) //= 0x64. The runtime will prefer the VTID if present
  @VTID(14)
  com.exceljava.com4j.office.MsoTriState getAccent();


  /**
   * <p>
   * Setter method for the COM property "Accent"
   * </p>
   * @param accent Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   */

  @DISPID(100) //= 0x64. The runtime will prefer the VTID if present
  @VTID(15)
  void setAccent(
    com.exceljava.com4j.office.MsoTriState accent);


  /**
   * <p>
   * Getter method for the COM property "Angle"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoCalloutAngleType
   */

  @DISPID(101) //= 0x65. The runtime will prefer the VTID if present
  @VTID(16)
  com.exceljava.com4j.office.MsoCalloutAngleType getAngle();


  /**
   * <p>
   * Setter method for the COM property "Angle"
   * </p>
   * @param angle Mandatory com.exceljava.com4j.office.MsoCalloutAngleType parameter.
   */

  @DISPID(101) //= 0x65. The runtime will prefer the VTID if present
  @VTID(17)
  void setAngle(
    com.exceljava.com4j.office.MsoCalloutAngleType angle);


  /**
   * <p>
   * Getter method for the COM property "AutoAttach"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTriState
   */

  @DISPID(102) //= 0x66. The runtime will prefer the VTID if present
  @VTID(18)
  com.exceljava.com4j.office.MsoTriState getAutoAttach();


  /**
   * <p>
   * Setter method for the COM property "AutoAttach"
   * </p>
   * @param autoAttach Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   */

  @DISPID(102) //= 0x66. The runtime will prefer the VTID if present
  @VTID(19)
  void setAutoAttach(
    com.exceljava.com4j.office.MsoTriState autoAttach);


  /**
   * <p>
   * Getter method for the COM property "AutoLength"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTriState
   */

  @DISPID(103) //= 0x67. The runtime will prefer the VTID if present
  @VTID(20)
  com.exceljava.com4j.office.MsoTriState getAutoLength();


  /**
   * <p>
   * Getter method for the COM property "Border"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTriState
   */

  @DISPID(104) //= 0x68. The runtime will prefer the VTID if present
  @VTID(21)
  com.exceljava.com4j.office.MsoTriState getBorder();


  /**
   * <p>
   * Setter method for the COM property "Border"
   * </p>
   * @param border Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   */

  @DISPID(104) //= 0x68. The runtime will prefer the VTID if present
  @VTID(22)
  void setBorder(
    com.exceljava.com4j.office.MsoTriState border);


  /**
   * <p>
   * Getter method for the COM property "Drop"
   * </p>
   * @return  Returns a value of type float
   */

  @DISPID(105) //= 0x69. The runtime will prefer the VTID if present
  @VTID(23)
  float getDrop();


  /**
   * <p>
   * Getter method for the COM property "DropType"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoCalloutDropType
   */

  @DISPID(106) //= 0x6a. The runtime will prefer the VTID if present
  @VTID(24)
  com.exceljava.com4j.office.MsoCalloutDropType getDropType();


  /**
   * <p>
   * Getter method for the COM property "Gap"
   * </p>
   * @return  Returns a value of type float
   */

  @DISPID(107) //= 0x6b. The runtime will prefer the VTID if present
  @VTID(25)
  float getGap();


  /**
   * <p>
   * Setter method for the COM property "Gap"
   * </p>
   * @param gap Mandatory float parameter.
   */

  @DISPID(107) //= 0x6b. The runtime will prefer the VTID if present
  @VTID(26)
  void setGap(
    float gap);


  /**
   * <p>
   * Getter method for the COM property "Length"
   * </p>
   * @return  Returns a value of type float
   */

  @DISPID(108) //= 0x6c. The runtime will prefer the VTID if present
  @VTID(27)
  float getLength();


  /**
   * <p>
   * Getter method for the COM property "Type"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoCalloutType
   */

  @DISPID(109) //= 0x6d. The runtime will prefer the VTID if present
  @VTID(28)
  com.exceljava.com4j.office.MsoCalloutType getType();


  /**
   * <p>
   * Setter method for the COM property "Type"
   * </p>
   * @param type Mandatory com.exceljava.com4j.office.MsoCalloutType parameter.
   */

  @DISPID(109) //= 0x6d. The runtime will prefer the VTID if present
  @VTID(29)
  void setType(
    com.exceljava.com4j.office.MsoCalloutType type);


  // Properties:
}
