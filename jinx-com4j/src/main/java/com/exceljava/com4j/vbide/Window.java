package com.exceljava.com4j.vbide  ;

import com4j.*;

@IID("{0002E16B-0000-0000-C000-000000000046}")
public interface Window extends Com4jObject {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "VBE"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.vbide.VBE
   */

  @DISPID(1) //= 0x1. The runtime will prefer the VTID if present
  @VTID(7)
  com.exceljava.com4j.vbide.VBE getVBE();


  /**
   * <p>
   * Getter method for the COM property "Collection"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.vbide._Windows
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(8)
  com.exceljava.com4j.vbide._Windows getCollection();


  /**
   */

  @DISPID(99) //= 0x63. The runtime will prefer the VTID if present
  @VTID(9)
  void close();


  /**
   * <p>
   * Getter method for the COM property "Caption"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(100) //= 0x64. The runtime will prefer the VTID if present
  @VTID(10)
  java.lang.String getCaption();


  /**
   * <p>
   * Getter method for the COM property "Visible"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(106) //= 0x6a. The runtime will prefer the VTID if present
  @VTID(11)
  boolean getVisible();


  /**
   * <p>
   * Setter method for the COM property "Visible"
   * </p>
   * @param pfVisible Mandatory boolean parameter.
   */

  @DISPID(106) //= 0x6a. The runtime will prefer the VTID if present
  @VTID(12)
  void setVisible(
    boolean pfVisible);


  /**
   * <p>
   * Getter method for the COM property "Left"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(101) //= 0x65. The runtime will prefer the VTID if present
  @VTID(13)
  int getLeft();


  /**
   * <p>
   * Setter method for the COM property "Left"
   * </p>
   * @param plLeft Mandatory int parameter.
   */

  @DISPID(101) //= 0x65. The runtime will prefer the VTID if present
  @VTID(14)
  void setLeft(
    int plLeft);


  /**
   * <p>
   * Getter method for the COM property "Top"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(103) //= 0x67. The runtime will prefer the VTID if present
  @VTID(15)
  int getTop();


  /**
   * <p>
   * Setter method for the COM property "Top"
   * </p>
   * @param plTop Mandatory int parameter.
   */

  @DISPID(103) //= 0x67. The runtime will prefer the VTID if present
  @VTID(16)
  void setTop(
    int plTop);


  /**
   * <p>
   * Getter method for the COM property "Width"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(105) //= 0x69. The runtime will prefer the VTID if present
  @VTID(17)
  int getWidth();


  /**
   * <p>
   * Setter method for the COM property "Width"
   * </p>
   * @param plWidth Mandatory int parameter.
   */

  @DISPID(105) //= 0x69. The runtime will prefer the VTID if present
  @VTID(18)
  void setWidth(
    int plWidth);


  /**
   * <p>
   * Getter method for the COM property "Height"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(107) //= 0x6b. The runtime will prefer the VTID if present
  @VTID(19)
  int getHeight();


  /**
   * <p>
   * Setter method for the COM property "Height"
   * </p>
   * @param plHeight Mandatory int parameter.
   */

  @DISPID(107) //= 0x6b. The runtime will prefer the VTID if present
  @VTID(20)
  void setHeight(
    int plHeight);


  /**
   * <p>
   * Getter method for the COM property "WindowState"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.vbide.vbext_WindowState
   */

  @DISPID(109) //= 0x6d. The runtime will prefer the VTID if present
  @VTID(21)
  com.exceljava.com4j.vbide.vbext_WindowState getWindowState();


  /**
   * <p>
   * Setter method for the COM property "WindowState"
   * </p>
   * @param plWindowState Mandatory com.exceljava.com4j.vbide.vbext_WindowState parameter.
   */

  @DISPID(109) //= 0x6d. The runtime will prefer the VTID if present
  @VTID(22)
  void setWindowState(
    com.exceljava.com4j.vbide.vbext_WindowState plWindowState);


  /**
   */

  @DISPID(111) //= 0x6f. The runtime will prefer the VTID if present
  @VTID(23)
  void setFocus();


  /**
   * <p>
   * Getter method for the COM property "Type"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.vbide.vbext_WindowType
   */

  @DISPID(112) //= 0x70. The runtime will prefer the VTID if present
  @VTID(24)
  com.exceljava.com4j.vbide.vbext_WindowType getType();


  /**
   * @param eKind Mandatory com.exceljava.com4j.vbide.vbext_WindowType parameter.
   */

  @DISPID(113) //= 0x71. The runtime will prefer the VTID if present
  @VTID(25)
  void setKind(
    com.exceljava.com4j.vbide.vbext_WindowType eKind);


  /**
   * <p>
   * Getter method for the COM property "LinkedWindows"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.vbide._LinkedWindows
   */

  @DISPID(116) //= 0x74. The runtime will prefer the VTID if present
  @VTID(26)
  com.exceljava.com4j.vbide._LinkedWindows getLinkedWindows();


  /**
   * <p>
   * Getter method for the COM property "LinkedWindowFrame"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.vbide.Window
   */

  @DISPID(117) //= 0x75. The runtime will prefer the VTID if present
  @VTID(27)
  com.exceljava.com4j.vbide.Window getLinkedWindowFrame();


  /**
   */

  @DISPID(118) //= 0x76. The runtime will prefer the VTID if present
  @VTID(28)
  void detach();


  /**
   * @param lWindowHandle Mandatory int parameter.
   */

  @DISPID(119) //= 0x77. The runtime will prefer the VTID if present
  @VTID(29)
  void attach(
    int lWindowHandle);


  /**
   * <p>
   * Getter method for the COM property "HWnd"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(120) //= 0x78. The runtime will prefer the VTID if present
  @VTID(30)
  int getHWnd();


  // Properties:
}
