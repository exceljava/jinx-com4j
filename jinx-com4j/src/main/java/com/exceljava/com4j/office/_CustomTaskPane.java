package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C033B-0000-0000-C000-000000000046}")
public interface _CustomTaskPane extends Com4jObject {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Title"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(0) //= 0x0. The runtime will prefer the VTID if present
  @VTID(7)
  @DefaultMethod
  java.lang.String getTitle();


  /**
   * <p>
   * Getter method for the COM property "Application"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(1) //= 0x1. The runtime will prefer the VTID if present
  @VTID(8)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getApplication();


  /**
   * <p>
   * Getter method for the COM property "Window"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(9)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getWindow();


  /**
   * <p>
   * Getter method for the COM property "Visible"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(3) //= 0x3. The runtime will prefer the VTID if present
  @VTID(10)
  boolean getVisible();


  /**
   * <p>
   * Setter method for the COM property "Visible"
   * </p>
   * @param prop Mandatory boolean parameter.
   */

  @DISPID(3) //= 0x3. The runtime will prefer the VTID if present
  @VTID(11)
  void setVisible(
    boolean prop);


  /**
   * <p>
   * Getter method for the COM property "ContentControl"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(4) //= 0x4. The runtime will prefer the VTID if present
  @VTID(12)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getContentControl();


  /**
   * <p>
   * Getter method for the COM property "Height"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(5) //= 0x5. The runtime will prefer the VTID if present
  @VTID(13)
  int getHeight();


  /**
   * <p>
   * Setter method for the COM property "Height"
   * </p>
   * @param prop Mandatory int parameter.
   */

  @DISPID(5) //= 0x5. The runtime will prefer the VTID if present
  @VTID(14)
  void setHeight(
    int prop);


  /**
   * <p>
   * Getter method for the COM property "Width"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(6) //= 0x6. The runtime will prefer the VTID if present
  @VTID(15)
  int getWidth();


  /**
   * <p>
   * Setter method for the COM property "Width"
   * </p>
   * @param prop Mandatory int parameter.
   */

  @DISPID(6) //= 0x6. The runtime will prefer the VTID if present
  @VTID(16)
  void setWidth(
    int prop);


  /**
   * <p>
   * Getter method for the COM property "DockPosition"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoCTPDockPosition
   */

  @DISPID(7) //= 0x7. The runtime will prefer the VTID if present
  @VTID(17)
  com.exceljava.com4j.office.MsoCTPDockPosition getDockPosition();


  /**
   * <p>
   * Setter method for the COM property "DockPosition"
   * </p>
   * @param prop Mandatory com.exceljava.com4j.office.MsoCTPDockPosition parameter.
   */

  @DISPID(7) //= 0x7. The runtime will prefer the VTID if present
  @VTID(18)
  void setDockPosition(
    com.exceljava.com4j.office.MsoCTPDockPosition prop);


  /**
   * <p>
   * Getter method for the COM property "DockPositionRestrict"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoCTPDockPositionRestrict
   */

  @DISPID(8) //= 0x8. The runtime will prefer the VTID if present
  @VTID(19)
  com.exceljava.com4j.office.MsoCTPDockPositionRestrict getDockPositionRestrict();


  /**
   * <p>
   * Setter method for the COM property "DockPositionRestrict"
   * </p>
   * @param prop Mandatory com.exceljava.com4j.office.MsoCTPDockPositionRestrict parameter.
   */

  @DISPID(8) //= 0x8. The runtime will prefer the VTID if present
  @VTID(20)
  void setDockPositionRestrict(
    com.exceljava.com4j.office.MsoCTPDockPositionRestrict prop);


  /**
   */

  @DISPID(9) //= 0x9. The runtime will prefer the VTID if present
  @VTID(21)
  void delete();


  // Properties:
}
