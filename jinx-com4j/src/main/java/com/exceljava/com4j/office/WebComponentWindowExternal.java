package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000CD101-0000-0000-C000-000000000046}")
public interface WebComponentWindowExternal extends Com4jObject {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "InterfaceVersion"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(1) //= 0x1. The runtime will prefer the VTID if present
  @VTID(7)
  int getInterfaceVersion();


  /**
   * <p>
   * Getter method for the COM property "ApplicationName"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(8)
  java.lang.String getApplicationName();


  /**
   * <p>
   * Getter method for the COM property "ApplicationVersion"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(3) //= 0x3. The runtime will prefer the VTID if present
  @VTID(9)
  int getApplicationVersion();


  /**
   * <p>
   * Getter method for the COM property "Application"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(4) //= 0x4. The runtime will prefer the VTID if present
  @VTID(10)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getApplication();


  /**
   */

  @DISPID(5) //= 0x5. The runtime will prefer the VTID if present
  @VTID(11)
  void closeWindow();


  /**
   * <p>
   * Getter method for the COM property "WebComponent"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.WebComponent
   */

  @DISPID(6) //= 0x6. The runtime will prefer the VTID if present
  @VTID(12)
  com.exceljava.com4j.office.WebComponent getWebComponent();


  // Properties:
}
