package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C171D-0000-0000-C000-000000000046}")
public interface ChartColorFormat extends Com4jObject {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Parent"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(7)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getParent();


  /**
   * <p>
   * Getter method for the COM property "SchemeColor"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(8)
  int getSchemeColor();


  /**
   * <p>
   * Setter method for the COM property "SchemeColor"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(9)
  void setSchemeColor(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "RGB"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(10)
  int getRGB();


  /**
   * <p>
   * Setter method for the COM property "RGB"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(11)
  void setRGB(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "_Default"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(12)
  @DefaultMethod
  int get_Default();


  /**
   * <p>
   * Getter method for the COM property "Type"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(13)
  int getType();


  /**
   * <p>
   * Getter method for the COM property "Application"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(14)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getApplication();


  /**
   * <p>
   * Getter method for the COM property "Creator"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(15)
  int getCreator();


  // Properties:
}
