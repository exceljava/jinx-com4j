package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C03F1-0000-0000-C000-000000000046}")
public interface ContactCard extends com.exceljava.com4j.office._IMsoDispObj {
  // Methods:
  /**
   */

  @DISPID(1) //= 0x1. The runtime will prefer the VTID if present
  @VTID(9)
  void close();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>boolean parameter showWithDelay is set to false</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * show(cardStyle, rectangleLeft, rectangleRight, rectangleTop, rectangleBottom, horizontalPosition, false);
   * </code>
   * </p>
   * @param cardStyle Mandatory com.exceljava.com4j.office.MsoContactCardStyle parameter.
   * @param rectangleLeft Mandatory int parameter.
   * @param rectangleRight Mandatory int parameter.
   * @param rectangleTop Mandatory int parameter.
   * @param rectangleBottom Mandatory int parameter.
   * @param horizontalPosition Mandatory int parameter.
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5}, optParamIndex = {6}, javaType = {boolean.class}, nativeType = {NativeType.VariantBool}, variantType = {Variant.Type.VT_BOOL}, literal = {"false"})
  void show(
    com.exceljava.com4j.office.MsoContactCardStyle cardStyle,
    int rectangleLeft,
    int rectangleRight,
    int rectangleTop,
    int rectangleBottom,
    int horizontalPosition);

  /**
   * @param cardStyle Mandatory com.exceljava.com4j.office.MsoContactCardStyle parameter.
   * @param rectangleLeft Mandatory int parameter.
   * @param rectangleRight Mandatory int parameter.
   * @param rectangleTop Mandatory int parameter.
   * @param rectangleBottom Mandatory int parameter.
   * @param horizontalPosition Mandatory int parameter.
   * @param showWithDelay Optional parameter. Default value is false
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(10)
  void show(
    com.exceljava.com4j.office.MsoContactCardStyle cardStyle,
    int rectangleLeft,
    int rectangleRight,
    int rectangleTop,
    int rectangleBottom,
    int horizontalPosition,
    @Optional @DefaultValue("0") boolean showWithDelay);


  // Properties:
}
