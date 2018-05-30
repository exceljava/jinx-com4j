package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{919AA22C-B9AD-11D3-8D59-0050048384E3}")
public interface ILicValidator extends Com4jObject {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Products"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(1) //= 0x1. The runtime will prefer the VTID if present
  @VTID(7)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getProducts();


  /**
   * <p>
   * Getter method for the COM property "Selection"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(8)
  int getSelection();


  /**
   * <p>
   * Setter method for the COM property "Selection"
   * </p>
   * @param piSel Mandatory int parameter.
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(9)
  void setSelection(
    int piSel);


  // Properties:
}
