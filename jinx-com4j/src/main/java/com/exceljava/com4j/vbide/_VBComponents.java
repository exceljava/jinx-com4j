package com.exceljava.com4j.vbide  ;

import com4j.*;

@IID("{EEE0091C-E393-11D1-BB03-00C04FB6C4A6}")
public interface _VBComponents extends com.exceljava.com4j.vbide._VBComponents_Old {
  // Methods:
  /**
   * @param progId Mandatory java.lang.String parameter.
   * @return  Returns a value of type com.exceljava.com4j.vbide._VBComponent
   */

  @DISPID(25) //= 0x19. The runtime will prefer the VTID if present
  @VTID(15)
  com.exceljava.com4j.vbide._VBComponent addCustom(
    java.lang.String progId);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>int parameter index is set to 0</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addMTDesigner(0);
   * </code>
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.vbide._VBComponent
   */

  @DISPID(26) //= 0x1a. The runtime will prefer the VTID if present
  @VTID(16)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {int.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"0"})
  @ReturnValue(index=1)
  com.exceljava.com4j.vbide._VBComponent addMTDesigner();

  /**
   * @param index Optional parameter. Default value is 0
   * @return  Returns a value of type com.exceljava.com4j.vbide._VBComponent
   */

  @DISPID(26) //= 0x1a. The runtime will prefer the VTID if present
  @VTID(16)
  com.exceljava.com4j.vbide._VBComponent addMTDesigner(
    @Optional @DefaultValue("0") int index);


  // Properties:
}
