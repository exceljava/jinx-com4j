package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C0377-0000-0000-C000-000000000046}")
public interface SmartDocument extends com.exceljava.com4j.office._IMsoDispObj {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "SolutionID"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(1) //= 0x1. The runtime will prefer the VTID if present
  @VTID(9)
  java.lang.String getSolutionID();


  /**
   * <p>
   * Setter method for the COM property "SolutionID"
   * </p>
   * @param pbstrID Mandatory java.lang.String parameter.
   */

  @DISPID(1) //= 0x1. The runtime will prefer the VTID if present
  @VTID(10)
  void setSolutionID(
    java.lang.String pbstrID);


  /**
   * <p>
   * Getter method for the COM property "SolutionURL"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(11)
  java.lang.String getSolutionURL();


  /**
   * <p>
   * Setter method for the COM property "SolutionURL"
   * </p>
   * @param pbstrUrl Mandatory java.lang.String parameter.
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(12)
  void setSolutionURL(
    java.lang.String pbstrUrl);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>boolean parameter considerAllSchemas is set to false</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * pickSolution(false);
   * </code>
   * </p>
   */

  @DISPID(3) //= 0x3. The runtime will prefer the VTID if present
  @VTID(13)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {boolean.class}, nativeType = {NativeType.VariantBool}, variantType = {Variant.Type.VT_BOOL}, literal = {"false"})
  void pickSolution();

  /**
   * @param considerAllSchemas Optional parameter. Default value is false
   */

  @DISPID(3) //= 0x3. The runtime will prefer the VTID if present
  @VTID(13)
  void pickSolution(
    @Optional @DefaultValue("0") boolean considerAllSchemas);


  /**
   */

  @DISPID(4) //= 0x4. The runtime will prefer the VTID if present
  @VTID(14)
  void refreshPane();


  // Properties:
}
