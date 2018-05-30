package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C03E6-0000-0000-C000-000000000046}")
public interface PickerDialog extends com.exceljava.com4j.office._IMsoDispObj {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "DataHandlerId"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(1) //= 0x1. The runtime will prefer the VTID if present
  @VTID(9)
  java.lang.String getDataHandlerId();


  /**
   * <p>
   * Setter method for the COM property "DataHandlerId"
   * </p>
   * @param id Mandatory java.lang.String parameter.
   */

  @DISPID(1) //= 0x1. The runtime will prefer the VTID if present
  @VTID(10)
  void setDataHandlerId(
    java.lang.String id);


  /**
   * <p>
   * Getter method for the COM property "Title"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(11)
  java.lang.String getTitle();


  /**
   * <p>
   * Setter method for the COM property "Title"
   * </p>
   * @param title Mandatory java.lang.String parameter.
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(12)
  void setTitle(
    java.lang.String title);


  /**
   * <p>
   * Getter method for the COM property "Properties"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.PickerProperties
   */

  @DISPID(3) //= 0x3. The runtime will prefer the VTID if present
  @VTID(13)
  com.exceljava.com4j.office.PickerProperties getProperties();


  @VTID(13)
  @ReturnValue(defaultPropertyThrough={com.exceljava.com4j.office.PickerProperties.class})
  com.exceljava.com4j.office.PickerProperty getProperties(
    int index);

  /**
   * @return  Returns a value of type com.exceljava.com4j.office.PickerResults
   */

  @DISPID(4) //= 0x4. The runtime will prefer the VTID if present
  @VTID(14)
  com.exceljava.com4j.office.PickerResults createPickerResults();


  @VTID(14)
  @ReturnValue(defaultPropertyThrough={com.exceljava.com4j.office.PickerResults.class})
  com.exceljava.com4j.office.PickerResult createPickerResults(
    int index);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>boolean parameter isMultiSelect is set to false</li><li>com.exceljava.com4j.office.PickerResults parameter existingResults is set to 0</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * show(false, 0);
   * </code>
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.PickerResults
   */

  @DISPID(5) //= 0x5. The runtime will prefer the VTID if present
  @VTID(15)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {boolean.class, com.exceljava.com4j.office.PickerResults.class}, nativeType = {NativeType.VariantBool, NativeType.ComObject}, variantType = {Variant.Type.VT_BOOL, Variant.Type.VT_I2}, literal = {"false", "0"})
  @ReturnValue(index=2)
  com.exceljava.com4j.office.PickerResults show();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.office.PickerResults parameter existingResults is set to 0</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * show(isMultiSelect, 0);
   * </code>
   * </p>
   * @param isMultiSelect Optional parameter. Default value is false
   * @return  Returns a value of type com.exceljava.com4j.office.PickerResults
   */

  @DISPID(5) //= 0x5. The runtime will prefer the VTID if present
  @VTID(15)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {com.exceljava.com4j.office.PickerResults.class}, nativeType = {NativeType.ComObject}, variantType = {Variant.Type.VT_I2}, literal = {"0"})
  @ReturnValue(index=2)
  com.exceljava.com4j.office.PickerResults show(
    @Optional @DefaultValue("-1") boolean isMultiSelect);

  /**
   * @param isMultiSelect Optional parameter. Default value is false
   * @param existingResults Optional parameter. Default value is 0
   * @return  Returns a value of type com.exceljava.com4j.office.PickerResults
   */

  @DISPID(5) //= 0x5. The runtime will prefer the VTID if present
  @VTID(15)
  com.exceljava.com4j.office.PickerResults show(
    @Optional @DefaultValue("-1") boolean isMultiSelect,
    @Optional @DefaultValue("0") com.exceljava.com4j.office.PickerResults existingResults);


  /**
   * @param tokenText Mandatory java.lang.String parameter.
   * @param duplicateDlgMode Mandatory int parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.PickerResults
   */

  @DISPID(6) //= 0x6. The runtime will prefer the VTID if present
  @VTID(16)
  com.exceljava.com4j.office.PickerResults resolve(
    java.lang.String tokenText,
    int duplicateDlgMode);


  // Properties:
}
