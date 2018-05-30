package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C0315-0000-0000-C000-000000000046}")
public interface FreeformBuilder extends com.exceljava.com4j.office._IMsoDispObj {
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
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>float parameter x2 is set to 0.0f</li><li>float parameter y2 is set to 0.0f</li><li>float parameter x3 is set to 0.0f</li><li>float parameter y3 is set to 0.0f</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addNodes(segmentType, editingType, x1, y1, 0.0f, 0.0f, 0.0f, 0.0f);
   * </code>
   * </p>
   * @param segmentType Mandatory com.exceljava.com4j.office.MsoSegmentType parameter.
   * @param editingType Mandatory com.exceljava.com4j.office.MsoEditingType parameter.
   * @param x1 Mandatory float parameter.
   * @param y1 Mandatory float parameter.
   */

  @DISPID(10) //= 0xa. The runtime will prefer the VTID if present
  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5, 6, 7}, javaType = {float.class, float.class, float.class, float.class}, nativeType = {NativeType.Float, NativeType.Float, NativeType.Float, NativeType.Float}, variantType = {Variant.Type.VT_R4, Variant.Type.VT_R4, Variant.Type.VT_R4, Variant.Type.VT_R4}, literal = {"0.0", "0.0", "0.0", "0.0"})
  void addNodes(
    com.exceljava.com4j.office.MsoSegmentType segmentType,
    com.exceljava.com4j.office.MsoEditingType editingType,
    float x1,
    float y1);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>float parameter y2 is set to 0.0f</li><li>float parameter x3 is set to 0.0f</li><li>float parameter y3 is set to 0.0f</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addNodes(segmentType, editingType, x1, y1, x2, 0.0f, 0.0f, 0.0f);
   * </code>
   * </p>
   * @param segmentType Mandatory com.exceljava.com4j.office.MsoSegmentType parameter.
   * @param editingType Mandatory com.exceljava.com4j.office.MsoEditingType parameter.
   * @param x1 Mandatory float parameter.
   * @param y1 Mandatory float parameter.
   * @param x2 Optional parameter. Default value is 0.0f
   */

  @DISPID(10) //= 0xa. The runtime will prefer the VTID if present
  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5, 6, 7}, javaType = {float.class, float.class, float.class}, nativeType = {NativeType.Float, NativeType.Float, NativeType.Float}, variantType = {Variant.Type.VT_R4, Variant.Type.VT_R4, Variant.Type.VT_R4}, literal = {"0.0", "0.0", "0.0"})
  void addNodes(
    com.exceljava.com4j.office.MsoSegmentType segmentType,
    com.exceljava.com4j.office.MsoEditingType editingType,
    float x1,
    float y1,
    @Optional @DefaultValue("0") float x2);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>float parameter x3 is set to 0.0f</li><li>float parameter y3 is set to 0.0f</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addNodes(segmentType, editingType, x1, y1, x2, y2, 0.0f, 0.0f);
   * </code>
   * </p>
   * @param segmentType Mandatory com.exceljava.com4j.office.MsoSegmentType parameter.
   * @param editingType Mandatory com.exceljava.com4j.office.MsoEditingType parameter.
   * @param x1 Mandatory float parameter.
   * @param y1 Mandatory float parameter.
   * @param x2 Optional parameter. Default value is 0.0f
   * @param y2 Optional parameter. Default value is 0.0f
   */

  @DISPID(10) //= 0xa. The runtime will prefer the VTID if present
  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5}, optParamIndex = {6, 7}, javaType = {float.class, float.class}, nativeType = {NativeType.Float, NativeType.Float}, variantType = {Variant.Type.VT_R4, Variant.Type.VT_R4}, literal = {"0.0", "0.0"})
  void addNodes(
    com.exceljava.com4j.office.MsoSegmentType segmentType,
    com.exceljava.com4j.office.MsoEditingType editingType,
    float x1,
    float y1,
    @Optional @DefaultValue("0") float x2,
    @Optional @DefaultValue("0") float y2);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>float parameter y3 is set to 0.0f</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addNodes(segmentType, editingType, x1, y1, x2, y2, x3, 0.0f);
   * </code>
   * </p>
   * @param segmentType Mandatory com.exceljava.com4j.office.MsoSegmentType parameter.
   * @param editingType Mandatory com.exceljava.com4j.office.MsoEditingType parameter.
   * @param x1 Mandatory float parameter.
   * @param y1 Mandatory float parameter.
   * @param x2 Optional parameter. Default value is 0.0f
   * @param y2 Optional parameter. Default value is 0.0f
   * @param x3 Optional parameter. Default value is 0.0f
   */

  @DISPID(10) //= 0xa. The runtime will prefer the VTID if present
  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6}, optParamIndex = {7}, javaType = {float.class}, nativeType = {NativeType.Float}, variantType = {Variant.Type.VT_R4}, literal = {"0.0"})
  void addNodes(
    com.exceljava.com4j.office.MsoSegmentType segmentType,
    com.exceljava.com4j.office.MsoEditingType editingType,
    float x1,
    float y1,
    @Optional @DefaultValue("0") float x2,
    @Optional @DefaultValue("0") float y2,
    @Optional @DefaultValue("0") float x3);

  /**
   * @param segmentType Mandatory com.exceljava.com4j.office.MsoSegmentType parameter.
   * @param editingType Mandatory com.exceljava.com4j.office.MsoEditingType parameter.
   * @param x1 Mandatory float parameter.
   * @param y1 Mandatory float parameter.
   * @param x2 Optional parameter. Default value is 0.0f
   * @param y2 Optional parameter. Default value is 0.0f
   * @param x3 Optional parameter. Default value is 0.0f
   * @param y3 Optional parameter. Default value is 0.0f
   */

  @DISPID(10) //= 0xa. The runtime will prefer the VTID if present
  @VTID(10)
  void addNodes(
    com.exceljava.com4j.office.MsoSegmentType segmentType,
    com.exceljava.com4j.office.MsoEditingType editingType,
    float x1,
    float y1,
    @Optional @DefaultValue("0") float x2,
    @Optional @DefaultValue("0") float y2,
    @Optional @DefaultValue("0") float x3,
    @Optional @DefaultValue("0") float y3);


  /**
   * @return  Returns a value of type com.exceljava.com4j.office.Shape
   */

  @DISPID(11) //= 0xb. The runtime will prefer the VTID if present
  @VTID(11)
  com.exceljava.com4j.office.Shape convertToShape();


  // Properties:
}
