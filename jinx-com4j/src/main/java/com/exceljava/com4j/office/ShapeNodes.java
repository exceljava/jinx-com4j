package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C0319-0000-0000-C000-000000000046}")
public interface ShapeNodes extends com.exceljava.com4j.office._IMsoDispObj,Iterable<Com4jObject> {
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
   * Getter method for the COM property "Count"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(10)
  int getCount();


  /**
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.ShapeNode
   */

  @DISPID(0) //= 0x0. The runtime will prefer the VTID if present
  @VTID(11)
  @DefaultMethod
  com.exceljava.com4j.office.ShapeNode item(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   * <p>
   * Getter method for the COM property "_NewEnum"
   * </p>
   */

  @DISPID(-4) //= 0xfffffffc. The runtime will prefer the VTID if present
  @VTID(12)
  java.util.Iterator<Com4jObject> iterator();

  /**
   * @param index Mandatory int parameter.
   */

  @DISPID(11) //= 0xb. The runtime will prefer the VTID if present
  @VTID(13)
  void delete(
    int index);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>float parameter x2 is set to 0.0f</li><li>float parameter y2 is set to 0.0f</li><li>float parameter x3 is set to 0.0f</li><li>float parameter y3 is set to 0.0f</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * insert(index, segmentType, editingType, x1, y1, 0.0f, 0.0f, 0.0f, 0.0f);
   * </code>
   * </p>
   * @param index Mandatory int parameter.
   * @param segmentType Mandatory com.exceljava.com4j.office.MsoSegmentType parameter.
   * @param editingType Mandatory com.exceljava.com4j.office.MsoEditingType parameter.
   * @param x1 Mandatory float parameter.
   * @param y1 Mandatory float parameter.
   */

  @DISPID(12) //= 0xc. The runtime will prefer the VTID if present
  @VTID(14)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5, 6, 7, 8}, javaType = {float.class, float.class, float.class, float.class}, nativeType = {NativeType.Float, NativeType.Float, NativeType.Float, NativeType.Float}, variantType = {Variant.Type.VT_R4, Variant.Type.VT_R4, Variant.Type.VT_R4, Variant.Type.VT_R4}, literal = {"0.0", "0.0", "0.0", "0.0"})
  void insert(
    int index,
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
   * insert(index, segmentType, editingType, x1, y1, x2, 0.0f, 0.0f, 0.0f);
   * </code>
   * </p>
   * @param index Mandatory int parameter.
   * @param segmentType Mandatory com.exceljava.com4j.office.MsoSegmentType parameter.
   * @param editingType Mandatory com.exceljava.com4j.office.MsoEditingType parameter.
   * @param x1 Mandatory float parameter.
   * @param y1 Mandatory float parameter.
   * @param x2 Optional parameter. Default value is 0.0f
   */

  @DISPID(12) //= 0xc. The runtime will prefer the VTID if present
  @VTID(14)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5}, optParamIndex = {6, 7, 8}, javaType = {float.class, float.class, float.class}, nativeType = {NativeType.Float, NativeType.Float, NativeType.Float}, variantType = {Variant.Type.VT_R4, Variant.Type.VT_R4, Variant.Type.VT_R4}, literal = {"0.0", "0.0", "0.0"})
  void insert(
    int index,
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
   * insert(index, segmentType, editingType, x1, y1, x2, y2, 0.0f, 0.0f);
   * </code>
   * </p>
   * @param index Mandatory int parameter.
   * @param segmentType Mandatory com.exceljava.com4j.office.MsoSegmentType parameter.
   * @param editingType Mandatory com.exceljava.com4j.office.MsoEditingType parameter.
   * @param x1 Mandatory float parameter.
   * @param y1 Mandatory float parameter.
   * @param x2 Optional parameter. Default value is 0.0f
   * @param y2 Optional parameter. Default value is 0.0f
   */

  @DISPID(12) //= 0xc. The runtime will prefer the VTID if present
  @VTID(14)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6}, optParamIndex = {7, 8}, javaType = {float.class, float.class}, nativeType = {NativeType.Float, NativeType.Float}, variantType = {Variant.Type.VT_R4, Variant.Type.VT_R4}, literal = {"0.0", "0.0"})
  void insert(
    int index,
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
   * insert(index, segmentType, editingType, x1, y1, x2, y2, x3, 0.0f);
   * </code>
   * </p>
   * @param index Mandatory int parameter.
   * @param segmentType Mandatory com.exceljava.com4j.office.MsoSegmentType parameter.
   * @param editingType Mandatory com.exceljava.com4j.office.MsoEditingType parameter.
   * @param x1 Mandatory float parameter.
   * @param y1 Mandatory float parameter.
   * @param x2 Optional parameter. Default value is 0.0f
   * @param y2 Optional parameter. Default value is 0.0f
   * @param x3 Optional parameter. Default value is 0.0f
   */

  @DISPID(12) //= 0xc. The runtime will prefer the VTID if present
  @VTID(14)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6, 7}, optParamIndex = {8}, javaType = {float.class}, nativeType = {NativeType.Float}, variantType = {Variant.Type.VT_R4}, literal = {"0.0"})
  void insert(
    int index,
    com.exceljava.com4j.office.MsoSegmentType segmentType,
    com.exceljava.com4j.office.MsoEditingType editingType,
    float x1,
    float y1,
    @Optional @DefaultValue("0") float x2,
    @Optional @DefaultValue("0") float y2,
    @Optional @DefaultValue("0") float x3);

  /**
   * @param index Mandatory int parameter.
   * @param segmentType Mandatory com.exceljava.com4j.office.MsoSegmentType parameter.
   * @param editingType Mandatory com.exceljava.com4j.office.MsoEditingType parameter.
   * @param x1 Mandatory float parameter.
   * @param y1 Mandatory float parameter.
   * @param x2 Optional parameter. Default value is 0.0f
   * @param y2 Optional parameter. Default value is 0.0f
   * @param x3 Optional parameter. Default value is 0.0f
   * @param y3 Optional parameter. Default value is 0.0f
   */

  @DISPID(12) //= 0xc. The runtime will prefer the VTID if present
  @VTID(14)
  void insert(
    int index,
    com.exceljava.com4j.office.MsoSegmentType segmentType,
    com.exceljava.com4j.office.MsoEditingType editingType,
    float x1,
    float y1,
    @Optional @DefaultValue("0") float x2,
    @Optional @DefaultValue("0") float y2,
    @Optional @DefaultValue("0") float x3,
    @Optional @DefaultValue("0") float y3);


  /**
   * @param index Mandatory int parameter.
   * @param editingType Mandatory com.exceljava.com4j.office.MsoEditingType parameter.
   */

  @DISPID(13) //= 0xd. The runtime will prefer the VTID if present
  @VTID(15)
  void setEditingType(
    int index,
    com.exceljava.com4j.office.MsoEditingType editingType);


  /**
   * @param index Mandatory int parameter.
   * @param x1 Mandatory float parameter.
   * @param y1 Mandatory float parameter.
   */

  @DISPID(14) //= 0xe. The runtime will prefer the VTID if present
  @VTID(16)
  void setPosition(
    int index,
    float x1,
    float y1);


  /**
   * @param index Mandatory int parameter.
   * @param segmentType Mandatory com.exceljava.com4j.office.MsoSegmentType parameter.
   */

  @DISPID(15) //= 0xf. The runtime will prefer the VTID if present
  @VTID(17)
  void setSegmentType(
    int index,
    com.exceljava.com4j.office.MsoSegmentType segmentType);


  // Properties:
}
