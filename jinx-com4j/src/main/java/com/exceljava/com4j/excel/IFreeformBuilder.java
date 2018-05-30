package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{0002443F-0001-0000-C000-000000000046}")
public interface IFreeformBuilder extends Com4jObject {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Application"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel._Application
   */

  @VTID(7)
  com.exceljava.com4j.excel._Application getApplication();


  /**
   * <p>
   * Getter method for the COM property "Creator"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlCreator
   */

  @VTID(8)
  com.exceljava.com4j.excel.XlCreator getCreator();


  /**
   * <p>
   * Getter method for the COM property "Parent"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(9)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getParent();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter x2 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter y2 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter x3 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter y3 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addNodes(segmentType, editingType, x1, y1, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param segmentType Mandatory com.exceljava.com4j.office.MsoSegmentType parameter.
   * @param editingType Mandatory com.exceljava.com4j.office.MsoEditingType parameter.
   * @param x1 Mandatory float parameter.
   * @param y1 Mandatory float parameter.
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4, 5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
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
     * <li>java.lang.Object parameter y2 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter x3 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter y3 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addNodes(segmentType, editingType, x1, y1, x2, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param segmentType Mandatory com.exceljava.com4j.office.MsoSegmentType parameter.
   * @param editingType Mandatory com.exceljava.com4j.office.MsoEditingType parameter.
   * @param x1 Mandatory float parameter.
   * @param y1 Mandatory float parameter.
   * @param x2 Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4}, optParamIndex = {5, 6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  void addNodes(
    com.exceljava.com4j.office.MsoSegmentType segmentType,
    com.exceljava.com4j.office.MsoEditingType editingType,
    float x1,
    float y1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object x2);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter x3 is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter y3 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addNodes(segmentType, editingType, x1, y1, x2, y2, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param segmentType Mandatory com.exceljava.com4j.office.MsoSegmentType parameter.
   * @param editingType Mandatory com.exceljava.com4j.office.MsoEditingType parameter.
   * @param x1 Mandatory float parameter.
   * @param y1 Mandatory float parameter.
   * @param x2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param y2 Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5}, optParamIndex = {6, 7}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  void addNodes(
    com.exceljava.com4j.office.MsoSegmentType segmentType,
    com.exceljava.com4j.office.MsoEditingType editingType,
    float x1,
    float y1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object x2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object y2);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter y3 is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addNodes(segmentType, editingType, x1, y1, x2, y2, x3, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param segmentType Mandatory com.exceljava.com4j.office.MsoSegmentType parameter.
   * @param editingType Mandatory com.exceljava.com4j.office.MsoEditingType parameter.
   * @param x1 Mandatory float parameter.
   * @param y1 Mandatory float parameter.
   * @param x2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param y2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param x3 Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 4, 5, 6}, optParamIndex = {7}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void addNodes(
    com.exceljava.com4j.office.MsoSegmentType segmentType,
    com.exceljava.com4j.office.MsoEditingType editingType,
    float x1,
    float y1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object x2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object y2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object x3);

  /**
   * @param segmentType Mandatory com.exceljava.com4j.office.MsoSegmentType parameter.
   * @param editingType Mandatory com.exceljava.com4j.office.MsoEditingType parameter.
   * @param x1 Mandatory float parameter.
   * @param y1 Mandatory float parameter.
   * @param x2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param y2 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param x3 Optional parameter. Default value is com4j.Variant.getMissing()
   * @param y3 Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @VTID(10)
  void addNodes(
    com.exceljava.com4j.office.MsoSegmentType segmentType,
    com.exceljava.com4j.office.MsoEditingType editingType,
    float x1,
    float y1,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object x2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object y2,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object x3,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object y3);


  /**
   * @return  Returns a value of type com.exceljava.com4j.excel.Shape
   */

  @VTID(11)
  com.exceljava.com4j.excel.Shape convertToShape();


  // Properties:
}
