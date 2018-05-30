package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020400-0000-0000-C000-000000000046}")
public interface Comment extends Com4jObject {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Application"
   * </p>
   */

  @DISPID(148)
  @PropGet
  com.exceljava.com4j.excel._Application getApplication();


  /**
   * <p>
   * Getter method for the COM property "Creator"
   * </p>
   */

  @DISPID(149)
  @PropGet
  com.exceljava.com4j.excel.XlCreator getCreator();


  /**
   * <p>
   * Getter method for the COM property "Parent"
   * </p>
   */

  @DISPID(150)
  @PropGet
  com4j.Com4jObject getParent();


  /**
   * <p>
   * Getter method for the COM property "Author"
   * </p>
   */

  @DISPID(574)
  @PropGet
  java.lang.String getAuthor();


  /**
   * <p>
   * Getter method for the COM property "Shape"
   * </p>
   */

  @DISPID(1582)
  @PropGet
  com.exceljava.com4j.excel.Shape getShape();


  /**
   * <p>
   * Getter method for the COM property "Visible"
   * </p>
   */

  @DISPID(558)
  @PropGet
  boolean getVisible();


  /**
   * <p>
   * Setter method for the COM property "Visible"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @DISPID(558)
  @PropPut
  void setVisible(
    boolean rhs);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter text is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter start is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter overwrite is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * text(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(138)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.String text();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter start is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter overwrite is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * text(text, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param text Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(138)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=-1)
  java.lang.String text(
    @Optional java.lang.Object text);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter overwrite is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * text(text, start, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param text Optional parameter. Default value is com4j.Variant.getMissing()
   * @param start Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(138)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=-1)
  java.lang.String text(
    @Optional java.lang.Object text,
    @Optional java.lang.Object start);

  /**
   * @param text Optional parameter. Default value is com4j.Variant.getMissing()
   * @param start Optional parameter. Default value is com4j.Variant.getMissing()
   * @param overwrite Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(138)
  java.lang.String text(
    @Optional java.lang.Object text,
    @Optional java.lang.Object start,
    @Optional java.lang.Object overwrite);


  /**
   */

  @DISPID(117)
  void delete();


  /**
   */

  @DISPID(502)
  com.exceljava.com4j.excel.Comment next();


  /**
   */

  @DISPID(503)
  com.exceljava.com4j.excel.Comment previous();


  // Properties:
}
