package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00024466-0001-0000-C000-000000000046}")
public interface ISpeech extends Com4jObject {
  // Methods:
  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter speakAsync is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter speakXML is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter purge is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * speak(text, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param text Mandatory java.lang.String parameter.
   */

  @VTID(7)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  void speak(
    java.lang.String text);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter speakXML is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter purge is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * speak(text, speakAsync, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param text Mandatory java.lang.String parameter.
   * @param speakAsync Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @VTID(7)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  void speak(
    java.lang.String text,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object speakAsync);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter purge is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * speak(text, speakAsync, speakXML, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param text Mandatory java.lang.String parameter.
   * @param speakAsync Optional parameter. Default value is com4j.Variant.getMissing()
   * @param speakXML Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @VTID(7)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void speak(
    java.lang.String text,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object speakAsync,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object speakXML);

  /**
   * @param text Mandatory java.lang.String parameter.
   * @param speakAsync Optional parameter. Default value is com4j.Variant.getMissing()
   * @param speakXML Optional parameter. Default value is com4j.Variant.getMissing()
   * @param purge Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @VTID(7)
  void speak(
    java.lang.String text,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object speakAsync,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object speakXML,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object purge);


  /**
   * <p>
   * Getter method for the COM property "Direction"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.XlSpeakDirection
   */

  @VTID(8)
  com.exceljava.com4j.excel.XlSpeakDirection getDirection();


  /**
   * <p>
   * Setter method for the COM property "Direction"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.excel.XlSpeakDirection parameter.
   */

  @VTID(9)
  void setDirection(
    com.exceljava.com4j.excel.XlSpeakDirection rhs);


  /**
   * <p>
   * Getter method for the COM property "SpeakCellOnEnter"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(10)
  boolean getSpeakCellOnEnter();


  /**
   * <p>
   * Setter method for the COM property "SpeakCellOnEnter"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(11)
  void setSpeakCellOnEnter(
    boolean rhs);


  // Properties:
}
