package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00024435-0001-0000-C000-000000000046}")
public interface IChartFillFormat extends Com4jObject {
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
   * @param style Mandatory com.exceljava.com4j.office.MsoGradientStyle parameter.
   * @param variant Mandatory int parameter.
   * @param degree Mandatory float parameter.
   */

  @VTID(10)
  void oneColorGradient(
    com.exceljava.com4j.office.MsoGradientStyle style,
    int variant,
    float degree);


  /**
   * @param style Mandatory com.exceljava.com4j.office.MsoGradientStyle parameter.
   * @param variant Mandatory int parameter.
   */

  @VTID(11)
  void twoColorGradient(
    com.exceljava.com4j.office.MsoGradientStyle style,
    int variant);


  /**
   * @param presetTexture Mandatory com.exceljava.com4j.office.MsoPresetTexture parameter.
   */

  @VTID(12)
  void presetTextured(
    com.exceljava.com4j.office.MsoPresetTexture presetTexture);


  /**
   */

  @VTID(13)
  void solid();


  /**
   * @param pattern Mandatory com.exceljava.com4j.office.MsoPatternType parameter.
   */

  @VTID(14)
  void patterned(
    com.exceljava.com4j.office.MsoPatternType pattern);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter pictureFile is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter pictureFormat is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter pictureStackUnit is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter picturePlacement is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * userPicture(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @VTID(15)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1, 2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  void userPicture();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter pictureFormat is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter pictureStackUnit is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter picturePlacement is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * userPicture(pictureFile, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param pictureFile Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @VTID(15)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1, 2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  void userPicture(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object pictureFile);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter pictureStackUnit is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter picturePlacement is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * userPicture(pictureFile, pictureFormat, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param pictureFile Optional parameter. Default value is com4j.Variant.getMissing()
   * @param pictureFormat Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @VTID(15)
  @UseDefaultValues(paramIndexMapping = {0, 1}, optParamIndex = {2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  void userPicture(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object pictureFile,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object pictureFormat);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter picturePlacement is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * userPicture(pictureFile, pictureFormat, pictureStackUnit, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param pictureFile Optional parameter. Default value is com4j.Variant.getMissing()
   * @param pictureFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param pictureStackUnit Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @VTID(15)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2}, optParamIndex = {3}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void userPicture(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object pictureFile,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object pictureFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object pictureStackUnit);

  /**
   * @param pictureFile Optional parameter. Default value is com4j.Variant.getMissing()
   * @param pictureFormat Optional parameter. Default value is com4j.Variant.getMissing()
   * @param pictureStackUnit Optional parameter. Default value is com4j.Variant.getMissing()
   * @param picturePlacement Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @VTID(15)
  void userPicture(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object pictureFile,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object pictureFormat,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object pictureStackUnit,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object picturePlacement);


  /**
   * @param textureFile Mandatory java.lang.String parameter.
   */

  @VTID(16)
  void userTextured(
    java.lang.String textureFile);


  /**
   * @param style Mandatory com.exceljava.com4j.office.MsoGradientStyle parameter.
   * @param variant Mandatory int parameter.
   * @param presetGradientType Mandatory com.exceljava.com4j.office.MsoPresetGradientType parameter.
   */

  @VTID(17)
  void presetGradient(
    com.exceljava.com4j.office.MsoGradientStyle style,
    int variant,
    com.exceljava.com4j.office.MsoPresetGradientType presetGradientType);


  /**
   * <p>
   * Getter method for the COM property "BackColor"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.ChartColorFormat
   */

  @VTID(18)
  com.exceljava.com4j.excel.ChartColorFormat getBackColor();


  /**
   * <p>
   * Getter method for the COM property "ForeColor"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.ChartColorFormat
   */

  @VTID(19)
  com.exceljava.com4j.excel.ChartColorFormat getForeColor();


  /**
   * <p>
   * Getter method for the COM property "GradientColorType"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoGradientColorType
   */

  @VTID(20)
  com.exceljava.com4j.office.MsoGradientColorType getGradientColorType();


  /**
   * <p>
   * Getter method for the COM property "GradientDegree"
   * </p>
   * @return  Returns a value of type float
   */

  @VTID(21)
  float getGradientDegree();


  /**
   * <p>
   * Getter method for the COM property "GradientStyle"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoGradientStyle
   */

  @VTID(22)
  com.exceljava.com4j.office.MsoGradientStyle getGradientStyle();


  /**
   * <p>
   * Getter method for the COM property "GradientVariant"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(23)
  int getGradientVariant();


  /**
   * <p>
   * Getter method for the COM property "Pattern"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoPatternType
   */

  @VTID(24)
  com.exceljava.com4j.office.MsoPatternType getPattern();


  /**
   * <p>
   * Getter method for the COM property "PresetGradientType"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoPresetGradientType
   */

  @VTID(25)
  com.exceljava.com4j.office.MsoPresetGradientType getPresetGradientType();


  /**
   * <p>
   * Getter method for the COM property "PresetTexture"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoPresetTexture
   */

  @VTID(26)
  com.exceljava.com4j.office.MsoPresetTexture getPresetTexture();


  /**
   * <p>
   * Getter method for the COM property "TextureName"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(27)
  java.lang.String getTextureName();


  /**
   * <p>
   * Getter method for the COM property "TextureType"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTextureType
   */

  @VTID(28)
  com.exceljava.com4j.office.MsoTextureType getTextureType();


  /**
   * <p>
   * Getter method for the COM property "Type"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoFillType
   */

  @VTID(29)
  com.exceljava.com4j.office.MsoFillType getType();


  /**
   * <p>
   * Getter method for the COM property "Visible"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTriState
   */

  @VTID(30)
  com.exceljava.com4j.office.MsoTriState getVisible();


  /**
   * <p>
   * Setter method for the COM property "Visible"
   * </p>
   * @param rhs Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   */

  @VTID(31)
  void setVisible(
    com.exceljava.com4j.office.MsoTriState rhs);


  // Properties:
}
