package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{00020880-0001-0000-C000-000000000046}")
public interface ICheckBoxes extends Com4jObject,Iterable<Com4jObject> {
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
   */

  @VTID(10)
  void _Dummy3();


  /**
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(11)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object bringToFront();


  /**
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(12)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object copy();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlPictureAppearance parameter appearance is set to 2</li><li>com.exceljava.com4j.excel.XlCopyPictureFormat parameter format is set to -4147</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * copyPicture(2, -4147);
   * </code>
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(13)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {com.exceljava.com4j.excel.XlPictureAppearance.class, com.exceljava.com4j.excel.XlCopyPictureFormat.class}, nativeType = {NativeType.Int32, NativeType.Int32}, variantType = {Variant.Type.VT_I4, Variant.Type.VT_I4}, literal = {"2", "-4147"})
  @ReturnValue(type=NativeType.VARIANT,index=2)
  java.lang.Object copyPicture();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>com.exceljava.com4j.excel.XlCopyPictureFormat parameter format is set to -4147</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * copyPicture(appearance, -4147);
   * </code>
   * </p>
   * @param appearance Optional parameter. Default value is 2
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(13)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {com.exceljava.com4j.excel.XlCopyPictureFormat.class}, nativeType = {NativeType.Int32}, variantType = {Variant.Type.VT_I4}, literal = {"-4147"})
  @ReturnValue(type=NativeType.VARIANT,index=2)
  java.lang.Object copyPicture(
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlPictureAppearance appearance);

  /**
   * @param appearance Optional parameter. Default value is 2
   * @param format Optional parameter. Default value is -4147
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(13)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object copyPicture(
    @Optional @DefaultValue("2") com.exceljava.com4j.excel.XlPictureAppearance appearance,
    @Optional @DefaultValue("-4147") com.exceljava.com4j.excel.XlCopyPictureFormat format);


  /**
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(14)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object cut();


  /**
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(15)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object delete();


  /**
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(16)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject duplicate();


  /**
   * <p>
   * Getter method for the COM property "Enabled"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(17)
  boolean getEnabled();


  /**
   * <p>
   * Setter method for the COM property "Enabled"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(18)
  void setEnabled(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "Height"
   * </p>
   * @return  Returns a value of type double
   */

  @VTID(19)
  double getHeight();


  /**
   * <p>
   * Setter method for the COM property "Height"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @VTID(20)
  void setHeight(
    double rhs);


  /**
   */

  @VTID(21)
  void _Dummy12();


  /**
   * <p>
   * Getter method for the COM property "Left"
   * </p>
   * @return  Returns a value of type double
   */

  @VTID(22)
  double getLeft();


  /**
   * <p>
   * Setter method for the COM property "Left"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @VTID(23)
  void setLeft(
    double rhs);


  /**
   * <p>
   * Getter method for the COM property "Locked"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(24)
  boolean getLocked();


  /**
   * <p>
   * Setter method for the COM property "Locked"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(25)
  void setLocked(
    boolean rhs);


  /**
   */

  @VTID(26)
  void _Dummy15();


  /**
   * <p>
   * Getter method for the COM property "OnAction"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(27)
  java.lang.String getOnAction();


  /**
   * <p>
   * Setter method for the COM property "OnAction"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(28)
  void setOnAction(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "Placement"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(29)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getPlacement();


  /**
   * <p>
   * Setter method for the COM property "Placement"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(30)
  void setPlacement(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "PrintObject"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(31)
  boolean getPrintObject();


  /**
   * <p>
   * Setter method for the COM property "PrintObject"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(32)
  void setPrintObject(
    boolean rhs);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter replace is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * select(com4j.Variant.getMissing());
   * </code>
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(33)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=1)
  java.lang.Object select();

  /**
   * @param replace Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(33)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object select(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object replace);


  /**
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(34)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object sendToBack();


  /**
   * <p>
   * Getter method for the COM property "Top"
   * </p>
   * @return  Returns a value of type double
   */

  @VTID(35)
  double getTop();


  /**
   * <p>
   * Setter method for the COM property "Top"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @VTID(36)
  void setTop(
    double rhs);


  /**
   */

  @VTID(37)
  void _Dummy22();


  /**
   * <p>
   * Getter method for the COM property "Visible"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(38)
  boolean getVisible();


  /**
   * <p>
   * Setter method for the COM property "Visible"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(39)
  void setVisible(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "Width"
   * </p>
   * @return  Returns a value of type double
   */

  @VTID(40)
  double getWidth();


  /**
   * <p>
   * Setter method for the COM property "Width"
   * </p>
   * @param rhs Mandatory double parameter.
   */

  @VTID(41)
  void setWidth(
    double rhs);


  /**
   * <p>
   * Getter method for the COM property "ZOrder"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(42)
  int getZOrder();


  /**
   * <p>
   * Getter method for the COM property "ShapeRange"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.ShapeRange
   */

  @VTID(43)
  com.exceljava.com4j.excel.ShapeRange getShapeRange();


  /**
   * <p>
   * Getter method for the COM property "Caption"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(44)
  java.lang.String getCaption();


  /**
   * <p>
   * Setter method for the COM property "Caption"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(45)
  void setCaption(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "Characters"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter start is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter length is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getCharacters(com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Characters
   */

  @VTID(46)
  @UseDefaultValues(paramIndexMapping = {2}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=2)
  com.exceljava.com4j.excel.Characters getCharacters();

  /**
   * <p>
   * Getter method for the COM property "Characters"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter length is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getCharacters(start, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param start Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.Characters
   */

  @VTID(46)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=2)
  com.exceljava.com4j.excel.Characters getCharacters(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object start);

  /**
   * <p>
   * Getter method for the COM property "Characters"
   * </p>
   * @param start Optional parameter. Default value is com4j.Variant.getMissing()
   * @param length Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.excel.Characters
   */

  @VTID(46)
  com.exceljava.com4j.excel.Characters getCharacters(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object start,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object length);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter customDictionary is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter ignoreUppercase is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter alwaysSuggest is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter spellLang is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * checkSpelling(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(47)
  @UseDefaultValues(paramIndexMapping = {4}, optParamIndex = {0, 1, 2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=4)
  java.lang.Object checkSpelling();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter ignoreUppercase is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter alwaysSuggest is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter spellLang is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * checkSpelling(customDictionary, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param customDictionary Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(47)
  @UseDefaultValues(paramIndexMapping = {0, 4}, optParamIndex = {1, 2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=4)
  java.lang.Object checkSpelling(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object customDictionary);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter alwaysSuggest is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter spellLang is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * checkSpelling(customDictionary, ignoreUppercase, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param customDictionary Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignoreUppercase Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(47)
  @UseDefaultValues(paramIndexMapping = {0, 1, 4}, optParamIndex = {2, 3}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=4)
  java.lang.Object checkSpelling(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object customDictionary,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignoreUppercase);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter spellLang is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * checkSpelling(customDictionary, ignoreUppercase, alwaysSuggest, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param customDictionary Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignoreUppercase Optional parameter. Default value is com4j.Variant.getMissing()
   * @param alwaysSuggest Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(47)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 4}, optParamIndex = {3}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=4)
  java.lang.Object checkSpelling(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object customDictionary,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignoreUppercase,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object alwaysSuggest);

  /**
   * @param customDictionary Optional parameter. Default value is com4j.Variant.getMissing()
   * @param ignoreUppercase Optional parameter. Default value is com4j.Variant.getMissing()
   * @param alwaysSuggest Optional parameter. Default value is com4j.Variant.getMissing()
   * @param spellLang Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(47)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object checkSpelling(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object customDictionary,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object ignoreUppercase,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object alwaysSuggest,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object spellLang);


  /**
   * <p>
   * Getter method for the COM property "LockedText"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(48)
  boolean getLockedText();


  /**
   * <p>
   * Setter method for the COM property "LockedText"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(49)
  void setLockedText(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "Text"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(50)
  java.lang.String getText();


  /**
   * <p>
   * Setter method for the COM property "Text"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(51)
  void setText(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "Accelerator"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(52)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getAccelerator();


  /**
   * <p>
   * Setter method for the COM property "Accelerator"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(53)
  void setAccelerator(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "Border"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Border
   */

  @VTID(54)
  com.exceljava.com4j.excel.Border getBorder();


  /**
   * <p>
   * Getter method for the COM property "_Default"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(55)
  @DefaultMethod
  int get_Default();


  /**
   * <p>
   * Setter method for the COM property "_Default"
   * </p>
   * @param rhs Mandatory int parameter.
   */

  @VTID(56)
  @DefaultMethod
  void set_Default(
    int rhs);


  /**
   * <p>
   * Getter method for the COM property "Display3DShading"
   * </p>
   * @return  Returns a value of type boolean
   */

  @VTID(57)
  boolean getDisplay3DShading();


  /**
   * <p>
   * Setter method for the COM property "Display3DShading"
   * </p>
   * @param rhs Mandatory boolean parameter.
   */

  @VTID(58)
  void setDisplay3DShading(
    boolean rhs);


  /**
   * <p>
   * Getter method for the COM property "Interior"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.Interior
   */

  @VTID(59)
  com.exceljava.com4j.excel.Interior getInterior();


  /**
   * <p>
   * Getter method for the COM property "LinkedCell"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @VTID(60)
  java.lang.String getLinkedCell();


  /**
   * <p>
   * Setter method for the COM property "LinkedCell"
   * </p>
   * @param rhs Mandatory java.lang.String parameter.
   */

  @VTID(61)
  void setLinkedCell(
    java.lang.String rhs);


  /**
   * <p>
   * Getter method for the COM property "PhoneticAccelerator"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(62)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getPhoneticAccelerator();


  /**
   * <p>
   * Setter method for the COM property "PhoneticAccelerator"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(63)
  void setPhoneticAccelerator(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * <p>
   * Getter method for the COM property "Value"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @VTID(64)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getValue();


  /**
   * <p>
   * Setter method for the COM property "Value"
   * </p>
   * @param rhs Mandatory java.lang.Object parameter.
   */

  @VTID(65)
  void setValue(
    @MarshalAs(NativeType.VARIANT) java.lang.Object rhs);


  /**
   * @param left Mandatory double parameter.
   * @param top Mandatory double parameter.
   * @param width Mandatory double parameter.
   * @param height Mandatory double parameter.
   * @return  Returns a value of type com.exceljava.com4j.excel.CheckBox
   */

  @VTID(66)
  com.exceljava.com4j.excel.CheckBox add(
    double left,
    double top,
    double width,
    double height);


  /**
   * <p>
   * Getter method for the COM property "Count"
   * </p>
   * @return  Returns a value of type int
   */

  @VTID(67)
  int getCount();


  /**
   * @return  Returns a value of type com.exceljava.com4j.excel.GroupObject
   */

  @VTID(68)
  com.exceljava.com4j.excel.GroupObject group();


  /**
   * @param index Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com4j.Com4jObject
   */

  @VTID(69)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject item(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   */

  @VTID(70)
  java.util.Iterator<Com4jObject> iterator();

  // Properties:
}
