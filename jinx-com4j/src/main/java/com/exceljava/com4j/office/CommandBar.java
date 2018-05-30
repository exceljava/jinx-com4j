package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C0304-0000-0000-C000-000000000046}")
public interface CommandBar extends com.exceljava.com4j.office._IMsoOleAccDispObj {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "BuiltIn"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1610874880) //= 0x60040000. The runtime will prefer the VTID if present
  @VTID(30)
  boolean getBuiltIn();


  /**
   * <p>
   * Getter method for the COM property "Context"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(1610874881) //= 0x60040001. The runtime will prefer the VTID if present
  @VTID(31)
  java.lang.String getContext();


  /**
   * <p>
   * Setter method for the COM property "Context"
   * </p>
   * @param pbstrContext Mandatory java.lang.String parameter.
   */

  @DISPID(1610874881) //= 0x60040001. The runtime will prefer the VTID if present
  @VTID(32)
  void setContext(
    java.lang.String pbstrContext);


  /**
   * <p>
   * Getter method for the COM property "Controls"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.CommandBarControls
   */

  @DISPID(1610874883) //= 0x60040003. The runtime will prefer the VTID if present
  @VTID(33)
  com.exceljava.com4j.office.CommandBarControls getControls();


  @VTID(33)
  @ReturnValue(defaultPropertyThrough={com.exceljava.com4j.office.CommandBarControls.class})
  com.exceljava.com4j.office.CommandBarControl getControls(
    @MarshalAs(NativeType.VARIANT) java.lang.Object index);

  /**
   */

  @DISPID(1610874884) //= 0x60040004. The runtime will prefer the VTID if present
  @VTID(34)
  void delete();


  /**
   * <p>
   * Getter method for the COM property "Enabled"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1610874885) //= 0x60040005. The runtime will prefer the VTID if present
  @VTID(35)
  boolean getEnabled();


  /**
   * <p>
   * Setter method for the COM property "Enabled"
   * </p>
   * @param pvarfEnabled Mandatory boolean parameter.
   */

  @DISPID(1610874885) //= 0x60040005. The runtime will prefer the VTID if present
  @VTID(36)
  void setEnabled(
    boolean pvarfEnabled);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter type is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter id is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter tag is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter visible is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter recursive is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * findControl(com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.CommandBarControl
   */

  @DISPID(1610874887) //= 0x60040007. The runtime will prefer the VTID if present
  @VTID(37)
  @UseDefaultValues(paramIndexMapping = {5}, optParamIndex = {0, 1, 2, 3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=5)
  com.exceljava.com4j.office.CommandBarControl findControl();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter id is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter tag is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter visible is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter recursive is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * findControl(type, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.office.CommandBarControl
   */

  @DISPID(1610874887) //= 0x60040007. The runtime will prefer the VTID if present
  @VTID(37)
  @UseDefaultValues(paramIndexMapping = {0, 5}, optParamIndex = {1, 2, 3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004", "80020004"})
  @ReturnValue(index=5)
  com.exceljava.com4j.office.CommandBarControl findControl(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter tag is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter visible is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter recursive is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * findControl(type, id, com4j.Variant.getMissing(), com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @param id Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.office.CommandBarControl
   */

  @DISPID(1610874887) //= 0x60040007. The runtime will prefer the VTID if present
  @VTID(37)
  @UseDefaultValues(paramIndexMapping = {0, 1, 5}, optParamIndex = {2, 3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004", "80020004"})
  @ReturnValue(index=5)
  com.exceljava.com4j.office.CommandBarControl findControl(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object id);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter visible is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter recursive is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * findControl(type, id, tag, com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @param id Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tag Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.office.CommandBarControl
   */

  @DISPID(1610874887) //= 0x60040007. The runtime will prefer the VTID if present
  @VTID(37)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 5}, optParamIndex = {3, 4}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  @ReturnValue(index=5)
  com.exceljava.com4j.office.CommandBarControl findControl(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object id,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tag);

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter recursive is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * findControl(type, id, tag, visible, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @param id Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tag Optional parameter. Default value is com4j.Variant.getMissing()
   * @param visible Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.office.CommandBarControl
   */

  @DISPID(1610874887) //= 0x60040007. The runtime will prefer the VTID if present
  @VTID(37)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3, 5}, optParamIndex = {4}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=5)
  com.exceljava.com4j.office.CommandBarControl findControl(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object id,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tag,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object visible);

  /**
   * @param type Optional parameter. Default value is com4j.Variant.getMissing()
   * @param id Optional parameter. Default value is com4j.Variant.getMissing()
   * @param tag Optional parameter. Default value is com4j.Variant.getMissing()
   * @param visible Optional parameter. Default value is com4j.Variant.getMissing()
   * @param recursive Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type com.exceljava.com4j.office.CommandBarControl
   */

  @DISPID(1610874887) //= 0x60040007. The runtime will prefer the VTID if present
  @VTID(37)
  com.exceljava.com4j.office.CommandBarControl findControl(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object type,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object id,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object tag,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object visible,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object recursive);


  /**
   * <p>
   * Getter method for the COM property "Height"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(1610874888) //= 0x60040008. The runtime will prefer the VTID if present
  @VTID(38)
  int getHeight();


  /**
   * <p>
   * Setter method for the COM property "Height"
   * </p>
   * @param pdy Mandatory int parameter.
   */

  @DISPID(1610874888) //= 0x60040008. The runtime will prefer the VTID if present
  @VTID(39)
  void setHeight(
    int pdy);


  /**
   * <p>
   * Getter method for the COM property "Index"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(1610874890) //= 0x6004000a. The runtime will prefer the VTID if present
  @VTID(40)
  int getIndex();


  /**
   * <p>
   * Getter method for the COM property "InstanceId"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(1610874891) //= 0x6004000b. The runtime will prefer the VTID if present
  @VTID(41)
  int getInstanceId();


  /**
   * <p>
   * Getter method for the COM property "Left"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(1610874892) //= 0x6004000c. The runtime will prefer the VTID if present
  @VTID(42)
  int getLeft();


  /**
   * <p>
   * Setter method for the COM property "Left"
   * </p>
   * @param pxpLeft Mandatory int parameter.
   */

  @DISPID(1610874892) //= 0x6004000c. The runtime will prefer the VTID if present
  @VTID(43)
  void setLeft(
    int pxpLeft);


  /**
   * <p>
   * Getter method for the COM property "Name"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(1610874894) //= 0x6004000e. The runtime will prefer the VTID if present
  @VTID(44)
  java.lang.String getName();


  /**
   * <p>
   * Setter method for the COM property "Name"
   * </p>
   * @param pbstrName Mandatory java.lang.String parameter.
   */

  @DISPID(1610874894) //= 0x6004000e. The runtime will prefer the VTID if present
  @VTID(45)
  void setName(
    java.lang.String pbstrName);


  /**
   * <p>
   * Getter method for the COM property "NameLocal"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(1610874896) //= 0x60040010. The runtime will prefer the VTID if present
  @VTID(46)
  java.lang.String getNameLocal();


  /**
   * <p>
   * Setter method for the COM property "NameLocal"
   * </p>
   * @param pbstrNameLocal Mandatory java.lang.String parameter.
   */

  @DISPID(1610874896) //= 0x60040010. The runtime will prefer the VTID if present
  @VTID(47)
  void setNameLocal(
    java.lang.String pbstrNameLocal);


  /**
   * <p>
   * Getter method for the COM property "Parent"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(1610874898) //= 0x60040012. The runtime will prefer the VTID if present
  @VTID(48)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getParent();


  /**
   * <p>
   * Getter method for the COM property "Position"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoBarPosition
   */

  @DISPID(1610874899) //= 0x60040013. The runtime will prefer the VTID if present
  @VTID(49)
  com.exceljava.com4j.office.MsoBarPosition getPosition();


  /**
   * <p>
   * Setter method for the COM property "Position"
   * </p>
   * @param ppos Mandatory com.exceljava.com4j.office.MsoBarPosition parameter.
   */

  @DISPID(1610874899) //= 0x60040013. The runtime will prefer the VTID if present
  @VTID(50)
  void setPosition(
    com.exceljava.com4j.office.MsoBarPosition ppos);


  /**
   * <p>
   * Getter method for the COM property "RowIndex"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(1610874901) //= 0x60040015. The runtime will prefer the VTID if present
  @VTID(51)
  int getRowIndex();


  /**
   * <p>
   * Setter method for the COM property "RowIndex"
   * </p>
   * @param piRow Mandatory int parameter.
   */

  @DISPID(1610874901) //= 0x60040015. The runtime will prefer the VTID if present
  @VTID(52)
  void setRowIndex(
    int piRow);


  /**
   * <p>
   * Getter method for the COM property "Protection"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoBarProtection
   */

  @DISPID(1610874903) //= 0x60040017. The runtime will prefer the VTID if present
  @VTID(53)
  com.exceljava.com4j.office.MsoBarProtection getProtection();


  /**
   * <p>
   * Setter method for the COM property "Protection"
   * </p>
   * @param pprot Mandatory com.exceljava.com4j.office.MsoBarProtection parameter.
   */

  @DISPID(1610874903) //= 0x60040017. The runtime will prefer the VTID if present
  @VTID(54)
  void setProtection(
    com.exceljava.com4j.office.MsoBarProtection pprot);


  /**
   */

  @DISPID(1610874905) //= 0x60040019. The runtime will prefer the VTID if present
  @VTID(55)
  void reset();


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter x is set to com4j.Variant.getMissing()</li><li>java.lang.Object parameter y is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * showPopup(com4j.Variant.getMissing(), com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(1610874906) //= 0x6004001a. The runtime will prefer the VTID if present
  @VTID(56)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0, 1}, javaType = {java.lang.Object.class, java.lang.Object.class}, nativeType = {NativeType.VARIANT, NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR, Variant.Type.VT_ERROR}, literal = {"80020004", "80020004"})
  void showPopup();

  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter y is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * showPopup(x, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param x Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1610874906) //= 0x6004001a. The runtime will prefer the VTID if present
  @VTID(56)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void showPopup(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object x);

  /**
   * @param x Optional parameter. Default value is com4j.Variant.getMissing()
   * @param y Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1610874906) //= 0x6004001a. The runtime will prefer the VTID if present
  @VTID(56)
  void showPopup(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object x,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object y);


  /**
   * <p>
   * Getter method for the COM property "Top"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(1610874907) //= 0x6004001b. The runtime will prefer the VTID if present
  @VTID(57)
  int getTop();


  /**
   * <p>
   * Setter method for the COM property "Top"
   * </p>
   * @param pypTop Mandatory int parameter.
   */

  @DISPID(1610874907) //= 0x6004001b. The runtime will prefer the VTID if present
  @VTID(58)
  void setTop(
    int pypTop);


  /**
   * <p>
   * Getter method for the COM property "Type"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoBarType
   */

  @DISPID(1610874909) //= 0x6004001d. The runtime will prefer the VTID if present
  @VTID(59)
  com.exceljava.com4j.office.MsoBarType getType();


  /**
   * <p>
   * Getter method for the COM property "Visible"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1610874910) //= 0x6004001e. The runtime will prefer the VTID if present
  @VTID(60)
  boolean getVisible();


  /**
   * <p>
   * Setter method for the COM property "Visible"
   * </p>
   * @param pvarfVisible Mandatory boolean parameter.
   */

  @DISPID(1610874910) //= 0x6004001e. The runtime will prefer the VTID if present
  @VTID(61)
  void setVisible(
    boolean pvarfVisible);


  /**
   * <p>
   * Getter method for the COM property "Width"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(1610874912) //= 0x60040020. The runtime will prefer the VTID if present
  @VTID(62)
  int getWidth();


  /**
   * <p>
   * Setter method for the COM property "Width"
   * </p>
   * @param pdx Mandatory int parameter.
   */

  @DISPID(1610874912) //= 0x60040020. The runtime will prefer the VTID if present
  @VTID(63)
  void setWidth(
    int pdx);


  /**
   * <p>
   * Getter method for the COM property "AdaptiveMenu"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1610874914) //= 0x60040022. The runtime will prefer the VTID if present
  @VTID(64)
  boolean getAdaptiveMenu();


  /**
   * <p>
   * Setter method for the COM property "AdaptiveMenu"
   * </p>
   * @param pvarfAdaptiveMenu Mandatory boolean parameter.
   */

  @DISPID(1610874914) //= 0x60040022. The runtime will prefer the VTID if present
  @VTID(65)
  void setAdaptiveMenu(
    boolean pvarfAdaptiveMenu);


  /**
   * <p>
   * Getter method for the COM property "Id"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(1610874916) //= 0x60040024. The runtime will prefer the VTID if present
  @VTID(66)
  int getId();


  /**
   * <p>
   * Getter method for the COM property "InstanceIdPtr"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(1610874917) //= 0x60040025. The runtime will prefer the VTID if present
  @VTID(67)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getInstanceIdPtr();


  // Properties:
}
