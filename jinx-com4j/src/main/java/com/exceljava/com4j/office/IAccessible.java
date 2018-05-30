package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{618736E0-3C3D-11CF-810C-00AA00389B71}")
public interface IAccessible extends Com4jObject {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "accParent"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(-5000) //= 0xffffec78. The runtime will prefer the VTID if present
  @VTID(7)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getAccParent();


  /**
   * <p>
   * Getter method for the COM property "accChildCount"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(-5001) //= 0xffffec77. The runtime will prefer the VTID if present
  @VTID(8)
  int getAccChildCount();


  /**
   * <p>
   * Getter method for the COM property "accChild"
   * </p>
   * @param varChild Mandatory java.lang.Object parameter.
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(-5002) //= 0xffffec76. The runtime will prefer the VTID if present
  @VTID(9)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getAccChild(
    @MarshalAs(NativeType.VARIANT) java.lang.Object varChild);


  /**
   * <p>
   * Getter method for the COM property "accName"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter varChild is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getAccName(com4j.Variant.getMissing());
   * </code>
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(-5003) //= 0xffffec75. The runtime will prefer the VTID if present
  @VTID(10)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=1)
  java.lang.String getAccName();

  /**
   * <p>
   * Getter method for the COM property "accName"
   * </p>
   * @param varChild Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(-5003) //= 0xffffec75. The runtime will prefer the VTID if present
  @VTID(10)
  java.lang.String getAccName(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varChild);


  /**
   * <p>
   * Getter method for the COM property "accValue"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter varChild is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getAccValue(com4j.Variant.getMissing());
   * </code>
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(-5004) //= 0xffffec74. The runtime will prefer the VTID if present
  @VTID(11)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=1)
  java.lang.String getAccValue();

  /**
   * <p>
   * Getter method for the COM property "accValue"
   * </p>
   * @param varChild Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(-5004) //= 0xffffec74. The runtime will prefer the VTID if present
  @VTID(11)
  java.lang.String getAccValue(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varChild);


  /**
   * <p>
   * Getter method for the COM property "accDescription"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter varChild is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getAccDescription(com4j.Variant.getMissing());
   * </code>
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(-5005) //= 0xffffec73. The runtime will prefer the VTID if present
  @VTID(12)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=1)
  java.lang.String getAccDescription();

  /**
   * <p>
   * Getter method for the COM property "accDescription"
   * </p>
   * @param varChild Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(-5005) //= 0xffffec73. The runtime will prefer the VTID if present
  @VTID(12)
  java.lang.String getAccDescription(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varChild);


  /**
   * <p>
   * Getter method for the COM property "accRole"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter varChild is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getAccRole(com4j.Variant.getMissing());
   * </code>
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(-5006) //= 0xffffec72. The runtime will prefer the VTID if present
  @VTID(13)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=1)
  java.lang.Object getAccRole();

  /**
   * <p>
   * Getter method for the COM property "accRole"
   * </p>
   * @param varChild Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(-5006) //= 0xffffec72. The runtime will prefer the VTID if present
  @VTID(13)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getAccRole(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varChild);


  /**
   * <p>
   * Getter method for the COM property "accState"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter varChild is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getAccState(com4j.Variant.getMissing());
   * </code>
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(-5007) //= 0xffffec71. The runtime will prefer the VTID if present
  @VTID(14)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=1)
  java.lang.Object getAccState();

  /**
   * <p>
   * Getter method for the COM property "accState"
   * </p>
   * @param varChild Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(-5007) //= 0xffffec71. The runtime will prefer the VTID if present
  @VTID(14)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getAccState(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varChild);


  /**
   * <p>
   * Getter method for the COM property "accHelp"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter varChild is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getAccHelp(com4j.Variant.getMissing());
   * </code>
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(-5008) //= 0xffffec70. The runtime will prefer the VTID if present
  @VTID(15)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=1)
  java.lang.String getAccHelp();

  /**
   * <p>
   * Getter method for the COM property "accHelp"
   * </p>
   * @param varChild Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(-5008) //= 0xffffec70. The runtime will prefer the VTID if present
  @VTID(15)
  java.lang.String getAccHelp(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varChild);


  /**
   * <p>
   * Getter method for the COM property "accHelpTopic"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter varChild is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getAccHelpTopic(pszHelpFile, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param pszHelpFile Mandatory Holder<java.lang.String> parameter.
   * @return  Returns a value of type int
   */

  @DISPID(-5009) //= 0xffffec6f. The runtime will prefer the VTID if present
  @VTID(16)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=2)
  int getAccHelpTopic(
    Holder<java.lang.String> pszHelpFile);

  /**
   * <p>
   * Getter method for the COM property "accHelpTopic"
   * </p>
   * @param pszHelpFile Mandatory Holder<java.lang.String> parameter.
   * @param varChild Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type int
   */

  @DISPID(-5009) //= 0xffffec6f. The runtime will prefer the VTID if present
  @VTID(16)
  int getAccHelpTopic(
    Holder<java.lang.String> pszHelpFile,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varChild);


  /**
   * <p>
   * Getter method for the COM property "accKeyboardShortcut"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter varChild is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getAccKeyboardShortcut(com4j.Variant.getMissing());
   * </code>
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(-5010) //= 0xffffec6e. The runtime will prefer the VTID if present
  @VTID(17)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=1)
  java.lang.String getAccKeyboardShortcut();

  /**
   * <p>
   * Getter method for the COM property "accKeyboardShortcut"
   * </p>
   * @param varChild Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(-5010) //= 0xffffec6e. The runtime will prefer the VTID if present
  @VTID(17)
  java.lang.String getAccKeyboardShortcut(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varChild);


  /**
   * <p>
   * Getter method for the COM property "accFocus"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(-5011) //= 0xffffec6d. The runtime will prefer the VTID if present
  @VTID(18)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getAccFocus();


  /**
   * <p>
   * Getter method for the COM property "accSelection"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(-5012) //= 0xffffec6c. The runtime will prefer the VTID if present
  @VTID(19)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getAccSelection();


  /**
   * <p>
   * Getter method for the COM property "accDefaultAction"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter varChild is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * getAccDefaultAction(com4j.Variant.getMissing());
   * </code>
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(-5013) //= 0xffffec6b. The runtime will prefer the VTID if present
  @VTID(20)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(index=1)
  java.lang.String getAccDefaultAction();

  /**
   * <p>
   * Getter method for the COM property "accDefaultAction"
   * </p>
   * @param varChild Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(-5013) //= 0xffffec6b. The runtime will prefer the VTID if present
  @VTID(20)
  java.lang.String getAccDefaultAction(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varChild);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter varChild is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * accSelect(flagsSelect, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param flagsSelect Mandatory int parameter.
   */

  @DISPID(-5014) //= 0xffffec6a. The runtime will prefer the VTID if present
  @VTID(21)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void accSelect(
    int flagsSelect);

  /**
   * @param flagsSelect Mandatory int parameter.
   * @param varChild Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(-5014) //= 0xffffec6a. The runtime will prefer the VTID if present
  @VTID(21)
  void accSelect(
    int flagsSelect,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varChild);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter varChild is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * accLocation(pxLeft, pyTop, pcxWidth, pcyHeight, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param pxLeft Mandatory Holder<Integer> parameter.
   * @param pyTop Mandatory Holder<Integer> parameter.
   * @param pcxWidth Mandatory Holder<Integer> parameter.
   * @param pcyHeight Mandatory Holder<Integer> parameter.
   */

  @DISPID(-5015) //= 0xffffec69. The runtime will prefer the VTID if present
  @VTID(22)
  @UseDefaultValues(paramIndexMapping = {0, 1, 2, 3}, optParamIndex = {4}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void accLocation(
    Holder<Integer> pxLeft,
    Holder<Integer> pyTop,
    Holder<Integer> pcxWidth,
    Holder<Integer> pcyHeight);

  /**
   * @param pxLeft Mandatory Holder<Integer> parameter.
   * @param pyTop Mandatory Holder<Integer> parameter.
   * @param pcxWidth Mandatory Holder<Integer> parameter.
   * @param pcyHeight Mandatory Holder<Integer> parameter.
   * @param varChild Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(-5015) //= 0xffffec69. The runtime will prefer the VTID if present
  @VTID(22)
  void accLocation(
    Holder<Integer> pxLeft,
    Holder<Integer> pyTop,
    Holder<Integer> pcxWidth,
    Holder<Integer> pcyHeight,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varChild);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter varStart is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * accNavigate(navDir, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param navDir Mandatory int parameter.
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(-5016) //= 0xffffec68. The runtime will prefer the VTID if present
  @VTID(23)
  @UseDefaultValues(paramIndexMapping = {0, 2}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  @ReturnValue(type=NativeType.VARIANT,index=2)
  java.lang.Object accNavigate(
    int navDir);

  /**
   * @param navDir Mandatory int parameter.
   * @param varStart Optional parameter. Default value is com4j.Variant.getMissing()
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(-5016) //= 0xffffec68. The runtime will prefer the VTID if present
  @VTID(23)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object accNavigate(
    int navDir,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varStart);


  /**
   * @param xLeft Mandatory int parameter.
   * @param yTop Mandatory int parameter.
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(-5017) //= 0xffffec67. The runtime will prefer the VTID if present
  @VTID(24)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object accHitTest(
    int xLeft,
    int yTop);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter varChild is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * accDoDefaultAction(com4j.Variant.getMissing());
   * </code>
   * </p>
   */

  @DISPID(-5018) //= 0xffffec66. The runtime will prefer the VTID if present
  @VTID(25)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void accDoDefaultAction();

  /**
   * @param varChild Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(-5018) //= 0xffffec66. The runtime will prefer the VTID if present
  @VTID(25)
  void accDoDefaultAction(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varChild);


  /**
   * <p>
   * Setter method for the COM property "accName"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter varChild is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setAccName(com4j.Variant.getMissing(), pszName);
   * </code>
   * </p>
   * @param pszName Mandatory java.lang.String parameter.
   */

  @DISPID(-5003) //= 0xffffec75. The runtime will prefer the VTID if present
  @VTID(26)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void setAccName(
    java.lang.String pszName);

  /**
   * <p>
   * Setter method for the COM property "accName"
   * </p>
   * @param varChild Optional parameter. Default value is com4j.Variant.getMissing()
   * @param pszName Mandatory java.lang.String parameter.
   */

  @DISPID(-5003) //= 0xffffec75. The runtime will prefer the VTID if present
  @VTID(26)
  void setAccName(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varChild,
    java.lang.String pszName);


  /**
   * <p>
   * Setter method for the COM property "accValue"
   * </p>
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter varChild is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * setAccValue(com4j.Variant.getMissing(), pszValue);
   * </code>
   * </p>
   * @param pszValue Mandatory java.lang.String parameter.
   */

  @DISPID(-5004) //= 0xffffec74. The runtime will prefer the VTID if present
  @VTID(27)
  @UseDefaultValues(paramIndexMapping = {1}, optParamIndex = {0}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void setAccValue(
    java.lang.String pszValue);

  /**
   * <p>
   * Setter method for the COM property "accValue"
   * </p>
   * @param varChild Optional parameter. Default value is com4j.Variant.getMissing()
   * @param pszValue Mandatory java.lang.String parameter.
   */

  @DISPID(-5004) //= 0xffffec74. The runtime will prefer the VTID if present
  @VTID(27)
  void setAccValue(
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object varChild,
    java.lang.String pszValue);


  // Properties:
}
