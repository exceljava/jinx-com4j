package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C030C-0000-0000-C000-000000000046}")
public interface _CommandBarComboBox extends com.exceljava.com4j.office.CommandBarControl {
  // Methods:
  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>java.lang.Object parameter index is set to com4j.Variant.getMissing()</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * addItem(text, com4j.Variant.getMissing());
   * </code>
   * </p>
   * @param text Mandatory java.lang.String parameter.
   */

  @DISPID(1610940416) //= 0x60050000. The runtime will prefer the VTID if present
  @VTID(83)
  @UseDefaultValues(paramIndexMapping = {0}, optParamIndex = {1}, javaType = {java.lang.Object.class}, nativeType = {NativeType.VARIANT}, variantType = {Variant.Type.VT_ERROR}, literal = {"80020004"})
  void addItem(
    java.lang.String text);

  /**
   * @param text Mandatory java.lang.String parameter.
   * @param index Optional parameter. Default value is com4j.Variant.getMissing()
   */

  @DISPID(1610940416) //= 0x60050000. The runtime will prefer the VTID if present
  @VTID(83)
  void addItem(
    java.lang.String text,
    @Optional @MarshalAs(NativeType.VARIANT) java.lang.Object index);


  /**
   */

  @DISPID(1610940417) //= 0x60050001. The runtime will prefer the VTID if present
  @VTID(84)
  void clear();


  /**
   * <p>
   * Getter method for the COM property "DropDownLines"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(1610940418) //= 0x60050002. The runtime will prefer the VTID if present
  @VTID(85)
  int getDropDownLines();


  /**
   * <p>
   * Setter method for the COM property "DropDownLines"
   * </p>
   * @param pcLines Mandatory int parameter.
   */

  @DISPID(1610940418) //= 0x60050002. The runtime will prefer the VTID if present
  @VTID(86)
  void setDropDownLines(
    int pcLines);


  /**
   * <p>
   * Getter method for the COM property "DropDownWidth"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(1610940420) //= 0x60050004. The runtime will prefer the VTID if present
  @VTID(87)
  int getDropDownWidth();


  /**
   * <p>
   * Setter method for the COM property "DropDownWidth"
   * </p>
   * @param pdx Mandatory int parameter.
   */

  @DISPID(1610940420) //= 0x60050004. The runtime will prefer the VTID if present
  @VTID(88)
  void setDropDownWidth(
    int pdx);


  /**
   * <p>
   * Getter method for the COM property "List"
   * </p>
   * @param index Mandatory int parameter.
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(1610940422) //= 0x60050006. The runtime will prefer the VTID if present
  @VTID(89)
  java.lang.String getList(
    int index);


  /**
   * <p>
   * Setter method for the COM property "List"
   * </p>
   * @param index Mandatory int parameter.
   * @param pbstrItem Mandatory java.lang.String parameter.
   */

  @DISPID(1610940422) //= 0x60050006. The runtime will prefer the VTID if present
  @VTID(90)
  void setList(
    int index,
    java.lang.String pbstrItem);


  /**
   * <p>
   * Getter method for the COM property "ListCount"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(1610940424) //= 0x60050008. The runtime will prefer the VTID if present
  @VTID(91)
  int getListCount();


  /**
   * <p>
   * Getter method for the COM property "ListHeaderCount"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(1610940425) //= 0x60050009. The runtime will prefer the VTID if present
  @VTID(92)
  int getListHeaderCount();


  /**
   * <p>
   * Setter method for the COM property "ListHeaderCount"
   * </p>
   * @param pcItems Mandatory int parameter.
   */

  @DISPID(1610940425) //= 0x60050009. The runtime will prefer the VTID if present
  @VTID(93)
  void setListHeaderCount(
    int pcItems);


  /**
   * <p>
   * Getter method for the COM property "ListIndex"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(1610940427) //= 0x6005000b. The runtime will prefer the VTID if present
  @VTID(94)
  int getListIndex();


  /**
   * <p>
   * Setter method for the COM property "ListIndex"
   * </p>
   * @param pi Mandatory int parameter.
   */

  @DISPID(1610940427) //= 0x6005000b. The runtime will prefer the VTID if present
  @VTID(95)
  void setListIndex(
    int pi);


  /**
   * @param index Mandatory int parameter.
   */

  @DISPID(1610940429) //= 0x6005000d. The runtime will prefer the VTID if present
  @VTID(96)
  void removeItem(
    int index);


  /**
   * <p>
   * Getter method for the COM property "Style"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoComboStyle
   */

  @DISPID(1610940430) //= 0x6005000e. The runtime will prefer the VTID if present
  @VTID(97)
  com.exceljava.com4j.office.MsoComboStyle getStyle();


  /**
   * <p>
   * Setter method for the COM property "Style"
   * </p>
   * @param pstyle Mandatory com.exceljava.com4j.office.MsoComboStyle parameter.
   */

  @DISPID(1610940430) //= 0x6005000e. The runtime will prefer the VTID if present
  @VTID(98)
  void setStyle(
    com.exceljava.com4j.office.MsoComboStyle pstyle);


  /**
   * <p>
   * Getter method for the COM property "Text"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(1610940432) //= 0x60050010. The runtime will prefer the VTID if present
  @VTID(99)
  java.lang.String getText();


  /**
   * <p>
   * Setter method for the COM property "Text"
   * </p>
   * @param pbstrText Mandatory java.lang.String parameter.
   */

  @DISPID(1610940432) //= 0x60050010. The runtime will prefer the VTID if present
  @VTID(100)
  void setText(
    java.lang.String pbstrText);


  /**
   * <p>
   * Getter method for the COM property "InstanceIdPtr"
   * </p>
   * @return  Returns a value of type java.lang.Object
   */

  @DISPID(1610940434) //= 0x60050012. The runtime will prefer the VTID if present
  @VTID(101)
  @ReturnValue(type=NativeType.VARIANT)
  java.lang.Object getInstanceIdPtr();


  // Properties:
}
