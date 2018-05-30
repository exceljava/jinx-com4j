package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C030E-0000-0000-C000-000000000046}")
public interface _CommandBarButton extends com.exceljava.com4j.office.CommandBarControl {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "BuiltInFace"
   * </p>
   * @return  Returns a value of type boolean
   */

  @DISPID(1610940416) //= 0x60050000. The runtime will prefer the VTID if present
  @VTID(83)
  boolean getBuiltInFace();


  /**
   * <p>
   * Setter method for the COM property "BuiltInFace"
   * </p>
   * @param pvarfBuiltIn Mandatory boolean parameter.
   */

  @DISPID(1610940416) //= 0x60050000. The runtime will prefer the VTID if present
  @VTID(84)
  void setBuiltInFace(
    boolean pvarfBuiltIn);


  /**
   */

  @DISPID(1610940418) //= 0x60050002. The runtime will prefer the VTID if present
  @VTID(85)
  void copyFace();


  /**
   * <p>
   * Getter method for the COM property "FaceId"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(1610940419) //= 0x60050003. The runtime will prefer the VTID if present
  @VTID(86)
  int getFaceId();


  /**
   * <p>
   * Setter method for the COM property "FaceId"
   * </p>
   * @param pid Mandatory int parameter.
   */

  @DISPID(1610940419) //= 0x60050003. The runtime will prefer the VTID if present
  @VTID(87)
  void setFaceId(
    int pid);


  /**
   */

  @DISPID(1610940421) //= 0x60050005. The runtime will prefer the VTID if present
  @VTID(88)
  void pasteFace();


  /**
   * <p>
   * Getter method for the COM property "ShortcutText"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(1610940422) //= 0x60050006. The runtime will prefer the VTID if present
  @VTID(89)
  java.lang.String getShortcutText();


  /**
   * <p>
   * Setter method for the COM property "ShortcutText"
   * </p>
   * @param pbstrText Mandatory java.lang.String parameter.
   */

  @DISPID(1610940422) //= 0x60050006. The runtime will prefer the VTID if present
  @VTID(90)
  void setShortcutText(
    java.lang.String pbstrText);


  /**
   * <p>
   * Getter method for the COM property "State"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoButtonState
   */

  @DISPID(1610940424) //= 0x60050008. The runtime will prefer the VTID if present
  @VTID(91)
  com.exceljava.com4j.office.MsoButtonState getState();


  /**
   * <p>
   * Setter method for the COM property "State"
   * </p>
   * @param pstate Mandatory com.exceljava.com4j.office.MsoButtonState parameter.
   */

  @DISPID(1610940424) //= 0x60050008. The runtime will prefer the VTID if present
  @VTID(92)
  void setState(
    com.exceljava.com4j.office.MsoButtonState pstate);


  /**
   * <p>
   * Getter method for the COM property "Style"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoButtonStyle
   */

  @DISPID(1610940426) //= 0x6005000a. The runtime will prefer the VTID if present
  @VTID(93)
  com.exceljava.com4j.office.MsoButtonStyle getStyle();


  /**
   * <p>
   * Setter method for the COM property "Style"
   * </p>
   * @param pstyle Mandatory com.exceljava.com4j.office.MsoButtonStyle parameter.
   */

  @DISPID(1610940426) //= 0x6005000a. The runtime will prefer the VTID if present
  @VTID(94)
  void setStyle(
    com.exceljava.com4j.office.MsoButtonStyle pstyle);


  /**
   * <p>
   * Getter method for the COM property "HyperlinkType"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoCommandBarButtonHyperlinkType
   */

  @DISPID(1610940428) //= 0x6005000c. The runtime will prefer the VTID if present
  @VTID(95)
  com.exceljava.com4j.office.MsoCommandBarButtonHyperlinkType getHyperlinkType();


  /**
   * <p>
   * Setter method for the COM property "HyperlinkType"
   * </p>
   * @param phlType Mandatory com.exceljava.com4j.office.MsoCommandBarButtonHyperlinkType parameter.
   */

  @DISPID(1610940428) //= 0x6005000c. The runtime will prefer the VTID if present
  @VTID(96)
  void setHyperlinkType(
    com.exceljava.com4j.office.MsoCommandBarButtonHyperlinkType phlType);


  /**
   * <p>
   * Getter method for the COM property "Picture"
   * </p>
   * @return  Returns a value of type Holder<com4j.stdole.IPictureDisp>
   */

  @DISPID(1610940430) //= 0x6005000e. The runtime will prefer the VTID if present
  @VTID(97)
  Holder<com4j.stdole.IPictureDisp> getPicture();


  /**
   * <p>
   * Setter method for the COM property "Picture"
   * </p>
   * @param ppdispPicture Mandatory Holder<com4j.stdole.IPictureDisp> parameter.
   */

  @DISPID(1610940430) //= 0x6005000e. The runtime will prefer the VTID if present
  @VTID(98)
  void setPicture(
    Holder<com4j.stdole.IPictureDisp> ppdispPicture);


  /**
   * <p>
   * Getter method for the COM property "Mask"
   * </p>
   * @return  Returns a value of type Holder<com4j.stdole.IPictureDisp>
   */

  @DISPID(1610940432) //= 0x60050010. The runtime will prefer the VTID if present
  @VTID(99)
  Holder<com4j.stdole.IPictureDisp> getMask();


  /**
   * <p>
   * Setter method for the COM property "Mask"
   * </p>
   * @param ppipictdispMask Mandatory Holder<com4j.stdole.IPictureDisp> parameter.
   */

  @DISPID(1610940432) //= 0x60050010. The runtime will prefer the VTID if present
  @VTID(100)
  void setMask(
    Holder<com4j.stdole.IPictureDisp> ppipictdispMask);


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
