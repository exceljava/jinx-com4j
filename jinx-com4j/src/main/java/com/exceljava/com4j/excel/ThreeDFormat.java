package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{000C0321-0000-0000-C000-000000000046}")
public interface ThreeDFormat extends com.exceljava.com4j.office._IMsoDispObj {
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
   * @param increment Mandatory float parameter.
   */

  @DISPID(10) //= 0xa. The runtime will prefer the VTID if present
  @VTID(10)
  void incrementRotationX(
    float increment);


  /**
   * @param increment Mandatory float parameter.
   */

  @DISPID(11) //= 0xb. The runtime will prefer the VTID if present
  @VTID(11)
  void incrementRotationY(
    float increment);


  /**
   */

  @DISPID(12) //= 0xc. The runtime will prefer the VTID if present
  @VTID(12)
  void resetRotation();


  /**
   * @param presetThreeDFormat Mandatory com.exceljava.com4j.office.MsoPresetThreeDFormat parameter.
   */

  @DISPID(13) //= 0xd. The runtime will prefer the VTID if present
  @VTID(13)
  void setThreeDFormat(
    com.exceljava.com4j.office.MsoPresetThreeDFormat presetThreeDFormat);


  /**
   * @param presetExtrusionDirection Mandatory com.exceljava.com4j.office.MsoPresetExtrusionDirection parameter.
   */

  @DISPID(14) //= 0xe. The runtime will prefer the VTID if present
  @VTID(14)
  void setExtrusionDirection(
    com.exceljava.com4j.office.MsoPresetExtrusionDirection presetExtrusionDirection);


  /**
   * <p>
   * Getter method for the COM property "Depth"
   * </p>
   * @return  Returns a value of type float
   */

  @DISPID(100) //= 0x64. The runtime will prefer the VTID if present
  @VTID(15)
  float getDepth();


  /**
   * <p>
   * Setter method for the COM property "Depth"
   * </p>
   * @param depth Mandatory float parameter.
   */

  @DISPID(100) //= 0x64. The runtime will prefer the VTID if present
  @VTID(16)
  void setDepth(
    float depth);


  /**
   * <p>
   * Getter method for the COM property "ExtrusionColor"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.ColorFormat
   */

  @DISPID(101) //= 0x65. The runtime will prefer the VTID if present
  @VTID(17)
  com.exceljava.com4j.excel.ColorFormat getExtrusionColor();


  /**
   * <p>
   * Getter method for the COM property "ExtrusionColorType"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoExtrusionColorType
   */

  @DISPID(102) //= 0x66. The runtime will prefer the VTID if present
  @VTID(18)
  com.exceljava.com4j.office.MsoExtrusionColorType getExtrusionColorType();


  /**
   * <p>
   * Setter method for the COM property "ExtrusionColorType"
   * </p>
   * @param extrusionColorType Mandatory com.exceljava.com4j.office.MsoExtrusionColorType parameter.
   */

  @DISPID(102) //= 0x66. The runtime will prefer the VTID if present
  @VTID(19)
  void setExtrusionColorType(
    com.exceljava.com4j.office.MsoExtrusionColorType extrusionColorType);


  /**
   * <p>
   * Getter method for the COM property "Perspective"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTriState
   */

  @DISPID(103) //= 0x67. The runtime will prefer the VTID if present
  @VTID(20)
  com.exceljava.com4j.office.MsoTriState getPerspective();


  /**
   * <p>
   * Setter method for the COM property "Perspective"
   * </p>
   * @param perspective Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   */

  @DISPID(103) //= 0x67. The runtime will prefer the VTID if present
  @VTID(21)
  void setPerspective(
    com.exceljava.com4j.office.MsoTriState perspective);


  /**
   * <p>
   * Getter method for the COM property "PresetExtrusionDirection"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoPresetExtrusionDirection
   */

  @DISPID(104) //= 0x68. The runtime will prefer the VTID if present
  @VTID(22)
  com.exceljava.com4j.office.MsoPresetExtrusionDirection getPresetExtrusionDirection();


  /**
   * <p>
   * Getter method for the COM property "PresetLightingDirection"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoPresetLightingDirection
   */

  @DISPID(105) //= 0x69. The runtime will prefer the VTID if present
  @VTID(23)
  com.exceljava.com4j.office.MsoPresetLightingDirection getPresetLightingDirection();


  /**
   * <p>
   * Setter method for the COM property "PresetLightingDirection"
   * </p>
   * @param presetLightingDirection Mandatory com.exceljava.com4j.office.MsoPresetLightingDirection parameter.
   */

  @DISPID(105) //= 0x69. The runtime will prefer the VTID if present
  @VTID(24)
  void setPresetLightingDirection(
    com.exceljava.com4j.office.MsoPresetLightingDirection presetLightingDirection);


  /**
   * <p>
   * Getter method for the COM property "PresetLightingSoftness"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoPresetLightingSoftness
   */

  @DISPID(106) //= 0x6a. The runtime will prefer the VTID if present
  @VTID(25)
  com.exceljava.com4j.office.MsoPresetLightingSoftness getPresetLightingSoftness();


  /**
   * <p>
   * Setter method for the COM property "PresetLightingSoftness"
   * </p>
   * @param presetLightingSoftness Mandatory com.exceljava.com4j.office.MsoPresetLightingSoftness parameter.
   */

  @DISPID(106) //= 0x6a. The runtime will prefer the VTID if present
  @VTID(26)
  void setPresetLightingSoftness(
    com.exceljava.com4j.office.MsoPresetLightingSoftness presetLightingSoftness);


  /**
   * <p>
   * Getter method for the COM property "PresetMaterial"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoPresetMaterial
   */

  @DISPID(107) //= 0x6b. The runtime will prefer the VTID if present
  @VTID(27)
  com.exceljava.com4j.office.MsoPresetMaterial getPresetMaterial();


  /**
   * <p>
   * Setter method for the COM property "PresetMaterial"
   * </p>
   * @param presetMaterial Mandatory com.exceljava.com4j.office.MsoPresetMaterial parameter.
   */

  @DISPID(107) //= 0x6b. The runtime will prefer the VTID if present
  @VTID(28)
  void setPresetMaterial(
    com.exceljava.com4j.office.MsoPresetMaterial presetMaterial);


  /**
   * <p>
   * Getter method for the COM property "PresetThreeDFormat"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoPresetThreeDFormat
   */

  @DISPID(108) //= 0x6c. The runtime will prefer the VTID if present
  @VTID(29)
  com.exceljava.com4j.office.MsoPresetThreeDFormat getPresetThreeDFormat();


  /**
   * <p>
   * Getter method for the COM property "RotationX"
   * </p>
   * @return  Returns a value of type float
   */

  @DISPID(109) //= 0x6d. The runtime will prefer the VTID if present
  @VTID(30)
  float getRotationX();


  /**
   * <p>
   * Setter method for the COM property "RotationX"
   * </p>
   * @param rotationX Mandatory float parameter.
   */

  @DISPID(109) //= 0x6d. The runtime will prefer the VTID if present
  @VTID(31)
  void setRotationX(
    float rotationX);


  /**
   * <p>
   * Getter method for the COM property "RotationY"
   * </p>
   * @return  Returns a value of type float
   */

  @DISPID(110) //= 0x6e. The runtime will prefer the VTID if present
  @VTID(32)
  float getRotationY();


  /**
   * <p>
   * Setter method for the COM property "RotationY"
   * </p>
   * @param rotationY Mandatory float parameter.
   */

  @DISPID(110) //= 0x6e. The runtime will prefer the VTID if present
  @VTID(33)
  void setRotationY(
    float rotationY);


  /**
   * <p>
   * Getter method for the COM property "Visible"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTriState
   */

  @DISPID(111) //= 0x6f. The runtime will prefer the VTID if present
  @VTID(34)
  com.exceljava.com4j.office.MsoTriState getVisible();


  /**
   * <p>
   * Setter method for the COM property "Visible"
   * </p>
   * @param visible Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   */

  @DISPID(111) //= 0x6f. The runtime will prefer the VTID if present
  @VTID(35)
  void setVisible(
    com.exceljava.com4j.office.MsoTriState visible);


  /**
   * @param presetCamera Mandatory com.exceljava.com4j.office.MsoPresetCamera parameter.
   */

  @DISPID(15) //= 0xf. The runtime will prefer the VTID if present
  @VTID(36)
  void setPresetCamera(
    com.exceljava.com4j.office.MsoPresetCamera presetCamera);


  /**
   * @param increment Mandatory float parameter.
   */

  @DISPID(16) //= 0x10. The runtime will prefer the VTID if present
  @VTID(37)
  void incrementRotationZ(
    float increment);


  /**
   * @param increment Mandatory float parameter.
   */

  @DISPID(17) //= 0x11. The runtime will prefer the VTID if present
  @VTID(38)
  void incrementRotationHorizontal(
    float increment);


  /**
   * @param increment Mandatory float parameter.
   */

  @DISPID(18) //= 0x12. The runtime will prefer the VTID if present
  @VTID(39)
  void incrementRotationVertical(
    float increment);


  /**
   * <p>
   * Getter method for the COM property "PresetLighting"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoLightRigType
   */

  @DISPID(112) //= 0x70. The runtime will prefer the VTID if present
  @VTID(40)
  com.exceljava.com4j.office.MsoLightRigType getPresetLighting();


  /**
   * <p>
   * Setter method for the COM property "PresetLighting"
   * </p>
   * @param presetLightRigType Mandatory com.exceljava.com4j.office.MsoLightRigType parameter.
   */

  @DISPID(112) //= 0x70. The runtime will prefer the VTID if present
  @VTID(41)
  void setPresetLighting(
    com.exceljava.com4j.office.MsoLightRigType presetLightRigType);


  /**
   * <p>
   * Getter method for the COM property "Z"
   * </p>
   * @return  Returns a value of type float
   */

  @DISPID(113) //= 0x71. The runtime will prefer the VTID if present
  @VTID(42)
  float getZ();


  /**
   * <p>
   * Setter method for the COM property "Z"
   * </p>
   * @param z Mandatory float parameter.
   */

  @DISPID(113) //= 0x71. The runtime will prefer the VTID if present
  @VTID(43)
  void setZ(
    float z);


  /**
   * <p>
   * Getter method for the COM property "BevelTopType"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoBevelType
   */

  @DISPID(114) //= 0x72. The runtime will prefer the VTID if present
  @VTID(44)
  com.exceljava.com4j.office.MsoBevelType getBevelTopType();


  /**
   * <p>
   * Setter method for the COM property "BevelTopType"
   * </p>
   * @param bevelTopType Mandatory com.exceljava.com4j.office.MsoBevelType parameter.
   */

  @DISPID(114) //= 0x72. The runtime will prefer the VTID if present
  @VTID(45)
  void setBevelTopType(
    com.exceljava.com4j.office.MsoBevelType bevelTopType);


  /**
   * <p>
   * Getter method for the COM property "BevelTopInset"
   * </p>
   * @return  Returns a value of type float
   */

  @DISPID(115) //= 0x73. The runtime will prefer the VTID if present
  @VTID(46)
  float getBevelTopInset();


  /**
   * <p>
   * Setter method for the COM property "BevelTopInset"
   * </p>
   * @param bevelTopInset Mandatory float parameter.
   */

  @DISPID(115) //= 0x73. The runtime will prefer the VTID if present
  @VTID(47)
  void setBevelTopInset(
    float bevelTopInset);


  /**
   * <p>
   * Getter method for the COM property "BevelTopDepth"
   * </p>
   * @return  Returns a value of type float
   */

  @DISPID(116) //= 0x74. The runtime will prefer the VTID if present
  @VTID(48)
  float getBevelTopDepth();


  /**
   * <p>
   * Setter method for the COM property "BevelTopDepth"
   * </p>
   * @param bevelTopDepth Mandatory float parameter.
   */

  @DISPID(116) //= 0x74. The runtime will prefer the VTID if present
  @VTID(49)
  void setBevelTopDepth(
    float bevelTopDepth);


  /**
   * <p>
   * Getter method for the COM property "BevelBottomType"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoBevelType
   */

  @DISPID(117) //= 0x75. The runtime will prefer the VTID if present
  @VTID(50)
  com.exceljava.com4j.office.MsoBevelType getBevelBottomType();


  /**
   * <p>
   * Setter method for the COM property "BevelBottomType"
   * </p>
   * @param bevelBottomType Mandatory com.exceljava.com4j.office.MsoBevelType parameter.
   */

  @DISPID(117) //= 0x75. The runtime will prefer the VTID if present
  @VTID(51)
  void setBevelBottomType(
    com.exceljava.com4j.office.MsoBevelType bevelBottomType);


  /**
   * <p>
   * Getter method for the COM property "BevelBottomInset"
   * </p>
   * @return  Returns a value of type float
   */

  @DISPID(118) //= 0x76. The runtime will prefer the VTID if present
  @VTID(52)
  float getBevelBottomInset();


  /**
   * <p>
   * Setter method for the COM property "BevelBottomInset"
   * </p>
   * @param bevelBottomInset Mandatory float parameter.
   */

  @DISPID(118) //= 0x76. The runtime will prefer the VTID if present
  @VTID(53)
  void setBevelBottomInset(
    float bevelBottomInset);


  /**
   * <p>
   * Getter method for the COM property "BevelBottomDepth"
   * </p>
   * @return  Returns a value of type float
   */

  @DISPID(119) //= 0x77. The runtime will prefer the VTID if present
  @VTID(54)
  float getBevelBottomDepth();


  /**
   * <p>
   * Setter method for the COM property "BevelBottomDepth"
   * </p>
   * @param bevelBottomDepth Mandatory float parameter.
   */

  @DISPID(119) //= 0x77. The runtime will prefer the VTID if present
  @VTID(55)
  void setBevelBottomDepth(
    float bevelBottomDepth);


  /**
   * <p>
   * Getter method for the COM property "PresetCamera"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoPresetCamera
   */

  @DISPID(120) //= 0x78. The runtime will prefer the VTID if present
  @VTID(56)
  com.exceljava.com4j.office.MsoPresetCamera getPresetCamera();


  /**
   * <p>
   * Getter method for the COM property "RotationZ"
   * </p>
   * @return  Returns a value of type float
   */

  @DISPID(121) //= 0x79. The runtime will prefer the VTID if present
  @VTID(57)
  float getRotationZ();


  /**
   * <p>
   * Setter method for the COM property "RotationZ"
   * </p>
   * @param rotationZ Mandatory float parameter.
   */

  @DISPID(121) //= 0x79. The runtime will prefer the VTID if present
  @VTID(58)
  void setRotationZ(
    float rotationZ);


  /**
   * <p>
   * Getter method for the COM property "ContourWidth"
   * </p>
   * @return  Returns a value of type float
   */

  @DISPID(122) //= 0x7a. The runtime will prefer the VTID if present
  @VTID(59)
  float getContourWidth();


  /**
   * <p>
   * Setter method for the COM property "ContourWidth"
   * </p>
   * @param width Mandatory float parameter.
   */

  @DISPID(122) //= 0x7a. The runtime will prefer the VTID if present
  @VTID(60)
  void setContourWidth(
    float width);


  /**
   * <p>
   * Getter method for the COM property "ContourColor"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.excel.ColorFormat
   */

  @DISPID(123) //= 0x7b. The runtime will prefer the VTID if present
  @VTID(61)
  com.exceljava.com4j.excel.ColorFormat getContourColor();


  /**
   * <p>
   * Getter method for the COM property "FieldOfView"
   * </p>
   * @return  Returns a value of type float
   */

  @DISPID(124) //= 0x7c. The runtime will prefer the VTID if present
  @VTID(62)
  float getFieldOfView();


  /**
   * <p>
   * Setter method for the COM property "FieldOfView"
   * </p>
   * @param fov Mandatory float parameter.
   */

  @DISPID(124) //= 0x7c. The runtime will prefer the VTID if present
  @VTID(63)
  void setFieldOfView(
    float fov);


  /**
   * <p>
   * Getter method for the COM property "ProjectText"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTriState
   */

  @DISPID(125) //= 0x7d. The runtime will prefer the VTID if present
  @VTID(64)
  com.exceljava.com4j.office.MsoTriState getProjectText();


  /**
   * <p>
   * Setter method for the COM property "ProjectText"
   * </p>
   * @param projectText Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   */

  @DISPID(125) //= 0x7d. The runtime will prefer the VTID if present
  @VTID(65)
  void setProjectText(
    com.exceljava.com4j.office.MsoTriState projectText);


  /**
   * <p>
   * Getter method for the COM property "LightAngle"
   * </p>
   * @return  Returns a value of type float
   */

  @DISPID(126) //= 0x7e. The runtime will prefer the VTID if present
  @VTID(66)
  float getLightAngle();


  /**
   * <p>
   * Setter method for the COM property "LightAngle"
   * </p>
   * @param lightAngle Mandatory float parameter.
   */

  @DISPID(126) //= 0x7e. The runtime will prefer the VTID if present
  @VTID(67)
  void setLightAngle(
    float lightAngle);


  // Properties:
}
