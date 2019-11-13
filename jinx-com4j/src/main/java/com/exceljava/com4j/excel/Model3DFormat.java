package com.exceljava.com4j.excel  ;

import com4j.*;

@IID("{000C03D8-0000-0000-C000-000000000046}")
public interface Model3DFormat extends com.exceljava.com4j.office._IMsoDispObj {
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
   * <p>
   * Getter method for the COM property "AutoFit"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTriState
   */

  @DISPID(100) //= 0x64. The runtime will prefer the VTID if present
  @VTID(10)
  com.exceljava.com4j.office.MsoTriState getAutoFit();


  /**
   * <p>
   * Setter method for the COM property "AutoFit"
   * </p>
   * @param autoFit Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   */

  @DISPID(100) //= 0x64. The runtime will prefer the VTID if present
  @VTID(11)
  void setAutoFit(
    com.exceljava.com4j.office.MsoTriState autoFit);


  /**
   * <p>
   * Getter method for the COM property "RotationX"
   * </p>
   * @return  Returns a value of type float
   */

  @DISPID(101) //= 0x65. The runtime will prefer the VTID if present
  @VTID(12)
  float getRotationX();


  /**
   * <p>
   * Setter method for the COM property "RotationX"
   * </p>
   * @param rotationX Mandatory float parameter.
   */

  @DISPID(101) //= 0x65. The runtime will prefer the VTID if present
  @VTID(13)
  void setRotationX(
    float rotationX);


  /**
   * <p>
   * Getter method for the COM property "RotationY"
   * </p>
   * @return  Returns a value of type float
   */

  @DISPID(102) //= 0x66. The runtime will prefer the VTID if present
  @VTID(14)
  float getRotationY();


  /**
   * <p>
   * Setter method for the COM property "RotationY"
   * </p>
   * @param rotationY Mandatory float parameter.
   */

  @DISPID(102) //= 0x66. The runtime will prefer the VTID if present
  @VTID(15)
  void setRotationY(
    float rotationY);


  /**
   * <p>
   * Getter method for the COM property "RotationZ"
   * </p>
   * @return  Returns a value of type float
   */

  @DISPID(103) //= 0x67. The runtime will prefer the VTID if present
  @VTID(16)
  float getRotationZ();


  /**
   * <p>
   * Setter method for the COM property "RotationZ"
   * </p>
   * @param rotationZ Mandatory float parameter.
   */

  @DISPID(103) //= 0x67. The runtime will prefer the VTID if present
  @VTID(17)
  void setRotationZ(
    float rotationZ);


  /**
   * <p>
   * Getter method for the COM property "FieldOfView"
   * </p>
   * @return  Returns a value of type float
   */

  @DISPID(104) //= 0x68. The runtime will prefer the VTID if present
  @VTID(18)
  float getFieldOfView();


  /**
   * <p>
   * Setter method for the COM property "FieldOfView"
   * </p>
   * @param fov Mandatory float parameter.
   */

  @DISPID(104) //= 0x68. The runtime will prefer the VTID if present
  @VTID(19)
  void setFieldOfView(
    float fov);


  /**
   * <p>
   * Getter method for the COM property "CameraPositionX"
   * </p>
   * @return  Returns a value of type float
   */

  @DISPID(105) //= 0x69. The runtime will prefer the VTID if present
  @VTID(20)
  float getCameraPositionX();


  /**
   * <p>
   * Setter method for the COM property "CameraPositionX"
   * </p>
   * @param cameraPositionX Mandatory float parameter.
   */

  @DISPID(105) //= 0x69. The runtime will prefer the VTID if present
  @VTID(21)
  void setCameraPositionX(
    float cameraPositionX);


  /**
   * <p>
   * Getter method for the COM property "CameraPositionY"
   * </p>
   * @return  Returns a value of type float
   */

  @DISPID(106) //= 0x6a. The runtime will prefer the VTID if present
  @VTID(22)
  float getCameraPositionY();


  /**
   * <p>
   * Setter method for the COM property "CameraPositionY"
   * </p>
   * @param cameraPositionY Mandatory float parameter.
   */

  @DISPID(106) //= 0x6a. The runtime will prefer the VTID if present
  @VTID(23)
  void setCameraPositionY(
    float cameraPositionY);


  /**
   * <p>
   * Getter method for the COM property "CameraPositionZ"
   * </p>
   * @return  Returns a value of type float
   */

  @DISPID(107) //= 0x6b. The runtime will prefer the VTID if present
  @VTID(24)
  float getCameraPositionZ();


  /**
   * <p>
   * Setter method for the COM property "CameraPositionZ"
   * </p>
   * @param cameraPositionZ Mandatory float parameter.
   */

  @DISPID(107) //= 0x6b. The runtime will prefer the VTID if present
  @VTID(25)
  void setCameraPositionZ(
    float cameraPositionZ);


  /**
   * <p>
   * Getter method for the COM property "LookAtPointX"
   * </p>
   * @return  Returns a value of type float
   */

  @DISPID(108) //= 0x6c. The runtime will prefer the VTID if present
  @VTID(26)
  float getLookAtPointX();


  /**
   * <p>
   * Setter method for the COM property "LookAtPointX"
   * </p>
   * @param lookAtPointX Mandatory float parameter.
   */

  @DISPID(108) //= 0x6c. The runtime will prefer the VTID if present
  @VTID(27)
  void setLookAtPointX(
    float lookAtPointX);


  /**
   * <p>
   * Getter method for the COM property "LookAtPointY"
   * </p>
   * @return  Returns a value of type float
   */

  @DISPID(109) //= 0x6d. The runtime will prefer the VTID if present
  @VTID(28)
  float getLookAtPointY();


  /**
   * <p>
   * Setter method for the COM property "LookAtPointY"
   * </p>
   * @param lookAtPointY Mandatory float parameter.
   */

  @DISPID(109) //= 0x6d. The runtime will prefer the VTID if present
  @VTID(29)
  void setLookAtPointY(
    float lookAtPointY);


  /**
   * <p>
   * Getter method for the COM property "LookAtPointZ"
   * </p>
   * @return  Returns a value of type float
   */

  @DISPID(110) //= 0x6e. The runtime will prefer the VTID if present
  @VTID(30)
  float getLookAtPointZ();


  /**
   * <p>
   * Setter method for the COM property "LookAtPointZ"
   * </p>
   * @param lookAtPointZ Mandatory float parameter.
   */

  @DISPID(110) //= 0x6e. The runtime will prefer the VTID if present
  @VTID(31)
  void setLookAtPointZ(
    float lookAtPointZ);


  /**
   * <p>
   * This method uses predefined default values for the following parameters:
   * </p>
   * <ul>
     * <li>boolean parameter resetSize is set to false</li></ul>
   * <p>
   * Therefore, using this method is equivalent to
   * <code>
   * resetModel(false);
   * </code>
   * </p>
   */

  @DISPID(111) //= 0x6f. The runtime will prefer the VTID if present
  @VTID(32)
  @UseDefaultValues(paramIndexMapping = {}, optParamIndex = {0}, javaType = {boolean.class}, nativeType = {NativeType.VariantBool}, variantType = {Variant.Type.VT_BOOL}, literal = {"false"})
  void resetModel();

  /**
   * @param resetSize Optional parameter. Default value is false
   */

  @DISPID(111) //= 0x6f. The runtime will prefer the VTID if present
  @VTID(32)
  void resetModel(
    @Optional @DefaultValue("0") boolean resetSize);


  /**
   * @param increment Mandatory float parameter.
   */

  @DISPID(112) //= 0x70. The runtime will prefer the VTID if present
  @VTID(33)
  void incrementRotationX(
    float increment);


  /**
   * @param increment Mandatory float parameter.
   */

  @DISPID(113) //= 0x71. The runtime will prefer the VTID if present
  @VTID(34)
  void incrementRotationY(
    float increment);


  /**
   * @param increment Mandatory float parameter.
   */

  @DISPID(114) //= 0x72. The runtime will prefer the VTID if present
  @VTID(35)
  void incrementRotationZ(
    float increment);


  // Properties:
}
