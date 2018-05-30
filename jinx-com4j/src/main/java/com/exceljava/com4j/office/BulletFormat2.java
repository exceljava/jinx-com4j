package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C03B9-0000-0000-C000-000000000046}")
public interface BulletFormat2 extends com.exceljava.com4j.office._IMsoDispObj {
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
   * Getter method for the COM property "Character"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(10)
  int getCharacter();


  /**
   * <p>
   * Setter method for the COM property "Character"
   * </p>
   * @param character Mandatory int parameter.
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(11)
  void setCharacter(
    int character);


  /**
   * <p>
   * Getter method for the COM property "Font"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.Font2
   */

  @DISPID(3) //= 0x3. The runtime will prefer the VTID if present
  @VTID(12)
  com.exceljava.com4j.office.Font2 getFont();


  /**
   * <p>
   * Getter method for the COM property "Number"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(4) //= 0x4. The runtime will prefer the VTID if present
  @VTID(13)
  int getNumber();


  /**
   * @param fileName Mandatory java.lang.String parameter.
   */

  @DISPID(5) //= 0x5. The runtime will prefer the VTID if present
  @VTID(14)
  void picture(
    java.lang.String fileName);


  /**
   * <p>
   * Getter method for the COM property "RelativeSize"
   * </p>
   * @return  Returns a value of type float
   */

  @DISPID(6) //= 0x6. The runtime will prefer the VTID if present
  @VTID(15)
  float getRelativeSize();


  /**
   * <p>
   * Setter method for the COM property "RelativeSize"
   * </p>
   * @param size Mandatory float parameter.
   */

  @DISPID(6) //= 0x6. The runtime will prefer the VTID if present
  @VTID(16)
  void setRelativeSize(
    float size);


  /**
   * <p>
   * Getter method for the COM property "StartValue"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(7) //= 0x7. The runtime will prefer the VTID if present
  @VTID(17)
  int getStartValue();


  /**
   * <p>
   * Setter method for the COM property "StartValue"
   * </p>
   * @param start Mandatory int parameter.
   */

  @DISPID(7) //= 0x7. The runtime will prefer the VTID if present
  @VTID(18)
  void setStartValue(
    int start);


  /**
   * <p>
   * Getter method for the COM property "Style"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoNumberedBulletStyle
   */

  @DISPID(8) //= 0x8. The runtime will prefer the VTID if present
  @VTID(19)
  com.exceljava.com4j.office.MsoNumberedBulletStyle getStyle();


  /**
   * <p>
   * Setter method for the COM property "Style"
   * </p>
   * @param style Mandatory com.exceljava.com4j.office.MsoNumberedBulletStyle parameter.
   */

  @DISPID(8) //= 0x8. The runtime will prefer the VTID if present
  @VTID(20)
  void setStyle(
    com.exceljava.com4j.office.MsoNumberedBulletStyle style);


  /**
   * <p>
   * Getter method for the COM property "Type"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoBulletType
   */

  @DISPID(9) //= 0x9. The runtime will prefer the VTID if present
  @VTID(21)
  com.exceljava.com4j.office.MsoBulletType getType();


  /**
   * <p>
   * Setter method for the COM property "Type"
   * </p>
   * @param type Mandatory com.exceljava.com4j.office.MsoBulletType parameter.
   */

  @DISPID(9) //= 0x9. The runtime will prefer the VTID if present
  @VTID(22)
  void setType(
    com.exceljava.com4j.office.MsoBulletType type);


  /**
   * <p>
   * Getter method for the COM property "UseTextColor"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTriState
   */

  @DISPID(10) //= 0xa. The runtime will prefer the VTID if present
  @VTID(23)
  com.exceljava.com4j.office.MsoTriState getUseTextColor();


  /**
   * <p>
   * Setter method for the COM property "UseTextColor"
   * </p>
   * @param useTextColor Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   */

  @DISPID(10) //= 0xa. The runtime will prefer the VTID if present
  @VTID(24)
  void setUseTextColor(
    com.exceljava.com4j.office.MsoTriState useTextColor);


  /**
   * <p>
   * Getter method for the COM property "UseTextFont"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTriState
   */

  @DISPID(11) //= 0xb. The runtime will prefer the VTID if present
  @VTID(25)
  com.exceljava.com4j.office.MsoTriState getUseTextFont();


  /**
   * <p>
   * Setter method for the COM property "UseTextFont"
   * </p>
   * @param useTextFont Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   */

  @DISPID(11) //= 0xb. The runtime will prefer the VTID if present
  @VTID(26)
  void setUseTextFont(
    com.exceljava.com4j.office.MsoTriState useTextFont);


  /**
   * <p>
   * Getter method for the COM property "Visible"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoTriState
   */

  @DISPID(12) //= 0xc. The runtime will prefer the VTID if present
  @VTID(27)
  com.exceljava.com4j.office.MsoTriState getVisible();


  /**
   * <p>
   * Setter method for the COM property "Visible"
   * </p>
   * @param visible Mandatory com.exceljava.com4j.office.MsoTriState parameter.
   */

  @DISPID(12) //= 0xc. The runtime will prefer the VTID if present
  @VTID(28)
  void setVisible(
    com.exceljava.com4j.office.MsoTriState visible);


  // Properties:
}
