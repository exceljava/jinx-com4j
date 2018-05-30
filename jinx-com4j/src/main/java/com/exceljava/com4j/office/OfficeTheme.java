package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C03A0-0000-0000-C000-000000000046}")
public interface OfficeTheme extends com.exceljava.com4j.office._IMsoDispObj {
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
   * Getter method for the COM property "ThemeColorScheme"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.ThemeColorScheme
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(10)
  com.exceljava.com4j.office.ThemeColorScheme getThemeColorScheme();


  @VTID(10)
  @ReturnValue(defaultPropertyThrough={com.exceljava.com4j.office.ThemeColorScheme.class})
  com.exceljava.com4j.office.ThemeColor getThemeColorScheme(
    com.exceljava.com4j.office.MsoThemeColorSchemeIndex index);

  /**
   * <p>
   * Getter method for the COM property "ThemeFontScheme"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.ThemeFontScheme
   */

  @DISPID(3) //= 0x3. The runtime will prefer the VTID if present
  @VTID(11)
  com.exceljava.com4j.office.ThemeFontScheme getThemeFontScheme();


  /**
   * <p>
   * Getter method for the COM property "ThemeEffectScheme"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.ThemeEffectScheme
   */

  @DISPID(4) //= 0x4. The runtime will prefer the VTID if present
  @VTID(12)
  com.exceljava.com4j.office.ThemeEffectScheme getThemeEffectScheme();


  // Properties:
}
