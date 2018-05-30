package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C1533-0000-0000-C000-000000000046}")
public interface ODSOFilter extends com.exceljava.com4j.office._IMsoDispObj {
  // Methods:
  /**
   * <p>
   * Getter method for the COM property "Index"
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(1) //= 0x1. The runtime will prefer the VTID if present
  @VTID(9)
  int getIndex();


  /**
   * <p>
   * Getter method for the COM property "Parent"
   * </p>
   * @return  Returns a value of type com4j.Com4jObject
   */

  @DISPID(2) //= 0x2. The runtime will prefer the VTID if present
  @VTID(10)
  @ReturnValue(type=NativeType.Dispatch)
  com4j.Com4jObject getParent();


  /**
   * <p>
   * Getter method for the COM property "Column"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(3) //= 0x3. The runtime will prefer the VTID if present
  @VTID(11)
  java.lang.String getColumn();


  /**
   * <p>
   * Setter method for the COM property "Column"
   * </p>
   * @param pbstrCol Mandatory java.lang.String parameter.
   */

  @DISPID(3) //= 0x3. The runtime will prefer the VTID if present
  @VTID(12)
  void setColumn(
    java.lang.String pbstrCol);


  /**
   * <p>
   * Getter method for the COM property "Comparison"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoFilterComparison
   */

  @DISPID(4) //= 0x4. The runtime will prefer the VTID if present
  @VTID(13)
  com.exceljava.com4j.office.MsoFilterComparison getComparison();


  /**
   * <p>
   * Setter method for the COM property "Comparison"
   * </p>
   * @param pComparison Mandatory com.exceljava.com4j.office.MsoFilterComparison parameter.
   */

  @DISPID(4) //= 0x4. The runtime will prefer the VTID if present
  @VTID(14)
  void setComparison(
    com.exceljava.com4j.office.MsoFilterComparison pComparison);


  /**
   * <p>
   * Getter method for the COM property "CompareTo"
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(5) //= 0x5. The runtime will prefer the VTID if present
  @VTID(15)
  java.lang.String getCompareTo();


  /**
   * <p>
   * Setter method for the COM property "CompareTo"
   * </p>
   * @param pbstrCompareTo Mandatory java.lang.String parameter.
   */

  @DISPID(5) //= 0x5. The runtime will prefer the VTID if present
  @VTID(16)
  void setCompareTo(
    java.lang.String pbstrCompareTo);


  /**
   * <p>
   * Getter method for the COM property "Conjunction"
   * </p>
   * @return  Returns a value of type com.exceljava.com4j.office.MsoFilterConjunction
   */

  @DISPID(6) //= 0x6. The runtime will prefer the VTID if present
  @VTID(17)
  com.exceljava.com4j.office.MsoFilterConjunction getConjunction();


  /**
   * <p>
   * Setter method for the COM property "Conjunction"
   * </p>
   * @param pConjunction Mandatory com.exceljava.com4j.office.MsoFilterConjunction parameter.
   */

  @DISPID(6) //= 0x6. The runtime will prefer the VTID if present
  @VTID(18)
  void setConjunction(
    com.exceljava.com4j.office.MsoFilterConjunction pConjunction);


  // Properties:
}
