package com.exceljava.com4j.office  ;

import com4j.*;

/**
 * ILicAgent Interface
 */
@IID("{00194002-D9C3-11D3-8D59-0050048384E3}")
public interface ILicAgent extends Com4jObject {
  // Methods:
  /**
   * <p>
   * method Initialize
   * </p>
   * @param dwBPC Mandatory int parameter.
   * @param dwMode Mandatory int parameter.
   * @param bstrLicSource Mandatory java.lang.String parameter.
   * @return  Returns a value of type int
   */

  @DISPID(1) //= 0x1. The runtime will prefer the VTID if present
  @VTID(7)
  int initialize(
    int dwBPC,
    int dwMode,
    java.lang.String bstrLicSource);


  /**
   * <p>
   * method GetFirstName
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(3) //= 0x3. The runtime will prefer the VTID if present
  @VTID(8)
  java.lang.String getFirstName();


  /**
   * <p>
   * method SetFirstName
   * </p>
   * @param bstrNewVal Mandatory java.lang.String parameter.
   */

  @DISPID(4) //= 0x4. The runtime will prefer the VTID if present
  @VTID(9)
  void setFirstName(
    java.lang.String bstrNewVal);


  /**
   * <p>
   * method GetLastName
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(5) //= 0x5. The runtime will prefer the VTID if present
  @VTID(10)
  java.lang.String getLastName();


  /**
   * <p>
   * method SetLastName
   * </p>
   * @param bstrNewVal Mandatory java.lang.String parameter.
   */

  @DISPID(6) //= 0x6. The runtime will prefer the VTID if present
  @VTID(11)
  void setLastName(
    java.lang.String bstrNewVal);


  /**
   * <p>
   * method GetOrgName
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(7) //= 0x7. The runtime will prefer the VTID if present
  @VTID(12)
  java.lang.String getOrgName();


  /**
   * <p>
   * method SetOrgName
   * </p>
   * @param bstrNewVal Mandatory java.lang.String parameter.
   */

  @DISPID(8) //= 0x8. The runtime will prefer the VTID if present
  @VTID(13)
  void setOrgName(
    java.lang.String bstrNewVal);


  /**
   * <p>
   * method GetEmail
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(9) //= 0x9. The runtime will prefer the VTID if present
  @VTID(14)
  java.lang.String getEmail();


  /**
   * <p>
   * method SetEmail
   * </p>
   * @param bstrNewVal Mandatory java.lang.String parameter.
   */

  @DISPID(10) //= 0xa. The runtime will prefer the VTID if present
  @VTID(15)
  void setEmail(
    java.lang.String bstrNewVal);


  /**
   * <p>
   * method GetPhone
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(11) //= 0xb. The runtime will prefer the VTID if present
  @VTID(16)
  java.lang.String getPhone();


  /**
   * <p>
   * method SetPhone
   * </p>
   * @param bstrNewVal Mandatory java.lang.String parameter.
   */

  @DISPID(12) //= 0xc. The runtime will prefer the VTID if present
  @VTID(17)
  void setPhone(
    java.lang.String bstrNewVal);


  /**
   * <p>
   * method GetAddress1
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(13) //= 0xd. The runtime will prefer the VTID if present
  @VTID(18)
  java.lang.String getAddress1();


  /**
   * <p>
   * method SetAddress1
   * </p>
   * @param bstrNewVal Mandatory java.lang.String parameter.
   */

  @DISPID(14) //= 0xe. The runtime will prefer the VTID if present
  @VTID(19)
  void setAddress1(
    java.lang.String bstrNewVal);


  /**
   * <p>
   * method GetCity
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(15) //= 0xf. The runtime will prefer the VTID if present
  @VTID(20)
  java.lang.String getCity();


  /**
   * <p>
   * method SetCity
   * </p>
   * @param bstrNewVal Mandatory java.lang.String parameter.
   */

  @DISPID(16) //= 0x10. The runtime will prefer the VTID if present
  @VTID(21)
  void setCity(
    java.lang.String bstrNewVal);


  /**
   * <p>
   * method GetState
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(17) //= 0x11. The runtime will prefer the VTID if present
  @VTID(22)
  java.lang.String getState();


  /**
   * <p>
   * method SetState
   * </p>
   * @param bstrNewVal Mandatory java.lang.String parameter.
   */

  @DISPID(18) //= 0x12. The runtime will prefer the VTID if present
  @VTID(23)
  void setState(
    java.lang.String bstrNewVal);


  /**
   * <p>
   * method GetCountryCode
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(19) //= 0x13. The runtime will prefer the VTID if present
  @VTID(24)
  java.lang.String getCountryCode();


  /**
   * <p>
   * method SetCountryCode
   * </p>
   * @param bstrNewVal Mandatory java.lang.String parameter.
   */

  @DISPID(20) //= 0x14. The runtime will prefer the VTID if present
  @VTID(25)
  void setCountryCode(
    java.lang.String bstrNewVal);


  /**
   * <p>
   * method GetCountryDesc
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(21) //= 0x15. The runtime will prefer the VTID if present
  @VTID(26)
  java.lang.String getCountryDesc();


  /**
   * <p>
   * method SetCountryDesc
   * </p>
   * @param bstrNewVal Mandatory java.lang.String parameter.
   */

  @DISPID(22) //= 0x16. The runtime will prefer the VTID if present
  @VTID(27)
  void setCountryDesc(
    java.lang.String bstrNewVal);


  /**
   * <p>
   * method GetZip
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(23) //= 0x17. The runtime will prefer the VTID if present
  @VTID(28)
  java.lang.String getZip();


  /**
   * <p>
   * method SetZip
   * </p>
   * @param bstrNewVal Mandatory java.lang.String parameter.
   */

  @DISPID(24) //= 0x18. The runtime will prefer the VTID if present
  @VTID(29)
  void setZip(
    java.lang.String bstrNewVal);


  /**
   * <p>
   * method GetIsoLanguage
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(25) //= 0x19. The runtime will prefer the VTID if present
  @VTID(30)
  int getIsoLanguage();


  /**
   * <p>
   * method SetIsoLanguage
   * </p>
   * @param dwNewVal Mandatory int parameter.
   */

  @DISPID(26) //= 0x1a. The runtime will prefer the VTID if present
  @VTID(31)
  void setIsoLanguage(
    int dwNewVal);


  /**
   * <p>
   * method GetMSUpdate
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(32) //= 0x20. The runtime will prefer the VTID if present
  @VTID(32)
  java.lang.String getMSUpdate();


  /**
   * <p>
   * method SetMSUpdate
   * </p>
   * @param bstrNewVal Mandatory java.lang.String parameter.
   */

  @DISPID(33) //= 0x21. The runtime will prefer the VTID if present
  @VTID(33)
  void setMSUpdate(
    java.lang.String bstrNewVal);


  /**
   * <p>
   * method GetMSOffer
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(34) //= 0x22. The runtime will prefer the VTID if present
  @VTID(34)
  java.lang.String getMSOffer();


  /**
   * <p>
   * method SetMSOffer
   * </p>
   * @param bstrNewVal Mandatory java.lang.String parameter.
   */

  @DISPID(35) //= 0x23. The runtime will prefer the VTID if present
  @VTID(35)
  void setMSOffer(
    java.lang.String bstrNewVal);


  /**
   * <p>
   * method GetOtherOffer
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(36) //= 0x24. The runtime will prefer the VTID if present
  @VTID(36)
  java.lang.String getOtherOffer();


  /**
   * <p>
   * method SetOtherOffer
   * </p>
   * @param bstrNewVal Mandatory java.lang.String parameter.
   */

  @DISPID(37) //= 0x25. The runtime will prefer the VTID if present
  @VTID(37)
  void setOtherOffer(
    java.lang.String bstrNewVal);


  /**
   * <p>
   * method GetAddress2
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(38) //= 0x26. The runtime will prefer the VTID if present
  @VTID(38)
  java.lang.String getAddress2();


  /**
   * <p>
   * method SetAddress2
   * </p>
   * @param bstrNewVal Mandatory java.lang.String parameter.
   */

  @DISPID(39) //= 0x27. The runtime will prefer the VTID if present
  @VTID(39)
  void setAddress2(
    java.lang.String bstrNewVal);


  /**
   * <p>
   * method CheckSystemClock
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(40) //= 0x28. The runtime will prefer the VTID if present
  @VTID(40)
  int checkSystemClock();


  /**
   * <p>
   * method GetExistingExpiryDate
   * </p>
   * @return  Returns a value of type java.util.Date
   */

  @DISPID(41) //= 0x29. The runtime will prefer the VTID if present
  @VTID(41)
  java.util.Date getExistingExpiryDate();


  /**
   * <p>
   * method GetNewExpiryDate
   * </p>
   * @return  Returns a value of type java.util.Date
   */

  @DISPID(42) //= 0x2a. The runtime will prefer the VTID if present
  @VTID(42)
  java.util.Date getNewExpiryDate();


  /**
   * <p>
   * method GetBillingFirstName
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(43) //= 0x2b. The runtime will prefer the VTID if present
  @VTID(43)
  java.lang.String getBillingFirstName();


  /**
   * <p>
   * method SetBillingFirstName
   * </p>
   * @param bstrNewVal Mandatory java.lang.String parameter.
   */

  @DISPID(44) //= 0x2c. The runtime will prefer the VTID if present
  @VTID(44)
  void setBillingFirstName(
    java.lang.String bstrNewVal);


  /**
   * <p>
   * method GetBillingLastName
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(45) //= 0x2d. The runtime will prefer the VTID if present
  @VTID(45)
  java.lang.String getBillingLastName();


  /**
   * <p>
   * method SetBillingLastName
   * </p>
   * @param bstrNewVal Mandatory java.lang.String parameter.
   */

  @DISPID(46) //= 0x2e. The runtime will prefer the VTID if present
  @VTID(46)
  void setBillingLastName(
    java.lang.String bstrNewVal);


  /**
   * <p>
   * method GetBillingPhone
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(47) //= 0x2f. The runtime will prefer the VTID if present
  @VTID(47)
  java.lang.String getBillingPhone();


  /**
   * <p>
   * method SetBillingPhone
   * </p>
   * @param bstrNewVal Mandatory java.lang.String parameter.
   */

  @DISPID(48) //= 0x30. The runtime will prefer the VTID if present
  @VTID(48)
  void setBillingPhone(
    java.lang.String bstrNewVal);


  /**
   * <p>
   * method GetBillingAddress1
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(49) //= 0x31. The runtime will prefer the VTID if present
  @VTID(49)
  java.lang.String getBillingAddress1();


  /**
   * <p>
   * method SetBillingAddress1
   * </p>
   * @param bstrNewVal Mandatory java.lang.String parameter.
   */

  @DISPID(50) //= 0x32. The runtime will prefer the VTID if present
  @VTID(50)
  void setBillingAddress1(
    java.lang.String bstrNewVal);


  /**
   * <p>
   * method GetBillingAddress2
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(51) //= 0x33. The runtime will prefer the VTID if present
  @VTID(51)
  java.lang.String getBillingAddress2();


  /**
   * <p>
   * method SetBillingAddress2
   * </p>
   * @param bstrNewVal Mandatory java.lang.String parameter.
   */

  @DISPID(52) //= 0x34. The runtime will prefer the VTID if present
  @VTID(52)
  void setBillingAddress2(
    java.lang.String bstrNewVal);


  /**
   * <p>
   * method GetBillingCity
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(53) //= 0x35. The runtime will prefer the VTID if present
  @VTID(53)
  java.lang.String getBillingCity();


  /**
   * <p>
   * method SetBillingCity
   * </p>
   * @param bstrNewVal Mandatory java.lang.String parameter.
   */

  @DISPID(54) //= 0x36. The runtime will prefer the VTID if present
  @VTID(54)
  void setBillingCity(
    java.lang.String bstrNewVal);


  /**
   * <p>
   * method GetBillingState
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(55) //= 0x37. The runtime will prefer the VTID if present
  @VTID(55)
  java.lang.String getBillingState();


  /**
   * <p>
   * method SetBillingState
   * </p>
   * @param bstrNewVal Mandatory java.lang.String parameter.
   */

  @DISPID(56) //= 0x38. The runtime will prefer the VTID if present
  @VTID(56)
  void setBillingState(
    java.lang.String bstrNewVal);


  /**
   * <p>
   * method GetBillingCountryCode
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(57) //= 0x39. The runtime will prefer the VTID if present
  @VTID(57)
  java.lang.String getBillingCountryCode();


  /**
   * <p>
   * method SetBillingCountryCode
   * </p>
   * @param bstrNewVal Mandatory java.lang.String parameter.
   */

  @DISPID(58) //= 0x3a. The runtime will prefer the VTID if present
  @VTID(58)
  void setBillingCountryCode(
    java.lang.String bstrNewVal);


  /**
   * <p>
   * method GetBillingZip
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(59) //= 0x3b. The runtime will prefer the VTID if present
  @VTID(59)
  java.lang.String getBillingZip();


  /**
   * <p>
   * method SetBillingZip
   * </p>
   * @param bstrNewVal Mandatory java.lang.String parameter.
   */

  @DISPID(60) //= 0x3c. The runtime will prefer the VTID if present
  @VTID(60)
  void setBillingZip(
    java.lang.String bstrNewVal);


  /**
   * <p>
   * method SaveBillingInfo
   * </p>
   * @param bSave Mandatory int parameter.
   * @return  Returns a value of type int
   */

  @DISPID(61) //= 0x3d. The runtime will prefer the VTID if present
  @VTID(61)
  int saveBillingInfo(
    int bSave);


  /**
   * <p>
   * method IsCCRenewalCountry
   * </p>
   * @param bstrCountryCode Mandatory java.lang.String parameter.
   * @return  Returns a value of type int
   */

  @DISPID(64) //= 0x40. The runtime will prefer the VTID if present
  @VTID(62)
  int isCCRenewalCountry(
    java.lang.String bstrCountryCode);


  /**
   * <p>
   * method GetVATLabel
   * </p>
   * @param bstrCountryCode Mandatory java.lang.String parameter.
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(65) //= 0x41. The runtime will prefer the VTID if present
  @VTID(63)
  java.lang.String getVATLabel(
    java.lang.String bstrCountryCode);


  /**
   * <p>
   * method GetCCRenewalExpiryDate
   * </p>
   * @return  Returns a value of type java.util.Date
   */

  @DISPID(66) //= 0x42. The runtime will prefer the VTID if present
  @VTID(64)
  java.util.Date getCCRenewalExpiryDate();


  /**
   * <p>
   * method SetVATNumber
   * </p>
   * @param bstrVATNumber Mandatory java.lang.String parameter.
   */

  @DISPID(67) //= 0x43. The runtime will prefer the VTID if present
  @VTID(65)
  void setVATNumber(
    java.lang.String bstrVATNumber);


  /**
   * <p>
   * method SetCreditCardType
   * </p>
   * @param bstrCCCode Mandatory java.lang.String parameter.
   */

  @DISPID(68) //= 0x44. The runtime will prefer the VTID if present
  @VTID(66)
  void setCreditCardType(
    java.lang.String bstrCCCode);


  /**
   * <p>
   * method SetCreditCardNumber
   * </p>
   * @param bstrCCNumber Mandatory java.lang.String parameter.
   */

  @DISPID(69) //= 0x45. The runtime will prefer the VTID if present
  @VTID(67)
  void setCreditCardNumber(
    java.lang.String bstrCCNumber);


  /**
   * <p>
   * method SetCreditCardExpiryYear
   * </p>
   * @param dwCCYear Mandatory int parameter.
   */

  @DISPID(70) //= 0x46. The runtime will prefer the VTID if present
  @VTID(68)
  void setCreditCardExpiryYear(
    int dwCCYear);


  /**
   * <p>
   * method SetCreditCardExpiryMonth
   * </p>
   * @param dwCCMonth Mandatory int parameter.
   */

  @DISPID(71) //= 0x47. The runtime will prefer the VTID if present
  @VTID(69)
  void setCreditCardExpiryMonth(
    int dwCCMonth);


  /**
   * <p>
   * method GetCreditCardCount
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(72) //= 0x48. The runtime will prefer the VTID if present
  @VTID(70)
  int getCreditCardCount();


  /**
   * <p>
   * method GetCreditCardCode
   * </p>
   * @param dwIndex Mandatory int parameter.
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(73) //= 0x49. The runtime will prefer the VTID if present
  @VTID(71)
  java.lang.String getCreditCardCode(
    int dwIndex);


  /**
   * <p>
   * method GetCreditCardName
   * </p>
   * @param dwIndex Mandatory int parameter.
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(74) //= 0x4a. The runtime will prefer the VTID if present
  @VTID(72)
  java.lang.String getCreditCardName(
    int dwIndex);


  /**
   * <p>
   * method GetVATNumber
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(75) //= 0x4b. The runtime will prefer the VTID if present
  @VTID(73)
  java.lang.String getVATNumber();


  /**
   * <p>
   * method GetCreditCardType
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(76) //= 0x4c. The runtime will prefer the VTID if present
  @VTID(74)
  java.lang.String getCreditCardType();


  /**
   * <p>
   * method GetCreditCardNumber
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(77) //= 0x4d. The runtime will prefer the VTID if present
  @VTID(75)
  java.lang.String getCreditCardNumber();


  /**
   * <p>
   * method GetCreditCardExpiryYear
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(78) //= 0x4e. The runtime will prefer the VTID if present
  @VTID(76)
  int getCreditCardExpiryYear();


  /**
   * <p>
   * method GetCreditCardExpiryMonth
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(79) //= 0x4f. The runtime will prefer the VTID if present
  @VTID(77)
  int getCreditCardExpiryMonth();


  /**
   * <p>
   * method GetDisconnectOption
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(80) //= 0x50. The runtime will prefer the VTID if present
  @VTID(78)
  int getDisconnectOption();


  /**
   * <p>
   * method SetDisconnectOption
   * </p>
   * @param bNewVal Mandatory int parameter.
   */

  @DISPID(81) //= 0x51. The runtime will prefer the VTID if present
  @VTID(79)
  void setDisconnectOption(
    int bNewVal);


  /**
   * <p>
   * method AsyncProcessHandshakeRequest
   * </p>
   * @param bReviseCustInfo Mandatory int parameter.
   */

  @DISPID(82) //= 0x52. The runtime will prefer the VTID if present
  @VTID(80)
  void asyncProcessHandshakeRequest(
    int bReviseCustInfo);


  /**
   * <p>
   * method AsyncProcessNewLicenseRequest
   * </p>
   */

  @DISPID(83) //= 0x53. The runtime will prefer the VTID if present
  @VTID(81)
  void asyncProcessNewLicenseRequest();


  /**
   * <p>
   * method AsyncProcessReissueLicenseRequest
   * </p>
   */

  @DISPID(84) //= 0x54. The runtime will prefer the VTID if present
  @VTID(82)
  void asyncProcessReissueLicenseRequest();


  /**
   * <p>
   * method AsyncProcessRetailRenewalLicenseRequest
   * </p>
   */

  @DISPID(85) //= 0x55. The runtime will prefer the VTID if present
  @VTID(83)
  void asyncProcessRetailRenewalLicenseRequest();


  /**
   * <p>
   * method AsyncProcessReviseCustInfoRequest
   * </p>
   */

  @DISPID(86) //= 0x56. The runtime will prefer the VTID if present
  @VTID(84)
  void asyncProcessReviseCustInfoRequest();


  /**
   * <p>
   * method AsyncProcessCCRenewalPriceRequest
   * </p>
   */

  @DISPID(87) //= 0x57. The runtime will prefer the VTID if present
  @VTID(85)
  void asyncProcessCCRenewalPriceRequest();


  /**
   * <p>
   * method AsyncProcessCCRenewalLicenseRequest
   * </p>
   */

  @DISPID(88) //= 0x58. The runtime will prefer the VTID if present
  @VTID(86)
  void asyncProcessCCRenewalLicenseRequest();


  /**
   * <p>
   * method GetAsyncProcessReturnCode
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(90) //= 0x5a. The runtime will prefer the VTID if present
  @VTID(87)
  int getAsyncProcessReturnCode();


  /**
   * <p>
   * method IsUpgradeAvailable
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(91) //= 0x5b. The runtime will prefer the VTID if present
  @VTID(88)
  int isUpgradeAvailable();


  /**
   * <p>
   * method WantUpgrade
   * </p>
   * @param bWantUpgrade Mandatory int parameter.
   */

  @DISPID(92) //= 0x5c. The runtime will prefer the VTID if present
  @VTID(89)
  void wantUpgrade(
    int bWantUpgrade);


  /**
   * <p>
   * method AsyncProcessDroppedLicenseRequest
   * </p>
   */

  @DISPID(93) //= 0x5d. The runtime will prefer the VTID if present
  @VTID(90)
  void asyncProcessDroppedLicenseRequest();


  /**
   * <p>
   * method GenerateInstallationId
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(94) //= 0x5e. The runtime will prefer the VTID if present
  @VTID(91)
  java.lang.String generateInstallationId();


  /**
   * <p>
   * method DepositConfirmationId
   * </p>
   * @param bstrVal Mandatory java.lang.String parameter.
   * @return  Returns a value of type int
   */

  @DISPID(95) //= 0x5f. The runtime will prefer the VTID if present
  @VTID(92)
  int depositConfirmationId(
    java.lang.String bstrVal);


  /**
   * <p>
   * method VerifyCheckDigits
   * </p>
   * @param bstrCIDIID Mandatory java.lang.String parameter.
   * @return  Returns a value of type int
   */

  @DISPID(96) //= 0x60. The runtime will prefer the VTID if present
  @VTID(93)
  int verifyCheckDigits(
    java.lang.String bstrCIDIID);


  /**
   * <p>
   * method GetCurrentExpiryDate
   * </p>
   * @return  Returns a value of type java.util.Date
   */

  @DISPID(97) //= 0x61. The runtime will prefer the VTID if present
  @VTID(94)
  java.util.Date getCurrentExpiryDate();


  /**
   * <p>
   * method CancelAsyncProcessRequest
   * </p>
   * @param bIsLicenseRequest Mandatory int parameter.
   */

  @DISPID(98) //= 0x62. The runtime will prefer the VTID if present
  @VTID(95)
  void cancelAsyncProcessRequest(
    int bIsLicenseRequest);


  /**
   * <p>
   * method GetCurrencyDescription
   * </p>
   * @param dwCurrencyIndex Mandatory int parameter.
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(100) //= 0x64. The runtime will prefer the VTID if present
  @VTID(96)
  java.lang.String getCurrencyDescription(
    int dwCurrencyIndex);


  /**
   * <p>
   * method GetPriceItemCount
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(101) //= 0x65. The runtime will prefer the VTID if present
  @VTID(97)
  int getPriceItemCount();


  /**
   * <p>
   * method GetPriceItemLabel
   * </p>
   * @param dwIndex Mandatory int parameter.
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(102) //= 0x66. The runtime will prefer the VTID if present
  @VTID(98)
  java.lang.String getPriceItemLabel(
    int dwIndex);


  /**
   * <p>
   * method GetPriceItemValue
   * </p>
   * @param dwCurrencyIndex Mandatory int parameter.
   * @param dwIndex Mandatory int parameter.
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(103) //= 0x67. The runtime will prefer the VTID if present
  @VTID(99)
  java.lang.String getPriceItemValue(
    int dwCurrencyIndex,
    int dwIndex);


  /**
   * <p>
   * method GetInvoiceText
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(104) //= 0x68. The runtime will prefer the VTID if present
  @VTID(100)
  java.lang.String getInvoiceText();


  /**
   * <p>
   * method GetBackendErrorMsg
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(105) //= 0x69. The runtime will prefer the VTID if present
  @VTID(101)
  java.lang.String getBackendErrorMsg();


  /**
   * <p>
   * method GetCurrencyOption
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(106) //= 0x6a. The runtime will prefer the VTID if present
  @VTID(102)
  int getCurrencyOption();


  /**
   * <p>
   * method SetCurrencyOption
   * </p>
   * @param dwCurrencyOption Mandatory int parameter.
   */

  @DISPID(107) //= 0x6b. The runtime will prefer the VTID if present
  @VTID(103)
  void setCurrencyOption(
    int dwCurrencyOption);


  /**
   * <p>
   * method GetEndOfLifeHtmlText
   * </p>
   * @return  Returns a value of type java.lang.String
   */

  @DISPID(108) //= 0x6c. The runtime will prefer the VTID if present
  @VTID(104)
  java.lang.String getEndOfLifeHtmlText();


  /**
   * <p>
   * method DisplaySSLCert
   * </p>
   * @return  Returns a value of type int
   */

  @DISPID(109) //= 0x6d. The runtime will prefer the VTID if present
  @VTID(105)
  int displaySSLCert();


  // Properties:
}
