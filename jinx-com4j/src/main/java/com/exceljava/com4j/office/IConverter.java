package com.exceljava.com4j.office  ;

import com4j.*;

@IID("{000C03D7-0000-0000-C000-000000000046}")
public interface IConverter extends Com4jObject {
  // Methods:
  /**
   * @param pcap Mandatory com.exceljava.com4j.office.IConverterApplicationPreferences parameter.
   * @param pcuic Mandatory com.exceljava.com4j.office.IConverterUICallback parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.IConverterPreferences
   */

  @VTID(3)
  @ReturnValue(index=1)
  com.exceljava.com4j.office.IConverterPreferences hrInitConverter(
    com.exceljava.com4j.office.IConverterApplicationPreferences pcap,
    com.exceljava.com4j.office.IConverterUICallback pcuic);


  /**
   * @param pcuic Mandatory com.exceljava.com4j.office.IConverterUICallback parameter.
   */

  @VTID(4)
  void hrUninitConverter(
    com.exceljava.com4j.office.IConverterUICallback pcuic);


  /**
   * @param bstrSourcePath Mandatory java.lang.String parameter.
   * @param bstrDestPath Mandatory java.lang.String parameter.
   * @param pcap Mandatory com.exceljava.com4j.office.IConverterApplicationPreferences parameter.
   * @param pcuic Mandatory com.exceljava.com4j.office.IConverterUICallback parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.IConverterPreferences
   */

  @VTID(5)
  @ReturnValue(index=3)
  com.exceljava.com4j.office.IConverterPreferences hrImport(
    java.lang.String bstrSourcePath,
    java.lang.String bstrDestPath,
    com.exceljava.com4j.office.IConverterApplicationPreferences pcap,
    com.exceljava.com4j.office.IConverterUICallback pcuic);


  /**
   * @param bstrSourcePath Mandatory java.lang.String parameter.
   * @param bstrDestPath Mandatory java.lang.String parameter.
   * @param bstrClass Mandatory java.lang.String parameter.
   * @param pcap Mandatory com.exceljava.com4j.office.IConverterApplicationPreferences parameter.
   * @param pcuic Mandatory com.exceljava.com4j.office.IConverterUICallback parameter.
   * @return  Returns a value of type com.exceljava.com4j.office.IConverterPreferences
   */

  @VTID(6)
  @ReturnValue(index=4)
  com.exceljava.com4j.office.IConverterPreferences hrExport(
    java.lang.String bstrSourcePath,
    java.lang.String bstrDestPath,
    java.lang.String bstrClass,
    com.exceljava.com4j.office.IConverterApplicationPreferences pcap,
    com.exceljava.com4j.office.IConverterUICallback pcuic);


  /**
   * @param bstrPath Mandatory java.lang.String parameter.
   * @param pbstrClass Mandatory Holder<java.lang.String> parameter.
   * @param pcap Mandatory com.exceljava.com4j.office.IConverterApplicationPreferences parameter.
   * @param ppcp Mandatory Holder<com.exceljava.com4j.office.IConverterPreferences> parameter.
   * @param pcuic Mandatory com.exceljava.com4j.office.IConverterUICallback parameter.
   */

  @VTID(7)
  void hrGetFormat(
    java.lang.String bstrPath,
    Holder<java.lang.String> pbstrClass,
    com.exceljava.com4j.office.IConverterApplicationPreferences pcap,
    Holder<com.exceljava.com4j.office.IConverterPreferences> ppcp,
    com.exceljava.com4j.office.IConverterUICallback pcuic);


  /**
   * @param hrErr Mandatory int parameter.
   * @param pcap Mandatory com.exceljava.com4j.office.IConverterApplicationPreferences parameter.
   * @return  Returns a value of type java.lang.String
   */

  @VTID(8)
  @ReturnValue(index=1)
  java.lang.String hrGetErrorString(
    int hrErr,
    com.exceljava.com4j.office.IConverterApplicationPreferences pcap);


  // Properties:
}
