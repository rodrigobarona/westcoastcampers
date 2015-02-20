<meta http-equiv="Content-type" content="text/html; charset=utf-8" />

<%
ignorehack = True
allowanything = False
%>
<!--#include file="a_server-checks.asp"-->
<!-- #include file="include_meta.asp" -->

<head>

<title></title>

</head>
<body  class="webstep6" >

<!-- #include file="include_header.asp" -->

<%
FUNCTION tidyup(thisString)
    thisString=LTrim(RTrim(Replace(thisString, "'" , "")))
      thisString=Replace(thisString, "”" , "")
      thisString=Replace(thisString, "<" , "")
      thisString=Replace(thisString, "http" , "")
      thisString=Replace(thisString, "://" , "")
      thisString=Replace(thisString, "=" , "")
       thisString=Replace(thisString, "?" , " ")
      thisString=Replace(thisString, "/" , "")
      thisString=Replace(thisString, "+" , "")
      thisString=Replace(thisString, "%" , "")
      tidyup=Replace(thisString, ">" , "")
END FUNCTION

'AJ 06-Feb-2015: Function to enable us to read UTF-8 Text Files within classic ASP
FUNCTION ReadFileUTF8(sFilePath)
      Dim objStream, strData
      strData = ""
      If (FileObject.FileExists(sFilePath)) Then
         Set objStream = CreateObject("ADODB.Stream")
         objStream.CharSet = "utf-8"
         objStream.Open
         objStream.LoadFromFile(sFilePath)
         strData = objStream.ReadText()
         objStream.Close
      End If
      ReadFileUTF8 = strData
END FUNCTION


%>



<%       if Session("RCM273_OnlineBookingNo")="" or ISNULL(Session("RCM273_OnlineBookingNo"))=True then
             Response.Redirect "webstep2.asp"
      end if

         CompanyCode="PtWestCoastCampers273"
         RCMURL="https://secure2.rentalcarmanager.com"


          imageURL = "http://secure2.rentalcarmanager.com"
        Session("RCM273_CarImageName")=Replace(Session("RCM273_CarImageName"), "." , "&#46;")
         theimage="http://secure2.rentalcarmanager.com/DB/"&CompanyCode&"/"&Session("RCM273_CarImageName")
         ServerID=2
         RCMCompanyID=273




         logo="logo.gif"
         FilePath=Server.MapPath("/DB")&"\"&CompanyCode
          Set FileObject=Server.CreateObject("Scripting.FileSystemObject")



   '----------connect RCM Database---------------------

       Dim webConn
         Set webConn = Server.CreateObject("ADODB.Connection")
         DatabaseFile="PtWestCoastCampers273"
         webconn.Open "Provider=SQLOLEDB;Data Source = 4461QVIRT;Initial Catalog = "&DatabaseFile&";Trusted_Connection=yes;"

              '--Feb 2015, we get the booking details from web log, in case if user open other window for different dates and vehicle, get new set of session it will mix up with the web log booking..
                   Set s_w=WebConn.Execute("select * from WebReservation where  ReservationNo='"&Clng(Session("RCM273_OnlineBookingNo"))&"' ")
                  if Not s_w.EOF then
                     Session("RCM273_CarSizeID") =   s_w("CarSizeID")
                     Session("RCM273_PickupLocationID") = s_w("PickupLocationID")
                     Session("RCM273_DropoffLocationID") = s_w("DropoffLocationID")
                     Session("RCM273_RequestPickDate") = Day(s_w("PickupDate"))&"/"&MonthName(Month(s_w("PickupDate")),3)&"/"&Year(s_w("PickupDate"))
                     Session("RCM273_RequestPickTime") = s_w("PickupTime")
                     Session("RCM273_RequestPickDateTime") = Session("RCM273_RequestPickDate")&" "&Session("RCM273_RequestPickTime")
                     Session("RCM273_RequestDropDate") = Day(s_w("DropoffDate"))&"/"&MonthName(Month(s_w("DropoffDate")),3)&"/"&Year(s_w("DropoffDate"))
                     Session("RCM273_RequestDropTime") = s_w("DropoffTime")
                     Session("RCM273_RequestDropDateTime") = Session("RCM273_RequestDropDate")&" "&Session("RCM273_RequestDropTime")

                  end if
                   s_w.Close
                   Set s_w=nothing

                            GSTInclusive = "Yes"
                  Set s_st=WebConn.Execute("SELECT * FROM SystemTable WHERE Code='GSINC'  ")
                  If not s_st.EOF then
                        GSTInclusive = s_st("syValue")
                 END IF
                  s_st.CLOSE
                 SET s_st=NOTHING


         if Request.QueryString("type")="Quote"  then
                   Session("RCM273_Notes")=left(tidyup(Request.Form("Notes")),250)
                   Session("RCM273_FirstName")=left(tidyup(Request.Form("firstname")),30)
                   Session("RCM273_LastName")=left(tidyup(Request.Form("lastname")),30)
                   Session("RCM273_countryID")=tidyup(Request.Form("countryID"))
                   Session("RCM273_CustomerEmail")=left(tidyup(Request.Form("CustomerEmail")),50)
                   Session("RCM273_phone")=left(tidyup(Request.Form("phone")),20)
                   Session("RCM273_NoTravelling")=left(tidyup(Request.Form("NoTravelling")),30)


                   Quotation=1
                   ReservationTypeID=3  '--for quotation
                   onlineType="Online Quotation"

         else


                    Quotation=0
                    if Session("RCM273_categoryStatus") = "2" then '---LIMITED AVAILABILTY

                           FreeSale=0
                           'ReservationTypeID=0 '--need release the on request menu in RCM
                           ReservationTypeID=1
                           onlineType="Online Booking Request" '---Might be available (place a request)

                     else
                           FreeSale=1
                           ReservationTypeID=1  '--confirmed booking
                           onlineType="Online Booking"

                     end if

                     DOBYear = CInt(Request.Form("DOBYear"))
                     DOBMonth = (Request.Form("DOBMonth"))
                     DOBDay  = CInt(Request.Form("DOBDay"))
                     Session("RCM273_DOB")= DOBDay&"/"&DOBMonth&"/"&DOBYear
                     if Day(Session("RCM273_RequestPickDate"))=29 and MonthName(Month(Session("RCM273_RequestPickDate")),3)="Feb"  then
                           MinDOB= MonthName(Month(Session("RCM273_RequestPickDate")),3)&"/"&(Day(Session("RCM273_RequestPickDate"))-1)&"/"&(Year(Session("RCM273_RequestPickDate"))-Session("RCM273_MinimunAge"))
                     else
                           MinDOB= MonthName(Month(Session("RCM273_RequestPickDate")),3)&"/"&Day(Session("RCM273_RequestPickDate"))&"/"&(Year(Session("RCM273_RequestPickDate"))-Session("RCM273_MinimunAge"))
                     end if
                     if Day(Session("RCM273_RequestPickDate"))=29 and MonthName(Month(Session("RCM273_RequestPickDate")),3)="Feb"  then
                           MinDOB= "28/"&MonthName(Month(Session("RCM273_RequestPickDate")),3)&"/"&(Year(Session("RCM273_RequestPickDate"))-Session("RCM273_MinimunAge"))
                     end if
                     '----check the Expiry Date
                     LicYear = CInt(Request.Form("licensYear"))
                     LicMonth = tidyup(Request.Form("licensMonth"))
                     LicDay  = CInt(Request.Form("licensDay"))
                     Session("RCM273_LicExpDate") = LicDay&"/"&LicMonth&"/"&LicYear
                     'CurrentExpDate= Day(Date)&"/"&MonthName(Month(Date),3)&"/"&Year(Date)
                      CurrentExpDate= Day(Session("RCM273_RequestDropDate"))&"/"&MonthName(Month(Session("RCM273_RequestDropDate")),3)&"/"&Year(Session("RCM273_RequestDropDate"))
                      Session("RCM273_FirstName")=Left(tidyup(Request.Form("firstname")),30)
                      Session("RCM273_LastName")=Left(tidyup(Request.Form("lastname")),30)
                      Session("RCM273_License")=Left(tidyup(Request.Form("License")),70)
                      Session("RCM273_LicenseIssued")=Left(tidyup(Request.Form("LicenseIssued")),30)
                      Session("RCM273_address")=Left(tidyup(Request.Form("address")),80)
                      Session("RCM273_city")=Left(tidyup(Request.Form("city")),50)
                     Session("RCM273_state")=Left(tidyup(Request.Form("state")),30)
                     Session("RCM273_postcode")=Left(tidyup(Request.Form("postcode")),10)
                     Session("RCM273_countryID")=tidyup(Request.Form("countryID"))
                     Session("RCM273_CustomerEmail")=Left(tidyup(Request.Form("CustomerEmail")),50)
                     Session("RCM273_phone")=Left(tidyup(Request.Form("phone")),20)
                     Session("RCM273_fax")=Left(tidyup(Request.Form("fax")),15)
                     Session("RCM273_Notes")=left(tidyup(Request.Form("Notes")),250)
                     Session("RCM273_Flight")=Left(tidyup(Request.Form("Flight")),50)
                     Session("RCM273_Flightout")=Left(tidyup(Request.Form("Flightout")),50)
                     Session("RCM273_NoTravelling")=Left(tidyup(Request.Form("NoTravelling")),30)
                     Session("RCM273_CollectionPoint") =Left(tidyup(Request.Form("CollectionPoint")),80)
                     Session("RCM273_ReturnPoint") =Left(tidyup(Request.Form("ReturnPoint")),120)
                     Session("RCM273_transmission")=Left(tidyup(Request.Form("transmission")),6)

                       if Session("RCM273_LastName")="" or Session("RCM273_CustomerEmail")="" then
                            Session("RCM273_ErrorMesage")="* Please enter Your Details."
                       elseif  Request.Form("DOBYear")="1900" then
                              Session("RCM273_ErrorMesage")="*  Please enter your DOB."
                              Response.Redirect "webstep4.asp?type=Booking"

                       elseif   IsDate(Session("RCM273_DOB"))<>"True"  then
                              Session("RCM273_ErrorMesage")="* The DOB date "&Session("RCM273_DOB")&" does not exit."
                              Response.Redirect "webstep4.asp?type=Booking"
                       elseif  IsDate(Session("RCM273_LicExpDate")) <>"True"  then
                              Session("RCM273_ErrorMesage")="* The License Expiry Date "&Session("RCM273_LicExpDate")&" does not exit."
                              Response.Redirect "webstep4.asp?type=Booking"
                        elseif  Session("RCM273_License")<>"" and  (DateValue(Session("RCM273_LicExpDate")) - DateValue(CurrentExpDate)< 0) then
                              Session("RCM273_ErrorMesage")="* The License Date "&Session("RCM273_LicExpDate")&" is Expired."
                              'Response.Write("<p><center><span class=red>&nbsp;&nbsp;&nbsp;* The License Date "&Session("RCM273_LicExpDate")&" is Expired.")
                              Response.Redirect "webstep4.asp?type=Booking"
                        elseif   (Session("RCM273_MinimunAge")<>0 and (DateValue(MinDOB) - DateValue(Session("RCM273_DOB")))< 0 ) then
                              Session("RCM273_ErrorMesage")="* Vehicles are not available for hire to drivers under "&Session("RCM273_MinimunAge")&" years of age."
                              'Response.Write("<p><center><span class=red>&nbsp;&nbsp;&nbsp;* Vehicles are not available for hire to drivers under "&Session("RCM273_MinimunAge")&" years of age.")
                              Response.Redirect "webstep4.asp?type=Booking"
                        else

                              Set s_c=webConn.Execute("SELECT * FROM Country where (ID)='"&Request.Form("countryID")&"' " )
                               Session("RCM273_country")=s_c("Country")
                               s_c.close
                               Set s_c=nothing


                        end if

           end if

                if Request.Form("RentalSource")="" then
                     Session("RCM273_RentalSource")="Online Booking"
                else
                     Session("RCM273_RentalSource")=Left(tidyup(Request.Form("RentalSource")),40)
                end if
                 if tidyup(Session("RCM273_RentalSource"))="" then
                        Session("RCM273_RentalSource")="Online Booking"
                 end if


               if Session("RCM273_DOB")="" then
                     Session("RCM273_DOB") ="1/Jan/1900"
               end if
               if isDate(Session("RCM273_LicExpDate"))<>True then
                  Session("RCM273_LicExpDate")="31/Dec/2100"
               end if
               if Session("RCM273_CountryID")="" then
                  Session("RCM273_CountryID")=7
               end if

         if Session("RCM273_CustomerID")=0 or Session("RCM273_CustomerID")="" then
                 SQL="INSERT INTO customers (CompanyID,acPostalAddress,acCity,acState,acPostcode,acCountryID,acFax,DoNotRentID,acDOB,acLicenseExpiry,acLicense,acLicenseIssued,acFirstName, acLastName, acEmail, acPhone)"
                SQL=SQL&"VALUES (0,'"&Left(tidyup(Session("RCM273_address")),80)&"','"&Left(tidyup(Session("RCM273_city")),50)&"','"&Left(tidyup(Session("RCM273_state")),30)&"','"&Left(tidyup(Session("RCM273_postcode")),10)&"','"&Session("RCM273_CountryID")&"','"&Left(tidyup(Session("RCM273_fax")),15)&"',0,'"&Session("RCM273_DOB")&"','"&Session("RCM273_LicExpDate")&"','"&Left(Tidyup(Session("RCM273_License")),70)&"','"&Left(Tidyup(Session("RCM273_LicenseIssued")),30)&"','"&Left(Tidyup(Session("RCM273_FirstName")),30)&"','"&Left(Tidyup(Session("RCM273_LastName")),30)&"','"&Left(Tidyup(Session("RCM273_CustomerEmail")),50)&"','"&Left(Tidyup(Session("RCM273_phone")),20)&"')"
                'Response.Write(SQL)
                webConn.Execute(SQL)

             '---get customerID
                  Set s_cus=webConn.Execute("SELECT Max(acID)  as customID FROM Customers where acLastName ='"&Session("RCM273_LastName")&"' ")
               CustomerID=s_cus("customID")
                s_cus.close
                set s_cus=nothing
       else
    
                CustomerID=Session("RCM273_CustomerID")
                SQL="update customers SET "
                SQL=SQL&" acPostalAddress='"&Left(tidyup(Session("RCM273_address")),80)&"', "
                SQL=SQL&" acCity='"&Left(tidyup(Session("RCM273_city")),50)&"', "
                SQL=SQL&" acState='"&Left(tidyup(Session("RCM273_state")),30)&"', "
                 SQL=SQL&" acPostcode='"&Left(tidyup(Session("RCM273_postcode")),20)&"', "
                 SQL=SQL&" acFax='"&Left(tidyup(Session("RCM273_fax")),15)&"', "
                 SQL=SQL&" acFirstName='"&Left(tidyup(Session("RCM273_FirstName")),30)&"', "
                SQL=SQL&" acLastName='"&Left(tidyup(Session("RCM273_LastName")),40)&"', "
                 SQL=SQL&" acEmail='"&Left(tidyup(Session("RCM273_CustomerEmail")),50)&"', "
                SQL=SQL&" acPhone='"&Left(tidyup(Session("RCM273_phone")),20)&"', "
                 if Request.QueryString("type")<>"Quote" then
                  SQL=SQL&" acDOB='"&Session("RCM273_DOB")&"', "
                  SQL=SQL&" acLicenseExpiry='"&tidyup(Session("RCM273_LicExpDate"))&"', "
                  SQL=SQL&" acLicense='"&Session("RCM273_License")&"', "
                  SQL=SQL&" acLicenseIssued='"&Left(Tidyup(Session("RCM273_LicenseIssued")),30)&"', "

                end if
                SQL=SQL&" acCountryID='"&Session("RCM273_CountryID")&"' "
                SQL=SQL&" WHERE acID='"&Session("RCM273_CustomerID")&"' "
               ' Response.Write(SQL)
             webConn.Execute(SQL)

         end if


   
if Session("RCM273_BookingBufferNo")=0 or Session("RCM273_BookingBufferNo")="" then
      
         '---Insert the booking details to ReservationBuffer table
            BrandID=0
             if webURLD<>"" then
            BrandID=webURLD
            end if

           Set s_p=webConn.Execute("SELECT lc_BrandID   FROM LocationCompanies where lc_LocationID ='"&CINT(Session("RCM273_PickupLocationID"))&"'  and lc_URLD= '"&webURLD&"' ")
            if NOT s_p.EOF then
               BrandID=s_p("lc_BrandID")
           end if
           s_p.close
           set s_p=nothing

            RcmReferralBooking=0
            if Session("RCM273_referral")="" then
               Session("RCM273_referral")=0
            end if
            if Session("RCM273_referral")<>0 then
               RcmReferralBooking=1
            end if


                  '--------------- Referral Commintion---------------
                  '---need add Session("RCM273_RentalCost") to step4
              if   Session("RCM273_referral")<>0  then
                  Set s_extra=webConn.Execute("SELECT Commission FROM Referrals WHERE  ID='"&Session("RCM273_referral")&"'  ")
                  EachExtraFees=0
                  if Not s_extra.EOF then
                          Session("RCM273_referralCommission")=Session("RCM273_RentalCost") * s_extra("Commission")/100
                          end if
               s_extra.Close
               Set s_extra=nothing
            end if


             DateEntered = Session("RCM273_LocalTime")
              if Session("RCM104_LocalTime")="" then
             DateEntered = Day(Now)&"/"&MonthName(Month(Now))&"/"&Year(Now)
             end if
              SQL="INSERT INTO ReservationBuffer (CampaignCode,Transmission,AreaofUsed,RcmReferralBooking,ReferralID,ReferralCommission, driverage,DateEntered,UpdatedDate,EnteredBy,LastUpdatedBy,BrandID,Quotation,AddKmsFee,KmsFree,NoTravelling,Flightout,Flight,CollectionPoint,ReturnPoint,StampDuty,GST,PickupDateTime,DropoffDateTime,CustomerID,RentalSource,CarSizeID,PickupTime,Pickupdate,DropoffTime,DropoffDate,PickupLocationID,DropoffLocationID,Notes,ReservationTypeID)"
              SQL=SQL&"VALUES ('"&Left(Session("RCM273_PromoCode"),30)&"','"&Session("RCM273_transmission")&"', '"&Left(Session("RCM273_AreaofUsed"),40)&"','"&RcmReferralBooking&"','"&Session("RCM273_referral")&"','"&Session("RCM273_referralCommission")&"', '"&Session("RCM273_driverage")&"','"&DateEntered&"','"&DateEntered&"', 2,2,'"&BrandID&"',"&Quotation&","&Session("RCM273_AddKmsFee")&","&Session("RCM273_KmsFree")&",'"&Left(Session("RCM273_NoTravelling"),20)&"','"&Left(Session("RCM273_Flightout"),50)&"','"&Left(Session("RCM273_Flight"),50)&"','"&Left(Session("RCM273_CollectionPoint"),80)&"','"&Left(Session("RCM273_ReturnPoint"),80)&"','"&Session("RCM273_StampDuty")&"','"&Session("RCM273_GST")&"','"&Session("RCM273_RequestPickDateTime")&"','"&Session("RCM273_RequestDropDateTime")&"' ,"&(CustomerID)&",'"&Left(tidyup(Session("RCM273_RentalSource")),40)&"',"&(CInt(Session("RCM273_CarSizeID")))&",'"&Session("RCM273_RequestPickTime")&"','"&Session("RCM273_RequestPickDate")&"','"&Session("RCM273_RequestDropTime")&"', '"&Session("RCM273_RequestDropDate")&"', "&CINT(Session("RCM273_PickupLocationID"))&", "&Session("RCM273_DropoffLocationID")&", '"&left(tidyup(Session("RCM273_Notes")),250)&"', '1')"
            'Response.Write(SQL)
              webConn.Execute(SQL)

        '---get the ReservationNo
           Set s_No=webConn.Execute("SELECT Max(ReservationNo)  as ResNo FROM ReservationBuffer where  CustomerID='"&CustomerID&"' ")
           'OnlineBookingNo=s_No("ResNo")
           Session("RCM273_BookingBufferNo")=s_No("ResNo")
           s_No.close
           set s_No=nothing

        '---Apr 2014, rcm changes, save the referenceID and pass to the link
             Function AlmostUniqueID()
                      Randomize
                      for iCtr = 1 to 10
                               sChar = Chr(Int((90 - 65 + 1) * Rnd) + 65)
                               sID = sID & sChar
                      Next
                      sID = sID & Month(Now) & Day(Now) & Year(Now) & Hour(Now) _
                                  & Minute(Now) & Second(Now)
                      AlmostUniqueID = sID
             End Function
             iUniqueID=AlmostUniqueID()
             rcmReferenceID=CustomerID&""&iUniqueID
             rcmReferenceID=Left(rcmReferenceID,30)

         '--insert Unallocaited RA# to WebReservaiton table,
            SQL="UPDATE ReservationBuffer SET "
            SQL=SQL&"rcmReferenceID ='"&rcmReferenceID&"' "
            SQL=SQL&"WHERE (ReservationNo) ='"&Session("RCM273_BookingBufferNo")&"'"
           'Response.Write(SQL)
            webConn.Execute(SQL)




        '--insert Unallocaited RA# to WebReservaiton table,

              SQL="UPDATE WebReservation SET "
             'SQL=SQL&" RentalSource ='"&tidyup(Session("RCM273_RentalSource"))&"',  "
              SQL=SQL&" KmsFree ="&Session("RCM273_KmsFree")&",  "
              SQL=SQL&" AddKmsFee ="&Session("RCM273_AddKmsFee")&",  "
              SQL=SQL&" Phone ='"&Left(tidyup(Request.Form("phone")),20)&"' , "
              SQL=SQL&" Email ='"&Left(tidyup(Request.Form("CustomerEmail")),50)&"',  "
              SQL=SQL&" Name ='"&Left(tidyup(Request.Form("lastname")),30)&"',  "
               SQL=SQL&" StampDuty ="&Session("RCM273_StampDuty")&",  "
               SQL=SQL&" GST ="&Session("RCM273_GST")&",  "
               'SQL=SQL&" BookingType ="&ReservationTypeID&",  "  '---Nov 2012, we should change it as ReservationTypeID=1 for cofirmed booking, ReservationTypeID=0 for request, no need the bookingType ,
               SQL=SQL&" ReservationTypeID ="&ReservationTypeID&",  "
               SQL=SQL&" UnallocatedRA ='"&Session("RCM273_BookingBufferNo")&"'  "
               SQL=SQL&"WHERE (ReservationNo) ='"&Clng(Session("RCM273_OnlineBookingNo"))&"'"
              webConn.Execute(SQL)
              ' Response.Write(SQL)

      Set s_km=webConn.Execute("SELECT * FROM WebReservationFees WHERE (rf_ReservationNo)='"&Clng(Session("RCM273_OnlineBookingNo"))&"' ")
       if Not s_km.EOF   then
             webConn.Execute("DELETE  FROM ReservationFeesBuff WHERE rf_ReservationNo="&Clng(Session("RCM273_OnlineBookingNo"))&" ")
            SQL="INSERT INTO ReservationFeesBuff (rf_ReservationNo,rf_MaxKmscharge,rf_DailyRate)"
            SQL=SQL&"VALUES ("&Session("RCM273_BookingBufferNo")&","&s_km("rf_MaxKmscharge")&","&s_km("rf_DailyRate")&")"
            'Response.Write(SQL)
            webConn.Execute(SQL)
      end if
   s_km.Close
   Set s_km=nothing


  '--get each season record from the WebPaymentDetail and insert to paymentdetailes table
   webConn.Execute("DELETE FROM PaymentDetailBuffer WHERE (ReservationNo)='"&Session("RCM273_BookingBufferNo")&"' ")
   Set s_rate=webConn.Execute("SELECT * FROM WebPaymentDetail WHERE (ReservationNo)='"&Clng(Session("RCM273_OnlineBookingNo"))&"' ")
   RentalCost=0
   While Not s_rate.EOF
             RentalCost=s_rate("Days")* s_rate("Rate") + RentalCost
            SQL="INSERT INTO PaymentDetailBuffer (NoHours,TripRates,DiscountID,DiscountType,DiscountName,DiscountPerc,StandardRate,Rate,ReservationNo,SeasonID,Days,RateName)"
             SQL=SQL&"VALUES ("&s_rate("NoHours")&", 0,'"&Session("RCM273_DiscountID")&"','"&Session("RCM273_DiscountType")&"','"&tidyup(Session("RCM273_DiscountName"))&"',"&Session("RCM273_DiscountRate")&","&s_rate("StandardRate")&","&s_rate("Rate")&",'"&Session("RCM273_BookingBufferNo")&"',"&s_rate("SeasonID")&","&s_rate("Days")&",'"&s_rate("RateName")&"')"
             webConn.Execute(SQL)
   s_rate.MoveNext
   Wend
   s_rate.Close
   Set l_s=nothing

    TotalCharges=RentalCost
    AgencyCommintionValue=RentalCost
     '-Insert Extra Fees (Mandatory and Selected)  to PaymentExtraFeesBuffer table
       webConn.Execute("DELETE FROM PaymentExtraFeesBuffer WHERE (ReservationNo)='"&Session("RCM273_BookingBufferNo")&"' ")
       'Set s_rate=webConn.Execute("SELECT * FROM WebPaymentExtraFees WHERE (ReservationNo)='"&Clng(Session("RCM273_OnlineBookingNo"))&"' ")
       Set s_rate=webConn.Execute("SELECT WebPaymentExtraFees.*, Name, Type,PayAgency FROM WebPaymentExtraFees, ExtraFees WHERE  (ReservationNo)='"&Clng(Session("RCM273_OnlineBookingNo"))&"' and ExtraFeesID = ExtraFees.ID ORDER BY TYPE, NAME ")
      While Not s_rate.EOF
         
               TotalCharges=TotalCharges+ s_rate("ExtraValue")
               if s_rate("PayAgency")=True then
                     AgencyCommintionValue=AgencyCommintionValue+ s_rate("ExtraValue")
               end if 
               SetMaxPrice=0
               if    s_rate("SetMaxPrice")=True then
                       SetMaxPrice=1
               end if
               SQL="INSERT INTO PaymentExtraFeesBuffer(ExtraValue,SetMaxPrice,QTY,ReservationNo,ExtraFeesID,Fees,Days)" 
               SQL=SQL&"VALUES ("&s_rate("ExtraValue")&","&SetMaxPrice&","&s_rate("QTY")&",'"&Session("RCM273_BookingBufferNo")&"',"&s_rate("ExtraFeesID")&","&s_rate("Fees")&","&s_rate("Days")&")"
                webConn.Execute(SQL)
                'Response.Write("<br>") 
               'Response.Write(SQL)
   s_rate.MoveNext 
   Wend
   s_rate.Close
   Set s_rate=nothing       
 
         
end if '---if Session("RCM273_BookingBufferNo")=0 or Session("RCM273_BookingBufferNo")=""

       DPS_PaymentAmount=0
      '--confirmed unallocated booking, use payment express-----------
       if   Session("RCM273_useDPSpayment")="Yes" and Request.Form("TxnType")<>""  then
                  '--1. confirmed unallocated booking pay 20% of booking value
                  '--2. if card is incorrect or expiry is incorrect it will allow re-entry
                  '--3. if card declines (eg insufficient funds ) turn into quote and send to Jucy?

                          Username="GMHDev" '---for testing, changed password 22 Feb 2011
                          Password="test1234"
                         Username="GMH_Dev" '---for testing, from wicked
                         Password="395c304b"

                         ExpiryDate=Request.form("ccDate1")&""&request.form("ccDate2")
                         DPS_ReservationBufferNo=Session("RCM273_BookingBufferNo")
                         DPS_ReservationNo=0
                         DPS_MerchantReference="U"&Session("RCM273_BookingBufferNo")
                         if Request.Form("PaymentOptions")="1" then
                              AmountInput=0.2*Session("RCM273_TotalEstimateofCharges")
                         elseif Request.Form("PaymentOptions")="2" then
                              AmountInput=Session("RCM273_TotalEstimateofCharges")
                         end if

                        AmountInput=Round(AmountInput,2)
                         AmountInput=Formatnumber(AmountInput,2)

                        AmountInput=Replace(AmountInput, ",", "")
                        PaymentCurrency="AUD"

                        sXmlAction = sXmlAction & "<Txn><PostUsername>"&Username&"</PostUsername>"
                        sXmlAction = sXmlAction & "<PostPassword>"&Password&"</PostPassword>"
                        sXmlAction = sXmlAction & "<TxnType>"&Request.Form("TxnType")&"</TxnType>"
                        sXmlAction = sXmlAction & "<CardHolderName>"&Request.Form("ccName")&"</CardHolderName>"
                        sXmlAction = sXmlAction & "<CardNumber>"&Request.Form("ccNumber")&"</CardNumber>"
                        sXmlAction = sXmlAction & "<Amount>"&AmountInput&"</Amount>"
                        sXmlAction = sXmlAction & "<InputCurrency>"&PaymentCurrency&"</InputCurrency>"
                        sXmlAction = sXmlAction & "<DateExpiry>"&ExpiryDate&"</DateExpiry>"
                        sXmlAction = sXmlAction & "<MerchantReference>"&DPS_MerchantReference&"</MerchantReference>"
                        sXmlAction = sXmlAction & "<ReceiptEmail>"&Session("RCM273_CustomerEmail")&"</ReceiptEmail></Txn>"

                     ' response.write(sXmlAction)

                       Dim objXMLhttp
                       Set objXMLhttp = server.Createobject("MSXML2.XMLHTTP")
                       Set objOutputXMLDoc = Server.CreateObject("Microsoft.XMLDOM")


                        objXMLhttp.Open "POST", "https://sec.paymentexpress.com/pxpost.aspx" ,False
                        objXMLhttp.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
                        objXMLhttp.send sXmlAction


                        'Load the XML response text onto a DOM dcoument
                        objOutputXMLDoc.loadXML objxmlhttp.responsetext
                         'response.write objxmlhttp.responsetext
                        'response.write "<p>"

                        'Check to make sure the load worked correctly
                        if objOutputXMLDoc.parseerror<>0 then
                              response.write "<bR><b>Parse Error:</b>" & objOutputXMLDoc.parseerror &objOutputXMLDoc.parseerror.reason & objOutputXMLDoc.parseerror.line
                              response.end
                              Session("RCM273_DPSErrorMesage")="You have entered an invalid card number, please check the details entered."
                              Response.Redirect "webStep4.asp?type=Booking"
                        else

                             ' check to make sure the element you are checking is found
                              if objOutputXMLDoc.getElementsByTagName("Authorized").length>0 then
                                 set AuthorizedElement =objOutputXMLDoc.getElementsByTagName("Authorized").item(0)
                                 strAuthorised = AuthorizedElement.text
                                 'response.write "The Authorised Element is : " & strAuthorised
                              End if

                              set MerchantReferenceElement =objOutputXMLDoc.getElementsByTagName("MerchantReference").item(0)
                              strMerchantReference = tidyup(MerchantReferenceElement.text)

                              set CardHolderNameElement =objOutputXMLDoc.getElementsByTagName("CardHolderName").item(0)
                              strCardHolderName = Left(tidyup(CardHolderNameElement.text),50)

                              set CardNumberElement =objOutputXMLDoc.getElementsByTagName("CardNumber").item(0)
                              strCardNumber = tidyup(CardNumberElement.text)

                              set DateExpiryElement =objOutputXMLDoc.getElementsByTagName("DateExpiry").item(0)
                              strDateExpiry = DateExpiryElement.text

                              set AmountElement =objOutputXMLDoc.getElementsByTagName("Amount").item(0)
                              strPaymentAmount = AmountElement.text

                              set CurrencyNameElement =objOutputXMLDoc.getElementsByTagName("CurrencyName").item(0)
                              strCurrencyName = tidyup(CurrencyNameElement.text)

                              set DateSettlementElement =objOutputXMLDoc.getElementsByTagName("DateSettlement").item(0)
                              strDateSettlement = tidyup(DateSettlementElement.text)

                              set DpsTxnRefElement =objOutputXMLDoc.getElementsByTagName("DpsTxnRef").item(0)
                              strDpsTxnRef = tidyup(DpsTxnRefElement.text)

                              set CardHolderResponseTextElement =objOutputXMLDoc.getElementsByTagName("CardHolderResponseText").item(0)
                              strCardHolderResponseText = tidyup(CardHolderResponseTextElement.text)

                              set MerchantResponseTextElement =objOutputXMLDoc.getElementsByTagName("MerchantResponseText").item(0)
                              strMerchantResponseText = tidyup(MerchantResponseTextElement.text)
                              Session("RCM273_DPSErrorMesage")=""
                              DPSErrorMesage=""
                              if strAuthorised="0"  then

                                       Session("RCM273_DPSErrorMesage")=strMerchantResponseText
                                       Session("RCM273_DPSErrorMesage2")=AmountInput
                                       Response.Redirect "webStep4.asp?type=Booking"
                              end if

                              SQL="INSERT INTO DPSPayment (DPS_Authorized,DPS_MerchantReference, DPS_CardHolderName, DPS_CardNumber,"
                              SQL=SQL&" DPS_DateExpiry, DPS_Amount,   DPS_CurrencyName,"
                              SQL=SQL&" DPS_DateSettlement,   DPS_DpsTxnRef,  DPS_CardHolderResponseText,"
                              SQL=SQL&" DPS_MerchantResponseText,   DPS_ReservationBufferNo,  DPS_ReservationNo)"
                              SQL=SQL&"VALUES ('"&strAuthorised&"','"&strMerchantReference&"','"&strCardHolderName&"', '"&strCardNumber&"', "
                              SQL=SQL&" '"&strDateExpiry&"','"&strPaymentAmount&"','"&strCurrencyName&"', "
                              SQL=SQL&" '"&strDateSettlement&"','"&strDpsTxnRef&"','"&strCardHolderResponseText&"', "
                              SQL=SQL&" '"&strMerchantResponseText&"','"&DPS_ReservationBufferNo&"','"&DPS_ReservationNo&"') "
                              'Response.Write("<p>")
                              'Response.Write(SQL)
                              webConn.Execute(SQL)

                              ' --can use above to diplay on tht screen
                              '--<CardHolderHelpText>INVALID AMOUNT
                              if strAuthorised="1" then '-- if transaction successful


                                     'response.write("<p>Success")
                                    '---insert payment to RCM

                                    if IsNumeric(strPaymentAmount)=True then
                                          DPS_PaymentAmount=CCur(strPaymentAmount)
                                    end if

                                    '--payment date should be the DateSettlement?
                                    PaymentDate=Day(Now)&"/"&MonthName(Month(Now))&"/"&Year(Now)
                                    BillingLocationID=1 '--brisbane for jucy
                                    LoginId=2 '--online booking login

                                     webConn.Execute("DELETE FROM ReservationPaymentBuffer WHERE (ReservationNo)='"&Session("RCM273_BookingBufferNo")&"' ")
                                     SQL="INSERT INTO ReservationPaymentBuffer(BillingLocationID,ReservationNo,PaymentType,Paid,PaymentDate,DMYCreated,LoginId)"
                                     SQL=SQL&"VALUES ("&BillingLocationID&","&Session("RCM273_BookingBufferNo")&",'"&ccType&"','"&DPS_PaymentAmount&"','"&PaymentDate&"','"&PaymentDate&"','"&LoginId&"')"
                                     ' Response.Write("<p>")
                                     'Response.Write(SQL)
                                      webConn.Execute(SQL)

                                     '--25 Oct 2009, for online full payment get $2 per day discount (ExtraFeeID= 227)

                                   ' DepostOptionAmount=0.2*Session("RCM273_TotalEstimateofCharges")
                                    'if Request.Form("AmountInput")="FullPayment" then    'Response.Write("===test")

                                       ' Set s_ds=webConn.Execute("SELECT Fees FROM  ExtraFees WHERE  ID=227 ")
                                       'if NOT s_ds.EOF then
                                           '  Fees=s_ds("Fees")
                                            ' ExtraValue=Session("RCM273_TotalRentalDays")*s_ds("Fees")
                                           '  ExtraFeeID= 227

                                           '  SQL="INSERT INTO PaymentExtraFeesBuffer(ExtraValue,SetMaxPrice,QTY,ReservationNo,ExtraFeesID,Fees,Days)"
                                           '  SQL=SQL&"VALUES ("&ExtraValue&",0,1,'"&Session("RCM273_BookingBufferNo")&"',"&ExtraFeeID&","&Fees&","&Session("RCM273_TotalRentalDays")&")"
                                           '  webConn.Execute(SQL)
                                             'Response.Write("<br>")
                                             'Response.Write(SQL)
                                      ' end if
                                      ' s_ds.close
                                       'set  s_ds=nothing

                                ' end if

                            end if

                  end if
                  Set objXMLhttp = nothing

            END IF '--END if   Session("RCM273_useDPSpayment")="Yes" and Request.Form("TxnType")<>""  then


       BookingNo = Session("RCM273_BookingBufferNo")
          SQL="SELECT QReservationforBuffer.*, TotalCost = dbo.BookingValueBuffer(ReservationNo, '"&GSTInclusive&"'), "
        SQL=SQL&" CASE WHEN dbo.CategoryPerNightRate(CarSizeID)=0 THEN 'day' ELSE 'night' END as sRateTypeDesc, "
         SQL=SQL&" dbo.BookingBalanceBuffer(ReservationNo,'"&GSTInclusive&"')  AS BalanceDue "
         SQL=SQL&" from QReservationforBuffer where ReservationNo="&CLng(BookingNo)&"  "
      Set s_d=webConn.Execute(SQL)
      If not s_d.EOF then


                   Set s_sy=webConn.Execute("SELECT * FROM SystemTable WHERE Code='CURCY'  ")
                  CurrencySample="$"
                  If not s_sy.EOF then
                        CurrencySample= s_sy("syValue")
                  END IF
                 s_sy.CLOSE
                  SET s_sy=NOTHING

                  Function DisplayCurrency(n)
                    DisplayCurrency = CurrencySample&""&FormatNumber(n,2)
                  End Function



               '-----use company info

                 if s_d("BrandID")<>0 then
                     Set s_com=webConn.Execute("SELECT * from Company WHERE ID="&s_d("BrandID")&"   ")
                 else
                     Set s_com=webConn.Execute("SELECT * from Company WHERE Defaulted=1  ")
                 end if
                 if s_com("BrandLogo")<>"" then
                     logo=s_com("BrandLogo")
                 end if
                 Company=s_com("Company")
                 CompanyEmail=s_com("Email")
                 BranchAddress=s_com("Address")
                 BranchCity = s_com("Suburb")
                 BranchPostCode = s_com("State")&"&nbsp;"&s_com("PostCode")&"&nbsp;"&s_com("Country")
                 BranchPhone=s_com("Phone")
                 BranchFax=s_com("Fax")
                 if s_com("WebSite")<>"" then
                     WebSite="<a href=http://"&s_com("WebSite")&">"&s_com("WebSite")&"</a>"
                 end if
                     FreePhone=""
                 if s_com("FreePhone")<>"" then
                     FreePhone="&nbsp;&nbsp;Free Phone: "&s_com("FreePhone")&""
                 end if
                  s_com.close
                  set  s_com=nothing

                 '--check use location or company address---
                Set RG=webConn.Execute("SELECT * FROM SystemTable WHERE Code='DOCS'  ")
                DOCSAddress="Yes"   '---Location Address
                If not RG.EOF then
                     DOCSAddress= RG("syValue")
                END IF
                RG.CLOSE
                SET RG=NOTHING

               if DOCSAddress="Yes" then '--use Location address
                   Set s_add=webConn.Execute("SELECT  QLocationState.* FROM QLocationState WHERE ID='"&CINT(Session("RCM273_PickupLocationID"))&"'  ")
                   if NOT s_add.EOF  then
                               if s_add("Email")<>"" then
                                     CompanyEmail=s_add("Email")
                               end if
                               BranchAddress = s_add("Address")
                               BranchCity = s_add("Suburb")&"&nbsp;"&s_add("City")
                               if s_add("StateCode")<>"N/A" then
                                  BranchPostCode = s_add("StateCode")&"&nbsp;"&s_add("PostCode")&"&nbsp;"&s_add("Country")
                               else
                                  BranchPostCode = s_add("PostCode")&"&nbsp;"&s_add("Country")
                               end if
                               BranchPhone = s_add("Phone")
                               BranchFax = s_add("Fax")
                      end if
                      s_add.close
                      Set s_add=nothing
               end if
           '-----Send a confirmation to customer

             if Request.QueryString("type")="Quote" then
                Subject1="Online Quotation with "&Company&" -  Ref #Q-"&Session("RCM273_DocPrFix")&""&Session("RCM273_BookingBufferNo")&" ("&Session("RCM273_PickupLocation")&")"

             else
                 if Session("RCM273_categoryStatus") = "2" then '---LIMITED AVAILABILTY
                   Subject1="Online Booking Request with "&Company&" - Ref #U-"&Session("RCM273_DocPrFix")&""&Session("RCM273_BookingBufferNo")&" ("&Session("RCM273_PickupLocation")&")"
                else
                   Subject1="Online Booking with "&Company&" - Ref #U-"&Session("RCM273_DocPrFix")&""&Session("RCM273_BookingBufferNo")&" ("&Session("RCM273_PickupLocation")&")"
                end if
            end if
          
              Dim  HTML
              Dim MyCDONTSMail

               Set MyCDONTSMail = CreateObject("CDONTS.NewMail")

               MyCDONTSMail.From = Company&"<"&CompanyEmail&">"
               MyCDONTSMail.To = Session("RCM273_CustomerEmail")
               MyCDONTSMail.Cc = Company&"<"&CompanyEmail&">"
               MyCDONTSMail.Subject=Subject1
               MyCDONTSMail.BodyFormat=0
               MyCDONTSMail.MailFormat=0
               MyCDONTSMail.Body=HTML
               if Request.QueryString("type")<>"Quote" then
                    MyCDONTSMail.Importance = 2
               end if

               HTML = "<html>"
               HTML = HTML & "<head>"
               HTML = HTML & "<META http-equiv=Content-Type content='text/html; charset=utf-8'>"
               HTML = HTML & "<STYLE>"
               HTML = HTML & ".text{  font-size: 12px; line-height: 15px; color:#010101;  font-family: Arial,Verdana;}   "
               HTML = HTML & ".greytext{ font-size: 12px; line-height: 15px; color:#010101;  font-family: Arial,Verdana;}  "
               HTML = HTML & ".Yellow{   font-size: 12px; line-height: 15px; color: #FFFFFF; font-weight: bold;  font-family: Arial,Verdana;}    "
               HTML = HTML & "A.linkBlue{  font-size: 12px;  color: blue;  font-weight: 900; font-family: Arial,Verdana;  text-decoration: none;}  "

               HTML = HTML & ".HighlightRow {background-color:#EEF1F4;} "
               HTML = HTML & ".TitleRow  {background-color:#376293;}"
               HTML = HTML & ".OpeningTD {     "
               HTML = HTML & "border:solid windowtext 1.0pt;   "
               HTML = HTML & "border-color:"&Session("RCM273_CompanyColour")&"; }  "
               HTML = HTML & "BODY   {      "
               HTML = HTML & "background-color:#FFFFFF;  "
               HTML = HTML & "scrollbar-3dlight-color:#000000; "
               HTML = HTML & "scrollbar-arrow-color:#4D6185;   "
               HTML = HTML & "scrollbar-base-color:#BDD1FB;"
               HTML = HTML & "scrollbar-darkshadow-color:#000000; "
               HTML = HTML & "scrollbar-face-color:#B4BDC5;   "
               HTML = HTML & "scrollbar-highlight-color:#ffffff;"
               HTML = HTML & "scrollbar-track-color:#BDD1FB; "
               HTML = HTML & "scrollbar-shadow-color:#BDD1FB;}"

               HTML = HTML & "</STYLE>"
               HTML = HTML & "<META content='MSHTML 6.00.2900.2722' name=GENERATOR></HEAD>"
               HTML = HTML & "<title>"&Subject1&" </title>"
               HTML = HTML & "</head>"
               HTML = HTML & "<body bgcolor='#FFFFFF'>"

               HTML = HTML & "<center><img  src='"&imageURL&"/db/"&CompanyCode&"/"&logo&"' /></center>"

               HTML = HTML & "<table style='border-collapse:collapse;  border:solid windowtext 1.0pt; border-color:"&Session("RCM273_CompanyColour")&"; ' cellspacing='0' cellpadding='2' width='600' align='center' 'bgcolor=#ffffff' >"
               HTML = HTML & "<tr style='border:solid windowtext 1.0px;  border-color:"&Session("RCM273_CompanyColour")&"'>"
               HTML = HTML & "<td style='border:solid windowtext 1.0px;height:25px;  border-color:"&Session("RCM273_CompanyColour")&"' bgcolor='"&Session("RCM273_CompanyColour")&"' colspan='2'  class='Yellow'>"

               HTML = HTML & Subject1
               HTML = HTML & "</td></tr>"

               HTML = HTML & "<tr><td colspan='2' class='text'>"

                DateEntered = Day(Session("RCM273_LocalTime"))&"/"&MonthName(Month(Session("RCM273_LocalTime")))&"/"&Year(Session("RCM273_LocalTime"))
                if Session("RCM273_LocalTime")="" then
                        DateEntered = Day(Now)&"/"&MonthName(Month(Now))&"/"&Year(Now)
                end if
               if s_d("Quotation")=True  then
                     HTML = HTML & "<div align=justify> "

                      quotLetter="QuotationLett"&CINT(Session("RCM273_PickupLocationID"))
                      If (FileObject.FileExists(FilePath & "\"&quotLetter&".txt" )=true) Then
                           theLetter=FilePath & "\"&quotLetter&".txt"
                     else
                           theLetter=FilePath & "\QuotationLett.txt"
                     end if

                     HTML = HTML & ReadFileUTF8(theLetter)



                     HTML = HTML & "</pre>"


                     HTML = HTML & "<tr><td colspan='2' class='text'>If you would like to proceed with this quotation :</td></tr>"

                     '---1. need pass CompanyID and ServerID ( not compayCode/DbName)
                     '---2. need pass Quotation No


                     rcmReferenceID=s_d("rcmReferenceID")
                     QuotationNo = s_d("ReservationNo")
                     ValidDateTo=""
                     if s_d("DaysQuoteValid")  <>0 and ISNull(s_d("DaysQuoteValid"))<>True then
                              ValidDateTo=Day(Now+s_d("DaysQuoteValid"))&"/"&Left(MonthName(Month(Now+s_d("DaysQuoteValid"))),3)&"/"&Year(Now+s_d("DaysQuoteValid"))
                     end if
                     HTML = HTML & "<tr ><td class=OpeningTD colspan='2'> "
                     HTML = HTML & "<table><tr><td class='text'><b>TURN YOUR QUOTE INTO A BOOKING REQUEST "
                     HTML = HTML & " <b><a href='"&rcmURL&"/s_QPay.asp?id="&RCMCompanyID&"&MC2="&ServerID&"&ferq="&s_d("ReservationNo")&"&refID="&rcmReferenceID&"' >"
                     HTML = HTML & " <font color='#FF0000'>CLICK HERE</font></a> </b></td><td><img src='"&imageURL&"/images/SafePayment.jpg'  border='0' /></td></tr>"


                     if s_d("DaysQuoteValid")  <>0 then
                              HTML = HTML & "<tr><td class='text'><b>Please note this quote is valid for "&s_d("DaysQuoteValid")&" days. </b></td></tr>"
                     end if
                     HTML = HTML & "<tr><td class='text'  colspan='2'>Note : If you are viewing this message in text mode and having difficulties opening the above link, please try copying and pasting the following entire link into the address bar of your Internet homepage.</td></tr>"
                     HTML = HTML & "<tr><td class='text' colspan='2'>"&rcmURL&"/s_QPay.asp?id="&RCMCompanyID&"&MC2="&ServerID&"&ferq="&s_d("ReservationNo")&"&refID="&rcmReferenceID&"&VD2="&ValidDateTo&"</td></tr>"
                     HTML = HTML & "</table></td></tr>"

                       HTML = HTML & "<tr><td colspan='2' class='text'>Your Quotation details are as follows:</td></tr>"
                       HTML = HTML & "<tr class='HighlightRow' bgcolor='#EEF1F4'><td class=OpeningTD colspan='2'><FONT class='text'>Quotation Date: "
                       HTML = HTML & DateEntered
                       HTML = HTML & "</font></td></tr>"
                       HTML = HTML & "<tr><td style='height:1px' colspan='2'  bgcolor="&Session("RCM273_CompanyColour")&"></td></tr>"
                       HTML = HTML & "<tr ><td class='GREYTEXT' style='width:120px' ><b>Ref:</b></td> <td class='text' align='left' style='width:520px'>#Q-"&Session("RCM273_DocPrFix")&""&Session("RCM273_BookingBufferNo")&" ("&Session("RCM273_PickupLocation")&")</td></tr>"
            else
                         HTML = HTML & "Thank you for your "&onlineType&" with "&Company&"."
                       HTML = HTML & "<tr><td colspan='2' class='text'>Your "&onlineType&" has now been forwarded to the Location - "
                       HTML = HTML & Session("RCM273_PickupLocation")
                       HTML = HTML & ".<tr><td colspan='2' class='text'>You will receive a email confirmation from the renting location shortly."
                       HTML = HTML & "<tr><td colspan='2' class='text'>Your booking details are as follows:</td></tr>"
                       HTML = HTML & "<tr class=HighlightRow bgcolor=#EEF1F4><td class=OpeningTD colspan='2'><FONT class='text'>Booking Date: "
                       HTML = HTML & DateEntered
                       HTML = HTML & "</font></td></tr>"
                       HTML = HTML & "<tr><td style='height:1px' colspan='2'  bgcolor='"&Session("RCM273_CompanyColour")&"'></td></tr>"
                       HTML = HTML & "<tr ><td class='GREYTEXT'><b> Ref:</b></td><td class='text' align='left'><b>#U-"&Session("RCM273_DocPrFix")&""&Session("RCM273_BookingBufferNo")&" ("&Session("RCM273_PickupLocation")&")</b></td></tr>"


            end if

              '-------get booking information
            HTML = HTML & "<tr ><td class='GREYTEXT' >Name: <td class='text' align='left' >"&Session("RCM273_FirstName")&"&nbsp;"&Session("RCM273_LastName")&"</td>"
         

               HTML = HTML & "<tr ><td class='GREYTEXT' >Email: <td class='text' align='left'>"&Session("RCM273_CustomerEmail")&"</td>"
               HTML = HTML & "<tr><td style='height:1px' colspan='2'  bgcolor="&Session("RCM273_CompanyColour")&"></td></tr>"
               HTML = HTML & "<tr ><td class='GREYTEXT' >Vehicle Type<td class='text' align='left'>"&Session("RCM273_CarType")&"</td>"

                HTML = HTML & "<tr ><td><td class='text' align='left'><img  src='"&theimage&"' width='192' /></td></tr>"
                HTML = HTML & "<tr ><td class='GREYTEXT' >Pickup Location: <td class='text' align='left'>"&Session("RCM273_PickupLocation")&"</td>"
                HTML = HTML & "<tr ><td class='GREYTEXT' >Pickup Date Time: <td class='text' align='left'>"
               HTML = HTML & Day(Session("RCM273_RequestPickDate"))&"/"&Left(MonthName(Month(Session("RCM273_RequestPickDate"))),3)&"/"&Year(Session("RCM273_RequestPickDate"))
               HTML = HTML & "&nbsp;"
               HTML = HTML & Session("RCM273_RequestPickTime")
               HTML = HTML & "</td></tr>"

             if Session("RCM273_Flight")<>"" then
                    HTML = HTML & "<tr ><td class='GREYTEXT' >Arrival: <td class='text' align='left'>"&Session("RCM273_Flight")&"</td>"
               end if

                  HTML = HTML & "<tr ><td class='GREYTEXT' >Dropoff Location: <td class='text' align='left'>"&Session("RCM273_DropoffLocation")&"</td>"
                  HTML = HTML & "<tr ><td class='GREYTEXT' >Dropoff Date Time: <td class='text' align='left'>"
                  HTML = HTML & Day(Session("RCM273_RequestDropDate"))&"/"&Left(MonthName(Month(Session("RCM273_RequestDropDate"))),3)&"/"&Year(Session("RCM273_RequestDropDate"))
                  HTML = HTML & "&nbsp;"
                  HTML = HTML & Session("RCM273_RequestDropTime")
                  HTML = HTML & "</td></tr>"

                if Session("RCM273_Flightout")<>"" then
                  HTML = HTML & "<tr ><td class='GREYTEXT' >Departure: <td class='text' align='left'>"&Session("RCM273_Flightout")&"</td>"
                end if

                if Session("RCM273_NoTravelling")<>"" then
                  HTML = HTML & "<tr ><td class='GREYTEXT' >No. of People Travelling: <td class='text' align='left'>"&Session("RCM273_NoTravelling")&"</td>"
               end if

                   HTML = HTML & "<tr ><td class='GREYTEXT' >No Days: <td class='text' align='left'>"&Session("RCM273_TotalRentalDays")&"</td>"
                   'HTML = HTML & "<tr><td style='height:1px' colspan='2'  bgcolor='"&Session("RCM273_CompanyColour")&"'></td></tr>"
                   HTML = HTML & "<tr class=HighlightRow><td class=OpeningTD colspan='2'><Span class='text'>Rental Rate and Fees</td></tr>"


                '----the seasonal rates
                HTML = HTML & "<tr><td colspan='2' >"

                HTML = HTML & "<table>"


              Set s_rate=webConn.Execute("msp_res_DailyRatesBuffer "&s_d("ReservationNo")&"  ")
             CarTotalCost=0
            cost=0


            DO WHILE NOT s_rate.EOF
                     LogNotes = LogNotes&""&s_rate("Days")&" days x "&s_rate("Rate")&", "
                   if s_rate("Season")="Default" then
                           Season="Rates"
                     else
                        Season=s_rate("Season")
                     end if
                      NoHours=s_rate("NoHours")
                     cost=s_rate("Days")*s_rate("Rate")
                     CarTotalCost=CartotalCost + s_rate("Days")*s_rate("Rate")
                     'TotalRentalDays=TotalRentalDays+s_rate("Days")
                     if NoHours=0 then

                             HTML = HTML & "<tr><td class='GREYTEXT'>"
                             HTML = HTML & Season
                             HTML = HTML & "&nbsp;&nbsp;"
                             HTML = HTML & s_rate("Days")
                             HTML = HTML & "&nbsp;days @&nbsp;"
                     else
                              HTML = HTML & "<tr><td class='GREYTEXT'>"
                              HTML = HTML & Season
                              HTML = HTML & "&nbsp;&nbsp;"
                              HTML = HTML & s_rate("NoHours")
                              HTML = HTML & "&nbsp;hours @&nbsp;"

                               NoDaysDescription=s_rate("Days")&" "&s_d("sRateTypeDesc")&"s @ "&DisplayCurrency(s_rate("Rate"))
                              if s_rate("Days")<=1 and InStr(s_rate("RateName"), "hour")>0 then
                                       NoDaysDescription=s_rate("RateName")
                              end if
                              if   s_rate("NoHours")>0 then '--for hour rate
                                    NoDaysDescription=s_rate("RateName")
                              end if
                              if s_rate("DiscountType")="p" then
                                       Discount="(include "&s_rate("DiscountPerc")&"% discount)"
                              else
                                       Discount="(include "&DisplayCurrency(s_rate("DiscountPerc"))&" discount)"
                              end if

                     end if


                        if s_rate("DiscountPerc")>0 then
                              HTML = HTML & "<tr><td class=greytext>"&s_rate("Season")&":"
                              HTML = HTML & "&nbsp;"&NoDaysDescription&"&nbsp;  "
                              if s_rate("NoHours")=0 or isNull(s_rate("NoHours"))=True then
                               HTML = HTML & "&nbsp;"&s_rate("DiscountName")&"  "&Discount&" Rate @ "&DisplayCurrency(s_rate("Rate"))&""
                               end if
                               HTML = HTML & "</td><td align=right class=text>&nbsp;"
                              HTML = HTML & DisplayCurrency(s_rate("Cost"))
                              HTML = HTML & "</td></tr>"
                        else
                              HTML = HTML & "<tr><td class=greytext>"&s_rate("Season")&":"
                              HTML = HTML & "&nbsp;"&NoDaysDescription&"&nbsp; </td>"
                              HTML = HTML & "<td align=right class=text>&nbsp;"
                              HTML = HTML & DisplayCurrency(s_rate("Cost"))
                              HTML = HTML & "</td></tr>"
                        end if

               s_rate.MoveNext
               Loop
               s_rate.close
                set s_rate=nothing
                
               if  Session("RCM273_KmsDesc")<>"" then
                     '  HTML = HTML & "<tr><td colspan='2' class='GREYTEXT'>"&Session("RCM273_KmsDesc")&" </td></tr>"
                     HTML = HTML & "<tr><td class='GREYTEXT'>"&Session("RCM273_KmsDesc")&" </td><td class='text'>"
                     if Session("RCM273_KmsCost")<>0 then
                           HTML = HTML & Session("RCM273_CompanyCurrency")
                           HTML = HTML & FormatNumber(Session("RCM273_KmsCost"),2)
                     end if
                     HTML = HTML & "</td></tr>"
               end if
               if Session("RCM273_AreaofUsed")<>"" then
                         HTML = HTML & "<tr><td class='GREYTEXT'>Area of Use:<td class='text'>"&Session("RCM273_AreaofUsed")&" </td></tr>"
               end if
                '--- extra fees
         'Set s_extra=webConn.Execute("SELECT * from QPaymentExtraFeesBuffer WHERE (ReservationNo)='"&Session("RCM273_BookingBufferNo")&"' ORDER BY TYPE, NAME ")
               Set s_extra=webConn.Execute("SELECT PaymentExtraFeesBuffer.*, Name,ExtraDesc, Type FROM PaymentExtraFeesBuffer, ExtraFees WHERE  (ReservationNo)='"&Session("RCM273_BookingBufferNo")&"' and ExtraFeesID = ExtraFees.ID ORDER BY TYPE, NAME ")
                 EachExtraFees=0
              TotalExtraFees=0
              DO WHILE NOT s_extra.EOF
                       ExtraFeesQTY=s_extra("QTY")
                       EachExtraFees=s_extra("ExtraValue")
                       theQTY=""
                       if ExtraFeesQTY>1 then
                       theQTY="(QTY "&ExtraFeesQTY&")"
                       end if
                       if s_extra("Type")="Percentage" then
                                 HTML = HTML & "<tr><td class='GREYTEXT'>"&s_extra("Name")&",   "&theQTY&":</td><td  align=right class='text'>"&Session("RCM273_CompanyCurrency")&""&FormatNumber(EachExtraFees,2)&"</td></tr>"
                       elseif s_extra("Type")="Daily" then
                             extraType =s_extra("Type")
                             if s_extra("SetMaxPrice")  =True then
                                extraType="Fixed"
                             end if
                             HTML = HTML & "<tr><td class='GREYTEXT'>"&s_extra("Name")&",  "&extraType&" at "&Session("RCM273_CompanyCurrency")&""&FormatNumber(s_extra("Fees"),2)&" "&theQTY&":</td><td  align=right class='text'>"&Session("RCM273_CompanyCurrency")&""&FormatNumber(EachExtraFees,2)&"</td></tr>"
                       elseif s_extra("Type")="Fixed" then
                             HTML = HTML & "<tr><td class='GREYTEXT'>"&s_extra("Name")&",  "&s_extra("Type")&" at "&Session("RCM273_CompanyCurrency")&""&FormatNumber(s_extra("Fees"),2)&" "&theQTY&":</td><td  align=right class='text'>"&Session("RCM273_CompanyCurrency")&""&FormatNumber(EachExtraFees,2)&"</td></tr>"
                       end if

                       TotalExtraFees=TotalExtraFees+EachExtraFees
                             'Response.Write("<br>"&s_extra("Name")&"--- "&extraType&" at "&Session("RCM273_CompanyCurrency")&""&FormatNumber(s_extra("Fees"),2)&" "&theQTY&":</td><td>"&Session("RCM273_CompanyCurrency")&""&FormatNumber(EachExtraFees,2)&"</br>")
                       if s_extra("ExtraDesc")<>"" then
                            HTML = HTML & "<tr><td class='GREYTEXT'><font color=#FF6600>"&s_extra("ExtraDesc")&"</td><td></td></tr>"
                         end if

              s_extra.MoveNext
              Loop
              s_extra.close
              set s_extra=nothing
               TotalCharges=0
                TotalCharges=CarTotalCost+TotalExtraFees + Session("RCM273_StampDuty") +Session("RCM273_KmsCost")

               if  Session("RCM273_StampDuty") >0 then
                    HTML = HTML & "<tr><td class='GREYTEXT'>"
                    HTML = HTML & Session("RCM273_TaxName2")
                    HTML = HTML & "</td><td  align=right class='text'>"&Session("RCM273_CompanyCurrency")&""
                    HTML = HTML & FormatNumber(Session("RCM273_StampDuty"),2)
                    HTML = HTML & "</td></tr>"
               end if

                 if Session("RCM273_GSTInclusive") = "Yes"  then
                          TotalCharges=CarTotalCost+TotalExtraFees +Session("RCM273_KmsCost") + Session("RCM273_StampDuty")
                          HTML = HTML & "<tr><td class='GREYTEXT'><b>Estimate of Charges:</b></td><td  align=right class='text'>"&Session("RCM273_Currency")&" "&Session("RCM273_CompanyCurrency")&""
                          HTML = HTML & FormatNumber(TotalCharges,2)
                          HTML = HTML & "</td></tr>"
                          HTML = HTML & "<tr><td class='GREYTEXT' colspan='2' align=right>(All Prices GST Inclusive)"
                          HTML = HTML & "</td></tr>"
                   else
                          TotalCharges=CarTotalCost+TotalExtraFees +Session("RCM273_KmsCost") + Session("RCM273_StampDuty") +Session("RCM273_GST")
                          if Session("RCM273_GST")>0 then
                          HTML = HTML & "<tr><td class='GREYTEXT'>"&Session("RCM273_TaxName1")&"</b></td><td  align=right class='text'>"&Session("RCM273_CompanyCurrency")&""
                          HTML = HTML & FormatNumber(Session("RCM273_GST"),2)
                          HTML = HTML & "</td></tr>"
                          end if
                          HTML = HTML & "<tr><td class='GREYTEXT'><b>Estimate of Charges:</b></td><td  align=right class='text'>"&Session("RCM273_Currency")&" "&Session("RCM273_CompanyCurrency")&""
                          HTML = HTML & FormatNumber(TotalCharges,2)


                          HTML = HTML & "</td></tr>"
                  end if

                   IF s_d("Quotation")=False then
                     '----- payment
                        PaidAmount=0
                       Set s_p=WebConn.Execute("msp_res_PaymentBuffer "&s_d("ReservationNo")&"  ")
                        While Not s_p.EOF
                                 HTML = HTML & "<tr><td class=greytext>"
                                 if s_p("PaymentType")<>"Agent Collected" then
                                       HTML = HTML & "Paid by "
                                 end if
                                 HTML = HTML & ""&s_p("PaymentType")&""&s_p("PaymentRefNo")&" ("&s_p("sPaymentDate")&")</td>"
                                 HTML = HTML & "<td class=text align=right>"&DisplayCurrency(s_p("Paid"))&"</td></tr>"
                                if s_p("PaymentType")="Agent Collected" then
                                 AgentCollected=s_p("Paid")
                                end if
                   s_p.MoveNext
                     Wend
                     s_p.close
                     set s_p=nothing
                     Balance=Round(TotalCharges,2)- PaidAmount

                     HTML = HTML & "<tr><td class=greytext><b>Balance Due:</td><td class=text align=right><b>"&Session("RCM273_Currency")&" "&Session("RCM273_CompanyCurrency")&" "&FormatNumber(Balance,2)&"</td></tr>"

                  end if


                  HTML = HTML & "</table>"
                  HTML = HTML & "</td></tr>"

               if s_d("Notes") <>"" then
                     HTML = HTML & "<tr><td colspan='2' >&nbsp;</td></tr>"
                     HTML = HTML & "<tr><td class='GREYTEXT' colspan='2'>Comments/Requests:</td></tr>"
                     HTML = HTML & "<tr><td class='GREYTEXT' colspan='2'>"&s_d("Notes")&"</td></tr>"
                end if


      IF s_d("Quotation")=True  then

                  '---if pickuplocation quotation text exist, use it,
                     quotFile="Quotation"&CINT(Session("RCM273_PickupLocationID"))

                     'Response.Write(FilePath)
                  If (FileObject.FileExists(FilePath & "\"&quotFile&".txt" )=true) Then
                      FileName=FilePath & "\"&quotFile&".txt"
                  else
                      FileName=FilePath+"\Quotation.txt"
                  end if



                  'Set MFile=FileObject.OpenTextFile(FileName, 1)
                  HTML = HTML & "<tr><td style='height:1px' colspan='2'  bgcolor='"&Session("RCM273_CompanyColour")&"'></td></tr>"
                  HTML = HTML & "<tr><td class='GREYTEXT' colspan='2'>"
                  'HTML = HTML & MFile.ReadAll
                 ' ---new code
                   HTML = HTML & ReadFileUTF8(FileName)
                  HTML = HTML & "</td></tr>"
                  'MFile.Close
                  'Set MFile=nothing

      ELSE
                '---if pickuplocation confirmation text exist, use it,
               EmailText="EmailConfirmation"&CINT(Session("RCM273_PickupLocationID"))
              
                'Response.Write(FilePath)
               If (FileObject.FileExists(FilePath & "\"&EmailText&".txt" )=true) Then
                   FileName=FilePath & "\"&EmailText&".txt"
               else
                   FileName=FilePath+"\EmailConfirmation.txt"
               end if


                 ' Set MFile=FileObject.OpenTextFile(FileName, 1)
                  HTML = HTML & "<tr><td style='height:1px' colspan='2'  bgcolor='"&Session("RCM273_CompanyColour")&"'></td></tr>"
                  HTML = HTML & "<tr><td class='GREYTEXT' colspan='2'>"
                 ' HTML = HTML & MFile.ReadAll
                 HTML = HTML & ReadFileUTF8(FileName)
                  HTML = HTML & "</td></tr>"
                  'MFile.Close
                  'Set MFile=nothing

      END IF

         HTML = HTML & "<tr><td class='text' colspan='2'><br><b>"&Company&"</b></br></td></tr>"
         HTML = HTML & "<tr><td class='text' colspan='2'>"&Session("RCM273_Pickuplocation")&"</td></tr>"
    
         HTML = HTML & "<tr><td style='border:solid windowtext 1.0px;  border-color:"&Session("RCM273_CompanyColour")&"' bgcolor='"&Session("RCM273_CompanyColour")&"' align=center colspan='2'  class=Yellow>"
         HTML = HTML & BranchAddress
         HTML = HTML & "<br>"
         HTML = HTML & BranchCity
         HTML = HTML & ",&nbsp;"
         HTML = HTML & BranchPostCode
         HTML = HTML & "</br>"
         HTML = HTML & "<br>Phone:"
           HTML = HTML & BranchPhone
           HTML = HTML & "&nbsp;&nbsp;&nbsp;Fax: "
           HTML = HTML & BranchFax
           HTML = HTML & FreePhone
           HTML = HTML & "</br>"
          HTML = HTML & "</td></tr>"
          HTML = HTML & "</table><center>"
          HTML = HTML & WebSite
            HTML = HTML & "</center>"


       MyCDONTSMail.SetLocaleIDs(65001)
       MyCDONTSMail.Body=HTML
       MyCDONTSMail.Send
       set MyCDONTSMail=Nothing
       
      'Response.write "<br/><br/>" & HTML




        '--insert into reslog
            NoteDateUpdate=Day(Now)&"/"&MonthName(Month(Now))&"/"&Year(Now)
            SQL="INSERT INTO Reslog (ReasonID,ReservationNo,resBufferNo,LoginID,DateUpdated,ResTypeID, rateNotes)"
               SQL=SQL&"VALUES (17,0,"&Session("RCM273_BookingBufferNo")&",2,'"&NoteDateUpdate&"',0,'"&Tidyup(Left(LogNotes,160))&"')"
              webConn.Execute(SQL)
               'Response.Write("<br>")
             ' Response.Write(SQL)
        '---print out on the screen

        Response.Write("<table   align='center' width='720'  bgcolor='"&Session("RCM273_CompanyColour")&"' cellspacing='1' cellpadding='1'  border='0'>")
        Response.Write("<tr><td>")
      Response.Write("<table width='100%' align='center' bgcolor='#FFFFFF'  cellspacing='0' cellpadding='2'  border='0'>")
     ' Response.Write("<tr><td  style='height:20px' align='center' bgcolor='"&Session("RCM273_CompanyColour")&"' class='header' >Thank you for your "&onlineType&"</td></tr>")
      
        
        if FreeSale  = 1 then
            Response.Write("<tr><td>&nbsp;</td></tr>")
            if s_d("Quotation")=True  then
                  Response.Write("<tr><td align='center' class='text'>This quote  has been forwarded to your email</td></tr>")
                  Response.Write("<tr><td>&nbsp;</td></tr>")
                  Response.Write("<tr><td class='text' align='center' >Any additional info, please contact us at <a href='mailto:"&CompanyEmail&"'  class='re'd><font color='#3cb4e8'>"&CompanyEmail&"</font></a></h1></td></tr>")
                  Response.Write("<tr><td align='center' class='text'>Thanks for choosing "&Company&"..</td></tr> ")
              Response.Write("<tr><td>&nbsp;</td></tr>")

             else
                  Response.Write("<tr><td align='center' class='text'>Your Online Booking  has been forwarded to "&Company&"</td></tr>")
                  Response.Write("<tr><td align='center' class='text'>You will be contacted shortly with pick-up details and confirmation.</td></tr>")
                  Response.Write("<tr><td align='center' class='text'>An Online Booking confirmation has also been emailed to you.</td></tr>")

                  Response.Write("<tr><td align='center' class='text'>If you have any queries regarding the status of your Online Booking please don't hesitate to email us at <a href='mailto:"&CompanyEmail&"'  class='red'><font color='#3cb4e8'>"&CompanyEmail&"</font></a></td></tr>  ")
                  Response.Write("<tr><td align='center' class='text'>Thanks for choosing  "&Company&"..</td></tr> ")
              Response.Write("<tr><td>&nbsp;</td></tr>")
           end if
      else
            if s_d("Quotation")=True  then
                     Response.Write("<tr><td align='center' class='text'>This quote  has been forwarded to your email</td></tr>")
                     Response.Write("<tr><td>&nbsp;</td></tr>")
                     Response.Write("<tr><td class='text' align=center >Any additional info, please contact us at <a href='mailto:"&CompanyEmail&"'  class=red><font color='#3cb4e8'>"&CompanyEmail&"</font></a></h1></td></tr>")
                    Response.Write("<tr><td  align='center' class='text'>Thanks for choosing "&Company&"..</td></tr> ")
               Response.Write("<tr><td>&nbsp;</td></tr>")

             else
                     Response.Write("<tr><td class='text' align=center>Your booking request has been forwarded to "&Company&" </td></tr>")
                     Response.Write("<tr><td>&nbsp;</td></tr>")
                     Response.Write("<tr><td class='text' align=center >As this vehicle is heavily booked during this period you will be contacted shortly to confirm availability.</h1></td></tr>")
                     'Response.Write("<tr><td class='text' align=center >Please note: No charge will be made against your credit card until your reservation is confirmed.</h1></td></tr>")
                     Response.Write("<tr><td>&nbsp;</td></tr>")

                    Response.Write("<tr><td align='center' class='text'>If you have any queries regarding the status of your Online Booking Request please don't hesitate to email us at <a href='mailto:"&CompanyEmail&"'  class=red><font color='#3cb4e8'>"&CompanyEmail&"</font></a></td></tr>  ")
                     Response.Write("<tr><td  class='text'></td></tr>")
                     Response.Write("<tr><td align='center' class='text'>Thanks for choosing  "&Company&".</td></tr> ")
                Response.Write("<tr><td>&nbsp;</td></tr>")
            end if
       end if
    Response.Write("<tr><td  class='text' align='center'><A HREF=http://"&homeURL&">RETURN TO HOME PAGE</a> ")

  ' Response.Write("<input type='rcmbutton' name='Back' class=rcmbutton value='Close Window' onclick='javascript:window.close()'>")
  'Response.Write("</form>")
    Response.Write("</td></tr> ")
    Response.Write("<tr><td style='height:40px'>&nbsp;</td></tr>")
    Response.Write("</table>")
    Response.Write("</td></tr>")
    Response.Write("</table> ")
    

  

    
end if

Session("RCM273_OnlineBookingNo")=""
Session.Abandon
    webConn.CLOSE
    SET webConn=nothing
%>

<!-- END RCM HTML CODE-->
<!-- #include file="include_footer.asp" -->
</body>
</html>
 
