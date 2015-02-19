<%
ignorehack = True
allowanything = False
%>
<!--#include file="a_server-checks.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN"  "http://www.w3.org/tr/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<meta http-equiv="Content-type" content="text/html;charset=UTF-8" />
<title></title>
<!-- Copyright 2003-2011 GMH SYSTEMS LTD  -->
<!-- Product Development - Rental Car Manager  ( www.rentalcarmanager.com )  -->
<!--  All Rights Reserved  -->
<!--  This product and related documentation is protected by copyright and   -->
<!--  distributed under license restricting its use, copying, distribution  -->
<!--  and decompilation. No part of this product or related documentation may  -->
<!--  reproduced in any form by any means without prior written consent of -->
<!--  GMH Systems LTD  -->
<!--  For more information contact support@rentalcarmanager.com -->
</head>
<style type="text/css">
.white {  font-family: Arial;font-size: 10pt;color: #FFFFFF; font-weight: bold; }  
.text  {  font-family: Arial;font-size: 8pt;color: #666666; }  
.formtext {  font-family: Arial;font-size: 8pt;color: #666666; font-weight: bold; }
.header {  font-family: Arial;font-size: 10pt;color: #FFFFFF; font-weight: bold;
   }
.rcmbutton {
   font-family: Arial;
   font-size: 10pt;
      border-radius: 3px;
      height: 24px;
   color: #FFFFFF;
   background:#29A9BF;
   border: 1pt solid #0099CC;
    cursor:hand;
   cursor:pointer;}

input  {   font-family: Arial;font-size: 8pt;color: #666666; }
.greytext{  font-family: Arial;font-size: 8pt;color: #575757;font-weight: bold; }  
</style>


<%
FUNCTION tidyup(thisString)
    thisString=LTrim(RTrim(Replace(thisString, "'" , "")))
      thisString=Replace(thisString, "�" , "")
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
%>


<!--<body leftmargin=0  rightmargin=0 topmargin=0 onLoad="JavaScript:checkRefresh();" onUnload="JavaScript:prepareForRefresh();">-->
<body leftmargin=0  rightmargin=0 topmargin=0 onload="javascript:click_modal()">
 <center><img src="Header.jpg"  border="0"> </center>
<!-- RCM HTML CODE-->

<%

         CompanyCode="PtWestCoastCampers273"
         RCMURL="https://secure2.rentalcarmanager.com"


         Session("RCM273_CarImageName")=Replace(Session("RCM273_CarImageName"), "." , "&#46;")
         theimage="http://secure2.rentalcarmanager.com/DB/"&CompanyCode&"/"&Session("RCM273_CarImageName")
         ServerID=2
         RCMCompanyID=273

         logo="logo.gif"
         FilePath=Server.MapPath("/DB")&"\"&CompanyCode

IF Session("RCM273_OnlineBookingNo") ="" or ISNULL(Session("RCM273_OnlineBookingNo"))=True then
        Response.Redirect "webstep2.asp"

else 'if Session("RCM273_OnlineBookingNo")<>"" then

   '----------connect RCM Database---------------------

          Dim webConn
         Set webConn = Server.CreateObject("ADODB.Connection")
         DatabaseFile="PtWestCoastCampers273"
         webconn.Open "Provider=SQLOLEDB;Data Source = 4461QVIRT; Initial Catalog = "&DatabaseFile&";Trusted_Connection=yes;"



         if Request.QueryString("type")="Quote"  then
               Session("RCM273_Notes")=left(tidyup(Request.Form("Notes")),250)
               Session("RCM273_FirstName")=left(tidyup(Request.Form("firstname")),30)
               Session("RCM273_LastName")=left(tidyup(Request.Form("lastname")),30)
               Session("RCM273_countryID")=tidyup(Request.Form("countryID"))
               Session("RCM273_CustomerEmail")=left(tidyup(Request.Form("CustomerEmail")),50)
               Session("RCM273_phone")=left(tidyup(Request.Form("phone")),20)
               Session("RCM273_NoTravelling")=left(tidyup(Request.Form("NoTravelling")),30)
               Session("RCM273_DOB")= "1/Jan/1900"
               'Session("RCM273_LicExpDate") = "1/Jan/2008"
               Quotation=1
               'ReservationTypeID=3  '--for quotation
                ReservationTypeID=1  '--to save to reservationBuffer table
               onlineType="Online Quotation"

         else

             Quotation=0
               ReservationTypeID=0 '--need release the on request menu in RCM
            '--Oct 2014, If card details are not entered we save it  as a quote 1st , Quotation=1
              '--use Session(RCM273_bookingType"), If card details are not entered we save it to reservationBuffer as a quote 1st
               '--if cc entered we will update it to booking next step
             'Quotation=0
             Quotation=1
             '--new change, limited booking saved as request
              if Session("RCM273_categoryStatus") = "2" then '---LIMITED AVAILABILTY
                     ReservationTypeID=0 '--need release the on request menu in RCM '--old steps still saved as ReservationTypeID=1
               else
                     ReservationTypeID=1  '--confirmed booking
               end if



               DOBYear = CInt(Request.Form("DOBYear"))
               DOBMonth = (Request.Form("DOBMonth"))
               DOBDay  = CInt(Request.Form("DOBDay"))
               Session("RCM273_DOB")= DOBDay&"/"&DOBMonth&"/"&DOBYear
            'Response.write Session("RCM273_DOB")
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
            SQL=SQL&"VALUES (0,'"&Left(tidyup(Session("RCM273_address")),80)&"','"&Left(tidyup(Session("RCM273_city")),50)&"','"&Left(tidyup(Session("RCM273_state")),30)&"','"&Left(tidyup(Session("RCM273_postcode")),10)&"','"&Session("RCM273_CountryID")&"','"&Left(tidyup(Session("RCM273_fax")),15)&"',0,'"&Session("RCM273_DOB")&"','"&Session("RCM273_LicExpDate")&"','"&Left(Tidyup(Session("RCM273_License")),70)&"','"&Left(Tidyup(Session("RCM273_LicenseIssued")),30)&"','"&Left(Tidyup(Session("RCM273_FirstName")),30)&"','"&Left(Tidyup(Session("RCM273_LastName")),30)&"','"&Left(Tidyup(Session("RCM273_CustomerEmail")),50)&"','"&Left(Tidyup(Session("RCM273_phone")),20)&"' )"
            'response.Write(SQL)
            webConn.Execute(SQL)

         '---get customerID
              Set s_cus=webConn.Execute("SELECT Max(acID)  as customID FROM Customers where acLastName ='"&Session("RCM273_LastName")&"' ")
           CustomerID=s_cus("customID")
           Session("RCM273_CustomerID")=CustomerID

           'response.Write(CustomerID)

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
    'response.end

Session("RCM273_CustomerID")=CustomerID
   
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
        Session("RCM273_referralCommission")=0
        if   Session("RCM273_referral")<>0  then
            Set s_extra=webConn.Execute("SELECT Commission FROM Referrals WHERE  ID='"&Session("RCM273_referral")&"'  ")
            EachExtraFees=0
            if Not s_extra.EOF then
                    Session("RCM273_referralCommission")=Session("RCM273_RentalCost") * s_extra("Commission")/100
                    end if
         s_extra.Close
         Set s_extra=nothing
      end if


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
        rcmReferenceID=Session("RCM273_CustomerID")&""&iUniqueID
        Session("RCM273_rcmReferenceID")=rcmReferenceID
       'if rcmReferenceID="" then
       'rcmReferenceID
       'end if
        rcmReferenceID=Left(rcmReferenceID,30)
        'response.write "ref-" & rcmReferenceID

        DateEntered = Session("RCM273_LocalTime")
          if Session("RCM104_LocalTime")="" then
        DateEntered = Day(Now)&"/"&MonthName(Month(Now))&"/"&Year(Now)
        end if
      SQL="INSERT INTO ReservationBuffer (rcmReferenceID, CampaignCode,Transmission,AreaofUsed,RcmReferralBooking,ReferralID,ReferralCommission, driverage,DateEntered,UpdatedDate,EnteredBy,LastUpdatedBy, "
         SQL=SQL&" BrandID,Quotation,AddKmsFee,KmsFree,NoTravelling,Flightout,Flight,CollectionPoint,ReturnPoint,StampDuty,GST,PickupDateTime,DropoffDateTime,CustomerID,RentalSource,CarSizeID, "
         SQL=SQL&" PickupTime,Pickupdate,DropoffTime,DropoffDate,PickupLocationID,DropoffLocationID,Notes,ReservationTypeID)"
         SQL=SQL&" VALUES ('"&rcmReferenceID&"', '"&Left(Session("RCM273_PromoCode"),30)&"','"&Session("RCM273_transmission")&"', '"&Left(Session("RCM273_AreaofUsed"),40)&"', "
         SQL=SQL&" '"&RcmReferralBooking&"','"&Session("RCM273_referral")&"','"&Session("RCM273_referralCommission")&"', '"&Session("RCM273_driverage")&"','"&DateEntered&"', "
         SQL=SQL&" '"&DateEntered&"', 2,2,'"&BrandID&"',"&Quotation&","&Session("RCM273_AddKmsFee")&","&Session("RCM273_KmsFree")&",'"&Left(Session("RCM273_NoTravelling"),20)&"', "
         SQL=SQL&" '"&Left(Session("RCM273_Flightout"),50)&"','"&Left(Session("RCM273_Flight"),50)&"','"&Left(Session("RCM273_CollectionPoint"),80)&"', "
         SQL=SQL&" '"&Left(Session("RCM273_ReturnPoint"),80)&"','"&Session("RCM273_StampDuty")&"','"&Session("RCM273_GST")&"','"&Session("RCM273_RequestPickDateTime")&"', "
         SQL=SQL&" '"&Session("RCM273_RequestDropDateTime")&"' ,"&Session("RCM273_CustomerID")&",'"&Left(tidyup(Session("RCM273_RentalSource")),40)&"',"&(CInt(Session("RCM273_CarSizeID")))&", "
         SQL=SQL&" '"&Session("RCM273_RequestPickTime")&"','"&Session("RCM273_RequestPickDate")&"','"&Session("RCM273_RequestDropTime")&"', '"&Session("RCM273_RequestDropDate")&"', "
         SQL=SQL&" "&CINT(Session("RCM273_PickupLocationID"))&", "&Session("RCM273_DropoffLocationID")&", '"&left(tidyup(Session("RCM273_Notes")),250)&"',  '"&ReservationTypeID&"' ) "
       'response.write SQL
        webConn.Execute(SQL)

    

   '---get the ReservationNo
      Set s_No=webConn.Execute("SELECT Max(ReservationNo)  as ResNo FROM ReservationBuffer where  CustomerID='"&CustomerID&"' ")
      'OnlineBookingNo=s_No("ResNo")
      Session("RCM273_BookingBufferNo")=s_No("ResNo")
      s_No.close
      set s_No=nothing



        '--insert Unallocaited RA# to WebReservaiton table,
        '--and if is request set ReservationTypeID=0
             ' if Session("RCM273_categoryStatus") = "2" then '---LIMITED AVAILABILTY
                  '  ReservationTypeID=0 '--need release the on request menu in RCM
             ' else
                  '  ReservationTypeID=1  '--confirmed booking
              'end if
              '--all saved as quote -------
                 ReservationTypeID=3
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
      LogNotes = ""
      While Not s_rate.EOF
             RentalCost=s_rate("Days")* s_rate("Rate") + RentalCost
             LogNotes = LogNotes&""&s_rate("Days")&" days x "&s_rate("Rate")&", "
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


        TotalCharges=TotalCharges+Session("RCM273_StampDuty")
        ' Response.Write(Session("RCM273_URL"))
        '  Response.Write(Request.QueryString("dir"))

         '--insert into reslog
             webConn.Execute("Delete from Reslog where resBufferNo="&Session("RCM273_BookingBufferNo")&" ")
        NoteDateUpdate=Day(Now)&"/"&MonthName(Month(Now))&"/"&Year(Now)
            SQL="INSERT INTO Reslog (ReasonID,ReservationNo,resBufferNo,LoginID,DateUpdated,ResTypeID, rateNotes,BookingValue, exRateNotes)"
            SQL=SQL&"VALUES (17,0,"&Session("RCM273_BookingBufferNo")&",2,'"&NoteDateUpdate&"',0, '"&Tidyup(Left(LogNotes,160))&"', '"&TotalCharges&"', 'No CC details, saved as quote,need follow up')"
            webConn.Execute(SQL)
               'Response.Write("<br>")
             ' Response.Write(SQL)

if Request.QueryString("type")="Quote"  then
      Response.Redirect "webstep6.asp"

else '---if Request.QueryString("type")<>"Quote"  then

'---------------  Auric Window ---------------------
  '---=============if Vault not working, check below================================
          '--1. step4, SUB PersonalInfoForm, add qstring sreload=1
          '--2. on step5. Body need add javascript onload
          '--3. on step5. need return sreturnURL


Session("RCM273_TOK_Supplier")=""
Session("RCM273_TOK_Supplier")=0
TOK_RetentionMnth="3 months"
 Set SW=webconn.Execute("SELECT TS_ID, TS_Supplier,TS_RetentionMth FROM syTokenSupplier WHERE TS_Supplier='AURIC'  ")
If not SW.EOF Then
     Session("RCM273_TOK_Supplier")=SW("TS_Supplier")
     Session("RCM273_TOK_SupplierID")=SW("TS_ID")
     TOK_SupplierID=Session("RCM273_TOK_SupplierID")
     TOK_RetentionMnth=SW("TS_RetentionMth") & " months"
      TS_RetentionMthint=SW("TS_RetentionMth")
      if  SW("TS_RetentionMth")=14 then
      TOK_RetentionMnth="big-year"
      end if
end if

     '-------reditect to Auric window to ADD cc details only
     'Response.write "res-" & Session("RCM273_BookingBufferNo") & "-" & Session("RCM273_TOK_Supplier")
      if Session("RCM273_TOK_Supplier")="AURIC" And Session("RCM273_BookingBufferNo")>0 then
            'Session("RCM273_TOK_SupplierID")=SW("TS_ID")
             '---resType="" '- -for booking

            '---need sreturnURL,  and on the s_TokenCallVault.asp we will use it to pass the returning  url
           sreturnURL="/ssl/"&DatabaseFile&"/webstep6.asp?type=" & Request.QueryString("type") & "&categoryStatus=" & request.QueryString("categoryStatus")
           resType="U" '--for unallocated  booking
         
            ResNo = Session("RCM273_BookingBufferNo")
            custID = Session("RCM273_CustomerID")
            rcmRefID = Session("RCM273_rcmReferenceID")
            dbname=DatabaseFile
            'response.Write "<br/>CustID-" & custID

           sessionCCV="No"
           Set s_st=webConn.Execute("SELECT * FROM SystemTable WHERE Code='CCV'  ")
           If not s_st.EOF then
          sessionCCV = s_st("syValue")
           END IF
           s_st.CLOSE
           SET s_st=NOTHING


                  TOK_ExpiryDay=day(Date+90)
                  TOK_ExpiryMonth=Month(Date+90)
                  TOK_ExpiryYear=Year(Date+90)
                  TOK_Expiry = TOK_ExpiryDay&"/"&MonthName(TOK_ExpiryMonth)&"/"& TOK_ExpiryYear


       IF TS_RetentionMthint = 6 then

                  TOK_ExpiryDay=day(Date+184)
                  TOK_ExpiryMonth=Month(Date+184)
                  TOK_ExpiryYear=Year(Date+184)
                  TOK_Expiry = TOK_ExpiryDay&"/"&MonthName(TOK_ExpiryMonth)&"/"& TOK_ExpiryYear
       END  if
       IF TS_RetentionMthint = 9 then

                  TOK_ExpiryDay=day(Date+275)
                  TOK_ExpiryMonth=Month(Date+275)
                  TOK_ExpiryYear=Year(Date+275)
                  TOK_Expiry = TOK_ExpiryDay&"/"&MonthName(TOK_ExpiryMonth)&"/"& TOK_ExpiryYear
       END  if
      IF TS_RetentionMthint = 14 then

                  TOK_ExpiryDay=day(Date+424)
                  TOK_ExpiryMonth=Month(Date+424)
                  TOK_ExpiryYear=Year(Date+424)
                  TOK_Expiry = TOK_ExpiryDay&"/"&MonthName(TOK_ExpiryMonth)&"/"& TOK_ExpiryYear
                  TOK_RetentionMnth="big-year"
       END  if


         ' Response.Write "<br>TOK Expiry="
      '  Response.Write TOK_Expiry

           'sessionCCV="Yes"
          'sessionCCV="No"



            %>
<!--#Include virtual="/s_TokenCallVault.asp"-->
            <script type="text/javascript">
              function click_modal() {

                $('#modalvault').trigger("click");
                $(".ui-dialog-buttonpane button:contains('Close')").attr("disabled", true).addClass("ui-state-disabled");
                //alert('vurl : <%=vURL%>');
              }
               var type = getUrlVars()["type"];
               var category = getUrlVars()["categoryStatus"]

               

               function getUrlVars() {
                  var vars = [], hash;
                  var hashes = window.location.href.slice(window.location.href.indexOf('?') + 1).split('&');
                  for (var i = 0; i < hashes.length; i++) {
                     hash = hashes[i].split('=');
                     vars.push(hash[0]);
                     vars[hash[0]] = hash[1];
                  }
                  return vars;
               }
            </script>

             <A class='modal' HREF='<%= vURL & "&dir=add"%>' id="modalvault" ></A>
          <% 'Response.write vURL

   end if

'------------------- end Auric Window ---------------------
END IF
       
    webConn.CLOSE
    SET webConn=nothing
end if         
%>

<!-- END RCM HTML CODE-->
<style>
.ui-widget-overlay
{
  background: none repeat scroll 0 0 #E0E0E0 !important;
}

</style>
</body>
</html>
 
