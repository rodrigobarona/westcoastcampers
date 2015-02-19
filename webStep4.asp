<%
ignorehack = True
allowanything = False
%>
<!--#include file="a_server-checks.asp"-->

<!-- #include file="include_meta.asp" -->


<script type="text/javascript">
//Webstep4 

function ismaxlength(obj){
var mlength=obj.getAttribute? parseInt(obj.getAttribute("maxlength")) : ""
if (obj.getAttribute && obj.value.length>mlength)
obj.value=obj.value.substring(0,mlength)
}


function open_new_window(url)
{
new_window = window.open(url,'window_name','toolbar=0,menubar=0,resizable=0,dependent=0,status=0,width=540,height=540,left=25,top=15')
}
function checkNumericField(objName,minval, maxval,comma,period,hyphen)
{
   var numberfield = objName;
   if (chkNumeric(objName,minval,maxval,comma,period,hyphen) == false)
   {     numberfield.select();
      numberfield.focus();
      return false;
   }
   else
   {     return true;
   }
}
function chkNumeric(objName,minval,maxval,comma,period,hyphen)
{
// only allow 0-9 be entered, plus any values passed
// if all numbers allow commas, periods, hyphens or whatever,
// just hard code it here and take out the passed parameters
var checkOK = "0123456789";
var checkStr = objName;
var allValid = true;
var decPoints = 0;
var allNum = "";

for (i = 0;  i < checkStr.value.length;  i++)
{
ch = checkStr.value.charAt(i);
for (j = 0;  j < checkOK.length;  j++)
if (ch == checkOK.charAt(j))
break;
if (j == checkOK.length)
{
allValid = false;
break;
}
if (ch != ",")
allNum += ch;
}
if (!allValid)
{
alertsay = "Your Credit Card Number or Expiry Date is invalid"
alert(alertsay);
return (false);
}
// set the minimum and maximum
var chkVal = allNum;
var prsVal = parseInt(allNum);
if (chkVal != "" && !(prsVal >= minval && prsVal <= maxval))
{
alertsay = "Your Credit Card Number or Expiry Date is invalid"

alert(alertsay);
return (false);
}
}

function ValidateEmail()
{
  if (document.theForm1.lastname.value == "")
  {      alert("Last Name required.");
      document.theForm1.lastname.focus();
      return (false);
  }
  if (document.theForm1.CustomerEmail.value == "")
  {      alert("Email Address required.");
      document.theForm1.CustomerEmail.focus();
      return (false);
  }
   var emailRegEx = /^[a-zA-Z0-9._-]*\@[a-zA-Z0-9._-]*$/;
      if(!emailRegEx.test(theForm1.CustomerEmail.value))
      {
         alert("Invalid Email address");
         return false;
      }
return (true);
}
function ValidateQuote()
{
  if (document.theForm.firstname.value == "")
  {      alert("First Name required.");
      document.theForm.firstname.focus();
      return (false);
  }
  if (document.theForm.lastname.value == "")
  {      alert("Last Name required.");
      document.theForm.lastname.focus();
      return (false);
  }
  if (document.theForm.CustomerEmail.value == "")
  {      alert("Email Address required.");
      document.theForm.CustomerEmail.focus();
      return (false);
  }
   var emailRegEx = /^[a-zA-Z0-9._-]*\@[a-zA-Z0-9._-]*$/;
      if(!emailRegEx.test(theForm.CustomerEmail.value))
      {
         alert("Invalid Email address");
         return false;
      }
   if (document.theForm.Phone.value == "")
  {      alert("Phone No required.");
      document.theForm.Phone.focus();
      return (false);
  }
return (true);
}


function Validate()
{

   if (document.theForm.firstname.value == "")
  {      alert("First Name required.");
      document.theForm.firstname.focus();
      return (false);
  }
  if (document.theForm.lastname.value == "")
  {      alert("Last Name required.");
      document.theForm.lastname.focus();
      return (false);
  }

    if (document.theForm.DOBYear.value == "1900")
   {           alert("Please enter Driver's D.O.B ");
               theForm.DOBYear.focus();

               return (false);
      }
  if (document.theForm.License.value == "")
  {      alert("License# required.");
      document.theForm.License.focus();
      return (false);
  }
  if (document.theForm.LicenseIssued.value == "")
  {      alert("License Issued required.");
      document.theForm.LicenseIssued.focus();
      return (false);
  }
  if (document.theForm.address.value == "")
  {      alert("Please enter your address  details.");
      document.theForm.address.focus();
      return (false);
  }
   if (document.theForm.city.value == "")
  {      alert("Please enter your address details.");
      document.theForm.city.focus();
      return (false);
  }
   if (document.theForm.CustomerEmail.value == "")
  {      alert("Email Address required.");
      document.theForm.CustomerEmail.focus();
      return (false);
  }
   if (document.theForm.Phone.value == "")
  {      alert("Phone No required.");
      document.theForm.Phone.focus();
      return (false);
  }
   if (document.theForm.NoTravelling.value == "")
  {      alert("No. of People Travelling required.");
      document.theForm.NoTravelling.focus();
      return (false);
  }

return (true);
}  
</script>

</head>
<body  class="webstep3" >

<!-- #include file="include_header.asp" -->

<%    if Session("RCM273_PickupLocationID")="" then
            Response.Redirect "webstep2.asp"
      end if

      CompanyCode="PtWestCoastCampers273"
      RCMURL="https://secure2.rentalcarmanager.com"

      theimage="https://secure2.rentalcarmanager.com/DB/"&CompanyCode&"/"&Session("RCM273_CarImageName")

      Dim webConn
      Set webConn = Server.CreateObject("ADODB.Connection")
      DatabaseFile="PtWestCoastCampers273"
      webconn.Open "Provider=SQLOLEDB;Data Source = 4461QVIRT; Initial Catalog = "&DatabaseFile&";Trusted_Connection=yes;"

if Request.QueryString("dir")="Rate"  then '--from step3

         Session("TOK_Supplier")=""
      Session("TOK_SupplierID")=0
      Set SW=webConn.Execute("SELECT TS_ID, TS_Supplier FROM syTokenSupplier WHERE TS_Supplier='AURIC'  ")
      If not SW.EOF then
      Session("RCM273_TOK_Supplier")="AURIC"
      Session("RCM273_TOK_SupplierID")=SW("TS_ID")
      end if
       SW.Close
         set SW=nothing

           if Session("RCM273_SaveCCV") ="" then
        Set s_st=webConn.Execute("SELECT * FROM SystemTable WHERE Code='CCV'  ")
         If not s_st.EOF then
               Session("RCM273_SaveCCV") = s_st("syValue")
        END IF
       s_st.CLOSE
       SET s_st=NOTHING
   end if

     '---reset the session in the begging of the page, when user open more than one windows to compare rates, all the session will be maxed up, and cause problems
   Session("RCM273_categoryStatus") = CInt(Request.Form("categoryStatus"))
   Session("RCM273_PickupLocationID") = CInt(Request.Form("PickupLocationID"))
   Session("RCM273_DropoffLocationID")= CInt(Request.Form("DropoffLocationID"))

   Session("RCM273_RequestPickDate")=Request.Form("PickupDate")
  Session("RCM273_RequestDropDate")=Request.Form("DropoffDate")
   Session("RCM273_RequestPickTime")= Left(Request.Form("PickupTime"),5)
   Session("RCM273_RequestDropTime")= Left(Request.Form("DropoffTime"),5)

   Session("RCM273_RequestPickDateTime") = Session("RCM273_RequestPickDate")&" "&Session("RCM273_RequestPickTime")
   Session("RCM273_RequestDropDateTime") = Session("RCM273_RequestDropDate")&" "&Session("RCM273_RequestDropTime")
    Session("RCM273_DiscountRate")=Request.Form("DiscountRate")
   Session("RCM273_DiscountID")=Request.Form("DiscountID")
   Session("RCM273_DiscountType")=Request.Form("DiscountType")
   Session("RCM273_TotalRentalDays")=CDbl(Request.Form("TotalRentalDays"))



   FlightNoReqd="False"
   Set l_s = webConn.Execute("select Location, FlightNoReqd FROM Location WHERE ID='"&CInt(Request.Form("PickupLocationID"))&"' ")
   if Not l_s.EOF then
          Session("RCM273_PickupLocation")=l_s("Location")
          FlightNoReqd=l_s("FlightNoReqd")
   end if
   l_s.Close
   Set l_s=nothing
    Set l_s = webConn.Execute("select Location FROM Location WHERE ID='"&CInt(Request.Form("DropoffLocationID"))&"' ")
   if Not l_s.EOF then
          Session("RCM273_DropoffLocation")=l_s("Location")
   end if
   l_s.Close
   Set l_s=nothing

   if Request.Form("submit")="Request Booking" or Request.Form("submit")="Book Now" or Request.Form("submit")="Book" then
            Session("RCM273_bookingType")="Booking" '--used pass from ExistingCustInfoForm, to show the "PersonalInfoForm" or "PersonalInfoFormQuote"
      end if
       if Request.Form("submit")="Email Me Quote"  then
            Session("RCM273_bookingType")="Quote" '--used pass from ExistingCustInfoForm, to show the "PersonalInfoForm" or "PersonalInfoFormQuote"
      end if


end if



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

<table width="720" align="center" border="0" cellpadding="0" cellspacing="0">

<tr><td align="center" valign="top">

<!-- RCM HTML CODE-->

<%      

SUB   KmsReSelectedInSubs  '--used in all subs
              webConn.Execute("DELETE  FROM WebReservationFees WHERE rf_ReservationNo="&Clng(Session("RCM273_OnlineBookingNo"))&" ")
                SQL="SELECT * from KmsFree  WHERE ID="&CInt(Request.Form("KmsFreeID"))&" "
                  Set s_km=webConn.Execute(SQL)
                  KmsFree=s_km("KmsFree")
                  AddKmsFee=s_km("AddKmsFee")
                  KmsMaxcharge=s_km("MaxPrice")
                  KmsDailyRate=s_km("DailyRate")
                  Session("RCM273_KmsFree") =s_km("KmsFree")
                 Session("RCM273_AddKmsFee") =s_km("AddKmsFee")
                   sKmsMaxprice=""
                   Session("RCM273_KmsDesc")=""
                   if s_km("AddKmsFee")=0 and s_km("KmsFree")=0 and s_km("DailyRate")=0 then
                            '--unlimited without daily charge
                            Session("RCM273_KmsDesc")="Unlimited Kms"
                             '--if  pass Unlimited, save the 0 Daily rate  to the Reservationfees table
                             '--if daily rate =0 then print unlimited in the documents
                               Response.Write("<input type='hidden' name='KmsSaveFees' size='5' value='1' />")
                               SQL="INSERT INTO WebReservationFees (rf_ReservationNo,rf_MaxKmscharge,rf_DailyRate)"
                              SQL=SQL&"VALUES ("&Clng(Session("RCM273_OnlineBookingNo"))&","&s_km("MaxPrice")&","&s_km("DailyRate")&")"
                                  'Response.Write(SQL)
                              webConn.Execute(SQL)
                   elseif s_km("AddKmsFee")>0 and s_km("DailyRate")>0 then
                             iKmsDailyRate=s_km("DailyRate")
                           Session("RCM273_KmsDesc")="Daily rate @"&Session("RCM273_CompanyCurrency")&""&FormatNumber(s_km("DailyRate"))&", "&s_km("KmsFree")&" "&Session("RCM273_Mileage")&" free per day, additional per "&Session("RCM273_Mileage")&" "&Session("RCM273_CompanyCurrency")&" "&FormatNumber(s_km("AddKmsFee"))
                          Response.Write("<input type='hidden' name='KmsSaveFees' size='5' value='1' />")
                           SQL="INSERT INTO WebReservationFees (rf_ReservationNo,rf_MaxKmscharge,rf_DailyRate)"
                           SQL=SQL&"VALUES ("&Clng(Session("RCM273_OnlineBookingNo"))&","&s_km("MaxPrice")&","&s_km("DailyRate")&")"
                           'Response.Write(SQL)
                           webConn.Execute(SQL)
                           Session("RCM273_KmsCost")= iKmsDailyRate*Session("RCM273_TotalRentalDays")



                   elseif s_km("AddKmsFee")=0 and s_km("KmsFree")=0 and s_km("DailyRate") >0 then
                           '--unlimited with daily charge
                           iKmsDailyRate=s_km("DailyRate")
                          Session("RCM273_KmsDesc")="Unlimited Kms, daily rate @"&Session("RCM273_CompanyCurrency")&""&FormatNumber(s_km("DailyRate"))
                          Response.Write("<input type='hidden' name='KmsSaveFees' size='5' value='1' />")
                           SQL="INSERT INTO WebReservationFees (rf_ReservationNo,rf_MaxKmscharge,rf_DailyRate)"
                           SQL=SQL&"VALUES ("&Clng(Session("RCM273_OnlineBookingNo"))&","&s_km("MaxPrice")&","&s_km("DailyRate")&")"
                           'Response.Write(SQL)
                           webConn.Execute(SQL)
                           Session("RCM273_KmsCost")= iKmsDailyRate*Session("RCM273_TotalRentalDays")
                     elseif  s_km("MaxPrice")>0 then
                           sKmsMaxprice=", max charge "&Session("RCM273_CompanyCurrency")&""&FormatNumber(s_km("MaxPrice"))&" per day"
                           Session("RCM273_KmsDesc")=s_km("KmsFree")&" "&Session("RCM273_Mileage")&" free per day, additional per "&Session("RCM273_Mileage")&" "&Session("RCM273_CompanyCurrency")&""&FormatNumber(s_km("AddKmsFee"))&" "&sKmsMaxprice
                            Response.Write("<input type='hidden' name='KmsSaveFees' size='5' value='1' />")
                            SQL="INSERT INTO WebReservationFees (rf_ReservationNo,rf_MaxKmscharge,rf_DailyRate)"
                           SQL=SQL&"VALUES ("&Clng(Session("RCM273_OnlineBookingNo"))&","&s_km("MaxPrice")&","&s_km("DailyRate")&")"
                              'Response.Write(SQL)
                           webConn.Execute(SQL)
                          ' Session("RCM273_KmsCost")=s_km("MaxPrice")
                     else
                           Session("RCM273_KmsDesc")=s_km("KmsFree")&" "&Session("RCM273_Mileage")&" free per day, additional per "&Session("RCM273_Mileage")&" "&Session("RCM273_CompanyCurrency")&""&FormatNumber(s_km("AddKmsFee"))
                            Response.Write("<input type='hidden' name='KmsSaveFees' size='5' value='0' />")
                      end if
                     Response.Write("<input type='hidden' name='KmsFree' size='5' value='"&KmsFree&"'  />")
                     Response.Write("<input type='hidden' name='AddKmsFee' size='5' value='"&AddKmsFee&"'  />")
                     Response.Write("<input type='hidden' name='KmsMaxcharge' size='5' value='"&KmsMaxcharge&"' />")
                     Response.Write("<input type='hidden' name='KmsDailyRate' size='5' value='"&KmsDailyRate&"' />")
                     s_km.CLOSE
                     set s_km=nothing
                     TotalCost=TotalCost+   Session("RCM273_KmsCost")

END SUB

SUB ExistingCustInfoForm
    Session("RCM273_ccNumber")=""
          '============a same page caonot have 2 forms with same name
      Response.Write("<form method=post action='webstep4.asp?dir=existingCust&bookingtype="&Session("RCM273_bookingType")&"'  name='theForm1'  onSubmit='return ValidateEmail();'>")

 %>
  <tr><td colspan="4" height=25 align="center" bgcolor="<%=Session("RCM273_CompanyColour")%>"  class="header">Have you booked with us before, please enter your Email and Last name below</td></tr>
         <tr><td></td><td class="text" colspan="3">
         <table width="100%">
         <tr><td class="text" align="left" ><font color="red">*</font>Your Last Name:</td>
         <td align="left"  class="text"><input type="text" name="lastname" maxlength="30" size="30" value="<%=Session("RCM273_LastName")%>" /></td></tr>
       <tr><td align="left" class="text"><font color="red">*</font>Email Address:</td>
       <td align="left" class="text"  colspan="2"><input type="text" name="CustomerEmail" maxlength="50" size="30" value="<%=Session("RCM273_CustomerEmail")%>" /></td></tr>
        <tr><td align="left" class="text"></td><td align="left" class="red"  colspan="2">(Please check your Email Address is correct)</td></tr>

      <!--<tr><td align="left" class="text"><font color="red">*</font>Drivers License No.:</td>
       <td align="left" class="text"  colspan="2"><input type="text" name=License maxlength="20" size="30" value="<%=Session("RCM273_License")%>" /></td></tr>-->

       <tr><td colspan="2"  class="text" align="right" bgcolor="#FFFFFF">
       <input name="submit1" type="submit" class=rcmbutton value="CONTINUE"  />
       </td>
       </tr></table></form>
        </td></tr>

<%
END SUB
SUB PersonalInfoForm
 '============a same page caonot have 2 forms with same name
          Response.Write("<h4 class='smallm_title centered bigger'><span>Your personal information</span></h4><div class='row'>")

         Response.Write("<form method=post action='webstep5.asp?type=Quote&categoryStatus="&Request.QueryString("categoryStatus")&"&sreload=1'  name='theForm'  onSubmit='return Validate();'>")

      
%>        

        
         <div class="table-responsive col-xs-12">
               <table class="table table-hover">

          <% if Session("RCM273_ErrorMesage")<>"" then
            Response.Write("<tr><td align='left' class='text' bgcolor='#FFFFCC' colspan='3'><FONT color='red'><b>"&Session("RCM273_ErrorMesage")&"</b></td></tr>")
            Session("RCM273_ErrorMesage")=""
   end if %>
        <tr><td align="left" class="text">First Name:<font color="red">*</font></td><td align="left" class="text"  colspan="2"><input type="text" name="firstname" maxlength="30" size="30" class="form-control" value="<%=Session("RCM273_FirstName")%>" /></td></tr>
       <tr><td align="left" class="text">Last Name:<font color="red">*</font></td><td align="left" class="text" colspan="2"><input type="text" name="lastname" maxlength="30" size="30" class="form-control"  value="<%=Session("RCM273_LastName")%>" /></td></tr>
       <tr>
       <td align="left" class="text">Date of birth:<font color="red">*</font></td>
       <td align="left" class="text"  colspan="2">
          <table border="0" cellspacing="0" cellpadding="0">
          <tr>
          <td align="left" class="formtext">
               <select name="DOBDay" class="form-control">
           <%  DOBDay=0
                DOBMonth=0
                DOBYear=1900
                if ISDate(Session("RCM273_DOB"))=True and Session("RCM273_DOB")<>"1/Jan/1900" then
                     DOBDay=Day(Session("RCM273_DOB"))
                     DOBMonth=Month(Session("RCM273_DOB"))
                     DOBYear=Year(Session("RCM273_DOB"))
                end if
               Response.Write("<option value='1' selected='selected'>Day</option>")
            for pd=1 to 31
                  if (pd)= DOBDay    then
                        Response.Write("<option value='"&pd&"' selected='selected'>"&pd&"</option>")
                  else
                  Response.Write("<option value='"&pd&"'>"&pd&"</option>")
                  end if
          next
           %>     </select>
           </td>
            <td align="left" class="formtext" >
            <select name="DOBMonth" class="form-control">
         <%    
              Response.Write("<option value='Jan' selected='selected'>Month</option>")

            for pm=1 to 12
                  theMonth=Left(MonthName(pm),3)
                 if pm=DOBMonth    then
                  Response.Write("<option value='"&theMonth&"' selected='selected'>"&theMonth&"</option>")
                  else
                  Response.Write("<option value='"&theMonth&"'>"&theMonth&"</option>")
                  end if
          next
           %>  </select>
         </td>
       
            <td align="left" class="formtext"><select name="DOBYear" class="form-control">
             <%
                    for y=(Year(Date)-90) to (Year(Date)-18)
                        if y=DOBYear  then
                        Response.Write("<option value='"&y&"' selected='selected'>"&y&"</option>")
                        else
                        Response.Write("<option value='"&y&"'>"&y&"</option>")
                        end if
                  next
                  if DOBYear=1900 or Session("RCM273_DOB")=""  then
                  Response.Write("<option value='1900' selected='selected'>Year</option>")
                  end if
           %>     </select></td>
           </tr>
           </table>
        </td></tr>
        <tr><td align="left" class="text">License No:<font color="red">*</font></td><td align="left" class="text"  colspan="2"><input type="text" name="License" maxlength="30"  class="form-control" size="30" value="<%=Session("RCM273_License")%>" /></td></tr>
        <tr><td align="left" class="text">Country of issue:<font color="red">*</font></td><td align="left" class="text"  colspan="2">
		 <%   Set s_c=webConn.Execute("select * FROM Country ORDER BY  Country " )
            Response.Write("<select name='LicenseIssued' class='form-control'>")
            WHILE NOT s_c.EOF
                  if  Cstr(s_c("ID"))=Session("RCM273_LicenseIssued") then
                               Response.Write("<option value='"&s_c("ID")&"' selected='selected'>"&s_c("Country")&"</option>")
                        else
                           if s_c("Defaulted")=-1 then
                              Response.Write("<option value='"&s_c("ID")&"' selected='selected'>"&s_c("Country")&"</option>")
                           else
                                 Response.Write("<option value='"&s_c("ID")&"' >"&s_c("Country")&"</option>")
                           end if
                        end if
            s_c.MoveNext 
            Wend
            s_c.Close
            Set s_c=nothing
            Response.Write("</select>")

%>
		
		
		</td></tr>
        <tr>
        <td align="left" class="text">Expiry Date:<font color="red">*</font></td>
        <td align="left" class="text"  colspan="2">
        <table border="0" cellspacing="0" cellpadding="0">
        <tr>
        <td align="left" class="formtext"><select name="licensDay" class="form-control">
           <%  for d=1 to 31
                 DayLicExpDate=Day(Date)
                 MonthLicExpDate=Left(MonthName(Month(Date)),3)
                 YearLicExpDate=Year(Date)
                  if  ISDate(Session("RCM273_LicExpDate"))=True  then
                        DayLicExpDate=Day(Session("RCM273_LicExpDate"))
                        MonthLicExpDate=Left(MonthName(Month(Session("RCM273_LicExpDate"))),3)
                        YearLicExpDate=Year(Session("RCM273_LicExpDate"))
                  end if
                     if  d=DayLicExpDate then
                  Response.Write("<option value='"&d&"' selected='selected'>"&d&"</option>")

                  else
                  Response.Write("<option value='"&d&"'>"&d&"</option>")
                  end if
          next
           %>     </select></td>
       <td align="left" class="formtext"><select name="licensMonth" class="form-control">
         <%    for m=1 to 12
                 theMonth=Left(MonthName(m),3)
                     if theMonth=MonthLicExpDate then
                       Response.Write("<option value='"&theMonth&"' selected='selected'>"&theMonth&"</option>")
                  else
                  Response.Write("<option value='"&theMonth&"'>"&theMonth&"</option>")
                  end if
          next
           %>  </select></td>
   <td align="left" class="formtext">   <select name="licensYear" class="form-control">
            <%  for theYear=Year(Date) to (Year(Date)+10)
                  if theYear=YearLicExpDate then
                        Response.Write("<option value='"&theYear&"' selected='selected'>"&theYear&"</option>")
                  else
                        Response.Write("<option value='"&theYear&"'>"&theYear&"</option>")
                  end if
                next
           %>   </select></td>
           </tr>
           </table>


           </td>
           </tr>
        <tr><td align="left" class="text">Address:<font color="red">*</font></td><td align="left" class="text"  colspan="2"><input type="text" name="address" maxlength="80" size="30" class="form-control" value="<%=Session("RCM273_address")%>" /></td></tr>
        <tr><td align="left" class="text">City:<font color="red">*</font></td><td align="left" class="text"  colspan="2"><input type="text" name="city" size="30" maxlength="50" class="form-control" value="<%=Session("RCM273_city")%>" /></td></tr>
        <!-- tr><td align="left" class="text">&nbsp;State/Province:</td><td align="left" class="text"  colspan="2"><input type="text" name="state" size="30" maxlength="30" class="form-control" value="<%=Session("RCM273_state")%>"  /></td></tr -->
        <tr><td align="left" class="text">&nbsp;Postcode/ZIP:</td><td align="left" class="text"  colspan="2"><input type="text" name="postcode" size="30" maxlength="10" class="form-control" value="<%=Session("RCM273_postcode")%>" /></td></tr>
        <tr>
        <td align="left" class="text">Country of Residence:<font color="red">*</font></td>
        <td align="left" class="text"  colspan="2">
             <%   Set s_c=webConn.Execute("select * FROM Country ORDER BY  Country " )
            Response.Write("<select name='CountryID' class='form-control'>")
            WHILE NOT s_c.EOF
                  if  Cstr(s_c("ID"))=Session("RCM273_countryID") then
                               Response.Write("<option value='"&s_c("ID")&"' selected='selected'>"&s_c("Country")&"</option>")
                        else
                           if s_c("Defaulted")=-1 then
                              Response.Write("<option value='"&s_c("ID")&"' selected='selected'>"&s_c("Country")&"</option>")
                           else
                                 Response.Write("<option value='"&s_c("ID")&"' >"&s_c("Country")&"</option>")
                           end if
                        end if
            s_c.MoveNext 
            Wend
            s_c.Close
            Set s_c=nothing
            Response.Write("</select>")

%>
   </td></tr>
   <tr><td align="left" class="text">Email Address:<font color="red">*</font></td><td align="left" class="text"  colspan="2"><input type="email" name="CustomerEmail" maxlength="50" class="form-control" size="30" value="<%=Session("RCM273_CustomerEmail")%>" /><small>Please check your Email Address is correct</small></td></tr>

   <tr><td align="left" class="text">Phone (incl area code):<font color="red">*</font></td><td align="left" class="text"  colspan="2"><input type="number" name="Phone" maxlength="20" class="form-control" size="30" value="<%=Session("RCM273_phone")%>" /></td></tr>
<% if  FlightNoReqd="True" then %>
         <input type="hidden" name="FlightNoReqd" size="30"  maxlength="50" value=1 />
   <tr>
   <td class="text" align="left"><font color="red">*</font>Arrival Flight # (if airport pickup):</td>
   <td class="text" align="left"><input type="text" name="Flight" size="30"  maxlength="50" value="<%=Session("RCM273_Flight")%>" /></td>
   </tr>
   <%   else %>

         <input type="hidden" name="FlightNoReqd"  value=0 />
         <input type="hidden" name="Flight"  value="" />
   <%  end if %>
   <tr>
   <td class="text" align="left">Pick Up Required From:</td>
   <td class="text" align="left"><input type="text"  class="form-control"name="CollectionPoint" size="30" maxlength="80" value="<%=Session("RCM273_CollectionPoint")%>" /></td>
   </tr>
   <tr>
   <td class="text" align="left">Drop Off Required to:</td>
   <td class="text" align="left"><input type="text" class="form-control" name="ReturnPoint" size="30" maxlength="80" value="<%=Session("RCM273_ReturnPoint")%>" /></td>
       </tr>
        <tr>
        <td class="text" align="left">No. of People Travelling:<font color="red">*</font></td>
       <td class="text" align="left"><input type="number" min="1" max="5" class="form-control" name="NoTravelling" size="30" maxlength="20" value="<%=Session("RCM273_NoTravelling")%>" /></td>
      </tr>
       <!-- tr>
        <td class="text" align="left">Preference for car transmission:</td>
       <td class="text" align="left" colspan="2"><select name="transmission" class="form-control">
       <option value="" >No preference</option>
      <% if Session("RCM273_transmission") = "Manual"  then %>
                      <option value="Auto" >Automatic preferred</option>
                      <option value="Manual" selected="selected">Manual preferred</option>
         <% elseif Session("RCM273_transmission") = "Auto"  then %>
                      <option value="Auto" selected="selected">Automatic preferred</option>
                            <option value="Manual">Manual preferred</option> 
         <% else %>       
                      <option value="Auto" >Automatic preferred</option>
                        <option value="Manual">Manual preferred</option> 
         <% end if %>
         </select></td>
         </tr -->
         <tr>
        <td align="left" class="text" style="vertical-align: top;">Special Requests or Additional Details:</td>
        <td align="left" class="text" colspan="2"><textarea name="Notes" class="form-control" maxlength=250 onkeyup="return ismaxlength(this)"  cols="35" rows="5"><%=Session("RCM273_Notes")%></textarea></td>
         </tr>

         <% Set l_s = webConn.Execute("SELECT * FROM RentalSource where WebAvaliable=1  order by  RentalSource")
         if Not l_s.EOF then
            Response.Write("<tr><td align='left' class='text'>How did you find us?</td>")
            Response.Write("<td align='left' class='text' colspan='2'>")
            Response.Write("<select name='RentalSource' >")
            Response.Write("<option value='Online Booking' >Please select</option>")
            While Not l_s.EOF
               if l_s("RentalSource")=Session("RCM273_RentalSource") then
                     Response.Write("<option value='"&l_s("RentalSource")&"' selected='selected'>"&l_s("RentalSource")&" </option>")
               else
                     Response.Write("<option value='"&l_s("RentalSource")&"' >"&l_s("RentalSource")&" </option>")
               end if
            l_s.MoveNext
            Wend

            Response.Write("</select>")
            'Response.Write("</td>")
            'Response.Write("</tr>")
         else
            Response.Write("<input type='hidden' name='RentalSource' size='40' value='Online Booking'  />")
         end if

         l_s.Close
         Set l_s=nothing       %>

         <!--<tr><td></td>
         <td align=left class=text><font color=red>*</font>Accept Terms and Conditions:</td><td align=left class=text  colspan=2><input type="checkbox" name="acceptTerms" id="acceptTerms"></td></tr>
         <tr><td></td>
         <td  align=left class=text valign=top nowrap></td>
         <td align=left><A HREF="http://<%=Terms%>" target=new class=link>Click here to view our Terms and Conditions</a></td></tr>
      -->
         </table></td></tr>
      <% Session("RCM273_useDPSpayment")="No" %>

         <tr><td  colspan="4" >
         <table>
       <%

       if Session("RCM273_DPSErrorMesage")<>"" then %>
            <tr><td colspan="3" class="text"><font color=red >DPS Payment <%=Session("RCM273_DPSErrorMesage")%></font>
            <font color=red>(payment Amount$<%=Session("RCM273_DPSErrorMesage2")%>)</font>
            </td></tr>
   <% end if
       ' FilePath=Server.MapPath("/DB")&"\"&CompanyCode
      ' Set FileObject=Server.CreateObject("Scripting.FileSystemObject")
     '   Response.Write("<tr><td  class=text colspan=4 align=center ><b>")
      '  Set FileObject=Server.CreateObject("Scripting.FileSystemObject")
      '  thismessage=FilePath & "\TemplateCreditCard.txt"
      '  Set MFile=FileObject.OpenTextFile(thismessage, 1)
      '  Response.Write(MFile.ReadAll)
       ' Response.Write("</td></tr>")
       ' MFile.Close
       ' Set MFile=nothing
      %>

        </table></td></tr>

        <%
     '   Session("RCM273_useDPSpayment")="No"
     if Session("RCM273_useDPSpayment")="Yes" then
         PaymentAmount1=0.2*Session("RCM273_TotalEstimateofCharges")
         PaymentAmount2=Session("RCM273_TotalEstimateofCharges")
         Response.Write("<tr><td  style=""height:22px"" align=center bgcolor='"&Session("RCM273_CompanyColour")&"' class=header  colspan=4>Payment Options</td></tr>")
         Response.Write("<tr><td  colspan=4><table width=100% >")
         Response.Write("<tr><td class=text width=150 align=right><input type=radio  name=PaymentOptions  Value=1 CHECKED>   ")
         Response.Write("<td class=text align=left>Pay 20% Deposit Now,  Total "&Session("RCM273_CompanyCurrency")&""&FormatNumber(PaymentAmount1,2)&"</td></tr>")
         Response.Write("<tr><td class=text width=150 align=right><input type=radio  name=PaymentOptions  Value=2 >  ")
         Response.Write("<td class=text align=left>Pay Full Amount Now, Total "&Session("RCM273_CompanyCurrency")&""&FormatNumber(PaymentAmount2,2)&"</td></tr>")
         ' Username="GMHDev" '---for testing, changed password 22 Feb 2011
          '--use card # 41111111111xxxx for test
          Response.Write("<input type=hidden name=TxnType value=Purchase>")
          Response.Write("<input type=hidden name=PaymentCurrency value=NZD>")
         Response.Write("<tr><td class=text colspan=2><small> Credit card details are protected with SSL encryption &amp; processed securely in real time by: <a href='http://www.paymentexpress.com/about/about_paymentexpress/privacy_policy.html' target=black border=0></td></tr>")
         Response.Write("<tr><td class=text colspan=2><img src='images/paymentexpress.gif' width=139 height=25  border=0></a>")
         Response.Write("<li><a href='http://www.paymentexpress.com/about/about_paymentexpress/privacy_policy.html' target=black border=0> Information about payment</a> </li></td></tr>")
         Response.Write("</table></td></tr>")
      end if    %>



      <tr><td colspan="4" align="right"   >
	  <hr />
      <input type="button" value="Back" class='btn btn-default'  onclick="javascript:history.back(-1)" style="float: left;" />&nbsp;&nbsp;

    <% if Session("RCM273_categoryStatus") = "2" then '---LIMITED AVAILABILTY %>
       <input  onClick="document.pressed=this.value" name="submit1" class='btn btn-success btn-lg'  type="submit"  value="Request Booking"  ></td></tr>
<% else %>
    <!--  <input  onClick="document.pressed=this.value" name="submit1" class=rcmbutton  type="submit"  value="Confirm Booking"  ></td></tr> -->
 <input  onClick="document.pressed=this.value" name="submit1" class='btn btn-success btn-lg'  type="submit"  value="Confirm Booking"  ></td></tr>

<% end if %>

  <tr>
      <td colspan="4"  class="text"  align="right">
      <font color="red">*</font>
      <font color="red">Required fields must be completed</font>
      </td>
      </tr>
<%    Response.Write("</table>")
      Response.Write("</form>")
END SUB

 SUB PersonalInfoFormQuote
  '============a same page caonot have 2 forms with same name
         Response.Write("<form method='post' action='webstep5.asp?type=Quote&categoryStatus="&Request.QueryString("categoryStatus")&"'  name='theForm'  onsubmit='return ValidateQuote();'>")
        '--pass below hidden fields with all the forms
         Response.Write("<table width='100%' border='0' cellpadding='0' cellspacing='0'>")
         %>

         <h4 class='smallm_title centered bigger'><span>Your personal information</span></h4><div class='row'>
        <div class="table-responsive col-xs-12">
        <table class="table table-hover">
          <%

        if Session("RCM273_ErrorMesage")<>"" then
            Response.Write("<tr><td align='left' class='text' bgcolor='#FFFFCC' colspan='3'><FONT color='red'><b>"&Session("RCM273_ErrorMesage")&"</b>")
            response.write("</td>")
            Session("RCM273_ErrorMesage")=""
         end if %>
         <tr><td align="left" class="text">First Name:<font color="red">*</font></td><td align="left" class="text"  colspan="2"><input type="text" name="firstname" maxlength="30" size="30" class="form-control" value="<%=Session("RCM273_FirstName")%>"  /></td></tr>
       <tr><td align="left" class="text">Last Name:<font color="red">*</font></td><td align="left" class="text" colspan="2"><input type="text" name="lastname" maxlength="30"  class="form-control" size="30" value="<%=Session("RCM273_LastName")%>" /></td></tr>

        <!-- tr><td align="left" class="text">&nbsp;Address:</td><td align="left" class="text"  colspan="2"><input type="text" name="address" maxlength="80" size="30" value="<%=Session("RCM273_address")%>" /></td></tr>
        <tr><td align="left" class="text">&nbsp;City:</td><td align="left" class="text"  colspan="2"><input type="text" name="city" size="30" maxlength="50" value="<%=Session("RCM273_city")%>" /></td></tr>
        <tr><td align="left" class="text">&nbsp;State/Province:</td><td align="left" class="text"  colspan="2"><input type="text" name="state" maxlength="30" size="30" value="<%=Session("RCM273_state")%>"  /></td></tr>
        <tr><td align="left" class="text">&nbsp;Postcode/ZIP:</td><td align="left" class="text"  colspan="2"><input type="text" name="postcode" maxlength="10" size="30" value="<%=Session("RCM273_postcode")%>" /></td></tr>
        <tr><td align="left" class="text">&nbsp;Country of Residence:</td><td align="left" class="text"  colspan="2">
         <%    Set s_c=webConn.Execute("select * FROM Country ORDER BY  Country " )
               Response.Write("<select name='CountryID'>")
               WHILE NOT s_c.EOF
                       if  Cstr(s_c("ID"))=Session("RCM273_countryID") then

                              Response.Write("<option value='"&s_c("ID")&"' selected='selected'>"&s_c("Country")&"</option>")
                     else
                              if s_c("Defaulted")=-1 then
                                 Response.Write("<option value='"&s_c("ID")&"' selected='selected'>"&s_c("Country")&"</option>")
                              else
                                 Response.Write("<option value='"&s_c("ID")&"' >"&s_c("Country")&"</option>")
                              end if
                     end if
            s_c.MoveNext 
            Wend
            s_c.Close
            Set s_c=nothing
            Response.Write("</select></td></tr -->")

            Set l_s = webConn.Execute("select * FROM RentalSource where WebAvaliable=1  order by  RentalSource")
            if Not l_s.EOF then
                  Response.Write("<tr><td align='left' class='text'>How did you find us?</td>")
                  Response.Write("<td align='left' class='text' colspan='2'>")
                  Response.Write("<select name='RentalSource' >")
                  Response.Write("<option value='Online Booking' >Please select</option>")
                  While Not l_s.EOF
                     if l_s("RentalSource")=Session("RCM273_RentalSource") then
                        Response.Write("<option value='"&l_s("RentalSource")&"' selected='selected'>"&l_s("RentalSource")&" </option>")
                     else
                        Response.Write("<option value='"&l_s("RentalSource")&"' >"&l_s("RentalSource")&" </option>")
                     end if
                 l_s.MoveNext
                  Wend
                  Response.Write("</select>")
                  Response.Write("</td>")
                  Response.Write("</tr>")
             else
                   Response.Write("<input type='hidden' name='RentalSource' size='30' value='Online Booking'  />")
            end if
           
            l_s.Close
            Set l_s=nothing       %>

        <tr><td align="left" class="text">Email Address:<font color="red">*</font></td><td align="left" class="text"  colspan="2"><input type="text" name="CustomerEmail"  class="form-control" maxlength="50" size="30" value="<%=Session("RCM273_CustomerEmail")%>" /><small>Please check your Email Address is correct</small></td></tr>
        
       <tr><td align="left" class="text">Phone (incl area code):<font color="red">*</font></td><td align="left" class="text"  colspan="2"><input type="text" name="Phone"  class="form-control" maxlength="20" size="30" value="<%=Session("RCM273_phone")%>" /></td><hr /></tr>
       <!-- tr>
       <td class="text" align="left">Preference for car transmission:</td>
       <td class="text" align="left" colspan="2"><select name="transmission">
       <option value="" >No preference</option>
       <% if Session("RCM273_transmission") = "Manual"  then %>
             <option value="Auto" >Automatic preferred</option>
             <option value="Manual" SELECTed>Manual preferred</option>
       <% elseif Session("RCM273_transmission") = "Auto"  then %>
            <option value="Auto" SELECTed>Automatic preferred</option>
            <option value="Manual">Manual preferred</option>
       <% else %>
            <option value="Auto" >Automatic preferred</option>
            <option value="Manual">Manual preferred</option>
       <% end if %>
       </select></td>
       </tr>
       <tr>
        <td align="left" class="text">Additional Details:</td>
        <td align="left" class="text" colspan="2"><textarea name="Notes" maxlength=250 onkeyup="return ismaxlength(this)" cols="30"   rows="5"><%=Session("RCM273_Notes")%></textarea></td>
      </tr -->

      </table>
        </td>
        </tr>
       
        
        <tr><td colspan="2" align="left" ><input type="button" name="Back" class='btn btn-default'  value="Back"  onclick="javascript:history.back(-1)" />
         
      <input  name="submit1" type="submit" class='btn btn-info btn-lg' style="float:right;"  value="EMAIL QUOTE"  />
     </td></tr>

   <tr>
   <td  colspan="4"  class="text"  align="right">
   <font color="red">*</font>
   <font color="#FF6600">Required fields must be completed</font>
   </td>
   </tr>

<%    Response.Write("</table></div>")
      Response.Write("</form>")
END SUB

'------ check, Not to let under 21years proceed with booking



'---from step3.asp, save the booking detais-------------
if Request.QueryString("dir")="Rate"  then
         if  Request.Form("CarSizeID")="" then
               Response.Redirect "webstep2.asp"
         end if
         '---Vehicle Type----------
         Session("RCM273_CarType")=""
         Session("RCM273_CarSizeID")=0
         Session("RCM273_CarImageName")=""
         Session("RCM273_AreaofUsed")= tidyup(Request.Form("AreaofUsed"))
         Session("RCM273_transmission")=tidyup(Request.Form("transmission"))
                
         SQL="SELECT * FROM CarSize WHERE (ID)='"&Left(tidyup(Request.Form("CarSizeID")),3)&"'"
         Set s_cs=webConn.Execute(SQL)
         Session("RCM273_CarType")=s_cs("Size")
         if s_cs("WebDesc")<>"" then
         Session("RCM273_CarType")=s_cs("WebDesc")
         end if
         Session("RCM273_CarSizeID")=Left(tidyup(Request.Form("CarSizeID")),3)
       
         Session("RCM273_CarImageName")=s_cs("ImageName")
         'Session("RCM273_Flight")=tidyup(Request.Form("Flight"))
         'Session("RCM273_Flightout")=tidyup(Request.Form("Flightout"))
         'Session("RCM273_NoTravelling")=tidyup(Request.Form("NoTravelling"))

        s_cs.CLOSE
        SET s_cs=nothing

    '---Insert the booking details to WebReservation table

     IF Session("RCM273_OnlineBookingNo") ="" or ISNULL(Session("RCM273_OnlineBookingNo"))=True then
            if Session("RCM273_referral")="" then
                  Session("RCM273_referral")=0
            end if
            DateEntered = Session("RCM273_LocalTime")
            SQL="INSERT INTO WebReservation (BookingDate,Referrals,ReferralID,BrandID,URL,Phone,Email,Name, NoTravelling,Flightout,Flight,PickupDateTime,DropoffDateTime,RentalSource,CarSizeID,PickupTime,Pickupdate,DropoffTime,DropoffDate,PickupLocationID,DropoffLocationID,ReservationTypeID)"
            SQL=SQL&"VALUES ('"&DateEntered&"','"&Left(Session("RCM273_PromoCode"),30)&"','"&tidyup(Session("RCM273_referral"))&"',0,'"&tidyup(Session("RCM273_URL"))&"','"&Left(tidyup(Session("RCM273_phone")),20)&"','"&Left(tidyup(Request.Form("CustomerEmail")),50)&"','"&Left(tidyup(Request.Form("lastname")),30)&"','"&Left(tidyup(Session("RCM273_NoTravelling")),30)&"', '"&Left(tidyup(Session("RCM273_Flightout")),50)&"','"&Left(tidyup(Session("RCM273_Flight")),50)&"','"&Session("RCM273_RequestPickDateTime")&"','"&Session("RCM273_RequestDropDateTime")&"' ,'Online Booking',"&(CInt(Session("RCM273_CarSizeID")))&",'"&Session("RCM273_RequestPickTime")&"','"&Session("RCM273_RequestPickDate")&"','"&Session("RCM273_RequestDropTime")&"', '"&Session("RCM273_RequestDropDate")&"', "&CINT(Session("RCM273_PickupLocationID"))&", "&Session("RCM273_DropoffLocationID")&", '1')"
           'Response.Write("<p>")
          ' Response.Write(SQL)
            webConn.Execute(SQL)

            '---get the ReservationNo
            'Set s_No=webConn.Execute("SELECT Max(ReservationNo)  as ResNo FROM WebReservation where Pickupdate= '"&Session("RCM273_RequestPickDate")&"' and  DropoffLocationID="&CInt(Session("RCM273_DropoffLocationID"))&" ")
             Set s_No=webConn.Execute("SELECT Max(ReservationNo)  as ResNo FROM WebReservation where CarSizeID='"&CInt(Session("RCM273_CarSizeID"))&"' and Dropoffdate= '"&Session("RCM273_RequestDropDate")&"' and  DropoffLocationID="&CInt(Session("RCM273_DropoffLocationID"))&" ")
           Session("RCM273_OnlineBookingNo")=s_No("ResNo")
            'Response.Write(Session("RCM273_OnlineBookingNo"))
            s_No.close
            set s_No=nothing
   ELSE '--IF Session("RCM273_OnlineBookingNo") <>""
         SQL="UPDATE WebReservation SET "

          SQL=SQL&"CarSizeID ="&CInt(Session("RCM273_CarSizeID"))&", "
         SQL=SQL&"Flight ='"&tidyup(Session("RCM273_Flight"))&"', "
         SQL=SQL&"Flightout ='"&tidyup(Session("RCM273_Flightout"))&"', "
         SQL=SQL&"PickupdateTime ='"&Session("RCM273_RequestPickDateTime")&"', "
         SQL=SQL&"DropoffDateTime ='"&Session("RCM273_RequestDropDateTime")&"', "
         SQL=SQL&"Pickupdate ='"&Session("RCM273_RequestPickDate")&"', "
         SQL=SQL&"PickupTime ='"&Session("RCM273_RequestPickTime")&"', "
         SQL=SQL&"DropoffDate ='"&Session("RCM273_RequestDropDate")&"', "
         SQL=SQL&"DropoffTime ='"&Session("RCM273_RequestDropTime")&"', "
         SQL=SQL&"PickupLocationID ="&CINT(Session("RCM273_PickupLocationID"))&", "
         SQL=SQL&"DropoffLocationID ="&CInt(Session("RCM273_DropoffLocationID"))&" "
         SQL=SQL&"WHERE (ReservationNo) ='"&Clng(Session("RCM273_OnlineBookingNo"))&"'"
         'Response.Write(SQL)
         webConn.Execute(SQL)
           '---if reflash the screen, delete the extra and rate then insert again
          'webConn.Execute("DELETE  FROM WebReservationPayment WHERE (ReservationNo)='"&Clng(Session("RCM273_OnlineBookingNo"))&"' ")
          webConn.Execute("DELETE  FROM WebPaymentDetail WHERE (ReservationNo)='"&Clng(Session("RCM273_OnlineBookingNo"))&"' ")
          webConn.Execute("DELETE  FROM WebPaymentExtraFees WHERE (ReservationNo)='"&Clng(Session("RCM273_OnlineBookingNo"))&"' ")

      END IF

        '----Relocation fee-------------
         Session("RCM273_RelocationFee")=0
         Session("RCM273_RelocationFeeID")=0
         Session("RCM273_RelocationFeeName")=""
         RelocationFeeGST=0
         RelocationFeeStampDuty=0
         '--1. check Relocation record (with caterory, date range)
         SQL="SELECT MinBookingDay,DaysNocharge,ExtraFees.ID,Name, Fees, GST,StampDuty FROM WebRelocationFees,ExtraFees "
         SQL=SQL&" WHERE OnewayFeeID=ExtraFees.ID and (CarSizeID)='"&CInt(Session("RCM273_CarSizeID"))&"'  and Mandatory=0 "
         SQL=SQL&"  AND (PickuplocationID)='"&CINT(Session("RCM273_PickupLocationID"))&"' AND (DropofflocationID)='"&CInt(Session("RCM273_DropoffLocationID"))&"'  "
         SQL=SQL&" AND PickupDateFrom<='"&Session("RCM273_RequestPickDate")&"'  and PickupDateTo >= '"&Session("RCM273_RequestPickDate")&"' "
         '--next line of code will return the max Minbookingday if there are 2 records entered for same conditin
         SQL=SQL&"  AND MinBookingDay<="&CInt(Session("RCM273_TotalRentalDays"))&" order by MinBookingDay desc "
         'Response.Write(SQL)
          Set s_o=webConn.Execute(SQL)
          if  s_o.EOF THEN
                  s_o.close
                  Set s_o=Nothing
                  '--2. if no vehicle category Relocation fee found, check Relocation record (with  date range only)
                  SQL="SELECT MinBookingDay,DaysNocharge,ExtraFees.ID,Name, Fees, GST,StampDuty FROM WebRelocationFees,ExtraFees "
                  SQL=SQL&" WHERE OnewayFeeID=ExtraFees.ID  and Mandatory=0  "
                 SQL=SQL&" AND CarSizeID=0 "
                  SQL=SQL&"  AND (PickuplocationID)='"&CINT(Session("RCM273_PickupLocationID"))&"' AND (DropofflocationID)='"&CInt(Session("RCM273_DropoffLocationID"))&"'  "
                  SQL=SQL&" AND PickupDateFrom<='"&Session("RCM273_RequestPickDate")&"'  and PickupDateTo >= '"&Session("RCM273_RequestPickDate")&"' "
                  SQL=SQL&"  AND MinBookingDay<="&CINT(Session("RCM273_TotalRentalDays"))&" order by MinBookingDay desc "
                  'Response.Write(SQL)
                   Set s_o=webConn.Execute(SQL)
                   if  s_o.EOF THEN
                        s_o.close
                        Set s_o=Nothing
                        '--3. if no vehicle category Relocation fee found, check Relocation record (with  catgory, no date ragne)
                        SQL="SELECT MinBookingDay,DaysNocharge,ExtraFees.ID,Name, Fees, GST,StampDuty FROM WebRelocationFees,ExtraFees "
                        SQL=SQL&" WHERE OnewayFeeID=ExtraFees.ID  and (CarSizeID)='"&CInt(Session("RCM273_CarSizeID"))&"' "
                        SQL=SQL&"  AND (PickuplocationID)='"&CINT(Session("RCM273_PickupLocationID"))&"' AND (DropofflocationID)='"&CInt(Session("RCM273_DropoffLocationID"))&"'  "
                        SQL=SQL&"  AND  Year(PickupDateFrom)=2100  and Mandatory=0 "
                        SQL=SQL&"  AND MinBookingDay<="&CINT(Session("RCM273_TotalRentalDays"))&" order by MinBookingDay desc "
                        'Response.Write(SQL)
                        Set s_o=webConn.Execute(SQL)
                        if  s_o.EOF THEN
                              s_o.close
                              Set s_o=Nothing
                                    '--4. if no vehicle category Relocation fee found, check Relocation record (with no catgory, no date ragne)
                              SQL="SELECT MinBookingDay,DaysNocharge,ExtraFees.ID,Name, Fees, GST,StampDuty FROM WebRelocationFees,ExtraFees "
                              SQL=SQL&" WHERE OnewayFeeID=ExtraFees.ID   and Mandatory=0 "
                              SQL=SQL&" AND CarSizeID=0 "
                              SQL=SQL&"  AND (PickuplocationID)='"&CINT(Session("RCM273_PickupLocationID"))&"' AND (DropofflocationID)='"&CINT(Session("RCM273_DropoffLocationID"))&"'  "
                              SQL=SQL&"  AND  Year(PickupDateFrom)=2100 "
                              SQL=SQL&"  AND MinBookingDay<="&CINT(Session("RCM273_TotalRentalDays"))&" order by MinBookingDay desc "
                              'Response.Write(SQL)
                              Set s_o=webConn.Execute(SQL)
                        end if
                  end if   

         end if

         
         webConn.Execute("DELETE  FROM WebPaymentExtraFees WHERE (ReservationNo)='"&Clng(Session("RCM273_OnlineBookingNo"))&"'  ")

         if NOT s_o.EOF THEN 
               DO WHILE NOT s_o.EOF
                  '---Jun 2014, new code, if user go back to open new window with different dates or locations caused sesson max up, we still check the min booking days
                  if s_o("MinBookingDay") > Session("RCM273_TotalRentalDays") then
                        Response.Redirect "webstep2.asp"
                     'Response.Write  Session("RCM273_TotalRentalDays")
                 end if
                  if  s_o("DaysNocharge")=0 or (s_o("DaysNocharge")>0 and Session("RCM273_TotalRentalDays")<s_o("DaysNocharge"))  then
                     Session("RCM273_RelocationFee")=Session("RCM273_RelocationFee")+s_o("Fees")
                     Session("RCM273_RelocationFeeID")=s_o("ID")
                     Session("RCM273_RelocationFeeName")=s_o("Name")
                     if s_o("GST")="True" then
                           RelocationFeeGST=s_o("Fees")
                     end if
                     if s_o("StampDuty")="True" then
                           RelocationFeeStampDuty=s_o("Fees")
                     end if
                 
                        if s_o("Fees")<>0 then
                           SQL="INSERT INTO WebPaymentExtraFees(ExtraValue,ReservationNo,ExtraFeesID,Fees,Days)"
                           SQL=SQL&"VALUES ("&s_o("Fees")&",'"&Clng(Session("RCM273_OnlineBookingNo"))&"','"&s_o("ID")&"','"&s_o("Fees")&"',1)"
                           'Response.Write("<br>") 
                                  'Response.Write(SQL)
                                 webConn.Execute(SQL)
                        end if
                 end if
                s_o.MoveNext
                Loop
         end if
         s_o.CLOSE
         SET s_o=NOTHING

                      '---Aug 2014, new WebExtralAdditionalFees, added for wicked AU, allow to add additional fee for date range
            SQL="SELECT DaysNocharge,ExtraFees.ID,Name, Fees, GST,StampDuty FROM WebExtralAdditionalFees,ExtraFees "
            SQL=SQL&" WHERE  PickuplocationID='"&CINT(Session("RCM273_PickupLocationID"))&"' "
            SQL=SQL&"  AND (DropofflocationID='"&CINT(Session("RCM273_DropoffLocationID"))&"' or DropofflocationID=0) "
            SQL=SQL&" and (CarSizeID='"&Cint(Session("RCM273_CarSizeID"))&"' or CarSizeID=0)  and Mandatory=0 AND ExtraFeeID=ExtraFees.ID "
            SQL=SQL&" AND PickupDateFrom<='"&Session("RCM273_RequestPickDate")&"'  and PickupDateTo >= '"&Session("RCM273_RequestPickDate")&"' "
            SQL=SQL&"  AND (DaysNocharge>"&CINT(Session("RCM273_TotalRentalDays"))&" or DaysNocharge=0 )"
           ' Response.Write SQL
            'Response.Write("<Br><br>")
            Set s_o=webConn.Execute(SQL)
            DO WHILE NOT s_o.EOF
                     if s_o("GST")="True" then
                        RelocationFeeGST=RelocationFeeGST+s_o("Fees")
                     end if
                     if s_o("StampDuty")="True" then
                        RelocationFeeStampDuty=RelocationFeeStampDuty+s_o("StampDuty")
                     end if
                     Session("RCM273_RelocationFee")=Session("RCM273_RelocationFee")+s_o("Fees")
                     if Session("RCM273_RelocationFeeID")=s_o("ID") then
                           ExtraValue=2*s_o("Fees")
                           SQL="update WebPaymentExtraFees set QTY=2, ExtraValue='"&ExtraValue&"' where ReservationNo='"&Clng(Session("RCM273_OnlineBookingNo"))&"' and ExtraFeesID='"&s_o("ID")&"'   "
                           webConn.Execute(SQL)
                     else
                           SQL="INSERT INTO WebPaymentExtraFees(QTY,ExtraValue,ReservationNo,ExtraFeesID,Fees,Days)"
                           SQL=SQL&"VALUES (1, "&s_o("Fees")&",'"&Clng(Session("RCM273_OnlineBookingNo"))&"','"&s_o("ID")&"','"&s_o("Fees")&"',1)"
                           webConn.Execute(SQL)
                      end if
                         'Response.Write("<br>")
                          ' Response.Write(SQL)
            s_o.MoveNext
            Loop
            s_o.CLOSE
            SET s_o=NOTHING


 '-----Pickup Location After hours and befores fee, check if the pickup time is between the office time-------------
        Session("RCM273_AfterHoursFee")=0
        Session("RCM273_PickupAfterHoursFee")=0
        AfterHoursFeeGST=0
        AfterHoursFeeStampDuty=0
        if Session("RCM273_PickupAfterHourFeeID")<>0  then 'check the time
               AfterHoursFee=0
               Set s_st=webConn.Execute("SELECT ID, Name, Fees,GST,StampDuty FROM ExtraFees WHERE (ID)='"&Cint(Session("RCM273_PickupAfterHourFeeID"))&"' ")
               if s_st("Fees")  <>0 then
                  AfterHoursFee=s_st("Fees")
                  AfterHoursGST=s_st("GST")
                  AfterHoursStampDuty=s_st("StampDuty")
                  PickupAfterHoursFeeID=s_st("ID")

               end if
               s_st.close
               SET s_st=nothing
               '---Pickup Location After Hour fees   'do not convert to CDate
                PickupOpeningTime=(Session("RCM273_RequestPickDate")&" "&Session("RCM273_PickupOfficeOpeningTime"))
               PickupClosingTime=(Session("RCM273_RequestPickDate")&" "&Session("RCM273_PickupOfficeClosingTime"))
                if (Session("RCM273_RequestPickDateTime") < PickupOpeningTime) or (Session("RCM273_RequestPickDateTime") > PickupClosingTime) THEN
                        Session("RCM273_PickupAfterHoursFee")=AfterHoursFee

                        if AfterHoursGST="True" then
                              AfterHoursFeeGST=Session("RCM273_PickupAfterHoursFee")
                        end if
                        if AfterHoursStampDuty="True" then
                              AfterHoursFeeStampDuty=Session("RCM273_PickupAfterHoursFee")
                        end if
                end if
      end if

     '------------Dropoff Location After Hour Fees ------------------
    Session("RCM273_DropoffAfterHoursFee")=0
    if Session("RCM273_DropoffAfterHourFeeID")<>0  then
           Set s_st=webConn.Execute("SELECT ID, Name, Fees,GST,StampDuty FROM ExtraFees WHERE (ID)='"&Cint(Session("RCM273_DropoffAfterHourFeeID"))&"' ")
           if s_st("Fees")  <>0 then
               AfterHoursFee=s_st("Fees")
               AfterHoursGST=s_st("GST")
               AfterHoursStampDuty=s_st("StampDuty")
               DropoffAfterHoursFeeID=s_st("ID")

            end if
            s_st.close
            SET s_st=nothing
            DropoffOpeningTime=(Session("RCM273_RequestDropDate")&" "&Session("RCM273_DropoffOfficeOpeningTime"))
            DropoffClosingTime=(Session("RCM273_RequestDropDate")&" "&Session("RCM273_DropoffOfficeClosingTime"))
            if (Session("RCM273_RequestDropDateTime") < DropoffOpeningTime) or (Session("RCM273_RequestDropDateTime") > DropoffClosingTime) THEN
                  Session("RCM273_DropoffAfterHoursFee")=  AfterHoursFee
                  if AfterHoursGST="True" then
                        AfterHoursFeeGST=AfterHoursFeeGST+Session("RCM273_DropoffAfterHoursFee")
                  end if
                  if AfterHoursStampDuty="True" then
                        AfterHoursFeeStampDuty=AfterHoursFeeStampDuty+Session("RCM273_DropoffAfterHoursFee")
                  end if
            end if
      end if
      ' AfterHoursFee=   Session("RCM273_PickupAfterHoursFee")+Session("RCM273_DropoffAfterHoursFee") '---, for insert to table, 9 Jun  2012 removed to below
     ' Session("RCM273_AfterHoursFee")=AfterHoursFee
      Session("RCM273_AfterHoursFee") =  Session("RCM273_PickupAfterHoursFee")+Session("RCM273_DropoffAfterHoursFee")'--for total cost
      if Session("RCM273_PickupAfterHoursFee")>0  and Session("RCM273_DropoffAfterHoursFee")=0 then
                 '---insert AfterHours Fee to WebPaymentExtraFees table
                  SQL="INSERT INTO WebPaymentExtraFees(ExtraValue,ReservationNo,ExtraFeesID,Fees,Days,QTY)"
                  SQL=SQL&"VALUES ("&Session("RCM273_PickupAfterHoursFee")&",'"&Clng(Session("RCM273_OnlineBookingNo"))&"',"&PickupAfterHoursFeeID&","&Session("RCM273_PickupAfterHoursFee")&",1,1)"
                  'Response.Write("<p>")
                  'Response.Write(SQL)
                  webConn.Execute(SQL)

        elseif Session("RCM273_PickupAfterHoursFee")=0 and Session("RCM273_DropoffAfterHoursFee")>0 then
                  '---insert AfterHours Fee to WebPaymentExtraFees table
                  SQL="INSERT INTO WebPaymentExtraFees(ExtraValue,ReservationNo,ExtraFeesID,Fees,Days,QTY)"
                  SQL=SQL&"VALUES ("&Session("RCM273_DropoffAfterHoursFee")&",'"&Clng(Session("RCM273_OnlineBookingNo"))&"',"&DropoffAfterHoursFeeID&","&Session("RCM273_DropoffAfterHoursFee")&",1,1)"
               '  Response.Write("<p>")
                ' Response.Write(SQL)
                  webConn.Execute(SQL)
        elseif Session("RCM273_PickupAfterHoursFee")>0 and Session("RCM273_DropoffAfterHoursFee")>0 then
               if PickupAfterHoursFeeID=DropoffAfterHoursFeeID then
                  AfterHoursFee=   Session("RCM273_PickupAfterHoursFee")+Session("RCM273_DropoffAfterHoursFee")
                  SQL="INSERT INTO WebPaymentExtraFees(ExtraValue,ReservationNo,ExtraFeesID,Fees,Days,QTY)"
                  SQL=SQL&"VALUES ("&AfterHoursFee&",'"&Clng(Session("RCM273_OnlineBookingNo"))&"',"&PickupAfterHoursFeeID&","&Session("RCM273_PickupAfterHoursFee")&",1,2)"
               '  Response.Write("<p>")
                ' Response.Write(SQL)
                  webConn.Execute(SQL)
                else
                     SQL="INSERT INTO WebPaymentExtraFees(ExtraValue,ReservationNo,ExtraFeesID,Fees,Days,QTY)"
                     SQL=SQL&"VALUES ("&Session("RCM273_PickupAfterHoursFee")&",'"&Clng(Session("RCM273_OnlineBookingNo"))&"',"&PickupAfterHoursFeeID&","&Session("RCM273_PickupAfterHoursFee")&",1,1)"
                      ' Response.Write("<p>")
                       'Response.Write(SQL)
                     webConn.Execute(SQL)
                    SQL="INSERT INTO WebPaymentExtraFees(ExtraValue,ReservationNo,ExtraFeesID,Fees,Days,QTY)"
                    '--bug fixSession("RCM168_PickupAfterHoursFee") changed Session("RCM273_DropoffAfterHoursFee")
                     SQL=SQL&"VALUES ("&Session("RCM273_DropoffAfterHoursFee")&",'"&Clng(Session("RCM273_OnlineBookingNo"))&"',"&DropoffAfterHoursFeeID&","&Session("RCM273_DropoffAfterHoursFee")&",1,1)"
                     '  Response.Write("<p>")
                     ' Response.Write(SQL)
                   webConn.Execute(SQL)
                end if
        end if



'----GST Inclusive--------------------------------
   Set s_st=webConn.Execute("SELECT * FROM SystemTable WHERE Code='GSINC'  ")
   Session("RCM273_GSTInclusive") = "Yes"
   If not s_st.EOF then
            Session("RCM273_GSTInclusive") = s_st("syValue")
   END IF
   s_st.CLOSE
   SET s_st=NOTHING
      

END IF
'--------END if Request.QueryString("dir")="Rate" ----------------------------------------------------------------------------------

   if Request.QueryString("dir")="existingCust"   then

         Session("RCM273_ccNumber")=""

         LastName=Left(tidyup(Request.Form("LastName")),40)
         CustomerEmail=Left(tidyup(Request.Form("CustomerEmail")),80)
         License=Left(tidyup(Request.Form("License")),30)

         SQL="SELECT Customers.*  from Customers where  acLastName='"&tidyup(LastName)&"' and acEmail='"&tidyup(CustomerEmail)&"' "
         'SQL="SELECT Customers.*  from Customers where  acLastName='"&tidyup(LastName)&"' and acLicense='"&License&"' "

         'Response.Write(SQL)
         Set s_b=WebConn.Execute(SQL)
         if  s_b.EOF  then
                        Session("RCM273_ErrorMesage")="*  Sorry we can not find your record. Please enter your details here"
         else
                        Session("RCM273_CustomerID")=s_b("acID")
                        Session("RCM273_FirstName")=s_b("acfirstname")
                        Session("RCM273_LastName")=s_b("acLastName")
                        Session("RCM273_License")=s_b("acLicense")
                        Session("RCM273_LicenseIssued")=s_b("acLicenseIssued")
                        Session("RCM273_LicExpDate") = s_b("acLicenseExpiry")

                        Session("RCM273_address")=s_b("acPostalAddress")
                        Session("RCM273_city")=s_b("acCity")
                        Session("RCM273_state")=s_b("acstate")
                        Session("RCM273_postcode")=s_b("acpostcode")
                        Session("RCM273_countryID")=s_b("acCountryID")
                        Session("RCM273_CustomerEmail")=s_b("acEmail")
                        Session("RCM273_phone")=s_b("acPhone")
                        Session("RCM273_fax")=s_b("acFax")
                        Session("RCM273_DOB")=s_b("acDOB")


         end if
         s_b.close
         Set s_b=nothing

         ' Response.Write(Session("RCM273_CustomerID"))
     end if


    %>
 <div class="container">
        <h4 class="smallm_title centered bigger"><span>Your quote details</span></h4>
		
		
    <%
        theimage=RCMURL&"/DB/"&CompanyCode&"/"&Session("RCM273_CarImageName")

         Response.Write("<div class='row'><div class='col-xs-5'> <p class='text-center'><img src='"&theimage&"'  class='vehicle-image' alt='' /></p></div>")

   '---Vehicle Type----------
         Response.Write("<div class='col-xs-7'><h2 style='text-align: left;'>"&Session("RCM273_CarType")&"</h2><hr />")
      '---Vehicle Type----------


                        
                        Response.Write("<div class='row'><div class='col-xs-2'><p class='text-left'><strong>Pickup:</strong></p></div>")
                        Response.Write("<div class='col-xs-3'><p class='text-right'>")
                        Response.Write(Session("RCM273_PickupLocation"))
                        Response.Write("</div>")
                       Response.Write("<div class='col-xs-1'>&nbsp;</div>")

                    
                        Response.Write("<div class='col-xs-2'><p class='text-left'><strong>Pickup Date:</strong></p></div>")
                        Response.Write("<div class='col-xs-4'><p class='text-right'>"&WeekdayName(WeekDay(Session("RCM273_RequestPickDate")))&", ")
                        'Response.Write(Session("RCM273_RequestPickDateTime"))
                        Response.Write(Day(Session("RCM273_RequestPickDate"))&"/"&Left(MonthName(Month(Session("RCM273_RequestPickDate"))),3)&"/"&Year(Session("RCM273_RequestPickDate")))
                        Response.Write("&nbsp;")
                        Response.Write(Session("RCM273_RequestPickTime"))
                        Response.Write("</p></div></div>")
                        
                        if Session("RCM273_Flight")<>"" then
                        Response.Write("<tr>")
                        Response.Write("<td></td><td align='left' class='text'>Arrival Details: </td>")
                        Response.Write("<td align='left' class='text' colspan='2'>")
                        Response.Write(Session("RCM273_Flight"))
                        Response.Write("</td>")
                       Response.Write("</tr>")
                        end if
                        
                  
                        Response.Write("<div class='row'><div class='col-xs-2'><p class='text-left'><strong>Return:</strong></p></div>")
                        Response.Write("<div class='col-xs-3'><p class='text-right'>")
                        Response.Write(Session("RCM273_DropoffLocation"))
                        Response.Write("</p></div>")
                       Response.Write("<div class='col-xs-1'>&nbsp;</div>")

                        
                        Response.Write("<div class='col-xs-2'><p class='text-left'><strong>Return Date:</strong></p></div>")
                        Response.Write("<div class='col-xs-4'><p class='text-right'>"&WeekdayName(WeekDay(Session("RCM273_RequestDropDate")))&", ")
                         'Response.Write(Session("RCM273_RequestDropDateTime"))
                        Response.Write(Day(Session("RCM273_RequestDropDate"))&"/"&Left(MonthName(Month(Session("RCM273_RequestDropDate"))),3)&"/"&Year(Session("RCM273_RequestDropDate")))
                        Response.Write("&nbsp;")
                        Response.Write(Session("RCM273_RequestDropTime"))
                       
                        Response.Write("</p></div>")
                      
                      
                       if Session("RCM273_Flightout")<>"" then
                        Response.Write("<tr>")
                        Response.Write("<td></td><td align='left' class='text'>Departure Details:</td>")
                        Response.Write("<td align='left' class='text' colspan='2'>")
                        Response.Write(Session("RCM273_Flightout"))
                        Response.Write("</td>")
                       Response.Write("</tr>")
                       end if 
                       
                        if Session("RCM273_NoTravelling")<>"" then
                        'Response.Write("<tr>")
                        Response.Write("<div class='col-xs-3'><p class='text-left'><strong>People Travelling:</strong></p></div>")
                        Response.Write("<div class='col-xs-2'><p class='text-right'>")
                        Response.Write(Session("RCM273_NoTravelling"))
                        Response.Write("</p></div>")
                       
                       end if 
                       Response.Write("</div><hr />")

        

   '-------------if Request.QueryString("dir")="Rate" -----------------
    RentalCost=0
    if Request.QueryString("dir")="Rate"  then
           '-------------  the rate--------------
           webConn.Execute("DELETE  FROM WebPaymentDetail WHERE (ReservationNo)='"&Clng(Session("RCM273_OnlineBookingNo"))&"' ")
          for i=1 to Request.Form("SeasonCount")
                  SeasonID=Request.Form("SeasonID"&i&"")
                  NoOfDaysEachSeason=0
                  NoHours=0
                  if Request.Form("NoOfDaysEachSeason"&i&"")  <>"" then
                     NoOfDaysEachSeason=Request.Form("NoOfDaysEachSeason"&i&"")
                  end if
                  if Request.Form("lessDayNumberOfHours")<>"0" and Request.Form("lessDayNumberOfHours")<>"" then
                     NoOfDaysEachSeason=Request.Form("NoOfDaysEachSeason")
                     NoHours=Request.Form("lessDayNumberOfHours")

                  end if
                  TotalNoDays=TotalNoDays+NoOfDaysEachSeason
                  StandardRate=0
                Rate=0
                StandardRate=Request.Form("StandardRate"&i&"")
                Rate=Request.Form("Rate"&i&"")
                costEachSeason=Rate*NoOfDaysEachSeason
                RentalCost=RentalCost+costEachSeason

                  SQL="INSERT INTO WebPaymentDetail (NoHours,DiscountID,DiscountName,DiscountType,DiscountPerc,StandardRate,Rate,ReservationNo,SeasonID,Days,RateName)"
                  SQL=SQL&"VALUES ("&NoHours&","&Session("RCM273_DiscountID")&",'"&Left(tidyup(Session("RCM273_DiscountName")),50)&"','"&(Session("RCM273_DiscountType"))&"',"&Session("RCM273_DiscountRate")&",'"&StandardRate&"','"&Rate&"','"&Clng(Session("RCM273_OnlineBookingNo"))&"','"&Request.Form("SeasonID"&i&"")&"','"&NoOfDaysEachSeason&"','"&Left(Request.Form("RateName"&i&""),50)&"')"
                  'Response.Write("<br>")
                  'Response.Write(SQL)
                  webConn.Execute(SQL)
        NEXT  '-------------  end the rate--------------


       '--------FreeDaysOffer-------------------
      FreeDaysValue=0
      if Request.Form("NoFreeDays")<>0 then
            FreeDayRate=-1*Request.Form("FreeDayRate")
               FreeDaysValue=Request.Form("NoFreeDays")*FreeDayRate
               SQL="INSERT INTO WebPaymentExtraFees(QTY,ExtraValue,ReservationNo,ExtraFeesID,Fees,Days)"
            SQL=SQL&"VALUES ('"&Request.Form("NoFreeDays")&"',"&FreeDaysValue&",'"&Clng(Session("RCM273_OnlineBookingNo"))&"','"&Request.Form("FreeDayExtraFeeID")&"','"&FreeDayRate&"',1)"
            'Response.Write("<br>")
                  'Response.Write(SQL)
                  webConn.Execute(SQL)

       end if
      
            '-----FixedDiscountRate------------
         FixedDiscountRate=0
      if Request.Form("FixedDiscountID")<>"0" and Request.Form("FixedDiscountID")<>"" then
               FixedDiscountRate=Request.Form("FixedDiscountRate")
               RentalCost=RentalCost+ FixedDiscountRate
               SQL="INSERT INTO WebPaymentExtraFees(QTY,ExtraValue,ReservationNo,ExtraFeesID,Fees,Days)"
               SQL=SQL&"VALUES (1,"&FixedDiscountRate&",'"&Clng(Session("RCM273_OnlineBookingNo"))&"','"&Request.Form("FixedDiscountID")&"','"&FixedDiscountRate&"',1)"
               'Response.Write("<br>")
                  'Response.Write(SQL)
                  webConn.Execute(SQL)
      end if

        TotalCost=RentalCost+ FreeDaysValue+ Session("RCM273_RelocationFee")+Session("RCM273_AfterHoursFee")
    
      'Response.Write(RentalCost)

      '-------------Selected Extra fees-----------------
     EachExtraFees=0
     TotalExtraFees=0
     ExtraStamp=0
     ExtraGST=0
     k=0
     For j=0 to CInt(Request.Form("ExtraFeeCount"))
            if Request.Form("ExtraFeesID"&j&"")<>"" then
                    SetMaxPrice=0
                    if Request.Form("ExtraFees"&j&"")<>"" then
                        ExtraFeesRate=Request.Form("ExtraFees"&j&"")
                    end if  
                    
                    if Request.Form("FeeType"&j&"")="Daily" then
                              if CCur(Request.Form("Maxprice"&j&""))>0 and Session("RCM273_TotalRentalDays")*Request.Form("ExtraFees"&j&"")>CCur(Request.Form("Maxprice"&j&"")) then
                                          EachExtraFees=Request.Form("Maxprice"&j&"")*Request.Form("QTY"&j&"")
                                       ExtraFeesRate=Request.Form("Maxprice"&j&"")
                                       ExtraFeesNoDays=1
                                       ExtraType="Fixed" 
                                       SetMaxPrice=1
                                 else 
                                       EachExtraFees=Session("RCM273_TotalRentalDays")*Request.Form("ExtraFees"&j&"")*Request.Form("QTY"&j&"")
                                       ExtraFeesNoDays=Session("RCM273_TotalRentalDays")
                                       ExtraType="Daily" 
                                   end if
                       elseif Request.Form("FeeType"&j&"")="Fixed" then
                                 EachExtraFees = Request.Form("ExtraFees"&j&"")*Request.Form("QTY"&j&"")
                                 ExtraFeesNoDays=1
                                 ExtraType="Fixed"

                      else
                                 EachExtraFees =(Request.Form("ExtraFees"&j&"")/100)*RentalCost* Request.Form("QTY"&j&"")
                                    '--27/Jun/2009 for % fees allow check max price
                                    if Request.Form("MaxPrice"&j&"") <>"0"  then
                                          if EachExtraFees > CCur(Request.Form("MaxPrice"&j&"") ) and CCur(Request.Form("MaxPrice"&j&"") )>0 then
                                                EachExtraFees=Request.Form("MaxPrice"&j&"")
                                          end if
                                           '--for discount, if set to 10% off and max -$30
                                          if CCur(Request.Form("MaxPrice"&j&"") )< 0 and  EachExtraFees<CCur(Request.Form("MaxPrice"&j&"")) then
                                             EachExtraFees=Request.Form("MaxPrice"&j&"")
                                          end if
                                    end if
                                 ExtraFeesNoDays=1
                                 ExtraType="Perscentage"

                       end if

                        TotalExtraFees=EachExtraFees+ TotalExtraFees

                        '----Stamp Duty for extraFees ---------
                        eachExtraStamp=0
                        if Request.Form("StampDuty"&j&"")="True" Then
                                  eachExtraStamp=EachExtraFees
                        end if
                        ExtraStamp=ExtraStamp+eachExtraStamp
                              
                           '----GST for extraFees ---------
                           eachExtraGST=0
                           if Request.Form("GST"&j&"")="True" Then
                                 eachExtraGST=EachExtraFees
                           end if
                           ExtraGST=ExtraGST+eachExtraGST    
                         '--insert the selected Extra Fees to WebPaymentExtraFees table
                        SQL="INSERT INTO WebPaymentExtraFees(ExtraValue,SetMaxPrice,QTY, ReservationNo,ExtraFeesID,Fees,Days)"
                        SQL=SQL&"VALUES ("&EachExtraFees&","&SetMaxPrice&",'"&Request.Form("QTY"&j&"")&"','"&Clng(Session("RCM273_OnlineBookingNo"))&"','"&Request.Form("ExtraFeesID"&j&"")&"','"&ExtraFeesRate&"','"&ExtraFeesNoDays&"')"
                
                        webConn.Execute(SQL)
                        
            end if
     NEXT
     TotalCost=TotalCost+TotalExtraFees
     



     '-----Insurance Fee------------    
        InsuranceExtraValue=0
         
        GSTInsurance=0
        StampInsurance=0
         if Request.Form("InsuranceID")<>"" then
                 Session("RCM273_InsuranceID")=tidyup(Request.Form("InsuranceID"))
                SQL="SELECT * from ExtraFees WHERE (ID)='"&Cint(Request.Form("InsuranceID"))&"' "
                 Set s_insu=webConn.Execute(SQL)
               if NOT s_insu.EOF  then
                   
                     if s_insu("Type")="Daily" then
                           InsuranceExtraValue=Session("RCM273_TotalRentalDays")*s_insu("Fees")
                     elseif s_insu("Type")="Fixed" then
                           InsuranceExtraValue=s_insu("Fees")
                     else
                        InsuranceExtraValue=(s_insu("Fees")/100)*RentalCost
                     end if   
                     InsuranceSetMaxPrice =0   
                        InsuranceRate=s_insu("Fees")
                  
                     InsuranceNoDays=Session("RCM273_TotalRentalDays")
                      if s_insu("MaxPrice") >0  and  InsuranceExtraValue>s_insu("MaxPrice") then
                              InsuranceExtraValue=s_insu("MaxPrice")
                              InsuranceSetMaxPrice =1 
                              InsuranceNoDays=1
                              InsuranceRate =s_insu("MaxPrice")
                     end if
                      if s_insu("GST")="True" Then
                              GSTInsurance=InsuranceExtraValue
                         end if 
                          if s_insu("StampDuty")="True" Then
                                  StampInsurance=InsuranceExtraValue
                        end if
                         SQL="INSERT INTO WebPaymentExtraFees(ExtraValue,SetMaxPrice,QTY, ReservationNo,ExtraFeesID,Fees,Days)"
                         SQL=SQL&"VALUES ("&InsuranceExtraValue&","&InsuranceSetMaxPrice&",1,'"&Clng(Session("RCM273_OnlineBookingNo"))&"',"&s_insu("ID")&","&InsuranceRate&",'"&InsuranceNoDays&"')"
                         'Response.Write("<br>")
                         'Response.Write(SQL)
                     webConn.Execute(SQL) 
               end if
               s_insu.close
               SET s_insu=Nothing
         end if
           
      TotalCost=TotalCost+ InsuranceExtraValue
   
        
   
         '-------------KmsFree per day -----------------
             KmsMaxcharge=0
        KmsDailyRate=0
        AddKmsFee=0 
        KmsFree=0 
        Session("RCM273_KmsFree") =0
        Session("RCM273_AddKmsFee") =0
          Session("RCM273_KmsCost") =0
           if Request.Form("KmsFreeID")<>"" then
                   KmsReSelectedInSubs
         end if
              
                  '----------------check Holiday Charges for pickup  from table HolidayextraFees  ---------------------
         SQL="SELECT ExtraFees.* FROM ExtraFees4Holiday,ExtraFees WHERE eh_ExtraFeeID=ExtraFees.ID and (eh_HolidayDate='"&Session("RCM273_RequestPickDate")&"' and  eh_LocationID='"&CINT(Session("RCM273_PickupLocationID"))&"')  "
         'Response.Write(SQL)
       Set s_ex=webConn.Execute(SQL)
       HolidayPickupExtraID=0 
       HolidayChargeQty=0
       if NOT s_ex.EOF THEN
               HolidayPickupExtraID=s_ex("ID")
               HolidayChargeQty=1
               SQL="INSERT INTO WebPaymentExtraFees(ExtraValue,ReservationNo,ExtraFeesID,Fees,Days)"
               SQL=SQL&"VALUES ("&s_ex("Fees")&",'"&Clng(Session("RCM273_OnlineBookingNo"))&"','"&s_ex("ID")&"','"&s_ex("Fees")&"',1)"
                'Response.Write("<br>") 
                'Response.Write(SQL)
               webConn.Execute(SQL)
               TotalCost=TotalCost+ s_ex("Fees")
      end if
      s_ex.close
      Set s_ex=Nothing

      '--check Holiday Charges for dropoff  from table HolidayextraFees
       SQL="SELECT ExtraFees.*  FROM ExtraFees4Holiday,ExtraFees WHERE   eh_ExtraFeeID=ExtraFees.ID and (eh_HolidayDate='"&Session("RCM273_RequestdropDate")&"' and  eh_LocationID='"&CInt(Session("RCM273_DropoffLocationID"))&"') "
       'Response.Write(SQL)
       Set s_ex=webConn.Execute(SQL)
       HolidayDropoffExtraID=0
       if NOT s_ex.EOF THEN
              HolidayDropoffExtraID=s_ex("ID")
               HolidayChargeQty=1
               if  HolidayDropoffExtraID=HolidayPickupExtraID  then
                  HolidayChargeQty=2
                  webConn.Execute("DELETE  FROM WebPaymentExtraFees WHERE (ReservationNo)='"&Clng(Session("RCM273_OnlineBookingNo"))&"' and ExtraFeesID='"&s_ex("ID")&"' ")

               end if
               SQL="INSERT INTO WebPaymentExtraFees(ExtraValue,ReservationNo,ExtraFeesID,Fees,Days,QTY)"
               SQL=SQL&"VALUES ("&(s_ex("Fees")*HolidayChargeQty)&",'"&Clng(Session("RCM273_OnlineBookingNo"))&"','"&s_ex("ID")&"','"&s_ex("Fees")&"',1,'"&HolidayChargeQty&"')"
                'Response.Write("<br>")
                'Response.Write(SQL)
                  TotalCost=TotalCost+ s_ex("Fees")
               webConn.Execute(SQL)
      end if
      s_ex.close
      Set s_ex=Nothing
          '----------------END check Holiday Charges for pickup  from table HolidayextraFees  ---------------------


       '-------------Web Mandatory Extra fees-----------------
   '---should not include InsuranceExtra (mandatory as default)
       SQL="SELECT * from ExtraFees WHERE (LocationID='"&CINT(Session("RCM273_PickupLocationID"))&"'  or LocationID=0) "
       SQL=SQL&" AND (VehicleSizeID='"&CInt(Session("RCM273_CarSizeID"))&"' or VehicleSizeID=0) "
       SQL=SQL&" and (CategoryTypeID ="&CINT(Session("RCM273_CategoryTypeID"))&" or CategoryTypeID =0 ) "
       SQL=SQL&" and WebItems=1 and InsuranceExtra=0  and Mandatory=1 AND inUse=1 and PercentageTotalCost=0 and MerchantFee=0 order by Type,Name"
     Set s_ex=webConn.Execute(SQL)
     j=0
     EachMandatoryExtraFees=0
     TotalMandatoryExtraFees=0
     MandatoryExtraStamp=0
     MandatoryExtraGST=0
     MandatoryExtraFeesNoDays=1
     BondAmount=0
     DO WHILE NOT s_ex.EOF
            MandatoryExtraFeesNoDays=1

             if s_ex("Type")="Daily" then

                  '---daily fees always use   Session("RCM273_TotalRentalDays") (fixed number)
                        MandatoryExtraFeesNoDays=Session("RCM273_TotalRentalDays")
                        EachMandatoryExtraFees=MandatoryExtraFeesNoDays*s_ex("Fees")
             elseif s_ex("Type")="Fixed" then
                        EachMandatoryExtraFees = s_ex("Fees")
             else
                       EachMandatoryExtraFees = (s_ex("Fees")/100)*RentalCost
             end if
                     '--27/Jun/2009 for Daily and % fees allow check max price in step4
                if s_ex("MaxPrice") >0  and  EachMandatoryExtraFees>s_ex("MaxPrice") then
                                 EachMandatoryExtraFees=s_ex("MaxPrice")
                end if

                if s_ex("MaxPrice")< 0 and EachMandatoryExtraFees<s_ex("MaxPrice") then
                                  EachMandatoryExtraFees=s_ex("MaxPrice")
                end if

                     TotalMandatoryExtraFees=EachMandatoryExtraFees+ TotalMandatoryExtraFees

                     if s_ex("Bond")=True then
                           BondAmount=s_ex("Fees")
                     end if
                 '----Stamp Duty for Mandatory extraFees ---------
                  eachMandatoryExtraStamp=0
                  if s_ex("StampDuty")="True" Then
                        eachMandatoryExtraStamp=EachMandatoryExtraFees
                 end if
                 MandatoryExtraStamp=MandatoryExtraStamp+eachMandatoryExtraStamp

                 '----GST for Mandatory extraFees ---------
                  eachMandatoryExtraGST=0
                  if s_ex("GST")="True" Then
                        eachMandatoryExtraGST=EachMandatoryExtraFees
                  end if
                 MandatoryExtraGST=MandatoryExtraGST+eachMandatoryExtraGST


                '--insert the Mandatory Extra Fees to WebPaymentExtraFees table

                  SQL="INSERT INTO WebPaymentExtraFees(ExtraValue,ReservationNo,ExtraFeesID,Fees,Days)"
                  SQL=SQL&"VALUES ("&EachMandatoryExtraFees&",'"&Clng(Session("RCM273_OnlineBookingNo"))&"','"&s_ex("ID")&"','"&s_ex("Fees")&"','"&MandatoryExtraFeesNoDays&"')"
                  'Response.Write("<br>")
                  'Response.Write(SQL)
                  webConn.Execute(SQL)

      s_ex.MoveNext
     j=j+1
     Loop
     s_ex.close
     SET s_ex=nothing
     TotalCost=TotalCost+TotalMandatoryExtraFees




'---Stamp Duty, GST  are location(state) based,  check if the extra fees are StampDuty, GST Apply
   StampDutyRate=0
   GSTRate= 0
   Set s_st=webConn.Execute("SELECT * FROM QStampDuty WHERE ID="&CINT(Session("RCM273_PickupLocationID"))&"  ")
   If not s_st.EOF then
           if ISNull(s_st("StampDuty"))<>True then
         StampDutyRate=s_st("StampDuty")/100
         end if
         if ISNull(s_st("GSTRate"))<>True then
         GSTRate= s_st("GSTRate")/100
         end if
         Session("RCM273_TaxName1") = s_st("SalesTax")
         Session("RCM273_TaxName2") = s_st("LocalTaxDesc")
  end if
  s_st.close
  set s_st=nothing


'----GST calculation------------------------------------
 GST=0 
 Session("RCM273_GST")=0

'--insert GST and StampDuty to WebReservation table
              
        '----STAMPDUTY GST diferent layout------------------------------------
         SubTotal=0
         StampDuty=0
         Session("RCM273_StampDuty")=0
  IF Session("RCM273_GSTInclusive") = "Yes"  then
         if StampDutyRate>0 then  '--if StampDutyRate=0 then no Stampduty applied
                 StampDuty=StampDutyRate*(RentalCost+ Session("RCM273_KmsCost")+ RelocationFeeStampDuty + AfterHoursFeeStampDuty + ExtraStamp + MandatoryextraStamp +StampInsurance)*(1-GSTRate)
                  StampDuty=Round(StampDuty,2)
                   Session("RCM273_StampDuty")=StampDuty
         end if
         TotalCost=StampDuty+ TotalCost
         GST=(RentalCost + Session("RCM273_KmsCost") + AfterHoursFeeGST + RelocationFeeGST + ExtraGST+ MandatoryExtraGST+ GSTInsurance)*(1-100/(100+GSTRate*100))
         GST=Round(GST,2)
         Session("RCM273_GST")=GST
                            '-------------Web Mandatory MerchantFee or airport fees-----------------
         SQL="SELECT * from ExtraFees WHERE (LocationID='"&CINT(Session("RCM273_PickupLocationID"))&"'  or LocationID=0) "
          SQL=SQL&" AND (VehicleSizeID='"&CInt(Session("RCM273_CarSizeID"))&"' or VehicleSizeID=0) "
            SQL=SQL&" and (CategoryTypeID ="&CINT(Session("RCM273_CategoryTypeID"))&" or CategoryTypeID =0 ) "
          SQL=SQL&" and WebItems=1 and Mandatory=1 AND inUse=1 and (PercentageTotalCost=1 or MerchantFee=1) and Type='Percentage' Order by PercentageTotalCost Desc,MerchantFee,Name"
         'Response.Write(SQL)
         Set s_ex=webConn.Execute(SQL)
         j=0
         EachMerchantFee=0
         TotalMerchantFee=0
         MerchantFeeStamp=0
         MerchantFeeNoDays=1
         BankBaseCalculationFee= TotalCost - BondAmount
         percentageBaseCalculationFee= TotalCost - BondAmount
         DO WHILE NOT s_ex.EOF
                  if s_ex("MerchantFee")=True then
                         EachMerchantFee =(s_ex("Fees")/100)*BankBaseCalculationFee
                  else
                        EachMerchantFee =(s_ex("Fees")/100)*percentageBaseCalculationFee
                  end if
                  '--27/Jun/2009 for % fees allow check max price in step4
                   if s_ex("MaxPrice") >0  and  EachMerchantFee>s_ex("MaxPrice") then
                         EachMerchantFee=s_ex("MaxPrice")
                end if

                if s_ex("MaxPrice")< 0 and  EachMerchantFee<s_ex("MaxPrice") then
                           EachMerchantFee=s_ex("MaxPrice")
                end if
                 EachMerchantFee =Round(EachMerchantFee,2)
                  TotalCost=TotalCost+EachMerchantFee
                  BankBaseCalculationFee=BankBaseCalculationFee+EachMerchantFee

                      '--insert the
                       SQL="INSERT INTO WebPaymentExtraFees(ExtraValue,ReservationNo,ExtraFeesID,Fees,Days)"
                     SQL=SQL&"VALUES ("&EachMerchantFee&","&Clng(Session("RCM273_OnlineBookingNo"))&","&s_ex("ID")&","&s_ex("Fees")&",'1')"
                     '  Response.Write("<br>")
                     'Response.Write(SQL)
                     webConn.Execute(SQL)
               
          s_ex.MoveNext
         j=j+1
         Loop
         s_ex.close
         SET s_ex=nothing
         TotalCostTotalCost=TotalCost
       
       
 ELSE '---GST exClusive----------------

         'TotalCost=StampDuty+ TotalCost
         '-------------Web Mandatory MerchantFee or airport fees-----------------
         SQL="SELECT * from ExtraFees WHERE (LocationID='"&CINT(Session("RCM273_PickupLocationID"))&"'  or LocationID=0) "
         SQL=SQL&" AND (VehicleSizeID='"&CInt(Session("RCM273_CarSizeID"))&"' or VehicleSizeID=0) "
           SQL=SQL&" and (CategoryTypeID ="&CINT(Session("RCM273_CategoryTypeID"))&" or CategoryTypeID =0 ) "
         SQL=SQL&" and WebItems=1 and Mandatory=1 AND inUse=1 and (PercentageTotalCost=1 or MerchantFee=1) and Type='Percentage' Order by PercentageTotalCost Desc,MerchantFee,Name"
         'Response.Write(SQL)
         Set s_ex=webConn.Execute(SQL)
         j=0
         EachMerchantFee=0
         TotalMerchantFee=0
         MerchantFeeStamp=0

         MerchantFeeNoDays=1
         BankBaseCalculationFee= TotalCost
           percentageBaseCalculationFee= TotalCost

         'Response.Write(BankBaseCalculationFee)
            ' Response.Write("<tr><td colspan='4' >"&rentalCost&"+"&TotalExtraFees&" +"&TotalMandatoryExtraFees&" + "&InsuranceFee&" + "&StampDuty&"</td></tr>")
                MerchantFeeGST=0
               MerchantFeeStamp=0
         '--TotalCost =rentalCost + extraFees + Insurance + Stamp
         DO WHILE NOT s_ex.EOF
                  if s_ex("MerchantFee")=True then
                         EachMerchantFee =(s_ex("Fees")/100)*BankBaseCalculationFee
                  else
                      EachMerchantFee =(s_ex("Fees")/100)*percentageBaseCalculationFee
                  end if
                 '--27/Jun/2009 for % fees allow check max price in step4
                   if s_ex("MaxPrice") >0  and  EachMerchantFee>s_ex("MaxPrice") then
                         EachMerchantFee=s_ex("MaxPrice")
                end if

                if s_ex("MaxPrice")< 0 and   EachMerchantFee<s_ex("MaxPrice") then
                           EachMerchantFee=s_ex("MaxPrice")
                end if
                 EachMerchantFee =Round(EachMerchantFee,2)
                  TotalCost=TotalCost+EachMerchantFee
                  TotalMerchantFee=TotalMerchantFee+EachMerchantFee
                  BankBaseCalculationFee=BankBaseCalculationFee+EachMerchantFee
                  if s_ex("GST")="True" then
                            MerchantFeeGST=MerchantFeeGST +EachMerchantFee
                  end if
                  if s_ex("StampDuty")="True" Then
                             MerchantFeeStamp=MerchantFeeStamp +EachMerchantFee
                  end if


                  SQL="INSERT INTO WebPaymentExtraFees(ExtraValue,ReservationNo,ExtraFeesID,Fees,Days)"
                  SQL=SQL&"VALUES ("&EachMerchantFee&","&Clng(Session("RCM273_OnlineBookingNo"))&","&s_ex("ID")&","&s_ex("Fees")&",'1')"
                  webConn.Execute(SQL)

            s_ex.MoveNext
         j=j+1
         Loop
         s_ex.close
         SET s_ex=nothing
         if StampDutyRate>0 then  '--if StampDutyRate=0 then no Stampduty applied
                   StampDuty=StampDutyRate*(RentalCost + Session("RCM273_KmsCost") + RelocationFeeStampDuty + AfterHoursFeeStampDuty + ExtraStamp+ MandatoryextraStamp + StampInsurance + MerchantFeeStamp)
                   StampDuty=Round(StampDuty,2)
                   Session("RCM273_StampDuty")=StampDuty
         end if

                   GST=( RentalCost + Session("RCM273_KmsCost") + AfterHoursFeeGST + RelocationFeeGST + ExtraGST+ MandatoryExtraGST + GSTInsurance + MerchantFeeGST)*GSTRate
                  GST=Round(GST,2)
                  TotalCostTotalCost=TotalCost + StampDuty+ GST

 END IF

   GST=Round(GST,2)
   Session("RCM273_GST")=GST
   Session("RCM273_TotalEstimateofCharges")=TotalCostTotalCost

 end if
    
    

    '---------display all rates and fees here
         Set s_rate=webConn.Execute("SELECT WebPaymentDetail.*,Season FROM WebPaymentDetail,Season WHERE (ReservationNo)='"&Clng(Session("RCM273_OnlineBookingNo"))&"' and  SeasonID = Season.ID ")
        RentalCost=0
        costEachSeason=0
       While Not s_rate.EOF
               costEachSeason=s_rate("Days")*s_rate("Rate")
                  CarTotalCost=CartotalCost + s_rate("Days")*s_rate("Rate")
                  RentalCost=s_rate("Days")* s_rate("Rate") + RentalCost 
                  if s_rate("Season")="Default" then
                     Season="Rates"
                  else
                     Season=s_rate("Season")
                  end if
                   NoHours=s_rate("NoHours")
                   Response.Write("<div class='row'><div class='col-xs-6'><p class='text-left'><strong>"&Season&"</strong></p></div>")
              
                  if  Session("RCM273_DiscountRate")>0 then
                        if Session("RCM273_DiscountType")="p" then
                           Discount=Session("RCM273_DiscountRate")&"% discount"
                        else
                           Discount=""&Session("RCM273_CompanyCurrency")&""&(Session("RCM273_DiscountRate"))&" discount"
                        end if

                        Response.Write("<div class='col-xs-4'><p class='text-left'>"&s_rate("Days")&" Days at <span class='text-muted'><s>"&Session("RCM273_CompanyCurrency")&" "&FormatNumber(s_rate("StandardRate"),2)&"</s></span> <span class='text-danger'>"&Session("RCM273_CompanyCurrency")&" "&FormatNumber(s_rate("Rate"),2)&"</span><br /><span class='label label-danger'>"&Discount&"</span></p></div><div class='col-xs-2'><p class='text-right'>"&Session("RCM273_CompanyCurrency")&" "&FormatNumber(costEachSeason,2)&"</p></div></div>")
                  else
                      if NoHours=0 then
                      Response.Write("<div class='col-xs-4'><p class='text-left'>"&s_rate("Days")&" Days at "&Session("RCM273_CompanyCurrency")&" "&FormatNumber(s_rate("Rate"),2)&"</p></div><div class='col-xs-2'><p class='text-right'>"&Session("RCM273_CompanyCurrency")&" "&FormatNumber(costEachSeason,2)&"</p></div></div>")
                        else
                        Response.Write("<td align='right' class='text'>"&NoHours&" hours  @ "&Session("RCM273_CompanyCurrency")&""&FormatNumber(costEachSeason,2)&"</td><td class='text' align='right'>"&Session("RCM273_CompanyCurrency")&""&FormatNumber(costEachSeason,2)&"</td></tr>")

                        end if
                  end if
       s_rate.MoveNext
      Wend
      s_rate.Close
      Set l_s=nothing
      Session("RCM273_RentalCost")=CarTotalCost
      
         if  Session("RCM273_KmsDesc")<>"" then
                  Response.Write("<tr><td align='left' colspan='3' class='text'>"&Session("RCM273_KmsDesc")&" </td>")
                  if Session("RCM273_KmsCost")<>0 then
                     Response.Write("<td class='text' align='right'>"&Session("RCM273_CompanyCurrency")&"")
                     Response.Write(FormatNumber(Session("RCM273_KmsCost"),2))
                     Response.Write("</td></tr>")
                  else
                  Response.Write("<td class='text' align='right'></td></tr>")
                  end if
                Response.Write("</td></tr>")
         
         end if
           
           
           
            if Session("RCM273_AreaofUsed")<>"" then
                  Response.Write("<tr><td></td><td align='left'  class='text'>Area of Use</td><td align='right'  class='text'> "&Session("RCM273_AreaofUsed")&" </td><td></td></tr>")
         end if
          '--- extra fees
        Set s_extra=webConn.Execute("SELECT WebPaymentExtraFees.*, Name,ExtraDesc, Type,PayAgency FROM WebPaymentExtraFees, ExtraFees WHERE  (ReservationNo)='"&Clng(Session("RCM273_OnlineBookingNo"))&"' and ExtraFeesID = ExtraFees.ID ORDER BY TYPE, NAME ")
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
                            Response.Write("<tr><td></td><td class='text'  align='left'>"&s_extra("Name")&"</td><td class='text' align='right'> "&theQTY&"</td><td  align='right' class='text'>"&Session("RCM273_CompanyCurrency")&""&FormatNumber(EachExtraFees,2)&"</td></tr>")
                  elseif s_extra("Type")="Daily" then
                        extraType =s_extra("Type")
                        if s_extra("SetMaxPrice")  =True then
                           extraType="Fixed"
                        end if
                        Response.Write("<div class='row'><div class='col-xs-6'><p class='text-left'><strong>"&s_extra("Name")&"</strong></p></div><div class='col-xs-4'><p class='text-left'>"&extraType&" at "&Session("RCM273_CompanyCurrency")&" "&FormatNumber(s_extra("Fees"),2)&" "&theQTY&"</p></div><div class='col-xs-2'><p class='text-right'>"&Session("RCM273_CompanyCurrency")&""&FormatNumber(EachExtraFees,2)&"</p></div></div>")
                   elseif s_extra("Type")="Fixed" then
                        Response.Write("<div class='row'><div class='col-xs-6'><p class='text-left'><strong>"&s_extra("Name")&"</strong></p></div><div class='col-xs-4'><p class='text-left'>"&s_extra("Type")&" at "&Session("RCM273_CompanyCurrency")&" "&FormatNumber(s_extra("Fees"),2)&" "&theQTY&"</p></div><div class='col-xs-2'><p class='text-right'>"&Session("RCM273_CompanyCurrency")&""&FormatNumber(EachExtraFees,2)&"</p></div></div>")
                   end if
                     
                  TotalExtraFees=TotalExtraFees+EachExtraFees
                     if s_extra("ExtraDesc")<>"" then
                       'Response.Write("<div class='row'><div class='col-xs-12'><p class='text-left'>"&s_extra("ExtraDesc")&"</p></div></div>")
                    end if
                
        s_extra.MoveNext
        Loop
         s_extra.close
         set s_extra=nothing

       if  Session("RCM273_StampDuty") >0 then
                Response.Write("<tr><td></td><td class='text'  colspan='2'  align='left'>"&Session("RCM273_TaxName2")&"</td>")
                Response.Write("<td  align='right' class='text'>"&Session("RCM273_CompanyCurrency")&""&FormatNumber(Session("RCM273_StampDuty"),2)&"</td></tr>")
        end if

         IF Session("RCM273_GSTInclusive") = "Yes" then
           Response.Write("<hr />")
            Response.Write("<div class='row'><div class='col-xs-6'><p class='lead text-left'><strong>Total estimate of charges:</strong></p></div><div class='col-xs-6'><p class='text-right'><span class='lead'><strong ><mark>"&Session("RCM273_CompanyCurrency")&" "&FormatNumber(Session("RCM273_TotalEstimateofCharges"),2)&"</mark></strong></span>")
            Response.Write("<br /><span class='text-muted'><em><small>includes "&Session("RCM273_TaxName1")&" "&Session("RCM273_CompanyCurrency")&" "&FormatNumber(Session("RCM273_GST"),2)&"</small></em></span></p></div></div></div>")
         ELSE
            if Session("RCM273_GST")>0 then
                   Response.Write("<tr><td></td><td class='text'  align='left'>"&Session("RCM273_TaxName1")&"</td><td class='text' colspan='2' align='right'>  "&Session("RCM273_CompanyCurrency")&""&FormatNumber(Session("RCM273_GST"),2)&"</td></tr>")
           end if
             Response.Write("<tr><td colspan='4'  bgcolor='"&Session("RCM273_CompanyColour")&"' style='height: 1px'></td></tr>")
            Response.Write("<tr><td></td><td align='left' class='text' colspan='2'><b>Total Estimate of Charges</b></td><td class='text' align='right'><b>"&Session("RCM273_Currency")&" "&Session("RCM273_CompanyCurrency")&""&FormatNumber(Session("RCM273_TotalEstimateofCharges"),2)&"</b></td></tr>")

      END IF

        Response.Write("</div>")
        'Response.Write("</td></tr>")


      'if   Session("RCM273_CustomerID")=0 then
         'ExistingCustInfoForm
     ' end if
         '--below code added in step3
      'Session("RCM273_CustomerID")=0
      'Session("RCM273_BookingBufferNo")=0

       'Response.Write("<div class='row'>")
      'Response.Write("<td colspan='4'>")
      'Response.Write("<table width='100%' border='0' cellpadding='0' cellspacing='0'>")
   'Response.Write("<tr>")
   'Response.Write("<td>")
  if  Session("RCM273_bookingType")="Booking" then
         PersonalInfoForm
  else
          PersonalInfoFormQuote
  end if

%>
   </td></tr></table>

<%
      webConn.CLOSE
      SET webConn=nothing
%>
      </td></tr></table>
      </td></tr></table>

<!-- END RCM HTML CODE-->

<!-- #include file="include_footer.asp" -->

</body>
</html>


