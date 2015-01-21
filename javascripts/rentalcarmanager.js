function openTarget (form, features, windowName) {
  if (!windowName)
    windowName = 'formTarget' + (new Date().getTime());
  form.target = windowName;
  open ('', windowName, features);
}
function Validate()
{
 if (document.theForm.driverage.value == "0")
  {      alert("Please select Age of Youngest Driver.");
      document.theForm.driverage.focus();
      return (false);
  }

 // if (document.theForm.firstname.value == "")
  {  //    alert("First Name required.");
     // document.theForm.firstname.focus();
     // return (false);
  }
 // if (document.theForm.lastname.value == "")
  {  //    alert("Last Name required.");
     // document.theForm.lastname.focus();
     // return (false);
  }
 // if (document.theForm.Phone.value == "")
  {   //   alert("Phone No required.");
     // document.theForm.Phone.focus();
     // return (false);
  }
 // if (document.theForm.CustomerEmail.value == "")
  {  //    alert("Email Address required.");
     // document.theForm.CustomerEmail.focus();
    //  return (false);
  }
  // var emailRegEx = /^[a-zA-Z0-9._-]*\@[a-zA-Z0-9._-]*$/;
  //    if(!emailRegEx.test(theForm.CustomerEmail.value))
      {
    //     alert("Invalid Email address");
    //     return false;
      }

//   if (document.theForm.CountryID.value == "0")
  { //     alert("Please select your Country of Residence.");
    //  document.theForm.CountryID.focus();
    //  return (false);
  }

window.open('', 'success', 'width=828,height=950,status=yes,resizable=yes,scrollbars=yes');
return (true);


}

function checkNumeric(objName)
{        var numberfield = objName;
   if (chkNumeric(objName) == false)
   {  numberfield.select();
      numberfield.focus();
      return false;
   }
   else
   {        return true;
   }
}
function chkNumeric(objName)
{   // only allow 0-9 be entered, plus any values passed
var checkOK = "0123456789";
var checkStr = objName;
var allValid = true;
var decPoints = 0;
var allNum = "";

   if  ((checkStr.value.length) != 0)
   {        for (i = 0;  i < checkStr.value.length;  i++)
            {        ch = checkStr.value.charAt(i);
               for (j = 0;  j < checkOK.length;  j++)
         if (ch == checkOK.charAt(j))
         break;
         if (j == checkOK.length)
         {     allValid = false;
            break;
               }
               if (ch != ",")
                  allNum += ch;
      }
            if (!allValid)
      {  alert("Please enter Numeric value");
            return (false);
      }
   }

}

jQuery(function () {
      jQuery('#txtStartDate').datepicker(
       {

         showOn: "button",
         buttonImage: "images/dynCalendar.png",
         buttonImageOnly: true,
         minDate: 0,
         maxDate: '+6M +10D',
         showAnim: 'fadeIn',
         numberOfMonths: 1,
         showButtonPanel: true,
         //altField: "#PickupDay",
         //altFormat: "dd",
         onClose: function (dateText, picker) {
           //alert(dateText);
           //alert("PickupDay = "+ dateText.split(/\//)[1]);
           //alert(getDayName(dateText, "/"));

           var dDefaultPickupDate = '12/06/2014';
           //alert('dateText = ' + dateText);
           //alert('dDefaultPickupDate = ' + dDefaultPickupDate);
           //alert(document.getElementById("txtEndDate").value);
           //PickupYearPickupMonthPickupDay
           document.forms['theform'].PickupDay.value = dateText.split(/\//)[1];
           document.forms['theform'].PickupMonth.value = dateText.split(/\//)[0];
           document.forms['theform'].PickupYear.value = dateText.split(/\//)[2];

           //As requested need to set the DropOff date as PickupDate
           if (dateText != dDefaultPickupDate) {
             dDefaultPickupDate = dateText;
             document.forms['theform'].DropoffDay.value = dateText.split(/\//)[1];
             document.forms['theform'].DropoffMonth.value = dateText.split(/\//)[0];
             document.forms['theform'].DropoffYear.value = dateText.split(/\//)[2];
             document.getElementById("txtEndDate").value = dDefaultPickupDate;
             //txtEndDate
           }
         }
       }); //txtStartDate

    jQuery('#txtEndDate').datepicker(
       {
         showOn: "button",
         buttonImage: "images/dynCalendar.png",
         buttonImageOnly: true,
         minDate: 0,
         maxDate: '+6M +10D',
         showAnim: 'fadeIn',
         numberOfMonths: 1,
         showButtonPanel: true,
         //altField: "#PickupDay",
         //altFormat: "dd",
         onClose: function (dateText, picker) {
           //alert(dateText);
           //alert(dateText.split(/\//)[0]);
           document.forms['theform'].DropoffDay.value = dateText.split(/\//)[1];
           document.forms['theform'].DropoffMonth.value = dateText.split(/\//)[0];
           document.forms['theform'].DropoffYear.value = dateText.split(/\//)[2];
         }
       }); //txtEndDate

  });   //jQuery function
  
  
  $(function () {
  $('[data-toggle="tooltip"]').tooltip()
})

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
