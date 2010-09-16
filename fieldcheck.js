//<script>

//Function for checking the value and putting it into an integer variable 
//Made by Lioren
function checkField(value, field){
  rePattern = /^(\s|\d)+$/;
  reDigit = /^\d$/;
  a = value.length > 0 ? true : false;
  if (a)
    a = rePattern.test(value);
  if (a) {
    temp = '';
    for (i=0; i<value.length; i++){
      temp += reDigit.test(value.substring(i,i+1)) ? value.substring(i,i+1) : '';
    }
  field.value = temp;
  }
  else {
    field.value = 0;
  };
}

//lightup functions
//made by ~Q

function LightUp(targ_obj){
	if(targ_obj.value == '0') {targ_obj.value = ''}
	targ_obj.style.background = '#000099';
}

function LightOut(targ_obj){
	if(targ_obj.value == '') {targ_obj.value = '0'}
	targ_obj.style.background = '#000033';
}

function ButtonLight(targ_obj){
	targ_obj.style.background = '#ee5900';
}

function ButtonDark(targ_obj){
	targ_obj.style.background = '#660000';
}

//Statistics Function
//made by Duke Brak

function statbar1(line1, line2){
	document.forms[0].elements[0].value=line1; 
	document.forms[0].elements[1].value=line2;
}
function statbarland(line1){
	document.forms[0].elements[0].value=line1; 
}


//Invade Functions
//made by Duke Brak





