var nodemailer = require("nodemailer");
var smtpTransport = nodemailer.createTransport("SMTP",{
  service: "Gmail",  // sets automatically host, port and connection security settings
  auth: {
    user: "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX",
    pass: "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
  }
});

var resuLink = 'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX';

if(typeof require !== 'undefined') {
	XLSX = require('xlsx');
}
var workbook = XLSX.readFile('MailList.xlsx');

/* Get worksheet */
var first_sheet_name = workbook.SheetNames[0];
var worksheet = workbook.Sheets[first_sheet_name];

var column_company = 'A';
var column_name = 'B';
var column_email = 'C';
var column_applied = 'D';
var column_emailed = 'E';

function email(name, email, message) {
	smtpTransport.sendMail({  //email options
		from: "XXXXXXXXXX <XXXXXXXXXXXX>", // sender address.  Must be the same as authenticated user if using Gmail.
		// to: ""+desired_cell_name.v+"<"+desired_cell_email.v+">", // receiver
		to: ""+name+"<"+email+">", // receiver
		subject: "XXXXXXXXXXXX", // subject
		text: message,

		}, function(error, response){  //callback
		  if(error){
		    console.log(error);
		  }
		  else{
		    console.log("Message sent: " + response.message);
		  }
	});
}

for(var i = 2;i>0;i++){
	row = i.toString();
	address_of_cell_company = column_company + row;
	desired_cell_company = worksheet[address_of_cell_company];
	if(desired_cell_company == undefined) {
		break;
	}
	address_of_cell_email = column_email + row;
	desired_cell_email = worksheet[address_of_cell_email];

	address_of_cell_name = column_name + row;
	desired_cell_name = worksheet[address_of_cell_name];

	address_of_cell_applied = column_applied + row;
	desired_cell_applied = worksheet[address_of_cell_applied];

	address_of_cell_emailed = column_emailed+row;
	desired_cell_emailed = worksheet[address_of_cell_emailed];


	// console.log(desired_cell_company.v, desired_cell_name.v,desired_cell_email.v, desired_cell_applied.v, desired_cell_emailed.v);

	//Haven't emailed yet
	if(desired_cell_emailed.v.toUpperCase() == 'NO' ) {
		var first_name = desired_cell_name.v.split(' ',1)[0];
		var company_name = desired_cell_company.v;

		var message = '';

		//Already applied and now emailing recruiters
		if(desired_cell_applied.v.toUpperCase() == 'YES') {
			message = 'Already Applied';
			email(desired_cell_name.v,desired_cell_email.v,message);
		}

		//Haven't applied to company yet 
		else{
			message = 'Didnt apply';
			email(desired_cell_name.v,desired_cell_email.v,message);
		}
	}
}

