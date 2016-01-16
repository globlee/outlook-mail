
var request = require("request")
module.exports=function(mysubject, message, sendto,  token, callback){

	var options = {
		url:"https://outlook.office.com/api/v2.0/me/sendmail",
		json:{
			Message: {
			    Subject: mysubject,
			    Body: {
			      ContentType: "Text",
			      Content: message
			    },
			    ToRecipients: [
			    ]
			  },
			  "SaveToSentItems": "true"
		},
		headers:{
			Authorization:"bearer "+token,
			Accept: 'text/*, application/xml, application/json; odata.metadata=none',
		}
	};
	options.headers['Content-type'] = 'application/json; odata.metadata=none'
	for (var i=0; i< sendto.length; i++){
		options.json.Message.ToRecipients.push({  EmailAddress: { Address:sendto[i] }   })
	}
	console.log(options, options.json.Message.ToRecipients)
	request.post(options, function(err, res, body){
		//console.log(res.statusCode, body, err)
		if (err){
			return callback(err)
		} else {
			return callback(null)
		}
	});

};
