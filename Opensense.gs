function get_URL(){ 
	return opensense.url.replace("[sense_key]",opensense.sense_key); 
}

function SendFeed(feed_id, value, ONSUCCESS_FN, ONERROR_FN){
  var request = new Request();
  
  writelog("Sending Feed : " + feed_id);
  
  if (feed_id == 0) {
    if (ONERROR_FN) ONERROR_FN(null,null,"FeedID = 0");
    return;
  }
  
  var postData = {
    "feed_id": feed_id,
    "value":value
  }; 
  
  var QUERY = {
    url: opensense.geturl(),
    'Content-type' : 'application/json',
    body: JSON.stringify(postData)
  };	
  
  request.post(QUERY, function ( err, res, body){    
    if (err || (res.getResponseCode() != 200 && res.getResponseCode() != 302)) {
      ONERROR_FN(err, res, body);      	 
    }
    else {
      ONSUCCESS_FN(err, res, body);
    }
  });		
}

function SendFeeds(){
  
  writelog("Sending Feeds");
  
  if(opensense.events.OnBeforeSendFeeds) opensense.events.OnBeforeSendFeeds();
  
  var feedbyidx = [];
  
  var SendFeed_FN = function (CALLBACK_FN,feedid_int) {
    var feed = opensense.feeds[feedbyidx[feedid_int]];
    var feed_id = feedbyidx[feedid_int];
    var feed_value = feed.value;			
        
    writelog("feedbyidx : "+feedbyidx +"\tfeedid_int : "+feedid_int);
    writelog(feedbyidx[feedid_int]);
    /*
    writelog("Sending Feed : " + feed_id);
    console.log("opensense.feeds : " + JSON.stringify(opensense.feeds,null,10));
    console.log("feedbyidx[feedid_int] : " + JSON.stringify(feedbyidx[feedid_int],null,10));
    console.log("feed_id : " + feed_id);
    console.log("feed : " + JSON.stringify(feed,null,10));
    */
    
    if(opensense.events.OnBeforeSendFeed) opensense.events.OnBeforeSendFeed(feed);	
    
    writelog("\Sending feed --> " + (""+feed_id) + " [processing...]",false);
    
    SendFeed(
      feed_id,
      feed_value,
      /*ONSUCCESS_FN*/ function (err, res, body){
        try
        {
          feed.json = JSON.parse(body);					
          
          writelog("\Sending feed --> " + (""+feed_id) + " [OK]");
          
          if(opensense.events.OnAfterSendFeed) opensense.events.OnAfterSendFeed(feed);	
          
          if (CALLBACK_FN) CALLBACK_FN();
        }
        catch(parseerr)
        {
          writelog("\tSending feed --> " + (""+feed_id)  + " [ERROR]");
          writelog("\tParse Query Result Error (device : " +(""+feed) + ") : " + parseerr);
          writelog("\tBody : " + body);
          if(opensense.events.OnAfterSendFeed) opensense.events.OnAfterSendFeed(null);	
        }
      },
      /*ONERROR_FN*/ function(err, res, body){ 
        writelog("[SendFeed_FN ONERROR_FN] "+body);
        if(opensense.events.OnAfterSendFeed) opensense.events.OnAfterSendFeed(null);	
      }
    );
  };
  
  //var subsequenty = require('sequenty');
  var DevicesSeqFun = [];
  
  for (var feedid in opensense.feeds)
  {
    writelog("feedid:" + feedid);
    feedbyidx.push(feedid);
    DevicesSeqFun.push(SendFeed_FN);	
  }       	
  
  DevicesSeqFun.push(function(CALLBACK_FN){
    if(opensense.events.OnAfterSendFeeds) opensense.events.OnAfterSendFeeds();
  });
  
  var seq = new Sequenty();
  seq.run(DevicesSeqFun);
}

var opensense = {
  showlog: true,
  show_datetime_in_log: true,
  url: "http://api.sen.se/events/?sense_key=[sense_key]",
  sense_key: "",
  feeds: {},
  sendfeeds: SendFeeds,
  geturl: get_URL,
  sendfeed: SendFeed,
  emptyFeeds: function() {
    this.feeds = {};
  },
  events: {
    OnBeforeSendFeeds:	null,
    OnAfterSendFeeds:	null,
    
    OnBeforeSendFeed: 	null,
    OnAfterSendFeed:	null
  }
};
