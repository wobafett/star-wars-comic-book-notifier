function newComicAlert() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName('Main');
  var archive = ss.getSheetByName('Archive');
  var lastRow = s.getLastRow();
  var archiveLastRow = archive.getLastRow();
  var data = s.getRange('C:H').getValues();
  var data2 = s.getDataRange().getValues();
  var archiveData = archive.getDataRange().getValues();
  var archiveArray = archive.getRange('A2:A').getValues();
  var todayDate = Utilities.formatDate(new Date(), 'PDT', 'MMMM dd, yyyy');
  var recepients = 'recipient@example.com'; // add recepients here
  var launches = [];
  var downloadLinks = '';
  
  var miUid = 0;
  var miType = 3;
  var miPublishDate = 5;
  var miSeries = 6;
  var miIssues = 7;
  var miNotes = 8;
  
  var aiPublishDate = 1;
  var aiSeries = 2;
  var aiIssues = 3;
  
  var dupe = false;
  
  // build Archive
  ////////////////
  for (var i=0; i<lastRow; i++)
  {
    if (data2[i][miType] == "Comic book") // Check at each row to see if it's a comic
    {
      for (var j=0; j<archiveLastRow; j++)
      {
        if (archiveArray[j][0] == data2[i][miUid]) // Check at every row of archive to see if the UID already exists
        {
          dupe = true;
          break;
        }
      }
      if (dupe == false) // If the UID doesn't exist, add it
      {
        archive.appendRow([data2[i][miUid],data2[i][miPublishDate],data2[i][miSeries],data2[i][miIssues],data2[i][miNotes]]);
      }
      else
      {
        dupe = false;
      }
    }
  }
  
  for (var i=0; i<archiveLastRow; i++)
  {
    if (Utilities.formatDate(new Date(archiveData[i][aiPublishDate]),'PDT','MMMM dd, yyyy') == todayDate)
    {
      launches.push(' '+archiveData[i][aiSeries]+' #'+archiveData[i][aiIssues]);
      downloadLinks = downloadLinks+'<a href="https://comicstore.marvel.com/search?search='+archiveData[i][aiSeries].toString().replace(/[ ]/g,'+').replace(/#/g,'%23')+'+%23'+archiveData[i][aiIssues]+'">'+archiveData[i][aiSeries]+' #'+archiveData[i][aiIssues]+'</a><br />'
    }
  }
  
  // sendEmail Arguments
  //////////////////////
  
  //// subject ////
  if (launches.length>1)
  {
    var subject = 'Star Wars Comic Alert:'+launches+' Launch Today!';
  }
  else
  {
    var subject = 'Star Wars Comic Alert:'+launches+' Launches Today!';
  }
  
  //// body ////
  var body = '<html><a href="https://docs.google.com/spreadsheets/d/1ZG7BLaQvM2sdF5ZE-gJzvs5-H88g7BHkFUOvKI3cFmY/edit#gid=1784356221&fvid=434512340">Check the tracker</a>'
  +'<br /><p><b>Download:</b></p>'
  +downloadLinks
  +'</html>';
  
  Logger.log(launches.length);
  if(launches.length > 0)
  {
  MailApp.sendEmail(recepients,
                    subject,
                    body,
                    {
                      name: 'Star Wars Comic Launch Day Notifier',
                      htmlBody: body
                    });
  }
}
