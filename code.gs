function myFunction() {
  var rowCount = ROW_COUNT;
  var colCount = COLUMNS_COUNT;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var currSheet = ss.getSheets()[0];
  var x = currSheet.getRange(rowCount+1, colCount+1).getValue();
  if(x === "") {
    x = "0,1,1";
  }
  var lastRow = parseInt(x.split(',')[0]);
  var lastCol = parseInt(x.split(',')[1]);
  var prev = parseInt(x.split(',')[2]);
  var isNext = lastRow == rowCount;
  var currRow = isNext ? 1 : lastRow + 1;
  var currCol =  isNext ? lastCol + 1 : lastCol;
  prev = isNext ? 1 : prev;

  var newSheet = isNext ? ss.insertSheet(ss.getNumSheets()) : ss.getSheets()[ss.getNumSheets()-1];

  var url = currSheet.getRange(currRow, currCol).getValue();

  var result = isNext ? [['VideoID', 'Name','Comment','Time','Likes','Reply Count']] : [];
  
  if (url === "") {
    return;
  }

  var vidID = url.split('?v=')[1].split('&')[0];
  var nextPageToken = undefined;

  Logger.log([currRow, currCol, prev, vidID]);

  while (true) {
    var data = YouTube.CommentThreads.list('snippet', {videoId: vidID, maxResults: 100, pageToken: nextPageToken});
    nextPageToken = data.nextPageToken
    data.items.forEach(function(item) {
      result.push([vidID,
                  item.snippet.topLevelComment.snippet.authorDisplayName,
                  item.snippet.topLevelComment.snippet.textDisplay,
                  item.snippet.topLevelComment.snippet.publishedAt,
                  item.snippet.topLevelComment.snippet.likeCount,
                  item.snippet.totalReplyCount
                  ]);
    });
    if (nextPageToken === null || nextPageToken === "" || typeof nextPageToken === "undefined") {
      break;
    }

    Logger.log(result.length);

  }
  newSheet.getRange(prev, 1, result.length, 6).setValues(result);
  prev += result.length;

  currSheet.getRange(rowCount+1, colCount+1).setValue(currRow.toString() + ',' + currCol.toString() + ',' + prev.toString());
}
