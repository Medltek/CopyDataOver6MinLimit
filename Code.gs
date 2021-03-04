// Author Matyas Dana
// No rights reserved
// use at own will

var sss = SpreadsheetApp.openById('inputYourSourceSpreadSheetID'); // sss = source spreadsheet
var ss = sss.getSheetByName('inputYourSourceSheetName'); // ss = source sheet
var ts = sss.getSheetByName('inputYourTargetSheetName'); // ts = target sheet

// function for coloring rows with specific(N) data found on column 15 gray, other rows are colored white
function ColorNs(){
  var arrayOfNOs = [];
  arrayOfNOs=ts.getRange(4,15, ss.getLastRow(), 1).getValues();
 
  for(var row=4;row<=ss.getLastRow();row++){
    if(arrayOfNOs[row-4] == "N" || arrayOfNOs[row-4] == "N(stock)" || arrayOfNOs[row-4] == "STOCK(N)" || arrayOfNOs[row-4] == "STOCK (N)" || arrayOfNOs[row-4] == "N (stock)" || arrayOfNOs[row-4] == "N(CS)" || arrayOfNOs[row-4] == "N (CS)"){
      ts.getRange(row,1,1,ts.getLastColumn()).setBackgroundRGB(128,128,128);
      Logger.log("gray " + row + " value: " + arrayOfNOs[row-4]);
    }
    else{
      ts.getRange(row,1,1,ts.getLastColumn()).setBackgroundRGB(255,255,255);
      Logger.log("else " + row + " value: " + arrayOfNOs[row-4]);
      }
    
  }
}

// simple sort fnc
function MySort(){
var range = ts.getRange(4,1,ts.getLastRow(), ts.getLastColumn());
range.sort(1);

}

// Main fuction to call
// sets the start time of processing data
function CopyAndPaste(){
  var startTime = (new Date()).getTime();

  getUnique(startTime);
  
}


// funtion which looks into userCache

// if there is null value under key "completed" start cycle from the row 4(1st row with data) which copies all rows with unique data in column 10(FabricageOrderNumbers) 
// until the execution time reaches 280 seconds at which point progress is saved to userCache and "completed" gets value "no"

// if there is value "no" under key "completed" it retrieves data from userCache(progress) and continues the main cycle
// if it overcomes the time limit again(280s) progress is saved to userCache again
// if it finishes all data in time, clears the userCache
function getUnique(startTime){
    var cache = CacheService.getUserCache();
    var cached = cache.get("completed");
    
  var col=10;// FO column
  var array = [];
  array = ss.getRange(4, col, ss.getLastRow(), 1).getValues();

    if(cached == null){
        Logger.log("new cycle, bcs completed was null");

        var uniqueArray = [];
        var y=4;
        var rowIterator = 0;
        
        // Loop through array values
        for(var i=array.length-1;i>=0;i--){
          //Logger.log("Current remaining iterations i = " + i);
          var currTime = (new Date()).getTime();
          Logger.log("Current execution time = " + (currTime - startTime));
          rowIterator=4;
          if(currTime - startTime >= 280000){
            cache.put("y", y, 3600);
            cache.put("completed", "no", 3600);
            cache.put("i", i, 3600);
            cache.put("uniqueArray", uniqueArray, 3600);
            Logger.log("Execution Time exceeded, Caches stored: " + " y " + y + "completed: no " + " i: " + i + " Array: " + uniqueArray)
            break;
          } 

          if(ss.getRange(rowIterator+i, 2).getValue() == ""){
            rowIterator=rowIterator+i;
            //Logger.log("space on line: " + rowIterator); 
            continue;
          }
          else if(array[i]=="-" || array[i]=="FO" || array[i]==" " || array[i]=="" || array[i]=="STOCK"){
            rowIterator=rowIterator+i;
            /*Logger.log("wrong format FO: "+ array[i] + " into line: " + (y)
             + "from line: " + rowIterator); 
            Logger.log("1. column value " + CellVal(rowIterator, y));*/
            CopyDataToNewFile(rowIterator,2, y);
            y=y+1;
            continue;
          } 
          else if(!uniqueArray.includes(parseInt(array[i]))){
            uniqueArray.push(parseInt(array[i]));
            /*Logger.log("UniqueArray value " + uniqueArray[uniqueArray.length-1] + " TypeOf "+ typeof(uniqueArray[uniqueArray.length-1]));
            Logger.log("Array value " + array[i] + " TypeOf "+ typeof(array[i]));
            rowIterator=rowIterator+i;
            Logger.log("Unique FO: "+ array[i] + " into line: " + (y) + "from line: " + rowIterator); 
            Logger.log("1. column value " + CellVal(rowIterator, y));*/
            CopyDataToNewFile(rowIterator,2, y);
            y=y+1;
          }
          else if(uniqueArray.includes(parseInt(array[i]))){
                  rowIterator=rowIterator+i;
                  //Logger.log("Duplicite FO on row:" + rowIterator);
              }
          else{
                rowIterator=rowIterator+i;
                //Logger.log("Else on cycle on row: " + rowIterator);
              }
          if(i==0){
            ts.getRange(y,1, ts.getLastRow(), ts.getLastColumn()).clear();
          }
        }
      return uniqueArray;
    }

    else if(cached == "no" && cache.get("uniqueArray") != null && cache.get("y") != null && cache.get("i") != null){
      var uniqueArray = [];
      var y = parseInt(cache.get("y"));
      var tempArray = [];
      tempArray = cache.get('uniqueArray').split(',');
      for(var j=0;j<tempArray.length-1;j++){
        uniqueArray[j] = parseInt(tempArray[j]);
        /*Logger.log("UniqueArray value " + uniqueArray[j] + " TypeOf "+ typeof(uniqueArray[j]));
        Logger.log("Array value " + tempArray[j] + " TypeOf "+ typeof(tempArray[j]));*/
        
      }
      
      /*Logger.log("Execution Time exceeded, Caches restored: " + " y " + y + "completed: no " + " i: " + cache.get("i") + " Unique Array filled: " + uniqueArray);*/

      for(var i=parseInt(cache.get("i"));i>=0;i--){

          var currTime = (new Date()).getTime();

          if(currTime - startTime >= 280000){
            cache.put("y", y, 3600);
            cache.put("completed", "no", 3600);
            cache.put("i", i, 3600);
            cache.put("uniqueArray", uniqueArray, 3600);
            //Logger.log("Execution Time exceeded, Caches stored: " + " y " + y + "completed: no " + " i: " + i + " Array: " + uniqueArray)
            break;
          }
          
          rowIterator=4;

          if(ss.getRange(rowIterator+i, 2).getValue() == ""){
            /*Logger.log("space on  "+ array[i] + " into line: " + (y)
             + "from line: " + rowIterator); */
            continue;
          }
          else if(array[i]=="-" || array[i]=="FO" || array[i]==" " || array[i]=="" || array[i]=="STOCK"){
            rowIterator=rowIterator+i;
            /*Logger.log("wrong format FO: "+ array[i] + " into line: " + (y)
             + "from line: " + rowIterator); 
            Logger.log("1. column value " + CellVal(rowIterator, y));*/
            CopyDataToNewFile(rowIterator,2, y);
            y=y+1;
            continue;
          } 
          else if(!uniqueArray.includes(parseInt(array[i]))){
            uniqueArray.push(parseInt(array[i]));
            /*Logger.log("UniqueArray value " + uniqueArray[uniqueArray.length-1] + " TypeOf "+ typeof(uniqueArray[uniqueArray.length-1]));
            Logger.log("Array value " + array[i] + " TypeOf "+ typeof(array[i]));
            rowIterator=rowIterator+i;
            Logger.log("Unique FO: "+ array[i] + " into line: " + (y)
             + "from line: " + rowIterator); 
            Logger.log("1. column value " + CellVal(rowIterator, y));*/
            CopyDataToNewFile(rowIterator,2, y);
            y=y+1;
          }
          else if(uniqueArray.includes(parseInt(array[i]))){
                  rowIterator=rowIterator+i;
                  //Logger.log("Duplicite FO on row:" + rowIterator);
              }
          else{
                rowIterator=rowIterator+i;
                //Logger.log("Else on cycle on row: " + rowIterator);
              }
          if(i==0){
            ts.getRange(y,1, ts.getLastRow(), ts.getLastColumn()).clear();
          }
        }
    }
    else{
      //Logger.log("Caches went very wrong");
      cache.removeAll(['y', 'i', 'completed', 'uniqueArray']);
    }
  
  //Logger.log("Full Unique Array: " + uniqueArray);
  cache.removeAll(['y', 'i', 'completed', 'uniqueArray']);
}

// function copying the desired range from source sheet to target sheet
function CopyDataToNewFile(row, col, y) {
  
  var range = ss.getRange(row,2, 1, 27); //assign the range you want to copy
  var copy = range.getValues();

  PutText(row, y);
  ts.getRange(y,2, 1, 27).setValues(copy) //new range you want to paste a value in
  
}

// function copying data of a merged cell(CellVall fnc) to normal cell
function PutText(row, y) { // pasting data from a pasted cell using CellVal fnc
  ts.getRange(y,1, 1, 1).setValue(CellVal(row,y));
} 

function CellVal(row,y) { // taking data from a merged cell
  var cell = ss.getRange(row, 1);
  return (cell.isPartOfMerge() ? cell.getMergedRanges()[0].getCell(1, 1) : cell).getValue();
}











