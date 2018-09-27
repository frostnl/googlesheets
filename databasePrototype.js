function TestOutputStockReport(){
  //var stock = FetchStock();
  var stock = getStockFromFile();
  var stockReport = new Database("Stock Report", "code", 1,2,1);
  stockReport.saveAll(stock);
}


function Database(sheetName, key, headerRow, dataStartRow, dataColumnStart){
  // return the spreadsheet object
  this.sheetName = sheetName;
  this.headerRow = headerRow;
  this.dataRowStart = dataStartRow;
  this.dataColumnStart = dataColumnStart;
  this.key = key;
  
  this.sheet = getSheet(this.sheetName);
  this.dataColumns = this.sheet.getLastColumn() - this.dataColumnStart +1; 
  this.dataRows = this.sheet.getLastRow() - this.dataRowStart +1

  /*
   * getDataArray
   * return data in a 2D array
   */   
  this.getDataArray = function(){
    return this.sheet.getRange(this.dataRowStart,this.dataColumnStart,this.dataRows,this.dataColumns).getValues(); 
  };
  

  /*
   * getKeyColumn
   */    
  this.getKeyColumn = function(){
    var headerArray = this.getHeaderArray();
    var index = headerArray.indexOf(this.key);
    
    return index + this.dataColumnStart -1; 
  }  
  
  
  
  /*
   * getKeysArray
   */    
  this.getKeysArray = function(){
    return this.sheet.getRange(this.dataRowStart,this.getKeyColumn(),this.dataRows,1).getValues(); 
  }

  /*
   * getHeaderArray
   * return the header fields as an array
   */
  this.getHeaderArray = function(){
    return this.sheet.getRange(this.headerRow,this.dataColumnStart,1,this.dataColumns).getValues()[0];
  }

  
  /*
   * getDataFields
   * return all fields as an object array
   */
  this.getDataFields = function(){
    var headerArray = this.getHeaderArray();
    
    var fields = [];
    for (var i=0;i<headerArray.length;i++){
      var field = [];
      field["index"] = i;
      field["label"] = headerArray[i];
      field["name"] = headerArray[i].split(' ').join('');
      fields[field["name"]] = field;
    }
    return fields;
  }
  
   
  /**
   ** fetchAll
   **/
  this.fetchAll = function() {
    var data = this.getDataArray();
    var fields = this.getDataFields();
    var results = [];
    
    for(var i=0;i<data.length;i++){
      if (data[i][0] != ""){
        var row = [];
        for (var fieldName in fields){
          var field = fields[fieldName]
          row[field.name] = data[i][field.index];
        }
        results.push(row);
      }// end if row is valid
    }
    return results;
  }
  
  
  /**
   ** fetchByKey
   **/  
  this.fetchByKey = function(key){
    var index = this.indexByKey(key);
    if (index >= 0){
      var rowData = this.getRowData(index);
      var fields = this.getFields();
      var row = [];
      for (var fieldName in fields){
        var field = fields[fieldName]
        var value = rowData[0][field.index];
        if (typeof value !== "undefined"){
          row[field.name] = value;
        } else {
          row[field.name] = "";  
        }       
      }
    } else {
      row = null;  
    }
     
    return row
  }
    
  

  
 
  
  /*
   * getRowData
   */  
  this.getRowData = function(index) {
    return this.sheet.getRange(index+this.dataRowStart,1,1,this.columns).getValues();  
  }
  
  
  
  
  
  /**
   ** saveRow
   **/  
  this.saveRow = function(obj){
    var key = obj[this.key];
    var fields = this.getFields();
    var sheet = this.getSheet;
    var valid = true;
    
    var outputRow = [];
    for (var fieldName in fields){
      if (fieldName in obj){
        outputRow[fields[fieldName].index] = obj[fieldName]
      } else {
        valid = false;
      }
    }
    
    if (valid){
      var outputArray = [];
      outputArray.push(outputRow);
      var index = this.indexByKey(key);
      if (index >= 0){
        row = index+this.dataRow;
      } else {
        row = this.rows + 1;
      }     
      this.sheet.getRange(row,1,1,outputArray[0].length).setValues(outputArray);
      
    }
  }
  
  
  /**
   ** saveAll
   **/   
  this.saveAll = function(objArray){
    var fields = this.getDataFields();
    var sheet = this.getSheet;
    var valid = true;
    
    var outputArray = [];
    for (var i=0; i<objArray.length; i++){
      var obj = objArray[i];
      var outputRow = [];
      for (var fieldName in fields){
        if (fieldName in obj){
          outputRow[fields[fieldName].index] = obj[fieldName]
        } else {
          valid = false;
        }
      }
      outputArray.push(outputRow);
    }
    
    if (valid){
      this.truncateTable();
      this.sheet.getRange(this.dataRowStart,this.dataColumnStart,outputArray.length,outputArray[0].length).setValues(outputArray);
    }   
 
  }
  
  
  /*
   * truncateTable
   */   
  this.truncateTable = function(){
    if (this.dataRows > 0 && this.dataColumns > 0){
      this.sheet.getRange(this.dataRowStart,this.dataColumnStart,this.dataRows,this.dataColumns).clearContent();
    }
  }
 
  
  /**
   ** saveAll
   **/   
  this.saveAllAssociative = function(objArray){
    var fields = this.getDataFields();
    var sheet = this.getSheet;
    var valid = true;
    
    var outputArray = [];
    for (var key in objArray){
      var obj = objArray[key];
      var outputRow = [];
      for (var fieldName in fields){
        if (fieldName in obj){
          outputRow[fields[fieldName].index] = obj[fieldName]
        } else {
          outputRow[fields[fieldName].index] = "";
          //valid = false;
        }
      }
      outputArray.push(outputRow);
    }
    
    if (valid){
      this.truncateTable();
      this.sheet.getRange(this.dataRowStart,this.dataColumnStart,outputArray.length,outputArray[0].length).setValues(outputArray);
    }   
 
  }  
  
  
  /**
   ** keyExists
   **/    
  this.keyExists = function(key) {
    var keys = this.getKeys().reduce(function(a,b){ return a.concat(b); });; 
    var index = keys.indexOf(key);
    return index >= 0;
  }
  
  
  /**
   ** indexByKey
   **/    
  this.indexByKey = function(key) {
    var keys = this.getKeys().reduce(function(a,b){ return a.concat(b); });; 
    var index = keys.indexOf(key); 
    return index;
  }

  
  
  /**
   ** newRow
   **/    
  this.newRow = function(){
    var data = [];
    var fields = this.getFields();
    for (var fieldName in fields){
      data[fieldName] = "";
    }
    return data
  }
  

  
  
  
  /**
   ** upsert
   **/
  this.upsert = function(obj){
    // test to see if it exists
    var key = obj[this.key];
    var fields = this.getFields();
    var sheet = this.getSheet;
    
    if (typeof key !== "undefined"){ 
      var data = this.fetchByKey(key);
      if (data == null){
        var data = this.newRow();
      }  
      for (var fieldName in obj){
        var field = fields[fieldName];
        if (typeof field !== "undefined"){
          data[field.name] = obj[fieldName];  
        } // end if
      } //end for fieldName in obj
      this.save(data);
    }
  }
  
  
  
  
  
   /**
   ** upsertAll
   **/ 
  this.upsertAll = function(objArray){
    var fields = this.getFields();
    var data = this.fetchAll();
    var databaseRows = [];
    
    // turn into associative array
    for (i=0; i<data.length; i++){
      row = data[i];
      key = row[this.key];
      databaseRows[key] = row;
    }
    
    // update value and add new rows
    for (i=0; i<objArray.length; i++){
      var obj = objArray[i];
      if (this.key in obj){
        var key = obj[this.key];
        if (key in databaseRows){
          var databaseRow = databaseRows[key];
        } else {
          var databaseRow = this.newRow();  
        }
          
        for (field in obj){
          if (field in databaseRow){
            databaseRow[field] = obj[field];  
          } else {
            errors.push("Error: " + field + " not found at row " + i); 
          }      
        }
        databaseRows[key] = databaseRow;
      }
    }
    
    // turn back to array.. in the right order
    for (var i=0; i<data.length; i++){
      key = data[i][this.key];
      if (key in databaseRows) {
        data[i] = databaseRows[key];
        delete databaseRows[key];
      } 
    }
    for (var row in databaseRows){
      data.push(databaseRows[row]); 
    }
    
    // save all
    this.saveAll(data);
      
  }
  
  
};
