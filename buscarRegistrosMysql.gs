/*
.DESCRIPTION      Buscar Registros de um banco de dados mysql e popular uma planilha do google
.NOTES  
  Version:        1.0
  Author:         Erich Oliveira https://www.linkedin.com/in/oliveiraerich/
  Modified:       19/02/2021
*/


//altere com os valores da sua conexao
var address = 'dominio.com:3306';
var user = 'SeuUsuarioBanco';
var userPwd = 'SuaSenha';
var db = 'datastudio';
var dbUrl = 'jdbc:mysql://' + address + '/' + db;


function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    {name: 'Atualizar Registros', functionName: 'buscarRegistros'}
  ];
  spreadsheet.addMenu('Menu Empresa', menuItems);
}

function getFirstEmptyRowByColumnArray(spreadSheet, column) {
  var column = spreadSheet.getRange(column + ":" + column);
  var values = column.getValues(); // get all data in one call
  var ct = 0;
  while ( values[ct] && values[ct][0] != "" ) {
    ct++;
  }
  return (ct+1);
}


function buscarRegistros() {
  
 
  var start = new Date(); // Retornar a data
  var conn = Jdbc.getConnection(dbUrl, user, userPwd);
  
  var dbMetaData = conn.getMetaData();
  var stmt = conn.createStatement();
  var results = stmt.executeQuery('SELECT CAMPO1, CAMPO2 FROM TABELA');

  //alterar o nome da tabela de acordo com sua estrutura de banco de dados
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // Retorna a aba ativa
  var doc = SpreadsheetApp.openById("1tut4ThDdyyZvi-6fKgqv0isHbmQeQKGGpdSVO5r2XdM").getSheetByName('DATABASE'); 
  var cell = doc.getRange('a1');
  var row = 0;

  //Contagem de nomes de colunas da tabela mysql.
  var getCount = results.getMetaData().getColumnCount(); 
  
  
  for (var i = 0; i < getCount; i++){  
    // O nome da coluna da tabela mysql será buscado e adicionado na planilha.
     cell.offset(row, i).setValue(results.getMetaData().getColumnName(i+1)); 
  }  
  
  var row = 1; 
  while (results.next()) {
    for (var col = 0; col < results.getMetaData().getColumnCount(); col++) { 
    // O nome da coluna da tabela mysql será buscado e adicionado na planilha.
      cell.offset(row, col).setValue(results.getString(col + 1)); 

    }
    row++;
  }
  
  results.close();
  stmt.close();
  conn.close();
  var end = new Date(); // Finalizar o script
  Logger.log('Tempo final do script: ' + (end.getTime() - start.getTime()));

}