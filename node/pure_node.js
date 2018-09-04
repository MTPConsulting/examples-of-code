const mysql      = require('mysql');
const pool = mysql.createPool({
  host     : 'localhost',
  user     : 'root',
  password : '',
  database : 'price-list_development'
});

pool.getConnection(function(err, connection) {
  // Use the connection
  connection.query('SELECT * FROM categories', function (error, results, fields) {

    // Handle error after the release.
    if (error) throw error;

    const data = [];
    results.map((item) => {
        data.push(item.name);    
    });

    // Use the connection
      connection.query('SELECT * FROM groups', function (error, results2, fields) {
        // And done with the connection.
        connection.release();

        results2.map((item2) => {
            data.push(item2.name);    
        });

        console.log(data);
    });
  });
});
