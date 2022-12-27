package _JDBC;

import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.sql.Statement;

public class JDBCParent {

    public static Statement statement;
    public static Connection connection;

    @BeforeTest
    public void DBConnectionOpen()
    {
        String url = " ";  // write your URL with table at the end here
        String username = " "; // write the username here
        String password = " ";  // write the password here 

        try{
            connection = DriverManager.getConnection(url, username, password);
            statement = connection.createStatement();
        } catch (SQLException e) {
            throw new RuntimeException(e);
        }
    }

    @AfterTest
    public void DBConnectionClose(){
        try{
            connection.close();
        } catch (SQLException e) {
            throw new RuntimeException(e);
        }
    }
}
