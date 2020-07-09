package cc.wangzijie.importsql.jdbc;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.sql.Statement;

public class DbConn {

    private static final String URL = "jdbc:oracle:thin:@10.73.59.104:1521:ORCL_QMS";
    private static final String USERNAME = "DO_QMS";
    private static final String PASSWORD = "mwdoqms@18666@2019";
    private static final String DRIVER = "oracle.jdbc.driver.OracleDriver";

    private Connection conn = null;
    private Statement statement = null;

    public DbConn( ) {
        try {
            Class.forName( DRIVER );
            conn = DriverManager.getConnection( URL, USERNAME, PASSWORD );
            statement = conn.createStatement( );
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public int insert(String sql) {
        try {
            return statement.executeUpdate( sql );
        } catch ( Exception e ) {
            e.printStackTrace();
        }
        return -1;
    }

    public void close() {
        if(conn != null) {
            try {
                conn.close();
            } catch (SQLException e) {
                e.printStackTrace();
            }
        }
    }
}