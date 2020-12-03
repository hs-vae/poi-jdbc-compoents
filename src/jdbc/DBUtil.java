package jdbc;

import java.sql.*;

/**
 * 数据库工具类
 */
public class DBUtil {
    Connection conn=null;
    PreparedStatement pstmt=null;
    ResultSet rs=null;
    /**
     * 得到数据库连接
     */
    public Connection getConnection(){
        //通过Config获取mysql数据库配置信息
        String driver=Config.getValue("driver");
        String url=Config.getValue("url");
        String user=Config.getValue("user");
        String password=Config.getValue("password");
        try{
            //指定驱动程序
            Class.forName(driver);
            //建立数据库连接
            conn= DriverManager.getConnection(url,user,password);
            return conn;
        }catch(Exception e){
            e.printStackTrace();
            return null;
        }
    }
    /**
     * 释放资源
     */
    public void closeAll(){
        //如果rs不空，关闭rs
        if(rs!=null){
            try{
                rs.close();
            }catch (SQLException e){
                e.printStackTrace();
            }
        }
        //如果pstmt不空，关闭rs
        if(pstmt!=null){
            try{
                pstmt.close();
            }catch (SQLException e){
                e.printStackTrace();
            }
        }
        //如果conn不空，关闭rs
        if(conn!=null){
            try{
                conn.close();
            }catch (SQLException e){
                e.printStackTrace();
            }
        }
    }
    /**
     * 执行sql语句，可以进行查询
     */
    public ResultSet executeQuery(String preparedSql,Object[] param){
        //处理SQL,执行Sql
        try{
            //得到PreparedStatement对象
            pstmt=conn.prepareStatement(preparedSql);
            if(param!=null){
                for(int i=0;i<param.length;i++){
                    //为预编译sql设置参数
                    pstmt.setObject(i+1,param[i]);
                }
            }
            rs=pstmt.executeQuery();
        }catch (SQLException e){
            e.printStackTrace();
        }
        return rs;
    }
    /**
     * 执行sql语句，可以进行增、删、改的操作，不能进行查询
     */
    public int executeUpdate(String preparedSql,Object[] param){
        int num=0;
        //处理Sql，执行SQL
        try {
            //得到PreparedStatement对象
            pstmt=conn.prepareStatement(preparedSql);
            if(param!=null){
                for(int i=0;i<param.length;i++){
                    //为预编译sql设置参数
                    pstmt.setObject(i+1,param[i]);
                }
            }
            num=pstmt.executeUpdate();
        }catch (SQLException e){
            e.printStackTrace();
        }
        return num;
        }

}
