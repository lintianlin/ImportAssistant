package com.sinfeeloo.importassistant;

import java.sql.*;
import java.util.*;

public class MainActivity {

    /**
     * 表名（需要手动修改）
     */
    private final String tableName = "t_recharge_type_setting";
    /**
     * 生成的文件名（需要手动修改）
     */
    private final String fileName = "RechargeTypeSetting";

    /**
     * 包名（需要手动修改）
     */
    private final String packageName = "com.intelliquor.cloud.shop.system";
    /**
     * 表注解（需要手动修改）
     */
    private final String tableAnnotation = "充值类型设置";
    /**
     * 数据库连接信息
     */
    private final String URL = "jdbc:mysql://39.106.0.169:3306/cloud-shop?serverTimezone=UTC&useSSL=true&useUnicode=true&characterEncoding=utf-8";
    private final String USER = "cloud_shop";
    private final String PASSWORD = "cloud_shop";
    private final String DRIVER = "com.mysql.cj.jdbc.Driver";

    /**
     * 生成java文件路径,手动改成自己的地址（需要手动修改）
     */
    private final String diskPath = "/Users/lintianlin/IdeaProjects/cloud-shop/shop-system/src/main/java/com/intelliquor/cloud/shop/system";
    /**
     * 生成mapper.xml文件路径,手动改成自己的地址（需要手动修改）
     */
    private final String mapperDiskPath = "/Users/lintianlin/IdeaProjects/cloud-shop/shop-system/src/main/resources";

    /**
     * 生成文件前缀名
     */
    private final String changeTableName = replaceUnderLineAndUpperCase(fileName, true);


    /**
     * 主方法
     *
     * @param args
     * @throws Exception
     */
    public static void main(String[] args) throws Exception {
        MainActivity codeGenerateUtils = new MainActivity();
        codeGenerateUtils.generate();
    }


    /**
     * 获取数据库连接
     *
     * @return
     * @throws Exception
     */
    public Connection getConnection() throws Exception {
        Class.forName(DRIVER);
        Connection connection = DriverManager.getConnection(URL, USER, PASSWORD);
        return connection;
    }


    /**
     * 生成代码
     *
     * @throws Exception
     */
    public void generate() throws Exception {
        try {
            Map<String, Object> dataMap = new HashMap<>();
            //获取数据库表列信息
            getTableColumnsForMySQL(dataMap);
            //生成Mapper文件

        } catch (Exception e) {
            throw new RuntimeException(e);
        } finally {

        }
    }


    /**
     * MySQL专用
     *
     * @param dataMap
     */
    public void getTableColumnsForMySQL(Map<String, Object> dataMap) {
        try {
            Connection conn = getConnection();
            String sql = "select * from " + tableName;
            PreparedStatement stmt = conn.prepareStatement(sql);
            ResultSet rs = stmt.executeQuery(sql);
            ResultSetMetaData data = rs.getMetaData();
            // 获得所有列的数目及实际列数
            int columnCount = data.getColumnCount();
            String[] columnNames = new String[columnCount];
            String[] propertyNames = new String[columnCount];
            String[] columnTypeNames = new String[columnCount];
            for (int i = 1; i <= columnCount; i++) {
                // 获得指定列的列名
                String columnName = data.getColumnName(i);
                columnNames[i - 1] = columnName;
                propertyNames[i - 1] = replaceUnderLineAndUpperCase(columnName, false);
                columnTypeNames[i - 1] = data.getColumnTypeName(i);
                // 获得指定列的列值
                int columnType = data.getColumnType(i);
                // 获得指定列的数据类型名
                String columnTypeName = data.getColumnTypeName(i);
                // 所在的Catalog名字
                String catalogName = data.getCatalogName(i);
                // 对应数据类型的类
                String columnClassName = data.getColumnClassName(i);
                // 在数据库中类型的最大字符个数
                int columnDisplaySize = data.getColumnDisplaySize(i);
                // 默认的列的标题
                String columnLabel = data.getColumnLabel(i);
                // 获得列的模式
                String schemaName = data.getSchemaName(i);
                // 某列类型的精确度(类型的长度)
                int precision = data.getPrecision(i);
                // 小数点后的位数
                int scale = data.getScale(i);
                // 获取某列对应的表名
                String tableName = data.getTableName(i);
                // 是否自动递增
                boolean isAutoInctement = data.isAutoIncrement(i);
                // 在数据库中是否为货币型
                boolean isCurrency = data.isCurrency(i);
                // 是否为空
                int isNullable = data.isNullable(i);
                // 是否为只读
                boolean isReadOnly = data.isReadOnly(i);
                // 能否出现在where中
                boolean isSearchable = data.isSearchable(i);
            }
            dataMap.put("columnNames", columnNames);
            dataMap.put("propertyNames", propertyNames);
            dataMap.put("columnTypeNames", columnTypeNames);
            dataMap.put("columnComments", getComments());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 获取数据库字段注释
     *
     * @return
     */
    public String[] getComments() throws Exception {
        Connection conn = getConnection();
        String sql = "show full columns from " + tableName;
        PreparedStatement stmt = conn.prepareStatement(sql);
        ResultSet rs = stmt.executeQuery(sql);
        ResultSetMetaData data = rs.getMetaData();
        int columnCount = data.getColumnCount();

        String[] comments = new String[100];
        int i = 0;
        while (rs.next()) {
            comments[i] = rs.getString("Comment");
            i++;
        }
        return comments;
    }
}
