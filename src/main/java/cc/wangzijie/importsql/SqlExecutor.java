package cc.wangzijie.importsql;

import cc.wangzijie.importsql.jdbc.DbConn;

import java.io.File;
import java.nio.file.Files;
import java.util.List;
import java.util.Scanner;

public class SqlExecutor {

    private static String folderPath = "D:/ImportData/CG-D031(26到67列)/1594255837475";
    private static String outputPath;

    public static void main( String[] args ) {
        System.out.println( "请输入待执行SQL所在目录的完整路径：" );
        System.out.println( "例如：\"D:/ImportData/CG-D031(26到67列)/1594255837475\"" );
        System.out.print( "SQL目录：" );
        Scanner input = new Scanner( System.in );
        folderPath = input.nextLine();
        outputPath = folderPath + "/" + System.currentTimeMillis();
        File folder = new File( folderPath );
        File[] files = folder.listFiles();
        for (File file : files) {
            if (file.isDirectory()) {
                continue;
            }
            processFile(file);
            break;
        }
    }

    public static void processFile( File file ) {
        DbConn dbConn = null;
        try {
            dbConn = new DbConn();
            String name = file.getName();
            List<String> lines = Files.readAllLines( file.toPath( ) );
            int size = lines.size();
            System.out.println( "从文件[ " + name + " ]中读入行：" + size );
            int index = 0;
            for (String line : lines) {
                index++;
                if (line.isEmpty()) {
                    System.out.println( "[ " + index + " of " + size + " ] 空行" );
                } else if (line.startsWith( "--" )) {
                    System.out.println( "[ " + index + " of " + size + " ] 注释行：" + line );
                } else {
                    if (line.endsWith( ";" )) {
                        line = line.substring( 0, line.length()-1 );
                    }
                    int affectRow = dbConn.insert( line );
                    System.out.println( "[ " + index + " of " + size + " ] SQL影响行数：" + affectRow );
                }
            }
        } catch ( Exception e ) {
            e.printStackTrace( );
        } finally {
            if (dbConn != null) {
                dbConn.close();
            }
        }
    }

}
