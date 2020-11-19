package cc.wangzijie.importsql;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.PrintStream;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;

public class SqlGenerator {

    private final static int MIN_WBS = 1;
    private final static int MAX_WBS = 6;
    private final static String FOLDER_PATH = "D:\\BOM\\GT-C072(1至6列)";
    private static String outputPath;

    public static void main( String[] args ) {
//        Scanner input = new Scanner( System.in );
//        System.out.println( "请输入待执行XlSX所在目录的完整路径：" );
//        System.out.println( "例如：\"D:/ImportData/CG-D031(26到67列)\"" );
//        System.out.print( "XLSX目录：" );
//        folderPath = input.nextLine();
//        System.out.println( "请输入最大的WBS号：" );
//        System.out.println( "例如：\"67\"" );
//        System.out.print( "最大的WBS号：" );
//        MAX_WBS = input.nextInt();
        outputPath = FOLDER_PATH + "/" + System.currentTimeMillis( );
        File folder = new File( FOLDER_PATH );
        File[] files = folder.listFiles();
        for (File file : files) {
            if (file.isDirectory()) {
                continue;
            }
            processFile(file);
            // break;
        }
    }


    private static void processFile(File originFile) {
        String fileName = originFile.getName();
        List<String> wbsList = generateWbsList(fileName);
        List<Map<String, String>> rowList = parseRowList( originFile );
        List<List<String>> listOfSqlList = new LinkedList<>();
        for (String wbs : wbsList) {
            // 遍历生成的WBS号
            System.out.println( "\n\nWBS = " + wbs );
            List<String> sqlList = new LinkedList<>();
            for (int i=0; i< rowList.size(); i++) {
                Map<String, String> cellMap = rowList.get( i );
                int rowIndex = i+1;
                String colA = nullToBlank(cellMap.get("A" + rowIndex));
                String colB = nullToBlank(cellMap.get("B" + rowIndex));
                String colC = nullToBlank(cellMap.get("C" + rowIndex));
                String colD = nullToBlank(cellMap.get("D" + rowIndex));
                String colE = nullToBlank(cellMap.get("E" + rowIndex));
                String colF = nullToBlank(cellMap.get("F" + rowIndex));
                String colG = nullToBlank(cellMap.get("G" + rowIndex));
                String colH = nullToBlank(cellMap.get("H" + rowIndex));
                String colI = nullToBlank(cellMap.get("I" + rowIndex));
                String colJ = nullToBlank(cellMap.get("J" + rowIndex));
                String colK = nullToBlank(cellMap.get("K" + rowIndex));
                String colL = nullToBlank(cellMap.get("L" + rowIndex));
                String colM = nullToBlank(cellMap.get("M" + rowIndex));
                String colN = nullToBlank(cellMap.get("N" + rowIndex));
                String colO = nullToBlank(cellMap.get("O" + rowIndex));

                // 使用生成的wbs替换列B
                colB = wbs;

//                String sql =
//                        "INSERT INTO T2_BAS_WBS_BOM ("
//                        + "ID, MATNR, PSPID, MATNO, WERKS, CREATE_TIME, IDNRK, MENGE, MEINS, SORTF, DRAWE, DRAW1, DRAWM, DRAW2, CHJLX, CREATE_BY, IS_DELETED, PROJECT_CODE"
//                        + ") select do_qms.t2_bas_WBS_bom_seq.nextval,"
//                        + "'" + colA + "',"
//                        + "'" + colB + "',"
//                        + "'" + colC + "',"
//                        + "'" + colD + "',"
//                        + "'" + colE + "',"
//                        + "'" + colF + "',"
//                        + "'" + colH + "',"
//                        + "'" + colI + "',"
//                        + "'" + colJ + "',"
//                        + "'" + colK + "',"
//                        + "'" + colL + "',"
//                        + "'" + colM + "',"
//                        + "'" + colN + "',"
//                        + "'" + colO + "',"
//                        + "'system',"
//                        + " 0,"
//                        + " substr('" + colB + "',0,7)"
//                        + " from dual where ('" + colB + "','" + colC + "','" + colF + "')"
//                        + " not in (select PSPID, MATNO, IDNRK from T2_BAS_WBS_BOM );";
                String sql = "INSERT INTO DO_QMS.T2_BAS_WBS_BOM"
                             + " (\"ID\","
                             + " \"MATNR\","
                             + " \"PSPID\","
                             + " \"MATNO\","
                             + " \"WERKS\","
                             + " \"CREATE_TIME\","
                             + " \"IDNRK\","
                             + " \"MENGE\","
                             + " \"MEINS\","
                             + " \"SORTF\","
                             + " \"DRAWE\","
                             + " \"DRAW1\","
                             + " \"DRAWM\","
                             + " \"DRAW2\","
                             + " \"CHJLX\","
                             + " \"CREATE_BY\","
                             + " \"IS_DELETED\","
                             + " \"PROJECT_CODE\")"
                             + " SELECT"
                             + " DO_QMS.T2_BAS_WBS_BOM_SEQ.NEXTVAL,"
                             + " '" + colC + "',"
                             + " '" + colB + "',"
                             + " '" + colA + "',"
                             + " '" + colD + "',"
                             + " '" + colE + "',"
                             + " '" + colF + "',"
                             + " '" + colH + "',"
                             + " '" + colI + "',"
                             + " '" + colJ + "',"
                             + " '" + colK + "',"
                             + " '" + colL + "',"
                             + " '" + colM + "',"
                             + " '" + colN + "',"
                             + " '" + colO + "',"
                             + " 'system',"
                             + " '0',"
                             + " substr('" + colB + "',0,7)"
                             + " from dual where not exists ("
                             + " select 1 from DO_QMS.T2_BAS_WBS_BOM"
                             + " where \"PSPID\"='" + colB
                             + "' and \"MATNO\"='" + colA
                             + "' and \"IDNRK\"='" + colF
                             + "' and UPPER(\"SORTF\")='" + colJ
                             + "' and \"MATNR\"='" + colC
                             + "' );";
                sqlList.add( sql );
            }
            listOfSqlList.add( sqlList );
        }
        printToFile( originFile, wbsList, listOfSqlList );
    }

    private static void printToFile(File originFile, List<String> wbsList, List<List<String>> listOfSqlList) {
        for (int i = 0; i < wbsList.size(); i++) {
            String wbs = wbsList.get( i );
            List<String> sqlList = listOfSqlList.get( i );

            File outputFile;
            try {
                String outputFileName =  originFile.getName() + "/" + wbs + ".sql";
                outputFile = new File(outputPath + "/" + outputFileName);
                outputFile.getParentFile().mkdirs();
            } catch ( Exception e ) {
                e.printStackTrace();
                return;
            }

            try (PrintStream printStream = new PrintStream( new FileOutputStream( outputFile, false), true )) {
//                    printStream.println( "-- [ " + wbs + " ] ======== BEGIN ========" );
//                    int ctr = 0;
                    for (String sql : sqlList) {
//                        ctr++;
                        printStream.println( sql );
//                        if (ctr >= 1000) {
//                            ctr = 0;
//                            printStream.println( "commit;" );
//                        }
                    }
//                    printStream.println( "commit;" );
//                    printStream.println( "-- [ " + wbs + " ] ======== END ========\n\n" );

            } catch ( Exception e ) {
                e.printStackTrace();
            }

        }
    }

    private static List<Map<String, String>> parseRowList(File file) {
        try (OPCPackage pkg = OPCPackage.open( file )) {
            XSSFWorkbook workbook = new XSSFWorkbook(pkg);
            XSSFSheet sheet = workbook.getSheetAt( 0 );
            String sheetName = sheet.getSheetName();
            int lastRowNum = sheet.getLastRowNum();
            //System.out.println( "sheetName = " + sheetName + "\tlastRowNum = " + lastRowNum );
            List<Map<String, String>> rowList = new LinkedList<>();
            for (int i=1; i<=lastRowNum; i++) {
                XSSFRow row = sheet.getRow( i );
                int firstCellNum = row.getFirstCellNum();
                int lastCellNum = row.getLastCellNum();
                Map<String, String> cellMap = new HashMap<>(2 + lastCellNum - firstCellNum);
                cellMap.put( "firstCellNum", String.valueOf( firstCellNum ) );
                cellMap.put( "lastCellNum", String.valueOf( lastCellNum ) );
                for (int j=firstCellNum; j<lastCellNum; j++) {
                    XSSFCell cell = row.getCell( j, Row.RETURN_BLANK_AS_NULL );
                    if (cell != null) {
                        String colLabel = getIndexLabel(j);
                        cellMap.put( colLabel + i, getStringValueFromCell(cell) );
                    }
                }
                rowList.add( cellMap );
            }
            return rowList;
        } catch ( Exception e ) {
            e.printStackTrace( );
        }
        return null;
    }
    private static String getStringValueFromCell( Cell cell ) {
        SimpleDateFormat sFormat = new SimpleDateFormat( "yyyy-MM-dd" );
        DecimalFormat decimalFormat = new DecimalFormat( "#.#" );
        String cellValue = "";
        if ( cell == null ) {
            return cellValue;
        } else if ( cell.getCellType( ) == Cell.CELL_TYPE_STRING ) {
            cellValue = cell.getStringCellValue( );
        } else if ( cell.getCellType( ) == Cell.CELL_TYPE_NUMERIC ) {
            if ( DateUtil.isCellDateFormatted( cell ) ) {
                double d = cell.getNumericCellValue( );
                Date date = DateUtil.getJavaDate( d );
                cellValue = sFormat.format( date );
            } else {
                // cellValue = decimalFormat.format( ( cell.getNumericCellValue( ) ) );
                cellValue = NumberToTextConverter.toText( cell.getNumericCellValue( ) );
            }
        } else if ( cell.getCellType( ) == Cell.CELL_TYPE_BLANK ) {
            cellValue = "";
        } else if ( cell.getCellType( ) == Cell.CELL_TYPE_BOOLEAN ) {
            cellValue = String.valueOf( cell.getBooleanCellValue( ) );
        } else if ( cell.getCellType( ) == Cell.CELL_TYPE_ERROR ) {
            cellValue = "";
        } else if ( cell.getCellType( ) == Cell.CELL_TYPE_FORMULA ) {
            cellValue = cell.getCellFormula( );
        }
        return cellValue;
    }

    private static String [] COLUMN_LABEL_SOURCE = new String[] {
            "A", "B", "C", "D", "E", "F", "G",
            "H", "I", "J", "K", "L", "M", "N",
            "O", "P", "Q", "R", "S", "T",
            "U", "V", "W", "X", "Y", "Z"
    };

    /**
     * 返回该列号对应的字母
     *
     * @param columnNo
     *            (xls的)第几列（从1开始）
     */
    private static String getCorrespondingLabel(int columnNo) {
        if (columnNo < 1 ) {
            throw new IllegalArgumentException();
        }

        StringBuilder sb = new StringBuilder(5);
        int remainder = columnNo % 26;
        if (remainder == 0) {
            sb.append("Z");
            remainder = 26;
        } else {
            sb.append(COLUMN_LABEL_SOURCE[remainder - 1]);
        }

        while ((columnNo = (columnNo - remainder) / 26 - 1) > -1) {
            remainder = columnNo % 26;
            sb.append(COLUMN_LABEL_SOURCE[remainder]);
        }

        return sb.reverse().toString();
    }

    /**
     * 列号转字母
     *
     * @param columnIndex
     *            poi里xls的列号（从0开始）
     * @throws IllegalArgumentException
     *             if columnIndex less than 0
     * @return 该列对应的字母
     */
    public static String getIndexLabel(int columnIndex) {
        return getCorrespondingLabel(columnIndex + 1);
    }

    private static List<String> generateWbsList(String fileName) {
        //System.out.println( "fileName = " + fileName );
        String wbsPrefix = fileName.substring( 0, 10 );
        String wbsSuffix = fileName.substring( 13, 16 );
        String startWbs = fileName.substring( 10, 13 );
        int iStartWbs = Integer.valueOf( startWbs );
        if (MIN_WBS > 0) {
            iStartWbs = MIN_WBS;
        }
        //System.out.println( "wbsPrefix = " + wbsPrefix );
        //System.out.println( "wbsSuffix = " + wbsSuffix );
        List<String> wbsList = new LinkedList<>();
        for (int i=iStartWbs; i<=MAX_WBS; i++) {
            String iStr = String.valueOf( i );
            String prefix = "000".substring( 0, 3-iStr.length() );
            String wbs = wbsPrefix + prefix + iStr + wbsSuffix;
            //System.out.println( "i = " + i + "\twbs = " + wbs );
            wbsList.add( wbs );
        }
        return wbsList;
    }

    private static String nullToBlank(String str) {
        if (str == null) {
            return "";
        }
        return str;
    }
}
