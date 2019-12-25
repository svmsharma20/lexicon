import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONArray;
import org.json.JSONObject;

import javax.net.ssl.HttpsURLConnection;
import java.io.*;
import java.net.MalformedURLException;
import java.net.URL;
import java.util.Iterator;
import java.util.Set;

public class Main
{
    public static final String INPUT_FILE ="/home/shivam/Downloads/G.xlsx";
    public static final String OUTPUT_FILE ="/home/shivam/Downloads/O.xlsx";

    public static void main(String[] args) throws IOException
    {
        XSSFWorkbook inputWorkbook = null;
        XSSFWorkbook outputWorkbook = null;
        try
        {
            OPCPackage opcPackage = OPCPackage.open(new File(INPUT_FILE));
            inputWorkbook = new XSSFWorkbook(opcPackage);
            XSSFSheet readSheet = inputWorkbook.getSheet("Wordlist-1");

//            OPCPackage opcPackage2 = OPCPackage.open(new File(OUTPUT_FILE));
//            outputWorkbook = new XSSFWorkbook(opcPackage);
              outputWorkbook = new XSSFWorkbook();

            int predictedRowCount = readSheet.getPhysicalNumberOfRows();

//            for(int i=0; i<predictedRowCount; i++)
//            {
//                XSSFRow row = readSheet.getRow(i);
//                if(row.getCell(0).getR)
//            }

            Font font = outputWorkbook.createFont();
            font.setFontHeightInPoints((short)14);
            font.setFontName("Open Sans");

            CellStyle style = outputWorkbook.createCellStyle();
            style.setFont(font);
            style.setWrapText(true);

            Iterator<Row> iterator = readSheet.iterator();
            int count=0;
//            iterator.next();
            while(iterator.hasNext())
            {
                Row row = iterator.next();
                Cell cell0 = row.getCell(0);
                Cell cell1 = row.getCell(1);

                if(cell0 == null || cell1 == null)
                    break;

                String sheetName = cell0.getStringCellValue().trim();
                String wordlist = cell1.getStringCellValue().trim();

                if(wordlist.length()<1 || sheetName.length()<1)
                    break;

                System.out.println("Creating sheet: "+sheetName);

                XSSFSheet sheet = outputWorkbook.createSheet(sheetName);
//                sheet.setColumnWidth(0,1800);
//                sheet.setColumnWidth(1,1800);
//                sheet.setColumnWidth(2,1800);
                sheet.autoSizeColumn(1);
                sheet.autoSizeColumn(2);
                int rowNum = 0;


                String wordsArray[] = wordlist.split(",");
                for(String word : wordsArray)
                {
                    word = word.trim();
                    System.out.println("\tfetching word "+word+"...");
//                    System.out.println(word + ":\n"+getMeaning(word));

                    Row inputRow = sheet.createRow(rowNum++);
                    inputRow.setHeightInPoints((short)21);
                    Cell sNoCell = inputRow.createCell(0);
                    Cell wordCell = inputRow.createCell(1);
                    Cell meaningCell = inputRow.createCell(2);

                    sNoCell.setCellStyle(style);
                    wordCell.setCellStyle(style);
                    meaningCell.setCellStyle(style);

                    sNoCell.setCellValue(rowNum);
                    wordCell.setCellValue(word);
                    meaningCell.setCellValue(getMeaning(word).toLowerCase());

                }
                FileOutputStream outputStream = new FileOutputStream(OUTPUT_FILE);
                outputWorkbook.write(outputStream);
//                System.out.println(sheetName+" : "+wordlist);
            }




        } catch (IOException | InvalidFormatException e)
        {
            e.printStackTrace();
        }
        finally
        {
            if(inputWorkbook != null)
                inputWorkbook.close();
        }
    }


    public static String getRequest(String link)
    {
        try{
            URL url=new  URL(link);
            HttpsURLConnection urlConnection=(HttpsURLConnection) url.openConnection();
            urlConnection.setRequestProperty("Accept", "application/json");

            // read the output from the server
            BufferedReader reader=new BufferedReader(new InputStreamReader(urlConnection.getInputStream()));
            StringBuilder stringBuilder=new StringBuilder();
            String meaning = "";
            String line = "";
            while ((line=reader.readLine()) !=null){
                stringBuilder.append(line + "\n");
            }
            return stringBuilder.toString();
        }catch (IOException e){
            e.printStackTrace();
            return e.toString();
        }
    }

    public static String buildURL(String word)
    {
        String urlStr = "https://googledictionaryapi.eu-gb.mybluemix.net/?define="+word+"&lang=en";
//        String urlStr = "https://dictionaryapi.com/api/v3/references/sd3/json/"+word+"?key=96b4227b-61de-4dea-b1f8-b97ce7578e1f";
        return urlStr;
    }

    public static String getMeaning(String word)
    {
        String url = buildURL(word.trim());
//        String  = "{ 'M':"+ getRequest(url)+"}";
        String jsonObject = getRequest(url);

        JSONObject meaningJSONObject = (JSONObject)((JSONObject)(new JSONArray(jsonObject.trim()).get(0))).get("meaning");

//        return meaningJSONObject.get("M.meaning").toString();
//        JSONObject meaningObject = meaningJSONObject.get(0)
        StringBuilder meaning = new StringBuilder();
        Set<String> keySet = meaningJSONObject.keySet();
        boolean newLineFlag = false;
        boolean indexFlag = keySet.size()>1 ? true : false;
        int i=1;
        for(String key: keySet)
        {
            if(newLineFlag)
                meaning.append("\n");
            if(indexFlag)
                meaning.append(i++ + ". ");
            meaning.append(((JSONObject)((JSONArray)meaningJSONObject.get(key)).get(0)).get("definition").toString());
            newLineFlag=true;
        }


        return meaning.toString();
    }
}
