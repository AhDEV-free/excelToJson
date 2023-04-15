package com.example.exceltojson.util;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;


/**
 * ExcelToJSON 工具
 * 树结构
 * 关联数据 SheetName#{Sheetname!ROWNumber}
 * eg: sheet: sheet2#{sheet1}
 * row:   sheet2#{sheet1!A1}
 * header: 第一二行当作标题行，第一行为重命名，默认使用第二行，中文&英文
 */
public class ExcelToJSONUtil {

    public static JSONArray exportData(InputStream in) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook(in);
        int numberOfSheets = workbook.getNumberOfSheets();

        // 树结构
        List<JSONObject> objList = new ArrayList<>();
        Map<String, JSONObject> objMap2 = new HashMap<>();
        for (int i = 0; i < numberOfSheets; i++) {
            XSSFSheet sheetAt = workbook.getSheetAt(i);

            String sheetName = sheetAt.getSheetName();
            Custom customSheet = new ExcelToJSONUtil().convertContentTo(sheetName);
            JSONObject sheetObj = new JSONObject();

            sheetObj.put("_pCustom", customSheet);

            XSSFRow row;
            XSSFCell cell;
            // get header value row
            row = sheetAt.getRow(0);
            if (row == null) {
                row = sheetAt.getRow(1);
                if (row == null) {
                    continue;
                }
            }
            Map<String, Object> objMap = new HashMap<>();
            for (int i1 = 0; i1 <= row.getLastCellNum(); i1++) {
                cell = row.getCell(i1);
                if (cell != null) {
                    objMap.put(String.valueOf(i1), getCellValue(cell));
                }
            }
            // get data
            JSONArray rows = new JSONArray();
            for (int i2 = 2; i2 <= sheetAt.getLastRowNum(); i2++) {
                row = sheetAt.getRow(i2);
                JSONObject rowObj = new JSONObject();
                for (int i3 = 0; i3 <= row.getLastCellNum(); i3++) {
                    cell = row.getCell(i3);
                    if (cell != null) {


                        rowObj.put(String.valueOf(objMap.get(String.valueOf(i3))), getCellValue(cell));
                    }
                }
                rows.add(rowObj);

            }

            sheetObj.put(customSheet.getName(), rows);
            objList.add(sheetObj);
            objMap2.put(customSheet.getId(), sheetObj);
        }

        // 封装JSON
        Iterator<JSONObject> iterator = objList.iterator();
        while (iterator.hasNext()) {
            JSONObject jsonObj = iterator.next();
            Custom pCustom = jsonObj.getObject("_pCustom", Custom.class);
            JSONArray rowsA = jsonObj.getJSONArray(pCustom.getName());
            if (pCustom != null) {
                JSONObject jsonObjectP = objMap2.get(pCustom.getPid());
                if (jsonObjectP != null) {
                    String toNameAs = pCustom.getToName();

                    JSONArray rows = jsonObjectP.getJSONArray(toNameAs);
                    if (rows != null) {
                        for (int i = 0; i < rows.size(); i++) {
                            JSONObject jsonObject = rows.getJSONObject(i);
                            if (pCustom.getToRow() != null && !pCustom.getToRow().equals("")) {
                                if (pCustom.getToRow().equals(String.valueOf(i + 2 + 1))) {
                                    // 同名的添加
                                    if (jsonObject.containsKey(toNameAs)) {
                                        JSONArray jsonArray = jsonObject.getJSONArray(toNameAs);
                                        jsonArray.addAll(rowsA);
                                    } else {
                                        jsonObject.put(pCustom.getName(), rowsA);
                                    }
                                }
                            } else {
                                // 同名的添加
                                if (jsonObject.containsKey(toNameAs)) {
                                    JSONArray jsonArray = jsonObject.getJSONArray(toNameAs);
                                    jsonArray.addAll(rowsA);
                                } else {
                                    jsonObject.put(pCustom.getName(), rowsA);
                                }
                            }
                        }
                    }

                    iterator.remove();
                }
            }
            jsonObj.remove("_pCustom");
        }

        return JSONArray.parseArray(objList.toString());

    }

    /**
     * 获取对象内容
     */
    public static Object getCellValue(XSSFCell cell) {
        Object value = null;
        switch (cell.getCellType()) {
            case NUMERIC: {
                value = cell.getNumericCellValue();
                break;
            }
            case STRING: {
                value = cell.getStringCellValue();
                break;
            }
        }

        return value;
    }


    /**
     * 获取连接内容
     *
     * @return [0] sheet 页
     */
    private Custom convertContentTo(String content) {
        Custom result = new Custom();
        result.setId(content);
        // {} to sheet
        /**
         * 匹配子数据
         * {} 关联数据
         * */
        String regex = "\\{.*}";
        Pattern pattern = Pattern.compile(regex);
        Matcher matcher = pattern.matcher(content);
        // Loop through all matches and print the parts of the input string that are not matched
        int lastMatchEnd = 0;
        String matchedString = "";
        String nonMatchedString = "";
        if (matcher.find()) {
            matchedString = matcher.group();
            int matchStart = matcher.start();
            int matchEnd = matcher.end();
            nonMatchedString += content.substring(lastMatchEnd, matchStart);
            lastMatchEnd = matchEnd;
        } else {
            nonMatchedString += content;
        }

        // Print the final part of the input string that is not matched
        String finalNonMatchedString = nonMatchedString;
        if (finalNonMatchedString.equals("")) {
            result.setName("_item");
        } else {
            result.setName(finalNonMatchedString);
        }

        matchedString = matchedString.replaceAll("^\\{", "");
        matchedString = matchedString.replaceAll("}$", "");

        // !number
        String regex3 = "!\\d$";
        Pattern pattern3 = Pattern.compile(regex3);
        Matcher matcher3 = pattern3.matcher(matchedString);
        if (matcher3.find()) {
            String matchStr = matcher3.group();
            result.setToRow(matchStr.replace("!", ""));
            String pidStr = matchedString.replaceAll(matchStr + "$", "");
            result.setPid(pidStr);
            result.setToName(getRegexStrPlus("\\{.*}", pidStr));
        } else {
            result.setPid(matchedString);
            result.setToName(getRegexStrPlus("\\{.*}", matchedString));
        }
        return result;
    }

    /**
     * 获取匹配字符串的剩余字符串
     */
    public static String getRegexStrPlus(String regex, String text) {
        String result = "";
        Pattern pattern3 = Pattern.compile(regex);
        Matcher matcher3 = pattern3.matcher(text);
        int lastMatchEnd = 0;
        String matchedString = "";
        String nonMatchedString = "";
        if (matcher3.find()) {
            matchedString = matcher3.group();
            int matchStart = matcher3.start();
            int matchEnd = matcher3.end();
            nonMatchedString += text.substring(lastMatchEnd, matchStart);
            lastMatchEnd = matchEnd;
        } else {
            nonMatchedString += text;
        }

        // Print the final part of the input string that is not matched
        String finalNonMatchedString = nonMatchedString;
        if (finalNonMatchedString.equals("")) {
            result  = "_item";
        } else {
            result = finalNonMatchedString;
        }

        return result;
    }


    /**
     * 自定义
     */
    public class Custom {
        /**
         * id
         */
        private String id = "";
        /**
         * pid
         */
        private String pid = "";
        /**
         * sheet名
         * key
         */
        private String name = "";
        /**
         * 关联行
         */
        private String toRow = "";
        /**
         * 关联列
         */
        private String toCol = "";
        /**
         * 关联 sheet名
         */
        private String ToName = "";

        public String getId() {
            return id;
        }

        public void setId(String id) {
            this.id = id;
        }

        public String getPid() {
            return pid;
        }

        public void setPid(String pid) {
            this.pid = pid;
        }

        public String getName() {
            return name;
        }

        public void setName(String name) {
            this.name = name;
        }

        public String getToRow() {
            return toRow;
        }

        public void setToRow(String toRow) {
            this.toRow = toRow;
        }

        public String getToCol() {
            return toCol;
        }

        public void setToCol(String toCol) {
            this.toCol = toCol;
        }

        public String getToName() {
            return ToName;
        }

        public void setToName(String toName) {
            ToName = toName;
        }
    }
}
