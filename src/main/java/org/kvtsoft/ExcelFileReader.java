package org.kvtsoft;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.io.ByteArrayInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Base64;

public class ExcelFileReader {

    public void Extraction() {
        try {
            // Decoding the encoded content "uploaded file"
            byte[] base64decodedBytes = Base64.getDecoder().decode("content");
            InputStream inputStream = new ByteArrayInputStream(base64decodedBytes);

            // Create Workbook instance holding reference to .xlsx file
            // XSSFWorkbook workbook = new XSSFWorkbook(file);

            // Get first/desired sheet from the workbook
            // XSSFSheet sheet = workbook.getSheetAt(0);

            HSSFWorkbook workbook = new HSSFWorkbook(inputStream);
            HSSFSheet sheet = workbook.getSheetAt(0);

            ArrayList<ColumnNames> candidateList = new ArrayList<>();

            // get row0 from the sheet
            Row row0 = sheet.getRow(0);

            // get cell values from 0-4 cell from the row0
            Cell ceZero = row0.getCell(0);
            Cell ceOne = row0.getCell(1);
            Cell ceTwo = row0.getCell(2);
            Cell ceThree = row0.getCell(3);

            String ceZeroString = ceZero.getStringCellValue().toString();
            String ceOneString = ceOne.getStringCellValue().toString();
            String ceTwoString = ceTwo.getStringCellValue().toString();
            String ceThreeString = ceThree.getStringCellValue().toString();
//
//            // compare the header for verifying file format
//            if (ceZeroString.equals("Name") && ceOneString.equals("Email") && ceTwoString.equals("Phone Number")
//                    && ceThreeString.equals("Vendor")) {
//
//                // ignoring header for that initializing loop from +1 (row1)
//                for (int i = sheet.getFirstRowNum() + 1; i <= sheet.getLastRowNum(); i++) {
//                    HrCandidate h = new HrCandidate();
//                    Row row = sheet.getRow(i);
//
//                    for (int cellno = row.getFirstCellNum(); cellno <= row.getLastCellNum(); cellno++) {
//                        Cell cell = row.getCell(cellno);
//
//                        if (cell != null && cell.getCellType() != Cell.CELL_TYPE_BLANK) {
//                            if (cellno == 0) {
//                                h.setName(cell.getStringCellValue());
//
//                            }
//
//                            if (cellno == 1) {
//                                h.setEmailId(cell.getStringCellValue());
//
//                            }
//
//                            if (cellno == 3) {
//                                h.setVendor(cell.getStringCellValue());
//
//                            }
//
//                            if (cellno == 2) {
//                                Integer mobile = (int) cell.getNumericCellValue();
//                                String mobileString = mobile.toString();
//                                h.setMobile(mobileString);
//
//                            }
//                        }
//                    }
//                    candidateList.add(h);
//                    response.put("candidateList size", candidateList.size());
//                    logger.log("candidateList size" + candidateList.size());
//
//                }
//
//                // mapping (setter) values into document
//                Document document = new Document();
//                List<Document> data = new ArrayList<>();
//                for (HrCandidate can : candidateList) {
//                    document = new Document();
//                    document.put("name", can.getName());
//                    document.put("email", can.getEmailId());
//                    document.put("mobile", can.getMobile());
//                    document.put("vendor", can.getVendor());
//
//                    data.add(document);
//
//                }
//
//                // check for null objects and rows
//                JSONObject object = new JSONObject();
//                object.put("data", data);
//                JSONArray dataArray = object.optJSONArray("data");
//                logger.log("dataArraySize " + dataArray.length());
//
//                boolean rowLenght = false;
//                for (int i = 0; i < dataArray.length(); i++) {
//                    JSONObject inputIndexObject = dataArray.optJSONObject(i);
//                    if (inputIndexObject.length() == 0)
//                        continue;
//
//                    rowLenght = inputIndexObject.length() == 4;
//                    if (rowLenght == false)
//                        break;
//                }
//
//                logger.log("rowLenghtState " + rowLenght);
//
//                if (rowLenght == true) {
//
//                    // below code fetch the max candidateId value and append it to the BasicDBobject
//                    // along with document values
//                    for (int i = 0; i < dataArray.length(); i++) {
//                        response.put("message", "inside for loop");
//                        JSONObject inputIndexObject = dataArray.optJSONObject(i);
//                        if (inputIndexObject.length() == 0)
//                            continue;
//
//                        String name = inputIndexObject.optString("name");
//                        String email = inputIndexObject.optString("email");
//                        String mobile = inputIndexObject.optString("mobile");
//                        String vendor = inputIndexObject.optString("vendor");
//
//                        BasicDBObject eachData = new BasicDBObject();
//
//                        // sort Document with individual max requiremenId from collection
//                        BasicDBObject sort = new BasicDBObject("candidateId", -1);
//                        List<Document> user_detailList = (List<Document>) user_details.find().sort(sort).limit(1)
//                                .into(new ArrayList<Document>());
//
//                        JSONObject userObject = new JSONObject();
//                        userObject.put("data", user_detailList);
//                        JSONArray userdataArray = userObject.optJSONArray("data");
//
//                        long maxReqIddb = 0;
//
//                        // as array may be empty due to no data in collection
//                        if (dataArray.length() > 0) {
//                            JSONObject userIndexObject = userdataArray.optJSONObject(0);
//                            maxReqIddb = userIndexObject.optLong("candidateId");
//                        }
//
//                        eachData.append("candidateId", maxReqIddb + 1);
//
//                        eachData.append("name", name);
//                        eachData.append("email", email);
//                        eachData.append("mobile", mobile);
//                        eachData.append("vendor", vendor);
//                        eachData.append("appliedBy", "Hr");

        } catch (Exception e) {
            System.out.println("Exception: " + e.getMessage());
        }
    }
}
