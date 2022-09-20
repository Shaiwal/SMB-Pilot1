package com.cisco.smb;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Currency;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Locale;
import java.util.Set;

import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.filechooser.FileSystemView;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.bson.Document;
import org.bson.conversions.Bson;
import org.json.JSONException;
import org.json.JSONObject;

import com.mongodb.MongoClient;
import com.mongodb.MongoClientOptions;
import com.mongodb.MongoCommandException;
import com.mongodb.MongoCredential;
import com.mongodb.ServerAddress;
import com.mongodb.client.FindIterable;
import com.mongodb.client.MongoCollection;
import com.mongodb.client.MongoCursor;
import com.mongodb.client.MongoDatabase;
import com.mongodb.client.model.Filters;
import com.mongodb.client.model.UpdateOptions;
import com.mongodb.client.result.UpdateResult;

public class SMBInit {

	public static void main(String[] args) {

		MongoClient mongoClient = null;
		final JPanel panel = new JPanel();
		StringBuilder strB = new StringBuilder();

		try {

			MongoCredential credential;
			// credential = MongoCredential.createCredential("cdcmodify", "bookmarksadmin",
			// "Mongo#1234".toCharArray());
			credential = MongoCredential.createCredential("cdcmodify", "cdcsmb", "2Jay5ssestg11".toCharArray());
			System.out.println("Connected to the database successfully");
		//	MongoClientOptions mongoClientOptions = new MongoClientOptions.Builder().build();
			 List list = Arrays.asList(new ServerAddress("mgdb-cdcstg-npd1-1", 27060), new
			 ServerAddress("mgdb-cdcstg-npd2-1", 27060), new
			 ServerAddress("mgdb-cdcstg-npd3-1", 27060));
//			List<ServerAddress> list = Arrays.asList(
//					new ServerAddress("cdcsmbprodrepo-nqonxs3xdev-3-cdcsmbprodrepo.cloudapps.cisco.com", 35641));
			mongoClient = new MongoClient(list, Arrays.asList(credential));
			// mongoClient = new MongoClient(list, credential, mongoClientOptions);

			// Accessing the database
			MongoDatabase database = mongoClient.getDatabase("cdcsmb");

			// Retrieving a collection
			MongoCollection<Document> collection = database.getCollection("cdc_guide_filter_product_data_test");
			System.out.println("Collection myCollection selected successfully" + collection.count());

			JFileChooser jfc = new JFileChooser(FileSystemView.getFileSystemView().getHomeDirectory());

			jfc.setDialogTitle("Choose SMB PriceSpider Excel: ");
			jfc.setFileSelectionMode(JFileChooser.FILES_ONLY);

			jfc.setAcceptAllFileFilterUsed(false);
			FileNameExtensionFilter filter = new FileNameExtensionFilter("Excel Files", "xls", "xlsx");
			jfc.addChoosableFileFilter(filter);

			int returnValue = jfc.showOpenDialog(null);

			if (returnValue == JFileChooser.APPROVE_OPTION) {
				File selectedFile = jfc.getSelectedFile();
				System.out.println(selectedFile.getAbsolutePath());

				List<JSONObject> skuJSONArr = new ArrayList<JSONObject>();
				JSONObject skuJSONObj = new JSONObject();
				int skuColumnNum = 0, sellerURLColumnNum = 0, sellerPriceColumnNum = 0, sellerNameColumnNum = 0,
						sellerCountryColumnNum = 0, sellerCurrenyColumnNum = 0;
				String countryISOCode = "";
				Set<String> uniqueSKU = new HashSet<String>();

				FileInputStream file;

				file = new FileInputStream(selectedFile);

				// Create Workbook instance holding reference to .xlsx file
				XSSFWorkbook workbook = new XSSFWorkbook(file);
				// Get first/desired sheet from the workbook
				XSSFSheet sheet = workbook.getSheetAt(0);
				System.out.println(sheet.getSheetName());
				// Iterate through each rows one by one
				Iterator<Row> rowIterator = sheet.iterator();
				while (rowIterator.hasNext()) {
					Row row = rowIterator.next();
					// System.out.println(row.getRowNum() + " .. " + skuColumnNum + " .. " +
					// sellerURLColumnNum + " .. " + sellerPriceColumnNum);
					Iterator<Cell> cellIterator = row.cellIterator();
					skuJSONObj = new JSONObject();
					// skuJSONObj.put("seller_name", "Bechtle");
					// skuJSONObj.put("seller_country", "Germany");
					skuJSONObj.put("seller_logo", "/c/dam/cdc/sse/pop-up-image.png");
					// skuJSONObj.put("price_currency", "€");
					// skuJSONObj.put("price_currency_code", "EUR");
					while (cellIterator.hasNext()) {
						Cell cell = cellIterator.next();
						if (row.getRowNum() == 0) {
							if (cell.getCellType() == Cell.CELL_TYPE_STRING
									&& cell.getStringCellValue().equalsIgnoreCase("SKU")) {
								skuColumnNum = cell.getColumnIndex();
							} else if (cell.getCellType() == Cell.CELL_TYPE_STRING
									&& cell.getStringCellValue().equalsIgnoreCase("URL")) {
								sellerURLColumnNum = cell.getColumnIndex();
							} else if (cell.getCellType() == Cell.CELL_TYPE_STRING
									&& cell.getStringCellValue().equalsIgnoreCase("UNIT PRICE")) {
								sellerPriceColumnNum = cell.getColumnIndex();
							} else if (cell.getCellType() == Cell.CELL_TYPE_STRING
									&& cell.getStringCellValue().equalsIgnoreCase("SELLER NAME")) {
								sellerNameColumnNum = cell.getColumnIndex();
							} else if (cell.getCellType() == Cell.CELL_TYPE_STRING
									&& cell.getStringCellValue().equalsIgnoreCase("SELLER COUNTRY")) {
								sellerCountryColumnNum = cell.getColumnIndex();
							} else if (cell.getCellType() == Cell.CELL_TYPE_STRING
									&& cell.getStringCellValue().equalsIgnoreCase("CURRENCY")) {
								sellerCurrenyColumnNum = cell.getColumnIndex();
							}
						} else {
							if (cell.getColumnIndex() == skuColumnNum) {
								skuJSONObj.put("product_pid", cell.getStringCellValue());
							} else if (cell.getColumnIndex() == sellerURLColumnNum) {
								skuJSONObj.put("seller_product_link", cell.getStringCellValue());
							} else if (cell.getColumnIndex() == sellerPriceColumnNum) {
								if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
									skuJSONObj.put("seller_price", cell.getStringCellValue());
								} else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
									skuJSONObj.put("seller_price", cell.getNumericCellValue());
								}
							} else if (cell.getColumnIndex() == sellerNameColumnNum) {
								skuJSONObj.put("seller_name", cell.getStringCellValue());
							} else if (cell.getColumnIndex() == sellerCountryColumnNum) {
								Locale l = new Locale("", cell.getStringCellValue());
								countryISOCode = cell.getStringCellValue();
								skuJSONObj.put("seller_country", l.getDisplayCountry());
							} else if (cell.getColumnIndex() == sellerCurrenyColumnNum) {
								skuJSONObj.put("price_currency", cell.getStringCellValue());
								Locale l = new Locale("", countryISOCode);
								Currency c = Currency.getInstance(l);
								skuJSONObj.put("price_currency_code", c.getCurrencyCode());
							}
						}

					}
					if (row.getRowNum() > 0) {
						skuJSONArr.add(skuJSONObj);
					}
				}
				Iterator<JSONObject> ite = skuJSONArr.iterator();
				List<Document> documentList = new ArrayList<Document>();
				String prevSKU = "";
				while (ite.hasNext()) {
					JSONObject js = (JSONObject) ite.next();
				//	System.out.println("JSON inside Arr>> " + js.toString());

//					  BasicDBObject searchQuery = new BasicDBObject().append("product_type",
//					  "smb").append("product_pid", js.get("product_pid"));
					Document updateDocument = new Document();
					
					if (uniqueSKU.add((String) js.get("product_pid")) && prevSKU != "") {
						String s = updateDB(collection, documentList, prevSKU);
						if(s!=null)
							if(strB.length() % 50 > 25)
								strB.append("\n, " + s);
							else
								strB.append(", " + s);
						
						documentList = new ArrayList<Document>();
						prevSKU = (String) js.get("product_pid");
					}
					
					prevSKU = (String) js.get("product_pid");
					updateDocument.append("seller_name", String.valueOf(js.get("seller_name")).split("\\(")[0].trim());
					updateDocument.append("seller_country", js.get("seller_country"));
					updateDocument.append("seller_logo", "/c/dam/cdc/sse/pop-up-image.png");
					if (js.has("seller_product_link"))
						updateDocument.append("seller_product_link", js.get("seller_product_link"));
					else
						updateDocument.append("seller_product_link", "Not Available");

					if (js.has("seller_price"))
						updateDocument.append("seller_price", js.get("seller_price"));
					else
						updateDocument.append("seller_price", "Not Available");
					
					if (js.has("price_currency"))
						updateDocument.append("price_currency", js.get("price_currency"));
					else
						updateDocument.append("price_currency", "Not Available");
					
					if (js.has("price_currency_code"))
						updateDocument.append("price_currency_code", js.get("price_currency_code"));
					else
						updateDocument.append("price_currency_code", "Not Available");

					
					documentList.add(updateDocument);
					
				}
				if(documentList.size()>0) {
					String s = updateDB(collection, documentList, prevSKU);
					if(s!=null)
						strB.append(", " + s);
				}
			}

		} catch (FileNotFoundException e) {
			e.printStackTrace();
			JOptionPane.showMessageDialog(panel, "FileNotFoundException " + e.getMessage(), "Error",
					JOptionPane.ERROR_MESSAGE);
		} catch (IOException e) {
			e.printStackTrace();
			JOptionPane.showMessageDialog(panel, "IOException " + e.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
		} catch (JSONException e) {
			e.printStackTrace();
			JOptionPane.showMessageDialog(panel, "JSONException " + e.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
		} catch (MongoCommandException e) {
			e.printStackTrace();
			JOptionPane.showMessageDialog(panel, "MongoCommandException " + e.getMessage(), "Error",
					JOptionPane.ERROR_MESSAGE);
		} catch (IllegalArgumentException e) {
			e.printStackTrace();
			JOptionPane.showMessageDialog(panel, "Exception " + e.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
		} catch (Exception e) {
			e.printStackTrace();
			JOptionPane.showMessageDialog(panel, "Exception " + e.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
		} finally {
			mongoClient.close();
			JOptionPane.showMessageDialog(panel, "Completed for following SKU(s):  \n" + strB.substring(1, strB.length()), "Success",
					JOptionPane.INFORMATION_MESSAGE);
		}
	}
	
	private static String updateDB(MongoCollection<Document> collection, List<Document> documentList, String product_pid ) {
		
		//System.out.println("INSIDE UPDATEDB"+product_pid);
		Bson filterQuery = Filters.and(Filters.eq("product_type", "smb"),
				Filters.eq("product_pid", product_pid));

		Bson setUpdateQuery = new Document().append("$set",
				new Document().append("seller_pricing", documentList));
		FindIterable<Document> fi = collection.find(filterQuery);
		MongoCursor<Document> mr = fi.iterator();
		while (mr.hasNext()) {
			Document d = (Document) mr.next();
			System.out.println(d.get("_id") + " .. " + d.get("product_type") + " .. " + d.get("product_pid"));

			  UpdateOptions options = new UpdateOptions(); 
			  options.upsert(false);
			  UpdateResult updatedDoc = collection.updateOne(filterQuery, setUpdateQuery, options);//(searchQuery, setNewFieldQuery);
			  System.out.println(updatedDoc.getMatchedCount() + " ... " + updatedDoc.getUpsertedId()); 
			 // System.out.println(updatedDoc.get("_id") + " .. " + updatedDoc.get("product_type") + " .. " + updatedDoc.get("product_pid") + " .. " + updatedDoc.get("seller_pricing"));
			  return d.getString("product_pid");

		}
		
		return null;
	}
}
