package com.example.myapp;

import android.app.Activity;
import android.os.Bundle;
import android.os.Environment;
import android.util.Log;
import android.widget.Button;
import android.widget.Toast;

import com.android.volley.Request;
import com.android.volley.RequestQueue;
import com.android.volley.Response;
import com.android.volley.toolbox.JsonObjectRequest;
import com.android.volley.toolbox.StringRequest;
import com.android.volley.toolbox.Volley;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

public class MainActivity extends Activity {

	Button download;

	@Override
	protected void onCreate(Bundle savedInstanceState) {
		super.onCreate(savedInstanceState);
		setContentView(R.layout.activity_main);

		download = findViewById(R.id.button);
		download.setOnClickListener(v -> {
			try {
				DownloadXL();
			} catch (Exception e) {
				e.printStackTrace();
			}
		});
	}
	private void DownloadXL() {
		Map<String, String> params = new HashMap<String, String>();
		params.put("booking_CheckIn", "2021-12-20 10:41:55");
		params.put("booking_CheckOut", "2022-12-31 12:11:55");

		JSONObject parameters = new JSONObject(params);
		Log.e("CheckReq", " "+parameters );
		String url = "http://173.82.238.166/madras_beach_house/MBHBookingSystem/getbyDate";

		JsonObjectRequest jsonRequest = new JsonObjectRequest(Request.Method.POST, url, parameters,
				response -> {
					Log.e("CheckRes : "," " + response);
					try {
						JSONObject obj = new JSONObject(String.valueOf(response));
						if (obj.optString("status").equals("1")) {
							JSONArray dataArray = obj.getJSONArray("data");
							//Blank workbook
							HSSFWorkbook workbook = new HSSFWorkbook();
							//Create a Mapping Func. to map the data
							Map<Integer, Object[]> data = new TreeMap<Integer, Object[]>();

							for (int i = 0; i < dataArray.length(); i++) {
								JSONObject dataobj = dataArray.getJSONObject(i);
								Log.e("check : ", "data object :" + dataobj);

								String villa_Name = dataobj.getString("villa_Name");
								JSONArray villa_Details = dataobj.getJSONArray("villa_Details");

								//Creating sheet on workbook by Villaname
								HSSFSheet sheet = workbook.createSheet(villa_Name);
								//Header of the Sheet
								data.put(1, new Object[]{"Id", "Name", "CheckIn", "CheckOut",
										"Mobile No", "Villa Name", "Total Amount", "Advance Amount",
										"Advance Date", "Transaction Ref.No", "Admin Name"});
								//Values for the next row
								for (int j = 0; j < villa_Details.length(); j++) {
									JSONObject villaDetails_obj = villa_Details.getJSONObject(j);
									int rowid=j+2;
									Log.e("check : ", "rowid = "+rowid);
									data.put( rowid , new Object[]{
											villaDetails_obj.getString("booking_Id"),
											villaDetails_obj.getString("booking_Name"),
											villaDetails_obj.getString("booking_CheckIn"),
											villaDetails_obj.getString("booking_CheckOut"),
											villaDetails_obj.getString("booking_Mobileno"),
											villaDetails_obj.getString("villa_Name"),
											villaDetails_obj.getString("booking_TotalAmount"),
											villaDetails_obj.getString("booking_AdvanceAmount"),
											villaDetails_obj.getString("booking_AdvanceDate"),
											villaDetails_obj.getString("booking_TransactionRefno"),
											villaDetails_obj.getString("admin_Name")}
									);
								}
								//Iterating Over the Data and import all data into particular sheets
								Set<Integer> keyset = data.keySet();
								int rownum = 0;
								for (int key : keyset) {
									//Create a row on each sheet
									HSSFRow row = sheet.createRow(rownum++);
									Object[] objectarray = data.get(key);
									int cellnum = 0;
									for (Object value : objectarray) {
										//Create a cell(column) on each sheet
										HSSFCell cell = row.createCell(cellnum++);
										//set the value in each cell
										cell.setCellValue(String.valueOf(value));
									}
								}
							}
							try {
								String filePath = String.valueOf(Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOWNLOADS));
								File file = null;
								file = new File(filePath, "MBH.xlsx");
								if (!file.exists()) {
									try {
										file.delete();
										file.createNewFile();
									} catch (IOException e) {
										e.printStackTrace();
									}
								}
								Log.e("check : ", "file : " + file);
								FileOutputStream fos = new FileOutputStream(file);
								workbook.write(fos);
								fos.flush();
								fos.close();
								Toast.makeText(this, "Downloaded Successfully ", Toast.LENGTH_SHORT).show();
							} catch (Exception e) {
								e.printStackTrace();
							}
						}
					} catch (JSONException e) {
						e.printStackTrace();
					}
				},
				error -> {
					Toast.makeText(this, error.getMessage(), Toast.LENGTH_SHORT).show();
				});
		RequestQueue requestQueue = Volley.newRequestQueue(this);
		requestQueue.add(jsonRequest);
	}
}

