package com.guangn.mylibrary;

import android.app.Activity;
import android.content.ContentUris;
import android.content.Context;
import android.database.Cursor;
import android.net.Uri;
import android.os.Build;
import android.os.Environment;
import android.os.StatFs;
import android.provider.DocumentsContract;
import android.provider.MediaStore;
import android.util.Log;
import android.widget.Toast;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

import jxl.Sheet;
import jxl.Workbook;
import jxl.format.Colour;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

/**
 * @author Guangnian
 * @data 2019-11-04 10:06:10
 * */
public class ExcelUtil {
	//导出地址
	public static final String ExcleUrl = Environment.getExternalStorageDirectory().toString() + File.separator  + File.separator + "ExcelDirectory";
	//内存地址
	public static String root = Environment.getExternalStorageDirectory()
			.getPath();

	/**
	 * context  	上下文
	 * fileName		文件名
	 * mListData	数据集
	 * ExcleUrl		存储路径
	 * */
	public static void writeExcel(Context context, String fileName, List<mData>mListData, String ExcleUrl) throws Exception {
		if (!Environment.getExternalStorageState().equals(Environment.MEDIA_MOUNTED)&&getAvailableStorage()>1000000) {
			Toast.makeText(context, "SD卡不可用", Toast.LENGTH_LONG).show();
			return;
		}
		File file;
		File dir = new File(ExcleUrl);
		file = new File(dir, fileName + ".xls");
		if (!dir.exists()) {
			dir.mkdirs();
		}
		// 创建Excel工作表
		WritableWorkbook wwb;
		OutputStream os = new FileOutputStream(file);
		wwb = Workbook.createWorkbook(os);
		// 添加第一个工作表并设置第一个Sheet的名字
		for(int i=0;i<mListData.size();i++){
			WritableSheet sheet = wwb.createSheet(mListData.get(i).SheetName, i);
			List<List<String>> ListData =  mListData.get(i).mListData;
			//遍历内容
			int h=0;
			for (List<String> d:ListData) {
				int l=0;
				for(String s:d){
					Label labels = new Label(l, h, s);
					sheet.addCell(labels);
					l++;
				}
				h++;
			}
		}
		Toast.makeText(context, "导出成功，路径："+ExcleUrl, Toast.LENGTH_LONG).show();
		// 写入数据
		wwb.write();
		// 关闭文件
		wwb.close();
	}

	/**
	 * context  	上下文
	 * fileName		文件名
	 * mListData	数据集
	 * ExcleUrl		存储路径
	 * */
	/**
	 *  支持03以上
	 * */
	public static void writeExcel03(Context context, String fileName, List<mData>mListData, String ExcleUrl) throws Exception {
		if (!Environment.getExternalStorageState().equals(Environment.MEDIA_MOUNTED)&&getAvailableStorage()>1000000) {
			Toast.makeText(context, "SD卡不可用", Toast.LENGTH_LONG).show();
			return;
		}
		File file;
		File dir = new File(ExcleUrl);
		file = new File(dir, fileName + ".xls");
		if (!dir.exists()) {
			dir.mkdirs();
		}
		HSSFWorkbook wb = new HSSFWorkbook();
		OutputStream os = new FileOutputStream(file);
		for(int i=0;i<mListData.size();i++){
			HSSFSheet sheet = wb.createSheet(mListData.get(i).SheetName);
			List<List<String>> ListData =  mListData.get(i).mListData;
			//遍历内容
			int h=0;
			for (List<String> d:ListData) {
				int l=0;
				for(String s:d){
					HSSFRow row1 = sheet.createRow(h);   //--->创建一行
					HSSFCell cell1 = row1.createCell((short)l);   //--->创建一个单元格
					cell1.setCellValue(s);
					l++;
				}
				h++;
			}
		}

		Toast.makeText(context, "导出成功，路径："+ExcleUrl, Toast.LENGTH_LONG).show();
		// 写入数据
		wb.write(os);
		// 关闭文件
		wb.close();
	}

	/**
	 * activity  	上下文
	 * filepath		文件名
	 * */
	/**
	 * 查询excel表格结果
	 * */
	public static List<mData> QueryUser(Activity activity, Uri filepath){
		InputStream is = null;
		Workbook workbook = null;
		List<mData> mData = new ArrayList<>();
		try {
			File dir=null;
				if (filepath != null) {
					String path = ExcelUtil.getPath(activity, filepath);
					File file = new File(path);
					if (file.exists()) {
						if(file.getName().contains(".xls")||file.getName().contains(".xlsx")){
							dir = new File( file.toString());
						}
					}
				}
			if(dir==null){
				return mData;
			}
			is = new FileInputStream(dir.getPath());//获取流
			workbook = Workbook.getWorkbook(is);
			int i=0;
			for(String n:workbook.getSheetNames()){
				mData mdata = new mData();
				mdata.SheetName=n;
				Sheet sheet = workbook.getSheet(i);
				for(int h = 0; h <sheet.getRows(); h++){
					List<String> mD =new ArrayList<>();
					for(int l=0;l<sheet.getColumns()-1;l++){
						mD.add(sheet.getCell(l,h ).getContents());
					}
					mdata.mListData.add(mD);
				}
				mData.add(mdata);
				i++;
			}
		} catch (Exception e) {
			e.printStackTrace();
			Log.d("error",e+"");
		} finally {
			if (workbook != null) {
				workbook.close();
			}
			if (is != null) {
				try {
					is.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return mData;
	}

	/**
	 * activity  	上下文
	 * filepath		文件名
	 * */
	/**
	 *  支持查询03以上
	 * */
	public List<mData> mData = new ArrayList<>();
	public List<mData> readExcel03(String fileName) {
		try {
			InputStream inputStream = new FileInputStream(fileName);
			org.apache.poi.ss.usermodel.Workbook workbook;
			if (fileName.endsWith(".xls")) {
				workbook = new HSSFWorkbook(inputStream);
			} else if (fileName.endsWith(".xlsx")) {
				workbook = new XSSFWorkbook(inputStream);
			} else {
				return mData;
			}
			for(int i=0;i<workbook.getNumberOfSheets();i++){
				mData mdata = new mData();
				mdata.SheetName=workbook.getSheetName(i);
				org.apache.poi.ss.usermodel.Sheet sheet = workbook.getSheetAt(i);
				FormulaEvaluator formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();
				for(int h = 0; h <sheet.getPhysicalNumberOfRows(); h++){
					List<String> mD =new ArrayList<>();
					Row row = sheet.getRow(h);
					for(int l=0;l< row.getPhysicalNumberOfCells()-1;l++){
						CellValue v0 = formulaEvaluator.evaluate(row.getCell(l));
						if(v0!=null){
							mD.add(v0.getStringValue());
						}else{
							mD.add("");
						}
					}
					mdata.mListData.add(mD);
				}
				mData.add(mdata);
			}
			workbook.close();
		} catch (IOException e) {
			e.printStackTrace();
			Log.d("error",e+"");
		}
		return mData;
	}

	/**
	 * 路径转换
	 * */
	public static String getPath(final Context context, final Uri uri) {

		final boolean isKitKat = Build.VERSION.SDK_INT >= Build.VERSION_CODES.KITKAT;

		// DocumentProvider
		if (isKitKat && DocumentsContract.isDocumentUri(context, uri)) {
			// ExternalStorageProvider
			if (isExternalStorageDocument(uri)) {
				final String docId = DocumentsContract.getDocumentId(uri);
//                Log.i(TAG,"isExternalStorageDocument***"+uri.toString());
//                Log.i(TAG,"docId***"+docId);
//                以下是打印示例：
//                isExternalStorageDocument***content://com.android.externalstorage.documents/document/primary%3ATset%2FROC2018421103253.wav
//                docId***primary:Test/ROC2018421103253.wav
				final String[] split = docId.split(":");
				final String type = split[0];

				if ("primary".equalsIgnoreCase(type)) {
					return Environment.getExternalStorageDirectory() + "/" + split[1];
				}
			}
			// DownloadsProvider
			else if (isDownloadsDocument(uri)) {
//                Log.i(TAG,"isDownloadsDocument***"+uri.toString());
				final String id = DocumentsContract.getDocumentId(uri);
				final Uri contentUri = ContentUris.withAppendedId(
						Uri.parse("content://downloads/public_downloads"), Long.valueOf(id));

				return getDataColumn(context, contentUri, null, null);
			}
			// MediaProvider
			else if (isMediaDocument(uri)) {
//                Log.i(TAG,"isMediaDocument***"+uri.toString());
				final String docId = DocumentsContract.getDocumentId(uri);
				final String[] split = docId.split(":");
				final String type = split[0];

				Uri contentUri = null;
				if ("image".equals(type)) {
					contentUri = MediaStore.Images.Media.EXTERNAL_CONTENT_URI;
				} else if ("video".equals(type)) {
					contentUri = MediaStore.Video.Media.EXTERNAL_CONTENT_URI;
				} else if ("audio".equals(type)) {
					contentUri = MediaStore.Audio.Media.EXTERNAL_CONTENT_URI;
				}

				final String selection = "_id=?";
				final String[] selectionArgs = new String[]{split[1]};

				return getDataColumn(context, contentUri, selection, selectionArgs);
			}
		}
		// MediaStore (and general)
		else if ("content".equalsIgnoreCase(uri.getScheme())) {
//            Log.i(TAG,"content***"+uri.toString());
			return getDataColumn(context, uri, null, null);
		}
		// Files
		else if ("file".equalsIgnoreCase(uri.getScheme())) {
//            Log.i(TAG,"file***"+uri.toString());
			return uri.getPath();
		}
		return null;
	}

	public static WritableCellFormat getHeader() {
		WritableFont font = new WritableFont(WritableFont.TIMES, 10,
				WritableFont.BOLD);// 定义字体
		try {
			font.setColour(Colour.BLUE);// 蓝色字体
		} catch (WriteException e1) {
			e1.printStackTrace();
		}
		WritableCellFormat format = new WritableCellFormat(font);
		try {
			format.setAlignment(jxl.format.Alignment.CENTRE);// 左右居中
			format.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);// 上下居中
			// format.setBorder(Border.ALL, BorderLineStyle.THIN,
			// Colour.BLACK);// 黑色边框
			// format.setBackground(Colour.YELLOW);// 黄色背景
		} catch (WriteException e) {
			e.printStackTrace();
		}
		return format;
	}

	/** 获取SD可用容量 */
	private static long getAvailableStorage() {

		StatFs statFs = new StatFs(root);
		long blockSize = statFs.getBlockSize();
		long availableBlocks = statFs.getAvailableBlocks();
		long availableSize = blockSize * availableBlocks;
		// Formatter.formatFileSize(context, availableSize);
		return availableSize;
	}

	/**
	 * Get the value of the data column for this Uri. This is useful for
	 * MediaStore Uris, and other file-基础地图 ContentProviders.
	 *
	 * @param context       The context.
	 * @param uri           The Uri to query.
	 * @param selection     (Optional) Filter used in the query.
	 * @param selectionArgs (Optional) Selection arguments used in the query.
	 * @return The value of the _data column, which is typically a file path.
	 */
	public static String getDataColumn(Context context, Uri uri, String selection,
									   String[] selectionArgs) {

		Cursor cursor = null;
		final String column = "_data";
		final String[] projection = {column};

		try {
			cursor = context.getContentResolver().query(uri, projection, selection, selectionArgs,
					null);
			if (cursor != null && cursor.moveToFirst()) {
				final int column_index = cursor.getColumnIndexOrThrow(column);
				return cursor.getString(column_index);
			}
		} finally {
			if (cursor != null)
				cursor.close();
		}
		return null;
	}

	public static boolean isExternalStorageDocument(Uri uri) {
		return "com.android.externalstorage.documents".equals(uri.getAuthority());
	}

	public static boolean isDownloadsDocument(Uri uri) {
		return "com.android.providers.downloads.documents".equals(uri.getAuthority());
	}

	public static boolean isMediaDocument(Uri uri) {
		return "com.android.providers.media.documents".equals(uri.getAuthority());
	}
}
