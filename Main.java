package paqu;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Formula;
import jxl.write.WritableHyperlink;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import net.sf.json.JSONObject;
import user.user;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.ArrayList;
import org.springframework.context.ApplicationContext;
import org.springframework.context.support.ClassPathXmlApplicationContext;
import aj.org.objectweb.asm.Label;

public class Main {
	
	public  ArrayList<user> list5;
/**
	 * @return the list5
	 */
	public ArrayList<user> getList5() {
		return list5;
	}
	/**
	 * @param list5 the list5 to set
	 */
	public void setList5(ArrayList<user> list5) {
		this.list5 = list5;
	}
/**
	 * @return the list5
	 */
	public Main(){};
public  static void main(String []args) throws IOException, BiffException{
	WritableWorkbook sheet= Workbook.createWorkbook(new File( "d.xls" ));   
	 System.out.print("asa");
	 WritableSheet sheets = sheet.getSheet(0);
	 System.out.print("asa");
String formu = "HYPERLINK(\"0photo.jpg\",\"查看图片\")";
Formula formula =new Formula(1,1,formu);
try {
	sheets.addCell(formula);
	sheet.close();
} catch (WriteException e) {
	// TODO Auto-generated catch block
	e.printStackTrace();
}
	}


public   ArrayList<user> list()throws BiffException, IOException{
	   ApplicationContext s=new ClassPathXmlApplicationContext("applicationContext.xml");
	   System.out.println("成功");
	   System.out.println("成功");
	   ArrayList<String> list=new ArrayList<String>();
	   System.out.println("成功");
	   String ww=System.getProperty("user.dir");
	   Main.class.getResourceAsStream("/test.txt"); 
	   Workbook sheet= Workbook.getWorkbook( Main.class.getResourceAsStream("/表格/用户记录.xls"));
	   System.out.println("成功");
       Sheet sheets = sheet.getSheet(0);
	  /* Workbook workbook = Workbook.getWorkbook(File); 
	   System.out.println("成功");
	    Sheet[] sheets = workbook.getSheets();
	    System.out.println("成功2");*/
	    if(sheets!=null){
	            // 获得行数
	            int rows = sheets.getRows();
	            // 获得列数  
	            int cols = sheets.getColumns();
	            // 读取数据
	            for (int row=3; row<10; row++)
	            {
	            	ArrayList<String> list2=new ArrayList<String>();
	            	list.add(sheets.getCell(0,row).getContents());
	            	for(int col=0;col<cols;col++){
	                 list2.add(sheets.getCell(col, row).getContents());
	                 /*System.out.println(sheets.getCell(col, row).getContents()); */   
	            	}
	            
	            	user u=(user) s.getBean("user");
	            	u.setLeixing(list2.get(2));
	            	u.setName(list2.get(1));
	            	u.setPhoto(list2.get(0));
	            	u.setShijian(list2.get(4));
	            	u.setWeizhi(list2.get(3));
	            	list5.add(u);	
	            }
	    	
	    }
	    sheet.close();
	    JSONObject a=null;
	    /*  网页链接*/
	   
	    System.out.println("方法list成功");
	    String strUr = "http://192.168.18.50";
		for(int i=0;i<list.size();i++){
			System.out.println(list.get(i));
		String  strUrl=strUr+list.get(i);
		URL url = new URL(strUrl);
		HttpURLConnection conn = (HttpURLConnection)url.openConnection();
		conn.setRequestProperty("User-Agent","Mozilla/5.0 (Windows NT 6.3; WOW64; Trident/7.0; rv:11.0) like Gecko");
		conn.connect();
		InputStream inStream = conn.getInputStream();
		ByteArrayOutputStream outStream = new ByteArrayOutputStream();
		byte [] buf = new byte[1024];
		int len = 0;
		while((len=inStream.read(buf))!=-1){
		outStream.write(buf,0,len);
		}
		inStream.close();
		outStream.close();
	    String a1=list.get(i);
	    String path=System.getProperty("user.dir");
		File file = new File(path+"/paqu/WebRoot/photo/"+i+"photo.jpg");
		String path1=path+"/paqu/WebRoot/photo/"+i+"photo.jpg";
		FileOutputStream op = new FileOutputStream(file);
		op.write(outStream.toByteArray());
		System.out.println(outStream.toByteArray());
		op.close();
		list5.get(i).setPhoto("photo/"+i+"photo.jpg");
		}
		System.out.print("成功");
	    return list5;
	
}
}
