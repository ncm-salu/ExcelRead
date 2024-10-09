package excel.read;

import java.io.IOException;
/*dependencies required for excel operations
 * 1.poi
 * 2.poi-ooxml
 */
public class ExcelMain {
	public static void main(String[] args) throws IOException {



		String name =ExcelRead.readStringData(1,0);



		System.out.println("Name :"+ "\t"+name);



		String id= ExcelRead.readIntegerData(1, 1);



		System.out.println("id :"+"\t"+ id);



		}
}
