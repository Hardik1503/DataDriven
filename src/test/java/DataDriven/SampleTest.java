package DataDriven;

import java.io.IOException;
import java.util.ArrayList;

public class SampleTest {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub

		    TestData TD = new TestData();
		   ArrayList data = TD.getdata("DeleteProfile");
		   
		   System.out.println(data.get(0));
		   System.out.println(data.get(1));
		   System.out.println(data.get(2));
		   System.out.println(data.get(3));
		    
	}

}
