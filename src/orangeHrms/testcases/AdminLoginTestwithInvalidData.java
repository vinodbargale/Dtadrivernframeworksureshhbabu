package orangeHrms.testcases;

import java.io.IOException;

import org.testng.Assert;
import org.testng.annotations.Test;

import orangehrm.library.LoginPage;
import utilis.AppUtils;
import utilis.Xlutils;
import utilis.Xlutils1;

public class AdminLoginTestwithInvalidData extends AppUtils
{

	String datafile = "C:\\Users\\DELL\\eclipse-workspace\\XLDDTPROJECT\\testdatafiles\\Data.xlsx";
	String datasheet = "AdminLogin_InvalidData";
	
	@Test
	public void checkAdminLogin() throws IOException
	{
		
		int rowcount = Xlutils1.getRowCount(datafile, datasheet);
		String uid,pwd;
		
		LoginPage lp = new LoginPage();
		
		for(int i=1;i<=rowcount;i++)
		{
			uid = Xlutils.getStringCellData(datafile, datasheet, i, 0);
			pwd = Xlutils1.getStringCellData(datafile, datasheet, i, 1);
			lp.login(uid, pwd);
			boolean res = lp.isErrMsgDisplayed();
			Assert.assertTrue(res);
			if(res)
			{
				Xlutils1.setCellData(datafile, datasheet, i, 2, "Pass");
				Xlutils1.fillGreenColor(datafile, datasheet, i, 2);
			}else
			{
				Xlutils1.setCellData(datafile, datasheet, i, 2, "Fail");
				Xlutils1.fillRedColor(datafile, datasheet, i, 2);
			}			
		}
		
	}
	
	
}
