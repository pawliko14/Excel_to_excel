package parameters;

public class Parameters {
	
	private Parameters Instance = new Parameters();
	
		// to tests
//		private static final String PATH_TO_READING_FILE = "C:\\Users\\el08\\Desktop\\Irek_exel\\WYSY£KI MASZYN.xls";
//	    private static final String PATH_TO_FOLDER = "C:\\Users\\el08\\Desktop\\Irek_exel\\";
//	    private static final String PATH_TO_IREK_FILE = "C:\\Users\\el08\\Desktop\\Irek_exel\\backups\\sales_by_year_test.xls";
//	    private static final String PATH_TO_IREK_FILE_backup = "C:\\Users\\el08\\Desktop\\Irek_exel\\backups\\sales_by_year_test.xls";
//	    
	    
		private static final String PATH_TO_READING_FILE = "\\\\192.168.90.203\\Common\\SalesByYear\\Testy\\WYSY£KI MASZYN.xls";
	    private static final String PATH_TO_FOLDER = "\\\\192.168.90.203\\Common\\SalesByYear\\Testy\\";
	    private static final String PATH_TO_IREK_FILE = "\\\\192.168.90.203\\Common\\SalesByYear\\Testy\\sales_by_year_test.xls";
	    private static final String PATH_TO_IREK_FILE_backup = "\\\\192.168.90.203\\Common\\SalesByYear\\Testy\\sales_by_year_test.xls";
	    
	    /*
	     *  finall path will be in dataserver in marketing share folder
	     */
	    
	    
	    
	    public static String getSampleXlsxFilePath() {
			return PATH_TO_READING_FILE;
		}




		public static String getPathToFolder() {
			return PATH_TO_FOLDER;
		}




		public static String getPathToIrekFile() {
			return PATH_TO_IREK_FILE;
		}




		public static String getPathToIrekFileBackup() {
			return PATH_TO_IREK_FILE_backup;
		}




		public Parameters getInstance()
	    {
	    	if(Instance == null)
	    	{
	    		Instance = new Parameters();
	    	}
	    	
			return Instance;
	    	
	    }
	    

}
