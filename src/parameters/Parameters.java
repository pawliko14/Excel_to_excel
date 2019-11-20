package parameters;

public class Parameters {
	
	private Parameters Instance = new Parameters();
	
		private static final String SAMPLE_XLSX_FILE_PATH = "C:\\Users\\el08\\Desktop\\Irek_exel\\WYSY£KI MASZYN.xls";
	    private static final String PATH_TO_FOLDER = "C:\\Users\\el08\\Desktop\\Irek_exel\\";
	    
	    public Parameters getInstance()
	    {
	    	if(Instance == null)
	    	{
	    		Instance = new Parameters();
	    	}
	    	
			return Instance;
	    	
	    }
	    

}
