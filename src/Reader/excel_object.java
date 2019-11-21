package Reader;

import java.lang.reflect.Field;

public class excel_object  {

	private String Country;
	private String Client;
	private String Machine_type;
	private String SN;
	private String Quantity;
	private String Date;
	private String Year;
	private String Value_EUR;
	private String Value_PLN;
	private String Kurs_EUR;

	
	public Field[] get_all_fields()
	{
		Field[] fields = getClass().getDeclaredFields();
		
		return fields;
		
	}
	
	public int get_number_of_fields()
	{
		return getClass().getDeclaredFields().length;
	}
	
	//default constructor, do not use, better use Builder()
	public excel_object(String country, String client, String machine_type, String sN, String quantity, String date,
			String year, String value_EUR, String value_PLN, String kurs_EUR) {
			Country = country;
			Client = client;
			Machine_type = machine_type;
			SN = sN;
			Quantity = quantity;
			Date = date;
			Year = year;
			Value_EUR = value_EUR;
			Value_PLN = value_PLN;
			Kurs_EUR = kurs_EUR;
		}
	
	//necessary to exist because of Builder()
	public excel_object()
	{
		
	}
	
	public void printObject()
	{	System.out.println("------------------------");
		System.out.println("Country: " + this.Country);
		System.out.println("Client: " + this.Client);
		System.out.println("Machine_type: " + this.Machine_type);
		System.out.println("SN: " + this.SN);
		System.out.println("Quantity: " + this.Quantity);
		System.out.println("Date: " + this.Date);
		System.out.println("Year: " + this.Year);
		System.out.println("Value_EUR: " + this.Value_EUR);
		System.out.println("Value_PLN: " + this.Value_PLN);
		System.out.println("Kurs_EUR: " + this.Kurs_EUR);
		
	}
		
	
	// typical Builder pattern, very usefull to create object instance with many parameters
	public static final class Builder{
		
		private String Country;
		private String Client;
		private String Machine_type;
		private String SN;
		private String Quantity;
		private String Date;
		private String Year;
		private String Value_EUR;
		private String Value_PLN;
		private String Kurs_EUR;
		
		public Builder Country(String Country)
		{
			this.Country = Country;
			return this;
		}
		public Builder Client(String Client)
		{
			this.Client = Client;
			return this;
		}
		public Builder Machine_type(String Machine_type)
		{
			this.Machine_type = Machine_type;
			return this;
		}	
		public Builder SN(String SN)
		{
			this.SN = SN;
			return this;
		}
		public Builder Quantity(String Quantity)
		{
			this.Quantity = Quantity;
			return this;
		}
		public Builder Date(String Date)
		{
			this.Date = Date;
			return this;
		}
		public Builder Year(String Year)
		{
			this.Year = Year;
			return this;
		}
		public Builder Value_EUR(String Value_EUR)
		{
			this.Value_EUR = Value_EUR;
			return this;
		}
		
		public Builder Value_PLN(String Value_PLN)
		{
			this.Value_PLN = Value_PLN;
			return this;
		}
		public Builder Kurs_EUR(String Kurs_EUR)
		{
			this.Kurs_EUR = Kurs_EUR;
			return this;
		}
		
		
		public excel_object build()
		{
			excel_object excel_ob = new excel_object();
			

			
			excel_ob.Country = this.Country;
			excel_ob.Client = this.Client;
			excel_ob.Machine_type = this.Machine_type;
			excel_ob.SN = this.SN;
			excel_ob.Quantity = this.Quantity;
			excel_ob.Date = this.Date;
			excel_ob.Year = this.Year;
			excel_ob.Value_EUR = this.Value_EUR;
			excel_ob.Value_PLN = this.Value_PLN;
			excel_ob.Kurs_EUR = this.Kurs_EUR;
			
			
			return excel_ob;
		}
		
	}
	
	

	
	
	public String getCountry() {
		return Country;
	}
	public String getClient() {
		return Client;
	}
	public String getMachine_type() {
		return Machine_type;
	}
	public String getSN() {
		return SN;
	}
	public String getQuantity() {
		return Quantity;
	}
	public String getDate() {
		return Date;
	}
	public String getYear() {
		return Year;
	}
	public String getValue_EUR() {
		return Value_EUR;
	}
	public String getValue_PLN() {
		return Value_PLN;
	}
	public String getKurs_EUR() {
		return Kurs_EUR;
	}
	
	


	
}
