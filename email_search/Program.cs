using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Text;



namespace email_search
{
	class MainClass
	{
		public static StringBuilder sb_master = new StringBuilder();
		public static void Main (string[] args)
		{
			dosya_listesi ();
		}

		public static void dosya_listesi()
		{


			DirectoryInfo di = new DirectoryInfo(Yol());

			//FileInfo tipinden bir değişken oluşturuyoruz.
			//çünkü di.GetFiles methodu, bize FileInfo tipinden bir dizi dönüyor.
			FileInfo[] rgFiles = di.GetFiles();
			//foreach döngümüzle fgFiles içinde dönüyoruz.
			foreach (FileInfo fi in rgFiles)
			{
				//fi.Name bize dosyanın adını dönüyor.
				//fi.FullName ise bize dosyasının dizin bilgisini döner.
				Console.WriteLine ( fi.Name);
				ExtractEmails (Yol(),fi.Name);
				// String[] str = GetExcelSheetNames(conn);
			}

			////verilen yol içinde export_master klasörü içine mail adreslerini kayıt yapıyor
			File.WriteAllText(Yol()+"/export_master/export_master.txt", sb_master.ToString());

		}


		public static void ExtractEmails(string inFilePath,string file_name)
		{
			Console.WriteLine("----"+inFilePath);
			string data = File.ReadAllText(inFilePath+file_name); //read File 
			//instantiate with this pattern 
			Regex emailRegex = new Regex(@"\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*",
				RegexOptions.IgnoreCase);
			//find items that matches with our pattern
			MatchCollection emailMatches = emailRegex.Matches(data);

			StringBuilder sb = new StringBuilder();

			foreach (Match emailMatch in emailMatches)
			{
				sb.AppendLine(emailMatch.Value);
				sb_master.AppendLine(emailMatch.Value);
			}
			//verilen yol içinde export klasörü içine mail adreslerini kayıt yapıyor
			File.WriteAllText(inFilePath+"/export/"+file_name+".txt", sb.ToString());
		}

		public static String Yol()
		{
			string path = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
			string yol = Path.Combine(path+"/mail/");
			return yol;
		}
	}
}
