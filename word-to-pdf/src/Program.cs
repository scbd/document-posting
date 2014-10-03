using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;
using System.Text.RegularExpressions;
using System.Threading;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;

namespace Posting
{
	static class ext
	{
		public static T Get<T>(this SortedList<string, object> props, string key)
		{
			return props.ContainsKey(key) ? (T)props[key] : default(T);
		}
	}

	class Program
	{
		static void Main(string[] args)
		{
			List<object> results = new List<object>();

			Word.Formats? targetFormat = Word.Formats.PDF;

			if(args.Length>1 && args[0]=="--nopdf") {
				targetFormat = null;
				args = args.Skip(1).ToArray();
			}

			foreach(string sourcePathX in args)
			{
				var filePattern = Path.GetFileName(sourcePathX);
				var directory   = Path.GetDirectoryName(sourcePathX);
 
				foreach(string sourcePath in Directory.GetFiles(directory, filePattern))
				{
					try
					{
						string targetPath = Path.ChangeExtension(sourcePath, ".pdf"); 

						if(!File.Exists(sourcePath))
							throw new Exception("Source file not found: " + sourcePath);

						if(!new Regex(@".*\.docx?$").IsMatch(sourcePath))
							throw new Exception("Source file must be a .doc - " + sourcePath);

						if(!new Regex(@".*\.pdf").IsMatch(targetPath))
							throw new Exception("Target file must be a .pdf - " + targetPath);


						Regex  validlocaleRE = new Regex(@"^[a-z]{2,3}$");
						Regex  localeRE      = new Regex(@".*-([a-z]{2,3})\.[a-z]+", RegexOptions.IgnoreCase);
						string locale        = localeRE.Replace(sourcePath, "$1").ToLower();

						if(!validlocaleRE.IsMatch(locale??""))
							locale = null;

						SortedList<string, object> result = new SortedList<string,object>();

						using(Word oWord = new Word()) 
						{
							Microsoft.Office.Interop.Word.Document doc = null;

							try
							{
								doc = oWord.OpenDocument(sourcePath);

								var props = oWord.GetProperties(doc);

								result["symbol"]         = props.Get<string>("Symbol")   ?? "";
								result["title"]          = props.Get<string>("Title")    ?? "";
								result["subject"]        = props.Get<string>("Subject")  ?? "";
								result["company"]        = props.Get<string>("Company")  ?? "";
								result["keywords"]       = props.Get<string>("Keywords") ?? "";
								result["comments"]       = props.Get<string>("Comments") ?? "";
								result["trackRevisions"] = doc.TrackRevisions;
								result["hasComments"]    = doc.Comments.Count!=0;
								result["pageCount"]      = doc.ComputeStatistics(Microsoft.Office.Interop.Word.WdStatistic.wdStatisticPages);
								result["wordCount"]      = doc.ComputeStatistics(Microsoft.Office.Interop.Word.WdStatistic.wdStatisticWords, true);

								if(targetFormat.HasValue)
									oWord.SaveDocument(doc, targetPath, Word.Formats.PDF);
							}
							finally
							{
								if(doc!=null)
									oWord.CloseDocument(doc);
							}
						}

						result["success"] = true;
						result["locale"]  = locale;
						result["source"]  = sourcePath;

						if(targetFormat.HasValue)
							result["target"]  = targetPath;

						results.Add(result);
					}
					catch(Exception ex)
					{
						results.Add(new { success = false, source=sourcePath, error = ex.Message });
						Environment.ExitCode = -1;
					}
				}
			}

			if(results.Count==1)  Console.WriteLine(JsonConvert.SerializeObject(results.Single(),  Formatting.Indented));
			else                  Console.WriteLine(JsonConvert.SerializeObject(results.ToArray(), Formatting.Indented));
		}
	}
}
