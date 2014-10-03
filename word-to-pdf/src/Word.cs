using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Collections;
using Microsoft.Office.Core;

namespace Posting
{
	class Word : IDisposable
	{
		public enum Formats {
			PDF,
			DOCX,
			DOC
		}

		// Create a new Microsoft Word application object
		private Application mApp;
		private object MISSING = System.Reflection.Missing.Value;
		private object TRUE    = (object)true;
		private object FALSE   = (object)false;

		//============================================================
		//
		//
		//============================================================
		public Word()
		{
			mApp = new Application();
			mApp.Visible = false;
			mApp.ScreenUpdating = false;
		}

		//============================================================
		//
		//
		//============================================================
		~Word()
		{
			((IDisposable)this).Dispose();
		}

		//============================================================
		//
		//
		//============================================================
		private void ShutdownApp()
		{
			if(mApp==null)
				return;

		}

		//============================================================
		//
		//
		//============================================================
		public Document OpenDocument(string filePath)
		{
			// Cast as Object for word Open method
			object filename = (object)filePath;

			// Use the dummy value as a placeholder for optional arguments
			Document doc = mApp.Documents.Open(ref filename, ref FALSE, TRUE, ref MISSING, ref MISSING, ref MISSING, ref TRUE, ref MISSING, ref MISSING, ref MISSING, ref MISSING, ref MISSING, ref MISSING, ref MISSING, ref MISSING, ref MISSING);

			doc.Activate();

			return doc;
		}

		//============================================================
		//
		//
		//============================================================
		public void SaveDocument(Document doc, string filePath, Formats? format)
		{
			object outputFileName = filePath;
			object fileFormat     = MISSING;

			if(format==Formats.PDF)  fileFormat = WdSaveFormat.wdFormatPDF;
			if(format==Formats.DOCX) fileFormat = WdSaveFormat.wdFormatDocument;
			if(format==Formats.DOC)  fileFormat = WdSaveFormat.wdFormatDocument97;

			// Save document into PDF Format
			doc.SaveAs(ref outputFileName, ref fileFormat, ref MISSING, ref MISSING, ref MISSING, ref MISSING, ref MISSING, ref MISSING, ref MISSING, ref MISSING, ref MISSING, ref MISSING, ref MISSING, ref MISSING, ref MISSING, ref MISSING);
		}

		//============================================================
		//
		//
		//============================================================
		public void CloseDocument(Document doc)
		{
			object saveChanges = WdSaveOptions.wdDoNotSaveChanges;

			((_Document)doc).Close(ref saveChanges, ref MISSING, ref MISSING);
		}

		//============================================================
		//
		//
		//============================================================
		public void Convert(string source, string target, Formats format)
		{
			// Use the dummy value as a placeholder for optional arguments
			Document doc = OpenDocument(source);

			SaveDocument(doc, target, format);

			CloseDocument(doc);
		}

		//============================================================
		//
		//
		//============================================================
		public SortedList<string, object> GetProperties(string source)
		{
			// Use the dummy value as a placeholder for optional arguments
			Document doc = null;
            
            try 
            {
                doc = OpenDocument(source);

                return GetProperties(doc);
            }
            finally
            {
                if(doc!=null)
                    CloseDocument(doc);
            }
        }

		//============================================================
		//
		//
		//============================================================
		public SortedList<string, object> GetProperties(Document doc)
		{
			SortedList<string, object> properties = new SortedList<string,object>();

			foreach (dynamic p in doc.BuiltInDocumentProperties)
			{
				try
				{
					if((p.Value as object) == null)
						continue;

					properties[(string)p.Name] = (object)p.Value;
				}
				catch
				{}
			}
			
			foreach (dynamic p in doc.CustomDocumentProperties)
			{
				try
				{
					if((p.Value as object) == null)
						continue;

					properties[(string)p.Name] = (object)p.Value;
				}
				catch
				{}
			}
			
			return properties;
		}

		void IDisposable.Dispose()
		{
			try
			{
				if(mApp!=null)
					((_Application)mApp).Quit(ref MISSING, ref MISSING, ref MISSING);
			}
			finally
			{
				mApp = null;

				GC.Collect();
				GC.WaitForPendingFinalizers();
				GC.Collect();
				GC.WaitForPendingFinalizers();
			}
		}
	}
}
