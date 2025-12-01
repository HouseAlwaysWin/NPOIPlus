using NPOI.SS.UserModel;
using NPOIPlus.Helpers;
using System.IO;

namespace NPOIPlus.Base
{
	public abstract class FluentWorkbookBase
	{
		protected IWorkbook _workbook;

		protected FluentWorkbookBase()
		{
		}

	protected FluentWorkbookBase(IWorkbook workbook)
	{
		_workbook = workbook;
	}

	public IWorkbook GetWorkbook()
	{
		return _workbook;
	}

	public FluentMemoryStream ToStream()
		{
			var ms = new FluentMemoryStream();
			ms.AllowClose = false;
			_workbook.Write(ms);
			ms.Flush();
			ms.Seek(0, SeekOrigin.Begin);
			ms.AllowClose = true;
			return ms;
		}

		public IWorkbook SaveToPath(string filePath)
		{
			using (FileStream outFile = new FileStream(filePath, FileMode.Create, FileAccess.Write))
			{
				_workbook.Write(outFile);
			}
			return _workbook;
		}
	}
}

