using System.IO;

namespace NPOIPlus.Helpers
{
	public class FluentMemoryStream : MemoryStream
	{
		public FluentMemoryStream()
		{
			// We always want to close streams by default to
			// force the developer to make the conscious decision
			// to disable it.  Then, they're more apt to remember
			// to re-enable it.  The last thing you want is to
			// enable memory leaks by default.  ;-)
			AllowClose = true;
		}

		public bool AllowClose { get; set; }

		public override void Close()
		{
			if (AllowClose)
				base.Close();
		}
	}
}

